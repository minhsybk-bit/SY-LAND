-- SỸ LAND Creator: dữ liệu dự án video, tác vụ xử lý và nhật ký hạn mức.
-- Chạy bằng vai trò postgres trong Supabase SQL Editor sau SUPABASE_SCHEMA.sql.

begin;

create table if not exists public.creator_projects (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id) on delete cascade,
  title text not null default '',
  source_type text not null check (source_type in ('url', 'upload')),
  source_url text,
  source_platform text not null default 'unknown',
  source_visibility text not null default 'unknown'
    check (source_visibility in ('public', 'unlisted', 'private', 'restricted', 'unknown')),
  ownership_status text not null default 'unverified'
    check (ownership_status in ('verified', 'declared', 'unverified')),
  rights_status text not null default 'pending'
    check (rights_status in ('pending', 'low', 'medium', 'high', 'blocked')),
  rights_evidence jsonb not null default '{}'::jsonb,
  rights_confirmed_at timestamptz,
  status text not null default 'draft'
    check (status in ('draft', 'ready', 'processing', 'completed', 'failed', 'cancelled')),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.creator_jobs (
  id uuid primary key default gen_random_uuid(),
  project_id uuid not null references public.creator_projects(id) on delete cascade,
  user_id uuid not null references auth.users(id) on delete cascade,
  status text not null default 'queued'
    check (status in ('queued', 'downloading', 'transcribing', 'translating', 'dubbing', 'rendering', 'completed', 'failed', 'cancelled')),
  progress integer not null default 0 check (progress between 0 and 100),
  provider text not null check (provider in ('openai', 'gemini')),
  voice text not null,
  whisper_model text not null default 'small',
  background_volume integer not null default 10 check (background_volume between 0 and 30),
  subtitle_size text not null default 'medium' check (subtitle_size in ('small', 'medium', 'large')),
  duration_seconds numeric(12,3),
  input_object_key text,
  output_object_key text,
  queue_task_id text,
  error_code text,
  error_message text,
  started_at timestamptz,
  completed_at timestamptz,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

alter table public.creator_jobs add column if not exists queue_task_id text;

create table if not exists public.creator_usage_ledger (
  id bigint generated always as identity primary key,
  user_id uuid not null references auth.users(id) on delete cascade,
  job_id uuid references public.creator_jobs(id) on delete set null,
  minutes numeric(12,3) not null check (minutes >= 0),
  event_type text not null check (event_type in ('reserve', 'consume', 'refund', 'adjustment')),
  note text not null default '',
  created_at timestamptz not null default now()
);

create table if not exists public.creator_plan_limits (
  plan text primary key,
  monthly_minutes numeric(12,3) not null check (monthly_minutes > 0),
  max_video_minutes numeric(12,3) not null check (max_video_minutes > 0),
  updated_at timestamptz not null default now()
);

-- Hạn mức khởi tạo cho bản thương mại. Có thể thay đổi trong SQL Editor mà
-- không cần sửa mã nguồn hoặc build lại website.
insert into public.creator_plan_limits (plan, monthly_minutes, max_video_minutes)
values
  ('Dùng thử', 10, 5),
  ('Go', 60, 10),
  ('Plus', 180, 15),
  ('Pro', 600, 30),
  ('Cá nhân', 60, 10),
  ('Văn phòng', 300, 20),
  ('Đơn vị', 1000, 30)
on conflict (plan) do nothing;

create index if not exists creator_projects_user_created_idx
  on public.creator_projects(user_id, created_at desc);
create index if not exists creator_jobs_user_created_idx
  on public.creator_jobs(user_id, created_at desc);
create index if not exists creator_jobs_status_created_idx
  on public.creator_jobs(status, created_at);
create index if not exists creator_usage_user_created_idx
  on public.creator_usage_ledger(user_id, created_at desc);
create unique index if not exists creator_usage_one_reserve_per_job_idx
  on public.creator_usage_ledger(job_id)
  where event_type = 'reserve';
create unique index if not exists creator_usage_one_consume_per_job_idx
  on public.creator_usage_ledger(job_id)
  where event_type = 'consume';
create unique index if not exists creator_usage_one_refund_per_job_idx
  on public.creator_usage_ledger(job_id)
  where event_type = 'refund';

alter table public.creator_projects enable row level security;
alter table public.creator_jobs enable row level security;
alter table public.creator_usage_ledger enable row level security;
alter table public.creator_plan_limits enable row level security;

drop policy if exists "creator_project_owner_read" on public.creator_projects;
create policy "creator_project_owner_read" on public.creator_projects
for select using (user_id = auth.uid() or public.is_syland_admin());

drop policy if exists "creator_project_owner_insert" on public.creator_projects;
create policy "creator_project_owner_insert" on public.creator_projects
for insert with check (user_id = auth.uid());

drop policy if exists "creator_project_owner_update" on public.creator_projects;
create policy "creator_project_owner_update" on public.creator_projects
for update using (user_id = auth.uid() or public.is_syland_admin())
with check (user_id = auth.uid() or public.is_syland_admin());

drop policy if exists "creator_job_owner_read" on public.creator_jobs;
create policy "creator_job_owner_read" on public.creator_jobs
for select using (user_id = auth.uid() or public.is_syland_admin());

-- Tác vụ chỉ được tạo qua máy chủ Creator. Trình duyệt không có quyền INSERT.
drop policy if exists "creator_usage_owner_read" on public.creator_usage_ledger;
create policy "creator_usage_owner_read" on public.creator_usage_ledger
for select using (user_id = auth.uid() or public.is_syland_admin());

drop policy if exists "creator_limits_authenticated_read" on public.creator_plan_limits;
create policy "creator_limits_authenticated_read" on public.creator_plan_limits
for select to authenticated using (true);

create or replace function public.creator_reserve_minutes(
  p_job_id uuid,
  p_minutes numeric
)
returns void
language plpgsql
security definer
set search_path = public, auth
as $$
declare
  v_user_id uuid;
  v_email text;
  v_plan text := 'Dùng thử';
  v_limit numeric;
  v_max_video numeric;
  v_used numeric;
begin
  if p_minutes <= 0 then
    raise exception 'CREATOR_INVALID_MINUTES';
  end if;

  select user_id into v_user_id
  from public.creator_jobs
  where id = p_job_id;
  if v_user_id is null then
    raise exception 'CREATOR_JOB_NOT_FOUND';
  end if;

  perform pg_advisory_xact_lock(hashtext(v_user_id::text));
  if exists (
    select 1 from public.creator_usage_ledger
    where job_id = p_job_id and event_type = 'reserve'
  ) then
    return;
  end if;

  select lower(coalesce(email, '')) into v_email
  from auth.users
  where id = v_user_id;

  select l.plan into v_plan
  from public.licenses l
  where lower(l.email) = v_email
    and l.status = 'Hoạt động'
    and l.expires_at > now()
  order by l.expires_at desc
  limit 1;
  v_plan := coalesce(v_plan, 'Dùng thử');

  select monthly_minutes, max_video_minutes
  into v_limit, v_max_video
  from public.creator_plan_limits
  where plan = v_plan;
  if v_limit is null then
    raise exception 'CREATOR_PLAN_NOT_CONFIGURED';
  end if;
  if p_minutes > v_max_video then
    raise exception 'CREATOR_VIDEO_TOO_LONG_FOR_PLAN';
  end if;

  select coalesce(sum(
    case
      when event_type = 'reserve' then minutes
      when event_type = 'refund' then -minutes
      else 0
    end
  ), 0)
  into v_used
  from public.creator_usage_ledger
  where user_id = v_user_id
    and created_at >= date_trunc('month', now());

  if v_used + p_minutes > v_limit then
    raise exception 'CREATOR_QUOTA_EXCEEDED';
  end if;

  insert into public.creator_usage_ledger (user_id, job_id, minutes, event_type, note)
  values (v_user_id, p_job_id, p_minutes, 'reserve', 'Giữ hạn mức trước khi xử lý');
end;
$$;

create or replace function public.creator_consume_minutes(
  p_job_id uuid,
  p_minutes numeric
)
returns void
language plpgsql
security definer
set search_path = public
as $$
declare
  v_user_id uuid;
begin
  select user_id into v_user_id from public.creator_jobs where id = p_job_id;
  if v_user_id is null then raise exception 'CREATOR_JOB_NOT_FOUND'; end if;
  if not exists (
    select 1 from public.creator_usage_ledger
    where job_id = p_job_id and event_type = 'consume'
  ) then
    insert into public.creator_usage_ledger (user_id, job_id, minutes, event_type, note)
    values (v_user_id, p_job_id, p_minutes, 'consume', 'Tác vụ hoàn thành');
  end if;
end;
$$;

create or replace function public.creator_refund_minutes(
  p_job_id uuid,
  p_minutes numeric
)
returns void
language plpgsql
security definer
set search_path = public
as $$
declare
  v_user_id uuid;
begin
  select user_id into v_user_id from public.creator_jobs where id = p_job_id;
  if v_user_id is null then raise exception 'CREATOR_JOB_NOT_FOUND'; end if;
  if exists (
    select 1 from public.creator_usage_ledger
    where job_id = p_job_id and event_type = 'reserve'
  ) and not exists (
    select 1 from public.creator_usage_ledger
    where job_id = p_job_id and event_type = 'refund'
  ) then
    insert into public.creator_usage_ledger (user_id, job_id, minutes, event_type, note)
    values (v_user_id, p_job_id, p_minutes, 'refund', 'Hoàn hạn mức do tác vụ thất bại');
  end if;
end;
$$;

revoke all on function public.creator_reserve_minutes(uuid, numeric) from public, anon, authenticated;
revoke all on function public.creator_consume_minutes(uuid, numeric) from public, anon, authenticated;
revoke all on function public.creator_refund_minutes(uuid, numeric) from public, anon, authenticated;
grant execute on function public.creator_reserve_minutes(uuid, numeric) to service_role;
grant execute on function public.creator_consume_minutes(uuid, numeric) to service_role;
grant execute on function public.creator_refund_minutes(uuid, numeric) to service_role;

commit;
