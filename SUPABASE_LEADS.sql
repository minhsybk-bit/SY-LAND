-- SỸ LAND: yêu cầu tư vấn và khách hàng tiềm năng
-- Chạy sau các tệp schema tài khoản, thiết bị và audit.

create table if not exists public.leads (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
  unit text not null check (char_length(unit) between 2 and 200),
  contact text not null check (char_length(contact) between 3 and 200),
  monthly_volume integer not null default 1 check (monthly_volume between 1 and 10000000),
  people integer not null default 1 check (people between 1 and 100000),
  needs text[] not null default '{}',
  proposed_plan text not null check (proposed_plan in ('Cá nhân', 'Văn phòng', 'Đơn vị')),
  note text not null default '' check (char_length(note) <= 5000),
  status text not null default 'Mới' check (status in ('Mới', 'Đang liên hệ', 'Đã chuyển đổi', 'Đã đóng')),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists leads_created_index on public.leads(created_at desc);
create index if not exists leads_status_index on public.leads(status);
alter table public.leads enable row level security;

drop policy if exists "lead_owner_insert" on public.leads;
create policy "lead_owner_insert" on public.leads
for insert to authenticated
with check (user_id = auth.uid() and status = 'Mới');

drop policy if exists "lead_owner_read" on public.leads;
create policy "lead_owner_read" on public.leads
for select to authenticated
using (user_id = auth.uid() or public.is_syland_admin());

drop policy if exists "lead_admin_update" on public.leads;
create policy "lead_admin_update" on public.leads
for update to authenticated
using (public.is_syland_admin())
with check (public.is_syland_admin());

create or replace function public.set_lead_updated_at()
returns trigger language plpgsql as $$
begin new.updated_at = now(); return new; end;
$$;

drop trigger if exists leads_updated_at_trigger on public.leads;
create trigger leads_updated_at_trigger before update on public.leads
for each row execute procedure public.set_lead_updated_at();

-- Ghi nhật ký khi admin thay đổi trạng thái khách hàng (nếu audit_events đã có).
create or replace function public.write_lead_audit()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  if old.status is distinct from new.status then
    insert into public.audit_events(actor_id, action, entity_type, entity_id, details)
    values (auth.uid(), 'lead_status_changed', 'leads', new.id::text,
      jsonb_build_object('unit', new.unit, 'from', old.status, 'to', new.status));
  end if;
  return null;
end;
$$;

drop trigger if exists leads_audit_trigger on public.leads;
create trigger leads_audit_trigger after update on public.leads
for each row execute procedure public.write_lead_audit();

