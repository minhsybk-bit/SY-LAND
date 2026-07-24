-- SỸ LAND 11.6.9 - đồng bộ ổn định tài khoản/gói cho Website và Windows.
-- Chạy TOÀN BỘ một lần trong Supabase Dashboard > SQL Editor.
-- Có thể chạy lại an toàn; không xóa tài khoản, đơn hàng hoặc mã bản quyền.

begin;

alter table public.licenses
  add column if not exists user_id uuid references auth.users(id) on delete set null;
alter table public.licenses
  add column if not exists seat_count integer not null default 1;
alter table public.licenses
  add column if not exists max_devices integer not null default 1;
alter table public.licenses
  add column if not exists max_parcels_per_run integer;

update public.licenses
set seat_count = greatest(1, coalesce(seat_count, max_devices, 1)),
    max_devices = greatest(1, coalesce(max_devices, seat_count, 1));
alter table public.licenses alter column seat_count set default 1;
alter table public.licenses alter column seat_count set not null;
alter table public.licenses alter column max_devices set default 1;
alter table public.licenses alter column max_devices set not null;

-- Bảo đảm mọi tài khoản cũ/mới/Google đều có đúng một profile theo auth.users.id.
insert into public.profiles(id, full_name, email, role)
select
  u.id,
  coalesce(nullif(u.raw_user_meta_data ->> 'full_name', ''), split_part(coalesce(u.email, ''), '@', 1)),
  coalesce(u.email, ''),
  case when lower(coalesce(u.email, '')) = 'minhsybk@gmail.com' then 'admin' else 'user' end
from auth.users u
on conflict (id) do update set
  full_name = coalesce(nullif(excluded.full_name, ''), public.profiles.full_name),
  email = excluded.email,
  role = case
    when lower(excluded.email) = 'minhsybk@gmail.com' then 'admin'
    else public.profiles.role
  end,
  updated_at = now();

-- Gắn các bản quyền lịch sử vào UUID tài khoản; email chỉ còn là khóa tương thích.
update public.licenses l
set user_id = u.id,
    email = lower(u.email),
    updated_at = now()
from auth.users u
where l.user_id is null
  and lower(trim(l.email)) = lower(trim(u.email));

create index if not exists licenses_user_active_index
  on public.licenses(user_id, status, expires_at desc);

-- Bản quyền mới luôn được gắn vào đúng auth.users.id, kể cả khi được
-- tạo từ quy trình thanh toán cũ vốn chỉ lưu email.
create or replace function public.link_license_to_auth_user()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  new.email := lower(trim(coalesce(new.email, '')));
  if new.user_id is null and new.email <> '' then
    select u.id into new.user_id
    from auth.users u
    where lower(trim(coalesce(u.email, ''))) = new.email
    limit 1;
  end if;
  return new;
end;
$$;

drop trigger if exists link_license_to_auth_user_trigger on public.licenses;
create trigger link_license_to_auth_user_trigger
before insert or update of email, user_id on public.licenses
for each row execute procedure public.link_license_to_auth_user();

create or replace function public.handle_new_user()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  insert into public.profiles(id, full_name, email, role)
  values (
    new.id,
    coalesce(nullif(new.raw_user_meta_data ->> 'full_name', ''), split_part(coalesce(new.email, ''), '@', 1)),
    lower(coalesce(new.email, '')),
    case when lower(coalesce(new.email, '')) = 'minhsybk@gmail.com' then 'admin' else 'user' end
  )
  on conflict (id) do update set
    full_name = coalesce(nullif(excluded.full_name, ''), public.profiles.full_name),
    email = excluded.email,
    role = case
      when lower(excluded.email) = 'minhsybk@gmail.com' then 'admin'
      else public.profiles.role
    end,
    updated_at = now();
  return new;
end;
$$;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
after insert or update of email, raw_user_meta_data on auth.users
for each row execute procedure public.handle_new_user();

create or replace function public.is_syland_admin()
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select exists (
    select 1
    from public.profiles p
    where p.id = auth.uid()
      and p.role = 'admin'
  ) or lower(coalesce(auth.jwt() ->> 'email', '')) = 'minhsybk@gmail.com';
$$;

drop policy if exists "license_owner_read" on public.licenses;
create policy "license_owner_read" on public.licenses
for select to authenticated
using (
  user_id = auth.uid()
  or (
    user_id is null
    and lower(email) = lower(coalesce(auth.jwt() ->> 'email', ''))
  )
  or public.is_syland_admin()
);

create or replace function public.get_my_syland_entitlements()
returns jsonb
language plpgsql
stable
security definer
set search_path = public
as $$
declare
  v_uid uuid := auth.uid();
  v_email text := lower(coalesce(auth.jwt() ->> 'email', ''));
  v_role text := 'user';
  v_license public.licenses%rowtype;
  v_plan text := 'Dùng thử';
  v_seats integer := 1;
  v_limit integer := 50;
  v_percent integer := 25;
  v_file_mb integer := 20;
  v_total_mb integer := 150;
  v_full boolean := false;
  v_reason text := 'Tài khoản dùng thử';
begin
  if v_uid is null then
    return jsonb_build_object('ok', false, 'message', 'Phiên đăng nhập không hợp lệ.');
  end if;

  select coalesce(p.role, 'user') into v_role
  from public.profiles p
  where p.id = v_uid;

  if v_role = 'admin' or v_email = 'minhsybk@gmail.com' then
    return jsonb_build_object(
      'ok', true, 'schema_version', 1, 'user_id', v_uid, 'email', v_email,
      'role', 'admin', 'plan', 'Quản trị viên', 'status', 'Hoạt động',
      'expires_at', null, 'seat_count', null, 'max_devices', null,
      'max_parcels_per_run', null, 'feature_percent', 100,
      'max_file_size_mb', 100, 'max_total_upload_mb', 2000,
      'max_uses_per_day', null, 'full_tools', true,
      'reason', 'Quản trị viên SỸ LAND · Full Access',
      'source', 'profiles.role'
    );
  end if;

  select l.* into v_license
  from public.licenses l
  where (l.user_id = v_uid or (l.user_id is null and lower(trim(l.email)) = v_email))
    and l.status = 'Hoạt động'
    and l.expires_at > now()
  order by
    case when l.user_id = v_uid then 0 else 1 end,
    l.expires_at desc,
    l.created_at desc
  limit 1;

  if v_license.id is not null then
    v_plan := case
      when lower(v_license.plan) like '%văn phòng%' then 'Văn phòng'
      when lower(v_license.plan) like '%đơn vị%' then 'Đơn vị'
      when lower(v_license.plan) like '%pro%' then 'Pro'
      when lower(v_license.plan) like '%plus%' or lower(v_license.plan) = 'cá nhân' then 'Plus'
      when lower(v_license.plan) like '%go%' then 'Go'
      else 'Dùng thử'
    end;
    v_seats := greatest(1, coalesce(v_license.seat_count, v_license.max_devices, 1));

    if v_plan = 'Go' then
      v_limit := 80; v_percent := 40; v_file_mb := 20; v_total_mb := 500;
      v_reason := 'Cá nhân Go · tối đa 80 thửa/lần';
    elsif v_plan = 'Plus' or (v_plan = 'Văn phòng' and v_seats < 5) then
      v_limit := 140; v_percent := 70; v_file_mb := 30; v_total_mb := 750;
      v_reason := case when v_plan = 'Văn phòng'
        then format('Văn phòng %s tài khoản · quyền tương đương Plus', v_seats)
        else 'Cá nhân Plus · tối đa 140 thửa/lần' end;
    elsif v_plan = 'Pro' then
      v_limit := 200; v_percent := 100; v_file_mb := 50; v_total_mb := 1000; v_full := true;
      v_reason := 'Cá nhân Pro · đầy đủ công cụ, tối đa 200 thửa/lần';
    elsif v_plan = 'Đơn vị' or (v_plan = 'Văn phòng' and v_seats >= 5) then
      v_limit := null; v_percent := 100; v_file_mb := 100; v_total_mb := 2000; v_full := true;
      v_reason := case when v_plan = 'Văn phòng'
        then format('Văn phòng %s tài khoản · không giới hạn', v_seats)
        else 'Gói Đơn vị · không giới hạn' end;
    end if;

    if v_license.max_parcels_per_run is not null then
      v_limit := case when v_limit is null then v_license.max_parcels_per_run
                      else least(v_limit, v_license.max_parcels_per_run) end;
    end if;
  end if;

  return jsonb_build_object(
    'ok', true, 'schema_version', 1, 'user_id', v_uid, 'email', v_email,
    'role', 'user', 'license_id', v_license.id, 'license_code', v_license.code,
    'plan', v_plan, 'status', case when v_license.id is null then 'Dùng thử' else v_license.status end,
    'expires_at', v_license.expires_at, 'seat_count', v_seats,
    'max_devices', coalesce(v_license.max_devices, 1),
    'max_parcels_per_run', v_limit, 'feature_percent', v_percent,
    'max_file_size_mb', v_file_mb, 'max_total_upload_mb', v_total_mb,
    'max_uses_per_day', case when v_license.id is null then 10 else null end,
    'full_tools', v_full, 'reason', v_reason,
    'source', case when v_license.id is null then 'trial' else 'licenses' end
  );
end;
$$;

revoke all on function public.get_my_syland_entitlements() from public;
grant execute on function public.get_my_syland_entitlements() to authenticated;

commit;

-- Kiểm tra dữ liệu sau migration (không làm thay đổi dữ liệu).
select
  (select count(*) from auth.users) as auth_users,
  (select count(*) from public.profiles) as profiles,
  (select count(*) from public.licenses where user_id is not null) as licenses_linked,
  (select count(*) from public.licenses where user_id is null) as legacy_licenses_unlinked,
  (select count(*) from public.profiles where role = 'admin') as admins;
