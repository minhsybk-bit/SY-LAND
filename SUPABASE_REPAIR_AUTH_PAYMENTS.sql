-- SỸ LAND: phục hồi tài khoản, thanh toán và cấp mã bản quyền.
-- Chạy TOÀN BỘ tệp này trong một New query của Supabase SQL Editor.
-- Có thể chạy lại an toàn. Không sửa trạng thái hay nội dung các đơn cũ.

begin;

create extension if not exists pgcrypto;
create sequence if not exists public.payment_order_number start 1001;

create table if not exists public.profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  full_name text not null default '',
  email text not null,
  role text not null default 'user',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.licenses (
  id uuid primary key default gen_random_uuid(),
  code text not null unique,
  customer text not null,
  email text not null,
  plan text not null,
  expires_at timestamptz not null,
  status text not null default 'Hoạt động',
  created_by uuid references auth.users(id),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);
alter table public.licenses add column if not exists max_devices integer not null default 1;
alter table public.licenses add column if not exists seat_count integer not null default 1;
alter table public.licenses add column if not exists max_parcels_per_run integer;

create table if not exists public.payment_orders (
  id uuid primary key default gen_random_uuid(),
  order_code text not null unique default
    ('PAY-' || to_char(now(), 'YYMMDD') || '-' ||
      lpad(nextval('public.payment_order_number')::text, 5, '0')),
  user_id uuid not null default auth.uid() references auth.users(id) on delete restrict,
  email text not null default lower(coalesce(auth.jwt() ->> 'email', '')),
  customer text not null default '',
  plan text not null,
  amount bigint not null,
  duration_months integer not null default 1,
  max_devices integer not null default 1,
  seat_count integer not null default 1,
  transfer_content text unique,
  status text not null default 'Chờ thanh toán',
  license_code text,
  confirmed_by uuid references auth.users(id),
  confirmed_at timestamptz,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);
alter table public.payment_orders add column if not exists seat_count integer not null default 1;

create table if not exists public.payment_settings (
  id text primary key default 'primary',
  bank_bin text not null,
  bank_name text not null,
  account_number text not null,
  account_name text not null,
  support_phone text not null default '',
  updated_by uuid references auth.users(id),
  updated_at timestamptz not null default now()
);

-- Loại bỏ trigger đời cũ trước khi thay hàm. Không UPDATE dữ liệu lịch sử.
drop trigger if exists prepare_payment_order_trigger on public.payment_orders;
drop trigger if exists issue_paid_license_trigger on public.payment_orders;
drop trigger if exists audit_payment_change_trigger on public.payment_orders;
drop trigger if exists touch_payment_settings_trigger on public.payment_settings;

-- Thay ràng buộc nhưng không quét các bản ghi cũ.
alter table public.payment_orders drop constraint if exists payment_orders_plan_check;
alter table public.payment_orders drop constraint if exists payment_orders_amount_check;
alter table public.payment_orders drop constraint if exists payment_orders_duration_months_check;
alter table public.payment_orders drop constraint if exists payment_orders_max_devices_check;
alter table public.payment_orders drop constraint if exists payment_orders_seat_count_check;
alter table public.payment_orders drop constraint if exists payment_orders_status_check;
alter table public.payment_orders
  add constraint payment_orders_plan_check
  check (plan in ('Go', 'Plus', 'Pro', 'Văn phòng')) not valid;
alter table public.payment_orders
  add constraint payment_orders_amount_check check (amount > 0) not valid;
alter table public.payment_orders
  add constraint payment_orders_duration_months_check
  check (duration_months in (1, 6, 12)) not valid;
alter table public.payment_orders
  add constraint payment_orders_max_devices_check
  check (max_devices between 1 and 500) not valid;
alter table public.payment_orders
  add constraint payment_orders_seat_count_check
  check (seat_count between 1 and 500) not valid;
alter table public.payment_orders
  add constraint payment_orders_status_check
  check (status in ('Chờ thanh toán', 'Chờ xác nhận', 'Đã thanh toán', 'Từ chối', 'Đã hủy')) not valid;

create or replace function public.is_syland_admin()
returns boolean language sql stable security definer set search_path = public as $$
  select exists (
    select 1 from public.profiles
    where id = auth.uid() and role = 'admin'
  );
$$;

create or replace function public.handle_new_user()
returns trigger language plpgsql security definer set search_path = public as $$
begin
  insert into public.profiles(id, full_name, email, role)
  values (
    new.id,
    coalesce(nullif(new.raw_user_meta_data ->> 'full_name', ''), split_part(coalesce(new.email, ''), '@', 1)),
    coalesce(new.email, ''),
    'user'
  )
  on conflict (id) do update set
    full_name = excluded.full_name,
    email = excluded.email,
    updated_at = now();
  return new;
end;
$$;
drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
after insert or update of email, raw_user_meta_data on auth.users
for each row execute procedure public.handle_new_user();

insert into public.profiles(id, full_name, email, role)
select id,
  coalesce(nullif(raw_user_meta_data ->> 'full_name', ''), split_part(coalesce(email, ''), '@', 1)),
  coalesce(email, ''),
  case when lower(coalesce(email, '')) = 'minhsybk@gmail.com' then 'admin' else 'user' end
from auth.users
on conflict (id) do update set
  full_name = excluded.full_name,
  email = excluded.email,
  role = case
    when lower(excluded.email) = 'minhsybk@gmail.com' then 'admin'
    else public.profiles.role
  end,
  updated_at = now();

alter table public.profiles enable row level security;
alter table public.licenses enable row level security;
alter table public.payment_orders enable row level security;
alter table public.payment_settings enable row level security;

drop policy if exists "profile_self_read" on public.profiles;
create policy "profile_self_read" on public.profiles for select to authenticated
using (id = auth.uid() or public.is_syland_admin());

drop policy if exists "license_owner_read" on public.licenses;
create policy "license_owner_read" on public.licenses for select to authenticated
using (
  lower(email) = lower(coalesce(auth.jwt() ->> 'email', ''))
  or public.is_syland_admin()
);
drop policy if exists "license_admin_insert" on public.licenses;
create policy "license_admin_insert" on public.licenses for insert to authenticated
with check (public.is_syland_admin());
drop policy if exists "license_admin_update" on public.licenses;
create policy "license_admin_update" on public.licenses for update to authenticated
using (public.is_syland_admin()) with check (public.is_syland_admin());

drop policy if exists "payment_owner_insert" on public.payment_orders;
create policy "payment_owner_insert" on public.payment_orders for insert to authenticated
with check (user_id = auth.uid() and status = 'Chờ thanh toán');
drop policy if exists "payment_owner_read" on public.payment_orders;
create policy "payment_owner_read" on public.payment_orders for select to authenticated
using (user_id = auth.uid() or public.is_syland_admin());
drop policy if exists "payment_owner_notify" on public.payment_orders;
create policy "payment_owner_notify" on public.payment_orders for update to authenticated
using (user_id = auth.uid() and status in ('Chờ thanh toán', 'Chờ xác nhận'))
with check (
  user_id = auth.uid()
  and status in ('Chờ xác nhận', 'Đã hủy')
  and license_code is null
  and confirmed_by is null
);
drop policy if exists "payment_admin_update" on public.payment_orders;
create policy "payment_admin_update" on public.payment_orders for update to authenticated
using (public.is_syland_admin()) with check (public.is_syland_admin());

drop policy if exists "payment_settings_authenticated_read" on public.payment_settings;
create policy "payment_settings_authenticated_read" on public.payment_settings
for select to authenticated using (true);
drop policy if exists "payment_settings_admin_insert" on public.payment_settings;
create policy "payment_settings_admin_insert" on public.payment_settings
for insert to authenticated with check (public.is_syland_admin());
drop policy if exists "payment_settings_admin_update" on public.payment_settings;
create policy "payment_settings_admin_update" on public.payment_settings
for update to authenticated
using (public.is_syland_admin()) with check (public.is_syland_admin());

create or replace function public.prepare_payment_order()
returns trigger language plpgsql security definer set search_path = public as $$
declare v_name text; v_monthly bigint;
begin
  if auth.uid() is null then raise exception 'Hãy đăng nhập trước khi tạo đơn'; end if;
  new.plan := case lower(trim(new.plan))
    when 'go' then 'Go'
    when 'plus' then 'Plus'
    when 'pro' then 'Pro'
    when 'office' then 'Văn phòng'
    when 'van phong' then 'Văn phòng'
    when 'văn phòng' then 'Văn phòng'
    when 'cá nhân' then 'Plus'
    else trim(new.plan)
  end;
  if new.duration_months not in (1, 6, 12) then
    raise exception 'Chu kỳ thanh toán không hợp lệ';
  end if;
  if new.plan in ('Go', 'Plus', 'Pro') then
    new.seat_count := 1;
    new.max_devices := 1;
    v_monthly := case new.plan
      when 'Go' then 99000 when 'Plus' then 199000 else 399000 end;
  elsif new.plan = 'Văn phòng' then
    if new.seat_count not between 2 and 500 then
      raise exception 'Gói Văn phòng cần từ 2 đến 500 tài khoản';
    end if;
    new.max_devices := new.seat_count;
    v_monthly := 298000 * new.seat_count;
  else
    raise exception 'Gói thanh toán không hợp lệ: %', new.plan;
  end if;
  new.amount := round(
    v_monthly * new.duration_months *
    case new.duration_months when 6 then 0.90 when 12 then 0.80 else 1 end
  )::bigint;
  new.user_id := auth.uid();
  new.email := lower(coalesce(auth.jwt() ->> 'email', ''));
  select full_name into v_name from public.profiles where id = auth.uid();
  new.customer := coalesce(nullif(v_name, ''), new.email);
  new.status := 'Chờ thanh toán';
  new.license_code := null;
  new.confirmed_by := null;
  new.confirmed_at := null;
  if coalesce(new.transfer_content, '') = '' then
    new.transfer_content := 'SYLAND ' || new.order_code;
  end if;
  return new;
end;
$$;

create or replace function public.issue_paid_license()
returns trigger language plpgsql security definer set search_path = public as $$
declare v_code text; v_hash text; v_limit integer;
begin
  new.updated_at := now();
  if not public.is_syland_admin() then
    if new.user_id is distinct from old.user_id
      or new.email is distinct from old.email
      or new.customer is distinct from old.customer
      or new.plan is distinct from old.plan
      or new.amount is distinct from old.amount
      or new.duration_months is distinct from old.duration_months
      or new.max_devices is distinct from old.max_devices
      or new.seat_count is distinct from old.seat_count
      or new.transfer_content is distinct from old.transfer_content
      or new.license_code is distinct from old.license_code
      or new.confirmed_by is distinct from old.confirmed_by
      or new.confirmed_at is distinct from old.confirmed_at then
      raise exception 'Người dùng chỉ được thông báo đã chuyển khoản';
    end if;
    if not (
      (old.status = 'Chờ thanh toán' and new.status = 'Chờ xác nhận')
      or (old.status in ('Chờ thanh toán', 'Chờ xác nhận') and new.status = 'Đã hủy')
    ) then
      raise exception 'Chuyển trạng thái thanh toán không hợp lệ';
    end if;
  end if;
  if new.status = 'Đã thanh toán' and old.status is distinct from new.status then
    if not public.is_syland_admin() then
      raise exception 'Chỉ quản trị viên được xác nhận thanh toán';
    end if;
    if new.license_code is not null then return new; end if;
    v_limit := case
      when new.plan = 'Go' then 80
      when new.plan = 'Plus' then 140
      when new.plan = 'Pro' then 200
      when new.plan = 'Văn phòng' and new.seat_count < 5 then 140
      else null
    end;
    v_hash := upper(md5(
      random()::text || clock_timestamp()::text || new.id::text || new.email
    ));
    v_code := 'SYL-' || substr(v_hash, 1, 5) || '-' || substr(v_hash, 6, 5);
    insert into public.licenses(
      code, customer, email, plan, expires_at, status, created_by,
      max_devices, seat_count, max_parcels_per_run
    ) values (
      v_code, new.customer, new.email, new.plan,
      now() + make_interval(months => new.duration_months),
      'Hoạt động', auth.uid(), new.max_devices, new.seat_count, v_limit
    );
    new.license_code := v_code;
    new.confirmed_by := auth.uid();
    new.confirmed_at := now();
  elsif old.status = 'Đã thanh toán' and new.status is distinct from old.status then
    raise exception 'Không thể đổi trạng thái đơn đã cấp bản quyền';
  end if;
  return new;
end;
$$;

create or replace function public.admin_confirm_payment(
  p_order_id uuid, p_status text
)
returns setof public.payment_orders
language plpgsql security definer set search_path = public as $$
begin
  if not public.is_syland_admin() then
    raise exception 'Tài khoản hiện tại chưa có quyền quản trị SỸ LAND';
  end if;
  if p_status not in ('Đã thanh toán', 'Từ chối') then
    raise exception 'Trạng thái xác nhận không hợp lệ';
  end if;
  return query update public.payment_orders
  set status = p_status
  where id = p_order_id and status = 'Chờ xác nhận'
  returning *;
  if not found then
    raise exception 'Không tìm thấy đơn đang chờ xác nhận hoặc đơn đã được xử lý';
  end if;
end;
$$;
revoke all on function public.admin_confirm_payment(uuid, text) from public;
grant execute on function public.admin_confirm_payment(uuid, text) to authenticated;

create or replace function public.cancel_my_payment(p_order_id uuid)
returns setof public.payment_orders
language plpgsql security definer set search_path = public as $$
begin
  if auth.uid() is null then
    raise exception 'Hãy đăng nhập trước khi hủy giao dịch';
  end if;
  return query update public.payment_orders
  set status = 'Đã hủy'
  where id = p_order_id
    and user_id = auth.uid()
    and status in ('Chờ thanh toán', 'Chờ xác nhận')
  returning *;
  if not found then
    raise exception 'Không tìm thấy giao dịch có thể hủy hoặc giao dịch đã được xử lý';
  end if;
end;
$$;
revoke all on function public.cancel_my_payment(uuid) from public;
grant execute on function public.cancel_my_payment(uuid) to authenticated;

create or replace function public.admin_save_payment_settings(
  p_bank_bin text, p_bank_name text, p_account_number text,
  p_account_name text, p_support_phone text default ''
)
returns setof public.payment_settings
language plpgsql security definer set search_path = public as $$
begin
  if auth.uid() is null or not public.is_syland_admin() then
    raise exception 'Tài khoản hiện tại chưa có quyền quản trị SỸ LAND';
  end if;
  if coalesce(trim(p_bank_bin), '') !~ '^[0-9]{6}$' then
    raise exception 'Mã BIN VietQR phải gồm đúng 6 chữ số';
  end if;
  if coalesce(trim(p_account_number), '') !~ '^[0-9]{6,30}$' then
    raise exception 'Số tài khoản chỉ được chứa từ 6 đến 30 chữ số';
  end if;
  return query insert into public.payment_settings(
    id, bank_bin, bank_name, account_number, account_name,
    support_phone, updated_by, updated_at
  ) values (
    'primary', trim(p_bank_bin), trim(p_bank_name),
    trim(p_account_number), upper(trim(p_account_name)),
    left(coalesce(trim(p_support_phone), ''), 30), auth.uid(), now()
  )
  on conflict (id) do update set
    bank_bin = excluded.bank_bin,
    bank_name = excluded.bank_name,
    account_number = excluded.account_number,
    account_name = excluded.account_name,
    support_phone = excluded.support_phone,
    updated_by = auth.uid(),
    updated_at = now()
  returning *;
end;
$$;
revoke all on function public.admin_save_payment_settings(text, text, text, text, text) from public;
grant execute on function public.admin_save_payment_settings(text, text, text, text, text) to authenticated;

create trigger prepare_payment_order_trigger
before insert on public.payment_orders
for each row execute procedure public.prepare_payment_order();
create trigger issue_paid_license_trigger
before update on public.payment_orders
for each row execute procedure public.issue_paid_license();

create index if not exists payment_orders_user_index
on public.payment_orders(user_id, created_at desc);
create index if not exists payment_orders_status_index
on public.payment_orders(status, created_at desc);
create index if not exists licenses_email_index
on public.licenses(lower(email));

commit;

-- Phải trả về một dòng với tất cả cột boolean = true.
select
  to_regprocedure('public.prepare_payment_order()') is not null as create_order_ready,
  to_regprocedure('public.admin_confirm_payment(uuid,text)') is not null as confirm_payment_ready,
  to_regprocedure('public.cancel_my_payment(uuid)') is not null as cancel_payment_ready,
  to_regprocedure('public.admin_save_payment_settings(text,text,text,text,text)') is not null as payment_settings_ready,
  exists(
    select 1 from public.profiles
    where lower(email) = 'minhsybk@gmail.com' and role = 'admin'
  ) as admin_ready,
  exists(
    select 1 from pg_trigger
    where tgname = 'prepare_payment_order_trigger' and not tgisinternal
  ) as create_order_trigger_ready,
  exists(
    select 1 from pg_trigger
    where tgname = 'issue_paid_license_trigger' and not tgisinternal
  ) as license_trigger_ready;
