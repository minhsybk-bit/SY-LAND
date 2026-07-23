-- SỸ LAND: đơn chuyển khoản, gói quyền và tự động cấp mã bản quyền.
-- Chạy sau SUPABASE_SCHEMA.sql, SUPABASE_DEVICE_LICENSE.sql và SUPABASE_ADMIN_AUDIT.sql.

create sequence if not exists public.payment_order_number start 1001;

-- Cấu hình nhận tiền được lưu trong Supabase, không đặt trong mã nguồn công khai.
create table if not exists public.payment_settings (
  id text primary key default 'primary' check (id = 'primary'),
  bank_bin text not null check (bank_bin ~ '^[0-9]{6}$'),
  bank_name text not null check (char_length(bank_name) between 2 and 160),
  account_number text not null check (account_number ~ '^[0-9]{6,30}$'),
  account_name text not null check (char_length(account_name) between 2 and 160),
  support_phone text not null default '' check (char_length(support_phone) <= 30),
  updated_by uuid references auth.users(id),
  updated_at timestamptz not null default now()
);
alter table public.payment_settings enable row level security;
drop policy if exists "payment_settings_authenticated_read" on public.payment_settings;
create policy "payment_settings_authenticated_read" on public.payment_settings
for select to authenticated using (true);
drop policy if exists "payment_settings_admin_insert" on public.payment_settings;
create policy "payment_settings_admin_insert" on public.payment_settings
for insert to authenticated with check (public.is_syland_admin());
drop policy if exists "payment_settings_admin_update" on public.payment_settings;
create policy "payment_settings_admin_update" on public.payment_settings
for update to authenticated using (public.is_syland_admin()) with check (public.is_syland_admin());
create or replace function public.touch_payment_settings()
returns trigger language plpgsql security definer set search_path = public as $$
begin
  if not public.is_syland_admin() then raise exception 'Chỉ quản trị viên được đổi cấu hình nhận tiền'; end if;
  new.updated_by := auth.uid(); new.updated_at := now(); return new;
end; $$;
drop trigger if exists touch_payment_settings_trigger on public.payment_settings;
create trigger touch_payment_settings_trigger before insert or update on public.payment_settings
for each row execute procedure public.touch_payment_settings();

create table if not exists public.payment_orders (
  id uuid primary key default gen_random_uuid(),
  order_code text not null unique default ('PAY-' || to_char(now(), 'YYMMDD') || '-' || lpad(nextval('public.payment_order_number')::text, 5, '0')),
  user_id uuid not null default auth.uid() references auth.users(id) on delete restrict,
  email text not null default lower(coalesce(auth.jwt() ->> 'email', '')),
  customer text not null default '',
  plan text not null,
  amount bigint not null,
  duration_months integer not null default 1,
  max_devices integer not null default 1,
  seat_count integer not null default 1,
  transfer_content text unique,
  status text not null default 'Chờ thanh toán' check (status in ('Chờ thanh toán', 'Chờ xác nhận', 'Đã thanh toán', 'Từ chối', 'Đã hủy')),
  license_code text,
  confirmed_by uuid references auth.users(id),
  confirmed_at timestamptz,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);
alter table public.payment_orders add column if not exists seat_count integer not null default 1;
alter table public.payment_orders drop constraint if exists payment_orders_plan_check;
alter table public.payment_orders drop constraint if exists payment_orders_amount_check;
alter table public.payment_orders drop constraint if exists payment_orders_duration_months_check;
alter table public.payment_orders drop constraint if exists payment_orders_max_devices_check;
alter table public.payment_orders drop constraint if exists payment_orders_seat_count_check;
update public.payment_orders set plan = 'Plus' where plan = 'Cá nhân';
update public.payment_orders set seat_count = greatest(2, max_devices)
where plan = 'Văn phòng' and seat_count < 2;
alter table public.payment_orders add constraint payment_orders_plan_check check (plan in ('Go', 'Plus', 'Pro', 'Văn phòng'));
alter table public.payment_orders add constraint payment_orders_amount_check check (amount > 0);
alter table public.payment_orders add constraint payment_orders_duration_months_check check (duration_months in (1, 6, 12));
alter table public.payment_orders add constraint payment_orders_max_devices_check check (max_devices between 1 and 500);
alter table public.payment_orders add constraint payment_orders_seat_count_check check (seat_count between 1 and 500);
create index if not exists payment_orders_user_index on public.payment_orders(user_id, created_at desc);
create index if not exists payment_orders_status_index on public.payment_orders(status, created_at desc);
alter table public.payment_orders enable row level security;

create or replace function public.prepare_payment_order()
returns trigger language plpgsql security definer set search_path = public as $$
declare v_name text; v_monthly bigint;
begin
  if auth.uid() is null then raise exception 'Hãy đăng nhập trước khi tạo đơn'; end if;
  if new.duration_months not in (1, 6, 12) then raise exception 'Chu kỳ thanh toán không hợp lệ'; end if;
  if new.plan in ('Go', 'Plus', 'Pro') then
    new.seat_count := 1; new.max_devices := 1;
    v_monthly := case new.plan when 'Go' then 99000 when 'Plus' then 199000 else 399000 end;
  elsif new.plan = 'Văn phòng' then
    if new.seat_count not between 2 and 500 then raise exception 'Gói Văn phòng cần từ 2 đến 500 tài khoản'; end if;
    new.max_devices := new.seat_count;
    v_monthly := 298000 * new.seat_count;
  else
    raise exception 'Gói thanh toán không hợp lệ';
  end if;
  new.amount := round(v_monthly * new.duration_months *
    case new.duration_months when 6 then 0.90 when 12 then 0.80 else 1 end)::bigint;
  new.user_id := auth.uid();
  new.email := lower(coalesce(auth.jwt() ->> 'email', ''));
  select full_name into v_name from public.profiles where id = auth.uid();
  new.customer := coalesce(nullif(v_name, ''), new.email);
  new.status := 'Chờ thanh toán';
  new.license_code := null; new.confirmed_by := null; new.confirmed_at := null;
  if new.transfer_content is null or new.transfer_content = '' then new.transfer_content := 'SYLAND ' || new.order_code; end if;
  return new;
end; $$;
drop trigger if exists prepare_payment_order_trigger on public.payment_orders;
create trigger prepare_payment_order_trigger before insert on public.payment_orders
for each row execute procedure public.prepare_payment_order();

drop policy if exists "payment_owner_insert" on public.payment_orders;
create policy "payment_owner_insert" on public.payment_orders for insert to authenticated
with check (user_id = auth.uid() and status = 'Chờ thanh toán');
drop policy if exists "payment_owner_read" on public.payment_orders;
create policy "payment_owner_read" on public.payment_orders for select to authenticated
using (user_id = auth.uid() or public.is_syland_admin());
drop policy if exists "payment_owner_notify" on public.payment_orders;
create policy "payment_owner_notify" on public.payment_orders for update to authenticated
using (user_id = auth.uid() and status = 'Chờ thanh toán')
with check (user_id = auth.uid() and status = 'Chờ xác nhận' and license_code is null and confirmed_by is null);
drop policy if exists "payment_admin_update" on public.payment_orders;
create policy "payment_admin_update" on public.payment_orders for update to authenticated
using (public.is_syland_admin()) with check (public.is_syland_admin());

create or replace function public.issue_paid_license()
returns trigger language plpgsql security definer set search_path = public as $$
declare v_code text; v_hash text; v_limit integer;
begin
  new.updated_at := now();
  if not public.is_syland_admin() then
    if new.user_id is distinct from old.user_id or new.email is distinct from old.email
      or new.customer is distinct from old.customer or new.plan is distinct from old.plan
      or new.amount is distinct from old.amount or new.duration_months is distinct from old.duration_months
      or new.max_devices is distinct from old.max_devices or new.seat_count is distinct from old.seat_count
      or new.transfer_content is distinct from old.transfer_content or new.license_code is distinct from old.license_code
      or new.confirmed_by is distinct from old.confirmed_by or new.confirmed_at is distinct from old.confirmed_at then
      raise exception 'Người dùng chỉ được thông báo đã chuyển khoản';
    end if;
    if old.status <> 'Chờ thanh toán' or new.status <> 'Chờ xác nhận' then
      raise exception 'Chuyển trạng thái thanh toán không hợp lệ';
    end if;
  end if;
  if new.status = 'Đã thanh toán' and old.status is distinct from new.status then
    if not public.is_syland_admin() then raise exception 'Chỉ quản trị viên được xác nhận thanh toán'; end if;
    if new.license_code is not null then return new; end if;
    v_limit := case
      when new.plan = 'Go' then 80
      when new.plan = 'Plus' then 140
      when new.plan = 'Pro' then 200
      when new.plan = 'Văn phòng' and new.seat_count < 5 then 140
      else null
    end;
    v_hash := upper(md5(random()::text || clock_timestamp()::text || new.id::text || new.email));
    v_code := 'SYL-' || substr(v_hash, 1, 5) || '-' || substr(v_hash, 6, 5);
    insert into public.licenses
      (code, customer, email, plan, expires_at, status, created_by, max_devices, seat_count, max_parcels_per_run)
    values
      (v_code, new.customer, new.email, new.plan, now() + make_interval(months => new.duration_months),
       'Hoạt động', auth.uid(), new.max_devices, new.seat_count, v_limit);
    new.license_code := v_code; new.confirmed_by := auth.uid(); new.confirmed_at := now();
  elsif old.status = 'Đã thanh toán' and new.status is distinct from old.status then
    raise exception 'Không thể đổi trạng thái đơn đã cấp bản quyền';
  end if;
  return new;
end; $$;
drop trigger if exists issue_paid_license_trigger on public.payment_orders;
create trigger issue_paid_license_trigger before update on public.payment_orders
for each row execute procedure public.issue_paid_license();

create or replace function public.admin_confirm_payment(p_order_id uuid, p_status text)
returns setof public.payment_orders language plpgsql security definer set search_path = public as $$
begin
  if not public.is_syland_admin() then raise exception 'Tài khoản hiện tại chưa có quyền quản trị SỸ LAND'; end if;
  if p_status not in ('Đã thanh toán', 'Từ chối') then raise exception 'Trạng thái xác nhận không hợp lệ'; end if;
  return query update public.payment_orders set status = p_status
    where id = p_order_id and status = 'Chờ xác nhận' returning *;
  if not found then raise exception 'Không tìm thấy đơn đang chờ xác nhận hoặc đơn đã được xử lý'; end if;
end; $$;
revoke all on function public.admin_confirm_payment(uuid, text) from public;
grant execute on function public.admin_confirm_payment(uuid, text) to authenticated;

create or replace function public.audit_payment_change()
returns trigger language plpgsql security definer set search_path = public as $$
begin
  if old.status is distinct from new.status then
    insert into public.audit_events(actor_id, action, entity_type, entity_id, details)
    values (auth.uid(), 'payment_status_changed', 'payment_orders', new.id::text,
      jsonb_build_object('order_code', new.order_code, 'from', old.status, 'to', new.status, 'license_code', new.license_code));
  end if;
  return null;
end; $$;
drop trigger if exists audit_payment_change_trigger on public.payment_orders;
create trigger audit_payment_change_trigger after update on public.payment_orders
for each row execute procedure public.audit_payment_change();
