-- SỸ LAND: đơn chuyển khoản và tự động cấp mã bản quyền
-- Chạy sau SUPABASE_SCHEMA.sql, SUPABASE_DEVICE_LICENSE.sql và SUPABASE_ADMIN_AUDIT.sql.

create sequence if not exists public.payment_order_number start 1001;
create table if not exists public.payment_orders (
  id uuid primary key default gen_random_uuid(),
  order_code text not null unique default ('PAY-' || to_char(now(), 'YYMMDD') || '-' || lpad(nextval('public.payment_order_number')::text, 5, '0')),
  user_id uuid not null default auth.uid() references auth.users(id) on delete restrict,
  email text not null default lower(coalesce(auth.jwt() ->> 'email', '')),
  customer text not null default '',
  plan text not null check (plan in ('Cá nhân', 'Văn phòng')),
  amount bigint not null check (amount in (199000, 1490000)),
  duration_months integer not null default 1 check (duration_months between 1 and 24),
  max_devices integer not null check (max_devices between 1 and 5),
  transfer_content text unique,
  status text not null default 'Chờ thanh toán' check (status in ('Chờ thanh toán', 'Chờ xác nhận', 'Đã thanh toán', 'Từ chối', 'Đã hủy')),
  license_code text,
  confirmed_by uuid references auth.users(id),
  confirmed_at timestamptz,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);
create index if not exists payment_orders_user_index on public.payment_orders(user_id, created_at desc);
create index if not exists payment_orders_status_index on public.payment_orders(status, created_at desc);
alter table public.payment_orders enable row level security;

create or replace function public.prepare_payment_order() returns trigger language plpgsql security definer set search_path = public as $$
declare v_name text;
begin
  if new.plan = 'Cá nhân' then new.amount := 199000; new.max_devices := 1;
  elsif new.plan = 'Văn phòng' then new.amount := 1490000; new.max_devices := 5;
  else raise exception 'Gói thanh toán không hợp lệ'; end if;
  new.user_id := auth.uid(); new.email := lower(coalesce(auth.jwt() ->> 'email', ''));
  select full_name into v_name from public.profiles where id = auth.uid(); new.customer := coalesce(v_name, new.email);
  if new.transfer_content is null or new.transfer_content = '' then new.transfer_content := 'SYLAND ' || new.order_code; end if;
  return new;
end; $$;
drop trigger if exists prepare_payment_order_trigger on public.payment_orders;
create trigger prepare_payment_order_trigger before insert on public.payment_orders for each row execute procedure public.prepare_payment_order();

drop policy if exists "payment_owner_insert" on public.payment_orders;
create policy "payment_owner_insert" on public.payment_orders for insert to authenticated with check (user_id = auth.uid() and status = 'Chờ thanh toán');
drop policy if exists "payment_owner_read" on public.payment_orders;
create policy "payment_owner_read" on public.payment_orders for select to authenticated using (user_id = auth.uid() or public.is_syland_admin());
drop policy if exists "payment_owner_notify" on public.payment_orders;
create policy "payment_owner_notify" on public.payment_orders for update to authenticated using (user_id = auth.uid() and status = 'Chờ thanh toán') with check (user_id = auth.uid() and status = 'Chờ xác nhận' and license_code is null and confirmed_by is null);
drop policy if exists "payment_admin_update" on public.payment_orders;
create policy "payment_admin_update" on public.payment_orders for update to authenticated using (public.is_syland_admin()) with check (public.is_syland_admin());

create or replace function public.issue_paid_license() returns trigger language plpgsql security definer set search_path = public as $$
declare v_code text;
begin
  new.updated_at := now();
  if not public.is_syland_admin() then
    if new.user_id is distinct from old.user_id or new.email is distinct from old.email
      or new.customer is distinct from old.customer or new.plan is distinct from old.plan
      or new.amount is distinct from old.amount or new.duration_months is distinct from old.duration_months
      or new.max_devices is distinct from old.max_devices or new.transfer_content is distinct from old.transfer_content
      or new.license_code is distinct from old.license_code or new.confirmed_by is distinct from old.confirmed_by
      or new.confirmed_at is distinct from old.confirmed_at then
      raise exception 'Người dùng chỉ được thông báo đã chuyển khoản';
    end if;
    if old.status <> 'Chờ thanh toán' or new.status <> 'Chờ xác nhận' then
      raise exception 'Chuyển trạng thái thanh toán không hợp lệ';
    end if;
  end if;
  if new.status = 'Đã thanh toán' and old.status is distinct from new.status then
    if not public.is_syland_admin() then raise exception 'Chỉ quản trị viên được xác nhận thanh toán'; end if;
    if new.license_code is not null then return new; end if;
    v_code := 'SYL-' || upper(substr(encode(gen_random_bytes(8), 'hex'), 1, 5)) || '-' || upper(substr(encode(gen_random_bytes(8), 'hex'), 1, 5));
    insert into public.licenses(code, customer, email, plan, expires_at, status, created_by, max_devices)
    values (v_code, new.customer, new.email, new.plan, now() + make_interval(months => new.duration_months), 'Hoạt động', auth.uid(), new.max_devices);
    new.license_code := v_code; new.confirmed_by := auth.uid(); new.confirmed_at := now();
  elsif old.status = 'Đã thanh toán' and new.status is distinct from old.status then
    raise exception 'Không thể đổi trạng thái đơn đã cấp bản quyền';
  end if;
  return new;
end; $$;
drop trigger if exists issue_paid_license_trigger on public.payment_orders;
create trigger issue_paid_license_trigger before update on public.payment_orders for each row execute procedure public.issue_paid_license();

create or replace function public.audit_payment_change() returns trigger language plpgsql security definer set search_path = public as $$
begin
  if old.status is distinct from new.status then insert into public.audit_events(actor_id, action, entity_type, entity_id, details) values (auth.uid(), 'payment_status_changed', 'payment_orders', new.id::text, jsonb_build_object('order_code', new.order_code, 'from', old.status, 'to', new.status, 'license_code', new.license_code)); end if; return null;
end; $$;
drop trigger if exists audit_payment_change_trigger on public.payment_orders;
create trigger audit_payment_change_trigger after update on public.payment_orders for each row execute procedure public.audit_payment_change();
