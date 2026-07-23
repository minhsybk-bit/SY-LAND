-- SỸ LAND: vá nhanh thanh toán và cấp bản quyền theo Go/Plus/Pro/Văn phòng.
-- Dùng khi đã chạy SUPABASE_PAYMENTS.sql nhưng xác nhận thanh toán chưa hoạt động.
-- Tệp không chứa email quản trị và không tự nâng quyền tài khoản. Vai trò admin
-- phải được thiết lập riêng trong Supabase trước khi sử dụng chức năng đối soát.

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
create policy "payment_settings_authenticated_read" on public.payment_settings for select to authenticated using (true);
drop policy if exists "payment_settings_admin_insert" on public.payment_settings;
create policy "payment_settings_admin_insert" on public.payment_settings for insert to authenticated with check (public.is_syland_admin());
drop policy if exists "payment_settings_admin_update" on public.payment_settings;
create policy "payment_settings_admin_update" on public.payment_settings for update to authenticated
using (public.is_syland_admin()) with check (public.is_syland_admin());

alter table public.payment_orders add column if not exists seat_count integer not null default 1;
alter table public.licenses add column if not exists seat_count integer not null default 1;
alter table public.licenses add column if not exists max_parcels_per_run integer;
alter table public.payment_orders enable row level security;
drop policy if exists "payment_owner_read" on public.payment_orders;
create policy "payment_owner_read" on public.payment_orders for select to authenticated
using (user_id = auth.uid() or public.is_syland_admin());
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
    v_limit := case when new.plan = 'Go' then 80 when new.plan = 'Plus' then 140
      when new.plan = 'Pro' then 200 when new.plan = 'Văn phòng' and new.seat_count < 5 then 140 else null end;
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

select
  exists(select 1 from public.profiles where role = 'admin') as admin_ready,
  to_regprocedure('public.admin_confirm_payment(uuid,text)') is not null as rpc_ready,
  (select count(*) from public.payment_orders where status = 'Chờ xác nhận') as pending_orders;
