-- SỸ LAND: BẢN VÁ THANH TOÁN MỘT LẦN
-- Chạy toàn bộ tệp trong Supabase Dashboard > SQL Editor.

-- 1. Bảo đảm tài khoản chủ hệ thống có quyền quản trị.
insert into public.profiles (id, full_name, email, role)
select id, coalesce(nullif(raw_user_meta_data ->> 'full_name', ''), 'Nguyễn Minh Sỹ'), coalesce(email, ''), 'admin'
from auth.users
where lower(email) = lower('minhsybk@gmail.com')
on conflict (id) do update set role = 'admin', email = excluded.email, updated_at = now();

-- 2. Cài lại chính sách đọc và cập nhật đơn cho admin.
alter table public.payment_orders enable row level security;
drop policy if exists "payment_owner_read" on public.payment_orders;
create policy "payment_owner_read" on public.payment_orders for select to authenticated
using (user_id = auth.uid() or public.is_syland_admin());
drop policy if exists "payment_admin_update" on public.payment_orders;
create policy "payment_admin_update" on public.payment_orders for update to authenticated
using (public.is_syland_admin()) with check (public.is_syland_admin());

-- 3. Sinh mã không phụ thuộc extension pgcrypto/gen_random_bytes.
create or replace function public.issue_paid_license()
returns trigger language plpgsql security definer set search_path = public as $$
declare v_code text; v_hash text;
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
    v_hash := upper(md5(random()::text || clock_timestamp()::text || new.id::text || new.email));
    v_code := 'SYL-' || substr(v_hash, 1, 5) || '-' || substr(v_hash, 6, 5);
    insert into public.licenses(code, customer, email, plan, expires_at, status, created_by, max_devices)
    values (v_code, new.customer, new.email, new.plan, now() + make_interval(months => new.duration_months), 'Hoạt động', auth.uid(), new.max_devices);
    new.license_code := v_code;
    new.confirmed_by := auth.uid();
    new.confirmed_at := now();
  elsif old.status = 'Đã thanh toán' and new.status is distinct from old.status then
    raise exception 'Không thể đổi trạng thái đơn đã cấp bản quyền';
  end if;
  return new;
end;
$$;
drop trigger if exists issue_paid_license_trigger on public.payment_orders;
create trigger issue_paid_license_trigger before update on public.payment_orders
for each row execute procedure public.issue_paid_license();

-- 4. RPC duy nhất để website xác nhận hoặc từ chối đơn.
create or replace function public.admin_confirm_payment(p_order_id uuid, p_status text)
returns setof public.payment_orders language plpgsql security definer set search_path = public as $$
begin
  if not public.is_syland_admin() then raise exception 'Tài khoản hiện tại chưa có quyền quản trị SỸ LAND'; end if;
  if p_status not in ('Đã thanh toán', 'Từ chối') then raise exception 'Trạng thái xác nhận không hợp lệ'; end if;
  return query update public.payment_orders set status = p_status
    where id = p_order_id and status = 'Chờ xác nhận' returning *;
  if not found then raise exception 'Không tìm thấy đơn đang chờ xác nhận hoặc đơn đã được xử lý'; end if;
end;
$$;
revoke all on function public.admin_confirm_payment(uuid, text) from public;
grant execute on function public.admin_confirm_payment(uuid, text) to authenticated;

-- 5. Kết quả kiểm tra: admin phải là true và rpc_ready phải là true.
select
  exists(select 1 from public.profiles where lower(email) = lower('minhsybk@gmail.com') and role = 'admin') as admin_ready,
  to_regprocedure('public.admin_confirm_payment(uuid,text)') is not null as rpc_ready,
  (select count(*) from public.payment_orders where status = 'Chờ xác nhận') as pending_orders;
