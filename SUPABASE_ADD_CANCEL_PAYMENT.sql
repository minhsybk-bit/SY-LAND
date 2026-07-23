-- SỸ LAND - bổ sung hủy giao dịch an toàn.
-- Chạy toàn bộ tệp này đúng 1 lần trong Supabase SQL Editor.
-- Không xóa đơn, mã bản quyền hoặc dữ liệu thanh toán cũ.

begin;

drop policy if exists "payment_owner_notify" on public.payment_orders;
create policy "payment_owner_notify" on public.payment_orders for update to authenticated
using (user_id = auth.uid() and status in ('Chờ thanh toán', 'Chờ xác nhận'))
with check (
  user_id = auth.uid()
  and status in ('Chờ xác nhận', 'Đã hủy')
  and license_code is null
  and confirmed_by is null
);

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
      raise exception 'Người dùng chỉ được thông báo đã chuyển khoản hoặc hủy giao dịch';
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
    v_hash := upper(md5(random()::text || clock_timestamp()::text || new.id::text || new.email));
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

commit;

-- Kết quả phải là true.
select to_regprocedure('public.cancel_my_payment(uuid)') is not null as cancel_payment_ready;
