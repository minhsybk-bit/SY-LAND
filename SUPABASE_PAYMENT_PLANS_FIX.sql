-- SỸ LAND: sửa lỗi "Gói thanh toán không hợp lệ".
-- Chạy TOÀN BỘ tệp này một lần trong Supabase > SQL Editor.
-- Bản vá giữ nguyên đơn cũ, chuẩn hóa tên gói và thay thủ tục tạo đơn đời cũ.

begin;

alter table public.payment_orders
  add column if not exists seat_count integer not null default 1;

alter table public.payment_orders drop constraint if exists payment_orders_plan_check;
alter table public.payment_orders drop constraint if exists payment_orders_amount_check;
alter table public.payment_orders drop constraint if exists payment_orders_duration_months_check;
alter table public.payment_orders drop constraint if exists payment_orders_max_devices_check;
alter table public.payment_orders drop constraint if exists payment_orders_seat_count_check;

update public.payment_orders
set plan = case lower(trim(plan))
  when 'cá nhân' then 'Plus'
  when 'go' then 'Go'
  when 'plus' then 'Plus'
  when 'pro' then 'Pro'
  when 'office' then 'Văn phòng'
  when 'van phong' then 'Văn phòng'
  when 'văn phòng' then 'Văn phòng'
  else plan
end;

update public.payment_orders
set seat_count = greatest(2, max_devices)
where plan = 'Văn phòng' and seat_count < 2;

alter table public.payment_orders
  add constraint payment_orders_plan_check
  check (plan in ('Go', 'Plus', 'Pro', 'Văn phòng'));
alter table public.payment_orders
  add constraint payment_orders_amount_check check (amount > 0);
alter table public.payment_orders
  add constraint payment_orders_duration_months_check
  check (duration_months in (1, 6, 12));
alter table public.payment_orders
  add constraint payment_orders_max_devices_check
  check (max_devices between 1 and 500);
alter table public.payment_orders
  add constraint payment_orders_seat_count_check
  check (seat_count between 1 and 500);

create or replace function public.prepare_payment_order()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
declare
  v_name text;
  v_monthly bigint;
begin
  if auth.uid() is null then
    raise exception 'Hãy đăng nhập trước khi tạo đơn';
  end if;

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
      when 'Go' then 99000
      when 'Plus' then 199000
      else 399000
    end;
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
  if new.transfer_content is null or new.transfer_content = '' then
    new.transfer_content := 'SYLAND ' || new.order_code;
  end if;
  return new;
end;
$$;

drop trigger if exists prepare_payment_order_trigger on public.payment_orders;
create trigger prepare_payment_order_trigger
before insert on public.payment_orders
for each row execute procedure public.prepare_payment_order();

commit;

-- Kết quả phải trả về 4 dòng: Go, Plus, Pro, Văn phòng.
select unnest(array['Go', 'Plus', 'Pro', 'Văn phòng']) as goi_da_kich_hoat;
