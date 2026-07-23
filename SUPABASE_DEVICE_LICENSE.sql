-- SỸ LAND: quản lý số thiết bị sử dụng bản quyền
-- Chạy sau SUPABASE_SCHEMA.sql trong Supabase SQL Editor.

alter table public.licenses
  add column if not exists max_devices integer not null default 1
  check (max_devices between 1 and 100);

create table if not exists public.device_activations (
  id uuid primary key default gen_random_uuid(),
  license_id uuid not null references public.licenses(id) on delete cascade,
  user_id uuid not null references auth.users(id) on delete cascade,
  device_hash text not null,
  device_name text not null default 'Windows PC',
  app_version text not null default '',
  status text not null default 'Hoạt động' check (status in ('Hoạt động', 'Đã hủy')),
  first_seen_at timestamptz not null default now(),
  last_seen_at timestamptz not null default now(),
  unique (license_id, device_hash)
);

create index if not exists device_activations_user_index on public.device_activations(user_id);
alter table public.device_activations enable row level security;

drop policy if exists "device_owner_read" on public.device_activations;
create policy "device_owner_read" on public.device_activations
for select to authenticated
using (user_id = auth.uid() or public.is_syland_admin());

drop policy if exists "device_admin_update" on public.device_activations;
create policy "device_admin_update" on public.device_activations
for update to authenticated
using (public.is_syland_admin())
with check (public.is_syland_admin());

create or replace function public.activate_syland_device(
  p_email text,
  p_device_hash text,
  p_device_name text default 'Windows PC',
  p_app_version text default ''
)
returns jsonb
language plpgsql
security definer
set search_path = public
as $$
declare
  v_license public.licenses%rowtype;
  v_active_count integer;
  v_activation public.device_activations%rowtype;
begin
  if auth.uid() is null then
    return jsonb_build_object('ok', false, 'message', 'Phiên đăng nhập không hợp lệ.');
  end if;

  if lower(coalesce(auth.jwt() ->> 'email', '')) <> lower(trim(coalesce(p_email, ''))) then
    return jsonb_build_object('ok', false, 'message', 'Email đăng nhập không khớp yêu cầu kích hoạt.');
  end if;

  if public.is_syland_admin() then
    return jsonb_build_object('ok', true, 'role', 'admin', 'plan', 'Đơn vị', 'seat_count', 1,
      'feature_percent', 100, 'max_parcels_per_run', null, 'max_file_size_mb', 100,
      'max_total_upload_mb', 2000, 'max_uses_per_day', null, 'full_access', true,
      'recommended_upgrade', null, 'message', 'Tài khoản quản trị được toàn quyền sử dụng.');
  end if;

  select * into v_license
  from public.licenses
  where lower(email) = lower(trim(p_email))
    and status = 'Hoạt động'
    and expires_at >= now()
  order by expires_at desc
  limit 1;

  if v_license.id is null then
    return jsonb_build_object('ok', false, 'message', 'Tài khoản chưa có bản quyền hoạt động hoặc bản quyền đã hết hạn.');
  end if;

  select * into v_activation
  from public.device_activations
  where license_id = v_license.id and device_hash = p_device_hash
  limit 1;

  if v_activation.id is not null then
    update public.device_activations
    set user_id = auth.uid(), device_name = left(coalesce(p_device_name, 'Windows PC'), 120),
        app_version = left(coalesce(p_app_version, ''), 40), status = 'Hoạt động', last_seen_at = now()
    where id = v_activation.id;
    return jsonb_build_object('ok', true, 'license_code', v_license.code, 'plan', v_license.plan,
      'expires_at', v_license.expires_at, 'max_devices', v_license.max_devices,
      'seat_count', v_license.seat_count, 'max_parcels_per_run', v_license.max_parcels_per_run,
      'feature_percent', case when v_license.plan = 'Go' then 40 when v_license.plan = 'Plus' or (v_license.plan = 'Văn phòng' and v_license.seat_count < 5) then 70 else 100 end,
      'max_file_size_mb', case when v_license.plan = 'Go' then 20 when v_license.plan = 'Plus' or (v_license.plan = 'Văn phòng' and v_license.seat_count < 5) then 30 when v_license.plan in ('Đơn vị') or (v_license.plan = 'Văn phòng' and v_license.seat_count >= 5) then 100 else 50 end,
      'max_total_upload_mb', case when v_license.plan = 'Go' then 500 when v_license.plan = 'Plus' or (v_license.plan = 'Văn phòng' and v_license.seat_count < 5) then 750 when v_license.plan in ('Đơn vị') or (v_license.plan = 'Văn phòng' and v_license.seat_count >= 5) then 2000 else 1000 end,
      'max_uses_per_day', null, 'recommended_upgrade',
        case when v_license.plan = 'Go' then 'Plus' when v_license.plan = 'Plus' then 'Pro' else null end,
      'full_access', (v_license.plan in ('Pro', 'Đơn vị') or (v_license.plan = 'Văn phòng' and v_license.seat_count >= 5)),
      'device_registered', true);
  end if;

  select count(*) into v_active_count
  from public.device_activations
  where license_id = v_license.id and status = 'Hoạt động';

  if v_active_count >= v_license.max_devices then
    return jsonb_build_object('ok', false, 'message', 'Bản quyền đã đạt giới hạn số thiết bị.',
      'max_devices', v_license.max_devices, 'active_devices', v_active_count);
  end if;

  insert into public.device_activations
    (license_id, user_id, device_hash, device_name, app_version)
  values
    (v_license.id, auth.uid(), left(p_device_hash, 128), left(coalesce(p_device_name, 'Windows PC'), 120), left(coalesce(p_app_version, ''), 40));

  return jsonb_build_object('ok', true, 'license_code', v_license.code, 'plan', v_license.plan,
    'expires_at', v_license.expires_at, 'max_devices', v_license.max_devices,
    'seat_count', v_license.seat_count, 'max_parcels_per_run', v_license.max_parcels_per_run,
    'feature_percent', case when v_license.plan = 'Go' then 40 when v_license.plan = 'Plus' or (v_license.plan = 'Văn phòng' and v_license.seat_count < 5) then 70 else 100 end,
    'max_file_size_mb', case when v_license.plan = 'Go' then 20 when v_license.plan = 'Plus' or (v_license.plan = 'Văn phòng' and v_license.seat_count < 5) then 30 when v_license.plan in ('Đơn vị') or (v_license.plan = 'Văn phòng' and v_license.seat_count >= 5) then 100 else 50 end,
    'max_total_upload_mb', case when v_license.plan = 'Go' then 500 when v_license.plan = 'Plus' or (v_license.plan = 'Văn phòng' and v_license.seat_count < 5) then 750 when v_license.plan in ('Đơn vị') or (v_license.plan = 'Văn phòng' and v_license.seat_count >= 5) then 2000 else 1000 end,
    'max_uses_per_day', null, 'recommended_upgrade',
      case when v_license.plan = 'Go' then 'Plus' when v_license.plan = 'Plus' then 'Pro' else null end,
    'full_access', (v_license.plan in ('Pro', 'Đơn vị') or (v_license.plan = 'Văn phòng' and v_license.seat_count >= 5)),
    'device_registered', true);
end;
$$;

revoke all on function public.activate_syland_device(text, text, text, text) from public;
grant execute on function public.activate_syland_device(text, text, text, text) to authenticated;
