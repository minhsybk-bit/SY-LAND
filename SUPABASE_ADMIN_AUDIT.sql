-- SỸ LAND: nhật ký quản trị bản quyền và thiết bị
-- Chạy sau SUPABASE_SCHEMA.sql và SUPABASE_DEVICE_LICENSE.sql.

create table if not exists public.audit_events (
  id bigint generated always as identity primary key,
  actor_id uuid references auth.users(id) on delete set null,
  action text not null,
  entity_type text not null,
  entity_id text not null default '',
  details jsonb not null default '{}'::jsonb,
  created_at timestamptz not null default now()
);

create index if not exists audit_events_created_index on public.audit_events(created_at desc);
alter table public.audit_events enable row level security;

drop policy if exists "audit_admin_read" on public.audit_events;
create policy "audit_admin_read" on public.audit_events
for select to authenticated
using (public.is_syland_admin());

create or replace function public.write_syland_audit()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
declare
  v_action text;
  v_id text;
  v_details jsonb;
begin
  v_action := TG_TABLE_NAME || '_' || lower(TG_OP);
  v_id := coalesce((case when TG_OP = 'DELETE' then old.id else new.id end)::text, '');
  if TG_TABLE_NAME = 'licenses' then
    v_details := jsonb_build_object(
      'email', case when TG_OP = 'DELETE' then old.email else new.email end,
      'code', case when TG_OP = 'DELETE' then old.code else new.code end,
      'status', case when TG_OP = 'DELETE' then old.status else new.status end
    );
  else
    v_details := jsonb_build_object(
      'device_name', case when TG_OP = 'DELETE' then old.device_name else new.device_name end,
      'status', case when TG_OP = 'DELETE' then old.status else new.status end,
      'app_version', case when TG_OP = 'DELETE' then old.app_version else new.app_version end
    );
  end if;
  insert into public.audit_events(actor_id, action, entity_type, entity_id, details)
  values (auth.uid(), v_action, TG_TABLE_NAME, v_id, v_details);
  return null;
end;
$$;

drop trigger if exists licenses_audit_trigger on public.licenses;
create trigger licenses_audit_trigger
after insert or update or delete on public.licenses
for each row execute procedure public.write_syland_audit();

drop trigger if exists devices_audit_trigger on public.device_activations;
create trigger devices_audit_trigger
after insert or update or delete on public.device_activations
for each row execute procedure public.write_syland_audit();

-- Thiết bị đã bị admin hủy không được tự kích hoạt lại bằng lần đăng nhập kế tiếp.
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
  if auth.uid() is null then return jsonb_build_object('ok', false, 'message', 'Phiên đăng nhập không hợp lệ.'); end if;
  if lower(coalesce(auth.jwt() ->> 'email', '')) <> lower(trim(coalesce(p_email, ''))) then
    return jsonb_build_object('ok', false, 'message', 'Email đăng nhập không khớp yêu cầu kích hoạt.');
  end if;
  if public.is_syland_admin() then
    return jsonb_build_object('ok', true, 'role', 'admin', 'plan', 'Đơn vị', 'seat_count', 1,
      'max_parcels_per_run', null, 'full_access', true, 'message', 'Tài khoản quản trị được toàn quyền sử dụng.');
  end if;
  select * into v_license from public.licenses
  where lower(email) = lower(trim(p_email)) and status = 'Hoạt động' and expires_at >= now()
  order by expires_at desc limit 1;
  if v_license.id is null then
    return jsonb_build_object('ok', false, 'message', 'Tài khoản chưa có bản quyền hoạt động hoặc bản quyền đã hết hạn.');
  end if;
  select * into v_activation from public.device_activations
  where license_id = v_license.id and device_hash = p_device_hash limit 1;
  if v_activation.id is not null and v_activation.status = 'Đã hủy' then
    return jsonb_build_object('ok', false, 'message', 'Thiết bị này đã bị quản trị viên hủy quyền sử dụng.');
  end if;
  if v_activation.id is not null then
    update public.device_activations set user_id = auth.uid(),
      device_name = left(coalesce(p_device_name, 'Windows PC'), 120),
      app_version = left(coalesce(p_app_version, ''), 40), last_seen_at = now()
    where id = v_activation.id;
    return jsonb_build_object('ok', true, 'license_code', v_license.code, 'plan', v_license.plan,
      'expires_at', v_license.expires_at, 'max_devices', v_license.max_devices,
      'seat_count', v_license.seat_count, 'max_parcels_per_run', v_license.max_parcels_per_run,
      'full_access', (v_license.plan in ('Pro', 'Đơn vị') or (v_license.plan = 'Văn phòng' and v_license.seat_count >= 5)),
      'device_registered', true);
  end if;
  select count(*) into v_active_count from public.device_activations
  where license_id = v_license.id and status = 'Hoạt động';
  if v_active_count >= v_license.max_devices then
    return jsonb_build_object('ok', false, 'message', 'Bản quyền đã đạt giới hạn số thiết bị.',
      'max_devices', v_license.max_devices, 'active_devices', v_active_count);
  end if;
  insert into public.device_activations(license_id, user_id, device_hash, device_name, app_version)
  values (v_license.id, auth.uid(), left(p_device_hash, 128), left(coalesce(p_device_name, 'Windows PC'), 120), left(coalesce(p_app_version, ''), 40));
  return jsonb_build_object('ok', true, 'license_code', v_license.code, 'plan', v_license.plan,
    'expires_at', v_license.expires_at, 'max_devices', v_license.max_devices,
    'seat_count', v_license.seat_count, 'max_parcels_per_run', v_license.max_parcels_per_run,
    'full_access', (v_license.plan in ('Pro', 'Đơn vị') or (v_license.plan = 'Văn phòng' and v_license.seat_count >= 5)),
    'device_registered', true);
end;
$$;

revoke all on function public.activate_syland_device(text, text, text, text) from public;
grant execute on function public.activate_syland_device(text, text, text, text) to authenticated;
