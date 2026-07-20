-- SỸ LAND: quyền dữ liệu cá nhân và yêu cầu xóa tài khoản
-- Chạy sau SUPABASE_SCHEMA.sql và SUPABASE_ADMIN_AUDIT.sql.

create sequence if not exists public.privacy_request_number start 1001;

create table if not exists public.privacy_requests (
  id uuid primary key default gen_random_uuid(),
  request_code text not null unique default ('PRV-' || to_char(now(), 'YYMM') || '-' || lpad(nextval('public.privacy_request_number')::text, 5, '0')),
  user_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
  request_type text not null check (request_type in ('Xuất dữ liệu', 'Xóa tài khoản')),
  note text not null default '' check (char_length(note) <= 1000),
  admin_note text not null default '' check (char_length(admin_note) <= 2000),
  status text not null default 'Mới' check (status in ('Mới', 'Đang xác minh', 'Đang xử lý', 'Hoàn tất', 'Từ chối')),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists privacy_requests_user_index on public.privacy_requests(user_id, created_at desc);
create index if not exists privacy_requests_status_index on public.privacy_requests(status, created_at desc);
alter table public.privacy_requests enable row level security;

drop policy if exists "privacy_owner_insert" on public.privacy_requests;
create policy "privacy_owner_insert" on public.privacy_requests for insert to authenticated
with check (user_id = auth.uid() and status = 'Mới' and admin_note = '');

drop policy if exists "privacy_owner_read" on public.privacy_requests;
create policy "privacy_owner_read" on public.privacy_requests for select to authenticated
using (user_id = auth.uid() or public.is_syland_admin());

drop policy if exists "privacy_admin_update" on public.privacy_requests;
create policy "privacy_admin_update" on public.privacy_requests for update to authenticated
using (public.is_syland_admin()) with check (public.is_syland_admin());

create or replace function public.set_privacy_request_updated_at()
returns trigger language plpgsql as $$ begin new.updated_at = now(); return new; end; $$;
drop trigger if exists privacy_request_updated_at_trigger on public.privacy_requests;
create trigger privacy_request_updated_at_trigger before update on public.privacy_requests
for each row execute procedure public.set_privacy_request_updated_at();

create or replace function public.write_privacy_request_audit()
returns trigger language plpgsql security definer set search_path = public as $$
begin
  if old.status is distinct from new.status then
    insert into public.audit_events(actor_id, action, entity_type, entity_id, details)
    values (auth.uid(), 'privacy_request_status_changed', 'privacy_requests', new.id::text,
      jsonb_build_object('request_code', new.request_code, 'from', old.status, 'to', new.status));
  end if;
  return null;
end; $$;
drop trigger if exists privacy_request_audit_trigger on public.privacy_requests;
create trigger privacy_request_audit_trigger after update on public.privacy_requests
for each row execute procedure public.write_privacy_request_audit();

-- Không tạo hàm xóa auth.users ở trình duyệt. Bước xóa cuối cùng phải chạy ở môi
-- trường quản trị an toàn có service_role, sau khi đã xác minh đúng chủ tài khoản.
