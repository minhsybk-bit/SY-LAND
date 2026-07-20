-- SỸ LAND: trung tâm hỗ trợ sau bán hàng
-- Chạy sau các tệp schema tài khoản và audit.

create sequence if not exists public.support_ticket_number start 1001;

create table if not exists public.support_tickets (
  id uuid primary key default gen_random_uuid(),
  ticket_code text not null unique default ('SYL-' || to_char(now(), 'YYMM') || '-' || lpad(nextval('public.support_ticket_number')::text, 5, '0')),
  user_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
  title text not null check (char_length(title) between 5 and 250),
  category text not null,
  priority text not null default 'Bình thường' check (priority in ('Thấp', 'Bình thường', 'Cao', 'Khẩn cấp')),
  details text not null check (char_length(details) between 10 and 10000),
  app_version text not null default '',
  status text not null default 'Mới' check (status in ('Mới', 'Đang xử lý', 'Chờ người dùng', 'Đã giải quyết', 'Đã đóng')),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists support_user_index on public.support_tickets(user_id);
create index if not exists support_status_index on public.support_tickets(status, created_at desc);
alter table public.support_tickets enable row level security;

drop policy if exists "support_owner_insert" on public.support_tickets;
create policy "support_owner_insert" on public.support_tickets
for insert to authenticated
with check (user_id = auth.uid() and status = 'Mới');

drop policy if exists "support_owner_read" on public.support_tickets;
create policy "support_owner_read" on public.support_tickets
for select to authenticated
using (user_id = auth.uid() or public.is_syland_admin());

drop policy if exists "support_admin_update" on public.support_tickets;
create policy "support_admin_update" on public.support_tickets
for update to authenticated
using (public.is_syland_admin())
with check (public.is_syland_admin());

create or replace function public.set_support_updated_at()
returns trigger language plpgsql as $$
begin new.updated_at = now(); return new; end;
$$;

drop trigger if exists support_updated_at_trigger on public.support_tickets;
create trigger support_updated_at_trigger before update on public.support_tickets
for each row execute procedure public.set_support_updated_at();

create or replace function public.write_support_audit()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  if old.status is distinct from new.status then
    insert into public.audit_events(actor_id, action, entity_type, entity_id, details)
    values (auth.uid(), 'support_status_changed', 'support_tickets', new.id::text,
      jsonb_build_object('ticket_code', new.ticket_code, 'from', old.status, 'to', new.status));
  end if;
  return null;
end;
$$;

drop trigger if exists support_audit_trigger on public.support_tickets;
create trigger support_audit_trigger after update on public.support_tickets
for each row execute procedure public.write_support_audit();

