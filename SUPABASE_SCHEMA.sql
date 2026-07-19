-- SỸ LAND: hồ sơ tài khoản và bản quyền tập trung
-- Chạy toàn bộ tệp một lần trong Supabase Dashboard > SQL Editor.

create extension if not exists pgcrypto;

create table if not exists public.profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  full_name text not null default '',
  email text not null,
  role text not null default 'user' check (role in ('admin', 'user')),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create unique index if not exists profiles_email_unique on public.profiles (lower(email));

create table if not exists public.licenses (
  id uuid primary key default gen_random_uuid(),
  code text not null unique,
  customer text not null,
  email text not null,
  plan text not null check (plan in ('Cá nhân', 'Văn phòng', 'Đơn vị')),
  expires_at timestamptz not null,
  status text not null default 'Hoạt động' check (status in ('Hoạt động', 'Đã khóa')),
  created_by uuid references auth.users(id),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists licenses_email_index on public.licenses (lower(email));

create or replace function public.is_syland_admin()
returns boolean
language sql
stable
security definer
set search_path = public
as $$
  select exists (
    select 1 from public.profiles
    where id = auth.uid() and role = 'admin'
  );
$$;

create or replace function public.handle_new_user()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  insert into public.profiles (id, full_name, email, role)
  values (
    new.id,
    coalesce(new.raw_user_meta_data ->> 'full_name', ''),
    coalesce(new.email, ''),
    'user'
  )
  on conflict (id) do update set
    full_name = excluded.full_name,
    email = excluded.email,
    updated_at = now();
  return new;
end;
$$;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
  after insert or update of email, raw_user_meta_data on auth.users
  for each row execute procedure public.handle_new_user();

-- Bổ sung hồ sơ cho tài khoản đã đăng ký trước khi chạy tệp này.
insert into public.profiles (id, full_name, email, role)
select id, coalesce(raw_user_meta_data ->> 'full_name', ''), coalesce(email, ''), 'user'
from auth.users
on conflict (id) do nothing;

alter table public.profiles enable row level security;
alter table public.licenses enable row level security;

drop policy if exists "profile_self_read" on public.profiles;
create policy "profile_self_read" on public.profiles
for select to authenticated
using (id = auth.uid() or public.is_syland_admin());

-- Không cho client tự đổi role. Thông tin hồ sơ được trigger đồng bộ từ Auth;
-- quyền admin chỉ cấp bằng SQL Editor hoặc một dịch vụ quản trị bảo mật.
drop policy if exists "profile_self_update" on public.profiles;

drop policy if exists "license_owner_read" on public.licenses;
create policy "license_owner_read" on public.licenses
for select to authenticated
using (lower(email) = lower(coalesce(auth.jwt() ->> 'email', '')) or public.is_syland_admin());

drop policy if exists "license_admin_insert" on public.licenses;
create policy "license_admin_insert" on public.licenses
for insert to authenticated
with check (public.is_syland_admin());

drop policy if exists "license_admin_update" on public.licenses;
create policy "license_admin_update" on public.licenses
for update to authenticated
using (public.is_syland_admin())
with check (public.is_syland_admin());

drop policy if exists "license_admin_delete" on public.licenses;
create policy "license_admin_delete" on public.licenses
for delete to authenticated
using (public.is_syland_admin());

-- Sau khi Nguyễn Minh Sỹ đã đăng ký tài khoản, thay email bên dưới nếu cần rồi
-- chỉ chạy riêng câu UPDATE này để cấp quyền quản trị. Không cấp admin từ client.
update public.profiles
set role = 'admin', updated_at = now()
where lower(email) = lower('minhsybk@gmail.com');
