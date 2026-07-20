-- SỸ LAND: cấp quyền quản trị cho tài khoản chủ hệ thống.
-- Đã được chủ tài khoản Nguyễn Minh Sỹ xác nhận ngày 20/07/2026.
-- Chạy một lần trong Supabase Dashboard > SQL Editor.

insert into public.profiles (id, full_name, email, role)
select id, coalesce(nullif(raw_user_meta_data ->> 'full_name', ''), 'Nguyễn Minh Sỹ'), coalesce(email, ''), 'admin'
from auth.users
where lower(email) = lower('minhsybk@gmail.com')
on conflict (id) do update set
  role = 'admin',
  full_name = coalesce(nullif(excluded.full_name, ''), public.profiles.full_name),
  email = excluded.email,
  updated_at = now();

-- Kết quả phải trả về đúng một dòng có email minhsybk@gmail.com và role admin.
select id, full_name, email, role, updated_at
from public.profiles
where lower(email) = lower('minhsybk@gmail.com');
