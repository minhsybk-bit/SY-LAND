-- SỸ LAND: ghi nhận phiên bản điều khoản người dùng đã chấp thuận
-- Chạy sau SUPABASE_SCHEMA.sql.

alter table public.profiles
  add column if not exists terms_version text,
  add column if not exists terms_accepted_at timestamptz;

create or replace function public.handle_new_user()
returns trigger
language plpgsql
security definer
set search_path = public
as $$
begin
  insert into public.profiles (id, full_name, email, role, terms_version, terms_accepted_at)
  values (
    new.id,
    coalesce(new.raw_user_meta_data ->> 'full_name', ''),
    coalesce(new.email, ''),
    'user',
    nullif(new.raw_user_meta_data ->> 'terms_version', ''),
    case
      when coalesce(new.raw_user_meta_data ->> 'terms_accepted_at', '') <> ''
      then (new.raw_user_meta_data ->> 'terms_accepted_at')::timestamptz
      else null
    end
  )
  on conflict (id) do update set
    full_name = excluded.full_name,
    email = excluded.email,
    terms_version = coalesce(excluded.terms_version, public.profiles.terms_version),
    terms_accepted_at = coalesce(excluded.terms_accepted_at, public.profiles.terms_accepted_at),
    updated_at = now();
  return new;
end;
$$;

-- Tài khoản cũ chưa có bản ghi chấp thuận vẫn đăng nhập bình thường. Có thể yêu
-- cầu chấp thuận lại trong một bản cập nhật giao diện sau; không tự gán ngày giả.

