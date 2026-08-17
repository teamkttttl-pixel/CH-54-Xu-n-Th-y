-- ==========================================================================
-- Quản lý mua bán điện thoại / máy cũ — Phase 1 schema
-- Chạy toàn bộ file này 1 lần trong Supabase SQL Editor (project mới, trống).
-- ==========================================================================

create extension if not exists pgcrypto;

-- --------------------------------------------------------------------------
-- EMPLOYEES: tài khoản nội bộ, liên kết với Supabase Auth qua user_id.
-- Quản lý "mời" nhân viên bằng cách insert 1 dòng (user_id = null, chỉ có
-- email). Khi người được mời tự đăng ký (Tạo tài khoản lần đầu) đúng email
-- đó, app sẽ tự UPDATE user_id để liên kết (xem resolveEmployee trong App.jsx).
-- --------------------------------------------------------------------------
create table if not exists employees (
  id uuid primary key default gen_random_uuid(),
  user_id uuid unique references auth.users(id) on delete set null,
  full_name text not null,
  email text not null unique,
  role text not null check (role in ('nhan_vien','quan_ly','ke_toan')),
  is_active boolean not null default true,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

-- --------------------------------------------------------------------------
-- CUSTOMERS
-- --------------------------------------------------------------------------
create table if not exists customers (
  id uuid primary key default gen_random_uuid(),
  customer_code text unique,
  full_name text not null,
  cccd text,
  cccd_issue_date date,
  cccd_issue_place text,
  address text,
  phone text,
  email text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now(),
  created_by uuid references employees(id),
  updated_by uuid references employees(id)
);
create unique index if not exists customers_cccd_uidx on customers (cccd) where cccd is not null and cccd <> '';
create unique index if not exists customers_phone_uidx on customers (phone) where phone is not null and phone <> '';

-- Auto-generate mã khách hàng dạng KH0001, KH0002... nếu không tự nhập
create sequence if not exists customer_code_seq;
create or replace function set_customer_code() returns trigger as $$
begin
  if new.customer_code is null then
    new.customer_code := 'KH' || lpad(nextval('customer_code_seq')::text, 4, '0');
  end if;
  new.updated_at := now();
  return new;
end;
$$ language plpgsql;
drop trigger if exists trg_customer_code on customers;
create trigger trg_customer_code before insert or update on customers
  for each row execute function set_customer_code();

-- --------------------------------------------------------------------------
-- AUDIT LOG (dùng chung cho các Phase sau, tạo trước để không phải sửa RLS)
-- --------------------------------------------------------------------------
create table if not exists audit_logs (
  id uuid primary key default gen_random_uuid(),
  table_name text not null,
  record_id uuid,
  action text not null,
  old_data jsonb,
  new_data jsonb,
  performed_by uuid references employees(id),
  created_at timestamptz not null default now()
);

-- --------------------------------------------------------------------------
-- ROW LEVEL SECURITY
-- --------------------------------------------------------------------------
alter table employees enable row level security;
alter table customers enable row level security;
alter table audit_logs enable row level security;

-- EMPLOYEES -----------------------------------------------------------------
-- Ai cũng đọc được danh sách nhân viên nếu đã đăng nhập & đang active (cần
-- để hiển thị tên/role bản thân + để employees module của quản lý hoạt động).
create policy "employees_select_authenticated" on employees for select
  using (auth.role() = 'authenticated');

-- Chỉ Quản lý được thêm nhân viên mới (mời bằng email).
create policy "employees_insert_manager" on employees for insert
  with check (exists (select 1 from employees e where e.user_id = auth.uid() and e.role = 'quan_ly' and e.is_active));

-- Quản lý sửa được mọi dòng (đổi role, khóa/mở khóa...).
create policy "employees_update_manager" on employees for update
  using (exists (select 1 from employees e where e.user_id = auth.uid() and e.role = 'quan_ly' and e.is_active));

-- Nhân viên được mời tự liên kết tài khoản của mình lần đầu đăng ký
-- (chỉ khi dòng đó đang user_id = null và đúng email JWT của họ).
create policy "employees_self_link" on employees for update
  using (user_id is null and lower(email) = lower(coalesce(auth.jwt() ->> 'email','')))
  with check (user_id = auth.uid() and lower(email) = lower(coalesce(auth.jwt() ->> 'email','')));

-- CUSTOMERS -------------------------------------------------------------
create policy "customers_select_authenticated" on customers for select
  using (exists (select 1 from employees e where e.user_id = auth.uid() and e.is_active));

create policy "customers_insert_authenticated" on customers for insert
  with check (exists (select 1 from employees e where e.user_id = auth.uid() and e.is_active));

create policy "customers_update_authenticated" on customers for update
  using (exists (select 1 from employees e where e.user_id = auth.uid() and e.is_active));

create policy "customers_delete_manager" on customers for delete
  using (exists (select 1 from employees e where e.user_id = auth.uid() and e.role = 'quan_ly' and e.is_active));

-- AUDIT LOGS --------------------------------------------------------------
create policy "audit_logs_select_manager_or_ketoan" on audit_logs for select
  using (exists (select 1 from employees e where e.user_id = auth.uid() and e.role in ('quan_ly','ke_toan') and e.is_active));

create policy "audit_logs_insert_authenticated" on audit_logs for insert
  with check (exists (select 1 from employees e where e.user_id = auth.uid() and e.is_active));

-- --------------------------------------------------------------------------
-- BƯỚC CUỐI — TẠO TÀI KHOẢN QUẢN LÝ ĐẦU TIÊN (bootstrap)
-- App dùng "tên đăng nhập" (số điện thoại) thay vì email thật — phía sau
-- app tự ghép thêm "@gmail.com" cho hợp lệ với Supabase Auth, người dùng
-- không cần biết việc này. Đăng nhập lần đầu trên app chỉ cần gõ đúng phần
-- số điện thoại (0914657111), không cần gõ đuôi @gmail.com.
-- --------------------------------------------------------------------------
insert into employees (full_name, email, role, is_active)
values ('Quản lý cửa hàng', '0914657111@gmail.com', 'quan_ly', true)
on conflict (email) do nothing;
