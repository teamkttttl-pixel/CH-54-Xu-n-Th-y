-- =========================================================================
-- Phase 5a: Đa cửa hàng (Multi-store) — nền tảng
-- Chạy SAU khi đã có đầy đủ schema.sql + phase2 + phase3 + phase4 (tất cả).
-- File này gắn "store_id" vào mọi bảng dữ liệu kinh doanh + viết lại toàn bộ
-- RLS để mỗi cửa hàng chỉ thấy đúng dữ liệu của mình.
-- =========================================================================

-- --------------------------------------------------------------------------
-- STORES
-- --------------------------------------------------------------------------
create table if not exists stores (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  address text,
  created_at timestamptz not null default now()
);

insert into stores (name, address) values
  ('CH 54 Xuân Thủy', null),
  ('CH 81 Xã Đàn', null),
  ('CH 380 Nguyễn Trãi', null)
on conflict do nothing;

-- --------------------------------------------------------------------------
-- Thêm cột store_id vào các bảng dữ liệu kinh doanh
-- --------------------------------------------------------------------------
alter table employees add column if not exists store_id uuid references stores(id);
alter table customers add column if not exists store_id uuid references stores(id);
alter table devices add column if not exists store_id uuid references stores(id);
alter table sales_orders add column if not exists store_id uuid references stores(id);
alter table order_payments add column if not exists store_id uuid references stores(id);
alter table contracts add column if not exists store_id uuid references stores(id);
alter table purchase_orders add column if not exists store_id uuid references stores(id);
alter table audit_logs add column if not exists store_id uuid references stores(id);

-- --------------------------------------------------------------------------
-- Gán toàn bộ dữ liệu hiện có (của CH 54 Xuân Thủy) về đúng cửa hàng đó
-- --------------------------------------------------------------------------
do $$
declare ch54_id uuid;
begin
  select id into ch54_id from stores where name = 'CH 54 Xuân Thủy' limit 1;

  update employees set store_id = ch54_id where store_id is null;
  update customers set store_id = ch54_id where store_id is null;
  update devices set store_id = ch54_id where store_id is null;
  update sales_orders set store_id = ch54_id where store_id is null;
  update purchase_orders set store_id = ch54_id where store_id is null;
  update contracts c set store_id = ch54_id where store_id is null;
  update order_payments set store_id = ch54_id where store_id is null;
  update audit_logs set store_id = ch54_id where store_id is null;
end $$;

alter table employees alter column store_id set not null;
alter table customers alter column store_id set not null;
alter table devices alter column store_id set not null;
alter table sales_orders alter column store_id set not null;
alter table order_payments alter column store_id set not null;
alter table contracts alter column store_id set not null;
alter table purchase_orders alter column store_id set not null;
-- audit_logs.store_id CHO PHÉP null (log hệ thống cũ có thể thiếu, không chặn ghi log mới)

create index if not exists employees_store_idx on employees (store_id);
create index if not exists customers_store_idx on customers (store_id);
create index if not exists devices_store_idx on devices (store_id);
create index if not exists sales_orders_store_idx on sales_orders (store_id);
create index if not exists purchase_orders_store_idx on purchase_orders (store_id);
create index if not exists audit_logs_store_idx on audit_logs (store_id);

-- --------------------------------------------------------------------------
-- Helper: lấy store_id của nhân viên đang đăng nhập (dùng lại trong mọi RLS)
-- --------------------------------------------------------------------------
create or replace function current_employee_store_id() returns uuid as $$
  select store_id from employees where user_id = auth.uid() and is_active limit 1;
$$ language sql stable security definer;

create or replace function current_employee_role() returns text as $$
  select role from employees where user_id = auth.uid() and is_active limit 1;
$$ language sql stable security definer;

-- --------------------------------------------------------------------------
-- STORES: ai đã đăng nhập cũng đọc được danh sách cửa hàng (để hiện tên trên
-- header/sidebar), không ai được sửa qua app (chỉ sửa tay qua SQL Editor).
-- --------------------------------------------------------------------------
alter table stores enable row level security;
drop policy if exists "stores_select_authenticated" on stores;
create policy "stores_select_authenticated" on stores for select
  using (auth.role() = 'authenticated');

-- --------------------------------------------------------------------------
-- EMPLOYEES — viết lại toàn bộ policy, giới hạn theo store_id
-- --------------------------------------------------------------------------
drop policy if exists "employees_select_authenticated" on employees;
create policy "employees_select_same_store" on employees for select
  using (store_id = current_employee_store_id() or user_id = auth.uid());

drop policy if exists "employees_insert_manager" on employees;
create policy "employees_insert_manager" on employees for insert
  with check (
    store_id = current_employee_store_id()
    and current_employee_role() = 'quan_ly'
  );

drop policy if exists "employees_update_manager" on employees;
create policy "employees_update_manager" on employees for update
  using (store_id = current_employee_store_id() and current_employee_role() = 'quan_ly');

-- self_link: giữ nguyên (không phụ thuộc store_id — người được mời tự liên kết
-- tài khoản của chính họ, store_id đã được Quản lý gán sẵn từ lúc mời)
-- (policy "employees_self_link" không đổi, vẫn còn hiệu lực từ schema.sql)

-- --------------------------------------------------------------------------
-- CUSTOMERS
-- --------------------------------------------------------------------------
drop policy if exists "customers_select_authenticated" on customers;
create policy "customers_select_same_store" on customers for select
  using (store_id = current_employee_store_id());

drop policy if exists "customers_insert_authenticated" on customers;
create policy "customers_insert_same_store" on customers for insert
  with check (store_id = current_employee_store_id());

drop policy if exists "customers_update_authenticated" on customers;
create policy "customers_update_same_store" on customers for update
  using (store_id = current_employee_store_id());

drop policy if exists "customers_delete_manager" on customers;
create policy "customers_delete_manager" on customers for delete
  using (store_id = current_employee_store_id() and current_employee_role() = 'quan_ly');

-- --------------------------------------------------------------------------
-- DEVICES
-- --------------------------------------------------------------------------
drop policy if exists "devices_select" on devices;
create policy "devices_select_same_store" on devices for select
  using (store_id = current_employee_store_id());

drop policy if exists "devices_insert" on devices;
create policy "devices_insert_same_store" on devices for insert
  with check (store_id = current_employee_store_id() and current_employee_role() in ('nhan_vien','quan_ly'));

drop policy if exists "devices_update" on devices;
create policy "devices_update_same_store" on devices for update
  using (store_id = current_employee_store_id() and current_employee_role() in ('nhan_vien','quan_ly'));

drop policy if exists "devices_delete" on devices;
create policy "devices_delete_same_store" on devices for delete
  using (store_id = current_employee_store_id() and current_employee_role() = 'quan_ly');

-- --------------------------------------------------------------------------
-- SALES_ORDERS
-- --------------------------------------------------------------------------
drop policy if exists "sales_orders_select_active" on sales_orders;
create policy "sales_orders_select_same_store" on sales_orders for select
  using (store_id = current_employee_store_id());

drop policy if exists "sales_orders_insert_staff" on sales_orders;
create policy "sales_orders_insert_same_store" on sales_orders for insert
  with check (store_id = current_employee_store_id() and current_employee_role() in ('nhan_vien','quan_ly'));

drop policy if exists "sales_orders_update_staff" on sales_orders;
create policy "sales_orders_update_same_store" on sales_orders for update
  using (store_id = current_employee_store_id() and current_employee_role() in ('nhan_vien','quan_ly'));

drop policy if exists "sales_orders_delete_manager" on sales_orders;
create policy "sales_orders_delete_same_store" on sales_orders for delete
  using (store_id = current_employee_store_id() and current_employee_role() = 'quan_ly');

-- --------------------------------------------------------------------------
-- ORDER_PAYMENTS
-- --------------------------------------------------------------------------
drop policy if exists "order_payments_select_active" on order_payments;
create policy "order_payments_select_same_store" on order_payments for select
  using (store_id = current_employee_store_id());

drop policy if exists "order_payments_insert_staff" on order_payments;
create policy "order_payments_insert_same_store" on order_payments for insert
  with check (store_id = current_employee_store_id() and current_employee_role() in ('nhan_vien','quan_ly'));

drop policy if exists "order_payments_delete_manager" on order_payments;
create policy "order_payments_delete_same_store" on order_payments for delete
  using (store_id = current_employee_store_id() and current_employee_role() = 'quan_ly');

-- --------------------------------------------------------------------------
-- CONTRACTS
-- --------------------------------------------------------------------------
drop policy if exists "contracts_select_active" on contracts;
create policy "contracts_select_same_store" on contracts for select
  using (store_id = current_employee_store_id());

drop policy if exists "contracts_insert_staff" on contracts;
create policy "contracts_insert_same_store" on contracts for insert
  with check (store_id = current_employee_store_id() and current_employee_role() in ('nhan_vien','quan_ly'));

-- --------------------------------------------------------------------------
-- PURCHASE_ORDERS
-- --------------------------------------------------------------------------
drop policy if exists "purchase_orders_select_active" on purchase_orders;
create policy "purchase_orders_select_same_store" on purchase_orders for select
  using (store_id = current_employee_store_id());

drop policy if exists "purchase_orders_insert_staff" on purchase_orders;
create policy "purchase_orders_insert_same_store" on purchase_orders for insert
  with check (store_id = current_employee_store_id() and current_employee_role() in ('nhan_vien','quan_ly'));

drop policy if exists "purchase_orders_delete_manager" on purchase_orders;
create policy "purchase_orders_delete_same_store" on purchase_orders for delete
  using (store_id = current_employee_store_id() and current_employee_role() = 'quan_ly');

-- --------------------------------------------------------------------------
-- AUDIT_LOGS
-- --------------------------------------------------------------------------
drop policy if exists "audit_logs_select_manager_or_ketoan" on audit_logs;
create policy "audit_logs_select_same_store" on audit_logs for select
  using (store_id = current_employee_store_id() and current_employee_role() in ('quan_ly','ke_toan'));

drop policy if exists "audit_logs_insert_authenticated" on audit_logs;
create policy "audit_logs_insert_same_store" on audit_logs for insert
  with check (store_id = current_employee_store_id());

-- --------------------------------------------------------------------------
-- BOOTSTRAP — tài khoản Quản lý đầu tiên cho 2 cửa hàng mới
-- --------------------------------------------------------------------------
insert into employees (full_name, email, role, is_active, store_id)
select 'Quản lý cửa hàng', '0914657111-1@gmail.com', 'quan_ly', true, id
from stores where name = 'CH 81 Xã Đàn'
on conflict (email) do nothing;

insert into employees (full_name, email, role, is_active, store_id)
select 'Quản lý cửa hàng', '0914657111-2@gmail.com', 'quan_ly', true, id
from stores where name = 'CH 380 Nguyễn Trãi'
on conflict (email) do nothing;
