-- =========================================================================
-- Phase 3: Đơn hàng bán + Phiếu thu/chi + Hợp đồng
-- Chạy file này trong Supabase SQL Editor SAU khi đã chạy schema.sql (Phase 1)
-- và phase2_devices.sql (Phase 2).
-- =========================================================================

-- --------------------------------------------------------------------------
-- SALES_ORDERS: 1 đơn = 1 khách hàng + 1 máy (theo IMEI) đang "Còn hàng".
-- Nhân viên tạo đơn là hoàn tất luôn, không cần Quản lý duyệt.
-- --------------------------------------------------------------------------
create table if not exists sales_orders (
  id uuid primary key default gen_random_uuid(),
  order_code text unique,
  customer_id uuid not null references customers(id),
  device_id uuid not null references devices(id),
  sale_price numeric not null check (sale_price >= 0),
  discount numeric not null default 0 check (discount >= 0),
  total_amount numeric not null check (total_amount >= 0),
  status text not null default 'completed' check (status in ('completed', 'cancelled')),
  notes text,
  created_by uuid references employees(id),
  updated_by uuid references employees(id),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create sequence if not exists order_code_seq;
create or replace function set_order_code() returns trigger as $$
begin
  if new.order_code is null then
    new.order_code := 'DH' || lpad(nextval('order_code_seq')::text, 5, '0');
  end if;
  return new;
end;
$$ language plpgsql;
drop trigger if exists trg_order_code on sales_orders;
create trigger trg_order_code before insert on sales_orders
  for each row execute function set_order_code();

-- Dùng lại function set_updated_at() đã tạo ở phase2_devices.sql
drop trigger if exists sales_orders_set_updated_at on sales_orders;
create trigger sales_orders_set_updated_at
  before update on sales_orders
  for each row execute function set_updated_at();

create index if not exists sales_orders_customer_idx on sales_orders (customer_id);
create index if not exists sales_orders_device_idx on sales_orders (device_id);

-- --------------------------------------------------------------------------
-- ORDER_PAYMENTS: 1 đơn có thể chia nhiều phiếu thu (tiền mặt / chuyển khoản
-- / trả góp) — tổng các phiếu = total_amount của đơn (kiểm tra ở phía app).
-- --------------------------------------------------------------------------
create table if not exists order_payments (
  id uuid primary key default gen_random_uuid(),
  payment_code text unique,
  order_id uuid not null references sales_orders(id) on delete cascade,
  method text not null check (method in ('cash', 'bank_transfer', 'installment')),
  amount numeric not null check (amount > 0),
  installment_provider text,
  installment_contract_code text,
  note text,
  created_by uuid references employees(id),
  created_at timestamptz not null default now()
);

create sequence if not exists payment_code_seq;
create or replace function set_payment_code() returns trigger as $$
begin
  if new.payment_code is null then
    new.payment_code := 'PT' || lpad(nextval('payment_code_seq')::text, 5, '0');
  end if;
  return new;
end;
$$ language plpgsql;
drop trigger if exists trg_payment_code on order_payments;
create trigger trg_payment_code before insert on order_payments
  for each row execute function set_payment_code();

create index if not exists order_payments_order_idx on order_payments (order_id);

-- --------------------------------------------------------------------------
-- CONTRACTS: hợp đồng mua bán, tự sinh 1 dòng khi đơn hàng được tạo.
-- Nội dung hợp đồng render trực tiếp từ dữ liệu đơn/khách/máy phía app,
-- bảng này chỉ lưu mã hợp đồng để tra cứu/in lại.
-- --------------------------------------------------------------------------
create table if not exists contracts (
  id uuid primary key default gen_random_uuid(),
  contract_code text unique,
  order_id uuid not null references sales_orders(id) on delete cascade,
  created_by uuid references employees(id),
  created_at timestamptz not null default now()
);

create sequence if not exists contract_code_seq;
create or replace function set_contract_code() returns trigger as $$
begin
  if new.contract_code is null then
    new.contract_code := 'HD' || lpad(nextval('contract_code_seq')::text, 5, '0');
  end if;
  return new;
end;
$$ language plpgsql;
drop trigger if exists trg_contract_code on contracts;
create trigger trg_contract_code before insert on contracts
  for each row execute function set_contract_code();

create index if not exists contracts_order_idx on contracts (order_id);

-- --------------------------------------------------------------------------
-- ROW LEVEL SECURITY
-- --------------------------------------------------------------------------
alter table sales_orders enable row level security;
alter table order_payments enable row level security;
alter table contracts enable row level security;

-- SALES_ORDERS ---------------------------------------------------------
create policy "sales_orders_select_active" on sales_orders for select
  using (exists (select 1 from employees e where e.user_id = auth.uid() and e.is_active));

create policy "sales_orders_insert_staff" on sales_orders for insert
  with check (exists (select 1 from employees e where e.user_id = auth.uid() and e.is_active and e.role in ('nhan_vien','quan_ly')));

create policy "sales_orders_update_staff" on sales_orders for update
  using (exists (select 1 from employees e where e.user_id = auth.uid() and e.is_active and e.role in ('nhan_vien','quan_ly')));

create policy "sales_orders_delete_manager" on sales_orders for delete
  using (exists (select 1 from employees e where e.user_id = auth.uid() and e.is_active and e.role = 'quan_ly'));

-- ORDER_PAYMENTS ---------------------------------------------------------
create policy "order_payments_select_active" on order_payments for select
  using (exists (select 1 from employees e where e.user_id = auth.uid() and e.is_active));

create policy "order_payments_insert_staff" on order_payments for insert
  with check (exists (select 1 from employees e where e.user_id = auth.uid() and e.is_active and e.role in ('nhan_vien','quan_ly')));

create policy "order_payments_delete_manager" on order_payments for delete
  using (exists (select 1 from employees e where e.user_id = auth.uid() and e.is_active and e.role = 'quan_ly'));

-- CONTRACTS ---------------------------------------------------------------
create policy "contracts_select_active" on contracts for select
  using (exists (select 1 from employees e where e.user_id = auth.uid() and e.is_active));

create policy "contracts_insert_staff" on contracts for insert
  with check (exists (select 1 from employees e where e.user_id = auth.uid() and e.is_active and e.role in ('nhan_vien','quan_ly')));
