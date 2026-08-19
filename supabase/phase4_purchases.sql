-- =========================================================================
-- Phase 4: Nhập máy/thu cũ (mua lại máy cũ từ khách + đổi trừ khi bán)
-- Chạy SAU khi đã có schema.sql + phase2_devices.sql + phase3_orders.sql
-- =========================================================================

-- --------------------------------------------------------------------------
-- PURCHASE_ORDERS: phiếu thu mua máy cũ — độc lập (mua hẳn, trả tiền khách)
-- hoặc gắn với 1 đơn bán (khách đổi máy cũ lấy máy mới, trừ vào tiền đơn).
-- --------------------------------------------------------------------------
create table if not exists purchase_orders (
  id uuid primary key default gen_random_uuid(),
  purchase_code text unique,
  customer_id uuid not null references customers(id),
  device_id uuid not null unique references devices(id),
  linked_sale_order_id uuid references sales_orders(id) on delete set null,
  purchase_price numeric not null check (purchase_price >= 0),
  payment_method text not null check (payment_method in ('cash', 'bank_transfer', 'trade_in')),
  notes text,
  created_by uuid references employees(id),
  created_at timestamptz not null default now()
);

create sequence if not exists purchase_code_seq;
create or replace function set_purchase_code() returns trigger as $$
begin
  if new.purchase_code is null then
    new.purchase_code := 'NM' || lpad(nextval('purchase_code_seq')::text, 5, '0');
  end if;
  return new;
end;
$$ language plpgsql;
drop trigger if exists trg_purchase_code on purchase_orders;
create trigger trg_purchase_code before insert on purchase_orders
  for each row execute function set_purchase_code();

create index if not exists purchase_orders_customer_idx on purchase_orders (customer_id);
create index if not exists purchase_orders_linked_order_idx on purchase_orders (linked_sale_order_id);

alter table purchase_orders enable row level security;

create policy "purchase_orders_select_active" on purchase_orders for select
  using (exists (select 1 from employees e where e.user_id = auth.uid() and e.is_active));

create policy "purchase_orders_insert_staff" on purchase_orders for insert
  with check (exists (select 1 from employees e where e.user_id = auth.uid() and e.is_active and e.role in ('nhan_vien','quan_ly')));

create policy "purchase_orders_delete_manager" on purchase_orders for delete
  using (exists (select 1 from employees e where e.user_id = auth.uid() and e.is_active and e.role = 'quan_ly'));

-- --------------------------------------------------------------------------
-- ORDER_PAYMENTS: thêm phương thức "trade_in" (đổi máy cũ trừ vào đơn bán)
-- và cột liên kết tới máy cũ được đổi (devices mới tạo từ lượt đổi này).
-- --------------------------------------------------------------------------
alter table order_payments drop constraint if exists order_payments_method_check;
alter table order_payments add constraint order_payments_method_check
  check (method in ('cash', 'bank_transfer', 'installment', 'trade_in'));

alter table order_payments add column if not exists trade_in_device_id uuid references devices(id);
