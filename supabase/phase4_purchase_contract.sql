-- =========================================================================
-- Phase 4 (bổ sung): Hợp đồng mua bán cho Nhập máy/Thu cũ
-- Chạy SAU khi đã có phase4_purchases.sql (cần bảng purchase_orders trước).
-- =========================================================================

-- contracts trước đây chỉ gắn với sales_orders (order_id not null). Giờ mở
-- rộng để 1 hợp đồng có thể gắn với purchase_orders (thu mua máy cũ) thay vì
-- đơn bán — mỗi hợp đồng chỉ thuộc đúng 1 trong 2 loại.
alter table contracts alter column order_id drop not null;
alter table contracts add column if not exists purchase_order_id uuid references purchase_orders(id) on delete cascade;

alter table contracts drop constraint if exists contracts_one_target;
alter table contracts add constraint contracts_one_target
  check (
    (order_id is not null and purchase_order_id is null)
    or (order_id is null and purchase_order_id is not null)
  );

create unique index if not exists contracts_purchase_order_uidx on contracts (purchase_order_id) where purchase_order_id is not null;
create index if not exists contracts_purchase_order_idx on contracts (purchase_order_id);
