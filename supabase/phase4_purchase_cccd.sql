-- =========================================================================
-- Phase 4 (bổ sung): Bộ hồ sơ thu mua hợp lệ — đính kèm ảnh CCCD
-- Chạy SAU khi đã có phase4_purchases.sql và phase4_purchase_contract.sql.
-- =========================================================================

-- Lưu link ảnh CCCD (mặt trước/sau) gắn với từng phiếu thu mua — mỗi giao
-- dịch có bộ hồ sơ riêng, kể cả khách quen mua bán nhiều lần.
alter table purchase_orders add column if not exists cccd_front_url text;
alter table purchase_orders add column if not exists cccd_back_url text;

-- Bổ sung Ngày sinh cho khách hàng (để hợp đồng đủ thông tin định danh pháp lý)
alter table customers add column if not exists date_of_birth date;

-- --------------------------------------------------------------------------
-- Storage bucket lưu ảnh CCCD — tạo qua SQL bên dưới (chạy trong SQL Editor
-- với quyền postgres là đủ, không cần vào Dashboard tạo tay).
-- --------------------------------------------------------------------------
insert into storage.buckets (id, name, public)
values ('purchase-cccd', 'purchase-cccd', true)
on conflict (id) do nothing;

drop policy if exists "purchase_cccd_select" on storage.objects;
create policy "purchase_cccd_select" on storage.objects for select
  using (bucket_id = 'purchase-cccd');

drop policy if exists "purchase_cccd_insert" on storage.objects;
create policy "purchase_cccd_insert" on storage.objects for insert
  with check (
    bucket_id = 'purchase-cccd'
    and exists (select 1 from employees e where e.user_id = auth.uid() and e.is_active and e.role in ('nhan_vien','quan_ly'))
  );

drop policy if exists "purchase_cccd_delete" on storage.objects;
create policy "purchase_cccd_delete" on storage.objects for delete
  using (
    bucket_id = 'purchase-cccd'
    and exists (select 1 from employees e where e.user_id = auth.uid() and e.is_active and e.role = 'quan_ly')
  );
