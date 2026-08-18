-- =========================================================================
-- Phase 2: Kho hàng theo IMEI
-- Chạy file này trong Supabase SQL Editor SAU khi đã chạy schema.sql (Phase 1)
-- =========================================================================

create table if not exists devices (
  id uuid primary key default gen_random_uuid(),
  imei text not null unique,
  model text not null,
  storage text,
  color text,
  condition text not null default 'used' check (condition in ('new', 'used')),
  status text not null default 'in_stock' check (status in ('in_stock', 'reserved', 'sold')),
  cost_price numeric,
  sale_price numeric,
  supplier text,
  import_date date,
  notes text,
  created_by uuid references employees(id),
  updated_by uuid references employees(id),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists devices_imei_idx on devices (imei);
create index if not exists devices_status_idx on devices (status);

-- Tự động cập nhật updated_at mỗi khi sửa
create or replace function set_updated_at()
returns trigger as $$
begin
  new.updated_at = now();
  return new;
end;
$$ language plpgsql;

drop trigger if exists devices_set_updated_at on devices;
create trigger devices_set_updated_at
  before update on devices
  for each row execute function set_updated_at();

-- Row Level Security
alter table devices enable row level security;

-- Mọi nhân viên đang hoạt động (đã đăng nhập, is_active=true) đều xem được kho
drop policy if exists devices_select on devices;
create policy devices_select on devices
  for select
  using (
    exists (
      select 1 from employees e
      where e.user_id = auth.uid() and e.is_active = true
    )
  );

-- Nhân viên/Quản lý được thêm máy mới (Kế toán không tham gia nhập kho)
drop policy if exists devices_insert on devices;
create policy devices_insert on devices
  for insert
  with check (
    exists (
      select 1 from employees e
      where e.user_id = auth.uid() and e.is_active = true
        and e.role in ('nhan_vien', 'quan_ly')
    )
  );

-- Nhân viên/Quản lý được sửa thông tin & trạng thái máy
drop policy if exists devices_update on devices;
create policy devices_update on devices
  for update
  using (
    exists (
      select 1 from employees e
      where e.user_id = auth.uid() and e.is_active = true
        and e.role in ('nhan_vien', 'quan_ly')
    )
  );

-- Chỉ Quản lý cửa hàng được xóa máy khỏi kho
drop policy if exists devices_delete on devices;
create policy devices_delete on devices
  for delete
  using (
    exists (
      select 1 from employees e
      where e.user_id = auth.uid() and e.is_active = true
        and e.role = 'quan_ly'
    )
  );
