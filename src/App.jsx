import React, { useEffect, useState, useCallback } from "react";
import { supabase } from "./supabaseClient";
import * as XLSX from "xlsx";
import {
  BarChart as RBarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip as RTooltip, ResponsiveContainer,
} from "recharts";
import {
  LayoutDashboard, Users, Smartphone, ShoppingCart, Receipt, FileText,
  BarChart3, UserCog, ScrollText, Settings, LogOut, Search, Plus, X,
  Loader2, ChevronRight, Menu, ShieldAlert, Pencil, Trash2, PackageSearch,
  History, Printer, Wallet, Landmark, CalendarClock, ChevronDown, FileSpreadsheet,
  Filter, TrendingUp, Package, Award,
} from "lucide-react";

/* ---------------------------------------------------------------------- */
/* Helpers                                                                */
/* ---------------------------------------------------------------------- */

const ROLE_LABELS = {
  nhan_vien: "Nhân viên",
  quan_ly: "Quản lý cửa hàng",
  ke_toan: "Kế toán thuế",
};

function classNames(...a) {
  return a.filter(Boolean).join(" ");
}

// Đăng nhập bằng "tên đăng nhập" (số điện thoại/username tự đặt) thay vì email
// thật — Supabase Auth vẫn cần định dạng email hợp lệ phía sau, nên ta tự
// ghép thêm đuôi cố định này. Người dùng không cần biết/gõ phần đuôi.
const LOGIN_SUFFIX = "@gmail.com";
function toLoginEmail(username) {
  const u = (username || "").trim().toLowerCase();
  if (!u) return "";
  return u.includes("@") ? u : `${u}${LOGIN_SUFFIX}`;
}
function fromLoginEmail(email) {
  return (email || "").replace(LOGIN_SUFFIX, "");
}

function fmtDate(d) {
  if (!d) return "—";
  try {
    return new Date(d).toLocaleDateString("vi-VN");
  } catch {
    return d;
  }
}

function fmtVND(n) {
  if (n === null || n === undefined || n === "") return "—";
  const num = Number(n);
  if (Number.isNaN(num)) return "—";
  return num.toLocaleString("vi-VN") + "đ";
}

const DEVICE_STATUS_LABELS = {
  in_stock: "Còn hàng",
  reserved: "Đang giữ chỗ",
  sold: "Đã bán",
};
const DEVICE_STATUS_STYLES = {
  in_stock: "bg-emerald-50 text-emerald-700",
  reserved: "bg-amber-50 text-amber-700",
  sold: "bg-slate-100 text-slate-500",
};
const DEVICE_CONDITION_LABELS = {
  new: "Máy mới",
  used: "Máy cũ",
};

const PAYMENT_METHOD_LABELS = {
  cash: "Tiền mặt",
  bank_transfer: "Chuyển khoản",
  installment: "Trả góp",
};
const PAYMENT_METHOD_ICONS = {
  cash: Wallet,
  bank_transfer: Landmark,
  installment: CalendarClock,
};

// Tạo "Mã hàng" KiotViet từ model+dung lượng+màu (gom các máy cùng loại về
// chung 1 mã sản phẩm, còn Serial/IMEI phân biệt từng máy cụ thể trong đó).
function toKiotVietProductCode(device) {
  const raw = [device.model, device.storage, device.color].filter(Boolean).join(" ");
  const noDiacritics = raw
    .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/gi, "d");
  return noDiacritics.toUpperCase().replace(/[^A-Z0-9]/g, "").slice(0, 30) || "SP";
}

/* ---------------------------------------------------------------------- */
/* Shared UI bits                                                        */
/* ---------------------------------------------------------------------- */

function Card({ className = "", children }) {
  return (
    <div className={classNames("bg-white rounded-3xl shadow-sm border border-slate-100", className)}>
      {children}
    </div>
  );
}

function TextField({ label, className = "", ...props }) {
  return (
    <label className={classNames("block", className)}>
      {label && <span className="text-xs font-medium text-slate-500 mb-1 block">{label}</span>}
      <input
        {...props}
        className="w-full rounded-xl border border-slate-200 px-3 py-2 text-sm outline-none focus:ring-2 focus:ring-brand-400 focus:border-brand-400 transition"
      />
    </label>
  );
}

function EmptyState({ icon: Icon, text }) {
  return (
    <div className="flex flex-col items-center justify-center py-16 text-slate-400">
      {Icon && <Icon size={36} className="mb-3 opacity-50" />}
      <p className="text-sm">{text}</p>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Auth / Login                                                          */
/* ---------------------------------------------------------------------- */

function LoginPage({ onLoggedIn }) {
  const [mode, setMode] = useState("login"); // "login" | "signup"
  const [email, setEmail] = useState("");
  const [password, setPassword] = useState("");
  const [confirm, setConfirm] = useState("");
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");
  const [info, setInfo] = useState("");

  const submit = async (e) => {
    e.preventDefault();
    setError(""); setInfo("");
    if (!email.trim() || !password) { setError("Vui lòng nhập đủ tên đăng nhập và mật khẩu."); return; }
    if (mode === "signup" && password !== confirm) { setError("Mật khẩu nhập lại không khớp."); return; }
    setLoading(true);
    try {
      const loginEmail = toLoginEmail(email);
      if (mode === "login") {
        const { error: err } = await supabase.auth.signInWithPassword({ email: loginEmail, password });
        if (err) throw err;
        onLoggedIn();
      } else {
        const { data, error: err } = await supabase.auth.signUp({ email: loginEmail, password });
        if (err) throw err;
        if (data.session) {
          onLoggedIn();
        } else {
          setInfo("Đã tạo tài khoản. Nếu hệ thống yêu cầu xác nhận email, hãy kiểm tra hộp thư rồi quay lại đăng nhập.");
          setMode("login");
        }
      }
    } catch (err) {
      setError(err.message || "Có lỗi xảy ra, vui lòng thử lại.");
    } finally {
      setLoading(false);
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-b from-brand-50 to-white flex items-center justify-center p-4">
      <Card className="w-full max-w-sm p-6 sm:p-8">
        <div className="flex flex-col items-center mb-6">
          <div className="w-14 h-14 rounded-2xl bg-brand-600 flex items-center justify-center mb-3">
            <Smartphone className="text-white" size={26} />
          </div>
          <h1 className="text-lg font-semibold text-slate-800">Quản lý mua bán điện thoại</h1>
          <p className="text-xs text-slate-400 mt-1">CH 54 Xuân Thủy</p>
        </div>

        <form onSubmit={submit} className="space-y-3">
          <TextField label="Tên đăng nhập (số điện thoại)" type="text" value={email} onChange={(e) => setEmail(e.target.value)} placeholder="0914657111" autoComplete="username" />
          <TextField label="Mật khẩu" type="password" value={password} onChange={(e) => setPassword(e.target.value)} placeholder="••••••••" autoComplete={mode === "login" ? "current-password" : "new-password"} />
          {mode === "signup" && (
            <TextField label="Nhập lại mật khẩu" type="password" value={confirm} onChange={(e) => setConfirm(e.target.value)} placeholder="••••••••" />
          )}
          {error && <p className="text-xs text-rose-600 bg-rose-50 rounded-lg px-3 py-2">{error}</p>}
          {info && <p className="text-xs text-emerald-700 bg-emerald-50 rounded-lg px-3 py-2">{info}</p>}
          <button
            type="submit"
            disabled={loading}
            className="w-full bg-brand-600 hover:bg-brand-700 text-white rounded-xl py-2.5 text-sm font-medium flex items-center justify-center gap-2 transition disabled:opacity-60"
          >
            {loading && <Loader2 size={15} className="animate-spin" />}
            {mode === "login" ? "Đăng nhập" : "Tạo tài khoản"}
          </button>
        </form>

        <button
          type="button"
          onClick={() => { setMode(mode === "login" ? "signup" : "login"); setError(""); setInfo(""); }}
          className="w-full text-center text-xs text-slate-400 hover:text-brand-600 mt-4"
        >
          {mode === "login"
            ? "Được cấp tên đăng nhập nhưng chưa có mật khẩu? Tạo tài khoản lần đầu"
            : "Đã có tài khoản? Quay lại đăng nhập"}
        </button>
      </Card>
    </div>
  );
}

function NotProvisioned({ email, onSignOut }) {
  return (
    <div className="min-h-screen bg-gradient-to-b from-brand-50 to-white flex items-center justify-center p-4">
      <Card className="w-full max-w-sm p-8 text-center">
        <ShieldAlert size={36} className="mx-auto text-amber-500 mb-3" />
        <h2 className="font-semibold text-slate-800 mb-1">Tài khoản chưa được cấp quyền</h2>
        <p className="text-sm text-slate-500 mb-4">
          Tên đăng nhập <span className="font-medium text-slate-700">{fromLoginEmail(email)}</span> chưa được Quản lý thêm vào hệ thống,
          hoặc không khớp với tên đăng nhập được cấp. Vui lòng liên hệ Quản lý cửa hàng để được cấp quyền.
        </p>
        <button onClick={onSignOut} className="text-sm text-brand-600 hover:underline">Đăng xuất</button>
      </Card>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Dashboard                                                             */
/* ---------------------------------------------------------------------- */

function StatCard({ label, value, icon: Icon, comingSoon }) {
  return (
    <Card className="p-4 sm:p-5">
      <div className="flex items-center justify-between mb-3">
        <span className="text-xs font-medium text-slate-400">{label}</span>
        <Icon size={16} className="text-slate-300" />
      </div>
      {comingSoon ? (
        <p className="text-xs text-slate-400">Sẽ có ở Phase sau</p>
      ) : (
        <p className="text-2xl font-semibold text-slate-800">{value}</p>
      )}
    </Card>
  );
}

function DashboardModule({ employee, customerCount, inStockCount }) {
  const [todayStats, setTodayStats] = useState({ orders: 0, revenue: 0 });
  const [loadingStats, setLoadingStats] = useState(true);

  useEffect(() => {
    let active = true;
    const startOfDay = new Date(); startOfDay.setHours(0, 0, 0, 0);
    supabase
      .from("sales_orders")
      .select("total_amount")
      .eq("status", "completed")
      .gte("created_at", startOfDay.toISOString())
      .then(({ data }) => {
        if (!active) return;
        const revenue = (data || []).reduce((s, o) => s + Number(o.total_amount || 0), 0);
        setTodayStats({ orders: (data || []).length, revenue });
        setLoadingStats(false);
      });
    return () => { active = false; };
  }, []);

  return (
    <div>
      <div className="mb-5">
        <h2 className="text-lg font-semibold text-slate-800">Chào {employee.full_name} 👋</h2>
        <p className="text-sm text-slate-400">Vai trò: {ROLE_LABELS[employee.role] || employee.role}</p>
      </div>
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 sm:gap-4">
        <StatCard label="Tổng khách hàng" value={customerCount} icon={Users} />
        <StatCard label="Tồn kho (IMEI)" value={inStockCount} icon={Smartphone} />
        <StatCard label="Đơn hàng hôm nay" value={loadingStats ? "…" : todayStats.orders} icon={ShoppingCart} />
        <StatCard label="Doanh thu hôm nay" value={loadingStats ? "…" : fmtVND(todayStats.revenue)} icon={BarChart3} />
      </div>
      <Card className="p-5 mt-4">
        <p className="text-sm font-medium text-slate-700 mb-1">Tiến độ triển khai</p>
        <p className="text-xs text-slate-400 leading-relaxed">
          Phase 1 (Đăng nhập, phân quyền, Khách hàng), Phase 2 (Kho hàng theo IMEI) và Phase 3
          (Đơn hàng bán, Phiếu thu/chi, Hợp đồng) đang hoạt động. Hóa đơn thuế + Báo cáo + Audit Log
          sẽ được bổ sung ở các lần cập nhật tiếp theo.
        </p>
      </Card>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Customers module                                                      */
/* ---------------------------------------------------------------------- */

function CustomerForm({ initial, onCancel, onSaved, employee, existingMatch, onUseExisting }) {
  const [form, setForm] = useState(initial || {
    full_name: "", cccd: "", cccd_issue_date: "", cccd_issue_place: "", address: "", phone: "", email: "",
  });
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));

  const submit = async (e) => {
    e.preventDefault();
    if (!form.full_name.trim()) { setError("Vui lòng nhập họ tên khách hàng."); return; }
    setError(""); setSaving(true);
    try {
      const payload = {
        full_name: form.full_name.trim(),
        cccd: form.cccd.trim() || null,
        cccd_issue_date: form.cccd_issue_date || null,
        cccd_issue_place: form.cccd_issue_place.trim() || null,
        address: form.address.trim() || null,
        phone: form.phone.trim() || null,
        email: form.email.trim() || null,
      };
      if (initial?.id) {
        payload.updated_by = employee.id;
        const { error: err } = await supabase.from("customers").update(payload).eq("id", initial.id);
        if (err) throw err;
      } else {
        payload.created_by = employee.id;
        payload.updated_by = employee.id;
        const { error: err } = await supabase.from("customers").insert(payload);
        if (err) throw err;
      }
      onSaved();
    } catch (err) {
      setError(err.message || "Không lưu được, vui lòng thử lại.");
    } finally {
      setSaving(false);
    }
  };

  return (
    <Card className="p-4 sm:p-5 mb-5">
      <div className="flex items-center justify-between mb-3">
        <p className="font-semibold text-slate-800 text-sm">{initial?.id ? "Sửa khách hàng" : "Thêm khách hàng"}</p>
        <button onClick={onCancel} className="text-slate-400 hover:text-rose-600"><X size={16} /></button>
      </div>

      {existingMatch && (
        <div className="bg-amber-50 border border-amber-200 rounded-xl px-3 py-2.5 mb-3 text-xs text-amber-800 flex items-center justify-between gap-2">
          <span>Đã có khách hàng trùng CCCD/SĐT: <span className="font-medium">{existingMatch.full_name}</span> ({existingMatch.phone || existingMatch.cccd})</span>
          <button type="button" onClick={() => onUseExisting(existingMatch)} className="shrink-0 text-brand-700 font-medium hover:underline">Dùng khách này</button>
        </div>
      )}

      <form onSubmit={submit} className="grid sm:grid-cols-2 gap-3">
        <TextField label="Họ và tên *" value={form.full_name} onChange={set("full_name")} />
        <TextField label="Số điện thoại" value={form.phone} onChange={set("phone")} />
        <TextField label="Số CCCD" value={form.cccd} onChange={set("cccd")} />
        <TextField label="Ngày cấp CCCD" type="date" value={form.cccd_issue_date || ""} onChange={set("cccd_issue_date")} />
        <TextField label="Nơi cấp" value={form.cccd_issue_place} onChange={set("cccd_issue_place")} />
        <TextField label="Email" value={form.email} onChange={set("email")} />
        <TextField label="Địa chỉ" value={form.address} onChange={set("address")} className="sm:col-span-2" />
        {error && <p className="text-xs text-rose-600 sm:col-span-2">{error}</p>}
        <div className="sm:col-span-2 flex gap-2 mt-1">
          <button type="submit" disabled={saving} className="bg-brand-600 hover:bg-brand-700 text-white rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-2 disabled:opacity-60">
            {saving && <Loader2 size={14} className="animate-spin" />} Lưu
          </button>
          <button type="button" onClick={onCancel} className="text-slate-500 text-sm px-4 py-2">Hủy</button>
        </div>
      </form>
    </Card>
  );
}

function CustomersModule({ employee, onCountChange }) {
  const [customers, setCustomers] = useState([]);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState("");
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [existingMatch, setExistingMatch] = useState(null);

  const canDelete = employee.role === "quan_ly";

  const load = useCallback(async () => {
    setLoading(true);
    const { data, error } = await supabase.from("customers").select("*").order("created_at", { ascending: false }).limit(1000);
    if (!error) {
      setCustomers(data || []);
      onCountChange?.((data || []).length);
    }
    setLoading(false);
  }, [onCountChange]);

  useEffect(() => { load(); }, [load]);

  const filtered = customers.filter((c) => {
    if (!search.trim()) return true;
    const q = search.trim().toLowerCase();
    return (
      c.full_name?.toLowerCase().includes(q) ||
      c.cccd?.toLowerCase().includes(q) ||
      c.phone?.toLowerCase().includes(q)
    );
  });

  const checkDuplicate = async (form) => {
    if (!form.cccd && !form.phone) { setExistingMatch(null); return; }
    const orParts = [];
    if (form.cccd) orParts.push(`cccd.eq.${form.cccd}`);
    if (form.phone) orParts.push(`phone.eq.${form.phone}`);
    const { data } = await supabase.from("customers").select("*").or(orParts.join(",")).limit(1);
    setExistingMatch(data && data.length ? data[0] : null);
  };

  const openNew = () => { setEditing(null); setExistingMatch(null); setShowForm(true); };
  const openEdit = (c) => { setEditing(c); setExistingMatch(null); setShowForm(true); };

  const remove = async (c) => {
    if (!confirm(`Xóa khách hàng "${c.full_name}"?`)) return;
    const { error } = await supabase.from("customers").delete().eq("id", c.id);
    if (error) { alert(error.message); return; }
    load();
  };

  return (
    <div>
      <div className="flex items-center justify-between mb-4 gap-2 flex-wrap">
        <div>
          <h2 className="text-lg font-semibold text-slate-800">Khách hàng</h2>
          <p className="text-xs text-slate-400">{customers.length} khách hàng</p>
        </div>
        <button onClick={openNew} className="bg-brand-600 hover:bg-brand-700 text-white rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-1.5">
          <Plus size={15} /> Thêm khách hàng
        </button>
      </div>

      {showForm && (
        <CustomerForm
          initial={editing}
          employee={employee}
          existingMatch={!editing ? existingMatch : null}
          onCancel={() => setShowForm(false)}
          onUseExisting={(c) => { openEdit(c); }}
          onSaved={() => { setShowForm(false); load(); }}
        />
      )}

      <Card className="p-0 overflow-hidden">
        <div className="p-3 border-b border-slate-100">
          <div className="relative">
            <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-300" />
            <input
              value={search}
              onChange={(e) => {
                setSearch(e.target.value);
                if (!editing) checkDuplicate({ cccd: e.target.value.trim(), phone: e.target.value.trim() });
              }}
              placeholder="Tìm theo tên, CCCD, SĐT..."
              className="w-full pl-9 pr-3 py-2 rounded-xl border border-slate-200 text-sm outline-none focus:ring-2 focus:ring-brand-400"
            />
          </div>
        </div>
        {loading ? (
          <div className="flex justify-center py-10"><Loader2 className="animate-spin text-slate-300" /></div>
        ) : filtered.length === 0 ? (
          <EmptyState icon={Users} text="Chưa có khách hàng nào." />
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead><tr className="text-left text-xs text-slate-400 border-b border-slate-100">
                <th className="px-3 py-2">Họ tên</th>
                <th className="px-3 py-2">SĐT</th>
                <th className="px-3 py-2">CCCD</th>
                <th className="px-3 py-2">Địa chỉ</th>
                <th className="px-3 py-2"></th>
              </tr></thead>
              <tbody>
                {filtered.map((c) => (
                  <tr key={c.id} className="border-b border-slate-50 last:border-0 hover:bg-slate-50/60">
                    <td className="px-3 py-2.5 font-medium text-slate-700">{c.full_name}</td>
                    <td className="px-3 py-2.5 text-slate-500">{c.phone || "—"}</td>
                    <td className="px-3 py-2.5 text-slate-500">{c.cccd || "—"}</td>
                    <td className="px-3 py-2.5 text-slate-500 max-w-[220px] truncate">{c.address || "—"}</td>
                    <td className="px-3 py-2.5 text-right whitespace-nowrap">
                      <button onClick={() => openEdit(c)} className="text-brand-600 hover:underline text-xs mr-3">Sửa</button>
                      {canDelete && <button onClick={() => remove(c)} className="text-rose-500 hover:underline text-xs">Xóa</button>}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Card>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Inventory module — devices tracked by IMEI                            */
/* ---------------------------------------------------------------------- */

function DeviceForm({ initial, onCancel, onSaved, employee, duplicateImei }) {
  const [form, setForm] = useState(initial || {
    imei: "", model: "", storage: "", color: "", condition: "used",
    cost_price: "", sale_price: "", supplier: "", import_date: "", notes: "",
  });
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));

  const submit = async (e) => {
    e.preventDefault();
    if (!form.imei.trim()) { setError("Vui lòng nhập IMEI."); return; }
    if (!form.model.trim()) { setError("Vui lòng nhập tên model máy."); return; }
    setError(""); setSaving(true);
    try {
      const payload = {
        imei: form.imei.trim(),
        model: form.model.trim(),
        storage: form.storage.trim() || null,
        color: form.color.trim() || null,
        condition: form.condition,
        cost_price: form.cost_price === "" ? null : Number(form.cost_price),
        sale_price: form.sale_price === "" ? null : Number(form.sale_price),
        supplier: form.supplier.trim() || null,
        import_date: form.import_date || null,
        notes: form.notes.trim() || null,
        updated_by: employee.id,
      };
      if (initial?.id) {
        const { data: updated, error: err } = await supabase.from("devices").update(payload).eq("id", initial.id).select().maybeSingle();
        if (err) throw err;
        await supabase.from("audit_logs").insert({
          table_name: "devices", record_id: initial.id, action: "update",
          old_data: initial, new_data: updated, performed_by: employee.id,
        });
      } else {
        payload.status = "in_stock";
        payload.created_by = employee.id;
        const { data: created, error: err } = await supabase.from("devices").insert(payload).select().maybeSingle();
        if (err) throw err;
        await supabase.from("audit_logs").insert({
          table_name: "devices", record_id: created?.id, action: "create",
          new_data: created, performed_by: employee.id,
        });
      }
      onSaved();
    } catch (err) {
      if (err.code === "23505" || /duplicate/i.test(err.message || "")) {
        setError("IMEI này đã tồn tại trong kho — không được nhập trùng.");
      } else {
        setError(err.message || "Không lưu được, vui lòng thử lại.");
      }
    } finally {
      setSaving(false);
    }
  };

  return (
    <Card className="p-4 sm:p-5 mb-5">
      <div className="flex items-center justify-between mb-3">
        <p className="font-semibold text-slate-800 text-sm">{initial?.id ? "Sửa thông tin máy" : "Nhập máy mới"}</p>
        <button onClick={onCancel} className="text-slate-400 hover:text-rose-600"><X size={16} /></button>
      </div>

      {duplicateImei && (
        <div className="bg-rose-50 border border-rose-200 rounded-xl px-3 py-2.5 mb-3 text-xs text-rose-700">
          IMEI <span className="font-medium">{duplicateImei.imei}</span> đã có trong kho ({DEVICE_STATUS_LABELS[duplicateImei.status]}) — kiểm tra lại trước khi lưu.
        </div>
      )}

      <form onSubmit={submit} className="grid sm:grid-cols-2 gap-3">
        <TextField label="Số IMEI *" value={form.imei} onChange={set("imei")} disabled={!!initial?.id} />
        <TextField label="Model máy *" value={form.model} onChange={set("model")} placeholder="iPhone 14 Pro Max" />
        <TextField label="Dung lượng" value={form.storage} onChange={set("storage")} placeholder="256GB" />
        <TextField label="Màu sắc" value={form.color} onChange={set("color")} placeholder="Tím" />
        <label className="block">
          <span className="text-xs font-medium text-slate-500 mb-1 block">Tình trạng</span>
          <select value={form.condition} onChange={set("condition")} className="w-full rounded-xl border border-slate-200 px-3 py-2 text-sm">
            <option value="new">Máy mới</option>
            <option value="used">Máy cũ</option>
          </select>
        </label>
        <TextField label="Nhà cung cấp / nguồn nhập" value={form.supplier} onChange={set("supplier")} />
        <TextField label="Giá vốn (đ)" type="number" value={form.cost_price} onChange={set("cost_price")} />
        <TextField label="Giá bán đề xuất (đ)" type="number" value={form.sale_price} onChange={set("sale_price")} />
        <TextField label="Ngày nhập" type="date" value={form.import_date || ""} onChange={set("import_date")} />
        <TextField label="Ghi chú" value={form.notes} onChange={set("notes")} className="sm:col-span-2" />
        {error && <p className="text-xs text-rose-600 sm:col-span-2">{error}</p>}
        <div className="sm:col-span-2 flex gap-2 mt-1">
          <button type="submit" disabled={saving} className="bg-brand-600 hover:bg-brand-700 text-white rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-2 disabled:opacity-60">
            {saving && <Loader2 size={14} className="animate-spin" />} Lưu
          </button>
          <button type="button" onClick={onCancel} className="text-slate-500 text-sm px-4 py-2">Hủy</button>
        </div>
      </form>
    </Card>
  );
}

function ImeiHistoryPanel({ imei, onClose }) {
  const [logs, setLogs] = useState([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    let active = true;
    setLoading(true);
    supabase
      .from("audit_logs")
      .select("*")
      .eq("table_name", "devices")
      .or(`new_data->>imei.eq.${imei},old_data->>imei.eq.${imei}`)
      .order("created_at", { ascending: false })
      .then(({ data }) => { if (active) { setLogs(data || []); setLoading(false); } });
    return () => { active = false; };
  }, [imei]);

  const ACTION_LABELS = { create: "Nhập kho", update: "Cập nhật", delete: "Xóa khỏi kho" };

  return (
    <Card className="p-4 sm:p-5 mb-5">
      <div className="flex items-center justify-between mb-3">
        <p className="font-semibold text-slate-800 text-sm">Lịch sử tra cứu IMEI {imei}</p>
        <button onClick={onClose} className="text-slate-400 hover:text-rose-600"><X size={16} /></button>
      </div>
      {loading ? (
        <div className="flex justify-center py-6"><Loader2 className="animate-spin text-slate-300" /></div>
      ) : logs.length === 0 ? (
        <p className="text-xs text-slate-400">Chưa có lịch sử ghi nhận cho máy này.</p>
      ) : (
        <ul className="space-y-2">
          {logs.map((l) => (
            <li key={l.id} className="text-xs text-slate-500 border-b border-slate-50 pb-2 last:border-0">
              <span className="font-medium text-slate-700">{ACTION_LABELS[l.action] || l.action}</span> — {fmtDate(l.created_at)}
              {l.action === "update" && l.old_data?.status !== l.new_data?.status && (
                <span> · {DEVICE_STATUS_LABELS[l.old_data?.status]} → {DEVICE_STATUS_LABELS[l.new_data?.status]}</span>
              )}
            </li>
          ))}
        </ul>
      )}
    </Card>
  );
}

function InventoryModule({ employee, onCountChange }) {
  const [devices, setDevices] = useState([]);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState("all");
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [duplicateImei, setDuplicateImei] = useState(null);
  const [historyImei, setHistoryImei] = useState(null);

  const canDelete = employee.role === "quan_ly";
  const canSeeCost = employee.role !== "nhan_vien";

  const load = useCallback(async () => {
    setLoading(true);
    const { data, error } = await supabase.from("devices").select("*").order("created_at", { ascending: false }).limit(2000);
    if (!error) {
      setDevices(data || []);
      onCountChange?.((data || []).filter((d) => d.status === "in_stock").length);
    }
    setLoading(false);
  }, [onCountChange]);

  useEffect(() => { load(); }, [load]);

  const filtered = devices.filter((d) => {
    if (statusFilter !== "all" && d.status !== statusFilter) return false;
    if (!search.trim()) return true;
    const q = search.trim().toLowerCase();
    return (
      d.imei?.toLowerCase().includes(q) ||
      d.model?.toLowerCase().includes(q) ||
      d.color?.toLowerCase().includes(q)
    );
  });

  const checkImei = async (imei) => {
    if (!imei.trim()) { setDuplicateImei(null); return; }
    const { data } = await supabase.from("devices").select("*").eq("imei", imei.trim()).maybeSingle();
    setDuplicateImei(data || null);
  };

  const openNew = () => { setEditing(null); setDuplicateImei(null); setShowForm(true); };
  const openEdit = (d) => { setEditing(d); setDuplicateImei(null); setShowForm(true); };

  const remove = async (d) => {
    if (!confirm(`Xóa máy IMEI "${d.imei}" khỏi kho?`)) return;
    const { error } = await supabase.from("devices").delete().eq("id", d.id);
    if (error) { alert(error.message); return; }
    await supabase.from("audit_logs").insert({
      table_name: "devices", record_id: d.id, action: "delete",
      old_data: d, performed_by: employee.id,
    });
    load();
  };

  const cycleStatus = async (d) => {
    const order = ["in_stock", "reserved", "sold"];
    const next = order[(order.indexOf(d.status) + 1) % order.length];
    const { data: updated, error } = await supabase.from("devices").update({ status: next, updated_by: employee.id }).eq("id", d.id).select().maybeSingle();
    if (error) { alert(error.message); return; }
    await supabase.from("audit_logs").insert({
      table_name: "devices", record_id: d.id, action: "update",
      old_data: d, new_data: updated, performed_by: employee.id,
    });
    load();
  };

  return (
    <div>
      <div className="flex items-center justify-between mb-4 gap-2 flex-wrap">
        <div>
          <h2 className="text-lg font-semibold text-slate-800">Kho hàng (IMEI)</h2>
          <p className="text-xs text-slate-400">
            {devices.length} máy · {devices.filter((d) => d.status === "in_stock").length} còn hàng
          </p>
        </div>
        <button onClick={openNew} className="bg-brand-600 hover:bg-brand-700 text-white rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-1.5">
          <Plus size={15} /> Nhập máy
        </button>
      </div>

      {showForm && (
        <DeviceForm
          initial={editing}
          employee={employee}
          duplicateImei={!editing ? duplicateImei : null}
          onCancel={() => setShowForm(false)}
          onSaved={() => { setShowForm(false); load(); }}
        />
      )}

      {historyImei && <ImeiHistoryPanel imei={historyImei} onClose={() => setHistoryImei(null)} />}

      <Card className="p-0 overflow-hidden">
        <div className="p-3 border-b border-slate-100 flex flex-col sm:flex-row gap-2">
          <div className="relative flex-1">
            <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-300" />
            <input
              value={search}
              onChange={(e) => { setSearch(e.target.value); if (!editing) checkImei(e.target.value); }}
              placeholder="Tìm theo IMEI, model, màu..."
              className="w-full pl-9 pr-3 py-2 rounded-xl border border-slate-200 text-sm outline-none focus:ring-2 focus:ring-brand-400"
            />
          </div>
          <select
            value={statusFilter}
            onChange={(e) => setStatusFilter(e.target.value)}
            className="rounded-xl border border-slate-200 px-3 py-2 text-sm"
          >
            <option value="all">Tất cả trạng thái</option>
            <option value="in_stock">Còn hàng</option>
            <option value="reserved">Đang giữ chỗ</option>
            <option value="sold">Đã bán</option>
          </select>
        </div>
        {loading ? (
          <div className="flex justify-center py-10"><Loader2 className="animate-spin text-slate-300" /></div>
        ) : filtered.length === 0 ? (
          <EmptyState icon={PackageSearch} text="Chưa có máy nào trong kho." />
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead><tr className="text-left text-xs text-slate-400 border-b border-slate-100">
                <th className="px-3 py-2">IMEI</th>
                <th className="px-3 py-2">Model</th>
                <th className="px-3 py-2">Tình trạng</th>
                {canSeeCost && <th className="px-3 py-2">Giá vốn</th>}
                <th className="px-3 py-2">Giá bán</th>
                <th className="px-3 py-2">Trạng thái</th>
                <th className="px-3 py-2"></th>
              </tr></thead>
              <tbody>
                {filtered.map((d) => (
                  <tr key={d.id} className="border-b border-slate-50 last:border-0 hover:bg-slate-50/60">
                    <td className="px-3 py-2.5 font-medium text-slate-700 whitespace-nowrap">{d.imei}</td>
                    <td className="px-3 py-2.5 text-slate-600">
                      {d.model}
                      <span className="text-slate-400"> {[d.storage, d.color].filter(Boolean).join(" · ")}</span>
                    </td>
                    <td className="px-3 py-2.5 text-slate-500">{DEVICE_CONDITION_LABELS[d.condition] || "—"}</td>
                    {canSeeCost && <td className="px-3 py-2.5 text-slate-500">{fmtVND(d.cost_price)}</td>}
                    <td className="px-3 py-2.5 text-slate-500">{fmtVND(d.sale_price)}</td>
                    <td className="px-3 py-2.5">
                      <button
                        onClick={() => cycleStatus(d)}
                        className={classNames("text-xs px-2 py-0.5 rounded-full", DEVICE_STATUS_STYLES[d.status])}
                        title="Bấm để chuyển trạng thái"
                      >
                        {DEVICE_STATUS_LABELS[d.status] || d.status}
                      </button>
                    </td>
                    <td className="px-3 py-2.5 text-right whitespace-nowrap">
                      <button onClick={() => setHistoryImei(d.imei)} className="text-slate-400 hover:text-brand-600 mr-3" title="Lịch sử IMEI">
                        <History size={14} className="inline" />
                      </button>
                      <button onClick={() => openEdit(d)} className="text-brand-600 hover:underline text-xs mr-3">
                        <Pencil size={12} className="inline mr-0.5" />Sửa
                      </button>
                      {canDelete && (
                        <button onClick={() => remove(d)} className="text-rose-500 hover:underline text-xs">
                          <Trash2 size={12} className="inline mr-0.5" />Xóa
                        </button>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Card>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Print-friendly document modal — Phiếu thu & Hợp đồng                  */
/* ---------------------------------------------------------------------- */

function PrintDocModal({ type, order, customer, device, payments, contract, onClose }) {
  const today = fmtDate(new Date());

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/40 flex items-start sm:items-center justify-center p-4 overflow-y-auto print:bg-white print:p-0">
      <div className="bg-white rounded-2xl print:rounded-none shadow-xl w-full max-w-2xl p-6 sm:p-10 my-6 print:my-0 print:shadow-none print:max-w-none">
        <div className="flex items-center justify-end gap-2 mb-4 print:hidden">
          <button onClick={() => window.print()} className="bg-brand-600 hover:bg-brand-700 text-white rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-1.5">
            <Printer size={15} /> In
          </button>
          <button onClick={onClose} className="text-slate-400 hover:text-rose-600 rounded-xl px-3 py-2 text-sm">Đóng</button>
        </div>

        <div className="text-center mb-6">
          <p className="text-xs text-slate-400">CH 54 Xuân Thủy — Quản lý mua bán điện thoại</p>
          <h2 className="text-lg font-bold text-slate-800 mt-1">
            {type === "contract" ? "HỢP ĐỒNG MUA BÁN ĐIỆN THOẠI" : "PHIẾU THU TIỀN"}
          </h2>
          <p className="text-xs text-slate-400 mt-1">
            Số: {type === "contract" ? contract?.contract_code : (payments || []).map((p) => p.payment_code).join(", ")} — Ngày {today}
          </p>
        </div>

        <div className="text-sm text-slate-700 space-y-1 mb-4">
          <p><span className="text-slate-400">Mã đơn hàng:</span> {order?.order_code}</p>
          <p><span className="text-slate-400">Khách hàng:</span> {customer?.full_name}</p>
          {customer?.phone && <p><span className="text-slate-400">SĐT:</span> {customer.phone}</p>}
          {customer?.cccd && <p><span className="text-slate-400">CCCD:</span> {customer.cccd} {customer.cccd_issue_date && `— cấp ${fmtDate(customer.cccd_issue_date)}`} {customer.cccd_issue_place}</p>}
          {customer?.address && <p><span className="text-slate-400">Địa chỉ:</span> {customer.address}</p>}
        </div>

        <div className="text-sm text-slate-700 space-y-1 mb-4 border-t border-dashed border-slate-200 pt-4">
          <p className="font-medium text-slate-800">Thông tin máy</p>
          <p>{device?.model} {[device?.storage, device?.color].filter(Boolean).join(" · ")} — {DEVICE_CONDITION_LABELS[device?.condition]}</p>
          <p><span className="text-slate-400">IMEI:</span> {device?.imei}</p>
        </div>

        <div className="text-sm text-slate-700 space-y-1 mb-4 border-t border-dashed border-slate-200 pt-4">
          <div className="flex justify-between"><span className="text-slate-400">Giá bán</span><span>{fmtVND(order?.sale_price)}</span></div>
          {order?.discount > 0 && <div className="flex justify-between"><span className="text-slate-400">Giảm giá</span><span>-{fmtVND(order.discount)}</span></div>}
          <div className="flex justify-between font-semibold text-slate-800"><span>Tổng tiền</span><span>{fmtVND(order?.total_amount)}</span></div>
        </div>

        <div className="text-sm text-slate-700 space-y-1.5 mb-6 border-t border-dashed border-slate-200 pt-4">
          <p className="font-medium text-slate-800 mb-1">
            {type === "contract" ? "Hình thức thanh toán" : "Chi tiết thu tiền"}
          </p>
          {payments?.map((p) => (
            <div key={p.id} className="flex justify-between">
              <span className="text-slate-500">
                {PAYMENT_METHOD_LABELS[p.method]}
                {p.method === "installment" && p.installment_provider ? ` (${p.installment_provider}${p.installment_contract_code ? " · " + p.installment_contract_code : ""})` : ""}
              </span>
              <span>{fmtVND(p.amount)}</span>
            </div>
          ))}
        </div>

        <div className="grid grid-cols-2 gap-4 text-center text-sm text-slate-600 mt-10">
          <div>
            <p className="font-medium text-slate-700 mb-12">Khách hàng</p>
            <p className="text-xs text-slate-400">(Ký, ghi rõ họ tên)</p>
          </div>
          <div>
            <p className="font-medium text-slate-700 mb-12">Đại diện cửa hàng</p>
            <p className="text-xs text-slate-400">(Ký, ghi rõ họ tên)</p>
          </div>
        </div>
      </div>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Orders module — sales order + multi-method payments + auto contract   */
/* ---------------------------------------------------------------------- */

function CustomerPicker({ value, onSelect }) {
  const [query, setQuery] = useState("");
  const [results, setResults] = useState([]);
  const [open, setOpen] = useState(false);
  const [loading, setLoading] = useState(false);

  useEffect(() => {
    if (!query.trim()) { setResults([]); return; }
    let active = true;
    setLoading(true);
    const t = setTimeout(async () => {
      const q = query.trim();
      const { data } = await supabase
        .from("customers")
        .select("*")
        .or(`full_name.ilike.%${q}%,phone.ilike.%${q}%,cccd.ilike.%${q}%`)
        .limit(8);
      if (active) { setResults(data || []); setLoading(false); }
    }, 300);
    return () => { active = false; clearTimeout(t); };
  }, [query]);

  if (value) {
    return (
      <div className="flex items-center justify-between bg-slate-50 rounded-xl px-3 py-2.5">
        <div>
          <p className="text-sm font-medium text-slate-700">{value.full_name}</p>
          <p className="text-xs text-slate-400">{value.phone || value.cccd || "—"}</p>
        </div>
        <button type="button" onClick={() => onSelect(null)} className="text-xs text-brand-600 hover:underline">Đổi</button>
      </div>
    );
  }

  return (
    <div className="relative">
      <div className="relative">
        <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-300" />
        <input
          value={query}
          onChange={(e) => { setQuery(e.target.value); setOpen(true); }}
          onFocus={() => setOpen(true)}
          placeholder="Tìm khách hàng theo tên, SĐT, CCCD..."
          className="w-full pl-9 pr-3 py-2 rounded-xl border border-slate-200 text-sm outline-none focus:ring-2 focus:ring-brand-400"
        />
      </div>
      {open && query.trim() && (
        <div className="absolute z-10 w-full bg-white border border-slate-100 rounded-xl shadow-lg mt-1 max-h-60 overflow-y-auto">
          {loading ? (
            <div className="p-3 text-xs text-slate-400 flex items-center gap-2"><Loader2 size={13} className="animate-spin" /> Đang tìm...</div>
          ) : results.length === 0 ? (
            <div className="p-3 text-xs text-slate-400">Không tìm thấy — kiểm tra lại tên/SĐT/CCCD.</div>
          ) : (
            results.map((c) => (
              <button
                type="button"
                key={c.id}
                onClick={() => { onSelect(c); setOpen(false); setQuery(""); }}
                className="w-full text-left px-3 py-2 hover:bg-slate-50 text-sm border-b border-slate-50 last:border-0"
              >
                <p className="font-medium text-slate-700">{c.full_name}</p>
                <p className="text-xs text-slate-400">{c.phone || "—"} {c.cccd ? `· CCCD ${c.cccd}` : ""}</p>
              </button>
            ))
          )}
          <p className="px-3 py-2 text-xs text-slate-400 border-t border-slate-50">
            Không thấy khách? Vào mục Khách hàng để thêm mới trước.
          </p>
        </div>
      )}
    </div>
  );
}

function DevicePicker({ value, onSelect }) {
  const [query, setQuery] = useState("");
  const [results, setResults] = useState([]);
  const [open, setOpen] = useState(false);
  const [loading, setLoading] = useState(false);

  useEffect(() => {
    let active = true;
    setLoading(true);
    const t = setTimeout(async () => {
      let req = supabase.from("devices").select("*").eq("status", "in_stock").order("created_at", { ascending: false }).limit(8);
      if (query.trim()) req = req.or(`imei.ilike.%${query.trim()}%,model.ilike.%${query.trim()}%`);
      const { data } = await req;
      if (active) { setResults(data || []); setLoading(false); }
    }, 300);
    return () => { active = false; clearTimeout(t); };
  }, [query]);

  if (value) {
    return (
      <div className="flex items-center justify-between bg-slate-50 rounded-xl px-3 py-2.5">
        <div>
          <p className="text-sm font-medium text-slate-700">{value.model} {[value.storage, value.color].filter(Boolean).join(" · ")}</p>
          <p className="text-xs text-slate-400">IMEI {value.imei}</p>
        </div>
        <button type="button" onClick={() => onSelect(null)} className="text-xs text-brand-600 hover:underline">Đổi</button>
      </div>
    );
  }

  return (
    <div className="relative">
      <div className="relative">
        <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-300" />
        <input
          value={query}
          onChange={(e) => { setQuery(e.target.value); setOpen(true); }}
          onFocus={() => setOpen(true)}
          placeholder="Tìm máy còn hàng theo IMEI, model..."
          className="w-full pl-9 pr-3 py-2 rounded-xl border border-slate-200 text-sm outline-none focus:ring-2 focus:ring-brand-400"
        />
      </div>
      {open && (
        <div className="absolute z-10 w-full bg-white border border-slate-100 rounded-xl shadow-lg mt-1 max-h-60 overflow-y-auto">
          {loading ? (
            <div className="p-3 text-xs text-slate-400 flex items-center gap-2"><Loader2 size={13} className="animate-spin" /> Đang tìm...</div>
          ) : results.length === 0 ? (
            <div className="p-3 text-xs text-slate-400">Không có máy còn hàng khớp tìm kiếm.</div>
          ) : (
            results.map((d) => (
              <button
                type="button"
                key={d.id}
                onClick={() => { onSelect(d); setOpen(false); setQuery(""); }}
                className="w-full text-left px-3 py-2 hover:bg-slate-50 text-sm border-b border-slate-50 last:border-0"
              >
                <p className="font-medium text-slate-700">{d.model} {[d.storage, d.color].filter(Boolean).join(" · ")}</p>
                <p className="text-xs text-slate-400">IMEI {d.imei} · Giá đề xuất {fmtVND(d.sale_price)}</p>
              </button>
            ))
          )}
        </div>
      )}
    </div>
  );
}

function PaymentRows({ rows, setRows, total }) {
  const paid = rows.reduce((s, r) => s + (Number(r.amount) || 0), 0);
  const remaining = total - paid;

  const update = (i, patch) => setRows((rs) => rs.map((r, idx) => (idx === i ? { ...r, ...patch } : r)));
  const addRow = () => setRows((rs) => [...rs, { method: "cash", amount: remaining > 0 ? remaining : "", installment_provider: "", installment_contract_code: "", note: "" }]);
  const removeRow = (i) => setRows((rs) => rs.filter((_, idx) => idx !== i));

  return (
    <div className="space-y-2">
      {rows.map((r, i) => {
        const Icon = PAYMENT_METHOD_ICONS[r.method] || Wallet;
        return (
          <div key={i} className="bg-slate-50 rounded-xl p-3 space-y-2">
            <div className="flex items-center gap-2">
              <Icon size={15} className="text-slate-400 shrink-0" />
              <select
                value={r.method}
                onChange={(e) => update(i, { method: e.target.value })}
                className="rounded-lg border border-slate-200 px-2 py-1.5 text-sm flex-1"
              >
                <option value="cash">Tiền mặt</option>
                <option value="bank_transfer">Chuyển khoản</option>
                <option value="installment">Trả góp</option>
              </select>
              <input
                type="number"
                value={r.amount}
                onChange={(e) => update(i, { amount: e.target.value })}
                placeholder="Số tiền"
                className="w-32 rounded-lg border border-slate-200 px-2 py-1.5 text-sm"
              />
              {rows.length > 1 && (
                <button type="button" onClick={() => removeRow(i)} className="text-rose-400 hover:text-rose-600"><X size={15} /></button>
              )}
            </div>
            {r.method === "installment" && (
              <div className="grid grid-cols-2 gap-2 pl-6">
                <input
                  value={r.installment_provider}
                  onChange={(e) => update(i, { installment_provider: e.target.value })}
                  placeholder="Đơn vị hỗ trợ (Mira...)"
                  className="rounded-lg border border-slate-200 px-2 py-1.5 text-xs"
                />
                <input
                  value={r.installment_contract_code}
                  onChange={(e) => update(i, { installment_contract_code: e.target.value })}
                  placeholder="Mã hồ sơ trả góp"
                  className="rounded-lg border border-slate-200 px-2 py-1.5 text-xs"
                />
              </div>
            )}
          </div>
        );
      })}
      <button type="button" onClick={addRow} className="text-xs text-brand-600 hover:underline flex items-center gap-1">
        <Plus size={13} /> Thêm hình thức thanh toán
      </button>
      <div className={classNames(
        "text-xs px-3 py-2 rounded-lg",
        remaining === 0 ? "bg-emerald-50 text-emerald-700" : "bg-amber-50 text-amber-700"
      )}>
        Đã phân bổ {fmtVND(paid)} / {fmtVND(total)} {remaining !== 0 && `— còn thiếu ${fmtVND(remaining)}`}
      </div>
    </div>
  );
}

function OrderForm({ onCancel, onSaved, employee }) {
  const [customer, setCustomer] = useState(null);
  const [device, setDevice] = useState(null);
  const [salePrice, setSalePrice] = useState("");
  const [discount, setDiscount] = useState("0");
  const [notes, setNotes] = useState("");
  const [payments, setPayments] = useState([{ method: "cash", amount: "", installment_provider: "", installment_contract_code: "", note: "" }]);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  useEffect(() => {
    if (device && !salePrice) setSalePrice(device.sale_price ?? "");
  }, [device]); // eslint-disable-line

  const total = Math.max(0, (Number(salePrice) || 0) - (Number(discount) || 0));

  const submit = async (e) => {
    e.preventDefault();
    setError("");
    if (!customer) { setError("Vui lòng chọn khách hàng."); return; }
    if (!device) { setError("Vui lòng chọn máy bán (còn hàng)."); return; }
    if (!salePrice || Number(salePrice) <= 0) { setError("Vui lòng nhập giá bán hợp lệ."); return; }
    const paidTotal = payments.reduce((s, r) => s + (Number(r.amount) || 0), 0);
    if (Math.round(paidTotal) !== Math.round(total)) { setError("Tổng các hình thức thanh toán phải bằng tổng tiền đơn hàng."); return; }
    if (payments.some((r) => !r.amount || Number(r.amount) <= 0)) { setError("Mỗi hình thức thanh toán cần nhập số tiền hợp lệ."); return; }

    setSaving(true);
    try {
      const orderPayload = {
        customer_id: customer.id,
        device_id: device.id,
        sale_price: Number(salePrice),
        discount: Number(discount) || 0,
        total_amount: total,
        notes: notes.trim() || null,
        created_by: employee.id,
        updated_by: employee.id,
      };
      const { data: order, error: orderErr } = await supabase.from("sales_orders").insert(orderPayload).select().maybeSingle();
      if (orderErr) throw orderErr;

      const paymentRows = payments.map((r) => ({
        order_id: order.id,
        method: r.method,
        amount: Number(r.amount),
        installment_provider: r.method === "installment" ? (r.installment_provider.trim() || null) : null,
        installment_contract_code: r.method === "installment" ? (r.installment_contract_code.trim() || null) : null,
        note: r.note?.trim() || null,
        created_by: employee.id,
      }));
      const { error: payErr } = await supabase.from("order_payments").insert(paymentRows);
      if (payErr) throw payErr;

      const { error: contractErr } = await supabase.from("contracts").insert({ order_id: order.id, created_by: employee.id });
      if (contractErr) throw contractErr;

      const { data: soldDevice, error: devErr } = await supabase
        .from("devices").update({ status: "sold", updated_by: employee.id }).eq("id", device.id).select().maybeSingle();
      if (devErr) throw devErr;

      await supabase.from("audit_logs").insert([
        { table_name: "sales_orders", record_id: order.id, action: "create", new_data: order, performed_by: employee.id },
        { table_name: "devices", record_id: device.id, action: "update", old_data: device, new_data: soldDevice, performed_by: employee.id },
      ]);

      onSaved();
    } catch (err) {
      setError(err.message || "Không tạo được đơn hàng, vui lòng thử lại.");
    } finally {
      setSaving(false);
    }
  };

  return (
    <Card className="p-4 sm:p-5 mb-5">
      <div className="flex items-center justify-between mb-4">
        <p className="font-semibold text-slate-800 text-sm">Tạo đơn hàng bán</p>
        <button onClick={onCancel} className="text-slate-400 hover:text-rose-600"><X size={16} /></button>
      </div>
      <form onSubmit={submit} className="space-y-4">
        <div>
          <span className="text-xs font-medium text-slate-500 mb-1 block">Khách hàng *</span>
          <CustomerPicker value={customer} onSelect={setCustomer} />
        </div>
        <div>
          <span className="text-xs font-medium text-slate-500 mb-1 block">Máy bán (IMEI) *</span>
          <DevicePicker value={device} onSelect={setDevice} />
        </div>
        <div className="grid sm:grid-cols-2 gap-3">
          <TextField label="Giá bán (đ) *" type="number" value={salePrice} onChange={(e) => setSalePrice(e.target.value)} />
          <TextField label="Giảm giá (đ)" type="number" value={discount} onChange={(e) => setDiscount(e.target.value)} />
        </div>
        <div className="flex justify-between items-center bg-brand-50 rounded-xl px-3 py-2.5 text-sm">
          <span className="text-slate-500">Tổng tiền đơn hàng</span>
          <span className="font-semibold text-brand-700">{fmtVND(total)}</span>
        </div>
        <div>
          <span className="text-xs font-medium text-slate-500 mb-1 block">Hình thức thanh toán *</span>
          <PaymentRows rows={payments} setRows={setPayments} total={total} />
        </div>
        <TextField label="Ghi chú" value={notes} onChange={(e) => setNotes(e.target.value)} />
        {error && <p className="text-xs text-rose-600 bg-rose-50 rounded-lg px-3 py-2">{error}</p>}
        <div className="flex gap-2">
          <button type="submit" disabled={saving} className="bg-brand-600 hover:bg-brand-700 text-white rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-2 disabled:opacity-60">
            {saving && <Loader2 size={14} className="animate-spin" />} Hoàn tất đơn hàng
          </button>
          <button type="button" onClick={onCancel} className="text-slate-500 text-sm px-4 py-2">Hủy</button>
        </div>
      </form>
    </Card>
  );
}

function OrderRow({ order, employee, onDeleted }) {
  const [expanded, setExpanded] = useState(false);
  const [detail, setDetail] = useState(null);
  const [loadingDetail, setLoadingDetail] = useState(false);
  const [printType, setPrintType] = useState(null);
  const canDelete = employee.role === "quan_ly";

  const loadDetail = async () => {
    if (detail) { setExpanded((s) => !s); return; }
    setLoadingDetail(true);
    const [{ data: customer }, { data: device }, { data: payments }, { data: contract }] = await Promise.all([
      supabase.from("customers").select("*").eq("id", order.customer_id).maybeSingle(),
      supabase.from("devices").select("*").eq("id", order.device_id).maybeSingle(),
      supabase.from("order_payments").select("*").eq("order_id", order.id).order("created_at", { ascending: true }),
      supabase.from("contracts").select("*").eq("order_id", order.id).maybeSingle(),
    ]);
    setDetail({ customer, device, payments: payments || [], contract });
    setLoadingDetail(false);
    setExpanded(true);
  };

  const remove = async () => {
    if (!confirm(`Xóa đơn hàng "${order.order_code}"? Máy sẽ không tự động chuyển lại trạng thái Còn hàng.`)) return;
    const { error } = await supabase.from("sales_orders").delete().eq("id", order.id);
    if (error) { alert(error.message); return; }
    await supabase.from("audit_logs").insert({ table_name: "sales_orders", record_id: order.id, action: "delete", old_data: order, performed_by: employee.id });
    onDeleted();
  };

  return (
    <>
      <tr className="border-b border-slate-50 last:border-0 hover:bg-slate-50/60 cursor-pointer" onClick={loadDetail}>
        <td className="px-3 py-2.5 font-medium text-slate-700 whitespace-nowrap">{order.order_code}</td>
        <td className="px-3 py-2.5 text-slate-500">{fmtDate(order.created_at)}</td>
        <td className="px-3 py-2.5 text-slate-600">{fmtVND(order.total_amount)}</td>
        <td className="px-3 py-2.5">
          <span className={classNames("text-xs px-2 py-0.5 rounded-full", order.status === "completed" ? "bg-emerald-50 text-emerald-700" : "bg-rose-50 text-rose-600")}>
            {order.status === "completed" ? "Hoàn tất" : "Đã hủy"}
          </span>
        </td>
        <td className="px-3 py-2.5 text-right" onClick={(e) => e.stopPropagation()}>
          {loadingDetail ? <Loader2 size={14} className="animate-spin inline text-slate-300" /> : (
            <ChevronDown size={15} className={classNames("inline text-slate-300 transition-transform", expanded && "rotate-180")} onClick={loadDetail} />
          )}
        </td>
      </tr>
      {expanded && detail && (
        <tr>
          <td colSpan={5} className="bg-slate-50/70 px-4 py-4">
            <div className="grid sm:grid-cols-2 gap-4 text-sm">
              <div>
                <p className="text-xs text-slate-400 mb-1">Khách hàng</p>
                <p className="font-medium text-slate-700">{detail.customer?.full_name}</p>
                <p className="text-xs text-slate-400">{detail.customer?.phone}</p>
              </div>
              <div>
                <p className="text-xs text-slate-400 mb-1">Máy đã bán</p>
                <p className="font-medium text-slate-700">{detail.device?.model} {[detail.device?.storage, detail.device?.color].filter(Boolean).join(" · ")}</p>
                <p className="text-xs text-slate-400">IMEI {detail.device?.imei}</p>
              </div>
            </div>
            <div className="mt-3 space-y-1">
              <div className="flex items-center justify-between mb-1">
                <p className="text-xs text-slate-400">Thanh toán ({detail.payments.length} hình thức)</p>
                {detail.payments.length > 1 && (
                  <button onClick={() => setPrintType({ kind: "receipt", payments: detail.payments })} className="text-brand-600 hover:underline text-xs flex items-center gap-1">
                    <Printer size={12} /> In phiếu thu gộp ({detail.payments.length} hình thức)
                  </button>
                )}
              </div>
              {detail.payments.map((p) => (
                <div key={p.id} className="flex items-center justify-between text-sm bg-white rounded-lg px-3 py-2">
                  <span className="text-slate-600">{p.payment_code} · {PAYMENT_METHOD_LABELS[p.method]}</span>
                  <span className="flex items-center gap-3">
                    {fmtVND(p.amount)}
                    <button onClick={() => setPrintType({ kind: "receipt", payments: [p] })} className="text-slate-400 hover:text-brand-600 hover:underline text-xs flex items-center gap-1">
                      <Printer size={12} /> In riêng
                    </button>
                  </span>
                </div>
              ))}
            </div>
            {order.notes && <p className="text-xs text-slate-400 mt-2">Ghi chú: {order.notes}</p>}
            <div className="flex gap-3 mt-3">
              <button onClick={() => setPrintType({ kind: "contract" })} className="text-brand-600 hover:underline text-xs flex items-center gap-1">
                <Printer size={12} /> In hợp đồng
              </button>
              {canDelete && (
                <button onClick={remove} className="text-rose-500 hover:underline text-xs flex items-center gap-1">
                  <Trash2 size={12} /> Xóa đơn
                </button>
              )}
            </div>
          </td>
        </tr>
      )}
      {printType && detail && (
        <PrintDocModal
          type={printType.kind}
          order={order}
          customer={detail.customer}
          device={detail.device}
          payments={printType.kind === "receipt" ? printType.payments : detail.payments}
          contract={detail.contract}
          onClose={() => setPrintType(null)}
        />
      )}
    </>
  );
}

function KiotVietExportModal({ onClose }) {
  const [fromDate, setFromDate] = useState("");
  const [toDate, setToDate] = useState("");
  const [exporting, setExporting] = useState(false);
  const [error, setError] = useState("");

  const runExport = async () => {
    setError(""); setExporting(true);
    try {
      let query = supabase
        .from("sales_orders")
        .select("*, customers(customer_code, full_name, phone, address), devices(imei, model, storage, color)")
        .eq("status", "completed")
        .order("created_at", { ascending: true });
      if (fromDate) query = query.gte("created_at", `${fromDate}T00:00:00`);
      if (toDate) query = query.lte("created_at", `${toDate}T23:59:59`);
      const { data, error: err } = await query;
      if (err) throw err;
      if (!data || data.length === 0) { setError("Không có đơn hàng nào trong khoảng ngày đã chọn."); setExporting(false); return; }

      const rows = data.map((o) => ({
        "Mã hàng": o.devices ? toKiotVietProductCode(o.devices) : "",
        "Tên hàng": o.devices ? [o.devices.model, o.devices.storage, o.devices.color].filter(Boolean).join(" ") : "",
        "Đơn giá": Number(o.sale_price) || 0,
        "Giảm giá": Number(o.discount) || 0,
        "Giảm giá (%)": "",
        "Số lượng": 1,
        "Serial/IMEI": o.devices?.imei || "",
        "Mã khách hàng": o.customers?.customer_code || "",
        "Tên khách hàng": o.customers?.full_name || "",
        "Số điện thoại": o.customers?.phone || "",
        "Loại khách": "Individual",
        "Địa chỉ": o.customers?.address || "",
        "Khu vực": "",
        "Phường xã": "",
      }));

      const ws = XLSX.utils.json_to_sheet(rows, {
        header: ["Mã hàng", "Tên hàng", "Đơn giá", "Giảm giá", "Giảm giá (%)", "Số lượng", "Serial/IMEI", "Mã khách hàng", "Tên khách hàng", "Số điện thoại", "Loại khách", "Địa chỉ", "Khu vực", "Phường xã"],
      });
      const wb = XLSX.utils.book_new();
      XLSX.utils.book_append_sheet(wb, ws, "ImportProductTemplate");
      const stamp = new Date().toISOString().slice(0, 10);
      XLSX.writeFile(wb, `KiotViet-DonHang-${stamp}.xlsx`);
      onClose();
    } catch (err) {
      setError(err.message || "Không xuất được file, vui lòng thử lại.");
    } finally {
      setExporting(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/40 flex items-center justify-center p-4">
      <Card className="w-full max-w-sm p-5">
        <div className="flex items-center justify-between mb-4">
          <p className="font-semibold text-slate-800 text-sm">Xuất Excel KiotViet</p>
          <button onClick={onClose} className="text-slate-400 hover:text-rose-600"><X size={16} /></button>
        </div>
        <p className="text-xs text-slate-400 mb-3">
          Xuất danh sách đơn hàng đã hoàn tất theo đúng cấu trúc file mẫu import KiotViet (để trống ngày = xuất toàn bộ).
        </p>
        <div className="grid grid-cols-2 gap-3 mb-4">
          <TextField label="Từ ngày" type="date" value={fromDate} onChange={(e) => setFromDate(e.target.value)} />
          <TextField label="Đến ngày" type="date" value={toDate} onChange={(e) => setToDate(e.target.value)} />
        </div>
        {error && <p className="text-xs text-rose-600 bg-rose-50 rounded-lg px-3 py-2 mb-3">{error}</p>}
        <div className="flex gap-2">
          <button onClick={runExport} disabled={exporting} className="bg-brand-600 hover:bg-brand-700 text-white rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-2 disabled:opacity-60">
            {exporting && <Loader2 size={14} className="animate-spin" />} Xuất file
          </button>
          <button onClick={onClose} className="text-slate-500 text-sm px-4 py-2">Hủy</button>
        </div>
      </Card>
    </div>
  );
}

function OrdersModule({ employee }) {
  const [orders, setOrders] = useState([]);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState("");
  const [showForm, setShowForm] = useState(false);
  const [showExport, setShowExport] = useState(false);

  const canCreate = employee.role !== "ke_toan";

  const load = useCallback(async () => {
    setLoading(true);
    const { data, error } = await supabase.from("sales_orders").select("*").order("created_at", { ascending: false }).limit(1000);
    if (!error) setOrders(data || []);
    setLoading(false);
  }, []);

  useEffect(() => { load(); }, [load]);

  const filtered = orders.filter((o) => {
    if (!search.trim()) return true;
    return o.order_code?.toLowerCase().includes(search.trim().toLowerCase());
  });

  return (
    <div>
      <div className="flex items-center justify-between mb-4 gap-2 flex-wrap">
        <div>
          <h2 className="text-lg font-semibold text-slate-800">Đơn hàng bán</h2>
          <p className="text-xs text-slate-400">{orders.length} đơn hàng</p>
        </div>
        <div className="flex gap-2">
          <button onClick={() => setShowExport(true)} className="bg-white border border-slate-200 hover:bg-slate-50 text-slate-600 rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-1.5">
            <FileSpreadsheet size={15} /> Xuất Excel KiotViet
          </button>
          {canCreate && (
            <button onClick={() => setShowForm((s) => !s)} className="bg-brand-600 hover:bg-brand-700 text-white rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-1.5">
              <Plus size={15} /> Tạo đơn hàng
            </button>
          )}
        </div>
      </div>

      {showExport && <KiotVietExportModal onClose={() => setShowExport(false)} />}

      {showForm && (
        <OrderForm employee={employee} onCancel={() => setShowForm(false)} onSaved={() => { setShowForm(false); load(); }} />
      )}

      <Card className="p-0 overflow-hidden">
        <div className="p-3 border-b border-slate-100">
          <div className="relative">
            <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-300" />
            <input
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="Tìm theo mã đơn hàng..."
              className="w-full pl-9 pr-3 py-2 rounded-xl border border-slate-200 text-sm outline-none focus:ring-2 focus:ring-brand-400"
            />
          </div>
        </div>
        {loading ? (
          <div className="flex justify-center py-10"><Loader2 className="animate-spin text-slate-300" /></div>
        ) : filtered.length === 0 ? (
          <EmptyState icon={ShoppingCart} text="Chưa có đơn hàng nào." />
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead><tr className="text-left text-xs text-slate-400 border-b border-slate-100">
                <th className="px-3 py-2">Mã đơn</th>
                <th className="px-3 py-2">Ngày tạo</th>
                <th className="px-3 py-2">Tổng tiền</th>
                <th className="px-3 py-2">Trạng thái</th>
                <th className="px-3 py-2"></th>
              </tr></thead>
              <tbody>
                {filtered.map((o) => <OrderRow key={o.id} order={o} employee={employee} onDeleted={load} />)}
              </tbody>
            </table>
          </div>
        )}
      </Card>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Receipts module — global ledger of all phiếu thu/chi                  */
/* ---------------------------------------------------------------------- */

function ReceiptsModule({ employee }) {
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [printData, setPrintData] = useState(null);

  const load = useCallback(async () => {
    setLoading(true);
    const { data } = await supabase
      .from("order_payments")
      .select("*, sales_orders(order_code, customer_id, device_id, sale_price, discount, total_amount, notes, created_at)")
      .order("created_at", { ascending: false })
      .limit(500);
    setRows(data || []);
    setLoading(false);
  }, []);

  useEffect(() => { load(); }, [load]);

  // Gộp các phiếu cùng 1 đơn thành 1 nhóm — mỗi đơn chỉ hiện 1 dòng trong bảng,
  // dù có bao nhiêu hình thức thanh toán, tránh lặp lại và nhầm lẫn.
  const groups = [];
  const groupIndexByOrder = {};
  for (const r of rows) {
    const key = r.sales_orders?.order_code || r.id;
    if (groupIndexByOrder[key] === undefined) {
      groupIndexByOrder[key] = groups.length;
      groups.push({ orderCode: key, sales_orders: r.sales_orders, payments: [r] });
    } else {
      groups[groupIndexByOrder[key]].payments.push(r);
    }
  }

  const openPrint = async (group) => {
    const first = group.payments[0];
    const [{ data: customer }, { data: device }] = await Promise.all([
      supabase.from("customers").select("*").eq("id", first.sales_orders.customer_id).maybeSingle(),
      supabase.from("devices").select("*").eq("id", first.sales_orders.device_id).maybeSingle(),
    ]);
    setPrintData({ row: first, customer, device, payments: group.payments });
  };

  return (
    <div>
      <div className="mb-4">
        <h2 className="text-lg font-semibold text-slate-800">Phiếu thu/chi</h2>
        <p className="text-xs text-slate-400">Mỗi đơn hàng 1 dòng — đơn có nhiều hình thức thanh toán in ra sẽ gộp thành 1 phiếu duy nhất</p>
      </div>
      <Card className="p-0 overflow-hidden">
        {loading ? (
          <div className="flex justify-center py-10"><Loader2 className="animate-spin text-slate-300" /></div>
        ) : groups.length === 0 ? (
          <EmptyState icon={Receipt} text="Chưa có phiếu thu/chi nào." />
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead><tr className="text-left text-xs text-slate-400 border-b border-slate-100">
                <th className="px-3 py-2">Mã phiếu</th>
                <th className="px-3 py-2">Đơn hàng</th>
                <th className="px-3 py-2">Hình thức</th>
                <th className="px-3 py-2">Số tiền</th>
                <th className="px-3 py-2">Ngày</th>
                <th className="px-3 py-2"></th>
              </tr></thead>
              <tbody>
                {groups.map((g) => {
                  const total = g.payments.reduce((s, p) => s + Number(p.amount || 0), 0);
                  const methods = g.payments.map((p) => PAYMENT_METHOD_LABELS[p.method]).join(" + ");
                  const codes = g.payments.map((p) => p.payment_code).join(", ");
                  return (
                    <tr key={g.orderCode} className="border-b border-slate-50 last:border-0 hover:bg-slate-50/60">
                      <td className="px-3 py-2.5 font-medium text-slate-700 whitespace-nowrap">{codes}</td>
                      <td className="px-3 py-2.5 text-slate-500">{g.orderCode}</td>
                      <td className="px-3 py-2.5 text-slate-500">{methods}</td>
                      <td className="px-3 py-2.5 text-slate-600">{fmtVND(total)}</td>
                      <td className="px-3 py-2.5 text-slate-500">{fmtDate(g.payments[0].created_at)}</td>
                      <td className="px-3 py-2.5 text-right whitespace-nowrap">
                        <button onClick={() => openPrint(g)} className="text-brand-600 hover:underline text-xs flex items-center gap-1 ml-auto">
                          <Printer size={12} /> {g.payments.length > 1 ? `In gộp (${g.payments.length} hình thức)` : "In"}
                        </button>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </Card>
      {printData && (
        <PrintDocModal
          type="receipt"
          order={printData.row.sales_orders}
          customer={printData.customer}
          device={printData.device}
          payments={printData.payments}
          onClose={() => setPrintData(null)}
        />
      )}
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Audit Log module — lịch sử thao tác toàn hệ thống (quan_ly + ke_toan)  */
/* ---------------------------------------------------------------------- */

const AUDIT_TABLE_LABELS = {
  devices: "Kho hàng",
  sales_orders: "Đơn hàng bán",
  customers: "Khách hàng",
};
const AUDIT_ACTION_LABELS = { create: "Tạo mới", update: "Cập nhật", delete: "Xóa" };
const AUDIT_ACTION_STYLES = {
  create: "bg-emerald-50 text-emerald-700",
  update: "bg-amber-50 text-amber-700",
  delete: "bg-rose-50 text-rose-600",
};

// Vài trường đáng chú ý để tóm tắt thay đổi (nếu có), tránh dump cả object JSON khó đọc
const AUDIT_FIELD_LABELS = {
  imei: "IMEI", model: "Model", status: "Trạng thái", cost_price: "Giá vốn", sale_price: "Giá bán",
  full_name: "Họ tên", phone: "SĐT", order_code: "Mã đơn", total_amount: "Tổng tiền",
};

function summarizeAuditChange(log) {
  if (log.action === "delete") {
    const d = log.old_data || {};
    return d.imei ? `IMEI ${d.imei}` : d.order_code || d.full_name || d.id || "";
  }
  if (log.action === "create") {
    const d = log.new_data || {};
    return d.imei ? `IMEI ${d.imei} — ${d.model || ""}` : d.order_code || d.full_name || "";
  }
  // update: liệt kê các trường thay đổi
  const oldD = log.old_data || {};
  const newD = log.new_data || {};
  const changes = [];
  for (const key of Object.keys(AUDIT_FIELD_LABELS)) {
    if (key in newD && String(oldD[key]) !== String(newD[key])) {
      const label = AUDIT_FIELD_LABELS[key];
      const isMoney = key.includes("price") || key === "total_amount";
      const fmt = isMoney ? fmtVND : (v) => (key === "status" ? (DEVICE_STATUS_LABELS[v] || v) : v);
      changes.push(`${label}: ${fmt(oldD[key])} → ${fmt(newD[key])}`);
    }
  }
  return changes.length ? changes.join(" · ") : "Cập nhật thông tin";
}

function AuditLogModule() {
  const [logs, setLogs] = useState([]);
  const [loading, setLoading] = useState(true);
  const [tableFilter, setTableFilter] = useState("all");
  const [search, setSearch] = useState("");

  const load = useCallback(async () => {
    setLoading(true);
    const { data, error } = await supabase
      .from("audit_logs")
      .select("*, employees(full_name)")
      .order("created_at", { ascending: false })
      .limit(500);
    if (!error) setLogs(data || []);
    setLoading(false);
  }, []);

  useEffect(() => { load(); }, [load]);

  const filtered = logs.filter((l) => {
    if (tableFilter !== "all" && l.table_name !== tableFilter) return false;
    if (!search.trim()) return true;
    const q = search.trim().toLowerCase();
    const summary = summarizeAuditChange(l).toLowerCase();
    const performer = (l.employees?.full_name || "").toLowerCase();
    return summary.includes(q) || performer.includes(q);
  });

  return (
    <div>
      <div className="mb-4">
        <h2 className="text-lg font-semibold text-slate-800">Audit Log</h2>
        <p className="text-xs text-slate-400">Lịch sử thao tác trên toàn hệ thống — chỉ Quản lý & Kế toán xem được</p>
      </div>
      <Card className="p-0 overflow-hidden">
        <div className="p-3 border-b border-slate-100 flex flex-col sm:flex-row gap-2">
          <div className="relative flex-1">
            <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-300" />
            <input
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="Tìm theo nội dung thay đổi, người thực hiện..."
              className="w-full pl-9 pr-3 py-2 rounded-xl border border-slate-200 text-sm outline-none focus:ring-2 focus:ring-brand-400"
            />
          </div>
          <select value={tableFilter} onChange={(e) => setTableFilter(e.target.value)} className="rounded-xl border border-slate-200 px-3 py-2 text-sm">
            <option value="all">Tất cả</option>
            <option value="devices">Kho hàng</option>
            <option value="sales_orders">Đơn hàng bán</option>
            <option value="customers">Khách hàng</option>
          </select>
        </div>
        {loading ? (
          <div className="flex justify-center py-10"><Loader2 className="animate-spin text-slate-300" /></div>
        ) : filtered.length === 0 ? (
          <EmptyState icon={ScrollText} text="Chưa có nhật ký thao tác nào." />
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead><tr className="text-left text-xs text-slate-400 border-b border-slate-100">
                <th className="px-3 py-2">Thời gian</th>
                <th className="px-3 py-2">Bảng</th>
                <th className="px-3 py-2">Hành động</th>
                <th className="px-3 py-2">Nội dung</th>
                <th className="px-3 py-2">Người thực hiện</th>
              </tr></thead>
              <tbody>
                {filtered.map((l) => (
                  <tr key={l.id} className="border-b border-slate-50 last:border-0 hover:bg-slate-50/60">
                    <td className="px-3 py-2.5 text-slate-500 whitespace-nowrap">{fmtDate(l.created_at)} {new Date(l.created_at).toLocaleTimeString("vi-VN", { hour: "2-digit", minute: "2-digit" })}</td>
                    <td className="px-3 py-2.5 text-slate-600">{AUDIT_TABLE_LABELS[l.table_name] || l.table_name}</td>
                    <td className="px-3 py-2.5">
                      <span className={classNames("text-xs px-2 py-0.5 rounded-full", AUDIT_ACTION_STYLES[l.action])}>
                        {AUDIT_ACTION_LABELS[l.action] || l.action}
                      </span>
                    </td>
                    <td className="px-3 py-2.5 text-slate-500 max-w-[320px] truncate" title={summarizeAuditChange(l)}>{summarizeAuditChange(l)}</td>
                    <td className="px-3 py-2.5 text-slate-500">{l.employees?.full_name || "—"}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Card>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Reports module — doanh thu/lợi nhuận (quan_ly + ke_toan)               */
/* ---------------------------------------------------------------------- */

function startOfMonth() {
  const d = new Date();
  return new Date(d.getFullYear(), d.getMonth(), 1).toISOString().slice(0, 10);
}
function todayStr() {
  return new Date().toISOString().slice(0, 10);
}

function ReportKpiCard({ label, value, icon: Icon, sub }) {
  return (
    <Card className="p-4 sm:p-5">
      <div className="flex items-center justify-between mb-2">
        <span className="text-xs font-medium text-slate-400">{label}</span>
        <Icon size={16} className="text-slate-300" />
      </div>
      <p className="text-xl font-semibold text-slate-800">{value}</p>
      {sub && <p className="text-xs text-slate-400 mt-1">{sub}</p>}
    </Card>
  );
}

function ReportsModule({ employee }) {
  const [fromDate, setFromDate] = useState(startOfMonth());
  const [toDate, setToDate] = useState(todayStr());
  const [orders, setOrders] = useState([]);
  const [inventoryValue, setInventoryValue] = useState(0);
  const [inStockCount, setInStockCount] = useState(0);
  const [loading, setLoading] = useState(true);

  const canSeeProfit = employee.role !== "nhan_vien";

  const load = useCallback(async () => {
    setLoading(true);
    const [{ data: orderData }, { data: deviceData }] = await Promise.all([
      supabase
        .from("sales_orders")
        .select("*, devices(model, storage, color, cost_price), employees:created_by(full_name)")
        .eq("status", "completed")
        .gte("created_at", `${fromDate}T00:00:00`)
        .lte("created_at", `${toDate}T23:59:59`)
        .order("created_at", { ascending: true }),
      supabase.from("devices").select("cost_price, status"),
    ]);
    setOrders(orderData || []);
    const inStock = (deviceData || []).filter((d) => d.status === "in_stock");
    setInventoryValue(inStock.reduce((s, d) => s + Number(d.cost_price || 0), 0));
    setInStockCount(inStock.length);
    setLoading(false);
  }, [fromDate, toDate]);

  useEffect(() => { load(); }, [load]);

  const totalRevenue = orders.reduce((s, o) => s + Number(o.total_amount || 0), 0);
  const totalProfit = orders.reduce((s, o) => s + (Number(o.total_amount || 0) - Number(o.devices?.cost_price || 0)), 0);
  const totalOrders = orders.length;
  const avgOrderValue = totalOrders ? totalRevenue / totalOrders : 0;

  // Doanh thu theo ngày cho biểu đồ
  const byDay = {};
  for (const o of orders) {
    const day = fmtDate(o.created_at);
    byDay[day] = (byDay[day] || 0) + Number(o.total_amount || 0);
  }
  const chartData = Object.entries(byDay).map(([day, revenue]) => ({ day, revenue }));

  // Top nhân viên theo doanh số
  const byEmployee = {};
  for (const o of orders) {
    const name = o.employees?.full_name || "Không rõ";
    if (!byEmployee[name]) byEmployee[name] = { revenue: 0, count: 0 };
    byEmployee[name].revenue += Number(o.total_amount || 0);
    byEmployee[name].count += 1;
  }
  const topEmployees = Object.entries(byEmployee).sort((a, b) => b[1].revenue - a[1].revenue).slice(0, 5);

  // Top model bán chạy
  const byModel = {};
  for (const o of orders) {
    const name = o.devices ? [o.devices.model, o.devices.storage, o.devices.color].filter(Boolean).join(" ") : "—";
    byModel[name] = (byModel[name] || 0) + 1;
  }
  const topModels = Object.entries(byModel).sort((a, b) => b[1] - a[1]).slice(0, 5);

  return (
    <div>
      <div className="flex items-center justify-between mb-4 gap-2 flex-wrap">
        <div>
          <h2 className="text-lg font-semibold text-slate-800">Báo cáo</h2>
          <p className="text-xs text-slate-400">Doanh thu, lợi nhuận và hiệu suất bán hàng</p>
        </div>
        <div className="flex items-center gap-2 bg-white border border-slate-200 rounded-xl px-3 py-1.5">
          <Filter size={14} className="text-slate-400" />
          <input type="date" value={fromDate} onChange={(e) => setFromDate(e.target.value)} className="text-sm outline-none" />
          <span className="text-slate-300">—</span>
          <input type="date" value={toDate} onChange={(e) => setToDate(e.target.value)} className="text-sm outline-none" />
        </div>
      </div>

      {loading ? (
        <div className="flex justify-center py-10"><Loader2 className="animate-spin text-slate-300" /></div>
      ) : (
        <>
          <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 sm:gap-4 mb-4">
            <ReportKpiCard label="Doanh thu" value={fmtVND(totalRevenue)} icon={TrendingUp} sub={`${totalOrders} đơn hàng`} />
            {canSeeProfit && <ReportKpiCard label="Lợi nhuận gộp" value={fmtVND(totalProfit)} icon={Award} sub={totalRevenue ? `Biên LN ${((totalProfit / totalRevenue) * 100).toFixed(1)}%` : ""} />}
            <ReportKpiCard label="Giá trị đơn TB" value={fmtVND(avgOrderValue)} icon={Receipt} />
            {canSeeProfit && <ReportKpiCard label="Tồn kho hiện tại" value={fmtVND(inventoryValue)} icon={Package} sub={`${inStockCount} máy còn hàng`} />}
          </div>

          <Card className="p-4 sm:p-5 mb-4">
            <p className="text-sm font-medium text-slate-700 mb-3">Doanh thu theo ngày</p>
            {chartData.length === 0 ? (
              <p className="text-xs text-slate-400 py-6 text-center">Không có dữ liệu trong khoảng ngày đã chọn.</p>
            ) : (
              <div style={{ width: "100%", height: 240 }}>
                <ResponsiveContainer>
                  <RBarChart data={chartData}>
                    <CartesianGrid strokeDasharray="3 3" stroke="#f1f5f9" />
                    <XAxis dataKey="day" tick={{ fontSize: 11 }} />
                    <YAxis tick={{ fontSize: 11 }} tickFormatter={(v) => `${(v / 1000000).toFixed(0)}tr`} />
                    <RTooltip formatter={(v) => fmtVND(v)} />
                    <Bar dataKey="revenue" fill="#2563eb" radius={[4, 4, 0, 0]} />
                  </RBarChart>
                </ResponsiveContainer>
              </div>
            )}
          </Card>

          <div className="grid sm:grid-cols-2 gap-4">
            <Card className="p-4 sm:p-5">
              <p className="text-sm font-medium text-slate-700 mb-3">Top nhân viên theo doanh số</p>
              {topEmployees.length === 0 ? <p className="text-xs text-slate-400">Chưa có dữ liệu.</p> : (
                <ul className="space-y-2">
                  {topEmployees.map(([name, v], i) => (
                    <li key={name} className="flex items-center justify-between text-sm">
                      <span className="text-slate-600">{i + 1}. {name} <span className="text-slate-400 text-xs">({v.count} đơn)</span></span>
                      <span className="font-medium text-slate-700">{fmtVND(v.revenue)}</span>
                    </li>
                  ))}
                </ul>
              )}
            </Card>
            <Card className="p-4 sm:p-5">
              <p className="text-sm font-medium text-slate-700 mb-3">Model bán chạy</p>
              {topModels.length === 0 ? <p className="text-xs text-slate-400">Chưa có dữ liệu.</p> : (
                <ul className="space-y-2">
                  {topModels.map(([name, count], i) => (
                    <li key={name} className="flex items-center justify-between text-sm">
                      <span className="text-slate-600">{i + 1}. {name}</span>
                      <span className="font-medium text-slate-700">{count} máy</span>
                    </li>
                  ))}
                </ul>
              )}
            </Card>
          </div>
        </>
      )}
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Employees module (quan_ly only) — invite by email                     */
/* ---------------------------------------------------------------------- */

function EmployeesModule({ employee }) {
  const [employees, setEmployees] = useState([]);
  const [loading, setLoading] = useState(true);
  const [showForm, setShowForm] = useState(false);
  const [form, setForm] = useState({ full_name: "", email: "", role: "nhan_vien" });
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const load = useCallback(async () => {
    setLoading(true);
    const { data, error } = await supabase.from("employees").select("*").order("created_at", { ascending: true });
    if (!error) setEmployees(data || []);
    setLoading(false);
  }, []);
  useEffect(() => { load(); }, [load]);

  const invite = async (e) => {
    e.preventDefault();
    if (!form.full_name.trim() || !form.email.trim()) { setError("Vui lòng nhập đủ tên và tên đăng nhập."); return; }
    setError(""); setSaving(true);
    try {
      const { error: err } = await supabase.from("employees").insert({
        full_name: form.full_name.trim(), email: toLoginEmail(form.email), role: form.role, is_active: true,
      });
      if (err) throw err;
      setForm({ full_name: "", email: "", role: "nhan_vien" });
      setShowForm(false);
      load();
    } catch (err) {
      setError(err.message || "Không thêm được, vui lòng thử lại.");
    } finally {
      setSaving(false);
    }
  };

  const toggleActive = async (emp) => {
    const { error } = await supabase.from("employees").update({ is_active: !emp.is_active }).eq("id", emp.id);
    if (error) { alert(error.message); return; }
    load();
  };

  return (
    <div>
      <div className="flex items-center justify-between mb-4">
        <h2 className="text-lg font-semibold text-slate-800">Nhân viên</h2>
        <button onClick={() => setShowForm((s) => !s)} className="bg-brand-600 hover:bg-brand-700 text-white rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-1.5">
          <Plus size={15} /> Mời nhân viên
        </button>
      </div>

      {showForm && (
        <Card className="p-4 sm:p-5 mb-5">
          <form onSubmit={invite} className="grid sm:grid-cols-3 gap-3">
            <TextField label="Họ tên" value={form.full_name} onChange={(e) => setForm((f) => ({ ...f, full_name: e.target.value }))} />
            <TextField label="Tên đăng nhập (số điện thoại)" type="text" value={form.email} onChange={(e) => setForm((f) => ({ ...f, email: e.target.value }))} placeholder="0914657111" />
            <label className="block">
              <span className="text-xs font-medium text-slate-500 mb-1 block">Vai trò</span>
              <select value={form.role} onChange={(e) => setForm((f) => ({ ...f, role: e.target.value }))} className="w-full rounded-xl border border-slate-200 px-3 py-2 text-sm">
                <option value="nhan_vien">Nhân viên</option>
                <option value="quan_ly">Quản lý cửa hàng</option>
                <option value="ke_toan">Kế toán thuế</option>
              </select>
            </label>
            {error && <p className="text-xs text-rose-600 sm:col-span-3">{error}</p>}
            <div className="sm:col-span-3 flex gap-2">
              <button type="submit" disabled={saving} className="bg-brand-600 hover:bg-brand-700 text-white rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-2 disabled:opacity-60">
                {saving && <Loader2 size={14} className="animate-spin" />} Thêm
              </button>
              <p className="text-xs text-slate-400 self-center">
                Nhân viên sẽ vào màn đăng nhập, chọn "Tạo tài khoản lần đầu" và đăng ký đúng tên đăng nhập này để tự kích hoạt.
              </p>
            </div>
          </form>
        </Card>
      )}

      <Card className="p-0 overflow-hidden">
        {loading ? (
          <div className="flex justify-center py-10"><Loader2 className="animate-spin text-slate-300" /></div>
        ) : (
          <table className="w-full text-sm">
            <thead><tr className="text-left text-xs text-slate-400 border-b border-slate-100">
              <th className="px-3 py-2">Họ tên</th>
              <th className="px-3 py-2">Tên đăng nhập</th>
              <th className="px-3 py-2">Vai trò</th>
              <th className="px-3 py-2">Trạng thái</th>
              <th className="px-3 py-2"></th>
            </tr></thead>
            <tbody>
              {employees.map((e) => (
                <tr key={e.id} className="border-b border-slate-50 last:border-0">
                  <td className="px-3 py-2.5 font-medium text-slate-700">{e.full_name}</td>
                  <td className="px-3 py-2.5 text-slate-500">{fromLoginEmail(e.email)}</td>
                  <td className="px-3 py-2.5 text-slate-500">{ROLE_LABELS[e.role] || e.role}</td>
                  <td className="px-3 py-2.5">
                    <span className={classNames("text-xs px-2 py-0.5 rounded-full", e.is_active ? "bg-emerald-50 text-emerald-700" : "bg-slate-100 text-slate-400")}>
                      {e.is_active ? "Đang hoạt động" : "Đã khóa"}
                    </span>
                    {!e.user_id && <span className="text-xs text-amber-600 ml-2">Chưa kích hoạt</span>}
                  </td>
                  <td className="px-3 py-2.5 text-right">
                    {e.id !== employee.id && (
                      <button onClick={() => toggleActive(e)} className="text-xs text-slate-500 hover:underline">
                        {e.is_active ? "Khóa" : "Mở khóa"}
                      </button>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        )}
      </Card>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Placeholder module for not-yet-built phases                           */
/* ---------------------------------------------------------------------- */

function ComingSoonModule({ title, icon: Icon, phase }) {
  return (
    <Card className="p-10 flex flex-col items-center text-center">
      <Icon size={32} className="text-slate-300 mb-3" />
      <h2 className="text-base font-semibold text-slate-700 mb-1">{title}</h2>
      <p className="text-xs text-slate-400">Sẽ được xây dựng ở {phase}</p>
    </Card>
  );
}

/* ---------------------------------------------------------------------- */
/* Shell: sidebar + bottom nav + routing between modules                 */
/* ---------------------------------------------------------------------- */

const NAV_ITEMS = [
  { key: "dashboard", label: "Dashboard", icon: LayoutDashboard },
  { key: "customers", label: "Khách hàng", icon: Users },
  { key: "inventory", label: "Kho hàng", icon: Smartphone },
  { key: "orders", label: "Đơn hàng bán", icon: ShoppingCart },
  { key: "invoices", label: "Hóa đơn", icon: FileText, phase: "Phase 4" },
  { key: "receipts", label: "Phiếu thu/chi", icon: Receipt },
  { key: "reports", label: "Báo cáo", icon: BarChart3, allowedRoles: ["quan_ly", "ke_toan"] },
  { key: "employees", label: "Nhân viên", icon: UserCog, managerOnly: true },
  { key: "audit", label: "Audit Log", icon: ScrollText, allowedRoles: ["quan_ly", "ke_toan"] },
];

function AppShell({ employee, onSignOut }) {
  const [tab, setTab] = useState("dashboard");
  const [customerCount, setCustomerCount] = useState(0);
  const [inStockCount, setInStockCount] = useState(0);
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);

  const visibleNav = NAV_ITEMS.filter((n) => {
    if (n.managerOnly && employee.role !== "quan_ly") return false;
    if (n.allowedRoles && !n.allowedRoles.includes(employee.role)) return false;
    return true;
  });

  const renderModule = () => {
    switch (tab) {
      case "dashboard":
        return <DashboardModule employee={employee} customerCount={customerCount} inStockCount={inStockCount} />;
      case "customers":
        return <CustomersModule employee={employee} onCountChange={setCustomerCount} />;
      case "inventory":
        return <InventoryModule employee={employee} onCountChange={setInStockCount} />;
      case "orders":
        return <OrdersModule employee={employee} />;
      case "receipts":
        return <ReceiptsModule employee={employee} />;
      case "reports":
        return ["quan_ly", "ke_toan"].includes(employee.role) ? <ReportsModule employee={employee} /> : null;
      case "audit":
        return ["quan_ly", "ke_toan"].includes(employee.role) ? <AuditLogModule /> : null;
      case "employees":
        return employee.role === "quan_ly" ? <EmployeesModule employee={employee} /> : null;
      default: {
        const item = NAV_ITEMS.find((n) => n.key === tab);
        return <ComingSoonModule title={item?.label} icon={item?.icon || Settings} phase={item?.phase || "giai đoạn sau"} />;
      }
    }
  };

  return (
    <div className="min-h-screen bg-gradient-to-b from-brand-50 to-white lg:flex">
      {/* Desktop sidebar */}
      <aside className="hidden lg:flex flex-col w-60 shrink-0 border-r border-slate-100 bg-white/70 p-4 print:hidden">
        <div className="flex items-center gap-2 px-2 mb-6">
          <div className="w-9 h-9 rounded-xl bg-brand-600 flex items-center justify-center">
            <Smartphone className="text-white" size={18} />
          </div>
          <div>
            <p className="text-sm font-semibold text-slate-800 leading-tight">CH 54 Xuân Thủy</p>
            <p className="text-[11px] text-slate-400">Quản lý mua bán máy</p>
          </div>
        </div>
        <nav className="flex-1 space-y-1">
          {visibleNav.map((item) => (
            <button
              key={item.key}
              onClick={() => setTab(item.key)}
              className={classNames(
                "w-full flex items-center gap-2.5 px-3 py-2 rounded-xl text-sm transition",
                tab === item.key ? "bg-brand-600 text-white" : "text-slate-500 hover:bg-slate-100"
              )}
            >
              <item.icon size={16} />
              <span className="flex-1 text-left">{item.label}</span>
              {item.phase && <span className={classNames("text-[10px]", tab === item.key ? "text-brand-100" : "text-slate-300")}>{item.phase}</span>}
            </button>
          ))}
        </nav>
        <button onClick={onSignOut} className="flex items-center gap-2.5 px-3 py-2 rounded-xl text-sm text-slate-400 hover:bg-rose-50 hover:text-rose-600">
          <LogOut size={16} /> Đăng xuất
        </button>
      </aside>

      {/* Mobile top bar */}
      <div className="lg:hidden flex items-center justify-between px-4 py-3 bg-white border-b border-slate-100 sticky top-0 z-10 print:hidden">
        <div className="flex items-center gap-2">
          <div className="w-8 h-8 rounded-lg bg-brand-600 flex items-center justify-center">
            <Smartphone className="text-white" size={16} />
          </div>
          <p className="text-sm font-semibold text-slate-800">CH 54 Xuân Thủy</p>
        </div>
        <button onClick={() => setMobileMenuOpen((s) => !s)} className="text-slate-500"><Menu size={20} /></button>
      </div>
      {mobileMenuOpen && (
        <div className="lg:hidden bg-white border-b border-slate-100 px-4 py-2">
          {visibleNav.map((item) => (
            <button
              key={item.key}
              onClick={() => { setTab(item.key); setMobileMenuOpen(false); }}
              className={classNames(
                "w-full flex items-center gap-2.5 px-2 py-2.5 rounded-lg text-sm",
                tab === item.key ? "text-brand-700 font-medium" : "text-slate-500"
              )}
            >
              <item.icon size={16} /><span className="flex-1 text-left">{item.label}</span>
              {tab === item.key && <ChevronRight size={14} />}
            </button>
          ))}
          <button onClick={onSignOut} className="w-full flex items-center gap-2.5 px-2 py-2.5 rounded-lg text-sm text-rose-500">
            <LogOut size={16} /> Đăng xuất
          </button>
        </div>
      )}

      {/* Main content */}
      <main className="flex-1 p-4 sm:p-6 pb-24 lg:pb-6 max-w-6xl mx-auto w-full">
        {renderModule()}
      </main>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Root App: session + employee resolution                               */
/* ---------------------------------------------------------------------- */

export default function App() {
  const [session, setSession] = useState(undefined); // undefined = loading, null = signed out
  const [employee, setEmployee] = useState(undefined); // undefined = loading, null = not provisioned

  useEffect(() => {
    supabase.auth.getSession().then(({ data }) => setSession(data.session));
    const { data: sub } = supabase.auth.onAuthStateChange((_event, s) => setSession(s));
    return () => sub.subscription.unsubscribe();
  }, []);

  const resolveEmployee = useCallback(async (s) => {
    if (!s) { setEmployee(null); return; }
    const uid = s.user.id;
    const userEmail = (s.user.email || "").toLowerCase();

    let { data: byUid } = await supabase.from("employees").select("*").eq("user_id", uid).maybeSingle();
    if (byUid) { setEmployee(byUid); return; }

    // First login after being invited: link by email if a placeholder row exists.
    const { data: byEmail } = await supabase.from("employees").select("*").eq("email", userEmail).is("user_id", null).maybeSingle();
    if (byEmail) {
      const { data: linked, error } = await supabase.from("employees").update({ user_id: uid }).eq("id", byEmail.id).select().maybeSingle();
      if (!error && linked) { setEmployee(linked); return; }
    }
    setEmployee(null);
  }, []);

  useEffect(() => {
    if (session === undefined) return;
    resolveEmployee(session);
  }, [session, resolveEmployee]);

  const signOut = async () => {
    await supabase.auth.signOut();
    setEmployee(undefined);
  };

  if (session === undefined || (session && employee === undefined)) {
    return (
      <div className="min-h-screen flex items-center justify-center bg-white">
        <Loader2 className="animate-spin text-brand-400" size={28} />
      </div>
    );
  }

  if (!session) {
    return <LoginPage onLoggedIn={() => {}} />;
  }

  if (employee === null || employee?.is_active === false) {
    return <NotProvisioned email={session.user.email} onSignOut={signOut} />;
  }

  return <AppShell employee={employee} onSignOut={signOut} />;
}
