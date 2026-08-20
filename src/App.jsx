import React, { useEffect, useState, useCallback } from "react";
import { supabase } from "./supabaseClient";
import * as XLSX from "xlsx";
import {
  BarChart as RBarChart, Bar, XAxis, YAxis, CartesianGrid, Tooltip as RTooltip, ResponsiveContainer,
} from "recharts";
import {
  LayoutDashboard, Users, Smartphone, ShoppingCart, Receipt,
  BarChart3, UserCog, ScrollText, Settings, LogOut, Search, Plus, X,
  Loader2, ChevronRight, Menu, ShieldAlert, Pencil, Trash2, PackageSearch,
  History, Printer, Wallet, Landmark, CalendarClock, ChevronDown, FileSpreadsheet,
  Filter, TrendingUp, Package, Award, ArrowLeftRight, Banknote,
  Building2,
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

const DEVICE_ORIGIN_LABELS = {
  supplier: "Nhập NCC",
  customer: "Thu khách lẻ",
  returned: "Khách trả lại",
  internal: "Nhận nội bộ",
  unknown: "Không rõ nguồn",
};
const DEVICE_ORIGIN_STYLES = {
  supplier: "bg-sky-50 text-sky-700 ring-1 ring-sky-200",
  customer: "bg-amber-50 text-amber-700 ring-1 ring-amber-200",
  returned: "bg-rose-50 text-rose-700 ring-1 ring-rose-200",
  internal: "bg-violet-50 text-violet-700 ring-1 ring-violet-200",
  unknown: "bg-slate-100 text-slate-400 ring-1 ring-slate-200",
};
// Vạch màu bên trái mỗi dòng trong Kho hàng
const DEVICE_ORIGIN_ROW = {
  supplier: "",
  customer: "shadow-[inset_3px_0_0_0_#f59e0b]",
  returned: "shadow-[inset_3px_0_0_0_#e11d48] bg-rose-50/30",
  internal: "shadow-[inset_3px_0_0_0_#7c3aed]",
  unknown: "shadow-[inset_3px_0_0_0_#cbd5e1]",
};

const DEVICE_STATUS_LABELS = {
  in_stock: "Còn hàng",
  reserved: "Đang giữ chỗ",
  sold: "Đã bán",
  pending_reconciliation: "Chờ đối soát",
};
const DEVICE_STATUS_STYLES = {
  in_stock: "bg-emerald-50 text-emerald-700",
  reserved: "bg-amber-50 text-amber-700",
  sold: "bg-slate-100 text-slate-500",
  pending_reconciliation: "bg-rose-50 text-rose-600",
};
const DEVICE_CONDITION_LABELS = {
  new: "Máy mới",
  used: "Máy cũ",
};

// Danh mục iPhone từ đời X (2017) trở lên + màu sắc chính hãng từng đời (cập
// nhật tới iPhone 17 series 2025). Dùng làm gợi ý nhập liệu (datalist), vẫn
// cho gõ tay tự do với các dòng máy/hãng khác không có trong danh sách.
const IPHONE_COLOR_CATALOG = {
  "iPhone X": ["Space Gray", "Silver"],
  "iPhone XR": ["White", "Black", "Blue", "Yellow", "Coral", "(PRODUCT)RED"],
  "iPhone XS": ["Silver", "Space Gray", "Gold"],
  "iPhone XS Max": ["Silver", "Space Gray", "Gold"],
  "iPhone 11": ["White", "Black", "Green", "Yellow", "Purple", "(PRODUCT)RED"],
  "iPhone 11 Pro": ["Midnight Green", "Space Gray", "Silver", "Gold"],
  "iPhone 11 Pro Max": ["Midnight Green", "Space Gray", "Silver", "Gold"],
  "iPhone SE (2020)": ["White", "Black", "(PRODUCT)RED"],
  "iPhone 12 mini": ["Black", "White", "Blue", "Green", "Purple", "(PRODUCT)RED"],
  "iPhone 12": ["Black", "White", "Blue", "Green", "Purple", "(PRODUCT)RED"],
  "iPhone 12 Pro": ["Graphite", "Silver", "Gold", "Pacific Blue"],
  "iPhone 12 Pro Max": ["Graphite", "Silver", "Gold", "Pacific Blue"],
  "iPhone 13 mini": ["Pink", "Blue", "Midnight", "Starlight", "(PRODUCT)RED", "Green"],
  "iPhone 13": ["Pink", "Blue", "Midnight", "Starlight", "(PRODUCT)RED", "Green", "Alpine Green"],
  "iPhone 13 Pro": ["Graphite", "Gold", "Silver", "Sierra Blue", "Alpine Green"],
  "iPhone 13 Pro Max": ["Graphite", "Gold", "Silver", "Sierra Blue", "Alpine Green"],
  "iPhone SE (2022)": ["Midnight", "Starlight", "(PRODUCT)RED"],
  "iPhone 14": ["Blue", "Purple", "Midnight", "Starlight", "(PRODUCT)RED", "Yellow"],
  "iPhone 14 Plus": ["Blue", "Purple", "Midnight", "Starlight", "(PRODUCT)RED", "Yellow"],
  "iPhone 14 Pro": ["Space Black", "Silver", "Gold", "Deep Purple"],
  "iPhone 14 Pro Max": ["Space Black", "Silver", "Gold", "Deep Purple"],
  "iPhone 15": ["Pink", "Yellow", "Green", "Blue", "Black"],
  "iPhone 15 Plus": ["Pink", "Yellow", "Green", "Blue", "Black"],
  "iPhone 15 Pro": ["Black Titanium", "White Titanium", "Blue Titanium", "Natural Titanium"],
  "iPhone 15 Pro Max": ["Black Titanium", "White Titanium", "Blue Titanium", "Natural Titanium"],
  "iPhone 16": ["Black", "White", "Pink", "Teal", "Ultramarine"],
  "iPhone 16 Plus": ["Black", "White", "Pink", "Teal", "Ultramarine"],
  "iPhone 16 Pro": ["Black Titanium", "White Titanium", "Natural Titanium", "Desert Titanium"],
  "iPhone 16 Pro Max": ["Black Titanium", "White Titanium", "Natural Titanium", "Desert Titanium"],
  "iPhone 16e": ["Black", "White"],
  "iPhone Air": ["Space Black", "Cloud White", "Light Gold", "Sky Blue"],
  "iPhone 17e": ["Black", "White"],
  "iPhone 17": ["Black", "White", "Lavender", "Mist Blue", "Sage"],
  "iPhone 17 Pro": ["Silver", "Deep Blue", "Cosmic Orange"],
  "iPhone 17 Pro Max": ["Silver", "Deep Blue", "Cosmic Orange"],
};
const IPHONE_MODEL_LIST = Object.keys(IPHONE_COLOR_CATALOG);
const GENERIC_COLOR_LIST = ["Đen", "Trắng", "Bạc", "Vàng", "Xanh", "Xanh lá", "Tím", "Hồng", "Đỏ", "Vàng đồng"];

function coloroptionsForModel(modelText) {
  const trimmed = (modelText || "").trim();
  const exact = IPHONE_COLOR_CATALOG[trimmed];
  if (exact) return exact;
  const key = Object.keys(IPHONE_COLOR_CATALOG).find(
    (m) => trimmed.toLowerCase().startsWith(m.toLowerCase())
  );
  return key ? IPHONE_COLOR_CATALOG[key] : GENERIC_COLOR_LIST;
}

const PAYMENT_METHOD_LABELS = {
  cash: "Tiền mặt",
  bank_transfer: "Chuyển khoản",
  installment: "Trả góp",
  trade_in: "Đổi máy cũ",
  debt_offset: "Bù trừ công nợ",
  debt: "Khách nợ shop",
};
const PAYMENT_METHOD_ICONS = {
  cash: Wallet,
  bank_transfer: Landmark,
  installment: CalendarClock,
  trade_in: ArrowLeftRight,
  debt_offset: Banknote,
  debt: CalendarClock,
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
    <div className={classNames("bg-white rounded-2xl shadow-card border border-slate-200/60", className)}>
      {children}
    </div>
  );
}

function TextField({ label, className = "", ...props }) {
  return (
    <label className={classNames("block", className)}>
      {label && <span className="text-xs font-medium text-slate-600 mb-1.5 block">{label}</span>}
      <input
        {...props}
        className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm text-slate-800 placeholder:text-slate-400 outline-none focus:ring-2 focus:ring-brand-400/40 focus:border-brand-500 transition"
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
    <div className="min-h-screen bg-slate-50 flex items-center justify-center p-4">
      <Card className="w-full max-w-sm p-6 sm:p-8">
        <div className="flex flex-col items-center mb-6">
          <div className="w-14 h-14 rounded-2xl bg-brand-600 flex items-center justify-center mb-3">
            <Smartphone className="text-white" size={26} />
          </div>
          <h1 className="text-lg font-semibold text-slate-800">Quản lý mua bán điện thoại</h1>
          <p className="text-xs text-slate-400 mt-1">Hệ thống dùng chung cho nhiều cửa hàng</p>
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
    <div className="min-h-screen bg-slate-50 flex items-center justify-center p-4">
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
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Customers module                                                      */
/* ---------------------------------------------------------------------- */

function CustomerForm({ initial, onCancel, onSaved, employee, existingMatch, onUseExisting }) {
  const [form, setForm] = useState(() => initial ? {
    full_name: initial.full_name || "", cccd: initial.cccd || "", cccd_issue_date: initial.cccd_issue_date || "",
    cccd_issue_place: initial.cccd_issue_place || "", date_of_birth: initial.date_of_birth || "",
    address: initial.address || "", phone: initial.phone || "", email: initial.email || "",
  } : {
    full_name: "", cccd: "", cccd_issue_date: "", cccd_issue_place: "", date_of_birth: "", address: "", phone: "", email: "",
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
        date_of_birth: form.date_of_birth || null,
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
        payload.store_id = employee.store_id;
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
        <TextField label="Ngày sinh" type="date" value={form.date_of_birth || ""} onChange={set("date_of_birth")} />
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
              className="w-full pl-9 pr-3 py-2 rounded-lg border border-slate-300 bg-white text-sm placeholder:text-slate-400 outline-none focus:ring-2 focus:ring-brand-400/40 focus:border-brand-500 transition"
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
  const [form, setForm] = useState(() => initial ? {
    imei: initial.imei || "", model: initial.model || "", storage: initial.storage || "", color: initial.color || "",
    condition: initial.condition || "used", condition_percent: initial.condition_percent ?? "",
    cost_price: initial.cost_price ?? "", sale_price: initial.sale_price ?? "",
    supplier: initial.supplier || "", import_date: initial.import_date || "", notes: initial.notes || "",
  } : {
    imei: "", model: "", storage: "", color: "", condition: "used", condition_percent: "",
    cost_price: "", sale_price: "", supplier: "", import_date: "", notes: "",
  });
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const isSold = initial?.status === "sold";
  const canManagerEditSold = isSold && employee.role === "quan_ly";
  const lockedForStaff = isSold && employee.role !== "quan_ly";

  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));

  const submit = async (e) => {
    e.preventDefault();
    if (lockedForStaff) { setError("Máy đã bán — chỉ Quản lý mới có quyền chỉnh sửa."); return; }
    if (!form.model.trim()) { setError("Vui lòng nhập tên model máy."); return; }
    if (form.imei.trim() && duplicateImei) {
      setError(`IMEI ${duplicateImei.imei} đã tồn tại trong kho — không được nhập trùng.`);
      return;
    }
    setError(""); setSaving(true);
    try {
      const payload = isSold ? {
        // Máy đã bán: chỉ cho phép sửa IMEI + Nhà cung cấp/nguồn nhập, mọi trường
        // khác giữ nguyên giá trị gốc để bảo vệ dữ liệu đã chốt.
        imei: form.imei.trim() || null,
        supplier: form.supplier.trim() || null,
        updated_by: employee.id,
      } : {
        imei: form.imei.trim() || null,
        model: form.model.trim(),
        storage: form.storage.trim() || null,
        color: form.color.trim() || null,
        condition: form.condition,
        condition_percent: form.condition_percent === "" ? null : Number(form.condition_percent),
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
          old_data: initial, new_data: updated, performed_by: employee.id, store_id: employee.store_id,
        });
      } else {
        payload.status = "in_stock";
        payload.created_by = employee.id;
        payload.store_id = employee.store_id;
        const { data: created, error: err } = await supabase.from("devices").insert(payload).select().maybeSingle();
        if (err) throw err;
        await supabase.from("audit_logs").insert({
          table_name: "devices", record_id: created?.id, action: "create",
          new_data: created, performed_by: employee.id, store_id: employee.store_id,
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
        <div className="bg-rose-50 border border-rose-200 rounded-xl px-3 py-2.5 mb-3 text-xs text-rose-700 flex items-center gap-2">
          <ShieldAlert size={15} className="shrink-0" />
          IMEI <span className="font-medium">{duplicateImei.imei}</span> đã tồn tại trong kho ({DEVICE_STATUS_LABELS[duplicateImei.status]}) — không thể lưu trùng, vui lòng kiểm tra lại.
        </div>
      )}

      {isSold && lockedForStaff && (
        <div className="bg-slate-100 border border-slate-200 rounded-xl px-3 py-2.5 mb-3 text-xs text-slate-600 flex items-center gap-2">
          <ShieldAlert size={15} className="shrink-0" />
          Máy này đã bán — chỉ Quản lý cửa hàng mới có quyền chỉnh sửa. Toàn bộ thông tin bên dưới chỉ để xem.
        </div>
      )}
      {isSold && canManagerEditSold && (
        <div className="bg-amber-50 border border-amber-200 rounded-xl px-3 py-2.5 mb-3 text-xs text-amber-800 flex items-center gap-2">
          <ShieldAlert size={15} className="shrink-0" />
          Máy này đã bán — chỉ sửa được IMEI và Nhà cung cấp/nguồn nhập (khắc phục sai sót nhập liệu), các thông tin khác được khóa để bảo vệ dữ liệu đã chốt.
        </div>
      )}

      <form onSubmit={submit} className="grid sm:grid-cols-2 gap-3">
        <TextField label="Số IMEI (để trống nếu chưa có)" value={form.imei} onChange={set("imei")} disabled={lockedForStaff || (!!initial?.id && !isSold)} />
        <label className="block">
          <span className="text-xs font-medium text-slate-500 mb-1 block">Model máy *</span>
          <ModelPicker value={form.model} onSelect={(v) => setForm((f) => ({ ...f, model: v }))} placeholder="iPhone 14 Pro Max" disabled={isSold} />
        </label>
        <TextField label="Dung lượng" value={form.storage} onChange={set("storage")} placeholder="256GB" list="dl-storage" disabled={isSold} />
        <TextField label="Màu sắc" value={form.color} onChange={set("color")} placeholder="Tím" list="dl-colors-device" disabled={isSold} />
        <datalist id="dl-colors-device">{coloroptionsForModel(form.model).map((c) => <option key={c} value={c} />)}</datalist>
        <datalist id="dl-storage"><option value="64GB" /><option value="128GB" /><option value="256GB" /><option value="512GB" /><option value="1TB" /></datalist>
        <label className="block">
          <span className="text-xs font-medium text-slate-500 mb-1 block">Tình trạng</span>
          <select value={form.condition} onChange={set("condition")} disabled={isSold} className="w-full rounded-xl border border-slate-200 px-3 py-2 text-sm disabled:bg-slate-50 disabled:text-slate-400">
            <option value="new">Máy mới</option>
            <option value="used">Máy cũ</option>
          </select>
        </label>
        <TextField label="Độ mới (%)" type="number" min="0" max="100" value={form.condition_percent} onChange={set("condition_percent")} placeholder="99" disabled={isSold} />
        <TextField label="Nhà cung cấp / nguồn nhập" value={form.supplier} onChange={set("supplier")} disabled={lockedForStaff} />
        <TextField label="Giá vốn (đ)" type="number" value={form.cost_price} onChange={set("cost_price")} disabled={isSold} />
        <TextField label="Giá bán đề xuất (đ)" type="number" value={form.sale_price} onChange={set("sale_price")} disabled={isSold} />
        <TextField label="Ngày nhập" type="date" value={form.import_date || ""} onChange={set("import_date")} disabled={isSold} />
        <TextField label="Ghi chú" value={form.notes} onChange={set("notes")} className="sm:col-span-2" disabled={isSold} />
        {error && <p className="text-xs text-rose-600 sm:col-span-2">{error}</p>}
        <div className="sm:col-span-2 flex gap-2 mt-1">
          {!lockedForStaff && (
            <button type="submit" disabled={saving} className="bg-brand-600 hover:bg-brand-700 text-white rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-2 disabled:opacity-60">
              {saving && <Loader2 size={14} className="animate-spin" />} Lưu
            </button>
          )}
          <button type="button" onClick={onCancel} className="text-slate-500 text-sm px-4 py-2">{lockedForStaff ? "Đóng" : "Hủy"}</button>
        </div>
      </form>
    </Card>
  );
}

const HISTORY_STYLES = {
  purchase:     { label: "Nhập kho",     dot: "bg-sky-500",     chip: "bg-sky-50 text-sky-700" },
  spa:          { label: "Spa",          dot: "bg-fuchsia-500", chip: "bg-fuchsia-50 text-fuchsia-700" },
  screen:       { label: "Thay màn",     dot: "bg-violet-500",  chip: "bg-violet-50 text-violet-700" },
  transfer_out: { label: "Xuất nội bộ",  dot: "bg-indigo-500",  chip: "bg-indigo-50 text-indigo-700" },
  sale:         { label: "Bán",          dot: "bg-emerald-500", chip: "bg-emerald-50 text-emerald-700" },
  return:       { label: "Khách trả",    dot: "bg-rose-500",    chip: "bg-rose-50 text-rose-700" },
};

function DeviceHistoryPanel({ device, employee, onClose }) {
  const [rows, setRows] = useState(null);
  const [loading, setLoading] = useState(true);
  const canSeeCost = employee.role !== "nhan_vien";

  useEffect(() => {
    let active = true;
    setLoading(true);
    supabase.rpc("device_history", { p_device_id: device.id })
      .then(({ data }) => { if (active) { setRows(data || []); setLoading(false); } });
    return () => { active = false; };
  }, [device.id]);

  const services = (rows || []).filter((r) => r.event_type === "spa" || r.event_type === "screen");
  const spaTotal = services.filter((r) => r.event_type === "spa")
    .reduce((s, r) => s + Number(r.amount || 0), 0);
  const screenNet = services.filter((r) => r.event_type === "screen")
    .reduce((s, r) => s + Number(r.amount || 0), 0);

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/40 flex items-center justify-center p-4 overflow-y-auto">
      <Card className="w-full max-w-2xl p-5 my-8">
        <div className="flex items-center justify-between mb-1">
          <p className="font-semibold text-slate-800 text-sm">Lý lịch máy</p>
          <button onClick={onClose} className="text-slate-400 hover:text-rose-600"><X size={16} /></button>
        </div>
        <p className="text-xs text-slate-400 mb-4">
          {[device.model, device.storage, device.color].filter(Boolean).join(" ")} · IMEI {device.imei || "—"}
        </p>

        {services.length > 0 && canSeeCost && (
          <div className="grid grid-cols-3 gap-3 mb-4">
            <div className="bg-slate-50 rounded-xl p-3">
              <p className="text-xs text-slate-400 mb-1">Số lần dịch vụ</p>
              <p className="text-base font-semibold text-slate-800">{services.length}</p>
            </div>
            <div className="bg-fuchsia-50 rounded-xl p-3">
              <p className="text-xs text-fuchsia-600 mb-1">Chi phí spa</p>
              <p className="text-base font-semibold text-fuchsia-700">{fmtVND(spaTotal)}</p>
            </div>
            <div className={classNames("rounded-xl p-3", screenNet >= 0 ? "bg-amber-50" : "bg-emerald-50")}>
              <p className={classNames("text-xs mb-1", screenNet >= 0 ? "text-amber-600" : "text-emerald-600")}>Chênh do thay màn</p>
              <p className={classNames("text-base font-semibold", screenNet >= 0 ? "text-amber-700" : "text-emerald-700")}>
                {screenNet >= 0 ? "+" : ""}{fmtVND(screenNet)}
              </p>
            </div>
          </div>
        )}

        {loading ? (
          <div className="flex justify-center py-8"><Loader2 className="animate-spin text-slate-300" /></div>
        ) : !rows || rows.length === 0 ? (
          <p className="text-xs text-slate-400 py-6 text-center">Chưa có sự kiện nào cho máy này.</p>
        ) : (
          <ol className="relative border-l-2 border-slate-100 ml-2 space-y-4">
            {rows.map((r, i) => {
              const st = HISTORY_STYLES[r.event_type] || { label: r.event_type, dot: "bg-slate-400", chip: "bg-slate-100 text-slate-500" };
              const before = r.cost_before == null ? null : Number(r.cost_before);
              const after  = r.cost_after  == null ? null : Number(r.cost_after);
              const changed = before != null && after != null && Math.round(before) !== Math.round(after);
              return (
                <li key={i} className="ml-4">
                  <span className={classNames("absolute -left-[7px] w-3 h-3 rounded-full ring-2 ring-white", st.dot)} />
                  <div className="flex items-center gap-2 flex-wrap">
                    <span className={classNames("text-[10px] px-1.5 py-0.5 rounded", st.chip)}>{st.label}</span>
                    <span className="text-sm font-medium text-slate-700">{r.title}</span>
                    {r.code && <span className="text-xs text-slate-400">{r.code}</span>}
                    <span className="text-xs text-slate-400">· {fmtDate(r.event_date)}</span>
                  </div>
                  {r.detail && <p className="text-xs text-slate-500 mt-0.5">{r.detail}</p>}
                  {canSeeCost && changed && (
                    <p className="text-xs mt-0.5">
                      <span className="text-slate-400">Giá vốn </span>
                      <span className="text-slate-500">{fmtVND(before)}</span>
                      <span className="text-slate-400"> → </span>
                      <span className={classNames("font-medium", after > before ? "text-amber-600" : "text-emerald-600")}>
                        {fmtVND(after)}
                      </span>
                      <span className={classNames("ml-1", after > before ? "text-amber-600" : "text-emerald-600")}>
                        ({after > before ? "+" : ""}{fmtVND(after - before)})
                      </span>
                    </p>
                  )}
                </li>
              );
            })}
          </ol>
        )}
      </Card>
    </div>
  );
}

function ScreenHistoryPanel({ screen, onClose }) {
  const [rows, setRows] = useState(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    let active = true;
    setLoading(true);
    supabase.rpc("screen_history", { p_screen_id: screen.id })
      .then(({ data }) => { if (active) { setRows(data || []); setLoading(false); } });
    return () => { active = false; };
  }, [screen.id]);

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/40 flex items-center justify-center p-4">
      <Card className="w-full max-w-md p-5">
        <div className="flex items-center justify-between mb-1">
          <p className="font-semibold text-slate-800 text-sm">Lý lịch màn {screen.screen_code}</p>
          <button onClick={onClose} className="text-slate-400 hover:text-rose-600"><X size={16} /></button>
        </div>
        <p className="text-xs text-slate-400 mb-4">
          {screen.model || "—"} · {screen.grade || "chưa đánh giá"} · {fmtVND(screen.unit_price)}
        </p>
        {loading ? (
          <div className="flex justify-center py-6"><Loader2 className="animate-spin text-slate-300" /></div>
        ) : !rows || rows.length === 0 ? (
          <p className="text-xs text-slate-400 py-4 text-center">Chưa có sự kiện nào.</p>
        ) : (
          <ol className="relative border-l-2 border-slate-100 ml-2 space-y-3">
            {rows.map((r, i) => (
              <li key={i} className="ml-4">
                <span className={classNames("absolute -left-[7px] w-3 h-3 rounded-full ring-2 ring-white",
                  r.event_type === "installed_to" ? "bg-sky-500" : "bg-slate-400")} />
                <p className="text-sm font-medium text-slate-700">{r.title}</p>
                <p className="text-xs text-slate-500">{r.detail}</p>
                <p className="text-xs text-slate-400">{fmtDate(r.event_date)}{r.code ? ` · ${r.code}` : ""}</p>
              </li>
            ))}
          </ol>
        )}
      </Card>
    </div>
  );
}

function InventoryModule({ employee, onCountChange }) {
  const [devices, setDevices] = useState([]);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState("");
  const [statusFilter, setStatusFilter] = useState("all");
  const [originFilter, setOriginFilter] = useState("all");
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [historyDevice, setHistoryDevice] = useState(null);
  const [incoming, setIncoming] = useState([]);
  const [receiving, setReceiving] = useState(null);
  const canReceive = employee.role !== "nhan_vien";

  const loadIncoming = useCallback(async () => {
    const { data } = await supabase.from("v_internal_transfers").select("*")
      .eq("status", "pending").eq("to_store_id", employee.store_id)
      .order("created_at", { ascending: true });
    setIncoming(data || []);
  }, [employee.store_id]);

  useEffect(() => { loadIncoming(); }, [loadIncoming]);

  const canDelete = employee.role === "quan_ly";
  const canManage = employee.role === "quan_ly";
  const canSeeCost = employee.role !== "nhan_vien";

  const load = useCallback(async () => {
    setLoading(true);
    const { data, error } = await supabase.from("v_devices_origin").select("*").order("created_at", { ascending: false }).limit(2000);
    if (!error) {
      setDevices(data || []);
      onCountChange?.((data || []).filter((d) => d.status === "in_stock").length);
    }
    setLoading(false);
  }, [onCountChange]);

  useEffect(() => { load(); }, [load]);

  const filtered = devices.filter((d) => {
    if (originFilter !== "all" && d.origin !== originFilter) return false;
    if (statusFilter === "missing_imei") return !d.imei;
    if (statusFilter !== "all" && d.status !== statusFilter) return false;
    if (!search.trim()) return true;
    const q = search.trim().toLowerCase();
    return (
      d.imei?.toLowerCase().includes(q) ||
      d.model?.toLowerCase().includes(q) ||
      d.color?.toLowerCase().includes(q)
    );
  });

  const openEdit = (d) => { setEditing(d); setShowForm(true); };

  const remove = async (d) => {
    if (!confirm(`Xóa máy ${d.imei ? `IMEI "${d.imei}"` : `"${d.model}" (chưa có IMEI)`} khỏi kho?`)) return;
    const { error } = await supabase.from("devices").delete().eq("id", d.id);
    if (error) { alert(error.message); return; }
    await supabase.from("audit_logs").insert({
      table_name: "devices", record_id: d.id, action: "delete",
      old_data: d, performed_by: employee.id, store_id: employee.store_id,
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
      </div>

      {incoming.length > 0 && (
        <Card className="p-4 mb-4 border-violet-200 bg-violet-50/40">
          <div className="flex items-center gap-2 mb-3">
            <ArrowLeftRight size={16} className="text-violet-600 shrink-0" />
            <p className="text-sm font-medium text-violet-800">
              {incoming.length} máy chờ nhận từ cửa hàng khác
            </p>
          </div>
          <div className="space-y-2">
            {incoming.map((t) => (
              <div key={t.id} className="flex items-center gap-3 bg-white border border-violet-200 rounded-xl px-3 py-2.5 flex-wrap">
                <div className="flex-1 min-w-0">
                  <p className="text-sm font-medium text-slate-700">
                    {[t.model, t.storage, t.color].filter(Boolean).join(" ")}
                  </p>
                  <p className="text-xs text-slate-400">
                    IMEI {t.imei || "—"} · {t.transfer_code} · từ <span className="font-medium">{t.from_store_name}</span> · {fmtDate(t.created_at)}
                  </p>
                  {t.note && <p className="text-xs text-slate-400">Ghi chú: {t.note}</p>}
                </div>
                <div className="text-right shrink-0">
                  <p className="text-sm font-semibold text-slate-700">{fmtVND(t.transfer_price)}</p>
                  <p className="text-[11px] text-slate-400">giá vốn khi nhận</p>
                </div>
                {canReceive ? (
                  <button onClick={() => setReceiving(t)}
                    className="bg-violet-600 hover:bg-violet-700 text-white rounded-lg px-3 py-1.5 text-xs font-medium whitespace-nowrap">
                    Nhận hàng
                  </button>
                ) : (
                  <span className="text-[11px] text-slate-400 whitespace-nowrap">Chờ kế toán/quản lý nhận</span>
                )}
              </div>
            ))}
          </div>
        </Card>
      )}

      <div className="flex flex-wrap items-center gap-3 mb-3 text-xs">
        <span className="text-slate-400">Nguồn gốc:</span>
        {["supplier", "customer", "returned", "internal", "unknown"].map((k) => {
          const n = devices.filter((d) => d.origin === k && d.status === "in_stock").length;
          return (
            <button
              key={k}
              onClick={() => setOriginFilter(originFilter === k ? "all" : k)}
              className={classNames(
                "px-2 py-1 rounded transition",
                DEVICE_ORIGIN_STYLES[k],
                originFilter === k ? "ring-2" : "opacity-90 hover:opacity-100"
              )}
            >
              {DEVICE_ORIGIN_LABELS[k]} · {n}
            </button>
          );
        })}
      </div>

      <div className="bg-slate-100 border border-slate-200 rounded-xl px-3 py-2.5 mb-4 text-xs text-slate-600 flex items-start gap-2">
        <ArrowLeftRight size={15} className="shrink-0 mt-0.5" />
        <span>
          Máy chỉ vào kho qua <span className="font-medium">Nhập máy/Thu cũ</span> hoặc khi khách đổi máy trong đơn bán —
          để mỗi chiếc luôn có phiếu nhập, biết mua của ai và ghi đúng công nợ.
          {!canManage && " Bạn chỉ có quyền xem Kho hàng."}
        </span>
      </div>

      {canManage && devices.some((d) => !d.imei) && (
        <div className="bg-amber-50 border border-amber-200 rounded-xl px-3 py-2.5 mb-4 text-xs text-amber-800 flex items-center gap-2">
          <ShieldAlert size={15} className="shrink-0" />
          Có <span className="font-medium">{devices.filter((d) => !d.imei).length} máy</span> chưa có số IMEI — vui lòng cập nhật khi có đủ thông tin (lọc theo "Thiếu IMEI" bên dưới).
        </div>
      )}

      {receiving && (
        <ReceiveTransferModal
          row={receiving} employee={employee}
          onClose={() => setReceiving(null)}
          onDone={() => { setReceiving(null); loadIncoming(); load(); }}
        />
      )}

      {canManage && showForm && editing && (
        <DeviceForm
          initial={editing}
          employee={employee}
          duplicateImei={null}
          onCancel={() => { setShowForm(false); setEditing(null); }}
          onSaved={() => { setShowForm(false); setEditing(null); load(); }}
        />
      )}

      {historyDevice && (
        <DeviceHistoryPanel device={historyDevice} employee={employee} onClose={() => setHistoryDevice(null)} />
      )}

      <Card className="p-0 overflow-hidden">
        <div className="p-3 border-b border-slate-100 flex flex-col sm:flex-row gap-2">
          <div className="relative flex-1">
            <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-300" />
            <input
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="Tìm theo IMEI, model, màu..."
              className="w-full pl-9 pr-3 py-2 rounded-lg border border-slate-300 bg-white text-sm placeholder:text-slate-400 outline-none focus:ring-2 focus:ring-brand-400/40 focus:border-brand-500 transition"
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
            <option value="missing_imei">⚠ Thiếu IMEI</option>
          </select>
          <select
            value={originFilter}
            onChange={(e) => setOriginFilter(e.target.value)}
            className="rounded-xl border border-slate-200 px-3 py-2 text-sm"
          >
            <option value="all">Mọi nguồn gốc</option>
            <option value="supplier">Nhập NCC</option>
            <option value="customer">Thu khách lẻ</option>
            <option value="returned">Khách trả lại</option>
            <option value="internal">Nhận nội bộ</option>
            <option value="unknown">Không rõ nguồn</option>
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
                  <tr key={d.id} className={classNames(
                    "border-b border-slate-50 last:border-0 hover:bg-slate-50/60",
                    DEVICE_ORIGIN_ROW[d.origin] || ""
                  )}>
                    <td className="px-3 py-2.5 font-medium text-slate-700 whitespace-nowrap">
                      {d.imei || (
                        <span className="inline-flex items-center gap-1 text-xs font-normal px-2 py-0.5 rounded-full bg-amber-50 text-amber-700">
                          <ShieldAlert size={11} /> Thiếu IMEI
                        </span>
                      )}
                    </td>
                    <td className="px-3 py-2.5 text-slate-600">
                      <div>
                        {d.model}
                        <span className="text-slate-400"> {[d.storage, d.color].filter(Boolean).join(" · ")}</span>
                      </div>
                      <span
                        title={d.origin === "returned"
                          ? `Khách trả lại — chứng từ ${d.return_code || ""} ngày ${d.return_date ? fmtDate(d.return_date) : ""}`
                          : d.origin === "internal"
                          ? `Nhận nội bộ theo phiếu ${d.transfer_code || ""}`
                          : d.purchase_code ? `Phiếu nhập ${d.purchase_code}` : "Không có phiếu nhập"}
                        className={classNames(
                          "inline-block mt-1 text-[10px] px-1.5 py-0.5 rounded",
                          DEVICE_ORIGIN_STYLES[d.origin] || DEVICE_ORIGIN_STYLES.unknown
                        )}
                      >
                        {DEVICE_ORIGIN_LABELS[d.origin] || "Không rõ nguồn"}
                      </span>
                    </td>
                    <td className="px-3 py-2.5 text-slate-500">
                      {DEVICE_CONDITION_LABELS[d.condition] || "—"}
                      {d.condition_percent != null && <span className="text-slate-400"> · {d.condition_percent}%</span>}
                    </td>
                    {canSeeCost && <td className="px-3 py-2.5 text-slate-500">{fmtVND(d.cost_price)}</td>}
                    <td className="px-3 py-2.5 text-slate-500">{fmtVND(d.sale_price)}</td>
                    <td className="px-3 py-2.5">
                      <span
                        className={classNames("text-xs px-2 py-0.5 rounded-full", DEVICE_STATUS_STYLES[d.status])}
                        title="Trạng thái do nghiệp vụ quyết định: bán hàng, xuất nội bộ, gửi spa/sửa, trả hàng"
                      >
                        {DEVICE_STATUS_LABELS[d.status] || d.status}
                      </span>
                    </td>
                    <td className="px-3 py-2.5 text-right whitespace-nowrap">
                      <button onClick={() => setHistoryDevice(d)} className="text-slate-400 hover:text-brand-600 mr-3" title="Lý lịch máy">
                        <History size={14} className="inline" />
                      </button>
                      {canManage && (d.status !== "sold" || employee.role === "quan_ly") && (
                        <button onClick={() => openEdit(d)} className="text-brand-600 hover:underline text-xs mr-3">
                          <Pencil size={12} className="inline mr-0.5" />Sửa
                        </button>
                      )}
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

function PrintDocModal({ type, order, customer, device, payments, contract, storeName, onClose }) {
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
          <p className="text-xs text-slate-400">{storeName || "Cửa hàng"} — Quản lý mua bán điện thoại</p>
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
          <p><span className="text-slate-400">IMEI:</span> {device?.imei || "Chưa có IMEI"}</p>
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
          className="w-full pl-9 pr-3 py-2 rounded-lg border border-slate-300 bg-white text-sm placeholder:text-slate-400 outline-none focus:ring-2 focus:ring-brand-400/40 focus:border-brand-500 transition"
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
          <p className="text-xs text-slate-400">{value.imei ? `IMEI ${value.imei}` : "⚠ Chưa có IMEI"}</p>
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
          className="w-full pl-9 pr-3 py-2 rounded-lg border border-slate-300 bg-white text-sm placeholder:text-slate-400 outline-none focus:ring-2 focus:ring-brand-400/40 focus:border-brand-500 transition"
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
                <p className="text-xs text-slate-400">{d.imei ? `IMEI ${d.imei}` : "⚠ Chưa có IMEI"} · Giá đề xuất {fmtVND(d.sale_price)}</p>
              </button>
            ))
          )}
        </div>
      )}
    </div>
  );
}

function ModelPicker({ value, onSelect, placeholder, disabled }) {
  const [query, setQuery] = useState(value || "");
  const [open, setOpen] = useState(false);

  useEffect(() => { setQuery(value || ""); }, [value]);

  const filtered = query.trim()
    ? IPHONE_MODEL_LIST.filter((m) => m.toLowerCase().includes(query.trim().toLowerCase()))
    : IPHONE_MODEL_LIST;

  const choose = (m) => { onSelect(m); setQuery(m); setOpen(false); };

  return (
    <div className="relative">
      <div className="relative">
        <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-300" />
        <input
          value={query}
          onChange={(e) => { setQuery(e.target.value); onSelect(e.target.value); setOpen(true); }}
          onFocus={() => !disabled && setOpen(true)}
          placeholder={placeholder || "Gõ hoặc chọn từ danh sách..."}
          disabled={disabled}
          className="w-full pl-9 pr-8 py-2 rounded-xl border border-slate-200 text-sm outline-none focus:ring-2 focus:ring-brand-400 disabled:bg-slate-50 disabled:text-slate-400"
        />
        <ChevronDown size={15} className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-300 pointer-events-none" />
      </div>
      {open && !disabled && (
        <div className="absolute z-10 w-full bg-white border border-slate-100 rounded-xl shadow-lg mt-1 max-h-60 overflow-y-auto">
          {filtered.length === 0 ? (
            <div className="p-3 text-xs text-slate-400">Không có trong danh mục — vẫn dùng được đúng tên bạn vừa gõ (máy hãng khác/đời cũ).</div>
          ) : (
            filtered.map((m) => (
              <button
                type="button"
                key={m}
                onClick={() => choose(m)}
                className="w-full text-left px-3 py-2 hover:bg-slate-50 text-sm border-b border-slate-50 last:border-0"
              >
                {m}
              </button>
            ))
          )}
          <button type="button" onClick={() => setOpen(false)} className="w-full text-center px-3 py-1.5 text-xs text-slate-400 hover:bg-slate-50 border-t border-slate-100">
            Đóng danh sách
          </button>
        </div>
      )}
    </div>
  );
}

function usePaymentOptions() {
  const [banks, setBanks] = useState([]);
  const [providers, setProviders] = useState([]);

  useEffect(() => {
    let active = true;
    (async () => {
      const [{ data: b }, { data: p }] = await Promise.all([
        supabase.from("bank_accounts").select("id, short_label, bank_name, account_number")
          .eq("is_active", true).order("sort_order"),
        supabase.from("installment_providers").select("id, code, name")
          .eq("is_active", true).order("sort_order"),
      ]);
      if (!active) return;
      setBanks(b || []);
      setProviders(p || []);
    })();
    return () => { active = false; };
  }, []);

  return { banks, providers };
}

function BankSelect({ banks, value, onChange, className = "" }) {
  return (
    <select
      value={value || ""} onChange={(e) => onChange(e.target.value || null)}
      className={classNames("rounded-lg border border-slate-200 px-2 py-1.5 text-xs", className)}
    >
      <option value="">— Chọn tài khoản nhận tiền —</option>
      {banks.map((b) => (
        <option key={b.id} value={b.id}>{(b.short_label || b.bank_name) + (b.account_number ? ` · ${b.account_number}` : "")}</option>
      ))}
    </select>
  );
}

function PaymentRows({ rows, setRows, total, supplierDebt = 0, supplierName = "" }) {
  const { banks, providers } = usePaymentOptions();
  const paid = rows.reduce((s, r) => s + (Number(r.amount) || 0), 0);
  const remaining = total - paid;
  const debtDeclared = rows.filter((r) => r.method === "debt")
    .reduce((s, r) => s + (Number(r.amount) || 0), 0);
  const offsetUsed = rows.filter((r) => r.method === "debt_offset")
    .reduce((s, r) => s + (Number(r.amount) || 0), 0);
  const offsetLeft = supplierDebt - offsetUsed;

  const update = (i, patch) => setRows((rs) => rs.map((r, idx) => (idx === i ? { ...r, ...patch } : r)));
  const addRow = () => setRows((rs) => [...rs, {
    method: "cash", amount: remaining > 0 ? remaining : "", bank_account_id: null,
    installment_provider: "", installment_contract_code: "", note: "",
  }]);
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
                <option value="debt">Khách nợ shop</option>
                {supplierDebt > 0 && <option value="debt_offset">Bù trừ công nợ</option>}
              </select>
              <input
                type="number"
                value={r.amount}
                onChange={(e) => update(i, { amount: e.target.value })}
                placeholder={r.method === "debt_offset" ? "Số tiền bù trừ" : r.method === "debt" ? "Số tiền ghi nợ" : "Số tiền"}
                className="w-32 rounded-lg border border-slate-200 px-2 py-1.5 text-sm"
              />
              {rows.length > 1 && (
                <button type="button" onClick={() => removeRow(i)} className="text-rose-400 hover:text-rose-600"><X size={15} /></button>
              )}
            </div>
            {r.method === "debt" && (
              <div className="pl-6 space-y-1.5">
                <p className="text-[11px] text-slate-400">
                  Không thu tiền. Khoản này hiện ở cột Công nợ của đơn, thu dần sau bằng nút "Thu nợ".
                </p>
                <button
                  type="button"
                  onClick={() => update(i, { amount: String(Math.max(0, (Number(r.amount) || 0) + remaining)) })}
                  className="text-[11px] text-brand-600 hover:underline"
                >
                  Ghi nợ phần còn lại ({fmtVND(Math.max(0, remaining))})
                </button>
              </div>
            )}
            {r.method === "bank_transfer" && (
              <div className="pl-6">
                <BankSelect
                  banks={banks}
                  value={r.bank_account_id}
                  onChange={(v) => update(i, { bank_account_id: v })}
                  className="w-full"
                />
                {banks.length === 0 && (
                  <p className="text-[11px] text-amber-600 mt-1">
                    Chưa khai báo tài khoản ngân hàng nào — nhờ Quản lý bổ sung.
                  </p>
                )}
              </div>
            )}
            {r.method === "debt_offset" && (
              <div className="pl-6 space-y-1.5">
                <p className="text-[11px] text-slate-400">
                  Trừ vào khoản phải trả với <span className="font-medium text-slate-500">{supplierName}</span>.
                  Hệ thống lập biên bản bù trừ và phân bổ vào các phiếu nhập máy còn nợ, cũ nhất trước.
                </p>
                <button
                  type="button"
                  onClick={() => update(i, { amount: String(Math.max(0, Math.min(supplierDebt, (Number(r.amount) || 0) + offsetLeft, total))) })}
                  className="text-[11px] text-brand-600 hover:underline"
                >
                  Bù trừ tối đa ({fmtVND(Math.max(0, Math.min(supplierDebt, total)))})
                </button>
                {offsetLeft < 0 && (
                  <p className="text-[11px] text-rose-500">
                    Vượt quá công nợ hiện có {fmtVND(supplierDebt)} — giảm bớt {fmtVND(-offsetLeft)}.
                  </p>
                )}
              </div>
            )}
            {r.method === "installment" && (
              <div className="grid grid-cols-2 gap-2 pl-6">
                <select
                  value={r.installment_provider}
                  onChange={(e) => update(i, { installment_provider: e.target.value })}
                  className="rounded-lg border border-slate-200 px-2 py-1.5 text-xs"
                >
                  <option value="">— Chọn đơn vị trả góp —</option>
                  {providers.map((p) => <option key={p.id} value={p.name}>{p.name}</option>)}
                </select>
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
        "text-xs px-3 py-2 rounded-lg space-y-0.5",
        remaining === 0 ? "bg-emerald-50 text-emerald-700"
          : remaining > 0 ? "bg-amber-50 text-amber-700"
          : "bg-rose-50 text-rose-600"
      )}>
        <div>Đã phân bổ {fmtVND(paid)} / {fmtVND(total)}
          {remaining > 0 && ` — còn thiếu ${fmtVND(remaining)}`}
          {remaining < 0 && ` — thừa ${fmtVND(-remaining)}, vui lòng giảm bớt`}
        </div>
        {debtDeclared > 0 && (
          <div className="opacity-80">
            Thực thu {fmtVND(paid - debtDeclared)} · khách nợ {fmtVND(debtDeclared)}
          </div>
        )}
      </div>
    </div>
  );
}

function OrderForm({ onCancel, onSaved, employee }) {
  const [customer, setCustomer] = useState(null);
  const [device, setDevice] = useState(null);
  const [manualDeviceMode, setManualDeviceMode] = useState(false);
  const [manualDevice, setManualDevice] = useState({ imei: "", model: "", storage: "", color: "", condition: "used" });
  const [salePrice, setSalePrice] = useState("");
  const [discount, setDiscount] = useState("0");
  const [notes, setNotes] = useState("");
  const [payments, setPayments] = useState([{
    method: "cash", amount: "", bank_account_id: null,
    installment_provider: "", installment_contract_code: "", note: "",
  }]);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const [dueDate, setDueDate] = useState("");
  const [linkedSupplier, setLinkedSupplier] = useState(null);
  const [supplierDebt, setSupplierDebt] = useState(0);

  useEffect(() => {
    if (device && !salePrice) setSalePrice(device.sale_price ?? "");
  }, [device]); // eslint-disable-line

  // Khách này có đồng thời là Nhà cung cấp không? Nếu có, lấy số cửa hàng đang nợ họ.
  useEffect(() => {
    let active = true;
    if (!customer) { setLinkedSupplier(null); setSupplierDebt(0); return; }
    (async () => {
      // Số dư lấy từ SỔ CÁI, không đọc cột cache
      const { data: rows } = await supabase.rpc("partner_balance_by_customer", { p_customer_id: customer.id });
      if (!active) return;
      const bal = Array.isArray(rows) ? rows[0] : rows;
      if (!bal) { setLinkedSupplier(null); setSupplierDebt(0); return; }
      const payable = Number(bal.payable) || 0;   // >0 = cửa hàng đang nợ họ
      setLinkedSupplier(payable > 0 ? { id: bal.partner_id, name: bal.partner_name } : null);
      setSupplierDebt(Math.max(0, payable));
    })();
    return () => { active = false; };
  }, [customer]);

  const total = Math.max(0, (Number(salePrice) || 0) - (Number(discount) || 0));
  const realPaidNow = payments.filter((r) => r.method !== "debt")
    .reduce((s, r) => s + (Number(r.amount) || 0), 0);

  const submit = async (e) => {
    e.preventDefault();
    setError("");
    if (!customer) { setError("Vui lòng chọn khách hàng."); return; }
    if (!manualDeviceMode && !device) { setError("Vui lòng chọn máy bán (còn hàng)."); return; }
    if (manualDeviceMode && (!manualDevice.imei.trim() || !manualDevice.model.trim())) {
      setError("Vui lòng nhập đủ IMEI và model của máy đang bán để chờ Quản lý đối soát.");
      return;
    }
    if (!salePrice || Number(salePrice) <= 0) { setError("Vui lòng nhập giá bán hợp lệ."); return; }
    // Dòng để trống = chưa thu -> bỏ qua, cho phép đơn nợ toàn bộ
    const filledRows = payments.filter((r) => String(r.amount).trim() !== "");
    // Dòng "Khách nợ shop" chỉ để khai báo cho rõ, KHÔNG tạo phiếu thu.
    // Phần chưa thu tự thành công nợ qua sổ cái.
    const declaredDebt = filledRows.filter((r) => r.method === "debt")
      .reduce((s, r) => s + (Number(r.amount) || 0), 0);
    const activePayments = filledRows.filter((r) => r.method !== "debt");
    const paidTotal = activePayments.reduce((s, r) => s + (Number(r.amount) || 0), 0);
    if (Math.round(paidTotal) > Math.round(total)) {
      setError("Tổng tiền thu đang lớn hơn tổng tiền đơn hàng."); return;
    }
    if (declaredDebt > 0 && Math.round(paidTotal + declaredDebt) !== Math.round(total)) {
      setError(`Tiền thu (${fmtVND(paidTotal)}) cộng khoản ghi nợ (${fmtVND(declaredDebt)}) phải bằng đúng tổng tiền đơn hàng (${fmtVND(total)}).`);
      return;
    }
    if (activePayments.some((r) => Number(r.amount) <= 0)) { setError("Số tiền thanh toán phải lớn hơn 0."); return; }

    if (activePayments.some((r) => r.method === "bank_transfer" && !r.bank_account_id)) {
      setError("Vui lòng chọn tài khoản ngân hàng nhận tiền cho dòng chuyển khoản.");
      return;
    }
    if (activePayments.some((r) => r.method === "installment" && !r.installment_provider?.trim())) {
      setError("Vui lòng chọn đơn vị hỗ trợ trả góp.");
      return;
    }

    const offsetTotal = activePayments.filter((r) => r.method === "debt_offset")
      .reduce((s, r) => s + (Number(r.amount) || 0), 0);
    if (offsetTotal > 0) {
      if (!linkedSupplier) { setError("Cửa hàng không có khoản phải trả nào với đối tác này nên không bù trừ được."); return; }
      const { data: rows } = await supabase.rpc("partner_balance_by_customer", { p_customer_id: customer.id });
      const bal = Array.isArray(rows) ? rows[0] : rows;
      const available = Math.max(0, Number(bal?.payable) || 0);
      setSupplierDebt(available);
      if (offsetTotal > available) {
        setError(`Cửa hàng chỉ đang nợ ${fmtVND(available)} — không cấn trừ được ${fmtVND(offsetTotal)}.`);
        return;
      }
    }
    if (manualDeviceMode) {
      const { data: existing } = await supabase.from("devices").select("id, status").eq("imei", manualDevice.imei.trim()).maybeSingle();
      if (existing) {
        setError(`IMEI ${manualDevice.imei.trim()} thực ra ĐÃ có trong kho (${DEVICE_STATUS_LABELS[existing.status]}) — vui lòng tắt "Bán tạm" và chọn máy này từ danh sách bình thường.`);
        return;
      }
    }

    setSaving(true);
    try {
      let saleDevice = device;
      if (manualDeviceMode) {
        const { data: newDevice, error: mdErr } = await supabase.from("devices").insert({
          imei: manualDevice.imei.trim(), model: manualDevice.model.trim(),
          storage: manualDevice.storage.trim() || null, color: manualDevice.color.trim() || null,
          condition: manualDevice.condition, status: "pending_reconciliation", cost_price: null,
          sale_price: Number(salePrice), supplier: "Bán tạm — chờ Quản lý đối soát kho",
          import_date: new Date().toISOString().slice(0, 10),
          created_by: employee.id, updated_by: employee.id, store_id: employee.store_id,
        }).select().maybeSingle();
        if (mdErr) throw mdErr;
        saleDevice = newDevice;
        await supabase.from("audit_logs").insert({
          table_name: "devices", record_id: newDevice.id, action: "create", new_data: newDevice, performed_by: employee.id, store_id: employee.store_id,
        });
      }

      const orderPayload = {
        customer_id: customer.id,
        device_id: saleDevice.id,
        sale_price: Number(salePrice),
        discount: Number(discount) || 0,
        total_amount: total,
        notes: notes.trim() || null,
        due_date: Math.round(paidTotal) < Math.round(total) ? (dueDate || null) : null,   // paidTotal đã loại dòng ghi nợ
        status: manualDeviceMode ? "pending_stock" : "completed",
        created_by: employee.id,
        updated_by: employee.id,
        store_id: employee.store_id,
      };
      const { data: order, error: orderErr } = await supabase.from("sales_orders").insert(orderPayload).select().maybeSingle();
      if (orderErr) throw orderErr;

      const paymentRows = activePayments.map((r) => ({
        order_id: order.id,
        method: r.method,
        amount: Number(r.amount),
        bank_account_id: r.method === "bank_transfer" ? (r.bank_account_id || null) : null,
        installment_provider: r.method === "installment" ? (r.installment_provider.trim() || null) : null,
        installment_contract_code: r.method === "installment" ? (r.installment_contract_code.trim() || null) : null,
        note: r.note?.trim() || null,
        created_by: employee.id,
        store_id: employee.store_id,
      }));
      // Sổ cái tự ghi bút toán khi phiếu thu được tạo. Riêng dòng "Cấn trừ công nợ"
      // sẽ tự lập BIÊN BẢN BÙ TRỪ và giảm cả hai bên phải thu/phải trả.
      if (paymentRows.length > 0) {
        const { error: payErr } = await supabase.from("order_payments").insert(paymentRows);
        if (payErr) throw payErr;
      }

      const { error: contractErr } = await supabase.from("contracts").insert({ order_id: order.id, created_by: employee.id, store_id: employee.store_id });
      if (contractErr) throw contractErr;

      const auditRows = [
        { table_name: "sales_orders", record_id: order.id, action: "create", new_data: order, performed_by: employee.id, store_id: employee.store_id },
      ];

      if (!manualDeviceMode) {
        const { data: soldDevice, error: devErr } = await supabase
          .from("devices").update({ status: "sold", updated_by: employee.id }).eq("id", saleDevice.id).select().maybeSingle();
        if (devErr) throw devErr;
        auditRows.push({ table_name: "devices", record_id: saleDevice.id, action: "update", old_data: saleDevice, new_data: soldDevice, performed_by: employee.id, store_id: employee.store_id });
      }

      await supabase.from("audit_logs").insert(auditRows);

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
          <div className="flex items-center justify-between mb-1">
            <span className="text-xs font-medium text-slate-500 block">Máy bán (IMEI) *</span>
            <button
              type="button"
              onClick={() => { setManualDeviceMode((s) => !s); setDevice(null); }}
              className={classNames("text-xs hover:underline", manualDeviceMode ? "text-rose-600" : "text-brand-600")}
            >
              {manualDeviceMode ? "✕ Tắt bán tạm, chọn máy trong kho" : "Không tìm thấy IMEI trong kho?"}
            </button>
          </div>
          {!manualDeviceMode ? (
            <DevicePicker value={device} onSelect={setDevice} />
          ) : (
            <div className="bg-rose-50 border border-rose-200 rounded-xl p-3 space-y-2">
              <p className="text-xs text-rose-700 flex items-center gap-1.5">
                <ShieldAlert size={14} className="shrink-0" />
                Bán tạm khi IMEI thực tế không khớp dữ liệu Kho — đơn sẽ ở trạng thái "Chờ đối soát kho" cho tới khi Quản lý xác nhận và cập nhật lại Kho hàng.
              </p>
              <div className="grid grid-cols-2 gap-2">
                <input
                  value={manualDevice.imei}
                  onChange={(e) => setManualDevice((f) => ({ ...f, imei: e.target.value }))}
                  placeholder="Số IMEI thực tế trên máy *"
                  className="rounded-lg border border-rose-200 px-2 py-1.5 text-xs bg-white"
                />
                <input
                  value={manualDevice.model}
                  onChange={(e) => setManualDevice((f) => ({ ...f, model: e.target.value }))}
                  placeholder="Model máy *"
                  list="dl-models-manual-order"
                  className="rounded-lg border border-rose-200 px-2 py-1.5 text-xs bg-white"
                />
                <input
                  value={manualDevice.storage}
                  onChange={(e) => setManualDevice((f) => ({ ...f, storage: e.target.value }))}
                  placeholder="Dung lượng"
                  list="dl-storage"
                  className="rounded-lg border border-rose-200 px-2 py-1.5 text-xs bg-white"
                />
                <input
                  value={manualDevice.color}
                  onChange={(e) => setManualDevice((f) => ({ ...f, color: e.target.value }))}
                  placeholder="Màu sắc"
                  list="dl-colors-manual-order"
                  className="rounded-lg border border-rose-200 px-2 py-1.5 text-xs bg-white"
                />
                <datalist id="dl-models-manual-order">{IPHONE_MODEL_LIST.map((m) => <option key={m} value={m} />)}</datalist>
                <datalist id="dl-colors-manual-order">{coloroptionsForModel(manualDevice.model).map((c) => <option key={c} value={c} />)}</datalist>
              </div>
            </div>
          )}
        </div>
        <div className="grid sm:grid-cols-2 gap-3">
          <TextField label="Giá bán (đ) *" type="number" value={salePrice} onChange={(e) => setSalePrice(e.target.value)} />
          <TextField label="Giảm giá (đ)" type="number" value={discount} onChange={(e) => setDiscount(e.target.value)} />
        </div>
        <div className="flex justify-between items-center bg-brand-50 rounded-xl px-3 py-2.5 text-sm">
          <span className="text-slate-500">Tổng tiền đơn hàng</span>
          <span className="font-semibold text-brand-700">{fmtVND(total)}</span>
        </div>
        {linkedSupplier && supplierDebt > 0 && (
          <div className="bg-indigo-50 border border-indigo-200 rounded-xl px-3 py-2.5 text-xs text-indigo-800 flex items-start gap-2">
            <Banknote size={15} className="shrink-0 mt-0.5" />
            <span>
              Cửa hàng đang có khoản <span className="font-medium">phải trả {fmtVND(supplierDebt)}</span> với
              đối tác <span className="font-medium">{linkedSupplier.name}</span> (họ từng bán máy cho cửa hàng).
              Chọn hình thức <span className="font-medium">Bù trừ công nợ</span> bên dưới để lập biên bản bù trừ.
            </span>
          </div>
        )}
        <div>
          <span className="text-xs font-medium text-slate-500 mb-1 block">Hình thức thanh toán *</span>
          <PaymentRows
            rows={payments} setRows={setPayments} total={total}
            supplierDebt={supplierDebt} supplierName={linkedSupplier?.name || ""}
          />
        </div>
        {realPaidNow < total && (
          <div className="bg-amber-50 border border-amber-200 rounded-xl p-3 space-y-2">
            <p className="text-xs text-amber-800">
              Khách còn nợ <span className="font-medium">{fmtVND(total - realPaidNow)}</span> — đơn sẽ được ghi nhận là công nợ và thu sau ở màn Đơn hàng bán.
            </p>
            <TextField label="Hẹn ngày thanh toán" type="date" value={dueDate} onChange={(e) => setDueDate(e.target.value)} />
          </div>
        )}
        <TextField label="Ghi chú" value={notes} onChange={(e) => setNotes(e.target.value)} />
        {error && <p className="text-xs text-rose-600 bg-rose-50 rounded-lg px-3 py-2">{error}</p>}
        <div className="flex gap-2">
          <button type="submit" disabled={saving} className="bg-brand-600 hover:bg-brand-700 text-white rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-2 disabled:opacity-60">
            {saving && <Loader2 size={14} className="animate-spin" />} {manualDeviceMode ? "Bán tạm — chờ đối soát kho" : "Hoàn tất đơn hàng"}
          </button>
          <button type="button" onClick={onCancel} className="text-slate-500 text-sm px-4 py-2">Hủy</button>
        </div>
      </form>
    </Card>
  );
}

function ReconcileModal({ order, device, employee, onClose, onDone }) {
  const [form, setForm] = useState({
    imei: device?.imei || "", model: device?.model || "", storage: device?.storage || "", color: device?.color || "",
    condition: device?.condition || "used", condition_percent: device?.condition_percent ?? "",
    cost_price: device?.cost_price ?? "",
    supplier: (device?.supplier === "Bán tạm — chờ Quản lý đối soát kho" ? "" : device?.supplier) || "",
  });
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));

  const submit = async (e) => {
    e.preventDefault();
    if (!form.imei.trim() || !form.model.trim()) { setError("Vui lòng nhập đủ IMEI và model."); return; }
    if (form.cost_price === "" || Number(form.cost_price) < 0) { setError("Vui lòng nhập giá vốn hợp lệ cho máy này."); return; }
    setError(""); setSaving(true);
    try {
      const payload = {
        imei: form.imei.trim(), model: form.model.trim(),
        storage: form.storage.trim() || null, color: form.color.trim() || null,
        condition: form.condition, condition_percent: form.condition_percent === "" ? null : Number(form.condition_percent),
        cost_price: Number(form.cost_price), status: "sold",
        supplier: form.supplier.trim() || null,
        updated_by: employee.id,
      };
      const { data: updatedDevice, error: devErr } = await supabase.from("devices").update(payload).eq("id", device.id).select().maybeSingle();
      if (devErr) throw devErr;

      const { data: updatedOrder, error: orderErr } = await supabase.from("sales_orders")
        .update({ status: "completed", updated_by: employee.id }).eq("id", order.id).select().maybeSingle();
      if (orderErr) throw orderErr;

      await supabase.from("audit_logs").insert([
        { table_name: "devices", record_id: device.id, action: "update", old_data: device, new_data: updatedDevice, performed_by: employee.id, store_id: employee.store_id },
        { table_name: "sales_orders", record_id: order.id, action: "update", old_data: order, new_data: updatedOrder, performed_by: employee.id, store_id: employee.store_id },
      ]);

      onDone();
    } catch (err) {
      if (err.code === "23505" || /duplicate/i.test(err.message || "")) {
        setError(`IMEI ${form.imei.trim()} đã tồn tại ở 1 máy khác trong kho — kiểm tra lại, có thể đây là bản ghi trùng cần gộp thủ công.`);
      } else {
        setError(err.message || "Không lưu được, vui lòng thử lại.");
      }
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/40 flex items-center justify-center p-4">
      <Card className="w-full max-w-md p-5">
        <div className="flex items-center justify-between mb-4">
          <p className="font-semibold text-slate-800 text-sm">Đối soát kho — {order.order_code}</p>
          <button onClick={onClose} className="text-slate-400 hover:text-rose-600"><X size={16} /></button>
        </div>
        <p className="text-xs text-slate-400 mb-3">
          Kiểm tra lại thông tin máy nhân viên đã nhập lúc bán, sửa nếu cần, bổ sung giá vốn rồi xác nhận — đơn hàng sẽ chuyển thành "Hoàn tất" và máy chuyển "Đã bán" trong Kho.
        </p>
        <form onSubmit={submit} className="grid grid-cols-2 gap-3">
          <TextField label="Số IMEI *" value={form.imei} onChange={set("imei")} className="col-span-2" />
          <label className="block col-span-2">
            <span className="text-xs font-medium text-slate-500 mb-1 block">Model máy *</span>
            <ModelPicker value={form.model} onSelect={(v) => setForm((f) => ({ ...f, model: v }))} />
          </label>
          <TextField label="Dung lượng" value={form.storage} onChange={set("storage")} list="dl-storage" />
          <TextField label="Màu sắc" value={form.color} onChange={set("color")} list="dl-colors-reconcile" />
          <datalist id="dl-colors-reconcile">{coloroptionsForModel(form.model).map((c) => <option key={c} value={c} />)}</datalist>
          <datalist id="dl-storage"><option value="64GB" /><option value="128GB" /><option value="256GB" /><option value="512GB" /><option value="1TB" /></datalist>
          <label className="block">
            <span className="text-xs font-medium text-slate-500 mb-1 block">Tình trạng</span>
            <select value={form.condition} onChange={set("condition")} className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm text-slate-800 outline-none focus:ring-2 focus:ring-brand-400/40 focus:border-brand-500 transition">
              <option value="new">Máy mới</option>
              <option value="used">Máy cũ</option>
            </select>
          </label>
          <TextField label="Độ mới (%)" type="number" min="0" max="100" value={form.condition_percent} onChange={set("condition_percent")} />
          <TextField label="Nhà cung cấp / nguồn nhập" value={form.supplier} onChange={set("supplier")} placeholder="Ví dụ: khách lẻ, tên NCC..." className="col-span-2" />
          <TextField label="Giá vốn (đ) *" type="number" value={form.cost_price} onChange={set("cost_price")} className="col-span-2" />
          {error && <p className="text-xs text-rose-600 bg-rose-50 rounded-lg px-3 py-2 col-span-2">{error}</p>}
          <div className="col-span-2 flex gap-2 mt-1">
            <button type="submit" disabled={saving} className="bg-brand-600 hover:bg-brand-700 text-white rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-2 disabled:opacity-60">
              {saving && <Loader2 size={14} className="animate-spin" />} Xác nhận đã cập nhật kho
            </button>
            <button type="button" onClick={onClose} className="text-slate-500 text-sm px-4 py-2">Hủy</button>
          </div>
        </form>
      </Card>
    </div>
  );
}


function CollectDebtModal({ order, employee, onClose, onDone }) {
  const debt = Math.max(0, Number(order.total_amount) - Number(order.paid_amount || 0));
  const [amount, setAmount] = useState(String(debt));
  const [method, setMethod] = useState("cash");
  const [bankId, setBankId] = useState(null);
  const [note, setNote] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const { banks } = usePaymentOptions();

  const submit = async () => {
    const n = Number(amount) || 0;
    if (n <= 0) { setError("Vui lòng nhập số tiền thu hợp lệ."); return; }
    if (n > debt) { setError(`Khách chỉ còn nợ ${fmtVND(debt)}.`); return; }
    if (method === "bank_transfer" && !bankId) { setError("Vui lòng chọn tài khoản nhận tiền."); return; }
    setSaving(true); setError("");
    const { data: pay, error: err } = await supabase.from("order_payments").insert({
      order_id: order.id, method, amount: n,
      bank_account_id: method === "bank_transfer" ? bankId : null,
      note: note.trim() || "Thu công nợ",
      created_by: employee.id, store_id: employee.store_id,
    }).select().maybeSingle();
    setSaving(false);
    if (err) { setError(err.message); return; }
    await supabase.from("audit_logs").insert({
      table_name: "order_payments", record_id: pay?.id, action: "create",
      new_data: pay, performed_by: employee.id, store_id: employee.store_id,
    });
    onDone();
  };

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/40 flex items-center justify-center p-4">
      <Card className="w-full max-w-md p-5">
        <div className="flex items-center justify-between mb-4">
          <p className="font-semibold text-slate-800 text-sm">Thu công nợ — {order.order_code}</p>
          <button onClick={onClose} className="text-slate-400 hover:text-rose-600"><X size={16} /></button>
        </div>
        <div className="space-y-3">
        <div className="bg-amber-50 rounded-xl px-3 py-2.5 text-xs text-amber-800">
          Tổng đơn {fmtVND(order.total_amount)} · đã thu {fmtVND(order.paid_amount || 0)} ·
          <span className="font-medium"> còn nợ {fmtVND(debt)}</span>
          {order.due_date && <> · hẹn trả {fmtDate(order.due_date)}</>}
        </div>
        <div>
          <span className="text-xs font-medium text-slate-500 mb-1 block">Hình thức thu</span>
          <select
            value={method} onChange={(e) => setMethod(e.target.value)}
            className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm text-slate-800 outline-none focus:ring-2 focus:ring-brand-400/40 focus:border-brand-500 transition"
          >
            <option value="cash">Tiền mặt</option>
            <option value="bank_transfer">Chuyển khoản</option>
          </select>
        </div>
        {method === "bank_transfer" && (
          <div>
            <span className="text-xs font-medium text-slate-500 mb-1 block">Tài khoản nhận tiền *</span>
            <BankSelect banks={banks} value={bankId} onChange={setBankId} className="w-full !text-sm !px-3 !py-2 !rounded-xl" />
          </div>
        )}
        <TextField label="Số tiền thu (đ)" type="number" value={amount} onChange={(e) => setAmount(e.target.value)} />
        <TextField label="Ghi chú" value={note} onChange={(e) => setNote(e.target.value)} />
        {error && <p className="text-xs text-rose-600 bg-rose-50 rounded-lg px-3 py-2">{error}</p>}
        <div className="flex gap-2">
          <button
            onClick={submit} disabled={saving}
            className="flex-1 bg-brand-600 hover:bg-brand-700 disabled:opacity-60 text-white rounded-xl py-2.5 text-sm font-medium flex items-center justify-center gap-2"
          >
            {saving && <Loader2 size={15} className="animate-spin" />} Xác nhận thu
          </button>
          <button onClick={onClose} className="px-4 rounded-xl border border-slate-200 text-sm text-slate-500 hover:bg-slate-50">Hủy</button>
        </div>
        </div>
      </Card>
    </div>
  );
}

function PayCustomerDiffModal({ order, employee, onClose, onDone }) {
  const debt = Math.max(0, Number(order.shop_debt || 0));
  const [amount, setAmount] = useState(String(debt));
  const [method] = useState("bank_transfer");
  const [bankId, setBankId] = useState(null);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const { banks } = usePaymentOptions();

  const submit = async () => {
    const n = Number(amount) || 0;
    if (n <= 0) { setError("Vui lòng nhập số tiền hợp lệ."); return; }
    if (n > debt) { setError(`Cửa hàng chỉ còn nợ khách ${fmtVND(debt)}.`); return; }
    if (method === "bank_transfer" && !bankId) { setError("Vui lòng chọn tài khoản chuyển tiền."); return; }
    setSaving(true); setError("");
    const label = method === "cash" ? "tiền mặt"
      : `chuyển khoản ${banks.find((b) => b.id === bankId)?.short_label || ""}`.trim();
    const { error: err } = await supabase.rpc("pay_customer_diff", {
      p_sales_order_id: order.id, p_amount: n,
      p_note: `Trả khách phần chênh máy đổi (${label})`,
    });
    setSaving(false);
    if (err) { setError(err.message); return; }
    await supabase.from("audit_logs").insert({
      table_name: "sales_orders", record_id: order.id, action: "update",
      new_data: { tra_khach: n, hinh_thuc: label }, performed_by: employee.id, store_id: employee.store_id,
    });
    onDone();
  };

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/40 flex items-center justify-center p-4">
      <Card className="w-full max-w-md p-5">
        <div className="flex items-center justify-between mb-4">
          <p className="font-semibold text-slate-800 text-sm">Trả tiền khách — {order.order_code}</p>
          <button onClick={onClose} className="text-slate-400 hover:text-rose-600"><X size={16} /></button>
        </div>
        <div className="space-y-3">
          <div className="bg-indigo-50 rounded-xl px-3 py-2.5 text-xs text-indigo-800">
            Máy khách đổi có giá trị cao hơn máy bán. Cửa hàng còn nợ khách{" "}
            <span className="font-medium">{fmtVND(debt)}</span>.
          </div>
          <div>
            <span className="text-xs font-medium text-slate-500 mb-1 block">Tài khoản chuyển tiền *</span>
            <BankSelect banks={banks} value={bankId} onChange={setBankId} className="w-full !text-sm !px-3 !py-2 !rounded-xl" />
            <p className="text-[11px] text-slate-400 mt-1">Mọi khoản chi ra đều qua chuyển khoản.</p>
          </div>
          <TextField label="Số tiền trả (đ)" type="number" value={amount} onChange={(e) => setAmount(e.target.value)} />
          {error && <p className="text-xs text-rose-600 bg-rose-50 rounded-lg px-3 py-2">{error}</p>}
          <div className="flex gap-2">
            <button onClick={submit} disabled={saving}
              className="flex-1 bg-brand-600 hover:bg-brand-700 disabled:opacity-60 text-white rounded-xl py-2.5 text-sm font-medium flex items-center justify-center gap-2">
              {saving && <Loader2 size={15} className="animate-spin" />} Xác nhận đã trả
            </button>
            <button onClick={onClose} className="px-4 rounded-xl border border-slate-200 text-sm text-slate-500 hover:bg-slate-50">Hủy</button>
          </div>
        </div>
      </Card>
    </div>
  );
}

function OrderRow({ order, employee, onDeleted, onReconciled }) {
  const [expanded, setExpanded] = useState(false);
  const [detail, setDetail] = useState(null);
  const [loadingDetail, setLoadingDetail] = useState(false);
  const [printType, setPrintType] = useState(null);
  const [showReconcile, setShowReconcile] = useState(false);
  const [showCollect, setShowCollect] = useState(false);
  const [showPayDiff, setShowPayDiff] = useState(false);
  const shopDebt = Math.max(0, Number(order.shop_debt || 0));
  const orderDebt = order.customer_debt !== undefined
    ? Math.max(0, Number(order.customer_debt))
    : Math.max(0, Number(order.total_amount) - Number(order.paid_amount || 0));
  const overdue = orderDebt > 0 && order.due_date && new Date(order.due_date) < new Date(new Date().toDateString());
  const canCollect = employee.role !== "ke_toan" && orderDebt > 0 && order.status !== "cancelled";
  const canPayDiff = employee.role !== "ke_toan" && shopDebt > 0 && order.status !== "cancelled";
  const canDelete = employee.role === "quan_ly" && order.status !== "completed";
  const canReconcile = employee.role === "quan_ly" && order.status === "pending_stock";

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
    if (order.status === "completed") { alert("Đơn hàng đã Hoàn tất — không thể xóa để bảo vệ dữ liệu đã chốt."); return; }
    if (!confirm(`Xóa đơn hàng "${order.order_code}"? Máy sẽ không tự động chuyển lại trạng thái Còn hàng.`)) return;
    const { error } = await supabase.from("sales_orders").delete().eq("id", order.id);
    if (error) { alert(error.message); return; }
    await supabase.from("audit_logs").insert({ table_name: "sales_orders", record_id: order.id, action: "delete", old_data: order, performed_by: employee.id, store_id: employee.store_id });
    onDeleted();
  };

  return (
    <>
      <tr className="border-b border-slate-50 last:border-0 hover:bg-slate-50/60 cursor-pointer" onClick={loadDetail}>
        <td className="px-3 py-2.5 font-medium text-slate-700 whitespace-nowrap">{order.order_code}</td>
        <td className="px-3 py-2.5 text-slate-500">{fmtDate(order.created_at)}</td>
        <td className="px-3 py-2.5 text-slate-600">{fmtVND(order.total_amount)}</td>
        <td className="px-3 py-2.5 whitespace-nowrap">
          {orderDebt > 0 ? (
            <span className={classNames(
              "text-xs px-2 py-0.5 rounded-full",
              overdue ? "bg-rose-50 text-rose-600" : "bg-amber-50 text-amber-700"
            )}>
              Khách nợ {fmtVND(orderDebt)}{overdue ? " · quá hạn" : ""}
            </span>
          ) : shopDebt > 0 ? (
            <span className="text-xs px-2 py-0.5 rounded-full bg-indigo-50 text-indigo-700">
              Shop nợ khách {fmtVND(shopDebt)}
            </span>
          ) : (
            <span className="text-xs px-2 py-0.5 rounded-full bg-emerald-50 text-emerald-700">Đã thu đủ</span>
          )}
        </td>
        <td className="px-3 py-2.5">
          <span className={classNames(
            "text-xs px-2 py-0.5 rounded-full",
            order.status === "completed" ? "bg-emerald-50 text-emerald-700"
              : order.status === "pending_stock" ? "bg-amber-50 text-amber-700"
              : "bg-rose-50 text-rose-600"
          )}>
            {order.status === "completed" ? "Hoàn tất" : order.status === "pending_stock" ? "Chờ đối soát kho" : "Đã hủy"}
          </span>
        </td>
        <td className="px-3 py-2.5 text-right whitespace-nowrap" onClick={(e) => e.stopPropagation()}>
          {canCollect && (
            <button
              onClick={() => setShowCollect(true)}
              className="text-xs text-brand-600 hover:underline mr-3"
            >
              Thu nợ
            </button>
          )}
          {canPayDiff && (
            <button
              onClick={() => setShowPayDiff(true)}
              className="text-xs text-indigo-600 hover:underline mr-3"
            >
              Trả khách
            </button>
          )}
          {showPayDiff && (
            <PayCustomerDiffModal
              order={order} employee={employee}
              onClose={() => setShowPayDiff(false)}
              onDone={() => { setShowPayDiff(false); setDetail(null); onDeleted(); }}
            />
          )}
          {showCollect && (
            <CollectDebtModal
              order={order} employee={employee}
              onClose={() => setShowCollect(false)}
              onDone={() => { setShowCollect(false); setDetail(null); onDeleted(); }}
            />
          )}
          {loadingDetail ? <Loader2 size={14} className="animate-spin inline text-slate-300" /> : (
            <ChevronDown size={15} className={classNames("inline text-slate-300 transition-transform", expanded && "rotate-180")} onClick={loadDetail} />
          )}
        </td>
      </tr>
      {expanded && detail && (
        <tr>
          <td colSpan={6} className="bg-slate-50/70 px-4 py-4">
            {order.status === "pending_stock" && (
              <div className="bg-amber-50 border border-amber-200 rounded-xl px-3 py-2.5 mb-3 text-xs text-amber-800 flex items-center gap-2">
                <ShieldAlert size={15} className="shrink-0" />
                Đơn này được tạo với IMEI không khớp Kho hàng (nhân viên nhập tay lúc bán) — đơn chỉ hợp lệ hoàn toàn sau khi Quản lý đối soát và cập nhật Kho.
              </div>
            )}
            <div className="grid sm:grid-cols-2 gap-4 text-sm">
              <div>
                <p className="text-xs text-slate-400 mb-1">Khách hàng</p>
                <p className="font-medium text-slate-700">{detail.customer?.full_name}</p>
                <p className="text-xs text-slate-400">{detail.customer?.phone}</p>
              </div>
              <div>
                <p className="text-xs text-slate-400 mb-1">Máy đã bán</p>
                <p className="font-medium text-slate-700">{detail.device?.model} {[detail.device?.storage, detail.device?.color].filter(Boolean).join(" · ")}</p>
                <p className="text-xs text-slate-400">{detail.device?.imei ? `IMEI ${detail.device.imei}` : "Chưa có IMEI"}</p>
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
            <div className="flex gap-3 mt-3 items-center">
              <button onClick={() => setPrintType({ kind: "contract" })} className="text-brand-600 hover:underline text-xs flex items-center gap-1">
                <Printer size={12} /> In hợp đồng
              </button>
              {canReconcile && (
                <button onClick={() => setShowReconcile(true)} className="bg-amber-500 hover:bg-amber-600 text-white rounded-lg px-3 py-1.5 text-xs font-medium flex items-center gap-1">
                  <ShieldAlert size={12} /> Đối soát kho
                </button>
              )}
              {canDelete && (
                <button onClick={remove} className="text-rose-500 hover:underline text-xs flex items-center gap-1">
                  <Trash2 size={12} /> Xóa đơn
                </button>
              )}
            </div>
          </td>
        </tr>
      )}
      {showReconcile && detail && (
        <ReconcileModal
          order={order} device={detail.device} employee={employee}
          onClose={() => setShowReconcile(false)}
          onDone={() => { setShowReconcile(false); setDetail(null); onReconciled?.(); }}
        />
      )}
      {printType && detail && (
        <PrintDocModal
          type={printType.kind}
          order={order}
          customer={detail.customer}
          device={detail.device}
          payments={printType.kind === "receipt" ? printType.payments : detail.payments}
          contract={detail.contract}
          storeName={employee.stores?.name}
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

/* -------------------------------------------------------------- */
/* Xuất bán nội bộ giữa các cửa hàng                              */
/* -------------------------------------------------------------- */

const TRANSFER_STATUS_LABELS = { pending: "Chờ nhận", received: "Đã nhận", cancelled: "Đã hủy" };
const TRANSFER_STATUS_STYLES = {
  pending: "bg-amber-50 text-amber-700",
  received: "bg-emerald-50 text-emerald-700",
  cancelled: "bg-slate-100 text-slate-400",
};

function InternalTransferForm({ employee, onCancel, onSaved }) {
  const [stores, setStores] = useState([]);
  const [devices, setDevices] = useState([]);
  const [toStore, setToStore] = useState("");
  const [deviceId, setDeviceId] = useState("");
  const [price, setPrice] = useState("");
  const [note, setNote] = useState("");
  const [search, setSearch] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const canSeeCost = employee.role !== "nhan_vien";

  useEffect(() => {
    (async () => {
      const [{ data: s }, { data: d }] = await Promise.all([
        supabase.from("stores").select("id, name").order("name"),
        supabase.from("devices").select("id, imei, model, storage, color, cost_price")
          .eq("status", "in_stock").order("created_at", { ascending: false }).limit(2000),
      ]);
      setStores((s || []).filter((x) => x.id !== employee.store_id));
      setDevices(d || []);
    })();
  }, [employee.store_id]);

  const picked = devices.find((d) => d.id === deviceId);
  const cost = Number(picked?.cost_price || 0);
  const margin = (Number(price) || 0) - cost;

  const filtered = devices.filter((d) => {
    if (!search.trim()) return true;
    const q = search.trim().toLowerCase();
    return [d.imei, d.model, d.color, d.storage].some((v) => v?.toLowerCase().includes(q));
  }).slice(0, 50);

  const submit = async () => {
    if (!deviceId) { setError("Vui lòng chọn máy cần xuất."); return; }
    if (!toStore) { setError("Vui lòng chọn cửa hàng nhận."); return; }
    const p = Number(price) || 0;
    if (p <= 0) { setError("Giá xuất bán nội bộ phải lớn hơn 0."); return; }
    setSaving(true); setError("");
    const { data, error: err } = await supabase.rpc("create_internal_transfer", {
      p_device_id: deviceId, p_to_store_id: toStore,
      p_transfer_price: p, p_note: note.trim() || null,
    });
    setSaving(false);
    if (err) { setError(err.message); return; }
    await supabase.from("audit_logs").insert({
      table_name: "internal_transfers", record_id: deviceId, action: "create",
      new_data: { ma: data, den: toStore, gia: p },
      performed_by: employee.id, store_id: employee.store_id,
    });
    onSaved(data);
  };

  return (
    <Card className="p-4 sm:p-5 mb-4">
      <p className="text-sm font-medium text-slate-700 mb-1">Xuất bán nội bộ</p>
      <p className="text-xs text-slate-400 mb-3">
        Máy sẽ chuyển sang trạng thái Giữ chỗ cho tới khi cửa hàng nhận xác nhận.
        Giá xuất bán ở đây sẽ thành giá vốn của cửa hàng nhận.
      </p>

      <div className="grid sm:grid-cols-2 gap-3">
        <label className="block">
          <span className="text-xs font-medium text-slate-500 mb-1 block">Cửa hàng nhận *</span>
          <select value={toStore} onChange={(e) => setToStore(e.target.value)}
            className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm text-slate-800 outline-none focus:ring-2 focus:ring-brand-400/40 focus:border-brand-500 transition">
            <option value="">— Chọn cửa hàng —</option>
            {stores.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
        </label>
        <TextField label="Giá xuất bán nội bộ (đ) *" type="number" value={price}
          onChange={(e) => setPrice(e.target.value)} />
      </div>

      <div className="mt-3">
        <span className="text-xs font-medium text-slate-500 mb-1 block">Máy cần xuất *</span>
        <div className="relative mb-2">
          <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-300" />
          <input value={search} onChange={(e) => setSearch(e.target.value)}
            placeholder="Tìm theo IMEI, model, màu..."
            className="w-full pl-9 pr-3 py-2 rounded-lg border border-slate-300 bg-white text-sm placeholder:text-slate-400 outline-none focus:ring-2 focus:ring-brand-400/40 focus:border-brand-500 transition" />
        </div>
        <div className="border border-slate-200 rounded-xl max-h-56 overflow-y-auto">
          {filtered.length === 0 ? (
            <p className="text-xs text-slate-400 py-4 text-center">Không có máy nào Còn hàng khớp tìm kiếm.</p>
          ) : filtered.map((d) => (
            <button key={d.id} type="button" onClick={() => { setDeviceId(d.id); setError(""); }}
              className={classNames("w-full text-left px-3 py-2 border-b border-slate-50 last:border-0 hover:bg-slate-50",
                deviceId === d.id && "bg-brand-50")}>
              <p className="text-sm text-slate-700">{[d.model, d.storage, d.color].filter(Boolean).join(" ")}</p>
              <p className="text-xs text-slate-400">
                IMEI {d.imei || "—"}{canSeeCost && d.cost_price ? ` · giá vốn ${fmtVND(d.cost_price)}` : ""}
              </p>
            </button>
          ))}
        </div>
      </div>

      {picked && canSeeCost && Number(price) > 0 && (
        <div className={classNames("mt-3 rounded-xl px-3 py-2 text-xs",
          margin >= 0 ? "bg-emerald-50 text-emerald-800" : "bg-rose-50 text-rose-700")}>
          Giá vốn {fmtVND(cost)} → xuất {fmtVND(Number(price))} ·
          lãi nội bộ <span className="font-medium">{fmtVND(margin)}</span>
          {margin < 0 && " — đang xuất dưới giá vốn"}
        </div>
      )}

      <div className="mt-3">
        <TextField label="Ghi chú" value={note} onChange={(e) => setNote(e.target.value)}
          placeholder="Điều chuyển theo yêu cầu CH nhận..." />
      </div>

      {error && <p className="text-xs text-rose-600 bg-rose-50 rounded-lg px-3 py-2 mt-3">{error}</p>}

      <div className="flex gap-2 mt-3">
        <button onClick={submit} disabled={saving}
          className="bg-brand-600 hover:bg-brand-700 disabled:opacity-60 text-white rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-2">
          {saving && <Loader2 size={15} className="animate-spin" />} Lập phiếu xuất
        </button>
        <button onClick={onCancel} className="px-4 rounded-xl border border-slate-200 text-sm text-slate-500 hover:bg-slate-50">Hủy</button>
      </div>
    </Card>
  );
}

function InternalTransferTab({ employee }) {
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [showForm, setShowForm] = useState(false);
  const [okCode, setOkCode] = useState(null);

  const canSeeCost = employee.role !== "nhan_vien";
  const canCancel = employee.role !== "nhan_vien";

  const load = useCallback(async () => {
    setLoading(true);
    const { data } = await supabase.from("v_internal_transfers").select("*")
      .order("created_at", { ascending: false }).limit(500);
    setRows(data || []);
    setLoading(false);
  }, []);

  useEffect(() => { load(); }, [load]);

  const outRows = rows.filter((r) => r.from_store_id === employee.store_id);
  const pendingOut = outRows.filter((r) => r.status === "pending");
  const receivedOut = outRows.filter((r) => r.status === "received");
  const internalRevenue = receivedOut.reduce((s, r) => s + Number(r.transfer_price || 0), 0);
  const internalMargin = receivedOut.reduce((s, r) => s + Number(r.internal_margin || 0), 0);

  const cancel = async (r) => {
    const reason = prompt(`Hủy phiếu ${r.transfer_code}? Nhập lý do:`);
    if (reason === null) return;
    const { error } = await supabase.rpc("cancel_internal_transfer", {
      p_transfer_id: r.id, p_reason: reason || null,
    });
    if (error) { alert(error.message); return; }
    load();
  };

  return (
    <div className="space-y-3">
      <div className="grid grid-cols-2 lg:grid-cols-3 gap-3">
        <Card className="p-4">
          <p className="text-xs text-slate-400 mb-1">Doanh thu nội bộ đã xuất</p>
          <p className="text-lg font-semibold text-slate-800">{fmtVND(internalRevenue)}</p>
          <p className="text-[11px] text-slate-400">{receivedOut.length} máy đã được nhận</p>
        </Card>
        {canSeeCost && (
          <Card className="p-4">
            <p className="text-xs text-slate-400 mb-1">Lãi nội bộ</p>
            <p className={classNames("text-lg font-semibold", internalMargin >= 0 ? "text-emerald-700" : "text-rose-600")}>
              {fmtVND(internalMargin)}
            </p>
          </Card>
        )}
        <Card className="p-4">
          <p className="text-xs text-slate-400 mb-1">Đang chờ bên kia nhận</p>
          <p className="text-lg font-semibold text-amber-600">{pendingOut.length} máy</p>
        </Card>
      </div>

      {!showForm && (
        <button onClick={() => { setShowForm(true); setOkCode(null); }}
          className="bg-brand-600 hover:bg-brand-700 text-white rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-1.5">
          <ArrowLeftRight size={15} /> Xuất bán nội bộ
        </button>
      )}

      {okCode && (
        <p className="text-xs text-emerald-700 bg-emerald-50 rounded-lg px-3 py-2">
          Đã lập phiếu <span className="font-semibold">{okCode}</span>. Chờ cửa hàng nhận xác nhận trong mục Kho hàng của họ.
        </p>
      )}

      {showForm && (
        <InternalTransferForm employee={employee}
          onCancel={() => setShowForm(false)}
          onSaved={(code) => { setShowForm(false); setOkCode(code); load(); }} />
      )}

      <Card className="p-0 overflow-hidden">
        {loading ? (
          <div className="flex justify-center py-10"><Loader2 className="animate-spin text-slate-300" /></div>
        ) : rows.length === 0 ? (
          <EmptyState icon={ArrowLeftRight} text="Chưa có phiếu xuất bán nội bộ nào." />
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead><tr className="text-left text-xs text-slate-400 border-b border-slate-100">
                <th className="px-3 py-2">Mã</th>
                <th className="px-3 py-2">Ngày</th>
                <th className="px-3 py-2">Chiều</th>
                <th className="px-3 py-2">Máy</th>
                <th className="px-3 py-2 text-right">Giá xuất</th>
                <th className="px-3 py-2">Trạng thái</th>
                <th className="px-3 py-2"></th>
              </tr></thead>
              <tbody>
                {rows.map((r) => {
                  const isOut = r.from_store_id === employee.store_id;
                  return (
                    <tr key={r.id} className="border-b border-slate-50 last:border-0">
                      <td className="px-3 py-2.5 text-slate-400 doc-code whitespace-nowrap">{r.transfer_code}</td>
                      <td className="px-3 py-2.5 text-slate-500 whitespace-nowrap">{fmtDate(r.created_at)}</td>
                      <td className="px-3 py-2.5 whitespace-nowrap">
                        <span className={classNames("text-xs px-2 py-0.5 rounded",
                          isOut ? "bg-sky-50 text-sky-700" : "bg-violet-50 text-violet-700")}>
                          {isOut ? `Xuất → ${r.to_store_name}` : `Nhận ← ${r.from_store_name}`}
                        </span>
                      </td>
                      <td className="px-3 py-2.5">
                        <p className="text-slate-700">{[r.model, r.storage, r.color].filter(Boolean).join(" ")}</p>
                        <p className="text-xs text-slate-400">IMEI {r.imei || "—"}</p>
                      </td>
                      <td className="px-3 py-2.5 text-right whitespace-nowrap">
                        <p className="font-medium text-slate-700">{fmtVND(r.transfer_price)}</p>
                        {canSeeCost && isOut && (
                          <p className={classNames("text-xs", Number(r.internal_margin) >= 0 ? "text-emerald-600" : "text-rose-500")}>
                            lãi {fmtVND(r.internal_margin)}
                          </p>
                        )}
                      </td>
                      <td className="px-3 py-2.5 whitespace-nowrap">
                        <span className={classNames("text-xs px-2 py-0.5 rounded-full", TRANSFER_STATUS_STYLES[r.status])}>
                          {TRANSFER_STATUS_LABELS[r.status]}
                        </span>
                      </td>
                      <td className="px-3 py-2.5 text-right whitespace-nowrap">
                        {canCancel && r.status === "pending" && (
                          <button onClick={() => cancel(r)} className="text-xs text-rose-500 hover:underline">Hủy phiếu</button>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </Card>
    </div>
  );
}

function ReceiveTransferModal({ row, employee, onClose, onDone }) {
  const [salePrice, setSalePrice] = useState("");
  const [note, setNote] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const [okCode, setOkCode] = useState(null);

  const submit = async () => {
    setSaving(true); setError("");
    const { data, error: err } = await supabase.rpc("receive_internal_transfer", {
      p_transfer_id: row.id,
      p_sale_price: salePrice === "" ? null : Number(salePrice),
      p_note: note.trim() || null,
    });
    setSaving(false);
    if (err) { setError(err.message); return; }
    setOkCode(data);
    await supabase.from("audit_logs").insert({
      table_name: "internal_transfers", record_id: row.id, action: "update",
      new_data: { nhan: data }, performed_by: employee.id, store_id: employee.store_id,
    });
  };

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/40 flex items-center justify-center p-4">
      <Card className="w-full max-w-md p-5">
        <div className="flex items-center justify-between mb-4">
          <p className="font-semibold text-slate-800 text-sm">Nhận hàng nội bộ — {row.transfer_code}</p>
          <button onClick={onClose} className="text-slate-400 hover:text-rose-600"><X size={16} /></button>
        </div>

        {okCode ? (
          <div className="space-y-3">
            <div className="bg-emerald-50 border border-emerald-200 rounded-xl px-3 py-3 text-sm text-emerald-800 space-y-1">
              <p>Đã nhận máy vào kho theo phiếu <span className="font-semibold">{okCode}</span>.</p>
              <p className="text-xs">Giá vốn ghi nhận {fmtVND(row.transfer_price)}. Công nợ nội bộ với {row.from_store_name} đã được treo.</p>
            </div>
            <button onClick={onDone} className="w-full bg-brand-600 hover:bg-brand-700 text-white rounded-xl py-2.5 text-sm font-medium">Xong</button>
          </div>
        ) : (
          <div className="space-y-3">
            <div className="bg-slate-50 rounded-xl px-3 py-2.5 text-xs space-y-0.5">
              <p className="font-medium text-slate-700">{[row.model, row.storage, row.color].filter(Boolean).join(" ")}</p>
              <p className="text-slate-400">IMEI {row.imei || "—"}</p>
              <p className="text-slate-400">Từ {row.from_store_name} · lập ngày {fmtDate(row.created_at)}</p>
              {row.note && <p className="text-slate-400">Ghi chú: {row.note}</p>}
              <div className="flex justify-between pt-1 mt-1 border-t border-slate-200">
                <span className="text-slate-500">Giá vốn khi vào kho</span>
                <span className="font-semibold text-slate-800">{fmtVND(row.transfer_price)}</span>
              </div>
            </div>
            <TextField label="Giá bán dự kiến (đ)" type="number" value={salePrice}
              onChange={(e) => setSalePrice(e.target.value)} placeholder="Để trống nếu chưa định giá" />
            <TextField label="Ghi chú khi nhận" value={note} onChange={(e) => setNote(e.target.value)} />
            <p className="text-[11px] text-slate-400">
              Xác nhận sẽ đưa máy vào kho và treo khoản phải trả {fmtVND(row.transfer_price)} với {row.from_store_name}.
            </p>
            {error && <p className="text-xs text-rose-600 bg-rose-50 rounded-lg px-3 py-2">{error}</p>}
            <div className="flex gap-2">
              <button onClick={submit} disabled={saving}
                className="flex-1 bg-brand-600 hover:bg-brand-700 disabled:opacity-60 text-white rounded-xl py-2.5 text-sm font-medium flex items-center justify-center gap-2">
                {saving && <Loader2 size={15} className="animate-spin" />} Xác nhận nhận hàng
              </button>
              <button onClick={onClose} className="px-4 rounded-xl border border-slate-200 text-sm text-slate-500 hover:bg-slate-50">Hủy</button>
            </div>
          </div>
        )}
      </Card>
    </div>
  );
}

function OrdersModule({ employee }) {
  const [tab, setTab] = useState("orders");
  const [orders, setOrders] = useState([]);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState("");
  const [showForm, setShowForm] = useState(false);
  const [showExport, setShowExport] = useState(false);

  const canCreate = employee.role !== "ke_toan";

  const load = useCallback(async () => {
    setLoading(true);
    const { data, error } = await supabase.from("v_sales_order_debt").select("*").order("created_at", { ascending: false }).limit(1000);
    if (!error) setOrders(data || []);
    setLoading(false);
  }, []);

  useEffect(() => { load(); }, [load]);

  const filtered = orders.filter((o) => {
    if (!search.trim()) return true;
    return o.order_code?.toLowerCase().includes(search.trim().toLowerCase());
  });

  const debtOrders = orders.filter(
    (o) => o.status !== "cancelled" && Number(o.customer_debt || 0) > 0
  );
  const totalOrderDebt = debtOrders.reduce((s, o) => s + Number(o.customer_debt || 0), 0);
  const debtOrderCount = debtOrders.length;
  const today = new Date(new Date().toDateString());
  const overdueCount = debtOrders.filter((o) => o.due_date && new Date(o.due_date) < today).length;

  const shopDebtOrders = orders.filter((o) => Number(o.shop_debt || 0) > 0);
  const totalShopDebt = shopDebtOrders.reduce((s, o) => s + Number(o.shop_debt || 0), 0);

  return (
    <div>
      <div className="flex items-center justify-between mb-4 gap-2 flex-wrap">
        <div>
          <h2 className="text-lg font-semibold text-slate-800">Đơn hàng bán</h2>
          <p className="text-xs text-slate-400">
            {tab === "orders" ? `${orders.length} đơn hàng` : "Điều chuyển máy giữa các cửa hàng trong hệ thống"}
          </p>
        </div>
        {tab === "orders" && (
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
        )}
      </div>

      <div className="flex gap-2 mb-4">
        {[["orders", "Đơn bán khách"], ["internal", "Xuất bán nội bộ"]].map(([k, label]) => (
          <button key={k} onClick={() => setTab(k)}
            className={classNames("px-4 py-2 rounded-xl text-sm border font-medium",
              tab === k ? "bg-brand-600 text-white border-brand-600" : "bg-white text-slate-500 border-slate-200 hover:bg-slate-50")}
          >{label}</button>
        ))}
      </div>

      {tab === "internal" && <InternalTransferTab employee={employee} />}
      {tab === "orders" && (<>

      {totalOrderDebt > 0 && (
        <div className="bg-amber-50 border border-amber-200 rounded-xl px-3 py-2.5 mb-4 text-xs text-amber-800 flex items-center gap-2">
          <Banknote size={15} className="shrink-0" />
          Khách đang nợ tổng cộng <span className="font-medium">{fmtVND(totalOrderDebt)}</span> trên {debtOrderCount} đơn
          {overdueCount > 0 && <span className="text-rose-600 font-medium">· {overdueCount} đơn quá hạn</span>}
        </div>
      )}

      {totalShopDebt > 0 && (
        <div className="bg-indigo-50 border border-indigo-200 rounded-xl px-3 py-2.5 mb-4 text-xs text-indigo-800 flex items-center gap-2">
          <Banknote size={15} className="shrink-0" />
          Cửa hàng đang nợ khách <span className="font-medium">{fmtVND(totalShopDebt)}</span> trên {shopDebtOrders.length} đơn
          <span className="text-slate-400">· phần chênh máy khách đổi</span>
        </div>
      )}

      {employee.role === "quan_ly" && orders.some((o) => o.status === "pending_stock") && (
        <div className="bg-amber-50 border border-amber-200 rounded-xl px-3 py-2.5 mb-4 text-xs text-amber-800 flex items-center gap-2">
          <ShieldAlert size={15} className="shrink-0" />
          Có <span className="font-medium">{orders.filter((o) => o.status === "pending_stock").length} đơn</span> đang chờ đối soát kho (IMEI nhân viên nhập tay lúc bán không khớp Kho hàng) — mở đơn để đối soát.
        </div>
      )}

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
              className="w-full pl-9 pr-3 py-2 rounded-lg border border-slate-300 bg-white text-sm placeholder:text-slate-400 outline-none focus:ring-2 focus:ring-brand-400/40 focus:border-brand-500 transition"
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
                <th className="px-3 py-2">Công nợ</th>
                <th className="px-3 py-2">Trạng thái</th>
                <th className="px-3 py-2"></th>
              </tr></thead>
              <tbody>
                {filtered.map((o) => <OrderRow key={o.id} order={o} employee={employee} onDeleted={load} onReconciled={load} />)}
              </tbody>
            </table>
          </div>
        )}
      </Card>
      </>)}
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Purchase module — nhập máy/thu cũ độc lập (mua trả tiền khách)         */
/* ---------------------------------------------------------------------- */

function SupplierPicker({ value, onSelect, employee }) {
  const [query, setQuery] = useState("");
  const [results, setResults] = useState([]);
  const [open, setOpen] = useState(false);
  const [loading, setLoading] = useState(false);
  const [showNew, setShowNew] = useState(false);
  const [newName, setNewName] = useState("");
  const [newPhone, setNewPhone] = useState("");
  const [linkCustomer, setLinkCustomer] = useState(null);
  const [creating, setCreating] = useState(false);

  useEffect(() => {
    let active = true;
    setLoading(true);
    const t = setTimeout(async () => {
      let req = supabase.from("suppliers").select("*").order("created_at", { ascending: false }).limit(8);
      if (query.trim()) req = req.ilike("name", `%${query.trim()}%`);
      const { data } = await req;
      if (active) { setResults(data || []); setLoading(false); }
    }, 250);
    return () => { active = false; clearTimeout(t); };
  }, [query]);

  const createNew = async () => {
    if (!newName.trim()) return;
    setCreating(true);
    const { data, error } = await supabase.from("suppliers").insert({
      name: newName.trim(), phone: newPhone.trim() || null, customer_id: linkCustomer?.id || null,
      store_id: employee.store_id, created_by: employee.id,
    }).select().maybeSingle();
    setCreating(false);
    if (error) { alert(error.message); return; }
    onSelect(data);
    setShowNew(false);
    setNewName(""); setNewPhone(""); setLinkCustomer(null);
  };

  if (value) {
    return (
      <div className="flex items-center justify-between bg-slate-50 rounded-xl px-3 py-2.5">
        <div>
          <p className="text-sm font-medium text-slate-700">{value.name}</p>
          <p className="text-xs text-slate-400">{value.phone || "—"}</p>
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
          placeholder="Tìm nhà cung cấp theo tên..."
          className="w-full pl-9 pr-3 py-2 rounded-lg border border-slate-300 bg-white text-sm placeholder:text-slate-400 outline-none focus:ring-2 focus:ring-brand-400/40 focus:border-brand-500 transition"
        />
      </div>
      {open && (
        <div className="absolute z-10 w-full bg-white border border-slate-100 rounded-xl shadow-lg mt-1 max-h-72 overflow-y-auto">
          {loading ? (
            <div className="p-3 text-xs text-slate-400 flex items-center gap-2"><Loader2 size={13} className="animate-spin" /> Đang tìm...</div>
          ) : results.length === 0 && !showNew ? (
            <div className="p-3 text-xs text-slate-400">Chưa có NCC nào khớp.</div>
          ) : (
            results.map((s) => (
              <button
                type="button" key={s.id}
                onClick={() => { onSelect(s); setOpen(false); setQuery(""); }}
                className="w-full text-left px-3 py-2 hover:bg-slate-50 text-sm border-b border-slate-50 last:border-0"
              >
                <p className="font-medium text-slate-700">{s.name}</p>
                <p className="text-xs text-slate-400">{s.phone || "—"}</p>
              </button>
            ))
          )}
          {!showNew ? (
            <button type="button" onClick={() => { setShowNew(true); setNewName(query); }} className="w-full text-left px-3 py-2.5 text-xs text-brand-600 hover:bg-slate-50 border-t border-slate-100 flex items-center gap-1">
              <Plus size={13} /> Thêm nhà cung cấp mới
            </button>
          ) : (
            <div className="p-3 border-t border-slate-100 space-y-2">
              <input value={newName} onChange={(e) => setNewName(e.target.value)} placeholder="Tên NCC *" className="w-full rounded-lg border border-slate-200 px-2 py-1.5 text-xs" />
              <input value={newPhone} onChange={(e) => setNewPhone(e.target.value)} placeholder="SĐT" className="w-full rounded-lg border border-slate-200 px-2 py-1.5 text-xs" />
              <div className="pt-1">
                <p className="text-[11px] text-slate-400 mb-1">
                  NCC này cũng mua máy của cửa hàng? Liên kết hồ sơ khách hàng để cấn trừ công nợ được.
                </p>
                <CustomerPicker value={linkCustomer} onSelect={setLinkCustomer} />
              </div>
              <div className="flex gap-2">
                <button type="button" onClick={createNew} disabled={creating || !newName.trim()} className="bg-brand-600 text-white text-xs px-3 py-1.5 rounded-lg disabled:opacity-50">
                  {creating ? "Đang lưu..." : "Lưu NCC"}
                </button>
                <button type="button" onClick={() => setShowNew(false)} className="text-xs text-slate-400 px-3 py-1.5">Hủy</button>
              </div>
            </div>
          )}
        </div>
      )}
    </div>
  );
}

function PurchaseForm({ onCancel, onSaved, employee }) {
  const { banks } = usePaymentOptions();
  const [sourceType, setSourceType] = useState("customer"); // "customer" | "supplier"
  const [customer, setCustomer] = useState(null);
  const [supplier, setSupplier] = useState(null);
  const [form, setForm] = useState({
    imei: "", model: "", storage: "", color: "", condition: "used", condition_percent: "",
    purchase_price: "", payment_method: "bank_transfer", bank_account_id: null, notes: "",
    paid_amount: "", due_date: "",
  });
  const [duplicateImei, setDuplicateImei] = useState(null);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const [partnerRec, setPartnerRec] = useState(0);   // đối tác đang nợ cửa hàng

  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));

  // Người bán máy này có đang nợ cửa hàng không?
  useEffect(() => {
    let active = true;
    const ref = sourceType === "customer" ? customer : supplier;
    if (!ref) { setPartnerRec(0); return; }
    (async () => {
      const { data } = sourceType === "customer"
        ? await supabase.rpc("partner_balance_by_customer", { p_customer_id: ref.id })
        : await supabase.rpc("partner_balance_by_supplier", { p_supplier_id: ref.id });
      if (!active) return;
      const bal = Array.isArray(data) ? data[0] : data;
      setPartnerRec(Math.max(0, Number(bal?.receivable) || 0));
    })();
    return () => { active = false; };
  }, [sourceType, customer, supplier]);

  const checkImei = async (imei) => {
    if (!imei.trim()) { setDuplicateImei(null); return; }
    const { data } = await supabase.from("devices").select("*").eq("imei", imei.trim()).maybeSingle();
    setDuplicateImei(data || null);
  };

  const price = Number(form.purchase_price) || 0;
  const paid = form.paid_amount === "" ? price : Number(form.paid_amount);
  const debt = Math.max(0, price - paid);

  const submit = async (e) => {
    e.preventDefault();
    setError("");
    if (sourceType === "customer" && !customer) { setError("Vui lòng chọn khách hàng bán máy."); return; }
    if (sourceType === "supplier" && !supplier) { setError("Vui lòng chọn nhà cung cấp."); return; }
    if (!form.model.trim()) { setError("Vui lòng nhập tên model máy."); return; }
    if (!form.purchase_price || Number(form.purchase_price) <= 0) { setError("Vui lòng nhập giá thu mua hợp lệ."); return; }
    if (form.imei.trim() && duplicateImei) { setError(`IMEI ${duplicateImei.imei} đã có trong kho, không thể thu mua trùng.`); return; }
    if (sourceType === "customer") {
      if (!customer.cccd) { setError("Khách hàng chưa có số CCCD — vào mục Khách hàng bổ sung CCCD trước khi lập hồ sơ thu mua (bắt buộc để hồ sơ hợp lệ)."); return; }
    }
    if (sourceType === "supplier" && paid > price) { setError("Số tiền đã thanh toán không được lớn hơn giá thu mua."); return; }

    setSaving(true);
    try {
      const { data: newDevice, error: devErr } = await supabase.from("devices").insert({
        imei: form.imei.trim() || null, model: form.model.trim(),
        storage: form.storage.trim() || null, color: form.color.trim() || null,
        condition: form.condition, condition_percent: form.condition_percent === "" ? null : Number(form.condition_percent),
        status: "in_stock", cost_price: Number(form.purchase_price),
        supplier: sourceType === "supplier" ? supplier.name : `Thu mua từ khách — ${customer.full_name}`,
        import_date: new Date().toISOString().slice(0, 10),
        created_by: employee.id, updated_by: employee.id, store_id: employee.store_id,
      }).select().maybeSingle();
      if (devErr) throw devErr;

      const debtStatus = sourceType === "supplier" ? (debt === 0 ? "paid" : (paid === 0 ? "unpaid" : "partial")) : "paid";

      const { data: purchase, error: poErr } = await supabase.from("purchase_orders").insert({
        source_type: sourceType,
        customer_id: sourceType === "customer" ? customer.id : null,
        supplier_id: sourceType === "supplier" ? supplier.id : null,
        device_id: newDevice.id, linked_sale_order_id: null,
        purchase_price: price, payment_method: form.payment_method,
        bank_account_id: form.payment_method === "bank_transfer" ? (form.bank_account_id || null) : null,
        notes: form.notes.trim() || null, created_by: employee.id,
        store_id: employee.store_id,
        paid_amount: form.payment_method === "debt_offset" ? 0
          : (sourceType === "supplier" ? paid : price),
        debt_status: debtStatus,
        due_date: sourceType === "supplier" && debt > 0 ? (form.due_date || null) : null,
      }).select().maybeSingle();
      if (poErr) throw poErr;

      if (sourceType === "customer") {
        const { error: contractErr } = await supabase.from("contracts").insert({ purchase_order_id: purchase.id, created_by: employee.id, store_id: employee.store_id });
        if (contractErr) throw contractErr;
      }

      await supabase.from("audit_logs").insert([
        { table_name: "devices", record_id: newDevice.id, action: "create", new_data: newDevice, performed_by: employee.id, store_id: employee.store_id },
        { table_name: "purchase_orders", record_id: purchase.id, action: "create", new_data: purchase, performed_by: employee.id, store_id: employee.store_id },
      ]);

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
      <div className="flex items-center justify-between mb-4">
        <p className="font-semibold text-slate-800 text-sm">Nhập máy mới</p>
        <button onClick={onCancel} className="text-slate-400 hover:text-rose-600"><X size={16} /></button>
      </div>

      <div className="flex gap-2 mb-4">
        {[
          { key: "customer", label: "Khách lẻ", icon: Users },
          { key: "supplier", label: "Nhà cung cấp", icon: Building2 },
        ].map((opt) => (
          <button
            key={opt.key} type="button"
            onClick={() => { setSourceType(opt.key); setCustomer(null); setSupplier(null); setError(""); }}
            className={classNames(
              "flex-1 flex items-center justify-center gap-1.5 rounded-xl px-3 py-2 text-sm font-medium border",
              sourceType === opt.key ? "bg-brand-600 text-white border-brand-600" : "bg-white text-slate-500 border-slate-200"
            )}
          >
            <opt.icon size={14} /> {opt.label}
          </button>
        ))}
      </div>

      <form onSubmit={submit} className="space-y-4">
        {sourceType === "customer" ? (
          <div>
            <span className="text-xs font-medium text-slate-500 mb-1 block">Khách hàng (bên bán máy) *</span>
            <CustomerPicker value={customer} onSelect={setCustomer} />
            {customer && !customer.cccd && (
              <p className="text-xs text-amber-600 mt-1.5">⚠ Khách này chưa có CCCD trong hồ sơ — cần vào mục Khách hàng bổ sung trước.</p>
            )}
          </div>
        ) : (
          <div>
            <span className="text-xs font-medium text-slate-500 mb-1 block">Nhà cung cấp *</span>
            <SupplierPicker value={supplier} onSelect={setSupplier} employee={employee} />
          </div>
        )}

        {duplicateImei && (
          <div className="bg-rose-50 border border-rose-200 rounded-xl px-3 py-2.5 text-xs text-rose-700 flex items-center gap-2">
            <ShieldAlert size={15} className="shrink-0" />
            IMEI <span className="font-medium">{duplicateImei.imei}</span> đã tồn tại trong kho ({DEVICE_STATUS_LABELS[duplicateImei.status]}) — không thể lưu trùng.
          </div>
        )}

        <div className="grid sm:grid-cols-2 gap-3">
          <TextField label="Số IMEI (để trống nếu chưa có)" value={form.imei} onChange={(e) => { set("imei")(e); checkImei(e.target.value); }} />
          <label className="block">
            <span className="text-xs font-medium text-slate-500 mb-1 block">Model máy *</span>
            <ModelPicker value={form.model} onSelect={(v) => setForm((f) => ({ ...f, model: v }))} placeholder="iPhone 13 128GB" />
          </label>
          <TextField label="Dung lượng" value={form.storage} onChange={set("storage")} placeholder="128GB" list="dl-storage" />
          <TextField label="Màu sắc" value={form.color} onChange={set("color")} list="dl-colors-purchase" />
          <datalist id="dl-colors-purchase">{coloroptionsForModel(form.model).map((c) => <option key={c} value={c} />)}</datalist>
          <datalist id="dl-storage"><option value="64GB" /><option value="128GB" /><option value="256GB" /><option value="512GB" /><option value="1TB" /></datalist>
          <label className="block">
            <span className="text-xs font-medium text-slate-500 mb-1 block">Tình trạng</span>
            <select value={form.condition} onChange={set("condition")} className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm text-slate-800 outline-none focus:ring-2 focus:ring-brand-400/40 focus:border-brand-500 transition">
              <option value="used">Máy cũ</option>
              <option value="new">Máy mới</option>
            </select>
          </label>
          <TextField label="Độ mới (%)" type="number" min="0" max="100" value={form.condition_percent} onChange={set("condition_percent")} placeholder="99" />
          <TextField label="Giá thu mua (đ) *" type="number" value={form.purchase_price} onChange={set("purchase_price")} />
        </div>

        <label className="block">
          <span className="text-xs font-medium text-slate-500 mb-1 block">Hình thức chi trả</span>
          <select value={form.payment_method} onChange={set("payment_method")} className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm text-slate-800 outline-none focus:ring-2 focus:ring-brand-400/40 focus:border-brand-500 transition">
            <option value="bank_transfer">Chuyển khoản</option>
            <option value="debt_offset">Bù trừ công nợ (không chi tiền)</option>
          </select>
          {form.payment_method === "debt_offset" && (
            <div className="mt-2 bg-indigo-50 border border-indigo-200 rounded-xl px-3 py-2.5 text-xs text-indigo-800 space-y-1">
              {partnerRec > 0 ? (
                <>
                  <p>
                    Người bán đang nợ cửa hàng <span className="font-medium">{fmtVND(partnerRec)}</span>.
                    Tiền mua máy trừ thẳng vào khoản đó, không chi tiền.
                  </p>
                  {price > 0 && (
                    <p>
                      Trừ được <span className="font-medium">{fmtVND(Math.min(price, partnerRec))}</span>
                      {price > partnerRec && <> — phần còn lại <span className="font-medium">{fmtVND(price - partnerRec)}</span> thành khoản cửa hàng nợ họ, sẽ tự bù trừ khi họ mua máy.</>}
                      {price <= partnerRec && <> — sau đó họ còn nợ <span className="font-medium">{fmtVND(partnerRec - price)}</span>.</>}
                    </p>
                  )}
                  <p className="text-indigo-500">Hệ thống lập biên bản bù trừ mã BT, phân bổ vào các đơn bán cũ nhất trước.</p>
                </>
              ) : (
                <>
                  <p>
                    Không chi tiền cho người bán.
                    {price > 0 && <> Toàn bộ <span className="font-medium">{fmtVND(price)}</span> ghi thành khoản cửa hàng nợ họ.</>}
                  </p>
                  <p className="text-indigo-500">
                    Khi khách mua máy mới, chọn "Bù trừ công nợ" ở phần thanh toán để trừ khoản này vào tiền máy — đúng luồng lên đời máy.
                  </p>
                </>
              )}
            </div>
          )}
          {partnerRec > 0 && form.payment_method !== "debt_offset" && (
            <p className="mt-1.5 text-[11px] text-amber-600">
              Người bán đang nợ cửa hàng {fmtVND(partnerRec)} — cân nhắc chọn "Bù trừ công nợ" thay vì chi tiền.
            </p>
          )}
          {form.payment_method === "bank_transfer" && (
            <div className="mt-2">
              <BankSelect
                banks={banks}
                value={form.bank_account_id}
                onChange={(v) => setForm((f) => ({ ...f, bank_account_id: v }))}
                className="w-full !text-sm !px-3 !py-2 !rounded-xl"
              />
            </div>
          )}
        </label>

        {sourceType === "supplier" && form.payment_method !== "debt_offset" && (
          <div className="bg-slate-50 rounded-xl p-3 space-y-3">
            <p className="text-xs font-medium text-slate-600">Công nợ NCC</p>
            <div className="grid sm:grid-cols-2 gap-3">
              <TextField label="Đã thanh toán (đ)" type="number" value={form.paid_amount} onChange={set("paid_amount")} placeholder={String(price)} />
              {debt > 0 && <TextField label="Hạn thanh toán" type="date" value={form.due_date} onChange={set("due_date")} />}
            </div>
            <div className={classNames("text-xs px-3 py-2 rounded-lg", debt === 0 ? "bg-emerald-50 text-emerald-700" : "bg-amber-50 text-amber-700")}>
              {debt === 0 ? "Đã thanh toán đủ" : `Còn nợ NCC: ${fmtVND(debt)}`}
            </div>
          </div>
        )}

        <TextField label="Ghi chú" value={form.notes} onChange={set("notes")} />

        {error && <p className="text-xs text-rose-600 bg-rose-50 rounded-lg px-3 py-2">{error}</p>}
        <div className="flex gap-2">
          <button type="submit" disabled={saving} className="bg-brand-600 hover:bg-brand-700 text-white rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-2 disabled:opacity-60">
            {saving && <Loader2 size={14} className="animate-spin" />} Xác nhận nhập máy
          </button>
          <button type="button" onClick={onCancel} className="text-slate-500 text-sm px-4 py-2">Hủy</button>
        </div>
      </form>
    </Card>
  );
}


const IMPORT_TEMPLATE_COLS = [
  "IMEI", "Model", "Dung lượng", "Màu sắc", "Tình trạng",
  "Độ mới (%)", "Giá thu mua", "Ghi chú",
];
const IMPORT_CONDITIONS = { "Máy mới": "new", "Máy cũ": "used" };
const IMPORT_STORAGES = ["64GB", "128GB", "256GB", "512GB", "1TB"];

function ImportPurchaseModal({ employee, supplier, onClose, onDone }) {
  const [rows, setRows] = useState(null);
  const [fileName, setFileName] = useState("");
  const [saving, setSaving] = useState(false);
  const [progress, setProgress] = useState(0);
  const [error, setError] = useState("");
  const [result, setResult] = useState(null);

  const readFile = async (file) => {
    setError(""); setResult(null); setFileName(file.name);
    try {
      const buf = await file.arrayBuffer();
      const wb = XLSX.read(buf, { type: "array" });
      const ws = wb.Sheets["NhapMay"] || wb.Sheets[wb.SheetNames[0]];
      // Biểu mẫu có dòng hướng dẫn ở đầu, tiêu đề nằm ở dòng 3
      let raw = XLSX.utils.sheet_to_json(ws, { range: 2, defval: "" });
      // Bỏ tiêu đề có dấu * và khoảng trắng thừa
      raw = raw.map((row) => {
        const o = {};
        for (const k of Object.keys(row)) o[k.replace(/\s*\*\s*$/, "").trim()] = row[k];
        return o;
      });
      // Bỏ dòng ví dụ mẫu nếu người dùng quên xóa
      raw = raw.filter((row) => String(row["IMEI"] ?? "").trim() !== "356789012345678");
      // Bỏ dòng trống hoàn toàn
      raw = raw.filter((row) => Object.values(row).some((v) => String(v ?? "").trim() !== ""));
      if (raw.length === 0) { setError("File không có dòng dữ liệu nào. Hãy điền từ dòng 4 trở xuống."); return; }

      const parsed = raw.map((r, idx) => {
        const model = String(r["Model"] ?? "").trim();
        const storage = String(r["Dung lượng"] ?? "").trim();
        const color = String(r["Màu sắc"] ?? "").trim();
        const condText = String(r["Tình trạng"] ?? "Máy cũ").trim();
        const price = Number(String(r["Giá thu mua"] ?? "").replace(/[^\d.-]/g, "")) || 0;
        const pct = String(r["Độ mới (%)"] ?? "").trim();
        const imei = String(r["IMEI"] ?? "").trim();

        const errs = [];
        if (!model) errs.push("thiếu Model");
        else if (!IPHONE_MODEL_LIST.includes(model)) errs.push(`Model "${model}" không có trong danh mục`);
        if (storage && !IMPORT_STORAGES.includes(storage)) errs.push(`Dung lượng "${storage}" không hợp lệ`);
        if (color && model && IPHONE_MODEL_LIST.includes(model)) {
          const valid = coloroptionsForModel(model);
          if (valid.length > 0 && !valid.includes(color)) errs.push(`Màu "${color}" không thuộc ${model}`);
        }
        if (!IMPORT_CONDITIONS[condText]) errs.push(`Tình trạng "${condText}" phải là Máy mới hoặc Máy cũ`);
        if (price <= 0) errs.push("Giá thu mua phải lớn hơn 0");
        if (pct !== "" && (Number(pct) < 0 || Number(pct) > 100)) errs.push("Độ mới phải từ 0 đến 100");

        return {
          line: idx + 4, imei, model, storage, color,
          condition: IMPORT_CONDITIONS[condText] || "used",
          condition_percent: pct === "" ? null : Number(pct),
          price, notes: String(r["Ghi chú"] ?? "").trim(),
          errors: errs,
        };
      });

      // IMEI trùng nhau trong file
      const seen = {};
      for (const p of parsed) {
        if (!p.imei) continue;
        if (seen[p.imei]) p.errors.push(`IMEI trùng với dòng ${seen[p.imei]}`);
        else seen[p.imei] = p.line;
      }

      // IMEI đã có trong kho
      const imeis = parsed.map((p) => p.imei).filter(Boolean);
      if (imeis.length > 0) {
        const { data: dup } = await supabase.from("devices").select("imei").in("imei", imeis);
        const dupSet = new Set((dup || []).map((x) => x.imei));
        parsed.forEach((p) => { if (p.imei && dupSet.has(p.imei)) p.errors.push("IMEI đã có trong kho"); });
      }

      setRows(parsed);
    } catch (e) {
      setError("Không đọc được file. Hãy dùng đúng file mẫu (.xlsx). " + e.message);
    }
  };

  const okRows = (rows || []).filter((r) => r.errors.length === 0);
  const badRows = (rows || []).filter((r) => r.errors.length > 0);
  const totalCost = okRows.reduce((s, r) => s + r.price, 0);

  const doImport = async () => {
    if (okRows.length === 0) return;
    setSaving(true); setError(""); setProgress(0);

    const payload = okRows.map((r) => ({
      line: r.line, imei: r.imei || null, model: r.model,
      storage: r.storage || null, color: r.color || null,
      condition: r.condition,
      condition_percent: r.condition_percent === null ? "" : String(r.condition_percent),
      price: r.price, notes: r.notes || null,
    }));

    // Gọi 1 lần, database xử lý từng dòng độc lập — dòng lỗi không chặn dòng khác
    const { data, error: err } = await supabase.rpc("import_purchase_batch", {
      p_supplier_id: supplier.id, p_rows: payload,
    });
    setProgress(100);
    setSaving(false);

    if (err) { setError(err.message); return; }
    const failed = (data?.errors || []).map((e) => ({ line: e.line, msg: e.message }));
    setResult({ done: data?.inserted || 0, failed });
    if ((data?.inserted || 0) > 0) onDone();
  };

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/40 flex items-center justify-center p-4 overflow-y-auto">
      <Card className="w-full max-w-3xl p-5 my-8">
        <div className="flex items-center justify-between mb-4">
          <p className="font-semibold text-slate-800 text-sm">Nhập máy hàng loạt từ Excel</p>
          <button onClick={onClose} className="text-slate-400 hover:text-rose-600"><X size={16} /></button>
        </div>

        <div className="bg-slate-50 rounded-xl px-3 py-2.5 text-xs text-slate-600 mb-3 space-y-1">
          <p>
            Nhập cho nhà cung cấp <span className="font-medium text-slate-700">{supplier?.name}</span>.
            Toàn bộ máy sẽ vào kho ở trạng thái Còn hàng, phiếu nhập ghi nhận công nợ phải trả (chưa thanh toán).
          </p>
          <p className="text-slate-400">
            Cột Model, Dung lượng, Tình trạng phải khớp danh mục — xem sheet "DanhMuc" trong file mẫu. Các cột khác gõ tự do.
          </p>
        </div>

        <div className="flex gap-2 flex-wrap mb-3">
          <a href="/Mau-nhap-may.xlsx" download
            className="border border-brand-300 text-brand-700 hover:bg-brand-50 rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-2">
            <FileSpreadsheet size={15} /> Tải biểu mẫu
          </a>
          <label className="bg-brand-600 hover:bg-brand-700 text-white rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-2 cursor-pointer">
            <Plus size={15} /> Chọn file Excel
            <input type="file" accept=".xlsx,.xls" className="hidden"
              onChange={(e) => e.target.files?.[0] && readFile(e.target.files[0])} />
          </label>
          {fileName && <span className="text-xs text-slate-400 self-center">{fileName}</span>}
        </div>

        {error && <p className="text-xs text-rose-600 bg-rose-50 rounded-lg px-3 py-2 mb-3">{error}</p>}

        {result && (
          <div className="bg-emerald-50 border border-emerald-200 rounded-xl px-3 py-3 text-sm text-emerald-800 mb-3">
            <p>Đã nhập thành công <span className="font-semibold">{result.done}</span> máy.</p>
            {result.failed.length > 0 && (
              <div className="mt-2 text-rose-600 text-xs space-y-0.5">
                <p className="font-medium">{result.failed.length} dòng không nhập được:</p>
                {result.failed.map((f, k) => <p key={k}>Dòng {f.line}: {f.msg}</p>)}
              </div>
            )}
          </div>
        )}

        {rows && !result && (
          <>
            <div className="flex gap-3 text-xs mb-2 flex-wrap">
              <span className="bg-emerald-50 text-emerald-700 px-2 py-1 rounded">Hợp lệ: {okRows.length} máy · {fmtVND(totalCost)}</span>
              {badRows.length > 0 && <span className="bg-rose-50 text-rose-600 px-2 py-1 rounded">Lỗi: {badRows.length} dòng</span>}
            </div>
            <div className="border border-slate-200 rounded-xl overflow-hidden max-h-80 overflow-y-auto mb-3">
              <table className="w-full text-xs">
                <thead className="bg-slate-50 sticky top-0"><tr className="text-left text-slate-400">
                  <th className="px-2 py-1.5">Dòng</th>
                  <th className="px-2 py-1.5">IMEI</th>
                  <th className="px-2 py-1.5">Máy</th>
                  <th className="px-2 py-1.5 text-right">Giá</th>
                  <th className="px-2 py-1.5">Kết quả</th>
                </tr></thead>
                <tbody>
                  {rows.map((r) => (
                    <tr key={r.line} className={classNames("border-t border-slate-100", r.errors.length > 0 && "bg-rose-50/50")}>
                      <td className="px-2 py-1.5 text-slate-400">{r.line}</td>
                      <td className="px-2 py-1.5 text-slate-600">{r.imei || "—"}</td>
                      <td className="px-2 py-1.5 text-slate-700">{[r.model, r.storage, r.color].filter(Boolean).join(" ")}</td>
                      <td className="px-2 py-1.5 text-right text-slate-600">{fmtVND(r.price)}</td>
                      <td className="px-2 py-1.5">
                        {r.errors.length === 0
                          ? <span className="text-emerald-600">Hợp lệ</span>
                          : <span className="text-rose-600">{r.errors.join("; ")}</span>}
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </>
        )}

        {saving && (
          <div className="mb-3">
            <div className="h-2 bg-slate-100 rounded-full overflow-hidden">
              <div className="h-full bg-brand-600 transition-all" style={{ width: `${progress}%` }} />
            </div>
            <p className="text-xs text-slate-400 mt-1">Đang nhập... {progress}%</p>
          </div>
        )}

        <div className="flex gap-2">
          {rows && !result && (
            <button onClick={doImport} disabled={saving || okRows.length === 0}
              className="flex-1 bg-brand-600 hover:bg-brand-700 disabled:opacity-60 text-white rounded-xl py-2.5 text-sm font-medium flex items-center justify-center gap-2">
              {saving && <Loader2 size={15} className="animate-spin" />}
              Nhập {okRows.length} máy vào kho
            </button>
          )}
          <button onClick={onClose} className="px-4 rounded-xl border border-slate-200 text-sm text-slate-500 hover:bg-slate-50">
            {result ? "Đóng" : "Hủy"}
          </button>
        </div>
      </Card>
    </div>
  );
}

function EditPurchaseModal({ purchase, employee, onClose, onDone }) {
  const d = purchase.devices || {};
  const [form, setForm] = useState({
    imei: d.imei || "", model: d.model || "", storage: d.storage || "", color: d.color || "",
    condition: d.condition || "used", condition_percent: d.condition_percent ?? "",
    notes: d.notes || "", purchase_price: String(purchase.purchase_price ?? ""),
  });
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const set = (k) => (e) => setForm((f) => ({ ...f, [k]: e.target.value }));

  const oldPrice = Number(purchase.purchase_price || 0);
  const newPrice = Number(form.purchase_price) || 0;
  const diff = newPrice - oldPrice;

  const submit = async () => {
    if (!form.model.trim()) { setError("Model máy không được để trống."); return; }
    if (newPrice <= 0) { setError("Giá thu mua phải lớn hơn 0."); return; }
    setSaving(true); setError("");
    const { error: err } = await supabase.rpc("update_purchase_order", {
      p_purchase_order_id: purchase.id,
      p_purchase_price: newPrice,
      p_imei: form.imei.trim() || null,
      p_model: form.model.trim(),
      p_storage: form.storage.trim() || null,
      p_color: form.color.trim() || null,
      p_condition: form.condition,
      p_condition_percent: form.condition_percent === "" ? null : Number(form.condition_percent),
      p_notes: form.notes.trim() || null,
    });
    setSaving(false);
    if (err) { setError(err.message); return; }
    await supabase.from("audit_logs").insert({
      table_name: "purchase_orders", record_id: purchase.id, action: "update",
      old_data: purchase, new_data: form, performed_by: employee.id, store_id: employee.store_id,
    });
    onDone();
  };

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/40 flex items-center justify-center p-4 overflow-y-auto">
      <Card className="w-full max-w-lg p-5 my-8">
        <div className="flex items-center justify-between mb-4">
          <p className="font-semibold text-slate-800 text-sm">Sửa phiếu nhập {purchase.purchase_code}</p>
          <button onClick={onClose} className="text-slate-400 hover:text-rose-600"><X size={16} /></button>
        </div>
        <div className="space-y-3">
          <div className="grid sm:grid-cols-2 gap-3">
            <TextField label="Số IMEI" value={form.imei} onChange={set("imei")} />
            <TextField label="Model máy *" value={form.model} onChange={set("model")} list="dl-models-edit" />
            <TextField label="Dung lượng" value={form.storage} onChange={set("storage")} list="dl-storage" />
            <TextField label="Màu sắc" value={form.color} onChange={set("color")} list="dl-colors-edit" />
            <label className="block">
              <span className="text-xs font-medium text-slate-500 mb-1 block">Tình trạng</span>
              <select value={form.condition} onChange={set("condition")}
                className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm text-slate-800 outline-none focus:ring-2 focus:ring-brand-400/40 focus:border-brand-500 transition">
                <option value="used">Máy cũ</option>
                <option value="new">Máy mới</option>
              </select>
            </label>
            <TextField label="Độ mới (%)" type="number" value={form.condition_percent} onChange={set("condition_percent")} />
          </div>
          <datalist id="dl-models-edit">{IPHONE_MODEL_LIST.map((m) => <option key={m} value={m} />)}</datalist>
          <datalist id="dl-colors-edit">{coloroptionsForModel(form.model).map((c) => <option key={c} value={c} />)}</datalist>

          <TextField label="Giá thu mua (đ) *" type="number" value={form.purchase_price} onChange={set("purchase_price")} />
          {diff !== 0 && newPrice > 0 && (
            <div className="bg-amber-50 border border-amber-200 rounded-xl px-3 py-2 text-xs text-amber-800">
              Giá thay đổi {diff > 0 ? "tăng" : "giảm"} <span className="font-medium">{fmtVND(Math.abs(diff))}</span>.
              Hệ thống ghi một bút toán điều chỉnh vào sổ công nợ, không sửa bút toán cũ.
            </div>
          )}
          <TextField label="Ghi chú" value={form.notes} onChange={set("notes")} />

          {error && <p className="text-xs text-rose-600 bg-rose-50 rounded-lg px-3 py-2">{error}</p>}
          <div className="flex gap-2">
            <button onClick={submit} disabled={saving}
              className="flex-1 bg-brand-600 hover:bg-brand-700 disabled:opacity-60 text-white rounded-xl py-2.5 text-sm font-medium flex items-center justify-center gap-2">
              {saving && <Loader2 size={15} className="animate-spin" />} Lưu thay đổi
            </button>
            <button onClick={onClose} className="px-4 rounded-xl border border-slate-200 text-sm text-slate-500 hover:bg-slate-50">Hủy</button>
          </div>
        </div>
      </Card>
    </div>
  );
}

function PurchaseModule({ employee }) {
  const [purchases, setPurchases] = useState([]);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState("");
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [showImport, setShowImport] = useState(false);
  const [runImport, setRunImport] = useState(false);
  const [importSupplier, setImportSupplier] = useState(null);
  const canEdit = employee.role !== "ke_toan";
  const canImport = true;   // cả 3 vai đều nhập hàng loạt được, quyền do RPC kiểm tra
  const [payingDebt, setPayingDebt] = useState(null);
  const [payAmount, setPayAmount] = useState("");
  const [payBankId, setPayBankId] = useState(null);
  const { banks } = usePaymentOptions();

  const canCreate = employee.role !== "ke_toan";
  const canDelete = employee.role === "quan_ly";

  const load = useCallback(async () => {
    setLoading(true);
    const { data, error } = await supabase
      .from("purchase_orders")
      .select("*, customers(full_name, phone, cccd, cccd_issue_date, cccd_issue_place, date_of_birth, address), suppliers(name, phone), devices(imei, model, storage, color, condition), sales_orders(order_code), contracts(contract_code)")
      .order("created_at", { ascending: false })
      .limit(1000);
    if (!error) setPurchases(data || []);
    setLoading(false);
  }, []);

  useEffect(() => { load(); }, [load]);

  const filtered = purchases.filter((p) => {
    if (!search.trim()) return true;
    const q = search.trim().toLowerCase();
    return p.purchase_code?.toLowerCase().includes(q) || p.devices?.imei?.toLowerCase().includes(q)
      || p.customers?.full_name?.toLowerCase().includes(q) || p.suppliers?.name?.toLowerCase().includes(q);
  });

  const totalDebt = purchases.reduce((s, p) => s + Math.max(0, Number(p.purchase_price || 0) - Number(p.paid_amount ?? p.purchase_price ?? 0)), 0);

  const remove = async (p) => {
    if (!confirm(`Xóa phiếu thu mua "${p.purchase_code}"? Máy đã nhập kho sẽ không tự động bị xóa.`)) return;
    const { error } = await supabase.from("purchase_orders").delete().eq("id", p.id);
    if (error) { alert(error.message); return; }
    await supabase.from("audit_logs").insert({ table_name: "purchase_orders", record_id: p.id, action: "delete", old_data: p, performed_by: employee.id, store_id: employee.store_id });
    load();
  };

  const openPayDebt = (p) => {
    const debt = Math.max(0, Number(p.purchase_price) - Number(p.paid_amount ?? 0));
    setPayingDebt(p);
    setPayAmount(String(debt));
  };

  const submitPayDebt = async () => {
    const p = payingDebt;
    const addAmount = Number(payAmount) || 0;
    if (addAmount <= 0) return;
    if (!payBankId) { alert("Vui lòng chọn tài khoản chuyển tiền."); return; }
    // Ghi vào SỔ CÁI qua RPC — cột paid_amount do trigger tự tính lại
    const { data: remain, error } = await supabase.rpc("pay_supplier_debt", {
      p_purchase_order_id: p.id, p_amount: addAmount, p_note: "Trả tiền nhà cung cấp",
      p_bank_account_id: payBankId,
    });
    if (error) { alert(error.message); return; }
    await supabase.from("audit_logs").insert({
      table_name: "purchase_orders", record_id: p.id, action: "update",
      old_data: p, new_data: { ...p, con_no: remain }, performed_by: employee.id, store_id: employee.store_id,
    });
    setPayingDebt(null); setPayBankId(null);
    load();
  };

  return (
    <div>
      <div className="flex items-center justify-between mb-4 gap-2 flex-wrap">
        <div>
          <h2 className="text-lg font-semibold text-slate-800">Nhập máy / Thu cũ</h2>
          <p className="text-xs text-slate-400">
            {purchases.length} phiếu — Khách lẻ / NCC / đổi trừ khi bán
            {totalDebt > 0 && <span className="text-amber-600 font-medium"> · Tổng công nợ NCC: {fmtVND(totalDebt)}</span>}
          </p>
        </div>
        <div className="flex gap-2 flex-wrap">
          {canImport && (
            <button onClick={() => setShowImport(true)} className="bg-white border border-slate-200 hover:bg-slate-50 text-slate-600 rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-1.5">
              <FileSpreadsheet size={15} /> Nhập từ Excel
            </button>
          )}
          {canCreate && (
            <button onClick={() => setShowForm((s) => !s)} className="bg-brand-600 hover:bg-brand-700 text-white rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-1.5">
              <Banknote size={15} /> Nhập máy mới
            </button>
          )}
        </div>
      </div>

      {showForm && (
        <PurchaseForm employee={employee} onCancel={() => setShowForm(false)} onSaved={() => { setShowForm(false); load(); }} />
      )}

      <Card className="p-0 overflow-hidden">
        <div className="p-3 border-b border-slate-100">
          <div className="relative">
            <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-300" />
            <input
              value={search}
              onChange={(e) => setSearch(e.target.value)}
              placeholder="Tìm theo mã phiếu, IMEI, tên khách/NCC..."
              className="w-full pl-9 pr-3 py-2 rounded-lg border border-slate-300 bg-white text-sm placeholder:text-slate-400 outline-none focus:ring-2 focus:ring-brand-400/40 focus:border-brand-500 transition"
            />
          </div>
        </div>
        {loading ? (
          <div className="flex justify-center py-10"><Loader2 className="animate-spin text-slate-300" /></div>
        ) : filtered.length === 0 ? (
          <EmptyState icon={ArrowLeftRight} text="Chưa có phiếu thu mua nào." />
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead><tr className="text-left text-xs text-slate-400 border-b border-slate-100">
                <th className="px-3 py-2">Mã phiếu</th>
                <th className="px-3 py-2">Ngày</th>
                <th className="px-3 py-2">Đối tác</th>
                <th className="px-3 py-2">Máy thu</th>
                <th className="px-3 py-2">Giá thu mua</th>
                <th className="px-3 py-2">Nguồn</th>
                <th className="px-3 py-2">Công nợ</th>
                <th className="px-3 py-2"></th>
              </tr></thead>
              <tbody>
                {filtered.map((p) => {
                  const debt = Math.max(0, Number(p.purchase_price) - Number(p.paid_amount ?? p.purchase_price ?? 0));
                  return (
                  <tr key={p.id} className="border-b border-slate-50 last:border-0 hover:bg-slate-50/60">
                    <td className="px-3 py-2.5 font-medium text-slate-700 whitespace-nowrap">{p.purchase_code}</td>
                    <td className="px-3 py-2.5 text-slate-500">{fmtDate(p.created_at)}</td>
                    <td className="px-3 py-2.5 text-slate-600">{p.customers?.full_name || p.suppliers?.name}</td>
                    <td className="px-3 py-2.5 text-slate-500">
                      {p.devices?.model} <span className="text-slate-400">· {p.devices?.imei ? `IMEI ${p.devices.imei}` : "Chưa có IMEI"}</span>
                    </td>
                    <td className="px-3 py-2.5 text-slate-600">{fmtVND(p.purchase_price)}</td>
                    <td className="px-3 py-2.5">
                      {p.sales_orders ? (
                        <span className="text-xs px-2 py-0.5 rounded-full bg-brand-50 text-brand-700">Đổi trừ đơn {p.sales_orders.order_code}</span>
                      ) : p.source_type === "supplier" ? (
                        <span className="text-xs px-2 py-0.5 rounded-full bg-purple-50 text-purple-700">Nhà cung cấp</span>
                      ) : (
                        <span className="text-xs px-2 py-0.5 rounded-full bg-slate-100 text-slate-500">Khách lẻ</span>
                      )}
                    </td>
                    <td className="px-3 py-2.5">
                      {p.source_type === "supplier" && debt > 0 ? (
                        <button onClick={() => openPayDebt(p)} className="text-xs px-2 py-0.5 rounded-full bg-amber-50 text-amber-700 hover:bg-amber-100">
                          Còn nợ {fmtVND(debt)}
                        </button>
                      ) : p.source_type === "supplier" ? (
                        <span className="text-xs px-2 py-0.5 rounded-full bg-emerald-50 text-emerald-700">Đã thanh toán</span>
                      ) : (
                        <span className="text-slate-300">—</span>
                      )}
                    </td>
                    <td className="px-3 py-2.5 text-right whitespace-nowrap">
                      {canEdit && p.devices?.status !== "sold" && (
                        <button onClick={() => setEditing(p)} className="text-brand-600 hover:underline text-xs mr-3">
                          <Pencil size={12} className="inline mr-0.5" />Sửa
                        </button>
                      )}
                      {canDelete && (
                        <button onClick={() => remove(p)} className="text-rose-500 hover:underline text-xs">
                          <Trash2 size={12} className="inline mr-0.5" />Xóa
                        </button>
                      )}
                    </td>
                  </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </Card>

      {payingDebt && (
        <div className="fixed inset-0 z-50 bg-slate-900/40 flex items-center justify-center p-4">
          <Card className="w-full max-w-sm p-5">
            <div className="flex items-center justify-between mb-4">
              <p className="font-semibold text-slate-800 text-sm">Thanh toán công nợ — {payingDebt.purchase_code}</p>
              <button onClick={() => setPayingDebt(null)} className="text-slate-400 hover:text-rose-600"><X size={16} /></button>
            </div>
            <p className="text-xs text-slate-400 mb-3">
              NCC: {payingDebt.suppliers?.name} · Còn nợ: {fmtVND(Math.max(0, Number(payingDebt.purchase_price) - Number(payingDebt.paid_amount ?? 0)))}
            </p>
            <TextField label="Số tiền thanh toán thêm (đ)" type="number" value={payAmount} onChange={(e) => setPayAmount(e.target.value)} />
            <div className="mt-3">
              <span className="text-xs font-medium text-slate-500 mb-1 block">Tài khoản chuyển tiền *</span>
              <BankSelect banks={banks} value={payBankId} onChange={setPayBankId} className="w-full !text-sm !px-3 !py-2 !rounded-xl" />
              <p className="text-[11px] text-slate-400 mt-1">Mọi khoản chi ra đều qua chuyển khoản.</p>
            </div>
            <div className="flex gap-2 mt-4">
              <button onClick={submitPayDebt} className="bg-brand-600 hover:bg-brand-700 text-white rounded-xl px-4 py-2 text-sm font-medium">Xác nhận thanh toán</button>
              <button onClick={() => setPayingDebt(null)} className="text-slate-500 text-sm px-4 py-2">Hủy</button>
            </div>
          </Card>
        </div>
      )}

      {showImport && (
        <div className="fixed inset-0 z-50 bg-slate-900/40 flex items-center justify-center p-4">
          <Card className="w-full max-w-md p-5">
            <div className="flex items-center justify-between mb-4">
              <p className="font-semibold text-slate-800 text-sm">Nhập từ Excel — chọn nhà cung cấp</p>
              <button onClick={() => { setShowImport(false); setImportSupplier(null); }} className="text-slate-400 hover:text-rose-600"><X size={16} /></button>
            </div>
            <p className="text-xs text-slate-400 mb-3">
              Nhập hàng loạt chỉ áp dụng cho hàng nhập từ nhà cung cấp. Máy thu của khách lẻ nhập từng chiếc để có đủ hồ sơ.
            </p>
            <SupplierPicker value={importSupplier} onSelect={setImportSupplier} employee={employee} />
            {importSupplier && (
              <button
                onClick={() => { setShowImport(false); setImportSupplier(importSupplier); setRunImport(true); }}
                className="w-full mt-3 bg-brand-600 hover:bg-brand-700 text-white rounded-xl py-2.5 text-sm font-medium"
              >
                Tiếp tục với {importSupplier.name}
              </button>
            )}
          </Card>
        </div>
      )}

      {runImport && importSupplier && (
        <ImportPurchaseModal
          employee={employee}
          supplier={importSupplier}
          onClose={() => { setRunImport(false); setImportSupplier(null); }}
          onDone={load}
        />
      )}

      {editing && (
        <EditPurchaseModal
          purchase={editing}
          employee={employee}
          onClose={() => setEditing(null)}
          onDone={() => { setEditing(null); load(); }}
        />
      )}
    </div>
  );
}

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
              className="w-full pl-9 pr-3 py-2 rounded-lg border border-slate-300 bg-white text-sm placeholder:text-slate-400 outline-none focus:ring-2 focus:ring-brand-400/40 focus:border-brand-500 transition"
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

function CapitalCard({ employee }) {
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [showForm, setShowForm] = useState(false);
  const [form, setForm] = useState({
    contribution_date: todayStr(), contributor_name: "", amount: "", note: "", kind: "in",
  });
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const canManage = employee.role !== "nhan_vien";

  const load = useCallback(async () => {
    setLoading(true);
    const { data } = await supabase.from("v_capital_contributions").select("*")
      .order("contribution_date", { ascending: true }).limit(500);
    setRows(data || []);
    setLoading(false);
  }, []);

  useEffect(() => { load(); }, [load]);

  const total = rows.reduce((s, r) => s + Number(r.amount || 0), 0);
  const byPerson = {};
  for (const r of rows) byPerson[r.contributor_name] = (byPerson[r.contributor_name] || 0) + Number(r.amount || 0);
  const people = Object.entries(byPerson).sort((a, b) => b[1] - a[1]);

  const submit = async () => {
    const raw = Number(form.amount) || 0;
    if (!form.contributor_name.trim()) { setError("Vui lòng nhập tên người góp vốn."); return; }
    if (raw <= 0) { setError("Số tiền phải lớn hơn 0."); return; }
    setSaving(true); setError("");
    const { error: err } = await supabase.from("capital_contributions").insert({
      contribution_date: form.contribution_date,
      contributor_name: form.contributor_name.trim(),
      amount: form.kind === "out" ? -raw : raw,
      note: form.note.trim() || null,
      store_id: employee.store_id, created_by: employee.id,
    });
    setSaving(false);
    if (err) { setError(err.message); return; }
    setShowForm(false);
    setForm({ contribution_date: todayStr(), contributor_name: "", amount: "", note: "", kind: "in" });
    load();
  };

  const remove = async (r) => {
    if (!confirm(`Xóa bản ghi ${r.capital_code} — ${r.contributor_name} ${fmtVND(r.amount)}?`)) return;
    const { error: err } = await supabase.from("capital_contributions").delete().eq("id", r.id);
    if (err) { alert(err.message); return; }
    load();
  };

  return (
    <Card className="p-4 sm:p-5 mb-4">
      <div className="flex items-center justify-between mb-1 gap-2 flex-wrap">
        <p className="text-sm font-medium text-slate-700">Vốn góp đầu tư</p>
        {canManage && (
          <button onClick={() => { setShowForm((s) => !s); setError(""); }}
            className="text-xs bg-brand-600 hover:bg-brand-700 text-white rounded-lg px-3 py-1.5 flex items-center gap-1.5">
            <Plus size={13} /> Ghi vốn góp
          </button>
        )}
      </div>
      <p className="text-[11px] text-slate-400 mb-3">
        Vốn góp không phải chi phí, không trừ vào lợi nhuận. Dùng làm mẫu số đánh giá hiệu quả kinh doanh.
      </p>

      <div className="bg-brand-50 rounded-xl p-3 mb-3">
        <p className="text-xs text-brand-600 mb-1">Tổng vốn góp hiện tại</p>
        <p className="text-xl font-semibold text-brand-700">{fmtVND(total)}</p>
      </div>

      {people.length > 0 && (
        <div className="grid grid-cols-2 lg:grid-cols-3 gap-2 mb-3">
          {people.map(([name, amt]) => (
            <div key={name} className="bg-slate-50 rounded-lg px-3 py-2">
              <p className="text-xs text-slate-400 truncate">{name}</p>
              <p className="text-sm font-medium text-slate-700">{fmtVND(amt)}</p>
              {total > 0 && <p className="text-[11px] text-slate-400">{((amt / total) * 100).toFixed(1)}%</p>}
            </div>
          ))}
        </div>
      )}

      {showForm && (
        <div className="bg-slate-50 rounded-xl p-3 mb-3 space-y-3">
          <div className="grid sm:grid-cols-2 gap-3">
            <TextField label="Ngày" type="date" value={form.contribution_date}
              onChange={(e) => setForm((f) => ({ ...f, contribution_date: e.target.value }))} />
            <TextField label="Người góp vốn *" value={form.contributor_name}
              placeholder="Nguyễn Văn A"
              onChange={(e) => setForm((f) => ({ ...f, contributor_name: e.target.value }))} />
            <label className="block">
              <span className="text-xs font-medium text-slate-500 mb-1 block">Loại</span>
              <select value={form.kind} onChange={(e) => setForm((f) => ({ ...f, kind: e.target.value }))}
                className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm text-slate-800 outline-none focus:ring-2 focus:ring-brand-400/40 focus:border-brand-500 transition">
                <option value="in">Góp thêm vốn</option>
                <option value="out">Rút vốn</option>
              </select>
            </label>
            <TextField label="Số tiền (đ) *" type="number" value={form.amount}
              onChange={(e) => setForm((f) => ({ ...f, amount: e.target.value }))} />
          </div>
          <TextField label="Ghi chú" value={form.note}
            onChange={(e) => setForm((f) => ({ ...f, note: e.target.value }))} />
          {error && <p className="text-xs text-rose-600 bg-rose-50 rounded-lg px-3 py-2">{error}</p>}
          <div className="flex gap-2">
            <button onClick={submit} disabled={saving}
              className="bg-brand-600 hover:bg-brand-700 disabled:opacity-60 text-white rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-2">
              {saving && <Loader2 size={15} className="animate-spin" />} Lưu
            </button>
            <button onClick={() => setShowForm(false)} className="px-4 rounded-xl border border-slate-200 text-sm text-slate-500 hover:bg-white">Hủy</button>
          </div>
        </div>
      )}

      {loading ? (
        <div className="flex justify-center py-6"><Loader2 className="animate-spin text-slate-300" /></div>
      ) : rows.length === 0 ? (
        <p className="text-xs text-slate-400 py-4 text-center">Chưa ghi nhận vốn góp nào.</p>
      ) : (
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead><tr className="text-left text-xs text-slate-400 border-b border-slate-100">
              <th className="px-2 py-2">Mã</th>
              <th className="px-2 py-2">Ngày</th>
              <th className="px-2 py-2">Người góp</th>
              <th className="px-2 py-2">Ghi chú</th>
              <th className="px-2 py-2 text-right">Số tiền</th>
              <th className="px-2 py-2"></th>
            </tr></thead>
            <tbody>
              {rows.map((r) => (
                <tr key={r.id} className="border-b border-slate-50 last:border-0">
                  <td className="px-2 py-2 text-slate-400 doc-code whitespace-nowrap">{r.capital_code}</td>
                  <td className="px-2 py-2 text-slate-500 whitespace-nowrap">{fmtDate(r.contribution_date)}</td>
                  <td className="px-2 py-2 text-slate-700">{r.contributor_name}</td>
                  <td className="px-2 py-2 text-slate-400 text-xs">{r.note || "—"}</td>
                  <td className={classNames("px-2 py-2 text-right font-medium whitespace-nowrap",
                    Number(r.amount) < 0 ? "text-rose-600" : "text-slate-700")}>
                    {Number(r.amount) < 0 ? "" : "+"}{fmtVND(r.amount)}
                  </td>
                  <td className="px-2 py-2 text-right">
                    {employee.role === "quan_ly" && (
                      <button onClick={() => remove(r)} className="text-slate-400 hover:text-rose-600"><Trash2 size={13} /></button>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </Card>
  );
}

function IncomeStatementCard({ employee, fromDate, toDate, storeName }) {
  const [d, setD] = useState(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    let active = true;
    setLoading(true);
    supabase.rpc("income_statement", { p_from: fromDate, p_to: toDate })
      .then(({ data }) => {
        if (!active) return;
        setD(Array.isArray(data) ? data[0] : data);
        setLoading(false);
      });
    return () => { active = false; };
  }, [fromDate, toDate]);

  if (loading) return (
    <Card className="p-4 sm:p-5 mb-4">
      <div className="flex justify-center py-8"><Loader2 className="animate-spin text-slate-300" /></div>
    </Card>
  );
  if (!d) return null;

  const num = (v) => Number(v || 0);
  const net = num(d.net_profit);
  const roi = num(d.capital_total) > 0 ? (net / num(d.capital_total)) * 100 : null;
  const marginPct = num(d.revenue) > 0 ? (num(d.gross_profit) / num(d.revenue)) * 100 : 0;

  const Row = ({ label, value, bold, indent, tone, hint }) => (
    <div className={classNames("flex justify-between py-1.5", indent && "pl-4",
      bold && "border-t border-slate-200 mt-1 pt-2")}>
      <span className={classNames("text-slate-500", bold && "font-medium text-slate-700")}>
        {label}
        {hint && <span className="text-[11px] text-slate-400 ml-1">{hint}</span>}
      </span>
      <span className={classNames("whitespace-nowrap",
        bold ? "font-semibold" : "",
        tone === "neg" ? "text-rose-600" : tone === "pos" ? "text-emerald-700" : "text-slate-700")}>
        {tone === "neg" && num(value) > 0 ? "−" : ""}{fmtVND(Math.abs(num(value)))}
      </span>
    </div>
  );

  return (
    <Card className="p-4 sm:p-5 mb-4">
      <p className="text-sm font-medium text-slate-700 mb-1">Kết quả kinh doanh — {storeName}</p>
      <p className="text-[11px] text-slate-400 mb-3">
        {fmtDate(fromDate)} đến {fmtDate(toDate)} · {d.order_count} đơn bán
      </p>

      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 mb-4">
        <div className="bg-emerald-50 rounded-xl p-3">
          <p className="text-xs text-emerald-600 mb-1">Doanh thu</p>
          <p className="text-base font-semibold text-emerald-700">{fmtVND(d.revenue)}</p>
        </div>
        <div className="bg-brand-50 rounded-xl p-3">
          <p className="text-xs text-brand-600 mb-1">Lãi gộp</p>
          <p className="text-base font-semibold text-brand-700">{fmtVND(d.gross_profit)}</p>
          <p className="text-[11px] text-brand-500">Biên {marginPct.toFixed(1)}%</p>
        </div>
        <div className="bg-rose-50 rounded-xl p-3">
          <p className="text-xs text-rose-500 mb-1">Tổng chi phí</p>
          <p className="text-base font-semibold text-rose-600">{fmtVND(d.exp_total)}</p>
        </div>
        <div className={classNames("rounded-xl p-3", net >= 0 ? "bg-emerald-50" : "bg-rose-50")}>
          <p className={classNames("text-xs mb-1", net >= 0 ? "text-emerald-600" : "text-rose-500")}>Lợi nhuận</p>
          <p className={classNames("text-base font-semibold", net >= 0 ? "text-emerald-700" : "text-rose-600")}>
            {fmtVND(net)}
          </p>
          {roi !== null && <p className={classNames("text-[11px]", net >= 0 ? "text-emerald-500" : "text-rose-400")}>
            {roi >= 0 ? "+" : ""}{roi.toFixed(1)}% trên vốn góp
          </p>}
        </div>
      </div>

      <div className="text-sm">
        <Row label="Doanh thu bán hàng" value={d.revenue} tone="pos" />
        <Row label="Giá vốn hàng bán" value={d.cogs} tone="neg" indent />
        <Row label="Lãi gộp" value={d.gross_profit} bold />

        <div className="mt-3">
          <Row label="Chi phí khác" value={d.exp_other} tone="neg" indent />
          <Row label="Chi thưởng nhân viên" value={d.exp_bonus} tone="neg" indent />
          <Row label="Marketing — quảng cáo" value={d.exp_marketing} tone="neg" indent />
          <Row label="Giảm trừ doanh thu" value={d.exp_sales_return} tone="neg" indent hint="khách trả máy" />
          <Row label="Tổng chi phí" value={d.exp_total} bold tone="neg" />
        </div>

        <div className={classNames("flex justify-between py-2.5 mt-2 px-3 rounded-xl",
          net >= 0 ? "bg-emerald-50" : "bg-rose-50")}>
          <span className={classNames("font-semibold", net >= 0 ? "text-emerald-800" : "text-rose-700")}>
            Lợi nhuận trong kỳ
          </span>
          <span className={classNames("font-bold", net >= 0 ? "text-emerald-800" : "text-rose-700")}>
            {fmtVND(net)}
          </span>
        </div>
      </div>

      {num(d.internal_out_revenue) > 0 && (
        <div className="mt-4 bg-indigo-50 rounded-xl p-3 text-sm">
          <p className="text-xs font-medium text-indigo-700 mb-1">Xuất bán nội bộ (không tính vào lợi nhuận trên)</p>
          <div className="flex justify-between text-xs text-indigo-700 py-0.5">
            <span>Giá trị xuất</span><span>{fmtVND(d.internal_out_revenue)}</span>
          </div>
          <div className="flex justify-between text-xs text-indigo-700 py-0.5">
            <span>Giá vốn máy xuất</span><span>{fmtVND(d.internal_out_cost)}</span>
          </div>
          <div className="flex justify-between text-xs font-semibold text-indigo-800 border-t border-indigo-200 pt-1 mt-1">
            <span>Lãi nội bộ</span><span>{fmtVND(d.internal_margin)}</span>
          </div>
        </div>
      )}

      <div className="mt-4 grid grid-cols-2 gap-3 text-sm">
        <div className="bg-slate-50 rounded-xl p-3">
          <p className="text-xs text-slate-400 mb-1">Vốn góp lũy kế</p>
          <p className="font-semibold text-slate-800">{fmtVND(d.capital_total)}</p>
        </div>
        <div className="bg-slate-50 rounded-xl p-3">
          <p className="text-xs text-slate-400 mb-1">Giá trị tồn kho</p>
          <p className="font-semibold text-slate-800">{fmtVND(d.inventory_value)}</p>
        </div>
      </div>
    </Card>
  );
}

function CommissionSection({ employee, fromDate, toDate }) {
  const [summary, setSummary] = useState([]);
  const [detail, setDetail] = useState([]);
  const [loading, setLoading] = useState(true);
  const [openId, setOpenId] = useState(null);
  const [showExcluded, setShowExcluded] = useState(false);

  const canSeeMargin = employee.role !== "nhan_vien";

  const load = useCallback(async () => {
    setLoading(true);
    const [{ data: s }, { data: d }] = await Promise.all([
      supabase.rpc("commission_by_period", { p_from: fromDate, p_to: toDate }),
      supabase.from("v_sales_commission").select("*")
        .gte("sale_date", fromDate).lte("sale_date", toDate)
        .order("sale_date", { ascending: false }).limit(3000),
    ]);
    setSummary(s || []);
    setDetail(d || []);
    setLoading(false);
  }, [fromDate, toDate]);

  useEffect(() => { load(); }, [load]);

  const totalDevices = summary.reduce((a, r) => a + Number(r.device_count || 0), 0);
  const totalRevenue = summary.reduce((a, r) => a + Number(r.revenue || 0), 0);
  const totalMargin = summary.reduce((a, r) => a + Number(r.gross_margin || 0), 0);
  const excluded = detail.filter((r) => !r.counts_for_commission);

  const exportExcel = () => {
    const rows = detail
      .filter((r) => showExcluded || r.counts_for_commission)
      .map((r) => ({
        "Ngày bán": fmtDate(r.sale_date),
        "Mã đơn": r.order_code,
        "Nhân viên": r.employee_name || "",
        "Khách hàng": r.customer_name || "",
        "IMEI": r.imei || "",
        "Máy": [r.model, r.storage, r.color].filter(Boolean).join(" "),
        "Giá bán": Number(r.total_amount || 0),
        ...(canSeeMargin ? {
          "Giá vốn": Number(r.cost_price || 0),
          "Lãi gộp": Number(r.gross_margin || 0),
        } : {}),
        "Tính hoa hồng": r.counts_for_commission ? "Có" : "Không",
        "Lý do loại": r.excluded_reason || "",
      }));

    const sum = summary.map((r) => ({
      "Nhân viên": r.employee_name,
      "Số máy": Number(r.device_count || 0),
      "Doanh thu": Number(r.revenue || 0),
      ...(canSeeMargin ? {
        "Lãi gộp": Number(r.gross_margin || 0),
        "Lãi TB/máy": Number(r.avg_margin || 0),
      } : {}),
      "Đơn bị loại": Number(r.excluded_count || 0),
    }));

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(sum), "TongHop");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows), "ChiTiet");
    XLSX.writeFile(wb, `Hoa-hong-${fromDate}_${toDate}.xlsx`);
  };

  return (
    <Card className="p-4 sm:p-5 mb-4">
      <div className="flex items-center justify-between mb-1 gap-2 flex-wrap">
        <p className="text-sm font-medium text-slate-700">Máy bán được theo nhân viên</p>
        <button onClick={exportExcel} disabled={detail.length === 0}
          className="text-xs border border-brand-300 text-brand-700 hover:bg-brand-50 disabled:opacity-50 rounded-lg px-3 py-1.5 flex items-center gap-1.5">
          <FileSpreadsheet size={13} /> Xuất Excel
        </button>
      </div>
      <p className="text-[11px] text-slate-400 mb-3">
        Chỉ tính đơn bán khách lẻ. Không tính đơn đã hủy, khách trả máy, bán sỉ cho nhà cung cấp, và xuất bán nội bộ.
      </p>

      <div className="grid grid-cols-3 gap-3 mb-4">
        <div className="bg-slate-50 rounded-xl p-3">
          <p className="text-xs text-slate-400 mb-1">Máy tính hoa hồng</p>
          <p className="text-lg font-semibold text-slate-800">{totalDevices}</p>
        </div>
        <div className="bg-emerald-50 rounded-xl p-3">
          <p className="text-xs text-emerald-600 mb-1">Doanh thu</p>
          <p className="text-lg font-semibold text-emerald-700">{fmtVND(totalRevenue)}</p>
        </div>
        {canSeeMargin ? (
          <div className="bg-brand-50 rounded-xl p-3">
            <p className="text-xs text-brand-600 mb-1">Lãi gộp</p>
            <p className="text-lg font-semibold text-brand-700">{fmtVND(totalMargin)}</p>
          </div>
        ) : (
          <div className="bg-slate-50 rounded-xl p-3">
            <p className="text-xs text-slate-400 mb-1">Đơn bị loại</p>
            <p className="text-lg font-semibold text-slate-600">{excluded.length}</p>
          </div>
        )}
      </div>

      {loading ? (
        <div className="flex justify-center py-8"><Loader2 className="animate-spin text-slate-300" /></div>
      ) : summary.length === 0 ? (
        <p className="text-xs text-slate-400 py-6 text-center">Không có đơn bán nào trong khoảng ngày đã chọn.</p>
      ) : (
        <div className="overflow-x-auto">
          <table className="w-full text-sm">
            <thead><tr className="text-left text-xs text-slate-400 border-b border-slate-100">
              <th className="px-2 py-2">Nhân viên</th>
              <th className="px-2 py-2 text-right">Số máy</th>
              <th className="px-2 py-2 text-right">Doanh thu</th>
              {canSeeMargin && <th className="px-2 py-2 text-right">Lãi gộp</th>}
              {canSeeMargin && <th className="px-2 py-2 text-right">Lãi TB/máy</th>}
              <th className="px-2 py-2 text-right">Bị loại</th>
              <th className="px-2 py-2"></th>
            </tr></thead>
            <tbody>
              {summary.map((r) => {
                const isOpen = openId === r.employee_id;
                const mine = detail.filter((x) => x.employee_id === r.employee_id
                  && (showExcluded || x.counts_for_commission));
                return (
                  <React.Fragment key={r.employee_id}>
                    <tr className="border-b border-slate-50 hover:bg-slate-50/60 cursor-pointer"
                      onClick={() => setOpenId(isOpen ? null : r.employee_id)}>
                      <td className="px-2 py-2.5 font-medium text-slate-700">{r.employee_name}</td>
                      <td className="px-2 py-2.5 text-right font-semibold text-slate-800">{r.device_count}</td>
                      <td className="px-2 py-2.5 text-right text-slate-600 whitespace-nowrap">{fmtVND(r.revenue)}</td>
                      {canSeeMargin && <td className="px-2 py-2.5 text-right text-brand-700 whitespace-nowrap">{fmtVND(r.gross_margin)}</td>}
                      {canSeeMargin && <td className="px-2 py-2.5 text-right text-slate-500 whitespace-nowrap">{fmtVND(r.avg_margin)}</td>}
                      <td className="px-2 py-2.5 text-right">
                        {Number(r.excluded_count) > 0
                          ? <span className="text-xs px-2 py-0.5 rounded-full bg-rose-50 text-rose-600">{r.excluded_count}</span>
                          : <span className="text-slate-300">—</span>}
                      </td>
                      <td className="px-2 py-2.5 text-right">
                        <ChevronDown size={15} className={classNames("inline text-slate-300 transition-transform", isOpen && "rotate-180")} />
                      </td>
                    </tr>
                    {isOpen && (
                      <tr>
                        <td colSpan={canSeeMargin ? 7 : 5} className="bg-slate-50/70 px-3 py-3">
                          {mine.length === 0 ? (
                            <p className="text-xs text-slate-400 text-center py-2">Không có đơn nào.</p>
                          ) : (
                            <table className="w-full text-xs">
                              <thead><tr className="text-left text-slate-400 border-b border-slate-200">
                                <th className="py-1.5">Ngày</th>
                                <th className="py-1.5">Đơn</th>
                                <th className="py-1.5">Máy</th>
                                <th className="py-1.5">Khách</th>
                                <th className="py-1.5 text-right">Giá bán</th>
                                {canSeeMargin && <th className="py-1.5 text-right">Lãi</th>}
                              </tr></thead>
                              <tbody>
                                {mine.map((x) => (
                                  <tr key={x.order_id} className={classNames("border-b border-slate-100 last:border-0",
                                    !x.counts_for_commission && "opacity-60")}>
                                    <td className="py-1.5 text-slate-500 whitespace-nowrap">{fmtDate(x.sale_date)}</td>
                                    <td className="py-1.5 text-slate-400">{x.order_code}</td>
                                    <td className="py-1.5 text-slate-600">
                                      {[x.model, x.storage, x.color].filter(Boolean).join(" ")}
                                      <span className="text-slate-400"> · {x.imei || "—"}</span>
                                      {!x.counts_for_commission && (
                                        <span className="ml-1 text-[10px] px-1 py-0.5 rounded bg-rose-50 text-rose-600">{x.excluded_reason}</span>
                                      )}
                                    </td>
                                    <td className="py-1.5 text-slate-500">{x.customer_name || "—"}</td>
                                    <td className="py-1.5 text-right text-slate-700 whitespace-nowrap">{fmtVND(x.total_amount)}</td>
                                    {canSeeMargin && <td className="py-1.5 text-right text-brand-700 whitespace-nowrap">{fmtVND(x.gross_margin)}</td>}
                                  </tr>
                                ))}
                              </tbody>
                            </table>
                          )}
                        </td>
                      </tr>
                    )}
                  </React.Fragment>
                );
              })}
            </tbody>
          </table>
        </div>
      )}

      {excluded.length > 0 && (
        <button onClick={() => setShowExcluded((s) => !s)}
          className="mt-3 text-xs text-brand-600 hover:underline">
          {showExcluded ? "Ẩn các đơn không tính hoa hồng" : `Hiện ${excluded.length} đơn không tính hoa hồng`}
        </button>
      )}
    </Card>
  );
}

function ReportsModule({ employee }) {
  const [fromDate, setFromDate] = useState(startOfMonth());
  const [toDate, setToDate] = useState(todayStr());
  const [orders, setOrders] = useState([]);
  const [payments, setPayments] = useState([]);
  const [inventoryValue, setInventoryValue] = useState(0);
  const [inStockCount, setInStockCount] = useState(0);
  const [loading, setLoading] = useState(true);

  const canSeeProfit = employee.role !== "nhan_vien";

  const load = useCallback(async () => {
    setLoading(true);
    const [{ data: orderData }, { data: deviceData }, { data: payData }] = await Promise.all([
      supabase
        .from("sales_orders")
        .select("*, devices(model, storage, color, cost_price), employees:created_by(full_name)")
        .eq("status", "completed")
        .gte("created_at", `${fromDate}T00:00:00`)
        .lte("created_at", `${toDate}T23:59:59`)
        .order("created_at", { ascending: true }),
      supabase.from("devices").select("cost_price, status"),
      // Phiếu thu phát sinh trong kỳ — gồm cả tiền thu nợ của đơn kỳ trước
      supabase
        .from("order_payments")
        .select("method, amount, installment_provider, bank_accounts(short_label, bank_name)")
        .gte("created_at", `${fromDate}T00:00:00`)
        .lte("created_at", `${toDate}T23:59:59`)
        .limit(5000),
    ]);
    setOrders(orderData || []);
    setPayments(payData || []);
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

  // ---- Tiền thu trong kỳ, chia theo hình thức ----
  const sumBy = (m) => payments.filter((p) => p.method === m)
    .reduce((s, p) => s + Number(p.amount || 0), 0);

  const cashIn        = sumBy("cash");
  const bankIn        = sumBy("bank_transfer");
  const installmentIn = sumBy("installment");
  const offsetIn      = sumBy("debt_offset");
  const tradeInIn     = sumBy("trade_in");
  const realCashIn    = cashIn + bankIn + installmentIn;          // tiền thật vào
  const totalCollected = realCashIn + offsetIn + tradeInIn;       // tổng đã tất toán

  // Công nợ phát sinh từ chính các đơn trong kỳ
  const periodDebt = orders.reduce(
    (s, o) => s + Math.max(0, Number(o.total_amount || 0) - Number(o.paid_amount || 0)), 0
  );

  // Chi tiết chuyển khoản theo từng tài khoản
  const byBank = {};
  for (const p of payments.filter((x) => x.method === "bank_transfer")) {
    const k = p.bank_accounts?.short_label || p.bank_accounts?.bank_name || "Chưa gán tài khoản";
    byBank[k] = (byBank[k] || 0) + Number(p.amount || 0);
  }
  const bankRows = Object.entries(byBank).sort((a, b) => b[1] - a[1]);

  // Chi tiết trả góp theo đơn vị
  const byProvider = {};
  for (const p of payments.filter((x) => x.method === "installment")) {
    const k = p.installment_provider || "Chưa ghi đơn vị";
    byProvider[k] = (byProvider[k] || 0) + Number(p.amount || 0);
  }
  const providerRows = Object.entries(byProvider).sort((a, b) => b[1] - a[1]);

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
            <ReportKpiCard label="Doanh thu ghi nhận" value={fmtVND(totalRevenue)} icon={TrendingUp} sub={`${totalOrders} đơn · gồm cả chưa thu`} />
            {canSeeProfit && <ReportKpiCard label="Lợi nhuận gộp" value={fmtVND(totalProfit)} icon={Award} sub={totalRevenue ? `Biên LN ${((totalProfit / totalRevenue) * 100).toFixed(1)}%` : ""} />}
            <ReportKpiCard label="Giá trị đơn TB" value={fmtVND(avgOrderValue)} icon={Receipt} />
            {canSeeProfit && <ReportKpiCard label="Tồn kho hiện tại" value={fmtVND(inventoryValue)} icon={Package} sub={`${inStockCount} máy còn hàng`} />}
          </div>


          <Card className="p-4 sm:p-5 mb-4">
            <div className="flex items-baseline justify-between mb-1">
              <p className="text-sm font-medium text-slate-700">Dòng tiền theo hình thức</p>
              <p className="text-xs text-slate-400">phiếu thu phát sinh trong kỳ</p>
            </div>
            <p className="text-[11px] text-slate-400 mb-3">
              Gồm cả tiền thu nợ của đơn kỳ trước, nên có thể lệch với Doanh thu ở trên.
            </p>

            <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 mb-4">
              <div className="bg-emerald-50 rounded-xl p-3">
                <p className="text-xs text-emerald-700 mb-1 flex items-center gap-1"><Wallet size={13} /> Tiền mặt</p>
                <p className="text-base font-semibold text-emerald-800">{fmtVND(cashIn)}</p>
              </div>
              <div className="bg-sky-50 rounded-xl p-3">
                <p className="text-xs text-sky-700 mb-1 flex items-center gap-1"><Landmark size={13} /> Chuyển khoản</p>
                <p className="text-base font-semibold text-sky-800">{fmtVND(bankIn)}</p>
              </div>
              <div className="bg-violet-50 rounded-xl p-3">
                <p className="text-xs text-violet-700 mb-1 flex items-center gap-1"><CalendarClock size={13} /> Trả góp</p>
                <p className="text-base font-semibold text-violet-800">{fmtVND(installmentIn)}</p>
              </div>
              <div className="bg-amber-50 rounded-xl p-3">
                <p className="text-xs text-amber-700 mb-1 flex items-center gap-1"><Banknote size={13} /> Công nợ phát sinh</p>
                <p className="text-base font-semibold text-amber-800">{fmtVND(periodDebt)}</p>
                <p className="text-[10px] text-amber-600 mt-0.5">từ đơn trong kỳ, chưa thu</p>
              </div>
            </div>

            <div className="flex flex-wrap gap-3 text-xs mb-4">
              <div className="bg-slate-50 rounded-lg px-3 py-2">
                <span className="text-slate-400">Tiền thật vào két/ngân hàng: </span>
                <span className="font-semibold text-slate-700">{fmtVND(realCashIn)}</span>
              </div>
              {offsetIn > 0 && (
                <div className="bg-indigo-50 rounded-lg px-3 py-2 text-indigo-700">
                  <span className="opacity-70">Bù trừ công nợ (không phải tiền): </span>
                  <span className="font-semibold">{fmtVND(offsetIn)}</span>
                </div>
              )}
              {tradeInIn > 0 && (
                <div className="bg-slate-50 rounded-lg px-3 py-2">
                  <span className="text-slate-400">Đổi máy cũ (đơn cũ): </span>
                  <span className="font-semibold text-slate-700">{fmtVND(tradeInIn)}</span>
                </div>
              )}
              <div className="bg-slate-50 rounded-lg px-3 py-2">
                <span className="text-slate-400">Tổng đã tất toán: </span>
                <span className="font-semibold text-slate-700">{fmtVND(totalCollected)}</span>
              </div>
            </div>

            <div className="grid sm:grid-cols-2 gap-4">
              <div>
                <p className="text-xs font-medium text-slate-500 mb-2">Chuyển khoản theo tài khoản</p>
                {bankRows.length === 0 ? (
                  <p className="text-xs text-slate-400">Không có giao dịch chuyển khoản.</p>
                ) : (
                  <div className="space-y-1.5">
                    {bankRows.map(([name, amt]) => (
                      <div key={name} className="flex justify-between text-sm">
                        <span className="text-slate-600">{name}</span>
                        <span className="font-medium text-slate-700">{fmtVND(amt)}</span>
                      </div>
                    ))}
                  </div>
                )}
              </div>
              <div>
                <p className="text-xs font-medium text-slate-500 mb-2">Trả góp theo đơn vị</p>
                {providerRows.length === 0 ? (
                  <p className="text-xs text-slate-400">Không có giao dịch trả góp.</p>
                ) : (
                  <div className="space-y-1.5">
                    {providerRows.map(([name, amt]) => (
                      <div key={name} className="flex justify-between text-sm">
                        <span className="text-slate-600">{name}</span>
                        <span className="font-medium text-slate-700">{fmtVND(amt)}</span>
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
          </Card>

          <IncomeStatementCard employee={employee} fromDate={fromDate} toDate={toDate}
            storeName={employee.stores?.name || "Cửa hàng"} />

          <CapitalCard employee={employee} />

          <CommissionSection employee={employee} fromDate={fromDate} toDate={toDate} />

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
        store_id: employee.store_id,
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
              <select value={form.role} onChange={(e) => setForm((f) => ({ ...f, role: e.target.value }))} className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm text-slate-800 outline-none focus:ring-2 focus:ring-brand-400/40 focus:border-brand-500 transition">
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

function PartnerLedgerPanel({ partner }) {
  const [entries, setEntries] = useState(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    let active = true;
    (async () => {
      setLoading(true);
      const { data } = await supabase
        .from("partner_ledger")
        .select("*, sales_orders(order_code), purchase_orders(purchase_code), debt_offsets:offset_id(offset_code)")
        .eq("partner_id", partner.partner_id)
        .order("entry_date", { ascending: true })
        .order("created_at", { ascending: true })
        .limit(500);
      if (active) { setEntries(data || []); setLoading(false); }
    })();
    return () => { active = false; };
  }, [partner.partner_id]);

  if (loading) return <div className="flex justify-center py-6"><Loader2 className="animate-spin text-slate-300" size={18} /></div>;
  if (!entries || entries.length === 0) return <p className="text-xs text-slate-400 py-4 text-center">Chưa có bút toán nào.</p>;

  let runR = 0, runP = 0;
  const rows = entries.map((e) => {
    if (e.account === "receivable") runR += Number(e.amount); else runP += Number(e.amount);
    return { ...e, runR, runP };
  });

  return (
    <div className="overflow-x-auto">
      <table className="w-full text-xs">
        <thead><tr className="text-left text-slate-400 border-b border-slate-200">
          <th className="px-2 py-1.5">Mã</th>
          <th className="px-2 py-1.5">Ngày</th>
          <th className="px-2 py-1.5">Nội dung</th>
          <th className="px-2 py-1.5">Chứng từ</th>
          <th className="px-2 py-1.5 text-right">Phải thu</th>
          <th className="px-2 py-1.5 text-right">Phải trả</th>
        </tr></thead>
        <tbody>
          {rows.map((e) => {
            const doc = e.sales_orders?.order_code || e.purchase_orders?.purchase_code || e.debt_offsets?.offset_code || "—";
            const amt = Number(e.amount);
            return (
              <tr key={e.id} className="border-b border-slate-100 last:border-0">
                <td className="px-2 py-1.5 text-slate-400 whitespace-nowrap">{e.entry_code}</td>
                <td className="px-2 py-1.5 text-slate-500 whitespace-nowrap">{fmtDate(e.entry_date)}</td>
                <td className="px-2 py-1.5 text-slate-600 whitespace-nowrap">
                  {LEDGER_TYPE_LABELS[e.entry_type] || e.entry_type}
                  {e.entry_type === "offset" && <span className="ml-1 text-[10px] bg-indigo-50 text-indigo-600 px-1.5 py-0.5 rounded">BT</span>}
                </td>
                <td className="px-2 py-1.5 text-slate-400 whitespace-nowrap">{doc}</td>
                <td className="px-2 py-1.5 text-right whitespace-nowrap">
                  {e.account === "receivable" ? (
                    <span className={amt > 0 ? "text-slate-700" : "text-emerald-600"}>
                      {amt > 0 ? "+" : ""}{fmtVND(amt)} <span className="text-slate-300">| {fmtVND(e.runR)}</span>
                    </span>
                  ) : <span className="text-slate-200">—</span>}
                </td>
                <td className="px-2 py-1.5 text-right whitespace-nowrap">
                  {e.account === "payable" ? (
                    <span className={amt > 0 ? "text-slate-700" : "text-emerald-600"}>
                      {amt > 0 ? "+" : ""}{fmtVND(amt)} <span className="text-slate-300">| {fmtVND(e.runP)}</span>
                    </span>
                  ) : <span className="text-slate-200">—</span>}
                </td>
              </tr>
            );
          })}
        </tbody>
      </table>
      <p className="text-[11px] text-slate-400 mt-2">
        Cột nhạt bên phải mỗi số là số dư lũy kế. Sổ cái bất biến — mọi sai sót được sửa bằng bút toán điều chỉnh, không xóa dòng cũ.
      </p>
    </div>
  );
}


function SettleInstallmentModal({ row, employee, onClose, onDone }) {
  const [received, setReceived] = useState(String(row.amount));
  const [fee, setFee] = useState("0");
  const [bankId, setBankId] = useState(null);
  const [settledAt, setSettledAt] = useState(todayStr());
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const { banks } = usePaymentOptions();

  const recv = Number(received) || 0;
  const f = Number(fee) || 0;
  const diff = Number(row.amount) - recv - f;

  const submit = async () => {
    if (recv <= 0) { setError("Số tiền nhận phải lớn hơn 0."); return; }
    if (Math.round(diff) !== 0) {
      setError(`Tiền nhận cộng phí phải bằng ${fmtVND(row.amount)} — đang lệch ${fmtVND(diff)}.`);
      return;
    }
    setSaving(true); setError("");
    const { error: err } = await supabase.rpc("settle_installment", {
      p_payment_id: row.payment_id, p_received: recv,
      p_bank_account_id: bankId, p_fee: f, p_settled_at: settledAt,
    });
    setSaving(false);
    if (err) { setError(err.message); return; }
    await supabase.from("audit_logs").insert({
      table_name: "order_payments", record_id: row.payment_id, action: "update",
      new_data: { da_nhan: recv, phi: f, ngay: settledAt },
      performed_by: employee.id, store_id: employee.store_id,
    });
    onDone();
  };

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/40 flex items-center justify-center p-4">
      <Card className="w-full max-w-md p-5">
        <div className="flex items-center justify-between mb-4">
          <p className="font-semibold text-slate-800 text-sm">Xác nhận nhận tiền trả góp</p>
          <button onClick={onClose} className="text-slate-400 hover:text-rose-600"><X size={16} /></button>
        </div>
        <div className="space-y-3">
          <div className="bg-violet-50 rounded-xl px-3 py-2.5 text-xs text-violet-800 space-y-0.5">
            <p className="font-medium">{row.installment_provider} · {fmtVND(row.amount)}</p>
            <p>Đơn {row.order_code} · {row.customer_name}</p>
            {row.installment_contract_code && <p>Mã hồ sơ: {row.installment_contract_code}</p>}
            <p className="text-violet-500">Bán ngày {fmtDate(row.order_date)}</p>
          </div>
          <TextField label="Ngày nhận theo sao kê" type="date" value={settledAt} onChange={(e) => setSettledAt(e.target.value)} />
          <div className="grid grid-cols-2 gap-2">
            <TextField label="Số thực nhận (đ)" type="number" value={received} onChange={(e) => setReceived(e.target.value)} />
            <TextField label="Phí giữ lại (đ)" type="number" value={fee} onChange={(e) => setFee(e.target.value)} />
          </div>
          <div>
            <span className="text-xs font-medium text-slate-500 mb-1 block">Tài khoản nhận tiền</span>
            <BankSelect banks={banks} value={bankId} onChange={setBankId} className="w-full !text-sm !px-3 !py-2 !rounded-xl" />
          </div>
          <div className={classNames(
            "text-xs px-3 py-2 rounded-lg",
            Math.round(diff) === 0 ? "bg-emerald-50 text-emerald-700" : "bg-rose-50 text-rose-600"
          )}>
            {Math.round(diff) === 0
              ? `Khớp: ${fmtVND(recv)} + phí ${fmtVND(f)} = ${fmtVND(row.amount)}`
              : `Lệch ${fmtVND(diff)} — kiểm tra lại số nhận hoặc phí.`}
          </div>
          {error && <p className="text-xs text-rose-600 bg-rose-50 rounded-lg px-3 py-2">{error}</p>}
          <div className="flex gap-2">
            <button onClick={submit} disabled={saving}
              className="flex-1 bg-brand-600 hover:bg-brand-700 disabled:opacity-60 text-white rounded-xl py-2.5 text-sm font-medium flex items-center justify-center gap-2">
              {saving && <Loader2 size={15} className="animate-spin" />} Xác nhận đã nhận
            </button>
            <button onClick={onClose} className="px-4 rounded-xl border border-slate-200 text-sm text-slate-500 hover:bg-slate-50">Hủy</button>
          </div>
        </div>
      </Card>
    </div>
  );
}

function InstallmentTracking({ employee }) {
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [status, setStatus] = useState("pending");
  const [provider, setProvider] = useState("all");
  const [search, setSearch] = useState("");
  const [settling, setSettling] = useState(null);

  const canSettle = employee.role !== "nhan_vien";

  const load = useCallback(async () => {
    setLoading(true);
    const { data } = await supabase
      .from("v_installment_tracking").select("*")
      .order("created_at", { ascending: false }).limit(2000);
    setRows(data || []);
    setLoading(false);
  }, []);

  useEffect(() => { load(); }, [load]);

  const providers = [...new Set(rows.map((r) => r.installment_provider).filter(Boolean))].sort();

  const filtered = rows.filter((r) => {
    if (status !== "all" && r.settlement_status !== status) return false;
    if (provider !== "all" && r.installment_provider !== provider) return false;
    if (!search.trim()) return true;
    const q = search.trim().toLowerCase();
    return [r.order_code, r.customer_name, r.customer_phone, r.installment_contract_code, r.imei]
      .some((v) => v?.toLowerCase().includes(q));
  });

  const pending = rows.filter((r) => r.settlement_status === "pending");
  const pendingByProvider = {};
  for (const r of pending) {
    const k = r.installment_provider || "Chưa ghi đơn vị";
    if (!pendingByProvider[k]) pendingByProvider[k] = { count: 0, amount: 0 };
    pendingByProvider[k].count += 1;
    pendingByProvider[k].amount += Number(r.amount || 0);
  }
  const pendingRows = Object.entries(pendingByProvider).sort((a, b) => b[1].amount - a[1].amount);
  const totalPending = pending.reduce((s, r) => s + Number(r.amount || 0), 0);

  return (
    <div>
      <div className="bg-violet-50 border border-violet-200 rounded-xl px-3 py-2.5 mb-4 text-xs text-violet-800 flex items-center gap-2">
        <CalendarClock size={15} className="shrink-0" />
        Đang chờ giải ngân <span className="font-medium">{fmtVND(totalPending)}</span> trên {pending.length} hồ sơ
        <span className="text-violet-400">· đối chiếu với sao kê cuối ngày</span>
      </div>

      {pendingRows.length > 0 && (
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 mb-4">
          {pendingRows.map(([name, v]) => (
            <button
              key={name} onClick={() => { setProvider(name); setStatus("pending"); }}
              className={classNames(
                "text-left rounded-xl p-3 border transition",
                provider === name ? "bg-violet-100 border-violet-300" : "bg-white border-slate-200 hover:border-violet-200"
              )}
            >
              <p className="text-xs text-slate-500 mb-1">{name}</p>
              <p className="text-base font-semibold text-slate-800">{fmtVND(v.amount)}</p>
              <p className="text-[11px] text-slate-400">{v.count} hồ sơ</p>
            </button>
          ))}
        </div>
      )}

      <Card className="p-0 overflow-hidden">
        <div className="p-3 border-b border-slate-100 flex gap-2 flex-wrap">
          <div className="relative flex-1 min-w-[180px]">
            <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-300" />
            <input
              value={search} onChange={(e) => setSearch(e.target.value)}
              placeholder="Tìm mã đơn, khách, mã hồ sơ, IMEI..."
              className="w-full pl-9 pr-3 py-2 rounded-lg border border-slate-300 bg-white text-sm placeholder:text-slate-400 outline-none focus:ring-2 focus:ring-brand-400/40 focus:border-brand-500 transition"
            />
          </div>
          <select value={provider} onChange={(e) => setProvider(e.target.value)}
            className="rounded-xl border border-slate-200 px-3 py-2 text-sm">
            <option value="all">Tất cả đơn vị</option>
            {providers.map((p) => <option key={p} value={p}>{p}</option>)}
          </select>
          {[["pending", "Chờ nhận"], ["settled", "Đã nhận"], ["all", "Tất cả"]].map(([k, label]) => (
            <button key={k} onClick={() => setStatus(k)}
              className={classNames("px-3 py-2 rounded-xl text-sm border",
                status === k ? "bg-brand-600 text-white border-brand-600" : "bg-white text-slate-500 border-slate-200")}
            >{label}</button>
          ))}
        </div>

        {loading ? (
          <div className="flex justify-center py-10"><Loader2 className="animate-spin text-slate-300" /></div>
        ) : filtered.length === 0 ? (
          <EmptyState icon={CalendarClock} text="Không có hồ sơ trả góp nào khớp bộ lọc." />
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead><tr className="text-left text-xs text-slate-400 border-b border-slate-100">
                <th className="px-3 py-2">Ngày bán</th>
                <th className="px-3 py-2">Đơn / Khách</th>
                <th className="px-3 py-2">Đơn vị</th>
                <th className="px-3 py-2">Mã hồ sơ</th>
                <th className="px-3 py-2 text-right">Số tiền</th>
                <th className="px-3 py-2">Trạng thái</th>
                <th className="px-3 py-2"></th>
              </tr></thead>
              <tbody>
                {filtered.map((r) => (
                  <tr key={r.payment_id} className="border-b border-slate-50 last:border-0">
                    <td className="px-3 py-2.5 text-slate-500 whitespace-nowrap">{fmtDate(r.order_date)}</td>
                    <td className="px-3 py-2.5">
                      <p className="font-medium text-slate-700">{r.order_code}</p>
                      <p className="text-xs text-slate-400">{r.customer_name}{r.customer_phone ? ` · ${r.customer_phone}` : ""}</p>
                    </td>
                    <td className="px-3 py-2.5 text-slate-600 whitespace-nowrap">{r.installment_provider || "—"}</td>
                    <td className="px-3 py-2.5 text-slate-400 text-xs">{r.installment_contract_code || "—"}</td>
                    <td className="px-3 py-2.5 text-right font-medium text-slate-700 whitespace-nowrap">{fmtVND(r.amount)}</td>
                    <td className="px-3 py-2.5 whitespace-nowrap">
                      {r.settlement_status === "settled" ? (
                        <span className="text-xs px-2 py-0.5 rounded-full bg-emerald-50 text-emerald-700">
                          Đã nhận {r.settled_at ? fmtDate(r.settled_at) : ""}
                          {Number(r.settlement_fee) > 0 && ` · phí ${fmtVND(r.settlement_fee)}`}
                        </span>
                      ) : (
                        <span className="text-xs px-2 py-0.5 rounded-full bg-violet-50 text-violet-700">Chờ giải ngân</span>
                      )}
                    </td>
                    <td className="px-3 py-2.5 text-right whitespace-nowrap">
                      {canSettle && r.settlement_status === "pending" && (
                        <button onClick={() => setSettling(r)} className="text-xs text-brand-600 hover:underline">
                          Đã nhận tiền
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

      {settling && (
        <SettleInstallmentModal
          row={settling} employee={employee}
          onClose={() => setSettling(null)}
          onDone={() => { setSettling(null); load(); }}
        />
      )}
    </div>
  );
}

function SettleAdvanceModal({ row, employee, onClose, onDone }) {
  const remaining = Number(row.remaining || 0);
  const [amount, setAmount] = useState(String(remaining));
  const [method, setMethod] = useState("bank_transfer");
  const [bankId, setBankId] = useState(null);
  const [date, setDate] = useState(todayStr());
  const [note, setNote] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const { banks } = usePaymentOptions();

  const submit = async () => {
    const n = Number(amount) || 0;
    if (n <= 0) { setError("Số tiền nộp phải lớn hơn 0."); return; }
    if (n > remaining) { setError(`Phiếu này chỉ còn ${fmtVND(remaining)}.`); return; }
    if (method === "bank_transfer" && !bankId) { setError("Vui lòng chọn tài khoản nhận."); return; }
    setSaving(true); setError("");
    const { error: err } = await supabase.rpc("settle_cash_advance", {
      p_advance_id: row.id, p_amount: n, p_method: method,
      p_bank_account_id: method === "bank_transfer" ? bankId : null,
      p_settled_date: date, p_note: note.trim() || null,
    });
    setSaving(false);
    if (err) { setError(err.message); return; }
    await supabase.from("audit_logs").insert({
      table_name: "cash_advances", record_id: row.id, action: "update",
      new_data: { nop: n, hinh_thuc: method, ngay: date },
      performed_by: employee.id, store_id: employee.store_id,
    });
    onDone();
  };

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/40 flex items-center justify-center p-4">
      <Card className="w-full max-w-md p-5">
        <div className="flex items-center justify-between mb-4">
          <p className="font-semibold text-slate-800 text-sm">Hoàn ứng — {row.advance_code}</p>
          <button onClick={onClose} className="text-slate-400 hover:text-rose-600"><X size={16} /></button>
        </div>
        <div className="space-y-3">
          <div className="bg-slate-50 rounded-xl px-3 py-2.5 text-xs space-y-0.5">
            <p className="font-medium text-slate-700">{row.holder_name}</p>
            <p className="text-slate-400">Tạm ứng ngày {fmtDate(row.advance_date)}</p>
            <div className="flex justify-between pt-1 mt-1 border-t border-slate-200">
              <span className="text-slate-400">Đã ứng</span>
              <span className="font-medium text-slate-700">{fmtVND(row.amount)}</span>
            </div>
            <div className="flex justify-between">
              <span className="text-slate-400">Đã nộp lại</span>
              <span className="text-slate-600">{fmtVND(row.settled_amount)}</span>
            </div>
            <div className="flex justify-between text-amber-700">
              <span>Còn phải nộp</span>
              <span className="font-medium">{fmtVND(remaining)}</span>
            </div>
          </div>
          <TextField label="Ngày nộp" type="date" value={date} onChange={(e) => setDate(e.target.value)} />
          <div>
            <span className="text-xs font-medium text-slate-500 mb-1 block">Hình thức nộp</span>
            <select value={method} onChange={(e) => setMethod(e.target.value)}
              className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm text-slate-800 outline-none focus:ring-2 focus:ring-brand-400/40 focus:border-brand-500 transition">
              <option value="bank_transfer">Chuyển khoản</option>
              <option value="cash">Nộp tiền mặt</option>
            </select>
          </div>
          {method === "bank_transfer" && (
            <div>
              <span className="text-xs font-medium text-slate-500 mb-1 block">Tài khoản nhận *</span>
              <BankSelect banks={banks} value={bankId} onChange={setBankId} className="w-full !text-sm !px-3 !py-2 !rounded-xl" />
            </div>
          )}
          <TextField label="Số tiền nộp (đ)" type="number" value={amount} onChange={(e) => setAmount(e.target.value)} />
          <TextField label="Ghi chú" value={note} onChange={(e) => setNote(e.target.value)} />
          {error && <p className="text-xs text-rose-600 bg-rose-50 rounded-lg px-3 py-2">{error}</p>}
          <div className="flex gap-2">
            <button onClick={submit} disabled={saving}
              className="flex-1 bg-brand-600 hover:bg-brand-700 disabled:opacity-60 text-white rounded-xl py-2.5 text-sm font-medium flex items-center justify-center gap-2">
              {saving && <Loader2 size={15} className="animate-spin" />} Xác nhận đã nộp
            </button>
            <button onClick={onClose} className="px-4 rounded-xl border border-slate-200 text-sm text-slate-500 hover:bg-slate-50">Hủy</button>
          </div>
        </div>
      </Card>
    </div>
  );
}

function CashAdvanceTab({ employee }) {
  const [date, setDate] = useState(todayStr());
  const [summary, setSummary] = useState(null);
  const [staff, setStaff] = useState([]);
  const [holderId, setHolderId] = useState("");
  const [amount, setAmount] = useState("");
  const [note, setNote] = useState("");
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const [okCode, setOkCode] = useState(null);
  const [settling, setSettling] = useState(null);
  const [showSettled, setShowSettled] = useState(false);

  const canManage = employee.role !== "nhan_vien";

  const loadSummary = useCallback(async () => {
    const { data } = await supabase.rpc("daily_cash_summary", { p_date: date });
    const s = Array.isArray(data) ? data[0] : data;
    setSummary(s || null);
    setAmount(s ? String(Math.max(0, Number(s.suggested) || 0)) : "");
  }, [date]);

  const loadRows = useCallback(async () => {
    setLoading(true);
    const { data } = await supabase.from("v_cash_advances").select("*")
      .order("advance_date", { ascending: false }).limit(500);
    setRows(data || []);
    setLoading(false);
  }, []);

  useEffect(() => { loadSummary(); }, [loadSummary]);
  useEffect(() => { loadRows(); }, [loadRows]);
  useEffect(() => {
    (async () => {
      const { data } = await supabase.from("employees")
        .select("id, full_name, role").eq("is_active", true).order("full_name");
      setStaff(data || []);
    })();
  }, []);

  const submit = async () => {
    if (!holderId) { setError("Vui lòng chọn nhân sự giữ tiền."); return; }
    const n = Number(amount) || 0;
    if (n <= 0) { setError("Số tiền tạm ứng phải lớn hơn 0."); return; }
    setSaving(true); setError(""); setOkCode(null);
    const { data, error: err } = await supabase.rpc("create_cash_advance", {
      p_date: date, p_employee_id: holderId, p_amount: n, p_note: note.trim() || null,
    });
    setSaving(false);
    if (err) { setError(err.message); return; }
    setOkCode(data); setNote("");
    loadSummary(); loadRows();
  };

  const open = rows.filter((r) => r.status !== "settled");
  const shown = showSettled ? rows : open;
  const totalOpen = open.reduce((s, r) => s + Number(r.remaining || 0), 0);

  const byHolder = {};
  for (const r of open) {
    byHolder[r.holder_name] = (byHolder[r.holder_name] || 0) + Number(r.remaining || 0);
  }
  const holderRows = Object.entries(byHolder).sort((a, b) => b[1] - a[1]);

  return (
    <div className="space-y-3">
      {totalOpen > 0 && (
        <div className="bg-amber-50 border border-amber-200 rounded-xl px-3 py-2.5 text-xs text-amber-800 flex items-center gap-2 flex-wrap">
          <Wallet size={15} className="shrink-0" />
          Đang tạm ứng <span className="font-medium">{fmtVND(totalOpen)}</span> ·
          {holderRows.map(([n, v]) => <span key={n}>{n}: <span className="font-medium">{fmtVND(v)}</span></span>)}
        </div>
      )}

      {canManage && (
        <Card className="p-4">
          <p className="text-sm font-medium text-slate-700 mb-1">Lập phiếu tạm ứng cuối ngày</p>
          <p className="text-xs text-slate-400 mb-3">
            Chọn ngày và nhân sự giữ tiền. Số tiền tự điền theo tồn quỹ tiền mặt, sửa được nếu kiểm đếm thực tế lệch.
          </p>

          <div className="grid sm:grid-cols-3 gap-3">
            <TextField label="Ngày" type="date" value={date} onChange={(e) => { setDate(e.target.value); setOkCode(null); }} />
            <label className="block">
              <span className="text-xs font-medium text-slate-500 mb-1 block">Nhân sự tạm ứng *</span>
              <select value={holderId} onChange={(e) => setHolderId(e.target.value)}
                className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm text-slate-800 outline-none focus:ring-2 focus:ring-brand-400/40 focus:border-brand-500 transition">
                <option value="">— Chọn nhân sự —</option>
                {staff.map((s) => (
                  <option key={s.id} value={s.id}>{s.full_name} ({ROLE_LABELS[s.role] || s.role})</option>
                ))}
              </select>
            </label>
            <TextField label="Số tiền tạm ứng (đ)" type="number" value={amount} onChange={(e) => setAmount(e.target.value)} />
          </div>

          {summary && (
            <div className="mt-3 bg-slate-50 rounded-xl p-3 text-xs space-y-1">
              <p className="font-medium text-slate-600 mb-1">Tiền mặt ngày {fmtDate(date)}</p>
              <div className="flex justify-between"><span className="text-slate-400">Thu tiền mặt</span><span className="text-emerald-700 font-medium">+{fmtVND(summary.cash_in)}</span></div>
              {Number(summary.cash_out) > 0 ? (
                <>
                  {Number(summary.out_purchase) > 0 && <div className="flex justify-between"><span className="text-rose-500">Chi mua máy bằng tiền mặt</span><span className="text-rose-600">−{fmtVND(summary.out_purchase)}</span></div>}
                  {Number(summary.out_refund) > 0 && <div className="flex justify-between"><span className="text-rose-500">Hoàn khách bằng tiền mặt</span><span className="text-rose-600">−{fmtVND(summary.out_refund)}</span></div>}
                  {Number(summary.out_expense) > 0 && <div className="flex justify-between"><span className="text-rose-500">Chi phí bằng tiền mặt</span><span className="text-rose-600">−{fmtVND(summary.out_expense)}</span></div>}
                  <p className="text-[11px] text-rose-500">Có khoản chi bằng tiền mặt — theo quy định mọi khoản chi phải qua chuyển khoản, kiểm tra lại.</p>
                </>
              ) : (
                <p className="text-[11px] text-slate-400">Mọi khoản chi ra đều qua chuyển khoản, nên tồn quỹ đúng bằng tiền mặt thu được.</p>
              )}
              <div className="flex justify-between border-t border-slate-200 pt-1 mt-1"><span className="text-slate-500">Tồn quỹ</span><span className="font-semibold text-slate-800">{fmtVND(summary.net_cash)}</span></div>
              {Number(summary.already_advanced) > 0 && (
                <div className="flex justify-between text-amber-700"><span>Đã tạm ứng cho ngày này</span><span className="font-medium">−{fmtVND(summary.already_advanced)}</span></div>
              )}
              <div className="flex justify-between"><span className="text-slate-500">Còn lại để ứng</span><span className="font-semibold text-brand-700">{fmtVND(summary.suggested)}</span></div>
            </div>
          )}

          <div className="mt-3">
            <TextField label="Ghi chú" value={note} onChange={(e) => setNote(e.target.value)} placeholder="Giao ca tối, nộp ngân hàng sáng mai..." />
          </div>

          {error && <p className="text-xs text-rose-600 bg-rose-50 rounded-lg px-3 py-2 mt-3">{error}</p>}
          {okCode && <p className="text-xs text-emerald-700 bg-emerald-50 rounded-lg px-3 py-2 mt-3">Đã lập phiếu <span className="font-semibold">{okCode}</span>.</p>}

          <button onClick={submit} disabled={saving}
            className="mt-3 bg-brand-600 hover:bg-brand-700 disabled:opacity-60 text-white rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-2">
            {saving && <Loader2 size={15} className="animate-spin" />} Lưu phiếu tạm ứng
          </button>
        </Card>
      )}

      <div className="flex items-center justify-between gap-2 flex-wrap">
        <p className="text-xs text-slate-400">{open.length} phiếu chưa nộp đủ</p>
        <button onClick={() => setShowSettled((s) => !s)}
          className="text-xs text-brand-600 hover:underline">
          {showSettled ? "Chỉ xem phiếu chưa nộp" : "Xem cả phiếu đã nộp đủ"}
        </button>
      </div>

      <Card className="p-0 overflow-hidden">
        {loading ? (
          <div className="flex justify-center py-10"><Loader2 className="animate-spin text-slate-300" /></div>
        ) : shown.length === 0 ? (
          <EmptyState icon={Wallet} text="Chưa có phiếu tạm ứng nào." />
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead><tr className="text-left text-xs text-slate-400 border-b border-slate-100">
                <th className="px-3 py-2">Mã</th>
                <th className="px-3 py-2">Ngày</th>
                <th className="px-3 py-2">Người giữ</th>
                <th className="px-3 py-2 text-right">Đã ứng</th>
                <th className="px-3 py-2 text-right">Còn nộp</th>
                <th className="px-3 py-2"></th>
              </tr></thead>
              <tbody>
                {shown.map((r) => (
                  <tr key={r.id} className="border-b border-slate-50 last:border-0">
                    <td className="px-3 py-2.5 text-slate-400 doc-code whitespace-nowrap">{r.advance_code}</td>
                    <td className="px-3 py-2.5 text-slate-500 whitespace-nowrap">{fmtDate(r.advance_date)}</td>
                    <td className="px-3 py-2.5">
                      <p className="text-slate-700">{r.holder_name}</p>
                      {r.note && <p className="text-xs text-slate-400">{r.note}</p>}
                    </td>
                    <td className="px-3 py-2.5 text-right text-slate-600 whitespace-nowrap">{fmtVND(r.amount)}</td>
                    <td className="px-3 py-2.5 text-right whitespace-nowrap">
                      {r.status === "settled"
                        ? <span className="text-xs px-2 py-0.5 rounded-full bg-emerald-50 text-emerald-700">Đã nộp đủ</span>
                        : <span className="font-medium text-amber-600">{fmtVND(r.remaining)}</span>}
                    </td>
                    <td className="px-3 py-2.5 text-right whitespace-nowrap">
                      {canManage && r.status !== "settled" && (
                        <button onClick={() => setSettling(r)} className="text-xs text-brand-600 hover:underline">Hoàn ứng</button>
                      )}
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Card>

      {settling && (
        <SettleAdvanceModal
          row={settling} employee={employee}
          onClose={() => setSettling(null)}
          onDone={() => { setSettling(null); loadRows(); }}
        />
      )}
    </div>
  );
}

function ManualOffsetModal({ partner, onClose, onDone }) {
  const rec = Number(partner.receivable), pay = Number(partner.payable);
  const maxOffset = Math.min(rec, pay);
  const [amount, setAmount] = useState(String(maxOffset));
  const [note, setNote] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const [done, setDone] = useState(null);

  const submit = async () => {
    const n = Number(amount) || 0;
    if (n <= 0) { setError("Số tiền bù trừ phải lớn hơn 0."); return; }
    if (n > maxOffset) { setError(`Chỉ bù trừ được tối đa ${fmtVND(maxOffset)}.`); return; }
    setSaving(true); setError("");
    const { data, error: err } = await supabase.rpc("create_manual_offset", {
      p_partner_id: partner.partner_id, p_amount: n, p_note: note.trim() || null,
    });
    setSaving(false);
    if (err) { setError(err.message); return; }
    setDone(data);
  };

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/40 flex items-center justify-center p-4">
      <Card className="w-full max-w-md p-5">
        <div className="flex items-center justify-between mb-4">
          <p className="font-semibold text-slate-800 text-sm">Lập biên bản bù trừ công nợ</p>
          <button onClick={onClose} className="text-slate-400 hover:text-rose-600"><X size={16} /></button>
        </div>

        {done ? (
          <div className="space-y-3">
            <div className="bg-emerald-50 border border-emerald-200 rounded-xl px-3 py-3 text-sm text-emerald-800">
              Đã lập biên bản <span className="font-semibold">{done}</span> cho {partner.name}.
              Cả hai bên phải thu và phải trả đều đã giảm {fmtVND(Number(amount))}.
            </div>
            <button onClick={onDone} className="w-full bg-brand-600 hover:bg-brand-700 text-white rounded-xl py-2.5 text-sm font-medium">Xong</button>
          </div>
        ) : (
          <div className="space-y-3">
            <div className="bg-slate-50 rounded-xl px-3 py-2.5 text-xs space-y-1">
              <p className="font-medium text-slate-700">{partner.name}</p>
              <div className="flex justify-between"><span className="text-slate-400">Đối tác nợ cửa hàng</span><span className="text-amber-600 font-medium">{fmtVND(rec)}</span></div>
              <div className="flex justify-between"><span className="text-slate-400">Cửa hàng nợ đối tác</span><span className="text-indigo-600 font-medium">{fmtVND(pay)}</span></div>
              <div className="flex justify-between border-t border-slate-200 pt-1 mt-1"><span className="text-slate-500">Bù trừ tối đa</span><span className="font-semibold text-slate-700">{fmtVND(maxOffset)}</span></div>
            </div>
            <TextField label="Số tiền bù trừ (đ)" type="number" value={amount} onChange={(e) => setAmount(e.target.value)} />
            <TextField label="Nội dung / ghi chú" value={note} onChange={(e) => setNote(e.target.value)} placeholder="Hai bên thống nhất bù trừ công nợ..." />
            <p className="text-[11px] text-slate-400">
              Hệ thống phân bổ vào các chứng từ cũ nhất trước ở cả hai bên và sinh biên bản có mã BT để tra cứu.
            </p>
            {error && <p className="text-xs text-rose-600 bg-rose-50 rounded-lg px-3 py-2">{error}</p>}
            <div className="flex gap-2">
              <button onClick={submit} disabled={saving || maxOffset <= 0}
                className="flex-1 bg-brand-600 hover:bg-brand-700 disabled:opacity-60 text-white rounded-xl py-2.5 text-sm font-medium flex items-center justify-center gap-2">
                {saving && <Loader2 size={15} className="animate-spin" />} Lập biên bản
              </button>
              <button onClick={onClose} className="px-4 rounded-xl border border-slate-200 text-sm text-slate-500 hover:bg-slate-50">Hủy</button>
            </div>
          </div>
        )}
      </Card>
    </div>
  );
}

function DebtModule({ employee }) {
  const [tab, setTab] = useState("partners");   // partners | installment
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState("");
  const [filter, setFilter] = useState("all"); // all | receivable | payable
  const [openId, setOpenId] = useState(null);
  const [offsetting, setOffsetting] = useState(null);
  const canOffset = employee.role !== "ke_toan";

  const load = useCallback(async () => {
    setLoading(true);
    const { data, error } = await supabase.from("v_partner_balance").select("*").limit(2000);
    if (!error) setRows(data || []);
    setLoading(false);
  }, []);

  useEffect(() => { load(); }, [load]);

  const active = rows.filter((r) => Number(r.receivable) !== 0 || Number(r.payable) !== 0);
  const filtered = active.filter((r) => {
    if (filter === "receivable" && Number(r.receivable) <= 0) return false;
    if (filter === "payable" && Number(r.payable) <= 0) return false;
    if (!search.trim()) return true;
    const q = search.trim().toLowerCase();
    return r.name?.toLowerCase().includes(q) || r.phone?.toLowerCase().includes(q) || r.partner_code?.toLowerCase().includes(q);
  }).sort((a, b) => (Number(b.receivable) + Number(b.payable)) - (Number(a.receivable) + Number(a.payable)));

  const totalR = active.reduce((s, r) => s + Math.max(0, Number(r.receivable)), 0);
  const totalP = active.reduce((s, r) => s + Math.max(0, Number(r.payable)), 0);

  return (
    <div>
      <div className="flex items-center justify-between mb-4 gap-2 flex-wrap">
        <div>
          <h2 className="text-lg font-semibold text-slate-800">Công nợ</h2>
          <p className="text-xs text-slate-400">
            {tab === "partners" ? `${active.length} đối tác đang có số dư`
              : tab === "installment" ? "Đối soát tiền về từ đơn vị trả góp"
              : "Tiền mặt cuối ngày giao cho nhân sự giữ"}
          </p>
        </div>
        {tab === "partners" && (
          <button onClick={load} className="bg-white border border-slate-200 hover:bg-slate-50 text-slate-600 rounded-xl px-3 py-2 text-sm flex items-center gap-1.5">
            <Loader2 size={14} className={loading ? "animate-spin" : ""} /> Tải lại
          </button>
        )}
      </div>

      <div className="flex gap-2 mb-4">
        {[["partners", "Đối tác"], ["installment", "Trả góp"], ["advance", "Tạm ứng"]].map(([k, label]) => (
          <button key={k} onClick={() => setTab(k)}
            className={classNames("px-4 py-2 rounded-xl text-sm border font-medium",
              tab === k ? "bg-brand-600 text-white border-brand-600" : "bg-white text-slate-500 border-slate-200 hover:bg-slate-50")}
          >{label}</button>
        ))}
      </div>

      {tab === "installment" && <InstallmentTracking employee={employee} />}
      {tab === "advance" && <CashAdvanceTab employee={employee} />}
      {tab === "partners" && (<>

      <div className="grid grid-cols-2 gap-3 mb-4">
        <Card className="p-4">
          <p className="text-xs text-slate-400 mb-1">Phải thu — khách nợ cửa hàng</p>
          <p className="text-xl font-semibold text-amber-600">{fmtVND(totalR)}</p>
        </Card>
        <Card className="p-4">
          <p className="text-xs text-slate-400 mb-1">Phải trả — cửa hàng nợ đối tác</p>
          <p className="text-xl font-semibold text-indigo-600">{fmtVND(totalP)}</p>
        </Card>
      </div>

      <Card className="p-0 overflow-hidden">
        <div className="p-3 border-b border-slate-100 flex gap-2 flex-wrap">
          <div className="relative flex-1 min-w-[180px]">
            <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-300" />
            <input
              value={search} onChange={(e) => setSearch(e.target.value)}
              placeholder="Tìm theo tên, SĐT, mã đối tác..."
              className="w-full pl-9 pr-3 py-2 rounded-lg border border-slate-300 bg-white text-sm placeholder:text-slate-400 outline-none focus:ring-2 focus:ring-brand-400/40 focus:border-brand-500 transition"
            />
          </div>
          {[["all", "Tất cả"], ["receivable", "Phải thu"], ["payable", "Phải trả"]].map(([k, label]) => (
            <button
              key={k} onClick={() => setFilter(k)}
              className={classNames(
                "px-3 py-2 rounded-xl text-sm border",
                filter === k ? "bg-brand-600 text-white border-brand-600" : "bg-white text-slate-500 border-slate-200"
              )}
            >{label}</button>
          ))}
        </div>

        {loading ? (
          <div className="flex justify-center py-10"><Loader2 className="animate-spin text-slate-300" /></div>
        ) : filtered.length === 0 ? (
          <EmptyState icon={Banknote} text="Không có đối tác nào đang có công nợ." />
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead><tr className="text-left text-xs text-slate-400 border-b border-slate-100">
                <th className="px-3 py-2">Đối tác</th>
                <th className="px-3 py-2">Vai trò</th>
                <th className="px-3 py-2 text-right">Phải thu</th>
                <th className="px-3 py-2 text-right">Phải trả</th>
                <th className="px-3 py-2"></th>
              </tr></thead>
              <tbody>
                {filtered.map((r) => {
                  const rec = Number(r.receivable), pay = Number(r.payable);
                  const isOpen = openId === r.partner_id;
                  return (
                    <React.Fragment key={r.partner_id}>
                      <tr
                        className="border-b border-slate-50 last:border-0 hover:bg-slate-50/60 cursor-pointer"
                        onClick={() => setOpenId(isOpen ? null : r.partner_id)}
                      >
                        <td className="px-3 py-2.5">
                          <p className="font-medium text-slate-700">{r.name}</p>
                          <p className="text-xs text-slate-400">{r.partner_code} {r.phone ? `· ${r.phone}` : ""}</p>
                        </td>
                        <td className="px-3 py-2.5">
                          <div className="flex gap-1 flex-wrap">
                            {r.customer_id && <span className="text-[10px] bg-slate-100 text-slate-500 px-1.5 py-0.5 rounded">Khách</span>}
                            {r.supplier_id && <span className="text-[10px] bg-indigo-50 text-indigo-600 px-1.5 py-0.5 rounded">NCC</span>}
                          </div>
                        </td>
                        <td className="px-3 py-2.5 text-right">
                          {rec > 0 ? <span className="text-amber-600 font-medium">{fmtVND(rec)}</span> : <span className="text-slate-300">—</span>}
                        </td>
                        <td className="px-3 py-2.5 text-right">
                          {pay > 0 ? <span className="text-indigo-600 font-medium">{fmtVND(pay)}</span> : <span className="text-slate-300">—</span>}
                        </td>
                        <td className="px-3 py-2.5 text-right">
                          <ChevronDown size={15} className={classNames("inline text-slate-300 transition-transform", isOpen && "rotate-180")} />
                        </td>
                      </tr>
                      {isOpen && (
                        <tr>
                          <td colSpan={5} className="bg-slate-50/70 px-4 py-4">
                            {rec > 0 && pay > 0 && (
                              <div className="bg-indigo-50 border border-indigo-200 rounded-xl px-3 py-2.5 mb-3 text-xs text-indigo-800 flex items-center justify-between gap-3 flex-wrap">
                                <span>
                                  Đối tác này vừa nợ cửa hàng {fmtVND(rec)} vừa được cửa hàng nợ {fmtVND(pay)} —
                                  bù trừ được tối đa <span className="font-medium">{fmtVND(Math.min(rec, pay))}</span>.
                                </span>
                                {canOffset && (
                                  <button
                                    onClick={() => setOffsetting(r)}
                                    className="bg-indigo-600 hover:bg-indigo-700 text-white rounded-lg px-3 py-1.5 text-xs font-medium whitespace-nowrap"
                                  >
                                    Lập biên bản bù trừ
                                  </button>
                                )}
                              </div>
                            )}
                            <PartnerLedgerPanel partner={r} />
                          </td>
                        </tr>
                      )}
                    </React.Fragment>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </Card>
      </>)}

      {offsetting && (
        <ManualOffsetModal
          partner={offsetting}
          onClose={() => setOffsetting(null)}
          onDone={() => { setOffsetting(null); setOpenId(null); load(); }}
        />
      )}
    </div>
  );
}


/* -------------------------------------------------------------- */
/* Chi phí — thay cho Phiếu thu/chi                                */
/* -------------------------------------------------------------- */

const EXPENSE_CATEGORY_LABELS = {
  bonus: "Chi thưởng",
  other: "Chi phí khác",
  marketing: "Marketing — Quảng cáo",
  sales_return: "Giảm trừ doanh thu",
};

function OtherExpenseTab({ employee, rows, loading, onChanged }) {
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [form, setForm] = useState({ expense_date: todayStr(), amount: "", description: "", payment_method: "bank_transfer", bank_account_id: null });
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const { banks } = usePaymentOptions();

  const canManage = employee.role !== "nhan_vien";
  const list = rows.filter((r) => r.category === "other");
  const total = list.reduce((s, r) => s + Number(r.amount || 0), 0);

  const openNew = () => {
    setEditing(null);
    setForm({ expense_date: todayStr(), amount: "", description: "", payment_method: "bank_transfer", bank_account_id: null });
    setError(""); setShowForm(true);
  };
  const openEdit = (r) => {
    setEditing(r);
    setForm({ expense_date: r.expense_date, amount: String(r.amount), description: r.description || "",
      payment_method: r.payment_method || "bank_transfer", bank_account_id: r.bank_account_id || null });
    setError(""); setShowForm(true);
  };

  const submit = async () => {
    const amt = Number(form.amount) || 0;
    if (amt <= 0) { setError("Số tiền phải lớn hơn 0."); return; }
    if (!form.description.trim()) { setError("Vui lòng nhập diễn giải."); return; }
    setSaving(true); setError("");
    if (form.payment_method === "bank_transfer" && !form.bank_account_id) {
      setSaving(false); setError("Vui lòng chọn tài khoản chi tiền."); return;
    }
    const payload = {
      expense_date: form.expense_date, amount: amt,
      description: form.description.trim(), category: "other",
      payment_method: form.payment_method,
      bank_account_id: form.payment_method === "bank_transfer" ? form.bank_account_id : null,
      store_id: employee.store_id, created_by: employee.id,
    };
    const { error: err } = editing
      ? await supabase.from("expenses").update(payload).eq("id", editing.id)
      : await supabase.from("expenses").insert(payload);
    setSaving(false);
    if (err) { setError(err.message); return; }
    setShowForm(false); onChanged();
  };

  const remove = async (r) => {
    if (!confirm(`Xóa chi phí ${r.expense_code} — ${fmtVND(r.amount)}?`)) return;
    const { error: err } = await supabase.from("expenses").delete().eq("id", r.id);
    if (err) { alert(err.message); return; }
    onChanged();
  };

  return (
    <div>
      <div className="flex items-center justify-between mb-3 gap-2 flex-wrap">
        <div className="bg-slate-100 rounded-xl px-3 py-2 text-sm">
          <span className="text-slate-400">Tổng trong kỳ: </span>
          <span className="font-semibold text-slate-700">{fmtVND(total)}</span>
          <span className="text-slate-400"> · {list.length} khoản</span>
        </div>
        {canManage && (
          <button onClick={openNew} className="bg-brand-600 hover:bg-brand-700 text-white rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-1.5">
            <Plus size={15} /> Thêm chi phí
          </button>
        )}
      </div>

      {showForm && (
        <Card className="p-4 mb-3">
          <p className="text-sm font-medium text-slate-700 mb-3">{editing ? `Sửa ${editing.expense_code}` : "Thêm chi phí khác"}</p>
          <div className="grid sm:grid-cols-2 gap-3">
            <TextField label="Ngày phát sinh" type="date" value={form.expense_date}
              onChange={(e) => setForm((f) => ({ ...f, expense_date: e.target.value }))} />
            <TextField label="Số tiền (đ) *" type="number" value={form.amount}
              onChange={(e) => setForm((f) => ({ ...f, amount: e.target.value }))} />
          </div>
          <div className="grid sm:grid-cols-2 gap-3 mt-3">
            <div>
              <span className="text-xs font-medium text-slate-500 mb-1 block">Tài khoản chi *</span>
              <BankSelect banks={banks} value={form.bank_account_id}
                onChange={(v) => setForm((f) => ({ ...f, bank_account_id: v }))}
                className="w-full !text-sm !px-3 !py-2 !rounded-xl" />
              <p className="text-[11px] text-slate-400 mt-1">Mọi khoản chi ra đều qua chuyển khoản.</p>
            </div>
          </div>
          <div className="mt-3">
            <TextField label="Diễn giải *" value={form.description}
              placeholder="Tiền điện, tiền nước, sửa chữa, phí ship..."
              onChange={(e) => setForm((f) => ({ ...f, description: e.target.value }))} />
          </div>
          {error && <p className="text-xs text-rose-600 bg-rose-50 rounded-lg px-3 py-2 mt-3">{error}</p>}
          <div className="flex gap-2 mt-3">
            <button onClick={submit} disabled={saving}
              className="bg-brand-600 hover:bg-brand-700 disabled:opacity-60 text-white rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-2">
              {saving && <Loader2 size={15} className="animate-spin" />} {editing ? "Lưu" : "Thêm"}
            </button>
            <button onClick={() => setShowForm(false)} className="px-4 rounded-xl border border-slate-200 text-sm text-slate-500 hover:bg-slate-50">Hủy</button>
          </div>
        </Card>
      )}

      <Card className="p-0 overflow-hidden">
        {loading ? (
          <div className="flex justify-center py-10"><Loader2 className="animate-spin text-slate-300" /></div>
        ) : list.length === 0 ? (
          <EmptyState icon={Receipt} text="Chưa có chi phí nào trong kỳ." />
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead><tr className="text-left text-xs text-slate-400 border-b border-slate-100">
                <th className="px-3 py-2">Mã</th>
                <th className="px-3 py-2">Ngày</th>
                <th className="px-3 py-2">Diễn giải</th>
                <th className="px-3 py-2">Hình thức</th>
                <th className="px-3 py-2 text-right">Số tiền</th>
                <th className="px-3 py-2"></th>
              </tr></thead>
              <tbody>
                {list.map((r) => (
                  <tr key={r.id} className="border-b border-slate-50 last:border-0">
                    <td className="px-3 py-2.5 text-slate-400 doc-code whitespace-nowrap">{r.expense_code}</td>
                    <td className="px-3 py-2.5 text-slate-500 whitespace-nowrap">{fmtDate(r.expense_date)}</td>
                    <td className="px-3 py-2.5 text-slate-700">{r.description}</td>
                    <td className="px-3 py-2.5 whitespace-nowrap">
                      <span className={classNames("text-xs px-2 py-0.5 rounded-full",
                        r.payment_method === "cash" ? "bg-emerald-50 text-emerald-700" : "bg-sky-50 text-sky-700")}>
                        {r.payment_method === "cash" ? "Tiền mặt" : "Chuyển khoản"}
                      </span>
                    </td>
                    <td className="px-3 py-2.5 text-right font-medium text-slate-700 whitespace-nowrap">{fmtVND(r.amount)}</td>
                    <td className="px-3 py-2.5 text-right whitespace-nowrap">
                      {canManage && (<>
                        <button onClick={() => openEdit(r)} className="text-slate-400 hover:text-brand-600 mr-2"><Pencil size={14} /></button>
                        <button onClick={() => remove(r)} className="text-slate-400 hover:text-rose-600"><Trash2 size={14} /></button>
                      </>)}
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

function MarketingTab({ employee, rows, onChanged }) {
  const [month, setMonth] = useState(todayStr().slice(0, 7));
  const [amount, setAmount] = useState("");
  const [note, setNote] = useState("");
  const [preview, setPreview] = useState(null);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState("");
  const [okCode, setOkCode] = useState(null);

  const canManage = employee.role !== "nhan_vien";
  const list = rows.filter((r) => r.category === "marketing");
  const total = list.reduce((s, r) => s + Number(r.amount || 0), 0);

  const doPreview = async () => {
    const amt = Number(amount) || 0;
    if (amt <= 0) { setError("Nhập số tiền lớn hơn 0."); return; }
    setBusy(true); setError(""); setOkCode(null);
    const { data, error: err } = await supabase.rpc("preview_marketing_allocation", {
      p_month: `${month}-01`, p_total: amt,
    });
    setBusy(false);
    if (err) { setError(err.message); return; }
    setPreview(data || []);
  };

  const doAllocate = async () => {
    setBusy(true); setError("");
    const { data, error: err } = await supabase.rpc("allocate_marketing_cost", {
      p_month: `${month}-01`, p_total: Number(amount), p_description: note.trim() || null,
    });
    setBusy(false);
    if (err) { setError(err.message); return; }
    setOkCode(data); setPreview(null); setAmount(""); setNote("");
    onChanged();
  };

  return (
    <div className="space-y-3">
      {canManage && (
        <Card className="p-4">
          <p className="text-sm font-medium text-slate-700 mb-1">Nhập chi phí quảng cáo toàn hệ thống</p>
          <p className="text-xs text-slate-400 mb-3">
            Số tổng sẽ tự chia cho các cửa hàng theo tỷ lệ lãi gộp (doanh thu trừ giá vốn) của tháng đã chọn.
          </p>
          <div className="grid sm:grid-cols-3 gap-3">
            <TextField label="Tháng" type="month" value={month} onChange={(e) => { setMonth(e.target.value); setPreview(null); }} />
            <TextField label="Tổng chi phí (đ) *" type="number" value={amount} onChange={(e) => { setAmount(e.target.value); setPreview(null); }} />
            <TextField label="Diễn giải" value={note} placeholder="Quảng cáo Facebook tháng 8" onChange={(e) => setNote(e.target.value)} />
          </div>

          {error && <p className="text-xs text-rose-600 bg-rose-50 rounded-lg px-3 py-2 mt-3">{error}</p>}
          {okCode && (
            <p className="text-xs text-emerald-700 bg-emerald-50 rounded-lg px-3 py-2 mt-3">
              Đã phân bổ xong, mã đợt <span className="font-semibold">{okCode}</span>.
            </p>
          )}

          {preview && (
            <div className="mt-3 bg-slate-50 rounded-xl p-3">
              <p className="text-xs font-medium text-slate-600 mb-2">Xem trước phân bổ</p>
              <table className="w-full text-sm">
                <thead><tr className="text-left text-xs text-slate-400">
                  <th className="py-1">Cửa hàng</th>
                  <th className="py-1 text-right">Lãi gộp</th>
                  <th className="py-1 text-right">Tỷ lệ</th>
                  <th className="py-1 text-right">Phân bổ</th>
                </tr></thead>
                <tbody>
                  {preview.map((p) => (
                    <tr key={p.store_name} className="border-t border-slate-200">
                      <td className="py-1.5 text-slate-700">{p.store_name}</td>
                      <td className="py-1.5 text-right text-slate-500">{fmtVND(p.gross_profit)}</td>
                      <td className="py-1.5 text-right text-slate-500">{p.ratio}%</td>
                      <td className="py-1.5 text-right font-medium text-slate-700">{fmtVND(p.allocated)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
              {preview.every((p) => Number(p.gross_profit) <= 0) && (
                <p className="text-[11px] text-amber-600 mt-2">
                  Tháng này không cửa hàng nào có lãi gộp dương — hệ thống sẽ chia đều.
                </p>
              )}
            </div>
          )}

          <div className="flex gap-2 mt-3">
            <button onClick={doPreview} disabled={busy}
              className="border border-brand-300 text-brand-700 hover:bg-brand-50 disabled:opacity-60 rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-2">
              {busy && <Loader2 size={15} className="animate-spin" />} Xem trước
            </button>
            {preview && (
              <button onClick={doAllocate} disabled={busy}
                className="bg-brand-600 hover:bg-brand-700 disabled:opacity-60 text-white rounded-xl px-4 py-2 text-sm font-medium">
                Xác nhận phân bổ
              </button>
            )}
          </div>
        </Card>
      )}

      <div className="bg-slate-100 rounded-xl px-3 py-2 text-sm inline-block">
        <span className="text-slate-400">Marketing phân bổ cho cửa hàng này trong kỳ: </span>
        <span className="font-semibold text-slate-700">{fmtVND(total)}</span>
      </div>

      <Card className="p-0 overflow-hidden">
        {list.length === 0 ? (
          <EmptyState icon={TrendingUp} text="Chưa có chi phí marketing nào trong kỳ." />
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead><tr className="text-left text-xs text-slate-400 border-b border-slate-100">
                <th className="px-3 py-2">Mã</th>
                <th className="px-3 py-2">Kỳ</th>
                <th className="px-3 py-2">Diễn giải</th>
                <th className="px-3 py-2 text-right">Phân bổ</th>
              </tr></thead>
              <tbody>
                {list.map((r) => (
                  <tr key={r.id} className="border-b border-slate-50 last:border-0">
                    <td className="px-3 py-2.5 text-slate-400 doc-code whitespace-nowrap">{r.expense_code}</td>
                    <td className="px-3 py-2.5 text-slate-500 whitespace-nowrap">{fmtDate(r.expense_date)}</td>
                    <td className="px-3 py-2.5 text-slate-600 text-xs">{r.description}</td>
                    <td className="px-3 py-2.5 text-right font-medium text-slate-700 whitespace-nowrap">{fmtVND(r.amount)}</td>
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

function ReturnTab({ employee, rows, onChanged }) {
  const [query, setQuery] = useState("");
  const [results, setResults] = useState(null);
  const [searching, setSearching] = useState(false);
  const [picked, setPicked] = useState(null);

  const list = rows.filter((r) => r.category === "sales_return");
  const total = list.reduce((s, r) => s + Number(r.amount || 0), 0);

  const search = async () => {
    const q = query.trim();
    if (q.length < 3) { setResults([]); return; }
    setSearching(true);
    const esc = q.replace(/[%,()]/g, "");
    const { data } = await supabase
      .from("v_returnable_sales").select("*")
      .or(`imei.ilike.%${esc}%,customer_cccd.ilike.%${esc}%,customer_name.ilike.%${esc}%,order_code.ilike.%${esc}%`)
      .order("sold_at", { ascending: false }).limit(20);
    setResults(data || []);
    setSearching(false);
  };

  return (
    <div className="space-y-3">
      <Card className="p-4">
        <p className="text-sm font-medium text-slate-700 mb-1">Nhận máy khách trả</p>
        <p className="text-xs text-slate-400 mb-3">
          Tra theo số IMEI, số CCCD, họ tên khách hoặc mã đơn. Máy sẽ nhập lại kho theo đúng giá vốn cũ,
          phần chênh so với giá bán ghi thành chi phí giảm trừ doanh thu.
        </p>
        <div className="flex gap-2">
          <div className="relative flex-1">
            <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-300" />
            <input
              value={query}
              onChange={(e) => setQuery(e.target.value)}
              onKeyDown={(e) => e.key === "Enter" && search()}
              placeholder="IMEI, CCCD, họ tên hoặc mã đơn..."
              className="w-full pl-9 pr-3 py-2 rounded-lg border border-slate-300 bg-white text-sm placeholder:text-slate-400 outline-none focus:ring-2 focus:ring-brand-400/40 focus:border-brand-500 transition"
            />
          </div>
          <button onClick={search} disabled={searching}
            className="bg-brand-600 hover:bg-brand-700 disabled:opacity-60 text-white rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-2">
            {searching && <Loader2 size={15} className="animate-spin" />} Tra cứu
          </button>
        </div>

        {results !== null && (
          <div className="mt-3">
            {results.length === 0 ? (
              <p className="text-xs text-slate-400 py-3 text-center">
                Không tìm thấy đơn nào khớp. Kiểm tra lại IMEI hoặc thông tin khách.
              </p>
            ) : (
              <div className="space-y-2">
                {results.map((r) => (
                  <div key={r.order_id} className="flex items-center gap-3 border border-slate-200 rounded-xl px-3 py-2.5">
                    <div className="flex-1 min-w-0">
                      <p className="text-sm font-medium text-slate-700">
                        {[r.model, r.storage, r.color].filter(Boolean).join(" ")}
                      </p>
                      <p className="text-xs text-slate-400">
                        IMEI {r.imei || "—"} · {r.order_code} · bán {fmtDate(r.sold_at)}
                      </p>
                      <p className="text-xs text-slate-400">
                        {r.customer_name}{r.customer_phone ? ` · ${r.customer_phone}` : ""}
                        {r.customer_cccd ? ` · CCCD ${r.customer_cccd}` : ""}
                      </p>
                    </div>
                    <div className="text-right shrink-0">
                      <p className="text-sm font-semibold text-slate-700">{fmtVND(r.total_amount)}</p>
                      <button onClick={() => setPicked(r)}
                        className="text-xs text-brand-600 hover:underline">
                        Nhận trả máy
                      </button>
                    </div>
                  </div>
                ))}
              </div>
            )}
          </div>
        )}
      </Card>

      <div className="bg-rose-50 rounded-xl px-3 py-2 text-sm inline-block">
        <span className="text-rose-500">Giảm trừ doanh thu trong kỳ: </span>
        <span className="font-semibold text-rose-700">{fmtVND(total)}</span>
        <span className="text-rose-400"> · {list.length} lượt trả</span>
      </div>

      <Card className="p-0 overflow-hidden">
        {list.length === 0 ? (
          <EmptyState icon={ArrowLeftRight} text="Chưa có lượt trả máy nào trong kỳ." />
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead><tr className="text-left text-xs text-slate-400 border-b border-slate-100">
                <th className="px-3 py-2">Mã</th>
                <th className="px-3 py-2">Ngày</th>
                <th className="px-3 py-2">Diễn giải</th>
                <th className="px-3 py-2 text-right">Giảm trừ</th>
              </tr></thead>
              <tbody>
                {list.map((r) => (
                  <tr key={r.id} className="border-b border-slate-50 last:border-0">
                    <td className="px-3 py-2.5 text-slate-400 doc-code whitespace-nowrap">{r.expense_code}</td>
                    <td className="px-3 py-2.5 text-slate-500 whitespace-nowrap">{fmtDate(r.expense_date)}</td>
                    <td className="px-3 py-2.5 text-slate-600 text-xs">{r.description}</td>
                    <td className="px-3 py-2.5 text-right font-medium text-rose-600 whitespace-nowrap">{fmtVND(r.amount)}</td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Card>

      {picked && (
        <ConfirmReturnModal
          sale={picked} employee={employee}
          onClose={() => setPicked(null)}
          onDone={() => { setPicked(null); setResults(null); setQuery(""); onChanged(); }}
        />
      )}
    </div>
  );
}

function ConfirmReturnModal({ sale, employee, onClose, onDone }) {
  const [method] = useState("bank_transfer");
  const [bankId, setBankId] = useState(null);
  const [reason, setReason] = useState("");
  const [returnDate, setReturnDate] = useState(todayStr());
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const [okCode, setOkCode] = useState(null);
  const { banks } = usePaymentOptions();

  const refund = Number(sale.paid_amount || 0);
  const debtLeft = Math.max(0, Number(sale.total_amount) - refund);

  const submit = async () => {
    if (method === "bank_transfer" && !bankId) { setError("Vui lòng chọn tài khoản chuyển hoàn."); return; }
    setSaving(true); setError("");
    const { data, error: err } = await supabase.rpc("create_sales_return", {
      p_sales_order_id: sale.order_id, p_refund_method: method,
      p_bank_account_id: bankId, p_reason: reason.trim() || null, p_return_date: returnDate,
    });
    setSaving(false);
    if (err) { setError(err.message); return; }
    setOkCode(data);
    await supabase.from("audit_logs").insert({
      table_name: "sales_returns", record_id: sale.order_id, action: "create",
      new_data: { ma: data, imei: sale.imei, hoan: refund },
      performed_by: employee.id, store_id: employee.store_id,
    });
  };

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/40 flex items-center justify-center p-4">
      <Card className="w-full max-w-md p-5">
        <div className="flex items-center justify-between mb-4">
          <p className="font-semibold text-slate-800 text-sm">Nhận máy khách trả</p>
          <button onClick={onClose} className="text-slate-400 hover:text-rose-600"><X size={16} /></button>
        </div>

        {okCode ? (
          <div className="space-y-3">
            <div className="bg-emerald-50 border border-emerald-200 rounded-xl px-3 py-3 text-sm text-emerald-800 space-y-1">
              <p>Đã lập chứng từ trả hàng <span className="font-semibold">{okCode}</span>.</p>
              <p className="text-xs">Máy đã về kho ở trạng thái Còn hàng theo giá vốn cũ. Phần chênh lệch đã ghi vào chi phí giảm trừ doanh thu.</p>
            </div>
            <button onClick={onDone} className="w-full bg-brand-600 hover:bg-brand-700 text-white rounded-xl py-2.5 text-sm font-medium">Xong</button>
          </div>
        ) : (
          <div className="space-y-3">
            <div className="bg-slate-50 rounded-xl px-3 py-2.5 text-xs space-y-0.5">
              <p className="font-medium text-slate-700">{[sale.model, sale.storage, sale.color].filter(Boolean).join(" ")}</p>
              <p className="text-slate-400">IMEI {sale.imei || "—"} · {sale.order_code}</p>
              <p className="text-slate-400">{sale.customer_name} · bán ngày {fmtDate(sale.sold_at)}</p>
              <div className="flex justify-between pt-1 mt-1 border-t border-slate-200">
                <span className="text-slate-400">Giá đã bán</span>
                <span className="font-medium text-slate-700">{fmtVND(sale.total_amount)}</span>
              </div>
              <div className="flex justify-between">
                <span className="text-slate-400">Khách đã trả — sẽ hoàn lại</span>
                <span className="font-medium text-slate-700">{fmtVND(refund)}</span>
              </div>
              {debtLeft > 0 && (
                <div className="flex justify-between text-amber-700">
                  <span>Công nợ còn lại — sẽ được xóa</span>
                  <span className="font-medium">{fmtVND(debtLeft)}</span>
                </div>
              )}
            </div>

            <TextField label="Ngày nhận máy" type="date" value={returnDate} onChange={(e) => setReturnDate(e.target.value)} />

            <div>
              <span className="text-xs font-medium text-slate-500 mb-1 block">Tài khoản chuyển hoàn *</span>
              <BankSelect banks={banks} value={bankId} onChange={setBankId} className="w-full !text-sm !px-3 !py-2 !rounded-xl" />
              <p className="text-[11px] text-slate-400 mt-1">Mọi khoản chi ra đều qua chuyển khoản.</p>
            </div>
            <TextField label="Lý do trả máy" value={reason} placeholder="Máy lỗi, khách đổi ý..." onChange={(e) => setReason(e.target.value)} />

            <p className="text-[11px] text-slate-400">
              Đơn gốc được giữ nguyên để tra cứu, hệ thống sinh chứng từ trả hàng riêng đối ứng lại.
            </p>

            {error && <p className="text-xs text-rose-600 bg-rose-50 rounded-lg px-3 py-2">{error}</p>}
            <div className="flex gap-2">
              <button onClick={submit} disabled={saving}
                className="flex-1 bg-brand-600 hover:bg-brand-700 disabled:opacity-60 text-white rounded-xl py-2.5 text-sm font-medium flex items-center justify-center gap-2">
                {saving && <Loader2 size={15} className="animate-spin" />} Xác nhận nhận máy
              </button>
              <button onClick={onClose} className="px-4 rounded-xl border border-slate-200 text-sm text-slate-500 hover:bg-slate-50">Hủy</button>
            </div>
          </div>
        )}
      </Card>
    </div>
  );
}

function BonusTab({ employee, rows, loading, onChanged, fromDate, toDate }) {
  const [staff, setStaff] = useState([]);
  const [showForm, setShowForm] = useState(false);
  const [editing, setEditing] = useState(null);
  const [form, setForm] = useState({
    expense_date: todayStr(), employee_id: "", amount: "", description: "", bank_account_id: null,
  });
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const { banks } = usePaymentOptions();

  const canManage = employee.role !== "nhan_vien";
  const list = rows.filter((r) => r.category === "bonus");
  const total = list.reduce((s, r) => s + Number(r.amount || 0), 0);

  useEffect(() => {
    (async () => {
      const { data } = await supabase.from("employees")
        .select("id, full_name, role").eq("is_active", true).order("full_name");
      setStaff(data || []);
    })();
  }, []);

  const nameOf = (id) => staff.find((s) => s.id === id)?.full_name || "—";

  const openNew = () => {
    setEditing(null);
    setForm({ expense_date: todayStr(), employee_id: "", amount: "", description: "", bank_account_id: null });
    setError(""); setShowForm(true);
  };
  const openEdit = (r) => {
    setEditing(r);
    setForm({
      expense_date: r.expense_date, employee_id: r.employee_id || "",
      amount: String(r.amount), description: r.description || "",
      bank_account_id: r.bank_account_id || null,
    });
    setError(""); setShowForm(true);
  };

  const submit = async () => {
    const amt = Number(form.amount) || 0;
    if (!form.employee_id) { setError("Vui lòng chọn nhân viên."); return; }
    if (amt <= 0) { setError("Số tiền thưởng phải lớn hơn 0."); return; }
    if (!form.bank_account_id) { setError("Vui lòng chọn tài khoản chi."); return; }
    setSaving(true); setError("");
    const payload = {
      expense_date: form.expense_date, amount: amt, category: "bonus",
      employee_id: form.employee_id,
      description: form.description.trim() || `Thưởng ngày ${fmtDate(form.expense_date)} — ${nameOf(form.employee_id)}`,
      payment_method: "bank_transfer", bank_account_id: form.bank_account_id,
      store_id: employee.store_id, created_by: employee.id,
    };
    const { error: err } = editing
      ? await supabase.from("expenses").update(payload).eq("id", editing.id)
      : await supabase.from("expenses").insert(payload);
    setSaving(false);
    if (err) { setError(err.message); return; }
    setShowForm(false); onChanged();
  };

  const remove = async (r) => {
    if (!confirm(`Xóa khoản thưởng ${fmtVND(r.amount)} ngày ${fmtDate(r.expense_date)}?`)) return;
    const { error: err } = await supabase.from("expenses").delete().eq("id", r.id);
    if (err) { alert(err.message); return; }
    onChanged();
  };

  // Gom theo ngày
  const byDay = {};
  for (const r of list) {
    (byDay[r.expense_date] = byDay[r.expense_date] || []).push(r);
  }
  const days = Object.keys(byDay).sort().reverse();

  // Gom theo nhân viên
  const byStaff = {};
  for (const r of list) {
    const k = r.employee_id || "none";
    byStaff[k] = (byStaff[k] || 0) + Number(r.amount || 0);
  }
  const staffRows = Object.entries(byStaff).sort((a, b) => b[1] - a[1]);

  return (
    <div>
      <div className="flex items-center justify-between mb-3 gap-2 flex-wrap">
        <div className="bg-slate-100 rounded-xl px-3 py-2 text-sm">
          <span className="text-slate-400">Tổng thưởng trong kỳ: </span>
          <span className="font-semibold text-slate-700">{fmtVND(total)}</span>
          <span className="text-slate-400"> · {list.length} lượt · {days.length} ngày</span>
        </div>
        {canManage && (
          <button onClick={openNew} className="bg-brand-600 hover:bg-brand-700 text-white rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-1.5">
            <Plus size={15} /> Chi thưởng
          </button>
        )}
      </div>

      {staffRows.length > 0 && (
        <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 mb-3">
          {staffRows.map(([id, amt]) => (
            <Card key={id} className="p-3">
              <p className="text-xs text-slate-400 mb-1 truncate">{id === "none" ? "Không rõ" : nameOf(id)}</p>
              <p className="text-base font-semibold text-slate-800">{fmtVND(amt)}</p>
            </Card>
          ))}
        </div>
      )}

      {showForm && (
        <Card className="p-4 mb-3">
          <p className="text-sm font-medium text-slate-700 mb-3">
            {editing ? `Sửa khoản thưởng ${editing.expense_code}` : "Chi thưởng cho nhân viên"}
          </p>
          <div className="grid sm:grid-cols-2 gap-3">
            <TextField label="Ngày chi thưởng *" type="date" value={form.expense_date}
              onChange={(e) => setForm((f) => ({ ...f, expense_date: e.target.value }))} />
            <label className="block">
              <span className="text-xs font-medium text-slate-500 mb-1 block">Nhân viên *</span>
              <select value={form.employee_id}
                onChange={(e) => setForm((f) => ({ ...f, employee_id: e.target.value }))}
                className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm text-slate-800 outline-none focus:ring-2 focus:ring-brand-400/40 focus:border-brand-500 transition">
                <option value="">— Chọn nhân viên —</option>
                {staff.map((s) => (
                  <option key={s.id} value={s.id}>{s.full_name} ({ROLE_LABELS[s.role] || s.role})</option>
                ))}
              </select>
            </label>
            <TextField label="Số tiền thưởng (đ) *" type="number" value={form.amount}
              onChange={(e) => setForm((f) => ({ ...f, amount: e.target.value }))} />
            <div>
              <span className="text-xs font-medium text-slate-500 mb-1 block">Tài khoản chi *</span>
              <BankSelect banks={banks} value={form.bank_account_id}
                onChange={(v) => setForm((f) => ({ ...f, bank_account_id: v }))}
                className="w-full !text-sm !px-3 !py-2 !rounded-xl" />
            </div>
          </div>
          <div className="mt-3">
            <TextField label="Lý do thưởng" value={form.description}
              placeholder="Thưởng doanh số ngày, thưởng chăm chỉ..."
              onChange={(e) => setForm((f) => ({ ...f, description: e.target.value }))} />
          </div>
          <p className="text-[11px] text-slate-400 mt-2">
            Khoản thưởng tính là chi phí của đúng ngày chọn ở trên, vào ngay báo cáo kết quả kinh doanh của ngày đó.
          </p>
          {error && <p className="text-xs text-rose-600 bg-rose-50 rounded-lg px-3 py-2 mt-3">{error}</p>}
          <div className="flex gap-2 mt-3">
            <button onClick={submit} disabled={saving}
              className="bg-brand-600 hover:bg-brand-700 disabled:opacity-60 text-white rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-2">
              {saving && <Loader2 size={15} className="animate-spin" />} {editing ? "Lưu" : "Chi thưởng"}
            </button>
            <button onClick={() => setShowForm(false)} className="px-4 rounded-xl border border-slate-200 text-sm text-slate-500 hover:bg-slate-50">Hủy</button>
          </div>
        </Card>
      )}

      <Card className="p-0 overflow-hidden">
        {loading ? (
          <div className="flex justify-center py-10"><Loader2 className="animate-spin text-slate-300" /></div>
        ) : list.length === 0 ? (
          <EmptyState icon={Award} text="Chưa có khoản thưởng nào trong kỳ." />
        ) : (
          <div className="divide-y divide-slate-100">
            {days.map((day) => {
              const items = byDay[day];
              const dayTotal = items.reduce((s, r) => s + Number(r.amount || 0), 0);
              return (
                <div key={day}>
                  <div className="flex items-center justify-between px-3 py-2 bg-slate-50/70">
                    <p className="text-xs font-medium text-slate-600">{fmtDate(day)}</p>
                    <p className="text-xs text-slate-500">{items.length} lượt · <span className="font-semibold text-slate-700">{fmtVND(dayTotal)}</span></p>
                  </div>
                  <table className="w-full text-sm">
                    <tbody>
                      {items.map((r) => (
                        <tr key={r.id} className="border-b border-slate-50 last:border-0">
                          <td className="px-3 py-2.5 text-slate-400 text-xs whitespace-nowrap w-24">{r.expense_code}</td>
                          <td className="px-3 py-2.5 text-slate-700 whitespace-nowrap">{nameOf(r.employee_id)}</td>
                          <td className="px-3 py-2.5 text-slate-500 text-xs">{r.description}</td>
                          <td className="px-3 py-2.5 text-right font-medium text-slate-700 whitespace-nowrap">{fmtVND(r.amount)}</td>
                          <td className="px-3 py-2.5 text-right whitespace-nowrap w-20">
                            {canManage && (<>
                              <button onClick={() => openEdit(r)} className="text-slate-400 hover:text-brand-600 mr-2"><Pencil size={14} /></button>
                              <button onClick={() => remove(r)} className="text-slate-400 hover:text-rose-600"><Trash2 size={14} /></button>
                            </>)}
                          </td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              );
            })}
          </div>
        )}
      </Card>
    </div>
  );
}

function ExpensesModule({ employee }) {
  const [tab, setTab] = useState("other");
  const [fromDate, setFromDate] = useState(startOfMonth());
  const [toDate, setToDate] = useState(todayStr());
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);

  const load = useCallback(async () => {
    setLoading(true);
    const { data } = await supabase.from("expenses").select("*")
      .gte("expense_date", fromDate).lte("expense_date", toDate)
      .order("expense_date", { ascending: false }).limit(2000);
    setRows(data || []);
    setLoading(false);
  }, [fromDate, toDate]);

  useEffect(() => { load(); }, [load]);

  const sumOf = (c) => rows.filter((r) => r.category === c).reduce((s, r) => s + Number(r.amount || 0), 0);
  const grand = rows.reduce((s, r) => s + Number(r.amount || 0), 0);

  const TABS = [
    ["other", "Chi phí khác", sumOf("other")],
    ["bonus", "Chi thưởng", sumOf("bonus")],
    ["marketing", "Marketing", sumOf("marketing")],
    ["sales_return", "Giảm trừ doanh thu", sumOf("sales_return")],
  ];

  return (
    <div>
      <div className="flex items-center justify-between mb-4 gap-2 flex-wrap">
        <div>
          <h2 className="text-lg font-semibold text-slate-800">Chi phí</h2>
          <p className="text-xs text-slate-400">Tổng chi phí trong kỳ: {fmtVND(grand)}</p>
        </div>
        <div className="flex items-center gap-2 bg-white border border-slate-200 rounded-xl px-3 py-1.5">
          <Filter size={14} className="text-slate-400" />
          <input type="date" value={fromDate} onChange={(e) => setFromDate(e.target.value)} className="text-sm outline-none" />
          <span className="text-slate-300">—</span>
          <input type="date" value={toDate} onChange={(e) => setToDate(e.target.value)} className="text-sm outline-none" />
        </div>
      </div>

      <div className="flex gap-2 mb-4 flex-wrap">
        {TABS.map(([k, label, amt]) => (
          <button key={k} onClick={() => setTab(k)}
            className={classNames("px-4 py-2 rounded-xl text-sm border text-left",
              tab === k ? "bg-brand-600 text-white border-brand-600" : "bg-white text-slate-600 border-slate-200 hover:bg-slate-50")}>
            <span className="font-medium">{label}</span>
            <span className={classNames("ml-2 text-xs", tab === k ? "text-white/80" : "text-slate-400")}>{fmtVND(amt)}</span>
          </button>
        ))}
      </div>

      {tab === "other" && <OtherExpenseTab employee={employee} rows={rows} loading={loading} onChanged={load} />}
      {tab === "bonus" && <BonusTab employee={employee} rows={rows} loading={loading} onChanged={load} fromDate={fromDate} toDate={toDate} />}
      {tab === "marketing" && <MarketingTab employee={employee} rows={rows} onChanged={load} />}
      {tab === "sales_return" && <ReturnTab employee={employee} rows={rows} onChanged={load} />}
    </div>
  );
}

/* -------------------------------------------------------------- */
/* Spa / Sửa chữa — thay đổi giá vốn máy + kho màn hình            */
/* -------------------------------------------------------------- */

const SERVICE_TYPE_LABELS = { spa: "Spa", screen_replace: "Thay màn" };
const SERVICE_STATUS_LABELS = { in_progress: "Đang xử lý", done: "Hoàn tất", cancelled: "Đã hủy" };
const SERVICE_STATUS_STYLES = {
  in_progress: "bg-amber-50 text-amber-700",
  done: "bg-emerald-50 text-emerald-700",
  cancelled: "bg-slate-100 text-slate-400",
};
const SCREEN_STATUS_LABELS = {
  in_stock: "Còn trong kho", installed: "Đã lắp máy", sold: "Đã bán", scrapped: "Đã loại",
};
const SCREEN_STATUS_STYLES = {
  in_stock: "bg-emerald-50 text-emerald-700",
  installed: "bg-sky-50 text-sky-700",
  sold: "bg-slate-100 text-slate-500",
  scrapped: "bg-rose-50 text-rose-600",
};
const SCREEN_GRADES = ["Zin bóc máy", "Zin ép kính", "Lô loại 1", "Lô loại 2", "Không rõ"];

function SendServiceForm({ employee, onCancel, onSaved }) {
  const [devices, setDevices] = useState([]);
  const [search, setSearch] = useState("");
  const [picked, setPicked] = useState(null);
  const [type, setType] = useState("spa");
  const [vendor, setVendor] = useState(null);
  const [sentAt, setSentAt] = useState(todayStr());
  const [note, setNote] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const canSeeCost = employee.role !== "nhan_vien";

  useEffect(() => {
    (async () => {
      const { data } = await supabase.from("devices")
        .select("id, imei, model, storage, color, cost_price")
        .eq("status", "in_stock").order("created_at", { ascending: false }).limit(1000);
      setDevices(data || []);
    })();
  }, []);

  const filtered = !search.trim() ? devices.slice(0, 25) : devices.filter((d) => {
    const q = search.trim().toLowerCase();
    return [d.imei, d.model, d.color, d.storage].some((v) => v?.toLowerCase().includes(q));
  }).slice(0, 25);

  const submit = async () => {
    if (!picked) { setError("Vui lòng chọn máy."); return; }
    setSaving(true); setError("");
    const { data, error: err } = await supabase.rpc("create_service_ticket", {
      p_device_id: picked.id, p_service_type: type,
      p_vendor_id: vendor?.id || null, p_note: note.trim() || null, p_sent_at: sentAt,
    });
    setSaving(false);
    if (err) { setError(err.message); return; }
    await supabase.from("audit_logs").insert({
      table_name: "service_tickets", record_id: picked.id, action: "create",
      new_data: { ma: data, loai: type, imei: picked.imei },
      performed_by: employee.id, store_id: employee.store_id,
    });
    onSaved(data);
  };

  return (
    <Card className="p-4 mb-3">
      <p className="text-sm font-medium text-slate-700 mb-3">Gửi máy đi Spa / Sửa chữa</p>
      <div className="space-y-3">
        <div>
          <span className="text-xs font-medium text-slate-500 mb-1 block">Chọn máy trong kho *</span>
          {picked ? (
            <div className="flex items-center justify-between bg-slate-50 rounded-xl px-3 py-2.5">
              <div>
                <p className="text-sm font-medium text-slate-700">{[picked.model, picked.storage, picked.color].filter(Boolean).join(" ")}</p>
                <p className="text-xs text-slate-400">
                  IMEI {picked.imei || "—"}{canSeeCost && <> · giá vốn hiện tại {fmtVND(picked.cost_price)}</>}
                </p>
              </div>
              <button onClick={() => setPicked(null)} className="text-xs text-brand-600 hover:underline">Đổi</button>
            </div>
          ) : (
            <>
              <div className="relative">
                <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-300" />
                <input value={search} onChange={(e) => setSearch(e.target.value)}
                  placeholder="Tìm IMEI, model, màu..."
                  className="w-full pl-9 pr-3 py-2 rounded-lg border border-slate-300 bg-white text-sm placeholder:text-slate-400 outline-none focus:ring-2 focus:ring-brand-400/40 focus:border-brand-500 transition" />
              </div>
              <div className="mt-2 max-h-48 overflow-y-auto border border-slate-200 rounded-xl divide-y divide-slate-100">
                {filtered.length === 0 ? (
                  <p className="text-xs text-slate-400 py-4 text-center">Không có máy Còn hàng nào khớp.</p>
                ) : filtered.map((d) => (
                  <button key={d.id} onClick={() => setPicked(d)} className="w-full text-left px-3 py-2 hover:bg-slate-50">
                    <p className="text-sm text-slate-700">{[d.model, d.storage, d.color].filter(Boolean).join(" ")}</p>
                    <p className="text-xs text-slate-400">IMEI {d.imei || "—"}{canSeeCost && <> · {fmtVND(d.cost_price)}</>}</p>
                  </button>
                ))}
              </div>
            </>
          )}
        </div>

        <div className="grid sm:grid-cols-2 gap-3">
          <label className="block">
            <span className="text-xs font-medium text-slate-500 mb-1 block">Loại dịch vụ *</span>
            <select value={type} onChange={(e) => setType(e.target.value)}
              className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm text-slate-800 outline-none focus:ring-2 focus:ring-brand-400/40 focus:border-brand-500 transition">
              <option value="spa">Spa — làm đẹp máy (giá vốn tăng)</option>
              <option value="screen_replace">Thay màn (giá vốn tăng hoặc giảm)</option>
            </select>
          </label>
          <TextField label="Ngày gửi" type="date" value={sentAt} onChange={(e) => setSentAt(e.target.value)} />
        </div>

        <div>
          <span className="text-xs font-medium text-slate-500 mb-1 block">
            Đơn vị {type === "spa" ? "Spa" : "cung cấp màn"}
          </span>
          <SupplierPicker value={vendor} onSelect={setVendor} employee={employee} />
          <p className="text-[11px] text-slate-400 mt-1">
            Chi phí sẽ treo thành công nợ phải trả đơn vị này. Để trống nếu tự làm, không phát sinh công nợ.
          </p>
        </div>

        <TextField label="Ghi chú" value={note} onChange={(e) => setNote(e.target.value)}
          placeholder="Ép kính, thay vỏ, đánh bóng..." />

        {error && <p className="text-xs text-rose-600 bg-rose-50 rounded-lg px-3 py-2">{error}</p>}
        <div className="flex gap-2">
          <button onClick={submit} disabled={saving}
            className="bg-brand-600 hover:bg-brand-700 disabled:opacity-60 text-white rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-2">
            {saving && <Loader2 size={15} className="animate-spin" />} Gửi máy đi
          </button>
          <button onClick={onCancel} className="px-4 rounded-xl border border-slate-200 text-sm text-slate-500 hover:bg-slate-50">Hủy</button>
        </div>
      </div>
    </Card>
  );
}

function CompleteServiceModal({ ticket, employee, onClose, onDone }) {
  const isSpa = ticket.service_type === "spa";
  const [spaCost, setSpaCost] = useState("");
  const [oldValue, setOldValue] = useState("");
  const [newCost, setNewCost] = useState("");
  const [useStock, setUseStock] = useState(false);
  const [stockScreens, setStockScreens] = useState([]);
  const [stockId, setStockId] = useState("");
  const [oldGrade, setOldGrade] = useState("");
  const [newGrade, setNewGrade] = useState("");
  const [doneAt, setDoneAt] = useState(todayStr());
  const [note, setNote] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const [okCost, setOkCost] = useState(null);

  useEffect(() => {
    if (isSpa) return;
    (async () => {
      const { data } = await supabase.from("v_screens").select("*")
        .eq("status", "in_stock").order("created_at", { ascending: false }).limit(200);
      setStockScreens(data || []);
    })();
  }, [isSpa]);

  const before = Number(ticket.cost_before || 0);
  const chosen = stockScreens.find((s) => s.id === stockId);
  const effNew = useStock ? Number(chosen?.unit_price || 0) : (Number(newCost) || 0);
  const after = isSpa ? before + (Number(spaCost) || 0)
                      : Math.max(0, before - (Number(oldValue) || 0) + effNew);

  const submit = async () => {
    setSaving(true); setError("");
    let res;
    if (isSpa) {
      const c = Number(spaCost) || 0;
      if (c < 0) { setSaving(false); setError("Chi phí spa không hợp lệ."); return; }
      res = await supabase.rpc("complete_spa_ticket", {
        p_ticket_id: ticket.id, p_spa_cost: c,
        p_note: note.trim() || null, p_done_at: doneAt,
      });
    } else {
      const ov = Number(oldValue) || 0;
      if (String(oldValue).trim() === "") { setSaving(false); setError("Vui lòng nhập đơn giá màn cũ."); return; }
      if (useStock && !stockId) { setSaving(false); setError("Vui lòng chọn màn trong kho."); return; }
      if (!useStock && String(newCost).trim() === "") { setSaving(false); setError("Vui lòng nhập đơn giá màn mới."); return; }
      res = await supabase.rpc("complete_screen_ticket", {
        p_ticket_id: ticket.id,
        p_old_screen_value: ov,
        p_new_screen_cost: useStock ? null : (Number(newCost) || 0),
        p_new_screen_id: useStock ? stockId : null,
        p_old_screen_grade: oldGrade || null,
        p_new_screen_grade: newGrade || null,
        p_note: note.trim() || null, p_done_at: doneAt,
      });
    }
    setSaving(false);
    if (res.error) { setError(res.error.message); return; }
    setOkCost(res.data);
    await supabase.from("audit_logs").insert({
      table_name: "service_tickets", record_id: ticket.id, action: "update",
      new_data: { ma: ticket.ticket_code, gia_von_moi: res.data },
      performed_by: employee.id, store_id: employee.store_id,
    });
  };

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/40 flex items-center justify-center p-4 overflow-y-auto">
      <Card className="w-full max-w-lg p-5 my-8">
        <div className="flex items-center justify-between mb-4">
          <p className="font-semibold text-slate-800 text-sm">
            Hoàn tất {SERVICE_TYPE_LABELS[ticket.service_type]} — {ticket.ticket_code}
          </p>
          <button onClick={onClose} className="text-slate-400 hover:text-rose-600"><X size={16} /></button>
        </div>

        {okCost !== null ? (
          <div className="space-y-3">
            <div className="bg-emerald-50 border border-emerald-200 rounded-xl px-3 py-3 text-sm text-emerald-800 space-y-1">
              <p>Máy đã về kho ở trạng thái Còn hàng.</p>
              <p>Giá vốn mới: <span className="font-semibold">{fmtVND(okCost)}</span></p>
              {!isSpa && <p className="text-xs">Màn cũ đã được ghi vào Kho màn hình.</p>}
            </div>
            <button onClick={onDone} className="w-full bg-brand-600 hover:bg-brand-700 text-white rounded-xl py-2.5 text-sm font-medium">Xong</button>
          </div>
        ) : (
          <div className="space-y-3">
            <div className="bg-slate-50 rounded-xl px-3 py-2.5 text-xs space-y-0.5">
              <p className="font-medium text-slate-700">{[ticket.model, ticket.storage, ticket.color].filter(Boolean).join(" ")}</p>
              <p className="text-slate-400">IMEI {ticket.imei || "—"} · gửi ngày {fmtDate(ticket.sent_at)}</p>
              {ticket.vendor_name && <p className="text-slate-400">Đơn vị: {ticket.vendor_name}</p>}
              <div className="flex justify-between pt-1 mt-1 border-t border-slate-200">
                <span className="text-slate-400">Giá vốn trước</span>
                <span className="font-medium text-slate-700">{fmtVND(before)}</span>
              </div>
            </div>

            <TextField label="Ngày hoàn tất" type="date" value={doneAt} onChange={(e) => setDoneAt(e.target.value)} />

            {isSpa ? (
              <TextField label="Chi phí spa (đ) *" type="number" value={spaCost} onChange={(e) => setSpaCost(e.target.value)} />
            ) : (
              <>
                <div className="grid sm:grid-cols-2 gap-3">
                  <TextField label="Đơn giá màn cũ tháo ra (đ) *" type="number" value={oldValue} onChange={(e) => setOldValue(e.target.value)} />
                  <label className="block">
                    <span className="text-xs font-medium text-slate-500 mb-1 block">Chất lượng màn cũ</span>
                    <select value={oldGrade} onChange={(e) => setOldGrade(e.target.value)}
                      className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm text-slate-800 outline-none focus:ring-2 focus:ring-brand-400/40 focus:border-brand-500 transition">
                      <option value="">— Chọn —</option>
                      {SCREEN_GRADES.map((g) => <option key={g} value={g}>{g}</option>)}
                    </select>
                  </label>
                </div>

                <div className="border-t border-slate-100 pt-3">
                  <label className="flex items-center gap-2 text-xs text-slate-600 mb-2">
                    <input type="checkbox" checked={useStock} onChange={(e) => setUseStock(e.target.checked)} />
                    Lắp lại một màn đang có trong Kho màn hình (không phát sinh công nợ)
                  </label>

                  {useStock ? (
                    <select value={stockId} onChange={(e) => setStockId(e.target.value)}
                      className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm text-slate-800 outline-none focus:ring-2 focus:ring-brand-400/40 focus:border-brand-500 transition">
                      <option value="">— Chọn màn trong kho —</option>
                      {stockScreens.map((s) => (
                        <option key={s.id} value={s.id}>
                          {s.screen_code} · {s.model || "—"} · {s.grade || "chưa đánh giá"} · {fmtVND(s.unit_price)}
                        </option>
                      ))}
                    </select>
                  ) : (
                    <div className="grid sm:grid-cols-2 gap-3">
                      <TextField label="Đơn giá màn mới (đ) *" type="number" value={newCost} onChange={(e) => setNewCost(e.target.value)} />
                      <label className="block">
                        <span className="text-xs font-medium text-slate-500 mb-1 block">Chất lượng màn mới</span>
                        <select value={newGrade} onChange={(e) => setNewGrade(e.target.value)}
                          className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm text-slate-800 outline-none focus:ring-2 focus:ring-brand-400/40 focus:border-brand-500 transition">
                          <option value="">— Chọn —</option>
                          {SCREEN_GRADES.map((g) => <option key={g} value={g}>{g}</option>)}
                        </select>
                      </label>
                    </div>
                  )}
                </div>
              </>
            )}

            <div className={classNames("text-xs px-3 py-2.5 rounded-lg space-y-0.5",
              after > before ? "bg-amber-50 text-amber-800" : after < before ? "bg-emerald-50 text-emerald-800" : "bg-slate-100 text-slate-600")}>
              <div className="flex justify-between"><span>Giá vốn trước</span><span>{fmtVND(before)}</span></div>
              {isSpa ? (
                <div className="flex justify-between"><span>Chi phí spa</span><span>+{fmtVND(Number(spaCost) || 0)}</span></div>
              ) : (
                <>
                  <div className="flex justify-between"><span>Trừ màn cũ</span><span>−{fmtVND(Number(oldValue) || 0)}</span></div>
                  <div className="flex justify-between"><span>Cộng màn mới</span><span>+{fmtVND(effNew)}</span></div>
                </>
              )}
              <div className="flex justify-between font-semibold border-t border-current/20 pt-1 mt-1">
                <span>Giá vốn mới</span><span>{fmtVND(after)}</span>
              </div>
            </div>

            <TextField label="Ghi chú" value={note} onChange={(e) => setNote(e.target.value)} />
            {error && <p className="text-xs text-rose-600 bg-rose-50 rounded-lg px-3 py-2">{error}</p>}
            <div className="flex gap-2">
              <button onClick={submit} disabled={saving}
                className="flex-1 bg-brand-600 hover:bg-brand-700 disabled:opacity-60 text-white rounded-xl py-2.5 text-sm font-medium flex items-center justify-center gap-2">
                {saving && <Loader2 size={15} className="animate-spin" />} Xác nhận hoàn tất
              </button>
              <button onClick={onClose} className="px-4 rounded-xl border border-slate-200 text-sm text-slate-500 hover:bg-slate-50">Hủy</button>
            </div>
          </div>
        )}
      </Card>
    </div>
  );
}

function ScreenStockTab({ employee }) {
  const [historyScreen, setHistoryScreen] = useState(null);
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [status, setStatus] = useState("in_stock");
  const [search, setSearch] = useState("");

  const load = useCallback(async () => {
    setLoading(true);
    const { data } = await supabase.from("v_screens").select("*")
      .order("created_at", { ascending: false }).limit(1000);
    setRows(data || []);
    setLoading(false);
  }, []);

  useEffect(() => { load(); }, [load]);

  const filtered = rows.filter((r) => {
    if (status !== "all" && r.status !== status) return false;
    if (!search.trim()) return true;
    const q = search.trim().toLowerCase();
    return [r.screen_code, r.model, r.grade, r.source_imei, r.installed_imei]
      .some((v) => v?.toLowerCase().includes(q));
  });

  const inStock = rows.filter((r) => r.status === "in_stock");
  const stockValue = inStock.reduce((s, r) => s + Number(r.unit_price || 0), 0);

  return (
    <div className="space-y-3">
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3">
        <Card className="p-3">
          <p className="text-xs text-slate-400 mb-1">Màn còn trong kho</p>
          <p className="text-lg font-semibold text-slate-800">{inStock.length}</p>
        </Card>
        <Card className="p-3">
          <p className="text-xs text-slate-400 mb-1">Giá trị tồn</p>
          <p className="text-lg font-semibold text-emerald-700">{fmtVND(stockValue)}</p>
        </Card>
        <Card className="p-3">
          <p className="text-xs text-slate-400 mb-1">Đã lắp máy</p>
          <p className="text-lg font-semibold text-sky-700">{rows.filter((r) => r.status === "installed").length}</p>
        </Card>
        <Card className="p-3">
          <p className="text-xs text-slate-400 mb-1">Tháo ra từ máy</p>
          <p className="text-lg font-semibold text-slate-600">{rows.filter((r) => r.origin === "removed").length}</p>
        </Card>
      </div>

      <Card className="p-0 overflow-hidden">
        <div className="p-3 border-b border-slate-100 flex gap-2 flex-wrap">
          <div className="relative flex-1 min-w-[180px]">
            <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-300" />
            <input value={search} onChange={(e) => setSearch(e.target.value)}
              placeholder="Tìm mã màn, loại, chất lượng, IMEI..."
              className="w-full pl-9 pr-3 py-2 rounded-lg border border-slate-300 bg-white text-sm placeholder:text-slate-400 outline-none focus:ring-2 focus:ring-brand-400/40 focus:border-brand-500 transition" />
          </div>
          <select value={status} onChange={(e) => setStatus(e.target.value)}
            className="rounded-xl border border-slate-200 px-3 py-2 text-sm">
            <option value="in_stock">Còn trong kho</option>
            <option value="installed">Đã lắp máy</option>
            <option value="sold">Đã bán</option>
            <option value="scrapped">Đã loại</option>
            <option value="all">Tất cả</option>
          </select>
        </div>

        {loading ? (
          <div className="flex justify-center py-10"><Loader2 className="animate-spin text-slate-300" /></div>
        ) : filtered.length === 0 ? (
          <EmptyState icon={Smartphone} text="Chưa có màn hình nào trong kho." />
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead><tr className="text-left text-xs text-slate-400 border-b border-slate-100">
                <th className="px-3 py-2">Mã</th>
                <th className="px-3 py-2">Loại màn</th>
                <th className="px-3 py-2">Chất lượng</th>
                <th className="px-3 py-2">Nguồn</th>
                <th className="px-3 py-2 text-right">Đơn giá</th>
                <th className="px-3 py-2">Trạng thái</th>
                <th className="px-3 py-2"></th>
              </tr></thead>
              <tbody>
                {filtered.map((r) => (
                  <tr key={r.id} className="border-b border-slate-50 last:border-0">
                    <td className="px-3 py-2.5 text-slate-400 doc-code whitespace-nowrap">{r.screen_code}</td>
                    <td className="px-3 py-2.5 text-slate-700">{r.model || "—"}</td>
                    <td className="px-3 py-2.5 text-slate-600 text-xs">{r.grade || "chưa đánh giá"}</td>
                    <td className="px-3 py-2.5 text-xs">
                      {r.origin === "removed" ? (
                        <span className="text-slate-500">Tháo từ {r.source_imei || "máy cũ"}</span>
                      ) : (
                        <span className="text-slate-500">Mua{r.vendor_name ? ` — ${r.vendor_name}` : ""}</span>
                      )}
                      {r.installed_imei && <p className="text-sky-600">Đang lắp ở {r.installed_imei}</p>}
                    </td>
                    <td className="px-3 py-2.5 text-right text-slate-700 whitespace-nowrap">{fmtVND(r.unit_price)}</td>
                    <td className="px-3 py-2.5 whitespace-nowrap">
                      <span className={classNames("text-xs px-2 py-0.5 rounded-full", SCREEN_STATUS_STYLES[r.status])}>
                        {SCREEN_STATUS_LABELS[r.status] || r.status}
                      </span>
                    </td>
                    <td className="px-3 py-2.5 text-right whitespace-nowrap">
                      <button onClick={() => setHistoryScreen(r)} className="text-slate-400 hover:text-brand-600" title="Lý lịch tấm màn">
                        <History size={14} className="inline" />
                      </button>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </Card>

      {historyScreen && (
        <ScreenHistoryPanel screen={historyScreen} onClose={() => setHistoryScreen(null)} />
      )}
    </div>
  );
}

function ScreenPurchaseForm({ employee, onCancel, onSaved }) {
  const [mode, setMode] = useState("single");   // single | batch
  const [vendor, setVendor] = useState(null);
  const [model, setModel] = useState("");
  const [grade, setGrade] = useState("");
  const [unitPrice, setUnitPrice] = useState("");
  const [qty, setQty] = useState("1");
  const [date, setDate] = useState(todayStr());
  const [note, setNote] = useState("");
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");

  const n = mode === "single" ? 1 : (Number(qty) || 0);
  const total = (Number(unitPrice) || 0) * n;

  const submit = async () => {
    if (!model.trim()) { setError("Vui lòng nhập loại màn."); return; }
    if ((Number(unitPrice) || 0) < 0) { setError("Đơn giá không hợp lệ."); return; }
    if (n < 1) { setError("Số lượng phải từ 1 trở lên."); return; }
    setSaving(true); setError("");
    const { data, error: err } = await supabase.rpc("create_screen_purchase", {
      p_vendor_id: vendor?.id || null,
      p_model: model.trim(),
      p_unit_price: Number(unitPrice) || 0,
      p_quantity: n,
      p_grade: grade || null,
      p_note: note.trim() || null,
      p_purchase_date: date,
    });
    setSaving(false);
    if (err) { setError(err.message); return; }
    await supabase.from("audit_logs").insert({
      table_name: "screen_purchases", record_id: null, action: "create",
      new_data: { ma: data, model: model.trim(), sl: n, tong: total },
      performed_by: employee.id, store_id: employee.store_id,
    });
    onSaved(data, n);
  };

  return (
    <Card className="p-4 mb-3">
      <p className="text-sm font-medium text-slate-700 mb-3">Nhập kho màn hình</p>

      <div className="flex gap-2 mb-3">
        {[["single", "Màn rời"], ["batch", "Nhập theo lô"]].map(([k, label]) => (
          <button key={k} onClick={() => { setMode(k); setQty(k === "single" ? "1" : ""); }}
            className={classNames("px-3 py-1.5 rounded-lg text-xs border font-medium",
              mode === k ? "bg-brand-600 text-white border-brand-600" : "bg-white text-slate-500 border-slate-200")}
          >{label}</button>
        ))}
      </div>

      <div className="space-y-3">
        <div className="grid sm:grid-cols-2 gap-3">
          <TextField label="Loại màn *" value={model} onChange={(e) => setModel(e.target.value)}
            placeholder="iPhone 13 Pro" list="dl-models-screen" />
          <label className="block">
            <span className="text-xs font-medium text-slate-500 mb-1 block">Chất lượng</span>
            <select value={grade} onChange={(e) => setGrade(e.target.value)}
              className="w-full rounded-lg border border-slate-300 bg-white px-3 py-2 text-sm text-slate-800 outline-none focus:ring-2 focus:ring-brand-400/40 focus:border-brand-500 transition">
              <option value="">— Chọn —</option>
              {SCREEN_GRADES.map((g) => <option key={g} value={g}>{g}</option>)}
            </select>
          </label>
        </div>
        <datalist id="dl-models-screen">{IPHONE_MODEL_LIST.map((m) => <option key={m} value={m} />)}</datalist>

        <div className="grid sm:grid-cols-3 gap-3">
          <TextField label="Đơn giá 1 tấm (đ) *" type="number" value={unitPrice} onChange={(e) => setUnitPrice(e.target.value)} />
          {mode === "batch" ? (
            <TextField label="Số lượng *" type="number" value={qty} onChange={(e) => setQty(e.target.value)} placeholder="10" />
          ) : (
            <div className="flex items-end pb-2 text-xs text-slate-400">Số lượng: 1 tấm</div>
          )}
          <TextField label="Ngày nhập" type="date" value={date} onChange={(e) => setDate(e.target.value)} />
        </div>

        <div>
          <span className="text-xs font-medium text-slate-500 mb-1 block">Nhà cung cấp</span>
          <SupplierPicker value={vendor} onSelect={setVendor} employee={employee} />
          <p className="text-[11px] text-slate-400 mt-1">
            Tiền hàng treo thành công nợ phải trả đơn vị này. Để trống nếu không theo dõi công nợ.
          </p>
        </div>

        <TextField label="Ghi chú" value={note} onChange={(e) => setNote(e.target.value)} />

        {n > 0 && (Number(unitPrice) || 0) > 0 && (
          <div className="bg-slate-50 rounded-xl px-3 py-2.5 text-xs space-y-0.5">
            <div className="flex justify-between"><span className="text-slate-400">Số tấm</span><span className="text-slate-700">{n}</span></div>
            <div className="flex justify-between"><span className="text-slate-400">Đơn giá</span><span className="text-slate-700">{fmtVND(Number(unitPrice))}</span></div>
            <div className="flex justify-between border-t border-slate-200 pt-1 mt-1">
              <span className="text-slate-500">Tổng tiền</span>
              <span className="font-semibold text-slate-800">{fmtVND(total)}</span>
            </div>
            {n > 1 && <p className="text-[11px] text-slate-400 pt-1">Hệ thống tạo {n} bản ghi màn riêng, mỗi tấm một mã MH để theo dõi lắp vào máy nào.</p>}
          </div>
        )}

        {error && <p className="text-xs text-rose-600 bg-rose-50 rounded-lg px-3 py-2">{error}</p>}
        <div className="flex gap-2">
          <button onClick={submit} disabled={saving}
            className="bg-brand-600 hover:bg-brand-700 disabled:opacity-60 text-white rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-2">
            {saving && <Loader2 size={15} className="animate-spin" />} Nhập kho
          </button>
          <button onClick={onCancel} className="px-4 rounded-xl border border-slate-200 text-sm text-slate-500 hover:bg-slate-50">Hủy</button>
        </div>
      </div>
    </Card>
  );
}

function PayScreenPurchaseModal({ row, employee, onClose, onDone }) {
  const remaining = Number(row.remaining || 0);
  const [amount, setAmount] = useState(String(remaining));
  const [bankId, setBankId] = useState(null);
  const [saving, setSaving] = useState(false);
  const [error, setError] = useState("");
  const { banks } = usePaymentOptions();

  const submit = async () => {
    const v = Number(amount) || 0;
    if (v <= 0) { setError("Số tiền phải lớn hơn 0."); return; }
    if (v > remaining) { setError(`Phiếu này chỉ còn nợ ${fmtVND(remaining)}.`); return; }
    if (!bankId) { setError("Vui lòng chọn tài khoản chuyển tiền."); return; }
    setSaving(true); setError("");
    const { error: err } = await supabase.rpc("pay_screen_purchase", {
      p_purchase_id: row.id, p_amount: v, p_bank_account_id: bankId, p_note: null,
    });
    setSaving(false);
    if (err) { setError(err.message); return; }
    onDone();
  };

  return (
    <div className="fixed inset-0 z-50 bg-slate-900/40 flex items-center justify-center p-4">
      <Card className="w-full max-w-sm p-5">
        <div className="flex items-center justify-between mb-4">
          <p className="font-semibold text-slate-800 text-sm">Thanh toán — {row.purchase_code}</p>
          <button onClick={onClose} className="text-slate-400 hover:text-rose-600"><X size={16} /></button>
        </div>
        <div className="space-y-3">
          <div className="bg-slate-50 rounded-xl px-3 py-2.5 text-xs space-y-0.5">
            <p className="font-medium text-slate-700">{row.model} · {row.quantity} tấm</p>
            <p className="text-slate-400">{row.vendor_name || "Không có NCC"}</p>
            <div className="flex justify-between pt-1 mt-1 border-t border-slate-200">
              <span className="text-slate-400">Còn nợ</span>
              <span className="font-medium text-amber-700">{fmtVND(remaining)}</span>
            </div>
          </div>
          <TextField label="Số tiền thanh toán (đ)" type="number" value={amount} onChange={(e) => setAmount(e.target.value)} />
          <div>
            <span className="text-xs font-medium text-slate-500 mb-1 block">Tài khoản chuyển tiền *</span>
            <BankSelect banks={banks} value={bankId} onChange={setBankId} className="w-full !text-sm !px-3 !py-2 !rounded-xl" />
          </div>
          {error && <p className="text-xs text-rose-600 bg-rose-50 rounded-lg px-3 py-2">{error}</p>}
          <div className="flex gap-2">
            <button onClick={submit} disabled={saving}
              className="flex-1 bg-brand-600 hover:bg-brand-700 disabled:opacity-60 text-white rounded-xl py-2.5 text-sm font-medium flex items-center justify-center gap-2">
              {saving && <Loader2 size={15} className="animate-spin" />} Xác nhận
            </button>
            <button onClick={onClose} className="px-4 rounded-xl border border-slate-200 text-sm text-slate-500 hover:bg-slate-50">Hủy</button>
          </div>
        </div>
      </Card>
    </div>
  );
}

function ScreenPurchaseTab({ employee }) {
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [showForm, setShowForm] = useState(false);
  const [okMsg, setOkMsg] = useState(null);
  const [paying, setPaying] = useState(null);

  const canCreate = true;   // cả 3 vai đều nhập kho màn được, quyền do RPC kiểm tra
  const canPay = employee.role !== "nhan_vien";

  const load = useCallback(async () => {
    setLoading(true);
    const { data } = await supabase.from("v_screen_purchases").select("*")
      .order("created_at", { ascending: false }).limit(500);
    setRows(data || []);
    setLoading(false);
  }, []);

  useEffect(() => { load(); }, [load]);

  const totalDebt = rows.reduce((s, r) => s + Math.max(0, Number(r.remaining || 0)), 0);

  return (
    <div className="space-y-3">
      <div className="flex items-center justify-between gap-2 flex-wrap">
        {totalDebt > 0 ? (
          <div className="bg-amber-50 border border-amber-200 rounded-xl px-3 py-2 text-xs text-amber-800">
            Còn nợ nhà cung cấp màn: <span className="font-medium">{fmtVND(totalDebt)}</span>
          </div>
        ) : <span className="text-xs text-slate-400">{rows.length} phiếu nhập màn</span>}
        {canCreate && (
          <button onClick={() => { setShowForm((s) => !s); setOkMsg(null); }}
            className="bg-brand-600 hover:bg-brand-700 text-white rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-1.5">
            <Plus size={15} /> Nhập màn
          </button>
        )}
      </div>

      {okMsg && <p className="text-xs text-emerald-700 bg-emerald-50 rounded-lg px-3 py-2">{okMsg}</p>}

      {showForm && (
        <ScreenPurchaseForm employee={employee}
          onCancel={() => setShowForm(false)}
          onSaved={(code, n) => {
            setShowForm(false);
            setOkMsg(`Đã nhập kho ${n} tấm màn theo phiếu ${code}.`);
            load();
          }} />
      )}

      <Card className="p-0 overflow-hidden">
        {loading ? (
          <div className="flex justify-center py-10"><Loader2 className="animate-spin text-slate-300" /></div>
        ) : rows.length === 0 ? (
          <EmptyState icon={Smartphone} text="Chưa có phiếu nhập màn nào." />
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-sm">
              <thead><tr className="text-left text-xs text-slate-400 border-b border-slate-100">
                <th className="px-3 py-2">Mã</th>
                <th className="px-3 py-2">Ngày</th>
                <th className="px-3 py-2">Loại màn</th>
                <th className="px-3 py-2">NCC</th>
                <th className="px-3 py-2 text-right">SL / Tồn</th>
                <th className="px-3 py-2 text-right">Tổng tiền</th>
                <th className="px-3 py-2 text-right">Còn nợ</th>
                <th className="px-3 py-2"></th>
              </tr></thead>
              <tbody>
                {rows.map((r) => {
                  const rem = Math.max(0, Number(r.remaining || 0));
                  return (
                    <tr key={r.id} className="border-b border-slate-50 last:border-0">
                      <td className="px-3 py-2.5 text-slate-400 doc-code whitespace-nowrap">{r.purchase_code}</td>
                      <td className="px-3 py-2.5 text-slate-500 whitespace-nowrap">{fmtDate(r.purchase_date)}</td>
                      <td className="px-3 py-2.5">
                        <p className="text-slate-700">{r.model}</p>
                        <p className="text-xs text-slate-400">{r.grade || "chưa đánh giá"}</p>
                      </td>
                      <td className="px-3 py-2.5 text-slate-500 text-xs">{r.vendor_name || "—"}</td>
                      <td className="px-3 py-2.5 text-right text-slate-600 whitespace-nowrap">
                        {r.quantity} / <span className="text-emerald-600">{r.still_in_stock}</span>
                      </td>
                      <td className="px-3 py-2.5 text-right text-slate-700 whitespace-nowrap">{fmtVND(r.total_amount)}</td>
                      <td className="px-3 py-2.5 text-right whitespace-nowrap">
                        {rem > 0
                          ? <span className="font-medium text-amber-600">{fmtVND(rem)}</span>
                          : <span className="text-xs px-2 py-0.5 rounded-full bg-emerald-50 text-emerald-700">Đã trả đủ</span>}
                      </td>
                      <td className="px-3 py-2.5 text-right whitespace-nowrap">
                        {canPay && rem > 0 && (
                          <button onClick={() => setPaying(r)} className="text-xs text-brand-600 hover:underline">Thanh toán</button>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        )}
      </Card>

      {paying && (
        <PayScreenPurchaseModal row={paying} employee={employee}
          onClose={() => setPaying(null)}
          onDone={() => { setPaying(null); load(); }} />
      )}
    </div>
  );
}

function ServiceModule({ employee }) {
  const [tab, setTab] = useState("tickets");
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [showForm, setShowForm] = useState(false);
  const [completing, setCompleting] = useState(null);
  const [okCode, setOkCode] = useState(null);
  const [showDone, setShowDone] = useState(false);

  const canManage = employee.role !== "ke_toan";

  const load = useCallback(async () => {
    setLoading(true);
    const { data } = await supabase.from("v_service_tickets").select("*")
      .order("created_at", { ascending: false }).limit(500);
    setRows(data || []);
    setLoading(false);
  }, []);

  useEffect(() => { load(); }, [load]);

  const open = rows.filter((r) => r.status === "in_progress");
  const shown = showDone ? rows : open;
  const openValue = open.reduce((s, r) => s + Number(r.cost_before || 0), 0);

  const cancel = async (r) => {
    if (!confirm(`Hủy phiếu ${r.ticket_code}? Máy sẽ quay lại kho ở trạng thái Còn hàng.`)) return;
    const { error } = await supabase.rpc("cancel_service_ticket", { p_ticket_id: r.id });
    if (error) { alert(error.message); return; }
    load();
  };

  return (
    <div>
      <div className="flex items-center justify-between mb-4 gap-2 flex-wrap">
        <div>
          <h2 className="text-lg font-semibold text-slate-800">Spa / Sửa chữa</h2>
          <p className="text-xs text-slate-400">
            {tab === "tickets"
              ? `${open.length} máy đang xử lý · giá vốn ${fmtVND(openValue)}`
              : tab === "screens" ? "Màn tháo ra và màn đã mua, theo dõi chất lượng"
              : "Nhập kho màn rời hoặc theo lô, treo công nợ nhà cung cấp"}
          </p>
        </div>
        {tab === "tickets" && canManage && (
          <button onClick={() => { setShowForm((s) => !s); setOkCode(null); }}
            className="bg-brand-600 hover:bg-brand-700 text-white rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-1.5">
            <Plus size={15} /> Gửi máy đi
          </button>
        )}
      </div>

      <div className="flex gap-2 mb-4">
        {[["tickets", "Phiếu dịch vụ"], ["screens", "Kho màn hình"], ["purchases", "Nhập màn"]].map(([k, label]) => (
          <button key={k} onClick={() => setTab(k)}
            className={classNames("px-4 py-2 rounded-xl text-sm border font-medium",
              tab === k ? "bg-brand-600 text-white border-brand-600" : "bg-white text-slate-500 border-slate-200 hover:bg-slate-50")}
          >{label}</button>
        ))}
      </div>

      {tab === "purchases" ? <ScreenPurchaseTab employee={employee} />
       : tab === "screens" ? <ScreenStockTab employee={employee} /> : (
        <>
          {okCode && (
            <p className="text-xs text-emerald-700 bg-emerald-50 rounded-lg px-3 py-2 mb-3">
              Đã lập phiếu <span className="font-semibold">{okCode}</span>. Máy chuyển sang trạng thái Đang spa/sửa, không bán được cho tới khi hoàn tất.
            </p>
          )}

          {showForm && (
            <SendServiceForm employee={employee}
              onCancel={() => setShowForm(false)}
              onSaved={(code) => { setShowForm(false); setOkCode(code); load(); }} />
          )}

          <div className="flex items-center justify-between gap-2 flex-wrap mb-2">
            <p className="text-xs text-slate-400">{open.length} phiếu đang xử lý</p>
            <button onClick={() => setShowDone((s) => !s)} className="text-xs text-brand-600 hover:underline">
              {showDone ? "Chỉ xem đang xử lý" : "Xem cả phiếu đã hoàn tất"}
            </button>
          </div>

          <Card className="p-0 overflow-hidden">
            {loading ? (
              <div className="flex justify-center py-10"><Loader2 className="animate-spin text-slate-300" /></div>
            ) : shown.length === 0 ? (
              <EmptyState icon={Settings} text="Chưa có phiếu spa/sửa chữa nào." />
            ) : (
              <div className="overflow-x-auto">
                <table className="w-full text-sm">
                  <thead><tr className="text-left text-xs text-slate-400 border-b border-slate-100">
                    <th className="px-3 py-2">Mã</th>
                    <th className="px-3 py-2">Máy</th>
                    <th className="px-3 py-2">Loại</th>
                    <th className="px-3 py-2">Đơn vị</th>
                    <th className="px-3 py-2 text-right">Giá vốn</th>
                    <th className="px-3 py-2">Trạng thái</th>
                    <th className="px-3 py-2"></th>
                  </tr></thead>
                  <tbody>
                    {shown.map((r) => {
                      const change = Number(r.cost_after || 0) - Number(r.cost_before || 0);
                      return (
                        <tr key={r.id} className="border-b border-slate-50 last:border-0">
                          <td className="px-3 py-2.5 text-slate-400 doc-code whitespace-nowrap">{r.ticket_code}</td>
                          <td className="px-3 py-2.5">
                            <p className="text-slate-700">{[r.model, r.storage, r.color].filter(Boolean).join(" ")}</p>
                            <p className="text-xs text-slate-400">IMEI {r.imei || "—"} · gửi {fmtDate(r.sent_at)}</p>
                          </td>
                          <td className="px-3 py-2.5 text-slate-600 whitespace-nowrap">{SERVICE_TYPE_LABELS[r.service_type]}</td>
                          <td className="px-3 py-2.5 text-slate-500 text-xs">{r.vendor_name || "Tự làm"}</td>
                          <td className="px-3 py-2.5 text-right whitespace-nowrap">
                            <p className="text-slate-600">{fmtVND(r.cost_before)}</p>
                            {r.status === "done" && (
                              <p className={classNames("text-xs", change > 0 ? "text-amber-600" : change < 0 ? "text-emerald-600" : "text-slate-400")}>
                                → {fmtVND(r.cost_after)} ({change >= 0 ? "+" : ""}{fmtVND(change)})
                              </p>
                            )}
                          </td>
                          <td className="px-3 py-2.5 whitespace-nowrap">
                            <span className={classNames("text-xs px-2 py-0.5 rounded-full", SERVICE_STATUS_STYLES[r.status])}>
                              {SERVICE_STATUS_LABELS[r.status]}
                            </span>
                          </td>
                          <td className="px-3 py-2.5 text-right whitespace-nowrap">
                            {canManage && r.status === "in_progress" && (
                              <>
                                <button onClick={() => setCompleting(r)} className="text-xs text-brand-600 hover:underline mr-3">Hoàn tất</button>
                                <button onClick={() => cancel(r)} className="text-xs text-rose-500 hover:underline">Hủy</button>
                              </>
                            )}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            )}
          </Card>
        </>
      )}

      {completing && (
        <CompleteServiceModal ticket={completing} employee={employee}
          onClose={() => setCompleting(null)}
          onDone={() => { setCompleting(null); load(); }} />
      )}
    </div>
  );
}

const NAV_ITEMS = [
  { key: "dashboard", label: "Tổng quan", icon: LayoutDashboard, group: null },
  { key: "customers", label: "Khách hàng", icon: Users, group: "Bán hàng" },
  { key: "orders", label: "Đơn hàng bán", icon: ShoppingCart, group: "Bán hàng" },
  { key: "inventory", label: "Kho hàng", icon: Smartphone, group: "Kho & dịch vụ" },
  { key: "purchases", label: "Nhập máy / Thu cũ", icon: ArrowLeftRight, group: "Kho & dịch vụ" },
  { key: "service", label: "Spa / Sửa chữa", icon: Settings, group: "Kho & dịch vụ" },
  { key: "debts", label: "Công nợ", icon: Banknote, group: "Tài chính" },
  { key: "expenses", label: "Chi phí", icon: Receipt, group: "Tài chính" },
  { key: "reports", label: "Báo cáo", icon: BarChart3, group: "Tài chính", allowedRoles: ["quan_ly", "ke_toan"] },
  { key: "employees", label: "Nhân viên", icon: UserCog, group: "Hệ thống", managerOnly: true },
  { key: "audit", label: "Nhật ký thao tác", icon: ScrollText, group: "Hệ thống", allowedRoles: ["quan_ly", "ke_toan"] },
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
      case "purchases":
        return <PurchaseModule employee={employee} />;
      case "service":
        return <ServiceModule employee={employee} />;
      case "expenses":
        return <ExpensesModule employee={employee} />;
      case "debts":
        return <DebtModule employee={employee} />;
      case "reports":
        return ["quan_ly", "ke_toan"].includes(employee.role) ? <ReportsModule employee={employee} /> : null;
      case "audit":
        return ["quan_ly", "ke_toan"].includes(employee.role) ? <AuditLogModule /> : null;
      case "employees":
        return employee.role === "quan_ly" ? <EmployeesModule employee={employee} /> : null;
      default:
        return null;
    }
  };

  return (
    <div className="min-h-screen bg-slate-50/70 lg:flex">
      {/* Desktop sidebar */}
      <aside className="hidden lg:flex flex-col w-64 shrink-0 bg-white border-r border-slate-200/80 print:hidden">
        {/* Măng-sét: tên cửa hàng đặt như tiêu đề sổ */}
        <div className="px-5 pt-6 pb-5">
          <p className="text-[10px] uppercase tracking-[0.16em] text-slate-400 mb-1.5">Hệ thống quản lý</p>
          <p className="text-[15px] font-semibold text-brand-800 leading-snug">
            {employee.stores?.name || "Cửa hàng"}
          </p>
          <p className="text-[11px] text-slate-400 mt-0.5">
            {employee.full_name} · {ROLE_LABELS[employee.role] || employee.role}
          </p>
        </div>

        <nav className="flex-1 px-3 pb-4 overflow-y-auto">
          {visibleNav.map((item, i) => {
            const active = tab === item.key;
            const newGroup = item.group && item.group !== visibleNav[i - 1]?.group;
            return (
              <div key={item.key}>
                {newGroup && (
                  <p className="text-[10px] uppercase tracking-[0.14em] text-slate-300 px-3 pt-4 pb-1.5">
                    {item.group}
                  </p>
                )}
                <button
                  onClick={() => setTab(item.key)}
                  aria-current={active ? "page" : undefined}
                  className={classNames(
                    "w-full flex items-center gap-3 pl-3 pr-2 py-2 rounded-lg text-[13.5px] transition-colors relative",
                    active
                      ? "bg-brand-50 text-brand-800 font-medium"
                      : "text-slate-500 hover:text-slate-800 hover:bg-slate-50"
                  )}
                >
                  {active && <span className="absolute left-0 top-1.5 bottom-1.5 w-[3px] rounded-full bg-brand-600" />}
                  <item.icon size={16} className={active ? "text-brand-600" : "text-slate-400"} />
                  <span className="flex-1 text-left">{item.label}</span>
                </button>
              </div>
            );
          })}
        </nav>

        <div className="px-3 pb-4 pt-3 border-t border-slate-100">
          <button onClick={onSignOut}
            className="w-full flex items-center gap-3 px-3 py-2 rounded-lg text-[13.5px] text-slate-400 hover:text-rose-600 hover:bg-rose-50/60 transition-colors">
            <LogOut size={16} /> Đăng xuất
          </button>
        </div>
      </aside>

      {/* Mobile top bar */}
      <div className="lg:hidden flex items-center justify-between px-4 py-3 bg-white border-b border-slate-200/80 sticky top-0 z-20 print:hidden">
        <div>
          <p className="text-sm font-semibold text-brand-800 leading-tight">{employee.stores?.name || "Cửa hàng"}</p>
          <p className="text-[11px] text-slate-400">{ROLE_LABELS[employee.role] || employee.role}</p>
        </div>
        <button onClick={() => setMobileMenuOpen((s) => !s)}
          aria-label="Mở menu" className="text-slate-400 hover:text-brand-700 p-1"><Menu size={20} /></button>
      </div>
      {mobileMenuOpen && (
        <div className="lg:hidden bg-white border-b border-slate-200/80 px-3 py-2 print:hidden">
          {visibleNav.map((item, i) => (
            <div key={item.key}>
              {item.group && item.group !== visibleNav[i - 1]?.group && (
                <p className="text-[10px] uppercase tracking-[0.14em] text-slate-300 px-3 pt-3 pb-1">{item.group}</p>
              )}
              <button
                onClick={() => { setTab(item.key); setMobileMenuOpen(false); }}
                className={classNames(
                  "w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm",
                  tab === item.key ? "bg-brand-50 text-brand-800 font-medium" : "text-slate-500"
                )}
              >
                <item.icon size={16} className={tab === item.key ? "text-brand-600" : "text-slate-400"} />
                <span className="flex-1 text-left">{item.label}</span>
                {tab === item.key && <ChevronRight size={14} />}
              </button>
            </div>
          ))}
          <button onClick={onSignOut} className="w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm text-slate-400">
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

    let { data: byUid } = await supabase.from("employees").select("*, stores(name, address)").eq("user_id", uid).maybeSingle();
    if (byUid) { setEmployee(byUid); return; }

    // First login after being invited: link by email if a placeholder row exists.
    const { data: byEmail } = await supabase.from("employees").select("*, stores(name, address)").eq("email", userEmail).is("user_id", null).maybeSingle();
    if (byEmail) {
      const { data: linked, error } = await supabase.from("employees").update({ user_id: uid }).eq("id", byEmail.id).select("*, stores(name, address)").maybeSingle();
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
