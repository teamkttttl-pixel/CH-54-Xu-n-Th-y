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
};
const PAYMENT_METHOD_ICONS = {
  cash: Wallet,
  bank_transfer: Landmark,
  installment: CalendarClock,
  trade_in: ArrowLeftRight,
  debt_offset: Banknote,
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
  const [historyImei, setHistoryImei] = useState(null);

  const canDelete = employee.role === "quan_ly";
  const canManage = employee.role === "quan_ly";
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

  const cycleStatus = async (d) => {
    if (!canManage) { alert("Bạn chỉ có quyền xem Kho hàng."); return; }
    if (d.status === "sold") { alert("Máy đã bán — không thể tự đổi trạng thái ở đây."); return; }
    const order = ["in_stock", "reserved", "sold"];
    const next = order[(order.indexOf(d.status) + 1) % order.length];
    const { data: updated, error } = await supabase.from("devices").update({ status: next, updated_by: employee.id }).eq("id", d.id).select().maybeSingle();
    if (error) { alert(error.message); return; }
    await supabase.from("audit_logs").insert({
      table_name: "devices", record_id: d.id, action: "update",
      old_data: d, new_data: updated, performed_by: employee.id, store_id: employee.store_id,
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

      {canManage && showForm && editing && (
        <DeviceForm
          initial={editing}
          employee={employee}
          duplicateImei={null}
          onCancel={() => { setShowForm(false); setEditing(null); }}
          onSaved={() => { setShowForm(false); setEditing(null); load(); }}
        />
      )}

      {historyImei && <ImeiHistoryPanel imei={historyImei} onClose={() => setHistoryImei(null)} />}

      <Card className="p-0 overflow-hidden">
        <div className="p-3 border-b border-slate-100 flex flex-col sm:flex-row gap-2">
          <div className="relative flex-1">
            <Search size={15} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-300" />
            <input
              value={search}
              onChange={(e) => setSearch(e.target.value)}
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
            <option value="missing_imei">⚠ Thiếu IMEI</option>
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
                    <td className="px-3 py-2.5 font-medium text-slate-700 whitespace-nowrap">
                      {d.imei || (
                        <span className="inline-flex items-center gap-1 text-xs font-normal px-2 py-0.5 rounded-full bg-amber-50 text-amber-700">
                          <ShieldAlert size={11} /> Thiếu IMEI
                        </span>
                      )}
                    </td>
                    <td className="px-3 py-2.5 text-slate-600">
                      {d.model}
                      <span className="text-slate-400"> {[d.storage, d.color].filter(Boolean).join(" · ")}</span>
                    </td>
                    <td className="px-3 py-2.5 text-slate-500">
                      {DEVICE_CONDITION_LABELS[d.condition] || "—"}
                      {d.condition_percent != null && <span className="text-slate-400"> · {d.condition_percent}%</span>}
                    </td>
                    {canSeeCost && <td className="px-3 py-2.5 text-slate-500">{fmtVND(d.cost_price)}</td>}
                    <td className="px-3 py-2.5 text-slate-500">{fmtVND(d.sale_price)}</td>
                    <td className="px-3 py-2.5">
                      {d.status === "sold" || !canManage ? (
                        <span
                          className={classNames("text-xs px-2 py-0.5 rounded-full", DEVICE_STATUS_STYLES[d.status])}
                          title={d.status === "sold" ? "Máy đã bán — không thể tự đổi trạng thái" : "Bạn chỉ có quyền xem"}
                        >
                          {DEVICE_STATUS_LABELS[d.status] || d.status}
                        </span>
                      ) : (
                        <button
                          onClick={() => cycleStatus(d)}
                          className={classNames("text-xs px-2 py-0.5 rounded-full", DEVICE_STATUS_STYLES[d.status])}
                          title="Bấm để chuyển trạng thái"
                        >
                          {DEVICE_STATUS_LABELS[d.status] || d.status}
                        </button>
                      )}
                    </td>
                    <td className="px-3 py-2.5 text-right whitespace-nowrap">
                      {d.imei && (
                        <button onClick={() => setHistoryImei(d.imei)} className="text-slate-400 hover:text-brand-600 mr-3" title="Lịch sử IMEI">
                          <History size={14} className="inline" />
                        </button>
                      )}
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
  const offsetUsed = rows.filter((r) => r.method === "debt_offset")
    .reduce((s, r) => s + (Number(r.amount) || 0), 0);
  const offsetLeft = supplierDebt - offsetUsed;

  const update = (i, patch) => setRows((rs) => rs.map((r, idx) => (idx === i ? { ...r, ...patch } : r)));
  const addRow = () => setRows((rs) => [...rs, {
    method: "cash", amount: remaining > 0 ? remaining : "", bank_account_id: null,
    installment_provider: "", installment_contract_code: "", note: "",
    trade_in_imei: "", trade_in_model: "", trade_in_storage: "", trade_in_color: "", trade_in_condition: "used",
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
                <option value="trade_in">Đổi máy cũ</option>
                {supplierDebt > 0 && <option value="debt_offset">Bù trừ công nợ</option>}
              </select>
              <input
                type="number"
                value={r.amount}
                onChange={(e) => update(i, { amount: e.target.value })}
                placeholder={r.method === "trade_in" ? "Giá thu mua" : r.method === "debt_offset" ? "Số tiền cấn trừ" : "Số tiền"}
                className="w-32 rounded-lg border border-slate-200 px-2 py-1.5 text-sm"
              />
              {rows.length > 1 && (
                <button type="button" onClick={() => removeRow(i)} className="text-rose-400 hover:text-rose-600"><X size={15} /></button>
              )}
            </div>
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
            {r.method === "trade_in" && (
              <div className="pl-6 space-y-2">
                <p className="text-[11px] text-slate-400">Máy khách đổi sẽ tự động nhập vào Kho hàng (Còn hàng) khi hoàn tất đơn.</p>
                <div className="grid grid-cols-2 gap-2">
                  <input
                    value={r.trade_in_imei}
                    onChange={(e) => update(i, { trade_in_imei: e.target.value })}
                    placeholder="Số IMEI máy đổi *"
                    className="rounded-lg border border-slate-200 px-2 py-1.5 text-xs"
                  />
                  <input
                    value={r.trade_in_model}
                    onChange={(e) => update(i, { trade_in_model: e.target.value })}
                    placeholder="Model máy đổi *"
                    list={`dl-models-tradein-${i}`}
                    className="rounded-lg border border-slate-200 px-2 py-1.5 text-xs"
                  />
                  <input
                    value={r.trade_in_storage}
                    onChange={(e) => update(i, { trade_in_storage: e.target.value })}
                    placeholder="Dung lượng"
                    list="dl-storage"
                    className="rounded-lg border border-slate-200 px-2 py-1.5 text-xs"
                  />
                  <input
                    value={r.trade_in_color}
                    onChange={(e) => update(i, { trade_in_color: e.target.value })}
                    placeholder="Màu sắc"
                    list={`dl-colors-tradein-${i}`}
                    className="rounded-lg border border-slate-200 px-2 py-1.5 text-xs"
                  />
                  <datalist id={`dl-models-tradein-${i}`}>{IPHONE_MODEL_LIST.map((m) => <option key={m} value={m} />)}</datalist>
                  <datalist id={`dl-colors-tradein-${i}`}>{coloroptionsForModel(r.trade_in_model).map((c) => <option key={c} value={c} />)}</datalist>
                </div>
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
        remaining === 0 ? "bg-emerald-50 text-emerald-700"
          : remaining > 0 ? "bg-amber-50 text-amber-700"
          : "bg-rose-50 text-rose-600"
      )}>
        Đã phân bổ {fmtVND(paid)} / {fmtVND(total)}
        {remaining > 0 && ` — khách còn nợ ${fmtVND(remaining)}`}
        {remaining < 0 && ` — thừa ${fmtVND(-remaining)}, vui lòng giảm bớt`}
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
    trade_in_imei: "", trade_in_model: "", trade_in_storage: "", trade_in_color: "", trade_in_condition: "used",
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
  const paidNow = payments.reduce((s, r) => s + (Number(r.amount) || 0), 0);

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
    const activePayments = payments.filter((r) => String(r.amount).trim() !== "");
    const paidTotal = activePayments.reduce((s, r) => s + (Number(r.amount) || 0), 0);
    const tradeInTotal = activePayments.filter((r) => r.method === "trade_in")
      .reduce((s, r) => s + (Number(r.amount) || 0), 0);
    const nonTradeIn = paidTotal - tradeInTotal;
    if (Math.round(nonTradeIn) > Math.round(total)) {
      setError("Tiền mặt/chuyển khoản/trả góp đang vượt tổng tiền đơn hàng."); return;
    }
    if (Math.round(paidTotal) > Math.round(total) && tradeInTotal <= 0) {
      setError("Tổng các hình thức thanh toán đang lớn hơn tổng tiền đơn hàng."); return;
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
    const tradeInRows = activePayments.filter((r) => r.method === "trade_in");
    for (const r of tradeInRows) {
      if (!r.trade_in_imei?.trim() || !r.trade_in_model?.trim()) {
        setError("Vui lòng nhập đủ IMEI và model cho máy khách đổi.");
        return;
      }
    }
    // Kiểm tra IMEI máy đổi không trùng với kho hiện có, không trùng nhau giữa các dòng
    const imeis = tradeInRows.map((r) => r.trade_in_imei.trim());
    if (new Set(imeis).size !== imeis.length) { setError("Các IMEI máy đổi bị trùng nhau."); return; }
    if (tradeInRows.length > 0) {
      const { data: dupCheck } = await supabase.from("devices").select("imei").in("imei", imeis);
      if (dupCheck && dupCheck.length > 0) {
        setError(`IMEI máy đổi "${dupCheck[0].imei}" đã tồn tại trong kho.`);
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
        due_date: Math.round(paidTotal) < Math.round(total) ? (dueDate || null) : null,
        status: manualDeviceMode ? "pending_stock" : "completed",
        created_by: employee.id,
        updated_by: employee.id,
        store_id: employee.store_id,
      };
      const { data: order, error: orderErr } = await supabase.from("sales_orders").insert(orderPayload).select().maybeSingle();
      if (orderErr) throw orderErr;

      // Với mỗi dòng "Đổi máy cũ": tạo máy mới vào Kho + phiếu thu mua liên kết đơn này
      const tradeInDeviceIds = {};
      for (const r of tradeInRows) {
        const { data: newDevice, error: newDevErr } = await supabase.from("devices").insert({
          imei: r.trade_in_imei.trim(), model: r.trade_in_model.trim(),
          storage: r.trade_in_storage?.trim() || null, color: r.trade_in_color?.trim() || null,
          condition: "used", status: "in_stock", cost_price: Number(r.amount),
          supplier: `Thu đổi từ khách — đơn ${order.order_code}`, import_date: new Date().toISOString().slice(0, 10),
          created_by: employee.id, updated_by: employee.id, store_id: employee.store_id,
        }).select().maybeSingle();
        if (newDevErr) throw newDevErr;
        tradeInDeviceIds[r.trade_in_imei.trim()] = newDevice.id;

        const { data: purchase, error: poErr } = await supabase.from("purchase_orders").insert({
          source_type: "customer",
          customer_id: customer.id, device_id: newDevice.id, linked_sale_order_id: order.id,
          purchase_price: Number(r.amount), payment_method: "trade_in", created_by: employee.id,
          paid_amount: 0, store_id: employee.store_id,
        }).select().maybeSingle();
        if (poErr) throw poErr;

        const { error: pcontractErr } = await supabase.from("contracts").insert({ purchase_order_id: purchase.id, created_by: employee.id, store_id: employee.store_id });
        if (pcontractErr) throw pcontractErr;

        await supabase.from("audit_logs").insert({
          table_name: "devices", record_id: newDevice.id, action: "create", new_data: newDevice, performed_by: employee.id, store_id: employee.store_id,
        });
      }

      const paymentRows = activePayments.map((r) => ({
        order_id: order.id,
        method: r.method,
        amount: Number(r.amount),
        bank_account_id: r.method === "bank_transfer" ? (r.bank_account_id || null) : null,
        installment_provider: r.method === "installment" ? (r.installment_provider.trim() || null) : null,
        installment_contract_code: r.method === "installment" ? (r.installment_contract_code.trim() || null) : null,
        trade_in_device_id: r.method === "trade_in" ? tradeInDeviceIds[r.trade_in_imei.trim()] : null,
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
        {paidNow > total && (
          <div className="bg-indigo-50 border border-indigo-200 rounded-xl px-3 py-2.5 text-xs text-indigo-800 flex items-start gap-2">
            <Banknote size={15} className="shrink-0 mt-0.5" />
            <span>
              Máy khách đổi có giá trị cao hơn máy bán — cửa hàng sẽ nợ khách{" "}
              <span className="font-medium">{fmtVND(paidNow - total)}</span>.
              Khoản này hiện ở danh sách đơn hàng để trả lại hoặc trừ vào lần mua sau.
            </span>
          </div>
        )}
        {paidNow < total && (
          <div className="bg-amber-50 border border-amber-200 rounded-xl p-3 space-y-2">
            <p className="text-xs text-amber-800">
              Khách còn nợ <span className="font-medium">{fmtVND(total - paidNow)}</span> — đơn sẽ được ghi nhận là công nợ và thu sau ở màn Đơn hàng bán.
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
            <select value={form.condition} onChange={set("condition")} className="w-full rounded-xl border border-slate-200 px-3 py-2 text-sm">
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
            className="w-full rounded-xl border border-slate-200 px-3 py-2 text-sm"
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
  const [method, setMethod] = useState("cash");
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
            <span className="text-xs font-medium text-slate-500 mb-1 block">Hình thức chi</span>
            <select
              value={method} onChange={(e) => setMethod(e.target.value)}
              className="w-full rounded-xl border border-slate-200 px-3 py-2 text-sm"
            >
              <option value="cash">Tiền mặt</option>
              <option value="bank_transfer">Chuyển khoản</option>
            </select>
          </div>
          {method === "bank_transfer" && (
            <div>
              <span className="text-xs font-medium text-slate-500 mb-1 block">Tài khoản chuyển tiền *</span>
              <BankSelect banks={banks} value={bankId} onChange={setBankId} className="w-full !text-sm !px-3 !py-2 !rounded-xl" />
            </div>
          )}
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

function OrdersModule({ employee }) {
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
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Purchase module — nhập máy/thu cũ độc lập (mua trả tiền khách)         */
/* ---------------------------------------------------------------------- */

function PrintPurchaseModal({ type, purchase, customer, device, contract, storeName, onClose }) {
  const today = fmtDate(new Date());
  const isContract = type === "contract";
  return (
    <div className="fixed inset-0 z-50 bg-slate-900/40 flex items-start sm:items-center justify-center p-4 overflow-y-auto print:bg-white print:p-0">
      <div className="bg-white rounded-2xl print:rounded-none shadow-xl w-full max-w-2xl p-6 sm:p-10 my-6 print:my-0 print:shadow-none print:max-w-none">
        <div className="flex items-center justify-end gap-2 mb-4 print:hidden">
          <button onClick={() => window.print()} className="bg-brand-600 hover:bg-brand-700 text-white rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-1.5">
            <Printer size={15} /> In
          </button>
          <button onClick={onClose} className="text-slate-400 hover:text-rose-600 rounded-xl px-3 py-2 text-sm">Đóng</button>
        </div>
        <div className="text-center mb-2">
          <p className="text-xs text-slate-400 uppercase tracking-wide">Cộng hòa xã hội chủ nghĩa Việt Nam</p>
          <p className="text-[11px] text-slate-400">Độc lập — Tự do — Hạnh phúc</p>
        </div>
        <div className="text-center mb-6">
          <p className="text-xs text-slate-400 mt-3">{storeName || "Cửa hàng"} — Quản lý mua bán điện thoại</p>
          <h2 className="text-lg font-bold text-slate-800 mt-1">
            {isContract ? "HỢP ĐỒNG MUA BÁN TÀI SẢN (THU MUA MÁY CŨ)" : "PHIẾU CHI TIỀN (THU MUA MÁY CŨ)"}
          </h2>
          <p className="text-xs text-slate-400 mt-1">
            Số: {isContract ? contract?.contract_code : purchase.purchase_code} — Ngày {today}
          </p>
        </div>

        {isContract && (
          <div className="text-sm text-slate-700 space-y-3 mb-4">
            <div>
              <p className="font-medium text-slate-800 mb-1">BÊN A (Bên mua):</p>
              <p className="pl-3">Cửa hàng {storeName || ""} — Kinh doanh mua bán điện thoại di động</p>
              <p className="pl-3 text-slate-500">Đại diện: {"{Tên đại diện cửa hàng}"} — Chức vụ: {"{Chức vụ}"}</p>
            </div>
          </div>
        )}
        <div className="text-sm text-slate-700 space-y-1 mb-4">
          {isContract ? (
            <p className="font-medium text-slate-800 mb-1">BÊN B (Bên bán):</p>
          ) : (
            <p><span className="text-slate-400">Khách hàng (bên bán):</span> {customer?.full_name}</p>
          )}
          <div className={isContract ? "pl-3 space-y-0.5" : ""}>
            {isContract && <p><span className="text-slate-400">Họ và tên:</span> {customer?.full_name}</p>}
            {customer?.date_of_birth && <p><span className="text-slate-400">Ngày sinh:</span> {fmtDate(customer.date_of_birth)}</p>}
            {customer?.phone && <p><span className="text-slate-400">SĐT:</span> {customer.phone}</p>}
            {customer?.cccd && <p><span className="text-slate-400">Số CCCD:</span> {customer.cccd} {customer.cccd_issue_date && `— cấp ngày ${fmtDate(customer.cccd_issue_date)}`} {customer.cccd_issue_place && `tại ${customer.cccd_issue_place}`}</p>}
            {customer?.address && <p><span className="text-slate-400">Địa chỉ thường trú:</span> {customer.address}</p>}
          </div>
        </div>
        <div className="text-sm text-slate-700 space-y-1 mb-4 border-t border-dashed border-slate-200 pt-4">
          <p className="font-medium text-slate-800">{isContract ? "Điều 1. Đối tượng hợp đồng" : "Thông tin máy thu mua"}</p>
          <p className={isContract ? "pl-3" : ""}>Điện thoại {device?.model} {[device?.storage, device?.color].filter(Boolean).join(" · ")} — {DEVICE_CONDITION_LABELS[device?.condition]}</p>
          <p className={isContract ? "pl-3" : ""}><span className="text-slate-400">Số IMEI:</span> {device?.imei || "Chưa có IMEI"}</p>
        </div>
        <div className="text-sm text-slate-700 space-y-1 mb-6 border-t border-dashed border-slate-200 pt-4">
          {isContract && <p className="font-medium text-slate-800">Điều 2. Giá cả và phương thức thanh toán</p>}
          <div className={classNames("flex justify-between", isContract && "pl-3")}><span className="text-slate-400">Hình thức chi trả</span><span>{PAYMENT_METHOD_LABELS[purchase.payment_method]}</span></div>
          <div className={classNames("flex justify-between font-semibold text-slate-800", isContract && "pl-3")}><span>Số tiền {isContract ? "thu mua (đã bao gồm mọi chi phí)" : "chi"}</span><span>{fmtVND(purchase.purchase_price)}</span></div>
        </div>

        {isContract && (
          <div className="text-xs text-slate-600 space-y-2.5 mb-6 border-t border-dashed border-slate-200 pt-4">
            <div>
              <p className="font-medium text-slate-700 mb-1">Điều 3. Cam kết của Bên B</p>
              <p className="pl-3">3.1. Tài sản nêu trên thuộc quyền sở hữu hợp pháp, đầy đủ của Bên B, không có tranh chấp, không phải tài sản do phạm tội mà có, không đang bị cầm cố, thế chấp hay thuộc diện tranh chấp với bên thứ ba nào.</p>
              <p className="pl-3">3.2. Thông tin cá nhân và giấy tờ tùy thân (CCCD) cung cấp cho Bên A là đúng sự thật, Bên B chịu trách nhiệm trước pháp luật về tính chính xác của các thông tin này.</p>
            </div>
            <div>
              <p className="font-medium text-slate-700 mb-1">Điều 4. Chuyển giao quyền sở hữu</p>
              <p className="pl-3">Ngay sau khi Bên A thanh toán đủ số tiền tại Điều 2 và Bên B bàn giao tài sản, quyền sở hữu đối với tài sản nêu trên chuyển giao hoàn toàn, không điều kiện cho Bên A.</p>
            </div>
            <div>
              <p className="font-medium text-slate-700 mb-1">Điều 5. Giải quyết tranh chấp</p>
              <p className="pl-3">Mọi tranh chấp phát sinh (nếu có) được hai bên ưu tiên giải quyết trên tinh thần thương lượng, hòa giải. Trường hợp không thỏa thuận được, tranh chấp sẽ được đưa ra cơ quan có thẩm quyền giải quyết theo quy định pháp luật.</p>
            </div>
            <div>
              <p className="font-medium text-slate-700 mb-1">Điều 6. Hiệu lực hợp đồng</p>
              <p className="pl-3">Hợp đồng có hiệu lực kể từ thời điểm hai bên ký tên, lập thành bộ hồ sơ gồm hợp đồng này kèm bản sao/ảnh chụp CCCD của Bên B.</p>
            </div>
          </div>
        )}

        <div className="grid grid-cols-2 gap-4 text-center text-sm text-slate-600 mt-10">
          <div><p className="font-medium text-slate-700 mb-12">{isContract ? "Bên B (Khách hàng)" : "Khách hàng"}</p><p className="text-xs text-slate-400">(Ký, ghi rõ họ tên)</p></div>
          <div><p className="font-medium text-slate-700 mb-12">{isContract ? "Bên A (Cửa hàng)" : "Đại diện cửa hàng"}</p><p className="text-xs text-slate-400">(Ký, ghi rõ họ tên)</p></div>
        </div>
      </div>
    </div>
  );
}

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
          className="w-full pl-9 pr-3 py-2 rounded-xl border border-slate-200 text-sm outline-none focus:ring-2 focus:ring-brand-400"
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
    purchase_price: "", payment_method: "cash", bank_account_id: null, notes: "",
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
            <select value={form.condition} onChange={set("condition")} className="w-full rounded-xl border border-slate-200 px-3 py-2 text-sm">
              <option value="used">Máy cũ</option>
              <option value="new">Máy mới</option>
            </select>
          </label>
          <TextField label="Độ mới (%)" type="number" min="0" max="100" value={form.condition_percent} onChange={set("condition_percent")} placeholder="99" />
          <TextField label="Giá thu mua (đ) *" type="number" value={form.purchase_price} onChange={set("purchase_price")} />
        </div>

        <label className="block">
          <span className="text-xs font-medium text-slate-500 mb-1 block">Hình thức chi trả</span>
          <select value={form.payment_method} onChange={set("payment_method")} className="w-full rounded-xl border border-slate-200 px-3 py-2 text-sm">
            <option value="cash">Tiền mặt</option>
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

function PurchaseModule({ employee }) {
  const [purchases, setPurchases] = useState([]);
  const [loading, setLoading] = useState(true);
  const [search, setSearch] = useState("");
  const [showForm, setShowForm] = useState(false);
  const [printData, setPrintData] = useState(null);
  const [payingDebt, setPayingDebt] = useState(null);
  const [payAmount, setPayAmount] = useState("");

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
    // Ghi vào SỔ CÁI qua RPC — cột paid_amount do trigger tự tính lại
    const { data: remain, error } = await supabase.rpc("pay_supplier_debt", {
      p_purchase_order_id: p.id, p_amount: addAmount, p_note: "Trả tiền nhà cung cấp",
    });
    if (error) { alert(error.message); return; }
    await supabase.from("audit_logs").insert({
      table_name: "purchase_orders", record_id: p.id, action: "update",
      old_data: p, new_data: { ...p, con_no: remain }, performed_by: employee.id, store_id: employee.store_id,
    });
    setPayingDebt(null);
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
        {canCreate && (
          <button onClick={() => setShowForm((s) => !s)} className="bg-brand-600 hover:bg-brand-700 text-white rounded-xl px-4 py-2 text-sm font-medium flex items-center gap-1.5">
            <Banknote size={15} /> Nhập máy mới
          </button>
        )}
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
              className="w-full pl-9 pr-3 py-2 rounded-xl border border-slate-200 text-sm outline-none focus:ring-2 focus:ring-brand-400"
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
                      {!p.sales_orders && p.source_type === "customer" && (
                        <button onClick={() => setPrintData({ p, kind: "receipt" })} className="text-brand-600 hover:underline text-xs mr-3">
                          <Printer size={12} className="inline mr-0.5" />Phiếu chi
                        </button>
                      )}
                      {p.source_type === "customer" && (
                        <button onClick={() => setPrintData({ p, kind: "contract" })} className="text-brand-600 hover:underline text-xs mr-3">
                          <Printer size={12} className="inline mr-0.5" />Hợp đồng
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
            <div className="flex gap-2 mt-4">
              <button onClick={submitPayDebt} className="bg-brand-600 hover:bg-brand-700 text-white rounded-xl px-4 py-2 text-sm font-medium">Xác nhận thanh toán</button>
              <button onClick={() => setPayingDebt(null)} className="text-slate-500 text-sm px-4 py-2">Hủy</button>
            </div>
          </Card>
        </div>
      )}

      {printData && (
        <PrintPurchaseModal
          type={printData.kind}
          purchase={printData.p}
          customer={printData.p.customers}
          device={printData.p.devices}
          contract={Array.isArray(printData.p.contracts) ? printData.p.contracts[0] : printData.p.contracts}
          storeName={employee.stores?.name}
          onClose={() => setPrintData(null)}
        />
      )}
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
          storeName={employee.stores?.name}
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


/* -------------------------------------------------------------- */
/* Công nợ đối tác — đọc trực tiếp từ sổ cái                       */
/* -------------------------------------------------------------- */

const LEDGER_TYPE_LABELS = {
  sale: "Bán máy",
  sale_receipt: "Thu tiền khách",
  purchase: "Nhập máy",
  purchase_payment: "Trả tiền NCC",
  offset: "Bù trừ công nợ",
  adjustment: "Điều chỉnh",
};

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
          <h2 className="text-lg font-semibold text-slate-800">Công nợ đối tác</h2>
          <p className="text-xs text-slate-400">{active.length} đối tác đang có số dư</p>
        </div>
        <button onClick={load} className="bg-white border border-slate-200 hover:bg-slate-50 text-slate-600 rounded-xl px-3 py-2 text-sm flex items-center gap-1.5">
          <Loader2 size={14} className={loading ? "animate-spin" : ""} /> Tải lại
        </button>
      </div>

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
              className="w-full pl-9 pr-3 py-2 rounded-xl border border-slate-200 text-sm outline-none focus:ring-2 focus:ring-brand-400"
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

const NAV_ITEMS = [
  { key: "dashboard", label: "Dashboard", icon: LayoutDashboard },
  { key: "customers", label: "Khách hàng", icon: Users },
  { key: "inventory", label: "Kho hàng", icon: Smartphone },
  { key: "orders", label: "Đơn hàng bán", icon: ShoppingCart },
  { key: "purchases", label: "Nhập máy/Thu cũ", icon: ArrowLeftRight },
  { key: "invoices", label: "Hóa đơn", icon: FileText, phase: "Phase 4" },
  { key: "receipts", label: "Phiếu thu/chi", icon: Receipt },
  { key: "debts", label: "Công nợ đối tác", icon: Banknote },
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
      case "purchases":
        return <PurchaseModule employee={employee} />;
      case "receipts":
        return <ReceiptsModule employee={employee} />;
      case "debts":
        return <DebtModule employee={employee} />;
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
            <p className="text-sm font-semibold text-slate-800 leading-tight">{employee.stores?.name || "Cửa hàng"}</p>
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
          <p className="text-sm font-semibold text-slate-800">{employee.stores?.name || "Cửa hàng"}</p>
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
