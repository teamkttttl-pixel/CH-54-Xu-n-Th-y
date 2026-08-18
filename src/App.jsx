import React, { useEffect, useState, useCallback } from "react";
import { supabase } from "./supabaseClient";
import {
  LayoutDashboard, Users, Smartphone, ShoppingCart, Receipt, FileText,
  BarChart3, UserCog, ScrollText, Settings, LogOut, Search, Plus, X,
  Loader2, ChevronRight, Menu, ShieldAlert, Pencil, Trash2, PackageSearch,
  History,
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
  return (
    <div>
      <div className="mb-5">
        <h2 className="text-lg font-semibold text-slate-800">Chào {employee.full_name} 👋</h2>
        <p className="text-sm text-slate-400">Vai trò: {ROLE_LABELS[employee.role] || employee.role}</p>
      </div>
      <div className="grid grid-cols-2 lg:grid-cols-4 gap-3 sm:gap-4">
        <StatCard label="Tổng khách hàng" value={customerCount} icon={Users} />
        <StatCard label="Tồn kho (IMEI)" value={inStockCount} icon={Smartphone} />
        <StatCard label="Đơn hàng hôm nay" icon={ShoppingCart} comingSoon />
        <StatCard label="Doanh thu hôm nay" icon={BarChart3} comingSoon />
      </div>
      <Card className="p-5 mt-4">
        <p className="text-sm font-medium text-slate-700 mb-1">Tiến độ triển khai</p>
        <p className="text-xs text-slate-400 leading-relaxed">
          Phase 1 (Đăng nhập, phân quyền, Khách hàng) và Phase 2 (Kho hàng theo IMEI) đang hoạt động.
          Đơn hàng bán + Phiếu thu/chi + Hợp đồng (Phase 3) sẽ được bổ sung ở lần cập nhật tiếp theo.
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
  { key: "orders", label: "Đơn hàng bán", icon: ShoppingCart, phase: "Phase 3" },
  { key: "invoices", label: "Hóa đơn", icon: FileText, phase: "Phase 3" },
  { key: "receipts", label: "Phiếu thu/chi", icon: Receipt, phase: "Phase 3" },
  { key: "reports", label: "Báo cáo", icon: BarChart3, phase: "Phase 4" },
  { key: "employees", label: "Nhân viên", icon: UserCog, managerOnly: true },
  { key: "audit", label: "Audit Log", icon: ScrollText, phase: "Phase 4" },
];

function AppShell({ employee, onSignOut }) {
  const [tab, setTab] = useState("dashboard");
  const [customerCount, setCustomerCount] = useState(0);
  const [inStockCount, setInStockCount] = useState(0);
  const [mobileMenuOpen, setMobileMenuOpen] = useState(false);

  const visibleNav = NAV_ITEMS.filter((n) => !n.managerOnly || employee.role === "quan_ly");

  const renderModule = () => {
    switch (tab) {
      case "dashboard":
        return <DashboardModule employee={employee} customerCount={customerCount} inStockCount={inStockCount} />;
      case "customers":
        return <CustomersModule employee={employee} onCountChange={setCustomerCount} />;
      case "inventory":
        return <InventoryModule employee={employee} onCountChange={setInStockCount} />;
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
      <aside className="hidden lg:flex flex-col w-60 shrink-0 border-r border-slate-100 bg-white/70 p-4">
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
      <div className="lg:hidden flex items-center justify-between px-4 py-3 bg-white border-b border-slate-100 sticky top-0 z-10">
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
