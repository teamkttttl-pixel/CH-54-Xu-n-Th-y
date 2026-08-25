/* ---------------------------------------------------------------------- */
/*  ReportsExcelModule.jsx                                                 */
/*                                                                         */
/*  Bo bao cao dung theo bieu mau Excel cu cua CH 81 Xa Dan.               */
/*  Muc dich: nhan su cu mo ra thay dung bang quen thuoc, khong phai       */
/*  hoc lai cach doc.                                                      */
/*                                                                         */
/*  File nay DOC LAP hoan toan voi App.jsx — tu dinh nghia helper rieng,   */
/*  khong can App.jsx export them gi. Hong file nay khong keo theo phan    */
/*  dang chay.                                                             */
/*                                                                         */
/*  Nguon du lieu: 14 ham SQL da tao o phase40 -> phase43a.               */
/* ---------------------------------------------------------------------- */

import React, { useEffect, useState, useCallback } from "react";
import { supabase } from "./supabaseClient";
import * as XLSX from "xlsx";
import {
  Loader2, FileSpreadsheet, Wallet, Users, AlertTriangle,
  Building2, Package, ArrowLeftRight, Scale,
} from "lucide-react";

/* ---------------------------------------------------------------------- */
/* Helpers                                                                 */
/* ---------------------------------------------------------------------- */

function cx(...a) {
  return a.filter(Boolean).join(" ");
}

function todayStr() {
  return new Date().toISOString().slice(0, 10);
}

/* Ky ke toan chay tu ngay 26 thang truoc den 25 thang nay */
function periodStart(dateStr) {
  const d = new Date(dateStr);
  const day = d.getDate();
  const m = day >= 26 ? d.getMonth() : d.getMonth() - 1;
  return new Date(d.getFullYear(), m, 26).toISOString().slice(0, 10);
}

function fmtVND(n) {
  if (n === null || n === undefined || n === "") return "";
  const num = Number(n);
  if (Number.isNaN(num)) return "";
  return num.toLocaleString("vi-VN");
}

function fmtDate(d) {
  if (!d) return "";
  try {
    return new Date(d).toLocaleDateString("vi-VN");
  } catch {
    return String(d);
  }
}

function Card({ className = "", children }) {
  return (
    <div className={cx("bg-white rounded-2xl shadow-card border border-slate-200/60", className)}>
      {children}
    </div>
  );
}

/* Xuat bang dang hien ra file Excel — giu dung thu tu cot dang xem */
function exportSheet(rows, columns, fileName) {
  if (!rows || rows.length === 0) return;
  const data = rows.map((r) => {
    const o = {};
    for (const c of columns) o[c.label] = c.raw ? c.raw(r) : r[c.key];
    return o;
  });
  const ws = XLSX.utils.json_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Sheet1");
  XLSX.writeFile(wb, `${fileName}.xlsx`);
}

/* Bang dung chung cho ca 8 man — mot noi sua, tat ca cung doi */
function DataTable({ columns, rows, loading, empty, rowClass }) {
  if (loading) {
    return (
      <div className="flex justify-center py-10">
        <Loader2 className="animate-spin text-slate-300" />
      </div>
    );
  }
  if (!rows || rows.length === 0) {
    return <p className="text-sm text-slate-400 py-10 text-center">{empty || "Chưa có dữ liệu"}</p>;
  }
  return (
    <div className="overflow-x-auto">
      <table className="w-full text-sm">
        <thead>
          <tr className="border-b border-slate-200 bg-slate-50/60">
            {columns.map((c) => (
              <th key={c.key}
                className={cx("px-3 py-2.5 font-medium text-slate-600 whitespace-nowrap",
                  c.align === "right" ? "text-right" : "text-left")}>
                {c.label}
              </th>
            ))}
          </tr>
        </thead>
        <tbody>
          {rows.map((r, i) => (
            <tr key={i}
              className={cx("border-b border-slate-100 hover:bg-slate-50/60 transition",
                rowClass ? rowClass(r) : "")}>
              {columns.map((c) => (
                <td key={c.key}
                  className={cx("px-3 py-2 whitespace-nowrap",
                    c.align === "right" ? "text-right tabular-nums" : "text-left",
                    c.strong ? "font-medium text-slate-800" : "text-slate-600")}>
                  {c.render ? c.render(r) : (r[c.key] ?? "")}
                </td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  );
}

/* Thanh cong cu tren moi man: chon ngay/ky + nut xuat Excel */
function Toolbar({ children, onExport, exportDisabled }) {
  return (
    <div className="flex items-center justify-between gap-2 flex-wrap mb-3">
      <div className="flex items-center gap-2 flex-wrap">{children}</div>
      <button onClick={onExport} disabled={exportDisabled}
        className={cx("flex items-center gap-1.5 text-sm rounded-xl px-3 py-1.5 border transition",
          exportDisabled
            ? "border-slate-200 text-slate-300 cursor-not-allowed"
            : "border-slate-200 text-slate-600 hover:bg-slate-50")}>
        <FileSpreadsheet size={14} /> Xuất Excel
      </button>
    </div>
  );
}

function DateBox({ label, value, onChange }) {
  return (
    <div className="flex items-center gap-1.5 bg-white border border-slate-200 rounded-xl px-3 py-1.5">
      <span className="text-xs text-slate-400">{label}</span>
      <input type="date" value={value} onChange={(e) => onChange(e.target.value)}
        className="text-sm outline-none" />
    </div>
  );
}

function ErrorNote({ msg }) {
  if (!msg) return null;
  return (
    <div className="mb-3 rounded-xl bg-rose-50 border border-rose-200 px-3 py-2 text-sm text-rose-700">
      {msg}
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* 1. So quy ngay — hai cot ton chay dom song song                        */
/* ---------------------------------------------------------------------- */

function SheetCashBook({ employee }) {
  const [date, setDate] = useState(todayStr());
  const [rows, setRows] = useState([]);
  const [opening, setOpening] = useState(null);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  const load = useCallback(async () => {
    setLoading(true); setError("");
    const [{ data, error: e1 }, { data: op }] = await Promise.all([
      supabase.rpc("report_cash_book_balance", {
        p_store_id: employee.store_id, p_date: date,
      }),
      supabase.rpc("report_cash_book_opening", {
        p_store_id: employee.store_id, p_date: date,
      }),
    ]);
    if (e1) setError(e1.message);
    setRows(data || []);
    setOpening(Array.isArray(op) ? op[0] : op);
    setLoading(false);
  }, [employee.store_id, date]);

  useEffect(() => { load(); }, [load]);

  const thu = rows.filter((r) => r.section === "thu");
  const chi = rows.filter((r) => r.section === "chi");

  const columns = [
    { key: "stt", label: "STT" },
    { key: "content", label: "Nội dung", strong: true },
    { key: "doc_code", label: "Chứng từ" },
    { key: "thu", label: "Thu", align: "right", render: (r) => fmtVND(r.thu), raw: (r) => Number(r.thu || 0) },
    { key: "chi", label: "Chi", align: "right", render: (r) => fmtVND(r.chi), raw: (r) => Number(r.chi || 0) },
    { key: "no_amount", label: "Nợ", align: "right", render: (r) => fmtVND(r.no_amount), raw: (r) => Number(r.no_amount || 0) },
    { key: "ton_tien_mat", label: "Tồn tiền mặt", align: "right", render: (r) => fmtVND(r.ton_tien_mat), raw: (r) => Number(r.ton_tien_mat || 0) },
    { key: "ton_ngan_hang", label: "Tồn ngân hàng", align: "right", render: (r) => fmtVND(r.ton_ngan_hang), raw: (r) => Number(r.ton_ngan_hang || 0) },
  ];

  const sum = (arr, f) => arr.reduce((s, r) => s + Number(r[f] || 0), 0);
  const last = rows[rows.length - 1];

  return (
    <div>
      <Toolbar onExport={() => exportSheet(rows, columns, `So quy ${date}`)}
        exportDisabled={rows.length === 0}>
        <DateBox label="Ngày" value={date} onChange={setDate} />
      </Toolbar>
      <ErrorNote msg={error} />

      {opening && (
        <div className="flex flex-wrap gap-3 text-xs mb-3">
          <div className="bg-emerald-50 rounded-lg px-3 py-2">
            <span className="text-emerald-600">Tồn tiền mặt đầu ngày: </span>
            <span className="font-semibold text-emerald-800">{fmtVND(opening.cash_bd)}</span>
          </div>
          <div className="bg-sky-50 rounded-lg px-3 py-2">
            <span className="text-sky-600">Tồn ngân hàng đầu ngày: </span>
            <span className="font-semibold text-sky-800">{fmtVND(opening.bank_bd)}</span>
          </div>
          <div className="bg-slate-50 rounded-lg px-3 py-2">
            <span className="text-slate-400">Đầu kỳ từ: </span>
            <span className="font-medium text-slate-600">{fmtDate(opening.period_start)}</span>
          </div>
        </div>
      )}

      <Card className="p-4 mb-4">
        <p className="text-sm font-medium text-emerald-700 mb-2">THU</p>
        <DataTable columns={columns} rows={thu} loading={loading} empty="Không có khoản thu" />
        {thu.length > 0 && (
          <p className="text-xs text-slate-500 mt-2 text-right">
            Tổng thu: <span className="font-semibold text-slate-700">{fmtVND(sum(thu, "thu"))}</span>
            {" · "}Nợ: <span className="font-semibold text-slate-700">{fmtVND(sum(thu, "no_amount"))}</span>
          </p>
        )}
      </Card>

      <Card className="p-4 mb-4">
        <p className="text-sm font-medium text-rose-700 mb-2">CHI</p>
        <DataTable columns={columns} rows={chi} loading={loading} empty="Không có khoản chi" />
        {chi.length > 0 && (
          <p className="text-xs text-slate-500 mt-2 text-right">
            Tổng chi: <span className="font-semibold text-slate-700">{fmtVND(sum(chi, "chi"))}</span>
            {" · "}Nợ: <span className="font-semibold text-slate-700">{fmtVND(sum(chi, "no_amount"))}</span>
          </p>
        )}
      </Card>

      {last && (
        <div className="flex flex-wrap gap-3 text-sm">
          <div className="bg-emerald-50 rounded-xl px-4 py-2.5">
            <span className="text-emerald-600 text-xs block">Tồn tiền mặt cuối ngày</span>
            <span className="font-semibold text-emerald-800">{fmtVND(last.ton_tien_mat)}</span>
          </div>
          <div className="bg-sky-50 rounded-xl px-4 py-2.5">
            <span className="text-sky-600 text-xs block">Tồn ngân hàng cuối ngày</span>
            <span className="font-semibold text-sky-800">{fmtVND(last.ton_ngan_hang)}</span>
          </div>
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* 2. Khach no                                                             */
/* ---------------------------------------------------------------------- */

const SECTION_LABELS = {
  bad_debt: "Nợ khó đòi (tổng)",
  khach: "Khách lẻ",
  ncc: "Nhà cung cấp / đối tác hai chiều",
};

function SheetCustomerDebt({ employee }) {
  const [date, setDate] = useState(todayStr());
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  const load = useCallback(async () => {
    setLoading(true); setError("");
    const { data, error: err } = await supabase.rpc("report_customer_debt", {
      p_store_id: employee.store_id, p_date: date,
    });
    if (err) setError(err.message);
    setRows(data || []);
    setLoading(false);
  }, [employee.store_id, date]);

  useEffect(() => { load(); }, [load]);

  const columns = [
    { key: "stt", label: "STT" },
    { key: "ngay", label: "Ngày", render: (r) => fmtDate(r.ngay) },
    { key: "khach", label: "Khách", strong: true },
    { key: "noi_dung", label: "Nội dung" },
    { key: "so_hd", label: "Công nợ phát sinh" },
    { key: "no_ps", label: "Nợ", align: "right", render: (r) => fmtVND(r.no_ps), raw: (r) => Number(r.no_ps || 0) },
    { key: "da_tra", label: "Trả", align: "right", render: (r) => fmtVND(r.da_tra), raw: (r) => Number(r.da_tra || 0) },
    { key: "ton_no", label: "Tồn nợ", align: "right", strong: true, render: (r) => fmtVND(r.ton_no), raw: (r) => Number(r.ton_no || 0) },
    { key: "ngay_tra", label: "Ngày trả", render: (r) => fmtDate(r.ngay_tra) },
  ];

  const groups = ["bad_debt", "khach", "ncc"]
    .map((s) => ({ key: s, rows: rows.filter((r) => r.section === s) }))
    .filter((g) => g.rows.length > 0);

  const tongTonNo = rows.reduce((s, r) => s + Number(r.ton_no || 0), 0);

  return (
    <div>
      <Toolbar onExport={() => exportSheet(rows, columns, `Khach no ${date}`)}
        exportDisabled={rows.length === 0}>
        <DateBox label="Tính đến" value={date} onChange={setDate} />
      </Toolbar>
      <ErrorNote msg={error} />

      {loading ? (
        <div className="flex justify-center py-10"><Loader2 className="animate-spin text-slate-300" /></div>
      ) : groups.length === 0 ? (
        <p className="text-sm text-slate-400 py-10 text-center">Chưa có công nợ khách hàng</p>
      ) : (
        <>
          {groups.map((g) => (
            <Card key={g.key} className="p-4 mb-4">
              <p className="text-sm font-medium text-slate-700 mb-2">{SECTION_LABELS[g.key]}</p>
              {g.key === "ncc" && (
                <p className="text-[11px] text-slate-500 mb-2">
                  Đối tác vừa mua vừa bán nên tồn nợ có thể âm — âm nghĩa là mình đang nợ họ.
                </p>
              )}
              <DataTable columns={columns} rows={g.rows} loading={false}
                rowClass={(r) => (Number(r.ton_no) < 0 ? "bg-amber-50/40" : "")} />
            </Card>
          ))}
          <p className="text-sm text-right">
            <span className="text-slate-500">Tổng khách nợ: </span>
            <span className="font-semibold text-slate-800">{fmtVND(tongTonNo)}</span>
          </p>
        </>
      )}
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* 3. No kho doi                                                           */
/* ---------------------------------------------------------------------- */

function SheetBadDebt({ employee }) {
  const [date, setDate] = useState(todayStr());
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  const load = useCallback(async () => {
    setLoading(true); setError("");
    const { data, error: err } = await supabase.rpc("report_bad_debt", {
      p_store_id: employee.store_id, p_date: date,
    });
    if (err) setError(err.message);
    setRows(data || []);
    setLoading(false);
  }, [employee.store_id, date]);

  useEffect(() => { load(); }, [load]);

  const columns = [
    { key: "stt", label: "STT" },
    { key: "ngay", label: "Ngày", render: (r) => fmtDate(r.ngay) },
    { key: "khach", label: "Khách", strong: true },
    { key: "sdt", label: "SĐT" },
    { key: "noi_dung", label: "Nội dung" },
    { key: "so_hd", label: "Công nợ phát sinh" },
    { key: "no_ps", label: "Nợ", align: "right", render: (r) => fmtVND(r.no_ps), raw: (r) => Number(r.no_ps || 0) },
    { key: "da_tra", label: "Trả", align: "right", render: (r) => fmtVND(r.da_tra), raw: (r) => Number(r.da_tra || 0) },
    { key: "ton_no", label: "Tồn nợ", align: "right", strong: true, render: (r) => fmtVND(r.ton_no), raw: (r) => Number(r.ton_no || 0) },
    { key: "ngay_tra", label: "Ngày trả", render: (r) => fmtDate(r.ngay_tra) },
    { key: "lan", label: "Lần", render: (r) => (r.lan ? `Lần ${r.lan}` : "") },
  ];

  const tong = rows.reduce((s, r) => s + Number(r.ton_no || 0), 0);

  return (
    <div>
      <Toolbar onExport={() => exportSheet(rows, columns, `No kho doi ${date}`)}
        exportDisabled={rows.length === 0}>
        <DateBox label="Tính đến" value={date} onChange={setDate} />
      </Toolbar>
      <ErrorNote msg={error} />

      <Card className="p-4">
        <DataTable columns={columns} rows={rows} loading={loading}
          empty="Chưa đánh dấu khoản nợ khó đòi nào" />
        {rows.length > 0 && (
          <p className="text-sm mt-3 text-right">
            <span className="text-slate-500">Tổng nợ khó đòi: </span>
            <span className="font-semibold text-rose-700">{fmtVND(tong)}</span>
          </p>
        )}
      </Card>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* 4. No doi tac                                                           */
/* ---------------------------------------------------------------------- */

function SheetPartnerDebt({ employee }) {
  const [date, setDate] = useState(todayStr());
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  const load = useCallback(async () => {
    setLoading(true); setError("");
    const { data, error: err } = await supabase.rpc("report_partner_debt", {
      p_store_id: employee.store_id, p_date: date,
    });
    if (err) setError(err.message);
    setRows(data || []);
    setLoading(false);
  }, [employee.store_id, date]);

  useEffect(() => { load(); }, [load]);

  const columns = [
    { key: "stt", label: "STT" },
    { key: "ngay", label: "Ngày", render: (r) => fmtDate(r.ngay) },
    { key: "doi_tac", label: "Đối tác", strong: true },
    { key: "noi_dung", label: "Nội dung" },
    { key: "so_tien", label: "Số tiền", align: "right", render: (r) => fmtVND(r.so_tien), raw: (r) => Number(r.so_tien || 0) },
    { key: "thanh_toan", label: "Thanh toán", align: "right", render: (r) => fmtVND(r.thanh_toan), raw: (r) => Number(r.thanh_toan || 0) },
    { key: "no_con_lai", label: "Nợ còn lại", align: "right", strong: true, render: (r) => fmtVND(r.no_con_lai), raw: (r) => Number(r.no_con_lai || 0) },
    { key: "ngay_tra", label: "Ngày trả", render: (r) => fmtDate(r.ngay_tra) },
  ];

  const tong = rows.reduce((s, r) => s + Number(r.no_con_lai || 0), 0);

  return (
    <div>
      <Toolbar onExport={() => exportSheet(rows, columns, `No doi tac ${date}`)}
        exportDisabled={rows.length === 0}>
        <DateBox label="Tính đến" value={date} onChange={setDate} />
      </Toolbar>
      <ErrorNote msg={error} />
      <p className="text-[11px] text-slate-500 mb-2">
        Chỉ gồm nhà cung cấp ký gửi máy. Lương nhân viên ghi ở mục Chi phí khác.
      </p>

      <Card className="p-4">
        <DataTable columns={columns} rows={rows} loading={loading}
          empty="Chưa có công nợ đối tác" />
        {rows.length > 0 && (
          <p className="text-sm mt-3 text-right">
            <span className="text-slate-500">Tổng nợ đối tác: </span>
            <span className="font-semibold text-slate-800">{fmtVND(tong)}</span>
          </p>
        )}
      </Card>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* 5. May ton — ca may con ton lan may da ban (lich su may)                */
/* ---------------------------------------------------------------------- */

function SheetInventory({ employee }) {
  const [date, setDate] = useState(todayStr());
  const [from, setFrom] = useState(periodStart(todayStr()));
  const [rows, setRows] = useState([]);
  const [summary, setSummary] = useState([]);
  const [onlyStock, setOnlyStock] = useState(false);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  const load = useCallback(async () => {
    setLoading(true); setError("");
    const [{ data, error: e1 }, { data: s }] = await Promise.all([
      supabase.rpc("report_device_inventory", {
        p_store_id: employee.store_id, p_date: date,
      }),
      supabase.rpc("report_device_inventory_summary", {
        p_store_id: employee.store_id, p_from: from, p_to: date,
      }),
    ]);
    if (e1) setError(e1.message);
    setRows(data || []);
    setSummary(s || []);
    setLoading(false);
  }, [employee.store_id, date, from]);

  useEffect(() => { load(); }, [load]);

  const shown = onlyStock ? rows.filter((r) => r.trang_thai === "Ton") : rows;

  const columns = [
    { key: "stt", label: "STT" },
    { key: "ngay", label: "Ngày nhập", render: (r) => fmtDate(r.ngay) },
    { key: "ten_may", label: "Tên máy", strong: true },
    { key: "doi_may", label: "Đời" },
    { key: "ma", label: "Mã" },
    { key: "imei", label: "IMEI" },
    { key: "so_luong", label: "SL", align: "right" },
    { key: "gia_nhap", label: "Giá nhập", align: "right", render: (r) => fmtVND(r.gia_nhap), raw: (r) => Number(r.gia_nhap || 0) },
    { key: "doi_tac", label: "Đối tác" },
    { key: "gia_von", label: "Giá vốn", align: "right", render: (r) => fmtVND(r.gia_von), raw: (r) => Number(r.gia_von || 0) },
    { key: "ngay_ban", label: "Ngày bán", render: (r) => fmtDate(r.ngay_ban) },
    {
      key: "trang_thai", label: "Trạng thái",
      render: (r) => (
        <span className={cx("text-xs rounded-full px-2 py-0.5",
          r.trang_thai === "Ton" ? "bg-emerald-50 text-emerald-700" : "bg-slate-100 text-slate-500")}>
          {r.trang_thai === "Ton" ? "Tồn" : "Đã bán"}
        </span>
      ),
      raw: (r) => (r.trang_thai === "Ton" ? "Tồn" : "Đã bán"),
    },
    { key: "ghi_chu", label: "Ghi chú" },
  ];

  return (
    <div>
      <Toolbar onExport={() => exportSheet(shown, columns, `May ton ${date}`)}
        exportDisabled={shown.length === 0}>
        <DateBox label="Từ" value={from} onChange={setFrom} />
        <DateBox label="Đến" value={date} onChange={setDate} />
        <label className="flex items-center gap-1.5 text-sm text-slate-600 bg-white border border-slate-200 rounded-xl px-3 py-1.5 cursor-pointer">
          <input type="checkbox" checked={onlyStock} onChange={(e) => setOnlyStock(e.target.checked)} />
          Chỉ máy còn tồn
        </label>
      </Toolbar>
      <ErrorNote msg={error} />

      {summary.length > 0 && (
        <div className="grid grid-cols-2 lg:grid-cols-5 gap-3 mb-4">
          {summary.map((s) => (
            <div key={s.chi_tieu} className="bg-slate-50 rounded-xl p-3">
              <p className="text-xs text-slate-500 mb-1">
                {{ DK: "Đầu kỳ", Nhap: "Nhập", Xuat: "Xuất", "Ky gui": "Ký gửi", Ton: "Tồn" }[s.chi_tieu] || s.chi_tieu}
              </p>
              <p className="text-base font-semibold text-slate-800">{s.so_luong} máy</p>
              <p className="text-[11px] text-slate-500">{fmtVND(s.gia_tri)}</p>
            </div>
          ))}
        </div>
      )}

      <Card className="p-4">
        <DataTable columns={columns} rows={shown} loading={loading} empty="Chưa có máy nào" />
      </Card>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* 6. Nhap may — nhat ky nhap theo thoi gian                               */
/* ---------------------------------------------------------------------- */

function SheetIntake({ employee }) {
  const [from, setFrom] = useState(periodStart(todayStr()));
  const [to, setTo] = useState(todayStr());
  const [rows, setRows] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  const load = useCallback(async () => {
    setLoading(true); setError("");
    const { data, error: err } = await supabase.rpc("report_device_intake", {
      p_store_id: employee.store_id, p_from: from, p_to: to,
    });
    if (err) setError(err.message);
    setRows(data || []);
    setLoading(false);
  }, [employee.store_id, from, to]);

  useEffect(() => { load(); }, [load]);

  const columns = [
    { key: "stt", label: "STT" },
    { key: "ngay", label: "Ngày", render: (r) => fmtDate(r.ngay) },
    { key: "ma_phieu", label: "Mã phiếu" },
    { key: "ten_may", label: "Tên máy", strong: true },
    { key: "imei", label: "IMEI" },
    { key: "so_luong", label: "SL", align: "right" },
    { key: "gia_nhap", label: "Giá nhập", align: "right", render: (r) => fmtVND(r.gia_nhap), raw: (r) => Number(r.gia_nhap || 0) },
    { key: "doi_tac", label: "Đối tác" },
    { key: "gia_von", label: "Giá vốn", align: "right", render: (r) => fmtVND(r.gia_von), raw: (r) => Number(r.gia_von || 0) },
    { key: "ngay_ban", label: "Ngày bán", render: (r) => fmtDate(r.ngay_ban) },
    { key: "trang_thai", label: "Trạng thái", render: (r) => (r.trang_thai === "Ton" ? "Tồn" : "Đã bán") },
    { key: "ghi_chu", label: "Ghi chú" },
  ];

  const tong = rows.reduce((s, r) => s + Number(r.gia_nhap || 0), 0);

  return (
    <div>
      <Toolbar onExport={() => exportSheet(rows, columns, `Nhap may ${from} - ${to}`)}
        exportDisabled={rows.length === 0}>
        <DateBox label="Từ" value={from} onChange={setFrom} />
        <DateBox label="Đến" value={to} onChange={setTo} />
      </Toolbar>
      <ErrorNote msg={error} />

      <Card className="p-4">
        <DataTable columns={columns} rows={rows} loading={loading} empty="Không có phiếu nhập trong kỳ" />
        {rows.length > 0 && (
          <p className="text-sm mt-3 text-right">
            <span className="text-slate-500">{rows.length} phiếu · Tổng giá nhập: </span>
            <span className="font-semibold text-slate-800">{fmtVND(tong)}</span>
          </p>
        )}
      </Card>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* 7. Cong no lien cua hang — co o chon cua hang doi ung                   */
/* ---------------------------------------------------------------------- */

function SheetInternalDebt({ employee }) {
  const [stores, setStores] = useState([]);
  const [counterpart, setCounterpart] = useState("");
  const [from, setFrom] = useState(periodStart(todayStr()));
  const [to, setTo] = useState(todayStr());
  const [rows, setRows] = useState([]);
  const [summary, setSummary] = useState([]);
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState("");

  useEffect(() => {
    (async () => {
      const { data } = await supabase.from("stores").select("id, name").order("name");
      const others = (data || []).filter((s) => s.id !== employee.store_id);
      setStores(others);
      if (others.length > 0) setCounterpart(others[0].id);
    })();
  }, [employee.store_id]);

  const load = useCallback(async () => {
    if (!counterpart) return;
    setLoading(true); setError("");
    const [{ data, error: e1 }, { data: s }] = await Promise.all([
      supabase.rpc("report_internal_debt", {
        p_store_id: employee.store_id, p_counterpart_id: counterpart,
        p_from: from, p_to: to,
      }),
      supabase.rpc("report_internal_debt_summary", {
        p_store_id: employee.store_id, p_counterpart_id: counterpart,
        p_from: from, p_to: to,
      }),
    ]);
    if (e1) setError(e1.message);
    setRows(data || []);
    setSummary(s || []);
    setLoading(false);
  }, [employee.store_id, counterpart, from, to]);

  useEffect(() => { load(); }, [load]);

  const tenDoiUng = stores.find((s) => s.id === counterpart)?.name || "cửa hàng bạn";

  const columns = [
    { key: "stt", label: "STT" },
    { key: "ngay", label: "Ngày", render: (r) => fmtDate(r.ngay) },
    { key: "loai", label: "Loại" },
    { key: "ma_ct", label: "Chứng từ" },
    { key: "noi_dung", label: "Nội dung", strong: true },
    { key: "minh_no_ho", label: `Mình nợ ${tenDoiUng}`, align: "right", render: (r) => fmtVND(r.minh_no_ho), raw: (r) => Number(r.minh_no_ho || 0) },
    { key: "ho_no_minh", label: `${tenDoiUng} nợ mình`, align: "right", render: (r) => fmtVND(r.ho_no_minh), raw: (r) => Number(r.ho_no_minh || 0) },
    { key: "du_no", label: "Dư nợ lũy kế", align: "right", strong: true, render: (r) => fmtVND(r.du_no), raw: (r) => Number(r.du_no || 0) },
    { key: "ghi_chu", label: "Ghi chú" },
  ];

  const canTru = summary.find((s) => s.loai === "CAN TRU CON LAI");

  return (
    <div>
      <Toolbar onExport={() => exportSheet(rows, columns, `Cong no lien cua hang ${from} - ${to}`)}
        exportDisabled={rows.length === 0}>
        <div className="flex items-center gap-1.5 bg-white border border-slate-200 rounded-xl px-3 py-1.5">
          <span className="text-xs text-slate-400">Đối ứng</span>
          <select value={counterpart} onChange={(e) => setCounterpart(e.target.value)}
            className="text-sm outline-none bg-transparent">
            {stores.map((s) => <option key={s.id} value={s.id}>{s.name}</option>)}
          </select>
        </div>
        <DateBox label="Từ" value={from} onChange={setFrom} />
        <DateBox label="Đến" value={to} onChange={setTo} />
      </Toolbar>
      <ErrorNote msg={error} />

      {canTru && (
        <div className={cx("rounded-xl px-4 py-3 mb-4 text-sm",
          Number(canTru.rong) > 0 ? "bg-amber-50 text-amber-800"
            : Number(canTru.rong) < 0 ? "bg-emerald-50 text-emerald-800"
              : "bg-slate-50 text-slate-600")}>
          <span className="font-semibold">Cấn trừ còn lại: </span>
          {canTru.ghi_chu}
        </div>
      )}

      {summary.filter((s) => s.loai !== "CAN TRU CON LAI").length > 0 && (
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 mb-4">
          {summary.filter((s) => s.loai !== "CAN TRU CON LAI").map((s) => (
            <div key={s.loai} className="bg-slate-50 rounded-xl p-3">
              <p className="text-xs text-slate-500 mb-1">{s.loai} · {s.so_ct} chứng từ</p>
              <p className="text-sm text-slate-700">
                Mình nợ: <span className="font-medium">{fmtVND(s.minh_no_ho)}</span>
              </p>
              <p className="text-sm text-slate-700">
                Họ nợ: <span className="font-medium">{fmtVND(s.ho_no_minh)}</span>
              </p>
            </div>
          ))}
        </div>
      )}

      <Card className="p-4">
        <DataTable columns={columns} rows={rows} loading={loading}
          empty="Không có phát sinh giữa hai cửa hàng trong kỳ" />
      </Card>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* 8. Doi soat cong no 3 don vi — CHI tai khoan chu vao duoc               */
/*    Man nay dung ngoai bo 7 sheet, cong no doi tac ngoai cua ca 3 CH.    */
/* ---------------------------------------------------------------------- */

function SheetOwnerReconcile() {
  const [date, setDate] = useState(todayStr());
  const [stores, setStores] = useState([]);
  const [matrix, setMatrix] = useState([]);
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  const load = useCallback(async () => {
    setLoading(true); setError("");
    const { data: st, error: e0 } = await supabase
      .from("stores").select("id, name").order("name");
    if (e0) { setError(e0.message); setLoading(false); return; }
    const list = st || [];
    setStores(list);

    // Goi report_partner_debt cho tung cua hang roi gop theo doi tac
    const results = await Promise.all(
      list.map((s) =>
        supabase.rpc("report_partner_debt", { p_store_id: s.id, p_date: date })
      )
    );

    const byPartner = {};
    results.forEach((res, i) => {
      const storeId = list[i].id;
      for (const r of res.data || []) {
        const name = r.doi_tac || "(không rõ)";
        if (!byPartner[name]) byPartner[name] = { doi_tac: name };
        byPartner[name][storeId] =
          Number(byPartner[name][storeId] || 0) + Number(r.no_con_lai || 0);
      }
    });

    const rows = Object.values(byPartner).filter((r) =>
      list.some((s) => Number(r[s.id] || 0) !== 0)
    );
    setMatrix(rows);
    setLoading(false);
  }, [date]);

  useEffect(() => { load(); }, [load]);

  const columns = [
    { key: "doi_tac", label: "Đối tác", strong: true },
    ...stores.map((s) => ({
      key: s.id, label: s.name, align: "right",
      render: (r) => fmtVND(r[s.id]),
      raw: (r) => Number(r[s.id] || 0),
    })),
    {
      key: "_tong", label: "Tổng", align: "right", strong: true,
      render: (r) => fmtVND(stores.reduce((t, s) => t + Number(r[s.id] || 0), 0)),
      raw: (r) => stores.reduce((t, s) => t + Number(r[s.id] || 0), 0),
    },
  ];

  const tongTheoCH = {};
  for (const s of stores) {
    tongTheoCH[s.id] = matrix.reduce((t, r) => t + Number(r[s.id] || 0), 0);
  }

  return (
    <div>
      <div className="mb-4">
        <h2 className="text-lg font-semibold text-slate-800">Đối soát công nợ 3 đơn vị</h2>
        <p className="text-xs text-slate-500">
          Nợ đối tác bên ngoài, chia theo cửa hàng. Chỉ tài khoản chủ xem được màn này.
        </p>
      </div>

      <Toolbar onExport={() => exportSheet(matrix, columns, `Doi soat cong no ${date}`)}
        exportDisabled={matrix.length === 0}>
        <DateBox label="Tính đến" value={date} onChange={setDate} />
      </Toolbar>
      <ErrorNote msg={error} />

      <Card className="p-4">
        <DataTable columns={columns} rows={matrix} loading={loading}
          empty="Chưa có công nợ đối tác ở cửa hàng nào" />
        {matrix.length > 0 && (
          <div className="flex flex-wrap gap-3 mt-4 text-sm">
            {stores.map((s) => (
              <div key={s.id} className="bg-slate-50 rounded-xl px-3 py-2">
                <span className="text-xs text-slate-500 block">{s.name}</span>
                <span className="font-semibold text-slate-800">{fmtVND(tongTheoCH[s.id])}</span>
              </div>
            ))}
          </div>
        )}
      </Card>
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* Vo ngoai: 7 sheet chung mot man co tab                                  */
/* ---------------------------------------------------------------------- */

const SHEETS = [
  { key: "cashbook", label: "Sổ quỹ", icon: Wallet, Comp: SheetCashBook },
  { key: "customer", label: "Khách nợ", icon: Users, Comp: SheetCustomerDebt },
  { key: "bad", label: "Nợ khó đòi", icon: AlertTriangle, Comp: SheetBadDebt },
  { key: "partner", label: "Nợ đối tác", icon: Building2, Comp: SheetPartnerDebt },
  { key: "stock", label: "Máy tồn", icon: Package, Comp: SheetInventory },
  { key: "intake", label: "Nhập máy", icon: ArrowLeftRight, Comp: SheetIntake },
  { key: "internal", label: "Công nợ liên CH", icon: Scale, Comp: SheetInternalDebt },
];

export function ReportsExcelModule({ employee }) {
  const [tab, setTab] = useState("cashbook");
  const active = SHEETS.find((s) => s.key === tab) || SHEETS[0];
  const Comp = active.Comp;

  return (
    <div>
      <div className="mb-4">
        <h2 className="text-lg font-semibold text-slate-800">Sổ sách</h2>
        <p className="text-xs text-slate-500">Bộ báo cáo theo đúng biểu mẫu Excel cũ</p>
      </div>

      <div className="flex gap-1.5 flex-wrap mb-4 border-b border-slate-200 pb-2">
        {SHEETS.map((s) => {
          const Icon = s.icon;
          return (
            <button key={s.key} onClick={() => setTab(s.key)}
              className={cx("flex items-center gap-1.5 text-sm rounded-xl px-3 py-1.5 transition",
                tab === s.key
                  ? "bg-brand-50 text-brand-700 font-medium"
                  : "text-slate-500 hover:bg-slate-50")}>
              <Icon size={14} /> {s.label}
            </button>
          );
        })}
      </div>

      <Comp employee={employee} />
    </div>
  );
}

export { SheetOwnerReconcile as OwnerReconcileModule };
export default ReportsExcelModule;
