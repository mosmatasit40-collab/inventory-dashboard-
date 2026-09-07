import { useState, useEffect, useMemo, useRef, useCallback } from "react";
import * as XLSX from "xlsx";
import Papa from "papaparse";
import { createClient } from "@supabase/supabase-js";
import {
  ResponsiveContainer,
  ComposedChart,
  Bar,
  Line,
  XAxis,
  YAxis,
  CartesianGrid,
  Tooltip,
  Legend,
} from "recharts";
import {
  Upload,
  RefreshCw,
  PackageSearch,
  TrendingDown,
  TrendingUp,
  AlertTriangle,
  Link as LinkIcon,
  Trash2,
  ChevronRight,
  BoxSelect,
  Download,
} from "lucide-react";

/* ---------------------------------------------------------------------- */
/* Design tokens                                                          */
/* ---------------------------------------------------------------------- */
const COLORS = {
  bg: "#12161A",
  surface: "#1B2127",
  surface2: "#212A32",
  border: "#2B3540",
  text: "#E7ECEF",
  muted: "#8C9AA6",
  faint: "#5C6873",
  accent: "#4FA3D1",
  accentDim: "#2E5A73",
  warn: "#E8A33D",
  danger: "#E0645A",
  success: "#5FB77E",
};

const FONT_LINK_ID = "inv-dash-fonts";

function useGoogleFonts() {
  useEffect(() => {
    if (document.getElementById(FONT_LINK_ID)) return;
    const link = document.createElement("link");
    link.id = FONT_LINK_ID;
    link.rel = "stylesheet";
    link.href =
      "https://fonts.googleapis.com/css2?family=IBM+Plex+Sans+Thai:wght@400;500;600;700&family=IBM+Plex+Mono:wght@400;500;600&display=swap";
    document.head.appendChild(link);
  }, []);
}

/* ---------------------------------------------------------------------- */
/* Storage helpers — shared cloud storage via Supabase, so every visitor  */
/* sees the same live data (not just their own browser).                 */
/* ---------------------------------------------------------------------- */
const K_PRODUCTS = "inv:products";
const K_TX = "inv:transactions";

const supabaseUrl = import.meta.env.VITE_SUPABASE_URL;
const supabaseKey = import.meta.env.VITE_SUPABASE_ANON_KEY;
const supabase = supabaseUrl && supabaseKey ? createClient(supabaseUrl, supabaseKey) : null;

async function loadJSON(key, fallback) {
  if (!supabase) return fallback; // env vars not configured yet — falls back gracefully
  try {
    const { data, error } = await supabase.from("kv_store").select("value").eq("id", key).maybeSingle();
    if (error || !data) return fallback;
    return data.value;
  } catch {
    return fallback;
  }
}
async function saveJSON(key, value) {
  if (!supabase) return;
  try {
    await supabase.from("kv_store").upsert({ id: key, value, updated_at: new Date().toISOString() });
  } catch (e) {
    console.error("storage set failed", key, e);
  }
}

/* ---------------------------------------------------------------------- */
/* Seed demo data (used only the very first time, so the dashboard isn't  */
/* empty before the user imports their own file)                          */
/* ---------------------------------------------------------------------- */
function seedData() {
  const today = new Date();
  const fmt = (d) => d.toISOString().slice(0, 10);
  const products = {
    "SKU-001": { sku: "SKU-001", name: "น้ำดื่มขวด 600ml", unit: "แพ็ค", stock: 84, leadTimeDays: 3, safetyStock: 20 },
    "SKU-002": { sku: "SKU-002", name: "กระดาษ A4 80แกรม", unit: "รีม", stock: 30, leadTimeDays: 5, safetyStock: 10 },
    "SKU-003": { sku: "SKU-003", name: "ถุงมือยาง size M", unit: "กล่อง", stock: 12, leadTimeDays: 7, safetyStock: 15 },
    "SKU-004": { sku: "SKU-004", name: "น้ำยาทำความสะอาด 5L", unit: "แกลลอน", stock: 46, leadTimeDays: 4, safetyStock: 8 },
  };
  const tx = [];
  let id = 1;
  for (let i = 13; i >= 0; i--) {
    const d = new Date(today);
    d.setDate(d.getDate() - i);
    const date = fmt(d);
    Object.values(products).forEach((p) => {
      const out = Math.max(0, Math.round(3 + Math.random() * 6 + (p.sku === "SKU-003" ? 2 : 0)));
      const inQty = Math.random() > 0.8 ? Math.round(20 + Math.random() * 30) : 0;
      tx.push({ id: id++, date, sku: p.sku, name: p.name, unit: p.unit, in: inQty, out, note: inQty ? "รับเข้าสต๊อก" : "" });
    });
  }
  return { products, tx };
}

/* ---------------------------------------------------------------------- */
/* Forecast math                                                          */
/* ---------------------------------------------------------------------- */
const TREND_WINDOW = 14; // days used to estimate average daily usage
const REVIEW_DAYS = 7; // additional buffer period covered by the suggested order

function metricsFor(product, transactions, horizon = 30) {
  const today = new Date();
  const since = new Date(today);
  since.setDate(since.getDate() - TREND_WINDOW);

  const rows = transactions.filter((t) => t.sku === product.sku && new Date(t.date) >= since);
  const totalOut = rows.reduce((s, r) => s + (Number(r.out) || 0), 0);
  const daysWithData = new Set(rows.map((r) => r.date)).size || 1;
  const avgDailyOut = totalOut / Math.max(daysWithData, TREND_WINDOW / 2, 1) || 0;

  const leadTimeDays = Number(product.leadTimeDays) || 3;
  const safetyStock = Number(product.safetyStock) || 0;
  const reorderPoint = avgDailyOut * leadTimeDays + safetyStock;
  const daysOfStockLeft = avgDailyOut > 0 ? product.stock / avgDailyOut : Infinity;
  const forecastDemand = avgDailyOut * (leadTimeDays + REVIEW_DAYS);
  const suggestedQty = Math.max(0, Math.ceil(forecastDemand + safetyStock - product.stock));
  const needsReorder = product.stock <= reorderPoint;

  // explicit forward-looking projection over the selected horizon (e.g. next 7 / 14 / 30 days)
  const forecastQtyHorizon = Math.round(avgDailyOut * horizon);
  const projectedStockAtHorizon = Math.round(product.stock - forecastQtyHorizon);
  const stockOutInDays = avgDailyOut > 0 ? Math.max(0, Math.floor(product.stock / avgDailyOut)) : null;

  let status = "ปกติ";
  if (product.stock <= safetyStock) status = "วิกฤต";
  else if (needsReorder) status = "ควรสั่งซื้อ";

  return {
    avgDailyOut,
    reorderPoint,
    daysOfStockLeft,
    suggestedQty,
    needsReorder,
    status,
    forecastQtyHorizon,
    projectedStockAtHorizon,
    stockOutInDays,
  };
}

function dailyTrend(transactions, days = TREND_WINDOW) {
  const today = new Date();
  const out = [];
  for (let i = days - 1; i >= 0; i--) {
    const d = new Date(today);
    d.setDate(d.getDate() - i);
    const key = d.toISOString().slice(0, 10);
    const rows = transactions.filter((t) => t.date === key);
    out.push({
      date: key,
      label: d.toLocaleDateString("th-TH", { day: "2-digit", month: "2-digit" }),
      in: rows.reduce((s, r) => s + (Number(r.in) || 0), 0),
      out: rows.reduce((s, r) => s + (Number(r.out) || 0), 0),
    });
  }
  return out;
}

/* ---------------------------------------------------------------------- */
/* Small UI atoms                                                         */
/* ---------------------------------------------------------------------- */
function Badge({ status }) {
  const map = {
    ปกติ: { bg: "rgba(95,183,126,0.14)", fg: COLORS.success, bd: "rgba(95,183,126,0.35)" },
    ควรสั่งซื้อ: { bg: "rgba(232,163,61,0.14)", fg: COLORS.warn, bd: "rgba(232,163,61,0.4)" },
    วิกฤต: { bg: "rgba(224,100,90,0.16)", fg: COLORS.danger, bd: "rgba(224,100,90,0.45)" },
  };
  const c = map[status] || map["ปกติ"];
  return (
    <span
      style={{
        background: c.bg,
        color: c.fg,
        border: `1px solid ${c.bd}`,
        padding: "3px 10px",
        borderRadius: 999,
        fontSize: 12.5,
        fontWeight: 600,
        whiteSpace: "nowrap",
      }}
    >
      {status}
    </span>
  );
}

function KPI({ icon: Icon, label, value, sub, tone }) {
  const toneColor = tone === "warn" ? COLORS.warn : tone === "danger" ? COLORS.danger : COLORS.accent;
  return (
    <div
      style={{
        background: COLORS.surface,
        border: `1px solid ${COLORS.border}`,
        borderRadius: 14,
        padding: "18px 20px",
        display: "flex",
        flexDirection: "column",
        gap: 10,
        minWidth: 0,
      }}
    >
      <div style={{ display: "flex", alignItems: "center", gap: 8, color: COLORS.muted, fontSize: 13 }}>
        <Icon size={16} color={toneColor} />
        <span>{label}</span>
      </div>
      <div style={{ fontFamily: "'IBM Plex Mono', monospace", fontSize: 28, fontWeight: 600, color: COLORS.text, lineHeight: 1 }}>
        {value}
      </div>
      {sub && <div style={{ fontSize: 12.5, color: COLORS.faint }}>{sub}</div>}
    </div>
  );
}

function FilterChip({ active, onClick, color, children }) {
  return (
    <button
      onClick={onClick}
      style={{
        background: active ? `${color}22` : "transparent",
        color: active ? color : COLORS.muted,
        border: `1px solid ${active ? color + "55" : COLORS.border}`,
        borderRadius: 999,
        padding: "7px 13px",
        fontSize: 12.5,
        fontWeight: 500,
        cursor: "pointer",
        fontFamily: "inherit",
        display: "flex",
        alignItems: "center",
        gap: 5,
        whiteSpace: "nowrap",
      }}
    >
      {children}
    </button>
  );
}

function Tab({ active, onClick, children }) {
  return (
    <button
      onClick={onClick}
      style={{
        background: active ? COLORS.surface2 : "transparent",
        color: active ? COLORS.text : COLORS.muted,
        border: `1px solid ${active ? COLORS.border : "transparent"}`,
        borderRadius: 10,
        padding: "8px 16px",
        fontSize: 14,
        fontWeight: 500,
        cursor: "pointer",
        fontFamily: "inherit",
      }}
    >
      {children}
    </button>
  );
}

/* ---------------------------------------------------------------------- */
/* Main component                                                        */
/* ---------------------------------------------------------------------- */
export default function InventoryDashboard() {
  useGoogleFonts();

  const [loading, setLoading] = useState(true);
  const [products, setProducts] = useState({});
  const [transactions, setTransactions] = useState([]);
  const [tab, setTab] = useState("overview");
  const [now, setNow] = useState(new Date());
  const [showImport, setShowImport] = useState(false);
  const [importStatus, setImportStatus] = useState(null);
  const [sheetUrl, setSheetUrl] = useState("");
  const [txSearch, setTxSearch] = useState("");
  const [txFilter, setTxFilter] = useState("all"); // "all" | "in" | "out"
  const [horizon, setHorizon] = useState(30); // forecast horizon in days for the forecast tab
  const fileRef = useRef(null);

  useEffect(() => {
    const t = setInterval(() => setNow(new Date()), 1000);
    return () => clearInterval(t);
  }, []);

  useEffect(() => {
    (async () => {
      let p = await loadJSON(K_PRODUCTS, null);
      let tx = await loadJSON(K_TX, null);
      if (!p || !tx) {
        const seed = seedData();
        p = seed.products;
        tx = seed.tx;
        await saveJSON(K_PRODUCTS, p);
        await saveJSON(K_TX, tx);
      }
      setProducts(p);
      setTransactions(tx);
      setLoading(false);
    })();
  }, []);

  // Real-time sync: whenever another visitor imports/edits data, everyone
  // watching the dashboard updates automatically without refreshing.
  useEffect(() => {
    if (!supabase) return;
    const channel = supabase
      .channel("kv_store-changes")
      .on("postgres_changes", { event: "*", schema: "public", table: "kv_store" }, (payload) => {
        const row = payload.new;
        if (!row) return;
        if (row.id === K_PRODUCTS) setProducts(row.value);
        if (row.id === K_TX) setTransactions(row.value);
      })
      .subscribe();
    return () => {
      supabase.removeChannel(channel);
    };
  }, []);

  const persist = useCallback(async (p, tx) => {
    setProducts(p);
    setTransactions(tx);
    await saveJSON(K_PRODUCTS, p);
    await saveJSON(K_TX, tx);
  }, []);

  /* ------------------------- ingest rows from any source ------------------------- */
  // Two supported sheet formats, auto-detected from the header row:
  //   1) explicit "in" / "out" columns (transaction log)
  //   2) a single running quantity column per SKU per date (stock snapshot) —
  //      the system compares it to the previous known reading for that SKU:
  //      number went down  -> treated as "จ่าย/เบิกออก" (out)
  //      number went up    -> treated as "รับเข้า" (in)
  const IN_KEYS = ["in", "รับเข้า", "จำนวนรับ"];
  const OUT_KEYS = ["out", "เบิกออก", "จำนวนเบิก", "จ่ายออก"];
  const QTY_KEYS = [
    "quantity",
    "qty",
    "stock",
    "balance",
    "available qty",
    "availableqty",
    "available quantity",
    "คงเหลือ",
    "จำนวน",
    "จำนวนคงเหลือ",
    "ยอดคงเหลือ",
  ];
  const SKU_KEYS = ["sku", "รหัสสินค้า", "code", "material number", "materialnumber", "material no", "part number", "item code"];
  const NAME_KEYS = ["name", "product", "ชื่อสินค้า", "productname", "material description", "materialdescription", "description", "item description"];
  const UNIT_KEYS = ["unit", "หน่วย", "basic unit of measure", "uom", "unit of measure"];

  const getField = (raw, keys) => {
    for (const k of keys) {
      const found = Object.keys(raw).find((rk) => rk.trim().toLowerCase() === k);
      if (found && raw[found] !== undefined && raw[found] !== "") return raw[found];
    }
    return undefined;
  };

  const normalizeDate = (date) => {
    if (date instanceof Date) return date.toISOString().slice(0, 10);
    if (typeof date === "number") return XLSX.SSF.format("yyyy-mm-dd", date);
    if (typeof date === "string" && date.trim()) return date.trim().slice(0, 10);
    return new Date().toISOString().slice(0, 10);
  };

  const ingestRows = useCallback(
    (rows) => {
      const p = { ...products };
      const tx = [...transactions];
      let nextId = tx.reduce((m, r) => Math.max(m, r.id || 0), 0) + 1;
      let count = 0;

      const headers = rows.length ? Object.keys(rows[0]).map((k) => k.trim().toLowerCase()) : [];
      const hasInOut = headers.some((h) => IN_KEYS.includes(h) || OUT_KEYS.includes(h));
      const hasQty = headers.some((h) => QTY_KEYS.includes(h));
      const snapshotMode = hasQty && !hasInOut;

      const applyMasterFields = (sku, name, unit, leadTimeDays, safetyStock) => {
        if (!p[sku]) {
          p[sku] = { sku, name: name || sku, unit: unit || "ชิ้น", stock: 0, leadTimeDays: Number(leadTimeDays) || 3, safetyStock: Number(safetyStock) || 0 };
        } else {
          if (name) p[sku].name = name;
          if (unit) p[sku].unit = unit;
          if (leadTimeDays !== undefined) p[sku].leadTimeDays = Number(leadTimeDays);
          if (safetyStock !== undefined) p[sku].safetyStock = Number(safetyStock);
        }
      };

      if (snapshotMode) {
        // parse + normalize every row first
        const parsed = rows
          .map((raw) => {
            const sku = String(getField(raw, SKU_KEYS) ?? "").trim();
            if (!sku) return null;
            return {
              sku,
              name: String(getField(raw, NAME_KEYS) ?? "").trim(),
              unit: String(getField(raw, UNIT_KEYS) ?? "").trim(),
              date: normalizeDate(getField(raw, ["date", "วันที่"])),
              quantity: Number(getField(raw, QTY_KEYS)),
              leadTimeDays: getField(raw, ["leadtimedays", "leadtime", "ระยะเวลาสั่งซื้อ"]),
              safetyStock: getField(raw, ["safetystock", "สต๊อกสำรอง"]),
              note: getField(raw, ["note", "หมายเหตุ"]) || "",
            };
          })
          .filter(Boolean);

        const bySku = {};
        parsed.forEach((r) => {
          (bySku[r.sku] = bySku[r.sku] || []).push(r);
        });

        Object.entries(bySku).forEach(([sku, items]) => {
          items.sort((a, b) => (a.date < b.date ? -1 : a.date > b.date ? 1 : 0));
          const first = items[0];
          applyMasterFields(sku, first.name, first.unit, first.leadTimeDays, first.safetyStock);

          let prevQty = p[sku].stock && p[sku].stock !== 0 ? Number(p[sku].stock) : null;
          // if the product already had a nonzero known stock, treat that as the baseline
          // before the very first row in this upload; otherwise the first row becomes
          // the opening balance (no in/out inferred for it).
          items.forEach((item) => {
            applyMasterFields(sku, item.name, item.unit, item.leadTimeDays, item.safetyStock);
            const qty = Number.isFinite(item.quantity) ? item.quantity : prevQty ?? 0;

            if (prevQty === null) {
              p[sku].stock = qty;
              tx.push({ id: nextId++, date: item.date, sku, name: p[sku].name, unit: p[sku].unit, in: 0, out: 0, note: item.note || "ยอดเริ่มต้น" });
            } else {
              const delta = qty - prevQty;
              const inQty = delta > 0 ? delta : 0;
              const outQty = delta < 0 ? -delta : 0;
              p[sku].stock = qty;
              if (delta !== 0) {
                tx.push({
                  id: nextId++,
                  date: item.date,
                  sku,
                  name: p[sku].name,
                  unit: p[sku].unit,
                  in: inQty,
                  out: outQty,
                  note: item.note || (delta > 0 ? "รับเข้า (คำนวณจากยอดที่เพิ่มขึ้น)" : "เบิกจ่าย (คำนวณจากยอดที่ลดลง)"),
                });
              }
            }
            prevQty = qty;
            count++;
          });
        });

        return { p, tx, count };
      }

      // --- legacy explicit in / out column mode ---
      rows.forEach((raw) => {
        const sku = String(getField(raw, SKU_KEYS) ?? "").trim();
        if (!sku) return;
        const name = String(getField(raw, NAME_KEYS) ?? sku).trim();
        const unit = String(getField(raw, UNIT_KEYS) ?? p[sku]?.unit ?? "ชิ้น").trim();
        const date = normalizeDate(getField(raw, ["date", "วันที่"]));
        const inQty = Number(getField(raw, IN_KEYS)) || 0;
        const outQty = Number(getField(raw, OUT_KEYS)) || 0;
        const leadTimeDays = getField(raw, ["leadtimedays", "leadtime", "ระยะเวลาสั่งซื้อ"]);
        const safetyStock = getField(raw, ["safetystock", "สต๊อกสำรอง"]);
        const note = getField(raw, ["note", "หมายเหตุ"]) || "";

        applyMasterFields(sku, name, unit, leadTimeDays, safetyStock);
        p[sku].stock = Number(p[sku].stock || 0) + inQty - outQty;

        tx.push({ id: nextId++, date, sku, name: p[sku].name, unit: p[sku].unit, in: inQty, out: outQty, note });
        count++;
      });

      return { p, tx, count };
    },
    [products, transactions]
  );

  const handleFile = useCallback(
    async (file) => {
      setImportStatus({ type: "loading", msg: "กำลังอ่านไฟล์..." });
      try {
        const buf = await file.arrayBuffer();
        const wb = XLSX.read(buf, { type: "array", cellDates: true });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        const { p, tx, count } = ingestRows(rows);
        await persist(p, tx);
        setImportStatus({ type: "success", msg: `นำเข้าสำเร็จ ${count} รายการ จาก ${file.name}` });
      } catch (e) {
        console.error(e);
        setImportStatus({ type: "error", msg: "อ่านไฟล์ไม่สำเร็จ ตรวจสอบรูปแบบไฟล์ Excel/CSV" });
      }
    },
    [ingestRows, persist]
  );

  const handleSheetUrl = useCallback(async () => {
    if (!sheetUrl.trim()) return;
    setImportStatus({ type: "loading", msg: "กำลังดึงข้อมูลจากลิงก์..." });
    try {
      const res = await fetch(sheetUrl.trim());
      if (!res.ok) throw new Error("fetch failed");
      const text = await res.text();
      const parsed = Papa.parse(text, { header: true, skipEmptyLines: true });
      const { p, tx, count } = ingestRows(parsed.data);
      await persist(p, tx);
      setImportStatus({ type: "success", msg: `นำเข้าสำเร็จ ${count} รายการ จาก Google Sheet` });
    } catch (e) {
      console.error(e);
      setImportStatus({
        type: "error",
        msg: "ดึงข้อมูลไม่สำเร็จ ตรวจสอบว่าลิงก์เป็นแบบเผยแพร่เป็น CSV (File → Share → Publish to web → .csv) และเป็นสาธารณะ",
      });
    }
  }, [sheetUrl, ingestRows, persist]);

  const resetDemo = useCallback(async () => {
    const seed = seedData();
    await persist(seed.products, seed.tx);
  }, [persist]);

  /* ------------------------------- derived data ------------------------------- */
  const productList = useMemo(() => Object.values(products), [products]);

  const enriched = useMemo(
    () =>
      productList
        .map((p) => ({ ...p, m: metricsFor(p, transactions, horizon) }))
        .sort((a, b) => {
          const rank = { วิกฤต: 0, ควรสั่งซื้อ: 1, ปกติ: 2 };
          return rank[a.m.status] - rank[b.m.status] || a.name.localeCompare(b.name, "th");
        }),
    [productList, transactions, horizon]
  );

  const trend = useMemo(() => dailyTrend(transactions), [transactions]);
  const todayStr = new Date().toISOString().slice(0, 10);
  const todayIn = transactions.filter((t) => t.date === todayStr).reduce((s, r) => s + (Number(r.in) || 0), 0);
  const todayOut = transactions.filter((t) => t.date === todayStr).reduce((s, r) => s + (Number(r.out) || 0), 0);
  const reorderCount = enriched.filter((p) => p.m.status !== "ปกติ").length;

  const recentTx = useMemo(
    () =>
      [...transactions]
        .sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : b.id - a.id))
        .filter((t) => !txSearch || t.sku.toLowerCase().includes(txSearch.toLowerCase()) || t.name.toLowerCase().includes(txSearch.toLowerCase()))
        .filter((t) => {
          if (txFilter === "in") return Number(t.in) > 0;
          if (txFilter === "out") return Number(t.out) > 0;
          return true;
        })
        .slice(0, 60),
    [transactions, txSearch, txFilter]
  );

  const clockStr = now.toLocaleTimeString("th-TH", { hour: "2-digit", minute: "2-digit", second: "2-digit" });
  const dateStr = now.toLocaleDateString("th-TH", { weekday: "long", day: "numeric", month: "long", year: "numeric" });

  /* ------------------------------- export to Excel ------------------------------- */
  // Exports two sheets: a per-SKU total summary, and the full (unlimited, filtered)
  // transaction log matching whatever search/filter is currently applied on screen.
  const exportExcel = useCallback(() => {
    const filteredForExport = [...transactions]
      .sort((a, b) => (a.date < b.date ? 1 : a.date > b.date ? -1 : b.id - a.id))
      .filter((t) => !txSearch || t.sku.toLowerCase().includes(txSearch.toLowerCase()) || t.name.toLowerCase().includes(txSearch.toLowerCase()))
      .filter((t) => {
        if (txFilter === "in") return Number(t.in) > 0;
        if (txFilter === "out") return Number(t.out) > 0;
        return true;
      });

    const txRows = filteredForExport.map((t) => ({
      วันที่: t.date,
      รหัสสินค้า: t.sku,
      ชื่อสินค้า: t.name,
      หน่วย: t.unit,
      รับเข้า: t.in || 0,
      เบิกจ่าย: t.out || 0,
      หมายเหตุ: t.note || "",
    }));

    const summaryMap = {};
    transactions.forEach((t) => {
      if (!summaryMap[t.sku]) {
        summaryMap[t.sku] = { รหัสสินค้า: t.sku, ชื่อสินค้า: t.name, หน่วย: t.unit, รับเข้ารวม: 0, เบิกจ่ายรวม: 0 };
      }
      summaryMap[t.sku].รับเข้ารวม += Number(t.in) || 0;
      summaryMap[t.sku].เบิกจ่ายรวม += Number(t.out) || 0;
    });
    const summaryRows = Object.values(summaryMap).map((s) => ({
      ...s,
      คงเหลือปัจจุบัน: products[s.รหัสสินค้า]?.stock ?? "",
    }));

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(summaryRows), "สรุปตามสินค้า");
    XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(txRows), "ประวัติรับ-จ่าย");
    XLSX.writeFile(wb, `รายงานรับจ่าย_${new Date().toISOString().slice(0, 10)}.xlsx`);
  }, [transactions, products, txSearch, txFilter]);

  /* ------------------------------- render ------------------------------- */
  return (
    <div
      style={{
        fontFamily: "'IBM Plex Sans Thai', 'IBM Plex Sans', system-ui, sans-serif",
        background: COLORS.bg,
        color: COLORS.text,
        minHeight: "100%",
        padding: "22px 26px 60px",
        boxSizing: "border-box",
      }}
    >
      {/* Header */}
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "flex-start", flexWrap: "wrap", gap: 14, marginBottom: 20 }}>
        <div>
          <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
            <BoxSelect size={22} color={COLORS.accent} />
            <h1 style={{ margin: 0, fontSize: 21, fontWeight: 700, letterSpacing: 0.2 }}>คลังสินค้า · แดชบอร์ดเรียลไทม์</h1>
          </div>
          <div style={{ color: COLORS.muted, fontSize: 13.5, marginTop: 4 }}>{dateStr}</div>
        </div>
        <div style={{ display: "flex", alignItems: "center", gap: 10 }}>
          <div
            style={{
              fontFamily: "'IBM Plex Mono', monospace",
              background: COLORS.surface,
              border: `1px solid ${COLORS.border}`,
              borderRadius: 10,
              padding: "8px 14px",
              fontSize: 15,
              color: COLORS.success,
              display: "flex",
              alignItems: "center",
              gap: 8,
            }}
          >
            <span style={{ width: 8, height: 8, borderRadius: 999, background: COLORS.success, boxShadow: `0 0 0 3px rgba(95,183,126,0.2)` }} />
            {clockStr}
          </div>
          <button
            onClick={() => setShowImport(true)}
            style={{
              background: COLORS.accent,
              color: "#0C1418",
              border: "none",
              borderRadius: 10,
              padding: "10px 16px",
              fontSize: 14,
              fontWeight: 600,
              cursor: "pointer",
              display: "flex",
              alignItems: "center",
              gap: 8,
            }}
          >
            <Upload size={16} /> นำเข้าข้อมูล
          </button>
        </div>
      </div>

      {/* KPI row */}
      <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(190px, 1fr))", gap: 14, marginBottom: 20 }}>
        <KPI icon={PackageSearch} label="จำนวนสินค้าทั้งหมด" value={productList.length} sub="รายการ SKU" />
        <KPI icon={TrendingUp} label="รับเข้าวันนี้" value={todayIn.toLocaleString("th-TH")} sub="หน่วยสะสม" />
        <KPI icon={TrendingDown} label="เบิกจ่ายวันนี้" value={todayOut.toLocaleString("th-TH")} sub="หน่วยสะสม" />
        <KPI
          icon={AlertTriangle}
          label="ต้องสั่งซื้อเพิ่ม"
          value={reorderCount}
          sub={reorderCount ? "ตรวจสอบที่แท็บพยากรณ์" : "สต๊อกเพียงพอทั้งหมด"}
          tone={reorderCount ? "warn" : undefined}
        />
      </div>

      {/* Tabs */}
      <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
        <Tab active={tab === "overview"} onClick={() => setTab("overview")}>
          ภาพรวม
        </Tab>
        <Tab active={tab === "forecast"} onClick={() => setTab("forecast")}>
          พยากรณ์ & สั่งซื้อ
        </Tab>
        <Tab active={tab === "transactions"} onClick={() => setTab("transactions")}>
          ประวัติรับ-จ่าย
        </Tab>
      </div>

      {loading ? (
        <div style={{ color: COLORS.muted, padding: 40, textAlign: "center" }}>กำลังโหลดข้อมูล...</div>
      ) : (
        <>
          {tab === "overview" && (
            <div style={{ display: "grid", gridTemplateColumns: "minmax(0,2fr) minmax(280px,1fr)", gap: 16 }}>
              <div style={{ background: COLORS.surface, border: `1px solid ${COLORS.border}`, borderRadius: 14, padding: "18px 20px" }}>
                <div style={{ fontSize: 14.5, fontWeight: 600, marginBottom: 12 }}>แนวโน้มรับเข้า / เบิกจ่าย ({TREND_WINDOW} วันล่าสุด)</div>
                <ResponsiveContainer width="100%" height={300}>
                  <ComposedChart data={trend}>
                    <CartesianGrid stroke={COLORS.border} vertical={false} />
                    <XAxis dataKey="label" stroke={COLORS.faint} fontSize={12} />
                    <YAxis stroke={COLORS.faint} fontSize={12} />
                    <Tooltip
                      contentStyle={{ background: COLORS.surface2, border: `1px solid ${COLORS.border}`, borderRadius: 8, fontSize: 13 }}
                      labelStyle={{ color: COLORS.text }}
                    />
                    <Legend wrapperStyle={{ fontSize: 12.5 }} />
                    <Bar dataKey="in" name="รับเข้า" fill={COLORS.accent} radius={[4, 4, 0, 0]} />
                    <Bar dataKey="out" name="เบิกจ่าย" fill={COLORS.warn} radius={[4, 4, 0, 0]} />
                    <Line type="monotone" dataKey="out" name="แนวโน้มเบิกจ่าย" stroke={COLORS.danger} dot={false} strokeWidth={2} />
                  </ComposedChart>
                </ResponsiveContainer>
              </div>

              <div style={{ background: COLORS.surface, border: `1px solid ${COLORS.border}`, borderRadius: 14, padding: "18px 20px" }}>
                <div style={{ fontSize: 14.5, fontWeight: 600, marginBottom: 12 }}>สินค้าที่ต้องเฝ้าระวัง</div>
                <div style={{ display: "flex", flexDirection: "column", gap: 4 }}>
                  {enriched
                    .filter((p) => p.m.status !== "ปกติ")
                    .slice(0, 8)
                    .map((p) => (
                      <div key={p.sku} style={{ display: "flex", justifyContent: "space-between", alignItems: "center", padding: "9px 0", fontSize: 13.5, borderBottom: `1px solid ${COLORS.border}` }}>
                        <div>
                          <div style={{ color: COLORS.text }}>{p.name}</div>
                          <div style={{ color: COLORS.faint, fontSize: 11.5, fontFamily: "'IBM Plex Mono', monospace" }}>
                            คงเหลือ {p.stock.toLocaleString("th-TH")} {p.unit}
                          </div>
                        </div>
                        <Badge status={p.m.status} />
                      </div>
                    ))}
                  {reorderCount === 0 && <div style={{ fontSize: 13, color: COLORS.faint }}>ไม่มีสินค้าที่ต้องเฝ้าระวังในขณะนี้</div>}
                </div>
              </div>
            </div>
          )}

          {tab === "forecast" && (
            <div style={{ background: COLORS.surface, border: `1px solid ${COLORS.border}`, borderRadius: 14, overflow: "hidden" }}>
              <div
                style={{
                  padding: "14px 20px",
                  borderBottom: `1px solid ${COLORS.border}`,
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "center",
                  flexWrap: "wrap",
                  gap: 12,
                }}
              >
                <div style={{ fontSize: 13, color: COLORS.muted, maxWidth: 480 }}>
                  พยากรณ์ความต้องการล่วงหน้าจากยอดเบิกจ่ายเฉลี่ยย้อนหลัง {TREND_WINDOW} วัน คูณด้วยช่วงเวลาที่เลือก แล้วเทียบกับสต๊อกคงเหลือ
                </div>
                <div style={{ display: "flex", gap: 6 }}>
                  {[7, 14, 30, 60].map((d) => (
                    <FilterChip key={d} active={horizon === d} onClick={() => setHorizon(d)} color={COLORS.accent}>
                      {d} วัน
                    </FilterChip>
                  ))}
                </div>
              </div>
              <div style={{ overflowX: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13.5 }}>
                  <thead>
                    <tr style={{ textAlign: "left", color: COLORS.muted, fontSize: 12.5 }}>
                      {[
                        "สินค้า",
                        "คงเหลือ",
                        "เฉลี่ย/วัน",
                        `พยากรณ์ความต้องการ ${horizon} วัน`,
                        `คาดว่าคงเหลือหลัง ${horizon} วัน`,
                        "ของจะหมดใน (วัน)",
                        "แนะนำสั่งเพิ่ม",
                        "สถานะ",
                      ].map((h) => (
                        <th key={h} style={{ padding: "10px 18px", fontWeight: 500, whiteSpace: "nowrap" }}>
                          {h}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {enriched.map((p) => (
                      <tr key={p.sku} style={{ borderTop: `1px solid ${COLORS.border}` }}>
                        <td style={{ padding: "12px 18px" }}>
                          <div style={{ fontWeight: 500 }}>{p.name}</div>
                          <div style={{ color: COLORS.faint, fontSize: 12, fontFamily: "'IBM Plex Mono', monospace" }}>{p.sku}</div>
                        </td>
                        <td style={numTd}>
                          {p.stock.toLocaleString("th-TH")} {p.unit}
                        </td>
                        <td style={numTd}>{p.m.avgDailyOut.toFixed(1)}</td>
                        <td style={{ ...numTd, fontWeight: 600 }}>
                          {p.m.forecastQtyHorizon.toLocaleString("th-TH")} {p.unit}
                        </td>
                        <td style={{ ...numTd, color: p.m.projectedStockAtHorizon < 0 ? COLORS.danger : COLORS.text, fontWeight: 600 }}>
                          {p.m.projectedStockAtHorizon.toLocaleString("th-TH")} {p.unit}
                        </td>
                        <td style={numTd}>{p.m.stockOutInDays !== null ? p.m.stockOutInDays.toLocaleString("th-TH") : "—"}</td>
                        <td style={{ ...numTd, color: p.m.suggestedQty > 0 ? COLORS.warn : COLORS.faint, fontWeight: 600 }}>
                          {p.m.suggestedQty > 0 ? `+${p.m.suggestedQty.toLocaleString("th-TH")} ${p.unit}` : "-"}
                        </td>
                        <td style={{ padding: "12px 18px" }}>
                          <Badge status={p.m.status} />
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}

          {tab === "transactions" && (
            <div style={{ background: COLORS.surface, border: `1px solid ${COLORS.border}`, borderRadius: 14, overflow: "hidden" }}>
              <div
                style={{
                  padding: "14px 20px",
                  borderBottom: `1px solid ${COLORS.border}`,
                  display: "flex",
                  justifyContent: "space-between",
                  alignItems: "center",
                  flexWrap: "wrap",
                  gap: 12,
                }}
              >
                <div style={{ display: "flex", gap: 10, flexWrap: "wrap", flex: 1, minWidth: 260 }}>
                  <input
                    placeholder="ค้นหาด้วยชื่อสินค้าหรือ SKU..."
                    value={txSearch}
                    onChange={(e) => setTxSearch(e.target.value)}
                    style={{ ...inpStyle, maxWidth: 300 }}
                  />
                  <div style={{ display: "flex", gap: 6 }}>
                    <FilterChip active={txFilter === "all"} onClick={() => setTxFilter("all")} color={COLORS.accent}>
                      ทั้งหมด
                    </FilterChip>
                    <FilterChip active={txFilter === "in"} onClick={() => setTxFilter("in")} color={COLORS.success}>
                      <TrendingUp size={13} /> รับเข้า
                    </FilterChip>
                    <FilterChip active={txFilter === "out"} onClick={() => setTxFilter("out")} color={COLORS.warn}>
                      <TrendingDown size={13} /> เบิกจ่าย
                    </FilterChip>
                  </div>
                </div>
                <div style={{ display: "flex", alignItems: "center", gap: 12, flexWrap: "wrap" }}>
                  <span style={{ color: COLORS.faint, fontSize: 12.5, whiteSpace: "nowrap" }}>
                    แสดง {recentTx.length.toLocaleString("th-TH")} จาก {transactions.length.toLocaleString("th-TH")} รายการ
                  </span>
                  <button
                    onClick={exportExcel}
                    style={{
                      background: COLORS.surface2,
                      color: COLORS.text,
                      border: `1px solid ${COLORS.border}`,
                      borderRadius: 8,
                      padding: "8px 14px",
                      fontSize: 13,
                      fontWeight: 600,
                      cursor: "pointer",
                      display: "flex",
                      alignItems: "center",
                      gap: 7,
                      whiteSpace: "nowrap",
                    }}
                  >
                    <Download size={14} color={COLORS.accent} /> ส่งออก Excel
                  </button>
                </div>
              </div>
              <div style={{ overflowX: "auto", maxHeight: 460, overflowY: "auto" }}>
                <table style={{ width: "100%", borderCollapse: "collapse", fontSize: 13.5 }}>
                  <thead>
                    <tr style={{ textAlign: "left", color: COLORS.muted, fontSize: 12.5, position: "sticky", top: 0, background: COLORS.surface }}>
                      {["วันที่", "สินค้า", "รับเข้า", "เบิกออก", "หมายเหตุ"].map((h) => (
                        <th key={h} style={{ padding: "10px 18px", fontWeight: 500 }}>
                          {h}
                        </th>
                      ))}
                    </tr>
                  </thead>
                  <tbody>
                    {recentTx.map((t) => (
                      <tr key={t.id} style={{ borderTop: `1px solid ${COLORS.border}` }}>
                        <td style={{ padding: "10px 18px", fontFamily: "'IBM Plex Mono', monospace", color: COLORS.muted }}>{t.date}</td>
                        <td style={{ padding: "10px 18px" }}>
                          {t.name} <span style={{ color: COLORS.faint, fontSize: 11.5 }}>({t.sku})</span>
                        </td>
                        <td style={{ ...numTd, color: t.in ? COLORS.success : COLORS.faint }}>{t.in || "-"}</td>
                        <td style={{ ...numTd, color: t.out ? COLORS.warn : COLORS.faint }}>{t.out || "-"}</td>
                        <td style={{ padding: "10px 18px", color: COLORS.faint }}>{t.note}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </>
      )}

      {/* Import modal */}
      {showImport && (
        <div
          onClick={() => setShowImport(false)}
          style={{ position: "fixed", inset: 0, background: "rgba(0,0,0,0.55)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 50 }}
        >
          <div
            onClick={(e) => e.stopPropagation()}
            style={{ background: COLORS.surface, border: `1px solid ${COLORS.border}`, borderRadius: 16, padding: 24, width: 460, maxWidth: "90vw" }}
          >
            <div style={{ fontSize: 16, fontWeight: 700, marginBottom: 4 }}>นำเข้าข้อมูลรับ-จ่ายสินค้า</div>
            <div style={{ fontSize: 12.5, color: COLORS.muted, marginBottom: 16 }}>
              รองรับ 2 รูปแบบไฟล์ (ตรวจจับอัตโนมัติจากหัวคอลัมน์):
              <br />• แบบระบุรับ-จ่ายตรง: date, sku, name, unit, in, out
              <br />• แบบตัวเลขคงเหลือต่อวัน: date, sku, name, unit, quantity (ระบบจะเทียบกับยอดครั้งก่อนของ SKU เดียวกัน — ตัวเลขลดลง = เบิกจ่าย, เพิ่มขึ้น = รับเข้า)
              <br />รองรับหัวคอลัมน์แบบระบบ ERP/SAP ด้วย เช่น Material Number, Material Description, Available Qty, Basic Unit of Measure
              <br />คอลัมน์เสริม: leadTimeDays, safetyStock, note (รองรับทั้งชื่อภาษาไทยและอังกฤษ)
            </div>

            <div
              onClick={() => fileRef.current?.click()}
              onDragOver={(e) => e.preventDefault()}
              onDrop={(e) => {
                e.preventDefault();
                if (e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
              }}
              style={{
                border: `1.5px dashed ${COLORS.border}`,
                borderRadius: 12,
                padding: "26px 16px",
                textAlign: "center",
                cursor: "pointer",
                color: COLORS.muted,
                marginBottom: 16,
              }}
            >
              <Upload size={20} style={{ marginBottom: 8 }} />
              <div style={{ fontSize: 13.5 }}>ลากไฟล์ .xlsx / .csv มาวาง หรือคลิกเพื่อเลือกไฟล์</div>
              <input
                ref={fileRef}
                type="file"
                accept=".xlsx,.xls,.csv"
                style={{ display: "none" }}
                onChange={(e) => e.target.files[0] && handleFile(e.target.files[0])}
              />
            </div>

            <div style={{ fontSize: 12.5, color: COLORS.muted, marginBottom: 8 }}>
              หรือวางลิงก์ Google Sheet ที่เผยแพร่เป็น CSV (File → Share → Publish to web → เลือก .csv)
            </div>
            <div style={{ display: "flex", gap: 8, marginBottom: 16 }}>
              <input value={sheetUrl} onChange={(e) => setSheetUrl(e.target.value)} placeholder="https://docs.google.com/.../pub?output=csv" style={{ ...inpStyle, flex: 1 }} />
              <button onClick={handleSheetUrl} style={{ ...primaryBtn, padding: "0 14px" }}>
                <LinkIcon size={15} />
              </button>
            </div>

            {importStatus && (
              <div
                style={{
                  fontSize: 13,
                  padding: "9px 12px",
                  borderRadius: 8,
                  marginBottom: 14,
                  background:
                    importStatus.type === "success" ? "rgba(95,183,126,0.12)" : importStatus.type === "error" ? "rgba(224,100,90,0.12)" : COLORS.surface2,
                  color: importStatus.type === "success" ? COLORS.success : importStatus.type === "error" ? COLORS.danger : COLORS.muted,
                }}
              >
                {importStatus.msg}
              </div>
            )}

            <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
              <button onClick={resetDemo} style={{ ...ghostBtn, color: COLORS.faint }}>
                <Trash2 size={13} /> ล้างข้อมูลตัวอย่าง
              </button>
              <button onClick={() => setShowImport(false)} style={ghostBtn}>
                ปิด <ChevronRight size={14} />
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
}

/* ---------------------------------------------------------------------- */
/* inline style helpers                                                   */
/* ---------------------------------------------------------------------- */
const inpStyle = {
  background: COLORS.surface2,
  border: `1px solid ${COLORS.border}`,
  borderRadius: 8,
  padding: "9px 12px",
  color: COLORS.text,
  fontSize: 13.5,
  fontFamily: "inherit",
  outline: "none",
  width: "100%",
  boxSizing: "border-box",
};
const selStyle = { ...inpStyle };
const numTd = { padding: "12px 18px", fontFamily: "'IBM Plex Mono', monospace", color: COLORS.text };
const primaryBtn = {
  background: COLORS.accent,
  color: "#0C1418",
  border: "none",
  borderRadius: 8,
  padding: "10px 14px",
  fontSize: 13.5,
  fontWeight: 600,
  cursor: "pointer",
  display: "flex",
  alignItems: "center",
  justifyContent: "center",
  gap: 6,
};
const ghostBtn = {
  background: "transparent",
  border: `1px solid ${COLORS.border}`,
  color: COLORS.muted,
  borderRadius: 8,
  padding: "8px 12px",
  fontSize: 12.5,
  cursor: "pointer",
  display: "flex",
  alignItems: "center",
  gap: 6,
};
