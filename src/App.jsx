import React, { useState, useMemo, useRef } from "react";
import * as XLSX from "xlsx";
import { Plus, Trash2, Download, FileSpreadsheet } from "lucide-react";

const uid = () => Math.random().toString(36).slice(2, 9);

const money = (n) =>
  (isFinite(n) ? n : 0).toLocaleString("en-SG", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

export default function EventBudgetGenerator() {
  const [docType, setDocType] = useState("Proposed Budget");
  const [orgName, setOrgName] = useState("Tampines Greencourt RN");
  const [eventName, setEventName] = useState("");
  const [eventDate, setEventDate] = useState("");
  const [eventTime, setEventTime] = useState("");

  const [income, setIncome] = useState([
    { id: uid(), label: "Registration Fee", amount: "" },
  ]);
  const [expenditure, setExpenditure] = useState([
    { id: uid(), label: "Item A", amount: "" },
    { id: uid(), label: "Item B", amount: "" },
  ]);

  const [preparedName, setPreparedName] = useState("");
  const [preparedDesignation, setPreparedDesignation] = useState("");
  const [approverName, setApproverName] = useState("Effendy Chua");
  const [approverDesignation, setApproverDesignation] = useState("Chairman");
  const [directorName, setDirectorName] = useState("Natalie Tan");
  const [directorOrg, setDirectorOrg] = useState("Tampines Boulevard CO");

  const fileRef = useRef(null);

  const totalIncome = useMemo(
    () => income.reduce((s, r) => s + (parseFloat(r.amount) || 0), 0),
    [income]
  );
  const totalExpenditure = useMemo(
    () => expenditure.reduce((s, r) => s + (parseFloat(r.amount) || 0), 0),
    [expenditure]
  );
  const net = totalIncome - totalExpenditure;
  const isDeficit = net < 0;

  const updateRow = (setter, id, field, value) =>
    setter((rows) => rows.map((r) => (r.id === id ? { ...r, [field]: value } : r)));
  const addRow = (setter, label = "") =>
    setter((rows) => [...rows, { id: uid(), label, amount: "" }]);
  const removeRow = (setter, id) =>
    setter((rows) => rows.filter((r) => r.id !== id));

  const missing = [];
  if (!eventName.trim()) missing.push("Event name");
  if (!eventDate.trim()) missing.push("Event date");
  if (!preparedName.trim()) missing.push("Prepared-by name");

  function buildWorkbook() {
    const rows = [];
    const push = (arr) => rows.push(arr);

    push([docType]);
    push([`Event: ${eventName || "[Event Name]"}`]);
    push([`Date: ${eventDate || "[Event Date]"}`]);
    push([`Time: ${eventTime || "[Event Time]"}`]);
    push([orgName]);
    push([]);
    push(["INCOME", "", "", "", "", "S$"]);
    const incomeStart = rows.length + 1;
    income.forEach((r) => push([r.label || "", "", "", "", "", parseFloat(r.amount) || 0]));
    const incomeEnd = rows.length;
    push([
      "TOTAL INCOME", "", "", "", "",
      { t: "n", f: `SUM(F${incomeStart}:F${incomeEnd})` },
    ]);
    const totalIncomeRow = rows.length;
    push([]);
    push(["EXPENDITURE", "", "", "", "", "S$"]);
    const expStart = rows.length + 1;
    expenditure.forEach((r) => push([r.label || "", "", "", "", "", parseFloat(r.amount) || 0]));
    const expEnd = rows.length;
    push([
      "TOTAL EXPENDITURE", "", "", "", "",
      { t: "n", f: `SUM(F${expStart}:F${expEnd})` },
    ]);
    const totalExpRow = rows.length;
    push([]);
    push([
      "SURPLUS / (DEFICIT)", "", "", "", "",
      { t: "n", f: `F${totalIncomeRow}-F${totalExpRow}` },
    ]);
    push([]);
    push([]);
    push(["Prepared by:", "", "", "", "Approved by:", ""]);
    push([]);
    push([]);
    push(["Name:", preparedName, "", "", "Name:", approverName]);
    push(["Designation:", preparedDesignation, "", "", "Designation:", approverDesignation]);
    push(["", "", "", "", "", orgName]);
    push([]);
    push(["", "", "", "", "Date of Approval by Committee:", ""]);
    push([]);
    push([]);
    push(["Certified Correct & True Copy by:"]);
    push([]);
    push([]);
    push([]);
    push([directorName]);
    push(["Constituency Director"]);
    push([directorOrg]);

    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [
      { wch: 22 }, { wch: 18 }, { wch: 12 }, { wch: 10 }, { wch: 20 }, { wch: 14 },
    ];
    ws["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } },
    ];
    // currency format on amount column (F)
    for (let r = 0; r < rows.length; r++) {
      const addr = XLSX.utils.encode_cell({ r, c: 5 });
      if (ws[addr]) ws[addr].z = "$#,##0.00;($#,##0.00)";
    }

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, docType === "Proposed Budget" ? "Budget" : "Statement of Accounts");
    return wb;
  }

  function handleDownload() {
    const wb = buildWorkbook();
    const safeName = (eventName || "event").replace(/[^a-z0-9]+/gi, "_");
    const fileLabel = docType === "Proposed Budget" ? "Budget" : "Statement_of_Accounts";
    XLSX.writeFile(wb, `${fileLabel}_${safeName}.xlsx`);
  }

  const inputCls =
    "w-full rounded-md border border-slate-300 bg-white px-3 py-2 text-sm text-slate-800 placeholder-slate-400 focus:outline-none focus:ring-2 focus:ring-[#1B2A4A] focus:border-transparent";
  const labelCls = "block text-xs font-semibold uppercase tracking-wide text-slate-500 mb-1";

  return (
    <div className="min-h-screen bg-slate-100 py-8 px-4">
      <div className="mx-auto max-w-6xl">
        <div className="mb-6 flex items-center gap-3">
          <div className="flex h-10 w-10 items-center justify-center rounded-md bg-[#1B2A4A] text-white">
            <FileSpreadsheet size={20} />
          </div>
          <div>
            <h1 className="text-xl font-bold text-slate-900">Event Financial Document Generator</h1>
            <p className="text-sm text-slate-500">Prototype — Budget approval &amp; statement of accounts, in one form</p>
          </div>
        </div>

        <div className="grid grid-cols-1 lg:grid-cols-5 gap-6">
          {/* FORM */}
          <div className="lg:col-span-3 space-y-5">
            <div className="rounded-lg border border-slate-200 bg-white p-5 shadow-sm">
              <div className="mb-4 flex gap-2">
                {["Proposed Budget", "Statement of Accounts"].map((t) => (
                  <button
                    key={t}
                    onClick={() => setDocType(t)}
                    className={`flex-1 rounded-md px-3 py-2 text-sm font-medium transition ${
                      docType === t
                        ? "bg-[#1B2A4A] text-white"
                        : "bg-slate-100 text-slate-600 hover:bg-slate-200"
                    }`}
                  >
                    {t}
                  </button>
                ))}
              </div>
              <p className="text-xs text-slate-500 mb-4">
                {docType === "Proposed Budget"
                  ? "Estimated figures, for Chairperson approval before the event."
                  : "Actual figures after the event, for certification and filing."}
              </p>

              <div className="grid grid-cols-2 gap-3">
                <div className="col-span-2">
                  <label className={labelCls}>Residents' Network</label>
                  <input className={inputCls} value={orgName} onChange={(e) => setOrgName(e.target.value)} />
                </div>
                <div className="col-span-2">
                  <label className={labelCls}>Event Name</label>
                  <input
                    className={inputCls}
                    placeholder="e.g. Mid-Autumn Family Carnival"
                    value={eventName}
                    onChange={(e) => setEventName(e.target.value)}
                  />
                </div>
                <div>
                  <label className={labelCls}>Event Date</label>
                  <input
                    className={inputCls}
                    type="date"
                    value={eventDate}
                    onChange={(e) => setEventDate(e.target.value)}
                  />
                </div>
                <div>
                  <label className={labelCls}>Event Time</label>
                  <input
                    className={inputCls}
                    placeholder="e.g. 4:00 PM – 7:00 PM"
                    value={eventTime}
                    onChange={(e) => setEventTime(e.target.value)}
                  />
                </div>
              </div>
            </div>

            <LineItemCard
              title="Income"
              rows={income}
              onChange={(id, f, v) => updateRow(setIncome, id, f, v)}
              onAdd={() => addRow(setIncome)}
              onRemove={(id) => removeRow(setIncome, id)}
              total={totalIncome}
            />
            <LineItemCard
              title="Expenditure"
              rows={expenditure}
              onChange={(id, f, v) => updateRow(setExpenditure, id, f, v)}
              onAdd={() => addRow(setExpenditure)}
              onRemove={(id) => removeRow(setExpenditure, id)}
              total={totalExpenditure}
            />

            <div className="rounded-lg border border-slate-200 bg-white p-5 shadow-sm">
              <h2 className="text-sm font-semibold text-slate-700 mb-3">Sign-off</h2>
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className={labelCls}>Prepared by (Name)</label>
                  <input className={inputCls} value={preparedName} onChange={(e) => setPreparedName(e.target.value)} />
                </div>
                <div>
                  <label className={labelCls}>Designation</label>
                  <input className={inputCls} value={preparedDesignation} onChange={(e) => setPreparedDesignation(e.target.value)} />
                </div>
                <div>
                  <label className={labelCls}>Approved by (Name)</label>
                  <input className={inputCls} value={approverName} onChange={(e) => setApproverName(e.target.value)} />
                </div>
                <div>
                  <label className={labelCls}>Designation</label>
                  <input className={inputCls} value={approverDesignation} onChange={(e) => setApproverDesignation(e.target.value)} />
                </div>
                <div>
                  <label className={labelCls}>Constituency Director</label>
                  <input className={inputCls} value={directorName} onChange={(e) => setDirectorName(e.target.value)} />
                </div>
                <div>
                  <label className={labelCls}>Constituency Office</label>
                  <input className={inputCls} value={directorOrg} onChange={(e) => setDirectorOrg(e.target.value)} />
                </div>
              </div>
            </div>
          </div>

          {/* PREVIEW */}
          <div className="lg:col-span-2">
            <div className="sticky top-6 space-y-3">
              <div
                className="rounded-lg border border-slate-300 shadow-sm p-6"
                style={{ background: "#FAF7F0", fontFamily: "Georgia, 'Times New Roman', serif" }}
              >
                <p className="text-center text-xs tracking-[0.2em] uppercase text-[#B08D57] mb-1">
                  {orgName || "Residents' Network"}
                </p>
                <h2 className="text-center text-lg font-bold text-[#1B2A4A] mb-4">{docType}</h2>
                <div className="text-sm text-slate-700 space-y-1 mb-4 border-b border-slate-300 pb-3">
                  <p><span className="text-slate-500">Event:</span> {eventName || "[Event Name]"}</p>
                  <p><span className="text-slate-500">Date:</span> {eventDate || "[Event Date]"}</p>
                  <p><span className="text-slate-500">Time:</span> {eventTime || "[Event Time]"}</p>
                </div>

                <p className="text-xs font-bold uppercase tracking-wide text-slate-500 mb-1">Income</p>
                {income.map((r) => (
                  <Row key={r.id} label={r.label || "—"} value={r.amount} />
                ))}
                <Row label="Total Income" value={totalIncome} bold />

                <p className="text-xs font-bold uppercase tracking-wide text-slate-500 mt-4 mb-1">Expenditure</p>
                {expenditure.map((r) => (
                  <Row key={r.id} label={r.label || "—"} value={r.amount} />
                ))}
                <Row label="Total Expenditure" value={totalExpenditure} bold />

                <div className="mt-4 pt-3 border-t-2 border-[#1B2A4A] flex justify-between items-baseline">
                  <span className="text-sm font-bold text-[#1B2A4A]">
                    {isDeficit ? "Deficit" : "Surplus"}
                  </span>
                  <span className={`text-base font-bold ${isDeficit ? "text-red-700" : "text-emerald-700"}`}>
                    S$ {money(Math.abs(net))}
                  </span>
                </div>

                <div className="mt-6 grid grid-cols-2 gap-4 text-xs text-slate-600">
                  <div>
                    <p className="border-t border-slate-400 pt-1">{preparedName || "[Prepared by]"}</p>
                    <p className="text-slate-400">{preparedDesignation || "Designation"}</p>
                  </div>
                  <div>
                    <p className="border-t border-slate-400 pt-1">{approverName || "[Approved by]"}</p>
                    <p className="text-slate-400">{approverDesignation || "Designation"}</p>
                  </div>
                </div>
              </div>

              {missing.length > 0 && (
                <p className="text-xs text-amber-700 bg-amber-50 border border-amber-200 rounded-md px-3 py-2">
                  Still needed before it's ready to sign: {missing.join(", ")}.
                </p>
              )}

              <button
                onClick={handleDownload}
                className="w-full flex items-center justify-center gap-2 rounded-md bg-[#1B2A4A] text-white px-4 py-3 text-sm font-semibold hover:bg-[#243a63] transition"
              >
                <Download size={16} />
                Download as Excel (.xlsx)
              </button>
              <p className="text-[11px] text-slate-400 text-center">
                Opens in Excel — ready to print and sign.
              </p>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

function Row({ label, value, bold }) {
  return (
    <div className={`flex justify-between text-sm ${bold ? "font-bold text-[#1B2A4A] mt-1" : "text-slate-700"}`}>
      <span>{label}</span>
      <span>S$ {money(parseFloat(value) || 0)}</span>
    </div>
  );
}

function LineItemCard({ title, rows, onChange, onAdd, onRemove, total }) {
  return (
    <div className="rounded-lg border border-slate-200 bg-white p-5 shadow-sm">
      <div className="flex items-center justify-between mb-3">
        <h2 className="text-sm font-semibold text-slate-700">{title}</h2>
        <span className="text-sm font-semibold text-slate-500">S$ {money(total)}</span>
      </div>
      <div className="space-y-2">
        {rows.map((r) => (
          <div key={r.id} className="flex gap-2">
            <input
              className="flex-1 rounded-md border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-[#1B2A4A]"
              placeholder="Item"
              value={r.label}
              onChange={(e) => onChange(r.id, "label", e.target.value)}
            />
            <input
              className="w-28 rounded-md border border-slate-300 px-3 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-[#1B2A4A]"
              placeholder="0.00"
              type="number"
              value={r.amount}
              onChange={(e) => onChange(r.id, "amount", e.target.value)}
            />
            <button
              onClick={() => onRemove(r.id)}
              className="rounded-md px-2 text-slate-400 hover:text-red-600 hover:bg-red-50"
              aria-label={`Remove ${r.label}`}
            >
              <Trash2 size={16} />
            </button>
          </div>
        ))}
      </div>
      <button
        onClick={onAdd}
        className="mt-3 flex items-center gap-1 text-sm font-medium text-[#1B2A4A] hover:underline"
      >
        <Plus size={14} /> Add item
      </button>
    </div>
  );
}