import React, { useState, useMemo } from "react";
import * as XLSX from "xlsx";
import { Plus, Trash2, Download, Pen } from "lucide-react";

const uid = () => Math.random().toString(36).slice(2, 9);

const money = (n) =>
  (isFinite(n) ? n : 0).toLocaleString("en-SG", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

export default function EventBudgetGenerator() {
  const [docType, setDocType] = useState("Proposed Budget");
  const [orgName, setOrgName] = useState("Tampines Greencourt RN");
  const [eventName, setEventName] = useState("Visit to Snail Farm");
  const [eventDate, setEventDate] = useState("2026-03-01");
  const [eventTime, setEventTime] = useState("10:30AM to 12:00PM");

  const [income, setIncome] = useState([
    { id: uid(), label: "Registration Fee", price: "10", qty: "45" },
    { id: uid(), label: "Infant Fee", price: "0", qty: "5" },
  ]);
  const [expenditure, setExpenditure] = useState([
    { id: uid(), label: "Entrance Fee", price: "20", qty: "45" },
    { id: uid(), label: "Bus", price: "200", qty: "1" },
  ]);

  const [preparedName, setPreparedName] = useState("Vicky Faith Khoo");
  const [preparedDesignation, setPreparedDesignation] = useState("Vice Chairman");

  // Calculations
  const totalIncome = useMemo(
    () => income.reduce((s, r) => s + (parseFloat(r.price) || 0) * (parseFloat(r.qty) || 0), 0),
    [income]
  );
  const totalExpenditure = useMemo(
    () => expenditure.reduce((s, r) => s + (parseFloat(r.price) || 0) * (parseFloat(r.qty) || 0), 0),
    [expenditure]
  );

  const updateRow = (setter, id, field, value) =>
    setter((rows) => rows.map((r) => (r.id === id ? { ...r, [field]: value } : r)));
  const addRow = (setter) =>
    setter((rows) => [...rows, { id: uid(), label: "", price: "", qty: "" }]);
  const removeRow = (setter, id) =>
    setter((rows) => rows.filter((r) => r.id !== id));

  function buildWorkbook() {
    const rows = [];
    const push = (arr) => rows.push(arr);

    push([docType]);
    push([`Event: ${eventName}`]);
    push([`Date: ${eventDate}`]);
    push([`Time: ${eventTime}`]);
    push([orgName]);
    push([]);
    
    push(["INCOME", "", "Unit Price", "Qty", "", "Total S$"]);
    const incomeStart = rows.length + 1;
    income.forEach((r) => {
      const p = parseFloat(r.price) || 0;
      const q = parseFloat(r.qty) || 0;
      push([r.label || "", "", p, q, "", p * q]);
    });
    const incomeEnd = rows.length;
    push([
      "TOTAL INCOME", "", "", "", "",
      { t: "n", f: `SUM(F${incomeStart}:F${incomeEnd})` },
    ]);
    const totalIncomeRow = rows.length;
    push([]);
    
    push(["EXPENDITURE", "", "Unit Price", "Qty", "", "Total S$"]);
    const expStart = rows.length + 1;
    expenditure.forEach((r) => {
      const p = parseFloat(r.price) || 0;
      const q = parseFloat(r.qty) || 0;
      push([r.label || "", "", p, q, "", p * q]);
    });
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
    push(["Prepared by:"]);
    push([preparedName]);
    push([preparedDesignation]);
    push([orgName]);

    const ws = XLSX.utils.aoa_to_sheet(rows);
    ws["!cols"] = [
      { wch: 22 }, { wch: 18 }, { wch: 12 }, { wch: 10 }, { wch: 20 }, { wch: 14 },
    ];
    
    for (let r = 0; r < rows.length; r++) {
      const addrPrice = XLSX.utils.encode_cell({ r, c: 2 });
      if (ws[addrPrice] && typeof ws[addrPrice].v === "number") ws[addrPrice].z = "$#,##0.00;($#,##0.00)";
      
      const addrTotal = XLSX.utils.encode_cell({ r, c: 5 });
      if (ws[addrTotal] && typeof ws[addrTotal].v === "number") ws[addrTotal].z = "$#,##0.00;($#,##0.00)";
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

  const labelCls = "block text-[11px] font-bold uppercase tracking-wider text-slate-500 mb-1.5";
  const inputCls = "w-full rounded-lg border border-slate-200 bg-white px-3 py-2.5 text-sm text-slate-900 focus:outline-none focus:border-slate-400 focus:ring-1 focus:ring-slate-400 transition-shadow";

  return (
    <div className="min-h-screen bg-[#F8F9FA] pb-32 font-sans selection:bg-slate-200">
      <div className="mx-auto max-w-lg p-4 space-y-6">
        
        {/* HEADER */}
        <div className="flex justify-between items-start pt-2">
          <div>
            <h1 className="text-xl font-bold text-slate-900">Budget & Accounts</h1>
            <p className="text-sm text-slate-500 mt-0.5">Event Financial Generator</p>
          </div>
          <span className="bg-blue-50 text-blue-600 px-3 py-1 rounded-full text-xs font-semibold tracking-wide">
            Draft
          </span>
        </div>

        {/* TABS */}
        <div className="bg-slate-200/60 p-1 rounded-xl flex gap-1">
          {["Proposed Budget", "Statement of Accounts"].map((t) => (
            <button
              key={t}
              onClick={() => setDocType(t)}
              className={`flex-1 rounded-lg px-3 py-2.5 text-sm font-semibold transition-all duration-200 ${
                docType === t
                  ? "bg-white text-slate-900 shadow-sm"
                  : "text-slate-500 hover:text-slate-700"
              }`}
            >
              {t}
            </button>
          ))}
        </div>

        {/* GENERAL INFO CARD */}
        <div className="bg-white rounded-2xl p-5 shadow-sm border border-slate-100">
          <div className="flex items-center gap-2.5 mb-5">
            <div className="w-1.5 h-5 rounded-full bg-slate-900" />
            <h2 className="text-lg font-bold text-slate-900">General Info</h2>
          </div>
          <div className="space-y-4">
            <div>
              <label className={labelCls}>Network</label>
              <input className={inputCls} value={orgName} onChange={(e) => setOrgName(e.target.value)} />
            </div>
            <div>
              <label className={labelCls}>Event Name</label>
              <input className={inputCls} value={eventName} onChange={(e) => setEventName(e.target.value)} />
            </div>
            <div className="grid grid-cols-2 gap-3">
              <div>
                <label className={labelCls}>Date</label>
                <input className={inputCls} type="date" value={eventDate} onChange={(e) => setEventDate(e.target.value)} />
              </div>
              <div>
                <label className={labelCls}>Time</label>
                <input className={inputCls} value={eventTime} onChange={(e) => setEventTime(e.target.value)} />
              </div>
            </div>
          </div>
        </div>

        {/* INCOME CARD */}
        <div className="bg-white rounded-2xl p-5 shadow-sm border border-slate-100">
          <div className="flex justify-between items-start mb-5">
            <div>
              <div className="flex items-center gap-2.5 mb-1">
                <div className="w-1.5 h-5 rounded-full bg-emerald-500" />
                <h2 className="text-lg font-bold text-slate-900">Income</h2>
              </div>
              <p className="text-xs text-slate-400 pl-4">Total revenue generated</p>
            </div>
            <div className="text-right">
              <p className="text-[10px] font-bold uppercase tracking-wider text-slate-400 mb-0.5">Total</p>
              <p className="text-lg font-bold text-emerald-500">S$ {money(totalIncome)}</p>
            </div>
          </div>
          
          <div className="space-y-3">
            {income.map((r) => (
              <LineItemRow key={r.id} row={r} onChange={(f, v) => updateRow(setIncome, r.id, f, v)} onRemove={() => removeRow(setIncome, r.id)} />
            ))}
          </div>
          
          <button onClick={() => addRow(setIncome)} className="w-full mt-4 py-3 border-2 border-dashed border-slate-200 hover:border-slate-300 hover:bg-slate-50 transition-colors rounded-xl text-slate-500 text-sm font-semibold flex items-center justify-center gap-2">
            <Plus size={16} /> Add Income Item
          </button>
        </div>

        {/* EXPENDITURE CARD */}
        <div className="bg-white rounded-2xl p-5 shadow-sm border border-slate-100">
          <div className="flex justify-between items-start mb-5">
            <div>
              <div className="flex items-center gap-2.5 mb-1">
                <div className="w-1.5 h-5 rounded-full bg-rose-500" />
                <h2 className="text-lg font-bold text-slate-900">Expenditure</h2>
              </div>
              <p className="text-xs text-slate-400 pl-4">Costs and expenses</p>
            </div>
            <div className="text-right">
              <p className="text-[10px] font-bold uppercase tracking-wider text-slate-400 mb-0.5">Total</p>
              <p className="text-lg font-bold text-rose-500">S$ {money(totalExpenditure)}</p>
            </div>
          </div>

          <div className="space-y-3">
            {expenditure.map((r) => (
              <LineItemRow key={r.id} row={r} onChange={(f, v) => updateRow(setExpenditure, r.id, f, v)} onRemove={() => removeRow(setExpenditure, r.id)} />
            ))}
          </div>

          <button onClick={() => addRow(setExpenditure)} className="w-full mt-4 py-3 border-2 border-dashed border-slate-200 hover:border-slate-300 hover:bg-slate-50 transition-colors rounded-xl text-slate-500 text-sm font-semibold flex items-center justify-center gap-2">
            <Plus size={16} /> Add Expense Item
          </button>
        </div>

        {/* SIGN-OFF CARD */}
        <div className="bg-white rounded-2xl p-5 shadow-sm border border-slate-100">
          <div className="flex items-center gap-2.5 mb-5">
            <div className="w-1.5 h-5 rounded-full bg-slate-400" />
            <h2 className="text-lg font-bold text-slate-900">Sign-off</h2>
          </div>

          <div className="space-y-5">
            <div>
              <label className={labelCls}>Prepared By</label>
              <div className="border border-slate-200 rounded-xl overflow-hidden focus-within:ring-1 focus-within:border-slate-400 focus-within:ring-slate-400">
                <input 
                  className="w-full px-4 py-3 text-sm text-slate-900 font-medium border-b border-slate-100 outline-none" 
                  value={preparedName} 
                  onChange={(e) => setPreparedName(e.target.value)}
                  placeholder="Full Name"
                />
                <input 
                  className="w-full px-4 py-2.5 text-sm text-slate-500 italic bg-slate-50 outline-none" 
                  value={preparedDesignation} 
                  onChange={(e) => setPreparedDesignation(e.target.value)}
                  placeholder="Designation"
                />
              </div>
            </div>

            <div>
              <label className={labelCls}>Approval Signature</label>
              <div className="border-2 border-dashed border-slate-200 rounded-xl bg-slate-50 flex flex-col items-center justify-center py-10 text-slate-400 cursor-pointer hover:bg-slate-100 transition-colors">
                <Pen size={22} className="mb-2 opacity-60" />
                <span className="text-sm font-medium">Tap to sign</span>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* FLOATING ACTION BUTTON */}
      <div className="fixed bottom-0 left-0 right-0 p-4 bg-white/80 backdrop-blur-lg border-t border-slate-100 flex justify-center pb-safe">
        <button
          onClick={handleDownload}
          className="w-full max-w-lg bg-[#0F172A] hover:bg-slate-800 text-white rounded-xl py-4 text-[15px] font-semibold flex items-center justify-center gap-2 shadow-lg shadow-slate-900/20 transition-all active:scale-[0.98]"
        >
          <Download size={18} />
          Generate Excel Report
        </button>
      </div>
    </div>
  );
}

function LineItemRow({ row, onChange, onRemove }) {
  return (
    <div className="flex gap-2 items-center">
      <input
        className="flex-1 rounded-lg border border-slate-200 px-3 py-2.5 text-sm text-slate-900 focus:outline-none focus:border-slate-400"
        placeholder="Item description"
        value={row.label}
        onChange={(e) => onChange("label", e.target.value)}
      />
      
      <div className="relative w-24 shrink-0">
        <span className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400 text-sm font-medium">$</span>
        <input
          className="w-full rounded-lg border border-slate-200 pl-7 pr-2 py-2.5 text-sm text-slate-900 focus:outline-none focus:border-slate-400"
          placeholder="0.00"
          type="number"
          value={row.price}
          onChange={(e) => onChange("price", e.target.value)}
        />
      </div>

      <div className="relative w-20 shrink-0">
        <span className="absolute left-2 top-1/2 -translate-y-1/2 text-slate-400 text-xs font-bold pl-1">x</span>
        <input
          className="w-full rounded-lg border border-slate-200 pl-6 pr-2 py-2.5 text-sm text-slate-900 focus:outline-none focus:border-slate-400"
          placeholder="Qty"
          type="number"
          value={row.qty}
          onChange={(e) => onChange("qty", e.target.value)}
        />
      </div>

      <button
        onClick={onRemove}
        className="text-slate-300 hover:text-rose-500 hover:bg-rose-50 p-2 rounded-lg transition-colors"
        aria-label="Remove item"
      >
        <Trash2 size={18} />
      </button>
    </div>
  );
}