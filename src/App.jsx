import React, { useState, useMemo } from "react";
import * as XLSX from "xlsx";
import { 
  Plus, Trash2, Download, Search, Bell, LayoutDashboard, 
  Wallet, FileText, Settings, HelpCircle, LogOut, 
  Calendar, Clock, TrendingUp, TrendingDown, PenTool, Save 
} from "lucide-react";

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
    { id: uid(), label: "Infant Fee", price: "10", qty: "5" },
  ]);
  const [expenditure, setExpenditure] = useState([
    { id: uid(), label: "Entrance Fee", price: "17.77", qty: "45" },
    { id: uid(), label: "Bus", price: "400", qty: "1" },
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
  const net = totalIncome - totalExpenditure;
  const isDeficit = net < 0;

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

  const inputCls = "w-full bg-slate-50 border border-slate-200 rounded px-3 py-2 text-sm text-slate-800 focus:outline-none focus:border-blue-500 focus:bg-white transition-colors";
  const labelCls = "block text-[11px] font-bold text-slate-700 mb-1.5";

  return (
    <div className="flex h-screen bg-[#F3F4F6] font-sans overflow-hidden selection:bg-blue-100">
      
      {/* SIDEBAR */}
      <div className="w-64 bg-[#181C25] flex flex-col shrink-0">
        <div className="p-6">
          <h1 className="text-white text-xl font-bold tracking-wide">EventFin Pro</h1>
          <p className="text-slate-400 text-xs mt-1">Planner Portal</p>
        </div>
        
        <div className="px-4 mb-6">
          <button className="w-full bg-[#2563EB] hover:bg-blue-600 text-white rounded-md py-2.5 text-sm font-semibold transition-colors">
            + New Event
          </button>
        </div>

        <nav className="flex-1 px-3 space-y-1">
          <SidebarItem icon={<LayoutDashboard size={18}/>} label="Dashboard" />
          <SidebarItem icon={<Wallet size={18}/>} label="Budgets" />
          <SidebarItem icon={<FileText size={18}/>} label="Accounts" active />
          <SidebarItem icon={<TrendingUp size={18}/>} label="Reports" />
          <SidebarItem icon={<Settings size={18}/>} label="Settings" />
        </nav>

        <div className="px-3 pb-6 space-y-1">
          <SidebarItem icon={<HelpCircle size={18}/>} label="Help" />
          <SidebarItem icon={<LogOut size={18}/>} label="Sign Out" />
        </div>
      </div>

      {/* MAIN CONTENT */}
      <div className="flex-1 flex flex-col overflow-hidden">
        
        {/* TOP HEADER */}
        <div className="h-24 bg-white border-b border-slate-200 flex items-center justify-between px-8 shrink-0">
          <div className="flex items-center gap-8">
            <div>
              <h2 className="text-2xl font-bold text-slate-900 leading-tight">Budget &<br/>Accounts</h2>
              <span className="inline-block mt-1 bg-blue-100 text-blue-800 text-[10px] font-bold px-2 py-0.5 rounded tracking-widest uppercase">
                Draft
              </span>
            </div>

            <div className="flex gap-6 h-full items-end pb-2 ml-4">
              {["Proposed Budget", "Statement of Accounts"].map((t) => (
                <button
                  key={t}
                  onClick={() => setDocType(t)}
                  className={`text-sm font-semibold pb-2 border-b-2 transition-colors ${
                    docType === t ? "border-blue-600 text-blue-700" : "border-transparent text-slate-500 hover:text-slate-700"
                  }`}
                >
                  {t.replace(" ", "\n")}
                </button>
              ))}
            </div>
          </div>

          <div className="flex items-center gap-6">
            <div className="relative">
              <Search size={16} className="absolute left-3 top-1/2 -translate-y-1/2 text-slate-400" />
              <input 
                type="text" 
                placeholder="Search records..." 
                className="bg-slate-100 border-none rounded-md pl-9 pr-4 py-2 text-sm w-64 focus:ring-2 focus:ring-blue-100 outline-none"
              />
            </div>
            <button 
              onClick={handleDownload}
              className="bg-black hover:bg-slate-800 text-white rounded text-xs font-bold px-4 py-2.5 flex items-center gap-2 transition-colors"
            >
              <Download size={16} /> Generate<br/>Excel Report
            </button>
            <button className="text-slate-400 hover:text-slate-600"><Bell size={20}/></button>
            <div className="w-8 h-8 rounded-full bg-slate-200 border-2 border-white shadow-sm overflow-hidden flex items-center justify-center shrink-0">
               <span className="text-xs font-bold text-slate-500">VK</span>
            </div>
          </div>
        </div>

        {/* SCROLLABLE WORKSPACE */}
        <div className="flex-1 overflow-auto p-8 relative">
          <div className="max-w-6xl mx-auto space-y-6 pb-20">
            
            {/* GENERAL INFO CARD */}
            <div className="bg-white border border-slate-200 rounded-lg p-6 shadow-sm">
              <div className="grid grid-cols-1 md:grid-cols-4 gap-6">
                <div>
                  <label className={labelCls}>Network</label>
                  <input className={inputCls} value={orgName} onChange={(e) => setOrgName(e.target.value)} />
                </div>
                <div>
                  <label className={labelCls}>Event Name</label>
                  <input className={inputCls} value={eventName} onChange={(e) => setEventName(e.target.value)} />
                </div>
                <div>
                  <label className={labelCls}>Date</label>
                  <div className="relative">
                    <input type="date" className={inputCls} value={eventDate} onChange={(e) => setEventDate(e.target.value)} />
                  </div>
                </div>
                <div>
                  <label className={labelCls}>Time Slot</label>
                  <div className="relative">
                    <input className={inputCls} value={eventTime} onChange={(e) => setEventTime(e.target.value)} />
                    <Clock size={16} className="absolute right-3 top-1/2 -translate-y-1/2 text-slate-400 pointer-events-none" />
                  </div>
                </div>
              </div>
            </div>

            {/* FINANCIALS SPLIT */}
            <div className="grid grid-cols-1 xl:grid-cols-2 gap-6">
              
              {/* INCOME */}
              <div className="bg-white border border-slate-200 rounded-lg shadow-sm flex flex-col">
                <div className="bg-slate-50 border-b border-slate-200 p-5 flex justify-between items-center rounded-t-lg">
                  <div className="flex items-center gap-2">
                    <TrendingUp size={20} className="text-emerald-500" />
                    <h2 className="text-lg font-bold text-slate-800">Income Items</h2>
                  </div>
                  <div className="text-right">
                    <p className="text-[9px] font-bold uppercase tracking-widest text-slate-500 mb-0.5">Total Income</p>
                    <p className="text-2xl font-mono font-bold text-emerald-600">S$ {money(totalIncome)}</p>
                  </div>
                </div>
                
                <div className="p-5 flex-1">
                  <div className="flex text-xs font-bold text-slate-500 border-b border-slate-200 pb-2 mb-3">
                    <div className="flex-1">Description</div>
                    <div className="w-16 text-center">Price</div>
                    <div className="w-16 text-center">Qty</div>
                    <div className="w-24 text-right">Amount ($)</div>
                    <div className="w-8"></div>
                  </div>
                  
                  <div className="space-y-2">
                    {income.map((r) => (
                      <FinancialRow key={r.id} row={r} onChange={(f, v) => updateRow(setIncome, r.id, f, v)} onRemove={() => removeRow(setIncome, r.id)} />
                    ))}
                  </div>

                  <button onClick={() => addRow(setIncome)} className="mt-5 flex items-center gap-2 text-sm font-semibold text-blue-600 hover:text-blue-800 transition-colors">
                    <Plus size={16} /> Add Income Line
                  </button>
                </div>
              </div>

              {/* EXPENDITURE */}
              <div className="bg-white border border-slate-200 rounded-lg shadow-sm flex flex-col">
                <div className="bg-slate-50 border-b border-slate-200 p-5 flex justify-between items-center rounded-t-lg">
                  <div className="flex items-center gap-2">
                    <TrendingDown size={20} className="text-rose-500" />
                    <h2 className="text-lg font-bold text-slate-800">Expenditure Items</h2>
                  </div>
                  <div className="text-right">
                    <p className="text-[9px] font-bold uppercase tracking-widest text-slate-500 mb-0.5">Total Expenses</p>
                    <p className="text-2xl font-mono font-bold text-rose-600">S$ {money(totalExpenditure)}</p>
                  </div>
                </div>
                
                <div className="p-5 flex-1">
                  <div className="flex text-xs font-bold text-slate-500 border-b border-slate-200 pb-2 mb-3">
                    <div className="flex-1">Description</div>
                    <div className="w-16 text-center">Price</div>
                    <div className="w-16 text-center">Qty</div>
                    <div className="w-24 text-right">Amount ($)</div>
                    <div className="w-8"></div>
                  </div>
                  
                  <div className="space-y-2">
                    {expenditure.map((r) => (
                      <FinancialRow key={r.id} row={r} color="text-rose-600" onChange={(f, v) => updateRow(setExpenditure, r.id, f, v)} onRemove={() => removeRow(setExpenditure, r.id)} />
                    ))}
                  </div>

                  <button onClick={() => addRow(setExpenditure)} className="mt-5 flex items-center gap-2 text-sm font-semibold text-blue-600 hover:text-blue-800 transition-colors">
                    <Plus size={16} /> Add Expense Line
                  </button>
                </div>
              </div>

            </div>

            {/* BOTTOM SECTION */}
            <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
              
              {/* BALANCE SUMMARY */}
              <div className="bg-white border border-slate-200 rounded-lg p-6 shadow-sm flex flex-col justify-between">
                <div>
                  <h3 className="text-xs font-bold tracking-widest text-slate-500 uppercase mb-2">Projected Net Balance</h3>
                  <div className={`text-4xl font-mono font-bold mt-2 ${isDeficit ? 'text-rose-600' : 'text-emerald-600'}`}>
                    {isDeficit ? "-S$" : "S$"} {money(Math.abs(net))}
                  </div>
                </div>
                <p className="text-xs text-slate-500 mt-6 leading-relaxed">
                  Note: This event {isDeficit ? `requires a subsidy of $${money(Math.abs(net))} from the main RC fund.` : `yields a surplus of $${money(net)} to be retained in the fund.`}
                </p>
              </div>

              {/* SIGN OFF */}
              <div className="md:col-span-2 bg-white border border-slate-200 rounded-lg p-6 shadow-sm">
                <h3 className="text-base font-bold text-slate-800 mb-6">Preparation & Authorization</h3>
                
                <div className="grid grid-cols-2 gap-8">
                  <div className="space-y-6 border-r border-slate-100 pr-8">
                    <div>
                      <label className={labelCls}>Prepared By</label>
                      <input className="w-full bg-transparent border-b border-slate-200 py-1 text-sm font-semibold text-slate-800 outline-none focus:border-blue-500" value={preparedName} onChange={(e) => setPreparedName(e.target.value)} />
                      <input className="w-full bg-transparent py-1 text-xs text-slate-500 outline-none mt-1" value={preparedDesignation} onChange={(e) => setPreparedDesignation(e.target.value)} />
                    </div>
                    <div>
                      <label className={labelCls}>Preparation Date</label>
                      <p className="text-sm text-slate-800">February 15, 2026</p>
                    </div>
                  </div>
                  
                  <div>
                    <label className={labelCls}>Digital Signature</label>
                    <div className="mt-2 border-2 border-dashed border-slate-200 rounded flex flex-col items-center justify-center h-28 text-slate-400 hover:bg-slate-50 cursor-pointer transition-colors">
                      <PenTool size={20} className="mb-2 opacity-70" />
                      <span className="text-xs font-semibold tracking-wider">TAP TO SIGN</span>
                    </div>
                  </div>
                </div>
              </div>

            </div>

            <div className="text-center pb-8">
              <p className="text-[10px] font-bold tracking-[0.2em] text-slate-400 uppercase">
                Confidential Document • Internal Use Only • EventFin Pro Cloud V4.2.0
              </p>
            </div>
            
          </div>
        </div>

      </div>

      {/* FLOATING SAVE BUTTON */}
      <button className="absolute bottom-8 right-8 bg-[#0F52BA] hover:bg-blue-700 text-white w-14 h-14 rounded-xl shadow-lg flex items-center justify-center transition-transform hover:scale-105">
        <Save size={24} />
      </button>

    </div>
  );
}

function SidebarItem({ icon, label, active }) {
  return (
    <button className={`w-full flex items-center gap-3 px-3 py-2.5 rounded-lg text-sm font-medium transition-colors ${
      active ? "bg-[#2563EB] text-white" : "text-slate-400 hover:bg-slate-800 hover:text-slate-200"
    }`}>
      <span className={active ? "opacity-100" : "opacity-70"}>{icon}</span>
      {label}
    </button>
  );
}

function FinancialRow({ row, color = "text-slate-800", onChange, onRemove }) {
  const lineTotal = (parseFloat(row.price) || 0) * (parseFloat(row.qty) || 0);
  
  return (
    <div className="flex items-center gap-2 group">
      <input
        className="flex-1 bg-transparent border-b border-transparent hover:border-slate-200 focus:border-blue-500 py-1.5 text-sm text-slate-700 outline-none transition-colors"
        placeholder="Description"
        value={row.label}
        onChange={(e) => onChange("label", e.target.value)}
      />
      <input
        className="w-16 bg-transparent border-b border-transparent hover:border-slate-200 focus:border-blue-500 py-1.5 text-sm text-slate-700 text-center outline-none transition-colors"
        placeholder="0.00"
        value={row.price}
        onChange={(e) => onChange("price", e.target.value)}
      />
      <input
        className="w-16 bg-transparent border-b border-transparent hover:border-slate-200 focus:border-blue-500 py-1.5 text-sm text-slate-700 text-center outline-none transition-colors"
        placeholder="1"
        value={row.qty}
        onChange={(e) => onChange("qty", e.target.value)}
      />
      <div className={`w-24 text-right font-mono text-sm font-medium ${color}`}>
        {money(lineTotal)}
      </div>
      <button
        onClick={onRemove}
        className="w-8 flex justify-center text-slate-300 opacity-0 group-hover:opacity-100 hover:text-rose-500 transition-all"
      >
        <Trash2 size={14} />
      </button>
    </div>
  );
}