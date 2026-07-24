import React, { useState, useMemo } from "react";
import * as XLSX from "xlsx";

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

  // Sign-off States
  const [preparedName, setPreparedName] = useState("Vicky Faith Khoo");
  const [preparedDesignation, setPreparedDesignation] = useState("Vice Chairman");
  const [preparedOrg, setPreparedOrg] = useState("Tampines GreenCourt RN");

  const [approverName, setApproverName] = useState("Effendy Chua");
  const [approverDesignation, setApproverDesignation] = useState("Chairman");
  const [approverOrg, setApproverOrg] = useState("Tampines GreenCourt RN");

  const [directorName, setDirectorName] = useState("Natalie Tan");
  const [directorDesignation, setDirectorDesignation] = useState("Constituency Director");
  const [directorOrg, setDirectorOrg] = useState("Tampines Boulevard CO");

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
    
    // INCOME TABLE
    push(["INCOME", "", "Unit Price", "Qty", "", "Total S$"]);
    const incomeStart = rows.length + 1; // 1-based row index for Excel formula start
    income.forEach((r) => {
      const p = parseFloat(r.price) || 0;
      const q = parseFloat(r.qty) || 0;
      // Formula for line amount: =C(row)*D(row)
      const currentRow = rows.length + 1;
      push([r.label || "", "", p, q, "", { t: "n", f: `C${currentRow}*D${currentRow}` }]);
    });
    const incomeEnd = rows.length; // 1-based row index for Excel formula end
    
    push([
      "TOTAL INCOME", "", "", "", "",
      { t: "n", f: `SUM(F${incomeStart}:F${incomeEnd})` },
    ]);
    const totalIncomeRow = rows.length;
    
    push([]);
    
    // EXPENDITURE TABLE
    push(["EXPENDITURE", "", "Unit Price", "Qty", "", "Total S$"]);
    const expStart = rows.length + 1;
    expenditure.forEach((r) => {
      const p = parseFloat(r.price) || 0;
      const q = parseFloat(r.qty) || 0;
      const currentRow = rows.length + 1;
      push([r.label || "", "", p, q, "", { t: "n", f: `C${currentRow}*D${currentRow}` }]);
    });
    const expEnd = rows.length;
    
    push([
      "TOTAL EXPENDITURE", "", "", "", "",
      { t: "n", f: `SUM(F${expStart}:F${expEnd})` },
    ]);
    const totalExpRow = rows.length;
    
    push([]);
    
    // SURPLUS / DEFICIT
    push([
      "SURPLUS / (DEFICIT)", "", "", "", "",
      { t: "n", f: `F${totalIncomeRow}-F${totalExpRow}` },
    ]);
    
    push([]);
    push([]);
    push(["Prepared by:", "", "", "", "Approved by:"]);
    push([]);
    push([]);
    push(["Name:", preparedName, "", "", "Name:", approverName]);
    push(["Designation:", preparedDesignation, "", "", "Designation:", approverDesignation]);
    push(["Organization:", preparedOrg, "", "", "Organization:", approverOrg]);
    push([]);
    push([]);
    push(["Certified Correct & True Copy by:"]);
    push([]);
    push([]);
    push(["Name:", directorName]);
    push(["Designation:", directorDesignation]);
    push(["Organization:", directorOrg]);

    const ws = XLSX.utils.aoa_to_sheet(rows);
    
    // Clean column widths for professional readability
    ws["!cols"] = [
      { wch: 30 }, { wch: 5 }, { wch: 14 }, { wch: 10 }, { wch: 5 }, { wch: 16 },
    ];
    
    ws["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 5 } },
    ];

    // Format currency columns (Price in Col C, Totals in Col F)
    for (let r = 0; r < rows.length; r++) {
      const addrPrice = XLSX.utils.encode_cell({ r, c: 2 });
      if (ws[addrPrice] && (typeof ws[addrPrice].v === "number" || ws[addrPrice].f)) {
        ws[addrPrice].z = "$#,##0.00;($#,##0.00)";
      }
      
      const addrTotal = XLSX.utils.encode_cell({ r, c: 5 });
      if (ws[addrTotal] && (typeof ws[addrTotal].v === "number" || ws[addrTotal].f)) {
        ws[addrTotal].z = "$#,##0.00;($#,##0.00)";
      }
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

  return (
    <div className="bg-[#f7f9fb] text-[#191c1e] font-sans overflow-hidden flex h-screen">
      
      {/* SideNavBar Shell */}
      <aside className="fixed left-0 top-0 h-screen w-[260px] bg-[#131b2e] border-r border-[#c6c6cd]/20 flex flex-col py-6 z-50 shrink-0">
        <div className="px-6 mb-10">
          <h1 className="text-xl font-bold text-white leading-tight">Tampines GreenCourt RN</h1>
          <p className="text-[#7c839b] text-sm mt-1">Planner Portal</p>
        </div>
        <div className="px-4 mb-8">
          <button className="w-full py-3 bg-[#2170e4] text-[#fefcff] font-semibold rounded shadow-sm hover:opacity-90 transition-opacity">
            + New Event
          </button>
        </div>
        <nav className="flex-1 space-y-1">
          <a className="flex items-center px-6 py-3 text-[#7c839b] hover:bg-[#3f465c]/50 hover:text-white transition-colors" href="#">
            <span className="material-symbols-outlined mr-3">dashboard</span>
            <span className="font-medium">Dashboard</span>
          </a>
          <a className="flex items-center px-6 py-3 text-[#7c839b] hover:bg-[#3f465c]/50 hover:text-white transition-colors" href="#">
            <span className="material-symbols-outlined mr-3">account_balance_wallet</span>
            <span className="font-medium">Budgets</span>
          </a>
          <a className="flex items-center px-6 py-3 text-[#fefcff] bg-[#2170e4] border-l-4 border-[#0058be] rounded-r-full mr-4 transition-colors" href="#">
            <span className="material-symbols-outlined mr-3">receipt_long</span>
            <span className="font-medium">Accounts</span>
          </a>
          <a className="flex items-center px-6 py-3 text-[#7c839b] hover:bg-[#3f465c]/50 hover:text-white transition-colors" href="#">
            <span className="material-symbols-outlined mr-3">assessment</span>
            <span className="font-medium">Reports</span>
          </a>
          <a className="flex items-center px-6 py-3 text-[#7c839b] hover:bg-[#3f465c]/50 hover:text-white transition-colors" href="#">
            <span className="material-symbols-outlined mr-3">settings</span>
            <span className="font-medium">Settings</span>
          </a>
        </nav>
        <div className="mt-auto border-t border-[#c6c6cd]/20 pt-4">
          <a className="flex items-center px-6 py-3 text-[#7c839b] hover:text-white transition-colors" href="#">
            <span className="material-symbols-outlined mr-3">help</span>
            <span className="font-medium">Help</span>
          </a>
          <a className="flex items-center px-6 py-3 text-[#7c839b] hover:text-white transition-colors" href="#">
            <span className="material-symbols-outlined mr-3">logout</span>
            <span className="font-medium">Sign Out</span>
          </a>
        </div>
      </aside>

      {/* Main Content Area */}
      <div className="ml-[260px] flex flex-col flex-1 overflow-hidden">
        
        {/* TopAppBar */}
        <header className="flex justify-between items-center w-full px-8 h-20 border-b border-[#c6c6cd] bg-[#f7f9fb] sticky top-0 z-40 shrink-0">
          <div className="flex items-center space-x-8">
            <div className="flex flex-col">
              <h2 className="text-2xl font-bold text-[#191c1e]">Budget & Accounts</h2>
              <div className="flex items-center mt-1">
                <span className="bg-[#d8e2ff] text-[#001a42] text-[10px] uppercase tracking-wider font-bold px-2 py-0.5 rounded">Draft</span>
              </div>
            </div>
            <nav className="hidden md:flex items-center space-x-6 h-full mt-4">
              {["Proposed Budget", "Statement of Accounts"].map((t) => (
                <button
                  key={t}
                  onClick={() => setDocType(t)}
                  className={`pb-4 font-semibold text-sm transition-all cursor-pointer ${
                    docType === t ? "text-[#0058be] border-b-2 border-[#0058be]" : "text-[#45464d] hover:text-[#191c1e]"
                  }`}
                >
                  {t}
                </button>
              ))}
            </nav>
          </div>
          <div className="flex items-center space-x-4">
            <div className="relative">
              <span className="absolute inset-y-0 left-0 pl-3 flex items-center text-[#45464d]">
                <span className="material-symbols-outlined">search</span>
              </span>
              <input className="bg-[#f2f4f6] border-none rounded-lg pl-10 pr-4 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-[#0058be]/20 w-64" placeholder="Search records..." type="text" />
            </div>
            <button onClick={handleDownload} className="bg-black text-white px-4 py-2 rounded text-sm font-semibold flex items-center hover:opacity-80 transition-all cursor-pointer">
              <span className="material-symbols-outlined mr-2">download</span>
              Generate Excel Report
            </button>
            <div className="flex space-x-2 border-l border-[#c6c6cd] pl-4">
              <button className="p-2 text-[#45464d] hover:bg-[#eceef0] rounded-full transition-colors cursor-pointer">
                <span className="material-symbols-outlined">notifications</span>
              </button>
              <div className="w-10 h-10 rounded-full border-2 border-[#76777d] overflow-hidden bg-white flex items-center justify-center font-bold text-[#45464d]">
                VK
              </div>
            </div>
          </div>
        </header>

        {/* Canvas */}
        <main className="flex-1 overflow-y-auto p-8 bg-[#f7f9fb]">
          <div className="max-w-[1440px] mx-auto space-y-6 pb-10">
            
            {/* General Info Card */}
            <section className="bg-white border border-[#c6c6cd] rounded-lg p-6 shadow-sm">
              <div className="grid grid-cols-1 md:grid-cols-4 gap-6">
                <div>
                  <label className="block text-xs font-semibold text-[#45464d] mb-2">Network</label>
                  <input className="w-full bg-[#f7f9fb] border border-[#c6c6cd] rounded p-3 text-sm focus:outline-none focus:border-[#0058be] transition-all" value={orgName} onChange={(e) => setOrgName(e.target.value)} />
                </div>
                <div>
                  <label className="block text-xs font-semibold text-[#45464d] mb-2">Event Name</label>
                  <input className="w-full bg-[#f7f9fb] border border-[#c6c6cd] rounded p-3 text-sm focus:outline-none focus:border-[#0058be] transition-all" value={eventName} onChange={(e) => setEventName(e.target.value)} />
                </div>
                <div>
                  <label className="block text-xs font-semibold text-[#45464d] mb-2">Date</label>
                  <div className="relative">
                    <span className="absolute inset-y-0 right-0 pr-3 flex items-center pointer-events-none text-[#45464d]">
                      <span className="material-symbols-outlined">calendar_today</span>
                    </span>
                    <input type="date" className="w-full bg-[#f7f9fb] border border-[#c6c6cd] rounded p-3 text-sm focus:outline-none focus:border-[#0058be] transition-all" value={eventDate} onChange={(e) => setEventDate(e.target.value)} />
                  </div>
                </div>
                <div>
                  <label className="block text-xs font-semibold text-[#45464d] mb-2">Time Slot</label>
                  <div className="relative">
                    <span className="absolute inset-y-0 right-0 pr-3 flex items-center pointer-events-none text-[#45464d]">
                      <span className="material-symbols-outlined">schedule</span>
                    </span>
                    <input className="w-full bg-[#f7f9fb] border border-[#c6c6cd] rounded p-3 text-sm focus:outline-none focus:border-[#0058be] transition-all" value={eventTime} onChange={(e) => setEventTime(e.target.value)} />
                  </div>
                </div>
              </div>
            </section>

            {/* Financials Grid */}
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
              
              {/* Income Section */}
              <section className="bg-white border border-[#c6c6cd] rounded-lg flex flex-col shadow-sm overflow-hidden">
                <div className="flex justify-between items-center px-6 py-4 border-b border-[#c6c6cd] bg-[#f2f4f6]">
                  <h3 className="text-lg font-semibold flex items-center text-[#191c1e]">
                    <span className="material-symbols-outlined text-[#009668] mr-2" style={{ fontVariationSettings: "'FILL' 1" }}>trending_up</span>
                    Income Items
                  </h3>
                  <div className="text-right">
                    <p className="text-[10px] font-bold text-[#45464d] uppercase tracking-tighter">Total Income</p>
                    <p className="font-mono text-2xl font-bold text-[#009668]">S$ {money(totalIncome)}</p>
                  </div>
                </div>
                <div className="p-6">
                  <table className="w-full border-collapse">
                    <thead>
                      <tr className="border-b border-[#c6c6cd]">
                        <th className="text-left py-3 text-xs font-semibold text-[#45464d]">Description</th>
                        <th className="text-center py-3 text-xs font-semibold text-[#45464d] w-20">Price</th>
                        <th className="text-center py-3 text-xs font-semibold text-[#45464d] w-16">Qty</th>
                        <th className="text-right py-3 text-xs font-semibold text-[#45464d] w-28">Amount ($)</th>
                        <th className="w-10"></th>
                      </tr>
                    </thead>
                    <tbody>
                      {income.map((r) => {
                        const lineTotal = (parseFloat(r.price) || 0) * (parseFloat(r.qty) || 0);
                        return (
                          <tr key={r.id} className="border-b border-[#c6c6cd]/50 group">
                            <td className="py-3">
                              <input
                                className="w-full bg-transparent border-b border-transparent hover:border-[#c6c6cd] focus:outline-none focus:border-[#0058be] py-1 text-sm transition-colors"
                                placeholder="Description"
                                value={r.label}
                                onChange={(e) => updateRow(setIncome, r.id, "label", e.target.value)}
                              />
                            </td>
                            <td className="py-3 px-1">
                              <input
                                className="w-full bg-transparent border-b border-transparent hover:border-[#c6c6cd] focus:outline-none focus:border-[#0058be] py-1 text-sm text-center transition-colors"
                                placeholder="0.00"
                                value={r.price}
                                onChange={(e) => updateRow(setIncome, r.id, "price", e.target.value)}
                              />
                            </td>
                            <td className="py-3 px-1">
                              <input
                                className="w-full bg-transparent border-b border-transparent hover:border-[#c6c6cd] focus:outline-none focus:border-[#0058be] py-1 text-sm text-center transition-colors"
                                placeholder="1"
                                value={r.qty}
                                onChange={(e) => updateRow(setIncome, r.id, "qty", e.target.value)}
                              />
                            </td>
                            <td className="py-3 text-right">
                              <span className="font-mono text-sm font-medium text-[#191c1e]">
                                {money(lineTotal)}
                              </span>
                            </td>
                            <td className="py-3 text-center">
                              <button
                                onClick={() => removeRow(setIncome, r.id)}
                                className="text-[#ba1a1a] opacity-0 group-hover:opacity-100 hover:scale-110 transition-all cursor-pointer"
                              >
                                <span className="material-symbols-outlined text-sm">delete</span>
                              </button>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                  <button onClick={() => addRow(setIncome)} className="mt-6 flex items-center text-[#0058be] text-sm font-semibold hover:underline cursor-pointer">
                    <span className="material-symbols-outlined mr-1">add_circle</span> Add Income Line
                  </button>
                </div>
              </section>

              {/* Expenditure Section */}
              <section className="bg-white border border-[#c6c6cd] rounded-lg flex flex-col shadow-sm overflow-hidden">
                <div className="flex justify-between items-center px-6 py-4 border-b border-[#c6c6cd] bg-[#f2f4f6]">
                  <h3 className="text-lg font-semibold flex items-center text-[#191c1e]">
                    <span className="material-symbols-outlined text-[#ba1a1a] mr-2" style={{ fontVariationSettings: "'FILL' 1" }}>trending_down</span>
                    Expenditure Items
                  </h3>
                  <div className="text-right">
                    <p className="text-[10px] font-bold text-[#45464d] uppercase tracking-tighter">Total Expenses</p>
                    <p className="font-mono text-2xl font-bold text-[#ba1a1a]">S$ {money(totalExpenditure)}</p>
                  </div>
                </div>
                <div className="p-6">
                  <table className="w-full border-collapse">
                    <thead>
                      <tr className="border-b border-[#c6c6cd]">
                        <th className="text-left py-3 text-xs font-semibold text-[#45464d]">Description</th>
                        <th className="text-center py-3 text-xs font-semibold text-[#45464d] w-20">Price</th>
                        <th className="text-center py-3 text-xs font-semibold text-[#45464d] w-16">Qty</th>
                        <th className="text-right py-3 text-xs font-semibold text-[#45464d] w-28">Amount ($)</th>
                        <th className="w-10"></th>
                      </tr>
                    </thead>
                    <tbody>
                      {expenditure.map((r) => {
                        const lineTotal = (parseFloat(r.price) || 0) * (parseFloat(r.qty) || 0);
                        return (
                          <tr key={r.id} className="border-b border-[#c6c6cd]/50 group">
                            <td className="py-3">
                              <input
                                className="w-full bg-transparent border-b border-transparent hover:border-[#c6c6cd] focus:outline-none focus:border-[#0058be] py-1 text-sm transition-colors"
                                placeholder="Description"
                                value={r.label}
                                onChange={(e) => updateRow(setExpenditure, r.id, "label", e.target.value)}
                              />
                            </td>
                            <td className="py-3 px-1">
                              <input
                                className="w-full bg-transparent border-b border-transparent hover:border-[#c6c6cd] focus:outline-none focus:border-[#0058be] py-1 text-sm text-center transition-colors"
                                placeholder="0.00"
                                value={r.price}
                                onChange={(e) => updateRow(setExpenditure, r.id, "price", e.target.value)}
                              />
                            </td>
                            <td className="py-3 px-1">
                              <input
                                className="w-full bg-transparent border-b border-transparent hover:border-[#c6c6cd] focus:outline-none focus:border-[#0058be] py-1 text-sm text-center transition-colors"
                                placeholder="1"
                                value={r.qty}
                                onChange={(e) => updateRow(setExpenditure, r.id, "qty", e.target.value)}
                              />
                            </td>
                            <td className="py-3 text-right">
                              <span className="font-mono text-sm font-medium text-[#ba1a1a]">
                                {money(lineTotal)}
                              </span>
                            </td>
                            <td className="py-3 text-center">
                              <button
                                onClick={() => removeRow(setExpenditure, r.id)}
                                className="text-[#ba1a1a] opacity-0 group-hover:opacity-100 hover:scale-110 transition-all cursor-pointer"
                              >
                                <span className="material-symbols-outlined text-sm">delete</span>
                              </button>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                  <button onClick={() => addRow(setExpenditure)} className="mt-6 flex items-center text-[#0058be] text-sm font-semibold hover:underline cursor-pointer">
                    <span className="material-symbols-outlined mr-1">add_circle</span> Add Expense Line
                  </button>
                </div>
              </section>

            </div>

            {/* Bottom Section: Net Result */}
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
              <div className={`lg:col-span-1 bg-white border-l-4 ${isDeficit ? 'border-[#ba1a1a]' : 'border-[#009668]'} border border-[#c6c6cd] rounded-lg p-6 flex flex-col justify-between shadow-sm`}>
                <div>
                  <h4 className="text-xs font-semibold text-[#45464d] uppercase tracking-widest mb-2">Projected Net Balance</h4>
                  <p className={`text-3xl font-mono font-bold ${isDeficit ? 'text-[#ba1a1a]' : 'text-[#009668]'}`}>
                    {isDeficit ? "-S$" : "S$"} {money(Math.abs(net))}
                  </p>
                </div>
                <p className="text-sm text-[#45464d] mt-6 leading-relaxed">
                  Note: This event {isDeficit ? `requires a subsidy of $${money(Math.abs(net))} from the main RC fund.` : `yields a surplus of $${money(net)} to be retained in the fund.`}
                </p>
              </div>
            </div>

            {/* Preparation & Authorization Card (Full Width) */}
            <div className="bg-white border border-[#c6c6cd] rounded-lg p-6 shadow-sm mt-6">
              <h4 className="text-lg font-semibold mb-6 text-[#191c1e]">Preparation & Authorization</h4>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                
                {/* Prepared By Block */}
                <div className="bg-[#f2f4f6]/40 p-5 rounded-lg border border-[#c6c6cd]/40 flex flex-col h-full">
                  <h5 className="text-xs font-bold text-[#45464d] uppercase tracking-widest mb-4 border-b border-[#c6c6cd]/30 pb-2">Prepared By</h5>
                  <div className="space-y-3 mb-6">
                    <div>
                      <label className="block text-[10px] font-semibold text-[#45464d] mb-1 uppercase">Name</label>
                      <input className="w-full bg-transparent border-b border-[#c6c6cd]/50 hover:border-[#c6c6cd] focus:border-[#0058be] py-1 text-sm font-semibold outline-none transition-colors" value={preparedName} onChange={(e) => setPreparedName(e.target.value)} />
                    </div>
                    <div>
                      <label className="block text-[10px] font-semibold text-[#45464d] mb-1 uppercase">Designation</label>
                      <input className="w-full bg-transparent border-b border-[#c6c6cd]/50 hover:border-[#c6c6cd] focus:border-[#0058be] py-1 text-xs text-[#45464d] outline-none transition-colors" value={preparedDesignation} onChange={(e) => setPreparedDesignation(e.target.value)} />
                    </div>
                    <div>
                      <label className="block text-[10px] font-semibold text-[#45464d] mb-1 uppercase">Organization</label>
                      <input className="w-full bg-transparent border-b border-[#c6c6cd]/50 hover:border-[#c6c6cd] focus:border-[#0058be] py-1 text-xs text-[#45464d] outline-none transition-colors" value={preparedOrg} onChange={(e) => setPreparedOrg(e.target.value)} />
                    </div>
                  </div>
                  <div className="mt-auto">
                    <label className="block text-[10px] font-semibold text-[#45464d] mb-2 uppercase">Digital Signature</label>
                    <div className="signature-pad w-full h-24 border border-dashed border-[#76777d] rounded flex flex-col items-center justify-center cursor-pointer hover:bg-[#f7f9fb] transition-colors group">
                      <span className="material-symbols-outlined text-[#45464d] group-hover:scale-110 transition-transform">draw</span>
                      <p className="text-[#45464d] text-[10px] mt-1 font-bold tracking-wider">TAP TO SIGN</p>
                    </div>
                  </div>
                </div>

                {/* Approved By Block */}
                <div className="bg-[#f2f4f6]/40 p-5 rounded-lg border border-[#c6c6cd]/40 flex flex-col h-full">
                  <h5 className="text-xs font-bold text-[#45464d] uppercase tracking-widest mb-4 border-b border-[#c6c6cd]/30 pb-2">Approved By</h5>
                  <div className="space-y-3 mb-6">
                    <div>
                      <label className="block text-[10px] font-semibold text-[#45464d] mb-1 uppercase">Name</label>
                      <input className="w-full bg-transparent border-b border-[#c6c6cd]/50 hover:border-[#c6c6cd] focus:border-[#0058be] py-1 text-sm font-semibold outline-none transition-colors" value={approverName} onChange={(e) => setApproverName(e.target.value)} />
                    </div>
                    <div>
                      <label className="block text-[10px] font-semibold text-[#45464d] mb-1 uppercase">Designation</label>
                      <input className="w-full bg-transparent border-b border-[#c6c6cd]/50 hover:border-[#c6c6cd] focus:border-[#0058be] py-1 text-xs text-[#45464d] outline-none transition-colors" value={approverDesignation} onChange={(e) => setApproverDesignation(e.target.value)} />
                    </div>
                    <div>
                      <label className="block text-[10px] font-semibold text-[#45464d] mb-1 uppercase">Organization</label>
                      <input className="w-full bg-transparent border-b border-[#c6c6cd]/50 hover:border-[#c6c6cd] focus:border-[#0058be] py-1 text-xs text-[#45464d] outline-none transition-colors" value={approverOrg} onChange={(e) => setApproverOrg(e.target.value)} />
                    </div>
                  </div>
                  <div className="mt-auto">
                    <label className="block text-[10px] font-semibold text-[#45464d] mb-2 uppercase">Digital Signature</label>
                    <div className="signature-pad w-full h-24 border border-dashed border-[#76777d] rounded flex flex-col items-center justify-center cursor-pointer hover:bg-[#f7f9fb] transition-colors group">
                      <span className="material-symbols-outlined text-[#45464d] group-hover:scale-110 transition-transform">draw</span>
                      <p className="text-[#45464d] text-[10px] mt-1 font-bold tracking-wider">TAP TO SIGN</p>
                    </div>
                  </div>
                </div>

                {/* Certified Correct Block */}
                <div className="bg-[#f2f4f6]/40 p-5 rounded-lg border border-[#c6c6cd]/40 flex flex-col h-full">
                  <h5 className="text-xs font-bold text-[#45464d] uppercase tracking-widest mb-4 border-b border-[#c6c6cd]/30 pb-2">Certified Correct & True Copy by</h5>
                  <div className="space-y-3 mb-6">
                    <div>
                      <label className="block text-[10px] font-semibold text-[#45464d] mb-1 uppercase">Name</label>
                      <input className="w-full bg-transparent border-b border-[#c6c6cd]/50 hover:border-[#c6c6cd] focus:border-[#0058be] py-1 text-sm font-semibold outline-none transition-colors" value={directorName} onChange={(e) => setDirectorName(e.target.value)} />
                    </div>
                    <div>
                      <label className="block text-[10px] font-semibold text-[#45464d] mb-1 uppercase">Designation</label>
                      <input className="w-full bg-transparent border-b border-[#c6c6cd]/50 hover:border-[#c6c6cd] focus:border-[#0058be] py-1 text-xs text-[#45464d] outline-none transition-colors" value={directorDesignation} onChange={(e) => setDirectorDesignation(e.target.value)} />
                    </div>
                    <div>
                      <label className="block text-[10px] font-semibold text-[#45464d] mb-1 uppercase">Organization</label>
                      <input className="w-full bg-transparent border-b border-[#c6c6cd]/50 hover:border-[#c6c6cd] focus:border-[#0058be] py-1 text-xs text-[#45464d] outline-none transition-colors" value={directorOrg} onChange={(e) => setDirectorOrg(e.target.value)} />
                    </div>
                  </div>
                  <div className="mt-auto">
                    <label className="block text-[10px] font-semibold text-[#45464d] mb-2 uppercase">Digital Signature</label>
                    <div className="signature-pad w-full h-24 border border-dashed border-[#76777d] rounded flex flex-col items-center justify-center cursor-pointer hover:bg-[#f7f9fb] transition-colors group">
                      <span className="material-symbols-outlined text-[#45464d] group-hover:scale-110 transition-transform">draw</span>
                      <p className="text-[#45464d] text-[10px] mt-1 font-bold tracking-wider">TAP TO SIGN</p>
                    </div>
                  </div>
                </div>

              </div>
            </div>

            {/* Footer Disclaimer */}
            <footer className="text-center pt-8">
              <p className="text-[#45464d] text-[11px] uppercase tracking-widest opacity-50">
                Confidential Document • Internal Use Only • EventFin Pro Cloud v4.2.0
              </p>
            </footer>

          </div>
        </main>

        {/* Floating Action Button */}
        <div className="fixed bottom-10 right-10 flex flex-col items-end space-y-4">
          <button className="bg-[#0058be] text-white w-14 h-14 rounded-full shadow-lg flex items-center justify-center hover:scale-110 transition-transform cursor-pointer" title="Save Draft">
            <span className="material-symbols-outlined">save</span>
          </button>
        </div>

      </div>
    </div>
  );
}