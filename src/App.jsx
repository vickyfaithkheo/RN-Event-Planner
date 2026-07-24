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
    ws["!cols"] = [
      { wch: 22 }, { wch: 18 }, { wch: 12 }, { wch: 10 }, { wch: 20 }, { wch: 14 },
    ];
    
    // Format currency columns
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

  return (
    <div className="bg-background text-on-surface font-body-md overflow-hidden flex h-screen">
      
      {/* SideNavBar Shell */}
      <aside className="fixed left-0 top-0 h-screen w-sidebar-width bg-primary-container border-r border-outline-variant flex flex-col py-6 z-50 shrink-0">
        <div className="px-6 mb-10">
          <h1 className="font-headline-md text-xl font-bold text-on-primary leading-tight">Tampines GreenCourt RN</h1>
          <p className="text-on-primary-container opacity-70 text-sm mt-1">Planner Portal</p>
        </div>
        <div className="px-4 mb-8">
          <button className="w-full py-3 bg-secondary-container text-on-secondary-container font-semibold rounded shadow-sm hover:scale-95 transition-transform duration-150">
            + New Event
          </button>
        </div>
        <nav className="flex-1 space-y-1">
          <SidebarLink icon="dashboard" label="Dashboard" />
          <SidebarLink icon="account_balance_wallet" label="Budgets" />
          <a className="flex items-center px-6 py-3 text-on-secondary-container bg-secondary-container border-l-4 border-secondary rounded-r-full mr-4 transition-colors" href="#">
            <span className="material-symbols-outlined mr-3">receipt_long</span>
            <span className="font-medium">Accounts</span>
          </a>
          <SidebarLink icon="assessment" label="Reports" />
          <SidebarLink icon="settings" label="Settings" />
        </nav>
        <div className="mt-auto border-t border-outline-variant/20 pt-4">
          <SidebarLink icon="help" label="Help" />
          <SidebarLink icon="logout" label="Sign Out" />
        </div>
      </aside>

      {/* Main Content Area */}
      <div className="ml-sidebar-width flex flex-col flex-1 overflow-hidden">
        
        {/* TopAppBar */}
        <header className="flex justify-between items-center w-full px-margin-page h-20 border-b border-outline-variant bg-surface sticky top-0 z-40 shrink-0">
          <div className="flex items-center space-x-8">
            <div className="flex flex-col">
              <h2 className="text-2xl font-bold text-on-surface">Budget & Accounts</h2>
              <div className="flex items-center mt-1">
                <span className="bg-secondary-fixed text-on-secondary-fixed text-[10px] uppercase tracking-wider font-bold px-2 py-0.5 rounded">Draft</span>
              </div>
            </div>
            <nav className="hidden md:flex items-center space-x-6 h-full mt-4">
              {["Proposed Budget", "Statement of Accounts"].map((t) => (
                <button
                  key={t}
                  onClick={() => setDocType(t)}
                  className={`pb-4 font-semibold text-sm transition-all ${
                    docType === t ? "text-secondary border-b-2 border-secondary" : "text-on-surface-variant hover:text-on-surface"
                  }`}
                >
                  {t}
                </button>
              ))}
            </nav>
          </div>
          <div className="flex items-center space-x-4">
            <div className="relative">
              <span className="absolute inset-y-0 left-0 pl-3 flex items-center text-on-surface-variant">
                <span className="material-symbols-outlined">search</span>
              </span>
              <input className="bg-surface-container-low border-none rounded-lg pl-10 pr-4 py-2 text-sm focus:outline-none focus:ring-2 focus:ring-secondary/20 w-64" placeholder="Search records..." type="text" />
            </div>
            <button onClick={handleDownload} className="bg-primary text-on-primary px-4 py-2 rounded text-sm font-semibold flex items-center hover:opacity-80 transition-all cursor-pointer">
              <span className="material-symbols-outlined mr-2">download</span>
              Generate Excel Report
            </button>
            <div className="flex space-x-2 border-l border-outline-variant pl-4">
              <button className="p-2 text-on-surface-variant hover:bg-surface-container rounded-full transition-colors">
                <span className="material-symbols-outlined">notifications</span>
              </button>
              <div className="w-10 h-10 rounded-full border-2 border-outline overflow-hidden bg-white flex items-center justify-center font-bold text-slate-500">
                VK
              </div>
            </div>
          </div>
        </header>

        {/* Canvas */}
        <main className="flex-1 overflow-y-auto p-margin-page bg-background">
          <div className="max-w-container-max mx-auto space-y-stack-lg pb-10">
            
            {/* General Info Card */}
            <section className="bg-surface-container-lowest border border-outline-variant rounded-lg p-6 shadow-[0px_4px_6px_-1px_rgba(15,23,42,0.05)]">
              <div className="grid grid-cols-1 md:grid-cols-4 gap-gutter">
                <div>
                  <label className="block text-xs font-semibold text-on-surface-variant mb-2">Network</label>
                  <input className="w-full bg-surface border border-outline-variant rounded p-3 text-sm focus:outline-none focus:border-secondary transition-all" value={orgName} onChange={(e) => setOrgName(e.target.value)} />
                </div>
                <div>
                  <label className="block text-xs font-semibold text-on-surface-variant mb-2">Event Name</label>
                  <input className="w-full bg-surface border border-outline-variant rounded p-3 text-sm focus:outline-none focus:border-secondary transition-all" value={eventName} onChange={(e) => setEventName(e.target.value)} />
                </div>
                <div>
                  <label className="block text-xs font-semibold text-on-surface-variant mb-2">Date</label>
                  <div className="relative">
                    <span className="absolute inset-y-0 right-0 pr-3 flex items-center pointer-events-none text-on-surface-variant">
                      <span className="material-symbols-outlined">calendar_today</span>
                    </span>
                    <input type="date" className="w-full bg-surface border border-outline-variant rounded p-3 text-sm focus:outline-none focus:border-secondary transition-all" value={eventDate} onChange={(e) => setEventDate(e.target.value)} />
                  </div>
                </div>
                <div>
                  <label className="block text-xs font-semibold text-on-surface-variant mb-2">Time Slot</label>
                  <div className="relative">
                    <span className="absolute inset-y-0 right-0 pr-3 flex items-center pointer-events-none text-on-surface-variant">
                      <span className="material-symbols-outlined">schedule</span>
                    </span>
                    <input className="w-full bg-surface border border-outline-variant rounded p-3 text-sm focus:outline-none focus:border-secondary transition-all" value={eventTime} onChange={(e) => setEventTime(e.target.value)} />
                  </div>
                </div>
              </div>
            </section>

            {/* Financials Grid */}
            <div className="grid grid-cols-1 lg:grid-cols-2 gap-gutter">
              
              {/* Income Section */}
              <section className="bg-surface-container-lowest border border-outline-variant rounded-lg flex flex-col shadow-[0px_4px_6px_-1px_rgba(15,23,42,0.05)] overflow-hidden">
                <div className="flex justify-between items-center px-6 py-4 border-b border-outline-variant bg-surface-container-low">
                  <h3 className="text-lg font-semibold flex items-center">
                    <span className="material-symbols-outlined text-on-tertiary-container mr-2" style={{ fontVariationSettings: "'FILL' 1" }}>trending_up</span>
                    Income Items
                  </h3>
                  <div className="text-right">
                    <p className="text-[10px] font-bold text-on-surface-variant uppercase tracking-tighter">Total Income</p>
                    <p className="font-mono text-2xl font-bold text-on-tertiary-container">S$ {money(totalIncome)}</p>
                  </div>
                </div>
                <div className="p-6">
                  <table className="w-full border-collapse">
                    <thead>
                      <tr className="border-b border-outline-variant">
                        <th className="text-left py-3 text-xs font-semibold text-on-surface-variant">Description</th>
                        <th className="text-center py-3 text-xs font-semibold text-on-surface-variant w-20">Price</th>
                        <th className="text-center py-3 text-xs font-semibold text-on-surface-variant w-16">Qty</th>
                        <th className="text-right py-3 text-xs font-semibold text-on-surface-variant w-28">Amount ($)</th>
                        <th className="w-10"></th>
                      </tr>
                    </thead>
                    <tbody>
                      {income.map((r) => (
                        <FinancialRow key={r.id} row={r} color="text-on-surface" onChange={(f, v) => updateRow(setIncome, r.id, f, v)} onRemove={() => removeRow(setIncome, r.id)} />
                      ))}
                    </tbody>
                  </table>
                  <button onClick={() => addRow(setIncome)} className="mt-6 flex items-center text-secondary text-sm font-semibold hover:underline cursor-pointer">
                    <span className="material-symbols-outlined mr-1">add_circle</span> Add Income Line
                  </button>
                </div>
              </section>

              {/* Expenditure Section */}
              <section className="bg-surface-container-lowest border border-outline-variant rounded-lg flex flex-col shadow-[0px_4px_6px_-1px_rgba(15,23,42,0.05)] overflow-hidden">
                <div className="flex justify-between items-center px-6 py-4 border-b border-outline-variant bg-surface-container-low">
                  <h3 className="text-lg font-semibold flex items-center">
                    <span className="material-symbols-outlined text-error mr-2" style={{ fontVariationSettings: "'FILL' 1" }}>trending_down</span>
                    Expenditure Items
                  </h3>
                  <div className="text-right">
                    <p className="text-[10px] font-bold text-on-surface-variant uppercase tracking-tighter">Total Expenses</p>
                    <p className="font-mono text-2xl font-bold text-error">S$ {money(totalExpenditure)}</p>
                  </div>
                </div>
                <div className="p-6">
                  <table className="w-full border-collapse">
                    <thead>
                      <tr className="border-b border-outline-variant">
                        <th className="text-left py-3 text-xs font-semibold text-on-surface-variant">Description</th>
                        <th className="text-center py-3 text-xs font-semibold text-on-surface-variant w-20">Price</th>
                        <th className="text-center py-3 text-xs font-semibold text-on-surface-variant w-16">Qty</th>
                        <th className="text-right py-3 text-xs font-semibold text-on-surface-variant w-28">Amount ($)</th>
                        <th className="w-10"></th>
                      </tr>
                    </thead>
                    <tbody>
                      {expenditure.map((r) => (
                        <FinancialRow key={r.id} row={r} color="text-error" onChange={(f, v) => updateRow(setExpenditure, r.id, f, v)} onRemove={() => removeRow(setExpenditure, r.id)} />
                      ))}
                    </tbody>
                  </table>
                  <button onClick={() => addRow(setExpenditure)} className="mt-6 flex items-center text-secondary text-sm font-semibold hover:underline cursor-pointer">
                    <span className="material-symbols-outlined mr-1">add_circle</span> Add Expense Line
                  </button>
                </div>
              </section>

            </div>

            {/* Bottom Section: Net Result & Authorizations */}
            <div className="grid grid-cols-1 lg:grid-cols-3 gap-gutter">
              
              {/* Net Result Card */}
              <div className="lg:col-span-1 bg-surface-container-lowest border-l-4 border-error border border-outline-variant rounded-lg p-6 flex flex-col justify-between shadow-[0px_4px_6px_-1px_rgba(15,23,42,0.05)]">
                <div>
                  <h4 className="text-xs font-semibold text-on-surface-variant uppercase tracking-widest mb-2">Projected Net Balance</h4>
                  <p className={`text-3xl font-mono font-bold ${isDeficit ? 'text-error' : 'text-on-tertiary-container'}`}>
                    {isDeficit ? "-S$" : "S$"} {money(Math.abs(net))}
                  </p>
                </div>
                <p className="text-sm text-on-surface-variant mt-6 leading-relaxed">
                  Note: This event {isDeficit ? `requires a subsidy of $${money(Math.abs(net))} from the main RC fund.` : `yields a surplus of $${money(net)} to be retained in the fund.`}
                </p>
              </div>

            </div>

            {/* Preparation & Authorization Card (Full Width) */}
            <div className="bg-surface-container-lowest border border-outline-variant rounded-lg p-6 shadow-[0px_4px_6px_-1px_rgba(15,23,42,0.05)] mt-6">
              <h4 className="text-lg font-semibold mb-6">Preparation & Authorization</h4>
              <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
                
                <SignOffBlock 
                  title="Prepared By" 
                  name={preparedName} setName={setPreparedName}
                  designation={preparedDesignation} setDesignation={setPreparedDesignation}
                  org={preparedOrg} setOrg={setPreparedOrg}
                />
                
                <SignOffBlock 
                  title="Approved By" 
                  name={approverName} setName={setApproverName}
                  designation={approverDesignation} setDesignation={setApproverDesignation}
                  org={approverOrg} setOrg={setApproverOrg}
                />

                <SignOffBlock 
                  title="Certified Correct & True Copy by" 
                  name={directorName} setName={setDirectorName}
                  designation={directorDesignation} setDesignation={setDirectorDesignation}
                  org={directorOrg} setOrg={setDirectorOrg}
                />

              </div>
            </div>

            {/* Footer Disclaimer */}
            <footer className="text-center pt-8">
              <p className="text-on-surface-variant text-[11px] uppercase tracking-widest opacity-50">
                Confidential Document • Internal Use Only • EventFin Pro Cloud v4.2.0
              </p>
            </footer>

          </div>
        </main>

        {/* Floating Action Button */}
        <div className="fixed bottom-10 right-10 flex flex-col items-end space-y-4">
          <button className="bg-secondary text-on-secondary w-14 h-14 rounded-full shadow-lg flex items-center justify-center hover:scale-110 transition-transform cursor-pointer" title="Save Draft">
            <span className="material-symbols-outlined">save</span>
          </button>
        </div>

      </div>
    </div>
  );
}

// Sidebar Helper Component
function SidebarLink({ icon, label }) {
  return (
    <a className="flex items-center px-6 py-3 text-on-primary-container opacity-70 hover:bg-primary-fixed-variant hover:opacity-100 transition-colors" href="#">
      <span className="material-symbols-outlined mr-3">{icon}</span>
      <span className="font-medium">{label}</span>
    </a>
  );
}

// Table Row Helper Component
function FinancialRow({ row, color, onChange, onRemove }) {
  const lineTotal = (parseFloat(row.price) || 0) * (parseFloat(row.qty) || 0);
  
  return (
    <tr className="border-b border-outline-variant/50 group">
      <td className="py-3">
        <input
          className="w-full bg-transparent border-b border-transparent hover:border-outline-variant focus:outline-none focus:border-secondary py-1 text-sm transition-colors"
          placeholder="Description"
          value={row.label}
          onChange={(e) => onChange("label", e.target.value)}
        />
      </td>
      <td className="py-3 px-1">
        <input
          className="w-full bg-transparent border-b border-transparent hover:border-outline-variant focus:outline-none focus:border-secondary py-1 text-sm text-center transition-colors"
          placeholder="0.00"
          value={row.price}
          onChange={(e) => onChange("price", e.target.value)}
        />
      </td>
      <td className="py-3 px-1">
        <input
          className="w-full bg-transparent border-b border-transparent hover:border-outline-variant focus:outline-none focus:border-secondary py-1 text-sm text-center transition-colors"
          placeholder="1"
          value={row.qty}
          onChange={(e) => onChange("qty", e.target.value)}
        />
      </td>
      <td className="py-3 text-right">
        <span className={`font-mono text-sm font-medium ${color}`}>
          {money(lineTotal)}
        </span>
      </td>
      <td className="py-3 text-center">
        <button
          onClick={onRemove}
          className="text-error opacity-0 group-hover:opacity-100 hover:scale-110 transition-all cursor-pointer"
        >
          <span className="material-symbols-outlined text-sm">delete</span>
        </button>
      </td>
    </tr>
  );
}

// Signature Block Helper Component
function SignOffBlock({ title, name, setName, designation, setDesignation, org, setOrg }) {
  return (
    <div className="bg-surface-container-low/40 p-5 rounded-lg border border-outline-variant/40 flex flex-col h-full">
      <h5 className="text-xs font-bold text-on-surface-variant uppercase tracking-widest mb-4 border-b border-outline-variant/30 pb-2">
        {title}
      </h5>
      
      <div className="space-y-3 mb-6">
        <div>
          <label className="block text-[10px] font-semibold text-on-surface-variant mb-1 uppercase">Name</label>
          <input className="w-full bg-transparent border-b border-outline-variant/50 hover:border-outline-variant focus:border-secondary py-1 text-sm font-semibold outline-none transition-colors" value={name} onChange={(e) => setName(e.target.value)} placeholder="Full Name" />
        </div>
        <div>
          <label className="block text-[10px] font-semibold text-on-surface-variant mb-1 uppercase">Designation</label>
          <input className="w-full bg-transparent border-b border-outline-variant/50 hover:border-outline-variant focus:border-secondary py-1 text-xs text-on-surface-variant outline-none transition-colors" value={designation} onChange={(e) => setDesignation(e.target.value)} placeholder="Designation" />
        </div>
        <div>
          <label className="block text-[10px] font-semibold text-on-surface-variant mb-1 uppercase">Organization</label>
          <input className="w-full bg-transparent border-b border-outline-variant/50 hover:border-outline-variant focus:border-secondary py-1 text-xs text-on-surface-variant outline-none transition-colors" value={org} onChange={(e) => setOrg(e.target.value)} placeholder="Organization" />
        </div>
      </div>

      <div className="mt-auto">
        <label className="block text-[10px] font-semibold text-on-surface-variant mb-2 uppercase">Digital Signature</label>
        <div className="signature-pad w-full h-24 border border-dashed border-outline rounded flex flex-col items-center justify-center cursor-pointer hover:bg-surface transition-colors group">
          <span className="material-symbols-outlined text-on-surface-variant group-hover:scale-110 transition-transform">draw</span>
          <p className="text-on-surface-variant text-[10px] mt-1 font-bold tracking-wider">TAP TO SIGN</p>
        </div>
      </div>
    </div>
  );
}