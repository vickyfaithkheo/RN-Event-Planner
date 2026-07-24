import React, { useState, useMemo } from "react";
import * as XLSX from "xlsx";

const uid = () => Math.random().toString(36).slice(2, 9);

const money = (n) =>
  (isFinite(n) ? n : 0).toLocaleString("en-SG", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });

const ACCOUNT_DESCRIPTIONS = [
  "Administrative Charges",
  "Temporary Staff",
  "OTC Credit Card Admin Charges",
  "Passion Card Admin Fee",
  "Bank Charges",
  "Newspapers & Periodicals",
  "Office Supplies",
  "Office Inventory",
  "Fixed Assets",
  "Public Utilities",
  "Security Services",
  "Property Tax",
  "GST Expended",
  "Admission Charges",
  "Decoration",
  "Event Expenses",
  "Food and Drinks",
  "Gifts & Souvenirs",
  "Lighting & Sound System",
  "Logistics",
  "Publicity",
  "Rentals",
  "Tentage",
  "Tour Charges",
  "Transport Expenses",
  "Volunteer Expenses",
];

export default function EventBudgetGenerator() {
  const [activeNav, setActiveNav] = useState("Proposed Budget");
  const [orgName, setOrgName] = useState("Tampines Greencourt RN");
  const [eventName, setEventName] = useState("Visit to Snail Farm");
  const [eventDate, setEventDate] = useState("2026-03-01");
  const [eventTime, setEventTime] = useState("10:30AM to 12:00PM");

  // Budget / Accounts items
  const [income, setIncome] = useState([
    { id: uid(), label: "Registration Fee", price: "10", qty: "45" },
    { id: uid(), label: "Infant Fee", price: "10", qty: "5" },
  ]);
  const [expenditure, setExpenditure] = useState([
    { id: uid(), label: "Entrance Fee", price: "17.77", qty: "45" },
    { id: uid(), label: "Bus", price: "400", qty: "1" },
  ]);

  // Funding Percentage States for Statement of Accounts
  const [rnFundingPct, setRnFundingPct] = useState("50");
  const [necdcFundingPct, setNecdcFundingPct] = useState("50");

  // Claim Details state
  const [claims, setClaims] = useState([
    {
      id: uid(),
      projectName: "Annual Gala 2024",
      vendor: "EventMaster Pte Ltd/2021004A",
      invoiceNo: "INV/2024/991",
      date: "12/10/2024",
      itemDetails: "Full venue booking and banquet service for 500pax",
      accountGroup: "Activity_Expenses",
      accountDesc: "Rentals",
      amount: "12500.00",
    },
    {
      id: uid(),
      projectName: "", 
      vendor: "PrintPress Solutions/1998221B",
      invoiceNo: "PP-1042",
      date: "15/10/2024",
      itemDetails: "Design and printing of 500 event brochures",
      accountGroup: "Activity_Expenses",
      accountDesc: "Publicity",
      amount: "2450.00",
    },
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

  // New Sign-off State for Statement of Accounts
  const [gcdodName, setGcdodName] = useState("Jason Goh");
  const [gcdodDesignation, setGcdodDesignation] = useState("Deputy Constituency Director");
  const [gcdodOrg, setGcdodOrg] = useState("Tampines Boulevard CO");

  // Calculations for Budget/Accounts
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

  // Calculated Fund Allocations based on Net Balance
  const rnAllocatedAmount = net * ((parseFloat(rnFundingPct) || 0) / 100);
  const necdcAllocatedAmount = net * ((parseFloat(necdcFundingPct) || 0) / 100);

  // Calculations for Claims Group Subtotals
  const activityExpensesTotal = useMemo(
    () => claims.filter(c => c.accountGroup === "Activity_Expenses").reduce((s, r) => s + (parseFloat(r.amount) || 0), 0),
    [claims]
  );
  const adminExpensesTotal = useMemo(
    () => claims.filter(c => c.accountGroup === "Administrative_Expenses").reduce((s, r) => s + (parseFloat(r.amount) || 0), 0),
    [claims]
  );
  const claimsGrandTotal = activityExpensesTotal + adminExpensesTotal;

  const updateRow = (setter, id, field, value) =>
    setter((rows) => rows.map((r) => (r.id === id ? { ...r, [field]: value } : r)));
  const addRow = (setter) =>
    setter((rows) => [...rows, { id: uid(), label: "", price: "", qty: "" }]);
  const removeRow = (setter, id) =>
    setter((rows) => rows.filter((r) => r.id !== id));

  const addClaimRow = () =>
    setClaims((rows) => [
      ...rows,
      {
        id: uid(),
        projectName: "",
        vendor: "",
        invoiceNo: "",
        date: "",
        itemDetails: "",
        accountGroup: "Activity_Expenses",
        accountDesc: "Event Expenses",
        amount: "0.00",
      },
    ]);
  const removeClaimRow = (id) => setClaims((rows) => rows.filter((r) => r.id !== id));

  function buildWorkbook() {
    const rows = [];
    const push = (arr) => rows.push(arr);

    const activeDocTitle = activeNav;

    push([activeDocTitle]);
    push([`Event: ${eventName}`]);
    push([`Date: ${eventDate}`]);
    push([`Time: ${eventTime}`]);
    push([orgName]);
    push([]);
    
    if (activeNav === "Claim Details") {
      push(["SN", "Project Name", "Vendor/UEN/NRIC", "Invoice No", "Date", "Item Details", "Account Group", "Account Description", "Amount ($)"]);
      claims.forEach((c, index) => {
        const pName = c.projectName || (claims[0] ? claims[0].projectName : "");
        push([index + 1, pName, c.vendor, c.invoiceNo, c.date, c.itemDetails, c.accountGroup, c.accountDesc, parseFloat(c.amount) || 0]);
      });
      push([]);
      push(["Activity_Expenses Subtotal", "", "", "", "", "", "", "", activityExpensesTotal]);
      push(["Administrative_Expenses Subtotal", "", "", "", "", "", "", "", adminExpensesTotal]);
      push(["TOTAL CLAIM AMOUNT", "", "", "", "", "", "", "", claimsGrandTotal]);
    } else {
      push(["INCOME", "", "Unit Price", "Qty", "", "Total S$"]);
      const incomeStart = rows.length + 1;
      income.forEach((r) => {
        const p = parseFloat(r.price) || 0;
        const q = parseFloat(r.qty) || 0;
        const currentRow = rows.length + 1;
        push([r.label || "", "", p, q, "", { t: "n", f: `C${currentRow}*D${currentRow}` }]);
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
      
      push([
        "NET BALANCE", "", "", "", "",
        { t: "n", f: `F${totalIncomeRow}-F${totalExpRow}` },
      ]);
      
      if (activeNav === "Statement of Accounts") {
        push([]);
        push(["RN Funding Allocation (%)", rnFundingPct + "%", "", "", "NECDC Funding Allocation (%)", necdcFundingPct + "%"]);
        push(["RN Funded Amount", rnAllocatedAmount, "", "", "NECDC Funded Amount", necdcAllocatedAmount]);
      }

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
      push(["Certified Correct & True Copy by:", "", "", "", "Certified Correct & True Copy by GCDO:"]);
      push([]);
      push([]);
      push(["Name:", directorName, "", "", "Name:", gcdodName]);
      push(["Designation:", directorDesignation, "", "", "Designation:", gcdodDesignation]);
      push(["Organization:", directorOrg, "", "", "Organization:", gcdodOrg]);
    }

    const ws = XLSX.utils.aoa_to_sheet(rows);
    
    ws["!cols"] = [
      { wch: 30 }, { wch: 18 }, { wch: 22 }, { wch: 15 }, { wch: 12 }, { wch: 30 }, { wch: 22 }, { wch: 25 }, { wch: 18 },
    ];
    
    ws["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 5 } },
    ];

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, activeDocTitle);
    return wb;
  }

  function handleDownload() {
    const wb = buildWorkbook();
    const safeName = (eventName || "event").replace(/[^a-z0-9]+/gi, "_");
    const fileLabel = activeNav.replace(/[^a-z0-9]+/gi, "_");
    XLSX.writeFile(wb, `${fileLabel}_${safeName}.xlsx`);
  }

  const primaryProjectName = claims.length > 0 ? claims[0].projectName : "";

  return (
    <div className="bg-[#f7f9fb] text-[#191c1e] font-sans overflow-hidden flex h-screen">
      
      {/* SideNavBar Shell */}
      <aside className="fixed left-0 top-0 h-screen w-[260px] bg-[#131b2e] border-r border-[#c6c6cd]/20 flex flex-col py-6 z-50 shrink-0">
        <div className="px-6 mb-10">
          <h1 className="text-xl font-bold text-white leading-tight">Tampines GreenCourt RN</h1>
          <p className="text-[#7c839b] text-sm mt-1">Planner Portal</p>
        </div>
        <div className="px-4 mb-8">
          <button className="w-full py-3 bg-[#2170e4] text-[#fefcff] font-semibold rounded shadow-sm hover:opacity-90 transition-opacity cursor-pointer">
            + New Event
          </button>
        </div>
        
        <nav className="flex-1 space-y-1">
          <SidebarButton 
            icon="account_balance_wallet" 
            label="Proposed Budget" 
            active={activeNav === "Proposed Budget"} 
            onClick={() => setActiveNav("Proposed Budget")} 
          />
          <SidebarButton 
            icon="dashboard" 
            label="EDM Generation" 
            active={activeNav === "EDM Generation"} 
            onClick={() => setActiveNav("EDM Generation")} 
          />
          <SidebarButton 
            icon="receipt_long" 
            label="Claim Details" 
            active={activeNav === "Claim Details"} 
            onClick={() => setActiveNav("Claim Details")} 
          />
          <SidebarButton 
            icon="assessment" 
            label="Statement of Accounts" 
            active={activeNav === "Statement of Accounts"} 
            onClick={() => setActiveNav("Statement of Accounts")} 
          />
          <SidebarButton 
            icon="settings" 
            label="Settings" 
            active={activeNav === "Settings"} 
            onClick={() => setActiveNav("Settings")} 
          />
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
              <h2 className="text-2xl font-bold text-[#191c1e]">{activeNav}</h2>
              <div className="flex items-center mt-1">
                <span className="bg-[#d8e2ff] text-[#001a42] text-[10px] uppercase tracking-wider font-bold px-2 py-0.5 rounded">Draft</span>
              </div>
            </div>
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
            
            {activeNav === "Claim Details" ? (
              
              /* ---------------- CLAIM DETAILS TAB VIEW ---------------- */
              <div className="space-y-6">
                <div className="flex justify-between items-end">
                  <div>
                    <h3 className="text-xl font-bold text-[#191c1e] mb-1">Claim Details</h3>
                    <p className="text-[#45464d] text-sm">Review and update itemized expenditure for the current financial period.</p>
                  </div>
                  <div className="bg-[#002113]/10 px-4 py-2 border-l-4 border-[#009668] rounded">
                    <span className="text-xs font-bold text-[#009668] block">PERIOD</span>
                    <span className="text-sm font-bold">FY2024 - Q3</span>
                  </div>
                </div>

                <div className="bg-white border border-[#c6c6cd] rounded-lg shadow-sm overflow-hidden flex flex-col">
                  <div className="p-4 border-b border-[#c6c6cd] bg-[#f2f4f6] flex justify-between items-center">
                    <h4 className="font-semibold text-sm text-[#191c1e]">Itemized Expenditure Table</h4>
                    <div className="flex gap-2">
                      <button className="p-2 hover:bg-[#e0e3e5] rounded transition-colors text-[#45464d] cursor-pointer"><span className="material-symbols-outlined text-sm">filter_list</span></button>
                      <button className="p-2 hover:bg-[#e0e3e5] rounded transition-colors text-[#45464d] cursor-pointer"><span className="material-symbols-outlined text-sm">settings</span></button>
                    </div>
                  </div>

                  <div className="overflow-x-auto">
                    <table className="w-full border-collapse text-left text-sm">
                      <thead>
                        <tr className="bg-[#eceef0] text-xs font-bold text-[#45464d] border-b border-[#c6c6cd]">
                          <th className="px-3 py-3 uppercase tracking-wider w-12">SN</th>
                          <th className="px-3 py-3 uppercase tracking-wider min-w-[160px]">Project Name</th>
                          <th className="px-3 py-3 uppercase tracking-wider min-w-[180px]">Vendor/UEN/NRIC</th>
                          <th className="px-3 py-3 uppercase tracking-wider w-28">Invoice No</th>
                          <th className="px-3 py-3 uppercase tracking-wider w-28">Date</th>
                          <th className="px-3 py-3 uppercase tracking-wider min-w-[220px]">Item Details</th>
                          <th className="px-3 py-3 uppercase tracking-wider min-w-[170px]">Account Group</th>
                          <th className="px-3 py-3 uppercase tracking-wider min-w-[200px]">Account Description</th>
                          <th className="px-3 py-3 uppercase tracking-wider text-right w-36">Amount ($)</th>
                          <th className="w-10"></th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-[#c6c6cd]/30">
                        {claims.map((c, index) => {
                          const isFirst = index === 0;
                          const snStr = String(index + 1).padStart(2, "0");
                          const displayedProjectName = isFirst ? c.projectName : (c.projectName || primaryProjectName);

                          return (
                            <tr key={c.id} className="group hover:bg-[#0058be]/5 transition-colors">
                              <td className="px-3 py-3 font-mono text-[#45464d]">{snStr}</td>
                              <td className="px-3 py-3">
                                {isFirst ? (
                                  <input 
                                    className="w-full bg-transparent border-none p-1 text-sm focus:bg-white rounded font-medium" 
                                    value={c.projectName} 
                                    onChange={(e) => updateRow(setClaims, c.id, "projectName", e.target.value)} 
                                    placeholder="Project Name" 
                                  />
                                ) : (
                                  <input 
                                    className="w-full bg-transparent border-none p-1 text-sm text-[#45464d]/60 focus:bg-white rounded italic" 
                                    value={displayedProjectName} 
                                    onChange={(e) => updateRow(setClaims, c.id, "projectName", e.target.value)} 
                                    placeholder={primaryProjectName || "Project Name"} 
                                  />
                                )}
                              </td>
                              <td className="px-3 py-3">
                                <input 
                                  className="w-full bg-transparent border-none p-1 text-sm focus:bg-white rounded" 
                                  value={c.vendor} 
                                  onChange={(e) => updateRow(setClaims, c.id, "vendor", e.target.value)} 
                                  placeholder="FullName/UEN" 
                                />
                              </td>
                              <td className="px-3 py-3">
                                <input 
                                  className="w-full bg-transparent border-none p-1 text-sm focus:bg-white rounded" 
                                  value={c.invoiceNo} 
                                  onChange={(e) => updateRow(setClaims, c.id, "invoiceNo", e.target.value)} 
                                  placeholder="INV-000" 
                                />
                              </td>
                              <td className="px-3 py-3">
                                <input 
                                  className="w-full bg-transparent border-none p-1 text-sm focus:bg-white rounded" 
                                  value={c.date} 
                                  onChange={(e) => updateRow(setClaims, c.id, "date", e.target.value)} 
                                  placeholder="dd/mm/yyyy" 
                                />
                              </td>
                              <td className="px-3 py-3">
                                <textarea 
                                  rows="2" 
                                  className="w-full bg-transparent border-none p-1 text-sm focus:bg-white rounded resize-none min-h-[48px]" 
                                  value={c.itemDetails} 
                                  onChange={(e) => updateRow(setClaims, c.id, "itemDetails", e.target.value)} 
                                  placeholder="Describe items..." 
                                />
                              </td>
                              <td className="px-3 py-3">
                                <select 
                                  className="w-full bg-transparent border-none p-1 text-sm focus:bg-white rounded cursor-pointer" 
                                  value={c.accountGroup} 
                                  onChange={(e) => updateRow(setClaims, c.id, "accountGroup", e.target.value)}
                                >
                                  <option value="Activity_Expenses">Activity_Expenses</option>
                                  <option value="Administrative_Expenses">Administrative_Expenses</option>
                                </select>
                              </td>
                              <td className="px-3 py-3">
                                <select 
                                  className="w-full bg-transparent border-none p-1 text-sm focus:bg-white rounded cursor-pointer" 
                                  value={c.accountDesc} 
                                  onChange={(e) => updateRow(setClaims, c.id, "accountDesc", e.target.value)}
                                >
                                  {ACCOUNT_DESCRIPTIONS.map((desc) => (
                                    <option key={desc} value={desc}>{desc}</option>
                                  ))}
                                </select>
                              </td>
                              <td className="px-3 py-3 text-right">
                                <input 
                                  type="number" 
                                  step="0.01" 
                                  className="w-full bg-transparent border-none p-1 text-sm text-right font-mono focus:bg-white rounded" 
                                  value={c.amount} 
                                  onChange={(e) => updateRow(setClaims, c.id, "amount", e.target.value)} 
                                />
                              </td>
                              <td className="px-3 py-3 text-center">
                                <button 
                                  onClick={() => removeClaimRow(c.id)} 
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
                  </div>

                  {/* Table Footer / Group Subtotals */}
                  <div className="p-6 bg-[#f2f4f6]/50 border-t border-[#c6c6cd] flex flex-col md:flex-row md:items-center justify-between gap-6">
                    <button onClick={addClaimRow} className="flex items-center gap-2 text-[#0058be] font-semibold hover:bg-[#0058be]/10 px-4 py-2 rounded transition-all cursor-pointer">
                      <span className="material-symbols-outlined">add_circle</span> Add New Row
                    </button>
                    <div className="flex flex-col items-end">
                      <div className="flex items-center gap-8 mb-2">
                        <span className="text-xs font-bold text-[#45464d] uppercase">Activity_Expenses Subtotal</span>
                        <span className="text-sm font-mono">${money(activityExpensesTotal)}</span>
                      </div>
                      <div className="flex items-center gap-8 mb-4">
                        <span className="text-xs font-bold text-[#45464d] uppercase">Administrative_Expenses Subtotal</span>
                        <span className="text-sm font-mono">${money(adminExpensesTotal)}</span>
                      </div>
                      <div className="flex items-center gap-12 bg-[#0058be] text-white px-8 py-4 rounded-xl shadow-lg">
                        <div>
                          <span className="text-[10px] font-bold opacity-80 block uppercase tracking-widest">Total Claim Amount</span>
                          <span className="text-2xl font-bold font-mono">${money(claimsGrandTotal)}</span>
                        </div>
                        <span className="material-symbols-outlined text-3xl opacity-50">account_balance</span>
                      </div>
                    </div>
                  </div>
                </div>
              </div>

            ) : (

              /* ---------------- BUDGET & ACCOUNTS VIEWS ---------------- */
              <>
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

                {/* Bottom Section: Net Balance & Funding Allocations (Extended Box for Statement of Accounts) */}
                <div className="grid grid-cols-1 gap-6">
                  <div className={`bg-white border-l-4 ${isDeficit ? 'border-[#ba1a1a]' : 'border-[#009668]'} border border-[#c6c6cd] rounded-lg p-6 shadow-sm`}>
                    <div className="grid grid-cols-1 md:grid-cols-3 gap-6 items-center">
                      
                      {/* Net Balance Col */}
                      <div>
                        <h4 className="text-xs font-semibold text-[#45464d] uppercase tracking-widest mb-2">Net Balance</h4>
                        <p className={`text-3xl font-mono font-bold ${isDeficit ? 'text-[#ba1a1a]' : 'text-[#009668]'}`}>
                          {isDeficit ? "-S$" : "S$"} {money(Math.abs(net))}
                        </p>
                      </div>

                      {/* RN Funding Col */}
                      <div className="border-t md:border-t-0 md:border-l border-[#c6c6cd]/50 pt-4 md:pt-0 md:pl-6">
                        <div className="flex justify-between items-center mb-2">
                          <label className="text-xs font-bold text-[#45464d] uppercase tracking-wider">RN Funding</label>
                          <div className="flex items-center space-x-1">
                            <input 
                              type="number" 
                              className="w-16 bg-[#f7f9fb] border border-[#c6c6cd] rounded p-1 text-xs text-right font-mono focus:outline-none focus:border-[#0058be]" 
                              value={rnFundingPct} 
                              onChange={(e) => setRnFundingPct(e.target.value)} 
                            />
                            <span className="text-xs text-[#45464d]">%</span>
                          </div>
                        </div>
                        <p className="font-mono text-xl font-bold text-[#191c1e]">
                          {isDeficit ? "-S$" : "S$"} {money(Math.abs(rnAllocatedAmount))}
                        </p>
                      </div>

                      {/* NECDC Funding Col */}
                      <div className="border-t md:border-t-0 md:border-l border-[#c6c6cd]/50 pt-4 md:pt-0 md:pl-6">
                        <div className="flex justify-between items-center mb-2">
                          <label className="text-xs font-bold text-[#45464d] uppercase tracking-wider">NECDC Funding</label>
                          <div className="flex items-center space-x-1">
                            <input 
                              type="number" 
                              className="w-16 bg-[#f7f9fb] border border-[#c6c6cd] rounded p-1 text-xs text-right font-mono focus:outline-none focus:border-[#0058be]" 
                              value={necdcFundingPct} 
                              onChange={(e) => setNecdcFundingPct(e.target.value)} 
                            />
                            <span className="text-xs text-[#45464d]">%</span>
                          </div>
                        </div>
                        <p className="font-mono text-xl font-bold text-[#191c1e]">
                          {isDeficit ? "-S$" : "S$"} {money(Math.abs(necdcAllocatedAmount))}
                        </p>
                      </div>

                    </div>
                  </div>
                </div>

                {/* Preparation & Authorization Card (Full Width with 4 parties) */}
                <div className="bg-white border border-[#c6c6cd] rounded-lg p-6 shadow-sm mt-6">
                  <h4 className="text-lg font-semibold mb-6 text-[#191c1e]">Preparation & Authorization</h4>
                  <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
                    
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
                      <h5 className="text-xs font-bold text-[#45464d] uppercase tracking-widest mb-4 border-b border-[#c6c6cd]/30 pb-2">Certified Correct</h5>
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

                    {/* Certified Correct by GCDO Block */}
                    <div className="bg-[#f2f4f6]/40 p-5 rounded-lg border border-[#c6c6cd]/40 flex flex-col h-full">
                      <h5 className="text-xs font-bold text-[#45464d] uppercase tracking-widest mb-4 border-b border-[#c6c6cd]/30 pb-2">Certified Correct by GCDO</h5>
                      <div className="space-y-3 mb-6">
                        <div>
                          <label className="block text-[10px] font-semibold text-[#45464d] mb-1 uppercase">Name</label>
                          <input className="w-full bg-transparent border-b border-[#c6c6cd]/50 hover:border-[#c6c6cd] focus:border-[#0058be] py-1 text-sm font-semibold outline-none transition-colors" value={gcdodName} onChange={(e) => setGcdodName(e.target.value)} />
                        </div>
                        <div>
                          <label className="block text-[10px] font-semibold text-[#45464d] mb-1 uppercase">Designation</label>
                          <input className="w-full bg-transparent border-b border-[#c6c6cd]/50 hover:border-[#c6c6cd] focus:border-[#0058be] py-1 text-xs text-[#45464d] outline-none transition-colors" value={gcdodDesignation} onChange={(e) => setGcdodDesignation(e.target.value)} />
                        </div>
                        <div>
                          <label className="block text-[10px] font-semibold text-[#45464d] mb-1 uppercase">Organization</label>
                          <input className="w-full bg-transparent border-b border-[#c6c6cd]/50 hover:border-[#c6c6cd] focus:border-[#0058be] py-1 text-xs text-[#45464d] outline-none transition-colors" value={gcdodOrg} onChange={(e) => setGcdodOrg(e.target.value)} />
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
              </>
            )}

            <footer className="text-center pt-8">
              <p className="text-[#45464d] text-[11px] uppercase tracking-widest opacity-50">
                Confidential Document • Internal Use Only • EventFin Pro Cloud v4.2.0
              </p>
            </footer>

          </div>
        </main>

        <div className="fixed bottom-10 right-10 flex flex-col items-end space-y-4">
          <button className="bg-[#0058be] text-white w-14 h-14 rounded-full shadow-lg flex items-center justify-center hover:scale-110 transition-transform cursor-pointer" title="Save Draft">
            <span className="material-symbols-outlined">save</span>
          </button>
        </div>

      </div>
    </div>
  );
}

function SidebarButton({ icon, label, active, onClick }) {
  return (
    <button 
      onClick={onClick}
      className={`w-full flex items-center px-6 py-3 transition-colors cursor-pointer ${
        active 
          ? "text-[#fefcff] bg-[#2170e4] border-l-4 border-[#0058be] rounded-r-full mr-4 font-semibold" 
          : "text-[#7c839b] hover:bg-[#3f465c]/50 hover:text-white"
      }`}
    >
      <span className="material-symbols-outlined mr-3">{icon}</span>
      <span>{label}</span>
    </button>
  );
}