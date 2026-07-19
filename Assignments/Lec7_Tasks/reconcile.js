const XLSX = require('xlsx');

const SRC_FILE = 'Expenses.xlsx';
const OUT_FILE = 'Expenses_Reconciled.xlsx';

const wb = XLSX.readFile(SRC_FILE);

function parseSheet(sheetName) {
  const ws = wb.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
  // Rows 0-2 are headers ("Bank/Personal Statement", "Date/Name/Amount", "Credit/Debit")
  const dataRows = rows.slice(3);

  // Aggregate credit/debit by date (handles duplicate same-day rows)
  const byDate = new Map();
  for (const row of dataRows) {
    const [date, name, credit, debit] = row;
    if (date === '' || date === undefined) continue;
    const entry = byDate.get(date) || { credit: 0, debit: 0, names: [] };
    entry.credit += Number(credit) || 0;
    entry.debit += Number(debit) || 0;
    if (name) entry.names.push(name);
    byDate.set(date, entry);
  }
  return byDate;
}

function excelDateToStr(serial) {
  // Excel serial date -> JS Date (1900 date system)
  const date = new Date(Math.round((serial - 25569) * 86400 * 1000));
  return date.toISOString().slice(0, 10);
}

const bank = parseSheet('Bank Statement');
const personal = parseSheet('Personal Statement');

// Union of all dates seen in either sheet
const allDates = new Set([...bank.keys(), ...personal.keys()]);
const sortedDates = [...allDates].sort((a, b) => a - b);

const report = [];

for (const date of sortedDates) {
  const b = bank.get(date) || { credit: 0, debit: 0, names: [] };
  const p = personal.get(date) || { credit: 0, debit: 0, names: [] };

  const creditMatch = b.credit === p.credit;
  const debitMatch = b.debit === p.debit;

  let status;
  if (b.credit === 0 && b.debit === 0 && p.credit === 0 && p.debit === 0) {
    continue; // no activity on this date in either sheet
  } else if (creditMatch && debitMatch) {
    status = 'Reconciled';
  } else if ((b.credit || b.debit) && !(p.credit || p.debit)) {
    status = 'Missing in Personal Statement';
  } else if ((p.credit || p.debit) && !(b.credit || b.debit)) {
    status = 'Missing in Bank Statement';
  } else {
    status = 'Amount Mismatch';
  }

  report.push({
    Date: excelDateToStr(date),
    'Bank Credit': b.credit || '',
    'Bank Debit': b.debit || '',
    'Bank Name(s)': b.names.join(', '),
    'Personal Credit': p.credit || '',
    'Personal Debit': p.debit || '',
    'Personal Name(s)': p.names.join(', '),
    Status: status,
  });
}

// --- Console summary ---
console.log('\n=== Reconciliation Report ===\n');
console.table(report);

const summary = report.reduce((acc, r) => {
  acc[r.Status] = (acc[r.Status] || 0) + 1;
  return acc;
}, {});
console.log('\n=== Summary ===');
console.table(summary);

const issues = report.filter(r => r.Status !== 'Reconciled');
console.log(`\n${issues.length} discrepanc${issues.length === 1 ? 'y' : 'ies'} found out of ${report.length} dates with activity.\n`);

// --- Write results as a new sheet in the workbook ---
const reportWs = XLSX.utils.json_to_sheet(report);
XLSX.utils.book_append_sheet(wb, reportWs, 'Reconciliation');
XLSX.writeFile(wb, OUT_FILE);
console.log(`Report written to ${OUT_FILE} (new "Reconciliation" sheet), original file untouched.`);
