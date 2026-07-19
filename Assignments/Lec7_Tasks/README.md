# Lec7_Tasks — Bank vs Personal Statement Reconciliation

This task reconciles transactions between the **Bank Statement** and **Personal Statement** sheets in `Expenses.xlsx`, flagging recurring/missing/mismatched entries.

## Files

- `Expenses.xlsx` — source workbook with `Bank Statement` and `Personal Statement` sheets.
- `reconcile.js` — Node.js script that performs the reconciliation.
- `Expenses_Reconciled.xlsx` — output workbook (original sheets + new `Reconciliation` sheet).

## Commands run in this session

### Command_1

Check Ms Excel file. Plan and show me the code which checks the recurring, or missing transaction between Bank Statement and Personal Statement. Reconcile at the end


```bash
# Locate the Excel file in the working directory
find . -iname "*.xlsx" -o -iname "*.xls" -o -iname "*.csv" 2>/dev/null | grep -v node_modules
```

```powershell
# Check for a Python interpreter (not available on this machine)
Get-Command python3, py -ErrorAction SilentlyContinue
py -3 --version 2>&1
python3 --version 2>&1
```

```powershell
# Check for Node.js (available)
node --version
npm --version
```

```bash
# Check/install the xlsx npm package to read the workbook
npm ls xlsx 2>&1
npm view xlsx version 2>&1 | tail -1
npm install xlsx --no-save --silent 2>&1 | tail -20
```

```bash
# Inspect sheet names and sample rows
node -e "
const XLSX = require('xlsx');
const wb = XLSX.readFile('Expenses.xlsx');
console.log('Sheets:', wb.SheetNames);
for (const name of wb.SheetNames) {
  const ws = wb.Sheets[name];
  const data = XLSX.utils.sheet_to_json(ws, {header:1, defval:''});
  console.log('--- ' + name + ' --- rows: ' + data.length);
  data.slice(0,8).forEach(r => console.log(JSON.stringify(r)));
}
"
```

```bash
# Dump full contents of both sheets
node -e "
const XLSX = require('xlsx');
const wb = XLSX.readFile('Expenses.xlsx');
for (const name of wb.SheetNames) {
  const ws = wb.Sheets[name];
  const data = XLSX.utils.sheet_to_json(ws, {header:1, defval:''});
  console.log('=== ' + name + ' === total rows: ' + data.length);
  data.forEach((r,i) => console.log(i, JSON.stringify(r)));
}
"
```
### Command_2
reconcile.js in the folder (created the the third reconciled sheet)


```bash
# Run the reconciliation script (generates Expenses_Reconciled.xlsx)
node reconcile.js
```

```bash
# Read back and display the generated Reconciliation sheet
node -e "
const XLSX = require('xlsx');
const wb = XLSX.readFile('Expenses_Reconciled.xlsx');
const ws = wb.Sheets['Reconciliation'];
const data = XLSX.utils.sheet_to_json(ws);
console.table(data);
"
```

```powershell
# Open the reconciled report in the default application (Excel)
Start-Process "Expenses_Reconciled.xlsx"
```

## Reconciliation logic (`reconcile.js`)

1. Parse both sheets, skipping the 3 header rows.
2. Aggregate Credit/Debit **by date** per sheet (handles duplicate same-day rows).
3. Match transactions on **Date + Amount** (not Name, since the two statements use different transaction labels/categories).
4. Classify each date as one of:
   - `Reconciled`
   - `Missing in Personal Statement`
   - `Missing in Bank Statement`
   - `Amount Mismatch`
5. Write results to a new `Reconciliation` sheet in `Expenses_Reconciled.xlsx` and print a summary to the console.

### Command_3
Open the report 

## Result summary

- 13 dates reconciled
- 4 discrepancies found:
  - `2026-06-07` — 16,000 debit ("Credit Card") missing in Personal Statement
  - `2026-06-19` — Waqas 10,000/10,000 entry missing in Bank Statement (duplicate of 06-18)
  - `2026-06-20` — 75,000 credit ("Zeeshan") missing in Personal Statement
  - `2026-06-21` — 75,000 debit ("Zeeshan") missing in Personal Statement


### Command_4
Create a README.md file in the Lec7_Tasks. Add all the commands run by me and you in this session

```
# Wrote README.md to the Lec7_Tasks folder documenting this session
Write README.md
```