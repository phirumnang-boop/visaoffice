const fs = require('fs');

// 1. YearlyReceivedK2ChecklistModal.tsx
let c = fs.readFileSync('src/components/YearlyReceivedK2ChecklistModal.tsx', 'utf8');
c = c.replace(
  "      if (isTransfer) {\n        transfer[vt] = (transfer[vt] || 0) + qty;\n      } else {\n        gross[vt] = (gross[vt] || 0) + qty;\n      }",
  "      if (isTransfer) {\n        return;\n      }\n      gross[vt] = (gross[vt] || 0) + qty;"
);
fs.writeFileSync('src/components/YearlyReceivedK2ChecklistModal.tsx', c, 'utf8');

// 2. YearlyK2IssuedToTeamsReport.tsx
let r1 = fs.readFileSync('src/components/YearlyK2IssuedToTeamsReport.tsx', 'utf8');
r1 = r1.replace(
  "export const isTransferDeductionRecord = (r: StockRecord): boolean => {",
  "export const isTransferDeductionRecord = (r: StockRecord): boolean => {\n  return false;"
);
fs.writeFileSync('src/components/YearlyK2IssuedToTeamsReport.tsx', r1, 'utf8');

// 3. YearlyK2ReceivedK1Report.tsx
let r2 = fs.readFileSync('src/components/YearlyK2ReceivedK1Report.tsx', 'utf8');
r2 = r2.replace(
  "export const isTransferDeductionRecord = (r: StockRecord): boolean => {",
  "export const isTransferDeductionRecord = (r: StockRecord): boolean => {\n  return false;"
);
fs.writeFileSync('src/components/YearlyK2ReceivedK1Report.tsx', r2, 'utf8');

console.log('Successfully updated transfer deduction handling across reports.');
