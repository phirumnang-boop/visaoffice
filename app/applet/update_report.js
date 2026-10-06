const fs = require('fs');
let str = fs.readFileSync('src/components/StickerTeamStockReport.tsx', 'utf8');

str = str.replace(
  "else if (rec.operationType === 'transferTeam') {          openMat[effectiveTeam][vt] = (openMat[effectiveTeam][vt] || 0) - qty;",
  "else if (rec.operationType === 'transferTeam') {          openMat[effectiveTeam][vt] = (openMat[effectiveTeam][vt] || 0) - qty;\n          const rawTarget = (rec as any).targetVisaTeamName || (rec as any).targetVisaTeamId || '';\n          const targetTeam = teamsList.find((t) => t === rawTarget || normalizeTeamName(t) === normalizeTeamName(rawTarget)) || rawTarget;\n          if (targetTeam) {\n            if (!openMat[targetTeam]) openMat[targetTeam] = {};\n            openMat[targetTeam][vt] = (openMat[targetTeam][vt] || 0) + qty;\n          }"
);

str = str.replace(
  "else if (rec.operationType === 'transferTeam') {          transMat[effectiveTeam][vt] = (transMat[effectiveTeam][vt] || 0) + qty;",
  "else if (rec.operationType === 'transferTeam') {          transMat[effectiveTeam][vt] = (transMat[effectiveTeam][vt] || 0) + qty;\n          const rawTarget = (rec as any).targetVisaTeamName || (rec as any).targetVisaTeamId || '';\n          const targetTeam = teamsList.find((t) => t === rawTarget || normalizeTeamName(t) === normalizeTeamName(rawTarget)) || rawTarget;\n          if (targetTeam) {\n            if (!transMat[targetTeam]) transMat[targetTeam] = {};\n            transMat[targetTeam][vt] = (transMat[targetTeam][vt] || 0) - qty;\n          }"
);

fs.writeFileSync('src/components/StickerTeamStockReport.tsx', str, 'utf8');
console.log('Successfully updated StickerTeamStockReport.tsx');
