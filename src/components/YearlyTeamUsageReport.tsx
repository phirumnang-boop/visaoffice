import React, { useState, useMemo, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { CategoriesState, Officer, StockRecord, UserRole } from '../types';
import { INITIAL_OFFICERS } from '../data/initialData';
import { useWorkspaceSettings } from '../context/WorkspaceSettingsContext';
import {
  Printer,
  Download,
  Calendar,
  CalendarDays,
  UserCheck,
  FileSpreadsheet,
  RotateCcw,
  Save,
  ZoomIn,
  ZoomOut,
  Sparkles,
  X,
  RefreshCw,
  HardDrive,
  ListChecks,
  Search,
  CheckCircle2,
  Eye,
  EyeOff,
} from 'lucide-react';
import { exportSinglePageA4LandscapePdf } from '../utils/pdfExportHelper';
import { printA4Document } from '../utils/printHelper';
import { CustomDatePicker } from './CustomDatePicker';
import { getKhmerLunarDate, getKhmerSolarParts, parseDateInput, toKhmerNum } from '../utils/khmerCalendar';
import { normalizeDateToISO, normalizeVisaType, normalizeTeamName, formatReportTeamName } from '../utils/teamNormalization';
import { isCeaRecord, isOldStockTeamRecord } from '../utils/teamStockCalculation';
import { TacteingLine, TacteingControlSelector, saveTacteingSettings, getSavedTacteingSettings, TacteingType } from './TacteingLine';
import { YearlyMonthChecklistModal } from './YearlyMonthChecklistModal';
import { YearlyDamagedMissingChecklistModal } from './YearlyDamagedMissingChecklistModal';
import { YearlyReceivedK2ChecklistModal } from './YearlyReceivedK2ChecklistModal';
import { isIssueTeamRecord, computeTeamStockFromRobokFormula, OFFICIAL_ROBOK_TEAM_BASELINE as IMPORTED_ROBOK_TEAM_BASELINE, isRecordForTeam, isRecordForRecipientTeam, isTransferTeamRecord, isDailyRecordForTeam } from './YearlyTeamDistributionReport';

export interface YearlyTeamUsageReportProps {
  stockRecords: StockRecord[];
  categories?: CategoriesState;
  officers?: Officer[];
  currentRole?: UserRole;
  userName?: string;
  assignedTeam?: string;
  onClose?: () => void;
}

// 13 Official Visa Types
export const REPORT_VISA_TYPES = ['T', 'T1', 'T2', 'T3', 'E', 'E1', 'E2', 'E3', 'D', 'K', 'A', 'B', 'C'] as const;
export type ReportVisaType = (typeof REPORT_VISA_TYPES)[number];

// Standard Month Names in Khmer
const KHMER_MONTHS_NAMES = [
  'មករា', 'កុម្ភៈ', 'មិនា', 'មេសា', 'ឧសភា', 'មិថុនា',
  'កក្កដា', 'សីហា', 'កញ្ញា', 'តុលា', 'វិច្ឆិកា', 'ធ្នូ'
];

export interface RowUsageData {
  visaType: string;
  openingStock: number; // សន្និធិចុងគ្រា
  monthlyValues: number[]; // 12 monthly usage
  totalMonthly: number; // សរុប (12 months sum of usage)
  receivedK2: number; // ចំនួន បើកពីក២
  totalAvailable: number; // ចំនួន សរុប (សន្និធិចុងគ្រា + បើកពីក២)
  used: number; // ប្រើប្រាស់ សរុប
  damagedMissing: number; // ទិដ្ឋាការ ខ្វះ & ខូច
  returnStock: number; // បង្វិលក២
  remaining: number; // សន្និធិនៅសល់ (ចំនួនសរុប - ប្រើប្រាស់សរុប - ខ្វះ&ខូច - បង្វិលក២)
}

// Official Baseline for Team Stock from Robok Total Stock (បក.សរុបការងារស្តុក)
export const OFFICIAL_ROBOK_TEAM_BASELINE: Record<string, number> = {
  T: 93675,
  T1: 2223,
  T2: 1927,
  T3: 1682,
  E: 37448,
  E1: 2144,
  E2: 2057,
  E3: 1881,
  D: 1349,
  K: 3648,
  A: 2060,
  B: 2512,
  C: 3500,
};

// Exact official figures for Usage Report Dec 2018 - Nov 2019
export const OFFICIAL_USAGE_2018_2019_DATA: Record<string, {
  opening: number;
  months: number[];
  receivedK2: number;
  damagedMissing: number;
  returnStock: number;
}> = {
  T: {
    opening: 93675,
    months: [244292, 259936, 243333, 227630, 168261, 143744, 125855, 149391, 146947, 107579, 132535, 169057],
    receivedK2: 2130050,
    damagedMissing: 82,
    returnStock: 0,
  },
  T1: {
    opening: 2223,
    months: [569, 463, 459, 737, 618, 559, 545, 586, 546, 523, 531, 531],
    receivedK2: 7200,
    damagedMissing: 3,
    returnStock: 0,
  },
  T2: {
    opening: 1927,
    months: [22, 33, 30, 46, 30, 29, 63, 66, 18, 29, 27, 46],
    receivedK2: 1300,
    damagedMissing: 0,
    returnStock: 0,
  },
  T3: {
    opening: 1682,
    months: [49, 64, 51, 64, 59, 47, 52, 74, 54, 62, 57, 54],
    receivedK2: 1550,
    damagedMissing: 1,
    returnStock: 0,
  },
  E: {
    opening: 37448,
    months: [43205, 30104, 55963, 110011, 78702, 80737, 80725, 87321, 75115, 51822, 42626, 32917],
    receivedK2: 785450,
    damagedMissing: 38,
    returnStock: 0,
  },
  E1: {
    opening: 2144,
    months: [705, 625, 532, 814, 749, 860, 871, 959, 1026, 1039, 1035, 1013],
    receivedK2: 11200,
    damagedMissing: 4,
    returnStock: 0,
  },
  E2: {
    opening: 2057,
    months: [39, 52, 39, 58, 39, 27, 58, 49, 39, 42, 42, 56],
    receivedK2: 1100,
    damagedMissing: 4,
    returnStock: 0,
  },
  E3: {
    opening: 1881,
    months: [98, 119, 104, 185, 156, 120, 140, 142, 122, 165, 169, 149],
    receivedK2: 2850,
    damagedMissing: 2,
    returnStock: 0,
  },
  D: {
    opening: 1349,
    months: [26, 50, 541, 42, 48, 40, 25, 23, 30, 15, 72, 29],
    receivedK2: 2300,
    damagedMissing: 4,
    returnStock: 0,
  },
  K: {
    opening: 3648,
    months: [4241, 3212, 2101, 3238, 2983, 2057, 3339, 4308, 2283, 1684, 2084, 2605],
    receivedK2: 35150,
    damagedMissing: 99,
    returnStock: 0,
  },
  A: {
    opening: 2060,
    months: [162, 95, 97, 109, 104, 131, 97, 64, 67, 81, 128, 156],
    receivedK2: 1750,
    damagedMissing: 1,
    returnStock: 0,
  },
  B: {
    opening: 2512,
    months: [111, 129, 90, 130, 96, 94, 105, 101, 97, 94, 121, 137],
    receivedK2: 1700,
    damagedMissing: 3,
    returnStock: 0,
  },
  C: {
    opening: 3500,
    months: [163, 263, 115, 197, 86, 104, 200, 136, 333, 186, 114, 589],
    receivedK2: 1600,
    damagedMissing: 3,
    returnStock: 0,
  },
};

export const OFFICIAL_USAGE_2024_2025_DATA = OFFICIAL_USAGE_2018_2019_DATA;

// Helper to verify if a record is strictly a Usage by Team (ប្រតិបត្តិការ: ប្រើប្រាស់តាមក្រុម)
export const isUseTeamRecord = (r: StockRecord): boolean => {
  if (r.stockType && r.stockType !== 'sticker') return false;
  if (isCeaRecord(r)) return false;
  if (isOldStockTeamRecord(r)) return false;
  const op = (r.operationType || '').trim().toLowerCase();

  return (
    op === 'useteam' ||
    op === 'use_team' ||
    op === 'used' ||
    op.includes('ប្រើប្រាស់តាមក្រុម') ||
    op.includes('ប្រើប្រាស់')
  );
};

// Helper to calculate auto data from stock records for Yearly Team Usage
export const computeAutoTeamUsageData = (
  startYear: number,
  startMonth: number,
  stockRecords: StockRecord[] = [],
  targetTeamName?: string
): Record<string, RowUsageData> => {
  const isSpecificTeam = Boolean(targetTeamName && targetTeamName !== 'ផ្នែករដ្ឋបាល' && targetTeamName !== 'all');
  const monthCols = [];
  for (let i = 0; i < 12; i++) {
    const curTotalMonths = (startMonth - 1) + i;
    const mNum = (curTotalMonths % 12) + 1;
    const y = startYear + Math.floor(curTotalMonths / 12);
    const mName = KHMER_MONTHS_NAMES[mNum - 1];
    monthCols.push({ monthNum: mNum, year: y, name: mName });
  }

  const startIso = `${startYear}-${String(startMonth).padStart(2, '0')}-01`;
  const lastCol = monthCols[11];
  const daysInLastMonth = new Date(lastCol.year, lastCol.monthNum, 0).getDate();
  const endIso = `${lastCol.year}-${String(lastCol.monthNum).padStart(2, '0')}-${String(daysInLastMonth).padStart(2, '0')}`;

  const computedOpeningStock = computeTeamStockFromRobokFormula(startIso, stockRecords, targetTeamName);

  const result: Record<string, RowUsageData> = {};
  REPORT_VISA_TYPES.forEach((vt) => {
    result[vt] = {
      visaType: vt,
      openingStock: computedOpeningStock[vt] !== undefined ? computedOpeningStock[vt] : (isSpecificTeam ? 0 : (OFFICIAL_ROBOK_TEAM_BASELINE[vt] || 0)),
      monthlyValues: Array(12).fill(0),
      totalMonthly: 0,
      receivedK2: 0,
      totalAvailable: 0,
      used: 0,
      damagedMissing: 0,
      returnStock: 0,
      remaining: 0,
    };
  });

  const cleanRecords = (stockRecords || []).filter(
    (r) => (r.stockType === 'sticker' || !r.stockType) && !isCeaRecord(r) && !isOldStockTeamRecord(r)
  );

  cleanRecords.forEach((r) => {
    const isTransfer = isTransferTeamRecord(r);
    const isSender = !isSpecificTeam || isRecordForTeam(r, targetTeamName);
    const isRecipient = isSpecificTeam && isRecordForRecipientTeam(r, targetTeamName);

    if (isSpecificTeam) {
      if (isTransfer) {
        if (!isSender && !isRecipient) return;
      } else {
        if (!isSender) return;
      }
    }

    const dIso = normalizeDateToISO(r.date || (r as any).createdAt || '');
    if (!dIso || dIso < startIso || dIso > endIso) return;

    const vt = normalizeVisaType(r.visaType);
    if (!vt || !result[vt]) return;

    const op = (r.operationType || '').toLowerCase();
    const qty = Number((r as any).quantity || (r as any).quantityBundles || r.totalSheets || ((r as any).count ? (r as any).count * 50 : 0)) || 0;
    if (!qty) return;

    const [recY, recM] = dIso.split('-').map(Number);
    const matchIdx = monthCols.findIndex((col) => col.year === recY && col.monthNum === recM);

    // Monthly Usage (ប្រើប្រាស់តាមក្រុម)
    if (isUseTeamRecord(r) && matchIdx >= 0 && matchIdx < 12) {
      result[vt].monthlyValues[matchIdx] += qty;
    }

    const isDamagedK2 =
      op === 'damagedk2' ||
      op === 'damaged' ||
      op === 'damagedoffice' ||
      op === 'voidk2' ||
      op === 'voidoffice' ||
      op === 'missingoffice' ||
      op.includes('ខូចក២') ||
      op.includes('មិនបានការក២') ||
      op.includes('ខូចក') ||
      op.includes('ក២ខូច') ||
      op.includes('សាកក') ||
      op.includes('សាកល្បង');

    const isDamagedTeam =
      !isDamagedK2 &&
      (op === 'damagedteam' ||
        op === 'voidteam' ||
        op === 'missingteam' ||
        op === 'teamdamaged' ||
        op.includes('ខូចក្រុម') ||
        op.includes('ខ្វះក្រុម') ||
        op.includes('ខូចតាមក្រុម') ||
        op.includes('ខ្វះតាមក្រុម') ||
        ((op.includes('ខូច') || op.includes('ខ្វះ')) &&
          !op.includes('ក២') &&
          !op.includes('ក១') &&
          !op.includes('ការិយាល័យ') &&
          (r.sourceFrom?.includes('ក្រុម') || (r as any).visaTeamRobokName)));

    // ទទួលពីការិយាល័យ (បើកពីក២ / បើកផ្តល់តាមក្រុម)
    const isReceiveFromOffice =
      !isDamagedTeam &&
      !isDamagedK2 &&
      !isTransfer &&
      op !== 'useteam' &&
      op !== 'used' &&
      !op.includes('ប្រើប្រាស់') &&
      (
        isIssueTeamRecord(r) ||
        op === 'issueteam' ||
        op === 'issue_team' ||
        op === 'issue' ||
        op.includes('បើកផ្តល់') ||
        op.includes('ចែកជូន') ||
        op.includes('ផ្តល់ជូន') ||
        op.includes('ទទួលពីការិយាល័យ') ||
        op.includes('បើកពីក២') ||
        op.includes('ទទួលពីក២')
      );

    // ផ្ទេរពីក្រុម (ទទួលផ្ទេរការប្រើប្រាស់ពីក្រុមផ្សេង)
    const isTransferIn = isTransfer && isSpecificTeam && isRecipient;

    // ប្រតិបត្តិការ (ទទួលពីការិយាល័យ ឬផ្ទេរពីក្រុម) ក្នុងជួរឈរ ចំនួន បើកពីក២ (picture 2)
    if (isReceiveFromOffice || isTransferIn) {
      result[vt].receivedK2 += qty;
    }

    // ផ្ទេរ/ខ្វះ/ខូច = ផ្ទេរការប្រើប្រាស់ + ទិដ្ឋាការខ្វះក្រុម + ទិដ្ឋាការខូចក្រុម (picture 1)
    // ផ្ទេរការប្រើប្រាស់ (ផ្ទេរចេញទៅក្រុមផ្សេង)
    const isTransferOut = isTransfer && isSpecificTeam && isSender && !isRecipient;

    if (isTransferOut || isDamagedTeam) {
      result[vt].damagedMissing += qty;
    }

    // Return to Stock (បង្វិលចូលឃ្លាំង)
    const isReturnTeam =
      op === 'returnteam' ||
      op === 'returnoffice' ||
      op === 'recycle' ||
      op.includes('បង្វិល');

    if (isReturnTeam) {
      result[vt].returnStock += qty;
    }
  });

  // Daily operations for Usage (ដូចក្នុងតារាងបើកផ្តល់)
  try {
    const savedDaily = localStorage.getItem('app_daily_team_operations_v5');
    if (savedDaily) {
      const parsedDaily = JSON.parse(savedDaily);
      if (Array.isArray(parsedDaily)) {
        const existingUseKeySet = new Set<string>();
        cleanRecords.forEach((sr) => {
          if (sr.operationType === 'useTeam' && !isCeaRecord(sr)) {
            if (isSpecificTeam && !isRecordForTeam(sr, targetTeamName)) return;
            const dIso = normalizeDateToISO(sr.date || '');
            const tNorm = normalizeTeamName(sr.visaTeamRobokName || sr.sourceFrom || '');
            const vNorm = normalizeVisaType(sr.visaType);
            if (dIso && tNorm && vNorm) {
              existingUseKeySet.add(`${dIso}_${tNorm}_${vNorm}`);
            }
          }
        });

        parsedDaily.forEach((dRec: any) => {
          if (dRec.categoryType === 'cEA' || (dRec.categoryType && dRec.categoryType !== 'Sticker')) return;
          if (isSpecificTeam && !isDailyRecordForTeam(dRec, targetTeamName)) return;
          const recDate = normalizeDateToISO(dRec.date || '');
          if (!dRec.values || !recDate) return;
          if (recDate < startIso || recDate > endIso) return;

          const rawTeamName = dRec.teamName || '';
          const normTeam = normalizeTeamName(rawTeamName);
          const [recY, recM] = recDate.split('-').map(Number);
          const matchIdx = monthCols.findIndex((col) => col.year === recY && col.monthNum === recM);

          REPORT_VISA_TYPES.forEach((vt) => {
            const item = dRec.values[vt];
            const qty = parseInt(item?.quantity || '', 10) || 0;
            if (qty <= 0) return;

            const lookupKey = `${recDate}_${normTeam}_${vt}`;
            if (!existingUseKeySet.has(lookupKey)) {
              if (matchIdx >= 0 && matchIdx < 12) {
                result[vt].monthlyValues[matchIdx] += qty;
              }
            }
          });
        });
      }
    }
  } catch (e) {
    console.error('Error computing daily usage for yearly usage report:', e);
  }

  // No fallback to official sample to keep table empty/0 as requested by the user when no data is uploaded yet
  REPORT_VISA_TYPES.forEach((vt) => {

    // Exact identical formula:
    // សរុបប្រើប្រាស់ = ផលបូក ១២ខែ
    result[vt].totalMonthly = result[vt].monthlyValues.reduce((a, b) => (Number(a) || 0) + (Number(b) || 0), 0);
    // សរុបមាន = សន្និធិដើមគ្រា + ទទួលបាន
    result[vt].totalAvailable = (Number(result[vt].openingStock) || 0) + (Number(result[vt].receivedK2) || 0);
    // បានប្រើប្រាស់ = សរុបប្រើប្រាស់ ១២ខែ
    result[vt].used = result[vt].totalMonthly;
    // សន្និធិចុងគ្រា = សរុបមាន - បានប្រើប្រាស់ - ទិដ្ឋាការខ្វះ&ខូច - បង្វិលចូលឃ្លាំង
    result[vt].remaining =
      result[vt].totalAvailable -
      (Number(result[vt].used) || 0) -
      (Number(result[vt].damagedMissing) || 0) -
      (Number(result[vt].returnStock) || 0);
  });

  return result;
};

export const YearlyTeamUsageReport: React.FC<YearlyTeamUsageReportProps> = ({
  stockRecords = [],
  categories,
  officers = [],
  currentRole,
  userName,
  assignedTeam,
  onClose,
}) => {
  const { scaleMode } = useWorkspaceSettings();
  const printAreaRef = useRef<HTMLDivElement>(null);

  // Period state (Selected in dropdowns vs Applied in report)
  const [startYear, setStartYear] = useState<number>(2018);
  const [startMonth, setStartMonth] = useState<number>(12); // Default to December (ធ្នូ)
  const [appliedStartYear, setAppliedStartYear] = useState<number>(2018);
  const [appliedStartMonth, setAppliedStartMonth] = useState<number>(12);
  const [isCalculating, setIsCalculating] = useState<boolean>(false);
  const [hasCalculated, setHasCalculated] = useState<boolean>(false);

  // Check if date selection has unapplied changes
  const isFilterPending = startYear !== appliedStartYear || startMonth !== appliedStartMonth;

  // Operating Full Team Name resolution
  const defaultFullTeamName = useMemo(() => {
    // សម្រាប់គណនីការិយាល័យ (Office / Secondary): រក្សាទុក «ផ្នែករដ្ឋបាល» មិនប្តូរទេ
    if (currentRole === 'Secondary' || currentRole === 'Admin') {
      return 'ផ្នែករដ្ឋបាល';
    }
    // សម្រាប់តែការងារស្តុកក្រុម (Team side only): បង្ហាញឈ្មោះក្រុមផ្តល់ទិដ្ឋាការពេញលេញ
    if (assignedTeam && assignedTeam.trim()) {
      return formatReportTeamName(assignedTeam);
    }
    if (categories?.visaTeams && categories.visaTeams.length > 0) {
      return categories.visaTeams[0].name;
    }
    if (categories?.teams && categories.teams.length > 0) {
      return formatReportTeamName(categories.teams[0].name);
    }
    return 'ផ្នែករដ្ឋបាល';
  }, [currentRole, assignedTeam, categories]);

  // Header and Metadata State (Editable in-place)
  const [ministryName, setMinistryName] = useState<string>('ក្រសួងមហាផ្ទៃ');
  const [departmentName, setDepartmentName] = useState<string>('អគ្គនាយកដ្ឋានអន្តោប្រវេសន៍');
  const [generalDeptName, setGeneralDeptName] = useState<string>('នាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត');
  const [officeName, setOfficeName] = useState<string>('ការិយាល័យទិដ្ឋាការចូល');
  const [sectionName, setSectionName] = useState<string>(() => defaultFullTeamName);

  useEffect(() => {
    setSectionName(defaultFullTeamName);
  }, [defaultFullTeamName]);

  // Active filtering team: office accounts are exempt and view all / administrative section
  const activeFilteringTeam = useMemo(() => {
    if (currentRole === 'Secondary' || currentRole === 'Admin') {
      return undefined;
    }
    if (assignedTeam && assignedTeam.trim()) {
      return assignedTeam;
    }
    return sectionName && sectionName !== 'ផ្នែករដ្ឋបាល' ? sectionName : (defaultFullTeamName !== 'ផ្នែករដ្ឋបាល' ? defaultFullTeamName : undefined);
  }, [currentRole, assignedTeam, sectionName, defaultFullTeamName]);

  const isSecondary = currentRole === 'Secondary' || currentRole === 'User (ការិយាល័យ)' || currentRole === 'Admin';

  // Signer Details
  const [reportDate, setReportDate] = useState<string>(() => {
    const now = new Date();
    return `${now.getFullYear()}-${String(now.getMonth() + 1).padStart(2, '0')}-${String(now.getDate()).padStart(2, '0')}`;
  });
  const [signerName, setSignerName] = useState<string>(() => (isSecondary ? 'អនុសេនីយ៍ឯក ច្រេង ថុល' : ''));
  const [signerRole, setSignerRole] = useState<string>(() => (isSecondary ? 'អ្នកធ្វើតារាង' : 'ប្រធានក្រុមផ្តល់ទិដ្ឋាការ'));

  useEffect(() => {
    if (!isSecondary) {
      if (signerRole === 'អ្នកធ្វើតារាង') {
        setSignerRole('ប្រធានក្រុមផ្តល់ទិដ្ឋាការ');
      }
      if (signerName === 'អនុសេនីយ៍ឯក ច្រេង ថុល' || signerName === 'អនុសេនីយ៍ឯក អ៊ុក រ័ត្នបញ្ញា') {
        setSignerName('');
      }
    }
  }, [isSecondary]);

  // Header Font Size (pt) - Standard 12pt as requested
  const [headerFontSize, setHeaderFontSize] = useState<number>(() => {
    try {
      const saved = localStorage.getItem('yearly_team_usage_header_font_size_pt');
      if (saved !== null) {
        const val = parseFloat(saved);
        if (!isNaN(val) && val >= 8 && val <= 24) return val;
      }
    } catch {}
    return 12;
  });

  useEffect(() => {
    try {
      localStorage.setItem('yearly_team_usage_header_font_size_pt', headerFontSize.toString());
    } catch {}
  }, [headerFontSize]);

  // Signature Vertical Shift (pt) - Shifted up as requested
  const [showFormattingToolbar, setShowFormattingToolbar] = useState<boolean>(() => {
    try {
      const saved = localStorage.getItem('yearly_team_usage_show_formatting_toolbar');
      if (saved !== null) return saved === 'true';
    } catch {}
    return true;
  });

  useEffect(() => {
    try {
      localStorage.setItem('yearly_team_usage_show_formatting_toolbar', showFormattingToolbar.toString());
    } catch {}
  }, [showFormattingToolbar]);

  const [signatureShiftY, setSignatureShiftY] = useState<number>(() => {
    try {
      const saved = localStorage.getItem('yearly_team_usage_signature_shift_y');
      if (saved !== null) {
        const val = parseFloat(saved);
        if (!isNaN(val) && val >= -40 && val <= 60) return val;
      }
    } catch {}
    return 2; // Shifted up (default 2pt instead of previous ~18pt mt-6)
  });

  useEffect(() => {
    try {
      localStorage.setItem('yearly_team_usage_signature_shift_y', signatureShiftY.toString());
    } catch {}
  }, [signatureShiftY]);

  // Spacing between signer title (អ្នកធ្វើតារាង) and signer rank/name (អនុសេនីយ៍ឯក ច្រេង ថុល)
  const [signerGapHeight, setSignerGapHeight] = useState<number>(() => {
    try {
      const saved = localStorage.getItem('yearly_team_usage_signer_gap_height');
      if (saved !== null) {
        const val = parseInt(saved, 10);
        if (!isNaN(val) && val >= 20 && val <= 200) return val;
      }
    } catch {}
    return 96; // Increased default spacing (from 64px to 96px)
  });

  useEffect(() => {
    try {
      localStorage.setItem('yearly_team_usage_signer_gap_height', signerGapHeight.toString());
    } catch {}
  }, [signerGapHeight]);

  // Display Settings
  const [zoomLevel, setZoomLevel] = useState<number>(100);

  // Synchronize with global workspace scale if not fit
  useEffect(() => {
    if (scaleMode && scaleMode !== 'fit') {
      const numeric = parseInt(scaleMode.replace('%', ''), 10);
      if (!isNaN(numeric)) {
        setZoomLevel(numeric);
      }
    }
  }, [scaleMode]);
  const [fontFamily, setFontFamily] = useState<string>('Khmer OS Siemreap');
  const [fontSize, setFontSize] = useState<number>(11);
  const [tableRowHeight, setTableRowHeight] = useState<number>(() => {
    try {
      const saved = localStorage.getItem('yearly_team_usage_row_height');
      if (saved !== null) {
        const parsed = parseInt(saved, 10);
        if (!isNaN(parsed) && parsed >= 14 && parsed <= 30) return parsed;
      }
    } catch {}
    return 19;
  });
  const [lineSpacing] = useState<number>(1.25);
  const [nameColor, setNameColor] = useState<string>('#C00000'); // Red (#C00000)
  const [dateNumberFormat, setDateNumberFormat] = useState<'khmer' | 'latin'>('khmer'); // Khmer numerals vs English/Arabic numerals

  // Tacteing Customization State
  const [tacteingSettings, setTacteingSettings] = useState<{ type: TacteingType; customImage: string | null }>(() => getSavedTacteingSettings());
  const [showTacteingSelector, setShowTacteingSelector] = useState<boolean>(false);

  useEffect(() => {
    const handleTacteingUpdate = () => {
      setTacteingSettings(getSavedTacteingSettings());
    };
    window.addEventListener('tacteing_settings_updated', handleTacteingUpdate);
    window.addEventListener('storage', handleTacteingUpdate);
    return () => {
      window.removeEventListener('tacteing_settings_updated', handleTacteingUpdate);
      window.removeEventListener('storage', handleTacteingUpdate);
    };
  }, []);

  // Effective Officers list (only real system officers, exclude legacy sample mock persons)
  const effectiveOfficers = useMemo(() => {
    const dummyIds = new Set(['off-01', 'off-02', 'off-03', 'off-04', 'off-05']);
    const dummyNames = ['ចាន់ សុក្ខា', 'ចាន់ សុភា', 'ម៉េង វណ្ណា', 'វង្ស រតនា', 'រង្សី រតនា', 'កែវ បុប្ផា', 'ហេង សំណាង'];
    
    const realFromApp = (officers || []).filter(
      (o) => !dummyIds.has(o.id) && !dummyNames.some((dn) => (o.nameKh || o.name || o.nameEn || '').includes(dn))
    );
    
    const combined = [...realFromApp];
    INITIAL_OFFICERS.forEach((initOf) => {
      const initName = initOf.nameKh || initOf.name || '';
      if (!combined.some((o) => (o.nameKh || o.name || '') === initName)) {
        combined.push(initOf);
      }
    });
    return combined;
  }, [officers]);

  // Helper to extract clean officer name without rank prefix for dropdown
  const getOfficerNameOnly = (of: Officer): string => {
    const raw = of.nameKh || of.name || of.nameEn || '';
    const knownRanks = [
      'ឧត្តមសេនីយ៍ឯក', 'ឧត្តមសេនីយ៍ទោ', 'ឧត្តមសេនីយ៍ត្រី',
      'វរសេនីយ៍ឯក', 'វរសេនីយ៍ទោ', 'វរសេនីយ៍ត្រី',
      'អនុសេនីយ៍ឯក', 'អនុសេនីយ៍ទោ', 'អនុសេនីយ៍ត្រី',
      'នាយចំណង់', 'ពលបាលឯក', 'ពលបាលទោ', 'ពលបាលត្រី',
      'Major', 'Captain', 'Lieutenant Colonel', 'Lieutenant'
    ];
    let cleaned = raw;
    for (const r of knownRanks) {
      if (cleaned.startsWith(r)) {
        cleaned = cleaned.substring(r.length).trim();
        break;
      }
    }
    return cleaned || raw;
  };

  // Helper to format full rank + name for the input and signature block
  const getOfficerFormattedRankAndName = (of: Officer): string => {
    const raw = of.nameKh || of.name || of.nameEn || '';
    const knownKhmerRanks = [
      'ឧត្តមសេនីយ៍ឯក', 'ឧត្តមសេនីយ៍ទោ', 'ឧត្តមសេនីយ៍ត្រី',
      'វរសេនីយ៍ឯក', 'វរសេនីយ៍ទោ', 'វរសេនីយ៍ត្រី',
      'អនុសេនីយ៍ឯក', 'អនុសេនីយ៍ទោ', 'អនុសេនីយ៍ត្រី',
      'នាយចំណង់', 'ពលបាលឯក', 'ពលបាលទោ', 'ពលបាលត្រី'
    ];
    
    // If it already starts with a Khmer rank, use it directly
    if (knownKhmerRanks.some(kr => raw.trim().startsWith(kr))) {
      return raw.trim();
    }
    
    let rankTitle = '';
    if (of.rankId && categories?.ranks) {
      const r = categories.ranks.find(rk => rk.id === of.rankId);
      if (r) rankTitle = r.name;
    }
    if (!rankTitle && of.rank && knownKhmerRanks.includes(of.rank)) {
      rankTitle = of.rank;
    }
    if (!rankTitle) {
      rankTitle = 'អនុសេនីយ៍ឯក';
    }
    
    const nameOnly = getOfficerNameOnly(of);
    return `${rankTitle} ${nameOnly}`.trim();
  };

  // Helper to render rank in Khmer OS Siemreap and name in Khmer OS Muol Light
  const renderSignerFormatted = (fullSigner: string) => {
    if (!fullSigner) return null;

    const knownKhmerRanks = [
      'នាយឧត្តមសេនីយ៍',
      'ឧត្តមសេនីយ៍ឯក',
      'ឧត្តមសេនីយ៍ទោ',
      'ឧត្តមសេនីយ៍ត្រី',
      'វរសេនីយ៍ឯក',
      'វរៈសេនីយ៍ឯក',
      'វរេសេនីយ៍ឯក',
      'វរសេនីយ៍ទោ',
      'វរៈសេនីយ៍ទោ',
      'វរេសេនីយ៍ទោ',
      'វរសេនីយ៍ត្រី',
      'វរៈសេនីយ៍ត្រី',
      'វរេសេនីយ៍ត្រី',
      'អនុសេនីយ៍ឯក',
      'អនុសេនីយ៍ទោ',
      'អនុសេនីយ៍ត្រី',
      'ព្រិន្ទបាលឯក',
      'ព្រិន្ទបាលទោ',
      'ព្រិន្ទបាលត្រី',
      'នាយចំណង់',
      'ពលបាលឯក',
      'ពលបាលទោ',
      'ពលបាលត្រី',
      'លោក',
      'លោកស្រី',
    ];

    if (categories?.ranks) {
      categories.ranks.forEach((r) => {
        if (r.name && !knownKhmerRanks.includes(r.name.trim())) {
          knownKhmerRanks.push(r.name.trim());
        }
      });
    }

    const sortedRanks = [...knownKhmerRanks].sort((a, b) => b.length - a.length);
    const trimmed = fullSigner.trim();
    const matchedRank = sortedRanks.find((rk) => trimmed.startsWith(rk));

    if (matchedRank) {
      const namePart = trimmed.substring(matchedRank.length).trim();
      return (
        <span className="inline-flex items-baseline justify-center gap-1.5 flex-wrap">
          <span className="font-siemreap font-bold text-[12pt]" style={{ fontSize: '12pt' }}>{matchedRank}</span>
          {namePart && (
            <span
              className="font-moul font-normal text-[12pt]"
              style={{
                fontSize: '12pt',
                fontFamily: "'Khmer Mool1', 'Khmer Mool', 'Khmer OS Muol Light', 'Khmer OS Muol', 'Moul', serif"
              }}
            >
              {namePart}
            </span>
          )}
        </span>
      );
    }

    const spaceIdx = trimmed.indexOf(' ');
    if (spaceIdx > 0) {
      const firstWord = trimmed.substring(0, spaceIdx).trim();
      const rest = trimmed.substring(spaceIdx + 1).trim();
      return (
        <span className="inline-flex items-baseline justify-center gap-1.5 flex-wrap">
          <span className="font-siemreap font-bold text-[12pt]" style={{ fontSize: '12pt' }}>{firstWord}</span>
          <span
            className="font-moul font-normal text-[12pt]"
            style={{
              fontSize: '12pt',
              fontFamily: "'Khmer Mool1', 'Khmer Mool', 'Khmer OS Muol Light', 'Khmer OS Muol', 'Moul', serif"
            }}
          >
            {rest}
          </span>
        </span>
      );
    }

    return (
      <span
        className="font-moul font-normal text-[12pt]"
        style={{
          fontSize: '12pt',
          fontFamily: "'Khmer Mool1', 'Khmer Mool', 'Khmer OS Muol Light', 'Khmer OS Muol', 'Moul', serif"
        }}
      >
        {trimmed}
      </span>
    );
  };

  // Auto Calculation vs Manual Entry Mode
  const [autoCalculate, setAutoCalculate] = useState<boolean>(true);
  const [activePreset, setActivePreset] = useState<string>('custom');
  const [isSavedRecently, setIsSavedRecently] = useState<boolean>(false);
  const [isGeneratingPdf, setIsGeneratingPdf] = useState<boolean>(false);

  // Month Checklist Modal State
  const [showMonthChecklistModal, setShowMonthChecklistModal] = useState<boolean>(false);
  const [checklistInitialMonth, setChecklistInitialMonth] = useState<number>(0);
  const [showDamagedChecklistModal, setShowDamagedChecklistModal] = useState<boolean>(false);
  const [showReceivedK2ChecklistModal, setShowReceivedK2ChecklistModal] = useState<boolean>(false);

  // Main Table Data State - starts clean with 0s and dashes before user clicks calculate
  const [tableData, setTableData] = useState<Record<string, RowUsageData>>(() => {
    const initial: Record<string, RowUsageData> = {};
    REPORT_VISA_TYPES.forEach((vt) => {
      initial[vt] = {
        visaType: vt,
        openingStock: 0,
        monthlyValues: Array(12).fill(0),
        totalMonthly: 0,
        receivedK2: 0,
        totalAvailable: 0,
        used: 0,
        damagedMissing: 0,
        returnStock: 0,
        remaining: 0,
      };
    });
    return initial;
  });

  // Calculate 12 Months Column Info based on currently applied period
  const monthColumnsInfo = useMemo(() => {
    const cols = [];
    for (let i = 0; i < 12; i++) {
      const curTotalMonths = (appliedStartMonth - 1) + i;
      const mNum = (curTotalMonths % 12) + 1;
      const y = appliedStartYear + Math.floor(curTotalMonths / 12);
      const mName = KHMER_MONTHS_NAMES[mNum - 1];
      cols.push({ monthNum: mNum, year: y, name: mName });
    }
    return cols;
  }, [appliedStartYear, appliedStartMonth]);

  // Note: Data is initially uncalculated (showing '-') until user selects date and clicks "បង្ហាញទិន្នន័យ"

  // Labels for opening and ending stock periods
  const openingStockLabel = useMemo(() => {
    const prevDate = new Date(appliedStartYear, appliedStartMonth - 1, 0); // Last day of previous month
    const d = prevDate.getDate();
    const m = prevDate.getMonth() + 1;
    const y = prevDate.getFullYear();
    return `${toKhmerNum(d)} ${KHMER_MONTHS_NAMES[m - 1]} ${toKhmerNum(y)}`;
  }, [appliedStartYear, appliedStartMonth]);

  const endingStockLabel = useMemo(() => {
    const lastCol = monthColumnsInfo[11];
    const lastDate = new Date(lastCol.year, lastCol.monthNum, 0); // Last day of 12th month
    const d = lastDate.getDate();
    return `${toKhmerNum(d)} ${lastCol.name} ${toKhmerNum(lastCol.year)}`;
  }, [monthColumnsInfo]);

  // Apply Filter / Show Result Handler (Called when user clicks "បង្ហាញទិន្នន័យ")
  const handleApplyFilter = (targetYear?: number, targetMonth?: number) => {
    const y = targetYear ?? startYear;
    const m = targetMonth ?? startMonth;
    setIsCalculating(true);

    if (targetYear !== undefined) setStartYear(targetYear);
    if (targetMonth !== undefined) setStartMonth(targetMonth);
    setAppliedStartYear(y);
    setAppliedStartMonth(m);

    setTimeout(() => {
      const computed = computeAutoTeamUsageData(y, m, stockRecords, activeFilteringTeam);
      setTableData(computed);
      setIsCalculating(false);
      setHasCalculated(true);
    }, 50);
  };

  // Handle preset switching
  const handleLoadOfficial2018 = () => {
    setStartYear(2018);
    setStartMonth(12);
    setAppliedStartYear(2018);
    setAppliedStartMonth(12);
    setActivePreset('official2018');

    const initial: Record<string, RowUsageData> = {};
    REPORT_VISA_TYPES.forEach((vt) => {
      const off = OFFICIAL_USAGE_2018_2019_DATA[vt];
      const monthlyValues = off ? [...off.months] : Array(12).fill(0);
      const totalMonthly = monthlyValues.reduce((a, b) => a + b, 0);
      const opening = off ? off.opening : (OFFICIAL_ROBOK_TEAM_BASELINE[vt] || 0);
      const receivedK2 = off ? off.receivedK2 : 0;
      const totalAvailable = opening + receivedK2;
      const used = totalMonthly;
      const damagedMissing = off ? off.damagedMissing : 0;
      const returnStock = off ? off.returnStock : 0;
      const remaining = totalAvailable - used - damagedMissing - returnStock;

      initial[vt] = {
        visaType: vt,
        openingStock: opening,
        monthlyValues,
        totalMonthly,
        receivedK2,
        totalAvailable,
        used,
        damagedMissing,
        returnStock,
        remaining,
      };
    });
    setTableData(initial);
    setHasCalculated(true);
  };

  const handleLoadOfficial2024 = () => {
    setStartYear(2024);
    setStartMonth(12);
    setAppliedStartYear(2024);
    setAppliedStartMonth(12);
    setActivePreset('official2024');

    const initial: Record<string, RowUsageData> = {};
    REPORT_VISA_TYPES.forEach((vt) => {
      const off = OFFICIAL_USAGE_2024_2025_DATA[vt];
      const monthlyValues = off ? [...off.months] : Array(12).fill(0);
      const totalMonthly = monthlyValues.reduce((a, b) => a + b, 0);
      const opening = off ? off.opening : (OFFICIAL_ROBOK_TEAM_BASELINE[vt] || 0);
      const receivedK2 = off ? off.receivedK2 : 0;
      const totalAvailable = opening + receivedK2;
      const used = totalMonthly;
      const damagedMissing = off ? off.damagedMissing : 0;
      const returnStock = off ? off.returnStock : 0;
      const remaining = totalAvailable - used - damagedMissing - returnStock;

      initial[vt] = {
        visaType: vt,
        openingStock: opening,
        monthlyValues,
        totalMonthly,
        receivedK2,
        totalAvailable,
        used,
        damagedMissing,
        returnStock,
        remaining,
      };
    });
    setTableData(initial);
    setHasCalculated(true);
  };

  const handleLoadFromRobokTotalStock = () => {
    const targetStartIso = `${appliedStartYear}-${String(appliedStartMonth).padStart(2, '0')}-01`;
    const computedStock = computeTeamStockFromRobokFormula(targetStartIso, stockRecords, activeFilteringTeam);

    setTableData((prev) => {
      const updated = { ...prev };
      REPORT_VISA_TYPES.forEach((vt) => {
        const opening = computedStock[vt] !== undefined ? computedStock[vt] : (activeFilteringTeam ? 0 : (OFFICIAL_ROBOK_TEAM_BASELINE[vt] || 0));
        const cur = updated[vt] || {
          visaType: vt,
          openingStock: opening,
          monthlyValues: Array(12).fill(0),
          totalMonthly: 0,
          receivedK2: 0,
          totalAvailable: 0,
          used: 0,
          damagedMissing: 0,
          returnStock: 0,
          remaining: 0,
        };
        const totalMonthly = cur.monthlyValues.reduce((a, b) => a + b, 0);
        const totalAvailable = opening + (cur.receivedK2 || 0);
        const used = totalMonthly;
        const remaining = totalAvailable - used - (cur.damagedMissing || 0) - (cur.returnStock || 0);

        updated[vt] = {
          ...cur,
          openingStock: opening,
          totalMonthly,
          totalAvailable,
          used,
          remaining,
        };
      });
      return updated;
    });
    setActivePreset('robokTotalStock');
  };

  const handleLoadFromRealStock = () => {
    const computed = computeAutoTeamUsageData(appliedStartYear, appliedStartMonth, stockRecords, activeFilteringTeam);
    setTableData(computed);
    setActivePreset('autoRealStock');
  };

  const handleSelectPeriodPreset = (year: number, month: number, presetKey: string) => {
    setStartYear(year);
    setStartMonth(month);
    setAppliedStartYear(year);
    setAppliedStartMonth(month);
    setActivePreset(presetKey);
    const computed = computeAutoTeamUsageData(year, month, stockRecords, activeFilteringTeam);
    setTableData(computed);
    setHasCalculated(true);
  };

  // Cell Change Handler for in-place editing
  const handleCellChange = (
    visaType: string,
    field: 'openingStock' | 'month' | 'receivedK2' | 'totalMonthly' | 'totalAvailable' | 'used' | 'damagedMissing' | 'returnStock' | 'remaining',
    rawVal: string,
    monthIdx?: number
  ) => {
    const sanitized = rawVal.replace(/[^0-9.-]/g, '');
    const numVal = parseInt(sanitized, 10) || 0;

    setTableData((prev) => {
      const row = { ...(prev[visaType] || {
        visaType,
        openingStock: 0,
        monthlyValues: Array(12).fill(0),
        totalMonthly: 0,
        receivedK2: 0,
        totalAvailable: 0,
        used: 0,
        damagedMissing: 0,
        returnStock: 0,
        remaining: 0,
      }) };

      if (field === 'openingStock') {
        row.openingStock = numVal;
      } else if (field === 'month' && monthIdx !== undefined && monthIdx >= 0 && monthIdx < 12) {
        const newMonths = [...row.monthlyValues];
        newMonths[monthIdx] = numVal;
        row.monthlyValues = newMonths;
      } else if (field === 'receivedK2') {
        row.receivedK2 = numVal;
      } else if (field === 'damagedMissing') {
        row.damagedMissing = numVal;
      } else if (field === 'returnStock') {
        row.returnStock = numVal;
      } else if (field === 'totalMonthly') {
        row.totalMonthly = numVal;
      } else if (field === 'totalAvailable') {
        row.totalAvailable = numVal;
      } else if (field === 'used') {
        row.used = numVal;
      } else if (field === 'remaining') {
        row.remaining = numVal;
      }

      if (autoCalculate) {
        row.totalMonthly = row.monthlyValues.reduce((a, b) => (Number(a) || 0) + (Number(b) || 0), 0);
        row.totalAvailable = (Number(row.openingStock) || 0) + (Number(row.receivedK2) || 0);
        row.used = row.totalMonthly;
        row.remaining =
          row.totalAvailable -
          (Number(row.used) || 0) -
          (Number(row.damagedMissing) || 0) -
          (Number(row.returnStock) || 0);
      }

      return { ...prev, [visaType]: row };
    });
  };

  // Apply actual transaction values for a specific month from the Checklist Modal
  const handleApplyMonthActualValues = (monthIdx: number, visaValues: Record<string, number>) => {
    setTableData((prev) => {
      const updated = { ...prev };
      REPORT_VISA_TYPES.forEach((vt) => {
        if (updated[vt]) {
          const nextMonthly = [...updated[vt].monthlyValues];
          nextMonthly[monthIdx] = visaValues[vt] || 0;
          const totalMonthly = nextMonthly.reduce((s, v) => s + v, 0);
          const totalAvailable = (Number(updated[vt].openingStock) || 0) + (Number(updated[vt].receivedK2) || 0);
          const used = totalMonthly;
          const remaining = totalAvailable - used - (Number(updated[vt].damagedMissing) || 0) - (Number(updated[vt].returnStock) || 0);
          updated[vt] = {
            ...updated[vt],
            monthlyValues: nextMonthly,
            totalMonthly,
            used,
            totalAvailable,
            remaining,
          };
        }
      });
      return updated;
    });
  };

  // Recalculate All
  const handleRecalculateAll = () => {
    setTableData((prev) => {
      const updated: Record<string, RowUsageData> = {};
      REPORT_VISA_TYPES.forEach((vt) => {
        const row = { ...(prev[vt] || {
          visaType: vt,
          openingStock: 0,
          monthlyValues: Array(12).fill(0),
          totalMonthly: 0,
          receivedK2: 0,
          totalAvailable: 0,
          used: 0,
          damagedMissing: 0,
          returnStock: 0,
          remaining: 0,
        }) };
        row.totalMonthly = row.monthlyValues.reduce((a, b) => (Number(a) || 0) + (Number(b) || 0), 0);
        row.totalAvailable = (Number(row.openingStock) || 0) + (Number(row.receivedK2) || 0);
        row.used = row.totalMonthly;
        row.remaining =
          row.totalAvailable -
          (Number(row.used) || 0) -
          (Number(row.damagedMissing) || 0) -
          (Number(row.returnStock) || 0);
        updated[vt] = row;
      });
      return updated;
    });
  };

  // Reset / Clear All
  const handleClearAll = () => {
    if (!window.confirm('តើអ្នកប្រាកដជាចង់កំណត់ទិន្នន័យជា ០ ទាំងអស់មែនទេ?')) return;
    const cleared: Record<string, RowUsageData> = {};
    REPORT_VISA_TYPES.forEach((vt) => {
      cleared[vt] = {
        visaType: vt,
        openingStock: 0,
        monthlyValues: Array(12).fill(0),
        totalMonthly: 0,
        receivedK2: 0,
        totalAvailable: 0,
        used: 0,
        damagedMissing: 0,
        returnStock: 0,
        remaining: 0,
      };
    });
    setTableData(cleared);
  };

  // Calculate Grand Totals across all 13 types
  const grandTotals = useMemo(() => {
    let opening = 0;
    const months = Array(12).fill(0);
    let monthlySum = 0;
    let receivedK2 = 0;
    let available = 0;
    let used = 0;
    let damagedMissing = 0;
    let returnStock = 0;
    let remaining = 0;

    REPORT_VISA_TYPES.forEach((vt) => {
      const row = tableData[vt];
      if (!row) return;
      opening += Number(row.openingStock) || 0;
      row.monthlyValues.forEach((val, idx) => {
        months[idx] += Number(val) || 0;
      });
      const rowMonthlySum = row.monthlyValues.reduce((a, b) => (Number(a) || 0) + (Number(b) || 0), 0);
      monthlySum += autoCalculate ? rowMonthlySum : (Number(row.totalMonthly) || rowMonthlySum);
      receivedK2 += Number(row.receivedK2) || 0;
      const rowAvailable = (Number(row.openingStock) || 0) + (Number(row.receivedK2) || 0);
      available += autoCalculate ? rowAvailable : (Number(row.totalAvailable) || rowAvailable);
      used += autoCalculate ? rowMonthlySum : (Number(row.used) || rowMonthlySum);
      damagedMissing += Number(row.damagedMissing) || 0;
      returnStock += Number(row.returnStock) || 0;
      const rowRemaining = rowAvailable - (Number(row.used) || rowMonthlySum) - (Number(row.damagedMissing) || 0) - (Number(row.returnStock) || 0);
      remaining += autoCalculate ? rowRemaining : (Number(row.remaining) || rowRemaining);
    });

    return {
      opening,
      months,
      monthlySum,
      receivedK2,
      available,
      used,
      damagedMissing,
      returnStock,
      remaining,
    };
  }, [tableData, autoCalculate]);

  // Khmer Dates for Footer
  const signatureSolarParts = useMemo(() => {
    const d = parseDateInput(reportDate);
    return getKhmerSolarParts(d);
  }, [reportDate]);

  const signatureLunarDate = useMemo(() => {
    const d = parseDateInput(reportDate);
    return getKhmerLunarDate(d);
  }, [reportDate]);

  // Load Saved Draft on mount
  useEffect(() => {
    try {
      const savedRaw = localStorage.getItem('yearly_team_usage_report_draft');
      if (savedRaw) {
        const saved = JSON.parse(savedRaw);
        if (saved && saved.tableData) {
          setTableData(saved.tableData);
          if (saved.startYear) setStartYear(saved.startYear);
          if (saved.startMonth) setStartMonth(saved.startMonth);
          if (saved.signerName) {
            if (!isSecondary && (saved.signerName === 'អនុសេនីយ៍ឯក ច្រេង ថុល' || saved.signerName === 'អនុសេនីយ៍ឯក អ៊ុក រ័ត្នបញ្ញា')) {
              setSignerName('');
            } else {
              setSignerName(saved.signerName);
            }
          }
          if (saved.signerRole) {
            if (!isSecondary && saved.signerRole === 'អ្នកធ្វើតារាង') {
              setSignerRole('ប្រធានក្រុមផ្តល់ទិដ្ឋាការ');
            } else {
              setSignerRole(saved.signerRole);
            }
          }
          if (saved.reportDate) setReportDate(saved.reportDate);
          if (saved.dateNumberFormat) setDateNumberFormat(saved.dateNumberFormat);
        }
      }
    } catch (err) {
      console.error('Error loading draft:', err);
    }
  }, []);

  // Save Draft
  const handleSaveDraft = () => {
    try {
      localStorage.setItem(
        'yearly_team_usage_report_draft',
        JSON.stringify({
          tableData,
          startYear,
          startMonth,
          signerName,
          signerRole,
          reportDate,
          dateNumberFormat,
          updatedAt: new Date().toISOString(),
        })
      );
      setIsSavedRecently(true);
      setTimeout(() => setIsSavedRecently(false), 3000);
    } catch (err) {
      console.error('Error saving draft:', err);
    }
  };

  // Export to Excel
  const handleExportExcel = () => {
    const wb = XLSX.utils.book_new();

    const headers1 = [
      'ប្រភេទ',
      `សន្និធិចុងគ្រា (${openingStockLabel})`,
      ...monthColumnsInfo.map((m) => m.name),
      'សរុប (១២ខែ)',
      'ចំនួន បើកពីក២',
      'ចំនួន សរុប',
      'ប្រើប្រាស់ សរុប',
      !isSecondary ? 'ផ្ទេរ/ខ្វះ/ខូច' : 'ទិដ្ឋាការ ខ្វះ&ខូច',
      'បង្វិលក២',
      `សន្និធិនៅសល់ (${endingStockLabel})`,
    ];

    const dataRows: any[][] = [];

    REPORT_VISA_TYPES.forEach((vt) => {
      const row = tableData[vt];
      if (!row) return;
      const dynamicMonthlySum = row.monthlyValues.reduce((a, b) => (Number(a) || 0) + (Number(b) || 0), 0);
      const displayAvailable = (Number(row.openingStock) || 0) + (Number(row.receivedK2) || 0);
      const displayRemaining = displayAvailable - (Number(row.used) || dynamicMonthlySum) - (Number(row.damagedMissing) || 0) - (Number(row.returnStock) || 0);

      dataRows.push([
        vt,
        row.openingStock || 0,
        ...row.monthlyValues,
        dynamicMonthlySum,
        row.receivedK2 || 0,
        displayAvailable,
        row.used || dynamicMonthlySum,
        row.damagedMissing || 0,
        row.returnStock || 0,
        displayRemaining,
      ]);
    });

    dataRows.push([
      'សរុប',
      grandTotals.opening,
      ...grandTotals.months,
      grandTotals.monthlySum,
      grandTotals.receivedK2,
      grandTotals.available,
      grandTotals.used,
      grandTotals.damagedMissing,
      grandTotals.returnStock,
      grandTotals.remaining,
    ]);

    const titleRows = [
      ['តារាងសន្លឹក ទិដ្ឋាការស្អិត តាមប្រភេទនីមួយៗ បានផ្តល់ជូនភ្ញៀវ'],
      [`ដោយគិតចាប់ពីថ្ងៃទី០១ ខែ${monthColumnsInfo[0].name} ឆ្នាំ${toKhmerNum(monthColumnsInfo[0].year)} ដល់ថ្ងៃទី${toKhmerNum(new Date(monthColumnsInfo[11].year, monthColumnsInfo[11].monthNum, 0).getDate())} ខែ${monthColumnsInfo[11].name} ឆ្នាំ${toKhmerNum(monthColumnsInfo[11].year)}`],
      [],
    ];

    const ws = XLSX.utils.aoa_to_sheet([...titleRows, headers1, ...dataRows]);
    XLSX.utils.book_append_sheet(wb, ws, 'តារាងប្រើប្រាស់តាមក្រុម');
    XLSX.writeFile(wb, `តារាងប្រើប្រាស់តាមក្រុម_${appliedStartYear}_${appliedStartMonth}.xlsx`);
  };

  // Export to PDF
  const handleExportPdf = async () => {
    if (!printAreaRef.current) return;
    setIsGeneratingPdf(true);
    try {
      await exportSinglePageA4LandscapePdf(
        printAreaRef.current,
        `តារាងប្រើប្រាស់តាមក្រុម_${startYear}_${startMonth}.pdf`,
        'yearly-team-usage-print-area'
      );
    } catch (err) {
      console.error('PDF export error:', err);
      handlePrint();
    } finally {
      setIsGeneratingPdf(false);
    }
  };

  // Direct Print
  const handlePrint = () => {
    printA4Document('yearly-team-usage-print-area', {
      orientation: 'landscape',
      documentTitle: `តារាងប្រើប្រាស់តាមក្រុម_${startYear}_${startMonth}`,
    });
  };

  // Helper to format table cell values (showing '-' when uncalculated / not yet queried)
  const renderCellValue = (val: number | undefined | null) => {
    if (!hasCalculated) return '-';
    if (val === undefined || val === null || isNaN(Number(val))) return '0';
    return Number(val).toLocaleString();
  };

  return (
    <div className="flex flex-col gap-4 pb-12 print:p-0 print:m-0">
      {/* Top Toolbar / Configuration Controls (Hidden during print) */}
      <div className="bg-white rounded-xl shadow-xs border border-gray-200 p-4 print:hidden space-y-4">
        {/* Title & Action Buttons Bar */}
        <div className="flex flex-wrap items-center justify-between gap-3 border-b border-gray-100 pb-3">
          <div className="flex items-center gap-2.5">
            <div className="w-9 h-9 rounded-lg bg-blue-100 text-blue-900 flex items-center justify-center font-bold">
              <Calendar className="w-5 h-5 text-blue-700" />
            </div>
            <div>
              <h2 className="text-base font-bold text-gray-800">
                តារាងប្រើប្រាស់តាមក្រុម — សន្លឹកទិដ្ឋាការស្អិត
              </h2>
              <p className="text-xs text-gray-500">
                តារាងសន្លឹកទិដ្ឋាការស្អិត តាមប្រភេទនីមួយៗ ប្រើប្រាស់តាមក្រុម ប្រចាំឆ្នាំ (១២ខែ)
              </p>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            {/* Check List / Monthly Breakdown Inspector Button */}
            <button
              onClick={() => {
                setChecklistInitialMonth(0);
                setShowMonthChecklistModal(true);
              }}
              type="button"
              className="w-36 h-9 rounded-lg text-xs font-medium bg-purple-700 hover:bg-purple-800 text-white flex items-center justify-center gap-1.5 shadow-xs transition cursor-pointer"
              title="ពិនិត្យផ្ទៀងផ្ទាត់បញ្ជី និងចំនួនសរុបតាមខែនីមួយៗ (Month Checklist & Total)"
            >
              <ListChecks className="w-3.5 h-3.5 shrink-0" />
              <span>ពិនិត្យបញ្ជីតាមខែ</span>
            </button>

            {/* Export PDF */}
            <button
              onClick={handleExportPdf}
              disabled={isGeneratingPdf}
              type="button"
              className="w-36 h-9 rounded-lg text-xs font-medium bg-rose-600 hover:bg-rose-700 text-white flex items-center justify-center gap-1.5 shadow-xs transition cursor-pointer disabled:opacity-50"
            >
              <Download className="w-3.5 h-3.5 shrink-0" />
              <span>{isGeneratingPdf ? 'កំពុងបង្កើត...' : 'ទាញយក PDF'}</span>
            </button>

            {/* Export Excel */}
            <button
              onClick={handleExportExcel}
              type="button"
              className="w-36 h-9 rounded-lg text-xs font-medium bg-emerald-700 hover:bg-emerald-800 text-white flex items-center justify-center gap-1.5 shadow-xs transition cursor-pointer"
            >
              <FileSpreadsheet className="w-3.5 h-3.5 shrink-0" />
              <span>Excel</span>
            </button>

            {/* Toggle Show/Hide Formatting Toolbar */}
            <button
              type="button"
              onClick={() => setShowFormattingToolbar(!showFormattingToolbar)}
              className={`h-9 px-3 rounded-lg text-xs font-bold border transition flex items-center gap-1.5 cursor-pointer shadow-2xs ${
                showFormattingToolbar
                  ? 'bg-slate-800 hover:bg-slate-900 text-white border-slate-900'
                  : 'bg-white hover:bg-gray-100 text-gray-700 border-gray-300'
              }`}
              title={showFormattingToolbar ? 'ចុចដើម្បីលាក់របារឧបករណ៍កែប្រែទម្រង់' : 'ចុចដើម្បីបង្ហាញរបារឧបករណ៍កែប្រែទម្រង់'}
            >
              {showFormattingToolbar ? (
                <>
                  <EyeOff className="w-3.5 h-3.5 text-amber-300" />
                  <span>លាក់របារឧបករណ៍</span>
                </>
              ) : (
                <>
                  <Eye className="w-3.5 h-3.5 text-blue-600" />
                  <span>បង្ហាញរបារឧបករណ៍</span>
                </>
              )}
            </button>

            {/* Close */}
            {onClose && (
              <button
                onClick={onClose}
                type="button"
                className="px-2.5 py-1.5 rounded-lg text-xs font-medium bg-gray-100 hover:bg-gray-200 text-gray-700 transition cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            )}
          </div>
        </div>

        {/* Date & Period Controls */}
        <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3.5 bg-white border border-slate-200/90 rounded-xl p-3.5 shadow-xs items-end">
          {/* Start Month */}
          <div>
            <label className="text-[12px] font-semibold text-slate-700 mb-1.5 flex items-center gap-1.5">
              <Calendar className="w-3.5 h-3.5 text-blue-600" />
              <span>ខែចាប់ផ្តើម ៖</span>
            </label>
            <select
              value={startMonth}
              onChange={(e) => {
                const newM = parseInt(e.target.value, 10);
                setStartMonth(newM);
                setActivePreset('custom');
                setHasCalculated(false);
              }}
              className={`w-full bg-slate-50/60 hover:bg-white border rounded-lg px-3 py-2 text-xs text-slate-800 focus:bg-white focus:outline-none focus:ring-2 font-medium transition-all cursor-pointer ${
                isFilterPending || !hasCalculated ? 'border-blue-400 ring-1 ring-blue-300' : 'border-slate-200 hover:border-slate-300 focus:border-blue-500 focus:ring-blue-500/20'
              }`}
            >
              {KHMER_MONTHS_NAMES.map((mName, idx) => (
                <option key={idx} value={idx + 1}>
                  ខែ{mName} (ខែទី{toKhmerNum(idx + 1)})
                </option>
              ))}
            </select>
          </div>

          {/* Start Year */}
          <div>
            <label className="text-[12px] font-semibold text-slate-700 mb-1.5 flex items-center gap-1.5">
              <CalendarDays className="w-3.5 h-3.5 text-blue-600" />
              <span>ឆ្នាំចាប់ផ្តើម ៖</span>
            </label>
            <select
              value={startYear}
              onChange={(e) => {
                const newY = parseInt(e.target.value, 10);
                setStartYear(newY);
                setActivePreset('custom');
                setHasCalculated(false);
              }}
              className={`w-full bg-slate-50/60 hover:bg-white border rounded-lg px-3 py-2 text-xs text-slate-800 focus:bg-white focus:outline-none focus:ring-2 font-medium transition-all cursor-pointer ${
                isFilterPending || !hasCalculated ? 'border-blue-400 ring-1 ring-blue-300' : 'border-slate-200 hover:border-slate-300 focus:border-blue-500 focus:ring-blue-500/20'
              }`}
            >
              {[2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025, 2026, 2027].map((yr) => (
                <option key={yr} value={yr}>
                  ឆ្នាំ {toKhmerNum(yr)} ({yr})
                </option>
              ))}
            </select>
          </div>

          {/* Show Result Button */}
          <div>
            <label className="text-[12px] font-semibold text-slate-700 mb-1.5 flex items-center gap-1.5">
              <Sparkles className="w-3.5 h-3.5 text-blue-600" />
              <span>សកម្មភាពទាញយក ៖</span>
            </label>
            <button
              type="button"
              onClick={() => handleApplyFilter()}
              disabled={isCalculating}
              className={`w-full py-2 px-3 rounded-lg text-xs font-bold flex items-center justify-center gap-1.5 shadow-xs transition-all cursor-pointer ${
                isFilterPending || !hasCalculated
                  ? 'bg-blue-600 hover:bg-blue-700 text-white ring-2 ring-blue-400/80 shadow-md animate-pulse'
                  : 'bg-gradient-to-r from-blue-700 to-indigo-800 hover:from-blue-800 hover:to-indigo-900 text-white'
              } disabled:opacity-50`}
              title="ជ្រើសកាលបរិច្ឆេទរួចហើយ ចុចប៊ូតុងនេះដើម្បីបង្ហាញទិន្នន័យ"
            >
              {isCalculating ? (
                <>
                  <RefreshCw className="w-3.5 h-3.5 animate-spin" />
                  <span>កំពុងទាញយក...</span>
                </>
              ) : (
                <>
                  <Search className="w-3.5 h-3.5 shrink-0" />
                  <span>{isFilterPending || !hasCalculated ? '⚡ បង្ហាញទិន្នន័យ' : '🔍 បង្ហាញទិន្នន័យ'}</span>
                </>
              )}
            </button>
          </div>

          {/* Signature Date */}
          <div>
            <label className="text-[12px] font-semibold text-slate-700 mb-1.5 flex items-center gap-1.5">
              <Calendar className="w-3.5 h-3.5 text-blue-600" />
              <span>កាលបរិច្ឆេទធ្វើតារាង ៖</span>
            </label>
            <CustomDatePicker
              value={reportDate}
              onChange={(d) => setReportDate(d)}
            />
          </div>

          {/* Signer Selection */}
          <div>
            <label className="text-[12px] font-semibold text-slate-700 mb-1.5 flex items-center gap-1.5">
              <UserCheck className="w-3.5 h-3.5 text-blue-600" />
              <span>មន្ត្រីចុះហត្ថលេខា ៖</span>
            </label>
            <div className="flex rounded-lg border border-slate-200 hover:border-slate-300 focus-within:border-blue-500 focus-within:ring-2 focus-within:ring-blue-500/20 bg-slate-50/60 focus-within:bg-white overflow-hidden transition-all">
              <input
                type="text"
                value={signerName}
                onChange={(e) => setSignerName(e.target.value)}
                placeholder="ឋានន្តរស័ក្កិ/ឈ្មោះ"
                className="flex-1 min-w-0 bg-transparent px-3 py-2 text-xs text-slate-800 focus:outline-none font-medium placeholder:text-slate-400"
              />
              {effectiveOfficers.length > 0 && (
                <select
                  value=""
                  onChange={(e) => {
                    const offId = e.target.value;
                    if (offId) {
                      const of = effectiveOfficers.find((o) => o.id === offId);
                      if (of) {
                        const fullTitle = getOfficerFormattedRankAndName(of);
                        setSignerName(fullTitle);
                      }
                    }
                  }}
                  className="bg-slate-100/90 hover:bg-slate-200/90 border-l border-slate-200 px-2 py-2 text-slate-700 focus:outline-none text-[11px] font-medium cursor-pointer transition-colors max-w-[110px]"
                  title="ជ្រើសរើសមន្ត្រីពីបញ្ជី"
                >
                  <option value="">ជ្រើសរើស</option>
                  {effectiveOfficers.map((of) => {
                    const nameOnly = getOfficerNameOnly(of);
                    const positionName =
                      categories?.positions?.find((p) => p.id === of.positionId)?.name ||
                      (of.rank && of.rank.includes('មន្ត្រី') ? of.rank : 'មន្ត្រី');
                    return (
                      <option key={of.id} value={of.id}>
                        {nameOnly} ({positionName})
                      </option>
                    );
                  })}
                </select>
              )}
            </div>
          </div>
        </div>

        {/* Action Controls & Formatting Tools */}
        {showFormattingToolbar && (
          <div className="flex flex-wrap items-center justify-between gap-3 text-xs pt-1 border-t border-gray-100 mt-1">
          {/* Zoom & Font Controls */}
          <div className="flex flex-wrap items-center gap-2">
            <span className="font-semibold text-gray-600">ទំហំបង្ហាញ (Zoom):</span>
            <button
              onClick={() => setZoomLevel((z) => Math.max(60, z - 10))}
              className="p-1 rounded bg-gray-100 hover:bg-gray-200 text-gray-700"
              title="បង្រួម (Zoom Out)"
            >
              <ZoomOut className="w-3.5 h-3.5" />
            </button>
            <span className="font-bold text-gray-700 min-w-[3rem] text-center">{zoomLevel}%</span>
            <button
              onClick={() => setZoomLevel((z) => Math.min(150, z + 10))}
              className="p-1 rounded bg-gray-100 hover:bg-gray-200 text-gray-700"
              title="ពង្រីក (Zoom In)"
            >
              <ZoomIn className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={() => setZoomLevel(100)}
              className="px-2 py-0.5 rounded bg-gray-100 hover:bg-gray-200 text-gray-700 text-[11px]"
            >
              ដើម (100%)
            </button>

            <div className="h-4 w-px bg-gray-300 mx-1" />

            <span className="font-semibold text-gray-600">ពុម្ពអក្សរ:</span>
            <select
              value={fontFamily}
              onChange={(e) => setFontFamily(e.target.value)}
              className="bg-white border border-gray-300 rounded px-2 py-1 focus:outline-none font-medium"
            >
              <option value="Khmer OS Siemreap">Khmer OS Siemreap (ស្តង់ដារ)</option>
              <option value="Khmer OS Muol Light">Khmer OS Muol Light</option>
              <option value="Times New Roman">Times New Roman</option>
            </select>

            <span className="font-semibold text-gray-600 ml-2">ទំហំក្បាលលិខិត:</span>
            <select
              value={headerFontSize}
              onChange={(e) => setHeaderFontSize(parseFloat(e.target.value))}
              className="bg-white border border-amber-300 text-amber-950 font-bold rounded px-2 py-1 focus:outline-none"
              title="កំណត់ទំហំអក្សរក្បាលលិខិត (ក្រសួង និង ព្រះរាជាណាចក្រកម្ពុជា)"
            >
              <option value={10}>10pt</option>
              <option value={11}>11pt</option>
              <option value={11.5}>11.5pt</option>
              <option value={12}>12pt (ស្តង់ដារ)</option>
              <option value={12.5}>12.5pt</option>
              <option value={13}>13pt</option>
              <option value={14}>14pt</option>
            </select>

            <span className="font-semibold text-gray-600 ml-2">ទំហំអក្សរតារាង:</span>
            <select
              value={fontSize}
              onChange={(e) => setFontSize(parseInt(e.target.value, 10))}
              className="bg-white border border-gray-300 rounded px-2 py-1 focus:outline-none"
            >
              <option value={10}>តូចខ្លាំង (10px)</option>
              <option value={11}>ស្តង់ដារ PDF (11px)</option>
              <option value={12}>មធ្យម (12px / 9pt)</option>
              <option value={13}>ធំល្មម (13px / 10pt)</option>
              <option value={14}>ធំ (14px / 10.5pt)</option>
              <option value={16}>ធំខ្លាំង (16px / 12pt)</option>
            </select>

            {/* Compact Table Row Height Control */}
            <div className="flex items-center border border-gray-300 rounded bg-white px-2 py-0.5 text-xs text-gray-700 font-siemreap ml-1">
              <span className="text-[11px] text-gray-600 mr-1 select-none font-medium">កម្ពស់ជួរ៖</span>
              <button
                type="button"
                onClick={() => {
                  const val = Math.max(15, tableRowHeight - 1);
                  setTableRowHeight(val);
                  localStorage.setItem('yearly_team_usage_row_height', String(val));
                }}
                className="w-4 h-4 flex items-center justify-center rounded hover:bg-gray-100 text-gray-700 font-bold cursor-pointer transition text-xs"
                title="បង្រួមកម្ពស់ជួរតារាង (-1px)"
              >
                -
              </button>
              <span className="font-mono text-xs font-bold px-1 text-[#002060] select-none">
                {tableRowHeight}px
              </span>
              <button
                type="button"
                onClick={() => {
                  const val = Math.min(30, tableRowHeight + 1);
                  setTableRowHeight(val);
                  localStorage.setItem('yearly_team_usage_row_height', String(val));
                }}
                className="w-4 h-4 flex items-center justify-center rounded hover:bg-gray-100 text-gray-700 font-bold cursor-pointer transition text-xs"
                title="ពង្រីកកម្ពស់ជួរតារាង (+1px)"
              >
                +
              </button>
            </div>

            <span className="font-semibold text-gray-600 ml-2">ពណ៌ឈ្មោះអ្នកធ្វើ:</span>
            <select
              value={nameColor}
              onChange={(e) => setNameColor(e.target.value)}
              className="bg-white border border-gray-300 rounded px-2 py-1 focus:outline-none font-medium"
            >
              <option value="#C00000" className="text-red-700 font-bold">ពណ៌ក្រហម (#C00000)</option>
              <option value="#000000" className="text-black font-bold">ពណ៌ខ្មៅ (ស្តង់ដារ)</option>
              <option value="#002060" className="text-blue-900 font-bold">ពណ៌ខៀវចាស់</option>
            </select>

            <div className="h-4 w-px bg-gray-300 mx-1" />

            {/* Date Number Format Selector (Khmer digits vs English digits) */}
            <span className="font-semibold text-gray-600 ml-1">លេខកាលបរិច្ឆេទ:</span>
            <select
              value={dateNumberFormat}
              onChange={(e) => setDateNumberFormat(e.target.value as 'khmer' | 'latin')}
              className="bg-white border border-blue-300 text-blue-950 font-bold rounded px-2 py-1 focus:outline-none"
              title="ជ្រើសរើសទម្រង់លេខខ្មែរ (០៨, ២០១៩) ឬលេខអង់គ្លេស (8, 2019) លើកាលបរិច្ឆេទ"
            >
              <option value="khmer">លេខខ្មែរ (០៨, ២០១៩)</option>
              <option value="latin">លេខអង់គ្លេស (8, 2019)</option>
            </select>

            {/* Quick Signature Shift Stepper (រំកិលហត្ថលេខាឡើង/ចុះ) */}
            <div className="flex items-center border border-teal-300 rounded bg-teal-50/70 px-2 py-1 text-xs text-teal-950 font-siemreap ml-1">
              <span className="text-[11px] text-teal-800 mr-1 select-none font-semibold">រំកិលហត្ថលេខា៖</span>
              <button
                type="button"
                onClick={() => setSignatureShiftY((prev) => Math.max(-30, prev - 2))}
                className="w-5 h-5 flex items-center justify-center rounded hover:bg-teal-200 text-teal-900 font-bold cursor-pointer transition"
                title="រំកិលហត្ថលេខាឡើងលើ (-2pt)"
              >
                -
              </button>
              <span className="font-mono text-xs font-bold px-1.5 text-teal-900 select-none min-w-[32px] text-center">
                {signatureShiftY}pt
              </span>
              <button
                type="button"
                onClick={() => setSignatureShiftY((prev) => Math.min(60, prev + 2))}
                className="w-5 h-5 flex items-center justify-center rounded hover:bg-teal-200 text-teal-900 font-bold cursor-pointer transition"
                title="ទម្លាក់ហត្ថលេខាចុះក្រោម (+2pt)"
              >
                +
              </button>
            </div>

            {/* Quick Signer Gap Stepper (គម្លាតពីអ្នកធ្វើតារាងទៅឈ្មោះមន្ត្រីចុះហត្ថលេខា) */}
            <div className="flex items-center border border-indigo-300 rounded bg-indigo-50/70 px-2 py-1 text-xs text-indigo-950 font-siemreap ml-1">
              <span className="text-[11px] text-indigo-800 mr-1 select-none font-semibold">គម្លាតឈ្មោះមន្ត្រី៖</span>
              <button
                type="button"
                onClick={() => setSignerGapHeight((prev) => Math.max(30, prev - 5))}
                className="w-5 h-5 flex items-center justify-center rounded hover:bg-indigo-200 text-indigo-900 font-bold cursor-pointer transition"
                title="បន្ថយគម្លាតចុះហត្ថលេខា (-5px)"
              >
                -
              </button>
              <span className="font-mono text-xs font-bold px-1.5 text-indigo-900 select-none min-w-[36px] text-center">
                {signerGapHeight}px
              </span>
              <button
                type="button"
                onClick={() => setSignerGapHeight((prev) => Math.min(180, prev + 5))}
                className="w-5 h-5 flex items-center justify-center rounded hover:bg-indigo-200 text-indigo-900 font-bold cursor-pointer transition"
                title="បន្ថែមគម្លាតចុះហត្ថលេខា (+5px)"
              >
                +
              </button>
            </div>

            {/* Tacteing Style Button */}
            <button
              type="button"
              onClick={() => setShowTacteingSelector((prev) => !prev)}
              className={`px-2.5 py-1 rounded border text-xs font-medium flex items-center gap-1.5 transition ml-1 ${
                showTacteingSelector
                  ? 'bg-amber-100 border-amber-400 text-amber-900 font-bold'
                  : 'bg-white border-gray-300 text-gray-700 hover:bg-gray-50'
              }`}
              title="ជ្រើសរើស ឬប្តូរទម្រង់តាក់តែង (Tacteing)"
            >
              <Sparkles className="w-3.5 h-3.5 text-amber-600" />
              <span>តាក់តែង ({tacteingSettings.type === 'image2-classic' ? 'បុរាណ' : tacteingSettings.type === 'symbol-cross' ? '-( + )-' : tacteingSettings.type === 'symbol-flower' ? '-( ❁ )-' : 'រូបភាព'})</span>
            </button>
          </div>

          {/* Recalculate & Reset */}
          <div className="flex items-center gap-2">
            <label className="flex items-center gap-1.5 text-gray-700 cursor-pointer font-medium select-none">
              <input
                type="checkbox"
                checked={autoCalculate}
                onChange={(e) => setAutoCalculate(e.target.checked)}
                className="rounded text-blue-600 focus:ring-blue-500"
              />
              <span>គណនាសរុបស្វ័យប្រវត្តិក្នងជួរ (Auto Sum)</span>
            </label>

            <button
              onClick={handleRecalculateAll}
              type="button"
              className="px-2.5 py-1 rounded bg-gray-100 hover:bg-gray-200 text-gray-700 text-xs font-medium flex items-center gap-1 transition"
              title="គណនាផលបូកឡើងវិញទាំងអស់"
            >
              <RefreshCw className="w-3 h-3" />
              <span>គណនាឡើងវិញ</span>
            </button>

            <button
              onClick={handleClearAll}
              type="button"
              className="px-2.5 py-1 rounded bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 text-xs font-medium flex items-center gap-1 transition"
              title="កំណត់ទិន្នន័យជា ០ ទាំងអស់"
            >
              <RotateCcw className="w-3 h-3" />
              <span>កំណត់ឡើងវិញ (Reset)</span>
            </button>

            {/* Quick Hide Button inside toolbar */}
            <button
              onClick={() => setShowFormattingToolbar(false)}
              type="button"
              className="px-2 py-1 rounded bg-gray-100 hover:bg-gray-200 text-gray-500 hover:text-gray-800 text-xs font-medium flex items-center gap-1 transition cursor-pointer"
              title="លាក់របារឧបករណ៍"
            >
              <EyeOff className="w-3 h-3 text-gray-500" />
              <span>លាក់</span>
            </button>
          </div>
        </div>
        )}

        {/* Tacteing Selector Panel when toggled */}
        {showFormattingToolbar && showTacteingSelector && (
          <div className="mt-2.5 pt-2.5 border-t border-amber-200">
            <TacteingControlSelector
              currentType={tacteingSettings.type}
              customImage={tacteingSettings.customImage}
              onChange={(t, img) => {
                const newSettings = { type: t, customImage: img || null };
                setTacteingSettings(newSettings);
                saveTacteingSettings(t, img);
              }}
            />
          </div>
        )}
      </div>

      {/* Printable Report Canvas Area */}
      <div className="overflow-x-auto w-full flex justify-center print:overflow-visible print:block">
        <div
          ref={printAreaRef}
          id="yearly-team-usage-print-area"
          style={{
            transform: `scale(${zoomLevel / 100})`,
            transformOrigin: 'top center',
            fontFamily,
            fontSize: `${fontSize}px`,
            lineHeight: lineSpacing,
          }}
          className="bg-white text-black shadow-lg print:shadow-none border border-gray-300 print:border-none p-6 md:p-8 w-[1120px] shrink-0 print:w-full print:p-2 transition-transform duration-100"
        >
          {/* Header Layout (Ministry on left, Kingdom on right) */}
          <div className="flex justify-between items-start mb-3 leading-tight">
            {/* Left Header: Department & Office */}
            <div
              style={{ fontSize: `${headerFontSize}pt` }}
              className="text-center flex flex-col items-center leading-tight pt-5"
            >
              <p
                contentEditable
                suppressContentEditableWarning
                onBlur={(e) => setMinistryName(e.currentTarget.textContent || '')}
                style={{ fontSize: `${headerFontSize}pt` }}
                className="font-moul leading-tight outline-none hover:bg-amber-50/50 rounded px-1"
              >
                {ministryName}
              </p>
              <p
                contentEditable
                suppressContentEditableWarning
                onBlur={(e) => setDepartmentName(e.currentTarget.textContent || '')}
                style={{ fontSize: `${headerFontSize}pt` }}
                className="font-moul leading-tight outline-none hover:bg-amber-50/50 rounded px-1 mt-0.5"
              >
                {departmentName}
              </p>
              <p
                contentEditable
                suppressContentEditableWarning
                onBlur={(e) => setGeneralDeptName(e.currentTarget.textContent || '')}
                style={{ fontSize: `${headerFontSize}pt` }}
                className="font-moul leading-tight outline-none hover:bg-amber-50/50 rounded px-1 mt-0.5"
              >
                {generalDeptName}
              </p>
              <p
                contentEditable
                suppressContentEditableWarning
                onBlur={(e) => setOfficeName(e.currentTarget.textContent || '')}
                style={{ fontSize: `${headerFontSize}pt` }}
                className="font-moul leading-tight outline-none hover:bg-amber-50/50 rounded px-1 mt-0.5"
              >
                {officeName}
              </p>
              <p
                contentEditable
                suppressContentEditableWarning
                onBlur={(e) => setSectionName(e.currentTarget.textContent || '')}
                style={{
                  fontSize: sectionName.length > 20 || sectionName.includes('បាវិត') ? '10pt' : '12pt',
                  fontFamily: "'Khmer Mool1', 'Khmer Mool', 'Khmer OS Muol Light', 'Khmer OS Muol', 'Moul', serif"
                }}
                className="font-moul leading-tight outline-none hover:bg-amber-50/50 rounded px-1 mt-0.5 whitespace-nowrap"
              >
                {sectionName}
              </p>
              <div className="mt-1">
                <TacteingLine
                  type={tacteingSettings.type}
                  customImage={tacteingSettings.customImage}
                  width={100}
                  height={8}
                />
              </div>
            </div>

            {/* Right Header: Kingdom */}
            <div
              style={{ fontSize: `${headerFontSize}pt` }}
              className="text-center flex flex-col items-center leading-tight"
            >
              <p
                style={{ fontSize: `${headerFontSize}pt` }}
                className="font-moul leading-tight tracking-wider"
              >
                ព្រះរាជាណាចក្រកម្ពុជា
              </p>
              <p
                style={{ fontSize: `${headerFontSize}pt` }}
                className="font-moul leading-tight tracking-wider mt-0.5"
              >
                ជាតិ សាសនា ព្រះមហាក្សត្រ
              </p>
              <div className="mt-1">
                <TacteingLine
                  type={tacteingSettings.type}
                  customImage={tacteingSettings.customImage}
                  width={100}
                  height={8}
                />
              </div>
            </div>
          </div>

          {/* Central Title */}
          <div className="text-center my-3">
            <h1
              style={{
                fontSize: '12pt',
                fontFamily: "'Khmer Mool1', 'Khmer Mool', 'Khmer OS Muol Light', 'Khmer OS Muol', 'Moul', serif"
              }}
              className="text-[12pt] font-moul text-black font-normal leading-relaxed"
            >
              តារាងសន្លឹក ទិដ្ឋាការស្អិត តាមប្រភេទនីមួយៗ បានផ្តល់ជូនភ្ញៀវ
            </h1>
            <p
              style={{ fontSize: '12pt' }}
              className="text-[12pt] font-siemreap font-bold mt-1 text-black"
            >
              ដោយគិតចាប់ពីថ្ងៃទី០១ ខែ{KHMER_MONTHS_NAMES[appliedStartMonth - 1]} ឆ្នាំ{toKhmerNum(appliedStartYear)} រហូតដល់ថ្ងៃទី{toKhmerNum(new Date(monthColumnsInfo[11].year, monthColumnsInfo[11].monthNum, 0).getDate())} ខែ{monthColumnsInfo[11].name} ឆ្នាំ{toKhmerNum(monthColumnsInfo[11].year)}
            </p>
          </div>

          {/* Main 13-Type Usage Table */}
          <div className="w-full overflow-x-auto print:overflow-visible">
            <table
              className="w-full border-collapse border border-black text-black text-center text-[11px] leading-normal select-text"
              style={{
                '--table-row-height': `${tableRowHeight}px`,
                fontFamily: fontFamily || 'Khmer OS Siemreap'
              } as React.CSSProperties}
            >
              <thead>
                {/* Header Row 1 */}
                <tr className="bg-[#D9E1F2] font-bold border-b border-black text-black h-[26px]">
                  <th
                    rowSpan={2}
                    className="border border-black px-0.5 py-1 min-w-[34px] max-w-[38px] text-center align-middle font-bold font-siemreap text-[10.5px] whitespace-nowrap"
                  >
                    ប្រភេទ
                  </th>
                  <th
                    rowSpan={2}
                    className="border border-black px-1 py-1 min-w-[95px] text-center align-middle font-bold font-siemreap"
                  >
                    <div className="whitespace-nowrap font-bold text-[11px] leading-tight">សន្និធិចុងគ្រា</div>
                    <div className="text-[9px] font-semibold text-gray-800 mt-0.5 whitespace-nowrap leading-tight">
                      {openingStockLabel}
                    </div>
                  </th>
                  <th
                    colSpan={13}
                    className="border border-black px-1 py-1 text-center font-bold font-siemreap text-[11px] leading-normal whitespace-nowrap h-[24px] align-middle"
                  >
                    គិតចាប់ពីថ្ងៃទី០១ ខែ{KHMER_MONTHS_NAMES[appliedStartMonth - 1]} ឆ្នាំ{toKhmerNum(appliedStartYear)} រហូតដល់ថ្ងៃទី{toKhmerNum(new Date(monthColumnsInfo[11].year, monthColumnsInfo[11].monthNum, 0).getDate())} ខែ{monthColumnsInfo[11].name} ឆ្នាំ{toKhmerNum(monthColumnsInfo[11].year)}
                  </th>
                  <th
                    rowSpan={2}
                    onClick={() => setShowReceivedK2ChecklistModal(true)}
                    className="border border-black px-0.5 py-1 min-w-[50px] max-w-[58px] text-center align-middle font-bold font-siemreap text-[10px] leading-tight hover:bg-blue-200 transition cursor-pointer select-none group"
                    title={!isSecondary ? "ចុចដើម្បីពិនិត្យផ្ទៀងផ្ទាត់បញ្ជីទិដ្ឋាការ ទទួលពីការិយាល័យ ឬផ្ទេរពីក្រុម (បើកពីក២) ទាំង ១២ខែ" : "ចុចដើម្បីពិនិត្យផ្ទៀងផ្ទាត់បញ្ជីទិដ្ឋាការ បើកពីក២ ទាំង ១២ខែ"}
                  >
                    <div className="group-hover:underline whitespace-nowrap">ចំនួន</div>
                    <div className="group-hover:underline whitespace-nowrap mt-0.5">បើកពីក២</div>
                  </th>
                  <th
                    rowSpan={2}
                    className="border border-black px-0.5 py-1 min-w-[46px] max-w-[54px] text-center align-middle font-bold font-siemreap text-[10px] leading-tight"
                  >
                    <div className="whitespace-nowrap">ចំនួន</div>
                    <div className="whitespace-nowrap mt-0.5">សរុប</div>
                  </th>
                  <th
                    rowSpan={2}
                    className="border border-black px-0.5 py-1 min-w-[48px] max-w-[56px] text-center align-middle font-bold font-siemreap text-[10px] leading-tight"
                  >
                    <div className="whitespace-nowrap">ប្រើប្រាស់</div>
                    <div className="whitespace-nowrap mt-0.5">សរុប</div>
                  </th>
                  <th
                    rowSpan={2}
                    onClick={() => setShowDamagedChecklistModal(true)}
                    className="border border-black px-0.5 py-1 min-w-[44px] max-w-[50px] text-center align-middle font-bold font-siemreap text-[10px] leading-tight hover:bg-red-200 transition cursor-pointer select-none group"
                    title={!isSecondary ? "ចុចដើម្បីពិនិត្យផ្ទៀងផ្ទាត់បញ្ជីទិដ្ឋាការ ផ្ទេរ / ខ្វះ / ខូច ទាំង ១២ខែ" : "ចុចដើម្បីពិនិត្យផ្ទៀងផ្ទាត់បញ្ជីទិដ្ឋាការ ខ្វះ & ខូច ទាំង ១២ខែ"}
                  >
                    {!isSecondary ? (
                      <div className="group-hover:underline whitespace-nowrap">ផ្ទេរ/ខ្វះ/ខូច</div>
                    ) : (
                      <>
                        <div className="group-hover:underline whitespace-nowrap">ទិដ្ឋាការ</div>
                        <div className="text-[8.5px] group-hover:underline whitespace-nowrap mt-0.5 leading-tight">ខ្វះ&ខូច</div>
                      </>
                    )}
                  </th>
                  <th
                    rowSpan={2}
                    className="border border-black px-0.5 py-1 min-w-[42px] max-w-[48px] text-center align-middle font-bold font-siemreap text-[10px] leading-tight"
                  >
                    <div className="whitespace-nowrap">បង្វិលក២</div>
                  </th>
                  <th
                    rowSpan={2}
                    className="border border-black px-1 py-1 min-w-[95px] text-center align-middle font-bold font-siemreap"
                  >
                    <div className="whitespace-nowrap font-bold text-[11px] leading-tight">សន្និធិនៅសល់</div>
                    <div className="text-[9px] font-semibold text-gray-800 mt-0.5 whitespace-nowrap leading-tight">
                      {endingStockLabel}
                    </div>
                  </th>
                </tr>

                {/* Header Row 2: 12 Months + Sum */}
                <tr className="bg-[#D9E1F2] font-bold border-b border-black text-[10.5px] font-siemreap leading-normal text-black h-[20px]">
                  {monthColumnsInfo.map((mCol, idx) => (
                    <th
                      key={idx}
                      onClick={() => {
                        setChecklistInitialMonth(idx);
                        setShowMonthChecklistModal(true);
                      }}
                      className="border border-black px-0.5 py-0.5 min-w-[44px] text-center hover:bg-blue-200 transition cursor-pointer select-none group whitespace-nowrap"
                      title={`ចុចដើម្បីពិនិត្យផ្ទៀងផ្ទាត់បញ្ជី និងចំនួនប្រើប្រាស់ខែ ${mCol.name} (${mCol.year})`}
                    >
                      <span className="group-hover:underline">{mCol.name}</span>
                    </th>
                  ))}
                  <th
                    onClick={() => {
                      setChecklistInitialMonth(-1);
                      setShowMonthChecklistModal(true);
                    }}
                    className="border border-black px-1 py-0.5 min-w-[52px] text-center font-bold bg-[#B4C6E7] hover:bg-blue-200 transition cursor-pointer select-none whitespace-nowrap"
                    title="ចុចដើម្បីពិនិត្យតារាងសរុបទាំង ១២ខែ"
                  >
                    សរុប
                  </th>
                </tr>
              </thead>

              <tbody>
                {REPORT_VISA_TYPES.map((vt, rowIdx) => {
                  const row = tableData[vt] || {
                    visaType: vt,
                    openingStock: 0,
                    monthlyValues: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
                    totalMonthly: 0,
                    receivedK2: 0,
                    totalAvailable: 0,
                    used: 0,
                    damagedMissing: 0,
                    returnStock: 0,
                    remaining: 0,
                  };

                  const dynamicMonthlySum = row.monthlyValues.reduce((a, b) => (Number(a) || 0) + (Number(b) || 0), 0);
                  const displayMonthlySum = autoCalculate ? dynamicMonthlySum : (row.totalMonthly ?? dynamicMonthlySum);
                  const displayAvailable = autoCalculate
                    ? (Number(row.openingStock) || 0) + (Number(row.receivedK2) || 0)
                    : (row.totalAvailable ?? ((Number(row.openingStock) || 0) + (Number(row.receivedK2) || 0)));
                  const displayUsed = autoCalculate ? dynamicMonthlySum : (row.used ?? dynamicMonthlySum);
                  const displayRemaining = autoCalculate
                    ? displayAvailable - (Number(displayUsed) || 0) - (Number(row.damagedMissing) || 0) - (Number(row.returnStock) || 0)
                    : (row.remaining ?? (displayAvailable - (Number(displayUsed) || 0) - (Number(row.damagedMissing) || 0) - (Number(row.returnStock) || 0)));

                  return (
                    <tr
                      key={vt}
                      style={{ height: `${tableRowHeight}px` }}
                      className={`hover:bg-blue-50/40 transition-colors ${
                        rowIdx % 2 === 1 ? 'bg-gray-50/30' : 'bg-white'
                      }`}
                    >
                      {/* Visa Type Code */}
                      <td
                        style={{ fontFamily: "'Times New Roman', Times, serif" }}
                        className="border border-black px-0.5 py-[1px] font-bold text-center font-times bg-gray-50/50 min-w-[34px] max-w-[38px] text-[11px] leading-tight"
                      >
                        {vt}
                      </td>

                      {/* Opening Stock (Editable) */}
                      <td
                        contentEditable
                        suppressContentEditableWarning
                        onBlur={(e) => handleCellChange(vt, 'openingStock', e.currentTarget.textContent || '0')}
                        className="border border-black px-1 py-[1px] text-right font-normal font-times outline-none hover:bg-yellow-100/70 focus:bg-yellow-100 focus:ring-1 focus:ring-blue-500 text-[11px] leading-tight"
                      >
                        {renderCellValue(row.openingStock)}
                      </td>

                      {/* 12 Monthly Values (Editable) */}
                      {monthColumnsInfo.map((_, mIdx) => {
                        const val = row.monthlyValues[mIdx] || 0;
                        return (
                          <td
                            key={mIdx}
                            contentEditable
                            suppressContentEditableWarning
                            onBlur={(e) =>
                              handleCellChange(vt, 'month', e.currentTarget.textContent || '0', mIdx)
                            }
                            className="border border-black px-0.5 py-[1px] text-right font-times outline-none hover:bg-yellow-100/70 focus:bg-yellow-100 focus:ring-1 focus:ring-blue-500 font-normal text-[10.5px] leading-tight"
                          >
                            {renderCellValue(val)}
                          </td>
                        );
                      })}

                      {/* Total Monthly Sum (12 Months Usage Sum) */}
                      <td
                        contentEditable={!autoCalculate}
                        suppressContentEditableWarning
                        onBlur={(e) =>
                          !autoCalculate &&
                          handleCellChange(vt, 'totalMonthly', e.currentTarget.textContent || '0')
                        }
                        className={`border border-black px-1 py-[1px] text-right font-normal font-times bg-blue-50/30 text-[11px] leading-tight ${
                          !autoCalculate ? 'outline-none hover:bg-yellow-100' : ''
                        }`}
                      >
                        {renderCellValue(displayMonthlySum)}
                      </td>

                      {/* Received from K2 (ចំនួន បើកពីក២) */}
                      <td
                        contentEditable
                        suppressContentEditableWarning
                        onBlur={(e) => handleCellChange(vt, 'receivedK2', e.currentTarget.textContent || '0')}
                        className="border border-black px-0.5 py-[1px] text-right font-normal font-times outline-none hover:bg-yellow-100/70 focus:bg-yellow-100 focus:ring-1 focus:ring-blue-500 text-[11px] leading-tight"
                      >
                        {renderCellValue(row.receivedK2)}
                      </td>

                      {/* Total Available (សរុប = ដើមគ្រា + បើកពីក២) */}
                      <td
                        contentEditable={!autoCalculate}
                        suppressContentEditableWarning
                        onBlur={(e) =>
                          !autoCalculate &&
                          handleCellChange(vt, 'totalAvailable', e.currentTarget.textContent || '0')
                        }
                        className={`border border-black px-0.5 py-[1px] text-right font-bold font-times bg-blue-50/40 text-[11px] leading-tight ${
                          !autoCalculate ? 'outline-none hover:bg-yellow-100' : ''
                        }`}
                      >
                        {renderCellValue(displayAvailable)}
                      </td>

                      {/* Used (ប្រើប្រាស់ សរុប) */}
                      <td
                        contentEditable={!autoCalculate}
                        suppressContentEditableWarning
                        onBlur={(e) =>
                          !autoCalculate &&
                          handleCellChange(vt, 'used', e.currentTarget.textContent || '0')
                        }
                        className={`border border-black px-0.5 py-[1px] text-right font-normal font-times text-[11px] leading-tight ${
                          !autoCalculate ? 'outline-none hover:bg-yellow-100' : ''
                        }`}
                      >
                        {renderCellValue(displayUsed)}
                      </td>

                      {/* Missing & Damaged (Editable) */}
                      <td
                        contentEditable
                        suppressContentEditableWarning
                        onBlur={(e) =>
                          handleCellChange(vt, 'damagedMissing', e.currentTarget.textContent || '0')
                        }
                        className="border border-black px-0.5 py-[1px] text-right font-normal font-times outline-none hover:bg-yellow-100/70 focus:bg-yellow-100 focus:ring-1 focus:ring-blue-500 text-[11px] leading-tight"
                      >
                        {renderCellValue(row.damagedMissing)}
                      </td>

                      {/* Return to Stock / បង្វិលក២ (Editable) */}
                      <td
                        contentEditable
                        suppressContentEditableWarning
                        onBlur={(e) =>
                          handleCellChange(vt, 'returnStock', e.currentTarget.textContent || '0')
                        }
                        className="border border-black px-0.5 py-[1px] text-right font-normal font-times outline-none hover:bg-yellow-100/70 focus:bg-yellow-100 focus:ring-1 focus:ring-blue-500 text-[11px] leading-tight"
                      >
                        {renderCellValue(row.returnStock)}
                      </td>

                      {/* Ending Remaining Stock */}
                      <td
                        contentEditable={!autoCalculate}
                        suppressContentEditableWarning
                        onBlur={(e) =>
                          !autoCalculate &&
                          handleCellChange(vt, 'remaining', e.currentTarget.textContent || '0')
                        }
                        className={`border border-black px-1 py-[1px] text-right font-bold font-times bg-blue-50/50 text-[11px] leading-tight ${
                          !autoCalculate ? 'outline-none hover:bg-yellow-100' : ''
                        }`}
                      >
                        {renderCellValue(displayRemaining)}
                      </td>
                    </tr>
                  );
                })}

                {/* Grand Total Row / សរុប */}
                <tr
                  style={{ height: `${Math.max(20, tableRowHeight + 2)}px` }}
                  className="bg-[#D9E1F2] font-bold border-t-2 border-black text-black"
                >
                  <td className="border border-black px-0.5 py-[1px] text-center font-bold font-siemreap text-[11px] min-w-[34px] max-w-[38px] leading-tight">
                    សរុប
                  </td>

                  {/* Opening Grand Total */}
                  <td className="border border-black px-1 py-[1px] text-right font-bold font-times text-[11px] leading-tight">
                    {renderCellValue(grandTotals.opening)}
                  </td>

                  {/* 12 Months Grand Totals */}
                  {grandTotals.months.map((mTotal, mIdx) => (
                    <td
                      key={mIdx}
                      className="border border-black px-0.5 py-[1px] text-right font-bold font-times text-[10.5px] leading-tight"
                    >
                      {renderCellValue(mTotal)}
                    </td>
                  ))}

                  {/* Monthly Sum Grand Total */}
                  <td className="border border-black px-1 py-[1px] text-right font-bold font-times bg-[#B4C6E7] text-[11px] leading-tight">
                    {renderCellValue(grandTotals.monthlySum)}
                  </td>

                  {/* Received from K2 Grand Total */}
                  <td
                    onClick={() => setShowReceivedK2ChecklistModal(true)}
                    className="border border-black px-0.5 py-[1px] text-right font-bold font-times hover:bg-blue-200 transition cursor-pointer text-[11px] leading-tight"
                    title={!isSecondary ? "ចុចដើម្បីពិនិត្យផ្ទៀងផ្ទាត់បញ្ជីទិដ្ឋាការ ទទួលពីការិយាល័យ ឬផ្ទេរពីក្រុម (បើកពីក២) ទាំង ១២ខែ" : "ចុចដើម្បីពិនិត្យផ្ទៀងផ្ទាត់បញ្ជីទិដ្ឋាការ បើកពីក២ ទាំង ១២ខែ"}
                  >
                    {renderCellValue(grandTotals.receivedK2)}
                  </td>

                  {/* Available Grand Total */}
                  <td className="border border-black px-0.5 py-[1px] text-right font-bold font-times bg-[#8EA9DB]/60 text-[11px] leading-tight">
                    {renderCellValue(grandTotals.available)}
                  </td>

                  {/* Used Grand Total */}
                  <td className="border border-black px-0.5 py-[1px] text-right font-bold font-times text-[11px] leading-tight">
                    {renderCellValue(grandTotals.used)}
                  </td>

                  {/* Missing/Damaged Grand Total */}
                  <td
                    onClick={() => setShowDamagedChecklistModal(true)}
                    className="border border-black px-0.5 py-[1px] text-right font-bold font-times hover:bg-red-200 transition cursor-pointer text-[11px] leading-tight"
                    title={!isSecondary ? "ចុចដើម្បីពិនិត្យផ្ទៀងផ្ទាត់បញ្ជីទិដ្ឋាការ ផ្ទេរ / ខ្វះ / ខូច ទាំង ១២ខែ" : "ចុចដើម្បីពិនិត្យផ្ទៀងផ្ទាត់បញ្ជីទិដ្ឋាការ ខ្វះ & ខូច ទាំង ១២ខែ"}
                  >
                    {renderCellValue(grandTotals.damagedMissing)}
                  </td>

                  {/* Return Grand Total */}
                  <td className="border border-black px-0.5 py-[1px] text-right font-bold font-times text-[11px] leading-tight">
                    {renderCellValue(grandTotals.returnStock)}
                  </td>

                  {/* Remaining Grand Total */}
                  <td className="border border-black px-1 py-[1px] text-right font-bold font-times bg-[#8EA9DB]/60 text-[11px] leading-tight">
                    {renderCellValue(grandTotals.remaining)}
                  </td>
                </tr>
              </tbody>
            </table>
          </div>

          {/* Footer Signature Section (Right-aligned matching Picture 2) */}
          <div className="flex justify-end" style={{ marginTop: `${signatureShiftY}pt` }}>
            <div className="text-center text-[12pt] min-w-[280px] leading-snug font-siemreap" style={{ fontSize: '12pt' }}>
              {/* Khmer Lunar Date */}
              <p
                contentEditable
                suppressContentEditableWarning
                style={{ fontSize: '12pt' }}
                className="outline-none hover:bg-amber-50/50 rounded px-1 text-slate-900 text-[12pt]"
              >
                {signatureLunarDate}
              </p>

              {/* Solar Date */}
              <p
                contentEditable
                suppressContentEditableWarning
                style={{ fontSize: '12pt' }}
                className="outline-none hover:bg-amber-50/50 rounded px-1 text-slate-900 font-medium mt-0.5 text-[12pt]"
              >
                ភ្នំពេញ, ថ្ងៃទី{dateNumberFormat === 'khmer' ? signatureSolarParts.khmerDay : signatureSolarParts.day} ខែ{signatureSolarParts.month} ឆ្នាំ{dateNumberFormat === 'khmer' ? signatureSolarParts.khmerYear : signatureSolarParts.year}
              </p>

              {/* Role Title in font-moul */}
              <p
                contentEditable
                suppressContentEditableWarning
                onBlur={(e) => setSignerRole(e.currentTarget.textContent || (isSecondary ? 'អ្នកធ្វើតារាង' : 'ប្រធានក្រុមផ្តល់ទិដ្ឋាការ'))}
                style={{
                  fontSize: '12pt',
                  fontFamily: "'Khmer Mool1', 'Khmer Mool', 'Khmer OS Muol Light', 'Khmer OS Muol', 'Moul', serif"
                }}
                className="font-moul text-[12pt] text-slate-900 font-normal outline-none hover:bg-amber-50/50 rounded px-1 mt-1"
              >
                {signerRole}
              </p>

              {/* Signature Graphic / Stamp Area (Clean Blank Space for Physical Signature & Stamp) */}
              <div style={{ height: `${signerGapHeight}px` }} className="my-1 transition-all" />

              {/* Officer Rank (Khmer OS Siemreap) & Name (Khmer OS Muol Light) in official Red font matching Picture */}
              {signerName ? (
                <div
                  style={{ color: nameColor || '#C00000', fontSize: '12pt' }}
                  className="outline-none hover:bg-amber-50/50 rounded px-1 tracking-wide text-[12pt]"
                >
                  {renderSignerFormatted(signerName)}
                </div>
              ) : null}
            </div>
          </div>
        </div>
      </div>

      {/* Yearly Month Checklist & Inspector Modal */}
      {showMonthChecklistModal && (
        <YearlyMonthChecklistModal
          isOpen={showMonthChecklistModal}
          onClose={() => setShowMonthChecklistModal(false)}
          startYear={appliedStartYear}
          startMonth={appliedStartMonth}
          monthColumnsInfo={monthColumnsInfo}
          tableData={tableData as any}
          stockRecords={stockRecords}
          initialMonthIndex={checklistInitialMonth}
          targetTeamName={activeFilteringTeam}
          mode="usage"
          onApplyMonthActualValues={handleApplyMonthActualValues}
        />
      )}

      {/* Yearly Damaged & Missing Visas Checklist Modal */}
      {showDamagedChecklistModal && (
        <YearlyDamagedMissingChecklistModal
          isOpen={showDamagedChecklistModal}
          onClose={() => setShowDamagedChecklistModal(false)}
          startYear={appliedStartYear}
          startMonth={appliedStartMonth}
          monthColumnsInfo={monthColumnsInfo}
          tableData={tableData as any}
          stockRecords={stockRecords}
          targetTeamName={activeFilteringTeam}
          mode="usage"
        />
      )}

      {/* Yearly Received from K2 Checklist Modal */}
      {showReceivedK2ChecklistModal && (
        <YearlyReceivedK2ChecklistModal
          isOpen={showReceivedK2ChecklistModal}
          onClose={() => setShowReceivedK2ChecklistModal(false)}
          startYear={appliedStartYear}
          startMonth={appliedStartMonth}
          monthColumnsInfo={monthColumnsInfo}
          tableData={tableData as any}
          stockRecords={stockRecords}
          targetTeamName={activeFilteringTeam}
          mode="usage"
        />
      )}
    </div>
  );
};
