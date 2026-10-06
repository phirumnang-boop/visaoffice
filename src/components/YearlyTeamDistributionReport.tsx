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
  Edit3,
  Check,
  FileText,
  Save,
  Bold,
  Italic,
  Underline,
  Strikethrough,
  AlignLeft,
  AlignCenter,
  AlignRight,
  AlignJustify,
  Type,
  Maximize2,
  ZoomIn,
  ZoomOut,
  Sliders,
  Sparkles,
  Undo,
  Redo,
  Highlighter,
  Palette,
  RemoveFormatting,
  Search,
  X,
  Minus,
  Plus,
  Layout,
  RefreshCw,
  Award,
  Layers,
  HelpCircle,
  Clock,
  HardDrive,
  ListChecks,
  Eye,
  EyeOff,
} from 'lucide-react';
import { getKhmerLunarDate, getKhmerSolarParts, parseDateInput, toKhmerNum } from '../utils/khmerCalendar';
import { exportSinglePageA4LandscapePdf } from '../utils/pdfExportHelper';
import { printA4Document } from '../utils/printHelper';
import { CustomDatePicker } from './CustomDatePicker';
import { VISA_TYPES, normalizeDateToISO, normalizeVisaType, normalizeTeamName, formatReportTeamName, matchTeamInList, OFFICIAL_29_TEAMS } from '../utils/teamNormalization';
import { isCeaRecord, isOldStockTeamRecord, DEFAULT_OPENING_MATRIX, resolveRecordTeamName, VTR_ID_MAP } from '../utils/teamStockCalculation';
import { TacteingLine, TacteingControlSelector, saveTacteingSettings, getSavedTacteingSettings, TacteingType } from './TacteingLine';
import { idbStorage } from '../utils/idbStorage';
import { YearlyMonthChecklistModal } from './YearlyMonthChecklistModal';
import { YearlyDamagedMissingChecklistModal } from './YearlyDamagedMissingChecklistModal';
import { YearlyReceivedK2ChecklistModal } from './YearlyReceivedK2ChecklistModal';

export interface YearlyTeamDistributionReportProps {
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

export interface RowDistributionData {
  visaType: string;
  openingStock: number; // សន្និធិចុងគ្រា (e.g. 30 Nov 2024)
  monthlyValues: number[]; // 12 months (e.g. Dec, Jan, Feb, ..., Nov)
  totalMonthly: number; // សរុប ១២ខែ
  totalAvailable: number; // ចំនួនសរុប (សន្និធិចុងគ្រា + សរុប ១២ខែ)
  used: number; // ប្រើប្រាស់សរុប
  damagedMissing: number; // ទិដ្ឋាការ ខ្វះ&ខូច
  returnStock: number; // បង្វិលកង
  remaining: number; // សន្និធិនៅសល់ (ចំនួនសរុប - ប្រើប្រាស់ - ខ្វះ&ខូច - បង្វិលកង)
}

// Official Baseline for K2 Office Closing Stock as of 30-Nov-2018
export const OFFICIAL_K2_DEC_2018_BASELINE: Record<string, number> = {
  T: 304000,
  T1: 3200,
  T2: 6200,
  T3: 5200,
  E: 330750,
  E1: 2650,
  E2: 4700,
  E3: 4950,
  D: 3600,
  K: 16750,
  A: 6100,
  B: 7250,
  C: 4850,
};

// Exact official figures for Dec 2018 - Nov 2019 from the verified official document
export const OFFICIAL_2018_2019_TEAM_DISTRIBUTION: Record<string, {
  opening: number;
  months: number[];
  used: number;
  damagedMissing: number;
  returnStock: number;
}> = {
  T: {
    opening: 93675,
    months: [210250, 286250, 252500, 249000, 199250, 146500, 85300, 119250, 165250, 149000, 121500, 146000],
    used: 2118560,
    damagedMissing: 82,
    returnStock: 0,
  },
  T1: {
    opening: 2223,
    months: [500, 50, 1000, 1250, 750, 250, 250, 800, 550, 500, 500, 800],
    used: 6667,
    damagedMissing: 3,
    returnStock: 0,
  },
  T2: {
    opening: 1927,
    months: [0, 0, 0, 550, 0, 0, 200, 50, 0, 0, 250, 250],
    used: 439,
    damagedMissing: 0,
    returnStock: 0,
  },
  T3: {
    opening: 1682,
    months: [250, 250, 0, 550, 0, 0, 200, 50, 0, 0, 250, 0],
    used: 687,
    damagedMissing: 1,
    returnStock: 0,
  },
  E: {
    opening: 37448,
    months: [46500, 54500, 15250, 128300, 104700, 70750, 71500, 87200, 109250, 50000, 24750, 22750],
    used: 769248,
    damagedMissing: 38,
    returnStock: 0,
  },
  E1: {
    opening: 2144,
    months: [500, 1000, 250, 900, 500, 1000, 1000, 2050, 1000, 250, 1250, 1500],
    used: 10228,
    damagedMissing: 4,
    returnStock: 0,
  },
  E2: {
    opening: 2057,
    months: [0, 0, 0, 800, 0, 0, 250, 50, 0, 0, 0, 0],
    used: 540,
    damagedMissing: 4,
    returnStock: 0,
  },
  E3: {
    opening: 1881,
    months: [0, 250, 0, 800, 0, 0, 250, 50, 250, 250, 500, 500],
    used: 1669,
    damagedMissing: 2,
    returnStock: 0,
  },
  D: {
    opening: 1349,
    months: [0, 0, 1000, 1250, 0, 0, 0, 50, 0, 0, 0, 0],
    used: 941,
    damagedMissing: 4,
    returnStock: 0,
  },
  K: {
    opening: 3648,
    months: [3600, 4550, 2000, 2800, 5850, 1500, 0, 6050, 5150, 1100, 2550, 0],
    used: 34135,
    damagedMissing: 99,
    returnStock: 0,
  },
  A: {
    opening: 2060,
    months: [250, 250, 100, 300, 0, 0, 300, 50, 500, 0, 0, 0],
    used: 1291,
    damagedMissing: 1,
    returnStock: 0,
  },
  B: {
    opening: 2512,
    months: [0, 500, 100, 300, 250, 0, 250, 50, 250, 0, 0, 0],
    used: 1305,
    damagedMissing: 3,
    returnStock: 0,
  },
  C: {
    opening: 3500,
    months: [0, 0, 350, 450, 300, 0, 250, 50, 0, 0, 0, 200],
    used: 2486,
    damagedMissing: 3,
    returnStock: 0,
  },
};

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

export const OFFICIAL_2018_2019_DATA = OFFICIAL_2018_2019_TEAM_DISTRIBUTION;
export const OFFICIAL_2024_2025_DATA = OFFICIAL_2018_2019_TEAM_DISTRIBUTION;

// Helper to check if a StockRecord belongs to a specific team (as primary owner or sender)
export const isRecordForTeam = (
  r: StockRecord | any,
  targetTeamName?: string
): boolean => {
  if (!targetTeamName || targetTeamName === 'ផ្នែករដ្ឋបាល' || targetTeamName === 'all' || targetTeamName.toLowerCase() === 'all') {
    return true;
  }
  const targetNorm = normalizeTeamName(targetTeamName);
  if (!targetNorm) return true;

  const rawTeam = resolveRecordTeamName(r);
  const normTeam = normalizeTeamName(rawTeam);
  const matchedOfficial = matchTeamInList(rawTeam, Array.from(OFFICIAL_29_TEAMS));
  const targetOfficial = matchTeamInList(targetTeamName, Array.from(OFFICIAL_29_TEAMS));

  if (rawTeam && rawTeam === targetTeamName) return true;
  if (normTeam && normTeam === targetNorm) return true;
  if (matchedOfficial && targetOfficial && matchedOfficial === targetOfficial) return true;
  if (normTeam && targetNorm && (normTeam.includes(targetNorm) || targetNorm.includes(normTeam))) return true;

  if (r.visaTeamRobokId && normalizeTeamName(r.visaTeamRobokId) === targetNorm) return true;
  if (r.visaTeamRobokName && normalizeTeamName(r.visaTeamRobokName) === targetNorm) return true;
  if (r.sourceFrom && normalizeTeamName(r.sourceFrom) === targetNorm) return true;
  if ((r as any).teamName && normalizeTeamName((r as any).teamName) === targetNorm) return true;

  // Only check destination/issuedTo if record is NOT a team transfer to another recipient
  if (!r.recipientTeamName && !r.recipientTeamId) {
    if ((r as any).issuedTo && normalizeTeamName((r as any).issuedTo) === targetNorm) return true;
    if ((r as any).destinationTo && normalizeTeamName((r as any).destinationTo) === targetNorm) return true;
    if ((r as any).targetTeam && normalizeTeamName((r as any).targetTeam) === targetNorm) return true;
  }

  return false;
};

// Helper to check if a StockRecord belongs to a specific team as RECIPIENT of a transfer
export const isRecordForRecipientTeam = (
  r: StockRecord | any,
  targetTeamName?: string
): boolean => {
  if (!targetTeamName || targetTeamName === 'ផ្នែករដ្ឋបាល' || targetTeamName === 'all' || targetTeamName.toLowerCase() === 'all') {
    return false;
  }
  const targetNorm = normalizeTeamName(targetTeamName);
  if (!targetNorm) return false;

  const targetOfficial = matchTeamInList(targetTeamName, Array.from(OFFICIAL_29_TEAMS));

  const recipName = (r.recipientTeamName || (r as any).destinationTo || (r as any).targetTeam || '').trim();
  const recipId = (r.recipientTeamId || '').trim();

  if (recipName) {
    if (recipName === targetTeamName) return true;
    const recipNorm = normalizeTeamName(recipName);
    if (recipNorm && recipNorm === targetNorm) return true;
    const recipOfficial = matchTeamInList(recipName, Array.from(OFFICIAL_29_TEAMS));
    if (recipOfficial && targetOfficial && recipOfficial === targetOfficial) return true;
    if (recipNorm && targetNorm && (recipNorm.includes(targetNorm) || targetNorm.includes(recipNorm))) return true;
  }

  if (recipId) {
    if (recipId === targetTeamName || recipId.toLowerCase() === targetTeamName.toLowerCase()) return true;
    const idNorm = normalizeTeamName(recipId);
    if (idNorm && idNorm === targetNorm) return true;
    const idOfficial = matchTeamInList(recipId, Array.from(OFFICIAL_29_TEAMS));
    if (idOfficial && targetOfficial && idOfficial === targetOfficial) return true;
    if (VTR_ID_MAP && VTR_ID_MAP[recipId]) {
      const mapped = VTR_ID_MAP[recipId];
      if (mapped === targetTeamName || normalizeTeamName(mapped) === targetNorm) return true;
    }
  }

  return false;
};

// Helper to check if a record is a transfer operation
export const isTransferTeamRecord = (r: StockRecord | any): boolean => {
  const op = (r.operationType || '').trim().toLowerCase();
  const src = (r.sourceFrom || '').trim().toLowerCase();
  const notes = ((r as any).notes || '').trim().toLowerCase();
  return (
    op === 'transferteam' ||
    op === 'transferuseteam' ||
    op === 'transfer' ||
    op.includes('transfer') ||
    op.includes('ផ្ទេរ') ||
    src.includes('ផ្ទេរ') ||
    notes.includes('ផ្ទេរ') ||
    Boolean(r.recipientTeamName || r.recipientTeamId)
  );
};

// Helper to check if a daily operation record belongs to a specific team
export const isDailyRecordForTeam = (
  dRec: any,
  targetTeamName?: string
): boolean => {
  if (!targetTeamName || targetTeamName === 'ផ្នែករដ្ឋបាល' || targetTeamName === 'all') {
    return true;
  }
  const targetNorm = normalizeTeamName(targetTeamName);
  if (!targetNorm) return true;

  const rawTeam = dRec.teamName || dRec.visaTeamRobokName || '';
  const normTeam = normalizeTeamName(rawTeam);
  const matchedOfficial = matchTeamInList(rawTeam, Array.from(OFFICIAL_29_TEAMS));
  const targetOfficial = matchTeamInList(targetTeamName, Array.from(OFFICIAL_29_TEAMS));

  if (rawTeam === targetTeamName) return true;
  if (normTeam === targetNorm) return true;
  if (matchedOfficial && targetOfficial && matchedOfficial === targetOfficial) return true;
  if (normTeam && targetNorm && (normTeam.includes(targetNorm) || targetNorm.includes(normTeam))) return true;

  return false;
};

// Helper to verify if a record is strictly an Issue to Team (ប្រតិបត្តិការ: បើកផ្តល់តាមក្រុម) in Sticker Visa data
export const isIssueTeamRecord = (r: StockRecord): boolean => {
  if (r.stockType && r.stockType !== 'sticker') return false;
  if (isCeaRecord(r)) return false;
  if (isOldStockTeamRecord(r)) return false;
  const op = (r.operationType || '').trim().toLowerCase();

  // Exclude Open K1 / opening central stock from K1
  if (
    op === 'openk1' ||
    op === 'open_k1' ||
    op === 'k1' ||
    op.includes('បើកពីk1') ||
    op.includes('បើកពី k1') ||
    op.includes('openk1')
  ) {
    return false;
  }

  // Exclude used, damaged, return, transfer
  if (
    op === 'useteam' ||
    op === 'used' ||
    op.includes('ប្រើប្រាស់') ||
    op === 'damagedteam' ||
    op === 'damaged' ||
    op === 'returnteam' ||
    op.includes('ខូច') ||
    op.includes('បង្វិល') ||
    op === 'transferteam' ||
    op.includes('ផ្ទេរ')
  ) {
    return false;
  }

  // Strictly issuance to teams (header ប្រតិបត្តិការ: បើកផ្តល់តាមក្រុម)
  return (
    op === 'issueteam' ||
    op === 'issue_team' ||
    op === 'openteam' ||
    op === 'issued' ||
    op.includes('បើកផ្តល់តាមក្រុម') ||
    op.includes('បើកផ្តល់ទៅក្រុម') ||
    op.includes('បើកផ្តល់')
  );
};

// Helper to compute team closing stock for any start date based on Robok Total Stock Work formulas
export const computeTeamStockFromRobokFormula = (
  targetStartDate: string,
  stockRecords: StockRecord[] = [],
  targetTeamName?: string
): Record<string, number> => {
  const normStart = normalizeDateToISO(targetStartDate);
  const isSpecificTeam = Boolean(targetTeamName && targetTeamName !== 'ផ្នែករដ្ឋបាល' && targetTeamName !== 'all');

  let targetTeamNorm = '';
  let targetOfficialTeam: string | null = null;
  if (isSpecificTeam && targetTeamName) {
    targetTeamNorm = normalizeTeamName(targetTeamName);
    targetOfficialTeam = matchTeamInList(targetTeamName, Array.from(OFFICIAL_29_TEAMS));
  }

  // Baseline
  const rows: Record<string, number> = {};
  if (isSpecificTeam && (targetOfficialTeam || targetTeamNorm)) {
    const teamBase =
      (targetOfficialTeam ? DEFAULT_OPENING_MATRIX[targetOfficialTeam] : null) ||
      DEFAULT_OPENING_MATRIX[targetTeamNorm] ||
      DEFAULT_OPENING_MATRIX[targetTeamName!] ||
      {};
    REPORT_VISA_TYPES.forEach((vt) => {
      rows[vt] = teamBase[vt] !== undefined ? teamBase[vt] : 0;
    });
  } else {
    REPORT_VISA_TYPES.forEach((vt) => {
      rows[vt] = (stockRecords && stockRecords.length > 0 ? (OFFICIAL_ROBOK_TEAM_BASELINE[vt] || 0) : 0);
    });
  }

  // Check if there is explicit recorded old stock for this team
  const recordedOldStockMap: Record<string, number> = {};
  let hasRecordedOldStock = false;

  (stockRecords || []).forEach((rec) => {
    if (!isOldStockTeamRecord(rec)) return;
    if (isSpecificTeam && !isRecordForTeam(rec, targetTeamName)) return;

    const vt = normalizeVisaType(rec.visaType);
    if (!rows[vt] && rows[vt] !== 0) return;

    const qty = Number(rec.quantityBundles || rec.totalSheets || (rec as any).quantity || 0);
    recordedOldStockMap[vt] = (recordedOldStockMap[vt] || 0) + qty;
    hasRecordedOldStock = true;
  });

  if (hasRecordedOldStock) {
    REPORT_VISA_TYPES.forEach((vt) => {
      if (recordedOldStockMap[vt] !== undefined) {
        rows[vt] = recordedOldStockMap[vt];
      }
    });
  }

  (stockRecords || []).forEach((rec) => {
    if (rec.stockType && rec.stockType !== 'sticker') return;
    if (isCeaRecord(rec)) return;
    if (isOldStockTeamRecord(rec)) return;

    const isTransfer = isTransferTeamRecord(rec);
    const isSender = !isSpecificTeam || isRecordForTeam(rec, targetTeamName);
    const isRecipient = isSpecificTeam && isRecordForRecipientTeam(rec, targetTeamName);

    if (isSpecificTeam) {
      if (isTransfer) {
        if (!isSender && !isRecipient) return;
      } else {
        if (!isSender) return;
      }
    }

    const vt = normalizeVisaType(rec.visaType);
    if (!rows[vt] && rows[vt] !== 0) return;

    const qty = Number(rec.quantityBundles || rec.totalSheets || (rec as any).quantity || 0);
    const recDate = normalizeDateToISO(rec.date || '');
    const op = (rec.operationType || '').toLowerCase();

    if (!recDate || recDate <= '2018-11-30') return;

    const isDamagedK2 =
      op === 'damagedk2' ||
      op === 'damaged' ||
      op === 'damagedoffice' ||
      op === 'voidk2' ||
      op === 'voidoffice' ||
      op.includes('ខូចក២') ||
      op.includes('មិនបានការក២') ||
      op.includes('ខូចក') ||
      op.includes('ក២ខូច');

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
          (rec.sourceFrom?.includes('ក្រុម') || (rec as any).visaTeamRobokName)));

    if (normStart && recDate < normStart) {
      if (isTransfer) {
        if (isRecipient) {
          rows[vt] += qty; // Received transfer before start date increases opening stock
        } else if (isSender) {
          rows[vt] -= qty; // Sent transfer before start date decreases opening stock
        }
      } else if (op === 'returnteam' || op === 'returnoffice' || op.includes('បង្វិល')) {
        rows[vt] -= qty;
      } else if (isIssueTeamRecord(rec)) {
        rows[vt] += qty;
      } else if (op === 'useteam' || op.includes('ប្រើប្រាស់')) {
        rows[vt] -= qty;
      } else if (isDamagedTeam) {
        rows[vt] -= qty;
      }
    }
  });

  // Daily operations
  try {
    const savedDaily1 = localStorage.getItem('app_daily_team_records_v1');
    const savedDaily5 = localStorage.getItem('app_daily_team_operations_v5');
    const allDailyRecords: any[] = [];
    if (savedDaily1) {
      try {
        const p1 = JSON.parse(savedDaily1);
        if (Array.isArray(p1)) allDailyRecords.push(...p1);
      } catch {}
    }
    if (savedDaily5) {
      try {
        const p5 = JSON.parse(savedDaily5);
        if (Array.isArray(p5)) allDailyRecords.push(...p5);
      } catch {}
    }

    if (allDailyRecords.length > 0) {
      const existingUseKeySet = new Set<string>();
      (stockRecords || []).forEach((sr) => {
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

      const processedDailyIds = new Set<string>();

      allDailyRecords.forEach((dRec: any) => {
        if (!dRec || !dRec.id || processedDailyIds.has(dRec.id)) return;
        processedDailyIds.add(dRec.id);

        if (dRec.categoryType === 'cEA' || (dRec.categoryType && dRec.categoryType !== 'Sticker')) return;
        const recDate = normalizeDateToISO(dRec.date || '');
        if (!dRec.values || !recDate || recDate <= '2018-11-30') return;

        if (isSpecificTeam && !isDailyRecordForTeam(dRec, targetTeamName)) return;

        const rawTeamName = dRec.teamName || dRec.visaTeamRobokName || '';
        const normTeam = normalizeTeamName(rawTeamName);

        Object.keys(rows).forEach((vt) => {
          const item = dRec.values[vt];
          let qty = 0;
          if (item) {
            if (Array.isArray(item.entries) && item.entries.length > 0) {
              qty = item.entries.reduce((sum: number, e: any) => sum + (parseInt(e?.quantity || '', 10) || 0), 0);
            } else {
              qty = parseInt(item?.quantity || item || '', 10) || 0;
            }
          }
          if (qty <= 0) return;

          const lookupKey = `${recDate}_${normTeam}_${vt}`;
          if (!existingUseKeySet.has(lookupKey)) {
            if (normStart && recDate < normStart) {
              rows[vt] -= qty;
            }
          }
        });
      });
    }
  } catch (e) {
    console.error('Error computing team stock from robok:', e);
  }

  return rows;
};

// Helper to check if a record is K2 receiving from K1 (បើកពីក១)
export const isReceiveFromK1Record = (r: StockRecord): boolean => {
  if (r.stockType && r.stockType !== 'sticker') return false;
  if (isCeaRecord(r)) return false;
  if (isOldStockTeamRecord(r)) return false;
  const op = (r.operationType || '').trim().toLowerCase();

  // Exclude operations that are issuance/usage/damage
  if (
    op === 'issueteam' ||
    op === 'issue_team' ||
    op === 'useteam' ||
    op === 'use_team' ||
    op === 'used' ||
    op === 'damaged' ||
    op === 'damagedk2' ||
    op === 'damagedteam' ||
    op.includes('បើកផ្តល់') ||
    op.includes('បើកផ្ដល់') ||
    op.includes('ប្រើប្រាស់') ||
    op.includes('ខូច')
  ) {
    return false;
  }

  return (
    op === 'openk1' ||
    op === 'open_k1' ||
    op === 'k1' ||
    op === 'receive_from_k1' ||
    op === 'receivek1' ||
    op.includes('បើកពីk1') ||
    op.includes('បើកពី k1') ||
    op.includes('បើកពីក១') ||
    op.includes('បើកពី ក១') ||
    op.includes('ទទួលពីក១') ||
    op.includes('ទទួលពី k1') ||
    op.includes('openk1') ||
    (Boolean(r.sourceFrom) && (r.sourceFrom?.includes('ក១') || r.sourceFrom?.includes('K1')))
  );
};

// Helper to check if a record is K2 damaged (ទិដ្ឋាការខូចក២)
export const isDamagedK2Record = (r: StockRecord): boolean => {
  if (r.stockType && r.stockType !== 'sticker') return false;
  if (isCeaRecord(r)) return false;
  if (isOldStockTeamRecord(r)) return false;
  const op = (r.operationType || '').trim().toLowerCase();

  return (
    op === 'damagedk2' ||
    op === 'damaged_k2' ||
    op === 'damaged' ||
    op === 'damagedoffice' ||
    op === 'voidk2' ||
    op === 'voidoffice' ||
    op.includes('ខូចក២') ||
    op.includes('មិនបានការក២') ||
    op.includes('ក២ខូច') ||
    op.includes('ខូចការិយាល័យ')
  );
};

// Compute opening stock for K2 office
export const computeK2OfficeOpeningStock = (
  targetStartDate: string,
  stockRecords: StockRecord[] = []
): Record<string, number> => {
  const normStart = normalizeDateToISO(targetStartDate);
  const rows: Record<string, number> = {};

  REPORT_VISA_TYPES.forEach((vt) => {
    rows[vt] = OFFICIAL_K2_DEC_2018_BASELINE[vt] || 0;
  });

  (stockRecords || []).forEach((rec) => {
    if (rec.stockType && rec.stockType !== 'sticker') return;
    if (isCeaRecord(rec)) return;
    if (isOldStockTeamRecord(rec)) return;

    const vt = normalizeVisaType(rec.visaType);
    if (!rows[vt] && rows[vt] !== 0) return;

    const qty = Number(rec.quantityBundles || rec.totalSheets || (rec as any).quantity || 0);
    const recDate = normalizeDateToISO(rec.date || '');
    const op = (rec.operationType || '').toLowerCase();

    if (!recDate || recDate <= '2018-11-30') return;

    if (normStart && recDate < normStart) {
      if (isReceiveFromK1Record(rec)) {
        rows[vt] += qty;
      } else if (isIssueTeamRecord(rec)) {
        rows[vt] -= qty;
      } else if (isDamagedK2Record(rec)) {
        rows[vt] -= qty;
      } else if (op === 'returnk1' || op === 'return_k1' || op.includes('បង្វិលក១')) {
        rows[vt] -= qty;
      }
    }
  });

  return rows;
};

// Full calculation of the 12-month team distribution report for ANY given start year & start month
export const computeYearlyDistributionReportData = (
  startYear: number,
  startMonth: number,
  stockRecords: StockRecord[] = [],
  targetTeamName?: string
): Record<string, RowDistributionData> => {
  const isSpecificTeam = Boolean(targetTeamName && targetTeamName !== 'ផ្នែករដ្ឋបាល' && targetTeamName !== 'all');

  // Build 12 columns info
  const cols: { monthNum: number; year: number }[] = [];
  for (let i = 0; i < 12; i++) {
    const totalOffset = (startMonth - 1) + i;
    const curYear = startYear + Math.floor(totalOffset / 12);
    const curMonth = (totalOffset % 12) + 1;
    cols.push({ monthNum: curMonth, year: curYear });
  }

  const startMStr = String(startMonth).padStart(2, '0');
  const startDateStr = `${startYear}-${startMStr}-01`;
  const lastCol = cols[11];
  const daysInEndMonth = new Date(lastCol.year, lastCol.monthNum, 0).getDate();
  const endDateStr = `${lastCol.year}-${String(lastCol.monthNum).padStart(2, '0')}-${String(daysInEndMonth).padStart(2, '0')}`;

  const computedOpening = computeTeamStockFromRobokFormula(startDateStr, stockRecords, targetTeamName);

  const result: Record<string, RowDistributionData> = {};
  REPORT_VISA_TYPES.forEach((vt) => {
    result[vt] = {
      visaType: vt,
      openingStock: computedOpening[vt] ?? (isSpecificTeam ? 0 : (stockRecords && stockRecords.length > 0 ? (OFFICIAL_ROBOK_TEAM_BASELINE[vt] ?? 0) : 0)),
      monthlyValues: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
      totalMonthly: 0,
      totalAvailable: 0,
      used: 0,
      damagedMissing: 0,
      returnStock: 0,
      remaining: 0,
    };
  });

  const cleanRecords = (stockRecords || []).filter(
    (r) => (r.stockType === 'sticker' || !r.stockType) && !isCeaRecord(r)
  );

  cleanRecords.forEach((r) => {
    if (isCeaRecord(r)) return;

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

    const vt = normalizeVisaType(r.visaType || '');
    if (!result[vt]) return;

    const dateStr = normalizeDateToISO(r.date || (r as any).createdAt || '');
    if (!dateStr) return;

    // Strict date bounds: ANY record outside [startDateStr, endDateStr] MUST NOT be included in this 12-month report
    if (dateStr < startDateStr || dateStr > endDateStr) return;

    const [yrStr, moStr] = dateStr.split('-');
    const recYear = parseInt(yrStr, 10);
    const recMonth = parseInt(moStr, 10);
    const op = (r.operationType || '').toLowerCase();

    const matchIdx = cols.findIndex(
      (c) => c.year === recYear && c.monthNum === recMonth
    );

    const qty = Number(r.quantityBundles || r.totalSheets || (r as any).quantity || ((r as any).count ? (r as any).count * 50 : 0));
    if (!qty) return;

    const isDamagedK2 =
      op === 'damagedk2' ||
      op === 'damaged' ||
      op === 'damagedoffice' ||
      op === 'voidk2' ||
      op === 'voidoffice' ||
      op.includes('ខូចក២') ||
      op.includes('មិនបានការក២') ||
      op.includes('ខូចក') ||
      op.includes('ក២ខូច');

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

    // Monthly Distribution to teams (strictly issuance to teams from header ប្រតិបត្តិការ: បើកផ្តល់តាមក្រុម only in Sticker Visa)
    if (isIssueTeamRecord(r)) {
      if (matchIdx !== -1 && !isDamagedTeam && !isTransfer && op !== 'useteam') {
        result[vt].monthlyValues[matchIdx] += qty;
      }
    }

    if (isTransfer && matchIdx !== -1) {
      if (isSpecificTeam) {
        // ONLY count transfers in monthly received values for the RECIPIENT team!
        if (isRecipient) {
          result[vt].monthlyValues[matchIdx] += qty;
        }
      } else {
        // Ignore team transfers for office distribution calculations
      }
    }

    // Usage records in the 12-month period
    if (op === 'useteam' || op === 'used' || op.includes('ប្រើប្រាស់')) {
      if (dateStr >= startDateStr && dateStr <= endDateStr) {
        result[vt].used += qty;
      }
    }

    // Missing & Damaged in the 12-month period
    if (isDamagedTeam) {
      if (dateStr >= startDateStr && dateStr <= endDateStr) {
        result[vt].damagedMissing += qty;
      }
    }

    if (isTransfer) {
      if (isSpecificTeam) {
        // ONLY the SENDER team has this deducted as transferred out under "ផ្ទេរ/ខ្វះ/ខូច"
        if (isSender && !isRecipient) {
          if (dateStr >= startDateStr && dateStr <= endDateStr) {
            result[vt].damagedMissing += qty;
          }
        }
      }
    }

    // Return to Stock in the 12-month period
    if (op === 'returnteam' || op === 'returnoffice' || op === 'recycle' || op.includes('បង្វិល')) {
      if (dateStr >= startDateStr && dateStr <= endDateStr) {
        result[vt].returnStock += qty;
      }
    }
  });

  // Daily operations for Usage
  try {
    const savedDaily1 = localStorage.getItem('app_daily_team_records_v1');
    const savedDaily5 = localStorage.getItem('app_daily_team_operations_v5');
    const allDailyRecords: any[] = [];
    if (savedDaily1) {
      try {
        const p1 = JSON.parse(savedDaily1);
        if (Array.isArray(p1)) allDailyRecords.push(...p1);
      } catch {}
    }
    if (savedDaily5) {
      try {
        const p5 = JSON.parse(savedDaily5);
        if (Array.isArray(p5)) allDailyRecords.push(...p5);
      } catch {}
    }

    if (allDailyRecords.length > 0) {
      const existingUseKeySet = new Set<string>();
      (stockRecords || []).forEach((sr) => {
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

      const processedDailyIds = new Set<string>();

      allDailyRecords.forEach((dRec: any) => {
        if (!dRec || !dRec.id || processedDailyIds.has(dRec.id)) return;
        processedDailyIds.add(dRec.id);

        if (dRec.categoryType === 'cEA' || (dRec.categoryType && dRec.categoryType !== 'Sticker')) return;
        const recDate = normalizeDateToISO(dRec.date || '');
        if (!dRec.values || !recDate) return;
        // Strictly within period
        if (recDate < startDateStr || recDate > endDateStr) return;

        if (isSpecificTeam && !isDailyRecordForTeam(dRec, targetTeamName)) return;

        const rawTeamName = dRec.teamName || dRec.visaTeamRobokName || '';
        const normTeam = normalizeTeamName(rawTeamName);

        REPORT_VISA_TYPES.forEach((vt) => {
          const item = dRec.values[vt];
          let qty = 0;
          if (item) {
            if (Array.isArray(item.entries) && item.entries.length > 0) {
              qty = item.entries.reduce((sum: number, e: any) => sum + (parseInt(e?.quantity || '', 10) || 0), 0);
            } else {
              qty = parseInt(item?.quantity || item || '', 10) || 0;
            }
          }
          if (qty <= 0) return;

          const lookupKey = `${recDate}_${normTeam}_${vt}`;
          if (!existingUseKeySet.has(lookupKey)) {
            result[vt].used += qty;
          }
        });
      });
    }
  } catch (e) {
    console.error('Error computing daily usage for yearly report:', e);
  }

  // No fallback to official sample to keep table empty/0 as requested by the user when no data is uploaded yet

  REPORT_VISA_TYPES.forEach((vt) => {
    result[vt].totalMonthly = result[vt].monthlyValues.reduce((a, b) => (a || 0) + (b || 0), 0);
    result[vt].totalAvailable = (result[vt].openingStock || 0) + result[vt].totalMonthly;
    result[vt].remaining =
      result[vt].totalAvailable -
      (result[vt].used || 0) -
      (result[vt].damagedMissing || 0) -
      (result[vt].returnStock || 0);
  });

  return result;
};

export const YearlyTeamDistributionReport: React.FC<YearlyTeamDistributionReportProps> = ({
  stockRecords,
  categories,
  officers = [],
  currentRole,
  userName,
  assignedTeam,
  onClose,
}) => {
  const { scaleMode } = useWorkspaceSettings();
  // Fiscal / Period Configuration (Default Dec 2018 to Nov 2019)
  const [startYear, setStartYear] = useState<number>(2018);
  const [startMonth, setStartMonth] = useState<number>(12); // Month index 1-12 (12 = Dec)
  const [appliedStartYear, setAppliedStartYear] = useState<number>(2018);
  const [appliedStartMonth, setAppliedStartMonth] = useState<number>(12);
  const [isCalculating, setIsCalculating] = useState<boolean>(false);
  const [hasCalculated, setHasCalculated] = useState<boolean>(false);
  const [reportDate, setReportDate] = useState<string>('2019-12-08'); // Signature date

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

  // Header and Metadata Customization
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
    return sectionName && sectionName !== 'ផ្នែករដ្ឋបាល' ? sectionName : (defaultFullTeamName !== 'ផ្នែករដ្ឋបាល' ? defaultFullTeamName : undefined);
  }, [currentRole, sectionName, defaultFullTeamName]);

  const isSecondary = currentRole === 'Secondary';

  // Signer Customization
  const [signerRole, setSignerRole] = useState<string>(() => (isSecondary ? 'អ្នកធ្វើតារាង' : 'ប្រធានក្រុមផ្តល់ទិដ្ឋាការ'));
  const [signerName, setSignerName] = useState<string>(() => (isSecondary ? 'អនុសេនីយ៍ឯក ច្រេង ថុល' : ''));

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

  // Auto calculation vs custom manual override mode
  const [autoCalculate, setAutoCalculate] = useState<boolean>(true);

  // Primary Table State: dictionary of row data by visa type (starts clean with 0s before user calculates)
  const [tableData, setTableData] = useState<Record<string, RowDistributionData>>(() => {
    const initial: Record<string, RowDistributionData> = {};
    REPORT_VISA_TYPES.forEach((vt) => {
      initial[vt] = {
        visaType: vt,
        openingStock: 0,
        monthlyValues: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
        totalMonthly: 0,
        totalAvailable: 0,
        used: 0,
        damagedMissing: 0,
        returnStock: 0,
        remaining: 0,
      };
    });
    return initial;
  });

  // Calculate 12 Month Column Headers dynamically based on appliedStartMonth and appliedStartYear
  const monthColumnsInfo = useMemo(() => {
    const cols: { monthNum: number; year: number; name: string }[] = [];
    for (let i = 0; i < 12; i++) {
      const totalOffset = (appliedStartMonth - 1) + i;
      const curYear = appliedStartYear + Math.floor(totalOffset / 12);
      const curMonth = (totalOffset % 12) + 1; // 1-12
      const name = KHMER_MONTHS_NAMES[curMonth - 1];
      cols.push({
        monthNum: curMonth,
        year: curYear,
        name,
      });
    }
    return cols;
  }, [appliedStartYear, appliedStartMonth]);

  // Derived Start and End Dates based on applied period
  const startDateStr = useMemo(() => {
    const mStr = String(appliedStartMonth).padStart(2, '0');
    return `${appliedStartYear}-${mStr}-01`;
  }, [appliedStartYear, appliedStartMonth]);

  const endDateStr = useMemo(() => {
    const lastCol = monthColumnsInfo[11];
    const daysInEndMonth = new Date(lastCol.year, lastCol.monthNum, 0).getDate();
    const mStr = String(lastCol.monthNum).padStart(2, '0');
    return `${lastCol.year}-${mStr}-${String(daysInEndMonth).padStart(2, '0')}`;
  }, [monthColumnsInfo]);

  // Previous Period Date for Opening Stock Label (e.g. "៣០ វិច្ឆិកា ២០២៤")
  const openingStockLabel = useMemo(() => {
    let prevM = appliedStartMonth - 1;
    let prevY = appliedStartYear;
    if (prevM < 1) {
      prevM = 12;
      prevY -= 1;
    }
    const daysInPrev = new Date(prevY, prevM, 0).getDate();
    const khmerDay = toKhmerNum(daysInPrev);
    const khmerMonth = KHMER_MONTHS_NAMES[prevM - 1];
    const khmerYear = toKhmerNum(prevY);
    return `${khmerDay} ${khmerMonth} ${khmerYear}`;
  }, [appliedStartYear, appliedStartMonth]);

  // Ending Period Date Label for Remaining Stock (e.g. "៣០ វិច្ឆិកា ២០២៥")
  const endingStockLabel = useMemo(() => {
    const lastCol = monthColumnsInfo[11];
    const daysInEndMonth = new Date(lastCol.year, lastCol.monthNum, 0).getDate();
    const khmerDay = toKhmerNum(daysInEndMonth);
    const khmerMonth = lastCol.name;
    const khmerYear = toKhmerNum(lastCol.year);
    return `${khmerDay} ${khmerMonth} ${khmerYear}`;
  }, [monthColumnsInfo]);

  // Lunar and Solar Date for Signatures
  const signatureSolarParts = useMemo(() => getKhmerSolarParts(reportDate), [reportDate]);
  const signatureLunarDate = useMemo(() => getKhmerLunarDate(reportDate), [reportDate]);

  // Editable Report Title
  const [reportTitle, setReportTitle] = useState<string>(
    'តារាងសន្លឹកទិដ្ឋាការស្អិត តាមប្រភេទនីមួយៗ បើកពីការិយាល័យទិដ្ឋាការចូល'
  );

  // Header Font Size (pt) - Standard 12pt as requested
  const [headerFontSize, setHeaderFontSize] = useState<number>(() => {
    try {
      const saved = localStorage.getItem('yearly_team_dist_header_font_size_pt');
      if (saved !== null) {
        const val = parseFloat(saved);
        if (!isNaN(val) && val >= 8 && val <= 24) return val;
      }
    } catch {}
    return 12;
  });

  useEffect(() => {
    try {
      localStorage.setItem('yearly_team_dist_header_font_size_pt', headerFontSize.toString());
    } catch {}
  }, [headerFontSize]);

  // Signature Vertical Shift (pt) - Shifted up as requested
  const [showFormattingToolbar, setShowFormattingToolbar] = useState<boolean>(() => {
    try {
      const saved = localStorage.getItem('yearly_team_dist_show_formatting_toolbar');
      if (saved !== null) return saved === 'true';
    } catch {}
    return true;
  });

  useEffect(() => {
    try {
      localStorage.setItem('yearly_team_dist_show_formatting_toolbar', showFormattingToolbar.toString());
    } catch {}
  }, [showFormattingToolbar]);

  const [signatureShiftY, setSignatureShiftY] = useState<number>(() => {
    try {
      const saved = localStorage.getItem('yearly_team_dist_signature_shift_y');
      if (saved !== null) {
        const val = parseFloat(saved);
        if (!isNaN(val) && val >= -40 && val <= 60) return val;
      }
    } catch {}
    return 2; // Shifted up (default 2pt instead of previous ~18pt mt-6)
  });

  useEffect(() => {
    try {
      localStorage.setItem('yearly_team_dist_signature_shift_y', signatureShiftY.toString());
    } catch {}
  }, [signatureShiftY]);

  // Spacing between signer title (អ្នកធ្វើតារាង) and signer rank/name (អនុសេនីយ៍ឯក ច្រេង ថុល)
  const [signerGapHeight, setSignerGapHeight] = useState<number>(() => {
    try {
      const saved = localStorage.getItem('yearly_team_dist_signer_gap_height');
      if (saved !== null) {
        const val = parseInt(saved, 10);
        if (!isNaN(val) && val >= 20 && val <= 200) return val;
      }
    } catch {}
    return 96; // Increased default spacing (from 64px to 96px)
  });

  useEffect(() => {
    try {
      localStorage.setItem('yearly_team_dist_signer_gap_height', signerGapHeight.toString());
    } catch {}
  }, [signerGapHeight]);

  // Formatting and UI Tools State
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
      const saved = localStorage.getItem('yearly_team_distribution_row_height');
      if (saved !== null) {
        const parsed = parseInt(saved, 10);
        if (!isNaN(parsed) && parsed >= 14 && parsed <= 30) return parsed;
      }
    } catch {}
    return 19;
  });
  const [lineSpacing, setLineSpacing] = useState<number>(1.25);
  const [isBold, setIsBold] = useState<boolean>(false);
  const [isItalic, setIsItalic] = useState<boolean>(false);
  const [isUnderline, setIsUnderline] = useState<boolean>(false);
  const [textColor, setTextColor] = useState<string>('#000000');
  const [nameColor, setNameColor] = useState<string>('#C00000');
  const [highlightColor, setHighlightColor] = useState<string>('transparent');
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

  const [isGeneratingPdf, setIsGeneratingPdf] = useState<boolean>(false);
  const [isSavedRecently, setIsSavedRecently] = useState<boolean>(false);
  const [activePreset, setActivePreset] = useState<'official2024' | 'autoRealStock' | 'custom'>('official2024');
  const [showMonthChecklistModal, setShowMonthChecklistModal] = useState<boolean>(false);
  const [checklistInitialMonth, setChecklistInitialMonth] = useState<number>(0);
  const [showDamagedChecklistModal, setShowDamagedChecklistModal] = useState<boolean>(false);
  const [showReceivedK2ChecklistModal, setShowReceivedK2ChecklistModal] = useState<boolean>(false);

  const printAreaRef = useRef<HTMLDivElement>(null);

  // Auto-Save Draft to IndexedDB & LocalStorage
  const handleSaveDraft = () => {
    try {
      localStorage.setItem('app_yearly_team_distribution_draft_v6', JSON.stringify(tableData));
      idbStorage.setItem('yearly_team_distribution_report_draft_v6', 'data', {
        tableData,
        startYear,
        startMonth,
        reportDate,
        signerName,
        signerRole,
        dateNumberFormat,
      });
      setIsSavedRecently(true);
      setTimeout(() => setIsSavedRecently(false), 2500);
    } catch (e) {
      console.warn('Could not save draft:', e);
    }
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
          const totalAvailable = updated[vt].openingStock + totalMonthly;
          const remaining = totalAvailable - updated[vt].used - updated[vt].damagedMissing - updated[vt].returnStock;
          updated[vt] = {
            ...updated[vt],
            monthlyValues: nextMonthly,
            totalMonthly,
            totalAvailable,
            remaining,
          };
        }
      });
      return updated;
    });
  };

  // Cell Edit Handler
  const handleCellChange = (
    visaType: string,
    field: 'openingStock' | 'month' | 'used' | 'damagedMissing' | 'returnStock' | 'remaining' | 'totalMonthly' | 'totalAvailable',
    value: string,
    monthIndex?: number
  ) => {
    const cleaned = value.replace(/,/g, '').trim();
    const num = cleaned === '' ? 0 : parseInt(cleaned, 10);
    if (isNaN(num)) return;

    setTableData((prev) => {
      const current = prev[visaType] || {
        visaType,
        openingStock: 0,
        monthlyValues: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
        totalMonthly: 0,
        totalAvailable: 0,
        used: 0,
        damagedMissing: 0,
        returnStock: 0,
        remaining: 0,
      };

      const updatedRow = { ...current, monthlyValues: [...current.monthlyValues] };

      if (field === 'openingStock') {
        updatedRow.openingStock = num;
      } else if (field === 'month' && monthIndex !== undefined) {
        updatedRow.monthlyValues[monthIndex] = num;
      } else if (field === 'used') {
        updatedRow.used = num;
      } else if (field === 'damagedMissing') {
        updatedRow.damagedMissing = num;
      } else if (field === 'returnStock') {
        updatedRow.returnStock = num;
      } else if (field === 'totalMonthly') {
        updatedRow.totalMonthly = num;
      } else if (field === 'totalAvailable') {
        updatedRow.totalAvailable = num;
      } else if (field === 'remaining') {
        updatedRow.remaining = num;
      }

      // Auto-recalculate row sums if autoCalculate is ON
      if (autoCalculate) {
        updatedRow.totalMonthly = updatedRow.monthlyValues.reduce((a, b) => a + b, 0);
        updatedRow.totalAvailable = updatedRow.openingStock + updatedRow.totalMonthly;
        updatedRow.remaining =
          updatedRow.totalAvailable - updatedRow.used - updatedRow.damagedMissing - updatedRow.returnStock;
      }

      return {
        ...prev,
        [visaType]: updatedRow,
      };
    });
  };

  // Grand Totals Calculation (Columns sum across all visa types)
  const grandTotals = useMemo(() => {
    let totalOpening = 0;
    const totalMonths: number[] = [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
    let totalMonthlySum = 0;
    let totalAvailableSum = 0;
    let totalUsed = 0;
    let totalDamagedMissing = 0;
    let totalReturnStock = 0;
    let totalRemaining = 0;

    REPORT_VISA_TYPES.forEach((vt) => {
      const row = tableData[vt];
      if (row) {
        const opening = Number(row.openingStock) || 0;
        const used = Number(row.used) || 0;
        const damagedMissing = Number(row.damagedMissing) || 0;
        const returnStock = Number(row.returnStock) || 0;

        totalOpening += opening;
        let rMonthlySum = 0;
        row.monthlyValues.forEach((val, idx) => {
          const mVal = Number(val) || 0;
          totalMonths[idx] = (totalMonths[idx] || 0) + mVal;
          rMonthlySum += mVal;
        });

        const effectiveMonthly = autoCalculate ? rMonthlySum : (row.totalMonthly ?? rMonthlySum);
        const rAvailable = autoCalculate ? opening + effectiveMonthly : (row.totalAvailable ?? (opening + effectiveMonthly));
        const rRemaining = autoCalculate
          ? rAvailable - used - damagedMissing - returnStock
          : (row.remaining ?? (rAvailable - used - damagedMissing - returnStock));

        totalMonthlySum += effectiveMonthly;
        totalAvailableSum += rAvailable;
        totalUsed += used;
        totalDamagedMissing += damagedMissing;
        totalReturnStock += returnStock;
        totalRemaining += rRemaining;
      }
    });

    return {
      opening: totalOpening,
      months: totalMonths,
      monthlySum: totalMonthlySum,
      available: totalAvailableSum,
      used: totalUsed,
      damagedMissing: totalDamagedMissing,
      returnStock: totalReturnStock,
      remaining: totalRemaining,
    };
  }, [tableData, autoCalculate]);

  // Note: Data is initially uncalculated (showing '-') until user selects date and clicks "បង្ហាញទិន្នន័យ"

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
      const computed = computeYearlyDistributionReportData(y, m, stockRecords, activeFilteringTeam);
      setTableData(computed);
      setIsCalculating(false);
      setHasCalculated(true);
    }, 50);
  };

  // Load Official 2018-2019 Template Preset
  const handleLoadOfficial2018 = () => {
    setStartYear(2018);
    setStartMonth(12);
    setAppliedStartYear(2018);
    setAppliedStartMonth(12);
    setReportDate('2019-12-08');
    setSignerRole(isSecondary ? 'អ្នកធ្វើតារាង' : 'ប្រធានក្រុមផ្តល់ទិដ្ឋាការ');
    setSignerName(isSecondary ? 'អនុសេនីយ៍ឯក អ៊ុក រ័ត្នបញ្ញា' : '');

    if (activeFilteringTeam && activeFilteringTeam !== 'ផ្នែករដ្ឋបាល') {
      const computed = computeYearlyDistributionReportData(2018, 12, stockRecords, activeFilteringTeam);
      setTableData(computed);
      setActivePreset('autoRealStock');
      setHasCalculated(true);
      return;
    }

    const nextData: Record<string, RowDistributionData> = {};
    REPORT_VISA_TYPES.forEach((vt) => {
      const template = OFFICIAL_2018_2019_DATA[vt] || {
        opening: 0,
        months: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
        used: 0,
        damagedMissing: 0,
        returnStock: 0,
      };
      const totalMonthly = template.months.reduce((a, b) => a + b, 0);
      const totalAvailable = template.opening + totalMonthly;
      const remaining = totalAvailable - template.used - template.damagedMissing - template.returnStock;

      nextData[vt] = {
        visaType: vt,
        openingStock: template.opening,
        monthlyValues: [...template.months],
        totalMonthly,
        totalAvailable,
        used: template.used,
        damagedMissing: template.damagedMissing,
        returnStock: template.returnStock,
        remaining,
      };
    });
    setTableData(nextData);
    setActivePreset('official2018');
    setHasCalculated(true);
    localStorage.setItem('app_yearly_team_distribution_draft_v6', JSON.stringify(nextData));
  };

  const handleLoadOfficial2024 = () => {
    handleSelectPeriodPreset(2024, 12, 'official2024');
  };

  // Preset Selectors for fast period switching
  const handleSelectPeriodPreset = (presetYear: number, presetMonth: number, presetTag: string) => {
    setStartYear(presetYear);
    setStartMonth(presetMonth);
    setAppliedStartYear(presetYear);
    setAppliedStartMonth(presetMonth);
    setActivePreset(presetTag as any);
    
    // Auto set appropriate report signing date
    const endY = presetMonth === 1 ? presetYear : presetYear + 1;
    const endM = presetMonth === 1 ? '12' : String(presetMonth - 1).padStart(2, '0');
    setReportDate(`${endY}-${endM}-08`);

    const computed = computeYearlyDistributionReportData(presetYear, presetMonth, stockRecords, activeFilteringTeam);
    setTableData(computed);
    setHasCalculated(true);
  };

  // Auto-Pull & Compute from Real System Stock Records
  const handleLoadFromRealStock = () => {
    const computed = computeYearlyDistributionReportData(appliedStartYear, appliedStartMonth, stockRecords, activeFilteringTeam);
    setTableData(computed);
    setActivePreset('autoRealStock');
  };

  // Dedicated action to load Team Closing Stock from Robok Total Stock Work Report (បក.សរុបការងារស្តុក)
  const handleLoadFromRobokTotalStock = () => {
    const computedRobokTeamOpening = computeTeamStockFromRobokFormula(startDateStr, stockRecords, activeFilteringTeam);
    setTableData((prev) => {
      const next: Record<string, RowDistributionData> = {};
      REPORT_VISA_TYPES.forEach((vt) => {
        const r = prev[vt] || {
          visaType: vt,
          openingStock: 0,
          monthlyValues: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
          totalMonthly: 0,
          totalAvailable: 0,
          used: 0,
          damagedMissing: 0,
          returnStock: 0,
          remaining: 0,
        };
        const newOpening = computedRobokTeamOpening[vt] ?? (activeFilteringTeam ? 0 : (stockRecords && stockRecords.length > 0 ? (OFFICIAL_ROBOK_TEAM_BASELINE[vt] ?? 0) : 0));
        const totalMonthly = r.monthlyValues.reduce((a, b) => a + b, 0);
        const totalAvailable = newOpening + totalMonthly;
        const remaining = totalAvailable - (r.used || 0) - (r.damagedMissing || 0) - (r.returnStock || 0);

        next[vt] = {
          ...r,
          openingStock: newOpening,
          totalMonthly,
          totalAvailable,
          remaining,
        };
      });
      return next;
    });
    setActivePreset('autoRobok');
  };

  // Re-Calculate all totals on demand
  const handleRecalculateAll = () => {
    setTableData((prev) => {
      const next: Record<string, RowDistributionData> = {};
      REPORT_VISA_TYPES.forEach((vt) => {
        const r = prev[vt] || {
          visaType: vt,
          openingStock: 0,
          monthlyValues: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
          totalMonthly: 0,
          totalAvailable: 0,
          used: 0,
          damagedMissing: 0,
          returnStock: 0,
          remaining: 0,
        };
        const totalMonthly = r.monthlyValues.reduce((a, b) => a + b, 0);
        const totalAvailable = (r.openingStock || 0) + totalMonthly;
        const remaining = totalAvailable - (r.used || 0) - (r.damagedMissing || 0) - (r.returnStock || 0);

        next[vt] = {
          ...r,
          totalMonthly,
          totalAvailable,
          remaining,
        };
      });
      return next;
    });
  };

  // Reset to Zero
  const handleClearAll = () => {
    if (!window.confirm('តើលោកអ្នកពិតជាចង់សម្អាតទិន្នន័យទាំងអស់ក្នុងតារាងនេះមែនទេ?')) return;
    const cleared: Record<string, RowDistributionData> = {};
    REPORT_VISA_TYPES.forEach((vt) => {
      cleared[vt] = {
        visaType: vt,
        openingStock: 0,
        monthlyValues: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
        totalMonthly: 0,
        totalAvailable: 0,
        used: 0,
        damagedMissing: 0,
        returnStock: 0,
        remaining: 0,
      };
    });
    setTableData(cleared);
  };

  // Export to Excel (.xlsx)
  const handleExportExcel = () => {
    const wb = XLSX.utils.book_new();

    // Sheet data array with formatted headers
    const wsData: any[][] = [];

    // Title rows
    wsData.push([ministryName]);
    wsData.push([departmentName]);
    wsData.push([generalDeptName]);
    wsData.push([officeName]);
    wsData.push([sectionName]);
    wsData.push([]);
    wsData.push([
      `តារាងសន្លឹកទិដ្ឋាការស្អិត តាមប្រភេទនីមួយៗ បើកពីការិយាល័យទិដ្ឋាការចូល`,
    ]);
    wsData.push([
      `ដោយគិតចាប់ពីថ្ងៃទី០១ ខែ${KHMER_MONTHS_NAMES[appliedStartMonth - 1]} ឆ្នាំ${toKhmerNum(appliedStartYear)} រហូតដល់ថ្ងៃទី${toKhmerNum(new Date(monthColumnsInfo[11].year, monthColumnsInfo[11].monthNum, 0).getDate())} ខែ${monthColumnsInfo[11].name} ឆ្នាំ${toKhmerNum(monthColumnsInfo[11].year)}`,
    ]);
    wsData.push([]);

    // Table Header 1
    const h1 = [
      'ប្រភេទ',
      `សន្និធិចុងគ្រា (${openingStockLabel})`,
      ...monthColumnsInfo.map((m) => m.name),
      'សរុប (១២ខែ)',
      'ចំនួនសរុប',
      'ប្រើប្រាស់សរុប',
      isSecondary ? 'ទិដ្ឋាការ ខ្វះ&ខូច' : 'ផ្ទេរ/ខ្វះ/ខូច',
      'បង្វិលកង',
      `សន្និធិនៅសល់ (${endingStockLabel})`,
    ];
    wsData.push(h1);

    // Data rows
    REPORT_VISA_TYPES.forEach((vt) => {
      const r = tableData[vt] || {
        openingStock: 0,
        monthlyValues: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
        totalMonthly: 0,
        totalAvailable: 0,
        used: 0,
        damagedMissing: 0,
        returnStock: 0,
        remaining: 0,
      };

      wsData.push([
        vt,
        r.openingStock,
        ...r.monthlyValues,
        r.totalMonthly,
        r.totalAvailable,
        r.used,
        r.damagedMissing,
        r.returnStock,
        r.remaining,
      ]);
    });

    // Grand total row
    wsData.push([
      'សរុប',
      grandTotals.opening,
      ...grandTotals.months,
      grandTotals.monthlySum,
      grandTotals.available,
      grandTotals.used,
      grandTotals.damagedMissing,
      grandTotals.returnStock,
      grandTotals.remaining,
    ]);

    // Signatures
    wsData.push([]);
    wsData.push([]);
    wsData.push(['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', signatureLunarDate]);
    wsData.push(['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', `ភ្នំពេញ, ថ្ងៃទី${signatureSolarParts.day} ខែ${signatureSolarParts.month} ឆ្នាំ${signatureSolarParts.year}`]);
    wsData.push(['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', signerRole]);
    wsData.push([]);
    wsData.push([]);
    wsData.push(['', '', '', '', '', '', '', '', '', '', '', '', '', '', '', signerName]);

    const ws = XLSX.utils.aoa_to_sheet(wsData);
    XLSX.utils.book_append_sheet(wb, ws, 'តារាងផ្តល់ប្រចាំឆ្នាំទៅក្រុម');
    XLSX.writeFile(wb, `Yearly_Sticker_Team_Distribution_${appliedStartYear}_${appliedStartYear + 1}.xlsx`);
  };

  // Export to PDF (Landscape A4)
  const handleExportPdf = async () => {
    const el = printAreaRef.current;
    if (!el) return;
    try {
      setIsGeneratingPdf(true);
      await exportSinglePageA4LandscapePdf(
        el,
        `តារាងផ្តល់ប្រចាំឆ្នាំទៅក្រុម_${startYear}_${startYear + 1}.pdf`,
        'yearly-team-distribution-print-area'
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
    printA4Document('yearly-team-distribution-print-area', {
      orientation: 'landscape',
      documentTitle: `តារាងផ្តល់ប្រចាំឆ្នាំទៅក្រុម_${appliedStartYear}_${appliedStartMonth}`,
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
        {/* Title & Presets Bar */}
        <div className="flex flex-wrap items-center justify-between gap-3 border-b border-gray-100 pb-3">
          <div className="flex items-center gap-2.5">
            <div className="w-9 h-9 rounded-lg bg-amber-100 text-amber-900 flex items-center justify-center font-bold">
              <Calendar className="w-5 h-5 text-amber-700" />
            </div>
            <div>
              <h2 className="text-base font-bold text-gray-800">
                តារាងផ្តល់ប្រចាំឆ្នាំទៅក្រុម — សន្លឹកទិដ្ឋាការស្អិត
              </h2>
              <p className="text-xs text-gray-500">
                តារាងសន្លឹកទិដ្ឋាការស្អិត តាមប្រភេទនីមួយៗ បើកពីការិយាល័យទិដ្ឋាការចូល ប្រចាំឆ្នាំ (១២ខែ)
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
              <Calendar className="w-3.5 h-3.5 text-amber-600" />
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
                isFilterPending || !hasCalculated ? 'border-amber-400 ring-1 ring-amber-300' : 'border-slate-200 hover:border-slate-300 focus:border-amber-500 focus:ring-amber-500/20'
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
              <CalendarDays className="w-3.5 h-3.5 text-amber-600" />
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
                isFilterPending || !hasCalculated ? 'border-amber-400 ring-1 ring-amber-300' : 'border-slate-200 hover:border-slate-300 focus:border-amber-500 focus:ring-amber-500/20'
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
              <Sparkles className="w-3.5 h-3.5 text-amber-600" />
              <span>សកម្មភាពទាញយក ៖</span>
            </label>
            <button
              type="button"
              onClick={() => handleApplyFilter()}
              disabled={isCalculating}
              className={`w-full py-2 px-3 rounded-lg text-xs font-bold flex items-center justify-center gap-1.5 shadow-xs transition-all cursor-pointer ${
                isFilterPending || !hasCalculated
                  ? 'bg-amber-600 hover:bg-amber-700 text-white ring-2 ring-amber-400/80 shadow-md animate-pulse'
                  : 'bg-gradient-to-r from-amber-700 to-yellow-800 hover:from-amber-800 hover:to-yellow-900 text-white'
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
              <Calendar className="w-3.5 h-3.5 text-amber-600" />
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
              <UserCheck className="w-3.5 h-3.5 text-amber-600" />
              <span>មន្ត្រីចុះហត្ថលេខា ៖</span>
            </label>
            <div className="flex rounded-lg border border-slate-200 hover:border-slate-300 focus-within:border-amber-500 focus-within:ring-2 focus-within:ring-amber-500/20 bg-slate-50/60 focus-within:bg-white overflow-hidden transition-all">
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
              className="bg-white border border-gray-300 rounded px-2 py-1 focus:outline-none"
            >
              <option value="Khmer OS Siemreap">Khmer OS Siemreap (ស្តង់ដារ PDF)</option>
              <option value="Khmer OS Battambang">Khmer OS Battambang</option>
              <option value="Khmer OS Muol Light">Khmer OS Muol Light</option>
              <option value="Moul">Moul</option>
              <option value="sans-serif">System Sans</option>
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
                  localStorage.setItem('yearly_team_distribution_row_height', String(val));
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
                  localStorage.setItem('yearly_team_distribution_row_height', String(val));
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
              className="bg-white border border-amber-300 text-amber-950 font-bold rounded px-2 py-1 focus:outline-none"
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
                className="rounded text-amber-600 focus:ring-amber-500"
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
          id="yearly-team-distribution-print-area"
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
              contentEditable
              suppressContentEditableWarning
              onBlur={(e) => setReportTitle(e.currentTarget.textContent || 'តារាងសន្លឹកទិដ្ឋាការស្អិត តាមប្រភេទនីមួយៗ បើកពីការិយាល័យទិដ្ឋាការចូល')}
              style={{
                fontSize: '12pt',
                fontFamily: "'Khmer Mool1', 'Khmer Mool', 'Khmer OS Muol Light', 'Khmer OS Muol', 'Moul', serif"
              }}
              className="text-[12pt] font-moul text-black font-normal leading-relaxed outline-none hover:bg-amber-50/50 rounded px-2"
            >
              {reportTitle}
            </h1>
            <p
              style={{ fontSize: '12pt' }}
              className="text-[12pt] font-siemreap font-bold mt-1 text-black"
            >
              ដោយគិតចាប់ពីថ្ងៃទី០១ ខែ{KHMER_MONTHS_NAMES[appliedStartMonth - 1]} ឆ្នាំ{toKhmerNum(appliedStartYear)} រហូតដល់ថ្ងៃទី{toKhmerNum(new Date(monthColumnsInfo[11].year, monthColumnsInfo[11].monthNum, 0).getDate())} ខែ{monthColumnsInfo[11].name} ឆ្នាំ{toKhmerNum(monthColumnsInfo[11].year)}
            </p>
          </div>

          {/* Main Distribution Table */}
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
                <tr className="bg-[#FCE4D6] font-bold border-b border-black text-black h-[26px]">
                  <th
                    rowSpan={2}
                    className="border border-black px-1 py-1 min-w-[36px] max-w-[40px] text-center align-middle font-bold font-siemreap text-[11px] whitespace-nowrap"
                  >
                    ប្រភេទ
                  </th>
                  <th
                    rowSpan={2}
                    className="border border-black px-1 py-1 min-w-[98px] text-center align-middle font-bold font-siemreap"
                  >
                    <div className="whitespace-nowrap font-bold text-[11px] leading-tight">សន្និធិចុងគ្រា</div>
                    <div className="text-[9px] font-semibold text-gray-800 mt-0.5 whitespace-nowrap leading-tight">
                      {openingStockLabel}
                    </div>
                  </th>
                  <th
                    colSpan={13}
                    className="border border-black px-1 py-1 text-center font-bold font-siemreap text-[11.5px] leading-normal whitespace-nowrap h-[24px] align-middle"
                  >
                    គិតចាប់ពីថ្ងៃទី០១ ខែ{KHMER_MONTHS_NAMES[appliedStartMonth - 1]} ឆ្នាំ{toKhmerNum(appliedStartYear)} រហូតដល់ថ្ងៃទី{toKhmerNum(new Date(monthColumnsInfo[11].year, monthColumnsInfo[11].monthNum, 0).getDate())} ខែ{monthColumnsInfo[11].name} ឆ្នាំ{toKhmerNum(monthColumnsInfo[11].year)}
                  </th>
                  <th
                    rowSpan={2}
                    onClick={() => setShowReceivedK2ChecklistModal(true)}
                    className="border border-black px-1 py-1 min-w-[48px] max-w-[56px] text-center align-middle font-bold font-siemreap text-[10.5px] leading-tight hover:bg-amber-300 transition cursor-pointer select-none group"
                    title="ចុចដើម្បីពិនិត្យផ្ទៀងផ្ទាត់បញ្ជីទិដ្ឋាការ បើកពីក២ (ចែកជូនក្រុម) ទាំង ១២ខែ"
                  >
                    <div className="group-hover:underline whitespace-nowrap">ចំនួន</div>
                    <div className="group-hover:underline whitespace-nowrap mt-0.5">សរុប</div>
                  </th>
                  <th
                    rowSpan={2}
                    className="border border-black px-1 py-1 min-w-[50px] max-w-[58px] text-center align-middle font-bold font-siemreap text-[10.5px] leading-tight"
                  >
                    <div className="whitespace-nowrap">ប្រើប្រាស់</div>
                    <div className="whitespace-nowrap mt-0.5">សរុប</div>
                  </th>
                  <th
                    rowSpan={2}
                    onClick={() => setShowDamagedChecklistModal(true)}
                    className="border border-black px-1 py-1 min-w-[46px] max-w-[52px] text-center align-middle font-bold font-siemreap text-[10.5px] leading-tight hover:bg-amber-300 transition cursor-pointer select-none group"
                    title={isSecondary ? "ចុចដើម្បីពិនិត្យផ្ទៀងផ្ទាត់បញ្ជីទិដ្ឋាការ ខ្វះ & ខូច ទាំង ១២ខែ" : "ចុចដើម្បីពិនិត្យផ្ទៀងផ្ទាត់បញ្ជីទិដ្ឋាការ ផ្ទេរ / ខ្វះ / ខូច ទាំង ១២ខែ"}
                  >
                    {isSecondary ? (
                      <>
                        <div className="group-hover:underline whitespace-nowrap">ទិដ្ឋាការ</div>
                        <div className="text-[9px] group-hover:underline whitespace-nowrap mt-0.5 leading-tight">ខ្វះ&ខូច</div>
                      </>
                    ) : (
                      <div className="group-hover:underline whitespace-nowrap">ផ្ទេរ/ខ្វះ/ខូច</div>
                    )}
                  </th>
                  <th
                    rowSpan={2}
                    className="border border-black px-1 py-1 min-w-[44px] max-w-[50px] text-center align-middle font-bold font-siemreap text-[10.5px] leading-tight"
                  >
                    <div className="whitespace-nowrap">បង្វិលក២</div>
                  </th>
                  <th
                    rowSpan={2}
                    className="border border-black px-1 py-1 min-w-[98px] text-center align-middle font-bold font-siemreap"
                  >
                    <div className="whitespace-nowrap font-bold text-[11px] leading-tight">សន្និធិនៅសល់</div>
                    <div className="text-[9px] font-semibold text-gray-800 mt-0.5 whitespace-nowrap leading-tight">
                      {endingStockLabel}
                    </div>
                  </th>
                </tr>

                {/* Header Row 2: 12 Months + Sum */}
                <tr className="bg-[#FCE4D6] font-bold border-b border-black text-[10.5px] font-siemreap leading-normal text-black h-[20px]">
                  {monthColumnsInfo.map((mCol, idx) => (
                    <th
                      key={idx}
                      onClick={() => {
                        setChecklistInitialMonth(idx);
                        setShowMonthChecklistModal(true);
                      }}
                      className="border border-black px-0.5 py-0.5 min-w-[44px] text-center align-middle hover:bg-amber-300 transition cursor-pointer select-none group whitespace-nowrap"
                      title={`ចុចដើម្បីពិនិត្យផ្ទៀងផ្ទាត់បញ្ជី និងចំនួនសរុបខែ ${mCol.name} (${mCol.year})`}
                    >
                      <span className="group-hover:underline">{mCol.name}</span>
                    </th>
                  ))}
                  <th
                    onClick={() => {
                      setChecklistInitialMonth(-1);
                      setShowMonthChecklistModal(true);
                    }}
                    className="border border-black px-1 py-0.5 min-w-[52px] text-center align-middle font-bold bg-[#F8CBAD] hover:bg-amber-300 transition cursor-pointer select-none whitespace-nowrap"
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
                    totalAvailable: 0,
                    used: 0,
                    damagedMissing: 0,
                    returnStock: 0,
                    remaining: 0,
                  };

                  const dynamicMonthlySum = row.monthlyValues.reduce((a, b) => (Number(a) || 0) + (Number(b) || 0), 0);
                  const displayMonthlySum = autoCalculate ? dynamicMonthlySum : (row.totalMonthly ?? dynamicMonthlySum);
                  const displayAvailable = autoCalculate
                    ? (Number(row.openingStock) || 0) + displayMonthlySum
                    : (row.totalAvailable ?? ((Number(row.openingStock) || 0) + displayMonthlySum));
                  const displayRemaining = autoCalculate
                    ? displayAvailable - (Number(row.used) || 0) - (Number(row.damagedMissing) || 0) - (Number(row.returnStock) || 0)
                    : (row.remaining ?? (displayAvailable - (Number(row.used) || 0) - (Number(row.damagedMissing) || 0) - (Number(row.returnStock) || 0)));

                  return (
                    <tr
                      key={vt}
                      style={{ height: `${tableRowHeight}px` }}
                      className={`hover:bg-amber-50/40 transition-colors ${
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
                        className="border border-black px-1 py-[1px] text-right font-normal font-times outline-none hover:bg-yellow-100/70 focus:bg-yellow-100 focus:ring-1 focus:ring-amber-500 text-[11px] leading-tight"
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
                            className="border border-black px-0.5 py-[1px] text-right font-times outline-none hover:bg-yellow-100/70 focus:bg-yellow-100 focus:ring-1 focus:ring-amber-500 font-normal text-[10.5px] leading-tight"
                          >
                            {renderCellValue(val)}
                          </td>
                        );
                      })}

                      {/* Total Monthly Sum (12 Months) */}
                      <td
                        contentEditable={!autoCalculate}
                        suppressContentEditableWarning
                        onBlur={(e) =>
                          !autoCalculate &&
                          handleCellChange(vt, 'totalMonthly', e.currentTarget.textContent || '0')
                        }
                        className={`border border-black px-1 py-[1px] text-right font-normal font-times bg-[#FDF5ED] text-[11px] leading-tight ${
                          !autoCalculate ? 'outline-none hover:bg-yellow-100' : ''
                        }`}
                      >
                        {renderCellValue(displayMonthlySum)}
                      </td>

                      {/* Total Available (Opening + Monthly Sum) */}
                      <td
                        contentEditable={!autoCalculate}
                        suppressContentEditableWarning
                        onBlur={(e) =>
                          !autoCalculate &&
                          handleCellChange(vt, 'totalAvailable', e.currentTarget.textContent || '0')
                        }
                        className={`border border-black px-0.5 py-[1px] text-right font-bold font-times bg-amber-50/40 text-[11px] leading-tight ${
                          !autoCalculate ? 'outline-none hover:bg-yellow-100' : ''
                        }`}
                      >
                        {renderCellValue(displayAvailable)}
                      </td>

                      {/* Used (Editable) */}
                      <td
                        contentEditable
                        suppressContentEditableWarning
                        onBlur={(e) => handleCellChange(vt, 'used', e.currentTarget.textContent || '0')}
                        className="border border-black px-0.5 py-[1px] text-right font-normal font-times outline-none hover:bg-yellow-100/70 focus:bg-yellow-100 focus:ring-1 focus:ring-amber-500 text-[11px] leading-tight"
                      >
                        {renderCellValue(row.used)}
                      </td>

                      {/* Missing & Damaged (Editable) */}
                      <td
                        contentEditable
                        suppressContentEditableWarning
                        onBlur={(e) =>
                          handleCellChange(vt, 'damagedMissing', e.currentTarget.textContent || '0')
                        }
                        className="border border-black px-0.5 py-[1px] text-right font-normal font-times outline-none hover:bg-yellow-100/70 focus:bg-yellow-100 focus:ring-1 focus:ring-amber-500 text-[11px] leading-tight"
                      >
                        {renderCellValue(row.damagedMissing)}
                      </td>

                      {/* Return to Stock / បង្វិលកង (Editable) */}
                      <td
                        contentEditable
                        suppressContentEditableWarning
                        onBlur={(e) =>
                          handleCellChange(vt, 'returnStock', e.currentTarget.textContent || '0')
                        }
                        className="border border-black px-0.5 py-[1px] text-right font-normal font-times outline-none hover:bg-yellow-100/70 focus:bg-yellow-100 focus:ring-1 focus:ring-amber-500 text-[11px] leading-tight"
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
                        className={`border border-black px-1 py-[1px] text-right font-bold font-times bg-amber-50/50 text-[11px] leading-tight ${
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
                  className="bg-[#FCE4D6] font-bold border-t-2 border-black text-black"
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
                  <td className="border border-black px-1 py-[1px] text-right font-bold font-times bg-[#F8CBAD] text-[11px] leading-tight">
                    {renderCellValue(grandTotals.monthlySum)}
                  </td>

                  {/* Available Grand Total */}
                  <td className="border border-black px-0.5 py-[1px] text-right font-bold font-times bg-[#F4B183]/60 text-[11px] leading-tight">
                    {renderCellValue(grandTotals.available)}
                  </td>

                  {/* Used Grand Total */}
                  <td className="border border-black px-0.5 py-[1px] text-right font-bold font-times text-[11px] leading-tight">
                    {renderCellValue(grandTotals.used)}
                  </td>

                  {/* Missing/Damaged Grand Total */}
                  <td className="border border-black px-0.5 py-[1px] text-right font-bold font-times text-[11px] leading-tight">
                    {renderCellValue(grandTotals.damagedMissing)}
                  </td>

                  {/* Return Grand Total */}
                  <td className="border border-black px-0.5 py-[1px] text-right font-bold font-times text-[11px] leading-tight">
                    {renderCellValue(grandTotals.returnStock)}
                  </td>

                  {/* Remaining Grand Total */}
                  <td className="border border-black px-1 py-[1px] text-right font-bold font-times bg-[#F4B183]/60 text-[11px] leading-tight">
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
          tableData={tableData}
          stockRecords={stockRecords}
          initialMonthIndex={checklistInitialMonth}
          targetTeamName={activeFilteringTeam}
          mode="distribution"
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
          tableData={tableData}
          stockRecords={stockRecords}
          targetTeamName={activeFilteringTeam}
          mode="distribution"
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
          tableData={tableData}
          stockRecords={stockRecords}
          targetTeamName={activeFilteringTeam}
          mode="distribution"
        />
      )}
    </div>
  );
};
