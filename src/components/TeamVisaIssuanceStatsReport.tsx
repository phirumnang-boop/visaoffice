import React, { useState, useEffect, useMemo, useRef, useCallback } from 'react';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import { CategoriesState, Officer, StockRecord, UserRole } from '../types';
import { sanitizeDocumentForHtml2Canvas, exportElementToPdf } from '../utils/pdfExportHelper';
import { useWorkspaceSettings } from '../context/WorkspaceSettingsContext';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import {
  Printer,
  Download,
  Calendar,
  Save,
  RotateCcw,
  FileSpreadsheet,
  FileText,
  ZoomIn,
  ZoomOut,
  ChevronDown,
  Sparkles,
  SlidersHorizontal,
  CheckCircle2,
  Eye,
  Layers,
  ArrowLeft,
  Search,
  AlertCircle,
} from 'lucide-react';
import { getKhmerLunarDate, getKhmerSolarParts, toKhmerNum, parseDateInput } from '../utils/khmerCalendar';
import { printA4Document } from '../utils/printHelper';
import { CustomDatePicker } from './CustomDatePicker';
import {
  formatReportTeamName,
  normalizeTeamName,
  normalizeDateToISO,
  normalizeVisaType,
  matchTeamInList,
  OFFICIAL_29_TEAMS,
  OFFICIAL_TEAM_FULL_NAMES,
} from '../utils/teamNormalization';
import {
  resolveRecordTeamName,
  isCeaRecord,
} from '../utils/teamStockCalculation';
import { TacteingLine, TacteingControlSelector, getSavedTacteingSettings, saveTacteingSettings, TacteingType } from './TacteingLine';

export interface TeamVisaIssuanceStatsReportProps {
  stockRecords?: StockRecord[];
  categories?: CategoriesState;
  officers?: Officer[];
  currentRole?: UserRole;
  userName?: string;
  assignedTeam?: string;
  onClose?: () => void;
  isBranchChiefMode?: boolean;
  singleDateMode?: boolean;
  storageKeyPrefix?: string;
}

// 13 Official Visa Types
export const VISA_TYPE_KEYS = [
  'T', 'T1', 'T2', 'T3', 'E', 'E1', 'E2', 'E3', 'D',
  'K', 'A', 'B', 'C'
] as const;

export type VisaTypeKey = (typeof VISA_TYPE_KEYS)[number];

// Group ก: ប្រភេទទិដ្ឋាការបង់អាករ (Taxable Visas)
export const TAXABLE_VISA_TYPES: Array<{ key: VisaTypeKey; labelKh: string; code: string }> = [
  { key: 'T', labelKh: '1-ទិដ្ឋាការទេសចរណ៍', code: 'T' },
  { key: 'T1', labelKh: '2-ទិដ្ឋាការទេសចរណ៍', code: 'T1' },
  { key: 'T2', labelKh: '3-ទិដ្ឋាការទេសចរណ៍', code: 'T2' },
  { key: 'T3', labelKh: '4-ទិដ្ឋាការទេសចរណ៍', code: 'T3' },
  { key: 'E', labelKh: '5-ទិដ្ឋាការធម្មតា', code: 'E' },
  { key: 'E1', labelKh: '6-ទិដ្ឋាការធម្មតា', code: 'E1' },
  { key: 'E2', labelKh: '7-ទិដ្ឋាការធម្មតា', code: 'E2' },
  { key: 'E3', labelKh: '8-ទិដ្ឋាការធម្មតា', code: 'E3' },
  { key: 'D', labelKh: '9-ទិដ្ឋាការឆ្លងកាត់', code: 'D' },
];

// Group ខ: ប្រភេទទិដ្ឋាការមិនបង់អាករ (Non-taxable / Gratis Visas)
export const NON_TAXABLE_VISA_TYPES: Array<{ key: VisaTypeKey; labelKh: string; code: string }> = [
  { key: 'K', labelKh: '1-ទិដ្ឋាការពិសេស', code: 'K' },
  { key: 'A', labelKh: '2-ទិដ្ឋាការការទូត', code: 'A' },
  { key: 'B', labelKh: '3-ទិដ្ឋាការផ្លូវការ', code: 'B' },
  { key: 'C', labelKh: '4-ទិដ្ឋាការគួរសម', code: 'C' },
];

// PDF Sample data from user's attached document (Q1 2026: 127,954 total)
export const SAMPLE_Q1_2026_VALUES: Record<VisaTypeKey, number> = {
  T: 107472,
  T1: 494,
  T2: 27,
  T3: 40,
  E: 13254,
  E1: 154,
  E2: 13,
  E3: 39,
  D: 458,
  K: 5572,
  A: 118,
  B: 269,
  C: 44,
};

// Realistic baseline & sample data for the 4 cEA teams
export const SAMPLE_CEA_TEAMS_Q1: Record<string, Record<VisaTypeKey, number>> = {
  'អាកាស តេជោ': {
    T: 107472,
    T1: 494,
    T2: 27,
    T3: 40,
    E: 13254,
    E1: 154,
    E2: 13,
    E3: 39,
    D: 458,
    K: 5572,
    A: 118,
    B: 269,
    C: 44,
  },
  'អាកាស សៀមរាប': {
    T: 85210,
    T1: 312,
    T2: 18,
    T3: 25,
    E: 9420,
    E1: 110,
    E2: 8,
    E3: 15,
    D: 215,
    K: 4102,
    A: 85,
    B: 174,
    C: 32,
  },
  'អាកាស ព្រះសីហនុ': {
    T: 18450,
    T1: 35,
    T2: 12,
    T3: 15,
    E: 12380,
    E1: 48,
    E2: 6,
    E3: 10,
    D: 180,
    K: 1240,
    A: 24,
    B: 52,
    C: 18,
  },
  'កំពង់ផែ ព្រះសីហនុ': {
    T: 6240,
    T1: 22,
    T2: 8,
    T3: 10,
    E: 1820,
    E1: 15,
    E2: 4,
    E3: 6,
    D: 140,
    K: 320,
    A: 12,
    B: 38,
    C: 14,
  },
};

export const KHMER_MONTHS = [
  'មករា', 'កុម្ភៈ', 'មីនា', 'មេសា', 'ឧសភា', 'មិថុនា',
  'កក្កដា', 'សីហា', 'កញ្ញា', 'តុលា', 'វិច្ឆិកា', 'ធ្នូ',
];

export const createEmptyVisaCounts = (): Record<VisaTypeKey, number> => ({
  T: 0, T1: 0, T2: 0, T3: 0,
  E: 0, E1: 0, E2: 0, E3: 0,
  D: 0, K: 0, A: 0, B: 0, C: 0,
});

export const TeamVisaIssuanceStatsReport: React.FC<TeamVisaIssuanceStatsReportProps> = ({
  stockRecords = [],
  categories,
  officers = [],
  currentRole = 'Secondary',
  userName = '',
  assignedTeam,
  onClose,
  isBranchChiefMode = false,
  singleDateMode = false,
  storageKeyPrefix = '',
}) => {
  const { scaleMode } = useWorkspaceSettings();
  const isSecondary = currentRole === 'Secondary';
  const baseKeyPrefix = isBranchChiefMode ? 'branch_chief_' : '';
  const keyPrefix = `${storageKeyPrefix || (singleDateMode ? 'single_day_' : '')}${baseKeyPrefix}`;
  const reportMainTitle = isBranchChiefMode 
    ? (singleDateMode ? 'ស្ថិតិរបស់ក្រុមប្រធានសាខា (១ ថ្ងៃ)' : 'ស្ថិតិរបស់ក្រុមប្រធានសាខា')
    : (singleDateMode ? 'ស្ថិតិផ្តល់ប្រចាំថ្ងៃរបស់ក្រុម' : 'ស្ថិតិផ្តល់ទិដ្ឋាការរបស់ក្រុម');
  const reportSubtitle = isBranchChiefMode
    ? 'របាយការណ៍ផ្លូវការស្ថិតិរបស់ក្រុមប្រធានសាខា'
    : 'របាយការណ៍ផ្លូវការចំនួនភ្ញៀវបរទេសស្នើសុំទិដ្ឋាការចូល';

  // Team Selection
  const [selectedTeam, setSelectedTeam] = useState<string>(() => {
    if (assignedTeam && assignedTeam.trim()) {
      return assignedTeam;
    }
    return 'ព្រំដែន បាវិត';
  });

  // Sync if assignedTeam prop changes
  useEffect(() => {
    if (assignedTeam && assignedTeam.trim() && assignedTeam !== selectedTeam) {
      setSelectedTeam(assignedTeam);
      setIsResultShown(false);
    }
  }, [assignedTeam]);

  // Date selection states (matching RobokTotalStockWorkReport)
  const [selectedMonth, setSelectedMonth] = useState<number>(7);
  const [selectedYear, setSelectedYear] = useState<number>(2026);
  const [startDate, setStartDate] = useState<string>(() => singleDateMode ? '2026-07-01' : '2026-01-01');
  const [endDate, setEndDate] = useState<string>(() => singleDateMode ? '2026-07-01' : '2026-03-31');
  const [activeDurationPreset, setActiveDurationPreset] = useState<
    'this_month' | 'last_month' | 'next_month' | 'q1' | 's1' | 'm9' | 'full_year' | null
  >(singleDateMode ? null : 'q1');

  // Helper to calculate start & end date for a given year, month, and duration
  const calculateDateRange = (year: number, month: number, durationMonths: number) => {
    const startMStr = String(month).padStart(2, '0');
    const s = `${year}-${startMStr}-01`;

    const totalMonthsOffset = (month - 1) + (durationMonths - 1);
    const endYear = year + Math.floor(totalMonthsOffset / 12);
    const endMonth = (totalMonthsOffset % 12) + 1;
    const daysInEndMonth = new Date(endYear, endMonth, 0).getDate();
    const endMStr = String(endMonth).padStart(2, '0');
    const e = `${endYear}-${endMStr}-${String(daysInEndMonth).padStart(2, '0')}`;

    return { startDate: s, endDate: e };
  };

  // Derive periodLabelKh and periodDetailsKh from startDate & endDate
  const { periodLabelKh, periodDetailsKh } = useMemo(() => {
    if (!startDate || !endDate) {
      return {
        periodLabelKh: '',
        periodDetailsKh: '',
      };
    }

    const sParts = startDate.split('-');
    const eParts = endDate.split('-');
    const sY = parseInt(sParts[0], 10) || 2026;
    const sM = parseInt(sParts[1], 10) || 1;
    const sD = parseInt(sParts[2], 10) || 1;
    const eY = parseInt(eParts[0], 10) || 2026;
    const eM = parseInt(eParts[1], 10) || 1;
    const eD = parseInt(eParts[2], 10) || 31;

    const sYearKh = toKhmerNum(sY);
    const eYearKh = toKhmerNum(eY);
    const sMonthKh = KHMER_MONTHS[sM - 1] || '';
    const eMonthKh = KHMER_MONTHS[eM - 1] || '';
    const sDayKh = toKhmerNum(sD);
    const eDayKh = toKhmerNum(eD);

    // 1. Single Day
    if (sY === eY && sM === eM && sD === eD) {
      return {
        periodLabelKh: `ប្រចាំថ្ងៃទី${sDayKh} ខែ${sMonthKh} ឆ្នាំ${sYearKh}។`,
        periodDetailsKh: '',
      };
    }

    // 2. Exact fiscal Q1 cross-year boundary (Dec 01 to Feb 28/29 of following year, e.g. 2018-12-01 to 2019-02-28)
    if (sD === 1 && sM === 12 && eM === 2 && eY === sY + 1 && (eD === 28 || eD === 29)) {
      return {
        periodLabelKh: `ប្រចាំត្រីមាសទី១ ឆ្នាំ${eYearKh}`,
        periodDetailsKh: `ដោយគិតចាប់ពីថ្ងៃទី០១ ខែធ្នូ ឆ្នាំ${sYearKh} រហូតដល់ថ្ងៃទី${eDayKh} ខែកុម្ភៈ ឆ្នាំ${eYearKh}`,
      };
    }

    // Check if both start and end fall on whole month boundaries (start day 1, end day last day of month)
    const lastDayOfEndMonth = new Date(eY, eM, 0).getDate();
    const isWholeMonthRange = sD === 1 && eD === lastDayOfEndMonth;

    if (isWholeMonthRange && sY === eY) {
      const monthDiff = eM - sM + 1;

      // Single Whole Month
      if (monthDiff === 1) {
        return {
          periodLabelKh: `ប្រចាំខែ${sMonthKh} ឆ្នាំ${sYearKh}`,
          periodDetailsKh: `ដោយគិតចាប់ពីថ្ងៃទី០១ រហូតដល់ថ្ងៃទី${eDayKh} ខែ${sMonthKh} ឆ្នាំ${sYearKh}។`,
        };
      }

      // 3 Months (Quarter) in same calendar year
      if (monthDiff === 3) {
        let qNum = 1;
        if (sM === 1) qNum = 1;
        else if (sM === 4) qNum = 2;
        else if (sM === 7) qNum = 3;
        else if (sM === 10) qNum = 4;
        else qNum = Math.ceil(sM / 3);

        const qKhmer = toKhmerNum(qNum);
        return {
          periodLabelKh: `ប្រចាំត្រីមាសទី${qKhmer} ឆ្នាំ${sYearKh}`,
          periodDetailsKh: `ដោយគិតចាប់ពីថ្ងៃទី០១ ខែ${sMonthKh} រហូតដល់ថ្ងៃទី${eDayKh} ខែ${eMonthKh} ឆ្នាំ${sYearKh}។`,
        };
      }

      // 6 Months (Semester)
      if (monthDiff === 6) {
        const semNum = sM <= 6 ? 1 : 2;
        return {
          periodLabelKh: `ប្រចាំឆមាសទី${toKhmerNum(semNum)} ឆ្នាំ${sYearKh}`,
          periodDetailsKh: `ដោយគិតចាប់ពីថ្ងៃទី០១ ខែ${sMonthKh} រហូតដល់ថ្ងៃទី${eDayKh} ខែ${eMonthKh} ឆ្នាំ${sYearKh}។`,
        };
      }

      // 9 Months
      if (monthDiff === 9) {
        return {
          periodLabelKh: `ប្រចាំនព្វមាស ឆ្នាំ${sYearKh}`,
          periodDetailsKh: `ដោយគិតចាប់ពីថ្ងៃទី០១ ខែ${sMonthKh} រហូតដល់ថ្ងៃទី${eDayKh} ខែ${eMonthKh} ឆ្នាំ${sYearKh}។`,
        };
      }

      // 12 Months (Full Year)
      if (monthDiff === 12) {
        return {
          periodLabelKh: `ប្រចាំឆ្នាំ${sYearKh}`,
          periodDetailsKh: `ដោយគិតចាប់ពីថ្ងៃទី០១ ខែ${sMonthKh} រហូតដល់ថ្ងៃទី${eDayKh} ខែ${eMonthKh} ឆ្នាំ${sYearKh}។`,
        };
      }
    }

    // 3. Custom range within SAME month
    if (sY === eY && sM === eM) {
      return {
        periodLabelKh: `ពីថ្ងៃទី${sDayKh} ដល់ថ្ងៃទី${eDayKh} ខែ${sMonthKh} ឆ្នាំ${sYearKh}`,
        periodDetailsKh: `ដោយគិតចាប់ពីថ្ងៃទី${sDayKh} រហូតដល់ថ្ងៃទី${eDayKh} ខែ${sMonthKh} ឆ្នាំ${sYearKh}។`,
      };
    }

    // 4. Custom range within SAME year across different months
    if (sY === eY && sM !== eM) {
      return {
        periodLabelKh: `ពីថ្ងៃទី${sDayKh} ខែ${sMonthKh} ដល់ថ្ងៃទី${eDayKh} ខែ${eMonthKh} ឆ្នាំ${sYearKh}`,
        periodDetailsKh: `ដោយគិតចាប់ពីថ្ងៃទី${sDayKh} ខែ${sMonthKh} រហូតដល់ថ្ងៃទី${eDayKh} ខែ${eMonthKh} ឆ្នាំ${sYearKh}។`,
      };
    }

    // 5. Custom range across different years
    return {
      periodLabelKh: `ពីថ្ងៃទី${sDayKh} ខែ${sMonthKh} ឆ្នាំ${sYearKh} ដល់ថ្ងៃទី${eDayKh} ខែ${eMonthKh} ឆ្នាំ${eYearKh}`,
      periodDetailsKh: `ដោយគិតចាប់ពីថ្ងៃទី${sDayKh} ខែ${sMonthKh} ឆ្នាំ${sYearKh} រហូតដល់ថ្ងៃទី${eDayKh} ខែ${eMonthKh} ឆ្នាំ${eYearKh}។`,
    };
  }, [startDate, endDate]);

  // Date handlers matching RobokTotalStockWorkReport
  const handleMonthYearChange = (year: number, month: number) => {
    setSelectedYear(year);
    setSelectedMonth(month);
    setIsResultShown(false);

    let duration = 1;
    if (activeDurationPreset === 'q1') duration = 3;
    else if (activeDurationPreset === 's1') duration = 6;
    else if (activeDurationPreset === 'm9') duration = 9;
    else if (activeDurationPreset === 'full_year') duration = 12;

    const range = calculateDateRange(year, month, duration);
    setStartDate(range.startDate);
    setEndDate(range.endDate);

    // Auto-update signingDate to 1st of the next month following end date
    const eParts = range.endDate.split('-');
    const eY = parseInt(eParts[0], 10);
    const eM = parseInt(eParts[1], 10);
    const nextY = eM === 12 ? eY + 1 : eY;
    const nextM = eM === 12 ? 1 : eM + 1;
    setSigningDate(`${nextY}-${String(nextM).padStart(2, '0')}-01`);
  };

  const handleStartDateChange = (val: string) => {
    setStartDate(val);
    setActiveDurationPreset(null);
    setIsResultShown(false);
    if (val) {
      const parts = val.split('-');
      if (parts.length === 3) {
        const y = parseInt(parts[0], 10);
        const m = parseInt(parts[1], 10);
        if (!isNaN(y) && y >= 1900 && y <= 2100) setSelectedYear(y);
        if (!isNaN(m) && m >= 1 && m <= 12) setSelectedMonth(m);
      }
    }
  };

  const handleEndDateChange = (val: string) => {
    setEndDate(val);
    setActiveDurationPreset(null);
    setIsResultShown(false);
  };

  const handleSigningDateChange = (val: string) => {
    setSigningDate(val);
    setCustomLunarDate('');
    setCustomSolarDate('');
  };

  const handlePresetRange = (
    preset: 'this_month' | 'last_month' | 'next_month' | 'q1' | 's1' | 'm9' | 'full_year'
  ) => {
    setIsResultShown(false);
    const y = selectedYear || 2026;
    const m = selectedMonth || 1;
    let s = '';
    let e = '';

    if (preset === 'this_month') {
      const cy = selectedYear || new Date().getFullYear();
      const cm = selectedMonth || (new Date().getMonth() + 1);
      const range = calculateDateRange(cy, cm, 1);
      s = range.startDate;
      e = range.endDate;
      setActiveDurationPreset('this_month');
    } else if (preset === 'last_month') {
      let prevM = (selectedMonth || 1) - 1;
      let prevY = selectedYear || 2026;
      if (prevM < 1) {
        prevM = 12;
        prevY -= 1;
      }
      let duration = 1;
      if (activeDurationPreset === 'q1') duration = 3;
      else if (activeDurationPreset === 's1') duration = 6;
      else if (activeDurationPreset === 'm9') duration = 9;
      else if (activeDurationPreset === 'full_year') duration = 12;

      const range = calculateDateRange(prevY, prevM, duration);
      s = range.startDate;
      e = range.endDate;
      setSelectedYear(prevY);
      setSelectedMonth(prevM);
    } else if (preset === 'next_month') {
      let nextM = (selectedMonth || 1) + 1;
      let nextY = selectedYear || 2026;
      if (nextM > 12) {
        nextM = 1;
        nextY += 1;
      }
      let duration = 1;
      if (activeDurationPreset === 'q1') duration = 3;
      else if (activeDurationPreset === 's1') duration = 6;
      else if (activeDurationPreset === 'm9') duration = 9;
      else if (activeDurationPreset === 'full_year') duration = 12;

      const range = calculateDateRange(nextY, nextM, duration);
      s = range.startDate;
      e = range.endDate;
      setSelectedYear(nextY);
      setSelectedMonth(nextM);
    } else if (preset === 'q1') {
      const range = calculateDateRange(y, m, 3);
      s = range.startDate;
      e = range.endDate;
      setActiveDurationPreset('q1');
    } else if (preset === 's1') {
      const range = calculateDateRange(y, m, 6);
      s = range.startDate;
      e = range.endDate;
      setActiveDurationPreset('s1');
    } else if (preset === 'm9') {
      const range = calculateDateRange(y, m, 9);
      s = range.startDate;
      e = range.endDate;
      setActiveDurationPreset('m9');
    } else if (preset === 'full_year') {
      const range = calculateDateRange(y, m, 12);
      s = range.startDate;
      e = range.endDate;
      setActiveDurationPreset('full_year');
    }

    if (s && e) {
      setStartDate(s);
      setEndDate(e);
      // Auto-update signingDate to 1st of next month after e
      const eParts = e.split('-');
      const eY = parseInt(eParts[0], 10);
      const eM = parseInt(eParts[1], 10);
      const nextY = eM === 12 ? eY + 1 : eY;
      const nextM = eM === 12 ? 1 : eM + 1;
      setSigningDate(`${nextY}-${String(nextM).padStart(2, '0')}-01`);
    }
  };

  const isPresetActive = (preset: 'q1' | 's1' | 'm9' | 'full_year' | 'this_month') => {
    if (activeDurationPreset === preset) return true;
    if (preset === 'q1') {
      const r = calculateDateRange(selectedYear, selectedMonth, 3);
      return startDate === r.startDate && endDate === r.endDate;
    }
    if (preset === 's1') {
      const r = calculateDateRange(selectedYear, selectedMonth, 6);
      return startDate === r.startDate && endDate === r.endDate;
    }
    if (preset === 'm9') {
      const r = calculateDateRange(selectedYear, selectedMonth, 9);
      return startDate === r.startDate && endDate === r.endDate;
    }
    if (preset === 'full_year') {
      const r = calculateDateRange(selectedYear, selectedMonth, 12);
      return startDate === r.startDate && endDate === r.endDate;
    }
    return false;
  };

  // Category filter: 'all' (default: Sticker + cEA), 'sticker' (Sticker only), 'cea' (cEA only)
  const [selectedCategory, setSelectedCategory] = useState<'all' | 'sticker' | 'cea'>('all');

  // Helper to gather all stock records from props and localStorage
  const getAllStockRecords = useCallback((): StockRecord[] => {
    const all: StockRecord[] = [...(stockRecords || [])];
    try {
      const savedStock = localStorage.getItem('app_stock_records');
      if (savedStock) {
        const parsed = JSON.parse(savedStock);
        if (Array.isArray(parsed)) {
          parsed.forEach((r: StockRecord) => {
            if (!all.some((ex) => ex.id === r.id)) {
              all.push(r);
            }
          });
        }
      }
    } catch {}
    return all;
  }, [stockRecords]);

  // Check if a team has any sticker usage in the selected date range
  const checkTeamHasStickerUsage = useCallback((teamName: string, actualStart?: string, actualEnd?: string): { hasUsage: boolean; totalSheets: number } => {
    const allStock = getAllStockRecords();
    const targetTeam = matchTeamInList(teamName, Array.from(OFFICIAL_29_TEAMS)) || teamName;
    const normTarget = normalizeTeamName(teamName);

    let totalSheets = 0;

    allStock.forEach((rec) => {
      if (!rec || !rec.date) return;
      // Skip cEA / evisa records
      if (isCeaRecord(rec) || (rec as any).stockType === 'evisa' || (rec.sourceFrom || '').toLowerCase() === 'cea') {
        return;
      }

      // Check team match
      const rawRecTeam = resolveRecordTeamName(rec);
      if (!rawRecTeam) return;
      const cleanRaw = rawRecTeam.trim();
      const recTarget = matchTeamInList(cleanRaw, Array.from(OFFICIAL_29_TEAMS)) || cleanRaw;
      const recNorm = normalizeTeamName(cleanRaw);

      const isMatch =
        cleanRaw === teamName ||
        recTarget === targetTeam ||
        (recNorm && normTarget && (recNorm === normTarget || recNorm.includes(normTarget) || normTarget.includes(recNorm)));

      if (!isMatch) return;

      // Check operation: must be team usage
      const op = (rec.operationType || '').toLowerCase().trim();
      const src = (rec.sourceFrom || '').toLowerCase();
      const desc = ((rec as any).description || (rec as any).notes || '').toLowerCase();

      const isUseTeam =
        op === 'useteam' ||
        op.includes('useteam') ||
        op.includes('ប្រើប្រាស់') ||
        op.includes('ប្រើប្រាស') ||
        op.includes('ប្រើ') ||
        src.includes('ប្រើប្រាស់') ||
        desc.includes('ប្រើប្រាស់');

      if (!isUseTeam) return;

      // Exclude lost, damaged, transfer
      if (op.includes('បាត់') || op.includes('ខូច') || op.includes('ផ្ទេរ') || op.includes('បញ្ជូន')) return;

      // Check date range
      const recIso = normalizeDateToISO(rec.date);
      if (actualStart && recIso < actualStart) return;
      if (actualEnd && recIso > actualEnd) return;

      const qty =
        rec.totalSheets !== undefined && rec.totalSheets !== null
          ? rec.totalSheets
          : (rec.quantityBundles || (rec as any).quantity || 0);

      if (qty > 0) {
        totalSheets += qty;
      }
    });

    // Also check daily operations for sticker usage
    try {
      const savedDaily5 = localStorage.getItem('app_daily_team_operations_v5');
      if (savedDaily5) {
        const p5 = JSON.parse(savedDaily5);
        if (Array.isArray(p5)) {
          p5.forEach((dRec: any) => {
            if (!dRec) return;
            const dRecCat = String(dRec.categoryType || '').toLowerCase();
            if (dRecCat === 'cea') return; // skip cEA
            const recIso = normalizeDateToISO(dRec.date || '');
            if (actualStart && recIso < actualStart) return;
            if (actualEnd && recIso > actualEnd) return;

            const rawT = dRec.teamName || '';
            const normT = normalizeTeamName(rawT);
            if (rawT === teamName || normT === normTarget) {
              if (dRec.values) {
                Object.values(dRec.values).forEach((item: any) => {
                  let q = 0;
                  if (item?.entries && Array.isArray(item.entries)) {
                    q = item.entries.reduce((s: number, e: any) => s + (parseInt(e?.quantity || '', 10) || 0), 0);
                  } else {
                    q = parseInt(item?.quantity ?? item?.total ?? item?.sheets ?? item?.value ?? item ?? '', 10) || 0;
                  }
                  if (q > 0) totalSheets += q;
                });
              }
            }
          });
        }
      }
    } catch {}

    return { hasUsage: totalSheets > 0, totalSheets };
  }, [getAllStockRecords]);

  // Check if a team has cEA usage (Paper Approval / ក្រដាសអនុម័ត cEA)
  const checkTeamHasCeaUsage = useCallback((teamName: string, actualStart?: string, actualEnd?: string): { hasUsage: boolean; totalSheets: number } => {
    const allStock = getAllStockRecords();
    const targetTeam = matchTeamInList(teamName, Array.from(OFFICIAL_29_TEAMS)) || teamName;
    const normTarget = normalizeTeamName(teamName);

    let totalSheets = 0;

    // 1. Check stock records for cEA
    allStock.forEach((rec) => {
      if (!rec || !rec.date) return;
      const isCea = isCeaRecord(rec) || (rec as any).stockType === 'evisa' || (rec.sourceFrom || '').toLowerCase() === 'cea';
      if (!isCea) return;

      const rawRecTeam = resolveRecordTeamName(rec);
      if (!rawRecTeam) return;
      const cleanRaw = rawRecTeam.trim();
      const recTarget = matchTeamInList(cleanRaw, Array.from(OFFICIAL_29_TEAMS)) || cleanRaw;
      const recNorm = normalizeTeamName(cleanRaw);

      const isMatch =
        cleanRaw === teamName ||
        recTarget === targetTeam ||
        (recNorm && normTarget && (recNorm === normTarget || recNorm.includes(normTarget) || normTarget.includes(recNorm)));

      if (!isMatch) return;

      const op = (rec.operationType || '').toLowerCase().trim();
      if (op.includes('បាត់') || op.includes('ខូច') || op.includes('ផ្ទេរ') || op.includes('បញ្ជូន')) return;

      const recIso = normalizeDateToISO(rec.date);
      if (actualStart && recIso < actualStart) return;
      if (actualEnd && recIso > actualEnd) return;

      const qty =
        rec.totalSheets !== undefined && rec.totalSheets !== null
          ? rec.totalSheets
          : (rec.quantityBundles || (rec as any).quantity || 0);

      if (qty > 0) {
        totalSheets += qty;
      }
    });

    // 2. Check daily operations for cEA
    try {
      const savedDaily1 = localStorage.getItem('app_daily_team_records_v1');
      const savedDaily5 = localStorage.getItem('app_daily_team_operations_v5');
      const allDaily: any[] = [];
      if (savedDaily1) {
        try {
          const p1 = JSON.parse(savedDaily1);
          if (Array.isArray(p1)) allDaily.push(...p1);
        } catch {}
      }
      if (savedDaily5) {
        try {
          const p5 = JSON.parse(savedDaily5);
          if (Array.isArray(p5)) allDaily.push(...p5);
        } catch {}
      }

      allDaily.forEach((dRec: any) => {
        if (!dRec) return;
        const dRecCat = String(dRec.categoryType || '').toLowerCase();
        if (dRecCat !== 'cea') return;
        const recIso = normalizeDateToISO(dRec.date || '');
        if (actualStart && recIso < actualStart) return;
        if (actualEnd && recIso > actualEnd) return;

        const rawT = dRec.teamName || '';
        const normT = normalizeTeamName(rawT);
        if (rawT === teamName || normT === normTarget) {
          if (dRec.values) {
            Object.values(dRec.values).forEach((item: any) => {
              let q = 0;
              if (item?.entries && Array.isArray(item.entries)) {
                q = item.entries.reduce((s: number, e: any) => s + (parseInt(e?.quantity || '', 10) || 0), 0);
              } else {
                q = parseInt(item?.quantity ?? item?.total ?? item?.sheets ?? item?.value ?? item ?? '', 10) || 0;
              }
              if (q > 0) totalSheets += q;
            });
          }
        }
      });
    } catch {}

    // 3. Known cEA teams (such as Airport / Port teams in SAMPLE_CEA_TEAMS_Q1 or saved options)
    const isKnownCeaTeam = Object.keys(SAMPLE_CEA_TEAMS_Q1).some(
      (k) => normalizeTeamName(k) === normTarget || k === teamName
    );

    let hasConfiguredEvisa = false;
    try {
      const savedOpts = localStorage.getItem('team_type_options_v1');
      if (savedOpts) {
        const parsedOpts = JSON.parse(savedOpts);
        if (parsedOpts && typeof parsedOpts === 'object') {
          for (const [k, v] of Object.entries(parsedOpts)) {
            if (normalizeTeamName(k) === normTarget && (v as any)?.evisa === true) {
              hasConfiguredEvisa = true;
              break;
            }
          }
        }
      }
    } catch {}

    const hasUsage = totalSheets > 0 || isKnownCeaTeam || hasConfiguredEvisa;
    return { hasUsage, totalSheets };
  }, [getAllStockRecords]);

  // Evaluate Sticker and cEA usage for currently selected team
  const {
    hasStickerUsage,
    stickerSheets,
    hasCeaUsage,
    ceaSheets,
    hasEligibleUsage,
  } = useMemo(() => {
    const normStart = normalizeDateToISO(startDate);
    const normEnd = normalizeDateToISO(endDate);
    const actualStart = normStart && normEnd && normStart > normEnd ? normEnd : normStart;
    const actualEnd = normStart && normEnd && normStart > normEnd ? normStart : normEnd;

    const stk = checkTeamHasStickerUsage(selectedTeam, actualStart, actualEnd);
    const cea = checkTeamHasCeaUsage(selectedTeam, actualStart, actualEnd);

    let eligible = false;
    if (selectedCategory === 'sticker') {
      // Must have Sticker usage
      eligible = stk.hasUsage;
    } else if (selectedCategory === 'cea') {
      // Must have cEA usage (teams that don't use Sticker but use cEA show results here)
      eligible = cea.hasUsage;
    } else {
      // 'all' category: shows result if team uses either Sticker or cEA
      eligible = stk.hasUsage || cea.hasUsage;
    }

    return {
      hasStickerUsage: stk.hasUsage,
      stickerSheets: stk.totalSheets,
      hasCeaUsage: cea.hasUsage,
      ceaSheets: cea.totalSheets,
      hasEligibleUsage: eligible,
    };
  }, [checkTeamHasStickerUsage, checkTeamHasCeaUsage, selectedTeam, selectedCategory, startDate, endDate]);

  // Notice state for attempts to show results on a team without eligible usage
  const [noUsageNotice, setNoUsageNotice] = useState<boolean>(false);

  // Storage key for caching user edits
  const storageKey = useMemo(() => {
    const tNorm = normalizeTeamName(selectedTeam) || 'default';
    return `${keyPrefix}team_visa_issuance_stats_${tNorm}_${selectedCategory}_${startDate}_${endDate}`;
  }, [keyPrefix, selectedTeam, selectedCategory, startDate, endDate]);

  // Visa Numbers Data State
  const [visaCounts, setVisaCounts] = useState<Record<VisaTypeKey, number>>(() => {
    try {
      const saved = localStorage.getItem(storageKey);
      if (saved) {
        const parsed = JSON.parse(saved);
        if (parsed && typeof parsed === 'object') return parsed;
      }
      const tNorm = normalizeTeamName(selectedTeam) || 'default';
      const legacySaved = localStorage.getItem(`team_visa_issuance_stats_${tNorm}_${startDate}_${endDate}`);
      if (legacySaved) {
        const parsed = JSON.parse(legacySaved);
        if (parsed && typeof parsed === 'object') return parsed;
      }
    } catch {}
    return createEmptyVisaCounts();
  });

  // Track if results are revealed (starts as false so dashes - are shown first until user clicks Show Result)
  const [isResultShown, setIsResultShown] = useState<boolean>(false);
  const [isSavedNotice, setIsSavedNotice] = useState<boolean>(false);

  // Function to pull and aggregate from stockRecords, daily operations, and cEA records
  const calculateFromRecords = () => {
    const newCounts = createEmptyVisaCounts();

    // If the team does not have eligible usage for the selected category, do not show result
    if (!hasEligibleUsage) {
      setVisaCounts(newCounts);
      setIsResultShown(false);
      try {
        localStorage.removeItem(storageKey);
      } catch {}
      return newCounts;
    }

    const normStart = normalizeDateToISO(startDate);
    const normEnd = normalizeDateToISO(endDate);
    const actualStart = normStart && normEnd && normStart > normEnd ? normEnd : normStart;
    const actualEnd = normStart && normEnd && normStart > normEnd ? normStart : normEnd;

    const targetSelectedTeam = matchTeamInList(selectedTeam, Array.from(OFFICIAL_29_TEAMS));
    const normSelectedTeam = normalizeTeamName(selectedTeam);

    const checkTeamMatches = (rawRecTeam: string): boolean => {
      if (!selectedTeam || selectedTeam === 'ទាំងអស់') return true;
      if (!rawRecTeam) return false;
      const rawClean = rawRecTeam.trim();
      if (rawClean === selectedTeam) return true;

      const recTarget = matchTeamInList(rawClean, Array.from(OFFICIAL_29_TEAMS));
      if (recTarget && targetSelectedTeam && recTarget === targetSelectedTeam) return true;

      const recNorm = normalizeTeamName(rawClean);
      if (recNorm && normSelectedTeam) {
        if (recNorm === normSelectedTeam) return true;
        if (recNorm.includes(normSelectedTeam) || normSelectedTeam.includes(recNorm)) return true;
      }
      return false;
    };

    let foundAny = false;
    const existingUseKeySet = new Set<string>();

    // 1. Gather from stock records (Sticker & cEA)
    const allStock = getAllStockRecords();

    allStock.forEach((rec) => {
      if (!rec || !rec.date) return;
      const isCea = isCeaRecord(rec) || (rec as any).stockType === 'evisa' || (rec.sourceFrom || '').toLowerCase() === 'cea';
      if (selectedCategory === 'sticker' && isCea) return;
      if (selectedCategory === 'cea' && !isCea) return;

      const op = (rec.operationType || '').toLowerCase().trim();
      const src = (rec.sourceFrom || '').toLowerCase();
      const desc = ((rec as any).description || (rec as any).notes || '').toLowerCase();

      const isUseTeam =
        op === 'useteam' ||
        op.includes('useteam') ||
        op.includes('ប្រើប្រាស់') ||
        op.includes('ប្រើប្រាស') ||
        op.includes('ប្រើ') ||
        src.includes('ប្រើប្រាស់') ||
        desc.includes('ប្រើប្រាស់') ||
        src === 'cea' ||
        src.includes('cea') ||
        (rec.stockType === 'evisa' && (op === 'useteam' || op.includes('use')));

      if (!isUseTeam) return;
      if (op.includes('បាត់') || op.includes('ខូច') || op.includes('ផ្ទេរ') || op.includes('បញ្ជូន')) return;

      const rawTeam = resolveRecordTeamName(rec);
      if (!checkTeamMatches(rawTeam)) return;

      const recIso = normalizeDateToISO(rec.date);
      if (actualStart && recIso < actualStart) return;
      if (actualEnd && recIso > actualEnd) return;

      const vtNorm = normalizeVisaType(rec.visaType);
      const matchedVt =
        VISA_TYPE_KEYS.find((v) => v === vtNorm) ||
        VISA_TYPE_KEYS.find((v) => v === (rec.visaType || '').toUpperCase().trim());

      const qty =
        rec.totalSheets !== undefined && rec.totalSheets !== null
          ? rec.totalSheets
          : (rec.quantityBundles || (rec as any).quantity || 0);

      if (qty > 0 && matchedVt) {
        foundAny = true;
        newCounts[matchedVt] = (newCounts[matchedVt] || 0) + qty;
        const normT = normalizeTeamName(rawTeam);
        existingUseKeySet.add(`${recIso}_${normT}_${matchedVt}_${isCea ? 'cea' : 'sticker'}`);
      }
    });

    // 2. Gather from app_daily_team_operations_v5 and app_daily_team_records_v1
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

      const processedDailyIds = new Set<string>();

      allDailyRecords.forEach((dRec: any) => {
        if (!dRec || !dRec.id || processedDailyIds.has(dRec.id)) return;
        processedDailyIds.add(dRec.id);

        const dRecCat = String(dRec.categoryType || '').toLowerCase();
        const isCea = dRecCat === 'cea';
        if (selectedCategory === 'sticker' && isCea) return;
        if (selectedCategory === 'cea' && !isCea) return;

        const recIso = normalizeDateToISO(dRec.date || '');
        if (!dRec.values || (actualStart && recIso < actualStart) || (actualEnd && recIso > actualEnd)) return;

        const rawTeamName = dRec.teamName || '';
        if (!checkTeamMatches(rawTeamName)) return;

        VISA_TYPE_KEYS.forEach((vt) => {
          const item = dRec.values[vt];
          let qty = 0;
          if (item) {
            if (Array.isArray(item.entries) && item.entries.length > 0) {
              qty = item.entries.reduce((sum: number, e: any) => sum + (parseInt(e?.quantity || '', 10) || 0), 0);
            } else {
              qty = parseInt(item?.quantity ?? item?.total ?? item?.sheets ?? item?.value ?? item ?? '', 10) || 0;
            }
          }
          if (qty <= 0) return;

          const normT = normalizeTeamName(rawTeamName);
          const lookupKey = `${recIso}_${normT}_${vt}_${isCea ? 'cea' : 'sticker'}`;
          if (!existingUseKeySet.has(lookupKey)) {
            foundAny = true;
            newCounts[vt] = (newCounts[vt] || 0) + qty;
          }
        });
      });
    } catch (e) {
      console.warn('Error reading daily operations', e);
    }

    // 3. No fallback baseline for cEA teams if no records entered in database yet to keep it empty/0 as requested by the user
    // (Removed SAMPLE_CEA_TEAMS_Q1 fallback)

    setVisaCounts(newCounts);
    return newCounts;
  };

  // Synchronize: auto-load saved custom data or auto-calculate dynamically whenever team, category or date range changes
  useEffect(() => {
    if (!hasEligibleUsage) {
      setVisaCounts(createEmptyVisaCounts());
      setIsResultShown(false);
      try {
        localStorage.removeItem(storageKey);
      } catch {}
      return;
    }
    try {
      const saved = localStorage.getItem(storageKey);
      if (saved) {
        const parsed = JSON.parse(saved);
        if (parsed && typeof parsed === 'object') {
          setVisaCounts(parsed);
          return;
        }
      }
    } catch (e) {}
    calculateFromRecords();
  }, [storageKey, selectedTeam, selectedCategory, startDate, endDate, stockRecords, hasEligibleUsage]);

  // Trigger calculate / load data and reveal results
  const handleShowResult = () => {
    if (!hasEligibleUsage) {
      setIsResultShown(false);
      setNoUsageNotice(true);
      setTimeout(() => setNoUsageNotice(false), 4500);
      return;
    }
    // Recalculate fresh counts from system records
    calculateFromRecords();
    setIsResultShown(true);
  };

  // Clear / reset display back to (-)
  const handleClearToDash = () => {
    setIsResultShown(false);
  };

  // Save current report data / state
  const handleSaveData = () => {
    if (!hasEligibleUsage) {
      setNoUsageNotice(true);
      setTimeout(() => setNoUsageNotice(false), 4500);
      return;
    }
    try {
      localStorage.setItem(storageKey, JSON.stringify(visaCounts));
      setIsSavedNotice(true);
      setIsResultShown(true);
      setTimeout(() => setIsSavedNotice(false), 2500);
    } catch (e) {
      console.error('Failed to save data', e);
    }
  };

  // Reset to PDF sample data
  const handleLoadPdfSample = () => {
    setSelectedTeam('អាកាស តេជោ');
    setSelectedYear(2018);
    setSelectedMonth(12);
    setStartDate('2018-12-01');
    setEndDate('2019-02-28');
    setActiveDurationPreset('q1');
    setSigningDate('2019-03-01');
    setCustomLocation('');
    setTeamHeaderName('ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអាកាសយានដ្ឋានអន្តរជាតិតេជោ');
    setChiefTitle('ប្រធានក្រុមផ្តល់ទិដ្ឋាការ');
    setVisaCounts({ ...SAMPLE_Q1_2026_VALUES });
    setIsResultShown(true);

    try {
      const tNorm = normalizeTeamName('អាកាស តេជោ') || 'techo';
      const key1 = `team_visa_issuance_stats_${tNorm}_2018-12-01_2019-02-28`;
      const key2 = `team_visa_issuance_stats_air_techo_2018-12-01_2019-02-28`;
      localStorage.setItem(key1, JSON.stringify(SAMPLE_Q1_2026_VALUES));
      localStorage.setItem(key2, JSON.stringify(SAMPLE_Q1_2026_VALUES));
    } catch (e) {}
  };

  // Grand Total Calculation
  const totalTaxable = useMemo(() => {
    return TAXABLE_VISA_TYPES.reduce((sum, item) => sum + (visaCounts[item.key] || 0), 0);
  }, [visaCounts]);

  const totalNonTaxable = useMemo(() => {
    return NON_TAXABLE_VISA_TYPES.reduce((sum, item) => sum + (visaCounts[item.key] || 0), 0);
  }, [visaCounts]);

  const grandTotal = useMemo(() => {
    return totalTaxable + totalNonTaxable;
  }, [totalTaxable, totalNonTaxable]);

  // Tacteing ornament settings
  const [tacteingSettings, setTacteingSettings] = useState<{ type: TacteingType; customImage: string | null }>(
    () => getSavedTacteingSettings()
  );

  useEffect(() => {
    const handleUpdate = () => {
      setTacteingSettings(getSavedTacteingSettings());
    };
    window.addEventListener('tacteing_settings_updated', handleUpdate);
    window.addEventListener('storage', handleUpdate);
    return () => {
      window.removeEventListener('tacteing_settings_updated', handleUpdate);
      window.removeEventListener('storage', handleUpdate);
    };
  }, []);

  // Header & Content Customizations
  const [ministryName, setMinistryName] = useState<string>('ក្រសួងមហាផ្ទៃ');
  const [deptName, setDeptName] = useState<string>('អគ្គនាយកដ្ឋានអន្តោប្រវេសន៍');
  const [firstGateDeptName, setFirstGateDeptName] = useState<string>('នាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត');
  const [officeName, setOfficeName] = useState<string>('ការិយាល័យទិដ្ឋាការចូល');
  const [teamHeaderName, setTeamHeaderName] = useState<string>(() => {
    const initialTeam = assignedTeam && assignedTeam.trim() ? assignedTeam : 'ព្រំដែន បាវិត';
    const norm = normalizeTeamName(initialTeam);
    if (norm === normalizeTeamName('អាកាស តេជោ')) return 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអាកាសយានដ្ឋានអន្តរជាតិតេជោ';
    if (norm === normalizeTeamName('អាកាស សៀមរាប')) return 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអាកាសយានដ្ឋានអន្តរជាតិសៀមរាប';
    if (norm === normalizeTeamName('អាកាស ព្រះសីហនុ')) return 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអាកាសយានដ្ឋានអន្តរជាតិព្រះសីហនុ';
    if (norm === normalizeTeamName('កំពង់ផែ ព្រះសីហនុ')) return 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារកំពង់ផែអន្តរជាតិព្រះសីហនុ';
    return formatReportTeamName(initialTeam);
  });

  // Center Salutation
  const [salutationTop, setSalutationTop] = useState<string>('សូមគោរពជូន');
  const [salutationRecipient, setSalutationRecipient] = useState<string>(() => {
    try {
      const saved = localStorage.getItem(`${keyPrefix}salutation_recipient`);
      if (saved) {
        // Automatically migrate legacy/placeholder "លោកប្រធានសាខា" to "លោកវរសេនីយ៍ឯក ប្រធានសាខាប្រមូលចំណូល"
        if (isBranchChiefMode && (saved === 'លោកប្រធានសាខា' || saved === 'ប្រធានសាខា')) {
          return 'លោកវរសេនីយ៍ឯក ប្រធានសាខាប្រមូលចំណូល';
        }
        return saved;
      }
    } catch {}
    return isBranchChiefMode ? 'លោកវរសេនីយ៍ឯក ប្រធានសាខាប្រមូលចំណូល' : 'លោកវរសេនីយ៍ទោ នាយការិយាល័យទិដ្ឋាការចូល';
  });

  // Location string used in Subject (optional, e.g. "នៅ...")
  const [customLocation, setCustomLocation] = useState<string>('');

  // Custom override for full subject if user wants to freely edit
  const [isCustomSubject, setIsCustomSubject] = useState<boolean>(false);
  const [customSubjectText, setCustomSubjectText] = useState<string>('');

  const defaultSubjectText = useMemo(() => {
    const locPart = customLocation && customLocation.trim() ? `នៅ${customLocation.trim()} ` : '';
    return `របាយការណ៍ស្តីពី ចំនួនភ្ញៀវបរទេស ដែលបានមកស្នើសុំទិដ្ឋាការចូលព្រះរាជាណាចក្រកម្ពុជា ${locPart}${periodLabelKh} ${periodDetailsKh}`;
  }, [customLocation, periodLabelKh, periodDetailsKh]);

  const effectiveSubjectText = isCustomSubject && customSubjectText.trim() ? customSubjectText : defaultSubjectText;

  // Font size setting (default 12pt as requested)
  const [fontSizePt, setFontSizePt] = useState<number>(() => {
    try {
      const saved = localStorage.getItem(`${keyPrefix}team_visa_stats_font_size`);
      if (saved) {
        const val = parseFloat(saved);
        if (!isNaN(val) && val >= 8 && val <= 20) {
          if (val === 11.5) return 12; // Upgrade legacy 11.5pt to 12pt
          return val;
        }
      }
    } catch {}
    return 12; // Default 12pt
  });

  // Line spacing setting (គម្លាតអក្សរពីលើចុះក្រោម / Vertical Line Spacing)
  const [lineSpacing, setLineSpacing] = useState<number>(() => {
    try {
      const saved = localStorage.getItem(`${keyPrefix}team_visa_stats_line_spacing`);
      if (saved) {
        const val = parseFloat(saved);
        // Automatically adopt 1.25 for users who had looser legacy values (>= 1.35)
        if (!isNaN(val) && val >= 0.9 && val <= 2.5 && val <= 1.30) return val;
      }
    } catch {}
    return 1.25; // Default 1.25 for condensed, elegant vertical spacing
  });

  // Section titles (ក. និង ខ.)
  const [sectionATitle, setSectionATitle] = useState<string>(() => {
    try {
      const saved = localStorage.getItem(`${keyPrefix}team_visa_stats_section_a_title`);
      if (saved) {
        if (saved.includes('អាការ')) {
          const corrected = saved.replace(/អាការ/g, 'អាករ');
          localStorage.setItem(`${keyPrefix}team_visa_stats_section_a_title`, corrected);
          return corrected;
        }
        return saved;
      }
    } catch {}
    return 'ក. ប្រភេទទិដ្ឋាការបង់អាករ';
  });

  const [sectionBTitle, setSectionBTitle] = useState<string>(() => {
    try {
      const saved = localStorage.getItem(`${keyPrefix}team_visa_stats_section_b_title`);
      if (saved) {
        if (saved.includes('អាការ')) {
          const corrected = saved.replace(/អាការ/g, 'អាករ');
          localStorage.setItem(`${keyPrefix}team_visa_stats_section_b_title`, corrected);
          return corrected;
        }
        return saved;
      }
    } catch {}
    return 'ខ. ប្រភេទទិដ្ឋាការមិនបង់អាករ';
  });

  // Tab indent setting (matching the width of 'កម្មវត្ថុ ៖' so that 'ក.', 'ខ.', 'គ.' align with 'របាយការណ៍ស្តីពី')
  const [tabIndentPx, setTabIndentPx] = useState<number>(() => {
    try {
      const saved = localStorage.getItem(`${keyPrefix}team_visa_stats_tab_indent`);
      if (saved) {
        const val = parseFloat(saved);
        if (!isNaN(val) && val >= 0) return val;
      }
    } catch {}
    return 88;
  });

  // Remarks / Section គ
  const [remarksNote, setRemarksNote] = useState<string>('គ្មាន ។');

  // Polite closing request
  const defaultClosing1 = isBranchChiefMode
    ? 'អាស្រ័យដូចបានជម្រាបជូនខាងលើ សូមលោកវរសេនីយ៍ឯក ប្រធានសាខាប្រមូលចំណូល មេត្តាពិនិត្យ និងជ្រាបជារបាយការណ៍ តាមការគួរ។'
    : 'អាស្រ័យដូចបានជម្រាបជូនខាងលើ សូមលោកវរសេនីយ៍ទោ នាយការិយាល័យទិដ្ឋាការចូល មេត្តាពិនិត្យ និងជ្រាបជារបាយការណ៍ ដ៏ខ្ពង់ខ្ពស់។';

  const defaultClosing2 = isBranchChiefMode
    ? 'សូម លោកប្រធានសាខា ទទួលនូវការរាប់អាន អំពីខ្ញុំ ។'
    : 'សូម លោកនាយការិយាល័យ ទទួលនូវការគោរពដ៏ខ្ពង់ខ្ពស់ អំពីយើងខ្ញុំ ។';

  const [closingParagraph1, setClosingParagraph1] = useState<string>(() => {
    try {
      const saved = localStorage.getItem(`${keyPrefix}closing_paragraph_1`);
      if (saved) {
        if (isBranchChiefMode && saved.includes('លោកវរសេនីយ៍ទោ')) {
          return defaultClosing1;
        }
        return saved;
      }
    } catch {}
    return defaultClosing1;
  });

  const [closingParagraph2, setClosingParagraph2] = useState<string>(() => {
    try {
      const saved = localStorage.getItem(`${keyPrefix}closing_paragraph_2`);
      if (saved) {
        if (isBranchChiefMode && saved.includes('លោកនាយការិយាល័យ')) {
          return defaultClosing2;
        }
        return saved;
      }
    } catch {}
    return defaultClosing2;
  });

  // Helper to format titles with Khmer OS Muol Light in closing text
  const renderClosingTextWithMuol = (text: string) => {
    const targets = [
      'លោកវរសេនីយ៍ទោ នាយការិយាល័យទិដ្ឋាការចូល',
      'លោកវរសេនីយ៍ឯក ប្រធានសាខាប្រមូលចំណូល',
      'លោកនាយការិយាល័យ',
      'លោកប្រធានសាខា',
      salutationRecipient.trim(),
    ].filter(Boolean);

    const uniqueTargets = Array.from(new Set(targets)).sort((a, b) => b.length - a.length);
    if (uniqueTargets.length === 0) return text;

    const regex = new RegExp(`(${uniqueTargets.map((t) => t.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|')})`, 'g');
    const parts = text.split(regex);

    return parts.map((part, idx) => {
      if (uniqueTargets.includes(part)) {
        return (
          <span
            key={idx}
            className="font-moul"
            style={{
              fontFamily: "'Khmer OS Muol Light', 'Khmer OS Moul Light', 'Khmer OS Muol', 'Moul', serif",
              fontSize: `${fontSizePt}pt`,
            }}
          >
            {part}
          </span>
        );
      }
      return <span key={idx}>{part}</span>;
    });
  };

  const formatClosingHtmlWithMuol = (text: string) => {
    const targets = [
      'លោកវរសេនីយ៍ទោ នាយការិយាល័យទិដ្ឋាការចូល',
      'លោកវរសេនីយ៍ឯក ប្រធានសាខាប្រមូលចំណូល',
      'លោកនាយការិយាល័យ',
      'លោកប្រធានសាខា',
      salutationRecipient.trim(),
    ].filter(Boolean);

    const uniqueTargets = Array.from(new Set(targets)).sort((a, b) => b.length - a.length);
    let result = text;
    for (const target of uniqueTargets) {
      if (target) {
        const regex = new RegExp(target.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g');
        result = result.replace(
          regex,
          `<span class="font-moul" style="font-family: 'Khmer OS Muol Light', 'Khmer OS Moul Light', 'Khmer OS Muol', 'Moul', serif; font-size: ${fontSizePt}pt;">${target}</span>`
        );
      }
    }
    return result;
  };

  const defaultChiefTitle = 'ប្រធានក្រុមផ្តល់ទិដ្ឋាការ';

  // Signing date & location
  const [signingDate, setSigningDate] = useState<string>('2026-04-01');
  const [signingLocation, setSigningLocation] = useState<string>('រាជធានីភ្នំពេញ');
  const [chiefTitle, setChiefTitle] = useState<string>(() => {
    try {
      const saved = localStorage.getItem(`${keyPrefix}chief_title`);
      if (saved) {
        if (isBranchChiefMode && (saved === 'ប្រធានសាខា' || saved === 'លោកប្រធានសាខា')) {
          return 'ប្រធានក្រុមផ្តល់ទិដ្ឋាការ';
        }
        if (saved === 'ប្រធានក្រុមផ្តល់ទិដ្ឋាការទី១') {
          return 'ប្រធានក្រុមផ្តល់ទិដ្ឋាការ';
        }
        return saved;
      }
    } catch {}
    return defaultChiefTitle;
  });
  const [chiefName, setChiefName] = useState<string>(() => {
    try {
      return localStorage.getItem(`${keyPrefix}chief_name`) || '';
    } catch {
      return '';
    }
  });
  const [customLunarDate, setCustomLunarDate] = useState<string>('');
  const [customSolarDate, setCustomSolarDate] = useState<string>('');

  // Auto-calculated Lunar Date
  const khmerLunarDate = useMemo(() => {
    return getKhmerLunarDate(signingDate) || 'ថ្ងៃពុធ ១៤កើត ខែចេត្រ ឆ្នាំម្សាញ់ សប្តស័ក ព.ស ២៥៦៩';
  }, [signingDate]);

  // Solar Date
  const khmerSolarDate = useMemo(() => {
    const parts = getKhmerSolarParts(signingDate);
    return `${signingLocation}, ថ្ងៃទី ${parts.khmerDay} ខែ ${parts.khmerMonth} ឆ្នាំ ${parts.khmerYear}`;
  }, [signingDate, signingLocation]);

  const displayLunarDate = customLunarDate.trim() ? customLunarDate : khmerLunarDate;
  const displaySolarDate = customSolarDate.trim() ? customSolarDate : khmerSolarDate;

  // Lunar Date on ONE single line (unless explicitly multiline with \n)
  const lunarLines = useMemo(() => {
    const text = displayLunarDate.trim();
    if (!text) return [];
    if (text.includes('\n')) return text.split('\n').map((s) => s.trim()).filter(Boolean);
    return [text];
  }, [displayLunarDate]);

  // Update team header automatically when selectedTeam changes
  useEffect(() => {
    const norm = normalizeTeamName(selectedTeam);
    if (norm === normalizeTeamName('អាកាស តេជោ')) {
      setTeamHeaderName('ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអាកាសយានដ្ឋានអន្តរជាតិតេជោ');
    } else if (norm === normalizeTeamName('អាកាស សៀមរាប')) {
      setTeamHeaderName('ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអាកាសយានដ្ឋានអន្តរជាតិសៀមរាប');
    } else if (norm === normalizeTeamName('អាកាស ព្រះសីហនុ')) {
      setTeamHeaderName('ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអាកាសយានដ្ឋានអន្តរជាតិព្រះសីហនុ');
    } else if (norm === normalizeTeamName('កំពង់ផែ ព្រះសីហនុ')) {
      setTeamHeaderName('ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារកំពង់ផែអន្តរជាតិព្រះសីហនុ');
    } else {
      let full = '';
      if (categories) {
        const robokList = categories.visaTeamsRobok || [];
        const fullList = categories.visaTeams || [];
        const rIdx = robokList.findIndex((item) => {
          if (!item?.name) return false;
          const rNorm = normalizeTeamName(item.name);
          return rNorm === norm || item.name.trim() === selectedTeam.trim();
        });
        if (rIdx !== -1 && fullList[rIdx]?.name && !/^ក្រុមទី\s*\d+$/i.test(fullList[rIdx].name.trim())) {
          full = fullList[rIdx].name.trim();
        }
      }
      if (!full) {
        full = formatReportTeamName(selectedTeam);
      }
      setTeamHeaderName(full || `ក្រុមផ្តល់ទិដ្ឋាការ ${selectedTeam}`);
    }
  }, [selectedTeam, categories]);

  // Settings drawer / panel
  const [showSettings, setShowSettings] = useState<boolean>(false);

  // Page Margins in cm (Default left: 2.80cm, right: 1.30cm, top: 1.00cm, bottom: 0.50cm)
  const [pageLeftMargin, setPageLeftMargin] = useState<number>(() => {
    try {
      const saved = localStorage.getItem(`${keyPrefix}team_visa_report_margin_left`);
      if (saved !== null) {
        const parsed = parseFloat(saved);
        if (!isNaN(parsed)) return parsed;
      }
    } catch {}
    return 2.80;
  });

  const [pageRightMargin, setPageRightMargin] = useState<number>(() => {
    try {
      const saved = localStorage.getItem(`${keyPrefix}team_visa_report_margin_right`);
      if (saved !== null) {
        const parsed = parseFloat(saved);
        if (!isNaN(parsed)) return parsed;
      }
    } catch {}
    return 1.30;
  });

  const [pageTopMargin, setPageTopMargin] = useState<number>(() => {
    try {
      const saved = localStorage.getItem(`${keyPrefix}team_visa_report_margin_top`);
      if (saved !== null) {
        const parsed = parseFloat(saved);
        if (!isNaN(parsed) && parsed >= 0.2) return parsed;
      }
    } catch {}
    return 0.80;
  });

  const [pageBottomMargin, setPageBottomMargin] = useState<number>(() => {
    try {
      const saved = localStorage.getItem(`${keyPrefix}team_visa_report_margin_bottom`);
      if (saved !== null) {
        const parsed = parseFloat(saved);
        if (!isNaN(parsed)) return parsed;
      }
    } catch {}
    return 0.50;
  });

  const [paragraphShiftY, setParagraphShiftY] = useState<number>(() => {
    try {
      const saved = localStorage.getItem(`${keyPrefix}team_visa_report_paragraph_shift_y`);
      if (saved !== null) {
        const parsed = parseFloat(saved);
        if (!isNaN(parsed) && parsed !== 16 && parsed !== 6) return parsed;
      }
    } catch {}
    return 2; // Default 2pt (compact and clean)
  });

  const [signatureShiftY, setSignatureShiftY] = useState<number>(() => {
    try {
      const saved = localStorage.getItem(`${keyPrefix}team_visa_report_signature_shift_y`);
      if (saved !== null) {
        const parsed = parseFloat(saved);
        if (!isNaN(parsed) && parsed !== 6) return parsed;
      }
    } catch {}
    return 4; // Default 4pt (clean and raised gap)
  });

  // Persist margins and line spacing to localStorage
  useEffect(() => {
    try {
      localStorage.setItem(`${keyPrefix}team_visa_report_margin_left`, String(pageLeftMargin));
      localStorage.setItem(`${keyPrefix}team_visa_report_margin_right`, String(pageRightMargin));
      localStorage.setItem(`${keyPrefix}team_visa_report_margin_top`, String(pageTopMargin));
      localStorage.setItem(`${keyPrefix}team_visa_report_margin_bottom`, String(pageBottomMargin));
      localStorage.setItem(`${keyPrefix}team_visa_report_paragraph_shift_y`, String(paragraphShiftY));
      localStorage.setItem(`${keyPrefix}team_visa_report_signature_shift_y`, String(signatureShiftY));
      localStorage.setItem(`${keyPrefix}team_visa_stats_line_spacing`, String(lineSpacing));
      localStorage.setItem(`${keyPrefix}team_visa_stats_section_a_title`, sectionATitle);
      localStorage.setItem(`${keyPrefix}team_visa_stats_section_b_title`, sectionBTitle);
      localStorage.setItem(`${keyPrefix}chief_title`, chiefTitle);
      localStorage.setItem(`${keyPrefix}chief_name`, chiefName);
      localStorage.setItem(`${keyPrefix}salutation_recipient`, salutationRecipient);
      localStorage.setItem(`${keyPrefix}closing_paragraph_1`, closingParagraph1);
      localStorage.setItem(`${keyPrefix}closing_paragraph_2`, closingParagraph2);
    } catch {}
  }, [
    keyPrefix,
    pageLeftMargin,
    pageRightMargin,
    pageTopMargin,
    pageBottomMargin,
    paragraphShiftY,
    signatureShiftY,
    lineSpacing,
    sectionATitle,
    sectionBTitle,
    chiefTitle,
    chiefName,
    salutationRecipient,
    closingParagraph1,
    closingParagraph2,
  ]);

  // Zoom Controls
  const [zoomScale, setZoomScale] = useState<number>(100);

  // Synchronize with global workspace scale if not fit
  useEffect(() => {
    if (scaleMode && scaleMode !== 'fit') {
      const numeric = parseInt(scaleMode.replace('%', ''), 10);
      if (!isNaN(numeric)) {
        setZoomScale(numeric);
      }
    }
  }, [scaleMode]);
  const [isExporting, setIsExporting] = useState<boolean>(false);

  // Printable container ref
  const reportContainerRef = useRef<HTMLDivElement>(null);

  // Print function
  const handlePrint = () => {
    printA4Document('team-visa-issuance-stats-report-pdf', {
      orientation: 'portrait',
      documentTitle: `${reportMainTitle}_${(selectedTeam || 'team').replace(/[\/\s]+/g, '_')}`,
    });
  };

  // PDF Export
  const handleExportPdf = async () => {
    if (!reportContainerRef.current) return;
    setIsExporting(true);
    const safeTeam = (selectedTeam || 'team').replace(/[\/\s]+/g, '_');
    const fileName = `${reportMainTitle}_${safeTeam}_${startDate}_${endDate}.pdf`;

    const oldZoom = zoomScale;
    setZoomScale(100);

    try {
      if (document.fonts && document.fonts.ready) {
        await document.fonts.ready;
      }
      await new Promise((r) => setTimeout(r, 150));

      await exportElementToPdf(
        reportContainerRef.current,
        fileName,
        { pixelRatio: 3, fitSinglePage: true, orientation: 'portrait' }
      );
    } catch (e) {
      console.error('Export PDF error, attempting fallback:', e);
      try {
        const element = reportContainerRef.current;
        const canvas = await html2canvas(element, {
          scale: 3,
          useCORS: true,
          allowTaint: true,
          backgroundColor: '#ffffff',
          windowWidth: 1200,
          windowHeight: 1600,
          onclone: (clonedDoc) => {
            sanitizeDocumentForHtml2Canvas(clonedDoc, 'team-visa-issuance-stats-report-pdf');
            const clonedEl = clonedDoc.getElementById('team-visa-issuance-stats-report-pdf');
            if (clonedEl) {
              clonedEl.style.width = '210mm';
              clonedEl.style.minWidth = '210mm';
              clonedEl.style.maxWidth = '210mm';
              clonedEl.style.boxShadow = 'none';
              clonedEl.style.backgroundColor = '#ffffff';
            }
          },
        });

        const imgData = canvas.toDataURL('image/jpeg', 0.98);
        const pdf = new jsPDF({
          orientation: 'portrait',
          unit: 'mm',
          format: 'a4',
          compress: true,
        });
        pdf.addImage(imgData, 'JPEG', 0, 0, 210, 297, undefined, 'FAST');
        pdf.save(fileName);
      } catch (err) {
        console.error('Fallback export failed:', err);
      }
    } finally {
      setZoomScale(oldZoom);
      setIsExporting(false);
    }
  };

  // Word Export (.doc)
  const handleExportWord = () => {
    const safeTeam = (selectedTeam || 'team').replace(/[\/\s]+/g, '_');
    const fileName = `${reportMainTitle}_${safeTeam}_${startDate}_${endDate}.doc`;

    const docContent = `
      <html xmlns:o='urn:schemas-microsoft-com:office:office' xmlns:w='urn:schemas-microsoft-com:office:word' xmlns='http://www.w3.org/TR/REC-html40'>
      <head>
        <meta charset='utf-8'>
        <title>${reportMainTitle}</title>
        <style>
          @page {
            size: 21.0cm 29.7cm;
            margin: ${pageTopMargin}cm ${pageRightMargin}cm ${pageBottomMargin}cm ${pageLeftMargin}cm;
            mso-page-orientation: portrait;
          }
          body {
            font-family: 'Khmer OS Siemreap', 'Siemreap', sans-serif;
            font-size: ${fontSizePt}pt;
            line-height: ${lineSpacing || 1.6};
            color: #000;
          }
          .font-moul {
            font-family: 'Khmer Mool1', 'Khmer Mool 1', 'Khmer Mool', 'Khmer OS Muol Light', 'Khmer OS Moul Light', 'Khmer OS Muol', 'Moul', serif;
            font-weight: normal;
          }
          .font-siemreap {
            font-family: 'Khmer OS Siemreap', 'Siemreap', sans-serif;
          }
          .font-times {
            font-family: 'Times New Roman', Times, serif;
          }
          .text-blue {
            color: #002060;
          }
          .header-table {
            width: 100%;
            border-collapse: collapse;
            margin-bottom: 8pt;
          }
          .items-table {
            width: 100%;
            border-collapse: collapse;
            margin-top: 4pt;
            margin-bottom: 6pt;
          }
          .items-table td {
            padding: 2.5pt 4pt;
            font-size: ${fontSizePt}pt;
            line-height: ${lineSpacing || 1.6};
          }
          .text-right { text-align: right; }
          .text-center { text-align: center; }
          .bold { font-weight: bold; }
        </style>
      </head>
      <body>
        <!-- Header -->
        <table class="header-table" style="line-height: ${Math.max(0.95, parseFloat((lineSpacing * 0.85).toFixed(2)))}; margin-bottom: ${Math.max(2, Math.round((lineSpacing - 0.7) * 4))}pt;">
          <tr>
            <td style="width: 50%; vertical-align: top; line-height: ${Math.max(0.95, parseFloat((lineSpacing * 0.85).toFixed(2)))};">
              <div style="visibility: hidden; font-size: 12pt; line-height: inherit;" class="font-moul">ព្រះរាជាណាចក្រកម្ពុជា</div>
              <div class="font-moul text-blue" style="font-size: 12pt; line-height: inherit;">${ministryName}</div>
              <div class="font-moul text-blue" style="font-size: 12pt; line-height: inherit;">${deptName}</div>
              <div class="font-moul text-blue" style="font-size: 12pt; line-height: inherit;">${firstGateDeptName}</div>
              <div class="font-moul text-blue" style="font-size: 12pt; line-height: inherit;">${officeName}</div>
              <div class="bold text-blue" style="font-size: 10pt; line-height: inherit;">${teamHeaderName}</div>
              <div style="font-size: 9pt; color: #002060; margin-top: 1pt; text-align: center; width: 140px;">─────── ❖ ───────</div>
            </td>
            <td style="width: 50%; vertical-align: top; text-align: center; line-height: ${Math.max(0.95, parseFloat((lineSpacing * 0.85).toFixed(2)))};">
              <div class="font-moul text-blue" style="font-size: 12pt; line-height: inherit;">ព្រះរាជាណាចក្រកម្ពុជា</div>
              <div class="font-moul text-blue" style="font-size: 12pt; line-height: inherit;">ជាតិ សាសនា ព្រះមហាក្សត្រ</div>
              <div style="font-size: 9pt; color: #002060; margin-top: 1pt;">─────── ❖ ───────</div>
            </td>
          </tr>
        </table>

        <!-- Salutation -->
        <div class="text-center" style="margin-top: ${Math.max(1, Math.round((lineSpacing - 0.7) * 3))}pt; margin-bottom: ${Math.max(2, Math.round((lineSpacing - 0.7) * 4))}pt; line-height: ${Math.max(1.0, parseFloat((lineSpacing * 0.88).toFixed(2)))};">
          <div class="font-moul text-blue" style="font-size: 14pt; line-height: inherit;">${salutationTop}</div>
          <div class="font-moul text-blue" style="font-size: 11.5pt; line-height: inherit;">${salutationRecipient}</div>
        </div>

        <!-- Subject -->
        <table style="width: 100%; margin-top: 5pt; margin-bottom: 6pt; border-collapse: collapse; line-height: ${Math.max(lineSpacing * 1.35, 1.72)};">
          <tr>
            <td style="width: ${tabIndentPx}px; vertical-align: top; font-weight: bold; white-space: nowrap; font-size: ${fontSizePt}pt;" class="font-moul">កម្មវត្ថុ ៖</td>
            <td style="text-align: justify; vertical-align: top; font-size: ${fontSizePt}pt; line-height: ${Math.max(lineSpacing * 1.35, 1.72)};">${effectiveSubjectText}</td>
          </tr>
        </table>

        <!-- Body content with Tab indent aligned with របាយការណ៍ស្តីពី -->
        <div style="margin-left: ${tabIndentPx}px;">
          <!-- Section A: Taxable -->
          <div class="bold font-moul" style="font-size: ${fontSizePt}pt; margin-top: 4pt; margin-bottom: 2pt;">
            ${sectionATitle}
          </div>
          <table class="items-table" style="margin-left: 20px; width: calc(100% - 20px);">
            ${TAXABLE_VISA_TYPES.map(
              (item) => {
                const parts = item.labelKh.split('-');
                const prefix = parts.length > 1 ? parts[0] : '';
                const title = parts.length > 1 ? parts.slice(1).join('-') : item.labelKh;
                return `
              <tr>
                <td style="width: 40%;" class="font-siemreap">${prefix ? `<span class="font-times">${prefix}-</span>` : ''}${title}</td>
                <td style="width: 15%;" class="bold font-times">${item.code}</td>
                <td style="width: 15%; text-align: right;" class="font-siemreap">ចំនួន<span class="font-times">=</span></td>
                <td style="width: 15%; text-align: right;" class="bold font-times">${isResultShown && hasEligibleUsage ? (visaCounts[item.key] || 0).toLocaleString() : '-'}</td>
                <td style="width: 15%;" class="font-siemreap">នាក់</td>
              </tr>
            `;
              }
            ).join('')}
          </table>

          <!-- Section B: Non-taxable -->
          <div class="bold font-moul" style="font-size: ${fontSizePt}pt; margin-top: 4pt; margin-bottom: 2pt;">
            ${sectionBTitle}
          </div>
          <table class="items-table" style="margin-left: 20px; width: calc(100% - 20px);">
            ${NON_TAXABLE_VISA_TYPES.map(
              (item) => {
                const parts = item.labelKh.split('-');
                const prefix = parts.length > 1 ? parts[0] : '';
                const title = parts.length > 1 ? parts.slice(1).join('-') : item.labelKh;
                return `
              <tr>
                <td style="width: 40%;" class="font-siemreap">${prefix ? `<span class="font-times">${prefix}-</span>` : ''}${title}</td>
                <td style="width: 15%;" class="bold font-times">${item.code}</td>
                <td style="width: 15%; text-align: right;" class="font-siemreap">ចំនួន<span class="font-times">=</span></td>
                <td style="width: 15%; text-align: right;" class="bold font-times">${isResultShown && hasEligibleUsage ? (visaCounts[item.key] || 0).toLocaleString() : '-'}</td>
                <td style="width: 15%;" class="font-siemreap">នាក់</td>
              </tr>
            `;
              }
            ).join('')}
            <tr>
              <td colspan="2" class="bold font-moul" style="text-align: right; padding-top: 2pt;">សរុប</td>
              <td style="text-align: right; padding-top: 2pt;" class="font-siemreap">ចំនួន<span class="font-times">=</span></td>
              <td style="text-align: right; padding-top: 2pt;" class="bold font-times">${isResultShown && hasEligibleUsage ? grandTotal.toLocaleString() : '-'}</td>
              <td style="padding-top: 2pt;" class="font-siemreap">នាក់</td>
            </tr>
          </table>

          <!-- Section C: Remarks -->
          <div style="margin-top: 4pt; margin-bottom: 6pt;">
            <span class="bold">គ. ចំណុចគួរអោយកត់សម្គាល់</span> &nbsp;&nbsp;&nbsp;&nbsp; ${remarksNote}
          </div>
        </div>

        <!-- Closing sentences -->
        <div style="text-align: justify; text-indent: ${tabIndentPx}px; margin-bottom: 3pt;" class="font-siemreap">
          ${formatClosingHtmlWithMuol(closingParagraph1)}
        </div>
        <div style="text-align: justify; text-indent: ${tabIndentPx}px; margin-bottom: 8pt;" class="font-siemreap">
          ${formatClosingHtmlWithMuol(closingParagraph2)}
        </div>

        <!-- Date & Signatures -->
        <table style="width: 100%; margin-top: ${signatureShiftY || 6}pt;">
          <tr>
            <td style="width: 35%;"></td>
            <td style="width: 65%; text-align: center; font-size: ${fontSizePt}pt;">
              ${lunarLines.map(line => `<div style="white-space: nowrap; line-height: 1.55; margin-bottom: 3.5pt;">${line}</div>`).join('')}
              <div style="white-space: nowrap; line-height: 1.55; margin-bottom: 3.5pt;">${displaySolarDate}</div>
              <div class="bold font-moul" style="font-size: ${fontSizePt}pt; line-height: 1.55; white-space: nowrap;">${chiefTitle}</div>
              <div style="height: 38pt;"></div>
              <div class="bold" style="font-size: ${fontSizePt}pt; white-space: nowrap;">${chiefName}</div>
            </td>
          </tr>
        </table>
      </body>
      </html>
    `;

    const blob = new Blob([docContent], { type: 'application/msword;charset=utf-8' });
    saveAs(blob, fileName);
  };

  // Excel Export (.xlsx)
  const handleExportExcel = () => {
    const safeTeam = (selectedTeam || 'team').replace(/[\/\s]+/g, '_');
    const fileName = `${reportMainTitle}_${safeTeam}_${startDate}_${endDate}.xlsx`;

    const rows: any[] = [
      ['ព្រះរាជាណាចក្រកម្ពុជា'],
      ['ជាតិ សាសនា ព្រះមហាក្សត្រ'],
      [''],
      [ministryName],
      [deptName],
      [firstGateDeptName],
      [officeName],
      [teamHeaderName],
      [''],
      [salutationTop],
      [salutationRecipient],
      [''],
      ['កម្មវត្ថុ៖', effectiveSubjectText],
      [''],
      [sectionATitle, '', '', '', ''],
      ['ល.រ', 'ប្រភេទទិដ្ឋាការ', 'កូដ', 'ចំនួន', 'ខ្នាត'],
    ];

    TAXABLE_VISA_TYPES.forEach((item, idx) => {
      rows.push([
        idx + 1,
        item.labelKh.replace(/^[0-9]+-/, ''),
        item.code,
        isResultShown && hasEligibleUsage ? (visaCounts[item.key] || 0) : '-',
        'នាក់',
      ]);
    });

    rows.push(['']);
    rows.push([sectionBTitle, '', '', '', '']);

    NON_TAXABLE_VISA_TYPES.forEach((item, idx) => {
      rows.push([
        idx + 1,
        item.labelKh.replace(/^[0-9]+-/, ''),
        item.code,
        isResultShown && hasEligibleUsage ? (visaCounts[item.key] || 0) : '-',
        'នាក់',
      ]);
    });

    rows.push(['សរុប', '', '', isResultShown && hasEligibleUsage ? grandTotal : '-', 'នាក់']);
    rows.push(['']);
    rows.push(['គ. ចំណុចគួរអោយកត់សម្គាល់', remarksNote]);
    rows.push(['']);
    rows.push([closingParagraph1]);
    rows.push([closingParagraph2]);
    rows.push(['']);
    lunarLines.forEach((l) => {
      rows.push(['', '', l]);
    });
    rows.push(['', '', displaySolarDate]);
    rows.push(['', '', chiefTitle]);
    if (chiefName) {
      rows.push(['', '', chiefName]);
    }

    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, reportMainTitle);
    XLSX.writeFile(wb, fileName);
  };

  return (
    <div className="min-h-screen bg-slate-100 flex flex-col print:bg-white print:m-0 print:p-0">
      {/* Top Action Toolbar (Hidden when printing) */}
      <div className="bg-white border-b border-gray-200 sticky top-0 z-30 shadow-sm px-4 py-3 print:hidden">
        <div className="max-w-7xl mx-auto flex flex-wrap items-center justify-between gap-3">
          {/* Left: Title & Team Selector */}
          <div className="flex items-center gap-3">
            {onClose && (
              <button
                onClick={onClose}
                className="p-1.5 rounded-md hover:bg-gray-100 text-gray-600 transition"
                title="ត្រឡប់ក្រោយ"
              >
                <ArrowLeft className="w-5 h-5" />
              </button>
            )}
            <div className="flex items-center gap-2">
              <FileText className="w-6 h-6 text-[#C6A15B]" />
              <div>
                <h1 className="text-base font-bold text-gray-900 font-siemreap leading-tight">
                  {reportMainTitle}
                </h1>
                <p className="text-xs text-gray-500 font-siemreap">
                  {reportSubtitle}
                </p>
              </div>
            </div>

            {/* Category Dropdown (All / Sticker / cEA) */}
            <div className="flex items-center gap-1.5 ml-2">
              <span className="text-xs font-medium text-gray-600 font-siemreap">ប្រភេទ៖</span>
              <select
                value={selectedCategory}
                onChange={(e) => {
                  setSelectedCategory(e.target.value as 'all' | 'sticker' | 'cea');
                  setIsResultShown(false);
                }}
                className="h-[36px] min-h-[36px] max-h-[36px] box-border text-xs border border-gray-300 rounded-md px-2.5 py-1 bg-white font-siemreap text-gray-800 focus:ring-1 focus:ring-[#C6A15B] focus:border-[#C6A15B] shadow-xs font-semibold"
                title="ជ្រើសរើសប្រភេទ៖ ទាំងអស់ (Sticker + cEA), សន្លឹកស្អិត (Sticker), ឬ ក្រដាសអនុម័ត (cEA)"
              >
                <option value="all">ទាំងអស់ (Sticker + cEA)</option>
                <option value="sticker">សន្លឹកស្អិត (Sticker)</option>
                <option value="cea">ក្រដាសអនុម័ត (cEA)</option>
              </select>
            </div>

            {/* Dynamic Usage Status Indicator based on selected Category */}
            <div className="flex items-center ml-1">
              {hasEligibleUsage ? (
                <span className="inline-flex items-center gap-1 text-[11px] font-medium text-emerald-700 bg-emerald-50 border border-emerald-200 px-2 py-1 rounded shadow-2xs whitespace-nowrap">
                  <span className="w-1.5 h-1.5 rounded-full bg-emerald-500"></span>
                  {selectedCategory === 'sticker'
                    ? `មាន Sticker (${stickerSheets.toLocaleString()} សន្លឹក)`
                    : selectedCategory === 'cea'
                    ? (ceaSheets > 0 ? `មាន cEA (${ceaSheets.toLocaleString()} សន្លឹក)` : 'មានទិន្នន័យ cEA')
                    : (hasStickerUsage && hasCeaUsage
                        ? 'មាន Sticker + cEA'
                        : hasStickerUsage
                        ? `មាន Sticker (${stickerSheets.toLocaleString()} សន្លឹក)`
                        : 'មានទិន្នន័យ cEA')}
                </span>
              ) : (
                <span
                  className="inline-flex items-center gap-1 text-[11px] font-medium text-amber-800 bg-amber-50 border border-amber-300 px-2 py-1 rounded shadow-2xs whitespace-nowrap"
                  title={
                    selectedCategory === 'sticker'
                      ? 'ក្រុមនេះមិនមានប្រើប្រាស់ Sticker ទេ ដូច្នេះមិនបង្ហាញលទ្ធផលឡើយ'
                      : selectedCategory === 'cea'
                      ? 'ក្រុមនេះមិនមានប្រើប្រាស់ cEA ទេ ដូច្នេះមិនបង្ហាញលទ្ធផលឡើយ'
                      : 'ក្រុមនេះមិនមានទិន្នន័យប្រើប្រាស់ Sticker ឬ cEA ទេ ដូច្នេះមិនបង្ហាញលទ្ធផលឡើយ'
                  }
                >
                  <AlertCircle className="w-3 h-3 text-amber-600" />
                  {selectedCategory === 'sticker'
                    ? 'គ្មាន Sticker (មិនបង្ហាញលទ្ធផល)'
                    : selectedCategory === 'cea'
                    ? 'គ្មាន cEA (មិនបង្ហាញលទ្ធផល)'
                    : 'គ្មានទិន្នន័យ (មិនបង្ហាញលទ្ធផល)'}
                </span>
              )}
            </div>
          </div>

          {/* Right: Actions & Tools */}
          <div className="flex items-center gap-2">
            {/* Save Button */}
            <button
              onClick={handleSaveData}
              className="flex items-center gap-1 text-xs px-2.5 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded-md font-siemreap transition shadow-sm cursor-pointer"
              title="រក្សាទុកទិន្នន័យរបាយការណ៍"
            >
              <Save className="w-3.5 h-3.5" />
              <span>រក្សាទុក</span>
            </button>

            {/* Settings Toggle */}
            <button
              onClick={() => setShowSettings(!showSettings)}
              className={`p-1.5 rounded-md border transition ${
                showSettings ? 'bg-[#C6A15B] text-white border-[#C6A15B]' : 'bg-white text-gray-700 border-gray-300 hover:bg-gray-50'
              }`}
              title="កំណត់ក្បាលលិខិត & ហត្ថលេខា"
            >
              <SlidersHorizontal className="w-4 h-4" />
            </button>

            {/* Zoom Controls */}
            <div className="flex items-center border border-gray-300 rounded-md overflow-hidden bg-white">
              <button
                onClick={() => setZoomScale((prev) => Math.max(prev - 10, 50))}
                className="p-1.5 hover:bg-gray-100 text-gray-600"
                title="Zoom Out"
              >
                <ZoomOut className="w-3.5 h-3.5" />
              </button>
              <span className="text-[11px] font-mono px-1.5 text-gray-700 select-none">
                {zoomScale}%
              </span>
              <button
                onClick={() => setZoomScale((prev) => Math.min(prev + 10, 150))}
                className="p-1.5 hover:bg-gray-100 text-gray-600"
                title="Zoom In"
              >
                <ZoomIn className="w-3.5 h-3.5" />
              </button>
            </div>

            {/* Quick Top Margin Stepper (ទម្លាក់គម្លាតពីលើចុះក្រោម) */}
            <div className="flex items-center border border-gray-300 rounded-md bg-white shadow-2xs px-2 py-1 text-xs text-gray-700 font-siemreap">
              <span className="text-[11px] text-gray-500 mr-1.5 select-none font-medium">គែមលើ៖</span>
              <button
                type="button"
                onClick={() => setPageTopMargin((prev) => Math.max(0.2, parseFloat((prev - 0.1).toFixed(2))))}
                className="w-5 h-5 flex items-center justify-center rounded hover:bg-gray-100 text-gray-700 font-bold cursor-pointer transition"
                title="បន្ថយគម្លាតពីលើ (-0.1cm)"
              >
                -
              </button>
              <span className="font-mono text-xs font-bold px-1.5 text-[#002060] select-none">
                {pageTopMargin.toFixed(1)}cm
              </span>
              <button
                type="button"
                onClick={() => setPageTopMargin((prev) => Math.min(3.5, parseFloat((prev + 0.1).toFixed(2))))}
                className="w-5 h-5 flex items-center justify-center rounded hover:bg-gray-100 text-gray-700 font-bold cursor-pointer transition"
                title="ទម្លាក់គម្លាតពីលើចុះ (+0.1cm)"
              >
                +
              </button>
            </div>

            {/* Quick Paragraph Shift Stepper (ទម្លាក់កថាខណ្ឌចុះក្រោម) */}
            <div className="flex items-center border border-amber-300 rounded-md bg-amber-50/60 shadow-2xs px-2 py-1 text-xs text-amber-950 font-siemreap">
              <span className="text-[11px] text-amber-800 mr-1.5 select-none font-medium">កថាខណ្ឌ៖</span>
              <button
                type="button"
                onClick={() => setParagraphShiftY((prev) => Math.max(0, prev - 2))}
                className="w-5 h-5 flex items-center justify-center rounded hover:bg-amber-200 text-amber-900 font-bold cursor-pointer transition"
                title="រំកិលកថាខណ្ឌឡើងលើ (-2pt)"
              >
                -
              </button>
              <span className="font-mono text-xs font-bold px-1.5 text-amber-900 select-none min-w-[34px] text-center">
                {paragraphShiftY}pt
              </span>
              <button
                type="button"
                onClick={() => setParagraphShiftY((prev) => Math.min(60, prev + 2))}
                className="w-5 h-5 flex items-center justify-center rounded hover:bg-amber-200 text-amber-900 font-bold cursor-pointer transition"
                title="ទម្លាក់កថាខណ្ឌចុះក្រោម (+2pt)"
              >
                +
              </button>
            </div>

            {/* Quick Signature Shift Stepper (រំកិលហត្ថលេខាឡើង/ចុះ) */}
            <div className="flex items-center border border-teal-300 rounded-md bg-teal-50/60 shadow-2xs px-2 py-1 text-xs text-teal-950 font-siemreap">
              <span className="text-[11px] text-teal-800 mr-1.5 select-none font-medium">ហត្ថលេខា៖</span>
              <button
                type="button"
                onClick={() => setSignatureShiftY((prev) => Math.max(-15, prev - 2))}
                className="w-5 h-5 flex items-center justify-center rounded hover:bg-teal-200 text-teal-900 font-bold cursor-pointer transition"
                title="រំកិលហត្ថលេខាឡើងលើ (-2pt)"
              >
                -
              </button>
              <span className="font-mono text-xs font-bold px-1.5 text-teal-900 select-none min-w-[30px] text-center">
                {signatureShiftY}pt
              </span>
              <button
                type="button"
                onClick={() => setSignatureShiftY((prev) => Math.min(40, prev + 2))}
                className="w-5 h-5 flex items-center justify-center rounded hover:bg-teal-200 text-teal-900 font-bold cursor-pointer transition"
                title="ទម្លាក់ហត្ថលេខាចុះក្រោម (+2pt)"
              >
                +
              </button>
            </div>

            {/* Quick Line Spacing Stepper (ទម្លាក់គម្លាតអក្សរពីលើចុះក្រោម / Line Spacing) */}
            <div className="flex items-center border border-indigo-300 rounded-md bg-indigo-50/60 shadow-2xs px-2 py-1 text-xs text-indigo-950 font-siemreap">
              <span className="text-[11px] text-indigo-800 mr-1.5 select-none font-medium">គម្លាតជួរ៖</span>
              <button
                type="button"
                onClick={() => setLineSpacing((prev) => Math.max(0.85, parseFloat((prev - 0.05).toFixed(2))))}
                className="w-5 h-5 flex items-center justify-center rounded hover:bg-indigo-200 text-indigo-900 font-bold cursor-pointer transition"
                title="បង្រួមគម្លាតជួរអក្សរ (-0.05)"
              >
                -
              </button>
              <span className="font-mono text-xs font-bold px-1.5 text-indigo-900 select-none min-w-[32px] text-center">
                {lineSpacing.toFixed(2)}
              </span>
              <button
                type="button"
                onClick={() => setLineSpacing((prev) => Math.min(2.3, parseFloat((prev + 0.05).toFixed(2))))}
                className="w-5 h-5 flex items-center justify-center rounded hover:bg-indigo-200 text-indigo-900 font-bold cursor-pointer transition"
                title="ទម្លាក់គម្លាតជួរអក្សរចុះ (+0.05)"
              >
                +
              </button>
            </div>

            {/* Export Dropdown / Actions */}
            <div className="flex items-center gap-1">
              <button
                onClick={handleExportPdf}
                disabled={isExporting}
                className="flex items-center gap-1 text-xs px-2.5 py-1.5 bg-red-600 hover:bg-red-700 text-white rounded-md font-siemreap transition shadow-sm disabled:opacity-50"
                title="ទាញយកជា PDF"
              >
                <Download className="w-3.5 h-3.5" />
                <span>PDF</span>
              </button>

              <button
                onClick={handleExportWord}
                className="flex items-center gap-1 text-xs px-2.5 py-1.5 bg-blue-700 hover:bg-blue-800 text-white rounded-md font-siemreap transition shadow-sm"
                title="ទាញយកជា Word (.doc)"
              >
                <FileText className="w-3.5 h-3.5" />
                <span>Word</span>
              </button>

              <button
                onClick={handleExportExcel}
                className="flex items-center gap-1 text-xs px-2.5 py-1.5 bg-green-700 hover:bg-green-800 text-white rounded-md font-siemreap transition shadow-sm"
                title="ទាញយកជា Excel (.xlsx)"
              >
                <FileSpreadsheet className="w-3.5 h-3.5" />
                <span>Excel</span>
              </button>
            </div>
          </div>
        </div>

        {/* Date Filters & Controls (Matching RobokTotalStockWorkReport) */}
        <div className="max-w-7xl mx-auto mt-2.5 bg-white border border-gray-200 rounded-xl p-3 shadow-xs">
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-5 gap-3 pt-1 text-xs">
            <div>
              <label className="block font-bold text-gray-700 mb-1 h-5 flex items-center text-xs truncate whitespace-nowrap">
                ជ្រើសរើសខែរបាយការណ៍ :
              </label>
              <select
                value={selectedMonth}
                onChange={(e) => handleMonthYearChange(selectedYear, parseInt(e.target.value, 10))}
                className="w-full h-[38px] min-h-[38px] max-h-[38px] box-border border border-gray-300 rounded px-2.5 py-1.5 font-sans text-xs focus:border-blue-500 focus:outline-none font-bold text-gray-800 bg-white cursor-pointer shadow-xs"
              >
                {KHMER_MONTHS.map((m, idx) => (
                  <option key={idx + 1} value={idx + 1}>
                    ខែ {m} (ខែ {toKhmerNum(idx + 1)})
                  </option>
                ))}
              </select>
            </div>

            <div>
              <label className="block font-bold text-gray-700 mb-1 h-5 flex items-center text-xs truncate whitespace-nowrap">
                ជ្រើសរើសឆ្នាំ :
              </label>
              <select
                value={selectedYear}
                onChange={(e) => handleMonthYearChange(parseInt(e.target.value, 10), selectedMonth)}
                className="w-full h-[38px] min-h-[38px] max-h-[38px] box-border border border-gray-300 rounded px-2.5 py-1.5 font-sans text-xs focus:border-blue-500 focus:outline-none font-bold text-gray-800 bg-white cursor-pointer shadow-xs"
              >
                {Array.from({ length: 15 }, (_, i) => 2018 + i).map((y) => (
                  <option key={y} value={y}>
                    ឆ្នាំ {toKhmerNum(y)} ({y})
                  </option>
                ))}
              </select>
            </div>

            {singleDateMode ? (
              <div className="sm:col-span-2 lg:col-span-2">
                <label className="block font-bold text-gray-700 mb-1 h-5 flex items-center text-xs truncate whitespace-nowrap">
                  📅 កាលបរិច្ឆេទរបាយការណ៍ (ជ្រើសរើស ១ ថ្ងៃគត់) :
                </label>
                <CustomDatePicker
                  value={startDate}
                  onChange={(val) => {
                    setStartDate(val);
                    setEndDate(val);
                    setActiveDurationPreset(null);
                    setIsResultShown(false);
                    if (val) {
                      const parts = val.split('-');
                      if (parts.length === 3) {
                        const y = parseInt(parts[0], 10);
                        const m = parseInt(parts[1], 10);
                        if (!isNaN(y) && y >= 1900 && y <= 2100) setSelectedYear(y);
                        if (!isNaN(m) && m >= 1 && m <= 12) setSelectedMonth(m);
                      }
                    }
                  }}
                  className="h-[38px] min-h-[38px] max-h-[38px] box-border border border-gray-300 rounded px-2.5 py-1.5 font-sans text-xs focus:border-blue-500 focus:outline-none font-bold text-gray-800 bg-white shadow-xs"
                />
              </div>
            ) : (
              <>
                <div>
                  <label className="block font-bold text-gray-700 mb-1 h-5 flex items-center text-xs truncate whitespace-nowrap">
                    កាលបរិច្ឆេទចាប់ផ្តើម (From Date) :
                  </label>
                  <CustomDatePicker
                    value={startDate}
                    onChange={handleStartDateChange}
                    className="h-[38px] min-h-[38px] max-h-[38px] box-border border border-gray-300 rounded px-2.5 py-1.5 font-sans text-xs focus:border-blue-500 focus:outline-none font-bold text-gray-800 bg-white shadow-xs"
                  />
                </div>

                <div>
                  <label className="block font-bold text-gray-700 mb-1 h-5 flex items-center text-xs truncate whitespace-nowrap">
                    កាលបរិច្ឆេទបញ្ចប់ (To Date) :
                  </label>
                  <CustomDatePicker
                    value={endDate}
                    onChange={handleEndDateChange}
                    className="h-[38px] min-h-[38px] max-h-[38px] box-border border border-gray-300 rounded px-2.5 py-1.5 font-sans text-xs focus:border-blue-500 focus:outline-none font-bold text-gray-800 bg-white shadow-xs"
                  />
                </div>
              </>
            )}

            {/* Tacteing: Set Document Date */}
            <div>
              <label className="block font-bold text-gray-700 mb-1 h-5 flex items-center text-xs truncate whitespace-nowrap">
                📅 ថ្ងៃខែឆ្នាំតាក់តែងលិខិត :
              </label>
              <CustomDatePicker
                value={signingDate}
                onChange={handleSigningDateChange}
                className="h-[38px] min-h-[38px] max-h-[38px] box-border border border-gray-300 rounded px-2.5 py-1.5 font-sans text-xs focus:border-blue-500 focus:outline-none bg-blue-50/40 text-blue-950 font-bold shadow-xs"
                title="កំណត់ថ្ងៃខែឆ្នាំតាក់តែងលិខិត (ចុចដើម្បីជ្រើសរើស)"
              />
            </div>
          </div>

          {/* Quick Range Presets & Action Button */}
          <div className="mt-2.5 pt-2.5 border-t border-gray-100 flex flex-wrap items-center justify-between gap-2 text-xs">
            <div className="flex flex-wrap items-center gap-1.5">
              <span className="text-gray-500 font-bold h-8 flex items-center">ចន្លោះកាលបរិច្ឆេទរហ័ស ៖</span>
              <button
                type="button"
                onClick={() => handlePresetRange('this_month')}
                className="h-8 px-2.5 rounded border border-gray-300 bg-gray-50 hover:bg-gray-100 text-gray-700 text-xs font-semibold transition cursor-pointer flex items-center justify-center"
              >
                ខែនេះ
              </button>
              <button
                type="button"
                onClick={() => handlePresetRange('last_month')}
                className="h-8 px-2.5 rounded border border-gray-300 bg-gray-50 hover:bg-gray-100 text-gray-700 text-xs font-semibold transition cursor-pointer flex items-center justify-center"
              >
                ខែមុន
              </button>
              <button
                type="button"
                onClick={() => handlePresetRange('next_month')}
                className="h-8 px-2.5 rounded border border-gray-300 bg-gray-50 hover:bg-gray-100 text-gray-700 text-xs font-semibold transition cursor-pointer flex items-center justify-center"
              >
                ខែបន្ទាប់
              </button>
              <button
                type="button"
                onClick={() => handlePresetRange('q1')}
                className={`h-8 px-2.5 rounded border text-xs font-semibold transition cursor-pointer flex items-center justify-center ${
                  isPresetActive('q1')
                    ? 'bg-blue-600 border-blue-600 text-white font-bold shadow-xs'
                    : 'bg-gray-50 hover:bg-gray-100 border-gray-300 text-gray-700'
                }`}
                title="ត្រីមាស (៣ខែ ចាប់ពីខែដែលបានជ្រើសរើស)"
              >
                ត្រីមាសទី១
              </button>
              <button
                type="button"
                onClick={() => handlePresetRange('s1')}
                className={`h-8 px-2.5 rounded border text-xs font-semibold transition cursor-pointer flex items-center justify-center ${
                  isPresetActive('s1')
                    ? 'bg-blue-600 border-blue-600 text-white font-bold shadow-xs'
                    : 'bg-gray-50 hover:bg-gray-100 border-gray-300 text-gray-700'
                }`}
                title="ឆមាស (៦ខែ ចាប់ពីខែដែលបានជ្រើសរើស)"
              >
                ឆមាសទី១
              </button>
              <button
                type="button"
                onClick={() => handlePresetRange('m9')}
                className={`h-8 px-2.5 rounded border text-xs font-semibold transition cursor-pointer flex items-center justify-center ${
                  isPresetActive('m9')
                    ? 'bg-blue-600 border-blue-600 text-white font-bold shadow-xs'
                    : 'bg-gray-50 hover:bg-gray-100 border-gray-300 text-gray-700'
                }`}
                title="នព្វមាស (៩ខែ ចាប់ពីខែដែលបានជ្រើសរើស)"
              >
                នព្វមាស
              </button>
              <button
                type="button"
                onClick={() => handlePresetRange('full_year')}
                className={`h-8 px-2.5 rounded border text-xs font-semibold transition cursor-pointer flex items-center justify-center ${
                  isPresetActive('full_year')
                    ? 'bg-blue-600 border-blue-600 text-white font-bold shadow-xs'
                    : 'bg-gray-50 hover:bg-gray-100 border-gray-300 text-gray-700'
                }`}
                title="រយៈពេល ១២ខែ (១២ខែ ចាប់ពីខែដែលបានជ្រើសរើស)"
              >
                រយៈពេល ១២ខែ
              </button>
            </div>

            <div className="flex items-center gap-1.5 flex-wrap">
              {/* Show Result Action Button */}
              <button
                type="button"
                onClick={handleShowResult}
                className={`h-8 px-3.5 rounded font-bold text-xs flex items-center justify-center gap-1.5 transition cursor-pointer shadow-sm ${
                  !hasEligibleUsage
                    ? 'bg-slate-300 hover:bg-slate-400 text-slate-700 font-moul tracking-wide'
                    : !isResultShown
                    ? 'bg-blue-600 hover:bg-blue-700 active:scale-95 text-white ring-2 ring-blue-400 font-moul tracking-wide shadow-md'
                    : 'bg-emerald-600 hover:bg-emerald-700 active:scale-95 text-white font-moul tracking-wide'
                }`}
                title={hasEligibleUsage ? "ចុចដើម្បីទាញយក និងបង្ហាញលទ្ធផលរបាយការណ៍" : "ក្រុមនេះគ្មានទិន្នន័យប្រើប្រាស់សម្រាប់ប្រភេទនេះទេ"}
              >
                <Search className="w-3.5 h-3.5" />
                <span>បង្ហាញលទ្ធផលរួចរាល់</span>
              </button>

              {/* Clear to Dash (-) Button */}
              <button
                type="button"
                onClick={handleClearToDash}
                className="h-8 flex items-center gap-1 text-xs px-2.5 bg-slate-100 hover:bg-slate-200 text-slate-700 rounded border border-slate-300 font-siemreap transition cursor-pointer justify-center"
                title="សម្អាតតួលេខឱ្យទៅជា (-)"
              >
                <RotateCcw className="w-3 h-3 text-slate-500" />
                <span>ជម្រះ (-)</span>
              </button>

              {/* Load PDF Sample Data */}
              <button
                onClick={handleLoadPdfSample}
                className="h-8 flex items-center gap-1.5 text-xs px-2.5 bg-amber-50 hover:bg-amber-100 text-amber-800 rounded border border-amber-300 font-siemreap transition cursor-pointer justify-center"
                title="ផ្ទុកទិន្នន័យគំរូពីឯកសារ PDF (ត្រីមាសទី១ ២០២៦ សរុប ១២៧,៩៥៤)"
              >
                <Sparkles className="w-3.5 h-3.5 text-amber-600" />
                <span>គំរូ PDF</span>
              </button>
            </div>
          </div>
        </div>

        {/* No Usage Notice Toast */}
        {noUsageNotice && (
          <div className="max-w-7xl mx-auto mt-2">
            <div className="bg-amber-50 border border-amber-300 text-amber-900 text-xs px-3.5 py-2 rounded-lg flex items-center gap-2 font-siemreap shadow-sm">
              <AlertCircle className="w-4 h-4 text-amber-600 shrink-0" />
              <span>
                <strong>មិនអាចបង្ហាញលទ្ធផលបានទេ៖</strong> ក្រុម «<strong>{selectedTeam}</strong>» មិនមានទិន្នន័យប្រើប្រាស់ {selectedCategory === 'sticker' ? 'Sticker' : selectedCategory === 'cea' ? 'cEA' : 'Sticker ឬ cEA'} ក្នុងកាលបរិច្ឆេទនេះទេ ដូច្នេះលទ្ធផលត្រូវបានកំណត់មិនបង្ហាញឡើយ (បង្ហាញសញ្ញា «-» ទាំងអស់)។
              </span>
            </div>
          </div>
        )}

        {/* Save Notice Toast */}
        {isSavedNotice && (
          <div className="max-w-7xl mx-auto mt-2">
            <div className="bg-emerald-50 border border-emerald-200 text-emerald-800 text-xs px-3.5 py-2 rounded-lg flex items-center gap-2 font-siemreap shadow-sm">
              <CheckCircle2 className="w-4 h-4 text-emerald-600" />
              <span>បានរក្សាទុកទិន្នន័យរបាយការណ៍ដោយជោគជ័យ!</span>
            </div>
          </div>
        )}

        {/* Settings Drawer */}
        {showSettings && (
          <div className="max-w-7xl mx-auto mt-3 p-4 bg-slate-50 border border-slate-300 rounded-lg shadow-inner font-siemreap text-xs grid grid-cols-1 md:grid-cols-3 gap-4">
            {/* Col 1: Header Titles */}
            <div className="space-y-2">
              <div className="font-bold text-gray-800 border-b pb-1">១. ក្បាលលិខិត (Header)</div>
              <div>
                <label className="text-gray-600 block text-xs mb-1">ក្រសួង៖</label>
                <input
                  type="text"
                  value={ministryName}
                  onChange={(e) => setMinistryName(e.target.value)}
                  className="w-full h-8 min-h-[32px] max-h-[32px] box-border border border-gray-300 rounded px-2.5 py-1 bg-white text-xs text-gray-800 focus:ring-1 focus:ring-blue-500 focus:outline-none"
                />
              </div>
              <div>
                <label className="text-gray-600 block text-xs mb-1">អគ្គនាយកដ្ឋាន៖</label>
                <input
                  type="text"
                  value={deptName}
                  onChange={(e) => setDeptName(e.target.value)}
                  className="w-full h-8 min-h-[32px] max-h-[32px] box-border border border-gray-300 rounded px-2.5 py-1 bg-white text-xs text-gray-800 focus:ring-1 focus:ring-blue-500 focus:outline-none"
                />
              </div>
              <div>
                <label className="text-gray-600 block text-xs mb-1">នាយកដ្ឋាន៖</label>
                <input
                  type="text"
                  value={firstGateDeptName}
                  onChange={(e) => setFirstGateDeptName(e.target.value)}
                  className="w-full h-8 min-h-[32px] max-h-[32px] box-border border border-gray-300 rounded px-2.5 py-1 bg-white text-xs text-gray-800 focus:ring-1 focus:ring-blue-500 focus:outline-none"
                />
              </div>
              <div>
                <label className="text-gray-600 block text-xs mb-1">ការិយាល័យ៖</label>
                <input
                  type="text"
                  value={officeName}
                  onChange={(e) => setOfficeName(e.target.value)}
                  className="w-full h-8 min-h-[32px] max-h-[32px] box-border border border-gray-300 rounded px-2.5 py-1 bg-white text-xs text-gray-800 focus:ring-1 focus:ring-blue-500 focus:outline-none"
                />
              </div>
              <div>
                <label className="text-gray-600 block text-xs mb-1">ក្រុម / អង្គភាព៖</label>
                <input
                  type="text"
                  value={teamHeaderName}
                  onChange={(e) => setTeamHeaderName(e.target.value)}
                  className="w-full h-8 min-h-[32px] max-h-[32px] box-border border border-gray-300 rounded px-2.5 py-1 bg-white text-xs text-gray-800 focus:ring-1 focus:ring-blue-500 focus:outline-none"
                />
              </div>
            </div>

            {/* Col 2: Salutation & Subject */}
            <div className="space-y-2">
              <div className="font-bold text-gray-800 border-b pb-1">២. គោរពជូន & កម្មវត្ថុ</div>
              <div>
                <label className="text-gray-600 block text-xs mb-1">ពាក្យគោរព៖</label>
                <input
                  type="text"
                  value={salutationTop}
                  onChange={(e) => setSalutationTop(e.target.value)}
                  className="w-full h-8 min-h-[32px] max-h-[32px] box-border border border-gray-300 rounded px-2.5 py-1 bg-white text-xs text-gray-800 focus:ring-1 focus:ring-blue-500 focus:outline-none"
                />
              </div>
              <div>
                <div className="flex items-center justify-between mb-1">
                  <label className="text-gray-600 block text-xs">អ្នកទទួល (Recipient)៖</label>
                  {isBranchChiefMode ? (
                    <button
                      type="button"
                      onClick={() => setSalutationRecipient('លោកវរសេនីយ៍ឯក ប្រធានសាខាប្រមូលចំណូល')}
                      className="text-[10px] text-blue-600 hover:underline cursor-pointer"
                    >
                      លំនាំដើម
                    </button>
                  ) : (
                    <button
                      type="button"
                      onClick={() => setSalutationRecipient('លោកវរសេនីយ៍ទោ នាយការិយាល័យទិដ្ឋាការចូល')}
                      className="text-[10px] text-blue-600 hover:underline cursor-pointer"
                    >
                      លំនាំដើម
                    </button>
                  )}
                </div>
                <input
                  type="text"
                  value={salutationRecipient}
                  onChange={(e) => setSalutationRecipient(e.target.value)}
                  className="w-full h-8 min-h-[32px] max-h-[32px] box-border border border-gray-300 rounded px-2.5 py-1 bg-white text-xs text-gray-800 focus:ring-1 focus:ring-blue-500 focus:outline-none"
                />
              </div>
              <div>
                <label className="text-gray-600 block text-xs mb-1">ទីតាំងបន្ថែមក្នុងកម្មវត្ថុ (ស្រេចចិត្ត)៖</label>
                <input
                  type="text"
                  placeholder="ទុកនៅទំនេរ ឬបញ្ចូលបន្ថែម (ឧ. អាកាសយានដ្ឋាន...)"
                  value={customLocation}
                  onChange={(e) => setCustomLocation(e.target.value)}
                  className="w-full h-8 min-h-[32px] max-h-[32px] box-border border border-gray-300 rounded px-2.5 py-1 bg-white text-xs text-gray-800 focus:ring-1 focus:ring-blue-500 focus:outline-none"
                />
              </div>
              <div className="pt-1">
                <label className="flex items-center gap-1.5 cursor-pointer text-gray-700 text-xs">
                  <input
                    type="checkbox"
                    checked={isCustomSubject}
                    onChange={(e) => {
                      setIsCustomSubject(e.target.checked);
                      if (e.target.checked && !customSubjectText) {
                        setCustomSubjectText(defaultSubjectText);
                      }
                    }}
                    className="rounded text-[#C6A15B]"
                  />
                  <span>កែប្រែកម្មវត្ថុដោយផ្ទាល់ (Custom Subject)</span>
                </label>
                {isCustomSubject && (
                  <textarea
                    rows={2}
                    value={customSubjectText}
                    onChange={(e) => setCustomSubjectText(e.target.value)}
                    className="w-full border border-gray-300 rounded px-2.5 py-1 mt-1 bg-white text-xs text-gray-800 focus:ring-1 focus:ring-blue-500 focus:outline-none"
                  />
                )}
              </div>
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="text-gray-600 block text-xs mb-1">ចំណងជើងផ្នែក ក.៖</label>
                  <input
                    type="text"
                    value={sectionATitle}
                    onChange={(e) => setSectionATitle(e.target.value)}
                    className="w-full h-8 min-h-[32px] max-h-[32px] box-border border border-gray-300 rounded px-2.5 py-1 bg-white text-xs text-gray-800 focus:ring-1 focus:ring-blue-500 focus:outline-none"
                    placeholder="ក. ប្រភេទទិដ្ឋាការបង់អាករ"
                  />
                </div>
                <div>
                  <label className="text-gray-600 block text-xs mb-1">ចំណងជើងផ្នែក ខ.៖</label>
                  <input
                    type="text"
                    value={sectionBTitle}
                    onChange={(e) => setSectionBTitle(e.target.value)}
                    className="w-full h-8 min-h-[32px] max-h-[32px] box-border border border-gray-300 rounded px-2.5 py-1 bg-white text-xs text-gray-800 focus:ring-1 focus:ring-blue-500 focus:outline-none"
                    placeholder="ខ. ប្រភេទទិដ្ឋាការមិនបង់អាករ"
                  />
                </div>
              </div>

              <div>
                <label className="text-gray-600 block text-xs mb-1">ចំណុចគួរអោយកត់សម្គាល់ (ផ្នែក គ)៖</label>
                <input
                  type="text"
                  value={remarksNote}
                  onChange={(e) => setRemarksNote(e.target.value)}
                  className="w-full h-8 min-h-[32px] max-h-[32px] box-border border border-gray-300 rounded px-2.5 py-1 bg-white text-xs text-gray-800 focus:ring-1 focus:ring-blue-500 focus:outline-none"
                />
              </div>
              <div>
                <label className="text-gray-600 block text-xs mb-1">
                  គម្លាត Tab ឈរស្មើនិង របាយការណ៍ស្តីពី (px)៖
                </label>
                <div className="flex items-center gap-2">
                  <input
                    type="number"
                    min="0"
                    max="250"
                    value={tabIndentPx}
                    onChange={(e) => {
                      const val = parseInt(e.target.value, 10) || 0;
                      setTabIndentPx(val);
                      try {
                        localStorage.setItem(`${keyPrefix}team_visa_stats_tab_indent`, String(val));
                      } catch {}
                    }}
                    className="w-20 h-8 min-h-[32px] max-h-[32px] box-border border border-gray-300 rounded px-2 py-1 bg-white text-xs font-bold text-gray-800 focus:ring-1 focus:ring-blue-500 focus:outline-none text-center"
                  />
                  <span className="text-[11px] text-gray-500">(លំនាំដើម 88px ស្មើពាក្យ កម្មវត្ថុ ៖)</span>
                </div>
              </div>

              <div>
                <label className="text-gray-600 block text-xs font-medium mb-1">
                  ទំហំពុម្ពអក្សររបាយការណ៍ (Font Size pt)៖
                </label>
                <div className="flex items-center gap-2">
                  <input
                    type="number"
                    step="0.5"
                    min="8"
                    max="18"
                    value={fontSizePt}
                    onChange={(e) => {
                      const val = parseFloat(e.target.value) || 12;
                      setFontSizePt(val);
                      try {
                        localStorage.setItem(`${keyPrefix}team_visa_stats_font_size`, String(val));
                      } catch {}
                    }}
                    className="w-20 h-8 min-h-[32px] max-h-[32px] box-border border border-gray-300 rounded px-2 py-1 bg-white text-xs font-bold font-times text-gray-800 focus:ring-1 focus:ring-blue-500 focus:outline-none text-center"
                  />
                  <div className="flex items-center gap-1">
                    {[10, 11, 11.5, 12, 12.5, 13, 14].map((sz) => (
                      <button
                        key={sz}
                        type="button"
                        onClick={() => {
                          setFontSizePt(sz);
                          try {
                            localStorage.setItem(`${keyPrefix}team_visa_stats_font_size`, String(sz));
                          } catch {}
                        }}
                        className={`h-8 px-2 rounded border transition cursor-pointer text-xs flex items-center justify-center ${
                          fontSizePt === sz
                            ? 'bg-blue-600 text-white font-bold border-blue-700 shadow-xs'
                            : 'bg-white text-gray-700 hover:bg-gray-100 border-gray-300'
                        }`}
                      >
                        {sz}pt
                      </button>
                    ))}
                  </div>
                </div>
              </div>

              <div>
                <label className="text-gray-600 block text-xs font-medium mb-1">
                  គម្លាតអក្សរពីលើចុះក្រោម (Line Spacing)៖
                </label>
                <div className="flex items-center gap-2">
                  <input
                    type="number"
                    step="0.05"
                    min="1.0"
                    max="2.5"
                    value={lineSpacing}
                    onChange={(e) => {
                      const val = parseFloat(e.target.value) || 1.6;
                      setLineSpacing(val);
                      try {
                        localStorage.setItem(`${keyPrefix}team_visa_stats_line_spacing`, String(val));
                      } catch {}
                    }}
                    className="w-20 h-8 min-h-[32px] max-h-[32px] box-border border border-gray-300 rounded px-2 py-1 bg-white text-xs font-bold font-times text-gray-800 focus:ring-1 focus:ring-blue-500 focus:outline-none text-center"
                  />
                  <div className="flex items-center gap-1">
                    {[1.1, 1.2, 1.25, 1.3, 1.4, 1.5].map((sp) => (
                      <button
                        key={sp}
                        type="button"
                        onClick={() => {
                          setLineSpacing(sp);
                          try {
                            localStorage.setItem(`${keyPrefix}team_visa_stats_line_spacing`, String(sp));
                          } catch {}
                        }}
                        className={`h-8 px-2 rounded border transition cursor-pointer text-xs flex items-center justify-center ${
                          lineSpacing === sp
                            ? 'bg-indigo-600 text-white font-bold border-indigo-700 shadow-xs'
                            : 'bg-white text-gray-700 hover:bg-gray-100 border-gray-300'
                        }`}
                      >
                        {sp}
                      </button>
                    ))}
                  </div>
                </div>
              </div>

              <div>
                <div className="flex items-center justify-between mb-1">
                  <label className="text-gray-600 block text-xs">កថាខណ្ឌស្នើសុំបញ្ចប់ (កថាខណ្ឌទី១)៖</label>
                  <button
                    type="button"
                    onClick={() => setClosingParagraph1(defaultClosing1)}
                    className="text-[10px] text-blue-600 hover:underline cursor-pointer"
                  >
                    លំនាំដើម
                  </button>
                </div>
                <textarea
                  rows={2}
                  value={closingParagraph1}
                  onChange={(e) => setClosingParagraph1(e.target.value)}
                  className="w-full border border-gray-300 rounded px-2.5 py-1 bg-white text-xs text-gray-800 focus:ring-1 focus:ring-blue-500 focus:outline-none"
                />
              </div>

              <div>
                <div className="flex items-center justify-between mb-1">
                  <label className="text-gray-600 block text-xs">កថាខណ្ឌគោរពបញ្ចប់ (កថាខណ្ឌទី២)៖</label>
                  <button
                    type="button"
                    onClick={() => setClosingParagraph2(defaultClosing2)}
                    className="text-[10px] text-blue-600 hover:underline cursor-pointer"
                  >
                    លំនាំដើម
                  </button>
                </div>
                <input
                  type="text"
                  value={closingParagraph2}
                  onChange={(e) => setClosingParagraph2(e.target.value)}
                  className="w-full h-8 min-h-[32px] max-h-[32px] box-border border border-gray-300 rounded px-2.5 py-1 bg-white text-xs text-gray-800 focus:ring-1 focus:ring-blue-500 focus:outline-none"
                />
              </div>
            </div>

            {/* Col 3: Signatures & Margins */}
            <div className="space-y-2">
              <div className="font-bold text-gray-800 border-b pb-1">៣. កាលបរិច្ឆេទ & ហត្ថលេខា & គែមទំព័រ</div>
              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="text-gray-600 block text-xs mb-1">កាលបរិច្ឆេទចុះហត្ថលេខា៖</label>
                  <CustomDatePicker
                    value={signingDate}
                    onChange={setSigningDate}
                    className="h-8 min-h-[32px] max-h-[32px] box-border border border-gray-300 rounded px-2 py-1 bg-white text-xs font-semibold text-gray-800 focus:ring-1 focus:ring-blue-500 focus:outline-none"
                  />
                </div>
                <div>
                  <label className="text-gray-600 block text-xs mb-1">ទីកន្លែង៖</label>
                  <input
                    type="text"
                    value={signingLocation}
                    onChange={(e) => setSigningLocation(e.target.value)}
                    className="w-full h-8 min-h-[32px] max-h-[32px] box-border border border-gray-300 rounded px-2.5 py-1 bg-white text-xs text-gray-800 focus:ring-1 focus:ring-blue-500 focus:outline-none"
                  />
                </div>
              </div>
              <div>
                <label className="text-gray-600 block text-xs mb-1">កាលបរិច្ឆេទចន្ទគតិ (កែប្រែដោយផ្ទាល់ប្រសិនបើចង់)៖</label>
                <input
                  type="text"
                  value={customLunarDate}
                  onChange={(e) => setCustomLunarDate(e.target.value)}
                  placeholder={khmerLunarDate}
                  className="w-full h-8 min-h-[32px] max-h-[32px] box-border border border-gray-300 rounded px-2.5 py-1 bg-white text-xs text-gray-800 focus:ring-1 focus:ring-blue-500 focus:outline-none"
                />
              </div>
              <div>
                <label className="text-gray-600 block text-xs mb-1">កាលបរិច្ឆេទសុរិយគតិ (កែប្រែដោយផ្ទាល់ប្រសិនបើចង់)៖</label>
                <input
                  type="text"
                  value={customSolarDate}
                  onChange={(e) => setCustomSolarDate(e.target.value)}
                  placeholder={khmerSolarDate}
                  className="w-full h-8 min-h-[32px] max-h-[32px] box-border border border-gray-300 rounded px-2.5 py-1 bg-white text-xs text-gray-800 focus:ring-1 focus:ring-blue-500 focus:outline-none"
                />
              </div>
              <div>
                <div className="flex items-center justify-between mb-1">
                  <label className="text-gray-600 block text-xs">តួនាទីអ្នកចុះហត្ថលេខា៖</label>
                  <button
                    type="button"
                    onClick={() => setChiefTitle(defaultChiefTitle)}
                    className="text-[10px] text-blue-600 hover:underline cursor-pointer"
                  >
                    លំនាំដើម
                  </button>
                </div>
                <input
                  type="text"
                  value={chiefTitle}
                  onChange={(e) => setChiefTitle(e.target.value)}
                  className="w-full h-8 min-h-[32px] max-h-[32px] box-border border border-gray-300 rounded px-2.5 py-1 bg-white text-xs text-gray-800 focus:ring-1 focus:ring-blue-500 focus:outline-none"
                />
              </div>
              <div>
                <label className="text-gray-600 block text-xs mb-1">ឈ្មោះអ្នកចុះហត្ថលេខា (ជម្រើស)៖</label>
                <input
                  type="text"
                  value={chiefName}
                  onChange={(e) => setChiefName(e.target.value)}
                  placeholder="ទុកទំនេរ ឬ បញ្ចូលឈ្មោះ"
                  className="w-full h-8 min-h-[32px] max-h-[32px] box-border border border-gray-300 rounded px-2.5 py-1 bg-white text-xs text-gray-800 focus:ring-1 focus:ring-blue-500 focus:outline-none"
                />
              </div>

              {/* Margins */}
              <div className="grid grid-cols-4 gap-1 pt-1">
                <div>
                  <label className="text-[10px] text-gray-500 block mb-0.5 text-center">ឆ្វេង (cm)</label>
                  <input
                    type="number"
                    step="0.1"
                    value={pageLeftMargin}
                    onChange={(e) => setPageLeftMargin(parseFloat(e.target.value) || 2.8)}
                    className="w-full h-8 min-h-[32px] max-h-[32px] box-border border border-gray-300 rounded px-1 py-1 bg-white text-xs text-center text-gray-800 focus:ring-1 focus:ring-blue-500 focus:outline-none font-times"
                  />
                </div>
                <div>
                  <label className="text-[10px] text-gray-500 block mb-0.5 text-center">ស្តាំ (cm)</label>
                  <input
                    type="number"
                    step="0.1"
                    value={pageRightMargin}
                    onChange={(e) => setPageRightMargin(parseFloat(e.target.value) || 1.3)}
                    className="w-full h-8 min-h-[32px] max-h-[32px] box-border border border-gray-300 rounded px-1 py-1 bg-white text-xs text-center text-gray-800 focus:ring-1 focus:ring-blue-500 focus:outline-none font-times"
                  />
                </div>
                <div>
                  <label className="text-[10px] text-gray-500 block mb-0.5 text-center">លើ (cm)</label>
                  <input
                    type="number"
                    step="0.1"
                    value={pageTopMargin}
                    onChange={(e) => setPageTopMargin(parseFloat(e.target.value) || 1.0)}
                    className="w-full h-8 min-h-[32px] max-h-[32px] box-border border border-gray-300 rounded px-1 py-1 bg-white text-xs text-center text-gray-800 focus:ring-1 focus:ring-blue-500 focus:outline-none font-times"
                  />
                </div>
                <div>
                  <label className="text-[10px] text-gray-500 block mb-0.5 text-center">ក្រោម (cm)</label>
                  <input
                    type="number"
                    step="0.1"
                    value={pageBottomMargin}
                    onChange={(e) => setPageBottomMargin(parseFloat(e.target.value) || 0.5)}
                    className="w-full h-8 min-h-[32px] max-h-[32px] box-border border border-gray-300 rounded px-1 py-1 bg-white text-xs text-center text-gray-800 focus:ring-1 focus:ring-blue-500 focus:outline-none font-times"
                  />
                </div>
              </div>
              <div className="pt-2">
                <label className="text-gray-600 block text-xs mb-1">គម្លាតប្លុកហត្ថលេខា (Signature Shift Y - pt)៖</label>
                <input
                  type="number"
                  value={signatureShiftY}
                  onChange={(e) => setSignatureShiftY(parseInt(e.target.value) || 0)}
                  className="w-full h-8 min-h-[32px] max-h-[32px] box-border border border-gray-300 rounded px-2.5 py-1 bg-white text-xs text-gray-800 focus:ring-1 focus:ring-blue-500 focus:outline-none"
                />
              </div>
            </div>
          </div>
        )}
      </div>

      {/* Main Canvas Scroll Area */}
      <div className="flex-1 overflow-auto p-4 md:p-8 flex justify-center print:p-0 print:m-0 print:overflow-visible">
        <div
          style={{
            transform: `scale(${zoomScale / 100})`,
            transformOrigin: 'top center',
            transition: 'transform 0.15s ease-out',
          }}
          className="print:transform-none"
        >
          {/* Banner if team has no eligible usage for selected category */}
          {!hasEligibleUsage && (
            <div className="w-[210mm] mx-auto mb-3 px-3.5 py-2.5 bg-amber-50/95 border border-amber-300 text-amber-900 rounded-lg text-xs flex items-center justify-between gap-3 font-siemreap shadow-xs print:hidden">
              <div className="flex items-center gap-2">
                <AlertCircle className="w-4 h-4 text-amber-600 shrink-0" />
                <span>
                  <strong>សម្គាល់៖</strong> ក្រុម «<strong>{selectedTeam}</strong>» មិនមានទិន្នន័យប្រើប្រាស់{' '}
                  {selectedCategory === 'sticker' ? 'Sticker' : selectedCategory === 'cea' ? 'cEA' : 'Sticker ឬ cEA'}{' '}
                  ក្នុងកាលបរិច្ឆេទនេះទេ ដូច្នេះតារាងលទ្ធផលត្រូវបានកំណត់មិនបង្ហាញឡើយ (បង្ហាញសញ្ញា «-» ទាំងអស់)។
                </span>
              </div>
              <span className="text-[11px] px-2 py-0.5 bg-amber-200/80 text-amber-950 rounded font-semibold whitespace-nowrap">
                {selectedCategory === 'sticker' ? 'គ្មានប្រើ Sticker' : selectedCategory === 'cea' ? 'គ្មានប្រើ cEA' : 'គ្មានទិន្នន័យ'}
              </span>
            </div>
          )}

          {/* A4 Paper Sheet */}
          <div
            id="team-visa-issuance-stats-report-pdf"
            ref={reportContainerRef}
            style={{
              width: '210mm',
              minHeight: '297mm',
              paddingLeft: `${pageLeftMargin}cm`,
              paddingRight: `${pageRightMargin}cm`,
              paddingTop: `${pageTopMargin}cm`,
              paddingBottom: `${pageBottomMargin}cm`,
              boxSizing: 'border-box',
              fontSize: `${fontSizePt}pt`,
            }}
            className="bg-white shadow-xl text-black font-siemreap relative print:shadow-none print:m-0 print:w-[210mm] print:min-h-[297mm]"
          >
            {/* 1. Header Section */}
            <div
              className="flex justify-between items-start text-black"
              style={{
                marginBottom: `${Math.max(2, Math.round((lineSpacing - 0.7) * 4))}px`
              }}
            >
              {/* Left Column: Organization Hierarchy */}
              <div className="text-center inline-block text-black">
                {/* Spacer matching height of ព្រះរាជាណាចក្រកម្ពុជា so ក្រសួងមហាផ្ទៃ aligns with ជាតិ សាសនា ព្រះមហាក្សត្រ */}
                <p
                  className="font-moul text-[12pt] whitespace-nowrap invisible select-none pointer-events-none m-0 p-0"
                  aria-hidden="true"
                  style={{ lineHeight: Math.max(0.95, parseFloat((lineSpacing * 0.85).toFixed(2))) }}
                >
                  ព្រះរាជាណាចក្រកម្ពុជា
                </p>
                <div
                  className="flex flex-col items-center"
                  style={{ gap: `${Math.max(0, Math.round((lineSpacing - 1.15) * 4))}px` }}
                >
                  <p
                    className="font-moul text-[12pt] whitespace-nowrap m-0 p-0"
                    style={{
                      fontFamily: "'Khmer OS Muol Light', 'Khmer OS Moul Light', 'Khmer OS Muol', 'Moul', serif",
                      lineHeight: Math.max(0.95, parseFloat((lineSpacing * 0.85).toFixed(2)))
                    }}
                  >
                    {ministryName}
                  </p>
                  <p
                    className="font-moul text-[12pt] whitespace-nowrap m-0 p-0"
                    style={{
                      fontFamily: "'Khmer OS Muol Light', 'Khmer OS Moul Light', 'Khmer OS Muol', 'Moul', serif",
                      lineHeight: Math.max(0.95, parseFloat((lineSpacing * 0.85).toFixed(2)))
                    }}
                  >
                    {deptName}
                  </p>
                  <p
                    className="font-moul text-[12pt] whitespace-nowrap m-0 p-0"
                    style={{
                      fontFamily: "'Khmer OS Muol Light', 'Khmer OS Moul Light', 'Khmer OS Muol', 'Moul', serif",
                      lineHeight: Math.max(0.95, parseFloat((lineSpacing * 0.85).toFixed(2)))
                    }}
                  >
                    {firstGateDeptName}
                  </p>
                  <p
                    className="font-moul text-[12pt] whitespace-nowrap m-0 p-0"
                    style={{
                      fontFamily: "'Khmer OS Muol Light', 'Khmer OS Moul Light', 'Khmer OS Muol', 'Moul', serif",
                      lineHeight: Math.max(0.95, parseFloat((lineSpacing * 0.85).toFixed(2)))
                    }}
                  >
                    {officeName}
                  </p>
                  <p
                    className="font-bold whitespace-nowrap m-0 p-0 font-siemreap"
                    style={{
                      fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif",
                      fontSize: '10pt',
                      lineHeight: Math.max(0.95, parseFloat((lineSpacing * 0.85).toFixed(2)))
                    }}
                  >
                    {teamHeaderName}
                  </p>
                  <div
                    className="flex justify-center text-black"
                    style={{ marginTop: `${Math.max(1, Math.round(lineSpacing * 1.5))}px` }}
                  >
                    <TacteingLine
                      type={tacteingSettings.type}
                      customImage={tacteingSettings.customImage}
                      width={110}
                      height={12}
                      className="text-black"
                    />
                  </div>
                </div>
              </div>

              {/* Right Column: Kingdom Motto & Tacteing */}
              <div
                className="text-center inline-block text-black"
                style={{
                  display: 'inline-flex',
                  flexDirection: 'column',
                  alignItems: 'center',
                  gap: `${Math.max(0, Math.round((lineSpacing - 1.15) * 4))}px`
                }}
              >
                <p
                  className="font-moul text-[12pt] whitespace-nowrap m-0 p-0"
                  style={{
                    fontFamily: "'Khmer OS Muol Light', 'Khmer OS Moul Light', 'Khmer OS Muol', 'Moul', serif",
                    lineHeight: Math.max(0.95, parseFloat((lineSpacing * 0.85).toFixed(2)))
                  }}
                >
                  ព្រះរាជាណាចក្រកម្ពុជា
                </p>
                <p
                  className="font-moul text-[12pt] whitespace-nowrap m-0 p-0"
                  style={{
                    fontFamily: "'Khmer OS Muol Light', 'Khmer OS Moul Light', 'Khmer OS Muol', 'Moul', serif",
                    lineHeight: Math.max(0.95, parseFloat((lineSpacing * 0.85).toFixed(2)))
                  }}
                >
                  ជាតិ សាសនា ព្រះមហាក្សត្រ
                </p>
                <div
                  className="flex justify-center text-black"
                  style={{ marginTop: `${Math.max(1, Math.round(lineSpacing * 1.5))}px` }}
                >
                  <TacteingLine
                    type={tacteingSettings.type}
                    customImage={tacteingSettings.customImage}
                    width={130}
                    height={14}
                    className="text-black"
                  />
                </div>
              </div>
            </div>

            {/* 2. Salutation (គោរពជូន) */}
            <div
              className="text-center flex flex-col items-center"
              style={{
                marginTop: `${Math.max(1, Math.round((lineSpacing - 0.7) * 3))}px`,
                marginBottom: `${Math.max(2, Math.round((lineSpacing - 0.7) * 4))}px`,
                gap: '0px'
              }}
            >
              <p
                className="font-moul text-[14pt] m-0 p-0"
                style={{
                  fontFamily: "'Khmer OS Muol Light', 'Khmer OS Moul Light', 'Khmer OS Muol', 'Moul', serif",
                  fontSize: '14pt',
                  lineHeight: Math.max(1.0, parseFloat((lineSpacing * 0.88).toFixed(2)))
                }}
              >
                {salutationTop}
              </p>
              <p
                className="font-moul text-[12pt] m-0 p-0"
                style={{
                  fontFamily: "'Khmer OS Muol Light', 'Khmer OS Moul Light', 'Khmer OS Muol', 'Moul', serif",
                  fontSize: '12pt',
                  lineHeight: Math.max(1.0, parseFloat((lineSpacing * 0.88).toFixed(2)))
                }}
              >
                {salutationRecipient}
              </p>
            </div>

            {/* 3. Details Block (កម្មវត្ថុ) */}
            <div
              className="font-siemreap mb-1"
              style={{
                lineHeight: Math.max(lineSpacing * 1.35, 1.72),
                fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif",
                fontSize: `${fontSizePt}pt`,
                marginTop: '5px'
              }}
            >
              <div className="flex items-baseline" style={{ lineHeight: Math.max(lineSpacing * 1.35, 1.72) }}>
                <div style={{ width: `${tabIndentPx}px` }} className="shrink-0 flex items-baseline justify-between pr-2">
                  <span className="font-moul" style={{ fontFamily: "'Khmer OS Muol Light', 'Khmer OS Moul Light', 'Khmer OS Muol', 'Moul', serif", fontSize: `${fontSizePt}pt` }}>កម្មវត្ថុ</span>
                  <span className="font-bold font-siemreap" style={{ fontSize: `${fontSizePt}pt` }}>៖</span>
                </div>
                <div className="text-justify font-siemreap flex-1" style={{ lineHeight: Math.max(lineSpacing * 1.35, 1.72), fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif", fontSize: `${fontSizePt}pt` }}>
                  {effectiveSubjectText}
                </div>
              </div>
            </div>

            {/* 4. Body Content (Indented so ក., ខ., គ. stand aligned with របាយការណ៍ស្តីពី) */}
            <div className="mb-1" style={{ paddingLeft: `${tabIndentPx}px`, fontSize: `${fontSizePt}pt`, marginTop: `${paragraphShiftY}pt` }}>
              {/* Group ក. ប្រភេទទិដ្ឋាការបង់អាការ */}
              <div>
                <div className="font-moul text-gray-950 mb-1 select-none" style={{ fontFamily: "'Khmer OS Muol Light', 'Khmer OS Moul Light', 'Khmer OS Muol', 'Moul', serif", fontSize: `${fontSizePt}pt`, lineHeight: lineSpacing }}>
                  {sectionATitle}
                </div>

                <div className="flex flex-col pl-6">
                  {TAXABLE_VISA_TYPES.map((item) => {
                    const count = visaCounts[item.key] || 0;
                    const parts = item.labelKh.split('-');
                    const prefix = parts.length > 1 ? parts[0] : '';
                    const title = parts.length > 1 ? parts.slice(1).join('-') : item.labelKh;

                    return (
                      <div
                        key={item.key}
                        className="flex items-center"
                        style={{
                          fontSize: `${fontSizePt}pt`,
                          lineHeight: lineSpacing,
                          paddingTop: '1.5px',
                          paddingBottom: '1.5px',
                        }}
                      >
                        {/* Visa Label (Fixed Width) */}
                        <div className="w-[185px] text-gray-950 select-none flex items-center font-siemreap" style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}>
                          {prefix ? (
                            <>
                              <span className="font-times" style={{ fontFamily: "'Times New Roman', Times, serif" }}>{prefix}-</span>
                              <span>{title}</span>
                            </>
                          ) : (
                            <span>{title}</span>
                          )}
                        </div>

                        {/* Code (e.g. T, T1, E) */}
                        <div className="w-[45px] font-bold text-gray-950 font-times" style={{ fontFamily: "'Times New Roman', Times, serif" }}>
                          {item.code}
                        </div>

                        {/* Prefix 'ចំនួន=' */}
                        <div className="w-[70px] text-right text-gray-900 select-none font-siemreap" style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}>
                          ចំនួន<span className="font-times" style={{ fontFamily: "'Times New Roman', Times, serif" }}>=</span>
                        </div>

                        {/* Number Quantity */}
                        <div className="w-[120px] text-right font-medium pr-2">
                          <span className="font-bold text-black tracking-wide font-times" style={{ fontFamily: "'Times New Roman', Times, serif" }}>
                            {isResultShown && hasEligibleUsage ? count.toLocaleString('en-US') : '-'}
                          </span>
                        </div>

                        {/* Unit 'នាក់' */}
                        <div className="w-[40px] text-gray-950 select-none font-siemreap" style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}>
                          នាក់
                        </div>
                      </div>
                    );
                  })}
                </div>
              </div>

              {/* Group ខ. ប្រភេទទិដ្ឋាការមិនបង់អាការ */}
              <div style={{ marginTop: '4px' }}>
                <div className="font-moul text-gray-950 mb-1 select-none" style={{ fontFamily: "'Khmer OS Muol Light', 'Khmer OS Moul Light', 'Khmer OS Muol', 'Moul', serif", fontSize: `${fontSizePt}pt`, lineHeight: lineSpacing }}>
                  {sectionBTitle}
                </div>

                <div className="flex flex-col pl-6">
                  {NON_TAXABLE_VISA_TYPES.map((item) => {
                    const count = visaCounts[item.key] || 0;
                    const parts = item.labelKh.split('-');
                    const prefix = parts.length > 1 ? parts[0] : '';
                    const title = parts.length > 1 ? parts.slice(1).join('-') : item.labelKh;

                    return (
                      <div
                        key={item.key}
                        className="flex items-center"
                        style={{
                          fontSize: `${fontSizePt}pt`,
                          lineHeight: lineSpacing,
                          paddingTop: '1.5px',
                          paddingBottom: '1.5px',
                        }}
                      >
                        {/* Visa Label (Fixed Width) */}
                        <div className="w-[185px] text-gray-950 select-none flex items-center font-siemreap" style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}>
                          {prefix ? (
                            <>
                              <span className="font-times" style={{ fontFamily: "'Times New Roman', Times, serif" }}>{prefix}-</span>
                              <span>{title}</span>
                            </>
                          ) : (
                            <span>{title}</span>
                          )}
                        </div>

                        {/* Code (e.g. K, A, B, C) */}
                        <div className="w-[45px] font-bold text-gray-950 font-times" style={{ fontFamily: "'Times New Roman', Times, serif" }}>
                          {item.code}
                        </div>

                        {/* Prefix 'ចំនួន=' */}
                        <div className="w-[70px] text-right text-gray-900 select-none font-siemreap" style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}>
                          ចំនួន<span className="font-times" style={{ fontFamily: "'Times New Roman', Times, serif" }}>=</span>
                        </div>

                        {/* Number Quantity */}
                        <div className="w-[120px] text-right font-medium pr-2">
                          <span className="font-bold text-black tracking-wide font-times" style={{ fontFamily: "'Times New Roman', Times, serif" }}>
                            {isResultShown && hasEligibleUsage ? count.toLocaleString('en-US') : '-'}
                          </span>
                        </div>

                        {/* Unit 'នាក់' */}
                        <div className="w-[40px] text-gray-950 select-none font-siemreap" style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}>
                          នាក់
                        </div>
                      </div>
                    );
                  })}

                  {/* Grand Total Row - Exactly aligned with items above */}
                  <div
                    className="flex items-center"
                    style={{
                      fontSize: `${fontSizePt}pt`,
                      lineHeight: lineSpacing,
                      paddingTop: '1.5px',
                      paddingBottom: '1.5px',
                    }}
                  >
                    <div className="w-[230px] text-right pr-3">
                      <span className="font-moul text-gray-950" style={{ fontFamily: "'Khmer OS Muol Light', 'Khmer OS Moul Light', 'Khmer OS Muol', 'Moul', serif", fontSize: `${fontSizePt}pt` }}>សរុប</span>
                    </div>
                    <div className="w-[70px] text-right text-gray-900 select-none font-siemreap" style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}>
                      ចំនួន<span className="font-times" style={{ fontFamily: "'Times New Roman', Times, serif" }}>=</span>
                    </div>
                    <div className="w-[120px] text-right font-bold text-black pr-2 tracking-wide font-times" style={{ fontFamily: "'Times New Roman', Times, serif", fontSize: `${fontSizePt}pt` }}>
                      {isResultShown && hasEligibleUsage ? grandTotal.toLocaleString('en-US') : '-'}
                    </div>
                    <div className="w-[40px] text-gray-950 font-bold select-none font-siemreap" style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif", fontSize: `${fontSizePt}pt` }}>
                      នាក់
                    </div>
                  </div>
                </div>
              </div>

              {/* Group គ. ចំណុចគួរអោយកត់សម្គាល់ */}
              <div className="flex items-baseline gap-4" style={{ fontSize: `${fontSizePt}pt`, lineHeight: lineSpacing, marginTop: '4px' }}>
                <div className="font-moul text-gray-950 select-none shrink-0" style={{ fontFamily: "'Khmer OS Muol Light', 'Khmer OS Moul Light', 'Khmer OS Muol', 'Moul', serif", fontSize: `${fontSizePt}pt` }}>
                  គ. ចំណុចគួរអោយកត់សម្គាល់
                </div>
                <div className="text-gray-900 font-siemreap" style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif", fontSize: `${fontSizePt}pt` }}>
                  {remarksNote}
                </div>
              </div>
            </div>

            {/* 5. Closing Formulas */}
            <div
              className="text-justify font-siemreap"
              style={{
                textIndent: `${tabIndentPx}px`,
                fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif",
                fontSize: `${fontSizePt}pt`,
                lineHeight: lineSpacing,
                marginTop: '4px',
                marginBottom: '4px'
              }}
            >
              <p className="m-0" style={{ lineHeight: lineSpacing, marginBottom: '2px' }}>{renderClosingTextWithMuol(closingParagraph1)}</p>
              <p className="m-0" style={{ lineHeight: lineSpacing }}>{renderClosingTextWithMuol(closingParagraph2)}</p>
            </div>

            {/* 6. Date & Signature Block */}
            <div className="flex justify-end" style={{ marginTop: `${signatureShiftY}pt` }}>
              <div
                className="text-center inline-flex flex-col items-center min-w-[340px]"
                style={{
                  fontSize: `${fontSizePt}pt`,
                  gap: '2px'
                }}
              >
                {lunarLines.map((line, idx) => (
                  <div
                    key={idx}
                    className="text-gray-950 font-siemreap whitespace-nowrap m-0 text-center"
                    style={{
                      fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif",
                      fontSize: `${fontSizePt}pt`,
                      lineHeight: lineSpacing
                    }}
                  >
                    {line}
                  </div>
                ))}
                <div
                  className="text-gray-950 font-siemreap whitespace-nowrap m-0 text-center"
                  style={{
                    fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif",
                    fontSize: `${fontSizePt}pt`,
                    lineHeight: lineSpacing
                  }}
                >
                  {displaySolarDate}
                </div>
                <div
                  className="font-moul text-gray-950 whitespace-nowrap m-0 text-center"
                  style={{
                    fontFamily: "'Khmer OS Muol Light', 'Khmer OS Moul Light', 'Khmer OS Muol', 'Moul', serif",
                    fontSize: `${fontSizePt}pt`,
                    lineHeight: lineSpacing
                  }}
                >
                  {chiefTitle}
                </div>

                {/* Vertical space for physical or stamp signature */}
                <div className="h-[44px]"></div>

                {/* Signer Officer Name (if specified) */}
                {chiefName && (
                  <div
                    className="font-bold text-gray-950 font-siemreap whitespace-nowrap m-0 text-center"
                    style={{
                      fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif",
                      fontSize: `${fontSizePt}pt`,
                      lineHeight: lineSpacing
                    }}
                  >
                    {chiefName}
                  </div>
                )}
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};
