import React, { useState, useEffect, useRef, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { CategoriesState, Officer, StockRecord, UserRole } from '../types';
import {
  Printer,
  Download,
  Calendar,
  Layers,
  ArrowLeft,
  Sliders,
  Sparkles,
  Edit3,
  Check,
  X,
  FileSpreadsheet,
  ChevronLeft,
  ChevronRight,
  CalendarCheck,
  Home,
} from 'lucide-react';
import html2canvas from 'html2canvas';
import { jsPDF } from 'jspdf';
import { getKhmerLunarDate, getKhmerSolarParts, toKhmerNum } from '../utils/khmerCalendar';
import { sanitizeDocumentForHtml2Canvas, exportElementToPdf } from '../utils/pdfExportHelper';
import { printA4Document } from '../utils/printHelper';
import { CustomDatePicker } from './CustomDatePicker';
import { TacteingLine, TacteingType, TacteingControlSelector, saveTacteingSettings, getSavedTacteingSettings } from './TacteingLink';

export interface StickerOfficeStockReportProps {
  stockRecords: StockRecord[];
  categories?: CategoriesState;
  officers?: Officer[];
  currentRole?: UserRole;
  userName?: string;
  onClose?: () => void;
}

// Initial baseline stock for each sticker visa type (Default 2026)
const INITIAL_STICKER_BASELINES: Record<string, number> = {
  T: 73000,
  T1: 2100,
  T2: 7250,
  T3: 3750,
  E: 111400,
  E1: 2750,
  E2: 4250,
  E3: 3700,
  D: 15850,
  K: 12050,
  A: 4050,
  B: 5550,
  C: 8550,
};

// Baseline for 30 Nov 2018 (matching official PDF report)
const DEC_2018_STICKER_BASELINES: Record<string, number> = {
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

const DEFAULT_VISA_TYPES = ['T', 'T1', 'T2', 'T3', 'E', 'E1', 'E2', 'E3', 'D', 'K', 'A', 'B', 'C'];

const KHMER_MONTH_NAMES = [
  'មករា (ខែ ០១)',
  'កុម្ភៈ (ខែ ០២)',
  'មីនា (ខែ ០៣)',
  'មេសា (ខែ ០៤)',
  'ឧសភា (ខែ ០៥)',
  'មិថុនា (ខែ ០៦)',
  'កក្កដា (ខែ ០៧)',
  'សីហា (ខែ ០៨)',
  'កញ្ញា (ខែ ០៩)',
  'តុលា (ខែ ១០)',
  'វិច្ឆិកា (ខែ ១១)',
  'ធ្នូ (ខែ ១២)',
];

const RIGHT_SIGNATURE_OPTIONS = [
  'អ្នកធ្វើតារាង',
  'នាយផ្នែក',
  'ជ.នាយផ្នែក',
  'ប្រធានក្រុមផ្តល់ទិដ្ឋាការ',
  'ជ.ប្រធានក្រុមផ្តល់ទិដ្ឋាការ',
  'ប្រធានក្រុមផ្តល់ទិដ្ឋាការស្តីទី',
];

// Helper function to convert OKLAB / OKLCH color strings to standard RGB/RGBA strings for html2canvas
function oklabToRgb(L: number, aVal: number, bVal: number, alpha: number = 1): string {
  const l_ = L + 0.3963377774 * aVal + 0.2158037573 * bVal;
  const m_ = L - 0.1055613458 * aVal - 0.0638541728 * bVal;
  const s_ = L - 0.0894841775 * aVal - 0.1291986507 * bVal;

  const l3 = l_ * l_ * l_;
  const m3 = m_ * m_ * m_;
  const s3 = s_ * s_ * s_;

  const rLin = +4.0767416621 * l3 - 3.3077115913 * m3 + 0.2309699292 * s3;
  const gLin = -1.2684380046 * l3 + 2.6097574011 * m3 - 0.3413193965 * s3;
  const bLin = -0.0041960863 * l3 - 0.7034186147 * m3 + 1.7076147010 * s3;

  const gamma = (c: number) => {
    const clamped = Math.max(0, Math.min(1, c));
    return clamped <= 0.0031308 ? 12.92 * clamped : 1.055 * Math.pow(clamped, 1 / 2.4) - 0.055;
  };

  const r = Math.round(gamma(rLin) * 255);
  const g = Math.round(gamma(gLin) * 255);
  const b = Math.round(gamma(bLin) * 255);

  if (alpha < 1) {
    return `rgba(${r}, ${g}, ${b}, ${Number(alpha.toFixed(3))})`;
  }
  return `rgb(${r}, ${g}, ${b})`;
}

function convertModernCssColors(input: string): string {
  if (!input || typeof input !== 'string') return input;
  let result = input;

  if (result.includes('oklch')) {
    const oklchRegex = /oklch\(\s*([-\d.%]+)(?:[\s,]+)([-\d.%]+)(?:[\s,]+)([-\d.deg%radturn]+)(?:\s*[\/,]\s*([-\d.%]+))?\s*\)/gi;
    result = result.replace(oklchRegex, (fullMatch, lStr, cStr, hStr, aStr) => {
      try {
        let L = parseFloat(lStr);
        if (lStr.endsWith('%')) L = L / 100;
        let C = parseFloat(cStr);
        if (cStr.endsWith('%')) C = C / 100;
        let H = parseFloat(hStr.replace(/(deg|rad|turn)/gi, ''));
        if (hStr.toLowerCase().endsWith('rad')) H = (H * 180) / Math.PI;
        else if (hStr.toLowerCase().endsWith('turn')) H = H * 360;

        let alpha = 1;
        if (aStr) {
          const alphaStr = aStr.replace(/[\s\/]/g, '');
          alpha = parseFloat(alphaStr);
          if (alphaStr.endsWith('%')) alpha = alpha / 100;
        }

        const hRad = (H * Math.PI) / 180;
        const aVal = C * Math.cos(hRad);
        const bVal = C * Math.sin(hRad);

        if (isNaN(L) || isNaN(aVal) || isNaN(bVal)) return fullMatch;
        return oklabToRgb(L, aVal, bVal, alpha);
      } catch {
        return fullMatch;
      }
    });
  }

  if (result.includes('oklab')) {
    const oklabRegex = /oklab\(\s*([-\d.%]+)(?:[\s,]+)([-\d.%]+)(?:[\s,]+)([-\d.%]+)(?:\s*[\/,]\s*([-\d.%]+))?\s*\)/gi;
    result = result.replace(oklabRegex, (fullMatch, lStr, aStr, bStr, alphaStr) => {
      try {
        let L = parseFloat(lStr);
        if (lStr.endsWith('%')) L = L / 100;
        let aVal = parseFloat(aStr);
        let bVal = parseFloat(bStr);

        let alpha = 1;
        if (alphaStr) {
          alpha = parseFloat(alphaStr);
          if (alphaStr.endsWith('%')) alpha = alpha / 100;
        }

        if (isNaN(L) || isNaN(aVal) || isNaN(bVal)) return fullMatch;
        return oklabToRgb(L, aVal, bVal, alpha);
      } catch {
        return fullMatch;
      }
    });
  }

  return result;
}

export const StickerOfficeStockReport: React.FC<StickerOfficeStockReportProps> = ({
  stockRecords,
  categories,
  officers,
  currentRole,
  userName = '',
  onClose,
}) => {
  // Date Filters (Defaults to current or selected period)
  const [startDate, setStartDate] = useState<string>('2026-04-01');
  const [endDate, setEndDate] = useState<string>('2026-07-09');
  const [reportDate, setReportDate] = useState<string>('2026-07-09');

  // Previous Period Date string (e.g., "៣១ មីនា ២០២៦")
  const [prevPeriodStr, setPrevPeriodStr] = useState<string>('៣១ មីនា ២០២៦');
  const [isManualPrevPeriodStr, setIsManualPrevPeriodStr] = useState<boolean>(false);

  // Custom Report Title
  const [reportTitle, setReportTitle] = useState<string>(
    'តារាងទិដ្ឋាការស្អិត តាមប្រភេទនីមួយៗ បើកពីក១ និងផ្តល់តាមបណ្តាប៉ុស្តិ៍ច្រកទ្វារអន្តរជាតិទាំង ២៩'
  );

  // Paper Margin Settings (Defaults: top: 0.5cm, bottom: 0.5cm, left: 0.8cm, right: 0.8cm)
  const [showMarginControls, setShowMarginControls] = useState<boolean>(false);
  const [customMargins, setCustomMargins] = useState<{
    top: number;
    bottom: number;
    left: number;
    right: number;
  }>(() => {
    try {
      const saved = localStorage.getItem('sticker_office_report_margins');
      if (saved) {
        const parsed = JSON.parse(saved);
        if (typeof parsed.top === 'number' && typeof parsed.bottom === 'number') {
          return parsed;
        }
      }
    } catch (e) {
      // ignore
    }
    return { top: 0.5, bottom: 0.5, left: 0.8, right: 0.8 };
  });

  const handleMarginChange = (key: 'top' | 'bottom' | 'left' | 'right', val: number) => {
    setCustomMargins((prev) => {
      const updated = { ...prev, [key]: val };
      try {
        localStorage.setItem('sticker_office_report_margins', JSON.stringify(updated));
      } catch (e) {}
      return updated;
    });
  };

  // Signature Controls
  const [showLeftSignature, setShowLeftSignature] = useState<boolean>(false);
  const [leftSignatureTitle, setLeftSignatureTitle] = useState<string>('នាយរងការិយាល័យទទួលបន្ទុក');
  const [rightSignatureTitle, setRightSignatureTitle] = useState<string>('អ្នកធ្វើតារាង');
  const [signerRank, setSignerRank] = useState<string>('អនុសេនីយ៍ឯក');
  const [signerNameOnly, setSignerNameOnly] = useState<string>('អ៊ុក រ័ត្នបញ្ញា');
  const [useRedSignerName, setUseRedSignerName] = useState<boolean>(true);
  const [selectedOfficerId, setSelectedOfficerId] = useState<string>('');
  const [customLunarDate, setCustomLunarDate] = useState<string>('ថ្ងៃព្រហស្បតិ៍ ១០រោច ខែ ឋមាសាឍ ឆ្នាំខាល ចត្វាស័ក ព.ស ២៥៧០');

  // Tacteing Decoration
  const [showTacteingControls, setShowTacteingControls] = useState<boolean>(false);
  const [tacteingType, setTacteingType] = useState<TacteingType>(() => {
    return getSavedTacteingSettings().type || 'image2-classic';
  });
  const [tacteingCustomImage, setTacteingCustomImage] = useState<string | null>(() => {
    return getSavedTacteingSettings().customImage || null;
  });

  // Whether to auto-calculate beginning stock from imported stock records / historical data before startDate
  const [useAutoBeginningStock, setUseAutoBeginningStock] = useState<boolean>(true);

  // Baseline stock balances (for manual mode or fallback)
  const [baselines, setBaselines] = useState<Record<string, number>>(() => {
    const saved = localStorage.getItem('sticker_office_report_baselines');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        const sum = Object.values(parsed).reduce((a: any, b: any) => (Number(a) || 0) + (Number(b) || 0), 0);
        // If it was the old default template (254,250) or empty, prioritize the official 700,200 DEC 2018 baseline
        if (sum === 254250 || sum === 0) {
          return DEC_2018_STICKER_BASELINES;
        }
        return parsed;
      } catch {
        // fallback
      }
    }
    return DEC_2018_STICKER_BASELINES;
  });

  const [isEditingBaselines, setIsEditingBaselines] = useState<boolean>(false);
  const [tempBaselines, setTempBaselines] = useState<Record<string, number>>(baselines);

  // Column Headers Customization
  const [issuedColTitle, setIssuedColTitle] = useState<string>('សន្លឹកទិដ្ឋាការបើកផ្តល់ទៅច្រកទ្វារ');
  const [damagedColTitle, setDamagedColTitle] = useState<string>('សន្លឹកទិដ្ឋាការមិនបានការ ក២');
  const [showOfficeSubLabel, setShowOfficeSubLabel] = useState<boolean>(false);

  const saveBaselines = () => {
    setBaselines(tempBaselines);
    localStorage.setItem('sticker_office_report_baselines', JSON.stringify(tempBaselines));
    setIsEditingBaselines(false);
  };

  const loadDec2018Baselines = () => {
    setTempBaselines(DEC_2018_STICKER_BASELINES);
    setBaselines(DEC_2018_STICKER_BASELINES);
    localStorage.setItem('sticker_office_report_baselines', JSON.stringify(DEC_2018_STICKER_BASELINES));
  };

  const resetBaselinesToDefault = () => {
    setTempBaselines(INITIAL_STICKER_BASELINES);
    setBaselines(INITIAL_STICKER_BASELINES);
    localStorage.setItem('sticker_office_report_baselines', JSON.stringify(INITIAL_STICKER_BASELINES));
  };

  const resetBaselinesToZero = () => {
    const zeroMap: Record<string, number> = {};
    DEFAULT_VISA_TYPES.forEach((vt) => {
      zeroMap[vt] = 0;
    });
    setTempBaselines(zeroMap);
    setBaselines(zeroMap);
    localStorage.setItem('sticker_office_report_baselines', JSON.stringify(zeroMap));
  };

  // Auto-fill officer rank and name when selected from dropdown
  useEffect(() => {
    if (selectedOfficerId && officers) {
      const found = officers.find((o) => o.id === selectedOfficerId);
      if (found) {
        if (found.rankName) setSignerRank(found.rankName);
        if (found.name) setSignerNameOnly(found.name);
      }
    }
  }, [selectedOfficerId, officers]);

  // Handle Date range formatting for solar & lunar
  const startParts = getKhmerSolarParts(startDate);
  const endParts = getKhmerSolarParts(endDate);
  const reportDateParts = getKhmerSolarParts(reportDate);
  const autoLunarStr = getKhmerLunarDate(reportDate);

  // Keep custom Lunar date updated when report date changes unless manually edited
  useEffect(() => {
    if (autoLunarStr) {
      setCustomLunarDate(autoLunarStr);
    }
  }, [reportDate]);

  // Update prevPeriodStr automatically when startDate changes if not manually overridden
  useEffect(() => {
    if (startDate && !isManualPrevPeriodStr) {
      try {
        const dt = new Date(startDate);
        dt.setDate(dt.getDate() - 1);
        const prevIso = dt.toISOString().split('T')[0];
        const pParts = getKhmerSolarParts(prevIso);
        setPrevPeriodStr(`${pParts.khmerDay} ${pParts.khmerMonth} ២០${pParts.khmerYear.slice(-2) || pParts.khmerYear}`);
      } catch {
        // ignore
      }
    }
  }, [startDate, isManualPrevPeriodStr]);

  // Format numbers
  const [useKhmerDigits, setUseKhmerDigits] = useState<boolean>(false);
  const formatNum = (num: number): string => {
    const formatted = (num || 0).toLocaleString();
    return useKhmerDigits ? toKhmerNum(formatted) : formatted;
  };

  // 1. Core cumulative historical calculation:
  // For each visa type:
  // Beginning Stock at startDate = Base Stock + (All net inflows - outflows prior to startDate)
  // This guarantees that Ending Stock of Month N seamlessly becomes Beginning Stock of Month N+1!
  const autoBeginningStockMap = useMemo(() => {
    const map: Record<string, number> = {};
    DEFAULT_VISA_TYPES.forEach((vt) => {
      map[vt] = 0;
    });

    // Check for explicit imported 'ស្តុកចាស់ / សន្និធិដើម' records
    let hasImportedOldStock = false;
    const oldStockFromRecords: Record<string, number> = {};
    const oldStockDateMap: Record<string, string> = {};
    DEFAULT_VISA_TYPES.forEach((vt) => {
      oldStockFromRecords[vt] = 0;
    });

    stockRecords.forEach((r) => {
      if (r.stockType !== 'sticker') return;
      const vt = (r.visaType || '').trim().toUpperCase();
      if (!DEFAULT_VISA_TYPES.includes(vt)) return;

      const isTeamRecord =
        Boolean(r.visaTeamRobokId) ||
        Boolean(r.visaTeamRobokName) ||
        r.operationType === 'damagedTeam' ||
        r.operationType === 'missingTeam' ||
        r.operationType === 'useTeam' ||
        r.operationType === 'issueTeam' ||
        r.operationType === 'returnTeam' ||
        r.operationType === 'oldStockTeam' ||
        r.sourceFrom?.includes('ក្រុម') ||
        r.operationType?.includes('ក្រុម');

      const isOldStockK2 =
        !isTeamRecord &&
        (r.operationType === 'oldStockK2' ||
          r.operationType === 'oldStock' ||
          r.operationType === 'initial' ||
          r.operationType?.toLowerCase().includes('oldstock') ||
          r.operationType?.includes('ស្តុកចាស់') ||
          r.sourceFrom?.includes('ស្តុកចាស់') ||
          r.sourceFrom?.includes('សន្និធិដើម') ||
          r.sourceFrom?.includes('សន្និធិដំបូង') ||
          r.sourceFrom?.includes('ចុងគ្រា') ||
          r.sourceFrom?.toLowerCase().includes('old stock') ||
          r.sourceFrom?.toLowerCase().includes('initial') ||
          r.notes?.includes('ស្តុកចាស់') ||
          r.notes?.includes('សន្និធិដើម'));

      if (isOldStockK2) {
        const sheets = r.totalSheets || r.quantityBundles || 0;
        oldStockFromRecords[vt] = (oldStockFromRecords[vt] || 0) + sheets;
        if (r.date) {
          oldStockDateMap[vt] = r.date;
        }
        hasImportedOldStock = true;
      }
    });

    // Step 1: Assign initial baseline (from imported Old Stock K2 records, or configured baseline)
    DEFAULT_VISA_TYPES.forEach((vt) => {
      if (hasImportedOldStock && oldStockFromRecords[vt] > 0) {
        map[vt] = oldStockFromRecords[vt];
      } else {
        map[vt] = baselines[vt] ?? 0;
      }
    });

    // Step 2: Accumulate all historical office movements before startDate (r.date < startDate)
    stockRecords.forEach((r) => {
      if (r.stockType !== 'sticker') return;
      const vt = (r.visaType || '').trim().toUpperCase();
      if (!DEFAULT_VISA_TYPES.includes(vt)) return;
      if (!r.date || r.date >= startDate) return;

      const sheets = r.totalSheets || r.quantityBundles || 0;
      if (sheets <= 0) return;

      const isTeamRecord =
        Boolean(r.visaTeamRobokId) ||
        Boolean(r.visaTeamRobokName) ||
        r.operationType === 'damagedTeam' ||
        r.operationType === 'missingTeam' ||
        r.operationType === 'useTeam' ||
        r.operationType === 'issueTeam' ||
        r.operationType === 'returnTeam' ||
        r.operationType === 'oldStockTeam' ||
        r.sourceFrom?.includes('ក្រុម') ||
        r.operationType?.includes('ក្រុម');

      const isOldStock =
        !isTeamRecord &&
        (r.operationType === 'oldStockK2' ||
          r.operationType === 'oldStock' ||
          r.operationType === 'initial' ||
          r.operationType?.toLowerCase().includes('oldstock') ||
          r.operationType?.includes('ស្តុកចាស់') ||
          r.sourceFrom?.includes('ស្តុកចាស់') ||
          r.sourceFrom?.includes('សន្និធិដើម') ||
          r.sourceFrom?.includes('សន្និធិដំបូង') ||
          r.sourceFrom?.includes('ចុងគ្រា') ||
          r.sourceFrom?.toLowerCase().includes('old stock') ||
          r.sourceFrom?.toLowerCase().includes('initial') ||
          r.notes?.includes('ស្តុកចាស់') ||
          r.notes?.includes('សន្និធិដើម'));

      // Skip old stock because it is already captured in the initial baseline step
      if (isOldStock) return;

      // If an explicit old stock record was imported on a specific date, only movements AFTER that date should modify it
      const baseOldDate = oldStockDateMap[vt];
      if (baseOldDate && r.date <= baseOldDate) {
        return;
      }

      const isReturnK1Record =
        r.operationType === 'returnK1' ||
        r.sourceFrom === 'បង្វិលក១' ||
        r.sourceFrom === 'ទិដ្ឋាការបង្វិលទៅក១' ||
        (Boolean(r.sourceFrom?.includes('បង្វិល') && r.sourceFrom?.includes('ក១')) && r.operationType !== 'openK1');

      if (
        !isReturnK1Record &&
        (r.operationType === 'returnTeam' ||
          r.operationType === 'returned' ||
          r.sourceFrom?.includes('បង្វិល') ||
          r.operationType?.includes('បង្វិល'))
      ) {
        map[vt] = (map[vt] || 0) + sheets;
      } else if (isReturnK1Record) {
        map[vt] = (map[vt] || 0) - sheets;
      } else if (
        r.operationType === 'transferTeam' ||
        r.operationType === 'transferUseTeam' ||
        r.operationType === 'transfer' ||
        r.operationType?.includes('transfer') ||
        r.sourceFrom?.includes('ផ្ទេរ') ||
        r.operationType?.includes('ផ្ទេរ') ||
        r.notes?.includes('ផ្ទេរ')
      ) {
        // Team-to-team transfers do not affect Office (K2) stock balance
      } else if (r.operationType === 'issueTeam') {
        map[vt] = (map[vt] || 0) - sheets;
      } else if (
        !isTeamRecord &&
        (r.operationType === 'testPrintK2' ||
          r.operationType === 'testK2' ||
          r.sourceFrom?.includes('សាក') ||
          r.sourceFrom?.includes('បោះពុម្ពសាកល្បង') ||
          r.operationType?.includes('សាក'))
      ) {
        map[vt] = (map[vt] || 0) - sheets;
      } else if (
        !isTeamRecord &&
        (r.operationType === 'damaged' ||
          r.operationType === 'invalid' ||
          r.operationType === 'damagedK2' ||
          r.sourceFrom?.includes('ខូចក២') ||
          r.sourceFrom?.includes('មិនបានការ') ||
          r.sourceFrom === 'ទិដ្ឋាការខូចក២' ||
          r.operationType?.includes('ខូច') ||
          r.operationType?.includes('មិនបានការ'))
      ) {
        map[vt] = (map[vt] || 0) - sheets;
      } else if (
        !isTeamRecord &&
        !r.sourceFrom?.includes('ស្តុកចាស់') &&
        !r.sourceFrom?.includes('សន្និធិដើម') &&
        (r.operationType === 'openK1' ||
          r.operationType === 'receive' ||
          r.operationType === 'received' ||
          r.operationType === 'import' ||
          r.operationType === 'in' ||
          r.sourceFrom === 'ក១' ||
          r.sourceFrom?.includes('ក១') ||
          r.sourceFrom?.includes('អគ្គ') ||
          r.sourceFrom?.includes('នាយកដ្ឋាន') ||
          r.sourceFrom?.includes('បើកពី') ||
          r.sourceFrom?.includes('បញ្ចូលស្តុក') ||
          r.notes?.includes('ក១') ||
          r.notes?.includes('បើកពីក១') ||
          r.notes?.includes('បញ្ចូលស្តុក'))
      ) {
        map[vt] = (map[vt] || 0) + sheets;
      }
    });

    return map;
  }, [stockRecords, startDate, baselines]);

  const totalAutoBeginning = useMemo(() => {
    return Object.values(autoBeginningStockMap).reduce((sum: number, v: number) => sum + (v || 0), 0);
  }, [autoBeginningStockMap]);

  // Compute stock report data for each visa type in the active date range [startDate, endDate]
  const reportRows = useMemo(() => {
    return DEFAULT_VISA_TYPES.map((visaType, index) => {
      const autoVal = autoBeginningStockMap[visaType] ?? 0;
      const beginningStock = useAutoBeginningStock ? autoVal : (baselines[visaType] ?? 0);

      const relevantRecords = stockRecords.filter((r) => {
        if (r.stockType !== 'sticker') return false;
        const recordType = (r.visaType || '').trim().toUpperCase();
        if (recordType !== visaType.toUpperCase()) return false;
        if (!r.date || r.date < startDate || r.date > endDate) return false;

        const isOldStock =
          r.operationType === 'oldStockK2' ||
          r.operationType === 'oldStockTeam' ||
          r.operationType === 'oldStock' ||
          r.operationType === 'initial' ||
          r.operationType?.toLowerCase().includes('oldstock') ||
          r.operationType?.includes('ស្តុកចាស់') ||
          r.sourceFrom?.includes('ស្តុកចាស់') ||
          r.sourceFrom?.includes('សន្និធិដើម') ||
          r.sourceFrom?.includes('សន្និធិដំបូង') ||
          r.notes?.includes('ស្តុកចាស់') ||
          r.notes?.includes('សន្និធិដើម');

        // If this record is old stock and already accounted for in beginning stock, skip
        if (isOldStock) return false;

        return true;
      });

      let receivedK1 = 0;
      let returnedFromTeam = 0;
      let issuedToTeam = 0;
      let damaged = 0;
      let testPrintK2 = 0;
      let transferUseTeam = 0;

      relevantRecords.forEach((r) => {
        const sheets = r.totalSheets || r.quantityBundles || 0;

        const isTeamRecord =
          Boolean(r.visaTeamRobokId) ||
          Boolean(r.visaTeamRobokName) ||
          r.operationType === 'damagedTeam' ||
          r.operationType === 'missingTeam' ||
          r.operationType === 'useTeam' ||
          r.operationType === 'issueTeam' ||
          r.operationType === 'returnTeam' ||
          r.operationType === 'transferTeam' ||
          r.sourceFrom?.includes('ក្រុម') ||
          r.operationType?.includes('ក្រុម');

        const isReturnK1Record =
          r.operationType === 'returnK1' ||
          r.sourceFrom === 'បង្វិលក១' ||
          r.sourceFrom === 'ទិដ្ឋាការបង្វិលទៅក១' ||
          (Boolean(r.sourceFrom?.includes('បង្វិល') && r.sourceFrom?.includes('ក១')) && r.operationType !== 'openK1');

        if (
          !isReturnK1Record &&
          (r.operationType === 'returnTeam' ||
            r.operationType === 'returned' ||
            r.sourceFrom?.includes('បង្វិល') ||
            r.operationType?.includes('បង្វិល'))
        ) {
          returnedFromTeam += sheets;
        } else if (
          r.operationType === 'transferTeam' ||
          r.operationType === 'transferUseTeam' ||
          r.operationType === 'transfer' ||
          r.operationType?.includes('transfer') ||
          r.sourceFrom?.includes('ផ្ទេរ') ||
          r.operationType?.includes('ផ្ទេរ') ||
          r.notes?.includes('ផ្ទេរ')
        ) {
          transferUseTeam += sheets;
        } else if (r.operationType === 'issueTeam') {
          issuedToTeam += sheets;
        } else if (
          !isTeamRecord &&
          (r.operationType === 'testPrintK2' ||
            r.operationType === 'testK2' ||
            r.sourceFrom?.includes('សាក') ||
            r.sourceFrom?.includes('បោះពុម្ពសាកល្បង') ||
            r.operationType?.includes('សាក'))
        ) {
          testPrintK2 += sheets;
        } else if (
          !isTeamRecord &&
          (r.operationType === 'damaged' ||
            r.operationType === 'invalid' ||
            r.operationType === 'damagedK2' ||
            r.sourceFrom?.includes('ខូចក២') ||
            r.sourceFrom?.includes('មិនបានការ') ||
            r.sourceFrom === 'ទិដ្ឋាការខូចក២' ||
            r.operationType?.includes('ខូច') ||
            r.operationType?.includes('មិនបានការ'))
        ) {
          damaged += sheets;
        } else if (
          !isTeamRecord &&
          !r.sourceFrom?.includes('ស្តុកចាស់') &&
          !r.sourceFrom?.includes('សន្និធិដើម') &&
          (r.operationType === 'openK1' ||
            r.operationType === 'receive' ||
            r.operationType === 'received' ||
            r.operationType === 'import' ||
            r.operationType === 'in' ||
            r.sourceFrom === 'ក១' ||
            r.sourceFrom?.includes('ក១') ||
            r.sourceFrom?.includes('អគ្គ') ||
            r.sourceFrom?.includes('នាយកដ្ឋាន') ||
            r.sourceFrom?.includes('បើកពី') ||
            r.sourceFrom?.includes('បញ្ចូលស្តុក') ||
            r.notes?.includes('ក១') ||
            r.notes?.includes('បើកពីក១') ||
            r.notes?.includes('បញ្ចូលស្តុក'))
        ) {
          receivedK1 += sheets;
        }
      });

      // User exact formula:
      // សន្និធិចុងគ្រា = សន្និធិដើមគ្រា + សន្លឹកទិដ្ឋាការបើកពីក១ + សន្លឹកទិដ្ឋាការបង្វិលពីក្រុម - សន្លឹកទិដ្ឋាការបើកផ្តល់ទៅច្រកទ្វារ(កុំរាប់បញ្ចូលផ្ទេរការប្រើប្រាស់) - សន្លឹកទិដ្ឋាការមិនបានការក២ - សន្លឹកទិដ្ឋាការបោះពុម្ពសាកល្បង ក២
      const endingStock = beginningStock + receivedK1 + returnedFromTeam - issuedToTeam - damaged - testPrintK2;

      return {
        no: index + 1,
        visaType,
        beginningStock,
        receivedK1,
        returnedFromTeam,
        issuedToTeam,
        damaged,
        testPrintK2,
        transferUseTeam,
        endingStock,
      };
    });
  }, [autoBeginningStockMap, useAutoBeginningStock, baselines, stockRecords, startDate, endDate]);

  // Calculate Totals
  const totals = useMemo(() => {
    return reportRows.reduce(
      (acc, row) => ({
        beginningStock: acc.beginningStock + row.beginningStock,
        receivedK1: acc.receivedK1 + row.receivedK1,
        returnedFromTeam: acc.returnedFromTeam + row.returnedFromTeam,
        issuedToTeam: acc.issuedToTeam + row.issuedToTeam,
        damaged: acc.damaged + row.damaged,
        testPrintK2: acc.testPrintK2 + row.testPrintK2,
        transferUseTeam: acc.transferUseTeam + (row.transferUseTeam || 0),
        endingStock: acc.endingStock + row.endingStock,
      }),
      {
        beginningStock: 0,
        receivedK1: 0,
        returnedFromTeam: 0,
        issuedToTeam: 0,
        damaged: 0,
        testPrintK2: 0,
        transferUseTeam: 0,
        endingStock: 0,
      }
    );
  }, [reportRows]);

  // Sync ending stock to manual baseline
  const syncEndingStockToBaselines = () => {
    const endingMap: Record<string, number> = {};
    reportRows.forEach((r) => {
      endingMap[r.visaType] = r.endingStock;
    });
    setTempBaselines(endingMap);
    setBaselines(endingMap);
    localStorage.setItem('sticker_office_report_baselines', JSON.stringify(endingMap));
  };

  const printAreaRef = useRef<HTMLDivElement>(null);

  const [isExportingExcel, setIsExportingExcel] = useState(false);

  // Export to Excel handler matching Paper A4 layout 100% using ExcelJS
  const handleExportExcel = async () => {
    setIsExportingExcel(true);
    try {
      const ExcelJS = await import('exceljs');
      const workbook = new ExcelJS.Workbook();
      workbook.creator = 'General Department of Immigration';
      workbook.lastModifiedBy = 'General Department of Immigration';
      workbook.created = new Date();
      workbook.modified = new Date();

      const worksheet = workbook.addWorksheet('របាយការណ៍ស្តុកការិយាល័យ', {
        pageSetup: {
          paperSize: 9, // A4
          orientation: 'landscape',
          fitToPage: true,
          fitToWidth: 1,
          fitToHeight: 0, // Automatic height, fit all columns on 1 page width
          horizontalCentered: true,
          verticalCentered: false,
          margins: {
            top: 0.5,
            bottom: 0.5,
            left: 0.5,
            right: 0.5,
            header: 0.2,
            footer: 0.2,
          },
        },
        views: [{ showGridLines: true }],
      });

      // 1. Column widths in exact Pixel equivalents (58px, 85px, 216px x 7)
      worksheet.columns = [
        { key: 'colA', width: 7.6 },  // 1: ល.រ (58 px)
        { key: 'colB', width: 11.4 }, // 2: ប្រភេទ (85 px)
        { key: 'colC', width: 30.1 }, // 3: សន្និធិចុងគ្រា (216 px)
        { key: 'colD', width: 30.1 }, // 4: បើកពីក១ (216 px)
        { key: 'colE', width: 30.1 }, // 5: បង្វិលពីក្រុម (216 px)
        { key: 'colF', width: 30.1 }, // 6: បើកផ្តល់ទៅក្រុម/ច្រកទ្វារ (216 px)
        { key: 'colG', width: 30.1 }, // 7: មិនបានការ (216 px)
        { key: 'colH', width: 30.1 }, // 8: បោះពុម្ពសាកល្បង ក២ (216 px)
        { key: 'colI', width: 30.1 }, // 9: សន្និធិសល់ (216 px)
      ];

      const FONT_MOUL = 'Khmer OS Muol Light';
      const FONT_SIEMREAP = 'Khmer OS Siemreap';

      const thinBorder = {
        top: { style: 'thin' as const, color: { argb: 'FF000000' } },
        left: { style: 'thin' as const, color: { argb: 'FF000000' } },
        bottom: { style: 'thin' as const, color: { argb: 'FF000000' } },
        right: { style: 'thin' as const, color: { argb: 'FF000000' } },
      };

      const totalRowBorder = {
        top: { style: 'thin' as const, color: { argb: 'FF000000' } },
        left: { style: 'thin' as const, color: { argb: 'FF000000' } },
        bottom: { style: 'double' as const, color: { argb: 'FF000000' } },
        right: { style: 'thin' as const, color: { argb: 'FF000000' } },
      };

      const headerFill = {
        type: 'pattern' as const,
        pattern: 'solid' as const,
        fgColor: { argb: 'FFF2F4F7' },
      };

      // --- 1. Cambodian Government Official Header (Row 1 to 7) ---
      // Row 1: Right Side Kingdom (Left is blank)
      worksheet.mergeCells('G1:I1');
      const k1 = worksheet.getCell('G1');
      k1.value = 'ព្រះរាជាណាចក្រកម្ពុជា';
      k1.font = { name: FONT_MOUL, size: 11, bold: true, color: { argb: 'FF000000' } };
      k1.alignment = { horizontal: 'center', vertical: 'middle' };
      worksheet.getRow(1).height = 20;

      // Row 2: Left: Ministry, Right: Motto
      worksheet.mergeCells('A2:D2');
      const m1 = worksheet.getCell('A2');
      m1.value = 'ក្រសួងមហាផ្ទៃ';
      m1.font = { name: FONT_MOUL, size: 11, bold: true, color: { argb: 'FF000000' } };
      m1.alignment = { horizontal: 'center', vertical: 'middle' };

      worksheet.mergeCells('G2:I2');
      const k2 = worksheet.getCell('G2');
      k2.value = 'ជាតិ សាសនា ព្រះមហាក្សត្រ';
      k2.font = { name: FONT_MOUL, size: 11, bold: true, color: { argb: 'FF000000' } };
      k2.alignment = { horizontal: 'center', vertical: 'middle' };
      worksheet.getRow(2).height = 20;

      // Row 3: Left: General Dept, Right: Tacteing under Motto
      worksheet.mergeCells('A3:D3');
      const m2 = worksheet.getCell('A3');
      m2.value = 'អគ្គនាយកដ្ឋានអន្តោប្រវេសន៍';
      m2.font = { name: FONT_MOUL, size: 11, bold: true, color: { argb: 'FF000000' } };
      m2.alignment = { horizontal: 'center', vertical: 'middle' };

      worksheet.mergeCells('G3:I3');
      const tRight = worksheet.getCell('G3');
      tRight.value = '6';
      tRight.font = { name: 'Tacteing', size: 12 };
      tRight.alignment = { horizontal: 'center', vertical: 'middle' };
      worksheet.getRow(3).height = 18;

      // Row 4: Left: Department
      worksheet.mergeCells('A4:D4');
      const m3 = worksheet.getCell('A4');
      m3.value = 'នាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត';
      m3.font = { name: FONT_MOUL, size: 10.5, bold: true, color: { argb: 'FF000000' } };
      m3.alignment = { horizontal: 'center', vertical: 'middle' };
      worksheet.getRow(4).height = 20;

      // Row 5: Left: Office
      worksheet.mergeCells('A5:D5');
      const m4 = worksheet.getCell('A5');
      m4.value = 'ការិយាល័យទិដ្ឋាការចូល';
      m4.font = { name: FONT_MOUL, size: 11, bold: true, color: { argb: 'FF000000' } };
      m4.alignment = { horizontal: 'center', vertical: 'middle' };
      worksheet.getRow(5).height = 20;

      // Row 6: Left: Administration Section
      worksheet.mergeCells('A6:D6');
      const m5 = worksheet.getCell('A6');
      m5.value = 'ផ្នែករដ្ឋបាល';
      m5.font = { name: FONT_MOUL, size: 11, bold: true, color: { argb: 'FF000000' } };
      m5.alignment = { horizontal: 'center', vertical: 'middle' };
      worksheet.getRow(6).height = 20;

      // Row 7: Left: Tacteing under Administration
      worksheet.mergeCells('A7:D7');
      const tLeft = worksheet.getCell('A7');
      tLeft.value = '6';
      tLeft.font = { name: 'Tacteing', size: 12 };
      tLeft.alignment = { horizontal: 'center', vertical: 'middle' };
      worksheet.getRow(7).height = 18;

      // --- 2. Document Title Header (Rows 8 & 9) ---
      // Row 8: Report Title (13pt Bold)
      worksheet.mergeCells('A8:I8');
      const titleCell = worksheet.getCell('A8');
      titleCell.value = reportTitle;
      titleCell.font = { name: FONT_MOUL, size: 13, bold: true, color: { argb: 'FF000000' } };
      titleCell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
      worksheet.getRow(8).height = 30;

      // Row 9: Subtitle date range (11pt Bold)
      worksheet.mergeCells('A9:I9');
      const subtitleCell = worksheet.getCell('A9');
      subtitleCell.value = `ដោយគិតចាប់ពីថ្ងៃទី${startParts.khmerDay} ខែ${startParts.khmerMonth} ឆ្នាំ${startParts.khmerYear} រហូតដល់ថ្ងៃទី${endParts.khmerDay} ខែ${endParts.khmerMonth} ឆ្នាំ${endParts.khmerYear}`;
      subtitleCell.font = { name: FONT_SIEMREAP, size: 11, bold: true, color: { argb: 'FF000000' } };
      subtitleCell.alignment = { horizontal: 'center', vertical: 'middle' };
      worksheet.getRow(9).height = 24;

      // Row 10: Blank spacing row
      worksheet.getRow(10).height = 12;

      // --- 3. Table Headers (Row 11 & 12) ---
      // Row 11: 40 pixels (30.0 pt)
      // Row 12: 99 pixels (74.3 pt) with Wrap Text
      const headerRow1 = worksheet.getRow(11);
      const headerRow2 = worksheet.getRow(12);
      headerRow1.height = 30.0;
      headerRow2.height = 74.3;

      // Col A (1): ល.រ
      worksheet.mergeCells('A11:A12');
      worksheet.getCell('A11').value = 'ល.រ';

      // Col B (2): ប្រភេទ
      worksheet.mergeCells('B11:B12');
      worksheet.getCell('B11').value = 'ប្រភេទ';

      // Col C (3): សន្និធិចុងគ្រា
      worksheet.mergeCells('C11:C12');
      worksheet.getCell('C11').value = `សន្និធិ\nចុងគ្រា ${prevPeriodStr}${showOfficeSubLabel ? '\nការិយាល័យ' : ''}`;

      // Col D-H (4-8): Date Range
      worksheet.mergeCells('D11:H11');
      worksheet.getCell('D11').value = `${startParts.khmerDay}-${startParts.khmerMonth}-${startParts.khmerYear} ដល់ ${endParts.khmerDay}-${endParts.khmerMonth}-${endParts.khmerYear}`;

      // Col I (9): សន្និធិសល់
      worksheet.mergeCells('I11:I12');
      worksheet.getCell('I11').value = `សន្និធិ\nសល់ ${endParts.khmerDay} ${endParts.khmerMonth} ${endParts.khmerYear}${showOfficeSubLabel ? '\nការិយាល័យ' : ''}`;

      // Row 12 Subheaders (Cols D-H)
      worksheet.getCell('D12').value = 'សន្លឹកទិដ្ឋាការបើកពី\nក១';
      worksheet.getCell('E12').value = 'សន្លឹកទិដ្ឋាការ\nបង្វិលពីក្រុម';
      worksheet.getCell('F12').value = `សន្លឹកទិដ្ឋាការ\n${issuedColTitle.includes('ច្រកទ្វារ') ? 'បើកផ្តល់ទៅច្រកទ្វារ' : 'បើកផ្តល់ទៅក្រុម'}`;
      worksheet.getCell('G12').value = `សន្លឹកទិដ្ឋាការ\n${damagedColTitle.includes('ក២') ? 'មិនបានការក២' : 'មិនបានការ'}`;
      worksheet.getCell('H12').value = 'សន្លឹកទិដ្ឋាការ\nបោះពុម្ពសាកល្បង ក២';

      // Format all header cells in Rows 11 and 12 (11pt Bold, Center & Middle, Wrap Text)
      for (let r = 11; r <= 12; r++) {
        for (let c = 1; c <= 9; c++) {
          const cell = worksheet.getRow(r).getCell(c);
          cell.font = { name: FONT_SIEMREAP, size: 11, bold: true, color: { argb: 'FF000000' } };
          cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
          cell.fill = headerFill;
          cell.border = thinBorder;
        }
      }

      // --- 4. Table Data Rows (Starting Row 13 to 25) ---
      // Rows 13 to 26: 40 pixels (30.0 pt) 100% across all rows
      let currentRowIdx = 13;
      reportRows.forEach((r) => {
        const row = worksheet.getRow(currentRowIdx);
        row.height = 30.0; // 40 pixels

        // Col 1 (A): No (Center Align)
        const cell1 = row.getCell(1);
        cell1.value = r.no;
        cell1.font = { name: FONT_SIEMREAP, size: 11, color: { argb: 'FF000000' } };
        cell1.alignment = { horizontal: 'center', vertical: 'middle' };
        cell1.border = thinBorder;

        // Col 2 (B): Visa Type (Center Align)
        const cell2 = row.getCell(2);
        cell2.value = r.visaType;
        cell2.font = { name: FONT_SIEMREAP, size: 11, bold: true, color: { argb: 'FF000000' } };
        cell2.alignment = { horizontal: 'center', vertical: 'middle' };
        cell2.border = thinBorder;

        // Cols 3-9 (C-I): Numbers (Right Align, #,##0 format)
        const numValues = [
          r.beginningStock,
          r.receivedK1,
          r.returnedFromTeam,
          r.issuedToTeam,
          r.damaged,
          r.testPrintK2,
          r.endingStock,
        ];

        numValues.forEach((val, i) => {
          const colNum = i + 3;
          const cell = row.getCell(colNum);
          cell.value = val;
          cell.numFmt = '#,##0';
          cell.font = {
            name: FONT_SIEMREAP,
            size: 11,
            color: { argb: 'FF000000' },
            bold: colNum === 9, // ending stock bold
          };
          cell.alignment = { horizontal: 'right', vertical: 'middle' };
          cell.border = thinBorder;
        });

        currentRowIdx++;
      });

      // --- 5. Total Summary Row (Row 26) ---
      // 40 pixels (30.0 pt), Bold, Double Bottom Border
      const totalRow = worksheet.getRow(currentRowIdx);
      totalRow.height = 30.0; // 40 pixels

      worksheet.mergeCells(`A${currentRowIdx}:B${currentRowIdx}`);
      const totalLabelCell = worksheet.getCell(`A${currentRowIdx}`);
      totalLabelCell.value = 'សរុប';
      totalLabelCell.font = { name: FONT_MOUL, size: 11, bold: true, color: { argb: 'FF000000' } };
      totalLabelCell.alignment = { horizontal: 'center', vertical: 'middle' };
      totalLabelCell.fill = headerFill;
      totalLabelCell.border = totalRowBorder;

      worksheet.getCell(`B${currentRowIdx}`).border = totalRowBorder;
      worksheet.getCell(`B${currentRowIdx}`).fill = headerFill;

      const totalValues = [
        totals.beginningStock,
        totals.receivedK1,
        totals.returnedFromTeam,
        totals.issuedToTeam,
        totals.damaged,
        totals.testPrintK2,
        totals.endingStock,
      ];

      totalValues.forEach((val, i) => {
        const colNum = i + 3;
        const cell = totalRow.getCell(colNum);
        cell.value = val;
        cell.numFmt = '#,##0';
        cell.font = { name: FONT_SIEMREAP, size: 11, bold: true, color: { argb: 'FF000000' } };
        cell.alignment = { horizontal: 'right', vertical: 'middle' };
        cell.fill = headerFill;
        cell.border = totalRowBorder;
      });

      currentRowIdx++;

      // --- 6. Bottom Signatures Block ---
      currentRowIdx += 1; // 1 blank row spacing

      if (showLeftSignature) {
        // Left & Right Signature
        worksheet.mergeCells(`A${currentRowIdx}:D${currentRowIdx}`);
        const c1 = worksheet.getCell(`A${currentRowIdx}`);
        c1.value = 'បានឃើញ និងគោរពជូន';
        c1.font = { name: FONT_SIEMREAP, size: 11, bold: true };
        c1.alignment = { horizontal: 'center', vertical: 'middle' };

        worksheet.mergeCells(`G${currentRowIdx}:I${currentRowIdx}`);
        const c2 = worksheet.getCell(`G${currentRowIdx}`);
        c2.value = customLunarDate || `ភ្នំពេញ, ថ្ងៃទី${reportDateParts.khmerDay} ខែ${reportDateParts.khmerMonth} ឆ្នាំ${reportDateParts.khmerYear}`;
        c2.font = { name: FONT_SIEMREAP, size: 11, bold: !customLunarDate };
        c2.alignment = { horizontal: 'center', vertical: 'middle' };
        currentRowIdx++;

        worksheet.mergeCells(`A${currentRowIdx}:E${currentRowIdx}`);
        const c3 = worksheet.getCell(`A${currentRowIdx}`);
        c3.value = 'លោកនាយការិយាល័យ មេត្តាជ្រាប ដោយក្តីអនុគ្រោះ។';
        c3.font = { name: FONT_SIEMREAP, size: 11 };
        c3.alignment = { horizontal: 'center', vertical: 'middle' };

        worksheet.mergeCells(`G${currentRowIdx}:I${currentRowIdx}`);
        const c4 = worksheet.getCell(`G${currentRowIdx}`);
        c4.value = customLunarDate
          ? `ភ្នំពេញ, ថ្ងៃទី${reportDateParts.khmerDay} ខែ${reportDateParts.khmerMonth} ឆ្នាំ${reportDateParts.khmerYear}`
          : rightSignatureTitle;
        c4.font = customLunarDate
          ? { name: FONT_SIEMREAP, size: 11, bold: true }
          : { name: FONT_MOUL, size: 11, bold: true };
        c4.alignment = { horizontal: 'center', vertical: 'middle' };
        currentRowIdx++;

        if (customLunarDate) {
          worksheet.mergeCells(`G${currentRowIdx}:I${currentRowIdx}`);
          const c5 = worksheet.getCell(`G${currentRowIdx}`);
          c5.value = rightSignatureTitle;
          c5.font = { name: FONT_MOUL, size: 11, bold: true };
          c5.alignment = { horizontal: 'center', vertical: 'middle' };
          currentRowIdx++;
        }

        worksheet.mergeCells(`A${currentRowIdx}:D${currentRowIdx}`);
        const c6 = worksheet.getCell(`A${currentRowIdx}`);
        c6.value = `ភ្នំពេញ, ថ្ងៃទី${reportDateParts.khmerDay} ខែ${reportDateParts.khmerMonth} ឆ្នាំ${reportDateParts.khmerYear}`;
        c6.font = { name: FONT_SIEMREAP, size: 11 };
        c6.alignment = { horizontal: 'center', vertical: 'middle' };
        currentRowIdx++;

        worksheet.mergeCells(`A${currentRowIdx}:D${currentRowIdx}`);
        const c7 = worksheet.getCell(`A${currentRowIdx}`);
        c7.value = leftSignatureTitle || 'នាយរងការិយាល័យទទួលបន្ទុក';
        c7.font = { name: FONT_MOUL, size: 11, bold: true };
        c7.alignment = { horizontal: 'center', vertical: 'middle' };
        currentRowIdx++;
      } else {
        // Only Right Signature
        if (customLunarDate) {
          worksheet.mergeCells(`G${currentRowIdx}:I${currentRowIdx}`);
          const cLunar = worksheet.getCell(`G${currentRowIdx}`);
          cLunar.value = customLunarDate;
          cLunar.font = { name: FONT_SIEMREAP, size: 10.5 };
          cLunar.alignment = { horizontal: 'center', vertical: 'middle' };
          currentRowIdx++;
        }

        worksheet.mergeCells(`G${currentRowIdx}:I${currentRowIdx}`);
        const cSolar = worksheet.getCell(`G${currentRowIdx}`);
        cSolar.value = `ភ្នំពេញ, ថ្ងៃទី${reportDateParts.khmerDay} ខែ${reportDateParts.khmerMonth} ឆ្នាំ${reportDateParts.khmerYear}`;
        cSolar.font = { name: FONT_SIEMREAP, size: 11, bold: true };
        cSolar.alignment = { horizontal: 'center', vertical: 'middle' };
        currentRowIdx++;

        worksheet.mergeCells(`G${currentRowIdx}:I${currentRowIdx}`);
        const cTitle = worksheet.getCell(`G${currentRowIdx}`);
        cTitle.value = rightSignatureTitle;
        cTitle.font = { name: FONT_MOUL, size: 11, bold: true };
        cTitle.alignment = { horizontal: 'center', vertical: 'middle' };
        currentRowIdx++;
      }

      // Space for signature stamp (3 blank rows)
      currentRowIdx += 3;

      // Signer Name
      worksheet.mergeCells(`G${currentRowIdx}:I${currentRowIdx}`);
      const signerCell = worksheet.getCell(`G${currentRowIdx}`);
      const signerFullName = `${signerRank ? signerRank + ' ' : ''}${signerNameOnly}`;
      signerCell.value = signerFullName;
      signerCell.font = {
        name: FONT_MOUL,
        size: 11.5,
        bold: true,
        color: useRedSignerName ? { argb: 'FFDC2626' } : { argb: 'FF000000' },
      };
      signerCell.alignment = { horizontal: 'center', vertical: 'middle' };

      // Write buffer and download as xlsx
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], {
        type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      });
      const url = window.URL.createObjectURL(blob);
      const anchor = document.createElement('a');
      anchor.href = url;
      anchor.download = `របាយការណ៍ស្តុកសន្លឹកទិដ្ឋាការ_ការិយាល័យ_${startDate}_ដល់_${endDate}.xlsx`;
      anchor.click();
      window.URL.revokeObjectURL(url);
    } catch (err) {
      console.error('Export Excel error:', err);
      alert('មានបញ្ហាក្នុងការទាញយក Excel សូមព្យាយាមម្តងទៀត');
    } finally {
      setIsExportingExcel(false);
    }
  };

  const syncFromImportedData = () => {
    setTempBaselines(autoBeginningStockMap);
    setBaselines(autoBeginningStockMap);
    localStorage.setItem('sticker_office_report_baselines', JSON.stringify(autoBeginningStockMap));
  };

  // Month navigation helpers
  const currentStartYear = useMemo(() => {
    try {
      return new Date(startDate).getFullYear() || 2026;
    } catch {
      return 2026;
    }
  }, [startDate]);

  const currentStartMonth = useMemo(() => {
    try {
      return new Date(startDate).getMonth(); // 0-indexed
    } catch {
      return 3;
    }
  }, [startDate]);

  const setMonthRange = (year: number, monthIndex: number) => {
    const startD = new Date(year, monthIndex, 1);
    const endD = new Date(year, monthIndex + 1, 0); // Last day of target month

    const yStr = String(startD.getFullYear());
    const mStr = String(startD.getMonth() + 1).padStart(2, '0');
    const startIso = `${yStr}-${mStr}-01`;
    const endIso = `${yStr}-${mStr}-${String(endD.getDate()).padStart(2, '0')}`;

    setIsManualPrevPeriodStr(false);
    setStartDate(startIso);
    setEndDate(endIso);
    setReportDate(endIso);
  };

  // Jump to Previous Month
  const handlePrevMonth = () => {
    try {
      const cur = new Date(startDate);
      const prev = new Date(cur.getFullYear(), cur.getMonth() - 1, 1);
      setMonthRange(prev.getFullYear(), prev.getMonth());
    } catch {
      // fallback
    }
  };

  // Jump to Next Month (Carrying ending remaining stock to beginning stock)
  const handleNextMonth = () => {
    try {
      const cur = new Date(startDate);
      const next = new Date(cur.getFullYear(), cur.getMonth() + 1, 1);
      setMonthRange(next.getFullYear(), next.getMonth());
    } catch {
      // fallback
    }
  };

  // Quick Presets
  const setPresetPeriod = (preset: 'thisMonth' | 'lastMonth' | 'nextMonth' | 'dec2018' | 'jan2019' | 'q1' | 's1' | 'm9' | 'year') => {
    const now = new Date();
    const curYear = now.getFullYear();
    setIsManualPrevPeriodStr(false);

    if (preset === 'dec2018') {
      setStartDate('2018-12-01');
      setEndDate('2018-12-31');
      setReportDate('2018-12-31');
      setPrevPeriodStr('៣០ វិច្ឆិកា ២០១៨');
      setTempBaselines(DEC_2018_STICKER_BASELINES);
      setBaselines(DEC_2018_STICKER_BASELINES);
      localStorage.setItem('sticker_office_report_baselines', JSON.stringify(DEC_2018_STICKER_BASELINES));
    } else if (preset === 'jan2019') {
      setStartDate('2019-01-01');
      setEndDate('2019-01-31');
      setReportDate('2019-01-31');
      setPrevPeriodStr('៣១ ធ្នូ ២០១៨');
    } else if (preset === 'thisMonth') {
      setMonthRange(curYear, now.getMonth());
    } else if (preset === 'lastMonth') {
      try {
        const cur = new Date(startDate);
        setMonthRange(cur.getFullYear(), cur.getMonth() - 1);
      } catch {
        setMonthRange(curYear, now.getMonth() - 1);
      }
    } else if (preset === 'nextMonth') {
      try {
        const cur = new Date(startDate);
        setMonthRange(cur.getFullYear(), cur.getMonth() + 1);
      } catch {
        setMonthRange(curYear, now.getMonth() + 1);
      }
    } else if (preset === 'q1') {
      const y = currentStartYear || (startDate ? new Date(startDate).getFullYear() : curYear);
      const m = (currentStartMonth !== undefined ? currentStartMonth : (startDate ? new Date(startDate).getMonth() : 0)) + 1;
      const startMStr = String(m).padStart(2, '0');
      const totalMonthsOffset = (m - 1) + (3 - 1);
      const endYear = y + Math.floor(totalMonthsOffset / 12);
      const endMonth = (totalMonthsOffset % 12) + 1;
      const daysInEndMonth = new Date(endYear, endMonth, 0).getDate();
      const endMStr = String(endMonth).padStart(2, '0');
      setStartDate(`${y}-${startMStr}-01`);
      setEndDate(`${endYear}-${endMStr}-${String(daysInEndMonth).padStart(2, '0')}`);
      setReportDate(`${endYear}-${endMStr}-${String(daysInEndMonth).padStart(2, '0')}`);
    } else if (preset === 's1') {
      const y = currentStartYear || (startDate ? new Date(startDate).getFullYear() : curYear);
      const m = (currentStartMonth !== undefined ? currentStartMonth : (startDate ? new Date(startDate).getMonth() : 0)) + 1;
      const startMStr = String(m).padStart(2, '0');
      const totalMonthsOffset = (m - 1) + (6 - 1);
      const endYear = y + Math.floor(totalMonthsOffset / 12);
      const endMonth = (totalMonthsOffset % 12) + 1;
      const daysInEndMonth = new Date(endYear, endMonth, 0).getDate();
      const endMStr = String(endMonth).padStart(2, '0');
      setStartDate(`${y}-${startMStr}-01`);
      setEndDate(`${endYear}-${endMStr}-${String(daysInEndMonth).padStart(2, '0')}`);
      setReportDate(`${endYear}-${endMStr}-${String(daysInEndMonth).padStart(2, '0')}`);
    } else if (preset === 'm9') {
      const y = currentStartYear || (startDate ? new Date(startDate).getFullYear() : curYear);
      const m = (currentStartMonth !== undefined ? currentStartMonth : (startDate ? new Date(startDate).getMonth() : 0)) + 1;
      const startMStr = String(m).padStart(2, '0');
      const totalMonthsOffset = (m - 1) + (9 - 1);
      const endYear = y + Math.floor(totalMonthsOffset / 12);
      const endMonth = (totalMonthsOffset % 12) + 1;
      const daysInEndMonth = new Date(endYear, endMonth, 0).getDate();
      const endMStr = String(endMonth).padStart(2, '0');
      setStartDate(`${y}-${startMStr}-01`);
      setEndDate(`${endYear}-${endMStr}-${String(daysInEndMonth).padStart(2, '0')}`);
      setReportDate(`${endYear}-${endMStr}-${String(daysInEndMonth).padStart(2, '0')}`);
    } else if (preset === 'year') {
      const y = currentStartYear || (startDate ? new Date(startDate).getFullYear() : curYear);
      const m = (currentStartMonth !== undefined ? currentStartMonth : (startDate ? new Date(startDate).getMonth() : 0)) + 1;
      const startMStr = String(m).padStart(2, '0');
      const totalMonthsOffset = (m - 1) + (12 - 1);
      const endYear = y + Math.floor(totalMonthsOffset / 12);
      const endMonth = (totalMonthsOffset % 12) + 1;
      const daysInEndMonth = new Date(endYear, endMonth, 0).getDate();
      const endMStr = String(endMonth).padStart(2, '0');
      setStartDate(`${y}-${startMStr}-01`);
      setEndDate(`${endYear}-${endMStr}-${String(daysInEndMonth).padStart(2, '0')}`);
      setReportDate(`${endYear}-${endMStr}-${String(daysInEndMonth).padStart(2, '0')}`);
    }
  };

  // PDF Export Function matching RobokDirectorStockWorkReport
  const [isExportingPdf, setIsExportingPdf] = useState(false);

  const handleDownloadPdf = async () => {
    if (!printAreaRef.current) return;
    setIsExportingPdf(true);
    try {
      await exportElementToPdf(
        printAreaRef.current,
        `របាយការណ៍ស្តុកសន្លឹកទិដ្ឋាការ_ការិយាល័យ_${startDate}_ដល់_${endDate}.pdf`,
        { pixelRatio: 3.2, fitSinglePage: true, orientation: 'landscape' }
      );
    } catch (e) {
      console.error('html-to-image PDF export error, attempting fallback:', e);
      try {
        if (document.fonts && document.fonts.ready) {
          await document.fonts.ready;
        }
        await new Promise((r) => setTimeout(r, 200));

        const element = printAreaRef.current;
        const canvas = await html2canvas(element, {
          scale: 3.2,
          useCORS: true,
          allowTaint: true,
          logging: false,
          backgroundColor: '#ffffff',
          windowWidth: 1600,
          windowHeight: 1200,
          imageTimeout: 0,
          onclone: (clonedDoc) => {
            sanitizeDocumentForHtml2Canvas(clonedDoc, 'sticker-office-stock-report-pdf');
            const clonedEl = clonedDoc.getElementById('sticker-office-stock-report-pdf');
            if (clonedEl) {
              clonedEl.style.width = '297mm';
              clonedEl.style.minWidth = '297mm';
              clonedEl.style.maxWidth = '297mm';
              clonedEl.style.transform = 'none';
              clonedEl.style.margin = '0 auto';
              clonedEl.style.boxShadow = 'none';
              clonedEl.style.backgroundColor = '#ffffff';
              clonedEl.style.textRendering = 'geometricPrecision';
              clonedEl.style.setProperty('-webkit-font-smoothing', 'antialiased');
              clonedEl.style.setProperty('-moz-osx-font-smoothing', 'grayscale');
            }
          },
        });

        const imgData = canvas.toDataURL('image/png', 1.0);
        const pdf = new jsPDF({
          orientation: 'landscape',
          unit: 'mm',
          format: 'a4',
          compress: true,
        });

        const imgWidth = 297;
        const imgHeight = (canvas.height * imgWidth) / canvas.width;

        pdf.addImage(imgData, 'PNG', 0, 0, imgWidth, imgHeight);
        pdf.save(`របាយការណ៍ស្តុកសន្លឹកទិដ្ឋាការ_ការិយាល័យ_${startDate}_ដល់_${endDate}.pdf`);
      } catch (fallbackErr) {
        console.error('PDF Export Error:', fallbackErr);
        alert('មានបញ្ហាក្នុងការទាញយក PDF សូមព្យាយាមម្តងទៀត');
      }
    } finally {
      setIsExportingPdf(false);
    }
  };

  const handlePrint = () => {
    printA4Document('sticker-office-stock-report-pdf', {
      orientation: 'landscape',
      documentTitle: `របាយការណ៍ស្តុកសន្លឹកទិដ្ឋាការការិយាល័យ_${currentStartYear}`,
    });
  };

  return (
    <div className="space-y-4 w-full animate-fade">
      {/* AdminLTE Content Header (Breadcrumbs) - Picture 2 style */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 bg-white px-4 py-3 rounded-md border border-gray-200 shadow-xs print:hidden">
        <div>
          <h1 className="text-base font-bold text-gray-800 flex items-center gap-2">
            <span>របាយការណ៍ស្តុកសន្លឹកទិដ្ឋាការការិយាល័យ</span>
          </h1>
          <p className="text-xs text-gray-500">
            តារាងតាមប្រភេទនីមួយៗ បើកពីក១ និងផ្តល់តាមបណ្តាប៉ុស្តិ៍ច្រកទ្វារអន្តរជាតិទាំង ២៩
          </p>
        </div>
        <div className="flex flex-wrap items-center gap-2 self-start sm:self-auto">
          <button
            type="button"
            onClick={() => setShowMarginControls(!showMarginControls)}
            className={`px-3 py-1.5 rounded-md text-xs font-bold border transition flex items-center gap-1.5 cursor-pointer shadow-2xs ${
              showMarginControls
                ? 'bg-blue-600 text-white border-blue-600'
                : 'bg-blue-50 hover:bg-blue-100 text-blue-900 border-blue-300'
            }`}
            title="កំណត់ទំហំគែមក្រដាស"
          >
            <Sliders className="w-3.5 h-3.5" />
            <span>គែមក្រដាស ({customMargins.top}cm / {customMargins.bottom}cm)</span>
          </button>
          <button
            type="button"
            onClick={handleExportExcel}
            disabled={isExportingExcel}
            className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1.5 rounded-md text-xs font-bold transition flex items-center gap-1.5 cursor-pointer shadow-2xs disabled:opacity-50"
            title="ទាញយកទិន្នន័យជា Excel"
          >
            <FileSpreadsheet className="w-3.5 h-3.5" />
            <span>{isExportingExcel ? 'កំពុងទាញយក...' : 'ទាញយក Excel'}</span>
          </button>
          <button
            type="button"
            onClick={handleDownloadPdf}
            disabled={isExportingPdf}
            className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1.5 rounded-md text-xs font-bold transition flex items-center gap-1.5 cursor-pointer shadow-2xs disabled:opacity-50"
            title="ទាញយក PDF (A4)"
          >
            <Download className="w-3.5 h-3.5" />
            <span>{isExportingPdf ? 'កំពុងទាញយក...' : 'ទាញយក PDF (A4)'}</span>
          </button>
          <div className="flex items-center gap-1.5 text-xs text-gray-500 bg-gray-100 px-3 py-1.5 rounded-md font-medium">
            <Home className="w-3.5 h-3.5 text-gray-500" />
            <span>ទំព័រដើម</span>
            <span>/</span>
            <span>ការងារស្តុក</span>
            <span>/</span>
            <span className="text-[#007bff] font-bold">របាយការណ៍ស្តុកសន្លឹក</span>
          </div>
        </div>
      </div>

      {/* Filter & Control Panel */}
      <div className="bg-white p-4 rounded-md border border-gray-200 shadow-2xs space-y-3 print:hidden">

        {/* Collapsible Paper Margins Settings */}
        {showMarginControls && (
          <div className="p-3 bg-slate-900 rounded-lg text-white border border-slate-700 text-xs space-y-3 animate-fadeIn">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-slate-800 pb-2">
              <span className="font-bold text-amber-300 flex items-center gap-1.5">
                <Sliders className="w-4 h-4 text-blue-400" />
                <span>កំណត់ទំហំគែមក្រដាស (Paper Margins Setup)</span>
              </span>
              <div className="flex items-center gap-1.5 flex-wrap">
                <button
                  type="button"
                  onClick={() => {
                    handleMarginChange('top', 0.5);
                    handleMarginChange('bottom', 0.5);
                    handleMarginChange('left', 0.8);
                    handleMarginChange('right', 0.8);
                  }}
                  className={`px-2 py-0.5 rounded text-[11px] font-bold border transition cursor-pointer ${
                    customMargins.top === 0.5 && customMargins.bottom === 0.5 && customMargins.left === 0.8 && customMargins.right === 0.8
                      ? 'bg-blue-600 border-blue-500 text-white'
                      : 'bg-slate-800 border-slate-700 text-gray-300 hover:text-white'
                  }`}
                >
                  ស្តង់ដារ (លើ 0.5cm, ក្រោម 0.5cm)
                </button>
                <button
                  type="button"
                  onClick={() => {
                    handleMarginChange('top', 0.8);
                    handleMarginChange('bottom', 0.8);
                    handleMarginChange('left', 1.0);
                    handleMarginChange('right', 1.0);
                  }}
                  className={`px-2 py-0.5 rounded text-[11px] font-bold border transition cursor-pointer ${
                    customMargins.top === 0.8 && customMargins.bottom === 0.8
                      ? 'bg-blue-600 border-blue-500 text-white'
                      : 'bg-slate-800 border-slate-700 text-gray-300 hover:text-white'
                  }`}
                >
                  មធ្យម (0.8cm / 0.8cm)
                </button>
                <button
                  type="button"
                  onClick={() => {
                    handleMarginChange('top', 1.2);
                    handleMarginChange('bottom', 1.2);
                    handleMarginChange('left', 1.5);
                    handleMarginChange('right', 1.2);
                  }}
                  className={`px-2 py-0.5 rounded text-[11px] font-bold border transition cursor-pointer ${
                    customMargins.top === 1.2 && customMargins.bottom === 1.2
                      ? 'bg-blue-600 border-blue-500 text-white'
                      : 'bg-slate-800 border-slate-700 text-gray-300 hover:text-white'
                  }`}
                >
                  ទូលាយ (1.2cm / 1.2cm)
                </button>
              </div>
            </div>

            <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
              <div>
                <label className="block text-gray-300 mb-1 font-bold">គែមលើ (Top): {customMargins.top} cm</label>
                <input
                  type="range"
                  min="0.2"
                  max="3.0"
                  step="0.1"
                  value={customMargins.top}
                  onChange={(e) => handleMarginChange('top', parseFloat(e.target.value))}
                  className="w-full accent-blue-500 cursor-pointer"
                />
              </div>
              <div>
                <label className="block text-gray-300 mb-1 font-bold">គែមក្រោម (Bottom): {customMargins.bottom} cm</label>
                <input
                  type="range"
                  min="0.2"
                  max="3.0"
                  step="0.1"
                  value={customMargins.bottom}
                  onChange={(e) => handleMarginChange('bottom', parseFloat(e.target.value))}
                  className="w-full accent-blue-500 cursor-pointer"
                />
              </div>
              <div>
                <label className="block text-gray-300 mb-1 font-bold">គែមឆ្វេង (Left): {customMargins.left} cm</label>
                <input
                  type="range"
                  min="0.4"
                  max="3.0"
                  step="0.1"
                  value={customMargins.left}
                  onChange={(e) => handleMarginChange('left', parseFloat(e.target.value))}
                  className="w-full accent-amber-500 cursor-pointer"
                />
              </div>
              <div>
                <label className="block text-gray-300 mb-1 font-bold">គែមស្តាំ (Right): {customMargins.right} cm</label>
                <input
                  type="range"
                  min="0.4"
                  max="3.0"
                  step="0.1"
                  value={customMargins.right}
                  onChange={(e) => handleMarginChange('right', parseFloat(e.target.value))}
                  className="w-full accent-blue-500 cursor-pointer"
                />
              </div>
            </div>
          </div>
        )}

        {/* Monthly Transition & Carry-over Control Bar */}
        <div className="bg-gradient-to-r from-slate-900 via-blue-950 to-slate-900 p-3 rounded-lg border border-blue-800/40 text-white shadow-sm space-y-2.5">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2.5">
            <div className="flex items-center gap-2">
              <CalendarCheck className="w-4 h-4 text-[#C6A15B] shrink-0" />
              <span className="font-bold text-xs text-amber-200">ជ្រើសរើសខែរបាយការណ៍ (ផ្ទេរសន្និធិពីខែមួយទៅខែមួយ)៖</span>
            </div>

            {/* Previous / Next Month and Direct Carry-Over Button */}
            <div className="flex flex-wrap items-center gap-1.5">
              <button
                type="button"
                onClick={handlePrevMonth}
                className="px-2.5 py-1 rounded bg-white/10 hover:bg-white/20 text-white text-xs font-bold transition flex items-center gap-1 cursor-pointer border border-white/10"
                title="ត្រឡប់ទៅខែមុន"
              >
                <ChevronLeft className="w-3.5 h-3.5" />
                <span>« ខែមុន</span>
              </button>

              <div className="flex items-center gap-1 bg-white/10 px-2 py-0.5 rounded border border-white/15">
                <select
                  value={currentStartMonth}
                  onChange={(e) => setMonthRange(currentStartYear, parseInt(e.target.value, 10))}
                  className="bg-transparent text-amber-300 font-bold text-xs focus:outline-none cursor-pointer"
                >
                  {KHMER_MONTH_NAMES.map((name, idx) => (
                    <option key={idx} value={idx} className="bg-slate-900 text-white">
                      {name}
                    </option>
                  ))}
                </select>
                <select
                  value={currentStartYear}
                  onChange={(e) => setMonthRange(parseInt(e.target.value, 10), currentStartMonth)}
                  className="bg-transparent text-amber-300 font-bold text-xs focus:outline-none cursor-pointer"
                >
                  {[2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025, 2026, 2027, 2028, 2029, 2030].map((y) => (
                    <option key={y} value={y} className="bg-slate-900 text-white">
                      {y}
                    </option>
                  ))}
                </select>
              </div>

              <button
                type="button"
                onClick={handleNextMonth}
                className="px-2.5 py-1 rounded bg-white/10 hover:bg-white/20 text-white text-xs font-bold transition flex items-center gap-1 cursor-pointer border border-white/10"
                title="ទៅខែបន្ទាប់"
              >
                <span>ខែបន្ទាប់ »</span>
                <ChevronRight className="w-3.5 h-3.5" />
              </button>

              {/* Special Carry Over / Next Month Action Button */}
              <button
                type="button"
                onClick={handleNextMonth}
                className="px-3 py-1 rounded bg-[#C6A15B] hover:bg-[#b5924d] text-slate-950 text-xs font-bold transition flex items-center gap-1.5 cursor-pointer shadow-xs ml-1"
                title="ផ្ទេរសន្និធិនៅសល់បច្ចុប្បន្នធ្វើជាសន្និធិដើមគ្រាសម្រាប់ខែបន្ទាប់"
              >
                <span>ផ្ទេរសន្និធិនៅសល់ ➡️ ទៅខែបន្ទាប់</span>
              </button>
            </div>
          </div>

          {/* Quick Preset Buttons and Explanatory Indicator */}
          <div className="flex flex-wrap items-center justify-between gap-2 pt-2 border-t border-blue-800/40 text-[11px]">
            <div className="flex flex-wrap items-center gap-1">
              <span className="text-gray-300 mr-1">គំរូកាលបរិច្ឆេទ៖</span>
              <button
                type="button"
                onClick={() => setPresetPeriod('dec2018')}
                className="px-2.5 py-0.5 rounded bg-amber-400/20 text-amber-300 font-bold border border-amber-400/40 hover:bg-amber-400/30 transition cursor-pointer"
                title="កំណត់កាលបរិច្ឆេទខែធ្នូ ឆ្នាំ២០១៨ ជាមួយនឹងសន្និធិដើម ៣០ វិច្ឆិកា ២០១៨ ចំនួន ៧០០,២០០ សន្លឹក"
              >
                ធ្នូ ២០១៨ (ស្តុកចាស់ ៣០ វិច្ឆិកា: ៧០០,២០០)
              </button>
              <button
                type="button"
                onClick={() => setPresetPeriod('jan2019')}
                className="px-2.5 py-0.5 rounded bg-emerald-400/20 text-emerald-300 font-bold border border-emerald-400/40 hover:bg-emerald-400/30 transition cursor-pointer"
                title="កំណត់កាលបរិច្ឆេទខែមករា ឆ្នាំ២០១៩ (សន្និធិដើម ៤៣៨,៣៥០ ផ្ទេរពីធ្នូ ២០១៨)"
              >
                មករា ២០១៩ (ផ្ទេរពី ធ្នូ ២០១៨)
              </button>
              <button
                type="button"
                onClick={() => setPresetPeriod('thisMonth')}
                className="px-2 py-0.5 rounded bg-white/10 hover:bg-white/20 text-gray-200 font-medium transition cursor-pointer"
              >
                ខែនេះ
              </button>
              <button
                type="button"
                onClick={() => setPresetPeriod('lastMonth')}
                className="px-2 py-0.5 rounded bg-white/10 hover:bg-white/20 text-gray-200 font-medium transition cursor-pointer"
              >
                ខែមុន
              </button>
              <button
                type="button"
                onClick={() => setPresetPeriod('nextMonth')}
                className="px-2 py-0.5 rounded bg-white/10 hover:bg-white/20 text-gray-200 font-medium transition cursor-pointer"
              >
                ខែបន្ទាប់
              </button>
              <button
                type="button"
                onClick={() => setPresetPeriod('q1')}
                className="px-2 py-0.5 rounded bg-white/10 hover:bg-white/20 text-gray-200 font-medium transition cursor-pointer"
              >
                ត្រីមាសទី១
              </button>
              <button
                type="button"
                onClick={() => setPresetPeriod('s1')}
                className="px-2 py-0.5 rounded bg-white/10 hover:bg-white/20 text-gray-200 font-medium transition cursor-pointer"
              >
                ឆមាសទី១
              </button>
              <button
                type="button"
                onClick={() => setPresetPeriod('m9')}
                className="px-2 py-0.5 rounded bg-white/10 hover:bg-white/20 text-gray-200 font-medium transition cursor-pointer"
              >
                នព្វមាស
              </button>
              <button
                type="button"
                onClick={() => setPresetPeriod('year')}
                className="px-2 py-0.5 rounded bg-white/10 hover:bg-white/20 text-gray-200 font-medium transition cursor-pointer"
              >
                រយៈពេល ១២ខែ
              </button>
            </div>

            <div className="flex items-center gap-1.5 text-amber-300 font-medium">
              <Sparkles className="w-3.5 h-3.5 text-amber-400" />
              <span>
                សន្និធិនៅសល់បច្ចុប្បន្ន៖ <strong className="text-white font-mono">{formatNum(totals.endingStock)}</strong> សន្លឹក ➜ ផ្ទេរទៅសន្និធិចុងគ្រាខែបន្ទាប់
              </span>
            </div>
          </div>
        </div>

        {/* Beginning Stock Source Indicator & Mode Selector */}
        <div className="flex flex-wrap items-center justify-between gap-2 p-2.5 bg-blue-50/80 border border-blue-200 rounded-md text-xs">
          <div className="flex items-center gap-2">
            <span className="font-bold text-blue-900 flex items-center gap-1">
              <Sparkles className="w-4 h-4 text-blue-700" />
              <span>ប្រភពសន្និធិដើមគ្រា ({prevPeriodStr})៖</span>
            </span>
            {totalAutoBeginning > 0 ? (
              <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded bg-emerald-100 text-emerald-800 font-bold text-[11px] border border-emerald-300">
                <Check className="w-3 h-3 text-emerald-600" />
                <span>បានគណនា និងផ្ទេរស្វ័យប្រវត្តិចំនួន៖ <strong>{formatNum(totalAutoBeginning)}</strong> សន្លឹក</span>
              </span>
            ) : (
              <span className="text-gray-600 text-[11px]">
                (មិនទាន់រកឃើញទិន្នន័យចំណាំ «ស្តុកចាស់ ក២» ក្នុងប្រព័ន្ធទេ — កំពុងប្រើប្រាស់តម្លៃកំណត់គំរូ ៣០ វិច្ឆិកា ២០១៨)
              </span>
            )}
          </div>

          <div className="flex items-center gap-2">
            <label className="flex items-center gap-1.5 cursor-pointer font-bold text-gray-700 select-none text-[11px]">
              <input
                type="radio"
                name="beginningStockMode"
                checked={useAutoBeginningStock}
                onChange={() => setUseAutoBeginningStock(true)}
                className="text-blue-600 focus:ring-blue-500"
              />
              <span>ស្វ័យប្រវត្តិតាមប្រតិបត្តិការ / ស្តុកចាស់ (Auto Carryover)</span>
            </label>
            <label className="flex items-center gap-1.5 cursor-pointer font-bold text-gray-700 select-none text-[11px] ml-2">
              <input
                type="radio"
                name="beginningStockMode"
                checked={!useAutoBeginningStock}
                onChange={() => setUseAutoBeginningStock(false)}
                className="text-blue-600 focus:ring-blue-500"
              />
              <span>កំណត់ដោយផ្ទាល់ដៃ</span>
            </label>
          </div>
        </div>

        {/* Expandable Initial Balances Editor */}
        {isEditingBaselines && (
          <div className="p-3.5 bg-amber-50/70 border border-amber-200 rounded-md space-y-3">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-amber-200/60 pb-2">
              <div>
                <h3 className="font-bold text-amber-900 text-xs flex items-center gap-1.5">
                  <Sliders className="w-4 h-4 text-amber-700" />
                  <span>កែសម្រួលសន្និធិចុងគ្រាមុន (Beginning Stock Balances)</span>
                </h3>
                <p className="text-[11px] text-amber-800">ជ្រើសរើសគំរូ ឬវាយបញ្ចូលចំនួនសន្និធិដើមគ្រាផ្ទាល់ខ្លួន</p>
              </div>
              <div className="flex flex-wrap items-center gap-1.5">
                <button
                  type="button"
                  onClick={syncEndingStockToBaselines}
                  className="px-2.5 py-1 text-[11px] text-emerald-900 hover:text-emerald-950 border border-emerald-400 rounded bg-emerald-100 hover:bg-emerald-200 font-bold cursor-pointer transition shadow-2xs"
                  title="ចម្លងលទ្ធផលសន្និធិនៅសល់នៃរបាយការណ៍បច្ចុប្បន្នដាក់ជា Baseline"
                >
                  ចម្លងសន្និធិនៅសល់ ({formatNum(totals.endingStock)})
                </button>
                {totalAutoBeginning > 0 && (
                  <button
                    type="button"
                    onClick={syncFromImportedData}
                    className="px-2.5 py-1 text-[11px] text-blue-900 hover:text-blue-950 border border-blue-300 rounded bg-blue-100 hover:bg-blue-200 font-bold cursor-pointer transition"
                  >
                    ទាញយកពីទិន្នន័យ Auto ({formatNum(totalAutoBeginning)})
                  </button>
                )}
                <button
                  type="button"
                  onClick={loadDec2018Baselines}
                  className="px-2.5 py-1 text-[11px] text-blue-900 hover:text-blue-950 border border-blue-300 rounded bg-blue-50 hover:bg-blue-100 font-bold cursor-pointer transition"
                >
                  គំរូ ៣០ វិច្ឆិកា ២០១៨ (៧០០,២០០)
                </button>
                <button
                  type="button"
                  onClick={resetBaselinesToDefault}
                  className="px-2.5 py-1 text-[11px] text-gray-700 hover:text-gray-900 border border-gray-300 rounded bg-white font-medium cursor-pointer"
                >
                  គំរូ ២០២៦ (២៥៤,២៥០)
                </button>
                <button
                  type="button"
                  onClick={resetBaselinesToZero}
                  className="px-2 py-1 text-[11px] text-gray-500 hover:text-red-700 border border-gray-300 rounded bg-white font-medium cursor-pointer"
                >
                  កំណត់ ០
                </button>
                <button
                  type="button"
                  onClick={saveBaselines}
                  className="px-3 py-1 text-[11px] text-white bg-amber-700 hover:bg-amber-800 rounded font-bold transition cursor-pointer flex items-center gap-1 ml-1"
                >
                  <Check className="w-3.5 h-3.5" />
                  <span>រក្សាទុក</span>
                </button>
              </div>
            </div>

            <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-7 gap-2">
              {DEFAULT_VISA_TYPES.map((vt) => (
                <div key={vt} className="bg-white p-2 rounded border border-amber-200 text-center">
                  <div className="flex items-center justify-between mb-1">
                    <span className="font-bold text-amber-900 text-xs">{vt}</span>
                    {autoBeginningStockMap[vt] > 0 && (
                      <span className="text-[9px] text-emerald-700 font-mono font-medium" title="ចំនួនពីការគណនាស្វ័យប្រវត្តិ">
                        Auto:{autoBeginningStockMap[vt].toLocaleString()}
                      </span>
                    )}
                  </div>
                  <input
                    type="number"
                    value={tempBaselines[vt] ?? 0}
                    onChange={(e) =>
                      setTempBaselines((prev) => ({
                        ...prev,
                        [vt]: parseInt(e.target.value, 10) || 0,
                      }))
                    }
                    className="w-full border border-gray-300 rounded px-1.5 py-1 text-center font-mono font-bold text-xs text-gray-900 focus:border-amber-600 focus:outline-none"
                  />
                </div>
              ))}
            </div>
          </div>
        )}

        {/* Date & Title Inputs Grid */}
        <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-3 pt-1 text-xs">
          <div>
            <label className="block font-bold text-gray-700 mb-1 h-5 flex items-center whitespace-nowrap truncate">កាលបរិច្ឆេទចាប់ផ្តើម (From Date):</label>
            <CustomDatePicker
              value={startDate}
              onChange={(d) => {
                setIsManualPrevPeriodStr(false);
                setStartDate(d);
              }}
              className="h-9 font-sans text-xs"
            />
          </div>

          <div>
            <label className="block font-bold text-gray-700 mb-1 h-5 flex items-center whitespace-nowrap truncate">រហូតដល់ថ្ងៃទី (To Date):</label>
            <CustomDatePicker
              value={endDate}
              onChange={(d) => setEndDate(d)}
              className="h-9 font-sans text-xs"
            />
          </div>

          <div>
            <label className="block font-bold text-gray-700 mb-1 h-5 flex items-center whitespace-nowrap truncate">កាលបរិច្ឆេទចេញរបាយការណ៍ (Report Date):</label>
            <CustomDatePicker
              value={reportDate}
              onChange={(d) => setReportDate(d)}
              className="h-9 font-sans text-xs"
            />
          </div>

          <div>
            <label className="block font-bold text-gray-700 mb-1 h-5 flex items-center whitespace-nowrap truncate">អត្ថបទកាលបរិច្ឆេទចុងគ្រាមុន:</label>
            <input
              type="text"
              value={prevPeriodStr}
              onChange={(e) => {
                setIsManualPrevPeriodStr(true);
                setPrevPeriodStr(e.target.value);
              }}
              placeholder="ឧទាហរណ៍៖ ៣០ វិច្ឆិកា ២០១៨"
              className="w-full h-9 border border-gray-300 rounded px-2.5 font-sans text-xs focus:border-blue-500 focus:outline-none box-border leading-normal bg-white text-gray-800"
            />
          </div>
        </div>

        {/* Signer Selection Grid */}
        <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-3 pt-1 text-xs">
          <div>
            <label className="block font-bold text-gray-700 mb-1 h-5 flex items-center whitespace-nowrap truncate">តួនាទីអ្នកចុះហត្ថលេខា:</label>
            <select
              value={rightSignatureTitle}
              onChange={(e) => setRightSignatureTitle(e.target.value)}
              className="w-full h-9 border border-gray-300 rounded px-2.5 font-sans text-xs focus:border-blue-500 focus:outline-none box-border leading-normal font-bold text-gray-800 bg-white"
            >
              {RIGHT_SIGNATURE_OPTIONS.map((opt) => (
                <option key={opt} value={opt}>
                  {opt}
                </option>
              ))}
            </select>
          </div>

          <div>
            <label className="block font-bold text-gray-700 mb-1 h-5 flex items-center whitespace-nowrap truncate">ជ្រើសរើសមន្ត្រី (អ្នកធ្វើតារាង):</label>
            <select
              value={selectedOfficerId}
              onChange={(e) => setSelectedOfficerId(e.target.value)}
              className="w-full h-9 border border-gray-300 rounded px-2.5 font-sans text-xs focus:border-blue-500 focus:outline-none box-border leading-normal text-gray-800 bg-white"
            >
              <option value="">-- បញ្ចូលឈ្មោះផ្ទាល់ ឬជ្រើសរើស --</option>
              {officers?.map((o) => (
                <option key={o.id} value={o.id}>
                  {o.rankName ? `${o.rankName} ` : ''}
                  {o.name}
                </option>
              ))}
            </select>
          </div>

          <div>
            <label className="block font-bold text-gray-700 mb-1 h-5 flex items-center whitespace-nowrap truncate">ថ្នាក់/តួនាទី (Rank):</label>
            <input
              type="text"
              value={signerRank}
              onChange={(e) => setSignerRank(e.target.value)}
              placeholder="ឧទាហរណ៍៖ អនុសេនីយ៍ឯក"
              className="w-full h-9 border border-gray-300 rounded px-2.5 font-sans text-xs focus:border-blue-500 focus:outline-none box-border leading-normal text-gray-800 bg-white"
            />
          </div>

          <div>
            <label className="block font-bold text-gray-700 mb-1 h-5 flex items-center whitespace-nowrap truncate">ឈ្មោះអ្នកចុះហត្ថលេខា:</label>
            <input
              type="text"
              value={signerNameOnly}
              onChange={(e) => setSignerNameOnly(e.target.value)}
              placeholder="ឧទាហរណ៍៖ អ៊ុក រ័ត្នបញ្ញា"
              className="w-full h-9 border border-gray-300 rounded px-2.5 font-sans text-xs focus:border-blue-500 focus:outline-none box-border leading-normal text-gray-800 bg-white font-bold"
            />
          </div>
        </div>

        {/* Signature Options Controls */}
        <div className="pt-2 border-t border-gray-100 flex flex-wrap items-center justify-between gap-3 text-xs">
          <div className="flex items-center gap-3">
            <label className="flex items-center gap-2 cursor-pointer font-bold text-gray-700 bg-blue-50/70 hover:bg-blue-100/70 px-3 py-1.5 rounded border border-blue-200 transition">
              <input
                type="checkbox"
                checked={showLeftSignature}
                onChange={(e) => setShowLeftSignature(e.target.checked)}
                className="w-4 h-4 text-blue-600 rounded border-gray-300 focus:ring-blue-500 cursor-pointer"
              />
              <span>បង្ហាញហត្ថលេខាផ្នែកខាងឆ្វេង (បានឃើញ និងគោរពជូន)</span>
            </label>
          </div>

          <div className="text-[11px] text-blue-900 bg-blue-50/80 px-2.5 py-1 rounded border border-blue-200 font-medium">
            <span className="font-bold">រូបមន្តគណនាសន្និធិចុងគ្រា៖</span> (ស្តុកចាស់ ក២ + សន្លឹកទិដ្ឋាការបើកពីក១ + សន្លឹកទិដ្ឋាការបង្វិលពីក្រុម - ការបើកផ្តល់តាមក្រុម - សន្លឹកទិដ្ឋាការមិនបានការក២ - សន្លឹកទិដ្ឋាការបោះពុម្ពសាកល្បង ក២ + ផ្ទេរការប្រើប្រាស់ក្រុម)
          </div>

          {showLeftSignature && (
            <div className="flex items-center gap-1.5">
              <span className="text-gray-600 font-medium">ចំណងជើងខាងឆ្វេង:</span>
              <input
                type="text"
                value={leftSignatureTitle}
                onChange={(e) => setLeftSignatureTitle(e.target.value)}
                placeholder="នាយរងការិយាល័យទទួលបន្ទុក"
                className="border border-gray-300 rounded px-2.5 py-1 text-xs focus:border-blue-500 focus:outline-none w-48 font-bold text-gray-800 bg-white"
              />
            </div>
          )}
        </div>
      </div>

      {/* Report A4 Document Canvas Preview Container matching EVisaStockReport style */}
      <div className="bg-gray-200/80 p-2 sm:p-6 rounded-md overflow-x-auto flex flex-col items-center print:p-0 print:bg-white print:block">
        <div
          ref={printAreaRef}
          id="sticker-office-stock-report-pdf"
          className="bg-white shadow-lg border border-gray-300 rounded-xs text-black leading-relaxed text-sm w-[297mm] min-h-[210mm] max-w-full print:shadow-none print:border-none print:m-0 print:w-[297mm] print:min-h-[210mm] print:max-w-none flex flex-col justify-start font-siemreap shrink-0"
          style={{
            boxSizing: 'border-box',
            paddingTop: `${customMargins.top}cm`,
            paddingLeft: `${customMargins.left}cm`,
            paddingRight: `${customMargins.right}cm`,
            paddingBottom: `${customMargins.bottom}cm`,
          }}
        >
          <div>
            {/* Cambodian Government Official Header */}
            <div className="flex justify-between items-start mb-4 pt-1 font-siemreap">
              <div className="text-center font-moul leading-relaxed space-y-0.5 text-black pt-6" style={{ fontSize: '12pt' }}>
                <p className="font-moul" style={{ fontSize: '12pt' }}>ក្រសួងមហាផ្ទៃ</p>
                <p className="font-moul" style={{ fontSize: '12pt' }}>អគ្គនាយកដ្ឋានអន្តោប្រវេសន៍</p>
                <p className="font-moul" style={{ fontSize: '12pt' }}>នាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត</p>
                <p className="font-moul" style={{ fontSize: '12pt' }}>ការិយាល័យទិដ្ឋាការចូល</p>
                <p className="font-moul" style={{ fontSize: '12pt' }}>ផ្នែករដ្ឋបាល</p>
                <div className="flex justify-center pt-1">
                  <TacteingLine type={tacteingType} customImage={tacteingCustomImage} width={140} height={14} />
                </div>
              </div>

              <div className="text-center font-moul leading-relaxed space-y-0.5 text-black" style={{ fontSize: '12pt' }}>
                <p className="font-moul" style={{ fontSize: '12pt' }}>ព្រះរាជាណាចក្រកម្ពុជា</p>
                <p className="font-moul" style={{ fontSize: '12pt' }}>ជាតិ សាសនា ព្រះមហាក្សត្រ</p>
                <div className="flex justify-center pt-1">
                  <TacteingLine type={tacteingType} customImage={tacteingCustomImage} width={150} height={14} />
                </div>
              </div>
            </div>

            {/* Document Title Header */}
            <div className="text-center my-4 space-y-1">
              <h1 className="font-moul text-[14.5px] text-black leading-normal max-w-5xl mx-auto px-2">
                {reportTitle}
              </h1>
              <p className="font-siemreap font-bold text-[13px] text-black">
                ដោយគិតចាប់ពីថ្ងៃទី{startParts.khmerDay} ខែ{startParts.khmerMonth} ឆ្នាំ{startParts.khmerYear} រហូតដល់ថ្ងៃទី{endParts.khmerDay} ខែ{endParts.khmerMonth} ឆ្នាំ{endParts.khmerYear}
              </p>
            </div>

            {/* Main Data Table matching EVisaStockReport design */}
            <div className="mt-4 mb-4">
              <table className="w-full border-collapse border border-black text-center text-[11.5px] text-black align-middle font-siemreap">
                <thead>
                  <tr className="bg-gray-100 text-black font-bold">
                    <th rowSpan={2} style={{ verticalAlign: 'middle' }} className="border border-black p-1.5 w-[38px] text-center font-bold bg-gray-100">
                      ល.រ
                    </th>
                    <th rowSpan={2} style={{ verticalAlign: 'middle' }} className="border border-black p-1.5 w-[65px] text-center font-bold bg-gray-100">
                      ប្រភេទ
                    </th>
                    <th rowSpan={2} style={{ verticalAlign: 'middle' }} className="border border-black p-1.5 w-[120px] text-center bg-gray-100">
                      <div className="font-bold leading-normal">សន្និធិ</div>
                      <div className="font-bold leading-normal">ចុងគ្រា {prevPeriodStr}</div>
                      {showOfficeSubLabel && <div className="font-bold leading-normal">ការិយាល័យ</div>}
                    </th>
                    <th colSpan={5} style={{ verticalAlign: 'middle' }} className="border border-black p-1.5 font-bold text-center bg-gray-100">
                      {startParts.khmerDay}-{startParts.khmerMonth}-{startParts.khmerYear} ដល់ {endParts.khmerDay}-{endParts.khmerMonth}-{endParts.khmerYear}
                    </th>
                    <th rowSpan={2} style={{ verticalAlign: 'middle' }} className="border border-black p-1.5 w-[120px] text-center bg-gray-100">
                      <div className="font-bold leading-normal">សន្និធិ</div>
                      <div className="font-bold leading-normal">សល់ {endParts.khmerDay} {endParts.khmerMonth} {endParts.khmerYear}</div>
                      {showOfficeSubLabel && <div className="font-bold leading-normal">ការិយាល័យ</div>}
                    </th>
                  </tr>
                  <tr className="bg-gray-100 text-black font-bold text-[10.5px]">
                    <th style={{ verticalAlign: 'middle' }} className="border border-black p-1.5 leading-tight w-[110px] bg-gray-100">
                      សន្លឹកទិដ្ឋាការបើកពី<br />ក១
                    </th>
                    <th style={{ verticalAlign: 'middle' }} className="border border-black p-1.5 leading-tight w-[110px] bg-gray-100">
                      សន្លឹកទិដ្ឋាការ<br />បង្វិលពីក្រុម
                    </th>
                    <th style={{ verticalAlign: 'middle' }} className="border border-black p-1.5 leading-tight w-[110px] bg-gray-100">
                      {issuedColTitle.includes('ច្រកទ្វារ') ? (
                        <>
                          សន្លឹកទិដ្ឋាការ<br />បើកផ្តល់ទៅច្រកទ្វារ
                        </>
                      ) : (
                        <>
                          សន្លឹកទិដ្ឋាការ<br />បើកផ្តល់ទៅក្រុម
                        </>
                      )}
                    </th>
                    <th style={{ verticalAlign: 'middle' }} className="border border-black p-1.5 leading-tight w-[110px] bg-gray-100">
                      {damagedColTitle.includes('ក២') ? (
                        <>
                          សន្លឹកទិដ្ឋាការ<br />មិនបានការក២
                        </>
                      ) : (
                        <>
                          សន្លឹកទិដ្ឋាការ<br />មិនបានការ
                        </>
                      )}
                    </th>
                    <th style={{ verticalAlign: 'middle' }} className="border border-black p-1.5 leading-tight w-[125px] bg-gray-100">
                      សន្លឹកទិដ្ឋាការ<br />បោះពុម្ពសាកល្បង ក២
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {reportRows.map((row) => (
                    <tr key={row.no} className="text-center hover:bg-gray-50/40">
                      <td className="border border-black py-1 px-1 font-times text-center text-[10.5pt]">{row.no}</td>
                      <td className="border border-black py-1 px-1 font-bold font-times text-center text-[10.5pt]">{row.visaType}</td>
                      <td className="border border-black py-1 px-2 text-right pr-3 font-times text-[10.5pt]">
                        {formatNum(row.beginningStock)}
                      </td>
                      <td className="border border-black py-1 px-2 text-right pr-3 font-times text-[10.5pt]">
                        {formatNum(row.receivedK1)}
                      </td>
                      <td className="border border-black py-1 px-2 text-right pr-3 font-times text-[10.5pt]">
                        {formatNum(row.returnedFromTeam)}
                      </td>
                      <td className="border border-black py-1 px-2 text-right pr-3 font-times text-[10.5pt]">
                        {formatNum(row.issuedToTeam)}
                      </td>
                      <td className="border border-black py-1 px-2 text-right pr-3 font-times text-[10.5pt]">
                        {formatNum(row.damaged)}
                      </td>
                      <td className="border border-black py-1 px-2 text-right pr-3 font-times text-[10.5pt]">
                        {formatNum(row.testPrintK2)}
                      </td>
                      <td className="border border-black py-1 px-2 text-right pr-3 font-times font-bold text-[10.5pt]">
                        {formatNum(row.endingStock)}
                      </td>
                    </tr>
                  ))}

                  {/* Total Summary Row */}
                  <tr className="text-center font-bold bg-gray-100 text-[11.5px]">
                    <td className="border border-black py-1.5 px-1 font-moul" colSpan={2}>
                      សរុប
                    </td>
                    <td className="border border-black py-1.5 px-2 text-right pr-3 font-times font-bold text-[10.5pt]">
                      {formatNum(totals.beginningStock)}
                    </td>
                    <td className="border border-black py-1.5 px-2 text-right pr-3 font-times font-bold text-[10.5pt]">
                      {formatNum(totals.receivedK1)}
                    </td>
                    <td className="border border-black py-1.5 px-2 text-right pr-3 font-times font-bold text-[10.5pt]">
                      {formatNum(totals.returnedFromTeam)}
                    </td>
                    <td className="border border-black py-1.5 px-2 text-right pr-3 font-times font-bold text-[10.5pt]">
                      {formatNum(totals.issuedToTeam)}
                    </td>
                    <td className="border border-black py-1.5 px-2 text-right pr-3 font-times font-bold text-[10.5pt]">
                      {formatNum(totals.damaged)}
                    </td>
                    <td className="border border-black py-1.5 px-2 text-right pr-3 font-times font-bold text-[10.5pt]">
                      {formatNum(totals.testPrintK2)}
                    </td>
                    <td className="border border-black py-1.5 px-2 text-right pr-3 font-times font-extrabold text-black text-[10.5pt]">
                      {formatNum(totals.endingStock)}
                    </td>
                  </tr>
                </tbody>
              </table>
            </div>

            {/* Bottom Signatures Block matching PDF exact signature stamp */}
            <div className="mt-5 text-[12.5px] font-siemreap text-black">
              <div className={`grid ${showLeftSignature ? 'grid-cols-2 gap-4' : 'grid-cols-1'} text-center items-start`}>
                {showLeftSignature && (
                  <div className="space-y-1">
                    <p className="font-bold text-black">បានឃើញ និងគោរពជូន</p>
                    <p className="text-black">លោកនាយការិយាល័យ មេត្តាជ្រាប ដោយក្តីអនុគ្រោះ។</p>
                    {customLunarDate && <p className="text-black text-[11.5px] pt-0.5">{customLunarDate}</p>}
                    <p className="text-black">
                      ភ្នំពេញ, ថ្ងៃទី{reportDateParts.khmerDay} ខែ{reportDateParts.khmerMonth} ឆ្នាំ{reportDateParts.khmerYear}
                    </p>
                    <p className="font-moul text-black pt-1.5" style={{ fontSize: '12pt' }}>{leftSignatureTitle || 'នាយរងការិយាល័យទទួលបន្ទុក'}</p>
                  </div>
                )}

                <div className={`space-y-1 ${!showLeftSignature ? 'ml-auto w-[42%] pr-4 text-center' : 'w-full pr-2 text-center'}`}>
                  {customLunarDate && <p className="text-black text-[11.5px]">{customLunarDate}</p>}
                  <p className="text-black font-bold">
                    ភ្នំពេញ, ថ្ងៃទី{reportDateParts.khmerDay} ខែ{reportDateParts.khmerMonth} ឆ្នាំ{reportDateParts.khmerYear}
                  </p>
                  <p className="font-moul text-black pt-1.5" style={{ fontSize: '12pt' }}>{rightSignatureTitle}</p>

                  <div className="h-14"></div>

                  <p className="text-[13px] flex items-center justify-center gap-1.5">
                    {signerRank && (
                      <span className={`font-siemreap font-bold ${useRedSignerName ? 'text-red-600' : 'text-black'}`}>
                        {signerRank}
                      </span>
                    )}
                    <span className={`font-moul font-bold ${useRedSignerName ? 'text-red-600' : 'text-black'}`}>
                      {signerNameOnly}
                    </span>
                  </p>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};
