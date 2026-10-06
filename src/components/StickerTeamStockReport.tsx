import React, { useState, useEffect, useRef, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { CategoriesState, Officer, StockRecord, UserRole } from '../types';
import {
  Printer,
  Download,
  Calendar,
  Layers,
  Sparkles,
  Edit3,
  Check,
  FileSpreadsheet,
  RotateCcw,
  UserCheck,
  Eye,
  EyeOff,
  FileText,
  Search,
  AlertCircle,
  HelpCircle,
  CheckCircle2,
  XCircle,
  Info,
  ChevronRight,
  Filter,
  Save,
  Undo,
  Redo,
  Type,
  Sliders,
  Bold,
  Italic,
  Underline,
  Strikethrough,
  Subscript,
  Superscript,
  Palette,
  Highlighter,
  RemoveFormatting,
  AlignLeft,
  AlignCenter,
  AlignRight,
  AlignJustify,
  List,
  ListOrdered,
  Outdent,
  Indent,
  Minus,
  Maximize2,
  Layout,
  Move,
  ZoomIn,
  ZoomOut,
  X,
  ChevronDown,
} from 'lucide-react';
import html2canvas from 'html2canvas';
import { jsPDF } from 'jspdf';
import { getKhmerLunarDate, getKhmerSolarParts, toKhmerNum } from '../utils/khmerCalendar';
import { sanitizeDocumentForHtml2Canvas, exportElementToPdf, exportElementsToPdf } from '../utils/pdfExportHelper';
import { printA4Document } from '../utils/printHelper';
import { CustomDatePicker } from './CustomDatePicker';
import { TacteingLine, TacteingType, TacteingControlSelector, saveTacteingSettings, getSavedTacteingSettings } from './TacteingLink';
import {
  OFFICIAL_29_TEAMS,
  VISA_TYPES,
  normalizeDateToISO,
  normalizeVisaType,
  normalizeTeamName,
  matchTeamInList,
} from '../utils/teamNormalization';

export interface StickerTeamStockReportProps {
  stockRecords: StockRecord[];
  categories?: CategoriesState;
  officers?: Officer[];
  currentRole?: UserRole;
  userName?: string;
  onClose?: () => void;
}

export type TeamReportTabType =
  | 'opening'     // 1. ស្តុកដើមគ្រា / ចុងគ្រាមុន
  | 'issued'      // 2. បើកពីក២
  | 'used'        // 3. ប្រើប្រាស់
  | 'transferred' // 4. ផ្ទេរការប្រើប្រាស់
  | 'damaged'     // 5. ខូចគុណភាព ខ្វះ និងបង្វិល
  | 'remaining'   // 6. នៅសល់ចុងគ្រា
  | 'allInOne';   // 7. បង្ហាញគ្រប់ទំព័រទាំង ៦

export { VISA_TYPES };
export const DEFAULT_29_TEAMS = Array.from(OFFICIAL_29_TEAMS);

// Default Baseline Values (Starts clean at 0 / empty until user records are uploaded or entered)
export const DEFAULT_OPENING_MATRIX: Record<string, Record<string, number>> = {
  'អាកាស តេជោ': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'អាកាស សៀមរាប': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'អាកាស ព្រះសីហនុ': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន ប៉ោយប៉ែត': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន បាវិត': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន ចាំយាម': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន ដូង': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន អូរស្មាច់': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន ព្រំ': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន បន្ទាយចក្រី': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន ជាំ': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន ត្រពាំងស្រែ': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន ត្រពាំងក្រៀល': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន ត្រពាំងផ្លុង': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន ភ្នំដិន': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន កោះរកា': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន ព្រៃវល្លិ៍': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន អូរយ៉ាដាវ': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន ក្អមសំណ': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន រអមសំណ': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន ភ្នំដី': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'កំពង់ផែ ឧកញ៉ាម៉ុង': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'កំពង់ផែ ស្ទឹងហាវ': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'កំពង់ផែ ព្រះសីហនុ': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'កំពង់ផែ ភ្នំពេញ': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន ព្រែកចាក': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន ម៉ឺនជ័យ': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន ស្ទឹងបត់': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'កំពង់ផែ កោះកុង': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'កំពង់ផែ កំពត': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
};

const RIGHT_SIGNATURE_OPTIONS = [
  'នាយផ្នែក',
  'ជ.នាយផ្នែក',
  'អ្នកធ្វើតារាង',
  'ប្រធានក្រុមផ្តល់ទិដ្ឋាការ',
  'ជ.ប្រធានក្រុមផ្តល់ទិដ្ឋាការ',
  'ប្រធានក្រុមផ្តល់ទិដ្ឋាការស្តីទី',
];

// Helper to convert modern CSS colors (oklch/oklab) for html2canvas
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
        if (aStr !== undefined && aStr !== null) {
          alpha = parseFloat(aStr);
          if (aStr.endsWith('%')) alpha = alpha / 100;
        }
        if (isNaN(L) || isNaN(C) || isNaN(H)) return fullMatch;
        const hRad = (H * Math.PI) / 180;
        const aVal = C * Math.cos(hRad);
        const bVal = C * Math.sin(hRad);
        return oklabToRgb(L, aVal, bVal, alpha);
      } catch {
        return fullMatch;
      }
    });
  }
  return result;
}

export const StickerTeamStockReport: React.FC<StickerTeamStockReportProps> = ({
  stockRecords,
  categories,
  officers = [],
  currentRole,
  userName,
  onClose,
}) => {
  const [activeTab, setActiveTab] = useState<TeamReportTabType>('remaining');

  // Matching helper for team names (hoisted function)
  function normalizeTeamName(name: string): string {
    if (!name) return '';
    const cleaned = name.trim().toLowerCase().replace(/\s+/g, ' ');

    // 1. Specific checks for Kaam Samnor / Leuk Daek (supports all spellings and aliases)
    if (
      cleaned.includes('ក្អមសំណ') ||
      cleaned.includes('ក្អម') ||
      cleaned.includes('រអមសំណ') ||
      cleaned.includes('ភូមិសាលា') ||
      cleaned.includes('លើកដែក')
    ) {
      return 'ព្រំដែន ក្អមសំណ';
    }

    // 2. Specific checks for other known border gates, airports & ports
    if (cleaned.includes('ត្រពាំងក្រៀល') || cleaned.includes('ត្រពាំងរូង')) return 'ព្រំដែន ត្រពាំងក្រៀល';
    if (cleaned.includes('តេជោ') || cleaned.includes('តេជា') || cleaned.includes('ពោធិ៍ចិនតុង') || (cleaned.includes('ភ្នំពេញ') && cleaned.includes('អាកាស'))) return 'អាកាស តេជោ';
    if (cleaned.includes('សៀមរាប')) return 'អាកាស សៀមរាប';
    if (cleaned.includes('ព្រះសីហនុ') && cleaned.includes('អាកាស')) return 'អាកាស ព្រះសីហនុ';
    if (cleaned.includes('ប៉ោយប៉ែត')) return 'ព្រំដែន ប៉ោយប៉ែត';
    if (cleaned.includes('បាវិត')) return 'ព្រំដែន បាវិត';
    if (cleaned.includes('ចាំយាម')) return 'ព្រំដែន ចាំយាម';
    if (cleaned.includes('ដូង')) return 'ព្រំដែន ដូង';
    if (cleaned.includes('អូរស្មាច់')) return 'ព្រំដែន អូរស្មាច់';
    if (cleaned.includes('បន្ទាយចក្រី')) return 'ព្រំដែន បន្ទាយចក្រី';
    if (cleaned.includes('ជាំ') && !cleaned.includes('ចាំ')) return 'ព្រំដែន ជាំ';
    if (cleaned.includes('ត្រពាំងស្រែ')) return 'ព្រំដែន ត្រពាំងស្រែ';
    if (cleaned.includes('ត្រពាំងផ្លុង')) return 'ព្រំដែន ត្រពាំងផ្លុង';
    if (cleaned.includes('ភ្នំដិន')) return 'ព្រំដែន ភ្នំដិន';
    if (cleaned.includes('កោះរកា')) return 'ព្រំដែន កោះរកា';
    if (cleaned.includes('ព្រៃវល្លិ៍')) return 'ព្រំដែន ព្រៃវល្លិ៍';
    if (cleaned.includes('អូរយ៉ាដាវ')) return 'ព្រំដែន អូរយ៉ាដាវ';
    if (cleaned.includes('ភ្នំដី')) return 'ព្រំដែន ភ្នំដី';
    if (cleaned.includes('ឧកញ៉ាម៉ុង')) return 'កំពង់ផែ ឧកញ៉ាម៉ុង';
    if (cleaned.includes('ស្ទឹងហាវ')) return 'កំពង់ផែ ស្ទឹងហាវ';
    if (cleaned.includes('ព្រះសីហនុ') && cleaned.includes('កំពង់ផែ')) return 'កំពង់ផែ ព្រះសីហនុ';
    if (cleaned.includes('ភ្នំពេញ') && cleaned.includes('កំពង់ផែ')) return 'កំពង់ផែ ភ្នំពេញ';
    if (cleaned.includes('ព្រែកចាក')) return 'ព្រំដែន ព្រែកចាក';
    if (cleaned.includes('ព្រែកបាក់')) return 'ព្រំដែន ព្រែកបាក់';
    if (cleaned.includes('ម៉ឺនជ័យ')) return 'ព្រំដែន ម៉ឺនជ័យ';
    if (cleaned.includes('ស្ទឹងបត់')) return 'ព្រំដែន ស្ទឹងបត់';
    if (cleaned.includes('កោះកុង')) return 'កំពង់ផែ កោះកុង';
    if (cleaned.includes('កំពត')) return 'កំពង់ផែ កំពត';

    // 3. Strict check for "ព្រំដែន ព្រំ" (Phrom Border - Pailin)
    const stripped = cleaned
      .replace(/^(ច្រកទ្វារព្រំដែនអន្តរជាតិ|ច្រកទ្វារព្រំដែន|ច្រកទ្វារអន្តរជាតិ|ច្រកទ្វារ|ព្រំដែន|ច្រក|ក្រុម|ប៉ុស្តិ៍|អន្តរជាតិ|\s)+/g, '')
      .trim();

    if (
      stripped === 'ព្រំ' ||
      cleaned === 'ព្រំ' ||
      cleaned.includes('ព្រំដែន ព្រំ') ||
      cleaned.includes('ព្រំដែនព្រំ') ||
      cleaned.includes('ច្រកព្រំ') ||
      cleaned.includes('ច្រកទ្វារព្រំ') ||
      cleaned.includes('ក្រុមព្រំ') ||
      cleaned.includes('ក្រុម ព្រំ') ||
      cleaned.includes('ផ្សារព្រំ') ||
      cleaned.includes('ប៉ៃលិន')
    ) {
      return 'ព្រំដែន ព្រំ';
    }

    return name.trim();
  }

  // Month & Year Filter
  const now = new Date();
  const [selectedYear, setSelectedYear] = useState<number>(now.getFullYear());
  const [selectedMonth, setSelectedMonth] = useState<number>(now.getMonth() + 1); // 1-12

  // Date Range Defaults
  const [startDate, setStartDate] = useState<string>(() => {
    const y = now.getFullYear();
    const m = String(now.getMonth() + 1).padStart(2, '0');
    return `${y}-${m}-01`;
  });
  const [endDate, setEndDate] = useState<string>(() => {
    const y = now.getFullYear();
    const m = now.getMonth() + 1;
    const lastDay = new Date(y, m, 0).getDate();
    return `${y}-${String(m).padStart(2, '0')}-${String(lastDay).padStart(2, '0')}`;
  });
  const [reportDate, setReportDate] = useState<string>(() => {
    const y = now.getFullYear();
    const m = now.getMonth() + 1;
    const lastDay = new Date(y, m, 0).getDate();
    return `${y}-${String(m).padStart(2, '0')}-${String(lastDay).padStart(2, '0')}`;
  });

  // When Month or Year changes, sync start/end date
  const handleMonthYearChange = (newMonth: number, newYear: number) => {
    setSelectedMonth(newMonth);
    setSelectedYear(newYear);
    const mStr = String(newMonth).padStart(2, '0');
    const lastDay = new Date(newYear, newMonth, 0).getDate();
    setStartDate(`${newYear}-${mStr}-01`);
    setEndDate(`${newYear}-${mStr}-${String(lastDay).padStart(2, '0')}`);
    setReportDate(`${newYear}-${mStr}-${String(lastDay).padStart(2, '0')}`);
  };

  // Signatures State & Defaults
  const [rightSignRole, setRightSignRole] = useState<string>('នាយផ្នែក');
  const [rightSignName, setRightSignName] = useState<string>('');
  const [midSignRole, setMidSignRole] = useState<string>('នាយរងការិយាល័យ');
  const [midSignName, setMidSignName] = useState<string>('');
  const [leftSignRole, setLeftSignRole] = useState<string>('នាយការិយាល័យ');
  const [leftSignName, setLeftSignName] = useState<string>('');

  // Tacteing Settings
  const [tacteingSettings, setTacteingSettings] = useState<{ type: TacteingType; customImage: string | null }>(() =>
    getSavedTacteingSettings()
  );
  const [showTacteingPicker, setShowTacteingPicker] = useState(false);

  const handleTacteingChange = (type: TacteingType, customImg?: string | null) => {
    setTacteingSettings({ type, customImage: customImg || null });
    saveTacteingSettings(type, customImg);
  };

  // Printable ref
  const reportRef = useRef<HTMLDivElement>(null);
  const allPagesRef = useRef<HTMLDivElement>(null);
  const [isExportingPdf, setIsExportingPdf] = useState(false);

  // MS Word-like Live Document Editing States
  const [isWordEditMode, setIsWordEditMode] = useState<boolean>(true);
  const [activeFont, setActiveFont] = useState<string>('Khmer OS Siemreap');
  const [activeFontSize, setActiveFontSize] = useState<string>('9.5pt');
  const [lineSpacing, setLineSpacing] = useState<string>('1.5');
  const [showColorPicker, setShowColorPicker] = useState<boolean>(false);
  const [showHighlightPicker, setShowHighlightPicker] = useState<boolean>(false);
  const [showFindReplace, setShowFindReplace] = useState<boolean>(false);
  const [findText, setFindText] = useState<string>('');
  const [replaceText, setReplaceText] = useState<string>('');
  const [matchCount, setMatchCount] = useState<number | null>(null);
  const [docZoom, setDocZoom] = useState<number>(100);
  const [showRuler, setShowRuler] = useState<boolean>(true);
  const [activeTabStop, setActiveTabStop] = useState<number>(2.25);
  const [customMargins, setCustomMargins] = useState<{ top: number; bottom: number; left: number; right: number }>(() => {
    try {
      const saved = localStorage.getItem('sticker_team_report_margins');
      if (saved) {
        const parsed = JSON.parse(saved);
        if (parsed && typeof parsed.top === 'number') return parsed;
      }
    } catch {}
    return {
      top: 0.4,
      bottom: 0.4,
      left: 1.8,
      right: 1.0,
    };
  });
  const [showMarginControls, setShowMarginControls] = useState<boolean>(false);
  const [showControlPanel, setShowControlPanel] = useState<boolean>(() => {
    try {
      const saved = localStorage.getItem('sticker_team_report_show_control_panel');
      if (saved !== null) return saved === 'true';
    } catch {}
    return true;
  });

  // Sync custom margins to localStorage
  useEffect(() => {
    try {
      localStorage.setItem('sticker_team_report_margins', JSON.stringify(customMargins));
    } catch {}
  }, [customMargins]);

  // Microsoft Word Page Setup Dialog States (Alt+P+S+P)
  const [showPageSetupDialog, setShowPageSetupDialog] = useState<boolean>(false);
  const [pageSetupTab, setPageSetupTab] = useState<'margins' | 'paper' | 'layout'>('margins');
  const [pageOrientation, setPageOrientation] = useState<'portrait' | 'landscape'>('portrait');
  const [paperSize, setPaperSize] = useState<'A4' | 'Letter' | 'Legal' | 'A3'>('A4');
  const [tempMargins, setTempMargins] = useState<{ top: number; bottom: number; left: number; right: number }>({
    top: 0.4,
    bottom: 0.4,
    left: 1.8,
    right: 1.0,
  });
  const [gutter, setGutter] = useState<number>(0);
  const [gutterPosition, setGutterPosition] = useState<'left' | 'top'>('left');
  const [headerMargin, setHeaderMargin] = useState<number>(1.0);
  const [footerMargin, setFooterMargin] = useState<number>(1.0);

  // Microsoft Word Font & Advanced (Ctrl+D) Dialog States
  const [showFontDialog, setShowFontDialog] = useState<boolean>(false);
  const [fontDialogTab, setFontDialogTab] = useState<'font' | 'advanced'>('advanced');
  const [characterScale, setCharacterScale] = useState<string>('100%');
  const [characterSpacing, setCharacterSpacing] = useState<'Normal' | 'Expanded' | 'Condensed'>('Condensed');
  const [spacingByPt, setSpacingByPt] = useState<number>(0.8);
  const [characterPosition, setCharacterPosition] = useState<'Normal' | 'Raised' | 'Lowered'>('Normal');
  const [positionByPt, setPositionByPt] = useState<number>(0);
  const [kerningEnabled, setKerningEnabled] = useState<boolean>(true);
  const [kerningPoints, setKerningPoints] = useState<number>(1);
  const [ligatures, setLigatures] = useState<string>('Standard and Contextual');
  const [numberSpacing, setNumberSpacing] = useState<string>('Default');
  const [numberForms, setNumberForms] = useState<string>('Default');
  const [stylisticSets, setStylisticSets] = useState<string>('Default');
  const [useContextualAlternates, setUseContextualAlternates] = useState<boolean>(false);

  // Font Tab States
  const [dialogFontFamily, setDialogFontFamily] = useState<string>('Khmer OS Siemreap');
  const [dialogFontStyle, setDialogFontStyle] = useState<string>('Regular');
  const [dialogFontSize, setDialogFontSize] = useState<string>('9.5pt');
  const [dialogFontColor, setDialogFontColor] = useState<string>('#000000');
  const [dialogUnderlineStyle, setDialogUnderlineStyle] = useState<string>('none');
  const [dialogEffects, setDialogEffects] = useState<{
    strikethrough: boolean;
    doubleStrike: boolean;
    superscript: boolean;
    subscript: boolean;
    smallCaps: boolean;
    allCaps: boolean;
    hidden: boolean;
  }>({
    strikethrough: false,
    doubleStrike: false,
    superscript: false,
    subscript: false,
    smallCaps: false,
    allCaps: false,
    hidden: false,
  });

  // Signature Positioning States
  const [showSignPositionPanel, setShowSignPositionPanel] = useState<boolean>(false);
  const [signSectionMarginTop, setSignSectionMarginTop] = useState<number>(0.2); // cm
  const [signShiftLeftX, setSignShiftLeftX] = useState<number>(0);
  const [signShiftLeftY, setSignShiftLeftY] = useState<number>(0);
  const [signShiftMidX, setSignShiftMidX] = useState<number>(0);
  const [signShiftMidY, setSignShiftMidY] = useState<number>(0);
  const [signShiftRightX, setSignShiftRightX] = useState<number>(0);
  const [signShiftRightY, setSignShiftRightY] = useState<number>(0);

  // Save state
  const [lastSavedTime, setLastSavedTime] = useState<string | null>(null);
  const [showSaveSuccessToast, setShowSaveSuccessToast] = useState<boolean>(false);

  const savedSelectionRef = useRef<Range | null>(null);

  // Track selection change
  useEffect(() => {
    const handleSelectionChange = () => {
      const sel = window.getSelection();
      if (sel && sel.rangeCount > 0 && (reportRef.current || allPagesRef.current)) {
        const range = sel.getRangeAt(0);
        const container = reportRef.current || allPagesRef.current;
        if (container && container.contains(range.commonAncestorContainer)) {
          savedSelectionRef.current = range.cloneRange();
        }
      }
    };
    document.addEventListener('selectionchange', handleSelectionChange);
    return () => {
      document.removeEventListener('selectionchange', handleSelectionChange);
    };
  }, []);

  // Keyboard Shortcuts (Ctrl+B, Ctrl+I, Ctrl+U, Ctrl+D, Ctrl+S, Ctrl+F, Tab)
  const handleEditorKeyDown = (e: React.KeyboardEvent<HTMLDivElement>) => {
    if (e.key === 'Tab') {
      e.preventDefault();
      handleInsertTab();
    } else if ((e.ctrlKey || e.metaKey) && (e.key === 'd' || e.key === 'D')) {
      e.preventDefault();
      const sel = window.getSelection();
      const container = reportRef.current || allPagesRef.current;
      if (sel && sel.rangeCount > 0 && container && container.contains(sel.getRangeAt(0).commonAncestorContainer)) {
        savedSelectionRef.current = sel.getRangeAt(0).cloneRange();
      }
      setFontDialogTab('advanced');
      setShowFontDialog(true);
    } else if ((e.ctrlKey || e.metaKey) && (e.key === 's' || e.key === 'S')) {
      e.preventDefault();
      handleSaveDocument();
    } else if ((e.ctrlKey || e.metaKey) && (e.key === 'f' || e.key === 'F')) {
      e.preventDefault();
      setShowFindReplace(true);
    }
  };

  const FONT_SIZE_STEPS = [
    '8pt',
    '8.5pt',
    '9pt',
    '9.5pt',
    '10pt',
    '10.5pt',
    '11pt',
    '12pt',
    '13pt',
    '14pt',
    '16pt',
    '18pt',
    '20pt',
    '22pt',
    '24pt',
    '28pt',
    '32pt',
  ];

  const executeCommand = (command: string, value: string | undefined = undefined) => {
    document.execCommand(command, false, value);
  };

  const handleApplyFont = (font: string) => {
    setActiveFont(font);
    const sel = window.getSelection();
    if (savedSelectionRef.current && (!sel || sel.rangeCount === 0 || sel.isCollapsed)) {
      sel?.removeAllRanges();
      sel?.addRange(savedSelectionRef.current);
    }

    const currentSel = window.getSelection();
    if (currentSel && currentSel.rangeCount > 0 && !currentSel.isCollapsed) {
      const range = currentSel.getRangeAt(0);
      const span = document.createElement('span');
      if (font.includes('Mool') || font.includes('Moul')) {
        span.style.fontFamily = `'Khmer OS Muol Light', 'Khmer OS Mool1', 'Khmer OS Moul', serif`;
      } else if (font.includes('Siemreap')) {
        span.style.fontFamily = `'Khmer OS Siemreap', 'Siemreap', sans-serif`;
      } else if (font.includes('Battambang')) {
        span.style.fontFamily = `'Khmer OS Battambang', 'Battambang', sans-serif`;
      } else if (font.includes('Times')) {
        span.style.fontFamily = `'Times New Roman', Times, serif`;
      } else {
        span.style.fontFamily = `'${font}', sans-serif`;
      }

      try {
        span.appendChild(range.extractContents());
        range.insertNode(span);
        currentSel.removeAllRanges();
        const newRange = document.createRange();
        newRange.selectNodeContents(span);
        currentSel.addRange(newRange);
        savedSelectionRef.current = newRange.cloneRange();
      } catch (err) {
        executeCommand('fontName', font);
      }
    }
  };

  const handleApplyFontSize = (fontSize: string) => {
    setActiveFontSize(fontSize);
    const sel = window.getSelection();
    if (savedSelectionRef.current && (!sel || sel.rangeCount === 0 || sel.isCollapsed)) {
      sel?.removeAllRanges();
      sel?.addRange(savedSelectionRef.current);
    }

    const currentSel = window.getSelection();
    if (currentSel && currentSel.rangeCount > 0 && !currentSel.isCollapsed) {
      const range = currentSel.getRangeAt(0);
      const span = document.createElement('span');
      span.style.fontSize = fontSize;

      try {
        span.appendChild(range.extractContents());
        range.insertNode(span);
        currentSel.removeAllRanges();
        const newRange = document.createRange();
        newRange.selectNodeContents(span);
        currentSel.addRange(newRange);
        savedSelectionRef.current = newRange.cloneRange();
      } catch (err) {
        document.execCommand('fontSize', false, '3');
      }
    }
  };

  const handleGrowFontSize = () => {
    const idx = FONT_SIZE_STEPS.indexOf(activeFontSize);
    if (idx < FONT_SIZE_STEPS.length - 1) {
      handleApplyFontSize(FONT_SIZE_STEPS[idx + 1]);
    }
  };

  const handleShrinkFontSize = () => {
    const idx = FONT_SIZE_STEPS.indexOf(activeFontSize);
    if (idx > 0) {
      handleApplyFontSize(FONT_SIZE_STEPS[idx - 1]);
    }
  };

  const handleApplyTextColor = (color: string) => {
    const sel = window.getSelection();
    if (savedSelectionRef.current && (!sel || sel.rangeCount === 0 || sel.isCollapsed)) {
      sel?.removeAllRanges();
      sel?.addRange(savedSelectionRef.current);
    }
    executeCommand('foreColor', color);
    setShowColorPicker(false);
  };

  const handleApplyHighlight = (color: string) => {
    const sel = window.getSelection();
    if (savedSelectionRef.current && (!sel || sel.rangeCount === 0 || sel.isCollapsed)) {
      sel?.removeAllRanges();
      sel?.addRange(savedSelectionRef.current);
    }
    if (color === 'transparent') {
      executeCommand('removeFormat');
    } else {
      executeCommand('hiliteColor', color);
    }
    setShowHighlightPicker(false);
  };

  const handleClearFormatting = () => {
    executeCommand('removeFormat');
  };

  const handleApplyLineSpacing = (spacing: string) => {
    setLineSpacing(spacing);
    const container = reportRef.current || allPagesRef.current;
    if (container) {
      container.style.lineHeight = spacing;
    }
  };

  const handleInsertDateStamp = () => {
    const now = new Date();
    const parts = getKhmerSolarParts(now);
    const dateStr = ` ថ្ងៃទី${parts.khmerDay} ខែ${parts.khmerMonth} ឆ្នាំ${parts.khmerYear} `;
    executeCommand('insertText', dateStr);
  };

  const handleInsertHorizontalRule = () => {
    executeCommand('insertHorizontalRule');
  };

  const handleInsertTab = () => {
    const sel = window.getSelection();
    if (savedSelectionRef.current && (!sel || sel.rangeCount === 0 || sel.isCollapsed)) {
      sel?.removeAllRanges();
      sel?.addRange(savedSelectionRef.current);
    }
    const span = document.createElement('span');
    span.style.display = 'inline-block';
    span.style.width = `${activeTabStop}cm`;
    span.innerHTML = '&nbsp;';

    const currentSel = window.getSelection();
    if (currentSel && currentSel.rangeCount > 0) {
      const range = currentSel.getRangeAt(0);
      range.insertNode(span);
      range.setStartAfter(span);
      range.setEndAfter(span);
      currentSel.removeAllRanges();
      currentSel.addRange(range);
      savedSelectionRef.current = range.cloneRange();
    } else {
      executeCommand('insertHTML', `<span style="display:inline-block;width:${activeTabStop}cm">&nbsp;</span>`);
    }
  };

  const handleInsertSpace = (count: number = 1) => {
    const spaces = '&nbsp;'.repeat(count);
    executeCommand('insertHTML', spaces);
  };

  const handleFindNext = () => {
    if (!findText) return;
    if ((window as any).find) {
      const found = (window as any).find(findText, false, false, true, false, false, false);
      if (!found) {
        (window as any).find(findText, false, false, true, false, true, false);
      }
    }
  };

  const handleReplaceAll = () => {
    const container = reportRef.current || allPagesRef.current;
    if (!findText || !container) return;
    const html = container.innerHTML;
    const regex = new RegExp(findText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g');
    const matches = html.match(regex);
    setMatchCount(matches ? matches.length : 0);
    container.innerHTML = html.replace(regex, replaceText);
  };

  const handleApplyFontDialog = () => {
    const sel = window.getSelection();
    if (savedSelectionRef.current && (!sel || sel.rangeCount === 0 || sel.isCollapsed)) {
      sel?.removeAllRanges();
      sel?.addRange(savedSelectionRef.current);
    }

    const currentSel = window.getSelection();
    if (currentSel && currentSel.rangeCount > 0 && !currentSel.isCollapsed) {
      const range = currentSel.getRangeAt(0);
      const span = document.createElement('span');

      if (characterSpacing === 'Condensed') {
        span.style.letterSpacing = `-${spacingByPt}pt`;
      } else if (characterSpacing === 'Expanded') {
        span.style.letterSpacing = `${spacingByPt}pt`;
      } else {
        span.style.letterSpacing = 'normal';
      }

      if (characterScale !== '100%') {
        const scaleVal = parseFloat(characterScale.replace('%', '')) / 100;
        span.style.display = 'inline-block';
        span.style.transform = `scaleX(${scaleVal})`;
        span.style.transformOrigin = 'left center';
      }

      if (characterPosition === 'Raised' && positionByPt > 0) {
        span.style.position = 'relative';
        span.style.top = `-${positionByPt}pt`;
      } else if (characterPosition === 'Lowered' && positionByPt > 0) {
        span.style.position = 'relative';
        span.style.top = `${positionByPt}pt`;
      }

      if (dialogFontFamily) {
        if (dialogFontFamily.includes('Mool') || dialogFontFamily.includes('Moul')) {
          span.style.fontFamily = `'Khmer OS Muol Light', 'Khmer OS Mool1', 'Khmer OS Moul', serif`;
        } else if (dialogFontFamily.includes('Siemreap')) {
          span.style.fontFamily = `'Khmer OS Siemreap', 'Siemreap', sans-serif`;
        } else if (dialogFontFamily.includes('Battambang')) {
          span.style.fontFamily = `'Khmer OS Battambang', 'Battambang', sans-serif`;
        } else if (dialogFontFamily.includes('Times')) {
          span.style.fontFamily = `'Times New Roman', Times, serif`;
        } else {
          span.style.fontFamily = `'${dialogFontFamily}', sans-serif`;
        }
      }

      if (dialogFontSize) {
        span.style.fontSize = dialogFontSize;
      }
      if (dialogFontStyle.includes('Bold')) {
        span.style.fontWeight = 'bold';
      }
      if (dialogFontStyle.includes('Italic')) {
        span.style.fontStyle = 'italic';
      }
      if (dialogFontColor) {
        span.style.color = dialogFontColor;
      }
      if (dialogUnderlineStyle === 'single') {
        span.style.textDecoration = 'underline';
      } else if (dialogEffects.strikethrough) {
        span.style.textDecoration = 'line-through';
      }

      try {
        span.appendChild(range.extractContents());
        range.insertNode(span);
        currentSel.removeAllRanges();
        const newRange = document.createRange();
        newRange.selectNodeContents(span);
        currentSel.addRange(newRange);
        savedSelectionRef.current = newRange.cloneRange();
      } catch (err) {
        console.error('Failed to apply font dialog formatting:', err);
      }
    }

    setShowFontDialog(false);
  };

  const handleRulerClick = (e: React.MouseEvent<HTMLDivElement>) => {
    const rect = e.currentTarget.getBoundingClientRect();
    const clickX = e.clientX - rect.left;
    const totalWidthPx = rect.width;
    const clickedTotalCm = (clickX / totalWidthPx) * 21.0;
    const relativeCm = clickedTotalCm - customMargins.left;
    const maxContentCm = 21.0 - customMargins.left - customMargins.right;
    if (relativeCm > 0.2 && relativeCm < maxContentCm) {
      const rounded = Math.round(relativeCm * 2) / 2;
      setActiveTabStop(rounded > 0 ? rounded : 0.5);
    }
  };

  const handleSaveDocument = () => {
    try {
      const now = new Date();
      const parts = getKhmerSolarParts(now);
      const hours = String(now.getHours()).padStart(2, '0');
      const minutes = String(now.getMinutes()).padStart(2, '0');
      const seconds = String(now.getSeconds()).padStart(2, '0');
      const formattedTime = `ម៉ោង ${hours}:${minutes}:${seconds} ថ្ងៃទី${parts.khmerDay} ខែ${parts.khmerMonth} ឆ្នាំ${parts.khmerYear}`;
      setLastSavedTime(formattedTime);
      setShowSaveSuccessToast(true);
      setTimeout(() => setShowSaveSuccessToast(false), 4000);

      const draftData = {
        startDate,
        endDate,
        selectedMonth,
        selectedYear,
        customMargins,
        activeTabStop,
        activeFont,
        activeFontSize,
        lineSpacing,
        rightSignRole,
        rightSignName,
        midSignRole,
        midSignName,
        leftSignRole,
        leftSignName,
        signSectionMarginTop,
        signShiftLeftX,
        signShiftLeftY,
        signShiftMidX,
        signShiftMidY,
        signShiftRightX,
        signShiftRightY,
        savedAt: now.toISOString(),
        savedAtFormatted: formattedTime,
      };
      localStorage.setItem(`sticker_team_report_draft_${startDate}_${endDate}`, JSON.stringify(draftData));
    } catch (err) {
      console.error('Error saving document:', err);
    }
  };

  // Manual Overrides / Baseline persistence in localStorage (starts clean at 0 unless user explicitly edited)
  const [manualOpeningOverrides, setManualOpeningOverrides] = useState<Record<string, Record<string, number>>>(() => {
    const saved = localStorage.getItem('sticker_team_opening_overrides_v2');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        // Clear if containing old mock sample data (e.g. 5405 for Phnom Penh port or 172510 for Techo)
        if (
          parsed['កំពង់ផែ ភ្នំពេញ']?.T === 5405 ||
          parsed['អាកាស តេជោ']?.T === 172510 ||
          parsed['អាកាស តេជោ']?.T === 183209 ||
          parsed['អាកាស សៀមរាប']?.T === 324000
        ) {
          localStorage.removeItem('sticker_team_opening_overrides_v2');
          return DEFAULT_OPENING_MATRIX;
        }
        return parsed;
      } catch {}
    }
    return DEFAULT_OPENING_MATRIX;
  });

  useEffect(() => {
    try {
      const saved = localStorage.getItem('sticker_team_opening_overrides_v2');
      if (saved) {
        const parsed = JSON.parse(saved);
        if (
          parsed['កំពង់ផែ ភ្នំពេញ']?.T === 5405 ||
          parsed['អាកាស តេជោ']?.T === 172510 ||
          parsed['អាកាស តេជោ']?.T === 183209 ||
          parsed['អាកាស សៀមរាប']?.T === 324000
        ) {
          localStorage.removeItem('sticker_team_opening_overrides_v2');
          setManualOpeningOverrides(DEFAULT_OPENING_MATRIX);
        }
      }
    } catch {}
  }, []);

  const [isEditingOpening, setIsEditingOpening] = useState(false);

  // Teams list (directly use categories.visaTeamsRobok / categories.visaTeams from Office system "ទិន្នន័យក្រុមផ្តល់ទិដ្ឋាការ", fallback to DEFAULT_29_TEAMS)
  const teamsList = useMemo(() => {
    const robokTeams = (categories?.visaTeamsRobok || [])
      .map((t) => t.name.trim())
      .filter(Boolean);
    if (robokTeams.length > 0) {
      return Array.from(new Set(robokTeams));
    }
    const fullTeams = (categories?.visaTeams || [])
      .map((t) => normalizeTeamName(t.name.trim()))
      .filter(Boolean);
    if (fullTeams.length > 0) {
      return Array.from(new Set(fullTeams));
    }
    return DEFAULT_29_TEAMS;
  }, [categories]);

  // Mapping of team short names to full official names from Office system categories
  const teamFullNameMap = useMemo(() => {
    const map: Record<string, string> = {};
    const fullItems = categories?.visaTeams || [];
    const robokItems = categories?.visaTeamsRobok || [];
    const maxLen = Math.max(fullItems.length, robokItems.length);

    for (let i = 0; i < maxLen; i++) {
      const full = fullItems[i]?.name?.trim();
      const robok = robokItems[i]?.name?.trim();
      if (robok && full) {
        map[robok] = full;
        map[normalizeTeamName(robok)] = full;
      }
    }
    return map;
  }, [categories?.visaTeams, categories?.visaTeamsRobok]);

  // Audit Modal State
  const [isAuditModalOpen, setIsAuditModalOpen] = useState<boolean>(false);
  const [selectedAuditTeam, setSelectedAuditTeam] = useState<string>('ព្រំដែន ព្រំ');
  const [auditSearchTerm, setAuditSearchTerm] = useState<string>('');
  const [auditTab, setAuditTab] = useState<'counted' | 'uncounted' | 'all'>('counted');

  // Helper to identify if a record belongs to cEA (electronic approval / e-Visa)
  const isCeaRecord = (rec: StockRecord): boolean => {
    if (rec.stockType === 'evisa') return true;
    const sf = (rec.sourceFrom || '').trim().toLowerCase();
    const id = (rec.id || '').toLowerCase();
    const vt = (rec.visaType || '').trim().toLowerCase();
    const op = (rec.operationType || '').toLowerCase();
    return (
      sf === 'cea' ||
      sf.includes('cea') ||
      id.includes('-cea-') ||
      vt.includes('cea') ||
      op.includes('cea')
    );
  };

  // Helper to identify oldStockTeam (ស្តុកចាស់ក្រុម) records (strictly Sticker only)
  const isOldStockTeamRecord = (rec: StockRecord): boolean => {
    if (rec.stockType && rec.stockType !== 'sticker') return false;
    if (isCeaRecord(rec)) return false;
    const op = (rec.operationType || '').trim().toLowerCase();
    const sf = (rec.sourceFrom || '').trim();

    if (
      op === 'oldstockteam' ||
      op === 'old_stock_team' ||
      op.includes('oldstockteam') ||
      op.includes('ស្តុកចាស់ក្រុម')
    ) {
      return true;
    }

    if (
      sf === 'ស្តុកចាស់ក្រុម' ||
      sf === 'ស្តុកចាស់របស់ក្រុម' ||
      sf.includes('ស្តុកចាស់ក្រុម') ||
      sf.includes('ស្តុកចាស់របស់ក្រុម') ||
      sf.includes('សន្និធិដើមក្រុម') ||
      sf.includes('ស្តុកចាស់តាមក្រុម')
    ) {
      return true;
    }

    const hasTeam = Boolean(
      rec.visaTeamRobokName ||
      rec.visaTeamRobokId ||
      (rec as any).team ||
      (rec as any).teamName ||
      sf.includes('ក្រុម') ||
      sf.includes('ប៉ុស្តិ៍') ||
      sf.includes('ព្រំដែន') ||
      sf.includes('អាកាស') ||
      sf.includes('កំពង់ផែ')
    );

    if (sf.includes('ស្តុកចាស់') && hasTeam && !sf.includes('ក២') && !sf.includes('ក១')) {
      return true;
    }

    return false;
  };

  // Helper to resolve team name from any field in StockRecord
  const resolveRecordTeamName = (r: StockRecord): string => {
    let raw = (
      r.visaTeamRobokName ||
      (r as any).teamName ||
      (r as any).team ||
      (r as any).visaTeam ||
      (r as any).station ||
      (r as any).port ||
      ''
    ).trim();

    if (!raw && r.visaTeamRobokId) {
      const found = (categories?.visaTeamsRobok || categories?.visaTeams || []).find(
        (t) => t.id === r.visaTeamRobokId || t.name === r.visaTeamRobokId
      );
      if (found) raw = found.name;
      else raw = r.visaTeamRobokId;
    }

    if (!raw && r.sourceFrom) {
      const cleaned = r.sourceFrom
        .replace(/ស្តុកចាស់របស់ក្រុម|ស្តុកចាស់ក្រុម|សន្និធិដើមក្រុម|ស្តុកចាស់|សន្និធិដើម|[\-:_]/g, '')
        .trim();
      if (cleaned) {
        raw = cleaned;
      }
    }

    return raw.trim();
  };

  // Extract latest date from oldStockTeam records if available (e.g. "2018-11-30")
  const oldStockTeamRawDate = useMemo(() => {
    const oldStockRecs = (stockRecords || []).filter(isOldStockTeamRecord);
    if (oldStockRecs.length > 0) {
      const dates = oldStockRecs.map((r) => r.date).filter(Boolean);
      if (dates.length > 0) {
        dates.sort();
        return dates[dates.length - 1];
      }
    }
    return null;
  }, [stockRecords]);

  const oldStockTeamSolarDate = useMemo(() => {
    if (oldStockTeamRawDate) {
      return getKhmerSolarParts(oldStockTeamRawDate);
    }
    return null;
  }, [oldStockTeamRawDate]);

  // Count how many oldStockTeam records were detected
  const detectedOldStockCount = useMemo(() => {
    return (stockRecords || []).filter(isOldStockTeamRecord).length;
  }, [stockRecords]);

  // Mode for Opening Stock data source
  const [openingSourceMode, setOpeningSourceMode] = useState<'auto' | 'recordedOnly' | 'baseline2018'>('auto');

  // Build Matrices from stockRecords & Overrides (High-performance Memoized)
  const {
    openingMatrix,
    issuedMatrix,
    usedMatrix,
    transferredMatrix,
    damagedMatrix,
    remainingMatrix,
  } = useMemo(() => {
    // 1. Initial Empty Matrices
    const openMat: Record<string, Record<string, number>> = {};
    const issMat: Record<string, Record<string, number>> = {};
    const useMat: Record<string, Record<string, number>> = {};
    const transMat: Record<string, Record<string, number>> = {};
    const damMat: Record<string, Record<string, number>> = {};
    const remMat: Record<string, Record<string, number>> = {};

    // 1.1 Map data recorded under 'oldStockTeam' (ស្តុកចាស់ក្រុម) in stockRecords
    const recordedOldStockMap: Record<string, Record<string, number>> = {};
    const teamsWithRecordedOldStock = new Set<string>();

    (stockRecords || []).forEach((rec) => {
      if (!isOldStockTeamRecord(rec)) return;

      const rawTeam = resolveRecordTeamName(rec);
      const targetTeam = matchTeamInList(rawTeam, teamsList);
      const normTeam = normalizeTeamName(rawTeam);
      if (!targetTeam && !normTeam && !rawTeam) return;

      const vt = normalizeVisaType(rec.visaType);
      if (!VISA_TYPES.includes(vt as any)) return;

      const qty = parseInt(String(rec.quantityBundles ?? rec.totalSheets ?? (rec as any).quantity ?? 0), 10) || 0;

      const matchedTeam = targetTeam || normTeam || rawTeam;
      teamsWithRecordedOldStock.add(matchedTeam);
      if (normTeam) teamsWithRecordedOldStock.add(normTeam);
      if (rawTeam) teamsWithRecordedOldStock.add(rawTeam);

      // Store by target team
      if (targetTeam) {
        if (!recordedOldStockMap[targetTeam]) recordedOldStockMap[targetTeam] = {};
        recordedOldStockMap[targetTeam][vt] = (recordedOldStockMap[targetTeam][vt] || 0) + qty;
      }

      // Store by normalized team
      if (normTeam && normTeam !== targetTeam) {
        if (!recordedOldStockMap[normTeam]) recordedOldStockMap[normTeam] = {};
        recordedOldStockMap[normTeam][vt] = (recordedOldStockMap[normTeam][vt] || 0) + qty;
      }

      // Store by raw team
      if (rawTeam && rawTeam !== targetTeam && rawTeam !== normTeam) {
        if (!recordedOldStockMap[rawTeam]) recordedOldStockMap[rawTeam] = {};
        recordedOldStockMap[rawTeam][vt] = (recordedOldStockMap[rawTeam][vt] || 0) + qty;
      }
    });

    const hasAnyRecordedOldStock = Object.keys(recordedOldStockMap).length > 0;
    const shouldUseRecorded = openingSourceMode === 'recordedOnly' || (openingSourceMode === 'auto' && hasAnyRecordedOldStock);

    teamsList.forEach((team) => {
      const normalizedTeam = normalizeTeamName(team);
      openMat[team] = {};
      issMat[team] = {};
      useMat[team] = {};
      transMat[team] = {};
      damMat[team] = {};
      remMat[team] = {};

      const hasTeamRecord =
        teamsWithRecordedOldStock.has(team) ||
        teamsWithRecordedOldStock.has(normalizedTeam) ||
        Boolean(Object.keys(recordedOldStockMap).find((k) => normalizeTeamName(k) === normalizedTeam));

      VISA_TYPES.forEach((vt) => {
        let initialVal = 0;

        if (recordedOldStockMap[team]?.[vt] !== undefined) {
          initialVal = recordedOldStockMap[team][vt];
        } else if (recordedOldStockMap[normalizedTeam]?.[vt] !== undefined) {
          initialVal = recordedOldStockMap[normalizedTeam][vt];
        } else {
          const matchKey = Object.keys(recordedOldStockMap).find((k) => normalizeTeamName(k) === normalizedTeam);
          if (matchKey && recordedOldStockMap[matchKey]?.[vt] !== undefined) {
            initialVal = recordedOldStockMap[matchKey][vt];
          } else if (hasTeamRecord) {
            initialVal = 0;
          } else if (manualOpeningOverrides[team]?.[vt] !== undefined && !shouldUseRecorded) {
            initialVal = manualOpeningOverrides[team][vt];
          } else if (manualOpeningOverrides[normalizedTeam]?.[vt] !== undefined && !shouldUseRecorded) {
            initialVal = manualOpeningOverrides[normalizedTeam][vt];
          } else if (!shouldUseRecorded || openingSourceMode === 'baseline2018') {
            initialVal = DEFAULT_OPENING_MATRIX[team]?.[vt] ?? DEFAULT_OPENING_MATRIX[normalizedTeam]?.[vt] ?? 0;
          } else {
            initialVal = 0;
          }
        }

        openMat[team][vt] = initialVal;
        issMat[team][vt] = 0;
        useMat[team][vt] = 0;
        transMat[team][vt] = 0;
        damMat[team][vt] = 0;
        remMat[team][vt] = initialVal;
      });
    });

    // 2. Iterate stockRecords for Sticker Visas (Strictly exclude cEA / electronic approvals)
    const normStartDate = normalizeDateToISO(startDate);
    const normEndDate = normalizeDateToISO(endDate);

    (stockRecords || []).forEach((rec) => {
      if (rec.stockType && rec.stockType !== 'sticker') return;
      if (isCeaRecord(rec)) return;
      if (isOldStockTeamRecord(rec)) return;

      const rawTeam = resolveRecordTeamName(rec);
      const targetTeam = matchTeamInList(rawTeam, teamsList);
      const normTeam = normalizeTeamName(rawTeam);
      const effectiveTeam = teamsList.find((t) => t === targetTeam || t === rawTeam || normalizeTeamName(t) === normTeam) || targetTeam || normTeam;
      if (!openMat[effectiveTeam]) return;

      const vt = normalizeVisaType(rec.visaType);
      if (!VISA_TYPES.includes(vt as any)) return;

      const qty = Number(rec.quantityBundles || rec.totalSheets || (rec as any).quantity || 0);
      const recDate = normalizeDateToISO(rec.date || '');

      // Check if record date is before startDate (adds/subtracts to Opening)
      const isAfterNov30_2018 = !recDate || recDate > '2018-11-30';
      const opLower = (rec.operationType || '').trim().toLowerCase();
      if (
        opLower === 'returnstub' ||
        opLower.includes('stub') ||
        opLower === 'openk1' ||
        opLower === 'oldstockk2' ||
        opLower === 'returnk1' ||
        opLower === 'testprintk2' ||
        opLower === 'damaged' ||
        (rec.sourceFrom || '').includes('គល់សន្លឹក')
      ) {
        return;
      }

      if (recDate && normStartDate && recDate < normStartDate && isAfterNov30_2018) {
        if (rec.operationType === 'oldStockTeam' || rec.operationType === 'oldStockK2') {
          // already counted
        } else if (rec.operationType === 'issueTeam') {
          openMat[effectiveTeam][vt] = (openMat[effectiveTeam][vt] || 0) + qty;
        } else if (rec.operationType === 'useTeam') {
          openMat[effectiveTeam][vt] = (openMat[effectiveTeam][vt] || 0) - qty;
        } else if (rec.operationType === 'transferTeam') {
          openMat[effectiveTeam][vt] = (openMat[effectiveTeam][vt] || 0) - qty;
          if (rec.recipientTeamName) {
            const rawRecip = rec.recipientTeamName;
            const targetRecip = matchTeamInList(rawRecip, teamsList);
            const normRecip = normalizeTeamName(rawRecip);
            const effectiveRecip = teamsList.find((t) => t === targetRecip || t === rawRecip || normalizeTeamName(t) === normRecip) || targetRecip || normRecip;
            if (effectiveRecip && openMat[effectiveRecip]) {
              openMat[effectiveRecip][vt] = (openMat[effectiveRecip][vt] || 0) + qty;
            }
          }
        } else if (
          rec.operationType === 'damagedTeam' ||
          rec.operationType === 'returnTeam' ||
          rec.operationType === 'missingTeam'
        ) {
          openMat[effectiveTeam][vt] = (openMat[effectiveTeam][vt] || 0) - qty;
        }
      }

      // Check if record date is within [startDate, endDate]
      const isInRange = (!normStartDate || recDate >= normStartDate) && (!normEndDate || recDate <= normEndDate);
      if (isInRange) {
        if (rec.operationType === 'issueTeam') {
          issMat[effectiveTeam][vt] = (issMat[effectiveTeam][vt] || 0) + qty;
        } else if (rec.operationType === 'useTeam') {
          useMat[effectiveTeam][vt] = (useMat[effectiveTeam][vt] || 0) + qty;
        } else if (rec.operationType === 'transferTeam') {
          transMat[effectiveTeam][vt] = (transMat[effectiveTeam][vt] || 0) + qty;
          if (rec.recipientTeamName) {
            const rawRecip = rec.recipientTeamName;
            const targetRecip = matchTeamInList(rawRecip, teamsList);
            const normRecip = normalizeTeamName(rawRecip);
            const effectiveRecip = teamsList.find((t) => t === targetRecip || t === rawRecip || normalizeTeamName(t) === normRecip) || targetRecip || normRecip;
            if (effectiveRecip && issMat[effectiveRecip]) {
              issMat[effectiveRecip][vt] = (issMat[effectiveRecip][vt] || 0) + qty;
            }
          }
        } else if (
          rec.operationType === 'damagedTeam' ||
          rec.operationType === 'returnTeam' ||
          rec.operationType === 'missingTeam'
        ) {
          damMat[effectiveTeam][vt] = (damMat[effectiveTeam][vt] || 0) + qty;
        }
      }
    });

    // 2.1 Include daily team operations strictly for Sticker with O(1) index
    try {
      const savedDaily = localStorage.getItem('app_daily_team_operations_v5');
      if (savedDaily) {
        const parsedDaily = JSON.parse(savedDaily);
        if (Array.isArray(parsedDaily)) {
          // Pre-index existing useTeam records for O(1) lookup
          const existingUseKeySet = new Set<string>();
          (stockRecords || []).forEach((sr) => {
            if (sr.operationType === 'useTeam' && !isCeaRecord(sr)) {
              const dIso = normalizeDateToISO(sr.date || '');
              const tNorm = normalizeTeamName(sr.visaTeamRobokName || sr.sourceFrom || '');
              const vNorm = normalizeVisaType(sr.visaType);
              if (dIso && tNorm && vNorm) {
                existingUseKeySet.add(`${dIso}_${tNorm}_${vNorm}`);
              }
            }
          });

          parsedDaily.forEach((dRec: any) => {
            const rawTeamName = dRec.teamName || '';
            const normTeam = normalizeTeamName(rawTeamName);

            const isCeaTeam =
              normTeam === 'អាកាស តេជោ' ||
              normTeam === 'អាកាស សៀមរាប' ||
              normTeam === 'អាកាស ព្រះសីហនុ' ||
              normTeam === 'កំពង់ផែ ព្រះសីហនុ' ||
              /អាកាស/i.test(normTeam) ||
              /តេជោ|សៀមរាប|ព្រះសីហនុ/i.test(normTeam);

            const isExplicitSticker = dRec.categoryType === 'Sticker' || dRec.selectedOption === 'Sticker';
            const isCeaRec = dRec.categoryType === 'cEA' || dRec.selectedOption === 'cEA' || dRec.sourceFrom === 'cEA' || (isCeaTeam && !isExplicitSticker);

            if (isCeaRec) return;
            const targetTeam = teamsList.find((t) => t === rawTeamName || normalizeTeamName(t) === normTeam) || normTeam;
            if (!targetTeam || !useMat[targetTeam]) return;
            const recDate = normalizeDateToISO(dRec.date || '');

            const isBeforeStart = Boolean(recDate && normStartDate && recDate < normStartDate);
            const isInRange = (!normStartDate || recDate >= normStartDate) && (!normEndDate || recDate <= normEndDate);
            if ((isBeforeStart || isInRange) && dRec.values) {
              VISA_TYPES.forEach((vt) => {
                const item = dRec.values[vt];
                const qty = parseInt(item?.quantity || '', 10) || 0;
                if (qty <= 0) return;

                const lookupKey = `${recDate}_${normTeam}_${vt}`;
                if (!existingUseKeySet.has(lookupKey)) {
                  if (isBeforeStart) {
                    openMat[targetTeam][vt] = (openMat[targetTeam][vt] || 0) - qty;
                  } else if (isInRange) {
                    useMat[targetTeam][vt] = (useMat[targetTeam][vt] || 0) + qty;
                  }
                }
              });
            }
          });
        }
      }
    } catch {}

    // 3. Compute Remaining Matrix
    teamsList.forEach((team) => {
      VISA_TYPES.forEach((vt) => {
        const op = openMat[team][vt] || 0;
        const iss = issMat[team][vt] || 0;
        const us = useMat[team][vt] || 0;
        const tr = transMat[team][vt] || 0;
        const dm = damMat[team][vt] || 0;
        remMat[team][vt] = op + iss - us - tr - dm;
      });
    });

    return {
      openingMatrix: openMat,
      issuedMatrix: issMat,
      usedMatrix: useMat,
      transferredMatrix: transMat,
      damagedMatrix: damMat,
      remainingMatrix: remMat,
    };
  }, [stockRecords, teamsList, startDate, endDate, manualOpeningOverrides]);

  // Active matrix based on active tab
  const currentMatrix = useMemo(() => {
    switch (activeTab) {
      case 'opening':
        return openingMatrix;
      case 'issued':
        return issuedMatrix;
      case 'used':
        return usedMatrix;
      case 'transferred':
        return transferredMatrix;
      case 'damaged':
        return damagedMatrix;
      case 'remaining':
      default:
        return remainingMatrix;
    }
  }, [activeTab, openingMatrix, issuedMatrix, usedMatrix, transferredMatrix, damagedMatrix, remainingMatrix]);

  // Compute column totals and row totals helper
  const getMatrixTotals = (matrix: Record<string, Record<string, number>>) => {
    const rTotals: Record<string, number> = {};
    const cTotals: Record<string, number> = {};
    VISA_TYPES.forEach((vt) => {
      cTotals[vt] = 0;
    });

    let gTotal = 0;

    teamsList.forEach((team) => {
      let rSum = 0;
      VISA_TYPES.forEach((vt) => {
        const val = matrix[team]?.[vt] || 0;
        rSum += val;
        cTotals[vt] = (cTotals[vt] || 0) + val;
      });
      rTotals[team] = rSum;
      gTotal += rSum;
    });

    return { rowTotals: rTotals, colTotals: cTotals, grandTotal: gTotal };
  };

  const { rowTotals, colTotals, grandTotal } = useMemo(
    () => getMatrixTotals(currentMatrix),
    [currentMatrix, teamsList]
  );

  // Handle manual edit change
  const handleCellEdit = (team: string, vt: string, val: number) => {
    setManualOpeningOverrides((prev) => {
      const copy = { ...prev };
      if (!copy[team]) copy[team] = {};
      copy[team] = { ...copy[team], [vt]: val };
      localStorage.setItem('sticker_team_opening_overrides_v2', JSON.stringify(copy));
      return copy;
    });
  };

  const handleResetOpening = () => {
    if (window.confirm('តើលោកអ្នកពិតជាចង់កំណត់ទិន្នន័យដើមគ្រាទៅទិន្នន័យផ្លូវការ (៣០ វិច្ឆិកា ២០១៨) វិញមែនទេ?')) {
      setManualOpeningOverrides(DEFAULT_OPENING_MATRIX);
      localStorage.setItem('sticker_team_opening_overrides_v2', JSON.stringify(DEFAULT_OPENING_MATRIX));
    }
  };

  // Khmer Dates formatting
  const khmerSolar = useMemo(() => getKhmerSolarParts(endDate || startDate), [endDate, startDate]);
  const startSolar = useMemo(() => getKhmerSolarParts(startDate), [startDate]);
  const khmerLunar = useMemo(() => getKhmerLunarDate(endDate || startDate), [endDate, startDate]);

  // Report Sign Dates formatting (from reportDate)
  const reportSolar = useMemo(() => getKhmerSolarParts(reportDate || endDate || startDate), [reportDate, endDate, startDate]);
  const reportLunar = useMemo(() => getKhmerLunarDate(reportDate || endDate || startDate), [reportDate, endDate, startDate]);

  // Previous day of startDate (or end of previous month) for Opening page title
  const prevDaySolar = useMemo(() => {
    if (!startDate) return { khmerDay: '៣០', khmerMonth: 'វិច្ឆិកា', khmerYear: '២០១៨' };
    try {
      const d = new Date(startDate);
      d.setDate(d.getDate() - 1);
      const prevIso = d.toISOString().split('T')[0];
      return getKhmerSolarParts(prevIso);
    } catch {
      return khmerSolar;
    }
  }, [startDate, khmerSolar]);

  // Titles mapping according to official PDF format
  const getTabTitle = (tab: TeamReportTabType) => {
    switch (tab) {
      case 'opening': {
        const is2018Baseline = !startDate || startDate <= '2018-12-01';
        const openingSolar = is2018Baseline
          ? { khmerDay: '៣០', khmerMonth: 'វិច្ឆិកា', khmerYear: '២០១៨' }
          : prevDaySolar;
        return {
          main: 'តារាងទិដ្ឋាការស្អិត តាមប្រភេទនីមួយៗ តាមបណ្តាប៉ុស្តិ៍ច្រកទ្វារអន្តរជាតិទាំង ២៩',
          sub: `ចុងគ្រា ថ្ងៃទី${openingSolar.khmerDay} ខែ${openingSolar.khmerMonth} ឆ្នាំ${openingSolar.khmerYear}`,
          sub2: null,
          short: '១. ស្តុកដើមគ្រា / ចុងគ្រាមុន',
        };
      }
      case 'issued':
        return {
          main: 'តារាងទិដ្ឋាការស្អិត តាមប្រភេទនីមួយៗ បើកពីក២ តាមបណ្តាប៉ុស្តិ៍ច្រកទ្វារអន្តរជាតិទាំង ២៩',
          sub: `គិតចាប់ពីថ្ងៃទី${startSolar.khmerDay} ខែ${startSolar.khmerMonth} ឆ្នាំ${startSolar.khmerYear} ដល់ថ្ងៃទី${khmerSolar.khmerDay} ខែ${khmerSolar.khmerMonth} ឆ្នាំ${khmerSolar.khmerYear}`,
          sub2: null,
          short: '២. បើកពីក២ (ចេញទៅក្រុម)',
        };
      case 'used':
        return {
          main: 'តារាងសន្លឹកទិដ្ឋាការស្អិត តាមប្រភេទនីមួយៗ ប្រើប្រាស់តាមបណ្តាប៉ុស្តិ៍ច្រកទ្វារអន្តរជាតិទាំង ២៩',
          sub: `គិតចាប់ពីថ្ងៃទី${startSolar.khmerDay} ខែ${startSolar.khmerMonth} ឆ្នាំ${startSolar.khmerYear} ដល់ថ្ងៃទី${khmerSolar.khmerDay} ខែ${khmerSolar.khmerMonth} ឆ្នាំ${khmerSolar.khmerYear}`,
          sub2: null,
          short: '៣. ការប្រើប្រាស់តាមក្រុម',
        };
      case 'transferred':
        return {
          main: 'តារាងសន្លឹកទិដ្ឋាការស្អិត តាមប្រភេទនីមួយៗ ផ្ទេរការប្រើប្រាស់តាមបណ្តាប៉ុស្តិ៍ច្រកទ្វារអន្តរជាតិទាំង ២៩',
          sub: `គិតចាប់ពីថ្ងៃទី${startSolar.khmerDay} ខែ${startSolar.khmerMonth} ឆ្នាំ${startSolar.khmerYear} ដល់ថ្ងៃទី${khmerSolar.khmerDay} ខែ${khmerSolar.khmerMonth} ឆ្នាំ${khmerSolar.khmerYear}`,
          sub2: null,
          short: '៤. ផ្ទេរការប្រើប្រាស់',
        };
      case 'damaged':
        return {
          main: 'តារាងសន្លឹកទិដ្ឋាការស្អិត តាមប្រភេទនីមួយៗ ខូចគុណភាព ខ្វះ និងបង្វិល តាមបណ្តាប៉ុស្តិ៍ច្រកទ្វារអន្តរជាតិទាំង២៩',
          sub: `គិតចាប់ពីថ្ងៃទី${startSolar.khmerDay} ខែ${startSolar.khmerMonth} ឆ្នាំ${startSolar.khmerYear} ដល់ថ្ងៃទី${khmerSolar.khmerDay} ខែ${khmerSolar.khmerMonth} ឆ្នាំ${khmerSolar.khmerYear}`,
          sub2: null,
          short: '៥. ខូចគុណភាព ខ្វះ និងបង្វិល',
        };
      case 'remaining':
      default:
        return {
          main: 'តារាងទិដ្ឋាការស្អិត តាមប្រភេទនីមួយៗ នៅសល់តាមបណ្តាប៉ុស្តិ៍ច្រកទ្វារអន្តរជាតិទាំង ២៩',
          sub: `គិតចាប់ពីថ្ងៃទី${startSolar.khmerDay} ខែ${startSolar.khmerMonth} ឆ្នាំ${startSolar.khmerYear} ដល់ថ្ងៃទី${khmerSolar.khmerDay} ខែ${khmerSolar.khmerMonth} ឆ្នាំ${khmerSolar.khmerYear}`,
          sub2: `សល់ត្រឹមថ្ងៃទី${khmerSolar.khmerDay} ខែ${khmerSolar.khmerMonth} ឆ្នាំ${khmerSolar.khmerYear}`,
          short: '៦. នៅសល់ចុងគ្រា (ស្តុកជាក់ស្តែង)',
        };
    }
  };

  const activeTitle = getTabTitle(activeTab);

  // Print
  const handlePrint = () => {
    const target = activeTab === 'allInOne' ? allPagesRef.current : reportRef.current;
    if (target) {
      printA4Document(target, {
        orientation: 'portrait',
        scale: 0.95,
        documentTitle: activeTab === 'allInOne' ? 'របាយការណ៍សរុបការងារស្តុកទិដ្ឋាការស្អិត' : activeTitle.main,
      });
    } else {
      window.print();
    }
  };

  // Export to Excel (All 6 Pages)
  const handleExportExcel = () => {
    const wb = XLSX.utils.book_new();

    const tabSheets = [
      { tab: 'opening' as const, name: '១.ដើមគ្រា', mat: openingMatrix },
      { tab: 'issued' as const, name: '២.បើកពីក២', mat: issuedMatrix },
      { tab: 'used' as const, name: '៣.ប្រើប្រាស់', mat: usedMatrix },
      { tab: 'transferred' as const, name: '៤.ផ្ទេរ', mat: transferredMatrix },
      { tab: 'damaged' as const, name: '៥.ខូច_បង្វិល', mat: damagedMatrix },
      { tab: 'remaining' as const, name: '៦.នៅសល់ចុងគ្រា', mat: remainingMatrix },
    ];

    tabSheets.forEach(({ tab, name, mat }) => {
      const headerRow = ['ល.រ', 'ក្រុមផ្តល់ទិដ្ឋាការ', ...VISA_TYPES, 'សរុបសន្លឹក'];
      const dataRows = teamsList.map((team, idx) => {
        const rowVals = VISA_TYPES.map((vt) => mat[team]?.[vt] || 0);
        const rowSum = rowVals.reduce((a, b) => a + b, 0);
        return [idx + 1, team, ...rowVals, rowSum];
      });

      // Total Row
      const colSums = VISA_TYPES.map((vt) => teamsList.reduce((sum, t) => sum + (mat[t]?.[vt] || 0), 0));
      const grandSum = colSums.reduce((a, b) => a + b, 0);
      const totalRow = ['', 'សរុបសន្លឹក', ...colSums, grandSum];

      const tabInfo = getTabTitle(tab);
      const headerTitles = [[tabInfo.main], [tabInfo.sub]];
      if (tabInfo.sub2) headerTitles.push([tabInfo.sub2]);

      const ws = XLSX.utils.aoa_to_sheet([
        ...headerTitles,
        [],
        headerRow,
        ...dataRows,
        totalRow,
      ]);

      XLSX.utils.book_append_sheet(wb, ws, name);
    });

    XLSX.writeFile(wb, `របាយការណ៍សន្លឹកទិដ្ឋាការស្អិតតាមក្រុម_${startDate}_ដល់_${endDate}.xlsx`);
  };

  // Export to PDF (Single Page or All 6 Pages)
  const handleExportPdf = async () => {
    setIsExportingPdf(true);
    try {
      if (activeTab === 'allInOne') {
        // Multi-page export from allPagesRef
        const pageElements = Array.from(document.querySelectorAll('.printable-report-page')) as HTMLElement[];
        if (pageElements.length === 0) return;

        await exportElementsToPdf(
          pageElements,
          `របាយការណ៍សន្លឹកទិដ្ឋាការស្អិត_គ្រប់៦ទំព័រ_${startDate}_ដល់_${endDate}.pdf`,
          { pixelRatio: 3 }
        );
      } else {
        // Single Page Export
        const element = (reportRef.current?.querySelector('.printable-report-page') as HTMLElement) || reportRef.current;
        if (!element) return;

        const tabTitle = getTabTitle(activeTab).main.replace(/\s+/g, '_');
        await exportElementToPdf(
          element,
          `របាយការណ៍តាមក្រុម_${tabTitle}_${startDate}_ដល់_${endDate}.pdf`,
          { pixelRatio: 3, fitSinglePage: true }
        );
      }
    } catch (err) {
      console.error('PDF export error:', err);
    } finally {
      setIsExportingPdf(false);
    }
  };

  // Reusable Single Page Table Renderer (Matching exact 100% PDF styling)
  const renderSinglePage = (
    tabType: TeamReportTabType,
    matrixData: Record<string, Record<string, number>>,
    showFullSignatures: boolean = false
  ) => {
    const tabInfo = getTabTitle(tabType);
    const { rowTotals: rTots, colTotals: cTots, grandTotal: gTot } = getMatrixTotals(matrixData);

    return (
      <div
        id="sticker-team-stock-report-pdf"
        contentEditable={isWordEditMode}
        suppressContentEditableWarning={true}
        className="printable-report-page bg-white shadow-xl border border-gray-300 rounded-xs text-black leading-relaxed text-sm w-[210mm] min-w-[210mm] max-w-[210mm] min-h-[297mm] print:shadow-none print:border-none print:m-0 print:w-[210mm] print:min-w-[210mm] print:max-w-[210mm] print:min-h-[297mm] flex flex-col justify-start font-siemreap shrink-0 box-border my-3 mx-auto page-break focus:outline-none"
        style={{
          boxSizing: 'border-box',
          width: '210mm',
          minWidth: '210mm',
          maxWidth: '210mm',
          minHeight: '297mm',
          paddingTop: `${customMargins.top}cm`,
          paddingLeft: `${customMargins.left}cm`,
          paddingRight: `${customMargins.right}cm`,
          paddingBottom: `${customMargins.bottom}cm`,
          backgroundColor: '#ffffff',
          lineHeight: lineSpacing,
        }}
      >
        <div>
          {/* Cambodian Government Official Header */}
          <div className="flex justify-between items-start mb-1.5 pt-0.5 font-siemreap">
            {/* Left Top: Ministry & Department (Aligned so ក្រសួងមហាផ្ទៃ starts level with ជាតិ សាសនា ព្រះមហាក្សត្រ) */}
            <div className="text-center space-y-0.5 inline-block pt-[1.35rem]">
              <p className="font-moul text-[11px] text-black whitespace-nowrap">ក្រសួងមហាផ្ទៃ</p>
              <p className="font-moul text-[11px] text-black whitespace-nowrap">អគ្គនាយកដ្ឋានអន្តោប្រវេសន៍</p>
              <p className="font-moul text-[10px] text-black whitespace-nowrap">នាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត</p>
              <p className="font-moul text-[10px] text-black whitespace-nowrap">ការិយាល័យទិដ្ឋាការចូល</p>
              <p className="font-moul text-[10px] text-black whitespace-nowrap">ផ្នែករដ្ឋបាល</p>
              <div className="flex justify-center pt-0.5">
                <TacteingLine
                  type={tacteingSettings.type}
                  customImage={tacteingSettings.customImage}
                  width={130}
                  height={11}
                />
              </div>
            </div>

            {/* Right Top: Kingdom of Cambodia */}
            <div className="text-center space-y-0.5 inline-block">
              <p className="font-moul text-[12px] text-black tracking-wide whitespace-nowrap">ព្រះរាជាណាចក្រកម្ពុជា</p>
              <p className="font-moul text-[11px] text-black tracking-wide whitespace-nowrap">ជាតិ សាសនា ព្រះមហាក្សត្រ</p>
              <div className="mt-1 flex justify-center">
                <TacteingLine
                  type={tacteingSettings.type}
                  customImage={tacteingSettings.customImage}
                  width={130}
                  height={11}
                />
              </div>
            </div>
          </div>

          {/* Report Title */}
          <div className="text-center my-1 space-y-0.5">
            <h2
              className="font-moul text-[9pt] text-black leading-snug"
              style={{ fontSize: '9pt' }}
            >
              {tabInfo.main}
            </h2>
            <p className="font-bold text-[10.5px] text-gray-950 font-siemreap">
              {tabInfo.sub}
            </p>
            {tabInfo.sub2 && (
              <p className="font-bold text-[10.5px] text-gray-950 font-siemreap">
                {tabInfo.sub2}
              </p>
            )}
          </div>

          {/* The 29-Team Matrix Table (Clean White Table Style) */}
          <div className="w-full my-0.5">
            <table className="w-full table-fixed border-collapse border border-black text-center text-[9.5px] leading-tight font-siemreap">
              <colgroup>
                {/* No (ល.រ) */}
                <col style={{ width: '3.5%' }} />
                {/* Team Name (ក្រុមផ្តល់ទិដ្ឋាការ) - exactly 16% as requested */}
                <col style={{ width: '16%' }} />
                {/* 13 Visa Types: T, T1, T2, T3, E, E1, E2, E3, D, K, A, B, C (5.7% each, sum 74.1%) */}
                {VISA_TYPES.map((vt) => (
                  <col key={vt} style={{ width: '5.7%' }} />
                ))}
                {/* Total (សន្លឹក) - 6.4% */}
                <col style={{ width: '6.4%' }} />
              </colgroup>

              <thead>
                {/* Header Row 1: Unified Orange #E69138 */}
                <tr className="bg-[#E69138] text-black font-bold border-b border-black">
                  <th
                    rowSpan={2}
                    className="border border-black py-1 px-0 text-center bg-[#E69138] font-bold text-[10px]"
                  >
                    ល.រ
                  </th>
                  <th
                    rowSpan={2}
                    className="border border-black py-1 px-1 text-center bg-[#E69138] font-bold truncate text-[10px]"
                  >
                    ក្រុមផ្តល់ទិដ្ឋាការ
                  </th>
                  <th
                    colSpan={13}
                    className="border border-black py-1 text-center bg-[#E69138] font-bold tracking-wide text-[10.5px]"
                  >
                    ប្រភេទទិដ្ឋាការ
                  </th>
                  <th
                    rowSpan={2}
                    className="border border-black py-1 px-0.5 text-center bg-[#E69138] font-bold text-[10px]"
                  >
                    សន្លឹក
                  </th>
                </tr>

                {/* Header Row 2: Same Unified Orange #E69138 */}
                <tr className="bg-[#E69138] text-black font-bold border-b border-black text-[10px]">
                  {VISA_TYPES.map((vt) => (
                    <th
                      key={vt}
                      className="border border-black py-1 px-0 text-center font-bold bg-[#E69138] font-times text-[10.5px]"
                    >
                      {vt}
                    </th>
                  ))}
                </tr>
              </thead>

              <tbody>
                {teamsList.map((team, idx) => {
                  const rTotal = rTots[team] || 0;
                  return (
                    <tr
                      key={team}
                      className="hover:bg-amber-50/40 transition border-b border-black text-black leading-tight h-[22px]"
                    >
                      {/* No */}
                      <td className="border border-black py-0.5 px-0 text-center font-medium font-times text-[10px] leading-none">
                        {idx + 1}
                      </td>

                      {/* Team Name */}
                      <td
                        onClick={() => {
                          setSelectedAuditTeam(team);
                          setIsAuditModalOpen(true);
                        }}
                        className="border border-black py-0.5 px-1 text-left font-medium text-black text-[9.5px] leading-tight cursor-pointer hover:bg-amber-200/70 hover:text-blue-900 group"
                        title={teamFullNameMap[team] || teamFullNameMap[normalizeTeamName(team)] || `ចុចដើម្បីពិនិត្យផ្ទៀងផ្ទាត់ទិន្នន័យ (បូក / មិនបូក) របស់ក្រុម «${team}»`}
                      >
                        <span className="inline-flex items-center gap-0.5">
                          <span>{team}</span>
                          <Eye className="w-2.5 h-2.5 opacity-0 group-hover:opacity-100 text-blue-700 transition shrink-0" />
                        </span>
                      </td>

                      {/* Visa Types T, T1, T2, T3, E, E1, E2, E3, D, K, A, B, C */}
                      {VISA_TYPES.map((vt) => {
                        const val = matrixData[team]?.[vt] || 0;

                        if (tabType === 'opening' && isEditingOpening && activeTab === 'opening') {
                          return (
                            <td
                              key={vt}
                              className="border border-black p-0 text-center bg-amber-50"
                            >
                              <input
                                type="number"
                                min="0"
                                value={manualOpeningOverrides[team]?.[vt] ?? ''}
                                onChange={(e) =>
                                  handleCellEdit(team, vt, Math.max(0, parseInt(e.target.value) || 0))
                                }
                                className="w-full text-center py-0.5 px-0 text-[10px] font-times border-none focus:outline-none focus:bg-amber-100 font-bold text-blue-900"
                              />
                            </td>
                          );
                        }

                        return (
                          <td
                            key={vt}
                            className="border border-black py-0.5 px-0 text-center font-times font-normal text-black text-[10px] leading-none"
                          >
                            {val === 0 ? '0' : val.toLocaleString()}
                          </td>
                        );
                      })}

                      {/* Row Total */}
                      <td className="border border-black py-0.5 px-0.5 text-center font-times font-normal text-black bg-gray-50/70 text-[10px] leading-none">
                        {rTotal.toLocaleString()}
                      </td>
                    </tr>
                  );
                })}

                {/* Bottom Grand Total Row: Same Unified Orange #E69138 */}
                <tr className="bg-[#E69138] font-bold text-black border-t-2 border-black h-[24px]">
                  <td
                    colSpan={2}
                    className="border border-black py-1 px-1 text-center font-bold text-[10.5px] bg-[#E69138]"
                  >
                    សរុបសន្លឹក
                  </td>
                  {VISA_TYPES.map((vt) => (
                    <td
                      key={vt}
                      className="border border-black py-1 px-0 text-center font-times font-bold bg-[#E69138] text-[10.5px]"
                    >
                      {(cTots[vt] || 0).toLocaleString()}
                    </td>
                  ))}
                  <td className="border border-black py-1 px-0.5 text-center font-times font-bold text-[10.5px] bg-[#E69138]">
                    {gTot.toLocaleString()}
                  </td>
                </tr>
              </tbody>
            </table>
          </div>
        </div>

        {/* Official Signatures Block - ONLY rendered on Page 6 (Shifted Up) */}
        {showFullSignatures && (
          <div
            className="pt-0.5 text-xs font-siemreap text-black"
            style={{ marginTop: `${signSectionMarginTop}cm` }}
          >
            <div className="grid grid-cols-3 gap-2 text-center items-start">
              {/* Left: នាយការិយាល័យ */}
              <div
                className="space-y-0.5 flex flex-col items-center"
                style={{
                  transform: `translate(${signShiftLeftX}px, ${signShiftLeftY}px)`,
                  transition: 'transform 0.1s ease-out',
                }}
              >
                <p className="font-bold text-[11px]">បានឃើញ និងឯកភាព</p>
                <p className="text-[9px] text-gray-800 whitespace-nowrap">
                  ថ្ងៃ............................ខែ.........ឆ្នាំ........ .......ស័ក ព.ស.២៥៧...
                </p>
                <p className="text-[9px] text-gray-800 whitespace-nowrap">
                  រាជធានីភ្នំពេញ, ថ្ងៃទី........ ខែ............ ឆ្នាំ២០២....
                </p>
                <div className="pt-0.5">
                  <p
                    className="font-moul text-[11px] text-black text-center tracking-normal"
                    style={{ fontFamily: "'Khmer OS Mool1', 'Moul', serif" }}
                  >
                    {leftSignRole || 'នាយការិយាល័យ'}
                  </p>
                </div>
              </div>

              {/* Middle: នាយរងការិយាល័យ */}
              <div
                className="space-y-0.5 flex flex-col items-center"
                style={{
                  transform: `translate(${signShiftMidX}px, ${signShiftMidY}px)`,
                  transition: 'transform 0.1s ease-out',
                }}
              >
                <p className="font-bold text-[11px]">បានឃើញ និងគោរពជូន</p>
                <p className="text-[9.5px] text-gray-800">
                  {reportLunar || `ថ្ងៃ... ទី... ខែ... ឆ្នាំ...`}
                </p>
                <p className="text-[9.5px] text-gray-800">
                  រាជធានីភ្នំពេញ, ថ្ងៃទី{reportSolar.khmerDay} ខែ{reportSolar.khmerMonth} ឆ្នាំ{reportSolar.khmerYear}
                </p>
                <div className="pt-0.5">
                  <p
                    className="font-moul text-[11px] text-black text-center tracking-normal"
                    style={{ fontFamily: "'Khmer OS Mool1', 'Moul', serif" }}
                  >
                    {midSignRole || 'នាយរងការិយាល័យ'}
                  </p>
                </div>
              </div>

              {/* Right: នាយផ្នែក / ជ.នាយផ្នែក / អ្នកធ្វើតារាង */}
              <div
                className="space-y-0.5 flex flex-col items-center"
                style={{
                  transform: `translate(${signShiftRightX}px, ${signShiftRightY}px)`,
                  transition: 'transform 0.1s ease-out',
                }}
              >
                <p className="text-[9.5px] text-gray-800 leading-tight">
                  {reportLunar || `ថ្ងៃ... ទី... ខែ... ឆ្នាំ...`}
                </p>
                <p className="text-[9.5px] text-gray-800">
                  ភ្នំពេញ, ថ្ងៃទី{reportSolar.khmerDay} ខែ{reportSolar.khmerMonth} ឆ្នាំ{reportSolar.khmerYear}
                </p>

                {/* Role Selector with pure Khmer OS Mool1 font like នាយការិយាល័យ */}
                <div className="relative pt-0.5 inline-block group cursor-pointer">
                  <p
                    className="font-moul text-[11px] text-black text-center tracking-normal hover:opacity-80 transition select-none"
                    style={{ fontFamily: "'Khmer OS Mool1', 'Moul', serif" }}
                  >
                    {rightSignRole || 'នាយផ្នែក'}
                  </p>
                  <select
                    value={rightSignRole}
                    onChange={(e) => setRightSignRole(e.target.value)}
                    className="absolute inset-0 opacity-0 cursor-pointer w-full h-full font-moul"
                    style={{ fontFamily: "'Khmer OS Mool1', 'Moul', serif" }}
                    title="ចុចដើម្បីប្តូរតួនាទី (នាយផ្នែក, ជ.នាយផ្នែក, អ្នកធ្វើតារាង...)"
                  >
                    {RIGHT_SIGNATURE_OPTIONS.map((opt) => (
                      <option
                        key={opt}
                        value={opt}
                        className="font-moul text-[11px]"
                        style={{ fontFamily: "'Khmer OS Mool1', 'Moul', serif" }}
                      >
                        {opt}
                      </option>
                    ))}
                  </select>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>
    );
  };

  return (
    <div className="space-y-4">
      {/* Top Action Bar / Controls */}
      <div className="bg-white rounded-lg border border-[#C6A15B]/30 shadow-xs p-3.5 print:hidden">
        <div className="flex flex-nowrap items-center justify-between gap-2.5 overflow-x-auto pb-1 whitespace-nowrap">
          {/* Left: Month/Year selector & Range */}
          <div className="flex items-center gap-2 shrink-0">
            <div className="flex items-center gap-1.5 bg-[#0B2545]/5 px-2.5 py-1 rounded-md border border-[#0B2545]/15 shrink-0">
              <Calendar className="w-4 h-4 text-[#0B2545] shrink-0" />
              <span className="text-xs font-bold text-[#0B2545]">ជ្រើសរើសខែ/ឆ្នាំ៖</span>
              <select
                value={selectedMonth}
                onChange={(e) => handleMonthYearChange(Number(e.target.value), selectedYear)}
                className="bg-white border border-gray-300 rounded px-2 py-0.5 text-xs font-bold text-gray-800 cursor-pointer"
              >
                {Array.from({ length: 12 }, (_, i) => i + 1).map((m) => (
                  <option key={m} value={m}>
                    ខែ {String(m).padStart(2, '0')}
                  </option>
                ))}
              </select>
              <select
                value={selectedYear}
                onChange={(e) => handleMonthYearChange(selectedMonth, Number(e.target.value))}
                className="bg-white border border-gray-300 rounded px-2 py-0.5 text-xs font-bold text-gray-800 cursor-pointer"
              >
                {[2018, 2019, 2020, 2021, 2022, 2023, 2024, 2025, 2026, 2027, 2028, 2029, 2030].map((y) => (
                  <option key={y} value={y}>
                    ឆ្នាំ {y}
                  </option>
                ))}
              </select>
            </div>

            <div className="flex items-center gap-1.5 text-xs text-gray-600 bg-gray-50 px-2 py-1 rounded border border-gray-200 shrink-0">
              <span>ពី:</span>
              <div className="w-28">
                <CustomDatePicker
                  value={startDate}
                  onChange={(d) => setStartDate(d)}
                  className="py-0.5 text-xs"
                />
              </div>
              <span>ដល់:</span>
              <div className="w-28">
                <CustomDatePicker
                  value={endDate}
                  onChange={(d) => setEndDate(d)}
                  className="py-0.5 text-xs"
                />
              </div>
            </div>

            <div className="flex items-center gap-1.5 text-xs text-gray-700 bg-amber-50/70 px-2 py-1 rounded border border-amber-200 shrink-0">
              <span className="font-bold text-amber-900">កាលបរិច្ឆេទចេញរបាយការណ៍:</span>
              <div className="w-32">
                <CustomDatePicker
                  value={reportDate}
                  onChange={(d) => setReportDate(d)}
                  className="py-0.5 text-xs"
                />
              </div>
            </div>
          </div>

          {/* Right: Print, Export, Tacteing */}
          <div className="flex items-center gap-1.5 sm:gap-2 shrink-0">
            {/* Tacteing Ornament Selector */}
            <div className="relative shrink-0">
              <button
                type="button"
                onClick={() => setShowTacteingPicker(!showTacteingPicker)}
                className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded text-xs font-bold bg-amber-50 hover:bg-amber-100 text-amber-900 border border-amber-300 transition cursor-pointer shrink-0 whitespace-nowrap"
              >
                <Sparkles className="w-3.5 h-3.5 text-amber-600 shrink-0" />
                <span>ម៉ូតតាក់តែង</span>
              </button>

              {showTacteingPicker && (
                <div className="absolute right-0 mt-2 z-50 w-80 bg-white shadow-xl rounded-lg border border-amber-300 p-2">
                  <TacteingControlSelector
                    currentType={tacteingSettings.type}
                    customImage={tacteingSettings.customImage}
                    onChange={handleTacteingChange}
                  />
                  <div className="text-right mt-1">
                    <button
                      type="button"
                      onClick={() => setShowTacteingPicker(false)}
                      className="text-[11px] text-gray-600 hover:text-gray-950 font-bold px-2 py-0.5 cursor-pointer"
                    >
                      បិទ
                    </button>
                  </div>
                </div>
              )}
            </div>

            {/* Team Audit / Inspection Button */}
            <button
              type="button"
              onClick={() => setIsAuditModalOpen(true)}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded bg-indigo-700 hover:bg-indigo-800 text-white text-xs font-bold transition shadow-2xs cursor-pointer shrink-0 whitespace-nowrap"
              title="ពិនិត្យផ្ទៀងផ្ទាត់ទិន្នន័យ (បូក / មិនបូក) តាមក្រុមច្រកទ្វារអន្តរជាតិទាំង ២៩"
            >
              <Search className="w-3.5 h-3.5 shrink-0" />
              <span>ផ្ទៀងផ្ទាត់ទិន្នន័យតាមក្រុម</span>
            </button>

            {/* In-place edit for opening baseline */}
            {activeTab === 'opening' && (
              <button
                type="button"
                onClick={() => setIsEditingOpening(!isEditingOpening)}
                className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded text-xs font-bold transition cursor-pointer shrink-0 whitespace-nowrap ${
                  isEditingOpening
                    ? 'bg-emerald-600 hover:bg-emerald-700 text-white shadow-2xs'
                    : 'bg-amber-50 hover:bg-amber-100 text-amber-900 border border-amber-300'
                }`}
              >
                {isEditingOpening ? (
                  <>
                    <Check className="w-3.5 h-3.5 shrink-0" />
                    <span>រក្សាទុកការកែប្រែ</span>
                  </>
                ) : (
                  <>
                    <Edit3 className="w-3.5 h-3.5 shrink-0" />
                    <span>កែសម្រួលតួលេខដើមគ្រា</span>
                  </>
                )}
              </button>
            )}

            {activeTab === 'opening' && isEditingOpening && (
              <button
                type="button"
                onClick={handleResetOpening}
                className="inline-flex items-center gap-1 px-2.5 py-1.5 rounded text-xs font-medium bg-gray-100 hover:bg-gray-200 text-gray-700 border border-gray-300 transition cursor-pointer shrink-0 whitespace-nowrap"
                title="កំណត់ទិន្នន័យដើមគ្រាទៅដើមវិញ"
              >
                <RotateCcw className="w-3 h-3 shrink-0" />
                <span>កំណត់ឡើងវិញ</span>
              </button>
            )}

            <button
              type="button"
              onClick={handleExportExcel}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded bg-emerald-700 hover:bg-emerald-800 text-white text-xs font-bold transition shadow-2xs cursor-pointer shrink-0 whitespace-nowrap"
              title="ទាញយកជា Excel គ្រប់ ៦ ទំព័រ"
            >
              <FileSpreadsheet className="w-3.5 h-3.5 shrink-0" />
              <span>Excel (៦ សន្លឹក)</span>
            </button>

            <button
              type="button"
              onClick={handleExportPdf}
              disabled={isExportingPdf}
              className="inline-flex items-center gap-1.5 px-3 py-1.5 rounded bg-red-700 hover:bg-red-800 text-white text-xs font-bold transition shadow-2xs cursor-pointer disabled:opacity-50 shrink-0 whitespace-nowrap"
            >
              <Download className="w-3.5 h-3.5 shrink-0" />
              <span>{isExportingPdf ? 'កំពុងបង្កើត PDF...' : activeTab === 'allInOne' ? 'PDF (គ្រប់ ៦ ទំព័រ)' : 'ទាញយក PDF'}</span>
            </button>

            {/* Toggle Show/Hide Picture Panel (រូបមន្តគណនា & របារ Word Formatting) */}
            <button
              type="button"
              onClick={() => {
                const nextVal = !showControlPanel;
                setShowControlPanel(nextVal);
                try {
                  localStorage.setItem('sticker_team_report_show_control_panel', String(nextVal));
                } catch {}
              }}
              className={`px-3 py-1.5 rounded text-xs font-bold border transition flex items-center gap-1.5 cursor-pointer shadow-2xs shrink-0 whitespace-nowrap ${
                showControlPanel
                  ? 'bg-slate-800 hover:bg-slate-900 text-white border-slate-900'
                  : 'bg-amber-50 hover:bg-amber-100 text-amber-900 border-amber-300'
              }`}
              title={showControlPanel ? 'ចុចដើម្បីលាក់ផ្ទាំងរូបមន្ត និងឧបករណ៍កែសម្រួល' : 'ចុចដើម្បីបង្ហាញផ្ទាំងរូបមន្ត និងឧបករណ៍កែសម្រួល'}
            >
              {showControlPanel ? (
                <>
                  <EyeOff className="w-3.5 h-3.5 text-amber-300 shrink-0" />
                  <span>លាក់ផ្ទាំងជម្រើស</span>
                </>
              ) : (
                <>
                  <Eye className="w-3.5 h-3.5 text-amber-600 shrink-0" />
                  <span>បង្ហាញផ្ទាំងជម្រើស</span>
                </>
              )}
            </button>
          </div>
        </div>

        {/* 7 Tabs Navigation (Matching exact PDF structure + All-in-One view) */}
        <div className="flex flex-wrap items-center gap-1.5 mt-3 pt-3 border-t border-gray-200">
          {(
            [
              { id: 'opening', num: '១', label: 'ស្តុកដើមគ្រា' },
              { id: 'issued', num: '២', label: 'បើកពីក២' },
              { id: 'used', num: '៣', label: 'ប្រើប្រាស់តាមក្រុម' },
              { id: 'transferred', num: '៤', label: 'ផ្ទេរការប្រើប្រាស់' },
              { id: 'damaged', num: '៥', label: 'ខូច / ខ្វះ / បង្វិល' },
              { id: 'remaining', num: '៦', label: 'នៅសល់ចុងគ្រា (ជាក់ស្តែង)' },
              { id: 'allInOne', num: '★', label: 'បង្ហាញគ្រប់ទំព័រទាំង ៦ (All 6 Pages)' },
            ] as const
          ).map((tab) => {
            const isActive = activeTab === tab.id;
            return (
              <button
                key={tab.id}
                type="button"
                onClick={() => setActiveTab(tab.id)}
                className={`flex items-center gap-1.5 px-3 py-1.5 rounded-md text-xs font-bold transition cursor-pointer ${
                  isActive
                    ? 'bg-[#0B2545] text-[#E4CD98] shadow-sm ring-1 ring-[#C6A15B]'
                    : 'bg-gray-100 hover:bg-gray-200 text-gray-700 border border-gray-200'
                }`}
              >
                <span className={`w-4 h-4 rounded-full text-[10px] flex items-center justify-center font-bold ${
                  isActive ? 'bg-[#C6A15B] text-[#0B2545]' : 'bg-gray-300 text-gray-700'
                }`}>
                  {tab.num}
                </span>
                <span>{tab.label}</span>
              </button>
            );
          })}
        </div>

        {/* Picture: Formula Explanatory Note & Word Formatting Toolbar - Toggled by លាក់ផ្ទាំងជម្រើស */}
        {showControlPanel && (
          <>
            {/* Formula Explanatory Note & Data Status */}
        <div className="mt-2.5 px-3 py-2 bg-amber-50/80 border border-amber-200 rounded-md text-xs text-amber-950 flex flex-wrap items-center justify-between gap-2">
          <div className="flex flex-wrap items-center gap-2">
            <span className="font-bold text-[#0B2545] bg-amber-200/80 px-2 py-0.5 rounded text-[11px]">
              រូបមន្តគណនា:
            </span>
            <span className="font-medium text-gray-800">
              <strong>ស្តុកដើមគ្រា</strong> = ស្តុកចាស់ក្រុម + បើកពីក២ <span className="text-gray-500 text-[10px]">(ក្រោយ ៣០ វិច្ឆិកា ២០១៨)</span> <span className="text-red-700 font-bold">- ប្រើប្រាស់</span> <span className="text-red-700 font-bold">- ផ្ទេរ</span> <span className="text-red-700 font-bold">- ខូច/ខ្វះ/បង្វិល</span> <span className="text-gray-500 text-[11px]">(គិតមុនកាលបរិច្ឆេទចាប់ផ្តើម)</span>
            </span>
            {detectedOldStockCount > 0 && (
              <span className="inline-flex items-center gap-1 bg-emerald-100 text-emerald-900 border border-emerald-300 px-2 py-0.5 rounded text-[11px] font-bold">
                <Check className="w-3 h-3 text-emerald-700" />
                ប្រភព៖ «ស្តុកចាស់ក្រុម» ពីទិន្នន័យស្តុក ({detectedOldStockCount} ជួរ)
              </span>
            )}
          </div>
          
          <div className="flex items-center gap-2">
            <span className="text-[11px] font-bold text-gray-700">ប្រភពដើមគ្រា៖</span>
            <select
              value={openingSourceMode}
              onChange={(e) => setOpeningSourceMode(e.target.value as any)}
              className="bg-white border border-amber-300 rounded px-2 py-0.5 text-[11px] font-bold text-[#0B2545] cursor-pointer shadow-2xs"
            >
              <option value="auto">ស្វ័យប្រវត្តិ (យកស្តុកចាស់ក្រុមពេលមានទិន្នន័យ)</option>
              <option value="recordedOnly">ទិន្នន័យ «ស្តុកចាស់ក្រុម» ពីប្រព័ន្ធតែប៉ុណ្ណោះ</option>
              <option value="baseline2018">ទិន្នន័យគំរូដើមគ្រាផ្លូវការ (៣០ វិច្ឆិកា ២០១៨)</option>
            </select>
          </div>
        </div>

        {/* Microsoft Word Document Formatting Toolbar / Ribbon */}
        <div className="mt-3 pt-2.5 border-t border-gray-200 flex flex-wrap items-center gap-1.5 text-xs text-gray-800">
          {/* Edit Mode Toggle */}
          <button
            type="button"
            onClick={() => setIsWordEditMode(!isWordEditMode)}
            className={`flex items-center gap-1 px-2.5 py-1.5 rounded font-bold transition cursor-pointer ${
              isWordEditMode
                ? 'bg-blue-600 text-white shadow-xs'
                : 'bg-gray-100 text-gray-700 hover:bg-gray-200 border border-gray-300'
            }`}
            title="បើក/បិទ មុខងារកែអក្សរផ្ទាល់លើក្រដាស (Word Mode)"
          >
            <Edit3 className="w-3.5 h-3.5" />
            <span>{isWordEditMode ? 'Word Mode: ON' : 'Word Mode: OFF'}</span>
          </button>

          {/* Quick Save Draft */}
          <button
            type="button"
            onClick={handleSaveDocument}
            className="flex items-center gap-1 px-2.5 py-1.5 rounded font-bold bg-emerald-600 hover:bg-emerald-700 text-white transition shadow-xs cursor-pointer"
            title="រក្សាទុកទម្រង់ និងការកែប្រែ (Ctrl+S)"
          >
            <Save className="w-3.5 h-3.5" />
            <span>រក្សាទុក (Ctrl+S)</span>
          </button>

          <div className="h-5 w-[1px] bg-gray-300 mx-0.5" />

          {/* Undo / Redo */}
          <button
            type="button"
            onClick={() => executeCommand('undo')}
            className="p-1.5 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded border border-gray-300 cursor-pointer"
            title="ថយក្រោយ (Undo - Ctrl+Z)"
          >
            <Undo className="w-3.5 h-3.5" />
          </button>
          <button
            type="button"
            onClick={() => executeCommand('redo')}
            className="p-1.5 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded border border-gray-300 cursor-pointer"
            title="ទៅមុខវិញ (Redo - Ctrl+Y)"
          >
            <Redo className="w-3.5 h-3.5" />
          </button>

          <div className="h-5 w-[1px] bg-gray-300 mx-0.5" />

          {/* Font Family Dropdown */}
          <select
            value={activeFont}
            onChange={(e) => handleApplyFont(e.target.value)}
            className="bg-white border border-gray-300 rounded px-2 py-1 text-xs font-medium text-gray-800 cursor-pointer shadow-2xs focus:outline-none focus:border-blue-500"
            title="ពុម្ពអក្សរ (Font Family)"
          >
            <option value="Khmer OS Siemreap">Khmer OS Siemreap</option>
            <option value="Khmer OS Mool1">Khmer OS Mool1</option>
            <option value="Khmer OS Battambang">Khmer OS Battambang</option>
            <option value="Times New Roman">Times New Roman</option>
            <option value="Arial">Arial</option>
            <option value="Calibri">Calibri</option>
          </select>

          {/* Font Size Dropdown */}
          <select
            value={activeFontSize}
            onChange={(e) => handleApplyFontSize(e.target.value)}
            className="bg-white border border-gray-300 rounded px-2 py-1 text-xs font-medium text-gray-800 cursor-pointer shadow-2xs focus:outline-none focus:border-blue-500"
            title="ទំហំអក្សរ (Font Size)"
          >
            {FONT_SIZE_STEPS.map((sz) => (
              <option key={sz} value={sz}>
                {sz}
              </option>
            ))}
          </select>

          {/* Font Size Increment / Decrement */}
          <button
            type="button"
            onClick={handleGrowFontSize}
            className="px-1.5 py-1 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded border border-gray-300 font-bold text-xs cursor-pointer"
            title="ពង្រីកទំហំអក្សរ (Grow Font)"
          >
            A+
          </button>
          <button
            type="button"
            onClick={handleShrinkFontSize}
            className="px-1.5 py-1 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded border border-gray-300 font-bold text-xs cursor-pointer"
            title="បង្រួមទំហំអក្សរ (Shrink Font)"
          >
            A-
          </button>

          {/* Condensed (Ctrl+D) / Font Dialog Launcher Button */}
          <button
            type="button"
            onClick={() => {
              const sel = window.getSelection();
              if (sel && sel.rangeCount > 0) {
                savedSelectionRef.current = sel.getRangeAt(0).cloneRange();
              }
              setShowFontDialog(true);
            }}
            className="flex items-center gap-1.5 px-2.5 py-1 bg-gradient-to-r from-amber-500 to-orange-500 hover:from-amber-600 hover:to-orange-600 text-white rounded font-bold text-xs shadow-xs cursor-pointer ring-1 ring-orange-400/50"
            title="ផ្ទាំងតម្រឹមកម្រិតខ្ពស់ Condensed & Font (Ctrl+D)"
          >
            <Sliders className="w-3.5 h-3.5" />
            <span>Condensed (Ctrl+D)</span>
            <span className="bg-black/20 text-[10px] px-1 py-0.2 rounded">Ctrl+D</span>
          </button>

          <div className="h-5 w-[1px] bg-gray-300 mx-0.5" />

          {/* Bold, Italic, Underline, Strikethrough */}
          <button
            type="button"
            onClick={() => executeCommand('bold')}
            className="p-1.5 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded border border-gray-300 cursor-pointer font-bold"
            title="អក្សរដិត Bold (Ctrl+B)"
          >
            <Bold className="w-3.5 h-3.5" />
          </button>
          <button
            type="button"
            onClick={() => executeCommand('italic')}
            className="p-1.5 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded border border-gray-300 cursor-pointer italic"
            title="អក្សរទ្រេត Italic (Ctrl+I)"
          >
            <Italic className="w-3.5 h-3.5" />
          </button>
          <button
            type="button"
            onClick={() => executeCommand('underline')}
            className="p-1.5 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded border border-gray-300 cursor-pointer underline"
            title="គូសបន្ទាត់ពីក្រោម Underline (Ctrl+U)"
          >
            <Underline className="w-3.5 h-3.5" />
          </button>
          <button
            type="button"
            onClick={() => executeCommand('strikeThrough')}
            className="p-1.5 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded border border-gray-300 cursor-pointer"
            title="គូសឆូត Strikethrough"
          >
            <Strikethrough className="w-3.5 h-3.5" />
          </button>
          <button
            type="button"
            onClick={() => executeCommand('subscript')}
            className="p-1.5 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded border border-gray-300 cursor-pointer"
            title="អក្សរជើង Subscript"
          >
            <Subscript className="w-3.5 h-3.5" />
          </button>
          <button
            type="button"
            onClick={() => executeCommand('superscript')}
            className="p-1.5 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded border border-gray-300 cursor-pointer"
            title="អក្សរលើ Superscript"
          >
            <Superscript className="w-3.5 h-3.5" />
          </button>

          {/* Color & Highlight */}
          <div className="relative">
            <button
              type="button"
              onClick={() => setShowColorPicker(!showColorPicker)}
              className="p-1.5 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded border border-gray-300 cursor-pointer flex items-center gap-1"
              title="ពណ៌អក្សរ (Font Color)"
            >
              <Palette className="w-3.5 h-3.5 text-blue-600" />
            </button>
            {showColorPicker && (
              <div className="absolute top-full mt-1 bg-white border border-gray-300 rounded shadow-lg p-2 z-50 grid grid-cols-5 gap-1.5 w-36">
                {[
                  '#000000', '#1E40AF', '#B91C1C', '#047857', '#B45309',
                  '#6B21A8', '#4B5563', '#DC2626', '#2563EB', '#059669',
                ].map((c) => (
                  <button
                    key={c}
                    type="button"
                    onClick={() => handleApplyTextColor(c)}
                    className="w-5 h-5 rounded border border-gray-400 cursor-pointer hover:scale-110 transition"
                    style={{ backgroundColor: c }}
                  />
                ))}
              </div>
            )}
          </div>

          <div className="relative">
            <button
              type="button"
              onClick={() => setShowHighlightPicker(!showHighlightPicker)}
              className="p-1.5 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded border border-gray-300 cursor-pointer flex items-center gap-1"
              title="ពណ៌រំលេចអត្ថបទ (Text Highlight Color)"
            >
              <Highlighter className="w-3.5 h-3.5 text-amber-500" />
            </button>
            {showHighlightPicker && (
              <div className="absolute top-full mt-1 bg-white border border-gray-300 rounded shadow-lg p-2 z-50 grid grid-cols-4 gap-1.5 w-32">
                {[
                  '#FEF08A', '#BBF7D0', '#BAE6FD', '#FBCFE8',
                  '#FED7AA', '#DDD6FE', '#E2E8F0', 'transparent',
                ].map((c) => (
                  <button
                    key={c}
                    type="button"
                    onClick={() => handleApplyHighlight(c)}
                    className="w-5 h-5 rounded border border-gray-400 cursor-pointer hover:scale-110 transition text-[9px] flex items-center justify-center font-bold"
                    style={{ backgroundColor: c === 'transparent' ? '#FFFFFF' : c }}
                  >
                    {c === 'transparent' ? '✕' : ''}
                  </button>
                ))}
              </div>
            )}
          </div>

          <button
            type="button"
            onClick={handleClearFormatting}
            className="p-1.5 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded border border-gray-300 cursor-pointer"
            title="លុបទម្រង់ (Clear Formatting)"
          >
            <RemoveFormatting className="w-3.5 h-3.5" />
          </button>

          <div className="h-5 w-[1px] bg-gray-300 mx-0.5" />

          {/* Alignment */}
          <button
            type="button"
            onClick={() => executeCommand('justifyLeft')}
            className="p-1.5 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded border border-gray-300 cursor-pointer"
            title="តម្រឹមឆ្វេង Align Left"
          >
            <AlignLeft className="w-3.5 h-3.5" />
          </button>
          <button
            type="button"
            onClick={() => executeCommand('justifyCenter')}
            className="p-1.5 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded border border-gray-300 cursor-pointer"
            title="តម្រឹមកណ្តាល Align Center"
          >
            <AlignCenter className="w-3.5 h-3.5" />
          </button>
          <button
            type="button"
            onClick={() => executeCommand('justifyRight')}
            className="p-1.5 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded border border-gray-300 cursor-pointer"
            title="តម្រឹមស្តាំ Align Right"
          >
            <AlignRight className="w-3.5 h-3.5" />
          </button>
          <button
            type="button"
            onClick={() => executeCommand('justifyFull')}
            className="p-1.5 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded border border-gray-300 cursor-pointer"
            title="តម្រឹមសងខាង Justify"
          >
            <AlignJustify className="w-3.5 h-3.5" />
          </button>

          {/* Line Spacing */}
          <select
            value={lineSpacing}
            onChange={(e) => handleApplyLineSpacing(e.target.value)}
            className="bg-white border border-gray-300 rounded px-2 py-1 text-xs text-gray-800 cursor-pointer shadow-2xs focus:outline-none"
            title="គម្លាតបន្ទាត់ (Line Spacing)"
          >
            <option value="1.0">1.0 (Single)</option>
            <option value="1.15">1.15</option>
            <option value="1.25">1.25</option>
            <option value="1.5">1.5 (Default)</option>
            <option value="2.0">2.0 (Double)</option>
          </select>

          <div className="h-5 w-[1px] bg-gray-300 mx-0.5" />

          {/* Tab Stop & Insert Tab */}
          <div className="flex items-center gap-1 bg-gray-50 px-2 py-0.5 rounded border border-gray-300">
            <span className="text-[10.5px] font-bold text-gray-600">Tab Stop:</span>
            <select
              value={activeTabStop}
              onChange={(e) => setActiveTabStop(parseFloat(e.target.value))}
              className="bg-white border border-gray-300 rounded px-1.5 py-0.5 text-xs text-black cursor-pointer font-mono font-bold"
            >
              <option value={1.5}>1.5 cm</option>
              <option value={2.0}>2.0 cm</option>
              <option value={2.25}>2.25 cm</option>
              <option value={2.5}>2.5 cm</option>
              <option value={3.0}>3.0 cm</option>
            </select>
            <button
              type="button"
              onClick={handleInsertTab}
              className="px-2 py-0.5 bg-blue-600 hover:bg-blue-700 text-white rounded text-[11px] font-bold cursor-pointer transition shadow-2xs"
              title="បញ្ចូលដកឃ្លា Tab Stop លើអត្ថបទដែលបានជ្រើស"
            >
              + ចុច Tab
            </button>
          </div>

          {/* Find & Replace Toggle */}
          <button
            type="button"
            onClick={() => setShowFindReplace(!showFindReplace)}
            className={`flex items-center gap-1 px-2.5 py-1 rounded font-bold text-xs transition cursor-pointer ${
              showFindReplace
                ? 'bg-purple-600 text-white shadow-xs'
                : 'bg-gray-100 text-gray-700 hover:bg-gray-200 border border-gray-300'
            }`}
            title="ស្វែងរក និងជំនួសពាក្យ Find & Replace (Ctrl+F)"
          >
            <Search className="w-3.5 h-3.5" />
            <span>រក/ជំនួស (Ctrl+F)</span>
          </button>

          {/* Ruler Toggle */}
          <button
            type="button"
            onClick={() => setShowRuler(!showRuler)}
            className={`flex items-center gap-1 px-2 py-1 rounded font-medium text-xs transition cursor-pointer ${
              showRuler
                ? 'bg-blue-100 text-blue-900 border border-blue-300'
                : 'bg-gray-100 text-gray-700 hover:bg-gray-200 border border-gray-300'
            }`}
            title="បង្ហាញ/លាក់ បន្ទាត់វាស់ (Ruler)"
          >
            <Sliders className="w-3.5 h-3.5" />
            <span>បន្ទាត់វាស់ Ruler</span>
          </button>

          {/* Quick Margins Toggle */}
          <button
            type="button"
            onClick={() => setShowMarginControls(!showMarginControls)}
            className={`flex items-center gap-1 px-2 py-1 rounded font-medium text-xs transition cursor-pointer ${
              showMarginControls
                ? 'bg-blue-100 text-blue-900 border border-blue-300'
                : 'bg-gray-100 text-gray-700 hover:bg-gray-200 border border-gray-300'
            }`}
            title="កែសម្រួលគែមក្រដាស A4"
          >
            <Layout className="w-3.5 h-3.5" />
            <span>គែមក្រដាស</span>
          </button>

          {/* Microsoft Word Page Setup Dialog Launcher */}
          <button
            type="button"
            onClick={() => {
              setTempMargins(customMargins);
              setShowPageSetupDialog(true);
            }}
            className="flex items-center gap-1 px-2 py-1 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded border border-gray-300 font-medium text-xs cursor-pointer"
            title="បើកផ្ទាំង Page Setup (Alt+P+S+P)"
          >
            <Maximize2 className="w-3.5 h-3.5 text-blue-600" />
            <span>Page Setup...</span>
          </button>

          {/* Signature Position Controls Toggle */}
          <button
            type="button"
            onClick={() => setShowSignPositionPanel(!showSignPositionPanel)}
            className={`flex items-center gap-1 px-2 py-1 rounded font-medium text-xs transition cursor-pointer ${
              showSignPositionPanel
                ? 'bg-purple-100 text-purple-900 border border-purple-300'
                : 'bg-gray-100 text-gray-700 hover:bg-gray-200 border border-gray-300'
            }`}
            title="រំកិលទីតាំងហត្ថលេខា ឆ្វេង កណ្តាល ស្តាំ"
          >
            <Move className="w-3.5 h-3.5 text-purple-600" />
            <span>រំកិលហត្ថលេខា</span>
          </button>
        </div>

        {/* Save Success Toast */}
        {showSaveSuccessToast && (
          <div className="mt-2 p-2 bg-emerald-100 border border-emerald-300 rounded text-emerald-900 text-xs font-bold flex items-center justify-between animate-fadeIn">
            <span className="flex items-center gap-1.5">
              <CheckCircle2 className="w-4 h-4 text-emerald-600" />
              <span>បានរក្សាទុកទម្រង់ និងការកែប្រែដោយជោគជ័យ! {lastSavedTime && `(${lastSavedTime})`}</span>
            </span>
            <button
              type="button"
              onClick={() => setShowSaveSuccessToast(false)}
              className="text-emerald-700 hover:text-emerald-950 font-bold cursor-pointer"
            >
              ✕
            </button>
          </div>
        )}

        {/* Find & Replace Panel */}
        {showFindReplace && (
          <div className="mt-2 p-2.5 bg-slate-900 text-white rounded-lg border border-purple-500/40 flex flex-wrap items-center gap-2 text-xs">
            <div className="flex items-center gap-1 bg-slate-800 px-2 py-1 rounded border border-slate-700">
              <span className="text-gray-400">ស្វែងរក:</span>
              <input
                type="text"
                placeholder="ពាក្យស្វែងរក..."
                value={findText}
                onChange={(e) => setFindText(e.target.value)}
                onKeyDown={(e) => e.key === 'Enter' && handleFindNext()}
                className="bg-transparent text-white focus:outline-none w-32 sm:w-44 text-xs font-medium"
              />
            </div>
            <button
              type="button"
              onClick={handleFindNext}
              className="px-2.5 py-1 bg-purple-600 hover:bg-purple-500 text-white rounded font-medium cursor-pointer transition"
            >
              រកបន្ទាប់ (Find Next)
            </button>

            <div className="flex items-center gap-1 bg-slate-800 px-2 py-1 rounded border border-slate-700">
              <span className="text-gray-400">ជំនួសដោយ:</span>
              <input
                type="text"
                placeholder="ពាក្យជំនួស..."
                value={replaceText}
                onChange={(e) => setReplaceText(e.target.value)}
                className="bg-transparent text-white focus:outline-none w-32 sm:w-44 text-xs font-medium"
              />
            </div>
            <button
              type="button"
              onClick={handleReplaceAll}
              className="px-2.5 py-1 bg-blue-600 hover:bg-blue-500 text-white rounded font-medium cursor-pointer transition"
            >
              ជំនួសទាំងអស់ (Replace All)
            </button>

            {matchCount !== null && (
              <span className="text-emerald-400 text-xs font-semibold ml-1">
                បានជំនួស {matchCount} កន្លែង
              </span>
            )}

            <button
              type="button"
              onClick={() => setShowFindReplace(false)}
              className="ml-auto p-1 text-gray-400 hover:text-white rounded cursor-pointer"
            >
              <X className="w-3.5 h-3.5" />
            </button>
          </div>
        )}

        {/* Quick Margin Controls Sub-bar */}
        {showMarginControls && (
          <div className="mt-2 bg-slate-900 text-white rounded-xl p-3 border border-slate-700 space-y-3 text-xs">
            <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-800 pb-2">
              <span className="text-gray-300 font-semibold flex items-center gap-1.5">
                <Sliders className="w-3.5 h-3.5 text-blue-400" />
                <span>កែសម្រួលគែមក្រដាស (Page Margins)</span>
              </span>
              <div className="flex flex-wrap items-center gap-1.5">
                <button
                  type="button"
                  onClick={() => setCustomMargins({ top: 0.4, bottom: 0.4, left: 1.8, right: 1.0 })}
                  className={`px-2.5 py-1 rounded text-[11px] font-medium border transition cursor-pointer ${
                    customMargins.left === 1.8 && customMargins.right === 1.0 && customMargins.top === 0.4 && customMargins.bottom === 0.4
                      ? 'bg-blue-600 border-blue-500 text-white'
                      : 'bg-slate-800 border-slate-700 text-gray-300 hover:text-white'
                  }`}
                >
                  ស្តង់ដារតារាង (ឆ្វេង 1.8cm, ស្តាំ 1.0cm, លើ 0.4cm, ក្រោម 0.4cm)
                </button>
                <button
                  type="button"
                  onClick={() => setCustomMargins({ top: 1.2, bottom: 1.2, left: 2.5, right: 1.5 })}
                  className={`px-2.5 py-1 rounded text-[11px] font-medium border transition cursor-pointer ${
                    customMargins.left === 2.5 && customMargins.right === 1.5 && customMargins.top === 1.2
                      ? 'bg-blue-600 border-blue-500 text-white'
                      : 'bg-slate-800 border-slate-700 text-gray-300 hover:text-white'
                  }`}
                >
                  ស្តង់ដាររដ្ឋបាល (2.5 / 1.5 / 1.2)
                </button>
                <button
                  type="button"
                  onClick={() => setCustomMargins({ top: 1.0, bottom: 1.0, left: 2.0, right: 1.2 })}
                  className={`px-2.5 py-1 rounded text-[11px] font-medium border transition cursor-pointer ${
                    customMargins.left === 2.0 && customMargins.right === 1.2
                      ? 'bg-blue-600 border-blue-500 text-white'
                      : 'bg-slate-800 border-slate-700 text-gray-300 hover:text-white'
                  }`}
                >
                  គែមត្បិត (ឆ្វេង 2.0cm)
                </button>
              </div>
            </div>

            <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
              <div>
                <label className="block text-gray-400 mb-1">គែមលើ (Top): {customMargins.top} cm</label>
                <input
                  type="range"
                  min="0.2"
                  max="3.0"
                  step="0.1"
                  value={customMargins.top}
                  onChange={(e) => setCustomMargins({ ...customMargins, top: parseFloat(e.target.value) })}
                  className="w-full accent-blue-500 cursor-pointer"
                />
              </div>
              <div>
                <label className="block text-gray-400 mb-1">គែមក្រោម (Bottom): {customMargins.bottom} cm</label>
                <input
                  type="range"
                  min="0.2"
                  max="3.0"
                  step="0.1"
                  value={customMargins.bottom}
                  onChange={(e) => setCustomMargins({ ...customMargins, bottom: parseFloat(e.target.value) })}
                  className="w-full accent-blue-500 cursor-pointer"
                />
              </div>
              <div>
                <label className="block text-gray-400 mb-1 font-bold text-amber-300">
                  គែមឆ្វេង (Left): {customMargins.left} cm
                </label>
                <input
                  type="range"
                  min="0.5"
                  max="4.0"
                  step="0.1"
                  value={customMargins.left}
                  onChange={(e) => setCustomMargins({ ...customMargins, left: parseFloat(e.target.value) })}
                  className="w-full accent-amber-500 cursor-pointer"
                />
              </div>
              <div>
                <label className="block text-gray-400 mb-1">គែមស្តាំ (Right): {customMargins.right} cm</label>
                <input
                  type="range"
                  min="0.5"
                  max="3.0"
                  step="0.1"
                  value={customMargins.right}
                  onChange={(e) => setCustomMargins({ ...customMargins, right: parseFloat(e.target.value) })}
                  className="w-full accent-blue-500 cursor-pointer"
                />
              </div>
            </div>
          </div>
        )}

        {/* Signature Position Controls Sub-bar */}
        {showSignPositionPanel && (
          <div className="mt-2 bg-slate-900 text-white rounded-xl p-3 border border-purple-500/50 space-y-3 text-xs">
            <div className="flex items-center justify-between border-b border-slate-800 pb-2">
              <span className="font-bold text-amber-400 flex items-center gap-1.5">
                <Move className="w-4 h-4" />
                <span>រំកិលទីតាំងហត្ថលេខា (Signature Shifts)</span>
              </span>
              <button
                type="button"
                onClick={() => {
                  setSignSectionMarginTop(0.2);
                  setSignShiftLeftX(0);
                  setSignShiftLeftY(0);
                  setSignShiftMidX(0);
                  setSignShiftMidY(0);
                  setSignShiftRightX(0);
                  setSignShiftRightY(0);
                }}
                className="px-2 py-0.5 bg-slate-800 hover:bg-slate-700 text-gray-300 rounded border border-slate-600 text-[10.5px] cursor-pointer"
              >
                កំណត់ដើមវិញ (Reset)
              </button>
            </div>

            <div className="grid grid-cols-1 sm:grid-cols-4 gap-3">
              {/* Overall Top Spacing */}
              <div className="bg-slate-800/80 p-2 rounded-lg border border-slate-700 space-y-1">
                <div className="text-[11px] font-bold text-purple-300">គម្លាតពីលើ (Top Margin):</div>
                <div className="flex items-center gap-1">
                  <button
                    type="button"
                    onClick={() => setSignSectionMarginTop((v) => Math.max(0, parseFloat((v - 0.1).toFixed(1))))}
                    className="px-2 py-0.5 bg-slate-700 hover:bg-purple-600 rounded text-white font-bold cursor-pointer"
                  >
                    ▲
                  </button>
                  <span className="flex-1 text-center font-mono font-bold text-purple-300 bg-slate-900 py-0.5 rounded text-[11px]">
                    {signSectionMarginTop}cm
                  </span>
                  <button
                    type="button"
                    onClick={() => setSignSectionMarginTop((v) => parseFloat((v + 0.1).toFixed(1)))}
                    className="px-2 py-0.5 bg-slate-700 hover:bg-purple-600 rounded text-white font-bold cursor-pointer"
                  >
                    ▼
                  </button>
                </div>
              </div>

              {/* Left Sign Offset */}
              <div className="bg-slate-800/80 p-2 rounded-lg border border-slate-700 space-y-1">
                <div className="text-[11px] font-bold text-blue-300">ហត្ថលេខាឆ្វេង (X / Y):</div>
                <div className="grid grid-cols-2 gap-1 text-[10px]">
                  <div className="flex items-center gap-0.5">
                    <button
                      type="button"
                      onClick={() => setSignShiftLeftX((x) => x - 5)}
                      className="px-1.5 py-0.5 bg-slate-700 hover:bg-blue-600 rounded cursor-pointer"
                    >
                      ◀
                    </button>
                    <span className="flex-1 text-center font-mono text-amber-300">{signShiftLeftX}px</span>
                    <button
                      type="button"
                      onClick={() => setSignShiftLeftX((x) => x + 5)}
                      className="px-1.5 py-0.5 bg-slate-700 hover:bg-blue-600 rounded cursor-pointer"
                    >
                      ▶
                    </button>
                  </div>
                  <div className="flex items-center gap-0.5">
                    <button
                      type="button"
                      onClick={() => setSignShiftLeftY((y) => y - 5)}
                      className="px-1.5 py-0.5 bg-slate-700 hover:bg-blue-600 rounded cursor-pointer"
                    >
                      ▲
                    </button>
                    <span className="flex-1 text-center font-mono text-cyan-300">{signShiftLeftY}px</span>
                    <button
                      type="button"
                      onClick={() => setSignShiftLeftY((y) => y + 5)}
                      className="px-1.5 py-0.5 bg-slate-700 hover:bg-blue-600 rounded cursor-pointer"
                    >
                      ▼
                    </button>
                  </div>
                </div>
              </div>

              {/* Middle Sign Offset */}
              <div className="bg-slate-800/80 p-2 rounded-lg border border-slate-700 space-y-1">
                <div className="text-[11px] font-bold text-emerald-300">ហត្ថលេខាកណ្តាល (X / Y):</div>
                <div className="grid grid-cols-2 gap-1 text-[10px]">
                  <div className="flex items-center gap-0.5">
                    <button
                      type="button"
                      onClick={() => setSignShiftMidX((x) => x - 5)}
                      className="px-1.5 py-0.5 bg-slate-700 hover:bg-emerald-600 rounded cursor-pointer"
                    >
                      ◀
                    </button>
                    <span className="flex-1 text-center font-mono text-amber-300">{signShiftMidX}px</span>
                    <button
                      type="button"
                      onClick={() => setSignShiftMidX((x) => x + 5)}
                      className="px-1.5 py-0.5 bg-slate-700 hover:bg-emerald-600 rounded cursor-pointer"
                    >
                      ▶
                    </button>
                  </div>
                  <div className="flex items-center gap-0.5">
                    <button
                      type="button"
                      onClick={() => setSignShiftMidY((y) => y - 5)}
                      className="px-1.5 py-0.5 bg-slate-700 hover:bg-emerald-600 rounded cursor-pointer"
                    >
                      ▲
                    </button>
                    <span className="flex-1 text-center font-mono text-cyan-300">{signShiftMidY}px</span>
                    <button
                      type="button"
                      onClick={() => setSignShiftMidY((y) => y + 5)}
                      className="px-1.5 py-0.5 bg-slate-700 hover:bg-emerald-600 rounded cursor-pointer"
                    >
                      ▼
                    </button>
                  </div>
                </div>
              </div>

              {/* Right Sign Offset */}
              <div className="bg-slate-800/80 p-2 rounded-lg border border-slate-700 space-y-1">
                <div className="text-[11px] font-bold text-orange-300">ហត្ថលេខាស្តាំ (X / Y):</div>
                <div className="grid grid-cols-2 gap-1 text-[10px]">
                  <div className="flex items-center gap-0.5">
                    <button
                      type="button"
                      onClick={() => setSignShiftRightX((x) => x - 5)}
                      className="px-1.5 py-0.5 bg-slate-700 hover:bg-orange-600 rounded cursor-pointer"
                    >
                      ◀
                    </button>
                    <span className="flex-1 text-center font-mono text-amber-300">{signShiftRightX}px</span>
                    <button
                      type="button"
                      onClick={() => setSignShiftRightX((x) => x + 5)}
                      className="px-1.5 py-0.5 bg-slate-700 hover:bg-orange-600 rounded cursor-pointer"
                    >
                      ▶
                    </button>
                  </div>
                  <div className="flex items-center gap-0.5">
                    <button
                      type="button"
                      onClick={() => setSignShiftRightY((y) => y - 5)}
                      className="px-1.5 py-0.5 bg-slate-700 hover:bg-orange-600 rounded cursor-pointer"
                    >
                      ▲
                    </button>
                    <span className="flex-1 text-center font-mono text-cyan-300">{signShiftRightY}px</span>
                    <button
                      type="button"
                      onClick={() => setSignShiftRightY((y) => y + 5)}
                      className="px-1.5 py-0.5 bg-slate-700 hover:bg-orange-600 rounded cursor-pointer"
                    >
                      ▼
                    </button>
                  </div>
                </div>
              </div>
            </div>
          </div>
        )}
          </>
        )}
      </div>

      {/* Main Content Area (A4 Document Canvas Container) */}
      <div className="bg-gray-200/80 p-2 sm:p-6 rounded-md overflow-x-auto flex justify-center print:p-0 print:bg-white print:block">
        <div
          style={{
            transform: docZoom !== 100 ? `scale(${docZoom / 100})` : undefined,
            transformOrigin: 'top center',
            transition: 'transform 0.15s ease-out',
          }}
        >
          {/* Microsoft Word Horizontal Ruler (Top of Page) */}
          {showRuler && (
            <div
              onClick={handleRulerClick}
              className="print:hidden relative mb-1.5 bg-slate-100 select-none shadow-xs rounded-t-md border border-slate-300 overflow-hidden font-sans cursor-crosshair group mx-auto"
              style={{ width: '210mm', height: '28px' }}
              title="ចុចលើ Ruler ដើម្បីកំណត់ចំណុច Tab Stop (ឧ. 2cm, 2.25cm, 2.5cm) សម្រាប់ចុច Tab ដកឃ្លា"
            >
              {/* Shaded Left Margin Area */}
              <div
                className="absolute top-0 bottom-0 left-0 bg-slate-300/80 border-r border-slate-400/80 pointer-events-none"
                style={{ width: `${customMargins.left}cm` }}
              />
              {/* White Content Area */}
              <div
                className="absolute top-0 bottom-0 bg-white pointer-events-none"
                style={{
                  left: `${customMargins.left}cm`,
                  width: `${21.0 - customMargins.left - customMargins.right}cm`,
                }}
              />
              {/* Shaded Right Margin Area */}
              <div
                className="absolute top-0 bottom-0 right-0 bg-slate-300/80 border-l border-slate-400/80 pointer-events-none"
                style={{ width: `${customMargins.right}cm` }}
              />

              {/* cm Scale Numbers & Ticks */}
              <div className="relative w-full h-full pointer-events-none">
                {Array.from({ length: 22 }).map((_, cmIdx) => {
                  const relativeCm = Math.round(cmIdx - customMargins.left);
                  return (
                    <div
                      key={cmIdx}
                      className="absolute top-0 flex flex-col items-center pointer-events-none"
                      style={{ left: `${cmIdx}cm` }}
                    >
                      <div className="w-[1px] h-3 bg-slate-600" />
                      {/* Half cm tick */}
                      {cmIdx < 21 && (
                        <div
                          className="absolute top-0 w-[1px] h-2 bg-slate-400"
                          style={{ left: '0.5cm' }}
                        />
                      )}
                      {/* Quarter cm ticks */}
                      {cmIdx < 21 && (
                        <>
                          <div
                            className="absolute top-0 w-[1px] h-1.5 bg-slate-300"
                            style={{ left: '0.25cm' }}
                          />
                          <div
                            className="absolute top-0 w-[1px] h-1.5 bg-slate-300"
                            style={{ left: '0.75cm' }}
                          />
                        </>
                      )}
                      {/* Number label */}
                      <span className="text-[8.5px] font-semibold text-slate-700 mt-0.5 -ml-1">
                        {relativeCm >= 0 ? relativeCm : ''}
                      </span>
                    </div>
                  );
                })}
              </div>

              {/* Active Tab Stop Marker (Classic Word Tab Stop Marker on Ruler) */}
              <div
                className="absolute top-0 z-20 -ml-1.5 flex flex-col items-center pointer-events-none"
                style={{ left: `${customMargins.left + activeTabStop}cm` }}
                title={`Tab Stop កំណត់នៅ: ${activeTabStop} cm`}
              >
                <div className="w-0 h-0 border-l-[4px] border-l-transparent border-r-[4px] border-r-transparent border-t-[5px] border-t-amber-600" />
                <div className="w-2.5 h-3 bg-amber-500 rounded-[1px] shadow-xs flex items-center justify-center text-[7px] text-black font-extrabold">
                  {activeTabStop}
                </div>
                <div className="w-0 h-0 border-l-[4px] border-l-transparent border-r-[4px] border-r-transparent border-b-[5px] border-b-amber-600" />
              </div>

              {/* Left Margin Indicator Marker */}
              <div
                className="absolute top-0 z-10 -ml-1.5 flex flex-col items-center cursor-pointer group pointer-events-none"
                style={{ left: `${customMargins.left}cm` }}
                title={`គែមឆ្វេង (Left Margin): ${customMargins.left} cm`}
              >
                <div className="w-0 h-0 border-l-[4px] border-l-transparent border-r-[4px] border-r-transparent border-t-[5px] border-t-blue-600" />
                <div className="w-2 h-2.5 bg-blue-600 rounded-[1px] shadow-xs flex items-center justify-center text-[6.5px] text-white font-bold">
                  L
                </div>
                <div className="w-0 h-0 border-l-[4px] border-l-transparent border-r-[4px] border-r-transparent border-b-[5px] border-b-blue-600" />
              </div>

              {/* Right Margin Indicator Marker */}
              <div
                className="absolute top-0 z-10 -ml-1.5 flex flex-col items-center cursor-pointer group pointer-events-none"
                style={{ left: `${21.0 - customMargins.right}cm` }}
                title={`គែមស្តាំ (Right Margin): ${customMargins.right} cm`}
              >
                <div className="w-0 h-0 border-l-[4px] border-l-transparent border-r-[4px] border-r-transparent border-t-[5px] border-t-blue-600" />
                <div className="w-2 h-2.5 bg-blue-600 rounded-[1px] shadow-xs flex items-center justify-center text-[6.5px] text-white font-bold">
                  R
                </div>
                <div className="w-0 h-0 border-l-[4px] border-l-transparent border-r-[4px] border-r-transparent border-b-[5px] border-b-blue-600" />
              </div>
            </div>
          )}
          {activeTab === 'allInOne' ? (
            <div ref={allPagesRef} className="space-y-6 flex flex-col items-center">
              {/* Page 1: Opening (No signatures) */}
              {renderSinglePage('opening', openingMatrix, false)}

              {/* Page 2: Issued (No signatures) */}
              {renderSinglePage('issued', issuedMatrix, false)}

              {/* Page 3: Used (No signatures) */}
              {renderSinglePage('used', usedMatrix, false)}

              {/* Page 4: Transferred (No signatures) */}
              {renderSinglePage('transferred', transferredMatrix, false)}

              {/* Page 5: Damaged (No signatures) */}
              {renderSinglePage('damaged', damagedMatrix, false)}

              {/* Page 6: Remaining (ONLY Page 6 has official signatures block) */}
              {renderSinglePage('remaining', remainingMatrix, true)}
            </div>
          ) : (
            <div ref={reportRef} className="flex flex-col items-center">
              {renderSinglePage(activeTab, currentMatrix, activeTab === 'remaining')}
            </div>
          )}
        </div>
      </div>

      {/* Team Data Audit & Inspection Modal */}
      {isAuditModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-3 sm:p-6 bg-black/60 backdrop-blur-xs animate-fadeIn">
          <div className="bg-white rounded-xl shadow-2xl border border-gray-200 w-full max-w-5xl max-h-[90vh] flex flex-col overflow-hidden">
            {/* Modal Header */}
            <div className="px-5 py-3.5 bg-[#0B2545] text-white flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <Search className="w-5 h-5 text-[#E4CD98]" />
                <div>
                  <h3 className="font-bold text-sm text-white">
                    ផ្ទៀងផ្ទាត់ទិន្នន័យ (បូក / មិនបូក) តាមបណ្តាក្រុមច្រកទ្វារអន្តរជាតិ
                  </h3>
                  <p className="text-[11px] text-[#E4CD98]">
                    ពិនិត្យមើលប្រតិបត្តិការជាក់ស្តែង និងមូលហេតុដែលទិន្នន័យត្រូវបានរាប់បញ្ចូល ឬមិនត្រូវបានរាប់បញ្ចូល
                  </p>
                </div>
              </div>

              <button
                type="button"
                onClick={() => setIsAuditModalOpen(false)}
                className="text-gray-300 hover:text-white p-1 rounded-md hover:bg-white/10 transition cursor-pointer"
              >
                ✕
              </button>
            </div>

            {/* Team Selection & Quick Chips */}
            <div className="p-4 bg-gray-50 border-b border-gray-200 space-y-3">
              <div className="flex flex-wrap items-center gap-2">
                <span className="text-xs font-bold text-gray-700">ជ្រើសរើសក្រុម/ប៉ុស្តិ៍៖</span>
                <select
                  value={selectedAuditTeam}
                  onChange={(e) => setSelectedAuditTeam(e.target.value)}
                  className="bg-white border border-gray-300 rounded-md px-3 py-1.5 text-xs font-bold text-[#0B2545] cursor-pointer shadow-2xs focus:ring-2 focus:ring-blue-500 focus:outline-none"
                >
                  {teamsList.map((t) => (
                    <option key={t} value={t}>
                      {t} {teamFullNameMap[t] ? `(${teamFullNameMap[t]})` : ''}
                    </option>
                  ))}
                </select>

                <div className="ml-auto flex items-center gap-2 text-xs text-gray-600 bg-white px-2.5 py-1 rounded border border-gray-200">
                  <Calendar className="w-3.5 h-3.5 text-gray-500" />
                  <span>ជួរកាលបរិច្ឆេទរបាយការណ៍៖</span>
                  <span className="font-bold text-gray-800 font-times">{startDate || 'ដើមគ្រា'}</span>
                  <span>ដល់</span>
                  <span className="font-bold text-gray-800 font-times">{endDate || 'ចុងគ្រា'}</span>
                </div>
              </div>

              {/* Quick Filter Chips for Key Border Gates */}
              <div className="flex flex-wrap items-center gap-1.5">
                <span className="text-[11px] font-semibold text-gray-500">ប៉ុស្តិ៍សំខាន់ៗ៖</span>
                {[
                  'ព្រំដែន ព្រំ',
                  'ព្រំដែន ព្រែកចាក',
                  'កំពង់ផែ ព្រះសីហនុ',
                  'ព្រំដែន ត្រពាំងស្រែ',
                  'ព្រំដែន បាវិត',
                  'ព្រំដែន ក្អមសំណ',
                  'ព្រំដែន អូរយ៉ាដាវ',
                  'កំពង់ផែ ភ្នំពេញ',
                ].map((gate) => {
                  const isSelected = selectedAuditTeam === gate;
                  return (
                    <button
                      key={gate}
                      type="button"
                      onClick={() => setSelectedAuditTeam(gate)}
                      className={`px-2 py-0.5 rounded text-[11px] font-medium transition cursor-pointer border ${
                        isSelected
                          ? 'bg-[#0B2545] text-white border-[#0B2545] shadow-2xs'
                          : 'bg-white text-gray-700 border-gray-300 hover:bg-gray-100'
                      }`}
                    >
                      {gate}
                    </button>
                  );
                })}
              </div>
            </div>

            {/* Team Summary Cards */}
            {(() => {
              const opSum = VISA_TYPES.reduce((s, vt) => s + (openingMatrix[selectedAuditTeam]?.[vt] || 0), 0);
              const issSum = VISA_TYPES.reduce((s, vt) => s + (issuedMatrix[selectedAuditTeam]?.[vt] || 0), 0);
              const useSum = VISA_TYPES.reduce((s, vt) => s + (usedMatrix[selectedAuditTeam]?.[vt] || 0), 0);
              const transSum = VISA_TYPES.reduce((s, vt) => s + (transferredMatrix[selectedAuditTeam]?.[vt] || 0), 0);
              const damSum = VISA_TYPES.reduce((s, vt) => s + (damagedMatrix[selectedAuditTeam]?.[vt] || 0), 0);
              const remSum = VISA_TYPES.reduce((s, vt) => s + (remainingMatrix[selectedAuditTeam]?.[vt] || 0), 0);

              return (
                <div className="grid grid-cols-2 sm:grid-cols-6 gap-2 p-4 bg-white border-b border-gray-200 text-center">
                  <div className="bg-amber-50/70 border border-amber-200 rounded p-2">
                    <p className="text-[10.5px] font-bold text-amber-900">១. ដើមគ្រា</p>
                    <p className="text-sm font-bold text-gray-900 font-times mt-0.5">{opSum.toLocaleString()}</p>
                  </div>
                  <div className="bg-blue-50/70 border border-blue-200 rounded p-2">
                    <p className="text-[10.5px] font-bold text-blue-900">២. បើកពីក២</p>
                    <p className="text-sm font-bold text-blue-950 font-times mt-0.5">{issSum.toLocaleString()}</p>
                  </div>
                  <div className="bg-emerald-50/70 border border-emerald-200 rounded p-2">
                    <p className="text-[10.5px] font-bold text-emerald-900">៣. ប្រើប្រាស់</p>
                    <p className="text-sm font-bold text-emerald-950 font-times mt-0.5">{useSum.toLocaleString()}</p>
                  </div>
                  <div className="bg-purple-50/70 border border-purple-200 rounded p-2">
                    <p className="text-[10.5px] font-bold text-purple-900">៤. ផ្ទេរ</p>
                    <p className="text-sm font-bold text-purple-950 font-times mt-0.5">{transSum.toLocaleString()}</p>
                  </div>
                  <div className="bg-rose-50/70 border border-rose-200 rounded p-2">
                    <p className="text-[10.5px] font-bold text-rose-900">៥. ខូច/ខ្វះ/បង្វិល</p>
                    <p className="text-sm font-bold text-rose-950 font-times mt-0.5">{damSum.toLocaleString()}</p>
                  </div>
                  <div className="bg-slate-100 border border-slate-300 rounded p-2">
                    <p className="text-[10.5px] font-bold text-slate-800">៦. នៅសល់ចុងគ្រា</p>
                    <p className="text-sm font-bold text-[#0B2545] font-times mt-0.5">{remSum.toLocaleString()}</p>
                  </div>
                </div>
              );
            })()}

            {/* Audit Records List & Filter */}
            {(() => {
              const normTarget = normalizeTeamName(selectedAuditTeam);
              const normStart = normalizeDateToISO(startDate);
              const normEnd = normalizeDateToISO(endDate);

              const allTeamRecords = (stockRecords || []).filter((rec) => {
                const rawTeam = resolveRecordTeamName(rec);
                const normRawTeam = normalizeTeamName(rawTeam);
                return (
                  normRawTeam === normTarget ||
                  (normRawTeam && normTarget && (normRawTeam.includes(normTarget) || normTarget.includes(normRawTeam))) ||
                  (rec.visaTeamRobokName && normalizeTeamName(rec.visaTeamRobokName) === normTarget) ||
                  (rec.sourceFrom && normalizeTeamName(rec.sourceFrom) === normTarget) ||
                  (rec.requesterName && normalizeTeamName(rec.requesterName) === normTarget)
                );
              });

              const analyzedRecords = allTeamRecords.map((rec) => {
                const recDate = normalizeDateToISO(rec.date || '');
                const isCea = isCeaRecord(rec);
                const isOldStock = isOldStockTeamRecord(rec);
                const vt = normalizeVisaType(rec.visaType);
                const isValidVt = VISA_TYPES.includes(vt as any);
                const qty = Number(rec.quantityBundles || rec.totalSheets || (rec as any).quantity || 0);
                const isSticker = !rec.stockType || rec.stockType === 'sticker';
                const isInRange = (!normStart || recDate >= normStart) && (!normEnd || recDate <= normEnd);

                let isCounted = false;
                let role = '';
                let reason = '';

                if (isCea) {
                  reason = 'ជាទិន្នន័យ cEA / e-Visa (មិនមែន Sticker)';
                } else if (!isSticker) {
                  reason = `ប្រភេទស្តុក «${rec.stockType}» មិនមែន Sticker`;
                } else if (!isValidVt) {
                  reason = `ប្រភេទទិដ្ឋាការ «${rec.visaType || 'ទទេ'}» មិនស្គាល់`;
                } else if (qty <= 0) {
                  reason = 'ចំនួនសន្លឹកស្មើ ០';
                } else if (activeTab === 'opening') {
                  if (isOldStock) {
                    isCounted = true;
                    role = 'ស្តុកចាស់ក្រុម (ដើមគ្រា)';
                  } else if (recDate && normStart && recDate < normStart) {
                    if (recDate <= '2018-11-30') {
                      reason = 'មុនថ្ងៃ ៣០ វិច្ឆិកា ២០១៨ (ច្បាប់កាត់ផ្តាច់)';
                    } else {
                      isCounted = true;
                      role = `ប្រតិបត្តិការមុនដើមគ្រា (${rec.operationType})`;
                    }
                  } else {
                    reason = `កាលបរិច្ឆេទ (${recDate}) មិនមែនមុនដើមគ្រា (${normStart})`;
                  }
                } else if (activeTab === 'issued') {
                  if (rec.operationType === 'issueTeam') {
                    if (isInRange) {
                      isCounted = true;
                      role = 'បើកពីក២ ចេញទៅក្រុម';
                    } else {
                      reason = `កាលបរិច្ឆេទ (${recDate}) ក្រៅជួរ (${normStart} ដល់ ${normEnd})`;
                    }
                  } else {
                    reason = `ប្រតិបត្តិការ «${rec.operationType}» មិនមែនការបើកពីក២ (issueTeam)`;
                  }
                } else if (activeTab === 'used') {
                  if (rec.operationType === 'useTeam') {
                    if (isInRange) {
                      isCounted = true;
                      role = 'ប្រើប្រាស់តាមក្រុម';
                    } else {
                      reason = `កាលបរិច្ឆេទ (${recDate}) ក្រៅជួរ (${normStart} ដល់ ${normEnd})`;
                    }
                  } else {
                    reason = `ប្រតិបត្តិការ «${rec.operationType}» មិនមែនការប្រើប្រាស់ (useTeam)`;
                  }
                } else if (activeTab === 'transferred') {
                  if (rec.operationType === 'transferTeam') {
                    if (isInRange) {
                      isCounted = true;
                      role = 'ផ្ទេរការប្រើប្រាស់';
                    } else {
                      reason = `កាលបរិច្ឆេទ (${recDate}) ក្រៅជួរ (${normStart} ដល់ ${normEnd})`;
                    }
                  } else {
                    reason = `ប្រតិបត្តិការ «${rec.operationType}» មិនមែនការផ្ទេរ (transferTeam)`;
                  }
                } else if (activeTab === 'damaged') {
                  if (
                    rec.operationType === 'damagedTeam' ||
                    rec.operationType === 'returnTeam' ||
                    rec.operationType === 'missingTeam'
                  ) {
                    if (isInRange) {
                      isCounted = true;
                      role = 'ខូច/ខ្វះ/បង្វិល';
                    } else {
                      reason = `កាលបរិច្ឆេទ (${recDate}) ក្រៅជួរ (${normStart} ដល់ ${normEnd})`;
                    }
                  } else {
                    reason = `ប្រតិបត្តិការ «${rec.operationType}» មិនមែនខូច/ខ្វះ/បង្វិល`;
                  }
                } else {
                  // remaining / allInOne
                  if (isInRange || (recDate && normStart && recDate < normStart)) {
                    isCounted = true;
                    role = `ប្រតិបត្តិការ ${rec.operationType}`;
                  } else {
                    reason = `កាលបរិច្ឆេទ (${recDate}) ក្រៅជួរ (${normStart} ដល់ ${normEnd})`;
                  }
                }

                return {
                  rec,
                  isCounted,
                  role: role || (isCounted ? 'រាប់បញ្ចូល' : ''),
                  reason: isCounted ? 'បានរាប់បញ្ចូលក្នុងតារាង' : reason,
                };
              });

              const countedList = analyzedRecords.filter((r) => r.isCounted);
              const uncountedList = analyzedRecords.filter((r) => !r.isCounted);

              const displayList = analyzedRecords.filter((item) => {
                if (auditTab === 'counted' && !item.isCounted) return false;
                if (auditTab === 'uncounted' && item.isCounted) return false;

                if (!auditSearchTerm) return true;
                const term = auditSearchTerm.toLowerCase();
                const r = item.rec;
                return (
                  (r.date && r.date.toLowerCase().includes(term)) ||
                  (r.visaType && r.visaType.toLowerCase().includes(term)) ||
                  (r.operationType && r.operationType.toLowerCase().includes(term)) ||
                  (r.sourceFrom && r.sourceFrom.toLowerCase().includes(term)) ||
                  (r.visaTeamRobokName && r.visaTeamRobokName.toLowerCase().includes(term)) ||
                  (r.startNumber && r.startNumber.toLowerCase().includes(term)) ||
                  (r.endNumber && r.endNumber.toLowerCase().includes(term)) ||
                  item.reason.toLowerCase().includes(term)
                );
              });

              return (
                <div className="flex-1 flex flex-col overflow-hidden bg-gray-50">
                  {/* Tab switch and search bar */}
                  <div className="p-3 bg-white border-b border-gray-200 flex flex-wrap items-center justify-between gap-3">
                    <div className="flex items-center gap-1.5">
                      <button
                        type="button"
                        onClick={() => setAuditTab('counted')}
                        className={`px-3 py-1.5 rounded-md text-xs font-bold transition cursor-pointer flex items-center gap-1.5 ${
                          auditTab === 'counted'
                            ? 'bg-emerald-700 text-white shadow-2xs'
                            : 'bg-gray-100 text-gray-700 hover:bg-gray-200'
                        }`}
                      >
                        <CheckCircle2 className="w-3.5 h-3.5" />
                        <span>ទិន្នន័យបានរាប់បញ្ចូល ({countedList.length})</span>
                      </button>

                      <button
                        type="button"
                        onClick={() => setAuditTab('uncounted')}
                        className={`px-3 py-1.5 rounded-md text-xs font-bold transition cursor-pointer flex items-center gap-1.5 ${
                          auditTab === 'uncounted'
                            ? 'bg-amber-600 text-white shadow-2xs'
                            : 'bg-gray-100 text-gray-700 hover:bg-gray-200'
                        }`}
                      >
                        <AlertCircle className="w-3.5 h-3.5" />
                        <span>ទិន្នន័យមិនបានរាប់ ({uncountedList.length})</span>
                      </button>

                      <button
                        type="button"
                        onClick={() => setAuditTab('all')}
                        className={`px-3 py-1.5 rounded-md text-xs font-bold transition cursor-pointer flex items-center gap-1.5 ${
                          auditTab === 'all'
                            ? 'bg-slate-700 text-white shadow-2xs'
                            : 'bg-gray-100 text-gray-700 hover:bg-gray-200'
                        }`}
                      >
                        <Layers className="w-3.5 h-3.5" />
                        <span>ទាំងអស់ ({analyzedRecords.length})</span>
                      </button>
                    </div>

                    <div className="relative w-full sm:w-64">
                      <Search className="w-3.5 h-3.5 absolute left-2.5 top-1/2 -translate-y-1/2 text-gray-400" />
                      <input
                        type="text"
                        placeholder="ស្វែងរកតាមលេខសន្លឹក, ប្រភេទ..."
                        value={auditSearchTerm}
                        onChange={(e) => setAuditSearchTerm(e.target.value)}
                        className="w-full bg-gray-50 border border-gray-300 rounded-md pl-8 pr-3 py-1 text-xs text-gray-800 placeholder-gray-400 focus:bg-white focus:outline-none focus:ring-1 focus:ring-blue-500"
                      />
                    </div>
                  </div>

                  {/* Records Table */}
                  <div className="flex-1 overflow-auto p-3">
                    {displayList.length === 0 ? (
                      <div className="text-center py-12 text-gray-500 text-xs">
                        <Info className="w-8 h-8 text-gray-400 mx-auto mb-2" />
                        <p className="font-semibold">ពុំមានទិន្នន័យត្រូវតាមលក្ខខណ្ឌស្វែងរកនេះទេ</p>
                      </div>
                    ) : (
                      <table className="w-full border-collapse text-xs bg-white rounded shadow-2xs overflow-hidden">
                        <thead>
                          <tr className="bg-gray-100 text-gray-700 text-[11px] font-bold border-b border-gray-200">
                            <th className="py-2 px-2 text-center w-10">ល.រ</th>
                            <th className="py-2 px-2 text-left">កាលបរិច្ឆេទ</th>
                            <th className="py-2 px-2 text-left">ប្រភេទប្រតិបត្តិការ</th>
                            <th className="py-2 px-2 text-center">ប្រភេទទិដ្ឋាការ</th>
                            <th className="py-2 px-2 text-right">ចំនួន (សន្លឹក)</th>
                            <th className="py-2 px-2 text-left">ចន្លោះលេខសន្លឹក</th>
                            <th className="py-2 px-2 text-left">ប្រភព/ក្រុមដើម</th>
                            <th className="py-2 px-2 text-left">ស្ថានភាពក្នុងតារាងបច្ចុប្បន្ន</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-gray-200">
                          {displayList.map((item, idx) => {
                            const r = item.rec;
                            const qty = Number(r.quantityBundles || r.totalSheets || (r as any).quantity || 0);

                            return (
                              <tr key={r.id || idx} className="hover:bg-gray-50 transition">
                                <td className="py-2 px-2 text-center text-gray-500 font-times">{idx + 1}</td>
                                <td className="py-2 px-2 font-times text-gray-800 whitespace-nowrap">{r.date || '—'}</td>
                                <td className="py-2 px-2">
                                  <span className="inline-block px-2 py-0.5 rounded text-[10.5px] font-semibold bg-gray-100 text-gray-800 border border-gray-200">
                                    {r.operationType === 'useTeam'
                                      ? 'ការប្រើប្រាស់តាមក្រុម'
                                      : r.operationType === 'issueTeam'
                                      ? 'ការបើកពីក២'
                                      : r.operationType === 'transferTeam'
                                      ? 'ផ្ទេរការប្រើប្រាស់'
                                      : r.operationType === 'damagedTeam'
                                      ? 'ខូចគុណភាពក្រុម'
                                      : r.operationType === 'oldStockTeam'
                                      ? 'ស្តុកចាស់ក្រុម'
                                      : r.operationType || 'ប្រតិបត្តិការ'}
                                  </span>
                                </td>
                                <td className="py-2 px-2 text-center font-bold font-times text-blue-900">
                                  {r.visaType || '—'}
                                </td>
                                <td className="py-2 px-2 text-right font-bold font-times text-gray-900">
                                  {qty.toLocaleString()}
                                </td>
                                <td className="py-2 px-2 font-times text-gray-600 text-[11px] whitespace-nowrap">
                                  {r.startNumber && r.endNumber
                                    ? `${r.startNumber} - ${r.endNumber}`
                                    : r.startNumber || r.endNumber || '—'}
                                </td>
                                <td className="py-2 px-2 text-gray-700 text-[11px]" title={r.visaTeamRobokName || r.sourceFrom || ''}>
                                  {r.operationType === 'transferTeam' ? (
                                    <span>
                                      {r.visaTeamRobokName || '—'} <strong className="text-blue-600">➔ ទទួល:</strong> {r.targetVisaTeamName || '—'}
                                    </span>
                                  ) : (
                                    r.visaTeamRobokName || r.sourceFrom || '—'
                                  )}
                                </td>
                                <td className="py-2 px-2">
                                  {item.isCounted ? (
                                    <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10.5px] font-bold bg-emerald-100 text-emerald-900 border border-emerald-300">
                                      <CheckCircle2 className="w-3 h-3 text-emerald-700 shrink-0" />
                                      <span>បូកបញ្ចូល: {item.role}</span>
                                    </span>
                                  ) : (
                                    <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10.5px] font-semibold bg-amber-100 text-amber-900 border border-amber-300">
                                      <AlertCircle className="w-3 h-3 text-amber-700 shrink-0" />
                                      <span>{item.reason}</span>
                                    </span>
                                  )}
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                      </table>
                    )}
                  </div>
                </div>
              );
            })()}

            {/* Modal Footer */}
            <div className="p-3 bg-gray-100 border-t border-gray-200 flex items-center justify-between text-xs">
              <span className="text-gray-500">
                ចំណាំ៖ ប្រព័ន្ធរក្សាទិន្នន័យដើមដែលបាន Import ដោយមិនកែប្រែ ឬលុបឡើយ។
              </span>
              <button
                type="button"
                onClick={() => setIsAuditModalOpen(false)}
                className="px-4 py-1.5 rounded bg-[#0B2545] hover:bg-[#071A33] text-white font-bold transition cursor-pointer"
              >
                បិទផ្ទាំង
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Microsoft Word Font Dialog Modal (Ctrl+D) */}
      {showFontDialog && (
        <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-[2px] flex items-center justify-center p-3">
          <div className="bg-[#F0F0F0] text-gray-900 rounded shadow-2xl border border-gray-400 w-full max-w-[500px] text-[11px] select-none font-sans overflow-hidden">
            {/* Title Bar */}
            <div className="flex items-center justify-between px-3 py-1.5 bg-gradient-to-r from-gray-100 to-gray-200 border-b border-gray-300">
              <div className="flex items-center gap-1.5 font-semibold text-gray-800 text-xs">
                <span className="text-blue-600 font-bold">A</span>
                <span>Font & Character Spacing (Condensed / Expanded)</span>
              </div>
              <div className="flex items-center gap-1">
                <button
                  type="button"
                  onClick={() => setShowFontDialog(false)}
                  className="w-5 h-5 flex items-center justify-center rounded hover:bg-red-500 hover:text-white text-gray-600 font-bold transition cursor-pointer"
                  title="Close"
                >
                  <X className="w-3.5 h-3.5" />
                </button>
              </div>
            </div>

            {/* Tabs Header */}
            <div className="flex px-2 pt-2 gap-1 border-b border-gray-300 bg-[#E8E8E8]">
              <button
                type="button"
                onClick={() => setFontDialogTab('font')}
                className={`px-4 py-1 rounded-t border-t border-l border-r text-xs font-medium cursor-pointer transition ${
                  fontDialogTab === 'font'
                    ? 'bg-[#F0F0F0] border-gray-400 text-black -mb-[1px] font-semibold'
                    : 'bg-gray-200 border-transparent text-gray-600 hover:text-black hover:bg-gray-100'
                }`}
              >
                Font
              </button>
              <button
                type="button"
                onClick={() => setFontDialogTab('advanced')}
                className={`px-4 py-1 rounded-t border-t border-l border-r text-xs font-medium cursor-pointer transition ${
                  fontDialogTab === 'advanced'
                    ? 'bg-[#F0F0F0] border-gray-400 text-black -mb-[1px] font-semibold'
                    : 'bg-gray-200 border-transparent text-gray-600 hover:text-black hover:bg-gray-100'
                }`}
              >
                Advanced (Condensed)
              </button>
            </div>

            {/* Tab Body */}
            <div className="p-3.5 space-y-3 bg-[#F0F0F0]">
              {fontDialogTab === 'advanced' ? (
                /* Advanced Tab (Character Spacing & OpenType Features) */
                <div className="space-y-3">
                  {/* Character Spacing Fieldset */}
                  <fieldset className="border border-gray-300 rounded p-2.5 bg-[#F7F7F7]">
                    <legend className="text-gray-700 px-1 font-semibold text-[11px]">
                      Character Spacing
                    </legend>
                    <div className="space-y-2 mt-1">
                      {/* Scale */}
                      <div className="grid grid-cols-[100px_1fr] items-center gap-2">
                        <label className="text-gray-700 font-medium">Scale:</label>
                        <select
                          value={characterScale}
                          onChange={(e) => setCharacterScale(e.target.value)}
                          className="bg-white border border-gray-300 rounded px-2 py-0.5 text-xs text-black focus:outline-none focus:border-blue-500 w-36"
                        >
                          <option value="200%">200%</option>
                          <option value="150%">150%</option>
                          <option value="100%">100%</option>
                          <option value="90%">90%</option>
                          <option value="80%">80%</option>
                          <option value="66%">66%</option>
                          <option value="50%">50%</option>
                          <option value="33%">33%</option>
                        </select>
                      </div>

                      {/* Spacing (Condensed / Expanded / Normal) */}
                      <div className="grid grid-cols-[100px_1fr_auto_1fr] items-center gap-2">
                        <label className="text-gray-700 font-medium">Spacing:</label>
                        <select
                          value={characterSpacing}
                          onChange={(e) => setCharacterSpacing(e.target.value as any)}
                          className="bg-white border border-gray-300 rounded px-2 py-0.5 text-xs text-black focus:outline-none focus:border-blue-500 font-bold"
                        >
                          <option value="Normal">Normal</option>
                          <option value="Expanded">Expanded</option>
                          <option value="Condensed">Condensed (ត្បិតអក្សរ)</option>
                        </select>
                        <label className="text-gray-700 ml-2">By:</label>
                        <div className="flex items-center gap-1">
                          <input
                            type="number"
                            step="0.1"
                            min="0"
                            max="20"
                            value={spacingByPt}
                            onChange={(e) => setSpacingByPt(parseFloat(e.target.value) || 0)}
                            className="bg-white border border-gray-300 rounded px-2 py-0.5 text-xs text-black w-20 focus:outline-none focus:border-blue-500 text-right font-medium"
                          />
                          <span className="text-gray-600">pt</span>
                        </div>
                      </div>

                      {/* Position */}
                      <div className="grid grid-cols-[100px_1fr_auto_1fr] items-center gap-2">
                        <label className="text-gray-700 font-medium">Position:</label>
                        <select
                          value={characterPosition}
                          onChange={(e) => setCharacterPosition(e.target.value as any)}
                          className="bg-white border border-gray-300 rounded px-2 py-0.5 text-xs text-black focus:outline-none focus:border-blue-500"
                        >
                          <option value="Normal">Normal</option>
                          <option value="Raised">Raised</option>
                          <option value="Lowered">Lowered</option>
                        </select>
                        <label className="text-gray-700 ml-2">By:</label>
                        <div className="flex items-center gap-1">
                          <input
                            type="number"
                            step="0.1"
                            min="0"
                            max="20"
                            value={positionByPt}
                            onChange={(e) => setPositionByPt(parseFloat(e.target.value) || 0)}
                            className="bg-white border border-gray-300 rounded px-2 py-0.5 text-xs text-black w-20 focus:outline-none focus:border-blue-500 text-right font-medium"
                          />
                          <span className="text-gray-600">pt</span>
                        </div>
                      </div>

                      {/* Kerning */}
                      <div className="flex items-center gap-2 pt-1">
                        <label className="flex items-center gap-1.5 cursor-pointer">
                          <input
                            type="checkbox"
                            checked={kerningEnabled}
                            onChange={(e) => setKerningEnabled(e.target.checked)}
                            className="rounded border-gray-300 text-blue-600"
                          />
                          <span>Kerning for fonts:</span>
                        </label>
                        <input
                          type="number"
                          step="1"
                          min="1"
                          max="72"
                          value={kerningPoints}
                          onChange={(e) => setKerningPoints(parseInt(e.target.value) || 1)}
                          disabled={!kerningEnabled}
                          className="bg-white border border-gray-300 rounded px-1.5 py-0.5 text-xs text-black w-14 text-center disabled:bg-gray-100 disabled:text-gray-400"
                        />
                        <span className="text-gray-600">Points and above</span>
                      </div>
                    </div>
                  </fieldset>

                  {/* OpenType Features Fieldset */}
                  <fieldset className="border border-gray-300 rounded p-2.5 bg-[#F7F7F7]">
                    <legend className="text-gray-700 px-1 font-semibold text-[11px]">
                      OpenType Features
                    </legend>
                    <div className="grid grid-cols-2 gap-x-4 gap-y-2 mt-1">
                      <div className="flex items-center justify-between gap-1">
                        <label className="text-gray-700">Ligatures:</label>
                        <select
                          value={ligatures}
                          onChange={(e) => setLigatures(e.target.value)}
                          className="bg-white border border-gray-300 rounded px-1.5 py-0.5 text-xs text-black w-32"
                        >
                          <option value="Standard and Contextual">Standard and Contextual</option>
                          <option value="None">None</option>
                          <option value="All">All</option>
                          <option value="Historical">Historical</option>
                        </select>
                      </div>

                      <div className="flex items-center justify-between gap-1">
                        <label className="text-gray-700">Number spacing:</label>
                        <select
                          value={numberSpacing}
                          onChange={(e) => setNumberSpacing(e.target.value)}
                          className="bg-white border border-gray-300 rounded px-1.5 py-0.5 text-xs text-black w-28"
                        >
                          <option value="Default">Default</option>
                          <option value="Proportional">Proportional</option>
                          <option value="Tabular">Tabular</option>
                        </select>
                      </div>

                      <div className="flex items-center justify-between gap-1">
                        <label className="text-gray-700">Number forms:</label>
                        <select
                          value={numberForms}
                          onChange={(e) => setNumberForms(e.target.value)}
                          className="bg-white border border-gray-300 rounded px-1.5 py-0.5 text-xs text-black w-32"
                        >
                          <option value="Default">Default</option>
                          <option value="Lining">Lining</option>
                          <option value="Oldstyle">Oldstyle</option>
                        </select>
                      </div>

                      <div className="flex items-center justify-between gap-1">
                        <label className="text-gray-700">Stylistic sets:</label>
                        <select
                          value={stylisticSets}
                          onChange={(e) => setStylisticSets(e.target.value)}
                          className="bg-white border border-gray-300 rounded px-1.5 py-0.5 text-xs text-black w-28"
                        >
                          <option value="Default">Default</option>
                          <option value="1">1</option>
                          <option value="2">2</option>
                          <option value="3">3</option>
                        </select>
                      </div>

                      <div className="col-span-2 pt-0.5">
                        <label className="flex items-center gap-1.5 cursor-pointer">
                          <input
                            type="checkbox"
                            checked={useContextualAlternates}
                            onChange={(e) => setUseContextualAlternates(e.target.checked)}
                            className="rounded border-gray-300 text-blue-600"
                          />
                          <span>Use Contextual Alternates</span>
                        </label>
                      </div>
                    </div>
                  </fieldset>
                </div>
              ) : (
                /* Font Tab */
                <div className="space-y-3">
                  <div className="grid grid-cols-3 gap-2">
                    {/* Font Family */}
                    <div className="space-y-1">
                      <label className="text-gray-700 font-semibold">Font:</label>
                      <select
                        size={5}
                        value={dialogFontFamily}
                        onChange={(e) => setDialogFontFamily(e.target.value)}
                        className="w-full bg-white border border-gray-300 rounded p-1 text-xs text-black focus:outline-none"
                      >
                        <option value="Khmer OS Siemreap">Khmer OS Siemreap</option>
                        <option value="Khmer OS Mool1">Khmer OS Mool1</option>
                        <option value="Khmer OS Battambang">Khmer OS Battambang</option>
                        <option value="Times New Roman">Times New Roman</option>
                        <option value="Arial">Arial</option>
                        <option value="Calibri">Calibri</option>
                      </select>
                    </div>

                    {/* Font Style */}
                    <div className="space-y-1">
                      <label className="text-gray-700 font-semibold">Font style:</label>
                      <select
                        size={5}
                        value={dialogFontStyle}
                        onChange={(e) => setDialogFontStyle(e.target.value)}
                        className="w-full bg-white border border-gray-300 rounded p-1 text-xs text-black focus:outline-none"
                      >
                        <option value="Regular">Regular</option>
                        <option value="Italic">Italic</option>
                        <option value="Bold">Bold</option>
                        <option value="Bold Italic">Bold Italic</option>
                      </select>
                    </div>

                    {/* Size */}
                    <div className="space-y-1">
                      <label className="text-gray-700 font-semibold">Size:</label>
                      <select
                        size={5}
                        value={dialogFontSize}
                        onChange={(e) => setDialogFontSize(e.target.value)}
                        className="w-full bg-white border border-gray-300 rounded p-1 text-xs text-black focus:outline-none"
                      >
                        {FONT_SIZE_STEPS.map((s) => (
                          <option key={s} value={s}>
                            {s}
                          </option>
                        ))}
                      </select>
                    </div>
                  </div>

                  {/* Font Color & Underline */}
                  <div className="grid grid-cols-2 gap-3 pt-1">
                    <div className="flex items-center justify-between">
                      <label className="text-gray-700">Font color:</label>
                      <input
                        type="color"
                        value={dialogFontColor}
                        onChange={(e) => setDialogFontColor(e.target.value)}
                        className="w-10 h-6 border border-gray-300 rounded cursor-pointer p-0.5 bg-white"
                      />
                    </div>
                    <div className="flex items-center justify-between">
                      <label className="text-gray-700">Underline style:</label>
                      <select
                        value={dialogUnderlineStyle}
                        onChange={(e) => setDialogUnderlineStyle(e.target.value)}
                        className="bg-white border border-gray-300 rounded px-2 py-0.5 text-xs text-black"
                      >
                        <option value="none">(none)</option>
                        <option value="single">Single line</option>
                      </select>
                    </div>
                  </div>

                  {/* Effects */}
                  <fieldset className="border border-gray-300 rounded p-2 bg-[#F7F7F7]">
                    <legend className="text-gray-700 px-1 font-semibold text-[11px]">
                      Effects
                    </legend>
                    <div className="grid grid-cols-2 gap-1.5 mt-1">
                      <label className="flex items-center gap-1.5 cursor-pointer">
                        <input
                          type="checkbox"
                          checked={dialogEffects.strikethrough}
                          onChange={(e) =>
                            setDialogEffects({ ...dialogEffects, strikethrough: e.target.checked })
                          }
                          className="rounded border-gray-300 text-blue-600"
                        />
                        <span>Strikethrough</span>
                      </label>
                      <label className="flex items-center gap-1.5 cursor-pointer">
                        <input
                          type="checkbox"
                          checked={dialogEffects.superscript}
                          onChange={(e) =>
                            setDialogEffects({ ...dialogEffects, superscript: e.target.checked })
                          }
                          className="rounded border-gray-300 text-blue-600"
                        />
                        <span>Superscript</span>
                      </label>
                      <label className="flex items-center gap-1.5 cursor-pointer">
                        <input
                          type="checkbox"
                          checked={dialogEffects.subscript}
                          onChange={(e) =>
                            setDialogEffects({ ...dialogEffects, subscript: e.target.checked })
                          }
                          className="rounded border-gray-300 text-blue-600"
                        />
                        <span>Subscript</span>
                      </label>
                      <label className="flex items-center gap-1.5 cursor-pointer">
                        <input
                          type="checkbox"
                          checked={dialogEffects.allCaps}
                          onChange={(e) =>
                            setDialogEffects({ ...dialogEffects, allCaps: e.target.checked })
                          }
                          className="rounded border-gray-300 text-blue-600"
                        />
                        <span>All caps</span>
                      </label>
                    </div>
                  </fieldset>
                </div>
              )}

              {/* Preview Box */}
              <fieldset className="border border-gray-300 rounded p-2.5 bg-[#F7F7F7]">
                <legend className="text-gray-700 px-1 font-semibold text-[11px]">
                  Preview
                </legend>
                <div className="h-16 bg-white border border-gray-400 p-2 flex items-center justify-center overflow-hidden rounded-sm shadow-inner">
                  <span
                    style={{
                      fontFamily: dialogFontFamily.includes('Times')
                        ? 'Times New Roman'
                        : dialogFontFamily.includes('Mool')
                        ? `'Khmer OS Muol Light', serif`
                        : `'Khmer OS Siemreap', 'Siemreap', sans-serif`,
                      fontSize: dialogFontSize,
                      fontWeight: dialogFontStyle.includes('Bold') ? 'bold' : 'normal',
                      fontStyle: dialogFontStyle.includes('Italic') ? 'italic' : 'normal',
                      color: dialogFontColor,
                      textDecoration: dialogUnderlineStyle === 'single' ? 'underline' : dialogEffects.strikethrough ? 'line-through' : 'none',
                      letterSpacing:
                        characterSpacing === 'Condensed'
                          ? `-${spacingByPt}pt`
                          : characterSpacing === 'Expanded'
                          ? `${spacingByPt}pt`
                          : 'normal',
                      transform:
                        characterScale !== '100%'
                          ? `scaleX(${parseFloat(characterScale.replace('%', '')) / 100})`
                          : undefined,
                      display: 'inline-block',
                      transformOrigin: 'center',
                      position:
                        characterPosition === 'Raised' || characterPosition === 'Lowered'
                          ? 'relative'
                          : undefined,
                      top:
                        characterPosition === 'Raised'
                          ? `-${positionByPt}pt`
                          : characterPosition === 'Lowered'
                          ? `${positionByPt}pt`
                          : undefined,
                    }}
                    className="text-center transition-all"
                  >
                    Sample របាយការណ៍ស្តុកតាមក្រុម 123456789
                  </span>
                </div>
                <p className="text-[10px] text-gray-500 mt-1">
                  This is the body theme font. The current document theme defines which font will be used.
                </p>
              </fieldset>
            </div>

            {/* Bottom Actions */}
            <div className="flex items-center justify-between px-3.5 py-2.5 bg-[#E8E8E8] border-t border-gray-300">
              <div className="flex items-center gap-1.5">
                <button
                  type="button"
                  onClick={() => alert('Default character spacing saved.')}
                  className="px-2.5 py-1 bg-white hover:bg-gray-50 text-gray-800 rounded border border-gray-300 font-medium text-[11px] shadow-sm cursor-pointer transition"
                >
                  Set As Default
                </button>
              </div>

              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={handleApplyFontDialog}
                  className="px-5 py-1 bg-blue-600 hover:bg-blue-700 text-white rounded border border-blue-700 font-medium text-xs shadow-sm cursor-pointer transition"
                >
                  OK
                </button>
                <button
                  type="button"
                  onClick={() => setShowFontDialog(false)}
                  className="px-4 py-1 bg-white hover:bg-gray-100 text-gray-800 rounded border border-gray-400 font-medium text-xs shadow-sm cursor-pointer transition"
                >
                  Cancel
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Microsoft Word Page Setup Dialog Modal (Alt+P+S+P) */}
      {showPageSetupDialog && (
        <div className="fixed inset-0 z-50 bg-black/65 backdrop-blur-[2px] flex items-center justify-center p-3 animate-fadeIn">
          <div className="bg-[#F0F0F0] text-gray-900 rounded shadow-2xl border border-gray-400 w-full max-w-[530px] text-[11px] select-none font-sans overflow-hidden">
            {/* Window Header */}
            <div className="flex items-center justify-between bg-gradient-to-r from-blue-800 to-indigo-900 text-white px-3 py-1.5 font-semibold text-xs border-b border-blue-900">
              <span className="flex items-center gap-1.5">
                <Layout className="w-3.5 h-3.5 text-blue-300" />
                <span>Page Setup (កំណត់ទំព័រ និងគែមក្រដាស A4)</span>
              </span>
              <button
                type="button"
                onClick={() => setShowPageSetupDialog(false)}
                className="hover:bg-red-600 text-white p-0.5 rounded transition cursor-pointer"
                title="បិទ"
              >
                <X className="w-3.5 h-3.5" />
              </button>
            </div>

            {/* Tab Navigation */}
            <div className="flex items-center bg-[#E5E5E5] px-3 pt-2 border-b border-gray-300 gap-1 text-xs">
              <button
                type="button"
                onClick={() => setPageSetupTab('margins')}
                className={`px-4 py-1 rounded-t border-t border-x font-medium cursor-pointer transition ${
                  pageSetupTab === 'margins'
                    ? 'bg-[#F0F0F0] border-gray-300 text-black -mb-[1px] font-semibold pb-1.5 shadow-sm'
                    : 'bg-transparent border-transparent text-gray-600 hover:text-black'
                }`}
              >
                Margins (គែមក្រដាស)
              </button>
              <button
                type="button"
                onClick={() => setPageSetupTab('paper')}
                className={`px-4 py-1 rounded-t border-t border-x font-medium cursor-pointer transition ${
                  pageSetupTab === 'paper'
                    ? 'bg-[#F0F0F0] border-gray-300 text-black -mb-[1px] font-semibold pb-1.5 shadow-sm'
                    : 'bg-transparent border-transparent text-gray-600 hover:text-black'
                }`}
              >
                Paper (ទំហំក្រដាស)
              </button>
              <button
                type="button"
                onClick={() => setPageSetupTab('layout')}
                className={`px-4 py-1 rounded-t border-t border-x font-medium cursor-pointer transition ${
                  pageSetupTab === 'layout'
                    ? 'bg-[#F0F0F0] border-gray-300 text-black -mb-[1px] font-semibold pb-1.5 shadow-sm'
                    : 'bg-transparent border-transparent text-gray-600 hover:text-black'
                }`}
              >
                Layout (ប្លង់ & គម្លាត)
              </button>
            </div>

            {/* Tab Content Body */}
            <div className="p-3.5 space-y-3 bg-[#F0F0F0]">
              {pageSetupTab === 'margins' && (
                <div className="space-y-3">
                  {/* Margins Inputs Box */}
                  <fieldset className="border border-gray-300 rounded p-2.5 bg-[#F7F7F7]">
                    <legend className="text-gray-700 px-1 font-semibold text-[11px]">
                      Margins (គែមក្រដាសគិតជា សង់ទីម៉ែត្រ cm)
                    </legend>
                    <div className="grid grid-cols-2 gap-x-4 gap-y-2 text-xs">
                      <div className="flex items-center justify-between">
                        <label className="text-gray-700">Top (លើ) :</label>
                        <div className="flex items-center">
                          <input
                            type="number"
                            step="0.1"
                            min="0.2"
                            max="5.0"
                            value={tempMargins.top}
                            onChange={(e) => setTempMargins({ ...tempMargins, top: parseFloat(e.target.value) || 0 })}
                            className="w-16 bg-white border border-gray-300 rounded px-1.5 py-0.5 text-right font-mono text-black"
                          />
                          <span className="ml-1 text-gray-500 text-[10px]">cm</span>
                        </div>
                      </div>

                      <div className="flex items-center justify-between">
                        <label className="text-gray-700">Bottom (ក្រោម) :</label>
                        <div className="flex items-center">
                          <input
                            type="number"
                            step="0.1"
                            min="0.2"
                            max="5.0"
                            value={tempMargins.bottom}
                            onChange={(e) => setTempMargins({ ...tempMargins, bottom: parseFloat(e.target.value) || 0 })}
                            className="w-16 bg-white border border-gray-300 rounded px-1.5 py-0.5 text-right font-mono text-black"
                          />
                          <span className="ml-1 text-gray-500 text-[10px]">cm</span>
                        </div>
                      </div>

                      <div className="flex items-center justify-between">
                        <label className="text-gray-900 font-bold">Left (ឆ្វេង) :</label>
                        <div className="flex items-center">
                          <input
                            type="number"
                            step="0.1"
                            min="0.2"
                            max="6.0"
                            value={tempMargins.left}
                            onChange={(e) => setTempMargins({ ...tempMargins, left: parseFloat(e.target.value) || 0 })}
                            className="w-16 bg-blue-50 border border-blue-400 font-bold rounded px-1.5 py-0.5 text-right font-mono text-blue-900"
                          />
                          <span className="ml-1 text-gray-500 text-[10px]">cm</span>
                        </div>
                      </div>

                      <div className="flex items-center justify-between">
                        <label className="text-gray-700">Right (ស្តាំ) :</label>
                        <div className="flex items-center">
                          <input
                            type="number"
                            step="0.1"
                            min="0.2"
                            max="5.0"
                            value={tempMargins.right}
                            onChange={(e) => setTempMargins({ ...tempMargins, right: parseFloat(e.target.value) || 0 })}
                            className="w-16 bg-white border border-gray-300 rounded px-1.5 py-0.5 text-right font-mono text-black"
                          />
                          <span className="ml-1 text-gray-500 text-[10px]">cm</span>
                        </div>
                      </div>

                      <div className="flex items-center justify-between pt-1">
                        <label className="text-gray-700">Gutter :</label>
                        <div className="flex items-center">
                          <input
                            type="number"
                            step="0.1"
                            min="0"
                            max="3"
                            value={gutter}
                            onChange={(e) => setGutter(parseFloat(e.target.value) || 0)}
                            className="w-16 bg-white border border-gray-300 rounded px-1.5 py-0.5 text-right font-mono text-black"
                          />
                          <span className="ml-1 text-gray-500 text-[10px]">cm</span>
                        </div>
                      </div>

                      <div className="flex items-center justify-between pt-1">
                        <label className="text-gray-700">Gutter pos :</label>
                        <select
                          value={gutterPosition}
                          onChange={(e) => setGutterPosition(e.target.value as 'left' | 'top')}
                          className="w-20 bg-white border border-gray-300 rounded px-1 py-0.5 text-black text-[11px]"
                        >
                          <option value="left">Left</option>
                          <option value="top">Top</option>
                        </select>
                      </div>
                    </div>

                    {/* Quick Presets Bar */}
                    <div className="mt-2.5 pt-2 border-t border-gray-200">
                      <p className="text-[10px] text-gray-500 mb-1 font-medium">ទម្រង់គែមរហ័ស (Quick Margin Presets) :</p>
                      <div className="flex flex-wrap gap-1">
                        <button
                          type="button"
                          onClick={() => setTempMargins({ top: 0.4, bottom: 0.4, left: 1.8, right: 1.0 })}
                          className="px-2 py-0.5 bg-blue-100 hover:bg-blue-200 text-blue-900 border border-blue-300 rounded text-[10.5px] font-semibold cursor-pointer"
                        >
                          🇰🇭 ស្តង់ដារតារាង (1.8 / 1.0 / 0.4 / 0.4)
                        </button>
                        <button
                          type="button"
                          onClick={() => setTempMargins({ top: 1.2, bottom: 1.2, left: 2.5, right: 1.5 })}
                          className="px-2 py-0.5 bg-gray-200 hover:bg-gray-300 text-gray-800 border border-gray-300 rounded text-[10.5px] cursor-pointer"
                        >
                          ស្តង់ដាររដ្ឋបាល (2.5 / 1.5 / 1.2)
                        </button>
                        <button
                          type="button"
                          onClick={() => setTempMargins({ top: 1.0, bottom: 1.0, left: 2.0, right: 1.2 })}
                          className="px-2 py-0.5 bg-gray-200 hover:bg-gray-300 text-gray-800 border border-gray-300 rounded text-[10.5px] cursor-pointer"
                        >
                          គែមត្បិត (ឆ្វេង 2.0cm)
                        </button>
                      </div>
                    </div>
                  </fieldset>

                  {/* Orientation Section */}
                  <fieldset className="border border-gray-300 rounded p-2.5 bg-[#F7F7F7]">
                    <legend className="text-gray-700 px-1 font-semibold text-[11px]">
                      Orientation (ទិសដៅទំព័រ)
                    </legend>
                    <div className="flex items-center gap-6 pt-0.5">
                      <label
                        onClick={() => setPageOrientation('portrait')}
                        className={`flex items-center gap-2 p-2 rounded border cursor-pointer transition ${
                          pageOrientation === 'portrait'
                            ? 'bg-blue-50 border-blue-500 text-blue-900 font-semibold ring-1 ring-blue-400'
                            : 'bg-white border-gray-300 text-gray-700 hover:bg-gray-50'
                        }`}
                      >
                        <input
                          type="radio"
                          name="orientation"
                          checked={pageOrientation === 'portrait'}
                          onChange={() => setPageOrientation('portrait')}
                          className="text-blue-600"
                        />
                        <div className="w-4 h-6 border-2 border-gray-700 bg-white rounded-xs"></div>
                        <span>Portrait (បញ្ឈរ)</span>
                      </label>

                      <label
                        onClick={() => setPageOrientation('landscape')}
                        className={`flex items-center gap-2 p-2 rounded border cursor-pointer transition ${
                          pageOrientation === 'landscape'
                            ? 'bg-blue-50 border-blue-500 text-blue-900 font-semibold ring-1 ring-blue-400'
                            : 'bg-white border-gray-300 text-gray-700 hover:bg-gray-50'
                        }`}
                      >
                        <input
                          type="radio"
                          name="orientation"
                          checked={pageOrientation === 'landscape'}
                          onChange={() => setPageOrientation('landscape')}
                          className="text-blue-600"
                        />
                        <div className="w-6 h-4 border-2 border-gray-700 bg-white rounded-xs"></div>
                        <span>Landscape (ផ្ដេក)</span>
                      </label>
                    </div>
                  </fieldset>
                </div>
              )}

              {pageSetupTab === 'paper' && (
                <div className="space-y-3">
                  <fieldset className="border border-gray-300 rounded p-2.5 bg-[#F7F7F7]">
                    <legend className="text-gray-700 px-1 font-semibold text-[11px]">
                      Paper Size (ទំហំក្រដាស)
                    </legend>
                    <div className="space-y-2.5">
                      <div className="flex items-center justify-between">
                        <label className="text-gray-700 font-semibold">Paper size :</label>
                        <select
                          value={paperSize}
                          onChange={(e) => setPaperSize(e.target.value as any)}
                          className="w-48 bg-white border border-gray-300 rounded px-2 py-1 text-black font-medium"
                        >
                          <option value="A4">A4 (210 x 297 mm)</option>
                          <option value="Letter">Letter (8.5 x 11 in)</option>
                          <option value="Legal">Legal (8.5 x 14 in)</option>
                          <option value="A3">A3 (297 x 420 mm)</option>
                        </select>
                      </div>

                      <div className="grid grid-cols-2 gap-3 pt-1">
                        <div className="flex items-center justify-between">
                          <label className="text-gray-700">Width :</label>
                          <div className="flex items-center">
                            <span className="font-mono bg-gray-100 border border-gray-300 px-2 py-0.5 rounded text-gray-800">
                              {paperSize === 'A4' ? '21.0' : paperSize === 'Letter' ? '21.59' : paperSize === 'Legal' ? '21.59' : '29.7'}
                            </span>
                            <span className="ml-1 text-gray-500 text-[10px]">cm</span>
                          </div>
                        </div>

                        <div className="flex items-center justify-between">
                          <label className="text-gray-700">Height :</label>
                          <div className="flex items-center">
                            <span className="font-mono bg-gray-100 border border-gray-300 px-2 py-0.5 rounded text-gray-800">
                              {paperSize === 'A4' ? '29.7' : paperSize === 'Letter' ? '27.94' : paperSize === 'Legal' ? '35.56' : '42.0'}
                            </span>
                            <span className="ml-1 text-gray-500 text-[10px]">cm</span>
                          </div>
                        </div>
                      </div>
                    </div>
                  </fieldset>
                </div>
              )}

              {pageSetupTab === 'layout' && (
                <div className="space-y-3">
                  <fieldset className="border border-gray-300 rounded p-2.5 bg-[#F7F7F7]">
                    <legend className="text-gray-700 px-1 font-semibold text-[11px]">
                      Headers and Footers (ក្បាល និងបាតទំព័រ)
                    </legend>
                    <div className="grid grid-cols-2 gap-3 text-xs">
                      <div className="flex items-center justify-between">
                        <label className="text-gray-700">From edge Header :</label>
                        <div className="flex items-center">
                          <input
                            type="number"
                            step="0.1"
                            value={headerMargin}
                            onChange={(e) => setHeaderMargin(parseFloat(e.target.value) || 0)}
                            className="w-16 bg-white border border-gray-300 rounded px-1.5 py-0.5 text-right font-mono text-black"
                          />
                          <span className="ml-1 text-gray-500 text-[10px]">cm</span>
                        </div>
                      </div>

                      <div className="flex items-center justify-between">
                        <label className="text-gray-700">From edge Footer :</label>
                        <div className="flex items-center">
                          <input
                            type="number"
                            step="0.1"
                            value={footerMargin}
                            onChange={(e) => setFooterMargin(parseFloat(e.target.value) || 0)}
                            className="w-16 bg-white border border-gray-300 rounded px-1.5 py-0.5 text-right font-mono text-black"
                          />
                          <span className="ml-1 text-gray-500 text-[10px]">cm</span>
                        </div>
                      </div>
                    </div>
                  </fieldset>

                  <fieldset className="border border-gray-300 rounded p-2.5 bg-[#F7F7F7]">
                    <legend className="text-gray-700 px-1 font-semibold text-[11px]">
                      Paragraph Indent & Tab Stop (ការចូលបន្ទាត់)
                    </legend>
                    <div className="flex items-center justify-between">
                      <label className="text-gray-700">First-line Tab Stop :</label>
                      <div className="flex items-center gap-1.5">
                        <input
                          type="number"
                          step="0.1"
                          min="0"
                          max="4"
                          value={activeTabStop}
                          onChange={(e) => setActiveTabStop(parseFloat(e.target.value) || 0)}
                          className="w-16 bg-white border border-gray-300 rounded px-1.5 py-0.5 text-right font-mono text-black"
                        />
                        <span className="text-gray-500 text-[10px]">cm</span>
                        <button
                          type="button"
                          onClick={() => setActiveTabStop(2.0)}
                          className="px-2 py-0.5 bg-gray-200 hover:bg-gray-300 text-gray-800 rounded text-[10px] cursor-pointer"
                        >
                          2.0cm
                        </button>
                      </div>
                    </div>
                  </fieldset>
                </div>
              )}

              {/* Live Preview Box */}
              <fieldset className="border border-gray-300 rounded p-2.5 bg-[#F7F7F7]">
                <legend className="text-gray-700 px-1 font-semibold text-[11px]">
                  Preview (ទិដ្ឋភាពជាក់ស្តែងនៃទំព័រ)
                </legend>
                <div className="flex items-center justify-center p-2 bg-white border border-gray-300 rounded shadow-inner">
                  <div
                    style={{
                      width: pageOrientation === 'portrait' ? '120px' : '170px',
                      height: pageOrientation === 'portrait' ? '170px' : '120px',
                      paddingTop: `${Math.max(2, (tempMargins.top / 29.7) * 170)}px`,
                      paddingBottom: `${Math.max(2, (tempMargins.bottom / 29.7) * 170)}px`,
                      paddingLeft: `${Math.max(2, (tempMargins.left / 21.0) * 120)}px`,
                      paddingRight: `${Math.max(2, (tempMargins.right / 21.0) * 120)}px`,
                    }}
                    className="border-2 border-gray-700 bg-white shadow flex flex-col justify-between transition-all relative overflow-hidden"
                  >
                    <div className="space-y-0.5 border border-dashed border-blue-300 p-0.5 bg-blue-50/50 rounded-xs h-full flex flex-col justify-between">
                      <div className="flex justify-between items-center">
                        <div className="w-1/3 h-1 bg-gray-400 rounded-xs"></div>
                        <div className="w-1/3 h-1 bg-gray-400 rounded-xs"></div>
                      </div>
                      <div className="space-y-0.5 py-1">
                        <div className="w-full h-0.5 bg-gray-300 rounded-xs"></div>
                        <div className="w-full h-0.5 bg-gray-300 rounded-xs"></div>
                        <div className="w-3/4 h-0.5 bg-gray-300 rounded-xs"></div>
                        <div className="w-full h-0.5 bg-gray-300 rounded-xs"></div>
                      </div>
                      <div className="flex justify-between items-center pt-1">
                        <div className="w-1/4 h-1 bg-gray-400 rounded-xs"></div>
                        <div className="w-1/4 h-1 bg-gray-400 rounded-xs"></div>
                      </div>
                    </div>
                  </div>
                </div>

                <div className="flex items-center justify-between pt-2 text-[10.5px]">
                  <span className="text-gray-600">Apply to :</span>
                  <select className="bg-white border border-gray-300 rounded px-2 py-0.5 text-black">
                    <option>Whole document (ឯកសារទាំងមូល)</option>
                    <option>This point forward</option>
                  </select>
                </div>
              </fieldset>
            </div>

            {/* Bottom Actions */}
            <div className="flex items-center justify-between px-3.5 py-2.5 bg-[#E8E8E8] border-t border-gray-300">
              <button
                type="button"
                onClick={() => {
                  setCustomMargins({ ...tempMargins });
                  alert('បានកំណត់គែមក្រដាសជាលំនាំដើម (Default Margins Saved)!');
                }}
                className="px-2.5 py-1 bg-white hover:bg-gray-50 text-gray-800 rounded border border-gray-300 font-medium text-[11px] shadow-sm cursor-pointer transition"
              >
                Set As Default
              </button>

              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => {
                    setCustomMargins({ ...tempMargins });
                    setShowPageSetupDialog(false);
                  }}
                  className="px-5 py-1 bg-blue-600 hover:bg-blue-700 text-white rounded border border-blue-700 font-semibold text-xs shadow-sm cursor-pointer transition"
                >
                  OK (យល់ព្រម)
                </button>
                <button
                  type="button"
                  onClick={() => setShowPageSetupDialog(false)}
                  className="px-4 py-1 bg-white hover:bg-gray-100 text-gray-800 rounded border border-gray-400 font-medium text-xs shadow-sm cursor-pointer transition"
                >
                  Cancel (បោះបង់)
                </button>
              </div>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
