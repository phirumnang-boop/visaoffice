import React, { useState, useMemo, useRef, useEffect } from 'react';
import * as XLSX from 'xlsx';
import { CategoriesState, Officer, StockRecord, UserRole } from '../types';
import {
  Printer,
  Download,
  Calendar,
  FileSpreadsheet,
  RotateCcw,
  Edit3,
  Check,
  FileText,
  Save,
  Clock,
  HardDrive,
  Bold,
  Italic,
  Underline,
  Strikethrough,
  Subscript,
  Superscript,
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
  List,
  ListOrdered,
  Indent,
  Outdent,
  Search,
  X,
  Minus,
  Plus,
  Layout,
  Move,
  RefreshCw,
  AlertCircle,
  Users,
  Building2,
  ChevronUp,
  ChevronDown,
  Eye,
  EyeOff,
} from 'lucide-react';
import html2canvas from 'html2canvas';
import { jsPDF } from 'jspdf';
import { getKhmerLunarDate, getKhmerSolarParts, parseDateInput, toKhmerNum } from '../utils/khmerCalendar';
import { sanitizeDocumentForHtml2Canvas, exportElementToPdf } from '../utils/pdfExportHelper';
import { printA4Document } from '../utils/printHelper';
import { CustomDatePicker } from './CustomDatePicker';
import {
  VISA_TYPES,
  normalizeDateToISO,
  normalizeVisaType,
  normalizeTeamName,
  OFFICIAL_29_TEAMS,
  matchTeamInList,
  formatReportTeamName,
} from '../utils/teamNormalization';
import {
  isCeaRecord,
  DEFAULT_OPENING_MATRIX,
  isOldStockTeamRecord,
  resolveRecordTeamName,
} from '../utils/teamStockCalculation';
import { TacteingLine, TacteingControlSelector, saveTacteingSettings, getSavedTacteingSettings, TacteingType } from './TacteingLine';
import { exportRobokToWord } from '../utils/exportRobokWord';
import { idbStorage } from '../utils/idbStorage';

export interface RobokTotalStockWorkReportProps {
  stockRecords: StockRecord[];
  categories?: CategoriesState;
  officers?: Officer[];
  currentRole?: UserRole;
  userName?: string;
  assignedTeam?: string;
  onClose?: () => void;
  variant?: 'standard' | 'director';
  initialViewMode?: 'team' | 'office';
  isRolling?: boolean;
  defaultDocThroughTitle?: string;
  proposalType?: 'section' | 'office';
}

export interface StubProposalRow {
  visaType: string;
  booklets: number;
  sheets: number;
  remarks?: string;
}

export const DEFAULT_STUB_PROPOSAL_ROWS: StubProposalRow[] = [
  { visaType: 'T', booklets: 6491, sheets: 324550, remarks: '' },
  { visaType: 'T1', booklets: 16, sheets: 800, remarks: '' },
  { visaType: 'E', booklets: 1724, sheets: 86200, remarks: '' },
  { visaType: 'E1', booklets: 3, sheets: 150, remarks: '' },
  { visaType: 'D', booklets: 492, sheets: 24600, remarks: '' },
  { visaType: 'K', booklets: 7, sheets: 350, remarks: '' },
  { visaType: 'C', booklets: 15, sheets: 750, remarks: '' },
];

export const ALL_13_VISA_TYPES = ['T', 'T1', 'T2', 'T3', 'E', 'E1', 'E2', 'E3', 'D', 'K', 'A', 'B', 'C'] as const;

export const DEFAULT_STUB_EMPTY_ROWS: StubProposalRow[] = ALL_13_VISA_TYPES.map((vt) => ({
  visaType: vt,
  booklets: 0,
  sheets: 0,
  remarks: '',
}));

// Official Baseline for "សន្និធិ ចុងគ្រា ៣០-វិច្ឆិកា-២០១៨" (30-Nov-2018 benchmark figures)
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

export const OFFICIAL_TEAMS_DEC_2018_BASELINE: Record<string, number> = {
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

const KHMER_MONTHS_NAMES = [
  'មករា', 'កុម្ភៈ', 'មីនា', 'មេសា', 'ឧសភា', 'មិថុនា',
  'កក្កដា', 'សីហា', 'កញ្ញា', 'តុលា', 'វិច្ឆិកា', 'ធ្នូ'
];

// Helper to sanitize saved report HTML
const sanitizeReportHtml = (html: string) => {
  return html
    .replace(/<tr[^>]*>\s*(?:<t[dh][^>]*>\s*\(?\s*ដុំ\s*\)?\s*<\/t[dh]>\s*)+<\/tr>/gi, '')
    .replace(/\s*\(\s*ដុំ\s*\)/g, '')
    .replace(/（ដុំ）/g, '')
    .replace(/transform:\s*translateX\(3(?:\.0)?cm\);?/gi, '')
    .replace(/margin-left:\s*3(?:\.0)?cm;?/gi, '')
    .replace(/translateX\(3(?:\.0)?cm\)/gi, '')
    // Auto-update department name in saved HTML
    .replace(/នាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន៍/g, 'នាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត')
    // Auto-update concluding salutation in saved HTML
    .replace(/ជ្រាបរបាយការណ៍/g, 'ជ្រាបជារបាយការណ៍')
    .replace(/និងជ្រាបរបាយការណ៍/g, 'និងជ្រាបជារបាយការណ៍');
};

export const RobokTotalStockWorkReport: React.FC<RobokTotalStockWorkReportProps> = ({
  stockRecords,
  categories,
  officers = [],
  currentRole,
  userName,
  assignedTeam,
  onClose,
  initialViewMode,
  isRolling = false,
  defaultDocThroughTitle,
  proposalType,
}) => {
  const storagePrefix = isRolling
    ? 'robok_rolling_report_saved_'
    : proposalType === 'section'
    ? 'robok_stub_section_saved_'
    : proposalType === 'office'
    ? 'robok_stub_office_saved_'
    : 'robok_report_saved_';
  const latestDraftKey = isRolling
    ? 'robok_rolling_report_latest_draft'
    : proposalType === 'section'
    ? 'robok_stub_section_latest_draft'
    : proposalType === 'office'
    ? 'robok_stub_office_latest_draft'
    : 'robok_report_latest_draft';
  const userPrefKey = isRolling
    ? 'robok_rolling_report_user_preferences'
    : proposalType === 'section'
    ? 'robok_stub_section_user_preferences'
    : proposalType === 'office'
    ? 'robok_stub_office_user_preferences'
    : 'robok_report_user_preferences';

  // Initial draft load helper to ensure saved report is maintained when navigating back
  const initialSavedDraft = useMemo(() => {
    try {
      const processDraft = (draft: any) => {
        if (!draft) return draft;
        if (draft.docRecipientTitle) {
          draft.docRecipientTitle = draft.docRecipientTitle.replace(/នាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន៍/g, 'នាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត');
        }
        if (proposalType === 'office') {
          if (!draft.docRecipientRank || draft.docRecipientRank.includes('ឯកឧត្តម')) {
            draft.docRecipientRank = 'លោកឧត្តមសេនីយ៍ទោ';
          }
          if (!draft.docRecipientTitle || draft.docRecipientTitle.includes('នាយការិយាល័យ')) {
            draft.docRecipientTitle = 'ប្រធាននាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត';
          }
        }
        if (Array.isArray(draft.ministryHierarchy)) {
          draft.ministryHierarchy = draft.ministryHierarchy.map((line: string) =>
            line.replace(/នាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន៍/g, 'នាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត')
          );
          if (proposalType === 'office') {
            draft.ministryHierarchy = draft.ministryHierarchy.filter((l: string) => !l.includes('ផ្នែករដ្ឋបាល'));
          }
        }
        return draft;
      };

      const defaultPeriodKey = isRolling ? '2018-12-16_2019-06-15' : '2018-12-01_2019-02-28';
      const defaultPeriod = localStorage.getItem(`${storagePrefix}${defaultPeriodKey}`);
      if (defaultPeriod) return processDraft(JSON.parse(defaultPeriod));
    } catch (e) {
      console.error(e);
    }
    return null;
  }, [storagePrefix, isRolling]);

  // Date range filters
  const [selectedYear, setSelectedYear] = useState<number>(() => initialSavedDraft?.selectedYear !== undefined ? initialSavedDraft.selectedYear : 2018);
  const [selectedMonth, setSelectedMonth] = useState<number>(() => initialSavedDraft?.selectedMonth !== undefined ? initialSavedDraft.selectedMonth : 12); // 1-12
  const [useCustomRange, setUseCustomRange] = useState<boolean>(() => initialSavedDraft?.useCustomRange !== undefined ? initialSavedDraft.useCustomRange : true);
  const [startDate, setStartDate] = useState<string>(() => {
    if (initialSavedDraft?.startDate) {
      if (isRolling && !initialSavedDraft.startDate.endsWith('-16')) {
        const parts = initialSavedDraft.startDate.split('-');
        if (parts.length === 3) return `${parts[0]}-${parts[1]}-16`;
      }
      return initialSavedDraft.startDate;
    }
    return isRolling ? '2018-12-16' : '2018-12-01';
  });
  const [endDate, setEndDate] = useState<string>(() => {
    if (initialSavedDraft?.endDate) {
      if (isRolling && !initialSavedDraft.endDate.endsWith('-15')) {
        const parts = initialSavedDraft.endDate.split('-');
        if (parts.length === 3) return `${parts[0]}-${parts[1]}-15`;
      }
      return initialSavedDraft.endDate;
    }
    return isRolling ? '2019-06-15' : '2019-02-28';
  });

  const [activeDurationPreset, setActiveDurationPreset] = useState<'this_month' | 'q1' | 's1' | 'm9' | 'full_year' | null>(() => isRolling ? 's1' : 'q1');
  const [isCalculated, setIsCalculated] = useState<boolean>(false);

  const cleanDocThroughTitle = (title?: string) => {
    if (!title) return '';
    return title.replace(/លោកនាយរងការិយាល័យទិទទួលបន្ទុក/g, 'លោកនាយរងការិយាល័យទទួលបន្ទុក')
                .replace(/ទិទទួលបន្ទុក/g, 'ទទួលបន្ទុក');
  };

  // Custom metadata editing
  const isStubProposal = proposalType === 'section' || proposalType === 'office';
  const defaultDocRecipientTitle = proposalType === 'office'
    ? 'ប្រធាននាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត'
    : 'លោកវរសេនីយ៍ទោ នាយការិយាល័យទិដ្ឋាការចូល';
  const defaultDocRecipientRank = proposalType === 'office' ? 'លោកឧត្តមសេនីយ៍ទោ' : '';

  const [docNumber, setDocNumber] = useState<string>(() => {
    if (proposalType === 'office') {
      if (initialSavedDraft?.docNumber && !initialSavedDraft.docNumber.includes('.....')) {
        return initialSavedDraft.docNumber;
      }
      return '.......................សណ/២៦';
    }
    return initialSavedDraft?.docNumber || '.......................';
  });
  const [docRecipientRank, setDocRecipientRank] = useState<string>(() => initialSavedDraft?.docRecipientRank || defaultDocRecipientRank);
  const [docRecipientTitle, setDocRecipientTitle] = useState<string>(() => initialSavedDraft?.docRecipientTitle || defaultDocRecipientTitle);
  const fallbackThroughTitle = defaultDocThroughTitle || (proposalType === 'section' ? 'លោកនាយរងការិយាល័យទទួលបន្ទុក' : proposalType === 'office' ? 'លោកអនុប្រធាននាយកដ្ឋានទទួលបន្ទុក' : 'លោកនាយការិយាល័យទិដ្ឋាការចូល');
  const [docThroughTitle, setDocThroughTitle] = useState<string>(() => cleanDocThroughTitle(initialSavedDraft?.docThroughTitle || fallbackThroughTitle));
  const [docReferenceText, setDocReferenceText] = useState<string>(() => initialSavedDraft?.docReferenceText || 'លិខិតប្រគល់ទទួលតាមក្រុមផ្តល់ទិដ្ឋាការ ។');
  const [docSignDate, setDocSignDate] = useState<string>(() => initialSavedDraft?.docSignDate || new Date().toISOString().split('T')[0]);

  // Result count from ប្រតិបត្តិការ (ប្រមូលគល់សន្លឹក) in ទិន្នន័យសន្លឹកទិដ្ឋាការ ក្នុងការងារស្តុកការិយាល័យ
  const stubRowsFromStock = useMemo(() => {
    if (!isCalculated && isStubProposal) {
      return [];
    }

    const isStubRec = (r: StockRecord) => {
      if (r.stockType && r.stockType !== 'sticker') return false;
      return (
        r.operationType === 'returnStub' ||
        Boolean(r.sourceFrom?.includes('គល់សន្លឹក')) ||
        Boolean(r.remarks?.includes('គល់សន្លឹក')) ||
        Boolean(r.id?.startsWith('stock-stub-'))
      );
    };

    const normStart = normalizeDateToISO(startDate);
    const normEnd = normalizeDateToISO(endDate);

    // Filter strictly by report date range
    const matchingRecords = (stockRecords || []).filter((r) => {
      if (!isStubRec(r)) return false;
      const d = normalizeDateToISO(r.date || '');
      if (normStart && d && d < normStart) return false;
      if (normEnd && d && d > normEnd) return false;
      return true;
    });

    const map: Record<string, { sheets: number; remarks?: string }> = {};

    matchingRecords.forEach((r) => {
      const vt = normalizeVisaType(r.visaType);
      const q = Number(r.totalSheets || r.quantityBundles || (r as any).quantity || 0);
      if (q > 0) {
        if (!map[vt]) map[vt] = { sheets: 0, remarks: '' };
        map[vt].sheets += q;
      }
    });

    // Also check localStorage app_visa_stub_collections_v1 if present and not already in stockRecords
    try {
      const savedStubs = localStorage.getItem('app_visa_stub_collections_v1');
      if (savedStubs) {
        const parsed = JSON.parse(savedStubs);
        if (Array.isArray(parsed)) {
          parsed.forEach((item: any) => {
            const d = normalizeDateToISO(item.date || '');
            const inRange = (!normStart || !d || d >= normStart) && (!normEnd || !d || d <= normEnd);
            const alreadyInStock = (stockRecords || []).some(
              (sr) => sr.id.startsWith(`stock-stub-${item.id}`) || (sr as any).stubId === item.id
            );
            if (!alreadyInStock && inRange) {
              if (item.values && typeof item.values === 'object') {
                Object.entries(item.values).forEach(([vt, val]: [string, any]) => {
                  const q = Number(val?.quantity || 0);
                  if (q > 0) {
                    const normVt = normalizeVisaType(vt);
                    if (!map[normVt]) map[normVt] = { sheets: 0, remarks: '' };
                    map[normVt].sheets += q;
                  }
                });
              }
            }
          });
        }
      }
    } catch (e) {
      console.error(e);
    }

    // Only visa types that have records / data > 0 are displayed ("ហើយប្រភេទ ណាមានចាំបង្ហាញ")
    const visaOrder = ['T', 'T1', 'T2', 'T3', 'E', 'E1', 'E2', 'E3', 'D', 'K', 'A', 'B', 'C'];
    const typesWithData = Object.keys(map).filter((vt) => map[vt].sheets > 0);

    typesWithData.sort((a, b) => {
      const idxA = visaOrder.indexOf(a);
      const idxB = visaOrder.indexOf(b);
      if (idxA !== -1 && idxB !== -1) return idxA - idxB;
      if (idxA !== -1) return -1;
      if (idxB !== -1) return 1;
      return a.localeCompare(b);
    });

    if (typesWithData.length > 0) {
      return typesWithData.map((vt) => {
        const sheets = map[vt].sheets;
        // booklets = sheets / 50
        const booklets = sheets % 50 === 0 ? sheets / 50 : Math.round((sheets / 50) * 100) / 100;
        return {
          visaType: vt,
          booklets,
          sheets,
          remarks: map[vt].remarks || '',
        };
      });
    }

    return [];
  }, [isCalculated, isStubProposal, stockRecords, startDate, endDate]);

  // Stub proposal rows state for proposalType === 'section' or 'office'
  const [stubProposalRows, setStubProposalRows] = useState<StubProposalRow[]>(() => {
    if (initialSavedDraft?.stubProposalRows && initialSavedDraft.stubProposalRows.length > 0) {
      return initialSavedDraft.stubProposalRows;
    }
    return DEFAULT_STUB_EMPTY_ROWS;
  });

  useEffect(() => {
    if (isCalculated && stubRowsFromStock.length > 0) {
      setStubProposalRows(stubRowsFromStock);
    }
  }, [isCalculated, stubRowsFromStock]);

  const displayStubRows = useMemo(() => {
    if (isStubProposal) {
      if (!isCalculated) {
        return DEFAULT_STUB_EMPTY_ROWS;
      }
      if (stubRowsFromStock.length > 0) {
        return stubRowsFromStock;
      }
      if (stubProposalRows.length > 0 && stubProposalRows !== DEFAULT_STUB_EMPTY_ROWS) {
        return stubProposalRows;
      }
      return DEFAULT_STUB_EMPTY_ROWS;
    }
    return stubProposalRows;
  }, [isStubProposal, isCalculated, stubRowsFromStock, stubProposalRows]);

  const totalStubBooklets = useMemo(() => {
    return displayStubRows.reduce((sum, r) => sum + (Number(r.booklets) || 0), 0);
  }, [displayStubRows]);

  const totalStubSheets = useMemo(() => {
    return displayStubRows.reduce((sum, r) => sum + (Number(r.sheets) || 0), 0);
  }, [displayStubRows]);

  const handleUpdateStubRow = (index: number, field: keyof StubProposalRow, value: any) => {
    setStubProposalRows((prev) => {
      const copy = [...prev];
      const target = { ...copy[index], [field]: value };
      if (field === 'booklets') {
        target.sheets = (Number(value) || 0) * 50;
      }
      copy[index] = target;
      return copy;
    });
  };

  useEffect(() => {
    if (defaultDocThroughTitle) {
      setDocThroughTitle(cleanDocThroughTitle(defaultDocThroughTitle));
    } else if (proposalType === 'section') {
      setDocThroughTitle('លោកនាយរងការិយាល័យទទួលបន្ទុក');
    } else if (proposalType === 'office') {
      setDocThroughTitle('លោកអនុប្រធាននាយកដ្ឋានទទួលបន្ទុក');
    }
  }, [defaultDocThroughTitle, proposalType]);

  useEffect(() => {
    if (proposalType === 'office') {
      if (!docRecipientRank) setDocRecipientRank('លោកឧត្តមសេនីយ៍ទោ');
      if (!docRecipientTitle || docRecipientTitle.includes('នាយការិយាល័យ')) {
        setDocRecipientTitle('ប្រធាននាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត');
      }
      if (!docNumber || docNumber === '.......................' || (docNumber.includes('.....') && !docNumber.includes('សណ/២៦')) || docNumber === '...................សណ/២៦') {
        setDocNumber('.......................សណ/២៦');
      }
    }
  }, [proposalType]);

  // Helper to ensure clean signature titles
  const cleanRightSignTitle = (title?: string) => {
    const defaultRight = proposalType === 'office' ? 'នាយការិយាល័យ' : 'នាយផ្នែក';
    if (!title) return defaultRight;
    if (title.includes('មេត្តាជ្រាប') || title.includes('ជារបាយការណ៍') || title.includes('លោកនាយការិយាល័យ') || title.includes('បានឃើញ')) {
      return defaultRight;
    }
    if (proposalType === 'office' && title === 'នាយផ្នែក') {
      return 'នាយការិយាល័យ';
    }
    return title;
  };

  const cleanLeftSignTitle = (title?: string) => {
    const defaultLeft = proposalType === 'office' ? 'អនុប្រធាននាយកដ្ឋានទទួលបន្ទុក' : 'នាយរងការិយាល័យទទួលបន្ទុក';
    if (!title) return defaultLeft;
    if (title.includes('មេត្តាជ្រាប') || title.includes('ជារបាយការណ៍') || title.includes('បានឃើញ')) {
      return defaultLeft;
    }
    if (proposalType === 'office' && title === 'នាយរងការិយាល័យទទួលបន្ទុក') {
      return 'អនុប្រធាននាយកដ្ឋានទទួលបន្ទុក';
    }
    return title;
  };

  const cleanOfficerName = (name?: string) => {
    if (!name) return '';
    if (name.includes('ជារបាយការណ៍') || name.includes('មេត្តាជ្រាប') || name.includes('លោកនាយការិយាល័យ') || name.includes('បានឃើញ') || name.includes('ដោយក្តីអនុគ្រោះ')) {
      return '';
    }
    return name;
  };

  // Officer signatures (Left & Right 2 columns with 4-way movement controls)
  const [leftSignTitle, setLeftSignTitle] = useState<string>(() => cleanLeftSignTitle(initialSavedDraft?.leftSignTitle));
  const [leftSignOfficer, setLeftSignOfficer] = useState<string>(() => cleanOfficerName(initialSavedDraft?.leftSignOfficer));
  const [leftSignShiftX, setLeftSignShiftX] = useState<number>(() => (initialSavedDraft?.leftSignShiftX !== undefined && initialSavedDraft.leftSignShiftX !== 3.0) ? initialSavedDraft.leftSignShiftX : ((initialSavedDraft?.leftSignShift !== undefined && initialSavedDraft.leftSignShift !== 3.0) ? initialSavedDraft.leftSignShift : 0)); // Shift in cm (Default 0: strictly no overlap)
  const [leftSignShiftY, setLeftSignShiftY] = useState<number>(() => initialSavedDraft?.leftSignShiftY !== undefined ? initialSavedDraft.leftSignShiftY : 0); // Shift in pt
  const [rightSignTitle, setRightSignTitle] = useState<string>(() => cleanRightSignTitle(initialSavedDraft?.rightSignTitle));
  const [rightSignOfficer, setRightSignOfficer] = useState<string>(() => cleanOfficerName(initialSavedDraft?.rightSignOfficer));
  const [rightSignShiftX, setRightSignShiftX] = useState<number>(() => initialSavedDraft?.rightSignShiftX !== undefined ? initialSavedDraft.rightSignShiftX : 0); // Shift in cm
  const [rightSignShiftY, setRightSignShiftY] = useState<number>(() => initialSavedDraft?.rightSignShiftY !== undefined ? initialSavedDraft.rightSignShiftY : 0); // Shift in pt
  const [signSectionMarginTop, setSignSectionMarginTop] = useState<number>(() => initialSavedDraft?.signSectionMarginTop !== undefined ? initialSavedDraft.signSectionMarginTop : 4); // Spacing in pt
  const [tableShiftY, setTableShiftY] = useState<number>(() => {
    if (initialSavedDraft?.tableShiftY !== undefined) return initialSavedDraft.tableShiftY;
    const prefStr = typeof localStorage !== 'undefined' ? localStorage.getItem('robok_total_report_user_preferences') : null;
    if (prefStr) {
      try {
        const pref = JSON.parse(prefStr);
        if (pref.tableShiftY !== undefined) return pref.tableShiftY;
      } catch (e) {}
    }
    return 0; // Default 0pt (clean spacing)
  });

  const handleTableShiftChange = (newShift: number) => {
    setTableShiftY(newShift);
    try {
      const prefStr = localStorage.getItem('robok_total_report_user_preferences');
      const pref = prefStr ? JSON.parse(prefStr) : {};
      pref.tableShiftY = newShift;
      localStorage.setItem('robok_total_report_user_preferences', JSON.stringify(pref));
    } catch (e) {}
  };

  const [paragraphShiftY, setParagraphShiftY] = useState<number>(() => {
    if (initialSavedDraft?.paragraphShiftY !== undefined) return initialSavedDraft.paragraphShiftY;
    const prefStr = typeof localStorage !== 'undefined' ? localStorage.getItem('robok_total_report_user_preferences') : null;
    if (prefStr) {
      try {
        const pref = JSON.parse(prefStr);
        if (pref.paragraphShiftY !== undefined && pref.paragraphShiftY !== 16) return pref.paragraphShiftY;
      } catch (e) {}
    }
    return 6; // Default 6pt for a clean, proportional spacing between subject and intro paragraph
  });

  const handleParagraphShiftChange = (newShift: number) => {
    setParagraphShiftY(newShift);
    try {
      const prefStr = localStorage.getItem('robok_total_report_user_preferences');
      const pref = prefStr ? JSON.parse(prefStr) : {};
      pref.paragraphShiftY = newShift;
      localStorage.setItem('robok_total_report_user_preferences', JSON.stringify(pref));
    } catch (e) {}
  };
  const [showSignPositionPanel, setShowSignPositionPanel] = useState<boolean>(false);
  const [showDocDateModal, setShowDocDateModal] = useState<boolean>(false);
  const [customSignLunar, setCustomSignLunar] = useState<string>('');
  const [customSignSolar, setCustomSignSolar] = useState<string>('');

  // Available Teams & Selection for Team Report View (PDF format)
  const availableTeams = useMemo(() => {
    if (categories?.teams && categories.teams.length > 0) {
      return categories.teams.map((t) => t.name);
    }
    return Array.from(OFFICIAL_29_TEAMS);
  }, [categories?.teams]);

  const defaultTeam = useMemo(() => {
    // 1. Check assignedTeam if provided
    if (assignedTeam) {
      const direct = availableTeams.find((t) => t === assignedTeam || normalizeTeamName(t) === normalizeTeamName(assignedTeam));
      if (direct) return direct;
      const partial = availableTeams.find((t) => normalizeTeamName(t).includes(normalizeTeamName(assignedTeam)) || normalizeTeamName(assignedTeam).includes(normalizeTeamName(t)));
      if (partial) return partial;
    }
    // 2. Check userName if it matches a real team (and is not an admin/system user)
    if (userName && !userName.toLowerCase().includes('admin') && !userName.toLowerCase().includes('k2') && !userName.toLowerCase().includes('director') && !userName.toLowerCase().includes('immigration')) {
      const direct = availableTeams.find((t) => t === userName || normalizeTeamName(t) === normalizeTeamName(userName));
      if (direct) return direct;
      const partial = availableTeams.find((t) => normalizeTeamName(t).includes(normalizeTeamName(userName)) || normalizeTeamName(userName).includes(normalizeTeamName(t)));
      if (partial) return partial;
    }
    // 3. Default to 'កំពង់ផែ កោះកុង' or first available team
    const kohKong = availableTeams.find((t) => t.includes('កោះកុង'));
    return kohKong || availableTeams[0] || 'កំពង់ផែ កោះកុង';
  }, [assignedTeam, userName, availableTeams]);

  const isOfficeUser = currentRole === 'Secondary' || currentRole === 'Admin';
  const isTeamUser = !isOfficeUser;

  const [selectedTeam, setSelectedTeam] = useState<string>(() => {
    if (isTeamUser || assignedTeam) {
      return defaultTeam;
    }
    if (initialSavedDraft?.selectedTeam) {
      // Validate that saved team is an actual team, not an admin username
      const s = initialSavedDraft.selectedTeam;
      if (!s.toLowerCase().includes('admin') && !s.toLowerCase().includes('immigration')) {
        return s;
      }
    }
    return defaultTeam;
  });

  useEffect(() => {
    if (isTeamUser && defaultTeam && selectedTeam !== defaultTeam) {
      setSelectedTeam(defaultTeam);
    }
  }, [isTeamUser, defaultTeam]);

  const [viewMode, setViewMode] = useState<'office' | 'team'>(() => {
    if (initialViewMode) return initialViewMode;
    if (initialSavedDraft?.viewMode) return initialSavedDraft.viewMode;
    return 'team';
  });

  useEffect(() => {
    if (initialViewMode) {
      setViewMode(initialViewMode);
      setIsCalculated(false);
    }
  }, [initialViewMode]);

  // Ministry hierarchy & customizable header lines
  const [ministryHierarchy, setMinistryHierarchy] = useState<string[]>(() => {
    if (initialSavedDraft?.ministryHierarchy && Array.isArray(initialSavedDraft.ministryHierarchy)) {
      if (proposalType === 'office') {
        return initialSavedDraft.ministryHierarchy.filter((line: string) => !line.includes('ផ្នែករដ្ឋបាល'));
      }
      return initialSavedDraft.ministryHierarchy;
    }
    if (proposalType === 'office') {
      return [
        'ក្រសួងមហាផ្ទៃ',
        'អគ្គនាយកដ្ឋានអន្តោប្រវេសន៍',
        'នាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត',
        'ការិយាល័យទិដ្ឋាការចូល',
      ];
    }
    return [
      'ក្រសួងមហាផ្ទៃ',
      'អគ្គនាយកដ្ឋានអន្តោប្រវេសន៍',
      'នាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត',
      'ការិយាល័យទិដ្ឋាការចូល',
      'ផ្នែករដ្ឋបាល',
    ];
  });

  useEffect(() => {
    if (proposalType === 'office') {
      setMinistryHierarchy((prev) => prev.filter((l) => !l.includes('ផ្នែករដ្ឋបាល')));
    }
  }, [proposalType]);

  // MS Word-like Live Document Editing States
  const [isWordEditMode, setIsWordEditMode] = useState<boolean>(true); // Editable by default as requested
  const [showRibbonToolbar, setShowRibbonToolbar] = useState<boolean>(true); // Toggle hide/show Word toolbar
  const [activeFont, setActiveFont] = useState<string>('Khmer OS Siemreap');
  const [activeFontSize, setActiveFontSize] = useState<string>('10pt');
  const [lineSpacing, setLineSpacing] = useState<string>('1.5');
  const [showColorPicker, setShowColorPicker] = useState<boolean>(false);
  const [showHighlightPicker, setShowHighlightPicker] = useState<boolean>(false);
  const [showFindReplace, setShowFindReplace] = useState<boolean>(false);
  const [findText, setFindText] = useState<string>('');
  const [replaceText, setReplaceText] = useState<string>('');
  const [matchCount, setMatchCount] = useState<number | null>(null);
  const [docZoom, setDocZoom] = useState<number>(100);
  const [showRuler, setShowRuler] = useState<boolean>(true); // Microsoft Word-like Ruler
  const [activeTabStop, setActiveTabStop] = useState<number>(2.25); // Tab Stop position in cm (e.g. 2.25cm)
  const [customMargins, setCustomMargins] = useState<{ top: number; bottom: number; left: number; right: number }>(() => {
    const DEFAULT_MARGINS = { top: 0.5, bottom: 0.5, left: 2.8, right: 1.5 };
    if (initialSavedDraft?.customMargins) {
      return { ...initialSavedDraft.customMargins, top: 0.5 };
    }
    try {
      const prefStr = typeof localStorage !== 'undefined' ? (localStorage.getItem('robok_report_user_preferences') || localStorage.getItem('robok_total_margins') || localStorage.getItem('robok_margins')) : null;
      if (prefStr) {
        const pref = JSON.parse(prefStr);
        const margins = pref.customMargins || pref;
        if (margins && typeof margins.top === 'number') {
          return {
            top: 0.5,
            bottom: margins.bottom ?? 0.5,
            left: margins.left ?? 2.8,
            right: margins.right ?? 1.5,
          };
        }
      }
    } catch (e) {}
    return DEFAULT_MARGINS;
  });
  const [showMarginControls, setShowMarginControls] = useState<boolean>(false);

  // Microsoft Word Page Setup Dialog States (Alt+P+S+P)
  const [showPageSetupDialog, setShowPageSetupDialog] = useState<boolean>(false);
  const [pageSetupTab, setPageSetupTab] = useState<'margins' | 'paper' | 'layout'>('margins');
  const [pageOrientation, setPageOrientation] = useState<'portrait' | 'landscape'>('portrait');
  const [paperSize, setPaperSize] = useState<'A4' | 'Letter' | 'Legal' | 'A3'>('A4');
  const [tempMargins, setTempMargins] = useState<{ top: number; bottom: number; left: number; right: number }>(() => ({
    top: customMargins?.top ?? 0.5,
    bottom: customMargins?.bottom ?? 0.5,
    left: customMargins?.left ?? 2.8,
    right: customMargins?.right ?? 1.5,
  }));
  const [gutter, setGutter] = useState<number>(0);
  const [gutterPosition, setGutterPosition] = useState<'left' | 'top'>('left');
  const [headerMargin, setHeaderMargin] = useState<number>(1.25);
  const [footerMargin, setFooterMargin] = useState<number>(1.25);

  // Page Adjust / Scaling State (Default 97% for printing and PDF generation)
  const [adjustPagePercent, setAdjustPagePercent] = useState<number>(() => {
    try {
      const prefStr = localStorage.getItem('robok_report_user_preferences');
      if (prefStr) {
        const pref = JSON.parse(prefStr);
        if (typeof pref.adjustPagePercent === 'number') return pref.adjustPagePercent;
      }
      const saved = localStorage.getItem('robok_total_adjust_scale');
      if (saved) return parseFloat(saved) || 97;
    } catch (e) {}
    return 97; // Default 97% as requested
  });
  const [tempAdjustPagePercent, setTempAdjustPagePercent] = useState<number>(adjustPagePercent);

  // Microsoft Word Font & Advanced (Ctrl+D) Dialog States
  const [showFontDialog, setShowFontDialog] = useState<boolean>(false);
  const [fontDialogTab, setFontDialogTab] = useState<'font' | 'advanced'>('advanced');
  
  // Advanced Tab States (Character Spacing & OpenType)
  const [characterScale, setCharacterScale] = useState<string>('100%');
  const [characterSpacing, setCharacterSpacing] = useState<'Normal' | 'Expanded' | 'Condensed'>('Condensed');
  const [spacingByPt, setSpacingByPt] = useState<number>(1.1);
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
  const [dialogFontSize, setDialogFontSize] = useState<string>('10pt');
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

  const printAreaRef = useRef<HTMLDivElement>(null);
  const savedSelectionRef = useRef<Range | null>(null);
  const [isExporting, setIsExporting] = useState<boolean>(false);
  const [keyResetCounter, setKeyResetCounter] = useState<number>(0);
  const [lastSavedTime, setLastSavedTime] = useState<string | null>(() => initialSavedDraft?.savedAtFormatted || initialSavedDraft?.savedAt || null);
  const [showSaveSuccessToast, setShowSaveSuccessToast] = useState<boolean>(false);
  const [hasSavedDraftForPeriod, setHasSavedDraftForPeriod] = useState<boolean>(() => !!initialSavedDraft);
  const isRestoringDraftRef = useRef<boolean>(false);
  const [showTacteingModal, setShowTacteingModal] = useState<boolean>(false);
  const [tacteingSettings, setTacteingSettings] = useState<{ type: TacteingType; customImage: string | null }>(() => getSavedTacteingSettings());

  // Track user selection in contentEditable for font/size formatting
  useEffect(() => {
    const handleSelectionChange = () => {
      const sel = window.getSelection();
      if (sel && sel.rangeCount > 0 && printAreaRef.current) {
        const range = sel.getRangeAt(0);
        if (printAreaRef.current.contains(range.commonAncestorContainer)) {
          savedSelectionRef.current = range.cloneRange();
        }
      }
    };
    document.addEventListener('selectionchange', handleSelectionChange);
    return () => {
      document.removeEventListener('selectionchange', handleSelectionChange);
    };
  }, []);

  // Helper to find signature elements with fallback support for saved drafts
  const getSignatureDOMElements = () => {
    if (!printAreaRef.current) return { leftElem: null, rightElem: null, sectionElem: null };
    const root = printAreaRef.current;

    const sectionElem = (
      root.querySelector('[data-signature-section="true"]') ||
      root.querySelector('[data-signature-section]') ||
      (root.querySelectorAll('.grid.grid-cols-2')[root.querySelectorAll('.grid.grid-cols-2').length - 1] as HTMLElement) ||
      null
    ) as HTMLElement | null;

    let leftElem = root.querySelector('[data-signature-block="left"]') as HTMLElement | null;
    let rightElem = root.querySelector('[data-signature-block="right"]') as HTMLElement | null;

    if (!leftElem && sectionElem && sectionElem.children.length >= 1) {
      leftElem = sectionElem.children[0] as HTMLElement;
      leftElem.setAttribute('data-signature-block', 'left');
    }
    if (!rightElem && sectionElem && sectionElem.children.length >= 2) {
      rightElem = sectionElem.children[1] as HTMLElement;
      rightElem.setAttribute('data-signature-block', 'right');
    }

    return { leftElem, rightElem, sectionElem };
  };

  // Helper to synchronize signature positions to both React State & DOM elements
  const updateSignaturePositions = (
    newLeftX?: number,
    newLeftY?: number,
    newRightX?: number,
    newRightY?: number,
    newSectionTop?: number
  ) => {
    const lx = newLeftX !== undefined ? newLeftX : leftSignShiftX;
    const ly = newLeftY !== undefined ? newLeftY : leftSignShiftY;
    const rx = newRightX !== undefined ? newRightX : rightSignShiftX;
    const ry = newRightY !== undefined ? newRightY : rightSignShiftY;
    const st = newSectionTop !== undefined ? newSectionTop : signSectionMarginTop;

    if (newLeftX !== undefined) setLeftSignShiftX(lx);
    if (newLeftY !== undefined) setLeftSignShiftY(ly);
    if (newRightX !== undefined) setRightSignShiftX(rx);
    if (newRightY !== undefined) setRightSignShiftY(ry);
    if (newSectionTop !== undefined) setSignSectionMarginTop(st);

    const { leftElem, rightElem, sectionElem } = getSignatureDOMElements();

    if (leftElem) {
      leftElem.style.position = 'relative';
      leftElem.style.left = `${lx}cm`;
      leftElem.style.top = `${ly}pt`;
      leftElem.style.marginLeft = '0px';
      leftElem.style.marginTop = '0px';
      leftElem.style.transform = 'none';
    }
    if (rightElem) {
      rightElem.style.position = 'relative';
      rightElem.style.left = `${rx}cm`;
      rightElem.style.top = `${ry}pt`;
      rightElem.style.marginLeft = '0px';
      rightElem.style.marginTop = '0px';
      rightElem.style.transform = 'none';
    }
    if (sectionElem) {
      sectionElem.style.marginTop = `${st}pt`;
    }
  };

  // Restore saved draft on mount or when period changes
  useEffect(() => {
    let isCancelled = false;

    const restoreSavedReport = async () => {
      try {
        const storageKey = `${storagePrefix}${startDate}_${endDate}`;
        let savedData: any = null;

        // 1. Check exact date key in localStorage
        const savedStr = localStorage.getItem(storageKey);
        if (savedStr) {
          try {
            const parsed = JSON.parse(savedStr);
            if (parsed && parsed.startDate === startDate && parsed.endDate === endDate) {
              savedData = parsed;
            }
          } catch (e) {}
        }

        // 2. Check IndexedDB storage if localStorage was empty for this EXACT storage key
        if (!savedData) {
          const idbDraft = await idbStorage.getItem<any>('daily_team_operations', storageKey);
          if (idbDraft && idbDraft.startDate === startDate && idbDraft.endDate === endDate) {
            savedData = idbDraft;
          }
        }

        if (isCancelled) return;

        if (savedData) {
          setHasSavedDraftForPeriod(true);
          setLastSavedTime(savedData.savedAtFormatted || savedData.savedAt || null);
          if (savedData.docNumber !== undefined) setDocNumber(savedData.docNumber);
          if (savedData.docRecipientRank !== undefined) setDocRecipientRank(savedData.docRecipientRank);
          if (savedData.docRecipientTitle !== undefined) setDocRecipientTitle(savedData.docRecipientTitle);
          if (proposalType === 'office') {
            if (!savedData.docRecipientRank || savedData.docRecipientRank.includes('ឯកឧត្តម')) {
              setDocRecipientRank('លោកឧត្តមសេនីយ៍ទោ');
            }
            if (!savedData.docRecipientTitle || savedData.docRecipientTitle.includes('នាយការិយាល័យ')) {
              setDocRecipientTitle('ប្រធាននាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត');
            }
            if (!savedData.docNumber || savedData.docNumber === '.......................' || (savedData.docNumber.includes('.....') && !savedData.docNumber.includes('សណ/២៦')) || savedData.docNumber === '...................សណ/២៦') {
              setDocNumber('.......................សណ/២៦');
            }
          }
          if (savedData.docThroughTitle !== undefined) setDocThroughTitle(cleanDocThroughTitle(savedData.docThroughTitle));
          if (savedData.docReferenceText !== undefined) setDocReferenceText(savedData.docReferenceText);
          if (savedData.docSignDate !== undefined) setDocSignDate(savedData.docSignDate);
          if (savedData.leftSignTitle !== undefined) setLeftSignTitle(cleanLeftSignTitle(savedData.leftSignTitle));
          if (savedData.leftSignOfficer !== undefined) setLeftSignOfficer(cleanOfficerName(savedData.leftSignOfficer));
          const lx = (savedData.leftSignShiftX !== undefined && savedData.leftSignShiftX !== 3.0) ? savedData.leftSignShiftX : ((savedData.leftSignShift !== undefined && savedData.leftSignShift !== 3.0) ? savedData.leftSignShift : 0);
          const ly = savedData.leftSignShiftY !== undefined ? savedData.leftSignShiftY : 0;
          const rx = savedData.rightSignShiftX !== undefined ? savedData.rightSignShiftX : 0;
          const ry = savedData.rightSignShiftY !== undefined ? savedData.rightSignShiftY : 0;
          const st = savedData.signSectionMarginTop !== undefined ? savedData.signSectionMarginTop : 4;
          
          setLeftSignShiftX(lx);
          setLeftSignShiftY(ly);
          if (savedData.rightSignTitle !== undefined) setRightSignTitle(cleanRightSignTitle(savedData.rightSignTitle));
          if (savedData.rightSignOfficer !== undefined) setRightSignOfficer(cleanOfficerName(savedData.rightSignOfficer));
          setRightSignShiftX(rx);
          setRightSignShiftY(ry);
          setSignSectionMarginTop(st);
          if (savedData.tableShiftY !== undefined) setTableShiftY(savedData.tableShiftY);
          if (savedData.paragraphShiftY !== undefined) setParagraphShiftY(savedData.paragraphShiftY);
          if (savedData.customMargins) {
            setCustomMargins({
              ...savedData.customMargins,
              top: savedData.customMargins.top === 1.2 ? 0.5 : (savedData.customMargins.top ?? 0.5),
            });
          }
          if (savedData.activeTabStop !== undefined) setActiveTabStop(savedData.activeTabStop);
          if (savedData.activeFont) setActiveFont(savedData.activeFont);
          if (savedData.activeFontSize) setActiveFontSize(savedData.activeFontSize);
          if (savedData.lineSpacing) setLineSpacing(savedData.lineSpacing);
        } else {
          setHasSavedDraftForPeriod(false);
          setLastSavedTime(null);
          // Load latest saved user profile preferences so names/titles/margins stay preserved
          const prefStr = localStorage.getItem(userPrefKey);
          if (prefStr) {
            const pref = JSON.parse(prefStr);
            if (pref.docRecipientRank !== undefined) setDocRecipientRank(pref.docRecipientRank);
            if (pref.docRecipientTitle) setDocRecipientTitle(pref.docRecipientTitle);
            if (proposalType === 'office') {
              if (!pref.docRecipientRank || pref.docRecipientRank.includes('ឯកឧត្តម')) {
                setDocRecipientRank('លោកឧត្តមសេនីយ៍ទោ');
              }
              if (!pref.docRecipientTitle || pref.docRecipientTitle.includes('នាយការិយាល័យ')) {
                setDocRecipientTitle('ប្រធាននាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត');
              }
            }
            if (pref.docThroughTitle) setDocThroughTitle(cleanDocThroughTitle(pref.docThroughTitle));
            if (pref.docReferenceText) setDocReferenceText(pref.docReferenceText);
            if (pref.leftSignTitle) setLeftSignTitle(cleanLeftSignTitle(pref.leftSignTitle));
            if (pref.leftSignOfficer) setLeftSignOfficer(cleanOfficerName(pref.leftSignOfficer));
            if (pref.leftSignShiftX !== undefined && pref.leftSignShiftX !== 3.0) setLeftSignShiftX(pref.leftSignShiftX);
            else setLeftSignShiftX(0);
            if (pref.leftSignShiftY !== undefined) setLeftSignShiftY(pref.leftSignShiftY);
            if (pref.rightSignTitle) setRightSignTitle(cleanRightSignTitle(pref.rightSignTitle));
            if (pref.rightSignOfficer) setRightSignOfficer(cleanOfficerName(pref.rightSignOfficer));
            if (pref.rightSignShiftX !== undefined) setRightSignShiftX(pref.rightSignShiftX);
            if (pref.rightSignShiftY !== undefined) setRightSignShiftY(pref.rightSignShiftY);
            if (pref.signSectionMarginTop !== undefined) setSignSectionMarginTop(pref.signSectionMarginTop);
            if (pref.tableShiftY !== undefined) setTableShiftY(pref.tableShiftY);
            if (pref.paragraphShiftY !== undefined) setParagraphShiftY(pref.paragraphShiftY);
            if (pref.customMargins) {
              setCustomMargins({
                ...pref.customMargins,
                top: pref.customMargins.top === 1.2 ? 0.5 : (pref.customMargins.top ?? 0.5),
              });
            }
            if (pref.activeTabStop !== undefined) setActiveTabStop(pref.activeTabStop);
          }
        }
      } catch (e) {
        console.error('Error loading saved draft:', e);
      }
    };

    restoreSavedReport();

    return () => {
      isCancelled = true;
    };
  }, [startDate, endDate]);

  // Auto-persist custom margins forever to localStorage
  useEffect(() => {
    try {
      const prefStr = localStorage.getItem('robok_report_user_preferences');
      const pref = prefStr ? JSON.parse(prefStr) : {};
      pref.customMargins = customMargins;
      localStorage.setItem('robok_report_user_preferences', JSON.stringify(pref));
      localStorage.setItem('robok_total_margins', JSON.stringify(customMargins));
      localStorage.setItem('robok_margins', JSON.stringify(customMargins));
    } catch (e) {
      console.error('Error auto-persisting total report margins:', e);
    }
  }, [customMargins]);

  // Global Shortcut Interception (Ctrl+D for Microsoft Word Font Dialog & Ctrl+S for Save)
  useEffect(() => {
    const handleGlobalKeyDown = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && (e.key === 'd' || e.key === 'D')) {
        e.preventDefault();
        const sel = window.getSelection();
        if (sel && sel.rangeCount > 0 && printAreaRef.current) {
          const range = sel.getRangeAt(0);
          if (printAreaRef.current.contains(range.commonAncestorContainer)) {
            savedSelectionRef.current = range.cloneRange();
          }
        }
        setFontDialogTab('advanced');
        setShowFontDialog(true);
      } else if ((e.ctrlKey || e.metaKey) && (e.key === 's' || e.key === 'S')) {
        e.preventDefault();
        handleSaveDocument();
      }
    };
    window.addEventListener('keydown', handleGlobalKeyDown);
    return () => {
      window.removeEventListener('keydown', handleGlobalKeyDown);
    };
  }, [
    startDate,
    endDate,
    selectedMonth,
    selectedYear,
    useCustomRange,
    docNumber,
    docRecipientTitle,
    docThroughTitle,
    docSignDate,
    leftSignTitle,
    leftSignOfficer,
    rightSignTitle,
    rightSignOfficer,
    customMargins,
    activeTabStop,
    activeFont,
    activeFontSize,
    lineSpacing,
  ]);

  // Helper to compute date range given start year, start month (1-12), and duration in months (1, 3, 6, 9, 12)
  const calculateDateRange = (year: number, month: number, durationMonths: number) => {
    if (isRolling) {
      const startMStr = String(month).padStart(2, '0');
      const s = `${year}-${startMStr}-16`;

      const totalMonthsOffset = (month - 1) + durationMonths;
      const endYear = year + Math.floor(totalMonthsOffset / 12);
      const endMonth = (totalMonthsOffset % 12) + 1;
      const endMStr = String(endMonth).padStart(2, '0');
      const e = `${endYear}-${endMStr}-15`;

      return { startDate: s, endDate: e };
    }

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

  // Trigger calculation and show result
  const handleShowResult = () => {
    setIsCalculated(true);
    setKeyResetCounter((prev) => prev + 1);
  };

  // Helper to format table numbers: returns '-' when pending calculation
  const renderNum = (val: number | string | undefined | null) => {
    if (!isCalculated) return '-';
    if (val === undefined || val === null || val === '') return '0';
    if (typeof val === 'number') {
      return val.toLocaleString();
    }
    const n = Number(val);
    return isNaN(n) ? val : n.toLocaleString();
  };

  // Handler when user edits start date
  const handleStartDateChange = (val: string) => {
    let finalVal = val;
    if (isRolling && val) {
      const parts = val.split('-');
      if (parts.length === 3) {
        finalVal = `${parts[0]}-${parts[1]}-16`;
      }
    }
    setStartDate(finalVal);
    setActiveDurationPreset(null);
    setIsCalculated(false);
    setKeyResetCounter((prev) => prev + 1);
    if (finalVal) {
      const parts = finalVal.split('-');
      if (parts.length === 3) {
        const y = parseInt(parts[0], 10);
        const m = parseInt(parts[1], 10);
        if (!isNaN(y) && y >= 1900 && y <= 2100) setSelectedYear(y);
        if (!isNaN(m) && m >= 1 && m <= 12) setSelectedMonth(m);
      }
    }
  };

  // Handler when user edits end date
  const handleEndDateChange = (val: string) => {
    let finalVal = val;
    if (isRolling && val) {
      const parts = val.split('-');
      if (parts.length === 3) {
        finalVal = `${parts[0]}-${parts[1]}-15`;
      }
    }
    setEndDate(finalVal);
    setActiveDurationPreset(null);
    setIsCalculated(false);
    setKeyResetCounter((prev) => prev + 1);
  };

  // Handler when user edits document signing/creation date
  const handleDocSignDateChange = (val: string) => {
    setDocSignDate(val);
    setCustomSignLunar('');
    setCustomSignSolar('');
    setKeyResetCounter((prev) => prev + 1);
  };

  // Sync date range when month/year changes
  const handleMonthYearChange = (year: number, month: number) => {
    setSelectedYear(year);
    setSelectedMonth(month);
    setIsCalculated(false);

    let duration = 1;
    if (activeDurationPreset === 'q1') duration = 3;
    else if (activeDurationPreset === 's1') duration = 6;
    else if (activeDurationPreset === 'm9') duration = 9;
    else if (activeDurationPreset === 'full_year') duration = 12;

    const range = calculateDateRange(year, month, duration);
    setStartDate(range.startDate);
    setEndDate(range.endDate);
    setKeyResetCounter((prev) => prev + 1); // Refresh editable DOM nodes to show new dates
  };

  // Quick Preset Ranges
  const handlePresetRange = (preset: 'default_2018_2019' | 'this_month' | 'last_month' | 'next_month' | 'q1' | 's1' | 'm9' | 'full_year') => {
    const y = selectedYear || 2018;
    const m = selectedMonth || 12;
    let s = '';
    let e = '';
    setIsCalculated(false);

    if (preset === 'default_2018_2019') {
      if (isRolling) {
        s = '2018-12-16';
        e = '2019-06-15';
        setSelectedYear(2018);
        setSelectedMonth(12);
        setActiveDurationPreset('s1');
      } else {
        s = '2018-12-01';
        e = '2019-02-28';
        setSelectedYear(2018);
        setSelectedMonth(12);
        setActiveDurationPreset('q1');
      }
    } else if (preset === 'this_month') {
      const cy = selectedYear || new Date().getFullYear();
      const cm = selectedMonth || (new Date().getMonth() + 1);
      const range = calculateDateRange(cy, cm, 1);
      s = range.startDate;
      e = range.endDate;
      setActiveDurationPreset('this_month');
    } else if (preset === 'last_month') {
      let prevM = (selectedMonth || 1) - 1;
      let prevY = selectedYear || 2018;
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
      let nextY = selectedYear || 2018;
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
      // 3 Months starting from selectedMonth & selectedYear
      const range = calculateDateRange(y, m, 3);
      s = range.startDate;
      e = range.endDate;
      setActiveDurationPreset('q1');
    } else if (preset === 's1') {
      // 6 Months starting from selectedMonth & selectedYear
      const range = calculateDateRange(y, m, 6);
      s = range.startDate;
      e = range.endDate;
      setActiveDurationPreset('s1');
    } else if (preset === 'm9') {
      // 9 Months starting from selectedMonth & selectedYear
      const range = calculateDateRange(y, m, 9);
      s = range.startDate;
      e = range.endDate;
      setActiveDurationPreset('m9');
    } else if (preset === 'full_year') {
      // 12 Months starting from selectedMonth & selectedYear
      const range = calculateDateRange(y, m, 12);
      s = range.startDate;
      e = range.endDate;
      setActiveDurationPreset('full_year');
    }

    if (s && e) {
      setStartDate(s);
      setEndDate(e);
      setKeyResetCounter((prev) => prev + 1);
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
    if (preset === 'this_month') {
      const r = calculateDateRange(selectedYear, selectedMonth, 1);
      return startDate === r.startDate && endDate === r.endDate;
    }
    return false;
  };

  // Formatted date labels
  const startSolar = useMemo(() => getKhmerSolarParts(startDate), [startDate]);
  const endSolar = useMemo(() => getKhmerSolarParts(endDate), [endDate]);
  const prevEndDateStr = useMemo(() => {
    try {
      const d = parseDateInput(startDate);
      d.setDate(d.getDate() - 1);
      const y = d.getFullYear();
      const m = String(d.getMonth() + 1).padStart(2, '0');
      const day = String(d.getDate()).padStart(2, '0');
      return `${y}-${m}-${day}`;
    } catch {
      return startDate;
    }
  }, [startDate]);
  const prevSolar = useMemo(() => getKhmerSolarParts(prevEndDateStr), [prevEndDateStr]);
  const signLunar = useMemo(() => getKhmerLunarDate(docSignDate), [docSignDate]);
  const signSolar = useMemo(() => getKhmerSolarParts(docSignDate), [docSignDate]);

  // Dynamic report period title:
  // Formatter for report period (matches Office Report & Team Report)
  // e.g., 01 Dec 2018 to 31 Dec 2018 -> "ប្រចាំខែ ធ្នូ ឆ្នាំ២០១៨"
  // e.g., 01 Dec 2018 to 28 Feb 2019 / multi-months -> "ប្រចាំរយៈពេល ០៣ខែ"
  const reportPeriodLabel = useMemo(() => {
    if (!startDate || !endDate) return isRolling ? 'ប្រចាំខែរំកិល' : `ប្រចាំខែ ${startSolar.khmerMonth} ឆ្នាំ${startSolar.khmerYear}`;
    try {
      const sParts = startDate.split('-');
      const eParts = endDate.split('-');

      if (sParts.length === 3 && eParts.length === 3) {
        const startY = parseInt(sParts[0], 10);
        const startM = parseInt(sParts[1], 10); // 1-12
        const startD = parseInt(sParts[2], 10);
        const endY = parseInt(eParts[0], 10);
        const endM = parseInt(eParts[1], 10); // 1-12
        const endD = parseInt(eParts[2], 10);

        // Dedicated period labels for របក.សរុបរួមរំកិល (Rolling Report)
        if (isRolling) {
          let totalMonths = (endY - startY) * 12 + (endM - startM);
          if (totalMonths <= 0) {
            totalMonths = ((endM - startM + 12) % 12);
            if (totalMonths === 0) totalMonths = 12;
          }

          if (totalMonths <= 1) {
            return 'ប្រចាំខែរំកិល';
          }
          if (totalMonths === 3) {
            return 'ប្រចាំរំកិលត្រីមាសទី១';
          }
          if (totalMonths === 6) {
            return 'ប្រចាំរំកិលឆមាសទី១';
          }
          if (totalMonths === 9) {
            return 'ប្រចាំរំកិលនព្វមាស';
          }
          if (totalMonths === 12) {
            return 'ប្រចាំរំកិល ១២ខែ';
          }
          const monthCountKhmer = toKhmerNum(String(totalMonths).padStart(2, '0'));
          return `ប្រចាំរំកិល ${monthCountKhmer}ខែ`;
        }

        // Standard Report Logic (Unchanged)
        // Same month and year
        if (startY === endY && startM === endM) {
          return `ប្រចាំខែ ${startSolar.khmerMonth} ឆ្នាំ${startSolar.khmerYear}`;
        }

        // Full year (Jan 1 to Dec 31)
        if (startY === endY && startM === 1 && startD === 1 && endM === 12 && endD === 31) {
          return `ប្រចាំឆ្នាំ ${startSolar.khmerYear}`;
        }

        // Semester 1 (Jan-Jun)
        if (startY === endY && startM === 1 && endM === 6) {
          return `ប្រចាំឆមាសទី១ ឆ្នាំ${endSolar.khmerYear}`;
        }
        // Semester 2 (Jul-Dec)
        if (startY === endY && startM === 7 && endM === 12) {
          return `ប្រចាំឆមាសទី២ ឆ្នាំ${endSolar.khmerYear}`;
        }

        // Trimester / Quarter
        if (startY === endY && startM === 1 && endM === 3) {
          return `ប្រចាំត្រីមាសទី១ ឆ្នាំ${endSolar.khmerYear}`;
        }
        if (startY === endY && startM === 4 && endM === 6) {
          return `ប្រចាំត្រីមាសទី២ ឆ្នាំ${endSolar.khmerYear}`;
        }
        if (startY === endY && startM === 7 && endM === 9) {
          return `ប្រចាំត្រីមាសទី៣ ឆ្នាំ${endSolar.khmerYear}`;
        }
        if (startY === endY && startM === 10 && endM === 12) {
          return `ប្រចាំត្រីមាសទី៤ ឆ្នាំ${endSolar.khmerYear}`;
        }

        // Calculate total months difference
        let totalMonths = (endY - startY) * 12 + (endM - startM) + 1;
        if (totalMonths <= 0) {
          totalMonths = ((endM - startM + 12) % 12) + 1;
        }

        if (totalMonths <= 1) {
          return `ប្រចាំខែ ${startSolar.khmerMonth} ឆ្នាំ${startSolar.khmerYear}`;
        }

        // Multi-month period (e.g., 2 months -> ០២ខែ, 3 months -> ០៣ខែ, 6 months -> ០៦ខែ, 9 months -> ០៩ខែ, 12 months -> ១២ខែ)
        const monthCountKhmer = toKhmerNum(String(totalMonths).padStart(2, '0'));
        return `ប្រចាំរយៈពេល ${monthCountKhmer}ខែ`;
      }

      return isRolling ? 'ប្រចាំខែរំកិល' : `ប្រចាំខែ ${startSolar.khmerMonth} ឆ្នាំ${startSolar.khmerYear}`;
    } catch {
      return isRolling ? 'ប្រចាំខែរំកិល' : `ប្រចាំខែ ${startSolar.khmerMonth} ឆ្នាំ${startSolar.khmerYear}`;
    }
  }, [startDate, endDate, startSolar, endSolar, isRolling]);

  // Main Calculation Engine for Sticker Visas (Table ក)
  const stickerData = useMemo(() => {
    const normStart = normalizeDateToISO(startDate);
    const normEnd = normalizeDateToISO(endDate);

    type RowData = {
      visaType: string;
      openK2: number;
      openTeams: number;
      k1ToK2: number;
      teamReturnedToK2: number;
      issuedK2ToTeams: number;
      teamsUsed: number;
      damagedTeam: number;
      damagedK2: number;
      testSampleK2: number;
      endingK2: number;
      endingTeams: number;
      grandTotalEnding: number;
    };

    const rows: Record<string, RowData> = {};
    VISA_TYPES.forEach((vt) => {
      rows[vt] = {
        visaType: vt,
        openK2: OFFICIAL_K2_DEC_2018_BASELINE[vt] || 0,
        openTeams: OFFICIAL_TEAMS_DEC_2018_BASELINE[vt] || 0,
        k1ToK2: 0,
        teamReturnedToK2: 0,
        issuedK2ToTeams: 0,
        teamsUsed: 0,
        damagedTeam: 0,
        damagedK2: 0,
        testSampleK2: 0,
        endingK2: 0,
        endingTeams: 0,
        grandTotalEnding: 0,
      };
    });

    // Process stockRecords
    (stockRecords || []).forEach((rec) => {
      if (rec.stockType && rec.stockType !== 'sticker') return;
      if (isCeaRecord(rec)) return;

      const vt = normalizeVisaType(rec.visaType);
      if (!VISA_TYPES.includes(vt as any)) return;

      const qty = Number(rec.quantityBundles || rec.totalSheets || (rec as any).quantity || 0);
      const recDate = normalizeDateToISO(rec.date || '');
      const op = (rec.operationType || '').toLowerCase();

      if (recDate && recDate <= '2018-11-30') return;

      const isDamagedTeam =
        op === 'damagedteam' ||
        op === 'voidteam' ||
        op === 'missingteam' ||
        op === 'teamdamaged' ||
        op.includes('ខូចក្រុម') ||
        op.includes('ខ្វះក្រុម') ||
        op.includes('ខូចតាមក្រុម') ||
        op.includes('ខ្វះតាមក្រុម') ||
        ((op.includes('ខូច') || op.includes('មិនបានការ')) &&
          !op.includes('ក២') &&
          !op.includes('ក១') &&
          !op.includes('ការិយាល័យ') &&
          (rec.sourceFrom?.includes('ក្រុម') || (rec as any).visaTeamRobokName));

      const isDamagedK2 =
        !isDamagedTeam &&
        (op === 'damaged' ||
          op === 'damagedk2' ||
          op === 'damagedoffice' ||
          op === 'void' ||
          op === 'voidk2' ||
          op === 'voidoffice' ||
          op === 'missingoffice' ||
          op.includes('ខូចក') ||
          op.includes('មិនបានការក') ||
          op.includes('ខូច') ||
          op.includes('មិនបានការ'));

      const isTestSampleK2 =
        op === 'testprintk2' ||
        op === 'testk2' ||
        op === 'samplek2' ||
        op === 'testoffice' ||
        op === 'sampleoffice' ||
        op === 'test' ||
        op === 'sample' ||
        op === 'testprint' ||
        op.includes('សាកក') ||
        op.includes('សាកល្បង') ||
        op.includes('សាក');

      const isTransferTeam =
        op === 'transferteam' ||
        op === 'transferuseteam' ||
        op === 'transfer' ||
        op.includes('transfer') ||
        op.includes('ផ្ទេរ') ||
        (rec.sourceFrom && rec.sourceFrom.includes('ផ្ទេរ')) ||
        (rec.notes && rec.notes.includes('ផ្ទេរ'));

      if (normStart && recDate && recDate < normStart) {
        if (isTransferTeam) {
          // Ignore team-to-team transfers for Office (K2) stock calculations
        } else if (op === 'openk1' || op === 'receivek1' || op === 'k1' || op.includes('បញ្ចូលស្តុក') || op.includes('ក១')) {
          rows[vt].openK2 += qty;
        } else if (op === 'returnteam' || op === 'returnoffice' || op.includes('បង្វិល')) {
          rows[vt].openK2 += qty;
          rows[vt].openTeams -= qty;
        } else if (op === 'issueteam' || op.includes('បើកផ្តល់')) {
          rows[vt].openK2 -= qty;
          rows[vt].openTeams += qty;
        } else if (op === 'useteam' || op.includes('ប្រើប្រាស់')) {
          rows[vt].openTeams -= qty;
        } else if (isDamagedK2) {
          rows[vt].openK2 -= qty;
        } else if (isTestSampleK2) {
          rows[vt].openK2 -= qty;
        } else if (isDamagedTeam) {
          rows[vt].openTeams -= qty;
        }
      } else if (normStart && normEnd && recDate && recDate >= normStart && recDate <= normEnd) {
        if (isTransferTeam) {
          // Ignore team-to-team transfers for Office (K2) stock calculations
        } else if (op === 'openk1' || op === 'receivek1' || op === 'k1' || op.includes('បញ្ចូលស្តុក') || op.includes('ក១')) {
          rows[vt].k1ToK2 += qty;
        } else if (op === 'returnteam' || op === 'returnoffice' || op.includes('បង្វិល')) {
          rows[vt].teamReturnedToK2 += qty;
        } else if (op === 'issueteam' || op.includes('បើកផ្តល់')) {
          rows[vt].issuedK2ToTeams += qty;
        } else if (op === 'useteam' || op.includes('ប្រើប្រាស់')) {
          rows[vt].teamsUsed += qty;
        } else if (isDamagedTeam) {
          rows[vt].damagedTeam += qty;
        } else if (isDamagedK2) {
          rows[vt].damagedK2 += qty;
        } else if (isTestSampleK2) {
          rows[vt].testSampleK2 += qty;
        }
      }
    });

    // Process local daily operations
    try {
      const savedDaily = localStorage.getItem('app_daily_team_operations_v5');
      if (savedDaily) {
        const parsedDaily = JSON.parse(savedDaily);
        if (Array.isArray(parsedDaily)) {
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
            const recDate = normalizeDateToISO(dRec.date || '');
            if (!dRec.values || (recDate && recDate <= '2018-11-30')) return;

            VISA_TYPES.forEach((vt) => {
              const item = dRec.values[vt];
              const qty = parseInt(item?.quantity || '', 10) || 0;
              if (qty <= 0) return;

              const lookupKey = `${recDate}_${normTeam}_${vt}`;
              if (!existingUseKeySet.has(lookupKey)) {
                if (normStart && recDate < normStart) {
                  rows[vt].openTeams -= qty;
                } else if (normStart && normEnd && recDate >= normStart && recDate <= normEnd) {
                  rows[vt].teamsUsed += qty;
                }
              }
            });
          });
        }
      }
    } catch (e) {
      console.error('Error parsing daily team operations:', e);
    }

    // Calculate Ending Stocks
    VISA_TYPES.forEach((vt) => {
      const r = rows[vt];
      r.endingK2 =
        r.openK2 +
        r.k1ToK2 +
        r.teamReturnedToK2 -
        r.issuedK2ToTeams -
        r.damagedK2 -
        r.testSampleK2;

      r.endingTeams = r.openTeams + r.issuedK2ToTeams - r.teamReturnedToK2 - r.teamsUsed - r.damagedTeam;
      r.grandTotalEnding = r.endingK2 + r.endingTeams;
    });

    const totalRow: RowData = {
      visaType: 'សរុប',
      openK2: 0,
      openTeams: 0,
      k1ToK2: 0,
      teamReturnedToK2: 0,
      issuedK2ToTeams: 0,
      teamsUsed: 0,
      damagedTeam: 0,
      damagedK2: 0,
      testSampleK2: 0,
      endingK2: 0,
      endingTeams: 0,
      grandTotalEnding: 0,
    };

    VISA_TYPES.forEach((vt) => {
      const r = rows[vt];
      totalRow.openK2 += r.openK2;
      totalRow.openTeams += r.openTeams;
      totalRow.k1ToK2 += r.k1ToK2;
      totalRow.teamReturnedToK2 += r.teamReturnedToK2;
      totalRow.issuedK2ToTeams += r.issuedK2ToTeams;
      totalRow.teamsUsed += r.teamsUsed;
      totalRow.damagedTeam += r.damagedTeam;
      totalRow.damagedK2 += r.damagedK2;
      totalRow.testSampleK2 += r.testSampleK2;
      totalRow.endingK2 += r.endingK2;
      totalRow.endingTeams += r.endingTeams;
      totalRow.grandTotalEnding += r.grandTotalEnding;
    });

    return { rows: VISA_TYPES.map((vt) => rows[vt]), totalRow };
  }, [stockRecords, startDate, endDate, keyResetCounter]);

  // Main Calculation Engine for eVisa (Table ខ)
  const eVisaData = useMemo(() => {
    const normStart = normalizeDateToISO(startDate);
    const normEnd = normalizeDateToISO(endDate);

    let openBundles = 0;
    let receivedK1Bundles = 0;
    let issuedToTeamsBundles = 0;

    (stockRecords || []).forEach((rec) => {
      if (!isCeaRecord(rec) && rec.stockType !== 'evisa') return;
      const recDate = normalizeDateToISO(rec.date || '');
      const op = (rec.operationType || '').toLowerCase();
      const qty = Number(rec.quantityBundles || (rec as any).quantity || 0);

      if (recDate && recDate <= '2018-11-30') return;

      const isTransferTeam =
        op === 'transferteam' ||
        op === 'transferuseteam' ||
        op === 'transfer' ||
        op.includes('transfer') ||
        op.includes('ផ្ទេរ') ||
        (rec.sourceFrom && rec.sourceFrom.includes('ផ្ទេរ')) ||
        (rec.notes && rec.notes.includes('ផ្ទេរ'));

      if (normStart && recDate && recDate < normStart) {
        if (isTransferTeam) {
          // Ignore team-to-team transfers for Office (K2) stock calculations
        } else if (op === 'openk1' || op === 'receivek1' || op === 'k1') {
          openBundles += qty;
        } else if (op === 'issueteam') {
          openBundles -= qty;
        }
      } else if (normStart && normEnd && recDate && recDate >= normStart && recDate <= normEnd) {
        if (isTransferTeam) {
          // Ignore team-to-team transfers for Office (K2) stock calculations
        } else if (op === 'openk1' || op === 'receivek1' || op === 'k1') {
          receivedK1Bundles += qty;
        } else if (op === 'issueteam') {
          issuedToTeamsBundles += qty;
        }
      }
    });

    const endingBundles = openBundles + receivedK1Bundles - issuedToTeamsBundles;

    return {
      openBundles,
      receivedK1Bundles,
      issuedToTeamsBundles,
      endingBundles,
    };
  }, [stockRecords, startDate, endDate, keyResetCounter]);

  // Team displayName and Leader title
  const teamDisplayName = useMemo(() => {
    const raw = selectedTeam || defaultTeam || 'កំពង់ផែ កោះកុង';
    return formatReportTeamName(raw);
  }, [selectedTeam, defaultTeam]);

  const teamLeaderTitle = useMemo(() => {
    return 'ប្រធានក្រុមផ្តល់ទិដ្ឋាការ';
  }, []);

  // Team Period Title Formatter (strictly synchronized with reportPeriodLabel of Office report)
  const { teamPeriodKhmerTitle, teamPeriodShortKhmerTitle } = useMemo(() => {
    return {
      teamPeriodKhmerTitle: reportPeriodLabel,
      teamPeriodShortKhmerTitle: reportPeriodLabel,
    };
  }, [reportPeriodLabel]);

  // Main Calculation Engine for Individual Team Stock Report (Matching PDF)
  // Mapping logic as requested:
  // ចំនួនសល់ខែចាស់ = ស្តុកចាស់ក្រុម (Opening stock / initial baseline + movements before startDate)
  // ចំនួនបើកថ្មី = ចំនួនបើកពីក២ (Visas issued from K2 within period [startDate, endDate])
  // ចំនួនបង្វិល/ខូច/ខ្វះ = ទិដ្ឋាការបង្វិលទៅក២ (returnTeam) / ទិដ្ឋាការខូចក្រុម (damagedTeam) / ទិដ្ឋាការខ្វះក្រុម (missingTeam)
  // ចំនួនប្រើ = ការប្រើប្រាស់តាមក្រុម (useTeam within period [startDate, endDate] + daily operations)
  // សល់ខែបន្ទាប់ = ចំនួនសល់ខែចាស់ + ចំនួនបើកថ្មី - ចំនួនបង្វិល/ខូច/ខ្វះ - ចំនួនប្រើ
  const teamStickerData = useMemo(() => {
    const normStart = normalizeDateToISO(startDate);
    const normEnd = normalizeDateToISO(endDate);
    const curTeam = selectedTeam || defaultTeam || 'កំពង់ផែ កោះកុង';
    const curTeamNorm = normalizeTeamName(curTeam);
    const currentTargetTeam = matchTeamInList(curTeam, Array.from(OFFICIAL_29_TEAMS)) || curTeam;

    const opening: Record<string, number> = {};
    const newIssued: Record<string, number> = {};
    const returnedDamagedMissing: Record<string, number> = {};
    const used: Record<string, number> = {};
    const ending: Record<string, number> = {};

    // 1. Check if there are recorded old stock records for this team
    const recordedOldStockMap: Record<string, number> = {};
    let hasRecordedOldStock = false;

    (stockRecords || []).forEach((rec) => {
      if (!isOldStockTeamRecord(rec)) return;
      const rawTeam = resolveRecordTeamName(rec);
      const targetTeam = matchTeamInList(rawTeam, Array.from(OFFICIAL_29_TEAMS));
      const normTeam = normalizeTeamName(rawTeam);

      const isThisTeam =
        rawTeam === curTeam ||
        normTeam === curTeamNorm ||
        (targetTeam && targetTeam === currentTargetTeam) ||
        (curTeamNorm && normTeam && (normTeam.includes(curTeamNorm) || curTeamNorm.includes(normTeam)));

      if (!isThisTeam) return;

      const vt = normalizeVisaType(rec.visaType);
      if (!vt || !VISA_TYPES.includes(vt as any)) return;

      const qty = Number(rec.quantityBundles || rec.totalSheets || (rec as any).quantity || 0);
      recordedOldStockMap[vt] = (recordedOldStockMap[vt] || 0) + qty;
      hasRecordedOldStock = true;
    });

    // Initialize baseline opening values
    VISA_TYPES.forEach((vt) => {
      let initialVal = 0;
      if (hasRecordedOldStock && recordedOldStockMap[vt] !== undefined) {
        initialVal = recordedOldStockMap[vt];
      } else {
        initialVal =
          DEFAULT_OPENING_MATRIX[curTeam]?.[vt] ??
          DEFAULT_OPENING_MATRIX[curTeamNorm]?.[vt] ??
          DEFAULT_OPENING_MATRIX[currentTargetTeam]?.[vt] ??
          0;
      }
      opening[vt] = initialVal;
      newIssued[vt] = 0;
      returnedDamagedMissing[vt] = 0;
      used[vt] = 0;
      ending[vt] = 0;
    });

    // 2. Iterate all stock records
    (stockRecords || []).forEach((rec) => {
      if (rec.stockType && rec.stockType !== 'sticker') return;
      if (isCeaRecord(rec)) return;
      if (isOldStockTeamRecord(rec)) return; // Recorded old stock is already loaded as opening baseline

      const vt = normalizeVisaType(rec.visaType);
      if (!vt || !VISA_TYPES.includes(vt as any)) return;

      const recDate = normalizeDateToISO(rec.date || '');
      // Benchmark rule: ignore any movements prior to or on 2018-11-30
      if (recDate && recDate <= '2018-11-30') return;

      const rawTeam = resolveRecordTeamName(rec);
      const targetTeam = matchTeamInList(rawTeam, Array.from(OFFICIAL_29_TEAMS));
      const normTeam = normalizeTeamName(rawTeam);
      const op = (rec.operationType || '').trim().toLowerCase();

      const isSender =
        rawTeam === curTeam ||
        normTeam === curTeamNorm ||
        (targetTeam && targetTeam === currentTargetTeam) ||
        (curTeamNorm && normTeam && (normTeam.includes(curTeamNorm) || curTeamNorm.includes(normTeam))) ||
        (curTeam && (rec as any).issuedTo && normalizeTeamName((rec as any).issuedTo) === curTeamNorm) ||
        (curTeam && (rec as any).destinationTo && normalizeTeamName((rec as any).destinationTo) === curTeamNorm) ||
        (curTeam && rec.sourceFrom && normalizeTeamName(rec.sourceFrom) === curTeamNorm);

      const isRecipient =
        op === 'transferteam' &&
        rec.recipientTeamName &&
        (rec.recipientTeamName === curTeam ||
         normalizeTeamName(rec.recipientTeamName) === curTeamNorm ||
         matchTeamInList(rec.recipientTeamName, Array.from(OFFICIAL_29_TEAMS)) === currentTargetTeam);

      const isThisTeam = isSender || isRecipient;

      if (!isThisTeam) return;

      const qty = Number(rec.quantityBundles || rec.totalSheets || (rec as any).quantity || 0);
      if (qty <= 0) return;

      const src = (rec.sourceFrom || '').trim().toLowerCase();
      const desc = (((rec as any).description || (rec as any).note || '') + '').trim().toLowerCase();

      // Operation definitions:
      // ចំនួនបើកថ្មី = ចំនួនបើកពីក២ (issueTeam) ឬ ទទួលបានផ្ទេរពីក្រុមផ្សេង (isRecipient)
      const isIssued =
        isRecipient ||
        op === 'issueteam' ||
        op.includes('issueteam') ||
        op.includes('បើកផ្តល់') ||
        op.includes('បើកពីក២') ||
        op.includes('បើកជូន') ||
        src.includes('បើកពីក២') ||
        src.includes('បើកផ្តល់') ||
        desc.includes('បើកពីក២');

      // ចំនួនបង្វិល/ខូច/ខ្វះ = ទិដ្ឋាការបង្វិលទៅក២ (returnTeam) / ទិដ្ឋាការខូចក្រុម (damagedTeam) / ទិដ្ឋាការខ្វះក្រុម (missingTeam)
      const isReturned =
        op === 'returnteam' ||
        op === 'returnoffice' ||
        op.includes('returnteam') ||
        op.includes('បង្វិល') ||
        src.includes('បង្វិល') ||
        desc.includes('បង្វិល');

      const isDamaged =
        op === 'damagedteam' ||
        op === 'damaged' ||
        op === 'voidteam' ||
        op === 'teamdamaged' ||
        op.includes('damagedteam') ||
        op.includes('ខូច') ||
        src.includes('ខូច') ||
        desc.includes('ខូច');

      const isMissing =
        op === 'missingteam' ||
        op === 'missing' ||
        op.includes('missing') ||
        op.includes('ខ្វះ') ||
        op.includes('បាត់') ||
        src.includes('ខ្វះ') ||
        src.includes('បាត់') ||
        desc.includes('ខ្វះ');

      const isTransfer =
        !isRecipient && (
          op === 'transferteam' ||
          op.includes('transfer') ||
          op.includes('ផ្ទេរ')
        );

      // ចំនួនប្រើ = ការប្រើប្រាស់តាមក្រុម (useTeam)
      const isUsed =
        op === 'useteam' ||
        op.includes('useteam') ||
        op.includes('ប្រើប្រាស់') ||
        op.includes('ប្រើប្រាស') ||
        op.includes('ប្រើ') ||
        src.includes('ប្រើប្រាស់') ||
        desc.includes('ប្រើប្រាស់');

      if (normStart && recDate && recDate < normStart) {
        // Transactions prior to startDate adjust the starting balance (ចំនួនសល់ខែចាស់ = ស្តុកចាស់ក្រុម)
        if (isIssued) {
          opening[vt] += qty;
        } else if (isReturned || isDamaged || isMissing || isUsed || isTransfer) {
          opening[vt] -= qty;
        }
      } else {
        const isInRange = (!normStart || recDate >= normStart) && (!normEnd || recDate <= normEnd);
        if (isInRange) {
          // Transactions within [startDate, endDate]
          if (isIssued) {
            newIssued[vt] += qty;
          } else if (isReturned || isDamaged || isMissing || isTransfer) {
            returnedDamagedMissing[vt] += qty;
          } else if (isUsed) {
            used[vt] += qty;
          }
        }
      }
    });

    // 3. Process daily team usage operations (app_daily_team_operations_v5 and app_daily_team_records_v1)
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
          if (
            (sr.operationType === 'useTeam' || sr.operationType?.toLowerCase().includes('useteam')) &&
            !isCeaRecord(sr)
          ) {
            const dIso = normalizeDateToISO(sr.date || '');
            const rawT = resolveRecordTeamName(sr);
            const tNorm = normalizeTeamName(rawT);
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
          const recDate = normalizeDateToISO(dRec.date || '');
          if (!dRec.values || (recDate && recDate <= '2018-11-30')) return;
          const targetTeam = matchTeamInList(rawTeamName, Array.from(OFFICIAL_29_TEAMS));

          const isThisTeam =
            rawTeamName === curTeam ||
            normTeam === curTeamNorm ||
            (targetTeam && targetTeam === currentTargetTeam) ||
            (curTeamNorm && normTeam && (normTeam.includes(curTeamNorm) || curTeamNorm.includes(normTeam)));

          if (!isThisTeam) return;

          VISA_TYPES.forEach((vt) => {
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
                opening[vt] -= qty;
              } else {
                const isInRange = (!normStart || recDate >= normStart) && (!normEnd || recDate <= normEnd);
                if (isInRange) {
                  used[vt] += qty;
                }
              }
            }
          });
        });
      }
    } catch (e) {}

    // 4. Calculate ending stock: សល់ខែបន្ទាប់ = ចំនួនសល់ខែចាស់ + ចំនួនបើកថ្មី - ចំនួនបង្វិល/ខូច/ខ្វះ - ចំនួនប្រើ
    VISA_TYPES.forEach((vt) => {
      ending[vt] = opening[vt] + newIssued[vt] - returnedDamagedMissing[vt] - used[vt];
    });

    const sumValues = (record: Record<string, number>) => {
      return VISA_TYPES.reduce((acc, vt) => acc + (record[vt] || 0), 0);
    };

    return {
      opening: { values: opening, total: sumValues(opening) },
      newIssued: { values: newIssued, total: sumValues(newIssued) },
      returnedDamagedMissing: { values: returnedDamagedMissing, total: sumValues(returnedDamagedMissing) },
      used: { values: used, total: sumValues(used) },
      ending: { values: ending, total: sumValues(ending) },
    };
  }, [stockRecords, startDate, endDate, selectedTeam, defaultTeam, keyResetCounter]);

  // Execute MS Word Rich Text command
  const executeCommand = (command: string, value: string = '') => {
    document.execCommand(command, false, value);
  };

  // Insert Tab at cursor based on active Tab Stop (in cm) set on the ruler
  const handleInsertTab = (customWidthCm?: number) => {
    const widthCm = customWidthCm !== undefined ? customWidthCm : (activeTabStop || 2.0);
    const sel = window.getSelection();
    if (sel && sel.rangeCount > 0) {
      const range = sel.getRangeAt(0);
      range.deleteContents();
      
      // Create inline styled tab span matching ruler tab stop
      const tabSpan = document.createElement('span');
      tabSpan.style.display = 'inline-block';
      tabSpan.style.width = `${widthCm}cm`;
      tabSpan.style.minWidth = `${widthCm}cm`;
      tabSpan.innerHTML = '&nbsp;';
      tabSpan.className = 'tab-space-marker';
      
      range.insertNode(tabSpan);
      range.setStartAfter(tabSpan);
      range.setEndAfter(tabSpan);
      sel.removeAllRanges();
      sel.addRange(range);
    }
  };

  // Insert Tab or Spaces at cursor
  const handleInsertSpace = (count: number = 4) => {
    const sel = window.getSelection();
    if (sel && sel.rangeCount > 0) {
      const range = sel.getRangeAt(0);
      range.deleteContents();
      const spacesText = '\u00A0'.repeat(count); // Non-breaking spaces
      const textNode = document.createTextNode(spacesText);
      range.insertNode(textNode);
      range.setStartAfter(textNode);
      range.setEndAfter(textNode);
      sel.removeAllRanges();
      sel.addRange(range);
    }
  };

  // Keyboard Tab & Shortcut Key interception inside editable paper
  const handleKeyDownOnPaper = (e: React.KeyboardEvent<HTMLDivElement>) => {
    if (!isWordEditMode) return;
    if (e.key === 'Tab') {
      e.preventDefault();
      handleInsertTab();
    } else if ((e.ctrlKey || e.metaKey) && (e.key === 'd' || e.key === 'D')) {
      e.preventDefault();
      const sel = window.getSelection();
      if (sel && sel.rangeCount > 0 && printAreaRef.current) {
        const range = sel.getRangeAt(0);
        if (printAreaRef.current.contains(range.commonAncestorContainer)) {
          savedSelectionRef.current = range.cloneRange();
        }
      }
      setFontDialogTab('advanced');
      setShowFontDialog(true);
    } else if ((e.ctrlKey || e.metaKey) && (e.key === 's' || e.key === 'S')) {
      e.preventDefault();
      handleSaveDocument();
    }
  };

  // Apply Microsoft Word Font & Character Spacing (Condensed) Settings
  const handleApplyFontDialog = () => {
    // Restore selection if available
    const sel = window.getSelection();
    if (savedSelectionRef.current && (!sel || sel.rangeCount === 0 || sel.isCollapsed)) {
      sel?.removeAllRanges();
      sel?.addRange(savedSelectionRef.current);
    }

    const currentSel = window.getSelection();
    if (currentSel && currentSel.rangeCount > 0 && !currentSel.isCollapsed) {
      const range = currentSel.getRangeAt(0);
      const span = document.createElement('span');

      // 1. Character Spacing (Condensed / Expanded / Normal)
      if (characterSpacing === 'Condensed') {
        span.style.letterSpacing = `-${spacingByPt}pt`;
      } else if (characterSpacing === 'Expanded') {
        span.style.letterSpacing = `${spacingByPt}pt`;
      } else {
        span.style.letterSpacing = 'normal';
      }

      // 2. Character Scale (e.g. 100%, 90%, 80%, 66%, 50%)
      if (characterScale !== '100%') {
        const scaleVal = parseFloat(characterScale.replace('%', '')) / 100;
        span.style.display = 'inline-block';
        span.style.transform = `scaleX(${scaleVal})`;
        span.style.transformOrigin = 'left center';
      }

      // 3. Vertical Position (Raised / Lowered / Normal)
      if (characterPosition === 'Raised' && positionByPt > 0) {
        span.style.position = 'relative';
        span.style.top = `-${positionByPt}pt`;
      } else if (characterPosition === 'Lowered' && positionByPt > 0) {
        span.style.position = 'relative';
        span.style.top = `${positionByPt}pt`;
      }

      // 4. Font Family
      if (dialogFontFamily) {
        if (dialogFontFamily.includes('Mool') || dialogFontFamily.includes('Moul')) {
          span.style.fontFamily = `'Khmer Mool1', 'Khmer Mool 1', 'Khmer Mool', 'Khmer OS Mool1', 'Khmer OS Muol Light', 'Moul', serif`;
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

      // 5. Font Size
      if (dialogFontSize) {
        span.style.fontSize = dialogFontSize;
      }

      // 6. Font Style (Bold / Italic)
      if (dialogFontStyle.includes('Bold')) {
        span.style.fontWeight = 'bold';
      }
      if (dialogFontStyle.includes('Italic')) {
        span.style.fontStyle = 'italic';
      }

      // 7. Font Color
      if (dialogFontColor) {
        span.style.color = dialogFontColor;
      }

      // 8. Underline & Effects
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
    } else if (printAreaRef.current) {
      // Apply letter-spacing to the document if no specific text selected
      if (characterSpacing === 'Condensed') {
        printAreaRef.current.style.letterSpacing = `-${spacingByPt}pt`;
      } else if (characterSpacing === 'Expanded') {
        printAreaRef.current.style.letterSpacing = `${spacingByPt}pt`;
      } else {
        printAreaRef.current.style.letterSpacing = 'normal';
      }
    }

    setShowFontDialog(false);
  };

  // Handle clicking on Ruler to set Tab Stop at that exact cm
  const handleRulerClick = (e: React.MouseEvent<HTMLDivElement>) => {
    const rect = e.currentTarget.getBoundingClientRect();
    const clickX = e.clientX - rect.left;
    const totalWidthPx = rect.width;
    const clickedTotalCm = (clickX / totalWidthPx) * 21.0;
    const relativeCm = clickedTotalCm - customMargins.left;
    const maxContentCm = 21.0 - customMargins.left - customMargins.right;
    if (relativeCm > 0.2 && relativeCm < maxContentCm) {
      const rounded = Math.round(relativeCm * 2) / 2; // nearest 0.5cm
      setActiveTabStop(rounded > 0 ? rounded : 0.5);
    }
  };

  const FONT_SIZE_STEPS = [
    '8pt',
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

  // Change font family for selection or document
  const handleApplyFont = (font: string) => {
    setActiveFont(font);

    // Restore selection if available
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
        span.style.fontFamily = `'Khmer Mool1', 'Khmer Mool 1', 'Khmer Mool', 'Khmer OS Mool1', 'Khmer OS Muol Light', 'Moul', serif`;
      } else if (font.includes('Siemreap')) {
        span.style.fontFamily = `'Khmer OS Siemreap', 'Siemreap', sans-serif`;
      } else if (font.includes('Battambang')) {
        span.style.fontFamily = `'Khmer OS Battambang', 'Battambang', sans-serif`;
      } else if (font.includes('Bokor')) {
        span.style.fontFamily = `'Khmer OS Bokor', 'Bokor', cursive`;
      } else if (font.includes('Freehand')) {
        span.style.fontFamily = `'Khmer OS Freehand', 'Freehand', cursive`;
      } else if (font.includes('Kantumruy')) {
        span.style.fontFamily = `'Kantumruy Pro', sans-serif`;
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
        document.execCommand('fontName', false, font);
      }
    } else if (printAreaRef.current) {
      printAreaRef.current.style.fontFamily = `'${font}', sans-serif`;
    }
  };

  // Change font size for selection or document
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
    } else if (printAreaRef.current) {
      printAreaRef.current.style.fontSize = fontSize;
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
    if (printAreaRef.current) {
      printAreaRef.current.style.lineHeight = spacing;
    }
  };

  const handleSelectAll = () => {
    if (printAreaRef.current) {
      const range = document.createRange();
      range.selectNodeContents(printAreaRef.current);
      const sel = window.getSelection();
      sel?.removeAllRanges();
      sel?.addRange(range);
      savedSelectionRef.current = range.cloneRange();
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

  const handleFindNext = () => {
    if (!findText) return;
    if ((window as any).find) {
      const found = (window as any).find(findText, false, false, true, false, false, false);
      if (!found) {
        // Wrap around to beginning
        (window as any).find(findText, false, false, true, false, true, false);
      }
    }
  };

  const handleReplaceAll = () => {
    if (!findText || !printAreaRef.current) return;
    const html = printAreaRef.current.innerHTML;
    const regex = new RegExp(findText.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'g');
    const matches = html.match(regex);
    setMatchCount(matches ? matches.length : 0);
    printAreaRef.current.innerHTML = html.replace(regex, replaceText);
  };

  // Save document state and customized content to local storage (Microsoft Word Save)
  const handleSaveDocument = () => {
    try {
      if (!printAreaRef.current) return;
      const htmlContent = printAreaRef.current.innerHTML;
      const now = new Date();
      const parts = getKhmerSolarParts(now);
      const hours = String(now.getHours()).padStart(2, '0');
      const minutes = String(now.getMinutes()).padStart(2, '0');
      const seconds = String(now.getSeconds()).padStart(2, '0');
      const formattedTime = `ម៉ោង ${hours}:${minutes}:${seconds} ថ្ងៃទី${parts.khmerDay} ខែ${parts.khmerMonth} ឆ្នាំ${parts.khmerYear}`;

      const storageKey = `${storagePrefix}${startDate}_${endDate}`;
      const draftData = {
        id: `draft_${startDate}_${endDate}`,
        startDate,
        endDate,
        selectedMonth,
        selectedYear,
        useCustomRange,
        docNumber,
        docRecipientRank,
        docRecipientTitle,
        docThroughTitle,
        docReferenceText,
        docSignDate,
        stubProposalRows,
        leftSignTitle,
        leftSignOfficer,
        leftSignShiftX,
        leftSignShiftY,
        rightSignTitle,
        rightSignOfficer,
        rightSignShiftX,
        rightSignShiftY,
        signSectionMarginTop,
        tableShiftY,
        paragraphShiftY,
        customMargins,
        activeTabStop,
        activeFont,
        activeFontSize,
        lineSpacing,
        savedHtml: htmlContent,
        savedAt: now.toISOString(),
        savedAtFormatted: formattedTime,
      };

      localStorage.setItem(storageKey, JSON.stringify(draftData));
      localStorage.setItem(latestDraftKey, JSON.stringify(draftData));
      // Mirror in IndexedDB for resilience
      idbStorage.setItem('daily_team_operations', storageKey, draftData);
      idbStorage.setItem('daily_team_operations', latestDraftKey, draftData);
      localStorage.setItem(
        userPrefKey,
        JSON.stringify({
          docRecipientRank,
          docRecipientTitle,
          docThroughTitle,
          docReferenceText,
          leftSignTitle,
          leftSignOfficer,
          leftSignShiftX,
          leftSignShiftY,
          rightSignTitle,
          rightSignOfficer,
          rightSignShiftX,
          rightSignShiftY,
          signSectionMarginTop,
          tableShiftY,
          paragraphShiftY,
          customMargins,
          activeTabStop,
        })
      );

      setLastSavedTime(formattedTime);
      setHasSavedDraftForPeriod(true);
      setShowSaveSuccessToast(true);
      setTimeout(() => {
        setShowSaveSuccessToast(false);
      }, 3500);
    } catch (err) {
      console.error('Failed to save document:', err);
      alert('មានបញ្ហាក្នុងការរក្សាទុកទិន្នន័យឯកសារ');
    }
  };

  // Auto-persist before unmount/navigation so edited report is safely kept
  useEffect(() => {
    return () => {
      try {
        if (printAreaRef.current) {
          const htmlContent = printAreaRef.current.innerHTML;
          if (htmlContent && htmlContent.length > 50) {
            const now = new Date();
            const parts = getKhmerSolarParts(now);
            const hours = String(now.getHours()).padStart(2, '0');
            const minutes = String(now.getMinutes()).padStart(2, '0');
            const seconds = String(now.getSeconds()).padStart(2, '0');
            const formattedTime = `ម៉ោង ${hours}:${minutes}:${seconds} ថ្ងៃទី${parts.khmerDay} ខែ${parts.khmerMonth} ឆ្នាំ${parts.khmerYear}`;

            const storageKey = `${storagePrefix}${startDate}_${endDate}`;
            const draftData = {
              id: `draft_${startDate}_${endDate}`,
              startDate,
              endDate,
              selectedMonth,
              selectedYear,
              useCustomRange,
              docNumber,
              docRecipientTitle,
              docThroughTitle,
              docReferenceText,
              docSignDate,
              stubProposalRows,
              leftSignTitle,
              leftSignOfficer,
              leftSignShiftX,
              leftSignShiftY,
              rightSignTitle,
              rightSignOfficer,
              rightSignShiftX,
              rightSignShiftY,
              signSectionMarginTop,
              tableShiftY,
              paragraphShiftY,
              customMargins,
              activeTabStop,
              activeFont,
              activeFontSize,
              lineSpacing,
              savedHtml: htmlContent,
              savedAt: now.toISOString(),
              savedAtFormatted: formattedTime,
            };

            localStorage.setItem(storageKey, JSON.stringify(draftData));
            localStorage.setItem(latestDraftKey, JSON.stringify(draftData));
            idbStorage.setItem('daily_team_operations', storageKey, draftData);
            idbStorage.setItem('daily_team_operations', latestDraftKey, draftData);
          }
        }
      } catch (e) {}
    };
  }, [
    startDate,
    endDate,
    selectedMonth,
    selectedYear,
    useCustomRange,
    docNumber,
    docRecipientTitle,
    docThroughTitle,
    docSignDate,
    leftSignTitle,
    leftSignOfficer,
    leftSignShiftX,
    leftSignShiftY,
    rightSignTitle,
    rightSignOfficer,
    rightSignShiftX,
    rightSignShiftY,
    signSectionMarginTop,
    tableShiftY,
    paragraphShiftY,
    customMargins,
    activeTabStop,
    activeFont,
    activeFontSize,
    lineSpacing,
  ]);

  // Reset all edits back to template calculation
  const handleResetDocument = () => {
    if (window.confirm('តើអ្នកពិតជាចង់កំណត់ទិន្នន័យឯកសារឡើងវិញតាមទម្រង់ដើម និងលុបការកែប្រែដែលបានរក្សាទុកមែនទេ?')) {
      const storageKey = `${storagePrefix}${startDate}_${endDate}`;
      localStorage.removeItem(storageKey);
      localStorage.removeItem(latestDraftKey);
      idbStorage.removeItem('daily_team_operations', storageKey);
      idbStorage.removeItem('daily_team_operations', latestDraftKey);
      setHasSavedDraftForPeriod(false);
      setLastSavedTime(null);
      setKeyResetCounter((prev) => prev + 1);
    }
  };

  // Print Handler
  const handlePrint = () => {
    printA4Document('robok-total-stock-work-pdf', {
      orientation: 'portrait',
      scale: (adjustPagePercent || 97) / 100,
      documentTitle: isRolling
        ? `របក_សរុបរួមរំកិល_${startSolar.khmerYear}`
        : `របក_សរុបការងារស្តុកការិយាល័យ_${startSolar.khmerYear}`,
    });
  };

  // Export to Excel Handler
  const handleExportExcel = () => {
    try {
      const startKh = `${toKhmerNum(startSolar.day)} ${startSolar.month} ${toKhmerNum(startSolar.year)}`;
      const endKh = `${toKhmerNum(endSolar.day)} ${endSolar.month} ${toKhmerNum(endSolar.year)}`;

      let wsData: any[][] = [];

      if (viewMode === 'team') {
        wsData = [
          ['ព្រះរាជាណាចក្រកម្ពុជា'],
          ['ជាតិ សាសនា ព្រះមហាក្សត្រ'],
          [],
          ['ក្រសួងមហាផ្ទៃ'],
          ['អគ្គនាយកដ្ឋានអន្តោប្រវេសន៍'],
          ['នាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត'],
          ['ការិយាល័យទិដ្ឋាការចូល'],
          [teamDisplayName],
          [],
          [`របាយការណ៍ ស្តីពីការប្រើប្រាស់សន្លឹកទិដ្ឋាការស្អិត ${teamDisplayName}`],
          [`កាលបរិច្ឆេទ ៖ ពីថ្ងៃទី ${startKh} ដល់ ${endKh}`],
          [],
          [
            'ប្រភេទសន្លឹកទិដ្ឋាការ',
            ...VISA_TYPES,
            'សរុប',
          ],
          [
            'ចំនួនសល់ខែចាស់',
            ...VISA_TYPES.map((vt) => teamStickerData.opening.values[vt] || 0),
            teamStickerData.opening.total,
          ],
          [
            'ចំនួនបើកថ្មី',
            ...VISA_TYPES.map((vt) => teamStickerData.newIssued.values[vt] || 0),
            teamStickerData.newIssued.total,
          ],
          [
            'ផ្ទេរ/បង្វិល/ខូច/ខ្វះ',
            ...VISA_TYPES.map((vt) => teamStickerData.returnedDamagedMissing.values[vt] || 0),
            teamStickerData.returnedDamagedMissing.total,
          ],
          [
            'ចំនួនប្រើ',
            ...VISA_TYPES.map((vt) => teamStickerData.used.values[vt] || 0),
            teamStickerData.used.total,
          ],
          [
            'សល់ខែបន្ទាប់',
            ...VISA_TYPES.map((vt) => teamStickerData.ending.values[vt] || 0),
            teamStickerData.ending.total,
          ],
        ];
      } else {
        wsData = [
          ['ព្រះរាជាណាចក្រកម្ពុជា'],
          ['ជាតិ សាសនា ព្រះមហាក្សត្រ'],
          [],
          ['ក្រសួងមហាផ្ទៃ'],
          ['អគ្គនាយកដ្ឋានអន្តោប្រវេសន៍'],
          ['នាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត'],
          ['ការិយាល័យទិដ្ឋាការចូល'],
          ['ផ្នែករដ្ឋបាល'],
          [`លេខៈ ${docNumber}`],
          [],
          [`របាយការណ៍ស្តីពីការបើក ការផ្តល់សន្លឹកទិដ្ឋាការស្អិត និងក្រដាសអនុម័តផ្តល់ទិដ្ឋាការអេឡិចត្រូនិកនៅពេលមកដល់ ចាប់ពីថ្ងៃទី ${startKh} ដល់ថ្ងៃទី ${endKh}`],
          [],
          ['ក. សន្លឹកទិដ្ឋាការស្អិត'],
          [
            'ល.រ',
            'ប្រភេទ',
            `សន្និធិ ចុងគ្រា (${startKh}) - ក២`,
            `សន្និធិ ចុងគ្រា (${startKh}) - ក្រុម`,
            `ចលនា (${startKh} ដល់ ${endKh}) - ទិដ្ឋាការបើកពីក១`,
            `ចលនា (${startKh} ដល់ ${endKh}) - ទិដ្ឋាការបង្វិលពីក្រុម`,
            `ចលនា (${startKh} ដល់ ${endKh}) - ទិដ្ឋាការបើកផ្តល់ទៅក្រុម`,
            `ចលនា (${startKh} ដល់ ${endKh}) - ក្រុមប្រើប្រាស់`,
            `ចលនា (${startKh} ដល់ ${endKh}) - មិនបានការ ក២`,
            `ចលនា (${startKh} ដល់ ${endKh}) - ទិដ្ឋាការខូច/លុប ក២`,
            `ចលនា (${startKh} ដល់ ${endKh}) - ទិដ្ឋាការសាកល្បង ក២`,
            `សន្និធិសល់ (${endKh}) - ក២`,
            `សន្និធិសល់ (${endKh}) - ក្រុម`,
            `សន្និធិសរុប ក២និងក្រុម (${endKh})`,
          ],
        ];

        stickerData.rows.forEach((r, idx) => {
          wsData.push([
            idx + 1,
            r.visaType,
            r.openK2,
            r.openTeams,
            r.k1ToK2,
            r.teamReturnedToK2,
            r.issuedK2ToTeams,
            r.teamsUsed,
            r.damagedK2,
            r.damagedOrLostK2,
            r.testSampleK2,
            r.endingK2,
            r.endingTeams,
            r.grandTotalEnding,
          ]);
        });

        const tot = stickerData.totalRow;
        wsData.push([
          '',
          tot.visaType,
          tot.openK2,
          tot.openTeams,
          tot.k1ToK2,
          tot.teamReturnedToK2,
          tot.issuedK2ToTeams,
          tot.teamsUsed,
          tot.damagedK2,
          tot.damagedOrLostK2,
          tot.testSampleK2,
          tot.endingK2,
          tot.endingTeams,
          tot.grandTotalEnding,
        ]);

        wsData.push([]);
        wsData.push(['ខ. ក្រដាសអនុម័តទិដ្ឋាការអេឡិចត្រូនិក']);
        wsData.push([
          'ល.រ',
          'សន្និធិដើមគ្រា',
          'បញ្ចូលស្តុក',
          'បើកផ្តល់ទៅតាមក្រុមផ្តល់ទិដ្ឋាការអេឡិចត្រូនិច',
          'សន្និធិនៅសល់',
        ]);
        wsData.push([
          '១',
          eVisaData.openBundles,
          eVisaData.receivedK1Bundles,
          eVisaData.issuedToTeamsBundles,
          eVisaData.endingBundles,
        ]);
      }

      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.aoa_to_sheet(wsData);
      const sheetName = isRolling ? 'របក.សរុបរួមរំកិល' : 'របក.សរុបការងារស្តុក';
      const excelFileName = isRolling
        ? `របក_សរុបរួមរំកិល_${startDate}_${endDate}.xlsx`
        : `របក_សរុបការងារស្តុក_${startDate}_${endDate}.xlsx`;
      XLSX.utils.book_append_sheet(wb, ws, sheetName);
      XLSX.writeFile(wb, excelFileName);
    } catch (err) {
      console.error('Export Excel failed:', err);
    }
  };

  // Export to PDF
  const handleExportPDF = async () => {
    if (!printAreaRef.current) return;
    setIsExporting(true);
    const pdfFileName = isRolling
      ? `របក_សរុបរួមរំកិល_${startDate}_${endDate}.pdf`
      : `របក_សរុបការងារស្តុក_${startDate}_${endDate}.pdf`;
    try {
      await exportElementToPdf(
        printAreaRef.current,
        pdfFileName,
        { pixelRatio: 3, fitSinglePage: true, adjustPagePercent: adjustPagePercent || 97 }
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
          scale: 3,
          useCORS: true,
          allowTaint: true,
          logging: false,
          backgroundColor: '#ffffff',
          windowWidth: 1200,
          windowHeight: 1600,
          imageTimeout: 0,
          onclone: (clonedDoc) => {
            sanitizeDocumentForHtml2Canvas(clonedDoc, 'robok-total-stock-work-pdf');
            const clonedElement = clonedDoc.getElementById('robok-total-stock-work-pdf');
            if (clonedElement) {
              clonedElement.style.width = '210mm';
              clonedElement.style.minWidth = '210mm';
              clonedElement.style.maxWidth = '210mm';
              clonedElement.style.margin = '0 auto';
              clonedElement.style.boxShadow = 'none';
              clonedElement.style.backgroundColor = '#ffffff';
              clonedElement.style.textRendering = 'geometricPrecision';
              clonedElement.style.setProperty('-webkit-font-smoothing', 'antialiased');
              clonedElement.style.setProperty('-moz-osx-font-smoothing', 'grayscale');
              clonedElement.style.transform = 'none';
            }
          },
        });
        const imgData = canvas.toDataURL('image/png', 1.0);
        const pdf = new jsPDF({
          orientation: 'p',
          unit: 'mm',
          format: 'a4',
          compress: true,
        });
        const pageWidth = pdf.internal.pageSize.getWidth();
        const pageHeight = pdf.internal.pageSize.getHeight();
        const calculatedHeight = (canvas.height * pageWidth) / canvas.width;
        const scaleFactor = (adjustPagePercent || 97) / 100;
        const scaledWidth = pageWidth * scaleFactor;
        const scaledHeight = calculatedHeight * scaleFactor;
        const offsetX = (pageWidth - scaledWidth) / 2;

        pdf.addImage(imgData, 'PNG', offsetX, 0, scaledWidth, Math.min(pageHeight, scaledHeight), undefined, 'FAST');
        pdf.save(pdfFileName);
      } catch (fallbackErr) {
        console.error('Export PDF fallback failed:', fallbackErr);
      }
    } finally {
      setIsExporting(false);
    }
  };

  // Export to Microsoft Word (.doc)
  const handleExportWord = () => {
    try {
      const tacteingSettings = getSavedTacteingSettings();
      const wordFileName = isRolling
        ? `របក_សរុបរួមរំកិល_${startDate}_${endDate}`
        : viewMode === 'team'
        ? `របក_សរុបការងារស្តុក_${teamDisplayName}_${startDate}_${endDate}`
        : `របក_សរុបការងារស្តុក_${startDate}_${endDate}`;
      const wordReportTitle = isRolling
        ? 'របក.សរុបរួមរំកិល'
        : viewMode === 'team'
        ? `របក.សរុបការងារស្តុក ${teamDisplayName}`
        : 'របក.សរុបការងារស្តុក';
      exportRobokToWord({
        fileName: wordFileName,
        reportTitle: wordReportTitle,
        customMargins,
        activeTabStop,
        paragraphShiftY,
        customTacteingImage: tacteingSettings.type === 'custom-image' ? tacteingSettings.customImage : undefined,
        ministryHierarchy,
        docNumber,
        docRecipientRank,
        docRecipientTitle,
        docThroughTitle: cleanDocThroughTitle(docThroughTitle),
        docReferenceText,
        reportPeriodLabel,
        introSalutation: proposalType === 'office'
          ? `${docRecipientRank || 'លោកឧត្តមសេនីយ៍ទោ'} ${docRecipientTitle || 'ប្រធាននាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត'}`
          : docRecipientTitle,
        startKhmerDate: { day: startSolar.khmerDay, month: startSolar.khmerMonth, year: startSolar.khmerYear },
        endKhmerDate: { day: endSolar.khmerDay, month: endSolar.khmerMonth, year: endSolar.khmerYear },
        prevKhmerDate: { day: prevSolar.khmerDay, month: prevSolar.khmerMonth, year: prevSolar.khmerYear },
        signKhmerDate: { day: signSolar.khmerDay, month: signSolar.khmerMonth, year: signSolar.khmerYear },
        signLunarDate: customSignLunar || signLunar || 'ថ្ងៃ.....................ខែ.............ឆ្នាំ..............សំរឹទ្ធិស័ក ព.ស. ២៥៦....',
        signSolarDate:
          customSignSolar ||
          `${selectedTeam ? (selectedTeam.includes('កំពង់ផែ') || selectedTeam.includes('ព្រំដែន') || selectedTeam.includes('អាកាស') ? selectedTeam : 'ច្រកទ្វារ' + selectedTeam) : ''}, ថ្ងៃទី${signSolar.khmerDay} ខែ${signSolar.khmerMonth} ឆ្នាំ${signSolar.khmerYear}`,
        viewMode,
        teamDisplayName,
        teamPeriodKhmerTitle,
        teamPeriodShortKhmerTitle,
        teamStickerData,
        stickerRows: stickerData.rows,
        stickerTotalRow: stickerData.totalRow,
        eVisaData,
        leftSignTitle,
        leftSignOfficer,
        rightSignTitle: viewMode === 'team' ? `ប្រធាន${teamDisplayName}` : rightSignTitle,
        rightSignOfficer,
        variant: 'standard',
        proposalType,
        stubProposalRows: displayStubRows,
        isCalculated,
      });
    } catch (err) {
      console.error('Export Word failed:', err);
    }
  };

  return (
    <div className="space-y-4 pb-20 print:p-0 print:m-0 print:space-y-0">
      {/* Top Filter & Control Panel (Hidden on print) */}
      <div className="bg-white p-4 rounded-md border border-gray-200 shadow-2xs space-y-3 print:hidden">
        <div className="flex flex-col gap-3 border-b border-gray-100 pb-3">
          <div className="flex items-start gap-2.5 min-w-0">
            <div className="w-9 h-9 rounded bg-blue-50 text-[#007bff] flex items-center justify-center font-bold shrink-0 mt-0.5">
              <FileText className="w-5 h-5" />
            </div>
            <div>
              <h2 className="text-sm sm:text-base font-bold text-gray-800 leading-snug">
                {isRolling
                  ? 'របក.សរុបរួមរំកិល'
                  : proposalType === 'section'
                  ? 'ផ្នែកស្នើសុំប្រគល់គល់ទិដ្ឋាការ (ការិយាល័យ)'
                  : proposalType === 'office'
                  ? 'ការិយាល័យស្នើសុំប្រគល់គល់ទិដ្ឋាការ'
                  : viewMode === 'team'
                  ? 'ការងារស្តុកក្រុម'
                  : 'របក.ស្តុកសរុបរួម (ការិយាល័យ)'}
              </h2>
              <p className="text-xs text-gray-500 mt-0.5">
                {isRolling
                  ? 'របាយការណ៍សរុបរួមរំកិលការងារស្តុកទិដ្ឋាការស្អិត និងក្រដាសអនុម័ត'
                  : proposalType === 'section'
                  ? 'របាយការណ៍សរុបការងារស្តុកទិដ្ឋាការស្អិត និងក្រដាសអនុម័ត ប្រចាំការិយាល័យ'
                  : proposalType === 'office'
                  ? 'លិខិតការិយាល័យស្នើសុំប្រគល់គល់ទិដ្ឋាការដែលបានប្រើប្រាស់រួច'
                  : viewMode === 'team'
                  ? 'របាយការណ៍ការងារស្តុកតាមក្រុមនីមួយៗ (គំរូ ៥ ជួរ ស្របតាមទម្រង់ផ្លូវការ)'
                  : 'របាយការណ៍សរុបការងារស្តុកទិដ្ឋាការស្អិត និងក្រដាសអនុម័ត ប្រចាំការិយាល័យ'}
              </p>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-1.5 sm:gap-2">
            <button
              type="button"
              onClick={() => setIsWordEditMode(!isWordEditMode)}
              className={`px-3 py-1.5 rounded text-xs font-bold border transition flex items-center gap-1.5 cursor-pointer shadow-2xs ${
                isWordEditMode
                  ? 'bg-blue-600 border-blue-700 text-white'
                  : 'bg-gray-100 hover:bg-gray-200 border-gray-300 text-gray-700'
              }`}
              title="បើក/បិទ របៀបកែប្រែអក្សរផ្ទាល់លើក្រដាស"
            >
              <Edit3 className="w-3.5 h-3.5" />
              <span>{isWordEditMode ? 'កំពុងបើករបៀប Word' : 'បើករបៀប Word'}</span>
            </button>

            {/* Button to Hide / Show Word Ribbon Toolbar */}
            <button
              type="button"
              onClick={() => setShowRibbonToolbar(!showRibbonToolbar)}
              className={`px-3 py-1.5 rounded text-xs font-bold border transition flex items-center gap-1.5 cursor-pointer shadow-2xs ${
                showRibbonToolbar
                  ? 'bg-slate-800 hover:bg-slate-900 text-white border-slate-900'
                  : 'bg-white hover:bg-gray-100 text-gray-700 border-gray-300'
              }`}
              title={showRibbonToolbar ? 'ចុចដើម្បីលាក់របារឧបករណ៍ Word' : 'ចុចដើម្បីបង្ហាញរបារឧបករណ៍ Word'}
            >
              {showRibbonToolbar ? (
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

            <button
              type="button"
              onClick={handleExportExcel}
              className="px-3 py-1.5 rounded text-xs font-bold bg-emerald-600 hover:bg-emerald-700 text-white transition flex items-center gap-1.5 cursor-pointer shadow-2xs"
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              <span>ទាញយក Excel</span>
            </button>

            <button
              type="button"
              onClick={handleExportWord}
              className="px-3 py-1.5 rounded text-xs font-bold bg-blue-600 hover:bg-blue-700 text-white transition flex items-center gap-1.5 cursor-pointer shadow-2xs"
              title="ទាញយកឯកសារជា Microsoft Word (.doc) តាមទម្រង់ A4 ស្តង់ដារ"
            >
              <FileText className="w-3.5 h-3.5" />
              <span>ទាញយក Word</span>
            </button>

            <button
              type="button"
              onClick={handleExportPDF}
              disabled={isExporting}
              className="px-3.5 py-1.5 rounded text-xs font-bold bg-red-600 hover:bg-red-700 text-white transition flex items-center gap-1.5 cursor-pointer shadow-2xs disabled:opacity-50"
            >
              <Download className="w-3.5 h-3.5" />
              <span>{isExporting ? 'កំពុងបង្កើត PDF...' : 'ទាញយក PDF'}</span>
            </button>

            <button
              type="button"
              onClick={handlePrint}
              className="px-3.5 py-1.5 rounded text-xs font-bold bg-slate-800 hover:bg-slate-900 text-white transition flex items-center gap-1.5 cursor-pointer shadow-2xs"
              title="បោះពុម្ព ឬរក្សាទុកជា PDF (Print / Save as PDF)"
            >
              <Printer className="w-3.5 h-3.5 text-blue-300" />
              <span>បោះពុម្ព (Print)</span>
            </button>
          </div>
        </div>

        {/* Collapsible Tacteing Control Selector (The same as in EVisaStockReport) */}
        {showTacteingModal && (
          <div className="pt-2 border-t border-gray-100">
            <div className="bg-amber-50/40 border border-amber-200 rounded p-3">
              <div className="flex items-center justify-between border-b border-amber-200/80 pb-2 mb-3">
                <span className="font-bold text-amber-900 text-xs flex items-center gap-1.5">
                  <Sparkles className="w-4 h-4 text-amber-600" />
                  កំណត់ និងជ្រើសរើសរូបភាពតាក់តែងលិខិត (Tacteing Line Ornament)
                </span>
                <button
                  type="button"
                  onClick={() => setShowTacteingModal(false)}
                  className="text-gray-400 hover:text-gray-700 p-1 rounded hover:bg-amber-100/50 cursor-pointer text-xs"
                >
                  ✕
                </button>
              </div>

              <TacteingControlSelector
                currentType={tacteingSettings.type}
                customImage={tacteingSettings.customImage}
                onChange={(t, img) => {
                  setTacteingSettings({ type: t, customImage: img || null });
                  saveTacteingSettings(t, img);
                }}
              />
            </div>
          </div>
        )}



        {/* Date Filters & Document Tacteing Controls */}
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
              {KHMER_MONTHS_NAMES.map((m, idx) => (
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

          <div>
            <label className="block font-bold text-gray-700 mb-1 h-5 flex items-center text-xs truncate whitespace-nowrap">
              កាលបរិច្ឆេទចាប់ផ្តើម (From Date) :
            </label>
            <CustomDatePicker
              value={startDate}
              onChange={handleStartDateChange}
              className="h-[38px] min-h-[38px] max-h-[38px] font-sans text-xs font-medium text-gray-800 bg-white"
            />
          </div>

          <div>
            <label className="block font-bold text-gray-700 mb-1 h-5 flex items-center text-xs truncate whitespace-nowrap">
              កាលបរិច្ឆេទបញ្ចប់ (To Date) :
            </label>
            <CustomDatePicker
              value={endDate}
              onChange={handleEndDateChange}
              className="h-[38px] min-h-[38px] max-h-[38px] font-sans text-xs font-medium text-gray-800 bg-white"
            />
          </div>

          {/* Tacteing: Set Document Date */}
          <div>
            <label className="block font-bold text-gray-700 mb-1 h-5 flex items-center text-xs truncate whitespace-nowrap">
              📅 ថ្ងៃខែឆ្នាំតាក់តែងលិខិត :
            </label>
            <CustomDatePicker
              value={docSignDate}
              onChange={handleDocSignDateChange}
              className="h-[38px] min-h-[38px] max-h-[38px] font-sans text-xs bg-blue-50/40 text-blue-950 font-semibold"
              title="កំណត់ថ្ងៃខែឆ្នាំតាក់តែងលិខិត (ចុចដើម្បីជ្រើសរើស)"
            />
          </div>
        </div>

        {/* Quick Range Presets & Action Button */}
        <div className="mt-2 pt-2 border-t border-gray-100 flex flex-wrap items-center justify-between gap-2 text-xs">
          <div className="flex flex-wrap items-center gap-1.5">
            <span className="text-gray-500 font-bold">ចន្លោះកាលបរិច្ឆេទរហ័ស ៖</span>
            <button
              type="button"
              onClick={() => handlePresetRange('this_month')}
              className="px-2.5 py-1 rounded border border-gray-300 bg-gray-50 hover:bg-gray-100 text-gray-700 text-xs font-semibold transition cursor-pointer"
            >
              {isRolling ? 'ខែរំកិល' : 'ខែនេះ'}
            </button>
            <button
              type="button"
              onClick={() => handlePresetRange('last_month')}
              className="px-2.5 py-1 rounded border border-gray-300 bg-gray-50 hover:bg-gray-100 text-gray-700 text-xs font-semibold transition cursor-pointer"
            >
              {isRolling ? 'ខែរំកិលមុន' : 'ខែមុន'}
            </button>
            <button
              type="button"
              onClick={() => handlePresetRange('next_month')}
              className="px-2.5 py-1 rounded border border-gray-300 bg-gray-50 hover:bg-gray-100 text-gray-700 text-xs font-semibold transition cursor-pointer"
            >
              {isRolling ? 'ខែបន្ទាប់រំកិល' : 'ខែបន្ទាប់'}
            </button>
            <button
              type="button"
              onClick={() => handlePresetRange('q1')}
              className={`px-2.5 py-1 rounded border text-xs font-semibold transition cursor-pointer ${
                isPresetActive('q1')
                  ? 'bg-blue-600 border-blue-600 text-white font-bold shadow-xs'
                  : 'bg-gray-50 hover:bg-gray-100 border-gray-300 text-gray-700'
              }`}
              title={isRolling ? 'រំកិល ត្រីមាសទី១ (៣ខែ ចាប់ពីខែដែលបានជ្រើសរើស)' : 'ត្រីមាស (៣ខែ ចាប់ពីខែដែលបានជ្រើសរើស)'}
            >
              {isRolling ? 'រំកិល ត្រីមាសទី១' : 'ត្រីមាសទី១'}
            </button>
            <button
              type="button"
              onClick={() => handlePresetRange('s1')}
              className={`px-2.5 py-1 rounded border text-xs font-semibold transition cursor-pointer ${
                isPresetActive('s1')
                  ? 'bg-blue-600 border-blue-600 text-white font-bold shadow-xs'
                  : 'bg-gray-50 hover:bg-gray-100 border-gray-300 text-gray-700'
              }`}
              title={isRolling ? 'រំកិល ឆមាសទី១់ (៦ខែ ចាប់ពីខែដែលបានជ្រើសរើស)' : 'ឆមាស (៦ខែ ចាប់ពីខែដែលបានជ្រើសរើស)'}
            >
              {isRolling ? 'រំកិល ឆមាសទី១់' : 'ឆមាសទី១'}
            </button>
            <button
              type="button"
              onClick={() => handlePresetRange('m9')}
              className={`px-2.5 py-1 rounded border text-xs font-semibold transition cursor-pointer ${
                isPresetActive('m9')
                  ? 'bg-blue-600 border-blue-600 text-white font-bold shadow-xs'
                  : 'bg-gray-50 hover:bg-gray-100 border-gray-300 text-gray-700'
              }`}
              title={isRolling ? 'រំកិល នព្វមាស (៩ខែ ចាប់ពីខែដែលបានជ្រើសរើស)' : 'នព្វមាស (៩ខែ ចាប់ពីខែដែលបានជ្រើសរើស)'}
            >
              {isRolling ? 'រំកិល នព្វមាស' : 'នព្វមាស'}
            </button>
            <button
              type="button"
              onClick={() => handlePresetRange('full_year')}
              className={`px-2.5 py-1 rounded border text-xs font-semibold transition cursor-pointer ${
                isPresetActive('full_year')
                  ? 'bg-blue-600 border-blue-600 text-white font-bold shadow-xs'
                  : 'bg-gray-50 hover:bg-gray-100 border-gray-300 text-gray-700'
              }`}
              title={isRolling ? 'រំកិលប្រចាំ ១២ខែ (១២ខែ ចាប់ពីខែដែលបានជ្រើសរើស)' : 'រយៈពេល ១២ខែ (១២ខែ ចាប់ពីខែដែលបានជ្រើសរើស)'}
            >
              {isRolling ? 'រំកិលប្រចាំ ១២ខែ' : 'រយៈពេល ១២ខែ'}
            </button>
          </div>

          {/* Show Result Action Button */}
          <button
            type="button"
            onClick={handleShowResult}
            className={`px-4 py-1.5 rounded font-bold text-xs flex items-center justify-center gap-1.5 transition cursor-pointer shadow-sm ${
              !isCalculated
                ? 'bg-blue-600 hover:bg-blue-700 active:scale-95 text-white ring-2 ring-blue-400 font-moul tracking-wide shadow-md'
                : 'bg-emerald-600 hover:bg-emerald-700 active:scale-95 text-white'
            }`}
            title="ចុចដើម្បីទាញយក និងបង្ហាញលទ្ធផលរបាយការណ៍"
          >
            <Search className="w-3.5 h-3.5" />
            <span>{!isCalculated ? 'បង្ហាញលទ្ធផល (Show Result)' : 'បង្ហាញលទ្ធផលរួចរាល់ (Refresh)'}</span>
          </button>
        </div>
      </div>

      {/* Microsoft Word-Style Floating / Fixed Ribbon Toolbar (Hidden on print) */}
      {showRibbonToolbar && (
        <div className="print:hidden sticky top-3 z-30 bg-slate-800 text-white rounded-xl p-2 shadow-2xl border border-slate-700 space-y-2 text-xs">
          {/* Main Ribbon Tool Groups */}
          <div className="flex flex-wrap items-center gap-1.5">
            {/* Quick Access Save (Floppy Disk) */}
            <button
              onClick={handleSaveDocument}
              className="px-2.5 py-1.5 bg-blue-600 hover:bg-blue-500 rounded-lg text-white font-bold transition flex items-center gap-1.5 cursor-pointer shadow border border-blue-400"
              title="រក្សាទុកឯកសារ (Save - Ctrl+S)"
            >
              <Save className="w-3.5 h-3.5 text-blue-100" />
              <span className="text-[11px]">រក្សាទុក</span>
            </button>

            {/* Quick Hide Toolbar Button */}
            <button
              type="button"
              onClick={() => setShowRibbonToolbar(false)}
              className="px-2 py-1.5 bg-slate-900/90 hover:bg-slate-700 rounded-lg text-amber-300 hover:text-white transition flex items-center gap-1 cursor-pointer border border-slate-700 text-[11px] font-bold"
              title="ចុចដើម្បីលាក់របារឧបករណ៍នេះ (Hide Toolbar)"
            >
              <ChevronUp className="w-3.5 h-3.5" />
              <span>លាក់របារ</span>
            </button>

          <div className="h-5 w-[1px] bg-slate-700 mx-0.5" />

          {/* History / Undo / Redo */}
          <div className="flex items-center bg-slate-900/90 rounded-lg p-0.5 border border-slate-700">
            <button
              onClick={() => executeCommand('undo')}
              className="p-1.5 hover:bg-slate-700 rounded text-gray-300 hover:text-white transition cursor-pointer"
              title="ត្រឡប់ក្រោយ (Undo - Ctrl+Z)"
            >
              <Undo className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={() => executeCommand('redo')}
              className="p-1.5 hover:bg-slate-700 rounded text-gray-300 hover:text-white transition cursor-pointer"
              title="ទៅមុខវិញ (Redo - Ctrl+Y)"
            >
              <Redo className="w-3.5 h-3.5" />
            </button>
          </div>

          <div className="h-5 w-[1px] bg-slate-700 mx-0.5" />

          {/* Font Family Selector */}
          <div className="flex items-center gap-1 bg-slate-900/90 px-2 py-1 rounded-lg border border-slate-700">
            <Type className="w-3.5 h-3.5 text-amber-400 shrink-0" />
            <select
              value={activeFont}
              onChange={(e) => handleApplyFont(e.target.value)}
              className="bg-transparent text-white text-xs font-medium focus:outline-none cursor-pointer pr-1"
              title="ប្តូរពុម្ពអក្សរ (Font Family)"
            >
              <option value="Khmer OS Siemreap" className="text-black">Khmer OS Siemreap (អក្សរជ្រុងស្តង់ដារ)</option>
              <option value="Khmer Mool1" className="text-black">Khmer Mool1 (អក្សរមូលចំណងជើង)</option>
              <option value="Khmer OS Mool1" className="text-black">Khmer OS Mool1</option>
              <option value="Khmer OS Battambang" className="text-black">Khmer OS Battambang</option>
              <option value="Khmer OS Bokor" className="text-black">Khmer OS Bokor</option>
              <option value="Khmer OS Freehand" className="text-black">Khmer OS Freehand</option>
              <option value="Kantumruy Pro" className="text-black">Kantumruy Pro</option>
              <option value="Times New Roman" className="text-black">Times New Roman (លេខ & អង់គ្លេស)</option>
              <option value="Arial" className="text-black">Arial</option>
              <option value="Calibri" className="text-black">Calibri</option>
            </select>
          </div>

          {/* Font Size Selector & Grow / Shrink */}
          <div className="flex items-center bg-slate-900/90 rounded-lg p-0.5 border border-slate-700">
            <select
              value={activeFontSize}
              onChange={(e) => handleApplyFontSize(e.target.value)}
              className="bg-transparent text-white text-xs font-medium focus:outline-none cursor-pointer px-1.5 py-0.5"
              title="ទំហំអក្សរ (Font Size)"
            >
              {FONT_SIZE_STEPS.map((size) => (
                <option key={size} value={size} className="text-black">
                  {size}
                </option>
              ))}
            </select>
            <button
              onClick={handleGrowFontSize}
              className="p-1 hover:bg-slate-700 rounded text-gray-200 hover:text-white transition cursor-pointer font-bold text-[11px] px-1.5"
              title="ពង្រីកទំហំអក្សរ (Grow Font A+)"
            >
              A<span className="text-[9px] align-top font-bold">+</span>
            </button>
            <button
              onClick={handleShrinkFontSize}
              className="p-1 hover:bg-slate-700 rounded text-gray-200 hover:text-white transition cursor-pointer font-bold text-[11px] px-1.5"
              title="បង្រួមទំហំអក្សរ (Shrink Font A-)"
            >
              A<span className="text-[9px] align-top font-bold">-</span>
            </button>
          </div>

          {/* Microsoft Word Font Dialog / Condensed Launcher (Ctrl+D) */}
          <button
            onClick={() => {
              const sel = window.getSelection();
              if (sel && sel.rangeCount > 0 && printAreaRef.current && printAreaRef.current.contains(sel.getRangeAt(0).commonAncestorContainer)) {
                savedSelectionRef.current = sel.getRangeAt(0).cloneRange();
              }
              setFontDialogTab('advanced');
              setShowFontDialog(true);
            }}
            className="px-2 py-1 hover:bg-slate-700 rounded text-amber-300 hover:text-white transition flex items-center gap-1.5 cursor-pointer bg-slate-900/90 border border-slate-700 font-bold"
            title="Microsoft Word Font Dialog & Condensed Character Spacing (Ctrl+D)"
          >
            <Sliders className="w-3.5 h-3.5 text-amber-400" />
            <span className="text-[11px]">Condensed (Ctrl+D)</span>
          </button>

          <div className="h-5 w-[1px] bg-slate-700 mx-0.5" />

          {/* Basic Text Formatting (B, I, U, S, Sub, Sup) */}
          <div className="flex items-center bg-slate-900/90 rounded-lg p-0.5 border border-slate-700">
            <button
              onClick={() => executeCommand('bold')}
              className="p-1.5 hover:bg-slate-700 rounded text-gray-200 hover:text-white transition cursor-pointer"
              title="អក្សរដិត (Bold - Ctrl+B)"
            >
              <Bold className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={() => executeCommand('italic')}
              className="p-1.5 hover:bg-slate-700 rounded text-gray-200 hover:text-white transition cursor-pointer"
              title="អក្សរទ្រេត (Italic - Ctrl+I)"
            >
              <Italic className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={() => executeCommand('underline')}
              className="p-1.5 hover:bg-slate-700 rounded text-gray-200 hover:text-white transition cursor-pointer"
              title="គូសបន្ទាត់ក្រោម (Underline - Ctrl+U)"
            >
              <Underline className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={() => executeCommand('strikeThrough')}
              className="p-1.5 hover:bg-slate-700 rounded text-gray-200 hover:text-white transition cursor-pointer"
              title="គូសបន្ទាត់ចំកណ្តាលអក្សរ (Strikethrough)"
            >
              <Strikethrough className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={() => executeCommand('subscript')}
              className="p-1.5 hover:bg-slate-700 rounded text-gray-200 hover:text-white transition cursor-pointer"
              title="អក្សរជើងក្រោម (Subscript X₂)"
            >
              <Subscript className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={() => executeCommand('superscript')}
              className="p-1.5 hover:bg-slate-700 rounded text-gray-200 hover:text-white transition cursor-pointer"
              title="អក្សរសន្ទស្សន៍លើ (Superscript X²)"
            >
              <Superscript className="w-3.5 h-3.5" />
            </button>
          </div>

          {/* Text Color & Highlight Dropdowns */}
          <div className="relative flex items-center bg-slate-900/90 rounded-lg p-0.5 border border-slate-700">
            {/* Font Color */}
            <div className="relative">
              <button
                onClick={() => {
                  setShowColorPicker(!showColorPicker);
                  setShowHighlightPicker(false);
                }}
                className="p-1.5 hover:bg-slate-700 rounded text-gray-200 hover:text-white transition flex items-center gap-0.5 cursor-pointer"
                title="ពណ៌អក្សរ (Font Color)"
              >
                <Palette className="w-3.5 h-3.5 text-red-400" />
                <span className="text-[10px] font-bold">A</span>
              </button>
              {showColorPicker && (
                <div className="absolute top-full left-0 mt-1.5 bg-slate-900 border border-slate-700 shadow-2xl rounded-lg p-2 z-50 w-44 grid grid-cols-5 gap-1.5">
                  {[
                    { name: 'ខ្មៅ', hex: '#000000' },
                    { name: 'ប្រផេះ', hex: '#4B5563' },
                    { name: 'ខៀវចាស់', hex: '#1E3A8A' },
                    { name: 'ខៀវ', hex: '#2563EB' },
                    { name: 'ក្រហម', hex: '#DC2626' },
                    { name: 'ត្នោត', hex: '#78350F' },
                    { name: 'បៃតង', hex: '#16A34A' },
                    { name: 'ស្វាយ', hex: '#9333EA' },
                    { name: 'ទឹកក្រូច', hex: '#D97706' },
                    { name: 'ស', hex: '#FFFFFF' },
                  ].map((c) => (
                    <button
                      key={c.hex}
                      onClick={() => handleApplyTextColor(c.hex)}
                      className="w-6 h-6 rounded border border-slate-600 hover:scale-110 transition cursor-pointer shadow-sm"
                      style={{ backgroundColor: c.hex }}
                      title={c.name}
                    />
                  ))}
                </div>
              )}
            </div>

            {/* Highlighter */}
            <div className="relative">
              <button
                onClick={() => {
                  setShowHighlightPicker(!showHighlightPicker);
                  setShowColorPicker(false);
                }}
                className="p-1.5 hover:bg-slate-700 rounded text-gray-200 hover:text-white transition flex items-center gap-0.5 cursor-pointer"
                title="ហាយឡាយពណ៌ផ្ទៃអក្សរ (Text Highlight Color)"
              >
                <Highlighter className="w-3.5 h-3.5 text-yellow-400" />
              </button>
              {showHighlightPicker && (
                <div className="absolute top-full left-0 mt-1.5 bg-slate-900 border border-slate-700 shadow-2xl rounded-lg p-2 z-50 w-44 space-y-1.5">
                  <div className="grid grid-cols-5 gap-1.5">
                    {[
                      { name: 'លឿង', hex: '#FEF08A' },
                      { name: 'បៃតងខ្ចី', hex: '#BBF7D0' },
                      { name: 'ផ្ទៃមេឃ', hex: '#A5F3FC' },
                      { name: 'ផ្កាឈូក', hex: '#FBCFE8' },
                      { name: 'ទឹកក្រូចខ្ចី', hex: '#FDE68A' },
                    ].map((c) => (
                      <button
                        key={c.hex}
                        onClick={() => handleApplyHighlight(c.hex)}
                        className="w-6 h-6 rounded border border-slate-600 hover:scale-110 transition cursor-pointer"
                        style={{ backgroundColor: c.hex }}
                        title={c.name}
                      />
                    ))}
                  </div>
                  <button
                    onClick={() => handleApplyHighlight('transparent')}
                    className="w-full text-[10px] text-center text-gray-300 hover:text-white hover:bg-slate-800 py-1 rounded transition cursor-pointer border border-slate-700"
                  >
                    គ្មានហាយឡាយ (No Color)
                  </button>
                </div>
              )}
            </div>

            {/* Clear Formatting */}
            <button
              onClick={handleClearFormatting}
              className="p-1.5 hover:bg-slate-700 rounded text-gray-300 hover:text-white transition cursor-pointer"
              title="លុបទម្រង់អក្សរដែលបានកែ (Clear Formatting)"
            >
              <RemoveFormatting className="w-3.5 h-3.5" />
            </button>
          </div>

          <div className="h-5 w-[1px] bg-slate-700 mx-0.5" />

          {/* Alignment */}
          <div className="flex items-center bg-slate-900/90 rounded-lg p-0.5 border border-slate-700">
            <button
              onClick={() => executeCommand('justifyLeft')}
              className="p-1.5 hover:bg-slate-700 rounded text-gray-200 hover:text-white transition cursor-pointer"
              title="តម្រឹមឆ្វេង (Align Left - Ctrl+L)"
            >
              <AlignLeft className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={() => executeCommand('justifyCenter')}
              className="p-1.5 hover:bg-slate-700 rounded text-gray-200 hover:text-white transition cursor-pointer"
              title="តម្រឹមកណ្តាល (Align Center - Ctrl+E)"
            >
              <AlignCenter className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={() => executeCommand('justifyRight')}
              className="p-1.5 hover:bg-slate-700 rounded text-gray-200 hover:text-white transition cursor-pointer"
              title="តម្រឹមស្តាំ (Align Right - Ctrl+R)"
            >
              <AlignRight className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={() => executeCommand('justifyFull')}
              className="p-1.5 hover:bg-slate-700 rounded text-gray-200 hover:text-white transition cursor-pointer"
              title="តម្រឹមសងខាងស្មើ (Justify - Ctrl+J)"
            >
              <AlignJustify className="w-3.5 h-3.5" />
            </button>
          </div>

          {/* Line Spacing */}
          <div className="flex items-center gap-1 bg-slate-900/90 px-1.5 py-1 rounded-lg border border-slate-700">
            <span className="text-[10px] text-gray-400">គម្លាត:</span>
            <select
              value={lineSpacing}
              onChange={(e) => handleApplyLineSpacing(e.target.value)}
              className="bg-transparent text-white text-xs font-medium focus:outline-none cursor-pointer"
              title="គម្លាតចន្លោះបន្ទាត់ (Line Spacing)"
            >
              <option value="1.0" className="text-black">1.0 (Single)</option>
              <option value="1.15" className="text-black">1.15</option>
              <option value="1.25" className="text-black">1.25</option>
              <option value="1.5" className="text-black">1.5 (Standard)</option>
              <option value="2.0" className="text-black">2.0 (Double)</option>
            </select>
          </div>

          {/* Lists & Indents */}
          <div className="flex items-center bg-slate-900/90 rounded-lg p-0.5 border border-slate-700">
            <button
              onClick={() => executeCommand('insertUnorderedList')}
              className="p-1.5 hover:bg-slate-700 rounded text-gray-200 hover:text-white transition cursor-pointer"
              title="បញ្ជីចំណុច (Bullets List)"
            >
              <List className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={() => executeCommand('insertOrderedList')}
              className="p-1.5 hover:bg-slate-700 rounded text-gray-200 hover:text-white transition cursor-pointer"
              title="បញ្ជីលេខរៀង (Numbered List)"
            >
              <ListOrdered className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={() => executeCommand('outdent')}
              className="p-1.5 hover:bg-slate-700 rounded text-gray-200 hover:text-white transition cursor-pointer"
              title="បន្ថយការដកឃ្លាបន្ទាត់ (Decrease Indent)"
            >
              <Outdent className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={() => executeCommand('indent')}
              className="p-1.5 hover:bg-slate-700 rounded text-gray-200 hover:text-white transition cursor-pointer"
              title="បង្កើនការដកឃ្លាបន្ទាត់ (Increase Indent)"
            >
              <Indent className="w-3.5 h-3.5" />
            </button>
          </div>

          {/* Tap Space / Indent / Tab Button with Ruler Tab Stop sync */}
          <div className="flex items-center gap-1 bg-amber-500/20 text-amber-200 border border-amber-500/40 rounded-lg px-2 py-1">
            <button
              onClick={() => handleInsertTab()}
              className="hover:text-white font-bold text-xs flex items-center gap-1 cursor-pointer"
              title={`ចុច Tab ដកឃ្លាតាមខ្នាត Ruler (${activeTabStop} cm)`}
            >
              <span>⇥ Tab ({activeTabStop} cm)</span>
            </button>
            <span className="text-amber-500/60">|</span>
            <select
              value={activeTabStop}
              onChange={(e) => setActiveTabStop(parseFloat(e.target.value))}
              className="bg-transparent text-amber-200 text-xs font-semibold focus:outline-none cursor-pointer"
              title="ជ្រើសរើសចម្ងាយ Tab Stop"
            >
              <option value="0.5" className="text-black">Tab 0.5 cm</option>
              <option value="1.0" className="text-black">Tab 1.0 cm</option>
              <option value="1.5" className="text-black">Tab 1.5 cm</option>
              <option value="2.0" className="text-black">Tab 2.0 cm (ស្តង់ដារ)</option>
              <option value="2.25" className="text-black">Tab 2.25 cm</option>
              <option value="2.3" className="text-black">Tab 2.30 cm</option>
              <option value="2.5" className="text-black">Tab 2.5 cm</option>
              <option value="2.55" className="text-black">Tab 2.55 cm</option>
              <option value="2.6" className="text-black">Tab 2.6 cm</option>
              <option value="3.0" className="text-black">Tab 3.0 cm</option>
              <option value="4.0" className="text-black">Tab 4.0 cm</option>
            </select>
            <span className="text-amber-500/60">|</span>
            <button
              onClick={() => handleInsertSpace(1)}
              className="hover:text-white font-medium text-xs px-1 cursor-pointer"
              title="បញ្ចូល ១ Space"
            >
              +1
            </button>
            <button
              onClick={() => handleInsertSpace(2)}
              className="hover:text-white font-medium text-xs px-1 cursor-pointer"
              title="បញ្ចូល ២ Space"
            >
              +2
            </button>
          </div>

          {/* Insert Tools: Date & Divider */}
          <div className="flex items-center bg-slate-900/90 rounded-lg p-0.5 border border-slate-700">
            <button
              onClick={handleInsertDateStamp}
              className="px-2 py-1 hover:bg-slate-700 rounded text-gray-200 hover:text-white transition flex items-center gap-1 cursor-pointer text-[11px]"
              title="បញ្ចូលកាលបរិច្ឆេទថ្ងៃនេះ (Insert Date Stamp)"
            >
              <Calendar className="w-3 h-3 text-emerald-400" />
              <span>កាលបរិច្ឆេទ</span>
            </button>
            <button
              onClick={handleInsertHorizontalRule}
              className="px-2 py-1 hover:bg-slate-700 rounded text-gray-200 hover:text-white transition flex items-center gap-1 cursor-pointer text-[11px]"
              title="បញ្ចូលបន្ទាត់ផ្តេកខណ្ឌ (Insert Horizontal Line)"
            >
              <Minus className="w-3 h-3 text-cyan-400" />
              <span>បន្ទាត់ខណ្ឌ</span>
            </button>
          </div>

          {/* Find & Replace Toggle */}
          <button
            onClick={() => setShowFindReplace(!showFindReplace)}
            className={`p-1.5 rounded-lg border transition flex items-center gap-1 cursor-pointer ${
              showFindReplace
                ? 'bg-purple-600 border-purple-500 text-white'
                : 'bg-slate-900/90 border-slate-700 text-gray-300 hover:text-white'
            }`}
            title="ស្វែងរក និងជំនួសពាក្យ (Find & Replace - Ctrl+F / Ctrl+H)"
          >
            <Search className="w-3.5 h-3.5" />
            <span className="text-[11px]">ស្វែងរក / ជំនួស</span>
          </button>

          {/* Ruler Toggle */}
          <button
            onClick={() => setShowRuler(!showRuler)}
            className={`p-1.5 rounded-lg border transition flex items-center gap-1 cursor-pointer ${
              showRuler
                ? 'bg-emerald-600 border-emerald-500 text-white shadow-sm'
                : 'bg-slate-900/90 border-slate-700 text-gray-300 hover:text-white'
            }`}
            title="បង្ហាញ/លាក់ បន្ទាត់វាស់ខ្នាត Ruler (Microsoft Word View > Ruler)"
          >
            <Maximize2 className="w-3.5 h-3.5" />
            <span className="text-[11px] font-medium">Ruler {showRuler ? '✓' : ''}</span>
          </button>

          {/* Margin Controls Toggle */}
          <button
            onClick={() => setShowMarginControls(!showMarginControls)}
            className={`p-1.5 rounded-lg border transition flex items-center gap-1 cursor-pointer ${
              showMarginControls
                ? 'bg-blue-600 border-blue-500 text-white'
                : 'bg-slate-900/90 border-slate-700 text-gray-300 hover:text-white'
            }`}
            title="កែសម្រួលខ្នាតគែម A4 (Quick Margins)"
          >
            <Sliders className="w-3.5 h-3.5" />
            <span className="text-[11px]">ខ្នាតគែម</span>
          </button>

          {/* Quick Top Margin Stepper (ទម្លាក់គម្លាតពីលើចុះក្រោម) */}
          <div className="flex items-center bg-slate-900/90 border border-slate-700 rounded-lg px-2 py-1 text-xs text-gray-200">
            <span className="text-[10.5px] text-gray-400 mr-1.5 select-none font-siemreap">គែមលើ៖</span>
            <button
              type="button"
              onClick={() => {
                const newTop = Math.max(0.2, parseFloat((customMargins.top - 0.1).toFixed(2)));
                setCustomMargins({ ...customMargins, top: newTop });
              }}
              className="w-5 h-5 flex items-center justify-center rounded bg-slate-800 hover:bg-slate-700 text-gray-200 font-bold cursor-pointer transition"
              title="បន្ថយគម្លាតពីលើ (-0.1cm)"
            >
              -
            </button>
            <span className="font-mono text-xs font-bold px-1.5 text-blue-300 select-none">
              {customMargins.top.toFixed(1)}cm
            </span>
            <button
              type="button"
              onClick={() => {
                const newTop = Math.min(3.5, parseFloat((customMargins.top + 0.1).toFixed(2)));
                setCustomMargins({ ...customMargins, top: newTop });
              }}
              className="w-5 h-5 flex items-center justify-center rounded bg-slate-800 hover:bg-slate-700 text-gray-200 font-bold cursor-pointer transition"
              title="ទម្លាក់គម្លាតពីលើចុះ (+0.1cm)"
            >
              +
            </button>
          </div>

          {/* Quick Paragraph Shift Stepper (ទម្លាក់កថាខណ្ឌចុះក្រោម) */}
          <div className="flex items-center bg-slate-900/90 border border-amber-700/60 rounded-lg px-2 py-1 text-xs text-gray-200 shadow-xs">
            <span className="text-[10.5px] text-amber-300 mr-1.5 select-none font-siemreap font-medium">កថាខណ្ឌ៖</span>
            <button
              type="button"
              onClick={() => handleParagraphShiftChange(Math.max(0, paragraphShiftY - 2))}
              className="w-5 h-5 flex items-center justify-center rounded bg-slate-800 hover:bg-amber-600 text-gray-200 font-bold cursor-pointer transition"
              title="រំកិលកថាខណ្ឌឡើងលើ (-2pt)"
            >
              -
            </button>
            <span className="font-mono text-xs font-bold px-1.5 text-amber-300 select-none min-w-[34px] text-center">
              {paragraphShiftY}pt
            </span>
            <button
              type="button"
              onClick={() => handleParagraphShiftChange(Math.min(60, paragraphShiftY + 2))}
              className="w-5 h-5 flex items-center justify-center rounded bg-slate-800 hover:bg-amber-600 text-gray-200 font-bold cursor-pointer transition"
              title="ទម្លាក់កថាខណ្ឌចុះក្រោម (+2pt)"
            >
              +
            </button>
          </div>

          {/* Page Setup Dialog (Microsoft Word Page Setup - Alt+P+S+P) */}
          <button
            onClick={() => {
              setTempMargins({ ...customMargins });
              setTempAdjustPagePercent(adjustPagePercent);
              setShowPageSetupDialog(true);
            }}
            className="px-2.5 py-1 rounded-lg border bg-gradient-to-r from-blue-700 to-indigo-700 hover:from-blue-600 hover:to-indigo-600 border-blue-400 text-white font-medium shadow flex items-center gap-1.5 cursor-pointer text-[11px] transition"
            title="បើកផ្ទាំងកំណត់ទំព័រ (Microsoft Word Page Setup Dialog)"
          >
            <Layout className="w-3.5 h-3.5 text-blue-200" />
            <span className="font-semibold">Page Setup</span>
          </button>

          {/* Quick Adjust Page Stepper (ពង្រីក/បង្រួមទំព័រ 97%) */}
          <div className="flex items-center bg-slate-900/90 border border-blue-500/50 rounded-lg px-2 py-1 text-xs text-gray-200 shadow-xs" title="ពង្រីក/បង្រួមទំព័រពេលបោះពុម្ព និងទាញយក PDF (Adjust Page: 97%)">
            <span className="text-[10.5px] text-blue-300 mr-1.5 select-none font-siemreap font-medium">ទំព័រ៖</span>
            <button
              type="button"
              onClick={() => {
                const next = Math.max(50, adjustPagePercent - 1);
                setAdjustPagePercent(next);
                try {
                  localStorage.setItem('robok_total_adjust_scale', next.toString());
                  const prefStr = localStorage.getItem('robok_report_user_preferences');
                  const pref = prefStr ? JSON.parse(prefStr) : {};
                  pref.adjustPagePercent = next;
                  localStorage.setItem('robok_report_user_preferences', JSON.stringify(pref));
                } catch (e) {}
              }}
              className="w-5 h-5 flex items-center justify-center rounded bg-slate-800 hover:bg-blue-600 text-gray-200 font-bold cursor-pointer transition"
              title="បន្ថយទំហំទំព័រ (-1%)"
            >
              -
            </button>
            <span className="font-mono text-xs font-bold px-1.5 text-blue-300 select-none min-w-[34px] text-center">
              {adjustPagePercent}%
            </span>
            <button
              type="button"
              onClick={() => {
                const next = Math.min(120, adjustPagePercent + 1);
                setAdjustPagePercent(next);
                try {
                  localStorage.setItem('robok_total_adjust_scale', next.toString());
                  const prefStr = localStorage.getItem('robok_report_user_preferences');
                  const pref = prefStr ? JSON.parse(prefStr) : {};
                  pref.adjustPagePercent = next;
                  localStorage.setItem('robok_report_user_preferences', JSON.stringify(pref));
                } catch (e) {}
              }}
              className="w-5 h-5 flex items-center justify-center rounded bg-slate-800 hover:bg-blue-600 text-gray-200 font-bold cursor-pointer transition"
              title="បង្កើនទំហំទំព័រ (+1%)"
            >
              +
            </button>
          </div>

          {/* Document Date Manager Button */}
          <button
            onClick={() => setShowDocDateModal(true)}
            className="px-2.5 py-1 rounded-lg border bg-gradient-to-r from-amber-600 to-orange-600 hover:from-amber-500 hover:to-orange-500 border-amber-400 text-white font-medium shadow flex items-center gap-1.5 cursor-pointer text-[11px] transition"
            title="កែប្រែថ្ងៃខែឆ្នាំតាក់តែងលិខិត (Edit Document Creation Date)"
          >
            <span>📅</span>
            <span className="font-semibold">ថ្ងៃតាក់តែង ៖ {signSolar.khmerDay}/{signSolar.khmerMonth}/{signSolar.khmerYear}</span>
          </button>

          {/* Tacteing Ornament Settings Button */}
          <div className="relative">
            <button
              onClick={() => setShowTacteingModal(!showTacteingModal)}
              className={`p-1.5 rounded-lg border transition flex items-center gap-1.5 cursor-pointer text-[11px] font-semibold ${
                showTacteingModal
                  ? 'bg-amber-600 border-amber-500 text-white shadow-md'
                  : 'bg-slate-900/90 border-slate-700 text-amber-300 hover:text-white hover:border-amber-400'
              }`}
              title="កំណត់ និងជ្រើសរើសរូបតាក់តែងបន្ទាត់ (Tacteing Line Ornament Settings)"
            >
              <Sparkles className="w-3.5 h-3.5 text-amber-400" />
              <span>រូបតាក់តែង</span>
            </button>

            {showTacteingModal && (
              <div className="absolute top-full right-0 mt-2 bg-slate-900 border border-amber-500/60 shadow-2xl rounded-xl p-4 z-50 w-96 space-y-3 text-xs text-white">
                <div className="flex items-center justify-between border-b border-slate-800 pb-2">
                  <span className="font-bold text-amber-400 flex items-center gap-1.5">
                    <Sparkles className="w-4 h-4 text-amber-400" />
                    កំណត់រូបភាពតាក់តែងលិខិត
                  </span>
                  <button
                    onClick={() => setShowTacteingModal(false)}
                    className="text-gray-400 hover:text-white p-1 rounded hover:bg-slate-800 cursor-pointer"
                  >
                    ✕
                  </button>
                </div>

                <TacteingControlSelector
                  currentType={tacteingSettings.type}
                  customImage={tacteingSettings.customImage}
                  onChange={(t, img) => {
                    setTacteingSettings({ type: t, customImage: img || null });
                    saveTacteingSettings(t, img);
                  }}
                />

                <div className="pt-1 flex justify-end">
                  <button
                    type="button"
                    onClick={() => setShowTacteingModal(false)}
                    className="px-3 py-1.5 bg-blue-600 hover:bg-blue-500 text-white font-bold rounded text-xs cursor-pointer shadow"
                  >
                    រួចរាល់
                  </button>
                </div>
              </div>
            )}
          </div>

          {/* Signature Position Controls Toggle */}
          <div className="relative">
            <button
              onClick={() => setShowSignPositionPanel(!showSignPositionPanel)}
              className={`p-1.5 rounded-lg border transition flex items-center gap-1 cursor-pointer ${
                showSignPositionPanel
                  ? 'bg-amber-600 border-amber-500 text-white shadow-md'
                  : 'bg-slate-900/90 border-slate-700 text-amber-300 hover:text-white'
              }`}
              title="កំណត់ទីតាំង និងរំកិលហត្ថលេខា (Signature Position Controls)"
            >
              <Move className="w-3.5 h-3.5 text-amber-400" />
              <span className="text-[11px] font-semibold">រំកិលហត្ថលេខា</span>
            </button>

            {showSignPositionPanel && (
              <div className="absolute top-full right-0 mt-2 bg-slate-900 border border-amber-500/50 shadow-2xl rounded-xl p-4 z-50 w-80 space-y-3 text-xs text-white">
                <div className="flex items-center justify-between border-b border-slate-800 pb-2">
                  <span className="font-bold text-amber-400 flex items-center gap-1.5">
                    <Move className="w-4 h-4" />
                    រំកិលទីតាំងហត្ថលេខា
                  </span>
                  <button
                    onClick={() => setShowSignPositionPanel(false)}
                    className="text-gray-400 hover:text-white p-1 rounded hover:bg-slate-800 cursor-pointer"
                  >
                    ✕
                  </button>
                </div>

                {/* Left Signature Controls */}
                <div className="bg-slate-800/80 p-2.5 rounded-lg border border-slate-700 space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="font-semibold text-blue-300">ហត្ថលេខាឆ្វេង ({leftSignTitle})</span>
                    <span className="text-[10px] text-gray-400 font-mono">X: {leftSignShiftX}cm | Y: {leftSignShiftY}pt</span>
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    <div className="space-y-1">
                      <label className="text-[10px] text-gray-400">ឆ្វេង-ស្តាំ (X):</label>
                      <div className="flex items-center gap-1">
                        <button
                          type="button"
                          onClick={() => updateSignaturePositions(parseFloat((leftSignShiftX - 0.5).toFixed(1)))}
                          className="px-2 py-1 bg-slate-700 hover:bg-blue-600 rounded text-white font-bold cursor-pointer"
                          title="រំកិលទៅឆ្វេង 0.5cm"
                        >
                          ◀
                        </button>
                        <span className="flex-1 text-center font-mono font-bold text-amber-300 bg-slate-900 py-0.5 rounded">
                          {leftSignShiftX}cm
                        </span>
                        <button
                          type="button"
                          onClick={() => updateSignaturePositions(parseFloat((leftSignShiftX + 0.5).toFixed(1)))}
                          className="px-2 py-1 bg-slate-700 hover:bg-blue-600 rounded text-white font-bold cursor-pointer"
                          title="រំកិលទៅស្តាំ 0.5cm"
                        >
                          ▶
                        </button>
                      </div>
                    </div>
                    <div className="space-y-1">
                      <label className="text-[10px] text-gray-400">ឡើង-ចុះ (Y):</label>
                      <div className="flex items-center gap-1">
                        <button
                          type="button"
                          onClick={() => updateSignaturePositions(undefined, leftSignShiftY - 4)}
                          className="px-2 py-1 bg-slate-700 hover:bg-blue-600 rounded text-white font-bold cursor-pointer"
                          title="រំកិលឡើងលើ 4pt"
                        >
                          ▲
                        </button>
                        <span className="flex-1 text-center font-mono font-bold text-cyan-300 bg-slate-900 py-0.5 rounded">
                          {leftSignShiftY}pt
                        </span>
                        <button
                          type="button"
                          onClick={() => updateSignaturePositions(undefined, leftSignShiftY + 4)}
                          className="px-2 py-1 bg-slate-700 hover:bg-blue-600 rounded text-white font-bold cursor-pointer"
                          title="រំកិលចុះក្រោម 4pt"
                        >
                          ▼
                        </button>
                      </div>
                    </div>
                  </div>
                </div>

                {/* Right Signature Controls */}
                <div className="bg-slate-800/80 p-2.5 rounded-lg border border-slate-700 space-y-2">
                  <div className="flex items-center justify-between">
                    <span className="font-semibold text-emerald-300">ហត្ថលេខាស្តាំ ({rightSignTitle})</span>
                    <span className="text-[10px] text-gray-400 font-mono">X: {rightSignShiftX}cm | Y: {rightSignShiftY}pt</span>
                  </div>
                  <div className="grid grid-cols-2 gap-2">
                    <div className="space-y-1">
                      <label className="text-[10px] text-gray-400">ឆ្វេង-ស្តាំ (X):</label>
                      <div className="flex items-center gap-1">
                        <button
                          type="button"
                          onClick={() => updateSignaturePositions(undefined, undefined, parseFloat((rightSignShiftX - 0.5).toFixed(1)))}
                          className="px-2 py-1 bg-slate-700 hover:bg-emerald-600 rounded text-white font-bold cursor-pointer"
                          title="រំកិលទៅឆ្វេង 0.5cm"
                        >
                          ◀
                        </button>
                        <span className="flex-1 text-center font-mono font-bold text-amber-300 bg-slate-900 py-0.5 rounded">
                          {rightSignShiftX}cm
                        </span>
                        <button
                          type="button"
                          onClick={() => updateSignaturePositions(undefined, undefined, parseFloat((rightSignShiftX + 0.5).toFixed(1)))}
                          className="px-2 py-1 bg-slate-700 hover:bg-emerald-600 rounded text-white font-bold cursor-pointer"
                          title="រំកិលទៅស្តាំ 0.5cm"
                        >
                          ▶
                        </button>
                      </div>
                    </div>
                    <div className="space-y-1">
                      <label className="text-[10px] text-gray-400">ឡើង-ចុះ (Y):</label>
                      <div className="flex items-center gap-1">
                        <button
                          type="button"
                          onClick={() => updateSignaturePositions(undefined, undefined, undefined, rightSignShiftY - 4)}
                          className="px-2 py-1 bg-slate-700 hover:bg-emerald-600 rounded text-white font-bold cursor-pointer"
                          title="រំកិលឡើងលើ 4pt"
                        >
                          ▲
                        </button>
                        <span className="flex-1 text-center font-mono font-bold text-cyan-300 bg-slate-900 py-0.5 rounded">
                          {rightSignShiftY}pt
                        </span>
                        <button
                          type="button"
                          onClick={() => updateSignaturePositions(undefined, undefined, undefined, rightSignShiftY + 4)}
                          className="px-2 py-1 bg-slate-700 hover:bg-emerald-600 rounded text-white font-bold cursor-pointer"
                          title="រំកិលចុះក្រោម 4pt"
                        >
                          ▼
                        </button>
                      </div>
                    </div>
                  </div>
                </div>

                {/* Overall Top Margin / Spacing */}
                <div className="flex items-center justify-between bg-slate-800/80 p-2.5 rounded-lg border border-slate-700">
                  <span className="text-[11px] text-gray-300 font-medium">គម្លាតពីលើ (Top Margin):</span>
                  <div className="flex items-center gap-1">
                    <button
                      type="button"
                      onClick={() => updateSignaturePositions(undefined, undefined, undefined, undefined, Math.max(0, signSectionMarginTop - 2))}
                      className="px-2 py-1 bg-slate-700 hover:bg-purple-600 rounded text-white font-bold cursor-pointer"
                      title="បន្ថយគម្លាត 2pt"
                    >
                      ▲
                    </button>
                    <span className="w-12 text-center font-mono font-bold text-purple-300 bg-slate-900 py-0.5 rounded text-[11px]">
                      {signSectionMarginTop}pt
                    </span>
                    <button
                      type="button"
                      onClick={() => updateSignaturePositions(undefined, undefined, undefined, undefined, signSectionMarginTop + 2)}
                      className="px-2 py-1 bg-slate-700 hover:bg-purple-600 rounded text-white font-bold cursor-pointer"
                      title="បង្កើនគម្លាត 2pt"
                    >
                      ▼
                    </button>
                  </div>
                </div>

                {/* Paragraph Vertical Position Shift */}
                <div className="flex items-center justify-between bg-slate-800/80 p-2.5 rounded-lg border border-amber-800/50">
                  <span className="text-[11px] text-amber-200 font-medium">ទីតាំងកថាខណ្ឌ (Paragraph Shift):</span>
                  <div className="flex items-center gap-1">
                    <button
                      type="button"
                      onClick={() => handleParagraphShiftChange(Math.max(0, paragraphShiftY - 2))}
                      className="px-2 py-1 bg-slate-700 hover:bg-amber-600 rounded text-white font-bold cursor-pointer transition"
                      title="រំកិលកថាខណ្ឌឡើងលើ 2pt"
                    >
                      ▲ ឡើង
                    </button>
                    <span className="w-14 text-center font-mono font-bold text-amber-300 bg-slate-900 py-0.5 rounded text-[11px]">
                      {paragraphShiftY}pt
                    </span>
                    <button
                      type="button"
                      onClick={() => handleParagraphShiftChange(Math.min(60, paragraphShiftY + 2))}
                      className="px-2 py-1 bg-slate-700 hover:bg-amber-600 rounded text-white font-bold cursor-pointer transition"
                      title="ទម្លាក់កថាខណ្ឌចុះក្រោម 2pt"
                    >
                      ▼ ចុះ
                    </button>
                  </div>
                </div>

                {/* Table Vertical Position Shift */}
                <div className="flex items-center justify-between bg-slate-800/80 p-2.5 rounded-lg border border-slate-700">
                  <span className="text-[11px] text-gray-300 font-medium">ទីតាំងតារាង (Table Shift):</span>
                  <div className="flex items-center gap-1">
                    <button
                      type="button"
                      onClick={() => handleTableShiftChange(tableShiftY - 2)}
                      className="px-2 py-1 bg-slate-700 hover:bg-indigo-600 rounded text-white font-bold cursor-pointer"
                      title="រំកិលតារាងឡើងលើ 2pt"
                    >
                      ▲ ឡើង
                    </button>
                    <span className="w-14 text-center font-mono font-bold text-indigo-300 bg-slate-900 py-0.5 rounded text-[11px]">
                      {tableShiftY}pt
                    </span>
                    <button
                      type="button"
                      onClick={() => handleTableShiftChange(tableShiftY + 2)}
                      className="px-2 py-1 bg-slate-700 hover:bg-indigo-600 rounded text-white font-bold cursor-pointer"
                      title="រំកិលតារាងចុះក្រោម 2pt"
                    >
                      ▼ ចុះ
                    </button>
                  </div>
                </div>

                {/* Reset all button */}
                <button
                  type="button"
                  onClick={() => {
                    updateSignaturePositions(0, 0, 0, 0, 4);
                    handleTableShiftChange(0);
                    handleParagraphShiftChange(16);
                  }}
                  className="w-full py-1.5 bg-slate-800 hover:bg-slate-700 text-gray-300 hover:text-white rounded-lg text-center text-xs font-semibold cursor-pointer border border-slate-700 transition"
                >
                  ↺ កំណត់ទីតាំងដើមឡើងវិញ (Reset to Default)
                </button>
              </div>
            )}
          </div>

          {/* Zoom Controls */}
          <div className="flex items-center gap-1 ml-auto bg-slate-900/90 px-2 py-1 rounded-lg border border-slate-700">
            <button
              onClick={() => setDocZoom((z) => Math.max(50, z - 10))}
              className="p-1 hover:text-white text-gray-300 transition cursor-pointer"
              title="បង្រួម (Zoom Out)"
            >
              <ZoomOut className="w-3.5 h-3.5" />
            </button>
            <button
              onClick={() => setDocZoom(100)}
              className="text-[11px] font-mono px-1 hover:bg-slate-800 rounded text-amber-300 cursor-pointer"
              title="កំណត់ទំហំ 100% ឡើងវិញ"
            >
              {docZoom}%
            </button>
            <button
              onClick={() => setDocZoom((z) => Math.min(160, z + 10))}
              className="p-1 hover:text-white text-gray-300 transition cursor-pointer"
              title="ពង្រីក (Zoom In)"
            >
              <ZoomIn className="w-3.5 h-3.5" />
            </button>
          </div>

          {/* Quick Hide Toolbar Button */}
          <button
            type="button"
            onClick={() => setShowRibbonToolbar(false)}
            className="p-1 px-2.5 py-1.5 rounded-lg bg-slate-900/90 hover:bg-slate-700 text-amber-300 hover:text-white transition cursor-pointer flex items-center gap-1 border border-slate-700 text-xs font-bold shrink-0 shadow-sm"
            title="លាក់របារឧបករណ៍នេះ (Hide Toolbar)"
          >
            <ChevronUp className="w-3.5 h-3.5" />
            <span>លាក់របារ</span>
          </button>
        </div>

        {/* Find & Replace Bar (collapsible) */}
        {showFindReplace && (
          <div className="bg-slate-900 p-2 rounded-lg border border-purple-500/40 flex flex-wrap items-center gap-2 text-xs">
            <div className="flex items-center gap-1 bg-slate-800 px-2 py-1 rounded border border-slate-700">
              <span className="text-gray-400">ស្វែងរក:</span>
              <input
                type="text"
                placeholder="ពាក្យស្វែងរក..."
                value={findText}
                onChange={(e) => setFindText(e.target.value)}
                onKeyDown={(e) => e.key === 'Enter' && handleFindNext()}
                className="bg-transparent text-white focus:outline-none w-32 sm:w-44 text-xs"
              />
            </div>
            <button
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
                className="bg-transparent text-white focus:outline-none w-32 sm:w-44 text-xs"
              />
            </div>
            <button
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
              onClick={() => setShowFindReplace(false)}
              className="ml-auto p-1 text-gray-400 hover:text-white rounded cursor-pointer"
            >
              <X className="w-3.5 h-3.5" />
            </button>
          </div>
        )}
      </div>
      )}

      {/* Margin Adjustment Sub-bar */}
      {showMarginControls && (
        <div className="print:hidden bg-slate-900 text-white rounded-xl p-3 border border-slate-700 space-y-3 text-xs">
          <div className="flex flex-wrap items-center justify-between gap-2 border-b border-slate-800 pb-2">
            <span className="text-gray-300 font-semibold flex items-center gap-1.5">
              <Sliders className="w-3.5 h-3.5 text-blue-400" />
              <span>កែសម្រួលគែមក្រដាស (Page Margins)</span>
            </span>
            <div className="flex flex-wrap items-center gap-1.5">
              <button
                onClick={() => setCustomMargins({ top: 0.5, bottom: 0.5, left: 2.8, right: 1.5 })}
                className={`px-2.5 py-1 rounded text-[11px] font-medium border transition cursor-pointer ${
                  customMargins.left === 2.8 && customMargins.right === 1.5 && customMargins.top === 0.5 && customMargins.bottom === 0.5
                    ? 'bg-blue-600 border-blue-500 text-white font-bold shadow'
                    : 'bg-slate-800 border-slate-700 text-gray-300 hover:text-white'
                }`}
              >
                ⭐ ស្តង់ដារលំនាំដើម (ឆ្វេង 2.8cm, ស្តាំ 1.5cm, លើ 0.5cm, ក្រោម 0.5cm)
              </button>
              <button
                onClick={() => setCustomMargins({ top: 1.2, bottom: 1.2, left: 2.5, right: 1.5 })}
                className={`px-2.5 py-1 rounded text-[11px] font-medium border transition cursor-pointer ${
                  customMargins.left === 2.5 && customMargins.right === 1.5 && customMargins.top === 1.2
                    ? 'bg-blue-600 border-blue-500 text-white'
                    : 'bg-slate-800 border-slate-700 text-gray-300 hover:text-white'
                }`}
              >
                ស្តង់ដាររដ្ឋបាល (ឆ្វេង 2.5cm, 1.2cm)
              </button>
              <button
                onClick={() => setCustomMargins({ top: 2.5, bottom: 2.5, left: 2.5, right: 2.5 })}
                className={`px-2.5 py-1 rounded text-[11px] font-medium border transition cursor-pointer ${
                  customMargins.left === 2.5 && customMargins.right === 2.5 && customMargins.top === 2.5
                    ? 'bg-blue-600 border-blue-500 text-white'
                    : 'bg-slate-800 border-slate-700 text-gray-300 hover:text-white'
                }`}
              >
                គែម 2.5cm ជុំវិញ
              </button>
            </div>
          </div>

          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
            <div>
              <label className="block text-gray-400 mb-1">គែមលើ (Top): {customMargins.top} cm</label>
              <input
                type="range"
                min="0.2"
                max="3"
                step="0.1"
                value={customMargins.top}
                onChange={(e) => setCustomMargins({ ...customMargins, top: parseFloat(e.target.value) })}
                className="w-full accent-blue-500"
              />
            </div>
            <div>
              <label className="block text-gray-400 mb-1">គែមក្រោម (Bottom): {customMargins.bottom} cm</label>
              <input
                type="range"
                min="0.2"
                max="3"
                step="0.1"
                value={customMargins.bottom}
                onChange={(e) => setCustomMargins({ ...customMargins, bottom: parseFloat(e.target.value) })}
                className="w-full accent-blue-500"
              />
            </div>
            <div>
              <label className="block text-gray-400 mb-1 font-bold text-amber-300">គែមឆ្វេង (Left): {customMargins.left} cm</label>
              <input
                type="range"
                min="0.5"
                max="4"
                step="0.1"
                value={customMargins.left}
                onChange={(e) => setCustomMargins({ ...customMargins, left: parseFloat(e.target.value) })}
                className="w-full accent-amber-500"
              />
            </div>
            <div>
              <label className="block text-gray-400 mb-1">គែមស្តាំ (Right): {customMargins.right} cm</label>
              <input
                type="range"
                min="0.5"
                max="3"
                step="0.1"
                value={customMargins.right}
                onChange={(e) => setCustomMargins({ ...customMargins, right: parseFloat(e.target.value) })}
                className="w-full accent-blue-500"
              />
            </div>
          </div>
        </div>
      )}

      {/* Official A4 Live Document Canvas with Zoom Container */}
      <div className="flex justify-center overflow-x-auto p-2 sm:p-6 bg-slate-200/60 rounded-2xl border border-slate-300/80">
        <div
          style={{
            transform: `scale(${docZoom / 100})`,
            transformOrigin: 'top center',
            transition: 'transform 0.15s ease-out',
          }}
        >
          {/* Microsoft Word Horizontal Ruler (Top of Page) */}
          {showRuler && (
            <div
              onClick={handleRulerClick}
              className="print:hidden relative mb-1.5 bg-slate-100 select-none shadow-sm rounded-t-md border border-slate-300 overflow-hidden font-sans cursor-crosshair group"
              style={{ width: '210mm', height: '28px' }}
              title="ចុចលើ Ruler ដើម្បីកំណត់ចំណុច Tab Stop (ឧ. 2cm, 2.5cm) សម្រាប់ចុច Tab ដកឃ្លា"
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

              {/* Active Tab Stop Marker (Classic Word Tab Stop L/Marker on Ruler) */}
              <div
                className="absolute top-0 z-20 -ml-1.5 flex flex-col items-center pointer-events-none"
                style={{ left: `${customMargins.left + activeTabStop}cm` }}
                title={`Tab Stop កំណត់នៅ: ${activeTabStop} cm`}
              >
                <div className="w-0 h-0 border-l-[4px] border-l-transparent border-r-[4px] border-r-transparent border-t-[5px] border-t-amber-600" />
                <div className="w-2.5 h-3 bg-amber-500 rounded-[1px] shadow-sm flex items-center justify-center text-[7px] text-black font-extrabold">
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
                <div className="w-2 h-2.5 bg-blue-600 rounded-[1px] shadow-sm flex items-center justify-center text-[6.5px] text-white font-bold">
                  L
                </div>
                <div className="w-0 h-0 border-l-[4px] border-l-transparent border-r-[4px] border-r-transparent border-b-[5px] border-b-blue-600" />
              </div>

              {/* Right Margin Indicator Marker */}
              <div
                className="absolute top-0 z-10 -mr-1.5 flex flex-col items-center cursor-pointer group pointer-events-none"
                style={{ right: `${customMargins.right}cm` }}
                title={`គែមស្តាំ (Right Margin): ${customMargins.right} cm`}
              >
                <div className="w-0 h-0 border-l-[4px] border-l-transparent border-r-[4px] border-r-transparent border-t-[5px] border-t-blue-600" />
                <div className="w-2 h-2.5 bg-blue-600 rounded-[1px] shadow-sm flex items-center justify-center text-[6.5px] text-white font-bold">
                  R
                </div>
                <div className="w-0 h-0 border-l-[4px] border-l-transparent border-r-[4px] border-r-transparent border-b-[5px] border-b-blue-600" />
              </div>
            </div>
          )}

            <div
              id="robok-total-stock-work-pdf"
              key={`${keyResetCounter}_fresh`}
              ref={printAreaRef}
              data-adjust-page={adjustPagePercent || 97}
              contentEditable={isWordEditMode}
              suppressContentEditableWarning={true}
              onKeyDown={handleKeyDownOnPaper}
              className={`bg-white shadow-2xl text-black font-siemreap text-[10pt] leading-relaxed transition-all ${
                isWordEditMode
                  ? 'outline-none ring-2 ring-blue-400/40 hover:ring-blue-500 cursor-text select-text'
                  : 'outline-none cursor-default'
              }`}
              style={{
                width: '210mm',
                minHeight: '297mm',
                paddingTop: `${customMargins.top}cm`,
                paddingBottom: `${customMargins.bottom}cm`,
                paddingLeft: `${customMargins.left}cm`,
                paddingRight: `${customMargins.right}cm`,
                boxSizing: 'border-box',
                backgroundColor: '#ffffff',
                ['--robok-print-scale' as any]: `${(adjustPagePercent || 97) / 100}`,
              }}
            >
              {viewMode === 'team' ? (
                /* ========================================================
                   ទម្រង់របាយការណ៍សរុបការងារស្តុកតាមក្រុម (PDF ផ្លូវការ ១០០%)
                   ======================================================== */
                <>
                  {/* Header section (Left: Ministry & Team, Right: Kingdom) */}
                  <div className="flex justify-between items-start mb-2.5">
                    {/* Left Ministry Info */}
                    <div className="text-center inline-block pt-1 text-black">
                      {/* Spacer matching height of ព្រះរាជាណាចក្រកម្ពុជា so ក្រសួងមហាផ្ទៃ aligns exactly with ជាតិ សាសនា ព្រះមហាក្សត្រ */}
                      <p className="font-moul text-[12pt] leading-[1.35] whitespace-nowrap invisible select-none pointer-events-none" aria-hidden="true">
                        ព្រះរាជាណាចក្រកម្ពុជា
                      </p>
                      <div className="mt-1 space-y-0.5">
                        <p className="font-moul text-[12pt] whitespace-nowrap leading-[1.35]">ក្រសួងមហាផ្ទៃ</p>
                        <p className="font-moul text-[12pt] whitespace-nowrap leading-[1.35]">អគ្គនាយកដ្ឋានអន្តោប្រវេសន៍</p>
                        <p className="font-moul text-[12pt] whitespace-nowrap leading-[1.35]">នាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត</p>
                        <p className="font-moul text-[12pt] whitespace-nowrap leading-[1.35]">ការិយាល័យទិដ្ឋាការចូល</p>
                        <p className="font-siemreap font-bold text-[10pt] whitespace-nowrap leading-[1.35]">{teamDisplayName}</p>
                        <div className="mt-1 mb-0 flex justify-center text-black">
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

                    {/* Right Kingdom Info */}
                    <div className="text-center inline-block text-black pt-1 space-y-1">
                      <p className="font-moul text-[12pt] leading-[1.35] whitespace-nowrap">ព្រះរាជាណាចក្រកម្ពុជា</p>
                      <p className="font-moul text-[12pt] leading-[1.35] whitespace-nowrap">ជាតិ សាសនា ព្រះមហាក្សត្រ</p>
                      <div className="mt-1 flex justify-center text-black">
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

                  {/* Letter To */}
                  <div className="text-center space-y-1 leading-snug mt-3.5 mb-3">
                    <p className="font-moul text-[14pt] leading-snug m-0 p-0" style={{ fontSize: '14pt' }}>សូមគោរពជូន</p>
                    <p className="font-moul text-[12pt] leading-snug m-0 p-0">លោកវរសេនីយ៍ទោ នាយការិយាល័យទិដ្ឋាការចូល</p>
                  </div>

                  {/* Details Block (កម្មវត្ថុ & Intro Paragraph) */}
                  <div className="text-[12pt] font-siemreap mb-2">
                    <div className="flex items-baseline leading-[1.7] mb-0.5">
                      <div className="w-[88px] shrink-0 flex items-baseline justify-between pr-2">
                        <span className="font-moul text-[12pt]" style={{ fontFamily: "'Khmer Mool1', 'Khmer Mool 1', 'Khmer Mool', 'Khmer OS Mool1', 'Khmer OS Muol Light', 'Moul', serif" }}>កម្មវត្ថុ</span>
                        <span className="font-bold text-[12pt]">៖</span>
                      </div>
                      <div className="text-justify font-siemreap text-[12pt] flex-1 leading-[1.7]">
                        របាយការណ៍ ស្តីពីការប្រើប្រាស់សន្លឹកទិដ្ឋាការស្អិត {teamPeriodKhmerTitle} ។
                      </div>
                    </div>

                    {/* Introductory Body Paragraph dropped down properly */}
                    <div
                      className="text-justify font-siemreap text-[12pt] leading-[1.75]"
                      style={{
                        textIndent: `${activeTabStop ? activeTabStop + 'cm' : '2.25cm'}`,
                        marginTop: `${paragraphShiftY}pt`,
                        marginBottom: '8pt'
                      }}
                    >
                      តបតាមកម្មវត្ថុខាងលើ {teamDisplayName} សូមជម្រាបជូន <span className="font-moul text-[12pt]">លោកវរសេនីយ៍ទោ នាយការិយាល័យទិដ្ឋាការចូល</span> មេត្តាជ្រាបថា៖ ការប្រើប្រាស់សន្លឹកទិដ្ឋាការស្អិត របស់{teamDisplayName} {teamPeriodShortKhmerTitle} ដោយគិតចាប់ពីថ្ងៃទី{startSolar.khmerDay} ខែ{startSolar.khmerMonth} ឆ្នាំ{startSolar.khmerYear} រហូតដល់ថ្ងៃទី{endSolar.khmerDay} ខែ{endSolar.khmerMonth} ឆ្នាំ{endSolar.khmerYear} មានចំនួនដូចខាងក្រោម ៖
                    </div>
                  </div>

                  {/* 5-Row Stock Summary Table Matching PDF 100% */}
                  <div className="mb-1">
                    <table className="w-full table-fixed text-center border-collapse border border-black text-[8.5pt] leading-tight bg-white">
                      <colgroup>
                        <col style={{ width: '14%' }} />
                        {VISA_TYPES.map((vt) => (
                          <col
                            key={vt}
                            style={{
                              width: vt === 'T' || vt === 'E' ? '8%' : vt === 'D' || vt === 'K' ? '7%' : '5.333%',
                            }}
                          />
                        ))}
                        <col style={{ width: '8%' }} />
                      </colgroup>
                      <thead>
                        <tr className="bg-white font-normal text-[8.5pt] h-[28px]">
                          <th className="border border-black p-0 w-[14%] h-[28px] relative bg-white overflow-hidden">
                            <svg className="absolute inset-0 w-full h-full" preserveAspectRatio="none">
                              <line x1="0" y1="0" x2="100%" y2="100%" stroke="#000000" strokeWidth="1.2" />
                            </svg>
                          </th>
                          {VISA_TYPES.map((vt) => {
                            const colWidth = vt === 'T' || vt === 'E' ? '8%' : vt === 'D' || vt === 'K' ? '7%' : '5.333%';
                            return (
                              <th
                                key={vt}
                                className="border border-black px-0.5 py-1.5 align-middle font-times font-bold bg-white text-[8.5pt] whitespace-nowrap overflow-hidden"
                                style={{
                                  width: colWidth,
                                  fontFamily: "'Times New Roman', Times, serif",
                                  fontSize: '8.5pt',
                                }}
                              >
                                {vt}
                              </th>
                            );
                          })}
                          <th
                            className="border border-black px-0.5 py-1.5 align-middle font-siemreap font-bold bg-white text-[8.5pt] whitespace-nowrap overflow-hidden"
                            style={{ width: '8%' }}
                          >
                            សរុប
                          </th>
                        </tr>
                      </thead>
                      <tbody>
                        {/* Row 1: ចំនួនសល់ខែចាស់ */}
                        <tr className="hover:bg-blue-50/20 h-[26px]">
                          <td className="border border-black px-1.5 py-1.5 text-left font-siemreap text-[8pt] whitespace-nowrap overflow-hidden" style={{ width: '14%' }}>
                            ចំនួនសល់ខែចាស់
                          </td>
                          {VISA_TYPES.map((vt) => {
                            const colWidth = vt === 'T' || vt === 'E' ? '8%' : vt === 'D' || vt === 'K' ? '7%' : '5.333%';
                            return (
                              <td
                                key={vt}
                                className="border border-black px-0.5 py-1.5 font-times text-[8.5pt] whitespace-nowrap overflow-hidden text-center"
                                style={{
                                  width: colWidth,
                                  fontFamily: "'Times New Roman', Times, serif",
                                  fontSize: '8.5pt',
                                }}
                              >
                                {renderNum(teamStickerData.opening.values[vt])}
                              </td>
                            );
                          })}
                          <td
                            className="border border-black px-0.5 py-1.5 font-times font-bold text-[8.5pt] bg-gray-50/50 whitespace-nowrap overflow-hidden text-center"
                            style={{ width: '8%', fontFamily: "'Times New Roman', Times, serif", fontSize: '8.5pt' }}
                          >
                            {renderNum(teamStickerData.opening.total)}
                          </td>
                        </tr>

                        {/* Row 2: ចំនួនបើកថ្មី */}
                        <tr className="hover:bg-blue-50/20 h-[26px]">
                          <td className="border border-black px-1.5 py-1.5 text-left font-siemreap text-[8pt] whitespace-nowrap overflow-hidden" style={{ width: '14%' }}>
                            ចំនួនបើកថ្មី
                          </td>
                          {VISA_TYPES.map((vt) => {
                            const colWidth = vt === 'T' || vt === 'E' ? '8%' : vt === 'D' || vt === 'K' ? '7%' : '5.333%';
                            return (
                              <td
                                key={vt}
                                className="border border-black px-0.5 py-1.5 font-times text-[8.5pt] whitespace-nowrap overflow-hidden text-center"
                                style={{
                                  width: colWidth,
                                  fontFamily: "'Times New Roman', Times, serif",
                                  fontSize: '8.5pt',
                                }}
                              >
                                {renderNum(teamStickerData.newIssued.values[vt])}
                              </td>
                            );
                          })}
                          <td
                            className="border border-black px-0.5 py-1.5 font-times font-bold text-[8.5pt] bg-gray-50/50 whitespace-nowrap overflow-hidden text-center"
                            style={{ width: '8%', fontFamily: "'Times New Roman', Times, serif", fontSize: '8.5pt' }}
                          >
                            {renderNum(teamStickerData.newIssued.total)}
                          </td>
                        </tr>

                        {/* Row 3: ផ្ទេរ/បង្វិល/ខូច/ខ្វះ */}
                        <tr className="hover:bg-blue-50/20 h-[26px]">
                          <td className="border border-black px-1.5 py-1.5 text-left font-siemreap text-[8pt] whitespace-nowrap overflow-hidden" style={{ width: '14%' }}>
                            ផ្ទេរ/បង្វិល/ខូច/ខ្វះ
                          </td>
                          {VISA_TYPES.map((vt) => {
                            const colWidth = vt === 'T' || vt === 'E' ? '8%' : vt === 'D' || vt === 'K' ? '7%' : '5.333%';
                            return (
                              <td
                                key={vt}
                                className="border border-black px-0.5 py-1.5 font-times text-[8.5pt] whitespace-nowrap overflow-hidden text-center"
                                style={{
                                  width: colWidth,
                                  fontFamily: "'Times New Roman', Times, serif",
                                  fontSize: '8.5pt',
                                }}
                              >
                                {renderNum(teamStickerData.returnedDamagedMissing.values[vt])}
                              </td>
                            );
                          })}
                          <td
                            className="border border-black px-0.5 py-1.5 font-times font-bold text-[8.5pt] bg-gray-50/50 whitespace-nowrap overflow-hidden text-center"
                            style={{ width: '8%', fontFamily: "'Times New Roman', Times, serif", fontSize: '8.5pt' }}
                          >
                            {renderNum(teamStickerData.returnedDamagedMissing.total)}
                          </td>
                        </tr>

                        {/* Row 4: ចំនួនប្រើ */}
                        <tr className="hover:bg-blue-50/20 h-[26px]">
                          <td className="border border-black px-1.5 py-1.5 text-left font-siemreap text-[8pt] whitespace-nowrap overflow-hidden" style={{ width: '14%' }}>
                            ចំនួនប្រើ
                          </td>
                          {VISA_TYPES.map((vt) => {
                            const colWidth = vt === 'T' || vt === 'E' ? '8%' : vt === 'D' || vt === 'K' ? '7%' : '5.333%';
                            return (
                              <td
                                key={vt}
                                className="border border-black px-0.5 py-1.5 font-times text-[8.5pt] whitespace-nowrap overflow-hidden text-center"
                                style={{
                                  width: colWidth,
                                  fontFamily: "'Times New Roman', Times, serif",
                                  fontSize: '8.5pt',
                                }}
                              >
                                {renderNum(teamStickerData.used.values[vt])}
                              </td>
                            );
                          })}
                          <td
                            className="border border-black px-0.5 py-1.5 font-times font-bold text-[8.5pt] bg-gray-50/50 whitespace-nowrap overflow-hidden text-center"
                            style={{ width: '8%', fontFamily: "'Times New Roman', Times, serif", fontSize: '8.5pt' }}
                          >
                            {renderNum(teamStickerData.used.total)}
                          </td>
                        </tr>

                        {/* Row 5: សល់ខែបន្ទាប់ */}
                        <tr className="hover:bg-blue-50/20 font-semibold h-[26px]">
                          <td className="border border-black px-1.5 py-1.5 text-left font-siemreap font-bold text-[8pt] whitespace-nowrap overflow-hidden" style={{ width: '14%' }}>
                            សល់ខែបន្ទាប់
                          </td>
                          {VISA_TYPES.map((vt) => {
                            const colWidth = vt === 'T' || vt === 'E' ? '8%' : vt === 'D' || vt === 'K' ? '7%' : '5.333%';
                            return (
                              <td
                                key={vt}
                                className="border border-black px-0.5 py-1.5 font-times font-semibold text-[8.5pt] whitespace-nowrap overflow-hidden text-center"
                                style={{
                                  width: colWidth,
                                  fontFamily: "'Times New Roman', Times, serif",
                                  fontSize: '8.5pt',
                                }}
                              >
                                {renderNum(teamStickerData.ending.values[vt])}
                              </td>
                            );
                          })}
                          <td
                            className="border border-black px-0.5 py-1.5 font-times font-bold text-[8.5pt] bg-gray-50/50 whitespace-nowrap overflow-hidden text-center"
                            style={{ width: '8%', fontFamily: "'Times New Roman', Times, serif", fontSize: '8.5pt' }}
                          >
                            {renderNum(teamStickerData.ending.total)}
                          </td>
                        </tr>
                      </tbody>
                    </table>
                  </div>

                  {/* Concluding text */}
                  <div className="text-[12pt] font-siemreap space-y-1.5 mt-3.5 mb-2.5 leading-[1.75]">
                    <p style={{ textIndent: `${activeTabStop ? activeTabStop + 'cm' : '2.25cm'}` }} className="text-justify m-0 p-0 leading-[1.75]">
                      អាស្រ័យដូចបានជម្រាបជូនខាងលើ សូម <span className="font-moul text-[12pt]">លោកវរសេនីយ៍ទោ នាយការិយាល័យទិដ្ឋាការចូល</span> មេត្តា ពិនិត្យ និងជ្រាប ជារាយការណ៍ដោយក្តីអនុគ្រោះ ។
                    </p>
                    <p style={{ textIndent: `${activeTabStop ? activeTabStop + 'cm' : '2.25cm'}` }} className="text-justify m-0 p-0 leading-[1.75]">
                      សូម<span className="font-moul text-[12pt]">លោកវរសេនីយ៍ទោ</span> មេត្តាទទួលនូវការគោរព អំពីខ្ញុំ ។
                    </p>
                  </div>

                  {/* Signature Section */}
                  <div
                    data-signature-section="true"
                    className="flex justify-end text-center text-[12pt] font-siemreap items-start"
                    style={{ marginTop: `${signSectionMarginTop}pt` }}
                  >
                    <div
                      data-signature-block="right"
                      className="space-y-0.5 relative group/rightsign min-w-[280px]"
                      style={{
                        position: 'relative',
                        left: `${rightSignShiftX}cm`,
                        top: `${rightSignShiftY}pt`,
                      }}
                    >
                      {/* On-screen quick shift adjuster */}
                      <div className="print:hidden opacity-0 group-hover/rightsign:opacity-100 transition-opacity absolute -top-8 right-0 bg-slate-900/90 backdrop-blur-xs text-white text-[10px] px-2 py-1 rounded shadow-lg flex items-center gap-1 z-30 whitespace-nowrap border border-slate-700">
                        <span className="text-gray-300 mr-0.5">ទីតាំង:</span>
                        <button
                          type="button"
                          onMouseDown={(e) => { e.preventDefault(); e.stopPropagation(); }}
                          onClick={(e) => {
                            e.preventDefault();
                            e.stopPropagation();
                            updateSignaturePositions(undefined, undefined, parseFloat((rightSignShiftX - 0.2).toFixed(1)));
                          }}
                          className="px-1.5 py-0.5 bg-slate-800 hover:bg-emerald-600 rounded text-white font-bold cursor-pointer transition active:scale-95"
                          title="រំកិលទៅឆ្វេង 0.2cm"
                        >
                          ◀
                        </button>
                        <span className="font-mono text-emerald-300 font-bold px-0.5">{rightSignShiftX}cm</span>
                        <button
                          type="button"
                          onMouseDown={(e) => { e.preventDefault(); e.stopPropagation(); }}
                          onClick={(e) => {
                            e.preventDefault();
                            e.stopPropagation();
                            updateSignaturePositions(undefined, undefined, parseFloat((rightSignShiftX + 0.2).toFixed(1)));
                          }}
                          className="px-1.5 py-0.5 bg-slate-800 hover:bg-emerald-600 rounded text-white font-bold cursor-pointer transition active:scale-95"
                          title="រំកិលទៅស្តាំ 0.2cm"
                        >
                          ▶
                        </button>
                        <span className="text-slate-500 mx-0.5">|</span>
                        <button
                          type="button"
                          onMouseDown={(e) => { e.preventDefault(); e.stopPropagation(); }}
                          onClick={(e) => {
                            e.preventDefault();
                            e.stopPropagation();
                            updateSignaturePositions(undefined, undefined, undefined, rightSignShiftY - 2);
                          }}
                          className="px-1.5 py-0.5 bg-slate-800 hover:bg-emerald-600 rounded text-white font-bold cursor-pointer transition active:scale-95"
                          title="រំកិលឡើងលើ 2pt"
                        >
                          ▲
                        </button>
                        <span className="font-mono text-cyan-300 font-bold px-0.5">{rightSignShiftY}pt</span>
                        <button
                          type="button"
                          onMouseDown={(e) => { e.preventDefault(); e.stopPropagation(); }}
                          onClick={(e) => {
                            e.preventDefault();
                            e.stopPropagation();
                            updateSignaturePositions(undefined, undefined, undefined, rightSignShiftY + 2);
                          }}
                          className="px-1.5 py-0.5 bg-slate-800 hover:bg-emerald-600 rounded text-white font-bold cursor-pointer transition active:scale-95"
                          title="រំកិលចុះក្រោម 2pt"
                        >
                          ▼
                        </button>
                        <span className="text-slate-500 mx-0.5">|</span>
                        <button
                          type="button"
                          onMouseDown={(e) => { e.preventDefault(); e.stopPropagation(); }}
                          onClick={(e) => {
                            e.preventDefault();
                            e.stopPropagation();
                            updateSignaturePositions(undefined, undefined, 0, 0);
                          }}
                          className="px-1 py-0.5 bg-slate-800 hover:bg-red-600 rounded text-amber-300 hover:text-white font-bold cursor-pointer transition active:scale-95 text-[9px]"
                          title="កំណត់ទីតាំងដើមឡើងវិញ"
                        >
                          Reset
                        </button>
                      </div>

                      <div
                        onClick={() => setShowDocDateModal(true)}
                        className="cursor-pointer hover:bg-amber-50/80 rounded px-1 py-0.5 transition-colors border border-transparent hover:border-amber-300 relative group/datebtn text-center"
                        title="ចុចត្រង់នេះដើម្បីផ្លាស់ប្តូរ ថ្ងៃខែឆ្នាំតាក់តែងលិខិត"
                      >
                        <p className="text-[12pt]">
                          {customSignLunar || signLunar || 'ថ្ងៃ.....................ខែ.............ឆ្នាំ..............សំរឹទ្ធិស័ក ព.ស. ២៥៦....'}
                        </p>
                        <p className="text-[12pt]">
                          {customSignSolar || `${selectedTeam ? (selectedTeam.includes('កំពង់ផែ') || selectedTeam.includes('ព្រំដែន') || selectedTeam.includes('អាកាស') ? selectedTeam : 'ច្រកទ្វារ' + selectedTeam) : ''}, ថ្ងៃទី${signSolar.khmerDay} ខែ${signSolar.khmerMonth} ឆ្នាំ${signSolar.khmerYear}`}
                        </p>
                        <span className="print:hidden opacity-0 group-hover/datebtn:opacity-100 transition-opacity absolute -top-5 left-1/2 -translate-x-1/2 bg-blue-900 text-white text-[9px] px-1.5 py-0.5 rounded shadow whitespace-nowrap z-20 pointer-events-none">
                          📅 ចុចកែប្រែកាលបរិច្ឆេទ
                        </span>
                      </div>
                      <p className="font-moul text-[12pt] mt-0.5">{teamLeaderTitle}</p>
                      <div className="h-16 flex items-end justify-center">
                        {rightSignOfficer && <p className="font-bold text-[12pt]">{rightSignOfficer}</p>}
                      </div>
                    </div>
                  </div>
                </>
              ) : (
                /* ========================================================
                   ទម្រង់សរុបការងារស្តុកការិយាល័យ (Office Overview Report)
                   ======================================================== */
                <>
                  {/* Header section (Left: Ministry, Right: Kingdom) */}
                  <div className="flex justify-between items-start mb-0.5">
                    {/* Left Ministry Info (Centered under itself, top margin aligned so line 1 starts level with 'ជាតិ សាសនា ព្រះមហាក្សត្រ') */}
                    <div className="text-center inline-block pt-1 text-black">
                      {/* Spacer matching height of ព្រះរាជាណាចក្រកម្ពុជា so ក្រសួងមហាផ្ទៃ aligns exactly with ជាតិ សាសនា ព្រះមហាក្សត្រ */}
                      <p className="font-moul text-[12pt] leading-tight whitespace-nowrap invisible select-none pointer-events-none" aria-hidden="true" style={{ fontSize: '12pt' }}>
                        ព្រះរាជាណាចក្រកម្ពុជា
                      </p>
                      <div className="mt-1 space-y-0.5">
                        {ministryHierarchy.map((line, idx) => (
                          <p key={idx} className="font-moul text-[12pt] whitespace-nowrap leading-tight" style={{ fontSize: '12pt' }}>{line}</p>
                        ))}
                        <div className="mt-0.5 mb-0 flex justify-center">
                          <TacteingLine
                            type={tacteingSettings.type}
                            customImage={tacteingSettings.customImage}
                            width={110}
                            height={12}
                          />
                        </div>
                        {proposalType === 'office' && (
                          <div
                            className="text-center font-siemreap text-[12pt] pt-0.5 whitespace-nowrap outline-none"
                            style={{ fontSize: '12pt' }}
                            contentEditable={isWordEditMode}
                            suppressContentEditableWarning={true}
                            onBlur={(e) => {
                              const txt = e.currentTarget.textContent?.trim();
                              if (txt) {
                                const clean = txt.replace(/^លេខ\s*[:៖]\s*/, '');
                                setDocNumber(clean);
                              }
                            }}
                          >
                            លេខ ៖ {(!docNumber || docNumber === '.......................' || (docNumber.includes('.....') && !docNumber.includes('សណ/២៦')) || docNumber === '...................សណ/២៦')
                              ? '.......................សណ/២៦'
                              : docNumber.replace(/^លេខ\s*[:៖]\s*/, '')}
                          </div>
                        )}
                      </div>
                    </div>

                    {/* Right Kingdom Info */}
                    <div className="text-center inline-block text-black pt-1 space-y-1">
                      <p className="font-moul text-[12pt] leading-tight whitespace-nowrap" style={{ fontSize: '12pt' }}>ព្រះរាជាណាចក្រកម្ពុជា</p>
                      <p className="font-moul text-[12pt] leading-tight whitespace-nowrap" style={{ fontSize: '12pt' }}>ជាតិ សាសនា ព្រះមហាក្សត្រ</p>
                      <div className="mt-1 flex justify-center text-black">
                        <TacteingLine
                          type={tacteingSettings.type}
                          customImage={tacteingSettings.customImage}
                          width={130}
                          height={14}
                        />
                      </div>
                    </div>
                  </div>

            {/* Letter To */}
            <div className="text-center space-y-0 leading-snug mt-1.5 mb-2">
              <p className="font-moul text-[14pt] leading-snug" style={{ fontSize: '14pt' }}>សូមគោរពជូន</p>
              {proposalType === 'office' ? (
                <>
                  <p
                    className="font-moul text-[12pt] leading-snug outline-none"
                    style={{ fontSize: '12pt' }}
                    contentEditable={isWordEditMode}
                    suppressContentEditableWarning={true}
                    onBlur={(e) => {
                      const txt = e.currentTarget.textContent?.trim();
                      if (txt) setDocRecipientRank(txt);
                    }}
                  >
                    {docRecipientRank || 'លោកឧត្តមសេនីយ៍ទោ'}
                  </p>
                  <p
                    className="font-moul text-[12pt] leading-snug outline-none"
                    style={{ fontSize: '12pt' }}
                    contentEditable={isWordEditMode}
                    suppressContentEditableWarning={true}
                    onBlur={(e) => {
                      const txt = e.currentTarget.textContent?.trim();
                      if (txt) setDocRecipientTitle(txt);
                    }}
                  >
                    {docRecipientTitle || 'ប្រធាននាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត'}
                  </p>
                </>
              ) : (
                <p
                  className="font-moul text-[12pt] leading-snug outline-none"
                  style={{ fontSize: '12pt' }}
                  contentEditable={isWordEditMode}
                  suppressContentEditableWarning={true}
                  onBlur={(e) => {
                    const txt = e.currentTarget.textContent?.trim();
                    if (txt) setDocRecipientTitle(txt);
                  }}
                >
                  {docRecipientTitle}
                </p>
              )}
            </div>

            {/* Details Block (តាមរយៈ and កម្មវត្ថុ) */}
            <div className="text-[12pt] space-y-1 leading-normal mb-1" style={{ fontSize: '12pt' }}>
              <div className="flex items-baseline">
                <div className="w-[88px] shrink-0 flex items-baseline justify-between pr-2">
                  <span className="font-moul text-[12pt]" style={{ fontFamily: "'Khmer OS Mool1', 'Khmer OS Muol Light', 'Khmer OS Muol', 'Moul', serif", fontSize: '12pt' }}>តាមរយៈ</span>
                  <span className="font-siemreap font-bold text-[12pt]" style={{ fontSize: '12pt' }}>៖</span>
                </div>
                <span
                  className="font-siemreap text-[12pt] flex-1 outline-none"
                  style={{ fontSize: '12pt' }}
                  contentEditable={isWordEditMode}
                  suppressContentEditableWarning={true}
                  onBlur={(e) => {
                    const txt = e.currentTarget.textContent?.trim();
                    if (txt) setDocThroughTitle(cleanDocThroughTitle(txt));
                  }}
                >
                  {cleanDocThroughTitle(docThroughTitle)}
                </span>
              </div>

              <div className="flex items-baseline">
                <div className="w-[88px] shrink-0 flex items-baseline justify-between pr-2">
                  <span className="font-moul text-[12pt]" style={{ fontFamily: "'Khmer OS Mool1', 'Khmer OS Muol Light', 'Khmer OS Muol', 'Moul', serif", fontSize: '12pt' }}>កម្មវត្ថុ</span>
                  <span className="font-siemreap font-bold text-[12pt]" style={{ fontSize: '12pt' }}>៖</span>
                </div>
                <div className="text-justify leading-normal font-siemreap text-[12pt] flex-1" style={{ fontSize: '12pt' }}>
                  {isStubProposal
                    ? 'សំណើសុំការអនុញ្ញាតប្រគល់ គល់ទិដ្ឋាការស្អិតដែលបានប្រើប្រាស់រួច ពីបណ្តាក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអន្តរជាតិនានា ជូនទៅការិយាល័យរដ្ឋបាលសរុប។'
                    : `របាយការណ៍ស្តីពីការបើក ការផ្តល់សន្លឹកទិដ្ឋាការស្អិត និងក្រដាសអនុម័តផ្តល់ទិដ្ឋាការអេឡិចត្រូនិកនៅពេលមកដល់ នៅតាមបណ្តាក្រុមផ្តល់ទិដ្ឋាការ ${reportPeriodLabel}។`}
                </div>
              </div>

              {isStubProposal && (
                <div className="flex items-baseline">
                  <div className="w-[88px] shrink-0 flex items-baseline justify-between pr-2">
                    <span className="font-moul text-[12pt]" style={{ fontFamily: "'Khmer OS Mool1', 'Khmer OS Muol Light', 'Khmer OS Muol', 'Moul', serif", fontSize: '12pt' }}>យោង</span>
                    <span className="font-siemreap font-bold text-[12pt]" style={{ fontSize: '12pt' }}>៖</span>
                  </div>
                  <div
                    className="text-justify leading-normal font-siemreap text-[12pt] flex-1 outline-none"
                    style={{ fontSize: '12pt' }}
                    contentEditable={isWordEditMode}
                    suppressContentEditableWarning={true}
                    onBlur={(e) => {
                      const txt = e.currentTarget.textContent?.trim();
                      if (txt) setDocReferenceText(txt);
                    }}
                  >
                    {docReferenceText}
                  </div>
                </div>
              )}

              {/* Introductory Body Paragraph dropped down properly */}
              <div
                className="text-justify leading-normal font-siemreap text-[12pt]"
                style={{
                  textIndent: `${activeTabStop ? activeTabStop + 'cm' : '2.25cm'}`,
                  marginTop: `${paragraphShiftY}pt`,
                  fontSize: '12pt'
                }}
              >
                {proposalType === 'office' ? (
                  <>
                    សេចក្តីដូចមានចែងក្នុងកម្មវត្ថុ និងយោងខាងលើ ការិយាល័យទិដ្ឋាការចូល មានកិត្តិយសសូមជម្រាបជូន<span className="font-moul text-[12pt]" style={{ fontFamily: "'Khmer OS Mool1', 'Khmer OS Muol Light', 'Khmer OS Muol', 'Moul', serif", fontSize: '12pt' }}>{docRecipientRank || 'លោកឧត្តមសេនីយ៍ទោ'} {docRecipientTitle || 'ប្រធាននាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត'}</span> មេត្តាជ្រាបថា ៖ ការិយាល័យស្នើសុំការអនុញ្ញាតប្រគល់ គល់ទិដ្ឋាការដែលប្រើប្រាស់រួច ពីបណ្តាក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអន្តរជាតិនានា ទៅការិយាល័យរដ្ឋបាលសរុប មានចំនួនដូចខាងក្រោម ៖
                  </>
                ) : proposalType === 'section' ? (
                  <>
                    សេចក្តីដូចមានចែងក្នុងកម្មវត្ថុ និងយោងខាងលើ ផ្នែករដ្ឋបាល មានកិត្តិយសសូមជម្រាប ជូន<span className="font-moul text-[12pt]" style={{ fontFamily: "'Khmer OS Mool1', 'Khmer OS Muol Light', 'Khmer OS Muol', 'Moul', serif", fontSize: '12pt' }}>{docRecipientTitle}</span> មេត្តាជ្រាបថា ៖ ផ្នែកស្នើសុំការអនុញ្ញាតប្រគល់ គល់ទិដ្ឋាការដែលប្រើប្រាស់រួច ពីបណ្តាក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអន្តរជាតិនានា ទៅការិយាល័យរដ្ឋបាលសរុប មានចំនួនដូចខាងក្រោម ៖
                  </>
                ) : (
                  <>
                    តបតាមកម្មវត្ថុខាងលើ ផ្នែករដ្ឋបាល មានកិត្តិយសសូមជម្រាបជូន <span className="font-moul text-[12pt]" style={{ fontFamily: "'Khmer OS Mool1', 'Khmer OS Muol Light', 'Khmer OS Muol', 'Moul', serif", fontSize: '12pt' }}>{docRecipientTitle.includes('ឯកឧត្តម') ? 'ឯកឧត្តម នាយក' : docRecipientTitle}</span> មេត្តាជ្រាបថា ៖ នៅក្នុងកិច្ចប្រតិបត្តិការ {reportPeriodLabel} ដោយគិតចាប់ពីថ្ងៃទី{startSolar.khmerDay} ខែ{startSolar.khmerMonth} ឆ្នាំ{startSolar.khmerYear} រហូតដល់ថ្ងៃទី{endSolar.khmerDay} ខែ{endSolar.khmerMonth} ឆ្នាំ{endSolar.khmerYear} ផ្នែកសម្រេចលទ្ធផលការបើក ការផ្តល់សន្លឹកទិដ្ឋាការ និងផ្តល់ក្រដាសអនុម័តផ្តល់ទិដ្ឋាការអេឡិចត្រូនិកពេលមកដល់ នៅតាមបណ្តាក្រុមផ្តល់ទិដ្ឋាការ មានដូចខាងក្រោម ៖
                  </>
                )}
              </div>
            </div>

            {isStubProposal ? (
              /* Picture 2 Table: Stub proposal table */
              <div className="my-1">
                <table className="w-full text-center border-separate border-spacing-0 border-t border-l border-black text-[12pt] leading-snug bg-white" style={{ fontSize: '12pt' }}>
                  <thead>
                    <tr className="bg-white font-bold text-[12pt]" style={{ fontSize: '12pt' }}>
                      <th rowSpan={2} className="border-r border-b border-black px-1.5 py-1 align-middle w-[10%] font-siemreap font-bold bg-white text-[12pt]" style={{ fontSize: '12pt' }}>
                        ល.រ
                      </th>
                      <th rowSpan={2} className="border-r border-b border-black px-1.5 py-1 align-middle w-[18%] font-siemreap font-bold bg-white text-[12pt]" style={{ fontSize: '12pt' }}>
                        ប្រភេទ
                      </th>
                      <th colSpan={2} className="border-r border-b border-black px-1.5 py-1 align-middle w-[48%] font-siemreap font-bold bg-white text-[12pt]" style={{ fontSize: '12pt' }}>
                        ឯកតា
                      </th>
                      <th rowSpan={2} className="border-r border-b border-black px-1.5 py-1 align-middle w-[24%] font-siemreap font-bold bg-white text-[12pt]" style={{ fontSize: '12pt' }}>
                        ផ្សេង ៗ
                      </th>
                    </tr>
                    <tr className="bg-white font-bold text-[12pt]" style={{ fontSize: '12pt' }}>
                      <th className="border-r border-b border-black px-1.5 py-1 align-middle w-[24%] font-siemreap font-bold bg-white text-[12pt]" style={{ fontSize: '12pt' }}>
                        ក្បាល
                      </th>
                      <th className="border-r border-b border-black px-1.5 py-1 align-middle w-[24%] font-siemreap font-bold bg-white text-[12pt]" style={{ fontSize: '12pt' }}>
                        សន្លឹក
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    {displayStubRows.map((row, idx) => (
                      <tr key={row.visaType}>
                        <td className="border-r border-b border-black px-1.5 py-0.5 align-middle text-center font-times text-[12pt] bg-white" style={{ fontSize: '12pt' }}>
                          {idx + 1}
                        </td>
                        <td className="border-r border-b border-black px-1.5 py-0.5 align-middle font-bold font-times text-center text-[12pt] bg-white" style={{ fontSize: '12pt' }}>
                          {row.visaType}
                        </td>
                        <td
                          className="border-r border-b border-black px-1.5 py-0.5 align-middle font-times text-center text-[12pt] bg-white outline-none"
                          style={{ fontSize: '12pt' }}
                          contentEditable={isWordEditMode && isCalculated}
                          suppressContentEditableWarning={true}
                          onBlur={(e) => {
                            const val = parseInt(e.currentTarget.textContent?.replace(/,/g, '') || '0', 10) || 0;
                            handleUpdateStubRow(idx, 'booklets', val);
                          }}
                        >
                          {isCalculated ? Number(row.booklets).toLocaleString() : '-'}
                        </td>
                        <td
                          className="border-r border-b border-black px-1.5 py-0.5 align-middle font-times text-center text-[12pt] bg-white outline-none"
                          style={{ fontSize: '12pt' }}
                          contentEditable={isWordEditMode && isCalculated}
                          suppressContentEditableWarning={true}
                          onBlur={(e) => {
                            const val = parseInt(e.currentTarget.textContent?.replace(/,/g, '') || '0', 10) || 0;
                            handleUpdateStubRow(idx, 'sheets', val);
                          }}
                        >
                          {isCalculated ? Number(row.sheets).toLocaleString() : '-'}
                        </td>
                        <td
                          className="border-r border-b border-black px-1.5 py-0.5 align-middle font-siemreap text-center text-[12pt] bg-white outline-none"
                          style={{ fontSize: '12pt' }}
                          contentEditable={isWordEditMode}
                          suppressContentEditableWarning={true}
                          onBlur={(e) => {
                            const val = e.currentTarget.textContent?.trim() || '';
                            handleUpdateStubRow(idx, 'remarks', val);
                          }}
                        >
                          {row.remarks || ''}
                        </td>
                      </tr>
                    ))}
                    {/* Total Row */}
                    <tr className="font-bold bg-white">
                      <td colSpan={2} className="border-r border-b border-black px-1.5 py-1 align-middle text-center font-moul text-[12pt] bg-white" style={{ fontSize: '12pt' }}>
                        សរុប
                      </td>
                      <td className="border-r border-b border-black px-1.5 py-1 align-middle font-times font-bold text-center text-[12pt] bg-white" style={{ fontSize: '12pt' }}>
                        {isCalculated ? totalStubBooklets.toLocaleString() : '-'}
                      </td>
                      <td className="border-r border-b border-black px-1.5 py-1 align-middle font-times font-bold text-center text-[12pt] bg-white" style={{ fontSize: '12pt' }}>
                        {isCalculated ? totalStubSheets.toLocaleString() : '-'}
                      </td>
                      <td className="border-r border-b border-black px-1.5 py-1 align-middle font-siemreap text-center text-[12pt] bg-white" style={{ fontSize: '12pt' }}>
                      </td>
                    </tr>
                  </tbody>
                </table>
              </div>
            ) : (
              <>
                {/* Section ក: សន្លឹកទិដ្ឋាការស្អិត */}
                <div className="mb-1">
                  <h2 className="font-moul text-[12pt] leading-tight mb-0.5 text-gray-950" style={{ fontSize: '12pt' }}>
                    ក. សន្លឹកទិដ្ឋាការស្អិត
                  </h2>

                  <table className="w-full text-center border-separate border-spacing-0 border-t border-l border-black text-[8pt] leading-tight bg-white">
                    <thead>
                      <tr className="bg-white font-normal text-[6.5pt]">
                        <th rowSpan={2} className="border-r border-b border-black px-0.5 py-0.5 align-middle w-[22px] font-siemreap font-normal bg-white">
                          ល.រ
                        </th>
                        <th rowSpan={2} className="border-r border-b border-black px-0.5 py-0.5 align-middle w-[32px] font-siemreap font-normal bg-white">
                          ប្រភេទ
                        </th>
                        <th colSpan={2} className="border-r border-b border-black px-0.5 py-0.5 align-middle font-siemreap font-normal bg-white">
                          សន្និធិ
                          <br />
                          ចុងគ្រា {prevSolar.khmerDay}-{prevSolar.khmerMonth}-{prevSolar.khmerYear}
                        </th>
                        <th colSpan={7} className="border-r border-b border-black px-0.5 py-0.5 align-middle font-siemreap font-normal bg-white">
                          {startSolar.khmerDay}-{startSolar.khmerMonth}-{startSolar.khmerYear} រហូតដល់ {endSolar.khmerDay}-{endSolar.khmerMonth}-{endSolar.khmerYear}
                        </th>
                        <th colSpan={2} className="border-r border-b border-black px-0.5 py-0.5 align-middle font-siemreap font-normal bg-white">
                          សន្និធិ
                          <br />
                          សល់ {endSolar.khmerDay}-{endSolar.khmerMonth}-{endSolar.khmerYear}
                        </th>
                        <th rowSpan={2} className="border-r border-b border-black px-0.5 py-0.5 align-middle font-siemreap font-normal w-[54px] bg-white">
                          សន្និធិសរុប
                          <br />
                          ក២និងក្រុម
                        </th>
                      </tr>
                      <tr className="bg-white text-[6.5pt] font-normal">
                        {/* សន្និធិ ចុងគ្រា */}
                        <th className="border-r border-b border-black px-0.5 py-0.5 align-middle font-siemreap font-normal bg-white">ក២</th>
                        <th className="border-r border-b border-black px-0.5 py-0.5 align-middle font-siemreap font-normal bg-white">ក្រុម</th>

                        {/* ចលនា (7 ជួរឈរ) */}
                        <th className="border-r border-b border-black px-0.5 py-0.5 align-middle font-siemreap font-normal bg-white">ទិដ្ឋាការបើក<br />ពីក១</th>
                        <th className="border-r border-b border-black px-0.5 py-0.5 align-middle font-siemreap font-normal bg-white">ទិដ្ឋាការបង្វិល<br />ពីក្រុម</th>
                        <th className="border-r border-b border-black px-0.5 py-0.5 align-middle font-siemreap font-normal bg-white">ទិដ្ឋាការបើក<br />ផ្តល់ទៅក្រុម</th>
                        <th className="border-r border-b border-black px-0.5 py-0.5 align-middle font-siemreap font-normal bg-white">ក្រុម<br />ប្រើប្រាស់</th>
                        <th className="border-r border-b border-black px-0.5 py-0.5 align-middle font-siemreap font-normal bg-white">ទិដ្ឋាការក្រុម<br />មិនបានការ</th>
                        <th className="border-r border-b border-black px-0.5 py-0.5 align-middle font-siemreap font-normal bg-white">មិនបានការ<br />ក២</th>
                        <th className="border-r border-b border-black px-0.5 py-0.5 align-middle font-siemreap font-normal bg-white">ទិដ្ឋាការសាក<br />ល្បង ក២</th>

                        {/* សន្និធិ សល់ */}
                        <th className="border-r border-b border-black px-0.5 py-0.5 align-middle font-siemreap font-normal bg-white">ក២</th>
                        <th className="border-r border-b border-black px-0.5 py-0.5 align-middle font-siemreap font-normal bg-white">ក្រុម</th>
                      </tr>
                    </thead>

                    <tbody>
                      {stickerData.rows.map((row, idx) => (
                        <tr key={row.visaType}>
                          <td className="border-r border-b border-black px-0.5 py-0.5 align-middle text-center font-times text-[8.5pt] bg-white">
                            {idx + 1}
                          </td>
                          <td className="border-r border-b border-black px-0.5 py-0.5 align-middle font-bold font-times text-center text-[8.5pt] bg-white">
                            {row.visaType}
                          </td>
                          <td className="border-r border-b border-black px-1 py-0.5 align-middle text-right font-times text-[8.5pt] bg-white">
                            {renderNum(row.openK2)}
                          </td>
                          <td className="border-r border-b border-black px-1 py-0.5 align-middle text-right font-times text-[8.5pt] bg-white">
                            {renderNum(row.openTeams)}
                          </td>
                          <td className="border-r border-b border-black px-1 py-0.5 align-middle text-right font-times text-[8.5pt] bg-white">
                            {renderNum(row.k1ToK2)}
                          </td>
                          <td className="border-r border-b border-black px-1 py-0.5 align-middle text-right font-times text-[8.5pt] bg-white">
                            {renderNum(row.teamReturnedToK2)}
                          </td>
                          <td className="border-r border-b border-black px-1 py-0.5 align-middle text-right font-times text-[8.5pt] bg-white">
                            {renderNum(row.issuedK2ToTeams)}
                          </td>
                          <td className="border-r border-b border-black px-1 py-0.5 align-middle text-right font-times text-[8.5pt] bg-white">
                            {renderNum(row.teamsUsed)}
                          </td>
                          <td className="border-r border-b border-black px-1 py-0.5 align-middle text-right font-times text-[8.5pt] bg-white">
                            {renderNum(row.damagedTeam)}
                          </td>
                          <td className="border-r border-b border-black px-1 py-0.5 align-middle text-right font-times text-[8.5pt] bg-white">
                            {renderNum(row.damagedK2)}
                          </td>
                          <td className="border-r border-b border-black px-1 py-0.5 align-middle text-right font-times text-[8.5pt] bg-white">
                            {renderNum(row.testSampleK2)}
                          </td>
                          <td className="border-r border-b border-black px-1 py-0.5 align-middle text-right font-times text-[8.5pt] bg-white">
                            {renderNum(row.endingK2)}
                          </td>
                          <td className="border-r border-b border-black px-1 py-0.5 align-middle text-right font-times text-[8.5pt] bg-white">
                            {renderNum(row.endingTeams)}
                          </td>
                          <td className="border-r border-b border-black px-1 py-0.5 align-middle text-right font-times text-[8.5pt] font-bold bg-white">
                            {renderNum(row.grandTotalEnding)}
                          </td>
                        </tr>
                      ))}

                      {/* Grand Total Row */}
                      <tr className="font-bold">
                        <td colSpan={2} className="border-r border-b border-black px-1 py-0.5 align-middle text-center font-moul text-[8.5pt] bg-white">
                          សរុប
                        </td>
                        <td className="border-r border-b border-black px-1 py-0.5 align-middle text-right font-times font-bold text-[8.5pt] bg-white">
                          {renderNum(stickerData.totalRow.openK2)}
                        </td>
                        <td className="border-r border-b border-black px-1 py-0.5 align-middle text-right font-times font-bold text-[8.5pt] bg-white">
                          {renderNum(stickerData.totalRow.openTeams)}
                        </td>
                        <td className="border-r border-b border-black px-1 py-0.5 align-middle text-right font-times font-bold text-[8.5pt] bg-white">
                          {renderNum(stickerData.totalRow.k1ToK2)}
                        </td>
                        <td className="border-r border-b border-black px-1 py-0.5 align-middle text-right font-times font-bold text-[8.5pt] bg-white">
                          {renderNum(stickerData.totalRow.teamReturnedToK2)}
                        </td>
                        <td className="border-r border-b border-black px-1 py-0.5 align-middle text-right font-times font-bold text-[8.5pt] bg-white">
                          {renderNum(stickerData.totalRow.issuedK2ToTeams)}
                        </td>
                        <td className="border-r border-b border-black px-1 py-0.5 align-middle text-right font-times font-bold text-[8.5pt] bg-white">
                          {renderNum(stickerData.totalRow.teamsUsed)}
                        </td>
                        <td className="border-r border-b border-black px-1 py-0.5 align-middle text-right font-times font-bold text-[8.5pt] bg-white">
                          {renderNum(stickerData.totalRow.damagedTeam)}
                        </td>
                        <td className="border-r border-b border-black px-1 py-0.5 align-middle text-right font-times font-bold text-[8.5pt] bg-white">
                          {renderNum(stickerData.totalRow.damagedK2)}
                        </td>
                        <td className="border-r border-b border-black px-1 py-0.5 align-middle text-right font-times font-bold text-[8.5pt] bg-white">
                          {renderNum(stickerData.totalRow.testSampleK2)}
                        </td>
                        <td className="border-r border-b border-black px-1 py-0.5 align-middle text-right font-times font-bold text-[8.5pt] bg-white">
                          {renderNum(stickerData.totalRow.endingK2)}
                        </td>
                        <td className="border-r border-b border-black px-1 py-0.5 align-middle text-right font-times font-bold text-[8.5pt] bg-white">
                          {renderNum(stickerData.totalRow.endingTeams)}
                        </td>
                        <td className="border-r border-b border-black px-1 py-0.5 align-middle text-right font-times font-bold text-[8.5pt] bg-white">
                          {renderNum(stickerData.totalRow.grandTotalEnding)}
                        </td>
                      </tr>
                    </tbody>
                  </table>
                </div>

                {/* Section ខ: ក្រដាសអនុម័តទិដ្ឋាការអេឡិចត្រូនិក (4 equal columns matching user image 100%) */}
                <div className="mb-1">
                  <h2 className="font-moul text-[12pt] leading-tight mb-0.5 text-gray-950" style={{ fontSize: '12pt' }}>
                    ខ. ក្រដាសអនុម័តទិដ្ឋាការអេឡិចត្រូនិក
                  </h2>

                  <table className="w-full table-fixed text-center border-separate border-spacing-0 border-t border-l border-black text-[8pt] leading-tight bg-white">
                    <thead>
                      <tr className="bg-white font-normal text-[6.5pt]">
                        <th className="border-r border-b border-black px-1 py-0.5 align-middle font-siemreap font-normal w-1/4 bg-white">
                          សន្និធិ
                          <br />
                          ដើមគ្រា {prevSolar.khmerDay}-{prevSolar.khmerMonth}-{prevSolar.khmerYear}
                        </th>
                        <th className="border-r border-b border-black px-1 py-0.5 align-middle font-siemreap font-normal w-1/4 bg-white">
                          បញ្ចូលស្តុក
                        </th>
                        <th className="border-r border-b border-black px-1 py-0.5 align-middle font-siemreap font-normal w-1/4 bg-white">
                          បើកផ្តល់ទៅតាមក្រុម
                          <br />
                          ផ្តល់ទិដ្ឋាការអេឡិចត្រូនិច
                        </th>
                        <th className="border-r border-b border-black px-1 py-0.5 align-middle font-siemreap font-normal w-1/4 bg-white">
                          សន្និធិនៅសល់
                          <br />
                          ត្រឹម {endSolar.khmerDay}-{endSolar.khmerMonth}-{endSolar.khmerYear}
                        </th>
                      </tr>
                    </thead>
                    <tbody>
                      <tr>
                        <td className="border-r border-b border-black px-3 py-0.5 align-middle font-times text-center text-[8.5pt] w-1/4 bg-white">
                          {renderNum(eVisaData.openBundles)}
                        </td>
                        <td className="border-r border-b border-black px-3 py-0.5 align-middle font-times text-center text-[8.5pt] w-1/4 bg-white">
                          {renderNum(eVisaData.receivedK1Bundles)}
                        </td>
                        <td className="border-r border-b border-black px-3 py-0.5 align-middle font-times text-center text-[8.5pt] w-1/4 bg-white">
                          {renderNum(eVisaData.issuedToTeamsBundles)}
                        </td>
                        <td className="border-r border-b border-black px-3 py-0.5 align-middle font-times text-center text-[8.5pt] font-bold w-1/4 bg-white">
                          {renderNum(eVisaData.endingBundles)}
                        </td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              </>
            )}

            {/* Concluding text */}
            <div className="text-[12pt] font-siemreap space-y-0.5 mt-1 mb-0.5 leading-normal" style={{ fontSize: '12pt' }}>
              <p style={{ textIndent: `${activeTabStop ? activeTabStop + 'cm' : '2.25cm'}` }} className="text-justify">
                {proposalType === 'office' ? (
                  <>
                    អាស្រ័យដូចបានជម្រាបជូនខាងលើ សូម<span className="font-moul text-[12pt]" style={{ fontSize: '12pt' }}>{docRecipientRank || 'លោកឧត្តមសេនីយ៍ទោ'} {docRecipientTitle || 'ប្រធាននាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត'}</span> មេត្តា ពិនិត្យ និងសម្រេចដោយក្តីអនុគ្រោះដ៏ខ្ពង់ខ្ពស់ ។
                  </>
                ) : proposalType === 'section' ? (
                  <>
                    អាស្រ័យដូចបានជម្រាបជូនខាងលើ សូម<span className="font-moul text-[12pt]" style={{ fontSize: '12pt' }}>លោកវរសេនីយ៍ទោ នាយការិយាល័យទិដ្ឋាការចូល</span> មេត្តា ពិនិត្យ និងសម្រេចដោយក្តីអនុគ្រោះ ។
                  </>
                ) : (
                  <>
                    អាស្រ័យដូចបានជម្រាបជូនខាងលើ សូម <span className="font-moul text-[12pt]" style={{ fontSize: '12pt' }}>លោកវរសេនីយ៍ទោ នាយការិយាល័យទិដ្ឋាការចូល</span> មេត្តា ពិនិត្យ និងជ្រាបជារបាយការណ៍ដោយក្តីអនុគ្រោះ ។
                  </>
                )}
              </p>
              <p style={{ textIndent: `${activeTabStop ? activeTabStop + 'cm' : '2.25cm'}` }} className="text-justify">
                {proposalType === 'office' ? (
                  <>
                    សូម<span className="font-moul text-[12pt]" style={{ fontSize: '12pt' }}>{docRecipientRank || 'លោកឧត្តមសេនីយ៍ទោ'}</span> មេត្តាទទួលនូវការគោរពដ៏ខ្ពង់ខ្ពស់ អំពីខ្ញុំបាទ ។
                  </>
                ) : (
                  <>
                    សូម <span className="font-moul text-[12pt]" style={{ fontSize: '12pt' }}>លោកនាយការិយាល័យ</span> មេត្តាទទួលនូវការគោរព ពីខ្ញុំបាទ ។
                  </>
                )}
              </p>
            </div>

            {/* Signature section (2 Columns matching user PDF 100% with clean non-overlapping positioning) */}
            <div
              data-signature-section="true"
              className="grid grid-cols-2 gap-4 text-center text-[9.5pt] font-siemreap items-start"
              style={{ marginTop: `${signSectionMarginTop}pt` }}
            >
              {/* Left Signature */}
              <div
                data-signature-block="left"
                className="space-y-[1px] leading-tight relative group/leftsign"
                style={{
                  position: 'relative',
                  left: `${leftSignShiftX}cm`,
                  top: `${leftSignShiftY}pt`,
                }}
              >
                {/* On-screen quick shift adjuster (hidden on print, contentEditable=false) */}
                <div
                  contentEditable={false}
                  data-no-print="true"
                  onMouseDown={(e) => { e.preventDefault(); e.stopPropagation(); }}
                  className="print:hidden select-none opacity-0 group-hover/leftsign:opacity-100 transition-opacity absolute -top-8 left-1/2 -translate-x-1/2 flex items-center gap-1 bg-slate-900/95 text-white text-[10px] px-2 py-1 rounded-md shadow-xl border border-slate-700 z-30 whitespace-nowrap"
                >
                  <span className="text-amber-300 font-semibold">ឆ្វេង:</span>
                  <button
                    type="button"
                    onMouseDown={(e) => { e.preventDefault(); e.stopPropagation(); }}
                    onClick={(e) => {
                      e.preventDefault();
                      e.stopPropagation();
                      updateSignaturePositions(parseFloat((leftSignShiftX - 0.2).toFixed(1)));
                    }}
                    className="px-1.5 py-0.5 bg-slate-800 hover:bg-blue-600 rounded text-white font-bold cursor-pointer transition active:scale-95"
                    title="រំកិលទៅឆ្វេង 0.2cm"
                  >
                    ◀
                  </button>
                  <span className="font-mono text-emerald-300 font-bold px-0.5">{leftSignShiftX}cm</span>
                  <button
                    type="button"
                    onMouseDown={(e) => { e.preventDefault(); e.stopPropagation(); }}
                    onClick={(e) => {
                      e.preventDefault();
                      e.stopPropagation();
                      updateSignaturePositions(parseFloat((leftSignShiftX + 0.2).toFixed(1)));
                    }}
                    className="px-1.5 py-0.5 bg-slate-800 hover:bg-blue-600 rounded text-white font-bold cursor-pointer transition active:scale-95"
                    title="រំកិលទៅស្តាំ 0.2cm"
                  >
                    ▶
                  </button>
                  <span className="text-slate-500 mx-0.5">|</span>
                  <button
                    type="button"
                    onMouseDown={(e) => { e.preventDefault(); e.stopPropagation(); }}
                    onClick={(e) => {
                      e.preventDefault();
                      e.stopPropagation();
                      updateSignaturePositions(undefined, leftSignShiftY - 2);
                    }}
                    className="px-1.5 py-0.5 bg-slate-800 hover:bg-blue-600 rounded text-white font-bold cursor-pointer transition active:scale-95"
                    title="រំកិលឡើងលើ 2pt"
                  >
                    ▲
                  </button>
                  <span className="font-mono text-cyan-300 font-bold px-0.5">{leftSignShiftY}pt</span>
                  <button
                    type="button"
                    onMouseDown={(e) => { e.preventDefault(); e.stopPropagation(); }}
                    onClick={(e) => {
                      e.preventDefault();
                      e.stopPropagation();
                      updateSignaturePositions(undefined, leftSignShiftY + 2);
                    }}
                    className="px-1.5 py-0.5 bg-slate-800 hover:bg-blue-600 rounded text-white font-bold cursor-pointer transition active:scale-95"
                    title="រំកិលចុះក្រោម 2pt"
                  >
                    ▼
                  </button>
                  <span className="text-slate-500 mx-0.5">|</span>
                  <button
                    type="button"
                    onMouseDown={(e) => { e.preventDefault(); e.stopPropagation(); }}
                    onClick={(e) => {
                      e.preventDefault();
                      e.stopPropagation();
                      updateSignaturePositions(0, 0);
                    }}
                    className="px-1 py-0.5 bg-slate-800 hover:bg-red-600 rounded text-amber-300 hover:text-white font-bold cursor-pointer transition active:scale-95 text-[9px]"
                    title="កំណត់ទីតាំងដើមឡើងវិញ"
                  >
                    Reset
                  </button>
                </div>

                <p className="leading-tight">បានឃើញ និងគោរពជូន</p>
                {proposalType === 'office' ? (
                  <>
                    <p className="leading-tight">លោកឧត្តមសេនីយ៍ទោ ប្រធាននាយកដ្ឋាន មេត្តាពិនិត្យ</p>
                    <p className="leading-tight">និងសម្រេច ដោយក្តីអនុគ្រោះ។</p>
                  </>
                ) : proposalType === 'section' ? (
                  <>
                    <p className="leading-tight">លោកនាយការិយាល័យទិដ្ឋាការចូល មេត្តាពិនិត្យ</p>
                    <p className="leading-tight">និងសម្រេច ដោយក្តីអនុគ្រោះ។</p>
                  </>
                ) : (
                  <>
                    <p className="leading-tight">លោកនាយការិយាល័យទិដ្ឋាការចូល មេត្តាជ្រាប</p>
                    <p className="leading-tight">ជារបាយការណ៍ ដោយក្តីអនុគ្រោះ។</p>
                  </>
                )}
                <div
                  onClick={() => setShowDocDateModal(true)}
                  className="cursor-pointer hover:bg-amber-50/80 rounded px-1 transition-colors border border-transparent hover:border-amber-300 relative group/datebtn text-center leading-tight my-0.5"
                  title="ចុចត្រង់នេះដើម្បីផ្លាស់ប្តូរ ថ្ងៃខែឆ្នាំតាក់តែងលិខិត (Click to change document date)"
                >
                  {(customSignLunar || signLunar) && (
                    <p className="text-[9pt] leading-tight">
                      {customSignLunar || signLunar}
                    </p>
                  )}
                  <p className="text-[9pt] leading-tight">
                    {customSignSolar || `ភ្នំពេញ, ថ្ងៃទី${signSolar.khmerDay} ខែ${signSolar.khmerMonth} ឆ្នាំ${signSolar.khmerYear}`}
                  </p>
                  <span className="print:hidden opacity-0 group-hover/datebtn:opacity-100 transition-opacity absolute -top-5 left-1/2 -translate-x-1/2 bg-blue-900 text-white text-[9px] px-1.5 py-0.5 rounded shadow whitespace-nowrap z-20 pointer-events-none">
                    📅 ចុចកែប្រែកាលបរិច្ឆេទ
                  </span>
                </div>
                <p className="font-moul text-[12pt] leading-tight mt-0.5" style={{ fontSize: '12pt' }}>{leftSignTitle}</p>
                <div className="h-10 flex items-end justify-center">
                  {leftSignOfficer && <p className="font-bold leading-tight">{leftSignOfficer}</p>}
                </div>
              </div>

              {/* Right Signature */}
              <div
                data-signature-block="right"
                className="space-y-[1px] leading-tight relative group/rightsign"
                style={{
                  position: 'relative',
                  left: `${rightSignShiftX}cm`,
                  top: `${rightSignShiftY}pt`,
                }}
              >
                {/* On-screen quick shift adjuster (hidden on print, contentEditable=false) */}
                <div
                  contentEditable={false}
                  data-no-print="true"
                  onMouseDown={(e) => { e.preventDefault(); e.stopPropagation(); }}
                  className="print:hidden select-none opacity-0 group-hover/rightsign:opacity-100 transition-opacity absolute -top-8 left-1/2 -translate-x-1/2 flex items-center gap-1 bg-slate-900/95 text-white text-[10px] px-2 py-1 rounded-md shadow-xl border border-slate-700 z-30 whitespace-nowrap"
                >
                  <span className="text-amber-300 font-semibold">ស្តាំ:</span>
                  <button
                    type="button"
                    onMouseDown={(e) => { e.preventDefault(); e.stopPropagation(); }}
                    onClick={(e) => {
                      e.preventDefault();
                      e.stopPropagation();
                      updateSignaturePositions(undefined, undefined, parseFloat((rightSignShiftX - 0.2).toFixed(1)));
                    }}
                    className="px-1.5 py-0.5 bg-slate-800 hover:bg-emerald-600 rounded text-white font-bold cursor-pointer transition active:scale-95"
                    title="រំកិលទៅឆ្វេង 0.2cm"
                  >
                    ◀
                  </button>
                  <span className="font-mono text-emerald-300 font-bold px-0.5">{rightSignShiftX}cm</span>
                  <button
                    type="button"
                    onMouseDown={(e) => { e.preventDefault(); e.stopPropagation(); }}
                    onClick={(e) => {
                      e.preventDefault();
                      e.stopPropagation();
                      updateSignaturePositions(undefined, undefined, parseFloat((rightSignShiftX + 0.2).toFixed(1)));
                    }}
                    className="px-1.5 py-0.5 bg-slate-800 hover:bg-emerald-600 rounded text-white font-bold cursor-pointer transition active:scale-95"
                    title="រំកិលទៅស្តាំ 0.2cm"
                  >
                    ▶
                  </button>
                  <span className="text-slate-500 mx-0.5">|</span>
                  <button
                    type="button"
                    onMouseDown={(e) => { e.preventDefault(); e.stopPropagation(); }}
                    onClick={(e) => {
                      e.preventDefault();
                      e.stopPropagation();
                      updateSignaturePositions(undefined, undefined, undefined, rightSignShiftY - 2);
                    }}
                    className="px-1.5 py-0.5 bg-slate-800 hover:bg-emerald-600 rounded text-white font-bold cursor-pointer transition active:scale-95"
                    title="រំកិលឡើងលើ 2pt"
                  >
                    ▲
                  </button>
                  <span className="font-mono text-cyan-300 font-bold px-0.5">{rightSignShiftY}pt</span>
                  <button
                    type="button"
                    onMouseDown={(e) => { e.preventDefault(); e.stopPropagation(); }}
                    onClick={(e) => {
                      e.preventDefault();
                      e.stopPropagation();
                      updateSignaturePositions(undefined, undefined, undefined, rightSignShiftY + 2);
                    }}
                    className="px-1.5 py-0.5 bg-slate-800 hover:bg-emerald-600 rounded text-white font-bold cursor-pointer transition active:scale-95"
                    title="រំកិលចុះក្រោម 2pt"
                  >
                    ▼
                  </button>
                  <span className="text-slate-500 mx-0.5">|</span>
                  <button
                    type="button"
                    onMouseDown={(e) => { e.preventDefault(); e.stopPropagation(); }}
                    onClick={(e) => {
                      e.preventDefault();
                      e.stopPropagation();
                      updateSignaturePositions(undefined, undefined, 0, 0);
                    }}
                    className="px-1 py-0.5 bg-slate-800 hover:bg-red-600 rounded text-amber-300 hover:text-white font-bold cursor-pointer transition active:scale-95 text-[9px]"
                    title="កំណត់ទីតាំងដើមឡើងវិញ"
                  >
                    Reset
                  </button>
                </div>

                <div
                  onClick={() => setShowDocDateModal(true)}
                  className="cursor-pointer hover:bg-amber-50/80 rounded px-1 transition-colors border border-transparent hover:border-amber-300 relative group/datebtn text-center leading-tight my-0.5"
                  title="ចុចត្រង់នេះដើម្បីផ្លាស់ប្តូរ ថ្ងៃខែឆ្នាំតាក់តែងលិខិត (Click to change document date)"
                >
                  {(customSignLunar || signLunar) && (
                    <p className="text-[9pt] leading-tight">
                      {customSignLunar || signLunar}
                    </p>
                  )}
                  <p className="text-[9pt] leading-tight">
                    {customSignSolar || `ភ្នំពេញ, ថ្ងៃទី${signSolar.khmerDay} ខែ${signSolar.khmerMonth} ឆ្នាំ${signSolar.khmerYear}`}
                  </p>
                  <span className="print:hidden opacity-0 group-hover/datebtn:opacity-100 transition-opacity absolute -top-5 left-1/2 -translate-x-1/2 bg-blue-900 text-white text-[9px] px-1.5 py-0.5 rounded shadow whitespace-nowrap z-20 pointer-events-none">
                    📅 ចុចកែប្រែកាលបរិច្ឆេទ
                  </span>
                </div>
                <p className="font-moul text-[12pt] leading-tight mt-0.5" style={{ fontSize: '12pt' }}>{rightSignTitle}</p>
                <div className="h-10 flex items-end justify-center">
                  {rightSignOfficer && <p className="font-bold leading-tight">{rightSignOfficer}</p>}
                </div>
              </div>
            </div>
          </>
        )}
      </div>
    </div>
      </div>

      {/* Microsoft Word Font Dialog Modal (Ctrl+D) */}
      {showFontDialog && (
        <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-[2px] flex items-center justify-center p-3">
          <div className="bg-[#F0F0F0] text-gray-900 rounded shadow-2xl border border-gray-400 w-full max-w-[500px] text-[11px] select-none font-sans overflow-hidden">
            {/* Title Bar */}
            <div className="flex items-center justify-between px-3 py-1.5 bg-gradient-to-r from-gray-100 to-gray-200 border-b border-gray-300">
              <div className="flex items-center gap-1.5 font-semibold text-gray-800 text-xs">
                <span className="text-blue-600 font-bold">A</span>
                <span>Font</span>
              </div>
              <div className="flex items-center gap-1">
                <button
                  onClick={() => setShowFontDialog(false)}
                  className="w-5 h-5 flex items-center justify-center rounded hover:bg-red-500 hover:text-white text-gray-600 font-bold transition"
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
                Advanced
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
                        <label className="text-gray-700">Scale:</label>
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
                        <label className="text-gray-700">Spacing:</label>
                        <select
                          value={characterSpacing}
                          onChange={(e) => setCharacterSpacing(e.target.value as any)}
                          className="bg-white border border-gray-300 rounded px-2 py-0.5 text-xs text-black focus:outline-none focus:border-blue-500"
                        >
                          <option value="Normal">Normal</option>
                          <option value="Expanded">Expanded</option>
                          <option value="Condensed">Condensed</option>
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
                        <label className="text-gray-700">Position:</label>
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
                        <option value="Khmer Mool1">Khmer Mool1</option>
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
                        : dialogFontFamily.includes('Mool') || dialogFontFamily.includes('Moul')
                        ? `'Khmer Mool1', 'Khmer Mool 1', 'Khmer Mool', 'Khmer OS Mool1', 'Khmer OS Muol Light', serif`
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
                    Sample របក.សរុបការងារស្តុក 123456789 ទិដ្ឋាការ
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
                <button
                  type="button"
                  className="px-2.5 py-1 bg-white hover:bg-gray-50 text-gray-800 rounded border border-gray-300 font-medium text-[11px] shadow-sm cursor-pointer transition"
                >
                  Text Effects...
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
                          onClick={() => setTempMargins({ top: 0.5, bottom: 0.5, left: 2.8, right: 1.5 })}
                          className="px-2 py-0.5 bg-blue-100 hover:bg-blue-200 text-blue-900 border border-blue-300 rounded text-[10.5px] font-semibold cursor-pointer"
                        >
                          ⭐ ស្តង់ដារលំនាំដើម (ឆ្វេង 2.8 / ស្តាំ 1.5 / លើ 0.5 / ក្រោម 0.5)
                        </button>
                        <button
                          type="button"
                          onClick={() => setTempMargins({ top: 1.2, bottom: 1.2, left: 2.5, right: 1.5 })}
                          className="px-2 py-0.5 bg-gray-200 hover:bg-gray-300 text-gray-800 border border-gray-300 rounded text-[10.5px] cursor-pointer"
                        >
                          🇰🇭 ស្តង់ដាររដ្ឋបាល (2.5 / 1.5 / 1.2)
                        </button>
                        <button
                          type="button"
                          onClick={() => setTempMargins({ top: 2.54, bottom: 2.54, left: 2.54, right: 2.54 })}
                          className="px-2 py-0.5 bg-gray-200 hover:bg-gray-300 text-gray-800 border border-gray-300 rounded text-[10.5px] cursor-pointer"
                        >
                          ធម្មតា Normal (2.54cm)
                        </button>
                        <button
                          type="button"
                          onClick={() => setTempMargins({ top: 1.27, bottom: 1.27, left: 1.27, right: 1.27 })}
                          className="px-2 py-0.5 bg-gray-200 hover:bg-gray-300 text-gray-800 border border-gray-300 rounded text-[10.5px] cursor-pointer"
                        >
                          ត្បិត Narrow (1.27cm)
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

                  <fieldset className="border border-gray-300 rounded p-2.5 bg-[#F7F7F7]">
                    <legend className="text-gray-700 px-1 font-semibold text-[11px]">
                      Scaling (ពង្រីក/បង្រួមទំព័រ - Adjust Page)
                    </legend>
                    <div className="space-y-2 text-xs">
                      <div className="flex items-center justify-between">
                        <label className="text-gray-900 font-bold">Adjust to (ទំហំទំព័រពេលបោះពុម្ព/PDF) :</label>
                        <div className="flex items-center gap-1.5">
                          <input
                            type="number"
                            step="1"
                            min="50"
                            max="150"
                            value={tempAdjustPagePercent}
                            onChange={(e) => setTempAdjustPagePercent(parseInt(e.target.value) || 97)}
                            className="w-16 bg-blue-50 border border-blue-400 font-bold rounded px-1.5 py-0.5 text-right font-mono text-blue-900"
                          />
                          <span className="text-gray-700 font-medium">% normal size</span>
                        </div>
                      </div>
                      <div className="flex flex-wrap items-center gap-1.5 pt-1">
                        <button
                          type="button"
                          onClick={() => setTempAdjustPagePercent(97)}
                          className={`px-2 py-0.5 rounded text-[10.5px] font-semibold cursor-pointer border ${
                            tempAdjustPagePercent === 97
                              ? 'bg-blue-600 text-white border-blue-700'
                              : 'bg-blue-100 hover:bg-blue-200 text-blue-900 border-blue-300'
                          }`}
                        >
                          ⭐ 97% (ស្តង់ដារលំនាំដើម)
                        </button>
                        <button
                          type="button"
                          onClick={() => setTempAdjustPagePercent(100)}
                          className={`px-2 py-0.5 rounded text-[10.5px] cursor-pointer border ${
                            tempAdjustPagePercent === 100
                              ? 'bg-blue-600 text-white border-blue-700 font-semibold'
                              : 'bg-gray-200 hover:bg-gray-300 text-gray-800 border-gray-300'
                          }`}
                        >
                          100% (ធម្មតា)
                        </button>
                        <button
                          type="button"
                          onClick={() => setTempAdjustPagePercent(95)}
                          className={`px-2 py-0.5 rounded text-[10.5px] cursor-pointer border ${
                            tempAdjustPagePercent === 95
                              ? 'bg-blue-600 text-white border-blue-700 font-semibold'
                              : 'bg-gray-200 hover:bg-gray-300 text-gray-800 border-gray-300'
                          }`}
                        >
                          95%
                        </button>
                        <button
                          type="button"
                          onClick={() => setTempAdjustPagePercent(90)}
                          className={`px-2 py-0.5 rounded text-[10.5px] cursor-pointer border ${
                            tempAdjustPagePercent === 90
                              ? 'bg-blue-600 text-white border-blue-700 font-semibold'
                              : 'bg-gray-200 hover:bg-gray-300 text-gray-800 border-gray-300'
                          }`}
                        >
                          90%
                        </button>
                      </div>
                      <p className="text-[10px] text-gray-500 italic">
                        * កំណត់ទំហំ 97% ជួយឱ្យតារាង និងហត្ថលេខាក្នុងទំព័ររបាយការណ៍បោះពុម្ព ឬទាញយកជា PDF សមល្មមស្អាតលើ ១ទំព័រ A4 មិនធ្លាក់ទៅទំព័រទី២។
                      </p>
                    </div>
                  </fieldset>

                  <fieldset className="border border-gray-300 rounded p-2.5 bg-[#F7F7F7]">
                    <legend className="text-gray-700 px-1 font-semibold text-[11px]">
                      Paper Source
                    </legend>
                    <div className="grid grid-cols-2 gap-2 text-xs">
                      <div>
                        <label className="text-gray-600 block mb-1">First page :</label>
                        <select className="w-full bg-white border border-gray-300 rounded px-1.5 py-0.5 text-black">
                          <option>Default tray (Auto Select)</option>
                          <option>Tray 1</option>
                          <option>Tray 2</option>
                        </select>
                      </div>
                      <div>
                        <label className="text-gray-600 block mb-1">Other pages :</label>
                        <select className="w-full bg-white border border-gray-300 rounded px-1.5 py-0.5 text-black">
                          <option>Default tray (Auto Select)</option>
                          <option>Tray 1</option>
                          <option>Tray 2</option>
                        </select>
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
                          onClick={() => setActiveTabStop(1.2)}
                          className="px-2 py-0.5 bg-gray-200 hover:bg-gray-300 text-gray-800 rounded text-[10px] cursor-pointer"
                        >
                          1.2cm
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
                  {/* Scaled Visual Page representation */}
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
                    {/* Top simulated lines */}
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
                  const defaultM = { ...tempMargins };
                  setCustomMargins(defaultM);
                  setAdjustPagePercent(tempAdjustPagePercent);
                  try {
                    const prefStr = localStorage.getItem('robok_report_user_preferences');
                    const pref = prefStr ? JSON.parse(prefStr) : {};
                    pref.customMargins = defaultM;
                    pref.adjustPagePercent = tempAdjustPagePercent;
                    localStorage.setItem('robok_report_user_preferences', JSON.stringify(pref));
                    localStorage.setItem('robok_total_margins', JSON.stringify(defaultM));
                    localStorage.setItem('robok_margins', JSON.stringify(defaultM));
                    localStorage.setItem('robok_total_adjust_scale', tempAdjustPagePercent.toString());
                  } catch (e) {}
                  alert(`បានកំណត់គែមក្រដាស និងទំហំទំព័រ Adjust Page=${tempAdjustPagePercent}% ជាលំនាំដើមជារៀងរហូត!`);
                }}
                className="px-2.5 py-1 bg-white hover:bg-gray-50 text-gray-800 rounded border border-gray-300 font-medium text-[11px] shadow-sm cursor-pointer transition"
              >
                Set As Default (កំណត់ជាលំនាំដើម)
              </button>

              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => {
                    setCustomMargins({ ...tempMargins });
                    setAdjustPagePercent(tempAdjustPagePercent);
                    try {
                      localStorage.setItem('robok_total_adjust_scale', tempAdjustPagePercent.toString());
                      const prefStr = localStorage.getItem('robok_report_user_preferences');
                      const pref = prefStr ? JSON.parse(prefStr) : {};
                      pref.adjustPagePercent = tempAdjustPagePercent;
                      pref.customMargins = { ...tempMargins };
                      localStorage.setItem('robok_report_user_preferences', JSON.stringify(pref));
                    } catch (e) {}
                    setShowPageSetupDialog(false);
                    setTimeout(() => {
                      handlePrint();
                    }, 200);
                  }}
                  className="px-4 py-1 bg-slate-800 hover:bg-slate-900 text-white rounded border border-slate-700 font-semibold text-xs shadow-sm cursor-pointer transition flex items-center gap-1.5"
                  title="រក្សាទុកការកំណត់ និងបើកផ្ទាំងបោះពុម្ព (Print)"
                >
                  <Printer className="w-3.5 h-3.5 text-blue-300" />
                  <span>Print (បោះពុម្ព)</span>
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setCustomMargins({ ...tempMargins });
                    setAdjustPagePercent(tempAdjustPagePercent);
                    try {
                      localStorage.setItem('robok_total_adjust_scale', tempAdjustPagePercent.toString());
                      const prefStr = localStorage.getItem('robok_report_user_preferences');
                      const pref = prefStr ? JSON.parse(prefStr) : {};
                      pref.adjustPagePercent = tempAdjustPagePercent;
                      localStorage.setItem('robok_report_user_preferences', JSON.stringify(pref));
                    } catch (e) {}
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

      {/* Document Date Manager Modal */}
      {showDocDateModal && (
        <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-[2px] flex items-center justify-center p-3">
          <div className="bg-white text-gray-900 rounded-xl shadow-2xl border border-gray-300 w-full max-w-[540px] text-xs font-sans overflow-hidden animate-in fade-in zoom-in-95 duration-150">
            {/* Header */}
            <div className="flex items-center justify-between px-4 py-3 bg-gradient-to-r from-blue-900 to-indigo-900 text-white">
              <div className="flex items-center gap-2 font-bold text-sm">
                <span>📅</span>
                <span>កំណត់ថ្ងៃខែឆ្នាំតាក់តែងលិខិត (Document Creation Date)</span>
              </div>
              <button
                onClick={() => setShowDocDateModal(false)}
                className="w-6 h-6 flex items-center justify-center rounded-full hover:bg-white/20 text-white font-bold transition cursor-pointer"
                title="បិទ (Close)"
              >
                ✕
              </button>
            </div>

            {/* Body */}
            <div className="p-4 space-y-4">
              <div>
                <label className="block text-xs font-bold text-gray-700 mb-1.5">
                  ជ្រើសរើសកាលបរិច្ឆេទតាក់តែង (Pick Date) ៖
                </label>
                <div className="flex items-center gap-2">
                  <div className="flex-1">
                    <CustomDatePicker
                      value={docSignDate}
                      onChange={handleDocSignDateChange}
                      className="text-sm bg-blue-50/60 border-2 border-blue-400 rounded-lg py-1.5 text-blue-950 font-bold"
                    />
                  </div>
                  <button
                    type="button"
                    onClick={() => {
                      const today = new Date().toISOString().split('T')[0];
                      handleDocSignDateChange(today);
                    }}
                    className="px-3 py-2 bg-blue-100 hover:bg-blue-200 text-blue-800 rounded-lg font-bold text-xs cursor-pointer transition border border-blue-300"
                  >
                    ថ្ងៃនេះ
                  </button>
                  <button
                    type="button"
                    onClick={() => {
                      if (endDate) {
                        handleDocSignDateChange(endDate);
                      }
                    }}
                    className="px-3 py-2 bg-indigo-100 hover:bg-indigo-200 text-indigo-800 rounded-lg font-bold text-xs cursor-pointer transition border border-indigo-300"
                    title={`តាមថ្ងៃបញ្ចប់របាយការណ៍ (${endDate})`}
                  >
                    តាមថ្ងៃបញ្ចប់
                  </button>
                </div>
              </div>

              {/* Preview Cards */}
              <div className="bg-slate-50 border border-slate-200 rounded-xl p-3.5 space-y-3">
                <div className="text-[11px] font-bold text-slate-600 uppercase tracking-wider">
                  ទិដ្ឋភាពបង្ហាញក្នុងហត្ថលេខា (Signature Preview) ៖
                </div>

                {/* Lunar Date */}
                <div className="space-y-1">
                  <div className="flex items-center justify-between">
                    <label className="text-[11px] font-semibold text-slate-700">
                      កាលបរិច្ឆេទចន្ទគតិ (Khmer Lunar Date) ៖
                    </label>
                    {customSignLunar && (
                      <button
                        type="button"
                        onClick={() => setCustomSignLunar('')}
                        className="text-[10px] text-amber-600 hover:underline cursor-pointer"
                      >
                        ប្រើស្វ័យប្រវត្តឡើងវិញ
                      </button>
                    )}
                  </div>
                  <input
                    type="text"
                    value={customSignLunar !== '' ? customSignLunar : signLunar}
                    onChange={(e) => setCustomSignLunar(e.target.value)}
                    placeholder="ឧទាហរណ៍៖ ថ្ងៃចន្ទ ១១កើត ខែស្រាពណ៍ ឆ្នាំម្សាញ់ សំរឹទ្ធិស័ក ព.ស២៥៧០"
                    className="w-full text-xs bg-white border border-gray-300 rounded-md px-2.5 py-1.5 text-gray-800 focus:outline-none focus:border-blue-500 font-medium"
                  />
                </div>

                {/* Solar Date */}
                <div className="space-y-1">
                  <div className="flex items-center justify-between">
                    <label className="text-[11px] font-semibold text-slate-700">
                      កាលបរិច្ឆេទសុរិយគតិ (Khmer Solar Date) ៖
                    </label>
                    {customSignSolar && (
                      <button
                        type="button"
                        onClick={() => setCustomSignSolar('')}
                        className="text-[10px] text-amber-600 hover:underline cursor-pointer"
                      >
                        ប្រើស្វ័យប្រវត្តឡើងវិញ
                      </button>
                    )}
                  </div>
                  <input
                    type="text"
                    value={
                      customSignSolar !== ''
                        ? customSignSolar
                        : viewMode === 'team'
                        ? `${selectedTeam ? (selectedTeam.includes('កំពង់ផែ') || selectedTeam.includes('ព្រំដែន') || selectedTeam.includes('អាកាស') ? selectedTeam : 'ច្រកទ្វារ' + selectedTeam) : 'ភ្នំពេញ'}, ថ្ងៃទី${signSolar.khmerDay} ខែ${signSolar.khmerMonth} ឆ្នាំ${signSolar.khmerYear}`
                        : `ភ្នំពេញ, ថ្ងៃទី${signSolar.khmerDay} ខែ${signSolar.khmerMonth} ឆ្នាំ${signSolar.khmerYear}`
                    }
                    onChange={(e) => setCustomSignSolar(e.target.value)}
                    placeholder={viewMode === 'team' ? "ឧទាហរណ៍៖ អាកាស ព្រះសីហនុ, ថ្ងៃទី... ខែ... ឆ្នាំ..." : "ឧទាហរណ៍៖ ភ្នំពេញ, ថ្ងៃទី... ខែ... ឆ្នាំ..."}
                    className="w-full text-xs bg-white border border-gray-300 rounded-md px-2.5 py-1.5 text-gray-800 focus:outline-none focus:border-blue-500 font-medium"
                  />
                </div>
              </div>
            </div>

            {/* Footer */}
            <div className="flex items-center justify-end gap-2 px-4 py-3 bg-gray-100 border-t border-gray-200">
              <button
                type="button"
                onClick={() => {
                  setShowDocDateModal(false);
                }}
                className="px-5 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-lg font-bold text-xs shadow cursor-pointer transition"
              >
                យល់ព្រម (Done)
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Floating Save Success Toast Notification (Microsoft Word Style) */}
      {showSaveSuccessToast && (
        <div className="fixed bottom-6 right-6 z-50 bg-slate-900 text-white px-4 py-3 rounded-xl shadow-2xl border border-emerald-500/60 flex items-center gap-3 animate-bounce">
          <div className="w-9 h-9 rounded-full bg-emerald-500/20 text-emerald-400 flex items-center justify-center shrink-0 border border-emerald-500/40">
            <Save className="w-5 h-5" />
          </div>
          <div>
            <p className="text-xs font-bold text-emerald-400 flex items-center gap-1.5">
              <span>✓ បានរក្សាទុកទិន្នន័យឯកសារជោគជ័យ!</span>
            </p>
            <p className="text-[11px] text-gray-300 font-['Khmer_OS_Siemreap','Siemreap',sans-serif] mt-0.5">{lastSavedTime}</p>
          </div>
        </div>
      )}
    </div>
  );
};
