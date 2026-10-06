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
  Check,
  ZoomIn,
  ZoomOut,
  Sparkles,
  Search,
  X,
  RefreshCw,
  ListChecks,
  Save,
  HardDrive,
  Eye,
  EyeOff,
} from 'lucide-react';
import { getKhmerLunarDate, getKhmerSolarParts, toKhmerNum } from '../utils/khmerCalendar';
import { exportSinglePageA4LandscapePdf } from '../utils/pdfExportHelper';
import { printA4Document } from '../utils/printHelper';
import { CustomDatePicker } from './CustomDatePicker';
import { normalizeDateToISO, normalizeVisaType } from '../utils/teamNormalization';
import { isCeaRecord, isOldStockTeamRecord } from '../utils/teamStockCalculation';
import {
  TacteingLine,
  TacteingControlSelector,
  saveTacteingSettings,
  getSavedTacteingSettings,
  TacteingType,
} from './TacteingLine';
import { idbStorage } from '../utils/idbStorage';
import { YearlyMonthChecklistModal } from './YearlyMonthChecklistModal';
import { YearlyDamagedMissingChecklistModal } from './YearlyDamagedMissingChecklistModal';
import { YearlyReceivedK2ChecklistModal } from './YearlyReceivedK2ChecklistModal';

// Helper to check if a record is issued to teams (self-contained for K2 report)
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
    op === 'receive_from_k1' ||
    op === 'receivek1' ||
    op.includes('បើកពីk1') ||
    op.includes('បើកពី k1') ||
    op.includes('បើកពីក១') ||
    op.includes('បើកពី ក១') ||
    op.includes('ទទួលពីក១') ||
    op.includes('ទទួលពី k1') ||
    op.includes('openk1')
  ) {
    return false;
  }

  // Exclude used, damaged, return, transfer
  if (
    op === 'useteam' ||
    op === 'use_team' ||
    op === 'used' ||
    op.includes('ប្រើប្រាស់') ||
    op === 'damagedteam' ||
    op === 'damagedk2' ||
    op === 'damaged' ||
    op.includes('ខូច') ||
    op === 'returnteam' ||
    op === 'return' ||
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
    op.includes('បើកផ្តល់') ||
    op.includes('បើកផ្ដល់') ||
    op.includes('ចែកជូន')
  );
};

// Helper to check if a record is a transfer deduction (ផ្ទេរការប្រើប្រាស់)
export const isTransferDeductionRecord = (r: StockRecord): boolean => {
  if (r.stockType && r.stockType !== 'sticker') return false;
  if (isCeaRecord(r)) return false;
  if (isOldStockTeamRecord(r)) return false;
  const op = (r.operationType || '').trim().toLowerCase();
  const notes = ((r.note || (r as any).notes || '') as string).trim().toLowerCase();
  const source = (((r as any).sourceFrom || '') as string).trim().toLowerCase();
  const reason = (((r as any).reason || '') as string).trim().toLowerCase();
  return (
    op === 'transferteam' ||
    op === 'transfer_team' ||
    op === 'transfer' ||
    op === 'transfer_usage' ||
    op === 'transferuseteam' ||
    op.includes('transfer') ||
    op.includes('ផ្ទេរ') ||
    notes.includes('ផ្ទេរ') ||
    source.includes('ផ្ទេរ') ||
    reason.includes('ផ្ទេរ')
  );
};

export interface YearlyK2ReceivedK1ReportProps {
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

export interface RowK1ReceivedData {
  visaType: string;
  openingStock: number; // សន្និធិចុងគ្រា ក២ (e.g. 30 Nov 2018 baseline)
  monthlyValues: number[]; // 12 months received from K1 (Dec, Jan, Feb, ..., Nov)
  totalMonthly: number; // សរុប ១២ខែ
  totalAvailable: number; // ចំនួនសរុប (សន្និធិចុងគ្រា + សរុប ១២ខែ)
  issuedToTeams: number; // បើកផ្តល់ជូនក្រុម
  damagedK2: number; // មិនបានការ សាកល្បងក២
  returnStock: number; // ទិដ្ឋាការបង្វិលពីក្រុម
  remaining: number; // សន្និធិនៅសល់ចុងគ្រា
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

// Verified official baseline data for K2 receipts from K1 (Dec 2018 - Nov 2019)
export const OFFICIAL_K2_RECEIVE_2018_2019: Record<string, {
  opening: number;
  months: number[];
  issuedToTeams: number;
  damagedK2: number;
  returnStock: number;
}> = {
  T: {
    opening: 304000,
    months: [0, 0, 0, 0, 400000, 0, 400000, 0, 400000, 200000, 400000, 0],
    issuedToTeams: 2130050,
    damagedK2: 0,
    returnStock: 0,
  },
  T1: {
    opening: 3200,
    months: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    issuedToTeams: 0,
    damagedK2: 0,
    returnStock: 0,
  },
  T2: {
    opening: 6200,
    months: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    issuedToTeams: 0,
    damagedK2: 0,
    returnStock: 0,
  },
  T3: {
    opening: 5200,
    months: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    issuedToTeams: 0,
    damagedK2: 0,
    returnStock: 0,
  },
  E: {
    opening: 330750,
    months: [0, 0, 0, 0, 200000, 0, 200000, 0, 0, 0, 200000, 0],
    issuedToTeams: 708450,
    damagedK2: 0,
    returnStock: 0,
  },
  E1: {
    opening: 2650,
    months: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    issuedToTeams: 0,
    damagedK2: 0,
    returnStock: 0,
  },
  E2: {
    opening: 4700,
    months: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    issuedToTeams: 0,
    damagedK2: 0,
    returnStock: 0,
  },
  E3: {
    opening: 4950,
    months: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    issuedToTeams: 0,
    damagedK2: 0,
    returnStock: 0,
  },
  D: {
    opening: 3600,
    months: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    issuedToTeams: 0,
    damagedK2: 0,
    returnStock: 0,
  },
  K: {
    opening: 16750,
    months: [0, 0, 0, 0, 20000, 0, 0, 0, 0, 0, 0, 0],
    issuedToTeams: 26000,
    damagedK2: 0,
    returnStock: 0,
  },
  A: {
    opening: 6100,
    months: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    issuedToTeams: 0,
    damagedK2: 0,
    returnStock: 0,
  },
  B: {
    opening: 7250,
    months: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    issuedToTeams: 0,
    damagedK2: 0,
    returnStock: 0,
  },
  C: {
    opening: 4850,
    months: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
    issuedToTeams: 0,
    damagedK2: 0,
    returnStock: 0,
  },
};

// Check if a record is K2 receiving from K1
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

// Check if a record is K2 Damaged or K2 Test (មិនបានការ សាកល្បងក២ = ទិដ្ឋាការខូចក២ + ទិដ្ឋាការសាកក២)
export const isDamagedOrTestK2Record = (r: StockRecord): boolean => {
  if (r.stockType && r.stockType !== 'sticker') return false;
  if (isCeaRecord(r)) return false;
  if (isOldStockTeamRecord(r)) return false;
  const op = (r.operationType || '').trim().toLowerCase();
  const src = (r.sourceFrom || '').trim().toLowerCase();
  const notes = (r.note || (r as any).notes || '').trim().toLowerCase();

  // Exclude operations that belong purely to team usage/distribution
  if (
    op === 'damagedteam' ||
    op === 'missingteam' ||
    op === 'useteam' ||
    op === 'issueteam' ||
    op === 'returnteam' ||
    op === 'oldstockteam' ||
    op === 'transferteam'
  ) {
    return false;
  }

  // 1. Check K2 Damaged (ទិដ្ឋាការខូចក២ / មិនបានការ)
  const isDamagedK2 =
    op === 'damagedk2' ||
    op === 'damaged_k2' ||
    op === 'voidk2' ||
    op === 'damagedoffice' ||
    op === 'voidoffice' ||
    op === 'invalidk2' ||
    op === 'damaged' ||
    op === 'invalid' ||
    op.includes('ខូចក២') ||
    op.includes('មិនបានការក២') ||
    op.includes('ក២ខូច') ||
    op.includes('ខូចក') ||
    op.includes('មិនបានការ') ||
    src.includes('ខូចក២') ||
    src.includes('មិនបានការ') ||
    src.includes('ខូចក') ||
    src === 'ទិដ្ឋាការខូចក២' ||
    src === 'ខូចក២' ||
    notes.includes('ខូចក២') ||
    notes.includes('មិនបានការក២');

  // 2. Check K2 Test / Sample (ទិដ្ឋាការសាកក២ / បោះពុម្ពសាកល្បងក២)
  const isTestK2 =
    op === 'testprintk2' ||
    op === 'test_print_k2' ||
    op === 'testk2' ||
    op === 'test_k2' ||
    op === 'samplek2' ||
    op === 'sample_k2' ||
    op === 'testoffice' ||
    op === 'sampleoffice' ||
    op === 'testprint' ||
    op === 'sample' ||
    op === 'test' ||
    op.includes('សាកក') ||
    op.includes('សាកល្បង') ||
    op.includes('សាក') ||
    src.includes('សាក') ||
    src.includes('បោះពុម្ពសាកល្បង') ||
    src.includes('ទិដ្ឋាការសាកក២') ||
    src.includes('សាកក២') ||
    notes.includes('សាក') ||
    notes.includes('បោះពុម្ពសាកល្បង');

  return isDamagedK2 || isTestK2;
};

// Backward-compatibility alias
export const isDamagedK2Record = isDamagedOrTestK2Record;

// Check if a record is Team return to K2 (ទិដ្ឋាការបង្វិលពីក្រុម)
export const isReturnFromTeamRecord = (r: StockRecord): boolean => {
  if (r.stockType && r.stockType !== 'sticker') return false;
  if (isCeaRecord(r)) return false;
  if (isOldStockTeamRecord(r)) return false;
  const op = (r.operationType || '').trim().toLowerCase();

  return (
    op === 'returnteam' ||
    op === 'returnoffice' ||
    op === 'return_team' ||
    op === 'return_office' ||
    op === 'recycle' ||
    op.includes('បង្វិល') ||
    op.includes('បង្វិលកង') ||
    op.includes('បង្វិលពីក្រុម')
  );
};

export const YearlyK2ReceivedK1Report: React.FC<YearlyK2ReceivedK1ReportProps> = ({
  stockRecords,
  categories,
  officers = [],
  currentRole,
  userName,
  assignedTeam,
  onClose,
}) => {
  const { scaleMode } = useWorkspaceSettings();
  const DRAFT_STORAGE_KEY = 'app_yearly_k2_receive_k1_draft_v6';
  const IDB_DRAFT_KEY = 'yearly_k2_receive_k1_report_draft_v6';

  // Read draft from localStorage if available
  const savedDraft = useMemo(() => {
    try {
      const raw = localStorage.getItem(DRAFT_STORAGE_KEY);
      if (raw) {
        const parsed = JSON.parse(raw);
        if (parsed && typeof parsed === 'object') return parsed;
      }
    } catch {}
    return null;
  }, []);

  // Fiscal / Period Configuration (Default Dec 2018 to Nov 2019)
  const [startYear, setStartYear] = useState<number>(() => savedDraft?.startYear ?? 2018);
  const [startMonth, setStartMonth] = useState<number>(() => savedDraft?.startMonth ?? 12); // Month index 1-12 (12 = Dec)
  const [appliedStartYear, setAppliedStartYear] = useState<number>(() => savedDraft?.appliedStartYear ?? 2018);
  const [appliedStartMonth, setAppliedStartMonth] = useState<number>(() => savedDraft?.appliedStartMonth ?? 12);
  const [isCalculating, setIsCalculating] = useState<boolean>(false);
  const [hasCalculated, setHasCalculated] = useState<boolean>(false);
  const [reportDate, setReportDate] = useState<string>(() => savedDraft?.reportDate ?? '2019-12-08'); // Signature date
  const [isSavedRecently, setIsSavedRecently] = useState<boolean>(false);

  // Check if date selection has unapplied changes
  const isFilterPending = startYear !== appliedStartYear || startMonth !== appliedStartMonth;

  // Header and Metadata Customization
  const [ministryName, setMinistryName] = useState<string>('ក្រសួងមហាផ្ទៃ');
  const [departmentName, setDepartmentName] = useState<string>('អគ្គនាយកដ្ឋានអន្តោប្រវេសន៍');
  const [generalDeptName, setGeneralDeptName] = useState<string>('នាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត');
  const [officeName, setOfficeName] = useState<string>('ការិយាល័យទិដ្ឋាការចូល');
  const [sectionName, setSectionName] = useState<string>('ផ្នែករដ្ឋបាល');
  const [reportTitle, setReportTitle] = useState<string>('តារាងសន្លឹកទិដ្ឋាការស្អិត តាមប្រភេទនីមួយៗ បើកពី ក១');

  // Signer Customization
  const [signerRole, setSignerRole] = useState<string>(() => savedDraft?.signerRole ?? 'អ្នកធ្វើតារាង');
  const [signerName, setSignerName] = useState<string>(() => savedDraft?.signerName ?? 'អនុសេនីយ៍ឯក ច្រេង ថុល');

  // Auto calculate vs manual cell override
  const [autoCalculate, setAutoCalculate] = useState<boolean>(true);

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
      const saved = localStorage.getItem('yearly_k2_received_k1_row_height');
      if (saved !== null) {
        const parsed = parseInt(saved, 10);
        if (!isNaN(parsed) && parsed >= 14 && parsed <= 30) return parsed;
      }
    } catch {}
    return 19;
  });
  const [nameColor, setNameColor] = useState<string>('#C00000');
  const [dateNumberFormat, setDateNumberFormat] = useState<'khmer' | 'latin'>(() => savedDraft?.dateNumberFormat ?? 'khmer');

  // Tacteing Customization State
  const [tacteingSettings, setTacteingSettings] = useState<{ type: TacteingType; customImage: string | null }>(
    getSavedTacteingSettings()
  );
  const [showTacteingSelector, setShowTacteingSelector] = useState<boolean>(false);

  // Toggle visibility of the formatting/actions toolbar (Zoom, Fonts, Auto Sum, Recalculate, Reset)
  const [showPictureToolbar, setShowPictureToolbar] = useState<boolean>(() => {
    try {
      const saved = localStorage.getItem('yearly_k2_received_k1_show_toolbar');
      return saved !== null ? saved === 'true' : true;
    } catch {
      return true;
    }
  });

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

  const [isExportingPdf, setIsExportingPdf] = useState<boolean>(false);
  const printContainerRef = useRef<HTMLDivElement>(null);

  // Modals
  const [isChecklistOpen, setIsChecklistOpen] = useState<boolean>(false);
  const [selectedModalMonthIdx, setSelectedModalMonthIdx] = useState<number>(0);
  const [isDamagedModalOpen, setIsDamagedModalOpen] = useState<boolean>(false);
  const [selectedDamagedVisaType, setSelectedDamagedVisaType] = useState<string>('ALL');
  const [isIssuedModalOpen, setIsIssuedModalOpen] = useState<boolean>(false);

  // Effective Officers list
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

  const getOfficerFormattedRankAndName = (of: Officer): string => {
    const raw = of.nameKh || of.name || of.nameEn || '';
    const knownKhmerRanks = [
      'ឧត្តមសេនីយ៍ឯក', 'ឧត្តមសេនីយ៍ទោ', 'ឧត្តមសេនីយ៍ត្រី',
      'វរសេនីយ៍ឯក', 'វរសេនីយ៍ទោ', 'វរសេនីយ៍ត្រី',
      'អនុសេនីយ៍ឯក', 'អនុសេនីយ៍ទោ', 'អនុសេនីយ៍ត្រី',
      'នាយចំណង់', 'ពលបាលឯក', 'ពលបាលទោ', 'ពលបាលត្រី'
    ];
    
    if (knownKhmerRanks.some((kr) => raw.trim().startsWith(kr))) {
      return raw.trim();
    }
    
    let rankTitle = '';
    if (of.rankId && categories?.ranks) {
      const r = categories.ranks.find((rk) => rk.id === of.rankId);
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
          <span
            className="font-siemreap font-bold text-[12pt]"
            style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif", fontSize: '12pt' }}
          >
            {matchedRank}
          </span>
          {namePart && (
            <span
              className="font-moul font-normal text-[12pt]"
              style={{
                color: nameColor || '#C00000',
                fontFamily: "'Khmer Mool1', 'Khmer Mool', 'Khmer OS Muol Light', 'Khmer OS Muol', 'Moul', serif",
                fontSize: '12pt',
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
          <span
            className="font-siemreap font-bold text-[12pt]"
            style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif", fontSize: '12pt' }}
          >
            {firstWord}
          </span>
          <span
            className="font-moul font-normal text-[12pt]"
            style={{
              color: nameColor || '#C00000',
              fontFamily: "'Khmer Mool1', 'Khmer Mool', 'Khmer OS Muol Light', 'Khmer OS Muol', 'Moul', serif",
              fontSize: '12pt',
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
        style={{ color: nameColor || '#C00000', fontSize: '12pt' }}
      >
        {trimmed}
      </span>
    );
  };

  // 12 Months computation based on appliedStartYear and appliedStartMonth
  const monthColumnsInfo = useMemo(() => {
    const cols = [];
    let curYear = appliedStartYear;
    let curMonth = appliedStartMonth;

    for (let i = 0; i < 12; i++) {
      cols.push({
        monthNum: curMonth,
        year: curYear,
        name: KHMER_MONTHS_NAMES[curMonth - 1],
      });
      curMonth++;
      if (curMonth > 12) {
        curMonth = 1;
        curYear++;
      }
    }
    return cols;
  }, [appliedStartYear, appliedStartMonth]);

  // Calculate table matrix for K2 receipts from K1 using REAL stock transactions
  const calculateK2ReportData = (
    year: number,
    month: number,
    forceSampleBaseline = false
  ): Record<string, RowK1ReceivedData> => {
    const result: Record<string, RowK1ReceivedData> = {};

    // 12 Months window
    const cols: { monthNum: number; year: number; startIso: string; endIso: string }[] = [];
    let curY = year;
    let curM = month;
    for (let i = 0; i < 12; i++) {
      const mStr = String(curM).padStart(2, '0');
      const daysInM = new Date(curY, curM, 0).getDate();
      cols.push({
        monthNum: curM,
        year: curY,
        startIso: `${curY}-${mStr}-01`,
        endIso: `${curY}-${mStr}-${String(daysInM).padStart(2, '0')}`,
      });
      curM++;
      if (curM > 12) {
        curM = 1;
        curY++;
      }
    }

    const startDateStr = cols[0].startIso;
    const endDateStr = cols[11].endIso;

    REPORT_VISA_TYPES.forEach((vt) => {
      const baseOpening = OFFICIAL_K2_DEC_2018_BASELINE[vt] || 0;
      const base2018 = OFFICIAL_K2_RECEIVE_2018_2019[vt];

      result[vt] = {
        visaType: vt,
        openingStock: year === 2018 && month === 12 && base2018 ? base2018.opening : baseOpening,
        monthlyValues: Array(12).fill(0),
        totalMonthly: 0,
        totalAvailable: 0,
        issuedToTeams: 0,
        damagedK2: 0,
        returnStock: 0,
        remaining: 0,
      };
    });

    let totalRealK1ReceivedCount = 0;

    // Populate from real stock transactions
    (stockRecords || []).forEach((r) => {
      if (r.stockType && r.stockType !== 'sticker') return;
      if (isCeaRecord(r)) return;
      if (isOldStockTeamRecord(r)) return;

      const vt = normalizeVisaType(r.visaType || '');
      if (!result[vt]) return;

      const qty = Number(
        r.quantityBundles ||
          r.totalSheets ||
          (r as any).quantity ||
          ((r as any).count ? (r as any).count * 50 : 0) ||
          0
      );
      if (!qty) return;

      const dateIso = normalizeDateToISO(r.date || (r as any).createdAt || '');
      if (!dateIso) return;

      // Handle prior transactions before current period to accurately build Opening Stock
      if (dateIso < startDateStr && dateIso > '2018-11-30') {
        if (isReceiveFromK1Record(r)) {
          result[vt].openingStock += qty;
        } else if (isReturnFromTeamRecord(r)) {
          result[vt].openingStock += qty;
        } else if (isIssueTeamRecord(r)) {
          result[vt].openingStock -= qty;
        } else if (isTransferDeductionRecord(r)) {
          // Ignore team transfers for Office stock
        } else if (isDamagedOrTestK2Record(r)) {
          result[vt].openingStock -= qty;
        }
        return;
      }

      // Inside current 12-month period
      if (dateIso >= startDateStr && dateIso <= endDateStr) {
        const mIdx = cols.findIndex((c) => dateIso >= c.startIso && dateIso <= c.endIso);

        if (isReceiveFromK1Record(r)) {
          if (mIdx !== -1) {
            result[vt].monthlyValues[mIdx] += qty;
            totalRealK1ReceivedCount += qty;
          }
        } else if (isReturnFromTeamRecord(r)) {
          result[vt].returnStock += qty;
        } else if (isIssueTeamRecord(r)) {
          result[vt].issuedToTeams += qty;
        } else if (isTransferDeductionRecord(r)) {
          // Ignore team transfers for Office stock
        } else if (isDamagedOrTestK2Record(r)) {
          result[vt].damagedK2 += qty;
        }
      }
    });

    // No fallback to sample baseline to keep table empty/0 as requested by the user when no data is uploaded yet
    if (forceSampleBaseline) {
      REPORT_VISA_TYPES.forEach((vt) => {
        const base2018 = OFFICIAL_K2_RECEIVE_2018_2019[vt];
        const isRowEmpty = result[vt].monthlyValues.every((v) => v === 0);
        if (base2018 && (isRowEmpty || forceSampleBaseline)) {
          result[vt].openingStock = base2018.opening;
          result[vt].monthlyValues = [...base2018.months];
          if (result[vt].issuedToTeams === 0 && base2018.issuedToTeams > 0) {
            result[vt].issuedToTeams = base2018.issuedToTeams;
          }
          if (result[vt].damagedK2 === 0 && base2018.damagedK2 > 0) {
            result[vt].damagedK2 = base2018.damagedK2;
          }
          if (result[vt].returnStock === 0 && base2018.returnStock > 0) {
            result[vt].returnStock = base2018.returnStock;
          }
        }
      });
    }

    // Compute totals and balance for each visa type
    REPORT_VISA_TYPES.forEach((vt) => {
      const row = result[vt];
      row.totalMonthly = row.monthlyValues.reduce((sum, v) => sum + v, 0);
      row.totalAvailable = row.openingStock + row.totalMonthly;
      row.remaining = row.totalAvailable - row.issuedToTeams - row.damagedK2 + row.returnStock;
    });

    return result;
  };

  // Primary Table State: dictionary of row data by visa type (starts clean until user clicks "បង្ហាញទិន្នន័យ")
  const [tableData, setTableData] = useState<Record<string, RowK1ReceivedData>>(() => {
    const initial: Record<string, RowK1ReceivedData> = {};
    REPORT_VISA_TYPES.forEach((vt) => {
      initial[vt] = {
        visaType: vt,
        openingStock: 0,
        monthlyValues: Array(12).fill(0),
        totalMonthly: 0,
        totalAvailable: 0,
        issuedToTeams: 0,
        damagedK2: 0,
        returnStock: 0,
        remaining: 0,
      };
    });
    return initial;
  });

  // Restore draft from IndexedDB on mount if available (metadata only; data shows on button click)
  useEffect(() => {
    let isMounted = true;
    idbStorage
      .getItem(IDB_DRAFT_KEY, 'data')
      .then((draft: any) => {
        if (!isMounted || !draft) return;
        if (draft.startYear) setStartYear(draft.startYear);
        if (draft.startMonth) setStartMonth(draft.startMonth);
        if (draft.appliedStartYear) setAppliedStartYear(draft.appliedStartYear);
        if (draft.appliedStartMonth) setAppliedStartMonth(draft.appliedStartMonth);
        if (draft.signerRole) setSignerRole(draft.signerRole);
        if (draft.signerName) setSignerName(draft.signerName);
        if (draft.reportDate) setReportDate(draft.reportDate);
        if (draft.dateNumberFormat) setDateNumberFormat(draft.dateNumberFormat);
      })
      .catch((err) => {
        console.warn('IDB draft load error:', err);
      });
    return () => {
      isMounted = false;
    };
  }, []);

  // Helper to persist draft to localStorage and IndexedDB
  const persistDraft = (
    currentData: Record<string, RowK1ReceivedData>,
    overrideExtras?: Partial<{
      startYear: number;
      startMonth: number;
      appliedStartYear: number;
      appliedStartMonth: number;
      reportDate: string;
      signerName: string;
      signerRole: string;
      dateNumberFormat: 'khmer' | 'latin';
      hasCalculated: boolean;
    }>
  ) => {
    try {
      const payload = {
        tableData: currentData,
        startYear: overrideExtras?.startYear ?? startYear,
        startMonth: overrideExtras?.startMonth ?? startMonth,
        appliedStartYear: overrideExtras?.appliedStartYear ?? appliedStartYear,
        appliedStartMonth: overrideExtras?.appliedStartMonth ?? appliedStartMonth,
        reportDate: overrideExtras?.reportDate ?? reportDate,
        signerName: overrideExtras?.signerName ?? signerName,
        signerRole: overrideExtras?.signerRole ?? signerRole,
        dateNumberFormat: overrideExtras?.dateNumberFormat ?? dateNumberFormat,
        hasCalculated: overrideExtras?.hasCalculated ?? hasCalculated,
        updatedAt: new Date().toISOString(),
      };
      localStorage.setItem(DRAFT_STORAGE_KEY, JSON.stringify(payload));
      idbStorage.setItem(IDB_DRAFT_KEY, 'data', payload);
    } catch (e) {
      console.warn('Could not persist K2 receive draft:', e);
    }
  };

  // Direct manual save draft
  const handleSaveDraft = () => {
    persistDraft(tableData, { hasCalculated: true });
    setIsSavedRecently(true);
    setTimeout(() => setIsSavedRecently(false), 2500);
  };

  // Apply Filter Action (calculate from real stock records for the selected period)
  const handleApplyFilter = () => {
    setIsCalculating(true);
    setTimeout(() => {
      setAppliedStartYear(startYear);
      setAppliedStartMonth(startMonth);
      const computed = calculateK2ReportData(startYear, startMonth, false);
      setTableData(computed);
      setIsCalculating(false);
      setHasCalculated(true);
      persistDraft(computed, {
        startYear,
        startMonth,
        appliedStartYear: startYear,
        appliedStartMonth: startMonth,
        hasCalculated: true,
      });
    }, 150);
  };

  // Quick Preset Selection
  const handleSelectPeriodPreset = (yr: number, m: number, forceSample = false) => {
    setStartYear(yr);
    setStartMonth(m);
    setAppliedStartYear(yr);
    setAppliedStartMonth(m);
    setIsCalculating(true);
    setTimeout(() => {
      const computed = calculateK2ReportData(yr, m, forceSample);
      setTableData(computed);
      setIsCalculating(false);
      setHasCalculated(true);
      persistDraft(computed, {
        startYear: yr,
        startMonth: m,
        appliedStartYear: yr,
        appliedStartMonth: m,
        hasCalculated: true,
      });
    }, 100);
  };

  const handleLoadOfficial2018 = () => {
    handleSelectPeriodPreset(2018, 12, false);
  };

  // Sync All 12 Months directly from actual stock transactions
  const handleSyncAllRealStock = () => {
    setIsCalculating(true);
    setTimeout(() => {
      const computed = calculateK2ReportData(appliedStartYear, appliedStartMonth, false);
      setTableData(computed);
      setIsCalculating(false);
      setHasCalculated(true);
      persistDraft(computed, { hasCalculated: true });
    }, 150);
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
          const totalAvailable = (updated[vt].openingStock || 0) + totalMonthly;
          const remaining =
            totalAvailable - (updated[vt].issuedToTeams || 0) - (updated[vt].damagedK2 || 0) + (updated[vt].returnStock || 0);
          updated[vt] = {
            ...updated[vt],
            monthlyValues: nextMonthly,
            totalMonthly,
            totalAvailable,
            remaining,
          };
        }
      });
      persistDraft(updated, { hasCalculated: true });
      return updated;
    });
    setHasCalculated(true);
  };

  // Apply actual Issued to Teams (បើកផ្តល់តាមក្រុម - ផ្ទេរការប្រើប្រាស់) from Checklist Modal
  const handleApplyAllIssuedValues = (visaValues: Record<string, number>) => {
    setTableData((prev) => {
      const updated = { ...prev };
      REPORT_VISA_TYPES.forEach((vt) => {
        if (updated[vt] && visaValues[vt] !== undefined) {
          const issued = Number(visaValues[vt]) || 0;
          const totalMonthly = Array.isArray(updated[vt].monthlyValues)
            ? updated[vt].monthlyValues.reduce((s, v) => s + v, 0)
            : 0;
          const totalAvailable = (updated[vt].openingStock || 0) + totalMonthly;
          const remaining =
            totalAvailable - issued - (updated[vt].damagedK2 || 0) + (updated[vt].returnStock || 0);
          updated[vt] = {
            ...updated[vt],
            issuedToTeams: issued,
            remaining,
          };
        }
      });
      persistDraft(updated, { hasCalculated: true });
      return updated;
    });
    setHasCalculated(true);
  };

  // Cell Edit Handler (support contentEditable text)
  const handleCellChange = (
    visaType: string,
    field:
      | 'openingStock'
      | 'month'
      | 'issuedToTeams'
      | 'damagedK2'
      | 'returnStock'
      | 'remaining'
      | 'totalMonthly'
      | 'totalAvailable',
    value: string,
    monthIndex?: number
  ) => {
    const cleaned = value.replace(/,/g, '').trim();
    const num = cleaned === '' || cleaned === '-' ? 0 : parseInt(cleaned, 10);
    if (isNaN(num)) return;

    setTableData((prev) => {
      const current = prev[visaType] || {
        visaType,
        openingStock: 0,
        monthlyValues: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
        totalMonthly: 0,
        totalAvailable: 0,
        issuedToTeams: 0,
        damagedK2: 0,
        returnStock: 0,
        remaining: 0,
      };

      const updatedRow = { ...current, monthlyValues: [...current.monthlyValues] };

      if (field === 'openingStock') {
        updatedRow.openingStock = num;
      } else if (field === 'month' && monthIndex !== undefined) {
        updatedRow.monthlyValues[monthIndex] = num;
      } else if (field === 'issuedToTeams') {
        updatedRow.issuedToTeams = num;
      } else if (field === 'damagedK2') {
        updatedRow.damagedK2 = num;
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
        updatedRow.totalAvailable = (updatedRow.openingStock || 0) + updatedRow.totalMonthly;
        updatedRow.remaining =
          updatedRow.totalAvailable - (updatedRow.issuedToTeams || 0) - (updatedRow.damagedK2 || 0) + (updatedRow.returnStock || 0);
      }

      const nextTable = {
        ...prev,
        [visaType]: updatedRow,
      };
      persistDraft(nextTable, { hasCalculated: true });
      return nextTable;
    });
  };

  // Re-Calculate all totals on demand
  const handleRecalculateAll = () => {
    setTableData((prev) => {
      const next: Record<string, RowK1ReceivedData> = {};
      REPORT_VISA_TYPES.forEach((vt) => {
        const r = prev[vt] || {
          visaType: vt,
          openingStock: 0,
          monthlyValues: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
          totalMonthly: 0,
          totalAvailable: 0,
          issuedToTeams: 0,
          damagedK2: 0,
          returnStock: 0,
          remaining: 0,
        };
        const totalMonthly = r.monthlyValues.reduce((a, b) => a + b, 0);
        const totalAvailable = (r.openingStock || 0) + totalMonthly;
        const remaining = totalAvailable - (r.issuedToTeams || 0) - (r.damagedK2 || 0) + (r.returnStock || 0);

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
    const cleared: Record<string, RowK1ReceivedData> = {};
    REPORT_VISA_TYPES.forEach((vt) => {
      cleared[vt] = {
        visaType: vt,
        openingStock: 0,
        monthlyValues: [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
        totalMonthly: 0,
        totalAvailable: 0,
        issuedToTeams: 0,
        damagedK2: 0,
        returnStock: 0,
        remaining: 0,
      };
    });
    setTableData(cleared);
  };

  // Grand totals across all 13 visa types
  const grandTotals = useMemo(() => {
    let openingStock = 0;
    const monthlyValues: number[] = Array(12).fill(0);
    let totalMonthly = 0;
    let totalAvailable = 0;
    let issuedToTeams = 0;
    let damagedK2 = 0;
    let returnStock = 0;
    let remaining = 0;

    REPORT_VISA_TYPES.forEach((vt) => {
      const row = tableData[vt];
      if (!row) return;
      openingStock += row.openingStock || 0;
      row.monthlyValues.forEach((val, idx) => {
        monthlyValues[idx] += val || 0;
      });
      totalMonthly += row.totalMonthly || 0;
      totalAvailable += row.totalAvailable || 0;
      issuedToTeams += row.issuedToTeams || 0;
      damagedK2 += row.damagedK2 || 0;
      returnStock += row.returnStock || 0;
      remaining += row.remaining || 0;
    });

    return {
      openingStock,
      monthlyValues,
      totalMonthly,
      totalAvailable,
      issuedToTeams,
      damagedK2,
      returnStock,
      remaining,
    };
  }, [tableData]);

  // Previous Period Date for Opening Stock Label (e.g. "៣០ វិច្ឆិកា ២០១៨")
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

  // Ending Period Date Label for Remaining Stock (e.g. "៣០ វិច្ឆិកា ២០១៩")
  const endingStockLabel = useMemo(() => {
    const lastCol = monthColumnsInfo[11];
    const daysInEndMonth = new Date(lastCol.year, lastCol.monthNum, 0).getDate();
    const khmerDay = toKhmerNum(daysInEndMonth);
    const khmerMonth = lastCol.name;
    const khmerYear = toKhmerNum(lastCol.year);
    return `${khmerDay} ${khmerMonth} ${khmerYear}`;
  }, [monthColumnsInfo]);

  // Solar and Lunar Signatures
  const signatureSolarParts = useMemo(() => {
    try {
      return getKhmerSolarParts(reportDate || '2019-12-08');
    } catch {
      return { day: '០៨', month: 'ធ្នូ', year: '២០១៩', khmerDay: '០៨', khmerYear: '២០១៩' };
    }
  }, [reportDate]);

  const signatureLunarDate = useMemo(() => {
    try {
      return getKhmerLunarDate(reportDate || '2019-12-08');
    } catch {
      return 'ថ្ងៃអាទិត្យ ១២កើត ខែមិគសិរ ឆ្នាំកុរ ឯកស័ក ព.ស២៥៦៣';
    }
  }, [reportDate]);

  // Excel Export
  const handleExportExcel = () => {
    const wb = XLSX.utils.book_new();
    const headers = [
      'ប្រភេទ',
      `សន្និធិមុនគ្រា (${openingStockLabel})`,
      ...monthColumnsInfo.map((m) => m.name),
      'សរុប ១២ខែ',
      'ចំនួនសរុប',
      'បើកផ្តល់ជូនក្រុម',
      'មិនបានការ សាកល្បងក២',
      'ទិដ្ឋាការបង្វិលពីក្រុម',
      `សន្និធិនៅសល់ (${endingStockLabel})`,
    ];

    const rows = REPORT_VISA_TYPES.map((vt) => {
      const r = tableData[vt];
      return [
        vt,
        r.openingStock,
        ...r.monthlyValues,
        r.totalMonthly,
        r.totalAvailable,
        r.issuedToTeams,
        r.damagedK2,
        r.returnStock,
        r.remaining,
      ];
    });

    // Grand total row
    rows.push([
      'សរុប',
      grandTotals.openingStock,
      ...grandTotals.monthlyValues,
      grandTotals.totalMonthly,
      grandTotals.totalAvailable,
      grandTotals.issuedToTeams,
      grandTotals.damagedK2,
      grandTotals.returnStock,
      grandTotals.remaining,
    ]);

    const ws = XLSX.utils.aoa_to_sheet([
      [`ព្រះរាជាណាចក្រកម្ពុជា`],
      [`ជាតិ សាសនា ព្រះមហាក្សត្រ`],
      [`${ministryName} - ${departmentName}`],
      [`${officeName} - ${sectionName}`],
      [`តារាងសន្លឹកទិដ្ឋាការស្អិត តាមប្រភេទនីមួយៗ បើកពី ក១ — ដោយគិតចាប់ពីថ្ងៃទី០១ ខែ${KHMER_MONTHS_NAMES[appliedStartMonth - 1]} ឆ្នាំ${toKhmerNum(appliedStartYear)} រហូតដល់ថ្ងៃទី${toKhmerNum(new Date(monthColumnsInfo[11].year, monthColumnsInfo[11].monthNum, 0).getDate())} ខែ${monthColumnsInfo[11].name} ឆ្នាំ${toKhmerNum(monthColumnsInfo[11].year)}`],
      [],
      headers,
      ...rows,
    ]);

    XLSX.utils.book_append_sheet(wb, ws, 'តារាងទិដ្ឋាការបើកពីក១');
    XLSX.writeFile(wb, `K2_Receive_From_K1_${appliedStartYear}_${appliedStartMonth}.xlsx`);
  };

  // Direct Print
  const handlePrint = () => {
    printA4Document('yearly-k2-received-k1-print-area', {
      orientation: 'landscape',
      documentTitle: `តារាងទិដ្ឋាការបើកពីក១_${appliedStartYear}_${appliedStartYear + 1}`,
    });
  };

  // PDF Export
  const handleExportPdf = async () => {
    if (!printContainerRef.current) return;
    setIsExportingPdf(true);
    try {
      await exportSinglePageA4LandscapePdf(
        printContainerRef.current,
        `តារាងទិដ្ឋាការបើកពីក១_${appliedStartYear}_${appliedStartYear + 1}.pdf`,
        'yearly-k2-received-k1-print-area'
      );
    } catch (err) {
      console.error('PDF Export Error:', err);
    } finally {
      setIsExportingPdf(false);
    }
  };

  // Cell Value Renderer (showing '-' when uncalculated or 0, or formatted number)
  const renderCellValue = (val: number | undefined | null) => {
    if (!hasCalculated) return '-';
    if (val === undefined || val === null || isNaN(Number(val)) || Number(val) === 0) return '-';
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
                តារាងទិដ្ឋាការបើកពីក១ — សន្លឹកទិដ្ឋាការស្អិត
              </h2>
              <p className="text-xs text-gray-500">
                តារាងសន្លឹកទិដ្ឋាការស្អិត តាមប្រភេទនីមួយៗ បើកពីនាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត (ក១) ប្រចាំឆ្នាំ (១២ខែ)
              </p>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2">
            {/* Save Draft Button */}
            <button
              onClick={handleSaveDraft}
              type="button"
              className={`w-32 h-9 rounded-lg text-xs font-semibold flex items-center justify-center gap-1.5 shadow-xs transition cursor-pointer ${
                isSavedRecently
                  ? 'bg-emerald-600 text-white ring-2 ring-emerald-400'
                  : 'bg-amber-600 hover:bg-amber-700 text-white'
              }`}
              title="រក្សាទុកទិន្នន័យតារាងនេះជាសេចក្តីព្រាង (Persistent Draft)"
            >
              {isSavedRecently ? (
                <>
                  <Check className="w-3.5 h-3.5 shrink-0" />
                  <span>បានរក្សាទុក!</span>
                </>
              ) : (
                <>
                  <Save className="w-3.5 h-3.5 shrink-0" />
                  <span>រក្សាទុក</span>
                </>
              )}
            </button>

            {/* Check List / Monthly Breakdown Inspector Button */}
            <button
              onClick={() => {
                setSelectedModalMonthIdx(0);
                setIsChecklistOpen(true);
              }}
              type="button"
              className="w-36 h-9 rounded-lg text-xs font-medium bg-purple-700 hover:bg-purple-800 text-white flex items-center justify-center gap-1.5 shadow-xs transition cursor-pointer"
              title="ពិនិត្យផ្ទៀងផ្ទាត់បញ្ជី និងចំនួនសរុបតាមខែនីមួយៗ"
            >
              <ListChecks className="w-3.5 h-3.5 shrink-0" />
              <span>ពិនិត្យបញ្ជីតាមខែ</span>
            </button>

            {/* Toggle Show/Hide Picture Controls */}
            <button
              type="button"
              onClick={() => {
                const nextVal = !showPictureToolbar;
                setShowPictureToolbar(nextVal);
                try {
                  localStorage.setItem('yearly_k2_received_k1_show_toolbar', String(nextVal));
                } catch {}
              }}
              className={`h-9 px-3 rounded-lg text-xs font-bold border transition flex items-center justify-center gap-1.5 cursor-pointer shadow-xs ${
                showPictureToolbar
                  ? 'bg-slate-800 hover:bg-slate-900 text-white border-slate-900'
                  : 'bg-amber-50 hover:bg-amber-100 text-amber-900 border-amber-300'
              }`}
              title={showPictureToolbar ? 'ចុចដើម្បីលាក់ផ្ទាំងជម្រើស' : 'ចុចដើម្បីបង្ហាញផ្ទាំងជម្រើស'}
            >
              {showPictureToolbar ? (
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

            {/* Export PDF */}
            <button
              onClick={handleExportPdf}
              disabled={isExportingPdf}
              type="button"
              className="w-32 h-9 rounded-lg text-xs font-medium bg-rose-600 hover:bg-rose-700 text-white flex items-center justify-center gap-1.5 shadow-xs transition cursor-pointer disabled:opacity-50"
            >
              <Download className="w-3.5 h-3.5 shrink-0" />
              <span>{isExportingPdf ? 'កំពុងបង្កើត...' : 'ទាញយក PDF'}</span>
            </button>

            {/* Export Excel */}
            <button
              onClick={handleExportExcel}
              type="button"
              className="w-28 h-9 rounded-lg text-xs font-medium bg-emerald-700 hover:bg-emerald-800 text-white flex items-center justify-center gap-1.5 shadow-xs transition cursor-pointer"
            >
              <FileSpreadsheet className="w-3.5 h-3.5 shrink-0" />
              <span>Excel</span>
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
                setHasCalculated(false);
              }}
              className={`w-full bg-slate-50/60 hover:bg-white border rounded-lg px-3 py-2 text-xs text-slate-800 focus:bg-white focus:outline-none focus:ring-2 font-medium transition-all cursor-pointer ${
                isFilterPending || !hasCalculated
                  ? 'border-amber-400 ring-1 ring-amber-300'
                  : 'border-slate-200 hover:border-slate-300 focus:border-amber-500 focus:ring-amber-500/20'
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
                setHasCalculated(false);
              }}
              className={`w-full bg-slate-50/60 hover:bg-white border rounded-lg px-3 py-2 text-xs text-slate-800 focus:bg-white focus:outline-none focus:ring-2 font-medium transition-all cursor-pointer ${
                isFilterPending || !hasCalculated
                  ? 'border-amber-400 ring-1 ring-amber-300'
                  : 'border-slate-200 hover:border-slate-300 focus:border-amber-500 focus:ring-amber-500/20'
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
              onClick={handleApplyFilter}
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

        {/* Action Controls & Formatting Tools (Toggled by showPictureToolbar) */}
        {showPictureToolbar && (
          <div className="flex flex-wrap items-center justify-between gap-3 text-xs pt-1">
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

              <span className="font-semibold text-gray-600 ml-2">ទំហំអក្សរ:</span>
              <select
                value={fontSize}
                onChange={(e) => setFontSize(parseInt(e.target.value, 10))}
                className="bg-white border border-gray-300 rounded px-2 py-1 focus:outline-none"
              >
                <option value={10}>តូចខ្លាំង (10px)</option>
                <option value={11}>ស្តង់ដារ PDF (11px)</option>
                <option value={12}>មធ្យម (12px)</option>
                <option value={13}>ធំល្មម (13px)</option>
                <option value={14}>ធំ (14px)</option>
              </select>

              {/* Compact Table Row Height Control */}
              <div className="flex items-center border border-gray-300 rounded bg-white px-2 py-0.5 text-xs text-gray-700 font-siemreap ml-1">
                <span className="text-[11px] text-gray-600 mr-1 select-none font-medium">កម្ពស់ជួរ៖</span>
                <button
                  type="button"
                  onClick={() => {
                    const val = Math.max(15, tableRowHeight - 1);
                    setTableRowHeight(val);
                    localStorage.setItem('yearly_k2_received_k1_row_height', String(val));
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
                    localStorage.setItem('yearly_k2_received_k1_row_height', String(val));
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

              {/* Date Number Format Selector */}
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

              {/* Tacteing Style Button */}
              <button
                type="button"
                onClick={() => setShowTacteingSelector((prev) => !prev)}
                className={`px-2.5 py-1 rounded border text-xs font-medium flex items-center gap-1.5 transition ml-1 cursor-pointer ${
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
                onClick={handleSyncAllRealStock}
                type="button"
                className="px-2.5 py-1 rounded bg-amber-100 hover:bg-amber-200 text-amber-900 text-xs font-semibold flex items-center gap-1 transition cursor-pointer border border-amber-300"
                title="ទាញទិន្នន័យពីប្រតិបត្តិការស្តុកជាក់ស្តែងទាំង ១២ខែ មកគណនាស្វ័យប្រវត្តិ"
              >
                <Sparkles className="w-3 h-3 text-amber-700" />
                <span>ទាញទិន្នន័យជាក់ស្តែង</span>
              </button>

              <button
                onClick={handleRecalculateAll}
                type="button"
                className="px-2.5 py-1 rounded bg-gray-100 hover:bg-gray-200 text-gray-700 text-xs font-medium flex items-center gap-1 transition cursor-pointer"
                title="គណនាផលបូកឡើងវិញទាំងអស់"
              >
                <RefreshCw className="w-3 h-3" />
                <span>គណនាឡើងវិញ</span>
              </button>

              <button
                onClick={handleClearAll}
                type="button"
                className="px-2.5 py-1 rounded bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-200 text-xs font-medium flex items-center gap-1 transition cursor-pointer"
                title="កំណត់ទិន្នន័យជា ០ ទាំងអស់"
              >
                <RotateCcw className="w-3 h-3" />
                <span>កំណត់ឡើងវិញ (Reset)</span>
              </button>
            </div>
          </div>
        )}

        {/* Tacteing Selector Panel when toggled */}
        {showPictureToolbar && showTacteingSelector && (
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

      {/* Main Printable Landscape Canvas Area */}
      <div className="overflow-x-auto w-full flex justify-center print:overflow-visible print:block">
        <div
          ref={printContainerRef}
          id="yearly-k2-received-k1-print-area"
          style={{
            transform: `scale(${zoomLevel / 100})`,
            transformOrigin: 'top center',
            fontFamily: fontFamily || 'Khmer OS Siemreap',
            fontSize: `${fontSize}px`,
          }}
          className="bg-white text-black shadow-lg print:shadow-none border border-gray-300 print:border-none p-6 md:p-8 w-[1120px] shrink-0 print:w-full print:p-2 transition-transform duration-100"
        >
          {/* Header Layout (Ministry on left, Kingdom on right) */}
          <div className="flex justify-between items-start mb-3 leading-tight">
            {/* Left Header */}
            <div className="text-center text-[11px] flex flex-col items-center leading-tight pt-5">
              <p
                contentEditable
                suppressContentEditableWarning
                onBlur={(e) => setMinistryName(e.currentTarget.textContent || '')}
                style={{ fontSize: '12pt' }}
                className="font-moul text-[12pt] leading-tight outline-none hover:bg-amber-50/50 rounded px-1"
              >
                {ministryName}
              </p>
              <p
                contentEditable
                suppressContentEditableWarning
                onBlur={(e) => setDepartmentName(e.currentTarget.textContent || '')}
                style={{ fontSize: '12pt' }}
                className="font-moul text-[12pt] leading-tight outline-none hover:bg-amber-50/50 rounded px-1 mt-0.5"
              >
                {departmentName}
              </p>
              <p
                contentEditable
                suppressContentEditableWarning
                onBlur={(e) => setGeneralDeptName(e.currentTarget.textContent || '')}
                style={{ fontSize: '12pt' }}
                className="font-moul text-[12pt] leading-tight outline-none hover:bg-amber-50/50 rounded px-1 mt-0.5"
              >
                {generalDeptName}
              </p>
              <p
                contentEditable
                suppressContentEditableWarning
                onBlur={(e) => setOfficeName(e.currentTarget.textContent || '')}
                style={{ fontSize: '12pt' }}
                className="font-moul text-[12pt] leading-tight outline-none hover:bg-amber-50/50 rounded px-1 mt-0.5"
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

            {/* Right Header */}
            <div className="text-center text-[11px] flex flex-col items-center leading-tight">
              <p
                className="font-moul text-[12pt] leading-tight tracking-wider"
                style={{ fontSize: '12pt' }}
              >
                ព្រះរាជាណាចក្រកម្ពុជា
              </p>
              <p
                className="font-moul text-[12pt] leading-tight tracking-wider mt-1"
                style={{ fontSize: '12pt' }}
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
              onBlur={(e) => setReportTitle(e.currentTarget.textContent || 'តារាងសន្លឹកទិដ្ឋាការស្អិត តាមប្រភេទនីមួយៗ បើកពី ក១')}
              style={{
                fontSize: '12pt',
                fontFamily: "'Khmer Mool1', 'Khmer Mool', 'Khmer OS Muol Light', 'Khmer OS Muol', 'Moul', serif"
              }}
              className="text-[12pt] font-moul text-black font-normal leading-relaxed outline-none hover:bg-amber-50/50 rounded px-2"
            >
              {reportTitle}
            </h1>
            <p
              contentEditable
              suppressContentEditableWarning
              style={{
                fontSize: '12pt',
                fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif"
              }}
              className="text-[12pt] font-siemreap font-bold mt-1 text-black outline-none hover:bg-amber-50/50 rounded px-1"
            >
              ដោយគិតចាប់ពីថ្ងៃទី០១ ខែ{KHMER_MONTHS_NAMES[appliedStartMonth - 1]} ឆ្នាំ{toKhmerNum(appliedStartYear)} រហូតដល់ថ្ងៃទី{toKhmerNum(new Date(monthColumnsInfo[11].year, monthColumnsInfo[11].monthNum, 0).getDate())} ខែ{monthColumnsInfo[11].name} ឆ្នាំ{toKhmerNum(monthColumnsInfo[11].year)}
            </p>
          </div>

          {/* Main Distribution Table with Orange / Peach Theme Matching YearlyTeamDistributionReport */}
          <div className="w-full overflow-x-auto print:overflow-visible">
            <table
              className="w-full border-collapse border border-black text-black text-center text-[11px] leading-normal select-text font-siemreap"
              style={{
                '--table-row-height': `${tableRowHeight}px`,
                fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif"
              } as React.CSSProperties}
            >
              <thead
                className="font-siemreap"
                style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}
              >
                {/* Header Row 1 */}
                <tr
                  className="bg-[#FCE4D6] font-bold border-b border-black text-black h-[26px] font-siemreap"
                  style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}
                >
                  <th
                    rowSpan={2}
                    className="border border-black px-1 py-1 min-w-[36px] max-w-[40px] text-center align-middle font-bold font-siemreap text-[11px] whitespace-nowrap"
                    style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}
                  >
                    ប្រភេទ
                  </th>
                  <th
                    rowSpan={2}
                    className="border border-black px-1 py-1 min-w-[98px] text-center align-middle font-bold font-siemreap"
                    style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}
                  >
                    <div className="whitespace-nowrap font-bold text-[11px] leading-tight font-siemreap">សន្និធិមុនគ្រា</div>
                    <div className="text-[9px] font-semibold text-gray-800 mt-0.5 whitespace-nowrap leading-tight font-siemreap">
                      {openingStockLabel}
                    </div>
                  </th>
                  <th
                    colSpan={12}
                    className="border border-black px-1 py-1 text-center font-bold font-siemreap text-[11.5px] leading-normal whitespace-nowrap h-[24px] align-middle"
                    style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}
                  >
                    គិតចាប់ពីថ្ងៃទី០១ ខែ{KHMER_MONTHS_NAMES[appliedStartMonth - 1]} ឆ្នាំ{toKhmerNum(appliedStartYear)} រហូតដល់ថ្ងៃទី{toKhmerNum(new Date(monthColumnsInfo[11].year, monthColumnsInfo[11].monthNum, 0).getDate())} ខែ{monthColumnsInfo[11].name} ឆ្នាំ{toKhmerNum(monthColumnsInfo[11].year)}
                  </th>
                  <th
                    rowSpan={2}
                    className="border border-black px-1 py-1 min-w-[48px] max-w-[56px] text-center align-middle font-bold font-siemreap text-[10.5px] leading-tight"
                    style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}
                  >
                    <div className="whitespace-nowrap font-siemreap">សរុប</div>
                    <div className="whitespace-nowrap mt-0.5 font-siemreap">១២ខែ</div>
                  </th>
                  <th
                    rowSpan={2}
                    className="border border-black px-1 py-1 min-w-[50px] max-w-[58px] text-center align-middle font-bold font-siemreap text-[10.5px] leading-tight"
                    style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}
                  >
                    <div className="whitespace-nowrap font-siemreap">ចំនួន</div>
                    <div className="whitespace-nowrap mt-0.5 font-siemreap">សរុប</div>
                  </th>
                  <th
                    rowSpan={2}
                    onClick={() => setIsIssuedModalOpen(true)}
                    className="border border-black px-1 py-1 min-w-[50px] max-w-[58px] text-center align-middle font-bold font-siemreap text-[10.5px] leading-tight cursor-pointer hover:bg-amber-300 transition-colors select-none"
                    title="ចុចដើម្បីពិនិត្យផ្ទៀងផ្ទាត់បញ្ជីទិដ្ឋាការ បើកផ្តល់ជូនក្រុម (បើកផ្តល់តាមក្រុម - ផ្ទេរការប្រើប្រាស់)"
                    style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}
                  >
                    <div className="whitespace-nowrap font-siemreap">បើកផ្តល់</div>
                    <div className="whitespace-nowrap mt-0.5 font-siemreap">ជូនក្រុម</div>
                  </th>
                  <th
                    rowSpan={2}
                    onClick={() => {
                      setSelectedDamagedVisaType('ALL');
                      setIsDamagedModalOpen(true);
                    }}
                    className="border border-black px-1 py-1 min-w-[52px] max-w-[60px] text-center align-middle font-bold font-siemreap text-[10.5px] leading-tight cursor-pointer hover:bg-amber-300 transition-colors select-none"
                    title="ចុចដើម្បីពិនិត្យផ្ទៀងផ្ទាត់បញ្ជីទិដ្ឋាការ មិនបានការ & សាកល្បងក២ (ទិដ្ឋាការខូចក២ + ទិដ្ឋាការសាកក២)"
                    style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}
                  >
                    <div className="whitespace-nowrap font-siemreap">មិនបានការ</div>
                    <div className="text-[9px] whitespace-nowrap mt-0.5 leading-tight font-siemreap">សាកល្បងក២</div>
                  </th>
                  <th
                    rowSpan={2}
                    className="border border-black px-1 py-1 min-w-[54px] max-w-[62px] text-center align-middle font-bold font-siemreap text-[10.5px] leading-tight"
                    style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}
                  >
                    <div className="whitespace-nowrap font-siemreap">ទិដ្ឋាការ</div>
                    <div className="text-[9.5px] whitespace-nowrap mt-0.5 leading-tight font-siemreap">បង្វិលពីក្រុម</div>
                  </th>
                  <th
                    rowSpan={2}
                    className="border border-black px-1 py-1 min-w-[98px] text-center align-middle font-bold font-siemreap"
                    style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}
                  >
                    <div className="whitespace-nowrap font-bold text-[11px] leading-tight font-siemreap">សន្និធិនៅសល់</div>
                    <div className="text-[9px] font-semibold text-gray-800 mt-0.5 whitespace-nowrap leading-tight font-siemreap">
                      {endingStockLabel}
                    </div>
                  </th>
                </tr>

                {/* Header Row 2: 12 Months */}
                <tr
                  className="bg-[#FCE4D6] font-bold border-b border-black text-[10.5px] font-siemreap leading-normal text-black h-[20px]"
                  style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}
                >
                  {monthColumnsInfo.map((mCol, idx) => (
                    <th
                      key={idx}
                      onClick={() => {
                        setSelectedModalMonthIdx(idx);
                        setIsChecklistOpen(true);
                      }}
                      className="border border-black px-0.5 py-0.5 min-w-[44px] text-center align-middle hover:bg-amber-300 transition cursor-pointer select-none group whitespace-nowrap font-siemreap"
                      style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}
                      title={`ចុចដើម្បីពិនិត្យផ្ទៀងផ្ទាត់បញ្ជី និងចំនួនសរុបខែ ${mCol.name} (${mCol.year})`}
                    >
                      <span className="group-hover:underline">{mCol.name}</span>
                    </th>
                  ))}
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
                    issuedToTeams: 0,
                    damagedK2: 0,
                    returnStock: 0,
                    remaining: 0,
                  };

                  const dynamicMonthlySum = row.monthlyValues.reduce((a, b) => (Number(a) || 0) + (Number(b) || 0), 0);
                  const displayMonthlySum = autoCalculate ? dynamicMonthlySum : (row.totalMonthly ?? dynamicMonthlySum);
                  const displayAvailable = autoCalculate
                    ? (Number(row.openingStock) || 0) + displayMonthlySum
                    : (row.totalAvailable ?? ((Number(row.openingStock) || 0) + displayMonthlySum));
                  const displayRemaining = autoCalculate
                    ? displayAvailable - (Number(row.issuedToTeams) || 0) - (Number(row.damagedK2) || 0) + (Number(row.returnStock) || 0)
                    : (row.remaining ?? (displayAvailable - (Number(row.issuedToTeams) || 0) - (Number(row.damagedK2) || 0) + (Number(row.returnStock) || 0)));

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
                            className="border border-black px-0.5 py-[1px] text-right font-times outline-none hover:bg-yellow-100/70 focus:bg-yellow-100 focus:ring-1 focus:ring-amber-500 text-[10.5px] leading-tight text-black font-normal"
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
                        className={`border border-black px-1 py-[1px] text-right font-bold font-times bg-[#FDF5ED] text-[11px] leading-tight ${
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

                      {/* Issued to Teams (Editable) */}
                      <td
                        contentEditable
                        suppressContentEditableWarning
                        onBlur={(e) => handleCellChange(vt, 'issuedToTeams', e.currentTarget.textContent || '0')}
                        className="border border-black px-0.5 py-[1px] text-right font-normal font-times outline-none hover:bg-yellow-100/70 focus:bg-yellow-100 focus:ring-1 focus:ring-amber-500 text-[11px] leading-tight"
                      >
                        {renderCellValue(row.issuedToTeams)}
                      </td>

                      {/* Damaged K2 (Editable) */}
                      <td
                        contentEditable
                        suppressContentEditableWarning
                        onBlur={(e) => handleCellChange(vt, 'damagedK2', e.currentTarget.textContent || '0')}
                        className="border border-black px-0.5 py-[1px] text-right font-normal font-times text-black outline-none hover:bg-yellow-100/70 focus:bg-yellow-100 focus:ring-1 focus:ring-amber-500 text-[11px] leading-tight"
                      >
                        {renderCellValue(row.damagedK2)}
                      </td>

                      {/* Return / Transfer (Editable) */}
                      <td
                        contentEditable
                        suppressContentEditableWarning
                        onBlur={(e) => handleCellChange(vt, 'returnStock', e.currentTarget.textContent || '0')}
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
                    {renderCellValue(grandTotals.openingStock)}
                  </td>

                  {/* 12 Months Grand Totals */}
                  {grandTotals.monthlyValues.map((mTotal, mIdx) => (
                    <td
                      key={mIdx}
                      onClick={() => {
                        setSelectedModalMonthIdx(mIdx);
                        setIsChecklistOpen(true);
                      }}
                      className="border border-black px-0.5 py-[1px] text-right font-bold font-times text-[10.5px] hover:bg-amber-300 transition cursor-pointer leading-tight"
                      title={`ចុចពិនិត្យផ្ទៀងផ្ទាត់បញ្ជីខែ ${monthColumnsInfo[mIdx]?.name}`}
                    >
                      {renderCellValue(mTotal)}
                    </td>
                  ))}

                  {/* Monthly Sum Grand Total */}
                  <td className="border border-black px-1 py-[1px] text-right font-bold font-times bg-[#F8CBAD] text-[11px] leading-tight">
                    {renderCellValue(grandTotals.totalMonthly)}
                  </td>

                  {/* Available Grand Total */}
                  <td className="border border-black px-0.5 py-[1px] text-right font-bold font-times bg-[#F4B183]/60 text-[11px] leading-tight">
                    {renderCellValue(grandTotals.totalAvailable)}
                  </td>

                  {/* Issued Grand Total */}
                  <td
                    onClick={() => setIsIssuedModalOpen(true)}
                    className="border border-black px-0.5 py-[1px] text-right font-bold font-times cursor-pointer hover:bg-amber-300 transition text-[11px] leading-tight"
                    title="ចុចដើម្បីពិនិត្យផ្ទៀងផ្ទាត់បញ្ជីទិដ្ឋាការ បើកផ្តល់ជូនក្រុម (បើកផ្តល់តាមក្រុម - ផ្ទេរការប្រើប្រាស់)"
                  >
                    {renderCellValue(grandTotals.issuedToTeams)}
                  </td>

                  {/* Damaged Grand Total */}
                  <td
                    onClick={() => {
                      setSelectedDamagedVisaType('ALL');
                      setIsDamagedModalOpen(true);
                    }}
                    className="border border-black px-0.5 py-[2px] text-right font-bold font-times text-black cursor-pointer hover:bg-amber-300 transition text-[11px] leading-tight"
                    title="ចុចដើម្បីពិនិត្យផ្ទៀងផ្ទាត់បញ្ជីទិដ្ឋាការ មិនបានការ & សាកល្បងក២ ទាំងអស់"
                  >
                    {renderCellValue(grandTotals.damagedK2)}
                  </td>

                  {/* Return Grand Total */}
                  <td className="border border-black px-0.5 py-[2px] text-right font-bold font-times text-[11px] leading-tight">
                    {renderCellValue(grandTotals.returnStock)}
                  </td>

                  {/* Remaining Grand Total */}
                  <td className="border border-black px-1 py-[2px] text-right font-bold font-times bg-[#F4B183]/60 text-[11px] leading-tight">
                    {renderCellValue(grandTotals.remaining)}
                  </td>
                </tr>
              </tbody>
            </table>
          </div>

          {/* Footer Signature Section (Matching Picture & Layout) */}
          <div className="mt-3 flex justify-end">
            <div
              className="text-center min-w-[280px] leading-relaxed font-siemreap"
              style={{
                fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif",
                fontSize: '12pt',
              }}
            >
              {/* Khmer Lunar Date */}
              <p
                contentEditable
                suppressContentEditableWarning
                style={{
                  fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif",
                  fontSize: '12pt',
                }}
                className="outline-none hover:bg-amber-50/50 rounded px-1 text-slate-900 font-siemreap text-[12pt] leading-tight"
              >
                {signatureLunarDate}
              </p>

              {/* Solar Date */}
              <p
                contentEditable
                suppressContentEditableWarning
                style={{
                  fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif",
                  fontSize: '12pt',
                }}
                className="outline-none hover:bg-amber-50/50 rounded px-1 text-slate-900 font-medium mt-0.5 font-siemreap text-[12pt] leading-tight"
              >
                រាជធានីភ្នំពេញ, ថ្ងៃទី{dateNumberFormat === 'khmer' ? signatureSolarParts.khmerDay : signatureSolarParts.day} ខែ{signatureSolarParts.month} ឆ្នាំ{dateNumberFormat === 'khmer' ? signatureSolarParts.khmerYear : signatureSolarParts.year}
              </p>

              {/* Role Title in font-moul */}
              <p
                contentEditable
                suppressContentEditableWarning
                onBlur={(e) => setSignerRole(e.currentTarget.textContent || 'អ្នកធ្វើតារាង')}
                style={{
                  fontFamily: "'Khmer Mool1', 'Khmer Mool', 'Khmer OS Muol Light', 'Khmer OS Muol', 'Moul', serif",
                  fontSize: '12pt',
                }}
                className="font-moul text-[12pt] text-slate-900 font-normal outline-none hover:bg-amber-50/50 rounded px-1 mt-1.5 leading-tight"
              >
                {signerRole}
              </p>

              {/* Signature Graphic / Stamp Area (Clean comfortable space for physical signature & stamp) */}
              <div className="h-20 my-1" />

              {/* Officer Rank (Khmer OS Siemreap) & Name (Khmer OS Muol Light) in official Red font matching Picture */}
              <div
                style={{ color: nameColor || '#C00000', fontSize: '12pt' }}
                contentEditable
                suppressContentEditableWarning
                onBlur={(e) => setSignerName(e.currentTarget.textContent || '')}
                className="outline-none hover:bg-amber-50/50 rounded px-2 py-0.5 cursor-text text-center text-[12pt] leading-tight"
                title="ចុចដើម្បីកែប្រែឈ្មោះមន្ត្រី ឬជ្រើសរើសពីបញ្ជីខាងលើ"
              >
                {renderSignerFormatted(signerName)}
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Month Checklist Modal */}
      <YearlyMonthChecklistModal
        isOpen={isChecklistOpen}
        onClose={() => setIsChecklistOpen(false)}
        startYear={appliedStartYear}
        startMonth={appliedStartMonth}
        monthColumnsInfo={monthColumnsInfo}
        tableData={tableData as any}
        stockRecords={stockRecords}
        initialMonthIndex={selectedModalMonthIdx}
        mode="receive_k1"
        onApplyMonthActualValues={handleApplyMonthActualValues}
      />

      {/* Damaged & Missing / Test Checklist Modal */}
      <YearlyDamagedMissingChecklistModal
        isOpen={isDamagedModalOpen}
        onClose={() => setIsDamagedModalOpen(false)}
        startYear={appliedStartYear}
        startMonth={appliedStartMonth}
        monthColumnsInfo={monthColumnsInfo}
        tableData={tableData as any}
        stockRecords={stockRecords}
        initialVisaFilter={selectedDamagedVisaType}
        mode="receive_k1"
        onUpdateDamagedMissing={(vt, newVal) => {
          handleCellChange(vt as any, 'damagedK2', String(newVal));
        }}
      />

      {/* Issued to Teams Checklist Modal (បើកផ្តល់ជូនក្រុម = បើកផ្តល់តាមក្រុម - ផ្ទេរការប្រើប្រាស់) */}
      <YearlyReceivedK2ChecklistModal
        isOpen={isIssuedModalOpen}
        onClose={() => setIsIssuedModalOpen(false)}
        startYear={appliedStartYear}
        startMonth={appliedStartMonth}
        monthColumnsInfo={monthColumnsInfo}
        tableData={tableData as any}
        stockRecords={stockRecords}
        mode="receive_k1"
        onApplyAllIssuedValues={handleApplyAllIssuedValues}
      />
    </div>
  );
};
