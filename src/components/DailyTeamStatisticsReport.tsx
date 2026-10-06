import React, { useState, useEffect, useMemo, useRef, useCallback } from 'react';
import * as XLSX from 'xlsx';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import { CategoriesState, Officer, StockRecord, UserRole } from '../types';
import { sanitizeDocumentForHtml2Canvas, exportElementToPdf } from '../utils/pdfExportHelper';
import { printA4Document } from '../utils/printHelper';
import { useWorkspaceSettings } from '../context/WorkspaceSettingsContext';
import { CustomDatePicker } from './CustomDatePicker';
import {
  Printer,
  Download,
  Calendar,
  Save,
  RotateCcw,
  RefreshCw,
  FileSpreadsheet,
  ZoomIn,
  ZoomOut,
  ChevronDown,
  ChevronUp,
  Building2,
  Users,
  CheckCircle2,
  SlidersHorizontal,
  Eye,
  EyeOff,
  FileText,
  CheckSquare,
  Sparkles,
} from 'lucide-react';
import { getKhmerLunarDate, getKhmerSolarParts, toKhmerNum, parseDateInput } from '../utils/khmerCalendar';
import {
  formatReportTeamName,
  normalizeTeamName,
  normalizeDateToISO,
  OFFICIAL_29_TEAMS,
  VISA_TYPES,
} from '../utils/teamNormalization';
import { TacteingLine, TacteingControlSelector, getSavedTacteingSettings, saveTacteingSettings, TacteingType } from './TacteingLine';
import { DailyTeamStatsChecklistModal } from './DailyTeamStatsChecklistModal';

export interface DailyTeamStatisticsReportProps {
  mode?: 'office' | 'team';
  stockRecords?: StockRecord[];
  categories?: CategoriesState;
  officers?: Officer[];
  currentRole?: UserRole;
  userName?: string;
  assignedTeam?: string;
  onClose?: () => void;
  onShowToast?: (message: string, type?: 'success' | 'error' | 'info') => void;
  onAddStockRecord?: (record: any) => void;
  onBatchImportStockRecords?: (records: any[]) => void;
  onUpdateStockRecord?: (record: any) => void;
  onDeleteStockRecord?: (id: string) => void;
  onDeleteBatchStockRecords?: (ids: string[]) => void;
}

// 13 Official Visa Types
const STATS_VISA_TYPES = ['T', 'T1', 'T2', 'T3', 'E', 'E1', 'E2', 'E3', 'D', 'K', 'A', 'B', 'C'] as const;
type StatsVisaType = (typeof STATS_VISA_TYPES)[number];

// English short month names for 01-Jul-2026 format
const SHORT_ENG_MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

// Khmer month names
const KHMER_MONTHS = [
  'មករា', 'កុម្ភៈ', 'មីនា', 'មេសា', 'ឧសភា', 'មិថុនា',
  'កក្កដា', 'សីហា', 'កញ្ញា', 'តុលា', 'វិច្ឆិកា', 'ធ្នូ',
];

// Pure calculation helper to aggregate daily usage data for a team or all teams
export const calculateDailyUsageData = (
  team: string,
  year: number,
  month: number,
  numDays: number,
  stockRecords?: StockRecord[],
  customExcludedIds?: string[],
  categoryFilter: 'all' | 'Sticker' | 'cEA' = 'Sticker',
  customExcludedStorageKey?: string
): { data: Record<number, Record<StatsVisaType, number>>; foundAny: boolean } => {
  const isAllTeams = team === 'សរុបបណ្តាក្រុម' || !team;
  const normTeam = normalizeTeamName(team);

  // Load excluded IDs from argument and localStorage
  const excludedSet = new Set<string>(customExcludedIds || []);
  try {
    const key = customExcludedStorageKey || 'app_daily_team_excluded_ids';
    const savedEx = localStorage.getItem(key);
    if (savedEx) {
      const parsedEx = JSON.parse(savedEx);
      if (Array.isArray(parsedEx)) {
        parsedEx.forEach((id: string) => excludedSet.add(id));
      }
    }
  } catch {}

  const isTeamMatch = (rawTeamName?: string) => {
    if (isAllTeams) return true;
    if (!rawTeamName || !normTeam) return false;
    const recNorm = normalizeTeamName(rawTeamName);
    return recNorm === normTeam;
  };

  const newDaily: Record<number, Record<StatsVisaType, number>> = {};
  for (let d = 1; d <= numDays; d++) {
    newDaily[d] = {
      T: 0, T1: 0, T2: 0, T3: 0,
      E: 0, E1: 0, E2: 0, E3: 0,
      D: 0, K: 0, A: 0, B: 0, C: 0,
    };
  }

  let foundAny = false;

  // Gather actual stock records from Visa Sheet Stock Data (ទិន្នន័យសន្លឹកទិដ្ឋាការ)
  const allStock: StockRecord[] = [...(stockRecords || [])];
  try {
    const savedStock = localStorage.getItem('app_stock_records');
    if (savedStock) {
      const parsed = JSON.parse(savedStock);
      if (Array.isArray(parsed)) {
        parsed.forEach((r: StockRecord) => {
          if (!allStock.some((ex) => ex.id === r.id)) {
            allStock.push(r);
          }
        });
      }
    }
  } catch (e) {
    console.warn('Error reading app_stock_records', e);
  }

  // Process strictly real stock records of type useTeam (ការប្រើប្រាស់តាមក្រុម) from visa sheet data
  allStock.forEach((rec) => {
    if (!rec || !rec.date) return;
    if (rec.id && excludedSet.has(rec.id)) return;

    // Filter out sample, mock, or fake test stock records
    const isSampleId =
      rec.id &&
      (rec.id.startsWith('stk-use-shv-') ||
        rec.id.startsWith('stk-use-phnomden-') ||
        rec.id.startsWith('stk-use-tpp-') ||
        rec.id.startsWith('stk-use-sample-') ||
        rec.id.startsWith('sample-') ||
        rec.id.startsWith('fake-') ||
        rec.id === 'stk-bavet-t' ||
        rec.id === 'stk-bavet-e');
    if (isSampleId) return;

    const op = (rec.operationType || '').toLowerCase();
    const sf = (rec.sourceFrom || '').toLowerCase();
    const isUseTeam =
      (op === 'useteam' ||
        op === 'use_team' ||
        (op.includes('ប្រើប្រាស់') &&
          !op.includes('ផ្ទេរ') &&
          !sf.includes('ផ្ទេរ') &&
          !op.includes('បើក') &&
          !op.includes('ចែក') &&
          !op.includes('បង្វិល') &&
          !op.includes('ខូច') &&
          !op.includes('បាត់') &&
          !op.includes('គល់'))) &&
      !op.includes('stub') &&
      !sf.includes('គល់សន្លឹក');
    if (!isUseTeam) return;

    const isCea =
      rec.sourceFrom === 'cEA' ||
      rec.remarks?.includes('cEA') ||
      (rec.visaType && rec.visaType.includes('cEA')) ||
      (rec.stockType && rec.stockType.toLowerCase().includes('evisa'));

    if (categoryFilter === 'Sticker' && isCea) return;
    if (categoryFilter === 'cEA' && !isCea) return;

    const rawTeam = rec.visaTeamRobokName || (rec as any).teamName || '';
    if (!isTeamMatch(rawTeam)) return;

    const isoDate = normalizeDateToISO(rec.date);
    const parts = isoDate.split('-');
    if (parts.length >= 3) {
      const recYear = parseInt(parts[0], 10);
      const recMonth = parseInt(parts[1], 10);
      const recDay = parseInt(parts[2], 10);

      if (recYear === year && recMonth === month && recDay >= 1 && recDay <= numDays) {
        const vtRaw = (rec.visaType || '').toUpperCase().trim();
        const matchedVt = STATS_VISA_TYPES.find((v) => v === vtRaw);
        const qty =
          rec.totalSheets !== undefined && rec.totalSheets !== null
            ? Number(rec.totalSheets)
            : Number(rec.quantityBundles || 0);

        // Strictly count only when visa type is valid in STATS_VISA_TYPES and quantity > 0
        // Never fabricate, guess, or assign to fallback types ('T' or 'E')
        if (qty > 0 && matchedVt) {
          foundAny = true;
          newDaily[recDay][matchedVt] = (newDaily[recDay][matchedVt] || 0) + qty;
        }
      }
    }
  });

  return { data: newDaily, foundAny };
};

export const DailyTeamStatisticsReport: React.FC<DailyTeamStatisticsReportProps> = ({
  mode,
  stockRecords = [],
  categories,
  officers = [],
  currentRole = 'Secondary',
  userName = '',
  assignedTeam,
  onClose,
}) => {
  const { scaleMode } = useWorkspaceSettings();
  const isOffice = mode ? mode === 'office' : (currentRole === 'Secondary' || currentRole === 'Admin');
  const isTeam = !isOffice;
  const isSecondary = isOffice;
  const storagePrefix = isOffice ? 'office_daily_stats_' : 'team_daily_stats_';
  const excludedStorageKey = isOffice ? 'office_daily_team_excluded_ids' : 'team_daily_team_excluded_ids';

  const getStoredNumber = (key: string, defaultVal: number): number => {
    try {
      const v = localStorage.getItem(storagePrefix + key);
      if (v !== null) {
        const num = parseFloat(v);
        if (!isNaN(num)) return num;
      }
      const legacy = localStorage.getItem('daily_team_stats_' + key);
      if (legacy !== null) {
        const num = parseFloat(legacy);
        if (!isNaN(num)) return num;
      }
    } catch {}
    return defaultVal;
  };

  const setStoredSetting = (key: string, val: string | number) => {
    try {
      localStorage.setItem(storagePrefix + key, String(val));
    } catch {}
  };

  const getStoredBoolean = (key: string, defaultVal: boolean): boolean => {
    try {
      const v = localStorage.getItem(storagePrefix + key);
      if (v !== null) return v === 'true';
      const legacy = localStorage.getItem('daily_team_stats_' + key);
      if (legacy !== null) return legacy === 'true';
    } catch {}
    return defaultVal;
  };

  const reportContainerRef = useRef<HTMLDivElement>(null);

  // Category Option Selector: Sticker vs cEA
  const optionStorageKey = `${storagePrefix}selected_option`;
  const [showControlPanel, setShowControlPanel] = useState<boolean>(() =>
    getStoredBoolean('show_control_panel', true)
  );
  const [selectedOption, setSelectedOption] = useState<'Sticker' | 'cEA'>(() => {
    try {
      const saved = localStorage.getItem(optionStorageKey);
      if (saved === 'cEA' || saved === 'Sticker') return saved;
      const legacy = localStorage.getItem('app_daily_stats_selected_option');
      if (legacy === 'cEA' || legacy === 'Sticker') return legacy;
    } catch {}
    return 'Sticker';
  });

  const handleSelectOption = (opt: 'Sticker' | 'cEA') => {
    setSelectedOption(opt);
    try {
      localStorage.setItem(optionStorageKey, opt);
    } catch {}
  };

  // Team options from CategoryManager (team_type_options_v1)
  const [teamTypeOptions, setTeamTypeOptions] = useState<Record<string | number, { sticker?: boolean; evisa?: boolean }>>(() => {
    try {
      const saved = localStorage.getItem('team_type_options_v1');
      return saved ? JSON.parse(saved) : {};
    } catch {
      return {};
    }
  });

  useEffect(() => {
    const handleUpdate = () => {
      try {
        const saved = localStorage.getItem('team_type_options_v1');
        setTeamTypeOptions(saved ? JSON.parse(saved) : {});
      } catch (e) {
        console.error(e);
      }
    };
    window.addEventListener('team_type_options_updated', handleUpdate);
    window.addEventListener('storage', handleUpdate);
    return () => {
      window.removeEventListener('team_type_options_updated', handleUpdate);
      window.removeEventListener('storage', handleUpdate);
    };
  }, []);

  // Helper to check if a team is configured for Sticker or cEA
  const isTeamAllowedCategory = useCallback(
    (teamName: string, category: 'Sticker' | 'cEA', idx?: number) => {
      if (!teamName) return false;
      const cleanTeam = teamName.trim().toLowerCase();
      const normTeam = normalizeTeamName(teamName);

      // 1. Check direct name or normalized name in teamTypeOptions
      for (const [key, val] of Object.entries(
        teamTypeOptions as Record<string, { sticker?: boolean; evisa?: boolean }>
      )) {
        if (
          String(key).trim().toLowerCase() === cleanTeam ||
          (normTeam && normalizeTeamName(String(key)) === normTeam)
        ) {
          if (category === 'Sticker') {
            return val?.sticker !== false;
          } else {
            return Boolean(val?.evisa);
          }
        }
      }

      // 2. Check in categories.visaTeamsRobok / visaTeams by index or ID
      const vtrList = categories?.visaTeamsRobok || [];
      let resolvedIdx = idx !== undefined && idx >= 0 ? idx : -1;
      if (resolvedIdx === -1) {
        resolvedIdx = vtrList.findIndex(
          (t) =>
            t.name.trim().toLowerCase() === cleanTeam ||
            (normTeam && normalizeTeamName(t.name) === normTeam)
        );
      }
      if (resolvedIdx !== -1) {
        const vtrItem = vtrList[resolvedIdx];
        const vtItem = categories?.visaTeams?.[resolvedIdx];
        const opt =
          (vtrItem?.id && teamTypeOptions[vtrItem.id]) ||
          (vtrItem?.name && teamTypeOptions[vtrItem.name]) ||
          (vtItem?.id && teamTypeOptions[vtItem.id]) ||
          (vtItem?.name && teamTypeOptions[vtItem.name]) ||
          teamTypeOptions[resolvedIdx];

        if (opt !== undefined) {
          if (category === 'Sticker') {
            return opt.sticker !== false;
          } else {
            return Boolean(opt.evisa);
          }
        }
      }

      // 3. Check if team has recorded operations in app_daily_team_operations_v5 or stockRecords
      if (category === 'cEA') {
        try {
          const savedOps = localStorage.getItem('app_daily_team_operations_v5');
          if (savedOps) {
            const ops = JSON.parse(savedOps);
            if (Array.isArray(ops)) {
              const hasCea = ops.some(
                (o: any) =>
                  o.categoryType === 'cEA' && normalizeTeamName(o.teamName) === normTeam
              );
              if (hasCea) return true;
            }
          }
        } catch {}
      }

      // 4. Fallback defaults: Key airports & ports default to cEA
      const defaultCeaTeams = [
        'អាកាស តេជោ',
        'អាកាស សៀមរាប',
        'អាកាស ព្រះសីហនុ',
        'កំពង់ផែ ព្រះសីហនុ',
      ];
      const isDefaultCea = defaultCeaTeams.some((k) => normalizeTeamName(k) === normTeam);

      if (category === 'cEA') {
        return isDefaultCea;
      }

      // All teams support Sticker by default unless explicitly disabled
      return true;
    },
    [teamTypeOptions, categories]
  );

  // Raw base teams from categories or official list
  const baseTeams = useMemo(() => {
    let list: { name: string; idx: number }[] = [];
    if (categories?.visaTeamsRobok && categories.visaTeamsRobok.length > 0) {
      list = categories.visaTeamsRobok
        .map((t, idx) => ({ name: t.name.trim(), idx }))
        .filter((t) => Boolean(t.name));
    } else {
      list = Array.from(OFFICIAL_29_TEAMS).map((name, idx) => ({ name, idx }));
    }
    return list;
  }, [categories?.visaTeamsRobok]);

  // Teams that use Sticker only/allowed
  const stickerTeams = useMemo(() => {
    return baseTeams
      .filter((t) => isTeamAllowedCategory(t.name, 'Sticker', t.idx))
      .map((t) => t.name);
  }, [baseTeams, isTeamAllowedCategory]);

  // Teams that use cEA allowed
  const ceaTeams = useMemo(() => {
    return baseTeams
      .filter((t) => isTeamAllowedCategory(t.name, 'cEA', t.idx))
      .map((t) => t.name);
  }, [baseTeams, isTeamAllowedCategory]);

  // Filtered teams list based on selectedOption (Sticker vs cEA)
  const allTeams = useMemo(() => {
    const list = selectedOption === 'cEA' ? ceaTeams : stickerTeams;
    if (isOffice) {
      return ['សរុបបណ្តាក្រុម', ...list];
    }
    return list;
  }, [selectedOption, ceaTeams, stickerTeams, isOffice]);

  // Check if current assigned team is allowed cEA (for User role)
  const isTeamEVisaAllowed = useMemo(() => {
    if (isOffice) return true;
    const targetTeam = assignedTeam;
    if (!targetTeam) return false;
    return isTeamAllowedCategory(targetTeam, 'cEA');
  }, [isOffice, assignedTeam, isTeamAllowedCategory]);

  // Selected Team (defaults to 'សរុបបណ្តាក្រុម' for Office / assignedTeam or 'ព្រំដែន ភ្នំដិន' for Team)
  const [selectedTeam, setSelectedTeam] = useState<string>(() => {
    if (isOffice) {
      return 'សរុបបណ្តាក្រុម';
    }
    return assignedTeam || 'ព្រំដែន ភ្នំដិន';
  });

  // Automatically adjust selectedTeam if current selection is not valid in the filtered teams list
  useEffect(() => {
    if (allTeams.length > 0 && !allTeams.includes(selectedTeam)) {
      if (selectedTeam !== 'សរុបបណ្តាក្រុម') {
        const firstSpecific = allTeams.find((t) => t !== 'សរុបបណ្តាក្រុម');
        setSelectedTeam(firstSpecific || allTeams[0]);
      } else {
        setSelectedTeam(allTeams[0]);
      }
    }
  }, [allTeams, selectedTeam]);

  // Check if viewing the combined office total report vs individual team table
  const isTotalReport = isOffice && selectedTeam === 'សរុបបណ្តាក្រុម';

  // Keep in sync if assignedTeam prop changes from parent
  useEffect(() => {
    if (isTeam && assignedTeam) {
      setSelectedTeam(assignedTeam);
    }
  }, [isTeam, assignedTeam]);

  // Selected Year & Month (Default to July 2026 matching the PDF)
  const [selectedYear, setSelectedYear] = useState<number>(2026);
  const [selectedMonth, setSelectedMonth] = useState<number>(7); // 7 = July (1-indexed)

  // Zoom scale for preview / print
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

  // Number of days in selected month and year
  const daysInMonth = useMemo(() => {
    return new Date(selectedYear, selectedMonth, 0).getDate();
  }, [selectedYear, selectedMonth]);

  // Daily values state: day (1..31) -> { T: 0, T1: 0, ..., C: 0 }
  // Starts with 0 so all cells display (-) initially until user clicks to load data
  const [dailyData, setDailyData] = useState<Record<number, Record<StatsVisaType, number>>>(() => {
    const initial: Record<number, Record<StatsVisaType, number>> = {};
    for (let d = 1; d <= 31; d++) {
      initial[d] = {
        T: 0, T1: 0, T2: 0, T3: 0,
        E: 0, E1: 0, E2: 0, E3: 0,
        D: 0, K: 0, A: 0, B: 0, C: 0,
      };
    }
    return initial;
  });

  // Checklist Audit Modal & Excluded IDs
  const [isChecklistOpen, setIsChecklistOpen] = useState<boolean>(false);
  const [excludedIds, setExcludedIds] = useState<string[]>(() => {
    try {
      const saved = localStorage.getItem(excludedStorageKey);
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed)) return parsed;
      }
      const legacy = localStorage.getItem('app_daily_team_excluded_ids');
      if (legacy) {
        const parsed = JSON.parse(legacy);
        if (Array.isArray(parsed)) return parsed;
      }
    } catch {}
    return [];
  });

  // Storage key for saving user edits per category-team-year-month
  const storageKey = useMemo(() => {
    const norm = normalizeTeamName(selectedTeam) || 'default';
    return `${storagePrefix}data_${selectedOption}_${norm}_${selectedYear}_${selectedMonth}`;
  }, [storagePrefix, selectedOption, selectedTeam, selectedYear, selectedMonth]);

  const legacyStorageKey = useMemo(() => {
    const norm = normalizeTeamName(selectedTeam) || 'default';
    return `daily_team_stats_data_${norm}_${selectedYear}_${selectedMonth}`;
  }, [selectedTeam, selectedYear, selectedMonth]);

  // Clean up any stale or ghost cached daily stats so only real visa data is shown
  useEffect(() => {
    try {
      const keysToRemove: string[] = [];
      for (let i = 0; i < localStorage.length; i++) {
        const k = localStorage.key(i);
        if (k && (k.startsWith('daily_team_stats_data_') || k.startsWith('daily_team_stats_fake_'))) {
          keysToRemove.push(k);
        }
      }
      keysToRemove.forEach((k) => localStorage.removeItem(k));
    } catch {}
  }, []);

  // Keep table synced: strictly populate real usage results from actual recorded visa sheet stock operations
  useEffect(() => {
    // Automatically compute and display usage results from real visa stock operations for the selected team & category
    const { data: autoData } = calculateDailyUsageData(
      selectedTeam,
      selectedYear,
      selectedMonth,
      daysInMonth,
      stockRecords,
      excludedIds,
      selectedOption,
      excludedStorageKey
    );
    setDailyData(autoData);
  }, [
    selectedOption,
    selectedTeam,
    selectedYear,
    selectedMonth,
    daysInMonth,
    stockRecords,
    excludedIds,
    excludedStorageKey,
  ]);

  // Tacteing ornament style
  const [tacteingSettings, setTacteingSettings] = useState<{ type: TacteingType; customImage: string | null }>(
    () => getSavedTacteingSettings()
  );

  // Tacteing ornament picture visibility states
  const [showTacteingPicture, setShowTacteingPicture] = useState<boolean>(() =>
    getStoredBoolean('show_tacteing_picture', true)
  );
  const [showLeftTacteing, setShowLeftTacteing] = useState<boolean>(() =>
    getStoredBoolean('show_left_tacteing', true)
  );
  const [showRightTacteing, setShowRightTacteing] = useState<boolean>(() =>
    getStoredBoolean('show_right_tacteing', true)
  );
  const [showTacteingDropdown, setShowTacteingDropdown] = useState<boolean>(false);

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

  // Header and Footer custom text states
  const [ministryName, setMinistryName] = useState<string>('ក្រសួងមហាផ្ទៃ');
  const [departmentName, setDepartmentName] = useState<string>('អគ្គនាយកដ្ឋានអន្តោប្រវេសន៍');
  const [generalDeptName, setGeneralDeptName] = useState<string>('នាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត');
  const [officeName, setOfficeName] = useState<string>('ការិយាល័យទិដ្ឋាការចូល');
  const [customTeamTitle, setCustomTeamTitle] = useState<string>(() => (isTotalReport ? 'ផ្នែករដ្ឋបាល' : ''));
  const [teamTitleFont, setTeamTitleFont] = useState<'siemreap' | 'moul'>('siemreap');
  const [rightHeaderMargin, setRightHeaderMargin] = useState<number>(0);
  const [signingDate, setSigningDate] = useState<string>('2026-08-01');
  const [signingLocation, setSigningLocation] = useState<string>('ភ្នំពេញ');
  const [customLunarDate, setCustomLunarDate] = useState<string>(
    'ថ្ងៃសៅរ៍ ៣រោច ខែទុតិយាសាឍ ឆ្នាំមមី អដ្ឋស័ក ព.ស២៥៧០'
  );
  const [chiefTitle, setChiefTitle] = useState<string>(() => (isOffice ? 'នាយការិយាល័យ' : 'ប្រធានក្រុមផ្តល់ទិដ្ឋាការ'));
  const [chiefTitleFont, setChiefTitleFont] = useState<'muol' | 'siemreap'>('muol');
  const [chiefName, setChiefName] = useState<string>('');
  const [signatureMarginTop, setSignatureMarginTop] = useState<number>(() => getStoredNumber('signature_mt', 0));

  // Page Margins in centimeters (Default left: 2.80cm, right: 1.30cm, top: 0.50cm, bottom: 0.50cm)
  const [pageLeftMargin, setPageLeftMargin] = useState<number>(() => getStoredNumber('left_margin', 2.80));
  const [pageRightMargin, setPageRightMargin] = useState<number>(() => getStoredNumber('right_margin', 1.30));
  const [pageTopMargin, setPageTopMargin] = useState<number>(() => getStoredNumber('top_margin', 0.5));
  const [pageBottomMargin, setPageBottomMargin] = useState<number>(() => getStoredNumber('bottom_margin', 0.5));
  const [showHeaderSettings, setShowHeaderSettings] = useState<boolean>(false);
  const [headerFontSizePt, setHeaderFontSizePt] = useState<number>(() => getStoredNumber('header_font_size', 12));
  const [teamTitleFontSizePt, setTeamTitleFontSizePt] = useState<number>(() => getStoredNumber('team_title_font_size', 10));
  // Font size for table results & cells
  const [tableFontSize, setTableFontSize] = useState<number>(() => getStoredNumber('table_font_size', 12.5));
  const [tableRowHeight, setTableRowHeight] = useState<number>(() => getStoredNumber('row_height', 19));
  const [isSavedNotice, setIsSavedNotice] = useState<boolean>(false);
  const [isExporting, setIsExporting] = useState<boolean>(false);

  // Auto-fill header team title and signature chief title when team changes
  useEffect(() => {
    if (isOffice && selectedTeam === 'សរុបបណ្តាក្រុម') {
      setCustomTeamTitle('ផ្នែករដ្ឋបាល');
      setChiefTitle('នាយការិយាល័យ');
      return;
    }

    if (isOffice) {
      setChiefTitle('នាយការិយាល័យ');
    } else {
      setChiefTitle('ប្រធានក្រុមផ្តល់ទិដ្ឋាការ');
    }

    let full = '';
    if (categories) {
      const robokList = categories.visaTeamsRobok || [];
      const fullList = categories.visaTeams || [];
      const norm = normalizeTeamName(selectedTeam);
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
    setCustomTeamTitle(full);
  }, [selectedTeam, categories, isOffice, isSecondary]);

  // Auto-generate lunar date when signing date changes (or manual)
  useEffect(() => {
    if (signingDate) {
      if (signingDate === '2026-08-01') {
        setCustomLunarDate('ថ្ងៃសៅរ៍ ៣រោច ខែទុតិយាសាឍ ឆ្នាំមមី អដ្ឋស័ក ព.ស២៥៧០');
        return;
      }
      const calcLunar = getKhmerLunarDate(signingDate);
      if (calcLunar) {
        setCustomLunarDate(calcLunar);
      }
    }
  }, [signingDate]);

  // Computed statistics of actual stock records of type 'useTeam' for the active team (or all teams)
  const teamUsageStats = useMemo(() => {
    const isAllTeams = isSecondary || selectedTeam === 'សរុបបណ្តាក្រុម';
    const normTeam = normalizeTeamName(selectedTeam);
    const allStock: StockRecord[] = [...(stockRecords || [])];
    try {
      const saved = localStorage.getItem('app_stock_records');
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed)) {
          parsed.forEach((r: StockRecord) => {
            if (!allStock.some((ex) => ex.id === r.id)) allStock.push(r);
          });
        }
      }
    } catch (e) {
      console.warn('Error reading app_stock_records', e);
    }

    const isTeamMatch = (rawTeamName?: string) => {
      if (isAllTeams) return true;
      if (!rawTeamName || !normTeam) return false;
      const recNorm = normalizeTeamName(rawTeamName);
      return recNorm === normTeam;
    };

    const teamOps = allStock.filter((rec) => {
      if (!rec || !rec.date) return false;
      if (rec.id && excludedIds.includes(rec.id)) return false;

      const isSampleId =
        rec.id &&
        (rec.id.startsWith('stk-use-shv-') ||
          rec.id.startsWith('stk-use-phnomden-') ||
          rec.id.startsWith('stk-use-tpp-') ||
          rec.id.startsWith('stk-use-sample-') ||
          rec.id.startsWith('sample-') ||
          rec.id.startsWith('fake-') ||
          rec.id === 'stk-bavet-t' ||
          rec.id === 'stk-bavet-e');
      if (isSampleId) return false;

      const op = (rec.operationType || '').toLowerCase();
      const isUseTeam =
        op === 'useteam' ||
        op.includes('useteam') ||
        (op.includes('ប្រើប្រាស់') &&
          !op.includes('បើក') &&
          !op.includes('ចែក') &&
          !op.includes('បង្វិល') &&
          !op.includes('ខូច') &&
          !op.includes('បាត់'));
      if (!isUseTeam) return false;

      const isCea =
        rec.sourceFrom === 'cEA' ||
        rec.remarks?.includes('cEA') ||
        (rec.visaType && rec.visaType.includes('cEA')) ||
        (rec.stockType && rec.stockType.toLowerCase().includes('evisa'));

      if (selectedOption === 'Sticker' && isCea) return false;
      if (selectedOption === 'cEA' && !isCea) return false;

      const rawTeam = rec.visaTeamRobokName || (rec as any).teamName || '';
      return isTeamMatch(rawTeam);
    });

    const monthOps = teamOps.filter((rec) => {
      const parts = rec.date.split('-');
      if (parts.length >= 2) {
        const y = parseInt(parts[0], 10);
        const m = parseInt(parts[1], 10);
        return y === selectedYear && m === selectedMonth;
      }
      return false;
    });

    const totalSheetsInMonth = monthOps.reduce((sum, r) => {
      const vtRaw = (r.visaType || '').toUpperCase().trim();
      const matchedVt = STATS_VISA_TYPES.find((v) => v === vtRaw);
      if (!matchedVt) return sum;
      const qty =
        r.totalSheets !== undefined && r.totalSheets !== null
          ? Number(r.totalSheets)
          : Number(r.quantityBundles || 0);
      return sum + qty;
    }, 0);
    const monthOpsCount = monthOps.length;

    return {
      allCount: teamOps.length,
      monthCount: monthOpsCount,
      totalSheetsInMonth,
    };
  }, [selectedTeam, selectedYear, selectedMonth, stockRecords, isSecondary, excludedIds, selectedOption]);

  // Pull actual data from stockRecords (operationType === 'useTeam' / 'ការប្រើប្រាស់តាមក្រុម') AND DailyTeamOperations
  const pullFromTeamUsage = () => {
    const { data: calculated } = calculateDailyUsageData(
      selectedTeam,
      selectedYear,
      selectedMonth,
      daysInMonth,
      stockRecords,
      excludedIds,
      selectedOption,
      excludedStorageKey
    );
    setDailyData(calculated);
    setIsSavedNotice(true);
    setTimeout(() => setIsSavedNotice(false), 3000);
  };

  // Handler when checklist updates excluded IDs
  const handleApplyExcludedIds = (newIds: string[]) => {
    setExcludedIds(newIds);
    try {
      localStorage.setItem(excludedStorageKey, JSON.stringify(newIds));
    } catch (e) {
      console.warn('Failed to save excluded IDs to localStorage', e);
    }
    const { data: calculated } = calculateDailyUsageData(
      selectedTeam,
      selectedYear,
      selectedMonth,
      daysInMonth,
      stockRecords,
      newIds,
      selectedOption,
      excludedStorageKey
    );
    setDailyData(calculated);
    setIsSavedNotice(true);
    setTimeout(() => setIsSavedNotice(false), 3000);
  };

  // Initialize empty month
  const initEmptyMonth = () => {
    const empty: Record<number, Record<StatsVisaType, number>> = {};
    for (let d = 1; d <= daysInMonth; d++) {
      empty[d] = {
        T: 0, T1: 0, T2: 0, T3: 0,
        E: 0, E1: 0, E2: 0, E3: 0,
        D: 0, K: 0, A: 0, B: 0, C: 0,
      };
    }
    setDailyData(empty);
  };

  // Save manual modifications to localStorage
  const handleSave = () => {
    try {
      localStorage.setItem(storageKey, JSON.stringify(dailyData));
      setIsSavedNotice(true);
      setTimeout(() => setIsSavedNotice(false), 2500);
    } catch (e) {
      alert('មិនអាចរក្សាទុកបានទេ');
    }
  };

  // Handle cell edit
  const handleCellChange = (day: number, visaType: StatsVisaType, valStr: string) => {
    const cleaned = valStr.replace(/[^\d]/g, '');
    const num = parseInt(cleaned, 10) || 0;
    setDailyData((prev) => ({
      ...prev,
      [day]: {
        ...(prev[day] || {}),
        [visaType]: num,
      },
    }));
  };

  // Computed row totals and column grand totals
  const { rowTotals, columnTotals, grandTotal } = useMemo(() => {
    const rTotals: Record<number, number> = {};
    const cTotals: Record<StatsVisaType, number> = {
      T: 0, T1: 0, T2: 0, T3: 0,
      E: 0, E1: 0, E2: 0, E3: 0,
      D: 0, K: 0, A: 0, B: 0, C: 0,
    };
    let gTotal = 0;

    for (let d = 1; d <= daysInMonth; d++) {
      let rowSum = 0;
      const dayValues = dailyData[d] || {};
      STATS_VISA_TYPES.forEach((vt) => {
        const val = dayValues[vt] || 0;
        rowSum += val;
        cTotals[vt] += val;
      });
      rTotals[d] = rowSum;
      gTotal += rowSum;
    }

    return { rowTotals: rTotals, columnTotals: cTotals, grandTotal: gTotal };
  }, [dailyData, daysInMonth]);

  // Formatted date string for row, e.g. "01-Jul-2026"
  const getRowDateLabel = (day: number) => {
    const dStr = String(day).padStart(2, '0');
    const mStr = SHORT_ENG_MONTHS[selectedMonth - 1] || 'Jul';
    return `${dStr}-${mStr}-${selectedYear}`;
  };

  // Subtitle date range in Khmer digits
  const khmerDateRangeSubtitle = useMemo(() => {
    const mName = KHMER_MONTHS[selectedMonth - 1] || 'កក្កដា';
    const yKh = toKhmerNum(selectedYear);
    const lastDayKh = toKhmerNum(String(daysInMonth).padStart(2, '0'));
    return `គិតពីថ្ងៃទី០១ ខែ${mName} ឆ្នាំ${yKh} ដល់ ថ្ងៃទី${lastDayKh} ខែ${mName} ឆ្នាំ${yKh}`;
  }, [selectedMonth, selectedYear, daysInMonth]);

  // Formatted signing date in Khmer
  const formattedSigningDateKhmer = useMemo(() => {
    if (!signingDate) return '';
    const parts = getKhmerSolarParts(signingDate);
    return `${signingLocation || 'ភ្នំពេញ'}, ថ្ងៃទី${parts.khmerDay} ខែ${parts.khmerMonth} ឆ្នាំ${parts.khmerYear}`;
  }, [signingDate, signingLocation]);

  // Print function
  const handlePrint = () => {
    const docPrefix = isOffice ? 'ស្ថិតិសរុបប្រចាំថ្ងៃរបស់ក្រុម' : 'តារាងស្ថិតិប្រចាំថ្ងៃរបស់ក្រុម';
    printA4Document(reportContainerRef.current || 'daily-team-statistics-report-pdf', {
      documentTitle: `${docPrefix}_${(selectedTeam || 'team').replace(/[\/\s]+/g, '_')}_${selectedYear}_M${selectedMonth}`,
      orientation: 'portrait',
      pageMargin: `${pageTopMargin}cm ${pageRightMargin}cm ${pageBottomMargin}cm ${pageLeftMargin}cm`,
    });
  };

  // Export to PDF (matching Robok Total Stock Work Report)
  const handleExportPDF = async () => {
    if (!reportContainerRef.current) return;
    setIsExporting(true);
    const safeTeam = (selectedTeam || 'team').replace(/[\/\s]+/g, '_');
    const filePrefix = isOffice ? 'ស្ថិតិសរុបប្រចាំថ្ងៃរបស់ក្រុម' : 'តារាងស្ថិតិប្រចាំថ្ងៃរបស់ក្រុម';
    const fileName = `${filePrefix}_${safeTeam}_${selectedYear}_M${selectedMonth}.pdf`;

    // Temporarily reset zoom scale to 100% for capture to ensure sharp render
    const currentZoom = zoomScale;
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
      console.error('html-to-image PDF export error, attempting fallback:', e);
      try {
        if (document.fonts && document.fonts.ready) {
          await document.fonts.ready;
        }
        await new Promise((r) => setTimeout(r, 200));

        const element = reportContainerRef.current;
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
            sanitizeDocumentForHtml2Canvas(clonedDoc, 'daily-team-statistics-report-pdf');
            const clonedElement = clonedDoc.getElementById('daily-team-statistics-report-pdf');
            if (clonedElement) {
              clonedElement.style.width = '210mm';
              clonedElement.style.minWidth = '210mm';
              clonedElement.style.maxWidth = '210mm';
              clonedElement.style.paddingLeft = `${pageLeftMargin}cm`;
              clonedElement.style.paddingRight = `${pageRightMargin}cm`;
              clonedElement.style.paddingTop = `${pageTopMargin}cm`;
              clonedElement.style.paddingBottom = `${pageBottomMargin}cm`;
              clonedElement.style.boxSizing = 'border-box';
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

        pdf.addImage(imgData, 'PNG', 0, 0, pageWidth, Math.min(pageHeight, calculatedHeight), undefined, 'FAST');
        pdf.save(fileName);
      } catch (fallbackErr) {
        console.error('Export PDF fallback failed:', fallbackErr);
      }
    } finally {
      setZoomScale(currentZoom);
      setIsExporting(false);
    }
  };

  // Export to Excel (XLSX)
  const handleExportExcel = () => {
    const rows: any[] = [];

    // Title rows
    rows.push(['ព្រះរាជាណាចក្រកម្ពុជា']);
    rows.push(['ជាតិ សាសនា ព្រះមហាក្សត្រ']);
    rows.push(['']);
    rows.push(['ក្រសួងមហាផ្ទៃ']);
    rows.push(['អគ្គនាយកដ្ឋានអន្តោប្រវេសន៍']);
    rows.push(['នាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត']);
    rows.push(['ការិយាល័យទិដ្ឋាការចូល']);
    const subOfficeTitle = isTotalReport
      ? (customTeamTitle || 'ផ្នែករដ្ឋបាល')
      : (customTeamTitle || formatReportTeamName(selectedTeam));
    rows.push([subOfficeTitle]);
    rows.push(['']);
    rows.push(['ស្ថិតិភ្ញៀវស្នើសុំទិដ្ឋាការចូលប្រទេស']);
    rows.push([khmerDateRangeSubtitle]);
    rows.push(['']);

    // Table Header
    const headerRow = ['ប្រចាំថ្ងៃ', ...STATS_VISA_TYPES, 'សរុប'];
    rows.push(headerRow);

    // Data rows
    for (let d = 1; d <= daysInMonth; d++) {
      const dateLabel = getRowDateLabel(d);
      const rowVals = STATS_VISA_TYPES.map((vt) => dailyData[d]?.[vt] || 0);
      const rowSum = rowTotals[d] || 0;
      rows.push([dateLabel, ...rowVals, rowSum]);
    }

    // Grand total row
    const totalRow = ['សរុប', ...STATS_VISA_TYPES.map((vt) => columnTotals[vt] || 0), grandTotal];
    rows.push(totalRow);

    // Footer signature rows
    rows.push(['']);
    rows.push(['', '', '', '', '', '', '', '', '', customLunarDate]);
    rows.push(['', '', '', '', '', '', '', '', '', formattedSigningDateKhmer]);
    rows.push(['', '', '', '', '', '', '', '', '', chiefTitle]);
    if (chiefName) {
      rows.push(['', '', '', '', '', '', '', '', '', '']);
      rows.push(['', '', '', '', '', '', '', '', '', chiefName]);
    }

    const ws = XLSX.utils.aoa_to_sheet(rows);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Daily_Team_Stats');
    const safeTeam = (selectedTeam || 'team').replace(/\s+/g, '_');
    const sheetPrefix = isOffice ? 'ស្ថិតិសរុបប្រចាំថ្ងៃរបស់ក្រុម' : 'តារាងស្ថិតិប្រចាំថ្ងៃរបស់ក្រុម';
    XLSX.writeFile(wb, `${sheetPrefix}_${safeTeam}_${selectedYear}_M${selectedMonth}.xlsx`);
  };

  return (
    <div className="min-h-screen bg-slate-100 p-2 sm:p-4 text-slate-800 font-siemreap">
      {/* Control Bar (Hidden on Print) */}
      <div className="print:hidden max-w-5xl mx-auto mb-4 bg-white border border-gray-200 rounded-md shadow-2xs p-3.5 sm:p-4 space-y-3 transition-all">
        {/* Row 1: Header Title & Main Action Buttons */}
        <div className="flex flex-wrap items-center justify-between gap-3 pb-3 border-b border-gray-200">
          <div className="flex items-start gap-2.5 min-w-0">
            <div className="w-9 h-9 rounded bg-blue-50 text-blue-600 flex items-center justify-center font-bold shrink-0 mt-0.5">
              <FileSpreadsheet className="w-5 h-5 text-blue-600" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h2 className="text-sm sm:text-base font-bold text-gray-800 leading-snug">
                  {isOffice ? 'ស្ថិតិសរុបប្រចាំថ្ងៃរបស់ក្រុម' : 'តារាងស្ថិតិប្រចាំថ្ងៃរបស់ក្រុម'}
                </h2>
                <span className={`px-2 py-0.5 rounded text-[11px] font-bold tracking-wide border ${
                  selectedOption === 'cEA'
                    ? 'bg-cyan-50 text-cyan-800 border-cyan-300'
                    : 'bg-blue-50 text-blue-800 border-blue-300'
                }`}>
                  {selectedOption === 'cEA' ? 'ក្រដាសអនុម័ត (cEA)' : 'សន្លឹកទិដ្ឋាការ (Sticker)'}
                </span>
              </div>
              <p className="text-xs text-gray-500 mt-0.5">
                {isOffice
                  ? `ស្ថិតិភ្ញៀវស្នើសុំទិដ្ឋាការចូលប្រទេស សរុបប្រចាំថ្ងៃតាមបណ្តាក្រុម (${selectedOption === 'cEA' ? 'cEA' : 'Sticker'})`
                  : `ស្ថិតិភ្ញៀវស្នើសុំទិដ្ឋាការចូលប្រទេស ប្រចាំថ្ងៃតាមបណ្តាក្រុម (${selectedOption === 'cEA' ? 'cEA' : 'Sticker'})`}
              </p>
            </div>
          </div>

          {/* Action Buttons */}
          <div className="flex flex-wrap items-center gap-1.5 sm:gap-2">
            <button
              type="button"
              onClick={handleSave}
              className="px-3 py-1.5 rounded text-xs font-bold bg-blue-600 hover:bg-blue-700 text-white transition flex items-center gap-1.5 cursor-pointer shadow-2xs"
              title="រក្សាទុកទិន្នន័យដែលបានកែសម្រួល"
            >
              <Save className="w-3.5 h-3.5" />
              <span>រក្សាទុក</span>
            </button>

            <button
              type="button"
              onClick={handleExportExcel}
              className="px-3 py-1.5 rounded text-xs font-bold bg-emerald-600 hover:bg-emerald-700 text-white transition flex items-center gap-1.5 cursor-pointer shadow-2xs"
              title="នាំចេញជាទម្រង់ Excel (.xlsx)"
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              <span>Excel</span>
            </button>

            <button
              type="button"
              onClick={handleExportPDF}
              disabled={isExporting}
              className="px-3.5 py-1.5 rounded text-xs font-bold bg-red-600 hover:bg-red-700 text-white transition flex items-center gap-1.5 cursor-pointer shadow-2xs disabled:opacity-50"
              title="ទាញយកឯកសារជា PDF គុណភាពខ្ពស់ដូចនៅ របក.សរុបការងារស្តុក"
            >
              <Download className="w-3.5 h-3.5" />
              <span>{isExporting ? 'កំពុងបង្កើត PDF...' : 'ទាញយក PDF'}</span>
            </button>

            {/* Toggle Show/Hide Controls Panel (Picture ទី២) */}
            <button
              type="button"
              onClick={() => {
                const nextVal = !showControlPanel;
                setShowControlPanel(nextVal);
                setStoredSetting('show_control_panel', String(nextVal));
              }}
              className={`px-3 py-1.5 rounded text-xs font-bold border transition flex items-center gap-1.5 cursor-pointer shadow-2xs ${
                showControlPanel
                  ? 'bg-slate-800 hover:bg-slate-900 text-white border-slate-900'
                  : 'bg-white hover:bg-gray-100 text-gray-700 border-gray-300'
              }`}
              title={showControlPanel ? 'ចុចដើម្បីលាក់ផ្ទាំងឧបករណ៍/ជម្រើសកែសម្រួល (រូបភាពទី២)' : 'ចុចដើម្បីបង្ហាញផ្ទាំងឧបករណ៍/ជម្រើសកែសម្រួល (រូបភាពទី២)'}
            >
              {showControlPanel ? (
                <>
                  <EyeOff className="w-3.5 h-3.5 text-amber-300" />
                  <span>លាក់ផ្ទាំងជម្រើស</span>
                </>
              ) : (
                <>
                  <Eye className="w-3.5 h-3.5 text-blue-600" />
                  <span>បង្ហាញផ្ទាំងជម្រើស</span>
                </>
              )}
            </button>

            {onClose && (
              <button
                type="button"
                onClick={onClose}
                className="px-3 py-1.5 rounded text-xs font-semibold bg-gray-100 hover:bg-gray-200 border border-gray-300 text-gray-700 transition flex items-center gap-1.5 cursor-pointer shadow-2xs"
              >
                ចាកចេញ
              </button>
            )}
          </div>
        </div>

        {/* Row 2: Combined 1 Row matching image (Sticker/cEA, Team, Month, Year, Data Buttons, Checklist, Clear) - Always Visible */}
        <div className="flex items-center gap-2 overflow-x-auto no-scrollbar flex-nowrap min-w-0 py-2 border-b border-gray-200">
          {/* Option Buttons: Sticker vs cEA */}
          <div className="flex items-center bg-gray-100 p-0.5 rounded border border-gray-300 shrink-0 h-[34px] shadow-2xs">
            <button
              type="button"
              onClick={() => handleSelectOption('Sticker')}
              className={`h-full px-3 rounded text-xs font-bold transition flex items-center gap-1.5 cursor-pointer whitespace-nowrap ${
                selectedOption === 'Sticker'
                  ? 'bg-blue-600 text-white shadow-xs'
                  : 'text-gray-600 hover:text-blue-900 hover:bg-gray-200'
              }`}
              title="បង្ហាញស្ថិតិ និងក្រុមប្រើប្រាស់ សន្លឹកទិដ្ឋាការ (Sticker)"
            >
              <span className={`w-2 h-2 rounded-full ${selectedOption === 'Sticker' ? 'bg-amber-300' : 'bg-gray-400'}`} />
              <span>Sticker</span>
              <span className={`text-[10px] px-1.5 py-0.2 rounded font-mono ${
                selectedOption === 'Sticker' ? 'bg-white/20 text-white' : 'bg-gray-200 text-gray-600'
              }`}>
                {stickerTeams.length}
              </span>
            </button>

            {(isSecondary || isTeamEVisaAllowed) && (
              <button
                type="button"
                onClick={() => handleSelectOption('cEA')}
                className={`h-full px-3 rounded text-xs font-bold transition flex items-center gap-1.5 cursor-pointer whitespace-nowrap ${
                  selectedOption === 'cEA'
                    ? 'bg-cyan-700 text-white shadow-xs'
                    : 'text-gray-600 hover:text-cyan-900 hover:bg-gray-200'
                }`}
                title="បង្ហាញស្ថិតិ និងក្រុមប្រើប្រាស់ ក្រដាសអនុម័ត (cEA)"
              >
                <span className={`w-2 h-2 rounded-full ${selectedOption === 'cEA' ? 'bg-cyan-200' : 'bg-gray-400'}`} />
                <span>cEA</span>
                <span className={`text-[10px] px-1.5 py-0.2 rounded font-mono ${
                  selectedOption === 'cEA' ? 'bg-white/20 text-white' : 'bg-gray-200 text-gray-600'
                }`}>
                  {ceaTeams.length}
                </span>
              </button>
            )}
          </div>

          {/* Team Selector (Only shown for Office user in ស្ថិតិសរុបប្រចាំថ្ងៃរបស់ក្រុម; hidden for Team user in តារាងស្ថិតិប្រចាំថ្ងៃរបស់ក្រុម) */}
          {isOffice && (
            <div className="flex items-center gap-1.5 shrink-0 h-[34px]">
              <span className="font-bold text-gray-700 whitespace-nowrap text-xs">
                {selectedOption === 'cEA' ? 'ក្រុម cEA៖' : 'ក្រុម Sticker៖'}
              </span>
              <select
                value={selectedTeam}
                onChange={(e) => setSelectedTeam(e.target.value)}
                className="h-[34px] min-h-[34px] max-h-[34px] box-border border border-gray-300 rounded px-2.5 py-1 font-sans text-xs focus:border-blue-500 focus:outline-none font-bold text-gray-800 bg-white cursor-pointer shadow-2xs"
              >
                {allTeams.map((t) => (
                  <option key={t} value={t}>
                    {t}
                  </option>
                ))}
              </select>
            </div>
          )}

          {/* Month Selector */}
          <div className="flex items-center gap-1.5 shrink-0 h-[34px]">
            <span className="font-bold text-gray-700 whitespace-nowrap text-xs">ខែ៖</span>
            <select
              value={selectedMonth}
              onChange={(e) => setSelectedMonth(parseInt(e.target.value, 10))}
              className="h-[34px] min-h-[34px] max-h-[34px] box-border border border-gray-300 rounded px-2.5 py-1 font-sans text-xs focus:border-blue-500 focus:outline-none font-bold text-gray-800 bg-white cursor-pointer shadow-2xs"
            >
              {KHMER_MONTHS.map((mName, idx) => (
                <option key={idx + 1} value={idx + 1}>
                  ខែ {mName} (M{idx + 1})
                </option>
              ))}
            </select>
          </div>

          {/* Year Selector */}
          <div className="flex items-center gap-1.5 shrink-0 h-[34px]">
            <span className="font-bold text-gray-700 whitespace-nowrap text-xs">ឆ្នាំ៖</span>
            <select
              value={selectedYear}
              onChange={(e) => setSelectedYear(parseInt(e.target.value, 10))}
              className="h-[34px] min-h-[34px] max-h-[34px] box-border border border-gray-300 rounded px-2.5 py-1 font-sans text-xs focus:border-blue-500 focus:outline-none font-bold text-gray-800 bg-white cursor-pointer shadow-2xs"
            >
              {[2024, 2025, 2026, 2027, 2028, 2029, 2030].map((y) => (
                <option key={y} value={y}>
                  ឆ្នាំ {y} ({toKhmerNum(y)})
                </option>
              ))}
            </select>
          </div>

          {/* Action Button: បង្ហាញទិន្នន័យ */}
          <button
            onClick={pullFromTeamUsage}
            className="h-[34px] shrink-0 inline-flex items-center gap-1.5 px-3.5 bg-blue-600 hover:bg-blue-700 text-white rounded text-xs font-bold shadow-2xs transition cursor-pointer whitespace-nowrap"
            title={isTotalReport ? "ចុចដើម្បីទាញ ឬធ្វើបច្ចុប្បន្នភាពទិន្នន័យប្រើប្រាស់សរុបគ្រប់ក្រុម" : "ចុចដើម្បីបង្ហាញទិន្នន័យពីប្រតិបត្តិការ «ការប្រើប្រាស់តាមក្រុម»"}
          >
            <RefreshCw className="w-3.5 h-3.5 text-blue-100 shrink-0" />
            <span>បង្ហាញទិន្នន័យ</span>
          </button>

          {/* Action Button: ផ្ទៀងផ្ទាត់ទិន្នន័យ */}
          <button
            type="button"
            onClick={() => setIsChecklistOpen(true)}
            className="h-[34px] shrink-0 inline-flex items-center gap-1.5 px-3.5 bg-indigo-600 hover:bg-indigo-700 text-white rounded text-xs font-bold shadow-2xs transition cursor-pointer whitespace-nowrap"
            title="ផ្ទៀងផ្ទាត់បញ្ជីទិន្នន័យ & ស្វែងរកទិន្នន័យរាប់លើស/ខុស (Checklist)"
          >
            <CheckSquare className="w-3.5 h-3.5 text-yellow-300 shrink-0" />
            <span>ផ្ទៀងផ្ទាត់ទិន្នន័យ</span>
          </button>

          {/* Action Button: ជម្រះ (-) */}
          <button
            onClick={initEmptyMonth}
            className="h-[34px] shrink-0 inline-flex items-center gap-1 px-3 bg-gray-100 hover:bg-gray-200 text-gray-700 border border-gray-300 rounded text-xs font-semibold shadow-2xs transition cursor-pointer whitespace-nowrap"
            title="សម្អាតតួលេខទាំងអស់ឱ្យទៅជា (-)"
          >
            <RotateCcw className="w-3.5 h-3.5 text-gray-500" />
            <span>ជម្រះ (-)</span>
          </button>
        </div>

        {/* Picture ទី២: Formatting Toolbar (Steppers, Zoom, Settings & Tacteing Line) - Toggled by លាក់ផ្ទាំងជម្រើស */}
        {showControlPanel && (
          <>
            {/* Row 3: Steppers & Tools Row (គែមលើ, អក្សរក្បាល, អក្សរតារាង, កម្ពស់ជួរ, Zoom, កែសម្រួលក្បាល/កន្ទុយ) */}
            <div className="flex flex-wrap items-center gap-2 pt-1 text-xs">
          {/* Quick Top Margin Stepper (ទម្លាក់/បង្រួមគម្លាតពីលើចុះក្រោម) */}
          <div className="flex items-center border border-gray-300 rounded bg-white shadow-2xs px-2 py-1 text-xs text-gray-700 font-siemreap">
            <span className="text-[11px] text-gray-500 mr-1.5 select-none font-medium">គែមលើ៖</span>
            <button
              type="button"
              onClick={() => {
                const val = Math.max(0.1, parseFloat((pageTopMargin - 0.1).toFixed(2)));
                setPageTopMargin(val);
                setStoredSetting('top_margin', String(val));
              }}
              className="w-5 h-5 flex items-center justify-center rounded hover:bg-gray-100 text-gray-700 font-bold cursor-pointer transition"
              title="បង្រួមគម្លាតពីលើ (-0.1cm)"
            >
              -
            </button>
            <span className="font-mono text-xs font-bold px-1.5 text-blue-900 select-none">
              {pageTopMargin.toFixed(1)}cm
            </span>
            <button
              type="button"
              onClick={() => {
                const val = Math.min(3.5, parseFloat((pageTopMargin + 0.1).toFixed(2)));
                setPageTopMargin(val);
                setStoredSetting('top_margin', String(val));
              }}
              className="w-5 h-5 flex items-center justify-center rounded hover:bg-gray-100 text-gray-700 font-bold cursor-pointer transition"
              title="ទម្លាក់គម្លាតពីលើចុះ (+0.1cm)"
            >
              +
            </button>
          </div>

          {/* Quick Header Font Size Stepper (ទំហំអក្សរក្បាលលិខិត) */}
          <div className="flex items-center border border-gray-300 rounded bg-white shadow-2xs px-2 py-1 text-xs text-gray-700 font-siemreap">
            <span className="text-[11px] text-gray-500 mr-1.5 select-none font-medium">អក្សរក្បាល៖</span>
            <button
              type="button"
              onClick={() => {
                const val = Math.max(8, parseFloat((headerFontSizePt - 0.5).toFixed(1)));
                setHeaderFontSizePt(val);
                setStoredSetting('header_font_size', String(val));
              }}
              className="w-5 h-5 flex items-center justify-center rounded hover:bg-gray-100 text-gray-700 font-bold cursor-pointer transition"
              title="បង្រួមទំហំអក្សរ (-0.5pt)"
            >
              -
            </button>
            <span className="font-mono text-xs font-bold px-1.5 text-blue-900 select-none">
              {headerFontSizePt}pt
            </span>
            <button
              type="button"
              onClick={() => {
                const val = Math.min(18, parseFloat((headerFontSizePt + 0.5).toFixed(1)));
                setHeaderFontSizePt(val);
                setStoredSetting('header_font_size', String(val));
              }}
              className="w-5 h-5 flex items-center justify-center rounded hover:bg-gray-100 text-gray-700 font-bold cursor-pointer transition"
              title="ពង្រីកទំហំអក្សរ (+0.5pt)"
            >
              +
            </button>
          </div>

          {/* Quick Table Font Size Stepper (ទំហំអក្សរក្នុងតារាង / Table Font Size) */}
          <div className="flex items-center border border-gray-300 rounded bg-white shadow-2xs px-2 py-1 text-xs text-gray-700 font-siemreap">
            <span className="text-[11px] text-gray-500 mr-1.5 select-none font-medium">អក្សរតារាង៖</span>
            <button
              type="button"
              onClick={() => {
                const val = Math.max(8, parseFloat((tableFontSize - 0.5).toFixed(1)));
                setTableFontSize(val);
                setStoredSetting('table_font_size', String(val));
              }}
              className="w-5 h-5 flex items-center justify-center rounded hover:bg-gray-100 text-gray-700 font-bold cursor-pointer transition"
              title="បន្ថយទំហំអក្សរក្នុងតារាង (-0.5px)"
            >
              -
            </button>
            <span className="font-mono text-xs font-bold px-1.5 text-blue-900 select-none">
              {tableFontSize}px
            </span>
            <button
              type="button"
              onClick={() => {
                const val = Math.min(18, parseFloat((tableFontSize + 0.5).toFixed(1)));
                setTableFontSize(val);
                setStoredSetting('table_font_size', String(val));
              }}
              className="w-5 h-5 flex items-center justify-center rounded hover:bg-gray-100 text-gray-700 font-bold cursor-pointer transition"
              title="តម្លើងទំហំអក្សរក្នុងតារាង (+0.5px)"
            >
              +
            </button>
          </div>

          {/* Quick Table Row Height Stepper (បង្រួម / ពង្រីកកម្ពស់ជួរតារាង) */}
          <div className="flex items-center border border-gray-300 rounded bg-white shadow-2xs px-2 py-1 text-xs text-gray-700 font-siemreap">
            <span className="text-[11px] text-gray-500 mr-1.5 select-none font-medium">កម្ពស់ជួរ៖</span>
            <button
              type="button"
              onClick={() => {
                const val = Math.max(15, tableRowHeight - 1);
                setTableRowHeight(val);
                setStoredSetting('row_height', String(val));
              }}
              className="w-5 h-5 flex items-center justify-center rounded hover:bg-gray-100 text-gray-700 font-bold cursor-pointer transition"
              title="បង្រួមកម្ពស់ជួរតារាង (-1px)"
            >
              -
            </button>
            <span className="font-mono text-xs font-bold px-1.5 text-blue-900 select-none">
              {tableRowHeight}px
            </span>
            <button
              type="button"
              onClick={() => {
                const val = Math.min(32, tableRowHeight + 1);
                setTableRowHeight(val);
                setStoredSetting('row_height', String(val));
              }}
              className="w-5 h-5 flex items-center justify-center rounded hover:bg-gray-100 text-gray-700 font-bold cursor-pointer transition"
              title="ពង្រីกกម្ពស់ជួរតារាង (+1px)"
            >
              +
            </button>
          </div>

          {/* Zoom Controls */}
          <div className="flex items-center border border-gray-300 rounded bg-white overflow-hidden shadow-2xs">
            <button
              onClick={() => setZoomScale((z) => Math.max(z - 10, 70))}
              className="px-2 py-1 hover:bg-gray-100 text-gray-700 cursor-pointer"
              title="បង្រួម (Zoom Out)"
            >
              <ZoomOut className="w-3.5 h-3.5" />
            </button>
            <span className="px-2 py-0.5 text-[11px] font-mono font-bold text-gray-700 min-w-[42px] text-center">
              {zoomScale}%
            </span>
            <button
              onClick={() => setZoomScale((z) => Math.min(z + 10, 130))}
              className="px-2 py-1 hover:bg-gray-100 text-gray-700 cursor-pointer"
              title="ពង្រីក (Zoom In)"
            >
              <ZoomIn className="w-3.5 h-3.5" />
            </button>
          </div>

          {/* Header/Footer Settings Toggle */}
          <button
            onClick={() => setShowHeaderSettings(!showHeaderSettings)}
            className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded border text-xs font-bold transition cursor-pointer shadow-2xs ${
              showHeaderSettings
                ? 'bg-blue-600 text-white border-blue-600'
                : 'bg-white text-gray-700 border-gray-300 hover:bg-gray-50'
            }`}
          >
            <SlidersHorizontal className="w-3.5 h-3.5" />
            <span>កែសម្រួលក្បាល/កន្ទុយ</span>
          </button>

          {/* Quick Picture / Tacteing Hide-Show Toggle Button with Options Dropdown */}
          <div className="relative inline-block">
            <div className="inline-flex rounded border border-gray-300 bg-white shadow-2xs overflow-hidden">
              <button
                type="button"
                onClick={() => {
                  const nextVal = !showTacteingPicture;
                  setShowTacteingPicture(nextVal);
                  setStoredSetting('show_tacteing_picture', String(nextVal));
                }}
                className={`inline-flex items-center gap-1.5 px-3 py-1.5 text-xs font-bold transition cursor-pointer ${
                  showTacteingPicture
                    ? 'bg-amber-50 hover:bg-amber-100 text-amber-900'
                    : 'bg-white hover:bg-gray-100 text-gray-500 line-through'
                }`}
                title={showTacteingPicture ? 'ចុចដើម្បីលាក់រូបភាពតាក់តែង' : 'ចុចដើម្បីបង្ហាញរូបភាពតាក់តែង'}
              >
                {showTacteingPicture ? (
                  <Eye className="w-3.5 h-3.5 text-amber-600" />
                ) : (
                  <EyeOff className="w-3.5 h-3.5 text-gray-400" />
                )}
                <span>{showTacteingPicture ? 'រូបភាពតាក់តែង (បង្ហាញ)' : 'រូបភាពតាក់តែង (លាក់)'}</span>
              </button>

              <button
                type="button"
                onClick={() => setShowTacteingDropdown((prev) => !prev)}
                className="px-1.5 border-l border-gray-300 hover:bg-gray-100 text-gray-600 transition cursor-pointer"
                title="ជម្រើសបង្ហាញ/លាក់រូបភាព"
              >
                <ChevronDown className="w-3.5 h-3.5" />
              </button>
            </div>

            {showTacteingDropdown && (
              <div className="absolute right-0 mt-1 w-64 bg-white border border-gray-200 rounded-lg shadow-xl z-50 p-2 text-xs">
                <div className="text-[11px] font-bold text-gray-700 px-2 py-1 mb-1 border-b border-gray-100 flex items-center justify-between">
                  <span>ជម្រើសបង្ហាញ/លាក់រូបភាពតាក់តែង</span>
                  <span className="text-[10px] text-gray-400 font-mono">Options</span>
                </div>

                <button
                  type="button"
                  onClick={() => {
                    setShowTacteingPicture(true);
                    setShowLeftTacteing(true);
                    setShowRightTacteing(true);
                    setStoredSetting('show_tacteing_picture', 'true');
                    setStoredSetting('show_left_tacteing', 'true');
                    setStoredSetting('show_right_tacteing', 'true');
                    setShowTacteingDropdown(false);
                  }}
                  className={`w-full text-left px-2 py-1.5 rounded flex items-center justify-between hover:bg-amber-50 transition cursor-pointer ${
                    showTacteingPicture && showLeftTacteing && showRightTacteing ? 'font-bold text-amber-900 bg-amber-50/70' : 'text-gray-700'
                  }`}
                >
                  <span>បង្ហាញទាំងសងខាង (ឆ្វេង & ស្តាំ)</span>
                  {showTacteingPicture && showLeftTacteing && showRightTacteing && (
                    <CheckCircle2 className="w-3.5 h-3.5 text-amber-600" />
                  )}
                </button>

                <button
                  type="button"
                  onClick={() => {
                    setShowTacteingPicture(true);
                    setShowLeftTacteing(false);
                    setShowRightTacteing(true);
                    setStoredSetting('show_tacteing_picture', 'true');
                    setStoredSetting('show_left_tacteing', 'false');
                    setStoredSetting('show_right_tacteing', 'true');
                    setShowTacteingDropdown(false);
                  }}
                  className={`w-full text-left px-2 py-1.5 rounded flex items-center justify-between hover:bg-amber-50 transition cursor-pointer ${
                    showTacteingPicture && !showLeftTacteing && showRightTacteing ? 'font-bold text-amber-900 bg-amber-50/70' : 'text-gray-700'
                  }`}
                >
                  <span>បង្ហាញតែខាងស្តាំ (ក្រោមជាតិ សាសនា...)</span>
                  {showTacteingPicture && !showLeftTacteing && showRightTacteing && (
                    <CheckCircle2 className="w-3.5 h-3.5 text-amber-600" />
                  )}
                </button>

                <button
                  type="button"
                  onClick={() => {
                    setShowTacteingPicture(true);
                    setShowLeftTacteing(true);
                    setShowRightTacteing(false);
                    setStoredSetting('show_tacteing_picture', 'true');
                    setStoredSetting('show_left_tacteing', 'true');
                    setStoredSetting('show_right_tacteing', 'false');
                    setShowTacteingDropdown(false);
                  }}
                  className={`w-full text-left px-2 py-1.5 rounded flex items-center justify-between hover:bg-amber-50 transition cursor-pointer ${
                    showTacteingPicture && showLeftTacteing && !showRightTacteing ? 'font-bold text-amber-900 bg-amber-50/70' : 'text-gray-700'
                  }`}
                >
                  <span>បង្ហាញតែខាងឆ្វេង (ក្រោមឈ្មោះក្រុម/ផ្នែក)</span>
                  {showTacteingPicture && showLeftTacteing && !showRightTacteing && (
                    <CheckCircle2 className="w-3.5 h-3.5 text-amber-600" />
                  )}
                </button>

                <button
                  type="button"
                  onClick={() => {
                    setShowTacteingPicture(false);
                    setStoredSetting('show_tacteing_picture', 'false');
                    setShowTacteingDropdown(false);
                  }}
                  className={`w-full text-left px-2 py-1.5 rounded flex items-center justify-between hover:bg-rose-50 transition cursor-pointer text-rose-600 mt-1 pt-1 border-t border-gray-100 ${
                    !showTacteingPicture ? 'font-bold bg-rose-50' : ''
                  }`}
                >
                  <span>លាក់រូបភាពទាំងអស់ (Hide All)</span>
                  {!showTacteingPicture && (
                    <CheckCircle2 className="w-3.5 h-3.5 text-rose-600" />
                  )}
                </button>
              </div>
            )}
          </div>
        </div>

        {/* Collapsible Header/Footer Settings Box */}
        {showHeaderSettings && (
          <div className="mt-3 pt-3 border-t border-slate-200 grid grid-cols-1 sm:grid-cols-3 gap-3 text-xs bg-slate-50 p-3 rounded-lg border border-slate-200">
            <div>
              <label className="block font-semibold text-slate-700 mb-1">
                ក្រសួង (Ministry Name)៖
              </label>
              <input
                type="text"
                value={ministryName}
                onChange={(e) => setMinistryName(e.target.value)}
                className="w-full h-8 min-h-[32px] max-h-[32px] box-border px-2.5 py-1 border border-slate-300 rounded bg-white text-xs font-medium focus:ring-1 focus:ring-blue-500 outline-none"
              />
            </div>

            <div>
              <label className="block font-semibold text-slate-700 mb-1">
                អគ្គនាយកដ្ឋាន (Department)៖
              </label>
              <input
                type="text"
                value={departmentName}
                onChange={(e) => setDepartmentName(e.target.value)}
                className="w-full h-8 min-h-[32px] max-h-[32px] box-border px-2.5 py-1 border border-slate-300 rounded bg-white text-xs font-medium focus:ring-1 focus:ring-blue-500 outline-none"
              />
            </div>

            <div>
              <label className="block font-semibold text-slate-700 mb-1">
                នាយកដ្ឋាន (General Department)៖
              </label>
              <input
                type="text"
                value={generalDeptName}
                onChange={(e) => setGeneralDeptName(e.target.value)}
                className="w-full h-8 min-h-[32px] max-h-[32px] box-border px-2.5 py-1 border border-slate-300 rounded bg-white text-xs font-medium focus:ring-1 focus:ring-blue-500 outline-none"
              />
            </div>

            <div>
              <label className="block font-semibold text-slate-700 mb-1">
                ការិយាល័យ (Office Name)៖
              </label>
              <input
                type="text"
                value={officeName}
                onChange={(e) => setOfficeName(e.target.value)}
                className="w-full h-8 min-h-[32px] max-h-[32px] box-border px-2.5 py-1 border border-slate-300 rounded bg-white text-xs font-medium focus:ring-1 focus:ring-blue-500 outline-none"
              />
            </div>

            <div>
              <label className="block font-semibold text-slate-700 mb-1">
                {isTotalReport ? 'ផ្នែកបង្ហាញនៅក្បាលទំព័រ (ក្រោមការិយាល័យទិដ្ឋាការចូល)៖' : 'ឈ្មោះក្រុមបង្ហាញនៅក្បាលទំព័រ៖'}
              </label>
              <input
                type="text"
                value={customTeamTitle}
                onChange={(e) => setCustomTeamTitle(e.target.value)}
                className="w-full h-8 min-h-[32px] max-h-[32px] box-border px-2.5 py-1 border border-slate-300 rounded bg-white text-xs font-medium focus:ring-1 focus:ring-blue-500 outline-none"
                placeholder={isTotalReport ? "ផ្នែករដ្ឋបាល" : "ឧ. ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិភ្នំដិន"}
              />
              <div className="mt-1.5 flex items-center gap-2">
                <span className="text-[11px] text-slate-600 font-medium">ពុម្ពអក្សរ៖</span>
                <button
                  type="button"
                  onClick={() => setTeamTitleFont('siemreap')}
                  className={`px-2 py-0.5 text-[11px] rounded border transition-colors ${
                    teamTitleFont === 'siemreap'
                      ? 'bg-blue-600 text-white border-blue-600 font-semibold shadow-xs'
                      : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
                  }`}
                >
                  Khmer OS Siemreap
                </button>
                <button
                  type="button"
                  onClick={() => setTeamTitleFont('moul')}
                  className={`px-2 py-0.5 text-[11px] rounded border transition-colors ${
                    teamTitleFont === 'moul'
                      ? 'bg-blue-600 text-white border-blue-600 font-semibold shadow-xs'
                      : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
                  }`}
                >
                  Khmer OS Mool
                </button>
              </div>
              <div className="mt-1.5 flex items-center gap-1.5 flex-wrap">
                <span className="text-[11px] text-slate-600 font-medium">ទំហំអក្សរឈ្មោះក្រុម៖</span>
                {[9, 9.5, 10, 10.5, 11, 12].map((sz) => (
                  <button
                    key={sz}
                    type="button"
                    onClick={() => {
                      setTeamTitleFontSizePt(sz);
                      localStorage.setItem('daily_team_stats_team_title_font_size', String(sz));
                    }}
                    className={`text-[10.5px] px-2 py-0.5 rounded border transition cursor-pointer ${
                      teamTitleFontSizePt === sz
                        ? 'bg-[#002060] text-white font-bold border-[#002060]'
                        : 'bg-white text-slate-700 hover:bg-slate-100 border-slate-300'
                    }`}
                  >
                    {sz}pt
                  </button>
                ))}
              </div>
            </div>

            <div>
              <label className="block font-semibold text-slate-700 mb-1">
                កាលបរិច្ឆេទចុះហត្ថលេខា (Signing Date)៖
              </label>
              <CustomDatePicker
                value={signingDate}
                onChange={setSigningDate}
                className="h-8 min-h-[32px] max-h-[32px] box-border px-2.5 py-1 border border-slate-300 rounded bg-white text-xs font-medium focus:ring-1 focus:ring-blue-500 outline-none"
              />
            </div>

            <div>
              <label className="block font-semibold text-slate-700 mb-1">
                ទីកន្លែងចុះហត្ថលេខា (Location)៖
              </label>
              <input
                type="text"
                value={signingLocation}
                onChange={(e) => setSigningLocation(e.target.value)}
                className="w-full h-8 min-h-[32px] max-h-[32px] box-border px-2.5 py-1 border border-slate-300 rounded bg-white text-xs font-medium focus:ring-1 focus:ring-blue-500 outline-none"
                placeholder="ឧ. ភ្នំពេញ ឬ ត្រពាំងផ្លុង"
              />
            </div>

            <div className="sm:col-span-2">
              <label className="block font-semibold text-slate-700 mb-1">
                កាលបរិច្ឆេទចន្ទគតិ (Khmer Lunar Calendar)៖
              </label>
              <input
                type="text"
                value={customLunarDate}
                onChange={(e) => setCustomLunarDate(e.target.value)}
                className="w-full h-8 min-h-[32px] max-h-[32px] box-border px-2.5 py-1 border border-slate-300 rounded bg-white text-xs font-medium focus:ring-1 focus:ring-blue-500 outline-none"
                placeholder="ឧ. ថ្ងៃសៅរ៍ ៣រោច ខែទុតិយាសាឍ ឆ្នាំម្សាញ់ អដ្ឋស័ក ព.ស ២៥៧០"
              />
            </div>

            <div>
              <label className="block font-semibold text-slate-700 mb-1">
                តួនាទីអ្នកចុះហត្ថលេខា៖
              </label>
              <input
                type="text"
                value={chiefTitle}
                onChange={(e) => setChiefTitle(e.target.value)}
                className="w-full h-8 min-h-[32px] max-h-[32px] box-border px-2.5 py-1 border border-slate-300 rounded bg-white text-xs font-medium focus:ring-1 focus:ring-blue-500 outline-none"
                placeholder="ឧ. ប្រធានក្រុមផ្តល់ទិដ្ឋាការ"
              />
              <div className="mt-1.5 flex items-center gap-2">
                <span className="text-[11px] text-slate-600 font-medium">ពុម្ពអក្សរ៖</span>
                <button
                  type="button"
                  onClick={() => setChiefTitleFont('muol')}
                  className={`px-2 py-0.5 text-[11px] rounded border transition-colors ${
                    chiefTitleFont === 'muol'
                      ? 'bg-blue-600 text-white border-blue-600 font-semibold shadow-xs'
                      : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
                  }`}
                >
                  Khmer OS Muol Light
                </button>
                <button
                  type="button"
                  onClick={() => setChiefTitleFont('siemreap')}
                  className={`px-2 py-0.5 text-[11px] rounded border transition-colors ${
                    chiefTitleFont === 'siemreap'
                      ? 'bg-blue-600 text-white border-blue-600 font-semibold shadow-xs'
                      : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
                  }`}
                >
                  Khmer OS Siemreap
                </button>
              </div>
            </div>

            <div>
              <label className="block font-semibold text-slate-700 mb-1">
                ឈ្មោះប្រធានក្រុម / អ្នកចុះហត្ថលេខា (Signer Name)៖
              </label>
              <input
                type="text"
                value={chiefName}
                onChange={(e) => setChiefName(e.target.value)}
                className="w-full h-8 min-h-[32px] max-h-[32px] box-border px-2.5 py-1 border border-slate-300 rounded bg-white text-xs font-medium focus:ring-1 focus:ring-blue-500 outline-none"
                placeholder="ឈ្មោះ (បើមាន)"
              />
            </div>

            <div>
              <label className="block font-semibold text-slate-700 mb-1">
                គម្លាតស្តាំក្បាលទំព័រ (Right Margin px)៖
              </label>
              <input
                type="number"
                min={-50}
                max={150}
                step={2}
                value={rightHeaderMargin}
                onChange={(e) => setRightHeaderMargin(Number(e.target.value) || 0)}
                className="w-full h-8 min-h-[32px] max-h-[32px] box-border px-2.5 py-1 border border-slate-300 rounded bg-white text-xs font-medium focus:ring-1 focus:ring-blue-500 outline-none"
              />
              <span className="text-[10px] text-slate-500">លំនាំដើម 0px (ស្មើគែមស្តាំតារាងដូចរូបភាពទី២)</span>
            </div>

            <div>
              <label className="block font-semibold text-slate-700 mb-1">
                គម្លាតហត្ថលេខាឡើងលើ/ចុះក្រោម (Signature Margin-Top)៖
              </label>
              <div className="flex items-center gap-2">
                <input
                  type="number"
                  min={-30}
                  max={60}
                  step={2}
                  value={signatureMarginTop}
                  onChange={(e) => {
                    const val = Number(e.target.value) || 0;
                    setSignatureMarginTop(val);
                    localStorage.setItem('daily_team_stats_signature_mt', String(val));
                  }}
                  className="w-24 h-8 min-h-[32px] max-h-[32px] box-border px-2.5 py-1 border border-slate-300 rounded bg-white text-xs font-medium focus:ring-1 focus:ring-blue-500 outline-none"
                />
                <span className="text-[10.5px] text-slate-500">px</span>
                <button
                  type="button"
                  onClick={() => {
                    setSignatureMarginTop(0);
                    localStorage.setItem('daily_team_stats_signature_mt', '0');
                  }}
                  className="h-8 px-2.5 bg-slate-100 hover:bg-slate-200 border border-slate-300 rounded text-xs text-slate-700 font-medium cursor-pointer"
                  title="ដោះការរំកិលកំណត់ទៅ 0px"
                >
                  ដោះការរំកិល (0px)
                </button>
              </div>
            </div>

            {/* Page Margins (Default left: 2.80cm, right: 1.30cm, top: 1.0cm, bottom: 0.5cm) */}
            <div>
              <label className="block font-semibold text-slate-700 mb-1">
                គម្លាតទំព័រខាងឆ្វេង (Left Margin cm) ៖
              </label>
              <div className="flex items-center gap-2">
                <input
                  type="number"
                  min={0.2}
                  max={5.0}
                  step={0.05}
                  value={pageLeftMargin}
                  onChange={(e) => {
                    const val = Number(e.target.value) || 2.80;
                    setPageLeftMargin(val);
                    localStorage.setItem('daily_team_stats_left_margin', String(val));
                  }}
                  className="w-24 h-8 min-h-[32px] max-h-[32px] box-border px-2.5 py-1 border border-slate-300 rounded bg-white text-xs font-medium focus:ring-1 focus:ring-blue-500 outline-none"
                />
                <span className="text-[10.5px] text-slate-500">cm (លំនាំដើម 2.80)</span>
              </div>
            </div>

            <div>
              <label className="block font-semibold text-slate-700 mb-1">
                គម្លាតទំព័រខាងស្តាំ (Right Margin cm) ៖
              </label>
              <div className="flex items-center gap-2">
                <input
                  type="number"
                  min={0.2}
                  max={5.0}
                  step={0.05}
                  value={pageRightMargin}
                  onChange={(e) => {
                    const val = Number(e.target.value) || 1.30;
                    setPageRightMargin(val);
                    setStoredSetting('right_margin', String(val));
                  }}
                  className="w-24 h-8 min-h-[32px] max-h-[32px] box-border px-2.5 py-1 border border-slate-300 rounded bg-white text-xs font-medium focus:ring-1 focus:ring-blue-500 outline-none"
                />
                <span className="text-[10.5px] text-slate-500">cm (លំនាំដើម 1.30)</span>
              </div>
            </div>

            <div>
              <label className="block font-semibold text-slate-700 mb-1">
                គម្លាតទំព័រលើ / ក្រោម (Top / Bottom Margin cm) ៖
              </label>
              <div className="flex items-center gap-3">
                <div className="flex items-center gap-1">
                  <span className="text-[11px] text-slate-600">លើ:</span>
                  <input
                    type="number"
                    min={0.2}
                    max={5.0}
                    step={0.05}
                    value={pageTopMargin}
                    onChange={(e) => {
                      const val = Number(e.target.value) || 1.0;
                      setPageTopMargin(val);
                      setStoredSetting('top_margin', String(val));
                    }}
                    className="w-16 h-8 min-h-[32px] max-h-[32px] box-border px-1.5 py-1 border border-slate-300 rounded bg-white text-xs font-medium focus:ring-1 focus:ring-blue-500 outline-none text-center"
                  />
                </div>
                <div className="flex items-center gap-1">
                  <span className="text-[11px] text-slate-600">ក្រោម:</span>
                  <input
                    type="number"
                    min={0.1}
                    max={5.0}
                    step={0.05}
                    value={pageBottomMargin}
                    onChange={(e) => {
                      const val = Number(e.target.value) || 0.5;
                      setPageBottomMargin(val);
                      setStoredSetting('bottom_margin', String(val));
                    }}
                    className="w-16 h-8 min-h-[32px] max-h-[32px] box-border px-1.5 py-1 border border-slate-300 rounded bg-white text-xs font-medium focus:ring-1 focus:ring-blue-500 outline-none text-center"
                  />
                </div>
                <button
                  type="button"
                  onClick={() => {
                    setPageLeftMargin(1.80);
                    setPageRightMargin(1.0);
                    setPageTopMargin(1.0);
                    setPageBottomMargin(0.5);
                    setStoredSetting('left_margin', '1.80');
                    setStoredSetting('right_margin', '1.0');
                    setStoredSetting('top_margin', '1.0');
                    setStoredSetting('bottom_margin', '0.5');
                  }}
                  className="h-8 px-2 py-0.5 text-[10.5px] rounded bg-slate-100 hover:bg-slate-200 text-slate-700 border border-slate-300 transition-colors cursor-pointer flex items-center"
                  title="កំណត់គម្លាតទំព័រឡើងវិញ៖ ឆ្វេង 1.80cm, ស្តាំ 1.0cm, លើ 1.0cm, ក្រោម 0.5cm"
                >
                  កំណត់ស្តង់ដារ (ឆ្វេង 1.80, ស្តាំ 1, លើ 1, ក្រោម 0.5)
                </button>
              </div>
            </div>

            {/* Header and Title Font Size in pt */}
            <div className="sm:col-span-2 bg-slate-50 border border-slate-200 rounded-lg p-2.5">
              <label className="block font-semibold text-slate-800 mb-1">
                ទំហំពុម្ពអក្សរក្បាលលិខិត និងចំណងជើង (Header & Title Font Size) ៖
              </label>
              <div className="flex items-center gap-2 flex-wrap">
                <input
                  type="number"
                  min={8}
                  max={20}
                  step={0.5}
                  value={headerFontSizePt}
                  onChange={(e) => {
                    const val = parseFloat(e.target.value) || 12;
                    setHeaderFontSizePt(val);
                    setStoredSetting('header_font_size', String(val));
                  }}
                  className="w-20 px-2 py-1 border border-slate-300 rounded bg-white text-xs font-bold text-slate-800 focus:ring-1 focus:ring-blue-500 outline-none font-times"
                />
                <span className="text-[11px] font-bold text-slate-700">pt</span>
                <div className="flex items-center gap-1 ml-2">
                  {[10, 11, 11.5, 12, 12.5, 13, 14].map((sz) => (
                    <button
                      key={sz}
                      type="button"
                      onClick={() => {
                        setHeaderFontSizePt(sz);
                        setStoredSetting('header_font_size', String(sz));
                      }}
                      className={`text-[11px] px-2 py-0.5 rounded border transition cursor-pointer ${
                        headerFontSizePt === sz
                          ? 'bg-[#002060] text-white font-bold border-[#002060]'
                          : 'bg-white text-slate-700 hover:bg-slate-100 border-slate-300'
                      }`}
                    >
                      {sz}pt
                    </button>
                  ))}
                </div>
                <span className="text-[10.5px] text-slate-500 ml-1">(លំនាំដើម 12pt សម្រាប់បឋមកថា និងចំណងជើង)</span>
              </div>
            </div>

            {/* Table Font Size in px */}
            <div className="sm:col-span-2 bg-slate-50 border border-slate-200 rounded-lg p-2.5">
              <label className="block font-semibold text-slate-800 mb-1">
                ទំហំពុម្ពអក្សរក្នុងតារាង និងលទ្ធផល (Table & Results Font Size) ៖
              </label>
              <div className="flex items-center gap-2 flex-wrap">
                <input
                  type="number"
                  min={8}
                  max={20}
                  step={0.5}
                  value={tableFontSize}
                  onChange={(e) => {
                    const val = parseFloat(e.target.value) || 11.5;
                    setTableFontSize(val);
                    setStoredSetting('table_font_size', String(val));
                  }}
                  className="w-20 px-2 py-1 border border-slate-300 rounded bg-white text-xs font-bold text-slate-800 focus:ring-1 focus:ring-blue-500 outline-none font-times"
                />
                <span className="text-[11px] font-bold text-slate-700">px</span>
                <div className="flex items-center gap-1 ml-2">
                  {[10, 11, 11.5, 12, 12.5, 13, 13.5, 14].map((sz) => (
                    <button
                      key={sz}
                      type="button"
                      onClick={() => {
                        setTableFontSize(sz);
                        setStoredSetting('table_font_size', String(sz));
                      }}
                      className={`text-[11px] px-2 py-0.5 rounded border transition cursor-pointer ${
                        tableFontSize === sz
                          ? 'bg-[#002060] text-white font-bold border-[#002060]'
                          : 'bg-white text-slate-700 hover:bg-slate-100 border-slate-300'
                      }`}
                    >
                      {sz}px
                    </button>
                  ))}
                </div>
                <span className="text-[10.5px] text-slate-500 ml-1">(លំនាំដើម 12.5px សម្រាប់ទិន្នន័យលទ្ធផលក្នុងតារាង)</span>
              </div>
            </div>

            {/* Table Row Height in px */}
            <div className="sm:col-span-2 bg-slate-50 border border-slate-200 rounded-lg p-2.5">
              <label className="block font-semibold text-slate-800 mb-1">
                កម្ពស់ជួរតារាង (Table Row Height) ៖
              </label>
              <div className="flex items-center gap-2 flex-wrap">
                <input
                  type="number"
                  min={15}
                  max={35}
                  step={1}
                  value={tableRowHeight}
                  onChange={(e) => {
                    const val = parseInt(e.target.value, 10) || 19;
                    setTableRowHeight(val);
                    setStoredSetting('row_height', String(val));
                  }}
                  className="w-20 px-2 py-1 border border-slate-300 rounded bg-white text-xs font-bold text-slate-800 focus:ring-1 focus:ring-blue-500 outline-none font-times"
                />
                <span className="text-[11px] font-bold text-slate-700">px</span>
                <div className="flex items-center gap-1 ml-2">
                  {[17, 18, 19, 20, 21, 22, 23].map((h) => (
                    <button
                      key={h}
                      type="button"
                      onClick={() => {
                        setTableRowHeight(h);
                        setStoredSetting('row_height', String(h));
                      }}
                      className={`text-[11px] px-2 py-0.5 rounded border transition cursor-pointer ${
                        tableRowHeight === h
                          ? 'bg-[#002060] text-white font-bold border-[#002060]'
                          : 'bg-white text-slate-700 hover:bg-slate-100 border-slate-300'
                      }`}
                    >
                      {h}px
                    </button>
                  ))}
                </div>
                <span className="text-[10.5px] text-slate-500 ml-1">(លំនាំដើម 19px បង្រួមជួរឱ្យសមរម្យសម្រាប់ទំព័រ A4 តែមួយ)</span>
              </div>
            </div>

            <div className="sm:col-span-3 bg-white border border-amber-200 rounded-lg p-3">
              <div className="flex items-center justify-between flex-wrap gap-2 mb-2 pb-2 border-b border-amber-100">
                <label className="block font-semibold text-slate-800 text-xs">
                  រូបភាពគំនូសជាតិតាក់តែង (Tacteing Ornament / Picture) ៖
                </label>
                <div className="flex items-center gap-3 flex-wrap">
                  {/* Master Toggle */}
                  <label className="flex items-center gap-1.5 text-xs font-bold text-amber-900 cursor-pointer select-none">
                    <input
                      type="checkbox"
                      checked={showTacteingPicture}
                      onChange={(e) => {
                        const val = e.target.checked;
                        setShowTacteingPicture(val);
                        setStoredSetting('show_tacteing_picture', String(val));
                      }}
                      className="rounded text-amber-600 focus:ring-amber-500"
                    />
                    <span>{showTacteingPicture ? 'បង្ហាញរូបភាព (Show Picture)' : 'លាក់រូបភាព (Hide Picture)'}</span>
                  </label>

                  {/* Left Side Toggle */}
                  <label className="flex items-center gap-1.5 text-xs cursor-pointer select-none">
                    <input
                      type="checkbox"
                      checked={showLeftTacteing}
                      disabled={!showTacteingPicture}
                      onChange={(e) => {
                        const val = e.target.checked;
                        setShowLeftTacteing(val);
                        setStoredSetting('show_left_tacteing', String(val));
                      }}
                      className="rounded text-blue-600 focus:ring-blue-500 disabled:opacity-50"
                    />
                    <span className={!showTacteingPicture ? 'text-slate-400' : 'text-slate-700'}>
                      ខាងឆ្វេង (ក្រោមឈ្មោះក្រុម/ផ្នែក)
                    </span>
                  </label>

                  {/* Right Side Toggle */}
                  <label className="flex items-center gap-1.5 text-xs cursor-pointer select-none">
                    <input
                      type="checkbox"
                      checked={showRightTacteing}
                      disabled={!showTacteingPicture}
                      onChange={(e) => {
                        const val = e.target.checked;
                        setShowRightTacteing(val);
                        setStoredSetting('show_right_tacteing', String(val));
                      }}
                      className="rounded text-blue-600 focus:ring-blue-500 disabled:opacity-50"
                    />
                    <span className={!showTacteingPicture ? 'text-slate-400' : 'text-slate-700'}>
                      ខាងស្តាំ (ក្រោមជាតិ សាសនា ព្រះមហាក្សត្រ)
                    </span>
                  </label>
                </div>
              </div>

              {showTacteingPicture ? (
                <TacteingControlSelector
                  currentType={tacteingSettings.type}
                  customImage={tacteingSettings.customImage}
                  onChange={(type, customImage) => {
                    setTacteingSettings({ type, customImage: customImage || null });
                    saveTacteingSettings(type, customImage);
                  }}
                />
              ) : (
                <div className="p-3 text-center text-xs text-slate-500 bg-slate-50 rounded border border-dashed border-slate-200">
                  <span>រូបភាពគំនូសជាតិតាក់តែងត្រូវបានលាក់ (Hidden)។ សូមគូសធីក «បង្ហាញរូបភាព» ដើម្បីបង្ហាញឡើងវិញ។</span>
                </div>
              )}
            </div>
          </div>
        )}
      </>
    )}

        {isSavedNotice && (
          <div className="mt-2 text-xs font-medium text-emerald-700 flex items-center gap-1">
            <CheckCircle2 className="w-3.5 h-3.5" />
            <span>បានរក្សាទុកការកែសម្រួលដោយជោគជ័យ!</span>
          </div>
        )}
      </div>

      {/* Printable Sheet Container (Matches A4 layout from user's PDF) */}
      <div className="overflow-x-auto flex justify-center pb-12">
        <div
          id="daily-team-statistics-report-pdf"
          ref={reportContainerRef}
          style={{
            transform: `scale(${zoomScale / 100})`,
            transformOrigin: 'top center',
            paddingLeft: `${pageLeftMargin}cm`,
            paddingRight: `${pageRightMargin}cm`,
            paddingTop: `${pageTopMargin}cm`,
            paddingBottom: `${pageBottomMargin}cm`,
            boxSizing: 'border-box',
          }}
          className="w-[210mm] min-w-[210mm] max-w-[210mm] min-h-[297mm] bg-white border border-slate-300 shadow-md text-black print:border-none print:shadow-none print:m-0 print:w-[210mm] print:transform-none transition-transform box-border"
        >
          {/* Header Section */}
          <div className="flex justify-between items-start mb-2 text-xs">
            {/* Top Left: Ministry and Department info - Aligned so ក្រសួងមហាផ្ទៃ starts level with ជាតិ សាសនា ព្រះមហាក្សត្រ */}
            <div className="text-center flex flex-col items-center leading-snug">
              {/* Invisible spacer matching height of ព្រះរាជាណាចក្រកម្ពុជា so ក្រសួងមហាផ្ទៃ aligns with ជាតិ សាសនា ព្រះមហាក្សត្រ */}
              <p
                className="font-moul whitespace-nowrap invisible select-none pointer-events-none m-0 p-0"
                aria-hidden="true"
                style={{
                  fontSize: `${headerFontSizePt}pt`,
                  lineHeight: 1.25,
                  fontFamily: "'Khmer OS Muol Light', 'Khmer OS Moul Light', 'Khmer OS Muol', 'Moul', serif",
                }}
              >
                ព្រះរាជាណាចក្រកម្ពុជា
              </p>
              <p
                className="font-moul leading-snug tracking-wide text-black whitespace-nowrap m-0 p-0"
                style={{
                  fontSize: `${headerFontSizePt}pt`,
                  fontFamily: "'Khmer OS Muol Light', 'Khmer OS Moul Light', 'Khmer OS Muol', 'Moul', serif",
                }}
              >
                {ministryName}
              </p>
              <p
                className="font-moul leading-snug text-black mt-0.5 whitespace-nowrap"
                style={{
                  fontSize: `${headerFontSizePt}pt`,
                  fontFamily: "'Khmer OS Muol Light', 'Khmer OS Moul Light', 'Khmer OS Muol', 'Moul', serif",
                }}
              >
                {departmentName}
              </p>
              <p
                className="font-moul leading-snug text-black mt-0.5 whitespace-nowrap"
                style={{
                  fontSize: `${headerFontSizePt}pt`,
                  fontFamily: "'Khmer OS Muol Light', 'Khmer OS Moul Light', 'Khmer OS Muol', 'Moul', serif",
                }}
              >
                {generalDeptName}
              </p>
              <p
                className="font-moul leading-snug text-black mt-0.5 whitespace-nowrap"
                style={{
                  fontSize: `${headerFontSizePt}pt`,
                  fontFamily: "'Khmer OS Muol Light', 'Khmer OS Moul Light', 'Khmer OS Muol', 'Moul', serif",
                }}
              >
                {officeName}
              </p>
              {isTotalReport ? (
                <p
                  className={`${
                    teamTitleFont === 'moul'
                      ? 'font-moul'
                      : 'font-siemreap font-bold'
                  } leading-snug text-black mt-0.5 whitespace-nowrap`}
                  style={{
                    fontFamily:
                      teamTitleFont === 'siemreap'
                        ? "'Khmer OS Siemreap', 'Siemreap', sans-serif"
                        : "'Khmer OS Muol Light', 'Khmer OS Moul Light', 'Khmer OS Muol', 'Moul', serif",
                    fontSize: `${teamTitleFontSizePt}pt`,
                  }}
                >
                  {customTeamTitle || 'ផ្នែករដ្ឋបាល'}
                </p>
              ) : (customTeamTitle || formatReportTeamName(selectedTeam)) ? (
                <p
                  className={`${
                    teamTitleFont === 'moul'
                      ? 'font-moul'
                      : 'font-siemreap font-bold'
                  } leading-snug text-black mt-0.5 whitespace-nowrap`}
                  style={{
                    fontFamily:
                      teamTitleFont === 'siemreap'
                        ? "'Khmer OS Siemreap', 'Siemreap', sans-serif"
                        : "'Khmer OS Muol Light', 'Khmer OS Moul Light', 'Khmer OS Muol', 'Moul', serif",
                    fontSize: `${teamTitleFontSizePt}pt`,
                  }}
                >
                  {customTeamTitle || formatReportTeamName(selectedTeam)}
                </p>
              ) : null}
              {showTacteingPicture && showLeftTacteing && (
                <div className="mt-1 flex justify-center">
                  <TacteingLine
                    type={tacteingSettings.type}
                    customImage={tacteingSettings.customImage}
                    width={110}
                    height={10}
                  />
                </div>
              )}
            </div>

            {/* Top Right: Kingdom of Cambodia - Font Moul aligned */}
            <div
              className="text-center flex flex-col items-center leading-snug shrink-0"
              style={{ marginRight: `${rightHeaderMargin}px` }}
            >
              <p
                className="font-moul leading-snug tracking-wider text-black mb-0.5 whitespace-nowrap"
                style={{
                  fontSize: `${headerFontSizePt}pt`,
                  lineHeight: 1.25,
                  fontFamily: "'Khmer OS Muol Light', 'Khmer OS Moul Light', 'Khmer OS Muol', 'Moul', serif",
                }}
              >
                ព្រះរាជាណាចក្រកម្ពុជា
              </p>
              <p
                className="font-moul leading-snug tracking-wider text-black whitespace-nowrap"
                style={{
                  fontSize: `${headerFontSizePt}pt`,
                  lineHeight: 1.25,
                  fontFamily: "'Khmer OS Muol Light', 'Khmer OS Moul Light', 'Khmer OS Muol', 'Moul', serif",
                }}
              >
                ជាតិ សាសនា ព្រះមហាក្សត្រ
              </p>
              {showTacteingPicture && showRightTacteing && (
                <div className="mt-1 flex justify-center">
                  <TacteingLine
                    type={tacteingSettings.type}
                    customImage={tacteingSettings.customImage}
                    width={110}
                    height={10}
                  />
                </div>
              )}
            </div>
          </div>

          {/* Document Title */}
          <div className="text-center my-2">
            <h1
              className="font-moul text-black font-bold tracking-normal mb-1"
              style={{
                fontSize: `${headerFontSizePt}pt`,
                fontFamily: "'Khmer OS Muol Light', 'Khmer OS Moul Light', 'Khmer OS Muol', 'Moul', serif",
              }}
            >
              ស្ថិតិភ្ញៀវស្នើសុំទិដ្ឋាការចូលប្រទេស
            </h1>
            <p
              className="font-bold text-black font-siemreap"
              style={{
                fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif",
                fontSize: `${headerFontSizePt}pt`,
              }}
            >
              {khmerDateRangeSubtitle}
            </p>
          </div>

          {/* Main Table (Exactly 15 columns matching user's PDF) */}
          <div className="w-full overflow-x-auto mt-2">
            <table
              className="w-full border-collapse border border-black text-center select-text"
              style={{ fontSize: `${tableFontSize}px` }}
            >
              <thead>
                <tr
                  className="bg-[#002060] text-white font-bold"
                  style={{ height: `${Math.max(20, tableRowHeight + 2)}px` }}
                >
                  {/* Column 1: Date */}
                  <th
                    className="border border-black px-1 py-0.5 font-siemreap w-[105px] whitespace-nowrap leading-tight"
                    style={{ fontSize: `${tableFontSize}px` }}
                  >
                    ប្រចាំថ្ងៃ
                  </th>

                  {/* Columns 2-14: 13 Visa Types */}
                  {STATS_VISA_TYPES.map((vt) => (
                    <th
                      key={vt}
                      className="border border-black px-0.5 py-0.5 font-times font-bold w-[42px] leading-tight"
                      style={{ fontSize: `${Math.max(10, tableFontSize + 0.5)}px` }}
                    >
                      {vt}
                    </th>
                  ))}

                  {/* Column 15: Total */}
                  <th
                    className="border border-black px-1 py-0.5 font-siemreap w-[58px] whitespace-nowrap leading-tight"
                    style={{ fontSize: `${tableFontSize}px` }}
                  >
                    សរុប
                  </th>
                </tr>
              </thead>

              <tbody>
                {/* 1 Row per Day in selected Month */}
                {Array.from({ length: daysInMonth }, (_, idx) => {
                  const day = idx + 1;
                  const dateLabel = getRowDateLabel(day);
                  const dayVals = dailyData[day] || {};
                  const rowSum = rowTotals[day] || 0;

                  return (
                    <tr
                      key={day}
                      style={{ height: `${tableRowHeight}px` }}
                      className="hover:bg-amber-50/60 transition-colors"
                    >
                      {/* Date Cell */}
                      <td
                        className="border border-black px-1 py-0 text-center font-times text-black whitespace-nowrap leading-none"
                        style={{ fontSize: `${tableFontSize}px` }}
                      >
                        {dateLabel}
                      </td>

                      {/* 13 Visa Values Cells (Editable inline, shows '-' when 0) */}
                      {STATS_VISA_TYPES.map((vt) => {
                        const val = dayVals[vt] || 0;

                        return (
                          <td
                            key={vt}
                            contentEditable
                            suppressContentEditableWarning
                            onFocus={(e) => {
                              if (e.currentTarget.textContent?.trim() === '-') {
                                e.currentTarget.textContent = '';
                              }
                            }}
                            onBlur={(e) =>
                              handleCellChange(day, vt, e.currentTarget.textContent || '0')
                            }
                            className="border border-black px-0.5 py-0 text-center font-times font-normal text-black outline-none hover:bg-yellow-100/70 focus:bg-yellow-100 focus:ring-1 focus:ring-blue-500 leading-none cursor-text"
                            style={{ fontSize: `${tableFontSize}px` }}
                          >
                            {val > 0 ? val : '-'}
                          </td>
                        );
                      })}

                      {/* Row Total Cell */}
                      <td
                        className="border border-black px-1 py-0 text-center font-times font-normal text-black leading-none"
                        style={{ fontSize: `${tableFontSize}px` }}
                      >
                        {rowSum > 0 ? rowSum : '-'}
                      </td>
                    </tr>
                  );
                })}

                {/* Grand Total Bottom Row */}
                <tr
                  className="bg-white font-bold"
                  style={{ height: `${Math.max(20, tableRowHeight + 2)}px` }}
                >
                  <td
                    className="border border-black px-1 py-0.5 text-center font-siemreap font-bold text-black leading-tight"
                    style={{ fontSize: `${tableFontSize}px` }}
                  >
                    សរុប
                  </td>

                  {/* 13 Visa Column Sums */}
                  {STATS_VISA_TYPES.map((vt) => (
                    <td
                      key={vt}
                      className="border border-black px-0.5 py-0.5 text-center font-times font-bold text-black leading-tight"
                      style={{ fontSize: `${tableFontSize}px` }}
                    >
                      {columnTotals[vt] > 0 ? columnTotals[vt] : '-'}
                    </td>
                  ))}

                  {/* Grand Total of All Visas */}
                  <td
                    className="border border-black px-1 py-0.5 text-center font-times font-bold text-black leading-tight"
                    style={{ fontSize: `${tableFontSize}px` }}
                  >
                    {grandTotal > 0 ? grandTotal : '-'}
                  </td>
                </tr>
              </tbody>
            </table>
          </div>

          {/* Footer & Signature Section */}
          <div
            className="flex justify-end font-siemreap text-black text-xs leading-normal relative"
            style={{ marginTop: `${signatureMarginTop}px` }}
          >
            <div className="text-center min-w-[320px]">
              <p
                className="font-normal text-black mb-1 leading-normal"
                style={{
                  fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif",
                  fontSize: '12pt',
                }}
              >
                {customLunarDate}
              </p>
              <p
                className="font-normal text-black mb-1 leading-normal"
                style={{
                  fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif",
                  fontSize: '12pt',
                }}
              >
                {formattedSigningDateKhmer}
              </p>
              <p
                className={`${
                  chiefTitleFont === 'muol'
                    ? 'font-moul leading-snug'
                    : 'font-siemreap font-bold leading-normal'
                } text-black mt-1 mb-14 whitespace-nowrap`}
                style={{
                  fontSize: '12pt',
                  fontFamily:
                    chiefTitleFont === 'muol'
                      ? "'Khmer OS Muol Light', 'Khmer OS Moul Light', 'Khmer OS Muol', 'Moul', serif"
                      : "'Khmer OS Siemreap', 'Siemreap', sans-serif",
                }}
              >
                {chiefTitle}
              </p>
              {chiefName && (
                <p
                  className="font-bold text-black underline underline-offset-4"
                  style={{
                    fontSize: '12pt',
                    fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif",
                  }}
                >
                  {chiefName}
                </p>
              )}
            </div>
          </div>
        </div>
      </div>

      {/* Checklist Audit Modal */}
      <DailyTeamStatsChecklistModal
        isOpen={isChecklistOpen}
        onClose={() => setIsChecklistOpen(false)}
        selectedTeam={selectedTeam}
        selectedYear={selectedYear}
        selectedMonth={selectedMonth}
        daysInMonth={daysInMonth}
        stockRecords={stockRecords}
        excludedIds={excludedIds}
        onApplyExcludedIds={handleApplyExcludedIds}
        allTeams={allTeams}
        categoryFilterProp={selectedOption}
      />

      {/* Print Stylesheet */}
      <style>{`
        @media print {
          @page {
            size: A4 portrait;
            margin: 10mm 8mm 10mm 8mm;
          }
          body {
            background: white !important;
            color: black !important;
            -webkit-print-color-adjust: exact !important;
            print-color-adjust: exact !important;
          }
          .print\\:hidden {
            display: none !important;
          }
          table {
            border-collapse: collapse !important;
          }
          th, td {
            border: 1px solid black !important;
          }
          tr {
            page-break-inside: avoid !important;
          }
        }
      `}</style>
    </div>
  );
};
