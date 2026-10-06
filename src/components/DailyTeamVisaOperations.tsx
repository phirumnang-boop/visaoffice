import React, { useState, useEffect, useMemo, useRef, useCallback, useDeferredValue } from 'react';
import { CategoriesState, StockRecord, UserRole } from '../types';
import { apiService } from '../services/apiService';
import * as XLSX from 'xlsx';
import {
  Calendar,
  Download,
  Printer,
  Layers,
  CheckSquare,
  Square,
  Check,
  Filter,
  Search,
  AlertTriangle,
  Trash2,
  Eye,
  EyeOff,
  RefreshCw,
  X,
  FileCheck,
  CheckCircle2,
  SlidersHorizontal,
  Calculator,
  Info,
  ArrowRight,
  Plus,
  Minus,
  Equal,
  HelpCircle,
  Hash,
  Tag,
  Database,
  Edit3,
  Lock,
  Unlock,
  Clock,
  ShieldCheck,
  Send,
  FileText,
  AlertCircle,
} from 'lucide-react';
import {
  calculateAllTeamsStickerStockAtDate,
  DEFAULT_OPENING_MATRIX,
  isOldStockTeamRecord,
  resolveRecordTeamName,
  isRecordForTeam,
  isRecordForRecipientTeam,
  isTransferTeamRecord,
  isIssueTeamRecord,
  isCeaRecord,
} from '../utils/teamStockCalculation';
import {
  normalizeTeamName,
  normalizeDateToISO,
  normalizeVisaType,
  matchTeamInList,
} from '../utils/teamNormalization';
import { CustomDatePicker } from './CustomDatePicker';

export type VisaCategoryOption = 'Sticker' | 'cEA';

export interface FormRowEntry {
  id: string;
  quantity: string;
  startSerial: string;
  endSerial: string;
  oldCode: string;
}

interface DailyTeamRecord {
  id: string;
  categoryType: VisaCategoryOption; // 'Sticker' or 'cEA'
  date: string;
  teamName: string;
  previousDate?: string;
  values: {
    [visaType: string]: {
      quantity?: number | string;
      startSerial?: string;
      endSerial?: string;
      oldCode?: string;
      entries?: FormRowEntry[];
    };
  };
  totalSheets: number;
  createdAt: string;
  updatedAt?: string;
}

interface DailyTeamVisaOperationsProps {
  categories: CategoriesState;
  stockRecords?: StockRecord[];
  currentRole: UserRole;
  userName: string;
  assignedTeam?: string;
  onAddStockRecord?: (record: StockRecord) => void;
  onBatchImportStockRecords?: (records: StockRecord[]) => void;
  onUpdateStockRecord?: (record: StockRecord) => void;
  onDeleteStockRecord?: (id: string) => void;
  onDeleteBatchStockRecords?: (ids: string[]) => void;
  onShowToast: (msg: string, type: 'success' | 'error') => void;
}

const VISA_TYPES = [
  { id: 'T', name: 'T', isRed: true },
  { id: 'T1', name: 'T1', isRed: false },
  { id: 'T2', name: 'T2', isRed: false },
  { id: 'T3', name: 'T3', isRed: false },
  { id: 'E', name: 'E', isRed: true },
  { id: 'E1', name: 'E1', isRed: false },
  { id: 'E2', name: 'E2', isRed: false },
  { id: 'E3', name: 'E3', isRed: false },
  { id: 'D', name: 'D', isRed: false },
  { id: 'K', name: 'K', isRed: true },
  { id: 'A', name: 'A', isRed: false },
  { id: 'B', name: 'B', isRed: false },
  { id: 'C', name: 'C', isRed: false },
];

const DEFAULT_TEAMS_LIST = [
  'អាកាស តេជោ',
  'អាកាស សៀមរាប',
  'អាកាស ព្រះសីហនុ',
  'ព្រំដែន ប៉ោយប៉ែត',
  'ព្រំដែន បាវិត',
  'ព្រំដែន ចាំយាម',
  'ព្រំដែន ដូង',
  'ព្រំដែន អូរស្មាច់',
  'ព្រំដែន ព្រំ',
  'ព្រំដែន បន្ទាយចក្រី',
  'ព្រំដែន ជាំ',
  'ព្រំដែន ត្រពាំងផ្លុង',
  'ព្រំដែន ត្រពាំងស្រែ',
  'ព្រំដែន ត្រពាំងរូង',
  'ព្រំដែន ភ្នំដិន',
  'ព្រំដែន កោះរកា',
  'ព្រំដែន ព្រែកចាក',
  'ព្រំដែន អូរយ៉ាដាវ',
  'ព្រំដែន ក្អមសំណ',
  'ព្រំដែន ភ្នំដី',
  'កំពង់ផែ ឧកញ៉ាម៉ុង',
  'កំពង់ផែ ស្ទឹងហាវ',
  'កំពង់ផែ ព្រះសីហនុ',
  'កំពង់ផែ ភ្នំពេញ',
  'ព្រំដែន ព្រែកបាក់',
  'ព្រំដែន ម៉ឺនជ័យ',
  'ព្រំដែន ស្ទឹងបត់',
  'កំពង់ផែ កោះកុង',
  'កំពង់ផែ កំពត',
];

// Helper function to increment serial strings accurately with prefix/zero-padding/BigInt support
function incrementSerial(serialStr: string, step: number = 1): string {
  if (!serialStr) return '';
  const trimmed = serialStr.trim();
  const match = trimmed.match(/^(.*?)(\d+)(.*?)$/);
  if (!match) return trimmed;
  const prefix = match[1];
  const digits = match[2];
  const suffix = match[3];

  try {
    const nextVal = BigInt(digits) + BigInt(step);
    const nextDigits = nextVal.toString().padStart(digits.length, '0');
    return `${prefix}${nextDigits}${suffix}`;
  } catch {
    const n = parseInt(digits, 10);
    if (isNaN(n)) return trimmed;
    const nextDigits = String(n + step).padStart(digits.length, '0');
    return `${prefix}${nextDigits}${suffix}`;
  }
}

// Helper function to decrement serial strings accurately
function decrementSerial(serialStr: string, step: number = 1): string {
  if (!serialStr) return '';
  const trimmed = serialStr.trim();
  const match = trimmed.match(/^(.*?)(\d+)(.*?)$/);
  if (!match) return trimmed;
  const prefix = match[1];
  const digits = match[2];
  const suffix = match[3];

  try {
    const val = BigInt(digits);
    const stepVal = BigInt(step);
    if (val <= stepVal) return trimmed;
    const prevVal = val - stepVal;
    const prevDigits = prevVal.toString().padStart(digits.length, '0');
    return `${prefix}${prevDigits}${suffix}`;
  } catch {
    const n = parseInt(digits, 10);
    if (isNaN(n) || n <= step) return trimmed;
    const prevDigits = String(n - step).padStart(digits.length, '0');
    return `${prefix}${prevDigits}${suffix}`;
  }
}

export const DailyTeamVisaOperations: React.FC<DailyTeamVisaOperationsProps> = ({
  categories,
  stockRecords = [],
  currentRole,
  userName,
  assignedTeam,
  onAddStockRecord,
  onBatchImportStockRecords,
  onUpdateStockRecord,
  onDeleteStockRecord,
  onDeleteBatchStockRecords,
  onShowToast,
}) => {
  // Option Button: Sticker vs cEA
  const [selectedOption, setSelectedOption] = useState<VisaCategoryOption>(() => {
    return (localStorage.getItem('app_daily_team_selected_option') as VisaCategoryOption) || 'Sticker';
  });

  // Team options (សន្លឹកទិដ្ឋាការ / ក្រដាសអនុម័ត) from CategoryManager ("ទិន្នន័យក្រុមផ្តល់ទិដ្ឋាការ")
  const [teamTypeOptions, setTeamTypeOptions] = useState<Record<string | number, { sticker: boolean; evisa: boolean }>>(() => {
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

  // Date states - Persist selected date across menu navigation
  const [selectedDate, setSelectedDate] = useState<string>(() => {
    const saved = localStorage.getItem('app_daily_team_selected_date');
    if (saved) return saved;
    return '2026-08-15';
  });

  // Daily operations database in LocalStorage
  const [records, setRecords] = useState<DailyTeamRecord[]>([]);

  // Fetch initial data from API
  useEffect(() => {
    apiService.getDailyTeamOperations().then(setRecords);
  }, []);

  // Sync to localStorage only for caching purposes
  useEffect(() => {
    localStorage.setItem('app_daily_team_operations_v5', JSON.stringify(records));
  }, [records]);

  // Team list strictly taken from categories (ក្រុមផ្តល់ទិដ្ឋាការ.របក) along with default teams
  const allTeams = useMemo(() => {
    const list: string[] = [];
    const seen = new Set<string>();

    const addTeam = (raw: string) => {
      if (!raw) return;
      const trimmed = raw.trim();
      if (!trimmed) return;
      const norm = normalizeTeamName(trimmed);
      if (!seen.has(norm)) {
        seen.add(norm);
        list.push(trimmed);
      }
    };

    if (categories?.visaTeamsRobok && categories.visaTeamsRobok.length > 0) {
      categories.visaTeamsRobok.forEach((t) => addTeam(t.name));
    } else {
      DEFAULT_TEAMS_LIST.forEach(addTeam);
    }

    return list;
  }, [categories?.visaTeamsRobok]);

  const [selectedTeam, setSelectedTeam] = useState<string>(() => {
    const savedTeam = localStorage.getItem('app_daily_team_selected_team');
    if (savedTeam && (currentRole !== 'User' || !assignedTeam || savedTeam === assignedTeam)) {
      return savedTeam;
    }
    if (currentRole === 'User' && assignedTeam) {
      return assignedTeam;
    }
    return savedTeam || '';
  });

  useEffect(() => {
    if (records) {
      localStorage.setItem('app_daily_team_operations_v5', JSON.stringify(records));
    }
  }, [records]);

  // Sync saved selections from localStorage on mount & when custom selection event fires
  useEffect(() => {
    const resyncFromStorage = () => {
      // First priority: Check if direct edit payload is provided
      const payloadStr = localStorage.getItem('app_daily_team_edit_payload');
      if (payloadStr) {
        try {
          const payload = JSON.parse(payloadStr);
          if (payload) {
            const pTeam = payload.targetTeamName || payload.visaTeamRobokName || payload.recipientTeamName || payload.sourceFrom || '';
            const pDate = normalizeDateToISO(payload.date || '');
            const pOption: VisaCategoryOption = payload.optionMode || (payload.sourceFrom === 'cEA' || payload.remarks?.includes('cEA') || (payload.visaType || '').includes('cEA') ? 'cEA' : 'Sticker');
            if (pTeam) {
              if (currentRole !== 'User' || !assignedTeam || pTeam === assignedTeam) {
                setSelectedTeam(pTeam);
              }
            }
            if (pDate) setSelectedDate(pDate);
            if (pOption) setSelectedOption(pOption);
            if (payload.id) setEditingRecordId(payload.id);

            setSelectedVisaTypeFilter('ALL');

            const normTeam = normalizeTeamName(pTeam);
            // 1. Check if DailyTeamRecord exists with all types
            const dtrMatch = (records || []).find((r) => {
              return (
                normalizeDateToISO(r.date) === pDate &&
                normalizeTeamName(r.teamName) === normTeam &&
                (r.categoryType === pOption || (!r.categoryType && pOption === 'Sticker'))
              );
            }) || (effectiveRecords || []).find((r) => {
              return (
                normalizeDateToISO(r.date) === pDate &&
                normalizeTeamName(r.teamName) === normTeam &&
                (r.categoryType === pOption || (!r.categoryType && pOption === 'Sticker'))
              );
            });

            // 2. Find ALL matching stockRecords in "ទិន្នន័យសន្លឹកទិដ្ឋាការ" for this team, date, and option
            const matchingStock = (stockRecords || []).filter((sr) => {
              const srDate = normalizeDateToISO(sr.date || '');
              const srTeam = normalizeTeamName(resolveRecordTeamName(sr) || sr.visaTeamRobokName || (sr as any).teamName || '');
              const isCea = sr.sourceFrom === 'cEA' || sr.remarks?.includes('cEA') || sr.visaType?.includes('cEA');
              const isOptMatch = pOption === 'cEA' ? isCea : !isCea;
              return srDate === pDate && srTeam === normTeam && isOptMatch;
            });

            const updated: Record<string, FormRowEntry[]> = {};
            VISA_TYPES.forEach((vt) => {
              const lastCode = typeof getLastUsedOldCode === 'function' ? getLastUsedOldCode(pTeam, vt.id, pDate, pOption) : '';

              if (dtrMatch?.values?.[vt.id]) {
                const dtrVal = dtrMatch.values[vt.id];
                if (Array.isArray(dtrVal.entries) && dtrVal.entries.length > 0) {
                  updated[vt.id] = dtrVal.entries.map((e: any, idx: number) => ({
                    id: e.id || `${vt.id}-${idx}`,
                    quantity: e.quantity !== undefined && e.quantity !== '' && Number(e.quantity) > 0 ? String(e.quantity) : '',
                    startSerial: e.startSerial || '',
                    endSerial: e.endSerial || '',
                    oldCode: e.oldCode || (idx === 0 ? lastCode : ''),
                  }));
                  return;
                } else if (dtrVal.quantity || dtrVal.startSerial) {
                  const q = dtrVal.quantity !== undefined && dtrVal.quantity !== '' && Number(dtrVal.quantity) > 0 ? String(dtrVal.quantity) : '';
                  updated[vt.id] = [
                    {
                      id: `${vt.id}-0`,
                      quantity: q,
                      startSerial: dtrVal.startSerial || '',
                      endSerial: dtrVal.endSerial || '',
                      oldCode: dtrVal.oldCode || lastCode || '',
                    },
                  ];
                  return;
                }
              }

              const stockForVt = matchingStock.filter((sr) => normalizeVisaType(sr.visaType) === vt.id);
              if (stockForVt.length > 0) {
                updated[vt.id] = stockForVt.map((sr, idx) => ({
                  id: `${vt.id}-${idx}`,
                  quantity: String(sr.totalSheets || (sr.quantityBundles ? Number(sr.quantityBundles) * 50 : '')),
                  startSerial: sr.startSerial || '',
                  endSerial: sr.endSerial || '',
                  oldCode: (sr as any).oldCode || (idx === 0 ? lastCode : ''),
                }));
              } else if (normalizeVisaType(payload.visaType) === vt.id) {
                const q = String(payload.totalSheets || (payload.quantityBundles ? Number(payload.quantityBundles) * 50 : ''));
                updated[vt.id] = [
                  {
                    id: `${vt.id}-0`,
                    quantity: q,
                    startSerial: payload.startSerial || '',
                    endSerial: payload.endSerial || '',
                    oldCode: payload.oldCode || lastCode || '',
                  },
                ];
              } else {
                updated[vt.id] = [
                  {
                    id: `${vt.id}-0`,
                    quantity: '',
                    startSerial: '',
                    endSerial: '',
                    oldCode: lastCode || '',
                  },
                ];
              }
            });
            setFormData(updated);
            localStorage.removeItem('app_daily_team_edit_payload');
            localStorage.removeItem('app_daily_team_edit_record_id');
            localStorage.removeItem('app_daily_team_auto_load_form');
            return;
          }
        } catch (e) {
          console.error(e);
        }
      }

      const savedTeam = localStorage.getItem('app_daily_team_selected_team');
      const savedDate = localStorage.getItem('app_daily_team_selected_date');
      const savedOption = localStorage.getItem('app_daily_team_selected_option') as VisaCategoryOption;

      if (savedTeam) {
        if (currentRole !== 'User' || !assignedTeam || savedTeam === assignedTeam) {
          setSelectedTeam(savedTeam);
        }
      }
      if (savedDate) {
        setSelectedDate(savedDate);
      }
      if (savedOption && (savedOption === 'Sticker' || savedOption === 'cEA')) {
        setSelectedOption(savedOption);
      }

      // Fallback: Attempt loading existing record if edit request ID is pending
      const editId = localStorage.getItem('app_daily_team_edit_record_id');
      const autoLoad = localStorage.getItem('app_daily_team_auto_load_form') === 'true';
      if (editId && autoLoad) {
        const teamToUse = savedTeam || selectedTeam;
        const dateToUse = savedDate || selectedDate;
        const optToUse = savedOption || selectedOption;
        if (teamToUse && dateToUse && optToUse) {
          const loaded = loadExistingRecordIntoForm(teamToUse, dateToUse, optToUse);
          if (loaded) {
            localStorage.removeItem('app_daily_team_auto_load_form');
            localStorage.removeItem('app_daily_team_edit_record_id');
          }
        }
      }
    };

    resyncFromStorage();

    window.addEventListener('app_daily_team_selection_updated', resyncFromStorage);
    return () => {
      window.removeEventListener('app_daily_team_selection_updated', resyncFromStorage);
    };
  }, [currentRole, assignedTeam]);

  // Helper to check if a specific team has "ក្រដាសអនុម័ត" (cEA / evisa) enabled in "ទិន្នន័យក្រុមផ្តល់ទិដ្ឋាការ" ក្រោម header "ក្រុមផ្តល់ទិដ្ឋាការ.របក"
  const isTeamCeaEnabled = useCallback(
    (teamName: string, idx?: number): boolean => {
      if (!teamName) return false;
      const cleanTeam = teamName.trim();
      const cleanLower = cleanTeam.toLowerCase();
      const normTeam = normalizeTeamName(teamName);

      const vtrList =
        categories?.visaTeamsRobok && categories.visaTeamsRobok.length > 0
          ? categories.visaTeamsRobok
          : DEFAULT_TEAMS_LIST.map((name, i) => ({ id: `vtr-${i + 1}`, name, createdAt: '' }));

      let resolvedIdx = idx !== undefined && idx >= 0 ? idx : -1;
      if (resolvedIdx === -1) {
        resolvedIdx = vtrList.findIndex(
          (t) =>
            t.name.trim().toLowerCase() === cleanLower ||
            (normTeam && normalizeTeamName(t.name) === normTeam)
        );
      }

      // If team is not under header ក្រុមផ្តល់ទិដ្ឋាការ.របក, strictly reject ("កុំយកអ្វីក្រៅពីក្នុង")
      if (resolvedIdx === -1) {
        return false;
      }

      const vtrItem = vtrList[resolvedIdx];

      // 1. Check direct ID, Name or Index in teamTypeOptions
      if (vtrItem?.id && teamTypeOptions[vtrItem.id]?.evisa !== undefined) {
        return Boolean(teamTypeOptions[vtrItem.id].evisa);
      }
      if (vtrItem?.name && teamTypeOptions[vtrItem.name]?.evisa !== undefined) {
        return Boolean(teamTypeOptions[vtrItem.name].evisa);
      }
      if (teamTypeOptions[resolvedIdx]?.evisa !== undefined) {
        return Boolean(teamTypeOptions[resolvedIdx].evisa);
      }

      // 2. Check in teamTypeOptions entries by key name
      for (const [key, val] of Object.entries(
        teamTypeOptions as Record<string, { sticker?: boolean; evisa?: boolean }>
      )) {
        const keyStr = String(key).trim();
        if (
          keyStr.toLowerCase() === cleanLower ||
          (normTeam && normalizeTeamName(keyStr) === normTeam)
        ) {
          if (val?.evisa !== undefined) {
            return Boolean(val.evisa);
          }
        }
      }

      // 3. Check corresponding full team name in visaTeams
      const vtList = categories?.visaTeams || [];
      if (resolvedIdx < vtList.length) {
        const vtItem = vtList[resolvedIdx];
        if (vtItem?.id && teamTypeOptions[vtItem.id]?.evisa !== undefined) {
          return Boolean(teamTypeOptions[vtItem.id].evisa);
        }
        if (vtItem?.name && teamTypeOptions[vtItem.name]?.evisa !== undefined) {
          return Boolean(teamTypeOptions[vtItem.name].evisa);
        }
      }

      // 4. Default fallback: if no stored options exist at all in localStorage, default 4 key airport/port teams
      const hasStoredOptions = Object.keys(teamTypeOptions).length > 0;
      if (!hasStoredOptions) {
        const defKeyTeams = [
          'អាកាស តេជោ',
          'អាកាស សៀមរាប',
          'អាកាស ព្រះសីហនុ',
          'កំពង់ផែ ព្រះសីហនុ',
        ];
        return defKeyTeams.some((k) => normalizeTeamName(k) === normTeam);
      }

      return false;
    },
    [teamTypeOptions, categories?.visaTeamsRobok, categories?.visaTeams]
  );

  // Helper to check if a specific team has "សន្លឹកទិដ្ឋាការ" (Sticker) enabled in "ទិន្នន័យក្រុមផ្តល់ទិដ្ឋាការ" ក្រោម header "ក្រុមផ្តល់ទិដ្ឋាការ.របក"
  const isTeamStickerEnabled = useCallback(
    (teamName: string, idx?: number): boolean => {
      if (!teamName) return false;
      const cleanTeam = teamName.trim();
      const cleanLower = cleanTeam.toLowerCase();
      const normTeam = normalizeTeamName(teamName);

      const vtrList =
        categories?.visaTeamsRobok && categories.visaTeamsRobok.length > 0
          ? categories.visaTeamsRobok
          : DEFAULT_TEAMS_LIST.map((name, i) => ({ id: `vtr-${i + 1}`, name, createdAt: '' }));

      let resolvedIdx = idx !== undefined && idx >= 0 ? idx : -1;
      if (resolvedIdx === -1) {
        resolvedIdx = vtrList.findIndex(
          (t) =>
            t.name.trim().toLowerCase() === cleanLower ||
            (normTeam && normalizeTeamName(t.name) === normTeam)
        );
      }

      // If team is NOT under header ក្រុមផ្តល់ទិដ្ឋាការ.របក, strictly reject ("កុំយកអ្វីក្រៅពីក្នុង")
      if (resolvedIdx === -1) {
        return false;
      }

      const vtrItem = vtrList[resolvedIdx];

      // 1. Direct ID, Name, or Index lookup in teamTypeOptions
      if (vtrItem?.id && teamTypeOptions[vtrItem.id]?.sticker !== undefined) {
        return Boolean(teamTypeOptions[vtrItem.id].sticker);
      }
      if (vtrItem?.name && teamTypeOptions[vtrItem.name]?.sticker !== undefined) {
        return Boolean(teamTypeOptions[vtrItem.name].sticker);
      }
      if (teamTypeOptions[resolvedIdx]?.sticker !== undefined) {
        return Boolean(teamTypeOptions[resolvedIdx].sticker);
      }
      if (teamTypeOptions[String(resolvedIdx)]?.sticker !== undefined) {
        return Boolean(teamTypeOptions[String(resolvedIdx)].sticker);
      }

      // 2. Check direct name or normalized name in teamTypeOptions entries
      for (const [key, val] of Object.entries(
        teamTypeOptions as Record<string, { sticker?: boolean; evisa?: boolean }>
      )) {
        const keyStr = String(key).trim();
        if (
          keyStr.toLowerCase() === cleanLower ||
          (normTeam && normalizeTeamName(keyStr) === normTeam)
        ) {
          if (val?.sticker !== undefined) {
            return Boolean(val.sticker);
          }
        }
      }

      // 3. Check corresponding full team name in visaTeams
      const vtList = categories?.visaTeams || [];
      if (resolvedIdx < vtList.length) {
        const vtItem = vtList[resolvedIdx];
        if (vtItem?.id && teamTypeOptions[vtItem.id]?.sticker !== undefined) {
          return Boolean(teamTypeOptions[vtItem.id].sticker);
        }
        if (vtItem?.name && teamTypeOptions[vtItem.name]?.sticker !== undefined) {
          return Boolean(teamTypeOptions[vtItem.name].sticker);
        }
      }

      // 4. If stored options exist, reject any team not explicitly marked sticker: true
      const hasStoredOptions = Object.keys(teamTypeOptions).length > 0;
      if (hasStoredOptions) {
        return false;
      }

      // 5. Default fallback: for valid teams under ក្រុមផ្តល់ទិដ្ឋាការ.របក when no options stored yet
      return true;
    },
    [teamTypeOptions, categories?.visaTeamsRobok, categories?.visaTeams]
  );

  // Check if Paper Approval (ក្រដាសអនុម័ត / cEA) is allowed for the team
  const isTeamEVisaAllowed = useMemo(() => {
    if (currentRole !== 'User') return true;
    const targetTeam = selectedTeam || assignedTeam;
    if (!targetTeam) return false;
    return isTeamCeaEnabled(targetTeam);
  }, [currentRole, selectedTeam, assignedTeam, isTeamCeaEnabled]);

  // Available teams for combobox (ក្រុម) depending on selectedOption (cEA vs Sticker)
  // Strictly shows teams under header ក្រុមផ្តល់ទិដ្ឋាការ.របក in "ទិន្នន័យក្រុមផ្តល់ទិដ្ឋាការ" ("កុំយកអ្វីក្រៅពីក្នុង")
  const availableTeamsForSelection = useMemo(() => {
    const rawRobokTeams =
      categories?.visaTeamsRobok && categories.visaTeamsRobok.length > 0
        ? categories.visaTeamsRobok
        : DEFAULT_TEAMS_LIST.map((name, i) => ({ id: `vtr-${i + 1}`, name, createdAt: '' }));

    if (selectedOption === 'cEA') {
      // In cEA mode, ONLY show teams that have (ក្រដាសអនុម័ត) clicked in "ទិន្នន័យក្រុមផ្តល់ទិដ្ឋាការ" ក្រោម header ក្រុមផ្តល់ទិដ្ឋាការ.របក
      return rawRobokTeams
        .filter((t, idx) => isTeamCeaEnabled(t.name, idx))
        .map((t) => t.name);
    }
    // In Sticker mode, ONLY show teams that have (សន្លឹកទិដ្ឋាការ) clicked in "ទិន្នន័យក្រុមផ្តល់ទិដ្ឋាការ" ក្រោម header ក្រុមផ្តល់ទិដ្ឋាការ.របក
    return rawRobokTeams
      .filter((t, idx) => isTeamStickerEnabled(t.name, idx))
      .map((t) => t.name);
  }, [categories?.visaTeamsRobok, selectedOption, isTeamCeaEnabled, isTeamStickerEnabled]);

  // Auto-lock and update selectedTeam when assignedTeam or currentRole changes
  useEffect(() => {
    if (currentRole === 'User' && assignedTeam) {
      setSelectedTeam(assignedTeam);
    }
  }, [currentRole, assignedTeam]);

  // Persist user selection changes to LocalStorage so returning from another menu keeps the exact date & option
  useEffect(() => {
    localStorage.setItem('app_daily_team_selected_date', selectedDate);
  }, [selectedDate]);

  useEffect(() => {
    localStorage.setItem('app_daily_team_selected_option', selectedOption);
  }, [selectedOption]);

  useEffect(() => {
    if (selectedTeam) {
      localStorage.setItem('app_daily_team_selected_team', selectedTeam);
    } else {
      localStorage.removeItem('app_daily_team_selected_team');
    }
  }, [selectedTeam]);

  // Calculate previous date safely without UTC timezone shifting
  const previousDateString = useMemo(() => {
    if (!selectedDate) return '';
    const norm = normalizeDateToISO(selectedDate);
    if (!norm) return '';
    const [y, m, d] = norm.split('-').map(Number);
    if (!y || !m || !d) return '';
    const dateObj = new Date(y, m - 1, d);
    dateObj.setDate(dateObj.getDate() - 1);
    const prevY = dateObj.getFullYear();
    const prevM = String(dateObj.getMonth() + 1).padStart(2, '0');
    const prevD = String(dateObj.getDate()).padStart(2, '0');
    return `${prevY}-${prevM}-${prevD}`;
  }, [selectedDate]);

  // Initial fetch from Cloud SQL
  useEffect(() => {
    let isMounted = true;
    const fetchCloudDaily = async () => {
      try {
        const cloudDaily = await apiService.getDailyTeamOperations();
        if (isMounted) {
          setRecords(cloudDaily);
        }
      } catch (err) {
        console.warn('DailyTeam Cloud SQL fetch notice:', err);
      }
    };
    fetchCloudDaily();
    return () => {
      isMounted = false;
    };
  }, []);

  // Save records to LocalStorage
  useEffect(() => {
    localStorage.setItem('app_daily_team_operations_v5', JSON.stringify(records));
  }, [records]);

  // Excluded record IDs from Checklist (so user can toggle off any overcounted / duplicate entries)
  const [excludedIds, setExcludedIds] = useState<string[]>(() => {
    try {
      const saved = localStorage.getItem('app_daily_team_excluded_ids');
      if (saved) return JSON.parse(saved);
    } catch {}
    return [];
  });

  useEffect(() => {
    localStorage.setItem('app_daily_team_excluded_ids', JSON.stringify(excludedIds));
  }, [excludedIds]);

  // Modal and audit filter states
  const [isAuditChecklistOpen, setIsAuditChecklistOpen] = useState<boolean>(false);
  const [auditSearch, setAuditSearch] = useState<string>('');
  const deferredAuditSearch = useDeferredValue(auditSearch);
  const [auditDateFilter, setAuditDateFilter] = useState<'current' | 'all' | 'custom'>('current');
  const [auditCustomDate, setAuditCustomDate] = useState<string>('');
  const [auditCategoryFilter, setAuditCategoryFilter] = useState<'all' | 'Sticker' | 'cEA'>('all');
  const [auditTeamFilter, setAuditTeamFilter] = useState<string>('all');
  const [auditShowOnlyOvercounted, setAuditShowOnlyOvercounted] = useState<boolean>(false);

  // Visa Type Quick Filter/Selector and Multi-batch Editing State
  const [selectedVisaTypeFilter, setSelectedVisaTypeFilter] = useState<string>('ALL');
  const [editingRecordId, setEditingRecordId] = useState<string | null>(() => {
    try {
      const payloadStr = localStorage.getItem('app_daily_team_edit_payload');
      if (payloadStr) {
        const payload = JSON.parse(payloadStr);
        if (payload?.id) return payload.id;
      }
    } catch {}
    return localStorage.getItem('app_daily_team_edit_record_id') || null;
  });

  // Delete & Update Confirmation Modal States
  const [deleteModalOpen, setDeleteModalOpen] = useState<boolean>(false);
  const [updateModalOpen, setUpdateModalOpen] = useState<boolean>(false);

  // Drill-down Detail Modal State (e.g. clicking on T1)
  const [drillDownVisa, setDrillDownVisa] = useState<{
    teamName: string;
    visaType: string; // 'T', 'T1', ... or 'TOTAL'
  } | null>(null);

  // Merge stockRecords (ទិន្នន័យសន្លឹកទិដ្ឋាការ) with DailyTeam records so tables can show all data seamlessly
  // Supports aggregating multiple entries per day for the same team (e.g., multiple entries of T or B on the same day)
  // Ensures records entered via form are not double-counted with their linked stock ledger entries
  const effectiveRecords = useMemo(() => {
    const map = new Map<string, DailyTeamRecord>();
    const excludedSet = new Set(excludedIds);

    // 1. Process explicit DailyTeam records saved in DailyTeamVisaOperations first (authoritative source)
    records.forEach((r) => {
      if (excludedSet.has(r.id)) return;
      const rDate = normalizeDateToISO(r.date || '');
      const rNormTeam = normalizeTeamName(r.teamName || '');
      const key = `${r.categoryType}_${rDate}_${rNormTeam}`;
      const existing = map.get(key);

      if (!existing) {
        // Deep copy values
        const copyValues: Record<string, { quantity?: number | string; startSerial?: string; endSerial?: string; oldCode?: string; entries?: FormRowEntry[] }> = {};
        Object.entries(r.values || {}).forEach(([vt, v]) => {
          if (v && typeof v === 'object') {
            const valObj = v as any;
            const subEntries: FormRowEntry[] = Array.isArray(valObj.entries) && valObj.entries.length > 0
              ? valObj.entries.map((e: any, idx: number) => ({
                  id: e.id || `${vt}-${idx}`,
                  quantity: e.quantity !== undefined && e.quantity !== '' ? String(e.quantity) : '',
                  startSerial: e.startSerial || '',
                  endSerial: e.endSerial || '',
                  oldCode: e.oldCode || '',
                }))
              : (valObj.quantity || valObj.startSerial)
              ? [
                  {
                    id: `${vt}-0`,
                    quantity: valObj.quantity !== undefined && valObj.quantity !== '' ? String(valObj.quantity) : '',
                    startSerial: valObj.startSerial || '',
                    endSerial: valObj.endSerial || '',
                    oldCode: valObj.oldCode || '',
                  },
                ]
              : [];
            copyValues[vt] = {
              ...valObj,
              entries: subEntries.length > 0 ? subEntries : undefined,
            };
          }
        });
        map.set(key, {
          ...r,
          date: rDate,
          values: copyValues,
        });
      } else {
        // Accumulate multiple distinct batch entries for the same day/team/category
        const mergedValues: Record<string, { quantity?: number | string; startSerial?: string; endSerial?: string; oldCode?: string; entries?: FormRowEntry[] }> = { ...existing.values };
        let newTotal = 0;

        VISA_TYPES.forEach((vt) => {
          const valObj = (r.values as any)?.[vt.id] as { quantity?: number | string; startSerial?: string; endSerial?: string; oldCode?: string; entries?: FormRowEntry[] } | undefined;
          const curVal = mergedValues[vt.id] || { quantity: 0, startSerial: '', endSerial: '', oldCode: '', entries: [] };
          const existingEntries: FormRowEntry[] = Array.isArray(curVal.entries) ? [...curVal.entries] : [];
          
          const incomingEntries: FormRowEntry[] = Array.isArray(valObj?.entries) && valObj.entries.length > 0
            ? valObj.entries
            : (valObj?.quantity || valObj?.startSerial)
            ? [
                {
                  id: `${vt.id}-${existingEntries.length}`,
                  quantity: String(valObj?.quantity || ''),
                  startSerial: valObj?.startSerial || '',
                  endSerial: valObj?.endSerial || '',
                  oldCode: valObj?.oldCode || '',
                },
              ]
            : [];

          const combinedEntries = [...existingEntries];
          incomingEntries.forEach((ie) => {
            const isDupe = combinedEntries.some(
              (ce) => ce.startSerial === ie.startSerial && ce.endSerial === ie.endSerial && String(ce.quantity) === String(ie.quantity)
            );
            if (!isDupe) {
              combinedEntries.push(ie);
            }
          });

          const totalQtyForVt = combinedEntries.reduce((sum, e) => sum + (Number(e.quantity) || 0), 0);
          if (totalQtyForVt > 0 || combinedEntries.length > 0) {
            mergedValues[vt.id] = {
              quantity: totalQtyForVt,
              startSerial: curVal.startSerial || valObj?.startSerial || '',
              endSerial: valObj?.endSerial || curVal.endSerial || '',
              oldCode: curVal.oldCode || valObj?.oldCode || '',
              entries: combinedEntries.length > 0 ? combinedEntries : undefined,
            };
            newTotal += totalQtyForVt;
          }
        });

        existing.values = mergedValues;
        existing.totalSheets = newTotal;
      }
    });

    // 2. Next incorporate external stockRecords (Excel imports, ledger entries), excluding any that were auto-created from DailyTeam form
    if (stockRecords && stockRecords.length > 0) {
      stockRecords.forEach((sr) => {
        // Skip records that originated from this Daily Team Operations form to avoid double counting
        if (
          sr.id.startsWith('stock-dtr-') ||
          sr.id.startsWith('dtr-') ||
          (sr as any).createdBy === 'DailyTeamVisaOperations'
        ) {
          return;
        }

        if (
          sr.stockType === 'sticker' &&
          (sr.operationType === 'useTeam' || sr.sourceFrom?.includes('ប្រើប្រាស់') || sr.operationType === 'transferTeam')
        ) {
          const rawTeam = resolveRecordTeamName(sr) || (sr.visaTeamRobokName || '').trim();
          const team = rawTeam.trim();
          const normTeam = normalizeTeamName(team);
          const rawDate = sr.date || '';
          const date = normalizeDateToISO(rawDate);
          if (!normTeam || !date) return;

          const cat: VisaCategoryOption =
            sr.sourceFrom === 'cEA' ||
            sr.remarks?.includes('cEA') ||
            (sr.visaType && sr.visaType.includes('cEA'))
              ? 'cEA'
              : 'Sticker';

          const recordId = `dtr-synced-${cat}-${date}-${normTeam.replace(/\s+/g, '-')}`;
          if (excludedSet.has(recordId) || excludedSet.has(sr.id)) return;

          const key = `${cat}_${date}_${normTeam}`;
          const existing = map.get(key);

          const vtId = normalizeVisaType(sr.visaType);
          const sheets = Number(sr.totalSheets) || (Number(sr.quantityBundles) ? Number(sr.quantityBundles) * 50 : 0);

          if (!existing) {
            const newValues: Record<string, any> = {};
            if (vtId && sheets > 0) {
              newValues[vtId] = {
                quantity: sheets,
                startSerial: sr.startSerial || '',
                endSerial: sr.endSerial || '',
                oldCode: '',
                entries: [
                  {
                    id: `${vtId}-0`,
                    quantity: String(sheets),
                    startSerial: sr.startSerial || '',
                    endSerial: sr.endSerial || '',
                    oldCode: '',
                  },
                ],
              };
            }
            map.set(key, {
              id: recordId,
              categoryType: cat,
              date,
              teamName: team,
              values: newValues,
              totalSheets: sheets,
              createdAt: sr.createdAt || new Date().toISOString(),
            });
          } else if (vtId && sheets > 0) {
            // Check if this specific serial/quantity entry is already present in existing record
            const curVal = existing.values[vtId] || { quantity: 0, startSerial: '', endSerial: '', oldCode: '', entries: [] };
            const subEntries: FormRowEntry[] = Array.isArray(curVal.entries) ? [...curVal.entries] : [];

            const isAlreadyIncluded = subEntries.some(
              (e) =>
                (e.startSerial && sr.startSerial && e.startSerial === sr.startSerial && e.endSerial === sr.endSerial) ||
                (Number(e.quantity) === sheets && !e.startSerial && !sr.startSerial)
            );

            if (!isAlreadyIncluded) {
              subEntries.push({
                id: `${vtId}-${subEntries.length}`,
                quantity: String(sheets),
                startSerial: sr.startSerial || '',
                endSerial: sr.endSerial || '',
                oldCode: '',
              });

              const newQty = (Number(curVal.quantity) || 0) + sheets;
              existing.values[vtId] = {
                quantity: newQty,
                startSerial: curVal.startSerial || sr.startSerial || '',
                endSerial: sr.endSerial || curVal.endSerial || '',
                oldCode: curVal.oldCode || '',
                entries: subEntries,
              };
              existing.totalSheets += sheets;
            }
          }
        }
      });
    }

    return Array.from(map.values());
  }, [records, stockRecords, excludedIds]);

  // All individual Raw Records for Audit & Checklist inspection - Calculated only when modal is open for maximum smoothness
  const auditAllEntries = useMemo(() => {
    if (!isAuditChecklistOpen) return [];

    const list: Array<{
      id: string;
      categoryType: VisaCategoryOption;
      date: string;
      teamName: string;
      operationTitle: string;
      operationType: string;
      values: { [vt: string]: { quantity?: number | string; startSerial?: string; endSerial?: string } };
      totalSheets: number;
      source: 'daily_form' | 'stock_ledger';
      isDuplicateOrOvercounted: boolean;
      overcountReason?: string;
    }> = [];

    const keyOccurrences = new Map<string, number>();

    // 1. Add daily records
    records.forEach((r) => {
      const rDate = normalizeDateToISO(r.date || '');
      const key = `${r.categoryType}_use_${rDate}_${normalizeTeamName(r.teamName)}`;
      const count = (keyOccurrences.get(key) || 0) + 1;
      keyOccurrences.set(key, count);

      list.push({
        id: r.id,
        categoryType: r.categoryType,
        date: rDate,
        teamName: r.teamName,
        operationTitle: 'ប្រើប្រាស់ប្រចាំថ្ងៃ (Form)',
        operationType: 'useTeam',
        values: r.values || {},
        totalSheets: r.totalSheets || 0,
        source: 'daily_form',
        isDuplicateOrOvercounted: count > 1,
        overcountReason: count > 1 ? 'មានទិន្នន័យកត់ត្រាស្ទួនលើសពី១ដងក្នុងថ្ងៃតែមួយ' : undefined,
      });
    });

    // 2. Add stock ledger records (usage, issue, old stock, damage, transfer)
    if (stockRecords && stockRecords.length > 0) {
      stockRecords.forEach((sr) => {
        const rawTeam =
          resolveRecordTeamName(sr) ||
          (sr.visaTeamRobokName || (sr as any).teamName || (sr as any).team || sr.sourceFrom || '').trim();
        const date = normalizeDateToISO(sr.date || '');
        if (!rawTeam) return;

        const cat: VisaCategoryOption =
          sr.sourceFrom === 'cEA' ||
          sr.remarks?.includes('cEA') ||
          (sr.visaType && sr.visaType.includes('cEA'))
            ? 'cEA'
            : 'Sticker';

        const vtId = normalizeVisaType(sr.visaType);
        const sheets =
          Number(sr.totalSheets) || (Number(sr.quantityBundles) ? Number(sr.quantityBundles) * 50 : 0);

        const isTransfer = isTransferTeamRecord(sr);
        const recipientTeam = (
          sr.recipientTeamName ||
          (sr as any).destinationTo ||
          (sr as any).targetTeam ||
          (sr as any).recipientTeamId ||
          ''
        ).trim();

        let opTitle = 'ប្រើប្រាស់ប្រចាំថ្ងៃ';
        let opType = sr.operationType || 'useTeam';

        if (sr.operationType === 'issueTeam' || isIssueTeamRecord(sr)) {
          opTitle = 'បើកផ្តល់ពីក២ (ទទួលពីការិយាល័យ)';
        } else if (
          sr.operationType === 'oldstockteam' ||
          sr.sourceFrom?.includes('ស្តុកចាស់') ||
          sr.sourceFrom?.includes('សន្និធិដើម')
        ) {
          opTitle = 'ស្តុកចាស់ក្រុម (ដើមគ្រា)';
          opType = 'oldstockteam';
        } else if (isTransfer) {
          opTitle = `ផ្ទេរការប្រើប្រាស់ (ផ្ទេរទៅ ${recipientTeam || 'ក្រុមផ្សេង'})`;
        } else if (
          sr.operationType === 'damagedTeam' ||
          sr.operationType === 'damaged' ||
          sr.operationType === 'returnTeam' ||
          sr.operationType === 'missingTeam'
        ) {
          opTitle = 'ខូច/បាត់/ប្រគល់ត្រឡប់';
        }

        const key = `${cat}_${opType}_${date}_${normalizeTeamName(rawTeam)}_${vtId}`;
        const count = (keyOccurrences.get(key) || 0) + 1;
        keyOccurrences.set(key, count);

        const valuesObj: any = {};
        if (vtId) {
          valuesObj[vtId] = {
            quantity: sheets,
            startSerial: sr.startSerial || '',
            endSerial: sr.endSerial || '',
          };
        }

        list.push({
          id: sr.id,
          categoryType: cat,
          date: date || '30-Nov-2018',
          teamName: rawTeam,
          operationTitle: opTitle,
          operationType: opType,
          values: valuesObj,
          totalSheets: sheets,
          source: 'stock_ledger',
          isDuplicateOrOvercounted: count > 1,
          overcountReason: count > 1 ? 'ទិន្នន័យស្ទួនរវាងតារាងប្រចាំថ្ងៃ និងសៀវភៅស្តុក' : undefined,
        });

        // If transfer, also list under recipient team as "បើកផ្តល់ពីក២ (ផ្ទេរពីក្រុម)"
        if (isTransfer && recipientTeam) {
          const recNormTeam = normalizeTeamName(recipientTeam);
          const recKey = `${cat}_transferIn_${date}_${recNormTeam}_${vtId}`;
          const recCount = (keyOccurrences.get(recKey) || 0) + 1;
          keyOccurrences.set(recKey, recCount);

          list.push({
            id: `${sr.id}-recip`,
            categoryType: cat,
            date: date || '30-Nov-2018',
            teamName: recipientTeam,
            operationTitle: `បើកផ្តល់ពីក២ (ផ្ទេរពី ${rawTeam})`,
            operationType: 'transferTeam',
            values: valuesObj,
            totalSheets: sheets,
            source: 'stock_ledger',
            isDuplicateOrOvercounted: recCount > 1,
            overcountReason: recCount > 1 ? 'ទិន្នន័យស្ទួនរវាងតារាងប្រចាំថ្ងៃ និងសៀវភៅស្តុក' : undefined,
          });
        }
      });
    }

    return list;
  }, [isAuditChecklistOpen, records, stockRecords]);

  // Filtered audit entries for Checklist Modal
  const filteredAuditEntries = useMemo(() => {
    if (!isAuditChecklistOpen || auditAllEntries.length === 0) return [];
    return auditAllEntries.filter((item) => {
      // Date filter
      if (auditDateFilter === 'current') {
        if (item.date !== selectedDate) return false;
      } else if (auditDateFilter === 'custom' && auditCustomDate) {
        if (item.date !== auditCustomDate) return false;
      }

      // Category filter
      if (auditCategoryFilter !== 'all') {
        if (item.categoryType !== auditCategoryFilter) return false;
      }

      // Team filter
      if (auditTeamFilter !== 'all') {
        if (normalizeTeamName(item.teamName) !== normalizeTeamName(auditTeamFilter)) return false;
      }

      // Overcounted only
      if (auditShowOnlyOvercounted && !item.isDuplicateOrOvercounted) {
        return false;
      }

      // Search term
      if (deferredAuditSearch.trim()) {
        const q = deferredAuditSearch.trim().toLowerCase();
        const matchTeam = item.teamName.toLowerCase().includes(q);
        const matchDate = item.date.toLowerCase().includes(q);
        const matchCat = item.categoryType.toLowerCase().includes(q);
        const matchOp = item.operationTitle.toLowerCase().includes(q);
        if (!matchTeam && !matchDate && !matchCat && !matchOp) return false;
      }

      return true;
    });
  }, [
    isAuditChecklistOpen,
    auditAllEntries,
    auditDateFilter,
    selectedDate,
    auditCustomDate,
    auditCategoryFilter,
    auditTeamFilter,
    auditShowOnlyOvercounted,
    deferredAuditSearch,
  ]);

  // Ultra-fast lightweight count of detected overcounted/duplicate items for active selectedDate
  const activeDateOvercountedCount = useMemo(() => {
    const normDate = normalizeDateToISO(selectedDate);
    let count = 0;
    const keyOccurrences = new Map<string, number>();

    records.forEach((r) => {
      if (normalizeDateToISO(r.date || '') === normDate) {
        const key = `${r.categoryType}_use_${normDate}_${normalizeTeamName(r.teamName)}`;
        const c = (keyOccurrences.get(key) || 0) + 1;
        keyOccurrences.set(key, c);
        if (c > 1) count++;
      }
    });

    if (stockRecords && stockRecords.length > 0) {
      stockRecords.forEach((sr) => {
        if (normalizeDateToISO(sr.date || '') === normDate) {
          const rawTeam = resolveRecordTeamName(sr) || (sr.visaTeamRobokName || '').trim();
          if (!rawTeam) return;
          const cat =
            sr.sourceFrom === 'cEA' || sr.remarks?.includes('cEA') || sr.visaType?.includes('cEA')
              ? 'cEA'
              : 'Sticker';
          const vtId = normalizeVisaType(sr.visaType);
          const opType = sr.operationType || 'useTeam';
          const key = `${cat}_${opType}_${normDate}_${normalizeTeamName(rawTeam)}_${vtId}`;
          const c = (keyOccurrences.get(key) || 0) + 1;
          keyOccurrences.set(key, c);
          if (c > 1) count++;
        }
      });
    }

    return count;
  }, [records, stockRecords, selectedDate]);

  // Pre-indexed fast lookup map for prior used serial codes by (normTeam, category, visaType)
  const priorCodesIndex = useMemo(() => {
    const map = new Map<string, Array<{ date: string; serial: string }>>();

    const addEntry = (team: string, cat: string, vt: string, date: string, serial: string) => {
      if (!team || !cat || !vt || !serial || !date) return;
      const key = `${normalizeTeamName(team)}_${cat}_${vt.toUpperCase()}`;
      let arr = map.get(key);
      if (!arr) {
        arr = [];
        map.set(key, arr);
      }
      arr.push({ date: normalizeDateToISO(date), serial: serial.trim() });
    };

    // 1. Index from effectiveRecords
    effectiveRecords.forEach((r) => {
      const cat = r.categoryType;
      const date = r.date;
      const team = r.teamName;
      if (!r.values) return;
      Object.entries(r.values).forEach(([vtId, v]: [string, any]) => {
        if (!v) return;
        if (Array.isArray(v.entries)) {
          v.entries.forEach((e: FormRowEntry) => {
            if (e.endSerial?.trim()) addEntry(team, cat, vtId, date, e.endSerial.trim());
            else if (e.startSerial?.trim()) addEntry(team, cat, vtId, date, e.startSerial.trim());
            else if (e.oldCode?.trim()) addEntry(team, cat, vtId, date, e.oldCode.trim());
          });
        } else {
          if (v.endSerial?.trim()) addEntry(team, cat, vtId, date, v.endSerial.trim());
          else if (v.startSerial?.trim()) addEntry(team, cat, vtId, date, v.startSerial.trim());
          else if (v.oldCode?.trim()) addEntry(team, cat, vtId, date, v.oldCode.trim());
        }
      });
    });

    // 2. Index from stockRecords (ONLY actual usage or recorded old stock baseline, NOT issued stock)
    if (stockRecords && stockRecords.length > 0) {
      stockRecords.forEach((sr) => {
        if (sr.stockType && sr.stockType !== 'sticker') return;
        const isUsage = sr.operationType === 'useTeam' || sr.sourceFrom?.includes('ប្រើប្រាស់');
        const isOld = isOldStockTeamRecord(sr);
        if (!isUsage && !isOld) return;

        const sTeam = resolveRecordTeamName(sr) || sr.visaTeamRobokName || (sr as any).teamName || '';
        if (!sTeam) return;
        const sDate = sr.date || '';
        const sVt = normalizeVisaType(sr.visaType);
        if (!sVt) return;
        const cat =
          sr.sourceFrom === 'cEA' || sr.remarks?.includes('cEA') || sr.visaType?.includes('cEA')
            ? 'cEA'
            : 'Sticker';
        if (sr.endSerial?.trim()) addEntry(sTeam, cat, sVt, sDate, sr.endSerial.trim());
        else if (sr.startSerial?.trim()) addEntry(sTeam, cat, sVt, sDate, sr.startSerial.trim());
      });
    }

    // Sort entries descending by date for fast find
    map.forEach((arr) => {
      arr.sort((a, b) => b.date.localeCompare(a.date));
    });

    return map;
  }, [effectiveRecords, stockRecords]);

  // Fast direct map of effectiveRecords and direct records for instant O(1) searches
  const recordLookupMap = useMemo(() => {
    const map = new Map<string, { record: DailyTeamRecord; source: 'direct' | 'effective' }>();
    effectiveRecords.forEach((r) => {
      const key = `${r.categoryType}_${normalizeDateToISO(r.date)}_${normalizeTeamName(r.teamName)}`;
      map.set(key, { record: r, source: 'effective' });
    });
    records.forEach((r) => {
      const key = `${r.categoryType}_${normalizeDateToISO(r.date)}_${normalizeTeamName(r.teamName)}`;
      map.set(key, { record: r, source: 'direct' });
    });
    return map;
  }, [records, effectiveRecords]);

  // Checklist Actions
  const handleToggleIncludeItem = (id: string) => {
    setExcludedIds((prev) => {
      if (prev.includes(id)) {
        return prev.filter((x) => x !== id);
      } else {
        return [...prev, id];
      }
    });
  };

  const handleSelectAllInChecklist = () => {
    const visibleIds = filteredAuditEntries.map((i) => i.id);
    setExcludedIds((prev) => prev.filter((id) => !visibleIds.includes(id)));
  };

  const handleUncheckAllInChecklist = () => {
    const visibleIds = filteredAuditEntries.map((i) => i.id);
    setExcludedIds((prev) => Array.from(new Set([...prev, ...visibleIds])));
  };

  const handleAutoUncheckOvercounted = () => {
    const overcountedIds = filteredAuditEntries
      .filter((i) => i.isDuplicateOrOvercounted)
      .map((i) => i.id);
    setExcludedIds((prev) => Array.from(new Set([...prev, ...overcountedIds])));
    onShowToast(`បានដោះធីកទិន្នន័យស្ទួន/រាប់លើសចំនួន ${overcountedIds.length} ជោគជ័យ!`, 'success');
  };

  const handleSelectRecordForEdit = (item: any) => {
    // 1. Close Checklist modal
    setIsAuditChecklistOpen(false);

    // 2. Resolve parameters
    const teamToUse = item.teamName;
    const dateToUse = normalizeDateToISO(item.date);
    const optionToUse = item.categoryType;

    setSelectedTeam(teamToUse);
    setSelectedDate(dateToUse);
    setSelectedOption(optionToUse);

    // 3. Find if there's a daily team record
    const match = findRecordForDateTeamOption(dateToUse, teamToUse, optionToUse);
    if (match && match.record) {
      handleEditSpecificRecord(match.record);
    } else {
      // Fallback to stock record edit
      const srMatch = (stockRecords || []).find((s) => s.id === item.id);
      if (srMatch) {
        // Load matching stock records
        const normTeam = normalizeTeamName(teamToUse);
        const matchingStock = (stockRecords || []).filter((sr) => {
          const isDateMatch = normalizeDateToISO(sr.date || '') === dateToUse;
          const isTeamMatch =
            normalizeTeamName(resolveRecordTeamName(sr) || sr.visaTeamRobokName || (sr as any).teamName || '') === normTeam;
          const isCatMatch =
            optionToUse === 'cEA'
              ? sr.sourceFrom === 'cEA' || sr.remarks?.includes('cEA') || sr.visaType?.includes('cEA')
              : sr.sourceFrom !== 'cEA' && !sr.remarks?.includes('cEA') && !sr.visaType?.includes('cEA');
          return isDateMatch && isTeamMatch && isCatMatch;
        });

        setEditingRecordId(srMatch.id);
        const updated: Record<string, FormRowEntry[]> = {};
        VISA_TYPES.forEach((vt) => {
          const lastCode = getLastUsedOldCode(teamToUse, vt.id, dateToUse, optionToUse);
          const stockForVt = matchingStock.filter((sr) => normalizeVisaType(sr.visaType) === vt.id);

          if (stockForVt.length > 0) {
            updated[vt.id] = stockForVt.map((sr, idx) => ({
              id: `${vt.id}-${idx}`,
              quantity: String(sr.totalSheets || (sr.quantityBundles ? Number(sr.quantityBundles) * 50 : '')),
              startSerial: sr.startSerial || '',
              endSerial: sr.endSerial || '',
              oldCode: (sr as any).oldCode || (idx === 0 ? lastCode : ''),
            }));
          } else {
            updated[vt.id] = [
              {
                id: `${vt.id}-0`,
                quantity: '',
                startSerial: '',
                endSerial: '',
                oldCode: lastCode || '',
              },
            ];
          }
        });
        setFormData(updated);
      }
    }
    onShowToast(`បានបើកទិន្នន័យសម្រាប់កែប្រែ (${teamToUse} - ${dateToUse})`, 'info');
  };

  const handleDeleteAuditRecord = async (id: string, source: 'daily_form' | 'stock_ledger') => {
    if (!window.confirm('តើអ្នកពិតជាចង់លុបទិន្នន័យប្រតិបត្តិការនេះចោលទាំងស្រុងឬទេ?')) return;

    if (source === 'daily_form') {
      const targetRecord = records.find((r) => r.id === id);
      setRecords((prev) => prev.filter((r) => r.id !== id));
      
      try {
        await apiService.deleteDailyTeamOperation(id);
      } catch (err) {
        onShowToast('មានបញ្ហាក្នុងការលុបទិន្នន័យនៅលើម៉ាស៊ីនមេ', 'error');
        // Revert? (Optional: fetch records again)
        return;
      }

      if (targetRecord) {
        const normTeam = normalizeTeamName(targetRecord.teamName);
        const isCea = targetRecord.categoryType === 'cEA';
        const targetDate = normalizeDateToISO(targetRecord.date);
        
        const linkedStockIds = (stockRecords || [])
          .filter((sr) => {
            const srTeam = normalizeTeamName(resolveRecordTeamName(sr) || sr.visaTeamRobokName || (sr as any).teamName || '');
            const srDate = normalizeDateToISO(sr.date || '');
            const srIsCea = sr.sourceFrom === 'cEA' || sr.remarks?.includes('cEA') || sr.visaType?.includes('cEA');
            
            return (
              srDate === targetDate &&
              srTeam === normTeam &&
              srIsCea === isCea &&
              (sr.operationType === 'useTeam' || sr.sourceFrom?.includes('ប្រើប្រាស់') || sr.id.startsWith('stock-dtr-') || sr.id.startsWith('dtr-'))
            );
          })
          .map((sr) => sr.id);

        if (linkedStockIds.length > 0) {
          if (onDeleteBatchStockRecords) {
            await onDeleteBatchStockRecords(linkedStockIds);
          } else if (onDeleteStockRecord) {
            for (const sid of linkedStockIds) {
              await onDeleteStockRecord(sid);
            }
          }
        }
      }
      onShowToast('បានលុបទិន្នន័យពីប្រព័ន្ធប្រចាំថ្ងៃ និងបញ្ជីស្តុកទូទៅរួចរាល់', 'success');
    } else {
      // Stock ledger source
      if (onDeleteStockRecord) {
        await onDeleteStockRecord(id);
      }
      setExcludedIds((prev) => Array.from(new Set([...prev, id])));
      onShowToast('បានលុបទិន្នន័យពីបញ្ជីស្តុកទូទៅរួចរាល់', 'success');
    }
  };

  // Helper to find the most recent previous record for this team (checks today's prior entries first, then indexed prior dates)
  const getLastUsedOldCode = (
    team: string,
    visaTypeId: string,
    beforeDate: string,
    categoryType: VisaCategoryOption
  ): string => {
    if (!team) return '';
    const normTeam = normalizeTeamName(team);
    const normBefore = normalizeDateToISO(beforeDate);

    // 1. Check today's records for this team (prior batches today)
    const todayRecords = records.filter(
      (r) =>
        r.categoryType === categoryType &&
        normalizeDateToISO(r.date) === normBefore &&
        (r.teamName === team || normalizeTeamName(r.teamName) === normTeam)
    );
    for (let i = todayRecords.length - 1; i >= 0; i--) {
      const v = todayRecords[i].values?.[visaTypeId];
      if (v) {
        if (Array.isArray(v.entries) && v.entries.length > 0) {
          const lastValid = [...v.entries].reverse().find((e) => e.endSerial?.trim() || e.startSerial?.trim());
          if (lastValid?.endSerial?.trim()) return lastValid.endSerial.trim();
          if (lastValid?.startSerial?.trim()) return lastValid.startSerial.trim();
        }
        if (v.endSerial && v.endSerial.trim()) return v.endSerial.trim();
        if (v.startSerial && v.startSerial.trim()) return v.startSerial.trim();
        if (v.oldCode && v.oldCode.trim()) return v.oldCode.trim();
      }
    }

    // 2. Query pre-indexed priorCodesIndex for fast O(1) retrieval
    const key = `${normTeam}_${categoryType}_${visaTypeId.toUpperCase()}`;
    const list = priorCodesIndex.get(key);
    if (list && list.length > 0) {
      const match = list.find((item) => item.date < normBefore);
      if (match) return match.serial;
    }

    // 3. Fallback: If no prior usage code found, check if this team received an issued batch from K2
    if (stockRecords && stockRecords.length > 0) {
      const issues = stockRecords
        .filter((sr) => {
          if (sr.stockType && sr.stockType !== 'sticker') return false;
          const isSender = isRecordForTeam(sr, team);
          const isRecip = isRecordForRecipientTeam(sr, team);
          const isTrans = isTransferTeamRecord(sr);
          const isIss = sr.operationType === 'issueTeam' || isIssueTeamRecord(sr);
          if (!(isIss && isSender) && !(isTrans && isRecip)) return false;

          const sVt = normalizeVisaType(sr.visaType);
          if (sVt !== visaTypeId.toUpperCase()) return false;
          const sDate = normalizeDateToISO(sr.date || '');
          return sDate && sDate <= normBefore && Boolean(sr.startSerial?.trim());
        })
        .sort((a, b) => (a.date || '').localeCompare(b.date || ''));

      if (issues.length > 0 && issues[0].startSerial) {
        const s = issues[0].startSerial.trim();
        const prevCode = decrementSerial(s, 1);
        return prevCode || s;
      }
    }

    return '';
  };

  // Today's specific entries for the active team & date & category (allows viewing and editing multiple entries per day)
  const todayTeamRecords = useMemo(() => {
    if (!selectedDate || !selectedTeam) return [];
    const normTeam = normalizeTeamName(selectedTeam);
    const normDate = normalizeDateToISO(selectedDate);
    return records.filter(
      (r) =>
        r.categoryType === selectedOption &&
        normalizeDateToISO(r.date) === normDate &&
        (r.teamName === selectedTeam || normalizeTeamName(r.teamName) === normTeam)
    );
  }, [records, selectedOption, selectedDate, selectedTeam]);

  // Helper to find the latest prior record for a team before a target date
  const getLatestPriorRecord = (
    team: string,
    beforeDate: string,
    categoryType: VisaCategoryOption
  ): DailyTeamRecord | null => {
    if (!team || !beforeDate) return null;
    const normTeam = normalizeTeamName(team);
    const normBefore = normalizeDateToISO(beforeDate);

    // 1. Look in effectiveRecords (which already aggregates multiple entries per day)
    const matches = effectiveRecords
      .filter(
        (r) =>
          r.categoryType === categoryType &&
          normalizeDateToISO(r.date) < normBefore &&
          (r.teamName === team || normalizeTeamName(r.teamName) === normTeam)
      )
      .sort((a, b) => normalizeDateToISO(b.date).localeCompare(normalizeDateToISO(a.date)));

    if (matches.length > 0) return matches[0];

    // 2. Fallback to raw records
    const rawMatches = records
      .filter(
        (r) =>
          r.categoryType === categoryType &&
          normalizeDateToISO(r.date) < normBefore &&
          (r.teamName === team || normalizeTeamName(r.teamName) === normTeam)
      )
      .sort((a, b) => normalizeDateToISO(b.date).localeCompare(normalizeDateToISO(a.date)));

    return rawMatches.length > 0 ? rawMatches[0] : null;
  };

  // Helper to construct form row entries for all visa types from the latest prior date or prior batch
  // Supports multi-row types (e.g. T having 2 rows from the prior day -> automatically shows 2 rows with both old codes on the next day)
  const getPriorFormEntriesForTeam = (
    team: string,
    targetDate: string,
    categoryType: VisaCategoryOption
  ): Record<string, FormRowEntry[]> => {
    const result: Record<string, FormRowEntry[]> = {};
    if (!team) {
      VISA_TYPES.forEach((vt) => {
        result[vt.id] = [{ id: `${vt.id}-0`, quantity: '', startSerial: '', endSerial: '', oldCode: '' }];
      });
      return result;
    }

    const normTeam = normalizeTeamName(team);
    const normTarget = normalizeDateToISO(targetDate);

    // 1. Check if there are already records on targetDate for this team (prior batches entered today)
    const todayRecs = records.filter(
      (r) =>
        r.categoryType === categoryType &&
        normalizeDateToISO(r.date) === normTarget &&
        (r.teamName === team || normalizeTeamName(r.teamName) === normTeam)
    );

    // 2. Find the latest prior record before targetDate
    const latestPrior = getLatestPriorRecord(team, targetDate, categoryType);

    VISA_TYPES.forEach((vt) => {
      // Priority 1: If today has prior records, get entries from today's latest batch
      if (todayRecs.length > 0) {
        for (let i = todayRecs.length - 1; i >= 0; i--) {
          const valObj = todayRecs[i].values?.[vt.id] as any;
          if (valObj) {
            const subEntries: any[] = Array.isArray(valObj.entries) && valObj.entries.length > 0
              ? valObj.entries
              : (valObj.quantity || valObj.startSerial)
              ? [valObj]
              : [];
            if (subEntries.length > 0) {
              result[vt.id] = subEntries.map((e: any, idx: number) => {
                const qNum = parseInt(e.quantity || '', 10) || 0;
                const end =
                  e.endSerial?.trim() ||
                  (e.startSerial && qNum > 0 ? incrementSerial(e.startSerial.trim(), qNum - 1) : '') ||
                  e.startSerial?.trim() ||
                  e.oldCode?.trim() ||
                  '';
                return {
                  id: `${vt.id}-${idx}`,
                  quantity: '',
                  startSerial: '',
                  endSerial: '',
                  oldCode: end,
                };
              });
              return;
            }
          }
        }
      }

      // Priority 2: Use entries from the latest prior record before targetDate (e.g., 2 rows of T from 8/29/2026)
      if (latestPrior && latestPrior.values?.[vt.id]) {
        const valObj = latestPrior.values[vt.id] as any;
        const subEntries: any[] = Array.isArray(valObj.entries) && valObj.entries.length > 0
          ? valObj.entries
          : (valObj.quantity || valObj.startSerial)
          ? [valObj]
          : [];

        if (subEntries.length > 0) {
          result[vt.id] = subEntries.map((e: any, idx: number) => {
            const qNum = parseInt(e.quantity || '', 10) || 0;
            const end =
              e.endSerial?.trim() ||
              (e.startSerial && qNum > 0 ? incrementSerial(e.startSerial.trim(), qNum - 1) : '') ||
              e.startSerial?.trim() ||
              e.oldCode?.trim() ||
              '';
            return {
              id: `${vt.id}-${idx}`,
              quantity: '',
              startSerial: '',
              endSerial: '',
              oldCode: end,
            };
          });
          return;
        }
      }

      // Priority 3: Fallback using getLastUsedOldCode (checks earlier dates, stock issues, etc.)
      const lastCode = getLastUsedOldCode(team, vt.id, targetDate, categoryType);
      result[vt.id] = [
        {
          id: `${vt.id}-0`,
          quantity: '',
          startSerial: '',
          endSerial: '',
          oldCode: lastCode || '',
        },
      ];
    });

    return result;
  };

  // Return all unique candidate prior codes for a team, type, and date
  const getAvailablePriorCodes = (
    team: string,
    typeId: string,
    targetDate: string,
    categoryType: VisaCategoryOption
  ): string[] => {
    if (!team || !targetDate) return [];
    const codes = new Set<string>();

    const normTeam = normalizeTeamName(team);
    const normTarget = normalizeDateToISO(targetDate);

    // 1. From today's records (prior entries today)
    const todayRecs = records.filter(
      (r) =>
        r.categoryType === categoryType &&
        normalizeDateToISO(r.date) === normTarget &&
        (r.teamName === team || normalizeTeamName(r.teamName) === normTeam)
    );
    todayRecs.forEach((r) => {
      const valObj = r.values?.[typeId] as any;
      if (valObj) {
        const subEntries: any[] = Array.isArray(valObj.entries) && valObj.entries.length > 0
          ? valObj.entries
          : (valObj.quantity || valObj.startSerial)
          ? [valObj]
          : [];
        subEntries.forEach((e: any) => {
          const qNum = parseInt(e.quantity || '', 10) || 0;
          const end =
            e.endSerial?.trim() ||
            (e.startSerial && qNum > 0 ? incrementSerial(e.startSerial.trim(), qNum - 1) : '') ||
            e.startSerial?.trim() ||
            e.oldCode?.trim();
          if (end) codes.add(end);
        });
      }
    });

    // 2. From latest prior record (e.g. on 8/29/2026 where T has 2 rows)
    const latestPrior = getLatestPriorRecord(team, targetDate, categoryType);
    if (latestPrior && latestPrior.values?.[typeId]) {
      const valObj = latestPrior.values[typeId] as any;
      const subEntries: any[] = Array.isArray(valObj.entries) && valObj.entries.length > 0
        ? valObj.entries
        : (valObj.quantity || valObj.startSerial)
        ? [valObj]
        : [];
      subEntries.forEach((e: any) => {
        const qNum = parseInt(e.quantity || '', 10) || 0;
        const end =
          e.endSerial?.trim() ||
          (e.startSerial && qNum > 0 ? incrementSerial(e.startSerial.trim(), qNum - 1) : '') ||
          e.startSerial?.trim() ||
          e.oldCode?.trim();
        if (end) codes.add(end);
      });
    }

    // 3. Fallback code from getLastUsedOldCode
    const fallback = getLastUsedOldCode(team, typeId, targetDate, categoryType);
    if (fallback) codes.add(fallback);

    return Array.from(codes).filter(Boolean);
  };

  // Date to display in the previous code badge (shows the actual latest date where prior codes were found)
  const displayPriorDate = useMemo(() => {
    if (!selectedDate || !selectedTeam) return previousDateString || selectedDate;
    const latest = getLatestPriorRecord(selectedTeam, selectedDate, selectedOption);
    return latest?.date || previousDateString || selectedDate;
  }, [selectedDate, selectedTeam, selectedOption, effectiveRecords, records, previousDateString]);

  // Helper to find a record across direct records, effective records, and stock records
  const findRecordForDateTeamOption = (
    dateStr: string,
    teamStr: string,
    optionStr: VisaCategoryOption
  ): { record: DailyTeamRecord; source: 'direct' | 'effective' | 'stock' } | null => {
    if (!dateStr || !teamStr) return null;
    const normTeam = normalizeTeamName(teamStr);
    const normDate = normalizeDateToISO(dateStr);
    const key = `${optionStr}_${normDate}_${normTeam}`;

    const lookup = recordLookupMap.get(key);
    if (lookup && lookup.record.values && Object.keys(lookup.record.values).length > 0) {
      return lookup;
    }

    // Fallback: Check stockRecords if not found
    if (stockRecords && stockRecords.length > 0) {
      const matchStock = stockRecords.filter((sr) => {
        const isDateMatch = normalizeDateToISO(sr.date || '') === normDate;
        const isTeamMatch =
          normalizeTeamName(resolveRecordTeamName(sr) || sr.visaTeamRobokName || (sr as any).teamName || '') ===
          normTeam;
        const isCatMatch =
          optionStr === 'cEA'
            ? sr.sourceFrom === 'cEA' || sr.remarks?.includes('cEA') || sr.visaType?.includes('cEA')
            : sr.sourceFrom !== 'cEA' && !sr.remarks?.includes('cEA') && !sr.visaType?.includes('cEA');
        return isDateMatch && isTeamMatch && isCatMatch;
      });

      if (matchStock.length > 0) {
        const values: any = {};
        let totalSheets = 0;
        matchStock.forEach((sr) => {
          const vtId = normalizeVisaType(sr.visaType);
          const sheets =
            Number(sr.totalSheets) || (Number(sr.quantityBundles) ? Number(sr.quantityBundles) * 50 : 0);
          if (vtId) {
            const cur = values[vtId] || { quantity: 0, startSerial: '', endSerial: '', oldCode: '' };
            values[vtId] = {
              quantity: (Number(cur.quantity) || 0) + sheets,
              startSerial: cur.startSerial || sr.startSerial || '',
              endSerial: sr.endSerial || cur.endSerial || '',
              oldCode: cur.oldCode || '',
            };
            totalSheets += sheets;
          }
        });
        return {
          record: {
            id: `stock-sync-${optionStr}-${normDate}-${teamStr.replace(/\s+/g, '-')}`,
            categoryType: optionStr,
            date: normDate,
            teamName: teamStr,
            values,
            totalSheets,
          } as DailyTeamRecord,
          source: 'stock',
        };
      }
    }

    return null;
  };

  // Helper to load existing record data directly into the Form textboxes
  const loadExistingRecordIntoForm = (
    teamName: string,
    dateStr: string,
    optionType: VisaCategoryOption
  ): boolean => {
    if (!teamName || !dateStr) return false;

    // Check if a specific record ID edit was requested via localStorage
    const editRecordId = localStorage.getItem('app_daily_team_edit_record_id');
    if (editRecordId) {
      const targetRecord = records.find((r) => r.id === editRecordId) || effectiveRecords.find((r) => r.id === editRecordId);
      if (targetRecord) {
        localStorage.removeItem('app_daily_team_edit_record_id');
        handleEditSpecificRecord(targetRecord);
        return true;
      }
      // Look up in stockRecords if from general ledger
      const srMatch = (stockRecords || []).find((s) => s.id === editRecordId);
      if (srMatch) {
        localStorage.removeItem('app_daily_team_edit_record_id');
        const srTeam = resolveRecordTeamName(srMatch) || srMatch.visaTeamRobokName || (srMatch as any).teamName || teamName;
        const srDate = normalizeDateToISO(srMatch.date || dateStr);
        const srIsCea = srMatch.sourceFrom === 'cEA' || srMatch.remarks?.includes('cEA') || (srMatch.visaType || '').includes('cEA');
        const srOption: VisaCategoryOption = srIsCea ? 'cEA' : 'Sticker';

        if (srTeam) setSelectedTeam(srTeam);
        if (srDate) setSelectedDate(srDate);
        setSelectedOption(srOption);

        const match = findRecordForDateTeamOption(srDate || dateStr, srTeam || teamName, srOption);
        if (match && match.record) {
          const found = match.record;
          setEditingRecordId(found.id);
          const updated: Record<string, FormRowEntry[]> = {};
          VISA_TYPES.forEach((vt) => {
            const val = found.values?.[vt.id];
            const lastCode = getLastUsedOldCode(srTeam || teamName, vt.id, srDate || dateStr, srOption);
            const subEntries = val?.entries;
            if (Array.isArray(subEntries) && subEntries.length > 0) {
              updated[vt.id] = subEntries.map((e, idx) => ({
                id: e.id || `${vt.id}-${idx}`,
                quantity: e.quantity ? String(e.quantity) : '',
                startSerial: e.startSerial || '',
                endSerial: e.endSerial || '',
                oldCode: e.oldCode || (idx === 0 ? lastCode : ''),
              }));
            } else {
              const vtMatch = normalizeVisaType(srMatch.visaType) === vt.id;
              const q = vtMatch
                ? String(srMatch.totalSheets || (srMatch.quantityBundles ? Number(srMatch.quantityBundles) * 50 : ''))
                : val?.quantity
                ? String(val.quantity)
                : '';
              updated[vt.id] = [
                {
                  id: `${vt.id}-0`,
                  quantity: q,
                  startSerial: vtMatch ? srMatch.startSerial || '' : val?.startSerial || '',
                  endSerial: vtMatch ? srMatch.endSerial || '' : val?.endSerial || '',
                  oldCode: val?.oldCode || lastCode || '',
                },
              ];
            }
          });
          setFormData(updated);
          return true;
        } else {
          // Fallback: If no DailyTeamRecord exists, load ALL matching stock records for this team, date, and option!
          const normTeam = normalizeTeamName(srTeam || teamName);
          const normDate = srDate || dateStr;
          const matchingStock = (stockRecords || []).filter((sr) => {
            const isDateMatch = normalizeDateToISO(sr.date || '') === normDate;
            const isTeamMatch =
              normalizeTeamName(resolveRecordTeamName(sr) || sr.visaTeamRobokName || (sr as any).teamName || '') === normTeam;
            const isCatMatch =
              srOption === 'cEA'
                ? sr.sourceFrom === 'cEA' || sr.remarks?.includes('cEA') || sr.visaType?.includes('cEA')
                : sr.sourceFrom !== 'cEA' && !sr.remarks?.includes('cEA') && !sr.visaType?.includes('cEA');
            return isDateMatch && isTeamMatch && isCatMatch;
          });

          setEditingRecordId(srMatch.id);
          const updated: Record<string, FormRowEntry[]> = {};
          VISA_TYPES.forEach((vt) => {
            const lastCode = getLastUsedOldCode(srTeam || teamName, vt.id, srDate || dateStr, srOption);
            const stockForVt = matchingStock.filter((sr) => normalizeVisaType(sr.visaType) === vt.id);

            if (stockForVt.length > 0) {
              updated[vt.id] = stockForVt.map((sr, idx) => ({
                id: `${vt.id}-${idx}`,
                quantity: String(sr.totalSheets || (sr.quantityBundles ? Number(sr.quantityBundles) * 50 : '')),
                startSerial: sr.startSerial || '',
                endSerial: sr.endSerial || '',
                oldCode: (sr as any).oldCode || (idx === 0 ? lastCode : ''),
              }));
            } else if (normalizeVisaType(srMatch.visaType) === vt.id) {
              const q = String(srMatch.totalSheets || (srMatch.quantityBundles ? Number(srMatch.quantityBundles) * 50 : ''));
              updated[vt.id] = [
                {
                  id: `${vt.id}-0`,
                  quantity: q,
                  startSerial: srMatch.startSerial || '',
                  endSerial: srMatch.endSerial || '',
                  oldCode: (srMatch as any).oldCode || lastCode || '',
                },
              ];
            } else {
              updated[vt.id] = [
                {
                  id: `${vt.id}-0`,
                  quantity: '',
                  startSerial: '',
                  endSerial: '',
                  oldCode: lastCode || '',
                },
              ];
            }
          });
          setFormData(updated);
          setSelectedVisaTypeFilter('ALL');
          return true;
        }
      }
    }

    // Check if a record exists for this team, date, and option
    const match = findRecordForDateTeamOption(dateStr, teamName, optionType);
    if (match && match.record && match.record.values && Object.keys(match.record.values).length > 0) {
      const found = match.record;
      if (found.id && !found.id.startsWith('stock-sync-')) {
        setEditingRecordId(found.id);
      } else {
        setEditingRecordId(null);
      }
      const updated: Record<string, FormRowEntry[]> = {};
      VISA_TYPES.forEach((vt) => {
        const val = found.values[vt.id];
        const lastCode = getLastUsedOldCode(teamName, vt.id, dateStr, optionType);
        const subEntries = val?.entries;
        if (Array.isArray(subEntries) && subEntries.length > 0) {
          updated[vt.id] = subEntries.map((e, idx) => ({
            id: e.id || `${vt.id}-${idx}`,
            quantity:
              e.quantity !== undefined && e.quantity !== '' && Number(e.quantity) > 0
                ? String(e.quantity)
                : '',
            startSerial: e.startSerial || '',
            endSerial: e.endSerial || '',
            oldCode: e.oldCode || (idx === 0 ? lastCode : ''),
          }));
        } else {
          const q =
            val?.quantity !== undefined && val?.quantity !== '' && Number(val.quantity) > 0
              ? String(val.quantity)
              : '';
          updated[vt.id] = [
            {
              id: `${vt.id}-0`,
              quantity: q,
              startSerial: val?.startSerial || '',
              endSerial: val?.endSerial || '',
              oldCode: val?.oldCode || lastCode || '',
            },
          ];
        }
      });
      setFormData(updated);
      return true;
    }
    return false;
  };

  // Form input rows for each visa type (supports multiple entries per visa type)
  const [formData, setFormData] = useState<{
    [visaType: string]: FormRowEntry[];
  }>(() => {
    try {
      const payloadStr = localStorage.getItem('app_daily_team_edit_payload');
      if (payloadStr) {
        const payload = JSON.parse(payloadStr);
        if (payload) {
          const pTeam = payload.targetTeamName || payload.visaTeamRobokName || payload.recipientTeamName || payload.sourceFrom || '';
          const pDate = normalizeDateToISO(payload.date || '');
          const pOption: VisaCategoryOption = payload.optionMode || (payload.sourceFrom === 'cEA' || payload.remarks?.includes('cEA') || (payload.visaType || '').includes('cEA') ? 'cEA' : 'Sticker');
          const normTeam = normalizeTeamName(pTeam);

          // Find ALL matching stockRecords in props.stockRecords if available
          const matchingStock = (stockRecords || []).filter((sr) => {
            const srDate = normalizeDateToISO(sr.date || '');
            const srTeam = normalizeTeamName(resolveRecordTeamName(sr) || sr.visaTeamRobokName || (sr as any).teamName || '');
            const isCea = sr.sourceFrom === 'cEA' || sr.remarks?.includes('cEA') || sr.visaType?.includes('cEA');
            const isOptMatch = pOption === 'cEA' ? isCea : !isCea;
            return srDate === pDate && srTeam === normTeam && isOptMatch;
          });

          const initialForm: Record<string, FormRowEntry[]> = {};
          VISA_TYPES.forEach((vt) => {
            const stockForVt = matchingStock.filter((sr) => normalizeVisaType(sr.visaType) === vt.id);
            if (stockForVt.length > 0) {
              initialForm[vt.id] = stockForVt.map((sr, idx) => ({
                id: `${vt.id}-${idx}`,
                quantity: String(sr.totalSheets || (sr.quantityBundles ? Number(sr.quantityBundles) * 50 : '')),
                startSerial: sr.startSerial || '',
                endSerial: sr.endSerial || '',
                oldCode: (sr as any).oldCode || '',
              }));
            } else if (normalizeVisaType(payload.visaType) === vt.id) {
              const q = String(payload.totalSheets || (payload.quantityBundles ? Number(payload.quantityBundles) * 50 : ''));
              initialForm[vt.id] = [
                {
                  id: `${vt.id}-0`,
                  quantity: q,
                  startSerial: payload.startSerial || '',
                  endSerial: payload.endSerial || '',
                  oldCode: payload.oldCode || '',
                },
              ];
            } else {
              initialForm[vt.id] = [
                {
                  id: `${vt.id}-0`,
                  quantity: '',
                  startSerial: '',
                  endSerial: '',
                  oldCode: '',
                },
              ];
            }
          });
          localStorage.removeItem('app_daily_team_edit_payload');
          localStorage.removeItem('app_daily_team_edit_record_id');
          localStorage.removeItem('app_daily_team_auto_load_form');
          return initialForm;
        }
      }
    } catch (e) {
      console.error(e);
    }
    return getPriorFormEntriesForTeam(selectedTeam, selectedDate, selectedOption);
  });


  // Handle Form change for each field with automatic range calculation
  const handleInputChange = (
    typeId: string,
    entryIndex: number,
    field: 'quantity' | 'startSerial' | 'endSerial' | 'oldCode',
    val: string
  ) => {
    setFormData((prev) => {
      const list = prev[typeId] || [
        { id: `${typeId}-0`, quantity: '', startSerial: '', endSerial: '', oldCode: '' },
      ];
      const current = list[entryIndex] || {
        id: `${typeId}-${entryIndex}`,
        quantity: '',
        startSerial: '',
        endSerial: '',
        oldCode: '',
      };
      const updated = { ...current, [field]: val };

      // Automatic range calculation
      const sNum = parseInt(updated.startSerial, 10);
      const eNum = parseInt(updated.endSerial, 10);
      const qNum = parseInt(updated.quantity, 10);

      if (field === 'startSerial') {
        if (updated.startSerial && !isNaN(qNum) && qNum > 0) {
          updated.endSerial = incrementSerial(updated.startSerial, qNum - 1);
        } else if (!isNaN(sNum) && !isNaN(eNum) && eNum >= sNum) {
          updated.quantity = String(eNum - sNum + 1);
        }
        if (!updated.oldCode && updated.startSerial) {
          if (entryIndex > 0 && list[entryIndex - 1]?.endSerial) {
            updated.oldCode = list[entryIndex - 1].endSerial;
          } else {
            updated.oldCode = decrementSerial(updated.startSerial, 1);
          }
        }
      } else if (field === 'endSerial') {
        if (!isNaN(sNum) && !isNaN(eNum) && eNum >= sNum) {
          updated.quantity = String(eNum - sNum + 1);
        }
      } else if (field === 'quantity') {
        if (updated.startSerial && !isNaN(qNum) && qNum > 0) {
          updated.endSerial = incrementSerial(updated.startSerial, qNum - 1);
        }
      }

      const newList = [...list];
      newList[entryIndex] = updated;

      // If this entry now has an endSerial, keep the next row's oldCode synchronized
      if (updated.endSerial && newList[entryIndex + 1]) {
        newList[entryIndex + 1] = {
          ...newList[entryIndex + 1],
          oldCode: updated.endSerial,
        };
      }

      return {
        ...prev,
        [typeId]: newList,
      };
    });
  };

  // Add an additional row for a Visa Type
  const handleAddRow = (typeId: string, entryIndex: number) => {
    setFormData((prev) => {
      const list = prev[typeId] || [];
      const current = list[entryIndex] || list[list.length - 1];
      const qNum = parseInt(current?.quantity || '', 10) || 0;
      const calculatedCurrentEnd =
        current?.startSerial && qNum > 0
          ? incrementSerial(current.startSerial, qNum - 1)
          : '';

      const prevEnd =
        current?.endSerial?.trim() ||
        calculatedCurrentEnd ||
        current?.oldCode?.trim() ||
        getLastUsedOldCode(selectedTeam, typeId, selectedDate, selectedOption);

      const nextStart = prevEnd ? incrementSerial(prevEnd, 1) : '';

      const newEntry: FormRowEntry = {
        id: `${typeId}-${Date.now()}-${Math.random().toString(36).substr(2, 4)}`,
        quantity: '',
        startSerial: nextStart,
        endSerial: '',
        oldCode: prevEnd || '',
      };

      const newList = [...list];
      newList.splice(entryIndex + 1, 0, newEntry);

      return {
        ...prev,
        [typeId]: newList,
      };
    });
    onShowToast(`បានបន្ថែមជួរ ${typeId} ថ្មី`, 'info');
  };

  // Remove a sub-row for a Visa Type
  const handleRemoveRow = (typeId: string, entryIndex: number) => {
    setFormData((prev) => {
      const list = prev[typeId] || [];
      if (list.length <= 1) {
        const fallback = getPriorFormEntriesForTeam(selectedTeam, selectedDate, selectedOption)[typeId] || [
          {
            id: `${typeId}-0`,
            quantity: '',
            startSerial: '',
            endSerial: '',
            oldCode: getLastUsedOldCode(selectedTeam, typeId, selectedDate, selectedOption) || '',
          },
        ];
        return {
          ...prev,
          [typeId]: fallback,
        };
      }
      const newList = list.filter((_, idx) => idx !== entryIndex);
      return {
        ...prev,
        [typeId]: newList,
      };
    });
  };

  // Apply old code to start serial for a specific row (+1 step automatically)
  const handleApplyOldCodeToStartSerial = (
    typeId: string,
    entryIndex: number,
    rowIndex?: number
  ) => {
    const list = formData[typeId] || [];
    const current = list[entryIndex];
    let oldCode = current?.oldCode?.trim();
    if (!oldCode && entryIndex > 0) {
      oldCode = list[entryIndex - 1]?.endSerial?.trim();
    }
    if (!oldCode && current?.startSerial?.trim()) {
      oldCode = decrementSerial(current.startSerial.trim(), 1);
    }
    if (!oldCode) {
      oldCode = getLastUsedOldCode(selectedTeam, typeId, selectedDate, selectedOption)?.trim();
    }
    if (!oldCode) {
      onShowToast(`មិនទាន់មានលេខកូដចាស់សម្រាប់ប្រភេទ ${typeId} ទេ`, 'info');
      return;
    }

    const nextStart = incrementSerial(oldCode, 1);
    setFormData((prev) => {
      const curList = prev[typeId] || [];
      const curEntry = curList[entryIndex] || {
        id: `${typeId}-${entryIndex}`,
        quantity: '',
        startSerial: '',
        endSerial: '',
        oldCode: '',
      };
      const qNum = parseInt(curEntry.quantity, 10) || 0;
      let calculatedEnd = curEntry.endSerial;

      if (qNum > 0) {
        calculatedEnd = incrementSerial(nextStart, qNum - 1);
      }

      const updated = {
        ...curEntry,
        oldCode: oldCode,
        startSerial: nextStart,
        endSerial: calculatedEnd,
      };

      const newList = [...curList];
      newList[entryIndex] = updated;

      return {
        ...prev,
        [typeId]: newList,
      };
    });

    onShowToast(`បានកំណត់ចាប់ពីលេខ: ${nextStart} (កូដចាស់ + 1)`, 'success');

    // Focus startSerial input or endSerial/next
    setTimeout(() => {
      const rIdx =
        rowIndex !== undefined
          ? rowIndex
          : VISA_TYPES.findIndex((v) => v.id === typeId);
      const startInput = document.getElementById(
        `form-input-${rIdx}-${entryIndex}-1`
      ) as HTMLInputElement | null;
      if (startInput) {
        startInput.focus();
        startInput.select?.();
      }
    }, 40);
  };

  // Keyboard navigation across form textboxes (Arrow Up, Down, Left, Right, Enter)
  const handleKeyDown = (
    e: React.KeyboardEvent<HTMLInputElement>,
    rowIndex: number,
    entryIndex: number,
    fieldIndex: number
  ) => {
    const isCEA = selectedOption === 'cEA';
    const totalFields = isCEA ? 1 : 4;
    const totalRows = VISA_TYPES.length;

    let targetRow = rowIndex;
    let targetEntry = entryIndex;
    let targetField = fieldIndex;

    if (e.key === 'ArrowUp') {
      e.preventDefault();
      if (targetEntry > 0) {
        targetEntry -= 1;
      } else {
        targetRow = rowIndex > 0 ? rowIndex - 1 : totalRows - 1;
        targetEntry = 0;
      }
    } else if (e.key === 'ArrowDown' || e.key === 'Enter') {
      e.preventDefault();
      if (e.key === 'Enter' && !isCEA && fieldIndex < totalFields - 1) {
        targetField = fieldIndex + 1;
      } else {
        targetRow = rowIndex < totalRows - 1 ? rowIndex + 1 : 0;
        targetEntry = 0;
      }
    } else if (e.key === 'ArrowLeft') {
      const input = e.currentTarget;
      if (input.selectionStart === 0 && input.selectionEnd === 0) {
        e.preventDefault();
        if (fieldIndex > 0) {
          targetField = fieldIndex - 1;
        } else if (rowIndex > 0) {
          targetRow = rowIndex - 1;
          targetEntry = 0;
          targetField = totalFields - 1;
        }
      }
    } else if (e.key === 'ArrowRight') {
      const input = e.currentTarget;
      if (input.selectionStart === input.value.length && input.selectionEnd === input.value.length) {
        e.preventDefault();
        if (fieldIndex < totalFields - 1) {
          targetField = fieldIndex + 1;
        } else if (rowIndex < totalRows - 1) {
          targetRow = rowIndex + 1;
          targetEntry = 0;
          targetField = 0;
        }
      }
    } else {
      return;
    }

    const nextInputId = `form-input-${targetRow}-${targetEntry}-${targetField}`;
    const nextEl = document.getElementById(nextInputId) as HTMLInputElement | null;
    if (nextEl) {
      nextEl.focus();
      nextEl.select?.();
    }
  };

  // Calculate Form Total
  const formTotalSheets = useMemo(() => {
    return Object.values(formData).reduce((sum: number, list: FormRowEntry[]) => {
      if (!Array.isArray(list)) return sum;
      return (
        sum +
        list.reduce(
          (subSum: number, item: FormRowEntry) =>
            subSum + (parseInt(item?.quantity, 10) || 0),
          0
        )
      );
    }, 0);
  }, [formData]);

  // Actions
  // Handle Edit a specific individual batch/entry from today's list
  const handleEditSpecificRecord = (rec: DailyTeamRecord) => {
    if (rec.teamName) {
      setSelectedTeam(rec.teamName);
    }
    if (rec.date) {
      setSelectedDate(rec.date);
    }
    if (rec.categoryType === 'cEA' || (rec as any).sourceFrom === 'cEA') {
      setSelectedOption('cEA');
    } else if (rec.categoryType === 'Sticker') {
      setSelectedOption('Sticker');
    }

    setEditingRecordId(rec.id);
    const updated: Record<string, FormRowEntry[]> = {};
    let nonZeroVisaType = '';
    let nonZeroCount = 0;

    VISA_TYPES.forEach((vt) => {
      const val = rec.values?.[vt.id];
      const subEntries = val?.entries;
      if (Array.isArray(subEntries) && subEntries.length > 0) {
        updated[vt.id] = subEntries.map((e, idx) => ({
          id: e.id || `${vt.id}-${idx}`,
          quantity:
            e.quantity !== undefined && e.quantity !== '' && Number(e.quantity) > 0
              ? String(e.quantity)
              : '',
          startSerial: e.startSerial || '',
          endSerial: e.endSerial || '',
          oldCode: e.oldCode || '',
        }));
        const totalQ = subEntries.reduce(
          (s, e) => s + (parseInt(e.quantity, 10) || 0),
          0
        );
        if (totalQ > 0) {
          nonZeroCount++;
          nonZeroVisaType = vt.id;
        }
      } else {
        const q =
          val?.quantity !== undefined && val?.quantity !== '' && Number(val.quantity) > 0
            ? String(val.quantity)
            : '';
        if (q) {
          nonZeroCount++;
          nonZeroVisaType = vt.id;
        }
        updated[vt.id] = [
          {
            id: `${vt.id}-0`,
            quantity: q,
            startSerial: val?.startSerial || '',
            endSerial: val?.endSerial || '',
            oldCode: val?.oldCode || '',
          },
        ];
      }
    });
    setFormData(updated);

    if (nonZeroCount === 1) {
      setSelectedVisaTypeFilter(nonZeroVisaType);
    }
    onShowToast(`បានបើកទិន្នន័យលើកនេះសម្រាប់កែប្រែ`, 'info');
  };

  // Cancel edit mode and reset form to empty with latest oldCodes
  const handleCancelEdit = () => {
    setEditingRecordId(null);
    localStorage.removeItem('app_daily_team_edit_payload');
    localStorage.removeItem('app_daily_team_edit_record_id');
    localStorage.removeItem('app_daily_team_auto_load_form');
    const emptyForm = getPriorFormEntriesForTeam(selectedTeam, selectedDate, selectedOption);
    setFormData(emptyForm);
    onShowToast(`បានបោះបង់ការកែប្រែ`, 'info');
  };

  // Delete a specific batch entry
  const handleDeleteSpecificRecord = async (recordId: string) => {
    if (!window.confirm('តើអ្នកពិតជាចង់លុបទិន្នន័យលើកនេះឬទេ?')) return;

    const targetRecord = records.find((r) => r.id === recordId);
    setRecords((prev) => prev.filter((r) => r.id !== recordId));
    
    try {
      await apiService.deleteDailyTeamOperation(recordId);
      // Refresh local state by re-fetching
      const refreshed = await apiService.getDailyTeamOperations();
      setRecords(refreshed);
    } catch (err) {
      console.error('Delete failed:', err);
      onShowToast('មានបញ្ហាក្នុងការលុបទិន្នន័យ សូមព្យាយាមម្តងទៀត', 'error');
      // Re-fetch to revert to actual state
      const refreshed = await apiService.getDailyTeamOperations();
      setRecords(refreshed);
    }

    // Delete associated stock records if available
    if (targetRecord) {
      const normTeam = normalizeTeamName(targetRecord.teamName);
      const isCea = targetRecord.categoryType === 'cEA';
      const targetDate = normalizeDateToISO(targetRecord.date);

      const linkedStockIds = (stockRecords || [])
        .filter((sr) => {
          const srTeam = normalizeTeamName(resolveRecordTeamName(sr) || sr.visaTeamRobokName || (sr as any).teamName || '');
          const srDate = normalizeDateToISO(sr.date || '');
          const srIsCea = sr.sourceFrom === 'cEA' || sr.remarks?.includes('cEA') || sr.visaType?.includes('cEA');
          
          return (
            srDate === targetDate &&
            srTeam === normTeam &&
            srIsCea === isCea &&
            (sr.operationType === 'useTeam' || sr.sourceFrom?.includes('ប្រើប្រាស់') || sr.id.startsWith('stock-dtr-') || sr.id.startsWith('dtr-'))
          );
        })
        .map((sr) => sr.id);

      if (linkedStockIds.length > 0) {
        if (onDeleteBatchStockRecords) {
          onDeleteBatchStockRecords(linkedStockIds);
        } else if (onDeleteStockRecord) {
          linkedStockIds.forEach((id) => onDeleteStockRecord(id));
        }
      }
    }

    if (editingRecordId === recordId) {
      handleCancelEdit();
    }
    onShowToast('បានលុបទិន្នន័យលើកនេះរួចរាល់', 'success');
  };

  // Database Button (Red): Insert/Save to Daily Team Operations AND to Visa Sheet Stock Data (ទិន្នន័យសន្លឹកទិដ្ឋាការ)
  // Supports saving multiple batches per day for the same team (e.g., multiple ranges of T on the same day)
  const handleSaveRecord = () => {
    if (!selectedDate) {
      onShowToast('សូមជ្រើសរើសកាលបរិច្ឆេទ (ប្រចាំថ្ងៃ)', 'error');
      return;
    }
    if (!selectedTeam) {
      onShowToast('សូមជ្រើសរើសក្រុមជាមុនសិន', 'error');
      return;
    }
    if (formTotalSheets === 0) {
      onShowToast('សូមបញ្ចូលចំនួនសន្លឹកយ៉ាងហោចណាស់១ប្រភេទ', 'error');
      return;
    }

    const savedValues: any = {};
    const normTeam = normalizeTeamName(selectedTeam);
    const stockItemsToSave: StockRecord[] = [];

    VISA_TYPES.forEach((vt) => {
      const entries = formData[vt.id] || [];
      const totalQ = entries.reduce(
        (sum, e) => sum + (parseInt(e?.quantity || '', 10) || 0),
        0
      );
      const firstValid = entries.find((e) => e.startSerial?.trim());
      const lastValid = [...entries].reverse().find((e) => e.endSerial?.trim());
      
      const processedEntries = entries.map((e, eIdx) => {
        let oc = e.oldCode?.trim();
        if (!oc && eIdx > 0) {
          oc = entries[eIdx - 1]?.endSerial?.trim() || (e.startSerial ? decrementSerial(e.startSerial, 1) : '');
        } else if (!oc && e.startSerial) {
          oc = decrementSerial(e.startSerial, 1);
        }
        return {
          ...e,
          oldCode: oc || '',
        };
      });

      const oldCodeVal = processedEntries[0]?.oldCode || '';

      savedValues[vt.id] = {
        quantity: totalQ,
        startSerial: firstValid?.startSerial || '',
        endSerial: lastValid?.endSerial || '',
        oldCode: oldCodeVal,
        entries: processedEntries,
      };

      // Collect items for stock ledger
      entries.forEach((e) => {
        const q = parseInt(e?.quantity || '', 10) || 0;
        if (q > 0) {
          const robokTeamObj = categories?.visaTeamsRobok?.find(
            (t) => normalizeTeamName(t.name) === normTeam
          );

          stockItemsToSave.push({
            id: `stock-dtr-${selectedOption.toLowerCase()}-${selectedDate}-${selectedTeam.replace(/\s+/g, '-')}-${vt.id}-${Date.now()}-${Math.random().toString(36).substr(2, 4)}`,
            stockType: 'sticker',
            operationType: 'useTeam',
            date: selectedDate,
            time: new Date().toLocaleTimeString('en-US', { hour12: false, hour: '2-digit', minute: '2-digit' }),
            visaType: vt.name,
            quantityBundles: q,
            totalSheets: q,
            startSerial: e?.startSerial || '',
            endSerial: e?.endSerial || '',
            visaTeamRobokId: robokTeamObj?.id || `team-${selectedTeam}`,
            visaTeamRobokName: selectedTeam,
            sourceFrom: selectedOption === 'cEA' ? 'cEA' : 'Sticker',
            createdBy: userName || 'Admin',
            createdAt: new Date().toISOString(),
          });
        }
      });
    });

    const isEdit = Boolean(editingRecordId);
    const recordId =
      editingRecordId ||
      `dtr-${selectedOption}-${selectedDate}-${selectedTeam.replace(/\s+/g, '-')}-${Date.now()}`;
    const newRecord: DailyTeamRecord = {
      id: recordId,
      categoryType: selectedOption,
      date: selectedDate,
      teamName: selectedTeam,
      previousDate: previousDateString,
      values: savedValues,
      totalSheets: formTotalSheets,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };

    setRecords((prev) => {
      if (isEdit) {
        const exists = prev.some((r) => r.id === recordId);
        if (exists) {
          return prev.map((r) => (r.id === recordId ? newRecord : r));
        } else {
          return [newRecord, ...prev];
        }
      } else {
        return [newRecord, ...prev];
      }
    });

    // Un-exclude if it was previously excluded
    setExcludedIds((prev) => prev.filter((id) => id !== recordId));

    // Cloud SQL save
    apiService.saveDailyTeamOperation(newRecord);

    // If we are editing an existing record, first remove any prior linked stock records so the updated items replace them cleanly
    if (isEdit && editingRecordId) {
      const targetRecord = records.find((r) => r.id === editingRecordId);
      if (targetRecord) {
        const normTargetTeam = normalizeTeamName(targetRecord.teamName);
        const isTargetCea = targetRecord.categoryType === 'cEA';
        const targetDate = normalizeDateToISO(targetRecord.date);

        const linkedOldStockIds = (stockRecords || [])
          .filter((sr) => {
            const srTeam = normalizeTeamName(resolveRecordTeamName(sr) || sr.visaTeamRobokName || (sr as any).teamName || '');
            const srDate = normalizeDateToISO(sr.date || '');
            const srIsCea = sr.sourceFrom === 'cEA' || sr.remarks?.includes('cEA') || sr.visaType?.includes('cEA');
            return (
              srDate === targetDate &&
              srTeam === normTargetTeam &&
              srIsCea === isTargetCea &&
              (sr.operationType === 'useTeam' || sr.sourceFrom?.includes('ប្រើប្រាស់') || sr.id.startsWith('stock-dtr-') || sr.id.startsWith('dtr-'))
            );
          })
          .map((sr) => sr.id);

        if (linkedOldStockIds.length > 0) {
          if (onDeleteBatchStockRecords) {
            onDeleteBatchStockRecords(linkedOldStockIds);
          } else if (onDeleteStockRecord) {
            linkedOldStockIds.forEach((id) => onDeleteStockRecord(id));
          }
        }
      }
    }

    // Insert items with quantity > 0 into the main Visa Sheet Stock Data (ទិន្នន័យសន្លឹកទិដ្ឋាការ)
    if ((onBatchImportStockRecords || onAddStockRecord) && stockItemsToSave.length > 0) {
      if (onBatchImportStockRecords) {
        onBatchImportStockRecords(stockItemsToSave);
      } else if (onAddStockRecord) {
        stockItemsToSave.forEach((stk) => onAddStockRecord(stk));
      }
    }

    setEditingRecordId(null);
    localStorage.removeItem('app_daily_team_edit_payload');
    localStorage.removeItem('app_daily_team_edit_record_id');
    localStorage.removeItem('app_daily_team_auto_load_form');

    // Reset/empty input form after insert, keeping calendar (selectedDate) and team (selectedTeam) unchanged
    const emptyForm = getPriorFormEntriesForTeam(selectedTeam, selectedDate, selectedOption);
    setFormData(emptyForm);

    const batchNum = todayTeamRecords.length + (isEdit ? 0 : 1);
    onShowToast(
      isEdit
        ? `បានកែសម្រួលប្រតិបត្តិការ [${selectedOption}] ក្រុម ${selectedTeam} ជោគជ័យ!`
        : `បានរក្សាទុកប្រតិបត្តិការ [${selectedOption}] ក្រុម ${selectedTeam} លើកទី ${batchNum} (${formTotalSheets} សន្លឹក) ជោគជ័យ!`,
      'success'
    );
  };

  // Open Update Confirmation Modal
  const handleUpdateClick = () => {
    if (!selectedDate) {
      onShowToast('សូមជ្រើសរើសកាលបរិច្ឆេទ (ប្រចាំថ្ងៃ)', 'error');
      return;
    }
    if (!selectedTeam) {
      onShowToast('សូមជ្រើសរើសក្រុមជាមុនសិន', 'error');
      return;
    }

    setUpdateModalOpen(true);
  };

  // Execute Update after confirmation in modal
  const handleConfirmUpdate = () => {
    setUpdateModalOpen(false);

    const savedValues: any = {};
    const normTeam = normalizeTeamName(selectedTeam);
    const stockItemsToSave: StockRecord[] = [];

    VISA_TYPES.forEach((vt) => {
      const entries = formData[vt.id] || [];
      const totalQ = entries.reduce(
        (sum, e) => sum + (parseInt(e?.quantity || '', 10) || 0),
        0
      );
      const firstValid = entries.find((e) => e.startSerial?.trim());
      const lastValid = [...entries].reverse().find((e) => e.endSerial?.trim());
      
      const processedEntries = entries.map((e, eIdx) => {
        let oc = e.oldCode?.trim();
        if (!oc && eIdx > 0) {
          oc = entries[eIdx - 1]?.endSerial?.trim() || (e.startSerial ? decrementSerial(e.startSerial, 1) : '');
        } else if (!oc && e.startSerial) {
          oc = decrementSerial(e.startSerial, 1);
        }
        return {
          ...e,
          oldCode: oc || '',
        };
      });

      const oldCodeVal = processedEntries[0]?.oldCode || '';

      savedValues[vt.id] = {
        quantity: totalQ,
        startSerial: firstValid?.startSerial || '',
        endSerial: lastValid?.endSerial || '',
        oldCode: oldCodeVal,
        entries: processedEntries,
      };

      // Collect items for stock ledger
      entries.forEach((e) => {
        const q = parseInt(e?.quantity || '', 10) || 0;
        if (q > 0) {
          const robokTeamObj = categories?.visaTeamsRobok?.find(
            (t) => normalizeTeamName(t.name) === normTeam
          );

          stockItemsToSave.push({
            id: `stock-dtr-${selectedOption.toLowerCase()}-${selectedDate}-${selectedTeam.replace(/\s+/g, '-')}-${vt.id}-${Date.now()}-${Math.random().toString(36).substr(2, 4)}`,
            stockType: 'sticker',
            operationType: 'useTeam',
            date: selectedDate,
            time: new Date().toLocaleTimeString('en-US', { hour12: false, hour: '2-digit', minute: '2-digit' }),
            visaType: vt.name,
            quantityBundles: q,
            totalSheets: q,
            startSerial: e?.startSerial || '',
            endSerial: e?.endSerial || '',
            visaTeamRobokId: robokTeamObj?.id || `team-${selectedTeam}`,
            visaTeamRobokName: selectedTeam,
            sourceFrom: selectedOption === 'cEA' ? 'cEA' : 'Sticker',
            createdBy: userName || 'Admin',
            createdAt: new Date().toISOString(),
          });
        }
      });
    });

    const isEdit = Boolean(editingRecordId);
    const recordId =
      editingRecordId ||
      `dtr-${selectedOption}-${selectedDate}-${selectedTeam.replace(/\s+/g, '-')}`;
    const newRecord: DailyTeamRecord = {
      id: recordId,
      categoryType: selectedOption,
      date: selectedDate,
      teamName: selectedTeam,
      previousDate: previousDateString,
      values: savedValues,
      totalSheets: formTotalSheets,
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };

    setRecords((prev) => {
      if (isEdit) {
        const exists = prev.some((r) => r.id === recordId);
        if (exists) {
          return prev.map((r) => (r.id === recordId ? newRecord : r));
        } else {
          return [newRecord, ...prev];
        }
      } else {
        const filtered = prev.filter(
          (r) =>
            !(
              r.categoryType === selectedOption &&
              r.date === selectedDate &&
              (r.teamName === selectedTeam || normalizeTeamName(r.teamName) === normTeam)
            )
        );
        return [newRecord, ...filtered];
      }
    });

    // Un-exclude if it was previously excluded
    setExcludedIds((prev) => prev.filter((id) => id !== recordId));

    // Cloud SQL save
    apiService.saveDailyTeamOperation(newRecord);

    // Update in Visa Sheet Stock Data (ទិន្នន័យសន្លឹកទិដ្ឋាការ):
    // 1. Delete previous stock records linked to this team/date/option
    const targetDate = normalizeDateToISO(selectedDate);
    const isCea = selectedOption === 'cEA';
    const oldStockIdsToDelete = (stockRecords || [])
      .filter((sr) => {
        const srTeam = normalizeTeamName(resolveRecordTeamName(sr) || sr.visaTeamRobokName || (sr as any).teamName || '');
        const srDate = normalizeDateToISO(sr.date || '');
        const srIsCea = sr.sourceFrom === 'cEA' || sr.remarks?.includes('cEA') || sr.visaType?.includes('cEA');
        const matchesCat = isCea ? srIsCea : !srIsCea;

        return (
          srDate === targetDate &&
          srTeam === normTeam &&
          matchesCat &&
          (sr.operationType === 'useTeam' || sr.sourceFrom?.includes('ប្រើប្រាស់') || sr.id.startsWith('stock-dtr-') || sr.id.startsWith('dtr-'))
        );
      })
      .map((sr) => sr.id);

    if (oldStockIdsToDelete.length > 0) {
      if (onDeleteBatchStockRecords) {
        onDeleteBatchStockRecords(oldStockIdsToDelete);
      } else if (onDeleteStockRecord) {
        oldStockIdsToDelete.forEach((id) => onDeleteStockRecord(id));
      }
    }

    // 2. Insert new updated stock records
    if ((onBatchImportStockRecords || onAddStockRecord) && stockItemsToSave.length > 0) {
      if (onBatchImportStockRecords) {
        onBatchImportStockRecords(stockItemsToSave);
      } else if (onAddStockRecord) {
        stockItemsToSave.forEach((stk) => onAddStockRecord(stk));
      }
    }

    setEditingRecordId(null);

    // Reset/empty input form after update, keeping calendar (selectedDate) and team (selectedTeam) unchanged
    const emptyForm = getPriorFormEntriesForTeam(selectedTeam, selectedDate, selectedOption);
    setFormData(emptyForm);

    onShowToast(
      `បានកែសម្រួល (Update) ទិន្នន័យប្រតិបត្តិការប្រចាំថ្ងៃតាមក្រុម [${selectedOption}] ក្រុម ${selectedTeam} សរុប ${formTotalSheets} សន្លឹក រួចរាល់!`,
      'success'
    );
  };

  // Open Delete Confirmation Modal
  const handleDeleteClick = () => {
    if (!selectedDate) {
      onShowToast('សូមជ្រើសរើសកាលបរិច្ឆេទ (ប្រចាំថ្ងៃ)', 'error');
      return;
    }
    if (!selectedTeam) {
      onShowToast('សូមជ្រើសរើសក្រុមជាមុនសិន', 'error');
      return;
    }

    setDeleteModalOpen(true);
  };

  // Execute Deletion after confirmation in modal
  const handleConfirmDelete = () => {
    setDeleteModalOpen(false);
    const normTeam = normalizeTeamName(selectedTeam);
    const recordId = editingRecordId || `dtr-${selectedOption}-${selectedDate}-${selectedTeam.replace(/\s+/g, '-')}`;

    // 1. Delete from local records state
    setRecords((prev) =>
      prev.filter(
        (r) =>
          r.id !== recordId &&
          !(
            r.categoryType === selectedOption &&
            r.date === selectedDate &&
            (r.teamName === selectedTeam || normalizeTeamName(r.teamName) === normTeam)
          )
      )
    );

    // 2. Cloud SQL deletion
    apiService.deleteDailyTeamOperation(recordId);

    // 3. Delete linked stock records
    const targetDate = normalizeDateToISO(selectedDate);
    const stockIdsToDelete = (stockRecords || [])
      .filter((sr) => {
        const srTeam = normalizeTeamName(resolveRecordTeamName(sr) || sr.visaTeamRobokName || (sr as any).teamName || '');
        const srDate = normalizeDateToISO(sr.date || '');
        const srIsCea = sr.sourceFrom === 'cEA' || sr.remarks?.includes('cEA') || sr.visaType?.includes('cEA');
        const matchesCat = selectedOption === 'cEA' ? srIsCea : !srIsCea;

        return (
          srDate === targetDate &&
          srTeam === normTeam &&
          matchesCat &&
          (sr.operationType === 'useTeam' || sr.sourceFrom?.includes('ប្រើប្រាស់') || sr.id.startsWith('stock-dtr-') || sr.id.startsWith('dtr-'))
        );
      })
      .map((sr) => sr.id);

    if (stockIdsToDelete.length > 0) {
      if (onDeleteBatchStockRecords) {
        onDeleteBatchStockRecords(stockIdsToDelete);
      } else if (onDeleteStockRecord) {
        stockIdsToDelete.forEach((id) => onDeleteStockRecord(id));
      }
    }

    // 4. Mark excluded
    setExcludedIds((prev) => Array.from(new Set([...prev, recordId, ...stockIdsToDelete])));

    // 5. Reset form to empty
    const emptyForm = getPriorFormEntriesForTeam(selectedTeam, selectedDate, selectedOption);
    setFormData(emptyForm);
    setEditingRecordId(null);

    onShowToast(
      `បានលុបទិន្នន័យ [${selectedOption}] ក្រុម ${selectedTeam} សម្រាប់ថ្ងៃ ${selectedDate} រួចរាល់!`,
      'success'
    );
  };

  // Search button handler: loads existing data for selected date and team, automatically detecting whether it is Sticker or cEA
  const handleSearch = () => {
    if (!selectedDate) {
      onShowToast('សូមជ្រើសរើសកាលបរិច្ឆេទ (ប្រចាំថ្ងៃ)', 'error');
      return;
    }

    let targetTeam = selectedTeam;
    if (!targetTeam) {
      // Find first team with records for this date in either Sticker or cEA
      const anyWithData = allTeams.find((t) => {
        const matchCur = findRecordForDateTeamOption(selectedDate, t, selectedOption);
        if (matchCur !== null && matchCur.record && matchCur.record.totalSheets > 0) return true;
        const otherOpt: VisaCategoryOption = selectedOption === 'Sticker' ? 'cEA' : 'Sticker';
        const matchOther = findRecordForDateTeamOption(selectedDate, t, otherOpt);
        return matchOther !== null && matchOther.record && matchOther.record.totalSheets > 0;
      });
      if (anyWithData) {
        targetTeam = anyWithData;
        setSelectedTeam(anyWithData);
      } else {
        onShowToast('សូមជ្រើសរើសក្រុមដើម្បីស្វែងរកទិន្នន័យ', 'error');
        return;
      }
    }

    // Search for the team's data, checking BOTH options if necessary to auto-detect
    let match = findRecordForDateTeamOption(selectedDate, targetTeam, selectedOption);
    let matchedOption: VisaCategoryOption = selectedOption;

    if (!match || !match.record || !match.record.values) {
      const otherOpt: VisaCategoryOption = selectedOption === 'Sticker' ? 'cEA' : 'Sticker';
      const otherMatch = findRecordForDateTeamOption(selectedDate, targetTeam, otherOpt);
      if (otherMatch && otherMatch.record && otherMatch.record.values) {
        match = otherMatch;
        matchedOption = otherOpt;
        setSelectedOption(otherOpt);
      }
    }

    if (match && match.record && match.record.values) {
      const found = match.record;
      const updated: Record<string, FormRowEntry[]> = {};
      let totalFound = 0;
      VISA_TYPES.forEach((vt) => {
        const val = found.values[vt.id];
        const lastCode = getLastUsedOldCode(targetTeam, vt.id, selectedDate, matchedOption);
        const subEntries = val?.entries;
        if (Array.isArray(subEntries) && subEntries.length > 0) {
          updated[vt.id] = subEntries.map((e, idx) => ({
            id: e.id || `${vt.id}-${idx}`,
            quantity:
              e.quantity !== undefined && e.quantity !== '' && Number(e.quantity) > 0
                ? String(e.quantity)
                : '',
            startSerial: e.startSerial || '',
            endSerial: e.endSerial || '',
            oldCode: e.oldCode || (idx === 0 ? lastCode : ''),
          }));
          const totalQ = subEntries.reduce(
            (s, e) => s + (parseInt(e.quantity, 10) || 0),
            0
          );
          totalFound += totalQ;
        } else {
          const q =
            val?.quantity !== undefined && val?.quantity !== '' && Number(val.quantity) > 0
              ? String(val.quantity)
              : '';
          if (q) totalFound += Number(q);
          updated[vt.id] = [
            {
              id: `${vt.id}-0`,
              quantity: q,
              startSerial: val?.startSerial || '',
              endSerial: val?.endSerial || '',
              oldCode: val?.oldCode || lastCode || '',
            },
          ];
        }
      });
      setFormData(updated);
      onShowToast(
        `បានទាញយកទិន្នន័យ (Search) [${matchedOption}] ក្រុម ${targetTeam} សម្រាប់ថ្ងៃ ${selectedDate} សរុប ${totalFound} សន្លឹក`,
        'success'
      );
    } else {
      // Clean form with last used old code(s)
      const emptyForm = getPriorFormEntriesForTeam(targetTeam, selectedDate, selectedOption);
      setFormData(emptyForm);
      onShowToast(
        `មិនទាន់មានទិន្នន័យសម្រាប់ថ្ងៃ ${selectedDate} (Sticker ឬ cEA) ក្រុម ${targetTeam} ទេ`,
        'error'
      );
    }
  };

  // Aggregated data for selected date and selectedOption
  const dateRecords = useMemo(() => {
    const normDate = normalizeDateToISO(selectedDate);
    return effectiveRecords.filter(
      (r) => r.categoryType === selectedOption && normalizeDateToISO(r.date) === normDate
    );
  }, [effectiveRecords, selectedDate, selectedOption]);

  // Sticker Summary per team on selected date (Bottom Table) - Optimized with Map lookups
  const stickerTeamRowData = useMemo(() => {
    const normDate = normalizeDateToISO(selectedDate);
    const stickerRecordsMap = new Map<string, DailyTeamRecord>();
    effectiveRecords.forEach((r) => {
      if (r.categoryType === 'Sticker' && normalizeDateToISO(r.date) === normDate) {
        stickerRecordsMap.set(r.teamName, r);
        stickerRecordsMap.set(normalizeTeamName(r.teamName), r);
      }
    });

    return allTeams.map((teamName, idx) => {
      const normTeam = normalizeTeamName(teamName);
      const record = stickerRecordsMap.get(teamName) || stickerRecordsMap.get(normTeam);
      const rowCounts: { [vt: string]: number } = {};
      let total = 0;

      VISA_TYPES.forEach((vt) => {
        const q = Number(record?.values[vt.id]?.quantity) || 0;
        rowCounts[vt.id] = q;
        total += q;
      });

      return {
        index: idx + 1,
        teamName,
        counts: rowCounts,
        total,
      };
    });
  }, [allTeams, effectiveRecords, selectedDate]);

  // Display only Sticker teams - If User role, display only that team's row
  const displayedStickerRows = useMemo(() => {
    if (currentRole === 'User') {
      const targetTeam = selectedTeam || assignedTeam;
      if (!targetTeam) return [];
      const normTarget = normalizeTeamName(targetTeam);
      const userTeamRows = stickerTeamRowData.filter(
        (row) => row.teamName === targetTeam || normalizeTeamName(row.teamName) === normTarget
      );
      return userTeamRows.map((row, idx) => ({
        ...row,
        displayIndex: idx + 1,
      }));
    }

    // Secondary / Office: show all teams with usage (>0)
    return stickerTeamRowData
      .filter((row) => row.total > 0)
      .map((row, idx) => ({
        ...row,
        displayIndex: idx + 1,
      }));
  }, [stickerTeamRowData, currentRole, selectedTeam, assignedTeam]);

  // Overall totals across Sticker teams (Bottom Table)
  const stickerOverallTotals = useMemo(() => {
    const totals: { [vt: string]: number } = {};
    let grandTotal = 0;

    VISA_TYPES.forEach((vt) => {
      const sum = displayedStickerRows.reduce((acc, row) => acc + (row.counts[vt.id] || 0), 0);
      totals[vt.id] = sum;
      grandTotal += sum;
    });

    return { totals, grandTotal };
  }, [displayedStickerRows]);

  // Top Table: Key Summary for teams that have tick on (ក្រដាសអនុម័ត / cEA) in "ទិន្នន័យក្រុមផ្តល់ទិដ្ឋាការ"
  const topCeaRows = useMemo(() => {
    // If User role, only show if this specific team is allowed cEA, and only show this team's row
    if (currentRole === 'User') {
      if (!isTeamEVisaAllowed) {
        return [];
      }
      const targetTeam = selectedTeam || assignedTeam;
      if (!targetTeam) return [];
      const normTarget = normalizeTeamName(targetTeam);

      const normDate = normalizeDateToISO(selectedDate);
      const record = effectiveRecords.find(
        (r) =>
          r.categoryType === 'cEA' &&
          normalizeDateToISO(r.date) === normDate &&
          (r.teamName === targetTeam || normalizeTeamName(r.teamName) === normTarget)
      );

      const rowCounts: { [vt: string]: number } = {};
      let total = 0;

      VISA_TYPES.forEach((vt) => {
        const q = Number(record?.values[vt.id]?.quantity) || 0;
        rowCounts[vt.id] = q;
        total += q;
      });

      return [
        {
          index: 1,
          teamName: targetTeam,
          counts: rowCounts,
          total,
        },
      ];
    }

    // Secondary / Office: show all teams with cEA enabled
    const rawRobokTeams =
      categories?.visaTeamsRobok && categories.visaTeamsRobok.length > 0
        ? categories.visaTeamsRobok
        : DEFAULT_TEAMS_LIST.map((name, i) => ({ id: `vtr-${i + 1}`, name, createdAt: '' }));

    const ceaTeams = rawRobokTeams.filter((team, idx) => {
      if (isTeamCeaEnabled(team.name, idx)) return true;
      const nName = normalizeTeamName(team.name);
      return (
        effectiveRecords.some(
          (r) => r.categoryType === 'cEA' && (r.teamName === team.name || normalizeTeamName(r.teamName) === nName)
        ) ||
        (stockRecords || []).some(
          (sr) => isCeaRecord(sr) && (resolveRecordTeamName(sr) === team.name || normalizeTeamName(resolveRecordTeamName(sr)) === nName)
        )
      );
    });

    const normDate = normalizeDateToISO(selectedDate);
    const ceaRecordsMap = new Map<string, DailyTeamRecord>();
    effectiveRecords.forEach((r) => {
      if (r.categoryType === 'cEA' && normalizeDateToISO(r.date) === normDate) {
        ceaRecordsMap.set(r.teamName, r);
        ceaRecordsMap.set(normalizeTeamName(r.teamName), r);
      }
    });

    return ceaTeams.map((team, idx) => {
      const name = team.name.trim();
      const normTeam = normalizeTeamName(name);
      const record = ceaRecordsMap.get(name) || ceaRecordsMap.get(normTeam);
      const rowCounts: { [vt: string]: number } = {};
      let total = 0;

      VISA_TYPES.forEach((vt) => {
        const q = Number(record?.values[vt.id]?.quantity) || 0;
        rowCounts[vt.id] = q;
        total += q;
      });

      return {
        index: idx + 1,
        teamName: name,
        counts: rowCounts,
        total,
      };
    });
  }, [
    currentRole,
    isTeamEVisaAllowed,
    selectedTeam,
    assignedTeam,
    categories?.visaTeamsRobok,
    categories?.visaTeams,
    teamTypeOptions,
    effectiveRecords,
    selectedDate,
  ]);

  const topCeaTotals = useMemo(() => {
    const totals: { [vt: string]: number } = {};
    let grandTotal = 0;

    VISA_TYPES.forEach((vt) => {
      const sum = topCeaRows.reduce((acc, row) => acc + (row.counts[vt.id] || 0), 0);
      totals[vt.id] = sum;
      grandTotal += sum;
    });

    return { totals, grandTotal };
  }, [topCeaRows]);

  // Calculate remaining stock for the selected team up to selectedDate
  // Formula: Old Stock (30-Nov-2018) + Issued from K2 - Used - Transferred - Damaged/Missing/Returned
  const selectedTeamRemainingStock = useMemo(() => {
    if (!selectedTeam) {
      return { counts: {} as Record<string, number>, total: 0 };
    }

    const calculated = calculateAllTeamsStickerStockAtDate(
      stockRecords || [],
      selectedDate,
      allTeams,
      excludedIds
    );

    const normSelected = normalizeTeamName(selectedTeam);
    const matchedKey =
      Object.keys(calculated.remainingAfterMatrix).find(
        (k) => k === selectedTeam || normalizeTeamName(k) === normSelected
      ) || selectedTeam;

    const teamCounts = calculated.remainingAfterMatrix[matchedKey] || {};
    const counts: Record<string, number> = {};
    let total = 0;

    VISA_TYPES.forEach((vt) => {
      const val = teamCounts[vt.id] || 0;
      counts[vt.id] = val;
      total += val;
    });

    return { counts, total };
  }, [stockRecords, selectedDate, selectedTeam, allTeams, excludedIds, selectedOption]);

  // Drill-down data breakdown for selected visa / team
  const drillDownDetails = useMemo(() => {
    if (!drillDownVisa || !drillDownVisa.teamName) return null;
    const { teamName, visaType } = drillDownVisa;
    const normTeam = normalizeTeamName(teamName);
    const excludedSet = new Set(excludedIds);

    const isTotal = visaType === 'TOTAL';
    const targetVts = isTotal
      ? VISA_TYPES.map((v) => v.id)
      : [visaType.toUpperCase()];

    // Baseline from 30-Nov-2018 benchmark figures
    let baselineSum = 0;
    const baselineByVt: { [vt: string]: number } = {};
    targetVts.forEach((vt) => {
      const bVal =
        DEFAULT_OPENING_MATRIX[teamName]?.[vt] ??
        DEFAULT_OPENING_MATRIX[normTeam]?.[vt] ??
        0;
      baselineByVt[vt] = bVal;
      baselineSum += bVal;
    });

    const contributingRecords: Array<{
      id: string;
      date: string;
      opType: string;
      opTitle: string;
      visaType: string;
      startSerial?: string;
      endSerial?: string;
      quantity: number;
      impact: '+' | '-';
      category: 'opening' | 'issued' | 'used' | 'damaged' | 'transferred';
      source: 'stock_ledger' | 'daily_form' | 'baseline';
      isExcluded: boolean;
      remarks?: string;
    }> = [];

    // Scan Stock Ledger records
    (stockRecords || []).forEach((sr) => {
      if (sr.stockType && sr.stockType !== 'sticker') return;

      const vt = normalizeVisaType(sr.visaType).toUpperCase();
      if (!targetVts.includes(vt)) return;

      const qty =
        Number(sr.totalSheets) ||
        (Number(sr.quantityBundles) ? Number(sr.quantityBundles) * 50 : 0) ||
        Number((sr as any).quantity) ||
        0;
      if (qty <= 0) return;

      const rawTeam = resolveRecordTeamName(sr);
      const isSender = isRecordForTeam(sr, teamName);
      const isRecipient = isRecordForRecipientTeam(sr, teamName);
      const isTrans = isTransferTeamRecord(sr);

      if (!isSender && !isRecipient) return;

      const isEx = excludedSet.has(sr.id);
      const isOld = isOldStockTeamRecord(sr);
      const isDamaged =
        sr.operationType === 'damagedTeam' ||
        sr.operationType === 'damaged' ||
        sr.operationType === 'missingTeam' ||
        sr.operationType === 'returnTeam';
      const isUse =
        !isOld &&
        !isTrans &&
        !isDamaged &&
        (sr.operationType === 'useTeam' || sr.sourceFrom?.includes('ប្រើប្រាស់'));
      const isIssue =
        !isOld &&
        !isTrans &&
        !isDamaged &&
        !isUse &&
        (sr.operationType === 'issueTeam' || isIssueTeamRecord(sr));

      const normRecDate = normalizeDateToISO(sr.date || '');
      const normSelDate = normalizeDateToISO(selectedDate);

      if (isOld) {
        if (!isSender) return;
        contributingRecords.push({
          id: sr.id,
          date: sr.date || '30-Nov-2018',
          opType: sr.operationType || 'oldstockteam',
          opTitle: 'ស្តុកចាស់ក្រុម (បញ្ចូលក្នុងសៀវភៅស្តុក)',
          visaType: vt,
          startSerial: sr.startSerial,
          endSerial: sr.endSerial,
          quantity: qty,
          impact: '+',
          category: 'opening',
          source: 'stock_ledger',
          isExcluded: isEx,
          remarks: sr.remarks || sr.sourceFrom || '',
        });
      } else if (isTrans && isRecipient) {
        // ផ្ទេរពីក្រុម (ទទួលពីក្រុមផ្សេង) -> Count under "បើកផ្តល់ពីក២"
        if (normRecDate && normSelDate && normRecDate > normSelDate) return;
        contributingRecords.push({
          id: `${sr.id}-recip`,
          date: sr.date || selectedDate,
          opType: 'transferTeam',
          opTitle: `បើកផ្តល់ពីក២ (ផ្ទេរពី ${rawTeam || 'ក្រុម'})`,
          visaType: vt,
          startSerial: sr.startSerial,
          endSerial: sr.endSerial,
          quantity: qty,
          impact: '+',
          category: 'issued',
          source: 'stock_ledger',
          isExcluded: isEx,
          remarks: sr.remarks || sr.sourceFrom || '',
        });
      } else if (isIssue && isSender) {
        // ទទួលពីការិយាល័យ (បើកពីក២) -> Count under "បើកផ្តល់ពីក២"
        if (normRecDate && normSelDate && normRecDate > normSelDate) return;
        contributingRecords.push({
          id: sr.id,
          date: sr.date || selectedDate,
          opType: 'issueTeam',
          opTitle: 'បើកផ្តល់ពីក២ (ទទួលពីការិយាល័យ)',
          visaType: vt,
          startSerial: sr.startSerial,
          endSerial: sr.endSerial,
          quantity: qty,
          impact: '+',
          category: 'issued',
          source: 'stock_ledger',
          isExcluded: isEx,
          remarks: sr.remarks || sr.sourceFrom || '',
        });
      } else if (isUse && isSender) {
        if (normRecDate && normSelDate && normRecDate > normSelDate) return;
        contributingRecords.push({
          id: sr.id,
          date: sr.date || selectedDate,
          opType: 'useTeam',
          opTitle: 'ប្រើប្រាស់ប្រចាំថ្ងៃ (សៀវភៅស្តុក)',
          visaType: vt,
          startSerial: sr.startSerial,
          endSerial: sr.endSerial,
          quantity: qty,
          impact: '-',
          category: 'used',
          source: 'stock_ledger',
          isExcluded: isEx,
          remarks: sr.remarks || sr.sourceFrom || '',
        });
      } else if (isDamaged && isSender) {
        if (normRecDate && normSelDate && normRecDate > normSelDate) return;
        contributingRecords.push({
          id: sr.id,
          date: sr.date || selectedDate,
          opType: sr.operationType || 'damagedTeam',
          opTitle: 'ខូច/បាត់/ប្រគល់ត្រឡប់',
          visaType: vt,
          startSerial: sr.startSerial,
          endSerial: sr.endSerial,
          quantity: qty,
          impact: '-',
          category: 'damaged',
          source: 'stock_ledger',
          isExcluded: isEx,
          remarks: sr.remarks || '',
        });
      } else if (isTrans && isSender && !isRecipient) {
        // ផ្ទេរចេញទៅក្រុមផ្សេង
        if (normRecDate && normSelDate && normRecDate > normSelDate) return;
        const recip = (sr.recipientTeamName || (sr as any).destinationTo || (sr as any).targetTeam || '').trim();
        contributingRecords.push({
          id: sr.id,
          date: sr.date || selectedDate,
          opType: 'transferTeam',
          opTitle: `ផ្ទេរការប្រើប្រាស់ (ផ្ទេរទៅ ${recip || 'ក្រុមផ្សេង'})`,
          visaType: vt,
          startSerial: sr.startSerial,
          endSerial: sr.endSerial,
          quantity: qty,
          impact: '-',
          category: 'transferred',
          source: 'stock_ledger',
          isExcluded: isEx,
          remarks: sr.remarks || '',
        });
      }
    });

    // Scan Daily Form records
    records.forEach((dr) => {
      if (dr.teamName !== teamName && normalizeTeamName(dr.teamName) !== normTeam) return;
      if (dr.categoryType !== 'Sticker') return;
      const normRecDate = normalizeDateToISO(dr.date || '');
      const normSelDate = normalizeDateToISO(selectedDate);
      if (normRecDate && normSelDate && normRecDate > normSelDate) return;

      targetVts.forEach((vt) => {
        const valObj = dr.values?.[vt];
        const qty = Number(valObj?.quantity) || 0;
        if (qty <= 0) return;

        const isEx = excludedSet.has(dr.id);
        contributingRecords.push({
          id: dr.id,
          date: dr.date,
          opType: 'useDailyForm',
          opTitle: 'ប្រើប្រាស់ប្រចាំថ្ងៃ (ទម្រង់ Form)',
          visaType: vt,
          startSerial: valObj?.startSerial,
          endSerial: valObj?.endSerial,
          quantity: qty,
          impact: '-',
          category: 'used',
          source: 'daily_form',
          isExcluded: isEx,
          remarks: 'បានកត់ត្រាក្នុងទម្រង់ប្រចាំថ្ងៃ',
        });
      });
    });

    const activeOldStockRecords = contributingRecords.filter(
      (r) => r.category === 'opening' && !r.isExcluded
    );
    const recordedOldSum = activeOldStockRecords.reduce(
      (acc, r) => acc + r.quantity,
      0
    );

    // If recorded old stock exists, it is used for opening; otherwise baseline
    const effectiveOpening =
      recordedOldSum > 0 ? recordedOldSum : baselineSum;

    const totalIssued = contributingRecords
      .filter((r) => r.category === 'issued' && !r.isExcluded)
      .reduce((acc, r) => acc + r.quantity, 0);

    const totalUsed = contributingRecords
      .filter((r) => r.category === 'used' && !r.isExcluded)
      .reduce((acc, r) => acc + r.quantity, 0);

    const totalDamaged = contributingRecords
      .filter(
        (r) =>
          (r.category === 'damaged' || r.category === 'transferred') &&
          !r.isExcluded
      )
      .reduce((acc, r) => acc + r.quantity, 0);

    const calculatedBalance =
      effectiveOpening + totalIssued - totalUsed - totalDamaged;

    return {
      teamName,
      visaType,
      isTotal,
      baselineSum,
      baselineByVt,
      recordedOldSum,
      activeOldStockRecordsCount: activeOldStockRecords.length,
      effectiveOpening,
      totalIssued,
      totalUsed,
      totalDamaged,
      calculatedBalance,
      contributingRecords,
      hasPotentialDuplicate:
        recordedOldSum > 0 &&
        baselineSum > 0 &&
        (recordedOldSum === baselineSum * 2 ||
          recordedOldSum === baselineSum ||
          activeOldStockRecords.length > 1),
    };
  }, [drillDownVisa, stockRecords, records, selectedDate, excludedIds]);

  // Export to Excel
  const handleExportExcel = () => {
    const headers = [
      'ល.រ',
      'ក្រុមផ្តល់ទិដ្ឋាការ',
      'T',
      'T1',
      'T2',
      'T3',
      'E',
      'E1',
      'E2',
      'E3',
      'D',
      'K',
      'A',
      'B',
      'C',
      'TOTAL',
    ];

    const data = displayedStickerRows.map((row) => [
      row.displayIndex,
      row.teamName,
      row.counts['T'] || 0,
      row.counts['T1'] || 0,
      row.counts['T2'] || 0,
      row.counts['T3'] || 0,
      row.counts['E'] || 0,
      row.counts['E1'] || 0,
      row.counts['E2'] || 0,
      row.counts['E3'] || 0,
      row.counts['D'] || 0,
      row.counts['K'] || 0,
      row.counts['A'] || 0,
      row.counts['B'] || 0,
      row.counts['C'] || 0,
      row.total,
    ]);

    // Add total row
    data.push([
      'សរុប',
      `សរុបក្រុមផ្តល់ទិដ្ឋាការ (Sticker)`,
      stickerOverallTotals.totals['T'] || 0,
      stickerOverallTotals.totals['T1'] || 0,
      stickerOverallTotals.totals['T2'] || 0,
      stickerOverallTotals.totals['T3'] || 0,
      stickerOverallTotals.totals['E'] || 0,
      stickerOverallTotals.totals['E1'] || 0,
      stickerOverallTotals.totals['E2'] || 0,
      stickerOverallTotals.totals['E3'] || 0,
      stickerOverallTotals.totals['D'] || 0,
      stickerOverallTotals.totals['K'] || 0,
      stickerOverallTotals.totals['A'] || 0,
      stickerOverallTotals.totals['B'] || 0,
      stickerOverallTotals.totals['C'] || 0,
      stickerOverallTotals.grandTotal,
    ]);

    const ws = XLSX.utils.aoa_to_sheet([
      [`របាយការណ៍ប្រតិបត្តិការប្រចាំថ្ងៃតាមក្រុម [Sticker] — កាលបរិច្ឆេទ៖ ${selectedDate}`],
      [],
      headers,
      ...data,
    ]);

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, `Daily_Sticker`);
    XLSX.writeFile(wb, `Daily_Team_Operations_Sticker_${selectedDate}.xlsx`);
    onShowToast(`បានទាញយកឯកសារ Excel សម្រាប់ [Sticker] រួចរាល់!`, 'success');
  };

  const handleUserTeamChange = (newTeam: string) => {
    setSelectedTeam(newTeam);
    localStorage.removeItem('app_daily_team_edit_payload');
    localStorage.removeItem('app_daily_team_edit_record_id');
    localStorage.removeItem('app_daily_team_auto_load_form');
    // Clear form when user manually changes team
    const emptyForm: Record<string, FormRowEntry[]> = {};
    VISA_TYPES.forEach((vt) => {
      emptyForm[vt.id] = [{ id: `${vt.id}-0`, quantity: '', startSerial: '', endSerial: '', oldCode: '' }];
    });
    setFormData(emptyForm);
    setEditingRecordId(null);
  };

  return (
    <div className="w-full text-[#1E293B] font-sans">
      {/* Main Layout: Fixed Width Form on Left, Flex-1 Full Width Tables on Right */}
      <div className="flex flex-col xl:flex-row items-start gap-3 w-full">
        {/* LEFT COLUMN: Input Form (Optimized Width for Clean, Smooth Display) */}
        <div className="w-full xl:w-[490px] 2xl:w-[510px] shrink-0 flex flex-col">
          <div className="bg-[#BFDBFE] border-2 border-[#1E40AF] rounded shadow-md overflow-hidden">
            {/* Header: Action Buttons & Option Selector (No Manu text) */}
            <div className="bg-[#93C5FD] border-b-2 border-[#1E40AF] py-1.5 px-2.5 flex items-center justify-between gap-2">
              {/* 4 Action Buttons (Longer, Proportional & Equal) */}
              <div className="flex items-center gap-1.5 sm:gap-2">
                {/* Database (Red) */}
                <button
                  type="button"
                  onClick={handleSaveRecord}
                  className="w-14 sm:w-16 h-7.5 bg-[#DC2626] hover:bg-[#B91C1C] active:scale-95 text-white rounded flex items-center justify-center cursor-pointer shadow-xs transition"
                  title="បញ្ចូលទិន្នន័យ (Database / Save)"
                  aria-label="Database"
                >
                  <Database className="w-4 h-4" />
                </button>

                {/* Search (Blue) */}
                <button
                  type="button"
                  onClick={handleSearch}
                  className="w-14 sm:w-16 h-7.5 bg-[#2563EB] hover:bg-[#1D4ED8] active:scale-95 text-white rounded flex items-center justify-center cursor-pointer shadow-xs transition"
                  title="ស្វែងរកទិន្នន័យ (Search)"
                  aria-label="Search"
                >
                  <Search className="w-4 h-4" />
                </button>

                {/* Update (Magenta/Purple) */}
                <button
                  type="button"
                  onClick={handleUpdateClick}
                  className="w-14 sm:w-16 h-7.5 bg-[#C026D3] hover:bg-[#A21CAF] active:scale-95 text-white rounded flex items-center justify-center cursor-pointer shadow-xs transition"
                  title="កែសម្រួលទិន្នន័យ (Update)"
                  aria-label="Update"
                >
                  <Edit3 className="w-4 h-4" />
                </button>

                {/* Delete (Green) */}
                <button
                  type="button"
                  onClick={handleDeleteClick}
                  className="w-14 sm:w-16 h-7.5 bg-[#16A34A] hover:bg-[#15803D] active:scale-95 text-white rounded flex items-center justify-center cursor-pointer shadow-xs transition"
                  title="លុបទិន្នន័យ (Delete)"
                  aria-label="Delete"
                >
                  <Trash2 className="w-4 h-4" />
                </button>
              </div>

              {/* Option Selector: Sticker / cEA */}
              <div className="flex items-center gap-1 bg-white/90 p-0.5 rounded border border-blue-300 shadow-2xs">
                <button
                  type="button"
                  onClick={() => setSelectedOption('Sticker')}
                  className={`px-3 py-1 rounded text-xs font-bold font-times transition cursor-pointer ${
                    selectedOption === 'Sticker'
                      ? 'bg-[#1E40AF] text-white shadow-xs'
                      : 'text-blue-900 hover:bg-blue-100'
                  }`}
                >
                  Sticker
                </button>
                {isTeamEVisaAllowed && (
                  <button
                    type="button"
                    onClick={() => setSelectedOption('cEA')}
                    className={`px-3 py-1 rounded text-xs font-bold font-times transition cursor-pointer ${
                      selectedOption === 'cEA'
                        ? 'bg-[#0E7490] text-white shadow-xs'
                        : 'text-cyan-900 hover:bg-cyan-100'
                    }`}
                  >
                    cEA
                  </button>
                )}
              </div>
            </div>

            {/* Date & Group Selector Bar (Harmonized same-width inputs) */}
            <div className="grid grid-cols-12 border-b border-[#1E40AF] bg-[#DBEAFE] text-xs">
              {/* Date & Team Selection Column */}
              <div className="col-span-7 p-2 border-r border-[#1E40AF] flex flex-col justify-center space-y-1.5">
                <div className="flex items-center gap-2">
                  <span className="font-bold text-[#1E3A8A] whitespace-nowrap text-xs w-14">ប្រចាំថ្ងៃ</span>
                  <div className="flex-1">
                    <CustomDatePicker
                      value={selectedDate}
                      onChange={(dateStr) => setSelectedDate(dateStr)}
                      className="text-xs text-center font-bold font-times text-[#0F172A] py-1"
                    />
                  </div>
                </div>
                <div className="flex items-center gap-2">
                  <span className="font-bold text-[#1E3A8A] text-xs w-14">ក្រុម</span>
                  <select
                    value={selectedTeam}
                    disabled={currentRole === 'User' && Boolean(assignedTeam)}
                    onChange={(e) => handleUserTeamChange(e.target.value)}
                    className={`flex-1 border border-gray-400 rounded px-2 py-1 text-xs font-semibold shadow-inner outline-hidden focus:border-blue-600 truncate ${
                      currentRole === 'User' && assignedTeam
                        ? 'bg-gray-100 text-blue-950 cursor-not-allowed font-bold'
                        : 'bg-white text-[#0F172A]'
                    }`}
                  >
                    <option value="">(ជ្រើសរើស)</option>
                    {availableTeamsForSelection.map((t) => (
                      <option key={t} value={t}>
                        {t}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              {/* Right Side Previous Code Box */}
              <div className="col-span-5 bg-[#0E7490] text-white text-center p-2 flex flex-col justify-center">
                <div className="text-sm font-bold font-times">{formTotalSheets} <span className="font-siemreap font-bold">សន្លឹក</span></div>
                <div className="text-[11px] font-medium font-times border-t border-b border-cyan-400/40 my-0.5 py-0.5">
                  {displayPriorDate}
                </div>
                <div className="text-[11px] font-bold text-cyan-100">លេខកូដចាស់ថ្ងៃចុងក្រោយ</div>
              </div>
            </div>

            {/* Column Headers for Form Textboxes: For cEA show only [ប្រភេទ, ចំនួន], For Sticker show all 4 textboxes */}
            {selectedOption === 'cEA' ? (
              <div className="grid grid-cols-[80px_1fr] gap-2 items-center bg-[#0E7490] text-white px-3 py-1.5 text-xs font-bold text-center border-b border-[#085a6e]">
                <div className="text-center">ប្រភេទ</div>
                <div className="text-center">ចំនួន (សន្លឹក)</div>
              </div>
            ) : (
              <div className="grid grid-cols-[36px_0.85fr_1.25fr_1.25fr_1.65fr] gap-1 items-center bg-[#1E3A8A] text-white px-1.5 py-1.5 text-[11px] font-bold text-center border-b border-blue-900">
                <div className="text-center">ប្រភេទ</div>
                <div className="text-center">ចំនួន</div>
                <div className="text-center">ចាប់ពីលេខ</div>
                <div className="text-center">ដល់លេខ</div>
                <div className="text-center">លេខកូដចាស់</div>
              </div>
            )}

            {/* Form Rows Table for Visa Types */}
            <div className="p-1.5 bg-[#E2E8F0] space-y-1">
              {VISA_TYPES.map((vt, rowIndex) => {
                const entries = formData[vt.id] || [
                  { id: `${vt.id}-0`, quantity: '', startSerial: '', endSerial: '', oldCode: '' },
                ];

                if (selectedOption === 'cEA') {
                  return (
                    <div key={vt.id} className="space-y-1">
                      {entries.map((entry, entryIndex) => (
                        <div
                          key={entry.id || `${vt.id}-${entryIndex}`}
                          className="grid grid-cols-[80px_1fr_auto] gap-1.5 items-center bg-white/95 px-2.5 py-1 rounded border border-gray-300 shadow-2xs"
                        >
                          {/* Visa Type Label */}
                          <div className="text-center flex items-center justify-center gap-1">
                            <span
                              className={`text-base font-extrabold font-times ${
                                vt.isRed ? 'text-red-600' : 'text-[#1E3A8A]'
                              }`}
                            >
                              {vt.name}
                              {entries.length > 1 && (
                                <span className="text-[10px] ml-0.5 text-blue-600 font-times">
                                  #{entryIndex + 1}
                                </span>
                              )}
                            </span>
                            <span className="text-[10px] text-gray-500 font-semibold">ទិដ្ឋាការ</span>
                          </div>

                          {/* 1. ចំនួន (Quantity only for cEA) */}
                          <div className="w-full">
                            <input
                              id={`form-input-${rowIndex}-${entryIndex}-0`}
                              type="text"
                              placeholder="បញ្ចូលចំនួនសន្លឹក..."
                              value={entry.quantity}
                              onChange={(e) =>
                                handleInputChange(vt.id, entryIndex, 'quantity', e.target.value)
                              }
                              onKeyDown={(e) => handleKeyDown(e, rowIndex, entryIndex, 0)}
                              className="w-full h-8 bg-white border border-gray-400 rounded px-3 text-sm text-center font-bold font-times text-blue-900 focus:bg-yellow-50 focus:border-blue-500 outline-hidden shadow-2xs placeholder:font-normal placeholder:tracking-normal placeholder:font-kantumruy placeholder:text-gray-400"
                            />
                          </div>

                          {/* Add / Remove buttons for cEA */}
                          <div className="flex items-center gap-0.5">
                            <button
                              type="button"
                              onClick={() => handleAddRow(vt.id, entryIndex)}
                              className="h-8 w-7 bg-[#16A34A] hover:bg-[#15803D] active:scale-95 text-white flex items-center justify-center rounded border border-green-800 cursor-pointer shadow-2xs transition"
                              title="ចុច (+) ដើម្បីបន្ថែមជួរទិន្នន័យ"
                            >
                              <Plus className="w-3.5 h-3.5 stroke-[3]" />
                            </button>
                            {entries.length > 1 && (
                              <button
                                type="button"
                                onClick={() => handleRemoveRow(vt.id, entryIndex)}
                                className="h-8 w-6 bg-red-600 hover:bg-red-700 active:scale-95 text-white flex items-center justify-center rounded border border-red-800 cursor-pointer shadow-2xs transition"
                                title="លុបជួរនេះ"
                              >
                                <Trash2 className="w-3 h-3" />
                              </button>
                            )}
                          </div>
                        </div>
                      ))}
                    </div>
                  );
                }

                return (
                  <div key={vt.id} className="space-y-1">
                    {entries.map((entry, entryIndex) => (
                      <div
                        key={entry.id || `${vt.id}-${entryIndex}`}
                        className="grid grid-cols-[36px_0.85fr_1.25fr_1.25fr_1.65fr] gap-1 items-center bg-white/95 px-1 py-1 rounded border border-gray-300 shadow-2xs"
                      >
                        {/* Visa Type Label */}
                        <div className="text-center">
                          <span
                            className={`text-xs font-extrabold font-times ${
                              vt.isRed ? 'text-red-600' : 'text-[#1E3A8A]'
                            }`}
                          >
                            {vt.name}
                          </span>
                          {entries.length > 1 && (
                            <span className="text-[9px] block text-blue-600 font-bold font-times leading-none mx-auto">
                              #{entryIndex + 1}
                            </span>
                          )}
                        </div>

                        {/* 1. ចំនួន (Quantity) */}
                        <div className="w-full">
                          <input
                            id={`form-input-${rowIndex}-${entryIndex}-0`}
                            type="text"
                            placeholder="ចំនួន"
                            value={entry.quantity}
                            onChange={(e) =>
                              handleInputChange(vt.id, entryIndex, 'quantity', e.target.value)
                            }
                            onKeyDown={(e) => handleKeyDown(e, rowIndex, entryIndex, 0)}
                            className="w-full h-7 bg-white border border-gray-400 rounded px-1 text-xs text-center font-bold font-times text-blue-900 focus:bg-yellow-50 focus:border-blue-500 outline-hidden shadow-2xs placeholder:font-normal placeholder:tracking-normal placeholder:font-kantumruy placeholder:text-gray-400"
                          />
                        </div>

                        {/* 2. ចាប់ពីលេខ (Start Serial) */}
                        <div className="w-full">
                          <input
                            id={`form-input-${rowIndex}-${entryIndex}-1`}
                            type="text"
                            placeholder="ពីលេខ"
                            value={entry.startSerial}
                            onChange={(e) =>
                              handleInputChange(vt.id, entryIndex, 'startSerial', e.target.value)
                            }
                            onKeyDown={(e) => handleKeyDown(e, rowIndex, entryIndex, 1)}
                            className="w-full h-7 bg-white border border-gray-400 rounded px-1 text-xs text-center font-bold font-times text-gray-900 focus:bg-yellow-50 focus:border-blue-500 outline-hidden shadow-2xs tracking-tight placeholder:tracking-normal placeholder:font-normal placeholder:font-kantumruy placeholder:text-gray-400"
                          />
                        </div>

                        {/* 3. ដល់លេខ (End Serial - Auto show value) */}
                        <div className="w-full">
                          <input
                            id={`form-input-${rowIndex}-${entryIndex}-2`}
                            type="text"
                            placeholder="ដល់លេខ"
                            value={entry.endSerial}
                            onChange={(e) =>
                              handleInputChange(vt.id, entryIndex, 'endSerial', e.target.value)
                            }
                            onKeyDown={(e) => handleKeyDown(e, rowIndex, entryIndex, 2)}
                            className="w-full h-7 bg-white border border-gray-400 rounded px-1 text-xs text-center font-bold font-times text-gray-900 focus:bg-yellow-50 focus:border-blue-500 outline-hidden shadow-2xs tracking-tight placeholder:tracking-normal placeholder:font-normal placeholder:font-kantumruy placeholder:text-gray-400"
                          />
                        </div>

                        {/* 4. លេខកូដចាស់ (Old Code) with Button Icon (+) on the right - Clean, Full Display, No Red Delete Button */}
                        <div className="w-full flex items-center">
                          <input
                            id={`form-input-${rowIndex}-${entryIndex}-3`}
                            type="text"
                            placeholder="កូដចាស់"
                            value={entry.oldCode}
                            onClick={() =>
                              handleApplyOldCodeToStartSerial(vt.id, entryIndex, rowIndex)
                            }
                            onChange={(e) =>
                              handleInputChange(vt.id, entryIndex, 'oldCode', e.target.value)
                            }
                            onKeyDown={(e) => handleKeyDown(e, rowIndex, entryIndex, 3)}
                            className="w-full min-w-0 h-7 bg-[#15803D] hover:bg-[#166534] active:bg-[#14532D] text-white rounded-l px-1.5 text-xs font-bold font-times text-center border border-green-800 outline-hidden shadow-2xs tracking-tight transition select-all focus:bg-green-700 cursor-pointer placeholder:tracking-normal placeholder:font-normal placeholder:font-kantumruy placeholder:text-emerald-100"
                            title={
                              entry.oldCode
                                ? `លេខកូដចាស់: ${entry.oldCode} (ចុចលើនេះដើម្បីកំណត់ [ចាប់ពីលេខ] = ${incrementSerial(
                                    entry.oldCode,
                                    1
                                  )})`
                                : 'លេខកូដចាស់ (ចុចដើម្បីទាញយក)'
                            }
                          />
                          <button
                            type="button"
                            onClick={() => handleAddRow(vt.id, entryIndex)}
                            className={`h-7 px-1.5 bg-[#16A34A] hover:bg-[#15803D] active:scale-95 text-white flex items-center justify-center border border-l-0 border-green-800 cursor-pointer shadow-2xs shrink-0 transition ${
                              entries.length > 1 ? '' : 'rounded-r'
                            }`}
                            title={
                              entry.oldCode || entry.endSerial
                                ? `ចុច (+) ដើម្បីបន្ថែមជួរទិន្នន័យថ្មីសម្រាប់ប្រភេទ ${vt.name}`
                                : `ចុច (+) ដើម្បីបន្ថែមជួរថ្មី`
                            }
                          >
                            <Plus className="w-3.5 h-3.5 stroke-[3]" />
                          </button>
                          {entries.length > 1 && (
                            <button
                              type="button"
                              onClick={() => handleRemoveRow(vt.id, entryIndex)}
                              className="h-7 px-1.5 bg-[#DC2626] hover:bg-[#B91C1C] active:scale-95 text-white flex items-center justify-center rounded-r border border-l-0 border-red-800 cursor-pointer shadow-2xs shrink-0 transition"
                              title={`លុបជួរ #${entryIndex + 1} នេះចេញ`}
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                            </button>
                          )}
                        </div>
                      </div>
                    ))}
                  </div>
                );
              })}
            </div>
          </div>
        </div>

        {/* RIGHT COLUMN: Operations Tables (Expands to Full Width) */}
        <div className="flex-1 min-w-0 space-y-3 w-full">
          {/* Top Key Airport/Port Table (Fixed to cEA) - Only shown if allowed or secondary */}
          {(currentRole === 'Secondary' || isTeamEVisaAllowed) && topCeaRows.length > 0 && (
            <div className="bg-white border-2 border-[#1E40AF] rounded overflow-x-auto">
              <table className="w-full text-xs text-center border-collapse table-fixed min-w-[1020px]">
                <colgroup>
                  <col className="w-[36px] min-w-[36px]" />
                  <col className="w-[115px] min-w-[115px]" />
                  {VISA_TYPES.map((vt) => (
                    <col key={vt.id} className="w-[62px] min-w-[62px]" />
                  ))}
                  <col className="w-[68px] min-w-[68px]" />
                </colgroup>
                <thead>
                  {/* 1. Header Column Titles */}
                  <tr className="bg-[#0F4C81] text-white text-xs font-bold">
                    <th className="p-1.5 border border-blue-900 w-[36px]">ល.រ</th>
                    <th className="p-1.5 border border-blue-900 text-left w-[115px] whitespace-nowrap px-1.5 truncate">ក្រុមផ្តល់ទិដ្ឋាការ</th>
                    {VISA_TYPES.map((vt) => (
                      <th key={vt.id} className="p-1.5 border border-blue-900 w-[62px] text-center font-times text-sm font-bold">
                        {vt.name}
                      </th>
                    ))}
                    <th className="p-1.5 border border-blue-900 bg-[#0B2545] font-bold w-[68px] text-center font-times text-sm">TOTAL</th>
                  </tr>

                  {/* 2. Top Red Header Summary Row (cEA) - Exactly aligned column by column */}
                  <tr className="bg-[#B91C1C] text-white font-bold text-xs">
                    <td colSpan={2} className="p-1.5 border border-red-800 text-left px-2">
                      <div className="flex items-center gap-1.5">
                        <span className="whitespace-nowrap font-bold text-xs">សរុបក្រុមផ្តល់ទិដ្ឋាការ</span>
                        <span className="px-1.5 py-0.5 bg-yellow-400 text-red-950 text-[10px] font-extrabold font-times rounded">
                          cEA
                        </span>
                      </div>
                    </td>
                    {VISA_TYPES.map((vt) => (
                      <td key={vt.id} className="p-1 border border-red-800 font-times font-bold text-center" style={{ fontFamily: "'Times New Roman', Times, serif" }}>
                        <span className="w-full max-w-[54px] mx-auto py-1 px-1 bg-black/25 rounded flex items-center justify-center text-xs font-times font-bold" style={{ fontFamily: "'Times New Roman', Times, serif" }}>
                          {topCeaTotals.totals[vt.id] || 0}
                        </span>
                      </td>
                    ))}
                    <td className="p-1 border border-red-800 font-times bg-[#991B1B] text-yellow-300 font-bold text-center" style={{ fontFamily: "'Times New Roman', Times, serif" }}>
                      <span className="w-full max-w-[60px] mx-auto py-1 px-1 bg-yellow-400 text-red-950 font-bold rounded flex items-center justify-center text-xs font-times" style={{ fontFamily: "'Times New Roman', Times, serif" }}>
                        {topCeaTotals.grandTotal}
                      </span>
                    </td>
                  </tr>
                </thead>
                <tbody>
                  {topCeaRows.map((row, idx) => (
                    <tr
                      key={row.teamName}
                      onClick={() => {
                        setSelectedTeam(row.teamName);
                        setSelectedOption('cEA');
                      }}
                      className={`cursor-pointer transition hover:bg-yellow-100 ${
                        selectedTeam === row.teamName && selectedOption === 'cEA'
                          ? 'bg-cyan-100 font-bold text-cyan-950'
                          : idx % 2 === 0
                          ? 'bg-white'
                          : 'bg-blue-50/20'
                      }`}
                    >
                      <td className="p-1.5 border border-gray-300 font-semibold font-times text-gray-700">{row.index}</td>
                      <td className="p-1.5 border border-gray-300 text-left font-medium whitespace-nowrap truncate px-2" title={row.teamName}>
                        {row.teamName}
                      </td>
                      {VISA_TYPES.map((vt) => {
                        const val = row.counts[vt.id] || 0;
                        return (
                          <td
                            key={vt.id}
                            className={`p-1.5 border border-gray-300 font-times text-xs text-center ${
                              val > 0 ? 'font-bold text-blue-900 bg-blue-50/50' : 'text-gray-400'
                            }`}
                          >
                            {val}
                          </td>
                        );
                      })}
                      <td className="p-1.5 border border-gray-300 font-bold font-times bg-blue-50/40 text-xs text-center">
                        {row.total}
                      </td>
                    </tr>
                  ))}
                  {topCeaRows.length === 0 && (
                    <tr>
                      <td colSpan={16} className="p-3 text-center text-xs text-gray-500 bg-white">
                        មិនទាន់មានក្រុមដែលបានជ្រើសរើស (ក្រដាសអនុម័ត) នៅក្នុង "ទិន្នន័យក្រុមផ្តល់ទិដ្ឋាការ" នៅឡើយទេ
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          )}

          {/* Main Teams Detailed Table (Fixed to Sticker) */}
          <div className="bg-white border-2 border-[#1E40AF] rounded overflow-x-auto">
            <table className="w-full text-xs text-center border-collapse table-fixed min-w-[1020px]">
              <colgroup>
                <col className="w-[36px] min-w-[36px]" />
                <col className="w-[115px] min-w-[115px]" />
                {VISA_TYPES.map((vt) => (
                  <col key={vt.id} className="w-[62px] min-w-[62px]" />
                ))}
                <col className="w-[68px] min-w-[68px]" />
              </colgroup>
              <thead>
                {/* 1. Date & Column Headers Row */}
                <tr className="bg-[#0369A1] text-white text-xs font-bold">
                  <th className="p-1.5 border border-sky-800 w-[36px]">ល.រ</th>
                  <th className="p-1.5 border border-sky-800 text-left w-[115px] whitespace-nowrap px-1.5 truncate">
                    <div className="flex items-center gap-1.5">
                      <span className="bg-yellow-300 text-blue-950 px-1.5 py-0.5 rounded font-bold font-times text-xs">
                        {selectedDate}
                      </span>
                    </div>
                  </th>
                  {VISA_TYPES.map((vt) => (
                    <th
                      key={vt.id}
                      onClick={() =>
                        setDrillDownVisa({
                          teamName: selectedTeam || 'ព្រំដែន អូរយ៉ាដាវ',
                          visaType: vt.id,
                        })
                      }
                      className="p-1.5 border border-sky-800 w-[62px] hover:bg-sky-700 cursor-pointer transition select-none text-center font-times text-sm font-bold"
                      title={`ចុចដើម្បីមើលរូបមន្ត និងប្រភពទិន្នន័យលម្អិត [${vt.name}]`}
                    >
                      {vt.name}
                    </th>
                  ))}
                  <th
                    onClick={() =>
                      setDrillDownVisa({
                        teamName: selectedTeam || 'ព្រំដែន អូរយ៉ាដាវ',
                        visaType: 'TOTAL',
                      })
                    }
                    className="p-1.5 border border-sky-800 bg-[#0B2545] hover:bg-blue-950 font-bold w-[68px] cursor-pointer transition select-none text-center font-times text-sm"
                    title="ចុចដើម្បីមើលរូបមន្ត និងប្រភពទិន្នន័យលម្អិតសរុប (TOTAL)"
                  >
                    TOTAL
                  </th>
                </tr>

                {/* 2. Active Selected Team Row (Blue Bar) - Displays Remaining Stock of Selected Team up to selectedDate */}
                <tr className="bg-[#0284C7] text-white font-bold text-xs">
                  <td colSpan={2} className="p-1.5 border border-sky-700 text-left px-2">
                    <div className="flex items-center gap-1.5">
                      <span className="text-xs tracking-wide font-bold">
                        {selectedTeam || 'ជ្រើសរើសក្រុម'}
                      </span>
                    </div>
                  </td>
                  {VISA_TYPES.map((vt) => {
                    const count = selectedTeam
                      ? selectedTeamRemainingStock.counts[vt.id] || 0
                      : formData[vt.id]?.quantity || 0;
                    return (
                      <td key={vt.id} className="p-1 border border-sky-700 font-times font-bold text-center" style={{ fontFamily: "'Times New Roman', Times, serif" }}>
                        <button
                          type="button"
                          onClick={() =>
                            setDrillDownVisa({
                              teamName: selectedTeam || 'ព្រំដែន អូរយ៉ាដាវ',
                              visaType: vt.id,
                            })
                          }
                          className="w-full max-w-[54px] mx-auto py-1 px-1 bg-black/25 hover:bg-yellow-400 hover:text-blue-950 active:scale-95 transition-all rounded font-times font-bold text-xs cursor-pointer shadow-xs leading-none flex items-center justify-center"
                          style={{ fontFamily: "'Times New Roman', Times, serif" }}
                          title={`ចុចដើម្បីមើលរូបមន្តបូកដក និងប្រភពទិន្នន័យ [${vt.name}] របស់ក្រុម ${selectedTeam || 'ព្រំដែន អូរយ៉ាដាវ'}`}
                        >
                          {count}
                        </button>
                      </td>
                    );
                  })}
                  <td className="p-1 border border-sky-700 font-times bg-[#0369A1] text-yellow-300 font-bold text-center" style={{ fontFamily: "'Times New Roman', Times, serif" }}>
                    <button
                      type="button"
                      onClick={() =>
                        setDrillDownVisa({
                          teamName: selectedTeam || 'ព្រំដែន អូរយ៉ាដាវ',
                          visaType: 'TOTAL',
                        })
                      }
                      className="w-full max-w-[60px] mx-auto py-1 px-1 bg-yellow-300 hover:bg-yellow-400 text-blue-950 font-bold rounded text-xs cursor-pointer shadow-xs active:scale-95 transition-all leading-none flex items-center justify-center font-times"
                      style={{ fontFamily: "'Times New Roman', Times, serif" }}
                      title={`ចុចដើម្បីមើលរូបមន្តបូកដក និងប្រភពទិន្នន័យសរុប (TOTAL) របស់ក្រុម ${selectedTeam || 'ព្រំដែន អូរយ៉ាដាវ'}`}
                    >
                      {selectedTeam
                        ? selectedTeamRemainingStock.total
                        : formTotalSheets}
                    </button>
                  </td>
                </tr>

                {/* 3. Overall Totals Row (Bright Red) */}
                <tr className="bg-[#DC2626] text-white font-bold text-xs">
                  <td className="p-1.5 border border-red-800 text-left px-2" colSpan={2}>
                    <div className="flex items-center gap-2">
                      <span className="font-bold text-xs">ប្រើប្រាស់ប្រចាំថ្ងៃ</span>
                    </div>
                  </td>
                  {VISA_TYPES.map((vt) => (
                    <td key={vt.id} className="p-1 border border-red-800 font-times font-bold text-center" style={{ fontFamily: "'Times New Roman', Times, serif" }}>
                      <span className="w-full max-w-[54px] mx-auto py-1 px-1 bg-black/25 rounded flex items-center justify-center text-xs font-times font-bold" style={{ fontFamily: "'Times New Roman', Times, serif" }}>
                        {stickerOverallTotals.totals[vt.id] || 0}
                      </span>
                    </td>
                  ))}
                  <td className="p-1 border border-red-800 font-times bg-[#991B1B] text-yellow-300 font-bold text-center" style={{ fontFamily: "'Times New Roman', Times, serif" }}>
                    <span className="w-full max-w-[60px] mx-auto py-1 px-1 bg-yellow-400 text-red-950 font-bold rounded flex items-center justify-center text-xs font-times" style={{ fontFamily: "'Times New Roman', Times, serif" }}>
                      {stickerOverallTotals.grandTotal}
                    </span>
                  </td>
                </tr>
              </thead>

              <tbody>
                {displayedStickerRows.length === 0 ? (
                  <tr>
                    <td
                      colSpan={VISA_TYPES.length + 3}
                      className="p-6 text-center text-gray-500 italic bg-white border border-gray-300"
                    >
                      ពុំទាន់មានក្រុមណាមួយបានប្រើប្រាស់ទិដ្ឋាការ (Sticker) ក្នុងថ្ងៃ {selectedDate} នៅឡើយទេ។ សូមជ្រើសរើសក្រុមក្នុង Form ខាងឆ្វេងដើម្បីបញ្ចូលទិន្នន័យ។
                    </td>
                  </tr>
                ) : (
                  displayedStickerRows.map((row, idx) => (
                    <tr
                      key={row.teamName}
                      onClick={() => {
                        setSelectedTeam(row.teamName);
                        setSelectedOption('Sticker');
                      }}
                      className={`cursor-pointer transition hover:bg-yellow-100 ${
                        selectedTeam === row.teamName && selectedOption === 'Sticker'
                          ? 'bg-blue-100 font-bold text-blue-900'
                          : idx % 2 === 0
                          ? 'bg-white'
                          : 'bg-blue-50/20'
                      }`}
                    >
                      <td className="p-1.5 border border-gray-300 text-gray-700 font-semibold font-times">{row.displayIndex}</td>
                      <td className="p-1.5 border border-gray-300 text-left font-medium text-slate-800 whitespace-nowrap truncate px-2" title={row.teamName}>
                        {row.teamName}
                      </td>
                      {VISA_TYPES.map((vt) => {
                        const count = row.counts[vt.id] || 0;
                        return (
                          <td
                            key={vt.id}
                            className={`p-1.5 border border-gray-300 font-times text-xs text-center ${
                              count > 0
                                ? 'font-bold text-blue-900 bg-blue-50'
                                : 'text-gray-400'
                            }`}
                          >
                            {count}
                          </td>
                        );
                      })}
                      <td
                        className={`p-1.5 border border-gray-300 font-times font-bold text-xs text-center ${
                          row.total > 0
                            ? 'bg-blue-200 text-blue-950'
                            : 'text-gray-400 bg-white'
                        }`}
                      >
                        {row.total}
                      </td>
                    </tr>
                  ))
                )}
              </tbody>
            </table>
          </div>
        </div>
      </div>

      {/* Drill-down Visa Detail & Formula Modal Dialog (Triggered by clicking on T, T1... or TOTAL) */}
      {drillDownVisa && drillDownDetails && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-xs flex items-center justify-center p-2 sm:p-4 z-50 animate-in fade-in duration-200">
          <div className="bg-white rounded-xl shadow-2xl border border-gray-300 w-full max-w-5xl max-h-[92vh] flex flex-col overflow-hidden">
            {/* Modal Header */}
            <div className="bg-[#1E3A8A] text-white px-4 py-3 flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <div className="p-2 bg-yellow-400 text-blue-950 rounded-lg shadow-sm">
                  <Calculator className="w-5 h-5" />
                </div>
                <div>
                  <h2 className="text-base sm:text-lg font-bold flex items-center gap-2">
                    <span>
                      រូបមន្តគណនា & ប្រភពទិន្នន័យលម្អិត [{drillDownDetails.visaType}]
                    </span>
                    <span className="text-xs font-extrabold bg-yellow-300 text-blue-950 px-2 py-0.5 rounded">
                      {drillDownDetails.teamName}
                    </span>
                  </h2>
                  <p className="text-xs text-blue-200">
                    កាលបរិច្ឆេទគិតត្រឹមថ្ងៃទី៖ <strong>{selectedDate}</strong>
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setDrillDownVisa(null)}
                className="text-white/80 hover:text-white p-1.5 rounded-lg hover:bg-white/10 transition cursor-pointer"
                title="បិទ (Close)"
              >
                <X className="w-6 h-6" />
              </button>
            </div>

            {/* Visa Quick Switcher Tabs */}
            <div className="bg-slate-100 border-b border-gray-300 px-3 py-2 flex items-center gap-1.5 overflow-x-auto">
              <span className="text-xs font-bold text-gray-600 whitespace-nowrap mr-1">
                ជ្រើសរើសទិដ្ឋាការ៖
              </span>
              <button
                type="button"
                onClick={() =>
                  setDrillDownVisa({
                    teamName: drillDownDetails.teamName,
                    visaType: 'TOTAL',
                  })
                }
                className={`px-2.5 py-1 rounded text-xs font-bold transition cursor-pointer ${
                  drillDownDetails.visaType === 'TOTAL'
                    ? 'bg-[#1E3A8A] text-white shadow-xs'
                    : 'bg-white text-gray-700 hover:bg-gray-200 border border-gray-300'
                }`}
              >
                សរុបទាំងអស់ (TOTAL)
              </button>
              {VISA_TYPES.map((vt) => (
                <button
                  key={vt.id}
                  type="button"
                  onClick={() =>
                    setDrillDownVisa({
                      teamName: drillDownDetails.teamName,
                      visaType: vt.id,
                    })
                  }
                  className={`px-2 py-1 rounded text-xs font-bold font-times transition cursor-pointer min-w-[32px] text-center ${
                    drillDownDetails.visaType === vt.id
                      ? 'bg-blue-600 text-white shadow-xs'
                      : 'bg-white text-gray-700 hover:bg-gray-200 border border-gray-300'
                  }`}
                >
                  {vt.name}
                </button>
              ))}
            </div>

            {/* Modal Body */}
            <div className="flex-1 overflow-y-auto p-4 space-y-4">
              {/* 1. Formula & Equation Visualizer */}
              <div className="bg-linear-to-r from-blue-50 to-indigo-50 border border-blue-200 rounded-xl p-4 shadow-xs">
                <div className="flex items-center gap-2 mb-3">
                  <span className="px-2 py-0.5 bg-blue-600 text-white text-[11px] font-bold rounded">
                    រូបមន្តគណនាស្តុក
                  </span>
                  <span className="text-xs text-blue-900 font-medium">
                    ស្តុកនៅសល់ចុងគ្រា = សន្និធិដើមគ្រា (Opening) + បើកផ្តល់ពីក២ (Received: ទទួលពីការិយាល័យ ឬផ្ទេរពីក្រុម) - ប្រើប្រាស់ (Used) - ផ្ទេរ/ខូច/បាត់ (Damaged/Transferred)
                  </span>
                </div>

                {/* Formula Visual Grid */}
                <div className="grid grid-cols-2 sm:grid-cols-5 gap-2 sm:gap-3 text-center">
                  {/* Opening */}
                  <div className="bg-white p-3 rounded-lg border border-blue-300 shadow-xs flex flex-col justify-between">
                    <span className="text-[11px] font-bold text-gray-600">
                      សន្និធិដើមគ្រា (Opening)
                    </span>
                    <div className="my-1">
                      <span className="text-xl sm:text-2xl font-black font-times text-blue-900">
                        {drillDownDetails.effectiveOpening}
                      </span>
                      <span className="text-[11px] text-gray-500 ml-1">សន្លឹក</span>
                    </div>
                    <span className="text-[10px] text-gray-500 font-times">
                      Baseline: {drillDownDetails.baselineSum}
                      {drillDownDetails.recordedOldSum > 0 &&
                        ` + ស្តុកចាស់: ${drillDownDetails.recordedOldSum}`}
                    </span>
                  </div>

                  {/* Issued */}
                  <div className="bg-white p-3 rounded-lg border border-emerald-300 shadow-xs flex flex-col justify-between">
                    <span className="text-[11px] font-bold text-emerald-800">
                      + បើកផ្តល់ពីក២ (Received)
                    </span>
                    <div className="my-1">
                      <span className="text-xl sm:text-2xl font-black font-times text-emerald-600">
                        +{drillDownDetails.totalIssued}
                      </span>
                      <span className="text-[11px] text-gray-500 ml-1">សន្លឹក</span>
                    </div>
                    <span className="text-[10px] text-emerald-700">
                      ទទួលពីការិយាល័យ ឬផ្ទេរពីក្រុម
                    </span>
                  </div>

                  {/* Used */}
                  <div className="bg-white p-3 rounded-lg border border-rose-300 shadow-xs flex flex-col justify-between">
                    <span className="text-[11px] font-bold text-rose-800">
                      - បានប្រើប្រាស់ (Used)
                    </span>
                    <div className="my-1">
                      <span className="text-xl sm:text-2xl font-black font-times text-rose-600">
                        -{drillDownDetails.totalUsed}
                      </span>
                      <span className="text-[11px] text-gray-500 ml-1">សន្លឹក</span>
                    </div>
                    <span className="text-[10px] text-rose-700">
                      បិទលើលិខិតឆ្លងដែន
                    </span>
                  </div>

                  {/* Damaged */}
                  <div className="bg-white p-3 rounded-lg border border-amber-300 shadow-xs flex flex-col justify-between">
                    <span className="text-[11px] font-bold text-amber-800">
                      - ខូច/បាត់/ត្រឡប់
                    </span>
                    <div className="my-1">
                      <span className="text-xl sm:text-2xl font-black font-times text-amber-600">
                        -{drillDownDetails.totalDamaged}
                      </span>
                      <span className="text-[11px] text-gray-500 ml-1">សន្លឹក</span>
                    </div>
                    <span className="text-[10px] text-amber-700">
                      ខូច បាត់ ឬប្រគល់ត្រឡប់
                    </span>
                  </div>

                  {/* Balance Result */}
                  <div className="col-span-2 sm:col-span-1 bg-yellow-300 p-3 rounded-lg border-2 border-yellow-500 shadow-sm flex flex-col justify-between">
                    <span className="text-[11px] font-black text-blue-950">
                      = ស្តុកនៅសល់ចុងគ្រា
                    </span>
                    <div className="my-1">
                      <span className="text-2xl sm:text-3xl font-black font-times text-blue-950">
                        {drillDownDetails.calculatedBalance}
                      </span>
                      <span className="text-xs text-blue-950 font-bold ml-1">សន្លឹក</span>
                    </div>
                    <span className="text-[10px] text-blue-900 font-bold">
                      តួលេខបង្ហាញក្នុងតារាង
                    </span>
                  </div>
                </div>
              </div>

              {/* 2. Diagnostic & Explanation Card (If Discrepancy / Duplicate Exists) */}
              {drillDownDetails.hasPotentialDuplicate && (
                <div className="bg-amber-50 border-l-4 border-amber-500 p-3.5 rounded-r-lg text-xs space-y-1.5 shadow-xs">
                  <div className="flex items-center gap-2 text-amber-900 font-bold">
                    <AlertTriangle className="w-4 h-4 text-amber-600 shrink-0" />
                    <span>
                      ការបកស្រាយអំពីមូលហេតុឡើងចំនួន {drillDownDetails.calculatedBalance} ជំនួសចំនួន {drillDownDetails.baselineSum}៖
                    </span>
                  </div>
                  <ul className="list-disc list-inside text-amber-950 space-y-1 ml-1 leading-relaxed">
                    <li>
                      <strong>តួលេខដើមគ្រាផ្លូវការ (Baseline 30-Nov-2018)៖</strong> មានចំនួន <strong className="font-times">{drillDownDetails.baselineSum}</strong> សន្លឹក។
                    </li>
                    <li>
                      <strong>ទិន្នន័យបញ្ចូលបន្ថែមក្នុងសៀវភៅស្តុក៖</strong> មានចំនួន <strong className="font-times">+{drillDownDetails.recordedOldSum}</strong> សន្លឹកទៀត (ជាហេតុធ្វើឱ្យបូកបញ្ចូលគ្នាឡើងដល់ <strong className="font-times">{drillDownDetails.calculatedBalance}</strong> សន្លឹក)។
                    </li>
                    <li>
                      <strong>ដំណោះស្រាយងាយស្រួល៖</strong> លោកអ្នកគ្រាន់តែចុច <strong>[ ✕ ដោះធីក ]</strong> ឬ <strong>[ 🗑️ លុប ]</strong> លើជួរទិន្នន័យស្ទួនខាងក្រោម តួលេខនឹងធ្លាក់មកនៅសល់ <strong className="font-times">{drillDownDetails.baselineSum}</strong> សន្លឹកត្រឹមត្រូវភ្លាមៗ!
                    </li>
                  </ul>
                </div>
              )}

              {/* 3. Detailed Transaction / Records Table */}
              <div className="bg-white border border-gray-300 rounded-xl overflow-hidden shadow-xs">
                <div className="bg-slate-100 px-4 py-2.5 border-b border-gray-300 flex items-center justify-between flex-wrap gap-2">
                  <div className="flex items-center gap-2">
                    <Hash className="w-4 h-4 text-blue-700" />
                    <span className="font-bold text-xs text-gray-800">
                      បញ្ជីប្រតិបត្តិការ & ប្រភពទិន្នន័យដែលបង្កើតជាផលបូក ({drillDownDetails.contributingRecords.length} ប្រតិបត្តិការ)
                    </span>
                  </div>
                  <div className="text-xs text-gray-500">
                    ចុចលើប្រអប់ធីកដើម្បី <strong className="text-blue-900">រាប់ / មិនរាប់</strong> ក្នុងរូបមន្ត
                  </div>
                </div>

                <div className="overflow-x-auto max-h-72">
                  {drillDownDetails.contributingRecords.length === 0 ? (
                    <div className="p-8 text-center text-gray-500">
                      <Info className="w-8 h-8 text-gray-400 mx-auto mb-2" />
                      <p className="font-semibold text-xs">
                        ពុំមានប្រតិបត្តិការកត់ត្រាផ្សេងទៀតឡើយ — ប្រើប្រាស់តែតួលេខដើមគ្រា Baseline ({drillDownDetails.baselineSum} សន្លឹក)
                      </p>
                    </div>
                  ) : (
                    <table className="w-full text-xs text-left border-collapse">
                      <thead>
                        <tr className="bg-slate-50 text-slate-700 font-bold border-b border-gray-200">
                          <th className="p-2 w-10 text-center">រាប់</th>
                          <th className="p-2 w-28">កាលបរិច្ឆេទ</th>
                          <th className="p-2 w-28">ប្រភេទ</th>
                          <th className="p-2">ប្រតិបត្តិការ</th>
                          <th className="p-2 w-32">លេខស៊េរី Start - End</th>
                          <th className="p-2 text-center w-20">បរិមាណ</th>
                          <th className="p-2 text-center w-20">ផលប៉ះពាល់</th>
                          <th className="p-2 text-center w-28">ស្ថានភាព</th>
                          <th className="p-2 text-center w-16">សកម្មភាព</th>
                        </tr>
                      </thead>
                      <tbody>
                        {drillDownDetails.contributingRecords.map((rec, idx) => {
                          const isEx = rec.isExcluded;
                          return (
                            <tr
                              key={rec.id}
                              className={`border-b border-gray-200 transition ${
                                isEx
                                  ? 'bg-gray-100 text-gray-400 opacity-60'
                                  : rec.category === 'opening'
                                  ? 'bg-amber-50/60 hover:bg-amber-100/60'
                                  : idx % 2 === 0
                                  ? 'bg-white hover:bg-blue-50/40'
                                  : 'bg-slate-50 hover:bg-blue-50/40'
                              }`}
                            >
                              {/* Checkbox */}
                              <td className="p-2 text-center">
                                <input
                                  type="checkbox"
                                  checked={!isEx}
                                  onChange={() => handleToggleIncludeItem(rec.id)}
                                  className="w-4 h-4 rounded text-blue-600 focus:ring-blue-500 cursor-pointer"
                                  title={isEx ? 'ចុចដើម្បីរាប់បញ្ចូល' : 'ចុចដើម្បីដោះធីកកុំឱ្យរាប់លើស'}
                                />
                              </td>

                              {/* Date */}
                              <td className="p-2 font-times font-medium text-gray-700">
                                {rec.date}
                              </td>

                              {/* Visa Type */}
                              <td className="p-2 font-bold text-blue-900 font-times">
                                {rec.visaType}
                              </td>

                              {/* Operation Title */}
                              <td className="p-2 font-medium text-gray-800">
                                <div>{rec.opTitle}</div>
                                {rec.remarks && (
                                  <div className="text-[10px] text-gray-500 font-times">{rec.remarks}</div>
                                )}
                              </td>

                              {/* Serial Range */}
                              <td className="p-2 font-times text-[11px] text-gray-600">
                                {rec.startSerial && rec.endSerial
                                  ? `${rec.startSerial} - ${rec.endSerial}`
                                  : '-'}
                              </td>

                              {/* Quantity */}
                              <td className="p-2 text-center font-times font-bold text-sm text-gray-900">
                                {rec.quantity}
                              </td>

                              {/* Impact (+ / -) */}
                              <td className="p-2 text-center">
                                <span
                                  className={`px-2 py-0.5 rounded font-times font-bold text-xs ${
                                    rec.impact === '+'
                                      ? 'bg-emerald-100 text-emerald-800'
                                      : 'bg-rose-100 text-rose-800'
                                  }`}
                                >
                                  {rec.impact}
                                  {rec.quantity}
                                </span>
                              </td>

                              {/* Status */}
                              <td className="p-2 text-center">
                                {isEx ? (
                                  <span className="inline-block px-2 py-0.5 bg-gray-200 text-gray-600 rounded-full font-bold text-[10px]">
                                    ✕ ដោះធីក (មិនរាប់)
                                  </span>
                                ) : (
                                  <span className="inline-block px-2 py-0.5 bg-emerald-100 text-emerald-800 rounded-full font-bold text-[10px]">
                                    ✓ កំពុងរាប់បញ្ចូល
                                  </span>
                                )}
                              </td>

                              {/* Action: Edit & Delete */}
                              <td className="p-2 text-center flex items-center justify-center gap-1.5">
                                <button
                                  type="button"
                                  onClick={() => {
                                    setDrillDownVisa(null);
                                    handleSelectRecordForEdit(rec);
                                  }}
                                  className="p-1 text-blue-600 hover:text-blue-800 hover:bg-blue-50 rounded transition cursor-pointer"
                                  title="កែប្រែទិន្នន័យនេះ"
                                >
                                  <Edit3 className="w-4 h-4" />
                                </button>
                                <button
                                  type="button"
                                  onClick={() => handleDeleteAuditRecord(rec.id, rec.source)}
                                  className="p-1 text-red-600 hover:text-red-800 hover:bg-red-50 rounded transition cursor-pointer"
                                  title="លុបទិន្នន័យនេះ"
                                >
                                  <Trash2 className="w-4 h-4" />
                                </button>
                              </td>
                            </tr>
                          );
                        })}
                      </tbody>
                    </table>
                  )}
                </div>
              </div>
            </div>

            {/* Modal Footer */}
            <div className="bg-slate-100 border-t border-gray-300 px-4 py-3 flex items-center justify-between">
              <div className="text-xs text-gray-600">
                លទ្ធផលគណនាចុងក្រោយ៖{' '}
                <strong className="text-blue-900 font-times text-sm">
                  {drillDownDetails.calculatedBalance} សន្លឹក
                </strong>
              </div>
              <button
                type="button"
                onClick={() => setDrillDownVisa(null)}
                className="px-4 py-1.5 bg-[#1E3A8A] hover:bg-blue-900 text-white rounded-lg font-bold text-xs shadow-sm cursor-pointer transition"
              >
                យល់ព្រម (Close)
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Checklist & Overcounted Audit Modal Dialog */}
      {isAuditChecklistOpen && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-xs flex items-center justify-center p-2 sm:p-4 z-50 animate-in fade-in duration-200">
          <div className="bg-white rounded-xl shadow-2xl border border-gray-300 w-full max-w-6xl max-h-[92vh] flex flex-col overflow-hidden">
            {/* Modal Header */}
            <div className="bg-[#1E3A8A] text-white px-4 py-3 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <div className="p-1.5 bg-yellow-400 text-blue-950 rounded-lg">
                  <CheckSquare className="w-5 h-5" />
                </div>
                <div>
                  <h2 className="text-base sm:text-lg font-bold flex items-center gap-2">
                    <span>ផ្ទៀងផ្ទាត់បញ្ជីទិន្នន័យ & ស្វែងរកទិន្នន័យរាប់លើស</span>
                    <span className="text-xs font-medium bg-blue-700 px-2 py-0.5 rounded-full text-blue-100">
                      Checklist Audit
                    </span>
                  </h2>
                  <p className="text-xs text-blue-200">
                    ពិនិត្យរាល់ទិន្នន័យដែលបានបញ្ចូល ដោះធីកដើម្បីកុំឱ្យរាប់លើស ឬលុបទិន្នន័យស្ទួន
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setIsAuditChecklistOpen(false)}
                className="text-white/80 hover:text-white p-1 rounded-lg hover:bg-white/10 transition cursor-pointer"
              >
                <X className="w-6 h-6" />
              </button>
            </div>

            {/* Filter & Search Bar */}
            <div className="bg-slate-50 border-b border-gray-200 p-3 flex flex-wrap items-center justify-between gap-2 text-xs">
              <div className="flex flex-wrap items-center gap-2 flex-1">
                {/* Search */}
                <div className="relative min-w-[180px] max-w-xs">
                  <Search className="w-3.5 h-3.5 absolute left-2.5 top-2.5 text-gray-400" />
                  <input
                    type="text"
                    placeholder="ស្វែងរកក្រុម, កាលបរិច្ឆេទ..."
                    value={auditSearch}
                    onChange={(e) => setAuditSearch(e.target.value)}
                    className="w-full pl-8 pr-3 py-1.5 bg-white border border-gray-300 rounded-md focus:border-blue-500 focus:outline-hidden"
                  />
                </div>

                {/* Date Filter */}
                <div className="flex items-center gap-1 bg-white px-2 py-1 border border-gray-300 rounded-md">
                  <Calendar className="w-3.5 h-3.5 text-gray-500" />
                  <select
                    value={auditDateFilter}
                    onChange={(e) => setAuditDateFilter(e.target.value as any)}
                    className="bg-transparent text-xs font-semibold text-gray-700 outline-hidden cursor-pointer"
                  >
                    <option value="current">ថ្ងៃនេះ ({selectedDate})</option>
                    <option value="all">គ្រប់កាលបរិច្ឆេទទាំងអស់</option>
                    <option value="custom">ជ្រើសរើសថ្ងៃផ្ទាល់ខ្លួន</option>
                  </select>
                  {auditDateFilter === 'custom' && (
                    <div className="ml-1 w-32 inline-block">
                      <CustomDatePicker
                        value={auditCustomDate}
                        onChange={(d) => setAuditCustomDate(d)}
                        className="py-0.5 text-xs"
                      />
                    </div>
                  )}
                </div>

                {/* Category Filter */}
                <div className="flex items-center gap-1 bg-white px-2 py-1 border border-gray-300 rounded-md">
                  <Layers className="w-3.5 h-3.5 text-gray-500" />
                  <select
                    value={auditCategoryFilter}
                    onChange={(e) => setAuditCategoryFilter(e.target.value as any)}
                    className="bg-transparent text-xs font-semibold text-gray-700 outline-hidden cursor-pointer"
                  >
                    <option value="all">គ្រប់ប្រភេទ (Sticker & cEA)</option>
                    <option value="Sticker">តែ Sticker</option>
                    <option value="cEA">តែ cEA</option>
                  </select>
                </div>

                {/* Show only duplicates */}
                <label className="flex items-center gap-1.5 bg-white px-2.5 py-1.5 border border-amber-300 rounded-md cursor-pointer hover:bg-amber-50">
                  <input
                    type="checkbox"
                    checked={auditShowOnlyOvercounted}
                    onChange={(e) => setAuditShowOnlyOvercounted(e.target.checked)}
                    className="rounded text-amber-600 focus:ring-amber-500"
                  />
                  <span className="font-bold text-amber-900 text-xs">⚠️ បង្ហាញតែទិន្នន័យស្ទួន/រាប់លើស</span>
                </label>
              </div>

              {/* Quick Batch Actions */}
              <div className="flex items-center gap-1.5">
                <button
                  type="button"
                  onClick={handleSelectAllInChecklist}
                  className="px-2.5 py-1 bg-blue-600 hover:bg-blue-700 text-white rounded text-xs font-semibold cursor-pointer shadow-xs"
                >
                  ✓ ធីកទាំងអស់
                </button>
                <button
                  type="button"
                  onClick={handleUncheckAllInChecklist}
                  className="px-2.5 py-1 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded text-xs font-semibold cursor-pointer"
                >
                  ✕ ដោះធីកទាំងអស់
                </button>
                <button
                  type="button"
                  onClick={handleAutoUncheckOvercounted}
                  className="px-2.5 py-1 bg-amber-600 hover:bg-amber-700 text-white rounded text-xs font-bold cursor-pointer shadow-xs flex items-center gap-1"
                >
                  <AlertTriangle className="w-3.5 h-3.5" />
                  <span>ដោះធីកទិន្នន័យស្ទួន</span>
                </button>
              </div>
            </div>

            {/* Checklist Table Content */}
            <div className="flex-1 overflow-y-auto p-3">
              {filteredAuditEntries.length === 0 ? (
                <div className="py-12 text-center text-gray-500 bg-gray-50 rounded-lg border border-dashed border-gray-300">
                  <CheckCircle2 className="w-10 h-10 text-gray-400 mx-auto mb-2" />
                  <p className="font-semibold text-sm">ពុំមានទិន្នន័យតាមលក្ខខណ្ឌចម្រាញ់នេះឡើយ</p>
                  <p className="text-xs text-gray-400 mt-1">សូមផ្លាស់ប្តូរលក្ខខណ្ឌកាលបរិច្ឆេទ ឬការស្វែងរក</p>
                </div>
              ) : (
                <table className="w-full text-xs text-left border-collapse">
                  <thead>
                    <tr className="bg-slate-100 text-slate-700 font-bold border-b border-slate-300">
                      <th className="p-2 w-10 text-center">រាប់</th>
                      <th className="p-2 w-28">កាលបរិច្ឆេទ</th>
                      <th className="p-2 w-20">ប្រភេទ</th>
                      <th className="p-2 w-32">ប្រភេទប្រតិបត្តិការ</th>
                      <th className="p-2">ឈ្មោះក្រុម</th>
                      <th className="p-2">បរិមាណតាមទិដ្ឋាការ</th>
                      <th className="p-2 text-center w-24">សរុបសន្លឹក</th>
                      <th className="p-2 text-center w-36">ស្ថានភាព</th>
                      <th className="p-2 text-center w-16">សកម្មភាព</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredAuditEntries.map((item, idx) => {
                      const isExcluded = excludedIds.includes(item.id);
                      const isOvercounted = item.isDuplicateOrOvercounted;

                      return (
                        <tr
                          key={item.id}
                          className={`border-b border-gray-200 transition ${
                            isExcluded
                              ? 'bg-gray-100 text-gray-400 opacity-60'
                              : isOvercounted
                              ? 'bg-amber-50/70 hover:bg-amber-100/70'
                              : idx % 2 === 0
                              ? 'bg-white hover:bg-blue-50/40'
                              : 'bg-slate-50 hover:bg-blue-50/40'
                          }`}
                        >
                          {/* Checkbox */}
                          <td className="p-2 text-center">
                            <input
                              type="checkbox"
                              checked={!isExcluded}
                              onChange={() => handleToggleIncludeItem(item.id)}
                              className="w-4 h-4 rounded text-blue-600 focus:ring-blue-500 cursor-pointer"
                              title={isExcluded ? 'ចុចដើម្បីរាប់បញ្ចូល' : 'ចុចដើម្បីដកចេញកុំឱ្យរាប់លើស'}
                            />
                          </td>

                          {/* Date */}
                          <td className="p-2 font-times font-medium">
                            <span
                              className={`px-1.5 py-0.5 rounded font-times ${
                                item.date === selectedDate
                                  ? 'bg-blue-100 text-blue-900 font-bold'
                                  : 'bg-gray-200 text-gray-700'
                              }`}
                            >
                              {item.date}
                            </span>
                          </td>

                          {/* Category */}
                          <td className="p-2">
                            <span
                              className={`px-2 py-0.5 rounded text-[11px] font-bold font-times ${
                                item.categoryType === 'Sticker'
                                  ? 'bg-blue-600 text-white'
                                  : 'bg-cyan-600 text-white'
                              }`}
                            >
                              {item.categoryType}
                            </span>
                          </td>

                          {/* Operation Title */}
                          <td className="p-2">
                            <span className="px-1.5 py-0.5 rounded text-[10px] font-semibold bg-slate-200 text-slate-800">
                              {item.operationTitle}
                            </span>
                          </td>

                          {/* Team Name */}
                          <td className="p-2 font-semibold text-slate-800">
                            {item.teamName}
                          </td>

                          {/* Visa breakdown */}
                          <td className="p-2">
                            <div className="flex flex-wrap gap-1">
                              {Object.entries(item.values || {}).map(([vt, val]) => {
                                const qty = Number((val as any)?.quantity) || 0;
                                if (qty <= 0) return null;
                                return (
                                  <span
                                    key={vt}
                                    className="px-1.5 py-0.5 bg-slate-200 text-slate-800 rounded font-times text-[10px] font-semibold"
                                  >
                                    {vt}: <strong className="text-blue-900 font-times">{qty}</strong>
                                  </span>
                                );
                              })}
                            </div>
                          </td>

                          {/* Total sheets */}
                          <td className="p-2 text-center font-times font-bold text-sm text-blue-950">
                            {item.totalSheets}
                          </td>

                          {/* Status */}
                          <td className="p-2 text-center">
                            {isExcluded ? (
                              <span className="inline-flex items-center gap-1 px-2 py-0.5 bg-gray-200 text-gray-700 rounded-full font-bold text-[10px]">
                                ✕ មិនរាប់បញ្ចូល
                              </span>
                            ) : isOvercounted ? (
                              <span className="inline-flex items-center gap-1 px-2 py-0.5 bg-amber-100 text-amber-900 border border-amber-300 rounded-full font-bold text-[10px]">
                                <AlertTriangle className="w-3 h-3 text-amber-600" />
                                ស្ទួន/រាប់លើស
                              </span>
                            ) : (
                              <span className="inline-flex items-center gap-1 px-2 py-0.5 bg-emerald-100 text-emerald-800 rounded-full font-bold text-[10px]">
                                ✓ រាប់បញ្ចូល
                              </span>
                            )}
                          </td>

                          {/* Action */}
                          <td className="p-2 text-center flex items-center justify-center gap-1.5">
                            <button
                              type="button"
                              onClick={() => handleSelectRecordForEdit(item)}
                              className="p-1 text-blue-600 hover:text-blue-800 hover:bg-blue-50 rounded transition cursor-pointer"
                              title="កែប្រែទិន្នន័យនេះ"
                            >
                              <Edit3 className="w-4 h-4" />
                            </button>
                            <button
                              type="button"
                              onClick={() => handleDeleteAuditRecord(item.id, item.source)}
                              className="p-1 text-red-600 hover:text-red-800 hover:bg-red-50 rounded transition cursor-pointer"
                              title="លុបទិន្នន័យនេះ"
                            >
                              <Trash2 className="w-4 h-4" />
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              )}
            </div>

            {/* Modal Footer Summary */}
            <div className="bg-slate-100 border-t border-gray-300 px-4 py-3 flex flex-wrap items-center justify-between gap-3 text-xs">
              <div className="flex items-center gap-4 flex-wrap">
                <span className="text-gray-600">
                  ទិន្នន័យសរុប៖ <strong className="text-gray-900 font-times">{filteredAuditEntries.length}</strong>
                </span>
                <span className="text-emerald-700">
                  បានរាប់បញ្ចូល៖{' '}
                  <strong className="font-times">
                    {filteredAuditEntries.filter((i) => !excludedIds.includes(i.id)).length}
                  </strong>
                </span>
                <span className="text-red-600">
                  បានដកចេញ (មិនរាប់)៖{' '}
                  <strong className="font-times">
                    {filteredAuditEntries.filter((i) => excludedIds.includes(i.id)).length}
                  </strong>
                </span>
              </div>

              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => setIsAuditChecklistOpen(false)}
                  className="px-4 py-1.5 bg-[#1E3A8A] hover:bg-blue-900 text-white rounded-lg font-bold text-xs shadow-sm cursor-pointer transition"
                >
                  យល់ព្រម & អនុវត្ត (Done)
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* In-App Delete Confirmation Modal (Safe from iframe restrictions) */}
      {deleteModalOpen && (
        <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl border border-red-200 max-w-md w-full overflow-hidden animate-in fade-in zoom-in-95 duration-150">
            <div className="bg-red-600 text-white px-4 py-3 flex items-center justify-between">
              <div className="flex items-center gap-2 font-bold text-sm">
                <Trash2 className="w-4 h-4" />
                <span>បញ្ជាក់ការលុបទិន្នន័យប្រតិបត្តិការ</span>
              </div>
              <button
                type="button"
                onClick={() => setDeleteModalOpen(false)}
                className="text-white/80 hover:text-white p-1 rounded-md transition cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
            <div className="p-4 space-y-3 text-sm text-gray-700">
              <p>
                តើអ្នកពិតជាចង់លុបទិន្នន័យប្រតិបត្តិការនេះចេញពីប្រព័ន្ធឬទេ?
              </p>
              <div className="bg-red-50 border border-red-200 rounded-lg p-3 space-y-1.5 text-xs">
                <div className="flex justify-between">
                  <span className="text-gray-500">ប្រភេទ៖</span>
                  <span className="font-bold text-blue-900">{selectedOption}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-gray-500">កាលបរិច្ឆេទ៖</span>
                  <span className="font-bold text-slate-800">{selectedDate}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-gray-500">ក្រុម៖</span>
                  <span className="font-bold text-red-700">{selectedTeam}</span>
                </div>
              </div>
              <p className="text-[11px] text-gray-500 italic">
                * ការលុបនេះនឹងដកទិន្នន័យទាំងចេញពីតារាងប្រតិបត្តិការប្រចាំថ្ងៃ និងទិន្នន័យសន្លឹកទិដ្ឋាការដែលពាក់ព័ន្ធ។
              </p>
            </div>
            <div className="bg-gray-50 px-4 py-3 border-t border-gray-200 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => setDeleteModalOpen(false)}
                className="px-3.5 py-1.5 bg-gray-200 hover:bg-gray-300 text-gray-800 rounded-lg text-xs font-semibold cursor-pointer transition"
              >
                បោះបង់ (Cancel)
              </button>
              <button
                type="button"
                onClick={handleConfirmDelete}
                className="px-4 py-1.5 bg-red-600 hover:bg-red-700 active:scale-95 text-white rounded-lg text-xs font-bold shadow-sm cursor-pointer transition flex items-center gap-1.5"
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>លុបចេញ (Confirm Delete)</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* In-App Update Confirmation Modal */}
      {updateModalOpen && (
        <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl border border-fuchsia-200 max-w-md w-full overflow-hidden animate-in fade-in zoom-in-95 duration-150">
            <div className="bg-[#A21CAF] text-white px-4 py-3 flex items-center justify-between">
              <div className="flex items-center gap-2 font-bold text-sm">
                <RefreshCw className="w-4 h-4" />
                <span>បញ្ជាក់ការកែសម្រួលទិន្នន័យ (Confirm Update)</span>
              </div>
              <button
                type="button"
                onClick={() => setUpdateModalOpen(false)}
                className="text-white/80 hover:text-white p-1 rounded-md transition cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
            <div className="p-4 space-y-3 text-sm text-gray-700">
              <p>
                តើអ្នកពិតជាចង់កែសម្រួលទិន្នន័យប្រតិបត្តិការប្រចាំថ្ងៃតាមក្រុមនេះឬទេ?
              </p>
              <div className="bg-fuchsia-50 border border-fuchsia-200 rounded-lg p-3 space-y-1.5 text-xs">
                <div className="flex justify-between">
                  <span className="text-gray-500">ប្រភេទ៖</span>
                  <span className="font-bold text-blue-900">{selectedOption}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-gray-500">កាលបរិច្ឆេទ៖</span>
                  <span className="font-bold text-slate-800">{selectedDate}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-gray-500">ក្រុម៖</span>
                  <span className="font-bold text-purple-800">{selectedTeam}</span>
                </div>
                <div className="flex justify-between border-t border-fuchsia-200 pt-1">
                  <span className="text-gray-600 font-medium">ចំនួនសរុបថ្មី៖</span>
                  <span className="font-extrabold text-fuchsia-700">{formTotalSheets} សន្លឹក</span>
                </div>
              </div>
              <p className="text-[11px] text-gray-500 italic">
                * ការកែសម្រួលនេះនឹងធ្វើបច្ចុប្បន្នភាពទិន្នន័យក្នុងប្រតិបត្តិការប្រចាំថ្ងៃ និងសម្អាតទម្រង់បញ្ចូលទិន្នន័យ។
              </p>
            </div>
            <div className="bg-gray-50 px-4 py-3 border-t border-gray-200 flex justify-end gap-2">
              <button
                type="button"
                onClick={() => setUpdateModalOpen(false)}
                className="px-3.5 py-1.5 bg-gray-200 hover:bg-gray-300 text-gray-800 rounded-lg text-xs font-semibold cursor-pointer transition"
              >
                បោះបង់ (Cancel)
              </button>
              <button
                type="button"
                onClick={handleConfirmUpdate}
                className="px-4 py-1.5 bg-[#C026D3] hover:bg-[#A21CAF] active:scale-95 text-white rounded-lg text-xs font-bold shadow-sm cursor-pointer transition flex items-center gap-1.5"
              >
                <CheckCircle2 className="w-3.5 h-3.5" />
                <span>កែសម្រួល (Confirm Update)</span>
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};

export default DailyTeamVisaOperations;
