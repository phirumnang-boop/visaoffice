import React, { useState, useMemo, useEffect } from 'react';
import { CategoriesState, StockRecord, UserRole, VisaRecord } from '../types';
import { apiService } from '../services/apiService';
import { Search, RotateCcw, Play, Copy, Check, FileText } from 'lucide-react';
import {
  VISA_TYPES,
  normalizeDateToISO,
  normalizeTeamName,
} from '../utils/teamNormalization';
import { formatKhmerDate } from '../utils/khmerCalendar';

export interface TeamCodeSearchProps {
  categories?: CategoriesState;
  stockRecords?: StockRecord[];
  visaRecords?: VisaRecord[];
  stickerActualStock?: StockRecord[];
  currentRole?: UserRole;
  userName?: string;
  onShowToast?: (msg: string, type: 'success' | 'error') => void;
}

interface FormRowEntry {
  id: string;
  quantity: string;
  startSerial: string;
  endSerial: string;
  oldCode: string;
}

interface DailyTeamRecord {
  id: string;
  categoryType: 'Sticker' | 'cEA';
  date: string;
  teamName: string;
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
}

export interface CodeSearchResultItem {
  id: string;
  visaType: string;
  operation: string; // 'ប្រចាំថ្ងៃ' | 'បើកផ្តល់' | 'ខូច/បាត់'
  khmerDateText: string;
  rawDate: string;
  teamName: string;
  startSerial: string;
  endSerial: string;
  oldCode?: string;
  quantity?: number;
}

/**
 * Fast parse helper for serial numbers without costly repeated regex
 */
function parseSerialNumeric(serialStr: string): { prefix: string; num: bigint | null } {
  if (!serialStr) return { prefix: '', num: null };
  const trimmed = serialStr.trim();
  const match = trimmed.match(/^(.*?)(\d+)(.*?)$/);
  if (!match) return { prefix: trimmed.toUpperCase(), num: null };
  try {
    return {
      prefix: match[1].toUpperCase(),
      num: BigInt(match[2]),
    };
  } catch {
    return { prefix: match[1].toUpperCase(), num: null };
  }
}

/**
 * Official Full Names Mapping for the 29 Visa Teams
 */
export const OFFICIAL_TEAM_FULL_NAMES: Record<string, string> = {
  'អាកាស តេជោ': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអាកាសយានដ្ឋានអន្តរជាតិតេជោ',
  'អាកាស សៀមរាប': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអាកាសយានដ្ឋានអន្តរជាតិសៀមរាបអង្គរ',
  'អាកាស ព្រះសីហនុ': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអាកាសយានដ្ឋានអន្តរជាតិក្រុងព្រះសីហនុ',
  'ព្រំដែន ប៉ោយប៉ែត': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិប៉ោយប៉ែត',
  'ព្រំដែន បាវិត': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិបាវិត',
  'ព្រំដែន ចាំយាម': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិចាំយាម',
  'ព្រំដែន ដូង': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិដូង',
  'ព្រំដែន អូរស្មាច់': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិអូរស្មាច់',
  'ព្រំដែន ព្រំ': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិព្រំ',
  'ព្រំដែន បន្ទាយចក្រី': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិបន្ទាយចក្រី',
  'ព្រំដែន ជាំ': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិជាំ',
  'ព្រំដែន ត្រពាំងផ្លុង': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិត្រពាំងផ្លុង',
  'ព្រំដែន ត្រពាំងស្រែ': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិត្រពាំងស្រែ',
  'ព្រំដែន ត្រពាំងក្រៀល': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិត្រពាំងក្រៀល',
  'ព្រំដែន ភ្នំដិន': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិភ្នំដិន',
  'ព្រំដែន កោះរកា': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិកោះរកា',
  'ព្រំដែន ព្រៃវល្លិ៍': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិព្រៃវល្លិ៍',
  'ព្រំដែន អូរយ៉ាដាវ': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិអូរយ៉ាដាវ',
  'ព្រំដែន ក្អមសំណ': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិក្អមសំណ',
  'ព្រំដែន ភ្នំដី': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិភ្នំដី',
  'កំពង់ផែ ឧកញ៉ាម៉ុង': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារកំពង់ផែអន្តរជាតិឧកញ៉ាម៉ុង',
  'កំពង់ផែ ស្ទឹងហាវ': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារកំពង់ផែអន្តរជាតិស្ទឹងហាវ',
  'កំពង់ផែ ព្រះសីហនុ': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារកំពង់ផែអន្តរជាតិក្រុងព្រះសីហនុ',
  'កំពង់ផែ ភ្នំពេញ': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារកំពង់ផែអន្តរជាតិភ្នំពេញ',
  'ព្រំដែន ព្រែកចាក': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិព្រែកចាក',
  'ព្រំដែន ម៉ឺនជ័យ': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិម៉ឺនជ័យ',
  'ព្រំដែន ស្ទឹងបត់': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិស្ទឹងបត់',
  'កំពង់ផែ កោះកុង': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារកំពង់ផែអន្តរជាតិកោះកុង',
  'កំពង់ផែ កំពត': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារកំពង់ផែអន្តរជាតិកំពត',
};

/**
 * Helper to format team name for the official report phrase
 */
function formatReportTeamName(teamName: string, categories?: CategoriesState): string {
  if (!teamName) return 'ក្រុមផ្តល់ទិដ្ឋាការ';
  const t = teamName.trim();
  const norm = normalizeTeamName(t);

  // 1. Check categories
  if (categories) {
    const robokList = categories.visaTeamsRobok || [];
    const fullList = categories.visaTeams || [];

    const rIdx = robokList.findIndex((item) => {
      if (!item?.name) return false;
      const rNorm = normalizeTeamName(item.name);
      return rNorm === norm || item.name.trim() === t || item.name.includes(t) || t.includes(item.name);
    });

    if (rIdx !== -1 && fullList[rIdx]?.name) {
      const candidateFull = fullList[rIdx].name.trim();
      if (candidateFull && !/^ក្រុមទី\s*\d+$/i.test(candidateFull)) {
        if (candidateFull.startsWith('ក្រុមផ្តល់ទិដ្ឋាការ')) return candidateFull;
        return `ក្រុមផ្តល់ទិដ្ឋាការប្រចាំ${candidateFull.replace(/^(ច្រកទ្វារ|ក្រុមផ្តល់ទិដ្ឋាការ|ប្រចាំ)+/, '').trim()}`;
      }
    }

    const directFull = fullList.find((item) => {
      if (!item?.name) return false;
      const fNorm = normalizeTeamName(item.name);
      return (fNorm === norm || item.name.includes(norm) || item.name.includes(t)) && !/^ក្រុមទី\s*\d+$/i.test(item.name.trim());
    });

    if (directFull?.name) {
      const dName = directFull.name.trim();
      if (dName.startsWith('ក្រុមផ្តល់ទិដ្ឋាការ')) return dName;
      return `ក្រុមផ្តល់ទិដ្ឋាការប្រចាំ${dName}`;
    }
  }

  // 2. Check predefined official 29 full titles dictionary
  if (OFFICIAL_TEAM_FULL_NAMES[norm]) {
    return OFFICIAL_TEAM_FULL_NAMES[norm];
  }
  if (OFFICIAL_TEAM_FULL_NAMES[t]) {
    return OFFICIAL_TEAM_FULL_NAMES[t];
  }

  // 3. Already has full official prefix
  if (t.startsWith('ក្រុមផ្តល់ទិដ្ឋាការ')) {
    return t;
  }

  // 4. Pattern formatting
  if (norm.startsWith('អាកាស') || t.startsWith('អាកាស')) {
    const airName = (norm.startsWith('អាកាស') ? norm : t).replace(/^អាកាស\s*/, '').trim();
    if (airName.includes('តេជោ')) return 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអាកាសយានដ្ឋានអន្តរជាតិតេជោ';
    if (airName.includes('សៀមរាប')) return 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអាកាសយានដ្ឋានអន្តរជាតិសៀមរាបអង្គរ';
    if (airName.includes('ព្រះសីហនុ')) return 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអាកាសយានដ្ឋានអន្តរជាតិក្រុងព្រះសីហនុ';
    return `ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអាកាសយានដ្ឋានអន្តរជាតិ ${airName}`;
  }

  if (norm.startsWith('កំពង់ផែ') || t.startsWith('កំពង់ផែ')) {
    const portName = (norm.startsWith('កំពង់ផែ') ? norm : t).replace(/^កំពង់ផែ\s*/, '').trim();
    return `ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារកំពង់ផែអន្តរជាតិ ${portName}`;
  }

  if (norm.startsWith('ព្រំដែន') || t.startsWith('ព្រំដែន')) {
    const borderName = (norm.startsWith('ព្រំដែន') ? norm : t).replace(/^ព្រំដែន\s*/, '').trim();
    return `ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិ ${borderName}`;
  }

  return `ក្រុមផ្តល់ទិដ្ឋាការ ${t}`;
}

/**
 * Checks if searchCode falls precisely between startSerial and endSerial
 */
function isCodeInRange(
  searchQuery: string,
  startSerial: string,
  endSerial: string,
  oldCode?: string
): boolean {
  if (!searchQuery) return false;
  const qClean = searchQuery.trim().toUpperCase();

  // 1. Direct exact match on start, end, or old code
  if (startSerial && startSerial.toUpperCase() === qClean) return true;
  if (endSerial && endSerial.toUpperCase() === qClean) return true;
  if (oldCode && oldCode.toUpperCase() === qClean) return true;

  // 2. Numerical Range match: start <= query <= end
  if (startSerial && endSerial) {
    const startParsed = parseSerialNumeric(startSerial);
    const endParsed = parseSerialNumeric(endSerial);
    const queryParsed = parseSerialNumeric(qClean);

    if (startParsed.num !== null && endParsed.num !== null && queryParsed.num !== null) {
      const prefixMatches = !queryParsed.prefix || queryParsed.prefix === startParsed.prefix;
      if (prefixMatches) {
        const minVal = startParsed.num < endParsed.num ? startParsed.num : endParsed.num;
        const maxVal = startParsed.num > endParsed.num ? startParsed.num : endParsed.num;

        if (queryParsed.num >= minVal && queryParsed.num <= maxVal) {
          return true;
        }
      }
    }
  }

  // 3. Fallback contains check if non-pure numeric
  if (
    (startSerial && startSerial.toUpperCase().includes(qClean)) ||
    (endSerial && endSerial.toUpperCase().includes(qClean))
  ) {
    return true;
  }

  return false;
}

export const TeamCodeSearch: React.FC<TeamCodeSearchProps> = ({
  categories,
  stockRecords = [],
  visaRecords = [],
  stickerActualStock = [],
  onShowToast,
}) => {
  // Input Criteria
  const [selectedVisaType, setSelectedVisaType] = useState<string>('T');
  const [inputSerialCode, setInputSerialCode] = useState<string>('');
  const [selectedOperation, setSelectedOperation] = useState<string>('ប្រចាំថ្ងៃ');

  // Search execution state (only displays results after user clicks Show)
  const [hasSearched, setHasSearched] = useState<boolean>(false);
  const [activeSearchedCode, setActiveSearchedCode] = useState<string>('');
  const [activeVisaType, setActiveVisaType] = useState<string>('T');
  const [activeOperation, setActiveOperation] = useState<string>('ប្រចាំថ្ងៃ');

  // Copy text status state
  const [copiedIndex, setCopiedIndex] = useState<number | null>(null);
  const [copiedAll, setCopiedAll] = useState<boolean>(false);

  // Daily team records state
  const [dailyTeamRecords, setDailyTeamRecords] = useState<DailyTeamRecord[]>(() => {
    try {
      const saved = localStorage.getItem('app_daily_team_operations_v5');
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed)) return parsed;
      }
    } catch {}
    return [];
  });

  useEffect(() => {
    let isMounted = true;
    const fetchCloudDaily = async () => {
      try {
        const cloudDaily = await apiService.getDailyTeamOperations();
        if (isMounted && Array.isArray(cloudDaily) && cloudDaily.length > 0) {
          setDailyTeamRecords(cloudDaily);
        }
      } catch (e) {
        // Silent fallback
      }
    };
    fetchCloudDaily();
    return () => {
      isMounted = false;
    };
  }, []);

  // Handle Search click
  const handleShowClick = (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    const code = inputSerialCode.trim();
    if (!code) {
      if (onShowToast) onShowToast('សូមវាយបញ្ចូលលេខកូដស៊េរីដែលត្រូវស្វែងរក', 'error');
      return;
    }

    setActiveSearchedCode(code);
    setActiveVisaType(selectedVisaType);
    setActiveOperation(selectedOperation);
    setHasSearched(true);
    setCopiedAll(false);
    setCopiedIndex(null);
  };

  // Reset form
  const handleReset = () => {
    setInputSerialCode('');
    setSelectedVisaType('T');
    setSelectedOperation('ប្រចាំថ្ងៃ');
    setHasSearched(false);
    setActiveSearchedCode('');
    setCopiedAll(false);
    setCopiedIndex(null);
  };

  // Ultra-fast search performed ONLY when user clicks Show, formatting dates ONLY for matched rows
  const displayedRows = useMemo<CodeSearchResultItem[]>(() => {
    if (!hasSearched || !activeSearchedCode) return [];

    const qClean = activeSearchedCode.trim();
    const results: CodeSearchResultItem[] = [];

    // 1. Search in Daily Team Records
    for (let i = 0; i < dailyTeamRecords.length; i++) {
      const dtr = dailyTeamRecords[i];
      if (!dtr.values) continue;

      const teamName = normalizeTeamName(dtr.teamName || '');
      const rawDate = normalizeDateToISO(dtr.date || '');

      for (const [vt, valObj] of Object.entries(dtr.values)) {
        if (!valObj) continue;
        const vTypeUpper = vt.toUpperCase();

        if (activeVisaType !== 'ALL' && vTypeUpper !== activeVisaType) continue;
        if (activeOperation !== 'ALL' && activeOperation !== 'ប្រចាំថ្ងៃ') continue;

        const val = valObj as any;
        const entries = val.entries;

        if (Array.isArray(entries) && entries.length > 0) {
          for (let sIdx = 0; sIdx < entries.length; sIdx++) {
            const subE = entries[sIdx];
            const start = subE.startSerial ? String(subE.startSerial).trim() : '';
            const end = subE.endSerial ? String(subE.endSerial).trim() : '';
            const old = subE.oldCode ? String(subE.oldCode).trim() : '';
            const qty = subE.quantity ? parseInt(String(subE.quantity), 10) || 0 : 0;

            if (isCodeInRange(qClean, start, end, old)) {
              const khmerFormatted = rawDate ? formatKhmerDate(rawDate) : '';
              results.push({
                id: `daily-${dtr.id}-${vt}-${sIdx}`,
                visaType: vTypeUpper,
                operation: 'ប្រចាំថ្ងៃ',
                khmerDateText: khmerFormatted ? `ប្រើប្រាស់${khmerFormatted}` : (dtr.date || ''),
                rawDate,
                teamName,
                startSerial: start,
                endSerial: end,
                oldCode: old,
                quantity: qty > 0 ? qty : 1,
              });
            }
          }
        } else {
          const start = val.startSerial ? String(val.startSerial).trim() : '';
          const end = val.endSerial ? String(val.endSerial).trim() : '';
          const old = val.oldCode ? String(val.oldCode).trim() : '';
          const qty = val.quantity ? parseInt(String(val.quantity), 10) || 0 : 0;

          if (isCodeInRange(qClean, start, end, old)) {
            const khmerFormatted = rawDate ? formatKhmerDate(rawDate) : '';
            results.push({
              id: `daily-${dtr.id}-${vt}-0`,
              visaType: vTypeUpper,
              operation: 'ប្រចាំថ្ងៃ',
              khmerDateText: khmerFormatted ? `ប្រើប្រាស់${khmerFormatted}` : (dtr.date || ''),
              rawDate,
              teamName,
              startSerial: start,
              endSerial: end,
              oldCode: old,
              quantity: qty > 0 ? qty : 1,
            });
          }
        }
      }
    }

    // 2. Search in Stock Records (Ledger)
    for (let i = 0; i < stockRecords.length; i++) {
      const stk = stockRecords[i];
      if (
        !stk.visaTeamRobokName ||
        stk.visaTeamRobokName.includes('ការិយាល័យ') ||
        stk.visaTeamRobokName.includes('ឃ្លាំង') ||
        stk.operationType === 'receive' ||
        stk.operationType === 'import'
      ) {
        continue;
      }

      const vt = (stk.visaType || 'T').toUpperCase();
      if (activeVisaType !== 'ALL' && vt !== activeVisaType) continue;

      let op = 'ប្រចាំថ្ងៃ';
      if (stk.operationType === 'issueTeam') op = 'បើកផ្តល់';
      else if (stk.operationType === 'damaged' || stk.operationType === 'damagedTeam') op = 'ខូច/បាត់';
      else if (stk.operationType === 'returnTeam') op = 'ប្រគល់ត្រឡប់';

      if (activeOperation !== 'ALL' && op !== activeOperation) continue;

      const start = stk.startSerial ? String(stk.startSerial).trim() : '';
      const end = stk.endSerial ? String(stk.endSerial).trim() : '';

      if (isCodeInRange(qClean, start, end)) {
        const teamName = normalizeTeamName(stk.visaTeamRobokName || '');
        const rawDate = normalizeDateToISO(stk.date || '');
        const khmerFormatted = rawDate ? formatKhmerDate(rawDate) : '';

        results.push({
          id: `stock-${stk.id}`,
          visaType: vt,
          operation: op,
          khmerDateText: khmerFormatted ? `ប្រើប្រាស់${khmerFormatted}` : (stk.date || ''),
          rawDate,
          teamName,
          startSerial: start,
          endSerial: end,
          quantity: stk.totalSheets || 0,
        });
      }
    }

    // 3. Search in Actual Stock Records
    for (let i = 0; i < stickerActualStock.length; i++) {
      const stk = stickerActualStock[i];
      if (
        !stk.visaTeamRobokName ||
        stk.visaTeamRobokName.includes('ការិយាល័យ') ||
        stk.visaTeamRobokName.includes('ឃ្លាំង')
      ) {
        continue;
      }

      const vt = (stk.visaType || 'T').toUpperCase();
      if (activeVisaType !== 'ALL' && vt !== activeVisaType) continue;
      if (activeOperation !== 'ALL' && activeOperation !== 'បើកផ្តល់') continue;

      const start = stk.startSerial ? String(stk.startSerial).trim() : '';
      const end = stk.endSerial ? String(stk.endSerial).trim() : '';

      if (isCodeInRange(qClean, start, end)) {
        const teamName = normalizeTeamName(stk.visaTeamRobokName || '');
        const rawDate = normalizeDateToISO(stk.date || '');
        const khmerFormatted = rawDate ? formatKhmerDate(rawDate) : '';

        results.push({
          id: `actual-${stk.id}`,
          visaType: vt,
          operation: 'បើកផ្តល់',
          khmerDateText: khmerFormatted ? `ប្រើប្រាស់${khmerFormatted}` : (stk.date || ''),
          rawDate,
          teamName,
          startSerial: start,
          endSerial: end,
          quantity: stk.totalSheets || 0,
        });
      }
    }

    // 4. Search in Individual Visa Records
    for (let i = 0; i < visaRecords.length; i++) {
      const visa = visaRecords[i];
      const docNum = visa.documentNumber ? String(visa.documentNumber).trim() : '';
      if (!docNum) continue;

      const vt = (visa.newVisaType || visa.oldVisaType || 'T').toUpperCase();
      if (activeVisaType !== 'ALL' && vt !== activeVisaType) continue;
      if (activeOperation !== 'ALL' && activeOperation !== 'ប្រចាំថ្ងៃ') continue;

      if (isCodeInRange(qClean, docNum, docNum)) {
        const teamObj = categories?.visaTeams?.find((t) => t.id === visa.visaTeamId);
        const teamName = normalizeTeamName(teamObj?.name || visa.visaTeamId || '');
        const rawDate = normalizeDateToISO(visa.issueDate || visa.applicationDate || '');
        const khmerFormatted = rawDate ? formatKhmerDate(rawDate) : '';

        results.push({
          id: `visa-${visa.id}`,
          visaType: vt,
          operation: 'ប្រចាំថ្ងៃ',
          khmerDateText: khmerFormatted ? `ប្រើប្រាស់${khmerFormatted}` : (visa.issueDate || ''),
          rawDate,
          teamName: teamName || 'ក្រុមផ្តល់ទិដ្ឋាការ',
          startSerial: docNum,
          endSerial: docNum,
          quantity: 1,
        });
      }
    }

    // Remove exact duplicates
    const uniqueResults = [];
    const seen = new Set();
    for (const item of results) {
      const key = `${item.visaType}-${item.operation}-${item.startSerial}-${item.endSerial}-${item.teamName}-${item.rawDate}`;
      if (!seen.has(key)) {
        seen.add(key);
        uniqueResults.push(item);
      }
    }
    return uniqueResults;
  }, [
    dailyTeamRecords,
    stockRecords,
    stickerActualStock,
    visaRecords,
    categories?.visaTeams,
    hasSearched,
    activeSearchedCode,
    activeVisaType,
    activeOperation,
  ]);

  /**
   * Generates the requested official report text for a given result row
   */
  const generateRowReportText = (row: CodeSearchResultItem): string => {
    const vType = activeVisaType !== 'ALL' ? activeVisaType : row.visaType;
    const formattedTeam = formatReportTeamName(row.teamName, categories);

    let datePhrase = row.khmerDateText;
    if (datePhrase.startsWith('ប្រើប្រាស់')) {
      datePhrase = `បាន${datePhrase}`;
    } else if (!datePhrase.startsWith('បាន')) {
      datePhrase = `បានប្រើប្រាស់${datePhrase}`;
    }

    return `សូមគោរពរាយការណ៍ជូនលោកវរសេនីយ៍ទោ នាយការិយាល័យទិដ្ឋាការចូល លេខទិដ្ឋាការប្រភេទ ${vType} ${activeSearchedCode} ${datePhrase} ដោយ${formattedTeam}។\nដោយសេចក្តីគោរព!`;
  };

  // Consolidated report text for the first/primary result
  const primaryReportText = useMemo(() => {
    if (displayedRows.length === 0) return '';
    return generateRowReportText(displayedRows[0]);
  }, [displayedRows, activeVisaType, activeSearchedCode, categories]);

  // Copy handler
  const handleCopyReport = async (text: string, index?: number) => {
    if (!text) return;
    try {
      if (navigator.clipboard && navigator.clipboard.writeText) {
        await navigator.clipboard.writeText(text);
      } else {
        const textArea = document.createElement('textarea');
        textArea.value = text;
        document.body.appendChild(textArea);
        textArea.select();
        document.execCommand('copy');
        document.body.removeChild(textArea);
      }

      if (index !== undefined) {
        setCopiedIndex(index);
        setTimeout(() => setCopiedIndex(null), 2500);
      } else {
        setCopiedAll(true);
        setTimeout(() => setCopiedAll(false), 2500);
      }

      if (onShowToast) onShowToast('បានចម្លងខ្លឹមសាររាយការណ៍ជោគជ័យ', 'success');
    } catch (err) {
      if (onShowToast) onShowToast('មិនអាចចម្លងបានទេ សូមសាកល្បងម្តងទៀត', 'error');
    }
  };

  return (
    <div className="w-full bg-white p-4 sm:p-6 rounded-xl shadow-sm border border-slate-200 font-['Khmer_OS_Siemreap','Siemreap',sans-serif]">
      {/* 1. Filter Input Form */}
      <form onSubmit={handleShowClick} className="flex flex-wrap items-end gap-3 mb-5 pb-5 border-b border-slate-200">
        {/* Visa Type Selection */}
        <div className="w-32">
          <label className="block text-xs font-bold text-slate-700 mb-1.5">
            ប្រភេទទិដ្ឋាការ ៖
          </label>
          <select
            value={selectedVisaType}
            onChange={(e) => setSelectedVisaType(e.target.value)}
            className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-bold text-slate-800 focus:border-[#0060B6] focus:bg-white outline-none cursor-pointer"
          >
            <option value="ALL">-- ទាំងអស់ --</option>
            {VISA_TYPES.map((vt) => (
              <option key={vt} value={vt}>
                {vt}
              </option>
            ))}
          </select>
        </div>

        {/* Serial Code Input */}
        <div className="flex-1 min-w-[220px]">
          <label className="block text-xs font-bold text-slate-700 mb-1.5">
            វាយបញ្ចូលលេខកូដ (Serial Code) ៖
          </label>
          <div className="relative">
            <input
              type="text"
              value={inputSerialCode}
              onChange={(e) => setInputSerialCode(e.target.value)}
              className="w-full pl-9 pr-3 py-2 bg-slate-50 border border-slate-300 focus:border-[#0060B6] focus:bg-white rounded-lg text-sm font-bold text-slate-900 font-mono outline-none transition shadow-inner"
            />
            <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
          </div>
        </div>

        {/* Operation Selection */}
        <div className="w-40">
          <label className="block text-xs font-bold text-slate-700 mb-1.5">
            ប្រតិបត្តិការ ៖
          </label>
          <select
            value={selectedOperation}
            onChange={(e) => setSelectedOperation(e.target.value)}
            className="w-full px-3 py-2 bg-slate-50 border border-slate-300 rounded-lg text-sm font-semibold text-slate-800 focus:border-[#0060B6] focus:bg-white outline-none cursor-pointer"
          >
            <option value="ប្រចាំថ្ងៃ">ប្រចាំថ្ងៃ</option>
            <option value="បើកផ្តល់">បើកផ្តល់</option>
            <option value="ខូច/បាត់">ខូច/បាត់</option>
            <option value="ALL">-- ទាំងអស់ --</option>
          </select>
        </div>

        {/* Show Button */}
        <button
          type="submit"
          className="inline-flex items-center justify-center gap-1.5 px-5 py-2 bg-[#0060B6] hover:bg-[#004f98] text-white text-sm font-bold rounded-lg transition shadow-md cursor-pointer active:scale-98"
        >
          <Play className="w-4 h-4 fill-white" />
          <span>Show (បង្ហាញ)</span>
        </button>

        {/* Reset Button */}
        <button
          type="button"
          onClick={handleReset}
          className="p-2 text-slate-500 hover:text-slate-700 bg-slate-100 hover:bg-slate-200 rounded-lg transition cursor-pointer"
          title="កំណត់ឡើងវិញ"
        >
          <RotateCcw className="w-4 h-4" />
        </button>
      </form>

      {/* 2. Automated Official Report Box with One-Click Copy */}
      {hasSearched && displayedRows.length > 0 && primaryReportText && (
        <div className="mb-5 p-4 bg-blue-50/80 border-2 border-blue-200 rounded-xl shadow-xs animate-in fade-in duration-200">
          <div className="flex flex-wrap items-center justify-between gap-3 mb-2.5 pb-2 border-b border-blue-200/70">
            <div className="flex items-center gap-2 text-blue-900 font-bold text-sm">
              <FileText className="w-4 h-4 text-[#0060B6]" />
              <span>ខ្លឹមសាររាយការណ៍ផ្លូវការ (Official Report Message)</span>
            </div>

            {/* Quick Copy Button */}
            <button
              type="button"
              onClick={() => handleCopyReport(primaryReportText)}
              className={`inline-flex items-center gap-2 px-3.5 py-1.5 rounded-lg text-xs font-bold transition shadow-xs cursor-pointer ${
                copiedAll
                  ? 'bg-emerald-600 text-white shadow-emerald-200'
                  : 'bg-[#0060B6] hover:bg-[#004f98] text-white shadow-blue-200 active:scale-95'
              }`}
            >
              {copiedAll ? (
                <>
                  <Check className="w-3.5 h-3.5" />
                  <span>បានចម្លងជោគជ័យ!</span>
                </>
              ) : (
                <>
                  <Copy className="w-3.5 h-3.5" />
                  <span>ទាញ/ចម្លងខ្លឹមសាររាយការណ៍</span>
                </>
              )}
            </button>
          </div>

          <div className="relative bg-white p-3.5 rounded-lg border border-blue-200 shadow-inner">
            <p className="text-slate-800 text-[13.5px] leading-relaxed whitespace-pre-line font-medium select-all">
              {primaryReportText}
            </p>
          </div>
        </div>
      )}

      {/* 3. Results Table with Exact Blue Header */}
      <div className="overflow-x-auto border border-[#0060B6] rounded-sm">
        <table className="w-full text-center border-collapse text-[13.5px]">
          <thead>
            <tr className="bg-[#0060B6] text-white font-bold h-11 border-b border-[#0060B6]">
              {/* Column 1: Visa Type (e.g. T ▾) */}
              <th className="py-2 px-3 border-r border-white/30 text-center w-[150px] min-w-[130px]">
                <div className="inline-flex items-center justify-center gap-1.5 font-bold text-sm">
                  <span>{activeVisaType !== 'ALL' ? activeVisaType : (selectedVisaType !== 'ALL' ? selectedVisaType : 'T')}</span>
                  <div className="w-4 h-4 bg-white/20 rounded flex items-center justify-center border border-white/40">
                    <span className="text-[9px] leading-none text-white font-mono">▼</span>
                  </div>
                </div>
              </th>

              {/* Column 2: ប្រតិបត្តិការ */}
              <th className="py-2 px-3 border-r border-white/30 text-center font-bold min-w-[120px]">
                ប្រតិបត្តិការ
              </th>

              {/* Column 3: ថ្ងៃខែឆ្នាំ */}
              <th className="py-2 px-3 border-r border-white/30 text-center font-bold min-w-[240px]">
                ថ្ងៃខែឆ្នាំ
              </th>

              {/* Column 4: ក្រុមផ្តល់ទិដ្ឋាការ */}
              <th className="py-2 px-3 border-r border-white/30 text-center font-bold min-w-[180px]">
                ក្រុមផ្តល់ទិដ្ឋាការ
              </th>

              {/* Column 5: ចាប់ពីលេខ */}
              <th className="py-2 px-3 border-r border-white/30 text-center font-bold min-w-[130px]">
                ចាប់ពីលេខ
              </th>

              {/* Column 6: ដល់លេខ */}
              <th className="py-2 px-3 border-r border-white/30 text-center font-bold min-w-[130px]">
                ដល់លេខ
              </th>

              {/* Column 7: សកម្មភាពចម្លង */}
              <th className="py-2 px-3 text-center font-bold w-[110px]">
                ចម្លងខ្លឹមសារ
              </th>
            </tr>
          </thead>
          <tbody className="divide-y divide-slate-200 text-slate-900 bg-white">
            {!hasSearched ? (
              <tr>
                <td colSpan={7} className="py-12 text-center text-slate-400 font-medium">
                  <div className="flex flex-col items-center justify-center gap-1.5">
                    <Search className="w-6 h-6 text-slate-300" />
                    <span>សូមជ្រើសរើសប្រភេទ វាយបញ្ចូលលេខកូដ និងជ្រើសរើសប្រតិបត្តិការ រួចចុចប៊ូតុង <strong>Show (បង្ហាញ)</strong></span>
                  </div>
                </td>
              </tr>
            ) : displayedRows.length === 0 ? (
              <tr>
                <td colSpan={7} className="py-10 text-center text-rose-500 font-medium">
                  មិនមានទិន្នន័យរកឃើញសម្រាប់លេខកូដ «{activeSearchedCode}» ក្នុងប្រភេទ «{activeVisaType}» និងប្រតិបត្តិការ «{activeOperation}» ទេ
                </td>
              </tr>
            ) : (
              displayedRows.map((row, idx) => {
                const rowReport = generateRowReportText(row);
                const isRowCopied = copiedIndex === idx;

                return (
                  <tr
                    key={row.id || idx}
                    className="h-10 hover:bg-blue-50/40 transition border-b border-slate-200"
                  >
                    {/* Column 1: Input Serial Number */}
                    <td className="py-2 px-3 border-r border-slate-200 font-mono text-[13.5px] font-bold text-slate-900">
                      {activeSearchedCode}
                    </td>

                    {/* Column 2: ប្រតិបត្តិការ */}
                    <td className="py-2 px-3 border-r border-slate-200 text-slate-800">
                      {row.operation}
                    </td>

                    {/* Column 3: ថ្ងៃខែឆ្នាំ */}
                    <td className="py-2 px-3 border-r border-slate-200 text-slate-800">
                      {row.khmerDateText}
                    </td>

                    {/* Column 4: ក្រុមផ្តល់ទិដ្ឋាការ */}
                    <td className="py-2 px-3 border-r border-slate-200 text-slate-900 font-semibold">
                      {row.teamName}
                    </td>

                    {/* Column 5: ចាប់ពីលេខ */}
                    <td className="py-2 px-3 border-r border-slate-200 font-mono text-[13px] text-slate-800">
                      {row.startSerial || '-'}
                    </td>

                    {/* Column 6: ដល់លេខ */}
                    <td className="py-2 px-3 border-r border-slate-200 font-mono text-[13px] text-slate-800">
                      {row.endSerial || '-'}
                    </td>

                    {/* Column 7: Action Copy */}
                    <td className="py-2 px-2 text-center">
                      <button
                        type="button"
                        onClick={() => handleCopyReport(rowReport, idx)}
                        className={`inline-flex items-center justify-center gap-1 px-2.5 py-1 rounded text-xs font-semibold transition cursor-pointer ${
                          isRowCopied
                            ? 'bg-emerald-100 text-emerald-800 font-bold'
                            : 'bg-blue-50 hover:bg-blue-100 text-[#0060B6]'
                        }`}
                        title="ចម្លងខ្លឹមសាររាយការណ៍នៃជួរនេះ"
                      >
                        {isRowCopied ? (
                          <>
                            <Check className="w-3 h-3 text-emerald-600" />
                            <span>បានចម្លង</span>
                          </>
                        ) : (
                          <>
                            <Copy className="w-3 h-3" />
                            <span>ចម្លង</span>
                          </>
                        )}
                      </button>
                    </td>
                  </tr>
                );
              })
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
};
