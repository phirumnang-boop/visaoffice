import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { StockRecord, CategoriesState, Officer, UserRole } from '../types';
import { printA4Document } from '../utils/printHelper';
import { CustomDatePicker } from './CustomDatePicker';
import {
  Package,
  Layers,
  Search,
  Filter,
  Printer,
  FileSpreadsheet,
  Calendar,
  AlertTriangle,
  CheckCircle2,
  BookOpen,
  ArrowUpDown,
  ChevronRight,
  Info,
  Building,
  Users,
  Eye,
  X,
  Sparkles,
  Upload,
  Download,
  PlusCircle,
  Clock,
  UserCheck,
  Send,
  Boxes,
  ArrowRight,
  Check,
  FileText,
  RefreshCw,
  Trash2,
  RotateCcw,
} from 'lucide-react';

interface StickerVisaStatusProps {
  stockRecords: StockRecord[];
  categories: CategoriesState;
  officers?: Officer[];
  currentRole?: UserRole;
  userName?: string;
  onAddStockRecord?: (record: StockRecord) => void;
  onAddK1StockRecord?: (record: StockRecord) => void;
  onBatchImportStockRecords?: (records: StockRecord[], overwrite?: boolean) => void;
  onReturnIssuedStock?: (id: string) => void;
  onReturnIssuedStockBatch?: (ids: string[]) => void;
  onReturnIssuedStockByVisaType?: (visaType: string) => void;
  onClearAllActualStock?: () => void;
  onDeleteActualStockRecord?: (id: string) => void;
  onDeleteBatchActualStockRecords?: (ids: string[]) => void;
  onDeleteActualStockByRange?: (visaType: string, startSerial: string, endSerial: string) => void;
  onDeleteStockRecord?: (id: string) => void;
  onDeleteBatchStockRecords?: (ids: string[]) => void;
  onShowToast?: (message: string, type?: 'success' | 'error' | 'info') => void;
  onNavigateToForm?: () => void;
  onClose?: () => void;
}

const DEFAULT_VISA_TYPES = ['T', 'T1', 'T2', 'T3', 'E', 'E1', 'E2', 'E3', 'D', 'K', 'A', 'B', 'C'];

// Baseline Dec 2018 (Empty by default for fresh user imports)
const DEC_2018_STICKER_BASELINES: Record<string, number> = {};

export interface BookDetail {
  bookIndex: number;
  visaType: string;
  startSerial: string;
  endSerial: string;
  sheetCount: number;
  isPartial: boolean;
  rangeIndex: number;
}

export interface SerialRangeInterval {
  start: bigint;
  end: bigint;
  startSerial: string;
  endSerial: string;
  count: number;
  prefix: string;
  padLength: number;
}

export interface TransactionRangeAudit {
  id: string;
  date: string;
  operationType: string;
  operationLabel: string;
  sourceOrTeam: string;
  startSerial: string;
  endSerial: string;
  sheets: number;
  type: 'inflow' | 'outflow';
  visaType?: string;
  record?: StockRecord;
}

export interface VisaStatusItem {
  visaType: string;
  beginningStock: number;
  receivedK1: number;
  returnedFromTeam: number;
  issuedToTeam: number;
  damagedK2: number;
  testPrintK2: number;
  teamUsed: number;
  teamDamaged: number;
  officeRemainingSheets: number;
  officeRemainingBooks: number;
  officeRemainingLoose: number;
  teamRemainingSheets: number;
  teamRemainingBooks: number;
  teamRemainingLoose: number;
  totalRemainingSheets: number;
  totalRemainingBooks: number;
  totalRemainingLoose: number;
  startSerial: string;
  endSerial: string;
  serialRangeDisplay: string;
  remainingRanges: SerialRangeInterval[];
  inflowAudits: TransactionRangeAudit[];
  outflowAudits: TransactionRangeAudit[];
  books: BookDetail[];
  status: 'sufficient' | 'medium' | 'low' | 'out';
}

// Helper: Parse serial number into prefix, numeric bigint value, and pad length
function parseSerialNumber(serial: string): { prefix: string; num: bigint; padLength: number; raw: string } | null {
  if (!serial || typeof serial !== 'string') return null;
  const clean = serial.trim();
  if (!clean) return null;

  // Match leading non-digits (prefix) and trailing digits
  const match = clean.match(/^([^\d]*)(\d+)$/);
  if (match) {
    const prefix = match[1] || '';
    const digitsStr = match[2];
    try {
      const num = BigInt(digitsStr);
      return {
        prefix,
        num,
        padLength: digitsStr.length,
        raw: clean,
      };
    } catch {
      return null;
    }
  }

  // Pure digits or other format
  const digitsOnly = clean.replace(/\D/g, '');
  if (digitsOnly) {
    try {
      const num = BigInt(digitsOnly);
      const prefix = clean.replace(/\d+$/, '');
      return {
        prefix,
        num,
        padLength: digitsOnly.length,
        raw: clean,
      };
    } catch {
      return null;
    }
  }

  return null;
}

// Helper: Format bigint back to formatted serial string
function formatSerialNumber(prefix: string, num: bigint, padLength: number): string {
  const digitsStr = num.toString().padStart(padLength, '0');
  return `${prefix}${digitsStr}`;
}

// Exact Interval Subtraction Algorithm
function subtractIntervals(
  inflows: Array<{ start: bigint; end: bigint; prefix: string; padLength: number }>,
  outStart: bigint,
  outEnd: bigint
): Array<{ start: bigint; end: bigint; prefix: string; padLength: number }> {
  const result: Array<{ start: bigint; end: bigint; prefix: string; padLength: number }> = [];

  for (const interval of inflows) {
    const { start: inS, end: inE, prefix, padLength } = interval;

    // 1. No overlap
    if (outEnd < inS || outStart > inE) {
      result.push(interval);
      continue;
    }

    // 2. Outflow covers beginning of inflow -> keep remainder on right
    if (outStart <= inS && outEnd < inE) {
      result.push({
        start: outEnd + 1n,
        end: inE,
        prefix,
        padLength,
      });
      continue;
    }

    // 3. Outflow covers end of inflow -> keep remainder on left
    if (outStart > inS && outEnd >= inE) {
      result.push({
        start: inS,
        end: outStart - 1n,
        prefix,
        padLength,
      });
      continue;
    }

    // 4. Outflow punches a hole in the middle -> split into two intervals
    if (outStart > inS && outEnd < inE) {
      result.push({
        start: inS,
        end: outStart - 1n,
        prefix,
        padLength,
      });
      result.push({
        start: outEnd + 1n,
        end: inE,
        prefix,
        padLength,
      });
      continue;
    }

    // 5. Outflow completely covers inflow (outStart <= inS && outEnd >= inE) -> interval is completely consumed
  }

  return result;
}

export const StickerVisaStatus: React.FC<StickerVisaStatusProps> = ({
  stockRecords,
  categories,
  officers = [],
  userName = '',
  onAddStockRecord,
  onAddK1StockRecord,
  onBatchImportStockRecords,
  onReturnIssuedStock,
  onReturnIssuedStockBatch,
  onReturnIssuedStockByVisaType,
  onClearAllActualStock,
  onDeleteActualStockRecord,
  onDeleteBatchActualStockRecords,
  onDeleteActualStockByRange,
  onDeleteStockRecord,
  onDeleteBatchStockRecords,
  onShowToast,
  onNavigateToForm,
  onClose,
}) => {
  // Navigation View Modes: rangeTable (matching user's image), summary, booksTable, transactions
  const [activeTab, setActiveTab] = useState<'rangeTable' | 'summary' | 'booksTable' | 'transactions'>('rangeTable');

  // Filters & State
  const [selectedStockScope, setSelectedStockScope] = useState<'office' | 'teams' | 'total'>('office');
  const [filterByDate, setFilterByDate] = useState<boolean>(false);
  const [asOfDate, setAsOfDate] = useState<string>(() => {
    // Default to latest record date or today
    const stickerDates = stockRecords
      .filter((r) => r.stockType === 'sticker' && r.date)
      .map((r) => r.date)
      .sort();
    return stickerDates.length > 0 ? stickerDates[stickerDates.length - 1] : new Date().toISOString().slice(0, 10);
  });
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [selectedVisaTypeFilter, setSelectedVisaTypeFilter] = useState<string>('ALL');
  const [selectedBookModalVisa, setSelectedBookModalVisa] = useState<VisaStatusItem | null>(null);
  const [modalTab, setModalTab] = useState<'books' | 'audit'>('books');

  // Actual Stock Deletion States (លុបតែក្នុងស្តុកជាក់ស្តែង K2 មិនប៉ះពាល់ទិន្នន័យសន្លឹកទិដ្ឋាការ)
  const [isClearAllConfirmOpen, setIsClearAllConfirmOpen] = useState<boolean>(false);
  const [isManageDeleteModalOpen, setIsManageDeleteModalOpen] = useState<boolean>(false);
  const [manageSearch, setManageSearch] = useState<string>('');
  const [manageVisaFilter, setManageVisaFilter] = useState<string>('ALL');
  const [selectedManageIds, setSelectedManageIds] = useState<string[]>([]);
  const [deletingRecord, setDeletingRecord] = useState<StockRecord | null>(null);
  const [deletingRange, setDeletingRange] = useState<{ visaType: string; startSerial: string; endSerial: string; count: number } | null>(null);
  const [isBatchDeleteConfirmOpen, setIsBatchDeleteConfirmOpen] = useState<boolean>(false);

  // Pagination for Books Data Table
  const [currentPage, setCurrentPage] = useState<number>(1);
  const [pageSize, setPageSize] = useState<number>(25);

  // Quick Issue to Team Modal State
  const [isIssueModalOpen, setIsIssueModalOpen] = useState<boolean>(false);
  const [issueVisaType, setIssueVisaType] = useState<string>('T');
  const [issueQuantitySheets, setIssueQuantitySheets] = useState<number>(50);
  const [issueStartSerial, setIssueStartSerial] = useState<string>('');
  const [issueEndSerial, setIssueEndSerial] = useState<string>('');
  const [issueDate, setIssueDate] = useState<string>(new Date().toISOString().slice(0, 10));
  const [issueTime, setIssueTime] = useState<string>('08:30');
  const [issueTeamId, setIssueTeamId] = useState<string>('');
  const [issueRequesterRank, setIssueRequesterRank] = useState<string>('');
  const [issueRequesterName, setIssueRequesterName] = useState<string>('');
  const [issueCollectorName, setIssueCollectorName] = useState<string>('');
  const [issueCollectorRole, setIssueCollectorRole] = useState<string>('');

  // Add Stock from K1 Modal State
  const [isAddK1ModalOpen, setIsAddK1ModalOpen] = useState<boolean>(false);
  const [k1VisaType, setK1VisaType] = useState<string>('T');
  const [k1Date, setK1Date] = useState<string>(new Date().toISOString().slice(0, 10));
  const [k1Time, setK1Time] = useState<string>('08:00');
  const [k1QuantityBundles, setK1QuantityBundles] = useState<string>('1');
  const [k1TotalSheets, setK1TotalSheets] = useState<string>('50');
  const [k1StartSerial, setK1StartSerial] = useState<string>('');
  const [k1EndSerial, setK1EndSerial] = useState<string>('');
  const [k1Sender, setK1Sender] = useState<string>('ក១');
  const [k1Receiver, setK1Receiver] = useState<string>(userName || '');
  const [k1Notes, setK1Notes] = useState<string>('');

  // Return Issued Data Modal State
  const [isReturnModalOpen, setIsReturnModalOpen] = useState<boolean>(false);
  const [returnVisaTypeFilter, setReturnVisaTypeFilter] = useState<string>('ALL');
  const [selectedRecordToReturn, setSelectedRecordToReturn] = useState<StockRecord | null>(null);
  const [isReturnAllConfirmOpen, setIsReturnAllConfirmOpen] = useState<boolean>(false);
  const [returnTypeConfirm, setReturnTypeConfirm] = useState<string | null>(null);

  // Excel Import Modal State
  const [isImportModalOpen, setIsImportModalOpen] = useState<boolean>(false);
  const [importFile, setImportFile] = useState<File | null>(null);
  const [parsedImportRows, setParsedImportRows] = useState<StockRecord[]>([]);
  const [importProcessing, setImportProcessing] = useState<boolean>(false);
  const [importDragActive, setImportDragActive] = useState<boolean>(false);
  const [importMode, setImportMode] = useState<'overwrite' | 'append'>('append');

  // Signatures for printing
  const [signerTitle, setSignerTitle] = useState<string>('អ្នកធ្វើតារាង');
  const [signerName, setSignerName] = useState<string>(userName || 'អ៊ុក រ័ត្នបញ្ញា');
  const [signerRank, setSignerRank] = useState<string>('អនុសេនីយ៍ឯក');

  // Visa Types list
  const allVisaTypes = useMemo(() => {
    const set = new Set(DEFAULT_VISA_TYPES);
    stockRecords
      .filter((r) => r.stockType === 'sticker' && r.visaType)
      .forEach((r) => {
        const v = (r.visaType || '').trim().toUpperCase();
        if (v) set.add(v);
      });
    return Array.from(set);
  }, [stockRecords]);

  // List of all issued records for return functionality
  const issuedRecords = useMemo(() => {
    return stockRecords.filter(
      (r) =>
        r.stockType === 'sticker' &&
        (r.operationType === 'issueTeam' || r.sourceFrom?.includes('បើកផ្តល់'))
    );
  }, [stockRecords]);

  // Filtered issued records for the return modal
  const filteredIssuedRecords = useMemo(() => {
    if (returnVisaTypeFilter === 'ALL') return issuedRecords;
    return issuedRecords.filter(
      (r) => (r.visaType || '').trim().toUpperCase() === returnVisaTypeFilter.trim().toUpperCase()
    );
  }, [issuedRecords, returnVisaTypeFilter]);

  // Handlers for returning issued stock
  const handleReturnSingle = (record: StockRecord) => {
    if (onReturnIssuedStock) {
      onReturnIssuedStock(record.id);
    } else if (onDeleteStockRecord) {
      onDeleteStockRecord(record.id);
      onShowToast?.('បានបង្វិលទិន្នន័យបើកផ្តល់ត្រឡប់ចូលស្តុកជាក់ស្តែង និងទិន្នន័យ (cEA) រួចរាល់!', 'success');
    }
    setSelectedRecordToReturn(null);
  };

  const handleReturnBatch = (ids: string[]) => {
    if (onReturnIssuedStockBatch) {
      onReturnIssuedStockBatch(ids);
    } else if (onDeleteBatchStockRecords) {
      onDeleteBatchStockRecords(ids);
      onShowToast?.(`បានបង្វិលទិន្នន័យបើកផ្តល់ចំនួន ${ids.length} ត្រឡប់ចូលស្តុកជាក់ស្តែង និងទិន្នន័យ (cEA) រួចរាល់!`, 'success');
    } else if (onDeleteStockRecord) {
      ids.forEach((id) => onDeleteStockRecord(id));
      onShowToast?.(`បានបង្វិលទិន្នន័យបើកផ្តល់ចំនួន ${ids.length} ត្រឡប់ចូលស្តុកជាក់ស្តែង និងទិន្នន័យ (cEA) រួចរាល់!`, 'success');
    }
    setIsReturnAllConfirmOpen(false);
    setIsReturnModalOpen(false);
  };

  const handleReturnByVisaType = (visaType: string) => {
    if (onReturnIssuedStockByVisaType) {
      onReturnIssuedStockByVisaType(visaType);
    } else {
      const targetType = (visaType || '').trim().toUpperCase();
      const idsToReturn = issuedRecords
        .filter((r) => (r.visaType || '').trim().toUpperCase() === targetType)
        .map((r) => r.id);
      if (idsToReturn.length > 0) {
        handleReturnBatch(idsToReturn);
      }
    }
    setReturnTypeConfirm(null);
  };

  // Filtered records for Actual Stock Management & Deletion Modal
  const manageStockRecords = useMemo(() => {
    return stockRecords.filter((r) => {
      if (r.stockType !== 'sticker') return false;
      if (manageVisaFilter !== 'ALL') {
        if ((r.visaType || '').trim().toUpperCase() !== manageVisaFilter.trim().toUpperCase()) return false;
      }
      if (manageSearch.trim()) {
        const q = manageSearch.toLowerCase();
        const vt = (r.visaType || '').toLowerCase();
        const start = (r.startSerial || '').toLowerCase();
        const end = (r.endSerial || '').toLowerCase();
        const date = (r.date || '').toLowerCase();
        const src = (r.sourceFrom || '').toLowerCase();
        const op = (r.operationType || '').toLowerCase();
        const team = (r.teamName || r.teamId || '').toLowerCase();
        const note = (r.notes || '').toLowerCase();
        if (
          !vt.includes(q) &&
          !start.includes(q) &&
          !end.includes(q) &&
          !date.includes(q) &&
          !src.includes(q) &&
          !op.includes(q) &&
          !team.includes(q) &&
          !note.includes(q)
        ) {
          return false;
        }
      }
      return true;
    });
  }, [stockRecords, manageVisaFilter, manageSearch]);

  const handleDeleteRecordFromActualStock = (recordId: string) => {
    if (onDeleteActualStockRecord) {
      onDeleteActualStockRecord(recordId);
    } else if (onDeleteStockRecord) {
      onDeleteStockRecord(recordId);
    }
    setSelectedManageIds((prev) => prev.filter((id) => id !== recordId));
    setDeletingRecord(null);
  };

  const handleBatchDeleteFromActualStock = (ids: string[]) => {
    if (onDeleteBatchActualStockRecords) {
      onDeleteBatchActualStockRecords(ids);
    } else if (onDeleteBatchStockRecords) {
      onDeleteBatchStockRecords(ids);
    } else if (onDeleteStockRecord) {
      ids.forEach((id) => onDeleteStockRecord(id));
    }
    setSelectedManageIds([]);
    setIsBatchDeleteConfirmOpen(false);
  };

  const handleDeleteRangeFromActualStock = (range: { visaType: string; startSerial: string; endSerial: string; count: number }) => {
    if (onDeleteActualStockByRange) {
      onDeleteActualStockByRange(range.visaType, range.startSerial, range.endSerial);
    } else {
      const cleanVt = range.visaType.trim().toUpperCase();
      const pStart = parseSerialNumber(range.startSerial);
      const pEnd = parseSerialNumber(range.endSerial);
      const matchingIds = stockRecords
        .filter((r) => {
          if (r.stockType !== 'sticker') return false;
          const rVt = (r.visaType || '').trim().toUpperCase();
          if (rVt !== cleanVt) return false;
          if (pStart && pEnd && r.startSerial) {
            const rStart = parseSerialNumber(r.startSerial);
            const rEnd = r.endSerial ? parseSerialNumber(r.endSerial) : rStart;
            if (rStart && rEnd && rStart.num >= pStart.num && rEnd.num <= pEnd.num) {
              return true;
            }
          }
          return r.startSerial === range.startSerial && r.endSerial === range.endSerial;
        })
        .map((r) => r.id);

      if (matchingIds.length > 0) {
        if (onDeleteBatchActualStockRecords) {
          onDeleteBatchActualStockRecords(matchingIds);
        } else if (onDeleteBatchStockRecords) {
          onDeleteBatchStockRecords(matchingIds);
        } else if (onDeleteStockRecord) {
          matchingIds.forEach((id) => onDeleteStockRecord(id));
        }
      }
    }
    setDeletingRange(null);
  };

  // Calculate Visa Status Items with Exact Serial Range Interval Deduction
  const statusItems: VisaStatusItem[] = useMemo(() => {
    return allVisaTypes.map((visaType) => {
      const baseline = DEC_2018_STICKER_BASELINES[visaType] || 0;
      let oldStockTotal = 0;
      let receivedK1 = 0;
      let returnedFromTeam = 0;
      let issuedToTeam = 0;
      let damagedK2 = 0;
      let testPrintK2 = 0;
      let teamUsed = 0;
      let teamDamaged = 0;

      const recordsBeforeOrOnDate = stockRecords.filter(
        (r) =>
          r.stockType === 'sticker' &&
          (r.visaType || '').trim().toUpperCase() === visaType.toUpperCase() &&
          (!filterByDate || !r.date || r.date <= asOfDate)
      );

      // Raw interval records
      const rawInflowIntervals: Array<{
        start: bigint;
        end: bigint;
        prefix: string;
        padLength: number;
        record: StockRecord;
      }> = [];

      const rawOutflowIntervals: Array<{
        start: bigint;
        end: bigint;
        prefix: string;
        padLength: number;
        record: StockRecord;
      }> = [];

      const inflowAudits: TransactionRangeAudit[] = [];
      const outflowAudits: TransactionRangeAudit[] = [];

      recordsBeforeOrOnDate.forEach((r) => {
        const sheets = r.totalSheets || (r.quantityBundles ? r.quantityBundles * 50 : 0);
        const isTeam =
          r.operationType === 'useTeam' ||
          r.operationType === 'issueTeam' ||
          r.operationType === 'returnTeam' ||
          r.operationType === 'oldStockTeam';

        const isOldStock =
          r.operationType === 'oldStockK2' ||
          r.operationType === 'oldStock' ||
          r.operationType === 'initial' ||
          r.sourceFrom?.includes('ស្តុកចាស់') ||
          r.sourceFrom?.includes('សន្និធិដើម');

        const isReturn =
          r.operationType === 'returnTeam' ||
          r.operationType === 'returned' ||
          r.sourceFrom?.includes('បង្វិល') ||
          r.operationType?.includes('បង្វិល');

        const isIssue =
          r.operationType === 'issueTeam' ||
          r.sourceFrom?.includes('បើកផ្តល់') ||
          r.operationType?.includes('issue');

        const isTest =
          !isTeam &&
          (r.operationType === 'testPrintK2' ||
            r.operationType === 'testK2' ||
            r.operationType?.includes('សាក'));

        const isDamagedK2 =
          !isTeam &&
          (r.operationType === 'damaged' ||
            r.operationType === 'invalid' ||
            r.sourceFrom?.includes('ខូច') ||
            r.operationType?.includes('ខូច'));

        const isOpenK1 =
          !isTeam &&
          (r.operationType === 'openK1' ||
            r.operationType === 'receive' ||
            r.operationType === 'received' ||
            r.operationType === 'in' ||
            r.sourceFrom?.includes('ក១'));

        const isTeamUse = r.operationType === 'useTeam' || r.sourceFrom?.includes('ប្រើប្រាស់');
        const isTeamDamage = r.operationType === 'damagedTeam' || r.operationType === 'missingTeam';

        if (isOldStock) {
          oldStockTotal += sheets;
        } else if (isReturn) {
          returnedFromTeam += sheets;
        } else if (isIssue) {
          issuedToTeam += sheets;
        } else if (isTest) {
          testPrintK2 += sheets;
        } else if (isDamagedK2) {
          damagedK2 += sheets;
        } else if (isTeamUse) {
          teamUsed += sheets;
        } else if (isTeamDamage) {
          teamDamaged += sheets;
        } else if (isOpenK1) {
          receivedK1 += sheets;
        }

        // Determine if this record is an Inflow or Outflow for the current selected scope
        let isScopeInflow = false;
        let isScopeOutflow = false;
        let opLabel = '';

        if (selectedStockScope === 'office') {
          if (isOpenK1 || isOldStock || isReturn) {
            isScopeInflow = true;
            opLabel = isOpenK1 ? 'បើកពី ក១' : isOldStock ? 'ស្តុកចាស់ ក២' : 'បង្វិលពីក្រុម';
          } else if (isIssue || isDamagedK2 || isTest) {
            isScopeOutflow = true;
            opLabel = isIssue ? 'បើកជូនក្រុម' : isDamagedK2 ? 'ខូច K2' : 'បោះពុម្ពសាក';
          }
        } else if (selectedStockScope === 'teams') {
          if (isIssue) {
            isScopeInflow = true;
            opLabel = 'បានទទួលពី K2';
          } else if (isReturn || isTeamUse || isTeamDamage) {
            isScopeOutflow = true;
            opLabel = isReturn ? 'បង្វិលជូន K2' : isTeamUse ? 'ប្រើប្រាស់រួច' : 'ខូចបាត់';
          }
        } else {
          // Total
          if (isOpenK1 || isOldStock) {
            isScopeInflow = true;
            opLabel = isOpenK1 ? 'បើកពី ក១' : 'ស្តុកដើម';
          } else if (isTeamUse || isTeamDamage || isDamagedK2 || isTest) {
            isScopeOutflow = true;
            opLabel = isTeamUse ? 'ប្រើប្រាស់រួច' : isDamagedK2 ? 'ខូច K2' : 'ខូចតាមក្រុម';
          }
        }

        // Extract serial interval if present
        if (r.startSerial) {
          const parsedStart = parseSerialNumber(r.startSerial);
          if (parsedStart) {
            let parsedEnd = r.endSerial ? parseSerialNumber(r.endSerial) : null;
            let endNum: bigint;

            if (parsedEnd && parsedEnd.num >= parsedStart.num) {
              endNum = parsedEnd.num;
            } else if (sheets > 0) {
              endNum = parsedStart.num + BigInt(sheets) - 1n;
            } else {
              endNum = parsedStart.num;
            }

            const itemSheets = Number(endNum - parsedStart.num + 1n);
            const formattedEnd = parsedEnd ? parsedEnd.raw : formatSerialNumber(parsedStart.prefix, endNum, parsedStart.padLength);

            if (isScopeInflow) {
              rawInflowIntervals.push({
                start: parsedStart.num,
                end: endNum,
                prefix: parsedStart.prefix,
                padLength: parsedStart.padLength,
                record: r,
              });
              inflowAudits.push({
                id: r.id,
                date: r.date || '',
                operationType: r.operationType || '',
                operationLabel: opLabel,
                sourceOrTeam: r.visaTeamRobokName || r.sourceFrom || '-',
                startSerial: parsedStart.raw,
                endSerial: formattedEnd,
                sheets: itemSheets,
                type: 'inflow',
                visaType: r.visaType || visaType,
                record: r,
              });
            } else if (isScopeOutflow) {
              rawOutflowIntervals.push({
                start: parsedStart.num,
                end: endNum,
                prefix: parsedStart.prefix,
                padLength: parsedStart.padLength,
                record: r,
              });
              outflowAudits.push({
                id: r.id,
                date: r.date || '',
                operationType: r.operationType || '',
                operationLabel: opLabel,
                sourceOrTeam: r.visaTeamRobokName || r.sourceFrom || '-',
                startSerial: parsedStart.raw,
                endSerial: formattedEnd,
                sheets: itemSheets,
                type: 'outflow',
                visaType: r.visaType || visaType,
                record: r,
              });
            }
          }
        }
      });

      const beginningStock = oldStockTotal > 0 ? oldStockTotal : baseline;

      // Mathematical Sheet Counts
      const officeRemainingSheets = Math.max(
        0,
        beginningStock + receivedK1 + returnedFromTeam - issuedToTeam - damagedK2 - testPrintK2
      );
      const officeRemainingBooks = Math.floor(officeRemainingSheets / 50);
      const officeRemainingLoose = officeRemainingSheets % 50;

      const teamRemainingSheets = Math.max(0, issuedToTeam - returnedFromTeam - teamUsed - teamDamaged);
      const teamRemainingBooks = Math.floor(teamRemainingSheets / 50);
      const teamRemainingLoose = teamRemainingSheets % 50;

      const totalRemainingSheets = officeRemainingSheets + teamRemainingSheets;
      const totalRemainingBooks = Math.floor(totalRemainingSheets / 50);
      const totalRemainingLoose = totalRemainingSheets % 50;

      // -------------------------------------------------------------
      // EXACT INTERVAL DEDUCTION (K1 Inflow Range - Team Issued Range)
      // -------------------------------------------------------------
      // 1. Sort & merge overlapping/contiguous inflows
      rawInflowIntervals.sort((a, b) => (a.start < b.start ? -1 : a.start > b.start ? 1 : 0));

      let mergedInflows: Array<{ start: bigint; end: bigint; prefix: string; padLength: number }> = [];
      for (const inf of rawInflowIntervals) {
        if (mergedInflows.length === 0) {
          mergedInflows.push({ start: inf.start, end: inf.end, prefix: inf.prefix, padLength: inf.padLength });
        } else {
          const last = mergedInflows[mergedInflows.length - 1];
          if (inf.start <= last.end + 1n && inf.prefix === last.prefix) {
            if (inf.end > last.end) {
              last.end = inf.end;
            }
          } else {
            mergedInflows.push({ start: inf.start, end: inf.end, prefix: inf.prefix, padLength: inf.padLength });
          }
        }
      }

      // 2. Subtract each outflow interval from the merged inflows
      let activeIntervals = [...mergedInflows];
      for (const out of rawOutflowIntervals) {
        activeIntervals = subtractIntervals(activeIntervals, out.start, out.end);
      }

      // 3. Format remaining intervals
      const remainingRanges: SerialRangeInterval[] = activeIntervals.map((intv) => {
        const count = Number(intv.end - intv.start + 1n);
        return {
          start: intv.start,
          end: intv.end,
          startSerial: formatSerialNumber(intv.prefix, intv.start, intv.padLength),
          endSerial: formatSerialNumber(intv.prefix, intv.end, intv.padLength),
          count,
          prefix: intv.prefix,
          padLength: intv.padLength,
        };
      });

      // 4. Build 50-sheet Books from the remaining intervals (safely capped to prevent browser freeze on huge serial intervals)
      const books: BookDetail[] = [];
      let bookCounter = 1;
      const MAX_BOOKS_PER_TYPE = 200;

      for (let rangeIdx = 0; rangeIdx < remainingRanges.length; rangeIdx++) {
        const range = remainingRanges[rangeIdx];
        let currentStart = range.start;
        // Limit loop iterations to prevent infinite/heavy loops if serial gap is massive
        let loopCount = 0;
        const maxLoops = 200;
        while (currentStart <= range.end && books.length < MAX_BOOKS_PER_TYPE && loopCount < maxLoops) {
          loopCount++;
          const currentEnd = currentStart + 49n <= range.end ? currentStart + 49n : range.end;
          const sheetCount = Number(currentEnd - currentStart + 1n);
          const isPartial = sheetCount < 50;

          books.push({
            bookIndex: bookCounter++,
            visaType,
            startSerial: formatSerialNumber(range.prefix, currentStart, range.padLength),
            endSerial: formatSerialNumber(range.prefix, currentEnd, range.padLength),
            sheetCount,
            isPartial,
            rangeIndex: rangeIdx + 1,
          });

          currentStart = currentEnd + 1n;
        }
        if (books.length >= MAX_BOOKS_PER_TYPE) break;
      }

      // Total sheets derived from exact remaining intervals
      const intervalSumSheets = remainingRanges.reduce((sum, r) => sum + r.count, 0);

      // Determine final active remaining sheets
      const finalRemainingSheets =
        remainingRanges.length > 0 || rawInflowIntervals.length > 0
          ? intervalSumSheets
          : selectedStockScope === 'office'
          ? officeRemainingSheets
          : selectedStockScope === 'teams'
          ? teamRemainingSheets
          : totalRemainingSheets;

      const finalRemainingBooks = Math.floor(finalRemainingSheets / 50);
      const finalRemainingLoose = finalRemainingSheets % 50;

      // Range display string
      let serialRangeDisplay = '-';
      let startSerial = '-';
      let endSerial = '-';

      if (remainingRanges.length === 1) {
        startSerial = remainingRanges[0].startSerial;
        endSerial = remainingRanges[0].endSerial;
        serialRangeDisplay = `${startSerial} ដល់ ${endSerial}`;
      } else if (remainingRanges.length > 1) {
        startSerial = remainingRanges[0].startSerial;
        endSerial = remainingRanges[remainingRanges.length - 1].endSerial;
        serialRangeDisplay = remainingRanges.map((r) => `${r.startSerial} - ${r.endSerial}`).join(', ');
      }

      // Status classification
      let status: 'sufficient' | 'medium' | 'low' | 'out' = 'sufficient';
      if (finalRemainingSheets === 0) status = 'out';
      else if (finalRemainingSheets < 100) status = 'low';
      else if (finalRemainingSheets < 500) status = 'medium';

      return {
        visaType,
        beginningStock,
        receivedK1,
        returnedFromTeam,
        issuedToTeam,
        damagedK2,
        testPrintK2,
        teamUsed,
        teamDamaged,
        officeRemainingSheets: selectedStockScope === 'office' && remainingRanges.length > 0 ? intervalSumSheets : officeRemainingSheets,
        officeRemainingBooks: selectedStockScope === 'office' && remainingRanges.length > 0 ? finalRemainingBooks : officeRemainingBooks,
        officeRemainingLoose: selectedStockScope === 'office' && remainingRanges.length > 0 ? finalRemainingLoose : officeRemainingLoose,
        teamRemainingSheets: selectedStockScope === 'teams' && remainingRanges.length > 0 ? intervalSumSheets : teamRemainingSheets,
        teamRemainingBooks: selectedStockScope === 'teams' && remainingRanges.length > 0 ? finalRemainingBooks : teamRemainingBooks,
        teamRemainingLoose: selectedStockScope === 'teams' && remainingRanges.length > 0 ? finalRemainingLoose : teamRemainingLoose,
        totalRemainingSheets: selectedStockScope === 'total' && remainingRanges.length > 0 ? intervalSumSheets : totalRemainingSheets,
        totalRemainingBooks: selectedStockScope === 'total' && remainingRanges.length > 0 ? finalRemainingBooks : totalRemainingBooks,
        totalRemainingLoose: selectedStockScope === 'total' && remainingRanges.length > 0 ? finalRemainingLoose : totalRemainingLoose,
        startSerial,
        endSerial,
        serialRangeDisplay,
        remainingRanges,
        inflowAudits,
        outflowAudits,
        books,
        status,
      };
    });
  }, [allVisaTypes, stockRecords, asOfDate, selectedStockScope, filterByDate]);

  // Flatten all books for the All Available Books Data Table
  const allAvailableBooks = useMemo(() => {
    const list: BookDetail[] = [];
    statusItems.forEach((item) => {
      if (selectedVisaTypeFilter === 'ALL' || item.visaType === selectedVisaTypeFilter) {
        item.books.forEach((b) => {
          list.push(b);
        });
      }
    });
    return list;
  }, [statusItems, selectedVisaTypeFilter]);

  // Filtered Books Data Table search
  const filteredBooks = useMemo(() => {
    if (!searchQuery.trim()) return allAvailableBooks;
    const q = searchQuery.toLowerCase().trim();
    return allAvailableBooks.filter(
      (b) =>
        b.visaType.toLowerCase().includes(q) ||
        b.startSerial.toLowerCase().includes(q) ||
        b.endSerial.toLowerCase().includes(q) ||
        String(b.bookIndex).includes(q)
    );
  }, [allAvailableBooks, searchQuery]);

  // Paginated Books
  const paginatedBooks = useMemo(() => {
    const start = (currentPage - 1) * pageSize;
    return filteredBooks.slice(start, start + pageSize);
  }, [filteredBooks, currentPage, pageSize]);

  const totalPages = Math.ceil(filteredBooks.length / pageSize) || 1;

  // Filtered Visa Summary Items
  const filteredStatusItems = useMemo(() => {
    return statusItems.filter((item) => {
      const matchType = selectedVisaTypeFilter === 'ALL' || item.visaType === selectedVisaTypeFilter;
      const matchQuery =
        !searchQuery ||
        item.visaType.toLowerCase().includes(searchQuery.toLowerCase()) ||
        item.serialRangeDisplay.toLowerCase().includes(searchQuery.toLowerCase()) ||
        item.startSerial.toLowerCase().includes(searchQuery.toLowerCase()) ||
        item.endSerial.toLowerCase().includes(searchQuery.toLowerCase());
      return matchType && matchQuery;
    });
  }, [statusItems, selectedVisaTypeFilter, searchQuery]);

  // Overall Totals
  const summary = useMemo(() => {
    const totalOfficeSheets = statusItems.reduce((acc, it) => acc + it.officeRemainingSheets, 0);
    const totalOfficeBooks = Math.floor(totalOfficeSheets / 50);
    const totalOfficeLoose = totalOfficeSheets % 50;

    const totalTeamSheets = statusItems.reduce((acc, it) => acc + it.teamRemainingSheets, 0);
    const totalTeamBooks = Math.floor(totalTeamSheets / 50);
    const totalTeamLoose = totalTeamSheets % 50;

    const totalIssuedSheets = statusItems.reduce((acc, it) => acc + it.issuedToTeam, 0);
    const totalReceivedK1 = statusItems.reduce((acc, it) => acc + it.receivedK1, 0);

    const totalSheets =
      selectedStockScope === 'office'
        ? totalOfficeSheets
        : selectedStockScope === 'teams'
        ? totalTeamSheets
        : totalOfficeSheets + totalTeamSheets;

    const totalBooks = Math.floor(totalSheets / 50);
    const totalLoose = totalSheets % 50;

    const totalAllBooksCount = statusItems.reduce((acc, it) => acc + it.books.length, 0);

    return {
      totalOfficeSheets,
      totalOfficeBooks,
      totalOfficeLoose,
      totalTeamSheets,
      totalTeamBooks,
      totalTeamLoose,
      totalIssuedSheets,
      totalReceivedK1,
      totalSheets,
      totalBooks,
      totalLoose,
      totalAllBooksCount,
      countSufficient: statusItems.filter((it) => it.status === 'sufficient').length,
      countMedium: statusItems.filter((it) => it.status === 'medium').length,
      countLow: statusItems.filter((it) => it.status === 'low').length,
      countOut: statusItems.filter((it) => it.status === 'out').length,
    };
  }, [statusItems, selectedStockScope]);

  // Open Quick Issue to Team Modal
  const handleOpenIssueModal = (visaType?: string) => {
    const vType = visaType || (statusItems[0]?.visaType || 'T');
    setIssueVisaType(vType);

    const targetItem = statusItems.find((s) => s.visaType === vType);
    if (targetItem && targetItem.remainingRanges.length > 0) {
      const firstRange = targetItem.remainingRanges[0];
      const startNum = firstRange.start;
      const defaultQty = Math.min(50, Number(firstRange.end - firstRange.start + 1n));
      const endNum = startNum + BigInt(defaultQty) - 1n;

      setIssueQuantitySheets(defaultQty);
      setIssueStartSerial(firstRange.startSerial);
      setIssueEndSerial(formatSerialNumber(firstRange.prefix, endNum, firstRange.padLength));
    } else {
      setIssueQuantitySheets(50);
      setIssueStartSerial('');
      setIssueEndSerial('');
    }

    if (categories.visaTeamsRobok && categories.visaTeamsRobok.length > 0) {
      setIssueTeamId(categories.visaTeamsRobok[0].name);
    }

    setIsIssueModalOpen(true);
  };

  // Change Issue Visa Type in Quick Issue Modal
  const handleIssueVisaTypeChange = (vType: string) => {
    setIssueVisaType(vType);
    const targetItem = statusItems.find((s) => s.visaType === vType);
    if (targetItem && targetItem.remainingRanges.length > 0) {
      const firstRange = targetItem.remainingRanges[0];
      const startNum = firstRange.start;
      const defaultQty = Math.min(50, Number(firstRange.end - firstRange.start + 1n));
      const endNum = startNum + BigInt(defaultQty) - 1n;

      setIssueQuantitySheets(defaultQty);
      setIssueStartSerial(firstRange.startSerial);
      setIssueEndSerial(formatSerialNumber(firstRange.prefix, endNum, firstRange.padLength));
    } else {
      setIssueQuantitySheets(50);
      setIssueStartSerial('');
      setIssueEndSerial('');
    }
  };

  // Handle Sheet Quantity change in Quick Issue Modal
  const handleIssueQuantityChange = (qty: number) => {
    setIssueQuantitySheets(qty);
    if (issueStartSerial) {
      const parsed = parseSerialNumber(issueStartSerial);
      if (parsed && qty > 0) {
        const endNum = parsed.num + BigInt(qty) - 1n;
        setIssueEndSerial(formatSerialNumber(parsed.prefix, endNum, parsed.padLength));
      }
    }
  };

  // Submit Quick Issue to Team
  const handleConfirmIssueToTeam = (e: React.FormEvent) => {
    e.preventDefault();
    if (!issueVisaType || issueQuantitySheets <= 0) {
      onShowToast?.('សូមជ្រើសរើសប្រភេទទិដ្ឋាការ និងចំនួនសន្លឹកត្រឹមត្រូវ!', 'error');
      return;
    }

    const newRecord: StockRecord = {
      id: `stock_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`,
      stockType: 'sticker',
      operationType: 'issueTeam',
      sourceFrom: 'ក១',
      date: issueDate,
      time: issueTime,
      visaType: issueVisaType,
      quantityBundles: Math.ceil(issueQuantitySheets / 50),
      totalSheets: issueQuantitySheets,
      startSerial: issueStartSerial,
      endSerial: issueEndSerial,
      visaTeamRobokName: issueTeamId,
      requestedRankName: issueRequesterRank,
      requesterName: issueRequesterName,
      collectorName: issueCollectorName,
      collectorRoleName: issueCollectorRole,
      createdAt: new Date().toISOString(),
    };

    if (onAddStockRecord) {
      onAddStockRecord(newRecord);
      onShowToast?.(
        `បានបើកផ្តល់ទិដ្ឋាការប្រភេទ ${issueVisaType} ចំនួន ${issueQuantitySheets} សន្លឹកជូន ${issueTeamId || 'ក្រុម'} រួចរាល់!`,
        'success'
      );
    }

    setIsIssueModalOpen(false);
  };

  // Download Sample Excel Template for Stock Import
  const handleDownloadExcelTemplate = () => {
    const headers = ['ប្រភេទ', 'ក្បាល', 'សន្លឹក', 'ចាប់ផ្តើម', 'ដល់លេខ'];

    const sampleRows = [
      ['T', 4, 200, '1903806801', '1903807000'],
      ['T', 1540, 77000, '1903849001', '1903926000'],
      ['E', 10, 500, '2001000001', '2001000500'],
      ['E1', 8, 400, '2101000001', '2101000400'],
      ['K', 20, 1000, '2201000001', '2201001000'],
    ];

    const worksheet = XLSX.utils.aoa_to_sheet([headers, ...sampleRows]);
    worksheet['!cols'] = [
      { wch: 14 },
      { wch: 14 },
      { wch: 16 },
      { wch: 20 },
      { wch: 20 },
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'ស្តុកជាក់ស្តែង');
    XLSX.writeFile(workbook, `Template_Stock_Sticker_Range_${asOfDate}.xlsx`);
    onShowToast?.('បានទាញយកទម្រង់គំរូ Excel (ប្រភេទ | ក្បាល | សន្លឹក | ចាប់ផ្តើម | ដល់លេខ) រួចរាល់!', 'success');
  };

  // Export Current Actual Stock in exact 5-column layout
  const handleExportRangeFormatExcel = () => {
    const headers = ['ប្រភេទ', 'ក្បាល', 'សន្លឹក', 'ចាប់ផ្តើម', 'ដល់លេខ'];
    const rows: any[] = [];

    filteredStatusItems.forEach((item) => {
      if (item.remainingRanges.length > 0) {
        item.remainingRanges.forEach((r) => {
          const sheets = r.count;
          const books = Math.floor(sheets / 50);
          rows.push([
            item.visaType,
            books,
            sheets,
            r.startSerial,
            r.endSerial,
          ]);
        });
      } else {
        const sheets =
          selectedStockScope === 'office'
            ? item.officeRemainingSheets
            : selectedStockScope === 'teams'
            ? item.teamRemainingSheets
            : item.totalRemainingSheets;
        if (sheets > 0) {
          const books = Math.floor(sheets / 50);
          rows.push([
            item.visaType,
            books,
            sheets,
            item.startSerial !== '-' ? item.startSerial : '',
            item.endSerial !== '-' ? item.endSerial : '',
          ]);
        }
      }
    });

    const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);
    worksheet['!cols'] = [
      { wch: 14 },
      { wch: 14 },
      { wch: 16 },
      { wch: 20 },
      { wch: 20 },
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'ចន្លោះស៊េរីស្តុក');
    XLSX.writeFile(workbook, `Stock_Sticker_Ranges_${asOfDate}.xlsx`);
    onShowToast?.('បាននាំចេញទិន្នន័យជាឯកសារ Excel តាមទម្រង់ចន្លោះស៊េរីរួចរាល់!', 'success');
  };

  // Process and Auto-Sort Excel File
  const handleProcessExcelFile = (file: File) => {
    setImportProcessing(true);
    setImportFile(file);

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array' });
        if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
          onShowToast?.('ឯកសារ Excel គ្មាន Sheet ទេ!', 'error');
          setImportProcessing(false);
          return;
        }

        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        const aoa = XLSX.utils.sheet_to_json<any[]>(worksheet, { header: 1, defval: '' });

        if (!aoa || aoa.length === 0) {
          onShowToast?.('ឯកសារ Excel គ្មានទិន្នន័យទេ!', 'error');
          setImportProcessing(false);
          return;
        }

        // Header Detection
        let headerRowIndex = 0;
        let colMap: Record<string, number> = {
          visaType: 0,
          books: -1,
          date: -1,
          time: -1,
          qty: 2,
          startSerial: 3,
          endSerial: 4,
          op: -1,
          source: -1,
          team: -1,
          requester: -1,
          collector: -1,
        };

        // Scan for matching headers
        for (let r = 0; r < Math.min(10, aoa.length); r++) {
          const row = aoa[r];
          if (!Array.isArray(row)) continue;
          let matched = 0;
          row.forEach((cell, idx) => {
            const str = String(cell || '').trim().toLowerCase();
            if (
              str.includes('ទិដ្ឋាការ') ||
              str.includes('visatype') ||
              str.includes('cea') ||
              str.includes('(cea)') ||
              str === 'ប្រភេទ' ||
              str.includes('ប្រភេទ') ||
              str === 'កូដ'
            ) {
              colMap.visaType = idx;
              matched++;
            } else if (str.includes('ក្បាល') || str.includes('book') || str.includes('ក្បាល(៥០សន្លឹក)')) {
              colMap.books = idx;
              matched++;
            } else if (str.includes('កាលបរិច្ឆេទ') || str.includes('date') || str.includes('ថ្ងៃ')) {
              colMap.date = idx;
              matched++;
            } else if (str.includes('ចំនួន') || str.includes('qty') || str.includes('សន្លឹក') || str.includes('sheet')) {
              colMap.qty = idx;
              matched++;
            } else if (str.includes('ចាប់ពី') || str.includes('ចាប់ផ្តើម') || str.includes('start') || str === 'ចាប់ផ្តើម') {
              colMap.startSerial = idx;
              matched++;
            } else if (str.includes('ដល់លេខ') || str.includes('បញ្ចប់') || str.includes('end') || str === 'ដល់លេខ' || str.includes('ដល់')) {
              colMap.endSerial = idx;
              matched++;
            } else if (str.includes('ប្រតិបត្តិការ') || str.includes('op')) {
              colMap.op = idx;
              matched++;
            } else if (str.includes('បើកពី') || str.includes('ប្រភព') || str.includes('source')) {
              colMap.source = idx;
              matched++;
            }
          });

          if (matched >= 2) {
            headerRowIndex = r + 1;
            break;
          }
        }

        const parsedRows: StockRecord[] = [];
        let lastSeenVisaType = '';

        for (let i = headerRowIndex; i < aoa.length; i++) {
          const row = aoa[i];
          if (!row || row.length === 0) continue;

          let rawVisaType = String(row[colMap.visaType] || '').trim().toUpperCase();
          // Support grouped/merged rows where visaType is blank for subsequent ranges of the same type
          if (!rawVisaType && lastSeenVisaType && (row[colMap.startSerial] || row[colMap.qty] || (colMap.books !== -1 && row[colMap.books]))) {
            rawVisaType = lastSeenVisaType;
          }
          if (!rawVisaType) continue;
          lastSeenVisaType = rawVisaType;

          let rawDate = colMap.date !== -1 ? String(row[colMap.date] || '').trim() : '';
          if (!rawDate) rawDate = asOfDate || new Date().toISOString().slice(0, 10);
          else if (rawDate.includes('/')) {
            const parts = rawDate.split('/');
            if (parts.length === 3) {
              rawDate = `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
            }
          }

          let rawBooks = colMap.books !== -1 ? parseInt(String(row[colMap.books] || '0').replace(/\D/g, ''), 10) || 0 : 0;
          let rawQty = colMap.qty !== -1 ? parseInt(String(row[colMap.qty] || '0').replace(/\D/g, ''), 10) || 0 : 0;
          let startSerial = colMap.startSerial !== -1 ? String(row[colMap.startSerial] || '').trim() : '';
          let endSerial = colMap.endSerial !== -1 ? String(row[colMap.endSerial] || '').trim() : '';

          if (rawQty === 0 && rawBooks > 0) {
            rawQty = rawBooks * 50;
          }

          if (startSerial && !endSerial && rawQty > 0) {
            const parsed = parseSerialNumber(startSerial);
            if (parsed) {
              const endNum = parsed.num + BigInt(rawQty) - 1n;
              endSerial = formatSerialNumber(parsed.prefix, endNum, parsed.padLength);
            }
          } else if (startSerial && endSerial && rawQty === 0) {
            const pStart = parseSerialNumber(startSerial);
            const pEnd = parseSerialNumber(endSerial);
            if (pStart && pEnd && pEnd.num >= pStart.num) {
              rawQty = Number(pEnd.num - pStart.num + 1n);
              if (rawBooks === 0) rawBooks = Math.floor(rawQty / 50);
            }
          }

          const opStr = colMap.op !== -1 ? String(row[colMap.op] || '').trim().toLowerCase() : '';
          let opType: StockRecord['operationType'] = 'openK1';
          if (opStr.includes('ចាស់') || opStr.includes('old')) opType = 'oldStockK2';
          else if (opStr.includes('បើកផ្តល់') || opStr.includes('issue')) opType = 'issueTeam';
          else if (opStr.includes('បង្វិល') || opStr.includes('return')) opType = 'returnTeam';
          else if (opStr.includes('ខូច') || opStr.includes('damage')) opType = 'damaged';

          const record: StockRecord = {
            id: `import_${Date.now()}_${i}_${Math.random().toString(36).substring(2, 6)}`,
            stockType: 'sticker',
            operationType: opType,
            sourceFrom: colMap.source !== -1 ? String(row[colMap.source] || 'ក១').trim() : 'ក១',
            date: rawDate,
            time: colMap.time !== -1 ? String(row[colMap.time] || '08:00').trim() : '08:00',
            visaType: rawVisaType,
            quantityBundles: rawBooks > 0 ? rawBooks : Math.ceil((rawQty || 50) / 50),
            totalSheets: rawQty || (rawBooks > 0 ? rawBooks * 50 : 50),
            startSerial,
            endSerial,
            visaTeamRobokName: colMap.team !== -1 ? String(row[colMap.team] || '').trim() : '',
            requesterName: colMap.requester !== -1 ? String(row[colMap.requester] || '').trim() : '',
            collectorName: colMap.collector !== -1 ? String(row[colMap.collector] || '').trim() : '',
            createdAt: new Date().toISOString(),
          };

          parsedRows.push(record);
        }

        // =========================================================================
        // AUTO-SORT & ORGANIZE IMPORTED DATA (តម្រៀបអោយស្អាត)
        // 1. Sort by Visa Type priority (T, T1, T2, T3, E, E1, E2, E3, D, K, A, B, C...)
        // 2. Sort by Serial Number ascending (Numeric bigint comparison)
        // 3. Sort by Date ascending
        // =========================================================================
        const visaPriorityMap: Record<string, number> = {};
        DEFAULT_VISA_TYPES.forEach((vt, idx) => {
          visaPriorityMap[vt] = idx;
        });

        parsedRows.sort((a, b) => {
          const vA = (a.visaType || '').toUpperCase();
          const vB = (b.visaType || '').toUpperCase();
          const pA = visaPriorityMap[vA] !== undefined ? visaPriorityMap[vA] : 999;
          const pB = visaPriorityMap[vB] !== undefined ? visaPriorityMap[vB] : 999;

          if (pA !== pB) return pA - pB;

          // Compare serial start numbers
          const sA = parseSerialNumber(a.startSerial || '');
          const sB = parseSerialNumber(b.startSerial || '');
          if (sA && sB) {
            if (sA.num < sB.num) return -1;
            if (sA.num > sB.num) return 1;
          }

          // Compare dates
          return (a.date || '').localeCompare(b.date || '');
        });

        setParsedImportRows(parsedRows);
        setImportProcessing(false);
        onShowToast?.(`បានអានទិន្នន័យពី Excel បានចំនួន ${parsedRows.length} ជួរ និងបានតម្រៀបស្អាតរួចរាល់!`, 'success');
      } catch (err: any) {
        setImportProcessing(false);
        onShowToast?.(`មានបញ្ហាក្នុងការអាន Excel: ${err?.message || 'ទម្រង់មិនត្រឹមត្រូវ'}`, 'error');
      }
    };

    reader.readAsArrayBuffer(file);
  };

  // Auto-calculate end serial for K1 when start serial or sheet quantity changes
  React.useEffect(() => {
    if (!k1StartSerial.trim()) {
      setK1EndSerial('');
      return;
    }
    const sheets = parseInt(k1TotalSheets, 10) || 0;
    if (sheets <= 0) return;
    const parsed = parseSerialNumber(k1StartSerial.trim());
    if (parsed) {
      const endNum = parsed.num + BigInt(sheets) - 1n;
      const formatted = formatSerialNumber(parsed.prefix, endNum, parsed.padLength);
      setK1EndSerial(formatted);
    }
  }, [k1StartSerial, k1TotalSheets]);

  const handleK1BundlesChange = (val: string) => {
    setK1QuantityBundles(val);
    const bundles = parseInt(val, 10) || 0;
    if (bundles > 0) {
      setK1TotalSheets(String(bundles * 50));
    }
  };

  const handleSubmitAddK1Stock = (e: React.FormEvent) => {
    e.preventDefault();
    const sheets = parseInt(k1TotalSheets, 10) || (parseInt(k1QuantityBundles, 10) || 0) * 50;
    if (sheets <= 0) {
      onShowToast?.('សូមបញ្ចូលចំនួនសន្លឹកអោយបានត្រឹមត្រូវ!', 'error');
      return;
    }
    const newRecord: StockRecord = {
      id: `k1_stock_${Date.now()}_${Math.random().toString(36).substr(2, 5)}`,
      stockType: 'sticker',
      operationType: 'openK1',
      sourceFrom: k1Sender.trim() || 'ក១',
      date: k1Date,
      time: k1Time,
      visaType: k1VisaType.trim().toUpperCase(),
      quantityBundles: parseInt(k1QuantityBundles, 10) || Math.ceil(sheets / 50),
      totalSheets: sheets,
      startSerial: k1StartSerial.trim() || undefined,
      endSerial: k1EndSerial.trim() || undefined,
      requesterName: k1Sender.trim() || 'ក១',
      collectorName: k1Receiver.trim() || userName,
      note: k1Notes.trim() || undefined,
      createdAt: new Date().toISOString(),
      createdBy: userName,
    };

    if (onAddK1StockRecord) {
      onAddK1StockRecord(newRecord);
    } else if (onAddStockRecord) {
      onAddStockRecord(newRecord);
    }

    onShowToast?.(`បានបញ្ចូលស្តុកពី ក១ (${newRecord.visaType}: ${sheets.toLocaleString()} សន្លឹក) ចូលក្នុងស្តុកជាក់ស្តែង K2 ជោគជ័យ!`, 'success');
    setIsAddK1ModalOpen(false);
    // Reset form
    setK1StartSerial('');
    setK1EndSerial('');
    setK1QuantityBundles('1');
    setK1TotalSheets('50');
    setK1Notes('');
  };

  // Confirm Import
  const handleConfirmBatchImport = () => {
    if (parsedImportRows.length === 0) {
      onShowToast?.('មិនមានទិន្នន័យត្រូវនាំចូលទេ!', 'error');
      return;
    }

    if (importMode === 'overwrite') {
      const confirmClean = window.confirm(
        '⚠️ ការព្រមានសុវត្ថិភាពទិន្នន័យ ៖\nលោកអ្នកបានជ្រើសរើស «ជំនួសទិន្នន័យចាស់ទាំងអស់ (Clean Replace)»។\nសកម្មភាពនេះនឹងជម្រះទិន្នន័យចាស់ទាំងអស់ក្នុងស្តុកជាក់ស្តែង K2 ហើយជំនួសដោយទិន្នន័យ Excel ថ្មីនេះ។\n\nតើលោកអ្នកពិតជាចង់ជំនួសទិន្នន័យចាស់មែនទេ?\n(ប្រសិនបើគ្រាន់តែចង់បន្ថែម សូមចុច Cancel ហើយជ្រើសរើសជម្រើស «បន្ថែមលើស្តុកដែលមានស្រាប់ (Append)»)'
      );
      if (!confirmClean) {
        return;
      }
    }

    if (onBatchImportStockRecords) {
      onBatchImportStockRecords(parsedImportRows, importMode === 'overwrite');
    } else if (onAddStockRecord) {
      parsedImportRows.forEach((r) => onAddStockRecord(r));
    }

    // Auto-update asOfDate to the latest date among the imported records so they are immediately visible
    const importedDates = parsedImportRows.filter((r) => r.date).map((r) => r.date as string).sort();
    if (importedDates.length > 0) {
      const maxImportDate = importedDates[importedDates.length - 1];
      if (maxImportDate > asOfDate) {
        setAsOfDate(maxImportDate);
      }
    }

    onShowToast?.(
      importMode === 'overwrite'
        ? `បានជំនួស និងនាំចូលទិន្នន័យស្តុកចំនួន ${parsedImportRows.length} ជួរចូលក្នុងស្តុកជាក់ស្តែងដោយជោគជ័យ!`
        : `បានបន្ថែមទិន្នន័យស្តុកចំនួន ${parsedImportRows.length} ជួរចូលក្នុងស្តុកជាក់ស្តែងដោយជោគជ័យ!`,
      'success'
    );
    setIsImportModalOpen(false);
    setParsedImportRows([]);
    setImportFile(null);
  };

  // Export Table to Excel
  const handleExportExcel = () => {
    const exportData = filteredStatusItems.map((it, idx) => ({
      'ល.រ': idx + 1,
      'ប្រភេទទិដ្ឋាការ': it.visaType,
      'ស្តុកដើម': it.beginningStock,
      'បើកពី ក១': it.receivedK1,
      'បង្វិលពីក្រុម': it.returnedFromTeam,
      'បើកជូនក្រុម': it.issuedToTeam,
      'ខូច/សាក K2': it.damagedK2 + it.testPrintK2,
      'ស្តុកជាក់ស្តែងនៅសល់ (សន្លឹក)': it.officeRemainingSheets,
      'គិតជាក្បាល (៥០សន្លឹក/ក្បាល)': `${it.officeRemainingBooks} ក្បាល ${it.officeRemainingLoose > 0 ? `+ ${it.officeRemainingLoose} សន្លឹក` : ''}`,
      'ចន្លោះលេខស៊េរីនៅសល់': it.serialRangeDisplay,
      'ស្ថានភាព':
        it.status === 'sufficient'
          ? 'គ្រប់គ្រាន់'
          : it.status === 'medium'
          ? 'មធ្យម'
          : it.status === 'low'
          ? 'ជិតអស់'
          : 'អស់ពីស្តុក',
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportData);
    worksheet['!cols'] = [
      { wch: 6 },
      { wch: 15 },
      { wch: 12 },
      { wch: 12 },
      { wch: 14 },
      { wch: 14 },
      { wch: 12 },
      { wch: 22 },
      { wch: 22 },
      { wch: 32 },
      { wch: 14 },
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'ស្តុកជាក់ស្តែង');
    XLSX.writeFile(workbook, `ស្តុកជាក់ស្តែងសម្រាប់បើកផ្តល់_${asOfDate}.xlsx`);
    onShowToast?.('បានទាញយកតារាងស្តុកជាក់ស្តែងជា Excel រួចរាល់!');
  };

  // Print Table
  const handlePrint = () => {
    printA4Document('sticker-status-print-document', {
      orientation: 'portrait',
      documentTitle: `តារាងស្តុកជាក់ស្តែង_${asOfDate}`,
    });
  };

  return (
    <div id="sticker-status-container" className="space-y-6 animate-fade font-sans text-slate-800">
      {/* Header Bar */}
      <div className="bg-white rounded-2xl border border-slate-200/80 p-5 shadow-xs flex flex-col md:flex-row md:items-center justify-between gap-4">
        <div>
          <div className="inline-flex items-center gap-2 text-xs font-bold text-blue-800 bg-blue-50/90 px-3 py-1 rounded-full mb-2 border border-blue-100">
            <Boxes className="w-3.5 h-3.5 text-blue-700" />
            <span>ការគ្រប់គ្រងស្តុកសន្លឹកទិដ្ឋាការស្អិតជាក់ស្តែង</span>
          </div>
          <h1 className="text-xl font-bold text-slate-900 flex items-center gap-2.5">
            <span>ស្តុកជាក់ស្តែងសម្រាប់ទុកធ្វើការបើកផ្តល់តាមក្រុម</span>
          </h1>
          <p className="text-xs text-slate-500 mt-1">
            តាមដាន និងទូទាត់ចន្លោះលេខស៊េរីសន្លឹកទិដ្ឋាការនៅសល់ជាក់ស្តែងក្នុងស្តុក K2 សម្រាប់បើកផ្តល់ជូនបណ្តាក្រុម/ច្រកទ្វារ
          </p>
        </div>

        {/* Action Buttons */}
        <div className="flex flex-wrap items-center gap-2.5">
          <button
            onClick={() => setIsAddK1ModalOpen(true)}
            className="px-4 py-2 bg-indigo-600 hover:bg-indigo-700 active:scale-98 text-white rounded-xl text-xs font-bold transition-all shadow-xs flex items-center gap-2 cursor-pointer"
          >
            <PlusCircle className="w-4 h-4" />
            <span>បញ្ចូលស្តុកពី ក១</span>
          </button>

          <button
            onClick={() => handleOpenIssueModal()}
            className="px-4 py-2 bg-emerald-600 hover:bg-emerald-700 active:scale-98 text-white rounded-xl text-xs font-bold transition-all shadow-xs flex items-center gap-2 cursor-pointer"
          >
            <Send className="w-4 h-4" />
            <span>បើកផ្តល់ជូនក្រុម</span>
          </button>

          {issuedRecords.length > 0 && (
            <button
              onClick={() => setIsReturnModalOpen(true)}
              title="បង្វិលទិន្នន័យដែលបានបើកផ្តល់សាកល្បងត្រឡប់ចូលស្តុកជាក់ស្តែង និងកែប្រែទិន្នន័យ (cEA) ឡើងវិញ"
              className="px-4 py-2 bg-amber-500 hover:bg-amber-600 active:scale-98 text-white rounded-xl text-xs font-bold transition-all shadow-xs flex items-center gap-2 cursor-pointer"
            >
              <RotateCcw className="w-4 h-4" />
              <span>បង្វិលទិន្នន័យបើកផ្តល់ ({issuedRecords.length})</span>
            </button>
          )}

          <button
            onClick={() => setIsImportModalOpen(true)}
            className="px-4 py-2 bg-blue-600 hover:bg-blue-700 active:scale-98 text-white rounded-xl text-xs font-bold transition-all shadow-xs flex items-center gap-2 cursor-pointer"
          >
            <Upload className="w-4 h-4" />
            <span>Import Excel ស្តុក</span>
          </button>

          {onClearAllActualStock && (
            <button
              onClick={() => setIsClearAllConfirmOpen(true)}
              title="លុបទិន្នន័យទាំងអស់ក្នុងតារាងស្តុកជាក់ស្តែងនេះ (ដើម្បី Import ឯកសារថ្មី)"
              className="px-3.5 py-2 bg-rose-600 hover:bg-rose-700 active:scale-98 text-white rounded-xl text-xs font-bold transition-all shadow-xs flex items-center gap-1.5 cursor-pointer"
            >
              <Trash2 className="w-4 h-4" />
              <span>លុបតារាងទាំងអស់</span>
            </button>
          )}

          <button
            onClick={() => setIsManageDeleteModalOpen(true)}
            title="ជ្រើសរើស និងលុបទិន្នន័យចេញពីស្តុកជាក់ស្តែង (មិនលុបក្នុងទិន្នន័យសន្លឹកទិដ្ឋាការ)"
            className="px-3.5 py-2 bg-rose-50 text-rose-700 border border-rose-200 hover:bg-rose-100 active:scale-98 rounded-xl text-xs font-bold transition-all shadow-2xs flex items-center gap-1.5 cursor-pointer"
          >
            <Trash2 className="w-4 h-4 text-rose-600" />
            <span>គ្រប់គ្រងលុបតាមជួរ</span>
          </button>

          <button
            onClick={handleExportExcel}
            className="px-3.5 py-2 bg-emerald-50 text-emerald-800 border border-emerald-300 hover:bg-emerald-100 rounded-xl text-xs font-bold transition-all flex items-center gap-1.5 cursor-pointer"
          >
            <FileSpreadsheet className="w-4 h-4 text-emerald-700" />
            <span>Export Excel</span>
          </button>

          {onClose && (
            <button
              onClick={onClose}
              className="p-2 text-slate-400 hover:text-slate-700 hover:bg-slate-100 rounded-xl transition"
            >
              <X className="w-5 h-5" />
            </button>
          )}
        </div>
      </div>

      {/* Empty Stock State Notice if no sticker records exist */}
      {stockRecords.filter((r) => r.stockType === 'sticker').length === 0 && (
        <div className="p-4 rounded-2xl bg-blue-50/80 border border-blue-200 text-blue-950 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 shadow-xs">
          <div className="flex items-center gap-3">
            <div className="p-2.5 bg-blue-100 text-blue-800 rounded-xl shrink-0">
              <Info className="w-5 h-5" />
            </div>
            <div>
              <h4 className="text-xs font-bold">ទិន្នន័យស្តុកជាក់ស្តែងទំនេរ (គ្មានទិន្នន័យ)</h4>
              <p className="text-xs text-blue-800 mt-0.5">
                ទិន្នន័យស្តុកត្រូវបានសម្អាតរួចរាល់។ លោកអ្នកអាចចុច «Import Excel ស្តុក» ដើម្បីនាំចូលទិន្នន័យជាក់ស្តែងរបស់អ្នកបានយ៉ាងងាយស្រួល។
              </p>
            </div>
          </div>
          <button
            onClick={() => setIsImportModalOpen(true)}
            className="px-4 py-2 bg-blue-600 hover:bg-blue-700 text-white rounded-xl text-xs font-bold transition flex items-center gap-2 cursor-pointer shadow-xs shrink-0"
          >
            <Upload className="w-4 h-4" />
            <span>Import Excel ស្តុកឥឡូវនេះ</span>
          </button>
        </div>
      )}

      {/* KPI Cards Overview */}
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-4">
        {/* Available Sheets in Office K2 */}
        <div className="bg-white rounded-2xl border border-blue-200/80 p-4 shadow-xs relative overflow-hidden flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-blue-950">ស្តុកជាក់ស្តែងសរុប (K2)</span>
              <div className="p-2 rounded-xl bg-blue-50 text-blue-700">
                <Package className="w-5 h-5" />
              </div>
            </div>
            <div className="mt-2 flex items-baseline gap-2">
              <span className="text-2xl font-black text-slate-900 tabular-nums">
                {summary.totalOfficeSheets.toLocaleString()}
              </span>
              <span className="text-xs font-bold text-slate-500">សន្លឹក</span>
            </div>
          </div>
          <div className="mt-2 text-xs text-blue-800 font-semibold pt-2 border-t border-blue-50">
            ស្មើនឹង <b className="tabular-nums">{summary.totalOfficeBooks.toLocaleString()}</b> ក្បាល{' '}
            {summary.totalOfficeLoose > 0 ? `+ ${summary.totalOfficeLoose} សន្លឹក` : ''}
          </div>
          <div className="absolute top-0 right-0 w-24 h-24 bg-blue-500/5 rounded-full -mr-6 -mt-6 pointer-events-none" />
        </div>

        {/* Total Available Books */}
        <div className="bg-white rounded-2xl border border-indigo-200/80 p-4 shadow-xs relative overflow-hidden flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-indigo-950">ចំនួនក្បាលជាក់ស្តែង (៥០សន្លឹក/ក្បាល)</span>
              <div className="p-2 rounded-xl bg-indigo-50 text-indigo-700">
                <BookOpen className="w-5 h-5" />
              </div>
            </div>
            <div className="mt-2 flex items-baseline gap-2">
              <span className="text-2xl font-black text-slate-900 tabular-nums">
                {summary.totalAllBooksCount.toLocaleString()}
              </span>
              <span className="text-xs font-bold text-slate-500">ក្បាល</span>
            </div>
          </div>
          <div className="mt-2 text-xs text-indigo-800 font-semibold pt-2 border-t border-indigo-50">
            បែងចែកលេខស៊េរីតាមក្បាលជាក់ស្តែង
          </div>
          <div className="absolute top-0 right-0 w-24 h-24 bg-indigo-500/5 rounded-full -mr-6 -mt-6 pointer-events-none" />
        </div>

        {/* Issued to Teams */}
        <div className="bg-white rounded-2xl border border-amber-200/80 p-4 shadow-xs relative overflow-hidden flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-amber-950">បានបើកជូនក្រុមសរុប</span>
              <div className="p-2 rounded-xl bg-amber-50 text-amber-700">
                <Send className="w-5 h-5" />
              </div>
            </div>
            <div className="mt-2 flex items-baseline gap-2">
              <span className="text-2xl font-black text-slate-900 tabular-nums">
                {summary.totalIssuedSheets.toLocaleString()}
              </span>
              <span className="text-xs font-bold text-slate-500">សន្លឹក</span>
            </div>
          </div>
          <div className="mt-2 text-xs text-amber-800 font-semibold pt-2 border-t border-amber-50">
            បានកាត់ទូទាត់ចេញពីស្តុកជាក់ស្តែងភ្លាមៗ
          </div>
          <div className="absolute top-0 right-0 w-24 h-24 bg-amber-500/5 rounded-full -mr-6 -mt-6 pointer-events-none" />
        </div>

        {/* Stock Status Summary */}
        <div className="bg-white rounded-2xl border border-emerald-200/80 p-4 shadow-xs relative overflow-hidden flex flex-col justify-between">
          <div>
            <div className="flex items-center justify-between">
              <span className="text-xs font-bold text-emerald-950">ស្ថានភាពគ្រប់គ្រងស្តុក</span>
              <div className="p-2 rounded-xl bg-emerald-50 text-emerald-700">
                <CheckCircle2 className="w-5 h-5" />
              </div>
            </div>
            <div className="mt-2 flex items-center gap-2 flex-wrap">
              <span className="text-xs px-2.5 py-0.5 rounded-lg bg-emerald-100 text-emerald-800 font-bold">
                គ្រប់គ្រាន់: {summary.countSufficient}
              </span>
              <span className="text-xs px-2.5 py-0.5 rounded-lg bg-amber-100 text-amber-800 font-bold">
                មធ្យម/ខ្វះ: {summary.countMedium + summary.countLow}
              </span>
              {summary.countOut > 0 && (
                <span className="text-xs px-2.5 py-0.5 rounded-lg bg-rose-100 text-rose-800 font-bold">
                  អស់: {summary.countOut}
                </span>
              )}
            </div>
          </div>
          <div className="mt-2 text-xs text-emerald-800 font-semibold pt-2 border-t border-emerald-50">
            ទូទាត់លេខស៊េរីច្បាស់លាស់
          </div>
          <div className="absolute top-0 right-0 w-24 h-24 bg-emerald-500/5 rounded-full -mr-6 -mt-6 pointer-events-none" />
        </div>
      </div>

      {/* Tabs Navigation & Filter Controls */}
      <div className="bg-white rounded-2xl border border-slate-200/80 p-4 shadow-xs space-y-4">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-4 border-b border-slate-100 pb-4">
          {/* Main View Tabs */}
          <div className="flex flex-wrap items-center gap-1.5 p-1 bg-slate-100 rounded-xl w-fit">
            <button
              onClick={() => setActiveTab('rangeTable')}
              className={`px-3.5 py-2 rounded-lg text-xs font-bold transition cursor-pointer flex items-center gap-2 ${
                activeTab === 'rangeTable'
                  ? 'bg-white text-blue-900 shadow-xs border border-slate-200/60'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <FileSpreadsheet className="w-4 h-4 text-emerald-600" />
              <span>១. តារាងតាមចន្លោះស៊េរី (ដូចគំរូរូបភាព)</span>
            </button>

            <button
              onClick={() => setActiveTab('summary')}
              className={`px-3.5 py-2 rounded-lg text-xs font-bold transition cursor-pointer flex items-center gap-2 ${
                activeTab === 'summary'
                  ? 'bg-white text-blue-900 shadow-xs border border-slate-200/60'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <Boxes className="w-4 h-4 text-blue-600" />
              <span>២. តារាងសង្ខេបទូទៅ</span>
            </button>

            <button
              onClick={() => setActiveTab('booksTable')}
              className={`px-3.5 py-2 rounded-lg text-xs font-bold transition cursor-pointer flex items-center gap-2 ${
                activeTab === 'booksTable'
                  ? 'bg-white text-blue-900 shadow-xs border border-slate-200/60'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <BookOpen className="w-4 h-4 text-indigo-600" />
              <span>៣. បញ្ជីក្បាលទិដ្ឋាការនៅសល់ ({filteredBooks.length} ក្បាល)</span>
            </button>

            <button
              onClick={() => setActiveTab('transactions')}
              className={`px-3.5 py-2 rounded-lg text-xs font-bold transition cursor-pointer flex items-center gap-2 ${
                activeTab === 'transactions'
                  ? 'bg-white text-blue-900 shadow-xs border border-slate-200/60'
                  : 'text-slate-600 hover:text-slate-900'
              }`}
            >
              <Layers className="w-4 h-4 text-amber-600" />
              <span>៤. ប្រវត្តិប្រតិបត្តិការស្តុក</span>
            </button>
          </div>

          {/* Scope Selector */}
          <div className="flex items-center gap-2">
            <span className="text-xs font-bold text-slate-500">មើលស្តុក៖</span>
            <div className="flex items-center bg-slate-100 p-1 rounded-xl">
              <button
                onClick={() => setSelectedStockScope('office')}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition cursor-pointer flex items-center gap-1.5 ${
                  selectedStockScope === 'office'
                    ? 'bg-white text-blue-900 shadow-xs border border-slate-200/60'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <Building className="w-3.5 h-3.5 text-blue-600" />
                <span>ស្តុក K2</span>
              </button>
              <button
                onClick={() => setSelectedStockScope('teams')}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition cursor-pointer flex items-center gap-1.5 ${
                  selectedStockScope === 'teams'
                    ? 'bg-white text-amber-900 shadow-xs border border-slate-200/60'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <Users className="w-3.5 h-3.5 text-amber-600" />
                <span>ស្តុកតាមក្រុម</span>
              </button>
              <button
                onClick={() => setSelectedStockScope('total')}
                className={`px-3 py-1.5 rounded-lg text-xs font-bold transition cursor-pointer flex items-center gap-1.5 ${
                  selectedStockScope === 'total'
                    ? 'bg-white text-emerald-900 shadow-xs border border-slate-200/60'
                    : 'text-slate-600 hover:text-slate-900'
                }`}
              >
                <Layers className="w-3.5 h-3.5 text-emerald-600" />
                <span>ស្តុកសរុប</span>
              </button>
            </div>
          </div>
        </div>

        {/* Filter Controls Row */}
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 pt-1">
          {/* Search Box */}
          <div className="relative">
            <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              placeholder="ស្វែងរកប្រភេទទិដ្ឋាការ, លេខស៊េរី, ក្បាល..."
              value={searchQuery}
              onChange={(e) => {
                setSearchQuery(e.target.value);
                setCurrentPage(1);
              }}
              className="w-full pl-9 pr-3 py-2 text-xs rounded-xl border border-slate-300 focus:outline-hidden focus:ring-2 focus:ring-blue-500 bg-white font-medium"
            />
          </div>

          {/* Visa Type Filter */}
          <div className="flex items-center gap-2">
            <label className="text-xs font-bold text-slate-600 shrink-0">ប្រភេទទិដ្ឋាការ:</label>
            <select
              value={selectedVisaTypeFilter}
              onChange={(e) => {
                setSelectedVisaTypeFilter(e.target.value);
                setCurrentPage(1);
              }}
              className="w-full px-3 py-2 text-xs rounded-xl border border-slate-300 focus:outline-hidden focus:ring-2 focus:ring-blue-500 bg-white font-medium"
            >
              <option value="ALL">-- គ្រប់ប្រភេទទិដ្ឋាការទាំងអស់ --</option>
              {allVisaTypes.map((vt) => (
                <option key={vt} value={vt}>
                  ទិដ្ឋាការប្រភេទ {vt}
                </option>
              ))}
            </select>
          </div>

          {/* As Of Date & Filter Toggle */}
          <div className="flex flex-col sm:flex-row sm:items-center gap-2">
            <div className="flex items-center gap-1.5 shrink-0">
              <input
                id="filter-by-date-checkbox"
                type="checkbox"
                checked={filterByDate}
                onChange={(e) => setFilterByDate(e.target.checked)}
                className="w-4 h-4 rounded text-blue-600 focus:ring-blue-500 cursor-pointer"
              />
              <label htmlFor="filter-by-date-checkbox" className="text-xs font-bold text-slate-700 cursor-pointer select-none">
                គិតត្រឹមថ្ងៃ:
              </label>
            </div>
            <div className="w-36">
              <CustomDatePicker
                value={asOfDate}
                disabled={!filterByDate}
                onChange={(d) => setAsOfDate(d)}
                className={`py-1.5 text-xs ${!filterByDate ? 'opacity-60 bg-slate-100 cursor-not-allowed' : ''}`}
              />
            </div>
            {!filterByDate && (
              <span className="text-[11px] text-emerald-700 font-semibold bg-emerald-50 px-2.5 py-1 rounded-lg border border-emerald-200 shrink-0">
                បង្ហាញស្តុកជាក់ស្តែងទាំងអស់
              </span>
            )}
          </div>
        </div>
      </div>

      {/* ========================================================================= */}
      {/* TAB 1: RANGE-BASED STOCK TABLE (EXACT MATCH TO USER IMAGE) */}
      {/* ========================================================================= */}
      {activeTab === 'rangeTable' && (
        <div className="bg-white rounded-2xl border border-slate-200/80 overflow-hidden shadow-xs">
          <div className="p-4 bg-slate-900 text-white flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div className="flex items-center gap-2.5">
              <FileSpreadsheet className="w-5 h-5 text-emerald-400" />
              <div>
                <h2 className="text-sm font-bold">តារាងចន្លោះលេខស៊េរីស្តុកជាក់ស្តែង (ទម្រង់ដូចគំរូរូបភាព)</h2>
                <p className="text-[11px] text-slate-300">
                  បង្ហាញទិន្នន័យច្បាស់លាស់តាមជួរឈរ៖ ប្រភេទ | ក្បាល | សន្លឹក | ចាប់ផ្តើម | ដល់លេខ
                </p>
              </div>
            </div>

            <div className="flex items-center gap-2">
              {onClearAllActualStock && (
                <button
                  onClick={() => setIsClearAllConfirmOpen(true)}
                  title="លុបទិន្នន័យតារាងនេះទាំងអស់ ដើម្បី Import ឡើងវិញ"
                  className="px-3 py-1.5 bg-rose-600/90 hover:bg-rose-600 text-white rounded-xl text-xs font-bold transition flex items-center gap-1.5 cursor-pointer shadow-xs"
                >
                  <Trash2 className="w-3.5 h-3.5" />
                  <span>លុបទិន្នន័យតារាងទាំងអស់</span>
                </button>
              )}
              <button
                onClick={handleExportRangeFormatExcel}
                className="px-3 py-1.5 bg-white/10 hover:bg-white/20 text-white rounded-xl text-xs font-bold transition flex items-center gap-1.5 border border-white/10 cursor-pointer"
              >
                <Download className="w-3.5 h-3.5 text-emerald-400" />
                <span>ទាញយក Excel ដូចគំរូ</span>
              </button>
            </div>
          </div>

          <div className="p-6 overflow-x-auto bg-slate-50/30 flex justify-center">
            <div className="w-full max-w-4xl bg-white p-4 sm:p-6 rounded-xl border border-slate-200 shadow-xs">
              <div className="text-center mb-4 pb-3 border-b border-slate-100">
                <h3 className="font-moul text-sm text-slate-900">តារាងស្តុកជាក់ស្តែងសម្រាប់ទុកធ្វើការបើកផ្តល់តាមក្រុម</h3>
                <p className="text-xs text-slate-500 mt-1 font-siemreap">
                  គិតត្រឹមថ្ងៃទី {asOfDate} ({selectedStockScope === 'office' ? 'ស្តុក K2' : selectedStockScope === 'teams' ? 'ស្តុកតាមក្រុម' : 'ស្តុកសរុប'})
                </p>
              </div>

              <table className="w-full border-collapse border border-slate-800 text-xs sm:text-sm font-siemreap">
                <thead>
                  <tr className="bg-slate-100 text-slate-900 font-bold border-b border-slate-800">
                    <th className="border border-slate-800 py-3 px-3 text-center w-24 sm:w-28 font-bold">
                      ប្រភេទ
                    </th>
                    <th className="border border-slate-800 py-3 px-3 text-center w-24 sm:w-28 font-bold">
                      ក្បាល
                    </th>
                    <th className="border border-slate-800 py-3 px-3 text-center w-28 sm:w-32 font-bold">
                      សន្លឹក
                    </th>
                    <th className="border border-slate-800 py-3 px-4 text-center font-bold">
                      ចាប់ផ្តើម
                    </th>
                    <th className="border border-slate-800 py-3 px-4 text-center font-bold">
                      ដល់លេខ
                    </th>
                    <th className="border border-slate-800 py-3 px-2 text-center w-16 font-bold">
                      សកម្មភាព
                    </th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-800">
                  {filteredStatusItems.map((item) => {
                    const ranges = item.remainingRanges;
                    const remainingSheets =
                      selectedStockScope === 'office'
                        ? item.officeRemainingSheets
                        : selectedStockScope === 'teams'
                        ? item.teamRemainingSheets
                        : item.totalRemainingSheets;
                    const remainingBooks =
                      selectedStockScope === 'office'
                        ? item.officeRemainingBooks
                        : selectedStockScope === 'teams'
                        ? item.teamRemainingBooks
                        : item.totalRemainingBooks;

                    if (ranges && ranges.length > 0) {
                      return ranges.map((range, rangeIdx) => {
                        const rangeSheets = range.count;
                        const rangeBooks = Math.floor(rangeSheets / 50);

                        return (
                          <tr
                            key={`${item.visaType}-${rangeIdx}`}
                            className="hover:bg-blue-50/30 transition text-slate-900"
                          >
                            {rangeIdx === 0 && (
                              <td
                                rowSpan={ranges.length}
                                className="border border-slate-800 py-3 px-3 text-center font-black text-base align-middle bg-slate-50/50"
                              >
                                <span className="inline-block font-extrabold text-base sm:text-lg text-slate-950">
                                  {item.visaType}
                                </span>
                              </td>
                            )}
                            <td className="border border-slate-800 py-2.5 px-3 text-center font-semibold text-sm sm:text-base tabular-nums">
                              {rangeBooks.toLocaleString()}
                            </td>
                            <td className="border border-slate-800 py-2.5 px-3 text-center font-semibold text-sm sm:text-base tabular-nums">
                              {rangeSheets.toLocaleString()}
                            </td>
                            <td className="border border-slate-800 py-2.5 px-4 text-center font-medium text-xs sm:text-sm tracking-wide tabular-nums">
                              {range.startSerial}
                            </td>
                            <td className="border border-slate-800 py-2.5 px-4 text-center font-medium text-xs sm:text-sm tracking-wide tabular-nums">
                              {range.endSerial}
                            </td>
                            <td className="border border-slate-800 py-2 px-2 text-center align-middle">
                              <button
                                type="button"
                                onClick={() =>
                                  setDeletingRange({
                                    visaType: item.visaType,
                                    startSerial: range.startSerial,
                                    endSerial: range.endSerial,
                                    count: rangeSheets,
                                  })
                                }
                                title={`លុបចន្លោះ ${range.startSerial} - ${range.endSerial} ពីស្តុកជាក់ស្តែង (មិនលុបក្នុងទិន្នន័យសន្លឹកទិដ្ឋាការ)`}
                                className="p-1 text-rose-600 hover:text-white hover:bg-rose-600 rounded-lg transition-colors cursor-pointer inline-flex items-center justify-center shadow-2xs"
                              >
                                <Trash2 className="w-4 h-4" />
                              </button>
                            </td>
                          </tr>
                        );
                      });
                    }

                    // Fallback row if no specific interval calculated
                    return (
                      <tr key={item.visaType} className="hover:bg-blue-50/30 transition text-slate-900">
                        <td className="border border-slate-800 py-3 px-3 text-center font-black text-base align-middle bg-slate-50/50">
                          <span className="inline-block font-extrabold text-base sm:text-lg text-slate-950">
                            {item.visaType}
                          </span>
                        </td>
                        <td className="border border-slate-800 py-2.5 px-3 text-center font-semibold text-sm sm:text-base tabular-nums">
                          {remainingBooks.toLocaleString()}
                        </td>
                        <td className="border border-slate-800 py-2.5 px-3 text-center font-semibold text-sm sm:text-base tabular-nums">
                          {remainingSheets.toLocaleString()}
                        </td>
                        <td className="border border-slate-800 py-2.5 px-4 text-center font-medium text-xs sm:text-sm tracking-wide tabular-nums">
                          {item.startSerial !== '-' ? item.startSerial : '-'}
                        </td>
                        <td className="border border-slate-800 py-2.5 px-4 text-center font-medium text-xs sm:text-sm tracking-wide tabular-nums">
                          {item.endSerial !== '-' ? item.endSerial : '-'}
                        </td>
                        <td className="border border-slate-800 py-2 px-2 text-center align-middle">
                          <button
                            type="button"
                            onClick={() => {
                              const cleanVt = item.visaType.trim().toUpperCase();
                              const recs = stockRecords.filter((r) => (r.visaType || '').trim().toUpperCase() === cleanVt);
                              if (recs.length === 1) {
                                setDeletingRecord(recs[0]);
                              } else if (recs.length > 1) {
                                setSelectedManageIds(recs.map((r) => r.id));
                                setIsBatchDeleteConfirmOpen(true);
                              }
                            }}
                            title={`លុបទិន្នន័យប្រភេទ ${item.visaType} ពីស្តុកជាក់ស្តែង (មិនលុបក្នុងទិន្នន័យសន្លឹកទិដ្ឋាការ)`}
                            className="p-1 text-rose-600 hover:text-white hover:bg-rose-600 rounded-lg transition-colors cursor-pointer inline-flex items-center justify-center shadow-2xs"
                          >
                            <Trash2 className="w-4 h-4" />
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
                <tfoot>
                  <tr className="bg-slate-100 font-bold text-slate-950 border-t-2 border-slate-800">
                    <td className="border border-slate-800 py-3 px-3 text-center font-black text-sm sm:text-base">
                      សរុបរួម
                    </td>
                    <td className="border border-slate-800 py-3 px-3 text-center font-black text-sm sm:text-base tabular-nums">
                      {summary.totalBooks.toLocaleString()}
                    </td>
                    <td className="border border-slate-800 py-3 px-3 text-center font-black text-sm sm:text-base tabular-nums">
                      {summary.totalSheets.toLocaleString()}
                    </td>
                    <td colSpan={3} className="border border-slate-800 py-3 px-4 text-center text-slate-600 font-normal text-xs">
                      សរុបគ្រប់ប្រភេទ និងគ្រប់ចន្លោះលេខ
                    </td>
                  </tr>
                </tfoot>
              </table>

              {summary.totalSheets === 0 && (
                <div className="mt-6 p-6 rounded-2xl bg-blue-50/70 border-2 border-dashed border-blue-200 text-center space-y-3">
                  <div className="w-12 h-12 mx-auto rounded-full bg-blue-100 text-blue-700 flex items-center justify-center">
                    <Upload className="w-6 h-6" />
                  </div>
                  <div>
                    <h4 className="text-sm font-bold text-slate-800">តារាងស្តុកជាក់ស្តែងបានសម្អាតរួចរាល់ (០ សន្លឹក)</h4>
                    <p className="text-xs text-slate-500 mt-0.5">លោកអ្នកអាចធ្វើការ Import ឯកសារ Excel ស្តុកថ្មីដែលត្រឹមត្រូវចូលក្នុងតារាងនេះបាន</p>
                  </div>
                  <button
                    onClick={() => setIsImportModalOpen(true)}
                    className="inline-flex items-center gap-2 px-4 py-2 bg-blue-600 hover:bg-blue-700 active:scale-98 text-white rounded-xl text-xs font-bold transition shadow-xs cursor-pointer"
                  >
                    <Upload className="w-4 h-4" />
                    <span>Import Excel ស្តុកថ្មី</span>
                  </button>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* TAB 2: SUMMARY TABLE BY VISA TYPE */}
      {/* ========================================================================= */}
      {activeTab === 'summary' && (
        <div className="bg-white rounded-2xl border border-slate-200/80 overflow-hidden shadow-xs">
          <div className="p-4 bg-slate-900 text-white flex items-center justify-between">
            <div className="flex items-center gap-2.5">
              <Boxes className="w-5 h-5 text-blue-300" />
              <div>
                <h2 className="text-sm font-bold">តារាងសង្ខេបស្តុកជាក់ស្តែងតាមប្រភេទទិដ្ឋាការ</h2>
                <p className="text-[11px] text-slate-300">
                  គណនាកាត់ចេញពីស្តុកជាក់ស្តែងភ្លាមៗពេលបើកជូនក្រុម និងបង្ហាញចន្លោះស៊េរីនៅសល់
                </p>
              </div>
            </div>

            <div className="text-xs bg-white/10 px-3 py-1.5 rounded-lg border border-white/10 font-bold tabular-nums">
              សរុប {summary.totalSheets.toLocaleString()} សន្លឹក ({summary.totalBooks.toLocaleString()} ក្បាល)
            </div>
          </div>

          <div className="overflow-x-auto">
            <table className="w-full text-xs text-left">
              <thead>
                <tr className="bg-slate-100 text-slate-700 font-bold border-b border-slate-200">
                  <th className="py-3 px-3 text-center w-12">ល.រ</th>
                  <th className="py-3 px-3 text-center w-28">ប្រភេទទិដ្ឋាការ</th>
                  <th className="py-3 px-3 text-right">ស្តុកដើម/ក១</th>
                  <th className="py-3 px-3 text-right">បង្វិលចូល</th>
                  <th className="py-3 px-3 text-right text-amber-800">បានបើកជូនក្រុម</th>
                  <th className="py-3 px-3 text-right text-rose-700">ខូច/សាក K2</th>
                  <th className="py-3 px-3 text-right bg-blue-50/60 text-blue-950 font-black">
                    ស្តុកជាក់ស្តែងនៅសល់
                  </th>
                  <th className="py-3 px-3 text-center bg-indigo-50/60 text-indigo-950 font-black">
                    គិតជាក្បាល (៥០សន្លឹក)
                  </th>
                  <th className="py-3 px-3 text-center">ចន្លោះលេខស៊េរីនៅសល់</th>
                  <th className="py-3 px-3 text-center w-24">ស្ថានភាព</th>
                  <th className="py-3 px-3 text-center w-36">សកម្មភាព</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200">
                {filteredStatusItems.map((item, idx) => {
                  const remainingSheets =
                    selectedStockScope === 'office'
                      ? item.officeRemainingSheets
                      : selectedStockScope === 'teams'
                      ? item.teamRemainingSheets
                      : item.totalRemainingSheets;

                  const remainingBooks =
                    selectedStockScope === 'office'
                      ? item.officeRemainingBooks
                      : selectedStockScope === 'teams'
                      ? item.teamRemainingBooks
                      : item.totalRemainingBooks;

                  const remainingLoose =
                    selectedStockScope === 'office'
                      ? item.officeRemainingLoose
                      : selectedStockScope === 'teams'
                      ? item.teamRemainingLoose
                      : item.totalRemainingLoose;

                  return (
                    <tr
                      key={item.visaType}
                      className="hover:bg-blue-50/40 transition group"
                    >
                      <td className="py-3 px-3 text-center text-slate-500 font-bold bg-slate-50/50 tabular-nums">
                        {idx + 1}
                      </td>
                      <td className="py-3 px-3 text-center">
                        <span className="inline-block px-2.5 py-1 rounded-md bg-blue-100 text-blue-900 font-black text-xs">
                          {item.visaType}
                        </span>
                      </td>
                      <td className="py-3 px-3 text-right font-semibold text-slate-700 tabular-nums">
                        {(item.beginningStock + item.receivedK1).toLocaleString()}
                      </td>
                      <td className="py-3 px-3 text-right font-semibold text-emerald-700 tabular-nums">
                        {item.returnedFromTeam > 0 ? `+${item.returnedFromTeam.toLocaleString()}` : '-'}
                      </td>
                      <td className="py-3 px-3 text-right font-bold text-amber-800 bg-amber-50/30 tabular-nums">
                        {item.issuedToTeam > 0 ? (
                          <div className="flex items-center justify-end gap-1.5">
                            <span>-{item.issuedToTeam.toLocaleString()}</span>
                            <button
                              onClick={() => setReturnTypeConfirm(item.visaType)}
                              title={`បង្វិលទិន្នន័យដែលបានបើកផ្តល់ប្រភេទ ${item.visaType} (សរុប ${item.issuedToTeam} សន្លឹក) ត្រឡប់ចូលស្តុកជាក់ស្តែង និងទិន្នន័យ (cEA) វិញ`}
                              className="px-1.5 py-0.5 rounded bg-amber-500 hover:bg-amber-600 text-white font-bold text-[10px] transition flex items-center gap-0.5 cursor-pointer shadow-2xs"
                            >
                              <RotateCcw className="w-2.5 h-2.5" />
                              <span>បង្វិល</span>
                            </button>
                          </div>
                        ) : (
                          '0'
                        )}
                      </td>
                      <td className="py-3 px-3 text-right font-semibold text-rose-700 tabular-nums">
                        {item.damagedK2 + item.testPrintK2 > 0
                          ? `-${(item.damagedK2 + item.testPrintK2).toLocaleString()}`
                          : '-'}
                      </td>
                      <td className="py-3 px-3 text-right font-black text-sm text-blue-950 bg-blue-50/60 tabular-nums">
                        {remainingSheets.toLocaleString()}{' '}
                        <span className="text-[10px] font-normal text-blue-700">សន្លឹក</span>
                      </td>
                      <td className="py-3 px-3 text-center font-bold text-indigo-950 bg-indigo-50/60">
                        {remainingBooks > 0 || remainingLoose > 0 ? (
                          <span>
                            <b className="tabular-nums">{remainingBooks}</b> ក្បាល {remainingLoose > 0 ? `+ ${remainingLoose} សន្លឹក` : ''}
                          </span>
                        ) : (
                          <span className="text-slate-400">0</span>
                        )}
                      </td>
                      <td className="py-3 px-3 text-center font-medium text-slate-800 text-[11px] tabular-nums">
                        {item.serialRangeDisplay}
                      </td>
                      <td className="py-3 px-3 text-center">
                        {item.status === 'sufficient' && (
                          <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full bg-emerald-100 text-emerald-800 text-[10px] font-bold">
                            <Check className="w-3 h-3" />
                            <span>គ្រប់គ្រាន់</span>
                          </span>
                        )}
                        {item.status === 'medium' && (
                          <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full bg-blue-100 text-blue-800 text-[10px] font-bold">
                            <span>មធ្យម</span>
                          </span>
                        )}
                        {item.status === 'low' && (
                          <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full bg-amber-100 text-amber-800 text-[10px] font-bold">
                            <AlertTriangle className="w-3 h-3" />
                            <span>ជិតអស់</span>
                          </span>
                        )}
                        {item.status === 'out' && (
                          <span className="inline-flex items-center gap-1 px-2.5 py-0.5 rounded-full bg-rose-100 text-rose-800 text-[10px] font-bold">
                            <span>អស់ពីស្តុក</span>
                          </span>
                        )}
                      </td>
                      <td className="py-3 px-3 text-center">
                        <div className="flex items-center justify-center gap-1.5">
                          <button
                            onClick={() => handleOpenIssueModal(item.visaType)}
                            disabled={remainingSheets <= 0}
                            title="បើកផ្តល់ជូនក្រុម"
                            className="px-2.5 py-1 bg-emerald-600 hover:bg-emerald-700 disabled:opacity-40 disabled:cursor-not-allowed text-white rounded-lg text-[11px] font-bold transition flex items-center gap-1 cursor-pointer"
                          >
                            <Send className="w-3 h-3" />
                            <span>បើក</span>
                          </button>

                          {item.issuedToTeam > 0 && (
                            <button
                              onClick={() => setReturnTypeConfirm(item.visaType)}
                              title={`បង្វិលទិន្នន័យបើកផ្តល់ប្រភេទ ${item.visaType} ត្រឡប់ចូលស្តុកវិញ`}
                              className="px-2 py-1 bg-amber-500 hover:bg-amber-600 text-white rounded-lg text-[11px] font-bold transition flex items-center gap-1 cursor-pointer shadow-2xs"
                            >
                              <RotateCcw className="w-3 h-3" />
                              <span>បង្វិល</span>
                            </button>
                          )}

                          <button
                            onClick={() => {
                              setSelectedBookModalVisa(item);
                              setModalTab('books');
                            }}
                            title="ពិនិត្យក្បាល និងចន្លោះស៊េរី"
                            className="px-2.5 py-1 bg-blue-50 text-blue-800 hover:bg-blue-100 border border-blue-200 rounded-lg text-[11px] font-bold transition flex items-center gap-1 cursor-pointer"
                          >
                            <Eye className="w-3 h-3" />
                            <span>ក្បាល ({item.books.length})</span>
                          </button>

                          <button
                            onClick={() => {
                              const cleanVt = item.visaType.trim().toUpperCase();
                              const recs = stockRecords.filter((r) => (r.visaType || '').trim().toUpperCase() === cleanVt);
                              if (recs.length === 1) {
                                setDeletingRecord(recs[0]);
                              } else if (recs.length > 1) {
                                setSelectedManageIds(recs.map((r) => r.id));
                                setIsBatchDeleteConfirmOpen(true);
                              }
                            }}
                            title={`លុបទិន្នន័យប្រភេទ ${item.visaType} ពីស្តុកជាក់ស្តែង (មិនលុបក្នុងទិន្នន័យសន្លឹកទិដ្ឋាការ)`}
                            className="p-1.5 text-rose-600 hover:text-white hover:bg-rose-600 border border-rose-200 rounded-lg text-[11px] font-bold transition flex items-center gap-1 cursor-pointer shadow-2xs"
                          >
                            <Trash2 className="w-3 h-3" />
                          </button>
                        </div>
                      </td>
                    </tr>
                  );
                })}
              </tbody>
              <tfoot>
                <tr className="bg-slate-100 font-bold text-slate-900 border-t-2 border-slate-300">
                  <td colSpan={2} className="py-3 px-3 text-center font-black">
                    សរុបរួម
                  </td>
                  <td className="py-3 px-3 text-right tabular-nums">
                    {statusItems
                      .reduce((acc, it) => acc + it.beginningStock + it.receivedK1, 0)
                      .toLocaleString()}
                  </td>
                  <td className="py-3 px-3 text-right text-emerald-700 tabular-nums">
                    +{statusItems.reduce((acc, it) => acc + it.returnedFromTeam, 0).toLocaleString()}
                  </td>
                  <td className="py-3 px-3 text-right text-amber-800 tabular-nums">
                    -{summary.totalIssuedSheets.toLocaleString()}
                  </td>
                  <td className="py-3 px-3 text-right text-rose-700 tabular-nums">
                    -{statusItems
                      .reduce((acc, it) => acc + it.damagedK2 + it.testPrintK2, 0)
                      .toLocaleString()}
                  </td>
                  <td className="py-3 px-3 text-right font-black text-blue-950 bg-blue-100/50 tabular-nums">
                    {summary.totalSheets.toLocaleString()}{' '}
                    <span className="text-[10px] font-normal">សន្លឹក</span>
                  </td>
                  <td className="py-3 px-3 text-center font-black text-indigo-950 bg-indigo-100/50">
                    <span className="tabular-nums">{summary.totalBooks.toLocaleString()}</span> ក្បាល{' '}
                    {summary.totalLoose > 0 ? `+ ${summary.totalLoose} សន្លឹក` : ''}
                  </td>
                  <td className="py-3 px-3 text-center text-slate-600">សរុបគ្រប់ប្រភេទ</td>
                  <td className="py-3 px-3 text-center">-</td>
                  <td className="py-3 px-3 text-center">
                    <button
                      onClick={() => handleOpenIssueModal()}
                      className="px-3 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-xs font-bold transition cursor-pointer"
                    >
                      បើកផ្តល់
                    </button>
                  </td>
                </tr>
              </tfoot>
            </table>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* TAB 2: ALL AVAILABLE BOOKS DATA TABLE (ដូចនៅទិន្នន័យសន្លឹកទិដ្ឋាការ) */}
      {/* ========================================================================= */}
      {activeTab === 'booksTable' && (
        <div className="bg-white rounded-2xl border border-slate-200/80 overflow-hidden shadow-xs space-y-3">
          <div className="p-4 bg-slate-900 text-white flex items-center justify-between">
            <div className="flex items-center gap-2.5">
              <BookOpen className="w-5 h-5 text-indigo-300" />
              <div>
                <h2 className="text-sm font-bold">តារាងទិន្នន័យក្បាលទិដ្ឋាការនៅសល់ជាក់ស្តែង (១ ក្បាល = ៥០ សន្លឹក)</h2>
                <p className="text-[11px] text-slate-300">
                  បង្ហាញបញ្ជីក្បាលទិដ្ឋាការទាំងអស់ដែលនៅសល់ក្នុងស្តុក K2 អាចស្វែងរក និងចម្រាញ់តាមប្រភេទ
                </p>
              </div>
            </div>

            <div className="flex items-center gap-2 text-xs">
              <span className="bg-white/10 px-3 py-1.5 rounded-lg border border-white/10 font-bold tabular-nums">
                សរុប {filteredBooks.length.toLocaleString()} ក្បាល
              </span>
            </div>
          </div>

          <div className="px-4 py-3 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3 text-xs">
            <div className="flex items-center gap-2">
              <span className="text-slate-600 font-medium">បង្ហាញក្នុងមួយទំព័រ:</span>
              <select
                value={pageSize}
                onChange={(e) => {
                  setPageSize(Number(e.target.value));
                  setCurrentPage(1);
                }}
                className="px-2.5 py-1 border border-slate-300 rounded-lg bg-white font-medium text-slate-800"
              >
                <option value={15}>15</option>
                <option value={25}>25</option>
                <option value={50}>50</option>
                <option value={100}>100</option>
              </select>
            </div>

            <div className="text-slate-600 font-medium">
              ទំព័រទី <b className="tabular-nums">{currentPage}</b> នៃ <b className="tabular-nums">{totalPages}</b> (សរុប <span className="tabular-nums">{filteredBooks.length}</span> ក្បាល)
            </div>
          </div>

          {filteredBooks.length === 0 ? (
            <div className="py-12 text-center text-slate-500 text-xs">
              <BookOpen className="w-8 h-8 mx-auto text-slate-300 mb-2" />
              <p>មិនមានក្បាលទិដ្ឋាការនៅសល់តាមលក្ខខណ្ឌស្វែងរកឡើយ</p>
            </div>
          ) : (
            <div className="overflow-x-auto">
              <table className="w-full text-xs text-left">
                <thead>
                  <tr className="bg-slate-100 text-slate-700 font-bold border-b border-slate-200">
                    <th className="py-3 px-3 text-center w-12">ល.រ</th>
                    <th className="py-3 px-3 text-center w-28">ប្រភេទទិដ្ឋាការ</th>
                    <th className="py-3 px-3 text-center w-24">ក្បាលទី</th>
                    <th className="py-3 px-3 text-center">លេខស៊េរីចាប់ផ្តើម</th>
                    <th className="py-3 px-3 text-center">លេខស៊េរីបញ្ចប់</th>
                    <th className="py-3 px-3 text-right w-28">ចំនួនសន្លឹក</th>
                    <th className="py-3 px-3 text-center w-36">ស្ថានភាពក្បាល</th>
                    <th className="py-3 px-3 text-center w-32">សកម្មភាព</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {paginatedBooks.map((book, idx) => {
                    const rowNumber = (currentPage - 1) * pageSize + idx + 1;
                    return (
                      <tr key={`${book.visaType}_${book.bookIndex}_${book.startSerial}`} className="hover:bg-blue-50/40 transition">
                        <td className="py-2.5 px-3 text-center text-slate-500 font-bold bg-slate-50/50 tabular-nums">
                          {rowNumber}
                        </td>
                        <td className="py-2.5 px-3 text-center">
                          <span className="inline-block px-2.5 py-0.5 rounded-md bg-blue-100 text-blue-900 font-black text-xs">
                            {book.visaType}
                          </span>
                        </td>
                        <td className="py-2.5 px-3 text-center font-bold text-slate-800">
                          ក្បាល #{book.bookIndex}
                        </td>
                        <td className="py-2.5 px-3 text-center font-bold text-slate-900 tabular-nums">
                          {book.startSerial}
                        </td>
                        <td className="py-2.5 px-3 text-center font-bold text-slate-900 tabular-nums">
                          {book.endSerial}
                        </td>
                        <td className="py-2.5 px-3 text-right font-black text-blue-900 tabular-nums">
                          {book.sheetCount} <span className="text-[10px] font-normal text-blue-700">សន្លឹក</span>
                        </td>
                        <td className="py-2.5 px-3 text-center">
                          {book.isPartial ? (
                            <span className="px-2.5 py-0.5 rounded-full bg-amber-100 text-amber-800 text-[10px] font-bold">
                              សល់រាយ ({book.sheetCount}/50)
                            </span>
                          ) : (
                            <span className="px-2.5 py-0.5 rounded-full bg-emerald-100 text-emerald-800 text-[10px] font-bold">
                              ពេញក្បាល (50/50)
                            </span>
                          )}
                        </td>
                        <td className="py-2.5 px-3 text-center">
                          <button
                            onClick={() => {
                              setIssueVisaType(book.visaType);
                              setIssueQuantitySheets(book.sheetCount);
                              setIssueStartSerial(book.startSerial);
                              setIssueEndSerial(book.endSerial);
                              setIsIssueModalOpen(true);
                            }}
                            className="px-2.5 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded-lg text-[11px] font-bold transition flex items-center gap-1 mx-auto cursor-pointer"
                          >
                            <Send className="w-3 h-3" />
                            <span>បើកក្បាលនេះ</span>
                          </button>
                        </td>
                      </tr>
                    );
                  })}
                </tbody>
              </table>
            </div>
          )}

          {/* Pagination Controls */}
          {totalPages > 1 && (
            <div className="p-4 border-t border-slate-200 flex items-center justify-between">
              <button
                onClick={() => setCurrentPage((p) => Math.max(1, p - 1))}
                disabled={currentPage === 1}
                className="px-3.5 py-1.5 border border-slate-300 rounded-xl text-xs font-bold text-slate-700 hover:bg-slate-50 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
              >
                ទំព័រមុន
              </button>

              <div className="flex items-center gap-1">
                {Array.from({ length: Math.min(5, totalPages) }, (_, i) => {
                  let pageNum = i + 1;
                  if (totalPages > 5 && currentPage > 3) {
                    pageNum = currentPage - 2 + i;
                    if (pageNum > totalPages) pageNum = totalPages - 4 + i;
                  }
                  return (
                    <button
                      key={pageNum}
                      onClick={() => setCurrentPage(pageNum)}
                      className={`w-8 h-8 rounded-lg text-xs font-bold transition cursor-pointer ${
                        currentPage === pageNum
                          ? 'bg-blue-600 text-white shadow-xs'
                          : 'border border-slate-300 text-slate-700 hover:bg-slate-50'
                      }`}
                    >
                      {pageNum}
                    </button>
                  );
                })}
              </div>

              <button
                onClick={() => setCurrentPage((p) => Math.min(totalPages, p + 1))}
                disabled={currentPage === totalPages}
                className="px-3.5 py-1.5 border border-slate-300 rounded-xl text-xs font-bold text-slate-700 hover:bg-slate-50 disabled:opacity-40 disabled:cursor-not-allowed cursor-pointer"
              >
                ទំព័របន្ទាប់
              </button>
            </div>
          )}
        </div>
      )}

      {/* ========================================================================= */}
      {/* TAB 3: TRANSACTION AUDIT & LOGS */}
      {/* ========================================================================= */}
      {activeTab === 'transactions' && (
        <div className="space-y-4">
          <div className="bg-white rounded-2xl border border-slate-200/80 p-4 shadow-xs">
            <h3 className="text-sm font-bold text-slate-900 mb-1 flex items-center gap-2">
              <Layers className="w-4 h-4 text-blue-600" />
              <span>ប្រវត្តិ និងចន្លោះលេខស៊េរីទូទាត់ជាក់ស្តែង</span>
            </h3>
            <p className="text-xs text-slate-500 leading-relaxed">
              រាល់ការបើកផ្តល់ជូនក្រុម ឬការបញ្ចូលស្តុកពី ក១ ត្រូវបានកត់ត្រាចន្លោះលេខស៊េរី និងទូទាត់ចេញ-ចូលក្នុងស្តុក K2 ដោយស្វ័យប្រវត្តិ។
            </p>
          </div>

          <div className="grid grid-cols-1 lg:grid-cols-2 gap-4">
            {/* Inflow Section */}
            <div className="bg-white rounded-2xl border border-emerald-200/80 overflow-hidden shadow-xs">
              <div className="bg-emerald-700 text-white p-3.5 text-xs font-bold flex items-center justify-between">
                <span>📥 ចន្លោះលេខស៊េរី បើកចូលពី ក១ / ស្តុកចាស់ (Inflow)</span>
              </div>
              <div className="p-3 max-h-96 overflow-y-auto">
                <table className="w-full text-xs text-left">
                  <thead>
                    <tr className="bg-emerald-50 text-emerald-950 font-bold border-b border-emerald-200">
                      <th className="py-2 px-2 text-center">ថ្ងៃខែ</th>
                      <th className="py-2 px-2 text-center">ប្រភេទ</th>
                      <th className="py-2 px-2 text-center">លេខចាប់ផ្តើម</th>
                      <th className="py-2 px-2 text-center">លេខបញ្ចប់</th>
                      <th className="py-2 px-2 text-right">សន្លឹក</th>
                      <th className="py-2 px-2 text-center">សកម្មភាព</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-emerald-100">
                    {statusItems.flatMap((s) => s.inflowAudits).map((inf, i) => (
                      <tr key={inf.id || i} className="hover:bg-emerald-50/50">
                        <td className="py-1.5 px-2 text-center tabular-nums text-slate-600">{inf.date}</td>
                        <td className="py-1.5 px-2 text-center font-bold text-emerald-900">{inf.operationLabel}</td>
                        <td className="py-1.5 px-2 text-center font-bold text-slate-800 tabular-nums">{inf.startSerial}</td>
                        <td className="py-1.5 px-2 text-center font-bold text-slate-800 tabular-nums">{inf.endSerial}</td>
                        <td className="py-1.5 px-2 text-right font-black text-emerald-900 tabular-nums">
                          {inf.sheets.toLocaleString()}
                        </td>
                        <td className="py-1.5 px-2 text-center">
                          <button
                            onClick={() => {
                              const found = stockRecords.find((r) => r.id === inf.id);
                              if (found) {
                                setDeletingRecord(found);
                              } else if (inf.id) {
                                handleDeleteRecordFromActualStock(inf.id);
                              }
                            }}
                            title="លុបកំណត់ត្រាបើកចូលនេះចេញពីស្តុកជាក់ស្តែង (មិនលុបក្នុងទិន្នន័យសន្លឹកទិដ្ឋាការ)"
                            className="px-2 py-0.5 rounded bg-rose-500 hover:bg-rose-600 text-white font-bold text-[10px] transition inline-flex items-center gap-0.5 cursor-pointer shadow-2xs"
                          >
                            <Trash2 className="w-2.5 h-2.5" />
                            <span>លុប</span>
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Outflow Section */}
            <div className="bg-white rounded-2xl border border-amber-200/80 overflow-hidden shadow-xs">
              <div className="bg-amber-700 text-white p-3.5 text-xs font-bold flex items-center justify-between">
                <span>📤 ចន្លោះលេខស៊េរី បានបើកជូនក្រុម/ច្រកទ្វារ (Outflow)</span>
                {issuedRecords.length > 0 && (
                  <button
                    onClick={() => setIsReturnModalOpen(true)}
                    className="px-2.5 py-1 bg-amber-800 hover:bg-amber-900 text-amber-100 rounded-lg text-[11px] font-bold transition flex items-center gap-1 cursor-pointer"
                  >
                    <RotateCcw className="w-3 h-3" />
                    <span>បង្វិលទិន្នន័យទាំងអស់</span>
                  </button>
                )}
              </div>
              <div className="p-3 max-h-96 overflow-y-auto">
                <table className="w-full text-xs text-left">
                  <thead>
                    <tr className="bg-amber-50 text-amber-950 font-bold border-b border-amber-200">
                      <th className="py-2 px-2 text-center">ថ្ងៃខែ</th>
                      <th className="py-2 px-2 text-center">ប្រភេទ</th>
                      <th className="py-2 px-2 text-center">ក្រុម/ច្រកទ្វារ</th>
                      <th className="py-2 px-2 text-center">លេខចាប់ផ្តើម</th>
                      <th className="py-2 px-2 text-center">លេខបញ្ចប់</th>
                      <th className="py-2 px-2 text-right">សន្លឹក</th>
                      <th className="py-2 px-2 text-center">សកម្មភាព</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-amber-100">
                    {statusItems.flatMap((s) => s.outflowAudits).map((outf, i) => (
                      <tr key={outf.id || i} className="hover:bg-amber-50/50">
                        <td className="py-1.5 px-2 text-center tabular-nums text-slate-600">{outf.date}</td>
                        <td className="py-1.5 px-2 text-center font-bold text-blue-900">{outf.visaType || '-'}</td>
                        <td className="py-1.5 px-2 text-center font-medium text-amber-900">{outf.sourceOrTeam}</td>
                        <td className="py-1.5 px-2 text-center font-bold text-slate-800 tabular-nums">{outf.startSerial}</td>
                        <td className="py-1.5 px-2 text-center font-bold text-slate-800 tabular-nums">{outf.endSerial}</td>
                        <td className="py-1.5 px-2 text-right font-black text-amber-900 tabular-nums">
                          {outf.sheets.toLocaleString()}
                        </td>
                        <td className="py-1.5 px-2 text-center">
                          <button
                            onClick={() => {
                              const found = stockRecords.find((r) => r.id === outf.id);
                              if (found) {
                                setSelectedRecordToReturn(found);
                              } else {
                                handleReturnBatch([outf.id]);
                              }
                            }}
                            title="បង្វិលសន្លឹកបើកផ្តល់នេះត្រឡប់ចូលស្តុកជាក់ស្តែងវិញ"
                            className="px-2 py-0.5 rounded bg-amber-500 hover:bg-amber-600 text-white font-bold text-[10px] transition inline-flex items-center gap-0.5 cursor-pointer shadow-2xs"
                          >
                            <RotateCcw className="w-2.5 h-2.5" />
                            <span>បង្វិល</span>
                          </button>
                        </td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* QUICK ISSUE TO TEAM MODAL */}
      {/* ========================================================================= */}
      {isIssueModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/60 backdrop-blur-xs animate-fade">
          <div className="bg-white rounded-2xl max-w-lg w-full shadow-2xl border border-gray-200 overflow-hidden flex flex-col max-h-[90vh]">
            <div className="bg-gradient-to-r from-emerald-800 to-teal-800 p-4 text-white flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <Send className="w-5 h-5 text-emerald-300" />
                <h3 className="text-sm font-bold">ទម្រង់បើកផ្តល់ទិដ្ឋាការជូនក្រុម (កាត់ស្តុកជាក់ស្តែង)</h3>
              </div>
              <button
                onClick={() => setIsIssueModalOpen(false)}
                className="p-1 rounded-lg hover:bg-white/20 text-white/80 hover:text-white transition"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <form onSubmit={handleConfirmIssueToTeam} className="p-5 overflow-y-auto space-y-4 flex-1">
              <div className="bg-emerald-50 border border-emerald-200 rounded-lg p-3 text-xs text-emerald-900">
                រាល់ការបើកផ្តល់ជូនក្រុមនឹងកាត់ទូទាត់ចំនួនសន្លឹក និងចន្លោះលេខស៊េរីចេញពីស្តុកជាក់ស្តែង K2 ភ្លាមៗ។
              </div>

              {/* Visa Type & Quantity */}
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-bold text-gray-700 mb-1">
                    ប្រភេទទិដ្ឋាការ <span className="text-red-500">*</span>
                  </label>
                  <select
                    value={issueVisaType}
                    onChange={(e) => handleIssueVisaTypeChange(e.target.value)}
                    className="w-full px-3 py-2 text-xs rounded-lg border border-gray-300 focus:ring-2 focus:ring-emerald-500 bg-white font-bold text-gray-800"
                    required
                  >
                    {allVisaTypes.map((vt) => (
                      <option key={vt} value={vt}>
                        ទិដ្ឋាការប្រភេទ {vt}
                      </option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="block text-xs font-bold text-gray-700 mb-1">
                    ចំនួនបើក (សន្លឹក) <span className="text-red-500">*</span>
                  </label>
                  <input
                    type="number"
                    min={1}
                    value={issueQuantitySheets}
                    onChange={(e) => handleIssueQuantityChange(Number(e.target.value))}
                    className="w-full px-3 py-2 text-xs rounded-lg border border-gray-300 focus:ring-2 focus:ring-emerald-500 bg-white font-mono font-bold text-gray-800"
                    required
                  />
                </div>
              </div>

              {/* Serial Range */}
              {(() => {
                const targetVisaItem = statusItems.find((s) => s.visaType === issueVisaType);
                const allRanges = targetVisaItem?.remainingRanges || [];
                const sufficientRanges =
                  issueQuantitySheets > 0
                    ? allRanges.filter((r) => r.count >= issueQuantitySheets)
                    : allRanges;

                return (
                  <div className="space-y-2">
                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                      <div>
                        <label className="block text-xs font-bold text-gray-700 mb-1">
                          លេខស៊េរីចាប់ពី (ជ្រើសពីស្តុកជាក់ស្តែង)
                        </label>
                        {allRanges.length > 0 ? (
                          <select
                            value={issueStartSerial}
                            onChange={(e) => {
                              const sStart = e.target.value;
                              setIssueStartSerial(sStart);
                              if (sStart && issueQuantitySheets > 0) {
                                const parsed = parseSerialNumber(sStart);
                                if (parsed) {
                                  const endNum = parsed.num + BigInt(issueQuantitySheets) - 1n;
                                  setIssueEndSerial(
                                    formatSerialNumber(parsed.prefix, endNum, parsed.padLength)
                                  );
                                }
                              }
                            }}
                            className="w-full px-3 py-2 text-xs rounded-lg border border-gray-300 focus:ring-2 focus:ring-emerald-500 bg-white font-mono font-bold cursor-pointer"
                          >
                            <option value="">
                              -- ជ្រើសរើសចន្លោះលេខស៊េរី ({sufficientRanges.length} ចន្លោះគ្រប់ {issueQuantitySheets} សន្លឹក) --
                            </option>
                            {sufficientRanges.map((r) => (
                              <option key={r.startSerial} value={r.startSerial}>
                                {r.startSerial} ដល់ {r.endSerial} (នៅសល់ {r.count.toLocaleString()} សន្លឹក)
                              </option>
                            ))}
                            {sufficientRanges.length === 0 && allRanges.length > 0 && (
                              <>
                                <option disabled value="">
                                  ⚠️ គ្មានចន្លោះលេខណាគ្រប់ {issueQuantitySheets} សន្លឹកទេ (ស្តុកអតិបរមា {Math.max(...allRanges.map((r) => r.count))} សន្លឹក)
                                </option>
                                {allRanges.map((r) => (
                                  <option key={r.startSerial} value={r.startSerial} disabled className="text-gray-400">
                                    [មិនគ្រប់ចំនួន] {r.startSerial} ដល់ {r.endSerial} (មាន {r.count} សន្លឹក)
                                  </option>
                                ))}
                              </>
                            )}
                          </select>
                        ) : (
                          <input
                            type="text"
                            value={issueStartSerial}
                            onChange={(e) => setIssueStartSerial(e.target.value)}
                            placeholder="0000001"
                            className="w-full px-3 py-2 text-xs rounded-lg border border-gray-300 focus:ring-2 focus:ring-emerald-500 bg-white font-mono font-bold"
                          />
                        )}
                      </div>

                      <div>
                        <label className="block text-xs font-bold text-gray-700 mb-1">ដល់លេខស៊េរី (ស្វ័យប្រវត្តិ)</label>
                        <input
                          type="text"
                          value={issueEndSerial}
                          onChange={(e) => setIssueEndSerial(e.target.value)}
                          placeholder="0000050"
                          className="w-full px-3 py-2 text-xs rounded-lg border border-gray-300 focus:ring-2 focus:ring-emerald-500 bg-white font-mono font-bold"
                        />
                      </div>
                    </div>

                    {targetVisaItem && (
                      <div className="text-[11px] text-gray-600 bg-gray-50 px-2.5 py-1 rounded border border-gray-200 flex flex-wrap items-center justify-between">
                        <span>
                          ស្តុកជាក់ស្តែងប្រភេទ <b>{issueVisaType}</b>: <b className="text-blue-700">{targetVisaItem.officeRemainingSheets.toLocaleString()} សន្លឹក</b> ({allRanges.length} ចន្លោះ)
                        </span>
                        {issueQuantitySheets > 0 && sufficientRanges.length === 0 && allRanges.length > 0 && (
                          <span className="text-amber-700 font-bold">
                            ⚠️ ចំនួន {issueQuantitySheets} សន្លឹក លើសពីស្តុកក្នុងចន្លោះនីមួយៗ
                          </span>
                        )}
                      </div>
                    )}
                  </div>
                );
              })()}

              {/* Date & Time */}
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-bold text-gray-700 mb-1">កាលបរិច្ឆេទ</label>
                  <CustomDatePicker
                    required
                    value={issueDate}
                    onChange={(d) => setIssueDate(d)}
                    className="py-1.5 text-xs"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-gray-700 mb-1">ម៉ោង</label>
                  <input
                    type="time"
                    value={issueTime}
                    onChange={(e) => setIssueTime(e.target.value)}
                    className="w-full px-3 py-2 text-xs rounded-lg border border-gray-300 focus:ring-2 focus:ring-emerald-500 bg-white"
                  />
                </div>
              </div>

              {/* Team Receiver */}
              <div>
                <label className="block text-xs font-bold text-gray-700 mb-1">
                  ក្រុមផ្តល់ទិដ្ឋាការ.របក <span className="text-red-500">*</span>
                </label>
                <select
                  value={issueTeamId}
                  onChange={(e) => setIssueTeamId(e.target.value)}
                  className="w-full px-3 py-2 text-xs rounded-lg border border-gray-300 focus:ring-2 focus:ring-emerald-500 bg-white font-medium"
                  required
                >
                  <option value="">-- ជ្រើសរើសក្រុម --</option>
                  {(categories.visaTeamsRobok || []).map((t) => (
                    <option key={t.id} value={t.name}>
                      {t.name}
                    </option>
                  ))}
                  {(categories.visaTeams || []).map((t) => (
                    <option key={t.id} value={t.name}>
                      {t.name}
                    </option>
                  ))}
                </select>
              </div>

              {/* Personnel fields */}
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block text-xs font-bold text-gray-700 mb-1">អ្នកស្នើសុំ</label>
                  <input
                    type="text"
                    value={issueRequesterName}
                    onChange={(e) => setIssueRequesterName(e.target.value)}
                    placeholder="ឈ្មោះអ្នកស្នើសុំ"
                    className="w-full px-3 py-2 text-xs rounded-lg border border-gray-300 focus:ring-2 focus:ring-emerald-500 bg-white"
                  />
                </div>

                <div>
                  <label className="block text-xs font-bold text-gray-700 mb-1">អ្នកមកបើក</label>
                  <input
                    type="text"
                    value={issueCollectorName}
                    onChange={(e) => setIssueCollectorName(e.target.value)}
                    placeholder="ឈ្មោះអ្នកមកបើក"
                    className="w-full px-3 py-2 text-xs rounded-lg border border-gray-300 focus:ring-2 focus:ring-emerald-500 bg-white"
                  />
                </div>
              </div>

              <div className="flex items-center justify-end gap-2 pt-4 border-t border-gray-200">
                <button
                  type="button"
                  onClick={() => setIsIssueModalOpen(false)}
                  className="px-4 py-2 rounded-lg text-xs font-bold text-gray-700 hover:bg-gray-100 transition cursor-pointer"
                >
                  បោះបង់
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 rounded-lg text-xs font-bold bg-emerald-600 hover:bg-emerald-700 text-white shadow-xs transition flex items-center gap-1.5 cursor-pointer"
                >
                  <Send className="w-4 h-4" />
                  <span>រក្សាទុក និងកាត់ស្តុក</span>
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* EXCEL IMPORT MODAL WITH AUTO-SORT (តម្រៀបអោយស្អាត) */}
      {/* ========================================================================= */}
      {isImportModalOpen && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/60 backdrop-blur-xs animate-fade">
          <div className="bg-white rounded-2xl max-w-2xl w-full shadow-2xl border border-gray-200 overflow-hidden flex flex-col max-h-[90vh]">
            <div className="bg-gradient-to-r from-blue-900 to-indigo-900 p-4 text-white flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <FileSpreadsheet className="w-5 h-5 text-blue-300" />
                <div>
                  <h3 className="text-sm font-bold">នាំចូលទិន្នន័យស្តុកជាក់ស្តែងពី Excel</h3>
                  <p className="text-[11px] text-blue-200">ប្រព័ន្ធនឹងអាន និងតម្រៀបទិន្នន័យអោយស្អាតដោយស្វ័យប្រវត្តិ</p>
                </div>
              </div>
              <button
                onClick={() => {
                  setIsImportModalOpen(false);
                  setParsedImportRows([]);
                  setImportFile(null);
                }}
                className="p-1 rounded-lg hover:bg-white/20 text-white/80 hover:text-white transition"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="p-5 overflow-y-auto space-y-4 flex-1">
              {/* Download Template Banner */}
              <div className="bg-blue-50 border border-blue-200 rounded-xl p-3.5 flex items-center justify-between">
                <div className="flex items-center gap-2.5">
                  <Download className="w-5 h-5 text-blue-600 shrink-0" />
                  <div>
                    <div className="text-xs font-bold text-blue-950">ទាញយកទម្រង់គំរូ Excel (Template)</div>
                    <div className="text-[11px] text-blue-700">ទាញយកឯកសារគំរូសម្រាប់បំពេញទិន្នន័យស្តុកសន្លឹកទិដ្ឋាការ</div>
                  </div>
                </div>
                <button
                  type="button"
                  onClick={handleDownloadExcelTemplate}
                  className="px-3 py-1.5 bg-blue-600 hover:bg-blue-700 text-white rounded-lg text-xs font-bold transition flex items-center gap-1.5 cursor-pointer shrink-0"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>ទាញយកទម្រង់</span>
                </button>
              </div>

              {/* Upload Dropzone */}
              <div
                onDragOver={(e) => {
                  e.preventDefault();
                  setImportDragActive(true);
                }}
                onDragLeave={() => setImportDragActive(false)}
                onDrop={(e) => {
                  e.preventDefault();
                  setImportDragActive(false);
                  if (e.dataTransfer.files && e.dataTransfer.files[0]) {
                    handleProcessExcelFile(e.dataTransfer.files[0]);
                  }
                }}
                className={`border-2 border-dashed rounded-xl p-6 text-center transition cursor-pointer ${
                  importDragActive ? 'border-blue-500 bg-blue-50/50' : 'border-gray-300 hover:border-blue-400 bg-gray-50/50'
                }`}
                onClick={() => document.getElementById('excel-file-input')?.click()}
              >
                <input
                  id="excel-file-input"
                  type="file"
                  accept=".xlsx, .xls, .csv"
                  className="hidden"
                  onChange={(e) => {
                    if (e.target.files && e.target.files[0]) {
                      handleProcessExcelFile(e.target.files[0]);
                    }
                  }}
                />
                <Upload className="w-8 h-8 text-blue-600 mx-auto mb-2" />
                <div className="text-xs font-bold text-gray-800">
                  {importFile ? importFile.name : 'ចុចទីនេះដើម្បីជ្រើសរើសឯកសារ Excel ឬ អូសទម្លាក់ទីនេះ'}
                </div>
                <div className="text-[11px] text-gray-500 mt-1">គាំទ្រឯកសារ .xlsx, .xls, .csv</div>
              </div>

              {/* Preview Sorted Results */}
              {parsedImportRows.length > 0 && (
                <div className="space-y-2 border border-gray-200 rounded-xl p-3.5 bg-gray-50/50">
                  <div className="flex items-center justify-between text-xs font-bold text-gray-800">
                    <span className="flex items-center gap-1.5 text-emerald-700">
                      <Sparkles className="w-4 h-4 text-emerald-600" />
                      ទិន្នន័យបានអាន និងតម្រៀបស្អាតរួច ({parsedImportRows.length} ជួរ)
                    </span>
                    <span className="font-mono text-gray-500">
                      សរុប {parsedImportRows.reduce((a, b) => a + (b.totalSheets || 0), 0).toLocaleString()} សន្លឹក
                    </span>
                  </div>

                  <div className="max-h-56 overflow-y-auto border border-gray-200 rounded-lg bg-white">
                    <table className="w-full text-[11px] text-left">
                      <thead>
                        <tr className="bg-gray-100 text-gray-700 font-bold border-b border-gray-200 sticky top-0">
                          <th className="py-2 px-2 text-center">ល.រ</th>
                          <th className="py-2 px-2 text-center">ប្រភេទ</th>
                          <th className="py-2 px-2 text-center">កាលបរិច្ឆេទ</th>
                          <th className="py-2 px-2 text-right">ចំនួនសន្លឹក</th>
                          <th className="py-2 px-2 font-mono text-center">លេខចាប់ពី</th>
                          <th className="py-2 px-2 font-mono text-center">ដល់លេខ</th>
                          <th className="py-2 px-2 text-center">ប្រតិបត្តិការ</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-gray-200">
                        {parsedImportRows.map((r, idx) => (
                          <tr key={r.id} className="hover:bg-blue-50/30">
                            <td className="py-1.5 px-2 text-center font-mono text-gray-500">{idx + 1}</td>
                            <td className="py-1.5 px-2 text-center font-bold font-mono text-blue-900">{r.visaType}</td>
                            <td className="py-1.5 px-2 text-center font-mono">{r.date}</td>
                            <td className="py-1.5 px-2 text-right font-black text-gray-900">{r.totalSheets?.toLocaleString()}</td>
                            <td className="py-1.5 px-2 text-center font-mono font-bold text-gray-800">{r.startSerial || '-'}</td>
                            <td className="py-1.5 px-2 text-center font-mono font-bold text-gray-800">{r.endSerial || '-'}</td>
                            <td className="py-1.5 px-2 text-center text-gray-600">{r.operationType}</td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}

              {/* Data Protection / Import Mode Selection */}
              {parsedImportRows.length > 0 && (
                <div className="bg-slate-50 border border-slate-200 rounded-xl p-3.5 space-y-2.5">
                  <div className="text-xs font-bold text-slate-800 flex items-center gap-2">
                    <Info className="w-4 h-4 text-blue-600" />
                    <span>ជម្រើសនៃការរក្សាទុក និងការពារទិន្នន័យស្តុក</span>
                  </div>
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-2 text-xs">
                    <label
                      onClick={() => setImportMode('append')}
                      className={`p-3 rounded-xl border flex items-start gap-2.5 cursor-pointer transition ${
                        importMode === 'append'
                          ? 'border-emerald-500 bg-emerald-50/80 text-emerald-950 ring-2 ring-emerald-500/40 shadow-xs'
                          : 'border-slate-200 hover:border-slate-300 bg-white text-slate-700'
                      }`}
                    >
                      <input
                        type="radio"
                        name="importMode"
                        checked={importMode === 'append'}
                        onChange={() => setImportMode('append')}
                        className="mt-0.5 text-emerald-600 focus:ring-emerald-500"
                      />
                      <div>
                        <div className="flex items-center gap-1.5 font-bold text-emerald-950">
                          <span>បន្ថែមលើស្តុកដែលមានស្រាប់ (Append)</span>
                          <span className="text-[10px] px-1.5 py-0.2 bg-emerald-100 text-emerald-800 rounded font-semibold border border-emerald-300">ណែនាំ</span>
                        </div>
                        <div className="text-[11px] text-slate-600 mt-0.5 leading-snug">
                          រក្សាទុកទិន្នន័យស្តុកចាស់ដែលមានស្រាប់ទាំងអស់ ហើយបន្ថែមទិន្នន័យពី Excel ថ្មីនេះចូលដោយមិនបាត់បង់ទិន្នន័យចាស់ឡើយ។
                        </div>
                      </div>
                    </label>

                    <label
                      onClick={() => setImportMode('overwrite')}
                      className={`p-3 rounded-xl border flex items-start gap-2.5 cursor-pointer transition ${
                        importMode === 'overwrite'
                          ? 'border-amber-500 bg-amber-50/80 text-amber-950 ring-2 ring-amber-500/40 shadow-xs'
                          : 'border-slate-200 hover:border-slate-300 bg-white text-slate-700'
                      }`}
                    >
                      <input
                        type="radio"
                        name="importMode"
                        checked={importMode === 'overwrite'}
                        onChange={() => setImportMode('overwrite')}
                        className="mt-0.5 text-amber-600 focus:ring-amber-500"
                      />
                      <div>
                        <div className="flex items-center gap-1.5 font-bold text-amber-950">
                          <span>ជំនួសទិន្នន័យចាស់ (Clean Replace)</span>
                          <span className="text-[10px] px-1.5 py-0.2 bg-amber-100 text-amber-900 rounded font-semibold border border-amber-300">ប្រុងប្រយ័ត្ន</span>
                        </div>
                        <div className="text-[11px] text-amber-800/80 mt-0.5 leading-snug">
                          ⚠️ នឹងជម្រះទិន្នន័យចាស់ទាំងអស់ក្នុងស្តុកជាក់ស្តែង K2 ហើយជំនួសដោយទិន្នន័យថ្មីនេះ (ប្រើពេលចង់កំណត់ស្តុកជាក់ស្តែងឡើងវិញសុទ្ធសាធ)។
                        </div>
                      </div>
                    </label>
                  </div>
                </div>
              )}
            </div>

            <div className="p-4 border-t border-gray-200 flex items-center justify-end gap-2 bg-gray-50">
              <button
                type="button"
                onClick={() => {
                  setIsImportModalOpen(false);
                  setParsedImportRows([]);
                  setImportFile(null);
                }}
                className="px-4 py-2 rounded-lg text-xs font-bold text-gray-700 hover:bg-gray-200 transition cursor-pointer"
              >
                បោះបង់
              </button>
              <button
                type="button"
                onClick={handleConfirmBatchImport}
                disabled={parsedImportRows.length === 0}
                className="px-5 py-2 rounded-lg text-xs font-bold bg-blue-600 hover:bg-blue-700 disabled:opacity-40 disabled:cursor-not-allowed text-white shadow-xs transition flex items-center gap-1.5 cursor-pointer"
              >
                <Check className="w-4 h-4" />
                <span>បញ្ជាក់ការនាំចូល ({parsedImportRows.length} ជួរ)</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* SINGLE VISA DETAIL MODAL */}
      {/* ========================================================================= */}
      {selectedBookModalVisa && (
        <div className="fixed inset-0 z-50 flex items-center justify-center p-4 bg-black/60 backdrop-blur-xs animate-fade">
          <div className="bg-white rounded-2xl max-w-3xl w-full shadow-2xl border border-gray-200 overflow-hidden flex flex-col max-h-[90vh]">
            <div className="bg-gradient-to-r from-blue-900 to-indigo-900 p-4 text-white flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <BookOpen className="w-5 h-5 text-blue-300" />
                <div>
                  <h3 className="text-sm font-bold">
                    ព័ត៌មានលម្អិតក្បាលទិដ្ឋាការប្រភេទ {selectedBookModalVisa.visaType}
                  </h3>
                  <p className="text-[11px] text-blue-200">
                    ចន្លោះលេខស៊េរីនៅសល់៖ {selectedBookModalVisa.serialRangeDisplay}
                  </p>
                </div>
              </div>
              <button
                onClick={() => setSelectedBookModalVisa(null)}
                className="p-1 rounded-lg hover:bg-white/20 text-white/80 hover:text-white transition"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="p-5 overflow-y-auto space-y-4 flex-1">
              <div className="flex items-center justify-between border-b border-gray-200 pb-3">
                <div className="flex items-center gap-2">
                  <button
                    onClick={() => setModalTab('books')}
                    className={`px-3 py-1.5 rounded-md text-xs font-bold transition cursor-pointer flex items-center gap-1.5 ${
                      modalTab === 'books'
                        ? 'bg-blue-600 text-white shadow-xs'
                        : 'bg-gray-100 text-gray-700 hover:bg-gray-200'
                    }`}
                  >
                    <BookOpen className="w-3.5 h-3.5" />
                    <span>បញ្ជីក្បាលនៅសល់ ({selectedBookModalVisa.books.length} ក្បាល)</span>
                  </button>

                  <button
                    onClick={() => setModalTab('audit')}
                    className={`px-3 py-1.5 rounded-md text-xs font-bold transition cursor-pointer flex items-center gap-1.5 ${
                      modalTab === 'audit'
                        ? 'bg-blue-600 text-white shadow-xs'
                        : 'bg-gray-100 text-gray-700 hover:bg-gray-200'
                    }`}
                  >
                    <Layers className="w-3.5 h-3.5" />
                    <span>តារាងទូទាត់ចន្លោះលេខស៊េរី (បើក ក១ - ជូនក្រុម = នៅសល់)</span>
                  </button>
                </div>

                <div className="text-xs text-gray-500 font-mono">
                  សរុប {selectedBookModalVisa.officeRemainingSheets.toLocaleString()} សន្លឹក
                </div>
              </div>

              {modalTab === 'books' ? (
                <>
                  <div className="bg-amber-50 border border-amber-200 rounded-lg p-3 text-xs text-amber-900 flex items-center gap-2">
                    <Info className="w-4 h-4 text-amber-700 shrink-0" />
                    <span>
                      រាល់ទិដ្ឋាការស្អិត ១ ក្បាល មាន ៥០ សន្លឹក។ តារាងខាងក្រោមកំណត់ការបែងចែកក្បាលនីមួយៗតាមចន្លោះលេខស៊េរីដែលនៅសល់ជាក់ស្តែង៖
                    </span>
                  </div>

                  {selectedBookModalVisa.books.length === 0 ? (
                    <div className="py-8 text-center text-gray-500">
                      មិនមានក្បាលទិដ្ឋាការនៅសល់ក្នុងស្តុកឡើយ (០ សន្លឹក)
                    </div>
                  ) : (
                    <div className="border border-gray-200 rounded-lg overflow-hidden">
                      <table className="w-full text-xs text-left">
                        <thead>
                          <tr className="bg-gray-100 text-gray-700 font-bold border-b border-gray-200">
                            <th className="py-2.5 px-3 text-center w-16">ក្បាលទី</th>
                            <th className="py-2.5 px-3 text-center w-36 font-mono">លេខស៊េរីចាប់ផ្តើម</th>
                            <th className="py-2.5 px-3 text-center w-36 font-mono">លេខស៊េរីបញ្ចប់</th>
                            <th className="py-2.5 px-3 text-right w-28">ចំនួនសន្លឹក</th>
                            <th className="py-2.5 px-3 text-center">ស្ថានភាពក្បាល</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-gray-200">
                          {selectedBookModalVisa.books.map((b) => (
                            <tr key={b.bookIndex} className="hover:bg-blue-50/30 transition">
                              <td className="py-2 px-3 text-center font-bold text-gray-700 bg-gray-50/50">
                                #{b.bookIndex}
                              </td>
                              <td className="py-2 px-3 text-center font-mono font-semibold text-gray-800">
                                {b.startSerial}
                              </td>
                              <td className="py-2 px-3 text-center font-mono font-semibold text-gray-800">
                                {b.endSerial}
                              </td>
                              <td className="py-2 px-3 text-right font-black text-blue-900 font-sans">
                                {b.sheetCount} សន្លឹក
                              </td>
                              <td className="py-2 px-3 text-center">
                                {b.isPartial ? (
                                  <span className="px-2 py-0.5 rounded bg-amber-100 text-amber-800 text-[10px] font-bold">
                                    សល់រាយ ({b.sheetCount}/50)
                                  </span>
                                ) : (
                                  <span className="px-2 py-0.5 rounded bg-emerald-100 text-emerald-800 text-[10px] font-bold">
                                    ពេញក្បាល (50/50)
                                  </span>
                                )}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                </>
              ) : (
                <div className="space-y-4">
                  <div className="bg-blue-50 border border-blue-200 rounded-lg p-3 text-xs text-blue-900 leading-relaxed">
                    <b>រូបមន្តទូទាត់ចន្លោះលេខស៊េរី៖</b> យកចន្លោះស៊េរី <b>បើកចូលពី ក១ / ស្តុកចាស់</b> ដកទូទាត់ជាមួយចន្លោះស៊េរី <b>បានបើកជូនក្រុម/ច្រកទ្វារ</b> ដើម្បីទទួលបានចន្លោះស៊េរីនៅសល់ជាក់ស្តែង។
                  </div>

                  {/* Inflow Section */}
                  <div className="border border-emerald-200 rounded-lg overflow-hidden bg-emerald-50/20">
                    <div className="bg-emerald-100/70 px-3 py-2 text-xs font-bold text-emerald-900 flex items-center justify-between">
                      <span>📥 ១. ចន្លោះលេខស៊េរី បើកចូលពី ក១ / ស្តុកចាស់ (Inflow)</span>
                      <span>{selectedBookModalVisa.inflowAudits.length} កំណត់ត្រា</span>
                    </div>
                    {selectedBookModalVisa.inflowAudits.length === 0 ? (
                      <div className="p-3 text-center text-xs text-gray-500">មិនមានទិន្នន័យបើកចូលស៊េរី</div>
                    ) : (
                      <table className="w-full text-xs text-left">
                        <thead>
                          <tr className="bg-emerald-50 text-gray-700 font-semibold border-b border-emerald-200">
                            <th className="py-2 px-3 text-center w-10">ល.រ</th>
                            <th className="py-2 px-3 text-center w-24">កាលបរិច្ឆេទ</th>
                            <th className="py-2 px-3 text-center w-28">ប្រតិបត្តិការ</th>
                            <th className="py-2 px-3 font-mono text-center">លេខចាប់ផ្តើម</th>
                            <th className="py-2 px-3 font-mono text-center">លេខបញ្ចប់</th>
                            <th className="py-2 px-3 text-right w-24 font-sans">ចំនួនសន្លឹក</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-emerald-100">
                          {selectedBookModalVisa.inflowAudits.map((inf, i) => (
                            <tr key={inf.id} className="bg-white">
                              <td className="py-1.5 px-3 text-center text-gray-500 font-mono">{i + 1}</td>
                              <td className="py-1.5 px-3 text-center font-mono">{inf.date}</td>
                              <td className="py-1.5 px-3 text-center font-medium text-emerald-800">{inf.operationLabel}</td>
                              <td className="py-1.5 px-3 text-center font-mono font-bold text-gray-800">{inf.startSerial}</td>
                              <td className="py-1.5 px-3 text-center font-mono font-bold text-gray-800">{inf.endSerial}</td>
                              <td className="py-1.5 px-3 text-right font-black text-emerald-900">{inf.sheets.toLocaleString()}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    )}
                  </div>

                  {/* Outflow Section */}
                  <div className="border border-amber-200 rounded-lg overflow-hidden bg-amber-50/20">
                    <div className="bg-amber-100/70 px-3 py-2 text-xs font-bold text-amber-900 flex items-center justify-between">
                      <span>📤 ២. ចន្លោះលេខស៊េរី បានបើកជូនក្រុម/ច្រកទ្វារ (Outflow)</span>
                      <span>{selectedBookModalVisa.outflowAudits.length} កំណត់ត្រា</span>
                    </div>
                    {selectedBookModalVisa.outflowAudits.length === 0 ? (
                      <div className="p-3 text-center text-xs text-gray-500">មិនមានទិន្នន័យបើកជូនក្រុម</div>
                    ) : (
                      <table className="w-full text-xs text-left">
                        <thead>
                          <tr className="bg-amber-50 text-gray-700 font-semibold border-b border-amber-200">
                            <th className="py-2 px-3 text-center w-10">ល.រ</th>
                            <th className="py-2 px-3 text-center w-24">កាលបរិច្ឆេទ</th>
                            <th className="py-2 px-3 text-center">ក្រុម/ច្រកទ្វារ</th>
                            <th className="py-2 px-3 font-mono text-center">លេខចាប់ផ្តើម</th>
                            <th className="py-2 px-3 font-mono text-center">លេខបញ្ចប់</th>
                            <th className="py-2 px-3 text-right w-24 font-sans">ចំនួនសន្លឹក</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-amber-100">
                          {selectedBookModalVisa.outflowAudits.map((outf, i) => (
                            <tr key={outf.id} className="bg-white">
                              <td className="py-1.5 px-3 text-center text-gray-500 font-mono">{i + 1}</td>
                              <td className="py-1.5 px-3 text-center font-mono">{outf.date}</td>
                              <td className="py-1.5 px-3 text-center font-medium text-amber-900">{outf.sourceOrTeam}</td>
                              <td className="py-1.5 px-3 text-center font-mono font-bold text-gray-800">{outf.startSerial}</td>
                              <td className="py-1.5 px-3 text-center font-mono font-bold text-gray-800">{outf.endSerial}</td>
                              <td className="py-1.5 px-3 text-right font-black text-amber-900">{outf.sheets.toLocaleString()}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    )}
                  </div>

                  {/* Result Remaining */}
                  <div className="border border-blue-200 rounded-lg overflow-hidden bg-blue-50/20">
                    <div className="bg-blue-100/70 px-3 py-2 text-xs font-bold text-blue-950 flex items-center justify-between">
                      <span>📊 ៣. ចន្លោះលេខស៊េរីនៅសល់ជាក់ស្តែង (Resulting Balance)</span>
                      <span>{selectedBookModalVisa.remainingRanges.length} ចន្លោះ</span>
                    </div>
                    {selectedBookModalVisa.remainingRanges.length === 0 ? (
                      <div className="p-3 text-center text-xs text-red-600 font-bold">អស់ពីស្តុក (០ សន្លឹក)</div>
                    ) : (
                      <table className="w-full text-xs text-left">
                        <thead>
                          <tr className="bg-blue-50 text-gray-700 font-semibold border-b border-blue-200">
                            <th className="py-2 px-3 text-center w-10">ល.រ</th>
                            <th className="py-2 px-3 font-mono text-center">លេខចាប់ផ្តើមនៅសល់</th>
                            <th className="py-2 px-3 font-mono text-center">លេខបញ្ចប់នៅសល់</th>
                            <th className="py-2 px-3 text-right font-sans">ចំនួនសន្លឹក</th>
                            <th className="py-2 px-3 text-center font-sans">ស្មើនឹងក្បាល</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-blue-100">
                          {selectedBookModalVisa.remainingRanges.map((rng, i) => (
                            <tr key={i} className="bg-white">
                              <td className="py-2 px-3 text-center text-gray-500 font-mono">{i + 1}</td>
                              <td className="py-2 px-3 text-center font-mono font-bold text-blue-900">{rng.startSerial}</td>
                              <td className="py-2 px-3 text-center font-mono font-bold text-blue-900">{rng.endSerial}</td>
                              <td className="py-2 px-3 text-right font-black text-blue-950 font-sans">{rng.count.toLocaleString()} សន្លឹក</td>
                              <td className="py-2 px-3 text-center font-bold text-amber-900">
                                {Math.floor(rng.count / 50)} ក្បាល {rng.count % 50 > 0 ? `+ ${rng.count % 50} សន្លឹក` : ''}
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    )}
                  </div>
                </div>
              )}
            </div>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* MANAGE & DELETE ACTUAL STOCK MODAL (លុបទិន្នន័យស្តុកជាក់ស្តែង K2 តែប៉ុណ្ណោះ) */}
      {/* ========================================================================= */}
      {isManageDeleteModalOpen && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 animate-fade">
          <div className="bg-white rounded-2xl max-w-4xl w-full shadow-2xl border border-slate-200 overflow-hidden flex flex-col max-h-[90vh]">
            {/* Header */}
            <div className="bg-gradient-to-r from-rose-700 to-rose-800 p-4 text-white flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <div className="p-2 bg-white/10 rounded-xl">
                  <Trash2 className="w-5 h-5 text-rose-200" />
                </div>
                <div>
                  <h3 className="text-sm font-bold">គ្រប់គ្រង និងលុបទិន្នន័យចេញពីស្តុកជាក់ស្តែង</h3>
                  <p className="text-[11px] text-rose-100 mt-0.5">
                    ជ្រើសរើសទិន្នន័យដើម្បីលុបចេញពីស្តុកជាក់ស្តែងសម្រាប់បើកផ្តល់តាមក្រុម (K2)
                  </p>
                </div>
              </div>
              <button
                onClick={() => setIsManageDeleteModalOpen(false)}
                className="p-1.5 text-white/80 hover:text-white hover:bg-white/10 rounded-lg transition cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Safe Notice Banner */}
            <div className="bg-rose-50 border-b border-rose-200 px-4 py-3 text-xs text-rose-950 flex items-start gap-2.5">
              <div className="p-1 bg-rose-100 rounded-lg text-rose-700 shrink-0 mt-0.5">
                <AlertTriangle className="w-4 h-4" />
              </div>
              <div className="leading-relaxed">
                <b>🛡️ សុវត្ថិភាពទិន្នន័យ៖</b> ការលុបទិន្នន័យនៅទីនេះ គឺលុបតែចេញពី <b>ស្តុកជាក់ស្តែងសម្រាប់បើកផ្តល់តាមក្រុម (K2)</b> ប៉ុណ្ណោះ ដោយ<b>រក្សាទុកទិន្នន័យក្នុង «ទិន្នន័យសន្លឹកទិដ្ឋាការ»</b> ទាំងអស់មិនឱ្យបាត់បង់ឡើយ។
              </div>
            </div>

            {/* Filters and Batch Actions */}
            <div className="p-4 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3">
              <div className="flex flex-wrap items-center gap-3 flex-1 min-w-[280px]">
                {/* Search */}
                <div className="relative flex-1 min-w-[200px]">
                  <Search className="w-4 h-4 text-slate-400 absolute left-3 top-1/2 -translate-y-1/2" />
                  <input
                    type="text"
                    value={manageSearch}
                    onChange={(e) => setManageSearch(e.target.value)}
                    placeholder="ស្វែងរកតាមលេខស៊េរី, ប្រភេទ, ថ្ងៃខែ..."
                    className="w-full pl-9 pr-3 py-1.5 bg-white border border-slate-300 rounded-xl text-xs focus:ring-2 focus:ring-rose-500 focus:border-rose-500"
                  />
                </div>

                {/* Filter Type */}
                <select
                  value={manageVisaFilter}
                  onChange={(e) => setManageVisaFilter(e.target.value)}
                  className="px-3 py-1.5 bg-white border border-slate-300 rounded-xl text-xs font-bold text-slate-700 focus:ring-2 focus:ring-rose-500"
                >
                  <option value="ALL">គ្រប់ប្រភេទ ({stockRecords.filter((r) => r.stockType === 'sticker').length})</option>
                  {allVisaTypes.map((vt) => (
                    <option key={vt} value={vt}>
                      ប្រភេទ {vt} ({stockRecords.filter((r) => r.stockType === 'sticker' && (r.visaType || '').trim().toUpperCase() === vt.toUpperCase()).length})
                    </option>
                  ))}
                </select>
              </div>

              {/* Batch Action Button */}
              {selectedManageIds.length > 0 && (
                <button
                  onClick={() => setIsBatchDeleteConfirmOpen(true)}
                  className="px-3.5 py-1.5 bg-rose-600 hover:bg-rose-700 active:scale-98 text-white rounded-xl text-xs font-bold transition flex items-center gap-1.5 cursor-pointer shadow-xs"
                >
                  <Trash2 className="w-3.5 h-3.5" />
                  <span>លុបដែលបានជ្រើស ({selectedManageIds.length}) ចេញពីស្តុក</span>
                </button>
              )}
            </div>

            {/* Records Table */}
            <div className="p-4 overflow-y-auto flex-1 max-h-[50vh]">
              {manageStockRecords.length === 0 ? (
                <div className="py-12 text-center text-slate-500 text-xs">
                  មិនមានទិន្នន័យស្តុកជាក់ស្តែងស្របតាមលក្ខខណ្ឌស្វែងរកឡើយ
                </div>
              ) : (
                <table className="w-full text-xs text-left">
                  <thead>
                    <tr className="bg-slate-100 text-slate-700 font-bold border-b border-slate-200">
                      <th className="py-2.5 px-3 text-center w-10">
                        <input
                          type="checkbox"
                          checked={
                            manageStockRecords.length > 0 &&
                            manageStockRecords.every((r) => selectedManageIds.includes(r.id))
                          }
                          onChange={(e) => {
                            if (e.target.checked) {
                              const allIds = manageStockRecords.map((r) => r.id);
                              setSelectedManageIds(Array.from(new Set([...selectedManageIds, ...allIds])));
                            } else {
                              const currentIds = new Set(manageStockRecords.map((r) => r.id));
                              setSelectedManageIds(selectedManageIds.filter((id) => !currentIds.has(id)));
                            }
                          }}
                          className="rounded text-rose-600 focus:ring-rose-500 cursor-pointer"
                        />
                      </th>
                      <th className="py-2.5 px-2 text-center w-12">ល.រ</th>
                      <th className="py-2.5 px-2 text-center w-16">ប្រភេទ</th>
                      <th className="py-2.5 px-2 text-center">កាលបរិច្ឆេទ</th>
                      <th className="py-2.5 px-2 text-center">ប្រតិបត្តិការ</th>
                      <th className="py-2.5 px-2 text-center">ប្រភព / ក្រុម</th>
                      <th className="py-2.5 px-2 text-center">លេខចាប់ផ្តើម</th>
                      <th className="py-2.5 px-2 text-center">លេខបញ្ចប់</th>
                      <th className="py-2.5 px-2 text-right">សន្លឹក</th>
                      <th className="py-2.5 px-2 text-center w-20">សកម្មភាព</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-200">
                    {manageStockRecords.map((rec, index) => {
                      const isSelected = selectedManageIds.includes(rec.id);
                      const isIssue =
                        rec.operationType === 'useTeam' ||
                        rec.operationType === 'issueTeam' ||
                        rec.sourceFrom?.includes('បើកផ្តល់');
                      const sheets = rec.totalSheets || (rec.quantityBundles ? rec.quantityBundles * 50 : 0);

                      return (
                        <tr
                          key={rec.id}
                          className={`hover:bg-rose-50/40 transition ${isSelected ? 'bg-rose-50/60' : ''}`}
                        >
                          <td className="py-2 px-3 text-center">
                            <input
                              type="checkbox"
                              checked={isSelected}
                              onChange={(e) => {
                                if (e.target.checked) {
                                  setSelectedManageIds((prev) => [...prev, rec.id]);
                                } else {
                                  setSelectedManageIds((prev) => prev.filter((id) => id !== rec.id));
                                }
                              }}
                              className="rounded text-rose-600 focus:ring-rose-500 cursor-pointer"
                            />
                          </td>
                          <td className="py-2 px-2 text-center text-slate-500 tabular-nums">{index + 1}</td>
                          <td className="py-2 px-2 text-center">
                            <span className="font-bold text-blue-900 bg-blue-100 px-2 py-0.5 rounded text-[11px]">
                              {rec.visaType || '-'}
                            </span>
                          </td>
                          <td className="py-2 px-2 text-center tabular-nums text-slate-600">{rec.date || '-'}</td>
                          <td className="py-2 px-2 text-center">
                            <span
                              className={`px-2 py-0.5 rounded text-[10px] font-bold ${
                                isIssue
                                  ? 'bg-amber-100 text-amber-800'
                                  : 'bg-emerald-100 text-emerald-800'
                              }`}
                            >
                              {rec.operationType === 'openK1'
                                ? 'បើកពី ក១'
                                : rec.operationType === 'issueTeam' || rec.operationType === 'useTeam'
                                ? 'បើកជូនក្រុម'
                                : rec.operationType === 'returnTeam'
                                ? 'ក្រុមបង្វិល'
                                : rec.operationType === 'oldStockK2' || rec.operationType === 'oldStock'
                                ? 'ស្តុកចាស់ K2'
                                : rec.operationType || 'បើកចូល'}
                            </span>
                          </td>
                          <td className="py-2 px-2 text-center text-slate-700">
                            {rec.teamName || rec.teamId || rec.sourceFrom || '-'}
                          </td>
                          <td className="py-2 px-2 text-center font-bold text-slate-900 tabular-nums">
                            {rec.startSerial || '-'}
                          </td>
                          <td className="py-2 px-2 text-center font-bold text-slate-900 tabular-nums">
                            {rec.endSerial || '-'}
                          </td>
                          <td className="py-2 px-2 text-right font-black text-slate-900 tabular-nums">
                            {sheets.toLocaleString()}
                          </td>
                          <td className="py-2 px-2 text-center">
                            <button
                              onClick={() => setDeletingRecord(rec)}
                              title="លុបកំណត់ត្រានេះចេញពីស្តុកជាក់ស្តែង (មិនលុបក្នុងទិន្នន័យសន្លឹកទិដ្ឋាការ)"
                              className="px-2 py-1 bg-rose-50 text-rose-700 border border-rose-200 hover:bg-rose-600 hover:text-white rounded-lg text-[10px] font-bold transition inline-flex items-center gap-1 cursor-pointer shadow-2xs"
                            >
                              <Trash2 className="w-3 h-3" />
                              <span>លុប</span>
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              )}
            </div>

            {/* Footer */}
            <div className="p-4 bg-slate-50 border-t border-slate-200 flex flex-wrap items-center justify-between gap-3">
              <div className="text-xs text-slate-500">
                បង្ហាញ {manageStockRecords.length} / {stockRecords.filter((r) => r.stockType === 'sticker').length} កំណត់ត្រាស្តុកជាក់ស្តែង
              </div>
              <div className="flex items-center gap-2">
                {onClearAllActualStock && (
                  <button
                    onClick={() => setIsClearAllConfirmOpen(true)}
                    className="px-3 py-2 bg-rose-100 hover:bg-rose-200 text-rose-800 rounded-xl text-xs font-bold transition cursor-pointer flex items-center gap-1"
                  >
                    <Trash2 className="w-3.5 h-3.5 text-rose-600" />
                    <span>លុបទិន្នន័យទាំងអស់ (សម្អាត)</span>
                  </button>
                )}
                <button
                  onClick={() => setIsManageDeleteModalOpen(false)}
                  className="px-4 py-2 bg-slate-200 hover:bg-slate-300 text-slate-800 rounded-xl text-xs font-bold transition cursor-pointer"
                >
                  បិទផ្ទាំង
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Confirmation Modal for CLEARING / DELETING ALL ACTUAL STOCK (លុបទិន្នន័យតារាងស្តុកជាក់ស្តែងទាំងអស់) */}
      {isClearAllConfirmOpen && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 animate-fade">
          <div className="bg-white rounded-2xl max-w-md w-full p-6 shadow-2xl border border-rose-200 space-y-4">
            <div className="flex items-center gap-3 text-rose-600">
              <div className="p-3 bg-rose-100 rounded-xl">
                <Trash2 className="w-6 h-6" />
              </div>
              <div>
                <h3 className="text-base font-bold text-gray-900">លុបទិន្នន័យក្នុងតារាងស្តុកជាក់ស្តែងទាំងអស់?</h3>
                <p className="text-xs text-gray-500">សម្អាតទិន្នន័យមិនត្រឹមត្រូវ ដើម្បី Import ឯកសារថ្មី</p>
              </div>
            </div>

            <div className="bg-amber-50 border border-amber-200 rounded-xl p-3 text-xs text-amber-900 space-y-1">
              <div className="font-bold flex items-center gap-1.5 text-amber-800">
                <AlertTriangle className="w-4 h-4" />
                <span>ចំណាំអំពីទិន្នន័យ៖</span>
              </div>
              <p className="text-[11px] leading-relaxed">
                • រាល់ទិន្នន័យក្នុង <b>តារាងស្តុកជាក់ស្តែងសម្រាប់បើកផ្តល់តាមក្រុម (K2)</b> នឹងត្រូវកំណត់ជា <b>០ សន្លឹក</b> ទាំងអស់។<br />
                • <b>ទិន្នន័យក្នុង «ទិន្នន័យសន្លឹកទិដ្ឋាការ» (Stock Data)</b> នឹងនៅតែរក្សាទុកដដែល មិនបាត់បង់ឡើយ។<br />
                • បន្ទាប់ពីលុបរួច លោកអ្នកអាចចុច <b>«Import Excel ស្តុក»</b> ដើម្បីបញ្ចូលទិន្នន័យស្តុកត្រឹមត្រូវឡើងវិញបានភ្លាមៗ។
              </p>
            </div>

            <div className="flex items-center justify-end gap-2 pt-2">
              <button
                onClick={() => setIsClearAllConfirmOpen(false)}
                className="px-4 py-2 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded-xl text-xs font-bold transition cursor-pointer"
              >
                បោះបង់
              </button>
              <button
                onClick={() => {
                  onClearAllActualStock?.();
                  setIsClearAllConfirmOpen(false);
                  setIsManageDeleteModalOpen(false);
                  setSelectedManageIds([]);
                }}
                className="px-4 py-2 bg-rose-600 hover:bg-rose-700 text-white rounded-xl text-xs font-bold transition cursor-pointer shadow-xs flex items-center gap-1.5"
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>យល់ព្រមលុបទិន្នន័យទាំងអស់</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Confirmation Modal for Single Record Deletion from Actual Stock */}
      {deletingRecord && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 animate-fade">
          <div className="bg-white rounded-2xl max-w-md w-full p-6 shadow-2xl border border-rose-100 space-y-4">
            <div className="flex items-center gap-3 text-rose-600">
              <div className="p-3 bg-rose-50 rounded-xl">
                <Trash2 className="w-6 h-6" />
              </div>
              <div>
                <h3 className="text-base font-bold text-gray-900">លុបទិន្នន័យចេញពីស្តុកជាក់ស្តែង?</h3>
                <p className="text-xs text-gray-500">លុបតែក្នុងស្តុក K2 មិនប៉ះពាល់ទិន្នន័យសន្លឹកទិដ្ឋាការ</p>
              </div>
            </div>

            <div className="bg-slate-50 p-3.5 rounded-xl border border-slate-200 text-xs space-y-2">
              <div className="flex justify-between">
                <span className="text-slate-500">ប្រភេទទិដ្ឋាការ:</span>
                <span className="font-bold text-slate-900">{deletingRecord.visaType || '-'}</span>
              </div>
              <div className="flex justify-between">
                <span className="text-slate-500">លេខស៊េរី:</span>
                <span className="font-mono font-bold text-slate-900">
                  {deletingRecord.startSerial || '-'} ដល់ {deletingRecord.endSerial || '-'}
                </span>
              </div>
              <div className="flex justify-between">
                <span className="text-slate-500">ចំនួនសន្លឹក:</span>
                <span className="font-bold text-rose-700">
                  {(deletingRecord.totalSheets || (deletingRecord.quantityBundles ? deletingRecord.quantityBundles * 50 : 0)).toLocaleString()} សន្លឹក
                </span>
              </div>
              {deletingRecord.date && (
                <div className="flex justify-between">
                  <span className="text-slate-500">កាលបរិច្ឆេទ:</span>
                  <span className="text-slate-700">{deletingRecord.date}</span>
                </div>
              )}
            </div>

            <p className="text-xs text-slate-600 leading-relaxed">
              ទិន្នន័យនេះនឹងត្រូវដកចេញពី <b>ស្តុកជាក់ស្តែងសម្រាប់បើកផ្តល់តាមក្រុម (K2)</b> ដើម្បីកែតម្រូវចំនួនសន្លឹកនៅសល់។ <b>ទិន្នន័យក្នុង «ទិន្នន័យសន្លឹកទិដ្ឋាការ» (Stock Data) នឹងនៅរក្សាទុកដដែល</b> ដោយមិនបាត់បង់ឡើយ។
            </p>

            <div className="flex items-center justify-end gap-2 pt-2">
              <button
                onClick={() => setDeletingRecord(null)}
                className="px-4 py-2 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded-xl text-xs font-bold transition cursor-pointer"
              >
                បោះបង់
              </button>
              <button
                onClick={() => handleDeleteRecordFromActualStock(deletingRecord.id)}
                className="px-4 py-2 bg-rose-600 hover:bg-rose-700 text-white rounded-xl text-xs font-bold transition cursor-pointer shadow-xs"
              >
                យល់ព្រមលុបពីស្តុកជាក់ស្តែង
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Confirmation Modal for Range Deletion from Actual Stock */}
      {deletingRange && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 animate-fade">
          <div className="bg-white rounded-2xl max-w-md w-full p-6 shadow-2xl border border-rose-100 space-y-4">
            <div className="flex items-center gap-3 text-rose-600">
              <div className="p-3 bg-rose-50 rounded-xl">
                <Trash2 className="w-6 h-6" />
              </div>
              <div>
                <h3 className="text-base font-bold text-gray-900">លុបចន្លោះលេខចេញពីស្តុកជាក់ស្តែង?</h3>
                <p className="text-xs text-gray-500">លុបតែក្នុងស្តុក K2 មិនប៉ះពាល់ទិន្នន័យសន្លឹកទិដ្ឋាការ</p>
              </div>
            </div>

            <div className="bg-slate-50 p-3.5 rounded-xl border border-slate-200 text-xs space-y-2">
              <div className="flex justify-between">
                <span className="text-slate-500">ប្រភេទទិដ្ឋាការ:</span>
                <span className="font-bold text-slate-900">{deletingRange.visaType}</span>
              </div>
              <div className="flex justify-between">
                <span className="text-slate-500">ចន្លោះលេខស៊េរី:</span>
                <span className="font-mono font-bold text-slate-900">
                  {deletingRange.startSerial} ដល់ {deletingRange.endSerial}
                </span>
              </div>
              <div className="flex justify-between">
                <span className="text-slate-500">ចំនួនសរុប:</span>
                <span className="font-bold text-rose-700">
                  {Math.floor(deletingRange.count / 50)} ក្បាល ({deletingRange.count.toLocaleString()} សន្លឹក)
                </span>
              </div>
            </div>

            <p className="text-xs text-slate-600 leading-relaxed">
              ចន្លោះលេខស៊េរីនេះនឹងត្រូវដកចេញពី <b>ស្តុកជាក់ស្តែងសម្រាប់បើកផ្តល់តាមក្រុម (K2)</b>។ <b>ទិន្នន័យក្នុង «ទិន្នន័យសន្លឹកទិដ្ឋាការ» នឹងនៅរក្សាទុកដដែល</b>។
            </p>

            <div className="flex items-center justify-end gap-2 pt-2">
              <button
                onClick={() => setDeletingRange(null)}
                className="px-4 py-2 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded-xl text-xs font-bold transition cursor-pointer"
              >
                បោះបង់
              </button>
              <button
                onClick={() => handleDeleteRangeFromActualStock(deletingRange)}
                className="px-4 py-2 bg-rose-600 hover:bg-rose-700 text-white rounded-xl text-xs font-bold transition cursor-pointer shadow-xs"
              >
                យល់ព្រមលុបពីស្តុកជាក់ស្តែង
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Confirmation Modal for Batch Deletion from Actual Stock */}
      {isBatchDeleteConfirmOpen && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 animate-fade">
          <div className="bg-white rounded-2xl max-w-md w-full p-6 shadow-2xl border border-rose-100 space-y-4">
            <div className="flex items-center gap-3 text-rose-600">
              <div className="p-3 bg-rose-50 rounded-xl">
                <Trash2 className="w-6 h-6" />
              </div>
              <div>
                <h3 className="text-base font-bold text-gray-900">លុបទិន្នន័យដែលបានជ្រើសរើស?</h3>
                <p className="text-xs text-gray-500">លុប {selectedManageIds.length} កំណត់ត្រាចេញពីស្តុកជាក់ស្តែង</p>
              </div>
            </div>

            <p className="text-xs text-slate-600 leading-relaxed">
              ទិន្នន័យចំនួន <b>{selectedManageIds.length} កំណត់ត្រា</b> នឹងត្រូវដកចេញពី <b>ស្តុកជាក់ស្តែងសម្រាប់បើកផ្តល់តាមក្រុម (K2)</b>។ រាល់ទិន្នន័យក្នុង «ទិន្នន័យសន្លឹកទិដ្ឋាការ» (Stock Data) នៅតែរក្សាទុកដដែល។
            </p>

            <div className="flex items-center justify-end gap-2 pt-2">
              <button
                onClick={() => setIsBatchDeleteConfirmOpen(false)}
                className="px-4 py-2 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded-xl text-xs font-bold transition cursor-pointer"
              >
                បោះបង់
              </button>
              <button
                onClick={() => handleBatchDeleteFromActualStock(selectedManageIds)}
                className="px-4 py-2 bg-rose-600 hover:bg-rose-700 text-white rounded-xl text-xs font-bold transition cursor-pointer shadow-xs"
              >
                យល់ព្រមលុប ({selectedManageIds.length}) ពីស្តុកជាក់ស្តែង
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* RETURN ISSUED DATA MODAL (បង្វិលទិន្នន័យបើកផ្តល់ត្រឡប់ចូលស្តុកជាក់ស្តែង) */}
      {/* ========================================================================= */}
      {isReturnModalOpen && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 animate-fade">
          <div className="bg-white rounded-2xl max-w-3xl w-full shadow-2xl border border-slate-200 overflow-hidden flex flex-col max-h-[90vh]">
            {/* Modal Header */}
            <div className="bg-gradient-to-r from-amber-600 to-amber-700 p-4 text-white flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <div className="p-2 bg-white/10 rounded-xl">
                  <RotateCcw className="w-5 h-5 text-amber-200" />
                </div>
                <div>
                  <h3 className="text-sm font-bold">បង្វិលទិន្នន័យបើកផ្តល់ ត្រឡប់ចូលស្តុកជាក់ស្តែង</h3>
                  <p className="text-[11px] text-amber-100 mt-0.5">
                    ជ្រើសរើសទិន្នន័យដែលបានបើកផ្តល់សាកល្បង ដើម្បីបង្វិលចូលក្នុងស្តុកជាក់ស្តែង និងកែប្រែទិន្នន័យ (cEA) ឡើងវិញ
                  </p>
                </div>
              </div>
              <button
                onClick={() => setIsReturnModalOpen(false)}
                className="p-1.5 text-white/80 hover:text-white hover:bg-white/10 rounded-lg transition"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Controls Bar */}
            <div className="p-4 bg-slate-50 border-b border-slate-200 flex flex-wrap items-center justify-between gap-3">
              <div className="flex items-center gap-2">
                <span className="text-xs font-bold text-slate-700">តម្រងតាមប្រភេទ:</span>
                <select
                  value={returnVisaTypeFilter}
                  onChange={(e) => setReturnVisaTypeFilter(e.target.value)}
                  className="px-3 py-1.5 bg-white border border-slate-300 rounded-lg text-xs font-bold text-slate-800 focus:ring-2 focus:ring-amber-500 outline-none"
                >
                  <option value="ALL">គ្រប់ប្រភេទ ({issuedRecords.length})</option>
                  {allVisaTypes.map((vt) => {
                    const count = issuedRecords.filter(
                      (r) => (r.visaType || '').trim().toUpperCase() === vt.trim().toUpperCase()
                    ).length;
                    return (
                      <option key={vt} value={vt}>
                        ប្រភេទ {vt} ({count})
                      </option>
                    );
                  })}
                </select>
              </div>

              <div className="flex items-center gap-2">
                {returnVisaTypeFilter !== 'ALL' && filteredIssuedRecords.length > 0 && (
                  <button
                    onClick={() => setReturnTypeConfirm(returnVisaTypeFilter)}
                    className="px-3 py-1.5 bg-amber-600 hover:bg-amber-700 text-white rounded-lg text-xs font-bold transition flex items-center gap-1.5 cursor-pointer shadow-xs"
                  >
                    <RotateCcw className="w-3.5 h-3.5" />
                    <span>បង្វិលប្រភេទ {returnVisaTypeFilter} ទាំងអស់</span>
                  </button>
                )}

                {issuedRecords.length > 0 && (
                  <button
                    onClick={() => setIsReturnAllConfirmOpen(true)}
                    className="px-3 py-1.5 bg-rose-600 hover:bg-rose-700 text-white rounded-lg text-xs font-bold transition flex items-center gap-1.5 cursor-pointer shadow-xs"
                  >
                    <Trash2 className="w-3.5 h-3.5" />
                    <span>បង្វិលទិន្នន័យទាំងអស់ ({issuedRecords.length})</span>
                  </button>
                )}
              </div>
            </div>

            {/* Issued Records List */}
            <div className="p-4 overflow-y-auto max-h-[60vh]">
              {filteredIssuedRecords.length === 0 ? (
                <div className="text-center py-10 text-slate-400 space-y-2">
                  <RotateCcw className="w-10 h-10 mx-auto text-slate-300 opacity-50" />
                  <p className="text-xs font-medium">គ្មានទិន្នន័យបើកផ្តល់ដែលត្រូវបង្វិលនោះឡើយ</p>
                </div>
              ) : (
                <div className="border border-slate-200 rounded-xl overflow-hidden shadow-2xs">
                  <table className="w-full text-xs text-left">
                    <thead>
                      <tr className="bg-slate-100 text-slate-700 font-bold border-b border-slate-200">
                        <th className="py-2.5 px-3 text-center w-12">ល.រ</th>
                        <th className="py-2.5 px-3 text-center w-20">ប្រភេទ</th>
                        <th className="py-2.5 px-3 text-center">ថ្ងៃខែ/ម៉ោង</th>
                        <th className="py-2.5 px-3 text-center">ក្រុមទទួល</th>
                        <th className="py-2.5 px-3 text-center">ចន្លោះលេខស៊េរី</th>
                        <th className="py-2.5 px-3 text-right">ចំនួនសន្លឹក</th>
                        <th className="py-2.5 px-3 text-center w-28">សកម្មភាព</th>
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-slate-200">
                      {filteredIssuedRecords.map((rec, idx) => {
                        const sheets = rec.totalSheets || (rec.quantityBundles ? rec.quantityBundles * 50 : 0);
                        return (
                          <tr key={rec.id} className="hover:bg-amber-50/40 transition">
                            <td className="py-2 px-3 text-center text-slate-500 font-bold tabular-nums">
                              {idx + 1}
                            </td>
                            <td className="py-2 px-3 text-center">
                              <span className="inline-block px-2 py-0.5 rounded bg-blue-100 text-blue-900 font-black text-[11px]">
                                {rec.visaType || '-'}
                              </span>
                            </td>
                            <td className="py-2 px-3 text-center text-slate-600 tabular-nums">
                              {rec.date} {rec.time ? `(${rec.time})` : ''}
                            </td>
                            <td className="py-2 px-3 text-center font-semibold text-slate-800">
                              {rec.visaTeamRobokName || '-'}
                            </td>
                            <td className="py-2 px-3 text-center font-mono text-[11px] font-bold text-slate-700 tabular-nums">
                              {rec.startSerial} - {rec.endSerial}
                            </td>
                            <td className="py-2 px-3 text-right font-black text-amber-900 tabular-nums">
                              {sheets.toLocaleString()} សន្លឹក
                            </td>
                            <td className="py-2 px-3 text-center">
                              <button
                                onClick={() => setSelectedRecordToReturn(rec)}
                                className="px-2.5 py-1 bg-amber-500 hover:bg-amber-600 text-white rounded-lg text-[11px] font-bold transition inline-flex items-center gap-1 cursor-pointer shadow-2xs"
                              >
                                <RotateCcw className="w-3 h-3" />
                                <span>បង្វិល</span>
                              </button>
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              )}
            </div>

            {/* Modal Footer */}
            <div className="p-3 bg-slate-100 border-t border-slate-200 flex items-center justify-between text-xs text-slate-600">
              <span>
                សរុបទិន្នន័យបើកផ្តល់: <b>{filteredIssuedRecords.length}</b> កំណត់ត្រា (
                <b>
                  {filteredIssuedRecords
                    .reduce((acc, r) => acc + (r.totalSheets || (r.quantityBundles ? r.quantityBundles * 50 : 0)), 0)
                    .toLocaleString()}
                </b>{' '}
                សន្លឹក)
              </span>
              <button
                onClick={() => setIsReturnModalOpen(false)}
                className="px-4 py-1.5 bg-slate-200 hover:bg-slate-300 text-slate-700 rounded-lg font-bold transition cursor-pointer"
              >
                បិទ
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Confirmation Modal: Return Single Record */}
      {selectedRecordToReturn && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4 z-60 animate-fade">
          <div className="bg-white rounded-2xl max-w-md w-full p-6 shadow-2xl space-y-4">
            <div className="flex items-center gap-3 text-amber-600">
              <div className="p-3 bg-amber-50 rounded-xl">
                <RotateCcw className="w-6 h-6" />
              </div>
              <div>
                <h3 className="text-base font-bold text-gray-900">តើអ្នកចង់បង្វិលទិន្នន័យបើកផ្តល់នេះត្រឡប់ចូលស្តុកវិញមែនទេ?</h3>
                <p className="text-xs text-gray-500">បង្វិលចូលស្តុកជាក់ស្តែង និងទិន្នន័យ (cEA)</p>
              </div>
            </div>

            <div className="bg-amber-50 border border-amber-200 rounded-xl p-3 text-xs text-amber-950 space-y-1">
              <div>
                <b>ប្រភេទទិដ្ឋាការ:</b> <span className="font-bold text-blue-900">{selectedRecordToReturn.visaType}</span>
              </div>
              <div>
                <b>ចំនួនសន្លឹក:</b>{' '}
                <b>
                  {(
                    selectedRecordToReturn.totalSheets ||
                    (selectedRecordToReturn.quantityBundles ? selectedRecordToReturn.quantityBundles * 50 : 0)
                  ).toLocaleString()}{' '}
                  សន្លឹក
                </b>
              </div>
              <div>
                <b>ចន្លោះលេខស៊េរី:</b>{' '}
                <span className="font-mono font-bold">
                  {selectedRecordToReturn.startSerial} ដល់ {selectedRecordToReturn.endSerial}
                </span>
              </div>
              <div>
                <b>ក្រុមទទួល:</b> {selectedRecordToReturn.visaTeamRobokName || '-'}
              </div>
            </div>

            <p className="text-xs text-gray-600 leading-relaxed">
              បន្ទាប់ពីបង្វិលរួច ចំនួនសន្លឹក និងចន្លោះលេខស៊េរីខាងលើនឹងត្រូវ<b>បន្ថែមត្រឡប់ចូលក្នុងស្តុកជាក់ស្តែង (Actual Stock)</b> និងធ្វើបច្ចុប្បន្នភាពទិន្នន័យ (cEA) ឡើងវិញដោយស្វ័យប្រវត្តិ។
            </p>

            <div className="flex items-center justify-end gap-2 pt-2">
              <button
                onClick={() => setSelectedRecordToReturn(null)}
                className="px-4 py-2 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded-lg text-xs font-bold transition cursor-pointer"
              >
                បោះបង់
              </button>
              <button
                onClick={() => handleReturnSingle(selectedRecordToReturn)}
                className="px-4 py-2 bg-amber-600 hover:bg-amber-700 text-white rounded-lg text-xs font-bold transition cursor-pointer"
              >
                យល់ព្រមបង្វិលចូលស្តុកវិញ
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Confirmation Modal: Return by Visa Type */}
      {returnTypeConfirm && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4 z-60 animate-fade">
          <div className="bg-white rounded-2xl max-w-md w-full p-6 shadow-2xl space-y-4">
            <div className="flex items-center gap-3 text-amber-600">
              <div className="p-3 bg-amber-50 rounded-xl">
                <RotateCcw className="w-6 h-6" />
              </div>
              <div>
                <h3 className="text-base font-bold text-gray-900">
                  តើអ្នកចង់បង្វិលទិន្នន័យប្រភេទ {returnTypeConfirm} ទាំងអស់ត្រឡប់ចូលស្តុកវិញមែនទេ?
                </h3>
                <p className="text-xs text-gray-500">បង្វិលរាល់ការបើកផ្តល់សម្រាប់ប្រភេទ {returnTypeConfirm}</p>
              </div>
            </div>

            <p className="text-xs text-gray-600 leading-relaxed">
              រាល់ទិន្នន័យដែលបានបើកផ្តល់សម្រាប់ប្រភេទទិដ្ឋាការ <b>{returnTypeConfirm}</b> នឹងត្រូវដកចេញពីបញ្ជីបើកផ្តល់ ហើយចន្លោះលេខស៊េរីទាំងអស់នឹងត្រូវ<b>បង្វិលចូលក្នុងស្តុកជាក់ស្តែងវិញ</b>ភ្លាមៗ។
            </p>

            <div className="flex items-center justify-end gap-2 pt-2">
              <button
                onClick={() => setReturnTypeConfirm(null)}
                className="px-4 py-2 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded-lg text-xs font-bold transition cursor-pointer"
              >
                បោះបង់
              </button>
              <button
                onClick={() => handleReturnByVisaType(returnTypeConfirm)}
                className="px-4 py-2 bg-amber-600 hover:bg-amber-700 text-white rounded-lg text-xs font-bold transition cursor-pointer"
              >
                យល់ព្រមបង្វិលប្រភេទ {returnTypeConfirm}
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Confirmation Modal: Return All Records */}
      {isReturnAllConfirmOpen && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4 z-60 animate-fade">
          <div className="bg-white rounded-2xl max-w-md w-full p-6 shadow-2xl space-y-4">
            <div className="flex items-center gap-3 text-rose-600">
              <div className="p-3 bg-rose-50 rounded-xl">
                <Trash2 className="w-6 h-6" />
              </div>
              <div>
                <h3 className="text-base font-bold text-gray-900">តើអ្នកចង់បង្វិលរាល់ទិន្នន័យបើកផ្តល់ទាំងអស់មែនទេ?</h3>
                <p className="text-xs text-gray-500">បង្វិលរាល់ការបើកផ្តល់សរុប {issuedRecords.length} កំណត់ត្រា</p>
              </div>
            </div>

            <p className="text-xs text-gray-600 leading-relaxed">
              រាល់ការបើកផ្តល់ទាំងអស់ដែលបានធ្វើឡើង នឹងត្រូវបង្វិលត្រឡប់ចូលក្នុងស្តុកជាក់ស្តែង (Actual Stock) និងទិន្នន័យ (cEA) វិញទាំងអស់។
            </p>

            <div className="flex items-center justify-end gap-2 pt-2">
              <button
                onClick={() => setIsReturnAllConfirmOpen(false)}
                className="px-4 py-2 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded-lg text-xs font-bold transition cursor-pointer"
              >
                បោះបង់
              </button>
              <button
                onClick={() => {
                  const allIds = issuedRecords.map((r) => r.id);
                  handleReturnBatch(allIds);
                }}
                className="px-4 py-2 bg-rose-600 hover:bg-rose-700 text-white rounded-lg text-xs font-bold transition cursor-pointer"
              >
                យល់ព្រមបង្វិលទាំងអស់
              </button>
            </div>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* MODAL: ADD STOCK FROM K1 (បញ្ចូលស្តុកពី ក១ ចូលក្នុងស្តុកជាក់ស្តែង K2) */}
      {/* ========================================================================= */}
      {isAddK1ModalOpen && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4 z-50 animate-fade">
          <div className="bg-white rounded-2xl max-w-xl w-full shadow-2xl border border-slate-200 overflow-hidden flex flex-col max-h-[90vh]">
            {/* Modal Header */}
            <div className="bg-gradient-to-r from-blue-900 via-indigo-900 to-blue-950 p-4 text-white flex items-center justify-between">
              <div className="flex items-center gap-2.5">
                <div className="p-2 bg-white/10 rounded-xl">
                  <PlusCircle className="w-5 h-5 text-blue-300" />
                </div>
                <div>
                  <h3 className="text-sm font-bold">បញ្ចូលស្តុកពី ក១ (សន្លឹកទិដ្ឋាការ)</h3>
                  <p className="text-[11px] text-blue-200 mt-0.5">
                    បញ្ចូលទិន្នន័យទទួលស្តុកពី ក១ ចូលក្នុងតារាងស្តុកជាក់ស្តែងសម្រាប់ទុកបើកផ្តល់តាមក្រុម
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => setIsAddK1ModalOpen(false)}
                className="p-1.5 text-white/80 hover:text-white hover:bg-white/10 rounded-lg transition cursor-pointer"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            {/* Modal Body / Form */}
            <form onSubmit={handleSubmitAddK1Stock} className="p-5 overflow-y-auto space-y-4 flex-1">
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                {/* Visa Type */}
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">
                    ប្រភេទទិដ្ឋាការ <span className="text-rose-500">*</span>
                  </label>
                  <select
                    value={k1VisaType}
                    onChange={(e) => setK1VisaType(e.target.value)}
                    className="w-full px-3 py-2 border border-slate-300 rounded-xl text-xs font-bold text-blue-950 focus:ring-2 focus:ring-blue-500 outline-none bg-white"
                  >
                    {allVisaTypes.map((vt) => (
                      <option key={vt} value={vt}>
                        ទិដ្ឋាការប្រភេទ {vt}
                      </option>
                    ))}
                  </select>
                </div>

                {/* Date */}
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">
                    កាលបរិច្ឆេទទទួល <span className="text-rose-500">*</span>
                  </label>
                  <CustomDatePicker
                    required
                    value={k1Date}
                    onChange={(d) => setK1Date(d)}
                    className="py-1.5 text-xs"
                  />
                </div>

                {/* Time */}
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">ម៉ោងទទួល</label>
                  <input
                    type="time"
                    value={k1Time}
                    onChange={(e) => setK1Time(e.target.value)}
                    className="w-full px-3 py-2 border border-slate-300 rounded-xl text-xs font-medium focus:ring-2 focus:ring-blue-500 outline-none bg-white"
                  />
                </div>

                {/* Source */}
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">ប្រភព / អ្នកប្រគល់</label>
                  <input
                    type="text"
                    value={k1Sender}
                    onChange={(e) => setK1Sender(e.target.value)}
                    placeholder="ក១"
                    className="w-full px-3 py-2 border border-slate-300 rounded-xl text-xs font-medium focus:ring-2 focus:ring-blue-500 outline-none bg-white"
                  />
                </div>

                {/* Bundles (50 sheets/bundle) */}
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">
                    ចំនួនក្បាល (៥០សន្លឹក/ក្បាល)
                  </label>
                  <input
                    type="number"
                    min="0"
                    value={k1QuantityBundles}
                    onChange={(e) => handleK1BundlesChange(e.target.value)}
                    className="w-full px-3 py-2 border border-slate-300 rounded-xl text-xs font-bold text-blue-900 focus:ring-2 focus:ring-blue-500 outline-none bg-white"
                  />
                </div>

                {/* Total Sheets */}
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">
                    ចំនួនសន្លឹកសរុប <span className="text-rose-500">*</span>
                  </label>
                  <input
                    type="number"
                    min="1"
                    value={k1TotalSheets}
                    onChange={(e) => {
                      setK1TotalSheets(e.target.value);
                      const sheets = parseInt(e.target.value, 10) || 0;
                      setK1QuantityBundles(String(Math.floor(sheets / 50)));
                    }}
                    required
                    className="w-full px-3 py-2 border border-slate-300 rounded-xl text-xs font-black text-blue-900 focus:ring-2 focus:ring-blue-500 outline-none bg-white"
                  />
                </div>

                {/* Start Serial */}
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">
                    លេខចាប់ពី (Start Serial)
                  </label>
                  <input
                    type="text"
                    value={k1StartSerial}
                    onChange={(e) => setK1StartSerial(e.target.value.toUpperCase())}
                    placeholder="ឧ. FA0185001"
                    className="w-full px-3 py-2 border border-slate-300 rounded-xl text-xs font-mono font-bold text-slate-900 focus:ring-2 focus:ring-blue-500 outline-none bg-white"
                  />
                </div>

                {/* End Serial */}
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">
                    ដល់លេខ (End Serial - ស្វ័យប្រវត្តិ)
                  </label>
                  <input
                    type="text"
                    value={k1EndSerial}
                    onChange={(e) => setK1EndSerial(e.target.value.toUpperCase())}
                    placeholder="ឧ. FA0185050"
                    className="w-full px-3 py-2 border border-slate-300 rounded-xl text-xs font-mono font-bold text-slate-900 focus:ring-2 focus:ring-blue-500 outline-none bg-slate-50"
                  />
                </div>

                {/* Receiver */}
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">អ្នកទទួល</label>
                  <input
                    type="text"
                    value={k1Receiver}
                    onChange={(e) => setK1Receiver(e.target.value)}
                    placeholder="ឈ្មោះអ្នកទទួល"
                    className="w-full px-3 py-2 border border-slate-300 rounded-xl text-xs font-medium focus:ring-2 focus:ring-blue-500 outline-none bg-white"
                  />
                </div>

                {/* Notes */}
                <div>
                  <label className="block text-xs font-bold text-slate-700 mb-1">កំណត់សម្គាល់</label>
                  <input
                    type="text"
                    value={k1Notes}
                    onChange={(e) => setK1Notes(e.target.value)}
                    placeholder="សម្គាល់ផ្សេងៗ..."
                    className="w-full px-3 py-2 border border-slate-300 rounded-xl text-xs font-medium focus:ring-2 focus:ring-blue-500 outline-none bg-white"
                  />
                </div>
              </div>

              {/* Information Banner */}
              <div className="bg-blue-50 border border-blue-200 rounded-xl p-3 text-xs text-blue-900 flex items-start gap-2.5">
                <Info className="w-4 h-4 text-blue-600 shrink-0 mt-0.5" />
                <div className="space-y-1">
                  <div className="font-bold">បញ្ចូលចូលក្នុងស្តុកជាក់ស្តែង K2 ភ្លាមៗ</div>
                  <p className="text-[11px] text-blue-700">
                    ទិន្នន័យនេះនឹងត្រូវបន្ថែមចូលក្នុងតារាងស្តុកជាក់ស្តែង ដើម្បីត្រៀមសម្រាប់បើកផ្តល់ជូនក្រុម និងចូលក្នុងកំណត់ត្រាស្តុកផ្លូវការដោយស្វ័យប្រវត្តិ។
                  </p>
                </div>
              </div>

              {/* Modal Actions */}
              <div className="pt-2 flex items-center justify-end gap-2 border-t border-slate-200">
                <button
                  type="button"
                  onClick={() => setIsAddK1ModalOpen(false)}
                  className="px-4 py-2 rounded-xl text-xs font-bold text-slate-700 hover:bg-slate-100 transition cursor-pointer"
                >
                  បោះបង់
                </button>
                <button
                  type="submit"
                  className="px-5 py-2 rounded-xl text-xs font-bold bg-blue-600 hover:bg-blue-700 text-white shadow-xs transition flex items-center gap-1.5 cursor-pointer"
                >
                  <Check className="w-4 h-4" />
                  <span>រក្សាទុក និងបញ្ចូលស្តុក</span>
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* ========================================================================= */}
      {/* PRINT-ONLY SECTION */}
      {/* ========================================================================= */}
      <div id="sticker-status-print-document" className="hidden print:block print:p-8 bg-white text-black font-siemreap">
        <div className="text-center mb-6">
          <div className="font-moul text-base text-slate-900">ព្រះរាជាណាចក្រកម្ពុជា</div>
          <div className="font-moul text-base text-slate-900">ជាតិ សាសនា ព្រះមហាក្សត្រ</div>
          <div className="text-xs mt-1">***</div>
          <div className="text-left text-xs mt-2">
            <div className="font-moul text-xs text-slate-900">អគ្គនាយកដ្ឋានអន្តោប្រវេសន៍</div>
            <div className="font-moul text-xs text-slate-900">នាយកដ្ឋានទិដ្ឋាការ និងប័ណ្ណស្នាក់នៅ</div>
          </div>
          <h2 className="font-moul text-sm text-slate-900 mt-4">
            តារាងស្តុកជាក់ស្តែងសម្រាប់ទុកធ្វើការបើកផ្តល់តាមក្រុម
          </h2>
          <div className="text-xs text-slate-600 mt-1 font-siemreap">គិតត្រឹមថ្ងៃទី {asOfDate}</div>
        </div>

        <table className="w-full text-xs border-collapse border border-black mb-6">
          <thead>
            <tr className="bg-slate-100 font-bold">
              <th className="border border-black p-2 text-center w-24">ប្រភេទ</th>
              <th className="border border-black p-2 text-center w-28">ក្បាល</th>
              <th className="border border-black p-2 text-center w-32">សន្លឹក</th>
              <th className="border border-black p-2 text-center">ចាប់ផ្តើម</th>
              <th className="border border-black p-2 text-center">ដល់លេខ</th>
            </tr>
          </thead>
          <tbody>
            {filteredStatusItems.map((item) => {
              const ranges = item.remainingRanges;
              const remainingSheets =
                selectedStockScope === 'office'
                  ? item.officeRemainingSheets
                  : selectedStockScope === 'teams'
                  ? item.teamRemainingSheets
                  : item.totalRemainingSheets;
              const remainingBooks =
                selectedStockScope === 'office'
                  ? item.officeRemainingBooks
                  : selectedStockScope === 'teams'
                  ? item.teamRemainingBooks
                  : item.totalRemainingBooks;

              if (ranges && ranges.length > 0) {
                return ranges.map((range, rangeIdx) => {
                  const rangeSheets = range.count;
                  const rangeBooks = Math.floor(rangeSheets / 50);

                  return (
                    <tr key={`${item.visaType}-${rangeIdx}`}>
                      {rangeIdx === 0 && (
                        <td
                          rowSpan={ranges.length}
                          className="border border-black p-2 text-center font-bold text-sm align-middle"
                        >
                          {item.visaType}
                        </td>
                      )}
                      <td className="border border-black p-2 text-center tabular-nums font-semibold">
                        {rangeBooks.toLocaleString()}
                      </td>
                      <td className="border border-black p-2 text-center tabular-nums font-semibold">
                        {rangeSheets.toLocaleString()}
                      </td>
                      <td className="border border-black p-2 text-center tabular-nums">
                        {range.startSerial}
                      </td>
                      <td className="border border-black p-2 text-center tabular-nums">
                        {range.endSerial}
                      </td>
                    </tr>
                  );
                });
              }

              return (
                <tr key={item.visaType}>
                  <td className="border border-black p-2 text-center font-bold text-sm align-middle">
                    {item.visaType}
                  </td>
                  <td className="border border-black p-2 text-center tabular-nums font-semibold">
                    {remainingBooks.toLocaleString()}
                  </td>
                  <td className="border border-black p-2 text-center tabular-nums font-semibold">
                    {remainingSheets.toLocaleString()}
                  </td>
                  <td className="border border-black p-2 text-center tabular-nums">
                    {item.startSerial !== '-' ? item.startSerial : '-'}
                  </td>
                  <td className="border border-black p-2 text-center tabular-nums">
                    {item.endSerial !== '-' ? item.endSerial : '-'}
                  </td>
                </tr>
              );
            })}
          </tbody>
          <tfoot>
            <tr className="font-bold bg-slate-100">
              <td className="border border-black p-2 text-center font-bold">សរុបរួម</td>
              <td className="border border-black p-2 text-center tabular-nums">{summary.totalBooks.toLocaleString()}</td>
              <td className="border border-black p-2 text-center tabular-nums">{summary.totalSheets.toLocaleString()}</td>
              <td colSpan={2} className="border border-black p-2 text-center font-normal text-slate-600">
                សរុបគ្រប់ប្រភេទ និងគ្រប់ចន្លោះលេខ
              </td>
            </tr>
          </tfoot>
        </table>

        {/* Signatures */}
        <div className="flex justify-between items-start mt-8 pt-4 text-xs">
          <div className="text-center">
            <div>បានឃើញ និងឯកភាព</div>
            <div className="font-bold mt-1">ប្រធានការិយាល័យ</div>
          </div>
          <div className="text-center">
            <div>រាជធានីភ្នំពេញ, ថ្ងៃទី..... ខែ..... ឆ្នាំ២០២...</div>
            <div className="font-bold mt-1">{signerTitle}</div>
            <div className="mt-14 font-bold">{signerRank} {signerName}</div>
          </div>
        </div>
      </div>
    </div>
  );
};
