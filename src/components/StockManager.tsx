import React, { useState, useEffect, useMemo, useDeferredValue } from 'react';
import * as XLSX from 'xlsx';
import { CategoriesState, Officer, StockRecord, UserRole, UserAccount } from '../types';
import { useWorkspaceSettings } from '../context/WorkspaceSettingsContext';
import { StockPdfModal } from './StockPdfModal';
import { StockImportModal } from './VisaImportModal';
import { EVisaStockReport, EVisaRobokReport, EVisaSingleRobokReport } from './EVisaStockReport';
import { StickerOfficeStockReport } from './StickerOfficeStockReport';
import { StickerVisaStatus } from './StickerVisaStatus';
import { DEFAULT_29_TEAMS, VISA_TYPES } from './StickerTeamStockReport';
import { normalizeDateToISO } from '../utils/teamNormalization';
import { printA4Document } from '../utils/printHelper';
import { CustomDatePicker } from './CustomDatePicker';
import {
  Boxes,
  Archive,
  Package,
  Calendar,
  Clock,
  ArrowLeft,
  Trash2,
  Search,
  Printer,
  Layers,
  Award,
  UserCheck,
  FileCheck,
  Check,
  ListFilter,
  User,
  Home,
  CheckCircle2,
  FileText,
  Pencil,
  X,
  FileSpreadsheet,
  Plus,
  RotateCcw,
  Undo2,
  AlertTriangle,
  AlertCircle,
  Info,
  Filter,
  HelpCircle,
  Lock,
} from 'lucide-react';

interface StockManagerProps {
  initialStockType?: 'sticker' | 'evisa';
  initialOperationType?: 'openK1' | 'issueTeam' | 'useTeam';
  hideOperationSelector?: boolean;
  initialViewMode?: 'all' | 'form' | 'data' | 'report' | 'robokReport' | 'singleRobokReport' | 'stickerOfficeReport' | 'stickerStatus' | 'handoverWorkspace';
  categories: CategoriesState;
  officers?: Officer[];
  stockRecords: StockRecord[];
  actualStockRecords?: StockRecord[];
  users?: UserAccount[];
  currentRole: UserRole;
  userName?: string;
  assignedTeam?: string;
  onAddStockRecord: (record: StockRecord) => void;
  onBatchImportStockRecords?: (records: StockRecord[]) => void;
  onUpdateStockRecord: (record: StockRecord) => void;
  onDeleteStockRecord: (id: string) => void;
  onDeleteBatchStockRecords?: (ids: string[]) => void;
  onReturnIssuedStock?: (record: StockRecord) => void;
  onNavigate?: (page: string) => void;
  onShowToast: (msg: string) => void;
}

export const StockManager: React.FC<StockManagerProps> = ({
  initialStockType = 'evisa',
  initialOperationType,
  hideOperationSelector = false,
  initialViewMode = 'all',
  categories,
  officers,
  stockRecords,
  actualStockRecords,
  users = [],
  currentRole,
  userName = '',
  assignedTeam = '',
  onAddStockRecord,
  onBatchImportStockRecords,
  onUpdateStockRecord,
  onDeleteStockRecord,
  onDeleteBatchStockRecords,
  onReturnIssuedStock,
  onNavigate,
  onShowToast,
}) => {
  const isSecondary = currentRole === 'Secondary' || currentRole === 'User (ការិយាល័យ)' || currentRole === 'Admin';
  const { containerWidthMode } = useWorkspaceSettings();
  // Tab state: 'evisa' (ក្រដាសអនុម័តទិដ្ឋាការអេឡិចត្រូនិក) or 'sticker' (សន្លឹកទិដ្ឋាការស្អិត)
  const [stockType, setStockType] = useState<'evisa' | 'sticker'>(initialStockType);
  const [viewMode, setViewMode] = useState<'all' | 'form' | 'data' | 'report' | 'robokReport' | 'singleRobokReport' | 'stickerOfficeReport' | 'stickerStatus' | 'handoverWorkspace'>(initialViewMode);
  const isExpandedWidth = containerWidthMode === 'full' || viewMode === 'data' || stockType === 'sticker';
  const [operationType, setOperationType] = useState<string>(() => {
    if (initialOperationType) return initialOperationType;
    if (!isSecondary) return initialStockType === 'evisa' ? 'useTeam' : 'oldStockTeam';
    return 'openK1';
  });

  // Source of truth for actual stock records (specifically for sticker stock distribution)
  const effectiveActualStockRecords = useMemo(
    () => (actualStockRecords && actualStockRecords.length > 0 ? actualStockRecords : stockRecords),
    [actualStockRecords, stockRecords]
  );

  // Filter groups using the system (those with user accounts in the database)
  const filteredVisaTeamsRobok = useMemo(() => {
    const rawList = categories.visaTeamsRobok || [];
    if (!users || users.length === 0) return rawList;
    
    // Find all team names that have user accounts
    const activeUserTeams = new Set<string>();
    users.forEach((u) => {
      if (u.assignedTeam) {
        activeUserTeams.add(u.assignedTeam.trim().toLowerCase());
      }
    });

    // If there are active user accounts, filter visaTeamsRobok by name match.
    // If no accounts exist yet, return everything as fallback.
    if (activeUserTeams.size === 0) return rawList;

    return rawList.filter((vtr) => {
      const cleanName = vtr.name.trim().toLowerCase();
      // Keep if there's a user account configured for this team, or if it's the currently selected team to avoid broken selection
      return activeUserTeams.has(cleanName) || cleanName.includes('ការិយាល័យ') || cleanName === assignedTeam.trim().toLowerCase();
    });
  }, [categories.visaTeamsRobok, users, assignedTeam]);

  // Dynamically extract custom operations imported or existing in records (e.g. any custom operation with 'ក២' or other tags)
  const customOperations = useMemo(() => {
    const standard = ['openK1', 'issueTeam', 'useTeam', 'testPrintK2', 'damaged', 'returnTeam', 'damagedTeam', 'missingTeam'];
    const customSet = new Set<string>();
    stockRecords.forEach((r) => {
      if (r.stockType === stockType && r.operationType && !standard.includes(r.operationType)) {
        customSet.add(r.operationType);
      }
    });
    return Array.from(customSet);
  }, [stockRecords, stockType]);

  useEffect(() => {
    setStockType(initialStockType);
  }, [initialStockType]);

  useEffect(() => {
    if (initialViewMode) {
      setViewMode(initialViewMode);
    }
  }, [initialViewMode]);

  useEffect(() => {
    if (initialOperationType) {
      setOperationType(initialOperationType);
    }
  }, [initialOperationType]);

  useEffect(() => {
    if (!isSecondary) {
      if (stockType === 'evisa') {
        if (operationType !== 'useTeam') {
          setOperationType('useTeam');
          setSourceFrom('ប្រើប្រាស់តាមក្រុម');
        }
      } else if (stockType === 'sticker') {
        if (['openK1', 'oldStockK2', 'damaged', 'testPrintK2', 'issueTeam', 'useTeam'].includes(operationType)) {
          setOperationType('oldStockTeam');
          setSourceFrom('ស្តុកចាស់ក្រុម');
        }
      }
    } else if (stockType === 'sticker' && operationType === 'useTeam') {
      setOperationType('openK1');
      setSourceFrom('ក១');
    }
  }, [isSecondary, stockType, operationType]);

  // Editing state
  const [editingRecordId, setEditingRecordId] = useState<string | null>(null);
  const [highlightedRecordIds, setHighlightedRecordIds] = useState<string[]>([]);

  useEffect(() => {
    if (highlightedRecordIds.length > 0 && (viewMode === 'data' || viewMode === 'all')) {
      const scrollTimer = setTimeout(() => {
        const firstId = highlightedRecordIds[0];
        const rowEl = document.getElementById(`stock-row-${firstId}`);
        if (rowEl) {
          rowEl.scrollIntoView({ behavior: 'smooth', block: 'center' });
        }
      }, 150);

      const clearTimer = setTimeout(() => {
        setHighlightedRecordIds([]);
      }, 6000);

      return () => {
        clearTimeout(scrollTimer);
        clearTimeout(clearTimer);
      };
    }
  }, [highlightedRecordIds, viewMode]);

  // PDF Preview Modal State
  const [selectedPdfRecord, setSelectedPdfRecord] = useState<StockRecord | null>(null);

  // Delete Confirmation Modal State
  const [deletingRecord, setDeletingRecord] = useState<StockRecord | null>(null);
  const [isDeleteAllModalOpen, setIsDeleteAllModalOpen] = useState<boolean>(false);

  // Operation Type dropdown: 'openK1' (បញ្ចូលស្តុកក២) or 'issueTeam' (បើកផ្តល់តាមក្រុម)

  // Form Fields State
  const [date, setDate] = useState<string>(new Date().toISOString().split('T')[0]);
  const [time, setTime] = useState<string>(
    new Date().toLocaleTimeString('en-US', { hour12: false, hour: '2-digit', minute: '2-digit' })
  );
  const [sourceFrom, setSourceFrom] = useState<string>(''); // បើកពី: 'ក១' ឬ 'ន៨'
  const [quantityBundles, setQuantityBundles] = useState<string>('');
  const [visaTeamRobokId, setVisaTeamRobokId] = useState<string>(() => {
    if (!isSecondary && assignedTeam) {
      const matched = (categories?.visaTeamsRobok || []).find(
        (t) => t.name.trim().toLowerCase() === assignedTeam.trim().toLowerCase() || t.id === assignedTeam
      );
      return matched ? matched.id : assignedTeam;
    }
    return '';
  });

  useEffect(() => {
    if (!isSecondary && assignedTeam) {
      const matched = (categories?.visaTeamsRobok || []).find(
        (t) => t.name.trim().toLowerCase() === assignedTeam.trim().toLowerCase() || t.id === assignedTeam
      );
      setVisaTeamRobokId(matched ? matched.id : assignedTeam);
    }
  }, [isSecondary, assignedTeam, categories?.visaTeamsRobok]);

  // Find selected team name if visaTeamRobokId is selected
  const selectedTeamObj = (categories?.visaTeamsRobok || []).find(
    (vtr) => vtr.id === visaTeamRobokId || vtr.name === visaTeamRobokId
  );
  const selectedTeamName = selectedTeamObj ? selectedTeamObj.name : visaTeamRobokId || '';
  const [visaType, setVisaType] = useState<string>('');
  const [requestedRankId, setRequestedRankId] = useState<string>('');
  const [requesterName, setRequesterName] = useState<string>('');
  const [collectorName, setCollectorName] = useState<string>('');
  const [collectorRoleId, setCollectorRoleId] = useState<string>('');
  const [startSerial, setStartSerial] = useState<string>('');
  const [endSerial, setEndSerial] = useState<string>('');
  const [totalSheets, setTotalSheets] = useState<string>('');

  // States for transfer Team (ផ្ទេរការប្រើប្រាស់ក្រុម)
  const [recipientTeamId, setRecipientTeamId] = useState<string>('');
  const [officeApproved, setOfficeApproved] = useState<boolean>(false);
  const [deptApprovalDate, setDeptApprovalDate] = useState<string>('');

  // List of all available recipient teams for transfer operations (all teams in the system, minus the current sender team)
  const availableRecipientTeams = useMemo(() => {
    let list = categories?.visaTeamsRobok && categories.visaTeamsRobok.length > 0
      ? categories.visaTeamsRobok
      : (categories?.visaTeams && categories.visaTeams.length > 0
          ? categories.visaTeams
          : DEFAULT_29_TEAMS.map((name, i) => ({ id: `vtr-def-${i + 1}`, name, createdAt: '' })));

    const currentSenderId = (visaTeamRobokId || '').trim().toLowerCase();
    const currentSenderName = (assignedTeam || selectedTeamName || '').trim().toLowerCase();

    return list.filter((vtr) => {
      const vtrIdClean = (vtr.id || '').trim().toLowerCase();
      const vtrNameClean = (vtr.name || '').trim().toLowerCase();

      // Exclude current sender team ID or Name
      if (currentSenderId && (vtrIdClean === currentSenderId || vtrNameClean === currentSenderId)) {
        return false;
      }
      if (currentSenderName && vtrNameClean === currentSenderName) {
        return false;
      }
      return true;
    });
  }, [categories?.visaTeamsRobok, categories?.visaTeams, visaTeamRobokId, assignedTeam, selectedTeamName]);

  // Sticker multi-item state
  const VISA_TYPE_OPTIONS = ['T', 'T1', 'T2', 'T3', 'E', 'E1', 'E2', 'E3', 'D', 'K', 'A', 'B', 'C'];
  const [stickerItems, setStickerItems] = useState<Array<{ id: string; visaType: string; quantity: string; startSerial: string; endSerial: string }>>([
    { id: '1', visaType: '', quantity: '', startSerial: '', endSerial: '' },
  ]);
  const [manualSerialMap, setManualSerialMap] = useState<Record<string, boolean>>({});

  const toggleManualSerial = (id: string) => {
    setManualSerialMap((prev) => ({ ...prev, [id]: !prev[id] }));
  };

  const addStickerItem = () => {
    setStickerItems((prev) => [
      ...prev,
      { id: Date.now().toString() + Math.random().toString(36).substring(2, 5), visaType: '', quantity: '', startSerial: '', endSerial: '' },
    ]);
  };

  const removeStickerItem = (id: string) => {
    if (stickerItems.length > 1) {
      setStickerItems((prev) => prev.filter((item) => item.id !== id));
      setManualSerialMap((prev) => {
        const next = { ...prev };
        delete next[id];
        return next;
      });
    }
  };

  const updateStickerItem = (id: string, field: 'visaType' | 'quantity' | 'startSerial' | 'endSerial', value: string) => {
    setStickerItems((prev) =>
      prev.map((item) => (item.id === id ? { ...item, [field]: value } : item))
    );
  };

  const updateStickerItemFields = (
    id: string,
    updates: Partial<{ visaType: string; quantity: string; startSerial: string; endSerial: string }>
  ) => {
    setStickerItems((prev) =>
      prev.map((item) => (item.id === id ? { ...item, ...updates } : item))
    );
  };

  // Helper: Parse serial number into prefix, numeric bigint value, and pad length
  const parseSerialNumber = (serial: string): { prefix: string; num: bigint; padLength: number; raw: string } | null => {
    if (!serial || typeof serial !== 'string') return null;
    const clean = serial.trim();
    if (!clean) return null;

    const match = clean.match(/^([^\d]*)(\d+)$/);
    if (match) {
      const prefix = match[1] || '';
      const digitsStr = match[2];
      try {
        const num = BigInt(digitsStr);
        return { prefix, num, padLength: digitsStr.length, raw: clean };
      } catch {
        return null;
      }
    }

    const digitsOnly = clean.replace(/\D/g, '');
    if (digitsOnly) {
      try {
        const num = BigInt(digitsOnly);
        const prefix = clean.replace(/\d+$/, '');
        return { prefix, num, padLength: digitsOnly.length, raw: clean };
      } catch {
        return null;
      }
    }

    return null;
  };

  // Helper: Format bigint back to formatted serial string
  const formatSerialNumber = (prefix: string, num: bigint, padLength: number): string => {
    const digitsStr = num.toString().padStart(padLength, '0');
    return `${prefix}${digitsStr}`;
  };

  // Exact Interval Subtraction Algorithm
  const subtractIntervals = (
    inflows: Array<{ start: bigint; end: bigint; prefix: string; padLength: number }>,
    outStart: bigint,
    outEnd: bigint
  ): Array<{ start: bigint; end: bigint; prefix: string; padLength: number }> => {
    const result: Array<{ start: bigint; end: bigint; prefix: string; padLength: number }> = [];

    for (const interval of inflows) {
      const { start: inS, end: inE, prefix, padLength } = interval;

      if (outEnd < inS || outStart > inE) {
        result.push(interval);
        continue;
      }

      if (outStart <= inS && outEnd < inE) {
        result.push({
          start: outEnd + 1n,
          end: inE,
          prefix,
          padLength,
        });
        continue;
      }

      if (outStart > inS && outEnd >= inE) {
        result.push({
          start: inS,
          end: outStart - 1n,
          prefix,
          padLength,
        });
        continue;
      }

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
      }
    }

    return result;
  };

  // Calculate available office stock ranges for a visa type
  const getAvailableOfficeRanges = (
    records: StockRecord[],
    vType: string,
    excludeRecordIds: string[] = []
  ): Array<{
    start: bigint;
    end: bigint;
    startSerial: string;
    endSerial: string;
    count: number;
    books: number;
    prefix: string;
    padLength: number;
  }> => {
    if (!vType) return [];

    const rawInflowIntervals: Array<{
      start: bigint;
      end: bigint;
      prefix: string;
      padLength: number;
    }> = [];

    const rawOutflowIntervals: Array<{
      start: bigint;
      end: bigint;
      prefix: string;
      padLength: number;
    }> = [];

    const upperVType = vType.trim().toUpperCase();

    // Check if actualStockRecords has entries for this visa type
    const actualHasType = (actualStockRecords || []).some(
      (r) => r.stockType === 'sticker' && (r.visaType || '').trim().toUpperCase() === upperVType
    );

    // If actualStockRecords has entries for this visa type, use it as the source of truth;
    // otherwise fallback to records or stockRecords.
    const pool = actualHasType
      ? actualStockRecords!
      : records && records.length > 0
      ? records
      : stockRecords;

    pool.forEach((r) => {
      if (
        r.stockType === 'sticker' &&
        (r.visaType || '').trim().toUpperCase() === upperVType &&
        r.startSerial
      ) {
        if (excludeRecordIds.includes(r.id)) return;

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

        const isReturnK1 =
          r.operationType === 'returnK1' ||
          r.sourceFrom === 'បង្វិលក១' ||
          r.sourceFrom === 'ទិដ្ឋាការបង្វិលទៅក១' ||
          (Boolean(r.sourceFrom?.includes('បង្វិល') && r.sourceFrom?.includes('ក១')) && r.operationType !== 'openK1');

        const isReturn =
          !isReturnK1 &&
          (r.operationType === 'returnTeam' ||
            r.operationType === 'returned' ||
            r.sourceFrom?.includes('បង្វិល') ||
            r.operationType?.includes('បង្វិល'));

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

        const isOfficeInflow = isOpenK1 || isOldStock || isReturn;
        const isOfficeOutflow = isIssue || isDamagedK2 || isTest || isReturnK1;

        const parsedStart = parseSerialNumber(r.startSerial);
        if (parsedStart) {
          let endNum: bigint;
          const parsedEnd = r.endSerial ? parseSerialNumber(r.endSerial) : null;
          const sheets = r.quantityBundles ? r.quantityBundles * 50 : (r.totalSheets || 0);

          if (parsedEnd && parsedEnd.num >= parsedStart.num) {
            endNum = parsedEnd.num;
          } else if (sheets > 0) {
            endNum = parsedStart.num + BigInt(sheets) - 1n;
          } else {
            endNum = parsedStart.num;
          }

          if (isOfficeInflow) {
            rawInflowIntervals.push({
              start: parsedStart.num,
              end: endNum,
              prefix: parsedStart.prefix,
              padLength: parsedStart.padLength,
            });
          } else if (isOfficeOutflow) {
            rawOutflowIntervals.push({
              start: parsedStart.num,
              end: endNum,
              prefix: parsedStart.prefix,
              padLength: parsedStart.padLength,
            });
          }
        }
      }
    });

    rawInflowIntervals.sort((a, b) => (a.start < b.start ? -1 : a.start > b.start ? 1 : 0));

    const mergedInflows: Array<{ start: bigint; end: bigint; prefix: string; padLength: number }> = [];
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

    let activeIntervals = [...mergedInflows];
    for (const out of rawOutflowIntervals) {
      activeIntervals = subtractIntervals(activeIntervals, out.start, out.end);
    }

    return activeIntervals.map((intv) => {
      const count = Number(intv.end - intv.start + 1n);
      return {
        start: intv.start,
        end: intv.end,
        startSerial: formatSerialNumber(intv.prefix, intv.start, intv.padLength),
        endSerial: formatSerialNumber(intv.prefix, intv.end, intv.padLength),
        count,
        books: Math.floor(count / 50),
        prefix: intv.prefix,
        padLength: intv.padLength,
      };
    });
  };

  const calculateEndSerial = (start: string, qtyStr: string) => {
    const qty = parseInt(qtyStr, 10);
    if (!start || isNaN(qty) || qty <= 0) return '';
    const parsed = parseSerialNumber(start);
    if (!parsed) return '';
    const endNum = parsed.num + BigInt(qty) - 1n;
    return formatSerialNumber(parsed.prefix, endNum, parsed.padLength);
  };

  // Search, Date Range, and Operation Type filter
  const [searchTerm, setSearchTerm] = useState<string>('');
  const deferredSearchTerm = useDeferredValue(searchTerm);
  const [filterOperationType, setFilterOperationType] = useState<string>('all');
  const [filterStartDate, setFilterStartDate] = useState<string>('');
  const [filterEndDate, setFilterEndDate] = useState<string>('');
  const [filterUncountedInTeamReport, setFilterUncountedInTeamReport] = useState<boolean>(false);
  const [uncountedCategoryFilter, setUncountedCategoryFilter] = useState<string>('all');
  const [deleteScope, setDeleteScope] = useState<'filtered' | 'all'>('filtered');
  const [tablePage, setTablePage] = useState<number>(1);
  const [tablePageSize, setTablePageSize] = useState<number>(50);

  // Pre-calculate count of records by date to avoid O(N^2) loops in table rendering
  const sameDateCountsMap = useMemo(() => {
    const map: Record<string, number> = {};
    stockRecords.forEach((r) => {
      if (r.stockType === 'sticker' && r.date) {
        map[r.date] = (map[r.date] || 0) + 1;
      }
    });
    return map;
  }, [stockRecords]);

  const isTransferRecord = (r: StockRecord) => {
    const op = (r.operationType || '').toLowerCase();
    return (
      op === 'transferteam' ||
      op === 'transferuseteam' ||
      op === 'transfer' ||
      op.includes('transfer') ||
      op.includes('ផ្ទេរ') ||
      Boolean(r.sourceFrom && r.sourceFrom.includes('ផ្ទេរ')) ||
      Boolean((r as any).notes && (r as any).notes.includes('ផ្ទេរ')) ||
      Boolean(r.remarks && r.remarks.includes('ផ្ទេរ'))
    );
  };

  const isRecordForTeamRaw = (r: StockRecord, team: string) => {
    if (!team) return true;
    const cleanTeam = team.trim().toLowerCase();
    const rTeamName = (r.visaTeamRobokName || '').trim().toLowerCase();
    const rTeamId = (r.visaTeamRobokId || '').trim().toLowerCase();
    if (rTeamName === cleanTeam || rTeamId === cleanTeam) return true;
    if (rTeamName && cleanTeam && (rTeamName.includes(cleanTeam) || cleanTeam.includes(rTeamName))) return true;
    const matched = (categories?.visaTeamsRobok || []).find(
      (t) => t.name.trim().toLowerCase() === cleanTeam || t.id.toLowerCase() === cleanTeam
    );
    if (matched) {
      if (
        r.visaTeamRobokId === matched.id ||
        (r.visaTeamRobokName && r.visaTeamRobokName.trim().toLowerCase() === matched.name.trim().toLowerCase())
      ) {
        return true;
      }
    }
    return false;
  };

  const isRecordForRecipientTeam = (r: StockRecord, team: string) => {
    if (!team) return false;
    const cleanTeam = team.trim().toLowerCase();
    const rTeamName = (r.recipientTeamName || '').trim().toLowerCase();
    const rTeamId = (r.recipientTeamId || '').trim().toLowerCase();
    if (rTeamName === cleanTeam || rTeamId === cleanTeam) return true;
    if (rTeamName && cleanTeam && (rTeamName.includes(cleanTeam) || cleanTeam.includes(rTeamName))) return true;
    const matched = (categories?.visaTeamsRobok || []).find(
      (t) => t.name.trim().toLowerCase() === cleanTeam || t.id.toLowerCase() === cleanTeam
    );
    if (matched) {
      if (
        r.recipientTeamId === matched.id ||
        (r.recipientTeamName && r.recipientTeamName.trim().toLowerCase() === matched.name.trim().toLowerCase())
      ) {
        return true;
      }
    }
    return false;
  };

  const matchOperation = (r: StockRecord, opType: string) => {
    if (opType === 'all') return true;
    if (opType === 'oldStockK2') {
      return (
        r.operationType === 'oldStockK2' ||
        (r.operationType === 'openK1' && (r.sourceFrom === 'ស្តុកចាស់ ក២' || r.sourceFrom === 'ស្តុកចាស់' || r.sourceFrom?.includes('ស្តុកចាស់ ក២') || (r.sourceFrom?.includes('ស្តុកចាស់') && !r.visaTeamRobokName && !r.sourceFrom?.includes('ក្រុម'))))
      );
    }
    if (opType === 'oldStockTeam') {
      return (
        r.operationType === 'oldStockTeam' ||
        ((r.operationType === 'issueTeam' || r.operationType === 'useTeam') && (r.sourceFrom === 'ស្តុកចាស់ក្រុម' || r.sourceFrom === 'ស្តុកចាស់របស់ក្រុម' || r.sourceFrom?.includes('ស្តុកចាស់ក្រុម') || r.sourceFrom?.includes('ស្តុកចាស់របស់ក្រុម'))) ||
        (r.sourceFrom?.includes('ស្តុកចាស់') && Boolean(r.visaTeamRobokName || r.sourceFrom?.includes('ក្រុម')))
      );
    }
    if (opType === 'openK1') {
      return r.operationType === 'openK1' && !r.sourceFrom?.includes('ស្តុកចាស់');
    }
    if (opType === 'issueTeam') {
      const isIssue = (r.operationType === 'issueTeam' || r.operationType === 'issue') && !r.sourceFrom?.includes('ស្តុកចាស់');
      const currentTeam = !isSecondary ? assignedTeam : visaTeamRobokId || selectedTeamName;
      const isTransferIn = r.operationType === 'transferTeam' && currentTeam && isRecordForRecipientTeam(r, currentTeam);
      return isIssue || isTransferIn;
    }
    if (opType === 'transferTeam') {
      const currentTeam = !isSecondary ? assignedTeam : visaTeamRobokId || selectedTeamName;
      return r.operationType === 'transferTeam' && (!currentTeam || isRecordForTeamRaw(r, currentTeam) || isRecordForRecipientTeam(r, currentTeam));
    }
    if (opType === 'returnK1') {
      return (
        r.operationType === 'returnK1' ||
        r.sourceFrom === 'បង្វិលក១' ||
        r.sourceFrom === 'ទិដ្ឋាការបង្វិលទៅក១' ||
        (Boolean(r.sourceFrom?.includes('បង្វិល') && r.sourceFrom?.includes('ក១')) && r.operationType !== 'openK1')
      );
    }
    if (opType === 'returnTeam') {
      if (
        r.operationType === 'returnK1' ||
        r.sourceFrom === 'បង្វិលក១' ||
        r.sourceFrom === 'ទិដ្ឋាការបង្វិលទៅក១' ||
        (Boolean(r.sourceFrom?.includes('បង្វិល') && r.sourceFrom?.includes('ក១')) && r.operationType !== 'openK1')
      ) {
        return false;
      }
      return (
        r.operationType === 'returnTeam' ||
        r.operationType === 'returned' ||
        r.operationType === 'return_from_team' ||
        Boolean(r.sourceFrom?.includes('បង្វិល') || r.operationType?.includes('បង្វិល'))
      );
    }
    if (opType === 'returnStub') {
      return (
        r.operationType === 'returnStub' ||
        Boolean(r.sourceFrom?.includes('គល់សន្លឹក')) ||
        Boolean(r.remarks?.includes('គល់សន្លឹក')) ||
        Boolean(r.id?.startsWith('stock-stub-'))
      );
    }
    if (opType === 'damaged' || opType === 'damagedK2') {
      return (
        r.operationType === 'damaged' ||
        r.operationType === 'damagedK2' ||
        r.operationType === 'invalid' ||
        Boolean(r.sourceFrom?.includes('ខូចក២') || r.sourceFrom === 'ទិដ្ឋាការខូចក២') ||
        Boolean(r.operationType?.includes('ខូច') && !r.operationType?.includes('ក្រុម') && !r.visaTeamRobokName)
      );
    }
    if (opType === 'testPrintK2') {
      return (
        r.operationType === 'testPrintK2' ||
        r.operationType === 'testK2' ||
        Boolean(r.sourceFrom?.includes('សាកក២') || r.sourceFrom === 'ទិដ្ឋាការសាកក២') ||
        Boolean(r.operationType?.includes('សាក') && !r.operationType?.includes('ក្រុម') && !r.visaTeamRobokName)
      );
    }
    if (opType === 'useTeam') {
      return (
        r.operationType === 'useTeam' ||
        r.operationType === 'use_team' ||
        r.operationType === 'use' ||
        r.operationType === 'useSticker' ||
        r.operationType === 'useEVisa' ||
        (Boolean(r.operationType?.includes('ប្រើប្រាស់') || r.sourceFrom?.includes('ប្រើប្រាស់')) &&
          !r.operationType?.includes('ផ្ទេរ') &&
          !r.sourceFrom?.includes('ផ្ទេរ') &&
          !r.sourceFrom?.includes('ស្តុកចាស់'))
      );
    }
    return r.operationType === opType;
  };

  const opLabelMap: Record<string, string> = {
    openK1: 'ការបញ្ចូលស្តុក (ក១)',
    oldStockK2: 'ស្តុកចាស់ ក២ (សន្និធិដើមគ្រា ក២)',
    oldStockTeam: 'ស្តុកចាស់ក្រុម (សន្និធិដើមគ្រាក្រុម)',
    issueTeam: isSecondary ? 'ការបើកផ្តល់តាមក្រុម' : 'ទទួលពីការិយាល័យ ឬផ្ទេរពីក្រុម',
    useTeam: 'ការប្រើប្រាស់តាមក្រុម',
    testPrintK2: 'ទិដ្ឋាការសាកក២',
    damaged: 'ទិដ្ឋាការខូចក២',
    damagedTeam: 'ទិដ្ឋាការខូចក្រុម',
    missingTeam: 'ទិដ្ឋាការខ្វះក្រុម',
    returnTeam: isSecondary ? 'ទិដ្ឋាការបង្វិលពីក្រុម' : 'ទិដ្ឋាការបង្វិលទៅក២',
    returnK1: 'បង្វិលក១',
    transferTeam: 'ផ្ទេរការប្រើប្រាស់ក្រុម',
    returnStub: 'ប្រមូលគល់សន្លឹកទិដ្ឋាការ',
  };

  const opShortLabelMap: Record<string, string> = {
    openK1: 'បញ្ចូលស្តុក (ក១)',
    oldStockK2: 'ស្តុកចាស់ ក២',
    oldStockTeam: 'ស្តុកចាស់ក្រុម',
    issueTeam: isSecondary ? 'បើកផ្តល់តាមក្រុម' : 'ទទួលពីការិយាល័យ ឬផ្ទេរពីក្រុម',
    useTeam: 'ប្រើប្រាស់តាមក្រុម',
    testPrintK2: 'ទិដ្ឋាការសាកក២',
    damaged: 'ទិដ្ឋាការខូចក២',
    damagedTeam: 'ទិដ្ឋាការខូចក្រុម',
    missingTeam: 'ទិដ្ឋាការខ្វះក្រុម',
    returnTeam: isSecondary ? 'ទិដ្ឋាការបង្វិលពីក្រុម' : 'ទិដ្ឋាការបង្វិលទៅក២',
    returnK1: 'បង្វិលក១',
    transferTeam: 'ផ្ទេរការប្រើប្រាស់',
    returnStub: 'ប្រមូលគល់សន្លឹក',
  };

  const getRecordOperationLabel = (item: StockRecord): string => {
    if (item.operationType === 'oldStockK2') return 'ស្តុកចាស់ ក២';
    if (item.operationType === 'oldStockTeam') return 'ស្តុកចាស់ក្រុម';
    if (item.operationType === 'testPrintK2') return 'ទិដ្ឋាការសាកក២';
    if (item.operationType === 'damaged') return 'ទិដ្ឋាការខូចក២';
    if (item.operationType === 'damagedTeam') return 'ទិដ្ឋាការខូចក្រុម';
    if (item.operationType === 'missingTeam') return 'ទិដ្ឋាការខ្វះក្រុម';
    if (item.operationType === 'returnK1') return 'បង្វិលក១';
    if (
      item.operationType === 'returnStub' ||
      item.sourceFrom?.includes('គល់សន្លឹក') ||
      item.remarks?.includes('គល់សន្លឹក') ||
      item.id?.startsWith('stock-stub-')
    ) {
      return 'ប្រមូលគល់សន្លឹកទិដ្ឋាការ';
    }
    if (item.operationType === 'returnTeam') {
      return !isSecondary ? 'ទិដ្ឋាការបង្វិលទៅក២' : 'ទិដ្ឋាការបង្វិលពីក្រុម';
    }
    if (item.operationType === 'transferTeam') {
      const currentTeam = targetTeamFilter || assignedTeam || '';
      if (currentTeam && isRecordForRecipientTeam(item, currentTeam)) {
        return 'ទទួលពីការិយាល័យ ឬផ្ទេរពីក្រុម';
      }
      return isSecondary ? 'ផ្ទេរការប្រើប្រាស់ក្រុម' : 'ផ្ទេរការប្រើប្រាស់ក្រុម';
    }
    if (item.operationType === 'issueTeam') {
      return isSecondary ? 'ការបើកផ្តល់តាមក្រុម' : 'ទទួលពីការិយាល័យ ឬផ្ទេរពីក្រុម';
    }
    if (item.operationType === 'openK1') {
      return item.sourceFrom?.includes('ស្តុកចាស់') ? 'ស្តុកចាស់ ក២' : 'ការបញ្ចូលស្តុក (ក១)';
    }
    if (item.operationType === 'useTeam') return 'ការប្រើប្រាស់តាមក្រុម';
    if (item.sourceFrom?.includes('ស្តុកចាស់')) return 'ស្តុកចាស់ក្រុម';
    if (item.operationType && opLabelMap[item.operationType]) {
      return opLabelMap[item.operationType];
    }
    return item.sourceFrom || 'ការបើកផ្តល់តាមក្រុម';
  };

  const getRecordCleanSource = (item: StockRecord): string => {
    let cleanSource = item.sourceFrom || '';
    if (cleanSource === 'openK1') cleanSource = 'ក១';
    else if (cleanSource === 'oldStockK2') cleanSource = 'ស្តុកចាស់ ក២';
    else if (cleanSource === 'oldStockTeam') cleanSource = 'ស្តុកចាស់ក្រុម';
    else if (cleanSource === 'issueTeam') cleanSource = isSecondary ? 'បើកផ្តល់តាមក្រុម' : 'ស្តុកបានទទួលពីការិយាល័យ';
    else if (cleanSource === 'useTeam') cleanSource = 'ការប្រើប្រាស់តាមក្រុម';
    else if (cleanSource === 'testPrintK2') cleanSource = 'ទិដ្ឋាការសាកក២';
    else if (cleanSource === 'damaged') cleanSource = 'ទិដ្ឋាការខូចក២';
    else if (cleanSource === 'damagedTeam') cleanSource = 'ទិដ្ឋាការខូចក្រុម';
    else if (cleanSource === 'missingTeam') cleanSource = 'ទិដ្ឋាការខ្វះក្រុម';
    else if (cleanSource === 'returnTeam') cleanSource = isSecondary ? 'ទិដ្ឋាការបង្វិលពីក្រុម' : 'ទិដ្ឋាការបង្វិលទៅក២';
    else if (cleanSource === 'transferTeam') cleanSource = 'ផ្ទេរការប្រើប្រាស់ក្រុម';
    else if (cleanSource === 'returnStub') cleanSource = 'ប្រមូលគល់សន្លឹកទិដ្ឋាការ';
    return cleanSource;
  };

  // Excel Import Modal State
  const [isImportModalOpen, setIsImportModalOpen] = useState<boolean>(false);

  // Team options filter (សន្លឹកទិដ្ឋាការ / ក្រដាសអនុម័ត) from CategoryManager
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

  const formattedToday = new Date().toLocaleDateString('en-US', {
    month: 'long',
    day: 'numeric',
    year: 'numeric',
  });

  const formatKhmerDate = (dStr: string) => {
    if (!dStr) return '';
    try {
      const parts = dStr.split('-');
      if (parts.length === 3) {
        const y = parseInt(parts[0], 10);
        const m = parseInt(parts[1], 10);
        const d = parseInt(parts[2], 10);
        const khmerMonths = [
          'មករា', 'កុម្ភៈ', 'មីនា', 'មេសា', 'ឧសភា', 'មិថុនា',
          'កក្កដា', 'សីហា', 'កញ្ញា', 'តុលា', 'វិច្ឆិកា', 'ធ្នូ',
        ];
        if (m >= 1 && m <= 12) {
          return `${d} ${khmerMonths[m - 1]} ${y}`;
        }
      }
      const dt = new Date(dStr);
      if (!isNaN(dt.getTime())) {
        const khmerMonths = [
          'មករា', 'កុម្ភៈ', 'មីនា', 'មេសា', 'ឧសភា', 'មិថុនា',
          'កក្កដា', 'សីហា', 'កញ្ញា', 'តុលា', 'វិច្ឆិកា', 'ធ្នូ',
        ];
        return `${dt.getDate()} ${khmerMonths[dt.getMonth()]} ${dt.getFullYear()}`;
      }
      return dStr;
    } catch {
      return dStr;
    }
  };

  // Records saved on the selected date for current stockType
  const sameDateRecords = useMemo(() => {
    if (!date) return [];
    return stockRecords.filter((r) => r.stockType === stockType && r.date === date);
  }, [stockRecords, stockType, date]);

  // Breakdown of quantities by visa type on this date
  const sameDateTypeSummary = useMemo(() => {
    const map: Record<string, number> = {};
    sameDateRecords.forEach((r) => {
      const type = r.visaType || 'ផ្សេងៗ';
      const qty = r.quantityBundles || r.totalSheets || 0;
      map[type] = (map[type] || 0) + qty;
    });
    return map;
  }, [sameDateRecords]);

  // Total sheets recorded on this date
  const totalSheetsOnDate = useMemo(() => {
    return sameDateRecords.reduce((sum, r) => sum + (r.quantityBundles || r.totalSheets || 0), 0);
  }, [sameDateRecords]);

  // Load all records of this date into the multi-row form
  const handleLoadAllSameDateRecords = () => {
    if (sameDateRecords.length === 0) return;
    const eligibleRecords = !isSecondary
      ? sameDateRecords.filter((r) => r.operationType !== 'issueTeam')
      : sameDateRecords;
    if (eligibleRecords.length === 0) {
      if (!isSecondary && sameDateRecords.some((r) => r.operationType === 'issueTeam')) {
        onShowToast('មិនអនុញ្ញាតកែប្រែប្រតិបត្តិការ (ការបើកផ្តល់តាមក្រុម) ដែលទទួលបានពីការិយាល័យឡើយ!');
      }
      return;
    }
    setStickerItems(
      eligibleRecords.map((r) => ({
        id: r.id,
        visaType: r.visaType || '',
        quantity: r.quantityBundles ? r.quantityBundles.toString() : (r.totalSheets ? r.totalSheets.toString() : ''),
        startSerial: r.startSerial || '',
        endSerial: r.endSerial || '',
      }))
    );
    if (!editingRecordId && eligibleRecords.length > 0) {
      setEditingRecordId(eligibleRecords[0].id);
    }
    onShowToast(`បានផ្ទុកទិន្នន័យ ${eligibleRecords.length} ប្រភេទទិដ្ឋាការក្នុងថ្ងៃនេះ ចូលក្នុងទម្រង់កែប្រែ!`);
  };

  // Append a specific record from the same date to the form
  const handleAppendRecordToForm = (rec: StockRecord) => {
    if (!isSecondary && rec.operationType === 'issueTeam') {
      onShowToast('មិនអនុញ្ញាតកែប្រែប្រតិបត្តិការ (ការបើកផ្តល់តាមក្រុម) ដែលទទួលបានពីការិយាល័យឡើយ!');
      return;
    }
    if (stickerItems.some((it) => it.id === rec.id)) {
      onShowToast(`ប្រភេទទិដ្ឋាការ ${rec.visaType || ''} មានក្នុងទម្រង់កែប្រែរួចហើយ!`);
      return;
    }
    setStickerItems((prev) => {
      const cleaned = prev.filter((it) => it.visaType || it.quantity || it.startSerial || it.endSerial);
      return [
        ...cleaned,
        {
          id: rec.id,
          visaType: rec.visaType || '',
          quantity: rec.quantityBundles ? rec.quantityBundles.toString() : (rec.totalSheets ? rec.totalSheets.toString() : ''),
          startSerial: rec.startSerial || '',
          endSerial: rec.endSerial || '',
        },
      ];
    });
    if (!editingRecordId) {
      setEditingRecordId(rec.id);
    }
    onShowToast(`បានបញ្ចូលប្រភេទ ${rec.visaType || ''} ទៅក្នុងទម្រង់កែប្រែ!`);
  };

  const handleStartEdit = (record: StockRecord) => {
    if (!isSecondary && record.operationType === 'issueTeam') {
      onShowToast('មិនអនុញ្ញាតកែប្រែប្រតិបត្តិការ (ការបើកផ្តល់តាមក្រុម) ដែលទទួលបានពីការិយាល័យឡើយ!');
      return;
    }

    // Redirect only 'useTeam' (ការប្រើប្រាស់តាមក្រុម) edit requests to 'ប្រតិបត្តិការប្រចាំថ្ងៃតាមក្រុម'
    // Ensure 'transferTeam' or 'ផ្ទេរ' related operations are EXCLUDED from this redirection.
    const isUseTeamOp =
      (record.operationType === 'useTeam' ||
        record.operationType === 'use_team' ||
        record.operationType === 'use' ||
        record.operationType === 'useSticker' ||
        record.operationType === 'useEVisa' ||
        record.operationType === 'use_cea' ||
        (record.operationType || '').toLowerCase().includes('use') ||
        (record.operationType || '').includes('ប្រើប្រាស់') ||
        (record.sourceFrom || '').includes('ប្រើប្រាស់')) &&
      !(record.operationType || '').includes('ផ្ទេរ') &&
      !(record.sourceFrom || '').includes('ផ្ទេរ') &&
      !record.id?.startsWith('stock-tr-'); // Ensure transfer IDs are not redirected

    if (isUseTeamOp) {
      const targetTeamName =
        record.visaTeamRobokName ||
        resolveRecordTeamName(record) ||
        (categories?.visaTeamsRobok || []).find((t) => t.id === record.visaTeamRobokId)?.name ||
        record.recipientTeamName ||
        record.sourceFrom ||
        '';

      if (record.date) {
        localStorage.setItem('app_daily_team_selected_date', record.date);
      }
      if (targetTeamName) {
        localStorage.setItem('app_daily_team_selected_team', targetTeamName);
      }
      const isCea =
        record.stockType === 'evisa' ||
        isCeaRecord(record) ||
        record.sourceFrom === 'cEA' ||
        record.remarks?.includes('cEA') ||
        (record.visaType || '').includes('cEA');
      const optionMode = isCea ? 'cEA' : 'Sticker';
      localStorage.setItem('app_daily_team_selected_option', optionMode);

      if (record.id) {
        localStorage.setItem('app_daily_team_edit_record_id', record.id);
        localStorage.setItem('app_daily_team_auto_load_form', 'true');
        localStorage.setItem(
          'app_daily_team_edit_payload',
          JSON.stringify({
            id: record.id,
            date: record.date,
            visaType: record.visaType,
            quantityBundles: record.quantityBundles,
            totalSheets: record.totalSheets,
            startSerial: record.startSerial,
            endSerial: record.endSerial,
            oldCode: (record as any).oldCode || '',
            targetTeamName,
            optionMode,
          })
        );
      }

      window.dispatchEvent(new Event('app_daily_team_selection_updated'));

      if (onNavigate) {
        onNavigate('stockStickerDailyTeam');
      }
      onShowToast(`បានប្តូរទៅកាន់ «ប្រតិបត្តិការប្រចាំថ្ងៃតាមក្រុម» ដើម្បីកែប្រែទិន្នន័យ (${targetTeamName || 'ក្រុម'} - ${record.date || ''}) [${optionMode === 'cEA' ? 'ប្រើប្រាស់ cEA' : 'ប្រើប្រាស់ Sticker'}]`);
      return;
    }

    setEditingRecordId(record.id);
    setViewMode('form');
    setStockType(record.stockType);
    setOperationType(record.operationType);
    setDate(record.date || new Date().toISOString().split('T')[0]);
    setTime(record.time || new Date().toLocaleTimeString('en-US', { hour12: false, hour: '2-digit', minute: '2-digit' }));
    setSourceFrom(record.sourceFrom || (record.operationType === 'issueTeam' ? 'បើកផ្តល់តាមក្រុម' : 'ក១'));
    setQuantityBundles(record.quantityBundles ? record.quantityBundles.toString() : '');

    // Resolve team robok ID (support matching by ID or by Name)
    let teamId = record.visaTeamRobokId || '';
    if (record.visaTeamRobokName) {
      const foundTeam = (categories.visaTeamsRobok || []).find(
        (vtr) => vtr.name.trim() === record.visaTeamRobokName?.trim() || vtr.id === record.visaTeamRobokName
      );
      if (foundTeam) {
        teamId = foundTeam.id;
      } else if (!teamId) {
        teamId = record.visaTeamRobokName;
      }
    } else if (teamId) {
      const foundTeam = (categories.visaTeamsRobok || []).find(
        (vtr) => vtr.name.trim() === teamId.trim() || vtr.id === teamId
      );
      if (foundTeam) {
        teamId = foundTeam.id;
      }
    }
    setVisaTeamRobokId(teamId);

    // Resolve requested rank ID
    let rankId = record.requestedRankId || '';
    if (record.requestedRankName) {
      const foundRank = (categories.ranks || []).find(
        (r) => r.name.trim() === record.requestedRankName?.trim() || r.id === record.requestedRankName
      );
      if (foundRank) {
        rankId = foundRank.id;
      } else if (!rankId) {
        rankId = record.requestedRankName;
      }
    } else if (rankId) {
      const foundRank = (categories.ranks || []).find(
        (r) => r.name.trim() === rankId.trim() || r.id === rankId
      );
      if (foundRank) {
        rankId = foundRank.id;
      }
    }
    setRequestedRankId(rankId);

    // Resolve collector role ID
    let roleId = record.collectorRoleId || '';
    if (record.collectorRoleName) {
      const foundRole = (categories.collectorRoles || []).find(
        (cr) => cr.name.trim() === record.collectorRoleName?.trim() || cr.id === record.collectorRoleName
      );
      if (foundRole) {
        roleId = foundRole.id;
      } else if (!roleId) {
        roleId = record.collectorRoleName;
      }
    } else if (roleId) {
      const foundRole = (categories.collectorRoles || []).find(
        (cr) => cr.name.trim() === roleId.trim() || cr.id === roleId
      );
      if (foundRole) {
        roleId = foundRole.id;
      }
    }
    setCollectorRoleId(roleId);

    setVisaType(record.visaType || '');
    setRequesterName(record.requesterName || '');
    setCollectorName(record.collectorName || '');
    setStartSerial(record.startSerial || '');
    setEndSerial(record.endSerial || '');
    setTotalSheets(record.totalSheets ? record.totalSheets.toString() : '');

    // Load transfer attributes
    let recId = record.recipientTeamId || '';
    if (record.recipientTeamName && (!recId || !categories?.visaTeamsRobok?.some((vt) => vt.id === recId))) {
      const foundRec = (categories?.visaTeamsRobok || []).find(
        (vt) => vt.name.trim().toLowerCase() === record.recipientTeamName?.trim().toLowerCase() || vt.id === record.recipientTeamName
      );
      if (foundRec) {
        recId = foundRec.id;
      } else if (!recId) {
        recId = record.recipientTeamName;
      }
    }
    setRecipientTeamId(recId);
    setOfficeApproved(!!record.officeApproved);
    setDeptApprovalDate(record.deptApprovalDate || '');

    if (record.stockType === 'sticker') {
      // Find all records belonging to the same batch (same date, operation, and team/source)
      const sameBatchRecords = stockRecords.filter(
        (r) =>
          r.stockType === 'sticker' &&
          r.date === record.date &&
          r.operationType === record.operationType &&
          ((!record.visaTeamRobokName && !r.visaTeamRobokName) ||
            r.visaTeamRobokName === record.visaTeamRobokName ||
            r.visaTeamRobokId === record.visaTeamRobokId) &&
          ((!record.sourceFrom && !r.sourceFrom) || r.sourceFrom === record.sourceFrom)
      );

      if (sameBatchRecords.length > 0) {
        // Place the currently clicked record first, followed by other records from the same batch
        const otherRecords = sameBatchRecords.filter((r) => r.id !== record.id);
        const combined = [record, ...otherRecords];
        setStickerItems(
          combined.map((r) => ({
            id: r.id,
            visaType: r.visaType || '',
            quantity: r.quantityBundles ? r.quantityBundles.toString() : (r.totalSheets ? r.totalSheets.toString() : ''),
            startSerial: r.startSerial || '',
            endSerial: r.endSerial || '',
          }))
        );
      } else {
        setStickerItems([
          {
            id: record.id,
            visaType: record.visaType || '',
            quantity: record.quantityBundles ? record.quantityBundles.toString() : (record.totalSheets ? record.totalSheets.toString() : ''),
            startSerial: record.startSerial || '',
            endSerial: record.endSerial || '',
          },
        ]);
      }
    } else {
      setStickerItems([{ id: '1', visaType: '', quantity: '', startSerial: '', endSerial: '' }]);
    }

    window.scrollTo({ top: 0, behavior: 'smooth' });
    document.querySelector('.custom-main-scroll')?.scrollTo({ top: 0, behavior: 'smooth' });
  };

  const handleCancelEdit = () => {
    setEditingRecordId(null);
    setSourceFrom('');
    setQuantityBundles('');
    setVisaType('');
    setRequesterName('');
    setCollectorName('');
    setVisaTeamRobokId('');
    setRequestedRankId('');
    setCollectorRoleId('');
    setStartSerial('');
    setEndSerial('');
    setTotalSheets('');
    setRecipientTeamId('');
    setOfficeApproved(false);
    setDeptApprovalDate('');
    setStickerItems([{ id: '1', visaType: '', quantity: '', startSerial: '', endSerial: '' }]);
    if (viewMode === 'form') {
      setViewMode('data');
    }
  };

  const handleConfirmDelete = () => {
    if (deletingRecord) {
      if (!isSecondary && deletingRecord.operationType === 'issueTeam') {
        onShowToast('មិនអនុញ្ញាតលុបទិន្នន័យប្រតិបត្តិការ (ការបើកផ្តល់តាមក្រុម) ដែលទទួលបានពីការិយាល័យឡើយ!');
        setDeletingRecord(null);
        return;
      }
      onDeleteStockRecord(deletingRecord.id);
      if (editingRecordId === deletingRecord.id) {
        handleCancelEdit();
      }
      onShowToast('បានលុបទិន្នន័យស្តុកជោគជ័យ!');
      setDeletingRecord(null);
    }
  };

  const handleConfirmDeleteAll = () => {
    const currentTypeRecords = stockRecords.filter((r) => r.stockType === stockType);
    if (currentTypeRecords.length === 0) {
      onShowToast('គ្មានទិន្នន័យសម្រាប់លុបទេ!');
      setIsDeleteAllModalOpen(false);
      return;
    }

    let targetRecords = currentTypeRecords;
    // Team users are not allowed to delete issueTeam (received from office) records
    if (!isSecondary) {
      targetRecords = targetRecords.filter((r) => r.operationType !== 'issueTeam');
    }
    let successMsg = `បានលុបទិន្នន័យ${stockType === 'evisa' ? 'ក្រដាសអនុម័ត' : 'សន្លឹកទិដ្ឋាការ'}ទាំងអស់ជោគជ័យ!`;

    if (deleteScope === 'filtered' && filterOperationType !== 'all') {
      targetRecords = currentTypeRecords.filter((r) => matchOperation(r, filterOperationType));
      if (!isSecondary) {
        targetRecords = targetRecords.filter((r) => r.operationType !== 'issueTeam');
      }
      const opName = opLabelMap[filterOperationType] || filterOperationType;
      successMsg = `បានលុបទិន្នន័យ «${opName}» ចំនួន ${targetRecords.length} ជួរជោគជ័យ!`;
    }

    if (targetRecords.length === 0) {
      onShowToast(!isSecondary && filterOperationType === 'issueTeam' ? 'មិនអនុញ្ញាតលុបទិន្នន័យប្រតិបត្តិការ (ការបើកផ្តល់តាមក្រុម) ដែលទទួលបានពីការិយាល័យឡើយ!' : 'គ្មានទិន្នន័យសម្រាប់លុបទេ!');
      setIsDeleteAllModalOpen(false);
      return;
    }

    const idsToDelete = targetRecords.map((r) => r.id);
    const targetOp = deleteScope === 'filtered' && filterOperationType !== 'all' ? filterOperationType : undefined;
    if (onDeleteBatchStockRecords) {
      onDeleteBatchStockRecords(idsToDelete, stockType, targetOp);
    } else {
      idsToDelete.forEach((id) => onDeleteStockRecord(id));
    }

    onShowToast(successMsg);
    setIsDeleteAllModalOpen(false);
  };

  // Handle Form Submit
  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();

    if (!date) {
      onShowToast('សូមជ្រើសរើសកាលបរិច្ឆេទ!');
      return;
    }

    // Special logic for sticker (សន្លឹកទិដ្ឋាការស្អិត)
    if (stockType === 'sticker') {
      const validItems = stickerItems.filter(
        (item) => item.visaType.trim() !== '' || item.quantity.trim() !== '' || item.startSerial.trim() !== ''
      );

      if (validItems.length === 0) {
        onShowToast('សូមជ្រើសរើសប្រភេទទិដ្ឋាការ និងបញ្ចូលចំនួនសន្លឹក!');
        return;
      }

      for (const item of validItems) {
        if (!item.visaType) {
          onShowToast('សូមជ្រើសរើសប្រភេទទិដ្ឋាការ!');
          return;
        }
        const q = parseInt(item.quantity, 10);
        if (isNaN(q) || q <= 0) {
          onShowToast(`សូមបញ្ចូលចំនួនសម្រាប់ប្រភេទទិដ្ឋាការ ${item.visaType} ឱ្យបានត្រឹមត្រូវ!`);
          return;
        }
      }

      let teamName = '';
      let rankName = '';
      let roleName = '';
      let recTeamName = '';

      const isTeamOp =
        operationType === 'issueTeam' ||
        operationType === 'useTeam' ||
        operationType === 'oldStockTeam' ||
        operationType === 'returnTeam' ||
        operationType === 'transferTeam' ||
        operationType === 'damagedTeam' ||
        operationType === 'missingTeam' ||
        operationType === 'returnStub' ||
        (!isSecondary && Boolean(assignedTeam));

      if (isTeamOp) {
        const effectiveTeamKey = visaTeamRobokId || assignedTeam;
        if (effectiveTeamKey) {
          teamName = categories.visaTeamsRobok.find((vt) => vt.id === effectiveTeamKey || vt.name === effectiveTeamKey)?.name || effectiveTeamKey;
        }
        if (operationType === 'issueTeam' || operationType === 'transferTeam') {
          rankName = categories.ranks.find((r) => r.id === requestedRankId)?.name || requestedRankId;
          roleName = categories.collectorRoles.find((cr) => cr.id === collectorRoleId)?.name || collectorRoleId;
        }
        if (operationType === 'transferTeam' && recipientTeamId) {
          recTeamName = categories.visaTeamsRobok.find((vt) => vt.id === recipientTeamId || vt.name === recipientTeamId)?.name || recipientTeamId;
        }
      }

      let effSource = sourceFrom;
      if (!effSource || (operationType === 'issueTeam' && (effSource === 'ក១' || !effSource))) {
        if (operationType === 'openK1') effSource = 'ក១';
        else if (operationType === 'issueTeam') effSource = 'បើកផ្តល់តាមក្រុម';
        else if (operationType === 'oldStockK2') effSource = 'ស្តុកចាស់ ក២';
        else if (operationType === 'oldStockTeam') effSource = 'ស្តុកចាស់ក្រុម';
        else if (operationType === 'testPrintK2') effSource = 'ទិដ្ឋាការសាកក២';
        else if (operationType === 'damaged') effSource = 'ទិដ្ឋាការខូចក២';
        else if (operationType === 'damagedTeam') effSource = 'ទិដ្ឋាការខូចក្រុម';
        else if (operationType === 'missingTeam') effSource = 'ទិដ្ឋាការខ្វះក្រុម';
        else if (operationType === 'returnTeam') effSource = 'ទិដ្ឋាការបង្វិលពីក្រុម';
        else if (operationType === 'transferTeam') effSource = 'ផ្ទេរការប្រើប្រាស់ក្រុម';
        else if (operationType === 'returnStub') effSource = 'ប្រមូលគល់សន្លឹកទិដ្ឋាការ';
        else if (operationType === 'returnK1') effSource = 'បង្វិលក១';
        else effSource = operationType;
      }
      const affectedIds: string[] = [];

      if (editingRecordId) {
        if (!isSecondary && operationType === 'issueTeam') {
          onShowToast('មិនអនុញ្ញាតកែប្រែប្រតិបត្តិការ (ការបើកផ្តល់តាមក្រុម) ដែលទទួលបានពីការិយាល័យឡើយ!');
          return;
        }
        validItems.forEach((it, idx) => {
          const q = parseInt(it.quantity, 10) || 0;
          const existingRecord = stockRecords.find((r) => r.id === it.id);

          if (existingRecord) {
            const updatedRecord: StockRecord = {
              ...existingRecord,
              stockType: 'sticker',
              operationType,
              date,
              time: time || undefined,
              sourceFrom: effSource || undefined,
              quantityBundles: q,
              totalSheets: q,
              visaTeamRobokId: isTeamOp ? (visaTeamRobokId || assignedTeam || undefined) : undefined,
              visaTeamRobokName: teamName || (isTeamOp ? assignedTeam : undefined),
              visaType: it.visaType,
              requestedRankId: operationType === 'issueTeam' ? requestedRankId : undefined,
              requestedRankName: rankName || undefined,
              requesterName: (operationType === 'issueTeam' || operationType === 'returnTeam' || operationType === 'returnStub' || operationType === 'damagedTeam' || operationType === 'missingTeam') ? requesterName : undefined,
              collectorName: operationType === 'issueTeam' ? collectorName : undefined,
              collectorRoleId: operationType === 'issueTeam' ? collectorRoleId : undefined,
              collectorRoleName: roleName || undefined,
              startSerial: it.startSerial.trim() || undefined,
              endSerial: it.endSerial.trim() || undefined,
              recipientTeamId: operationType === 'transferTeam' ? (recipientTeamId || undefined) : undefined,
              recipientTeamName: operationType === 'transferTeam' ? (recTeamName || undefined) : undefined,
              officeApproved: operationType === 'transferTeam' ? officeApproved : undefined,
              deptApprovalDate: operationType === 'transferTeam' ? (deptApprovalDate || undefined) : undefined,
            };
            onUpdateStockRecord(updatedRecord);
            affectedIds.push(updatedRecord.id);
          } else {
            const newId = `stock-${Date.now()}-${idx}-${Math.random().toString(36).substr(2, 4)}`;
            const newRecord: StockRecord = {
              id: newId,
              stockType: 'sticker',
              operationType,
              date,
              time: time || undefined,
              sourceFrom: effSource || undefined,
              quantityBundles: q,
              totalSheets: q,
              visaTeamRobokId: isTeamOp ? (visaTeamRobokId || assignedTeam || undefined) : undefined,
              visaTeamRobokName: teamName || (isTeamOp ? assignedTeam : undefined),
              visaType: it.visaType,
              requestedRankId: operationType === 'issueTeam' ? requestedRankId : undefined,
              requestedRankName: rankName || undefined,
              requesterName: (operationType === 'issueTeam' || operationType === 'returnTeam' || operationType === 'returnStub' || operationType === 'damagedTeam' || operationType === 'missingTeam') ? requesterName : undefined,
              collectorName: operationType === 'issueTeam' ? collectorName : undefined,
              collectorRoleId: operationType === 'issueTeam' ? collectorRoleId : undefined,
              collectorRoleName: roleName || undefined,
              startSerial: it.startSerial.trim() || undefined,
              endSerial: it.endSerial.trim() || undefined,
              recipientTeamId: operationType === 'transferTeam' ? (recipientTeamId || undefined) : undefined,
              recipientTeamName: operationType === 'transferTeam' ? (recTeamName || undefined) : undefined,
              officeApproved: operationType === 'transferTeam' ? officeApproved : undefined,
              deptApprovalDate: operationType === 'transferTeam' ? (deptApprovalDate || undefined) : undefined,
              createdAt: new Date().toISOString(),
              createdBy: userName,
            };
            onAddStockRecord(newRecord);
            affectedIds.push(newRecord.id);
          }
        });
        onShowToast(`បានកែប្រែទិន្នន័យស្តុកចំនួន ${validItems.length} ប្រភេទទិដ្ឋាការ ជោគជ័យ!`);
        setEditingRecordId(null);
      } else {
        validItems.forEach((it, idx) => {
          const q = parseInt(it.quantity, 10) || 0;
          const newId = `stock-${Date.now()}-${idx}-${Math.random().toString(36).substr(2, 4)}`;
          const newRecord: StockRecord = {
            id: newId,
            stockType: 'sticker',
            operationType,
            date,
            time: time || undefined,
            sourceFrom: effSource || undefined,
            quantityBundles: q,
            totalSheets: q,
            visaTeamRobokId: isTeamOp ? (visaTeamRobokId || assignedTeam || undefined) : undefined,
            visaTeamRobokName: teamName || (isTeamOp ? assignedTeam : undefined),
            visaType: it.visaType,
            requestedRankId: operationType === 'issueTeam' ? requestedRankId : undefined,
            requestedRankName: rankName || undefined,
            requesterName: (operationType === 'issueTeam' || operationType === 'returnTeam' || operationType === 'returnStub' || operationType === 'damagedTeam' || operationType === 'missingTeam') ? requesterName : undefined,
            collectorName: operationType === 'issueTeam' ? collectorName : undefined,
            collectorRoleId: operationType === 'issueTeam' ? collectorRoleId : undefined,
            collectorRoleName: roleName || undefined,
            startSerial: it.startSerial.trim() || undefined,
            endSerial: it.endSerial.trim() || undefined,
            recipientTeamId: operationType === 'transferTeam' ? (recipientTeamId || undefined) : undefined,
            recipientTeamName: operationType === 'transferTeam' ? (recTeamName || undefined) : undefined,
            officeApproved: operationType === 'transferTeam' ? officeApproved : undefined,
            deptApprovalDate: operationType === 'transferTeam' ? (deptApprovalDate || undefined) : undefined,
            createdAt: new Date().toISOString(),
            createdBy: userName,
          };
          onAddStockRecord(newRecord);
          affectedIds.push(newRecord.id);
        });
        onShowToast(`បានរក្សាទុកទិន្នន័យស្តុក ${validItems.length} ប្រភេទទិដ្ឋាការ ជោគជ័យ!`);
      }

      if (affectedIds.length > 0) setHighlightedRecordIds(affectedIds);

      // Reset
      setSourceFrom('ក១');
      setQuantityBundles('');
      setVisaType('');
      setRequesterName('');
      setCollectorName('');
      setVisaTeamRobokId('');
      setRequestedRankId('');
      setCollectorRoleId('');
      setStartSerial('');
      setEndSerial('');
      setTotalSheets('');
      setRecipientTeamId('');
      setOfficeApproved(false);
      setDeptApprovalDate('');
      setStickerItems([{ id: '1', visaType: '', quantity: '', startSerial: '', endSerial: '' }]);

      if (operationType === 'issueTeam') {
        const primaryRecord: StockRecord = {
          id: affectedIds[0] || `stock-${Date.now()}`,
          stockType: 'sticker',
          operationType: 'issueTeam',
          date,
          time: time || undefined,
          sourceFrom: effSource || undefined,
          quantityBundles: parseInt(validItems[0]?.quantity, 10) || 0,
          totalSheets: parseInt(validItems[0]?.quantity, 10) || 0,
          visaTeamRobokId: isTeamOp ? (visaTeamRobokId || assignedTeam || undefined) : undefined,
          visaTeamRobokName: teamName || (isTeamOp ? assignedTeam : undefined),
          visaType: validItems[0]?.visaType,
          requestedRankId: requestedRankId || undefined,
          requestedRankName: rankName || undefined,
          requesterName: requesterName || undefined,
          collectorName: collectorName || undefined,
          collectorRoleId: collectorRoleId || undefined,
          collectorRoleName: roleName || undefined,
          startSerial: validItems[0]?.startSerial?.trim() || undefined,
          endSerial: validItems[0]?.endSerial?.trim() || undefined,
          createdAt: new Date().toISOString(),
          createdBy: userName,
        };
        setSelectedPdfRecord(primaryRecord);
        setViewMode('handoverWorkspace');
        onShowToast('បានរក្សាទុកការបើកផ្តល់តាមក្រុមជោគជ័យ! កំពុងបញ្ជូនទៅកាន់លិខិតប្រគល់ទទួល...');
      } else {
        setViewMode('data');
      }
      return;
    }

    if ((operationType === 'openK1' || operationType === 'oldStockK2') && !sourceFrom) {
      if (operationType === 'oldStockK2') {
        setSourceFrom('ស្តុកចាស់ ក២');
      } else {
        onShowToast('សូមជ្រើសរើសប្រភព (បើកពី)!');
        return;
      }
    }

    const qty = parseInt(quantityBundles, 10);
    if (isNaN(qty) || qty <= 0) {
      onShowToast('សូមបញ្ចូលចំនួនដុំឲ្យបានត្រឹមត្រូវ!');
      return;
    }

    let teamName = '';
    let rankName = '';
    let roleName = '';
    let recTeamName = '';

    if (operationType === 'issueTeam' || operationType === 'useTeam' || operationType === 'oldStockTeam' || operationType === 'transferTeam') {
      if (!visaTeamRobokId) {
        onShowToast('សូមជ្រើសរើសក្រុមផ្តល់ទិដ្ឋាការ.របក!');
        return;
      }
      teamName = categories.visaTeamsRobok.find((vt) => vt.id === visaTeamRobokId)?.name || '';
      if (operationType === 'issueTeam') {
        rankName = categories.ranks.find((r) => r.id === requestedRankId)?.name || '';
        roleName = categories.collectorRoles.find((cr) => cr.id === collectorRoleId)?.name || '';
      }
      if (operationType === 'transferTeam' && recipientTeamId) {
        recTeamName = categories.visaTeamsRobok.find((vt) => vt.id === recipientTeamId || vt.name === recipientTeamId)?.name || recipientTeamId;
      }
    }

    const targetId = editingRecordId || `stock-${Date.now()}`;
    const parsedSheets = totalSheets ? parseInt(totalSheets, 10) : qty * 100;
    const finalSourceFrom =
      operationType === 'oldStockK2'
        ? sourceFrom || 'ស្តុកចាស់ ក២'
        : operationType === 'oldStockTeam'
        ? sourceFrom || 'ស្តុកចាស់ក្រុម'
        : operationType === 'openK1'
        ? sourceFrom
        : undefined;

    if (editingRecordId) {
      if (!isSecondary && operationType === 'issueTeam') {
        onShowToast('មិនអនុញ្ញាតកែប្រែប្រតិបត្តិការ (ការបើកផ្តល់តាមក្រុម) ដែលទទួលបានពីការិយាល័យឡើយ!');
        return;
      }
      const original = stockRecords.find((r) => r.id === editingRecordId);
      const updatedRecord: StockRecord = {
        id: editingRecordId,
        stockType,
        operationType,
        date,
        time: operationType === 'issueTeam' ? time : undefined,
        sourceFrom: finalSourceFrom,
        quantityBundles: qty,
        visaTeamRobokId: (operationType === 'issueTeam' || operationType === 'useTeam' || operationType === 'oldStockTeam' || operationType === 'transferTeam') ? visaTeamRobokId : undefined,
        visaTeamRobokName: teamName,
        visaType: visaType.trim() || undefined,
        requestedRankId: operationType === 'issueTeam' ? requestedRankId : undefined,
        requestedRankName: rankName,
        requesterName: operationType === 'issueTeam' ? requesterName : undefined,
        collectorName: operationType === 'issueTeam' ? collectorName : undefined,
        collectorRoleId: operationType === 'issueTeam' ? collectorRoleId : undefined,
        collectorRoleName: roleName,
        startSerial: startSerial.trim() || undefined,
        endSerial: endSerial.trim() || undefined,
        totalSheets: isNaN(parsedSheets) ? undefined : parsedSheets,
        recipientTeamId: operationType === 'transferTeam' ? (recipientTeamId || undefined) : undefined,
        recipientTeamName: operationType === 'transferTeam' ? (recTeamName || undefined) : undefined,
        officeApproved: operationType === 'transferTeam' ? officeApproved : undefined,
        deptApprovalDate: operationType === 'transferTeam' ? (deptApprovalDate || undefined) : undefined,
        createdAt: original?.createdAt || new Date().toISOString(),
        createdBy: original?.createdBy || userName,
      };

      onUpdateStockRecord(updatedRecord);
      onShowToast('បានកែប្រែទិន្នន័យស្តុកជោគជ័យ!');
      setEditingRecordId(null);
    } else {
      const newRecord: StockRecord = {
        id: targetId,
        stockType,
        operationType,
        date,
        time: operationType === 'issueTeam' ? time : undefined,
        sourceFrom: finalSourceFrom,
        quantityBundles: qty,
        visaTeamRobokId: (operationType === 'issueTeam' || operationType === 'useTeam' || operationType === 'oldStockTeam' || operationType === 'transferTeam') ? visaTeamRobokId : undefined,
        visaTeamRobokName: teamName,
        visaType: visaType.trim() || undefined,
        requestedRankId: operationType === 'issueTeam' ? requestedRankId : undefined,
        requestedRankName: rankName,
        requesterName: operationType === 'issueTeam' ? requesterName : undefined,
        collectorName: operationType === 'issueTeam' ? collectorName : undefined,
        collectorRoleId: operationType === 'issueTeam' ? collectorRoleId : undefined,
        collectorRoleName: roleName,
        startSerial: startSerial.trim() || undefined,
        endSerial: endSerial.trim() || undefined,
        totalSheets: isNaN(parsedSheets) ? undefined : parsedSheets,
        recipientTeamId: operationType === 'transferTeam' ? (recipientTeamId || undefined) : undefined,
        recipientTeamName: operationType === 'transferTeam' ? (recTeamName || undefined) : undefined,
        officeApproved: operationType === 'transferTeam' ? officeApproved : undefined,
        deptApprovalDate: operationType === 'transferTeam' ? (deptApprovalDate || undefined) : undefined,
        createdAt: new Date().toISOString(),
        createdBy: userName,
      };

      onAddStockRecord(newRecord);
      onShowToast('បានរក្សាទុកទិន្នន័យស្តុកជោគជ័យ!');
    }

    setHighlightedRecordIds([targetId]);

    // Reset Form
    setSourceFrom('');
    setQuantityBundles('');
    setVisaType('');
    setRequesterName('');
    setCollectorName('');
    setVisaTeamRobokId('');
    setRequestedRankId('');
    setCollectorRoleId('');
    setStartSerial('');
    setEndSerial('');
    setTotalSheets('');
    setRecipientTeamId('');
    setOfficeApproved(false);
    setDeptApprovalDate('');

    if (operationType === 'issueTeam') {
      const primaryRecord: StockRecord = {
        id: targetId,
        stockType: 'evisa',
        operationType: 'issueTeam',
        date,
        time: time || undefined,
        sourceFrom: finalSourceFrom || undefined,
        quantityBundles: qty,
        totalSheets: isNaN(parsedSheets) ? undefined : parsedSheets,
        visaTeamRobokId: visaTeamRobokId || undefined,
        visaTeamRobokName: teamName,
        visaType: visaType.trim() || undefined,
        requestedRankId: requestedRankId || undefined,
        requestedRankName: rankName,
        requesterName: requesterName || undefined,
        collectorName: collectorName || undefined,
        collectorRoleId: collectorRoleId || undefined,
        collectorRoleName: roleName,
        startSerial: startSerial.trim() || undefined,
        endSerial: endSerial.trim() || undefined,
        createdAt: new Date().toISOString(),
        createdBy: userName,
      };
      setSelectedPdfRecord(primaryRecord);
      setViewMode('handoverWorkspace');
      onShowToast('បានរក្សាទុកការបើកផ្តល់តាមក្រុមជោគជ័យ! កំពុងបញ្ជូនទៅកាន់លិខិតប្រគល់ទទួល...');
    } else {
      // Go to data table view ("ទិន្នន័យក្រដាសអនុម័ត")
      setViewMode('data');
    }
  };

  // Helper functions to check team report inclusion
  const normalizeTeamName = (name: string): string => {
    if (!name) return '';
    const cleaned = name.trim().toLowerCase().replace(/\s+/g, ' ');

    if (
      cleaned.includes('ក្អមសំណ') ||
      cleaned.includes('ក្អម') ||
      cleaned.includes('រអមសំណ') ||
      cleaned.includes('ភូមិសាលា') ||
      cleaned.includes('លើកដែក')
    ) {
      return 'ព្រំដែន ក្អមសំណ';
    }

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
  };

  const normalizeVisaType = (vRaw?: string): string => {
    if (!vRaw) return '';
    const cleaned = vRaw.trim().toUpperCase();
    if (cleaned.includes('T-1') || cleaned === 'T1' || cleaned === 'T 1') return 'T1';
    if (cleaned.includes('T-2') || cleaned === 'T2' || cleaned === 'T 2') return 'T2';
    if (cleaned.includes('T-3') || cleaned === 'T3' || cleaned === 'T 3') return 'T3';
    if (cleaned === 'T' || cleaned.startsWith('T ')) return 'T';
    if (cleaned.includes('E-1') || cleaned === 'E1' || cleaned === 'E 1') return 'E1';
    if (cleaned.includes('E-2') || cleaned === 'E2' || cleaned === 'E 2') return 'E2';
    if (cleaned.includes('E-3') || cleaned === 'E3' || cleaned === 'E 3') return 'E3';
    if (cleaned === 'E' || cleaned.startsWith('E ')) return 'E';
    if (cleaned.startsWith('D')) return 'D';
    if (cleaned.startsWith('K')) return 'K';
    if (cleaned.startsWith('A')) return 'A';
    if (cleaned.startsWith('B')) return 'B';
    if (cleaned.startsWith('C')) return 'C';
    return cleaned;
  };

  const isCeaRecord = (rec: StockRecord): boolean => {
    if (rec.stockType === 'evisa') return true;
    const sf = (rec.sourceFrom || '').trim().toLowerCase();
    const id = (rec.id || '').toLowerCase();
    const vt = (rec.visaType || '').trim().toLowerCase();
    const op = (rec.operationType || '').toLowerCase();
    const team = (rec.visaTeamRobokName || '').toLowerCase();

    return (
      sf.includes('cea') ||
      sf.includes('c-ea') ||
      sf.includes('c ea') ||
      sf.includes('e-visa') ||
      sf.includes('evisa') ||
      id.includes('cea') ||
      vt.includes('cea') ||
      vt.includes('c-ea') ||
      vt.includes('c ea') ||
      vt.includes('evisa') ||
      vt.includes('e-visa') ||
      op.includes('cea') ||
      team.includes('cea') ||
      team.includes('e-visa')
    );
  };

  const isOldStockTeamRecord = (rec: StockRecord): boolean => {
    if (rec.operationType === 'oldStockTeam') return true;
    const sf = (rec.sourceFrom || '').trim().toLowerCase();
    const op = (rec.operationType || '').toLowerCase();
    if (
      (op === 'issueTeam' || op === 'useTeam') &&
      (sf === 'ស្តុកចាស់ក្រុម' || sf === 'ស្តុកចាស់របស់ក្រុម' || sf.includes('ស្តុកចាស់ក្រុម') || sf.includes('ស្តុកចាស់របស់ក្រុម'))
    ) {
      return true;
    }
    if (sf.includes('ស្តុកចាស់') && Boolean(rec.visaTeamRobokName || sf.includes('ក្រុម'))) {
      return true;
    }
    return false;
  };

  const resolveRecordTeamName = (r: StockRecord): string => {
    if (r.visaTeamRobokName && r.visaTeamRobokName.trim() !== '') {
      return r.visaTeamRobokName.trim();
    }
    if (r.sourceFrom && r.sourceFrom.trim() !== '') {
      const s = r.sourceFrom.trim();
      if (!s.includes('ស្តុកចាស់ ក២') && !s.includes('ក១') && !s.includes('នាយកដ្ឋាន')) {
        return s;
      }
    }
    if (r.requesterName && r.requesterName.trim() !== '') {
      return r.requesterName.trim();
    }
    return '';
  };

  const matchTeamInList = (rawName: string, tList: string[]): string | undefined => {
    if (!rawName) return undefined;
    const normRaw = normalizeTeamName(rawName);

    const exact = tList.find((t) => t === rawName || normalizeTeamName(t) === normRaw);
    if (exact) return exact;

    const matched = tList.find((t) => {
      const normT = normalizeTeamName(t);
      return (
        normT === normRaw ||
        normRaw.includes(normT) ||
        normT.includes(normRaw) ||
        (normRaw.includes('ក្អមសំណ') && normT.includes('ក្អមសំណ')) ||
        (normRaw.includes('លើកដែក') && normT.includes('ក្អមសំណ')) ||
        (normRaw.includes('តេជោ') && normT.includes('តេជោ')) ||
        (normRaw.includes('ពោធិ៍ចិនតុង') && normT.includes('តេជោ')) ||
        (normRaw.includes('ព្រំ') && normT.includes('ព្រំ') && !normT.includes('ព្រំដែន ត្រពាំង'))
      );
    });
    return matched;
  };

  // Teams list for report checking
  const reportTeamsList = useMemo(() => {
    const catTeams = (categories?.visaTeamsRobok || [])
      .map((t) => t.name.trim())
      .filter(Boolean);
    if (catTeams.length > 0) {
      return catTeams;
    }
    return DEFAULT_29_TEAMS;
  }, [categories]);

  // Check inclusion status for team stock report
  const checkTeamReportInclusion = (
    rec: StockRecord,
    tList: string[]
  ): {
    isIncluded: boolean;
    reason: string;
    category: 'unmatched_team' | 'unmatched_visa_type' | 'k2_internal_op' | 'pre_nov_2018' | 'zero_qty' | 'cea_evisa' | 'unknown_op';
    categoryLabel: string;
  } => {
    // 1. Check if cEA or evisa
    if (rec.stockType === 'evisa' || isCeaRecord(rec)) {
      return {
        isIncluded: false,
        category: 'cea_evisa',
        categoryLabel: 'ក្រដាសអនុម័ត / cEA',
        reason: 'ជាទិន្នន័យក្រដាសអនុម័តអេឡិចត្រូនិក (cEA) មិនរាប់ក្នុងរបាយការណ៍សន្លឹកស្អិត',
      };
    }

    // 2. Check quantity
    const qty = Number(rec.quantityBundles || rec.totalSheets || (rec as any).quantity || 0);
    if (qty <= 0 || isNaN(qty)) {
      return {
        isIncluded: false,
        category: 'zero_qty',
        categoryLabel: 'ចំនួនស្មើ ០',
        reason: 'ចំនួនសន្លឹកស្មើ ០ ឬមិនត្រឹមត្រូវ',
      };
    }

    // 3. Check visa type
    const vt = normalizeVisaType(rec.visaType);
    if (!VISA_TYPES.includes(vt as any)) {
      return {
        isIncluded: false,
        category: 'unmatched_visa_type',
        categoryLabel: 'មិនស្គាល់ប្រភេទទិដ្ឋាការ',
        reason: `ប្រភេទទិដ្ឋាការ «${rec.visaType || 'ទទេ'}» មិនត្រូវនឹងប្រភេទទាំង ១៣ (T, T1, T2, T3, E, E1, E2, E3, D, K, A, B, C)`,
      };
    }

    // 4. Resolve team
    const rawTeam = resolveRecordTeamName(rec);
    const targetTeam = matchTeamInList(rawTeam, tList);
    const normTeam = normalizeTeamName(rawTeam);
    const effectiveTeam = tList.find((t) => t === targetTeam || t === rawTeam || normalizeTeamName(t) === normTeam) || targetTeam;

    // 5. Check if it's Old Stock Team
    if (isOldStockTeamRecord(rec)) {
      if (!effectiveTeam) {
        return {
          isIncluded: false,
          category: 'unmatched_team',
          categoryLabel: 'មិនស្គាល់ឈ្មោះក្រុម',
          reason: `ស្តុកចាស់ក្រុម តែមិនស្គាល់ឈ្មោះក្រុម «${rawTeam || 'ទទេ'}» (មិនត្រូវនឹងក្រុមទាំង ២៩)`,
        };
      }
      return {
        isIncluded: true,
        category: 'k2_internal_op',
        categoryLabel: 'រាប់បញ្ចូល',
        reason: 'បានរាប់បញ្ចូលជាសន្និធិដើមគ្រាក្រុម (Base Opening)',
      };
    }

    // 6. Check K2/K1 Internal Warehouse Operations
    if (rec.operationType === 'openK1') {
      return {
        isIncluded: false,
        category: 'k2_internal_op',
        categoryLabel: 'បញ្ចូលស្តុកក២ (ក១)',
        reason: 'ការបញ្ចូលស្តុកក២ពីក១ (ជាប្រតិបត្តិការឃ្លាំងក២ មិនពាក់ព័ន្ធក្រុម)',
      };
    }
    if (rec.operationType === 'oldStockK2') {
      return {
        isIncluded: false,
        category: 'k2_internal_op',
        categoryLabel: 'ស្តុកចាស់ ក២',
        reason: 'ស្តុកចាស់ក២ (ជាប្រតិបត្តិការឃ្លាំងក២ មិនពាក់ព័ន្ធក្រុម)',
      };
    }
    if (rec.operationType === 'testPrintK2') {
      return {
        isIncluded: false,
        category: 'k2_internal_op',
        categoryLabel: 'ទិដ្ឋាការសាកក២',
        reason: 'ទិដ្ឋាការសាកក២ (ជាប្រតិបត្តិការឃ្លាំងក២ មិនពាក់ព័ន្ធក្រុម)',
      };
    }
    if (rec.operationType === 'damaged' && !rec.visaTeamRobokName && !rec.sourceFrom?.includes('ក្រុម')) {
      return {
        isIncluded: false,
        category: 'k2_internal_op',
        categoryLabel: 'ទិដ្ឋាការខូចក២',
        reason: 'ទិដ្ឋាការខូចក២ (ជាប្រតិបត្តិការឃ្លាំងក២ មិនពាក់ព័ន្ធក្រុម)',
      };
    }

    // 7. Team Operations
    const teamOps = ['issueTeam', 'useTeam', 'transferTeam', 'damagedTeam', 'damaged', 'returnTeam', 'missingTeam'];
    if (teamOps.includes(rec.operationType)) {
      if (!effectiveTeam) {
        return {
          isIncluded: false,
          category: 'unmatched_team',
          categoryLabel: 'មិនស្គាល់ឈ្មោះក្រុម',
          reason: `មិនស្គាល់ឈ្មោះក្រុម ឬមិនបានជ្រើសរើសក្រុម («${rawTeam || 'ទទេ'}») មិនត្រូវនឹងក្រុមទាំង ២៩`,
        };
      }

      // Cutoff date rule: <= 2018-11-30 (exclude from modifying opening stock)
      if (rec.date && rec.date <= '2018-11-30') {
        return {
          isIncluded: false,
          category: 'pre_nov_2018',
          categoryLabel: 'ក្រោមថ្ងៃ 30-11-2018',
          reason: 'កាលបរិច្ឆេទក្រោម ឬស្មើថ្ងៃ ៣០-១១-២០១៨ (មិនយកមកបូកជាមួយស្តុកចាស់ក្រុម តាមលក្ខខណ្ឌកំណត់)',
        };
      }

      return {
        isIncluded: true,
        category: 'k2_internal_op',
        categoryLabel: 'រាប់បញ្ចូល',
        reason: 'បានរាប់បញ្ចូលក្នុងរបាយការណ៍ក្រុម',
      };
    }

    return {
      isIncluded: false,
      category: 'unknown_op',
      categoryLabel: 'ប្រតិបត្តិការផ្សេងៗ',
      reason: `ប្រភេទប្រតិបត្តិការ «${rec.operationType || 'ទទេ'}» មិនរាប់បញ្ចូលក្នុងរបាយការណ៍ក្រុម`,
    };
  };

  // Helper to match a stock record to a team name
  const isRecordForTeam = useMemo(() => {
    return (r: StockRecord, team: string) => {
      if (!team) return true;
      const cleanTeam = team.trim().toLowerCase();
      const rTeamName = (r.visaTeamRobokName || '').trim().toLowerCase();
      const rTeamId = (r.visaTeamRobokId || '').trim().toLowerCase();
      if (rTeamName === cleanTeam || rTeamId === cleanTeam) return true;
      if (rTeamName && cleanTeam && (rTeamName.includes(cleanTeam) || cleanTeam.includes(rTeamName))) return true;
      const matched = (categories?.visaTeamsRobok || []).find(
        (t) => t.name.trim().toLowerCase() === cleanTeam || t.id.toLowerCase() === cleanTeam
      );
      if (matched) {
        if (
          r.visaTeamRobokId === matched.id ||
          (r.visaTeamRobokName && r.visaTeamRobokName.trim().toLowerCase() === matched.name.trim().toLowerCase())
        ) {
          return true;
        }
      }
      return false;
    };
  }, [categories?.visaTeamsRobok]);

  // Helper to match a stock record to a recipient team name
  const isRecordForRecipientTeamMemo = useMemo(() => {
    return (r: StockRecord, team: string) => {
      if (!team) return false;
      const cleanTeam = team.trim().toLowerCase();
      const rTeamName = (r.recipientTeamName || '').trim().toLowerCase();
      const rTeamId = (r.recipientTeamId || '').trim().toLowerCase();
      if (rTeamName === cleanTeam || rTeamId === cleanTeam) return true;
      if (rTeamName && cleanTeam && (rTeamName.includes(cleanTeam) || cleanTeam.includes(rTeamName))) return true;
      const matched = (categories?.visaTeamsRobok || []).find(
        (t) => t.name.trim().toLowerCase() === cleanTeam || t.id.toLowerCase() === cleanTeam
      );
      if (matched) {
        if (
          r.recipientTeamId === matched.id ||
          (r.recipientTeamName && r.recipientTeamName.trim().toLowerCase() === matched.name.trim().toLowerCase())
        ) {
          return true;
        }
      }
      return false;
    };
  }, [categories?.visaTeamsRobok]);

  // Merge stockRecords with any un-mirrored daily team usage operations from app_daily_team_operations_v5
  const synthesizedDailyTeamRecords = useMemo(() => {
    try {
      const savedOps = localStorage.getItem('app_daily_team_operations_v5');
      if (!savedOps) return [];
      const parsed: any[] = JSON.parse(savedOps);
      if (!Array.isArray(parsed) || parsed.length === 0) return [];

      const existingKeySet = new Set<string>();
      stockRecords.forEach((sr) => {
        if (
          sr.operationType === 'useTeam' ||
          sr.operationType === 'use_team' ||
          sr.sourceFrom?.includes('ប្រើប្រាស់')
        ) {
          const dIso = normalizeDateToISO(sr.date || '');
          const tNorm = normalizeTeamName(sr.visaTeamRobokName || sr.recipientTeamName || (sr as any).teamName || '');
          const vNorm = normalizeVisaType(sr.visaType);
          if (dIso && tNorm && vNorm) {
            existingKeySet.add(`${dIso}_${tNorm}_${vNorm}`);
          }
        }
      });

      const synthesized: StockRecord[] = [];
      parsed.forEach((dRec) => {
        if (!dRec || !dRec.values || !dRec.date || !dRec.teamName) return;
        const normDate = normalizeDateToISO(dRec.date);
        const normTeam = normalizeTeamName(dRec.teamName);
        const isCea = dRec.categoryType === 'cEA' || (dRec as any).sourceFrom === 'cEA';
        const stType = 'sticker';

        Object.keys(dRec.values).forEach((vtKey) => {
          const val = dRec.values[vtKey];
          if (!val) return;
          const entries = Array.isArray(val.entries) && val.entries.length > 0
            ? val.entries
            : [{ quantity: val.quantity, startSerial: val.startSerial, endSerial: val.endSerial, oldCode: val.oldCode }];

          entries.forEach((e: any, idx: number) => {
            const q = parseInt(e?.quantity || '', 10) || 0;
            if (q <= 0) return;
            const vNorm = normalizeVisaType(vtKey);
            const lookupKey = `${normDate}_${normTeam}_${vNorm}`;
            if (existingKeySet.has(lookupKey)) return;

            synthesized.push({
              id: `stock-dtr-auto-${dRec.id}-${vtKey}-${idx}`,
              stockType: stType,
              operationType: 'useTeam',
              date: dRec.date,
              time: '00:00',
              visaType: vtKey,
              quantityBundles: q,
              totalSheets: q,
              startSerial: e?.startSerial || '',
              endSerial: e?.endSerial || '',
              visaTeamRobokName: dRec.teamName,
              sourceFrom: isCea ? 'cEA' : 'Sticker',
              remarks: 'កត់ត្រាពីប្រតិបត្តិការប្រចាំថ្ងៃ',
              createdAt: dRec.createdAt || new Date().toISOString(),
            });
          });
        });
      });

      return synthesized;
    } catch {
      return [];
    }
  }, [stockRecords]);

  // Filter records for current stockType (defaulting undefined stockType to sticker)
  // For non-Secondary users, restrict records strictly to their assigned team (or records where they are the recipient of a transfer)
  const currentRecords = useMemo(() => {
    const combinedStock = synthesizedDailyTeamRecords.length > 0
      ? [...stockRecords, ...synthesizedDailyTeamRecords]
      : stockRecords;

    let records = combinedStock.filter(
      (r) => r.stockType === stockType || (!r.stockType && stockType === 'sticker')
    );
    if (!isSecondary && assignedTeam) {
      records = records.filter((r) => isRecordForTeam(r, assignedTeam) || isRecordForRecipientTeamMemo(r, assignedTeam));
    }
    return records;
  }, [stockRecords, synthesizedDailyTeamRecords, stockType, isSecondary, assignedTeam, isRecordForTeam, isRecordForRecipientTeamMemo]);

  // Map of inclusion status for current records
  const inclusionStatusMap = useMemo(() => {
    const map = new Map<string, { isIncluded: boolean; reason: string; category: 'unmatched_team' | 'unmatched_visa_type' | 'k2_internal_op' | 'pre_nov_2018' | 'zero_qty' | 'cea_evisa' | 'unknown_op'; categoryLabel: string }>();
    currentRecords.forEach((r) => {
      map.set(r.id, checkTeamReportInclusion(r, reportTeamsList));
    });
    return map;
  }, [currentRecords, reportTeamsList]);

  // List of all uncounted records for current stockType
  const uncountedInTeamReportRecords = useMemo(() => {
    return currentRecords.filter((r) => {
      const status = inclusionStatusMap.get(r.id);
      return status && !status.isIncluded;
    });
  }, [currentRecords, inclusionStatusMap]);

  // Category counts for uncounted records
  const uncountedCategoryCounts = useMemo(() => {
    const counts: Record<string, number> = {
      all: uncountedInTeamReportRecords.length,
      unmatched_team: 0,
      unmatched_visa_type: 0,
      k2_internal_op: 0,
      pre_nov_2018: 0,
      zero_qty: 0,
      cea_evisa: 0,
      unknown_op: 0,
    };
    uncountedInTeamReportRecords.forEach((r) => {
      const status = inclusionStatusMap.get(r.id);
      if (status && status.category && counts[status.category] !== undefined) {
        counts[status.category] += 1;
      }
    });
    return counts;
  }, [uncountedInTeamReportRecords, inclusionStatusMap]);

  const filteredRecords = useMemo(() => {
    return currentRecords.filter((r) => {
      // Filter by Uncounted in Team Report Finder
      if (filterUncountedInTeamReport) {
        const status = inclusionStatusMap.get(r.id);
        if (!status || status.isIncluded) {
          return false;
        }
        if (uncountedCategoryFilter !== 'all' && status.category !== uncountedCategoryFilter) {
          return false;
        }
      }

      // Filter by operation type
      if (filterOperationType !== 'all') {
        if (!matchOperation(r, filterOperationType)) {
          return false;
        }
      }

      // Filter by Date Range (startDate and endDate)
      if (filterStartDate) {
        if (!r.date || r.date < filterStartDate) {
          return false;
        }
      }
      if (filterEndDate) {
        if (!r.date || r.date > filterEndDate) {
          return false;
        }
      }

      if (!deferredSearchTerm) return true;
      const term = deferredSearchTerm.toLowerCase();
      const isK1 =
        r.operationType === 'testPrintK2'
          ? 'ទិដ្ឋាការសាកក២'
          : r.operationType === 'damaged'
          ? 'ទិដ្ឋាការខូចក២'
          : r.operationType === 'damagedTeam'
          ? 'ទិដ្ឋាការខូចក្រុម'
          : r.operationType === 'missingTeam'
          ? 'ទិដ្ឋាការខ្វះក្រុម'
          : r.operationType === 'returnTeam'
          ? 'ទិដ្ឋាការបង្វិលពីក្រុម'
          : r.operationType === 'transferTeam'
          ? 'ផ្ទេរការប្រើប្រាស់ក្រុម'
          : r.operationType === 'oldStockK2' || (r.operationType === 'openK1' && r.sourceFrom?.includes('ស្តុកចាស់'))
          ? 'ស្តុកចាស់ ក២'
          : r.operationType === 'oldStockTeam' || ((r.operationType === 'issueTeam' || r.operationType === 'useTeam') && r.sourceFrom?.includes('ស្តុកចាស់'))
          ? 'ស្តុកចាស់ក្រុម'
          : r.operationType === 'openK1'
          ? 'ការបញ្ចូលស្តុក'
          : r.operationType === 'useTeam'
          ? 'ការប្រើប្រាស់តាមក្រុម'
          : r.operationType || 'ការបើកផ្តល់តាមក្រុម';
      return (
        isK1.toLowerCase().includes(term) ||
        (r.operationType && r.operationType.toLowerCase().includes(term)) ||
        (r.sourceFrom && r.sourceFrom.toLowerCase().includes(term)) ||
        (r.visaTeamRobokName && r.visaTeamRobokName.toLowerCase().includes(term)) ||
        (r.visaType && r.visaType.toLowerCase().includes(term)) ||
        (r.requesterName && r.requesterName.toLowerCase().includes(term)) ||
        (r.collectorName && r.collectorName.toLowerCase().includes(term)) ||
        (r.date && r.date.includes(term))
      );
    });
  }, [
    currentRecords,
    filterUncountedInTeamReport,
    inclusionStatusMap,
    uncountedCategoryFilter,
    filterOperationType,
    filterStartDate,
    filterEndDate,
    deferredSearchTerm,
  ]);

  // Table pagination calculations
  const totalTablePages = Math.ceil(filteredRecords.length / tablePageSize) || 1;
  const safeCurrentTablePage = Math.min(Math.max(1, tablePage), totalTablePages);
  const paginatedRecords = useMemo(() => {
    if (tablePageSize >= 999999) return filteredRecords;
    const start = (safeCurrentTablePage - 1) * tablePageSize;
    return filteredRecords.slice(start, start + tablePageSize);
  }, [filteredRecords, safeCurrentTablePage, tablePageSize]);

  // Calculate totals
  const totalK1 = currentRecords
    .filter((r) => r.operationType === 'openK1' || r.operationType === 'oldStockK2')
    .reduce((sum, r) => sum + r.quantityBundles, 0);

  const totalIssueTeam = currentRecords
    .filter((r) => r.operationType === 'issueTeam' || r.operationType === 'oldStockTeam')
    .reduce((sum, r) => sum + r.quantityBundles, 0);

  const totalTeamUsed = currentRecords
    .filter((r) => r.operationType === 'useTeam')
    .reduce((sum, r) => sum + r.quantityBundles, 0);

  // Team specific calculations (when in useTeam mode or team selected)
  const targetTeamFilter = !isSecondary ? assignedTeam : visaTeamRobokId || selectedTeamName;

  const teamIssuedStock = currentRecords
    .filter((r) => {
      const isDirectIssue = (r.operationType === 'issueTeam' || r.operationType === 'oldStockTeam') && (!targetTeamFilter || isRecordForTeam(r, targetTeamFilter));
      const isTransferIn = r.operationType === 'transferTeam' && targetTeamFilter && isRecordForRecipientTeam(r, targetTeamFilter);
      return isDirectIssue || isTransferIn;
    })
    .reduce((sum, r) => sum + (r.quantityBundles || r.totalSheets || 0), 0);

  const teamUsedStock = currentRecords
    .filter(
      (r) =>
        (r.operationType === 'useTeam' || r.operationType === 'damagedTeam' || r.operationType === 'missingTeam' || r.operationType === 'returnTeam' || r.operationType === 'transferTeam') &&
        (!targetTeamFilter || isRecordForTeam(r, targetTeamFilter))
    )
    .reduce((sum, r) => sum + (r.quantityBundles || r.totalSheets || 0), 0);

  // Sticker Stock Data view for Office (isSecondary) - in ទម្រង់កត់ត្រាស្តុកសន្លឹក & ទិន្នន័យសន្លឹកទិដ្ឋាការ
  const isStickerOfficeData = isSecondary && stockType === 'sticker' && operationType !== 'useTeam';

  const isRecordFromDec2018 = (r: StockRecord) => {
    if (!r.date) return true;
    const dIso = normalizeDateToISO(r.date);
    return !dIso || dIso >= '2018-12-01';
  };



  const isTeamViewActive = !isSecondary || operationType === 'useTeam' || Boolean(visaTeamRobokId);

  // Office stock calculations (01-Dec-2018 to now)
  const totalOldStockK2 = currentRecords
    .filter((r) => matchOperation(r, 'oldStockK2'))
    .reduce((sum, r) => sum + (r.quantityBundles || r.totalSheets || 0), 0);

  const totalOpenK1 = currentRecords
    .filter((r) => matchOperation(r, 'openK1') && isRecordFromDec2018(r))
    .reduce((sum, r) => sum + (r.quantityBundles || r.totalSheets || 0), 0);

  const totalReturnTeam = currentRecords
    .filter((r) => matchOperation(r, 'returnTeam') && isRecordFromDec2018(r))
    .reduce((sum, r) => sum + (r.quantityBundles || r.totalSheets || 0), 0);

  const totalIssueTeamK2 = currentRecords
    .filter((r) => matchOperation(r, 'issueTeam') && isRecordFromDec2018(r))
    .reduce((sum, r) => sum + (r.quantityBundles || r.totalSheets || 0), 0);

  const totalDamagedK2 = currentRecords
    .filter((r) => matchOperation(r, 'damaged') && isRecordFromDec2018(r))
    .reduce((sum, r) => sum + (r.quantityBundles || r.totalSheets || 0), 0);

  const totalMissingK2 = currentRecords
    .filter((r) => (r.operationType === 'missing' || r.operationType === 'missingTeam' || (r.operationType as string)?.includes('ខ្វះ')) && isRecordFromDec2018(r))
    .reduce((sum, r) => sum + (r.quantityBundles || r.totalSheets || 0), 0);

  const totalTestPrintK2 = currentRecords
    .filter((r) => matchOperation(r, 'testPrintK2') && isRecordFromDec2018(r))
    .reduce((sum, r) => sum + (r.quantityBundles || r.totalSheets || 0), 0);

  const totalReturnK1 = currentRecords
    .filter((r) => matchOperation(r, 'returnK1') && isRecordFromDec2018(r))
    .reduce((sum, r) => sum + (r.quantityBundles || r.totalSheets || 0), 0);

  const totalTransferUseTeam = currentRecords
    .filter((r) => isTransferRecord(r) && isRecordFromDec2018(r))
    .reduce((sum, r) => sum + (r.quantityBundles || r.totalSheets || 0), 0);

  const remainingOfficeStock =
    totalOldStockK2 +
    totalOpenK1 +
    totalReturnTeam -
    totalIssueTeamK2 -
    totalDamagedK2 -
    totalTestPrintK2 -
    totalReturnK1;

  // Team specific sticker calculations
  const teamOldStock = currentRecords
    .filter((r) => matchOperation(r, 'oldStockTeam') || ((r.operationType === 'issueTeam' || r.operationType === 'useTeam') && r.sourceFrom?.includes('ស្តុកចាស់')))
    .filter((r) => !targetTeamFilter || isRecordForTeam(r, targetTeamFilter))
    .reduce((sum, r) => sum + (r.quantityBundles || r.totalSheets || 0), 0);

  const teamOpenK1 = currentRecords
    .filter((r) => {
      if (!isRecordFromDec2018(r)) return false;
      const isIssue = r.operationType === 'issueTeam' || r.operationType === 'openK1';
      const isTransferIn = r.operationType === 'transferTeam' && targetTeamFilter && isRecordForRecipientTeam(r, targetTeamFilter);
      if (isIssue) {
        return !targetTeamFilter || isRecordForTeam(r, targetTeamFilter);
      }
      if (isTransferIn) {
        return true;
      }
      return false;
    })
    .reduce((sum, r) => sum + (r.quantityBundles || r.totalSheets || 0), 0);

  const teamReturnTeam = currentRecords
    .filter((r) => matchOperation(r, 'returnTeam') && isRecordFromDec2018(r))
    .filter((r) => !targetTeamFilter || isRecordForTeam(r, targetTeamFilter))
    .reduce((sum, r) => sum + (r.quantityBundles || r.totalSheets || 0), 0);

  const teamUseTeam = currentRecords
    .filter((r) => matchOperation(r, 'useTeam') && isRecordFromDec2018(r))
    .filter((r) => !targetTeamFilter || isRecordForTeam(r, targetTeamFilter))
    .filter((r) => !isCeaRecord(r) && r.stockType !== 'evisa')
    .reduce((sum, r) => sum + (r.quantityBundles || r.totalSheets || 0), 0);

  const teamDamaged = currentRecords
    .filter((r) => (matchOperation(r, 'damaged') || r.operationType === 'damagedTeam') && isRecordFromDec2018(r))
    .filter((r) => !targetTeamFilter || isRecordForTeam(r, targetTeamFilter))
    .reduce((sum, r) => sum + (r.quantityBundles || r.totalSheets || 0), 0);

  const teamMissing = currentRecords
    .filter((r) => (r.operationType === 'missingTeam' || r.operationType === 'missing' || (r.operationType as string)?.includes('ខ្វះ')) && isRecordFromDec2018(r))
    .filter((r) => !targetTeamFilter || isRecordForTeam(r, targetTeamFilter))
    .reduce((sum, r) => sum + (r.quantityBundles || r.totalSheets || 0), 0);

  const teamTestPrint = currentRecords
    .filter((r) => (matchOperation(r, 'testPrintK2') || r.operationType === 'testPrintTeam') && isRecordFromDec2018(r))
    .filter((r) => !targetTeamFilter || isRecordForTeam(r, targetTeamFilter))
    .reduce((sum, r) => sum + (r.quantityBundles || r.totalSheets || 0), 0);

  const teamTransfer = currentRecords
    .filter((r) => isTransferRecord(r) && isRecordFromDec2018(r))
    .filter((r) => !targetTeamFilter || isRecordForTeam(r, targetTeamFilter))
    .reduce((sum, r) => sum + (r.quantityBundles || r.totalSheets || 0), 0);

  const activeTransferVal = isTeamViewActive ? teamTransfer : totalTransferUseTeam;

  const teamRemainingStock =
    teamOldStock +
    teamOpenK1 -
    teamReturnTeam -
    teamUseTeam -
    teamDamaged -
    teamMissing -
    teamTransfer;

  const officeStickerIssuedToTeams = useMemo(() => {
    let issueRecords: Array<{
      id: string;
      vt: string;
      start: string;
      end: string;
      date: string;
      team: string;
      qty: number;
    }> = [];
    let transferRecords: Array<{
      vt: string;
      start: string;
      end: string;
      date: string;
      team: string;
      qty: number;
    }> = [];

    (currentRecords || []).forEach((rec) => {
      if (isCeaRecord(rec)) return;
      const recDate = normalizeDateToISO(rec.date || '');
      if (recDate && recDate <= '2018-11-30') return;

      const vt = normalizeVisaType(rec.visaType || '');
      const start = (rec.startSerial || '').trim();
      const end = (rec.endSerial || '').trim();
      const qty = Number(rec.quantityBundles || rec.totalSheets || (rec as any).quantity || 0);
      const op = (rec.operationType || '').toLowerCase();
      const team = resolveRecordTeamName(rec);

      const isTransferTeam =
        op === 'transferteam' ||
        op === 'transferuseteam' ||
        op === 'transfer' ||
        op.includes('transfer') ||
        op.includes('ផ្ទេរ') ||
        (rec.sourceFrom && rec.sourceFrom.includes('ផ្ទេរ')) ||
        (rec.notes && (rec.notes as string).includes('ផ្ទេរ'));

      const isIssueTeam = (op === 'issueteam' || op === 'issue' || op.includes('បើកផ្តល់')) && !rec.sourceFrom?.includes('ស្តុកចាស់');

      if (isTransferTeam) {
        transferRecords.push({ vt, start, end, date: recDate, team, qty });
      } else if (isIssueTeam) {
        issueRecords.push({ id: rec.id, vt, start, end, date: recDate, team, qty });
      }
    });

    // Subtract transfer quantities from matching issue records, ensuring Port vs Airport mismatch is NOT matched
    transferRecords.forEach((tr) => {
      let matched = issueRecords.find((ir) => {
        if (ir.vt !== tr.vt || ir.date !== tr.date || ir.qty <= 0) return false;

        // Team check: if issue record is Sihanoukville Port and transfer record is Sihanoukville Airport (or vice versa), do not match
        if (ir.team && tr.team) {
          const t1 = ir.team.toLowerCase();
          const t2 = tr.team.toLowerCase();
          if (
            (t1.includes('កំពង់ផែ') && t2.includes('អាកាស')) ||
            (t1.includes('អាកាស') && t2.includes('កំពង់ផែ'))
          ) {
            return false;
          }
        }

        if (tr.start && tr.end && ir.start && ir.end) {
          return ir.start === tr.start && ir.end === tr.end;
        }
        return true;
      });

      if (matched) {
        matched.qty -= tr.qty;
        if (matched.qty < 0) matched.qty = 0;
      }
    });

    return issueRecords.reduce((sum, ir) => sum + ir.qty, 0);
  }, [currentRecords]);

  const netIssueTeamK2 = stockType === 'sticker' ? officeStickerIssuedToTeams : (totalIssueTeamK2 - totalTransferUseTeam);
  const displayOldStock = isTeamViewActive ? teamOldStock : totalOldStockK2;
  const displayOpenStock = isTeamViewActive ? teamOpenK1 : totalOpenK1;
  const displayReturnStock = isTeamViewActive ? teamReturnTeam : totalReturnTeam;
  const displayUsageStock = isSecondary ? netIssueTeamK2 : (isTeamViewActive ? teamUseTeam : netIssueTeamK2);
  const displayDamagedStock = isTeamViewActive ? teamDamaged : totalDamagedK2;
  const displayMissingStockVal = isTeamViewActive ? teamMissing : totalMissingK2;
  const displayTestPrintStock = isTeamViewActive ? teamTestPrint : totalTestPrintK2;
  const displayRemainingStock = isTeamViewActive ? teamRemainingStock : remainingOfficeStock;

  const showSevenCards = stockType === 'sticker';

  // Card values and labels depending on operationType
  const card1Label =
    !isSecondary
      ? assignedTeam
        ? `ស្តុកបានទទួល (${assignedTeam})`
        : 'ស្តុកទទួល / បើកផ្តល់'
      : operationType === 'useTeam'
      ? selectedTeamName
        ? `ស្តុក ${selectedTeamName}`
        : 'ស្តុកក្រុមផ្តល់ទិដ្ឋាការ.របក'
      : 'សរុបបញ្ចូលស្តុកក២';

  const card1Value = !isSecondary || operationType === 'useTeam' ? teamIssuedStock : totalK1;

  const card2Label =
    !isSecondary
      ? assignedTeam
        ? `សរុបប្រើប្រាស់ / កាត់ (${assignedTeam})`
        : 'សរុបប្រើប្រាស់តាមក្រុម'
      : operationType === 'useTeam'
      ? selectedTeamName
        ? `សរុបប្រើប្រាស់ ${selectedTeamName}`
        : 'សរុបប្រើប្រាស់តាមក្រុម'
      : 'ការបើកផ្តល់តាមក្រុម';

  const card2Value = !isSecondary || operationType === 'useTeam' ? teamUsedStock : totalIssueTeam;

  const remainingStock =
    !isSecondary || operationType === 'useTeam'
      ? teamIssuedStock - teamUsedStock
      : stockType === 'sticker'
      ? remainingOfficeStock
      : totalK1 - totalIssueTeam;

  const handlePrint = () => {
    printA4Document('stock-manager-print-table', {
      orientation: 'landscape',
      scale: 0.95,
      documentTitle: `របាយការណ៍ស្តុក — ${stockType === 'evisa' ? 'ក្រដាសអនុម័តទិដ្ឋាការអេឡិចត្រូនិក' : 'សន្លឹកទិដ្ឋាការ'}`,
    });
  };

  const handleExportExcel = () => {
    if (filteredRecords.length === 0) {
      onShowToast('គ្មានទិន្នន័យសម្រាប់ទាញយកទេ!');
      return;
    }

    const exportData = filteredRecords.map((item, index) => ({
      'ល.រ': index + 1,
      'ប្រភេទស្តុក': item.stockType === 'evisa' ? 'ក្រដាសអនុម័ត' : 'សន្លឹកទិដ្ឋាការ',
      'ប្រភេទប្រតិបត្តិការ': getRecordOperationLabel(item),
      'បើកពី': getRecordCleanSource(item),
      'កាលបរិច្ឆេទ': item.date || '',
      'ម៉ោង': item.time || '',
      'ប្រភេទទិដ្ឋាការ': item.visaType || '',
      'ចំនួន (សន្លឹក/ដុំ)': item.quantityBundles || item.totalSheets || 0,
      'ចាប់ពី': item.startSerial || '',
      'ដល់លេខ': item.endSerial || '',
      'ក្រុមផ្តល់ទិដ្ឋាការ.របក': item.visaTeamRobokName || '',
      'ឋានន្តរស័ក្កិអ្នកស្នើសុំ': item.requestedRankName || '',
      'អ្នកស្នើសុំ': item.requesterName || '',
      'អ្នកមកបើក': item.collectorName || '',
      'តួនាទីអ្នកមកបើក': item.collectorRoleName || '',
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportData);

    worksheet['!cols'] = [
      { wch: 6 },  // ល.រ
      { wch: 15 }, // ប្រភេទស្តុក
      { wch: 22 }, // ប្រភេទប្រតិបត្តិការ
      { wch: 18 }, // បើកពី
      { wch: 14 }, // កាលបរិច្ឆេទ
      { wch: 10 }, // ម៉ោង
      { wch: 15 }, // ប្រភេទទិដ្ឋាការ
      { wch: 16 }, // ចំនួន (សន្លឹក/ដុំ)
      { wch: 15 }, // ចាប់ពី
      { wch: 15 }, // ដល់លេខ
      { wch: 25 }, // ក្រុមផ្តល់ទិដ្ឋាការ.របក
      { wch: 22 }, // ឋានន្តរស័ក្កិអ្នកស្នើសុំ
      { wch: 20 }, // អ្នកស្នើសុំ
      { wch: 20 }, // អ្នកមកបើក
      { wch: 25 }, // តួនាទីអ្នកមកបើក
    ];

    const workbook = XLSX.utils.book_new();
    const sheetTitle = stockType === 'evisa' ? 'ទិន្នន័យក្រដាសអនុម័ត' : 'ទិន្នន័យសន្លឹកទិដ្ឋាការ';
    XLSX.utils.book_append_sheet(workbook, worksheet, sheetTitle);

    const opKhmerSuffix = filterOperationType !== 'all' ? `_${opShortLabelMap[filterOperationType] || filterOperationType}` : '';
    const dateSuffix = filterStartDate && filterEndDate
      ? `_${filterStartDate}_to_${filterEndDate}`
      : filterStartDate
      ? `_from_${filterStartDate}`
      : filterEndDate
      ? `_to_${filterEndDate}`
      : '';
    const fileName = stockType === 'evisa'
      ? `ទិន្នន័យស្តុក_ក្រដាសអនុម័ត${opKhmerSuffix}${dateSuffix}_${new Date().toISOString().split('T')[0]}.xlsx`
      : `ទិន្នន័យស្តុក_សន្លឹកទិដ្ឋាការ${opKhmerSuffix}${dateSuffix}_${new Date().toISOString().split('T')[0]}.xlsx`;

    XLSX.writeFile(workbook, fileName);
    onShowToast(`បានទាញយកទិន្នន័យស្តុកចំនួន ${filteredRecords.length} ជួរជា Excel រួចរាល់!`);
  };

  const handleBatchImportRecords = (newRecords: StockRecord[]) => {
    if (onBatchImportStockRecords) {
      onBatchImportStockRecords(newRecords);
    } else {
      newRecords.forEach((rec) => {
        onAddStockRecord(rec);
      });
    }
  };

  const renderEVisaNavbar = () => (
    <div className="bg-white p-2 rounded-md border border-gray-200 shadow-2xs flex flex-wrap items-center gap-1.5 text-xs print:hidden">
      <div className="flex items-center gap-1.5 text-slate-700 font-bold px-2 py-1 mr-1 border-r border-gray-200">
        <FileCheck className="w-4 h-4 text-[#007bff]" />
        <span>ក្រដាសអនុម័ត:</span>
      </div>

      {isSecondary && (
        <button
          type="button"
          onClick={() => {
            setViewMode('form');
            setOperationType('openK1');
            setSourceFrom('ក១');
            setEditingRecordId(null);
          }}
          className={`px-3 py-1.5 rounded text-xs font-bold transition flex items-center gap-1.5 cursor-pointer ${
            viewMode === 'form' && operationType !== 'useTeam'
              ? 'bg-[#007bff] text-white shadow-2xs'
              : 'bg-gray-100 text-gray-700 hover:bg-gray-200'
          }`}
        >
          <FileText className="w-3.5 h-3.5" />
          <span>ទម្រង់កត់ត្រាស្តុកក្រដាស</span>
        </button>
      )}

      <button
        type="button"
        onClick={() => {
          setViewMode('form');
          setOperationType('useTeam');
          setSourceFrom('ប្រើប្រាស់តាមក្រុម');
          setEditingRecordId(null);
        }}
        className={`px-3 py-1.5 rounded text-xs font-bold transition flex items-center gap-1.5 cursor-pointer ${
          viewMode === 'form' && operationType === 'useTeam'
            ? 'bg-[#007bff] text-white shadow-2xs'
            : 'bg-gray-100 text-gray-700 hover:bg-gray-200'
        }`}
      >
        <UserCheck className="w-3.5 h-3.5" />
        <span>ទម្រង់ប្រើប្រាស់ក្រដាសតាមក្រុម</span>
      </button>

      <button
        type="button"
        onClick={() => {
          setViewMode('data');
          setEditingRecordId(null);
        }}
        className={`px-3 py-1.5 rounded text-xs font-bold transition flex items-center gap-1.5 cursor-pointer ${
          viewMode === 'data'
            ? 'bg-[#007bff] text-white shadow-2xs'
            : 'bg-gray-100 text-gray-700 hover:bg-gray-200'
        }`}
      >
        <Layers className="w-3.5 h-3.5" />
        <span>ទិន្នន័យក្រដាសអនុម័ត</span>
      </button>

      {isSecondary && (
        <>
          <div className="h-4 w-px bg-gray-200 mx-1 hidden sm:block" />

          <button
            type="button"
            onClick={() => {
              setViewMode('robokReport');
              setEditingRecordId(null);
            }}
            className={`px-3 py-1.5 rounded text-xs font-bold transition flex items-center gap-1.5 cursor-pointer ${
              viewMode === 'robokReport'
                ? 'bg-amber-600 text-white shadow-2xs'
                : 'bg-amber-50 hover:bg-amber-100 text-amber-900 border border-amber-300'
            }`}
          >
            <Layers className="w-3.5 h-3.5" />
            <span>របកក្រដាសអនុម័តតាមបណ្តាក្រុម</span>
          </button>

          <button
            type="button"
            onClick={() => {
              setViewMode('report');
              setEditingRecordId(null);
            }}
            className={`px-3 py-1.5 rounded text-xs font-bold transition flex items-center gap-1.5 cursor-pointer ${
              viewMode === 'report'
                ? 'bg-blue-600 text-white shadow-2xs'
                : 'bg-blue-50 hover:bg-blue-100 text-blue-900 border border-blue-300'
            }`}
          >
            <Calendar className="w-3.5 h-3.5" />
            <span>របាយការណ៍ក្រដាសអនុម័ត</span>
          </button>

          <button
            type="button"
            onClick={() => {
              setViewMode('singleRobokReport');
              setEditingRecordId(null);
            }}
            className={`px-3 py-1.5 rounded text-xs font-bold transition flex items-center gap-1.5 cursor-pointer ${
              viewMode === 'singleRobokReport'
                ? 'bg-purple-600 text-white shadow-2xs'
                : 'bg-purple-50 hover:bg-purple-100 text-purple-900 border border-purple-300'
            }`}
          >
            <FileCheck className="w-3.5 h-3.5" />
            <span>របកក្រដាសបង្ហាញតែ១ក្រុម</span>
          </button>
        </>
      )}
    </div>
  );

  if (viewMode === 'report') {
    return (
      <div className={`space-y-4 animate-fade ${containerWidthMode === 'full' ? 'w-full' : 'max-w-7xl mx-auto'}`}>
        {stockType === 'evisa' && renderEVisaNavbar()}
        <EVisaStockReport
          stockRecords={stockRecords}
          categories={categories}
          officers={officers}
          onClose={() => setViewMode('all')}
          onSwitchReport={(m) => setViewMode(m)}
          currentReportMode="report"
          hideSwitcher={stockType === 'evisa'}
        />
      </div>
    );
  }

  if (viewMode === 'robokReport') {
    return (
      <div className={`space-y-4 animate-fade ${containerWidthMode === 'full' ? 'w-full' : 'max-w-7xl mx-auto'}`}>
        {stockType === 'evisa' && renderEVisaNavbar()}
        <EVisaRobokReport
          stockRecords={stockRecords}
          categories={categories}
          officers={officers}
          onClose={() => setViewMode('all')}
          onSwitchReport={(m) => setViewMode(m)}
          currentReportMode="robokReport"
          hideSwitcher={stockType === 'evisa'}
        />
      </div>
    );
  }

  if (viewMode === 'singleRobokReport') {
    return (
      <div className={`space-y-4 animate-fade ${containerWidthMode === 'full' ? 'w-full' : 'max-w-7xl mx-auto'}`}>
        {stockType === 'evisa' && renderEVisaNavbar()}
        <EVisaSingleRobokReport
          stockRecords={stockRecords}
          categories={categories}
          officers={officers}
          onClose={() => setViewMode('all')}
          onSwitchReport={(m) => setViewMode(m)}
          currentReportMode="singleRobokReport"
          hideSwitcher={stockType === 'evisa'}
        />
      </div>
    );
  }

  if (viewMode === 'stickerOfficeReport') {
    return (
      <StickerOfficeStockReport
        stockRecords={stockRecords}
        categories={categories}
        officers={officers}
        currentRole={currentRole}
        userName={userName}
        onClose={() => setViewMode('all')}
      />
    );
  }

  if (viewMode === 'stickerStatus') {
    return (
      <StickerVisaStatus
        stockRecords={stockRecords}
        categories={categories}
        officers={officers}
        currentRole={currentRole}
        userName={userName}
        onNavigateToForm={() => {
          setStockType('sticker');
          setViewMode('form');
        }}
        onClose={() => setViewMode('all')}
      />
    );
  }

  if (viewMode === 'handoverWorkspace') {
    return (
      <div className={`space-y-4 animate-fade ${containerWidthMode === 'full' ? 'w-full' : 'max-w-7xl mx-auto'}`}>
        {stockType === 'evisa' ? (
          <div className="space-y-2">
            {renderEVisaNavbar()}
            <div className="bg-white p-2 rounded-md border border-gray-200 shadow-2xs flex flex-wrap items-center justify-between gap-2 text-xs print:hidden">
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => {
                    setSelectedPdfRecord(null);
                    setViewMode('data');
                  }}
                  className="px-3 py-1.5 rounded text-xs font-bold bg-blue-50 text-blue-700 hover:bg-blue-100 border border-blue-200 flex items-center gap-1.5 cursor-pointer transition"
                >
                  <ArrowLeft className="w-3.5 h-3.5" />
                  <span>ត្រឡប់ទៅតារាងក្រដាសអនុម័ត</span>
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setSelectedPdfRecord(null);
                    setViewMode('form');
                    setOperationType('issueTeam');
                  }}
                  className="px-3 py-1.5 rounded text-xs font-bold bg-emerald-50 text-emerald-700 hover:bg-emerald-100 border border-emerald-200 flex items-center gap-1.5 cursor-pointer transition"
                >
                  <Plus className="w-3.5 h-3.5" />
                  <span>បន្តបើកផ្តល់ក្រដាសអនុម័តទៀត</span>
                </button>
              </div>
              <div className="text-xs text-gray-500 font-medium">
                ផ្ទាំងការងារលិខិតប្រគល់ទទួលក្រដាសអនុម័ត (Full Workspace Mode)
              </div>
            </div>
          </div>
        ) : (
          <div className="bg-white p-2 rounded-md border border-gray-200 shadow-2xs flex flex-wrap items-center justify-between gap-2 text-xs print:hidden">
            <div className="flex items-center gap-2">
              <button
                type="button"
                onClick={() => {
                  setSelectedPdfRecord(null);
                  setViewMode('data');
                }}
                className="px-3 py-1.5 rounded text-xs font-bold bg-blue-50 text-blue-700 hover:bg-blue-100 border border-blue-200 flex items-center gap-1.5 cursor-pointer transition"
              >
                <ArrowLeft className="w-3.5 h-3.5" />
                <span>ត្រឡប់ទៅតារាងសន្លឹកទិដ្ឋាការ</span>
              </button>
              <button
                type="button"
                onClick={() => {
                  setSelectedPdfRecord(null);
                  setViewMode('form');
                  setOperationType('issueTeam');
                }}
                className="px-3 py-1.5 rounded text-xs font-bold bg-emerald-50 text-emerald-700 hover:bg-emerald-100 border border-emerald-200 flex items-center gap-1.5 cursor-pointer transition"
              >
                <Plus className="w-3.5 h-3.5" />
                <span>បន្តបើកផ្តល់សន្លឹកទិដ្ឋាការទៀត</span>
              </button>
            </div>
            <div className="text-xs text-gray-500 font-medium">
              ផ្ទាំងការងារលិខិតប្រគល់ទទួលសន្លឹកទិដ្ឋាការ (Full Workspace Mode)
            </div>
          </div>
        )}
        <StockPdfModal
          displayMode="workspace"
          record={
            selectedPdfRecord
              ? {
                  ...selectedPdfRecord,
                  stockType: selectedPdfRecord.stockType || stockType,
                }
              : undefined
          }
          allRecords={stockRecords}
          categories={categories}
          officers={officers}
          userName={userName}
          onClose={() => {
            setSelectedPdfRecord(null);
            setViewMode('data');
          }}
        />
      </div>
    );
  }

  return (
    <div className={`space-y-5 animate-fade ${isExpandedWidth ? 'w-full' : 'max-w-7xl mx-auto'}`}>
      {/* Quick Navigation Tabs for Stock Pages */}
      {stockType === 'evisa' ? (
        renderEVisaNavbar()
      ) : (
        <div className="bg-white p-2 rounded-md border border-gray-200 shadow-2xs flex flex-wrap items-center gap-1.5 text-xs print:hidden">
          <div className="flex items-center gap-1.5 text-slate-700 font-bold px-2 py-1 mr-1 border-r border-gray-200">
            <Package className="w-4 h-4 text-[#007bff]" />
            <span>សន្លឹកទិដ្ឋាការ:</span>
          </div>

          <button
            type="button"
            onClick={() => {
              setViewMode('form');
              setEditingRecordId(null);
            }}
            className={`px-3 py-1.5 rounded text-xs font-bold transition flex items-center gap-1.5 cursor-pointer ${
              viewMode === 'form'
                ? 'bg-[#007bff] text-white shadow-2xs'
                : 'bg-gray-100 text-gray-700 hover:bg-gray-200'
            }`}
          >
            <FileText className="w-3.5 h-3.5" />
            <span>ទម្រង់កត់ត្រាស្តុកសន្លឹកទិដ្ឋាការ</span>
          </button>

          <div className="flex items-center gap-1">
            <button
              type="button"
              onClick={() => {
                setViewMode('data');
                setEditingRecordId(null);
              }}
              className={`px-3 py-1.5 rounded text-xs font-bold transition flex items-center gap-1.5 cursor-pointer ${
                viewMode === 'data'
                  ? 'bg-[#007bff] text-white shadow-2xs'
                  : 'bg-gray-100 text-gray-700 hover:bg-gray-200'
              }`}
            >
              <Layers className="w-3.5 h-3.5" />
              <span>ទិន្នន័យសន្លឹកទិដ្ឋាការ</span>
              {!isSecondary && assignedTeam && (
                <span className="text-[10px] bg-white/20 text-white px-1.5 py-0.5 rounded font-normal">
                  {assignedTeam}
                </span>
              )}
            </button>
          </div>

          {isSecondary && (
            <>
              <div className="h-4 w-px bg-gray-200 mx-1 hidden sm:block" />

              <button
                type="button"
                onClick={() => {
                  setViewMode('stickerOfficeReport');
                  setEditingRecordId(null);
                }}
                className="px-3 py-1.5 rounded text-xs font-bold bg-amber-50 hover:bg-amber-100 text-amber-900 border border-amber-300 transition flex items-center gap-1.5 cursor-pointer"
              >
                <FileSpreadsheet className="w-3.5 h-3.5 text-amber-700" />
                <span>របាយការណ៍ស្តុកការិយាល័យ</span>
              </button>

              <button
                type="button"
                onClick={() => {
                  setViewMode('stickerStatus');
                  setEditingRecordId(null);
                }}
                className="px-3 py-1.5 rounded text-xs font-bold bg-blue-50 hover:bg-blue-100 text-blue-900 border border-blue-300 transition flex items-center gap-1.5 cursor-pointer"
              >
                <CheckCircle2 className="w-3.5 h-3.5 text-blue-700" />
                <span>ស្ថានភាពប្រើប្រាស់សន្លឹកទិដ្ឋាការ</span>
              </button>
            </>
          )}
        </div>
      )}


      {/* Summary Stat Cards */}
      {isStickerOfficeData ? (
        <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-8 gap-1 sm:gap-1.5 xl:gap-2 print:hidden">
          {/* 1. ស្តុកចាស់ ក២ */}
          <div
            onClick={() => setFilterOperationType(filterOperationType === 'oldStockK2' ? 'all' : 'oldStockK2')}
            className={`bg-white p-1 sm:p-1.5 xl:p-2 rounded-md border transition-all cursor-pointer shadow-2xs flex items-center gap-1 sm:gap-1.5 xl:gap-2 ${
              filterOperationType === 'oldStockK2'
                ? 'border-indigo-500 ring-2 ring-indigo-200 bg-indigo-50/30'
                : 'border-gray-200 hover:border-indigo-300 hover:shadow-xs'
            }`}
            title="ចុចដើម្បីត្រងទិន្នន័យ (ស្តុកចាស់ ក២)"
          >
            <div className="w-5 h-5 sm:w-6 sm:h-6 xl:w-7 xl:h-7 rounded bg-indigo-50 text-indigo-600 flex items-center justify-center shrink-0">
              <Archive className="w-3 h-3 sm:w-3.5 sm:h-3.5 xl:w-4 xl:h-4" />
            </div>
            <div className="min-w-0 flex-1">
              <p className="text-[10px] xl:text-[11px] text-gray-700 font-bold truncate leading-tight" title="ស្តុកចាស់ ក២">
                ស្តុកចាស់ ក២
              </p>
              <p className="text-[8px] xl:text-[9px] text-gray-400 font-medium truncate leading-tight" title="សន្និធិដើមគ្រា">សន្និធិដើមគ្រា</p>
              <p className="text-[11px] xl:text-[13px] font-bold text-gray-900 leading-tight mt-0.5 truncate" title={`${totalOldStockK2.toLocaleString()} សន្លឹក`}>
                {totalOldStockK2.toLocaleString()} <span className="text-[8px] xl:text-[9px] font-normal text-gray-500">សន្លឹក</span>
              </p>
            </div>
          </div>

          {/* 2. បញ្ចូលស្តុក */}
          <div
            onClick={() => setFilterOperationType(filterOperationType === 'openK1' ? 'all' : 'openK1')}
            className={`bg-white p-1 sm:p-1.5 xl:p-2 rounded-md border transition-all cursor-pointer shadow-2xs flex items-center gap-1 sm:gap-1.5 xl:gap-2 ${
              filterOperationType === 'openK1'
                ? 'border-blue-500 ring-2 ring-blue-200 bg-blue-50/30'
                : 'border-gray-200 hover:border-blue-300 hover:shadow-xs'
            }`}
            title="ចុចដើម្បីត្រងទិន្នន័យ (បញ្ចូលស្តុក)"
          >
            <div className="w-5 h-5 sm:w-6 sm:h-6 xl:w-7 xl:h-7 rounded bg-blue-50 text-[#007bff] flex items-center justify-center shrink-0">
              <Boxes className="w-3 h-3 sm:w-3.5 sm:h-3.5 xl:w-4 xl:h-4" />
            </div>
            <div className="min-w-0 flex-1">
              <p className="text-[10px] xl:text-[11px] text-gray-700 font-bold truncate leading-tight" title="បញ្ចូលស្តុក">
                បញ្ចូលស្តុក
              </p>
              <p className="text-[8px] xl:text-[9px] text-blue-600 font-medium truncate leading-tight" title="01-Dec-2018 to now">01-Dec-2018 to now</p>
              <p className="text-[11px] xl:text-[13px] font-bold text-gray-900 leading-tight mt-0.5 truncate" title={`${totalOpenK1.toLocaleString()} សន្លឹក`}>
                {totalOpenK1.toLocaleString()} <span className="text-[8px] xl:text-[9px] font-normal text-gray-500">សន្លឹក</span>
              </p>
            </div>
          </div>

          {/* 3. ទិដ្ឋាការបង្វិលពីក្រុម */}
          <div
            onClick={() => setFilterOperationType(filterOperationType === 'returnTeam' ? 'all' : 'returnTeam')}
            className={`bg-white p-1 sm:p-1.5 xl:p-2 rounded-md border transition-all cursor-pointer shadow-2xs flex items-center gap-1 sm:gap-1.5 xl:gap-2 ${
              filterOperationType === 'returnTeam'
                ? 'border-teal-500 ring-2 ring-teal-200 bg-teal-50/30'
                : 'border-gray-200 hover:border-teal-300 hover:shadow-xs'
            }`}
            title="ចុចដើម្បីត្រងទិន្នន័យ (ទិដ្ឋាការបង្វិលពីក្រុម)"
          >
            <div className="w-5 h-5 sm:w-6 sm:h-6 xl:w-7 xl:h-7 rounded bg-teal-50 text-teal-600 flex items-center justify-center shrink-0">
              <RotateCcw className="w-3 h-3 sm:w-3.5 sm:h-3.5 xl:w-4 xl:h-4" />
            </div>
            <div className="min-w-0 flex-1">
              <p className="text-[10px] xl:text-[11px] text-gray-700 font-bold truncate leading-tight" title="ទិដ្ឋាការបង្វិលពីក្រុម">
                ទិដ្ឋាការបង្វិលពីក្រុម
              </p>
              <p className="text-[8px] xl:text-[9px] text-teal-600 font-medium truncate leading-tight" title="01-Dec-2018 to now">01-Dec-2018 to now</p>
              <p className="text-[11px] xl:text-[13px] font-bold text-gray-900 leading-tight mt-0.5 truncate" title={`${totalReturnTeam.toLocaleString()} សន្លឹក`}>
                {totalReturnTeam.toLocaleString()} <span className="text-[8px] xl:text-[9px] font-normal text-gray-500">សន្លឹក</span>
              </p>
            </div>
          </div>

          {/* 4. បើកផ្តល់តាមក្រុម */}
          <div
            onClick={() => setFilterOperationType(filterOperationType === 'issueTeam' ? 'all' : 'issueTeam')}
            className={`bg-white p-1 sm:p-1.5 xl:p-2 rounded-md border transition-all cursor-pointer shadow-2xs flex items-center gap-1 sm:gap-1.5 xl:gap-2 ${
              filterOperationType === 'issueTeam'
                ? 'border-amber-500 ring-2 ring-amber-200 bg-amber-50/30'
                : 'border-gray-200 hover:border-amber-300 hover:shadow-xs'
            }`}
            title="ចុចដើម្បីត្រងទិន្នន័យ (បើកផ្តល់តាមក្រុម)"
          >
            <div className="w-5 h-5 sm:w-6 sm:h-6 xl:w-7 xl:h-7 rounded bg-amber-50 text-amber-600 flex items-center justify-center shrink-0">
              <Layers className="w-3 h-3 sm:w-3.5 sm:h-3.5 xl:w-4 xl:h-4" />
            </div>
            <div className="min-w-0 flex-1">
              <p className="text-[10px] xl:text-[11px] text-gray-700 font-bold truncate leading-tight" title="បើកផ្តល់តាមក្រុម">
                បើកផ្តល់តាមក្រុម
              </p>
              <p className="text-[8px] xl:text-[9px] text-amber-600 font-medium truncate leading-tight" title="01-Dec-2018 to now">01-Dec-2018 to now</p>
              <p className="text-[11px] xl:text-[13px] font-bold text-gray-900 leading-tight mt-0.5 truncate" title={`${totalIssueTeamK2.toLocaleString()} សន្លឹក`}>
                {totalIssueTeamK2.toLocaleString()} <span className="text-[8px] xl:text-[9px] font-normal text-gray-500">សន្លឹក</span>
              </p>
            </div>
          </div>

          {/* 5. ទិដ្ឋាការខូចក២ */}
          <div
            onClick={() => setFilterOperationType(filterOperationType === 'damaged' ? 'all' : 'damaged')}
            className={`bg-white p-1 sm:p-1.5 xl:p-2 rounded-md border transition-all cursor-pointer shadow-2xs flex items-center gap-1 sm:gap-1.5 xl:gap-2 ${
              filterOperationType === 'damaged'
                ? 'border-rose-500 ring-2 ring-rose-200 bg-rose-50/30'
                : 'border-gray-200 hover:border-rose-300 hover:shadow-xs'
            }`}
            title="ចុចដើម្បីត្រងទិន្នន័យ (ទិដ្ឋាការខូចក២)"
          >
            <div className="w-5 h-5 sm:w-6 sm:h-6 xl:w-7 xl:h-7 rounded bg-rose-50 text-rose-600 flex items-center justify-center shrink-0">
              <AlertTriangle className="w-3 h-3 sm:w-3.5 sm:h-3.5 xl:w-4 xl:h-4" />
            </div>
            <div className="min-w-0 flex-1">
              <p className="text-[10px] xl:text-[11px] text-gray-700 font-bold truncate leading-tight" title="ទិដ្ឋាការខូចក២">
                ទិដ្ឋាការខូចក២
              </p>
              <p className="text-[8px] xl:text-[9px] text-rose-600 font-medium truncate leading-tight" title="01-Dec-2018 to now">01-Dec-2018 to now</p>
              <p className="text-[11px] xl:text-[13px] font-bold text-gray-900 leading-tight mt-0.5 truncate" title={`${totalDamagedK2.toLocaleString()} សន្លឹក`}>
                {totalDamagedK2.toLocaleString()} <span className="text-[8px] xl:text-[9px] font-normal text-gray-500">សន្លឹក</span>
              </p>
            </div>
          </div>

          {/* 6. ទិដ្ឋាការសាកក២ */}
          <div
            onClick={() => setFilterOperationType(filterOperationType === 'testPrintK2' ? 'all' : 'testPrintK2')}
            className={`bg-white p-1 sm:p-1.5 xl:p-2 rounded-md border transition-all cursor-pointer shadow-2xs flex items-center gap-1 sm:gap-1.5 xl:gap-2 ${
              filterOperationType === 'testPrintK2'
                ? 'border-purple-500 ring-2 ring-purple-200 bg-purple-50/30'
                : 'border-gray-200 hover:border-purple-300 hover:shadow-xs'
            }`}
            title="ចុចដើម្បីត្រងទិន្នន័យ (ទិដ្ឋាការសាកក២)"
          >
            <div className="w-5 h-5 sm:w-6 sm:h-6 xl:w-7 xl:h-7 rounded bg-purple-50 text-purple-600 flex items-center justify-center shrink-0">
              <Printer className="w-3 h-3 sm:w-3.5 sm:h-3.5 xl:w-4 xl:h-4" />
            </div>
            <div className="min-w-0 flex-1">
              <p className="text-[10px] xl:text-[11px] text-gray-700 font-bold truncate leading-tight" title="ទិដ្ឋាការសាកក២">
                ទិដ្ឋាការសាកក២
              </p>
              <p className="text-[8px] xl:text-[9px] text-purple-600 font-medium truncate leading-tight" title="01-Dec-2018 to now">01-Dec-2018 to now</p>
              <p className="text-[11px] xl:text-[13px] font-bold text-gray-900 leading-tight mt-0.5 truncate" title={`${totalTestPrintK2.toLocaleString()} សន្លឹក`}>
                {totalTestPrintK2.toLocaleString()} <span className="text-[8px] xl:text-[9px] font-normal text-gray-500">សន្លឹក</span>
              </p>
            </div>
          </div>

          {/* 7. បង្វិលក១ */}
          <div
            onClick={() => setFilterOperationType(filterOperationType === 'returnK1' ? 'all' : 'returnK1')}
            className={`bg-white p-1 sm:p-1.5 xl:p-2 rounded-md border transition-all cursor-pointer shadow-2xs flex items-center gap-1 sm:gap-1.5 xl:gap-2 ${
              filterOperationType === 'returnK1'
                ? 'border-sky-500 ring-2 ring-sky-200 bg-sky-50/30'
                : 'border-gray-200 hover:border-sky-300 hover:shadow-xs'
            }`}
            title="ចុចដើម្បីត្រងទិន្នន័យ (បង្វិលក១)"
          >
            <div className="w-5 h-5 sm:w-6 sm:h-6 xl:w-7 xl:h-7 rounded bg-sky-50 text-sky-600 flex items-center justify-center shrink-0">
              <Undo2 className="w-3 h-3 sm:w-3.5 sm:h-3.5 xl:w-4 xl:h-4" />
            </div>
            <div className="min-w-0 flex-1">
              <p className="text-[10px] xl:text-[11px] text-gray-700 font-bold truncate leading-tight" title="បង្វិលក១">
                បង្វិលក១
              </p>
              <p className="text-[8px] xl:text-[9px] text-sky-600 font-medium truncate leading-tight" title="01-Dec-2018 to now">01-Dec-2018 to now</p>
              <p className="text-[11px] xl:text-[13px] font-bold text-gray-900 leading-tight mt-0.5 truncate" title={`${totalReturnK1.toLocaleString()} សន្លឹក`}>
                {totalReturnK1.toLocaleString()} <span className="text-[8px] xl:text-[9px] font-normal text-gray-500">សន្លឹក</span>
              </p>
            </div>
          </div>

          {/* 8. ស្តុកនៅសល់ */}
          <div
            onClick={() => setFilterOperationType('all')}
            className={`bg-white p-1 sm:p-1.5 xl:p-2 rounded-md border transition-all cursor-pointer shadow-2xs flex items-center gap-1 sm:gap-1.5 xl:gap-2 ${
              filterOperationType === 'all'
                ? 'border-emerald-500 ring-2 ring-emerald-200 bg-emerald-50/20'
                : 'border-gray-200 hover:border-emerald-300 hover:shadow-xs'
            }`}
            title="ចុចដើម្បីបង្ហាញទិន្នន័យទាំងអស់ (All)"
          >
            <div
              className={`w-5 h-5 sm:w-6 sm:h-6 xl:w-7 xl:h-7 rounded ${
                remainingOfficeStock >= 0 ? 'bg-emerald-50 text-emerald-600' : 'bg-rose-50 text-rose-600'
              } flex items-center justify-center shrink-0`}
            >
              <CheckCircle2 className="w-3 h-3 sm:w-3.5 sm:h-3.5 xl:w-4 xl:h-4" />
            </div>
            <div className="min-w-0 flex-1">
              <p className="text-[10px] xl:text-[11px] text-gray-700 font-bold truncate leading-tight" title="ស្តុកនៅសល់">
                ស្តុកនៅសល់
              </p>
              <p className="text-[8px] xl:text-[9px] text-emerald-600 font-medium truncate leading-tight" title="សន្និធិចុងគ្រា">សន្និធិចុងគ្រា</p>
              <p
                className={`text-[11px] xl:text-[13px] font-bold leading-tight mt-0.5 truncate ${
                  remainingOfficeStock >= 0 ? 'text-emerald-700' : 'text-rose-600'
                }`}
                title={`${remainingOfficeStock.toLocaleString()} សន្លឹក`}
              >
                {remainingOfficeStock.toLocaleString()}{' '}
                <span className="text-[8px] xl:text-[9px] font-normal text-gray-500">សន្លឹក</span>
              </p>
            </div>
          </div>
        </div>
      ) : showSevenCards ? (
        <div className={`grid grid-cols-2 sm:grid-cols-4 ${teamTransfer > 0 ? 'lg:grid-cols-8' : 'lg:grid-cols-7'} gap-2 sm:gap-2.5 print:hidden`}>
          {/* 1. ស្តុកចាស់ */}
          <div
            onClick={() => setFilterOperationType(filterOperationType === 'oldStockK2' ? 'all' : 'oldStockK2')}
            className={`bg-white p-2 sm:p-2.5 rounded-md border transition-all cursor-pointer shadow-2xs flex items-center gap-2 ${
              filterOperationType === 'oldStockK2'
                ? 'border-indigo-500 ring-2 ring-indigo-200 bg-indigo-50/30'
                : 'border-gray-200 hover:border-indigo-300 hover:shadow-xs'
            }`}
            title="ចុចដើម្បីត្រងទិន្នន័យ (ស្តុកចាស់)"
          >
            <div className="w-7 h-7 sm:w-8 sm:h-8 rounded bg-indigo-50 text-indigo-600 flex items-center justify-center shrink-0">
              <Archive className="w-3.5 h-3.5 sm:w-4 sm:h-4" />
            </div>
            <div className="min-w-0 flex-1">
              <p className="text-[11px] sm:text-xs text-gray-700 font-bold truncate leading-tight">
                ស្តុកចាស់
              </p>
              <p className="text-[9px] sm:text-[10px] text-gray-400 font-medium truncate leading-tight" title="សន្និធិដើមគ្រា">សន្និធិដើមគ្រា</p>
              <p className="text-xs sm:text-[13px] xl:text-sm font-bold text-gray-900 leading-tight mt-0.5 truncate" title={`${displayOldStock.toLocaleString()} សន្លឹក`}>
                {displayOldStock.toLocaleString()} <span className="text-[10px] font-normal text-gray-500">សន្លឹក</span>
              </p>
            </div>
          </div>

          {/* 2. បញ្ចូលស្តុក / បើកទទួល */}
          <div
            onClick={() => setFilterOperationType(filterOperationType === 'openK1' ? 'all' : 'openK1')}
            className={`bg-white p-2 sm:p-2.5 rounded-md border transition-all cursor-pointer shadow-2xs flex items-center gap-2 ${
              filterOperationType === 'openK1'
                ? 'border-blue-500 ring-2 ring-blue-200 bg-blue-50/30'
                : 'border-gray-200 hover:border-blue-300 hover:shadow-xs'
            }`}
            title="ចុចដើម្បីត្រងទិន្នន័យ"
          >
            <div className="w-7 h-7 sm:w-8 sm:h-8 rounded bg-blue-50 text-[#007bff] flex items-center justify-center shrink-0">
              <Boxes className="w-3.5 h-3.5 sm:w-4 sm:h-4" />
            </div>
            <div className="min-w-0 flex-1">
              <p className="text-[11px] sm:text-xs text-gray-700 font-bold truncate leading-tight">
                {isTeamViewActive ? 'បើកទទួល' : 'បញ្ចូលស្តុក'}
              </p>
              <p className="text-[9px] sm:text-[10px] text-blue-600 font-medium truncate leading-tight" title="01-Dec-2018 to now">01-Dec-2018 to now</p>
              <p className="text-xs sm:text-[13px] xl:text-sm font-bold text-gray-900 leading-tight mt-0.5 truncate" title={`${displayOpenStock.toLocaleString()} សន្លឹក`}>
                {displayOpenStock.toLocaleString()} <span className="text-[10px] font-normal text-gray-500">សន្លឹក</span>
              </p>
            </div>
          </div>

          {/* 3. ទិដ្ឋាការបង្វិល */}
          <div
            onClick={() => setFilterOperationType(filterOperationType === 'returnTeam' ? 'all' : 'returnTeam')}
            className={`bg-white p-2 sm:p-2.5 rounded-md border transition-all cursor-pointer shadow-2xs flex items-center gap-2 ${
              filterOperationType === 'returnTeam'
                ? 'border-teal-500 ring-2 ring-teal-200 bg-teal-50/30'
                : 'border-gray-200 hover:border-teal-300 hover:shadow-xs'
            }`}
            title="ចុចដើម្បីត្រងទិន្នន័យ (ទិដ្ឋាការបង្វិល)"
          >
            <div className="w-7 h-7 sm:w-8 sm:h-8 rounded bg-teal-50 text-teal-600 flex items-center justify-center shrink-0">
              <RotateCcw className="w-3.5 h-3.5 sm:w-4 sm:h-4" />
            </div>
            <div className="min-w-0 flex-1">
              <p className="text-[11px] sm:text-xs text-gray-700 font-bold truncate leading-tight">
                ទិដ្ឋាការបង្វិល
              </p>
              <p className="text-[9px] sm:text-[10px] text-teal-600 font-medium truncate leading-tight" title="01-Dec-2018 to now">01-Dec-2018 to now</p>
              <p className="text-xs sm:text-[13px] xl:text-sm font-bold text-gray-900 leading-tight mt-0.5 truncate" title={`${displayReturnStock.toLocaleString()} សន្លឹក`}>
                {displayReturnStock.toLocaleString()} <span className="text-[10px] font-normal text-gray-500">សន្លឹក</span>
              </p>
            </div>
          </div>

          {/* 4. ការប្រើប្រាស់ */}
          <div
            onClick={() => setFilterOperationType(filterOperationType === 'issueTeam' ? 'all' : 'issueTeam')}
            className={`bg-white p-2 sm:p-2.5 rounded-md border transition-all cursor-pointer shadow-2xs flex items-center gap-2 ${
              filterOperationType === 'issueTeam'
                ? 'border-amber-500 ring-2 ring-amber-200 bg-amber-50/30'
                : 'border-gray-200 hover:border-amber-300 hover:shadow-xs'
            }`}
            title="ចុចដើម្បីត្រងទិន្នន័យ"
          >
            <div className="w-7 h-7 sm:w-8 sm:h-8 rounded bg-amber-50 text-amber-600 flex items-center justify-center shrink-0">
              <Layers className="w-3.5 h-3.5 sm:w-4 sm:h-4" />
            </div>
            <div className="min-w-0 flex-1">
              <p className="text-[11px] sm:text-xs text-gray-700 font-bold truncate leading-tight">
                {isTeamViewActive ? 'ការប្រើប្រាស់' : 'បើកផ្តល់តាមក្រុម'}
              </p>
              <p className="text-[9px] sm:text-[10px] text-amber-600 font-medium truncate leading-tight" title="01-Dec-2018 to now">01-Dec-2018 to now</p>
              <p className="text-xs sm:text-[13px] xl:text-sm font-bold text-gray-900 leading-tight mt-0.5 truncate" title={`${displayUsageStock.toLocaleString()} សន្លឹក`}>
                {displayUsageStock.toLocaleString()} <span className="text-[10px] font-normal text-gray-500">សន្លឹក</span>
              </p>
            </div>
          </div>

          {/* 5. ទិដ្ឋាការខូច */}
          <div
            onClick={() => setFilterOperationType(filterOperationType === 'damaged' ? 'all' : 'damaged')}
            className={`bg-white p-2 sm:p-2.5 rounded-md border transition-all cursor-pointer shadow-2xs flex items-center gap-2 ${
              filterOperationType === 'damaged'
                ? 'border-rose-500 ring-2 ring-rose-200 bg-rose-50/30'
                : 'border-gray-200 hover:border-rose-300 hover:shadow-xs'
            }`}
            title="ចុចដើម្បីត្រងទិន្នន័យ (ទិដ្ឋាការខូច)"
          >
            <div className="w-7 h-7 sm:w-8 sm:h-8 rounded bg-rose-50 text-rose-600 flex items-center justify-center shrink-0">
              <AlertTriangle className="w-3.5 h-3.5 sm:w-4 sm:h-4" />
            </div>
            <div className="min-w-0 flex-1">
              <p className="text-[11px] sm:text-xs text-gray-700 font-bold truncate leading-tight">
                ទិដ្ឋាការខូច
              </p>
              <p className="text-[9px] sm:text-[10px] text-rose-600 font-medium truncate leading-tight" title="01-Dec-2018 to now">01-Dec-2018 to now</p>
              <p className="text-xs sm:text-[13px] xl:text-sm font-bold text-gray-900 leading-tight mt-0.5 truncate" title={`${displayDamagedStock.toLocaleString()} សន្លឹក`}>
                {displayDamagedStock.toLocaleString()} <span className="text-[10px] font-normal text-gray-500">សន្លឹក</span>
              </p>
            </div>
          </div>

          {/* 6. ទិដ្ឋាការខ្វះ */}
          <div
            onClick={() => setFilterOperationType(filterOperationType === 'missing' ? 'all' : 'missing')}
            className={`bg-white p-2 sm:p-2.5 rounded-md border transition-all cursor-pointer shadow-2xs flex items-center gap-2 ${
              filterOperationType === 'missing'
                ? 'border-orange-500 ring-2 ring-orange-200 bg-orange-50/30'
                : 'border-gray-200 hover:border-orange-300 hover:shadow-xs'
            }`}
            title="ចុចដើម្បីត្រងទិន្នន័យ (ទិដ្ឋាការខ្វះ)"
          >
            <div className="w-7 h-7 sm:w-8 sm:h-8 rounded bg-orange-50 text-orange-600 flex items-center justify-center shrink-0">
              <AlertTriangle className="w-3.5 h-3.5 sm:w-4 sm:h-4" />
            </div>
            <div className="min-w-0 flex-1">
              <p className="text-[11px] sm:text-xs text-gray-700 font-bold truncate leading-tight">
                ទិដ្ឋាការខ្វះ
              </p>
              <p className="text-[9px] sm:text-[10px] text-orange-600 font-medium truncate leading-tight" title="01-Dec-2018 to now">01-Dec-2018 to now</p>
              <p className="text-xs sm:text-[13px] xl:text-sm font-bold text-gray-900 leading-tight mt-0.5 truncate" title={`${displayMissingStockVal.toLocaleString()} សន្លឹក`}>
                {displayMissingStockVal.toLocaleString()} <span className="text-[10px] font-normal text-gray-500">សន្លឹក</span>
              </p>
            </div>
          </div>

                    {/* ផ្ទេរការប្រើប្រាស់ */}
          {teamTransfer > 0 && (
            <div
              onClick={() => setFilterOperationType(filterOperationType === "transferTeam" ? "all" : "transferTeam")}
              className={`bg-white p-2 sm:p-2.5 rounded-md border transition-all cursor-pointer shadow-2xs flex items-center gap-2 ${
                filterOperationType === "transferTeam"
                  ? "border-cyan-500 ring-2 ring-cyan-200 bg-cyan-50/30"
                  : "border-gray-200 hover:border-cyan-300 hover:shadow-xs"
              }`}
              title="ចុចដើម្បីត្រងទិន្នន័យ (ផ្ទេរការប្រើប្រាស់)"
            >
              <div className="w-7 h-7 sm:w-8 sm:h-8 rounded bg-cyan-50 text-cyan-600 flex items-center justify-center shrink-0">
                <RotateCcw className="w-3.5 h-3.5 sm:w-4 sm:h-4" />
              </div>
              <div className="min-w-0 flex-1">
                <p className="text-[11px] sm:text-xs text-gray-700 font-bold truncate leading-tight">
                  ផ្ទេរការប្រើប្រាស់
                </p>
                <p className="text-[9px] sm:text-[10px] text-cyan-600 font-medium truncate leading-tight" title="01-Dec-2018 to now">01-Dec-2018 to now</p>
                <p className="text-xs sm:text-[13px] xl:text-sm font-bold text-gray-900 leading-tight mt-0.5 truncate" title={`${teamTransfer.toLocaleString()} សន្លឹក`}>
                  {teamTransfer.toLocaleString()} <span className="text-[10px] font-normal text-gray-500">សន្លឹក</span>
                </p>
              </div>
            </div>
          )}

          {/* 7. ស្តុកនៅសល់ */}
          <div
            onClick={() => setFilterOperationType('all')}
            className={`bg-white p-2 sm:p-2.5 rounded-md border transition-all cursor-pointer shadow-2xs flex items-center gap-2 ${
              filterOperationType === 'all'
                ? 'border-emerald-500 ring-2 ring-emerald-200 bg-emerald-50/20'
                : 'border-gray-200 hover:border-emerald-300 hover:shadow-xs'
            }`}
            title="ចុចដើម្បីបង្ហាញទិន្នន័យទាំងអស់ (All)"
          >
            <div
              className={`w-7 h-7 sm:w-8 sm:h-8 rounded ${
                displayRemainingStock >= 0 ? 'bg-emerald-50 text-emerald-600' : 'bg-rose-50 text-rose-600'
              } flex items-center justify-center shrink-0`}
            >
              <CheckCircle2 className="w-3.5 h-3.5 sm:w-4 sm:h-4" />
            </div>
            <div className="min-w-0 flex-1">
              <p className="text-[11px] sm:text-xs text-gray-700 font-bold truncate leading-tight">
                ស្តុកនៅសល់
              </p>
              <p className="text-[9px] sm:text-[10px] text-emerald-600 font-medium truncate leading-tight" title="សន្និធិចុងគ្រា">សន្និធិចុងគ្រា</p>
              <p
                className={`text-xs sm:text-[13px] xl:text-sm font-bold leading-tight mt-0.5 truncate ${
                  displayRemainingStock >= 0 ? 'text-emerald-700' : 'text-rose-600'
                }`}
                title={`${displayRemainingStock.toLocaleString()} សន្លឹក`}
              >
                {displayRemainingStock.toLocaleString()}{' '}
                <span className="text-[10px] font-normal text-gray-500">សន្លឹក</span>
              </p>
            </div>
          </div>
        </div>
      ) : (
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 print:hidden">
          <div className="bg-white p-3.5 rounded-md border border-gray-200 shadow-2xs flex items-center gap-3">
            <div className="w-9 h-9 rounded bg-blue-50 text-[#007bff] flex items-center justify-center shrink-0">
              <Boxes className="w-5 h-5" />
            </div>
            <div>
              <p className="text-xs text-gray-500 font-medium">{card1Label}</p>
              <p className="text-base font-bold text-gray-800">{card1Value} {stockType === 'evisa' ? 'ដុំ' : 'សន្លឹក'}</p>
            </div>
          </div>

          <div className="bg-white p-3.5 rounded-md border border-gray-200 shadow-2xs flex items-center gap-3">
            <div className="w-9 h-9 rounded bg-amber-50 text-amber-600 flex items-center justify-center shrink-0">
              <Layers className="w-5 h-5" />
            </div>
            <div>
              <p className="text-xs text-gray-500 font-medium">{card2Label}</p>
              <p className="text-base font-bold text-gray-800">{card2Value} {stockType === 'evisa' ? 'ដុំ' : 'សន្លឹក'}</p>
            </div>
          </div>

          <div className="bg-white p-3.5 rounded-md border border-gray-200 shadow-2xs flex items-center gap-3">
            <div className="w-9 h-9 rounded bg-emerald-50 text-emerald-600 flex items-center justify-center shrink-0">
              <CheckCircle2 className="w-5 h-5" />
            </div>
            <div>
              <p className="text-xs text-gray-500 font-medium">ស្តុកនៅសល់</p>
              <p
                className={`text-base font-bold ${
                  remainingStock >= 0 ? 'text-emerald-700' : 'text-rose-600'
                }`}
              >
                {remainingStock} {stockType === 'evisa' ? 'ដុំ' : 'សន្លឹក'}
              </p>
            </div>
          </div>
        </div>
      )}

      {/* AdminLTE Main Input Form Card */}
      {(viewMode === 'all' || viewMode === 'form' || editingRecordId !== null) && (
        <div className={`bg-white rounded-md border ${editingRecordId ? 'border-amber-400 border-t-4 border-t-amber-500' : 'border-gray-300 border-t-4 border-t-[#007bff]'} shadow-xs overflow-hidden print:hidden`}>
        {/* Card Header */}
        <div className={`${editingRecordId ? 'bg-amber-600' : 'bg-[#007bff]'} text-white px-4 py-3 flex items-center justify-between transition-colors`}>
          <div className="flex items-center gap-2 font-bold text-sm">
            {editingRecordId ? <Pencil className="w-4 h-4 text-white" /> : <FileText className="w-4 h-4 text-white" />}
            <span>
              {editingRecordId
                ? 'កែប្រែទិន្នន័យស្តុក'
                : stockType === 'evisa'
                ? operationType === 'useTeam'
                  ? 'ទម្រង់ប្រើប្រាស់ក្រដាសតាមក្រុម'
                  : operationType === 'issueTeam'
                  ? 'ទម្រង់បើកផ្តល់ក្រដាសតាមក្រុម'
                  : 'ទម្រង់បញ្ចូលស្តុកក្រដាស (ក១/ន៨)'
                : !isSecondary
                ? operationType === 'useTeam'
                  ? `ទម្រង់ប្រើប្រាស់សន្លឹកទិដ្ឋាការ${assignedTeam ? ` (${assignedTeam})` : 'តាមក្រុម'}`
                  : operationType === 'oldStockTeam'
                  ? `ទម្រង់កត់ត្រាស្តុកចាស់${assignedTeam ? ` (${assignedTeam})` : 'តាមក្រុម'}`
                  : operationType === 'returnTeam'
                  ? (!isSecondary ? `ទម្រង់បង្វិលសន្លឹកទិដ្ឋាការទៅក២${assignedTeam ? ` (${assignedTeam})` : ''}` : `ទម្រង់បង្វិលសន្លឹកទិដ្ឋាការពីក្រុម${assignedTeam ? ` (${assignedTeam})` : 'តាមក្រុម'}`)
                  : operationType === 'transferTeam'
                  ? `ទម្រង់ផ្ទេរសន្លឹកទិដ្ឋាការ${assignedTeam ? ` (${assignedTeam})` : 'តាមក្រុម'}`
                  : operationType === 'damagedTeam'
                  ? `ទម្រង់កត់ត្រាសន្លឹកខូច${assignedTeam ? ` (${assignedTeam})` : 'តាមក្រុម'}`
                  : operationType === 'missingTeam'
                  ? `ទម្រង់កត់ត្រាសន្លឹកខ្វះ${assignedTeam ? ` (${assignedTeam})` : 'តាមក្រុម'}`
                  : `ទម្រង់កត់ត្រាស្តុកសន្លឹកទិដ្ឋាការ${assignedTeam ? ` (${assignedTeam})` : 'តាមក្រុម'}`
                : operationType === 'useTeam'
                ? 'ទម្រង់ប្រើប្រាស់ក្រដាសតាមក្រុម'
                : 'ទម្រង់កត់ត្រាស្តុកការិយាល័យ'}
            </span>
          </div>
          <div className="flex items-center gap-1.5 text-xs bg-white/20 px-2.5 py-1 rounded font-mono">
            <Calendar className="w-3.5 h-3.5" />
            <span>{formattedToday}</span>
          </div>
        </div>

        {/* Card Body */}
        <form onSubmit={handleSubmit} className="p-6 space-y-6">
          {/* Operation Type Selection Dropdown */}
          {!hideOperationSelector && (
            <div className="bg-gray-50 p-3.5 rounded-md border border-gray-200 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
              <label className="text-xs font-bold text-gray-800 flex items-center gap-2">
                <ListFilter className="w-4 h-4 text-[#007bff]" />
                <span>ជ្រើសរើសប្រភេទប្រតិបត្តិការស្តុក:</span>
              </label>

              {stockType === 'evisa' ? (
                operationType === 'useTeam' ? (
                  <select
                    value="useTeam"
                    disabled
                    className="border border-gray-300 rounded-sm px-3 py-1.5 text-xs font-bold text-gray-800 bg-gray-100 focus:outline-none transition cursor-default w-full sm:w-80 disabled:opacity-100 disabled:text-gray-800 disabled:bg-gray-100"
                    title="ទម្រង់ប្រើប្រាស់ក្រដាសតាមក្រុម (ថេរ)"
                  >
                    <option value="useTeam">ការប្រើប្រាស់តាមក្រុម</option>
                  </select>
                ) : (
                  <select
                    value={operationType}
                    onChange={(e) => {
                      const val = e.target.value;
                      setOperationType(val);
                      if (val === 'openK1') {
                        setSourceFrom('ក១');
                      } else if (val === 'issueTeam') {
                        setSourceFrom('បើកផ្តល់តាមក្រុម');
                      } else if (val === 'useTeam') {
                        setSourceFrom('ប្រើប្រាស់តាមក្រុម');
                      } else {
                        setSourceFrom(val);
                      }
                    }}
                    className="border border-gray-300 rounded-sm px-3 py-1.5 text-xs font-bold text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition cursor-pointer w-full sm:w-80"
                  >
                    {isSecondary ? (
                      <>
                        <option value="openK1">១. ការបញ្ចូលស្តុក (ក១/ន៨)</option>
                        <option value="issueTeam">២. ការបើកផ្តល់តាមក្រុម</option>
                      </>
                    ) : (
                      <>
                        <option value="useTeam">ការប្រើប្រាស់តាមក្រុម</option>
                      </>
                    )}
                  </select>
                )
              ) : (
                <select
                  value={operationType}
                  onChange={(e) => {
                    const val = e.target.value;
                    setOperationType(val);
                    if (val === 'oldStockK2') {
                      setSourceFrom('ស្តុកចាស់ ក២');
                    } else if (val === 'oldStockTeam') {
                      setSourceFrom('ស្តុកចាស់ក្រុម');
                    } else if (val === 'testPrintK2') {
                      setSourceFrom('ទិដ្ឋាការសាកក២');
                    } else if (val === 'damaged') {
                      setSourceFrom('ទិដ្ឋាការខូចក២');
                    } else if (val === 'damagedTeam') {
                      setSourceFrom('ទិដ្ឋាការខូចក្រុម');
                    } else if (val === 'missingTeam') {
                      setSourceFrom('ទិដ្ឋាការខ្វះក្រុម');
                    } else if (val === 'returnTeam') {
                      setSourceFrom(!isSecondary ? 'ទិដ្ឋាការបង្វិលទៅក២' : 'ទិដ្ឋាការបង្វិលពីក្រុម');
                    } else if (val === 'transferTeam') {
                      setSourceFrom('ផ្ទេរការប្រើប្រាស់ក្រុម');
                    } else if (val === 'useTeam') {
                      setSourceFrom('ប្រើប្រាស់តាមក្រុម');
                    } else if (val === 'issueTeam') {
                      setSourceFrom('បើកផ្តល់តាមក្រុម');
                    } else if (val === 'openK1') {
                      setSourceFrom('ក១');
                    } else if (val === 'returnK1') {
                      setSourceFrom('បង្វិលក១');
                    } else if (val && !['openK1', 'issueTeam', 'useTeam', 'oldStockK2', 'oldStockTeam', 'returnK1'].includes(val)) {
                      setSourceFrom(val);
                    }
                  }}
                  className="border border-gray-300 rounded-sm px-3 py-1.5 text-xs font-bold text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition cursor-pointer w-full sm:w-80"
                >
                  {isSecondary ? (
                    <>
                      <option value="openK1">១. ការបញ្ចូលស្តុក (ក១)</option>
                      <option value="oldStockK2">២. ស្តុកចាស់ ក២</option>
                      <option value="oldStockTeam">៣. ស្តុកចាស់ក្រុម</option>
                      <option value="issueTeam">៤. ការបើកផ្តល់តាមក្រុម</option>
                      <option value="testPrintK2">៥. ទិដ្ឋាការសាកក២</option>
                      <option value="damaged">៦. ទិដ្ឋាការខូចក២</option>
                      <option value="damagedTeam">៧. ទិដ្ឋាការខូចក្រុម</option>
                      <option value="missingTeam">៨. ទិដ្ឋាការខ្វះក្រុម</option>
                      <option value="returnTeam">៩. ទិដ្ឋាការបង្វិលពីក្រុម</option>
                      <option value="transferTeam">១០. ផ្ទេរការប្រើប្រាស់ក្រុម</option>
                      <option value="returnK1">១១. បង្វិលក១</option>
                    </>
                  ) : (
                    <>
                      <option value="oldStockTeam">១. ស្តុកចាស់ក្រុម</option>
                      <option value="returnTeam">២. ទិដ្ឋាការបង្វិលទៅក២</option>
                      <option value="transferTeam">៣. ផ្ទេរការប្រើប្រាស់ក្រុម</option>
                      <option value="damagedTeam">៤. ទិដ្ឋាការខូចក្រុម</option>
                      <option value="missingTeam">៥. ទិដ្ឋាការខ្វះក្រុម</option>
                    </>
                  )}
                </select>
              )}
            </div>
          )}

          {/* Form fields for "សន្លឹកទិដ្ឋាការស្អិត" (Sticker stock) */}
          {stockType === 'sticker' ? (
            <div className="space-y-5">
              {/* Row 1: ជួរទី១ (កាលបរិច្ឆេទ *, បើកពីក១) */}
              <div className="space-y-4">
                <h4 className="text-xs font-bold text-gray-800 uppercase tracking-wide border-b border-gray-200 pb-2 flex items-center justify-between">
                  <span>ជួរទី១៖ ព័ត៌មានប្រតិបត្តិការស្តុកសន្លឹកទិដ្ឋាការ</span>
                  <span className="text-[11px] text-blue-600 font-normal border border-blue-200 bg-blue-50 px-2 py-0.5 rounded">
                    {operationType === 'oldStockK2'
                      ? 'ស្តុកចាស់ ក២'
                      : operationType === 'oldStockTeam'
                      ? 'ស្តុកចាស់ក្រុម'
                      : operationType === 'testPrintK2'
                      ? 'ទិដ្ឋាការសាកក២ (បោះពុម្ពសាកល្បង)'
                      : operationType === 'damaged'
                      ? 'ទិដ្ឋាការខូចក២ (មិនបានការក២)'
                      : operationType === 'damagedTeam'
                      ? 'ទិដ្ឋាការខូចក្រុម'
                      : operationType === 'missingTeam'
                      ? 'ទិដ្ឋាការខ្វះក្រុម'
                      : operationType === 'returnTeam'
                      ? (!isSecondary ? 'ទិដ្ឋាការបង្វិលទៅក២' : 'ទិដ្ឋាការបង្វិលពីក្រុម')
                      : operationType === 'transferTeam'
                      ? 'ផ្ទេរការប្រើប្រាស់ក្រុម'
                      : operationType === 'returnK1'
                      ? 'បង្វិលក១'
                      : operationType === 'openK1'
                      ? 'ការបញ្ចូលស្តុកក២'
                      : operationType === 'issueTeam'
                      ? 'ការបើកផ្តល់តាមក្រុម'
                      : operationType}
                  </span>
                </h4>

                {operationType === 'openK1' && (
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    {/* កាលបរិច្ឆេទ */}
                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        កាលបរិច្ឆេទ <span className="text-red-500">*</span>
                      </label>
                      <CustomDatePicker
                        required
                        value={date}
                        onChange={(d) => setDate(d)}
                        className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                      />
                    </div>

                    {/* បើកពី: (default 'ក១', options: 'ក១', 'ស្តុកចាស់ ក២', 'ន៨') */}
                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        បើកពី <span className="text-red-500">*</span>
                      </label>
                      <select
                        required
                        value={sourceFrom || 'ក១'}
                        onChange={(e) => setSourceFrom(e.target.value)}
                        className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs font-semibold text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition cursor-pointer"
                      >
                        <option value="ក១">ក១ (ការបញ្ចូលស្តុកថ្មី)</option>
                        <option value="ស្តុកចាស់ ក២">ស្តុកចាស់ ក២ (សន្និធិដើមគ្រា)</option>
                        <option value="ន៨">ន៨</option>
                      </select>
                    </div>
                  </div>
                )}

                {operationType === 'oldStockK2' && (
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        កាលបរិច្ឆេទ <span className="text-red-500">*</span>
                      </label>
                      <CustomDatePicker
                        required
                        value={date}
                        onChange={(d) => setDate(d)}
                        className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        ប្រភពស្តុក / សម្គាល់ <span className="text-red-500">*</span>
                      </label>
                      <input
                        type="text"
                        required
                        value={sourceFrom || 'ស្តុកចាស់ ក២'}
                        onChange={(e) => setSourceFrom(e.target.value)}
                        placeholder="ស្តុកចាស់ ក២ (សន្និធិដើមគ្រា)"
                        className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                      />
                    </div>
                  </div>
                )}

                 {operationType === 'oldStockTeam' && (
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        កាលបរិច្ឆេទ <span className="text-red-500">*</span>
                      </label>
                      <CustomDatePicker
                        required
                        value={date}
                        onChange={(d) => setDate(d)}
                        className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1 flex items-center justify-between">
                        <span>ក្រុមផ្តល់ទិដ្ឋាការ.របក <span className="text-red-500">*</span></span>
                      </label>
                      <select
                        required
                        disabled={!isSecondary && Boolean(assignedTeam)}
                        value={visaTeamRobokId}
                        onChange={(e) => setVisaTeamRobokId(e.target.value)}
                        className={`w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs font-semibold text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none transition ${!isSecondary && assignedTeam ? 'bg-gray-100 cursor-not-allowed text-gray-700 font-semibold' : 'bg-white cursor-pointer'}`}
                      >
                        <option value="">-- ជ្រើសរើស ក្រុមផ្តល់ទិដ្ឋាការ.របក --</option>
                        {(filteredVisaTeamsRobok || []).map((vtr) => (
                          <option key={vtr.id} value={vtr.id}>{vtr.name}</option>
                        ))}
                      </select>
                    </div>
                  </div>
                )}

                {operationType === 'testPrintK2' && (
                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        ម៉ោង <span className="text-gray-400 font-normal">(លុបបើមិនត្រូវការ)</span>
                      </label>
                      <input
                        type="time"
                        value={time}
                        onChange={(e) => setTime(e.target.value)}
                        className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition font-mono"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        កាលបរិច្ឆេទ <span className="text-red-500">*</span>
                      </label>
                      <CustomDatePicker
                        required
                        value={date}
                        onChange={(d) => setDate(d)}
                        className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        ប្រភេទ / ប្រភពសាកល្បង
                      </label>
                      <input
                        type="text"
                        value={sourceFrom || 'ទិដ្ឋាការសាកក២'}
                        onChange={(e) => setSourceFrom(e.target.value)}
                        placeholder="ទិដ្ឋាការសាកក២ / បោះពុម្ពសាកល្បង"
                        className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                      />
                    </div>
                  </div>
                )}

                {operationType === 'damaged' && (
                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        ម៉ោង <span className="text-gray-400 font-normal">(លុបបើមិនត្រូវការ)</span>
                      </label>
                      <input
                        type="time"
                        value={time}
                        onChange={(e) => setTime(e.target.value)}
                        className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition font-mono"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        កាលបរិច្ឆេទ <span className="text-red-500">*</span>
                      </label>
                      <CustomDatePicker
                        required
                        value={date}
                        onChange={(d) => setDate(d)}
                        className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        មូលហេតុ / ចំណាំ
                      </label>
                      <input
                        type="text"
                        value={sourceFrom || 'ទិដ្ឋាការខូចក២'}
                        onChange={(e) => setSourceFrom(e.target.value)}
                        placeholder="ទិដ្ឋាការខូចក២ / មិនបានការក២"
                        className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                      />
                    </div>
                  </div>
                )}

                {operationType === 'returnTeam' && (
                  <div className="space-y-3">
                    <div className={`grid ${isSecondary ? 'grid-cols-1 sm:grid-cols-3' : 'grid-cols-2'} gap-3`}>
                      <div>
                        <label className="block text-xs font-bold text-gray-700 mb-1">
                          កាលបរិច្ឆេទ <span className="text-red-500">*</span>
                        </label>
                        <CustomDatePicker
                          required
                          value={date}
                          onChange={(d) => setDate(d)}
                          className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                        />
                      </div>

                      <div>
                        <label className="block text-xs font-bold text-gray-700 mb-1">
                          ម៉ោង <span className="text-gray-400 font-normal">(លុបបើមិនត្រូវការ)</span>
                        </label>
                        <input
                          type="time"
                          value={time}
                          onChange={(e) => setTime(e.target.value)}
                          className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition font-mono"
                        />
                      </div>

                      {isSecondary && (
                        <div>
                          <label className="block text-xs font-bold text-gray-700 mb-1 flex items-center justify-between">
                            <span>ក្រុមផ្តល់ទិដ្ឋាការ.របក (ក្រុមបង្វិល) <span className="text-red-500">*</span></span>
                          </label>
                          <select
                            required
                            value={visaTeamRobokId}
                            onChange={(e) => setVisaTeamRobokId(e.target.value)}
                            className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none transition bg-white cursor-pointer"
                          >
                            <option value="">-- ជ្រើសរើស ក្រុមផ្តល់ទិដ្ឋាការ.របក --</option>
                            {(filteredVisaTeamsRobok || []).map((vtr) => (
                              <option key={vtr.id} value={vtr.id}>
                                {vtr.name}
                              </option>
                            ))}
                          </select>
                        </div>
                      )}
                    </div>

                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                      <div>
                        <label className="block text-xs font-bold text-gray-700 mb-1">ឈ្មោះអ្នកបង្វិល / មកប្រគល់</label>
                        <input
                          type="text"
                          placeholder="បញ្ចូលឈ្មោះអ្នកបង្វិល"
                          value={requesterName}
                          onChange={(e) => setRequesterName(e.target.value)}
                          className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none transition"
                        />
                      </div>
                      <div>
                        <label className="block text-xs font-bold text-gray-700 mb-1">សម្គាល់ / មូលហេតុបង្វិល</label>
                        <input
                          type="text"
                          placeholder={!isSecondary ? 'ទិដ្ឋាការបង្វិលទៅក២' : 'ទិដ្ឋាការបង្វិលពីក្រុម'}
                          value={sourceFrom || (!isSecondary ? 'ទិដ្ឋាការបង្វិលទៅក២' : 'ទិដ្ឋាការបង្វិលពីក្រុម')}
                          onChange={(e) => setSourceFrom(e.target.value)}
                          className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none transition"
                        />
                      </div>
                    </div>
                  </div>
                )}

                
                {operationType === 'transferTeam' && (
                  <div className="space-y-3">
                    <div className={`grid ${isSecondary ? 'grid-cols-1 sm:grid-cols-3' : 'grid-cols-2'} gap-3`}>
                      <div>
                        <label className="block text-xs font-bold text-gray-700 mb-1">
                          កាលបរិច្ឆេទ <span className="text-red-500">*</span>
                        </label>
                        <CustomDatePicker
                          required
                          value={date}
                          onChange={(d) => setDate(d)}
                          className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                        />
                      </div>
                      <div>
                        <label className="block text-xs font-bold text-gray-700 mb-1">
                          ម៉ោង <span className="text-gray-400 font-normal">(លុបបើមិនត្រូវការ)</span>
                        </label>
                        <input
                          type="time"
                          value={time}
                          onChange={(e) => setTime(e.target.value)}
                          className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition font-mono"
                        />
                      </div>
                      {isSecondary && (
                        <div>
                          <label className="block text-xs font-bold text-gray-700 mb-1 flex items-center justify-between">
                            <span>ក្រុមផ្តល់ទិដ្ឋាការ.របក (ក្រុមផ្ទេរ) <span className="text-red-500">*</span></span>
                          </label>
                          <select
                            required
                            value={visaTeamRobokId}
                            onChange={(e) => setVisaTeamRobokId(e.target.value)}
                            className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none transition bg-white cursor-pointer"
                          >
                            <option value="">-- ជ្រើសរើស ក្រុមផ្តល់ទិដ្ឋាការ.របក --</option>
                            {(filteredVisaTeamsRobok || []).map((vtr) => (
                              <option key={vtr.id} value={vtr.id}>
                                {vtr.name}
                              </option>
                            ))}
                          </select>
                        </div>
                      )}
                    </div>
                    {/* New fields for Transfer Team rearranged in a single row */}
                    <div className="grid grid-cols-1 sm:grid-cols-4 gap-3 border-t border-dashed border-gray-200 pt-3">
                      <div>
                        <label className="block text-xs font-bold text-gray-700 mb-1">សម្គាល់ / ព័ត៌មានផ្ទេរ</label>
                        <input
                          type="text"
                          placeholder="ផ្ទេរការប្រើប្រាស់ក្រុម"
                          value={sourceFrom || 'ផ្ទេរការប្រើប្រាស់ក្រុម'}
                          onChange={(e) => setSourceFrom(e.target.value)}
                          className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none transition"
                        />
                      </div>

                      <div>
                        <label className="block text-xs font-bold text-gray-700 mb-1">
                          ក្រុមដែលទទួលទិដ្ឋាការ <span className="text-red-500">*</span>
                        </label>
                        <select
                          required
                          value={recipientTeamId}
                          onChange={(e) => setRecipientTeamId(e.target.value)}
                          className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white cursor-pointer"
                        >
                          <option value="">-- ជ្រើសរើស ក្រុមដែលទទួល --</option>
                          {(availableRecipientTeams || []).map((vtr) => (
                            <option key={vtr.id} value={vtr.id}>
                              {vtr.name}
                            </option>
                          ))}
                        </select>
                      </div>

                      <div>
                        <label className="block text-xs font-bold text-gray-700 mb-1">
                          កាលបរិច្ឆេទឯកភាពពីនាយកដ្ឋាន <span className="text-red-500">*</span>
                        </label>
                        <CustomDatePicker
                          required
                          value={deptApprovalDate}
                          onChange={(d) => setDeptApprovalDate(d)}
                          className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                        />
                      </div>

                      <div className="flex items-center h-full pt-5">
                        <label className="flex items-center space-x-2 cursor-pointer select-none">
                          <input
                            type="checkbox"
                            checked={officeApproved}
                            onChange={(e) => setOfficeApproved(e.target.checked)}
                            className="w-4 h-4 text-blue-600 border-gray-300 rounded-sm focus:ring-[#007bff] cursor-pointer"
                          />
                          <span className="text-xs font-bold text-gray-700">ទទួលបានការអនុញ្ញាតពីការិយាល័យ</span>
                        </label>
                      </div>
                    </div>
                  </div>
                )}
                {operationType === 'damagedTeam' && (
                  <div className="space-y-3">
                    <div className={`grid ${isSecondary ? 'grid-cols-1 sm:grid-cols-3' : 'grid-cols-2'} gap-3`}>
                      <div>
                        <label className="block text-xs font-bold text-gray-700 mb-1">
                          កាលបរិច្ឆេទ <span className="text-red-500">*</span>
                        </label>
                        <CustomDatePicker
                          required
                          value={date}
                          onChange={(d) => setDate(d)}
                          className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                        />
                      </div>

                      <div>
                        <label className="block text-xs font-bold text-gray-700 mb-1">
                          ម៉ោង <span className="text-gray-400 font-normal">(លុបបើមិនត្រូវការ)</span>
                        </label>
                        <input
                          type="time"
                          value={time}
                          onChange={(e) => setTime(e.target.value)}
                          className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition font-mono"
                        />
                      </div>

                      {isSecondary && (
                        <div>
                          <label className="block text-xs font-bold text-gray-700 mb-1 flex items-center justify-between">
                            <span>ក្រុមផ្តល់ទិដ្ឋាការ.របក (ក្រុមខូច) <span className="text-red-500">*</span></span>
                          </label>
                          <select
                            required
                            value={visaTeamRobokId}
                            onChange={(e) => setVisaTeamRobokId(e.target.value)}
                            className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none transition bg-white cursor-pointer"
                          >
                            <option value="">-- ជ្រើសរើស ក្រុមផ្តល់ទិដ្ឋាការ.របក --</option>
                            {(categories.visaTeamsRobok || []).map((vtr) => (
                              <option key={vtr.id} value={vtr.id}>
                                {vtr.name}
                              </option>
                            ))}
                          </select>
                        </div>
                      )}
                    </div>

                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                      <div>
                        <label className="block text-xs font-bold text-gray-700 mb-1">ឈ្មោះមន្ត្រី ឬអ្នករាយការណ៍</label>
                        <input
                          type="text"
                          placeholder="បញ្ចូលឈ្មោះមន្ត្រី ឬអ្នករាយការណ៍"
                          value={requesterName}
                          onChange={(e) => setRequesterName(e.target.value)}
                          className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none transition"
                        />
                      </div>
                      <div>
                        <label className="block text-xs font-bold text-gray-700 mb-1">សម្គាល់ / មូលហេតុខូច</label>
                        <input
                          type="text"
                          placeholder="ទិដ្ឋាការខូចក្រុម"
                          value={sourceFrom || 'ទិដ្ឋាការខូចក្រុម'}
                          onChange={(e) => setSourceFrom(e.target.value)}
                          className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none transition"
                        />
                      </div>
                    </div>
                  </div>
                )}

                {operationType === 'missingTeam' && (
                  <div className="space-y-3">
                    <div className={`grid ${isSecondary ? 'grid-cols-1 sm:grid-cols-3' : 'grid-cols-2'} gap-3`}>
                      <div>
                        <label className="block text-xs font-bold text-gray-700 mb-1">
                          កាលបរិច្ឆេទ <span className="text-red-500">*</span>
                        </label>
                        <CustomDatePicker
                          required
                          value={date}
                          onChange={(d) => setDate(d)}
                          className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                        />
                      </div>

                      <div>
                        <label className="block text-xs font-bold text-gray-700 mb-1">
                          ម៉ោង <span className="text-gray-400 font-normal">(លុបបើមិនត្រូវការ)</span>
                        </label>
                        <input
                          type="time"
                          value={time}
                          onChange={(e) => setTime(e.target.value)}
                          className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition font-mono"
                        />
                      </div>

                      {isSecondary && (
                        <div>
                          <label className="block text-xs font-bold text-gray-700 mb-1 flex items-center justify-between">
                            <span>ក្រុមផ្តល់ទិដ្ឋាការ.របក (ក្រុមខ្វះ) <span className="text-red-500">*</span></span>
                          </label>
                          <select
                            required
                            value={visaTeamRobokId}
                            onChange={(e) => setVisaTeamRobokId(e.target.value)}
                            className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none transition bg-white cursor-pointer"
                          >
                            <option value="">-- ជ្រើសរើស ក្រុមផ្តល់ទិដ្ឋាការ.របក --</option>
                            {(categories.visaTeamsRobok || []).map((vtr) => (
                              <option key={vtr.id} value={vtr.id}>
                                {vtr.name}
                              </option>
                            ))}
                          </select>
                        </div>
                      )}
                    </div>

                    <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                      <div>
                        <label className="block text-xs font-bold text-gray-700 mb-1">ឈ្មោះមន្ត្រី ឬអ្នករាយការណ៍</label>
                        <input
                          type="text"
                          placeholder="បញ្ចូលឈ្មោះមន្ត្រី ឬអ្នករាយការណ៍"
                          value={requesterName}
                          onChange={(e) => setRequesterName(e.target.value)}
                          className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none transition"
                        />
                      </div>
                      <div>
                        <label className="block text-xs font-bold text-gray-700 mb-1">សម្គាល់ / មូលហេតុខ្វះ</label>
                        <input
                          type="text"
                          placeholder="ទិដ្ឋាការខ្វះក្រុម"
                          value={sourceFrom || 'ទិដ្ឋាការខ្វះក្រុម'}
                          onChange={(e) => setSourceFrom(e.target.value)}
                          className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none transition"
                        />
                      </div>
                    </div>
                  </div>
                )}

                {operationType === 'returnK1' && (
                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        កាលបរិច្ឆេទ <span className="text-red-500">*</span>
                      </label>
                      <CustomDatePicker
                        required
                        value={date}
                        onChange={(d) => setDate(d)}
                        className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        ម៉ោង <span className="text-gray-400 font-normal">(លុបបើមិនត្រូវការ)</span>
                      </label>
                      <input
                        type="time"
                        value={time}
                        onChange={(e) => setTime(e.target.value)}
                        className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition font-mono"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        ប្រភព / សម្គាល់ <span className="text-red-500">*</span>
                      </label>
                      <input
                        type="text"
                        required
                        value={sourceFrom || 'បង្វិលក១'}
                        onChange={(e) => setSourceFrom(e.target.value)}
                        placeholder="បង្វិលក១"
                        className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                      />
                    </div>
                  </div>
                )}

                {!['openK1', 'oldStockK2', 'oldStockTeam', 'issueTeam', 'useTeam', 'testPrintK2', 'damaged', 'returnTeam', 'transferTeam', 'damagedTeam', 'missingTeam', 'returnK1'].includes(operationType) && (
                  <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        ម៉ោង <span className="text-gray-400 font-normal">(លុបបើមិនត្រូវការ)</span>
                      </label>
                      <input
                        type="time"
                        value={time}
                        onChange={(e) => setTime(e.target.value)}
                        className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition font-mono"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        កាលបរិច្ឆេទ <span className="text-red-500">*</span>
                      </label>
                      <CustomDatePicker
                        required
                        value={date}
                        onChange={(d) => setDate(d)}
                        className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        ប្រភព / ប្រតិបត្តិការ
                      </label>
                      <input
                        type="text"
                        value={sourceFrom || operationType}
                        onChange={(e) => setSourceFrom(e.target.value)}
                        placeholder="បញ្ចូលប្រភព ឬឈ្មោះប្រតិបត្តិការ"
                        className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                      />
                    </div>
                  </div>
                )}

                {operationType === 'issueTeam' && (
                  <div className="space-y-3">
                    <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
                      {/* ម៉ោង */}
                      <div>
                        <label className="block text-xs font-bold text-gray-700 mb-1">
                          ម៉ោង <span className="text-gray-400 font-normal">(លុបបើមិនត្រូវការ)</span>
                        </label>
                        <input
                          type="time"
                          value={time}
                          onChange={(e) => setTime(e.target.value)}
                          className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition font-mono"
                        />
                      </div>

                      {/* កាលបរិច្ឆេទ */}
                      <div>
                        <label className="block text-xs font-bold text-gray-700 mb-1">
                          កាលបរិច្ឆេទ <span className="text-red-500">*</span>
                        </label>
                        <CustomDatePicker
                          required
                          value={date}
                          onChange={(d) => setDate(d)}
                          className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                        />
                      </div>

                      {/* ក្រុមផ្តល់ទិដ្ឋាការ.របក */}
                      <div>
                        <label className="block text-xs font-bold text-gray-700 mb-1 flex items-center justify-between">
                          <span>ក្រុមផ្តល់ទិដ្ឋាការ.របក <span className="text-red-500">*</span></span>
                          {!isSecondary && assignedTeam && (
                            <span className="text-[10px] text-amber-700 font-normal bg-amber-50 border border-amber-200 px-1.5 py-0.5 rounded">
                              🔒 {assignedTeam}
                            </span>
                          )}
                        </label>
                        <select
                          required
                          disabled={!isSecondary && Boolean(assignedTeam)}
                          value={visaTeamRobokId}
                          onChange={(e) => setVisaTeamRobokId(e.target.value)}
                          className={`w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none transition ${!isSecondary && assignedTeam ? 'bg-gray-100 cursor-not-allowed text-gray-700 font-semibold' : 'bg-white cursor-pointer'}`}
                        >
                          <option value="">-- ជ្រើសរើស ក្រុមផ្តល់ទិដ្ឋាការ.របក --</option>
                          {(categories.visaTeamsRobok || [])
                            .filter((vtr, idx) => {
                              if (vtr.id === visaTeamRobokId || vtr.name === visaTeamRobokId) return true;
                              const opt =
                                (vtr.id && teamTypeOptions[vtr.id]) ||
                                (vtr.name && teamTypeOptions[vtr.name]) ||
                                teamTypeOptions[idx];
                              if (!opt) return true;
                              return opt.sticker !== false;
                            })
                            .map((vtr) => (
                              <option key={vtr.id} value={vtr.id}>
                                {vtr.name}
                              </option>
                            ))}
                        </select>
                      </div>
                    </div>

                    <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-3 pt-2">
                      <div>
                        <label className="block text-xs font-bold text-gray-700 mb-1">ឋានន្តរស័ក្កិស្នើសុំ</label>
                        <select
                          value={requestedRankId}
                          onChange={(e) => setRequestedRankId(e.target.value)}
                          className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition cursor-pointer"
                        >
                          <option value="">-- ជ្រើសរើស --</option>
                          {(categories.ranks || []).map((r) => (
                            <option key={r.id} value={r.id}>{r.name}</option>
                          ))}
                        </select>
                      </div>

                      <div>
                        <label className="block text-xs font-bold text-gray-700 mb-1">ឈ្មោះស្នើសុំ</label>
                        <input
                          type="text"
                          placeholder="បញ្ចូលឈ្មោះអ្នកស្នើសុំ"
                          value={requesterName}
                          onChange={(e) => setRequesterName(e.target.value)}
                          className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none transition"
                        />
                      </div>

                      <div>
                        <label className="block text-xs font-bold text-gray-700 mb-1">ឈ្មោះអ្នកមកបើក</label>
                        <input
                          type="text"
                          placeholder="បញ្ចូលឈ្មោះអ្នកមកបើក"
                          value={collectorName}
                          onChange={(e) => setCollectorName(e.target.value)}
                          className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none transition"
                        />
                      </div>

                      <div>
                        <label className="block text-xs font-bold text-gray-700 mb-1">តួនាទីអ្នកមកបើក</label>
                        <select
                          value={collectorRoleId}
                          onChange={(e) => setCollectorRoleId(e.target.value)}
                          className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition cursor-pointer"
                        >
                          <option value="">-- ជ្រើសរើស --</option>
                          {(categories.collectorRoles || []).map((cr) => (
                            <option key={cr.id} value={cr.id}>{cr.name}</option>
                          ))}
                        </select>
                      </div>
                    </div>
                  </div>
                )}

                {operationType === 'useTeam' && (
                  <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        កាលបរិច្ឆេទ <span className="text-red-500">*</span>
                      </label>
                      <input
                        type="date"
                        required
                        value={date}
                        onChange={(e) => setDate(e.target.value)}
                        className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        ក្រុមផ្តល់ទិដ្ឋាការ.របក <span className="text-red-500">*</span>
                      </label>
                      <select
                        required
                        disabled={!isSecondary && Boolean(assignedTeam)}
                        value={visaTeamRobokId}
                        onChange={(e) => setVisaTeamRobokId(e.target.value)}
                        className={`w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none transition ${!isSecondary && assignedTeam ? 'bg-gray-100 cursor-not-allowed text-gray-700 font-semibold' : 'bg-white cursor-pointer'}`}
                      >
                        <option value="">-- ជ្រើសរើស ក្រុមផ្តល់ទិដ្ឋាការ.របក --</option>
                        {(filteredVisaTeamsRobok || []).map((vtr) => (
                          <option key={vtr.id} value={vtr.id}>{vtr.name}</option>
                        ))}
                      </select>
                    </div>
                  </div>
                )}
              </div>

              {/* Row 2: ជួរទី២ (ប្រភេទទិដ្ឋាការ=T,T1,T2,T3,E,E1,E2,E3,D,K,A,B,C, ចំនួន, ចាប់ផ្តើម, ដល់លេខ) - អាច Add Form បាន */}
              <div className="space-y-3 pt-3 border-t border-gray-200">
                <div className="flex items-center justify-between">
                  <h4 className="text-xs font-bold text-gray-800 uppercase tracking-wide flex items-center gap-2">
                    <span>ជួរទី២៖ ព័ត៌មានប្រភេទទិដ្ឋាការ, ចំនួន និងលេខស៊េរី</span>
                  </h4>
                  <button
                    type="button"
                    onClick={addStickerItem}
                    className="inline-flex items-center gap-1.5 text-xs text-blue-700 hover:text-blue-800 font-bold bg-blue-50 hover:bg-blue-100 border border-blue-200 px-3 py-1.5 rounded transition cursor-pointer shadow-2xs"
                  >
                    <Plus className="w-4 h-4" />
                    <span>+ បន្ថែមប្រភេទទិដ្ឋាការ (Add Form)</span>
                  </button>
                </div>

                <div className="space-y-3">
                    {stickerItems.map((item, index) => {
                    const isOfficeOutflowOp =
                      operationType === 'issueTeam' ||
                      operationType === 'transferTeam' ||
                      operationType === 'damaged' ||
                      operationType === 'testPrintK2';

                    // Non-hook helper function for range calculation
                    const getAvailableRangesForItem = (item: any, index: number) => {
                      let rawRanges = item.visaType
                        ? getAvailableOfficeRanges(
                            effectiveActualStockRecords,
                            item.visaType,
                            editingRecordId ? [item.id, editingRecordId] : [item.id]
                          )
                        : [];

                      // Deduct ranges consumed by previous items in the same form
                      for (let i = 0; i < index; i++) {
                        const prev = stickerItems[i];
                        if (prev.visaType === item.visaType && prev.startSerial) {
                          const pStart = parseSerialNumber(prev.startSerial);
                          const pQty = parseInt(prev.quantity, 10) || 0;
                          if (pStart) {
                            let pEndNum: bigint;
                            const pEnd = prev.endSerial ? parseSerialNumber(prev.endSerial) : null;
                            if (pEnd && pEnd.num >= pStart.num) {
                              pEndNum = pEnd.num;
                            } else if (pQty > 0) {
                              pEndNum = pStart.num + BigInt(pQty) - 1n;
                            } else {
                              pEndNum = pStart.num;
                            }

                            const curIntervals = rawRanges.map((r) => ({
                              start: r.start,
                              end: r.end,
                              prefix: r.prefix,
                              padLength: r.padLength,
                            }));
                            const remaining = subtractIntervals(curIntervals, pStart.num, pEndNum);
                            rawRanges = remaining.map((intv) => {
                              const count = Number(intv.end - intv.start + 1n);
                              return {
                                start: intv.start,
                                end: intv.end,
                                startSerial: formatSerialNumber(intv.prefix, intv.start, intv.padLength),
                                endSerial: formatSerialNumber(intv.prefix, intv.end, intv.padLength),
                                count,
                                books: Math.floor(count / 50),
                                prefix: intv.prefix,
                                padLength: intv.padLength,
                              };
                            });
                          }
                        }
                      }
                      return { availableRanges: rawRanges };
                    };

                    const { availableRanges } = getAvailableRangesForItem(item, index);

                    const qtyNum = parseInt(item.quantity, 10) || 0;
                    const sufficientRanges =
                      qtyNum > 0
                        ? availableRanges.filter((r) => r.count >= qtyNum)
                        : availableRanges;
                    const insufficientRanges =
                      qtyNum > 0 ? availableRanges.filter((r) => r.count < qtyNum) : [];

                    const isManual = manualSerialMap[item.id] || false;
                    const totalAvailableSheets = availableRanges.reduce((sum, r) => sum + r.count, 0);
                    const maxSingleRangeCount =
                      availableRanges.length > 0
                        ? Math.max(...availableRanges.map((r) => r.count))
                        : 0;

                    return (
                      <div
                        key={item.id}
                        className="flex flex-col bg-gray-50 p-3.5 rounded-md border border-gray-200 hover:border-blue-300 transition space-y-2.5"
                      >
                        <div className="flex flex-col sm:flex-row gap-3 items-start sm:items-end">
                          {/* ប្រភេទទិដ្ឋាការ (T, T1, T2, T3, E, E1, E2, E3, D, K, A, B, C) */}
                          <div className="flex-1 min-w-0 w-full sm:w-1/4">
                            <label className="block text-xs font-bold text-gray-700 mb-1">
                              ប្រភេទទិដ្ឋាការ {stickerItems.length > 1 ? `#${index + 1}` : ''} <span className="text-red-500">*</span>
                            </label>
                            <select
                              required
                              value={item.visaType}
                              onChange={(e) => {
                                const newType = e.target.value;
                                if (isOfficeOutflowOp && newType) {
                                  const ranges = getAvailableOfficeRanges(
                                    effectiveActualStockRecords,
                                    newType,
                                    editingRecordId ? [item.id, editingRecordId] : [item.id]
                                  );
                                  const q = parseInt(item.quantity, 10) || 0;
                                  const suff = q > 0 ? ranges.filter((r) => r.count >= q) : ranges;
                                  let nextStart = '';
                                  if (suff.length > 0) {
                                    nextStart = suff[0].startSerial;
                                  } else if (ranges.length > 0) {
                                    nextStart = ranges[0].startSerial;
                                  }
                                  const nextEnd = nextStart && item.quantity ? calculateEndSerial(nextStart, item.quantity) : '';
                                  updateStickerItemFields(item.id, { visaType: newType, startSerial: nextStart, endSerial: nextEnd });
                                } else {
                                  updateStickerItemFields(item.id, { visaType: newType, startSerial: '', endSerial: '' });
                                }
                              }}
                              className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition cursor-pointer"
                            >
                              <option value="">-- សូមជ្រើសរើសប្រភេទទិដ្ឋាការ --</option>
                              {VISA_TYPE_OPTIONS.map((vt) => (
                                <option key={vt} value={vt} className="text-gray-800">
                                  {vt}
                                </option>
                              ))}
                            </select>
                          </div>

                          {/* ចំនួន */}
                          <div className="flex-1 min-w-0 w-full sm:w-1/4">
                            <label className="block text-xs font-bold text-gray-700 mb-1">
                              ចំនួន (សន្លឹក) <span className="text-red-500">*</span>
                            </label>
                            <input
                              type="number"
                              min="1"
                              required
                              placeholder="បញ្ចូលចំនួនសន្លឹក"
                              value={item.quantity}
                              onChange={(e) => {
                                const val = e.target.value;
                                const q = parseInt(val, 10) || 0;
                                let nextStart = item.startSerial;

                                if (isOfficeOutflowOp && item.visaType) {
                                  const ranges = getAvailableOfficeRanges(
                                    effectiveActualStockRecords,
                                    item.visaType,
                                    editingRecordId ? [item.id, editingRecordId] : [item.id]
                                  );
                                  const suff = q > 0 ? ranges.filter((r) => r.count >= q) : ranges;
                                  const currentRange = ranges.find((r) => {
                                    const parsed = parseSerialNumber(item.startSerial);
                                    return parsed && parsed.num >= r.start && parsed.num <= r.end;
                                  });

                                  if (currentRange && q > 0) {
                                    const parsed = parseSerialNumber(item.startSerial);
                                    const availableInCurrent = parsed ? Number(currentRange.end - parsed.num + 1n) : currentRange.count;
                                    if (availableInCurrent < q) {
                                      if (suff.length > 0) {
                                        nextStart = suff[0].startSerial;
                                      } else if (ranges.length > 0) {
                                        nextStart = ranges[0].startSerial;
                                      }
                                    }
                                  } else if (!nextStart && suff.length > 0) {
                                    nextStart = suff[0].startSerial;
                                  } else if (!nextStart && ranges.length > 0) {
                                    nextStart = ranges[0].startSerial;
                                  } else if (suff.length > 0 && !suff.some((r) => r.startSerial === nextStart)) {
                                    nextStart = suff[0].startSerial;
                                  }
                                }

                                const autoEnd = nextStart && val ? calculateEndSerial(nextStart, val) : '';
                                updateStickerItemFields(item.id, { quantity: val, startSerial: nextStart, endSerial: autoEnd });
                              }}
                              className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                            />
                          </div>

                          {/* ចាប់ផ្តើម */}
                          <div className="flex-1 min-w-0 w-full sm:w-1/4">
                            <label className="block text-xs font-bold text-gray-700 mb-1">
                              ចាប់ផ្តើម {isOfficeOutflowOp && <span className="text-red-500">*</span>}
                            </label>
                            <div className="relative w-full">
                              <input
                                type="text"
                                list={`start-serials-${item.id}`}
                                required={isOfficeOutflowOp}
                                placeholder="ឧ. 1903806901"
                                value={item.startSerial}
                                onChange={(e) => {
                                  const val = e.target.value;
                                  if (val && item.quantity) {
                                    const autoEnd = calculateEndSerial(val, item.quantity);
                                    updateStickerItemFields(item.id, { startSerial: val, endSerial: autoEnd });
                                  } else {
                                    updateStickerItem(item.id, 'startSerial', val);
                                  }
                                }}
                                className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                              />
                              {isOfficeOutflowOp && (
                                <datalist id={`start-serials-${item.id}`}>
                                  {sufficientRanges.map((r) => (
                                    <option key={r.startSerial} value={r.startSerial}>
                                      {r.startSerial} (សល់ {r.count.toLocaleString()} សន្លឹក)
                                    </option>
                                  ))}
                                  {insufficientRanges.map((r) => (
                                    <option key={r.startSerial} value={r.startSerial}>
                                      {r.startSerial} (មានត្រឹម {r.count.toLocaleString()} សន្លឹក)
                                    </option>
                                  ))}
                                </datalist>
                              )}
                            </div>
                          </div>

                          {/* ដល់លេខ (Textbox) */}
                          <div className="flex-1 min-w-0 w-full sm:w-1/4">
                            <label className="block text-xs font-bold text-gray-700 mb-1">
                              ដល់លេខ {isOfficeOutflowOp && <span className="text-red-500">*</span>}
                            </label>
                            <input
                              type="text"
                              required={isOfficeOutflowOp}
                              placeholder="ស្វ័យប្រវត្តិ"
                              value={item.endSerial}
                              onChange={(e) => updateStickerItem(item.id, 'endSerial', e.target.value)}
                              className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                            />
                          </div>

                          {/* Action button */}
                          <div className="flex-none flex justify-end pb-0.5">
                            {stickerItems.length > 1 ? (
                              <button
                                type="button"
                                onClick={() => removeStickerItem(item.id)}
                                className="p-1.5 text-rose-600 hover:text-rose-800 hover:bg-rose-100 rounded border border-rose-200 transition cursor-pointer"
                                title="លុបជួរនេះ"
                              >
                                <Trash2 className="w-4 h-4" />
                              </button>
                            ) : (
                              <div className="w-8 h-8 hidden sm:block" />
                            )}
                          </div>
                        </div>

                        {/* Helper Stock Info Badge for this item */}
                        {isOfficeOutflowOp && item.visaType && (
                          <div className="flex flex-wrap items-center justify-between gap-2 pt-1 border-t border-gray-200/60 text-[11px]">
                            {item.startSerial && item.endSerial ? (
                              <div className="flex flex-wrap items-center gap-1.5 text-emerald-800 font-semibold">
                                <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600 inline" />
                                <span>លេខស៊េរីត្រូវបើក៖</span>
                                <span className="font-mono font-bold text-blue-700 bg-white px-1.5 py-0.5 rounded border border-blue-300">
                                  {item.startSerial}
                                </span>
                                <span>ដល់</span>
                                <span className="font-mono font-bold text-blue-700 bg-white px-1.5 py-0.5 rounded border border-blue-300">
                                  {item.endSerial}
                                </span>
                                <span className="text-gray-600 font-normal">({qtyNum} សន្លឹក)</span>
                              </div>
                            ) : availableRanges.length > 0 ? (
                              <span className="text-gray-600">
                                📊 ស្តុកជាក់ស្តែងសម្រាប់បើកផ្តល់តាមក្រុម ប្រភេទ <b>{item.visaType}</b>៖{' '}
                                <b className="text-blue-700">{totalAvailableSheets.toLocaleString()} សន្លឹក</b> ({availableRanges.length} ចន្លោះ)
                              </span>
                            ) : (
                              <span className="text-red-600 font-bold">
                                ⚠️ គ្មានស្តុកជាក់ស្តែងសម្រាប់បើកផ្តល់តាមក្រុម ប្រភេទ {item.visaType} នៅឡើយទេ
                              </span>
                            )}

                            {availableRanges.length > 0 && (
                              <span className="text-[10px] text-gray-500">
                                ស្តុកជាក់ស្តែងសរុប៖ <b>{totalAvailableSheets.toLocaleString()}</b> សន្លឹក ({availableRanges.length} ចន្លោះ)
                              </span>
                            )}
                          </div>
                        )}
                      </div>
                    );
                  })}
                </div>
              </div>
            </div>
          ) : (
            /* Standard E-Visa Form fields */
            <>
              {operationType === 'openK1' && (
                <div className="space-y-4">
                  <h4 className="text-xs font-bold text-gray-800 uppercase tracking-wide border-b border-gray-200 pb-2 flex items-center gap-2">
                    <span>ព័ត៌មានការបញ្ចូលស្តុក (ក១)</span>
                  </h4>

                  <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                    {/* កាលបរិច្ឆេទ */}
                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        កាលបរិច្ឆេទ <span className="text-red-500">*</span>
                      </label>
                      <CustomDatePicker
                        required
                        value={date}
                        onChange={(d) => setDate(d)}
                        className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                      />
                    </div>

                    {/* បើកពី: (ក១, ន៨) */}
                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        បើកពី <span className="text-red-500">*</span>
                      </label>
                      <select
                        required
                        value={sourceFrom}
                        onChange={(e) => setSourceFrom(e.target.value)}
                        className="w-full h-[32px] border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition cursor-pointer font-normal"
                      >
                        <option value="" className="font-normal">-- ជ្រើសរើស --</option>
                        <option value="ក១" className="font-normal">ក១</option>
                        <option value="ន៨" className="font-normal">ន៨</option>
                      </select>
                    </div>

                    {/* ចំនួន (ដុំ) */}
                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        ចំនួន (ដុំ) <span className="text-red-500">*</span>
                      </label>
                      <input
                        type="number"
                        min="1"
                        required
                        placeholder="បញ្ចូលចំនួនដុំ"
                        value={quantityBundles}
                        onChange={(e) => setQuantityBundles(e.target.value)}
                        className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none transition"
                      />
                    </div>
                  </div>
                </div>
              )}

              {operationType === 'oldStockK2' && (
                <div className="space-y-4">
                  <h4 className="text-xs font-bold text-gray-800 uppercase tracking-wide border-b border-gray-200 pb-2 flex items-center gap-2">
                    <span>ព័ត៌មានស្តុកចាស់ ក២ (សន្និធិដើមគ្រា)</span>
                  </h4>

                  <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        កាលបរិច្ឆេទ <span className="text-red-500">*</span>
                      </label>
                      <CustomDatePicker
                        required
                        value={date}
                        onChange={(d) => setDate(d)}
                        className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        ប្រភព / សម្គាល់ <span className="text-red-500">*</span>
                      </label>
                      <input
                        type="text"
                        required
                        value={sourceFrom || 'ស្តុកចាស់ ក២'}
                        onChange={(e) => setSourceFrom(e.target.value)}
                        className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        ចំនួន (ដុំ) <span className="text-red-500">*</span>
                      </label>
                      <input
                        type="number"
                        min="1"
                        required
                        placeholder="បញ្ចូលចំនួនដុំ"
                        value={quantityBundles}
                        onChange={(e) => setQuantityBundles(e.target.value)}
                        className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none transition"
                      />
                    </div>
                  </div>
                </div>
              )}

              {operationType === 'oldStockTeam' && (
                <div className="space-y-4">
                  <h4 className="text-xs font-bold text-gray-800 uppercase tracking-wide border-b border-gray-200 pb-2 flex items-center gap-2">
                    <span>ព័ត៌មានស្តុកចាស់ក្រុម</span>
                  </h4>

                  <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        កាលបរិច្ឆេទ <span className="text-red-500">*</span>
                      </label>
                      <CustomDatePicker
                        required
                        value={date}
                        onChange={(d) => setDate(d)}
                        className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        ក្រុមផ្តល់ទិដ្ឋាការ.របក <span className="text-red-500">*</span>
                      </label>
                      <select
                        required
                        disabled={!isSecondary && Boolean(assignedTeam)}
                        value={visaTeamRobokId}
                        onChange={(e) => setVisaTeamRobokId(e.target.value)}
                        className={`w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none transition ${!isSecondary && assignedTeam ? 'bg-gray-100 cursor-not-allowed text-gray-700 font-semibold' : 'bg-white cursor-pointer'}`}
                      >
                        <option value="">-- ជ្រើសរើស ក្រុមផ្តល់ទិដ្ឋាការ.របក --</option>
                        {(filteredVisaTeamsRobok || []).map((vtr) => (
                          <option key={vtr.id} value={vtr.id}>{vtr.name}</option>
                        ))}
                      </select>
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        ចំនួន (ដុំ) <span className="text-red-500">*</span>
                      </label>
                      <input
                        type="number"
                        min="1"
                        required
                        placeholder="បញ្ចូលចំនួនដុំ"
                        value={quantityBundles}
                        onChange={(e) => setQuantityBundles(e.target.value)}
                        className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none transition"
                      />
                    </div>
                  </div>
                </div>
              )}

              {operationType === 'issueTeam' && (
                <div className="space-y-4">
                  <h4 className="text-xs font-bold text-gray-800 uppercase tracking-wide border-b border-gray-200 pb-2">
                    ព័ត៌មានការបើកផ្តល់តាមក្រុម
                  </h4>

                  <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-3">
                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        ម៉ោង <span className="text-gray-400 font-normal">(លុបបើមិនត្រូវការ)</span>
                      </label>
                      <input
                        type="time"
                        value={time}
                        onChange={(e) => setTime(e.target.value)}
                        className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition font-mono"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        កាលបរិច្ឆេទ <span className="text-red-500">*</span>
                      </label>
                      <CustomDatePicker
                        required
                        value={date}
                        onChange={(d) => setDate(d)}
                        className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        ក្រុមផ្តល់ទិដ្ឋាការ.របក <span className="text-red-500">*</span>
                      </label>
                      <select
                        required
                        value={visaTeamRobokId}
                        onChange={(e) => setVisaTeamRobokId(e.target.value)}
                        className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition cursor-pointer"
                      >
                        <option value="">-- ជ្រើសរើស ក្រុមផ្តល់ទិដ្ឋាការ.របក --</option>
                        {(filteredVisaTeamsRobok || [])
                          .filter((vtr, idx) => {
                            if (vtr.id === visaTeamRobokId || vtr.name === visaTeamRobokId) return true;
                            const opt =
                              (vtr.id && teamTypeOptions[vtr.id]) ||
                              (vtr.name && teamTypeOptions[vtr.name]) ||
                              teamTypeOptions[idx];
                            if (!opt) return true;
                            return opt.evisa !== false;
                          })
                          .map((vtr) => (
                            <option key={vtr.id} value={vtr.id}>
                              {vtr.name}
                            </option>
                          ))}
                      </select>
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        ចំនួន (ដុំ) <span className="text-red-500">*</span>
                      </label>
                      <input
                        type="number"
                        min="1"
                        required
                        placeholder="បញ្ចូលចំនួនដុំ"
                        value={quantityBundles}
                        onChange={(e) => setQuantityBundles(e.target.value)}
                        className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none transition"
                      />
                    </div>
                  </div>

                  <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-3 pt-3 border-t border-gray-200">
                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">ឋានន្តរស័ក្កិស្នើសុំ</label>
                      <select
                        value={requestedRankId}
                        onChange={(e) => setRequestedRankId(e.target.value)}
                        className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition cursor-pointer"
                      >
                        <option value="">-- ជ្រើសរើស --</option>
                        {(categories.ranks || []).map((r) => (
                          <option key={r.id} value={r.id}>{r.name}</option>
                        ))}
                      </select>
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">ឈ្មោះស្នើសុំ</label>
                      <input
                        type="text"
                        placeholder="បញ្ចូលឈ្មោះអ្នកស្នើសុំ"
                        value={requesterName}
                        onChange={(e) => setRequesterName(e.target.value)}
                        className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none transition"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">ឈ្មោះអ្នកមកបើក</label>
                      <input
                        type="text"
                        placeholder="បញ្ចូលឈ្មោះអ្នកមកបើក"
                        value={collectorName}
                        onChange={(e) => setCollectorName(e.target.value)}
                        className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none transition"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">តួនាទីអ្នកមកបើក</label>
                      <select
                        value={collectorRoleId}
                        onChange={(e) => setCollectorRoleId(e.target.value)}
                        className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition cursor-pointer"
                      >
                        <option value="">-- ជ្រើសរើស --</option>
                        {(categories.collectorRoles || []).map((cr) => (
                          <option key={cr.id} value={cr.id}>{cr.name}</option>
                        ))}
                      </select>
                    </div>
                  </div>
                </div>
              )}

              {operationType === 'useTeam' && (
                <div className="space-y-4">
                  <h4 className="text-xs font-bold text-gray-800 uppercase tracking-wide border-b border-gray-200 pb-2">
                    ព័ត៌មានប្រើប្រាស់ក្រដាសតាមក្រុម
                  </h4>

                  <div className="grid grid-cols-1 md:grid-cols-3 gap-4">
                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        កាលបរិច្ឆេទ <span className="text-red-500">*</span>
                      </label>
                      <CustomDatePicker
                        required
                        value={date}
                        onChange={(d) => setDate(d)}
                        className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                      />
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        ក្រុមផ្តល់ទិដ្ឋាការ.របក <span className="text-red-500">*</span>
                      </label>
                      <select
                        required
                        disabled={!isSecondary && Boolean(assignedTeam)}
                        value={visaTeamRobokId}
                        onChange={(e) => setVisaTeamRobokId(e.target.value)}
                        className={`w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none transition ${!isSecondary && assignedTeam ? 'bg-gray-100 cursor-not-allowed text-gray-700 font-semibold' : 'bg-white cursor-pointer'}`}
                      >
                        <option value="">-- ជ្រើសរើស ក្រុមផ្តល់ទិដ្ឋាការ.របក --</option>
                        {(filteredVisaTeamsRobok || []).map((vtr) => (
                          <option key={vtr.id} value={vtr.id}>{vtr.name}</option>
                        ))}
                      </select>
                    </div>

                    <div>
                      <label className="block text-xs font-bold text-gray-700 mb-1">
                        ចំនួនប្រើប្រាស់ដុំក្រដាស <span className="text-red-500">*</span>
                      </label>
                      <input
                        type="number"
                        min="1"
                        required
                        placeholder="បញ្ចូលចំនួនដុំក្រដាសប្រើប្រាស់"
                        value={quantityBundles}
                        onChange={(e) => setQuantityBundles(e.target.value)}
                        className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none transition"
                      />
                    </div>
                  </div>
                </div>
              )}
            </>
          )}

          {/* Card Footer with Buttons */}
          <div className="pt-4 border-t border-gray-200 flex items-center justify-end gap-2 bg-gray-50 -mx-6 -mb-6 px-6 py-3.5 rounded-b-sm">
            {editingRecordId && (
              <button
                type="button"
                onClick={handleCancelEdit}
                className="bg-gray-500 hover:bg-gray-600 text-white px-4 py-2 rounded text-xs font-bold flex items-center gap-1.5 shadow-2xs transition cursor-pointer"
              >
                <X className="w-4 h-4" />
                <span>បោះបង់ (Cancel)</span>
              </button>
            )}

            <button
              type="submit"
              className={`${
                editingRecordId
                  ? 'bg-amber-600 hover:bg-amber-700'
                  : 'bg-[#28a745] hover:bg-[#218838]'
              } text-white px-6 py-2 rounded text-xs font-bold flex items-center gap-2 shadow-2xs transition cursor-pointer`}
            >
              <Check className="w-4 h-4" />
              <span>
                {editingRecordId
                  ? '✓ រក្សាទុកការកែប្រែ (Update Record)'
                  : '✓ រក្សាទុកទិន្នន័យ (Save Record)'}
              </span>
            </button>
          </div>
        </form>
      </div>
      )}

      {/* History Log Table Container */}
      {(viewMode === 'all' || viewMode === 'data') && (
        <div id="stock-manager-print-table" className="bg-white rounded-md border border-gray-300 shadow-xs overflow-hidden">
        {/* Table Header Controls */}
        <div className="p-3.5 border-b border-gray-200 bg-gray-50 flex flex-col gap-2.5 print:hidden">
          <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3">
            <div className="flex flex-wrap items-center gap-2">
              <ListFilter className="w-4 h-4 text-gray-600" />
              <h3 className="text-xs font-bold text-gray-800">
                បញ្ជីប្រតិបត្តិស្តុក ({
                  stockType === 'evisa'
                    ? 'ក្រដាសអនុម័តទិដ្ឋាការអេឡិចត្រូនិក'
                    : 'សន្លឹកទិដ្ឋាការ'
                })
              </h3>
              {!isSecondary && assignedTeam && (
                <span className="text-[11px] font-bold px-2 py-0.5 rounded-md bg-amber-100 text-amber-900 border border-amber-300">
                  ក្រុម: {assignedTeam}
                </span>
              )}
              <span className="text-[11px] font-bold px-2 py-0.5 rounded-full bg-blue-100 text-blue-800 border border-blue-200">
                {filteredRecords.length} / {currentRecords.length} ជួរ
              </span>
            </div>

            <div className="flex flex-wrap items-center gap-2">
              {/* Operation Type Filter Dropdown */}
              <div className="flex items-center gap-1">
                {stockType === 'evisa' ? (
                  <select
                    value={filterOperationType}
                    onChange={(e) => setFilterOperationType(e.target.value)}
                    className="bg-white border border-gray-300 rounded-sm px-2.5 py-1 text-xs font-bold text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none cursor-pointer"
                    title="ចម្រាញ់តាមប្រភេទប្រតិបត្តិការ"
                  >
                    <option value="all">🔍 ប្រតិបត្តិការ: ទាំងអស់ ({currentRecords.length})</option>
                    {isSecondary ? (
                      <>
                        <option value="openK1">១. ការបញ្ចូលស្តុក ({currentRecords.filter((r) => matchOperation(r, 'openK1')).length})</option>
                        <option value="issueTeam">២. ការបើកផ្តល់តាមក្រុម ({currentRecords.filter((r) => matchOperation(r, 'issueTeam')).length})</option>
                        <option value="useTeam">៣. ការប្រើប្រាស់តាមក្រុម ({currentRecords.filter((r) => matchOperation(r, 'useTeam')).length})</option>
                      </>
                    ) : (
                      <>
                        <option value="issueTeam">១. ស្តុកបានទទួលពីការិយាល័យ ({currentRecords.filter((r) => matchOperation(r, 'issueTeam')).length})</option>
                        <option value="useTeam">២. ការប្រើប្រាស់តាមក្រុម ({currentRecords.filter((r) => matchOperation(r, 'useTeam')).length})</option>
                      </>
                    )}
                  </select>
                ) : (
                  <select
                    value={filterOperationType}
                    onChange={(e) => setFilterOperationType(e.target.value)}
                    className="bg-white border border-gray-300 rounded-sm px-2.5 py-1 text-xs font-bold text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none cursor-pointer"
                    title="ចម្រាញ់តាមប្រភេទប្រតិបត្តិការ"
                  >
                    <option value="all">🔍 ប្រតិបត្តិការ: ទាំងអស់ ({currentRecords.length})</option>
                    {isSecondary ? (
                      <>
                        <option value="openK1">១. ការបញ្ចូលស្តុក (ក១) ({currentRecords.filter((r) => matchOperation(r, 'openK1')).length})</option>
                        <option value="oldStockK2">២. ស្តុកចាស់ ក២ ({currentRecords.filter((r) => matchOperation(r, 'oldStockK2')).length})</option>
                        <option value="oldStockTeam">៣. ស្តុកចាស់ក្រុម ({currentRecords.filter((r) => matchOperation(r, 'oldStockTeam')).length})</option>
                        <option value="issueTeam">៤. ការបើកផ្តល់តាមក្រុម ({currentRecords.filter((r) => matchOperation(r, 'issueTeam')).length})</option>
                        <option value="useTeam">៥. ការប្រើប្រាស់តាមក្រុម ({currentRecords.filter((r) => matchOperation(r, 'useTeam')).length})</option>
                        <option value="testPrintK2">៦. ទិដ្ឋាការសាកក២ ({currentRecords.filter((r) => matchOperation(r, 'testPrintK2')).length})</option>
                        <option value="damaged">៧. ទិដ្ឋាការខូចក២ ({currentRecords.filter((r) => matchOperation(r, 'damaged')).length})</option>
                        <option value="damagedTeam">៨. ទិដ្ឋាការខូចក្រុម ({currentRecords.filter((r) => matchOperation(r, 'damagedTeam')).length})</option>
                        <option value="missingTeam">៩. ទិដ្ឋាការខ្វះក្រុម ({currentRecords.filter((r) => matchOperation(r, 'missingTeam')).length})</option>
                        <option value="returnTeam">១០. ទិដ្ឋាការបង្វិលពីក្រុម ({currentRecords.filter((r) => matchOperation(r, 'returnTeam')).length})</option>
                        <option value="transferTeam">១១. ផ្ទេរការប្រើប្រាស់ក្រុម ({currentRecords.filter((r) => matchOperation(r, 'transferTeam')).length})</option>
                        <option value="returnK1">១២. បង្វិលក១ ({currentRecords.filter((r) => matchOperation(r, 'returnK1')).length})</option>
                        <option value="returnStub">១៣. ប្រមូលគល់សន្លឹកទិដ្ឋាការ ({currentRecords.filter((r) => matchOperation(r, 'returnStub')).length})</option>
                      </>
                    ) : (
                      <>
                        <option value="issueTeam">១. ស្តុកបានទទួលពីការិយាល័យ ({currentRecords.filter((r) => matchOperation(r, 'issueTeam')).length})</option>
                        <option value="oldStockTeam">២. ស្តុកចាស់ក្រុម ({currentRecords.filter((r) => matchOperation(r, 'oldStockTeam')).length})</option>
                        <option value="returnTeam">៣. ទិដ្ឋាការបង្វិលទៅក២ ({currentRecords.filter((r) => matchOperation(r, 'returnTeam')).length})</option>
                        <option value="transferTeam">៤. ផ្ទេរការប្រើប្រាស់ក្រុម ({currentRecords.filter((r) => matchOperation(r, 'transferTeam')).length})</option>
                        <option value="damagedTeam">៥. ទិដ្ឋាការខូចក្រុម ({currentRecords.filter((r) => matchOperation(r, 'damagedTeam')).length})</option>
                        <option value="missingTeam">៦. ទិដ្ឋាការខ្វះក្រុម ({currentRecords.filter((r) => matchOperation(r, 'missingTeam')).length})</option>
                        {currentRecords.some((r) => matchOperation(r, 'useTeam')) && (
                          <option value="useTeam">៧. ការប្រើប្រាស់តាមក្រុម ({currentRecords.filter((r) => matchOperation(r, 'useTeam')).length})</option>
                        )}
                      </>
                    )}
                  </select>
                )}
              </div>

              {/* Search input */}
              <div className="relative">
                <Search className="w-3.5 h-3.5 text-gray-400 absolute left-2.5 top-1/2 -translate-y-1/2" />
                <input
                  type="text"
                  placeholder="ស្វែងរកទិន្នន័យ..."
                  value={searchTerm}
                  onChange={(e) => setSearchTerm(e.target.value)}
                  className="bg-white border border-gray-300 rounded-sm pl-8 pr-3 py-1 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none w-44"
                />
              </div>

              {/* Find records not counted in Team Stock Report (only for sticker and Secondary role) */}
              {isSecondary && stockType === 'sticker' && (
                <button
                  type="button"
                  onClick={() => {
                    setFilterUncountedInTeamReport(!filterUncountedInTeamReport);
                    setUncountedCategoryFilter('all');
                    setTablePage(1);
                  }}
                  className={`px-3 py-1 rounded text-xs font-bold transition flex items-center gap-1.5 cursor-pointer shadow-2xs ${
                    filterUncountedInTeamReport
                      ? 'bg-amber-600 hover:bg-amber-700 text-white ring-2 ring-amber-300'
                      : 'bg-amber-50 hover:bg-amber-100 text-amber-900 border border-amber-300'
                  }`}
                  title="ស្វែងរក និងបង្ហាញទិន្នន័យដែលមិនត្រូវបានរាប់បញ្ចូល/បូក ក្នុងរបាយការណ៍តាមក្រុម (សន្លឹកទិដ្ឋាការស្អិត)"
                >
                  <Search className="w-3.5 h-3.5" />
                  <span>ទិន្នន័យមិនរាប់ក្នុងរបាយការណ៍ក្រុម</span>
                  <span
                    className={`px-1.5 py-0.2 rounded-full text-[10px] font-extrabold ${
                      filterUncountedInTeamReport ? 'bg-white text-amber-800' : 'bg-amber-200 text-amber-900'
                    }`}
                  >
                    {uncountedInTeamReportRecords.length}
                  </span>
                </button>
              )}

              {/* Export Excel button */}
              <button
                type="button"
                onClick={handleExportExcel}
                className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1 rounded text-xs font-bold transition flex items-center gap-1.5 cursor-pointer shadow-2xs"
                title="ទាញយកទិន្នន័យដែលបានចម្រាញ់ជាឯកសារ Excel"
              >
                <FileSpreadsheet className="w-3.5 h-3.5" />
                <span>ទាញយក Excel {filterOperationType !== 'all' ? '(ចម្រាញ់)' : ''}</span>
              </button>

              {/* Import Excel button */}
              <button
                type="button"
                onClick={() => setIsImportModalOpen(true)}
                className="bg-teal-600 hover:bg-teal-700 text-white px-3 py-1 rounded text-xs font-bold transition flex items-center gap-1.5 cursor-pointer shadow-2xs"
                title="នាំចូលទិន្នន័យពីឯកសារ Excel តាមប្រភេទប្រតិបត្តិការ"
              >
                <FileSpreadsheet className="w-3.5 h-3.5" />
                <span>នាំចូល Excel</span>
              </button>

              {/* Delete All button */}
              <button
                type="button"
                onClick={() => {
                  setDeleteScope(filterOperationType !== 'all' ? 'filtered' : 'all');
                  setIsDeleteAllModalOpen(true);
                }}
                className="bg-rose-600 hover:bg-rose-700 text-white px-3 py-1 rounded text-xs font-bold transition flex items-center gap-1.5 cursor-pointer shadow-2xs"
                title={
                  filterOperationType !== 'all'
                    ? `លុបទិន្នន័យ «${opLabelMap[filterOperationType] || filterOperationType}» (${currentRecords.filter((r) => matchOperation(r, filterOperationType)).length} ជួរ)`
                    : 'លុបទិន្នន័យទាំងអស់ក្នុងផ្នែកនេះ'
                }
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>
                  {filterOperationType !== 'all'
                    ? `លុប «${opShortLabelMap[filterOperationType] || filterOperationType}» (${currentRecords.filter((r) => matchOperation(r, filterOperationType)).length})`
                    : 'លុបទិន្នន័យទាំងអស់'}
                </span>
              </button>
            </div>
          </div>

          {/* Date Range & Quick Presets Filter Row */}
          <div className="flex flex-wrap items-center justify-between gap-2.5 pt-2 border-t border-gray-200">
            {/* Start Date & End Date Filter Inputs */}
            <div className="flex flex-wrap items-center gap-2 text-xs">
              <span className="flex items-center gap-1 text-[11px] font-bold text-gray-700">
                <Calendar className="w-3.5 h-3.5 text-blue-600" />
                <span>កាលបរិច្ឆេទ:</span>
              </span>
              <div className="flex items-center gap-1">
                <label className="text-[11px] text-gray-600 font-medium">ចាប់ពី:</label>
                <div className="w-28">
                  <CustomDatePicker
                    value={filterStartDate}
                    onChange={(d) => {
                      setFilterStartDate(d);
                      setTablePage(1);
                    }}
                    placeholder="YYYY-MM-DD"
                    className="py-0.5 text-xs"
                  />
                </div>
              </div>
              <div className="flex items-center gap-1">
                <label className="text-[11px] text-gray-600 font-medium">ដល់:</label>
                <div className="w-28">
                  <CustomDatePicker
                    value={filterEndDate}
                    onChange={(d) => {
                      setFilterEndDate(d);
                      setTablePage(1);
                    }}
                    placeholder="YYYY-MM-DD"
                    className="py-0.5 text-xs"
                  />
                </div>
              </div>

              {/* Quick Date Presets */}
              <div className="flex items-center gap-1 pl-1">
                <button
                  type="button"
                  onClick={() => {
                    const todayStr = new Date().toISOString().split('T')[0];
                    setFilterStartDate(todayStr);
                    setFilterEndDate(todayStr);
                    setTablePage(1);
                  }}
                  className="px-2 py-0.5 text-[11px] font-medium rounded bg-white hover:bg-gray-100 border border-gray-300 text-gray-700 transition cursor-pointer"
                  title="ចម្រាញ់យកទិន្នន័យថ្ងៃនេះ"
                >
                  ថ្ងៃនេះ
                </button>
                <button
                  type="button"
                  onClick={() => {
                    const now = new Date();
                    const year = now.getFullYear();
                    const month = String(now.getMonth() + 1).padStart(2, '0');
                    const lastDay = new Date(year, now.getMonth() + 1, 0).getDate();
                    setFilterStartDate(`${year}-${month}-01`);
                    setFilterEndDate(`${year}-${month}-${String(lastDay).padStart(2, '0')}`);
                    setTablePage(1);
                  }}
                  className="px-2 py-0.5 text-[11px] font-medium rounded bg-white hover:bg-gray-100 border border-gray-300 text-gray-700 transition cursor-pointer"
                  title="ចម្រាញ់យកទិន្នន័យខែនេះ"
                >
                  ខែនេះ
                </button>
                {(filterStartDate || filterEndDate) && (
                  <button
                    type="button"
                    onClick={() => {
                      setFilterStartDate('');
                      setFilterEndDate('');
                      setTablePage(1);
                    }}
                    className="flex items-center gap-1 px-2 py-0.5 text-[11px] font-bold rounded bg-rose-50 hover:bg-rose-100 border border-rose-200 text-rose-700 transition cursor-pointer"
                    title="សម្អាតការចម្រាញ់តាមកាលបរិច្ឆេទ"
                  >
                    <X className="w-3 h-3" />
                    <span>សម្អាតថ្ងៃ</span>
                  </button>
                )}
              </div>
            </div>

            {/* Active Date Tag */}
            {(filterStartDate || filterEndDate) && (
              <div className="flex items-center gap-1.5 text-[11px] bg-blue-50 border border-blue-200 text-blue-800 px-2.5 py-0.5 rounded shadow-2xs">
                <span>ចន្លោះថ្ងៃចម្រាញ់:</span>
                <span className="font-bold font-mono text-blue-900">
                  {filterStartDate || '...'} ដល់ {filterEndDate || '...'}
                </span>
              </div>
            )}
          </div>

          {/* Quick Filter Operation Pills */}
          <div className="flex flex-wrap items-center gap-1.5 pt-1 border-t border-gray-200">
            <span className="text-[11px] font-semibold text-gray-500 mr-1">ប្រភេទប្រតិបត្តិការ:</span>
            <button
              type="button"
              onClick={() => setFilterOperationType('all')}
              className={`px-2 py-0.5 rounded text-[11px] font-bold transition cursor-pointer ${
                filterOperationType === 'all'
                  ? 'bg-blue-600 text-white shadow-2xs'
                  : 'bg-white text-gray-700 border border-gray-300 hover:bg-gray-100'
              }`}
            >
              ទាំងអស់ ({currentRecords.length})
            </button>
            {isSecondary && (
              <button
                type="button"
                onClick={() => setFilterOperationType('openK1')}
                className={`px-2 py-0.5 rounded text-[11px] font-bold transition cursor-pointer ${
                  filterOperationType === 'openK1'
                    ? 'bg-blue-600 text-white shadow-2xs'
                    : 'bg-white text-gray-700 border border-gray-300 hover:bg-gray-100'
                }`}
              >
                បញ្ចូលស្តុក ({currentRecords.filter((r) => matchOperation(r, 'openK1')).length})
              </button>
            )}
            {stockType === 'sticker' && isSecondary && (
              <button
                type="button"
                onClick={() => setFilterOperationType('oldStockK2')}
                className={`px-2 py-0.5 rounded text-[11px] font-bold transition cursor-pointer ${
                  filterOperationType === 'oldStockK2'
                    ? 'bg-emerald-600 text-white shadow-2xs'
                    : 'bg-white text-emerald-800 border border-emerald-300 hover:bg-emerald-50'
                }`}
              >
                ស្តុកចាស់ ក២ ({currentRecords.filter((r) => matchOperation(r, 'oldStockK2')).length})
              </button>
            )}
            <button
              type="button"
              onClick={() => setFilterOperationType('issueTeam')}
              className={`px-2 py-0.5 rounded text-[11px] font-bold transition cursor-pointer ${
                filterOperationType === 'issueTeam'
                  ? 'bg-amber-600 text-white shadow-2xs'
                  : 'bg-white text-gray-700 border border-gray-300 hover:bg-gray-100'
              }`}
            >
              {isSecondary ? 'បើកផ្តល់តាមក្រុម' : 'ទទួលពីការិយាល័យ'} ({currentRecords.filter((r) => matchOperation(r, 'issueTeam')).length})
            </button>
            <button
              type="button"
              onClick={() => setFilterOperationType('useTeam')}
              className={`px-2 py-0.5 rounded text-[11px] font-bold transition cursor-pointer ${
                filterOperationType === 'useTeam'
                  ? 'bg-purple-600 text-white shadow-2xs'
                  : 'bg-white text-gray-700 border border-gray-300 hover:bg-gray-100'
              }`}
            >
              ប្រើប្រាស់តាមក្រុម ({currentRecords.filter((r) => matchOperation(r, 'useTeam')).length})
            </button>
            {stockType === 'sticker' && (
              <>
                <button
                  type="button"
                  onClick={() => setFilterOperationType('oldStockTeam')}
                  className={`px-2 py-0.5 rounded text-[11px] font-bold transition cursor-pointer ${
                    filterOperationType === 'oldStockTeam'
                      ? 'bg-teal-600 text-white shadow-2xs'
                      : 'bg-white text-teal-800 border border-teal-300 hover:bg-teal-50'
                  }`}
                >
                  ស្តុកចាស់ក្រុម ({currentRecords.filter((r) => matchOperation(r, 'oldStockTeam')).length})
                </button>
                {isSecondary && (
                  <>
                    <button
                      type="button"
                      onClick={() => setFilterOperationType('testPrintK2')}
                      className={`px-2 py-0.5 rounded text-[11px] font-bold transition cursor-pointer ${
                        filterOperationType === 'testPrintK2'
                          ? 'bg-amber-700 text-white shadow-2xs'
                          : 'bg-white text-gray-700 border border-gray-300 hover:bg-gray-100'
                      }`}
                    >
                      ទិដ្ឋាការសាកក២ ({currentRecords.filter((r) => matchOperation(r, 'testPrintK2')).length})
                    </button>
                    <button
                      type="button"
                      onClick={() => setFilterOperationType('damaged')}
                      className={`px-2 py-0.5 rounded text-[11px] font-bold transition cursor-pointer ${
                        filterOperationType === 'damaged'
                          ? 'bg-rose-700 text-white shadow-2xs'
                          : 'bg-white text-gray-700 border border-gray-300 hover:bg-gray-100'
                      }`}
                    >
                      ទិដ្ឋាការខូចក២ ({currentRecords.filter((r) => matchOperation(r, 'damaged')).length})
                    </button>
                  </>
                )}
                <button
                  type="button"
                  onClick={() => setFilterOperationType('damagedTeam')}
                  className={`px-2 py-0.5 rounded text-[11px] font-bold transition cursor-pointer ${
                    filterOperationType === 'damagedTeam'
                      ? 'bg-orange-700 text-white shadow-2xs'
                      : 'bg-white text-gray-700 border border-gray-300 hover:bg-gray-100'
                  }`}
                >
                  ទិដ្ឋាការខូចក្រុម ({currentRecords.filter((r) => matchOperation(r, 'damagedTeam')).length})
                </button>
                <button
                  type="button"
                  onClick={() => setFilterOperationType('missingTeam')}
                  className={`px-2 py-0.5 rounded text-[11px] font-bold transition cursor-pointer ${
                    filterOperationType === 'missingTeam'
                      ? 'bg-red-700 text-white shadow-2xs'
                      : 'bg-white text-gray-700 border border-gray-300 hover:bg-gray-100'
                  }`}
                >
                  ទិដ្ឋាការខ្វះក្រុម ({currentRecords.filter((r) => matchOperation(r, 'missingTeam')).length})
                </button>
                <button
                  type="button"
                  onClick={() => setFilterOperationType('returnTeam')}
                  className={`px-2 py-0.5 rounded text-[11px] font-bold transition cursor-pointer ${
                    filterOperationType === 'returnTeam'
                      ? 'bg-cyan-700 text-white shadow-2xs'
                      : 'bg-white text-gray-700 border border-gray-300 hover:bg-gray-100'
                  }`}
                >
                  {!isSecondary ? 'ទិដ្ឋាការបង្វិលទៅក២' : 'ទិដ្ឋាការបង្វិលពីក្រុម'} ({currentRecords.filter((r) => matchOperation(r, 'returnTeam')).length})
                </button>
                <button
                  type="button"
                  onClick={() => setFilterOperationType('transferTeam')}
                  className={`px-2 py-0.5 rounded text-[11px] font-bold transition cursor-pointer ${
                    filterOperationType === 'transferTeam'
                      ? 'bg-indigo-700 text-white shadow-2xs'
                      : 'bg-white text-gray-700 border border-gray-300 hover:bg-gray-100'
                  }`}
                >
                  ផ្ទេរការប្រើប្រាស់ ({currentRecords.filter((r) => matchOperation(r, 'transferTeam')).length})
                </button>
                {isSecondary && (
                  <button
                    type="button"
                    onClick={() => setFilterOperationType(filterOperationType === 'returnStub' ? 'all' : 'returnStub')}
                    className={`px-2 py-0.5 rounded text-[11px] font-bold transition cursor-pointer ${
                      filterOperationType === 'returnStub'
                        ? 'bg-teal-700 text-white shadow-2xs'
                        : 'bg-white text-teal-800 border border-teal-300 hover:bg-teal-50'
                    }`}
                  >
                    ប្រមូលគល់សន្លឹក ({currentRecords.filter((r) => matchOperation(r, 'returnStub')).length})
                  </button>
                )}

                {/* Quick Uncounted In Team Report Pill */}
                {isSecondary && (
                  <button
                    type="button"
                    onClick={() => {
                      setFilterUncountedInTeamReport(!filterUncountedInTeamReport);
                      setUncountedCategoryFilter('all');
                      setTablePage(1);
                    }}
                    className={`px-2 py-0.5 rounded text-[11px] font-bold transition cursor-pointer flex items-center gap-1 ${
                      filterUncountedInTeamReport
                        ? 'bg-amber-600 text-white shadow-2xs ring-2 ring-amber-300'
                        : 'bg-amber-50 text-amber-900 border border-amber-300 hover:bg-amber-100'
                    }`}
                    title="ស្វែងរក និងបង្ហាញទិន្នន័យដែលមិនត្រូវបានរាប់បញ្ចូល/បូក ក្នុងរបាយការណ៍តាមក្រុម"
                  >
                    <span>⚠️ មិនបូកក្នុងរបាយការណ៍ក្រុម ({uncountedInTeamReportRecords.length})</span>
                  </button>
                )}
              </>
            )}

            {/* Quick Action: Delete only the active operation type */}
            {filterOperationType !== 'all' && (
              <button
                type="button"
                onClick={() => {
                  setDeleteScope('filtered');
                  setIsDeleteAllModalOpen(true);
                }}
                className="ml-auto bg-rose-50 hover:bg-rose-100 text-rose-700 border border-rose-300 px-2 py-0.5 rounded text-[11px] font-bold transition flex items-center gap-1 cursor-pointer shadow-2xs"
                title={`លុបទិន្នន័យ «${opLabelMap[filterOperationType] || filterOperationType}» ចំនួន ${currentRecords.filter((r) => matchOperation(r, filterOperationType)).length} ជួរ`}
              >
                <Trash2 className="w-3 h-3 text-rose-600" />
                <span>លុបតែ «{opShortLabelMap[filterOperationType] || filterOperationType}» ({currentRecords.filter((r) => matchOperation(r, filterOperationType)).length})</span>
              </button>
            )}
          </div>
        </div>

        {/* Uncounted In Team Report Audit Banner */}
        {filterUncountedInTeamReport && (
          <div className="mx-3.5 my-3 p-3 bg-amber-50 border border-amber-300 rounded-md shadow-xs print:hidden">
            <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 border-b border-amber-200 pb-2">
              <div className="flex items-center gap-2">
                <span className="p-1 bg-amber-200 text-amber-900 rounded font-bold text-sm">⚠️</span>
                <div>
                  <h4 className="text-xs font-bold text-amber-900">
                    កំពុងបង្ហាញទិន្នន័យដែល «មិនត្រូវបានរាប់បញ្ចូល / មិនបូក» ក្នុងរបាយការណ៍តាមក្រុម ({filteredRecords.length} ជួរ)
                  </h4>
                  <p className="text-[11px] text-amber-800">
                    ទិន្នន័យខាងក្រោមនេះមិនត្រូវបានរាប់បញ្ចូលក្នុងតារាងរបាយការណ៍ទាំង ៦ ទំព័រឡើយ ដោយសារលក្ខខណ្ឌកំណត់ ប្រតិបត្តិការផ្ទៃក្នុងក២ ឬភាពមិនស៊ីគ្នានៃទិន្នន័យ។
                  </p>
                </div>
              </div>
              <button
                type="button"
                onClick={() => {
                  setFilterUncountedInTeamReport(false);
                  setUncountedCategoryFilter('all');
                  setTablePage(1);
                }}
                className="self-start sm:self-auto px-2.5 py-1 bg-white hover:bg-gray-100 border border-amber-300 text-gray-700 text-xs font-bold rounded shadow-2xs cursor-pointer flex items-center gap-1"
              >
                <X className="w-3.5 h-3.5" />
                <span>បិទការចម្រាញ់ (បង្ហាញធម្មតា)</span>
              </button>
            </div>

            {/* Category Filter Chips */}
            <div className="flex flex-wrap items-center gap-1.5 pt-2">
              <span className="text-[11px] font-bold text-amber-900 mr-1">ចម្រាញ់តាមមូលហេតុ:</span>
              <button
                type="button"
                onClick={() => {
                  setUncountedCategoryFilter('all');
                  setTablePage(1);
                }}
                className={`px-2 py-0.5 rounded text-[11px] font-bold transition cursor-pointer ${
                  uncountedCategoryFilter === 'all'
                    ? 'bg-amber-700 text-white shadow-2xs'
                    : 'bg-white text-amber-900 border border-amber-300 hover:bg-amber-100'
                }`}
              >
                ទាំងអស់ ({uncountedCategoryCounts.all})
              </button>
              {uncountedCategoryCounts.unmatched_team > 0 && (
                <button
                  type="button"
                  onClick={() => {
                    setUncountedCategoryFilter('unmatched_team');
                    setTablePage(1);
                  }}
                  className={`px-2 py-0.5 rounded text-[11px] font-bold transition cursor-pointer ${
                    uncountedCategoryFilter === 'unmatched_team'
                      ? 'bg-red-700 text-white shadow-2xs'
                      : 'bg-white text-red-800 border border-red-300 hover:bg-red-50'
                  }`}
                >
                  ⚠️ មិនស្គាល់ឈ្មោះក្រុម ({uncountedCategoryCounts.unmatched_team})
                </button>
              )}
              {uncountedCategoryCounts.unmatched_visa_type > 0 && (
                <button
                  type="button"
                  onClick={() => {
                    setUncountedCategoryFilter('unmatched_visa_type');
                    setTablePage(1);
                  }}
                  className={`px-2 py-0.5 rounded text-[11px] font-bold transition cursor-pointer ${
                    uncountedCategoryFilter === 'unmatched_visa_type'
                      ? 'bg-rose-700 text-white shadow-2xs'
                      : 'bg-white text-rose-800 border border-rose-300 hover:bg-rose-50'
                  }`}
                >
                  ⚠️ មិនស្គាល់ប្រភេទទិដ្ឋាការ ({uncountedCategoryCounts.unmatched_visa_type})
                </button>
              )}
              {uncountedCategoryCounts.k2_internal_op > 0 && (
                <button
                  type="button"
                  onClick={() => {
                    setUncountedCategoryFilter('k2_internal_op');
                    setTablePage(1);
                  }}
                  className={`px-2 py-0.5 rounded text-[11px] font-bold transition cursor-pointer ${
                    uncountedCategoryFilter === 'k2_internal_op'
                      ? 'bg-blue-700 text-white shadow-2xs'
                      : 'bg-white text-blue-800 border border-blue-300 hover:bg-blue-50'
                  }`}
                >
                  ℹ️ ប្រតិបត្តិការផ្ទៃក្នុងក២/ក១ ({uncountedCategoryCounts.k2_internal_op})
                </button>
              )}
              {uncountedCategoryCounts.pre_nov_2018 > 0 && (
                <button
                  type="button"
                  onClick={() => {
                    setUncountedCategoryFilter('pre_nov_2018');
                    setTablePage(1);
                  }}
                  className={`px-2 py-0.5 rounded text-[11px] font-bold transition cursor-pointer ${
                    uncountedCategoryFilter === 'pre_nov_2018'
                      ? 'bg-purple-700 text-white shadow-2xs'
                      : 'bg-white text-purple-800 border border-purple-300 hover:bg-purple-50'
                  }`}
                >
                  ℹ️ ក្រោមថ្ងៃ ៣០-១១-២០១៨ ({uncountedCategoryCounts.pre_nov_2018})
                </button>
              )}
              {uncountedCategoryCounts.zero_qty > 0 && (
                <button
                  type="button"
                  onClick={() => {
                    setUncountedCategoryFilter('zero_qty');
                    setTablePage(1);
                  }}
                  className={`px-2 py-0.5 rounded text-[11px] font-bold transition cursor-pointer ${
                    uncountedCategoryFilter === 'zero_qty'
                      ? 'bg-gray-700 text-white shadow-2xs'
                      : 'bg-white text-gray-800 border border-gray-300 hover:bg-gray-100'
                  }`}
                >
                  ⚠️ ចំនួនស្មើ ០ ({uncountedCategoryCounts.zero_qty})
                </button>
              )}
            </div>
          </div>
        )}

        {/* Printable Title Header */}
        <div className="hidden print:block text-center p-4 border-b">
          <h2 className="text-lg font-bold">
            របាយការណ៍ស្តុក —{' '}
            {stockType === 'evisa'
              ? 'ក្រដាសអនុម័តទិដ្ឋាការអេឡិចត្រូនិក'
              : 'សន្លឹកទិដ្ឋាការ'}
          </h2>
          <p className="text-xs text-gray-500">កាលបរិច្ឆេទបោះពុម្ព: {new Date().toLocaleDateString('km-KH')}</p>
        </div>

        {/* Table Content */}
        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse text-xs align-middle">
            <thead>
              <tr className="bg-gray-100 text-gray-700 font-bold border-b border-gray-200 align-middle">
                <th className="py-2.5 px-3 border-r border-gray-200 text-center w-12 align-middle">ល.រ</th>
                <th className="py-2.5 px-3 border-r border-gray-200 align-middle">ប្រតិបត្តិការ</th>
                <th className="py-2.5 px-3 border-r border-gray-200 align-middle">ក្រុមផ្តល់ទិដ្ឋាការ</th>
                {stockType === 'sticker' && (
                  <th className="py-2.5 px-3 border-r border-gray-200 text-center align-middle">
                    <div className="flex flex-col items-center justify-center">
                      <span>ប្រភេទទិដ្ឋាការ</span>
                    </div>
                  </th>
                )}
                <th className="py-2.5 px-3 border-r border-gray-200 text-center align-middle">
                  {stockType === 'evisa' ? 'ចំនួន(ដុំ)' : 'ចំនួន(សន្លឹក)'}
                </th>
                <th className="py-2.5 px-3 border-r border-gray-200 text-center align-middle">ចាប់ពី</th>
                <th className="py-2.5 px-3 border-r border-gray-200 text-center align-middle">ដល់លេខ</th>
                <th className="py-2.5 px-3 border-r border-gray-200 align-middle">អ្នកស្នើសុំ (មានឋានន្តរស័ក្តិ/ឈ្មោះ)</th>
                <th className="py-2.5 px-3 border-r border-gray-200 align-middle">អ្នកមកបើក (តួនាទី/ឈ្មោះ)</th>
                <th className="py-2.5 px-3 text-center print:hidden w-24 align-middle">សកម្មភាព</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-200 align-middle">
              {filteredRecords.length === 0 ? (
                <tr>
                  <td colSpan={stockType === 'sticker' ? 10 : 9} className="py-8 text-center text-gray-500 align-middle">
                    មិនមានទិន្នន័យស្តុកនៅឡើយទេ!
                  </td>
                </tr>
              ) : (
                paginatedRecords.map((item, idx) => {
                  const isHighlighted = highlightedRecordIds.includes(item.id);
                  const rowNumber = (safeCurrentTablePage - 1) * (tablePageSize >= 999999 ? 0 : tablePageSize) + idx + 1;
                  const itemInclusion = inclusionStatusMap.get(item.id);
                  const isUncounted = itemInclusion && !itemInclusion.isIncluded;

                  return (
                    <tr
                      id={`stock-row-${item.id}`}
                      key={item.id}
                      className={`align-middle transition-all duration-300 ${
                        isHighlighted
                          ? 'bg-amber-100/90 ring-2 ring-amber-500 font-semibold shadow-xs'
                          : editingRecordId === item.id
                          ? 'bg-amber-50'
                          : isUncounted && filterUncountedInTeamReport
                          ? 'bg-amber-50/50 hover:bg-amber-100/50'
                          : 'hover:bg-blue-50/40'
                      }`}
                    >
                      <td className="py-2.5 px-3 border-r border-gray-200 text-center font-medium align-middle">
                        <div className="flex flex-col items-center justify-center gap-1">
                          <span>{rowNumber}</span>
                          {isHighlighted && (
                            <span className="inline-flex items-center gap-0.5 px-1.5 py-0.2 rounded text-[9px] font-bold bg-amber-500 text-white shadow-2xs">
                              <Check className="w-2.5 h-2.5" />
                              <span>ទើបកែ</span>
                            </span>
                          )}
                          {isUncounted && (
                            <span
                              className="inline-flex items-center gap-0.5 px-1 py-0.2 rounded text-[8.5px] font-bold bg-amber-100 text-amber-900 border border-amber-300"
                              title={itemInclusion.reason}
                            >
                              មិនបូក
                            </span>
                          )}
                        </div>
                      </td>
                    <td className="py-2.5 px-3 border-r border-gray-200 align-middle">
                      <span
                        className={`inline-block px-2 py-0.5 rounded text-[11px] font-bold ${
                          item.operationType === 'oldStockK2'
                            ? 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                            : item.operationType === 'oldStockTeam'
                            ? 'bg-teal-100 text-teal-800 border border-teal-300'
                            : item.operationType === 'testPrintK2'
                            ? 'bg-amber-100 text-amber-900 border border-amber-300'
                            : item.operationType === 'damaged'
                            ? 'bg-rose-100 text-rose-800 border border-rose-300'
                            : item.operationType === 'damagedTeam'
                            ? 'bg-orange-100 text-orange-800 border border-orange-300'
                            : item.operationType === 'missingTeam'
                            ? 'bg-red-100 text-red-800 border border-red-300'
                            : item.operationType === 'returnStub' || item.sourceFrom?.includes('គល់សន្លឹក') || item.remarks?.includes('គល់សន្លឹក') || item.id?.startsWith('stock-stub-')
                            ? 'bg-teal-100 text-teal-800 border border-teal-300'
                            : item.operationType === 'returnTeam'
                            ? 'bg-cyan-100 text-cyan-800 border border-cyan-300'
                            : item.operationType === 'transferTeam'
                            ? 'bg-indigo-100 text-indigo-800 border border-indigo-300'
                            : item.operationType === 'issueTeam'
                            ? 'bg-blue-100 text-blue-900 border border-blue-300'
                            : item.operationType === 'openK1'
                            ? item.sourceFrom?.includes('ស្តុកចាស់')
                              ? 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                              : 'bg-blue-100 text-blue-800'
                            : item.operationType === 'useTeam'
                            ? 'bg-purple-100 text-purple-800'
                            : item.sourceFrom?.includes('ស្តុកចាស់')
                            ? 'bg-teal-100 text-teal-800 border border-teal-300'
                            : 'bg-amber-100 text-amber-800'
                        }`}
                      >
                        {getRecordOperationLabel(item)}
                      </span>
                      {item.sourceFrom && item.operationType !== 'useTeam' && item.operationType !== 'issueTeam' && (
                        <span className="block text-[10px] text-gray-500 font-normal mt-0.5">
                          {item.operationType === 'openK1' ? 'បើកពី: ' : 'សម្គាល់: '}
                          <strong className="text-gray-700">{item.sourceFrom}</strong>
                        </span>
                      )}
                      <span className="block text-[10px] text-gray-400 font-normal mt-0.5">
                        {item.date} {item.time ? `(${item.time})` : ''}
                      </span>
                      {stockType === 'sticker' && item.date && (sameDateCountsMap[item.date] || 0) > 1 && (
                        <span className="inline-block mt-0.5 px-1.5 py-0.5 rounded text-[10px] font-semibold bg-blue-50 text-blue-700 border border-blue-200" title="មានប្រភេទផ្សេងទៀតកត់ត្រាក្នុងថ្ងៃតែមួយ">
                          ថ្ងៃនេះមាន {sameDateCountsMap[item.date]} ប្រភេទ
                        </span>
                      )}
                      {isUncounted && (
                        <div className="mt-1 p-1 bg-amber-50 border border-amber-300 rounded text-[9.5px] text-amber-900 leading-tight">
                          <strong className="text-amber-800">⚠️ មិនបូកក្នុងរបាយការណ៍: </strong>
                          <span>{itemInclusion.reason}</span>
                        </div>
                      )}
                    </td>
                    <td className="py-2.5 px-3 border-r border-gray-200 font-medium align-middle">
                      <div>
                        {(() => {
                          const currentTeam = targetTeamFilter || assignedTeam || '';
                          const isRecipient = item.operationType === 'transferTeam' && currentTeam && isRecordForRecipientTeam(item, currentTeam);
                          
                          if (isRecipient) {
                            return (
                              <div>
                                <span className="text-emerald-800 font-bold">📥 ទទួលផ្ទេរពី៖ </span>
                                <strong>{item.visaTeamRobokName || '-'}</strong>
                                {item.operationType === 'transferTeam' && (
                                  <div className="mt-1 p-1.5 bg-indigo-50/60 border border-indigo-100 rounded text-[10px] space-y-0.5 text-indigo-900 font-normal leading-tight text-left">
                                    {item.recipientTeamName && (
                                      <div>
                                        <span className="font-semibold text-indigo-950">➡️ ទៅក្រុម (ខ្លួនឯង)៖ </span>
                                        <strong className="text-indigo-800">{item.recipientTeamName}</strong>
                                      </div>
                                    )}
                                    <div>
                                      <span className="font-semibold text-indigo-950">🏢 ការិយាល័យ៖ </span>
                                      <span className={item.officeApproved ? "text-emerald-700 font-semibold" : "text-amber-700 font-semibold"}>
                                        {item.officeApproved ? "✓ បានអនុញ្ញាត" : "✗ មិនទាន់អនុញ្ញាត"}
                                      </span>
                                    </div>
                                    {item.deptApprovalDate && (
                                      <div>
                                        <span className="font-semibold text-indigo-950">📅 នាយកដ្ឋាន៖ </span>
                                        <strong className="text-indigo-800">{formatKhmerDate(item.deptApprovalDate)}</strong>
                                      </div>
                                    )}
                                  </div>
                                )}
                              </div>
                            );
                          } else {
                            return (
                              <div>
                                {item.visaTeamRobokName || '-'}
                                {item.operationType === 'transferTeam' && (
                                  <div className="mt-1 p-1.5 bg-indigo-50/60 border border-indigo-100 rounded text-[10px] space-y-0.5 text-indigo-900 font-normal leading-tight text-left">
                                    {item.recipientTeamName && (
                                      <div>
                                        <span className="font-semibold text-indigo-950">➡️ ទៅក្រុម៖ </span>
                                        <strong className="text-indigo-800">{item.recipientTeamName}</strong>
                                      </div>
                                    )}
                                    <div>
                                      <span className="font-semibold text-indigo-950">🏢 ការិយាល័យ៖ </span>
                                      <span className={item.officeApproved ? "text-emerald-700 font-semibold" : "text-amber-700 font-semibold"}>
                                        {item.officeApproved ? "✓ បានអនុញ្ញាត" : "✗ មិនទាន់អនុញ្ញាត"}
                                      </span>
                                    </div>
                                    {item.deptApprovalDate && (
                                      <div>
                                        <span className="font-semibold text-indigo-950">📅 នាយកដ្ឋាន៖ </span>
                                        <strong className="text-indigo-800">{formatKhmerDate(item.deptApprovalDate)}</strong>
                                      </div>
                                    )}
                                  </div>
                                )}
                              </div>
                            );
                          }
                        })()}
                      </div>
                    </td>
                    {stockType === 'sticker' && (
                      <td className="py-2.5 px-3 border-r border-gray-200 text-center font-bold align-middle">
                        {item.visaType ? (
                          <span className="inline-block px-2 py-0.5 bg-blue-50 text-blue-700 rounded border border-blue-200 text-xs font-bold">
                            {item.visaType}
                          </span>
                        ) : (
                          '-'
                        )}
                      </td>
                    )}
                    <td className="py-2.5 px-3 border-r border-gray-200 text-center font-bold text-[#007bff] align-middle">
                      {item.quantityBundles ? item.quantityBundles.toLocaleString() : item.totalSheets ? item.totalSheets.toLocaleString() : '-'}
                      {item.quantityBundles ? (
                        <span className="block text-[10px] text-gray-400 font-normal">
                          ({item.quantityBundles} {item.stockType === 'evisa' ? 'ដុំ' : 'សន្លឹក'})
                        </span>
                      ) : null}
                    </td>
                    <td className="py-2.5 px-3 border-r border-gray-200 text-center font-mono font-medium align-middle">
                      {item.startSerial || '-'}
                    </td>
                    <td className="py-2.5 px-3 border-r border-gray-200 text-center font-mono font-medium align-middle">
                      {item.endSerial || '-'}
                    </td>
                    <td className="py-2.5 px-3 border-r border-gray-200 align-middle">
                      {item.operationType === 'issueTeam' ? (
                        <span>
                          {item.requestedRankName ? `${item.requestedRankName} ` : ''}
                          {item.requesterName || '-'}
                        </span>
                      ) : (
                        '-'
                      )}
                    </td>
                    <td className="py-2.5 px-3 border-r border-gray-200 align-middle">
                      {item.operationType === 'issueTeam' ? (
                        <span>
                          {item.collectorRoleName ? `[${item.collectorRoleName}] ` : ''}
                          {item.collectorName || '-'}
                        </span>
                      ) : (
                        '-'
                      )}
                    </td>
                    <td className="py-2.5 px-3 text-center print:hidden align-middle">
                      <div className="flex items-center justify-center gap-1.5">
                        {isSecondary && (item.operationType === 'issueTeam' || item.operationType === 'transferTeam') && (
                          <button
                            onClick={() => {
                              setSelectedPdfRecord(item);
                              setViewMode('handoverWorkspace');
                            }}
                            className="px-2 py-1 text-red-700 bg-red-50 hover:bg-red-100 border border-red-200 rounded text-[11px] font-bold transition cursor-pointer flex items-center gap-1"
                            title="បើក ផ្ទាំងការងារលិខិតប្រគល់ទទួល (PDF)"
                          >
                            <FileText className="w-3.5 h-3.5 text-red-600" />
                            <span>PDF</span>
                          </button>
                        )}
                        {stockType === 'sticker' && (item.operationType === 'issueTeam' || item.operationType === 'transferTeam') && onReturnIssuedStock && (
                          <button
                            onClick={() => {
                              if (window.confirm(`តើអ្នកពិតជាចង់បង្វិលការបើកផ្តល់ប្រភេទ ${item.visaType || ''} (${item.totalSheets || item.quantityBundles || 0} សន្លឹក) ត្រឡប់ចូលស្តុកជាក់ស្តែងវិញមែនទេ?`)) {
                                onReturnIssuedStock(item);
                              }
                            }}
                            className="px-2 py-1 text-amber-800 bg-amber-50 hover:bg-amber-100 border border-amber-200 rounded text-[11px] font-bold transition cursor-pointer flex items-center gap-1"
                            title="បង្វិលទិន្នន័យបើកផ្តល់នេះ ត្រឡប់ចូលក្នុងស្តុកជាក់ស្តែងវិញ"
                          >
                            <RotateCcw className="w-3.5 h-3.5 text-amber-600" />
                            <span>បង្វិល</span>
                          </button>
                        )}
                            <>
                              <button
                                onClick={() => handleStartEdit(item)}
                                className="p-1 text-blue-600 hover:text-blue-800 hover:bg-blue-50 rounded transition cursor-pointer"
                                title="កែប្រែ"
                              >
                                <Pencil className="w-3.5 h-3.5" />
                              </button>
                              <button
                                onClick={() => setDeletingRecord(item)}
                                className="p-1 text-red-500 hover:text-red-700 hover:bg-red-50 rounded transition cursor-pointer"
                                title="លុប"
                              >
                                <Trash2 className="w-3.5 h-3.5" />
                              </button>
                            </>
                      </div>
                    </td>
                  </tr>
                );
              })
            )}
            </tbody>
          </table>
        </div>

        {/* Pagination Bar */}
        {filteredRecords.length > 0 && (
          <div className="flex flex-wrap items-center justify-between gap-2 px-3 py-2.5 bg-gray-50 border-t border-gray-200 text-xs text-gray-600 print:hidden">
            <div className="flex items-center gap-2">
              <span>
                បង្ហាញជួរទី{' '}
                <b className="text-gray-900">
                  {filteredRecords.length === 0 ? 0 : (safeCurrentTablePage - 1) * (tablePageSize >= 999999 ? 0 : tablePageSize) + 1}
                </b>{' '}
                ដល់{' '}
                <b className="text-gray-900">
                  {Math.min(safeCurrentTablePage * tablePageSize, filteredRecords.length)}
                </b>{' '}
                នៃសរុប <b className="text-blue-700">{filteredRecords.length.toLocaleString()}</b> ជួរ
              </span>

              <span className="text-gray-300">|</span>

              <div className="flex items-center gap-1">
                <span>ក្នុងមួយទំព័រ:</span>
                <select
                  value={tablePageSize}
                  onChange={(e) => {
                    setTablePageSize(Number(e.target.value));
                    setTablePage(1);
                  }}
                  className="bg-white border border-gray-300 rounded px-1.5 py-0.5 text-xs text-gray-800 focus:outline-none focus:border-blue-500 cursor-pointer font-bold"
                >
                  <option value={25}>25</option>
                  <option value={50}>50</option>
                  <option value={100}>100</option>
                  <option value={200}>200</option>
                  <option value={500}>500</option>
                  <option value={999999}>ទាំងអស់</option>
                </select>
              </div>
            </div>

            {totalTablePages > 1 && (
              <div className="flex items-center gap-1">
                <button
                  type="button"
                  disabled={safeCurrentTablePage <= 1}
                  onClick={() => setTablePage(1)}
                  className="px-2 py-0.5 rounded border border-gray-300 bg-white text-gray-700 hover:bg-gray-100 disabled:opacity-40 disabled:cursor-not-allowed text-xs font-semibold cursor-pointer"
                >
                  « ដើម
                </button>
                <button
                  type="button"
                  disabled={safeCurrentTablePage <= 1}
                  onClick={() => setTablePage((p) => Math.max(1, p - 1))}
                  className="px-2 py-0.5 rounded border border-gray-300 bg-white text-gray-700 hover:bg-gray-100 disabled:opacity-40 disabled:cursor-not-allowed text-xs font-semibold cursor-pointer"
                >
                  ‹ មុន
                </button>

                <span className="px-2 font-bold text-gray-800">
                  ទំព័រ {safeCurrentTablePage} / {totalTablePages}
                </span>

                <button
                  type="button"
                  disabled={safeCurrentTablePage >= totalTablePages}
                  onClick={() => setTablePage((p) => Math.min(totalTablePages, p + 1))}
                  className="px-2 py-0.5 rounded border border-gray-300 bg-white text-gray-700 hover:bg-gray-100 disabled:opacity-40 disabled:cursor-not-allowed text-xs font-semibold cursor-pointer"
                >
                  បន្ទាប់ ›
                </button>
                <button
                  type="button"
                  disabled={safeCurrentTablePage >= totalTablePages}
                  onClick={() => setTablePage(totalTablePages)}
                  className="px-2 py-0.5 rounded border border-gray-300 bg-white text-gray-700 hover:bg-gray-100 disabled:opacity-40 disabled:cursor-not-allowed text-xs font-semibold cursor-pointer"
                >
                  ចុង »
                </button>
              </div>
            )}
          </div>
        )}
      </div>
      )}

      {/* Delete Stock Record Confirmation Modal */}
      {deletingRecord && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-2xs flex items-center justify-center p-4 print:hidden">
          <div className="bg-white rounded-md border border-gray-300 border-t-4 border-t-[#dc3545] shadow-xl w-full max-w-sm p-5 animate-fade">
            <div className="flex items-center gap-3 text-[#dc3545] mb-2">
              <div className="w-8 h-8 rounded-full bg-rose-100 flex items-center justify-center shrink-0">
                <Trash2 className="w-4 h-4 text-[#dc3545]" />
              </div>
              <h3 className="text-sm font-bold text-gray-800">បញ្ជាក់ការលុបទិន្នន័យស្តុក</h3>
            </div>
            <p className="text-xs text-gray-600 mb-4 leading-relaxed">
              តើអ្នកពិតជាចង់លុបទិន្នន័យស្តុក (<span className="font-bold text-gray-800">{deletingRecord.operationType === 'testPrintK2' ? 'ទិដ្ឋាការសាកក២' : deletingRecord.operationType === 'openK1' ? 'ការបញ្ចូលស្តុក' : deletingRecord.operationType === 'useTeam' ? 'ការប្រើប្រាស់តាមក្រុម' : deletingRecord.operationType === 'issueTeam' ? 'ការបើកផ្តល់តាមក្រុម' : 'ការបើកផ្តល់តាមក្រុម'}</span> - <span className="font-bold text-gray-800">{deletingRecord.quantityBundles || deletingRecord.totalSheets || 0} {deletingRecord.stockType === 'evisa' ? 'ដុំ' : 'សន្លឹក'}</span>) នេះចេញមែនទេ?
            </p>
            <div className="flex items-center justify-end gap-2">
              <button
                type="button"
                onClick={() => setDeletingRecord(null)}
                className="min-w-[100px] bg-[#6c757d] hover:bg-[#5a6268] text-white px-3.5 py-1.5 rounded text-xs font-bold transition flex items-center justify-center cursor-pointer"
              >
                <span>បោះបង់</span>
              </button>
              <button
                type="button"
                onClick={handleConfirmDelete}
                className="min-w-[100px] bg-[#dc3545] hover:bg-[#c82333] text-white px-3.5 py-1.5 rounded text-xs font-bold shadow-2xs transition flex items-center justify-center gap-1.5 cursor-pointer"
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>លុបចេញ</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Delete All / Filtered Stock Records Confirmation Modal */}
      {isDeleteAllModalOpen && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-2xs flex items-center justify-center p-4 print:hidden">
          <div className="bg-white rounded-md border border-gray-300 border-t-4 border-t-[#dc3545] shadow-xl w-full max-w-md p-5 animate-fade">
            <div className="flex items-center gap-3 text-[#dc3545] mb-2">
              <div className="w-8 h-8 rounded-full bg-rose-100 flex items-center justify-center shrink-0">
                <Trash2 className="w-4 h-4 text-[#dc3545]" />
              </div>
              <h3 className="text-sm font-bold text-gray-800">
                {filterOperationType !== 'all' && deleteScope === 'filtered'
                  ? `បញ្ជាក់ការលុបទិន្នន័យ «${opShortLabelMap[filterOperationType] || filterOperationType}»`
                  : 'បញ្ជាក់ការលុបទិន្នន័យស្តុក'}
              </h3>
            </div>

            {filterOperationType !== 'all' ? (
              <div className="space-y-3 mb-4">
                <p className="text-xs text-gray-700 leading-relaxed">
                  អ្នកបានជ្រើសរើសចម្រាញ់លើប្រភេទ «<span className="font-bold text-blue-700">{opLabelMap[filterOperationType] || filterOperationType}</span>»។ សូមជ្រើសរើសជម្រើសលុប៖
                </p>

                <div className="space-y-2 bg-gray-50 border border-gray-200 rounded p-3 text-xs">
                  <label className={`flex items-start gap-2.5 p-2 rounded cursor-pointer transition ${deleteScope === 'filtered' ? 'bg-rose-50/90 border border-rose-300' : 'hover:bg-gray-100 border border-transparent'}`}>
                    <input
                      type="radio"
                      name="deleteScopeOption"
                      checked={deleteScope === 'filtered'}
                      onChange={() => setDeleteScope('filtered')}
                      className="mt-0.5 text-rose-600 focus:ring-rose-500 cursor-pointer"
                    />
                    <div className="leading-tight">
                      <span className="font-bold text-rose-700">
                        លុបតែ «{opLabelMap[filterOperationType] || filterOperationType}»
                      </span>
                      <span className="block text-[11px] text-gray-600 mt-0.5">
                        លុបតែទិន្នន័យប្រភេទនេះចំនួន <strong className="text-rose-700 font-bold">{currentRecords.filter((r) => r.operationType === filterOperationType).length}</strong> ជួរ (រក្សាទុកទិន្នន័យប្រតិបត្តិការផ្សេងទៀតទាំងអស់)
                      </span>
                    </div>
                  </label>

                  <label className={`flex items-start gap-2.5 p-2 rounded cursor-pointer transition ${deleteScope === 'all' ? 'bg-rose-50/90 border border-rose-300' : 'hover:bg-gray-100 border border-transparent'}`}>
                    <input
                      type="radio"
                      name="deleteScopeOption"
                      checked={deleteScope === 'all'}
                      onChange={() => setDeleteScope('all')}
                      className="mt-0.5 text-rose-600 focus:ring-rose-500 cursor-pointer"
                    />
                    <div className="leading-tight">
                      <span className="font-bold text-gray-800">
                        លុបទិន្នន័យទាំងអស់នៃ{stockType === 'evisa' ? 'ក្រដាសអនុម័ត' : 'សន្លឹកទិដ្ឋាការ'}
                      </span>
                      <span className="block text-[11px] text-gray-600 mt-0.5">
                        លុបគ្រប់ប្រតិបត្តិការទាំងអស់សរុប <strong className="text-red-600 font-bold">{currentRecords.length}</strong> ជួរ
                      </span>
                    </div>
                  </label>
                </div>

                <span className="text-red-600 text-[11px] font-semibold block">
                  ⚠️ សកម្មភាពនេះមិនអាចត្រឡប់ក្រោយបានឡើយ!
                </span>
              </div>
            ) : (
              <p className="text-xs text-gray-600 mb-4 leading-relaxed">
                តើអ្នកពិតជាចង់លុបទិន្នន័យស្តុក ({stockType === 'evisa' ? 'ក្រដាសអនុម័តទិដ្ឋាការអេឡិចត្រូនិក' : 'សន្លឹកទិដ្ឋាការស្អិត'}) ទាំងអស់ចំនួន <span className="font-bold text-red-600">{currentRecords.length}</span> កំណត់ត្រានេះចេញមែនទេ? <br />
                <span className="text-red-500 font-semibold block mt-1">⚠️ សកម្មភាពនេះមិនអាចត្រឡប់ក្រោយបានឡើយ!</span>
              </p>
            )}

            <div className="flex items-center justify-end gap-2">
              <button
                type="button"
                onClick={() => setIsDeleteAllModalOpen(false)}
                className="min-w-[100px] bg-[#6c757d] hover:bg-[#5a6268] text-white px-3.5 py-1.5 rounded text-xs font-bold transition flex items-center justify-center cursor-pointer"
              >
                <span>បោះបង់</span>
              </button>
              <button
                type="button"
                onClick={handleConfirmDeleteAll}
                className="min-w-[100px] bg-[#dc3545] hover:bg-[#c82333] text-white px-3.5 py-1.5 rounded text-xs font-bold shadow-2xs transition flex items-center justify-center gap-1.5 cursor-pointer"
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>
                  {deleteScope === 'filtered' && filterOperationType !== 'all'
                    ? `លុប «${opShortLabelMap[filterOperationType] || filterOperationType}» (${currentRecords.filter((r) => r.operationType === filterOperationType).length})`
                    : `លុបទាំងអស់ (${currentRecords.length})`}
                </span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Stock PDF Modal - Only for Office (isSecondary) */}
      {isSecondary && selectedPdfRecord && (
        <StockPdfModal
          record={{
            ...selectedPdfRecord,
            stockType: selectedPdfRecord.stockType || stockType,
          }}
          allRecords={stockRecords}
          categories={categories}
          officers={officers}
          userName={userName}
          onClose={() => setSelectedPdfRecord(null)}
        />
      )}

      {/* Stock Excel Import Modal */}
      <StockImportModal
        isOpen={isImportModalOpen}
        onClose={() => setIsImportModalOpen(false)}
        stockType={stockType}
        categories={categories}
        existingRecords={stockRecords}
        selectedOperationType={filterOperationType}
        onImportRecords={handleBatchImportRecords}
        onShowToast={(msg) => onShowToast(msg)}
      />
    </div>
  );
};
