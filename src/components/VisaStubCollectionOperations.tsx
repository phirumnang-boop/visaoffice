import React, { useState, useEffect, useMemo } from 'react';
import { CategoriesState, StockRecord, UserRole } from '../types';
import * as XLSX from 'xlsx';
import {
  Calendar,
  Download,
  Printer,
  Search,
  Trash2,
  X,
  CheckCircle2,
  Plus,
  Database,
  Edit3,
  BookmarkCheck,
  FileSpreadsheet,
  RotateCcw,
  Eye,
  ArrowRightLeft,
  RefreshCw,
  Send,
  Check,
  AlertCircle,
} from 'lucide-react';
import {
  OFFICIAL_29_TEAMS,
  normalizeTeamName,
  normalizeDateToISO,
} from '../utils/teamNormalization';
import { printA4Document } from '../utils/printHelper';
import { CustomDatePicker } from './CustomDatePicker';

export interface StubEntry {
  id: string;
  quantity: string;
  startSerial: string;
  endSerial: string;
}

export interface StubCollectionRecord {
  id: string;
  date: string;
  teamName: string;
  collectedBy: string;
  remarks?: string;
  values: Record<
    string,
    {
      quantity: string;
      startSerial: string;
      endSerial: string;
      entries: StubEntry[];
    }
  >;
  totalSheets: number;
  createdAt: string;
  updatedAt?: string;
}

interface VisaStubCollectionOperationsProps {
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
  onShowToast: (msg: string, type: 'success' | 'error' | 'info') => void;
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
    if (nextVal < 0n) return trimmed;
    const nextDigits = nextVal.toString().padStart(digits.length, '0');
    return `${prefix}${nextDigits}${suffix}`;
  } catch {
    const n = parseInt(digits, 10);
    if (isNaN(n)) return trimmed;
    const nextDigits = String(n + step).padStart(digits.length, '0');
    return `${prefix}${nextDigits}${suffix}`;
  }
}

function calculateQtyFromSerials(start: string, end: string): string {
  if (!start || !end) return '';
  const mStart = start.trim().match(/^(.*?)(\d+)(.*?)$/);
  const mEnd = end.trim().match(/^(.*?)(\d+)(.*?)$/);
  if (!mStart || !mEnd) return '';
  try {
    const s = BigInt(mStart[2]);
    const e = BigInt(mEnd[2]);
    if (e >= s) {
      return (e - s + 1n).toString();
    }
  } catch {
    const s = parseInt(mStart[2], 10);
    const e = parseInt(mEnd[2], 10);
    if (!isNaN(s) && !isNaN(e) && e >= s) {
      return String(e - s + 1);
    }
  }
  return '';
}

const buildInitialFormData = (): Record<string, StubEntry[]> => {
  const initial: Record<string, StubEntry[]> = {};
  VISA_TYPES.forEach((vt) => {
    initial[vt.id] = [
      {
        id: `${vt.id}-0`,
        quantity: '',
        startSerial: '',
        endSerial: '',
      },
    ];
  });
  return initial;
};

export const VisaStubCollectionOperations: React.FC<VisaStubCollectionOperationsProps> = ({
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
  // Date State
  const [selectedDate, setSelectedDate] = useState<string>(() => {
    return (
      localStorage.getItem('app_visa_stub_selected_date') ||
      new Date().toISOString().split('T')[0]
    );
  });

  // Date Range State for Matrix Table (តារាងស្ថិតិប្រមូលគល់សន្លឹក (Sticker))
  const [matrixStartDate, setMatrixStartDate] = useState<string>(() => {
    return (
      localStorage.getItem('app_visa_stub_matrix_start_date') ||
      localStorage.getItem('app_visa_stub_selected_date') ||
      new Date().toISOString().split('T')[0]
    );
  });

  const [matrixEndDate, setMatrixEndDate] = useState<string>(() => {
    return (
      localStorage.getItem('app_visa_stub_matrix_end_date') ||
      localStorage.getItem('app_visa_stub_selected_date') ||
      new Date().toISOString().split('T')[0]
    );
  });

  const [selectedTeam, setSelectedTeam] = useState<string>(() => {
    return localStorage.getItem('app_visa_stub_selected_team') || '';
  });

  const [remarks, setRemarks] = useState<string>('');
  const [editingRecordId, setEditingRecordId] = useState<string | null>(null);

  // Form entries for all visa types
  const [formData, setFormData] = useState<Record<string, StubEntry[]>>(buildInitialFormData);

  // Saved records in LocalStorage
  const [records, setRecords] = useState<StubCollectionRecord[]>(() => {
    try {
      const saved = localStorage.getItem('app_visa_stub_collections_v1');
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed)) return parsed;
      }
    } catch (e) {
      console.error(e);
    }
    return [];
  });

  // Team list taken from categories (visaTeamsRobok) or fallback to official teams, plus any team names from records
  const allTeams = useMemo(() => {
    const list: string[] = [];
    if (categories?.visaTeamsRobok && categories.visaTeamsRobok.length > 0) {
      list.push(...categories.visaTeamsRobok.map((t) => t.name.trim()).filter(Boolean));
    } else {
      list.push(...OFFICIAL_29_TEAMS);
    }
    records.forEach((r) => {
      if (r.teamName?.trim() && !list.includes(r.teamName.trim())) {
        list.push(r.teamName.trim());
      }
    });
    (stockRecords || []).forEach((sr) => {
      const name = (sr.visaTeamRobokName || (sr as any).teamName || '').trim();
      if (name && !list.includes(name)) {
        list.push(name);
      }
    });
    return Array.from(new Set(list));
  }, [categories?.visaTeamsRobok, records, stockRecords]);

  // Combine local records with any stockRecords that represent stub collection in the main stock data
  const effectiveRecords = useMemo(() => {
    const map = new Map<string, StubCollectionRecord>();
    records.forEach((r) => {
      const key = `${normalizeDateToISO(r.date)}_${normalizeTeamName(r.teamName)}`;
      map.set(key, r);
    });

    // Group stockRecords by date & team for stub collection operations
    const stockStubGroups = new Map<string, StockRecord[]>();
    (stockRecords || []).forEach((sr) => {
      const isStub =
        sr.sourceFrom?.includes('គល់សន្លឹក') ||
        sr.remarks?.includes('គល់សន្លឹក') ||
        sr.id.startsWith('stock-stub-') ||
        sr.operationType === 'returnStub';
      if (!isStub) return;

      const team = sr.visaTeamRobokName || (sr as any).teamName || '';
      const date = sr.date || '';
      if (!team || !date) return;

      const key = `${normalizeDateToISO(date)}_${normalizeTeamName(team)}`;
      if (!stockStubGroups.has(key)) {
        stockStubGroups.set(key, []);
      }
      stockStubGroups.get(key)!.push(sr);
    });

    // If there are stock records in main data store not present in local records, synthesize them
    stockStubGroups.forEach((items, key) => {
      if (!map.has(key)) {
        const first = items[0];
        const date = first.date;
        const teamName = first.visaTeamRobokName || (first as any).teamName || '';
        const values: Record<
          string,
          { quantity: string; startSerial: string; endSerial: string; entries: StubEntry[] }
        > = {};
        let totalSheets = 0;

        items.forEach((item) => {
          const vt = (item.visaType || '').toUpperCase();
          const q = item.totalSheets || item.quantityBundles || 0;
          totalSheets += q;
          if (!values[vt]) {
            values[vt] = {
              quantity: String(q),
              startSerial: item.startSerial || '',
              endSerial: item.endSerial || '',
              entries: [
                {
                  id: item.id,
                  quantity: String(q),
                  startSerial: item.startSerial || '',
                  endSerial: item.endSerial || '',
                },
              ],
            };
          } else {
            const currentQ = parseInt(values[vt].quantity, 10) || 0;
            values[vt].quantity = String(currentQ + q);
            values[vt].entries.push({
              id: item.id,
              quantity: String(q),
              startSerial: item.startSerial || '',
              endSerial: item.endSerial || '',
            });
          }
        });

        map.set(key, {
          id: `stub-stock-${key}`,
          date,
          teamName,
          collectedBy: first.createdBy || 'ការិយាល័យ',
          remarks: first.remarks,
          values,
          totalSheets,
          createdAt: first.createdAt || new Date().toISOString(),
          updatedAt: first.createdAt || new Date().toISOString(),
        });
      }
    });

    return Array.from(map.values());
  }, [records, stockRecords]);

  // Sync to LocalStorage
  useEffect(() => {
    localStorage.setItem('app_visa_stub_collections_v1', JSON.stringify(records));
  }, [records]);

  useEffect(() => {
    localStorage.setItem('app_visa_stub_selected_date', selectedDate);
  }, [selectedDate]);

  useEffect(() => {
    localStorage.setItem('app_visa_stub_matrix_start_date', matrixStartDate);
  }, [matrixStartDate]);

  useEffect(() => {
    localStorage.setItem('app_visa_stub_matrix_end_date', matrixEndDate);
  }, [matrixEndDate]);

  useEffect(() => {
    if (selectedTeam) {
      localStorage.setItem('app_visa_stub_selected_team', selectedTeam);
    } else {
      localStorage.removeItem('app_visa_stub_selected_team');
    }
  }, [selectedTeam]);

  // Search & Filter in historical records table
  const [searchFilter, setSearchFilter] = useState<string>('');
  const [viewDetailRecord, setViewDetailRecord] = useState<StubCollectionRecord | null>(null);
  const [isPrintModalOpen, setIsPrintModalOpen] = useState<boolean>(false);
  const [isRefreshing, setIsRefreshing] = useState<boolean>(false);
  const [deleteConfirmTarget, setDeleteConfirmTarget] = useState<StubCollectionRecord | null>(null);
  const [deleteModalOpen, setDeleteModalOpen] = useState<boolean>(false);
  const [updateModalOpen, setUpdateModalOpen] = useState<boolean>(false);

  // Set of keys for records currently present in stockRecords ("ទិន្នន័យសន្លឹកទិដ្ឋាការ")
  const transferredStubKeys = useMemo(() => {
    const set = new Set<string>();
    (stockRecords || []).forEach((sr) => {
      const isStub =
        sr.sourceFrom?.includes('គល់សន្លឹក') ||
        sr.remarks?.includes('គល់សន្លឹក') ||
        sr.id.startsWith('stock-stub-') ||
        sr.operationType === 'returnStub';
      if (!isStub) return;

      const team = sr.visaTeamRobokName || (sr as any).teamName || '';
      const date = sr.date || '';
      if (!team || !date) return;

      set.add(`${normalizeDateToISO(date)}_${normalizeTeamName(team)}`);
    });
    return set;
  }, [stockRecords]);

  const isRecordTransferred = (rec: StubCollectionRecord): boolean => {
    const key = `${normalizeDateToISO(rec.date)}_${normalizeTeamName(rec.teamName)}`;
    if (transferredStubKeys.has(key)) return true;
    return (stockRecords || []).some(
      (sr) => sr.id.startsWith(`stock-stub-${rec.id}`) || (sr as any).stubId === rec.id
    );
  };

  const untransferredCount = useMemo(() => {
    return effectiveRecords.filter((r) => !isRecordTransferred(r)).length;
  }, [effectiveRecords, transferredStubKeys, stockRecords]);

  // Calculate total sheets currently typed in the form
  const formTotalSheets = useMemo(() => {
    let sum = 0;
    (Object.values(formData) as StubEntry[][]).forEach((list) => {
      list.forEach((entry) => {
        const q = parseInt(entry.quantity, 10);
        if (!isNaN(q) && q > 0) sum += q;
      });
    });
    return sum;
  }, [formData]);

  // Total rows across all visa types for grid arrow navigation
  const totalRows = useMemo(() => {
    let count = 0;
    VISA_TYPES.forEach((vt) => {
      count += formData[vt.id]?.length || 1;
    });
    return count;
  }, [formData]);

  // Navigation across textboxes using Arrow keys (Up, Down, Left, Right) and Enter
  const navigateGrid = (targetRow: number, targetCol: number) => {
    const target = document.querySelector<HTMLInputElement>(
      `input[data-stub-row="${targetRow}"][data-stub-col="${targetCol}"]`
    );
    if (target) {
      target.focus();
      target.select();
      target.scrollIntoView({ block: 'nearest', inline: 'nearest', behavior: 'smooth' });
    }
  };

  const handleKeyDown = (
    e: React.KeyboardEvent<HTMLInputElement>,
    flatRowIndex: number,
    colIndex: number,
    totalRowsCount: number
  ) => {
    const el = e.currentTarget;
    const isAllSelected =
      el.selectionStart === 0 &&
      el.selectionEnd === el.value.length;
    const isAtStart = el.selectionStart === 0 && el.selectionEnd === 0;
    const isAtEnd = el.selectionStart === el.value.length && el.selectionEnd === el.value.length;
    const isEmpty = !el.value;

    if (e.key === 'ArrowUp') {
      e.preventDefault();
      if (flatRowIndex > 0) {
        navigateGrid(flatRowIndex - 1, colIndex);
      }
    } else if (e.key === 'ArrowDown') {
      e.preventDefault();
      if (flatRowIndex < totalRowsCount - 1) {
        navigateGrid(flatRowIndex + 1, colIndex);
      }
    } else if (e.key === 'ArrowLeft') {
      if (isEmpty || isAllSelected || isAtStart) {
        e.preventDefault();
        if (colIndex > 0) {
          navigateGrid(flatRowIndex, colIndex - 1);
        } else if (flatRowIndex > 0) {
          navigateGrid(flatRowIndex - 1, 2); // Wrap to previous row's endSerial
        }
      }
    } else if (e.key === 'ArrowRight') {
      if (isEmpty || isAllSelected || isAtEnd) {
        e.preventDefault();
        if (colIndex < 2) {
          navigateGrid(flatRowIndex, colIndex + 1);
        } else if (flatRowIndex < totalRowsCount - 1) {
          navigateGrid(flatRowIndex + 1, 0); // Wrap to next row's quantity
        }
      }
    } else if (e.key === 'Enter') {
      e.preventDefault();
      if (colIndex < 2) {
        navigateGrid(flatRowIndex, colIndex + 1);
      } else if (flatRowIndex < totalRowsCount - 1) {
        navigateGrid(flatRowIndex + 1, 0);
      }
    }
  };

  // Handle Input Changes with auto-calculations
  const handleInputChange = (
    typeId: string,
    entryIndex: number,
    field: 'quantity' | 'startSerial' | 'endSerial',
    value: string
  ) => {
    setFormData((prev) => {
      const list = prev[typeId] || [];
      const current = list[entryIndex] || {
        id: `${typeId}-${entryIndex}`,
        quantity: '',
        startSerial: '',
        endSerial: '',
      };

      const updated = { ...current, [field]: value };
      const qNum = parseInt(updated.quantity, 10);

      if (field === 'startSerial') {
        if (updated.startSerial && !isNaN(qNum) && qNum > 0) {
          updated.endSerial = incrementSerial(updated.startSerial, qNum - 1);
        } else if (updated.startSerial && updated.endSerial) {
          const calcQty = calculateQtyFromSerials(updated.startSerial, updated.endSerial);
          if (calcQty) updated.quantity = calcQty;
        }
      } else if (field === 'endSerial') {
        if (updated.startSerial && updated.endSerial) {
          const calcQty = calculateQtyFromSerials(updated.startSerial, updated.endSerial);
          if (calcQty) updated.quantity = calcQty;
        }
      } else if (field === 'quantity') {
        if (updated.startSerial && !isNaN(qNum) && qNum > 0) {
          updated.endSerial = incrementSerial(updated.startSerial, qNum - 1);
        }
      }

      const newList = [...list];
      newList[entryIndex] = updated;
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
      const prevEnd = current?.endSerial?.trim() || '';

      const nextStart = prevEnd ? incrementSerial(prevEnd, 1) : '';

      const newEntry: StubEntry = {
        id: `${typeId}-${Date.now()}-${Math.random().toString(36).substr(2, 4)}`,
        quantity: '',
        startSerial: nextStart,
        endSerial: '',
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

  // Remove a row for a Visa Type
  const handleRemoveRow = (typeId: string, entryIndex: number) => {
    setFormData((prev) => {
      const list = prev[typeId] || [];
      if (list.length <= 1) {
        return {
          ...prev,
          [typeId]: [
            {
              id: `${typeId}-0`,
              quantity: '',
              startSerial: '',
              endSerial: '',
            },
          ],
        };
      }
      const newList = list.filter((_, idx) => idx !== entryIndex);
      return {
        ...prev,
        [typeId]: newList,
      };
    });
  };

  // Clear Form
  const handleClearForm = (showToast = true, clearTeam = false) => {
    setFormData(buildInitialFormData());
    setEditingRecordId(null);
    setRemarks('');
    if (clearTeam) {
      setSelectedTeam('');
    }
    if (showToast) {
      onShowToast('បានសម្អាតទម្រង់បញ្ចូលទិន្នន័យ', 'info');
    }
  };

  // Save / Insert Record
  const handleSaveRecord = (isUpdate = false) => {
    if (!selectedTeam) {
      onShowToast('សូមជ្រើសរើសក្រុមផ្តល់ទិដ្ឋាការជាមុនសិន!', 'error');
      return;
    }
    if (!selectedDate) {
      onShowToast('សូមជ្រើសរើសកាលបរិច្ឆេទប្រមូល!', 'error');
      return;
    }

    if (formTotalSheets === 0) {
      onShowToast('សូមបញ្ចូលចំនួនគល់សន្លឹកយ៉ាងតិច ១ សន្លឹក!', 'error');
      return;
    }

    // Build structured values object
    const valuesPayload: Record<
      string,
      {
        quantity: string;
        startSerial: string;
        endSerial: string;
        entries: StubEntry[];
      }
    > = {};

    (Object.entries(formData) as [string, StubEntry[]][]).forEach(([vtId, entries]) => {
      const validEntries = entries.filter((e) => {
        const q = parseInt(e.quantity, 10);
        return (!isNaN(q) && q > 0) || Boolean(e.startSerial.trim());
      });

      const totalQtyForVt = validEntries.reduce((acc, curr) => {
        const q = parseInt(curr.quantity, 10);
        return acc + (isNaN(q) ? 0 : q);
      }, 0);

      const firstStart = validEntries[0]?.startSerial || '';
      const lastEnd = validEntries[validEntries.length - 1]?.endSerial || '';

      valuesPayload[vtId] = {
        quantity: totalQtyForVt > 0 ? String(totalQtyForVt) : '',
        startSerial: firstStart,
        endSerial: lastEnd,
        entries: validEntries.length > 0 ? validEntries : entries,
      };
    });

    const normTeam = normalizeTeamName(selectedTeam);
    const normDate = normalizeDateToISO(selectedDate);

    // Check if record already exists for this team and date
    const existing = effectiveRecords.find(
      (r) =>
        normalizeTeamName(r.teamName) === normTeam &&
        normalizeDateToISO(r.date) === normDate
    );

    const recordId = editingRecordId || existing?.id || `stub-${Date.now()}-${Math.random().toString(36).substr(2, 5)}`;

    const newRecord: StubCollectionRecord = {
      id: recordId,
      date: selectedDate,
      teamName: selectedTeam,
      collectedBy: userName || 'មន្ត្រីការិយាល័យ',
      remarks: remarks.trim() || undefined,
      values: valuesPayload,
      totalSheets: formTotalSheets,
      createdAt: existing ? existing.createdAt : new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };

    // 1. Update local records state and synchronously persist to localStorage
    setRecords((prev) => {
      const filtered = prev.filter(
        (r) =>
          r.id !== recordId &&
          !(normalizeTeamName(r.teamName) === normTeam && normalizeDateToISO(r.date) === normDate)
      );
      const updated = [newRecord, ...filtered];
      try {
        localStorage.setItem('app_visa_stub_collections_v1', JSON.stringify(updated));
      } catch (e) {
        console.error(e);
      }
      return updated;
    });

    // 2. Sync to main application stock data: "ទិន្នន័យសន្លឹកទិដ្ឋាការ" (StockRecord)
    const stockItemsToSave: StockRecord[] = [];
    const robokTeamObj = categories?.visaTeamsRobok?.find(
      (t) => normalizeTeamName(t.name) === normTeam
    );

    (Object.entries(valuesPayload) as [string, any][]).forEach(([vtId, data]) => {
      const vt = VISA_TYPES.find((v) => v.id.toUpperCase() === vtId.toUpperCase()) || {
        name: vtId,
        id: vtId,
      };
      const entries: StubEntry[] = data.entries || [];
      entries.forEach((e, idx) => {
        const q = parseInt(e.quantity, 10);
        if (!isNaN(q) && q > 0) {
          stockItemsToSave.push({
            id: `stock-stub-${recordId}-${vt.id}-${idx}`,
            stockType: 'sticker',
            operationType: 'returnStub',
            date: selectedDate,
            time: new Date().toLocaleTimeString('en-US', { hour12: false, hour: '2-digit', minute: '2-digit' }),
            visaType: vt.name,
            quantityBundles: q,
            totalSheets: q,
            startSerial: e.startSerial || '',
            endSerial: e.endSerial || '',
            visaTeamRobokId: robokTeamObj?.id || `team-${selectedTeam}`,
            visaTeamRobokName: selectedTeam,
            sourceFrom: 'ប្រមូលគល់សន្លឹក',
            remarks: remarks.trim() ? `ប្រមូលគល់សន្លឹក: ${remarks.trim()}` : `ប្រមូលគល់សន្លឹកពីក្រុម ${selectedTeam}`,
            createdBy: userName || 'Admin',
            createdAt: new Date().toISOString(),
          });
        }
      });
    });

    // 3. Remove existing linked stock records from "ទិន្នន័យសន្លឹកទិដ្ឋាការ"
    const linkedOldStockIds = (stockRecords || [])
      .filter((sr) => {
        if (sr.id.startsWith(`stock-stub-${recordId}`)) return true;
        const srTeam = normalizeTeamName(sr.visaTeamRobokName || (sr as any).teamName || '');
        const srDate = normalizeDateToISO(sr.date || '');
        const isStub =
          sr.sourceFrom?.includes('គល់សន្លឹក') ||
          sr.remarks?.includes('គល់សន្លឹក') ||
          sr.operationType === 'returnStub' ||
          sr.id.includes(recordId);
        return (
          (srDate === normDate && srTeam === normTeam && isStub) ||
          sr.id.startsWith(`stock-stub-${recordId}`)
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

    // 4. Insert new/updated items into "ទិន្នន័យសន្លឹកទិដ្ឋាការ"
    if ((onBatchImportStockRecords || onAddStockRecord) && stockItemsToSave.length > 0) {
      if (onBatchImportStockRecords) {
        onBatchImportStockRecords(stockItemsToSave);
      } else if (onAddStockRecord) {
        stockItemsToSave.forEach((stk) => onAddStockRecord(stk));
      }
    }

    setEditingRecordId(recordId);
    if (isUpdate) {
      onShowToast(`បានកែសម្រួល និងធ្វើបច្ចុប្បន្នភាពទិន្នន័យគល់សន្លឹករបស់ «${selectedTeam}» ចំនួន ${formTotalSheets.toLocaleString()} សន្លឹកជោគជ័យ!`, 'success');
    } else {
      onShowToast(`បានរក្សាទុក និងផ្ទេរទៅ «ទិន្នន័យសន្លឹកទិដ្ឋាការ» ចំនួន ${formTotalSheets.toLocaleString()} សន្លឹកជោគជ័យ!`, 'success');
    }

    // Automatically empty form so user can enter the next team smoothly
    handleClearForm(false, true);
  };

  // Transfer single stub record to "ទិន្នន័យសន្លឹកទិដ្ឋាការ"
  const handleTransferSingleRecord = (rec: StubCollectionRecord) => {
    const normTeam = normalizeTeamName(rec.teamName);
    const normDate = normalizeDateToISO(rec.date);
    const robokTeamObj = categories?.visaTeamsRobok?.find(
      (t) => normalizeTeamName(t.name) === normTeam
    );

    const stockItemsToSave: StockRecord[] = [];
    if (rec.values) {
      Object.entries(rec.values).forEach(([vtId, data]) => {
        const vt = VISA_TYPES.find((v) => v.id.toUpperCase() === vtId.toUpperCase()) || {
          name: vtId,
          id: vtId,
        };
        const entries: StubEntry[] =
          Array.isArray(data.entries) && data.entries.length > 0
            ? data.entries
            : data.quantity
            ? [
                {
                  id: `${vtId}-0`,
                  quantity: data.quantity,
                  startSerial: data.startSerial || '',
                  endSerial: data.endSerial || '',
                },
              ]
            : [];

        entries.forEach((e, idx) => {
          const q = parseInt(e.quantity, 10);
          if (!isNaN(q) && q > 0) {
            stockItemsToSave.push({
              id: `stock-stub-${rec.id}-${vt.id}-${idx}`,
              stockType: 'sticker',
              operationType: 'returnStub',
              date: rec.date,
              time: new Date().toLocaleTimeString('en-US', {
                hour12: false,
                hour: '2-digit',
                minute: '2-digit',
              }),
              visaType: vt.name,
              quantityBundles: q,
              totalSheets: q,
              startSerial: e.startSerial || '',
              endSerial: e.endSerial || '',
              visaTeamRobokId: robokTeamObj?.id || `team-${rec.teamName}`,
              visaTeamRobokName: rec.teamName,
              sourceFrom: 'ប្រមូលគល់សន្លឹក',
              remarks: rec.remarks
                ? `ប្រមូលគល់សន្លឹក: ${rec.remarks}`
                : `ប្រមូលគល់សន្លឹកពីក្រុម ${rec.teamName}`,
              createdBy: rec.collectedBy || userName || 'Admin',
              createdAt: rec.createdAt || new Date().toISOString(),
            });
          }
        });
      });
    }

    if (stockItemsToSave.length === 0) {
      onShowToast('ពុំមានទិន្នន័យសន្លឹកទិដ្ឋាការដើម្បីផ្ទេរឡើយ!', 'error');
      return;
    }

    // Delete existing linked stock records first to prevent duplicates
    const linkedOldStockIds = (stockRecords || [])
      .filter((sr) => {
        if (sr.id.startsWith(`stock-stub-${rec.id}`)) return true;
        const srTeam = normalizeTeamName(sr.visaTeamRobokName || (sr as any).teamName || '');
        const srDate = normalizeDateToISO(sr.date || '');
        const isStub =
          sr.sourceFrom?.includes('គល់សន្លឹក') ||
          sr.remarks?.includes('គល់សន្លឹក') ||
          sr.operationType === 'returnStub' ||
          sr.id.includes(rec.id);
        return (
          (srDate === normDate && srTeam === normTeam && isStub) ||
          sr.id.startsWith(`stock-stub-${rec.id}`)
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

    if (onBatchImportStockRecords) {
      onBatchImportStockRecords(stockItemsToSave);
    } else if (onAddStockRecord) {
      stockItemsToSave.forEach((stk) => onAddStockRecord(stk));
    }

    handleClearForm(false, true);

    onShowToast(
      `បានផ្ទេរទិន្នន័យប្រមូលគល់សន្លឹក ${rec.totalSheets.toLocaleString()} សន្លឹក របស់ «${rec.teamName}» ទៅ «ទិន្នន័យសន្លឹកទិដ្ឋាការ» ជោគជ័យ!`,
      'success'
    );
  };

  // Transfer all stub records into "ទិន្នន័យសន្លឹកទិដ្ឋាការ"
  const handleTransferAllToStock = () => {
    const untransferred = effectiveRecords.filter((r) => !isRecordTransferred(r));
    const toTransfer = untransferred.length > 0 ? untransferred : effectiveRecords;

    if (toTransfer.length === 0) {
      onShowToast('ពុំមានទិន្នន័យប្រមូលគល់សន្លឹកសម្រាប់ផ្ទេរទេ!', 'info');
      return;
    }

    let allStockItems: StockRecord[] = [];
    const allLinkedOldIds: string[] = [];

    toTransfer.forEach((rec) => {
      const normTeam = normalizeTeamName(rec.teamName);
      const normDate = normalizeDateToISO(rec.date);
      const robokTeamObj = categories?.visaTeamsRobok?.find(
        (t) => normalizeTeamName(t.name) === normTeam
      );

      if (rec.values) {
        Object.entries(rec.values).forEach(([vtId, data]: [string, any]) => {
          const vt = VISA_TYPES.find((v) => v.id.toUpperCase() === vtId.toUpperCase()) || {
            name: vtId,
            id: vtId,
          };
          const entries: StubEntry[] =
            Array.isArray(data.entries) && data.entries.length > 0
              ? data.entries
              : data.quantity
              ? [
                  {
                    id: `${vtId}-0`,
                    quantity: data.quantity,
                    startSerial: data.startSerial || '',
                    endSerial: data.endSerial || '',
                  },
                ]
              : [];

          entries.forEach((e, idx) => {
            const q = parseInt(e.quantity, 10);
            if (!isNaN(q) && q > 0) {
              allStockItems.push({
                id: `stock-stub-${rec.id}-${vt.id}-${idx}`,
                stockType: 'sticker',
                operationType: 'returnStub',
                date: rec.date,
                time: new Date().toLocaleTimeString('en-US', {
                  hour12: false,
                  hour: '2-digit',
                  minute: '2-digit',
                }),
                visaType: vt.name,
                quantityBundles: q,
                totalSheets: q,
                startSerial: e.startSerial || '',
                endSerial: e.endSerial || '',
                visaTeamRobokId: robokTeamObj?.id || `team-${rec.teamName}`,
                visaTeamRobokName: rec.teamName,
                sourceFrom: 'ប្រមូលគល់សន្លឹក',
                remarks: rec.remarks
                  ? `ប្រមូលគល់សន្លឹក: ${rec.remarks}`
                  : `ប្រមូលគល់សន្លឹកពីក្រុម ${rec.teamName}`,
                createdBy: rec.collectedBy || userName || 'Admin',
                createdAt: rec.createdAt || new Date().toISOString(),
              });
            }
          });
        });
      }

      // Old linked IDs
      (stockRecords || []).forEach((sr) => {
        const srTeam = normalizeTeamName(sr.visaTeamRobokName || (sr as any).teamName || '');
        const srDate = normalizeDateToISO(sr.date || '');
        const isStub =
          sr.sourceFrom?.includes('គល់សន្លឹក') ||
          sr.remarks?.includes('គល់សន្លឹក') ||
          sr.operationType === 'returnStub' ||
          sr.id.includes(rec.id);
        if (
          (srDate === normDate && srTeam === normTeam && isStub) ||
          sr.id.startsWith(`stock-stub-${rec.id}`)
        ) {
          if (!allLinkedOldIds.includes(sr.id)) {
            allLinkedOldIds.push(sr.id);
          }
        }
      });
    });

    if (allStockItems.length === 0) {
      onShowToast('ពុំមានទិន្នន័យសម្រាប់ផ្ទេរទេ!', 'error');
      return;
    }

    if (allLinkedOldIds.length > 0) {
      if (onDeleteBatchStockRecords) {
        onDeleteBatchStockRecords(allLinkedOldIds);
      } else if (onDeleteStockRecord) {
        allLinkedOldIds.forEach((id) => onDeleteStockRecord(id));
      }
    }

    if (onBatchImportStockRecords) {
      onBatchImportStockRecords(allStockItems);
    } else if (onAddStockRecord) {
      allStockItems.forEach((stk) => onAddStockRecord(stk));
    }

    handleClearForm(false, true);

    onShowToast(
      `បានផ្ទេរទិន្នន័យប្រមូលគល់សន្លឹកទាំងអស់ (${allStockItems.length} ប្រតិបត្តិការ) ទៅ «ទិន្នន័យសន្លឹកទិដ្ឋាការ» ជោគជ័យ!`,
      'success'
    );
  };

  // Refresh data handler
  const handleRefreshData = () => {
    setIsRefreshing(true);
    try {
      const saved = localStorage.getItem('app_visa_stub_collections_v1');
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed)) {
          setRecords(parsed);
        }
      }
    } catch (e) {
      console.error(e);
    }
    setTimeout(() => {
      setIsRefreshing(false);
      onShowToast('បានផ្ទុកទិន្នន័យឡើងវិញជោគជ័យ!', 'success');
    }, 350);
  };

  // Search existing record for current selected team & date
  const handleSearch = () => {
    if (!selectedTeam) {
      onShowToast('សូមជ្រើសរើសក្រុមដើម្បីស្វែងរក!', 'error');
      return;
    }
    const normTeam = normalizeTeamName(selectedTeam);
    const normDate = normalizeDateToISO(selectedDate);

    const found = effectiveRecords.find(
      (r) =>
        normalizeTeamName(r.teamName) === normTeam &&
        normalizeDateToISO(r.date) === normDate
    );

    if (found) {
      loadRecordIntoForm(found);
      onShowToast(`បានរកឃើញទិន្នន័យប្រមូលគល់សន្លឹករបស់ក្រុម ${selectedTeam} ថ្ងៃ ${selectedDate}`, 'success');
    } else {
      onShowToast(`ពុំមានទិន្នន័យប្រមូលគល់សន្លឹកសម្រាប់ក្រុម ${selectedTeam} ថ្ងៃ ${selectedDate} ទេ`, 'info');
    }
  };

  // Load record into form
  const loadRecordIntoForm = (rec: StubCollectionRecord) => {
    setSelectedDate(rec.date);
    setSelectedTeam(rec.teamName);
    setRemarks(rec.remarks || '');
    setEditingRecordId(rec.id);

    const nextForm = buildInitialFormData();
    if (rec.values) {
      Object.entries(rec.values).forEach(([vtId, data]) => {
        const targetVtId =
          VISA_TYPES.find((v) => v.id.toUpperCase() === vtId.toUpperCase())?.id || vtId;
        if (Array.isArray(data.entries) && data.entries.length > 0) {
          nextForm[targetVtId] = data.entries.map((e, idx) => ({
            id: e.id || `${targetVtId}-${idx}`,
            quantity: e.quantity || '',
            startSerial: e.startSerial || '',
            endSerial: e.endSerial || '',
          }));
        } else if (data.quantity || data.startSerial || data.endSerial) {
          nextForm[targetVtId] = [
            {
              id: `${targetVtId}-0`,
              quantity: data.quantity || '',
              startSerial: data.startSerial || '',
              endSerial: data.endSerial || '',
            },
          ];
        }
      });
    }
    setFormData(nextForm);

    const formEl = document.getElementById('stub-collection-form-container');
    if (formEl) {
      formEl.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
    onShowToast(`បានទាញទិន្នន័យរបស់ «${rec.teamName}» មកក្នុង Form ដើម្បីកែសម្រួល!`, 'info');
  };

  // Update existing record
  const handleUpdateClick = () => {
    if (!selectedTeam) {
      onShowToast('សូមជ្រើសរើសក្រុមដើម្បីកែសម្រួល!', 'error');
      return;
    }
    if (!selectedDate) {
      onShowToast('សូមជ្រើសរើសកាលបរិច្ឆេទ!', 'error');
      return;
    }

    const normTeam = normalizeTeamName(selectedTeam);
    const normDate = normalizeDateToISO(selectedDate);
    const found = effectiveRecords.find(
      (r) =>
        (editingRecordId && r.id === editingRecordId) ||
        (normalizeTeamName(r.teamName) === normTeam && normalizeDateToISO(r.date) === normDate)
    );

    // If form is empty (0 sheets) but record exists in database, load it into form for editing!
    if (formTotalSheets === 0) {
      if (found) {
        loadRecordIntoForm(found);
        onShowToast(`បានទាញទិន្នន័យរបស់ «${found.teamName}» មកក្នុង Form ដើម្បីកែសម្រួល!`, 'info');
      } else {
        onShowToast(`ពុំទាន់មានទិន្នន័យប្រមូលគល់សន្លឹកសម្រាប់ក្រុម «${selectedTeam}» ថ្ងៃទី ${selectedDate} ទេ`, 'info');
      }
      return;
    }

    if (found && !editingRecordId) {
      setEditingRecordId(found.id);
    }

    setUpdateModalOpen(true);
  };

  const handleConfirmUpdate = () => {
    setUpdateModalOpen(false);
    handleSaveRecord(true);
  };

  // Delete record from toolbar
  const handleDeleteClick = () => {
    if (!selectedTeam) {
      onShowToast('សូមជ្រើសរើសក្រុមដើម្បីលុប!', 'error');
      return;
    }
    if (!selectedDate) {
      onShowToast('សូមជ្រើសរើសកាលបរិច្ឆេទ!', 'error');
      return;
    }

    const normTeam = normalizeTeamName(selectedTeam);
    const normDate = normalizeDateToISO(selectedDate);

    const target = effectiveRecords.find(
      (r) =>
        r.id === editingRecordId ||
        (normalizeTeamName(r.teamName) === normTeam && normalizeDateToISO(r.date) === normDate)
    );

    if (!target) {
      if (formTotalSheets > 0) {
        handleClearForm();
        onShowToast(`បានសម្អាតទម្រង់បញ្ចូលទិន្នន័យរបស់ក្រុម «${selectedTeam}»`, 'info');
        return;
      }
      onShowToast(`ពុំមានទិន្នន័យប្រមូលគល់សន្លឹកសម្រាប់ក្រុម «${selectedTeam}» ថ្ងៃទី ${selectedDate} ដើម្បីលុបឡើយ`, 'info');
      return;
    }

    setDeleteConfirmTarget(target);
    setDeleteModalOpen(true);
  };

  const handleConfirmDelete = () => {
    if (!deleteConfirmTarget) {
      setDeleteModalOpen(false);
      return;
    }
    executeDelete(deleteConfirmTarget);
    setDeleteModalOpen(false);
  };

  // Delete specific record by ID from table
  const handleDeleteRecordById = (id: string) => {
    const target = effectiveRecords.find((r) => r.id === id);
    if (!target) return;
    setDeleteConfirmTarget(target);
    setDeleteModalOpen(true);
  };

  // Execute deletion without window.confirm (works inside iframes)
  const executeDelete = (target: StubCollectionRecord) => {
    const targetNormTeam = normalizeTeamName(target.teamName);
    const targetNormDate = normalizeDateToISO(target.date);

    // 1. Delete linked stock records from "ទិន្នន័យសន្លឹកទិដ្ឋាការ"
    const linkedStockIds = (stockRecords || [])
      .filter((sr) => {
        if (sr.id.startsWith(`stock-stub-${target.id}`)) return true;
        const srTeam = normalizeTeamName(sr.visaTeamRobokName || (sr as any).teamName || '');
        const srDate = normalizeDateToISO(sr.date || '');
        const isStub =
          sr.sourceFrom?.includes('គល់សន្លឹក') ||
          sr.remarks?.includes('គល់សន្លឹក') ||
          sr.operationType === 'returnStub' ||
          sr.id.includes(target.id);
        return (
          (srDate === targetNormDate && srTeam === targetNormTeam && isStub) ||
          sr.id.startsWith(`stock-stub-${target.id}`)
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

    // 2. Delete from local state and localStorage
    setRecords((prev) => {
      const updated = prev.filter((r) => {
        if (r.id === target.id) return false;
        if (
          normalizeDateToISO(r.date) === targetNormDate &&
          normalizeTeamName(r.teamName) === targetNormTeam
        )
          return false;
        return true;
      });
      try {
        localStorage.setItem('app_visa_stub_collections_v1', JSON.stringify(updated));
      } catch (e) {
        console.error(e);
      }
      return updated;
    });

    if (
      editingRecordId === target.id ||
      (normalizeDateToISO(selectedDate) === targetNormDate &&
        normalizeTeamName(selectedTeam) === targetNormTeam)
    ) {
      handleClearForm(false);
    }

    setDeleteConfirmTarget(null);
    onShowToast(`បានលុបទិន្នន័យប្រមូលគល់សន្លឹករបស់ «${target.teamName}» ពី «ទិន្នន័យសន្លឹកទិដ្ឋាការ» ជោគជ័យ!`, 'success');
  };

  // Matrix calculation for the selected date range across all teams (តារាងស្ថិតិប្រមូលគល់សន្លឹកប្រចាំថ្ងៃ Sticker)
  const dateMatrix = useMemo(() => {
    const normStart = normalizeDateToISO(matrixStartDate) || '1970-01-01';
    const normEnd = normalizeDateToISO(matrixEndDate) || '2099-12-31';
    const [effectiveStart, effectiveEnd] =
      normStart <= normEnd ? [normStart, normEnd] : [normEnd, normStart];

    const rangeRecords = effectiveRecords.filter((r) => {
      const d = normalizeDateToISO(r.date);
      if (!d) return false;
      return d >= effectiveStart && d <= effectiveEnd;
    });

    const teamRows: Array<{
      teamName: string;
      counts: Record<string, number>;
      total: number;
      recordId?: string;
      recordsCount: number;
      dates: string[];
    }> = [];

    const dateTotals: Record<string, number> = {};
    VISA_TYPES.forEach((vt) => {
      dateTotals[vt.id] = 0;
    });
    let dateGrandTotal = 0;

    allTeams.forEach((team) => {
      const normT = normalizeTeamName(team);
      const teamRecs = rangeRecords.filter((r) => normalizeTeamName(r.teamName) === normT);

      const counts: Record<string, number> = {};
      let rowTotal = 0;
      const datesSet = new Set<string>();

      VISA_TYPES.forEach((vt) => {
        let count = 0;
        teamRecs.forEach((rec) => {
          if (rec.date) datesSet.add(rec.date);
          if (rec.values?.[vt.id]) {
            const v = rec.values[vt.id];
            if (Array.isArray(v.entries)) {
              count += v.entries.reduce((acc, e) => {
                const q = parseInt(e.quantity, 10);
                return acc + (isNaN(q) ? 0 : q);
              }, 0);
            } else {
              const q = parseInt(v.quantity, 10);
              if (!isNaN(q)) count += q;
            }
          }
        });
        counts[vt.id] = count;
        rowTotal += count;
        dateTotals[vt.id] += count;
      });

      dateGrandTotal += rowTotal;

      teamRows.push({
        teamName: team,
        counts,
        total: rowTotal,
        recordId: teamRecs.length === 1 ? teamRecs[0].id : undefined,
        recordsCount: teamRecs.length,
        dates: Array.from(datesSet).sort(),
      });
    });

    // Only show teams that have collections (total > 0)
    const activeRows = teamRows
      .filter((r) => r.total > 0)
      .map((row, idx) => ({
        ...row,
        displayIndex: idx + 1,
      }));

    return {
      effectiveStart,
      effectiveEnd,
      allTeamRows: teamRows,
      rows: activeRows,
      totals: dateTotals,
      grandTotal: dateGrandTotal,
    };
  }, [effectiveRecords, matrixStartDate, matrixEndDate, allTeams]);

  // Selected Team numbers for the day
  const selectedTeamStats = useMemo(() => {
    if (!selectedTeam) return null;
    const row = dateMatrix.allTeamRows.find(
      (r) => normalizeTeamName(r.teamName) === normalizeTeamName(selectedTeam)
    );
    return row || null;
  }, [dateMatrix.allTeamRows, selectedTeam]);

  // Overall Statistics
  const overallStats = useMemo(() => {
    const totalAllSheets = effectiveRecords.reduce((acc, r) => acc + (r.totalSheets || 0), 0);
    const uniqueTeams = new Set(effectiveRecords.map((r) => normalizeTeamName(r.teamName))).size;
    const normSelected = normalizeDateToISO(selectedDate);
    const todaySheets = effectiveRecords
      .filter((r) => normalizeDateToISO(r.date) === normSelected)
      .reduce((acc, r) => acc + (r.totalSheets || 0), 0);
    return {
      totalAllSheets,
      uniqueTeams,
      todaySheets,
      recordsCount: effectiveRecords.length,
    };
  }, [effectiveRecords, selectedDate]);

  // Filtered saved records for historical list
  const filteredRecords = useMemo(() => {
    let list = [...effectiveRecords];
    if (searchFilter.trim()) {
      const q = searchFilter.trim().toLowerCase();
      list = list.filter(
        (r) =>
          r.teamName.toLowerCase().includes(q) ||
          r.date.includes(q) ||
          (r.remarks && r.remarks.toLowerCase().includes(q)) ||
          (r.collectedBy && r.collectedBy.toLowerCase().includes(q))
      );
    }
    return list.sort((a, b) => b.date.localeCompare(a.date));
  }, [effectiveRecords, searchFilter]);

  // Export to Excel
  const handleExportExcel = () => {
    if (effectiveRecords.length === 0) {
      onShowToast('មិនទាន់មានទិន្នន័យប្រមូលគល់សន្លឹកសម្រាប់ Export ទេ', 'info');
      return;
    }

    const headers = [
      'ល.រ',
      'កាលបរិច្ឆេទ',
      'ក្រុមផ្តល់ទិដ្ឋាការ',
      ...VISA_TYPES.map((v) => v.name),
      'សរុប (សន្លឹក)',
      'លម្អិតលេខកូដ',
      'អ្នកកត់ត្រា',
      'ចំណាំ',
    ];

    const data = filteredRecords.map((r, idx) => {
      const serialDetails: string[] = [];
      VISA_TYPES.forEach((vt) => {
        const val = r.values?.[vt.id];
        if (val) {
          if (Array.isArray(val.entries)) {
            val.entries.forEach((e) => {
              if (e.quantity && parseInt(e.quantity, 10) > 0) {
                serialDetails.push(
                  `${vt.name}: ${e.quantity} [${e.startSerial || '-'} ដល់ ${e.endSerial || '-'}]`
                );
              }
            });
          } else if (val.quantity) {
            serialDetails.push(
              `${vt.name}: ${val.quantity} [${val.startSerial || '-'} ដល់ ${val.endSerial || '-'}]`
            );
          }
        }
      });

      return [
        idx + 1,
        r.date,
        r.teamName,
        ...VISA_TYPES.map((vt) => {
          const val = r.values?.[vt.id];
          if (!val) return 0;
          if (Array.isArray(val.entries)) {
            return val.entries.reduce((acc, e) => acc + (parseInt(e.quantity, 10) || 0), 0);
          }
          return parseInt(val.quantity, 10) || 0;
        }),
        r.totalSheets,
        serialDetails.join(' | '),
        r.collectedBy,
        r.remarks || '',
      ];
    });

    const ws = XLSX.utils.aoa_to_sheet([
      ['តារាងស្ថិតិប្រមូលគល់សន្លឹកទិដ្ឋាការស្អិត (ការិយាល័យ)'],
      [`កាលបរិច្ឆេទគិតត្រឹម៖ ${selectedDate}`],
      [],
      headers,
      ...data,
    ]);

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Stub_Collection');
    XLSX.writeFile(wb, `Stub_Collection_Report_${selectedDate}.xlsx`);
    onShowToast('បានទាញយកឯកសារ Excel រួចរាល់!', 'success');
  };

  // Export Matrix Table to Excel for the selected date range
  const handleExportMatrixExcel = () => {
    if (dateMatrix.rows.length === 0) {
      onShowToast('មិនទាន់មានទិន្នន័យក្រុមដែលប្រគល់ក្នុងចន្លោះកាលបរិច្ឆេទនេះសម្រាប់ Export ទេ', 'info');
      return;
    }

    const headers = [
      'ល.រ',
      'ក្រុមផ្តល់ទិដ្ឋាការ',
      ...VISA_TYPES.map((v) => v.name),
      'សរុប (សន្លឹក)',
    ];

    const data = dateMatrix.rows.map((row, idx) => [
      idx + 1,
      row.teamName,
      ...VISA_TYPES.map((vt) => row.counts[vt.id] || 0),
      row.total,
    ]);

    const totalRow = [
      '',
      dateMatrix.effectiveStart === dateMatrix.effectiveEnd
        ? 'សរុបប្រមូលប្រចាំថ្ងៃ'
        : 'សរុបប្រមូលក្នុងចន្លោះកាលបរិច្ឆេទ',
      ...VISA_TYPES.map((vt) => dateMatrix.totals[vt.id] || 0),
      dateMatrix.grandTotal,
    ];

    const rangeLabel =
      dateMatrix.effectiveStart === dateMatrix.effectiveEnd
        ? `កាលបរិច្ឆេទ៖ ${dateMatrix.effectiveStart}`
        : `កាលបរិច្ឆេទចាប់ផ្តើម៖ ${dateMatrix.effectiveStart} ដល់ថ្ងៃ៖ ${dateMatrix.effectiveEnd}`;

    const ws = XLSX.utils.aoa_to_sheet([
      ['តារាងស្ថិតិប្រមូលគល់សន្លឹកប្រចាំថ្ងៃ (Sticker) — ក្រុមដែលប្រគល់'],
      [rangeLabel],
      [],
      headers,
      ...data,
      totalRow,
    ]);

    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Matrix_Stubs');
    const fileNameDate =
      dateMatrix.effectiveStart === dateMatrix.effectiveEnd
        ? dateMatrix.effectiveStart
        : `${dateMatrix.effectiveStart}_to_${dateMatrix.effectiveEnd}`;
    XLSX.writeFile(wb, `Daily_Stub_Matrix_${fileNameDate}.xlsx`);
    onShowToast('បានទាញយកតារាងស្ថិតិជា Excel រួចរាល់!', 'success');
  };

  return (
    <div className="w-full text-[#1E293B] font-sans pb-10">
      {/* Top Banner / Breadcrumb & Status */}
      <div className="mb-2 flex flex-wrap items-center justify-between gap-2 bg-linear-to-r from-blue-900 to-indigo-900 text-white px-3 py-2 rounded-lg shadow-sm border border-blue-800">
        <div className="flex items-center gap-2">
          <div className="p-1.5 bg-[#C6A15B] text-blue-950 rounded shadow-xs">
            <BookmarkCheck className="w-4 h-4" />
          </div>
          <div>
            <h1 className="text-sm sm:text-base font-bold flex items-center gap-2">
              <span>ប្រមូលគល់សន្លឹកទិដ្ឋាការ (សន្លឹកទិដ្ឋាការស្អិត)</span>
              <span className="text-[10px] font-medium bg-amber-400/30 text-amber-200 border border-amber-400/40 px-2 py-0.5 rounded-full">
                ស្តុកការិយាល័យ
              </span>
            </h1>
            <p className="text-[11px] text-blue-200">
              ទម្រង់កត់ត្រា និងគ្រប់គ្រងការប្រមូលគល់សន្លឹកទិដ្ឋាការដែលក្រុមបានប្រើប្រាស់រួច
            </p>
          </div>
        </div>

        {/* Global Stats Badges */}
        <div className="flex items-center gap-2 text-xs">
          <div className="bg-black/30 px-2.5 py-1 rounded border border-white/10 flex items-center gap-1.5">
            <span className="text-blue-200 text-[11px]">សរុបថ្ងៃនេះ៖</span>
            <span className="font-times font-bold text-yellow-300 text-sm">
              {overallStats.todaySheets.toLocaleString()}
            </span>
            <span
              className="text-[10px] text-white/70 font-siemreap"
              style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}
            >
              សន្លឹក
            </span>
          </div>

          <div className="bg-black/30 px-2.5 py-1 rounded border border-white/10 hidden sm:flex items-center gap-1.5">
            <span className="text-blue-200 text-[11px]">សរុបទាំងអស់៖</span>
            <span className="font-times font-bold text-emerald-300 text-sm">
              {overallStats.totalAllSheets.toLocaleString()}
            </span>
            <span
              className="text-[10px] text-white/70 font-siemreap"
              style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}
            >
              សន្លឹក
            </span>
          </div>

          <button
            type="button"
            onClick={handleExportExcel}
            className="px-2.5 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded text-xs font-medium flex items-center gap-1 cursor-pointer transition shadow-xs"
            title="ទាញយកជា Excel"
          >
            <Download className="w-3.5 h-3.5" />
            <span className="hidden md:inline">Excel</span>
          </button>

        </div>
      </div>

      {/* Main Two-Column Layout */}
      <div className="flex flex-col xl:flex-row items-start gap-3 w-full">
        {/* LEFT COLUMN: Input Form */}
        <div id="stub-collection-form-container" className="w-full xl:w-[480px] shrink-0 flex flex-col scroll-mt-20">
          <div className="bg-[#BFDBFE] border-2 border-[#1E40AF] rounded shadow-md overflow-hidden">
            {/* Header: Action Buttons */}
            <div className="bg-[#93C5FD] border-b-2 border-[#1E40AF] py-1.5 px-2.5 flex items-center justify-between gap-2">
              {/* 4 Action Buttons */}
              <div className="flex items-center gap-1.5 sm:gap-2">
                {/* Save (Red) */}
                <button
                  type="button"
                  onClick={() => handleSaveRecord(false)}
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

              {/* Title / Clear Button */}
              <div className="flex items-center gap-1">
                <button
                  type="button"
                  onClick={() => handleClearForm(true, true)}
                  className="px-2 py-1 bg-white/80 hover:bg-white text-blue-900 rounded border border-blue-300 text-[11px] font-bold flex items-center gap-1 cursor-pointer transition shadow-2xs"
                  title="សម្អាតទម្រង់"
                >
                  <RotateCcw className="w-3 h-3 text-blue-700" />
                  <span>សម្អាត</span>
                </button>
              </div>
            </div>

            {/* Date & Group Selector Bar */}
            <div className="grid grid-cols-12 border-b border-[#1E40AF] bg-[#DBEAFE] text-xs">
              {/* Date & Team Selection */}
              <div className="col-span-7 p-2 border-r border-[#1E40AF] flex flex-col justify-center space-y-1.5">
                <div className="flex items-center gap-2">
                  <span className="font-bold text-[#1E3A8A] whitespace-nowrap text-xs w-14">
                    កាលបរិច្ឆេទ
                  </span>
                  <div className="flex-1">
                    <CustomDatePicker
                      value={selectedDate}
                      onChange={(d) => setSelectedDate(d)}
                      className="text-xs text-center font-bold text-[#0F172A] py-1"
                    />
                  </div>
                </div>
                <div className="flex items-center gap-2">
                  <span className="font-bold text-[#1E3A8A] text-xs w-14">
                    ក្រុម
                  </span>
                  <select
                    value={selectedTeam}
                    onChange={(e) => setSelectedTeam(e.target.value)}
                    className="flex-1 bg-white border border-gray-400 rounded px-2 py-1 text-xs font-semibold shadow-inner outline-hidden focus:border-blue-600 truncate text-[#0F172A]"
                  >
                    <option value="">(ជ្រើសរើសក្រុម)</option>
                    {allTeams.map((t) => (
                      <option key={t} value={t}>
                        {t}
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              {/* Right Side Summary Box */}
              <div className="col-span-5 bg-[#0E7490] text-white text-center p-2 flex flex-col justify-center">
                <div className="text-sm font-bold">
                  <span className="font-times">{formTotalSheets.toLocaleString()}</span>{' '}
                  <span
                    className="font-bold font-siemreap"
                    style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}
                  >
                    សន្លឹក
                  </span>
                </div>
                <div className="text-[11px] font-medium font-times border-t border-b border-cyan-400/40 my-0.5 py-0.5">
                  {selectedDate}
                </div>
                <div
                  className="text-[11px] font-bold text-cyan-100 font-siemreap"
                  style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}
                >
                  គល់សន្លឹកទិដ្ឋាការស្អិត
                </div>
              </div>
            </div>

            {/* Column Headers for Form Textboxes: [ប្រភេទ, ចំនួន, ចាប់ពីលេខ, ដល់លេខ, សកម្មភាព] */}
            {/* Note: EXACTLY NO "លេខកូដចាស់" column as requested, but (+) button is kept! */}
            <div className="grid grid-cols-[42px_1fr_1.45fr_1.45fr_auto] gap-1 items-center bg-[#1E3A8A] text-white px-1.5 py-1.5 text-[11px] font-bold text-center border-b border-blue-900">
              <div className="text-center">ប្រភេទ</div>
              <div className="text-center">ចំនួន</div>
              <div className="text-center">ចាប់ពីលេខ</div>
              <div className="text-center">ដល់លេខ</div>
              <div className="w-12 text-center">បន្ថែម</div>
            </div>

            {/* Form Rows Table for Visa Types */}
            <div className="p-1.5 bg-[#E2E8F0] space-y-1 max-h-[560px] overflow-y-auto">
              {(() => {
                let runningFlatRowIndex = 0;
                return VISA_TYPES.map((vt) => {
                  const entries = formData[vt.id] || [
                    { id: `${vt.id}-0`, quantity: '', startSerial: '', endSerial: '' },
                  ];

                  return (
                    <div key={vt.id} className="space-y-1">
                      {entries.map((entry, entryIndex) => {
                        const flatRowIndex = runningFlatRowIndex++;
                        return (
                          <div
                            key={entry.id || `${vt.id}-${entryIndex}`}
                            className="grid grid-cols-[42px_1fr_1.45fr_1.45fr_auto] gap-1 items-center bg-white/95 px-1 py-1 rounded border border-gray-300 shadow-2xs"
                          >
                            {/* Visa Type Label */}
                            <div className="text-center">
                              <span
                                className={`text-xs font-extrabold font-times ${
                                  vt.isRed ? 'text-red-600' : 'text-[#1E3A8A]'
                                }`}
                              >
                                {vt.name}
                                {entries.length > 1 && (
                                  <span className="text-[9px] block text-blue-600 font-bold font-times leading-none">
                                    #{entryIndex + 1}
                                  </span>
                                )}
                              </span>
                            </div>

                            {/* 1. ចំនួន (Quantity) */}
                            <div className="w-full">
                              <input
                                type="text"
                                data-stub-row={flatRowIndex}
                                data-stub-col={0}
                                id={`stub-input-${flatRowIndex}-0`}
                                placeholder="ចំនួន"
                                value={entry.quantity}
                                onChange={(e) =>
                                  handleInputChange(vt.id, entryIndex, 'quantity', e.target.value)
                                }
                                onFocus={(e) => e.target.select()}
                                onKeyDown={(e) => handleKeyDown(e, flatRowIndex, 0, totalRows)}
                                className="w-full h-7 bg-white border border-gray-400 rounded px-1 text-xs text-center font-bold font-times text-blue-900 focus:bg-yellow-50 focus:border-blue-500 outline-hidden shadow-2xs placeholder:font-normal placeholder:tracking-normal placeholder:font-siemreap placeholder:text-gray-400"
                              />
                            </div>

                            {/* 2. ចាប់ពីលេខ (Start Serial) */}
                            <div className="w-full">
                              <input
                                type="text"
                                data-stub-row={flatRowIndex}
                                data-stub-col={1}
                                id={`stub-input-${flatRowIndex}-1`}
                                placeholder="ពីលេខ"
                                value={entry.startSerial}
                                onChange={(e) =>
                                  handleInputChange(vt.id, entryIndex, 'startSerial', e.target.value)
                                }
                                onFocus={(e) => e.target.select()}
                                onKeyDown={(e) => handleKeyDown(e, flatRowIndex, 1, totalRows)}
                                className="w-full h-7 bg-white border border-gray-400 rounded px-1 text-xs text-center font-bold font-times text-gray-900 focus:bg-yellow-50 focus:border-blue-500 outline-hidden shadow-2xs tracking-tight placeholder:tracking-normal placeholder:font-normal placeholder:font-siemreap placeholder:text-gray-400"
                              />
                            </div>

                            {/* 3. ដល់លេខ (End Serial) */}
                            <div className="w-full">
                              <input
                                type="text"
                                data-stub-row={flatRowIndex}
                                data-stub-col={2}
                                id={`stub-input-${flatRowIndex}-2`}
                                placeholder="ដល់លេខ"
                                value={entry.endSerial}
                                onChange={(e) =>
                                  handleInputChange(vt.id, entryIndex, 'endSerial', e.target.value)
                                }
                                onFocus={(e) => e.target.select()}
                                onKeyDown={(e) => handleKeyDown(e, flatRowIndex, 2, totalRows)}
                                className="w-full h-7 bg-white border border-gray-400 rounded px-1 text-xs text-center font-bold font-times text-gray-900 focus:bg-yellow-50 focus:border-blue-500 outline-hidden shadow-2xs tracking-tight placeholder:tracking-normal placeholder:font-normal placeholder:font-siemreap placeholder:text-gray-400"
                              />
                            </div>

                            {/* 4. (+) Button and Trash Button (NO OLD CODE TEXTBOX) */}
                            <div className="flex items-center gap-0.5 shrink-0">
                              <button
                                type="button"
                                onClick={() => handleAddRow(vt.id, entryIndex)}
                                className="h-7 px-2 bg-[#16A34A] hover:bg-[#15803D] active:scale-95 text-white flex items-center justify-center rounded border border-green-800 cursor-pointer shadow-2xs transition"
                                title={`ចុច (+) ដើម្បីបន្ថែមជួរថ្មីសម្រាប់ប្រភេទ ${vt.name}`}
                              >
                                <Plus className="w-3.5 h-3.5 stroke-[3]" />
                              </button>
                              {entries.length > 1 && (
                                <button
                                  type="button"
                                  onClick={() => handleRemoveRow(vt.id, entryIndex)}
                                  className="h-7 px-1.5 bg-red-600 hover:bg-red-700 active:scale-95 text-white flex items-center justify-center rounded border border-red-800 cursor-pointer shadow-2xs transition"
                                  title="លុបជួរនេះចេញ"
                                >
                                  <Trash2 className="w-3 h-3" />
                                </button>
                              )}
                            </div>
                          </div>
                        );
                      })}
                    </div>
                  );
                });
              })()}
            </div>
          </div>
        </div>

        {/* RIGHT COLUMN: Real-Time Team Summary & History Database */}
        <div className="flex-1 w-full min-w-0 space-y-3">
          {/* 1. Daily Team Operations Matrix Table */}
          <div className="bg-white border-2 border-[#1E40AF] rounded shadow-sm overflow-hidden">
            {/* Table Header Bar */}
            <div className="bg-[#0369A1] text-white px-3 py-2 flex flex-wrap items-center justify-between gap-2 border-b border-sky-800">
              <div className="flex items-center gap-2">
                <FileSpreadsheet className="w-4 h-4 text-yellow-300" />
                <div className="flex flex-wrap items-baseline gap-1.5">
                  <span className="font-bold text-xs sm:text-sm">
                    តារាងស្ថិតិប្រមូលគល់សន្លឹក (Sticker) — ក្រុមដែលប្រគល់
                  </span>
                  <span className="text-xs text-sky-200 font-semibold inline-flex items-center gap-1">
                    <span className="font-times">[</span>
                    {dateMatrix.effectiveStart === dateMatrix.effectiveEnd ? (
                      <>
                        <span
                          className="font-siemreap"
                          style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}
                        >
                          ថ្ងៃទី
                        </span>
                        <span className="font-times font-semibold">{dateMatrix.effectiveStart}</span>
                      </>
                    ) : (
                      <>
                        <span
                          className="font-siemreap"
                          style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}
                        >
                          ពី
                        </span>
                        <span className="font-times font-semibold">{dateMatrix.effectiveStart}</span>
                        <span
                          className="font-siemreap"
                          style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}
                        >
                          ដល់
                        </span>
                        <span className="font-times font-semibold">{dateMatrix.effectiveEnd}</span>
                      </>
                    )}
                    <span className="font-times">]</span>
                  </span>
                </div>
              </div>

              <div className="flex items-center gap-2">
                <span className="text-xs text-sky-100 hidden sm:inline">
                  ក្រុមបានប្រគល់៖{' '}
                  <span className="font-bold text-yellow-300 font-times">{dateMatrix.rows.length}</span>{' '}
                  ក្រុម
                </span>
                <span className="text-xs bg-yellow-300 text-blue-950 px-2.5 py-0.5 rounded font-bold shadow-xs flex items-center gap-1">
                  <span className="text-[11px] text-blue-900 font-siemreap">សរុប៖</span>
                  <span className="font-times">{dateMatrix.grandTotal.toLocaleString()}</span>{' '}
                  <span
                    className="font-bold font-siemreap"
                    style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}
                  >
                    សន្លឹក
                  </span>
                </span>
              </div>
            </div>

            {/* Date Range Selection & Filter Toolbar */}
            <div className="bg-sky-50/90 border-b border-sky-200 px-3 py-2 flex flex-wrap items-center justify-between gap-2 text-xs">
              <div className="flex flex-wrap items-center gap-2">
                {/* Start Date */}
                <div className="flex items-center gap-1.5 bg-white border border-sky-300 rounded px-2 py-0.5 shadow-2xs">
                  <Calendar className="w-3.5 h-3.5 text-sky-700 shrink-0" />
                  <span
                    className="text-[11px] font-bold text-slate-700 whitespace-nowrap font-siemreap"
                    style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}
                  >
                    ចាប់ផ្តើម៖
                  </span>
                  <div className="w-28">
                    <CustomDatePicker
                      value={matrixStartDate}
                      onChange={(d) => setMatrixStartDate(d)}
                      className="py-0.5 text-xs font-semibold"
                    />
                  </div>
                </div>

                {/* End Date */}
                <div className="flex items-center gap-1.5 bg-white border border-sky-300 rounded px-2 py-0.5 shadow-2xs">
                  <Calendar className="w-3.5 h-3.5 text-sky-700 shrink-0" />
                  <span
                    className="text-[11px] font-bold text-slate-700 whitespace-nowrap font-siemreap"
                    style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}
                  >
                    ដល់ថ្ងៃ៖
                  </span>
                  <div className="w-28">
                    <CustomDatePicker
                      value={matrixEndDate}
                      onChange={(d) => setMatrixEndDate(d)}
                      className="py-0.5 text-xs font-semibold"
                    />
                  </div>
                </div>

                {/* Quick Presets */}
                <div className="flex items-center gap-1">
                  <button
                    type="button"
                    onClick={() => {
                      const today = new Date().toISOString().split('T')[0];
                      setMatrixStartDate(today);
                      setMatrixEndDate(today);
                      onShowToast('បានកំណត់កាលបរិច្ឆេទ «ថ្ងៃនេះ»', 'info');
                    }}
                    className={`px-2 py-1 rounded text-[11px] font-medium border transition cursor-pointer ${
                      matrixStartDate === matrixEndDate &&
                      matrixStartDate === new Date().toISOString().split('T')[0]
                        ? 'bg-sky-600 border-sky-700 text-white font-bold shadow-2xs'
                        : 'bg-white hover:bg-sky-100 border-sky-300 text-sky-900'
                    }`}
                    title="កំណត់កាលបរិច្ឆេទថ្ងៃនេះ"
                  >
                    ថ្ងៃនេះ
                  </button>

                  <button
                    type="button"
                    onClick={() => {
                      setMatrixStartDate(selectedDate);
                      setMatrixEndDate(selectedDate);
                      onShowToast(`បានកំណត់ស្មើនឹងថ្ងៃកត់ត្រា៖ ${selectedDate}`, 'info');
                    }}
                    className={`px-2 py-1 rounded text-[11px] font-medium border transition cursor-pointer ${
                      matrixStartDate === matrixEndDate && matrixStartDate === selectedDate
                        ? 'bg-sky-600 border-sky-700 text-white font-bold shadow-2xs'
                        : 'bg-white hover:bg-sky-100 border-sky-300 text-sky-900'
                    }`}
                    title="កំណត់ស្មើនឹងថ្ងៃកត់ត្រានៅក្នុងទម្រង់ខាងឆ្វេង"
                  >
                    តាមថ្ងៃកត់ត្រា
                  </button>

                  <button
                    type="button"
                    onClick={() => {
                      const now = new Date();
                      const y = now.getFullYear();
                      const m = String(now.getMonth() + 1).padStart(2, '0');
                      const firstDay = `${y}-${m}-01`;
                      const today = now.toISOString().split('T')[0];
                      setMatrixStartDate(firstDay);
                      setMatrixEndDate(today);
                      onShowToast('បានកំណត់ចន្លោះកាលបរិច្ឆេទ «ខែនេះ»', 'info');
                    }}
                    className="px-2 py-1 bg-white hover:bg-sky-100 border border-sky-300 text-sky-900 rounded text-[11px] font-medium transition cursor-pointer"
                    title="កំណត់ពីថ្ងៃទី ១ ដើមខែនេះរហូតដល់ថ្ងៃនេះ"
                  >
                    ខែនេះ
                  </button>
                </div>
              </div>

              {/* Matrix Table Excel Export */}
              <button
                type="button"
                onClick={handleExportMatrixExcel}
                className="px-2.5 py-1 bg-emerald-600 hover:bg-emerald-700 text-white rounded text-[11px] font-medium flex items-center gap-1 cursor-pointer transition shadow-2xs ml-auto"
                title="ទាញយកតារាងស្ថិតិនេះជា Excel"
              >
                <Download className="w-3.5 h-3.5" />
                <span>Excel តារាងស្ថិតិ</span>
              </button>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-xs text-center border-collapse table-fixed min-w-[960px]">
                <colgroup>
                  <col className="w-[36px] min-w-[36px]" />
                  <col className="w-[125px] min-w-[125px]" />
                  {VISA_TYPES.map((vt) => (
                    <col key={vt.id} className="w-[56px] min-w-[56px]" />
                  ))}
                  <col className="w-[68px] min-w-[68px]" />
                </colgroup>
                <thead>
                  {/* Column Headers */}
                  <tr className="bg-[#0F4C81] text-white text-xs font-bold">
                    <th className="p-1.5 border border-blue-900 w-[36px]">
                      ល.រ
                    </th>
                    <th className="p-1.5 border border-blue-900 text-left w-[125px] whitespace-nowrap px-2 truncate">
                      ក្រុមផ្តល់ទិដ្ឋាការ
                    </th>
                    {VISA_TYPES.map((vt) => (
                      <th
                        key={vt.id}
                        className="p-1.5 border border-blue-900 w-[56px] text-center font-times text-xs font-bold"
                      >
                        {vt.name}
                      </th>
                    ))}
                    <th className="p-1.5 border border-blue-900 bg-[#0B2545] font-bold w-[68px] text-center font-times text-xs">
                      TOTAL
                    </th>
                  </tr>

                  {/* Active Selected Team Row - Only display when it has data */}
                  {selectedTeam && (selectedTeamStats?.total || 0) > 0 && (
                    <tr className="bg-[#0284C7] text-white font-bold text-xs">
                      <td
                        colSpan={2}
                        className="p-1.5 border border-sky-700 text-left px-2"
                      >
                        <div className="flex items-center gap-1.5">
                          <CheckCircle2 className="w-3.5 h-3.5 text-yellow-300 shrink-0" />
                          <span className="text-xs tracking-wide font-bold">
                            {selectedTeam} (កំពុងជ្រើសរើស)
                          </span>
                        </div>
                      </td>
                      {VISA_TYPES.map((vt) => {
                        const count = selectedTeamStats?.counts[vt.id] || 0;
                        return (
                          <td
                            key={vt.id}
                            className="p-1 border border-sky-700 font-times font-bold text-center"
                          >
                            <span className="w-full max-w-[50px] mx-auto py-0.5 px-1 bg-black/25 rounded flex items-center justify-center text-xs font-times font-bold">
                              {count}
                            </span>
                          </td>
                        );
                      })}
                      <td className="p-1 border border-sky-700 font-times bg-[#0369A1] text-yellow-300 font-bold text-center">
                        <span className="w-full max-w-[58px] mx-auto py-0.5 px-1 bg-yellow-400 text-blue-950 font-bold rounded flex items-center justify-center text-xs font-times">
                          {selectedTeamStats?.total || 0}
                        </span>
                      </td>
                    </tr>
                  )}

                  {/* Overall Totals Row (Red) */}
                  <tr className="bg-[#DC2626] text-white font-bold text-xs">
                    <td
                      colSpan={2}
                      className="p-1.5 border border-red-800 text-left px-2"
                    >
                      <div className="flex items-center gap-2">
                        <span className="font-bold text-xs">
                          {dateMatrix.effectiveStart === dateMatrix.effectiveEnd
                            ? 'សរុបប្រមូលប្រចាំថ្ងៃ'
                            : 'សរុបប្រមូលក្នុងចន្លោះកាលបរិច្ឆេទ'}
                        </span>
                      </div>
                    </td>
                    {VISA_TYPES.map((vt) => (
                      <td
                        key={vt.id}
                        className="p-1 border border-red-800 font-times font-bold text-center"
                      >
                        <span className="w-full max-w-[50px] mx-auto py-0.5 px-1 bg-black/25 rounded flex items-center justify-center text-xs font-times font-bold">
                          {dateMatrix.totals[vt.id] || 0}
                        </span>
                      </td>
                    ))}
                    <td className="p-1 border border-red-800 font-times bg-[#991B1B] text-yellow-300 font-bold text-center">
                      <span className="w-full max-w-[58px] mx-auto py-0.5 px-1 bg-yellow-400 text-red-950 font-bold rounded flex items-center justify-center text-xs font-times">
                        {dateMatrix.grandTotal}
                      </span>
                    </td>
                  </tr>
                </thead>

                <tbody>
                  {dateMatrix.rows.length === 0 ? (
                    <tr>
                      <td
                        colSpan={VISA_TYPES.length + 3}
                        className="p-8 text-center text-gray-500 italic bg-white border border-gray-300"
                      >
                        {dateMatrix.effectiveStart === dateMatrix.effectiveEnd
                          ? `ពុំទាន់មានទិន្នន័យប្រមូលគល់សន្លឹក (Sticker) ក្នុងថ្ងៃ ${dateMatrix.effectiveStart} នៅឡើយទេ។ (ពេលមានទិន្នន័យប្រគល់ពីក្រុម នឹងបង្ហាញនៅទីនេះដោយស្វ័យប្រវត្តិ)`
                          : `ពុំទាន់មានទិន្នន័យប្រមូលគល់សន្លឹក (Sticker) ក្នុងចន្លោះថ្ងៃ ${dateMatrix.effectiveStart} ដល់ ${dateMatrix.effectiveEnd} នៅឡើយទេ។ (ពេលមានទិន្នន័យប្រគល់ពីក្រុម នឹងបង្ហាញនៅទីនេះដោយស្វ័យប្រវត្តិ)`}
                      </td>
                    </tr>
                  ) : (
                    dateMatrix.rows.map((row, idx) => (
                      <tr
                        key={row.teamName}
                        onClick={() => {
                          setSelectedTeam(row.teamName);
                          if (row.recordId) {
                            const found = effectiveRecords.find((r) => r.id === row.recordId);
                            if (found) loadRecordIntoForm(found);
                          } else {
                            const found =
                              effectiveRecords.find(
                                (r) =>
                                  normalizeTeamName(r.teamName) === normalizeTeamName(row.teamName) &&
                                  normalizeDateToISO(r.date) === normalizeDateToISO(selectedDate)
                              ) ||
                              effectiveRecords.find(
                                (r) => normalizeTeamName(r.teamName) === normalizeTeamName(row.teamName)
                              );
                            if (found) loadRecordIntoForm(found);
                          }
                        }}
                        className={`cursor-pointer transition hover:bg-yellow-100 ${
                          selectedTeam === row.teamName
                            ? 'bg-blue-100 font-bold text-blue-900'
                            : idx % 2 === 0
                            ? 'bg-white'
                            : 'bg-blue-50/20'
                        }`}
                      >
                        <td className="p-1.5 border border-gray-300 text-gray-700 font-semibold font-times">
                          {row.displayIndex || idx + 1}
                        </td>
                        <td
                          className="p-1.5 border border-gray-300 text-left font-medium text-slate-800 whitespace-nowrap truncate px-2"
                          title={row.teamName}
                        >
                          {row.teamName}
                        </td>
                        {VISA_TYPES.map((vt) => {
                          const count = row.counts[vt.id] || 0;
                          return (
                            <td
                              key={vt.id}
                              className={`p-1.5 border border-gray-300 font-times text-xs text-center ${
                                count > 0 ? 'font-bold text-blue-900 bg-blue-50' : 'text-gray-400'
                              }`}
                            >
                              {count}
                            </td>
                          );
                        })}
                        <td
                          className={`p-1.5 border border-gray-300 font-times font-bold text-xs text-center ${
                            row.total > 0 ? 'bg-blue-200 text-blue-950' : 'text-gray-400 bg-white'
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
      </div>

      {/* Record Details Modal Dialog */}
      {viewDetailRecord && (
        <div className="fixed inset-0 bg-black/50 backdrop-blur-xs flex items-center justify-center p-3 z-50 animate-in fade-in duration-150">
          <div className="bg-white rounded-xl shadow-2xl border border-gray-300 w-full max-w-2xl overflow-hidden flex flex-col max-h-[90vh]">
            <div className="bg-[#1E3A8A] text-white px-4 py-3 flex items-center justify-between">
              <div className="flex items-center gap-2">
                <BookmarkCheck className="w-5 h-5 text-yellow-300" />
                <h3 className="font-bold text-sm sm:text-base">
                  ព័ត៌មានលម្អិតនៃការប្រមូលគល់សន្លឹក — {viewDetailRecord.teamName}
                </h3>
              </div>
              <button
                type="button"
                onClick={() => setViewDetailRecord(null)}
                className="text-white/80 hover:text-white p-1 rounded hover:bg-white/10"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="p-4 overflow-y-auto space-y-3 text-xs">
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 bg-slate-50 p-3 rounded border border-slate-200">
                <div>
                  <span className="text-gray-500 block">កាលបរិច្ឆេទ</span>
                  <strong className="text-gray-900 font-times text-sm">
                    {viewDetailRecord.date}
                  </strong>
                </div>
                <div>
                  <span className="text-gray-500 block">ក្រុម</span>
                  <strong className="text-blue-900 text-sm">{viewDetailRecord.teamName}</strong>
                </div>
                <div>
                  <span className="text-gray-500 block">សរុបសន្លឹក</span>
                  <strong className="text-emerald-700 text-sm">
                    <span className="font-times">{viewDetailRecord.totalSheets.toLocaleString()}</span>{' '}
                    <span
                      className="font-siemreap"
                      style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}
                    >
                      សន្លឹក
                    </span>
                  </strong>
                </div>
                <div>
                  <span className="text-gray-500 block">អ្នកកត់ត្រា</span>
                  <strong className="text-gray-800">{viewDetailRecord.collectedBy}</strong>
                </div>
              </div>

              {viewDetailRecord.remarks && (
                <div className="p-2 bg-amber-50 border border-amber-200 rounded text-amber-900">
                  <strong>ចំណាំ៖</strong> {viewDetailRecord.remarks}
                </div>
              )}

              {/* Table of items */}
              <div className="border border-gray-300 rounded overflow-hidden">
                <table className="w-full text-xs">
                  <thead className="bg-[#0F4C81] text-white font-bold">
                    <tr>
                      <th className="p-2 text-center w-16">ប្រភេទ</th>
                      <th className="p-2 text-center w-20">ចំនួន</th>
                      <th className="p-2 text-center">ចាប់ពីលេខ</th>
                      <th className="p-2 text-center">ដល់លេខ</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-200 text-center">
                    {VISA_TYPES.map((vt) => {
                      const v = viewDetailRecord.values?.[vt.id];
                      if (!v) return null;
                      const entries = Array.isArray(v.entries) ? v.entries : [v as any];
                      const valid = entries.filter((e) => e.quantity && parseInt(e.quantity, 10) > 0);
                      if (valid.length === 0) return null;

                      return valid.map((entry, eIdx) => (
                        <tr key={`${vt.id}-${eIdx}`} className="hover:bg-blue-50">
                          <td className="p-2 font-times font-bold text-blue-900">
                            {vt.name}
                            {valid.length > 1 && (
                              <span className="text-[10px] text-gray-500 ml-1">#{eIdx + 1}</span>
                            )}
                          </td>
                          <td className="p-2 font-times font-bold text-emerald-700">
                            {entry.quantity}
                          </td>
                          <td className="p-2 font-times">{entry.startSerial || '-'}</td>
                          <td className="p-2 font-times">{entry.endSerial || '-'}</td>
                        </tr>
                      ));
                    })}
                  </tbody>
                </table>
              </div>
            </div>

            <div className="bg-slate-100 px-4 py-2.5 border-t border-slate-200 flex justify-end">
              <button
                type="button"
                onClick={() => setViewDetailRecord(null)}
                className="px-4 py-1.5 bg-[#1E3A8A] hover:bg-blue-900 text-white rounded text-xs font-bold"
              >
                បិទ (Close)
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Print Preview Modal Dialog */}
      {isPrintModalOpen && (
        <div className="fixed inset-0 bg-black/60 backdrop-blur-xs flex items-center justify-center p-2 sm:p-4 z-50 animate-in fade-in duration-200">
          <div className="bg-white rounded-xl shadow-2xl border border-gray-300 w-full max-w-4xl max-h-[92vh] flex flex-col overflow-hidden">
            <div className="bg-[#1E3A8A] text-white px-4 py-3 flex items-center justify-between print:hidden">
              <div className="flex items-center gap-2">
                <Printer className="w-5 h-5 text-yellow-300" />
                <h3 className="font-bold text-sm sm:text-base">
                  ទម្រង់បោះពុម្ព — របាយការណ៍ប្រមូលគល់សន្លឹកទិដ្ឋាការស្អិត
                </h3>
              </div>
              <div className="flex items-center gap-2">
                <button
                  type="button"
                  onClick={() => setIsPrintModalOpen(false)}
                  className="text-white/80 hover:text-white p-1 rounded hover:bg-white/10"
                >
                  <X className="w-5 h-5" />
                </button>
              </div>
            </div>

            {/* Printable Document Sheet */}
            <div id="visa-stub-printable-document" className="p-6 sm:p-8 overflow-y-auto bg-white text-black font-sans print:p-0">
              <div className="text-center space-y-1 mb-6">
                <p className="font-moul text-sm sm:text-base">ព្រះរាជាណាចក្រកម្ពុជា</p>
                <p className="font-moul text-xs sm:text-sm">ជាតិ សាសនា ព្រះមហាក្សត្រ</p>
                <div className="h-0.5 w-24 bg-black mx-auto mt-1" />
              </div>

              <div className="text-center mb-6">
                <h2 className="font-moul text-sm sm:text-base text-blue-950">
                  តារាងស្ថិតិប្រមូលគល់សន្លឹកទិដ្ឋាការស្អិត
                </h2>
                <p className="text-xs text-gray-700 mt-1">
                  កាលបរិច្ឆេទគិតត្រឹមថ្ងៃទី {selectedDate} — សរុបចំនួន{' '}
                  <strong className="font-times">{overallStats.totalAllSheets.toLocaleString()}</strong>{' '}
                  <span
                    className="font-siemreap"
                    style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}
                  >
                    សន្លឹក
                  </span>
                </p>
              </div>

              <table className="w-full text-xs border border-black border-collapse text-center">
                <thead>
                  <tr className="bg-gray-100 font-bold">
                    <th className="border border-black p-1.5 w-10">ល.រ</th>
                    <th className="border border-black p-1.5 w-24">កាលបរិច្ឆេទ</th>
                    <th className="border border-black p-1.5 text-left px-2">ក្រុមផ្តល់ទិដ្ឋាការ</th>
                    <th className="border border-black p-1.5 w-24">សរុប (សន្លឹក)</th>
                    <th className="border border-black p-1.5 text-left px-2">លម្អិតជួរលេខកូដសន្លឹក</th>
                    <th className="border border-black p-1.5 w-28">អ្នកកត់ត្រា</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredRecords.map((r, idx) => {
                    const descList: string[] = [];
                    VISA_TYPES.forEach((vt) => {
                      const v = r.values?.[vt.id];
                      if (v) {
                        if (Array.isArray(v.entries)) {
                          v.entries.forEach((e) => {
                            if (e.quantity && parseInt(e.quantity, 10) > 0) {
                              descList.push(
                                `${vt.name}:${e.quantity}${e.startSerial ? `(${e.startSerial}-${e.endSerial})` : ''}`
                              );
                            }
                          });
                        } else if (v.quantity && parseInt(v.quantity, 10) > 0) {
                          descList.push(
                            `${vt.name}:${v.quantity}${v.startSerial ? `(${v.startSerial}-${v.endSerial})` : ''}`
                          );
                        }
                      }
                    });

                    return (
                      <tr key={r.id}>
                        <td className="border border-black p-1.5 font-times">{idx + 1}</td>
                        <td className="border border-black p-1.5 font-times">{r.date}</td>
                        <td className="border border-black p-1.5 text-left px-2 font-bold">
                          {r.teamName}
                        </td>
                        <td className="border border-black p-1.5 font-times font-bold">
                          {r.totalSheets}
                        </td>
                        <td className="border border-black p-1.5 text-left px-2 font-times text-[11px]">
                          {descList.join(', ')}
                        </td>
                        <td className="border border-black p-1.5 text-center">{r.collectedBy}</td>
                      </tr>
                    );
                  })}
                  <tr className="bg-gray-100 font-bold">
                    <td colSpan={3} className="border border-black p-1.5 text-center">
                      សរុបរួម
                    </td>
                    <td className="border border-black p-1.5 font-times text-sm">
                      {overallStats.totalAllSheets.toLocaleString()}
                    </td>
                    <td colSpan={2} className="border border-black p-1.5" />
                  </tr>
                </tbody>
              </table>

              <div className="mt-8 flex justify-between text-xs px-6">
                <div className="text-center">
                  <p className="font-bold">អ្នកផ្ទៀងផ្ទាត់</p>
                  <p className="mt-12 text-gray-400 font-italic">(ហត្ថលេខា & ឈ្មោះ)</p>
                </div>
                <div className="text-center">
                  <p className="font-bold">រាជធានីភ្នំពេញ, ថ្ងៃទី....... ខែ....... ឆ្នាំ២០២...</p>
                  <p className="font-bold mt-1">អ្នកធ្វើរបាយការណ៍</p>
                  <p className="mt-12 font-bold">{userName || 'មន្ត្រីការិយាល័យ'}</p>
                </div>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* Delete Confirmation Modal */}
      {deleteModalOpen && deleteConfirmTarget && (
        <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl border border-red-200 max-w-md w-full overflow-hidden animate-in fade-in zoom-in-95 duration-150">
            <div className="bg-red-600 text-white px-4 py-3 flex items-center justify-between">
              <div className="flex items-center gap-2 font-bold text-sm">
                <Trash2 className="w-4 h-4" />
                <span>បញ្ជាក់ការលុបទិន្នន័យប្រមូលគល់សន្លឹក</span>
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
              <p className="font-medium">
                តើអ្នកពិតជាចង់លុបទិន្នន័យប្រមូលគល់សន្លឹកនេះចេញពីប្រព័ន្ធ និងទិន្នន័យស្តុកឬទេ?
              </p>
              <div className="bg-red-50 border border-red-200 rounded-lg p-3 space-y-1.5 text-xs">
                <div className="flex justify-between">
                  <span className="text-gray-500">កាលបរិច្ឆេទ៖</span>
                  <span className="font-bold text-slate-800 font-times">{deleteConfirmTarget.date}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-gray-500">ក្រុមផ្តល់ទិដ្ឋាការ៖</span>
                  <span className="font-bold text-red-700">{deleteConfirmTarget.teamName}</span>
                </div>
                <div className="flex justify-between border-t border-red-200 pt-1">
                  <span className="text-gray-600">ចំនួនសរុប៖</span>
                  <span className="font-bold text-red-800">
                    <span className="font-times">{deleteConfirmTarget.totalSheets.toLocaleString()}</span>{' '}
                    <span
                      className="font-siemreap"
                      style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}
                    >
                      សន្លឹក
                    </span>
                  </span>
                </div>
                {deleteConfirmTarget.remarks && (
                  <div className="flex justify-between text-[11px] text-gray-600">
                    <span>ចំណាំ៖</span>
                    <span className="truncate max-w-[200px]">{deleteConfirmTarget.remarks}</span>
                  </div>
                )}
              </div>
              <p className="text-[11px] text-gray-500 italic">
                * ការលុបនេះនឹងដកទិន្នន័យទាំងស្រុងពីបញ្ជីប្រមូលគល់សន្លឹក និងពី «ទិន្នន័យសន្លឹកទិដ្ឋាការ» (ស្តុក) ដោយស្វ័យប្រវត្តិ។
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
                className="px-4 py-1.5 bg-red-600 hover:bg-red-700 active:scale-95 text-white rounded-lg text-xs font-bold shadow-xs cursor-pointer transition flex items-center gap-1.5"
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>លុបចេញ (Confirm Delete)</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Update Confirmation Modal */}
      {updateModalOpen && (
        <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl border border-fuchsia-200 max-w-md w-full overflow-hidden animate-in fade-in zoom-in-95 duration-150">
            <div className="bg-[#A21CAF] text-white px-4 py-3 flex items-center justify-between">
              <div className="flex items-center gap-2 font-bold text-sm">
                <Edit3 className="w-4 h-4" />
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
              <p className="font-medium">
                តើអ្នកពិតជាចង់កែសម្រួលទិន្នន័យប្រមូលគល់សន្លឹកនេះ និងធ្វើបច្ចុប្បន្នភាពក្នុងស្តុកឬទេ?
              </p>
              <div className="bg-fuchsia-50 border border-fuchsia-200 rounded-lg p-3 space-y-1.5 text-xs">
                <div className="flex justify-between">
                  <span className="text-gray-500">កាលបរិច្ឆេទ៖</span>
                  <span className="font-bold text-slate-800 font-times">{selectedDate}</span>
                </div>
                <div className="flex justify-between">
                  <span className="text-gray-500">ក្រុមផ្តល់ទិដ្ឋាការ៖</span>
                  <span className="font-bold text-purple-800">{selectedTeam}</span>
                </div>
                <div className="flex justify-between border-t border-fuchsia-200 pt-1">
                  <span className="text-gray-600 font-medium">ចំនួនសរុបដែលត្រូវកែប្រែ៖</span>
                  <span className="font-extrabold text-fuchsia-700">
                    <span className="font-times">{formTotalSheets.toLocaleString()}</span>{' '}
                    <span
                      className="font-siemreap"
                      style={{ fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif" }}
                    >
                      សន្លឹក
                    </span>
                  </span>
                </div>
              </div>
              <p className="text-[11px] text-gray-500 italic">
                * ការកែសម្រួលនេះនឹងធ្វើបច្ចុប្បន្នភាពក្នុងបញ្ជីប្រមូលគល់សន្លឹក និងធ្វើសមកាលកម្មជាមួយ «ទិន្នន័យសន្លឹកទិដ្ឋាការ» (ស្តុក) ភ្លាមៗ។
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
                className="px-4 py-1.5 bg-[#C026D3] hover:bg-[#A21CAF] active:scale-95 text-white rounded-lg text-xs font-bold shadow-xs cursor-pointer transition flex items-center gap-1.5"
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
