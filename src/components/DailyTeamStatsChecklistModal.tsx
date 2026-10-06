import React, { useState, useMemo, useEffect } from 'react';
import { apiService } from '../services/apiService';
import {
  X,
  Calendar,
  Layers,
  Search,
  CheckCircle2,
  AlertTriangle,
  RotateCcw,
  CheckSquare,
  Square,
  Filter,
  Trash2,
  Info,
} from 'lucide-react';
import { StockRecord } from '../types';
import { normalizeTeamName, normalizeDateToISO } from '../utils/teamNormalization';
import { toKhmerNum } from '../utils/khmerCalendar';

export interface DailyUsageAuditItem {
  id: string;
  date: string;
  dayNum: number;
  teamName: string;
  categoryType: 'Sticker' | 'cEA';
  operationTitle: string;
  operationType: string;
  source: 'daily_form' | 'stock_ledger';
  values: Record<string, { quantity: number; startSerial?: string; endSerial?: string }>;
  totalSheets: number;
  isDuplicateOrOvercounted: boolean;
  overcountReason?: string;
  rawRecord?: any;
}

export interface DailyTeamStatsChecklistModalProps {
  isOpen: boolean;
  onClose: () => void;
  selectedTeam: string;
  selectedYear: number;
  selectedMonth: number;
  daysInMonth: number;
  stockRecords?: StockRecord[];
  excludedIds: string[];
  onApplyExcludedIds: (newExcludedIds: string[]) => void;
  allTeams?: string[];
  categoryFilterProp?: 'Sticker' | 'cEA' | 'all';
}

export const DailyTeamStatsChecklistModal: React.FC<DailyTeamStatsChecklistModalProps> = ({
  isOpen,
  onClose,
  selectedTeam,
  selectedYear,
  selectedMonth,
  daysInMonth,
  stockRecords = [],
  excludedIds: initialExcludedIds,
  onApplyExcludedIds,
  allTeams = [],
  categoryFilterProp,
}) => {
  if (!isOpen) return null;

  const isAllTeams = selectedTeam === 'សរុបបណ្តាក្រុម' || !selectedTeam;
  const normSelectedTeam = normalizeTeamName(selectedTeam);

  // Local state for excluded IDs while user is checking/unchecking in modal
  const [localExcludedIds, setLocalExcludedIds] = useState<string[]>(() => [...initialExcludedIds]);

  // Filters
  const [searchQuery, setSearchQuery] = useState<string>('');
  const [dayFilter, setDayFilter] = useState<'all' | number>('all');
  const [categoryFilter, setCategoryFilter] = useState<'Sticker' | 'cEA' | 'all'>(
    () => categoryFilterProp || 'Sticker'
  );

  useEffect(() => {
    if (categoryFilterProp) {
      setCategoryFilter(categoryFilterProp);
    }
  }, [categoryFilterProp, isOpen]);
  const [showOnlyOvercounted, setShowOnlyOvercounted] = useState<boolean>(false);
  const [activeTeamFilter, setActiveTeamFilter] = useState<string>(selectedTeam);

  // Collect all raw operations from app_daily_team_operations_v5 & stockRecords
  const allAuditItems = useMemo<DailyUsageAuditItem[]>(() => {
    const items: DailyUsageAuditItem[] = [];
    const occurrencesMap = new Map<string, number>();

    const normFilterTeam = normalizeTeamName(activeTeamFilter);
    const filterIsAllTeams = activeTeamFilter === 'សរុបបណ្តាក្រុម' || !activeTeamFilter;

    const isMatch = (rawTeam?: string) => {
      if (filterIsAllTeams) return true;
      if (!rawTeam || !normFilterTeam) return false;
      const recNorm = normalizeTeamName(rawTeam);
      return recNorm === normFilterTeam;
    };

    // 1. Load app_daily_team_operations_v5
    try {
      const savedOps = localStorage.getItem('app_daily_team_operations_v5');
      if (savedOps) {
        const records = JSON.parse(savedOps);
        if (Array.isArray(records)) {
          records.forEach((rec: any) => {
            if (!rec || !rec.date) return;
            if (!isMatch(rec.teamName)) return;

            const isoDate = normalizeDateToISO(rec.date);
            const parts = isoDate.split('-');
            if (parts.length < 3) return;

            const y = parseInt(parts[0], 10);
            const m = parseInt(parts[1], 10);
            const d = parseInt(parts[2], 10);

            // Month and year check
            if (y !== selectedYear || m !== selectedMonth) return;

            const cat: 'Sticker' | 'cEA' = rec.categoryType === 'cEA' ? 'cEA' : 'Sticker';

            // Extract values
            const valObj: Record<string, { quantity: number; startSerial?: string; endSerial?: string }> = {};
            let sumSheets = 0;

            if (rec.values && typeof rec.values === 'object') {
              Object.entries(rec.values).forEach(([vt, valData]: [string, any]) => {
                let q = 0;
                if (typeof valData?.quantity === 'number') q = valData.quantity;
                else if (typeof valData?.quantity === 'string') q = parseInt(valData.quantity, 10) || 0;

                if (q > 0) {
                  valObj[vt] = {
                    quantity: q,
                    startSerial: valData?.startSerial || '',
                    endSerial: valData?.endSerial || '',
                  };
                  sumSheets += q;
                }
              });
            }

            if (sumSheets === 0 && rec.totalSheets) {
              sumSheets = Number(rec.totalSheets) || 0;
            }

            // Tracking duplicate keys
            const key = `daily_form_${cat}_${isoDate}_${normalizeTeamName(rec.teamName)}`;
            const count = (occurrencesMap.get(key) || 0) + 1;
            occurrencesMap.set(key, count);

            items.push({
              id: rec.id || `dtr-${cat}-${isoDate}-${rec.teamName}`,
              date: isoDate,
              dayNum: d,
              teamName: rec.teamName,
              categoryType: cat,
              operationTitle: 'កត់ត្រាប្រចាំថ្ងៃ (Daily Form)',
              operationType: 'useTeam',
              source: 'daily_form',
              values: valObj,
              totalSheets: sumSheets,
              isDuplicateOrOvercounted: count > 1,
              overcountReason: count > 1 ? 'មានទិន្នន័យកត់ត្រាស្ទួនលើសពី១ដងក្នុងថ្ងៃតែមួយ' : undefined,
              rawRecord: rec,
            });
          });
        }
      }
    } catch (e) {
      console.warn('Error reading app_daily_team_operations_v5 in checklist', e);
    }

    // 2. Load stockRecords (props + localStorage)
    const combinedStock: StockRecord[] = [...stockRecords];
    try {
      const savedStock = localStorage.getItem('app_stock_records');
      if (savedStock) {
        const parsed = JSON.parse(savedStock);
        if (Array.isArray(parsed)) {
          parsed.forEach((r: StockRecord) => {
            if (!combinedStock.some((ex) => ex.id === r.id)) {
              combinedStock.push(r);
            }
          });
        }
      }
    } catch (e) {
      console.warn('Error reading app_stock_records in checklist', e);
    }

    combinedStock.forEach((sr) => {
      if (!sr || !sr.date) return;
      if (
        sr.id &&
        (sr.id.startsWith('stk-use-shv-') ||
          sr.id.startsWith('stk-use-phnomden-') ||
          sr.id.startsWith('stk-use-tpp-') ||
          sr.id.startsWith('stk-use-sample-') ||
          sr.id === 'stk-bavet-t' ||
          sr.id === 'stk-bavet-e')
      ) {
        return;
      }
      const op = (sr.operationType || '').toLowerCase();
      const isUse = op === 'useteam' || op.includes('use') || op.includes('ប្រើប្រាស់');
      if (!isUse) return;

      const rawTeam = sr.visaTeamRobokName || (sr as any).teamName || '';
      if (!isMatch(rawTeam)) return;

      const isoDate = normalizeDateToISO(sr.date);
      const parts = isoDate.split('-');
      if (parts.length < 3) return;

      const y = parseInt(parts[0], 10);
      const m = parseInt(parts[1], 10);
      const d = parseInt(parts[2], 10);

      if (y !== selectedYear || m !== selectedMonth) return;

      const cat: 'Sticker' | 'cEA' =
        sr.sourceFrom === 'cEA' ||
        sr.remarks?.includes('cEA') ||
        (sr.visaType && sr.visaType.includes('cEA'))
          ? 'cEA'
          : 'Sticker';

      const vt = (sr.visaType || '').toUpperCase().trim();
      const sheets = sr.totalSheets !== undefined && sr.totalSheets !== null
        ? sr.totalSheets
        : (sr.quantityBundles || 0);

      // Check if this stock record was generated automatically by DailyTeamVisaOperations
      // (e.g. id starts with stock-dtr-) and mirrors an already loaded daily_form entry
      const isMirroredFromDailyForm = sr.id && sr.id.startsWith('stock-dtr-');

      const key = `stock_ledger_${cat}_${isoDate}_${normalizeTeamName(rawTeam)}_${vt}`;
      const count = (occurrencesMap.get(key) || 0) + 1;
      occurrencesMap.set(key, count);

      const hasDailyFormSameDay = items.some(
        (it) => it.source === 'daily_form' && it.date === isoDate && normalizeTeamName(it.teamName) === normalizeTeamName(rawTeam) && it.categoryType === cat
      );

      const isDup = isMirroredFromDailyForm || count > 1 || (hasDailyFormSameDay && isMirroredFromDailyForm);
      let reason: string | undefined = undefined;
      if (isMirroredFromDailyForm) {
        reason = 'ទិន្នន័យចម្លងស្ទួនពីតារាងប្រតិបត្តិការប្រចាំថ្ងៃ (Auto-mirrored stock entry)';
      } else if (count > 1) {
        reason = 'មានទិន្នន័យស្ទួនក្នុងសៀវភៅស្តុក';
      }

      items.push({
        id: sr.id,
        date: isoDate,
        dayNum: d,
        teamName: rawTeam,
        categoryType: cat,
        operationTitle: 'សៀវភៅស្តុក (Stock Ledger)',
        operationType: sr.operationType || 'useTeam',
        source: 'stock_ledger',
        values: {
          [vt]: {
            quantity: sheets,
            startSerial: sr.startSerial || '',
            endSerial: sr.endSerial || '',
          },
        },
        totalSheets: sheets,
        isDuplicateOrOvercounted: isDup,
        overcountReason: reason,
        rawRecord: sr,
      });
    });

    // Sort by day descending (or date ascending)
    return items.sort((a, b) => a.date.localeCompare(b.date) || b.totalSheets - a.totalSheets);
  }, [activeTeamFilter, selectedYear, selectedMonth, stockRecords]);

  // Filter items based on active controls
  const filteredItems = useMemo(() => {
    return allAuditItems.filter((item) => {
      // Day filter
      if (dayFilter !== 'all' && item.dayNum !== dayFilter) return false;

      // Category filter
      if (categoryFilter !== 'all' && item.categoryType !== categoryFilter) return false;

      // Overcounted filter
      if (showOnlyOvercounted && !item.isDuplicateOrOvercounted) return false;

      // Search query
      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase().trim();
        const inDate = item.date.toLowerCase().includes(q);
        const inTeam = item.teamName.toLowerCase().includes(q);
        const inTitle = item.operationTitle.toLowerCase().includes(q);
        const inValues = Object.entries(item.values).some(
          ([vt, val]: [string, any]) =>
            vt.toLowerCase().includes(q) ||
            val?.startSerial?.toLowerCase().includes(q) ||
            val?.endSerial?.toLowerCase().includes(q)
        );
        if (!inDate && !inTeam && !inTitle && !inValues) return false;
      }

      return true;
    });
  }, [allAuditItems, dayFilter, categoryFilter, showOnlyOvercounted, searchQuery]);

  // Statistics calculation
  const stats = useMemo(() => {
    let totalItems = allAuditItems.length;
    let totalSheets = 0;
    let includedCount = 0;
    let includedSheets = 0;
    let excludedCount = 0;
    let excludedSheets = 0;
    let overcountedCount = 0;
    let overcountedSheets = 0;

    allAuditItems.forEach((item) => {
      totalSheets += item.totalSheets;
      const isExcluded = localExcludedIds.includes(item.id);

      if (isExcluded) {
        excludedCount++;
        excludedSheets += item.totalSheets;
      } else {
        includedCount++;
        includedSheets += item.totalSheets;
      }

      if (item.isDuplicateOrOvercounted) {
        overcountedCount++;
        overcountedSheets += item.totalSheets;
      }
    });

    return {
      totalItems,
      totalSheets,
      includedCount,
      includedSheets,
      excludedCount,
      excludedSheets,
      overcountedCount,
      overcountedSheets,
    };
  }, [allAuditItems, localExcludedIds]);

  // Toggle single item
  const handleToggleItem = (id: string) => {
    setLocalExcludedIds((prev) =>
      prev.includes(id) ? prev.filter((x) => x !== id) : [...prev, id]
    );
  };

  // Select all (include all filtered items)
  const handleIncludeAllFiltered = () => {
    const filteredIds = new Set(filteredItems.map((it) => it.id));
    setLocalExcludedIds((prev) => prev.filter((id) => !filteredIds.has(id)));
  };

  // Exclude all filtered items
  const handleExcludeAllFiltered = () => {
    const newExcluded = new Set(localExcludedIds);
    filteredItems.forEach((it) => newExcluded.add(it.id));
    setLocalExcludedIds(Array.from(newExcluded));
  };

  // Deselect all duplicate/overcounted items automatically!
  const handleExcludeAllDuplicates = () => {
    const newExcluded = new Set(localExcludedIds);
    allAuditItems.forEach((it) => {
      if (it.isDuplicateOrOvercounted) {
        newExcluded.add(it.id);
      }
    });
    setLocalExcludedIds(Array.from(newExcluded));
  };

  // Reset to initial
  const handleReset = () => {
    setLocalExcludedIds([...initialExcludedIds]);
  };

  // Apply changes to parent
  const handleApply = () => {
    onApplyExcludedIds(localExcludedIds);
    onClose();
  };

  // Permanently delete a raw duplicate record if from app_daily_team_operations_v5
  const handleDeleteOperation = (item: DailyUsageAuditItem) => {
    if (!window.confirm(`តើអ្នកពិតជាចង់លុបទិន្នន័យនេះជាអចិន្ត្រៃយ៍មែនទេ?\nថ្ងៃទី: ${item.date}\nក្រុម: ${item.teamName}\nសរុប: ${item.totalSheets} សន្លឹក`)) {
      return;
    }

    try {
      if (item.source === 'daily_form') {
        const saved = localStorage.getItem('app_daily_team_operations_v5');
        if (saved) {
          const records = JSON.parse(saved);
          if (Array.isArray(records)) {
            const filtered = records.filter((r: any) => r.id !== item.id);
            localStorage.setItem('app_daily_team_operations_v5', JSON.stringify(filtered));
          }
        }
        apiService.deleteDailyTeamOperation(item.id).catch(console.error);
      } else if (item.source === 'stock_ledger') {
        const saved = localStorage.getItem('app_stock_records');
        if (saved) {
          const records = JSON.parse(saved);
          if (Array.isArray(records)) {
            const filtered = records.filter((r: any) => r.id !== item.id);
            localStorage.setItem('app_stock_records', JSON.stringify(filtered));
          }
        }
        apiService.deleteStockRecord(item.id).catch(console.error);
      }
      // Also remove from excludedIds if present
      setLocalExcludedIds((prev) => prev.filter((id) => id !== item.id));
      alert('បានលុបទិន្នន័យដោយជោគជ័យ! សូម Refresh ទំព័រ (F5) ដើម្បីអាប់ដេតទិន្នន័យថ្មី។');
      window.dispatchEvent(new Event('storage'));
    } catch (e) {
      console.error('Delete error', e);
    }
  };

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center p-2 sm:p-4 bg-black/60 backdrop-blur-xs animate-in fade-in duration-200">
      <div className="bg-white rounded-xl shadow-2xl border border-slate-300 w-full max-w-6xl max-h-[92vh] flex flex-col overflow-hidden font-sans">
        {/* Modal Header */}
        <div className="bg-[#002060] text-white px-4 py-3 flex items-center justify-between border-b border-blue-900">
          <div className="flex items-center gap-2.5">
            <div className="p-2 bg-amber-400 text-slate-900 rounded-lg shadow-sm">
              <CheckSquare className="w-5 h-5" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h2 className="text-base sm:text-lg font-bold">
                  ផ្ទៀងផ្ទាត់បញ្ជីទិន្នន័យ & ស្វែងរកទិន្នន័យរាប់លើស (Checklist Audit)
                </h2>
                <span className="px-2 py-0.5 rounded-full text-[11px] font-semibold bg-blue-800 text-blue-100 border border-blue-600">
                  ខែ {selectedMonth} ឆ្នាំ {selectedYear}
                </span>
              </div>
              <p className="text-xs text-blue-200 mt-0.5">
                ពិនិត្យគ្រប់ទិន្នន័យប្រើប្រាស់ ធីកដើម្បីរាប់ ឬដោះធីកដើម្បីកុំឱ្យរាប់លើស/ស្ទួនលើតារាងស្ថិតិប្រចាំថ្ងៃ
              </p>
            </div>
          </div>
          <button
            type="button"
            onClick={onClose}
            className="text-white/80 hover:text-white p-1.5 rounded-lg hover:bg-white/10 transition cursor-pointer"
            title="បិទ (Close)"
          >
            <X className="w-6 h-6" />
          </button>
        </div>

        {/* Summary Metric Cards */}
        <div className="bg-slate-50 border-b border-slate-200 p-3 grid grid-cols-2 sm:grid-cols-4 gap-2 text-xs">
          <div className="bg-white p-2.5 rounded-lg border border-slate-200 shadow-2xs">
            <div className="text-slate-500 font-medium">សរុបទិន្នន័យរកឃើញ</div>
            <div className="text-base font-bold text-slate-800 mt-0.5">
              {stats.totalItems} <span className="text-xs font-normal text-slate-500">ប្រតិបត្តិការ</span>
            </div>
            <div className="text-[11px] text-slate-500 font-medium mt-0.5">
              សរុប {stats.totalSheets.toLocaleString()} សន្លឹក
            </div>
          </div>

          <div className="bg-emerald-50/60 p-2.5 rounded-lg border border-emerald-200 shadow-2xs">
            <div className="text-emerald-700 font-semibold flex items-center justify-between">
              <span>✓ បានរាប់បញ្ចូល</span>
              <span className="text-[10px] bg-emerald-200 text-emerald-900 px-1.5 py-0.2 rounded-full font-bold">
                សកម្ម
              </span>
            </div>
            <div className="text-base font-bold text-emerald-800 mt-0.5">
              {stats.includedCount} <span className="text-xs font-normal text-emerald-600">ប្រតិបត្តិការ</span>
            </div>
            <div className="text-[11px] text-emerald-700 font-bold mt-0.5">
              {stats.includedSheets.toLocaleString()} សន្លឹក
            </div>
          </div>

          <div className="bg-amber-50/60 p-2.5 rounded-lg border border-amber-200 shadow-2xs">
            <div className="text-amber-800 font-semibold flex items-center justify-between">
              <span>✕ បានដកចេញ (មិនរាប់)</span>
              {stats.excludedCount > 0 && (
                <span className="text-[10px] bg-amber-200 text-amber-900 px-1.5 py-0.2 rounded-full font-bold">
                  {stats.excludedCount}
                </span>
              )}
            </div>
            <div className="text-base font-bold text-amber-800 mt-0.5">
              {stats.excludedCount} <span className="text-xs font-normal text-amber-600">ប្រតិបត្តិការ</span>
            </div>
            <div className="text-[11px] text-amber-700 font-bold mt-0.5">
              {stats.excludedSheets.toLocaleString()} សន្លឹក
            </div>
          </div>

          <div className="bg-red-50/60 p-2.5 rounded-lg border border-red-200 shadow-2xs">
            <div className="text-red-700 font-semibold flex items-center justify-between">
              <span>⚠️ ស្ទួន / រាប់លើស</span>
              {stats.overcountedCount > 0 && (
                <span className="text-[10px] bg-red-200 text-red-900 px-1.5 py-0.2 rounded-full font-bold animate-pulse">
                  រកឃើញ
                </span>
              )}
            </div>
            <div className="text-base font-bold text-red-800 mt-0.5">
              {stats.overcountedCount} <span className="text-xs font-normal text-red-600">ប្រតិបត្តិការ</span>
            </div>
            <div className="text-[11px] text-red-600 font-medium mt-0.5">
              ស្មើនឹង {stats.overcountedSheets.toLocaleString()} សន្លឹក
            </div>
          </div>
        </div>

        {/* Filter and Search Bar */}
        <div className="bg-white border-b border-slate-200 p-3 flex flex-wrap items-center justify-between gap-2.5 text-xs">
          <div className="flex flex-wrap items-center gap-2 flex-1">
            {/* Search Box */}
            <div className="relative min-w-[200px] max-w-xs">
              <Search className="w-3.5 h-3.5 absolute left-2.5 top-2.5 text-slate-400" />
              <input
                type="text"
                placeholder="ស្វែងរកថ្ងៃ, ក្រុម, ស៊េរី, ប្រភេទ..."
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                className="w-full pl-8 pr-3 py-1.5 bg-slate-50 border border-slate-300 rounded-md focus:bg-white focus:border-blue-500 focus:outline-hidden text-xs"
              />
            </div>

            {/* Team Filter (if Secondary or multiple teams) */}
            {allTeams && allTeams.length > 0 && (
              <div className="flex items-center gap-1 bg-slate-50 px-2 py-1 border border-slate-300 rounded-md">
                <span className="text-slate-600 font-medium text-[11px]">ក្រុម៖</span>
                <select
                  value={activeTeamFilter}
                  onChange={(e) => setActiveTeamFilter(e.target.value)}
                  className="bg-transparent text-xs font-semibold text-slate-800 outline-hidden cursor-pointer"
                >
                  <option value="សរុបបណ្តាក្រុម">សរុបបណ្តាក្រុម (All Teams)</option>
                  {allTeams.map((t) => (
                    <option key={t} value={t}>
                      {t}
                    </option>
                  ))}
                </select>
              </div>
            )}

            {/* Day of Month Selector */}
            <div className="flex items-center gap-1 bg-slate-50 px-2 py-1 border border-slate-300 rounded-md">
              <Calendar className="w-3.5 h-3.5 text-slate-500" />
              <select
                value={dayFilter === 'all' ? 'all' : dayFilter}
                onChange={(e) => setDayFilter(e.target.value === 'all' ? 'all' : parseInt(e.target.value, 10))}
                className="bg-transparent text-xs font-semibold text-slate-800 outline-hidden cursor-pointer"
              >
                <option value="all">គ្រប់ថ្ងៃក្នុងខែ (Day 1..{daysInMonth})</option>
                {Array.from({ length: daysInMonth }, (_, i) => i + 1).map((d) => (
                  <option key={d} value={d}>
                    ថ្ងៃទី {d} ({toKhmerNum(d)})
                  </option>
                ))}
              </select>
            </div>

            {/* Category Filter */}
            <div className="flex items-center gap-1 bg-slate-50 px-2 py-1 border border-slate-300 rounded-md">
              <Layers className="w-3.5 h-3.5 text-slate-500" />
              <select
                value={categoryFilter}
                onChange={(e) => setCategoryFilter(e.target.value as any)}
                className="bg-transparent text-xs font-semibold text-slate-800 outline-hidden cursor-pointer"
              >
                <option value="Sticker">តែទិដ្ឋាការស្ទីកគ័រ (Sticker)</option>
                <option value="cEA">តែទិដ្ឋាការ cEA</option>
                <option value="all">គ្រប់ប្រភេទ (All)</option>
              </select>
            </div>

            {/* Overcounted Only Checkbox */}
            <label className="flex items-center gap-1.5 bg-amber-50 px-2.5 py-1.5 border border-amber-300 rounded-md cursor-pointer hover:bg-amber-100 transition">
              <input
                type="checkbox"
                checked={showOnlyOvercounted}
                onChange={(e) => setShowOnlyOvercounted(e.target.checked)}
                className="w-3.5 h-3.5 rounded text-amber-600 focus:ring-amber-500 cursor-pointer"
              />
              <span className="font-bold text-amber-900 text-xs">⚠️ បង្ហាញតែទិន្នន័យស្ទួន/រាប់លើស</span>
            </label>
          </div>

          {/* Quick Batch Action Buttons */}
          <div className="flex items-center gap-1.5">
            <button
              type="button"
              onClick={handleIncludeAllFiltered}
              className="px-2.5 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded text-xs font-semibold cursor-pointer shadow-xs transition"
              title="ធីកបញ្ចូលរាល់ទិន្នន័យក្នុងតារាងខាងក្រោម"
            >
              ✓ ធីកទាំងអស់
            </button>

            {stats.overcountedCount > 0 && (
              <button
                type="button"
                onClick={handleExcludeAllDuplicates}
                className="px-2.5 py-1.5 bg-amber-600 hover:bg-amber-700 text-white rounded text-xs font-bold cursor-pointer shadow-xs transition flex items-center gap-1"
                title="ដោះធីករាល់ទិន្នន័យដែលស្ទួន ឬរាប់លើសដើម្បីកុំឱ្យបូកចូលតារាង"
              >
                <AlertTriangle className="w-3.5 h-3.5" />
                <span>ដោះធីកទិន្នន័យស្ទួន ({stats.overcountedCount})</span>
              </button>
            )}

            <button
              type="button"
              onClick={handleExcludeAllFiltered}
              className="px-2.5 py-1.5 bg-slate-200 hover:bg-slate-300 text-slate-800 rounded text-xs font-semibold cursor-pointer transition"
              title="ដោះធីករាល់ទិន្នន័យក្នុងតារាងខាងក្រោម"
            >
              ✕ ដោះធីកទាំងអស់
            </button>

            <button
              type="button"
              onClick={handleReset}
              className="p-1.5 text-slate-600 hover:text-slate-900 hover:bg-slate-100 rounded transition cursor-pointer"
              title="កំណត់ដើមឡើងវិញ (Reset)"
            >
              <RotateCcw className="w-4 h-4" />
            </button>
          </div>
        </div>

        {/* Records Table Area */}
        <div className="flex-1 overflow-y-auto p-3">
          {filteredItems.length === 0 ? (
            <div className="py-16 text-center text-slate-500 bg-slate-50 rounded-lg border border-dashed border-slate-300">
              <CheckCircle2 className="w-10 h-10 text-slate-400 mx-auto mb-2" />
              <p className="font-semibold text-sm">ពុំមានទិន្នន័យតាមលក្ខខណ្ឌចម្រាញ់នេះឡើយ</p>
              <p className="text-xs text-slate-400 mt-1">
                សូមផ្លាស់ប្តូរថ្ងៃខែ ការស្វែងរក ឬប្រភេទប្រតិបត្តិការ
              </p>
            </div>
          ) : (
            <div className="border border-slate-200 rounded-lg overflow-hidden shadow-2xs">
              <table className="w-full text-xs text-left border-collapse">
                <thead>
                  <tr className="bg-slate-100 text-slate-700 font-bold border-b border-slate-300">
                    <th className="p-2.5 w-12 text-center">រាប់</th>
                    <th className="p-2.5 w-28">កាលបរិច្ឆេទ</th>
                    <th className="p-2.5 w-24">ប្រភេទ</th>
                    <th className="p-2.5 w-36">ប្រភព/ប្រតិបត្តិការ</th>
                    <th className="p-2.5">ឈ្មោះក្រុម</th>
                    <th className="p-2.5">បរិមាណតាមទិដ្ឋាការ</th>
                    <th className="p-2.5 text-center w-24">សរុបសន្លឹក</th>
                    <th className="p-2.5 text-center w-36">ស្ថានភាព</th>
                    <th className="p-2.5 text-center w-14">លុប</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredItems.map((item, idx) => {
                    const isExcluded = localExcludedIds.includes(item.id);
                    const isOvercounted = item.isDuplicateOrOvercounted;

                    return (
                      <tr
                        key={item.id}
                        className={`border-b border-slate-200 transition ${
                          isExcluded
                            ? 'bg-slate-100/70 text-slate-400 opacity-60'
                            : isOvercounted
                            ? 'bg-amber-50 hover:bg-amber-100/70'
                            : idx % 2 === 0
                            ? 'bg-white hover:bg-blue-50/50'
                            : 'bg-slate-50/60 hover:bg-blue-50/50'
                        }`}
                      >
                        {/* Checkbox (Include / Exclude) */}
                        <td className="p-2.5 text-center">
                          <input
                            type="checkbox"
                            checked={!isExcluded}
                            onChange={() => handleToggleItem(item.id)}
                            className="w-4 h-4 rounded text-blue-600 focus:ring-blue-500 cursor-pointer"
                            title={isExcluded ? 'ចុចដើម្បីរាប់បញ្ចូល' : 'ចុចដើម្បីដកចេញកុំឱ្យរាប់'}
                          />
                        </td>

                        {/* Date */}
                        <td className="p-2.5 font-times font-medium whitespace-nowrap">
                          <span
                            className={`px-1.5 py-0.5 rounded ${
                              item.dayNum === dayFilter
                                ? 'bg-blue-600 text-white font-bold'
                                : 'bg-slate-200 text-slate-800'
                            }`}
                          >
                            {item.date}
                          </span>
                        </td>

                        {/* Category */}
                        <td className="p-2.5">
                          <span
                            className={`px-2 py-0.5 rounded text-[10px] font-bold font-times ${
                              item.categoryType === 'Sticker'
                                ? 'bg-blue-600 text-white'
                                : 'bg-cyan-600 text-white'
                            }`}
                          >
                            {item.categoryType}
                          </span>
                        </td>

                        {/* Source / Operation */}
                        <td className="p-2.5">
                          <span
                            className={`px-2 py-0.5 rounded text-[10px] font-semibold ${
                              item.source === 'daily_form'
                                ? 'bg-indigo-100 text-indigo-800'
                                : 'bg-amber-100 text-amber-900'
                            }`}
                          >
                            {item.operationTitle}
                          </span>
                        </td>

                        {/* Team Name */}
                        <td className="p-2.5 font-semibold text-slate-800">
                          {item.teamName}
                        </td>

                        {/* Breakdown by Visa Type */}
                        <td className="p-2.5">
                          <div className="flex flex-wrap gap-1">
                            {Object.entries(item.values).map(([vt, valData]: [string, any]) => (
                              <span
                                key={vt}
                                className="px-1.5 py-0.5 bg-slate-200 text-slate-800 rounded font-times text-[10px] font-semibold flex items-center gap-1"
                                title={valData?.startSerial ? `ស៊េរី: ${valData.startSerial} - ${valData.endSerial}` : undefined}
                              >
                                <span>{vt}:</span>
                                <strong className="text-blue-900">{valData?.quantity || 0}</strong>
                              </span>
                            ))}
                          </div>
                        </td>

                        {/* Total Sheets */}
                        <td className="p-2.5 text-center font-bold font-times text-slate-900 text-xs">
                          {item.totalSheets.toLocaleString()}
                        </td>

                        {/* Status */}
                        <td className="p-2.5 text-center">
                          {isExcluded ? (
                            <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-bold bg-slate-200 text-slate-600">
                              ✕ ដកចេញពីការរាប់
                            </span>
                          ) : isOvercounted ? (
                            <span
                              className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-bold bg-amber-100 text-amber-900 border border-amber-300"
                              title={item.overcountReason || 'ទិន្នន័យស្ទួន ឬកត់ត្រាលើស'}
                            >
                              <AlertTriangle className="w-3 h-3 text-amber-600" />
                              <span>ស្ទួន / រាប់លើស</span>
                            </span>
                          ) : (
                            <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[10px] font-bold bg-emerald-100 text-emerald-800">
                              ✓ ធម្មតា
                            </span>
                          )}
                        </td>

                        {/* Action - Delete single record */}
                        <td className="p-2.5 text-center">
                          <button
                            type="button"
                            onClick={() => handleDeleteOperation(item)}
                            className="p-1 text-slate-400 hover:text-red-600 hover:bg-red-50 rounded transition cursor-pointer"
                            title="លុបទិន្នន័យនេះចេញ"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
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
        <div className="bg-slate-50 border-t border-slate-200 px-4 py-3 flex flex-wrap items-center justify-between gap-3 text-xs">
          <div className="flex items-center gap-2 text-slate-600">
            <Info className="w-4 h-4 text-blue-600 shrink-0" />
            <span>
              ចុច <strong>«យល់ព្រម & អនុវត្តលើតារាង»</strong> ដើម្បីធ្វើបច្ចុប្បន្នភាពតារាងស្ថិតិប្រចាំថ្ងៃភ្លាមៗ។
            </span>
          </div>

          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={onClose}
              className="px-4 py-2 bg-white border border-slate-300 hover:bg-slate-100 text-slate-700 rounded-lg font-semibold transition cursor-pointer shadow-2xs"
            >
              បោះបង់ (Cancel)
            </button>
            <button
              type="button"
              onClick={handleApply}
              className="px-5 py-2 bg-[#002060] hover:bg-[#1b3a7a] text-white rounded-lg font-bold shadow-sm transition cursor-pointer flex items-center gap-1.5"
            >
              <CheckSquare className="w-4 h-4 text-yellow-300" />
              <span>យល់ព្រម & អនុវត្តលើតារាង (Apply)</span>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
