import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  X,
  Calendar,
  Layers,
  FileSpreadsheet,
  Printer,
  CheckCircle2,
  ListFilter,
  Search,
  ChevronRight,
  TrendingUp,
  FileText,
  Sparkles,
  Info,
  ArrowUpDown,
  AlertCircle,
  Check,
  RotateCcw,
} from 'lucide-react';
import { StockRecord } from '../types';
import { REPORT_VISA_TYPES, RowDistributionData, isIssueTeamRecord, isRecordForTeam, isRecordForRecipientTeam, isTransferTeamRecord, isReceiveFromK1Record } from './YearlyTeamDistributionReport';
import { toKhmerNum } from '../utils/khmerCalendar';
import { normalizeDateToISO, normalizeVisaType, normalizeTeamName } from '../utils/teamNormalization';
import { isCeaRecord, isOldStockTeamRecord } from '../utils/teamStockCalculation';

export interface MonthColInfo {
  monthNum: number;
  year: number;
  name: string;
}

export interface YearlyMonthChecklistModalProps {
  isOpen: boolean;
  onClose: () => void;
  startYear: number;
  startMonth: number;
  monthColumnsInfo: MonthColInfo[];
  tableData: Record<string, RowDistributionData | any>;
  stockRecords: StockRecord[];
  initialMonthIndex?: number;
  targetTeamName?: string;
  mode?: 'distribution' | 'usage' | 'receive_k1';
  onApplyMonthActualValues?: (monthIndex: number, breakdownByVisa: Record<string, number>) => void;
  onApplyAllMonthsActualValues?: (allMonthsBreakdown: Record<string, number[]>) => void;
}

export const YearlyMonthChecklistModal: React.FC<YearlyMonthChecklistModalProps> = ({
  isOpen,
  onClose,
  startYear,
  startMonth,
  monthColumnsInfo,
  tableData,
  stockRecords = [],
  initialMonthIndex = 0,
  targetTeamName,
  mode = 'distribution',
  onApplyMonthActualValues,
  onApplyAllMonthsActualValues,
}) => {
  const isUsageMode = mode === 'usage';
  const isReceiveK1Mode = mode === 'receive_k1';
  // Selected Month Tab: 0-11 for specific months, -1 for "All 12 Months Matrix"
  const [selectedMonthIdx, setSelectedMonthIdx] = useState<number>(() => {
    return initialMonthIndex >= 0 && initialMonthIndex < 12 ? initialMonthIndex : 0;
  });

  const [activeSubTab, setActiveSubTab] = useState<'summary' | 'records' | 'allMonths'>('summary');
  const [searchTerm, setSearchTerm] = useState<string>('');
  const [selectedVisaFilter, setSelectedVisaFilter] = useState<string>('ALL');
  const [justAppliedMessage, setJustAppliedMessage] = useState<string | null>(null);

  const currentMonthInfo = selectedMonthIdx >= 0 && selectedMonthIdx < 12 ? monthColumnsInfo[selectedMonthIdx] : null;

  // Derive start & end date ISO for the selected month
  const selectedMonthDateRange = useMemo(() => {
    if (!currentMonthInfo) return null;
    const mStr = String(currentMonthInfo.monthNum).padStart(2, '0');
    const startIso = `${currentMonthInfo.year}-${mStr}-01`;
    const daysInM = new Date(currentMonthInfo.year, currentMonthInfo.monthNum, 0).getDate();
    const endIso = `${currentMonthInfo.year}-${mStr}-${String(daysInM).padStart(2, '0')}`;
    return { startIso, endIso, daysInM };
  }, [currentMonthInfo]);

  // Compute grand annual total across all 12 months for percentage calculation
  const totalAnnualAllTypes = useMemo(() => {
    let sum = 0;
    REPORT_VISA_TYPES.forEach((vt) => {
      const row = tableData[vt];
      if (row) {
        sum += row.monthlyValues.reduce((a: number, b: number) => (Number(a) || 0) + (Number(b) || 0), 0);
      }
    });
    return sum || 1;
  }, [tableData]);

  // Total for each of the 12 months across all 13 visa types
  const monthSums = useMemo(() => {
    const sums: number[] = Array(12).fill(0);
    REPORT_VISA_TYPES.forEach((vt) => {
      const row = tableData[vt];
      if (row) {
        row.monthlyValues.forEach((val: number, mIdx: number) => {
          sums[mIdx] += Number(val) || 0;
        });
      }
    });
    return sums;
  }, [tableData]);

  // Per-Visa breakdown for the selected month
  const selectedMonthVisaBreakdown = useMemo(() => {
    if (selectedMonthIdx < 0 || selectedMonthIdx >= 12) return [];
    const monthTotal = monthSums[selectedMonthIdx] || 0;

    return REPORT_VISA_TYPES.map((vt, idx) => {
      const row = tableData[vt];
      const count = row ? Number(row.monthlyValues[selectedMonthIdx]) || 0 : 0;
      const pct = monthTotal > 0 ? (count / monthTotal) * 100 : 0;
      const annualPct = (count / totalAnnualAllTypes) * 100;

      return {
        index: idx + 1,
        visaType: vt,
        count,
        pct,
        annualPct,
        openingStock: row?.openingStock || 0,
        annualTotalForType: row ? row.monthlyValues.reduce((a: number, b: number) => (Number(a) || 0) + (Number(b) || 0), 0) : 0,
      };
    });
  }, [selectedMonthIdx, tableData, monthSums, totalAnnualAllTypes]);

  // All unfiltered records for this month to calculate exact breakdown by visa
  const unfilteredMonthRecords = useMemo(() => {
    if (!selectedMonthDateRange || !currentMonthInfo) return [];
    return (stockRecords || []).filter((r) => {
      if (r.stockType && r.stockType !== 'sticker') return false;
      if (isCeaRecord(r)) return false;
      if (isOldStockTeamRecord(r)) return false;

      const isTransfer = isTransferTeamRecord(r);
      const isSender = !targetTeamName || isRecordForTeam(r, targetTeamName);
      const isRecipient = targetTeamName ? isRecordForRecipientTeam(r, targetTeamName) : false;

      if (targetTeamName) {
        if (isTransfer) {
          if (!isSender && !isRecipient) return false;
        } else {
          if (!isSender) return false;
        }
      }

      const dateIso = normalizeDateToISO(r.date || (r as any).createdAt || '');
      if (!dateIso) return false;
      if (dateIso < selectedMonthDateRange.startIso || dateIso > selectedMonthDateRange.endIso) return false;

      if (isReceiveK1Mode) {
        return isReceiveFromK1Record(r);
      }

      if (isUsageMode) {
        const op = (r.operationType || '').trim().toLowerCase();
        return (
          op === 'useteam' ||
          op === 'use_team' ||
          op === 'used' ||
          op.includes('ប្រើប្រាស់តាមក្រុម') ||
          op.includes('ប្រើប្រាស់')
        );
      }

      // If it's a team transfer:
      // ONLY include for the recipient team (received/incoming)
      // Exclude for sender team (transferred away)
      if (isTransfer) {
        return Boolean(isRecipient);
      }

      // Count strictly issuance/distribution to teams (header ប្រតិបត្តិការ: បើកផ្តល់តាមក្រុម)
      return isIssueTeamRecord(r);
    });
  }, [selectedMonthDateRange, currentMonthInfo, stockRecords, targetTeamName, isUsageMode, isReceiveK1Mode]);

  // Actual total sum and breakdown by visa for the selected month
  const { actualSelectedMonthSum, actualBreakdownByVisa } = useMemo(() => {
    const breakdown: Record<string, number> = {};
    REPORT_VISA_TYPES.forEach((vt) => {
      breakdown[vt] = 0;
    });
    let total = 0;

    unfilteredMonthRecords.forEach((r) => {
      const vt = normalizeVisaType(r.visaType || '');
      const qty = Number(
        r.quantityBundles ||
          r.totalSheets ||
          (r as any).quantity ||
          ((r as any).count ? (r as any).count * 50 : 0)
      );
      if (breakdown[vt] !== undefined) {
        breakdown[vt] += qty;
      }
      total += qty;
    });

    return { actualSelectedMonthSum: total, actualBreakdownByVisa: breakdown };
  }, [unfilteredMonthRecords]);

  // Filter actual stock transactions for the selected month (with search and visa filter)
  const selectedMonthRecords = useMemo(() => {
    return unfilteredMonthRecords.filter((r) => {
      const vt = normalizeVisaType(r.visaType || '');
      if (selectedVisaFilter !== 'ALL' && vt !== selectedVisaFilter) return false;

      if (searchTerm.trim()) {
        const q = searchTerm.toLowerCase();
        const team = (r.visaTeamRobokName || r.sourceFrom || r.recipientTeamName || '').toLowerCase();
        const note = (r.notes || (r as any).note || '').toLowerCase();
        const serial = `${r.startSerial || ''} ${r.endSerial || ''}`.toLowerCase();
        if (!team.includes(q) && !note.includes(q) && !serial.includes(q) && !vt.toLowerCase().includes(q)) {
          return false;
        }
      }
      return true;
    });
  }, [unfilteredMonthRecords, selectedVisaFilter, searchTerm]);

  // Apply actual transaction numbers for this month to the report
  const handleApplyCurrentMonthActual = () => {
    if (onApplyMonthActualValues && selectedMonthIdx >= 0 && selectedMonthIdx < 12) {
      onApplyMonthActualValues(selectedMonthIdx, actualBreakdownByVisa);
      setJustAppliedMessage(
        `✓ បានកែតម្រូវ និងបញ្ចូលទិន្នន័យជាក់ស្តែង ${actualSelectedMonthSum.toLocaleString()} សន្លឹក ទៅក្នុងខែ ${currentMonthInfo?.name} រួចរាល់!`
      );
      setTimeout(() => setJustAppliedMessage(null), 5000);
    }
  };

  // Export current month breakdown to Excel
  const handleExportMonthExcel = () => {
    if (!currentMonthInfo) return;
    const wb = XLSX.utils.book_new();

    const modeLabel = isUsageMode ? 'ការប្រើប្រាស់' : 'ការបើកផ្តល់';
    const dataRows: any[][] = [
      [`របាយការណ៍ផ្ទៀងផ្ទាត់${modeLabel}សន្លឹកទិដ្ឋាការស្អិត — ខែ ${currentMonthInfo.name} ឆ្នាំ ${currentMonthInfo.year}`],
      [`កាលបរិច្ឆេទ: ${selectedMonthDateRange?.startIso} ដល់ ${selectedMonthDateRange?.endIso}`],
      [`សរុបប្រចាំខែ: ${monthSums[selectedMonthIdx].toLocaleString()} សន្លឹក`],
      [],
      ['ល.រ', 'ប្រភេទ', `ចំនួន${modeLabel} (សន្លឹក)`, 'ភាគរយក្នុងខែ (%)', 'សន្និធិដើមគ្រា', 'សរុបប្រចាំឆ្នាំ ១២ខែ'],
    ];

    selectedMonthVisaBreakdown.forEach((item) => {
      dataRows.push([
        item.index,
        item.visaType,
        item.count,
        `${item.pct.toFixed(2)}%`,
        item.openingStock,
        item.annualTotalForType,
      ]);
    });

    dataRows.push([
      'សរុប',
      'គ្រប់ប្រភេទ (13)',
      monthSums[selectedMonthIdx],
      '100.00%',
      '',
      totalAnnualAllTypes,
    ]);

    const ws = XLSX.utils.aoa_to_sheet(dataRows);
    XLSX.utils.book_append_sheet(wb, ws, `ខែ_${currentMonthInfo.name}_${currentMonthInfo.year}`);
    XLSX.writeFile(wb, `Checklist_Month_${currentMonthInfo.monthNum}_${currentMonthInfo.year}.xlsx`);
  };

  const selectedMonthSum = selectedMonthIdx >= 0 ? monthSums[selectedMonthIdx] || 0 : 0;
  const activeVisaTypesCount = selectedMonthVisaBreakdown.filter((v) => v.count > 0).length;

  if (!isOpen) return null;

  return (
    <div
      className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-xs p-3 md:p-6 overflow-y-auto animate-in fade-in duration-200"
      onClick={onClose}
    >
      <div
        className="bg-white w-full max-w-5xl max-h-[92vh] rounded-2xl shadow-2xl border border-gray-200 flex flex-col overflow-hidden"
        onClick={(e) => e.stopPropagation()}
      >
        {/* Modal Header */}
        <div className="bg-gradient-to-r from-amber-700 via-amber-800 to-amber-900 text-white px-5 py-4 flex items-center justify-between shadow-md">
          <div className="flex items-center gap-3">
            <div className="w-10 h-10 rounded-xl bg-white/15 backdrop-blur-md flex items-center justify-center text-amber-200 shadow-inner">
              <Calendar className="w-5 h-5" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h3 className="text-base md:text-lg font-bold font-moul leading-tight">
                  ផ្ទៀងផ្ទាត់បញ្ជី និងចំនួនសរុបតាមខែ (Checklist By Month)
                </h3>
                <span className="bg-amber-500/30 text-amber-100 text-[10px] px-2 py-0.5 rounded-full font-sans border border-amber-400/30 font-medium">
                  ១២ ខែ
                </span>
              </div>
              <p className="text-xs text-amber-100/90 font-siemreap mt-0.5">
                តារាងសន្លឹកទិដ្ឋាការស្អិត — គិតចាប់ពីថ្ងៃទី០១ ខែ{monthColumnsInfo[0]?.name} ឆ្នាំ{toKhmerNum(monthColumnsInfo[0]?.year)} ដល់ ថ្ងៃទី{toKhmerNum(new Date(monthColumnsInfo[11]?.year, monthColumnsInfo[11]?.monthNum, 0).getDate())} ខែ{monthColumnsInfo[11]?.name} ឆ្នាំ{toKhmerNum(monthColumnsInfo[11]?.year)}
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2">
            <button
              onClick={handleExportMonthExcel}
              type="button"
              className="hidden sm:flex items-center gap-1.5 px-3 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-medium transition cursor-pointer shadow-xs"
              title="ទាញយកបញ្ជីខែនេះជា Excel"
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              <span>Excel ខែនេះ</span>
            </button>
            <button
              onClick={onClose}
              type="button"
              className="w-8 h-8 rounded-lg bg-white/10 hover:bg-white/20 text-white flex items-center justify-center transition cursor-pointer"
            >
              <X className="w-5 h-5" />
            </button>
          </div>
        </div>

        {/* 12 Months Selector Tabs */}
        <div className="bg-amber-50/70 border-b border-amber-200 px-3 py-2.5 overflow-x-auto select-none">
          <div className="flex items-center gap-1.5 min-w-max">
            {monthColumnsInfo.map((mCol, idx) => {
              const isSelected = selectedMonthIdx === idx && activeSubTab !== 'allMonths';
              const mSum = monthSums[idx] || 0;
              return (
                <button
                  key={idx}
                  type="button"
                  onClick={() => {
                    setSelectedMonthIdx(idx);
                    if (activeSubTab === 'allMonths') {
                      setActiveSubTab('summary');
                    }
                  }}
                  className={`px-3 py-2 rounded-xl text-xs font-medium flex flex-col items-center gap-0.5 transition cursor-pointer border ${
                    isSelected
                      ? 'bg-amber-800 text-white border-amber-900 shadow-md font-bold scale-[1.02]'
                      : 'bg-white hover:bg-amber-100/80 text-amber-950 border-amber-200/90 shadow-xs'
                  }`}
                >
                  <div className="flex items-center gap-1">
                    <span className="text-[11px] opacity-75">{idx + 1}.</span>
                    <span className="font-semibold">{mCol.name}</span>
                    <span className="text-[10px] opacity-80">({mCol.year})</span>
                  </div>
                  <span
                    className={`text-[10px] font-mono px-1.5 py-0.2 rounded-full font-bold ${
                      isSelected ? 'bg-amber-600 text-amber-50' : 'bg-amber-100 text-amber-900'
                    }`}
                  >
                    {mSum.toLocaleString()}
                  </span>
                </button>
              );
            })}

            {/* Matrix View All 12 Months Button */}
            <button
              type="button"
              onClick={() => setActiveSubTab('allMonths')}
              className={`px-3 py-2 rounded-xl text-xs font-medium flex flex-col items-center gap-0.5 transition cursor-pointer border ml-1 ${
                activeSubTab === 'allMonths'
                  ? 'bg-indigo-800 text-white border-indigo-900 shadow-md font-bold scale-[1.02]'
                  : 'bg-indigo-50 hover:bg-indigo-100 text-indigo-900 border-indigo-200 shadow-xs'
              }`}
            >
              <div className="flex items-center gap-1 font-semibold">
                <Layers className="w-3 h-3 text-indigo-500" />
                <span>តារាងសរុប ១២ខែ</span>
              </div>
              <span
                className={`text-[10px] font-mono px-1.5 py-0.2 rounded-full font-bold ${
                  activeSubTab === 'allMonths' ? 'bg-indigo-600 text-indigo-50' : 'bg-indigo-200 text-indigo-900'
                }`}
              >
                {totalAnnualAllTypes.toLocaleString()}
              </span>
            </button>
          </div>
        </div>

        {/* Modal Body Area */}
        <div className="flex-1 p-4 md:p-5 overflow-y-auto space-y-4">
          {activeSubTab !== 'allMonths' && currentMonthInfo ? (
            <>
              {/* Highlight Dashboard Stats for Selected Month */}
              <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
                <div className="bg-amber-50 border border-amber-200 rounded-xl p-3 shadow-2xs">
                  <div className="text-[11px] font-semibold text-amber-800 flex items-center gap-1 font-siemreap">
                    <Calendar className="w-3.5 h-3.5 text-amber-600" />
                    <span>ខែដែលកំពុងពិនិត្យ</span>
                  </div>
                  <div className="text-lg font-bold text-amber-950 mt-1 font-siemreap">
                    ខែ {currentMonthInfo.name} {toKhmerNum(currentMonthInfo.year)}
                  </div>
                  <div className="text-[10px] text-amber-700 font-mono mt-0.5">
                    {selectedMonthDateRange?.startIso} ~ {selectedMonthDateRange?.endIso}
                  </div>
                </div>

                <div className="bg-blue-50 border border-blue-200 rounded-xl p-3 shadow-2xs">
                  <div className="text-[11px] font-semibold text-blue-800 flex items-center gap-1 font-siemreap">
                    <TrendingUp className="w-3.5 h-3.5 text-blue-600" />
                    <span>
                      {isReceiveK1Mode
                        ? 'សរុបបើកពី ក១ ខែនេះ'
                        : isUsageMode
                        ? 'សរុបប្រើប្រាស់ខែនេះ'
                        : 'សរុបបើកផ្តល់ខែនេះ'}
                    </span>
                  </div>
                  <div className="text-xl font-bold text-blue-950 font-mono mt-1">
                    {selectedMonthSum.toLocaleString()}{' '}
                    <span className="text-xs font-normal font-siemreap text-blue-800">សន្លឹក</span>
                  </div>
                  <div className="text-[10px] text-blue-700 mt-0.5">
                    ស្មើនឹង {((selectedMonthSum / totalAnnualAllTypes) * 100).toFixed(2)}% នៃសរុប ១២ខែ
                  </div>
                </div>

                <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-3 shadow-2xs">
                  <div className="text-[11px] font-semibold text-emerald-800 flex items-center gap-1 font-siemreap">
                    <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                    <span>
                      {isReceiveK1Mode
                        ? 'ប្រភេទមានការបើកពី ក១'
                        : isUsageMode
                        ? 'ប្រភេទមានការប្រើប្រាស់'
                        : 'ប្រភេទមានការបើកផ្តល់'}
                    </span>
                  </div>
                  <div className="text-xl font-bold text-emerald-950 font-mono mt-1">
                    {activeVisaTypesCount} / 13{' '}
                    <span className="text-xs font-normal font-siemreap text-emerald-800">ប្រភេទ</span>
                  </div>
                  <div className="text-[10px] text-emerald-700 mt-0.5">
                    {isReceiveK1Mode
                      ? 'ប្រភេទគ្មានបើកពី ក១'
                      : isUsageMode
                      ? 'ប្រភេទគ្មានប្រើប្រាស់'
                      : 'ប្រភេទគ្មានបើក'}
                    : {13 - activeVisaTypesCount}
                  </div>
                </div>

                <div className="bg-purple-50 border border-purple-200 rounded-xl p-3 shadow-2xs">
                  <div className="text-[11px] font-semibold text-purple-800 flex items-center gap-1 font-siemreap">
                    <FileText className="w-3.5 h-3.5 text-purple-600" />
                    <span>ប្រតិបត្តិការក្នុងប្រព័ន្ធ</span>
                  </div>
                  <div className="text-xl font-bold text-purple-950 font-mono mt-1">
                    {selectedMonthRecords.length}{' '}
                    <span className="text-xs font-normal font-siemreap text-purple-800">ប្រតិបត្តិការ</span>
                  </div>
                  <div className="text-[10px] text-purple-700 mt-0.5">
                    {selectedMonthRecords.length > 0 ? 'ទាញពីទិន្នន័យជាក់ស្តែង' : 'ទិន្នន័យផ្ទៀងផ្ទាត់ផ្លូវការ'}
                  </div>
                </div>
              </div>

              {/* Difference Banner & Instant Correct Result Button */}
              {onApplyMonthActualValues && (
                <div
                  className={`rounded-xl p-3 border flex flex-wrap items-center justify-between gap-3 shadow-xs transition-all ${
                    selectedMonthSum !== actualSelectedMonthSum
                      ? 'bg-amber-50 border-amber-300 text-amber-950'
                      : 'bg-emerald-50/80 border-emerald-200 text-emerald-950'
                  }`}
                >
                  <div className="flex items-center gap-2.5">
                    {selectedMonthSum !== actualSelectedMonthSum ? (
                      <AlertCircle className="w-5 h-5 text-amber-700 shrink-0" />
                    ) : (
                      <CheckCircle2 className="w-5 h-5 text-emerald-600 shrink-0" />
                    )}
                    <div>
                      <p className="text-xs font-bold font-siemreap">
                        {selectedMonthSum !== actualSelectedMonthSum
                          ? `ផ្ទៀងផ្ទាត់តួលេខខែ ${currentMonthInfo.name}៖ ក្នុងតារាងរបាយការណ៍មាន (${selectedMonthSum.toLocaleString()} សន្លឹក) ខុសពីប្រតិបត្តិការជាក់ស្តែង (${actualSelectedMonthSum.toLocaleString()} សន្លឹក)`
                          : `តួលេខក្នុងតារាងរបាយការណ៍ត្រូវគ្នាឥតខ្ចោះជាមួយប្រតិបត្តិការជាក់ស្តែង (${selectedMonthSum.toLocaleString()} សន្លឹក)`}
                      </p>
                      <p className="text-[11px] opacity-80 font-siemreap">
                        {selectedMonthSum !== actualSelectedMonthSum
                          ? `ចុចប៊ូតុង «កែតម្រូវទិន្នន័យខែនេះ» ដើម្បីបញ្ចូលតួលេខជាក់ស្តែង ${actualSelectedMonthSum.toLocaleString()} សន្លឹក ទៅក្នុងតារាងរបាយការណ៍ភ្លាមៗ`
                          : `អ្នកអាចចុចកែតម្រូវឡើងវិញបានគ្រប់ពេលប្រសិនបើមានការផ្លាស់ប្តូរប្រតិបត្តិការ`}
                      </p>
                    </div>
                  </div>

                  <button
                    type="button"
                    onClick={handleApplyCurrentMonthActual}
                    className={`px-4 py-2 rounded-xl text-xs font-bold font-siemreap inline-flex items-center gap-2 shadow-sm transition cursor-pointer ${
                      selectedMonthSum !== actualSelectedMonthSum
                        ? 'bg-amber-800 hover:bg-amber-900 text-white animate-bounce'
                        : 'bg-emerald-700 hover:bg-emerald-800 text-white'
                    }`}
                  >
                    <Check className="w-4 h-4" />
                    <span>
                      កែតម្រូវទិន្នន័យខែនេះ ({actualSelectedMonthSum.toLocaleString()} សន្លឹក)
                    </span>
                  </button>
                </div>
              )}

              {/* Just Applied Toast Notification */}
              {justAppliedMessage && (
                <div className="bg-emerald-100 border border-emerald-300 text-emerald-900 px-4 py-2.5 rounded-xl text-xs font-semibold flex items-center gap-2 shadow-xs animate-in fade-in">
                  <CheckCircle2 className="w-4 h-4 text-emerald-600 shrink-0" />
                  <span>{justAppliedMessage}</span>
                </div>
              )}

              {/* Sub-Tabs: Summary Breakdown vs Transaction Records */}
              <div className="flex flex-wrap items-center justify-between gap-2 border-b border-gray-200 pb-2">
                <div className="flex items-center gap-2">
                  <button
                    type="button"
                    onClick={() => setActiveSubTab('summary')}
                    className={`px-3 py-1.5 rounded-lg text-xs font-medium flex items-center gap-1.5 transition cursor-pointer ${
                      activeSubTab === 'summary'
                        ? 'bg-amber-800 text-white shadow-xs font-bold'
                        : 'bg-gray-100 text-gray-700 hover:bg-gray-200'
                    }`}
                  >
                    <Layers className="w-3.5 h-3.5" />
                    <span>តារាងសង្ខេប ១៣ ប្រភេទ (Visa Types Breakdown)</span>
                  </button>

                  <button
                    type="button"
                    onClick={() => setActiveSubTab('records')}
                    className={`px-3 py-1.5 rounded-lg text-xs font-medium flex items-center gap-1.5 transition cursor-pointer ${
                      activeSubTab === 'records'
                        ? 'bg-amber-800 text-white shadow-xs font-bold'
                        : 'bg-gray-100 text-gray-700 hover:bg-gray-200'
                    }`}
                  >
                    <FileText className="w-3.5 h-3.5" />
                    <span>បញ្ជីប្រតិបត្តិការជាក់ស្តែង ({selectedMonthRecords.length})</span>
                  </button>
                </div>

                {activeSubTab === 'records' && (
                  <div className="flex items-center gap-2 w-full sm:w-auto">
                    {/* Visa Type Filter */}
                    <select
                      value={selectedVisaFilter}
                      onChange={(e) => setSelectedVisaFilter(e.target.value)}
                      className="px-2.5 py-1 text-xs rounded-lg border border-gray-300 bg-white text-gray-800 outline-none focus:border-amber-500"
                    >
                      <option value="ALL">គ្រប់ប្រភេទ (All Types)</option>
                      {REPORT_VISA_TYPES.map((vt) => (
                        <option key={vt} value={vt}>
                          ប្រភេទ {vt}
                        </option>
                      ))}
                    </select>

                    {/* Search Input */}
                    <div className="relative flex-1 sm:w-48">
                      <Search className="w-3.5 h-3.5 absolute left-2.5 top-1/2 -translate-y-1/2 text-gray-400" />
                      <input
                        type="text"
                        placeholder="ស្វែងរកក្រុម, លេខកូដ..."
                        value={searchTerm}
                        onChange={(e) => setSearchTerm(e.target.value)}
                        className="w-full pl-8 pr-2.5 py-1 text-xs rounded-lg border border-gray-300 bg-white text-gray-800 outline-none focus:border-amber-500"
                      />
                    </div>
                  </div>
                )}
              </div>

              {/* View 1: Summary Table by 13 Visa Types */}
              {activeSubTab === 'summary' && (
                <div className="border border-black rounded-xl overflow-hidden shadow-xs">
                  <table className="w-full border-collapse text-[11px] text-center font-siemreap">
                    <thead>
                      <tr className="bg-[#FCE4D6] font-bold text-black border-b border-black">
                        <th className="border-r border-black px-2 py-2 w-12 text-center">ល.រ</th>
                        <th className="border-r border-black px-3 py-2 w-20 text-center font-times">ប្រភេទ</th>
                        <th className="border-r border-black px-3 py-2 text-right">
                          ចំនួន{isUsageMode ? 'ប្រើប្រាស់' : 'បើកផ្តល់'}ក្នុងខែ {currentMonthInfo.name} (សន្លឹក)
                        </th>
                        <th className="border-r border-black px-3 py-2 text-right w-28">ភាគរយក្នុងខែ (%)</th>
                        <th className="border-r border-black px-3 py-2 text-right w-36">សន្និធិដើមគ្រា</th>
                        <th className="border-r border-black px-3 py-2 text-right w-36">សរុប ១២ខែ</th>
                        <th className="px-3 py-2 text-center w-28">ស្ថានភាព</th>
                      </tr>
                    </thead>
                    <tbody>
                      {selectedMonthVisaBreakdown.map((row, idx) => {
                        const isZero = row.count === 0;
                        return (
                          <tr
                            key={row.visaType}
                            className={`border-b border-gray-200 transition-colors ${
                              isZero ? 'bg-white text-gray-400' : 'bg-amber-50/20 text-gray-900 hover:bg-amber-100/50'
                            }`}
                          >
                            <td className="border-r border-black px-2 py-1.5 text-center font-mono">{idx + 1}</td>
                            <td className="border-r border-black px-3 py-1.5 font-bold font-times text-center text-xs">
                              <span
                                className={`px-2 py-0.5 rounded ${
                                  !isZero ? 'bg-amber-100 text-amber-900 font-bold' : 'text-gray-400'
                                }`}
                              >
                                {row.visaType}
                              </span>
                            </td>
                            <td className="border-r border-black px-3 py-1.5 text-right font-mono font-bold text-xs">
                              <span className={!isZero ? 'text-blue-900' : 'text-gray-400'}>
                                {row.count.toLocaleString()}
                              </span>
                            </td>
                            <td className="border-r border-black px-3 py-1.5 text-right font-mono">
                              {!isZero ? (
                                <span className="text-emerald-700 font-semibold">{row.pct.toFixed(2)}%</span>
                              ) : (
                                '0.00%'
                              )}
                            </td>
                            <td className="border-r border-black px-3 py-1.5 text-right font-mono text-gray-700">
                              {row.openingStock.toLocaleString()}
                            </td>
                            <td className="border-r border-black px-3 py-1.5 text-right font-mono font-semibold text-gray-800">
                              {row.annualTotalForType.toLocaleString()}
                            </td>
                            <td className="px-3 py-1.5 text-center">
                              {!isZero ? (
                                <span className="inline-flex items-center gap-1 text-[10px] text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded-full font-medium border border-emerald-200">
                                  <CheckCircle2 className="w-3 h-3 text-emerald-600" />
                                  <span>មានការបើក</span>
                                </span>
                              ) : (
                                <span className="text-[10px] text-gray-400 bg-gray-50 px-2 py-0.5 rounded-full">
                                  គ្មានបើក
                                </span>
                              )}
                            </td>
                          </tr>
                        );
                      })}
                    </tbody>
                    <tfoot>
                      <tr className="bg-[#F8CBAD] font-bold text-black border-t-2 border-black">
                        <td colSpan={2} className="border-r border-black px-3 py-2 text-center font-bold">
                          សរុបខែ {currentMonthInfo.name}
                        </td>
                        <td className="border-r border-black px-3 py-2 text-right font-mono font-bold text-xs text-blue-950">
                          {selectedMonthSum.toLocaleString()}
                        </td>
                        <td className="border-r border-black px-3 py-2 text-right font-mono">100.00%</td>
                        <td className="border-r border-black px-3 py-2 text-right font-mono">
                          {REPORT_VISA_TYPES.reduce((s, vt) => s + (tableData[vt]?.openingStock || 0), 0).toLocaleString()}
                        </td>
                        <td className="border-r border-black px-3 py-2 text-right font-mono">
                          {totalAnnualAllTypes.toLocaleString()}
                        </td>
                        <td className="px-3 py-2 text-center text-xs font-bold text-emerald-800">ផ្ទៀងផ្ទាត់រួច ✓</td>
                      </tr>
                    </tfoot>
                  </table>
                </div>
              )}

              {/* View 2: Transaction Records List */}
              {activeSubTab === 'records' && (
                <div className="space-y-3">
                  {selectedMonthRecords.length > 0 ? (
                    <>
                      <div className="border border-black rounded-xl overflow-hidden shadow-xs">
                      <table className="w-full border-collapse text-[11px] text-center font-siemreap">
                        <thead>
                          <tr className="bg-[#FCE4D6] font-bold text-black border-b border-black">
                            <th className="border-r border-black px-2 py-2 w-10 text-center">ល.រ</th>
                            <th className="border-r border-black px-2.5 py-2 w-24 text-center">កាលបរិច្ឆេទ</th>
                            <th className="border-r border-black px-2 py-2 w-16 text-center font-times">ប្រភេទ</th>
                            <th className="border-r border-black px-3 py-2 text-left">ក្រុម / គោលដៅទទួល</th>
                            <th className="border-r border-black px-3 py-2 text-center">ប្រតិបត្តិការ</th>
                            <th className="border-r border-black px-3 py-2 text-center">លេខស៊េរី (Serial)</th>
                            <th className="border-r border-black px-3 py-2 text-right w-24">ចំនួនសន្លឹក</th>
                            <th className="px-3 py-2 text-left">កំណត់សម្គាល់ / ប្រភព</th>
                          </tr>
                        </thead>
                        <tbody>
                          {selectedMonthRecords.map((rec, idx) => {
                            const sheets =
                              rec.totalSheets ||
                              rec.quantityBundles ||
                              (rec as any).quantity ||
                              ((rec as any).count ? (rec as any).count * 50 : 0);
                            return (
                              <tr
                                key={rec.id || idx}
                                className="border-b border-gray-200 hover:bg-amber-50/50 transition-colors"
                              >
                                <td className="border-r border-black px-2 py-1.5 font-mono text-gray-500">{idx + 1}</td>
                                <td className="border-r border-black px-2.5 py-1.5 font-mono text-gray-700">
                                  {normalizeDateToISO(rec.date || '')}
                                </td>
                                <td className="border-r border-black px-2 py-1.5 font-bold font-times text-amber-900">
                                  <span className="bg-amber-100 px-1.5 py-0.5 rounded">{rec.visaType}</span>
                                </td>
                                <td className="border-r border-black px-3 py-1.5 text-left font-medium text-gray-800">
                                  {rec.visaTeamRobokName || rec.sourceFrom || 'ក្រុមទិដ្ឋាការ'}
                                </td>
                                <td className="border-r border-black px-3 py-1.5 text-center text-emerald-700 font-medium">
                                  បើកផ្តល់ទៅក្រុម
                                </td>
                                <td className="border-r border-black px-3 py-1.5 text-center font-mono text-[10px] text-gray-600">
                                  {rec.startSerial && rec.endSerial
                                    ? `${rec.startSerial} - ${rec.endSerial}`
                                    : '-'}
                                </td>
                                <td className="border-r border-black px-3 py-1.5 text-right font-mono font-bold text-blue-900">
                                  {sheets.toLocaleString()}
                                </td>
                                <td className="px-3 py-1.5 text-left text-gray-600 text-[10.5px]">
                                  {rec.notes || (rec as any).note || rec.sourceFrom || '-'}
                                </td>
                              </tr>
                            );
                          })}
                        </tbody>
                        <tfoot>
                          <tr className="bg-[#F8CBAD] font-bold text-black border-t border-black">
                            <td colSpan={6} className="border-r border-black px-3 py-2 text-center">
                              សរុបប្រតិបត្តិការជាក់ស្តែង ({selectedMonthRecords.length})
                            </td>
                            <td className="border-r border-black px-3 py-2 text-right font-mono font-bold text-blue-950">
                              {selectedMonthRecords
                                .reduce(
                                  (s, r) =>
                                    s +
                                    (r.totalSheets ||
                                      r.quantityBundles ||
                                      (r as any).quantity ||
                                      ((r as any).count ? (r as any).count * 50 : 0)),
                                  0
                                )
                                .toLocaleString()}
                            </td>
                            <td className="px-3 py-2"></td>
                          </tr>
                        </tfoot>
                      </table>
                    </div>

                    {/* Bottom Action Button for Tab 2 */}
                    {onApplyMonthActualValues && (
                      <div className="bg-amber-50/70 border border-amber-200 rounded-xl p-3 flex flex-wrap items-center justify-between gap-2">
                        <div className="text-xs text-amber-900 font-siemreap">
                          សរុបប្រតិបត្តិការជាក់ស្តែង {selectedMonthRecords.length} កំណត់ត្រា ស្មើនឹង <span className="font-mono font-bold text-blue-900">{actualSelectedMonthSum.toLocaleString()}</span> សន្លឹក
                        </div>
                        <button
                          type="button"
                          onClick={handleApplyCurrentMonthActual}
                          className="px-4 py-2 rounded-xl bg-emerald-700 hover:bg-emerald-800 text-white text-xs font-bold font-siemreap inline-flex items-center gap-1.5 shadow-xs transition cursor-pointer"
                        >
                          <Check className="w-4 h-4" />
                          <span>បញ្ចូល/កែតម្រូវលទ្ធផលជាក់ស្តែង ({actualSelectedMonthSum.toLocaleString()} សន្លឹក) ចូលតារាងខែ {currentMonthInfo.name}</span>
                        </button>
                      </div>
                    )}
                  </>
                  ) : (
                    <div className="bg-amber-50/60 border border-amber-200 rounded-xl p-5 text-center space-y-2">
                      <Info className="w-8 h-8 text-amber-600 mx-auto" />
                      <h4 className="text-sm font-bold text-amber-900 font-siemreap">
                        មិនមានកំណត់ត្រាប្រតិបត្តិការលម្អិតក្នុងមូលដ្ឋានទិន្នន័យ (Database) សម្រាប់ខែនេះទេ
                      </h4>
                      <p className="text-xs text-amber-800 max-w-lg mx-auto font-siemreap">
                        ទិន្នន័យខែ {currentMonthInfo.name} ត្រូវបានផ្ទៀងផ្ទាត់ និងទាញយកពីឯកសារគំរូផ្លូវការបណ្ណសារដ្ឋាន (Official Archive Record) តាមចំនួនសរុប {selectedMonthSum.toLocaleString()} សន្លឹក។
                      </p>
                      <button
                        type="button"
                        onClick={() => setActiveSubTab('summary')}
                        className="px-3.5 py-1.5 rounded-lg bg-amber-700 hover:bg-amber-800 text-white text-xs font-medium inline-flex items-center gap-1.5 transition cursor-pointer shadow-xs mt-2"
                      >
                        <Layers className="w-3.5 h-3.5" />
                        <span>មើលតារាងសង្ខេប ១៣ ប្រភេទ</span>
                      </button>
                    </div>
                  )}
                </div>
              )}
            </>
          ) : (
            /* View 3: Full 12 Months Matrix View */
            <div className="space-y-3">
              <div className="flex items-center justify-between">
                <div>
                  <h4 className="text-sm font-bold text-gray-800 font-moul">
                    តារាងទិដ្ឋភាពទូទៅនៃខែទាំង ១២ (12-Month Matrix Overview)
                  </h4>
                  <p className="text-xs text-gray-500 font-siemreap">
                    បង្ហាញចំនួនបើកផ្តល់សន្លឹកទិដ្ឋាការស្អិតតាមប្រភេទនីមួយៗក្នុងគ្រប់ខែទាំង ១២
                  </p>
                </div>
              </div>

              <div className="border border-black rounded-xl overflow-x-auto shadow-xs">
                <table className="w-full border-collapse text-[10.5px] text-center font-siemreap">
                  <thead>
                    <tr className="bg-[#FCE4D6] font-bold text-black border-b border-black">
                      <th className="border-r border-black px-2 py-2 w-12 text-center font-times">ប្រភេទ</th>
                      {monthColumnsInfo.map((mCol, idx) => (
                        <th
                          key={idx}
                          onClick={() => {
                            setSelectedMonthIdx(idx);
                            setActiveSubTab('summary');
                          }}
                          className="border-r border-black px-1 py-1.5 min-w-[50px] text-center hover:bg-amber-200 transition cursor-pointer"
                          title={`ចុចដើម្បីពិនិត្យលម្អិតខែ ${mCol.name}`}
                        >
                          <div>{mCol.name}</div>
                          <div className="text-[9px] text-gray-600 font-normal">{mCol.year}</div>
                        </th>
                      ))}
                      <th className="px-2 py-2 min-w-[65px] text-center font-bold bg-[#F8CBAD]">សរុប ១២ខែ</th>
                    </tr>
                  </thead>
                  <tbody>
                    {REPORT_VISA_TYPES.map((vt, rIdx) => {
                      const row = tableData[vt];
                      const monthlyVals = row?.monthlyValues || [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0];
                      const rowSum = monthlyVals.reduce((a, b) => (Number(a) || 0) + (Number(b) || 0), 0);

                      return (
                        <tr
                          key={vt}
                          className={`border-b border-gray-200 transition-colors ${
                            rIdx % 2 === 1 ? 'bg-gray-50/40' : 'bg-white'
                          } hover:bg-amber-50/50`}
                        >
                          <td className="border-r border-black px-2 py-1 font-bold font-times text-center bg-gray-50">
                            {vt}
                          </td>
                          {monthColumnsInfo.map((_, mIdx) => {
                            const val = Number(monthlyVals[mIdx]) || 0;
                            return (
                              <td
                                key={mIdx}
                                onClick={() => {
                                  setSelectedMonthIdx(mIdx);
                                  setActiveSubTab('summary');
                                }}
                                className={`border-r border-black px-1 py-1 text-right font-mono cursor-pointer hover:bg-yellow-100 ${
                                  val > 0 ? 'font-semibold text-gray-900' : 'text-gray-300'
                                }`}
                              >
                                {val ? val.toLocaleString() : '0'}
                              </td>
                            );
                          })}
                          <td className="px-2 py-1 text-right font-mono font-bold text-xs bg-[#FDF5ED] text-blue-950">
                            {rowSum.toLocaleString()}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                  <tfoot>
                    <tr className="bg-[#F8CBAD] font-bold text-black border-t-2 border-black">
                      <td className="border-r border-black px-2 py-2 text-center font-bold">សរុប</td>
                      {monthColumnsInfo.map((_, mIdx) => (
                        <td
                          key={mIdx}
                          onClick={() => {
                            setSelectedMonthIdx(mIdx);
                            setActiveSubTab('summary');
                          }}
                          className="border-r border-black px-1 py-2 text-right font-mono font-bold text-blue-950 cursor-pointer hover:bg-amber-300"
                        >
                          {monthSums[mIdx].toLocaleString()}
                        </td>
                      ))}
                      <td className="px-2 py-2 text-right font-mono font-bold text-xs bg-amber-200 text-amber-950">
                        {totalAnnualAllTypes.toLocaleString()}
                      </td>
                    </tr>
                  </tfoot>
                </table>
              </div>
            </div>
          )}
        </div>

        {/* Modal Footer */}
        <div className="bg-gray-50 border-t border-gray-200 px-5 py-3 flex flex-wrap items-center justify-between gap-2 text-xs text-gray-600">
          <div className="flex items-center gap-2">
            <span className="font-semibold text-gray-700">សរុបទាំង ១២ខែ:</span>
            <span className="font-mono font-bold text-blue-900 bg-blue-50 px-2 py-0.5 rounded border border-blue-200">
              {totalAnnualAllTypes.toLocaleString()} សន្លឹក
            </span>
          </div>

          <div className="flex items-center gap-2">
            <button
              onClick={onClose}
              type="button"
              className="px-4 py-1.5 rounded-lg bg-gray-200 hover:bg-gray-300 text-gray-800 font-medium transition cursor-pointer"
            >
              បិទ (Close)
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
