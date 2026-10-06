import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  X,
  FileSpreadsheet,
  Printer,
  CheckCircle2,
  ListFilter,
  Search,
  FileText,
  Sparkles,
  Info,
  Layers,
  Calendar,
  Building,
  Hash,
  ArrowDownToLine,
  TrendingUp,
  Share2,
} from 'lucide-react';
import { StockRecord } from '../types';
import { REPORT_VISA_TYPES, isIssueTeamRecord, isRecordForTeam, isRecordForRecipientTeam, isTransferTeamRecord } from './YearlyTeamDistributionReport';
import { toKhmerNum, KHMER_MONTHS } from '../utils/khmerCalendar';
import { normalizeDateToISO, normalizeVisaType, normalizeTeamName } from '../utils/teamNormalization';
import { isCeaRecord, isOldStockTeamRecord } from '../utils/teamStockCalculation';
import { printA4Document } from '../utils/printHelper';

export interface MonthColInfo {
  monthNum: number;
  year: number;
  name: string;
}

export interface YearlyReceivedK2ChecklistModalProps {
  isOpen: boolean;
  onClose: () => void;
  startYear: number;
  startMonth: number;
  monthColumnsInfo: MonthColInfo[];
  tableData: Record<string, any>;
  stockRecords: StockRecord[];
  mode?: 'distribution' | 'usage' | 'receive_k1' | 'k2_issue_team';
  targetTeamName?: string;
  onApplyAllIssuedValues?: (visaValues: Record<string, number>) => void;
}

// Visa type full Khmer names
const VISA_TYPE_NAMES: Record<string, string> = {
  T: 'ទេសចរណ៍ (T)',
  T1: 'ទេសចរណ៍ពិសេស (T1)',
  T2: 'ទេសចរណ៍ពិសេស (T2)',
  T3: 'ទេសចរណ៍ពិសេស (T3)',
  E: 'ធម្មតា (E)',
  E1: 'ធម្មតាពិសេស (E1)',
  E2: 'ធម្មតាពិសេស (E2)',
  E3: 'ធម្មតាពិសេស (E3)',
  D: 'ការទូត (D)',
  K: 'ពិសេស (K)',
  A: 'ផ្លូវការ (A)',
  B: 'រដ្ឋបាល (B)',
  C: 'គួរសម (C)',
};

export const YearlyReceivedK2ChecklistModal: React.FC<YearlyReceivedK2ChecklistModalProps> = ({
  isOpen,
  onClose,
  startYear,
  startMonth,
  monthColumnsInfo,
  tableData,
  stockRecords = [],
  mode = 'usage',
  targetTeamName,
  onApplyAllIssuedValues,
}) => {
  const [activeTab, setActiveTab] = useState<'summary' | 'records' | 'monthly'>('summary');
  const [searchTerm, setSearchTerm] = useState<string>('');
  const [selectedVisaFilter, setSelectedVisaFilter] = useState<string>('ALL');
  const [selectedMonthFilter, setSelectedMonthFilter] = useState<string>('ALL');
  const [appliedSuccessMsg, setAppliedSuccessMsg] = useState<string | null>(null);

  // Date range
  const startDateStr = `${startYear}-${String(startMonth).padStart(2, '0')}-01`;
  const endYear = monthColumnsInfo[11]?.year || startYear;
  const endMonth = monthColumnsInfo[11]?.monthNum || startMonth;
  const daysInEndMonth = new Date(endYear, endMonth, 0).getDate();
  const endDateStr = `${endYear}-${String(endMonth).padStart(2, '0')}-${String(daysInEndMonth).padStart(2, '0')}`;

  // Find all records of issuance from K2 to teams and transfer deductions in stockRecords for the 12-month period
  const receivedK2Records = useMemo(() => {
    if (!stockRecords || stockRecords.length === 0) return [];

    return stockRecords.filter((rec) => {
      if (rec.stockType && rec.stockType !== 'sticker') return false;
      if (isCeaRecord(rec)) return false;
      if (isOldStockTeamRecord(rec)) return false;

      const isTransfer = isTransferTeamRecord(rec);
      const isSender = !targetTeamName || isRecordForTeam(rec, targetTeamName);
      const isRecipient = targetTeamName ? isRecordForRecipientTeam(rec, targetTeamName) : false;

      if (targetTeamName) {
        if (isTransfer) {
          if (!isSender && !isRecipient) return false;
        } else {
          if (!isSender) return false;
        }
      }

      const dIso = normalizeDateToISO(rec.date || (rec as any).createdAt || '');
      if (!dIso || dIso < startDateStr || dIso > endDateStr) return false;

      // Count strictly issuance/distribution to teams (បើកផ្តល់តាមក្រុម / បើកពីក២) OR transfer
      return isIssueTeamRecord(rec) || isTransfer;
    });
  }, [stockRecords, startDateStr, endDateStr, targetTeamName]);

  // Actual Gross Issue and Transfer Deduction computed from raw records
  const { grossIssueByVisa, transferDeductionByVisa, netCalculatedByVisa } = useMemo(() => {
    const gross: Record<string, number> = {};
    const transfer: Record<string, number> = {};
    const net: Record<string, number> = {};

    REPORT_VISA_TYPES.forEach((vt) => {
      gross[vt] = 0;
      transfer[vt] = 0;
      net[vt] = 0;
    });

    receivedK2Records.forEach((rec) => {
      const vt = normalizeVisaType(rec.visaType);
      if (!vt || gross[vt] === undefined) {
        if (!REPORT_VISA_TYPES.includes(vt as any)) return;
      }

      const qty = Number(
        (rec as any).quantity ||
          (rec as any).quantityBundles ||
          rec.totalSheets ||
          ((rec as any).count ? (rec as any).count * 50 : 0) ||
          0
      );
      if (!qty) return;

      const isTransfer = isTransferTeamRecord(rec);
      const isRecipient = targetTeamName ? isRecordForRecipientTeam(rec, targetTeamName) : false;

      if (isTransfer) {
        if (targetTeamName && isRecipient) {
          // For recipient team: this transfer is received stock (gross)!
          gross[vt] = (gross[vt] || 0) + qty;
        } else {
          // For sender team or office: this transfer is outgoing deduction!
          transfer[vt] = (transfer[vt] || 0) + qty;
        }
      } else {
        gross[vt] = (gross[vt] || 0) + qty;
      }
    });

    REPORT_VISA_TYPES.forEach((vt) => {
      net[vt] = (gross[vt] || 0) - (transfer[vt] || 0);
    });

    return {
      grossIssueByVisa: gross,
      transferDeductionByVisa: transfer,
      netCalculatedByVisa: net,
    };
  }, [receivedK2Records, targetTeamName]);

  // Compute total received from K2 per visa type from tableData or records
  const receivedK2ByVisa = useMemo(() => {
    const map: Record<string, number> = {};
    REPORT_VISA_TYPES.forEach((vt) => {
      const row = tableData[vt];
      // In receive_k1/k2_issue_team mode, default to netCalculatedByVisa or row.issuedToTeams
      if (mode === 'receive_k1' || mode === 'k2_issue_team') {
        if (typeof row?.issuedToTeams === 'number') {
          map[vt] = row.issuedToTeams;
        } else {
          map[vt] = netCalculatedByVisa[vt] || 0;
        }
      } else if (row) {
        if (typeof row.receivedK2 === 'number') {
          map[vt] = row.receivedK2;
        } else if (typeof row.totalReceived === 'number') {
          map[vt] = row.totalReceived;
        } else if (Array.isArray(row.monthlyValues)) {
          map[vt] = row.monthlyValues.reduce((a: number, b: number) => (Number(a) || 0) + (Number(b) || 0), 0);
        } else {
          map[vt] = 0;
        }
      } else {
        map[vt] = netCalculatedByVisa[vt] || 0;
      }
    });
    return map;
  }, [tableData, mode, netCalculatedByVisa]);

  // Total gross and net stats
  const totalGrossIssued = useMemo(() => {
    return (Object.values(grossIssueByVisa) as number[]).reduce((a, b) => Number(a) + (Number(b) || 0), 0);
  }, [grossIssueByVisa]);

  const totalTransferDeduction = useMemo(() => {
    return (Object.values(transferDeductionByVisa) as number[]).reduce((a, b) => Number(a) + (Number(b) || 0), 0);
  }, [transferDeductionByVisa]);

  const totalNetCalculated = useMemo(() => {
    return (Object.values(netCalculatedByVisa) as number[]).reduce((a, b) => Number(a) + (Number(b) || 0), 0);
  }, [netCalculatedByVisa]);

  // Monthly values for each visa type across the 12 months
  const monthlyBreakdownByVisa = useMemo(() => {
    const map: Record<string, number[]> = {};
    REPORT_VISA_TYPES.forEach((vt) => {
      const row = tableData[vt];
      if (row && Array.isArray(row.monthlyValues) && mode === 'distribution') {
        map[vt] = [...row.monthlyValues];
      } else {
        // Compute from receivedK2Records (issuance minus transfer deductions)
        const months = Array(12).fill(0);
        receivedK2Records.forEach((r) => {
          if (normalizeVisaType(r.visaType) === vt) {
            const dIso = normalizeDateToISO(r.date);
            if (dIso) {
              const [y, m] = dIso.split('-').map(Number);
              const mIdx = monthColumnsInfo.findIndex((col) => col.year === y && col.monthNum === m);
              if (mIdx >= 0 && mIdx < 12) {
                const qty = Number((r as any).quantity || (r as any).quantityBundles || r.totalSheets || ((r as any).count ? (r as any).count * 50 : 0)) || 0;
                const op = (r.operationType || '').toLowerCase();
                const isTransfer =
                  op === 'transferteam' ||
                  op === 'transferuseteam' ||
                  op === 'transfer' ||
                  op.includes('transfer') ||
                  op.includes('ផ្ទេរ') ||
                  (r.sourceFrom && r.sourceFrom.includes('ផ្ទេរ')) ||
                  ((r as any).notes && (r as any).notes.includes('ផ្ទេរ'));

                if (isTransfer) {
                  months[mIdx] -= qty;
                } else {
                  months[mIdx] += qty;
                }
              }
            }
          }
        });
        map[vt] = months;
      }
    });
    return map;
  }, [tableData, receivedK2Records, monthColumnsInfo, mode]);

  // Grand Total Received from K2
  const totalReceivedK2 = useMemo(() => {
    return (Object.values(receivedK2ByVisa) as number[]).reduce((acc: number, val: number) => acc + (Number(val) || 0), 0);
  }, [receivedK2ByVisa]);

  // Filtered Records
  const filteredRecords = useMemo(() => {
    return receivedK2Records.filter((rec) => {
      const vt = normalizeVisaType(rec.visaType);
      if (selectedVisaFilter !== 'ALL' && vt !== selectedVisaFilter) return false;

      if (selectedMonthFilter !== 'ALL') {
        const mIdx = Number(selectedMonthFilter);
        const colInfo = monthColumnsInfo[mIdx];
        if (colInfo) {
          const dIso = normalizeDateToISO(rec.date);
          if (dIso) {
            const [y, m] = dIso.split('-').map(Number);
            if (y !== colInfo.year || m !== colInfo.monthNum) return false;
          }
        }
      }

      if (!searchTerm) return true;
      const term = searchTerm.toLowerCase();
      const team = (rec.teamName || (rec as any).sourceFrom || '').toLowerCase();
      const officer = (rec.responsibleOfficer || (rec as any).receiver || (rec as any).giver || '').toLowerCase();
      const sn = `${rec.startSerial || ''} ${rec.endSerial || ''}`.toLowerCase();
      const notes = (rec.notes || (rec as any).memo || '').toLowerCase();
      const doc = (rec.documentNumber || (rec as any).voucherNo || '').toLowerCase();

      return (
        vt.toLowerCase().includes(term) ||
        team.includes(term) ||
        officer.includes(term) ||
        sn.includes(term) ||
        notes.includes(term) ||
        doc.includes(term)
      );
    });
  }, [receivedK2Records, selectedVisaFilter, selectedMonthFilter, searchTerm, monthColumnsInfo]);

  // Export to Excel
  const handleExportExcel = () => {
    const wb = XLSX.utils.book_new();

    const titleRows = [
      ['ព្រះរាជាណាចក្រកម្ពុជា'],
      ['ជាតិ សាសនា ព្រះមហាក្សត្រ'],
      ['***'],
      ['តារាងផ្ទៀងផ្ទាត់បញ្ជីទិដ្ឋាការ បើកពីក២ (ការិយាល័យទិដ្ឋាការចូល) តាមប្រភេទនីមួយៗ'],
      [`ដោយគិតចាប់ពីថ្ងៃទី០១ ខែ${KHMER_MONTHS[startMonth - 1] || startMonth} ឆ្នាំ${toKhmerNum(startYear)} ដល់ថ្ងៃទី${toKhmerNum(daysInEndMonth)} ខែ${monthColumnsInfo[11]?.name} ឆ្នាំ${toKhmerNum(endYear)}`],
      [],
      ['ល.រ', 'ប្រភេទវីសា', 'ឈ្មោះពេញ', 'ចំនួនបើកពីក២ (សន្លឹក)', 'ភាគរយ (%)', 'ស្ថានភាព'],
    ];

    const dataRows = REPORT_VISA_TYPES.map((vt, index) => {
      const count = receivedK2ByVisa[vt] || 0;
      const pct = totalReceivedK2 > 0 ? ((count / totalReceivedK2) * 100).toFixed(2) + '%' : '0%';
      return [
        index + 1,
        vt,
        VISA_TYPE_NAMES[vt] || vt,
        count,
        pct,
        count > 0 ? 'មានការបើកផ្តល់' : 'គ្មាន',
      ];
    });

    const totalRow = [
      '',
      'សរុប',
      'សរុបទូទាំងគ្រប់ប្រភេទ',
      totalReceivedK2,
      '100%',
      '',
    ];

    const ws = XLSX.utils.aoa_to_sheet([...titleRows, ...dataRows, totalRow]);
    XLSX.utils.book_append_sheet(wb, ws, 'បញ្ជីបើកពីក២_សង្ខេប');

    // Sheet 2: Monthly Matrix
    const monthHeader = ['ល.រ', 'ប្រភេទ', ...monthColumnsInfo.map((m) => `${m.name} ${m.year}`), 'សរុប'];
    const monthlyMatrixRows = REPORT_VISA_TYPES.map((vt, idx) => {
      const mVals = monthlyBreakdownByVisa[vt] || Array(12).fill(0);
      const rowSum = mVals.reduce((a, b) => a + b, 0);
      return [idx + 1, vt, ...mVals, rowSum];
    });
    const monthlyTotalRow = [
      '',
      'សរុប',
      ...Array(12).fill(0).map((_, mIdx) => {
        return REPORT_VISA_TYPES.reduce((s, vt) => s + (monthlyBreakdownByVisa[vt]?.[mIdx] || 0), 0);
      }),
      totalReceivedK2,
    ];
    const ws2 = XLSX.utils.aoa_to_sheet([
      ['តារាងសន្លឹកទិដ្ឋាការស្អិត បើកពីក២ ប្រចាំ ១២ខែ'],
      [],
      monthHeader,
      ...monthlyMatrixRows,
      monthlyTotalRow,
    ]);
    XLSX.utils.book_append_sheet(wb, ws2, 'ម៉ាទ្រីសបើកពីក២_១២ខែ');

    XLSX.writeFile(wb, `តារាងផ្ទៀងផ្ទាត់ទិដ្ឋាការបើកពីក២_${startYear}_${endYear}.xlsx`);
  };

  const handlePrint = () => {
    printA4Document('yearly-received-k2-checklist-container', {
      orientation: 'landscape',
      documentTitle: 'ផ្ទៀងផ្ទាត់បញ្ជីទិដ្ឋាការបើកពីក២',
    });
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-xs p-2 sm:p-4 overflow-y-auto animate-fade-in">
      <div className="bg-white w-full max-w-5xl rounded-2xl shadow-2xl border border-gray-200 flex flex-col max-h-[92vh] overflow-hidden my-auto">
        {/* Modal Header */}
        <div className="bg-linear-to-r from-blue-900 via-indigo-900 to-slate-950 text-white px-5 py-3.5 flex items-center justify-between shadow-md">
          <div className="flex items-center gap-3">
            <div className="p-2 bg-white/10 rounded-xl backdrop-blur-xs border border-white/20">
              <ArrowDownToLine className="w-5 h-5 text-cyan-300" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h2 className="text-base sm:text-lg font-bold font-siemreap tracking-wide">
                  ផ្ទៀងផ្ទាត់បញ្ជីទិដ្ឋាការ បើកពីក២ (ការិយាល័យទិដ្ឋាការចូល) ប្រចាំឆ្នាំ (១២ខែ)
                </h2>
                <span className="text-[11px] font-medium bg-blue-700/80 border border-blue-500/40 text-cyan-200 px-2 py-0.5 rounded-full">
                  {mode === 'usage' ? 'តារាងប្រើប្រាស់តាមក្រុម' : 'តារាងបើកផ្តល់តាមក្រុម'}
                </span>
              </div>
              <p className="text-xs text-blue-200 font-siemreap mt-0.5">
                ដោយគិតចាប់ពីថ្ងៃទី០១ ខែ{KHMER_MONTHS[startMonth - 1] || startMonth} ឆ្នាំ{toKhmerNum(startYear)} រហូតដល់ថ្ងៃទី{toKhmerNum(daysInEndMonth)} ខែ{monthColumnsInfo[11]?.name} ឆ្នាំ{toKhmerNum(endYear)}
              </p>
            </div>
          </div>

          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={handleExportExcel}
              className="bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold px-3 py-1.5 rounded-lg flex items-center gap-1.5 transition shadow-xs cursor-pointer border border-emerald-400"
              title="ទាញយកជា Excel"
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              <span className="hidden sm:inline">Excel</span>
            </button>
            <button
              type="button"
              onClick={onClose}
              className="p-1.5 rounded-lg bg-white/10 hover:bg-white/20 text-white transition cursor-pointer"
              title="បិទ"
            >
              <X className="w-5 h-5" />
            </button>
          </div>
        </div>

        {/* Highlight Stats Dashboard */}
        <div className="bg-gray-50 border-b border-gray-200 px-5 py-3">
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
            <div className="bg-blue-50 border border-blue-200 rounded-xl p-3 shadow-2xs">
              <div className="text-[11px] font-semibold text-blue-800 flex items-center gap-1 font-siemreap">
                <ArrowDownToLine className="w-3.5 h-3.5 text-blue-600" />
                <span>{mode === 'receive_k1' || mode === 'k2_issue_team' ? 'សរុប បើកផ្តល់ជូនក្រុម' : 'សរុប បើកពីក២'}</span>
              </div>
              <div className="text-2xl font-bold text-blue-950 font-times mt-1">
                {(mode === 'receive_k1' || mode === 'k2_issue_team' ? totalNetCalculated : totalReceivedK2).toLocaleString()}{' '}
                <span className="text-xs font-normal font-siemreap text-blue-800">សន្លឹក</span>
              </div>
              <div className="text-[10px] text-blue-700 mt-0.5">
                {mode === 'receive_k1' || mode === 'k2_issue_team'
                  ? `បើកផ្តល់ ${totalGrossIssued.toLocaleString()} - ផ្ទេរ ${totalTransferDeduction.toLocaleString()}`
                  : 'សរុប ១២ខែ គ្រប់ប្រភេទវីសា'}
              </div>
            </div>

            <div className="bg-indigo-50 border border-indigo-200 rounded-xl p-3 shadow-2xs">
              <div className="text-[11px] font-semibold text-indigo-800 flex items-center gap-1 font-siemreap">
                <Hash className="w-3.5 h-3.5 text-indigo-600" />
                <span>ប្រភេទ T (ទេសចរណ៍)</span>
              </div>
              <div className="text-xl font-bold text-indigo-950 font-times mt-1">
                {((mode === 'receive_k1' || mode === 'k2_issue_team' ? netCalculatedByVisa['T'] : receivedK2ByVisa['T']) || 0).toLocaleString()}{' '}
                <span className="text-xs font-normal font-siemreap text-indigo-800">សន្លឹក</span>
              </div>
              <div className="text-[10px] text-indigo-700 mt-0.5">
                {(mode === 'receive_k1' || mode === 'k2_issue_team' ? totalNetCalculated : totalReceivedK2) > 0
                  ? ((((mode === 'receive_k1' || mode === 'k2_issue_team' ? netCalculatedByVisa['T'] : receivedK2ByVisa['T']) || 0) / (mode === 'receive_k1' || mode === 'k2_issue_team' ? totalNetCalculated : totalReceivedK2)) * 100).toFixed(1)
                  : 0}
                % នៃចំនួនសរុប
              </div>
            </div>

            <div className="bg-amber-50 border border-amber-200 rounded-xl p-3 shadow-2xs">
              <div className="text-[11px] font-semibold text-amber-800 flex items-center gap-1 font-siemreap">
                <Hash className="w-3.5 h-3.5 text-amber-600" />
                <span>ប្រភេទ E & K (ធម្មតា/ពិសេស)</span>
              </div>
              <div className="text-xl font-bold text-amber-950 font-times mt-1">
                {(
                  (mode === 'receive_k1' || mode === 'k2_issue_team'
                    ? (netCalculatedByVisa['E'] || 0) + (netCalculatedByVisa['K'] || 0)
                    : (receivedK2ByVisa['E'] || 0) + (receivedK2ByVisa['K'] || 0))
                ).toLocaleString()}{' '}
                <span className="text-xs font-normal font-siemreap text-amber-800">សន្លឹក</span>
              </div>
              <div className="text-[10px] text-amber-700 mt-0.5">
                E: {((mode === 'receive_k1' || mode === 'k2_issue_team' ? netCalculatedByVisa['E'] : receivedK2ByVisa['E']) || 0).toLocaleString()} | K: {((mode === 'receive_k1' || mode === 'k2_issue_team' ? netCalculatedByVisa['K'] : receivedK2ByVisa['K']) || 0).toLocaleString()} សន្លឹក
              </div>
            </div>

            <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-3 shadow-2xs">
              <div className="text-[11px] font-semibold text-emerald-800 flex items-center gap-1 font-siemreap">
                <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                <span>ប្រភេទមានចរន្តបើក</span>
              </div>
              <div className="text-xl font-bold text-emerald-950 font-times mt-1">
                {(Object.values(mode === 'receive_k1' || mode === 'k2_issue_team' ? netCalculatedByVisa : receivedK2ByVisa) as number[]).filter((v: number) => (Number(v) || 0) > 0).length} / 13{' '}
                <span className="text-xs font-normal font-siemreap text-emerald-800">ប្រភេទ</span>
              </div>
              <div className="text-[10px] text-emerald-700 mt-0.5">
                គ្មានចរន្តបើក: {(Object.values(mode === 'receive_k1' || mode === 'k2_issue_team' ? netCalculatedByVisa : receivedK2ByVisa) as number[]).filter((v: number) => (Number(v) || 0) === 0).length} ប្រភេទ
              </div>
            </div>
          </div>
        </div>

        {/* Tab Navigation */}
        <div className="bg-white border-b border-gray-200 px-5 pt-3 flex items-center justify-between gap-2">
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={() => setActiveTab('summary')}
              className={`px-4 py-2 text-xs sm:text-sm font-semibold rounded-t-xl transition-all cursor-pointer flex items-center gap-1.5 ${
                activeTab === 'summary'
                  ? 'bg-blue-600 text-white shadow-xs'
                  : 'text-gray-600 hover:text-gray-900 hover:bg-gray-100'
              }`}
            >
              <FileText className="w-4 h-4" />
              <span>តារាងសង្ខេបតាមប្រភេទ (១៣ ប្រភេទ)</span>
            </button>

            <button
              type="button"
              onClick={() => setActiveTab('records')}
              className={`px-4 py-2 text-xs sm:text-sm font-semibold rounded-t-xl transition-all cursor-pointer flex items-center gap-1.5 ${
                activeTab === 'records'
                  ? 'bg-blue-600 text-white shadow-xs'
                  : 'text-gray-600 hover:text-gray-900 hover:bg-gray-100'
              }`}
            >
              <ListFilter className="w-4 h-4" />
              <span>បញ្ជីកំណត់ត្រាប្រតិបត្តិការ ({receivedK2Records.length})</span>
            </button>

            <button
              type="button"
              onClick={() => setActiveTab('monthly')}
              className={`px-4 py-2 text-xs sm:text-sm font-semibold rounded-t-xl transition-all cursor-pointer flex items-center gap-1.5 ${
                activeTab === 'monthly'
                  ? 'bg-blue-600 text-white shadow-xs'
                  : 'text-gray-600 hover:text-gray-900 hover:bg-gray-100'
              }`}
            >
              <Calendar className="w-4 h-4" />
              <span>ម៉ាទ្រីសប្រចាំ ១២ខែ</span>
            </button>
          </div>
        </div>

        {/* Modal Body */}
        <div id="yearly-received-k2-checklist-container" className="p-5 overflow-y-auto grow space-y-4 bg-white">
          {/* TAB 1: SUMMARY TABLE */}
          {activeTab === 'summary' && (
            <div className="space-y-4">
              {/* Action Banner for 1-click correct values */}
              {(mode === 'receive_k1' || mode === 'k2_issue_team') && onApplyAllIssuedValues && (
                <div className="bg-amber-50 border border-amber-300 rounded-xl p-3.5 flex flex-wrap items-center justify-between gap-3 shadow-xs">
                  <div className="flex items-center gap-2">
                    <Sparkles className="w-5 h-5 text-amber-600 shrink-0" />
                    <div>
                      <div className="font-bold text-amber-950 text-xs sm:text-sm font-siemreap">
                        រូបមន្តគណនា៖ <span className="text-blue-900 font-semibold">បើកផ្តល់ជូនក្រុម = បើកផ្តល់តាមក្រុម ({totalGrossIssued.toLocaleString()}) - ផ្ទេរការប្រើប្រាស់ ({totalTransferDeduction.toLocaleString()}) = <span className="text-emerald-700 font-bold">{totalNetCalculated.toLocaleString()} សន្លឹក</span></span>
                      </div>
                      <div className="text-[11px] text-amber-800">
                        ចុចប៊ូតុងខាងស្តាំដើម្បីបញ្ចូល និងកែតម្រូវទិន្នន័យជាក់ស្តែងចូលក្នុងតារាងរបាយការណ៍ស្វ័យប្រវត្តិ
                      </div>
                    </div>
                  </div>

                  <button
                    type="button"
                    onClick={() => {
                      onApplyAllIssuedValues(netCalculatedByVisa);
                      setAppliedSuccessMsg(`✓ បានកែតម្រូវទិន្នន័យបើកផ្តល់ជូនក្រុមសរុប ${totalNetCalculated.toLocaleString()} សន្លឹក ចូលក្នុងតារាងរបាយការណ៍រួចរាល់!`);
                      setTimeout(() => setAppliedSuccessMsg(null), 4000);
                    }}
                    className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs px-4 py-2 rounded-lg shadow-sm flex items-center gap-1.5 transition cursor-pointer"
                  >
                    <CheckCircle2 className="w-4 h-4" />
                    <span>កែតម្រូវទិន្នន័យបើកផ្តល់ជូនក្រុម ({totalNetCalculated.toLocaleString()} សន្លឹក)</span>
                  </button>
                </div>
              )}

              {appliedSuccessMsg && (
                <div className="bg-emerald-100 border border-emerald-400 text-emerald-900 px-4 py-2.5 rounded-xl text-xs font-semibold flex items-center gap-2 animate-bounce">
                  <CheckCircle2 className="w-4 h-4 text-emerald-600" />
                  <span>{appliedSuccessMsg}</span>
                </div>
              )}

              <div className="overflow-x-auto rounded-xl border border-gray-200 shadow-xs">
                <table className="w-full text-xs font-siemreap text-left border-collapse">
                  <thead>
                    <tr className="bg-[#D9E1F2] text-gray-900 font-bold border-b border-gray-300">
                      <th className="border border-black px-3 py-2.5 text-center w-12">ល.រ</th>
                      <th className="border border-black px-3 py-2.5 text-center w-24">ប្រភេទទិដ្ឋាការ</th>
                      <th className="border border-black px-4 py-2.5">ឈ្មោះពេញទិដ្ឋាការ</th>
                      {mode === 'receive_k1' || mode === 'k2_issue_team' ? (
                        <>
                          <th className="border border-black px-3 py-2.5 text-right w-32 bg-blue-100/70">បើកផ្តល់តាមក្រុម</th>
                          <th className="border border-black px-3 py-2.5 text-right w-32 bg-rose-100/70 text-rose-900">ផ្ទេរការប្រើប្រាស់ (-)</th>
                          <th className="border border-black px-4 py-2.5 text-right w-36 bg-emerald-100/80 text-emerald-950 font-bold">បើកផ្តល់ជូនក្រុម (=)</th>
                        </>
                      ) : (
                        <th className="border border-black px-4 py-2.5 text-right w-40">ចំនួនបើកពីក២ (សន្លឹក)</th>
                      )}
                      <th className="border border-black px-3 py-2.5 text-center w-28">ភាគរយ (%)</th>
                      <th className="border border-black px-3 py-2.5 text-center w-32">ស្ថានភាព</th>
                      <th className="border border-black px-3 py-2.5 text-center w-24">សកម្មភាព</th>
                    </tr>
                  </thead>
                  <tbody>
                    {REPORT_VISA_TYPES.map((vt, index) => {
                      const grossCount = grossIssueByVisa[vt] || 0;
                      const transferCount = transferDeductionByVisa[vt] || 0;
                      const netCount = mode === 'receive_k1' || mode === 'k2_issue_team' ? (netCalculatedByVisa[vt] || 0) : (receivedK2ByVisa[vt] || 0);
                      const baseTotal = mode === 'receive_k1' || mode === 'k2_issue_team' ? totalNetCalculated : totalReceivedK2;
                      const pct = baseTotal > 0 ? (netCount / baseTotal) * 100 : 0;
                      const hasData = netCount > 0 || grossCount > 0;

                      return (
                        <tr
                          key={vt}
                          className={`border-b border-gray-200 hover:bg-blue-50/60 transition ${
                            index % 2 === 1 ? 'bg-gray-50/50' : 'bg-white'
                          }`}
                        >
                          <td className="border border-black px-3 py-2 text-center text-gray-600 font-medium">
                            {toKhmerNum(index + 1)}
                          </td>
                          <td className="border border-black px-3 py-2 text-center font-bold text-blue-900 font-times text-sm">
                            {vt}
                          </td>
                          <td className="border border-black px-4 py-2 font-medium text-gray-900">
                            {VISA_TYPE_NAMES[vt] || vt}
                          </td>
                          {mode === 'receive_k1' || mode === 'k2_issue_team' ? (
                            <>
                              <td className="border border-black px-3 py-2 text-right font-mono text-gray-800">
                                {grossCount.toLocaleString()}
                              </td>
                              <td className="border border-black px-3 py-2 text-right font-mono text-rose-700">
                                {transferCount > 0 ? `-${transferCount.toLocaleString()}` : '0'}
                              </td>
                              <td className="border border-black px-4 py-2 text-right font-bold text-emerald-950 font-times text-sm bg-emerald-50/50">
                                {netCount.toLocaleString()}
                              </td>
                            </>
                          ) : (
                            <td className="border border-black px-4 py-2 text-right font-bold text-black font-times text-sm">
                              {netCount.toLocaleString()}
                            </td>
                          )}
                          <td className="border border-black px-3 py-2 text-center font-semibold text-gray-700">
                            {pct > 0 ? (
                              <div className="flex items-center justify-center gap-1.5">
                                <div className="w-12 bg-gray-200 rounded-full h-1.5 overflow-hidden">
                                  <div
                                    className="bg-blue-600 h-full rounded-full"
                                    style={{ width: `${Math.min(pct, 100)}%` }}
                                  />
                                </div>
                                <span className="text-[11px]">{pct.toFixed(1)}%</span>
                              </div>
                            ) : (
                              <span className="text-gray-400">-</span>
                            )}
                          </td>
                          <td className="border border-black px-3 py-2 text-center">
                            {hasData ? (
                              <span className="inline-flex items-center gap-1 text-[10.5px] font-semibold text-emerald-800 bg-emerald-100 px-2 py-0.5 rounded-full">
                                <CheckCircle2 className="w-3 h-3 text-emerald-600" />
                                <span>មានការបើកផ្តល់</span>
                              </span>
                            ) : (
                              <span className="inline-flex items-center gap-1 text-[10.5px] font-medium text-gray-500 bg-gray-100 px-2 py-0.5 rounded-full">
                                <span>គ្មាន</span>
                              </span>
                            )}
                          </td>
                          <td className="border border-black px-3 py-2 text-center">
                            <button
                              type="button"
                              onClick={() => {
                                setSelectedVisaFilter(vt);
                                setActiveTab('records');
                              }}
                              className="text-[11px] font-semibold text-blue-700 hover:text-blue-900 hover:underline cursor-pointer"
                            >
                              មើលកំណត់ត្រា
                            </button>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                  <tfoot>
                    <tr className="bg-[#D9E1F2] font-bold text-gray-950 border-t-2 border-black">
                      <td colSpan={3} className="border border-black px-4 py-2.5 text-center font-bold font-siemreap">
                        សរុបរួម (១៣ ប្រភេទ)
                      </td>
                      {mode === 'receive_k1' || mode === 'k2_issue_team' ? (
                        <>
                          <td className="border border-black px-3 py-2.5 text-right font-bold text-blue-950 font-times text-sm">
                            {totalGrossIssued.toLocaleString()}
                          </td>
                          <td className="border border-black px-3 py-2.5 text-right font-bold text-rose-900 font-times text-sm">
                            {totalTransferDeduction > 0 ? `-${totalTransferDeduction.toLocaleString()}` : '0'}
                          </td>
                          <td className="border border-black px-4 py-2.5 text-right font-bold text-emerald-950 font-times text-base bg-emerald-100/60">
                            {totalNetCalculated.toLocaleString()}
                          </td>
                        </>
                      ) : (
                        <td className="border border-black px-4 py-2.5 text-right font-bold text-blue-950 font-times text-base">
                          {totalReceivedK2.toLocaleString()}
                        </td>
                      )}
                      <td className="border border-black px-3 py-2.5 text-center font-bold">100%</td>
                      <td className="border border-black px-3 py-2.5 text-center font-semibold text-emerald-900">
                        {(Object.values(mode === 'receive_k1' || mode === 'k2_issue_team' ? netCalculatedByVisa : receivedK2ByVisa) as number[]).filter((v: number) => (Number(v) || 0) > 0).length} / 13 ប្រភេទ
                      </td>
                      <td className="border border-black px-3 py-2.5 text-center">-</td>
                    </tr>
                  </tfoot>
                </table>
              </div>

              {/* Information Note */}
              <div className="bg-blue-50 border border-blue-200 rounded-xl p-4 flex items-start gap-3">
                <Info className="w-5 h-5 text-blue-700 shrink-0 mt-0.5" />
                <div className="text-xs text-blue-900 leading-relaxed font-siemreap">
                  <span className="font-bold">កំណត់សម្គាល់ការផ្ទៀងផ្ទាត់ទិដ្ឋាការ បើកពីក២ ៖</span>
                  <p className="mt-1">
                    ទិន្នន័យទិដ្ឋាការបើកពីក២ គឺជាចំនួនសរុបនៃសន្លឹកទិដ្ឋាការស្អិតដែលការិយាល័យទិដ្ឋាការចូល (ក២) បានបើកផ្តល់ចែកជូនទៅតាមបណ្តាក្រុមច្រកទ្វារអន្តរជាតិនានាក្នុងអំឡុងពេល ១២ខែ។ ទិន្នន័យនេះស្របគ្នា ១០០% ជាមួយតារាងសន្និធិ និងតារាងចលនាស្តុកផ្លូវការ។
                  </p>
                </div>
              </div>
            </div>
          )}

          {/* TAB 2: DETAILED RECORDS */}
          {activeTab === 'records' && (
            <div className="space-y-4">
              {/* Search and Filters */}
              <div className="bg-gray-50 border border-gray-200 rounded-xl p-3.5 flex flex-wrap items-center justify-between gap-3">
                <div className="flex flex-wrap items-center gap-3 grow">
                  {/* Search Input */}
                  <div className="relative min-w-[220px] grow sm:grow-0">
                    <Search className="w-4 h-4 text-gray-400 absolute left-3 top-1/2 -translate-y-1/2" />
                    <input
                      type="text"
                      placeholder="ស្វែងរកតាម ក្រុម, លេខស៊េរី, មន្ត្រី, ប័ណ្ណ..."
                      value={searchTerm}
                      onChange={(e) => setSearchTerm(e.target.value)}
                      className="w-full pl-9 pr-3 py-1.5 text-xs bg-white border border-gray-300 rounded-lg focus:outline-hidden focus:ring-2 focus:ring-blue-500 font-siemreap"
                    />
                    {searchTerm && (
                      <button
                        type="button"
                        onClick={() => setSearchTerm('')}
                        className="absolute right-2.5 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600 text-xs"
                      >
                        ✕
                      </button>
                    )}
                  </div>

                  {/* Visa Type Filter */}
                  <div className="flex items-center gap-1.5">
                    <span className="text-xs font-semibold text-gray-700 font-siemreap">ប្រភេទ ៖</span>
                    <select
                      value={selectedVisaFilter}
                      onChange={(e) => setSelectedVisaFilter(e.target.value)}
                      className="text-xs bg-white border border-gray-300 rounded-lg px-2.5 py-1.5 focus:outline-hidden focus:ring-2 focus:ring-blue-500 font-siemreap cursor-pointer font-medium"
                    >
                      <option value="ALL">គ្រប់ប្រភេទ ({receivedK2Records.length})</option>
                      {REPORT_VISA_TYPES.map((vt) => (
                        <option key={vt} value={vt}>
                          {vt} - {VISA_TYPE_NAMES[vt] || vt} ({(receivedK2ByVisa[vt] || 0).toLocaleString()})
                        </option>
                      ))}
                    </select>
                  </div>

                  {/* Month Filter */}
                  <div className="flex items-center gap-1.5">
                    <span className="text-xs font-semibold text-gray-700 font-siemreap">ខែ ៖</span>
                    <select
                      value={selectedMonthFilter}
                      onChange={(e) => setSelectedMonthFilter(e.target.value)}
                      className="text-xs bg-white border border-gray-300 rounded-lg px-2.5 py-1.5 focus:outline-hidden focus:ring-2 focus:ring-blue-500 font-siemreap cursor-pointer font-medium"
                    >
                      <option value="ALL">គ្រប់ខែ (១២ ខែ)</option>
                      {monthColumnsInfo.map((m, idx) => (
                        <option key={idx} value={String(idx)}>
                          {m.name} {toKhmerNum(m.year)}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>

                <div className="text-xs text-gray-600 font-siemreap">
                  រកឃើញ <span className="font-bold text-blue-900">{filteredRecords.length}</span> កំណត់ត្រា
                </div>
              </div>

              {/* Records Table */}
              <div className="overflow-x-auto rounded-xl border border-gray-200 shadow-xs max-h-[500px] overflow-y-auto">
                <table className="w-full text-xs font-siemreap text-left border-collapse">
                  <thead className="sticky top-0 bg-[#D9E1F2] z-10">
                    <tr className="text-gray-900 font-bold border-b border-gray-300">
                      <th className="border border-black px-2.5 py-2 text-center w-10">ល.រ</th>
                      <th className="border border-black px-2.5 py-2 text-center w-24">កាលបរិច្ឆេទ</th>
                      <th className="border border-black px-2 py-2 text-center w-14">ប្រភេទ</th>
                      <th className="border border-black px-3 py-2 text-right w-24">ចំនួន (សន្លឹក)</th>
                      <th className="border border-black px-3 py-2 text-center w-36">លេខស៊េរី (ពី - ដល់)</th>
                      <th className="border border-black px-3 py-2 w-36">ក្រុម/ច្រកទ្វារទទួល</th>
                      <th className="border border-black px-3 py-2 w-32">អ្នកប្រគល់/ទទួល</th>
                      <th className="border border-black px-2.5 py-2 text-center w-24">ប័ណ្ណបើក</th>
                      <th className="border border-black px-3 py-2">កំណត់សម្គាល់</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredRecords.length > 0 ? (
                      filteredRecords.map((rec, index) => {
                        const vt = normalizeVisaType(rec.visaType);
                        const rawQty = Number((rec as any).quantity || (rec as any).quantityBundles || rec.totalSheets || ((rec as any).count ? (rec as any).count * 50 : 0)) || 0;
                        const dIso = normalizeDateToISO(rec.date);
                        const op = (rec.operationType || '').toLowerCase();
                        const isTransfer =
                          op === 'transferteam' ||
                          op === 'transferuseteam' ||
                          op === 'transfer' ||
                          op.includes('transfer') ||
                          op.includes('ផ្ទេរ') ||
                          (rec.sourceFrom && rec.sourceFrom.includes('ផ្ទេរ')) ||
                          ((rec as any).notes && (rec as any).notes.includes('ផ្ទេរ'));

                        return (
                          <tr
                            key={rec.id || index}
                            className={`border-b border-gray-200 hover:bg-blue-50/70 transition ${
                              isTransfer ? 'bg-amber-50/60' : index % 2 === 1 ? 'bg-gray-50/40' : 'bg-white'
                            }`}
                          >
                            <td className="border border-black px-2.5 py-2 text-center text-gray-500 font-medium">
                              {toKhmerNum(index + 1)}
                            </td>
                            <td className="border border-black px-2.5 py-2 text-center font-medium text-gray-800 whitespace-nowrap">
                              {dIso || rec.date}
                            </td>
                            <td className="border border-black px-2 py-2 text-center font-bold text-blue-900 font-times">
                              <span className="bg-blue-100 text-blue-950 px-1.5 py-0.5 rounded text-[11px] font-bold">
                                {vt}
                              </span>
                            </td>
                            <td className={`border border-black px-3 py-2 text-right font-bold font-times ${isTransfer ? 'text-amber-700' : 'text-black'}`}>
                              {isTransfer ? `-${rawQty.toLocaleString()}` : rawQty.toLocaleString()}
                            </td>
                            <td className="border border-black px-3 py-2 text-center font-mono text-[11px] text-gray-700 whitespace-nowrap">
                              {rec.startSerial && rec.endSerial ? (
                                <span>
                                  {rec.startSerial} - {rec.endSerial}
                                </span>
                              ) : (
                                <span className="text-gray-400">-</span>
                              )}
                            </td>
                            <td className="border border-black px-3 py-2 font-medium text-gray-900">
                              {rec.teamName || (rec as any).sourceFrom || (rec as any).visaTeamRobokName || (
                                <span className="text-gray-400">មិនបានបញ្ជាក់</span>
                              )}
                              {isTransfer && (
                                <span className="ml-1.5 inline-block text-[10px] px-1.5 py-0.2 bg-amber-100 text-amber-800 border border-amber-300 rounded">
                                  ផ្ទេរការប្រើប្រាស់
                                </span>
                              )}
                            </td>
                            <td className="border border-black px-3 py-2 text-gray-700 text-[11px]">
                              {rec.responsibleOfficer || (rec as any).receiver || (rec as any).giver || '-'}
                            </td>
                            <td className="border border-black px-2.5 py-2 text-center font-mono text-[11px] text-gray-600">
                              {rec.documentNumber || (rec as any).voucherNo || '-'}
                            </td>
                            <td className="border border-black px-3 py-2 text-gray-600 text-[11px]">
                              {rec.notes || (rec as any).memo || (isTransfer ? 'ផ្ទេរការប្រើប្រាស់' : '-')}
                            </td>
                          </tr>
                        );
                      })
                    ) : (
                      <tr>
                        <td colSpan={9} className="border border-black py-8 text-center text-gray-500 font-siemreap">
                          មិនមានទិន្នន័យកំណត់ត្រាបើកពីក២ ស្របតាមលក្ខខណ្ឌស្វែងរកនេះឡើយ
                        </td>
                      </tr>
                    )}
                  </tbody>
                  {filteredRecords.length > 0 && (
                    <tfoot className="sticky bottom-0 bg-[#D9E1F2] z-10">
                      <tr className="font-bold text-gray-950 border-t-2 border-black">
                        <td colSpan={3} className="border border-black px-3 py-2 text-center font-bold font-siemreap">
                          សរុបកំណត់ត្រាដែលបានជ្រើសរើស
                        </td>
                        <td className="border border-black px-3 py-2 text-right font-bold text-blue-950 font-times">
                          {filteredRecords
                            .reduce(
                              (acc, r) =>
                                acc +
                                (Number((r as any).quantity || (r as any).quantityBundles || r.totalSheets || ((r as any).count ? (r as any).count * 50 : 0)) || 0),
                              0
                            )
                            .toLocaleString()}
                        </td>
                        <td colSpan={5} className="border border-black px-3 py-2 text-left font-siemreap text-gray-700">
                          ({filteredRecords.length} កំណត់ត្រា)
                        </td>
                      </tr>
                    </tfoot>
                  )}
                </table>
              </div>
            </div>
          )}

          {/* TAB 3: 12-MONTH MATRIX */}
          {activeTab === 'monthly' && (
            <div className="space-y-4">
              <div className="overflow-x-auto rounded-xl border border-gray-200 shadow-xs">
                <table className="w-full text-xs font-siemreap text-center border-collapse">
                  <thead>
                    <tr className="bg-[#D9E1F2] text-gray-900 font-bold border-b border-gray-300">
                      <th className="border border-black px-2 py-2 text-center w-10">ល.រ</th>
                      <th className="border border-black px-2 py-2 text-center w-16">ប្រភេទ</th>
                      {monthColumnsInfo.map((m, idx) => (
                        <th key={idx} className="border border-black px-1.5 py-2 text-center min-w-[55px]">
                          <div className="whitespace-nowrap font-bold text-[10.5px]">{m.name}</div>
                          <div className="text-[9px] text-gray-700 font-normal">{toKhmerNum(m.year)}</div>
                        </th>
                      ))}
                      <th className="border border-black px-2.5 py-2 text-right font-bold bg-blue-100 min-w-[70px]">
                        សរុប
                      </th>
                    </tr>
                  </thead>
                  <tbody>
                    {REPORT_VISA_TYPES.map((vt, index) => {
                      const mVals = monthlyBreakdownByVisa[vt] || Array(12).fill(0);
                      const rowTotal = mVals.reduce((a, b) => a + b, 0);

                      return (
                        <tr
                          key={vt}
                          className={`border-b border-gray-200 hover:bg-blue-50/60 transition ${
                            index % 2 === 1 ? 'bg-gray-50/50' : 'bg-white'
                          }`}
                        >
                          <td className="border border-black px-2 py-1.5 text-center text-gray-500 font-medium">
                            {toKhmerNum(index + 1)}
                          </td>
                          <td className="border border-black px-2 py-1.5 font-bold text-blue-900 font-times text-sm">
                            {vt}
                          </td>
                          {mVals.map((val, mIdx) => (
                            <td
                              key={mIdx}
                              className={`border border-black px-1.5 py-1.5 font-times text-xs ${
                                val > 0 ? 'font-bold text-black' : 'text-gray-400'
                              }`}
                            >
                              {val > 0 ? val.toLocaleString() : '-'}
                            </td>
                          ))}
                          <td className="border border-black px-2.5 py-1.5 text-right font-bold text-blue-950 font-times text-sm bg-blue-50/60">
                            {rowTotal.toLocaleString()}
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                  <tfoot>
                    <tr className="bg-[#D9E1F2] font-bold text-gray-950 border-t-2 border-black">
                      <td colSpan={2} className="border border-black px-2 py-2 text-center font-bold font-siemreap">
                        សរុបប្រចាំខែ
                      </td>
                      {Array(12)
                        .fill(0)
                        .map((_, mIdx) => {
                          const mSum = REPORT_VISA_TYPES.reduce(
                            (s, vt) => s + (monthlyBreakdownByVisa[vt]?.[mIdx] || 0),
                            0
                          );
                          return (
                            <td
                              key={mIdx}
                              className="border border-black px-1.5 py-2 text-center font-bold font-times text-xs"
                            >
                              {mSum.toLocaleString()}
                            </td>
                          );
                        })}
                      <td className="border border-black px-2.5 py-2 text-right font-bold text-blue-950 font-times text-sm bg-blue-200">
                        {totalReceivedK2.toLocaleString()}
                      </td>
                    </tr>
                  </tfoot>
                </table>
              </div>
            </div>
          )}
        </div>

        {/* Modal Footer */}
        <div className="bg-gray-50 border-t border-gray-200 px-5 py-3 flex items-center justify-between">
          <div className="text-xs text-gray-600 font-siemreap flex items-center gap-2">
            <span className="font-semibold text-gray-800">សរុបបើកពីក២ ៖</span>
            <span className="font-bold text-blue-900 font-times text-sm">
              {totalReceivedK2.toLocaleString()}
            </span>
            <span>សន្លឹក</span>
          </div>

          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={onClose}
              className="bg-gray-200 hover:bg-gray-300 text-gray-800 text-xs font-semibold px-4 py-2 rounded-xl transition cursor-pointer font-siemreap"
            >
              បិទ
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};
