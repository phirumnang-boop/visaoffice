import React, { useState, useMemo } from 'react';
import * as XLSX from 'xlsx';
import {
  X,
  FileSpreadsheet,
  Printer,
  CheckCircle2,
  ListFilter,
  Search,
  AlertTriangle,
  FileText,
  Sparkles,
  Info,
  Layers,
  Calendar,
  Building,
  Hash,
} from 'lucide-react';
import { StockRecord } from '../types';
import { REPORT_VISA_TYPES, isRecordForTeam, isRecordForRecipientTeam, isTransferTeamRecord } from './YearlyTeamDistributionReport';
import { toKhmerNum, KHMER_MONTHS } from '../utils/khmerCalendar';
import { normalizeDateToISO, normalizeVisaType, normalizeTeamName } from '../utils/teamNormalization';
import { isCeaRecord } from '../utils/teamStockCalculation';
import { printA4Document } from '../utils/printHelper';

export interface MonthColInfo {
  monthNum: number;
  year: number;
  name: string;
}

export interface YearlyDamagedMissingChecklistModalProps {
  isOpen: boolean;
  onClose: () => void;
  startYear: number;
  startMonth: number;
  monthColumnsInfo: MonthColInfo[];
  tableData: Record<string, any>;
  stockRecords: StockRecord[];
  targetTeamName?: string;
  mode?: 'distribution' | 'usage' | 'receive_k1';
  initialVisaFilter?: string;
  onUpdateDamagedMissing?: (visaType: string, newValue: number) => void;
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

// Format operation type to Khmer display name
const formatOperationTypeKhmer = (op: string, src?: string): string => {
  const low = (op || '').toLowerCase();
  const lowSrc = (src || '').toLowerCase();
  if (low === 'transferteam' || low === 'transferuseteam' || low.includes('transfer') || low.includes('ផ្ទេរ') || lowSrc.includes('ផ្ទេរ')) {
    return 'ផ្ទេរការប្រើប្រាស់';
  }
  if (low === 'testprintk2' || low === 'testk2' || low.includes('សាក') || lowSrc.includes('សាក')) {
    return 'ទិដ្ឋាការសាកក២';
  }
  if (low === 'damagedk2' || low === 'damagedoffice' || low.includes('ខូចក២') || low.includes('មិនបានការ')) {
    return 'ទិដ្ឋាការខូចក២';
  }
  if (low === 'missingteam' || low === 'missing' || low.includes('ខ្វះ')) {
    return 'ទិដ្ឋាការខ្វះក្រុម';
  }
  if (low === 'damagedteam' || low.includes('ខូចក្រុម')) {
    return 'ទិដ្ឋាការខូចក្រុម';
  }
  if (low === 'damaged' || low.includes('ខូច')) {
    return 'ទិដ្ឋាការខូច';
  }
  if (low === 'returnteam' || low.includes('បង្វិល')) {
    return 'ទិដ្ឋាការបង្វិលពីក្រុម';
  }
  return op || 'ខូច/ខ្វះ';
};

export const YearlyDamagedMissingChecklistModal: React.FC<YearlyDamagedMissingChecklistModalProps> = ({
  isOpen,
  onClose,
  startYear,
  startMonth,
  monthColumnsInfo,
  tableData,
  stockRecords = [],
  targetTeamName,
  mode = 'usage',
  initialVisaFilter = 'ALL',
  onUpdateDamagedMissing,
}) => {
  const [activeTab, setActiveTab] = useState<'summary' | 'records'>('summary');
  const [searchTerm, setSearchTerm] = useState<string>('');
  const [selectedVisaFilter, setSelectedVisaFilter] = useState<string>(initialVisaFilter || 'ALL');

  // Sync initialVisaFilter when modal opens
  React.useEffect(() => {
    if (isOpen) {
      setSelectedVisaFilter(initialVisaFilter || 'ALL');
    }
  }, [isOpen, initialVisaFilter]);

  // Date range
  const startDateStr = `${startYear}-${String(startMonth).padStart(2, '0')}-01`;
  const endYear = monthColumnsInfo[11]?.year || startYear;
  const endMonth = monthColumnsInfo[11]?.monthNum || startMonth;
  const daysInEndMonth = new Date(endYear, endMonth, 0).getDate();
  const endDateStr = `${endYear}-${String(endMonth).padStart(2, '0')}-${String(daysInEndMonth).padStart(2, '0')}`;

  // Find all damaged/missing/test records in stockRecords for the 12-month period
  const damagedRecords = useMemo(() => {
    if (!stockRecords || stockRecords.length === 0) return [];

    return stockRecords.filter((rec) => {
      if (rec.stockType && rec.stockType !== 'sticker') return false;
      if (isCeaRecord(rec)) return false;

      const dIso = normalizeDateToISO(rec.date);
      if (!dIso || dIso < startDateStr || dIso > endDateStr) return false;

      const op = (rec.operationType || '').toLowerCase();
      const src = (rec.sourceFrom || '').toLowerCase();

      if (mode === 'receive_k1') {
        const isDamagedK2 =
          op === 'damagedk2' ||
          op === 'damaged' ||
          op === 'damagedoffice' ||
          op === 'voidk2' ||
          op === 'voidoffice' ||
          op === 'invalidk2' ||
          op === 'invalid' ||
          op.includes('ខូចក២') ||
          op.includes('មិនបានការក២') ||
          op.includes('ខូចក') ||
          op.includes('ក២ខូច') ||
          op.includes('មិនបានការ') ||
          src.includes('ខូចក២') ||
          src.includes('មិនបានការ');

        const isTestK2 =
          op === 'testprintk2' ||
          op === 'testk2' ||
          op === 'samplek2' ||
          op === 'testoffice' ||
          op === 'sampleoffice' ||
          op === 'testprint' ||
          op === 'sample' ||
          op === 'test' ||
          op.includes('សាកក') ||
          op.includes('សាកល្បង') ||
          op.includes('សាក') ||
          src.includes('សាក') ||
          src.includes('បោះពុម្ពសាកល្បង');

        return isDamagedK2 || isTestK2;
      }

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

      if (isTransfer) {
        // Only include in "ផ្ទេរ/ខ្វះ/ខូច" for the sender team who transferred out
        return Boolean(targetTeamName && isSender && !isRecipient);
      }

      const isDamagedK2 =
        op === 'damagedk2' ||
        op === 'damaged' ||
        op === 'damagedoffice' ||
        op === 'voidk2' ||
        op === 'voidoffice' ||
        op === 'missingoffice' ||
        op.includes('ខូចក២') ||
        op.includes('មិនបានការក២') ||
        op.includes('ខូចក') ||
        op.includes('ក២ខូច');

      const isDamaged =
        !isDamagedK2 &&
        (op === 'damagedteam' ||
          op === 'voidteam' ||
          op === 'missingteam' ||
          op === 'teamdamaged' ||
          op.includes('ខូចក្រុម') ||
          op.includes('ខ្វះក្រុម') ||
          op.includes('ខូចតាមក្រុម') ||
          op.includes('ខ្វះតាមក្រុម') ||
          ((op.includes('ខូច') || op.includes('ខ្វះ')) &&
            !op.includes('ក២') &&
            !op.includes('ក១') &&
            !op.includes('ការិយាល័យ') &&
            (rec.sourceFrom?.includes('ក្រុម') || (rec as any).visaTeamRobokName)));

      return isDamaged;
    });
  }, [stockRecords, startDateStr, endDateStr, mode, targetTeamName]);

  // Compute breakdown of actual Damaged K2 vs Test K2 per visa type
  const actualStatsByVisa = useMemo(() => {
    const res: Record<
      string,
      { damagedK2: number; testK2: number; totalActual: number; records: StockRecord[] }
    > = {};

    REPORT_VISA_TYPES.forEach((vt) => {
      res[vt] = { damagedK2: 0, testK2: 0, totalActual: 0, records: [] };
    });

    damagedRecords.forEach((rec) => {
      const vt = normalizeVisaType(rec.visaType || '');
      if (!res[vt]) return;
      const qty = Number(rec.quantityBundles || rec.totalSheets || (rec as any).quantity || 0);
      const op = (rec.operationType || '').toLowerCase();
      const src = (rec.sourceFrom || '').toLowerCase();
      const isTest =
        op === 'testprintk2' ||
        op === 'testk2' ||
        op === 'samplek2' ||
        op.includes('សាក') ||
        src.includes('សាក');

      if (isTest) {
        res[vt].testK2 += qty;
      } else {
        res[vt].damagedK2 += qty;
      }
      res[vt].totalActual += qty;
      res[vt].records.push(rec);
    });

    return res;
  }, [damagedRecords]);

  // Compute total damaged per visa type from current tableData
  const tableValuesByVisa = useMemo(() => {
    const map: Record<string, number> = {};
    REPORT_VISA_TYPES.forEach((vt) => {
      const row = tableData[vt];
      map[vt] = Number(row?.damagedK2 ?? row?.damagedMissing) || 0;
    });
    return map;
  }, [tableData]);

  // Grand Totals
  const totalTableDamaged = useMemo(() => {
    return (Object.values(tableValuesByVisa) as number[]).reduce(
      (acc: number, val: number) => acc + (Number(val) || 0),
      0
    );
  }, [tableValuesByVisa]);

  const totalActualAll = useMemo(() => {
    return REPORT_VISA_TYPES.reduce((sum, vt) => sum + (actualStatsByVisa[vt]?.totalActual || 0), 0);
  }, [actualStatsByVisa]);

  const totalDamagedK2All = useMemo(() => {
    return REPORT_VISA_TYPES.reduce((sum, vt) => sum + (actualStatsByVisa[vt]?.damagedK2 || 0), 0);
  }, [actualStatsByVisa]);

  const totalTestK2All = useMemo(() => {
    return REPORT_VISA_TYPES.reduce((sum, vt) => sum + (actualStatsByVisa[vt]?.testK2 || 0), 0);
  }, [actualStatsByVisa]);

  // Filtered Records for Records Tab
  const filteredRecords = useMemo(() => {
    return damagedRecords.filter((rec) => {
      const vt = normalizeVisaType(rec.visaType);
      if (selectedVisaFilter !== 'ALL' && vt !== selectedVisaFilter) return false;

      if (!searchTerm) return true;
      const term = searchTerm.toLowerCase();
      const team = (rec.teamName || '').toLowerCase();
      const officer = (rec.responsibleOfficer || '').toLowerCase();
      const sn = `${rec.startSerial || ''} ${rec.endSerial || ''}`.toLowerCase();
      const notes = (rec.notes || (rec as any).note || '').toLowerCase();
      const op = (rec.operationType || '').toLowerCase();

      return (
        vt.toLowerCase().includes(term) ||
        team.includes(term) ||
        officer.includes(term) ||
        sn.includes(term) ||
        notes.includes(term) ||
        op.includes(term)
      );
    });
  }, [damagedRecords, selectedVisaFilter, searchTerm]);

  // Sync a single visa type value to tableData
  const handleSyncSingle = (vt: string) => {
    if (!onUpdateDamagedMissing) return;
    const actual = actualStatsByVisa[vt]?.totalActual || 0;
    onUpdateDamagedMissing(vt, actual);
  };

  // Sync all visa types values to tableData
  const handleSyncAll = () => {
    if (!onUpdateDamagedMissing) return;
    if (!window.confirm('តើលោកអ្នកចង់អាប់ដេតទិន្នន័យជាក់ស្តែងទាំងអស់ (ខូចក២ + សាកក២) ទៅក្នុងតារាងមែនទេ?')) return;
    REPORT_VISA_TYPES.forEach((vt) => {
      const actual = actualStatsByVisa[vt]?.totalActual || 0;
      onUpdateDamagedMissing(vt, actual);
    });
  };

  // Export to Excel
  const handleExportExcel = () => {
    const wb = XLSX.utils.book_new();

    const titleRows = [
      ['ព្រះរាជាណាចក្រកម្ពុជា'],
      ['ជាតិ សាសនា ព្រះមហាក្សត្រ'],
      ['***'],
      ['តារាងផ្ទៀងផ្ទាត់បញ្ជីទិដ្ឋាការ ខ្វះ & ខូច តាមប្រភេទនីមួយៗ'],
      [`ដោយគិតចាប់ពីថ្ងៃទី០១ ខែ${KHMER_MONTHS[startMonth - 1] || startMonth} ឆ្នាំ${toKhmerNum(startYear)} ដល់ថ្ងៃទី${toKhmerNum(daysInEndMonth)} ខែ${monthColumnsInfo[11]?.name} ឆ្នាំ${toKhmerNum(endYear)}`],
      [],
      ['ល.រ', 'ប្រភេទវីសា', 'ឈ្មោះពេញ', 'ចំនួនខ្វះ&ខូច (សន្លឹក)', 'ភាគរយ (%)', 'ស្ថានភាព'],
    ];

    const dataRows = REPORT_VISA_TYPES.map((vt, index) => {
      const stats = actualStatsByVisa[vt] || { damagedK2: 0, testK2: 0, totalActual: 0 };
      const count = stats.totalActual;
      const pct = totalActualAll > 0 ? ((count / totalActualAll) * 100).toFixed(2) + '%' : '0%';
      return [
        index + 1,
        vt,
        VISA_TYPE_NAMES[vt] || vt,
        stats.damagedK2,
        stats.testK2,
        count,
        pct,
        count > 0 ? 'មានការខូច/សាកល្បង' : 'គ្មាន',
      ];
    });

    const totalRow = [
      '',
      'សរុប',
      'សរុបទូទាំងគ្រប់ប្រភេទ',
      totalDamagedK2All,
      totalTestK2All,
      totalActualAll,
      '100%',
      '',
    ];

    const ws = XLSX.utils.aoa_to_sheet([...titleRows, ...dataRows, totalRow]);
    XLSX.utils.book_append_sheet(wb, ws, 'បញ្ជីទិដ្ឋាការខ្វះខូច');
    XLSX.writeFile(wb, `តារាងផ្ទៀងផ្ទាត់ទិដ្ឋាការខ្វះខូច_${startYear}_${endYear}.xlsx`);
  };

  const handlePrint = () => {
    printA4Document('yearly-damaged-missing-checklist-container', {
      orientation: 'landscape',
      documentTitle: mode === 'receive_k1'
        ? 'ផ្ទៀងផ្ទាត់បញ្ជីទិដ្ឋាការ_មិនបានការ_សាកល្បងក២'
        : 'ផ្ទៀងផ្ទាត់បញ្ជីទិដ្ឋាការ_ខ្វះ_ខូច_ប្រចាំឆ្នាំ',
    });
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 flex items-center justify-center bg-black/60 backdrop-blur-xs p-2 sm:p-4 overflow-y-auto animate-fade-in">
      <div className="bg-white w-full max-w-5xl rounded-2xl shadow-2xl border border-gray-200 flex flex-col max-h-[92vh] overflow-hidden my-auto">
        {/* Modal Header */}
        <div className="bg-linear-to-r from-red-800 via-rose-900 to-red-950 text-white px-5 py-3.5 flex items-center justify-between shadow-md">
          <div className="flex items-center gap-3">
            <div className="p-2 bg-white/10 rounded-xl backdrop-blur-xs border border-white/20">
              <AlertTriangle className="w-5 h-5 text-amber-300" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h2 className="text-base sm:text-lg font-bold font-siemreap tracking-wide">
                  {mode === 'receive_k1'
                    ? 'ផ្ទៀងផ្ទាត់បញ្ជីទិដ្ឋាការ មិនបានការ & សាកល្បងក២ (១២ខែ)'
                    : 'ផ្ទៀងផ្ទាត់បញ្ជីទិដ្ឋាការ ខ្វះ & ខូច ប្រចាំឆ្នាំ (១២ខែ)'}
                </h2>
                <span className="text-[11px] font-medium bg-red-700/80 border border-red-500/40 text-amber-200 px-2 py-0.5 rounded-full">
                  {mode === 'receive_k1'
                    ? 'តារាងបើកពីក១ (ខូចក២ + សាកក២)'
                    : mode === 'usage'
                    ? 'តារាងប្រើប្រាស់តាមក្រុម'
                    : 'តារាងបើកផ្តល់តាមក្រុម'}
                </span>
              </div>
              <p className="text-xs text-red-200 font-siemreap mt-0.5">
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
            {mode === 'receive_k1' ? (
              <>
                <div className="bg-rose-50 border border-rose-200 rounded-xl p-3 shadow-2xs">
                  <div className="text-[11px] font-semibold text-rose-800 flex items-center gap-1 font-siemreap">
                    <AlertTriangle className="w-3.5 h-3.5 text-rose-600" />
                    <span>សរុបមិនបានការ & សាកក២</span>
                  </div>
                  <div className="text-2xl font-bold text-rose-950 font-times mt-1">
                    {totalActualAll.toLocaleString()}{' '}
                    <span className="text-xs font-normal font-siemreap text-rose-800">សន្លឹក</span>
                  </div>
                  <div className="text-[10px] text-rose-700 mt-0.5">
                    ទិន្នន័យជាក់ស្តែងពីប្រតិបត្តិការ
                  </div>
                </div>

                <div className="bg-red-50 border border-red-200 rounded-xl p-3 shadow-2xs">
                  <div className="text-[11px] font-semibold text-red-800 flex items-center gap-1 font-siemreap">
                    <Hash className="w-3.5 h-3.5 text-red-600" />
                    <span>ទិដ្ឋាការខូចក២ (Damaged)</span>
                  </div>
                  <div className="text-xl font-bold text-red-950 font-times mt-1">
                    {totalDamagedK2All.toLocaleString()}{' '}
                    <span className="text-xs font-normal font-siemreap text-red-800">សន្លឹក</span>
                  </div>
                  <div className="text-[10px] text-red-700 mt-0.5">
                    {totalActualAll > 0 ? ((totalDamagedK2All / totalActualAll) * 100).toFixed(1) : 0}% នៃចំនួនសរុប
                  </div>
                </div>

                <div className="bg-amber-50 border border-amber-200 rounded-xl p-3 shadow-2xs">
                  <div className="text-[11px] font-semibold text-amber-800 flex items-center gap-1 font-siemreap">
                    <Sparkles className="w-3.5 h-3.5 text-amber-600" />
                    <span>ទិដ្ឋាការសាកក២ (Test Print)</span>
                  </div>
                  <div className="text-xl font-bold text-amber-950 font-times mt-1">
                    {totalTestK2All.toLocaleString()}{' '}
                    <span className="text-xs font-normal font-siemreap text-amber-800">សន្លឹក</span>
                  </div>
                  <div className="text-[10px] text-amber-700 mt-0.5">
                    {totalActualAll > 0 ? ((totalTestK2All / totalActualAll) * 100).toFixed(1) : 0}% នៃចំនួនសរុប
                  </div>
                </div>

                <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-3 shadow-2xs">
                  <div className="text-[11px] font-semibold text-emerald-800 flex items-center gap-1 font-siemreap">
                    <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                    <span>តម្លៃក្នុងតារាងបច្ចុប្បន្ន</span>
                  </div>
                  <div className="text-xl font-bold text-emerald-950 font-times mt-1">
                    {totalTableDamaged.toLocaleString()}{' '}
                    <span className="text-xs font-normal font-siemreap text-emerald-800">សន្លឹក</span>
                  </div>
                  <div className="text-[10px] text-emerald-700 mt-0.5">
                    {totalTableDamaged === totalActualAll ? (
                      <span className="text-emerald-700 font-semibold">✓ ផ្ទៀងផ្ទាត់ត្រូវគ្នា ១០០%</span>
                    ) : (
                      <span className="text-rose-700 font-semibold">
                        ខុសគ្នា {Math.abs(totalTableDamaged - totalActualAll)} សន្លឹក
                      </span>
                    )}
                  </div>
                </div>
              </>
            ) : (
              <>
                <div className="bg-red-50 border border-red-200 rounded-xl p-3 shadow-2xs">
                  <div className="text-[11px] font-semibold text-red-800 flex items-center gap-1 font-siemreap">
                    <AlertTriangle className="w-3.5 h-3.5 text-red-600" />
                    <span>សរុប ខ្វះ & ខូច</span>
                  </div>
                  <div className="text-2xl font-bold text-red-950 font-times mt-1">
                    {totalTableDamaged.toLocaleString()}{' '}
                    <span className="text-xs font-normal font-siemreap text-red-800">សន្លឹក</span>
                  </div>
                  <div className="text-[10px] text-red-700 mt-0.5">
                    សរុប ១២ខែ គ្រប់ប្រភេទវីសា
                  </div>
                </div>

                <div className="bg-blue-50 border border-blue-200 rounded-xl p-3 shadow-2xs">
                  <div className="text-[11px] font-semibold text-blue-800 flex items-center gap-1 font-siemreap">
                    <Hash className="w-3.5 h-3.5 text-blue-600" />
                    <span>ប្រភេទ T (ទេសចរណ៍)</span>
                  </div>
                  <div className="text-xl font-bold text-blue-950 font-times mt-1">
                    {(tableValuesByVisa['T'] || 0).toLocaleString()}{' '}
                    <span className="text-xs font-normal font-siemreap text-blue-800">សន្លឹក</span>
                  </div>
                  <div className="text-[10px] text-blue-700 mt-0.5">
                    {totalTableDamaged > 0
                      ? (((tableValuesByVisa['T'] || 0) / totalTableDamaged) * 100).toFixed(1)
                      : 0}
                    % នៃចំនួនខូចសរុប
                  </div>
                </div>

                <div className="bg-amber-50 border border-amber-200 rounded-xl p-3 shadow-2xs">
                  <div className="text-[11px] font-semibold text-amber-800 flex items-center gap-1 font-siemreap">
                    <Hash className="w-3.5 h-3.5 text-amber-600" />
                    <span>ប្រភេទ K & E (ពិសេស/ធម្មតា)</span>
                  </div>
                  <div className="text-xl font-bold text-amber-950 font-times mt-1">
                    {((tableValuesByVisa['K'] || 0) + (tableValuesByVisa['E'] || 0)).toLocaleString()}{' '}
                    <span className="text-xs font-normal font-siemreap text-amber-800">សន្លឹក</span>
                  </div>
                  <div className="text-[10px] text-amber-700 mt-0.5">
                    K: {tableValuesByVisa['K'] || 0} | E: {tableValuesByVisa['E'] || 0} សន្លឹក
                  </div>
                </div>

                <div className="bg-emerald-50 border border-emerald-200 rounded-xl p-3 shadow-2xs">
                  <div className="text-[11px] font-semibold text-emerald-800 flex items-center gap-1 font-siemreap">
                    <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                    <span>ប្រភេទមានការខូច</span>
                  </div>
                  <div className="text-xl font-bold text-emerald-950 font-times mt-1">
                    {(Object.values(tableValuesByVisa) as number[]).filter((v: number) => (Number(v) || 0) > 0).length} / 13{' '}
                    <span className="text-xs font-normal font-siemreap text-emerald-800">ប្រភេទ</span>
                  </div>
                  <div className="text-[10px] text-emerald-700 mt-0.5">
                    គ្មានខូច: {(Object.values(tableValuesByVisa) as number[]).filter((v: number) => (Number(v) || 0) === 0).length} ប្រភេទ
                  </div>
                </div>
              </>
            )}
          </div>
        </div>

        {/* Tab Navigation */}
        <div className="bg-white border-b border-gray-200 px-5 pt-3 flex flex-wrap items-center justify-between gap-2">
          <div className="flex items-center gap-2">
            <button
              type="button"
              onClick={() => setActiveTab('summary')}
              className={`px-3.5 py-2 rounded-t-xl text-xs font-medium flex items-center gap-1.5 transition cursor-pointer border-b-2 ${
                activeTab === 'summary'
                  ? 'border-red-600 text-red-800 font-bold bg-red-50/50'
                  : 'border-transparent text-gray-600 hover:text-gray-900 hover:bg-gray-50'
              }`}
            >
              <Layers className="w-3.5 h-3.5" />
              <span>តារាងសង្ខេបផ្ទៀងផ្ទាត់ (Summary Matrix)</span>
            </button>

            <button
              type="button"
              onClick={() => setActiveTab('records')}
              className={`px-3.5 py-2 rounded-t-xl text-xs font-medium flex items-center gap-1.5 transition cursor-pointer border-b-2 ${
                activeTab === 'records'
                  ? 'border-red-600 text-red-800 font-bold bg-red-50/50'
                  : 'border-transparent text-gray-600 hover:text-gray-900 hover:bg-gray-50'
              }`}
            >
              <FileText className="w-3.5 h-3.5" />
              <span>បញ្ជីកំណត់ត្រាប្រតិបត្តិការជាក់ស្តែង ({damagedRecords.length})</span>
            </button>
          </div>

          {mode === 'receive_k1' && onUpdateDamagedMissing && (
            <button
              type="button"
              onClick={handleSyncAll}
              className="bg-red-700 hover:bg-red-800 text-white text-xs font-semibold px-3 py-1.5 rounded-lg flex items-center gap-1.5 transition shadow-xs cursor-pointer font-siemreap mb-2"
              title="អាប់ដេតទិន្នន័យជាក់ស្តែងទាំងអស់ទៅក្នុងតារាង"
            >
              <Sparkles className="w-3.5 h-3.5 text-amber-300" />
              <span>អាប់ដេតទិន្នន័យជាក់ស្តែងទាំងអស់ចូលក្នុងតារាង</span>
            </button>
          )}
        </div>

        {/* Body Content */}
        <div id="yearly-damaged-missing-checklist-container" className="flex-1 p-5 overflow-y-auto space-y-4 bg-white">
          {activeTab === 'summary' && (
            <div className="space-y-4">
              {/* Summary Table */}
              <div className="border border-black rounded-xl overflow-hidden shadow-xs bg-white">
                <table className="w-full border-collapse border border-black text-black text-center text-xs">
                  <thead>
                    <tr className="bg-[#D9E1F2] font-bold border-b border-black text-black font-siemreap">
                      <th className="border border-black px-2 py-2 w-10 text-center">ល.រ</th>
                      <th className="border border-black px-2 py-2 w-16 text-center">ប្រភេទ</th>
                      <th className="border border-black px-3 py-2 text-left">ឈ្មោះប្រភេទពេញ</th>
                      {mode === 'receive_k1' ? (
                        <>
                          <th className="border border-black px-2 py-2 w-28 text-right bg-red-50">ខូចក២ (សន្លឹក)</th>
                          <th className="border border-black px-2 py-2 w-28 text-right bg-amber-50">សាកក២ (សន្លឹក)</th>
                          <th className="border border-black px-2 py-2 w-28 text-right bg-rose-100">សរុបជាក់ស្តែង</th>
                          <th className="border border-black px-2 py-2 w-28 text-right bg-blue-50">តម្លៃក្នុងតារាង</th>
                          <th className="border border-black px-2 py-2 w-28 text-center">ស្ថានភាព</th>
                          {onUpdateDamagedMissing && (
                            <th className="border border-black px-2 py-2 w-28 text-center">សកម្មភាព</th>
                          )}
                        </>
                      ) : (
                        <>
                          <th className="border border-black px-3 py-2 w-32 text-right">ចំនួនខ្វះ&ខូច (សន្លឹក)</th>
                          <th className="border border-black px-3 py-2 w-24 text-right">ភាគរយ</th>
                          <th className="border border-black px-3 py-2 w-48 text-center">របារភាគរយ</th>
                          <th className="border border-black px-3 py-2 w-28 text-center">ស្ថានភាព</th>
                        </>
                      )}
                    </tr>
                  </thead>
                  <tbody>
                    {REPORT_VISA_TYPES.map((vt, idx) => {
                      const stat = actualStatsByVisa[vt] || { damagedK2: 0, testK2: 0, totalActual: 0 };
                      const tableVal = tableValuesByVisa[vt] || 0;
                      const isMatch = tableVal === stat.totalActual;
                      const pct = totalTableDamaged > 0 ? (tableVal / totalTableDamaged) * 100 : 0;

                      return (
                        <tr
                          key={vt}
                          className={`hover:bg-red-50/40 transition-colors ${
                            idx % 2 === 1 ? 'bg-gray-50/40' : 'bg-white'
                          }`}
                        >
                          <td className="border border-black px-2 py-1.5 text-center font-times">
                            {idx + 1}
                          </td>
                          <td className="border border-black px-2 py-1.5 text-center font-bold font-times text-blue-900">
                            {vt}
                          </td>
                          <td className="border border-black px-3 py-1.5 text-left font-siemreap text-gray-800">
                            {VISA_TYPE_NAMES[vt] || vt}
                          </td>

                          {mode === 'receive_k1' ? (
                            <>
                              <td className="border border-black px-2 py-1.5 text-right font-times text-red-700 bg-red-50/30">
                                {stat.damagedK2.toLocaleString()}
                              </td>
                              <td className="border border-black px-2 py-1.5 text-right font-times text-amber-800 bg-amber-50/30">
                                {stat.testK2.toLocaleString()}
                              </td>
                              <td className="border border-black px-2 py-1.5 text-right font-bold font-times text-rose-900 bg-rose-50/50">
                                {stat.totalActual.toLocaleString()}
                              </td>
                              <td className="border border-black px-2 py-1.5 text-right font-bold font-times text-blue-900 bg-blue-50/30">
                                {tableVal.toLocaleString()}
                              </td>
                              <td className="border border-black px-2 py-1.5 text-center font-siemreap">
                                {isMatch ? (
                                  <span className="inline-flex items-center gap-1 text-[11px] font-semibold text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded-md">
                                    <CheckCircle2 className="w-3 h-3 text-emerald-600" />
                                    <span>ត្រឹមត្រូវ</span>
                                  </span>
                                ) : (
                                  <span className="inline-flex items-center gap-1 text-[11px] font-semibold text-rose-700 bg-rose-100 px-2 py-0.5 rounded-md">
                                    <AlertTriangle className="w-3 h-3 text-rose-600" />
                                    <span>ខុសគ្នា {tableVal - stat.totalActual > 0 ? `+${tableVal - stat.totalActual}` : tableVal - stat.totalActual}</span>
                                  </span>
                                )}
                              </td>
                              {onUpdateDamagedMissing && (
                                <td className="border border-black px-1.5 py-1 text-center font-siemreap">
                                  <button
                                    type="button"
                                    onClick={() => handleSyncSingle(vt)}
                                    className="px-2 py-0.5 bg-blue-600 hover:bg-blue-700 text-white rounded text-[10px] font-medium transition cursor-pointer"
                                    title="ដាក់តម្លៃជាក់ស្តែងចូលក្នុងតារាង"
                                  >
                                    ដាក់តម្លៃជាក់ស្តែង
                                  </button>
                                </td>
                              )}
                            </>
                          ) : (
                            <>
                              <td className="border border-black px-3 py-1.5 text-right font-bold font-times text-red-700 bg-red-50/20">
                                {tableVal.toLocaleString()}
                              </td>
                              <td className="border border-black px-3 py-1.5 text-right font-times text-gray-700">
                                {pct.toFixed(1)}%
                              </td>
                              <td className="border border-black px-3 py-1.5 text-center align-middle">
                                <div className="w-full bg-gray-200 rounded-full h-2 overflow-hidden">
                                  <div
                                    className="bg-red-600 h-2 rounded-full transition-all duration-500"
                                    style={{ width: `${Math.min(pct * 1.5, 100)}%` }}
                                  />
                                </div>
                              </td>
                              <td className="border border-black px-2 py-1.5 text-center font-siemreap">
                                {tableVal > 0 ? (
                                  <span className="inline-flex items-center gap-1 text-[11px] font-semibold text-red-700 bg-red-100 px-2 py-0.5 rounded-md">
                                    <AlertTriangle className="w-3 h-3" />
                                    <span>ខូច {tableVal}</span>
                                  </span>
                                ) : (
                                  <span className="inline-flex items-center gap-1 text-[11px] font-semibold text-emerald-700 bg-emerald-50 px-2 py-0.5 rounded-md">
                                    <CheckCircle2 className="w-3 h-3 text-emerald-600" />
                                    <span>គ្មាន</span>
                                  </span>
                                )}
                              </td>
                            </>
                          )}
                        </tr>
                      );
                    })}

                    {/* Total Row */}
                    <tr className="bg-[#D9E1F2] font-bold border-t-2 border-black text-black">
                      <td colSpan={3} className="border border-black px-3 py-2 text-center font-siemreap font-bold">
                        សរុបរួម (Grand Total)
                      </td>
                      {mode === 'receive_k1' ? (
                        <>
                          <td className="border border-black px-2 py-2 text-right font-bold font-times text-red-900 bg-red-100">
                            {totalDamagedK2All.toLocaleString()}
                          </td>
                          <td className="border border-black px-2 py-2 text-right font-bold font-times text-amber-900 bg-amber-100">
                            {totalTestK2All.toLocaleString()}
                          </td>
                          <td className="border border-black px-2 py-2 text-right font-bold font-times text-rose-950 bg-rose-200">
                            {totalActualAll.toLocaleString()}
                          </td>
                          <td className="border border-black px-2 py-2 text-right font-bold font-times text-blue-950 bg-blue-100">
                            {totalTableDamaged.toLocaleString()}
                          </td>
                          <td colSpan={onUpdateDamagedMissing ? 2 : 1} className="border border-black px-2 py-2 text-center font-siemreap text-xs">
                            {totalTableDamaged === totalActualAll ? (
                              <span className="text-emerald-800 font-bold">✓ ទិន្នន័យស៊ីសង្វាក់គ្នា ១០០%</span>
                            ) : (
                              <span className="text-rose-800 font-bold">
                                ខុសគ្នា {Math.abs(totalTableDamaged - totalActualAll)} សន្លឹក
                              </span>
                            )}
                          </td>
                        </>
                      ) : (
                        <>
                          <td className="border border-black px-3 py-2 text-right font-bold font-times text-red-900 bg-[#B4C6E7]">
                            {totalTableDamaged.toLocaleString()}
                          </td>
                          <td className="border border-black px-3 py-2 text-right font-bold font-times">
                            100.0%
                          </td>
                          <td colSpan={2} className="border border-black px-3 py-2 text-center font-siemreap text-xs text-gray-700">
                            បានផ្ទៀងផ្ទាត់ត្រឹមត្រូវតាមទិន្នន័យជាក់ស្តែង
                          </td>
                        </>
                      )}
                    </tr>
                  </tbody>
                </table>
              </div>

              {/* Information Notice */}
              <div className="bg-amber-50 border border-amber-200 rounded-xl p-3 text-xs text-amber-900 flex items-start gap-2.5 font-siemreap">
                <Info className="w-4 h-4 text-amber-700 shrink-0 mt-0.5" />
                <div>
                  <p className="font-semibold text-amber-950">
                    រូបមន្តគណនាទិដ្ឋាការ មិនបានការ សាកល្បងក២ ៖
                  </p>
                  <p className="mt-0.5 text-amber-800 leading-relaxed font-times text-[12px]">
                    <span className="font-bold font-siemreap">មិនបានការ សាកល្បងក២ = </span>
                    <span>ទិដ្ឋាការខូចក២ (Damaged K2) + ទិដ្ឋាការសាកក២ (Test Prints K2)</span>
                  </p>
                  <p className="mt-1 text-gray-600 font-siemreap text-[11px]">
                    លោកអ្នកអាចពិនិត្យផ្ទៀងផ្ទាត់ និងចុច «ដាក់តម្លៃជាក់ស្តែង» ឬ «អាប់ដេតទិន្នន័យជាក់ស្តែងទាំងអស់ចូលក្នុងតារាង» ដើម្បីឱ្យតម្លៃក្នុងតារាងដើរស្របគ្នាជាមួយកំណត់ត្រាប្រតិបត្តិការជាក់ស្តែង។
                  </p>
                </div>
              </div>
            </div>
          )}

          {activeTab === 'records' && (
            <div className="space-y-3">
              {/* Search and Filters */}
              <div className="flex flex-wrap items-center justify-between gap-2 bg-gray-50 p-2.5 rounded-xl border border-gray-200">
                <div className="flex items-center gap-2 flex-1 min-w-[240px]">
                  <div className="relative flex-1">
                    <Search className="w-4 h-4 absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" />
                    <input
                      type="text"
                      value={searchTerm}
                      onChange={(e) => setSearchTerm(e.target.value)}
                      placeholder="ស្វែងរកតាមក្រុម, លេខសន្លឹក, មន្ត្រីទទួលខុសត្រូវ, កំណត់សម្គាល់..."
                      className="w-full bg-white border border-gray-300 rounded-lg pl-9 pr-3 py-1.5 text-xs focus:outline-none focus:ring-2 focus:ring-red-500 font-siemreap"
                    />
                  </div>
                  {searchTerm && (
                    <button
                      type="button"
                      onClick={() => setSearchTerm('')}
                      className="text-xs text-gray-500 hover:text-gray-700 underline"
                    >
                      សម្អាត
                    </button>
                  )}
                </div>

                <div className="flex items-center gap-2">
                  <span className="text-xs font-semibold text-gray-600 font-siemreap flex items-center gap-1">
                    <ListFilter className="w-3.5 h-3.5" />
                    <span>ប្រភេទ ៖</span>
                  </span>
                  <select
                    value={selectedVisaFilter}
                    onChange={(e) => setSelectedVisaFilter(e.target.value)}
                    className="bg-white border border-gray-300 rounded-lg px-2.5 py-1.5 text-xs focus:outline-none font-semibold text-gray-700"
                  >
                    <option value="ALL">គ្រប់ប្រភេទ ({REPORT_VISA_TYPES.length})</option>
                    {REPORT_VISA_TYPES.map((vt) => (
                      <option key={vt} value={vt}>
                        {vt} - {VISA_TYPE_NAMES[vt] || vt} ({mode === 'receive_k1' ? actualStatsByVisa[vt]?.totalActual || 0 : tableValuesByVisa[vt] || 0})
                      </option>
                    ))}
                  </select>
                </div>
              </div>

              {/* Records Table */}
              <div className="border border-black rounded-xl overflow-hidden shadow-xs bg-white">
                <table className="w-full border-collapse border border-black text-black text-center text-xs">
                  <thead>
                    <tr className="bg-[#D9E1F2] font-bold border-b border-black text-black font-siemreap">
                      <th className="border border-black px-2 py-2 w-10 text-center">ល.រ</th>
                      <th className="border border-black px-2 py-2 w-24 text-center">កាលបរិច្ឆេទ</th>
                      <th className="border border-black px-2 py-2 w-16 text-center">ប្រភេទ</th>
                      <th className="border border-black px-2 py-2 text-left">ក្រុម / ប្រភព</th>
                      <th className="border border-black px-2 py-2 w-20 text-right">ចំនួន (សន្លឹក)</th>
                      <th className="border border-black px-2 py-2 w-36 text-center">ពីលេខ ~ ដល់លេខ</th>
                      <th className="border border-black px-2 py-2 w-32 text-center">ប្រភេទប្រតិបត្តិការ</th>
                      <th className="border border-black px-2 py-2 text-left">មន្ត្រីទទួលខុសត្រូវ / កំណត់សម្គាល់</th>
                    </tr>
                  </thead>
                  <tbody>
                    {filteredRecords.length > 0 ? (
                      filteredRecords.map((rec, idx) => {
                        const opLow = (rec.operationType || '').toLowerCase();
                        const srcLow = (rec.sourceFrom || '').toLowerCase();
                        const isTest = opLow.includes('test') || opLow.includes('sample') || opLow.includes('សាក') || srcLow.includes('សាក');

                        return (
                          <tr
                            key={rec.id || idx}
                            className={`hover:bg-red-50/40 transition-colors ${
                              idx % 2 === 1 ? 'bg-gray-50/40' : 'bg-white'
                            }`}
                          >
                            <td className="border border-black px-2 py-1.5 text-center font-times">
                              {idx + 1}
                            </td>
                            <td className="border border-black px-2 py-1.5 text-center font-times">
                              {rec.date}
                            </td>
                            <td className="border border-black px-2 py-1.5 text-center font-bold font-times text-blue-900">
                              {rec.visaType}
                            </td>
                            <td className="border border-black px-2 py-1.5 text-left font-siemreap">
                              {rec.teamName || rec.sourceFrom || 'ទូទៅ'}
                            </td>
                            <td className="border border-black px-2 py-1.5 text-right font-bold font-times text-red-700">
                              {Number(rec.quantityBundles || rec.totalSheets || (rec as any).quantity || 0).toLocaleString()}
                            </td>
                            <td className="border border-black px-2 py-1.5 text-center font-mono text-[11px]">
                              {rec.startSerial && rec.endSerial
                                ? `${rec.startSerial} ~ ${rec.endSerial}`
                                : '-'}
                            </td>
                            <td className="border border-black px-2 py-1.5 text-center font-siemreap text-xs">
                              <span
                                className={`px-2 py-0.5 rounded text-[11px] font-semibold border ${
                                  isTest
                                    ? 'bg-amber-100 text-amber-900 border-amber-300'
                                    : 'bg-rose-100 text-rose-800 border-rose-300'
                                }`}
                              >
                                {formatOperationTypeKhmer(rec.operationType, rec.sourceFrom)}
                              </span>
                            </td>
                            <td className="border border-black px-2 py-1.5 text-left font-siemreap text-xs text-gray-700">
                              {rec.responsibleOfficer || rec.notes || (rec as any).note || '-'}
                            </td>
                          </tr>
                        );
                      })
                    ) : (
                      <tr>
                        <td
                          colSpan={8}
                          className="border border-black py-8 text-center text-gray-500 font-siemreap"
                        >
                          <div className="flex flex-col items-center justify-center gap-1.5">
                            <CheckCircle2 className="w-7 h-7 text-emerald-500" />
                            <p className="font-semibold text-gray-700">
                              ទិន្នន័យខ្វះ&ខូចត្រូវបានផ្ទៀងផ្ទាត់រួចរាល់
                            </p>
                            <p className="text-[11px] text-gray-400">
                              គ្មានកំណត់ត្រាបុគ្គលបន្ថែមក្នុងប្រព័ន្ធសម្រាប់ចន្លោះកាលបរិច្ឆេទនេះទេ
                            </p>
                          </div>
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>

              {/* Diagnostics / Troubleshooting Guide Box */}
              <div className="bg-slate-50 border border-slate-300 rounded-xl p-3.5 text-xs text-slate-800 space-y-1.5 font-siemreap">
                <div className="flex items-center gap-1.5 font-bold text-slate-900">
                  <Info className="w-4 h-4 text-blue-600" />
                  <span>ឧបករណ៍វិភាគ និងផ្ទៀងផ្ទាត់ទិន្នន័យ (Diagnostic Details) ៖</span>
                </div>
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-2 pt-1 font-times text-[11.5px]">
                  <div className="bg-white p-2 rounded-lg border border-slate-200">
                    <span className="text-gray-500 font-siemreap">ចន្លោះកាលបរិច្ឆេទ ៖ </span>
                    <span className="font-bold text-slate-900">{startDateStr} ដល់ {endDateStr}</span>
                  </div>
                  <div className="bg-white p-2 rounded-lg border border-slate-200">
                    <span className="text-gray-500 font-siemreap">កំណត់ត្រាដែលបានរាប់ ៖ </span>
                    <span className="font-bold text-slate-900">{damagedRecords.length} ប្រតិបត្តិការ</span>
                  </div>
                  <div className="bg-white p-2 rounded-lg border border-slate-200">
                    <span className="text-gray-500 font-siemreap">សរុបសន្លឹកជាក់ស្តែង ៖ </span>
                    <span className="font-bold text-rose-700">{totalActualAll} សន្លឹក</span>
                  </div>
                </div>
              </div>
            </div>
          )}
        </div>

        {/* Modal Footer */}
        <div className="bg-gray-100 border-t border-gray-300 px-5 py-3 flex items-center justify-between">
          <div className="text-xs text-gray-600 font-siemreap flex items-center gap-2">
            <span className="font-semibold text-gray-800">
              {mode === 'receive_k1' ? 'សរុបមិនបានការ & សាកក២ ៖' : 'សរុបខ្វះ & ខូច ៖'}
            </span>
            <span className="font-bold text-red-700 font-times text-sm">
              {(mode === 'receive_k1' ? totalActualAll : totalTableDamaged).toLocaleString()}
            </span>
            <span>សន្លឹក</span>
          </div>

          <button
            type="button"
            onClick={onClose}
            className="px-5 py-1.5 bg-gray-800 hover:bg-gray-900 text-white text-xs font-semibold rounded-xl transition cursor-pointer font-siemreap"
          >
            បិទផ្ទាំង
          </button>
        </div>
      </div>
    </div>
  );
};
