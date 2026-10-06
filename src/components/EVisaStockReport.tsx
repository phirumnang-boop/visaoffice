import React, { useState, useEffect, useRef } from 'react';
import { CategoriesState, Officer, StockRecord } from '../types';
import {
  Download,
  Printer,
  Calendar,
  FileText,
  ArrowLeft,
  Check,
  X,
  Sparkles,
  Layers,
  FileCheck,
  FolderDown,
  CheckCircle2,
} from 'lucide-react';
import { exportElementToPdf } from '../utils/pdfExportHelper';
import { printA4Document } from '../utils/printHelper';
import { CustomDatePicker } from './CustomDatePicker';
import { getKhmerLunarDate, getKhmerSolarParts, parseDateInput, toKhmerNum } from '../utils/khmerCalendar';
import { normalizeDateToISO, normalizeTeamName } from '../utils/teamNormalization';
import {
  TacteingLine,
  TacteingType,
  getSavedTacteingSettings,
  saveTacteingSettings,
  TacteingControlSelector,
} from './TacteingLink';

export interface EVisaStockReportProps {
  stockRecords: StockRecord[];
  categories?: CategoriesState;
  officers?: Officer[];
  onClose?: () => void;
  isSingleTeam?: boolean;
  onSwitchReport?: (mode: 'report' | 'robokReport' | 'singleRobokReport') => void;
  currentReportMode?: 'report' | 'robokReport' | 'singleRobokReport';
  hideSwitcher?: boolean;
  assignedTeam?: string;
  lockTeamSelection?: boolean;
}

const RIGHT_SIGNATURE_OPTIONS = [
  'នាយផ្នែក',
  'ជ.នាយផ្នែក',
  'អ្នកធ្វើតារាង',
  'ប្រធានក្រុមផ្តល់ទិដ្ឋាការ',
  'ជ.ប្រធានក្រុមផ្តល់ទិដ្ឋាការ',
  'ប្រធានក្រុមផ្តល់ទិដ្ឋាការស្តីទី',
];

const SINGLE_TEAM_RIGHT_SIGNATURE_OPTIONS = [
  'ប្រធានក្រុមផ្តល់ទិដ្ឋាការ',
  'ជ.ប្រធានក្រុមផ្តល់ទិដ្ឋាការ',
  'ប្រធានក្រុមផ្តល់ទិដ្ឋាការស្តីទី',
  'នាយផ្នែក',
  'ជ.នាយផ្នែក',
  'អ្នកធ្វើតារាង',
];

function formatReportDateLine(dateStr: string): string {
  const parts = getKhmerSolarParts(dateStr);
  return `ថ្ងៃទី${parts.khmerDay} ខែ${parts.khmerMonth} ឆ្នាំ២០${parts.khmerYear.slice(-2) || parts.khmerYear}`;
}

// Universal high-definition A4 PDF capture engine matching RobokTotalStockWorkReport
async function captureElementToPdf(elementId: string, fileName: string): Promise<string> {
  const element = document.getElementById(elementId) || document.getElementById(`${elementId}-offscreen`);
  if (!element) throw new Error(`Element with id '${elementId}' not found`);

  return await exportElementToPdf(element, fileName, {
    pixelRatio: 3,
    fitSinglePage: true,
  });
}

// Navigation Tab Switcher Bar for the 3 E-Visa Stock Reports
const ReportTypeSwitchBar: React.FC<{
  currentMode: 'report' | 'robokReport' | 'singleRobokReport';
  onSwitchMode?: (mode: 'report' | 'robokReport' | 'singleRobokReport') => void;
  onClose?: () => void;
}> = ({ currentMode, onSwitchMode, onClose }) => {
  if (!onSwitchMode && !onClose) return null;

  return (
    <div className="bg-gradient-to-r from-slate-900 via-blue-950 to-slate-900 p-2 rounded-lg border border-blue-800/40 shadow-sm flex flex-wrap items-center justify-between gap-2 text-xs mb-3 print:hidden">
      <div className="flex items-center gap-1.5 text-amber-300 font-bold px-2 py-1">
        <FolderDown className="w-4 h-4 text-amber-400 shrink-0" />
        <span className="hidden sm:inline">ជ្រើសរើសរបាយការណ៍:</span>
      </div>

      <div className="flex flex-wrap items-center gap-1.5 flex-1 justify-center sm:justify-start">
        {onSwitchMode && (
          <>
            <button
              type="button"
              onClick={() => onSwitchMode('singleRobokReport')}
              className={`px-3 py-1.5 rounded-md text-xs font-bold transition flex items-center gap-1.5 cursor-pointer ${
                currentMode === 'singleRobokReport'
                  ? 'bg-[#C6A15B] text-slate-950 shadow-xs ring-1 ring-amber-300'
                  : 'bg-white/10 text-white hover:bg-white/20'
              }`}
            >
              <FileCheck className="w-3.5 h-3.5" />
              <span>១. របកក្រដាសបង្ហាញតែ១ក្រុម (Single Team)</span>
            </button>

            <button
              type="button"
              onClick={() => onSwitchMode('robokReport')}
              className={`px-3 py-1.5 rounded-md text-xs font-bold transition flex items-center gap-1.5 cursor-pointer ${
                currentMode === 'robokReport'
                  ? 'bg-[#C6A15B] text-slate-950 shadow-xs ring-1 ring-amber-300'
                  : 'bg-white/10 text-white hover:bg-white/20'
              }`}
            >
              <Layers className="w-3.5 h-3.5" />
              <span>២. របកក្រដាសអនុម័តតាមបណ្តាក្រុម (All Teams)</span>
            </button>

            <button
              type="button"
              onClick={() => onSwitchMode('report')}
              className={`px-3 py-1.5 rounded-md text-xs font-bold transition flex items-center gap-1.5 cursor-pointer ${
                currentMode === 'report'
                  ? 'bg-[#C6A15B] text-slate-950 shadow-xs ring-1 ring-amber-300'
                  : 'bg-white/10 text-white hover:bg-white/20'
              }`}
            >
              <Calendar className="w-3.5 h-3.5" />
              <span>៣. របាយការណ៍ក្រដាសអនុម័ត (Date-based)</span>
            </button>
          </>
        )}
      </div>

      {onClose && (
        <button
          type="button"
          onClick={onClose}
          className="px-3 py-1.5 rounded-md text-xs font-bold bg-white/15 hover:bg-white/25 text-white transition flex items-center gap-1 cursor-pointer ml-auto"
        >
          <ArrowLeft className="w-3.5 h-3.5" />
          <span>ត្រឡប់ទៅទម្រង់/ទិន្នន័យ</span>
        </button>
      )}
    </div>
  );
};

// 1. Dedicated Report Component for "របាយការណ៍ក្រដាសអនុម័ត" (Date-based Report)
export const EVisaStockReport: React.FC<EVisaStockReportProps> = ({
  stockRecords,
  categories,
  onClose,
  onSwitchReport,
  currentReportMode = 'report',
  hideSwitcher = false,
}) => {
  const today = new Date().toISOString().split('T')[0];
  const now = new Date();
  const firstDayOfMonth = new Date(now.getFullYear(), now.getMonth(), 1)
    .toISOString()
    .split('T')[0];

  const [startDate, setStartDate] = useState<string>(firstDayOfMonth);
  const [endDate, setEndDate] = useState<string>(today);
  const [reportDate, setReportDate] = useState<string>(today);
  const [isDownloading, setIsDownloading] = useState<boolean>(false);
  const [isBatchDownloading, setIsBatchDownloading] = useState<boolean>(false);
  const [batchStatus, setBatchStatus] = useState<string | null>(null);
  const [downloadSuccess, setDownloadSuccess] = useState<boolean>(false);
  const [downloadedBlobUrls, setDownloadedBlobUrls] = useState<{ label: string; url: string }[]>([]);

  const [showLeftSignature, setShowLeftSignature] = useState<boolean>(() => {
    try {
      const saved = localStorage.getItem('report_show_left_signature');
      return saved !== null ? saved === 'true' : true;
    } catch {
      return true;
    }
  });
  const [leftSignatureTitle, setLeftSignatureTitle] = useState<string>('នាយរងការិយាល័យទទួលបន្ទុក');
  const [rightSignatureTitle, setRightSignatureTitle] = useState<string>('នាយផ្នែក');

  const savedTacteing = getSavedTacteingSettings();
  const [showTacteingControls, setShowTacteingControls] = useState<boolean>(false);
  const [tacteingType, setTacteingType] = useState<TacteingType>(savedTacteing.type);
  const [tacteingCustomImage, setTacteingCustomImage] = useState<string | null>(savedTacteing.customImage);

  const reportRef = useRef<HTMLDivElement>(null);

  const startParts = getKhmerSolarParts(startDate);
  const endParts = getKhmerSolarParts(endDate);

  const sDateObj = parseDateInput(startDate);
  const eDateObj = parseDateInput(endDate);
  const startYearNum = sDateObj.getFullYear();
  const startMonthNum = sDateObj.getMonth();
  const endYearNum = eDateObj.getFullYear();
  const endMonthNum = eDateObj.getMonth();
  const totalMonths = (endYearNum - startYearNum) * 12 + (endMonthNum - startMonthNum) + 1;

  let reportPeriodLabel = `ប្រចាំខែ ${startParts.khmerMonth} ឆ្នាំ${startParts.khmerYear}`;
  if (totalMonths > 1) {
    const monthNumStr = toKhmerNum(String(totalMonths).padStart(2, '0'));
    reportPeriodLabel = `ប្រចាំរយៈពេល ${monthNumStr}ខែ`;
  }

  const startD = new Date(startDate);
  const dayBeforeStart = new Date(startD);
  dayBeforeStart.setDate(dayBeforeStart.getDate() - 1);
  const dayBeforeParts = getKhmerSolarParts(dayBeforeStart);

  const openingStockDateLabel = `${dayBeforeParts.khmerDay}-${dayBeforeParts.khmerMonth}-${dayBeforeParts.khmerYear}`;
  const endingStockDateLabel = `${endParts.khmerDay}-${endParts.khmerMonth}-${endParts.khmerYear}`;

  const reportLunar = getKhmerLunarDate(reportDate);

  const normStartDate = normalizeDateToISO(startDate);
  const normEndDate = normalizeDateToISO(endDate);

  const evisaRecords = stockRecords.filter((r) => r.stockType === 'evisa' || !r.stockType);

  const openingK1 = evisaRecords
    .filter(
      (r) =>
        (r.operationType === 'openK1' ||
          r.operationType === 'oldStockK2' ||
          r.operationType === 'oldStock' ||
          r.operationType === 'initial' ||
          r.sourceFrom?.includes('ស្តុកចាស់') ||
          r.sourceFrom?.includes('សន្និធិដើម')) &&
        normalizeDateToISO(r.date) < normStartDate
    )
    .reduce((sum, r) => sum + (r.quantityBundles || 0), 0);

  const openingIssue = evisaRecords
    .filter((r) => r.operationType === 'issueTeam' && normalizeDateToISO(r.date) < normStartDate)
    .reduce((sum, r) => sum + (r.quantityBundles || 0), 0);

  const officeOpeningStock = Math.max(0, openingK1 - openingIssue);

  const periodK1Records = evisaRecords.filter(
    (r) => {
      const rd = normalizeDateToISO(r.date);
      return r.operationType === 'openK1' && (!normStartDate || rd >= normStartDate) && (!normEndDate || rd <= normEndDate);
    }
  );

  const periodIssueRecords = evisaRecords.filter(
    (r) => {
      const rd = normalizeDateToISO(r.date);
      return r.operationType === 'issueTeam' && (!normStartDate || rd >= normStartDate) && (!normEndDate || rd <= normEndDate);
    }
  );

  const officeInward = periodK1Records.reduce((sum, r) => sum + (r.quantityBundles || 0), 0);
  const officeUsed = periodIssueRecords.reduce((sum, r) => sum + (r.quantityBundles || 0), 0);
  const officeEndingStock = officeOpeningStock + officeInward - officeUsed;

  const totalOpening = officeOpeningStock;
  const totalInward = officeInward;
  const totalUsed = officeUsed;
  const totalEnding = officeEndingStock;

  // Single PDF Download
  const handleDownloadPdf = async () => {
    try {
      setIsDownloading(true);
      const cleanStart = startDate.replace(/[/\\?%*:|"<>]/g, '-').trim();
      const cleanEnd = endDate.replace(/[/\\?%*:|"<>]/g, '-').trim();
      const fileName = `របាយការណ៍ក្រដាសអនុម័តទិដ្ឋាការអេឡិចត្រូនិក(${cleanStart}_ដល់_${cleanEnd}).pdf`;

      const blobUrl = await captureElementToPdf('official-pdf-report-document', fileName);

      setDownloadedBlobUrls([{ label: 'របាយការណ៍ក្រដាសអនុម័ត', url: blobUrl }]);
      setDownloadSuccess(true);
    } catch (err) {
      console.error('Error downloading PDF report:', err);
      alert('មានបញ្ហាក្នុងការទាញយក PDF។ សូមព្យាយាមម្តងទៀត។');
    } finally {
      setIsDownloading(false);
    }
  };

  // Batch Download All 3 PDFs at once
  const handleBatchDownloadAllThree = async () => {
    try {
      setIsBatchDownloading(true);
      setDownloadSuccess(false);

      const cleanStart = startDate.replace(/[/\\?%*:|"<>]/g, '-').trim();
      const cleanEnd = endDate.replace(/[/\\?%*:|"<>]/g, '-').trim();

      // 1. Single Team Report PDF
      setBatchStatus('កំពុងបង្កើត PDF របាយការណ៍ទី ១/៣ (របកក្រដាសបង្ហាញតែ១ក្រុម)...');
      const url1 = await captureElementToPdf(
        'official-pdf-single-team-document',
        `១_របកក្រដាសអនុម័ត_បង្ហាញតែ១ក្រុម(${cleanStart}_ដល់_${cleanEnd}).pdf`
      );

      // 2. All Teams Report PDF
      setBatchStatus('កំពុងបង្កើត PDF របាយការណ៍ទី ២/៣ (របកក្រដាសអនុម័តតាមបណ្តាក្រុម)...');
      const url2 = await captureElementToPdf(
        'official-pdf-robok-document',
        `២_របកក្រដាសអនុម័ត_តាមបណ្តាក្រុម(${cleanStart}_ដល់_${cleanEnd}).pdf`
      );

      // 3. Date-based Report PDF
      setBatchStatus('កំពុងបង្កើត PDF របាយការណ៍ទី ៣/៣ (របាយការណ៍ក្រដាសអនុម័ត)...');
      const url3 = await captureElementToPdf(
        'official-pdf-report-document',
        `៣_របាយការណ៍ក្រដាសអនុម័ត_${cleanStart}_ដល់_${cleanEnd}.pdf`
      );

      setDownloadedBlobUrls([
        { label: '១. របកក្រដាសបង្ហាញតែ១ក្រុម', url: url1 },
        { label: '២. របកក្រដាសអនុម័តតាមបណ្តាក្រុម', url: url2 },
        { label: '៣. របាយការណ៍ក្រដាសអនុម័ត', url: url3 },
      ]);
      setDownloadSuccess(true);
    } catch (err) {
      console.error('Batch Download Error:', err);
      alert('មានបញ្ហាក្នុងការទាញយក PDF ទាំង ៣ របាយការណ៍។ សូមព្យាយាមម្តងទៀត។');
    } finally {
      setIsBatchDownloading(false);
      setBatchStatus(null);
    }
  };

  const handlePrint = () => {
    printA4Document(reportRef.current || 'official-pdf-report-document', {
      orientation: 'portrait',
      documentTitle: 'របាយការណ៍ស្តុក_e-Visa',
    });
  };

  return (
    <div className="space-y-4 w-full animate-fade">
      {/* Top Switcher Navigation Bar */}
      {!hideSwitcher && (
        <ReportTypeSwitchBar currentMode={currentReportMode} onSwitchMode={onSwitchReport} onClose={onClose} />
      )}

      {/* Progress Indicator for Batch Download */}
      {isBatchDownloading && batchStatus && (
        <div className="bg-blue-600 text-white p-3.5 rounded-md shadow-md flex items-center justify-between text-xs font-bold animate-pulse print:hidden">
          <div className="flex items-center gap-2">
            <FolderDown className="w-4 h-4 animate-bounce" />
            <span>{batchStatus}</span>
          </div>
          <span className="text-[11px] opacity-90">សូមរង់ចាំបន្តិច...</span>
        </div>
      )}

      {/* Success Alert Banner with direct opening links */}
      {downloadSuccess && (
        <div className="bg-emerald-50 border border-emerald-300 rounded-md p-3 px-4 flex flex-wrap items-center justify-between gap-3 text-xs text-emerald-900 animate-fade print:hidden">
          <div className="flex items-center gap-2">
            <div className="w-6 h-6 rounded-full bg-emerald-600 text-white flex items-center justify-center font-bold text-xs shrink-0">
              <Check className="w-4 h-4" />
            </div>
            <div>
              <span className="font-bold text-xs sm:text-sm">បានបង្កើត និងទាញយកឯកសារ PDF រួចរាល់ដោយជោគជ័យ!</span>
              <span className="block text-[11px] text-emerald-700 mt-0.5">
                ឯកសារត្រូវបានរក្សាទុកក្នុងកម្រិតច្បាស់ខ្ពស់ (HD/A4) តាមទម្រង់ដែលបានរៀបចំ។
              </span>
            </div>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            {downloadedBlobUrls.map((item, idx) => (
              <a
                key={idx}
                href={item.url}
                target="_blank"
                rel="noopener noreferrer"
                className="bg-emerald-700 hover:bg-emerald-800 text-white px-2.5 py-1.5 rounded text-[11px] font-bold transition flex items-center gap-1 cursor-pointer shadow-2xs"
              >
                <FileText className="w-3.5 h-3.5" />
                <span>បើកមើល {item.label}</span>
              </a>
            ))}
            <button
              type="button"
              onClick={() => setDownloadSuccess(false)}
              className="p-1 hover:bg-emerald-200/60 rounded text-emerald-700 transition cursor-pointer ml-1"
              title="លាក់ដំណឹង"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        </div>
      )}

      {/* Control Panel */}
      <div className="bg-white p-4 rounded-md border border-gray-200 shadow-2xs space-y-3 print:hidden">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-3 border-b border-gray-100 pb-3">
          <div className="flex items-center gap-2">
            <div className="w-8 h-8 rounded bg-blue-50 text-[#007bff] flex items-center justify-center font-bold">
              <Calendar className="w-4 h-4" />
            </div>
            <div>
              <h2 className="text-sm font-bold text-gray-800">របាយការណ៍ក្រដាសអនុម័តទិដ្ឋាការអេឡិចត្រូនិក (គិតតាមកាលបរិច្ឆេទ)</h2>
              <p className="text-xs text-gray-500">កំណត់ចន្លោះកាលបរិច្ឆេទ ចាប់ផ្តើម និងដល់ថ្ងៃ ដើម្បីទាញយក ឬបោះពុម្ពរបាយការណ៍</p>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2 shrink-0">
            <button
              type="button"
              onClick={() => setShowTacteingControls(!showTacteingControls)}
              className="px-3 py-1.5 rounded text-xs font-bold bg-amber-50 hover:bg-amber-100 text-amber-900 border border-amber-300 transition flex items-center gap-1.5 cursor-pointer"
            >
              <Sparkles className="w-3.5 h-3.5 text-amber-600" />
              <span>កំណត់ Tacteing</span>
            </button>

            {onClose && (
              <button
                type="button"
                onClick={onClose}
                className="px-3 py-1.5 rounded text-xs font-bold bg-gray-100 hover:bg-gray-200 text-gray-700 transition flex items-center gap-1.5 cursor-pointer"
              >
                <ArrowLeft className="w-3.5 h-3.5" />
                <span>ត្រឡប់ក្រោយ</span>
              </button>
            )}

            <button
              type="button"
              onClick={handleDownloadPdf}
              disabled={isDownloading || isBatchDownloading}
              className="px-3.5 py-1.5 rounded text-xs font-bold bg-emerald-600 hover:bg-emerald-700 text-white transition flex items-center gap-1.5 cursor-pointer shadow-2xs disabled:opacity-50"
            >
              <Download className="w-3.5 h-3.5" />
              <span>{isDownloading ? 'កំពុងទាញយក...' : 'ទាញយក PDF (A4)'}</span>
            </button>
          </div>
        </div>

        {/* Collapsible Tacteing Control Selector */}
        {showTacteingControls && (
          <div className="pt-2 border-t border-gray-100">
            <TacteingControlSelector
              currentType={tacteingType}
              customImage={tacteingCustomImage}
              onChange={(type, customImg) => {
                setTacteingType(type);
                setTacteingCustomImage(customImg || null);
                saveTacteingSettings(type, customImg);
              }}
            />
          </div>
        )}

        {/* Date Filters Panel */}
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 pt-1 text-xs">
          <div>
            <label className="block font-bold text-gray-700 mb-1 h-5 flex items-center">កាលបរិច្ឆេទចាប់ផ្តើម (From Date):</label>
            <CustomDatePicker
              value={startDate}
              onChange={setStartDate}
              className="h-9 border border-gray-300 rounded px-2.5 font-sans text-xs focus:border-blue-500 focus:outline-none box-border leading-normal bg-white text-gray-800"
            />
          </div>

          <div>
            <label className="block font-bold text-gray-700 mb-1 h-5 flex items-center">រហូតដល់ថ្ងៃទី (To Date):</label>
            <CustomDatePicker
              value={endDate}
              onChange={setEndDate}
              className="h-9 border border-gray-300 rounded px-2.5 font-sans text-xs focus:border-blue-500 focus:outline-none box-border leading-normal bg-white text-gray-800"
            />
          </div>

          <div>
            <label className="block font-bold text-gray-700 mb-1 h-5 flex items-center">កាលបរិច្ឆេទចេញរបាយការណ៍ (Report Date):</label>
            <CustomDatePicker
              value={reportDate}
              onChange={setReportDate}
              className="h-9 border border-gray-300 rounded px-2.5 font-sans text-xs focus:border-blue-500 focus:outline-none box-border leading-normal bg-white text-gray-800"
            />
          </div>
        </div>

        {/* Signature Options Controls */}
        <div className="pt-2 border-t border-gray-100 flex flex-wrap items-center justify-between gap-3 text-xs">
          <div className="flex items-center gap-3">
            <label className="flex items-center gap-2 cursor-pointer font-bold text-gray-700 bg-blue-50/70 hover:bg-blue-100/70 px-3 py-1.5 rounded border border-blue-200 transition">
              <input
                type="checkbox"
                checked={showLeftSignature}
                onChange={(e) => {
                  const val = e.target.checked;
                  setShowLeftSignature(val);
                  try {
                    localStorage.setItem('report_show_left_signature', String(val));
                  } catch (err) {
                    console.error(err);
                  }
                }}
                className="w-4 h-4 text-blue-600 rounded border-gray-300 focus:ring-blue-500 cursor-pointer"
              />
              <span>បង្ហាញហត្ថលេខាផ្នែកខាងឆ្វេង (បានឃើញ និងគោរពជូន)</span>
            </label>
          </div>

          <div className="flex flex-wrap items-center gap-3">
            {showLeftSignature && (
              <div className="flex items-center gap-1.5">
                <span className="text-gray-600 font-medium">ចំណងជើងខាងឆ្វេង:</span>
                <input
                  type="text"
                  value={leftSignatureTitle}
                  onChange={(e) => setLeftSignatureTitle(e.target.value)}
                  placeholder="នាយរងការិយាល័យទទួលបន្ទុក"
                  className="border border-gray-300 rounded px-2.5 py-1 text-xs focus:border-blue-500 focus:outline-none w-48 font-bold text-gray-800 bg-white"
                />
              </div>
            )}
            <div className="flex items-center gap-1.5">
              <span className="text-gray-600 font-medium">ចំណងជើងខាងស្តាំ:</span>
              <select
                value={RIGHT_SIGNATURE_OPTIONS.includes(rightSignatureTitle) ? rightSignatureTitle : 'CUSTOM'}
                onChange={(e) => {
                  if (e.target.value === 'CUSTOM') {
                    setRightSignatureTitle('');
                  } else {
                    setRightSignatureTitle(e.target.value);
                  }
                }}
                className="border border-gray-300 rounded px-2.5 py-1 text-xs focus:border-blue-500 focus:outline-none font-bold text-gray-800 bg-white cursor-pointer shadow-xs"
              >
                {RIGHT_SIGNATURE_OPTIONS.map((opt) => (
                  <option key={opt} value={opt}>
                    {opt}
                  </option>
                ))}
                <option value="CUSTOM">ផ្សេងៗ (បញ្ចូលផ្ទាល់)...</option>
              </select>

              {!RIGHT_SIGNATURE_OPTIONS.includes(rightSignatureTitle) && (
                <input
                  type="text"
                  value={rightSignatureTitle}
                  onChange={(e) => setRightSignatureTitle(e.target.value)}
                  placeholder="បញ្ចូលចំណងជើង..."
                  className="border border-gray-300 rounded px-2.5 py-1 text-xs focus:border-blue-500 focus:outline-none w-36 font-bold text-gray-800 bg-white"
                />
              )}
            </div>
          </div>
        </div>
      </div>

      {/* Report A4 Document Canvas Preview Container */}
      <div className="bg-gray-200/80 p-2 sm:p-6 rounded-md overflow-x-auto flex flex-col items-center print:p-0 print:bg-white print:block">
        <div
          ref={reportRef}
          id="official-pdf-report-document"
          className="bg-white shadow-lg border border-gray-300 rounded-xs text-black leading-relaxed text-sm w-[210mm] min-h-[297mm] max-w-full print:shadow-none print:border-none print:m-0 print:w-[210mm] print:min-h-[297mm] print:max-w-none flex flex-col justify-start font-siemreap shrink-0"
          style={{
            boxSizing: 'border-box',
            paddingTop: '1.2cm',
            paddingLeft: '2.5cm',
            paddingRight: '1.5cm',
            paddingBottom: '1.2cm',
            backgroundColor: '#ffffff',
          }}
        >
          <div>
            {/* Cambodian Government Official Header */}
            <div className="flex justify-between items-start mb-6 pt-1 font-siemreap">
              <div className="text-center font-moul leading-relaxed space-y-0.5 text-black pt-6" style={{ fontSize: '12pt' }}>
                <p className="font-moul" style={{ fontSize: '12pt' }}>ក្រសួងមហាផ្ទៃ</p>
                <p className="font-moul" style={{ fontSize: '12pt' }}>អគ្គនាយកដ្ឋានអន្តោប្រវេសន៍</p>
                <p className="font-moul" style={{ fontSize: '12pt' }}>នាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត</p>
                <p className="font-moul" style={{ fontSize: '12pt' }}>ការិយាល័យទិដ្ឋាការចូល</p>
                <p className="font-moul" style={{ fontSize: '12pt' }}>ផ្នែករដ្ឋបាល</p>
                <TacteingLine type={tacteingType} customImage={tacteingCustomImage} width={160} height={16} />
              </div>

              <div className="text-center font-moul leading-relaxed space-y-0.5 text-black" style={{ fontSize: '12pt' }}>
                <p className="font-moul" style={{ fontSize: '12pt' }}>ព្រះរាជាណាចក្រកម្ពុជា</p>
                <p className="font-moul" style={{ fontSize: '12pt' }}>ជាតិ សាសនា ព្រះមហាក្សត្រ</p>
                <TacteingLine type={tacteingType} customImage={tacteingCustomImage} width={160} height={16} />
              </div>
            </div>

            {/* Document Title Header */}
            <div className="text-center my-6 space-y-1.5">
              <h1 className="font-moul text-[15px] text-black leading-normal">
                របាយការណ៍ក្រដាសអនុម័តទិដ្ឋាការអេឡិចត្រូនិក
              </h1>
              <p className="font-siemreap font-bold text-[13.5px] text-black">
                ការបញ្ចូលស្តុក និងផ្តល់តាមក្រុមទិដ្ឋាការអេឡិចត្រូនិក {reportPeriodLabel}
              </p>
              <p className="font-siemreap text-[13px] text-black">
                ដោយគិតចាប់ពីថ្ងៃទី{startParts.khmerDay} ខែ{startParts.khmerMonth} ឆ្នាំ{startParts.khmerYear} រហូតដល់ថ្ងៃទី{endParts.khmerDay} ខែ{endParts.khmerMonth} ឆ្នាំ{endParts.khmerYear}
              </p>
            </div>

            {/* Main Report Table */}
            <div className="mt-5 mb-4">
              <table className="w-full border-collapse border border-black text-center text-[13px] text-black align-middle">
                <thead>
                  <tr className="text-black font-bold h-9">
                    <th rowSpan={2} style={{ verticalAlign: 'middle' }} className="border border-black px-2 py-2 w-10 text-black text-center align-middle font-bold bg-white">
                      ល.រ
                    </th>
                    <th rowSpan={2} style={{ verticalAlign: 'middle' }} className="border border-black px-3 py-2 min-w-[130px] text-black text-center align-middle font-bold bg-white">
                      ស្តុក
                    </th>
                    <th rowSpan={2} style={{ verticalAlign: 'middle' }} className="border border-black px-2 py-2 w-36 text-black text-center align-middle bg-white">
                      <div className="font-bold leading-snug">សន្និធិដើមគ្រា</div>
                      <div className="text-[12px] font-normal text-black mt-0.5 leading-snug whitespace-nowrap">
                        {openingStockDateLabel}
                      </div>
                    </th>
                    <th colSpan={2} style={{ verticalAlign: 'middle' }} className="border border-black px-2 py-2 text-black font-bold text-center align-middle bg-white whitespace-nowrap">
                      ឯកតា ដុំ
                    </th>
                    <th rowSpan={2} style={{ verticalAlign: 'middle' }} className="border border-black px-2 py-2 w-36 text-black text-center align-middle bg-white">
                      <div className="font-bold leading-snug">សន្និធិនៅសល់</div>
                      <div className="text-[12px] font-normal text-black mt-0.5 leading-snug whitespace-nowrap">
                        {endingStockDateLabel}
                      </div>
                    </th>
                  </tr>
                  <tr className="text-black font-bold h-8">
                    <th style={{ verticalAlign: 'middle' }} className="border border-black px-2 py-1.5 w-36 text-black text-center align-middle font-bold bg-white whitespace-nowrap">
                      បញ្ចូលស្តុក
                    </th>
                    <th style={{ verticalAlign: 'middle' }} className="border border-black px-2 py-1.5 w-36 text-black text-center align-middle font-bold bg-white whitespace-nowrap">
                      ផ្តល់ទៅក្រុម
                    </th>
                  </tr>
                </thead>
                <tbody className="text-black">
                  <tr className="hover:bg-gray-50/50 h-9">
                    <td style={{ verticalAlign: 'middle' }} className="border border-black px-2 py-2 font-bold text-center align-middle leading-normal">១</td>
                    <td style={{ verticalAlign: 'middle' }} className="border border-black px-3 py-2 text-left font-bold align-middle leading-normal">ការិយាល័យ</td>
                    <td style={{ verticalAlign: 'middle' }} className="border border-black px-2 py-2 font-bold text-center align-middle leading-normal">{toKhmerNum(officeOpeningStock)}</td>
                    <td style={{ verticalAlign: 'middle' }} className="border border-black px-2 py-2 font-bold text-center align-middle leading-normal">{toKhmerNum(officeInward)}</td>
                    <td style={{ verticalAlign: 'middle' }} className="border border-black px-2 py-2 font-bold text-center align-middle leading-normal">{toKhmerNum(officeUsed)}</td>
                    <td style={{ verticalAlign: 'middle' }} className="border border-black px-2 py-2 font-bold text-center align-middle leading-normal">{toKhmerNum(officeEndingStock)}</td>
                  </tr>

                  <tr className="bg-[#faece1] font-bold text-black h-9">
                    <td colSpan={2} style={{ verticalAlign: 'middle' }} className="border border-black px-2 py-2 text-center font-moul text-[13px] text-black align-middle leading-normal">
                      សរុបចំនួន
                    </td>
                    <td style={{ verticalAlign: 'middle' }} className="border border-black px-2 py-2 font-bold text-[14px] text-black text-center align-middle leading-normal">{toKhmerNum(totalOpening)}</td>
                    <td style={{ verticalAlign: 'middle' }} className="border border-black px-2 py-2 font-bold text-[14px] text-black text-center align-middle leading-normal">{toKhmerNum(totalInward)}</td>
                    <td style={{ verticalAlign: 'middle' }} className="border border-black px-2 py-2 font-bold text-[14px] text-black text-center align-middle leading-normal">{toKhmerNum(totalUsed)}</td>
                    <td style={{ verticalAlign: 'middle' }} className="border border-black px-2 py-2 font-bold text-[14px] text-black text-center align-middle leading-normal">{toKhmerNum(totalEnding)}</td>
                  </tr>
                </tbody>
              </table>
            </div>

            {/* Bottom Signatures Block */}
            <div className="mt-5 text-[13px] font-siemreap text-black">
              <div className={`grid ${showLeftSignature ? 'grid-cols-2 gap-4' : 'grid-cols-1'} text-center items-start`}>
                {showLeftSignature && (
                  <div className="space-y-1">
                    <p className="font-bold text-black">បានឃើញ និងគោរពជូន</p>
                    <p className="text-black">លោកនាយការិយាល័យ មេត្តាជ្រាប ដោយក្តីអនុគ្រោះ។</p>
                    {reportLunar && <p className="text-black text-[12px] pt-0.5">{reportLunar}</p>}
                    <p className="text-black">
                      ភ្នំពេញ, {formatReportDateLine(reportDate)}
                    </p>
                    <p className="font-moul text-black pt-1.5 text-[13px]">{leftSignatureTitle || 'នាយរងការិយាល័យទទួលបន្ទុក'}</p>
                  </div>
                )}

                <div className={`space-y-1 ${!showLeftSignature ? 'ml-auto w-1/2 text-center' : ''}`}>
                  {reportLunar && <p className="text-black text-[12px]">{reportLunar}</p>}
                  <p className="text-black">
                    ភ្នំពេញ, {formatReportDateLine(reportDate)}
                  </p>
                  <p className="font-moul text-black pt-1.5 text-[13px]">{rightSignatureTitle || 'នាយផ្នែក'}</p>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Offscreen Canvas DOMs for Batch Generation of Reports 1 and 2 */}
      <OffscreenReportContainers
        stockRecords={stockRecords}
        categories={categories}
        startDate={startDate}
        endDate={endDate}
        reportDate={reportDate}
        showLeftSignature={showLeftSignature}
        leftSignatureTitle={leftSignatureTitle}
        tacteingType={tacteingType}
        tacteingCustomImage={tacteingCustomImage}
      />
    </div>
  );
};

// Dedicated Report Component for "របកក្រដាសអនុម័តតាមបណ្តាក្រុម" and "របកក្រដាសបង្ហាញតែ១ក្រុម"
export const EVisaRobokReport: React.FC<EVisaStockReportProps> = ({
  stockRecords,
  categories,
  onClose,
  isSingleTeam = false,
  onSwitchReport,
  currentReportMode = isSingleTeam ? 'singleRobokReport' : 'robokReport',
  hideSwitcher = false,
  assignedTeam,
  lockTeamSelection = false,
}) => {
  const reportRef = useRef<HTMLDivElement>(null);
  const [isDownloading, setIsDownloading] = useState(false);
  const [isBatchDownloading, setIsBatchDownloading] = useState(false);
  const [batchStatus, setBatchStatus] = useState<string | null>(null);
  const [downloadSuccess, setDownloadSuccess] = useState(false);
  const [downloadedBlobUrls, setDownloadedBlobUrls] = useState<{ label: string; url: string }[]>([]);

  const today = new Date();
  const firstDayOfMonth = new Date(today.getFullYear(), today.getMonth(), 1)
    .toISOString()
    .split('T')[0];
  const todayStr = today.toISOString().split('T')[0];

  const [startDate, setStartDate] = useState<string>(firstDayOfMonth);
  const [endDate, setEndDate] = useState<string>(todayStr);
  const [reportDate, setReportDate] = useState<string>(todayStr);

  const [showLeftSignature, setShowLeftSignature] = useState<boolean>(() => {
    try {
      const saved = localStorage.getItem('robok_show_left_signature');
      return saved !== null ? saved === 'true' : true;
    } catch {
      return true;
    }
  });
  const [leftSignatureTitle, setLeftSignatureTitle] = useState<string>('នាយរងការិយាល័យទទួលបន្ទុក');
  const [rightSignatureTitle, setRightSignatureTitle] = useState<string>(
    isSingleTeam ? 'ប្រធានក្រុមផ្តល់ទិដ្ឋាការ' : 'នាយផ្នែក'
  );

  const savedTacteing = getSavedTacteingSettings();
  const [showTacteingControls, setShowTacteingControls] = useState<boolean>(false);
  const [tacteingType, setTacteingType] = useState<TacteingType>(savedTacteing.type);
  const [tacteingCustomImage, setTacteingCustomImage] = useState<string | null>(savedTacteing.customImage);

  const startParts = getKhmerSolarParts(startDate);
  const endParts = getKhmerSolarParts(endDate);
  const reportLunar = getKhmerLunarDate(reportDate);

  const sDateObj = parseDateInput(startDate);
  const eDateObj = parseDateInput(endDate);
  const startYearNum = sDateObj.getFullYear();
  const startMonthNum = sDateObj.getMonth();
  const endYearNum = eDateObj.getFullYear();
  const endMonthNum = eDateObj.getMonth();
  const totalMonths = (endYearNum - startYearNum) * 12 + (endMonthNum - startMonthNum) + 1;

  let reportPeriodLabel = `ប្រចាំខែ ${startParts.khmerMonth} ឆ្នាំ${startParts.khmerYear}`;
  if (totalMonths > 1) {
    const monthNumStr = toKhmerNum(String(totalMonths).padStart(2, '0'));
    reportPeriodLabel = `ប្រចាំរយៈពេល ${monthNumStr}ខែ`;
  }

  const startD = new Date(startDate);
  const dayBeforeStart = new Date(startD);
  dayBeforeStart.setDate(dayBeforeStart.getDate() - 1);
  const dayBeforeParts = getKhmerSolarParts(dayBeforeStart);

  const openingStockDateLabel = `${dayBeforeParts.khmerDay}-${dayBeforeParts.khmerMonth}-${dayBeforeParts.khmerYear}`;
  const endingStockDateLabel = `${endParts.khmerDay}-${endParts.khmerMonth}-${endParts.khmerYear}`;

  const evisaRecords = stockRecords.filter((r) => r.stockType === 'evisa' || !r.stockType);

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

  const rawRobokTeams = categories?.visaTeamsRobok && categories.visaTeamsRobok.length > 0
    ? categories.visaTeamsRobok
    : [
        { id: 'vtr-1', name: 'អាកាស តេជោ', createdAt: '2026-01-01' },
        { id: 'vtr-2', name: 'អាកាស សៀមរាប', createdAt: '2026-01-01' },
        { id: 'vtr-3', name: 'អាកាស ព្រះសីហនុ', createdAt: '2026-01-01' },
        { id: 'vtr-4', name: 'ក្រុមទី ១', createdAt: '2026-01-01' },
        { id: 'vtr-5', name: 'ក្រុមទី ២', createdAt: '2026-01-01' },
      ];

  const robokTeams = rawRobokTeams.filter((team, idx) => {
    const opt =
      (team.id && teamTypeOptions[team.id]) ||
      (team.name && teamTypeOptions[team.name]) ||
      teamTypeOptions[idx];
    return opt ? opt.evisa !== false : true;
  });

  const [selectedTeamId, setSelectedTeamId] = useState<string>(() => {
    if (assignedTeam) {
      const normInit = normalizeTeamName(assignedTeam);
      const found = robokTeams.find(
        (t) => t.id === assignedTeam || t.name === assignedTeam || normalizeTeamName(t.name) === normInit
      );
      if (found) return found.id || found.name;
    }
    if (isSingleTeam && robokTeams.length > 0) {
      return robokTeams[0].id || robokTeams[0].name;
    }
    return 'ALL';
  });

  useEffect(() => {
    if (assignedTeam) {
      const normInit = normalizeTeamName(assignedTeam);
      const found = robokTeams.find(
        (t) => t.id === assignedTeam || t.name === assignedTeam || normalizeTeamName(t.name) === normInit
      );
      if (found && selectedTeamId !== (found.id || found.name)) {
        setSelectedTeamId(found.id || found.name);
      }
    } else if (isSingleTeam && (selectedTeamId === 'ALL' || !selectedTeamId) && robokTeams.length > 0) {
      setSelectedTeamId(robokTeams[0].id || robokTeams[0].name);
    }
  }, [isSingleTeam, robokTeams, assignedTeam]);

  const normStartDate = normalizeDateToISO(startDate);
  const normEndDate = normalizeDateToISO(endDate);

  const allTeamRows = robokTeams.map((team) => {
    const teamPeriodUsage = evisaRecords
      .filter(
        (r) => {
          const rd = normalizeDateToISO(r.date);
          return (
            r.operationType === 'useTeam' &&
            (r.visaTeamRobokId === team.id || r.visaTeamRobokName === team.name) &&
            (!normStartDate || rd >= normStartDate) &&
            (!normEndDate || rd <= normEndDate)
          );
        }
      )
      .reduce((sum, r) => sum + (r.quantityBundles || 0), 0);

    const teamPeriodInward = evisaRecords
      .filter(
        (r) => {
          const rd = normalizeDateToISO(r.date);
          return (
            r.operationType === 'issueTeam' &&
            (r.visaTeamRobokId === team.id || r.visaTeamRobokName === team.name) &&
            (!normStartDate || rd >= normStartDate) &&
            (!normEndDate || rd <= normEndDate)
          );
        }
      )
      .reduce((sum, r) => sum + (r.quantityBundles || 0), 0);

    const teamPriorInward = evisaRecords
      .filter(
        (r) => {
          const rd = normalizeDateToISO(r.date);
          return (
            r.operationType === 'issueTeam' &&
            (r.visaTeamRobokId === team.id || r.visaTeamRobokName === team.name) &&
            normStartDate && rd < normStartDate
          );
        }
      )
      .reduce((sum, r) => sum + (r.quantityBundles || 0), 0);

    const teamPriorUsage = evisaRecords
      .filter(
        (r) => {
          const rd = normalizeDateToISO(r.date);
          return (
            r.operationType === 'useTeam' &&
            (r.visaTeamRobokId === team.id || r.visaTeamRobokName === team.name) &&
            normStartDate && rd < normStartDate
          );
        }
      )
      .reduce((sum, r) => sum + (r.quantityBundles || 0), 0);

    const baseOpening = 0;
    const opening = Math.max(0, baseOpening + teamPriorInward - teamPriorUsage);
    const ending = Math.max(0, opening + teamPeriodInward - teamPeriodUsage);

    return {
      id: team.id,
      name: team.name,
      opening,
      inward: teamPeriodInward,
      used: teamPeriodUsage,
      ending,
    };
  });

  const teamRows = allTeamRows.filter((team) => {
    if (!selectedTeamId || selectedTeamId === 'ALL') return true;
    return team.id === selectedTeamId || team.name === selectedTeamId;
  });

  const selectedTeamObj = robokTeams.find(
    (t) => t.id === selectedTeamId || t.name === selectedTeamId
  );

  const getFullTeamName = (teamObj?: { id: string; name: string }) => {
    if (!teamObj) return 'ផ្នែករដ្ឋបាល';
    const vtrList = categories?.visaTeamsRobok || [];
    const vtList = categories?.visaTeams || [];

    let matchedIdx = vtrList.findIndex((item) => item.id === teamObj.id || item.name === teamObj.name);
    if (matchedIdx !== -1 && vtList[matchedIdx]?.name) {
      return vtList[matchedIdx].name;
    }

    const directMatch = vtList.find((item) => item.id === teamObj.id || item.name === teamObj.name);
    if (directMatch?.name) {
      return directMatch.name;
    }

    return teamObj.name || 'ផ្នែករដ្ឋបាល';
  };

  const headerDepartmentName = selectedTeamObj ? getFullTeamName(selectedTeamObj) : 'ផ្នែករដ្ឋបាល';

  const totalOpening = teamRows.reduce((sum, t) => sum + t.opening, 0);
  const totalInward = teamRows.reduce((sum, t) => sum + t.inward, 0);
  const totalUsed = teamRows.reduce((sum, t) => sum + t.used, 0);
  const totalEnding = teamRows.reduce((sum, t) => sum + t.ending, 0);

  // Single PDF Download
  const handleDownloadPdf = async () => {
    try {
      setIsDownloading(true);
      const cleanStart = startDate.replace(/[/\\?%*:|"<>]/g, '-').trim();
      const cleanEnd = endDate.replace(/[/\\?%*:|"<>]/g, '-').trim();
      const teamSuffix = selectedTeamObj ? `_${selectedTeamObj.name}` : '';
      const docElementId = isSingleTeam ? 'official-pdf-single-team-document' : 'official-pdf-robok-document';
      const fileName = isSingleTeam
        ? `របកក្រដាសអនុម័ត_ក្រុម${teamSuffix}(${cleanStart}_ដល់_${cleanEnd}).pdf`
        : `របកក្រដាសអនុម័ត_តាមបណ្តាក្រុម(${cleanStart}_ដល់_${cleanEnd}).pdf`;

      const blobUrl = await captureElementToPdf(docElementId, fileName);

      setDownloadedBlobUrls([
        {
          label: isSingleTeam ? `របកក្រដាស${selectedTeamObj?.name || '១ក្រុម'}` : 'របកក្រដាសអនុម័តតាមបណ្តាក្រុម',
          url: blobUrl,
        },
      ]);
      setDownloadSuccess(true);
    } catch (err) {
      console.error('Error downloading PDF report:', err);
      alert('មានបញ្ហាក្នុងការទាញយក PDF។ សូមព្យាយាមម្តងទៀត។');
    } finally {
      setIsDownloading(false);
    }
  };

  // Batch Download All 3 PDFs at once
  const handleBatchDownloadAllThree = async () => {
    try {
      setIsBatchDownloading(true);
      setDownloadSuccess(false);

      const cleanStart = startDate.replace(/[/\\?%*:|"<>]/g, '-').trim();
      const cleanEnd = endDate.replace(/[/\\?%*:|"<>]/g, '-').trim();

      // 1. Single Team Report PDF
      setBatchStatus('កំពុងបង្កើត PDF របាយការណ៍ទី ១/៣ (របកក្រដាសបង្ហាញតែ១ក្រុម)...');
      const url1 = await captureElementToPdf(
        'official-pdf-single-team-document',
        `១_របកក្រដាសអនុម័ត_បង្ហាញតែ១ក្រុម(${cleanStart}_ដល់_${cleanEnd}).pdf`
      );

      // 2. All Teams Report PDF
      setBatchStatus('កំពុងបង្កើត PDF របាយការណ៍ទី ២/៣ (របកក្រដាសអនុម័តតាមបណ្តាក្រុម)...');
      const url2 = await captureElementToPdf(
        'official-pdf-robok-document',
        `២_របកក្រដាសអនុម័ត_តាមបណ្តាក្រុម(${cleanStart}_ដល់_${cleanEnd}).pdf`
      );

      // 3. Date-based Report PDF
      setBatchStatus('កំពុងបង្កើត PDF របាយការណ៍ទី ៣/៣ (របាយការណ៍ក្រដាសអនុម័ត)...');
      const url3 = await captureElementToPdf(
        'official-pdf-report-document',
        `៣_របាយការណ៍ក្រដាសអនុម័ត_${cleanStart}_ដល់_${cleanEnd}.pdf`
      );

      setDownloadedBlobUrls([
        { label: '១. របកក្រដាសបង្ហាញតែ១ក្រុម', url: url1 },
        { label: '២. របកក្រដាសអនុម័តតាមបណ្តាក្រុម', url: url2 },
        { label: '៣. របាយការណ៍ក្រដាសអនុម័ត', url: url3 },
      ]);
      setDownloadSuccess(true);
    } catch (err) {
      console.error('Batch Download Error:', err);
      alert('មានបញ្ហាក្នុងការទាញយក PDF ទាំង ៣ របាយការណ៍។ សូមព្យាយាមម្តងទៀត។');
    } finally {
      setIsBatchDownloading(false);
      setBatchStatus(null);
    }
  };

  const handlePrint = () => {
    printA4Document(reportRef.current || (isSingleTeam ? 'official-pdf-single-team-document' : 'official-pdf-robok-document'), {
      orientation: 'portrait',
      documentTitle: isSingleTeam ? 'របកក្រដាស_SingleTeam' : 'របកក្រដាសអនុម័តតាមបណ្តាក្រុម',
    });
  };

  return (
    <div className="space-y-4 w-full animate-fade">
      {/* Top Switcher Navigation Bar */}
      {!hideSwitcher && (
        <ReportTypeSwitchBar currentMode={currentReportMode} onSwitchMode={onSwitchReport} onClose={onClose} />
      )}

      {/* Progress Indicator for Batch Download */}
      {isBatchDownloading && batchStatus && (
        <div className="bg-blue-600 text-white p-3.5 rounded-md shadow-md flex items-center justify-between text-xs font-bold animate-pulse print:hidden">
          <div className="flex items-center gap-2">
            <FolderDown className="w-4 h-4 animate-bounce" />
            <span>{batchStatus}</span>
          </div>
          <span className="text-[11px] opacity-90">សូមរង់ចាំបន្តិច...</span>
        </div>
      )}

      {/* Toast Notification for Download */}
      {downloadSuccess && (
        <div className="bg-emerald-50 border border-emerald-300 p-3.5 rounded-md shadow-sm flex flex-wrap items-center justify-between gap-3 text-emerald-800 text-xs print:hidden animate-fade">
          <div className="flex items-center gap-2">
            <div className="w-6 h-6 rounded-full bg-emerald-200 text-emerald-800 flex items-center justify-center font-bold shrink-0">
              <Check className="w-3.5 h-3.5" />
            </div>
            <div>
              <p className="font-bold">ទាញយកឯកសារ PDF ជោគជ័យ!</p>
              <p className="text-[11px] text-emerald-700">
                ឯកសារត្រូវបានរក្សាទុកក្នុងកុំព្យូទ័ររបស់អ្នកដោយស្វ័យប្រវត្តិ។
              </p>
            </div>
          </div>
          <div className="flex flex-wrap items-center gap-2">
            {downloadedBlobUrls.map((item, idx) => (
              <a
                key={idx}
                href={item.url}
                target="_blank"
                rel="noreferrer"
                className="px-2.5 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded text-[11px] font-bold transition flex items-center gap-1 shadow-2xs"
              >
                <FileText className="w-3 h-3" />
                <span>បើកមើល {item.label}</span>
              </a>
            ))}
            <button
              type="button"
              onClick={() => setDownloadSuccess(false)}
              className="p-1 hover:bg-emerald-200/60 rounded text-emerald-700 transition cursor-pointer ml-1"
            >
              <X className="w-4 h-4" />
            </button>
          </div>
        </div>
      )}

      {/* Top Filter & Control Panel */}
      <div className="bg-white p-4 rounded-md border border-gray-200 shadow-2xs space-y-3 print:hidden">
        <div className="flex flex-col lg:flex-row lg:items-center justify-between gap-3 border-b border-gray-100 pb-3">
          <div className="flex items-center gap-2">
            <div className="w-8 h-8 rounded bg-blue-50 text-[#007bff] flex items-center justify-center font-bold">
              <Calendar className="w-4 h-4" />
            </div>
            <div>
              <h2 className="text-sm font-bold text-gray-800">
                {isSingleTeam ? 'របកក្រដាសបង្ហាញតែ១ក្រុម (Single Team Report)' : 'របកក្រដាសអនុម័តតាមបណ្តាក្រុមអេឡិចត្រូនិក'}
              </h2>
              <p className="text-xs text-gray-500">
                {isSingleTeam
                  ? 'ជ្រើសរើសក្រុម និងកំណត់ចន្លោះកាលបរិច្ឆេទ ដើម្បីទាញយក ឬបោះពុម្ពរបាយការណ៍'
                  : 'កំណត់ចន្លោះកាលបរិច្ឆេទ ដើម្បីទាញយក ឬបោះពុម្ពរបាយការណ៍តាមក្រុម'}
              </p>
            </div>
          </div>

          <div className="flex flex-wrap items-center gap-2 shrink-0">
            <button
              type="button"
              onClick={() => setShowTacteingControls(!showTacteingControls)}
              className="px-3 py-1.5 rounded text-xs font-bold bg-amber-50 hover:bg-amber-100 text-amber-900 border border-amber-300 transition flex items-center gap-1.5 cursor-pointer"
            >
              <Sparkles className="w-3.5 h-3.5 text-amber-600" />
              <span>កំណត់ Tacteing</span>
            </button>

            {onClose && (
              <button
                type="button"
                onClick={onClose}
                className="px-3 py-1.5 rounded text-xs font-bold bg-gray-100 hover:bg-gray-200 text-gray-700 transition flex items-center gap-1.5 cursor-pointer"
              >
                <ArrowLeft className="w-3.5 h-3.5" />
                <span>ត្រឡប់ក្រោយ</span>
              </button>
            )}

            <button
              type="button"
              onClick={handleDownloadPdf}
              disabled={isDownloading || isBatchDownloading}
              className="px-3.5 py-1.5 rounded text-xs font-bold bg-emerald-600 hover:bg-emerald-700 text-white transition flex items-center gap-1.5 cursor-pointer shadow-2xs disabled:opacity-50"
            >
              <Download className="w-3.5 h-3.5" />
              <span>{isDownloading ? 'កំពុងទាញយក...' : 'ទាញយក PDF (A4)'}</span>
            </button>
          </div>
        </div>

        {/* Collapsible Tacteing Control Selector */}
        {showTacteingControls && (
          <div className="pt-2 border-t border-gray-100">
            <TacteingControlSelector
              currentType={tacteingType}
              customImage={tacteingCustomImage}
              onChange={(type, customImg) => {
                setTacteingType(type);
                setTacteingCustomImage(customImg || null);
                saveTacteingSettings(type, customImg);
              }}
            />
          </div>
        )}

        {/* Date Filters & Team Selection Panel */}
        <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-3 pt-1 text-xs">
          {(!lockTeamSelection || !assignedTeam) && (
            <div>
              <label className="block font-bold text-gray-700 mb-1 h-5 flex items-center">
                {isSingleTeam ? 'ជ្រើសរើសក្រុម (Select Team):' : 'តម្រងតាមក្រុម (Filter Team):'}
              </label>
              <select
                value={selectedTeamId}
                onChange={(e) => setSelectedTeamId(e.target.value)}
                className="w-full h-9 border border-blue-400 rounded px-2.5 font-sans text-xs focus:border-blue-600 focus:outline-none font-bold text-blue-900 bg-blue-50/50 box-border leading-normal"
              >
                {!isSingleTeam && <option value="ALL">-- បង្ហាញគ្រប់ក្រុមទាំងអស់ --</option>}
                {robokTeams.map((team) => (
                  <option key={team.id} value={team.id || team.name}>
                    {team.name}
                  </option>
                ))}
              </select>
            </div>
          )}

          <div>
            <label className="block font-bold text-gray-700 mb-1 h-5 flex items-center">កាលបរិច្ឆេទចាប់ផ្តើម (From Date):</label>
            <CustomDatePicker
              value={startDate}
              onChange={setStartDate}
              className="h-9 border border-gray-300 rounded px-2.5 font-sans text-xs focus:border-blue-500 focus:outline-none box-border leading-normal bg-white text-gray-800"
            />
          </div>

          <div>
            <label className="block font-bold text-gray-700 mb-1 h-5 flex items-center">រហូតដល់ថ្ងៃទី (To Date):</label>
            <CustomDatePicker
              value={endDate}
              onChange={setEndDate}
              className="h-9 border border-gray-300 rounded px-2.5 font-sans text-xs focus:border-blue-500 focus:outline-none box-border leading-normal bg-white text-gray-800"
            />
          </div>

          <div>
            <label className="block font-bold text-gray-700 mb-1 h-5 flex items-center">កាលបរិច្ឆេទចេញរបាយការណ៍ (Report Date):</label>
            <CustomDatePicker
              value={reportDate}
              onChange={setReportDate}
              className="h-9 border border-gray-300 rounded px-2.5 font-sans text-xs focus:border-blue-500 focus:outline-none box-border leading-normal bg-white text-gray-800"
            />
          </div>
        </div>

        {/* Signature Options Controls */}
        <div className="pt-2 border-t border-gray-100 flex flex-wrap items-center justify-between gap-3 text-xs">
          <div className="flex items-center gap-3">
            <label className="flex items-center gap-2 cursor-pointer font-bold text-gray-700 bg-blue-50/70 hover:bg-blue-100/70 px-3 py-1.5 rounded border border-blue-200 transition">
              <input
                type="checkbox"
                checked={showLeftSignature}
                onChange={(e) => {
                  const val = e.target.checked;
                  setShowLeftSignature(val);
                  try {
                    localStorage.setItem('robok_show_left_signature', String(val));
                  } catch (err) {
                    console.error(err);
                  }
                }}
                className="w-4 h-4 text-blue-600 rounded border-gray-300 focus:ring-blue-500 cursor-pointer"
              />
              <span>បង្ហាញហត្ថលេខាផ្នែកខាងឆ្វេង (បានឃើញ និងគោរពជូន)</span>
            </label>
          </div>

          <div className="flex flex-wrap items-center gap-3">
            {showLeftSignature && (
              <div className="flex items-center gap-1.5">
                <span className="text-gray-600 font-medium">ចំណងជើងខាងឆ្វេង:</span>
                <input
                  type="text"
                  value={leftSignatureTitle}
                  onChange={(e) => setLeftSignatureTitle(e.target.value)}
                  placeholder="នាយរងការិយាល័យទទួលបន្ទុក"
                  className="border border-gray-300 rounded px-2.5 py-1 text-xs focus:border-blue-500 focus:outline-none w-48 font-bold text-gray-800 bg-white"
                />
              </div>
            )}
            <div className="flex items-center gap-1.5">
              <span className="text-gray-600 font-medium">ចំណងជើងខាងស្តាំ:</span>
              <select
                value={(isSingleTeam ? SINGLE_TEAM_RIGHT_SIGNATURE_OPTIONS : RIGHT_SIGNATURE_OPTIONS).includes(rightSignatureTitle) ? rightSignatureTitle : 'CUSTOM'}
                onChange={(e) => {
                  if (e.target.value === 'CUSTOM') {
                    setRightSignatureTitle('');
                  } else {
                    setRightSignatureTitle(e.target.value);
                  }
                }}
                className="border border-gray-300 rounded px-2.5 py-1 text-xs focus:border-blue-500 focus:outline-none font-bold text-gray-800 bg-white cursor-pointer shadow-xs"
              >
                {(isSingleTeam ? SINGLE_TEAM_RIGHT_SIGNATURE_OPTIONS : RIGHT_SIGNATURE_OPTIONS).map((opt) => (
                  <option key={opt} value={opt}>
                    {opt}
                  </option>
                ))}
                <option value="CUSTOM">ផ្សេងៗ (បញ្ចូលផ្ទាល់)...</option>
              </select>

              {!(isSingleTeam ? SINGLE_TEAM_RIGHT_SIGNATURE_OPTIONS : RIGHT_SIGNATURE_OPTIONS).includes(rightSignatureTitle) && (
                <input
                  type="text"
                  value={rightSignatureTitle}
                  onChange={(e) => setRightSignatureTitle(e.target.value)}
                  placeholder="បញ្ចូលចំណងជើង..."
                  className="border border-gray-300 rounded px-2.5 py-1 text-xs focus:border-blue-500 focus:outline-none w-36 font-bold text-gray-800 bg-white"
                />
              )}
            </div>
          </div>
        </div>
      </div>

      {/* Report A4 Document Canvas Preview Container */}
      <div className="bg-gray-200/80 p-2 sm:p-6 rounded-md overflow-x-auto flex flex-col items-center print:p-0 print:bg-white print:block">
        <div
          ref={reportRef}
          id={isSingleTeam ? 'official-pdf-single-team-document' : 'official-pdf-robok-document'}
          className="bg-white shadow-lg border border-gray-300 rounded-xs text-black leading-relaxed text-sm w-[210mm] min-h-[297mm] max-w-full print:shadow-none print:border-none print:m-0 print:w-[210mm] print:min-h-[297mm] print:max-w-none flex flex-col justify-start font-siemreap shrink-0"
          style={{
            boxSizing: 'border-box',
            paddingTop: '1.2cm',
            paddingLeft: '2.5cm',
            paddingRight: '1.5cm',
            paddingBottom: '1.2cm',
            backgroundColor: '#ffffff',
          }}
        >
          <div>
            {/* Cambodian Government Official Header */}
            <div className="flex justify-between items-start mb-6 pt-1 font-siemreap">
              <div className="text-center font-moul leading-relaxed space-y-0.5 text-black pt-6" style={{ fontSize: '12pt' }}>
                <p className="font-moul" style={{ fontSize: '12pt' }}>ក្រសួងមហាផ្ទៃ</p>
                <p className="font-moul" style={{ fontSize: '12pt' }}>អគ្គនាយកដ្ឋានអន្តោប្រវេសន៍</p>
                <p className="font-moul" style={{ fontSize: '12pt' }}>នាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត</p>
                <p className="font-moul" style={{ fontSize: '12pt' }}>ការិយាល័យទិដ្ឋាការចូល</p>
                <p className={headerDepartmentName === 'ផ្នែករដ្ឋបាល' ? 'font-moul' : 'font-siemreap font-bold'} style={{ fontSize: isSingleTeam || headerDepartmentName !== 'ផ្នែករដ្ឋបាល' ? '10pt' : '12pt' }}>
                  {headerDepartmentName}
                </p>
                <TacteingLine type={tacteingType} customImage={tacteingCustomImage} width={160} height={16} />
              </div>

              <div className="text-center font-moul leading-relaxed space-y-0.5 text-black" style={{ fontSize: '12pt' }}>
                <p className="font-moul" style={{ fontSize: '12pt' }}>ព្រះរាជាណាចក្រកម្ពុជា</p>
                <p className="font-moul" style={{ fontSize: '12pt' }}>ជាតិ សាសនា ព្រះមហាក្សត្រ</p>
                <TacteingLine type={tacteingType} customImage={tacteingCustomImage} width={160} height={16} />
              </div>
            </div>

            {/* Document Title Header */}
            <div className="text-center my-6 space-y-1.5">
              <h1 className="font-moul text-[15px] text-black leading-normal">
                {isSingleTeam
                  ? 'របាយការណ៍ក្រដាសអនុម័តទិដ្ឋាការអេឡិចត្រូនិក'
                  : 'របកក្រដាសអនុម័តតាមបណ្តាក្រុមអេឡិចត្រូនិក'}
              </h1>
              <p className="font-siemreap font-bold text-[13.5px] text-black">
                {isSingleTeam
                  ? `ការបើកពីក២ និងការប្រើប្រាស់ក្រដាសអនុម័ត ${reportPeriodLabel}`
                  : `ការបើកពីក២ និងការប្រើប្រាស់តាមក្រុមទិដ្ឋាការអេឡិចត្រូនិក ${reportPeriodLabel}`}
              </p>
              <p className="font-siemreap text-[13px] text-black">
                ដោយគិតចាប់ពីថ្ងៃទី{startParts.khmerDay} ខែ{startParts.khmerMonth} ឆ្នាំ{startParts.khmerYear} រហូតដល់ថ្ងៃទី{endParts.khmerDay} ខែ{endParts.khmerMonth} ឆ្នាំ{endParts.khmerYear}
              </p>
            </div>

            {/* Main Report Table */}
            <div className="mt-5 mb-4">
              <table className="w-full border-collapse border border-black text-center text-[13px] text-black align-middle">
                <thead>
                  <tr className="text-black font-bold h-9">
                    <th rowSpan={2} style={{ verticalAlign: 'middle' }} className="border border-black px-2 py-2 w-10 text-black text-center align-middle font-bold bg-white">
                      ល.រ
                    </th>
                    <th rowSpan={2} style={{ verticalAlign: 'middle' }} className="border border-black px-3 py-2 min-w-[130px] text-black text-center align-middle font-bold bg-white">
                      {isSingleTeam ? 'ស្តុក' : 'ក្រុមផ្តល់ទិដ្ឋាការ'}
                    </th>
                    <th rowSpan={2} style={{ verticalAlign: 'middle' }} className="border border-black px-2 py-2 w-36 text-black text-center align-middle bg-white">
                      <div className="font-bold leading-snug">សន្និធិដើមគ្រា</div>
                      <div className="text-[12px] font-normal text-black mt-0.5 leading-snug whitespace-nowrap">
                        {openingStockDateLabel}
                      </div>
                    </th>
                    <th colSpan={2} style={{ verticalAlign: 'middle' }} className="border border-black px-2 py-2 text-black font-bold text-center align-middle bg-white whitespace-nowrap">
                      ឯកតា ដុំ
                    </th>
                    <th rowSpan={2} style={{ verticalAlign: 'middle' }} className="border border-black px-2 py-2 w-36 text-black text-center align-middle bg-white">
                      <div className="font-bold leading-snug">សន្និធិនៅសល់</div>
                      <div className="text-[12px] font-normal text-black mt-0.5 leading-snug whitespace-nowrap">
                        {endingStockDateLabel}
                      </div>
                    </th>
                  </tr>
                  <tr className="text-black font-bold h-8">
                    <th style={{ verticalAlign: 'middle' }} className="border border-black px-2 py-1.5 w-36 text-black text-center align-middle font-bold bg-white whitespace-nowrap">
                      បើកពីក២
                    </th>
                    <th style={{ verticalAlign: 'middle' }} className="border border-black px-2 py-1.5 w-36 text-black text-center align-middle font-bold bg-white whitespace-nowrap">
                      ការប្រើប្រាស់
                    </th>
                  </tr>
                </thead>
                <tbody className="text-black">
                  {teamRows.map((team, idx) => (
                    <tr key={team.id} className="hover:bg-gray-50/50 h-9">
                      <td style={{ verticalAlign: 'middle' }} className="border border-black px-2 py-2 text-center text-black align-middle leading-normal">{toKhmerNum(idx + 1)}</td>
                      <td style={{ verticalAlign: 'middle' }} className="border border-black px-3 py-2 text-left text-black font-medium align-middle leading-normal">{team.name}</td>
                      <td style={{ verticalAlign: 'middle' }} className="border border-black px-2 py-2 text-center text-black align-middle leading-normal">{toKhmerNum(team.opening)}</td>
                      <td style={{ verticalAlign: 'middle' }} className="border border-black px-2 py-2 text-center text-black align-middle leading-normal">{toKhmerNum(team.inward)}</td>
                      <td style={{ verticalAlign: 'middle' }} className="border border-black px-2 py-2 text-center text-black font-bold align-middle leading-normal">{toKhmerNum(team.used)}</td>
                      <td style={{ verticalAlign: 'middle' }} className="border border-black px-2 py-2 text-center text-black align-middle leading-normal">{toKhmerNum(team.ending)}</td>
                    </tr>
                  ))}

                  <tr className="bg-[#faece1] font-bold text-black h-9">
                    <td colSpan={2} style={{ verticalAlign: 'middle' }} className="border border-black px-2 py-2 text-center font-moul text-[13px] text-black align-middle leading-normal">
                      សរុបចំនួន
                    </td>
                    <td style={{ verticalAlign: 'middle' }} className="border border-black px-2 py-2 font-bold text-[14px] text-black text-center align-middle leading-normal">{toKhmerNum(totalOpening)}</td>
                    <td style={{ verticalAlign: 'middle' }} className="border border-black px-2 py-2 font-bold text-[14px] text-black text-center align-middle leading-normal">{toKhmerNum(totalInward)}</td>
                    <td style={{ verticalAlign: 'middle' }} className="border border-black px-2 py-2 font-bold text-[14px] text-black text-center align-middle leading-normal">{toKhmerNum(totalUsed)}</td>
                    <td style={{ verticalAlign: 'middle' }} className="border border-black px-2 py-2 font-bold text-[14px] text-black text-center align-middle leading-normal">{toKhmerNum(totalEnding)}</td>
                  </tr>
                </tbody>
              </table>
            </div>

            {/* Bottom Signatures Block */}
            <div className="mt-5 text-[13px] font-siemreap text-black">
              <div className={`grid ${showLeftSignature ? 'grid-cols-2 gap-4' : 'grid-cols-1'} text-center items-start`}>
                {showLeftSignature && (
                  <div className="space-y-1">
                    <p className="font-bold text-black">បានឃើញ និងគោរពជូន</p>
                    <p className="text-black">លោកនាយការិយាល័យ មេត្តាជ្រាប ដោយក្តីអនុគ្រោះ។</p>
                    {reportLunar && <p className="text-black text-[12px] pt-0.5">{reportLunar}</p>}
                    <p className="text-black">
                      ភ្នំពេញ, {formatReportDateLine(reportDate)}
                    </p>
                    <p className="font-moul text-black pt-1.5 text-[13px]">{leftSignatureTitle || 'នាយរងការិយាល័យទទួលបន្ទុក'}</p>
                  </div>
                )}

                <div className={`space-y-1 ${!showLeftSignature ? 'ml-auto w-1/2 text-center' : ''}`}>
                  {reportLunar && <p className="text-black text-[12px]">{reportLunar}</p>}
                  <p className="text-black">
                    ភ្នំពេញ, {formatReportDateLine(reportDate)}
                  </p>
                  <p className="font-moul text-black pt-1.5 text-[13px]">{rightSignatureTitle || 'នាយផ្នែក'}</p>
                </div>
              </div>
            </div>
          </div>
        </div>
      </div>

      {/* Offscreen Canvas DOMs for Batch Generation */}
      <OffscreenReportContainers
        stockRecords={stockRecords}
        categories={categories}
        startDate={startDate}
        endDate={endDate}
        reportDate={reportDate}
        showLeftSignature={showLeftSignature}
        leftSignatureTitle={leftSignatureTitle}
        tacteingType={tacteingType}
        tacteingCustomImage={tacteingCustomImage}
      />
    </div>
  );
};

// Export Dedicated Single Team Robok Report Component
export const EVisaSingleRobokReport: React.FC<EVisaStockReportProps> = (props) => {
  return <EVisaRobokReport {...props} isSingleTeam={true} currentReportMode={props.currentReportMode || 'singleRobokReport'} />;
};

// Helper Component to render offscreen DOM nodes for all 3 reports so html2canvas can capture them anytime
const OffscreenReportContainers: React.FC<{
  stockRecords: StockRecord[];
  categories?: CategoriesState;
  startDate: string;
  endDate: string;
  reportDate: string;
  showLeftSignature: boolean;
  leftSignatureTitle: string;
  tacteingType: TacteingType;
  tacteingCustomImage: string | null;
}> = ({
  stockRecords,
  categories,
  startDate,
  endDate,
  reportDate,
  showLeftSignature,
  leftSignatureTitle,
  tacteingType,
  tacteingCustomImage,
}) => {
  const startParts = getKhmerSolarParts(startDate);
  const endParts = getKhmerSolarParts(endDate);
  const reportLunar = getKhmerLunarDate(reportDate);

  const startD = new Date(startDate);
  const dayBeforeStart = new Date(startD);
  dayBeforeStart.setDate(dayBeforeStart.getDate() - 1);
  const dayBeforeParts = getKhmerSolarParts(dayBeforeStart);

  const openingStockDateLabel = `${dayBeforeParts.khmerDay}-${dayBeforeParts.khmerMonth}-${dayBeforeParts.khmerYear}`;
  const endingStockDateLabel = `${endParts.khmerDay}-${endParts.khmerMonth}-${endParts.khmerYear}`;

  const normStartDate = normalizeDateToISO(startDate);
  const normEndDate = normalizeDateToISO(endDate);

  const evisaRecords = stockRecords.filter((r) => r.stockType === 'evisa' || !r.stockType);

  // Office Stats
  const openingK1 = evisaRecords
    .filter(
      (r) =>
        (r.operationType === 'openK1' ||
          r.operationType === 'oldStockK2' ||
          r.operationType === 'oldStock' ||
          r.operationType === 'initial' ||
          r.sourceFrom?.includes('ស្តុកចាស់') ||
          r.sourceFrom?.includes('សន្និធិដើម')) &&
        normalizeDateToISO(r.date) < normStartDate
    )
    .reduce((sum, r) => sum + (r.quantityBundles || 0), 0);
  const openingIssue = evisaRecords
    .filter((r) => r.operationType === 'issueTeam' && normalizeDateToISO(r.date) < normStartDate)
    .reduce((sum, r) => sum + (r.quantityBundles || 0), 0);
  const officeOpeningStock = Math.max(0, openingK1 - openingIssue);

  const officeInward = evisaRecords
    .filter((r) => {
      const rd = normalizeDateToISO(r.date);
      return r.operationType === 'openK1' && (!normStartDate || rd >= normStartDate) && (!normEndDate || rd <= normEndDate);
    })
    .reduce((sum, r) => sum + (r.quantityBundles || 0), 0);
  const officeUsed = evisaRecords
    .filter((r) => {
      const rd = normalizeDateToISO(r.date);
      return r.operationType === 'issueTeam' && (!normStartDate || rd >= normStartDate) && (!normEndDate || rd <= normEndDate);
    })
    .reduce((sum, r) => sum + (r.quantityBundles || 0), 0);
  const officeEndingStock = officeOpeningStock + officeInward - officeUsed;

  // Teams Stats
  const rawRobokTeams = categories?.visaTeamsRobok && categories.visaTeamsRobok.length > 0
    ? categories.visaTeamsRobok
    : [
        { id: 'vtr-1', name: 'អាកាស តេជោ', createdAt: '2026-01-01' },
        { id: 'vtr-2', name: 'អាកាស សៀមរាប', createdAt: '2026-01-01' },
        { id: 'vtr-3', name: 'អាកាស ព្រះសីហនុ', createdAt: '2026-01-01' },
        { id: 'vtr-4', name: 'ក្រុមទី ១', createdAt: '2026-01-01' },
        { id: 'vtr-5', name: 'ក្រុមទី ២', createdAt: '2026-01-01' },
      ];

  const robokTeams = rawRobokTeams;
  const singleTeam = robokTeams[0] || { id: 'vtr-1', name: 'អាកាស តេជោ' };

  const allTeamRows = robokTeams.map((team) => {
    const teamPeriodUsage = evisaRecords
      .filter((r) => {
        const rd = normalizeDateToISO(r.date);
        return (
          r.operationType === 'useTeam' &&
          (r.visaTeamRobokId === team.id || r.visaTeamRobokName === team.name) &&
          (!normStartDate || rd >= normStartDate) &&
          (!normEndDate || rd <= normEndDate)
        );
      })
      .reduce((sum, r) => sum + (r.quantityBundles || 0), 0);
    const teamPeriodInward = evisaRecords
      .filter((r) => {
        const rd = normalizeDateToISO(r.date);
        return (
          r.operationType === 'issueTeam' &&
          (r.visaTeamRobokId === team.id || r.visaTeamRobokName === team.name) &&
          (!normStartDate || rd >= normStartDate) &&
          (!normEndDate || rd <= normEndDate)
        );
      })
      .reduce((sum, r) => sum + (r.quantityBundles || 0), 0);
    const teamPriorInward = evisaRecords
      .filter((r) => {
        const rd = normalizeDateToISO(r.date);
        return (
          r.operationType === 'issueTeam' &&
          (r.visaTeamRobokId === team.id || r.visaTeamRobokName === team.name) &&
          normStartDate && rd < normStartDate
        );
      })
      .reduce((sum, r) => sum + (r.quantityBundles || 0), 0);
    const teamPriorUsage = evisaRecords
      .filter((r) => {
        const rd = normalizeDateToISO(r.date);
        return (
          r.operationType === 'useTeam' &&
          (r.visaTeamRobokId === team.id || r.visaTeamRobokName === team.name) &&
          normStartDate && rd < normStartDate
        );
      })
      .reduce((sum, r) => sum + (r.quantityBundles || 0), 0);

    const opening = Math.max(0, teamPriorInward - teamPriorUsage);
    const ending = Math.max(0, opening + teamPeriodInward - teamPeriodUsage);

    return { id: team.id, name: team.name, opening, inward: teamPeriodInward, used: teamPeriodUsage, ending };
  });

  const totalOpening = officeOpeningStock;
  const totalInward = officeInward;
  const totalUsed = officeUsed;
  const totalEnding = officeEndingStock;

  const allTeamTotalOpening = allTeamRows.reduce((s, t) => s + t.opening, 0);
  const allTeamTotalInward = allTeamRows.reduce((s, t) => s + t.inward, 0);
  const allTeamTotalUsed = allTeamRows.reduce((s, t) => s + t.used, 0);
  const allTeamTotalEnding = allTeamRows.reduce((s, t) => s + t.ending, 0);

  const singleTeamRow = allTeamRows.find((t) => t.id === singleTeam.id) || allTeamRows[0] || { id: '1', name: singleTeam.name, opening: 0, inward: 0, used: 0, ending: 0 };

  return (
    <div style={{ position: 'absolute', left: '-9999px', top: '0', width: '210mm', opacity: '0.01', pointerEvents: 'none' }}>
      {/* 1. Offscreen Single Team Document */}
      <div
        id="official-pdf-single-team-document-offscreen"
        className="bg-white text-black font-siemreap"
        style={{ width: '210mm', minHeight: '297mm', boxSizing: 'border-box', paddingTop: '1.2cm', paddingLeft: '2.5cm', paddingRight: '1.5cm', paddingBottom: '1.2cm', backgroundColor: '#ffffff' }}
      >
        <div className="flex justify-between items-start mb-6 pt-1 font-siemreap">
          <div className="text-center font-moul text-[13px] space-y-0.5 text-black pt-6">
            <p className="font-moul text-[13px]">ក្រសួងមហាផ្ទៃ</p>
            <p className="font-moul text-[13px]">អគ្គនាយកដ្ឋានអន្តោប្រវេសន៍</p>
            <p className="font-moul text-[13px]">នាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត</p>
            <p className="font-moul text-[13px]">ការិយាល័យទិដ្ឋាការចូល</p>
            <p className="font-siemreap font-bold text-[12px]">{singleTeam.name}</p>
            <TacteingLine type={tacteingType} customImage={tacteingCustomImage} width={160} height={16} />
          </div>
          <div className="text-center font-moul text-[13px] space-y-0.5 text-black">
            <p className="font-moul text-[13px]">ព្រះរាជាណាចក្រកម្ពុជា</p>
            <p className="font-moul text-[13px]">ជាតិ សាសនា ព្រះមហាក្សត្រ</p>
            <TacteingLine type={tacteingType} customImage={tacteingCustomImage} width={160} height={16} />
          </div>
        </div>

        <div className="text-center my-6 space-y-1.5">
          <h1 className="font-moul text-[15px] text-black">របាយការណ៍ក្រដាសអនុម័តទិដ្ឋាការអេឡិចត្រូនិក</h1>
          <p className="font-siemreap font-bold text-[13.5px] text-black">
            ការបើកពីក២ និងការប្រើប្រាស់ក្រដាសអនុម័ត ប្រចាំខែ {startParts.khmerMonth} ឆ្នាំ{startParts.khmerYear}
          </p>
          <p className="font-siemreap text-[13px] text-black">
            ដោយគិតចាប់ពីថ្ងៃទី{startParts.khmerDay} ខែ{startParts.khmerMonth} ឆ្នាំ{startParts.khmerYear} រហូតដល់ថ្ងៃទី{endParts.khmerDay} ខែ{endParts.khmerMonth} ឆ្នាំ{endParts.khmerYear}
          </p>
        </div>

        <div className="mt-5 mb-4">
          <table className="w-full border-collapse border border-black text-center text-[13px] text-black align-middle">
            <thead>
              <tr className="font-bold bg-white h-9">
                <th rowSpan={2} className="border border-black px-2 py-2 w-10">ល.រ</th>
                <th rowSpan={2} className="border border-black px-3 py-2 min-w-[130px]">ស្តុក</th>
                <th rowSpan={2} className="border border-black px-2 py-2 w-36">
                  <div className="leading-snug font-bold">សន្និធិដើមគ្រា</div>
                  <div className="text-[12px] font-normal mt-0.5 leading-snug whitespace-nowrap">{openingStockDateLabel}</div>
                </th>
                <th colSpan={2} className="border border-black px-2 py-2">ឯកតា ដុំ</th>
                <th rowSpan={2} className="border border-black px-2 py-2 w-36">
                  <div className="leading-snug font-bold">សន្និធិនៅសល់</div>
                  <div className="text-[12px] font-normal mt-0.5 leading-snug whitespace-nowrap">{endingStockDateLabel}</div>
                </th>
              </tr>
              <tr className="font-bold bg-white h-8">
                <th className="border border-black px-2 py-1.5 w-36">បើកពីក២</th>
                <th className="border border-black px-2 py-1.5 w-36">ការប្រើប្រាស់</th>
              </tr>
            </thead>
            <tbody>
              <tr className="h-9">
                <td className="border border-black px-2 py-2 font-bold">១</td>
                <td className="border border-black px-3 py-2 text-left font-medium">{singleTeamRow.name}</td>
                <td className="border border-black px-2 py-2 font-bold">{toKhmerNum(singleTeamRow.opening)}</td>
                <td className="border border-black px-2 py-2 font-bold">{toKhmerNum(singleTeamRow.inward)}</td>
                <td className="border border-black px-2 py-2 font-bold">{toKhmerNum(singleTeamRow.used)}</td>
                <td className="border border-black px-2 py-2 font-bold">{toKhmerNum(singleTeamRow.ending)}</td>
              </tr>
              <tr className="bg-[#faece1] font-bold h-9">
                <td colSpan={2} className="border border-black px-2 py-2 text-center font-moul text-[13px]">សរុបចំនួន</td>
                <td className="border border-black px-2 py-2 font-bold">{toKhmerNum(singleTeamRow.opening)}</td>
                <td className="border border-black px-2 py-2 font-bold">{toKhmerNum(singleTeamRow.inward)}</td>
                <td className="border border-black px-2 py-2 font-bold">{toKhmerNum(singleTeamRow.used)}</td>
                <td className="border border-black px-2 py-2 font-bold">{toKhmerNum(singleTeamRow.ending)}</td>
              </tr>
            </tbody>
          </table>
        </div>

        <div className="mt-5 text-[13px]">
          <div className={`grid ${showLeftSignature ? 'grid-cols-2 gap-4' : 'grid-cols-1'} text-center items-start`}>
            {showLeftSignature && (
              <div className="space-y-1">
                <p className="font-bold">បានឃើញ និងគោរពជូន</p>
                <p>លោកនាយការិយាល័យ មេត្តាជ្រាប ដោយក្តីអនុគ្រោះ។</p>
                {reportLunar && <p className="text-[12px] pt-0.5">{reportLunar}</p>}
                <p>ភ្នំពេញ, {formatReportDateLine(reportDate)}</p>
                <p className="font-moul pt-1.5 text-[13px]">{leftSignatureTitle || 'នាយរងការិយាល័យទទួលបន្ទុក'}</p>
              </div>
            )}
            <div className={`space-y-1 ${!showLeftSignature ? 'ml-auto w-1/2 text-center' : ''}`}>
              {reportLunar && <p className="text-[12px]">{reportLunar}</p>}
              <p>ភ្នំពេញ, {formatReportDateLine(reportDate)}</p>
              <p className="font-moul pt-1.5 text-[13px]">ប្រធានក្រុមផ្តល់ទិដ្ឋាការ</p>
            </div>
          </div>
        </div>
      </div>

      {/* 2. Offscreen All Teams Document */}
      <div
        id="official-pdf-robok-document-offscreen"
        className="bg-white text-black font-siemreap"
        style={{ width: '210mm', minHeight: '297mm', boxSizing: 'border-box', paddingTop: '1.2cm', paddingLeft: '2.5cm', paddingRight: '1.5cm', paddingBottom: '1.2cm', backgroundColor: '#ffffff' }}
      >
        <div className="flex justify-between items-start mb-6 pt-1 font-siemreap">
          <div className="text-center font-moul leading-relaxed space-y-0.5 text-black pt-6" style={{ fontSize: '12pt' }}>
            <p className="font-moul" style={{ fontSize: '12pt' }}>ក្រសួងមហាផ្ទៃ</p>
            <p className="font-moul" style={{ fontSize: '12pt' }}>អគ្គនាយកដ្ឋានអន្តោប្រវេសន៍</p>
            <p className="font-moul" style={{ fontSize: '12pt' }}>នាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត</p>
            <p className="font-moul" style={{ fontSize: '12pt' }}>ការិយាល័យទិដ្ឋាការចូល</p>
            <p className="font-moul" style={{ fontSize: '12pt' }}>ផ្នែករដ្ឋបាល</p>
            <TacteingLine type={tacteingType} customImage={tacteingCustomImage} width={160} height={16} />
          </div>
          <div className="text-center font-moul leading-relaxed space-y-0.5 text-black" style={{ fontSize: '12pt' }}>
            <p className="font-moul" style={{ fontSize: '12pt' }}>ព្រះរាជាណាចក្រកម្ពុជា</p>
            <p className="font-moul" style={{ fontSize: '12pt' }}>ជាតិ សាសនា ព្រះមហាក្សត្រ</p>
            <TacteingLine type={tacteingType} customImage={tacteingCustomImage} width={160} height={16} />
          </div>
        </div>

        <div className="text-center my-6 space-y-1.5">
          <h1 className="font-moul text-[15px] text-black">របកក្រដាសអនុម័តតាមបណ្តាក្រុមអេឡិចត្រូនិក</h1>
          <p className="font-siemreap font-bold text-[13.5px] text-black">
            ការបើកពីក២ និងការប្រើប្រាស់តាមក្រុមទិដ្ឋាការអេឡិចត្រូនិក ប្រចាំខែ {startParts.khmerMonth} ឆ្នាំ{startParts.khmerYear}
          </p>
          <p className="font-siemreap text-[13px] text-black">
            ដោយគិតចាប់ពីថ្ងៃទី{startParts.khmerDay} ខែ{startParts.khmerMonth} ឆ្នាំ{startParts.khmerYear} រហូតដល់ថ្ងៃទី{endParts.khmerDay} ខែ{endParts.khmerMonth} ឆ្នាំ{endParts.khmerYear}
          </p>
        </div>

        <div className="mt-5 mb-4">
          <table className="w-full border-collapse border border-black text-center text-[13px] text-black align-middle">
            <thead>
              <tr className="font-bold bg-white h-9">
                <th rowSpan={2} className="border border-black px-2 py-2 w-10">ល.រ</th>
                <th rowSpan={2} className="border border-black px-3 py-2 min-w-[130px]">ក្រុមផ្តល់ទិដ្ឋាការ</th>
                <th rowSpan={2} className="border border-black px-2 py-2 w-36">
                  <div className="leading-snug font-bold">សន្និធិដើមគ្រា</div>
                  <div className="text-[12px] font-normal mt-0.5 leading-snug whitespace-nowrap">{openingStockDateLabel}</div>
                </th>
                <th colSpan={2} className="border border-black px-2 py-2">ឯកតា ដុំ</th>
                <th rowSpan={2} className="border border-black px-2 py-2 w-36">
                  <div className="leading-snug font-bold">សន្និធិនៅសល់</div>
                  <div className="text-[12px] font-normal mt-0.5 leading-snug whitespace-nowrap">{endingStockDateLabel}</div>
                </th>
              </tr>
              <tr className="font-bold bg-white h-8">
                <th className="border border-black px-2 py-1.5 w-36">បើកពីក២</th>
                <th className="border border-black px-2 py-1.5 w-36">ការប្រើប្រាស់</th>
              </tr>
            </thead>
            <tbody>
              {allTeamRows.map((team, idx) => (
                <tr key={team.id} className="h-9">
                  <td className="border border-black px-2 py-2 font-bold">{toKhmerNum(idx + 1)}</td>
                  <td className="border border-black px-3 py-2 text-left font-medium">{team.name}</td>
                  <td className="border border-black px-2 py-2 font-bold">{toKhmerNum(team.opening)}</td>
                  <td className="border border-black px-2 py-2 font-bold">{toKhmerNum(team.inward)}</td>
                  <td className="border border-black px-2 py-2 font-bold">{toKhmerNum(team.used)}</td>
                  <td className="border border-black px-2 py-2 font-bold">{toKhmerNum(team.ending)}</td>
                </tr>
              ))}
              <tr className="bg-[#faece1] font-bold h-9">
                <td colSpan={2} className="border border-black px-2 py-2 text-center font-moul text-[13px]">សរុបចំនួន</td>
                <td className="border border-black px-2 py-2 font-bold">{toKhmerNum(allTeamTotalOpening)}</td>
                <td className="border border-black px-2 py-2 font-bold">{toKhmerNum(allTeamTotalInward)}</td>
                <td className="border border-black px-2 py-2 font-bold">{toKhmerNum(allTeamTotalUsed)}</td>
                <td className="border border-black px-2 py-2 font-bold">{toKhmerNum(allTeamTotalEnding)}</td>
              </tr>
            </tbody>
          </table>
        </div>

        <div className="mt-5 text-[13px]">
          <div className={`grid ${showLeftSignature ? 'grid-cols-2 gap-4' : 'grid-cols-1'} text-center items-start`}>
            {showLeftSignature && (
              <div className="space-y-1">
                <p className="font-bold">បានឃើញ និងគោរពជូន</p>
                <p>លោកនាយការិយាល័យ មេត្តាជ្រាប ដោយក្តីអនុគ្រោះ។</p>
                {reportLunar && <p className="text-[12px] pt-0.5">{reportLunar}</p>}
                <p>ភ្នំពេញ, {formatReportDateLine(reportDate)}</p>
                <p className="font-moul pt-1.5 text-[13px]">{leftSignatureTitle || 'នាយរងការិយាល័យទទួលបន្ទុក'}</p>
              </div>
            )}
            <div className={`space-y-1 ${!showLeftSignature ? 'ml-auto w-1/2 text-center' : ''}`}>
              {reportLunar && <p className="text-[12px]">{reportLunar}</p>}
              <p>ភ្នំពេញ, {formatReportDateLine(reportDate)}</p>
              <p className="font-moul pt-1.5 text-[13px]">នាយផ្នែក</p>
            </div>
          </div>
        </div>
      </div>

      {/* 3. Offscreen Official Visa Approval Paper Report */}
      <div
        id="official-pdf-report-document-offscreen"
        className="bg-white text-black font-siemreap"
        style={{ width: '210mm', minHeight: '297mm', boxSizing: 'border-box', paddingTop: '1.2cm', paddingLeft: '2.5cm', paddingRight: '1.5cm', paddingBottom: '1.2cm', backgroundColor: '#ffffff' }}
      >
        <div className="flex justify-between items-start mb-6 pt-1 font-siemreap">
          <div className="text-center font-moul leading-relaxed space-y-0.5 text-black pt-6" style={{ fontSize: '12pt' }}>
            <p className="font-moul" style={{ fontSize: '12pt' }}>ក្រសួងមហាផ្ទៃ</p>
            <p className="font-moul" style={{ fontSize: '12pt' }}>អគ្គនាយកដ្ឋានអន្តោប្រវេសន៍</p>
            <p className="font-moul" style={{ fontSize: '12pt' }}>នាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត</p>
            <p className="font-moul" style={{ fontSize: '12pt' }}>ការិយាល័យទិដ្ឋាការចូល</p>
            <p className="font-moul" style={{ fontSize: '12pt' }}>ផ្នែករដ្ឋបាល</p>
            <TacteingLine type={tacteingType} customImage={tacteingCustomImage} width={160} height={16} />
          </div>
          <div className="text-center font-moul leading-relaxed space-y-0.5 text-black" style={{ fontSize: '12pt' }}>
            <p className="font-moul" style={{ fontSize: '12pt' }}>ព្រះរាជាណាចក្រកម្ពុជា</p>
            <p className="font-moul" style={{ fontSize: '12pt' }}>ជាតិ សាសនា ព្រះមហាក្សត្រ</p>
            <TacteingLine type={tacteingType} customImage={tacteingCustomImage} width={160} height={16} />
          </div>
        </div>

        <div className="text-center my-6 space-y-1.5">
          <h1 className="font-moul text-[15px] text-black leading-normal">
            របាយការណ៍ក្រដាសអនុម័តទិដ្ឋាការអេឡិចត្រូនិក
          </h1>
          <p className="font-siemreap font-bold text-[13.5px] text-black">
            ការបញ្ចូលស្តុក និងផ្តល់តាមក្រុមទិដ្ឋាការអេឡិចត្រូនិក ប្រចាំខែ {startParts.khmerMonth} ឆ្នាំ{startParts.khmerYear}
          </p>
          <p className="font-siemreap text-[13px] text-black">
            ដោយគិតចាប់ពីថ្ងៃទី{startParts.khmerDay} ខែ{startParts.khmerMonth} ឆ្នាំ{startParts.khmerYear} រហូតដល់ថ្ងៃទី{endParts.khmerDay} ខែ{endParts.khmerMonth} ឆ្នាំ{endParts.khmerYear}
          </p>
        </div>

        <div className="mt-5 mb-4">
          <table className="w-full border-collapse border border-black text-center text-[13px] text-black align-middle">
            <thead>
              <tr className="text-black font-bold h-9">
                <th rowSpan={2} className="border border-black px-2 py-2 w-10 text-center font-bold bg-white">
                  ល.រ
                </th>
                <th rowSpan={2} className="border border-black px-3 py-2 min-w-[130px] text-center font-bold bg-white">
                  ស្តុក
                </th>
                <th rowSpan={2} className="border border-black px-2 py-2 w-36 text-center bg-white">
                  <div className="font-bold leading-snug">សន្និធិដើមគ្រា</div>
                  <div className="text-[12px] font-normal text-black mt-0.5 leading-snug whitespace-nowrap">
                    {openingStockDateLabel}
                  </div>
                </th>
                <th colSpan={2} className="border border-black px-2 py-2 font-bold text-center bg-white whitespace-nowrap">
                  ឯកតា ដុំ
                </th>
                <th rowSpan={2} className="border border-black px-2 py-2 w-36 text-center bg-white">
                  <div className="font-bold leading-snug">សន្និធិនៅសល់</div>
                  <div className="text-[12px] font-normal text-black mt-0.5 leading-snug whitespace-nowrap">
                    {endingStockDateLabel}
                  </div>
                </th>
              </tr>
              <tr className="text-black font-bold h-8">
                <th className="border border-black px-2 py-1.5 w-36 text-center font-bold bg-white whitespace-nowrap">
                  បញ្ចូលស្តុក
                </th>
                <th className="border border-black px-2 py-1.5 w-36 text-center font-bold bg-white whitespace-nowrap">
                  ផ្តល់ទៅក្រុម
                </th>
              </tr>
            </thead>
            <tbody className="text-black">
              <tr className="h-9">
                <td className="border border-black px-2 py-2 font-bold text-center leading-normal">១</td>
                <td className="border border-black px-3 py-2 text-left font-bold leading-normal">ការិយាល័យ</td>
                <td className="border border-black px-2 py-2 font-bold text-center leading-normal">{toKhmerNum(officeOpeningStock)}</td>
                <td className="border border-black px-2 py-2 font-bold text-center leading-normal">{toKhmerNum(officeInward)}</td>
                <td className="border border-black px-2 py-2 font-bold text-center leading-normal">{toKhmerNum(officeUsed)}</td>
                <td className="border border-black px-2 py-2 font-bold text-center leading-normal">{toKhmerNum(officeEndingStock)}</td>
              </tr>
              <tr className="bg-[#faece1] font-bold text-black h-9">
                <td colSpan={2} className="border border-black px-2 py-2 text-center font-moul text-[13px] text-black leading-normal">
                  សរុបចំនួន
                </td>
                <td className="border border-black px-2 py-2 font-bold text-[14px] text-black text-center leading-normal">{toKhmerNum(totalOpening)}</td>
                <td className="border border-black px-2 py-2 font-bold text-[14px] text-black text-center leading-normal">{toKhmerNum(totalInward)}</td>
                <td className="border border-black px-2 py-2 font-bold text-[14px] text-black text-center leading-normal">{toKhmerNum(totalUsed)}</td>
                <td className="border border-black px-2 py-2 font-bold text-[14px] text-black text-center leading-normal">{toKhmerNum(totalEnding)}</td>
              </tr>
            </tbody>
          </table>
        </div>

        <div className="mt-5 text-[13px] font-siemreap text-black">
          <div className={`grid ${showLeftSignature ? 'grid-cols-2 gap-4' : 'grid-cols-1'} text-center items-start`}>
            {showLeftSignature && (
              <div className="space-y-1">
                <p className="font-bold text-black">បានឃើញ និងគោរពជូន</p>
                <p className="text-black">លោកនាយការិយាល័យ មេត្តាជ្រាប ដោយក្តីអនុគ្រោះ។</p>
                {reportLunar && <p className="text-black text-[12px]">{reportLunar}</p>}
                <p className="text-black">ភ្នំពេញ, {formatReportDateLine(reportDate)}</p>
                <p className="font-moul text-black pt-1.5 text-[13px]">{leftSignatureTitle || 'នាយរងការិយាល័យទទួលបន្ទុក'}</p>
              </div>
            )}
            <div className={`space-y-1 ${!showLeftSignature ? 'ml-auto w-1/2 text-center' : ''}`}>
              {reportLunar && <p className="text-black text-[12px]">{reportLunar}</p>}
              <p className="text-black">
                ភ្នំពេញ, {formatReportDateLine(reportDate)}
              </p>
              <p className="font-moul text-black pt-1.5 text-[13px]">នាយផ្នែក</p>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};


