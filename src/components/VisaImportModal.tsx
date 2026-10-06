import React, { useState, useEffect, useMemo } from 'react';
import * as XLSX from 'xlsx';
import { CategoriesState, StockRecord, VisaRecord, CategoryItem, Officer } from '../types';
import { Upload, FileSpreadsheet, Download, X, CheckCircle2, AlertTriangle, FileCheck, RefreshCw, ShieldCheck, Copy, Filter, Layers } from 'lucide-react';

export interface StockImportModalProps {
  isOpen: boolean;
  onClose: () => void;
  stockType: 'evisa' | 'sticker';
  categories: CategoriesState;
  existingRecords?: StockRecord[];
  selectedOperationType?: string;
  onImportRecords: (records: StockRecord[]) => void;
  onShowToast: (msg: string, type?: 'success' | 'error' | 'info') => void;
}

export interface ParsedStockRecordWithDuplicate extends StockRecord {
  isDuplicate?: boolean;
  duplicateReason?: string;
  duplicateSource?: 'system' | 'file';
  duplicateIndex?: number;
  rowNumber?: number;
}

const MONTH_NAMES: Record<string, string> = {
  jan: '01', feb: '02', mar: '03', apr: '04', may: '05', jun: '06',
  jul: '07', aug: '08', sep: '09', oct: '10', nov: '11', dec: '12',
  january: '01', february: '02', march: '03', april: '04', june: '06',
  july: '07', august: '08', september: '09', october: '10', november: '11', december: '12'
};

const normalizeKhmerDigits = (val: any): string => {
  if (val === null || val === undefined) return '';
  const str = String(val).trim();
  const khmerDigits = ['០', '១', '២', '៣', '៤', '៥', '៦', '៧', '៨', '៩'];
  let res = '';
  for (let i = 0; i < str.length; i++) {
    const ch = str[i];
    const idx = khmerDigits.indexOf(ch);
    if (idx !== -1) {
      res += idx.toString();
    } else {
      res += ch;
    }
  }
  return res;
};

const calculateEndSerialFromStart = (start: string, quantity: number): string => {
  if (!start || quantity <= 0) return '';
  const cleanStart = normalizeKhmerDigits(start).trim();
  const match = cleanStart.match(/^([A-Za-z]*)(\d+)$/);
  if (!match) return cleanStart;
  const prefix = match[1];
  const numStr = match[2];
  const numLen = numStr.length;
  const startNum = parseInt(numStr, 10);
  if (isNaN(startNum)) return cleanStart;
  const endNum = startNum + quantity - 1;
  return `${prefix}${String(endNum).padStart(numLen, '0')}`;
};

const calculateQtyFromSerials = (start: string, end: string): number => {
  if (!start || !end) return 0;
  const sClean = normalizeKhmerDigits(start).trim();
  const eClean = normalizeKhmerDigits(end).trim();
  const sMatch = sClean.match(/^([A-Za-z]*)(\d+)$/);
  const eMatch = eClean.match(/^([A-Za-z]*)(\d+)$/);
  if (sMatch && eMatch) {
    const sNum = parseInt(sMatch[2], 10);
    const eNum = parseInt(eMatch[2], 10);
    if (!isNaN(sNum) && !isNaN(eNum) && eNum >= sNum) {
      return eNum - sNum + 1;
    }
  }
  return 0;
};

// Helper function to format JS Date without timezone shifts (handles UTC vs Local midnight)
const formatJsDate = (val: Date): string => {
  if (isNaN(val.getTime())) {
    const now = new Date();
    const y = now.getFullYear();
    const m = String(now.getMonth() + 1).padStart(2, '0');
    const d = String(now.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  
  const utcH = val.getUTCHours();
  if (utcH >= 18) {
    return `${val.getFullYear()}-${String(val.getMonth() + 1).padStart(2, '0')}-${String(val.getDate()).padStart(2, '0')}`;
  }
  if (val.getHours() >= 18) {
    return `${val.getUTCFullYear()}-${String(val.getUTCMonth() + 1).padStart(2, '0')}-${String(val.getUTCDate()).padStart(2, '0')}`;
  }
  if (utcH === 0 && val.getUTCMinutes() === 0) {
    return `${val.getUTCFullYear()}-${String(val.getUTCMonth() + 1).padStart(2, '0')}-${String(val.getUTCDate()).padStart(2, '0')}`;
  }
  return `${val.getFullYear()}-${String(val.getMonth() + 1).padStart(2, '0')}-${String(val.getDate()).padStart(2, '0')}`;
};

// Helper function to parse Excel dates accurately (handles Excel serial numbers, JS Date objects, MM/DD/YYYY, DD/MM/YYYY, YYYY-MM-DD, DD-MMM-YYYY)
const parseExcelDate = (val: any): string => {
  if (val === null || val === undefined || val === '') {
    const now = new Date();
    const y = now.getFullYear();
    const m = String(now.getMonth() + 1).padStart(2, '0');
    const d = String(now.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  // 1. If it's a JavaScript Date object
  if (val instanceof Date) {
    return formatJsDate(val);
  }

  // 2. Normalize any Khmer digits in date strings
  const strRaw = normalizeKhmerDigits(val).trim();

  // 3. If it's a Number or numeric string (Excel serial date number, e.g. 43434 = 2018-11-30)
  const num = typeof val === 'number' ? val : (/^\d+(\.\d+)?$/.test(strRaw) ? Number(strRaw) : NaN);
  if (!isNaN(num) && num > 1000 && num < 100000) {
    try {
      if (typeof XLSX !== 'undefined' && XLSX.SSF && XLSX.SSF.parse_date_code) {
        const parsed = XLSX.SSF.parse_date_code(num);
        if (parsed && parsed.y && parsed.m && parsed.d) {
          const y = parsed.y;
          const m = String(parsed.m).padStart(2, '0');
          const d = String(parsed.d).padStart(2, '0');
          return `${y}-${m}-${d}`;
        }
      }
      const date = new Date(Math.round((num - 25569) * 86400 * 1000));
      const y = date.getUTCFullYear();
      const m = String(date.getUTCMonth() + 1).padStart(2, '0');
      const d = String(date.getUTCDate()).padStart(2, '0');
      return `${y}-${m}-${d}`;
    } catch {
      // ignore
    }
  }

  if (!strRaw) {
    const now = new Date();
    const y = now.getFullYear();
    const m = String(now.getMonth() + 1).padStart(2, '0');
    const d = String(now.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  // 4. ISO pattern: YYYY-MM-DD or YYYY/MM/DD or YYYY.MM.DD
  const ymdMatch = strRaw.match(/^(\d{4})[-/.](0?[1-9]|1[0-2])[-/.](0?[1-9]|[12]\d|3[01])(\s.*)?$/);
  if (ymdMatch) {
    const y = ymdMatch[1];
    const m = ymdMatch[2].padStart(2, '0');
    const d = ymdMatch[3].padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  // 5. DD-MMM-YYYY or DD/MMM/YYYY or DD MMM YYYY (e.g. 02-Dec-2024, 29-Nov-2024)
  const ddMmmYyyyMatch = strRaw.match(/^(0?[1-9]|[12]\d|3[01])[-/.\s]+([A-Za-z]{3,9})[-/.\s]+(\d{4}|\d{2})$/);
  if (ddMmmYyyyMatch) {
    const d = ddMmmYyyyMatch[1].padStart(2, '0');
    const mName = ddMmmYyyyMatch[2].toLowerCase();
    const m = MONTH_NAMES[mName];
    let y = ddMmmYyyyMatch[3];
    if (y.length === 2) y = `20${y}`;
    if (m) {
      return `${y}-${m}-${d}`;
    }
  }

  // 6. MMM-DD-YYYY or MMM DD YYYY (e.g. Nov 30, 2018)
  const mmmDdMatch = strRaw.match(/^([A-Za-z]{3,9})[-/.\s]+(0?[1-9]|[12]\d|3[01])[-/,\s]+(\d{4}|\d{2})$/);
  if (mmmDdMatch) {
    const mName = mmmDdMatch[1].toLowerCase();
    const d = mmmDdMatch[2].padStart(2, '0');
    const m = MONTH_NAMES[mName];
    let y = mmmDdMatch[3];
    if (y.length === 2) y = `20${y}`;
    if (m) return `${y}-${m}-${d}`;
  }

  // 7. YYYY-MMM-DD or YYYY/MMM/DD or YYYY MMM DD
  const yyyyMmmDdMatch = strRaw.match(/^(\d{4})[-/.\s]+([A-Za-z]{3,9})[-/.\s]+(0?[1-9]|[12]\d|3[01])$/);
  if (yyyyMmmDdMatch) {
    const y = yyyyMmmDdMatch[1];
    const mName = yyyyMmmDdMatch[2].toLowerCase();
    const m = MONTH_NAMES[mName];
    const d = yyyyMmmDdMatch[3].padStart(2, '0');
    if (m) {
      return `${y}-${m}-${d}`;
    }
  }

  // 8. 3-part numeric dates: Handles MM/DD/YYYY, DD/MM/YYYY, M/D/YY, D/M/YY, MM-DD-YYYY, etc.
  const threePartMatch = strRaw.match(/^(\d{1,4})[-/.](\d{1,4})[-/.](\d{1,4})/);
  if (threePartMatch) {
    const p1 = threePartMatch[1];
    const p2 = threePartMatch[2];
    const p3 = threePartMatch[3];

    // If P1 is 4 digits -> YYYY-MM-DD
    if (p1.length === 4) {
      const y = p1;
      const m = String(Math.min(12, Math.max(1, parseInt(p2, 10)))).padStart(2, '0');
      const d = String(Math.min(31, Math.max(1, parseInt(p3, 10)))).padStart(2, '0');
      return `${y}-${m}-${d}`;
    }

    // P3 is Year (4 digits or 2 digits)
    let y = p3;
    if (y.length === 2) y = `20${y}`;
    if (y.length === 4) {
      const n1 = parseInt(p1, 10);
      const n2 = parseInt(p2, 10);

      let d = 1;
      let m = 1;

      if (n1 > 12 && n2 <= 12) {
        // DD/MM/YYYY e.g. 30/11/2018
        d = n1;
        m = n2;
      } else if (n2 > 12 && n1 <= 12) {
        // MM/DD/YYYY e.g. 11/30/2018 -> month 11, day 30
        m = n1;
        d = n2;
      } else if (n1 <= 12 && n2 <= 12) {
        // Both <= 12: Default to DD/MM/YYYY (Standard format)
        d = n1;
        m = n2;
      } else {
        d = n1;
        m = Math.min(12, n2);
      }

      const dStr = String(Math.min(31, Math.max(1, d))).padStart(2, '0');
      const mStr = String(Math.min(12, Math.max(1, m))).padStart(2, '0');
      return `${y}-${mStr}-${dStr}`;
    }
  }

  // 9. General JS Date parse fallback
  const parsedDate = new Date(strRaw);
  if (!isNaN(parsedDate.getTime())) {
    return formatJsDate(parsedDate);
  }

  const now = new Date();
  const y = now.getFullYear();
  const m = String(now.getMonth() + 1).padStart(2, '0');
  const d = String(now.getDate()).padStart(2, '0');
  return `${y}-${m}-${d}`;
};

export const StockImportModal: React.FC<StockImportModalProps> = ({
  isOpen,
  onClose,
  stockType,
  categories,
  existingRecords = [],
  selectedOperationType,
  onImportRecords,
  onShowToast,
}) => {
  const [file, setFile] = useState<File | null>(null);
  const [parsedRecords, setParsedRecords] = useState<ParsedStockRecordWithDuplicate[]>([]);
  const [invalidRowsCount, setInvalidRowsCount] = useState<number>(0);
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [isDragOver, setIsDragOver] = useState<boolean>(false);
  const [targetOpFilter, setTargetOpFilter] = useState<string>(selectedOperationType || 'all');
  const [importOnlyFiltered, setImportOnlyFiltered] = useState<boolean>(false);
  const [skipDuplicates, setSkipDuplicates] = useState<boolean>(false);
  const [duplicateFilterView, setDuplicateFilterView] = useState<'all' | 'valid' | 'duplicate'>('all');

  useEffect(() => {
    if (isOpen) {
      const initialOp = selectedOperationType || 'all';
      setTargetOpFilter(initialOp);
      setImportOnlyFiltered(initialOp !== 'all');
      setSkipDuplicates(false);
      setDuplicateFilterView('all');
    }
  }, [isOpen, selectedOperationType]);

  const stockTypeName = stockType === 'evisa' ? 'ក្រដាសអនុម័តទិដ្ឋាការអេឡិចត្រូនិក' : 'សន្លឹកទិដ្ឋាការស្អិត';

  const opLabelMap: Record<string, string> = {
    openK1: 'ការបញ្ចូលស្តុក (ក១)',
    oldStockK2: 'ស្តុកចាស់ ក២',
    oldStockTeam: 'ស្តុកចាស់ក្រុម',
    issueTeam: 'ការបើកផ្តល់តាមក្រុម',
    useTeam: 'ការប្រើប្រាស់តាមក្រុម',
    testPrintK2: 'ទិដ្ឋាការសាកក២',
    damaged: 'ទិដ្ឋាការខូចក២',
    damagedTeam: 'ទិដ្ឋាការខូចក្រុម',
    missingTeam: 'ទិដ្ឋាការខ្វះក្រុម',
    returnTeam: 'ទិដ្ឋាការបង្វិលពីក្រុម',
    transferTeam: 'ផ្ទេរការប្រើប្រាស់ក្រុម',
  };

  // Download Excel Template for Stock Import
  const handleDownloadTemplate = () => {
    const isSticker = stockType === 'sticker';
    const headers = isSticker
      ? [
          'ប្រភេទប្រតិបត្តិការ',
          'បើកពី',
          'កាលបរិច្ឆេទ',
          'ម៉ោង',
          'ប្រភេទទិដ្ឋាការ',
          'ចំនួន (សន្លឹក)',
          'ចាប់ផ្តើម',
          'ដល់លេខ',
          'ក្រុមផ្តល់ទិដ្ឋាការ.របក',
          'ឋានន្តរស័ក្កិអ្នកស្នើសុំ',
          'អ្នកស្នើសុំ',
          'អ្នកមកបើក',
          'តួនាទីអ្នកមកបើក',
        ]
      : [
          'ប្រភេទប្រតិបត្តិការ',
          'បើកពី',
          'កាលបរិច្ឆេទ',
          'ម៉ោង',
          'ចំនួន (ដុំ)',
          'ក្រុមផ្តល់ទិដ្ឋាការ.របក',
          'ឋានន្តរស័ក្កិអ្នកស្នើសុំ',
          'អ្នកស្នើសុំ',
          'អ្នកមកបើក',
          'តួនាទីអ្នកមកបើក',
        ];

    let sampleRows: any[] = [];
    const today = new Date().toISOString().split('T')[0];

    if (isSticker) {
      if (targetOpFilter === 'openK1') {
        sampleRows = [
          ['បញ្ចូលស្តុក (ក១)', 'ក១', today, '08:00', 'T', 500, '0000001', '0000500', '', '', '', '', ''],
          ['បញ្ចូលស្តុក (ក១)', 'ក១', today, '08:15', 'E', 300, '0000501', '0000800', '', '', '', '', ''],
        ];
      } else if (targetOpFilter === 'oldStockK2') {
        sampleRows = [
          ['ស្តុកចាស់ ក២', 'ស្តុកចាស់ ក២', today, '08:00', 'T', 400, '0000101', '0000500', '', '', '', '', ''],
          ['ស្តុកចាស់ ក២', 'ស្តុកចាស់ ក២', today, '08:15', 'E', 200, '0000501', '0000700', '', '', '', '', ''],
        ];
      } else if (targetOpFilter === 'oldStockTeam') {
        sampleRows = [
          ['ស្តុកចាស់ក្រុម', 'ស្តុកចាស់របស់ក្រុម', today, '08:00', 'T', 150, '0000701', '0000850', 'អាកាស តេជោ', 'វរៈសេនីយ៍ឯក', 'សុខ វាសនា', 'ចាន់ សុផល', 'ប្រធានក្រុមផ្តល់ទិដ្ឋាការ'],
        ];
      } else if (targetOpFilter === 'testPrintK2') {
        sampleRows = [
          ['ទិដ្ឋាការសាកក២', 'ទិដ្ឋាការសាកក២', today, '08:30', 'E', 10, '0000701', '0000710', '', '', '', '', ''],
          ['ទិដ្ឋាការសាកក២', 'ទិដ្ឋាការសាកក២', today, '09:00', 'T', 5, '0000711', '0000715', '', '', '', '', ''],
        ];
      } else if (targetOpFilter === 'damaged') {
        sampleRows = [
          ['ទិដ្ឋាការខូចក២', 'ទិដ្ឋាការខូចក២', today, '09:00', 'T', 10, '0000801', '0000810', '', '', '', '', ''],
          ['ទិដ្ឋាការខូចក២', 'ទិដ្ឋាការខូចក២', today, '09:30', 'E', 5, '0000811', '0000815', '', '', '', '', ''],
        ];
      } else if (targetOpFilter === 'damagedTeam') {
        sampleRows = [
          ['ទិដ្ឋាការខូចក្រុម', 'ទិដ្ឋាការខូចក្រុម', today, '09:05', 'T', 5, '0000811', '0000815', 'អាកាស តេជោ', '', 'សុខ វាសនា', '', ''],
        ];
      } else if (targetOpFilter === 'missingTeam') {
        sampleRows = [
          ['ទិដ្ឋាការខ្វះក្រុម', 'ទិដ្ឋាការខ្វះក្រុម', today, '09:10', 'T', 2, '0000816', '0000817', 'អាកាស តេជោ', '', 'សុខ វាសនា', '', ''],
        ];
      } else if (targetOpFilter === 'returnTeam') {
        sampleRows = [
          ['ទិដ្ឋាការបង្វិលពីក្រុម', 'បង្វិលពីក្រុម', today, '09:15', 'T', 20, '0000818', '0000837', 'អាកាស តេជោ', 'វរៈសេនីយ៍ឯក', 'សុខ វាសនា', 'ចាន់ សុផល', 'ប្រធានក្រុមផ្តល់ទិដ្ឋាការ'],
        ];
      } else if (targetOpFilter === 'transferTeam') {
        sampleRows = [
          ['ផ្ទេរការប្រើប្រាស់', 'ផ្ទេរការប្រើប្រាស់ក្រុម', today, '09:45', 'T', 30, '0000838', '0000867', 'អាកាស តេជោ', 'វរៈសេនីយ៍ឯក', 'សុខ វាសនា', 'ចាន់ សុផល', 'ប្រធានក្រុមផ្តល់ទិដ្ឋាការ'],
        ];
      } else if (targetOpFilter === 'issueTeam') {
        sampleRows = [
          ['ការបើកផ្តល់តាមក្រុម', 'ក១', today, '10:15', 'E', 50, '0000831', '0000880', 'អាកាស តេជោ', 'វរៈសេនីយ៍ឯក', 'សុខ វាសនា', 'ចាន់ សុផល', 'អនុប្រធានក្រុមផ្តល់ទិដ្ឋាការ'],
          ['ការបើកផ្តល់តាមក្រុម', 'ក១', today, '10:30', 'T', 100, '0000881', '0000980', 'អាកាស តេជោ', 'វរៈសេនីយ៍ឯក', 'សុខ វាសនា', 'ចាន់ សុផល', 'អនុប្រធានក្រុមផ្តល់ទិដ្ឋាការ'],
        ];
      } else if (targetOpFilter === 'useTeam') {
        sampleRows = [
          ['ការប្រើប្រាស់តាមក្រុម', '', today, '11:00', 'T', 50, '0000881', '0000930', 'អាកាស តេជោ', '', '', '', ''],
        ];
      } else {
        sampleRows = [
          ['បញ្ចូលស្តុក (ក១)', 'ក១', today, '08:00', 'T', 500, '0000001', '0000500', '', '', '', '', ''],
          ['ស្តុកចាស់ ក២', 'ស្តុកចាស់ ក២', today, '08:15', 'T', 300, '0000501', '0000800', '', '', '', '', ''],
          ['ស្តុកចាស់ក្រុម', 'ស្តុកចាស់របស់ក្រុម', today, '08:20', 'T', 200, '0000801', '0001000', 'អាកាស តេជោ', 'វរៈសេនីយ៍ឯក', 'សុខ វាសនា', 'ចាន់ សុផល', 'ប្រធានក្រុមផ្តល់ទិដ្ឋាការ'],
          ['ទិដ្ឋាការសាកក២', 'ទិដ្ឋាការសាកក២', today, '08:30', 'E', 10, '0001001', '0001010', '', '', '', '', ''],
          ['ទិដ្ឋាការខូចក២', 'ទិដ្ឋាការខូចក២', today, '09:00', 'T', 10, '0001011', '0001020', '', '', '', '', ''],
          ['ទិដ្ឋាការខូចក្រុម', 'ទិដ្ឋាការខូចក្រុម', today, '09:05', 'T', 5, '0001021', '0001025', 'អាកាស តេជោ', '', 'សុខ វាសនា', '', ''],
          ['ទិដ្ឋាការខ្វះក្រុម', 'ទិដ្ឋាការខ្វះក្រុម', today, '09:10', 'T', 2, '0001026', '0001027', 'អាកាស តេជោ', '', 'សុខ វាសនា', '', ''],
          ['ទិដ្ឋាការបង្វិលពីក្រុម', 'បង្វិលពីក្រុម', today, '09:15', 'T', 20, '0001028', '0001047', 'អាកាស តេជោ', 'វរៈសេនីយ៍ឯក', 'សុខ វាសនា', 'ចាន់ សុផល', 'ប្រធានក្រុមផ្តល់ទិដ្ឋាការ'],
          ['ការបើកផ្តល់តាមក្រុម', 'ក១', today, '10:15', 'E', 50, '0001048', '0001097', 'អាកាស តេជោ', 'វរៈសេនីយ៍ឯក', 'សុខ វាសនា', 'ចាន់ សុផល', 'អនុប្រធានក្រុមផ្តល់ទិដ្ឋាការ'],
          ['ការប្រើប្រាស់តាមក្រុម', '', today, '11:00', 'T', 50, '0001048', '0001097', 'អាកាស តេជោ', '', '', '', ''],
        ];
      }
    } else {
      if (targetOpFilter === 'openK1') {
        sampleRows = [
          ['បញ្ចូលស្តុក (ក១)', 'ក១', today, '08:00', 100, '', '', '', '', ''],
        ];
      } else if (targetOpFilter === 'oldStockK2') {
        sampleRows = [
          ['ស្តុកចាស់ ក២', 'ស្តុកចាស់ ក២', today, '08:15', 50, '', '', '', '', ''],
        ];
      } else if (targetOpFilter === 'oldStockTeam') {
        sampleRows = [
          ['ស្តុកចាស់ក្រុម', 'ស្តុកចាស់របស់ក្រុម', today, '08:20', 30, 'អាកាស តេជោ', 'វរៈសេនីយ៍ឯក', 'សុខ វាសនា', 'ចាន់ សុផល', 'ប្រធានក្រុមផ្តល់ទិដ្ឋាការ'],
        ];
      } else if (targetOpFilter === 'issueTeam') {
        sampleRows = [
          ['ការបើកផ្តល់តាមក្រុម', 'ក១', today, '09:15', 5, 'អាកាស តេជោ', 'វរៈសេនីយ៍ឯក', 'សុខ វាសនា', 'ចាន់ សុផល', 'អនុប្រធានក្រុមផ្តល់ទិដ្ឋាការ'],
        ];
      } else if (targetOpFilter === 'useTeam') {
        sampleRows = [
          ['ការប្រើប្រាស់តាមក្រុម', '', today, '10:00', 4, 'អាកាស តេជោ', '', '', '', ''],
        ];
      } else {
        sampleRows = [
          ['បញ្ចូលស្តុក (ក១)', 'ក១', today, '08:00', 100, '', '', '', '', ''],
          ['ស្តុកចាស់ ក២', 'ស្តុកចាស់ ក២', today, '08:15', 50, '', '', '', '', ''],
          ['ស្តុកចាស់ក្រុម', 'ស្តុកចាស់របស់ក្រុម', today, '08:20', 30, 'អាកាស តេជោ', 'វរៈសេនីយ៍ឯក', 'សុខ វាសនា', 'ចាន់ សុផល', 'ប្រធានក្រុមផ្តល់ទិដ្ឋាការ'],
          ['ការបើកផ្តល់តាមក្រុម', 'ក១', today, '09:15', 5, 'អាកាស តេជោ', 'វរៈសេនីយ៍ឯក', 'សុខ វាសនា', 'ចាន់ សុផល', 'អនុប្រធានក្រុមផ្តល់ទិដ្ឋាការ'],
          ['ការប្រើប្រាស់តាមក្រុម', '', today, '10:00', 4, 'អាកាស តេជោ', '', '', '', ''],
        ];
      }
    }

    const worksheet = XLSX.utils.aoa_to_sheet([headers, ...sampleRows]);

    worksheet['!cols'] = isSticker
      ? [
          { wch: 20 }, // ប្រភេទប្រតិបត្តិការ
          { wch: 12 }, // បើកពី
          { wch: 15 }, // កាលបរិច្ឆេទ
          { wch: 10 }, // ម៉ោង
          { wch: 15 }, // ប្រភេទទិដ្ឋាការ
          { wch: 14 }, // ចំនួន (សន្លឹក)
          { wch: 15 }, // ចាប់ផ្តើម
          { wch: 15 }, // ដល់លេខ
          { wch: 25 }, // ក្រុមផ្តល់ទិដ្ឋាការ.របក
          { wch: 22 }, // ឋានន្តរស័ក្កិអ្នកស្នើសុំ
          { wch: 20 }, // អ្នកស្នើសុំ
          { wch: 20 }, // អ្នកមកបើក
          { wch: 25 }, // តួនាទីអ្នកមកបើក
        ]
      : [
          { wch: 20 }, // ប្រភេទប្រតិបត្តិការ
          { wch: 12 }, // បើកពី
          { wch: 15 }, // កាលបរិច្ឆេទ
          { wch: 10 }, // ម៉ោង
          { wch: 14 }, // ចំនួន (ដុំ)
          { wch: 25 }, // ក្រុមផ្តល់ទិដ្ឋាការ.របក
          { wch: 22 }, // ឋានន្តរស័ក្កិអ្នកស្នើសុំ
          { wch: 20 }, // អ្នកស្នើសុំ
          { wch: 20 }, // អ្នកមកបើក
          { wch: 25 }, // តួនាទីអ្នកមកបើក
        ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'ទម្រង់ស្តុក');
    const opSlug = targetOpFilter !== 'all' ? `_${targetOpFilter}` : '';
    XLSX.writeFile(
      workbook,
      `Stock_Import_Template_${isSticker ? 'Sticker_Visa' : 'EVisa'}${opSlug}.xlsx`
    );
  };

  // Parse Uploaded File
  const handleFileProcess = (selectedFile: File) => {
    setIsProcessing(true);
    setFile(selectedFile);

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(data, { type: 'array', cellDates: false });
        if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
          onShowToast('ឯកសារ Excel គ្មាន Sheet ទេ!', 'error');
          setIsProcessing(false);
          return;
        }

        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];

        // 1. Read sheet as 2D Array (AOA)
        const aoa = XLSX.utils.sheet_to_json<any[]>(worksheet, { header: 1, defval: '' });

        if (!aoa || aoa.length === 0) {
          onShowToast('ឯកសារ Excel គ្មានទិន្នន័យទេ!', 'error');
          setIsProcessing(false);
          return;
        }

        // 2. Multi-row header detection
        let headerRowIndex = -1;
        let bestScore = 0;
        let detectedColMap: Record<string, number> = {};

        const isHeaderMatch = (cellText: string, patterns: string[]): boolean => {
          if (!cellText) return false;
          const clean = cellText.replace(/[^\w\u1780-\u17FF]/g, '').toLowerCase();
          return patterns.some((p) => {
            const pClean = p.replace(/[^\w\u1780-\u17FF]/g, '').toLowerCase();
            return clean.includes(pClean);
          });
        };

        const scanHeaderRow = (rowCells: any[]): { score: number; colMap: Record<string, number> } => {
          const colMap: Record<string, number> = {};
          let score = 0;

          rowCells.forEach((cellVal, colIdx) => {
            const cellStr = String(cellVal || '').trim();
            if (!cellStr) return;
            const cleanStr = cellStr.replace(/[^\w\u1780-\u17FF]/g, '').toLowerCase();

            // 1. Operation Type: match ប្រភេទប្រតិបត្តិការ, ប្រតិបត្តិការ, optype, operation, etc.
            // Ensure we do not accidentally match 'ប្រភេទទិដ្ឋាការ' as operation!
            if (
              colMap.op === undefined &&
              (isHeaderMatch(cellStr, ['ប្រភេទប្រតិបត្តិការ', 'ប្រតិបត្តិការ', 'optype', 'operation', 'op_type', 'operationtype', 'ប្រភេទប្រតិបតិ្តការ', 'មុខសញ្ញា']) ||
               (cleanStr.includes('ប្រតិបត្តិ') || cleanStr.includes('operation')))
            ) {
              colMap.op = colIdx;
              score += 4;
            } else if (
              colMap.visaType === undefined &&
              (isHeaderMatch(cellStr, ['ប្រភេទទិដ្ឋាការ', 'visatype', 'ប្រភេទ visa', 'visas', 'visa_type', 'ប្រភេទvisa']) ||
               (cleanStr.includes('ទិដ្ឋាការ') && !cleanStr.includes('ក្រុម') && !cleanStr.includes('សាក') && !cleanStr.includes('ខូច') && !cleanStr.includes('បង្វិល')))
            ) {
              colMap.visaType = colIdx;
              score += 3;
            } else if (
              colMap.qty === undefined &&
              isHeaderMatch(cellStr, ['ចំនួន', 'qty', 'quantity', 'សន្លឹក', 'ដុំ', 'សរុប', 'total', 'sheets', 'bundles', 'amount'])
            ) {
              colMap.qty = colIdx;
              score += 3;
            } else if (
              colMap.date === undefined &&
              isHeaderMatch(cellStr, ['កាលបរិច្ឆេទ', 'កាលបរិច្ឆែទ', 'ថ្ងៃខែ', 'ថ្ងៃទី', 'date', 'ថ្ងៃ', 'datetime'])
            ) {
              colMap.date = colIdx;
              score += 3;
            } else if (
              colMap.source === undefined &&
              isHeaderMatch(cellStr, ['បើកពី', 'ប្រភព', 'source', 'sourcefrom', 'from', 'ចេញពី', 'ប្រភពដើម'])
            ) {
              colMap.source = colIdx;
              score += 2;
            } else if (
              colMap.time === undefined &&
              isHeaderMatch(cellStr, ['ម៉ោង', 'time', 'hours', 'timeissued'])
            ) {
              colMap.time = colIdx;
              score += 2;
            } else if (
              colMap.startSerial === undefined &&
              isHeaderMatch(cellStr, ['ចាប់ផ្តើម', 'ចាប់ពី', 'start', 'startserial', 'serialstart', 'លេខចាប់ផ្តើម', 'លេខស៊េរីចាប់ពី', 'fromserial'])
            ) {
              colMap.startSerial = colIdx;
              score += 2;
            } else if (
              colMap.endSerial === undefined &&
              isHeaderMatch(cellStr, ['ដល់លេខ', 'end', 'endserial', 'serialend', 'លេខបញ្ចប់', 'លេខស៊េរីដល់', 'toserial'])
            ) {
              colMap.endSerial = colIdx;
              score += 2;
            } else if (
              colMap.team === undefined &&
              isHeaderMatch(cellStr, ['ក្រុមផ្តល់ទិដ្ឋាការ', 'ក្រុមរបក', 'ក្រុម', 'របក', 'team', 'robok', 'section', 'counter'])
            ) {
              colMap.team = colIdx;
              score += 2;
            } else if (
              colMap.rank === undefined &&
              isHeaderMatch(cellStr, ['ឋានន្តរស័ក្កិ', 'ឋានន្តរស័ក្តិ', 'ឋានន្តរសក្កិ', 'rank', 'ឋានន្តរ'])
            ) {
              colMap.rank = colIdx;
              score += 2;
            } else if (
              colMap.requester === undefined &&
              isHeaderMatch(cellStr, ['អ្នកស្នើសុំ', 'ស្នើសុំ', 'requester', 'applicant'])
            ) {
              colMap.requester = colIdx;
              score += 2;
            } else if (
              colMap.collector === undefined &&
              isHeaderMatch(cellStr, ['អ្នកមកបើក', 'អ្នកទទួល', 'អ្នកបើក', 'collector', 'receiver', 'អ្នកយក'])
            ) {
              colMap.collector = colIdx;
              score += 2;
            } else if (
              colMap.role === undefined &&
              isHeaderMatch(cellStr, ['តួនាទីអ្នកមកបើក', 'តួនាទី', 'role', 'position', 'មុខងារ'])
            ) {
              colMap.role = colIdx;
              score += 2;
            }
          });

          return { score, colMap };
        };

        const maxRowsToScan = Math.min(20, aoa.length);
        for (let r = 0; r < maxRowsToScan; r++) {
          const row = aoa[r];
          if (!Array.isArray(row)) continue;
          const { score, colMap } = scanHeaderRow(row);
          if (score > bestScore) {
            bestScore = score;
            headerRowIndex = r;
            detectedColMap = colMap;
          }
        }

        // If operation column wasn't explicitly detected by header, check Column C (index 2)
        if (detectedColMap.op === undefined) {
          // If column index 2 (Column C in Excel) exists and is not already assigned to date or qty, use it as op
          if (detectedColMap.date !== 2 && detectedColMap.qty !== 2) {
            detectedColMap.op = 2;
          }
        }

        // If no recognizable header row was found, fall back to smart column mapping
        let dataStartRow = 0;
        if (bestScore >= 2 && headerRowIndex !== -1) {
          dataStartRow = headerRowIndex + 1;
        } else {
          // Positional fallback (handles Column A: No/Date, Column B: Date/Source, Column C: Op, etc.)
          if (stockType === 'sticker') {
            detectedColMap = {
              op: 2, // Column C: ប្រភេទប្រតិបត្តិការ
              date: 1, // Column B: កាលបរិច្ឆេទ
              source: 3, // Column D: បើកពី / ប្រភព
              time: 4, // Column E: ម៉ោង
              visaType: 5, // Column F: ប្រភេទទិដ្ឋាការ
              qty: 6, // Column G: ចំនួន (សន្លឹក)
              startSerial: 7, // Column H: ចាប់ផ្តើម
              endSerial: 8, // Column I: ដល់លេខ
              team: 9, // Column J: ក្រុមផ្តល់ទិដ្ឋាការ
              rank: 10, // Column K: ឋានន្តរស័ក្តិ
              requester: 11, // Column L: អ្នកស្នើសុំ
              collector: 12, // Column M: អ្នកមកបើក
              role: 13, // Column N: តួនាទីអ្នកមកបើក
            };
          } else {
            detectedColMap = {
              op: 2, // Column C: ប្រភេទប្រតិបត្តិការ
              date: 1, // Column B: កាលបរិច្ឆេទ
              source: 3, // Column D: បើកពី / ប្រភព
              time: 4, // Column E: ម៉ោង
              qty: 5, // Column F: ចំនួន (ដុំ)
              team: 6, // Column G: ក្រុមផ្តល់ទិដ្ឋាការ
              rank: 7, // Column H: ឋានន្តរស័ក្តិ
              requester: 8, // Column I: អ្នកស្នើសុំ
              collector: 9, // Column J: អ្នកមកបើក
              role: 10, // Column K: តួនាទីអ្នកមកបើក
            };
          }
          dataStartRow = 0;
        }

        const validRecords: ParsedStockRecordWithDuplicate[] = [];
        let invalidCount = 0;

        const norm = (s?: string) => (s ? normalizeKhmerDigits(s).trim().toLowerCase() : '');
        const normVisa = (s?: string) => (s ? s.trim().toUpperCase() : '');

        // Fast O(1) Lookup Maps for existing records
        const existingSerialMap = new Map<string, StockRecord>();
        const existingQtyMap = new Map<string, StockRecord>();

        if (existingRecords && existingRecords.length > 0) {
          existingRecords.forEach((ex) => {
            if (ex.stockType !== stockType) return;
            const op = ex.operationType || '';
            const dt = ex.date || '';

            if (stockType === 'sticker' && ex.startSerial) {
              const sStart = norm(ex.startSerial);
              const sEnd = norm(ex.endSerial || ex.startSerial);
              const vType = normVisa(ex.visaType || '');
              const team = norm(ex.visaTeamRobokName);
              const src = norm(ex.sourceFrom);
              const qtyVal = ex.totalSheets || ex.quantityBundles || 0;
              const key = `${op}|${dt}|${sStart}|${sEnd}|${vType}|${team}|${src}|${qtyVal}`;
              if (!existingSerialMap.has(key)) {
                existingSerialMap.set(key, ex);
              }
            } else {
              const team = norm(ex.visaTeamRobokName);
              const src = norm(ex.sourceFrom);
              const vType = normVisa(ex.visaType || '');
              const qtyVal = ex.quantityBundles || ex.totalSheets || 0;
              const key = `${op}|${dt}|${qtyVal}|${team}|${src}|${vType}`;
              if (!existingQtyMap.has(key)) {
                existingQtyMap.set(key, ex);
              }
            }
          });
        }

        // Fast O(1) in-file duplicate tracking maps
        const fileSerialMap = new Map<string, number>();
        const fileQtyMap = new Map<string, number>();

        let consecutiveEmptyCount = 0;

        for (let r = dataStartRow; r < aoa.length; r++) {
          const row = aoa[r];
          if (!row || !Array.isArray(row)) {
            consecutiveEmptyCount++;
            if (consecutiveEmptyCount > 30) break;
            continue;
          }

          // Check if row is completely empty
          const hasAnyValue = row.some((c) => c !== null && c !== undefined && String(c).trim() !== '');
          if (!hasAnyValue) {
            consecutiveEmptyCount++;
            if (consecutiveEmptyCount > 30) break;
            continue;
          }
          consecutiveEmptyCount = 0;

          const getCellVal = (key: string): any => {
            const colIdx = detectedColMap[key];
            if (colIdx !== undefined && colIdx >= 0 && colIdx < row.length) {
              return row[colIdx];
            }
            return '';
          };

          const getCellStr = (key: string): string => {
            const val = getCellVal(key);
            if (val === null || val === undefined) return '';
            return String(val).trim();
          };

          // Get Operation Type: first from detected column, or fallback to Column C (index 2)
          let opTypeRaw = getCellStr('op');
          if (!opTypeRaw && row.length > 2 && row[2] !== null && row[2] !== undefined) {
            opTypeRaw = String(row[2]).trim();
          }

          const sourceFrom = getCellStr('source');
          
          // Get Date: from detected column, or fallback to Column B (index 1) or Column A (index 0)
          let rawDateVal = getCellVal('date');
          if ((rawDateVal === null || rawDateVal === undefined || rawDateVal === '') && row.length > 1) {
            rawDateVal = row[1];
          }
          if ((rawDateVal === null || rawDateVal === undefined || rawDateVal === '') && row.length > 0) {
            rawDateVal = row[0];
          }
          const dateStr = parseExcelDate(rawDateVal);
          
          const timeStr = getCellStr('time');
          const visaType = getCellStr('visaType');

          const rawQtyVal = getCellVal('qty');
          let qty = 0;
          if (rawQtyVal !== null && rawQtyVal !== undefined && rawQtyVal !== '') {
            if (typeof rawQtyVal === 'number' && !isNaN(rawQtyVal)) {
              qty = Math.round(rawQtyVal);
            } else {
              const cleanDigits = normalizeKhmerDigits(rawQtyVal).replace(/[^\d.-]/g, '');
              const parsedN = parseFloat(cleanDigits);
              if (!isNaN(parsedN)) {
                qty = Math.round(parsedN);
              }
            }
          }

          let startSerial = getCellStr('startSerial');
          let endSerial = getCellStr('endSerial');
          const teamRobokRaw = getCellStr('team');
          const reqRankRaw = getCellStr('rank');
          const requesterName = getCellStr('requester');
          const collectorName = getCellStr('collector');
          const collectorRoleRaw = getCellStr('role');

          // If quantity is missing but startSerial and endSerial exist, calculate quantity!
          if (qty <= 0 && startSerial && endSerial) {
            qty = calculateQtyFromSerials(startSerial, endSerial);
          }

          // If quantity exists and startSerial exists but endSerial is missing, calculate endSerial!
          if (qty > 0 && startSerial && !endSerial) {
            endSerial = calculateEndSerialFromStart(startSerial, qty);
          }

          // Determine operationType strictly following data from the Excel row (Column C / op column)
          let operationType: string = 'openK1';
          let finalSourceFrom = sourceFrom;

          const opClean = opTypeRaw.trim();
          const opLower = opClean.toLowerCase();

          if (
            opClean.includes('សាក') ||
            opClean.includes('បោះពុម្ព') ||
            opLower.includes('test') ||
            sourceFrom.includes('សាក')
          ) {
            operationType = 'testPrintK2';
            if (!finalSourceFrom) finalSourceFrom = opClean.includes('ក២') ? opClean : 'ទិដ្ឋាការសាកក២';
          } else if (
            opClean.includes('ខូចក្រុម') ||
            opClean.includes('ខូចតាមក្រុម') ||
            (opClean.includes('ខូច') && opClean.includes('ក្រុម')) ||
            opLower.includes('damagedteam') ||
            sourceFrom.includes('ខូចក្រុម') ||
            sourceFrom.includes('ខូចតាមក្រុម')
          ) {
            operationType = 'damagedTeam';
            if (!finalSourceFrom) finalSourceFrom = 'ទិដ្ឋាការខូចក្រុម';
          } else if (
            opClean.includes('ខ្វះក្រុម') ||
            opClean.includes('ខ្វះតាមក្រុម') ||
            opClean.includes('ខ្វះ') ||
            opClean.includes('បាត់') ||
            opLower.includes('missing') ||
            sourceFrom.includes('ខ្វះ')
          ) {
            operationType = 'missingTeam';
            if (!finalSourceFrom) finalSourceFrom = 'ទិដ្ឋាការខ្វះក្រុម';
          } else if (
            opClean.includes('ខូច') ||
            opClean.includes('មិនបានការ') ||
            opLower.includes('damaged') ||
            opLower.includes('invalid') ||
            opLower.includes('void') ||
            opLower.includes('spoiled') ||
            sourceFrom.includes('ខូច') ||
            sourceFrom.includes('មិនបានការ')
          ) {
            if (teamRobokRaw && !opClean.includes('ក២')) {
              operationType = 'damagedTeam';
              if (!finalSourceFrom) finalSourceFrom = 'ទិដ្ឋាការខូចក្រុម';
            } else {
              operationType = 'damaged';
              if (!finalSourceFrom) finalSourceFrom = opClean.includes('ក២') ? opClean : 'ទិដ្ឋាការខូចក២';
            }
          } else if (
            opClean.includes('គល់សន្លឹក') ||
            opClean.includes('ប្រមូលគល់') ||
            opLower.includes('returnstub') ||
            opLower.includes('stub')
          ) {
            operationType = 'returnStub';
            if (!finalSourceFrom) finalSourceFrom = 'ប្រមូលគល់សន្លឹកទិដ្ឋាការ';
          } else if (
            opClean.includes('បង្វិល') ||
            opClean.includes('ត្រឡប់') ||
            opClean.includes('ប្រគល់') ||
            opLower.includes('return') ||
            sourceFrom.includes('បង្វិល')
          ) {
            operationType = 'returnTeam';
            if (!finalSourceFrom) finalSourceFrom = 'ទិដ្ឋាការបង្វិលពីក្រុម';
          } else if (
            opClean.includes('ស្តុកចាស់ ក២') ||
            opClean.includes('ស្តុកចាស់ក២') ||
            opClean.includes('សន្និធិដើម ក២') ||
            opClean.includes('សន្និធិដើមក២') ||
            opLower.includes('oldstockk2') ||
            opLower.includes('old stock k2') ||
            (opClean.includes('ស្តុកចាស់') && opClean.includes('ក២')) ||
            sourceFrom.includes('ស្តុកចាស់ ក២')
          ) {
            operationType = 'oldStockK2';
            if (!finalSourceFrom) finalSourceFrom = 'ស្តុកចាស់ ក២';
          } else if (
            opClean.includes('ស្តុកចាស់ក្រុម') ||
            opClean.includes('ស្តុកចាស់របស់ក្រុម') ||
            opClean.includes('សន្និធិដើមក្រុម') ||
            opClean.includes('ស្តុកចាស់តាមក្រុម') ||
            opLower.includes('oldstockteam') ||
            opLower.includes('old stock team') ||
            (opClean.includes('ស្តុកចាស់') && opClean.includes('ក្រុម')) ||
            sourceFrom.includes('ស្តុកចាស់ក្រុម') ||
            sourceFrom.includes('ស្តុកចាស់របស់ក្រុម')
          ) {
            operationType = 'oldStockTeam';
            if (!finalSourceFrom) finalSourceFrom = 'ស្តុកចាស់របស់ក្រុម';
          } else if (
            opClean.includes('ស្តុកចាស់') ||
            opClean.includes('សន្និធិដើម') ||
            opClean.includes('សន្និធិដំបូង') ||
            opLower.includes('initial') ||
            opLower.includes('opening') ||
            opLower.includes('old stock')
          ) {
            if (opClean.includes('ក្រុម') || (teamRobokRaw && !opClean.includes('ក២') && !opClean.includes('ក១'))) {
              operationType = 'oldStockTeam';
              if (!finalSourceFrom) finalSourceFrom = 'ស្តុកចាស់របស់ក្រុម';
            } else {
              operationType = 'oldStockK2';
              if (!finalSourceFrom) finalSourceFrom = 'ស្តុកចាស់ ក២';
            }
          } else if (
            opClean.includes('ផ្ទេរ') ||
            opLower.includes('transfer') ||
            sourceFrom.includes('ផ្ទេរ')
          ) {
            operationType = 'transferTeam';
            if (!finalSourceFrom) finalSourceFrom = 'ផ្ទេរការប្រើប្រាស់ក្រុម';
          } else if (
            opClean.includes('ប្រើប្រាស់') ||
            opClean.includes('បានប្រើ') ||
            opClean.includes('ប្រើ') ||
            opLower.includes('use') ||
            opLower.includes('consumption')
          ) {
            operationType = 'useTeam';
          } else if (
            opClean.includes('បើកពីក១') ||
            opClean.includes('បើកពី ក១') ||
            opClean.includes('បើកពី') ||
            opClean.includes('សន្លឹកទិដ្ឋាការបើកពីក១') ||
            opClean.includes('ទិដ្ឋាការបើកពីក១') ||
            opClean.includes('បញ្ចូលស្តុក') ||
            opClean.includes('បញ្ចូលពីក១') ||
            opClean.includes('បញ្ចូល') ||
            opClean.includes('ក១') ||
            opClean.includes('ក.១') ||
            opLower.includes('openk1') ||
            opLower.includes('open') ||
            opLower.includes('inward') ||
            opLower.includes('import') ||
            opLower.includes('receive') ||
            sourceFrom.includes('ក១') ||
            sourceFrom.includes('បើកពីក១') ||
            sourceFrom.includes('បញ្ចូលស្តុក')
          ) {
            operationType = 'openK1';
            if (!finalSourceFrom) finalSourceFrom = 'ក១';
          } else if (
            opClean.includes('បើកផ្តល់') ||
            opClean.includes('ផ្តល់តាមក្រុម') ||
            opClean.includes('បើកជូន') ||
            opClean.includes('ផ្តល់ជូន') ||
            opClean.includes('បើកតាមក្រុម') ||
            opClean.includes('បើក') ||
            opLower.includes('issue') ||
            opLower.includes('distribute')
          ) {
            operationType = 'issueTeam';
          } else if (opClean.includes('ក២')) {
            if (opClean.includes('សាក')) {
              operationType = 'testPrintK2';
            } else if (opClean.includes('ខូច') || opClean.includes('មិនបានការ')) {
              operationType = 'damaged';
            } else {
              operationType = 'openK1';
            }
            if (!finalSourceFrom) finalSourceFrom = opClean;
          } else if (teamRobokRaw) {
            operationType = 'issueTeam';
          } else if (targetOpFilter !== 'all') {
            operationType = targetOpFilter;
          } else {
            operationType = 'openK1';
          }

          if (qty <= 0) {
            invalidCount++;
            continue;
          }

          // Match categories
          let visaTeamRobokId: string | undefined;
          let visaTeamRobokName = teamRobokRaw;
          if (teamRobokRaw && categories.visaTeamsRobok) {
            const matchedTeam = categories.visaTeamsRobok.find(
              (t) => t.name.trim().toLowerCase() === teamRobokRaw.toLowerCase()
            );
            if (matchedTeam) {
              visaTeamRobokId = matchedTeam.id;
              visaTeamRobokName = matchedTeam.name;
            }
          }

          let requestedRankId: string | undefined;
          let requestedRankName = reqRankRaw;
          if (reqRankRaw && categories.ranks) {
            const matchedRank = categories.ranks.find(
              (r) => r.name.trim().toLowerCase() === reqRankRaw.toLowerCase()
            );
            if (matchedRank) {
              requestedRankId = matchedRank.id;
              requestedRankName = matchedRank.name;
            }
          }

          let collectorRoleId: string | undefined;
          let collectorRoleName = collectorRoleRaw;
          if (collectorRoleRaw && categories.positions) {
            const matchedPos = categories.positions.find(
              (p) => p.name.trim().toLowerCase() === collectorRoleRaw.toLowerCase()
            );
            if (matchedPos) {
              collectorRoleId = matchedPos.id;
              collectorRoleName = matchedPos.name;
            }
          }

          // Fast O(1) duplicate checking logic
          let isDup = false;
          let dupReason = '';
          let dupSource: 'system' | 'file' | undefined;
          let dupIdx: number | undefined;

          const sStartKey = startSerial ? norm(startSerial) : '';
          const sEndKey = (endSerial || startSerial) ? norm(endSerial || startSerial) : '';
          const vTypeKey = normVisa(visaType);
          const teamKey = norm(visaTeamRobokName);
          const srcKey = norm(finalSourceFrom);

          const serialKey = `${operationType}|${dateStr}|${sStartKey}|${sEndKey}|${vTypeKey}|${teamKey}|${srcKey}|${qty}`;
          const qtyKey = `${operationType}|${dateStr}|${qty}|${teamKey}|${srcKey}|${vTypeKey}`;

          // 1. Check against existing database records via O(1) Map
          if (stockType === 'sticker' && startSerial) {
            const systemDup = existingSerialMap.get(serialKey);
            if (systemDup) {
              isDup = true;
              dupSource = 'system';
              dupReason = `ស្ទួនជាមួយទិន្នន័យមានស្រាប់ក្នុងប្រព័ន្ធ (ថ្ងៃ ${systemDup.date}, ស៊េរី: ${systemDup.startSerial}${systemDup.endSerial ? `-${systemDup.endSerial}` : ''}${systemDup.visaType ? `, ប្រភេទ ${systemDup.visaType}` : ''})`;
            }
          } else {
            const systemDup = existingQtyMap.get(qtyKey);
            if (systemDup) {
              isDup = true;
              dupSource = 'system';
              dupReason = `ស្ទួនជាមួយទិន្នន័យមានស្រាប់ក្នុងប្រព័ន្ធ (ថ្ងៃ ${systemDup.date}, ចំនួន: ${systemDup.quantityBundles})`;
            }
          }

          // 2. Check against prior parsed rows in this current uploaded file via O(1) Map
          if (!isDup) {
            if (stockType === 'sticker' && startSerial) {
              const priorRow = fileSerialMap.get(serialKey);
              if (priorRow !== undefined) {
                isDup = true;
                dupSource = 'file';
                dupIdx = priorRow;
                dupReason = `ស្ទួនជាមួយជួរទី ${priorRow} ក្នុង File នេះ`;
              }
            } else {
              const priorRow = fileQtyMap.get(qtyKey);
              if (priorRow !== undefined) {
                isDup = true;
                dupSource = 'file';
                dupIdx = priorRow;
                dupReason = `ស្ទួនជាមួយជួរទី ${priorRow} ក្នុង File នេះ`;
              }
            }
          }

          // Register into in-file O(1) map for subsequent rows
          const currentRowNum = r + 1;
          if (stockType === 'sticker' && startSerial) {
            if (!fileSerialMap.has(serialKey)) {
              fileSerialMap.set(serialKey, currentRowNum);
            }
          } else {
            if (!fileQtyMap.has(qtyKey)) {
              fileQtyMap.set(qtyKey, currentRowNum);
            }
          }

          const record: ParsedStockRecordWithDuplicate = {
            id: `stock-import-${Date.now()}-${r}-${Math.random().toString(36).substring(2, 6)}`,
            stockType,
            operationType,
            sourceFrom: finalSourceFrom || (operationType === 'openK1' ? 'ក១' : undefined),
            date: dateStr,
            time: timeStr || '08:00',
            quantityBundles: qty,
            totalSheets: qty,
            visaType: visaType || undefined,
            startSerial: startSerial || undefined,
            endSerial: endSerial || undefined,
            visaTeamRobokId,
            visaTeamRobokName,
            requestedRankId,
            requestedRankName,
            requesterName,
            collectorName,
            collectorRoleId,
            collectorRoleName,
            createdAt: new Date().toISOString(),
            isDuplicate: isDup,
            duplicateReason: dupReason,
            duplicateSource: dupSource,
            duplicateIndex: dupIdx,
            rowNumber: r + 1,
          };

          validRecords.push(record);
        }

        setParsedRecords(validRecords);
        setInvalidRowsCount(invalidCount);
        setIsProcessing(false);

        const dupCount = validRecords.filter((r) => r.isDuplicate).length;

        if (validRecords.length === 0) {
          onShowToast('ពុំមានទិន្នន័យត្រឹមត្រូវសម្រាប់នាំចូលទេ (សូមពិនិត្យមើលចំនួន ឬ លេខស៊េរី)', 'error');
        } else if (dupCount > 0) {
          onShowToast(`បានអាន ${validRecords.length} ជួរ (រកឃើញទិន្នន័យស្ទួន ${dupCount} ជួរ)`, 'info');
        } else {
          onShowToast(`បានអានទិន្នន័យស្តុកចំនួន ${validRecords.length} ជោគជ័យ!`, 'info');
        }
      } catch (err) {
        console.error('Excel parse error:', err);
        onShowToast('មានបញ្ហាក្នុងការអាន File! សូមប្រើប្រាស់ទម្រង់ Excel (.xlsx, .xls, .csv)', 'error');
        setIsProcessing(false);
      }
    };

    reader.readAsArrayBuffer(selectedFile);
  };

  const handleDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDragOver(false);

    if (e.dataTransfer.files && e.dataTransfer.files.length > 0) {
      const droppedFile = e.dataTransfer.files[0];
      if (
        droppedFile.type.includes('sheet') ||
        droppedFile.type.includes('csv') ||
        droppedFile.name.endsWith('.xlsx') ||
        droppedFile.name.endsWith('.xls') ||
        droppedFile.name.endsWith('.csv')
      ) {
        handleFileProcess(droppedFile);
      } else {
        onShowToast('សូមជ្រើសរើសប្រភេទ File Excel (.xlsx, .xls, .csv)', 'error');
      }
    }
  };

  const opCounts = useMemo(() => {
    const counts: Record<string, number> = {};
    parsedRecords.forEach((r) => {
      const op = r.operationType || 'other';
      counts[op] = (counts[op] || 0) + 1;
    });
    return counts;
  }, [parsedRecords]);

  const totalDuplicatesCount = useMemo(() => {
    return parsedRecords.filter((r) => r.isDuplicate).length;
  }, [parsedRecords]);

  const totalUniqueCount = useMemo(() => {
    return parsedRecords.filter((r) => !r.isDuplicate).length;
  }, [parsedRecords]);

  const filteredByOp = useMemo(() => {
    if (targetOpFilter === 'all') return parsedRecords;
    return parsedRecords.filter((r) => r.operationType === targetOpFilter);
  }, [parsedRecords, targetOpFilter]);

  const displayedRecords = useMemo(() => {
    if (duplicateFilterView === 'valid') {
      return filteredByOp.filter((r) => !r.isDuplicate);
    }
    if (duplicateFilterView === 'duplicate') {
      return filteredByOp.filter((r) => r.isDuplicate);
    }
    return filteredByOp;
  }, [filteredByOp, duplicateFilterView]);

  const recordsToImport = useMemo(() => {
    const pool = importOnlyFiltered && targetOpFilter !== 'all' ? filteredByOp : parsedRecords;
    if (skipDuplicates) {
      return pool.filter((r) => !r.isDuplicate);
    }
    return pool;
  }, [parsedRecords, filteredByOp, importOnlyFiltered, targetOpFilter, skipDuplicates]);

  const handleConfirmImport = () => {
    if (recordsToImport.length === 0) {
      onShowToast('គ្មានទិន្នន័យសម្រាប់នាំចូលទេ (ទិន្នន័យទាំងអស់ជាទិន្នន័យស្ទួន)!', 'error');
      return;
    }
    onImportRecords(recordsToImport);

    const basePool = importOnlyFiltered && targetOpFilter !== 'all' ? filteredByOp : parsedRecords;
    const skippedCount = basePool.length - recordsToImport.length;

    const opText = importOnlyFiltered && targetOpFilter !== 'all' && opLabelMap[targetOpFilter]
      ? ` (ប្រភេទ: ${opLabelMap[targetOpFilter]})`
      : '';
    const skipText = skippedCount > 0 ? ` (បានរំលងទិន្នន័យស្ទួន ${skippedCount} ជួរ)` : '';

    onShowToast(`បាននាំចូលទិន្នន័យស្តុកចំនួន ${recordsToImport.length} ជួរជោគជ័យ${opText}${skipText}!`, 'success');
    setParsedRecords([]);
    setFile(null);
    onClose();
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4 animate-fade">
      <div className="bg-white rounded-lg shadow-xl border border-gray-200 w-full max-w-4xl overflow-hidden flex flex-col max-h-[90vh]">
        {/* Header */}
        <div className="bg-gradient-to-r from-emerald-700 to-teal-800 text-white px-5 py-3.5 flex items-center justify-between shrink-0">
          <div className="flex items-center gap-2.5">
            <FileSpreadsheet className="w-5 h-5 text-emerald-300" />
            <div>
              <h3 className="text-sm font-bold">នាំចូលទិន្នន័យស្តុកពី Excel (.xlsx, .csv)</h3>
              <p className="text-[11px] text-emerald-100 font-normal">
                {stockTypeName}
              </p>
            </div>
          </div>
          <button
            onClick={onClose}
            className="text-emerald-100 hover:text-white p-1 rounded-full hover:bg-white/10 transition cursor-pointer"
          >
            <X className="w-5 h-5" />
          </button>
        </div>

        {/* Modal Body */}
        <div className="p-5 overflow-y-auto space-y-4 text-xs">
          {/* Operation Filter Selector for Template and Import */}
          <div className="bg-gray-50 border border-gray-200 rounded-md p-3 flex flex-col sm:flex-row sm:items-center justify-between gap-2.5">
            <div className="flex items-center gap-2">
              <span className="font-bold text-gray-800 text-xs flex items-center gap-1.5 shrink-0">
                <span>🎯 ប្រភេទប្រតិបត្តិការគោលដៅ:</span>
              </span>
              {stockType === 'evisa' ? (
                <select
                  value={targetOpFilter}
                  onChange={(e) => {
                    setTargetOpFilter(e.target.value);
                    if (e.target.value !== 'all') {
                      setImportOnlyFiltered(true);
                    }
                  }}
                  className="bg-white border border-gray-300 rounded px-2.5 py-1 text-xs font-bold text-gray-800 focus:outline-none focus:ring-1 focus:ring-emerald-500 cursor-pointer"
                >
                  <option value="all">ទាំងអស់ (គ្រប់ប្រតិបត្តិការ)</option>
                  <option value="openK1">១. ការបញ្ចូលស្តុក (ក១/ន៨)</option>
                  <option value="issueTeam">២. ការបើកផ្តល់តាមក្រុម</option>
                  <option value="useTeam">៣. ការប្រើប្រាស់តាមក្រុម</option>
                </select>
              ) : (
                <select
                  value={targetOpFilter}
                  onChange={(e) => {
                    setTargetOpFilter(e.target.value);
                    if (e.target.value !== 'all') {
                      setImportOnlyFiltered(true);
                    }
                  }}
                  className="bg-white border border-gray-300 rounded px-2.5 py-1 text-xs font-bold text-gray-800 focus:outline-none focus:ring-1 focus:ring-emerald-500 cursor-pointer"
                >
                  <option value="all">ទាំងអស់ (គ្រប់ប្រតិបត្តិការ)</option>
                  <option value="openK1">១. ការបញ្ចូលស្តុក (ក១)</option>
                  <option value="oldStockK2">២. ស្តុកចាស់ ក២</option>
                  <option value="oldStockTeam">៣. ស្តុកចាស់ក្រុម</option>
                  <option value="issueTeam">៤. ការបើកផ្តល់តាមក្រុម</option>
                  <option value="useTeam">៥. ការប្រើប្រាស់តាមក្រុម</option>
                  <option value="testPrintK2">៦. ទិដ្ឋាការសាកក២</option>
                  <option value="damaged">៧. ទិដ្ឋាការខូចក២</option>
                  <option value="damagedTeam">៨. ទិដ្ឋាការខូចក្រុម</option>
                  <option value="missingTeam">៩. ទិដ្ឋាការខ្វះក្រុម</option>
                  <option value="returnTeam">១០. ទិដ្ឋាការបង្វិលពីក្រុម</option>
                  <option value="transferTeam">១១. ផ្ទេរការប្រើប្រាស់ក្រុម</option>
                </select>
              )}
            </div>

            {targetOpFilter !== 'all' && (
              <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded text-[11px] font-bold bg-emerald-100 text-emerald-800 border border-emerald-300">
                <span>គំរូ & ការចម្រាញ់: {opLabelMap[targetOpFilter] || targetOpFilter}</span>
              </span>
            )}
          </div>

          {/* Step 1: Download Sample Template */}
          <div className="bg-emerald-50/80 border border-emerald-200 rounded-md p-3.5 flex items-center justify-between gap-3">
            <div className="space-y-0.5">
              <p className="font-bold text-emerald-900 text-xs">
                ទម្រង់គំរូ Excel ស្តុក {targetOpFilter !== 'all' ? `(សម្រាប់ ${opLabelMap[targetOpFilter] || targetOpFilter})` : '(គ្រប់ប្រតិបត្តិការ)'}
              </p>
              <p className="text-[11px] text-emerald-700">
                {targetOpFilter !== 'all'
                  ? `ទាញយកទម្រង់គំរូដែលមានទិន្នន័យឧទាហរណ៍ជាក់លាក់សម្រាប់ «${opLabelMap[targetOpFilter] || targetOpFilter}»`
                  : 'សូមទាញយកទម្រង់គំរូជាមុនសិន ដើម្បីបញ្ចូលទិន្នន័យឱ្យត្រូវតាមជួរឈរ (Columns)'}
              </p>
            </div>
            <button
              type="button"
              onClick={handleDownloadTemplate}
              className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1.5 rounded text-xs font-bold transition flex items-center gap-1.5 shrink-0 cursor-pointer shadow-2xs"
            >
              <Download className="w-3.5 h-3.5" />
              <span>ទាញយកទម្រង់គំរូ</span>
            </button>
          </div>

          {/* Drag & Drop File Upload Box */}
          <div
            onDragOver={(e) => {
              e.preventDefault();
              setIsDragOver(true);
            }}
            onDragLeave={() => setIsDragOver(false)}
            onDrop={handleDrop}
            className={`border-2 border-dashed rounded-lg p-6 text-center transition cursor-pointer ${
              isDragOver
                ? 'border-emerald-500 bg-emerald-50/60'
                : 'border-gray-300 hover:border-emerald-400 bg-gray-50/50'
            }`}
          >
            <input
              type="file"
              accept=".xlsx, .xls, .csv"
              onChange={(e) => {
                if (e.target.files && e.target.files[0]) {
                  handleFileProcess(e.target.files[0]);
                }
              }}
              className="hidden"
              id="excel-stock-file-input"
            />
            <label htmlFor="excel-stock-file-input" className="cursor-pointer block space-y-2">
              <div className="w-12 h-12 bg-emerald-100 text-emerald-700 rounded-full flex items-center justify-center mx-auto">
                {isProcessing ? (
                  <RefreshCw className="w-6 h-6 animate-spin" />
                ) : (
                  <Upload className="w-6 h-6" />
                )}
              </div>
              <div className="space-y-1">
                <p className="font-bold text-gray-800 text-xs">
                  {file ? file.name : 'ចុចទីនេះ ឬ អូសទម្លាក់ File Excel (.xlsx, .csv)'}
                </p>
                <p className="text-[11px] text-gray-500">
                  គាំទ្រប្រភេទទម្រង់ Microsoft Excel (.xlsx, .xls) និង CSV
                </p>
              </div>
            </label>
          </div>

          {/* Duplicate Warning Box */}
          {parsedRecords.length > 0 && totalDuplicatesCount > 0 && (
            <div className="bg-gradient-to-r from-amber-50 to-orange-50 border-2 border-amber-300 rounded-lg p-3.5 flex items-start gap-3 text-xs shadow-2xs animate-fade">
              <div className="w-8 h-8 rounded-full bg-amber-100 border border-amber-300 flex items-center justify-center shrink-0 text-amber-700 font-bold mt-0.5">
                <AlertTriangle className="w-4 h-4 text-amber-600" />
              </div>
              <div className="grow space-y-1.5">
                <div className="flex flex-wrap items-center justify-between gap-1">
                  <h4 className="font-bold text-amber-950 text-xs flex items-center gap-1.5">
                    <span>⚠️ រកឃើញទិន្នន័យស្ទួនចំនួន <span className="text-red-700 underline font-extrabold">{totalDuplicatesCount}</span> ជួរ!</span>
                  </h4>
                  <span className="text-[11px] text-amber-800 font-medium">
                    (ទិន្នន័យថ្មីត្រឹមត្រូវ: <strong className="text-emerald-700">{totalUniqueCount}</strong> ជួរ / សរុប: {parsedRecords.length} ជួរ)
                  </span>
                </div>
                <p className="text-[11px] text-amber-800 leading-relaxed">
                  ទិន្នន័យស្ទួនត្រូវបានសម្គាល់ដោយស្លាកពណ៌លឿង «<strong className="text-amber-950">ស្ទួន</strong>» ក្នុងតារាងខាងក្រោម ដើម្បីកុំឱ្យមានការកើនទិន្នន័យស្ទួនក្នុងប្រព័ន្ធ។
                </p>
                <div className="pt-1 flex flex-wrap items-center gap-3">
                  <label className="flex items-center gap-2 cursor-pointer bg-white px-2.5 py-1 rounded border border-amber-300 font-bold text-amber-950 shadow-2xs hover:bg-amber-100/60 transition">
                    <input
                      type="checkbox"
                      checked={skipDuplicates}
                      onChange={(e) => setSkipDuplicates(e.target.checked)}
                      className="w-4 h-4 text-emerald-600 rounded focus:ring-emerald-500 cursor-pointer"
                    />
                    <span className="flex items-center gap-1">
                      <ShieldCheck className="w-3.5 h-3.5 text-emerald-600" />
                      <span>រំលងទិន្នន័យស្ទួនដោយស្វ័យប្រវត្តិ (Skip Duplicates - ណែនាំ)</span>
                    </span>
                  </label>
                  <span className="text-[11px] font-semibold text-gray-700">
                    {skipDuplicates ? (
                      <span className="text-emerald-700">👉 នឹងនាំចូលតែ {recordsToImport.length} ជួរដែលជាទិន្នន័យថ្មីប៉ុណ្ណោះ</span>
                    ) : (
                      <span className="text-rose-700 font-bold">⚠️ នឹងនាំចូលទិន្នន័យទាំងអស់រួមទាំងទិន្នន័យស្ទួន</span>
                    )}
                  </span>
                </div>
              </div>
            </div>
          )}

          {/* Parsed Preview Table */}
          {parsedRecords.length > 0 && (
            <div className="space-y-2.5 pt-2">
              <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 text-xs">
                <div className="flex items-center gap-2">
                  <span className="font-bold text-gray-800 flex items-center gap-1.5">
                    <FileCheck className="w-4 h-4 text-emerald-600" />
                    <span>ទិន្នន័យដែលបានអាន (សរុប {parsedRecords.length} ជួរ)</span>
                  </span>
                  {targetOpFilter !== 'all' && (
                    <span className="text-[11px] font-bold px-2 py-0.5 bg-blue-50 text-blue-700 rounded border border-blue-200">
                      បង្ហាញ: {displayedRecords.length} ជួរ
                    </span>
                  )}
                </div>

                <div className="flex items-center gap-2">
                  {/* Duplicate Filter Tabs */}
                  {totalDuplicatesCount > 0 && (
                    <div className="inline-flex rounded-md shadow-2xs border border-gray-300 p-0.5 bg-gray-100 text-[11px]">
                      <button
                        type="button"
                        onClick={() => setDuplicateFilterView('all')}
                        className={`px-2 py-0.5 rounded font-bold transition cursor-pointer ${
                          duplicateFilterView === 'all'
                            ? 'bg-white text-gray-900 shadow-2xs'
                            : 'text-gray-600 hover:text-gray-900'
                        }`}
                      >
                        ទាំងអស់ ({filteredByOp.length})
                      </button>
                      <button
                        type="button"
                        onClick={() => setDuplicateFilterView('valid')}
                        className={`px-2 py-0.5 rounded font-bold transition cursor-pointer ${
                          duplicateFilterView === 'valid'
                            ? 'bg-emerald-600 text-white shadow-2xs'
                            : 'text-emerald-700 hover:text-emerald-900'
                        }`}
                      >
                        ✅ ថ្មី ({filteredByOp.filter((r) => !r.isDuplicate).length})
                      </button>
                      <button
                        type="button"
                        onClick={() => setDuplicateFilterView('duplicate')}
                        className={`px-2 py-0.5 rounded font-bold transition cursor-pointer ${
                          duplicateFilterView === 'duplicate'
                            ? 'bg-amber-600 text-white shadow-2xs'
                            : 'text-amber-800 hover:text-amber-950'
                        }`}
                      >
                        ⚠️ ស្ទួន ({filteredByOp.filter((r) => r.isDuplicate).length})
                      </button>
                    </div>
                  )}

                  {invalidRowsCount > 0 && (
                    <span className="text-amber-700 font-medium text-[11px] bg-amber-50 px-2 py-0.5 rounded border border-amber-200 flex items-center gap-1">
                      <AlertTriangle className="w-3 h-3" />
                      <span>រំលង {invalidRowsCount} ជួរ (គ្មានចំនួន)</span>
                    </span>
                  )}
                </div>
              </div>

              {/* Operation Category Filter Pills */}
              <div className="flex flex-wrap items-center gap-1.5 bg-gray-50 p-2 rounded border border-gray-200">
                <span className="text-[11px] font-semibold text-gray-500 mr-1">ចម្រាញ់តាម:</span>
                <button
                  type="button"
                  onClick={() => setTargetOpFilter('all')}
                  className={`px-2 py-0.5 rounded-full text-[11px] font-bold transition cursor-pointer ${
                    targetOpFilter === 'all'
                      ? 'bg-emerald-600 text-white shadow-2xs'
                      : 'bg-white text-gray-700 border border-gray-300 hover:bg-gray-100'
                  }`}
                >
                  ទាំងអស់ ({parsedRecords.length})
                </button>
                {Object.keys(opCounts).map((op) => (
                  <button
                    key={op}
                    type="button"
                    onClick={() => {
                      setTargetOpFilter(op);
                      setImportOnlyFiltered(true);
                    }}
                    className={`px-2 py-0.5 rounded-full text-[11px] font-bold transition cursor-pointer ${
                      targetOpFilter === op
                        ? 'bg-emerald-600 text-white shadow-2xs'
                        : 'bg-white text-gray-700 border border-gray-300 hover:bg-gray-100'
                    }`}
                  >
                    {opLabelMap[op] || op} ({opCounts[op]})
                  </button>
                ))}
              </div>

              {/* Import Choice Selector if Filtered */}
              {targetOpFilter !== 'all' && (
                <div className="bg-emerald-50/70 border border-emerald-200 rounded p-2.5 flex items-center justify-between gap-2 text-xs">
                  <span className="text-emerald-900 font-medium">
                    ជ្រើសរើសជម្រើសនាំចូល:
                  </span>
                  <div className="flex items-center gap-3">
                    <label className="flex items-center gap-1.5 cursor-pointer font-bold text-emerald-800">
                      <input
                        type="radio"
                        name="importChoice"
                        checked={importOnlyFiltered}
                        onChange={() => setImportOnlyFiltered(true)}
                        className="text-emerald-600 focus:ring-emerald-500"
                      />
                      <span>នាំចូលតែ «{opLabelMap[targetOpFilter] || targetOpFilter}» ({recordsToImport.length} ជួរ)</span>
                    </label>
                    <label className="flex items-center gap-1.5 cursor-pointer font-medium text-gray-700">
                      <input
                        type="radio"
                        name="importChoice"
                        checked={!importOnlyFiltered}
                        onChange={() => setImportOnlyFiltered(false)}
                        className="text-emerald-600 focus:ring-emerald-500"
                      />
                      <span>នាំចូលទាំងអស់ ({skipDuplicates ? totalUniqueCount : parsedRecords.length} ជួរ)</span>
                    </label>
                  </div>
                </div>
              )}

              <div className="overflow-x-auto max-h-60 border border-gray-200 rounded">
                <table className="w-full text-left text-[11px] border-collapse">
                  <thead className="bg-gray-100 sticky top-0 font-bold text-gray-700 border-b border-gray-200">
                    <tr>
                      <th className="p-2 border-r border-gray-200 text-center w-8">ល.រ</th>
                      <th className="p-2 border-r border-gray-200 text-center w-24">ស្ថានភាព</th>
                      <th className="p-2 border-r border-gray-200">ប្រតិបត្តិការ</th>
                      <th className="p-2 border-r border-gray-200">ថ្ងៃទី / ម៉ោង</th>
                      {stockType === 'sticker' && (
                        <th className="p-2 border-r border-gray-200 text-center">ប្រភេទទិដ្ឋាការ</th>
                      )}
                      <th className="p-2 border-r border-gray-200 text-center">
                        {stockType === 'sticker' ? 'ចំនួនសន្លឹក' : 'ចំនួនដុំ'}
                      </th>
                      {stockType === 'sticker' && (
                        <>
                          <th className="p-2 border-r border-gray-200 text-center">ចាប់ពី</th>
                          <th className="p-2 border-r border-gray-200 text-center">ដល់លេខ</th>
                        </>
                      )}
                      <th className="p-2 border-r border-gray-200">ក្រុមរបក</th>
                      <th className="p-2 border-r border-gray-200">អ្នកស្នើសុំ</th>
                      <th className="p-2">អ្នកមកបើក</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-200 bg-white">
                    {displayedRecords.length === 0 ? (
                      <tr>
                        <td colSpan={stockType === 'sticker' ? 10 : 8} className="p-4 text-center text-gray-500">
                          {duplicateFilterView === 'duplicate'
                            ? 'គ្មានទិន្នន័យស្ទួនក្នុងប្រភេទប្រតិបត្តិការនេះទេ! 🎉'
                            : 'គ្មានទិន្នន័យសម្រាប់បង្ហាញទេ!'}
                        </td>
                      </tr>
                    ) : (
                      displayedRecords.slice(0, 20).map((rec, i) => (
                        <tr
                          key={i}
                          className={`transition ${
                            rec.isDuplicate
                              ? 'bg-amber-50/90 hover:bg-amber-100/80 text-amber-950 font-medium'
                              : 'hover:bg-gray-50'
                          }`}
                        >
                          <td className="p-2 border-r border-gray-200 text-center font-medium text-gray-500">
                            {rec.rowNumber || i + 1}
                          </td>
                          <td className="p-2 border-r border-gray-200 text-center">
                            {rec.isDuplicate ? (
                              <div className="flex flex-col items-center">
                                <span className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded text-[10px] font-bold bg-amber-200 text-amber-950 border border-amber-400">
                                  <AlertTriangle className="w-3 h-3 text-amber-800" />
                                  <span>ស្ទួន</span>
                                </span>
                                <span
                                  className="text-[9px] text-amber-800 font-medium leading-tight mt-0.5 max-w-[130px] truncate"
                                  title={rec.duplicateReason}
                                >
                                  {rec.duplicateReason}
                                </span>
                              </div>
                            ) : (
                              <span className="inline-flex items-center gap-1 px-1.5 py-0.5 rounded text-[10px] font-bold bg-emerald-100 text-emerald-800 border border-emerald-300">
                                <CheckCircle2 className="w-3 h-3 text-emerald-600" />
                                <span>ថ្មី</span>
                              </span>
                            )}
                          </td>
                          <td className="p-2 border-r border-gray-200">
                            <span
                              className={`px-1.5 py-0.5 rounded text-[10px] font-bold ${
                                rec.operationType === 'testPrintK2'
                                  ? 'bg-amber-100 text-amber-900 border border-amber-300'
                                  : rec.operationType === 'damaged'
                                  ? 'bg-rose-100 text-rose-800 border border-rose-300'
                                  : rec.operationType === 'damagedTeam'
                                  ? 'bg-orange-100 text-orange-800 border border-orange-300'
                                  : rec.operationType === 'missingTeam'
                                  ? 'bg-red-100 text-red-800 border border-red-300'
                                  : rec.operationType === 'returnTeam'
                                  ? 'bg-cyan-100 text-cyan-800 border border-cyan-300'
                                  : rec.operationType === 'openK1'
                                  ? rec.sourceFrom?.includes('ស្តុកចាស់')
                                    ? 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                                    : 'bg-blue-100 text-blue-800'
                                  : rec.operationType === 'useTeam'
                                  ? 'bg-purple-100 text-purple-800'
                                  : rec.sourceFrom?.includes('ស្តុកចាស់')
                                  ? 'bg-teal-100 text-teal-800 border border-teal-300'
                                  : 'bg-amber-100 text-amber-800'
                              }`}
                            >
                              {rec.operationType === 'testPrintK2'
                                ? 'ទិដ្ឋាការសាកក២'
                                : rec.operationType === 'damaged'
                                ? 'ទិដ្ឋាការខូចក២'
                                : rec.operationType === 'damagedTeam'
                                ? 'ទិដ្ឋាការខូចក្រុម'
                                : rec.operationType === 'missingTeam'
                                ? 'ទិដ្ឋាការខ្វះក្រុម'
                                : rec.operationType === 'returnTeam'
                                ? 'ទិដ្ឋាការបង្វិលពីក្រុម'
                                : rec.operationType === 'openK1'
                                ? rec.sourceFrom?.includes('ស្តុកចាស់')
                                  ? 'ស្តុកចាស់ ក២'
                                  : 'ការបញ្ចូលស្តុក'
                                : rec.operationType === 'useTeam'
                                ? 'ការប្រើប្រាស់តាមក្រុម'
                                : rec.operationType === 'issueTeam'
                                ? 'ការបើកផ្តល់តាមក្រុម'
                                : rec.sourceFrom?.includes('ស្តុកចាស់')
                                ? 'ស្តុកចាស់របស់ក្រុម'
                                : rec.operationType || 'ការបើកផ្តល់តាមក្រុម'}
                            </span>
                          </td>
                          <td className="p-2 border-r border-gray-200 text-gray-700 font-medium">
                            {rec.date} {rec.time ? `(${rec.time})` : ''}
                          </td>
                          {stockType === 'sticker' && (
                            <td className="p-2 border-r border-gray-200 text-center font-bold text-blue-700">
                              {rec.visaType || '-'}
                            </td>
                          )}
                          <td className="p-2 border-r border-gray-200 text-center font-bold text-emerald-700">
                            {rec.quantityBundles}
                          </td>
                          {stockType === 'sticker' && (
                            <>
                              <td className="p-2 border-r border-gray-200 text-center font-mono">
                                {rec.startSerial || '-'}
                              </td>
                              <td className="p-2 border-r border-gray-200 text-center font-mono">
                                {rec.endSerial || '-'}
                              </td>
                            </>
                          )}
                          <td className="p-2 border-r border-gray-200 text-gray-800 font-medium">
                            {rec.visaTeamRobokName || '-'}
                          </td>
                          <td className="p-2 border-r border-gray-200 text-gray-700">
                            {rec.requestedRankName ? `${rec.requestedRankName} ` : ''}
                            {rec.requesterName || '-'}
                          </td>
                          <td className="p-2 text-gray-700">
                            {rec.collectorRoleName ? `${rec.collectorRoleName} ` : ''}
                            {rec.collectorName || '-'}
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
              {displayedRecords.length > 20 && (
                <p className="text-[10px] text-gray-500 text-center">
                  ... និងទិន្នន័យ {displayedRecords.length - 20} ជួរផ្សេងទៀត
                </p>
              )}
            </div>
          )}
        </div>

        {/* Footer Actions */}
        <div className="bg-gray-50 px-5 py-3 border-t border-gray-200 flex items-center justify-between shrink-0">
          <button
            type="button"
            onClick={onClose}
            className="px-4 py-2 rounded text-xs font-bold text-gray-600 hover:bg-gray-200 transition cursor-pointer"
          >
            បោះបង់
          </button>

          <button
            type="button"
            onClick={handleConfirmImport}
            disabled={recordsToImport.length === 0}
            className="bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white px-5 py-2 rounded text-xs font-bold flex items-center gap-2 shadow-2xs transition cursor-pointer"
          >
            <CheckCircle2 className="w-4 h-4" />
            <span>
              រក្សាទុកទិន្នន័យ ({recordsToImport.length} ជួរ)
              {skipDuplicates && totalDuplicatesCount > 0 && ` [រំលងស្ទួន ${parsedRecords.length - recordsToImport.length}]`}
            </span>
          </button>
        </div>
      </div>
    </div>
  );
};

export interface VisaImportModalProps {
  isOpen: boolean;
  onClose: () => void;
  onImportRecords: (records: Omit<VisaRecord, 'id' | 'createdAt'>[]) => void;
  onShowToast: (msg: string, type?: 'success' | 'error' | 'info') => void;
  existingRecords?: VisaRecord[];
  visaTeams?: CategoryItem[];
  organizations?: CategoryItem[];
  officers?: Officer[];
}

export interface ParsedVisaRecordWithDuplicate extends Omit<VisaRecord, 'id' | 'createdAt'> {
  isDuplicate?: boolean;
  duplicateReason?: string;
  rowNumber?: number;
}

export const VisaImportModal: React.FC<VisaImportModalProps> = ({
  isOpen,
  onClose,
  onImportRecords,
  onShowToast,
  existingRecords = [],
  visaTeams = [],
  organizations = [],
  officers = [],
}) => {
  const [file, setFile] = useState<File | null>(null);
  const [parsedRecords, setParsedRecords] = useState<ParsedVisaRecordWithDuplicate[]>([]);
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [skipDuplicates, setSkipDuplicates] = useState<boolean>(false);
  const [filterView, setFilterView] = useState<'all' | 'valid' | 'duplicate'>('all');

  useEffect(() => {
    if (isOpen) {
      setSkipDuplicates(false);
      setFilterView('all');
    }
  }, [isOpen]);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const uploadedFile = e.target.files?.[0];
    if (!uploadedFile) return;

    setFile(uploadedFile);
    setIsProcessing(true);

    const reader = new FileReader();
    reader.onload = (evt) => {
      setTimeout(() => {
        try {
          const buffer = evt.target?.result;
          const workbook = XLSX.read(buffer, { type: 'binary', cellDates: false });
          const sheetName = workbook.SheetNames[0];
          const sheet = workbook.Sheets[sheetName];
          const rawData = XLSX.utils.sheet_to_json<any>(sheet, { blankrows: false });

          // Pre-build O(1) Lookup Map for existing visa records
          const existingVisaMap = new Map<string, VisaRecord>();
          if (existingRecords && existingRecords.length > 0) {
            existingRecords.forEach((ex) => {
              const pass = (ex.passportNumber || '').trim().toLowerCase();
              const dt = ex.applicationDate || '';
              const vType = (ex.newVisaType || '').trim().toLowerCase();
              const name = (ex.fullName || '').trim().toLowerCase();
              const ext = (ex.extension || '').trim().toLowerCase();
              const key = `${pass}|${dt}|${vType}|${name}|${ext}`;
              if (!existingVisaMap.has(key)) {
                existingVisaMap.set(key, ex);
              }
            });
          }

          // In-file O(1) Map
          const fileVisaMap = new Map<string, number>();

          const records: ParsedVisaRecordWithDuplicate[] = [];

          rawData.forEach((row, rIdx) => {
            const passportNumber = String(row['លេខលិខិតឆ្លងដែន'] || row['Passport'] || row['passportNumber'] || row['passportNo'] || '').trim();
            if (!passportNumber) return;

            const rawDateVal = row['កាលបរិច្ឆេទ'] || row['Date'] || row['DateApplication'] || row['applicationDate'] || row['ថ្ងៃខែ'] || row['ថ្ងៃទី'];
            const appDate = parseExcelDate(rawDateVal);
            const newVisaType = String(row['ប្រភេទទិដ្ឋាការ'] || row['VisaType'] || 'E').trim();
            const fullName = String(row['ឈ្មោះអ្នកកាន់លិខិតឆ្លងដែន'] || row['FullName'] || '').trim();
            const extension = String(row['ពន្យារ'] || row['Extension'] || '1M').trim();

            const rowNum = rIdx + 2;
            const lookupKey = `${passportNumber.toLowerCase()}|${appDate}|${newVisaType.toLowerCase()}|${fullName.toLowerCase()}|${extension.toLowerCase()}`;

            // Duplicate detection via O(1) Map
            let isDup = false;
            let dupReason = '';

            // 1. Check against existing database records
            const systemDup = existingVisaMap.get(lookupKey);
            if (systemDup) {
              isDup = true;
              dupReason = `ស្ទួនក្នុងប្រព័ន្ធ (ថ្ងៃ ${systemDup.applicationDate}, លិខិតឆ្លងដែន ${systemDup.passportNumber})`;
            }

            // 2. Check against previous rows in this file
            if (!isDup) {
              const priorRowNum = fileVisaMap.get(lookupKey);
              if (priorRowNum !== undefined) {
                isDup = true;
                dupReason = `ស្ទួនជាមួយជួរទី ${priorRowNum} ក្នុង File នេះ`;
              }
            }

            if (!fileVisaMap.has(lookupKey)) {
              fileVisaMap.set(lookupKey, rowNum);
            }

            records.push({
              applicationDate: appDate,
              passportNumber,
              fullName: String(row['ឈ្មោះអ្នកកាន់លិខិតឆ្លងដែន'] || row['FullName'] || '').trim(),
              gender: String(row['ភេទ'] || row['Sex'] || 'M').trim(),
              nationality: String(row['សញ្ជាតិ'] || row['Nationality'] || '').trim(),
              newVisaType,
              extension: String(row['ពន្យារ'] || row['Extension'] || '1M').trim(),
              fee: Number(row['តម្លៃ'] || row['Fee'] || 50),
              organization: String(row['អង្គភាព'] || row['Organization'] || organizations[0]?.name || 'អគ្គនាយកដ្ឋានអន្តោប្រវេសន៍').trim(),
              officerName: String(row['អ្នកប្រមូល'] || row['CollectorName'] || officers[0]?.name || '').trim(),
              notes: String(row['កំណត់សម្គាល់'] || row['Remarks'] || '').trim(),
              status: 'អនុម័តរួច',
              isDuplicate: isDup,
              duplicateReason: dupReason,
              rowNumber: rowNum,
            });
          });

          setParsedRecords(records);
          const dupCount = records.filter((r) => r.isDuplicate).length;
          if (dupCount > 0) {
            onShowToast(`បានអាន ${records.length} ជួរ (រកឃើញទិន្នន័យស្ទួន ${dupCount} ជួរ)`, 'info');
          } else {
            onShowToast(`បានអានទិន្នន័យចំនួន ${records.length} ជួរ`, 'success');
          }
        } catch (err) {
          console.error('Failed to parse file:', err);
          onShowToast('មានកំហុសក្នុងការអានឯកសារ Excel', 'error');
        } finally {
          setIsProcessing(false);
        }
      }, 30);
    };

    reader.readAsBinaryString(uploadedFile);
  };

  const totalDuplicates = useMemo(() => parsedRecords.filter((r) => r.isDuplicate).length, [parsedRecords]);
  const totalUnique = useMemo(() => parsedRecords.filter((r) => !r.isDuplicate).length, [parsedRecords]);

  const displayedRecords = useMemo(() => {
    if (filterView === 'valid') return parsedRecords.filter((r) => !r.isDuplicate);
    if (filterView === 'duplicate') return parsedRecords.filter((r) => r.isDuplicate);
    return parsedRecords;
  }, [parsedRecords, filterView]);

  const recordsToImport = useMemo(() => {
    if (skipDuplicates) {
      return parsedRecords.filter((r) => !r.isDuplicate);
    }
    return parsedRecords;
  }, [parsedRecords, skipDuplicates]);

  const handleConfirm = () => {
    if (recordsToImport.length === 0) {
      onShowToast('គ្មានទិន្នន័យសម្រាប់នាំចូលទេ!', 'error');
      return;
    }
    onImportRecords(recordsToImport);
    const skippedCount = parsedRecords.length - recordsToImport.length;
    const skipText = skippedCount > 0 ? ` (បានរំលងស្ទួន ${skippedCount} ជួរ)` : '';
    onShowToast(`បាននាំចូលទិន្នន័យចំនួន ${recordsToImport.length} ជោគជ័យ${skipText}`, 'success');
    onClose();
  };

  if (!isOpen) return null;

  return (
    <div className="fixed inset-0 z-50 bg-black/60 backdrop-blur-xs flex items-center justify-center p-4">
      <div className="bg-white rounded-xl shadow-2xl max-w-2xl w-full overflow-hidden flex flex-col max-h-[90vh]">
        <div className="bg-emerald-800 text-white px-5 py-3.5 flex items-center justify-between shrink-0">
          <div className="flex items-center gap-2 font-bold text-sm">
            <FileSpreadsheet className="w-5 h-5 text-amber-300" />
            <span>នាំចូលទិន្នន័យទិដ្ឋាការ (Excel Import)</span>
          </div>
          <button onClick={onClose} className="p-1 hover:bg-emerald-700 rounded text-emerald-100 transition cursor-pointer">
            <X className="w-5 h-5" />
          </button>
        </div>

        <div className="p-5 space-y-4 overflow-y-auto grow text-xs">
          <div className="border-2 border-dashed border-gray-300 rounded-lg p-5 text-center hover:border-emerald-500 transition relative bg-gray-50/50">
            <input
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={handleFileUpload}
              className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
            />
            <Upload className="w-8 h-8 text-emerald-600 mx-auto mb-1.5" />
            <p className="font-semibold text-gray-700">ចុច ឬទម្លាក់ឯកសារ Excel នៅទីនេះ</p>
            <p className="text-[11px] text-gray-500 mt-0.5">គាំទ្រត្រឹម .xlsx, .xls, .csv</p>
            {file && <p className="text-xs text-emerald-600 font-bold mt-1.5">ឯកសារ៖ {file.name}</p>}
          </div>

          {isProcessing && <p className="text-center text-amber-600 font-semibold animate-pulse">កំពុងដំណើរការអានឯកសារ...</p>}

          {/* Duplicate Warning Box */}
          {parsedRecords.length > 0 && totalDuplicates > 0 && (
            <div className="bg-amber-50 border border-amber-300 rounded-lg p-3 space-y-2 text-xs">
              <div className="flex items-center justify-between">
                <span className="font-bold text-amber-950 flex items-center gap-1.5">
                  <AlertTriangle className="w-4 h-4 text-amber-600" />
                  <span>រកឃើញទិន្នន័យស្ទួន <strong className="text-red-700">{totalDuplicates}</strong> ជួរ</span>
                </span>
                <span className="text-[11px] text-amber-800">
                  (ទិន្នន័យថ្មី: <strong className="text-emerald-700">{totalUnique}</strong> / សរុប: {parsedRecords.length})
                </span>
              </div>
              <label className="flex items-center gap-2 cursor-pointer font-bold text-amber-950 bg-white p-1.5 rounded border border-amber-200">
                <input
                  type="checkbox"
                  checked={skipDuplicates}
                  onChange={(e) => setSkipDuplicates(e.target.checked)}
                  className="w-4 h-4 text-emerald-600 rounded"
                />
                <span className="flex items-center gap-1">
                  <ShieldCheck className="w-3.5 h-3.5 text-emerald-600" />
                  <span>រំលងទិន្នន័យស្ទួនដោយស្វ័យប្រវត្តិ (Skip Duplicates)</span>
                </span>
              </label>
            </div>
          )}

          {parsedRecords.length > 0 && (
            <div className="space-y-2">
              <div className="flex items-center justify-between">
                <span className="font-bold text-gray-800">
                  ទិន្នន័យដែលបានអាន (សរុប {parsedRecords.length} ជួរ)
                </span>
                {totalDuplicates > 0 && (
                  <div className="inline-flex rounded border border-gray-300 p-0.5 bg-gray-100 text-[11px]">
                    <button
                      type="button"
                      onClick={() => setFilterView('all')}
                      className={`px-2 py-0.5 rounded font-bold ${filterView === 'all' ? 'bg-white text-gray-900 shadow-xs' : 'text-gray-600'}`}
                    >
                      ទាំងអស់ ({parsedRecords.length})
                    </button>
                    <button
                      type="button"
                      onClick={() => setFilterView('valid')}
                      className={`px-2 py-0.5 rounded font-bold ${filterView === 'valid' ? 'bg-emerald-600 text-white shadow-xs' : 'text-emerald-700'}`}
                    >
                      ✅ ថ្មី ({totalUnique})
                    </button>
                    <button
                      type="button"
                      onClick={() => setFilterView('duplicate')}
                      className={`px-2 py-0.5 rounded font-bold ${filterView === 'duplicate' ? 'bg-amber-600 text-white shadow-xs' : 'text-amber-800'}`}
                    >
                      ⚠️ ស្ទួន ({totalDuplicates})
                    </button>
                  </div>
                )}
              </div>

              <div className="overflow-x-auto max-h-48 border border-gray-200 rounded">
                <table className="w-full text-left text-[11px] border-collapse">
                  <thead className="bg-gray-100 sticky top-0 font-bold text-gray-700 border-b border-gray-200">
                    <tr>
                      <th className="p-2 border-r text-center w-8">ល.រ</th>
                      <th className="p-2 border-r text-center w-20">ស្ថានភាព</th>
                      <th className="p-2 border-r">លិខិតឆ្លងដែន</th>
                      <th className="p-2 border-r">ឈ្មោះ</th>
                      <th className="p-2 border-r text-center">ប្រភេទ</th>
                      <th className="p-2 border-r">កាលបរិច្ឆេទ</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-gray-200 bg-white">
                    {displayedRecords.map((r, i) => (
                      <tr key={i} className={r.isDuplicate ? 'bg-amber-50' : 'hover:bg-gray-50'}>
                        <td className="p-2 border-r text-center text-gray-500">{i + 1}</td>
                        <td className="p-2 border-r text-center">
                          {r.isDuplicate ? (
                            <span className="px-1 py-0.5 rounded text-[9px] font-bold bg-amber-200 text-amber-900">
                              ស្ទួន
                            </span>
                          ) : (
                            <span className="px-1 py-0.5 rounded text-[9px] font-bold bg-emerald-100 text-emerald-800">
                              ថ្មី
                            </span>
                          )}
                        </td>
                        <td className="p-2 border-r font-mono font-bold text-blue-700">{r.passportNumber}</td>
                        <td className="p-2 border-r">{r.fullName || '-'}</td>
                        <td className="p-2 border-r text-center font-bold">{r.newVisaType} ({r.extension})</td>
                        <td className="p-2 font-mono text-gray-600">{r.applicationDate}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </div>

        <div className="bg-gray-50 px-5 py-3 border-t border-gray-200 flex items-center justify-between shrink-0">
          <button onClick={onClose} className="px-4 py-2 rounded text-xs font-bold text-gray-600 hover:bg-gray-200 transition">
            បោះបង់
          </button>
          <button
            onClick={handleConfirm}
            disabled={recordsToImport.length === 0}
            className="bg-emerald-600 hover:bg-emerald-700 disabled:opacity-50 text-white px-5 py-2 rounded text-xs font-bold flex items-center gap-2 transition cursor-pointer"
          >
            <CheckCircle2 className="w-4 h-4" />
            <span>រក្សាទុកទិន្នន័យ ({recordsToImport.length})</span>
          </button>
        </div>
      </div>
    </div>
  );
};
