import React, { useState, useRef, useMemo } from 'react';
import { VisaRecord } from '../types';
import { CalendarSearch, Search, Printer, CheckCircle2, Clock, Eye, AlertCircle, FileSpreadsheet, FileText, Loader2, Home } from 'lucide-react';
import * as XLSX from 'xlsx';
import html2canvas from 'html2canvas';
import jsPDF from 'jspdf';
import { sanitizeDocumentForHtml2Canvas } from '../utils/pdfExportHelper';
import { printA4Document } from '../utils/printHelper';
import { CustomDatePicker } from './CustomDatePicker';

interface VisaDateSearchProps {
  records: VisaRecord[];
}

export const VisaDateSearch: React.FC<VisaDateSearchProps> = ({ records }) => {
  const todayStr = new Date().toISOString().split('T')[0];

  // Input states (change as user types or picks date)
  const [fromDate, setFromDate] = useState(todayStr);
  const [toDate, setToDate] = useState(todayStr);
  const [keyword, setKeyword] = useState('');

  // Applied filter states (only update when user clicks 'ស្វែងរក' / Search)
  const [appliedFromDate, setAppliedFromDate] = useState(todayStr);
  const [appliedToDate, setAppliedToDate] = useState(todayStr);
  const [appliedKeyword, setAppliedKeyword] = useState('');
  const [hasSearched, setHasSearched] = useState(false);

  const [isExportingPDF, setIsExportingPDF] = useState(false);

  const reportRef = useRef<HTMLDivElement>(null);

  const handleSearchSubmit = (e?: React.FormEvent) => {
    if (e) e.preventDefault();
    setAppliedFromDate(fromDate);
    setAppliedToDate(toDate);
    setAppliedKeyword(keyword);
    setHasSearched(true);
  };

  const handleClearFilters = () => {
    setFromDate('');
    setToDate('');
    setKeyword('');
    setAppliedFromDate('');
    setAppliedToDate('');
    setAppliedKeyword('');
    setHasSearched(false);
  };

  // Preset quick filters
  const handlePreset = (preset: 'today' | 'week' | 'month' | 'year' | 'all') => {
    const now = new Date();
    const today = now.toISOString().split('T')[0];

    let newFrom = '';
    let newTo = '';

    if (preset === 'today') {
      newFrom = today;
      newTo = today;
    } else if (preset === 'week') {
      const firstDayOfWeek = new Date(now);
      firstDayOfWeek.setDate(now.getDate() - now.getDay());
      newFrom = firstDayOfWeek.toISOString().split('T')[0];
      newTo = today;
    } else if (preset === 'month') {
      const firstDayOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
      newFrom = firstDayOfMonth.toISOString().split('T')[0];
      newTo = today;
    } else if (preset === 'year') {
      const firstDayOfYear = new Date(now.getFullYear(), 0, 1);
      newFrom = firstDayOfYear.toISOString().split('T')[0];
      newTo = today;
    } else if (preset === 'all') {
      newFrom = '';
      newTo = '';
    }

    setFromDate(newFrom);
    setToDate(newTo);
    setAppliedFromDate(newFrom);
    setAppliedToDate(newTo);
    setHasSearched(true);
  };

  // Filter records by APPLIED date range and keyword (not live on date input change)
  const filteredRecords = records.filter((r) => {
    let inRange = true;
    if (appliedFromDate) {
      inRange = inRange && r.applicationDate >= appliedFromDate;
    }
    if (appliedToDate) {
      inRange = inRange && r.applicationDate <= appliedToDate;
    }

    let matchesKeyword = true;
    if (appliedKeyword.trim()) {
      const q = appliedKeyword.toLowerCase().trim();
      matchesKeyword =
        r.passportNumber.toLowerCase().includes(q) ||
        r.fullName.toLowerCase().includes(q) ||
        r.nationality.toLowerCase().includes(q);
    }

    return inRange && matchesKeyword;
  });

  const rangeCount = filteredRecords.length;
  const rangeApproved = filteredRecords.filter((r) => r.status === 'អនុម័តរួច').length;

  const rangeUniqueLetters = useMemo(() => {
    const uniqueKeys = new Set<string>();
    let noDocCount = 0;

    filteredRecords.forEach((r) => {
      const docNum = (r.documentNumber || '').trim();
      const issueDt = (r.issueDate || r.applicationDate || '').trim();

      if (docNum) {
        uniqueKeys.add(`${docNum}___${issueDt}`);
      } else {
        noDocCount += 1;
      }
    });

    return uniqueKeys.size + noDocCount;
  }, [filteredRecords]);

  const toKhmerDigits = (str: string) => {
    const khmerNums = ['០', '១', '២', '៣', '៤', '៥', '៦', '៧', '៨', '៩'];
    return str.replace(/[0-9]/g, (w) => khmerNums[parseInt(w, 10)]);
  };

  const khmerMonths = [
    'មករា', 'កុម្ភៈ', 'មីនា', 'មេសា', 'ឧសភា', 'មិថុនា',
    'កក្កដា', 'សីហា', 'កញ្ញា', 'តុលា', 'វិច្ឆិកា', 'ធ្នូ'
  ];

  const formatKhmerDateLong = (dateStr: string) => {
    if (!dateStr) return '';
    const parts = dateStr.split('-');
    if (parts.length !== 3) return dateStr;
    const year = toKhmerDigits(parts[0]);
    const monthIdx = parseInt(parts[1], 10) - 1;
    const monthName = khmerMonths[monthIdx] || parts[1];
    const day = toKhmerDigits(parts[2]);
    return `ថ្ងៃទី${day} ខែ${monthName} ឆ្នាំ${year}`;
  };

  const getDateRangeSubtitle = () => {
    if (appliedFromDate && appliedToDate) {
      if (appliedFromDate === appliedToDate) {
        return `ដោយគិតត្រឹម ${formatKhmerDateLong(appliedFromDate)} ( ${appliedFromDate} ដល់ ${appliedToDate} )`;
      }
      return `ដោយគិតចាប់ពី ${formatKhmerDateLong(appliedFromDate)} រហូតដល់ ${formatKhmerDateLong(appliedToDate)}`;
    } else if (appliedFromDate) {
      return `ដោយគិតចាប់ពី ${formatKhmerDateLong(appliedFromDate)}`;
    } else if (appliedToDate) {
      return `ដោយគិតត្រឹម ${formatKhmerDateLong(appliedToDate)}`;
    }
    return `លទ្ធផលទិន្នន័យទាំងអស់`;
  };

  const handlePrint = () => {
    printA4Document(reportRef.current || 'visa-date-search-report', {
      orientation: 'landscape',
      documentTitle: `បញ្ជីប្តូរប្រភេទទិដ្ឋាការ_${fromDate || 'all'}_${toDate || 'all'}`,
    });
  };

  const handleExportExcel = async () => {
    if (filteredRecords.length === 0) {
      alert('មិនមានទិន្នន័យសម្រាប់ទាញយកឡើយ!');
      return;
    }

    try {
      const ExcelJS = await import('exceljs');
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('បញ្ជីទិដ្ឋាការ');

      // 1. Title Row (A1:L1)
      worksheet.mergeCells('A1:L1');
      const titleCell = worksheet.getCell('A1');
      titleCell.value = 'បញ្ជីប្តូរប្រភេទទិដ្ឋាការ និងពន្យារទិដ្ឋាការ';
      titleCell.font = { name: 'Khmer OS Siemreap', size: 14, bold: true };
      titleCell.alignment = { vertical: 'middle', horizontal: 'center' };

      // 2. Subtitle Row (A2:L2)
      worksheet.mergeCells('A2:L2');
      const subtitleCell = worksheet.getCell('A2');
      subtitleCell.value = getDateRangeSubtitle();
      subtitleCell.font = { name: 'Khmer OS Siemreap', size: 12 };
      subtitleCell.alignment = { vertical: 'middle', horizontal: 'center' };

      // 3. Headers at Row 3 (A3:L3)
      const headers = [
        'ល.រ',
        'ថ្ងៃខែ',
        'ឈ្មោះ',
        'ភេទ',
        'លិខិតឆ្លងដែន',
        'សញ្ជាតិ',
        'ទិដ្ឋាការដំបូង',
        'ប្តូរ',
        'លេខលិខិត',
        'ថ្ងៃចេញ',
        'ស្ថាប័នដើម',
        'ពន្យារ'
      ];

      const headerRow = worksheet.getRow(3);
      headers.forEach((header, colIdx) => {
        const cell = headerRow.getCell(colIdx + 1);
        cell.value = header;
        cell.font = { name: 'Khmer OS Siemreap', size: 12, bold: false };
        cell.alignment = {
          vertical: 'middle',
          horizontal: [0, 3, 6, 7, 11].includes(colIdx) ? 'center' : 'left'
        };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' }
        };
      });

      // 4. Data rows starting at Row 4
      filteredRecords.forEach((r, idx) => {
        const rowNum = idx + 4;
        const row = worksheet.getRow(rowNum);
        const rowValues = [
          idx + 1,
          r.applicationDate,
          r.fullName,
          r.gender === 'ប្រុស' ? 'M' : r.gender === 'ស្រី' ? 'F' : r.gender,
          r.passportNumber,
          r.nationality,
          r.oldVisaType ? r.oldVisaType.split(' ')[0] : '-',
          r.newVisaType ? r.newVisaType.split(' ')[0] : '-',
          r.documentNumber || '-',
          r.issueDate || r.applicationDate,
          r.organization || '-',
          r.extension || r.duration
        ];

        rowValues.forEach((val, colIdx) => {
          const cell = row.getCell(colIdx + 1);
          cell.value = val;
          cell.font = { name: 'Khmer OS Siemreap', size: 11 };
          cell.alignment = {
            vertical: 'middle',
            horizontal: [0, 3, 6, 7, 11].includes(colIdx) ? 'center' : 'left'
          };
          cell.border = {
            top: { style: 'thin' },
            left: { style: 'thin' },
            bottom: { style: 'thin' },
            right: { style: 'thin' }
          };
        });
      });

      // 5. Column widths matching the image layout
      worksheet.columns = [
        { width: 6 },  // A: ល.រ
        { width: 14 }, // B: ថ្ងៃខែ
        { width: 22 }, // C: ឈ្មោះ
        { width: 8 },  // D: ភេទ
        { width: 16 }, // E: លិខិតឆ្លងដែន
        { width: 14 }, // F: សញ្ជាតិ
        { width: 16 }, // G: ទិដ្ឋាការដំបូង
        { width: 10 }, // H: ប្តូរ
        { width: 16 }, // I: លេខលិខិត
        { width: 14 }, // J: ថ្ងៃចេញ
        { width: 26 }, // K: ស្ថាប័នដើម
        { width: 10 }, // L: ពន្យារ
      ];

      // 6. Write buffer and initiate download
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = window.URL.createObjectURL(blob);
      const anchor = document.createElement('a');
      const dateSuffix = fromDate && toDate ? `${fromDate}_to_${toDate}` : 'all';
      anchor.href = url;
      anchor.download = `Visa_Report_${dateSuffix}.xlsx`;
      anchor.click();
      window.URL.revokeObjectURL(url);
    } catch (err) {
      console.error('Excel export error:', err);
      alert('មានបញ្ហាក្នុងការទាញយក Excel');
    }
  };

  const handleExportPDF = async () => {
    if (filteredRecords.length === 0) {
      alert('មិនមានទិន្នន័យសម្រាប់បង្កើត PDF ឡើយ!');
      return;
    }

    setIsExportingPDF(true);

    try {
      // 1. Try direct jsPDF export with offscreen clean DOM rendering
      const offscreenDiv = document.createElement('div');
      offscreenDiv.style.position = 'absolute';
      offscreenDiv.style.left = '-9999px';
      offscreenDiv.style.top = '-9999px';
      offscreenDiv.style.width = '1200px';
      offscreenDiv.style.backgroundColor = '#ffffff';
      offscreenDiv.style.padding = '20px';
      offscreenDiv.style.fontFamily = "'Kantumruy Pro', Khmer, sans-serif";

      const rowsHtml = filteredRecords.map((r, idx) => `
        <tr style="border-bottom: 1px solid #e2e8f0; font-size: 11px;">
          <td style="padding: 6px; text-align: center; border: 1px solid #cbd5e1;">${idx + 1}</td>
          <td style="padding: 6px; border: 1px solid #cbd5e1; white-space: nowrap;">${r.applicationDate}</td>
          <td style="padding: 6px; border: 1px solid #cbd5e1; font-weight: bold; color: #0B2545;">${r.fullName}</td>
          <td style="padding: 6px; border: 1px solid #cbd5e1; text-align: center;">${r.gender === 'ប្រុស' ? 'M' : r.gender === 'ស្រី' ? 'F' : r.gender}</td>
          <td style="padding: 6px; border: 1px solid #cbd5e1; font-weight: bold;">${r.passportNumber}</td>
          <td style="padding: 6px; border: 1px solid #cbd5e1;">${r.nationality}</td>
          <td style="padding: 6px; border: 1px solid #cbd5e1; text-align: center;">${r.oldVisaType ? r.oldVisaType.split(' ')[0] : '-'}</td>
          <td style="padding: 6px; border: 1px solid #cbd5e1; text-align: center; font-weight: bold; color: #047857;">${r.newVisaType ? r.newVisaType.split(' ')[0] : '-'}</td>
          <td style="padding: 6px; border: 1px solid #cbd5e1;">${r.documentNumber || '-'}</td>
          <td style="padding: 6px; border: 1px solid #cbd5e1; white-space: nowrap;">${r.issueDate || r.applicationDate}</td>
          <td style="padding: 6px; border: 1px solid #cbd5e1;">${r.organization || '-'}</td>
          <td style="padding: 6px; border: 1px solid #cbd5e1; text-align: center; font-weight: bold; color: #b45309;">${r.extension || r.duration}</td>
        </tr>
      `).join('');

      offscreenDiv.innerHTML = `
        <div style="text-align: center; margin-bottom: 15px; background: #0B2545; color: #E4CD98; padding: 14px; border-radius: 8px;">
          <h2 style="margin: 0; font-size: 18px; font-weight: bold;">បញ្ជីប្តូរប្រភេទទិដ្ឋាការ និងពន្យារទិដ្ឋាការ</h2>
          <p style="margin: 4px 0 0 0; font-size: 13px; color: #ffffff;">${getDateRangeSubtitle()}</p>
        </div>
        <table style="width: 100%; border-collapse: collapse; font-size: 11px; text-align: left;">
          <thead>
            <tr style="background-color: #f1f5f9; color: #0B2545;">
              <th style="padding: 8px 6px; border: 1px solid #cbd5e1; text-align: center;">លេខ</th>
              <th style="padding: 8px 6px; border: 1px solid #cbd5e1;">ថ្ងៃខែ</th>
              <th style="padding: 8px 6px; border: 1px solid #cbd5e1;">ឈ្មោះ</th>
              <th style="padding: 8px 6px; border: 1px solid #cbd5e1; text-align: center;">ភេទ</th>
              <th style="padding: 8px 6px; border: 1px solid #cbd5e1;">លិខិតឆ្លងដែន</th>
              <th style="padding: 8px 6px; border: 1px solid #cbd5e1;">សញ្ជាតិ</th>
              <th style="padding: 8px 6px; border: 1px solid #cbd5e1; text-align: center;">ទិដ្ឋាការដំបូង</th>
              <th style="padding: 8px 6px; border: 1px solid #cbd5e1; text-align: center;">ប្តូរ</th>
              <th style="padding: 8px 6px; border: 1px solid #cbd5e1;">លេខលិខិត</th>
              <th style="padding: 8px 6px; border: 1px solid #cbd5e1;">ថ្ងៃចេញ</th>
              <th style="padding: 8px 6px; border: 1px solid #cbd5e1;">ស្ថាប័នដើម</th>
              <th style="padding: 8px 6px; border: 1px solid #cbd5e1; text-align: center;">ពន្យារ</th>
            </tr>
          </thead>
          <tbody>
            ${rowsHtml}
          </tbody>
        </table>
        <div style="margin-top: 15px; text-align: right; font-size: 11px; color: #475569;">
          សរុប៖ <strong>${filteredRecords.length}</strong> ករណី
        </div>
      `;

      document.body.appendChild(offscreenDiv);

      const canvas = await html2canvas(offscreenDiv, {
        scale: 2,
        useCORS: true,
        allowTaint: true,
        backgroundColor: '#ffffff',
        windowWidth: 1200,
        onclone: (clonedDoc) => {
          sanitizeDocumentForHtml2Canvas(clonedDoc);
        },
      });

      document.body.removeChild(offscreenDiv);

      const imgData = canvas.toDataURL('image/png');
      const pdf = new jsPDF({
        orientation: 'landscape',
        unit: 'mm',
        format: 'a4',
      });

      const pdfWidth = pdf.internal.pageSize.getWidth();
      const pdfHeight = pdf.internal.pageSize.getHeight();
      const imgWidth = 280; // Margin 8.5mm left & right
      const imgHeight = (canvas.height * imgWidth) / canvas.width;

      let heightLeft = imgHeight;
      let position = 8;

      pdf.addImage(imgData, 'PNG', 8.5, position, imgWidth, imgHeight);
      heightLeft -= (pdfHeight - 16);

      while (heightLeft > 0) {
        position = heightLeft - imgHeight + 8;
        pdf.addPage();
        pdf.addImage(imgData, 'PNG', 8.5, position, imgWidth, imgHeight);
        heightLeft -= (pdfHeight - 16);
      }

      const dateSuffix = fromDate && toDate ? `${fromDate}_to_${toDate}` : 'all';
      pdf.save(`Visa_Report_${dateSuffix}.pdf`);
    } catch (err) {
      console.error('PDF export error:', err);
      // Fallback to printer dialog
      handlePrint();
    } finally {
      setIsExportingPDF(false);
    }
  };

  return (
    <div className="space-y-5 max-w-7xl mx-auto animate-fade">
      {/* AdminLTE Content Header (Breadcrumbs) */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 bg-white px-4 py-3 rounded-md border border-gray-200 shadow-xs print:hidden">
        <div>
          <h1 className="text-base font-bold text-gray-800 flex items-center gap-2">
            <span>ស្វែងរកតាមកាលបរិច្ឆេទ (Date Search)</span>
          </h1>
          <p className="text-xs text-gray-500">
            កំណត់ចន្លោះកាលបរិច្ឆេទ ពីថ្ងៃណា ដល់ថ្ងៃណា ដើម្បីទាញយកទិន្នន័យ និងរបាយការណ៍ប្តូរទិដ្ឋាការ
          </p>
        </div>
        <div className="flex items-center gap-1.5 text-xs text-gray-500 bg-gray-100 px-3 py-1.5 rounded-md font-medium self-start sm:self-auto">
          <Home className="w-3.5 h-3.5 text-gray-500" />
          <span>ទំព័រដើម</span>
          <span>/</span>
          <span>ការងារទិដ្ឋាការ</span>
          <span>/</span>
          <span className="text-[#007bff] font-bold">ស្វែងរកតាមកាលបរិច្ឆេទ</span>
        </div>
      </div>

      {/* Control Actions */}
      <div className="flex justify-end gap-2 print:hidden">
        <button
          onClick={handleExportExcel}
          className="min-w-[110px] bg-[#28a745] hover:bg-[#218838] text-white px-3.5 py-1.5 rounded text-xs font-bold transition flex items-center justify-center gap-1.5 shadow-2xs cursor-pointer"
          title="ទាញយកជាឯកសារ Excel (.xlsx)"
        >
          <FileSpreadsheet className="w-3.5 h-3.5" />
          <span>Excel</span>
        </button>

        <button
          onClick={handleExportPDF}
          disabled={isExportingPDF}
          className="min-w-[110px] bg-[#17a2b8] hover:bg-[#138496] text-white px-3.5 py-1.5 rounded text-xs font-bold transition flex items-center justify-center gap-1.5 shadow-2xs cursor-pointer disabled:opacity-50"
          title="បង្កើត និងទាញយកជា PDF"
        >
          {isExportingPDF ? (
            <Loader2 className="w-3.5 h-3.5 animate-spin" />
          ) : (
            <FileText className="w-3.5 h-3.5" />
          )}
          <span>{isExportingPDF ? 'កំពុងបង្កើត PDF...' : 'PDF'}</span>
        </button>
      </div>

      {/* FILTER BAR WITH PRESETS CARD */}
      <div className="bg-white rounded-md border border-gray-300 border-t-4 border-t-[#007bff] shadow-xs p-4 space-y-4 print:hidden">
        <div className="flex flex-wrap items-center justify-between gap-3 border-b border-gray-200 pb-3">
          <span className="text-xs text-gray-800 font-bold">ជ្រើសរើសចន្លោះកាលបរិច្ឆេទលឿន៖</span>
          <div className="flex flex-wrap items-center gap-1.5">
            <button
              onClick={() => handlePreset('today')}
              className="bg-[#007bff]/10 hover:bg-[#007bff]/20 text-[#007bff] font-bold px-3 py-1 rounded text-xs transition cursor-pointer"
            >
              ថ្ងៃនេះ
            </button>
            <button
              onClick={() => handlePreset('week')}
              className="bg-[#007bff]/10 hover:bg-[#007bff]/20 text-[#007bff] font-bold px-3 py-1 rounded text-xs transition cursor-pointer"
            >
              សប្តាហ៍នេះ
            </button>
            <button
              onClick={() => handlePreset('month')}
              className="bg-[#007bff]/10 hover:bg-[#007bff]/20 text-[#007bff] font-bold px-3 py-1 rounded text-xs transition cursor-pointer"
            >
              ខែនេះ
            </button>
            <button
              onClick={() => handlePreset('year')}
              className="bg-[#007bff]/10 hover:bg-[#007bff]/20 text-[#007bff] font-bold px-3 py-1 rounded text-xs transition cursor-pointer"
            >
              ឆ្នាំនេះ
            </button>
            <button
              onClick={() => handlePreset('all')}
              className="bg-gray-200 hover:bg-gray-300 text-gray-700 font-bold px-3 py-1 rounded text-xs transition cursor-pointer"
            >
              ទាំងអស់
            </button>
          </div>
        </div>

        <form onSubmit={handleSearchSubmit} className="space-y-4">
          <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
            <div>
              <label className="block text-xs font-bold text-gray-700 mb-1">ចាប់ពីកាលបរិច្ឆេទ (From Date)</label>
              <CustomDatePicker
                value={fromDate}
                onChange={(d) => setFromDate(d)}
                placeholder="YYYY-MM-DD"
                className="py-1.5 text-xs"
              />
            </div>

            <div>
              <label className="block text-xs font-bold text-gray-700 mb-1">រហូតដល់កាលបរិច្ឆេទ (To Date)</label>
              <CustomDatePicker
                value={toDate}
                onChange={(d) => setToDate(d)}
                placeholder="YYYY-MM-DD"
                className="py-1.5 text-xs"
              />
            </div>

            <div>
              <label className="block text-xs font-bold text-gray-700 mb-1">ពាក្យគន្លឹះបន្ថែម (Keyword Search)</label>
              <div className="relative">
                <Search className="w-4 h-4 text-gray-400 absolute left-2.5 top-1/2 -translate-y-1/2" />
                <input
                  type="text"
                  placeholder="លេខលិខិតឆ្លងដែន, ឈ្មោះ..."
                  value={keyword}
                  onChange={(e) => setKeyword(e.target.value)}
                  className="w-full pl-8 pr-3 py-1.5 text-xs border border-gray-300 rounded-sm focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                />
              </div>
            </div>
          </div>

          <div className="flex items-center justify-end gap-2 pt-3 border-t border-gray-200">
            <button
              type="button"
              onClick={handleClearFilters}
              className="min-w-[140px] bg-[#6c757d] hover:bg-[#5a6268] text-white px-4 py-1.5 rounded text-xs font-bold transition flex items-center justify-center gap-1.5 cursor-pointer shadow-2xs"
            >
              <span>សម្អាត</span>
            </button>
            <button
              type="submit"
              className="min-w-[140px] bg-[#007bff] hover:bg-[#0069d9] text-white px-4 py-1.5 rounded text-xs font-bold transition flex items-center justify-center gap-1.5 cursor-pointer shadow-2xs"
            >
              <Search className="w-3.5 h-3.5" />
              <span>ស្វែងរកទិន្នន័យ</span>
            </button>
          </div>
        </form>
      </div>

      {!hasSearched ? (
        <div className="bg-white rounded-2xl border border-[#C6A15B]/30 shadow-sm p-12 text-center text-gray-500">
          <CalendarSearch className="w-10 h-10 text-gray-300 mx-auto mb-3" />
          <p className="text-sm font-semibold text-gray-700">សូមជ្រើសរើសកាលបរិច្ឆេទ រួចចុចប៊ូតុង ស្វែងរកទិន្នន័យ</p>
          <p className="text-xs text-gray-400 mt-1">លទ្ធផលស្វែងរកនឹងត្រូវបង្ហាញបន្ទាប់ពីអ្នកចុចស្វែងរក</p>
        </div>
      ) : (
        <>
          {/* FILTER RESULT SUMMARY STATS */}
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 print:hidden">
            <div className="bg-white rounded-2xl p-4 border border-[#C6A15B]/30 shadow-sm flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-blue-100 text-blue-600 flex items-center justify-center shrink-0">
                <CalendarSearch className="w-5 h-5" />
              </div>
              <div>
                <p className="text-[11px] text-gray-500 font-semibold">ចំនួនសំណើក្នុងចន្លោះពេល</p>
                <p className="text-xl font-bold text-[#0B2545]">{rangeCount} ករណី</p>
              </div>
            </div>

            <div className="bg-white rounded-2xl p-4 border border-[#C6A15B]/30 shadow-sm flex items-center gap-3">
              <div className="w-10 h-10 rounded-xl bg-emerald-100 text-emerald-600 flex items-center justify-center shrink-0">
                <CheckCircle2 className="w-5 h-5" />
              </div>
              <div>
                <p className="text-[11px] text-gray-500 font-semibold">សរុបលិខិត</p>
                <p className="text-xl font-bold text-emerald-700">{rangeUniqueLetters} លិខិត</p>
              </div>
            </div>
          </div>

      {/* FILTERED RESULTS TABLE */}
      <div id="visa-date-search-report" ref={reportRef} className="bg-white rounded-2xl shadow-md border border-[#C6A15B]/30 overflow-hidden">
        <div className="px-6 py-4 bg-[#0B2545] text-[#E4CD98] relative flex flex-col md:flex-row items-center justify-between gap-3">
          <div className="flex-1 text-center md:text-center">
            <h3 className="font-display text-sm md:text-base font-bold text-[#E4CD98] tracking-wide">
              បញ្ជីប្តូរប្រភេទទិដ្ឋាការ និងពន្យារទិដ្ឋាការ
            </h3>
            <p className="text-xs text-white/90 font-medium mt-1">
              {getDateRangeSubtitle()}
            </p>
          </div>
          <div className="flex items-center gap-2 shrink-0 print:hidden">
            <span className="text-xs text-white/70 hidden sm:inline mr-1">
              សរុប {filteredRecords.length} ករណី
            </span>
            <button
              onClick={handleExportExcel}
              className="px-2.5 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white font-bold rounded-lg text-xs shadow-sm transition flex items-center gap-1 cursor-pointer"
              title="ទាញយក Excel (.xlsx)"
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              <span>Excel</span>
            </button>
            <button
              onClick={handleExportPDF}
              disabled={isExportingPDF}
              className="px-2.5 py-1.5 bg-[#C6A15B] hover:bg-[#b08d48] text-[#0B2545] font-bold rounded-lg text-xs shadow-sm transition flex items-center gap-1 cursor-pointer disabled:opacity-50"
              title="បង្កើត PDF"
            >
              {isExportingPDF ? (
                <Loader2 className="w-3.5 h-3.5 animate-spin" />
              ) : (
                <FileText className="w-3.5 h-3.5" />
              )}
              <span>PDF</span>
            </button>
          </div>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="bg-slate-100 text-[#0B2545] font-bold text-xs border-b border-gray-200">
                <th className="px-3 py-3 text-center border-r border-gray-200">លេខ</th>
                <th className="px-3 py-3 border-r border-gray-200">ថ្ងៃខែ</th>
                <th className="px-3 py-3 border-r border-gray-200">ឈ្មោះ</th>
                <th className="px-3 py-3 text-center border-r border-gray-200">ភេទ</th>
                <th className="px-3 py-3 border-r border-gray-200">លិខិតឆ្លងដែន</th>
                <th className="px-3 py-3 border-r border-gray-200">សញ្ជាតិ</th>
                <th className="px-3 py-3 text-center border-r border-gray-200">ទិដ្ឋាការដំបូង</th>
                <th className="px-3 py-3 text-center border-r border-gray-200">ប្តូរ</th>
                <th className="px-3 py-3 border-r border-gray-200">លេខលិខិត</th>
                <th className="px-3 py-3 border-r border-gray-200">ថ្ងៃចេញ</th>
                <th className="px-3 py-3 border-r border-gray-200">ស្ថាប័នដើម</th>
                <th className="px-3 py-3 text-center">ពន្យារ</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100 text-xs text-gray-800">
              {filteredRecords.length > 0 ? (
                filteredRecords.map((r, idx) => (
                  <tr key={r.id} className="hover:bg-amber-50/40 transition">
                    <td className="px-3 py-2.5 text-center font-semibold text-gray-500 border-r border-gray-100">
                      {idx + 1}
                    </td>
                    <td className="px-3 py-2.5 font-mono whitespace-nowrap border-r border-gray-100">
                      {r.applicationDate}
                    </td>
                    <td className="px-3 py-2.5 font-bold text-[#0B2545] uppercase whitespace-nowrap border-r border-gray-100">
                      {r.fullName}
                    </td>
                    <td className="px-3 py-2.5 text-center border-r border-gray-100">
                      <span className="inline-flex items-center justify-center w-6 h-6 rounded-full bg-sky-100 text-sky-800 font-bold text-[11px]">
                        {r.gender === 'ប្រុស' ? 'M' : r.gender === 'ស្រី' ? 'F' : r.gender}
                      </span>
                    </td>
                    <td className="px-3 py-2.5 text-gray-900 border-r border-gray-100 font-mono font-bold">
                      {r.passportNumber}
                    </td>
                    <td className="px-3 py-2.5 whitespace-nowrap border-r border-gray-100">
                      {r.nationality}
                    </td>
                    <td className="px-3 py-2.5 text-center border-r border-gray-100">
                      <span className="px-2 py-0.5 rounded bg-gray-100 text-gray-800 font-mono font-bold border border-gray-200">
                        {r.oldVisaType ? r.oldVisaType.split(' ')[0] : '-'}
                      </span>
                    </td>
                    <td className="px-3 py-2.5 text-center border-r border-gray-100">
                      <span className="px-2 py-0.5 rounded bg-emerald-100 text-emerald-800 font-mono font-bold border border-emerald-200">
                        {r.newVisaType ? r.newVisaType.split(' ')[0] : '-'}
                      </span>
                    </td>
                    <td className="px-3 py-2.5 font-mono border-r border-gray-100">
                      {r.documentNumber || '-'}
                    </td>
                    <td className="px-3 py-2.5 font-mono whitespace-nowrap border-r border-gray-100">
                      {r.issueDate || r.applicationDate}
                    </td>
                    <td className="px-3 py-2.5 border-r border-gray-100 whitespace-nowrap">
                      {r.organization || '-'}
                    </td>
                    <td className="px-3 py-2.5 text-center font-bold text-amber-700">
                      {r.extension || r.duration}
                    </td>
                  </tr>
                ))
              ) : (
                <tr>
                  <td colSpan={12} className="text-center py-12 text-gray-400">
                    មិនមានទិន្នន័យទិដ្ឋាការក្នុងចន្លោះកាលបរិច្ឆេទដែលបានជ្រើសរើសឡើយ
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>
      </div>
        </>
      )}
    </div>
  );
};

