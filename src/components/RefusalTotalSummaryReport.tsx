import React, { useState, useMemo } from 'react';
import { RefusalDeportationRecord } from '../types';
import { toKhmerNum, getKhmerSolarParts } from '../utils/khmerCalendar';
import { FileSpreadsheet, FileText } from 'lucide-react';
import { exportElementToPdf } from '../utils/pdfExportHelper';
import { CustomDatePicker } from './CustomDatePicker';

interface RefusalTotalSummaryReportProps {
  records: RefusalDeportationRecord[];
  categories?: {
    refusalReasons?: Array<{ id: string; name: string }>;
  };
}

export const RefusalTotalSummaryReport: React.FC<RefusalTotalSummaryReportProps> = ({ records, categories }) => {
  const [startDate, setStartDate] = useState<string>('');
  const [endDate, setEndDate] = useState<string>('');
  const [appliedStartDate, setAppliedStartDate] = useState<string>('');
  const [appliedEndDate, setAppliedEndDate] = useState<string>('');
  const [isExporting, setIsExporting] = useState<boolean>(false);

  const filteredRecords = useMemo(() => {
    return records.filter((record) => {
      if (!record.date) return true;
      if (appliedStartDate && record.date < appliedStartDate) return false;
      if (appliedEndDate && record.date > appliedEndDate) return false;
      return true;
    });
  }, [records, appliedStartDate, appliedEndDate]);

  const allReasons = useMemo(() => {
    if (categories?.refusalReasons && categories.refusalReasons.length > 0) {
      return categories.refusalReasons.map((r) => r.name);
    }
    return [
      'លិខិតឆ្លងដែន ឬឯកសារធ្វើដំណើរគ្មានតម្លៃ (១)',
      'គ្មានទិដ្ឋាការចូល (២)',
      'វិកលចរិក (៣)',
      'មានឈ្មោះក្នុងបញ្ជីហាមចូលប្រទេស (៤)',
      'មានជំងឺឆ្លងកាចសាហាវដែលបានប្រកាសដោយរាជរដ្ឋាភិបាល (៥)',
      'កំពុងជាប់ពាក់ព័ន្ធនឹងការប្រព្រឹត្តបទល្មើសនៅក្រៅប្រទេស (៦)',
      'ប្រវត្តិការធ្វើដំណើរមិនច្បាស់លាស់ (៧)',
      'គោលបំណងនៃការធ្វើដំណើរមិនច្បាស់លាស់ (៨)',
      'គ្មានប្រាក់គ្រប់គ្រាន់ (៩)',
      'ឯកសារក្លែងបន្លំ និងកែច្នៃ (១០)',
      'ធ្លាប់បានបណ្តេញចេញពីព្រះរាជាណាចក្រកម្ពុជា (១១)',
      'បានលួចចូលដោយ បន្លំ បើកផ្លូវ ឬប្តូរឈ្មោះចូលមកប្រទេសកម្ពុជា (១២)',
      'មានសកម្មភាពធ្វើអោយប៉ះពាល់ដល់សន្តិសុខជាតិ (១៣)'
    ];
  }, [categories?.refusalReasons]);

  const getMappedReason = (rawReason?: string): string | null => {
    if (!rawReason) return null;
    if (allReasons.includes(rawReason)) return rawReason;

    // Check by extracting any number from the reason string (e.g. "ករណីទី៣" -> 3, "(៣)" -> 3, etc.)
    const matchNum = rawReason.match(/(\d+)/);
    if (matchNum) {
      const num = parseInt(matchNum[1], 10);
      if (num >= 1 && num <= allReasons.length) {
        return allReasons[num - 1];
      }
    }

    for (let i = 0; i < allReasons.length; i++) {
      const targetReason = allReasons[i];
      if (rawReason.includes(`(${i + 1})`) || rawReason.includes(`ករណីទី${i + 1}`)) {
        return targetReason;
      }
    }

    for (const targetReason of allReasons) {
      const shortName = targetReason.split('(')[0].trim();
      if (rawReason.includes(shortName)) {
        return targetReason;
      }
    }

    // Default fallback to first reason if unmapped
    return allReasons[0];
  };

  const summary = useMemo(() => {
    const caseCounts: Record<string, number> = {};
    const peopleCounts: Record<string, number> = {};
    let grandCaseTotal = 0;
    let grandPeopleTotal = 0;

    allReasons.forEach(r => { caseCounts[r] = 0; peopleCounts[r] = 0; });

    filteredRecords.forEach(record => {
      const reasonsInCase = new Set<string>();
      record.people.forEach(person => {
        const mapped = getMappedReason(person.reason);
        if (mapped) reasonsInCase.add(mapped);
      });

      reasonsInCase.forEach(reason => {
        caseCounts[reason] = (caseCounts[reason] || 0) + 1;
        grandCaseTotal += 1;
      });
      
      record.people.forEach(person => {
        const mapped = getMappedReason(person.reason);
        if (mapped) {
          peopleCounts[mapped] = (peopleCounts[mapped] || 0) + 1;
          grandPeopleTotal += 1;
        }
      });
    });

    return { caseCounts, peopleCounts, grandCaseTotal, grandPeopleTotal };
  }, [filteredRecords, allReasons]);

  const formatKhmerDateRange = (start?: string, end?: string) => {
    if (start && end) {
      const s = getKhmerSolarParts(start);
      const e = getKhmerSolarParts(end);
      return `គិតចាប់ពីថ្ងៃទី ${s.khmerDay} ខែ ${s.khmerMonth} ឆ្នាំ ${s.khmerYear} ដល់ថ្ងៃទី ${e.khmerDay} ខែ ${e.khmerMonth} ឆ្នាំ ${e.khmerYear}`;
    } else if (start) {
      const s = getKhmerSolarParts(start);
      return `គិតចាប់ពីថ្ងៃទី ${s.khmerDay} ខែ ${s.khmerMonth} ឆ្នាំ ${s.khmerYear}`;
    } else if (end) {
      const e = getKhmerSolarParts(end);
      return `គិតដល់ថ្ងៃទី ${e.khmerDay} ខែ ${e.khmerMonth} ឆ្នាំ ${e.khmerYear}`;
    } else {
      return 'គិតចាប់ពីថ្ងៃទី ........ ខែ ........ ឆ្នាំ ២០២..... ដល់ថ្ងៃទី ........ ខែ ........ ឆ្នាំ ២០២.....';
    }
  };

  const handleDownloadExcel = () => {
    const tableHTML = document.getElementById('refusal-total-summary-table-container')?.innerHTML || '';
    const dateRangeText = formatKhmerDateRange(appliedStartDate, appliedEndDate);
    const template = `
      <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
      <head>
        <meta charset="utf-8" />
        <style>
          table { border-collapse: collapse; font-family: 'Khmer OS Siemreap', sans-serif; font-size: 11px; width: 100%; }
          th, td { border: 1px solid black; padding: 6px; vertical-align: middle; }
          .font-times { font-family: 'Times New Roman', serif; }
        </style>
      </head>
      <body>
        <h2 style="text-align: center; font-family: 'Khmer OS Siemreap', sans-serif; font-size: 16px; font-weight: bold; margin-bottom: 5px;">សរុបទិន្នន័យបដិសេធ</h2>
        <p style="text-align: center; font-family: 'Khmer OS Siemreap', sans-serif; font-size: 12px; margin-bottom: 15px;">${dateRangeText}</p>
        ${tableHTML}
      </body>
      </html>
    `;
    const blob = new Blob(['\ufeff' + template], { type: 'application/vnd.ms-excel' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `សរុបទិន្នន័យបដិសេធ_${appliedStartDate || 'ទាំងអស់'}_ដល់_${appliedEndDate || 'បច្ចុប្បន្ន'}.xls`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const handleDownloadPDF = async () => {
    const container = document.getElementById('refusal-total-summary-printable-area');
    if (!container) return;
    try {
      setIsExporting(true);
      await exportElementToPdf(
        container,
        `សរុបទិន្នន័យបដិសេធ_${appliedStartDate || 'ទាំងអស់'}_ដល់_${appliedEndDate || 'បច្ចុប្បន្ន'}.pdf`,
        {
          orientation: 'portrait',
          fitSinglePage: true,
          pixelRatio: 3,
        }
      );
    } catch (err) {
      console.error('PDF export failed:', err);
      window.print();
    } finally {
      setIsExporting(false);
    }
  };

  return (
    <div className="p-5 bg-white rounded-lg shadow-sm font-siemreap">
      {/* Top Controls & Filter Section */}
      <div className="mb-5 flex flex-col md:flex-row justify-between items-center gap-4 print:hidden">
        <div>
          <h2 className="text-lg font-bold text-blue-900 font-siemreap">សរុបទិន្នន័យបដិសេធ</h2>
        </div>

        {/* Date Filter & Export Controls */}
        <div className="flex flex-wrap items-center gap-2 text-xs">
          <div className="flex items-center gap-1">
            <span className="text-gray-700 font-bold font-siemreap">ពីថ្ងៃទី:</span>
            <div className="w-28">
              <CustomDatePicker
                value={startDate}
                onChange={(d) => setStartDate(d)}
                className="py-0.5 text-xs font-times"
              />
            </div>
          </div>
          <div className="flex items-center gap-1">
            <span className="text-gray-700 font-bold font-siemreap">ដល់ថ្ងៃទី:</span>
            <div className="w-28">
              <CustomDatePicker
                value={endDate}
                onChange={(d) => setEndDate(d)}
                className="py-0.5 text-xs font-times"
              />
            </div>
          </div>
          <button
            onClick={() => {
              setAppliedStartDate(startDate);
              setAppliedEndDate(endDate);
            }}
            className="px-3 py-1 bg-blue-600 hover:bg-blue-700 text-white rounded font-bold transition text-xs font-siemreap cursor-pointer shadow-sm"
          >
            បង្ហាញលទ្ធផល (Show)
          </button>
          {(startDate || endDate || appliedStartDate || appliedEndDate) && (
            <button
              onClick={() => { 
                setStartDate(''); 
                setEndDate(''); 
                setAppliedStartDate(''); 
                setAppliedEndDate(''); 
              }}
              className="px-2.5 py-1 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded border border-gray-300 transition text-xs font-siemreap cursor-pointer"
            >
              សម្អាត (Reset)
            </button>
          )}

          <div className="flex items-center gap-1.5 ml-2 border-l pl-2 border-gray-300">
            <button
              onClick={handleDownloadExcel}
              className="px-3 py-1.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded font-bold flex items-center gap-1 transition text-xs font-siemreap cursor-pointer shadow-sm"
              title="ទាញយកឯកសារ Excel"
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              <span>Excel</span>
            </button>
            <button
              onClick={handleDownloadPDF}
              disabled={isExporting}
              className="px-3 py-1.5 bg-rose-600 hover:bg-rose-700 disabled:bg-rose-400 text-white rounded font-bold flex items-center gap-1 transition text-xs font-siemreap cursor-pointer shadow-sm"
              title="ទាញយក PDF"
            >
              <FileText className="w-3.5 h-3.5" />
              <span>{isExporting ? 'កំពុងបង្កើត PDF...' : 'PDF'}</span>
            </button>
          </div>
        </div>
      </div>

      {/* Printable / Exportable Area */}
      <div id="refusal-total-summary-printable-area" className="p-6 bg-white max-w-4xl mx-auto border border-gray-200 shadow-xs rounded">
        {/* Centered Title & Date Range */}
        <div className="text-center mb-6">
          <h2 className="text-base font-bold text-black font-siemreap tracking-wider mb-2" style={{ fontFamily: "'Khmer OS Muol Light', 'Khmer OS Siemreap', sans-serif" }}>សរុបទិន្នន័យបដិសេធ</h2>
          <p className="text-xs text-black font-siemreap">
            {appliedStartDate && appliedEndDate ? (
              (() => {
                const s = getKhmerSolarParts(appliedStartDate);
                const e = getKhmerSolarParts(appliedEndDate);
                return `គិតចាប់ពីថ្ងៃទី ${s.khmerDay} ខែ ${s.khmerMonth} ឆ្នាំ ${s.khmerYear} ដល់ថ្ងៃទី ${e.khmerDay} ខែ ${e.khmerMonth} ឆ្នាំ ${e.khmerYear}`;
              })()
            ) : appliedStartDate ? (
              (() => {
                const s = getKhmerSolarParts(appliedStartDate);
                return `គិតចាប់ពីថ្ងៃទី ${s.khmerDay} ខែ ${s.khmerMonth} ឆ្នាំ ${s.khmerYear}`;
              })()
            ) : appliedEndDate ? (
              (() => {
                const e = getKhmerSolarParts(appliedEndDate);
                return `គិតដល់ថ្ងៃទី ${e.khmerDay} ខែ ${e.khmerMonth} ឆ្នាំ ${e.khmerYear}`;
              })()
            ) : (
              <span>
                គិតចាប់ពីថ្ងៃទី <span className="underline underline-offset-4 px-3">................</span> ខែ <span className="underline underline-offset-4 px-5">................</span> ឆ្នាំ ២០២..... ដល់ថ្ងៃទី <span className="underline underline-offset-4 px-3">................</span> ខែ <span className="underline underline-offset-4 px-5">................</span> ឆ្នាំ ២០២.....
              </span>
            )}
          </p>
        </div>

        <div id="refusal-total-summary-table-container" className="overflow-x-auto">
          <table className="w-full text-left border-collapse text-xs border border-black" style={{ fontFamily: "'Khmer OS Siemreap', sans-serif" }}>
            <thead>
              <tr className="bg-gray-100 font-bold text-center border-b border-black font-siemreap">
                <th className="py-2.5 px-2 border border-black w-12 font-siemreap">ល.រ</th>
                <th className="py-2.5 px-3 border border-black text-left font-siemreap">ការបដិសេធផ្ដល់ទិដ្ឋាការ</th>
                <th className="py-2.5 px-3 border border-black w-28 text-center font-siemreap">ករណី</th>
                <th className="py-2.5 px-3 border border-black w-28 text-center font-siemreap">ចំនួនមនុស្ស</th>
              </tr>
            </thead>
            <tbody>
              {allReasons.map((reason, idx) => {
                const numKh = toKhmerNum(idx + 1);
                const cCount = summary.caseCounts[reason] || 0;
                const pCount = summary.peopleCounts[reason] || 0;
                return (
                  <tr key={idx} className="border-b border-black hover:bg-gray-50">
                    <td className="py-2 px-2 border border-black text-center font-siemreap">{numKh}</td>
                    <td className="py-2 px-3 border border-black font-siemreap text-gray-900">{reason}</td>
                    <td className="py-2 px-3 border border-black text-center font-siemreap">
                      {toKhmerNum(cCount)}
                    </td>
                    <td className="py-2 px-3 border border-black text-center font-siemreap">
                      {toKhmerNum(pCount)}
                    </td>
                  </tr>
                );
              })}
              {/* Grand Total Row */}
              <tr className="bg-gray-100 font-bold border-t-2 border-black">
                <td colSpan={2} className="py-2.5 px-3 border border-black text-right font-siemreap text-sm">សរុបចំនួន</td>
                <td className="py-2.5 px-3 border border-black text-center font-siemreap text-sm">
                  {toKhmerNum(summary.grandCaseTotal)} ករណី
                </td>
                <td className="py-2.5 px-3 border border-black text-center font-siemreap text-sm">
                  {toKhmerNum(summary.grandPeopleTotal)} នាក់
                </td>
              </tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};
