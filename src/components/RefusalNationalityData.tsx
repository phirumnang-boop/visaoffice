import React, { useState, useMemo } from 'react';
import { RefusalDeportationRecord } from '../types';
import { toKhmerNum, getKhmerSolarParts } from '../utils/khmerCalendar';
import { Printer, Download, FileSpreadsheet, FileText } from 'lucide-react';
import { exportElementToPdf } from '../utils/pdfExportHelper';
import { CustomDatePicker } from './CustomDatePicker';

interface RefusalNationalityDataProps {
  records: RefusalDeportationRecord[];
  categories?: {
    refusalReasons?: Array<{ id: string; name: string }>;
  };
}

const NATIONALITY_MAP: Record<string, { kh: string; en: string }> = {
  'CHN': { kh: 'ចិន', en: 'China' },
  'VNM': { kh: 'វៀតណាម', en: 'Vietnam' },
  'THA': { kh: 'ថៃ', en: 'Thailand' },
  'LAO': { kh: 'ឡាវ', en: 'Laos' },
  'USA': { kh: 'អាមេរិក', en: 'United States' },
  'FRA': { kh: 'បារាំង', en: 'France' },
  'GBR': { kh: 'អង់គ្លេស', en: 'United Kingdom' },
  'JPN': { kh: 'ជប៉ុន', en: 'Japan' },
  'KOR': { kh: 'កូរ៉េខាងត្បូង', en: 'South Korea' },
  'SGP': { kh: 'សិង្ហបុរី', en: 'Singapore' },
  'MYS': { kh: 'ម៉ាឡេស៊ី', en: 'Malaysia' },
  'IND': { kh: 'ឥណ្ឌា', en: 'India' },
  'RUS': { kh: 'រុស្ស៊ី', en: 'Russia' },
  'SOM': { kh: 'សូម៉ាលី', en: 'Somalia' },
};

const getNationalityName = (codeOrName: string, english: boolean) => {
  if (!codeOrName) return english ? 'Unknown' : 'មិនស្គាល់';
  const upper = codeOrName.trim().toUpperCase();
  if (NATIONALITY_MAP[upper]) {
    return english ? NATIONALITY_MAP[upper].en : NATIONALITY_MAP[upper].kh;
  }
  for (const [code, val] of Object.entries(NATIONALITY_MAP)) {
    if (codeOrName === val.kh || codeOrName.toLowerCase() === val.en.toLowerCase()) {
      return english ? val.en : val.kh;
    }
  }
  return codeOrName;
};

export const RefusalNationalityData: React.FC<RefusalNationalityDataProps> = ({ records, categories }) => {
  const [startDate, setStartDate] = useState<string>('');
  const [endDate, setEndDate] = useState<string>('');
  const [appliedStartDate, setAppliedStartDate] = useState<string>('');
  const [appliedEndDate, setAppliedEndDate] = useState<string>('');
  const [isExporting, setIsExporting] = useState<boolean>(false);
  const [showEnglishCountry, setShowEnglishCountry] = useState<boolean>(false);

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
      'វិកលចរឹក (៣)',
      'មានឈ្មោះក្នុងបញ្ជីហាមចូលប្រទេស (៤)',
      'មានជំងឺឆ្លងកាចសាហាវដែលបានប្រកាសដោយរាជរដ្ឋាភិបាល (៥)',
      'កំពុងជាប់ពាក់ព័ន្ធនឹងការប្រព្រឹត្តបទល្មើសនៅក្រៅប្រទេស (៦)',
      'ប្រវត្តិការធ្វើដំណើរមិនច្បាស់លាស់ (៧)',
      'គោលបំណងនៃការធ្វើដំណើរមិនច្បាស់លាស់ (៨)',
      'គ្មានប្រាក់គ្រប់គ្រាន់ (៩)',
      'ឯកសារក្លែងបន្លំ និងកែច្នៃ (១០)',
      'ធ្លាប់បានបណ្តេញចេញពីព្រះរាជាណាចក្រកម្ពុជា (១១)',
      'បានលួចចូលដោយ បន្លំ បើកផ្លូវ ឬប្តូរឈ្មោះចូលមកប្រទេសកម្ពុជា (១២)',
      'មានសកម្មភាពធ្វើអោយប៉ះពាល់ដល់សន្តិសុខជាតិ (១៣)',
    ];
  }, [categories?.refusalReasons]);

  const getMappedReason = (rawReason?: string): string | null => {
    if (!rawReason) return null;
    if (allReasons.includes(rawReason)) return rawReason;

    // Check by number in parentheses
    for (let i = 0; i < allReasons.length; i++) {
      const targetReason = allReasons[i];
      const numMatch = targetReason.match(/\(([០-៩0-9]+)\)/);
      if (numMatch) {
        const numStr = numMatch[1];
        if (rawReason.includes(numStr) || rawReason.includes(`(${numStr})`)) {
          return targetReason;
        }
      }
      if (rawReason.includes(`(${i + 1})`) || rawReason.includes(`ករណីទី${i + 1}`)) {
        return targetReason;
      }
    }

    // Fallback substring matching
    for (const targetReason of allReasons) {
      const shortName = targetReason.split('(')[0].trim();
      if (rawReason.includes(shortName)) {
        return targetReason;
      }
    }

    return null;
  };

  const tableData = useMemo(() => {
    const data: Record<string, Record<string, number>> = {};
    const nationalities = new Set<string>();

    filteredRecords.forEach((record) => {
      record.people.forEach((person) => {
        const nat = getNationalityName(person.nationality || '', showEnglishCountry);
        const mappedReason = getMappedReason(person.reason);
        if (mappedReason) {
          if (!data[nat]) data[nat] = {};
          data[nat][mappedReason] = (data[nat][mappedReason] || 0) + 1;
          nationalities.add(nat);
        }
      });
    });

    return { data, nationalities: Array.from(nationalities).sort() };
  }, [filteredRecords, showEnglishCountry]);

  const summary = useMemo(() => {
    const caseCounts: Record<string, number> = {};
    const peopleCounts: Record<string, number> = {};
    const rowTotals: Record<string, number> = {};
    const colTotals: Record<string, number> = {};
    let grandTotal = 0;
    let grandCaseTotal = 0;
    let grandPeopleTotal = 0;

    allReasons.forEach(r => { caseCounts[r] = 0; peopleCounts[r] = 0; colTotals[r] = 0; });

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

    tableData.nationalities.forEach((nat) => {
      rowTotals[nat] = 0;
      allReasons.forEach((reason) => {
        const count = tableData.data[nat]?.[reason] || 0;
        rowTotals[nat] += count;
        colTotals[reason] = (colTotals[reason] || 0) + count;
        grandTotal += count;
      });
    });

    return { caseCounts, peopleCounts, rowTotals, colTotals, grandTotal, grandCaseTotal, grandPeopleTotal };
  }, [filteredRecords, tableData]);

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
    const tableHTML = document.getElementById('refusal-nationality-table-container')?.innerHTML || '';
    const dateRangeText = formatKhmerDateRange(appliedStartDate, appliedEndDate);
    const template = `
      <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
      <head>
        <meta charset="utf-8" />
        <style>
          table { border-collapse: collapse; font-family: 'Khmer OS Siemreap', sans-serif; font-size: 11px; width: 100%; }
          th, td { border: 1px solid black; padding: 4px; text-align: center; vertical-align: middle; }
          .font-times { font-family: 'Times New Roman', serif; }
        </style>
      </head>
      <body>
        <h2 style="text-align: center; font-family: 'Khmer OS Siemreap', sans-serif; font-size: 16px; font-weight: bold; margin-bottom: 5px;">ទិន្នន័យសញ្ជាតិបដិសេធ</h2>
        <p style="text-align: center; font-family: 'Khmer OS Siemreap', sans-serif; font-size: 12px; margin-bottom: 15px;">${dateRangeText}</p>
        ${tableHTML}
      </body>
      </html>
    `;
    const blob = new Blob(['\ufeff' + template], { type: 'application/vnd.ms-excel' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `ទិន្នន័យសញ្ជាតិបដិសេធ_${appliedStartDate || 'ទាំងអស់'}_ដល់_${appliedEndDate || 'បច្ចុប្បន្ន'}.xls`;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  };

  const handleDownloadPDF = async () => {
    const container = document.getElementById('refusal-nationality-printable-area');
    if (!container) return;
    try {
      setIsExporting(true);
      await exportElementToPdf(
        container,
        `ទិន្នន័យសញ្ជាតិបដិសេធ_${appliedStartDate || 'ទាំងអស់'}_ដល់_${appliedEndDate || 'បច្ចុប្បន្ន'}.pdf`,
        {
          orientation: 'landscape',
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
      {/* Top Controls & Filter Section (Hidden in PDF/Print) */}
      <div className="mb-5 flex flex-col md:flex-row justify-between items-center gap-4 print:hidden">
        <div>
          <h2 className="text-lg font-bold text-blue-900 font-siemreap">តារាងទិន្នន័យសញ្ជាតិបដិសេធ</h2>
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

          <button
            onClick={() => setShowEnglishCountry(!showEnglishCountry)}
            className={`px-3 py-1 rounded border text-xs font-siemreap font-medium transition-all ${
              showEnglishCountry ? 'bg-indigo-600 text-white border-indigo-600 shadow' : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
            }`}
            title="ប្តូររវាងឈ្មោះប្រទេសជាភាសាខ្មែរ និងអង់គ្លេស"
          >
            🌐 {showEnglishCountry ? 'ឈ្មោះប្រទេស (Khmer)' : 'English Country'}
          </button>

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
      <div id="refusal-nationality-printable-area" className="p-4 bg-white">
        {/* Centered Title & Date Range */}
        <div className="text-center mb-5">
          <h2 className="text-xl font-bold text-black font-siemreap tracking-wider">ទិន្នន័យសញ្ជាតិបដិសេធ</h2>
          <p className="text-xs text-black font-siemreap mt-1">
            {formatKhmerDateRange(appliedStartDate, appliedEndDate)}
          </p>
        </div>

        <div id="refusal-nationality-table-container" className="overflow-x-auto">
          <table className="w-full text-left border-collapse text-xs border-2 border-black" style={{ fontFamily: "'Khmer OS Siemreap', sans-serif" }}>
            <thead>
              {/* Top row: Vertical reasons headers */}
              <tr className="bg-[#89CFF0]">
                <th colSpan={2} className="p-2 border border-black text-center align-middle font-normal" style={{ backgroundColor: '#89CFF0', height: '200px', width: '160px', fontFamily: "'Khmer OS Siemreap', sans-serif", fontWeight: 400 }}>
                  <div style={{ writingMode: 'vertical-rl', transform: 'rotate(180deg)', margin: '0 auto', whiteSpace: 'nowrap' }}>មូលហេតុបដិសេធ</div>
                </th>
                {allReasons.map((reason, idx) => (
                  <th key={idx} className="p-2 border border-black text-center align-middle font-normal" style={{ height: '200px', width: '42px', backgroundColor: idx % 2 === 0 ? '#89CFF0' : '#D3D3D3', fontFamily: "'Khmer OS Siemreap', sans-serif", fontWeight: 400 }}>
                    <div style={{ writingMode: 'vertical-rl', transform: 'rotate(180deg)', margin: '0 auto', whiteSpace: 'nowrap' }}>{reason}</div>
                  </th>
                ))}
                <th className="p-2 border border-black text-center align-middle bg-[#D3D3D3] font-normal" style={{ height: '200px', width: '50px', fontFamily: "'Khmer OS Siemreap', sans-serif", fontWeight: 400 }}>
                  <div style={{ writingMode: 'vertical-rl', transform: 'rotate(180deg)', margin: '0 auto', whiteSpace: 'nowrap' }}>សរុប</div>
                </th>
              </tr>
              
              {/* Counts rows */}
              <tr className="bg-white border-b border-black">
                <th colSpan={2} className="p-2 border border-black font-normal" style={{ textAlign: 'left', fontFamily: "'Khmer OS Siemreap', sans-serif", fontWeight: 400 }}>ករណី</th>
                {allReasons.map((reason, idx) => (
                  <th key={idx} className="p-2 border border-black text-center font-normal" style={{ fontFamily: "'Times New Roman', serif", fontWeight: 400 }}>{summary.caseCounts[reason] || 0}</th>
                ))}
                <th className="p-2 border border-black text-center font-normal" style={{ fontFamily: "'Times New Roman', serif", fontWeight: 400 }}>{summary.grandCaseTotal}</th>
              </tr>
              <tr className="bg-white border-b border-black">
                <th colSpan={2} className="p-2 border border-black font-normal" style={{ textAlign: 'left', fontFamily: "'Khmer OS Siemreap', sans-serif", fontWeight: 400 }}>ចំនួនមនុស្ស</th>
                {allReasons.map((reason, idx) => (
                  <th key={idx} className="p-2 border border-black text-center font-normal" style={{ fontFamily: "'Times New Roman', serif", fontWeight: 400 }}>{summary.peopleCounts[reason] || 0}</th>
                ))}
                <th className="p-2 border border-black text-center font-normal" style={{ fontFamily: "'Times New Roman', serif", fontWeight: 400 }}>{summary.grandPeopleTotal}</th>
              </tr>
              
              {/* Nationalities header */}
              <tr className="bg-[#6A0DAD] text-white">
                <th className="p-2 border border-black text-center font-normal" style={{ width: '40px', fontFamily: "'Khmer OS Siemreap', sans-serif", fontWeight: 400 }}>ល.រ</th>
                <th className="p-2 border border-black font-normal" style={{ textAlign: 'left', width: '130px', fontFamily: "'Khmer OS Siemreap', sans-serif", fontWeight: 400 }}>សញ្ជាតិ</th>
                {allReasons.map((_, idx) => (
                  <th key={idx} className="p-2 border border-black text-center font-normal" style={{ width: '42px', fontFamily: "'Khmer OS Siemreap', sans-serif", fontWeight: 400 }}>{toKhmerNum(idx + 1)}</th>
                ))}
                <th className="p-2 border border-black text-center font-normal" style={{ width: '50px', fontFamily: "'Khmer OS Siemreap', sans-serif", fontWeight: 400 }}>សរុប</th>
              </tr>
            </thead>
            <tbody>
              {tableData.nationalities.length === 0 ? (
                <tr>
                  <td colSpan={16} className="text-center py-6 text-gray-500 font-siemreap">
                    មិនទាន់មានទិន្នន័យបដិសេធតាមកាលបរិច្ឆេទនេះនៅឡើយ
                  </td>
                </tr>
              ) : (
                tableData.nationalities.map((nat, idx) => (
                  <tr key={nat} className="border-b border-black">
                    <td className="p-2 border border-black text-center" style={{ fontFamily: "'Khmer OS Siemreap', sans-serif", fontWeight: 400 }}>{toKhmerNum(idx + 1)}</td>
                    <td className="p-2 border border-black" style={{ fontFamily: showEnglishCountry ? "'Times New Roman', serif" : "'Khmer OS Siemreap', sans-serif", fontWeight: 400 }}>{nat}</td>
                    {allReasons.map((reason, rIdx) => (
                      <td key={rIdx} className="p-2 border border-black text-center" style={{ fontFamily: "'Times New Roman', serif", fontWeight: 400 }}>
                        {tableData.data[nat]?.[reason] || 0}
                      </td>
                    ))}
                    <td className="p-2 border border-black text-center" style={{ fontFamily: "'Times New Roman', serif", fontWeight: 400 }}>{summary.rowTotals[nat] || 0}</td>
                  </tr>
                ))
              )}
            </tbody>
          </table>
        </div>
      </div>
    </div>
  );
};
