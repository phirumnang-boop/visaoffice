import React, { useState } from 'react';
import { VisaRecord, CategoryItem } from '../types';
import { CustomDatePicker } from './CustomDatePicker';
import {
  Calendar,
  Search,
  ClipboardList,
  Check,
  Download,
  Home,
  BarChart3,
  FileSpreadsheet,
} from 'lucide-react';

interface VisaSummaryProps {
  records: VisaRecord[];
  visaTeams: CategoryItem[];
}

export const VisaSummary: React.FC<VisaSummaryProps> = ({ records }) => {
  const [fromDate, setFromDate] = useState('');
  const [toDate, setToDate] = useState('');

  // Applied filter state (only updates when user clicks 'ស្វែងរក' button)
  const [appliedFromDate, setAppliedFromDate] = useState('');
  const [appliedToDate, setAppliedToDate] = useState('');
  const [hasSearched, setHasSearched] = useState(false);

  const formattedToday = new Date().toLocaleDateString('en-US', {
    month: 'long',
    day: 'numeric',
    year: 'numeric',
  });

  const handleSearch = (e: React.FormEvent) => {
    e.preventDefault();
    setAppliedFromDate(fromDate);
    setAppliedToDate(toDate);
    setHasSearched(true);
  };

  const handleClear = () => {
    setFromDate('');
    setToDate('');
    setAppliedFromDate('');
    setAppliedToDate('');
    setHasSearched(false);
  };

  // Filter records based on APPLIED date range (only when user clicks Search)
  const filteredRecords = records.filter((r) => {
    let valid = true;
    if (appliedFromDate) valid = valid && r.applicationDate >= appliedFromDate;
    if (appliedToDate) valid = valid && r.applicationDate <= appliedToDate;
    return valid;
  });

  const totalApplications = filteredRecords.length;

  // Calculate total unique letters/documents based on documentNumber and issueDate
  const totalUniqueLetters = React.useMemo(() => {
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

  // Clean visa code helper
  const cleanVisaType = (typeStr: string) => {
    if (!typeStr) return '';
    return typeStr.trim().split(' ')[0] || typeStr.trim();
  };

  // Calculate visa pairs (Old Visa -> New Visa)
  const pairCounts: Record<string, { oldVisa: string; newVisa: string; count: number }> = {};

  filteredRecords.forEach((r) => {
    const oldCode = cleanVisaType(r.oldVisaType) || 'N/A';
    const newCode = cleanVisaType(r.newVisaType) || 'N/A';
    const key = `${oldCode} → ${newCode}`;

    if (!pairCounts[key]) {
      pairCounts[key] = { oldVisa: oldCode, newVisa: newCode, count: 0 };
    }
    pairCounts[key].count += 1;
  });

  // Convert to list & filter to ONLY show pairs that have data (count > 0)
  const activePairs = Object.entries(pairCounts)
    .map(([key, item]) => ({
      label: key,
      oldVisa: item.oldVisa,
      newVisa: item.newVisa,
      count: item.count,
      percentage: totalApplications > 0 ? Number(((item.count / totalApplications) * 100).toFixed(1)) : 0,
    }))
    .filter((item) => item.count > 0)
    .sort((a, b) => b.count - a.count);

  // CSV Export handler
  const handleExportCSV = () => {
    if (activePairs.length === 0) return;

    const headers = ['លេខ', 'ប្រភេទទិដ្ឋាការ', 'ទិដ្ឋាការដំបូង', 'ប្តូរទៅ', 'ចំនួន', 'សមាមាត្រ (%)'];
    const rows = activePairs.map((item, index) => [
      index + 1,
      `"${item.label}"`,
      `"${item.oldVisa}"`,
      `"${item.newVisa}"`,
      item.count,
      `${item.percentage}%`,
    ]);

    const csvContent = '\uFEFF' + [headers.join(','), ...rows.map((r) => r.join(','))].join('\n');
    const blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.setAttribute('href', url);
    link.setAttribute('download', `visa_summary_${appliedFromDate || 'all'}_${appliedToDate || 'all'}.csv`);
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
  };

  const getVisaBadgeStyle = (code: string) => {
    switch (code.toUpperCase()) {
      case 'T':
        return 'bg-blue-100 text-blue-700 border border-blue-200';
      case 'E':
      case 'EB':
      case 'EG':
        return 'bg-emerald-100 text-emerald-700 border border-emerald-200';
      case 'C':
      case 'K':
        return 'bg-amber-100 text-amber-700 border border-amber-200';
      case 'A':
      case 'B':
        return 'bg-indigo-100 text-indigo-700 border border-indigo-200';
      default:
        return 'bg-gray-100 text-gray-700 border border-gray-200';
    }
  };

  return (
    <div className="space-y-5 max-w-7xl mx-auto animate-fade">
      {/* AdminLTE Content Header (Breadcrumbs) */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 bg-white px-4 py-3 rounded-md border border-gray-200 shadow-xs">
        <div>
          <h1 className="text-base font-bold text-gray-800 flex items-center gap-2">
            <span>សរុបទិន្នន័យប្តូរទិដ្ឋាការ (Visa Summary & Statistics)</span>
          </h1>
          <p className="text-xs text-gray-500">
            សរុបតាមប្រភេទទិដ្ឋាការ និងចន្លោះកាលបរិច្ឆេទក្នុងប្រព័ន្ធ
          </p>
        </div>
        <div className="flex items-center gap-1.5 text-xs text-gray-500 bg-gray-100 px-3 py-1.5 rounded-md font-medium self-start sm:self-auto">
          <Home className="w-3.5 h-3.5 text-gray-500" />
          <span>ទំព័រដើម</span>
          <span>/</span>
          <span>ការងារទិដ្ឋាការ</span>
          <span>/</span>
          <span className="text-[#007bff] font-bold">សរុបទិន្នន័យ</span>
        </div>
      </div>

      {/* Date Search Card */}
      <div className="bg-white rounded-md border border-gray-300 border-t-4 border-t-[#007bff] shadow-xs p-4 space-y-3">
        <div className="flex items-center justify-between border-b border-gray-200 pb-2">
          <h3 className="text-xs font-bold text-gray-800 flex items-center gap-2">
            <Search className="w-4 h-4 text-[#007bff]" />
            <span>ស្វែងរកតាមកាលបរិច្ឆេទ (Date Filter)</span>
          </h3>
          <span className="text-xs text-gray-500 flex items-center gap-1 font-medium bg-gray-100 px-2.5 py-1 rounded">
            <Calendar className="w-3.5 h-3.5 text-[#007bff]" />
            <span>{formattedToday}</span>
          </span>
        </div>

        <form onSubmit={handleSearch} className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 gap-3 items-end pt-1">
          <div>
            <label className="block text-xs font-bold text-gray-700 mb-1">ចាប់ពីថ្ងៃ (From Date)</label>
            <CustomDatePicker
              value={fromDate}
              onChange={(d) => setFromDate(d)}
              placeholder="YYYY-MM-DD"
              className="py-1.5 text-xs"
            />
          </div>

          <div>
            <label className="block text-xs font-bold text-gray-700 mb-1">ដល់ថ្ងៃ (To Date)</label>
            <CustomDatePicker
              value={toDate}
              onChange={(d) => setToDate(d)}
              placeholder="YYYY-MM-DD"
              className="py-1.5 text-xs"
            />
          </div>

          <div className="flex items-center gap-2 md:col-span-2">
            <button
              type="submit"
              className="min-w-[140px] bg-[#007bff] hover:bg-[#0069d9] text-white px-4 py-1.5 rounded text-xs font-bold flex items-center justify-center gap-1.5 shadow-2xs transition cursor-pointer"
            >
              <Search className="w-3.5 h-3.5" />
              <span>ស្វែងរកទិន្នន័យ</span>
            </button>

            <button
              type="button"
              onClick={handleClear}
              className="min-w-[140px] bg-[#6c757d] hover:bg-[#5a6268] text-white px-4 py-1.5 rounded text-xs font-bold flex items-center justify-center gap-1.5 shadow-2xs transition cursor-pointer"
            >
              <span>សម្អាត</span>
            </button>
          </div>
        </form>
      </div>

      {!hasSearched ? (
        <div className="bg-white rounded-md border border-gray-300 p-12 text-center text-gray-500 shadow-xs">
          <BarChart3 className="w-12 h-12 text-gray-300 mx-auto mb-3" />
          <p className="text-sm font-bold text-gray-700">សូមជ្រើសរើសកាលបរិច្ឆេទ រួចចុចប៊ូតុង ស្វែងរកទិន្នន័យ</p>
          <p className="text-xs text-gray-400 mt-1">របាយការណ៍សរុបតាមប្រភេទ និងសមាមាត្រនឹងត្រូវបង្ហាញជូន</p>
        </div>
      ) : (
        <>
          {/* Overview Stat Cards */}
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            <div className="bg-white rounded-md p-4 border border-gray-300 border-l-4 border-l-[#007bff] shadow-xs flex items-center justify-between">
              <div>
                <p className="text-xs font-bold text-gray-500 uppercase">សរុបទិន្នន័យក្នុងកាលបរិច្ឆេទ</p>
                <p className="text-2xl font-bold text-gray-800 leading-tight mt-1">{totalApplications} កំណត់ត្រា</p>
              </div>
              <div className="w-11 h-11 rounded bg-blue-50 text-[#007bff] flex items-center justify-center shrink-0 border border-blue-100">
                <ClipboardList className="w-6 h-6" />
              </div>
            </div>

            <div className="bg-white rounded-md p-4 border border-gray-300 border-l-4 border-l-[#28a745] shadow-xs flex items-center justify-between">
              <div>
                <p className="text-xs font-bold text-gray-500 uppercase">សរុបចំនួនលិខិត</p>
                <p className="text-2xl font-bold text-gray-800 leading-tight mt-1">{totalUniqueLetters} លិខិត</p>
              </div>
              <div className="w-11 h-11 rounded bg-emerald-50 text-[#28a745] flex items-center justify-center shrink-0 border border-emerald-100">
                <Check className="w-6 h-6" />
              </div>
            </div>
          </div>

          {/* Table Card */}
          <div className="bg-white rounded-md border border-gray-300 border-t-4 border-t-[#007bff] shadow-xs overflow-hidden">
            <div className="bg-[#007bff] text-white px-4 py-3 flex items-center justify-between">
              <h3 className="text-sm font-bold flex items-center gap-2">
                <FileSpreadsheet className="w-4 h-4 text-white" />
                <span>សរុបទិន្នន័យប្តូរទិដ្ឋាការ ({activePairs.length} ប្រភេទ)</span>
              </h3>
              <button
                type="button"
                onClick={handleExportCSV}
                className="bg-[#28a745] hover:bg-[#218838] text-white px-3 py-1 rounded text-xs font-bold flex items-center gap-1.5 shadow-2xs transition cursor-pointer"
              >
                <Download className="w-3.5 h-3.5" />
                <span>នាំចេញ CSV</span>
              </button>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs">
                <thead>
                  <tr className="bg-gray-100 text-gray-800 font-bold border-b border-gray-200">
                    <th className="py-2.5 px-3 w-12 text-center">ល.រ</th>
                    <th className="py-2.5 px-4 font-bold">ប្រភេទទិដ្ឋាការ</th>
                    <th className="py-2.5 px-4">ទិដ្ឋាការដំបូង</th>
                    <th className="py-2.5 px-4">ប្តូរទៅ</th>
                    <th className="py-2.5 px-4 text-center font-bold">ចំនួន</th>
                    <th className="py-2.5 px-4 w-64">សមាមាត្រ (%)</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-200">
                  {activePairs.length > 0 ? (
                    activePairs.map((item, idx) => (
                      <tr key={item.label} className="hover:bg-blue-50/30 transition">
                        <td className="py-2.5 px-3 text-center text-gray-500 font-bold">{idx + 1}</td>
                        <td className="py-2.5 px-4 font-bold text-gray-800">{item.label}</td>
                        <td className="py-2.5 px-4">
                          <span className={`inline-flex items-center justify-center px-2 py-0.5 rounded text-xs font-bold ${getVisaBadgeStyle(item.oldVisa)}`}>
                            {item.oldVisa}
                          </span>
                        </td>
                        <td className="py-2.5 px-4">
                          <span className={`inline-flex items-center justify-center px-2 py-0.5 rounded text-xs font-bold ${getVisaBadgeStyle(item.newVisa)}`}>
                            {item.newVisa}
                          </span>
                        </td>
                        <td className="py-2.5 px-4 text-center font-bold text-gray-800 text-sm">{item.count}</td>
                        <td className="py-2.5 px-4">
                          <div className="flex items-center gap-3">
                            <div className="flex-1 bg-gray-200 h-2 rounded-full overflow-hidden">
                              <div
                                className="bg-[#007bff] h-full rounded-full transition-all duration-300"
                                style={{ width: `${item.percentage}%` }}
                              />
                            </div>
                            <span className="text-xs font-bold text-gray-700 min-w-[42px] text-right">
                              {item.percentage}%
                            </span>
                          </div>
                        </td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={6} className="text-center py-10 text-gray-400 font-medium">
                        មិនមានទិន្នន័យប្តូរទិដ្ឋាការទេ
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

