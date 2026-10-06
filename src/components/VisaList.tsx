import React, { useState } from 'react';
import { VisaRecord, Officer, CategoryItem } from '../types';
import { VisaImportModal } from './VisaImportModal';
import { printA4Document } from '../utils/printHelper';
import { CustomDatePicker } from './CustomDatePicker';
import {
  Search,
  Filter,
  Eye,
  Edit,
  Trash2,
  FileSpreadsheet,
  Printer,
  ChevronLeft,
  ChevronRight,
  RefreshCw,
  X,
  Save,
  CheckCircle2,
  AlertCircle,
  Clock,
  Upload,
  Calendar,
  AlertTriangle,
  Home,
} from 'lucide-react';

const VISA_TYPES = [
  'T',
  'E',
  'EB',
  'EG',
  'ER',
  'ES',
  'EP',
  'B',
  'C',
  'K',
];

const EXTENSIONS = [
  { code: '1M', label: '1M (1 ខែ)', fee: 50 },
  { code: '3M', label: '3M (3 ខែ)', fee: 80 },
  { code: '6M', label: '6M (6 ខែ)', fee: 180 },
  { code: '1Y', label: '1Y (12 ខែ / 1 ឆ្នាំ)', fee: 290 },
];

interface VisaListProps {
  records: VisaRecord[];
  officers: Officer[];
  visaTeams: CategoryItem[];
  organizations?: CategoryItem[];
  onUpdateRecord: (record: VisaRecord) => void;
  onDeleteRecord: (id: string) => void;
  onBatchImportRecords?: (records: Omit<VisaRecord, 'id' | 'createdAt'>[]) => void;
  onDeleteRecordsByDateRange?: (startDate: string, endDate: string) => void;
  onShowToast: (msg: string, type: 'success' | 'error' | 'info') => void;
  onNavigateToForm: () => void;
}

export const VisaList: React.FC<VisaListProps> = ({
  records,
  officers,
  visaTeams,
  organizations = [],
  onUpdateRecord,
  onDeleteRecord,
  onBatchImportRecords,
  onDeleteRecordsByDateRange,
  onShowToast,
  onNavigateToForm,
}) => {
  const [searchTerm, setSearchTerm] = useState('');
  const [selectedVisaType, setSelectedVisaType] = useState('ALL');
  const [selectedStatus, setSelectedStatus] = useState('ALL');

  // Modal States
  const [viewingRecord, setViewingRecord] = useState<VisaRecord | null>(null);
  const [editingRecord, setEditingRecord] = useState<VisaRecord | null>(null);
  const [deletingRecord, setDeletingRecord] = useState<{ id: string; name: string } | null>(null);
  const [isImportModalOpen, setIsImportModalOpen] = useState(false);

  // Bulk Delete States
  const [isBulkDeleteModalOpen, setIsBulkDeleteModalOpen] = useState(false);
  const [bulkStartDate, setBulkStartDate] = useState('');
  const [bulkEndDate, setBulkEndDate] = useState('');
  const [showConfirmBulkModal, setShowConfirmBulkModal] = useState(false);

  // Pagination
  const [currentPage, setCurrentPage] = useState(1);
  const itemsPerPage = 8;

  // Filter Logic
  const filteredRecords = records.filter((r) => {
    const matchesSearch =
      r.passportNumber.toLowerCase().includes(searchTerm.toLowerCase()) ||
      r.fullName.toLowerCase().includes(searchTerm.toLowerCase()) ||
      r.nationality.toLowerCase().includes(searchTerm.toLowerCase());

    const matchesType =
      selectedVisaType === 'ALL' || r.newVisaType.includes(selectedVisaType);

    const matchesStatus =
      selectedStatus === 'ALL' || r.status === selectedStatus;

    return matchesSearch && matchesType && matchesStatus;
  });

  const totalPages = Math.ceil(filteredRecords.length / itemsPerPage) || 1;
  const paginatedRecords = filteredRecords.slice(
    (currentPage - 1) * itemsPerPage,
    currentPage * itemsPerPage
  );

  const totalFees = filteredRecords.reduce((sum, r) => sum + (r.fee || 0), 0);

  // Delete Action
  const handleConfirmDelete = () => {
    if (deletingRecord) {
      onDeleteRecord(deletingRecord.id);
      onShowToast(`បានលុបទិន្នន័យទិដ្ឋាការរបស់ "${deletingRecord.name}" រួចរាល់`, 'success');
      setDeletingRecord(null);
    }
  };

  // Bulk Delete calculation and action handlers
  const matchingBulkRecords = records.filter((r) => {
    if (!bulkStartDate || !bulkEndDate) return false;
    const d = r.applicationDate || r.issueDate || '';
    return d >= bulkStartDate && d <= bulkEndDate;
  });

  const handleOpenBulkDeleteModal = () => {
    const today = new Date().toISOString().split('T')[0];
    const firstDay = today.substring(0, 8) + '01';
    setBulkStartDate(firstDay);
    setBulkEndDate(today);
    setIsBulkDeleteModalOpen(true);
  };

  const handleRequestBulkDelete = () => {
    if (!bulkStartDate || !bulkEndDate) {
      onShowToast('សូមជ្រើសរើសថ្ងៃចាប់ផ្តើម និងថ្ងៃបញ្ចប់!', 'error');
      return;
    }

    if (bulkStartDate > bulkEndDate) {
      onShowToast('ថ្ងៃចាប់ផ្តើមមិនអាចធំជាងថ្ងៃបញ្ចប់ឡើយ!', 'error');
      return;
    }

    if (matchingBulkRecords.length === 0) {
      onShowToast(`មិនមានទិន្នន័យនៅក្នុងចន្លោះថ្ងៃ ${bulkStartDate} ដល់ ${bulkEndDate} ឡើយ!`, 'info');
      return;
    }

    setShowConfirmBulkModal(true);
  };

  const handleExecuteBulkDelete = () => {
    if (onDeleteRecordsByDateRange) {
      onDeleteRecordsByDateRange(bulkStartDate, bulkEndDate);
      onShowToast(`បានលុបទិន្នន័យចំនួន ${matchingBulkRecords.length} កំណត់ត្រាដោយជោគជ័យ!`, 'success');
    } else {
      matchingBulkRecords.forEach((r) => onDeleteRecord(r.id));
      onShowToast(`បានលុបទិន្នន័យចំនួន ${matchingBulkRecords.length} កំណត់ត្រាដោយជោគជ័យ!`, 'success');
    }

    setShowConfirmBulkModal(false);
    setIsBulkDeleteModalOpen(false);
  };

  // Edit Submit
  const handleEditSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingRecord) return;

    if (!editingRecord.passportNumber.trim() || !editingRecord.fullName.trim()) {
      onShowToast('សូមបំពេញព័ត៌មានឱ្យបានគ្រប់គ្រាន់', 'error');
      return;
    }

    onUpdateRecord(editingRecord);
    onShowToast(`បានកែប្រែទិន្នន័យ ${editingRecord.fullName} ដោយជោគជ័យ!`, 'success');
    setEditingRecord(null);
  };

  const handlePrint = () => {
    printA4Document('visa-list-print-table', {
      orientation: 'landscape',
      documentTitle: `បញ្ជីស្នើសុំប្តូរទិដ្ឋាការ_${filteredRecords.length}`,
    });
  };

  return (
    <div className="space-y-5 w-full animate-fade">
      {/* AdminLTE Content Header (Breadcrumbs) */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 bg-white px-4 py-3 rounded-md border border-gray-200 shadow-xs">
        <div>
          <h1 className="text-base font-bold text-gray-800 flex items-center gap-2">
            <span>ទិន្នន័យបញ្ជីស្នើសុំប្តូរទិដ្ឋាការ (Visa List)</span>
          </h1>
          <p className="text-xs text-gray-500">
            គ្រប់គ្រង និងស្វែងរកបញ្ជីឈ្មោះជនបរទេសស្នើសុំប្តូរប្រភេទទិដ្ឋាការ (សរុប {records.length} កំណត់ត្រា)
          </p>
        </div>
        <div className="flex items-center gap-1.5 text-xs text-gray-500 bg-gray-100 px-3 py-1.5 rounded-md font-medium self-start sm:self-auto">
          <Home className="w-3.5 h-3.5 text-gray-500" />
          <span>ទំព័រដើម</span>
          <span>/</span>
          <span>ការងារទិដ្ឋាការ</span>
          <span>/</span>
          <span className="text-[#007bff] font-bold">បញ្ជីស្នើសុំប្តូរ</span>
        </div>
      </div>

      {/* Control Bar with Actions */}
      <div className="bg-white rounded-md border border-gray-300 border-t-4 border-t-[#007bff] shadow-xs p-4 space-y-3">
        <div className="flex flex-col md:flex-row items-stretch md:items-center justify-between gap-3">
          {/* SEARCH BOX */}
          <div className="relative flex-1">
            <Search className="w-4 h-4 text-gray-400 absolute left-3 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              placeholder="ស្វែងរកតាម លេខលិខិតឆ្លងដែន, ឈ្មោះ, សញ្ជាតិ..."
              value={searchTerm}
              onChange={(e) => {
                setSearchTerm(e.target.value);
                setCurrentPage(1);
              }}
              className="w-full pl-9 pr-8 py-1.5 text-xs border border-gray-300 rounded-sm focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
            />
            {searchTerm && (
              <button
                onClick={() => setSearchTerm('')}
                className="absolute right-2.5 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600"
              >
                <X className="w-3.5 h-3.5" />
              </button>
            )}
          </div>

          {/* Action Buttons */}
          <div className="flex flex-wrap items-center gap-2">
            <button
              onClick={handleOpenBulkDeleteModal}
              className="bg-[#dc3545] hover:bg-[#c82333] text-white px-3.5 py-1.5 rounded text-xs font-bold transition flex items-center gap-1.5 shadow-2xs cursor-pointer"
              title="ជ្រើសរើសកាលបរិច្ឆេទដើម្បីលុបទិន្នន័យច្រើន"
            >
              <Trash2 className="w-3.5 h-3.5" />
              <span>លុបតាមកាលបរិច្ឆេទ</span>
            </button>
            <button
              onClick={() => setIsImportModalOpen(true)}
              className="bg-[#28a745] hover:bg-[#218838] text-white px-3.5 py-1.5 rounded text-xs font-bold transition flex items-center gap-1.5 shadow-2xs cursor-pointer"
            >
              <Upload className="w-3.5 h-3.5" />
              <span>នាំចូល Excel / CSV</span>
            </button>
            <button
              onClick={onNavigateToForm}
              className="bg-[#007bff] hover:bg-[#0069d9] text-white px-3.5 py-1.5 rounded text-xs font-bold transition flex items-center gap-1.5 shadow-2xs cursor-pointer"
            >
              <RefreshCw className="w-3.5 h-3.5" />
              <span>+ ស្នើសុំប្តូរថ្មី</span>
            </button>
          </div>
        </div>

        {/* Filters Row */}
        <div className="flex flex-wrap items-center gap-4 pt-2 border-t border-gray-200 text-xs">
          {/* Visa Type Filter */}
          <div className="flex items-center gap-2">
            <span className="font-bold text-gray-700">ប្រភេទទិដ្ឋាការ៖</span>
            <select
              value={selectedVisaType}
              onChange={(e) => {
                setSelectedVisaType(e.target.value);
                setCurrentPage(1);
              }}
              className="border border-gray-300 rounded-sm px-2.5 py-1 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
            >
              <option value="ALL">ទាំងអស់</option>
              <option value="EB">EB (Business)</option>
              <option value="EG">EG (General)</option>
              <option value="ER">ER (Retirement)</option>
              <option value="ES">ES (Student)</option>
              <option value="T">T (Tourist)</option>
              <option value="E">E (Ordinary)</option>
            </select>
          </div>
        </div>
      </div>

      {/* TABLE DATA CONTAINER */}
      <div className="bg-white rounded-md border border-gray-300 border-t-4 border-t-[#007bff] shadow-xs overflow-hidden">
        <div className="bg-[#007bff] text-white px-4 py-3 flex items-center justify-between">
          <h3 className="font-bold text-sm flex items-center gap-2">
            <FileSpreadsheet className="w-4 h-4 text-white" />
            <span>ទិន្នន័យទិដ្ឋាការ ({filteredRecords.length})</span>
          </h3>
        </div>

        <div id="visa-list-print-table" className="overflow-x-auto bg-white">
          <table className="w-full text-left border-collapse">
            <thead>
              <tr className="bg-gray-100 text-gray-700 font-bold text-xs border-b border-gray-200">
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
                <th className="px-3 py-3 text-center border-r border-gray-200">ពន្យារ</th>
                <th className="px-3 py-3 text-center">សកម្មភាព</th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-200 text-xs text-gray-800">
              {paginatedRecords.length > 0 ? (
                paginatedRecords.map((r, idx) => (
                  <tr key={r.id} className="hover:bg-slate-50 transition">
                    <td className="px-3 py-2.5 text-center font-semibold text-gray-500 border-r border-gray-100">
                      {(currentPage - 1) * itemsPerPage + idx + 1}
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
                    <td className="px-3 py-2.5 text-gray-900 border-r border-gray-100">
                      {r.passportNumber}
                    </td>
                    <td className="px-3 py-2.5 whitespace-nowrap border-r border-gray-100">
                      {r.nationality}
                    </td>
                    <td className="px-3 py-2.5 text-center border-r border-gray-100">
                      <span className="px-2 py-0.5 rounded bg-gray-100 text-gray-800 font-mono font-bold border border-gray-200">
                        {r.oldVisaType.split(' ')[0]}
                      </span>
                    </td>
                    <td className="px-3 py-2.5 text-center border-r border-gray-100">
                      <span className="px-2 py-0.5 rounded bg-emerald-100 text-emerald-800 font-mono font-bold border border-emerald-200">
                        {r.newVisaType.split(' ')[0]}
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
                    <td className="px-3 py-2.5 text-center font-bold text-amber-700 border-r border-gray-100">
                      {r.extension || r.duration}
                    </td>
                    <td className="px-3 py-2.5 text-center">
                      <div className="flex items-center justify-center gap-1.5">
                        <button
                          onClick={() => setViewingRecord(r)}
                          className="p-1 text-blue-600 hover:bg-blue-50 rounded transition"
                          title="មើលលម្អិត"
                        >
                          <Eye className="w-4 h-4" />
                        </button>
                        <button
                          onClick={() => setEditingRecord({ ...r })}
                          className="p-1 text-amber-600 hover:bg-amber-50 rounded transition"
                          title="កែប្រែ"
                        >
                          <Edit className="w-4 h-4" />
                        </button>
                        <button
                          onClick={() => setDeletingRecord({ id: r.id, name: `${r.fullName} (${r.passportNumber})` })}
                          className="p-1 text-rose-600 hover:bg-rose-50 rounded transition"
                          title="លុប"
                        >
                          <Trash2 className="w-4 h-4" />
                        </button>
                      </div>
                    </td>
                  </tr>
                ))
              ) : (
                <tr>
                  <td colSpan={13} className="text-center py-12 text-gray-400">
                    មិនមានទិន្នន័យទិដ្ឋាការដែលត្រូវស្វែងរកឡើយ
                  </td>
                </tr>
              )}
            </tbody>
          </table>
        </div>

        {/* SUMMARY BAR & PAGINATION */}
        <div className="bg-slate-50 px-6 py-4 border-t border-gray-200 flex flex-col md:flex-row items-center justify-between gap-4">
          <div className="text-xs text-gray-600 flex items-center gap-4">
            <span>
              បង្ហាញ <strong className="text-[#0B2545]">{paginatedRecords.length}</strong> នៃ{' '}
              <strong className="text-[#0B2545]">{filteredRecords.length}</strong> ទិន្នន័យ
            </span>
          </div>

          {/* Pagination Controls */}
          <div className="flex items-center gap-1.5">
            <button
              onClick={() => setCurrentPage((p) => Math.max(p - 1, 1))}
              disabled={currentPage === 1}
              className="p-1.5 rounded-lg border border-gray-300 bg-white text-gray-600 hover:bg-gray-100 disabled:opacity-40 transition"
            >
              <ChevronLeft className="w-4 h-4" />
            </button>
            <span className="px-3 py-1 text-xs font-bold text-[#0B2545]">
              {currentPage} / {totalPages}
            </span>
            <button
              onClick={() => setCurrentPage((p) => Math.min(p + 1, totalPages))}
              disabled={currentPage === totalPages}
              className="p-1.5 rounded-lg border border-gray-300 bg-white text-gray-600 hover:bg-gray-100 disabled:opacity-40 transition"
            >
              <ChevronRight className="w-4 h-4" />
            </button>
          </div>
        </div>
      </div>

      {/* VIEW DETAIL MODAL */}
      {viewingRecord && (
        <div className="fixed inset-0 z-50 bg-[#071A33]/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl shadow-2xl border border-[#C6A15B]/40 w-full max-w-lg overflow-hidden animate-fade">
            <div className="bg-[#0B2545] text-white px-6 py-4 flex items-center justify-between">
              <h3 className="font-display text-sm text-[#E4CD98]">ព័ត៌មានលម្អិតការប្តូរទិដ្ឋាការ</h3>
              <button onClick={() => setViewingRecord(null)} className="text-white/70 hover:text-white">
                <X className="w-5 h-5" />
              </button>
            </div>
            <div className="p-6 space-y-4 text-xs">
              <div className="grid grid-cols-2 gap-3 p-3 bg-amber-50/50 rounded-xl border border-amber-200">
                <div>
                  <p className="text-gray-500 font-semibold">លេខលិខិតឆ្លងដែន</p>
                  <p className="font-bold text-sm text-[#0B2545]">{viewingRecord.passportNumber}</p>
                </div>
                <div>
                  <p className="text-gray-500 font-semibold">ឈ្មោះពេញ</p>
                  <p className="font-bold text-sm text-[#0B2545] uppercase">{viewingRecord.fullName}</p>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <p className="text-gray-500">សញ្ជាតិ</p>
                  <p className="font-bold text-gray-800">{viewingRecord.nationality}</p>
                </div>
                <div>
                  <p className="text-gray-500">ភេទ</p>
                  <p className="font-bold text-gray-800">{viewingRecord.gender}</p>
                </div>
                <div>
                  <p className="text-gray-500">ប្រភេទទិដ្ឋាការដើម</p>
                  <p className="font-bold text-gray-800">{viewingRecord.oldVisaType}</p>
                </div>
                <div>
                  <p className="text-gray-500">ប្រភេទទិដ្ឋាការថ្មី</p>
                  <p className="font-bold text-emerald-700">{viewingRecord.newVisaType}</p>
                </div>
                <div>
                  <p className="text-gray-500">សុពលភាព / រយៈពេល</p>
                  <p className="font-bold text-gray-800">{viewingRecord.duration}</p>
                </div>
                <div>
                  <p className="text-gray-500">កាលបរិច្ឆេទស្នើសុំ</p>
                  <p className="font-mono text-gray-800">{viewingRecord.applicationDate}</p>
                </div>
                <div>
                  <p className="text-gray-500">មន្ត្រីទទួលបន្ទុក</p>
                  <p className="font-bold text-gray-800">{viewingRecord.officerName || '-'}</p>
                </div>
              </div>

              {viewingRecord.notes && (
                <div className="pt-2 border-t border-gray-100">
                  <p className="text-gray-500 font-semibold">កំណត់សម្គាល់</p>
                  <p className="text-gray-700 bg-gray-50 p-2.5 rounded-lg border border-gray-200 mt-1">
                    {viewingRecord.notes}
                  </p>
                </div>
              )}
            </div>
            <div className="bg-gray-50 px-6 py-3 text-right">
              <button
                onClick={() => setViewingRecord(null)}
                className="px-4 py-2 bg-[#0B2545] text-white rounded-lg font-bold text-xs hover:bg-[#071A33]"
              >
                បិទ
              </button>
            </div>
          </div>
        </div>
      )}

      {/* EDIT MODAL */}
      {editingRecord && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-2xs flex items-center justify-center p-4">
          <div className="bg-white rounded-md shadow-xl border border-gray-300 border-t-4 border-t-[#007bff] w-full max-w-lg overflow-hidden animate-fade">
            <div className="bg-[#007bff] text-white px-4 py-3 flex items-center justify-between">
              <h3 className="font-bold text-sm text-white">កែប្រែទិន្នន័យទិដ្ឋាការ (Edit Visa Record)</h3>
              <button onClick={() => setEditingRecord(null)} className="text-white/80 hover:text-white">
                <X className="w-4 h-4" />
              </button>
            </div>

            <form onSubmit={handleEditSubmit} className="p-5 space-y-3.5 text-xs">
              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-bold text-gray-700 mb-1">លេខលិខិតឆ្លងដែន</label>
                  <input
                    type="text"
                    required
                    value={editingRecord.passportNumber}
                    onChange={(e) => setEditingRecord({ ...editingRecord, passportNumber: e.target.value.toUpperCase() })}
                    className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white uppercase transition"
                  />
                </div>
                <div>
                  <label className="block font-bold text-gray-700 mb-1">ឈ្មោះពេញ</label>
                  <input
                    type="text"
                    required
                    value={editingRecord.fullName}
                    onChange={(e) => setEditingRecord({ ...editingRecord, fullName: e.target.value.toUpperCase() })}
                    className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white uppercase transition"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-bold text-gray-700 mb-1">ទិដ្ឋាការដំបូង → ប្តូរ</label>
                  <div className="flex items-center gap-1">
                    <select
                      value={editingRecord.oldVisaType}
                      onChange={(e) => setEditingRecord({ ...editingRecord, oldVisaType: e.target.value })}
                      className="w-1/2 border border-gray-300 rounded-sm px-2 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                    >
                      <option value="">-- ជ្រើសរើស --</option>
                      {VISA_TYPES.map((vt) => (
                        <option key={vt} value={vt}>{vt}</option>
                      ))}
                      {editingRecord.oldVisaType && !VISA_TYPES.includes(editingRecord.oldVisaType) && (
                        <option value={editingRecord.oldVisaType}>{editingRecord.oldVisaType}</option>
                      )}
                    </select>
                    <span>→</span>
                    <select
                      value={editingRecord.newVisaType}
                      onChange={(e) => setEditingRecord({ ...editingRecord, newVisaType: e.target.value })}
                      className="w-1/2 border border-gray-300 rounded-sm px-2 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                    >
                      <option value="">-- ជ្រើសរើស --</option>
                      {VISA_TYPES.map((vt) => (
                        <option key={vt} value={vt}>{vt}</option>
                      ))}
                      {editingRecord.newVisaType && !VISA_TYPES.includes(editingRecord.newVisaType) && (
                        <option value={editingRecord.newVisaType}>{editingRecord.newVisaType}</option>
                      )}
                    </select>
                  </div>
                </div>
                <div>
                  <label className="block font-bold text-gray-700 mb-1">លេខលិខិត (Doc No.)</label>
                  <input
                    type="text"
                    value={editingRecord.documentNumber || ''}
                    onChange={(e) => setEditingRecord({ ...editingRecord, documentNumber: e.target.value })}
                    className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                  />
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3">
                <div>
                  <label className="block font-bold text-gray-700 mb-1">ថ្ងៃចេញ (Issue Date)</label>
                  <CustomDatePicker
                    value={editingRecord.issueDate || editingRecord.applicationDate}
                    onChange={(d) => setEditingRecord({ ...editingRecord, issueDate: d })}
                    placeholder="YYYY-MM-DD"
                    className="py-1.5 text-xs"
                  />
                </div>
                <div>
                  <label className="block font-bold text-gray-700 mb-1">ស្ថាប័នដើម (Organization)</label>
                  <select
                    value={editingRecord.organization || ''}
                    onChange={(e) => setEditingRecord({ ...editingRecord, organization: e.target.value })}
                    className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                  >
                    <option value="">-- ជ្រើសរើសស្ថាប័ន --</option>
                    {organizations.map((org) => (
                      <option key={org.id} value={org.name}>
                        {org.name}
                      </option>
                    ))}
                    {editingRecord.organization && !organizations.some((o) => o.name === editingRecord.organization) && (
                      <option value={editingRecord.organization}>{editingRecord.organization}</option>
                    )}
                  </select>
                </div>
              </div>

              <div>
                <label className="block font-bold text-gray-700 mb-1">ពន្យារ (Extension)</label>
                <select
                  value={editingRecord.extension || ''}
                  onChange={(e) => {
                    const extCode = e.target.value;
                    const foundExt = EXTENSIONS.find((x) => x.code === extCode);
                    setEditingRecord({
                      ...editingRecord,
                      extension: extCode,
                      fee: foundExt ? foundExt.fee : editingRecord.fee,
                    });
                  }}
                  className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                >
                  <option value="">-- ជ្រើសរើសពន្យារ --</option>
                  {EXTENSIONS.map((ext) => (
                    <option key={ext.code} value={ext.code}>
                      {ext.label}
                    </option>
                  ))}
                  {editingRecord.extension && !EXTENSIONS.some((x) => x.code === editingRecord.extension) && (
                    <option value={editingRecord.extension}>{editingRecord.extension}</option>
                  )}
                </select>
              </div>

              <div>
                <label className="block font-bold text-gray-700 mb-1">កាលបរិច្ឆេទ</label>
                <CustomDatePicker
                  value={editingRecord.applicationDate}
                  onChange={(d) => setEditingRecord({ ...editingRecord, applicationDate: d })}
                  placeholder="YYYY-MM-DD"
                  className="py-1.5 text-xs"
                />
              </div>

              <div className="flex items-center justify-end gap-2 pt-4 border-t border-gray-200">
                <button
                  type="button"
                  onClick={() => setEditingRecord(null)}
                  className="min-w-[120px] bg-[#6c757d] hover:bg-[#5a6268] text-white px-4 py-1.5 rounded text-xs font-bold transition flex items-center justify-center cursor-pointer"
                >
                  <span>បោះបង់</span>
                </button>
                <button
                  type="submit"
                  className="min-w-[120px] bg-[#28a745] hover:bg-[#218838] text-white px-4 py-1.5 rounded text-xs font-bold transition flex items-center justify-center gap-1.5 cursor-pointer shadow-2xs"
                >
                  <Save className="w-3.5 h-3.5 text-white" />
                  <span>រក្សាទុក</span>
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* DELETE CONFIRMATION MODAL */}
      {deletingRecord && (
        <div className="fixed inset-0 z-50 bg-[#071A33]/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl border border-rose-200 shadow-2xl w-full max-w-sm p-6 animate-fade">
            <div className="flex items-center gap-3 text-rose-600 mb-3">
              <div className="w-10 h-10 rounded-full bg-rose-100 flex items-center justify-center shrink-0">
                <Trash2 className="w-5 h-5" />
              </div>
              <h3 className="font-display text-[#0B2545] text-base font-bold">បញ្ជាក់ការលុបទិន្នន័យ</h3>
            </div>
            <p className="text-xs text-gray-600 mb-5 leading-relaxed">
              តើអ្នកពិតជាចង់លុបទិន្នន័យទិដ្ឋាការ <span className="font-bold text-[#0B2545]">"{deletingRecord.name}"</span> នេះមែនទេ?
            </p>
            <div className="flex items-center justify-end gap-2.5">
              <button
                type="button"
                onClick={() => setDeletingRecord(null)}
                className="min-w-[120px] px-4 py-2 text-xs font-bold text-gray-700 bg-gray-100 hover:bg-gray-200 border border-gray-300 rounded-lg transition flex items-center justify-center cursor-pointer"
              >
                <span>បោះបង់</span>
              </button>
              <button
                type="button"
                onClick={handleConfirmDelete}
                className="min-w-[120px] px-4 py-2 text-xs font-bold text-white bg-rose-600 hover:bg-rose-700 rounded-lg shadow-xs transition flex items-center justify-center gap-1.5 cursor-pointer"
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>លុបចេញ</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* Import Modal */}
      <VisaImportModal
        isOpen={isImportModalOpen}
        onClose={() => setIsImportModalOpen(false)}
        onImportRecords={onBatchImportRecords || (() => {})}
        onShowToast={onShowToast}
        existingRecords={records}
        visaTeams={visaTeams}
        organizations={organizations}
        officers={officers}
      />

      {/* BULK DELETE DATE RANGE SELECTION MODAL */}
      {isBulkDeleteModalOpen && (
        <div className="fixed inset-0 z-50 bg-[#071A33]/70 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl border border-rose-300 shadow-2xl w-full max-w-lg p-6 animate-fade">
            <div className="flex items-center justify-between border-b border-gray-100 pb-3 mb-4">
              <div className="flex items-center gap-3 text-rose-600">
                <div className="w-10 h-10 rounded-xl bg-rose-100 flex items-center justify-center shrink-0">
                  <Trash2 className="w-5 h-5 text-rose-600" />
                </div>
                <div>
                  <h3 className="font-display text-[#0B2545] text-base font-bold">
                    លុបទិន្នន័យតាមកាលបរិច្ឆេទ (Bulk Delete)
                  </h3>
                  <p className="text-xs text-gray-500">
                    ជ្រើសរើសថ្ងៃចាប់ផ្តើម និងថ្ងៃបញ្ចប់ដើម្បីលុបទិន្នន័យច្រើនក្នុងពេលតែមួយ
                  </p>
                </div>
              </div>
              <button
                onClick={() => setIsBulkDeleteModalOpen(false)}
                className="text-gray-400 hover:text-gray-600 p-1.5 rounded-lg transition"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-4 text-xs text-gray-700">
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                <div>
                  <label className="block font-bold text-gray-700 mb-1 flex items-center gap-1">
                    <Calendar className="w-3.5 h-3.5 text-rose-600" />
                    <span>ថ្ងៃចាប់ផ្តើម (Start Date)</span>
                  </label>
                  <CustomDatePicker
                    value={bulkStartDate}
                    onChange={(d) => setBulkStartDate(d)}
                    placeholder="YYYY-MM-DD"
                    className="p-2 text-xs"
                  />
                </div>
                <div>
                  <label className="block font-bold text-gray-700 mb-1 flex items-center gap-1">
                    <Calendar className="w-3.5 h-3.5 text-rose-600" />
                    <span>ថ្ងៃបញ្ចប់ (End Date)</span>
                  </label>
                  <CustomDatePicker
                    value={bulkEndDate}
                    onChange={(d) => setBulkEndDate(d)}
                    placeholder="YYYY-MM-DD"
                    className="p-2 text-xs"
                  />
                </div>
              </div>

              {/* Data count box */}
              <div className="bg-rose-50 border border-rose-200 rounded-xl p-3.5 flex items-center justify-between">
                <div className="flex items-center gap-2">
                  <AlertCircle className="w-4 h-4 text-rose-600 shrink-0" />
                  <span className="text-xs font-semibold text-gray-800">
                    ទិន្នន័យត្រូវលុបក្នុងចន្លោះនេះ៖
                  </span>
                </div>
                <span className="font-bold text-rose-700 text-sm font-mono">
                  {matchingBulkRecords.length} កំណត់ត្រា
                </span>
              </div>
            </div>

            <div className="flex items-center justify-end gap-2.5 pt-5 mt-4 border-t border-gray-100">
              <button
                type="button"
                onClick={() => setIsBulkDeleteModalOpen(false)}
                className="px-4 py-2 text-xs font-bold text-gray-600 hover:bg-gray-100 rounded-xl transition"
              >
                បោះបង់
              </button>
              <button
                type="button"
                onClick={handleRequestBulkDelete}
                disabled={matchingBulkRecords.length === 0}
                className="px-5 py-2 text-xs font-bold text-white bg-rose-600 hover:bg-rose-700 disabled:opacity-50 rounded-xl shadow-sm transition flex items-center gap-1.5 cursor-pointer"
              >
                <Trash2 className="w-4 h-4" />
                <span>លុបទិន្នន័យ ({matchingBulkRecords.length})</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* BULK DELETE CONFIRMATION MODAL */}
      {showConfirmBulkModal && (
        <div className="fixed inset-0 z-[60] bg-[#071A33]/80 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl border-2 border-rose-500 shadow-2xl w-full max-w-md p-6 animate-fade">
            <div className="flex items-center gap-3 text-rose-600 mb-3">
              <div className="w-12 h-12 rounded-2xl bg-rose-100 flex items-center justify-center shrink-0 border border-rose-200">
                <AlertTriangle className="w-6 h-6 text-rose-600" />
              </div>
              <div>
                <h3 className="font-display text-[#0B2545] text-base font-bold">
                  សារបញ្ជាក់៖ ពិតជាចង់លុបទិន្នន័យមែនទេ?
                </h3>
                <p className="text-xs text-rose-600 font-semibold">
                  ការលុបនេះមិនអាចស្តារឡើងវិញបានឡើយ!
                </p>
              </div>
            </div>

            <div className="bg-rose-50/70 border border-rose-200 rounded-xl p-3.5 my-4 text-xs text-gray-700 leading-relaxed space-y-2">
              <p>
                តើអ្នកពិតជាចង់លុបទិន្នន័យបញ្ជីស្នើសុំប្តូរទិដ្ឋាការចំនួន{' '}
                <strong className="text-rose-700 font-mono text-sm underline">
                  {matchingBulkRecords.length} កំណត់ត្រា
                </strong>{' '}
                ចាប់ពីថ្ងៃទី <strong className="text-[#0B2545] font-mono">{bulkStartDate}</strong> ដល់ថ្ងៃទី{' '}
                <strong className="text-[#0B2545] font-mono">{bulkEndDate}</strong> មែនទេ?
              </p>
            </div>

            <div className="flex items-center justify-end gap-2.5 pt-2">
              <button
                type="button"
                onClick={() => setShowConfirmBulkModal(false)}
                className="px-4 py-2 text-xs font-bold text-gray-700 hover:bg-gray-100 border border-gray-300 rounded-xl transition"
              >
                បោះបង់
              </button>
              <button
                type="button"
                onClick={handleExecuteBulkDelete}
                className="px-5 py-2 text-xs font-bold text-white bg-rose-600 hover:bg-rose-700 rounded-xl shadow-md transition flex items-center gap-1.5 cursor-pointer"
              >
                <Trash2 className="w-4 h-4" />
                <span>ពិតជាចង់លុប</span>
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
