import React, { useState } from 'react';
import { CategoriesState, Officer, UserRole } from '../types';
import { formatPhoneNumber } from '../utils/khmerCalendar';
import { formatReportTeamName } from '../utils/teamNormalization';
import { isOfficeOfficer, isTeamUser, isOfficerInTeam } from '../utils/officerAccess';
import { printA4Document } from '../utils/printHelper';
import { CustomDatePicker } from './CustomDatePicker';
import {
  Search,
  Filter,
  Edit2,
  Trash2,
  Eye,
  X,
  Save,
  Printer,
  Download,
  Award,
  Briefcase,
  Layers,
  Calendar,
  Phone,
  User as UserIcon,
  Home,
  Users,
  ShieldAlert,
  Lock,
  Hash,
  Building2,
} from 'lucide-react';

interface OfficerListProps {
  officers: Officer[];
  categories: CategoriesState;
  currentRole: UserRole;
  assignedTeam?: string;
  hideOfficeOfficersFromTeams?: boolean;
  onUpdateOfficer: (id: string, updated: Partial<Officer>) => void;
  onDeleteOfficer: (id: string) => void;
  onNavigateCategoryManager?: (type: 'ranks' | 'positions' | 'teamPositions' | 'visaTeams' | 'visaTeamsData' | any) => void;
  onShowToast: (msg: string, type: 'success' | 'error') => void;
}

export const OfficerList: React.FC<OfficerListProps> = ({
  officers,
  categories,
  currentRole,
  assignedTeam,
  hideOfficeOfficersFromTeams = true,
  onUpdateOfficer,
  onDeleteOfficer,
  onNavigateCategoryManager,
  onShowToast,
}) => {
  const isTeam = isTeamUser(currentRole, assignedTeam);
  const isOfficeRestricted = hideOfficeOfficersFromTeams && isTeam;

  const [searchQuery, setSearchQuery] = useState('');
  const [filterRank, setFilterRank] = useState('');
  const [filterPosition, setFilterPosition] = useState('');
  const [filterOfficeWork, setFilterOfficeWork] = useState('');
  const [filterVisaTeam, setFilterVisaTeam] = useState('');

  // Modals state
  const [editingOfficer, setEditingOfficer] = useState<Officer | null>(null);
  const [viewingOfficer, setViewingOfficer] = useState<Officer | null>(null);
  const [deletingOfficer, setDeletingOfficer] = useState<{ id: string; name: string } | null>(null);

  // Edit form internal fields
  const [editName, setEditName] = useState('');
  const [editRankId, setEditRankId] = useState('');
  const [editGender, setEditGender] = useState<'ប' | 'ស'>('ប');
  const [editNumber, setEditNumber] = useState('');
  const [editDob, setEditDob] = useState('');
  const [editPositionId, setEditPositionId] = useState('');
  const [editTeamPositionId, setEditTeamPositionId] = useState('');
  const [editOfficeWork, setEditOfficeWork] = useState('');
  const [editVisaTeamId, setEditVisaTeamId] = useState('');
  const [editPhone, setEditPhone] = useState('');

  // Category helpers
  const getRankName = (id: string) => categories.ranks.find((r) => r.id === id)?.name || '-';
  const getPosName = (id: string) =>
    categories.positions.find((p) => p.id === id)?.name ||
    categories.teamPositions?.find((p) => p.id === id)?.name ||
    categories.collectorRoles?.find((p) => p.id === id)?.name ||
    id ||
    '-';
  const getTeamPosName = (id?: string) => {
    if (!id) return '-';
    return (
      categories.teamPositions?.find((p) => p.id === id)?.name ||
      categories.collectorRoles?.find((p) => p.id === id)?.name ||
      categories.positions.find((p) => p.id === id)?.name ||
      id ||
      '-'
    );
  };
  const getVisaTeamName = (id: string) => {
    if (!id) return '-';
    const found = categories.visaTeams.find((v) => v.id === id || v.name === id);
    const rawName = found ? found.name : id;
    return formatReportTeamName(rawName);
  };

  // Filter Position options (all positions across categories)
  const filterPositionOptions = React.useMemo(() => {
    const combined = [
      ...(categories.positions || []),
      ...(categories.teamPositions || []),
      ...(categories.collectorRoles || []),
    ];
    const seen = new Set<string>();
    const roles: Array<{ id: string; name: string }> = [];

    combined.forEach((item) => {
      if (item && item.name && item.name.trim()) {
        const nameClean = item.name.trim();
        const idKey = item.id || nameClean;
        if (!seen.has(nameClean.toLowerCase())) {
          seen.add(nameClean.toLowerCase());
          roles.push({ id: idKey, name: nameClean });
        }
      }
    });

    return roles;
  }, [categories.positions, categories.teamPositions, categories.collectorRoles]);
  const officePositionOptions = React.useMemo(() => {
    const mainPos = categories.positions || [];
    if (editPositionId && !mainPos.some((r) => r.id === editPositionId || r.name === editPositionId)) {
      const posName = getPosName(editPositionId);
      return [...mainPos, { id: editPositionId, name: posName }];
    }
    return mainPos;
  }, [categories.positions, editPositionId]);

  // Team Position options (for "តួនាទីតាមក្រុម" - categories.teamPositions/collectorRoles)
  const editTeamPositionOptions = React.useMemo(() => {
    const list = (categories.teamPositions && categories.teamPositions.length > 0)
      ? categories.teamPositions
      : categories.collectorRoles || [];
    if (editTeamPositionId && !list.some((r) => r.id === editTeamPositionId || r.name === editTeamPositionId)) {
      const posName = getPosName(editTeamPositionId);
      return [...list, { id: editTeamPositionId, name: posName }];
    }
    return list;
  }, [categories.teamPositions, categories.collectorRoles, editTeamPositionId]);

  // Filter logic:
  // For team users with an assignedTeam, ONLY show officers belonging to that team.
  // For office users (Admin/Secondary), show all officers (filterable via search & dropdowns).
  const filteredOfficers = officers.filter((of) => {
    if (isTeam && assignedTeam && assignedTeam.trim()) {
      const belongsToTeam = isOfficerInTeam(of, assignedTeam, categories.visaTeams);
      if (!belongsToTeam) {
        return false;
      }
    }

    const matchesSearch =
      !searchQuery.trim() ||
      of.name.toLowerCase().includes(searchQuery.toLowerCase()) ||
      of.officerNumber.toLowerCase().includes(searchQuery.toLowerCase()) ||
      of.phone.includes(searchQuery);

    const matchesRank = !filterRank || of.rankId === filterRank;
    const matchesPos = !filterPosition || of.positionId === filterPosition || of.teamPositionId === filterPosition;
    const matchesOffice = !filterOfficeWork || of.officeWork === filterOfficeWork;
    const matchesVisa = !filterVisaTeam || of.visaTeamId === filterVisaTeam || isOfficerInTeam(of, filterVisaTeam, categories.visaTeams);

    return matchesSearch && matchesRank && matchesPos && matchesOffice && matchesVisa;
  });

  const visibleOfficersCount = officers.length;

  const handleOpenEdit = (of: Officer) => {
    setEditingOfficer(of);
    setEditName(of.name || '');
    setEditRankId(of.rankId || '');
    setEditGender((of.gender as 'ប' | 'ស') || 'ប');
    setEditNumber(of.officerNumber || '');
    setEditDob(of.dob || '');
    setEditPositionId(of.positionId || '');
    setEditTeamPositionId(of.teamPositionId || '');
    setEditOfficeWork(of.officeWork || '');
    setEditVisaTeamId(of.visaTeamId || '');
    setEditPhone(formatPhoneNumber(of.phone || ''));
  };

  const handleEditSelectTeamPosition = (val: string) => {
    setEditTeamPositionId(val);
  };

  const handleEditSelectOfficeWork = (val: string) => {
    setEditOfficeWork(val);
  };

  const handleSaveEdit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingOfficer) return;

    if (!editName.trim()) {
      onShowToast('សូមបញ្ចូលឈ្មោះមន្ត្រី', 'error');
      return;
    }
    if (!editRankId) {
      onShowToast('សូមជ្រើសរើសឋានន្តរស័ក្កិ', 'error');
      return;
    }

    onUpdateOfficer(editingOfficer.id, {
      name: editName.trim(),
      rankId: editRankId,
      gender: editGender,
      officerNumber: editNumber.trim(),
      dob: editDob,
      positionId: editPositionId || editTeamPositionId,
      teamPositionId: editTeamPositionId,
      officeWork: editOfficeWork,
      visaTeamId: editVisaTeamId,
      phone: editPhone.trim(),
    });

    setEditingOfficer(null);
    onShowToast('កែប្រែទិន្នន័យមន្ត្រីជោគជ័យ', 'success');
  };

  const handleDelete = (id: string, name: string) => {
    setDeletingOfficer({ id, name });
  };

  const handleConfirmDelete = () => {
    if (deletingOfficer) {
      onDeleteOfficer(deletingOfficer.id);
      onShowToast(`បានលុបទិន្នន័យមន្ត្រី "${deletingOfficer.name}" រួចរាល់`, 'success');
      setDeletingOfficer(null);
    }
  };

  const handlePrint = () => {
    printA4Document('officer-list-print-table', {
      orientation: 'landscape',
      documentTitle: 'បញ្ជីឈ្មោះមន្ត្រី',
    });
  };

  return (
    <div className="space-y-5 w-full animate-fade">
      {/* AdminLTE Content Header (Breadcrumbs) */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 bg-white px-4 py-3 rounded-md border border-gray-200 shadow-xs">
        <div>
          <h1 className="text-base font-bold text-gray-800 flex items-center gap-2">
            <span>បញ្ជីឈ្មោះមន្ត្រី (Officers List)</span>
          </h1>
          <p className="text-xs text-gray-500">
            គ្រប់គ្រង និងស្វែងរកទិន្នន័យមន្ត្រីក្នុងប្រព័ន្ធ
          </p>
        </div>
        <div className="flex items-center gap-1.5 text-xs text-gray-500 bg-gray-100 px-3 py-1.5 rounded-md font-medium self-start sm:self-auto">
          <Home className="w-3.5 h-3.5 text-gray-500" />
          <span>ទំព័រដើម</span>
          <span>/</span>
          <span>គ្រប់គ្រងមន្ត្រី</span>
          <span>/</span>
          <span className="text-[#007bff] font-bold">បញ្ជីឈ្មោះមន្ត្រី</span>
        </div>
      </div>

      {/* Controls Bar: Search & Filters Card */}
      <div className="bg-white rounded-md border border-gray-300 border-t-4 border-t-[#007bff] shadow-xs p-4 space-y-3">
        <div className="flex flex-col md:flex-row items-stretch md:items-center justify-between gap-3">
          {/* Search Box */}
          <div className="relative flex-1">
            <Search className="w-4 h-4 text-gray-400 absolute left-3 top-1/2 -translate-y-1/2" />
            <input
              type="text"
              placeholder="ស្វែងរកតាមឈ្មោះ, អត្តលេខ, ឬទូរស័ព្ទ..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="w-full pl-9 pr-8 py-1.5 text-xs border border-gray-300 rounded-sm focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
            />
            {searchQuery && (
              <button
                onClick={() => setSearchQuery('')}
                className="absolute right-2.5 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600"
              >
                <X className="w-3.5 h-3.5" />
              </button>
            )}
          </div>
        </div>

        {/* Filters Row */}
        <div className="grid grid-cols-1 sm:grid-cols-4 gap-3 pt-2 border-t border-gray-200">
          <div>
            <label className="block text-xs font-bold text-gray-700 mb-1">
              តម្រងតាមឋានន្តរស័ក្កិ
            </label>
            <select
              value={filterRank}
              onChange={(e) => setFilterRank(e.target.value)}
              className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
            >
              <option value="">-- ទាំងអស់ --</option>
              {categories.ranks.map((r) => (
                <option key={r.id} value={r.id}>
                  {r.name}
                </option>
              ))}
            </select>
          </div>

          <div>
            <label className="block text-xs font-bold text-gray-700 mb-1">
              តម្រងតាមតួនាទី
            </label>
            <select
              value={filterPosition}
              onChange={(e) => setFilterPosition(e.target.value)}
              className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
            >
              <option value="">-- ទាំងអស់ --</option>
              {filterPositionOptions.map((p) => (
                <option key={p.id} value={p.id}>
                  {p.name}
                </option>
              ))}
            </select>
          </div>

          <div>
            <label className="block text-xs font-bold text-gray-700 mb-1 flex items-center justify-between">
              <span>បំរើការងារនៅការិយាល័យ</span>
              {isOfficeRestricted && (
                <span className="text-[10px] text-amber-600 font-semibold flex items-center gap-1">
                  <Lock className="w-3 h-3" />
                  <span>សិទ្ធិការិយាល័យ</span>
                </span>
              )}
            </label>
            <select
              disabled={isOfficeRestricted}
              value={isOfficeRestricted ? '' : filterOfficeWork}
              onChange={(e) => setFilterOfficeWork(e.target.value)}
              className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 disabled:bg-gray-100 disabled:text-gray-500 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
            >
              {isOfficeRestricted ? (
                <option value="">🔒 រឹតបន្តឹងសិទ្ធិ (ការិយាល័យប៉ុណ្ណោះ)</option>
              ) : (
                <>
                  <option value="">-- ទាំងអស់ --</option>
                  <option value="ក១">ក១</option>
                  <option value="ក២">ក២</option>
                  <option value="ក៣">ក៣</option>
                  <option value="ក៤">ក៤</option>
                  <option value="ក៥">ក៥</option>
                </>
              )}
            </select>
          </div>

          <div>
            <label className="block text-xs font-bold text-gray-700 mb-1">
              តម្រងតាមក្រុមទិដ្ឋាការ
            </label>
            <select
              value={filterVisaTeam}
              onChange={(e) => setFilterVisaTeam(e.target.value)}
              className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
            >
              <option value="">-- ទាំងអស់ --</option>
              {categories.visaTeams.map((v) => (
                <option key={v.id} value={v.id}>
                  {v.name}
                </option>
              ))}
            </select>
          </div>
        </div>
      </div>

      {/* Privacy Notice Banner for Team Users */}
      {isOfficeRestricted && (
        <div className="bg-amber-50 border-l-4 border-amber-500 p-3.5 rounded-md text-amber-900 text-xs flex items-start gap-2.5 shadow-2xs">
          <ShieldAlert className="w-4 h-4 text-amber-600 shrink-0 mt-0.5" />
          <div className="space-y-0.5">
            <p className="font-bold flex items-center gap-1.5">
              <span>ការកំណត់សុវត្ថិភាពទិន្នន័យ (Access Control Restricted)</span>
              <span className="bg-amber-200/80 text-amber-900 text-[10px] px-1.5 py-0.2 rounded font-semibold">
                SET ON
              </span>
            </p>
            <p className="text-[11px] text-amber-800 leading-relaxed">
              ទិន្នន័យមន្ត្រីនៅការិយាល័យ មិនអនុញ្ញាតឱ្យរបស់ក្រុមមើលឃើញឡើយ។ ប្រព័ន្ធកំពុងបង្ហាញតែទិន្នន័យមន្ត្រីរបស់ក្រុមប៉ុណ្ណោះ។
            </p>
          </div>
        </div>
      )}

      {/* Officers Table Container */}
      <div id="officer-list-print-table" className="bg-white rounded-md border border-gray-300 border-t-4 border-t-[#007bff] shadow-xs overflow-hidden">
        <div className="bg-[#007bff] text-white px-4 py-3 flex items-center justify-between">
          <h3 className="font-bold text-sm flex items-center gap-2">
            <Users className="w-4 h-4 text-white" />
            <span>
              {isOfficeRestricted
                ? `បញ្ជីឈ្មោះមន្ត្រីរបស់ក្រុម (${filteredOfficers.length})`
                : `បញ្ជីឈ្មោះមន្ត្រី (${filteredOfficers.length})`}
            </span>
          </h3>
          <div className="flex items-center gap-3">
            <span className="text-xs text-white/90 font-medium hidden sm:inline">
              {isOfficeRestricted
                ? `${visibleOfficersCount} មន្ត្រីរបស់ក្រុម (មន្ត្រីការិយាល័យត្រូវបានលាក់)`
                : `${officers.length} មន្ត្រីក្នុងប្រព័ន្ធ`}
            </span>
          </div>
        </div>

        {filteredOfficers.length === 0 ? (
          <div className="text-center py-12 text-gray-400 text-xs">
            មិនមានទិន្នន័យមន្ត្រីត្រូវតាមការស្វែងរក ឬតម្រងឡើយ។
          </div>
        ) : (
          <div className="overflow-x-auto">
            <table className="w-full text-xs text-left whitespace-nowrap">
              <thead>
                <tr className="bg-[#e9ecef] text-[#0056b3] border-b border-gray-300 font-bold text-xs">
                  <th className="px-3 py-3 font-bold text-[#0056b3] text-center w-12">ល.រ</th>
                  <th className="px-3.5 py-3 font-bold text-[#0056b3]">ឈ្មោះ</th>
                  <th className="px-3.5 py-3 font-bold text-[#0056b3]">ឋានន្តរស័ក្តិ</th>
                  <th className="px-3.5 py-3 font-bold text-[#0056b3] text-center">ភេទ</th>
                  <th className="px-3.5 py-3 font-bold text-[#0056b3]">អត្តលេខ</th>
                  <th className="px-3.5 py-3 font-bold text-[#0056b3]">ថ្ងៃកំណើត</th>
                  <th className="px-3.5 py-3 font-bold text-[#0056b3]">តួនាទី</th>
                  <th className="px-3.5 py-3 font-bold text-[#0056b3]">តួនាទីតាមក្រុម</th>
                  <th className="px-3.5 py-3 font-bold text-[#0056b3]">បំរើការងារនៅការិយាល័យ</th>
                  <th className="px-3.5 py-3 font-bold text-[#0056b3]">ក្រុមផ្តល់ទិដ្ឋាការ</th>
                  <th className="px-3.5 py-3 font-bold text-[#0056b3]">ទូរស័ព្ទ</th>
                  <th className="px-3.5 py-3 font-bold text-[#0056b3] text-center">សកម្មភាព</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-100">
                {filteredOfficers.map((of, idx) => (
                  <tr key={of.id} className="hover:bg-amber-50/40 transition">
                    <td className="px-3 py-3 text-center text-gray-500 font-medium">{idx + 1}</td>
                    <td className="px-3.5 py-3 font-bold text-[#0B2545]">{of.name}</td>
                    <td className="px-3.5 py-3 text-gray-700">{getRankName(of.rankId)}</td>
                    <td className="px-3.5 py-3 text-center">
                      <span
                        className={`inline-block px-2 py-0.5 rounded text-[10px] font-bold ${
                          of.gender === 'ប'
                            ? 'bg-blue-100 text-blue-800'
                            : 'bg-rose-100 text-rose-800'
                        }`}
                      >
                        {of.gender === 'ប' ? 'ប្រុស' : 'ស្រី'}
                      </span>
                    </td>
                    <td className="px-3.5 py-3 font-mono text-gray-600">{of.officerNumber}</td>
                    <td className="px-3.5 py-3 text-gray-600">{of.dob}</td>
                    <td className="px-3.5 py-3 text-gray-700">{getPosName(of.positionId)}</td>
                    <td className="px-3.5 py-3 text-gray-700">{getTeamPosName(of.teamPositionId)}</td>
                    <td className="px-3.5 py-3 text-gray-700">{of.officeWork || '-'}</td>
                    <td className="px-3.5 py-3 text-gray-700">{getVisaTeamName(of.visaTeamId)}</td>
                    <td className="px-3.5 py-3 text-gray-600 font-mono">{formatPhoneNumber(of.phone)}</td>
                    <td className="px-3.5 py-3 text-center">
                      <div className="flex items-center justify-center gap-1.5">
                        <button
                          onClick={() => setViewingOfficer(of)}
                          title="មើលលម្អិត"
                          className="p-1.5 text-blue-600 hover:bg-blue-50 rounded transition"
                        >
                          <Eye className="w-3.5 h-3.5" />
                        </button>

                        <button
                          onClick={() => handleOpenEdit(of)}
                          title="កែប្រែ"
                          className="p-1.5 text-[#0B2545] hover:bg-amber-100 rounded transition"
                        >
                          <Edit2 className="w-3.5 h-3.5" />
                        </button>

                        <button
                          onClick={() => handleDelete(of.id, of.name)}
                          title="លុប"
                          className="p-1.5 text-rose-600 hover:bg-rose-50 rounded transition"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    </td>
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        )}
      </div>

      {/* VIEW OFFICER MODAL */}
      {viewingOfficer && (
        <div className="fixed inset-0 z-50 bg-[#071A33]/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl border border-[#C6A15B]/50 shadow-2xl w-full max-w-2xl p-6 animate-fade">
            <div className="flex items-center justify-between border-b border-gray-200 pb-3 mb-4">
              <h3 className="font-display text-[#0B2545] text-base flex items-center gap-2">
                <UserIcon className="w-5 h-5 text-[#C6A15B]" />
                <span>ព័ត៌មានលម្អិតរបស់មន្ត្រី</span>
              </h3>
              <button
                onClick={() => setViewingOfficer(null)}
                className="text-gray-400 hover:text-gray-600 p-1 rounded"
              >
                <X className="w-5 h-5" />
              </button>
            </div>

            <div className="space-y-3 text-xs">
              <div className="bg-amber-50/50 p-3 rounded-lg border border-[#C6A15B]/30 flex items-center gap-3">
                <div className="w-10 h-10 rounded-full bg-[#0B2545] text-[#E4CD98] font-display text-sm flex items-center justify-center font-bold shrink-0">
                  {viewingOfficer.name.substring(0, 1)}
                </div>
                <div>
                  <p className="font-bold text-sm text-[#0B2545]">{viewingOfficer.name}</p>
                  <p className="text-gray-500 font-mono text-[11px]">
                    អត្តលេខ៖ {viewingOfficer.officerNumber}
                  </p>
                </div>
              </div>

              <div className="grid grid-cols-2 gap-3 pt-2">
                <div>
                  <span className="text-gray-400 block text-[10px]">ឋានន្តរស័ក្តិ</span>
                  <span className="font-semibold text-gray-800">
                    {getRankName(viewingOfficer.rankId)}
                  </span>
                </div>
                <div>
                  <span className="text-gray-400 block text-[10px]">ភេទ</span>
                  <span className="font-semibold text-gray-800">
                    {viewingOfficer.gender === 'ប' ? 'ប្រុស' : 'ស្រី'}
                  </span>
                </div>
                <div>
                  <span className="text-gray-400 block text-[10px]">ថ្ងៃខែឆ្នាំកំណើត</span>
                  <span className="font-semibold text-gray-800">{viewingOfficer.dob}</span>
                </div>
                <div>
                  <span className="text-gray-400 block text-[10px]">ទូរស័ព្ទលេខ</span>
                  <span className="font-semibold text-gray-800 font-mono">{formatPhoneNumber(viewingOfficer.phone)}</span>
                </div>
                <div>
                  <span className="text-gray-400 block text-[10px]">តួនាទី</span>
                  <span className="font-semibold text-gray-800">
                    {getPosName(viewingOfficer.positionId)}
                  </span>
                </div>
                <div>
                  <span className="text-gray-400 block text-[10px]">តួនាទីតាមក្រុម</span>
                  <span className="font-semibold text-gray-800">
                    {getTeamPosName(viewingOfficer.teamPositionId)}
                  </span>
                </div>
                <div className="col-span-2">
                  <span className="text-gray-400 block text-[10px]">បំរើការងារនៅការិយាល័យ</span>
                  <span className="font-semibold text-gray-800">
                    {viewingOfficer.officeWork || '-'}
                  </span>
                </div>
                <div className="col-span-2">
                  <span className="text-gray-400 block text-[10px]">ក្រុមផ្តល់ទិដ្ឋាការ</span>
                  <span className="font-semibold text-gray-800">
                    {getVisaTeamName(viewingOfficer.visaTeamId)}
                  </span>
                </div>
              </div>
            </div>

            <div className="mt-6 pt-3 border-t border-gray-100 flex justify-end">
              <button
                onClick={() => setViewingOfficer(null)}
                className="bg-[#0B2545] text-[#E4CD98] font-bold text-xs px-4 py-2 rounded-lg"
              >
                បិទ
              </button>
            </div>
          </div>
        </div>
      )}

      {/* EDIT OFFICER MODAL */}
      {editingOfficer && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-2xs flex items-center justify-center p-4">
          <div className="bg-white rounded-md border border-gray-300 border-t-4 border-t-[#007bff] shadow-xl w-full max-w-4xl overflow-hidden animate-fade max-h-[90vh] flex flex-col">
            {/* Modal Header */}
            <div className="bg-[#007bff] text-white px-4 py-3 flex items-center justify-between shrink-0">
              <div className="flex items-center gap-2 font-bold text-sm">
                <Edit2 className="w-4 h-4 text-white" />
                <span>ទម្រង់កែប្រែទិន្នន័យមន្ត្រី (Officer Edit Form)</span>
              </div>
              <button
                onClick={() => setEditingOfficer(null)}
                className="text-white/80 hover:text-white p-1 rounded transition cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <form onSubmit={handleSaveEdit} className="p-6 space-y-5 overflow-y-auto text-xs">
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
                {/* Officer Name */}
                <div>
                  <label className="block text-xs font-bold text-gray-700 uppercase tracking-wide mb-1.5">
                    ឈ្មោះមន្ត្រី (FULL NAME) <span className="text-red-500">*</span>
                  </label>
                  <div className="flex rounded-sm shadow-2xs">
                    <span className="inline-flex items-center px-3 rounded-l-sm border border-r-0 border-gray-300 bg-gray-100 text-gray-500 text-xs">
                      <UserIcon className="w-3.5 h-3.5" />
                    </span>
                    <input
                      type="text"
                      required
                      placeholder="ឧទាហរណ៍៖ ស៊ន សុភ័ក្ត្រ"
                      value={editName}
                      onChange={(e) => setEditName(e.target.value)}
                      className="w-full border border-gray-300 rounded-r-sm px-3 py-2 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none transition"
                    />
                  </div>
                </div>

                {/* Phone Number */}
                <div>
                  <label className="block text-xs font-bold text-gray-700 uppercase tracking-wide mb-1.5">
                    ទូរស័ព្ទលេខ (PHONE NUMBER) <span className="text-red-500">*</span>
                  </label>
                  <div className="flex rounded-sm shadow-2xs">
                    <span className="inline-flex items-center px-3 rounded-l-sm border border-r-0 border-gray-300 bg-gray-100 text-gray-500 text-xs">
                      <Phone className="w-3.5 h-3.5" />
                    </span>
                    <input
                      type="tel"
                      required
                      placeholder="ឧទាហរណ៍៖ 012 345 678"
                      value={editPhone}
                      onChange={(e) => setEditPhone(formatPhoneNumber(e.target.value))}
                      className="w-full border border-gray-300 rounded-r-sm px-3 py-2 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none transition"
                    />
                  </div>
                </div>

                {/* Rank Select */}
                <div>
                  <div className="flex items-center justify-between mb-1.5">
                    <label className="text-xs font-bold text-gray-700 uppercase tracking-wide">
                      ឋានន្តរស័ក្តិ (RANK) <span className="text-red-500">*</span>
                    </label>
                    {onNavigateCategoryManager && (
                      <button
                        type="button"
                        onClick={() => onNavigateCategoryManager('ranks')}
                        className="text-[11px] text-[#007bff] hover:underline font-bold cursor-pointer"
                      >
                        + បន្ថែមប្រភេទ
                      </button>
                    )}
                  </div>
                  <div className="flex rounded-sm shadow-2xs">
                    <span className="inline-flex items-center px-3 rounded-l-sm border border-r-0 border-gray-300 bg-gray-100 text-gray-500 text-xs">
                      <Award className="w-3.5 h-3.5" />
                    </span>
                    <select
                      required
                      value={editRankId}
                      onChange={(e) => setEditRankId(e.target.value)}
                      className="w-full border border-gray-300 rounded-r-sm px-3 py-2 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                    >
                      <option value="">-- ជ្រើសរើស --</option>
                      {categories.ranks.map((r) => (
                        <option key={r.id} value={r.id}>
                          {r.name}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>

                {/* Gender Radios */}
                <div>
                  <label className="block text-xs font-bold text-gray-700 uppercase tracking-wide mb-1.5">
                    ភេទ (GENDER) <span className="text-red-500">*</span>
                  </label>
                  <div className="flex items-center gap-6 h-[34px] px-3 border border-gray-300 rounded-sm bg-gray-50/50">
                    <label className="flex items-center gap-2 cursor-pointer text-xs font-semibold text-gray-800">
                      <input
                        type="radio"
                        name="editOfGender"
                        value="ប"
                        checked={editGender === 'ប'}
                        onChange={() => setEditGender('ប')}
                        className="w-3.5 h-3.5 text-[#007bff] focus:ring-[#007bff]"
                      />
                      <span>ប្រុស (Male)</span>
                    </label>
                    <label className="flex items-center gap-2 cursor-pointer text-xs font-semibold text-gray-800">
                      <input
                        type="radio"
                        name="editOfGender"
                        value="ស"
                        checked={editGender === 'ស'}
                        onChange={() => setEditGender('ស')}
                        className="w-3.5 h-3.5 text-[#007bff] focus:ring-[#007bff]"
                      />
                      <span>ស្រី (Female)</span>
                    </label>
                  </div>
                </div>

                {/* Officer ID Code */}
                <div>
                  <label className="block text-xs font-bold text-gray-700 uppercase tracking-wide mb-1.5">
                    អត្តលេខ (OFFICER ID) <span className="text-red-500">*</span>
                  </label>
                  <div className="flex rounded-sm shadow-2xs">
                    <span className="inline-flex items-center px-3 rounded-l-sm border border-r-0 border-gray-300 bg-gray-100 text-gray-500 text-xs">
                      <Hash className="w-3.5 h-3.5" />
                    </span>
                    <input
                      type="text"
                      required
                      placeholder="ឧទាហរណ៍៖ KHM-00891"
                      value={editNumber}
                      onChange={(e) => setEditNumber(e.target.value)}
                      className="w-full border border-gray-300 rounded-r-sm px-3 py-2 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none transition uppercase"
                    />
                  </div>
                </div>

                {/* Date of Birth */}
                <div>
                  <label className="block text-xs font-bold text-gray-700 uppercase tracking-wide mb-1.5">
                    ថ្ងៃខែឆ្នាំកំណើត (DOB) <span className="text-red-500">*</span>
                  </label>
                  <CustomDatePicker
                    required
                    value={editDob}
                    onChange={setEditDob}
                    className="border border-gray-300 rounded-sm px-3 py-2 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                  />
                </div>

                {/* Position Select */}
                <div>
                  <div className="flex items-center justify-between mb-1.5">
                    <label className="text-xs font-bold text-gray-700 uppercase tracking-wide">
                      តួនាទី (POSITION) <span className="text-red-500">*</span>
                    </label>
                    {onNavigateCategoryManager && (
                      <button
                        type="button"
                        onClick={() => onNavigateCategoryManager('positions')}
                        className="text-[11px] text-[#007bff] hover:underline font-bold cursor-pointer"
                      >
                        + បន្ថែមប្រភេទ
                      </button>
                    )}
                  </div>
                  <div className="flex rounded-sm shadow-2xs">
                    <span className="inline-flex items-center px-3 rounded-l-sm border border-r-0 border-gray-300 bg-gray-100 text-gray-500 text-xs">
                      <Briefcase className="w-3.5 h-3.5" />
                    </span>
                    <select
                      value={editPositionId}
                      onChange={(e) => setEditPositionId(e.target.value)}
                      className="w-full border border-gray-300 rounded-r-sm px-3 py-2 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                    >
                      <option value="">-- ជ្រើសរើស --</option>
                      {officePositionOptions.map((p) => (
                        <option key={p.id} value={p.id}>
                          {p.name}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>

                {/* Office Work Dropdown */}
                <div>
                  <div className="flex items-center justify-between mb-1.5">
                    <label className="text-xs font-bold text-gray-700 uppercase tracking-wide">
                      បំរើការងារនៅការិយាល័យ
                    </label>
                  </div>
                  <div className="flex rounded-sm shadow-2xs">
                    <span className="inline-flex items-center px-3 rounded-l-sm border border-r-0 border-gray-300 bg-gray-100 text-gray-500 text-xs">
                      <Building2 className="w-3.5 h-3.5" />
                    </span>
                    <select
                      value={editOfficeWork}
                      onChange={(e) => handleEditSelectOfficeWork(e.target.value)}
                      className="w-full border border-gray-300 rounded-r-sm px-3 py-2 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                    >
                      <option value="">-- ជ្រើសរើស --</option>
                      <option value="ក១">ក១</option>
                      <option value="ក២">ក២</option>
                      <option value="ក៣">ក៣</option>
                      <option value="ក៤">ក៤</option>
                      <option value="ក៥">ក៥</option>
                    </select>
                  </div>
                </div>

                {/* Team Position Select (Left side) */}
                <div>
                  <div className="flex items-center justify-between mb-1.5">
                    <label className="text-xs font-bold text-gray-700 uppercase tracking-wide">
                      តួនាទីតាមក្រុម (TEAM POSITION)
                    </label>
                    {onNavigateCategoryManager && (
                      <button
                        type="button"
                        onClick={() => onNavigateCategoryManager('teamPositions')}
                        className="text-[11px] text-[#007bff] hover:underline font-bold cursor-pointer"
                      >
                        + បន្ថែមប្រភេទ
                      </button>
                    )}
                  </div>
                  <div className="flex rounded-sm shadow-2xs">
                    <span className="inline-flex items-center px-3 rounded-l-sm border border-r-0 border-gray-300 bg-gray-100 text-gray-500 text-xs">
                      <Briefcase className="w-3.5 h-3.5 text-emerald-600" />
                    </span>
                    <select
                      value={editTeamPositionId}
                      onChange={(e) => handleEditSelectTeamPosition(e.target.value)}
                      className="w-full border border-gray-300 rounded-r-sm px-3 py-2 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                    >
                      <option value="">-- ជ្រើសរើស --</option>
                      {editTeamPositionOptions.map((p) => (
                        <option key={p.id} value={p.id}>
                          {p.name}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>

                {/* Visa Team Select (Right side) */}
                <div>
                  <div className="flex items-center justify-between mb-1.5">
                    <label className="text-xs font-bold text-gray-700 uppercase tracking-wide">
                      ក្រុមផ្តល់ទិដ្ឋាការ (VISA TEAM) <span className="text-red-500">*</span>
                    </label>
                    {onNavigateCategoryManager && (
                      <button
                        type="button"
                        onClick={() => onNavigateCategoryManager('visaTeams')}
                        className="text-[11px] text-[#007bff] hover:underline font-bold cursor-pointer"
                      >
                        + បន្ថែមប្រភេទ
                      </button>
                    )}
                  </div>
                  <div className="flex rounded-sm shadow-2xs">
                    <span className="inline-flex items-center px-3 rounded-l-sm border border-r-0 border-gray-300 bg-gray-100 text-gray-500 text-xs">
                      <Layers className="w-3.5 h-3.5" />
                    </span>
                    <select
                      value={editVisaTeamId}
                      onChange={(e) => setEditVisaTeamId(e.target.value)}
                      className="w-full border border-gray-300 rounded-r-sm px-3 py-2 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                    >
                      <option value="">-- ជ្រើសរើស --</option>
                      {categories.visaTeams.map((vt) => (
                        <option key={vt.id} value={vt.id}>
                          {vt.name}
                        </option>
                      ))}
                    </select>
                  </div>
                </div>
              </div>

              {/* Modal Footer */}
              <div className="pt-4 border-t border-gray-200 flex items-center justify-end gap-3 bg-gray-50 -mx-6 -mb-6 px-6 py-3.5 rounded-b-sm">
                <button
                  type="button"
                  onClick={() => setEditingOfficer(null)}
                  className="min-w-[140px] bg-[#6c757d] hover:bg-[#5a6268] text-white font-bold text-xs px-4 py-2 rounded shadow-2xs transition flex items-center justify-center gap-1.5 cursor-pointer"
                >
                  <X className="w-3.5 h-3.5" />
                  <span>បោះបង់ (Cancel)</span>
                </button>

                <button
                  type="submit"
                  className="min-w-[180px] bg-[#007bff] hover:bg-[#0069d9] text-white font-bold text-xs px-4 py-2 rounded shadow-2xs transition flex items-center justify-center gap-2 cursor-pointer"
                >
                  <Save className="w-4 h-4" />
                  <span>រក្សាទុក (Save)</span>
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* DELETE OFFICER CONFIRMATION MODAL */}
      {deletingOfficer && (
        <div className="fixed inset-0 z-50 bg-[#071A33]/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl border border-rose-200 shadow-2xl w-full max-w-sm p-6 animate-fade">
            <div className="flex items-center gap-3 text-rose-600 mb-3">
              <div className="w-10 h-10 rounded-full bg-rose-100 flex items-center justify-center shrink-0">
                <Trash2 className="w-5 h-5" />
              </div>
              <h3 className="font-display text-[#0B2545] text-base font-bold">បញ្ជាក់ការលុបទិន្នន័យ</h3>
            </div>
            <p className="text-xs text-gray-600 mb-5 leading-relaxed">
              តើអ្នកពិតជាចង់លុបទិន្នន័យមន្ត្រី <span className="font-bold text-[#0B2545]">"{deletingOfficer.name}"</span> នេះមែនទេ? សកម្មភាពនេះមិនអាចត្រឡប់វិញបានឡើយ។
            </p>
            <div className="flex items-center justify-end gap-2.5">
              <button
                type="button"
                onClick={() => setDeletingOfficer(null)}
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
    </div>
  );
};
