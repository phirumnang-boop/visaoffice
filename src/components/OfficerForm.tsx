import React, { useState, useEffect } from 'react';
import { CategoriesState, Officer, UserRole } from '../types';
import { formatPhoneNumber } from '../utils/khmerCalendar';
import { isTeamUser } from '../utils/officerAccess';
import { CustomDatePicker } from './CustomDatePicker';
import {
  PlusCircle,
  Award,
  Briefcase,
  Building2,
  Layers,
  Save,
  RefreshCw,
  User,
  Hash,
  Calendar,
  Phone,
  Lock,
  ShieldAlert,
} from 'lucide-react';

interface OfficerFormProps {
  categories: CategoriesState;
  currentRole: UserRole;
  currentUserId: string;
  assignedTeam?: string;
  hideOfficeOfficersFromTeams?: boolean;
  onAddOfficer: (officer: Omit<Officer, 'id' | 'createdAt'>) => void;
  onNavigateCategoryManager: (type: 'ranks' | 'positions' | 'visaTeams' | 'visaTeamsData' | any) => void;
  onShowToast: (msg: string, type: 'success' | 'error') => void;
}

export const OfficerForm: React.FC<OfficerFormProps> = ({
  categories,
  currentRole,
  currentUserId,
  assignedTeam,
  hideOfficeOfficersFromTeams = true,
  onAddOfficer,
  onNavigateCategoryManager,
  onShowToast,
}) => {
  const isTeam = isTeamUser(currentRole, assignedTeam);

  const [name, setName] = useState('');
  const [rankId, setRankId] = useState('');
  const [gender, setGender] = useState<'ប' | 'ស'>('ប');
  const [officerNumber, setOfficerNumber] = useState('');
  const [dob, setDob] = useState('');
  const [positionId, setPositionId] = useState('');
  const [teamPositionId, setTeamPositionId] = useState('');
  const [officeWork, setOfficeWork] = useState('');
  const [visaTeamId, setVisaTeamId] = useState('');
  const [phone, setPhone] = useState('');
  const [isSubmitting, setIsSubmitting] = useState(false);

  // Position options for 'តួនាទី (Position)'
  const positionOptions = React.useMemo(() => {
    const pos = categories.positions || [];
    if (pos.length > 0) return pos;
    return [
      { id: 'pos-1', name: 'ប្រធានក្រុម' },
      { id: 'pos-2', name: 'អនុប្រធានក្រុម' },
      { id: 'pos-3', name: 'មន្ត្រី' },
    ];
  }, [categories.positions]);

  // Team position options from Category Manager ('តួនាទីតាមក្រុម')
  const teamPositionOptions = React.useMemo(() => {
    const list = (categories.teamPositions && categories.teamPositions.length > 0)
      ? categories.teamPositions
      : categories.collectorRoles || [];
    if (list.length === 0) {
      return [
        { id: 'tp-1', name: 'ប្រធានក្រុម' },
        { id: 'tp-2', name: 'អនុប្រធានក្រុម' },
        { id: 'tp-3', name: 'មន្ត្រី' },
        { id: 'tp-4', name: 'អ្នកមកបើក' },
      ];
    }
    return list;
  }, [categories.teamPositions, categories.collectorRoles]);

  useEffect(() => {
    if (isTeam && assignedTeam && !visaTeamId) {
      const matchedTeam =
        categories.visaTeams.find((vt) => vt.name.trim().toLowerCase() === assignedTeam.trim().toLowerCase()) ||
        categories.visaTeamsRobok.find((vt) => vt.name.trim().toLowerCase() === assignedTeam.trim().toLowerCase());
      if (matchedTeam) {
        setVisaTeamId(matchedTeam.id);
      }
    }
  }, [isTeam, assignedTeam, categories.visaTeams, categories.visaTeamsRobok, visaTeamId]);

  const handleSelectTeamPosition = (val: string) => {
    setTeamPositionId(val);
  };

  const handleSelectOfficeWork = (val: string) => {
    setOfficeWork(val);
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();

    if (!name.trim()) {
      onShowToast('សូមបញ្ចូលឈ្មោះមន្ត្រី', 'error');
      return;
    }
    if (!rankId) {
      onShowToast('សូមជ្រើសរើសឋានន្តរស័ក្កិ', 'error');
      return;
    }
    if (!officerNumber.trim()) {
      onShowToast('សូមបញ្ចូលអត្តលេខ', 'error');
      return;
    }
    if (!dob) {
      onShowToast('សូមជ្រើសរើសថ្ងៃខែឆ្នាំកំណើត', 'error');
      return;
    }
    if (!positionId && !teamPositionId) {
      onShowToast('សូមជ្រើសរើសតួនាទី ឬ តួនាទីតាមក្រុម', 'error');
      return;
    }
    if ((teamPositionId || !officeWork) && !visaTeamId) {
      onShowToast('សូមជ្រើសរើសក្រុមផ្តល់ទិដ្ឋាការ', 'error');
      return;
    }
    if (!phone.trim()) {
      onShowToast('សូមបញ្ចូលលេខទូរស័ព្ទ', 'error');
      return;
    }

    setIsSubmitting(true);

    try {
      onAddOfficer({
        name: name.trim(),
        rankId,
        gender,
        officerNumber: officerNumber.trim(),
        dob,
        positionId: positionId || teamPositionId,
        teamPositionId,
        officeWork,
        visaTeamId,
        phone: phone.trim(),
        createdBy: currentUserId,
      });

      setName('');
      setRankId('');
      setGender('ប');
      setOfficerNumber('');
      setDob('');
      setPositionId('');
      setTeamPositionId('');
      setOfficeWork('');
      setVisaTeamId('');
      setPhone('');

      onShowToast('រក្សាទុកទិន្នន័យមន្ត្រីបានជោគជ័យ', 'success');
    } catch (err: any) {
      onShowToast(err?.message || 'មានបញ្ហាក្នុងការរក្សាទុក', 'error');
    } finally {
      setIsSubmitting(false);
    }
  };

  return (
    <div className="w-full space-y-5 animate-fade">
      {/* Breadcrumb Header */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 bg-white px-4 py-3 rounded-md border border-gray-200 shadow-xs">
        <div>
          <h1 className="text-base font-bold text-gray-800 flex items-center gap-2">
            <span>បញ្ចូលទិន្នន័យមន្ត្រី (Officer Form)</span>
          </h1>
          <p className="text-xs text-gray-500">
            ទម្រង់ផ្លូវការសម្រាប់បញ្ចូល និងកត់ត្រាព័ត៌មានមន្ត្រី
          </p>
        </div>
        <div className="flex items-center gap-1.5 text-xs text-gray-500 bg-gray-100 px-3 py-1.5 rounded-md font-medium self-start sm:self-auto">
          <span>ទំព័រដើម</span>
          <span>/</span>
          <span>ព័ត៌មានមន្ត្រី</span>
          <span>/</span>
          <span className="text-[#007bff] font-bold">ទម្រង់បញ្ចូល</span>
        </div>
      </div>

      {/* Form Card */}
      <div className="bg-white rounded-md border border-gray-300 border-t-4 border-t-[#007bff] shadow-xs overflow-hidden">
        <div className="bg-[#007bff] text-white px-4 py-3 flex items-center justify-between">
          <div className="flex items-center gap-2 font-bold text-sm">
            <PlusCircle className="w-4 h-4 text-white" />
            <span>ទម្រង់បញ្ចូលទិន្នន័យមន្ត្រី (Officer Registration)</span>
          </div>
        </div>

        <form onSubmit={handleSubmit} className="p-6 space-y-5">
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-4">
            {/* Officer Name */}
            <div>
              <label className="block text-xs font-bold text-gray-700 uppercase tracking-wide mb-1.5">
                ឈ្មោះមន្ត្រី (Full Name) <span className="text-red-500">*</span>
              </label>
              <div className="flex rounded-sm shadow-2xs">
                <span className="inline-flex items-center px-3 rounded-l-sm border border-r-0 border-gray-300 bg-gray-100 text-gray-500 text-xs">
                  <User className="w-3.5 h-3.5" />
                </span>
                <input
                  type="text"
                  required
                  placeholder="ឧទាហរណ៍៖ ស៊ន សុភ័ក្ត្រ"
                  value={name}
                  onChange={(e) => setName(e.target.value)}
                  className="w-full border border-gray-300 rounded-r-sm px-3 py-2 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none transition"
                />
              </div>
            </div>

            {/* Phone Number (Placed side-by-side with Officer Name) */}
            <div>
              <label className="block text-xs font-bold text-gray-700 uppercase tracking-wide mb-1.5">
                ទូរស័ព្ទលេខ (Phone Number) <span className="text-red-500">*</span>
              </label>
              <div className="flex rounded-sm shadow-2xs">
                <span className="inline-flex items-center px-3 rounded-l-sm border border-r-0 border-gray-300 bg-gray-100 text-gray-500 text-xs">
                  <Phone className="w-3.5 h-3.5" />
                </span>
                <input
                  type="tel"
                  required
                  placeholder="ឧទាហរណ៍៖ 012 345 678"
                  value={phone}
                  onChange={(e) => setPhone(formatPhoneNumber(e.target.value))}
                  className="w-full border border-gray-300 rounded-r-sm px-3 py-2 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none transition"
                />
              </div>
            </div>

            {/* Rank Select */}
            <div>
              <div className="flex items-center justify-between mb-1.5">
                <label className="text-xs font-bold text-gray-700 uppercase tracking-wide">
                  ឋានន្តរស័ក្តិ (Rank) <span className="text-red-500">*</span>
                </label>
                <button
                  type="button"
                  onClick={() => onNavigateCategoryManager('ranks')}
                  className="text-[11px] text-[#007bff] hover:underline font-bold cursor-pointer"
                >
                  + បន្ថែមប្រភេទ
                </button>
              </div>
              <div className="flex rounded-sm shadow-2xs">
                <span className="inline-flex items-center px-3 rounded-l-sm border border-r-0 border-gray-300 bg-gray-100 text-gray-500 text-xs">
                  <Award className="w-3.5 h-3.5" />
                </span>
                <select
                  required
                  value={rankId}
                  onChange={(e) => setRankId(e.target.value)}
                  className="w-full border border-gray-300 rounded-r-sm px-3 py-2 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                >
                  <option value="">-- ជ្រើសរើស --</option>
                  {categories.ranks.map((rank) => (
                    <option key={rank.id} value={rank.id}>
                      {rank.name}
                    </option>
                  ))}
                </select>
              </div>
            </div>

            {/* Gender Radios */}
            <div>
              <label className="block text-xs font-bold text-gray-700 uppercase tracking-wide mb-1.5">
                ភេទ (Gender) <span className="text-red-500">*</span>
              </label>
              <div className="flex items-center gap-6 h-[34px] px-3 border border-gray-300 rounded-sm bg-gray-50/50">
                <label className="flex items-center gap-2 cursor-pointer text-xs font-semibold text-gray-800">
                  <input
                    type="radio"
                    name="ofGender"
                    value="ប"
                    checked={gender === 'ប'}
                    onChange={() => setGender('ប')}
                    className="w-3.5 h-3.5 text-[#007bff] focus:ring-[#007bff]"
                  />
                  <span>ប្រុស (Male)</span>
                </label>
                <label className="flex items-center gap-2 cursor-pointer text-xs font-semibold text-gray-800">
                  <input
                    type="radio"
                    name="ofGender"
                    value="ស"
                    checked={gender === 'ស'}
                    onChange={() => setGender('ស')}
                    className="w-3.5 h-3.5 text-[#007bff] focus:ring-[#007bff]"
                  />
                  <span>ស្រី (Female)</span>
                </label>
              </div>
            </div>

            {/* Officer ID Code */}
            <div>
              <label className="block text-xs font-bold text-gray-700 uppercase tracking-wide mb-1.5">
                អត្តលេខ (Officer ID) <span className="text-red-500">*</span>
              </label>
              <div className="flex rounded-sm shadow-2xs">
                <span className="inline-flex items-center px-3 rounded-l-sm border border-r-0 border-gray-300 bg-gray-100 text-gray-500 text-xs">
                  <Hash className="w-3.5 h-3.5" />
                </span>
                <input
                  type="text"
                  required
                  placeholder="ឧទាហរណ៍៖ KHM-00891"
                  value={officerNumber}
                  onChange={(e) => setOfficerNumber(e.target.value)}
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
                value={dob}
                onChange={(d) => setDob(d)}
                className="py-1.5 text-xs"
              />
            </div>

            {/* Position Select */}
            <div>
              <div className="flex items-center justify-between mb-1.5">
                <label className="text-xs font-bold text-gray-700 uppercase tracking-wide">
                  តួនាទី (Position) <span className="text-red-500">*</span>
                </label>
                <button
                  type="button"
                  onClick={() => onNavigateCategoryManager('positions')}
                  className="text-[11px] text-[#007bff] hover:underline font-bold cursor-pointer"
                >
                  + បន្ថែមប្រភេទ
                </button>
              </div>
              <div className="flex rounded-sm shadow-2xs">
                <span className="inline-flex items-center px-3 rounded-l-sm border border-r-0 border-gray-300 bg-gray-100 text-gray-500 text-xs">
                  <Briefcase className="w-3.5 h-3.5" />
                </span>
                <select
                  value={positionId}
                  onChange={(e) => setPositionId(e.target.value)}
                  className="w-full border border-gray-300 rounded-r-sm px-3 py-2 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                >
                  <option value="">-- ជ្រើសរើស --</option>
                  {positionOptions.map((pos) => (
                    <option key={pos.id} value={pos.id}>
                      {pos.name}
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
                  value={officeWork}
                  onChange={(e) => handleSelectOfficeWork(e.target.value)}
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
                  តួនាទីតាមក្រុម (Team Position)
                </label>
                <button
                  type="button"
                  onClick={() => onNavigateCategoryManager('teamPositions')}
                  className="text-[11px] text-[#007bff] hover:underline font-bold cursor-pointer"
                >
                  + បន្ថែមប្រភេទ
                </button>
              </div>
              <div className="flex rounded-sm shadow-2xs">
                <span className="inline-flex items-center px-3 rounded-l-sm border border-r-0 border-gray-300 bg-gray-100 text-gray-500 text-xs">
                  <Briefcase className="w-3.5 h-3.5 text-emerald-600" />
                </span>
                <select
                  value={teamPositionId}
                  onChange={(e) => handleSelectTeamPosition(e.target.value)}
                  className="w-full border border-gray-300 rounded-r-sm px-3 py-2 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                >
                  <option value="">-- ជ្រើសរើស --</option>
                  {teamPositionOptions.map((tp) => (
                    <option key={tp.id} value={tp.id}>
                      {tp.name}
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
                <button
                  type="button"
                  onClick={() => onNavigateCategoryManager('visaTeams')}
                  className="text-[11px] text-[#007bff] hover:underline font-bold cursor-pointer"
                >
                  + បន្ថែមប្រភេទ
                </button>
              </div>
              <div className="flex rounded-sm shadow-2xs">
                <span className="inline-flex items-center px-3 rounded-l-sm border border-r-0 border-gray-300 bg-gray-100 text-gray-500 text-xs">
                  <Layers className="w-3.5 h-3.5" />
                </span>
                <select
                  value={visaTeamId}
                  onChange={(e) => setVisaTeamId(e.target.value)}
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

          {/* Card Footer */}
          <div className="pt-4 border-t border-gray-200 flex items-center justify-end gap-3 bg-gray-50 -mx-6 -mb-5 px-6 py-3.5 rounded-b-sm">
            <button
              type="button"
              onClick={() => {
                setName('');
                setRankId('');
                setGender('ប');
                setOfficerNumber('');
                setDob('');
                setPositionId('');
                setTeamPositionId('');
                setOfficeWork('');
                setVisaTeamId('');
                setPhone('');
              }}
              className="min-w-[180px] bg-[#6c757d] hover:bg-[#5a6268] text-white font-bold text-xs px-4 py-2 rounded shadow-2xs transition flex items-center justify-center gap-1.5 cursor-pointer"
            >
              <RefreshCw className="w-3.5 h-3.5" />
              <span>សម្អាតទម្រង់</span>
            </button>

            <button
              type="submit"
              disabled={isSubmitting}
              className="min-w-[180px] bg-[#007bff] hover:bg-[#0069d9] text-white font-bold text-xs px-4 py-2 rounded shadow-2xs transition flex items-center justify-center gap-2 disabled:opacity-50 cursor-pointer"
            >
              <Save className="w-4 h-4" />
              <span>{isSubmitting ? 'កំពុងរក្សាទុក...' : 'រក្សាទុកទិន្នន័យ (Save)'}</span>
            </button>
          </div>
        </form>
      </div>
    </div>
  );
};
