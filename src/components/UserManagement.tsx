import React, { useState } from 'react';
import { CategoriesState, UserAccount, UserRole, UserPermissions } from '../types';
import { UserPlus, Users, Edit, Trash2, X, Save, Shield, Eye, EyeOff, Building, Lock, CheckCircle2, Search, SlidersHorizontal } from 'lucide-react';
import { OFFICIAL_29_TEAMS } from '../utils/teamNormalization';

interface UserManagementProps {
  users: UserAccount[];
  viewMode: 'create' | 'list';
  categories?: CategoriesState;
  hideOfficeOfficersFromTeams?: boolean;
  hideOfficerMenuFromTeams?: boolean;
  hideUserMenuFromTeams?: boolean;
  hideDailyOpsFromTeams?: boolean;
  hideRefusalFromTeams?: boolean;
  hideVisaWorkFromTeams?: boolean;
  hideOfficeStockFromTeams?: boolean;
  hideSearchCodeFromTeams?: boolean;
  hideCategoriesFromTeams?: boolean;
  onToggleHideOfficeOfficersFromTeams?: (enabled: boolean) => void;
  onToggleHideOfficerMenuFromTeams?: (enabled: boolean) => void;
  onToggleHideUserMenuFromTeams?: (enabled: boolean) => void;
  onToggleHideDailyOpsFromTeams?: (enabled: boolean) => void;
  onToggleHideRefusalFromTeams?: (enabled: boolean) => void;
  onToggleHideVisaWorkFromTeams?: (enabled: boolean) => void;
  onToggleHideOfficeStockFromTeams?: (enabled: boolean) => void;
  onToggleHideSearchCodeFromTeams?: (enabled: boolean) => void;
  onToggleHideCategoriesFromTeams?: (enabled: boolean) => void;
  onAddUser: (user: Omit<UserAccount, 'id' | 'createdAt'>) => void;
  onUpdateUser: (id: string, updated: Partial<UserAccount>) => void;
  onDeleteUser: (id: string) => void;
  onShowToast: (msg: string, type: 'success' | 'error') => void;
}

export const UserManagement: React.FC<UserManagementProps> = ({
  users,
  viewMode,
  categories,
  hideOfficeOfficersFromTeams = true,
  hideOfficerMenuFromTeams = false,
  hideUserMenuFromTeams = true,
  hideDailyOpsFromTeams = false,
  hideRefusalFromTeams = false,
  hideVisaWorkFromTeams = true,
  hideOfficeStockFromTeams = true,
  hideSearchCodeFromTeams = false,
  hideCategoriesFromTeams = true,
  onToggleHideOfficeOfficersFromTeams,
  onToggleHideOfficerMenuFromTeams,
  onToggleHideUserMenuFromTeams,
  onToggleHideDailyOpsFromTeams,
  onToggleHideRefusalFromTeams,
  onToggleHideVisaWorkFromTeams,
  onToggleHideOfficeStockFromTeams,
  onToggleHideSearchCodeFromTeams,
  onToggleHideCategoriesFromTeams,
  onAddUser,
  onUpdateUser,
  onDeleteUser,
  onShowToast,
}) => {
  const teamOptions = React.useMemo(() => {
    if (categories?.visaTeamsRobok && categories.visaTeamsRobok.length > 0) {
      return categories.visaTeamsRobok.map((t) => t.name.trim()).filter(Boolean);
    }
    return Array.from(OFFICIAL_29_TEAMS);
  }, [categories?.visaTeamsRobok]);

  // Filter & Search State
  const [filterVisibility, setFilterVisibility] = useState<'all' | 'visible' | 'hidden'>('all');
  const [searchQuery, setSearchQuery] = useState('');

  // Create Form State
  const [username, setUsername] = useState('');
  const [officerNumber, setOfficerNumber] = useState('');
  const [password, setPassword] = useState('');
  const [showPassword, setShowPassword] = useState(false);
  const [role, setRole] = useState<UserRole>('User');
  const [assignedTeam, setAssignedTeam] = useState<string>('');

  // Edit Modal State
  const [editingUser, setEditingUser] = useState<UserAccount | null>(null);
  const [editUsername, setEditUsername] = useState('');
  const [editOfficerNumber, setEditOfficerNumber] = useState('');
  const [editPassword, setEditPassword] = useState('');
  const [showEditPassword, setShowEditPassword] = useState(false);
  const [editRole, setEditRole] = useState<UserRole>('User');
  const [editAssignedTeam, setEditAssignedTeam] = useState<string>('');
  const [editIsVisible, setEditIsVisible] = useState(true);
  const [deletingUser, setDeletingUser] = useState<{ id: string; name: string } | null>(null);

  // Password visibility map for the user list
  const [visiblePasswords, setVisiblePasswords] = useState<Record<string, boolean>>({});

  const togglePasswordVisibility = (userId: string) => {
    setVisiblePasswords((prev) => ({
      ...prev,
      [userId]: !prev[userId],
    }));
  };

  // Toggle single user account visibility (ឃើញ ឬមិនឃើញ / Show or Hide this specific username)
  const handleToggleUserVisibility = (u: UserAccount) => {
    const isCurrentlyVisible = u.isVisible !== false && !u.hidden;
    const nextVisible = !isCurrentlyVisible;
    onUpdateUser(u.id, {
      isVisible: nextVisible,
      hidden: !nextVisible,
    });
    onShowToast(
      `គណនី "${u.username}" ត្រូវបានកំណត់៖ ${nextVisible ? '👁️ មើលឃើញ (Visible)' : '🚫 លាក់ (Hidden)'}`,
      nextVisible ? 'success' : 'error'
    );
  };

  // Toggle specific menu permission for this individual user
  const handleToggleUserPermission = (
    u: UserAccount,
    key: keyof UserPermissions,
    globalHide: boolean,
    menuLabel: string
  ) => {
    const currentVal = u.permissions?.[key] !== undefined ? !!u.permissions[key] : !globalHide;
    const nextVal = !currentVal;
    onUpdateUser(u.id, {
      permissions: {
        ...(u.permissions || {}),
        [key]: nextVal,
      },
    });
    onShowToast(
      `បានកំណត់សិទ្ធិ «${menuLabel}» សម្រាប់ ${u.username}៖ ${nextVal ? '👁️ បង្ហាញ (ឃើញ)' : '🚫 លាក់ (មិនឃើញ)'}`,
      nextVal ? 'success' : 'error'
    );
  };

  // Filtered users for list display
  const filteredUsers = React.useMemo(() => {
    return users.filter((u) => {
      // visibility filter
      const isVis = u.isVisible !== false && !u.hidden;
      if (filterVisibility === 'visible' && !isVis) return false;
      if (filterVisibility === 'hidden' && isVis) return false;

      // search query filter
      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase().trim();
        const matchName = u.username.toLowerCase().includes(q);
        const matchNumber = (u.officerNumber || u.email || '').toLowerCase().includes(q);
        const matchTeam = (u.assignedTeam || '').toLowerCase().includes(q);
        return matchName || matchNumber || matchTeam;
      }
      return true;
    });
  }, [users, filterVisibility, searchQuery]);

  const handleCreateSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!username.trim() || !officerNumber.trim() || !password) {
      onShowToast('សូមបំពេញព័ត៌មានឱ្យបានគ្រប់គ្រាន់', 'error');
      return;
    }
    if (password.length < 6) {
      onShowToast('ពាក្យសម្ងាត់ត្រូវមានយ៉ាងតិច ៦ តួអក្សរ', 'error');
      return;
    }

    onAddUser({
      username: username.trim(),
      officerNumber: officerNumber.trim(),
      email: officerNumber.trim(),
      password: password,
      role,
      assignedTeam: role === 'User' ? (assignedTeam || undefined) : undefined,
      isVisible: true,
      hidden: false,
    });

    setUsername('');
    setOfficerNumber('');
    setPassword('');
    setRole('User');
    setAssignedTeam('');

    onShowToast(`បង្កើតគណនី "${username}" ជោគជ័យ!`, 'success');
  };

  const handleOpenEdit = (user: UserAccount) => {
    setEditingUser(user);
    setEditUsername(user.username);
    setEditOfficerNumber(user.officerNumber || user.email || '');
    setEditPassword(user.password || '123456');
    setShowEditPassword(false);
    setEditRole(user.role);
    setEditAssignedTeam(user.assignedTeam || '');
    setEditIsVisible(user.isVisible !== false && !user.hidden);
  };

  const handleSaveEdit = (e: React.FormEvent) => {
    e.preventDefault();
    if (!editingUser) return;

    if (!editUsername.trim() || !editOfficerNumber.trim()) {
      onShowToast('សូមបំពេញព័ត៌មានឱ្យបានគ្រប់គ្រាន់', 'error');
      return;
    }

    if (editPassword && editPassword.length < 6) {
      onShowToast('ពាក្យសម្ងាត់ត្រូវមានយ៉ាងតិច ៦ តួអក្សរ', 'error');
      return;
    }

    onUpdateUser(editingUser.id, {
      username: editUsername.trim(),
      officerNumber: editOfficerNumber.trim(),
      email: editOfficerNumber.trim(),
      password: editPassword || editingUser.password || '123456',
      role: editRole,
      assignedTeam: editRole === 'User' ? (editAssignedTeam || undefined) : undefined,
      isVisible: editIsVisible,
      hidden: !editIsVisible,
    });

    setEditingUser(null);
    onShowToast('កែប្រែគណនីជោគជ័យ', 'success');
  };

  const handleDelete = (id: string, name: string) => {
    setDeletingUser({ id, name });
  };

  const handleConfirmDelete = () => {
    if (deletingUser) {
      onDeleteUser(deletingUser.id);
      onShowToast(`បានលុបគណនី "${deletingUser.name}" រួចរាល់`, 'success');
      setDeletingUser(null);
    }
  };

  if (viewMode === 'create') {
    return (
      <div className="max-w-xl mx-auto bg-white rounded-md border border-gray-300 border-t-4 border-t-[#007bff] shadow-xs animate-fade overflow-hidden">
        <div className="bg-[#007bff] text-white px-4 py-3 flex items-center justify-between">
          <h3 className="font-bold text-sm flex items-center gap-2">
            <UserPlus className="w-4 h-4 text-white" />
            <span>បង្កើតគណនីប្រើប្រាស់ថ្មី (Create New Account)</span>
          </h3>
        </div>

        <form onSubmit={handleCreateSubmit} className="p-6 space-y-4 text-xs">
          <div>
            <label className="block font-bold text-gray-700 mb-1">
              ឈ្មោះអ្នកប្រើប្រាស់ <span className="text-red-500">*</span>
            </label>
            <input
              type="text"
              required
              placeholder="ឧទាហរណ៍៖ សុខ ចាន់ដារ៉ា ឬ ក្រុមផ្តល់ទិដ្ឋាការ ប៉ោយប៉ែត"
              value={username}
              onChange={(e) => setUsername(e.target.value)}
              className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
            />
          </div>

          <div>
            <label className="block font-bold text-gray-700 mb-1">
              អត្តលេខ <span className="text-red-500">*</span>
            </label>
            <input
              type="text"
              required
              placeholder="ឧទាហរណ៍៖ KHM-00891"
              value={officerNumber}
              onChange={(e) => setOfficerNumber(e.target.value)}
              className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
            />
          </div>

          <div>
            <label className="block font-bold text-gray-700 mb-1">
              ពាក្យសម្ងាត់ (យ៉ាងតិច ៦ តួអក្សរ) <span className="text-red-500">*</span>
            </label>
            <div className="relative">
              <input
                type={showPassword ? 'text' : 'password'}
                required
                minLength={6}
                placeholder="••••••••"
                value={password}
                onChange={(e) => setPassword(e.target.value)}
                className="w-full border border-gray-300 rounded-sm pl-2.5 pr-9 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
              />
              <button
                type="button"
                onClick={() => setShowPassword(!showPassword)}
                className="absolute right-2.5 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-700 p-1 transition"
                title={showPassword ? 'លាក់ពាក្យសម្ងាត់' : 'បង្ហាញពាក្យសម្ងាត់'}
              >
                {showPassword ? <EyeOff className="w-3.5 h-3.5" /> : <Eye className="w-3.5 h-3.5" />}
              </button>
            </div>
          </div>

          <div>
            <label className="block font-bold text-gray-700 mb-1">
              តួនាទីសិទ្ធិ <span className="text-red-500">*</span>
            </label>
            <select
              value={role}
              onChange={(e) => setRole(e.target.value as UserRole)}
              className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
            >
              <option value="User">User (មន្ត្រី / សមាជិកក្រុម)</option>
              <option value="User (ការិយាល័យ)">User (ការិយាល័យ)</option>
              <option value="Secondary">Secondary (ការិយាល័យ / Office)</option>
            </select>
          </div>

          {role === 'User' && (
            <div className="bg-amber-50/70 p-3 rounded border border-amber-200">
              <label className="block font-bold text-amber-900 mb-1 flex items-center gap-1.5">
                <Building className="w-3.5 h-3.5 text-amber-700" />
                <span>កំណត់ក្រុមផ្ទាល់ខ្លួន (Assigned Team)</span>
              </label>
              <select
                value={assignedTeam}
                onChange={(e) => setAssignedTeam(e.target.value)}
                className="w-full border border-amber-300 rounded px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:outline-none bg-white transition"
              >
                <option value="">-- មើលគ្រប់ក្រុម (All Teams) --</option>
                {teamOptions.map((t) => (
                  <option key={t} value={t}>
                    {t}
                  </option>
                ))}
              </select>
              <p className="text-[11px] text-amber-700 mt-1">
                * ប្រសិនបើជ្រើសរើសក្រុមណាមួយ ពេលគណនីនេះចូលប្រើប្រាស់ នឹងបង្ហាញតែទិន្នន័យក្រុមនោះដោយស្វ័យប្រវត្តិ។
              </p>
            </div>
          )}

          <div className="pt-4 flex justify-end border-t border-gray-200">
            <button
              type="submit"
              className="bg-[#28a745] hover:bg-[#218838] text-white font-bold text-xs px-5 py-2 rounded transition flex items-center gap-1.5 shadow-2xs cursor-pointer"
            >
              <UserPlus className="w-4 h-4 text-white" />
              <span>បង្កើតគណនី</span>
            </button>
          </div>
        </form>
      </div>
    );
  }

  return (
    <div className="w-full space-y-6 animate-fade">
      {/* Officer Access Control Card */}
      <div className="bg-white rounded-md border border-gray-300 border-t-4 border-t-amber-500 shadow-xs p-5">
        <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-3 border-b border-gray-100 pb-3">
          <div className="flex items-center gap-2.5">
            <div className="w-8 h-8 rounded bg-amber-100 text-amber-700 flex items-center justify-center font-bold">
              <Shield className="w-4.5 h-4.5" />
            </div>
            <div>
              <h4 className="font-bold text-gray-900 text-xs sm:text-sm">
                ការកំណត់សិទ្ធិមើលទិន្នន័យ និងមឺនុយការិយាល័យ (Office Access & Menu Control)
              </h4>
              <p className="text-[11px] text-gray-500">
                កំណត់កម្រិតសុវត្ថិភាព និងការលាក់ទិន្នន័យ/មឺនុយ រវាងការិយាល័យ និងក្រុម
              </p>
            </div>
          </div>
          <div className="flex items-center gap-2">
            <span
              className={`inline-flex items-center gap-1 px-2.5 py-1 rounded-full text-[11px] font-bold ${
                hideOfficeOfficersFromTeams
                  ? 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                  : 'bg-rose-100 text-rose-800 border border-rose-300'
              }`}
            >
              {hideOfficeOfficersFromTeams ? (
                <>
                  <CheckCircle2 className="w-3.5 h-3.5 text-emerald-600" />
                  <span>បើកដំណើរការ (SET ON)</span>
                </>
              ) : (
                <>
                  <X className="w-3.5 h-3.5 text-rose-600" />
                  <span>បានបិទ (SET OFF)</span>
                </>
              )}
            </span>
          </div>
        </div>

        {/* 1. Toggle to hide Office Officers' data from teams */}
        <div className="mt-3 flex flex-col md:flex-row md:items-center justify-between gap-4 bg-amber-50/70 p-3.5 rounded-md border border-amber-200">
          <div className="space-y-1">
            <div className="font-bold text-amber-950 text-xs flex items-center gap-1.5">
              <Lock className="w-3.5 h-3.5 text-amber-700" />
              <span>ទិន្នន័យមន្ត្រី នៅការិយាល័យ មិនអនុញ្ញាតអោយ របស់ក្រុម មើលឃើញឡើយ</span>
            </div>
            <p className="text-[11px] text-amber-800 leading-relaxed">
              នៅពេលបើកដំណើរការ (SET ON)៖ រាល់ទិន្នន័យមន្ត្រីដែលបំរើការងារនៅការិយាល័យ (ក១, ក២, ក៣, ក៤, ក៥ ឬផ្នែករដ្ឋបាលការិយាល័យ) នឹងត្រូវបានរឹតបន្តឹង មិនឱ្យគណនីរបស់ក្រុម (User/Team) មើលឃើញ ស្វែងរក ឬកែប្រែជាដាច់ខាត។ មានតែគណនីការិយាល័យ (Secondary) ប៉ុណ្ណោះដែលអាចមើលឃើញបាន។
            </p>
          </div>
          <button
            type="button"
            onClick={() => onToggleHideOfficeOfficersFromTeams?.(!hideOfficeOfficersFromTeams)}
            className={`px-4 py-2 rounded text-xs font-bold transition shrink-0 cursor-pointer shadow-2xs flex items-center justify-center gap-1.5 ${
              hideOfficeOfficersFromTeams
                ? 'bg-emerald-600 hover:bg-emerald-700 text-white'
                : 'bg-gray-600 hover:bg-gray-700 text-white'
            }`}
          >
            {hideOfficeOfficersFromTeams ? '✅ កំពុងបើក (SET ON)' : '🔘 ចុចបើកដំណើរការ (Turn ON)'}
          </button>
        </div>



      </div>

      <div className="bg-white rounded-md border border-gray-300 border-t-4 border-t-[#007bff] shadow-xs overflow-hidden">
        <div className="bg-[#007bff] text-white px-4 py-3 flex flex-col sm:flex-row sm:items-center justify-between gap-2">
          <div className="flex items-center gap-2.5">
            <div className="w-7 h-7 rounded bg-white/20 flex items-center justify-center">
              <Users className="w-4 h-4 text-white" />
            </div>
            <div>
              <h3 className="font-bold text-sm leading-tight flex items-center gap-2">
                <span>មឺនុយ៖ ទិន្នន័យគណនីអ្នកប្រើប្រាស់</span>
                <span className="text-xs bg-white/20 px-2 py-0.5 rounded-full font-semibold">({users.length})</span>
              </h3>
              <p className="text-[11px] text-white/90 font-medium">
                {hideUserMenuFromTeams 
                  ? '* សម្រាប់តែ «ការិយាល័យ (Office)» (Secondary & User (ការិយាល័យ)) មើលឃើញ និងគ្រប់គ្រង'
                  : '* បានអនុញ្ញាតឱ្យក្រុម (User) មើលឃើញ និងសហការគ្រប់គ្រង'}
              </p>
            </div>
          </div>
          <span className="bg-white/15 text-white border border-white/25 px-2.5 py-1 rounded text-xs font-bold flex items-center gap-1.5 shrink-0 self-start sm:self-auto">
            <Building className="w-3.5 h-3.5 text-amber-300" />
            <span>{hideUserMenuFromTeams ? 'សិទ្ធិមើល៖ ការិយាល័យ (Office Only)' : 'សិទ្ធិមើល៖ ទាំងអស់គ្នា (All Users)'}</span>
          </span>
        </div>

        {/* Search, Filter & Quick Help Bar */}
        <div className="p-3 bg-gray-50/90 border-b border-gray-200 flex flex-col sm:flex-row items-stretch sm:items-center justify-between gap-3">
          <div className="flex flex-wrap items-center gap-2">
            <span className="text-[11px] font-bold text-gray-700 flex items-center gap-1">
              <SlidersHorizontal className="w-3.5 h-3.5 text-gray-500" />
              <span>ចម្រោះការបង្ហាញ៖</span>
            </span>
            <div className="inline-flex rounded-md shadow-2xs border border-gray-300 bg-white p-0.5 text-xs">
              <button
                type="button"
                onClick={() => setFilterVisibility('all')}
                className={`px-2.5 py-1 rounded text-xs font-bold transition cursor-pointer ${
                  filterVisibility === 'all'
                    ? 'bg-[#007bff] text-white shadow-2xs'
                    : 'text-gray-600 hover:text-gray-900'
                }`}
              >
                ទាំងអស់ ({users.length})
              </button>
              <button
                type="button"
                onClick={() => setFilterVisibility('visible')}
                className={`px-2.5 py-1 rounded text-xs font-bold transition cursor-pointer flex items-center gap-1 ${
                  filterVisibility === 'visible'
                    ? 'bg-emerald-600 text-white shadow-2xs'
                    : 'text-gray-600 hover:text-gray-900'
                }`}
              >
                <Eye className="w-3 h-3" />
                <span>ឃើញ ({users.filter((u) => u.isVisible !== false && !u.hidden).length})</span>
              </button>
              <button
                type="button"
                onClick={() => setFilterVisibility('hidden')}
                className={`px-2.5 py-1 rounded text-xs font-bold transition cursor-pointer flex items-center gap-1 ${
                  filterVisibility === 'hidden'
                    ? 'bg-rose-600 text-white shadow-2xs'
                    : 'text-gray-600 hover:text-gray-900'
                }`}
              >
                <EyeOff className="w-3 h-3" />
                <span>មិនឃើញ ({users.filter((u) => u.isVisible === false || u.hidden).length})</span>
              </button>
            </div>
          </div>

          <div className="relative min-w-[220px] sm:w-64">
            <Search className="w-3.5 h-3.5 absolute left-2.5 top-1/2 -translate-y-1/2 text-gray-400" />
            <input
              type="text"
              placeholder="ស្វែងរកឈ្មោះ អត្តលេខ ឬក្រុម..."
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              className="w-full pl-8 pr-2.5 py-1 text-xs border border-gray-300 rounded-md focus:border-[#007bff] focus:outline-none bg-white"
            />
          </div>
        </div>

        <div className="overflow-x-auto">
          <table className="w-full text-xs text-left border-collapse align-middle">
            <thead>
              <tr className="bg-[#0B2545]/10 text-[#0B2545] font-bold border-b border-[#C6A15B]/30 align-middle">
                <th className="px-4 py-3 align-middle text-[#0B2545] min-w-[140px]">
                  <div className="flex flex-col">
                    <span>ឈ្មោះអ្នកប្រើប្រាស់</span>
                    <span className="text-[10px] text-gray-500 font-normal">ចុចប្តូរ ឃើញ / មិនឃើញ</span>
                  </div>
                </th>
                <th className="px-4 py-3 align-middle text-[#0B2545]">អត្តលេខ</th>
                <th className="px-4 py-3 align-middle text-[#0B2545]">ក្រុមប្រចាំការ (Team)</th>
                <th className="px-4 py-3 align-middle text-[#0B2545]">លេខសម្ងាត់ (Password)</th>
                <th className="px-4 py-3 text-center align-middle text-[#0B2545]">សកម្មភាព</th>
                <th className="px-4 py-3 align-middle text-[#0B2545]">តួនាទី</th>

                {/* Integrated Sidebar Menu Columns with Global Toggles */}
                {/* 1. ទិន្នន័យគណនីអ្នកប្រើប្រាស់ (userList) */}
                <th className="px-3 py-2 text-center border-l border-gray-200/50 align-middle min-w-[125px]">
                  <div className="flex flex-col items-center gap-1">
                    <span className="font-bold text-[11px] leading-tight text-[#0B2545]">ទិន្នន័យគណនីអ្នកប្រើប្រាស់</span>
                    <button
                      type="button"
                      onClick={() => onToggleHideUserMenuFromTeams?.(!hideUserMenuFromTeams)}
                      className={`px-2 py-0.5 rounded-[3px] text-[9px] font-bold transition cursor-pointer shadow-2xs border ${
                        hideUserMenuFromTeams
                          ? 'bg-rose-600 border-rose-700 text-white hover:bg-rose-700'
                          : 'bg-emerald-600 border-emerald-700 text-white hover:bg-emerald-700'
                      }`}
                      title="ចុចកំណត់លាក់/បង្ហាញរួមសម្រាប់គ្រប់ក្រុមទាំងអស់ (All Teams default)"
                    >
                      {hideUserMenuFromTeams ? 'លាក់រួម (ON)' : 'បង្ហាញរួម (OFF)'}
                    </button>
                  </div>
                </th>

                {/* 2. ប្រតិបត្តិការប្រចាំថ្ងៃតាមក្រុម (dailyOps) */}
                <th className="px-3 py-2 text-center border-l border-gray-200/50 align-middle min-w-[125px]">
                  <div className="flex flex-col items-center gap-1">
                    <span className="font-bold text-[11px] leading-tight text-[#0B2545]">ប្រតិបត្តិការប្រចាំថ្ងៃតាមក្រុម</span>
                    <button
                      type="button"
                      onClick={() => onToggleHideDailyOpsFromTeams?.(!hideDailyOpsFromTeams)}
                      className={`px-2 py-0.5 rounded-[3px] text-[9px] font-bold transition cursor-pointer shadow-2xs border ${
                        hideDailyOpsFromTeams
                          ? 'bg-rose-600 border-rose-700 text-white hover:bg-rose-700'
                          : 'bg-emerald-600 border-emerald-700 text-white hover:bg-emerald-700'
                      }`}
                      title="ចុចកំណត់លាក់/បង្ហាញរួមសម្រាប់គ្រប់ក្រុមទាំងអស់ (All Teams default)"
                    >
                      {hideDailyOpsFromTeams ? 'លាក់រួម (ON)' : 'បង្ហាញរួម (OFF)'}
                    </button>
                  </div>
                </th>

                {/* 3. បដិសេធនិងបញ្ជូនចេញ (refusal) */}
                <th className="px-3 py-2 text-center border-l border-gray-200/50 align-middle min-w-[125px]">
                  <div className="flex flex-col items-center gap-1">
                    <span className="font-bold text-[11px] leading-tight text-[#0B2545]">បដិសេធនិងបញ្ជូនចេញ</span>
                    <button
                      type="button"
                      onClick={() => onToggleHideRefusalFromTeams?.(!hideRefusalFromTeams)}
                      className={`px-2 py-0.5 rounded-[3px] text-[9px] font-bold transition cursor-pointer shadow-2xs border ${
                        hideRefusalFromTeams
                          ? 'bg-rose-600 border-rose-700 text-white hover:bg-rose-700'
                          : 'bg-emerald-600 border-emerald-700 text-white hover:bg-emerald-700'
                      }`}
                      title="ចុចកំណត់លាក់/បង្ហាញរួមសម្រាប់គ្រប់ក្រុមទាំងអស់ (All Teams default)"
                    >
                      {hideRefusalFromTeams ? 'លាក់រួម (ON)' : 'បង្ហាញរួម (OFF)'}
                    </button>
                  </div>
                </th>

                {/* 4. គណីអ្នកប្រើប្រាស់ (userMenu) */}
                <th className="px-3 py-2 text-center border-l border-gray-200/50 align-middle min-w-[125px]">
                  <div className="flex flex-col items-center gap-1">
                    <span className="font-bold text-[11px] leading-tight text-[#0B2545]">គណីអ្នកប្រើប្រាស់</span>
                    <button
                      type="button"
                      onClick={() => onToggleHideUserMenuFromTeams?.(!hideUserMenuFromTeams)}
                      className={`px-2 py-0.5 rounded-[3px] text-[9px] font-bold transition cursor-pointer shadow-2xs border ${
                        hideUserMenuFromTeams
                          ? 'bg-rose-600 border-rose-700 text-white hover:bg-rose-700'
                          : 'bg-emerald-600 border-emerald-700 text-white hover:bg-emerald-700'
                      }`}
                      title="ចុចកំណត់លាក់/បង្ហាញរួមសម្រាប់គ្រប់ក្រុមទាំងអស់ (All Teams default)"
                    >
                      {hideUserMenuFromTeams ? 'លាក់រួម (ON)' : 'បង្ហាញរួម (OFF)'}
                    </button>
                  </div>
                </th>

                {/* 5. ព័ត៏មានមន្ត្រី (officerMenu) */}
                <th className="px-3 py-2 text-center border-l border-gray-200/50 align-middle min-w-[125px]">
                  <div className="flex flex-col items-center gap-1">
                    <span className="font-bold text-[11px] leading-tight text-[#0B2545]">ព័ត៏មានមន្ត្រី</span>
                    <button
                      type="button"
                      onClick={() => onToggleHideOfficerMenuFromTeams?.(!hideOfficerMenuFromTeams)}
                      className={`px-2 py-0.5 rounded-[3px] text-[9px] font-bold transition cursor-pointer shadow-2xs border ${
                        hideOfficerMenuFromTeams
                          ? 'bg-rose-600 border-rose-700 text-white hover:bg-rose-700'
                          : 'bg-emerald-600 border-emerald-700 text-white hover:bg-emerald-700'
                      }`}
                      title="ចុចកំណត់លាក់/បង្ហាញរួមសម្រាប់គ្រប់ក្រុមទាំងអស់ (All Teams default)"
                    >
                      {hideOfficerMenuFromTeams ? 'លាក់រួម (ON)' : 'បង្ហាញរួម (OFF)'}
                    </button>
                  </div>
                </th>

                {/* 6. ការងារទិដ្ឋាការ់ (visaWork) */}
                <th className="px-3 py-2 text-center border-l border-gray-200/50 align-middle min-w-[125px]">
                  <div className="flex flex-col items-center gap-1">
                    <span className="font-bold text-[11px] leading-tight text-[#0B2545]">ការងារទិដ្ឋាការ់</span>
                    <button
                      type="button"
                      onClick={() => onToggleHideVisaWorkFromTeams?.(!hideVisaWorkFromTeams)}
                      className={`px-2 py-0.5 rounded-[3px] text-[9px] font-bold transition cursor-pointer shadow-2xs border ${
                        hideVisaWorkFromTeams
                          ? 'bg-rose-600 border-rose-700 text-white hover:bg-rose-700'
                          : 'bg-emerald-600 border-emerald-700 text-white hover:bg-emerald-700'
                      }`}
                      title="ចុចកំណត់លាក់/បង្ហាញរួមសម្រាប់គ្រប់ក្រុមទាំងអស់ (All Teams default)"
                    >
                      {hideVisaWorkFromTeams ? 'លាក់រួម (ON)' : 'បង្ហាញរួម (OFF)'}
                    </button>
                  </div>
                </th>

                {/* 7. ការងារស្តុកការិយាល័យ (officeStock) */}
                <th className="px-3 py-2 text-center border-l border-gray-200/50 align-middle min-w-[125px]">
                  <div className="flex flex-col items-center gap-1">
                    <span className="font-bold text-[11px] leading-tight text-[#0B2545]">ការងារស្តុកការិយាល័យ</span>
                    <button
                      type="button"
                      onClick={() => onToggleHideOfficeStockFromTeams?.(!hideOfficeStockFromTeams)}
                      className={`px-2 py-0.5 rounded-[3px] text-[9px] font-bold transition cursor-pointer shadow-2xs border ${
                        hideOfficeStockFromTeams
                          ? 'bg-rose-600 border-rose-700 text-white hover:bg-rose-700'
                          : 'bg-emerald-600 border-emerald-700 text-white hover:bg-emerald-700'
                      }`}
                      title="ចុចកំណត់លាក់/បង្ហាញរួមសម្រាប់គ្រប់ក្រុមទាំងអស់ (All Teams default)"
                    >
                      {hideOfficeStockFromTeams ? 'លាក់រួម (ON)' : 'បង្ហាញរួម (OFF)'}
                    </button>
                  </div>
                </th>

                {/* 8. បញ្ចូលប្រភេទ (categories) */}
                <th className="px-3 py-2 text-center border-l border-gray-200/50 align-middle min-w-[125px]">
                  <div className="flex flex-col items-center gap-1">
                    <span className="font-bold text-[11px] leading-tight text-[#0B2545]">បញ្ចូលប្រភេទ</span>
                    <button
                      type="button"
                      onClick={() => onToggleHideCategoriesFromTeams?.(!hideCategoriesFromTeams)}
                      className={`px-2 py-0.5 rounded-[3px] text-[9px] font-bold transition cursor-pointer shadow-2xs border ${
                        hideCategoriesFromTeams
                          ? 'bg-rose-600 border-rose-700 text-white hover:bg-rose-700'
                          : 'bg-emerald-600 border-emerald-700 text-white hover:bg-emerald-700'
                      }`}
                      title="ចុចកំណត់លាក់/បង្ហាញរួមសម្រាប់គ្រប់ក្រុមទាំងអស់ (All Teams default)"
                    >
                      {hideCategoriesFromTeams ? 'លាក់រួម (ON)' : 'បង្ហាញរួម (OFF)'}
                    </button>
                  </div>
                </th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {filteredUsers.length === 0 ? (
                <tr>
                  <td colSpan={14} className="px-4 py-8 text-center text-gray-500 italic">
                    មិនមានទិន្នន័យគណនីស្របតាមលក្ខខណ្ឌចម្រោះឡើយ
                  </td>
                </tr>
              ) : (
                filteredUsers.map((u) => {
                  const isVisibleAccount = u.isVisible !== false && !u.hidden;

                  // Per-user menu permissions
                  const isUserListVisible = u.permissions?.userList !== undefined ? u.permissions.userList : !hideUserMenuFromTeams;
                  const isDailyOpsVisible = u.permissions?.dailyOps !== undefined ? u.permissions.dailyOps : !hideDailyOpsFromTeams;
                  const isRefusalVisible = u.permissions?.refusal !== undefined ? u.permissions.refusal : !hideRefusalFromTeams;
                  const isUserMenuVisible = u.permissions?.userMenu !== undefined ? u.permissions.userMenu : !hideUserMenuFromTeams;
                  const isOfficerMenuVisible = u.permissions?.officerMenu !== undefined ? u.permissions.officerMenu : !hideOfficerMenuFromTeams;
                  const isVisaWorkVisible = u.permissions?.visaWork !== undefined ? u.permissions.visaWork : !hideVisaWorkFromTeams;
                  const isOfficeStockVisible = u.permissions?.officeStock !== undefined ? u.permissions.officeStock : !hideOfficeStockFromTeams;
                  const isCategoriesVisible = u.permissions?.categories !== undefined ? u.permissions.categories : !hideCategoriesFromTeams;

                  return (
                    <tr
                      key={u.id}
                      className={`hover:bg-amber-50/40 transition align-middle ${
                        !isVisibleAccount ? 'bg-gray-50/60 opacity-80' : ''
                      }`}
                    >
                      {/* 1. Username with Direct Show/Hide Toggle */}
                      <td className="px-4 py-3 align-middle">
                        <div className="flex flex-col gap-1 items-start">
                          <div className="flex items-center gap-1.5 flex-wrap">
                            <span className={`font-bold ${isVisibleAccount ? 'text-[#0B2545]' : 'text-gray-500 line-through'}`}>
                              {u.username}
                            </span>
                            {u.role === 'Secondary' && (
                              <span className="text-[9px] bg-amber-100 text-amber-800 border border-amber-300 px-1.5 py-0.2 rounded font-bold">
                                Office Admin
                              </span>
                            )}
                          </div>

                          {/* Clickable Visibility Toggle Button */}
                          <button
                            type="button"
                            onClick={() => handleToggleUserVisibility(u)}
                            className={`inline-flex items-center gap-1 px-2 py-0.5 rounded-full text-[10px] font-bold border transition-all cursor-pointer shadow-2xs hover:scale-105 active:scale-95 ${
                              isVisibleAccount
                                ? 'bg-emerald-50 text-emerald-700 border-emerald-300 hover:bg-emerald-100'
                                : 'bg-rose-50 text-rose-700 border-rose-300 hover:bg-rose-100'
                            }`}
                            title="ចុចដើម្បីកំណត់ឱ្យ ឃើញ ឬ មិនឃើញ ឈ្មោះអ្នកប្រើប្រាស់នេះ (Click to Show/Hide this username)"
                          >
                            {isVisibleAccount ? (
                              <>
                                <Eye className="w-3 h-3 text-emerald-600" />
                                <span>ឃើញ (Visible)</span>
                              </>
                            ) : (
                              <>
                                <EyeOff className="w-3 h-3 text-rose-600" />
                                <span>មិនឃើញ (Hidden)</span>
                              </>
                            )}
                          </button>
                        </div>
                      </td>

                      <td className="px-4 py-3 text-gray-600 font-mono font-semibold align-middle">{u.officerNumber || u.email}</td>

                      <td className="px-4 py-3 align-middle">
                        {u.assignedTeam ? (
                          <span className="inline-block px-2 py-0.5 rounded bg-emerald-50 text-emerald-800 border border-emerald-300 font-semibold text-[11px]">
                            {u.assignedTeam}
                          </span>
                        ) : (
                          <span className="text-gray-400 italic text-[11px]">
                            {u.role === 'Secondary' ? 'គ្រប់គ្រងរួម (Office)' : 'ទូទៅ (All)'}
                          </span>
                        )}
                      </td>

                      <td className="px-4 py-3 align-middle">
                        <div className="flex items-center gap-1.5 font-mono">
                          <span className="font-semibold text-gray-800 bg-gray-100 px-2 py-0.5 rounded border border-gray-200">
                            {visiblePasswords[u.id] ? (u.password || '123456') : '••••••••'}
                          </span>
                          <button
                            type="button"
                            onClick={() => togglePasswordVisibility(u.id)}
                            className="text-gray-400 hover:text-[#0B2545] p-1 transition rounded hover:bg-gray-100 cursor-pointer"
                            title={visiblePasswords[u.id] ? 'លាក់ពាក្យសម្ងាត់' : 'បង្ហាញពាក្យសម្ងាត់'}
                          >
                            {visiblePasswords[u.id] ? <EyeOff className="w-3.5 h-3.5" /> : <Eye className="w-3.5 h-3.5" />}
                          </button>
                        </div>
                      </td>

                      {/* Actions Column (Column 5) */}
                      <td className="px-4 py-3 text-center align-middle">
                        <div className="flex items-center justify-center gap-2">
                          <button
                            onClick={() => handleOpenEdit(u)}
                            className="text-[#0B2545] hover:bg-amber-100 p-1.5 rounded transition cursor-pointer"
                            title="កែប្រែ"
                          >
                            <Edit className="w-3.5 h-3.5" />
                          </button>
                          <button
                            onClick={() => handleDelete(u.id, u.username)}
                            className="text-rose-600 hover:bg-rose-50 p-1.5 rounded transition cursor-pointer"
                            title="លុប"
                          >
                            <Trash2 className="w-3.5 h-3.5" />
                          </button>
                        </div>
                      </td>

                      {/* Role Column (Column 6) */}
                      <td className="px-4 py-3 align-middle">
                        <span
                          className={`inline-block px-2.5 py-0.5 rounded-full text-[10px] font-bold ${
                            u.role === 'Secondary'
                              ? 'bg-amber-100 text-amber-800 border border-amber-300'
                              : u.role === 'User (ការិយាល័យ)'
                              ? 'bg-purple-100 text-purple-800 border border-purple-300'
                              : 'bg-blue-100 text-blue-800 border border-blue-300'
                          }`}
                        >
                          {u.role === 'Secondary'
                            ? 'Secondary (ការិយាល័យ)'
                            : u.role === 'User (ការិយាល័យ)'
                            ? 'User (ការិយាល័យ)'
                            : 'User (ក្រុម)'}
                        </span>
                      </td>

                      {/* 1. User List (ទិន្នន័យគណនីអ្នកប្រើប្រាស់) - Per User Toggle */}
                      <td className="px-3 py-3 text-center border-l border-gray-100 align-middle">
                        <button
                          type="button"
                          onClick={() => handleToggleUserPermission(u, 'userList', hideUserMenuFromTeams, 'ទិន្នន័យគណនីអ្នកប្រើប្រាស់')}
                          className={`inline-flex items-center gap-1 px-2 py-0.5 rounded font-bold text-[10px] border transition-all cursor-pointer hover:scale-105 active:scale-95 shadow-2xs ${
                            isUserListVisible
                              ? 'bg-blue-50 text-blue-700 border-blue-200 hover:bg-blue-100'
                              : 'bg-rose-50 text-rose-700 border-rose-200 hover:bg-rose-100'
                          }`}
                          title={`ចុចដើម្បីកំណត់ឱ្យ ${u.username} មើលឃើញ ឬ មិនឃើញ មឺនុយនេះ`}
                        >
                          <span>{isUserListVisible ? '👁️ បង្ហាញ' : '🚫 លាក់'}</span>
                          {u.permissions?.userList !== undefined && (
                            <span className="text-[8px] bg-blue-200 text-blue-900 px-1 rounded-full font-extrabold">*ផ្ទាល់</span>
                          )}
                        </button>
                      </td>

                      {/* 2. Daily Operations (ប្រតិបត្តិការប្រចាំថ្ងៃតាមក្រុម) - Per User Toggle */}
                      <td className="px-3 py-3 text-center border-l border-gray-100 align-middle">
                        <button
                          type="button"
                          onClick={() => handleToggleUserPermission(u, 'dailyOps', hideDailyOpsFromTeams, 'ប្រតិបត្តិការប្រចាំថ្ងៃតាមក្រុម')}
                          className={`inline-flex items-center gap-1 px-2 py-0.5 rounded font-bold text-[10px] border transition-all cursor-pointer hover:scale-105 active:scale-95 shadow-2xs ${
                            isDailyOpsVisible
                              ? 'bg-blue-50 text-blue-700 border-blue-200 hover:bg-blue-100'
                              : 'bg-rose-50 text-rose-700 border-rose-200 hover:bg-rose-100'
                          }`}
                          title={`ចុចដើម្បីកំណត់ឱ្យ ${u.username} មើលឃើញ ឬ មិនឃើញ មឺនុយនេះ`}
                        >
                          <span>{isDailyOpsVisible ? '👁️ បង្ហាញ' : '🚫 លាក់'}</span>
                          {u.permissions?.dailyOps !== undefined && (
                            <span className="text-[8px] bg-blue-200 text-blue-900 px-1 rounded-full font-extrabold">*ផ្ទាល់</span>
                          )}
                        </button>
                      </td>

                      {/* 3. Refusal (បដិសេធនិងបញ្ជូនចេញ) - Per User Toggle */}
                      <td className="px-3 py-3 text-center border-l border-gray-100 align-middle">
                        <button
                          type="button"
                          onClick={() => handleToggleUserPermission(u, 'refusal', hideRefusalFromTeams, 'បដិសេធនិងបញ្ជូនចេញ')}
                          className={`inline-flex items-center gap-1 px-2 py-0.5 rounded font-bold text-[10px] border transition-all cursor-pointer hover:scale-105 active:scale-95 shadow-2xs ${
                            isRefusalVisible
                              ? 'bg-blue-50 text-blue-700 border-blue-200 hover:bg-blue-100'
                              : 'bg-rose-50 text-rose-700 border-rose-200 hover:bg-rose-100'
                          }`}
                          title={`ចុចដើម្បីកំណត់ឱ្យ ${u.username} មើលឃើញ ឬ មិនឃើញ មឺនុយនេះ`}
                        >
                          <span>{isRefusalVisible ? '👁️ បង្ហាញ' : '🚫 លាក់'}</span>
                          {u.permissions?.refusal !== undefined && (
                            <span className="text-[8px] bg-blue-200 text-blue-900 px-1 rounded-full font-extrabold">*ផ្ទាល់</span>
                          )}
                        </button>
                      </td>

                      {/* 4. User Account (គណីអ្នកប្រើប្រាស់) - Per User Toggle */}
                      <td className="px-3 py-3 text-center border-l border-gray-100 align-middle">
                        <button
                          type="button"
                          onClick={() => handleToggleUserPermission(u, 'userMenu', hideUserMenuFromTeams, 'គណីអ្នកប្រើប្រាស់')}
                          className={`inline-flex items-center gap-1 px-2 py-0.5 rounded font-bold text-[10px] border transition-all cursor-pointer hover:scale-105 active:scale-95 shadow-2xs ${
                            isUserMenuVisible
                              ? 'bg-blue-50 text-blue-700 border-blue-200 hover:bg-blue-100'
                              : 'bg-rose-50 text-rose-700 border-rose-200 hover:bg-rose-100'
                          }`}
                          title={`ចុចដើម្បីកំណត់ឱ្យ ${u.username} មើលឃើញ ឬ មិនឃើញ មឺនុយនេះ`}
                        >
                          <span>{isUserMenuVisible ? '👁️ បង្ហាញ' : '🚫 លាក់'}</span>
                          {u.permissions?.userMenu !== undefined && (
                            <span className="text-[8px] bg-blue-200 text-blue-900 px-1 rounded-full font-extrabold">*ផ្ទាល់</span>
                          )}
                        </button>
                      </td>

                      {/* 5. Officer Info (ព័ត៏មានមន្ត្រី) - Per User Toggle */}
                      <td className="px-3 py-3 text-center border-l border-gray-100 align-middle">
                        <button
                          type="button"
                          onClick={() => handleToggleUserPermission(u, 'officerMenu', hideOfficerMenuFromTeams, 'ព័ត៏មានមន្ត្រី')}
                          className={`inline-flex items-center gap-1 px-2 py-0.5 rounded font-bold text-[10px] border transition-all cursor-pointer hover:scale-105 active:scale-95 shadow-2xs ${
                            isOfficerMenuVisible
                              ? 'bg-blue-50 text-blue-700 border-blue-200 hover:bg-blue-100'
                              : 'bg-rose-50 text-rose-700 border-rose-200 hover:bg-rose-100'
                          }`}
                          title={`ចុចដើម្បីកំណត់ឱ្យ ${u.username} មើលឃើញ ឬ មិនឃើញ មឺនុយនេះ`}
                        >
                          <span>{isOfficerMenuVisible ? '👁️ បង្ហាញ' : '🚫 លាក់'}</span>
                          {u.permissions?.officerMenu !== undefined && (
                            <span className="text-[8px] bg-blue-200 text-blue-900 px-1 rounded-full font-extrabold">*ផ្ទាល់</span>
                          )}
                        </button>
                      </td>

                      {/* 6. Visa Work (ការងារទិដ្ឋាការ់) - Per User Toggle */}
                      <td className="px-3 py-3 text-center border-l border-gray-100 align-middle">
                        <button
                          type="button"
                          onClick={() => handleToggleUserPermission(u, 'visaWork', hideVisaWorkFromTeams, 'ការងារទិដ្ឋាការ់')}
                          className={`inline-flex items-center gap-1 px-2 py-0.5 rounded font-bold text-[10px] border transition-all cursor-pointer hover:scale-105 active:scale-95 shadow-2xs ${
                            isVisaWorkVisible
                              ? 'bg-blue-50 text-blue-700 border-blue-200 hover:bg-blue-100'
                              : 'bg-rose-50 text-rose-700 border-rose-200 hover:bg-rose-100'
                          }`}
                          title={`ចុចដើម្បីកំណត់ឱ្យ ${u.username} មើលឃើញ ឬ មិនឃើញ មឺនុយនេះ`}
                        >
                          <span>{isVisaWorkVisible ? '👁️ បង្ហាញ' : '🚫 លាក់'}</span>
                          {u.permissions?.visaWork !== undefined && (
                            <span className="text-[8px] bg-blue-200 text-blue-900 px-1 rounded-full font-extrabold">*ផ្ទាល់</span>
                          )}
                        </button>
                      </td>

                      {/* 7. Office Stock (ការងារស្តុកការិយាល័យ) - Per User Toggle */}
                      <td className="px-3 py-3 text-center border-l border-gray-100 align-middle">
                        <button
                          type="button"
                          onClick={() => handleToggleUserPermission(u, 'officeStock', hideOfficeStockFromTeams, 'ការងារស្តុកការិយាល័យ')}
                          className={`inline-flex items-center gap-1 px-2 py-0.5 rounded font-bold text-[10px] border transition-all cursor-pointer hover:scale-105 active:scale-95 shadow-2xs ${
                            isOfficeStockVisible
                              ? 'bg-blue-50 text-blue-700 border-blue-200 hover:bg-blue-100'
                              : 'bg-rose-50 text-rose-700 border-rose-200 hover:bg-rose-100'
                          }`}
                          title={`ចុចដើម្បីកំណត់ឱ្យ ${u.username} មើលឃើញ ឬ មិនឃើញ មឺនុយនេះ`}
                        >
                          <span>{isOfficeStockVisible ? '👁️ បង្ហាញ' : '🚫 លាក់'}</span>
                          {u.permissions?.officeStock !== undefined && (
                            <span className="text-[8px] bg-blue-200 text-blue-900 px-1 rounded-full font-extrabold">*ផ្ទាល់</span>
                          )}
                        </button>
                      </td>

                      {/* 8. Category Manager (បញ្ចូលប្រភេទ) - Per User Toggle */}
                      <td className="px-3 py-3 text-center border-l border-gray-100 align-middle">
                        <button
                          type="button"
                          onClick={() => handleToggleUserPermission(u, 'categories', hideCategoriesFromTeams, 'បញ្ចូលប្រភេទ')}
                          className={`inline-flex items-center gap-1 px-2 py-0.5 rounded font-bold text-[10px] border transition-all cursor-pointer hover:scale-105 active:scale-95 shadow-2xs ${
                            isCategoriesVisible
                              ? 'bg-blue-50 text-blue-700 border-blue-200 hover:bg-blue-100'
                              : 'bg-rose-50 text-rose-700 border-rose-200 hover:bg-rose-100'
                          }`}
                          title={`ចុចដើម្បីកំណត់ឱ្យ ${u.username} មើលឃើញ ឬ មិនឃើញ មឺនុយនេះ`}
                        >
                          <span>{isCategoriesVisible ? '👁️ បង្ហាញ' : '🚫 លាក់'}</span>
                          {u.permissions?.categories !== undefined && (
                            <span className="text-[8px] bg-blue-200 text-blue-900 px-1 rounded-full font-extrabold">*ផ្ទាល់</span>
                          )}
                        </button>
                      </td>
                    </tr>
                  );
                })
              )}
            </tbody>
          </table>
        </div>
      </div>

      {/* EDIT USER MODAL */}
      {editingUser && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-2xs flex items-center justify-center p-4">
          <div className="bg-white rounded-md border border-gray-300 border-t-4 border-t-[#007bff] shadow-xl w-full max-w-sm p-5 animate-fade">
            <div className="flex items-center justify-between border-b border-gray-200 pb-3 mb-4">
              <h3 className="font-bold text-gray-800 text-sm">កែប្រែគណនី (Edit User)</h3>
              <button
                onClick={() => setEditingUser(null)}
                className="text-gray-400 hover:text-gray-600 p-1"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <form onSubmit={handleSaveEdit} className="space-y-3.5 text-xs">
              <div>
                <label className="block font-bold text-gray-700 mb-1">
                  ឈ្មោះអ្នកប្រើប្រាស់
                </label>
                <input
                  type="text"
                  required
                  value={editUsername}
                  onChange={(e) => setEditUsername(e.target.value)}
                  className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                />
              </div>

              <div>
                <label className="block font-bold text-gray-700 mb-1">
                  អត្តលេខ
                </label>
                <input
                  type="text"
                  required
                  value={editOfficerNumber}
                  onChange={(e) => setEditOfficerNumber(e.target.value)}
                  className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                />
              </div>

              <div>
                <label className="block font-bold text-gray-700 mb-1">
                  លេខសម្ងាត់ (Password)
                </label>
                <div className="relative">
                  <input
                    type={showEditPassword ? 'text' : 'password'}
                    required
                    minLength={6}
                    value={editPassword}
                    onChange={(e) => setEditPassword(e.target.value)}
                    className="w-full border border-gray-300 rounded-sm pl-2.5 pr-8 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                  />
                  <button
                    type="button"
                    onClick={() => setShowEditPassword(!showEditPassword)}
                    className="absolute right-2.5 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-700 p-1"
                  >
                    {showEditPassword ? <EyeOff className="w-3.5 h-3.5" /> : <Eye className="w-3.5 h-3.5" />}
                  </button>
                </div>
              </div>

              <div>
                <label className="block font-bold text-gray-700 mb-1">តួនាទី</label>
                <select
                  value={editRole}
                  onChange={(e) => setEditRole(e.target.value as UserRole)}
                  className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                >
                  <option value="User">User (ក្រុម)</option>
                  <option value="User (ការិយាល័យ)">User (ការិយាល័យ)</option>
                  <option value="Secondary">Secondary (ការិយាល័យ)</option>
                </select>
              </div>

              <div>
                <label className="block font-bold text-gray-700 mb-1">
                  ការបង្ហាញឈ្មោះអ្នកប្រើប្រាស់ (Account Visibility)
                </label>
                <div className="flex items-center gap-4 bg-gray-50 p-2 rounded border border-gray-200">
                  <label className="flex items-center gap-1.5 cursor-pointer font-bold text-xs text-emerald-800">
                    <input
                      type="radio"
                      name="editIsVisible"
                      checked={editIsVisible}
                      onChange={() => setEditIsVisible(true)}
                      className="text-emerald-600 focus:ring-emerald-500"
                    />
                    <span>👁️ ឃើញ / បង្ហាញ (Visible)</span>
                  </label>
                  <label className="flex items-center gap-1.5 cursor-pointer font-bold text-xs text-rose-800">
                    <input
                      type="radio"
                      name="editIsVisible"
                      checked={!editIsVisible}
                      onChange={() => setEditIsVisible(false)}
                      className="text-rose-600 focus:ring-rose-500"
                    />
                    <span>🚫 មិនឃើញ / លាក់ (Hidden)</span>
                  </label>
                </div>
              </div>

              {editRole === 'User' && (
                <div className="bg-amber-50/70 p-2.5 rounded border border-amber-200">
                  <label className="block font-bold text-amber-900 mb-1 flex items-center gap-1.5">
                    <Building className="w-3.5 h-3.5 text-amber-700" />
                    <span>ក្រុមផ្ទាល់ខ្លួន (Assigned Team)</span>
                  </label>
                  <select
                    value={editAssignedTeam}
                    onChange={(e) => setEditAssignedTeam(e.target.value)}
                    className="w-full border border-amber-300 rounded px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:outline-none bg-white transition"
                  >
                    <option value="">-- មើលគ្រប់ក្រុម (All Teams) --</option>
                    {teamOptions.map((t) => (
                      <option key={t} value={t}>
                        {t}
                      </option>
                    ))}
                  </select>
                </div>
              )}

              <div className="flex justify-end gap-2 pt-3 border-t border-gray-200">
                <button
                  type="button"
                  onClick={() => setEditingUser(null)}
                  className="min-w-[120px] bg-[#6c757d] hover:bg-[#5a6268] text-white px-4 py-1.5 rounded text-xs font-bold transition flex items-center justify-center cursor-pointer"
                >
                  <span>បោះបង់</span>
                </button>
                <button
                  type="submit"
                  className="min-w-[120px] bg-[#28a745] hover:bg-[#218838] text-white font-bold text-xs px-4 py-1.5 rounded flex items-center justify-center gap-1.5 transition cursor-pointer shadow-2xs"
                >
                  <Save className="w-3.5 h-3.5" />
                  <span>រក្សាទុក</span>
                </button>
              </div>
            </form>
          </div>
        </div>
      )}

      {/* DELETE USER CONFIRMATION MODAL */}
      {deletingUser && (
        <div className="fixed inset-0 z-50 bg-[#071A33]/60 backdrop-blur-xs flex items-center justify-center p-4">
          <div className="bg-white rounded-2xl border border-rose-200 shadow-2xl w-full max-w-sm p-6 animate-fade">
            <div className="flex items-center gap-3 text-rose-600 mb-3">
              <div className="w-10 h-10 rounded-full bg-rose-100 flex items-center justify-center shrink-0">
                <Trash2 className="w-5 h-5" />
              </div>
              <h3 className="font-display text-[#0B2545] text-base font-bold">បញ្ជាក់ការលុបគណនី</h3>
            </div>
            <p className="text-xs text-gray-600 mb-5 leading-relaxed">
              តើអ្នកពិតជាចង់លុបគណនី <span className="font-bold text-[#0B2545]">"{deletingUser.name}"</span> នេះមែនទេ?
            </p>
            <div className="flex items-center justify-end gap-2.5">
              <button
                type="button"
                onClick={() => setDeletingUser(null)}
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

