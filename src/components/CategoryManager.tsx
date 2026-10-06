import React, { useState } from 'react';
import * as XLSX from 'xlsx';
import { CategoriesState, CategoryType, UserRole } from '../types';
import { toKhmerNum } from '../utils/khmerCalendar';
import {
  FolderPlus,
  Award,
  Briefcase,
  Layers,
  Building2,
  UserCheck,
  Plus,
  Trash2,
  Edit2,
  Save,
  X,
  ArrowRight,
  CheckCircle,
  Home,
  Table,
  Download,
  Upload,
  FileSpreadsheet,
} from 'lucide-react';
import { formatReportTeamName } from '../utils/teamNormalization';

interface CategoryManagerProps {
  categories: CategoriesState;
  activeCatType: CategoryType;
  currentRole: UserRole;
  onAddCategory: (type: CategoryType, name: string) => void;
  onUpdateCategory: (type: CategoryType, id: string, name: string) => void;
  onDeleteCategory: (type: CategoryType, id: string) => void;
  onNavigateToOfficerForm: () => void;
  onSelectCatType: (type: CategoryType) => void;
  onShowToast: (msg: string, type: 'success' | 'error') => void;
}

export const CategoryManager: React.FC<CategoryManagerProps> = ({
  categories,
  activeCatType,
  currentRole,
  onAddCategory,
  onUpdateCategory,
  onDeleteCategory,
  onNavigateToOfficerForm,
  onSelectCatType,
  onShowToast,
}) => {
  const [newItemName, setNewItemName] = useState('');
  const [recentlyAdded, setRecentlyAdded] = useState<string | null>(null);

  // Edit state for single category
  const [editingId, setEditingId] = useState<string | null>(null);
  const [editingName, setEditingName] = useState('');
  const [deletingItem, setDeletingItem] = useState<{ id: string; name: string } | null>(null);

  // State for combined visa teams pair
  const [newTeamFullName, setNewTeamFullName] = useState('');
  const [newTeamRobokName, setNewTeamRobokName] = useState('');
  const [editingPairIdx, setEditingPairIdx] = useState<number | null>(null);
  const [editingPairFull, setEditingPairFull] = useState('');
  const [editingPairRobok, setEditingPairRobok] = useState('');
  const [deletingPair, setDeletingPair] = useState<{ idx: number; full: string; robok: string } | null>(null);

  // State for team option selections (សន្លឹកទិដ្ឋាការ, ក្រដាសអនុម័ត)
  const [showImportModal, setShowImportModal] = useState(false);
  const [teamTypeOptions, setTeamTypeOptions] = useState<Record<string | number, { sticker: boolean; evisa: boolean }>>(() => {
    try {
      const saved = localStorage.getItem('team_type_options_v1');
      return saved ? JSON.parse(saved) : {};
    } catch {
      return {};
    }
  });

  // Export Excel Function for Visa Teams
  const handleExportExcel = () => {
    const maxLen = Math.max((categories.visaTeams || []).length, (categories.visaTeamsRobok || []).length);
    if (maxLen === 0) {
      onShowToast('គ្មានទិន្នន័យដើម្បីទាញយក!', 'error');
      return;
    }

    const rows = [];
    for (let i = 0; i < maxLen; i++) {
      const vtItem = categories.visaTeams?.[i];
      const vtrItem = categories.visaTeamsRobok?.[i];
      const full = vtItem?.name || '';
      const robok = vtrItem?.name || '';

      const currentOpt =
        (vtrItem?.id && teamTypeOptions[vtrItem.id]) ||
        (vtrItem?.name && teamTypeOptions[vtrItem.name]) ||
        (vtItem?.id && teamTypeOptions[vtItem.id]) ||
        (vtItem?.name && teamTypeOptions[vtItem.name]) ||
        teamTypeOptions[i] ||
        { sticker: true, evisa: true };

      let usageStr = '';
      if (currentOpt.sticker && currentOpt.evisa) usageStr = 'សន្លឹកទិដ្ឋាការ + ក្រដាសអនុម័ត';
      else if (currentOpt.sticker) usageStr = 'សន្លឹកទិដ្ឋាការ';
      else if (currentOpt.evisa) usageStr = 'ក្រដាសអនុម័ត';
      else usageStr = 'គ្មាន';

      rows.push({
        'ល.រ': i + 1,
        'ក្រុមផ្តល់ទិដ្ឋាការ (ឈ្មោះពេញ)': full,
        'ក្រុមផ្តល់ទិដ្ឋាការ.របក (ឈ្មោះខ្លី)': robok,
        'ប្រភេទប្រើប្រាស់': usageStr,
      });
    }

    const worksheet = XLSX.utils.json_to_sheet(rows);
    worksheet['!cols'] = [
      { wch: 8 },
      { wch: 55 },
      { wch: 30 },
      { wch: 35 },
    ];

    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'ទិន្នន័យក្រុមផ្តល់ទិដ្ឋាការ');
    const dateStr = new Date().toISOString().split('T')[0];
    XLSX.writeFile(workbook, `Visa_Teams_Data_${dateStr}.xlsx`);
    onShowToast('បានទាញយកឯកសារ Excel រួចរាល់!', 'success');
  };

  // Download Sample Import Template
  const handleDownloadImportTemplate = () => {
    const sampleRows = [
      {
        'ល.រ': 1,
        'ក្រុមផ្តល់ទិដ្ឋាការ (ឈ្មោះពេញ)': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអាកាសយានដ្ឋានអន្តរជាតិតេជោ',
        'ក្រុមផ្តល់ទិដ្ឋាការ.របក (ឈ្មោះខ្លី)': 'អាកាស តេជោ',
        'ប្រភេទប្រើប្រាស់': 'សន្លឹកទិដ្ឋាការ + ក្រដាសអនុម័ត',
      },
      {
        'ល.រ': 2,
        'ក្រុមផ្តល់ទិដ្ឋាការ (ឈ្មោះពេញ)': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអាកាសយានដ្ឋានអន្តរជាតិភ្នំពេញ',
        'ក្រុមផ្តល់ទិដ្ឋាការ.របក (ឈ្មោះខ្លី)': 'អាកាស ភ្នំពេញ',
        'ប្រភេទប្រើប្រាស់': 'សន្លឹកទិដ្ឋាការ + ក្រដាសអនុម័ត',
      },
      {
        'ល.រ': 3,
        'ក្រុមផ្តល់ទិដ្ឋាការ (ឈ្មោះពេញ)': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអាកាសយានដ្ឋានអន្តរជាតិសៀមរាប',
        'ក្រុមផ្តល់ទិដ្ឋាការ.របក (ឈ្មោះខ្លី)': 'អាកាស សៀមរាប',
        'ប្រភេទប្រើប្រាស់': 'សន្លឹកទិដ្ឋាការ + ក្រដាសអនុម័ត',
      },
    ];

    const worksheet = XLSX.utils.json_to_sheet(sampleRows);
    worksheet['!cols'] = [
      { wch: 8 },
      { wch: 55 },
      { wch: 30 },
      { wch: 35 },
    ];
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'ទម្រង់គំរូក្រុម');
    XLSX.writeFile(workbook, 'Visa_Teams_Import_Template.xlsx');
    onShowToast('បានទាញយកទម្រង់គំរូ Excel រួចរាល់!', 'success');
  };

  // File Import Handler
  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const buffer = evt.target?.result;
        if (!buffer) return;
        const workbook = XLSX.read(buffer, { type: 'binary', cellDates: true });
        const firstSheetName = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheetName];
        const rawRows = XLSX.utils.sheet_to_json<Record<string, any>>(sheet, { defval: '' });

        if (rawRows.length === 0) {
          onShowToast('ឯកសារ Excel គ្មានទិន្នន័យ!', 'error');
          return;
        }

        let importedCount = 0;
        const newTypeOptions = { ...teamTypeOptions };

        rawRows.forEach((row, idx) => {
          let fullName = '';
          let robokName = '';
          let usageStr = '';

          Object.keys(row).forEach((key) => {
            const lowerKey = key.toLowerCase();
            const val = String(row[key] || '').trim();

            if (lowerKey.includes('របក') || lowerKey.includes('ឈ្មោះខ្លី') || lowerKey.includes('short')) {
              if (!robokName) robokName = val;
            } else if (lowerKey.includes('ក្រុម') || lowerKey.includes('ឈ្មោះពេញ') || lowerKey.includes('full')) {
              if (!fullName) fullName = val;
            } else if (lowerKey.includes('ប្រភេទ') || lowerKey.includes('ប្រើប្រាស់') || lowerKey.includes('usage') || lowerKey.includes('type')) {
              if (!usageStr) usageStr = val;
            }
          });

          if (!fullName && row['ក្រុមផ្តល់ទិដ្ឋាការ (ឈ្មោះពេញ)']) fullName = String(row['ក្រុមផ្តល់ទិដ្ឋាការ (ឈ្មោះពេញ)']).trim();
          if (!fullName && row['ក្រុមផ្តល់ទិដ្ឋាការ']) fullName = String(row['ក្រុមផ្តល់ទិដ្ឋាការ']).trim();
          if (!robokName && row['ក្រុមផ្តល់ទិដ្ឋាការ.របក (ឈ្មោះខ្លី)']) robokName = String(row['ក្រុមផ្តល់ទិដ្ឋាការ.របក (ឈ្មោះខ្លី)']).trim();
          if (!robokName && row['ក្រុមផ្តល់ទិដ្ឋាការ.របក']) robokName = String(row['ក្រុមផ្តល់ទិដ្ឋាការ.របក']).trim();
          if (!usageStr && row['ប្រភេទប្រើប្រាស់']) usageStr = String(row['ប្រភេទប្រើប្រាស់']).trim();

          if (fullName || robokName) {
            if (fullName) {
              const existsFull = (categories.visaTeams || []).some(
                (item) => item.name.trim().toLowerCase() === fullName.toLowerCase()
              );
              if (!existsFull) {
                onAddCategory('visaTeams', fullName);
              }
            }

            if (robokName) {
              const existsRobok = (categories.visaTeamsRobok || []).some(
                (item) => item.name.trim().toLowerCase() === robokName.toLowerCase()
              );
              if (!existsRobok) {
                onAddCategory('visaTeamsRobok', robokName);
              }
            }

            if (usageStr) {
              let sticker = true;
              let evisa = true;
              if (usageStr.includes('សន្លឹក') && !usageStr.includes('ក្រដាស') && !usageStr.includes('អនុម័ត')) {
                evisa = false;
              } else if ((usageStr.includes('ក្រដាស') || usageStr.includes('អនុម័ត') || usageStr.toLowerCase().includes('evisa')) && !usageStr.includes('សន្លឹក') && !usageStr.toLowerCase().includes('sticker')) {
                sticker = false;
              }

              const keyVal = robokName || fullName || idx;
              newTypeOptions[keyVal] = { sticker, evisa };
            }

            importedCount++;
          }
        });

        setTeamTypeOptions(newTypeOptions);
        try {
          localStorage.setItem('team_type_options_v1', JSON.stringify(newTypeOptions));
          window.dispatchEvent(new Event('team_type_options_updated'));
        } catch (err) {
          console.error('Failed to save team type options', err);
        }

        onShowToast(`បាននាំចូលទិន្នន័យក្រុមផ្តល់ទិដ្ឋាការចំនួន ${toKhmerNum(importedCount)} រួចរាល់!`, 'success');
        setShowImportModal(false);
      } catch (err) {
        console.error('Import excel error:', err);
        onShowToast('មានបញ្ហាក្នុងការអាន File Excel! សូមពិនិត្យមើលទម្រង់ File!', 'error');
      }
    };
    reader.readAsBinaryString(file);
    e.target.value = '';
  };

  const handleToggleTypeOption = (idx: number, type: 'sticker' | 'evisa') => {
    const vtrItem = (categories.visaTeamsRobok || [])[idx];
    const vtItem = (categories.visaTeams || [])[idx];

    const current =
      (vtrItem?.id && teamTypeOptions[vtrItem.id]) ||
      (vtrItem?.name && teamTypeOptions[vtrItem.name]) ||
      (vtItem?.id && teamTypeOptions[vtItem.id]) ||
      (vtItem?.name && teamTypeOptions[vtItem.name]) ||
      teamTypeOptions[idx] ||
      { sticker: true, evisa: true };

    const updatedVal = {
      ...current,
      [type]: !current[type],
    };

    const updated = {
      ...teamTypeOptions,
      [idx]: updatedVal,
    };

    if (vtrItem?.id) updated[vtrItem.id] = updatedVal;
    if (vtrItem?.name) updated[vtrItem.name] = updatedVal;
    if (vtItem?.id) updated[vtItem.id] = updatedVal;
    if (vtItem?.name) updated[vtItem.name] = updatedVal;

    setTeamTypeOptions(updated);
    try {
      localStorage.setItem('team_type_options_v1', JSON.stringify(updated));
      window.dispatchEvent(new Event('team_type_options_updated'));
    } catch (e) {
      console.error('Failed to save team type options', e);
    }
  };

  const handleSetAllTypeOption = (type: 'all' | 'sticker' | 'evisa') => {
    const maxLen = Math.max((categories.visaTeams || []).length, (categories.visaTeamsRobok || []).length);
    const updated: Record<string | number, { sticker: boolean; evisa: boolean }> = {};
    for (let i = 0; i < maxLen; i++) {
      let val = { sticker: true, evisa: true };
      if (type === 'all') val = { sticker: true, evisa: true };
      else if (type === 'sticker') val = { sticker: true, evisa: false };
      else if (type === 'evisa') val = { sticker: false, evisa: true };

      updated[i] = val;
      const vtrItem = (categories.visaTeamsRobok || [])[i];
      const vtItem = (categories.visaTeams || [])[i];
      if (vtrItem?.id) updated[vtrItem.id] = val;
      if (vtrItem?.name) updated[vtrItem.name] = val;
      if (vtItem?.id) updated[vtItem.id] = val;
      if (vtItem?.name) updated[vtItem.name] = val;
    }
    setTeamTypeOptions(updated);
    try {
      localStorage.setItem('team_type_options_v1', JSON.stringify(updated));
      window.dispatchEvent(new Event('team_type_options_updated'));
    } catch (e) {
      console.error('Failed to save team type options', e);
    }
  };

  const catTypeLabels: Record<CategoryType, string> = {
    ranks: 'ឋានន្តរស័ក្កិ',
    positions: 'តួនាទី',
    teamPositions: 'តួនាទីតាមក្រុម',
    visaTeams: 'ក្រុមផ្តល់ទិដ្ឋាការ',
    visaTeamsRobok: 'ក្រុមផ្តល់ទិដ្ឋាការ.របក',
    visaTeamsData: 'ទិន្នន័យក្រុមផ្តល់ទិដ្ឋាការ',
    collectorRoles: 'តួនាទីអ្នកមកបើក',
    organizations: 'ស្ថាប័នដើម',
    refusalReasons: 'ករណីបដិសេធ',
  };

  const handleAdd = (e: React.FormEvent) => {
    e.preventDefault();
    const name = newItemName.trim();

    if (!name) {
      onShowToast('សូមបញ្ចូលឈ្មោះប្រភេទជាមុនសិន', 'error');
      return;
    }

    // Check duplicate
    const currentList = categories[activeCatType] || [];
    const exists = currentList.some(
      (item) => item.name.trim().toLowerCase() === name.toLowerCase()
    );

    if (exists) {
      onShowToast(`ឈ្មោះ "${name}" មានរួចហើយនៅក្នុង ${catTypeLabels[activeCatType]}`, 'error');
      return;
    }

    onAddCategory(activeCatType, name);
    setNewItemName('');
    setRecentlyAdded(name);
    onShowToast(`បន្ថែម "${name}" ទៅក្នុង ${catTypeLabels[activeCatType]} រួចរាល់!`, 'success');
  };

  const handleAddTeamPair = (e: React.FormEvent) => {
    e.preventDefault();
    let full = newTeamFullName.trim();
    const robok = newTeamRobokName.trim();

    if (!full && !robok) {
      onShowToast('សូមបញ្ចូលឈ្មោះក្រុមផ្តល់ទិដ្ឋាការ ឬឈ្មោះក្រុម.របក', 'error');
      return;
    }

    if (!full && robok) {
      full = formatReportTeamName(robok);
    }

    if (full) {
      onAddCategory('visaTeams', full);
    }
    if (robok) {
      onAddCategory('visaTeamsRobok', robok);
    }

    setNewTeamFullName('');
    setNewTeamRobokName('');
    setRecentlyAdded(full || robok);
    onShowToast('បានបន្ថែមទិន្នន័យក្រុមផ្តល់ទិដ្ឋាការរួចរាល់!', 'success');
  };

  const handleStartEdit = (id: string, name: string) => {
    setEditingId(id);
    setEditingName(name);
  };

  const handleCancelEdit = () => {
    setEditingId(null);
    setEditingName('');
  };

  const handleSaveEdit = (id: string) => {
    const trimmed = editingName.trim();
    if (!trimmed) {
      onShowToast('ឈ្មោះប្រភេទមិនអាចទទេបានឡើយ', 'error');
      return;
    }

    // Check duplicate (excluding current item)
    const currentList = categories[activeCatType] || [];
    const exists = currentList.some(
      (item) => item.id !== id && item.name.trim().toLowerCase() === trimmed.toLowerCase()
    );

    if (exists) {
      onShowToast(`ឈ្មោះ "${trimmed}" មានរួចហើយនៅក្នុង ${catTypeLabels[activeCatType]}`, 'error');
      return;
    }

    onUpdateCategory(activeCatType, id, trimmed);
    setEditingId(null);
    setEditingName('');
    onShowToast(`កែប្រែឈ្មោះ "${trimmed}" ជោគជ័យ!`, 'success');
  };

  const handleDelete = (id: string, name: string) => {
    setDeletingItem({ id, name });
  };

  const handleConfirmDelete = () => {
    if (deletingItem) {
      onDeleteCategory(activeCatType, deletingItem.id);
      onShowToast(`បានលុប "${deletingItem.name}" រួចរាល់`, 'success');
      setDeletingItem(null);
    }
  };

  const handleStartEditPair = (idx: number, full: string, robok: string) => {
    setEditingPairIdx(idx);
    let resolvedFull = full;
    if ((!resolvedFull || /^ក្រុមទី\s*\d+$/i.test(resolvedFull.trim())) && robok) {
      resolvedFull = formatReportTeamName(robok);
    }
    setEditingPairFull(resolvedFull);
    setEditingPairRobok(robok);
  };

  const handleCancelEditPair = () => {
    setEditingPairIdx(null);
    setEditingPairFull('');
    setEditingPairRobok('');
  };

  const handleSaveEditPair = (idx: number) => {
    const fullTrim = editingPairFull.trim();
    const robokTrim = editingPairRobok.trim();

    const vtItem = categories.visaTeams[idx];
    const vtrItem = categories.visaTeamsRobok[idx];

    if (vtItem) {
      if (fullTrim) {
        onUpdateCategory('visaTeams', vtItem.id, fullTrim);
      }
    } else if (fullTrim) {
      onAddCategory('visaTeams', fullTrim);
    }

    if (vtrItem) {
      if (robokTrim) {
        onUpdateCategory('visaTeamsRobok', vtrItem.id, robokTrim);
      }
    } else if (robokTrim) {
      onAddCategory('visaTeamsRobok', robokTrim);
    }

    setEditingPairIdx(null);
    onShowToast('បានកែប្រែទិន្នន័យក្រុមផ្តល់ទិដ្ឋាការជោគជ័យ!', 'success');
  };

  const handleConfirmDeletePair = () => {
    if (deletingPair) {
      const { idx } = deletingPair;
      const vtItem = categories.visaTeams[idx];
      const vtrItem = categories.visaTeamsRobok[idx];

      if (vtItem) onDeleteCategory('visaTeams', vtItem.id);
      if (vtrItem) onDeleteCategory('visaTeamsRobok', vtrItem.id);

      onShowToast('បានលុបទិន្នន័យក្រុមផ្តល់ទិដ្ឋាការរួចរាល់', 'success');
      setDeletingPair(null);
    }
  };

  const currentList = categories[activeCatType] || [];
  const maxPairLength = Math.max((categories.visaTeams || []).length, (categories.visaTeamsRobok || []).length);

  return (
    <div className="max-w-4xl mx-auto space-y-5 animate-fade">
      {/* AdminLTE Content Header (Breadcrumbs) */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 bg-white px-4 py-3 rounded-md border border-gray-200 shadow-xs">
        <div>
          <h1 className="text-base font-bold text-gray-800 flex items-center gap-2">
            <span>គ្រប់គ្រងប្រភេទផ្សេងៗ (Category Manager)</span>
          </h1>
          <p className="text-xs text-gray-500">
            បន្ថែម កែប្រែ ឬលុប ឋានន្តរស័ក្កិ តួនាទី ក្រុមផ្តល់ទិដ្ឋាការ និងស្ថាប័នដើម
          </p>
        </div>
        <div className="flex items-center gap-1.5 text-xs text-gray-500 bg-gray-100 px-3 py-1.5 rounded-md font-medium self-start sm:self-auto">
          <Home className="w-3.5 h-3.5 text-gray-500" />
          <span>ទំព័រដើម</span>
          <span>/</span>
          <span>ព័ត៌មានមន្ត្រី</span>
          <span>/</span>
          <span className="text-[#007bff] font-bold">គ្រប់គ្រងប្រភេទ</span>
        </div>
      </div>

      {/* Category Type Switcher Tabs */}
      <div className="bg-white rounded-md border border-gray-300 p-1.5 shadow-xs flex flex-wrap gap-1.5">
        <button
          onClick={() => onSelectCatType('ranks')}
          className={`flex-1 min-w-[120px] py-2 px-2.5 rounded-sm text-xs font-bold flex items-center justify-center gap-1.5 transition cursor-pointer ${
            activeCatType === 'ranks'
              ? 'bg-[#007bff] text-white shadow-2xs'
              : 'bg-gray-100 text-gray-700 hover:bg-gray-200'
          }`}
        >
          <Award className="w-4 h-4" />
          <span>ឋានន្តរស័ក្កិ ({categories.ranks.length})</span>
        </button>

        <button
          onClick={() => onSelectCatType('positions')}
          className={`flex-1 min-w-[120px] py-2 px-2.5 rounded-sm text-xs font-bold flex items-center justify-center gap-1.5 transition cursor-pointer ${
            activeCatType === 'positions'
              ? 'bg-[#007bff] text-white shadow-2xs'
              : 'bg-gray-100 text-gray-700 hover:bg-gray-200'
          }`}
        >
          <Briefcase className="w-4 h-4" />
          <span>តួនាទី ({categories.positions.length})</span>
        </button>

        <button
          onClick={() => onSelectCatType('visaTeamsData')}
          className={`flex-1 min-w-[150px] py-2 px-2.5 rounded-sm text-xs font-bold flex items-center justify-center gap-1.5 transition cursor-pointer ${
            activeCatType === 'visaTeamsData' || activeCatType === 'visaTeams' || activeCatType === 'visaTeamsRobok'
              ? 'bg-[#007bff] text-white shadow-2xs'
              : 'bg-gray-100 text-gray-700 hover:bg-gray-200'
          }`}
        >
          <Table className="w-4 h-4" />
          <span>ទិន្នន័យក្រុមផ្តល់ទិដ្ឋាការ ({maxPairLength})</span>
        </button>

        <button
          onClick={() => onSelectCatType('collectorRoles')}
          className={`flex-1 min-w-[130px] py-2 px-2.5 rounded-sm text-xs font-bold flex items-center justify-center gap-1.5 transition cursor-pointer ${
            activeCatType === 'collectorRoles'
              ? 'bg-[#007bff] text-white shadow-2xs'
              : 'bg-gray-100 text-gray-700 hover:bg-gray-200'
          }`}
        >
          <UserCheck className="w-4 h-4" />
          <span>តួនាទីអ្នកមកបើក ({(categories.collectorRoles || []).length})</span>
        </button>

        <button
          onClick={() => onSelectCatType('organizations')}
          className={`flex-1 min-w-[120px] py-2 px-2.5 rounded-sm text-xs font-bold flex items-center justify-center gap-1.5 transition cursor-pointer ${
            activeCatType === 'organizations'
              ? 'bg-[#007bff] text-white shadow-2xs'
              : 'bg-gray-100 text-gray-700 hover:bg-gray-200'
          }`}
        >
          <Building2 className="w-4 h-4" />
          <span>ស្ថាប័នដើម ({(categories.organizations || []).length})</span>
        </button>

        <button
          onClick={() => onSelectCatType('refusalReasons')}
          className={`flex-1 min-w-[120px] py-2 px-2.5 rounded-sm text-xs font-bold flex items-center justify-center gap-1.5 transition cursor-pointer ${
            activeCatType === 'refusalReasons'
              ? 'bg-[#007bff] text-white shadow-2xs'
              : 'bg-gray-100 text-gray-700 hover:bg-gray-200'
          }`}
        >
          <FolderPlus className="w-4 h-4" />
          <span>ករណីបដិសេធ ({(categories.refusalReasons || []).length})</span>
        </button>
      </div>

      {/* Recently Added Confirmation Banner */}
      {recentlyAdded && (
        <div className="bg-emerald-50 border border-emerald-300 rounded-md p-3.5 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-3 animate-fade shadow-xs">
          <div className="flex items-center gap-2.5">
            <CheckCircle className="w-5 h-5 text-emerald-600 shrink-0" />
            <div>
              <p className="text-xs font-bold text-emerald-900">
                បានបន្ថែម "{recentlyAdded}" ទៅក្នុងប្រភេទ {catTypeLabels[activeCatType]} រួចរាល់!
              </p>
              <p className="text-[11px] text-emerald-700 mt-0.5">
                នៅក្នុង "ទម្រង់បញ្ចូលមន្ត្រី" ត្រង់ ({catTypeLabels[activeCatType]}) នឹងបង្ហាញ "-- ជ្រើសរើស --" ជាលំនាំដើម ព្រមទាំងមានជម្រើសថ្មីនេះ!
              </p>
            </div>
          </div>
          <button
            onClick={onNavigateToOfficerForm}
            className="bg-[#28a745] hover:bg-[#218838] text-white font-bold text-xs px-3.5 py-1.5 rounded transition shrink-0 flex items-center gap-1.5 cursor-pointer shadow-2xs"
          >
            <span>ទៅកាន់ទម្រង់បញ្ចូលមន្ត្រី</span>
            <ArrowRight className="w-3.5 h-3.5" />
          </button>
        </div>
      )}

      {/* Main Category Card */}
      <div className="bg-white rounded-md border border-gray-300 border-t-4 border-t-[#007bff] p-5 shadow-xs">
        <div className="border-b border-gray-200 pb-3 mb-4 flex flex-col sm:flex-row sm:items-center justify-between gap-3">
          <div>
            <h3 className="text-sm font-bold text-gray-800 flex items-center gap-2">
              {activeCatType === 'visaTeamsData' ? (
                <Table className="w-4 h-4 text-[#007bff]" />
              ) : (
                <FolderPlus className="w-4 h-4 text-[#007bff]" />
              )}
              <span>
                {activeCatType === 'visaTeamsData'
                  ? 'តារាងទិន្នន័យក្រុមផ្តល់ទិដ្ឋាការ (Visa Teams Combined Table)'
                  : `បញ្ចូលប្រភេទ៖ ${catTypeLabels[activeCatType]}`}
              </span>
            </h3>
            <p className="text-xs text-gray-500 mt-0.5">
              {activeCatType === 'visaTeamsData'
                ? 'បង្ហាញឈ្មោះក្រុមផ្តល់ទិដ្ឋាការពេញលេញ ទន្ទឹមគ្នាជាមួយឈ្មោះក្រុមផ្តល់ទិដ្ឋាការ.របក'
                : 'ប្រភេទដែលបានបន្ថែមនៅទីនេះ នឹងបង្ហាញក្នុងបញ្ជីជ្រើសរើសនៃទម្រង់បញ្ចូលមន្ត្រី'}
            </p>
          </div>

          <div className="flex items-center gap-2 flex-wrap">
            {(activeCatType === 'visaTeamsData' || activeCatType === 'visaTeams' || activeCatType === 'visaTeamsRobok') && (
              <>
                <button
                  type="button"
                  onClick={handleExportExcel}
                  className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs px-3 py-1.5 rounded transition flex items-center gap-1.5 shadow-2xs cursor-pointer"
                  title="ទាញយកទិន្នន័យជាប្រព័ន្ធ Excel (.xlsx)"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>Export Excel</span>
                </button>

                <button
                  type="button"
                  onClick={() => setShowImportModal(true)}
                  className="bg-indigo-600 hover:bg-indigo-700 text-white font-bold text-xs px-3 py-1.5 rounded transition flex items-center gap-1.5 shadow-2xs cursor-pointer"
                  title="នាំចូលទិន្នន័យក្រុមពី File Excel (.xlsx)"
                >
                  <Upload className="w-3.5 h-3.5" />
                  <span>Import Excel</span>
                </button>
              </>
            )}
            <span className="bg-[#007bff]/10 text-[#007bff] font-bold text-xs px-2.5 py-1 rounded border border-[#007bff]/20">
              {catTypeLabels[activeCatType]}
            </span>
          </div>
        </div>

        {activeCatType === 'visaTeamsData' || activeCatType === 'visaTeams' || activeCatType === 'visaTeamsRobok' ? (
          /* COMBINED VISA TEAMS TABLE VIEW */
          <div className="space-y-5">
            {/* Form to add team pair */}
            <form onSubmit={handleAddTeamPair} className="mb-5 bg-gray-50 p-3.5 rounded border border-gray-200">
              <h4 className="text-xs font-bold text-gray-700 mb-2 flex items-center gap-1.5">
                <Plus className="w-4 h-4 text-[#007bff]" />
                <span>បន្ថែមទិន្នន័យក្រុមផ្តល់ទិដ្ឋាការថ្មី (ជ្រើសរើស ឬបញ្ចូល)</span>
              </h4>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 mb-3">
                <div>
                  <label className="block text-[11px] font-bold text-gray-600 mb-1">
                    ក្រុមផ្តល់ទិដ្ឋាការ (ឈ្មោះពេញ)
                  </label>
                  <div className="space-y-1.5">
                    {(categories.visaTeams || []).length > 0 && (
                      <select
                        value={newTeamFullName}
                        onChange={(e) => setNewTeamFullName(e.target.value)}
                        className="w-full border border-gray-300 rounded-sm px-2.5 py-1 text-xs text-gray-700 bg-white focus:border-[#007bff] focus:outline-none"
                      >
                        <option value="">-- ជ្រើសរើសពីបញ្ជីដែលមានស្រាប់ --</option>
                        {(categories.visaTeams || []).map((vt) => (
                          <option key={vt.id} value={vt.name}>
                            {vt.name}
                          </option>
                        ))}
                      </select>
                    )}
                    <input
                      type="text"
                      value={newTeamFullName}
                      onChange={(e) => setNewTeamFullName(e.target.value)}
                      placeholder="ឬបញ្ចូលឈ្មោះពេញថ្មី៖ ឧ. ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារ..."
                      className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                    />
                  </div>
                </div>

                <div>
                  <label className="block text-[11px] font-bold text-gray-600 mb-1">
                    ក្រុមផ្តល់ទិដ្ឋាការ.របក (ឈ្មោះខ្លី)
                  </label>
                  <div className="space-y-1.5">
                    {(categories.visaTeamsRobok || []).length > 0 && (
                      <select
                        value={newTeamRobokName}
                        onChange={(e) => setNewTeamRobokName(e.target.value)}
                        className="w-full border border-gray-300 rounded-sm px-2.5 py-1 text-xs text-gray-700 bg-white focus:border-[#007bff] focus:outline-none"
                      >
                        <option value="">-- ជ្រើសរើសពីបញ្ជីដែលមានស្រាប់ --</option>
                        {(categories.visaTeamsRobok || []).map((vtr) => (
                          <option key={vtr.id} value={vtr.name}>
                            {vtr.name}
                          </option>
                        ))}
                      </select>
                    )}
                    <input
                      type="text"
                      value={newTeamRobokName}
                      onChange={(e) => setNewTeamRobokName(e.target.value)}
                      placeholder="ឬបញ្ចូលឈ្មោះខ្លីថ្មី៖ ឧ. អាកាស តេជោ..."
                      className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                    />
                  </div>
                </div>
              </div>

              <div className="flex justify-end">
                <button
                  type="submit"
                  className="bg-[#007bff] hover:bg-[#0069d9] text-white font-bold text-xs px-4 py-1.5 rounded transition flex items-center gap-1.5 shadow-2xs cursor-pointer"
                >
                  <Plus className="w-3.5 h-3.5" />
                  <span>បន្ថែមទិន្នន័យក្រុម</span>
                </button>
              </div>
            </form>

            {/* Combined Table matching official document formatting */}
            <div className="overflow-x-auto">
              <table className="w-full border-collapse border border-black text-xs text-black">
                <thead>
                  <tr className="bg-white text-center font-bold">
                    <th className="border border-black px-2 py-2 w-12 text-center text-sm font-bold">
                      ល.រ
                    </th>
                    <th className="border border-black px-4 py-2 text-center text-sm font-bold">
                      ក្រុមផ្តល់ទិដ្ឋាការ
                    </th>
                    <th className="border border-black px-4 py-2 text-center text-sm font-bold">
                      ក្រុមផ្តល់ទិដ្ឋាការ.របក
                    </th>
                    <th className="border border-black px-3 py-2 text-center text-xs font-bold min-w-[210px]">
                      <div className="flex flex-col items-center justify-center gap-1">
                        <span className="text-sm font-bold">ប្រភេទប្រើប្រាស់ (សន្លឹកទិដ្ឋាការ, ក្រដាសអនុម័ត)</span>
                        <div className="flex items-center gap-1 text-[11px] font-normal print:hidden mt-0.5">
                          <span className="text-gray-600 font-medium">ជ្រើសរើសទាំងអស់:</span>
                          <button
                            type="button"
                            onClick={() => handleSetAllTypeOption('all')}
                            className="px-1.5 py-0.5 rounded bg-gray-200 hover:bg-gray-300 text-gray-800 text-[10px] font-bold cursor-pointer transition"
                          >
                            ទាំងអស់
                          </button>
                          <button
                            type="button"
                            onClick={() => handleSetAllTypeOption('sticker')}
                            className="px-1.5 py-0.5 rounded bg-blue-100 hover:bg-blue-200 text-blue-800 text-[10px] font-bold cursor-pointer transition"
                          >
                            សន្លឹកទិដ្ឋាការ
                          </button>
                          <button
                            type="button"
                            onClick={() => handleSetAllTypeOption('evisa')}
                            className="px-1.5 py-0.5 rounded bg-amber-100 hover:bg-amber-200 text-amber-800 text-[10px] font-bold cursor-pointer transition"
                          >
                            ក្រដាសអនុម័ត
                          </button>
                        </div>
                      </div>
                    </th>
                    <th className="border border-black px-2 py-2 w-24 text-center text-xs font-bold print:hidden">
                      សកម្មភាព
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {Array.from({ length: maxPairLength }).map((_, idx) => {
                    const vtItem = categories.visaTeams?.[idx];
                    const vtrItem = categories.visaTeamsRobok?.[idx];
                    const robok = vtrItem?.name || '';
                    let full = vtItem?.name || '';
                    if ((!full || /^ក្រុមទី\s*\d+$/i.test(full.trim())) && robok) {
                      full = formatReportTeamName(robok);
                    }
                    const isEditing = editingPairIdx === idx;
                    const opt = teamTypeOptions[idx] || { sticker: true, evisa: true };

                    return (
                      <tr key={idx} className="hover:bg-gray-50/80 transition">
                        <td className="border border-black px-2 py-2 text-center font-bold text-sm">
                          {toKhmerNum(idx + 1)}
                        </td>
                        <td className="border border-black px-3 py-2 text-left font-normal text-sm font-siemreap">
                          {isEditing ? (
                            <div className="space-y-1">
                              {(categories.visaTeams || []).length > 0 && (
                                <select
                                  value={editingPairFull}
                                  onChange={(e) => setEditingPairFull(e.target.value)}
                                  className="w-full border border-gray-300 rounded px-1.5 py-0.5 text-xs bg-white font-siemreap"
                                >
                                  <option value="">-- ជ្រើសរើស --</option>
                                  {(categories.visaTeams || []).map((vt) => (
                                    <option key={vt.id} value={vt.name}>
                                      {vt.name}
                                    </option>
                                  ))}
                                </select>
                              )}
                              <input
                                type="text"
                                value={editingPairFull}
                                onChange={(e) => setEditingPairFull(e.target.value)}
                                className="w-full border border-gray-400 rounded px-2 py-1 text-xs focus:outline-none focus:border-[#007bff] font-siemreap"
                              />
                            </div>
                          ) : (
                            <span className="font-siemreap font-medium text-[13px] text-gray-900 leading-snug block">
                              {full || '-'}
                            </span>
                          )}
                        </td>
                        <td className="border border-black px-3 py-2 text-left font-normal text-sm font-siemreap">
                          {isEditing ? (
                            <div className="space-y-1">
                              {(categories.visaTeamsRobok || []).length > 0 && (
                                <select
                                  value={editingPairRobok}
                                  onChange={(e) => setEditingPairRobok(e.target.value)}
                                  className="w-full border border-gray-300 rounded px-1.5 py-0.5 text-xs bg-white font-siemreap"
                                >
                                  <option value="">-- ជ្រើសរើស --</option>
                                  {(categories.visaTeamsRobok || []).map((vtr) => (
                                    <option key={vtr.id} value={vtr.name}>
                                      {vtr.name}
                                    </option>
                                  ))}
                                </select>
                              )}
                              <input
                                type="text"
                                value={editingPairRobok}
                                onChange={(e) => setEditingPairRobok(e.target.value)}
                                className="w-full border border-gray-400 rounded px-2 py-1 text-xs focus:outline-none focus:border-[#007bff] font-siemreap"
                              />
                            </div>
                          ) : (
                            <span className="font-siemreap font-medium text-[13px] text-gray-900 leading-snug block">
                              {robok || '-'}
                            </span>
                          )}
                        </td>
                        <td className="border border-black px-3 py-2 text-center">
                          <div className="flex items-center justify-center gap-1.5 flex-wrap">
                            <button
                              type="button"
                              onClick={() => handleToggleTypeOption(idx, 'sticker')}
                              className={`px-2.5 py-1 rounded-md text-[11px] font-bold transition flex items-center gap-1.5 cursor-pointer border ${
                                opt.sticker
                                  ? 'bg-blue-600 text-white border-blue-700 shadow-2xs'
                                  : 'bg-gray-100 text-gray-500 border-gray-300 hover:bg-gray-200'
                              }`}
                              title="ចុចដើម្បី បើក/បិទ ជម្រើស សន្លឹកទិដ្ឋាការ"
                            >
                              <span className="text-[12px]">{opt.sticker ? '✓' : '○'}</span>
                              <span>សន្លឹកទិដ្ឋាការ</span>
                            </button>

                            <button
                              type="button"
                              onClick={() => handleToggleTypeOption(idx, 'evisa')}
                              className={`px-2.5 py-1 rounded-md text-[11px] font-bold transition flex items-center gap-1.5 cursor-pointer border ${
                                opt.evisa
                                  ? 'bg-amber-600 text-white border-amber-700 shadow-2xs'
                                  : 'bg-gray-100 text-gray-500 border-gray-300 hover:bg-gray-200'
                              }`}
                              title="ចុចដើម្បី បើក/បិទ ជម្រើស ក្រដាសអនុម័ត"
                            >
                              <span className="text-[12px]">{opt.evisa ? '✓' : '○'}</span>
                              <span>ក្រដាសអនុម័ត</span>
                            </button>
                          </div>
                        </td>
                        <td className="border border-black px-2 py-2 text-center print:hidden">
                          {isEditing ? (
                            <div className="flex items-center justify-center gap-1">
                              <button
                                type="button"
                                onClick={() => handleSaveEditPair(idx)}
                                className="bg-emerald-600 hover:bg-emerald-700 text-white p-1 rounded transition cursor-pointer"
                                title="រក្សាទុក"
                              >
                                <Save className="w-3.5 h-3.5" />
                              </button>
                              <button
                                type="button"
                                onClick={handleCancelEditPair}
                                className="bg-gray-500 hover:bg-gray-600 text-white p-1 rounded transition cursor-pointer"
                                title="បោះបង់"
                              >
                                <X className="w-3.5 h-3.5" />
                              </button>
                            </div>
                          ) : (
                            <div className="flex items-center justify-center gap-1">
                              <button
                                type="button"
                                onClick={() => handleStartEditPair(idx, full, robok)}
                                className="text-[#007bff] hover:bg-blue-50 p-1 rounded transition cursor-pointer"
                                title="កែប្រែ"
                              >
                                <Edit2 className="w-3.5 h-3.5" />
                              </button>
                              <button
                                type="button"
                                onClick={() =>
                                  setDeletingPair({ idx, full, robok })
                                }
                                className="text-gray-400 hover:text-rose-600 hover:bg-rose-50 p-1 rounded transition cursor-pointer"
                                title="លុប"
                              >
                                <Trash2 className="w-3.5 h-3.5" />
                              </button>
                            </div>
                          )}
                        </td>
                      </tr>
                    );
                  })}
                  {maxPairLength === 0 && (
                    <tr>
                      <td colSpan={5} className="border border-black p-4 text-center text-gray-500 text-xs">
                        មិនទាន់មានទិន្នន័យនៅឡើយទេ
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
          </div>
        ) : (
          /* STANDARD CATEGORY MANAGEMENT FORM & LIST */
          <>
            {/* Input Form to add new item */}
            <form onSubmit={handleAdd} className="mb-5">
              <label className="block text-xs font-bold text-gray-700 mb-1">
                ឈ្មោះ {catTypeLabels[activeCatType]} ថ្មី
              </label>
              <div className="flex gap-2">
                <input
                  type="text"
                  value={newItemName}
                  onChange={(e) => setNewItemName(e.target.value)}
                  placeholder={`បញ្ចូលឈ្មោះ ${catTypeLabels[activeCatType]}...`}
                  className="flex-1 border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                />
                <button
                  type="submit"
                  className="bg-[#007bff] hover:bg-[#0069d9] text-white font-bold text-xs px-4 py-1.5 rounded transition flex items-center gap-1.5 shadow-2xs shrink-0 cursor-pointer"
                >
                  <Plus className="w-3.5 h-3.5" />
                  <span>បន្ថែម</span>
                </button>
              </div>
            </form>

            {/* List of Existing Category Items */}
            <div>
              <h4 className="text-xs font-bold text-gray-700 uppercase tracking-wider mb-2.5">
                បញ្ជី {catTypeLabels[activeCatType]} ដែលមានស្រាប់ ({currentList.length})
              </h4>

              {currentList.length === 0 ? (
                <div className="text-center py-6 border border-dashed border-gray-300 rounded text-gray-400 text-xs">
                  មិនទាន់មានទិន្នន័យនៅឡើយទេ។ សូមបញ្ចូលឈ្មោះខាងលើ។
                </div>
              ) : (
                <ul className="divide-y divide-gray-200 border border-gray-200 rounded overflow-hidden bg-gray-50/50">
                  {currentList.map((item, idx) => (
                    <li
                      key={item.id}
                      className="px-3.5 py-2 flex items-center justify-between hover:bg-white transition text-xs gap-2"
                    >
                      {editingId === item.id ? (
                        <div className="flex items-center gap-2 w-full">
                          <span className="w-5 h-5 rounded-full bg-[#007bff]/10 text-[#007bff] font-bold text-[10px] flex items-center justify-center shrink-0">
                            {idx + 1}
                          </span>
                          <input
                            type="text"
                            value={editingName}
                            onChange={(e) => setEditingName(e.target.value)}
                            autoFocus
                            onKeyDown={(e) => {
                              if (e.key === 'Enter') handleSaveEdit(item.id);
                              if (e.key === 'Escape') handleCancelEdit();
                            }}
                            className="flex-1 border border-gray-300 rounded-sm px-2.5 py-1 text-xs focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white text-gray-800"
                          />
                          <button
                            type="button"
                            onClick={() => handleSaveEdit(item.id)}
                            title="រក្សាទុក"
                            className="bg-[#28a745] hover:bg-[#218838] text-white text-xs px-2.5 py-1 rounded flex items-center gap-1 font-bold transition shrink-0 cursor-pointer"
                          >
                            <Save className="w-3 h-3" />
                            <span>រក្សាទុក</span>
                          </button>
                          <button
                            type="button"
                            onClick={handleCancelEdit}
                            title="បោះបង់"
                            className="bg-[#6c757d] hover:bg-[#5a6268] text-white text-xs px-2.5 py-1 rounded transition shrink-0 cursor-pointer"
                          >
                            <X className="w-3 h-3" />
                          </button>
                        </div>
                      ) : (
                        <>
                          <div className="flex items-center gap-2.5 min-w-0">
                            <span className="w-5 h-5 rounded-full bg-[#007bff]/10 text-[#007bff] font-bold text-[10px] flex items-center justify-center shrink-0">
                              {idx + 1}
                            </span>
                            <span className="font-semibold text-gray-800 truncate">{item.name}</span>
                          </div>

                          <div className="flex items-center gap-1 shrink-0">
                            <button
                              type="button"
                              onClick={() => handleStartEdit(item.id, item.name)}
                              title="កែប្រែ"
                              className="text-[#007bff] hover:bg-[#007bff]/10 p-1.5 rounded transition cursor-pointer"
                            >
                              <Edit2 className="w-3.5 h-3.5" />
                            </button>

                            <button
                              type="button"
                              onClick={() => handleDelete(item.id, item.name)}
                              title="លុប"
                              className="text-gray-400 hover:text-rose-600 hover:bg-rose-50 p-1.5 rounded transition cursor-pointer"
                            >
                              <Trash2 className="w-3.5 h-3.5" />
                            </button>
                          </div>
                        </>
                      )}
                    </li>
                  ))}
                </ul>
              )}
            </div>
          </>
        )}
      </div>

      {/* DELETE CONFIRMATION MODAL (SINGLE ITEM) */}
      {deletingItem && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-2xs flex items-center justify-center p-4">
          <div className="bg-white rounded-md border border-gray-300 border-t-4 border-t-[#dc3545] shadow-xl w-full max-w-sm p-5 animate-fade">
            <div className="flex items-center gap-3 text-[#dc3545] mb-2">
              <div className="w-8 h-8 rounded-full bg-rose-100 flex items-center justify-center shrink-0">
                <Trash2 className="w-4 h-4 text-[#dc3545]" />
              </div>
              <h3 className="text-sm font-bold text-gray-800">បញ្ជាក់ការលុប</h3>
            </div>
            <p className="text-xs text-gray-600 mb-4 leading-relaxed">
              តើអ្នកពិតជាចង់លុប <span className="font-bold text-gray-800">"{deletingItem.name}"</span> នេះចេញពី {catTypeLabels[activeCatType]} មែនទេ?
            </p>
            <div className="flex items-center justify-end gap-2">
              <button
                type="button"
                onClick={() => setDeletingItem(null)}
                className="min-w-[120px] bg-[#6c757d] hover:bg-[#5a6268] text-white px-4 py-1.5 rounded text-xs font-bold transition flex items-center justify-center cursor-pointer"
              >
                <span>បោះបង់</span>
              </button>
              <button
                type="button"
                onClick={handleConfirmDelete}
                className="min-w-[120px] bg-[#dc3545] hover:bg-[#c82333] text-white px-4 py-1.5 rounded text-xs font-bold shadow-2xs transition flex items-center justify-center gap-1.5 cursor-pointer"
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>លុបចេញ</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* DELETE CONFIRMATION MODAL (COMBINED PAIR) */}
      {deletingPair && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-2xs flex items-center justify-center p-4">
          <div className="bg-white rounded-md border border-gray-300 border-t-4 border-t-[#dc3545] shadow-xl w-full max-w-sm p-5 animate-fade">
            <div className="flex items-center gap-3 text-[#dc3545] mb-2">
              <div className="w-8 h-8 rounded-full bg-rose-100 flex items-center justify-center shrink-0">
                <Trash2 className="w-4 h-4 text-[#dc3545]" />
              </div>
              <h3 className="text-sm font-bold text-gray-800">បញ្ជាក់ការលុបទិន្នន័យក្រុម</h3>
            </div>
            <p className="text-xs text-gray-600 mb-4 leading-relaxed">
              តើអ្នកពិតជាចង់លុបទិន្នន័យក្រុមទិដ្ឋាការជួរទី <span className="font-bold text-gray-800">{toKhmerNum(deletingPair.idx + 1)}</span> នេះចេញមែនទេ?
            </p>
            <div className="flex items-center justify-end gap-2">
              <button
                type="button"
                onClick={() => setDeletingPair(null)}
                className="min-w-[120px] bg-[#6c757d] hover:bg-[#5a6268] text-white px-4 py-1.5 rounded text-xs font-bold transition flex items-center justify-center cursor-pointer"
              >
                <span>បោះបង់</span>
              </button>
              <button
                type="button"
                onClick={handleConfirmDeletePair}
                className="min-w-[120px] bg-[#dc3545] hover:bg-[#c82333] text-white px-4 py-1.5 rounded text-xs font-bold shadow-2xs transition flex items-center justify-center gap-1.5 cursor-pointer"
              >
                <Trash2 className="w-3.5 h-3.5" />
                <span>លុបចេញ</span>
              </button>
            </div>
          </div>
        </div>
      )}

      {/* IMPORT EXCEL MODAL */}
      {showImportModal && (
        <div className="fixed inset-0 z-50 bg-black/50 backdrop-blur-2xs flex items-center justify-center p-4">
          <div className="bg-white rounded-md border border-gray-300 border-t-4 border-t-indigo-600 shadow-xl w-full max-w-md p-5 animate-fade">
            <div className="flex items-center justify-between border-b border-gray-200 pb-3 mb-4">
              <div className="flex items-center gap-2 text-indigo-700">
                <FileSpreadsheet className="w-5 h-5" />
                <h3 className="text-sm font-bold">នាំចូលទិន្នន័យក្រុមពី Excel (.xlsx, .csv)</h3>
              </div>
              <button
                onClick={() => setShowImportModal(false)}
                className="text-gray-400 hover:text-gray-600 p-1 rounded transition cursor-pointer"
              >
                <X className="w-4 h-4" />
              </button>
            </div>

            <p className="text-xs text-gray-600 mb-4 leading-relaxed">
              លោកអ្នកអាចទាញយកទម្រង់គំរូ ឬជ្រើសរើស File Excel ដែលមានជួរឈរ <span className="font-bold text-gray-800">"ក្រុមផ្តល់ទិដ្ឋាការ (ឈ្មោះពេញ)"</span>, <span className="font-bold text-gray-800">"ក្រុមផ្តល់ទិដ្ឋាការ.របក (ឈ្មោះខ្លី)"</span> និង <span className="font-bold text-gray-800">"ប្រភេទប្រើប្រាស់"</span> ដើម្បីបញ្ចូលក្នុងប្រព័ន្ធ។
            </p>

            <div className="flex items-center justify-between mb-4 bg-gray-50 p-2.5 rounded border border-gray-200">
              <span className="text-xs text-gray-700 font-medium">មិនទាន់មានទម្រង់?</span>
              <button
                type="button"
                onClick={handleDownloadImportTemplate}
                className="bg-emerald-600 hover:bg-emerald-700 text-white font-bold text-xs px-3 py-1 rounded transition flex items-center gap-1 cursor-pointer"
              >
                <Download className="w-3.5 h-3.5" />
                <span>ទាញយកទម្រង់គំរូ (.xlsx)</span>
              </button>
            </div>

            <div className="mb-5">
              <label className="block text-xs font-bold text-gray-700 mb-2">
                ជ្រើសរើស File Excel ឬ CSV៖
              </label>
              <div className="border-2 border-dashed border-gray-300 hover:border-indigo-500 rounded-md p-4 text-center bg-gray-50/50 transition relative cursor-pointer">
                <input
                  type="file"
                  accept=".xlsx, .xls, .csv"
                  onChange={handleFileUpload}
                  className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
                />
                <Upload className="w-8 h-8 text-indigo-500 mx-auto mb-1.5" />
                <p className="text-xs font-bold text-gray-800">ចុចទីនេះ ដើម្បីជ្រើសរើស File Excel</p>
                <p className="text-[11px] text-gray-500 mt-0.5">គាំទ្រប្រភេទ .xlsx, .xls, .csv</p>
              </div>
            </div>

            <div className="flex items-center justify-end gap-2">
              <button
                type="button"
                onClick={() => setShowImportModal(false)}
                className="bg-gray-500 hover:bg-gray-600 text-white px-4 py-1.5 rounded text-xs font-bold transition cursor-pointer"
              >
                <span>បិទ</span>
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
