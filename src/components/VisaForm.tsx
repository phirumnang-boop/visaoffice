import React, { useState } from 'react';
import { Officer, CategoryItem, VisaRecord } from '../types';
import { Info, Plus, Trash2, Check, Calendar, Home, FileText, Users } from 'lucide-react';
import { CustomDatePicker } from './CustomDatePicker';

interface VisaFormProps {
  officers: Officer[];
  visaTeams: CategoryItem[];
  organizations?: CategoryItem[];
  onAddVisaRecord: (record: Omit<VisaRecord, 'id' | 'createdAt'>) => void;
  onShowToast: (msg: string, type: 'success' | 'error' | 'info') => void;
  onNavigateToList: () => void;
}

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
  { code: '1M', label: '1M (1 ខែ)', fee: 50, duration: '1 ខែ' },
  { code: '3M', label: '3M (3 ខែ)', fee: 80, duration: '1 ខែ' },
  { code: '6M', label: '6M (6 ខែ)', fee: 180, duration: '6 ខែ' },
  { code: '1Y', label: '1Y (12 ខែ / 1 ឆ្នាំ)', fee: 290, duration: '12 ខែ' },
];

interface PersonRow {
  id: string;
  fullName: string;
  gender: 'ប្រុស' | 'ស្រី';
  passportNumber: string;
  nationality: string;
}

export const VisaForm: React.FC<VisaFormProps> = ({
  officers,
  visaTeams,
  organizations = [],
  onAddVisaRecord,
  onShowToast,
  onNavigateToList,
}) => {
  const todayStr = new Date().toISOString().split('T')[0];
  const formattedToday = new Date().toLocaleDateString('en-US', {
    month: 'long',
    day: 'numeric',
    year: 'numeric',
  });

  // Common form fields
  const [applicationDate, setApplicationDate] = useState(todayStr);
  const [oldVisaType, setOldVisaType] = useState('');
  const [newVisaType, setNewVisaType] = useState('');
  const [documentNumber, setDocumentNumber] = useState('');
  const [issueDate, setIssueDate] = useState('');
  const [organization, setOrganization] = useState('');
  const [extension, setExtension] = useState('');

  // Person rows list
  const [people, setPeople] = useState<PersonRow[]>([
    {
      id: '1',
      fullName: '',
      gender: 'ប្រុស',
      passportNumber: '',
      nationality: '',
    },
  ]);

  const handleAddPerson = () => {
    setPeople((prev) => [
      ...prev,
      {
        id: Date.now().toString(),
        fullName: '',
        gender: 'ប្រុស',
        passportNumber: '',
        nationality: '',
      },
    ]);
  };

  const handleRemovePerson = (id: string) => {
    if (people.length === 1) {
      onShowToast('ត្រូវតែមានយ៉ាងហោចណាស់ម្នាក់ក្នុងបញ្ជី', 'error');
      return;
    }
    setPeople((prev) => prev.filter((p) => p.id !== id));
  };

  const updatePerson = (id: string, field: keyof PersonRow, value: any) => {
    setPeople((prev) =>
      prev.map((p) => (p.id === id ? { ...p, [field]: value } : p))
    );
  };

  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();

    // Check if at least passport and name are filled for all people
    const hasInvalid = people.some(
      (p) => !p.fullName.trim() || !p.passportNumber.trim()
    );

    if (hasInvalid) {
      onShowToast('សូមបំពេញឈ្មោះ និងលេខលិខិតឆ្លងដែនសម្រាប់បុគ្គលគ្រប់រូប', 'error');
      return;
    }

    const selectedExtensionObj = EXTENSIONS.find((e) => e.code === extension);
    const durationVal = selectedExtensionObj ? selectedExtensionObj.duration : '1 ខែ';
    const feeVal = selectedExtensionObj ? selectedExtensionObj.fee : 50;

    const defaultOfficer = officers[0];
    const defaultTeam = visaTeams[0];

    people.forEach((p) => {
      onAddVisaRecord({
        passportNumber: p.passportNumber.trim().toUpperCase(),
        fullName: p.fullName.trim().toUpperCase(),
        nationality: p.nationality.trim() || 'ចិន',
        gender: p.gender,
        oldVisaType: oldVisaType || 'T',
        newVisaType: newVisaType || 'C',
        documentNumber: documentNumber.trim(),
        issueDate: issueDate || applicationDate,
        organization: organization.trim(),
        extension: extension || '1M',
        duration: durationVal,
        fee: feeVal,
        applicationDate,
        officerId: defaultOfficer?.id,
        officerName: defaultOfficer?.name,
        visaTeamId: defaultTeam?.id,
        status: 'អនុម័តរួច',
        notes: '',
      });
    });

    onShowToast(`បានរក្សាទុកទិន្នន័យប្តូរទិដ្ឋាការចំនួន ${people.length} នាក់ដោយជោគជ័យ!`, 'success');
    onNavigateToList();
  };

  return (
    <div className="space-y-5 w-full animate-fade">
      {/* AdminLTE Content Header (Breadcrumbs) */}
      <div className="flex flex-col sm:flex-row sm:items-center justify-between gap-2 bg-white px-4 py-3 rounded-md border border-gray-200 shadow-xs">
        <div>
          <h1 className="text-base font-bold text-gray-800 flex items-center gap-2">
            <span>ទម្រង់ប្តូរប្រភេទទិដ្ឋាការ (Visa Category Change Form)</span>
          </h1>
          <p className="text-xs text-gray-500">
            បញ្ចូលទិន្នន័យស្នើសុំប្តូរប្រភេទទិដ្ឋាការសម្រាប់បុគ្គល ឬក្រុមបុគ្គល
          </p>
        </div>
        <div className="flex items-center gap-1.5 text-xs text-gray-500 bg-gray-100 px-3 py-1.5 rounded-md font-medium self-start sm:self-auto">
          <Home className="w-3.5 h-3.5 text-gray-500" />
          <span>ទំព័រដើម</span>
          <span>/</span>
          <span>ការងារទិដ្ឋាការ</span>
          <span>/</span>
          <span className="text-[#007bff] font-bold">ប្តូរប្រភេទទិដ្ឋាការ</span>
        </div>
      </div>

      {/* AdminLTE Card Container */}
      <div className="bg-white rounded-md border border-gray-300 border-t-4 border-t-[#007bff] shadow-xs overflow-hidden">
        {/* Card Header */}
        <div className="bg-[#007bff] text-white px-4 py-3 flex items-center justify-between">
          <div className="flex items-center gap-2 font-bold text-sm">
            <FileText className="w-4 h-4 text-white" />
            <span>ទម្រង់ប្តូរប្រភេទទិដ្ឋាការ (Visa Change Request)</span>
          </div>
          <div className="flex items-center gap-1.5 text-xs bg-white/20 px-2.5 py-1 rounded font-mono">
            <Calendar className="w-3.5 h-3.5" />
            <span>{formattedToday}</span>
          </div>
        </div>

        {/* Card Body */}
        <form onSubmit={handleSubmit} className="p-6 space-y-6">
          {/* Info Callout Banner */}
          <div className="bg-[#17a2b8]/10 border-l-4 border-[#17a2b8] p-3.5 rounded-r-md text-[#17a2b8] text-xs flex items-center gap-2.5">
            <Info className="w-4 h-4 shrink-0" />
            <span className="text-gray-700 font-medium">
              ទម្រង់នេះអនុញ្ញាតឱ្យអ្នកបញ្ចូលបុគ្គលច្រើនក្នុងការបញ្ចូលតែមួយលើក។ ទិន្នន័យរួម (ថ្ងៃខែ, ទិដ្ឋាការ, ស្ថាប័ន, លិខិត, ពន្យារ) នឹងត្រូវបានអនុវត្តចំពោះមនុស្សទាំងអស់។
            </span>
          </div>

          {/* Section 1: ព័ត៌មានទូទៅ */}
          <div className="space-y-4">
            <h4 className="text-xs font-bold text-gray-800 uppercase tracking-wide border-b border-gray-200 pb-2">
              ១. ព័ត៌មានទូទៅ (General Information)
            </h4>

            {/* All General Information fields in 1 horizontal row on large screens */}
            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-4 lg:grid-cols-7 gap-2.5">
              <div>
                <label className="block text-xs font-bold text-gray-700 mb-1">
                  ថ្ងៃខែឆ្នាំ <span className="text-red-500">*</span>
                </label>
                <CustomDatePicker
                  required
                  value={applicationDate}
                  onChange={(d) => setApplicationDate(d)}
                  className="py-1 text-xs"
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-gray-700 mb-1">
                  ទិដ្ឋាការដំបូង
                </label>
                <select
                  value={oldVisaType}
                  onChange={(e) => setOldVisaType(e.target.value)}
                  className="w-full border border-gray-300 rounded-sm px-2 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                >
                  <option value="">-- ជ្រើសរើស --</option>
                  {VISA_TYPES.map((vt) => (
                    <option key={vt} value={vt}>
                      {vt}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-xs font-bold text-gray-700 mb-1">
                  ប្តូរទិដ្ឋាការ
                </label>
                <select
                  value={newVisaType}
                  onChange={(e) => setNewVisaType(e.target.value)}
                  className="w-full border border-gray-300 rounded-sm px-2 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                >
                  <option value="">-- ជ្រើសរើស --</option>
                  {VISA_TYPES.map((vt) => (
                    <option key={vt} value={vt}>
                      {vt}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-xs font-bold text-gray-700 mb-1">
                  លេខលិខិត
                </label>
                <input
                  type="text"
                  placeholder="លេខលិខិត"
                  value={documentNumber}
                  onChange={(e) => setDocumentNumber(e.target.value)}
                  className="w-full border border-gray-300 rounded-sm px-2 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none transition"
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-gray-700 mb-1">
                  ថ្ងៃចេញលិខិត
                </label>
                <CustomDatePicker
                  value={issueDate}
                  onChange={(d) => setIssueDate(d)}
                  className="py-1 text-xs"
                />
              </div>

              <div>
                <label className="block text-xs font-bold text-gray-700 mb-1">
                  ស្ថាប័នដើម
                </label>
                <select
                  value={organization}
                  onChange={(e) => setOrganization(e.target.value)}
                  className="w-full border border-gray-300 rounded-sm px-2 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                >
                  <option value="">-- ជ្រើសរើស --</option>
                  {organizations.map((org) => (
                    <option key={org.id} value={org.name}>
                      {org.name}
                    </option>
                  ))}
                </select>
              </div>

              <div>
                <label className="block text-xs font-bold text-gray-700 mb-1">
                  ពន្យារ (Extension)
                </label>
                <select
                  value={extension}
                  onChange={(e) => setExtension(e.target.value)}
                  className="w-full border border-gray-300 rounded-sm px-2 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                >
                  <option value="">-- ជ្រើសរើស --</option>
                  {EXTENSIONS.map((ext) => (
                    <option key={ext.code} value={ext.code}>
                      {ext.label}
                    </option>
                  ))}
                </select>
              </div>
            </div>
          </div>

          {/* Section 2: បញ្ជីបុគ្គល */}
          <div className="space-y-3 pt-2">
            <div className="flex items-center justify-between border-b border-gray-200 pb-2">
              <h4 className="text-xs font-bold text-gray-800 uppercase tracking-wide flex items-center gap-2">
                <Users className="w-4 h-4 text-[#007bff]" />
                <span>២. បញ្ជីបុគ្គល (Applicants List) - ({people.length} នាក់)</span>
              </h4>
            </div>

            <div className="space-y-2.5">
              {people.map((p, index) => (
                <div
                  key={p.id}
                  className="flex items-center gap-2 bg-gray-50 p-2 rounded-sm border border-gray-200 hover:border-gray-300 transition"
                >
                  <span className="text-xs font-bold bg-[#007bff] text-white w-6 h-6 rounded-full flex items-center justify-center shrink-0">
                    {index + 1}
                  </span>

                  <div className="flex-1 min-w-[140px]">
                    <input
                      type="text"
                      required
                      placeholder="ឈ្មោះ (NAME)"
                      value={p.fullName}
                      onChange={(e) =>
                        updatePerson(p.id, 'fullName', e.target.value.toUpperCase())
                      }
                      className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs uppercase text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                    />
                  </div>

                  <div className="w-24 shrink-0">
                    <select
                      value={p.gender}
                      onChange={(e) =>
                        updatePerson(p.id, 'gender', e.target.value as any)
                      }
                      className="w-full border border-gray-300 rounded-sm px-2 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                    >
                      <option value="ប្រុស">ប្រុស</option>
                      <option value="ស្រី">ស្រី</option>
                    </select>
                  </div>

                  <div className="flex-1 min-w-[140px]">
                    <input
                      type="text"
                      required
                      placeholder="លេខលិខិតឆ្លងដែន (PASSPORT)"
                      value={p.passportNumber}
                      onChange={(e) =>
                        updatePerson(p.id, 'passportNumber', e.target.value.toUpperCase())
                      }
                      className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs uppercase text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                    />
                  </div>

                  <div className="flex-1 min-w-[120px]">
                    <input
                      type="text"
                      placeholder="សញ្ជាតិ (Nationality)"
                      value={p.nationality}
                      onChange={(e) =>
                        updatePerson(p.id, 'nationality', e.target.value)
                      }
                      className="w-full border border-gray-300 rounded-sm px-2.5 py-1.5 text-xs text-gray-800 focus:border-[#007bff] focus:ring-1 focus:ring-[#007bff] focus:outline-none bg-white transition"
                    />
                  </div>

                  {people.length > 1 && (
                    <button
                      type="button"
                      onClick={() => handleRemovePerson(p.id)}
                      className="p-1.5 text-red-500 hover:text-red-700 hover:bg-red-50 rounded transition shrink-0 cursor-pointer"
                      title="លុប"
                    >
                      <Trash2 className="w-4 h-4" />
                    </button>
                  )}
                </div>
              ))}

              <button
                type="button"
                onClick={handleAddPerson}
                className="w-full py-2 border border-dashed border-[#007bff] rounded-sm text-xs font-bold text-[#007bff] bg-blue-50/40 hover:bg-[#007bff] hover:text-white transition flex items-center justify-center gap-1.5 mt-2 cursor-pointer"
              >
                <Plus className="w-4 h-4" />
                <span>+ បន្ថែមបុគ្គល (Add Applicant)</span>
              </button>
            </div>
          </div>

          {/* Card Footer */}
          <div className="pt-4 border-t border-gray-200 flex items-center justify-end bg-gray-50 -mx-6 -mb-6 px-6 py-3.5 rounded-b-sm">
            <button
              type="submit"
              className="bg-[#28a745] hover:bg-[#218838] text-white px-6 py-2 rounded text-xs font-bold flex items-center gap-2 shadow-2xs transition cursor-pointer"
            >
              <Check className="w-4 h-4" />
              <span>✓ រក្សាទុកទិន្នន័យ (Save Records)</span>
            </button>
          </div>
        </form>
      </div>
    </div>
  );
};

