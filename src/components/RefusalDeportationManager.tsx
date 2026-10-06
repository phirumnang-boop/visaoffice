import React, { useState, useEffect, useMemo, useRef } from 'react';
import { RefusalDeportationRecord, RefusalDeportationPerson, UserRole, CategoriesState } from '../types';
import {
  Plus,
  Trash2,
  Edit2,
  FileText,
  Search,
  Calendar,
  Layers,
  CheckCircle2,
  Printer,
  Download,
  ArrowLeft,
  X,
  Globe,
  Plane,
  ShieldAlert,
  Users,
  Info,
  Lock,
  Unlock,
} from 'lucide-react';
import { normalizeTeamName } from '../utils/teamNormalization';
import { toKhmerNum } from '../utils/khmerCalendar';
import { exportElementToPdf } from '../utils/pdfExportHelper';
import { printA4Document } from '../utils/printHelper';
import { CustomDatePicker } from './CustomDatePicker';

interface RefusalDeportationManagerProps {
  currentRole: UserRole;
  assignedTeam?: string;
  userName: string;
  categories?: CategoriesState;
  records?: RefusalDeportationRecord[];
  onUpdateRecords?: (records: RefusalDeportationRecord[]) => void;
  onClose?: () => void;
}

const COMMON_INCOMING_FLIGHTS = [
  { code: 'KR701', origin: 'ចិន (China)' },
  { code: 'QD602', origin: 'ចិន (China)' },
  { code: 'CA431', origin: 'ចិន (China)' },
  { code: 'CZ323', origin: 'ចិន (China)' },
  { code: 'MU513', origin: 'ចិន (China)' },
  { code: 'VN852', origin: 'វៀតណាម (Vietnam)' },
  { code: 'TG584', origin: 'ថៃ (Thailand)' },
  { code: 'KE690', origin: 'កូរ៉េខាងត្បូង (South Korea)' },
  { code: 'SQ158', origin: 'សិង្ហបុរី (Singapore)' },
  { code: 'AK538', origin: 'ម៉ាឡេស៊ី (Malaysia)' },
  { code: 'MH754', origin: 'ម៉ាឡេស៊ី (Malaysia)' },
  { code: 'PG931', origin: 'ថៃ (Thailand)' },
  { code: 'FD610', origin: 'ថៃ (Thailand)' },
  { code: 'LQ908', origin: 'ចិន (China)' },
];

const COMMON_OUTGOING_FLIGHTS = [
  { code: 'KR702', dest: 'ចិន (China)' },
  { code: 'QD603', dest: 'ចិន (China)' },
  { code: 'CA432', dest: 'ចិន (China)' },
  { code: 'CZ324', dest: 'ចិន (China)' },
  { code: 'MU514', dest: 'ចិន (China)' },
  { code: 'VN853', dest: 'វៀតណាម (Vietnam)' },
  { code: 'TG585', dest: 'ថៃ (Thailand)' },
  { code: 'KE691', dest: 'កូរ៉េខាងត្បូង (South Korea)' },
  { code: 'SQ159', dest: 'សិង្ហបុរី (Singapore)' },
  { code: 'AK539', dest: 'ម៉ាឡេស៊ី (Malaysia)' },
  { code: 'MH755', dest: 'ម៉ាឡេស៊ី (Malaysia)' },
  { code: 'PG932', dest: 'ថៃ (Thailand)' },
  { code: 'FD611', dest: 'ថៃ (Thailand)' },
  { code: 'LQ909', dest: 'ចិន (China)' },
];

const BORDER_COUNTRIES = ['វៀតណាម (Vietnam)', 'ថៃ (Thailand)', 'ឡាវ (Laos)'];

const PASSPORT_NATIONALITIES = [
  { code: 'CHN', nameKh: 'ចិន', nameEn: 'Chinese' },
  { code: 'VNM', nameKh: 'វៀតណាម', nameEn: 'Vietnamese' },
  { code: 'THA', nameKh: 'ថៃ', nameEn: 'Thai' },
  { code: 'LAO', nameKh: 'ឡាវ', nameEn: 'Laotian' },
  { code: 'USA', nameKh: 'អាមេរិក', nameEn: 'American' },
  { code: 'FRA', nameKh: 'បារាំង', nameEn: 'French' },
  { code: 'GBR', nameKh: 'អង់គ្លេស', nameEn: 'British' },
  { code: 'JPN', nameKh: 'ជប៉ុន', nameEn: 'Japanese' },
  { code: 'KOR', nameKh: 'កូរ៉េខាងត្បូង', nameEn: 'Korean' },
  { code: 'SGP', nameKh: 'សិង្ហបុរី', nameEn: 'Singaporean' },
  { code: 'MYS', nameKh: 'ម៉ាឡេស៊ី', nameEn: 'Malaysian' },
  { code: 'IND', nameKh: 'ឥណ្ឌា', nameEn: 'Indian' },
  { code: 'RUS', nameKh: 'រុស្ស៊ី', nameEn: 'Russian' },
];

const PREDEFINED_REASONS: string[] = [
  'លិខិតឆ្លងដែន ឬឯកសារធ្វើដំណើរគ្មានតម្លៃ (១)',
  'គ្មានទិដ្ឋាការចូល (២)',
  'ក្លែងបន្លំ (៣)',
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

export const RefusalDeportationManager: React.FC<RefusalDeportationManagerProps> = ({
  currentRole,
  assignedTeam,
  userName,
  categories,
  records: externalRecords,
  onUpdateRecords,
  onClose,
}) => {
  const isSecondary = currentRole === 'Secondary' || currentRole === 'Admin' || currentRole === 'User (ការិយាល័យ)';

  // Load / Save state of records from localStorage or external prop
  const [internalRecords, setInternalRecords] = useState<RefusalDeportationRecord[]>(() => {
    if (externalRecords) return externalRecords;
    try {
      const saved = localStorage.getItem('app_refusal_deportation_records_v1');
      return saved ? JSON.parse(saved) : [];
    } catch {
      return [];
    }
  });

  const records = externalRecords || internalRecords;
  const setRecords = (updater: React.SetStateAction<RefusalDeportationRecord[]>) => {
    const next = typeof updater === 'function' ? updater(records) : updater;
    if (!externalRecords) {
      setInternalRecords(next);
    }
    if (onUpdateRecords) {
      onUpdateRecords(next);
    }
    try {
      localStorage.setItem('app_refusal_deportation_records_v1', JSON.stringify(next));
    } catch {}
  };

  // Determine Gate Type automatically for logged-in team to filter/default
  const autoGateType = useMemo(() => {
    if (!assignedTeam) return 'airport';
    const norm = normalizeTeamName(assignedTeam);
    if (/អាកាស|airport|តេជោ|សៀមរាប|ព្រះសីហនុ/i.test(norm)) {
      return 'airport';
    }
    return 'border';
  }, [assignedTeam]);

  // UI state
  const [viewMode, setViewMode] = useState<'list' | 'add' | 'edit' | 'report'>('list');
  const [selectedRecordId, setSelectedRecordId] = useState<string | null>(null);
  const [expandedRecordIds, setExpandedRecordIds] = useState<Record<string, boolean>>({});

  // Dynamic reasons from Category Manager
  const dynamicReasons = useMemo(() => {
    if (categories?.refusalReasons && categories.refusalReasons.length > 0) {
      return categories.refusalReasons.map((r) => r.name);
    }
    return PREDEFINED_REASONS;
  }, [categories?.refusalReasons]);

  // Form states
  const [formGateType, setFormGateType] = useState<'airport' | 'border'>(autoGateType);
  const [formDate, setFormDate] = useState<string>(new Date().toISOString().split('T')[0]);
  const [formTeamName, setFormTeamName] = useState<string>(assignedTeam || 'អាកាស តេជោ');

  // List of people under this case (reusing VisaForm paradigm)
  const [people, setPeople] = useState<RefusalDeportationPerson[]>([
    {
      id: '1',
      fullName: '',
      gender: '',
      dob: '',
      nationality: 'CHN',
      passportNumber: '',
      reason: '',
    },
  ]);

  // Airport-specific states
  const [formIncomingFlight, setFormIncomingFlight] = useState<string>('');
  const [formOriginCountry, setFormOriginCountry] = useState<string>('');
  const [formReturnFlight, setFormReturnFlight] = useState<string>('');
  const [formDestinationCountry, setFormDestinationCountry] = useState<string>('');

  // Border-specific states
  const [formComingFromBorderCountry, setFormComingFromBorderCountry] = useState<string>('វៀតណាម (Vietnam)');
  const [formReturningToBorderCountry, setFormReturningToBorderCountry] = useState<string>('វៀតណាម (Vietnam)');

  // Filter states for list view
  const [filterTeam, setFilterTeam] = useState<string>('ALL');
  const [filterGateType, setFilterGateType] = useState<string>('ALL');
  const [filterSearch, setFilterSearch] = useState<string>('');
  const [filterStartDate, setFilterStartDate] = useState<string>('');
  const [filterEndDate, setFilterEndDate] = useState<string>('');

  // Success toast
  const [toast, setToast] = useState<{ message: string; type: 'success' | 'error' } | null>(null);
  const triggerToast = (msg: string, type: 'success' | 'error' = 'success') => {
    setToast({ message: msg, type });
    setTimeout(() => setToast(null), 3000);
  };

  // Custom Delete Confirmation Modal state (Replaces window.confirm to avoid iframe sandbox blocks)
  const [deletePending, setDeletePending] = useState<{
    type: 'case' | 'person';
    recordId: string;
    personId?: string;
    personIndex?: number;
    title: string;
  } | null>(null);

  // Sync Form Gate Type when autoGateType changes
  useEffect(() => {
    if (!isSecondary) {
      setFormGateType(autoGateType);
    }
  }, [autoGateType, isSecondary]);

  // Handle Airport Flight Number Auto-Complete & Country Auto-Detection
  const handleIncomingFlightChange = (val: string) => {
    setFormIncomingFlight(val);
    const cleaned = val.trim().toUpperCase();
    const matched = COMMON_INCOMING_FLIGHTS.find((f) => f.code === cleaned);
    if (matched) {
      setFormOriginCountry(matched.origin);
    } else if (cleaned.startsWith('CZ') || cleaned.startsWith('CA') || cleaned.startsWith('MU') || cleaned.startsWith('MF') || cleaned.startsWith('ZH') || cleaned.startsWith('HU') || cleaned.startsWith('FM') || cleaned.startsWith('3U') || cleaned.startsWith('KR')) {
      setFormOriginCountry('ចិន (China)');
    } else if (cleaned.startsWith('VN') || cleaned.startsWith('BL') || cleaned.startsWith('QH')) {
      setFormOriginCountry('វៀតណាម (Vietnam)');
    } else if (cleaned.startsWith('TG') || cleaned.startsWith('FD') || cleaned.startsWith('SL') || cleaned.startsWith('WE') || cleaned.startsWith('PG')) {
      setFormOriginCountry('ថៃ (Thailand)');
    } else if (cleaned.startsWith('KE') || cleaned.startsWith('OZ') || cleaned.startsWith('LJ') || cleaned.startsWith('7C') || cleaned.startsWith('TW')) {
      setFormOriginCountry('កូរ៉េខាងត្បូង (South Korea)');
    } else if (cleaned.startsWith('SQ') || cleaned.startsWith('TR')) {
      setFormOriginCountry('សិង្ហបុរី (Singapore)');
    } else if (cleaned.startsWith('MH') || cleaned.startsWith('AK') || cleaned.startsWith('OD')) {
      setFormOriginCountry('ម៉ាឡេស៊ី (Malaysia)');
    } else if (cleaned.startsWith('JL') || cleaned.startsWith('NH')) {
      setFormOriginCountry('ជប៉ុន (Japan)');
    } else if (cleaned.startsWith('EK') || cleaned.startsWith('FZ')) {
      setFormOriginCountry('អារ៉ាប់រួម (UAE)');
    } else if (cleaned.startsWith('QR')) {
      setFormOriginCountry('កាតា (Qatar)');
    } else if (cleaned.length >= 2) {
      setFormOriginCountry('អន្តរជាតិ (International)');
    }
  };

  const handleReturnFlightChange = (val: string) => {
    setFormReturnFlight(val);
    const cleaned = val.trim().toUpperCase();
    const matched = COMMON_OUTGOING_FLIGHTS.find((f) => f.code === cleaned);
    if (matched) {
      setFormDestinationCountry(matched.dest);
    } else if (cleaned.startsWith('CZ') || cleaned.startsWith('CA') || cleaned.startsWith('MU') || cleaned.startsWith('MF') || cleaned.startsWith('ZH') || cleaned.startsWith('HU') || cleaned.startsWith('FM') || cleaned.startsWith('3U') || cleaned.startsWith('KR')) {
      setFormDestinationCountry('ចិន (China)');
    } else if (cleaned.startsWith('VN') || cleaned.startsWith('BL') || cleaned.startsWith('QH')) {
      setFormDestinationCountry('វៀតណាម (Vietnam)');
    } else if (cleaned.startsWith('TG') || cleaned.startsWith('FD') || cleaned.startsWith('SL') || cleaned.startsWith('WE') || cleaned.startsWith('PG')) {
      setFormDestinationCountry('ថៃ (Thailand)');
    } else if (cleaned.startsWith('KE') || cleaned.startsWith('OZ') || cleaned.startsWith('LJ') || cleaned.startsWith('7C') || cleaned.startsWith('TW')) {
      setFormDestinationCountry('កូរ៉េខាងត្បូង (South Korea)');
    } else if (cleaned.startsWith('SQ') || cleaned.startsWith('TR')) {
      setFormDestinationCountry('សិង្ហបុរី (Singapore)');
    } else if (cleaned.startsWith('MH') || cleaned.startsWith('AK') || cleaned.startsWith('OD')) {
      setFormDestinationCountry('ម៉ាឡេស៊ី (Malaysia)');
    } else if (cleaned.startsWith('JL') || cleaned.startsWith('NH')) {
      setFormDestinationCountry('ជប៉ុន (Japan)');
    } else if (cleaned.startsWith('EK') || cleaned.startsWith('FZ')) {
      setFormDestinationCountry('អារ៉ាប់រួម (UAE)');
    } else if (cleaned.startsWith('QR')) {
      setFormDestinationCountry('កាតា (Qatar)');
    } else if (cleaned.length >= 2) {
      setFormDestinationCountry('អន្តរជាតិ (International)');
    }
  };

  // Handle Passport Number auto-detection of Nationality
  const handlePassportChange = (personId: string, passportVal: string) => {
    const upper = passportVal.trim().toUpperCase();
    setPeople((prev) =>
      prev.map((p) => {
        if (p.id !== personId) return p;
        let nat = p.nationality;
        if (upper.startsWith('E') || upper.startsWith('G') || upper.startsWith('S') || upper.startsWith('D') || upper.startsWith('CHN')) {
          nat = 'CHN';
        } else if (upper.startsWith('B') || upper.startsWith('VNM')) {
          nat = 'VNM';
        } else if (upper.startsWith('A') || upper.startsWith('THA')) {
          nat = 'THA';
        } else if (upper.startsWith('L') || upper.startsWith('LAO')) {
          nat = 'LAO';
        } else if (upper.startsWith('US') || upper.startsWith('USA')) {
          nat = 'USA';
        } else if (upper.startsWith('FR') || upper.startsWith('FRA')) {
          nat = 'FRA';
        } else if (upper.startsWith('JP') || upper.startsWith('JPN')) {
          nat = 'JPN';
        } else if (upper.startsWith('KR') || upper.startsWith('KOR')) {
          nat = 'KOR';
        }
        return { ...p, passportNumber: upper, nationality: nat };
      })
    );
  };

  // Dynamic row additions like VisaForm (Auto-inheriting reason for refusal from previous person)
  const handleAddPerson = () => {
    setPeople((prev) => {
      const lastPerson = prev.length > 0 ? prev[prev.length - 1] : null;
      const inheritedReason = lastPerson ? lastPerson.reason : '';
      const inheritedNationality = lastPerson ? lastPerson.nationality : 'CHN';
      const inheritedGender = lastPerson ? lastPerson.gender : '';

      return [
        ...prev,
        {
          id: Date.now().toString(),
          fullName: '',
          gender: inheritedGender,
          dob: '',
          nationality: inheritedNationality,
          passportNumber: '',
          reason: inheritedReason,
        },
      ];
    });
  };

  const handleRemovePerson = (id: string) => {
    if (people.length === 1) {
      triggerToast('ត្រូវតែមានយ៉ាងហោចណាស់ជនបរទេសម្នាក់ក្នុងបញ្ជី!', 'error');
      return;
    }
    setPeople((prev) => prev.filter((p) => p.id !== id));
  };

  const updatePerson = (id: string, field: keyof RefusalDeportationPerson, value: any) => {
    setPeople((prev) =>
      prev.map((p) => (p.id === id ? { ...p, [field]: value } : p))
    );
  };

  // Check if a record can be edited or deleted (Full permission for both Team and Office)
  const canModifyRecord = (_rec: RefusalDeportationRecord): boolean => {
    return true; // Always allow edit and delete for both Team and Office
  };

  // Toggle Office allowance for team to edit/delete
  const toggleAllowByOffice = (id: string) => {
    setRecords((prev) =>
      prev.map((r) => {
        if (r.id === id) {
          const nextState = !r.allowedByOffice;
          if (nextState) {
            triggerToast('បានបើកសោរអនុញ្ញាតឱ្យក្រុមអាចកែប្រែ/លុបករណីនេះបាន!', 'success');
          } else {
            triggerToast('បានបិទការអនុញ្ញាតកែប្រែ/លុបករណីនេះ!', 'error');
          }
          return { ...r, allowedByOffice: nextState };
        }
        return r;
      })
    );
  };

  // Open respective form
  const handleOpenAddForm = (type: 'airport' | 'border') => {
    resetForm();
    setFormGateType(type);
    setViewMode('add');
  };

  // Pre-fill form when editing
  const startEdit = (rec: RefusalDeportationRecord) => {
    if (!canModifyRecord(rec)) {
      triggerToast('មិនអាចកែប្រែបានទេ! ដោយសារផុតកំណត់ ៣០នាទី ក្រោយបញ្ចូល (ត្រូវមានការអនុញ្ញាតពីការិយាល័យ)', 'error');
      return;
    }

    setSelectedRecordId(rec.id);
    setFormGateType(rec.gateType);
    setFormDate(rec.date);
    setFormTeamName(rec.teamName || assignedTeam || 'អាកាស តេជោ');
    setPeople(rec.people && rec.people.length > 0 ? rec.people : [
      { id: '1', fullName: '', gender: '', dob: '', nationality: 'CHN', passportNumber: '', reason: '' }
    ]);

    setFormIncomingFlight(rec.incomingFlight || '');
    setFormOriginCountry(rec.originCountry || '');
    setFormReturnFlight(rec.returnFlight || '');
    setFormDestinationCountry(rec.destinationCountry || '');

    setFormComingFromBorderCountry(rec.comingFromBorderCountry || 'វៀតណាម (Vietnam)');
    setFormReturningToBorderCountry(rec.returningToBorderCountry || 'វៀតណាម (Vietnam)');

    setViewMode('edit');
  };

  // Reset form
  const resetForm = () => {
    setSelectedRecordId(null);
    setFormTeamName(assignedTeam || 'អាកាស តេជោ');
    setFormIncomingFlight('');
    setFormOriginCountry('');
    setFormReturnFlight('');
    setFormDestinationCountry('');
    setFormComingFromBorderCountry('វៀតណាម (Vietnam)');
    setFormReturningToBorderCountry('វៀតណាម (Vietnam)');
    setPeople([
      {
        id: '1',
        fullName: '',
        gender: '',
        dob: '',
        nationality: 'CHN',
        passportNumber: '',
        reason: '',
      },
    ]);
  };

  // Submit Form (Add / Edit)
  const handleSubmit = (e: React.FormEvent) => {
    e.preventDefault();

    // Check if at least passport and name are filled for all people
    const hasInvalid = people.some(
      (p) => !p.fullName.trim() || !p.passportNumber.trim()
    );

    if (hasInvalid) {
      triggerToast('សូមបំពេញឈ្មោះ និងលេខលិខិតឆ្លងដែនសម្រាប់ជនបរទេសគ្រប់រូប!', 'error');
      return;
    }

    const payload: RefusalDeportationRecord = {
      id: selectedRecordId || `rd-${Date.now()}`,
      gateType: formGateType,
      teamName: isSecondary ? formTeamName : (assignedTeam || 'ក្រុមការងារ'),
      date: formDate,
      people: people.map((p) => ({
        ...p,
        fullName: p.fullName.trim().toUpperCase(),
        passportNumber: p.passportNumber.trim().toUpperCase(),
      })),
      createdAt: selectedRecordId ? records.find(r => r.id === selectedRecordId)?.createdAt || new Date().toISOString() : new Date().toISOString(),
      createdBy: userName,
      ...(formGateType === 'airport'
        ? {
            incomingFlight: formIncomingFlight,
            originCountry: formOriginCountry,
            returnFlight: formReturnFlight,
            destinationCountry: formDestinationCountry,
          }
        : {
            comingFromBorderCountry: formComingFromBorderCountry,
            returningToBorderCountry: formReturningToBorderCountry,
          }),
    };

    if (viewMode === 'edit') {
      setRecords((prev) => prev.map((r) => (r.id === selectedRecordId ? payload : r)));
      triggerToast('កែប្រែទិន្នន័យជោគជ័យ!');
    } else {
      setRecords((prev) => [payload, ...prev]);
      triggerToast('បញ្ចូលទិន្នន័យថ្មីជោគជ័យ!');
    }

    resetForm();
    setViewMode('list');
  };

  // Open delete case confirmation
  const handleDelete = (id: string) => {
    setDeletePending({
      type: 'case',
      recordId: id,
      title: 'តើលោកអ្នកពិតជាចង់លុបករណីនេះទាំងមូល (រួមទាំងជនបរទេសទាំងអស់ក្នុងករណីនេះ) ចេញពីបញ្ជីមែនទេ?',
    });
  };

  // Open delete individual person confirmation
  const handleDeletePerson = (recId: string, personId: string, personName?: string, index?: number) => {
    const displayName = personName ? `«${personName}»` : 'ជនបរទេសនេះ';
    setDeletePending({
      type: 'person',
      recordId: recId,
      personId: personId,
      personIndex: index,
      title: `តើលោកអ្នកពិតជាចង់លុប ${displayName} ចេញពីបញ្ជីមែនទេ?`,
    });
  };

  // Confirm delete action
  const confirmDeleteAction = () => {
    if (!deletePending) return;
    const { type, recordId, personIndex } = deletePending;

    if (type === 'case') {
      setRecords((prev) => prev.filter((r) => r.id !== recordId));
      triggerToast('បានលុបករណីនេះចេញពីបញ្ជីជោគជ័យ!', 'success');
    } else if (type === 'person') {
      const rec = records.find((r) => r.id === recordId);
      if (rec) {
        if (rec.people && rec.people.length > 1) {
          setRecords((prev) =>
            prev.map((r) => {
              if (r.id === recordId) {
                return {
                  ...r,
                  people: r.people.filter((_, pIdx) => pIdx !== personIndex),
                };
              }
              return r;
            })
          );
          triggerToast('បានលុបជនបរទេសចេញពីបញ្ជីជោគជ័យ!', 'success');
        } else {
          // If only 1 person in the case, delete the whole case
          setRecords((prev) => prev.filter((r) => r.id !== recordId));
          triggerToast('បានលុបករណីនេះចេញពីបញ្ជីជោគជ័យ!', 'success');
        }
      }
    }
    setDeletePending(null);
  };

  // Pre-populate mock database with some realistic records if empty
  const handleSeedMockData = () => {
    if (records.length > 0) return;
    const mock: RefusalDeportationRecord[] = [
      {
        id: 'rd-mock-1',
        gateType: 'airport',
        teamName: 'អាកាស សៀមរាប',
        date: '2026-09-24',
        people: [
          {
            id: 'm1-p1',
            caseNumber: '១',
            fullName: 'WANG WEI',
            gender: 'ប្រុស',
            dob: '1989-05-12',
            nationality: 'CHN',
            passportNumber: 'E98124552',
            reason: 'ករណីទី៣៖ គ្មានទិដ្ឋាការ ឬទិដ្ឋាការមិនត្រឹមត្រូវ',
          },
          {
            id: 'm1-p2',
            caseNumber: '២',
            fullName: 'ZHANG MIN',
            gender: 'ស្រី',
            dob: '1992-08-25',
            nationality: 'CHN',
            passportNumber: 'E44238511',
            reason: 'ករណីទី៣៖ គ្មានទិដ្ឋាការ ឬទិដ្ឋាការមិនត្រឹមត្រូវ',
          }
        ],
        incomingFlight: 'KR701',
        originCountry: 'ចិន (China)',
        returnFlight: 'KR702',
        destinationCountry: 'ចិន (China)',
        createdAt: new Date().toISOString(),
        createdBy: 'អ្នកធ្វើតារាង',
      },
      {
        id: 'rd-mock-2',
        gateType: 'airport',
        teamName: 'អាកាស តេជោ',
        date: '2026-09-25',
        people: [
          {
            id: 'm2-p1',
            caseNumber: '១',
            fullName: 'LE MINH',
            gender: 'ប្រុស',
            dob: '1995-11-20',
            nationality: 'VNM',
            passportNumber: 'C12948123',
            reason: 'ករណីទី១០៖ គ្មានឯកសារបញ្ជាក់ពីគោលបំណងពិតប្រាកដនៃការចូលប្រទេស',
          }
        ],
        incomingFlight: 'VN852',
        originCountry: 'វៀតណាម (Vietnam)',
        returnFlight: 'VN853',
        destinationCountry: 'វៀតណាម (Vietnam)',
        createdAt: new Date().toISOString(),
        createdBy: 'អ្នកធ្វើតាង',
      },
      {
        id: 'rd-mock-3',
        gateType: 'border',
        teamName: 'ព្រំដែន ប៉ោយប៉ែត',
        date: '2026-09-23',
        people: [
          {
            id: 'm3-p1',
            caseNumber: '១',
            fullName: 'SOMCHAI JAIYEN',
            gender: 'ប្រុស',
            dob: '1985-02-15',
            nationality: 'THA',
            passportNumber: 'AA7712398',
            reason: 'ករណីទី១៤ (ន៧)៖ នាយកដ្ឋានស៊ើបអង្កេត និងអនុវត្តនីតិវិធី (ន៧)',
          }
        ],
        comingFromBorderCountry: 'ថៃ (Thailand)',
        returningToBorderCountry: 'ថៃ (Thailand)',
        createdAt: new Date().toISOString(),
        createdBy: 'អ្នកធ្វើតាង',
      },
    ];
    setRecords(mock);
    triggerToast('បានបញ្ចូលទិន្នន័យគំរូជោគជ័យ!');
  };

  // Filtered records based on filter options and user role
  const filteredRecords = useMemo(() => {
    return records.filter((r) => {
      // 1. Team filter
      if (!isSecondary) {
        // Team user can only see records for their own team
        const normAssigned = normalizeTeamName(assignedTeam || '');
        const normTeam = normalizeTeamName(r.teamName);
        if (normTeam !== normAssigned) return false;
      } else {
        // Office user can filter by any team
        if (filterTeam !== 'ALL') {
          const normFilter = normalizeTeamName(filterTeam);
          const normTeam = normalizeTeamName(r.teamName);
          if (normTeam !== normFilter) return false;
        }
      }

      // 2. Gate type filter
      if (filterGateType !== 'ALL') {
        if (r.gateType !== filterGateType) return false;
      }

      // 3. Search query (passport, name, nationality inside list of people, or reason)
      if (filterSearch.trim()) {
        const query = filterSearch.trim().toLowerCase();
        const hasPeopleMatch = r.people && r.people.some((p) => {
          return (
            p.fullName.toLowerCase().includes(query) ||
            p.passportNumber.toLowerCase().includes(query) ||
            p.nationality.toLowerCase().includes(query) ||
            p.reason.toLowerCase().includes(query) ||
            p.caseNumber.includes(query)
          );
        });
        if (!hasPeopleMatch) return false;
      }

      // 4. Date range filter
      if (filterStartDate) {
        if (r.date < filterStartDate) return false;
      }
      if (filterEndDate) {
        if (r.date > filterEndDate) return false;
      }

      return true;
    });
  }, [records, isSecondary, assignedTeam, filterTeam, filterGateType, filterSearch, filterStartDate, filterEndDate]);

  // List of active teams for office filter
  const activeTeamsList = useMemo(() => {
    const list = new Set<string>();
    records.forEach((r) => {
      if (r.teamName) list.add(r.teamName);
    });
    return Array.from(list);
  }, [records]);

  // List of teams for office form dropdown
  const formTeamsList = useMemo(() => {
    const fromCat = (categories?.visaTeamsRobok || categories?.visaTeams || []).map((t) => t.name);
    if (fromCat.length > 0) return Array.from(new Set(fromCat));
    return [
      'អាកាស តេជោ', 'អាកាស សៀមរាប', 'អាកាស ព្រះសីហនុ', 'ព្រំដែន ប៉ោយប៉ែត', 'ព្រំដែន បាវិត', 'ព្រំដែន ចាំយាម', 'ព្រំដែន ដូង', 'ព្រំដែន អូរស្មាច់', 'ព្រំដែន ព្រំ', 'ព្រំដែន បន្ទាយចក្រី', 'ព្រំដែន ជាំ', 'ព្រំដែន ត្រពាំងស្រែ', 'ព្រំដែន ត្រពាំងក្រៀល', 'ព្រំដែន ត្រពាំងផ្លុង', 'ព្រំដែន ភ្នំដិន', 'ព្រំដែន កោះរកា', 'ព្រំដែន ព្រៃវល្លិ៍', 'ព្រំដែន អូរយ៉ាដាវ', 'ព្រំដែន ក្អមសំណ', 'ព្រំដែន ភ្នំដី', 'កំពង់ផែ ឧកញ៉ាម៉ុង', 'កំពង់ផែ ស្ទឹងហាវ', 'កំពង់ផែ ព្រះសីហនុ', 'កំពង់ផែ ភ្នំពេញ', 'ព្រំដែន ព្រែកចាក'
    ];
  }, [categories]);

  // Toggle expand / collapse of foreigner table in list
  const toggleExpandRecord = (id: string) => {
    setExpandedRecordIds((prev) => ({
      ...prev,
      [id]: !prev[id],
    }));
  };

  // Generate A4 report print
  const handlePrint = () => {
    printA4Document('refusal-report-print-container');
  };

  // Export to PDF (A4 Layout)
  const handleExportPdf = () => {
    const el = document.getElementById('refusal-report-print-container');
    if (el) {
      exportElementToPdf(el, `Refusal_and_Deportation_Report_${new Date().toISOString().split('T')[0]}`);
    } else {
      triggerToast('មិនអាចទាញយក PDF បានទេ!', 'error');
    }
  };

  return (
    <div className="w-full h-full bg-[#f4f6f9] text-[#1F2937] flex flex-col focus:outline-none relative">
      {/* Toast Notification */}
      {toast && (
        <div className={`fixed top-4 right-4 z-50 px-4 py-3 rounded-lg shadow-lg flex items-center gap-2 text-white text-xs font-bold font-sans transition-all transform duration-300 translate-y-0 ${
          toast.type === 'success' ? 'bg-emerald-600' : 'bg-rose-600'
        }`}>
          <CheckCircle2 className="w-4 h-4 shrink-0" />
          <span>{toast.message}</span>
        </div>
      )}

      {/* Control Header */}
      <div className="bg-white border-b border-gray-200/80 px-4 sm:px-6 py-3 shrink-0 flex flex-wrap items-center justify-between gap-3 shadow-2xs">
        <div className="flex items-center gap-2.5">
          <div className="w-9 h-9 bg-orange-100 rounded-lg flex items-center justify-center text-orange-600">
            <ShieldAlert className="w-5 h-5" />
          </div>
          <div>
            <h1 className="text-base sm:text-lg font-black text-blue-900 font-sans tracking-wide">
              ការងារបដិសេធនិងបញ្ជូនចេញ (Refusal & Deportation)
            </h1>
            <p className="text-[11px] text-gray-500 font-medium">
              {!isSecondary 
                ? `ក្រុមផ្តល់ទិដ្ឋាការ៖ ${assignedTeam || 'អាកាស តេជោ'}` 
                : 'ទិន្នន័យគ្រប់គ្រងដោយ៖ ការិយាល័យផ្តល់ទិដ្ឋាការ'}
            </p>
          </div>
        </div>

        <div className="flex items-center gap-2">
          {viewMode === 'list' ? (
            <>
              {/* Form Option 1: Airways Gate Form */}
              {(!isSecondary && autoGateType === 'border') ? null : (
                <button
                  onClick={() => handleOpenAddForm('airport')}
                  className="h-9 px-3.5 bg-blue-600 hover:bg-blue-700 text-white rounded font-sans text-xs font-bold flex items-center gap-1.5 transition-all shadow-xs cursor-pointer"
                >
                  <Plane className="w-4 h-4" />
                  <span>ទម្រង់ផ្លូវអាកាស (Airways Form)</span>
                </button>
              )}

              {/* Form Option 2: Border Gate Form */}
              {(!isSecondary && autoGateType === 'airport') ? null : (
                <button
                  onClick={() => handleOpenAddForm('border')}
                  className="h-9 px-3.5 bg-amber-600 hover:bg-amber-700 text-white rounded font-sans text-xs font-bold flex items-center gap-1.5 transition-all shadow-xs cursor-pointer"
                >
                  <Globe className="w-4 h-4" />
                  <span>ទម្រង់ព្រំដែនគោក (Border Form)</span>
                </button>
              )}

              <button
                onClick={() => setViewMode('report')}
                className="h-9 px-3.5 bg-orange-500 hover:bg-orange-600 text-white rounded font-sans text-xs font-bold flex items-center gap-1.5 transition-all shadow-xs cursor-pointer"
              >
                <FileText className="w-4 h-4" />
                <span>របាយការណ៍សង្ខេប</span>
              </button>
              {records.length === 0 && (
                <button
                  onClick={handleSeedMockData}
                  className="h-9 px-3 border border-dashed border-blue-400 text-blue-600 hover:bg-blue-50/50 rounded font-sans text-xs font-bold transition-all cursor-pointer"
                >
                  បញ្ចូលទិន្នន័យគំរូ (Seed Mock)
                </button>
              )}
            </>
          ) : (
            <button
              onClick={() => {
                resetForm();
                setViewMode('list');
              }}
              className="h-9 px-4 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded font-sans text-xs font-bold flex items-center gap-1.5 transition-all cursor-pointer"
            >
              <ArrowLeft className="w-4 h-4" />
              <span>ត្រឡប់ក្រោយ (Back)</span>
            </button>
          )}

          {onClose && (
            <button
              onClick={onClose}
              className="w-9 h-9 flex items-center justify-center border border-gray-200 rounded hover:bg-gray-50 text-gray-400 hover:text-gray-600 transition cursor-pointer"
            >
              <X className="w-5 h-5" />
            </button>
          )}
        </div>
      </div>

      {/* Main Content Area */}
      <div className="flex-1 overflow-y-auto p-4 sm:p-5">
        {viewMode === 'list' && (
          <div className="space-y-4">
            {/* Filter Bar */}
            <div className="bg-white border border-gray-200/90 rounded-lg p-3 sm:p-4 shadow-3xs flex flex-wrap items-end gap-3 text-xs">
              {isSecondary && (
                <div className="w-full sm:w-48">
                  <label className="block font-bold text-gray-700 mb-1">តម្រងតាមក្រុម (Filter Team):</label>
                  <select
                    value={filterTeam}
                    onChange={(e) => setFilterTeam(e.target.value)}
                    className="w-full h-9 border border-gray-300 rounded px-2.5 font-sans focus:border-blue-500 focus:outline-none font-bold bg-white text-blue-900"
                  >
                    <option value="ALL">-- គ្រប់ក្រុមទាំងអស់ --</option>
                    {activeTeamsList.map((t) => (
                      <option key={t} value={t}>
                        {t}
                      </option>
                    ))}
                  </select>
                </div>
              )}

              {isSecondary && (
                <div className="w-full sm:w-36">
                  <label className="block font-bold text-gray-700 mb-1">ប្រភេទច្រកទ្វារ (Gate Type):</label>
                  <select
                    value={filterGateType}
                    onChange={(e) => setFilterGateType(e.target.value)}
                    className="w-full h-9 border border-gray-300 rounded px-2.5 font-sans focus:border-blue-500 focus:outline-none font-medium bg-white text-xs"
                  >
                    <option value="ALL">-- គ្រប់ប្រភេទ --</option>
                    <option value="airport">ផ្លូវអាកាស (Airport)</option>
                    <option value="border">ព្រំដែន (Land Border)</option>
                  </select>
                </div>
              )}

              <div className="flex-1 min-w-[200px]">
                <label className="block font-bold text-gray-700 mb-1">ស្វែងរក (លិខិតឆ្លងដែន, ឈ្មោះ, សញ្ជាតិ...):</label>
                <div className="relative">
                  <input
                    type="text"
                    value={filterSearch}
                    onChange={(e) => setFilterSearch(e.target.value)}
                    placeholder="ស្វែងរកលេខ Passport ឬឈ្មោះជនបរទេស..."
                    className="w-full h-9 border border-gray-300 rounded pl-8 pr-3 font-sans focus:border-blue-500 focus:outline-none"
                  />
                  <Search className="w-4 h-4 text-gray-400 absolute left-2.5 top-2.5" />
                </div>
              </div>

              <div className="w-full sm:w-36">
                <label className="block font-bold text-gray-700 mb-1">ចាប់ពីថ្ងៃ (From):</label>
                <CustomDatePicker
                  value={filterStartDate}
                  onChange={setFilterStartDate}
                  placeholder="YYYY-MM-DD"
                  className="h-9"
                />
              </div>

              <div className="w-full sm:w-36">
                <label className="block font-bold text-gray-700 mb-1">ដល់ថ្ងៃ (To):</label>
                <CustomDatePicker
                  value={filterEndDate}
                  onChange={setFilterEndDate}
                  placeholder="YYYY-MM-DD"
                  className="h-9"
                />
              </div>

              <button
                onClick={() => {
                  setFilterTeam('ALL');
                  setFilterGateType('ALL');
                  setFilterSearch('');
                  setFilterStartDate('');
                  setFilterEndDate('');
                }}
                className="h-9 px-3 border border-gray-300 hover:bg-gray-50 rounded font-sans text-gray-600 transition"
              >
                សម្អាត (Clear)
              </button>
            </div>

            {/* Stats calculation for header banner */}
            {(() => {
              const totalCases = filteredRecords.length;
              let totalPeople = 0;
              const nats = new Set<string>();

              filteredRecords.forEach((r) => {
                if (r.people && r.people.length > 0) {
                  totalPeople += r.people.length;
                  r.people.forEach((p) => {
                    if (p.nationality) nats.add(p.nationality.trim().toUpperCase());
                  });
                }
              });

              return (
                <div className="bg-white border border-gray-200/90 rounded-lg shadow-2xs overflow-hidden">
                  <div className="p-3 bg-blue-50/50 border-b border-gray-200/90 flex justify-between items-center">
                    <span className="text-xs font-black text-blue-900 font-sans">
                      បញ្ជីឈ្មោះជនបរទេសដែលត្រូវបានបដិសេធក្នុងករណី = {toKhmerNum(totalCases)} ករណី សរុបជនបរទេស ៖ {toKhmerNum(totalPeople)} នាក់, សញ្ជាតិ៖ {toKhmerNum(nats.size)} សញ្ជាតិ
                    </span>
                  </div>

                  <div className="overflow-x-auto">
                    <table className="w-full text-left border-collapse text-[11px] font-sans">
                      <thead>
                        <tr className="bg-gray-100 text-gray-800 font-bold border-b border-gray-200 text-center">
                          <th className="py-2.5 px-2 border-r border-gray-200 w-12">ករណី</th>
                          <th className="py-2.5 px-2 border-r border-gray-200 w-10">ល.រ</th>
                          <th className="py-2.5 px-2.5 border-r border-gray-200 w-24">កាលបរិច្ឆេទ</th>
                          {isSecondary && (
                            <th className="py-2.5 px-2.5 border-r border-gray-200 min-w-[120px]">ឈ្មោះក្រុម</th>
                          )}
                          <th className="py-2.5 px-2.5 border-r border-gray-200">ជើងយន្តហោះចូល</th>
                          <th className="py-2.5 px-2.5 border-r border-gray-200">មកពី</th>
                          <th className="py-2.5 px-3 border-r border-gray-200 text-left min-w-[140px]">ឈ្មោះពេញ ( FULL NAME )</th>
                          <th className="py-2.5 px-2 border-r border-gray-200 w-12">ភេទ</th>
                          <th className="py-2.5 px-2.5 border-r border-gray-200 w-24">ថ្ងៃខែឆ្នាំកំណើត</th>
                          <th className="py-2.5 px-2 border-r border-gray-200 w-16">សញ្ជាតិ</th>
                          <th className="py-2.5 px-2.5 border-r border-gray-200">លេខលិខិតឆ្លងដែន</th>
                          <th className="py-2.5 px-3 border-r border-gray-200 text-left min-w-[180px]">មូលហេតុបដិសេធ</th>
                          <th className="py-2.5 px-2.5 border-r border-gray-200">ជើងយន្តហោះចេញ</th>
                          <th className="py-2.5 px-2.5 border-r border-gray-200">ត្រឡប់ទៅ</th>
                          <th className="py-2.5 px-2.5 w-16">សកម្មភាព</th>
                        </tr>
                      </thead>
                      <tbody>
                        {filteredRecords.length === 0 ? (
                          <tr>
                            <td colSpan={isSecondary ? 15 : 14} className="py-8 text-center text-gray-400 font-sans">
                              មិនមានកំណត់ត្រាស្របតាមតម្រងដែលបានជ្រើសរើសឡើយ!
                            </td>
                          </tr>
                        ) : (
                          filteredRecords.map((r, caseIdx) => {
                            const peopleList = r.people && r.people.length > 0 ? r.people : [
                              { id: '1', fullName: '---', gender: '---', dob: '---', nationality: '---', passportNumber: '---', reason: '---' }
                            ];

                            return peopleList.map((p, personIdx) => {
                              const isFirstPerson = personIdx === 0;
                              return (
                                <tr key={`${r.id}-${p.id || personIdx}`} className="hover:bg-blue-50/30 border-b border-gray-200 text-[11px]">
                                  {/* ករណី (Case Number) */}
                                  {isFirstPerson ? (
                                    <td rowSpan={peopleList.length} className="py-2 px-2 text-center font-bold text-gray-900 border-r border-gray-200 bg-gray-50/50 align-middle">
                                      {caseIdx + 1}
                                    </td>
                                  ) : null}

                                  {/* ល.រ (Seq Number inside case) */}
                                  <td className="py-2 px-2 text-center font-bold text-gray-700 border-r border-gray-200 align-middle">
                                    {personIdx + 1}
                                  </td>

                                  {/* កាលបរិច្ឆេទ (Date) */}
                                  {isFirstPerson ? (
                                    <td rowSpan={peopleList.length} className="py-2 px-2.5 text-center font-medium text-gray-700 border-r border-gray-200 whitespace-nowrap align-middle">
                                      {r.date}
                                    </td>
                                  ) : null}

                                  {/* ឈ្មោះក្រុម (Team Name) - Office Level Only */}
                                  {isSecondary && isFirstPerson ? (
                                    <td rowSpan={peopleList.length} className="py-2 px-2.5 text-center font-bold text-blue-900 border-r border-gray-200 whitespace-nowrap align-middle bg-blue-50/20">
                                      {r.teamName || '---'}
                                    </td>
                                  ) : null}

                                  {/* ជើងយន្តហោះចូល (Flight In) */}
                                  {isFirstPerson ? (
                                    <td rowSpan={peopleList.length} className="py-2 px-2.5 text-center font-bold text-blue-900 border-r border-gray-200 whitespace-nowrap align-middle">
                                      {r.gateType === 'airport' ? (r.incomingFlight || '---') : '---'}
                                    </td>
                                  ) : null}

                                  {/* មកពី (Origin) */}
                                  {isFirstPerson ? (
                                    <td rowSpan={peopleList.length} className="py-2 px-2.5 text-center font-medium text-gray-800 border-r border-gray-200 align-middle">
                                      {r.gateType === 'airport' ? (r.originCountry || '---') : (r.comingFromBorderCountry || '---')}
                                    </td>
                                  ) : null}

                                  {/* ឈ្មោះពេញ (FULL NAME) */}
                                  <td className="py-2 px-3 font-black text-gray-900 uppercase tracking-wide border-r border-gray-200 align-middle">
                                    {p.fullName || '---'}
                                  </td>

                                  {/* ភេទ (Gender) */}
                                  <td className="py-2 px-2 text-center text-gray-700 border-r border-gray-200 align-middle">
                                    {p.gender || '---'}
                                  </td>

                                  {/* ថ្ងៃខែឆ្នាំកំណើត (DOB) */}
                                  <td className="py-2 px-2.5 text-center text-gray-700 border-r border-gray-200 whitespace-nowrap align-middle">
                                    {p.dob || '---'}
                                  </td>

                                  {/* សញ្ជាតិ (Nationality) */}
                                  <td className="py-2 px-2 text-center font-bold text-blue-800 border-r border-gray-200 align-middle">
                                    {p.nationality || '---'}
                                  </td>

                                  {/* លេខលិខិតឆ្លងដែន (Passport) */}
                                  <td className="py-2 px-2.5 text-center font-bold text-gray-900 tracking-wider border-r border-gray-200 align-middle">
                                    {p.passportNumber || '---'}
                                  </td>

                                  {/* មូលហេតុបដិសេធ (Reason) */}
                                  <td className="py-2 px-3 font-semibold text-red-700 border-r border-gray-200 align-middle">
                                    {p.reason || '---'}
                                  </td>

                                  {/* ជើងយន្តហោះចេញ (Flight Out) */}
                                  {isFirstPerson ? (
                                    <td rowSpan={peopleList.length} className="py-2 px-2.5 text-center font-bold text-orange-900 border-r border-gray-200 whitespace-nowrap align-middle">
                                      {r.gateType === 'airport' ? (r.returnFlight || '---') : '---'}
                                    </td>
                                  ) : null}

                                  {/* ត្រឡប់ទៅ (Destination) */}
                                  {isFirstPerson ? (
                                    <td rowSpan={peopleList.length} className="py-2 px-2.5 text-center font-medium text-gray-800 border-r border-gray-200 align-middle">
                                      {r.gateType === 'airport' ? (r.destinationCountry || '---') : (r.returningToBorderCountry || '---')}
                                    </td>
                                  ) : null}

                                  {/* សកម្មភាព (Actions) */}
                                  {isFirstPerson ? (
                                    <td rowSpan={peopleList.length} className="py-2 px-2.5 text-center align-middle">
                                      <div className="inline-flex gap-1 justify-center items-center">
                                        <button
                                          onClick={() => startEdit(r)}
                                          className="w-6 h-6 flex items-center justify-center bg-blue-50 text-blue-600 hover:bg-blue-100 rounded transition cursor-pointer"
                                          title="កែប្រែករណី"
                                        >
                                          <Edit2 className="w-3 h-3" />
                                        </button>
                                        <button
                                          type="button"
                                          onClick={(e) => {
                                            e.stopPropagation();
                                            handleDelete(r.id);
                                          }}
                                          className="w-6 h-6 flex items-center justify-center bg-red-50 text-red-600 hover:bg-red-100 rounded transition cursor-pointer"
                                          title="លុបករណី"
                                        >
                                          <Trash2 className="w-3.5 h-3.5 text-red-600" />
                                        </button>
                                      </div>
                                    </td>
                                  ) : null}
                                </tr>
                              );
                            });
                          })
                        )}
                      </tbody>
                    </table>
                  </div>
                </div>
              );
            })()}
          </div>
        )}

        {(viewMode === 'add' || viewMode === 'edit') && (
          <div className="max-w-[1250px] mx-auto bg-white border border-gray-200/90 rounded-lg shadow-md overflow-hidden">
            <div className="p-4 bg-blue-900 text-white font-sans flex justify-between items-center">
              <span className="text-sm font-bold flex items-center gap-2">
                <ShieldAlert className="w-4 h-4" />
                ទម្រង់កត់ត្រាជនបរទេសបដិសេធ និងបញ្ជូនចេញ (Refusal & Deportation Form)
              </span>
            </div>

            <form onSubmit={handleSubmit} className="p-5 sm:p-6 space-y-5 text-xs font-sans">
              {/* Header Info Banner */}
              <div className="bg-sky-50 border-l-4 border-sky-500 p-3 rounded-r text-sky-800 text-[11px] flex items-center gap-2">
                <Info className="w-4 h-4 shrink-0" />
                <span>
                  លោកអ្នកកំពុងបំពេញទម្រង់កត់ត្រាការបដិសេធ និងបញ្ជូនចេញជនបរទេស។ សូមបំពេញព័ត៌មានជើងហោះហើរ និងតារាងព័ត៌មានជនបរទេសខាងក្រោម។
                </span>
              </div>

              {/* General Information */}
              <div className="space-y-4">
                <h4 className="text-[11px] font-bold text-gray-800 uppercase tracking-wide border-b border-gray-200 pb-2 flex items-center gap-1.5">
                  <span className="w-2.5 h-2.5 bg-blue-600 rounded-full inline-block"></span>
                  ១. ព័ត៌មានទូទៅនៃកាលបរិច្ឆេទ និងជើងហោះហើរ/ប្រទេស (GENERAL INFORMATION)
                </h4>

                <div className={`grid grid-cols-1 ${isSecondary ? 'sm:grid-cols-6' : 'sm:grid-cols-5'} gap-3 text-xs font-sans`}>
                  <div>
                    <label className="block font-bold text-gray-700 mb-1">កាលបរិច្ឆេទ (Date):</label>
                    <CustomDatePicker
                      value={formDate}
                      onChange={setFormDate}
                      placeholder="YYYY-MM-DD"
                      className="h-10 text-xs"
                    />
                  </div>

                  {isSecondary && (
                    <div>
                      <label className="block font-bold text-gray-700 mb-1">ក្រុម/ប៉ុស្តិ៍ច្រកទ្វារ:</label>
                      <select
                        value={formTeamName}
                        onChange={(e) => setFormTeamName(e.target.value)}
                        className="w-full h-10 border border-gray-300 rounded-md px-2.5 focus:border-blue-500 focus:outline-none bg-white font-bold text-blue-900 text-xs"
                      >
                        {formTeamsList.map((t) => (
                          <option key={t} value={t}>
                            {t}
                          </option>
                        ))}
                      </select>
                    </div>
                  )}

                  {/* Airport Specific Fields */}
                  {formGateType === 'airport' && (
                    <>
                      <div>
                        <label className="block font-bold text-gray-700 mb-1">ជើងហោះហើរមក:</label>
                        <input
                          type="text"
                          list="incoming-flights-opt"
                          value={formIncomingFlight}
                          onChange={(e) => handleIncomingFlightChange(e.target.value)}
                          placeholder="KR701 / ---"
                          className="w-full h-10 border border-gray-300 rounded-md px-3 focus:border-blue-500 focus:outline-none font-bold uppercase font-sans text-xs"
                        />
                        <datalist id="incoming-flights-opt">
                          {COMMON_INCOMING_FLIGHTS.map((f) => (
                            <option key={f.code} value={f.code}>
                              {f.code} - {f.origin}
                            </option>
                          ))}
                        </datalist>
                      </div>

                      <div>
                        <label className="block font-bold text-gray-700 mb-1">ជើងហោះហើរត្រឡប់:</label>
                        <input
                          type="text"
                          list="return-flights-opt"
                          value={formReturnFlight}
                          onChange={(e) => handleReturnFlightChange(e.target.value)}
                          placeholder="KR702 / ---"
                          className="w-full h-10 border border-gray-300 rounded-md px-3 focus:border-blue-500 focus:outline-none font-bold uppercase font-sans text-xs"
                        />
                        <datalist id="return-flights-opt">
                          {COMMON_OUTGOING_FLIGHTS.map((f) => (
                            <option key={f.code} value={f.code}>
                              {f.code} - {f.dest}
                            </option>
                          ))}
                        </datalist>
                      </div>

                      <div>
                        <label className="block font-bold text-gray-700 mb-1">មកពីប្រទេស:</label>
                        <input
                          type="text"
                          value={formOriginCountry}
                          onChange={(e) => setFormOriginCountry(e.target.value)}
                          placeholder="ប្រទេស..."
                          className="w-full h-10 border border-gray-300 rounded-md px-3 focus:border-blue-500 focus:outline-none font-bold text-gray-700 font-sans text-xs"
                        />
                      </div>

                      <div>
                        <label className="block font-bold text-gray-700 mb-1">ត្រឡប់ទៅប្រទេស:</label>
                        <input
                          type="text"
                          value={formDestinationCountry}
                          onChange={(e) => setFormDestinationCountry(e.target.value)}
                          placeholder="ប្រទេស..."
                          className="w-full h-10 border border-gray-300 rounded-md px-3 focus:border-blue-500 focus:outline-none font-bold text-gray-700 font-sans text-xs"
                        />
                      </div>
                    </>
                  )}

                  {/* Border Specific Fields */}
                  {formGateType === 'border' && (
                    <>
                      <div className="sm:col-span-2">
                        <label className="block font-bold text-gray-700 mb-1">មកពីប្រទេស:</label>
                        <select
                          value={formComingFromBorderCountry}
                          onChange={(e) => setFormComingFromBorderCountry(e.target.value)}
                          className="w-full h-10 border border-gray-300 rounded-md px-3 focus:border-blue-500 focus:outline-none font-bold text-gray-800 text-xs"
                        >
                          {BORDER_COUNTRIES.map((c) => (
                            <option key={c} value={c}>
                              {c}
                            </option>
                          ))}
                        </select>
                      </div>

                      <div className="sm:col-span-2">
                        <label className="block font-bold text-gray-700 mb-1">ត្រឡប់ទៅប្រទេស:</label>
                        <select
                          value={formReturningToBorderCountry}
                          onChange={(e) => setFormReturningToBorderCountry(e.target.value)}
                          className="w-full h-10 border border-gray-300 rounded-md px-3 focus:border-blue-500 focus:outline-none font-bold text-gray-800 text-xs"
                        >
                          {BORDER_COUNTRIES.map((c) => (
                            <option key={c} value={c}>
                              {c}
                            </option>
                          ))}
                        </select>
                      </div>
                    </>
                  )}
                </div>
              </div>

              {/* Dynamic list of people under this case (exactly matching switch visa form, with caseNumber & reason in the row) */}
              <div className="border-t border-gray-200 pt-4">
                <div className="flex items-center justify-between border-b border-gray-200 pb-2 mb-3">
                  <h4 className="text-xs font-bold text-gray-800 uppercase tracking-wide flex items-center gap-2">
                    <Users className="w-4 h-4 text-blue-800" />
                    <span>២. បញ្ជីឈ្មោះជនបរទេសដែលត្រូវបានបដិសេធ (Applicants List) - ({people.length} នាក់)</span>
                  </h4>
                  <button
                    type="button"
                    onClick={handleAddPerson}
                    className="h-7 px-3 bg-blue-50 text-blue-700 hover:bg-blue-100 rounded border border-blue-200 font-sans text-[11px] font-bold flex items-center gap-1 transition cursor-pointer"
                  >
                    <Plus className="w-3.5 h-3.5" />
                    <span>បន្ថែមជនបរទេស (Add Row)</span>
                  </button>
                </div>

                  <div className="space-y-2">
                  {people.map((p, index) => (
                    <div
                      key={p.id}
                      className="flex flex-nowrap items-center gap-1.5 bg-gray-50/80 p-2.5 rounded-lg border border-gray-200 hover:border-gray-300 transition text-xs overflow-x-auto"
                    >
                      {/* Row index indicator */}
                      <span className="font-bold bg-blue-900 text-white w-6 h-6 rounded-full flex items-center justify-center shrink-0 text-xs">
                        {index + 1}
                      </span>

                      {/* 1. Full Name (ឈ្មោះ) */}
                      <div className="flex-1 min-w-[140px]">
                        <label className="block text-[9px] text-gray-500 font-bold mb-0.5">ឈ្មោះពេញ (FULL NAME):</label>
                        <input
                          type="text"
                          required
                          placeholder="ឈ្មោះ (NAME)"
                          value={p.fullName}
                          onChange={(e) => updatePerson(p.id, 'fullName', e.target.value.toUpperCase())}
                          className="w-full border border-gray-300 rounded px-2 py-1 text-xs uppercase text-gray-900 font-bold focus:border-blue-500 focus:outline-none bg-white h-8"
                        />
                      </div>

                      {/* 2. Gender (ភេទ) */}
                      <div className="w-24 shrink-0">
                        <label className="block text-[9px] text-gray-500 font-bold mb-0.5">ភេទ:</label>
                        <select
                          value={p.gender}
                          onChange={(e) => updatePerson(p.id, 'gender', e.target.value as any)}
                          className="w-full border border-gray-300 rounded px-1.5 py-1 text-xs text-gray-800 focus:border-blue-500 focus:outline-none bg-white font-bold h-8"
                        >
                          <option value="">-- ជ្រើសរើស --</option>
                          <option value="ប្រុស">ប្រុស</option>
                          <option value="ស្រី">ស្រី</option>
                        </select>
                      </div>

                      {/* 3. DOB (ថ្ងៃកំណើត) */}
                      <div className="w-32 shrink-0">
                        <label className="block text-[9px] text-gray-500 font-bold mb-0.5">ថ្ងៃកំណើត (DOB):</label>
                        <CustomDatePicker
                          value={p.dob}
                          onChange={(val) => updatePerson(p.id, 'dob', val)}
                          placeholder="YYYY-MM-DD"
                          className="h-8 text-xs px-2"
                        />
                      </div>

                      {/* 4. Passport Number (លិខិតឆ្លងដែន) */}
                      <div className="w-36 shrink-0">
                        <label className="block text-[9px] text-gray-500 font-bold mb-0.5">លេខ Passport:</label>
                        <input
                          type="text"
                          required
                          placeholder="លេខ PASSPORT"
                          value={p.passportNumber}
                          onChange={(e) => handlePassportChange(p.id, e.target.value)}
                          className="w-full border border-gray-300 rounded px-2 py-1 text-xs uppercase text-gray-900 font-bold focus:border-blue-500 focus:outline-none bg-white h-8 tracking-wider"
                        />
                      </div>

                      {/* 5. Nationality (សញ្ជាតិ) */}
                      <div className="w-28 shrink-0">
                        <div className="flex justify-between items-center mb-0.5">
                          <span className="text-[9px] text-blue-700 font-bold truncate max-w-[50px]" title={PASSPORT_NATIONALITIES.find(n => n.code === p.nationality)?.nameKh}>
                            {PASSPORT_NATIONALITIES.find(n => n.code === p.nationality)?.nameKh || 'ចិន'}
                          </span>
                          <label className="text-[9px] text-gray-500 font-bold">សញ្ជាតិ (NAT):</label>
                        </div>
                        <input
                          type="text"
                          list="passport-nationalities-opt"
                          value={p.nationality}
                          onChange={(e) => updatePerson(p.id, 'nationality', e.target.value.toUpperCase())}
                          placeholder="CHN"
                          className="w-full border border-gray-300 rounded px-2 py-1 text-xs uppercase text-gray-900 font-bold focus:border-blue-500 focus:outline-none bg-white h-8"
                        />
                        <datalist id="passport-nationalities-opt">
                          {PASSPORT_NATIONALITIES.map((nat) => (
                            <option key={nat.code} value={nat.code} label={nat.nameKh} />
                          ))}
                        </datalist>
                      </div>

                      {/* 6. Reason for Refusal (មូលហេតុ) */}
                      <div className="flex-1 min-w-[220px]">
                        <label className="block text-[9px] text-gray-500 font-bold mb-0.5">មូលហេតុបដិសេធ (Reason):</label>
                        <select
                          value={p.reason}
                          onChange={(e) => updatePerson(p.id, 'reason', e.target.value)}
                          className="w-full border border-gray-300 rounded px-1.5 py-1 text-xs text-red-600 font-bold focus:border-blue-500 focus:outline-none bg-white h-8"
                        >
                          <option value="" className="text-red-600 font-bold">-- ជ្រើសរើស --</option>
                          {!dynamicReasons.includes(p.reason) && p.reason && (
                            <option value={p.reason}>{p.reason}</option>
                          )}
                          {dynamicReasons.map((r, i) => (
                            <option key={i} value={r}>
                              {r}
                            </option>
                          ))}
                        </select>
                      </div>

                      {/* Remove row button */}
                      <div className="self-end pb-0.5 shrink-0 ml-1">
                        <button
                          type="button"
                          onClick={() => handleRemovePerson(p.id)}
                          className="w-8 h-8 flex items-center justify-center bg-red-50 hover:bg-red-100 text-red-600 rounded transition border border-red-200 cursor-pointer"
                          title="លុបជួរនេះ"
                        >
                          <Trash2 className="w-3.5 h-3.5" />
                        </button>
                      </div>
                    </div>
                  ))}
                </div>
              </div>

              {/* Form Controls */}
              <div className="border-t border-gray-100 pt-4 flex justify-end gap-2.5">
                <button
                  type="button"
                  onClick={() => {
                    resetForm();
                    setViewMode('list');
                  }}
                  className="h-10 px-5 border border-gray-300 hover:bg-gray-100 rounded text-gray-600 transition font-bold"
                >
                  បោះបង់ (Cancel)
                </button>
                <button
                  type="submit"
                  className="h-10 px-6 bg-blue-600 hover:bg-blue-700 text-white rounded font-bold shadow-xs flex items-center gap-1.5 transition-all cursor-pointer"
                >
                  <CheckCircle2 className="w-4.5 h-4.5" />
                  <span>{viewMode === 'add' ? 'រក្សាទុកទិន្នន័យ (Save)' : 'រក្សាទុកការកែប្រែ (Save Changes)'}</span>
                </button>
              </div>
            </form>
          </div>
        )}

        {viewMode === 'report' && (
          <div className="space-y-4">
            {/* Toolbar for Report */}
            <div className="bg-white border border-gray-200/90 rounded-lg p-3 shadow-3xs flex justify-between items-center gap-2">
              <span className="text-xs font-black text-blue-900 font-sans">
                របាយការណ៍បោះពុម្ពសង្ខេប / Summary Printable Document
              </span>
              <div className="flex items-center gap-2">
                <button
                  onClick={handleExportPdf}
                  className="h-8 px-3.5 bg-emerald-600 hover:bg-emerald-700 text-white rounded font-sans text-xs font-bold flex items-center gap-1.5 transition-all shadow-xs cursor-pointer"
                >
                  <Download className="w-3.5 h-3.5" />
                  <span>ទាញយកជា PDF (Export)</span>
                </button>
              </div>
            </div>

            {/* Printable Document Layout */}
            <div className="max-w-[900px] mx-auto bg-white border border-gray-200 p-8 sm:p-12 shadow-md relative min-h-[1100px]" id="refusal-report-print-container">
              {/* Report Header */}
              <div className="text-center font-sans space-y-1.5 mb-8">
                <h2 className="text-sm font-bold uppercase tracking-widest text-gray-800">ព្រះរាជាណាចក្រកម្ពុជា</h2>
                <h3 className="text-xs font-bold text-gray-800">ជាតិ សាសនា ព្រះមហាក្សត្រ</h3>
                <div className="flex justify-center py-1">
                  <div className="w-20 border-b border-double border-gray-400"></div>
                </div>
                <div className="text-left mt-4 text-[11px] leading-tight space-y-0.5">
                  <span className="block font-bold">ក្រសួងមហាផ្ទៃ</span>
                  <span className="block font-bold">អគ្គនាយកដ្ឋានអន្តោប្រវេសន៍</span>
                  <span className="block font-bold text-blue-900">ការិយាល័យផ្តល់ទិដ្ឋាការ និងការអនុញ្ញាតស្នាក់នៅ</span>
                </div>
              </div>

              {/* Title */}
              <div className="text-center space-y-2 mb-6">
                <h1 className="text-sm font-black text-gray-900 uppercase">
                  របាយការណ៍សង្ខេបករណីបដិសេធ និងបញ្ជូនចេញ
                </h1>
                <p className="text-[10px] text-gray-500 font-sans italic">
                  គិតត្រឹមថ្ងៃទី ៖ {new Date().toLocaleDateString('kh-KH')}
                </p>
              </div>

              {/* Table listing all consolidated people */}
              <table className="w-full text-left border-collapse border border-gray-400 text-[10px] font-sans">
                <thead>
                  <tr className="bg-gray-100 text-center font-bold border-b border-gray-400">
                    <th className="py-2 px-1.5 border border-gray-400 w-8">ល.រ</th>
                    <th className="py-2 px-1.5 border border-gray-400 w-16">កាលបរិច្ឆេទ</th>
                    <th className="py-2 px-1.5 border border-gray-400 w-12">ករណី</th>
                    <th className="py-2 px-1.5 border border-gray-400">ឈ្មោះពេញ (NAME)</th>
                    <th className="py-2 px-1.5 border border-gray-400 w-10">ភេទ</th>
                    <th className="py-2 px-1.5 border border-gray-400 w-12">សញ្ជាតិ</th>
                    <th className="py-2 px-1.5 border border-gray-400 w-24">លេខលិខិតឆ្លងដែន</th>
                    <th className="py-2 px-1.5 border border-gray-400 w-28">មកពី / ត្រឡប់ទៅ</th>
                    <th className="py-2 px-1.5 border border-gray-400">មូលហេតុបដិសេធ</th>
                  </tr>
                </thead>
                <tbody>
                  {filteredRecords.length === 0 ? (
                    <tr>
                      <td colSpan={9} className="py-6 text-center text-gray-400 border border-gray-400">
                        មិនមានទិន្នន័យឡើយ
                      </td>
                    </tr>
                  ) : (
                    (() => {
                      let globalIdx = 1;
                      return filteredRecords.flatMap((r) => {
                        return (r.people || []).map((p, pIdx) => {
                          const isFirstPerson = pIdx === 0;
                          return (
                            <tr key={`${r.id}-${p.id}`} className="hover:bg-gray-50/50">
                              <td className="py-2 px-1.5 border border-gray-400 text-center">{globalIdx++}</td>
                              
                              {isFirstPerson ? (
                                <td className="py-2 px-1.5 border border-gray-400 text-center font-sans text-[9px]" rowSpan={r.people.length}>
                                  {r.date}
                                </td>
                              ) : null}

                              <td className="py-2 px-1.5 border border-gray-400 text-center font-bold">
                                {p.caseNumber}
                              </td>

                              <td className="py-2 px-1.5 border border-gray-400 font-bold">{p.fullName}</td>
                              <td className="py-2 px-1.5 border border-gray-400 text-center">{p.gender}</td>
                              <td className="py-2 px-1.5 border border-gray-400 text-center font-bold text-indigo-800">{p.nationality}</td>
                              <td className="py-2 px-1.5 border border-gray-400 font-mono tracking-wide text-center">{p.passportNumber}</td>

                              {isFirstPerson ? (
                                <td className="py-2 px-1.5 border border-gray-400 text-left font-sans text-[9px]" rowSpan={r.people.length}>
                                  {r.gateType === 'airport' ? (
                                    <div className="leading-tight">
                                      <span className="font-bold text-blue-900">In: {r.incomingFlight || '---'}</span>
                                      <span className="block text-[8px] text-gray-500 mt-0.5">Out: {r.returnFlight || '---'}</span>
                                    </div>
                                  ) : (
                                    <div className="leading-tight">
                                      <span className="font-bold text-gray-800">In: {r.comingFromBorderCountry}</span>
                                      <span className="block text-[8px] text-gray-500 mt-0.5">Out: {r.returningToBorderCountry}</span>
                                    </div>
                                  )}
                                </td>
                              ) : null}

                              <td className="py-2 px-1.5 border border-gray-400 text-left font-bold text-red-800 text-[9px]">
                                {p.reason}
                              </td>
                            </tr>
                          );
                        });
                      });
                    })()
                  )}
                </tbody>
              </table>

              {/* Signature section */}
              <div className="mt-12 flex justify-between items-start text-[11px] font-sans">
                <div className="text-center w-48">
                  <span className="block font-bold">រៀបចំដោយ</span>
                  <span className="block text-gray-400 text-[10px] mt-10">(ហត្ថលេខា និងឈ្មោះ)</span>
                </div>
                <div className="text-center w-56 space-y-1">
                  <span className="block text-gray-600 italic">ថ្ងៃទី ៖ .............................................</span>
                  <span className="block font-bold">ប្រធានការិយាល័យ / ក្រុមការងារ</span>
                  <span className="block text-gray-400 text-[10px] mt-10">(ហត្ថលេខា និងត្រា)</span>
                </div>
              </div>
            </div>
          </div>
        )}
      </div>
      {/* Delete Confirmation Modal (Replaces window.confirm to avoid iframe sandbox blocks) */}
      {deletePending && (
        <div className="fixed inset-0 bg-black/50 z-50 flex items-center justify-center p-4">
          <div className="bg-white rounded-xl shadow-2xl border border-gray-200 max-w-md w-full p-5 space-y-4 font-sans animate-in fade-in zoom-in-95 duration-150">
            <div className="flex items-start gap-3.5">
              <div className="w-10 h-10 rounded-full bg-red-100 flex items-center justify-center shrink-0">
                <Trash2 className="w-5 h-5 text-red-600" />
              </div>
              <div className="flex-1">
                <h3 className="text-sm font-bold text-gray-900">បញ្ជាក់ការលុបទិន្នន័យ</h3>
                <p className="text-xs text-gray-700 mt-1.5 leading-relaxed">{deletePending.title}</p>
              </div>
            </div>
            <div className="flex justify-end gap-2 pt-3 border-t border-gray-100">
              <button
                type="button"
                onClick={() => setDeletePending(null)}
                className="px-4 py-2 border border-gray-300 rounded-md text-xs font-bold text-gray-700 hover:bg-gray-50 transition cursor-pointer"
              >
                បោះបង់
              </button>
              <button
                type="button"
                onClick={confirmDeleteAction}
                className="px-4 py-2 bg-red-600 text-white rounded-md text-xs font-bold hover:bg-red-700 transition cursor-pointer flex items-center gap-1.5 shadow-xs"
              >
                <Trash2 className="w-3.5 h-3.5 text-white" />
                <span>យល់ព្រមលុប</span>
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
