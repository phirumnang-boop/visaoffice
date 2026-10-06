import React, { useState, useEffect, useMemo, useRef } from 'react';
import * as XLSX from 'xlsx';
import { saveAs } from 'file-saver';
import { CategoriesState, Officer, StockRecord, UserRole } from '../types';
import { exportElementsToPdf, exportElementToPdf } from '../utils/pdfExportHelper';
import { printA4Document } from '../utils/printHelper';
import { getKhmerLunarDate, getKhmerSolarParts, toKhmerNum, parseDateInput } from '../utils/khmerCalendar';
import {
  formatReportTeamName,
  normalizeTeamName,
  matchTeamInList,
  OFFICIAL_29_TEAMS,
  VISA_TYPES,
} from '../utils/teamNormalization';
import {
  resolveRecordTeamName,
  isCeaRecord,
  DEFAULT_OPENING_MATRIX,
  isOldStockTeamRecord,
  calculateAllTeamsStickerStockAtDate,
  isRecordForTeam,
  isRecordForRecipientTeam,
  isTransferTeamRecord,
  isIssueTeamRecord,
} from '../utils/teamStockCalculation';
import {
  WordPageSetupDialog,
  CustomMargins,
} from './WordPageSetupDialog';
import { CustomDatePicker } from './CustomDatePicker';
import {
  TacteingLine,
  TacteingControlSelector,
  getSavedTacteingSettings,
  saveTacteingSettings,
  TacteingType,
} from './TacteingLine';
import {
  Printer,
  Download,
  Calendar,
  Save,
  RotateCcw,
  FileSpreadsheet,
  FileText,
  ZoomIn,
  ZoomOut,
  SlidersHorizontal,
  CheckCircle2,
  Eye,
  Edit3,
  Layers,
  ArrowLeft,
  Search,
  Plus,
  Trash2,
  Building2,
  Users,
  Sparkles,
  ChevronDown,
  Layout,
  X,
} from 'lucide-react';

export interface DailyStickerVisaOperationReportProps {
  stockRecords?: StockRecord[];
  categories?: CategoriesState;
  officers?: Officer[];
  currentRole?: UserRole;
  userName?: string;
  assignedTeam?: string;
  onClose?: () => void;
  onShowToast?: (text: string, type?: 'success' | 'error') => void;
}

// 13 Official Visa Types
export const VISA_TYPE_KEYS = [
  'T', 'T1', 'T2', 'T3', 'E', 'E1', 'E2', 'E3', 'D',
  'K', 'A', 'B', 'C'
] as const;

export type VisaTypeKey = (typeof VISA_TYPE_KEYS)[number];

// Refused Foreigner Row Interface
export interface RefusedForeignerRow {
  id: string;
  no: number;
  name: string;
  gender: string;
  birthYear: string;
  nationality: string;
  passportNo: string;
  reason: string;
}

// Used Visa Detail Row Interface
export interface UsedVisaDetailRow {
  id: string;
  no: number;
  visaType: string;
  count: number;
  fromSerial: string;
  toSerial: string;
  note: string;
}

// Exact sample data from user's attached official PDF document (01-July-2024, Techo Airport)
export const PDF_SAMPLE_MATRIX_ROWS = {
  opening: {
    T: 183209, T1: 2259, T2: 1949, T3: 1147,
    E: 93185, E1: 1560, E2: 1253, E3: 557,
    D: 443, K: 7627, A: 936, B: 595, C: 3241,
  },
  received: {
    T: 0, T1: 0, T2: 0, T3: 0,
    E: 0, E1: 0, E2: 0, E3: 0,
    D: 0, K: 0, A: 0, B: 0, C: 0,
  },
  damaged: {
    T: 0, T1: 0, T2: 0, T3: 0,
    E: 0, E1: 0, E2: 0, E3: 0,
    D: 0, K: 0, A: 0, B: 0, C: 0,
  },
  used: {
    T: 977, T1: 9, T2: 0, T3: 4,
    E: 1055, E1: 10, E2: 0, E3: 0,
    D: 0, K: 150, A: 0, B: 1, C: 1,
  },
};

export const PDF_SAMPLE_USED_DETAILS: UsedVisaDetailRow[] = [
  { id: 'u-1', no: 1, visaType: 'T', count: 71, fromSerial: '1903010696', toSerial: '1903010766', note: '' },
  { id: 'u-2', no: 2, visaType: 'T1', count: 1, fromSerial: '9753', toSerial: '9753', note: '' },
  { id: 'u-3', no: 3, visaType: 'E', count: 78, fromSerial: '2000558543', toSerial: '2000558620', note: '' },
  { id: 'u-4', no: 4, visaType: 'E1', count: 2, fromSerial: '1900007721', toSerial: '1900007722', note: '' },
  { id: 'u-5', no: 5, visaType: 'K', count: 1, fromSerial: '2200122430', toSerial: '2200122430', note: '' },
  { id: 'u-6', no: 6, visaType: 'T', count: 906, fromSerial: '1903188465', toSerial: '1903189370', note: '' },
  { id: 'u-7', no: 7, visaType: 'T1', count: 8, fromSerial: '1900001510', toSerial: '1900001517', note: '' },
  { id: 'u-8', no: 8, visaType: 'T3', count: 4, fromSerial: '862', toSerial: '865', note: '' },
  { id: 'u-9', no: 9, visaType: 'E', count: 977, fromSerial: '2000531140', toSerial: '2000532116', note: '' },
  { id: 'u-10', no: 10, visaType: 'E1', count: 8, fromSerial: '1900007187', toSerial: '1900007194', note: '' },
  { id: 'u-11', no: 11, visaType: 'K', count: 149, fromSerial: '2200120120', toSerial: '2200120268', note: '' },
  { id: 'u-12', no: 12, visaType: 'B', count: 1, fromSerial: '99279', toSerial: '99279', note: '' },
  { id: 'u-13', no: 13, visaType: 'C', count: 1, fromSerial: '101222', toSerial: '101222', note: '' },
];

export const DailyStickerVisaOperationReport: React.FC<DailyStickerVisaOperationReportProps> = ({
  stockRecords = [],
  categories,
  officers = [],
  currentRole,
  userName = '',
  assignedTeam = '',
  onClose,
  onShowToast,
}) => {
  // Team selection
  const isTeam = currentRole !== 'Secondary' && currentRole !== 'Admin';

  // Teams derived from Office system categories (ទិន្នន័យក្រុមផ្តល់ទិដ្ឋាការ ពីប្រព័ន្ធការិយាល័យ)
  const officeTeamPairs = useMemo(() => {
    const fullItems = categories?.visaTeams || [];
    const robokItems = categories?.visaTeamsRobok || [];
    const maxCount = Math.max(fullItems.length, robokItems.length);

    if (maxCount > 0) {
      const list: Array<{ id: string; shortName: string; fullName: string }> = [];
      const seenIds = new Set<string>();
      for (let i = 0; i < maxCount; i++) {
        const full = fullItems[i]?.name?.trim() || '';
        let robok = robokItems[i]?.name?.trim() || '';
        let id = robokItems[i]?.id || fullItems[i]?.id || ('team-pair-' + (i + 1));

        if (full.includes('ព្រៃវល្លិ៍') && robok.includes('ព្រែកចាក')) {
          robok = 'ព្រំដែន ព្រៃវល្លិ៍';
        }
        if (full.includes('ត្រពាំងក្រៀល') && robok.includes('ត្រពាំងរូង')) {
          robok = 'ព្រំដែន ត្រពាំងក្រៀល';
        }

        const shortName = robok || (full ? normalizeTeamName(full) : ('ក្រុម ' + (i + 1)));
        const fullName = full || formatReportTeamName(shortName);

        if (seenIds.has(id)) {
          id = id + '_' + i;
        }
        seenIds.add(id);

        list.push({ id, shortName, fullName });
      }
      return list;
    }

    return OFFICIAL_29_TEAMS.map((t, idx) => ({
      id: `default-team-${idx + 1}`,
      shortName: t,
      fullName: formatReportTeamName(t),
    }));
  }, [categories?.visaTeams, categories?.visaTeamsRobok]);

  const resolvedAssignedTeam = useMemo(() => {
    if (!assignedTeam?.trim()) return '';
    const match = officeTeamPairs.find(
      (p) =>
        p.shortName === assignedTeam.trim() ||
        p.fullName === assignedTeam.trim() ||
        normalizeTeamName(p.shortName) === normalizeTeamName(assignedTeam.trim()) ||
        normalizeTeamName(p.fullName) === normalizeTeamName(assignedTeam.trim())
    );
    if (match) return match.shortName;
    return matchTeamInList(assignedTeam.trim(), Array.from(OFFICIAL_29_TEAMS));
  }, [assignedTeam, officeTeamPairs]);

  const initialTeam = isTeam && resolvedAssignedTeam ? resolvedAssignedTeam : (officeTeamPairs[0]?.shortName || 'អាកាស តេជោ');
  const [selectedTeam, setSelectedTeam] = useState<string>(() => {
    if (isTeam && resolvedAssignedTeam) {
      return resolvedAssignedTeam;
    }
    const saved = localStorage.getItem('daily_sticker_op_selected_team');
    return (saved && (officeTeamPairs.find(p => p.shortName === saved || normalizeTeamName(p.shortName) === normalizeTeamName(saved))?.shortName || matchTeamInList(saved, Array.from(OFFICIAL_29_TEAMS)))) || initialTeam;
  });

  // Keep selectedTeam strictly synchronized with assignedTeam when logged in as a team user
  useEffect(() => {
    if (isTeam && resolvedAssignedTeam) {
      setSelectedTeam(resolvedAssignedTeam);
    }
  }, [isTeam, resolvedAssignedTeam]);

  // Custom Editable Team Header Title state
  const [customTeamHeaderTitle, setCustomTeamHeaderTitle] = useState<string | null>(null);

  // Custom Margins State (in cm, defaults to standard A4 margins: Top: 1.5, Bottom: 1.5, Left: 2.0, Right: 1.5)
  const [customMargins, setCustomMargins] = useState<CustomMargins>(() => {
    try {
      const saved = localStorage.getItem('daily_sticker_report_margins');
      if (saved) {
        const parsed = JSON.parse(saved);
        if (typeof parsed?.top === 'number') {
          return {
            top: parsed.top,
            bottom: parsed.bottom ?? 1.5,
            left: parsed.left ?? 2.0,
            right: parsed.right ?? 1.5,
          };
        }
      }
    } catch {}
    return {
      top: 1.0,
      bottom: 1.5,
      left: 2.0,
      right: 1.5,
    };
  });

  const [showMarginControls, setShowMarginControls] = useState<boolean>(false);
  const [showPageSetupDialog, setShowPageSetupDialog] = useState<boolean>(false);
  const [pageSetupTab, setPageSetupTab] = useState<'margins' | 'paper' | 'layout'>('margins');
  const [pageOrientation, setPageOrientation] = useState<'portrait' | 'landscape'>('portrait');
  const [paperSize, setPaperSize] = useState<string>('A4');
  const [gutter, setGutter] = useState<number>(0);
  const [gutterPosition, setGutterPosition] = useState<'left' | 'top'>('left');
  const [headerMargin, setHeaderMargin] = useState<number>(1.25);
  const [footerMargin, setFooterMargin] = useState<number>(1.25);
  const [activeTabStop, setActiveTabStop] = useState<number>(1.27);
  const [tempMargins, setTempMargins] = useState<CustomMargins>(() => ({ ...customMargins }));

  // Auto-persist custom margins to localStorage
  useEffect(() => {
    try {
      localStorage.setItem('daily_sticker_report_margins', JSON.stringify(customMargins));
    } catch {}
  }, [customMargins]);

  // Tacteing Settings
  const [showTacteingModal, setShowTacteingModal] = useState<boolean>(false);
  const [tacteingSettings, setTacteingSettings] = useState<{ type: TacteingType; customImage: string | null }>(() =>
    getSavedTacteingSettings()
  );

  useEffect(() => {
    const handleTactChange = () => {
      setTacteingSettings(getSavedTacteingSettings());
    };
    window.addEventListener('tacteing_settings_updated', handleTactChange);
    window.addEventListener('storage', handleTactChange);
    return () => {
      window.removeEventListener('tacteing_settings_updated', handleTactChange);
      window.removeEventListener('storage', handleTactChange);
    };
  }, []);

  const handleTacteingChange = (type: TacteingType, customImg?: string | null) => {
    setTacteingSettings({ type, customImage: customImg || null });
    saveTacteingSettings(type, customImg);
    if (onShowToast) {
      onShowToast('បានប្តូរម៉ូតតាក់តែងបន្ទាត់ដោយជោគជ័យ!', 'success');
    }
  };

  const handleApplyMargins = (margins: CustomMargins) => {
    setCustomMargins(margins);
    setShowPageSetupDialog(false);
    if (onShowToast) {
      onShowToast('បានអនុវត្តការកំណត់គែមក្រដាស (Margins) ដោយជោគជ័យ!', 'success');
    }
  };

  // Date selection (defaults to 2018-12-01 as requested for official baseline and operations)
  const [selectedDate, setSelectedDate] = useState<string>(() => {
    const saved = localStorage.getItem('daily_sticker_op_selected_date');
    return saved || '2018-12-01';
  });

  // View & UI controls
  const [zoomLevel, setZoomLevel] = useState<number>(100);
  const [isEditMode, setIsEditMode] = useState<boolean>(false);
  const [viewLayout, setViewLayout] = useState<'two-pages' | 'continuous'>('two-pages');
  const [isExportingPdf, setIsExportingPdf] = useState<boolean>(false);

  // During PDF export, disable editing inputs so clean, official static document renders
  const isEditing = isEditMode && !isExportingPdf;

  // Section I: Staff State
  const [staffInfo, setStaffInfo] = useState({
    total: '០២ នាក់',
    present: '០២ នាក់',
    absentWithReason: '០០ នាក់',
    absentWithoutReason: '០០ នាក់',
  });

  // Section II: Refused Foreigners
  const [refusalSummaryText, setRefusalSummaryText] = useState('១ ករណី ចំនួន ២');
  const [refusedForeigners, setRefusedForeigners] = useState<RefusedForeignerRow[]>([
    { id: 'rf-1', no: 1, name: '', gender: '', birthYear: '', nationality: '', passportNo: '', reason: '' },
    { id: 'rf-2', no: 2, name: '', gender: '', birthYear: '', nationality: '', passportNo: '', reason: '' },
    { id: 'rf-3', no: 3, name: '', gender: '', birthYear: '', nationality: '', passportNo: '', reason: '' },
    { id: 'rf-4', no: 4, name: '', gender: '', birthYear: '', nationality: '', passportNo: '', reason: '' },
    { id: 'rf-5', no: 5, name: '', gender: '', birthYear: '', nationality: '', passportNo: '', reason: '' },
  ]);

  const emptyZeros: Record<VisaTypeKey, number> = {
    T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0,
  };

  // Section III: Matrix Table Data (13 Visa Types) - Starts clean with emptyZeros (-)
  const [matrixData, setMatrixData] = useState<{
    opening: Record<VisaTypeKey, number>;
    received: Record<VisaTypeKey, number>;
    damaged: Record<VisaTypeKey, number>;
    used: Record<VisaTypeKey, number>;
  }>({
    opening: { ...emptyZeros },
    received: { ...emptyZeros },
    damaged: { ...emptyZeros },
    used: { ...emptyZeros },
  });

  // Section III: Used Visa Detail Rows
  const [usedDetails, setUsedDetails] = useState<UsedVisaDetailRow[]>([]);

  // Sections IV, V, VI: Challenges, Requests, Directions
  const [challengeText, setChallengeText] = useState('គ្មាន');
  const [requestText, setRequestText] = useState('គ្មាន');
  const [directionText, setDirectionText] = useState('គ្មាន');

  // Sign-off
  const [signatoryTitle, setSignatoryTitle] = useState('ប្រធានក្រុមផ្តល់ទិដ្ឋាការ');
  const [signatoryName, setSignatoryName] = useState('');
  const [signatoryCity, setSignatoryCity] = useState('រាជធានីភ្នំពេញ');

  // Print ref
  const printContainerRef = useRef<HTMLDivElement>(null);

  // Storage key for caching customizations per team & date (v2 ensures clean slate from legacy cache)
  const storageKey = useMemo(() => {
    return `daily_sticker_report_v2_${normalizeTeamName(selectedTeam)}_${selectedDate}`;
  }, [selectedTeam, selectedDate]);

  // Save current data to localStorage
  const handleSaveData = () => {
    try {
      const payload = {
        staffInfo,
        refusalSummaryText,
        refusedForeigners,
        matrixData,
        usedDetails,
        challengeText,
        requestText,
        directionText,
        signatoryTitle,
        signatoryName,
        signatoryCity,
        customTeamHeaderTitle,
        customMargins,
        savedAt: new Date().toISOString(),
      };
      localStorage.setItem(storageKey, JSON.stringify(payload));
      localStorage.setItem('daily_sticker_op_selected_team', selectedTeam);
      localStorage.setItem('daily_sticker_op_selected_date', selectedDate);
      if (onShowToast) {
        onShowToast('បានរក្សាទុកទិន្នន័យរបាយការណ៍សន្លឹកទិដ្ឋាការរួចរាល់!', 'success');
      }
    } catch (e) {
      if (onShowToast) onShowToast('មិនអាចរក្សាទុកទិន្នន័យបានទេ', 'error');
    }
  };

  // Helper to convert any Latin digits to Khmer digits
  const toKhmerDigits = (val: string) => {
    if (!val) return '០';
    const khmerDigits = ['០', '១', '២', '៣', '៤', '៥', '៦', '៧', '៨', '៩'];
    return val.replace(/[0-9]/g, (d) => khmerDigits[parseInt(d, 10)]);
  };

  // Helper to render staff count using Khmer numbers
  const renderStaffDisplay = (val: string) => {
    const khmerVal = toKhmerDigits(val);
    return <span className="font-siemreap font-bold">{khmerVal}</span>;
  };

  // Reset to PDF reference sample data
  const handleResetToSample = () => {
    setMatrixData({
      opening: { ...PDF_SAMPLE_MATRIX_ROWS.opening },
      received: { ...PDF_SAMPLE_MATRIX_ROWS.received },
      damaged: { ...PDF_SAMPLE_MATRIX_ROWS.damaged },
      used: { ...PDF_SAMPLE_MATRIX_ROWS.used },
    });
    setUsedDetails([...PDF_SAMPLE_USED_DETAILS]);
    setStaffInfo({
      total: '02 នាក់',
      present: '02 នាក់',
      absentWithReason: '00 នាក់',
      absentWithoutReason: '00 នាក់',
    });
    setRefusalSummaryText('១ ករណី ចំនួន ២');
    setRefusedForeigners([
      { id: 'rf-1', no: 1, name: '', gender: '', birthYear: '', nationality: '', passportNo: '', reason: '' },
      { id: 'rf-2', no: 2, name: '', gender: '', birthYear: '', nationality: '', passportNo: '', reason: '' },
      { id: 'rf-3', no: 3, name: '', gender: '', birthYear: '', nationality: '', passportNo: '', reason: '' },
      { id: 'rf-4', no: 4, name: '', gender: '', birthYear: '', nationality: '', passportNo: '', reason: '' },
      { id: 'rf-5', no: 5, name: '', gender: '', birthYear: '', nationality: '', passportNo: '', reason: '' },
    ]);
    setChallengeText('គ្មាន');
    setRequestText('គ្មាន');
    setDirectionText('គ្មាន');
    if (onShowToast) {
      onShowToast('បានកំណត់ឡើងវិញជាទិន្នន័យគំរូផ្លូវការ (PDF Sample Data)', 'success');
    }
  };

  // Helper to calculate team staff numbers based on actual officers data in the system
  const calculateTeamStaffInfo = (teamName: string, dateStr: string) => {
    try {
      const normTeam = normalizeTeamName(teamName);
      if (!normTeam) {
        return {
          total: '០០ នាក់',
          present: '០០ នាក់',
          absentWithReason: '០០ នាក់',
          absentWithoutReason: '០០ នាក់',
        };
      }

      // 1. Gather all officers from props or localStorage
      let allOfficers: Officer[] = [...(officers || [])];
      try {
        const keys = ['app_officers', 'app_officers_records_v1', 'officers', 'app_officers_list', 'officer_records'];
        keys.forEach((k) => {
          const saved = localStorage.getItem(k);
          if (saved) {
            const parsed = JSON.parse(saved);
            if (Array.isArray(parsed) && parsed.length > 0) {
              const map = new Map<string, Officer>();
              allOfficers.forEach((o) => { if (o && o.id) map.set(o.id, o); });
              parsed.forEach((o: Officer) => { if (o && o.id) map.set(o.id, o); });
              allOfficers = Array.from(map.values());
            }
          }
        });
      } catch {}

      // 2. Build map of category IDs and normalized names for the selected team
      const teamCategoryIds = new Set<string>();
      const teamNamesToMatch = new Set<string>();
      teamNamesToMatch.add(normTeam);
      teamNamesToMatch.add(normalizeTeamName(selectedTeam));

      (categories?.visaTeams || []).forEach((vt) => {
        const vtNorm = normalizeTeamName(vt.name);
        if (vtNorm === normTeam || vt.name === selectedTeam || (vtNorm && normTeam && (vtNorm.includes(normTeam) || normTeam.includes(vtNorm)))) {
          teamCategoryIds.add(vt.id);
          teamNamesToMatch.add(vtNorm);
        }
      });
      (categories?.visaTeamsRobok || []).forEach((vt) => {
        const vtNorm = normalizeTeamName(vt.name);
        if (vtNorm === normTeam || vt.name === selectedTeam || (vtNorm && normTeam && (vtNorm.includes(normTeam) || normTeam.includes(vtNorm)))) {
          teamCategoryIds.add(vt.id);
          teamNamesToMatch.add(vtNorm);
        }
      });

      // 3. Filter officers belonging to selected team
      const matchingOfficers = allOfficers.filter((of) => {
        if (!of) return false;

        // Match by visaTeamId
        if (of.visaTeamId) {
          if (teamCategoryIds.has(of.visaTeamId)) return true;
          if (normalizeTeamName(of.visaTeamId) === normTeam) return true;
          const matchedCategory = categories?.visaTeams?.find((c) => c.id === of.visaTeamId) ||
                                  categories?.visaTeamsRobok?.find((c) => c.id === of.visaTeamId);
          if (matchedCategory && normalizeTeamName(matchedCategory.name) === normTeam) return true;
        }

        // Match by text fields
        const fieldsToSearch = [
          of.officeWork,
          of.station,
          of.department,
          (of as any).visaTeamName,
          (of as any).visaTeamRobokName,
        ];

        for (const field of fieldsToSearch) {
          if (field) {
            const fieldNorm = normalizeTeamName(field);
            if (fieldNorm && (teamNamesToMatch.has(fieldNorm) || fieldNorm === normTeam || fieldNorm.includes(normTeam) || normTeam.includes(fieldNorm))) {
              return true;
            }
          }
        }

        return false;
      });

      const totalCount = matchingOfficers.length;

      // 4. Check if daily attendance record exists in localStorage for this team & date
      const targetDateISO = normalizeDateToISO(dateStr) || dateStr.trim();
      let presentCount = totalCount;
      let absentReasonCount = 0;
      let absentNoReasonCount = 0;

      try {
        const savedAttendance =
          localStorage.getItem(`app_team_attendance_${normTeam}_${targetDateISO}`) ||
          localStorage.getItem('app_officer_attendance_records');
        if (savedAttendance) {
          const parsedAtt = JSON.parse(savedAttendance);
          if (Array.isArray(parsedAtt)) {
            const dateAtt = parsedAtt.filter((a: any) => {
              const aDate = normalizeDateToISO(a.date || a.createdAt || '');
              const aTeam = normalizeTeamName(a.teamName || a.officeWork || '');
              return aDate === targetDateISO && aTeam === normTeam;
            });
            if (dateAtt.length > 0) {
              const pCount = dateAtt.filter((a: any) => a.status === 'present' || a.status === 'វត្តមាន').length;
              const arCount = dateAtt.filter((a: any) => a.status === 'absent_reason' || a.status === 'មានមូលហេតុ').length;
              const anrCount = dateAtt.filter((a: any) => a.status === 'absent_no_reason' || a.status === 'អត់មូលហេតុ').length;
              if (pCount + arCount + anrCount > 0) {
                presentCount = pCount;
                absentReasonCount = arCount;
                absentNoReasonCount = anrCount;
              }
            }
          }
        }
      } catch {}

      // Format as Khmer numbers e.g. "០២ នាក់" or "១៥ នាក់"
      const formatStaffNum = (num: number) => {
        const kh = toKhmerNum(num);
        const padded = num < 10 && !kh.startsWith('០') ? `០${kh}` : kh;
        return `${padded} នាក់`;
      };

      return {
        total: formatStaffNum(totalCount),
        present: formatStaffNum(presentCount),
        absentWithReason: formatStaffNum(absentReasonCount),
        absentWithoutReason: formatStaffNum(absentNoReasonCount),
      };
    } catch {
      return {
        total: '០០ នាក់',
        present: '០០ នាក់',
        absentWithReason: '០០ នាក់',
        absentWithoutReason: '០០ នាក់',
      };
    }
  };

  // Helper to format DOB to full DD-MM-YYYY format
  const formatDobToDDMMYYYY = (dobStr: string): string => {
    if (!dobStr || !dobStr.trim()) return '---';
    const clean = dobStr.trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(clean)) {
      const [y, m, d] = clean.split('-');
      return `${d}-${m}-${y}`;
    }
    if (/^\d{4}\/\d{2}\/\d{2}$/.test(clean)) {
      const [y, m, d] = clean.split('/');
      return `${d}-${m}-${y}`;
    }
    return clean;
  };

  // Helper to pull refusal/deportation data from app_refusal_deportation_records_v1 for target team and date
  const pullRefusalDataForTeamAndDate = (teamName: string, dateStr: string) => {
    try {
      const refusalSaved = localStorage.getItem('app_refusal_deportation_records_v1');
      if (!refusalSaved) return { summaryText: '០ ករណី ចំនួន ០ នាក់', rows: [] };

      const parsedRefusal = JSON.parse(refusalSaved);
      if (!Array.isArray(parsedRefusal)) return { summaryText: '០ ករណី ចំនួន ០ នាក់', rows: [] };

      const targetDateISO = normalizeDateToISO(dateStr) || dateStr.trim();
      const normTeam = normalizeTeamName(teamName);

      const matchingRecords = parsedRefusal.filter((r: any) => {
        const recDate = normalizeDateToISO(r.date || r.createdAt || '');
        if (recDate !== targetDateISO) return false;

        const recTeam = normalizeTeamName(r.teamName || '');
        return recTeam === normTeam;
      });

      const allPeople: Array<{
        fullName: string;
        gender: string;
        dob: string;
        nationality: string;
        passportNumber: string;
        reason: string;
      }> = [];

      matchingRecords.forEach((r: any) => {
        if (r.people && Array.isArray(r.people)) {
          r.people.forEach((p: any) => {
            allPeople.push(p);
          });
        }
      });

      const totalCases = matchingRecords.length;
      const totalPeople = allPeople.length;

      let summaryText = '០ ករណី ចំនួន ០ នាក់';
      if (totalPeople > 0) {
        summaryText = `${toKhmerNum(totalCases)} ករណី ចំនួន ${toKhmerNum(totalPeople)} នាក់`;
      }

      if (allPeople.length === 0) {
        return { summaryText, rows: [] };
      }

      const pulledRows: RefusedForeignerRow[] = allPeople.map((p, idx) => ({
        id: `rf-pulled-${idx + 1}-${Date.now()}`,
        no: idx + 1,
        name: p.fullName || '---',
        gender: p.gender || '---',
        birthYear: formatDobToDDMMYYYY(p.dob),
        nationality: p.nationality || '---',
        passportNo: p.passportNumber || '---',
        reason: p.reason || '---',
      }));

      return { summaryText, rows: pulledRows };
    } catch {
      return { summaryText: '០ ករណី ចំនួន ០ នាក់', rows: [] };
    }
  };

  // Helper to normalize visa type
  const normalizeVisaType = (val: any): string => {
    if (!val) return '';
    const s = String(val).trim().toUpperCase();
    if (s.startsWith('VISA ')) return s.replace('VISA ', '').trim();
    return s;
  };

  // Helper to normalize date string to YYYY-MM-DD
  const normalizeDateToISO = (val: any): string => {
    if (!val) return '';
    const s = String(val).trim();
    if (s.includes('T')) return s.slice(0, 10);
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    if (s.includes('/')) {
      const parts = s.split('/');
      if (parts.length === 3) {
        if (parts[0].length === 4) return `${parts[0]}-${parts[1].padStart(2, '0')}-${parts[2].padStart(2, '0')}`;
        return `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
      }
    }
    return s.slice(0, 10);
  };

  // Helper to get previous date in YYYY-MM-DD format
  const getPreviousDateISO = (dateStr: string): string => {
    if (!dateStr) return '';
    const [y, m, d] = dateStr.split('-').map(Number);
    if (!y || !m || !d) return '';
    const dt = new Date(Date.UTC(y, m - 1, d));
    dt.setUTCDate(dt.getUTCDate() - 1);
    return dt.toISOString().slice(0, 10);
  };

  // Clear report data to dashes (-)
  const handleClearToDashes = () => {
    setMatrixData({
      opening: { ...emptyZeros },
      received: { ...emptyZeros },
      damaged: { ...emptyZeros },
      used: { ...emptyZeros },
    });
    setUsedDetails([]);
    try {
      localStorage.setItem(storageKey, JSON.stringify({ isCleared: true }));
    } catch {}
    if (onShowToast) {
      onShowToast('បានកំណត់តារាងជា (-) ទាំងអស់រួចរាល់', 'info');
    }
  };

  // Pull actual stock data from system stock records and daily team operations with automatic roll over
  const handlePullSystemData = (showFeedback: boolean = true) => {
    try {
      const normTeam = normalizeTeamName(selectedTeam);
      const targetDateISO = normalizeDateToISO(selectedDate) || selectedDate.trim() || '2018-12-01';

      // 1. Calculate stock from system using official calculateAllTeamsStickerStockAtDate
      const calculated = calculateAllTeamsStickerStockAtDate(
        stockRecords || [],
        targetDateISO,
        Array.from(OFFICIAL_29_TEAMS)
      );

      const matchedKey =
        Object.keys(calculated.remainingBeforeMatrix).find(
          (k) => k === selectedTeam || normalizeTeamName(k) === normTeam
        ) || selectedTeam;

      // 2. Base Opening Stock
      // If 2018-12-01 or earlier: strictly the 30-Nov-2018 baseline old stock
      let openingStock: Record<VisaTypeKey, number> = {
        T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0,
      };

      const teamRemBefore = calculated.remainingBeforeMatrix[matchedKey] || {};
      const fallbackBaseline = DEFAULT_OPENING_MATRIX[normTeam] || DEFAULT_OPENING_MATRIX[selectedTeam] || {};

      VISA_TYPE_KEYS.forEach((vt) => {
        if (targetDateISO <= '2018-12-01') {
          openingStock[vt] = fallbackBaseline[vt] ?? teamRemBefore[vt] ?? 0;
        } else {
          openingStock[vt] = teamRemBefore[vt] ?? fallbackBaseline[vt] ?? 0;
        }
      });

      // 3. Roll Over from previous day if saved report exists in localStorage
      const prevDate = getPreviousDateISO(targetDateISO);
      if (prevDate && targetDateISO > '2018-12-01') {
        try {
          const prevSaved =
            localStorage.getItem(`daily_sticker_report_v2_${normTeam}_${prevDate}`) ||
            localStorage.getItem(`daily_sticker_report_${normTeam}_${prevDate}`);
          if (prevSaved) {
            const parsed = JSON.parse(prevSaved);
            if (parsed && parsed.matrixData) {
              const prevOp = parsed.matrixData.opening || {};
              const prevRec = parsed.matrixData.received || {};
              const prevDam = parsed.matrixData.damaged || {};
              const prevUsed = parsed.matrixData.used || {};
              VISA_TYPE_KEYS.forEach((vt) => {
                const op = Number(prevOp[vt]) || 0;
                const rc = Number(prevRec[vt]) || 0;
                const dm = Number(prevDam[vt]) || 0;
                const us = Number(prevUsed[vt]) || 0;
                openingStock[vt] = Math.max(0, op + rc - dm - us);
              });
            }
          }
        } catch (e) {
          console.error('Error reading prev saved report:', e);
        }
      }

      // 4. Issued, Damaged on Target Date (Strictly for target date and this team only, preventing cumulative bloat)
      const receivedStock: Record<VisaTypeKey, number> = {
        T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0,
      };
      const damagedStock: Record<VisaTypeKey, number> = {
        T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0,
      };
      let usedStock: Record<VisaTypeKey, number> = {
        T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0,
      };

      const teamUsedToday = calculated.usedTodayMatrix[matchedKey] || {};
      VISA_TYPE_KEYS.forEach((vt) => {
        usedStock[vt] = teamUsedToday[vt] || 0;
      });

      // Filter stockRecords strictly on targetDateISO for this specific team
      (stockRecords || []).forEach((rec) => {
        if (rec.stockType && rec.stockType !== 'sticker') return;
        if (isCeaRecord(rec)) return;
        if (isOldStockTeamRecord(rec)) return;

        // Skip auto-mirrored records created by DailyTeamVisaOperations to avoid duplicates
        if (rec.id && (rec.id.startsWith('stock-dtr-') || rec.id.startsWith('dtr-'))) {
          return;
        }

        const recDate = normalizeDateToISO(rec.date || rec.createdAt || '');
        if (recDate !== targetDateISO) return;

        const vt = normalizeVisaType(rec.visaType) as VisaTypeKey;
        if (!VISA_TYPE_KEYS.includes(vt)) return;

        const qty = Number(rec.totalSheets || (rec.quantityBundles ? rec.quantityBundles * 50 : 0) || (rec as any).quantity || 0);
        if (qty <= 0) return;

        const op = (rec.operationType || '').trim().toLowerCase();
        const src = (rec.sourceFrom || '').trim().toLowerCase();
        const rem = (rec.remarks || '').trim().toLowerCase();
        const notes = ((rec as any).notes || '').trim().toLowerCase();

        // Check if transfer operation
        const isTransfer =
          isTransferTeamRecord(rec) ||
          op === 'transferteam' ||
          op === 'transferuseteam' ||
          op === 'transfer' ||
          op.includes('transfer') ||
          op.includes('ផ្ទេរ') ||
          src.includes('ផ្ទេរ') ||
          rem.includes('ផ្ទេរ') ||
          notes.includes('ផ្ទេរ') ||
          Boolean(rec.recipientTeamName || (rec as any).recipientTeamId);

        const isSenderOrOwner = isRecordForTeam(rec, selectedTeam);
        const isRecipient = isRecordForRecipientTeam(rec, selectedTeam);

        // ១. បើកថ្មី = ទទួលពីការិយាល័យ ឬផ្ទេរពីក្រុម
        // Case 1a: ទទួលពីការិយាល័យ
        const isReceiveFromOffice =
          isSenderOrOwner &&
          !isTransfer &&
          (
            isIssueTeamRecord(rec) ||
            op === 'issueteam' ||
            op === 'issue_team' ||
            op === 'openteam' ||
            op === 'distribution' ||
            op === 'stock_in' ||
            op === 'issue' ||
            op.includes('បើក') ||
            op.includes('issue') ||
            op.includes('ទទួលពីការិយាល័យ') ||
            op.includes('ទទួលពីក២')
          );

        // Case 1b: ផ្ទេរពីក្រុម (ទទួលផ្ទេរការប្រើប្រាស់ពីក្រុមផ្សេង)
        const isTransferIn = isTransfer && isRecipient;

        if (isReceiveFromOffice || isTransferIn) {
          receivedStock[vt] = (receivedStock[vt] || 0) + qty;
          return;
        }

        // ២. ផ្ទេរ/ខូច/ខ្វះ/បង្វិល = ផ្ទេរការប្រើប្រាស់ + ទិដ្ឋាការខូចក្រុម + ទិដ្ឋាការខ្វះក្រុម + ទិដ្ឋាការបង្វិលទៅក២
        // Case 2a: ផ្ទេរការប្រើប្រាស់ (ផ្ទេរចេញទៅក្រុមផ្សេង)
        const isTransferOut = isTransfer && isSenderOrOwner && !isRecipient;

        // Case 2b: ទិដ្ឋាការខូចក្រុម
        const isDamaged =
          isSenderOrOwner &&
          !isTransfer &&
          (op === 'damagedteam' || op === 'damaged' || op.includes('ខូច'));

        // Case 2c: ទិដ្ឋាការខ្វះក្រុម
        const isMissing =
          isSenderOrOwner &&
          !isTransfer &&
          (op === 'missingteam' || op === 'missing' || op.includes('ខ្វះ') || op.includes('បាត់'));

        // Case 2d: ទិដ្ឋាការបង្វិលទៅក២
        const isReturnK2 =
          isSenderOrOwner &&
          !isTransfer &&
          (op === 'returnteam' || op === 'returnoffice' || op === 'recycle' || op.includes('បង្វិល') || op.includes('ប្រគល់'));

        if (isTransferOut || isDamaged || isMissing || isReturnK2) {
          damagedStock[vt] = (damagedStock[vt] || 0) + qty;
          return;
        }
      });

      // Check app_daily_team_operations_v5 for any extra operations on target date (if not in stockRecords)
      try {
        const savedOps = localStorage.getItem('app_daily_team_operations_v5');
        if (savedOps) {
          const parsedOps = JSON.parse(savedOps);
          if (Array.isArray(parsedOps)) {
            parsedOps.forEach((dRec: any) => {
              if (dRec.categoryType === 'cEA' || (dRec.categoryType && dRec.categoryType !== 'Sticker')) return;
              const dDate = normalizeDateToISO(dRec.date || '');
              if (dDate !== targetDateISO) return;

              // Avoid duplicate if this record is already in stockRecords
              if (dRec.id && (stockRecords || []).some((sr) => sr.id === dRec.id)) return;

              const op = (dRec.operationType || '').trim().toLowerCase();
              if (!op || op === 'useteam' || op === 'team_usage' || op === 'usage' || op.includes('ប្រើ')) return;

              const isSenderOrOwner = isRecordForTeam(dRec, selectedTeam);
              const isRecipient = isRecordForRecipientTeam(dRec, selectedTeam);
              const isTransfer = isTransferTeamRecord(dRec) || op === 'transferteam' || op.includes('ផ្ទេរ');

              VISA_TYPE_KEYS.forEach((vt) => {
                const item = dRec.values?.[vt];
                if (!item) return;
                const qty = Number(item.quantity || 0);
                if (qty <= 0) return;

                // 1. បើកថ្មី = ទទួលពីការិយាល័យ ឬផ្ទេរពីក្រុម
                const isReceiveFromOffice =
                  isSenderOrOwner &&
                  !isTransfer &&
                  (op === 'issueteam' || op === 'openteam' || op.includes('បើក') || op.includes('ទទួល'));
                const isTransferIn = isTransfer && isRecipient;

                if (isReceiveFromOffice || isTransferIn) {
                  receivedStock[vt] = (receivedStock[vt] || 0) + qty;
                  return;
                }

                // 2. ផ្ទេរ/ខូច/ខ្វះ/បង្វិល = ផ្ទេរការប្រើប្រាស់ + ទិដ្ឋាការខូចក្រុម + ទិដ្ឋាការខ្វះក្រុម + ទិដ្ឋាការបង្វិលទៅក២
                const isTransferOut = isTransfer && isSenderOrOwner && !isRecipient;
                const isDamaged = isSenderOrOwner && !isTransfer && (op.includes('ខូច') || op === 'damagedteam');
                const isMissing = isSenderOrOwner && !isTransfer && (op.includes('ខ្វះ') || op === 'missingteam');
                const isReturnK2 = isSenderOrOwner && !isTransfer && (op.includes('បង្វិល') || op === 'returnteam');

                if (isTransferOut || isDamaged || isMissing || isReturnK2) {
                  damagedStock[vt] = (damagedStock[vt] || 0) + qty;
                }
              });
            });
          }
        }
      } catch (e) {
        console.warn('Error reading app_daily_team_operations_v5 for non-usage ops:', e);
      }

      // 5. Collect Used details on Target Date with strict deduplication
      const rawDetails: UsedVisaDetailRow[] = [];

      // 5a. First collect from primary daily operations (app_daily_team_operations_v5)
      // This has the rich form breakdown with oldCode (កូដចាស់) and individual serial ranges
      try {
        const saved5 = localStorage.getItem('app_daily_team_operations_v5');
        if (saved5) {
          const parsed5 = JSON.parse(saved5);
          if (Array.isArray(parsed5)) {
            parsed5.forEach((dRec: any) => {
              if (dRec.categoryType === 'cEA' || (dRec.categoryType && dRec.categoryType !== 'Sticker')) return;
              const dTeam = normalizeTeamName(dRec.teamName || '');
              if (dTeam !== normTeam) return;
              const dDate = normalizeDateToISO(dRec.date || '');
              if (dDate !== targetDateISO) return;

              if (dRec.values) {
                VISA_TYPE_KEYS.forEach((vt) => {
                  const item = dRec.values[vt];
                  if (!item) return;

                  if (item.entries && Array.isArray(item.entries) && item.entries.length > 0) {
                    item.entries.forEach((ent: any) => {
                      const q = parseInt(ent.quantity, 10) || 0;
                      if (q > 0) {
                        rawDetails.push({
                          id: `sys-entry-${ent.id || Math.random()}`,
                          no: 0,
                          visaType: vt,
                          count: q,
                          fromSerial: (ent.startSerial || '').trim(),
                          toSerial: (ent.endSerial || '').trim(),
                          note: '',
                        });
                      }
                    });
                  } else {
                    const q = parseInt(String(item.quantity || ''), 10) || 0;
                    if (q > 0) {
                      rawDetails.push({
                        id: `sys-val-${Math.random()}`,
                        no: 0,
                        visaType: vt,
                        count: q,
                        fromSerial: (item.startSerial || '').trim(),
                        toSerial: (item.endSerial || '').trim(),
                        note: '',
                      });
                    }
                  }
                });
              }
            });
          }
        }
      } catch (e) {
        console.warn('Error reading app_daily_team_operations_v5 in DailyStickerVisaOperationReport', e);
      }

      // 5b. Then collect from stockRecords (skip auto-mirrored stock-dtr-* records to avoid duplicate entries)
      (stockRecords || []).forEach((rec) => {
        if (rec.stockType && rec.stockType !== 'sticker') return;
        if (isCeaRecord(rec)) return;
        if (isOldStockTeamRecord(rec)) return;

        // Skip auto-mirrored records created by DailyTeamVisaOperations
        if (rec.id && (rec.id.startsWith('stock-dtr-') || rec.id.startsWith('dtr-'))) {
          return;
        }

        const rawTeam = resolveRecordTeamName(rec) || (rec as any).teamName || rec.visaTeamRobokName || '';
        if (normalizeTeamName(rawTeam) !== normTeam) return;

        const recDate = normalizeDateToISO(rec.date || rec.createdAt || '');
        if (recDate !== targetDateISO) return;

        const vt = normalizeVisaType(rec.visaType) as VisaTypeKey;
        if (!VISA_TYPE_KEYS.includes(vt)) return;

        const qty = Number(rec.totalSheets || (rec.quantityBundles ? rec.quantityBundles * 50 : 0) || (rec as any).quantity || 0);
        const op = (rec.operationType || '').toLowerCase();

        if (op === 'useteam' || op === 'team_usage' || op === 'usage' || op.includes('ប្រើ')) {
          if (qty > 0) {
            rawDetails.push({
              id: `sys-sr-${rec.id || Math.random()}`,
              no: 0,
              visaType: vt,
              count: qty,
              fromSerial: (rec.startSerial || '').trim(),
              toSerial: (rec.endSerial || '').trim(),
              note: '',
            });
          }
        }
      });

      // 5c. Deduplicate rawDetails by visaType + serial numbers
      const deduplicatedMap = new Map<string, UsedVisaDetailRow>();
      rawDetails.forEach((row) => {
        // Unique signature key
        const sig = `${row.visaType}_${row.fromSerial}_${row.toSerial}_${row.count}`;
        if (!deduplicatedMap.has(sig)) {
          deduplicatedMap.set(sig, row);
        }
      });

      const detailsOnDate: UsedVisaDetailRow[] = Array.from(deduplicatedMap.values()).map((row, index) => ({
        ...row,
        no: index + 1,
      }));

      // If details exist, reconcile usedStock with detailsOnDate
      if (detailsOnDate.length > 0) {
        const detailsUsage: Record<VisaTypeKey, number> = {
          T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0,
        };
        detailsOnDate.forEach((row) => {
          const vt = (row.visaType || '').trim().toUpperCase() as VisaTypeKey;
          if (VISA_TYPE_KEYS.includes(vt)) {
            detailsUsage[vt] = (detailsUsage[vt] || 0) + (Number(row.count) || 0);
          }
        });
        if (Object.values(detailsUsage).some((v) => v > 0)) {
          usedStock = detailsUsage;
        }
      }

      // 6. Pull Refusal / Deportation Data and Staff Info for this team & date
      const refusalData = pullRefusalDataForTeamAndDate(selectedTeam, targetDateISO);
      let activeRefusalSummaryText = refusalData.summaryText;
      let activeRefusedRows = refusalData.rows || [];

      const activeStaffInfo = calculateTeamStaffInfo(selectedTeam, targetDateISO);
      setStaffInfo(activeStaffInfo);

      setRefusalSummaryText(activeRefusalSummaryText);
      setRefusedForeigners(activeRefusedRows);

      // Update State immediately
      setMatrixData({
        opening: openingStock,
        received: receivedStock,
        damaged: damagedStock,
        used: usedStock,
      });
      setUsedDetails(detailsOnDate);

      // Save to localStorage for persistence
      try {
        const toSave = {
          staffInfo: activeStaffInfo,
          refusalSummaryText: activeRefusalSummaryText,
          refusedForeigners: activeRefusedRows,
          matrixData: {
            opening: openingStock,
            received: receivedStock,
            damaged: damagedStock,
            used: usedStock,
          },
          usedDetails: detailsOnDate,
          challengeText,
          requestText,
          directionText,
          signatoryTitle,
          signatoryName,
          signatoryCity,
          savedAt: new Date().toISOString(),
          isCleared: false,
        };
        localStorage.setItem(storageKey, JSON.stringify(toSave));
      } catch {}

      if (showFeedback && onShowToast) {
        if (targetDateISO === '2018-12-01') {
          onShowToast('បានបង្ហាញលទ្ធផលទិន្នន័យថ្ងៃទី ០១ ខែ ធ្នូ ឆ្នាំ ២០១៨ (សល់ពីថ្ងៃចាស់ = ស្តុកចាស់ក្រុម ៣០ វិច្ឆិកា ២០១៨) ជោគជ័យ', 'success');
        } else {
          onShowToast(`បានបង្ហាញលទ្ធផលទិន្នន័យ ${selectedTeam} សម្រាប់ថ្ងៃទី ${targetDateISO} ដោយជោគជ័យ!`, 'success');
        }
      }
    } catch (err: any) {
      console.error('Error pulling system data:', err);
      if (showFeedback && onShowToast) {
        onShowToast('មានបញ្ហាក្នុងការបង្ហាញលទ្ធផល៖ ' + (err?.message || ''), 'error');
      }
    }
  };

  // Load custom saved data on team/date change or auto-calculate and show live results
  useEffect(() => {
    const refusalData = pullRefusalDataForTeamAndDate(selectedTeam, selectedDate);
    const calculatedStaff = calculateTeamStaffInfo(selectedTeam, selectedDate);

    try {
      const saved = localStorage.getItem(storageKey);
      if (saved) {
        const parsed = JSON.parse(saved);
        if (parsed.isCleared) {
          setMatrixData({
            opening: { ...emptyZeros },
            received: { ...emptyZeros },
            damaged: { ...emptyZeros },
            used: { ...emptyZeros },
          });
          setUsedDetails([]);
          return;
        }
        setStaffInfo(calculatedStaff);

        // Sync refusal data if saved version has filled names OR fallback to fresh refusal records
        const hasSavedRefusedNames = parsed.refusedForeigners && Array.isArray(parsed.refusedForeigners) && parsed.refusedForeigners.some((r: any) => r.name && r.name.trim().length > 0);
        if (hasSavedRefusedNames) {
          setRefusedForeigners(parsed.refusedForeigners);
          if (parsed.refusalSummaryText !== undefined) setRefusalSummaryText(parsed.refusalSummaryText);
        } else if (refusalData.rows && refusalData.rows.length > 0) {
          setRefusedForeigners(refusalData.rows);
          setRefusalSummaryText(refusalData.summaryText);
        } else {
          if (parsed.refusalSummaryText !== undefined) setRefusalSummaryText(parsed.refusalSummaryText);
          if (parsed.refusedForeigners) setRefusedForeigners(parsed.refusedForeigners);
        }

        if (parsed.matrixData) {
          if (parsed.matrixData.opening && parsed.matrixData.opening.T === 32002) {
            parsed.matrixData.opening = { ...PDF_SAMPLE_MATRIX_ROWS.opening };
          }
          setMatrixData(parsed.matrixData);
        }
        if (parsed.usedDetails && Array.isArray(parsed.usedDetails)) {
          setUsedDetails(
            parsed.usedDetails.map((u: any) => ({
              ...u,
              note: u.note && !u.note.includes('កូដចាស់') && u.note.trim() !== '-' ? u.note : '',
            }))
          );
        }
        if (parsed.challengeText !== undefined) setChallengeText(parsed.challengeText);
        if (parsed.requestText !== undefined) setRequestText(parsed.requestText);
        if (parsed.directionText !== undefined) setDirectionText(parsed.directionText);
        if (parsed.signatoryTitle !== undefined) setSignatoryTitle(parsed.signatoryTitle);
        if (parsed.signatoryName !== undefined) setSignatoryName(parsed.signatoryName);
        if (parsed.signatoryCity !== undefined) setSignatoryCity(parsed.signatoryCity);
        if (parsed.customTeamHeaderTitle !== undefined) setCustomTeamHeaderTitle(parsed.customTeamHeaderTitle);
        if (parsed.customMargins) setCustomMargins(parsed.customMargins);
        return;
      }
    } catch (e) {
      console.error('Error loading saved daily sticker report:', e);
    }

    // If no saved data, populate staff info & refusal data for target date if present
    setStaffInfo(calculatedStaff);
    if (refusalData.rows && refusalData.rows.length > 0) {
      setRefusedForeigners(refusalData.rows);
      setRefusalSummaryText(refusalData.summaryText);
    } else {
      setRefusalSummaryText('០ ករណី ចំនួន ០ នាក់');
      setRefusedForeigners([]);
    }

    // Initialize with dashes (-) and wait for user to click "បង្ហាញលទ្ធផល" button
    setMatrixData({
      opening: { ...emptyZeros },
      received: { ...emptyZeros },
      damaged: { ...emptyZeros },
      used: { ...emptyZeros },
    });
    setUsedDetails([]);
  }, [storageKey, selectedTeam, selectedDate, officers]);

  // Effective used derived from usedDetails table or fallback to matrixData.used
  const effectiveUsed = useMemo(() => {
    const usage: Record<VisaTypeKey, number> = {
      T: 0, T1: 0, T2: 0, T3: 0,
      E: 0, E1: 0, E2: 0, E3: 0,
      D: 0, K: 0, A: 0, B: 0, C: 0,
    };
    usedDetails.forEach((row) => {
      const vt = (row.visaType || '').trim().toUpperCase() as VisaTypeKey;
      if (VISA_TYPE_KEYS.includes(vt)) {
        usage[vt] = (usage[vt] || 0) + (Number(row.count) || 0);
      }
    });
    const hasDetailUsage = Object.values(usage).some((v) => v > 0);
    return hasDetailUsage ? usage : matrixData.used;
  }, [usedDetails, matrixData.used]);

  // Matrix calculations
  const calculatedEndingBalance = useMemo(() => {
    const ending: Record<VisaTypeKey, number> = {} as any;
    VISA_TYPE_KEYS.forEach((vt) => {
      const op = matrixData.opening[vt] || 0;
      const rec = matrixData.received[vt] || 0;
      const dam = matrixData.damaged[vt] || 0;
      const u = effectiveUsed[vt] || 0;
      ending[vt] = op + rec - dam - u;
    });
    return ending;
  }, [matrixData, effectiveUsed]);

  // Row Totals for Matrix Table
  const rowTotals = useMemo(() => {
    const sumOpening = (Object.values(matrixData.opening) as number[]).reduce((a, b) => a + (Number(b) || 0), 0);
    const sumReceived = (Object.values(matrixData.received) as number[]).reduce((a, b) => a + (Number(b) || 0), 0);
    const sumDamaged = (Object.values(matrixData.damaged) as number[]).reduce((a, b) => a + (Number(b) || 0), 0);
    const sumUsed = (Object.values(effectiveUsed) as number[]).reduce((a, b) => a + (Number(b) || 0), 0);
    const sumEnding = sumOpening + sumReceived - sumDamaged - sumUsed;
    return {
      opening: sumOpening,
      received: sumReceived,
      damaged: sumDamaged,
      used: sumUsed,
      ending: sumEnding,
    };
  }, [matrixData, effectiveUsed]);

  // Detail table total count
  const totalUsedDetailsCount = useMemo(() => {
    return usedDetails.reduce((sum, row) => sum + (Number(row.count) || 0), 0);
  }, [usedDetails]);

  // Full Khmer Lunar & Solar dates
  const { lunarDateStr, solarHeaderDateStr, solarSignDateStr } = useMemo(() => {
    const lunar = getKhmerLunarDate(selectedDate);
    const solarParts = getKhmerSolarParts(selectedDate);
    return {
      lunarDateStr: lunar || 'ថ្ងៃចន្ទ ១០រោច ខែបឋមាសាឍ ឆ្នាំរោង ឆស័ក ព.ស ២៥៦៨',
      solarHeaderDateStr: `ថ្ងៃទី ${solarParts.khmerDay} ខែ ${solarParts.khmerMonth} ឆ្នាំ ${solarParts.khmerYear}`,
      solarSignDateStr: `${signatoryCity}, ថ្ងៃទី${solarParts.khmerDay} ខែ${solarParts.khmerMonth} ឆ្នាំ${solarParts.khmerYear}`,
    };
  }, [selectedDate, signatoryCity]);

  // Full administrative title of team strictly taken from Office system categories (ទិន្នន័យក្រុមផ្តល់ទិដ្ឋាការ ពីប្រព័ន្ធការិយាល័យ)
  const fullTeamHeaderTitle = useMemo(() => {
    if (customTeamHeaderTitle && customTeamHeaderTitle.trim()) {
      return customTeamHeaderTitle.trim();
    }
    if (!selectedTeam) return 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារ';

    // 1. Direct match in officeTeamPairs
    const matched = officeTeamPairs.find(
      (pair) =>
        pair.shortName === selectedTeam ||
        pair.fullName === selectedTeam ||
        normalizeTeamName(pair.shortName) === normalizeTeamName(selectedTeam) ||
        normalizeTeamName(pair.fullName) === normalizeTeamName(selectedTeam)
    );
    if (matched && matched.fullName) {
      return matched.fullName;
    }

    // 2. Direct match in categories.visaTeams
    const catFull = (categories?.visaTeams || []).find(
      (vt) =>
        vt.name === selectedTeam ||
        normalizeTeamName(vt.name) === normalizeTeamName(selectedTeam) ||
        vt.name.includes(selectedTeam)
    );
    if (catFull && catFull.name) return catFull.name;

    // 3. Fallback
    return formatReportTeamName(selectedTeam);
  }, [customTeamHeaderTitle, selectedTeam, officeTeamPairs, categories?.visaTeams]);

  // Matrix cell change helper
  const handleMatrixChange = (
    rowKey: 'opening' | 'received' | 'damaged' | 'used',
    vt: VisaTypeKey,
    valStr: string
  ) => {
    const num = parseInt(valStr.replace(/[^\d]/g, ''), 10) || 0;
    setMatrixData((prev) => ({
      ...prev,
      [rowKey]: {
        ...prev[rowKey],
        [vt]: num,
      },
    }));
  };

  // Add / remove used detail rows
  const handleAddUsedRow = () => {
    setUsedDetails((prev) => [
      ...prev,
      {
        id: `u-${Date.now()}`,
        no: prev.length + 1,
        visaType: 'T',
        count: 0,
        fromSerial: '',
        toSerial: '',
        note: '',
      },
    ]);
  };

  const handleRemoveUsedRow = (id: string) => {
    setUsedDetails((prev) => {
      const filtered = prev.filter((r) => r.id !== id);
      return filtered.map((r, idx) => ({ ...r, no: idx + 1 }));
    });
  };

  const handleUpdateUsedRow = (id: string, field: keyof UsedVisaDetailRow, val: any) => {
    setUsedDetails((prev) =>
      prev.map((r) => (r.id === id ? { ...r, [field]: val } : r))
    );
  };

  // Add / remove refused foreigner rows
  const handleAddRefusedRow = () => {
    setRefusedForeigners((prev) => [
      ...prev,
      {
        id: `rf-${Date.now()}`,
        no: prev.length + 1,
        name: '',
        gender: '',
        birthYear: '',
        nationality: '',
        passportNo: '',
        reason: '',
      },
    ]);
  };

  const handleRemoveRefusedRow = (id: string) => {
    setRefusedForeigners((prev) => {
      const filtered = prev.filter((r) => r.id !== id);
      return filtered.map((r, idx) => ({ ...r, no: idx + 1 }));
    });
  };

  const handleUpdateRefusedRow = (id: string, field: keyof RefusedForeignerRow, val: any) => {
    setRefusedForeigners((prev) =>
      prev.map((r) => (r.id === id ? { ...r, [field]: val } : r))
    );
  };

  // High-fidelity Print
  const handlePrint = () => {
    if (!printContainerRef.current) return;
    printA4Document(printContainerRef.current, {
      orientation: 'portrait',
      documentTitle: `Daily_Sticker_Report_${selectedTeam}_${selectedDate}`,
      pageMargin: '0',
    });
  };

  // High-fidelity PDF Export matching workspace 100%
  const handleExportPdf = async () => {
    const page1 = document.getElementById('daily-sticker-report-page-1');
    const page2 = document.getElementById('daily-sticker-report-page-2');
    const pages = [page1, page2].filter(Boolean) as HTMLElement[];
    if (pages.length === 0 && printContainerRef.current) {
      pages.push(printContainerRef.current);
    }
    if (pages.length === 0) return;

    // Temporarily reset zoom transform on printContainerRef so capture is 100% true native scale
    const originalTransform = printContainerRef.current?.style.transform;

    try {
      setIsExportingPdf(true);
      if (printContainerRef.current) {
        printContainerRef.current.style.transform = 'none';
      }

      // Wait a moment for React to re-render in pristine non-editing view
      await new Promise((resolve) => setTimeout(resolve, 200));

      const cleanTeam = (selectedTeam || 'ក្រុម').replace(/[/\\?%*:|"<>]/g, '_').trim();
      const fileName = `របាយការណ៍សន្លឹកទិដ្ឋាការ_${cleanTeam}_${selectedDate}.pdf`;

      await exportElementsToPdf(pages, fileName, {
        pixelRatio: 3,
        orientation: 'portrait',
      });

      if (onShowToast) onShowToast('បានទាញយកឯកសារ PDF ជោគជ័យ!', 'success');
    } catch (err) {
      console.error('PDF export failed:', err);
      if (onShowToast) onShowToast('ការទាញយក PDF បានបរាជ័យ', 'error');
    } finally {
      if (printContainerRef.current && originalTransform !== undefined) {
        printContainerRef.current.style.transform = originalTransform;
      }
      setIsExportingPdf(false);
    }
  };

  // Excel Export
  const handleExportExcel = () => {
    const wb = XLSX.utils.book_new();

    // Sheet 1: Matrix
    const matrixRows = [
      ['កិច្ចប្រតិបត្តិការទិដ្ឋាការប្រចាំថ្ងៃ - សន្លឹកទិដ្ឋាការស្អិត'],
      ['ក្រុម:', fullTeamHeaderTitle],
      ['កាលបរិច្ឆេទ:', selectedDate, solarHeaderDateStr],
      [],
      ['ល.រ', 'ប្រតិបត្តិការ', ...VISA_TYPE_KEYS, 'សរុប'],
      ['១', 'សល់ពីថ្ងៃចាស់', ...VISA_TYPE_KEYS.map((k) => matrixData.opening[k]), rowTotals.opening],
      ['២', 'បើកថ្មី', ...VISA_TYPE_KEYS.map((k) => matrixData.received[k]), rowTotals.received],
      ['៣', 'ផ្ទេរ/ខូច/ខ្វះ/បង្វិល', ...VISA_TYPE_KEYS.map((k) => matrixData.damaged[k]), rowTotals.damaged],
      ['៤', 'ប្រើប្រាស់', ...VISA_TYPE_KEYS.map((k) => matrixData.used[k]), rowTotals.used],
      ['៥', 'សរុបរួម', ...VISA_TYPE_KEYS.map((k) => calculatedEndingBalance[k]), rowTotals.ending],
    ];

    const wsMatrix = XLSX.utils.aoa_to_sheet(matrixRows);
    XLSX.utils.book_append_sheet(wb, wsMatrix, 'Matrix');

    // Sheet 2: Used details
    const detailRows = [
      ['បញ្ជីទិដ្ឋាការដែលបានប្រើប្រាស់រួច'],
      ['កាលបរិច្ឆេទ:', selectedDate],
      [],
      ['ល.រ', 'ប្រភេទទិដ្ឋាការ', 'ចំនួន', 'ចាប់ពីលេខ', 'ដល់លេខ', 'ផ្សេងៗ'],
      ...usedDetails.map((r) => [
        r.no,
        r.visaType,
        r.count,
        r.fromSerial,
        r.toSerial,
        r.note && !r.note.includes('កូដចាស់') && r.note.trim() !== '-' ? r.note : '',
      ]),
      ['សរុបការប្រើប្រាស់', '', totalUsedDetailsCount, '', '', ''],
    ];

    const wsDetails = XLSX.utils.aoa_to_sheet(detailRows);
    XLSX.utils.book_append_sheet(wb, wsDetails, 'Used_Details');

    XLSX.writeFile(wb, `Daily_Sticker_Visa_Report_${selectedTeam}_${selectedDate}.xlsx`);
    if (onShowToast) onShowToast('បានទាញយកឯកសារ Excel ជោគជ័យ!', 'success');
  };

  // Splitting usedDetails for Page 1 and Page 2 (all on Page 1 now to fit in single page)
  const page1UsedRows = useMemo(() => {
    return usedDetails;
  }, [usedDetails]);

  const page2UsedRows = useMemo(() => {
    return [];
  }, []);

  return (
    <div className="min-h-screen bg-slate-100 flex flex-col font-sans pb-16">
      {/* Top Application Toolbar */}
      <div className="sticky top-0 z-40 bg-white/95 backdrop-blur-md border-b border-slate-200 shadow-xs px-4 py-2.5">
        <div className="max-w-7xl mx-auto flex flex-wrap items-center justify-between gap-3">
          {/* Left: Back / Title / Team info */}
          <div className="flex items-center gap-3">
            {onClose && (
              <button
                onClick={onClose}
                className="p-1.5 rounded-lg border border-slate-200 text-slate-600 hover:bg-slate-100 hover:text-slate-900 transition cursor-pointer"
                title="ត្រឡប់ក្រោយ"
              >
                <ArrowLeft className="w-4 h-4" />
              </button>
            )}
            <div>
              <div className="flex items-center gap-2">
                <FileText className="w-5 h-5 text-blue-700" />
                <h1 className="text-base font-bold text-slate-900 leading-tight">
                  របាយការណ៍សន្លឹកទិដ្ឋាការប្រចាំថ្ងៃ (A4)
                </h1>
                <span className="text-[11px] px-2 py-0.5 rounded-full font-semibold bg-blue-50 text-blue-700 border border-blue-200">
                  ស្តង់ដារ A4
                </span>
              </div>
              <p className="text-xs text-slate-500 mt-0.5">
                ទម្រង់ក្រដាស A4 ផ្លូវការ៖ កិច្ចប្រតិបត្តិការ និងការប្រើប្រាស់សន្លឹកទិដ្ឋាការតាមក្រុម
              </p>
            </div>
          </div>

          {/* Center: Team & Date Selectors */}
          <div className="flex items-center gap-2 flex-wrap">
            {/* Team Selector - Only for Secondary Office */}
            {!isTeam && (
              <div className="flex items-center gap-1.5 bg-slate-50 border border-slate-200 rounded-lg px-2.5 py-1 text-xs">
                <Building2 className="w-3.5 h-3.5 text-blue-600" />
                <label className="text-slate-500 font-medium">ក្រុម៖</label>
                <select
                  value={selectedTeam}
                  onChange={(e) => {
                    setSelectedTeam(e.target.value);
                    try {
                      localStorage.setItem('daily_sticker_op_selected_team', e.target.value);
                    } catch {}
                  }}
                  className="bg-transparent font-bold text-slate-800 focus:outline-none cursor-pointer text-xs"
                >
                  {officeTeamPairs.map((t) => (
                    <option key={t.id} value={t.shortName}>
                      {t.shortName} ({t.fullName})
                    </option>
                  ))}
                </select>
              </div>
            )}

            {/* Date Selector */}
            <div className="flex items-center gap-1.5 bg-slate-50 border border-slate-200 rounded-lg px-2.5 py-1 text-xs">
              <Calendar className="w-3.5 h-3.5 text-emerald-600" />
              <label className="text-slate-500 font-medium">ថ្ងៃ៖</label>
              <CustomDatePicker
                value={selectedDate}
                onChange={(dateStr) => {
                  setSelectedDate(dateStr);
                  try {
                    localStorage.setItem('daily_sticker_op_selected_date', dateStr);
                  } catch {}
                }}
                placeholder="YYYY-MM-DD"
              />
            </div>

            {/* Show Result Button */}
            <button
              onClick={() => handlePullSystemData(true)}
              className="flex items-center gap-1.5 px-3.5 py-1.5 rounded-lg border border-blue-600 bg-blue-600 text-white hover:bg-blue-700 text-xs font-bold shadow-xs hover:shadow transition cursor-pointer"
              title="ចុចដើម្បីបង្ហាញលទ្ធផលទិន្នន័យជាក់ស្តែងភ្លាមៗ (Click to show result)"
            >
              <Search className="w-3.5 h-3.5" />
              <span>បង្ហាញលទ្ធផល</span>
            </button>

            {/* Clear to dashes button */}
            <button
              onClick={handleClearToDashes}
              className="flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg border border-slate-300 bg-white text-slate-700 hover:bg-slate-100 text-xs font-medium transition cursor-pointer"
              title="កំណត់ទិន្នន័យក្នុងតារាងជាសញ្ញាដក (-) ទាំងអស់ឡើងវិញ"
            >
              <RotateCcw className="w-3.5 h-3.5 text-slate-500" />
              <span>សម្អាត (-)</span>
            </button>

            {/* Quick Sample Data */}
            <button
              onClick={handleResetToSample}
              className="flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg border border-amber-200 bg-amber-50 text-amber-700 hover:bg-amber-100 text-xs font-medium transition cursor-pointer"
              title="ផ្ទុកទិន្នន័យគំរូដូចក្នុងឯកសារយោង PDF (០១ កក្កដា ២០២៤)"
            >
              <Sparkles className="w-3.5 h-3.5" />
              <span>គំរូ PDF</span>
            </button>
          </div>

          {/* Right: Actions, Margins, Tacteing, Zoom & Export */}
          <div className="flex items-center gap-2 flex-wrap">
            {/* Quick Margins Toggle */}
            <button
              type="button"
              onClick={() => setShowMarginControls(!showMarginControls)}
              className={`flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-xs font-medium transition cursor-pointer border shadow-2xs ${
                showMarginControls
                  ? 'bg-blue-100 text-blue-900 border-blue-300 font-bold'
                  : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
              }`}
              title="កែសម្រួលគែមក្រដាស A4 (Margins Quick Panel)"
            >
              <SlidersHorizontal className="w-3.5 h-3.5 text-indigo-600" />
              <span>គែមក្រដាស</span>
            </button>

            {/* Page Setup Dialog Launcher */}
            <button
              type="button"
              onClick={() => {
                setTempMargins({ ...customMargins });
                setShowPageSetupDialog(true);
              }}
              className="flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg bg-white border border-slate-300 text-slate-700 hover:bg-slate-50 text-xs font-medium transition cursor-pointer shadow-2xs"
              title="កំណត់ទំព័រ និងគែមក្រដាស A4 (Page Setup & Margins)"
            >
              <Layout className="w-3.5 h-3.5 text-blue-600" />
              <span>Page Setup...</span>
            </button>

            {/* Tacteing Ornament Settings Button */}
            <div className="relative">
              <button
                type="button"
                onClick={() => setShowTacteingModal(!showTacteingModal)}
                className={`flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-xs font-medium transition cursor-pointer border shadow-2xs ${
                  showTacteingModal
                    ? 'bg-amber-100 text-amber-900 border-amber-300 font-bold'
                    : 'bg-white text-slate-700 border-slate-300 hover:bg-slate-50'
                }`}
                title="កំណត់ និងជ្រើសរើសរូបតាក់តែងបន្ទាត់ (Tacteing Line Ornament)"
              >
                <Sparkles className="w-3.5 h-3.5 text-amber-600" />
                <span>ម៉ូតតាក់តែង</span>
              </button>

              {showTacteingModal && (
                <div className="absolute top-full mt-2 right-0 z-50 bg-white rounded-lg shadow-2xl border border-amber-300 p-3 min-w-[340px] animate-fadeIn">
                  <div className="flex items-center justify-between pb-2 mb-2 border-b border-gray-100">
                    <span className="font-bold text-xs text-gray-800 flex items-center gap-1.5">
                      <Sparkles className="w-3.5 h-3.5 text-amber-600" />
                      <span>កំណត់រូបភាពតាក់តែងបន្ទាត់ (Tacteing Ornament)</span>
                    </span>
                    <button
                      type="button"
                      onClick={() => setShowTacteingModal(false)}
                      className="text-gray-400 hover:text-gray-600 p-0.5 rounded cursor-pointer"
                    >
                      <X className="w-3.5 h-3.5" />
                    </button>
                  </div>
                  <TacteingControlSelector
                    currentType={tacteingSettings.type}
                    customImage={tacteingSettings.customImage}
                    onChange={handleTacteingChange}
                  />
                  <div className="text-right mt-2 pt-1 border-t border-gray-100">
                    <button
                      type="button"
                      onClick={() => setShowTacteingModal(false)}
                      className="text-xs text-gray-600 hover:text-gray-900 font-semibold px-2 py-0.5 rounded bg-gray-100 hover:bg-gray-200 cursor-pointer"
                    >
                      បិទ
                    </button>
                  </div>
                </div>
              )}
            </div>

            {/* Edit Mode Toggle */}
            <button
              onClick={() => setIsEditMode(!isEditMode)}
              className={`flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg text-xs font-medium transition cursor-pointer border ${
                isEditMode
                  ? 'bg-amber-600 text-white border-amber-700 shadow-xs'
                  : 'bg-slate-50 text-slate-700 border-slate-200 hover:bg-slate-100'
              }`}
              title="បើក/បិទ ការកែប្រែទិន្នន័យលើក្រដាសដោយផ្ទាល់"
            >
              <Edit3 className="w-3.5 h-3.5" />
              <span>{isEditMode ? 'កំពុងកែសម្រួល' : 'កែប្រែ'}</span>
            </button>

            {/* Save Button */}
            <button
              onClick={handleSaveData}
              className="flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg bg-emerald-600 hover:bg-emerald-700 text-white text-xs font-semibold shadow-xs transition cursor-pointer"
              title="រក្សាទុកការកែប្រែ"
            >
              <Save className="w-3.5 h-3.5" />
              <span>រក្សាទុក</span>
            </button>

            {/* PDF Export */}
            <button
              onClick={handleExportPdf}
              disabled={isExportingPdf}
              className="flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg bg-red-600 hover:bg-red-700 disabled:opacity-50 text-white text-xs font-semibold shadow-xs transition cursor-pointer"
              title="ទាញយកជា PDF"
            >
              <Download className="w-3.5 h-3.5" />
              <span>{isExportingPdf ? 'កំពុងទាញ...' : 'PDF'}</span>
            </button>

            {/* Excel Export */}
            <button
              onClick={handleExportExcel}
              className="flex items-center gap-1.5 px-2.5 py-1.5 rounded-lg bg-emerald-700 hover:bg-emerald-800 text-white text-xs font-semibold shadow-xs transition cursor-pointer"
              title="ទាញយកជា Excel"
            >
              <FileSpreadsheet className="w-3.5 h-3.5" />
              <span>Excel</span>
            </button>

            {/* Zoom Controls */}
            <div className="flex items-center gap-1 bg-slate-100 rounded-lg p-0.5 border border-slate-200">
              <button
                onClick={() => setZoomLevel((z) => Math.max(z - 10, 60))}
                className="p-1 hover:bg-white rounded text-slate-600 transition cursor-pointer"
                title="បង្រួម"
              >
                <ZoomOut className="w-3.5 h-3.5" />
              </button>
              <span className="text-[11px] font-bold px-1.5 text-slate-700 select-none">
                {zoomLevel}%
              </span>
              <button
                onClick={() => setZoomLevel((z) => Math.min(z + 10, 150))}
                className="p-1 hover:bg-white rounded text-slate-600 transition cursor-pointer"
                title="ពង្រីក"
              >
                <ZoomIn className="w-3.5 h-3.5" />
              </button>
            </div>
          </div>
        </div>

        {/* Quick Margin Controls Sub-bar */}
        {showMarginControls && (
          <div className="mt-2.5 pt-2 border-t border-slate-200 flex flex-wrap items-center justify-between gap-3 text-xs animate-fadeIn bg-slate-50/80 p-2 rounded-lg">
            <div className="flex items-center gap-2 flex-wrap">
              <span className="font-bold text-slate-700 flex items-center gap-1">
                <SlidersHorizontal className="w-3.5 h-3.5 text-indigo-600" />
                <span>គែមក្រដាស (ស.ម):</span>
              </span>
              {/* Presets */}
              <button
                type="button"
                onClick={() => setCustomMargins({ top: 1.5, bottom: 1.5, left: 2.0, right: 1.5 })}
                className="px-2 py-0.5 bg-white border border-slate-300 rounded hover:bg-slate-100 font-medium text-slate-700 cursor-pointer shadow-2xs"
              >
                ស្តង់ដារ (១.៥ / ២.០)
              </button>
              <button
                type="button"
                onClick={() => setCustomMargins({ top: 1.0, bottom: 1.0, left: 1.5, right: 1.0 })}
                className="px-2 py-0.5 bg-white border border-slate-300 rounded hover:bg-slate-100 font-medium text-slate-700 cursor-pointer shadow-2xs"
              >
                តូច (១.០ / ១.៥)
              </button>
              <button
                type="button"
                onClick={() => setCustomMargins({ top: 2.5, bottom: 2.5, left: 2.5, right: 2.0 })}
                className="px-2 py-0.5 bg-white border border-slate-300 rounded hover:bg-slate-100 font-medium text-slate-700 cursor-pointer shadow-2xs"
              >
                ធំ (២.៥ / ២.៥)
              </button>
            </div>

            {/* Inputs */}
            <div className="flex items-center gap-3 flex-wrap">
              <label className="flex items-center gap-1 text-slate-600">
                <span>លើ:</span>
                <input
                  type="number"
                  step="0.1"
                  min="0.2"
                  max="6"
                  value={customMargins.top}
                  onChange={(e) => setCustomMargins((m) => ({ ...m, top: parseFloat(e.target.value) || 0 }))}
                  className="w-14 px-1.5 py-0.5 border border-slate-300 rounded bg-white text-center font-bold text-slate-800"
                />
                <span>cm</span>
              </label>
              <label className="flex items-center gap-1 text-slate-600">
                <span>ក្រោម:</span>
                <input
                  type="number"
                  step="0.1"
                  min="0.2"
                  max="6"
                  value={customMargins.bottom}
                  onChange={(e) => setCustomMargins((m) => ({ ...m, bottom: parseFloat(e.target.value) || 0 }))}
                  className="w-14 px-1.5 py-0.5 border border-slate-300 rounded bg-white text-center font-bold text-slate-800"
                />
                <span>cm</span>
              </label>
              <label className="flex items-center gap-1 text-slate-600">
                <span>ឆ្វេង:</span>
                <input
                  type="number"
                  step="0.1"
                  min="0.2"
                  max="6"
                  value={customMargins.left}
                  onChange={(e) => setCustomMargins((m) => ({ ...m, left: parseFloat(e.target.value) || 0 }))}
                  className="w-14 px-1.5 py-0.5 border border-slate-300 rounded bg-white text-center font-bold text-slate-800"
                />
                <span>cm</span>
              </label>
              <label className="flex items-center gap-1 text-slate-600">
                <span>ស្តាំ:</span>
                <input
                  type="number"
                  step="0.1"
                  min="0.2"
                  max="6"
                  value={customMargins.right}
                  onChange={(e) => setCustomMargins((m) => ({ ...m, right: parseFloat(e.target.value) || 0 }))}
                  className="w-14 px-1.5 py-0.5 border border-slate-300 rounded bg-white text-center font-bold text-slate-800"
                />
                <span>cm</span>
              </label>

              <button
                type="button"
                onClick={() => {
                  setTempMargins({ ...customMargins });
                  setShowPageSetupDialog(true);
                }}
                className="text-blue-600 hover:text-blue-800 font-bold underline cursor-pointer ml-1 text-xs"
              >
                Page Setup...
              </button>
            </div>
          </div>
        )}
      </div>

      {/* Word Page Setup Dialog Component */}
      <WordPageSetupDialog
        show={showPageSetupDialog}
        onClose={() => setShowPageSetupDialog(false)}
        pageSetupTab={pageSetupTab}
        setPageSetupTab={setPageSetupTab}
        tempMargins={tempMargins}
        setTempMargins={setTempMargins}
        gutter={gutter}
        setGutter={setGutter}
        gutterPosition={gutterPosition}
        setGutterPosition={setGutterPosition}
        pageOrientation={pageOrientation}
        setPageOrientation={setPageOrientation}
        paperSize={paperSize}
        setPaperSize={setPaperSize}
        headerMargin={headerMargin}
        setHeaderMargin={setHeaderMargin}
        footerMargin={footerMargin}
        setFooterMargin={setFooterMargin}
        activeTabStop={activeTabStop}
        setActiveTabStop={setActiveTabStop}
        onApply={handleApplyMargins}
      />

      {/* Print Margins & Page Size CSS */}
      <style>{`
        @page {
          size: ${paperSize === 'Letter' ? 'letter' : paperSize === 'Legal' ? 'legal' : 'A4'} ${pageOrientation};
          margin: ${customMargins.top}cm ${customMargins.right}cm ${customMargins.bottom}cm ${customMargins.left}cm;
        }
        @media print {
          body {
            background: white !important;
            -webkit-print-color-adjust: exact;
            print-color-adjust: exact;
          }
          #daily-sticker-report-page-1,
          #daily-sticker-report-page-2 {
            padding: 0 !important;
            margin: 0 !important;
            width: 100% !important;
            min-height: 100% !important;
            box-shadow: none !important;
            border: none !important;
            page-break-after: always;
            break-after: page;
          }
        }
      `}</style>

      {/* Main Preview Container */}
      <div className="flex-1 overflow-auto p-4 md:p-8 flex justify-center items-start">
        <div
          ref={printContainerRef}
          id="daily-sticker-visa-operation-report-pdf"
          style={{
            transform: `scale(${zoomLevel / 100})`,
            transformOrigin: 'top center',
          }}
          className="transition-transform duration-150 flex flex-col gap-8 print:transform-none print:m-0 print:p-0 print:gap-0"
        >
          {/* ============================================================== */}
          {/* PAGE 1 (Header, Staff, Refusals, Operations Matrix & Used 1-4) */}
          {/* ============================================================== */}
          <div
            id="daily-sticker-report-page-1"
            className="w-[210mm] min-h-[297mm] bg-white shadow-xl rounded-sm relative flex flex-col justify-between text-slate-900 border border-slate-200 print:shadow-none print:border-none print:m-0 print:w-[210mm] print:min-h-[297mm] print:break-after-page"
            style={{
              boxSizing: 'border-box',
              paddingTop: `calc(${customMargins.top}cm - 0.5cm)`,
              paddingBottom: `${customMargins.bottom}cm`,
              paddingLeft: `${customMargins.left}cm`,
              paddingRight: `${customMargins.right}cm`,
            }}
          >
            <div>
              {/* TOP HEADER: Ministry / Office on Left, Kingdom on Right */}
              <div className="flex justify-between items-start mb-2">
                {/* Left Side: Ministry, Department & Team */}
                <div className="text-center flex flex-col items-center">
                  {/* Invisible spacer matching height of ព្រះរាជាណាចក្រកម្ពុជា so ក្រសួងមហាផ្ទៃ aligns with ជាតិ សាសនា ព្រះមហាក្សត្រ */}
                  <p
                    className="font-moul whitespace-nowrap invisible select-none pointer-events-none m-0 p-0"
                    aria-hidden="true"
                    style={{
                      fontSize: '10pt',
                      lineHeight: 1.1,
                      fontFamily: "'Khmer OS Muol Light', 'Khmer OS Moul Light', 'Khmer OS Muol', 'Moul', serif",
                    }}
                  >
                    ព្រះរាជាណាចក្រកម្ពុជា
                  </p>
                  <div className="font-moul text-black text-[12pt] leading-tight">
                    ក្រសួងមហាផ្ទៃ
                  </div>
                  <div className="font-moul text-black text-[12pt] leading-tight mt-0.5">
                    អគ្គនាយកដ្ឋានអន្តោប្រវេសន៍
                  </div>
                  <div className="font-moul text-black text-[12pt] leading-tight mt-0.5">
                    នាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត
                  </div>
                  <div className="font-moul text-black text-[12pt] leading-tight mt-0.5">
                    ការិយាល័យទិដ្ឋាការចូល
                  </div>
                  {isEditing ? (
                    <input
                      type="text"
                      value={customTeamHeaderTitle !== null ? customTeamHeaderTitle : fullTeamHeaderTitle}
                      onChange={(e) => setCustomTeamHeaderTitle(e.target.value)}
                      className="font-siemreap font-bold text-blue-900 bg-blue-50/70 border border-blue-400 rounded px-1.5 py-0.5 text-center leading-snug mt-0.5 w-72 focus:outline-none focus:ring-1 focus:ring-blue-500"
                      style={{
                        fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif",
                        fontSize: '11pt',
                      }}
                      placeholder="បញ្ចូលឈ្មោះក្រុម..."
                      title="ចុចដើម្បីកែប្រែឈ្មោះក្រុមក្នុងរូបភាពរបាយការណ៍"
                    />
                  ) : (
                    <div
                      className="font-siemreap font-bold text-black leading-tight mt-0.5"
                      style={{
                        fontFamily: "'Khmer OS Siemreap', 'Siemreap', sans-serif",
                        fontSize: '11pt',
                      }}
                    >
                      {customTeamHeaderTitle !== null ? customTeamHeaderTitle : fullTeamHeaderTitle}
                    </div>
                  )}
                  <div className="my-0.5 flex justify-center">
                    <TacteingLine
                      type={tacteingSettings.type}
                      customImage={tacteingSettings.customImage}
                      width={110}
                      height={10}
                    />
                  </div>
                </div>

                {/* Right Side: Kingdom Motto & Flourish */}
                <div className="text-center flex flex-col items-center">
                  <div className="font-moul text-black text-[12pt] leading-tight">
                    ព្រះរាជាណាចក្រកម្ពុជា
                  </div>
                  <div className="font-moul text-black text-[12pt] leading-tight">
                    ជាតិ សាសនា ព្រះមហាក្សត្រ
                  </div>
                  <div className="my-0.5 flex justify-center">
                    <TacteingLine
                      type={tacteingSettings.type}
                      customImage={tacteingSettings.customImage}
                      width={130}
                      height={13}
                    />
                  </div>
                </div>
              </div>

              {/* CENTER TITLE */}
              <div className="text-center mb-2">
                <h2 className="font-moul text-slate-900 text-[12pt] tracking-wide mb-0">
                  របាយការណ៍
                </h2>
                <div className="font-siemreap font-bold text-slate-900 text-[11pt] mb-0">
                  កិច្ចប្រតិបត្តិការណ៍ប្រចាំ{solarHeaderDateStr}
                </div>
              </div>

              {/* SECTION I: Staff Management (ការងារគ្រប់គ្រងកម្លាំង) */}
              <div className="mb-2">
                <h3 className="font-siemreap font-bold text-[10.5pt] text-slate-900 mb-1">
                  I. ការងារគ្រប់គ្រងកម្លាំង
                </h3>
                <div className="pl-6 space-y-0.5 font-siemreap text-[10pt] text-slate-800">
                  <div className="flex items-center gap-2">
                    <span className="w-56 inline-block">១. ចំនួនរួម</span>
                    <span className="font-bold">៖</span>
                    {isEditing ? (
                      <input
                        type="text"
                        value={staffInfo.total}
                        onChange={(e) => setStaffInfo({ ...staffInfo, total: e.target.value })}
                        className="border border-slate-300 rounded px-1.5 py-0.5 w-28 text-center text-xs"
                      />
                    ) : (
                      renderStaffDisplay(staffInfo.total)
                    )}
                  </div>
                  <div className="flex items-center gap-2">
                    <span className="w-56 inline-block">២. ចំនួនវត្តមានធ្វើការងារ</span>
                    <span className="font-bold">៖</span>
                    {isEditing ? (
                      <input
                        type="text"
                        value={staffInfo.present}
                        onChange={(e) => setStaffInfo({ ...staffInfo, present: e.target.value })}
                        className="border border-slate-300 rounded px-1.5 py-0.5 w-28 text-center text-xs"
                      />
                    ) : (
                      renderStaffDisplay(staffInfo.present)
                    )}
                  </div>
                  <div className="flex items-center gap-2">
                    <span className="w-56 inline-block">៣. ចំនួនអវត្តមាន មានមូលហេតុ</span>
                    <span className="font-bold">៖</span>
                    {isEditing ? (
                      <input
                        type="text"
                        value={staffInfo.absentWithReason}
                        onChange={(e) => setStaffInfo({ ...staffInfo, absentWithReason: e.target.value })}
                        className="border border-slate-300 rounded px-1.5 py-0.5 w-28 text-center text-xs"
                      />
                    ) : (
                      renderStaffDisplay(staffInfo.absentWithReason)
                    )}
                  </div>
                  <div className="flex items-center gap-2">
                    <span className="w-56 inline-block">៤. ចំនួនអវត្តមាន មិនមានមូលហេតុ</span>
                    <span className="font-bold">៖</span>
                    {isEditing ? (
                      <input
                        type="text"
                        value={staffInfo.absentWithoutReason}
                        onChange={(e) => setStaffInfo({ ...staffInfo, absentWithoutReason: e.target.value })}
                        className="border border-slate-300 rounded px-1.5 py-0.5 w-28 text-center text-xs"
                      />
                    ) : (
                      renderStaffDisplay(staffInfo.absentWithoutReason)
                    )}
                  </div>
                </div>
              </div>

              {/* SECTION II: Situation (សភាពការណ៍) & Table of Refused Foreigners */}
              <div className="mb-2">
                <h3 className="font-siemreap font-bold text-[11pt] text-slate-900 mb-0.5">
                  II. សភាពការណ៍
                </h3>
                <div className="pl-3 mb-1 flex items-center gap-2 font-siemreap text-[10.5pt] text-slate-800">
                  <span>- ការបដិសេធផ្តល់ទិដ្ឋាការ និងបញ្ជូនជនបរទេស ចំនួន ៖</span>
                  {isEditing ? (
                    <input
                      type="text"
                      value={refusalSummaryText}
                      onChange={(e) => setRefusalSummaryText(e.target.value)}
                      className="border border-slate-300 rounded px-1.5 py-0.5 text-xs w-48 font-siemreap"
                    />
                  ) : (
                    <span className="font-bold">{refusalSummaryText}</span>
                  )}
                </div>

                {/* Table of Refused Foreigners */}
                <div className="overflow-x-auto">
                  <table className="w-full border-collapse border border-black text-[9.5pt] font-siemreap">
                    <thead>
                      <tr className="bg-slate-50 font-bold text-center">
                        <th className="border border-black px-0.5 py-1 w-6 text-center">ល.រ</th>
                        <th className="border border-black px-1.5 py-1 text-center min-w-[120px]">ឈ្មោះ</th>
                        <th className="border border-black px-0.5 py-1 w-9 text-center">ភេទ</th>
                        <th className="border border-black px-0.5 py-1 w-[72px] text-center">ឆ្នាំកំណើត</th>
                        <th className="border border-black px-0.5 py-1 w-11 text-center">សញ្ជាតិ</th>
                        <th className="border border-black px-0.5 py-1 w-[80px] text-center">លិខិតឆ្លងដែន</th>
                        <th className="border border-black px-1.5 py-1 text-center">មូលហេតុ</th>
                        {isEditing && <th className="border border-black px-0.5 py-1 w-8 text-center print:hidden">លុប</th>}
                      </tr>
                    </thead>
                    <tbody>
                      {(() => {
                        const activeRows = isEditing
                          ? refusedForeigners
                          : refusedForeigners.filter(
                              (r) =>
                                (r.name && r.name.trim() !== '' && r.name.trim() !== '---') ||
                                (r.passportNo && r.passportNo.trim() !== '' && r.passportNo.trim() !== '---')
                            );

                        if (activeRows.length === 0) {
                          return [1].map((num) => (
                            <tr key={`empty-rf-${num}`} className="h-6.5 text-center">
                              <td className="border border-black px-0.5 py-0.5 font-times">{num}</td>
                              <td className="border border-black px-1.5 py-0.5 text-left"></td>
                              <td className="border border-black px-0.5 py-0.5"></td>
                              <td className="border border-black px-0.5 py-0.5 font-times"></td>
                              <td className="border border-black px-0.5 py-0.5 font-times font-bold uppercase"></td>
                              <td className="border border-black px-0.5 py-0.5 font-times font-bold"></td>
                              <td className="border border-black px-1.5 py-0.5 text-left"></td>
                              {isEditing && <td className="border border-black px-0.5 py-0.5 print:hidden"></td>}
                            </tr>
                          ));
                        }

                        return activeRows.map((row, idx) => (
                          <tr key={row.id || idx} className="h-6.5 text-center">
                            <td className="border border-black px-0.5 py-0.5 font-times">{idx + 1}</td>
                            <td className="border border-black px-1.5 py-0.5 text-left font-times font-bold uppercase">
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={row.name}
                                  onChange={(e) => handleUpdateRefusedRow(row.id, 'name', e.target.value)}
                                  className="w-full border-none p-0 text-xs focus:ring-0 font-times font-bold uppercase"
                                />
                              ) : (
                                row.name
                              )}
                            </td>
                            <td className="border border-black px-0.5 py-0.5">
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={row.gender}
                                  onChange={(e) => handleUpdateRefusedRow(row.id, 'gender', e.target.value)}
                                  className="w-full border-none p-0 text-xs text-center focus:ring-0"
                                />
                              ) : (
                                row.gender
                              )}
                            </td>
                            <td className="border border-black px-0.5 py-0.5 font-times">
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={row.birthYear}
                                  onChange={(e) => handleUpdateRefusedRow(row.id, 'birthYear', e.target.value)}
                                  className="w-full border-none p-0 text-xs text-center focus:ring-0 font-times"
                                />
                              ) : (
                                row.birthYear
                              )}
                            </td>
                            <td className="border border-black px-0.5 py-0.5 font-times font-bold uppercase">
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={row.nationality}
                                  onChange={(e) => handleUpdateRefusedRow(row.id, 'nationality', e.target.value)}
                                  className="w-full border-none p-0 text-xs text-center focus:ring-0 font-times font-bold uppercase"
                                />
                              ) : (
                                row.nationality
                              )}
                            </td>
                            <td className="border border-black px-0.5 py-0.5 font-times font-bold">
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={row.passportNo}
                                  onChange={(e) => handleUpdateRefusedRow(row.id, 'passportNo', e.target.value)}
                                  className="w-full border-none p-0 text-xs text-center focus:ring-0 font-times font-bold"
                                />
                              ) : (
                                row.passportNo
                              )}
                            </td>
                            <td className="border border-black px-1.5 py-0.5 text-left">
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={row.reason}
                                  onChange={(e) => handleUpdateRefusedRow(row.id, 'reason', e.target.value)}
                                  className="w-full border-none p-0 text-xs focus:ring-0"
                                />
                              ) : (
                                row.reason
                              )}
                            </td>
                            {isEditing && (
                              <td className="border border-black px-1 py-0.5 text-center print:hidden">
                                <button
                                  onClick={() => handleRemoveRefusedRow(row.id)}
                                  className="text-red-500 hover:text-red-700"
                                >
                                  <Trash2 className="w-3.5 h-3.5 mx-auto" />
                                </button>
                              </td>
                            )}
                          </tr>
                        ));
                      })()}
                    </tbody>
                  </table>
                  {isEditing && (
                    <div className="mt-1 flex justify-start print:hidden">
                      <button
                        onClick={handleAddRefusedRow}
                        className="flex items-center gap-1 text-[11px] text-blue-600 hover:text-blue-800 font-semibold"
                      >
                        <Plus className="w-3 h-3" /> បន្ថែមជួរជនបរទេស
                      </button>
                    </div>
                  )}
                </div>
              </div>

              {/* SECTION III: Operations (កិច្ចប្រតិបត្តិការណ៍) */}
              <div className="mb-3">
                <h3 className="font-siemreap font-bold text-[11pt] text-slate-900 mb-1">
                  III. កិច្ចប្រតិបត្តិការណ៍
                </h3>
                <div className="pl-3 mb-1.5 font-siemreap text-[10.5pt] text-slate-800">
                  - ការផ្តល់ទិដ្ឋាការជូនជនបរទេស និងប្រើប្រាស់សន្លឹកទិដ្ឋាការ
                </div>

                {/* Table 1: Matrix of 13 Visa Types */}
                <div className="overflow-x-auto mb-3">
                  <table className="w-full border-collapse border border-black text-[8.5pt] font-siemreap text-center">
                    <thead>
                      <tr className="bg-slate-50 font-bold">
                        <th className="border border-black px-1 py-1 w-24">ប្រភេទ</th>
                        {VISA_TYPE_KEYS.map((k) => (
                          <th key={k} className="border border-black px-0.5 py-1 font-times w-8">
                            {k}
                          </th>
                        ))}
                        <th className="border border-black px-1 py-1 w-14 font-bold">សរុប</th>
                      </tr>
                    </thead>
                    <tbody className="font-times text-center">
                      {/* Row 1: សល់ពីថ្ងៃចាស់ */}
                      <tr className="h-6">
                        <td className="border border-black px-1 py-0.5 font-siemreap font-semibold text-left whitespace-nowrap">
                          សល់ពីថ្ងៃចាស់
                        </td>
                        {VISA_TYPE_KEYS.map((k) => (
                          <td key={k} className="border border-black px-0.5 py-0.5 font-times">
                            {isEditing ? (
                              <input
                                type="text"
                                value={matrixData.opening[k]}
                                onChange={(e) => handleMatrixChange('opening', k, e.target.value)}
                                className="w-full text-center text-xs p-0 border-none focus:ring-0 font-times"
                              />
                            ) : (
                              matrixData.opening[k] > 0 ? matrixData.opening[k].toLocaleString() : '-'
                            )}
                          </td>
                        ))}
                        <td className="border border-black px-1 py-0.5 font-bold bg-slate-50 font-times">
                          {rowTotals.opening > 0 ? rowTotals.opening.toLocaleString() : '-'}
                        </td>
                      </tr>

                      {/* Row 2: បើកថ្មី */}
                      <tr className="h-6">
                        <td
                          className="border border-black px-1 py-0.5 font-siemreap font-semibold text-left whitespace-nowrap cursor-help"
                          title="បើកថ្មី = ទទួលពីការិយាល័យ ឬផ្ទេរពីក្រុម"
                        >
                          បើកថ្មី
                        </td>
                        {VISA_TYPE_KEYS.map((k) => (
                          <td key={k} className="border border-black px-0.5 py-0.5 font-times">
                            {isEditing ? (
                              <input
                                type="text"
                                value={matrixData.received[k]}
                                onChange={(e) => handleMatrixChange('received', k, e.target.value)}
                                className="w-full text-center text-xs p-0 border-none focus:ring-0 font-times"
                              />
                            ) : (
                              matrixData.received[k] > 0 ? matrixData.received[k].toLocaleString() : '-'
                            )}
                          </td>
                        ))}
                        <td className="border border-black px-1 py-0.5 font-bold bg-slate-50 font-times">
                          {rowTotals.received > 0 ? rowTotals.received.toLocaleString() : '-'}
                        </td>
                      </tr>

                      {/* Row 3: ផ្ទេរ/ខូច/ខ្វះ/បង្វិល */}
                      <tr className="h-6">
                        <td
                          className="border border-black px-1 py-0.5 font-siemreap font-semibold text-left whitespace-nowrap cursor-help"
                          title="ផ្ទេរ/ខូច/ខ្វះ/បង្វិល = ផ្ទេរការប្រើប្រាស់ + ទិដ្ឋាការខូចក្រុម + ទិដ្ឋាការខ្វះក្រុម + ទិដ្ឋាការបង្វិលទៅក២"
                        >
                          ផ្ទេរ/ខូច/ខ្វះ/បង្វិល
                        </td>
                        {VISA_TYPE_KEYS.map((k) => (
                          <td key={k} className="border border-black px-0.5 py-0.5 font-times">
                            {isEditing ? (
                              <input
                                type="text"
                                value={matrixData.damaged[k]}
                                onChange={(e) => handleMatrixChange('damaged', k, e.target.value)}
                                className="w-full text-center text-xs p-0 border-none focus:ring-0 font-times"
                              />
                            ) : (
                              matrixData.damaged[k] > 0 ? matrixData.damaged[k].toLocaleString() : '-'
                            )}
                          </td>
                        ))}
                        <td className="border border-black px-1 py-0.5 font-bold bg-slate-50 font-times">
                          {rowTotals.damaged > 0 ? rowTotals.damaged.toLocaleString() : '-'}
                        </td>
                      </tr>

                      {/* Row 4: ប្រើប្រាស់ */}
                      <tr className="h-6">
                        <td className="border border-black px-1 py-0.5 font-siemreap font-semibold text-left whitespace-nowrap">
                          ប្រើប្រាស់
                        </td>
                        {VISA_TYPE_KEYS.map((k) => (
                          <td key={k} className="border border-black px-0.5 py-0.5 font-times">
                            {isEditing ? (
                              <input
                                type="text"
                                value={matrixData.used[k]}
                                onChange={(e) => handleMatrixChange('used', k, e.target.value)}
                                className="w-full text-center text-xs p-0 border-none focus:ring-0 font-bold text-blue-800 font-times"
                              />
                            ) : (
                              effectiveUsed[k] > 0 ? effectiveUsed[k].toLocaleString() : '-'
                            )}
                          </td>
                        ))}
                        <td className="border border-black px-1 py-0.5 font-bold bg-slate-50 text-blue-900 font-times">
                          {rowTotals.used > 0 ? rowTotals.used.toLocaleString() : '-'}
                        </td>
                      </tr>

                      {/* Row 5: សរុបរួម */}
                      <tr className="h-6 font-bold bg-slate-100 font-times">
                        <td className="border border-black px-1 py-0.5 font-siemreap text-left whitespace-nowrap">
                          សរុបរួម
                        </td>
                        {VISA_TYPE_KEYS.map((k) => (
                          <td key={k} className="border border-black px-0.5 py-0.5 font-times">
                            {calculatedEndingBalance[k] > 0
                              ? calculatedEndingBalance[k].toLocaleString()
                              : '-'}
                          </td>
                        ))}
                        <td className="border border-black px-1 py-0.5 bg-slate-200 font-times">
                          {rowTotals.ending > 0 ? rowTotals.ending.toLocaleString() : '-'}
                        </td>
                      </tr>
                    </tbody>
                  </table>
                </div>

                {/* Subtitle: Used Visas Breakdown */}
                <div className="pl-3 mb-1.5 font-siemreap text-[10.5pt] text-slate-800">
                  - ទិដ្ឋាការដែលបានប្រើប្រាស់រួច
                </div>

                {/* Table 2: Part 1 (Rows 1 to 4 on Page 1) */}
                <div className="overflow-x-auto">
                  <table className="w-full border-collapse border border-black text-[9pt] font-siemreap">
                    <thead>
                      <tr className="bg-slate-50 font-bold text-center">
                        <th className="border border-black px-1 py-1 w-10">ល.រ</th>
                        <th className="border border-black px-1.5 py-1 w-32">ប្រភេទទិដ្ឋាការ</th>
                        <th className="border border-black px-1.5 py-1 w-20">ចំនួន</th>
                        <th className="border border-black px-1.5 py-1 w-36">ចាប់ពីលេខ</th>
                        <th className="border border-black px-1.5 py-1 w-36">ដល់លេខ</th>
                        <th className="border border-black px-2 py-1">ផ្សេងៗ</th>
                        {isEditing && <th className="border border-black px-1 py-1 w-10 print:hidden">លុប</th>}
                      </tr>
                    </thead>
                    <tbody className="font-times text-center">
                      {page1UsedRows.length === 0 ? (
                        <tr className="h-6.5 font-times text-center">
                          <td className="border border-black px-1 py-0.5 font-times">-</td>
                          <td className="border border-black px-1 py-0.5 font-times">-</td>
                          <td className="border border-black px-1 py-0.5 font-times">-</td>
                          <td className="border border-black px-1 py-0.5 font-times">-</td>
                          <td className="border border-black px-1 py-0.5 font-times">-</td>
                          <td className="border border-black px-1 py-0.5 font-times"></td>
                          {isEditing && <td className="border border-black px-1 py-0.5">-</td>}
                        </tr>
                      ) : (
                        page1UsedRows.map((row) => (
                          <tr key={row.id} className="h-6.5 font-times">
                            <td className="border border-black px-1 py-0.5 font-times">{row.no}</td>
                            <td className="border border-black px-1 py-0.5 font-times">
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={row.visaType}
                                  onChange={(e) => handleUpdateUsedRow(row.id, 'visaType', e.target.value)}
                                  className="w-full border-none p-0 text-xs text-center focus:ring-0 font-times"
                                />
                              ) : (
                                row.visaType
                              )}
                            </td>
                            <td className="border border-black px-1 py-0.5 font-semibold font-times">
                              {isEditing ? (
                                <input
                                  type="number"
                                  value={row.count}
                                  onChange={(e) => handleUpdateUsedRow(row.id, 'count', Number(e.target.value) || 0)}
                                  className="w-full border-none p-0 text-xs text-center focus:ring-0 font-times"
                                />
                              ) : (
                                row.count > 0 ? row.count.toLocaleString() : '-'
                              )}
                            </td>
                            <td className="border border-black px-1 py-0.5 font-times">
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={row.fromSerial}
                                  onChange={(e) => handleUpdateUsedRow(row.id, 'fromSerial', e.target.value)}
                                  className="w-full border-none p-0 text-xs text-center focus:ring-0 font-times"
                                />
                              ) : (
                                row.fromSerial || '-'
                              )}
                            </td>
                            <td className="border border-black px-1 py-0.5 font-times">
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={row.toSerial}
                                  onChange={(e) => handleUpdateUsedRow(row.id, 'toSerial', e.target.value)}
                                  className="w-full border-none p-0 text-xs text-center focus:ring-0 font-times"
                                />
                              ) : (
                                row.toSerial || '-'
                              )}
                            </td>
                            <td className="border border-black px-1 py-0.5 text-left font-siemreap">
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={row.note}
                                  onChange={(e) => handleUpdateUsedRow(row.id, 'note', e.target.value)}
                                  className="w-full border-none p-0 text-xs focus:ring-0"
                                />
                              ) : (
                                row.note && !row.note.includes('កូដចាស់') && row.note.trim() !== '-' ? row.note : ''
                              )}
                            </td>
                            {isEditing && (
                              <td className="border border-black px-1 py-0.5 text-center print:hidden">
                                <button
                                  onClick={() => handleRemoveUsedRow(row.id)}
                                  className="text-red-500 hover:text-red-700"
                                >
                                  <Trash2 className="w-3.5 h-3.5 mx-auto" />
                                </button>
                              </td>
                            )}
                          </tr>
                        ))
                      )}

                      {/* CONDITIONAL SUMMARY ROW inside Page 1 table if NOT using Page 2 */}
                      {page2UsedRows.length === 0 && (
                        <tr className="h-7 bg-slate-100 font-bold">
                          <td colSpan={2} className="border border-black px-2 py-1 text-center font-siemreap font-bold">
                            សរុបការប្រើប្រាស់សន្លឹកទិដ្ឋាការ
                          </td>
                          <td className="border border-black px-1.5 py-1 text-center font-times text-blue-900 text-[10pt]">
                            {totalUsedDetailsCount > 0 ? totalUsedDetailsCount.toLocaleString() : '-'}
                          </td>
                          <td colSpan={isEditing ? 4 : 3} className="border border-black px-2 py-1 text-left font-siemreap">
                            សន្លឹក
                          </td>
                        </tr>
                      )}
                    </tbody>
                  </table>
                  {isEditing && page2UsedRows.length === 0 && (
                    <div className="mt-1 flex justify-start print:hidden">
                      <button
                        onClick={handleAddUsedRow}
                        className="flex items-center gap-1 text-[11px] text-blue-600 hover:text-blue-800 font-semibold"
                      >
                        <Plus className="w-3 h-3" /> បន្ថែមជួរទិដ្ឋាការប្រើប្រាស់
                      </button>
                    </div>
                  )}
                </div>
              </div>

              {/* CONDITIONAL SECTION IV, V, VI + SIGN-OFF inside Page 1 if NOT using Page 2 */}
              {page2UsedRows.length === 0 && (
                <>
                  {/* SECTION IV, V, VI */}
                  <div className="space-y-0.5 mt-2 mb-2 font-siemreap text-[10pt]">
                    {/* IV. បញ្ហាប្រឈម */}
                    <div className="flex items-start gap-2">
                      <span className="font-bold whitespace-nowrap">IV. បញ្ហាប្រឈម ៖</span>
                      {isEditing ? (
                        <input
                          type="text"
                          value={challengeText}
                          onChange={(e) => setChallengeText(e.target.value)}
                          className="border border-slate-300 rounded px-2 py-0.5 w-full text-xs"
                        />
                      ) : (
                        <span>{challengeText}</span>
                      )}
                    </div>

                    {/* V. សំណូមពរ */}
                    <div className="flex items-start gap-2">
                      <span className="font-bold whitespace-nowrap">V. សំណូមពរ ៖</span>
                      {isEditing ? (
                        <input
                          type="text"
                          value={requestText}
                          onChange={(e) => setRequestText(e.target.value)}
                          className="border border-slate-300 rounded px-2 py-0.5 w-full text-xs"
                        />
                      ) : (
                        <span>{requestText}</span>
                      )}
                    </div>

                    {/* VI. ទិសដៅបន្ត */}
                    <div className="flex items-start gap-2">
                      <span className="font-bold whitespace-nowrap">VI. ទិសដៅបន្ត ៖</span>
                      {isEditing ? (
                        <input
                          type="text"
                          value={directionText}
                          onChange={(e) => setDirectionText(e.target.value)}
                          className="border border-slate-300 rounded px-2 py-0.5 w-full text-xs"
                        />
                      ) : (
                        <span>{directionText}</span>
                      )}
                    </div>
                  </div>

                  {/* SIGN-OFF BLOCK (Bottom Right) */}
                  <div className="flex justify-end mt-4">
                    <div className="text-center font-siemreap w-96 flex flex-col items-center">
                      <div className="text-[12pt] whitespace-nowrap text-slate-800 mb-0.5">
                        {lunarDateStr}
                      </div>
                      <div className="text-[12pt] whitespace-nowrap font-medium text-slate-800 mb-2">
                        {solarSignDateStr}
                      </div>

                      <div className="font-moul text-slate-900 text-[11.5pt] mb-10">
                        {isEditing ? (
                          <input
                            type="text"
                            value={signatoryTitle}
                            onChange={(e) => setSignatoryTitle(e.target.value)}
                            className="text-center border border-slate-300 rounded px-1.5 py-0.5 text-xs w-64"
                          />
                        ) : (
                          signatoryTitle
                        )}
                      </div>

                      {/* Signee Name (If entered) */}
                      {signatoryName ? (
                        <div className="font-moul text-slate-900 text-[11pt]">
                          {signatoryName}
                        </div>
                      ) : isEditing ? (
                        <input
                          type="text"
                          placeholder="[ឈ្មោះអ្នកចុះហត្ថលេខា]"
                          value={signatoryName}
                          onChange={(e) => setSignatoryName(e.target.value)}
                          className="text-center border border-slate-300 rounded px-1.5 py-0.5 text-xs w-64 font-moul"
                        />
                      ) : null}
                    </div>
                  </div>
                </>
              )}
            </div>

            {/* Page 1 Footer */}
            <div className="pt-2 flex justify-between items-center text-[9pt] font-times text-slate-500 border-t border-slate-200 print:border-none">
              <span className="font-siemreap text-[8pt] text-slate-400">
                {fullTeamHeaderTitle} — របាយការណ៍សន្លឹកទិដ្ឋាការ
              </span>
              <span className="font-bold">1</span>
            </div>
          </div>

          {/* ============================================================== */}
          {/* PAGE 2 (Continuation of Table 2, IV, V, VI, Sign-off Block)     */}
          {/* ============================================================== */}
          {page2UsedRows.length > 0 && (
            <div
              id="daily-sticker-report-page-2"
              className="w-[210mm] min-h-[297mm] bg-white shadow-xl rounded-sm relative flex flex-col justify-between text-slate-900 border border-slate-200 print:shadow-none print:border-none print:m-0 print:w-[210mm] print:min-h-[297mm]"
              style={{
                boxSizing: 'border-box',
                paddingTop: `${customMargins.top}cm`,
                paddingBottom: `${customMargins.bottom}cm`,
                paddingLeft: `${customMargins.left}cm`,
                paddingRight: `${customMargins.right}cm`,
              }}
            >
              <div>
                {/* Table 2: Part 2 (Rows 5 to 13 + Summary Total Row) */}
                <div className="mb-6">
                  <div className="overflow-x-auto">
                    <table className="w-full border-collapse border border-black text-[9pt] font-siemreap">
                      <thead>
                        <tr className="bg-slate-50 font-bold text-center">
                          <th className="border border-black px-1 py-1 w-10">ល.រ</th>
                          <th className="border border-black px-1.5 py-1 w-32">ប្រភេទទិដ្ឋាការ</th>
                          <th className="border border-black px-1.5 py-1 w-20">ចំនួន</th>
                          <th className="border border-black px-1.5 py-1 w-36">ចាប់ពីលេខ</th>
                          <th className="border border-black px-1.5 py-1 w-36">ដល់លេខ</th>
                          <th className="border border-black px-2 py-1">ផ្សេងៗ</th>
                          {isEditing && <th className="border border-black px-1 py-1 w-10 print:hidden">លុប</th>}
                        </tr>
                      </thead>
                      <tbody className="font-times text-center">
                        {page2UsedRows.map((row) => (
                          <tr key={row.id} className="h-6.5 font-times">
                            <td className="border border-black px-1 py-0.5 font-times">{row.no}</td>
                            <td className="border border-black px-1 py-0.5 font-times">
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={row.visaType}
                                  onChange={(e) => handleUpdateUsedRow(row.id, 'visaType', e.target.value)}
                                  className="w-full border-none p-0 text-xs text-center focus:ring-0 font-times"
                                />
                              ) : (
                                row.visaType
                              )}
                            </td>
                            <td className="border border-black px-1 py-0.5 font-semibold font-times">
                              {isEditing ? (
                                <input
                                  type="number"
                                  value={row.count}
                                  onChange={(e) => handleUpdateUsedRow(row.id, 'count', Number(e.target.value) || 0)}
                                  className="w-full border-none p-0 text-xs text-center focus:ring-0 font-times"
                                />
                              ) : (
                                row.count > 0 ? row.count.toLocaleString() : '-'
                              )}
                            </td>
                            <td className="border border-black px-1 py-0.5 font-times">
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={row.fromSerial}
                                  onChange={(e) => handleUpdateUsedRow(row.id, 'fromSerial', e.target.value)}
                                  className="w-full border-none p-0 text-xs text-center focus:ring-0 font-times"
                                />
                              ) : (
                                row.fromSerial || '-'
                              )}
                            </td>
                            <td className="border border-black px-1 py-0.5 font-times">
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={row.toSerial}
                                  onChange={(e) => handleUpdateUsedRow(row.id, 'toSerial', e.target.value)}
                                  className="w-full border-none p-0 text-xs text-center focus:ring-0 font-times"
                                />
                              ) : (
                                row.toSerial || '-'
                              )}
                            </td>
                            <td className="border border-black px-1 py-0.5 text-left font-siemreap">
                              {isEditing ? (
                                <input
                                  type="text"
                                  value={row.note}
                                  onChange={(e) => handleUpdateUsedRow(row.id, 'note', e.target.value)}
                                  className="w-full border-none p-0 text-xs focus:ring-0"
                                />
                              ) : (
                                row.note && !row.note.includes('កូដចាស់') && row.note.trim() !== '-' ? row.note : ''
                              )}
                            </td>
                            {isEditing && (
                              <td className="border border-black px-1 py-0.5 text-center print:hidden">
                                <button
                                  onClick={() => handleRemoveUsedRow(row.id)}
                                  className="text-red-500 hover:text-red-700"
                                >
                                  <Trash2 className="w-3.5 h-3.5 mx-auto" />
                                </button>
                              </td>
                            )}
                          </tr>
                        ))}

                        {/* SUMMARY ROW: សរុបការប្រើប្រាស់សន្លឹកទិដ្ឋាការ */}
                        <tr className="h-7 bg-slate-100 font-bold">
                          <td colSpan={2} className="border border-black px-2 py-1 text-center font-siemreap font-bold">
                            សរុបការប្រើប្រាស់សន្លឹកទិដ្ឋាការ
                          </td>
                          <td className="border border-black px-1.5 py-1 text-center font-times text-blue-900 text-[10pt]">
                            {totalUsedDetailsCount > 0 ? totalUsedDetailsCount.toLocaleString() : '-'}
                          </td>
                          <td colSpan={isEditing ? 4 : 3} className="border border-black px-2 py-1 text-left font-siemreap">
                            សន្លឹក
                          </td>
                        </tr>
                      </tbody>
                    </table>
                    {isEditing && (
                      <div className="mt-1 flex justify-start print:hidden">
                        <button
                          onClick={handleAddUsedRow}
                          className="flex items-center gap-1 text-[11px] text-blue-600 hover:text-blue-800 font-semibold"
                        >
                          <Plus className="w-3 h-3" /> បន្ថែមជួរទិដ្ឋាការប្រើប្រាស់
                        </button>
                      </div>
                    )}
                  </div>
                </div>

                {/* SECTION IV, V, VI */}
                <div className="space-y-0.5 -mt-2 mb-2 font-siemreap text-[10pt]">
                  {/* IV. បញ្ហាប្រឈម */}
                  <div className="flex items-start gap-2">
                    <span className="font-bold whitespace-nowrap">IV. បញ្ហាប្រឈម ៖</span>
                    {isEditing ? (
                      <input
                        type="text"
                        value={challengeText}
                        onChange={(e) => setChallengeText(e.target.value)}
                        className="border border-slate-300 rounded px-2 py-0.5 w-full text-xs"
                      />
                    ) : (
                      <span>{challengeText}</span>
                    )}
                  </div>

                  {/* V. សំណូមពរ */}
                  <div className="flex items-start gap-2">
                    <span className="font-bold whitespace-nowrap">V. សំណូមពរ ៖</span>
                    {isEditing ? (
                      <input
                        type="text"
                        value={requestText}
                        onChange={(e) => setRequestText(e.target.value)}
                        className="border border-slate-300 rounded px-2 py-0.5 w-full text-xs"
                      />
                    ) : (
                      <span>{requestText}</span>
                    )}
                  </div>

                  {/* VI. ទិសដៅបន្ត */}
                  <div className="flex items-start gap-2">
                    <span className="font-bold whitespace-nowrap">VI. ទិសដៅបន្ត ៖</span>
                    {isEditing ? (
                      <input
                        type="text"
                        value={directionText}
                        onChange={(e) => setDirectionText(e.target.value)}
                        className="border border-slate-300 rounded px-2 py-0.5 w-full text-xs"
                      />
                    ) : (
                      <span>{directionText}</span>
                    )}
                  </div>
                </div>

                {/* SIGN-OFF BLOCK (Bottom Right) */}
                <div className="flex justify-end mt-1">
                  <div className="text-center font-siemreap w-96 flex flex-col items-center">
                    <div className="text-[12pt] whitespace-nowrap text-slate-800 mb-0.5">
                      {lunarDateStr}
                    </div>
                    <div className="text-[12pt] whitespace-nowrap font-medium text-slate-800 mb-1">
                      {solarSignDateStr}
                    </div>

                    <div className="font-moul text-slate-900 text-[11.5pt] mb-8">
                      {isEditing ? (
                        <input
                          type="text"
                          value={signatoryTitle}
                          onChange={(e) => setSignatoryTitle(e.target.value)}
                          className="text-center border border-slate-300 rounded px-1.5 py-0.5 text-xs w-64"
                        />
                      ) : (
                        signatoryTitle
                      )}
                    </div>

                    {/* Signee Name (If entered) */}
                    {signatoryName ? (
                      <div className="font-moul text-slate-900 text-[11pt]">
                        {signatoryName}
                      </div>
                    ) : isEditing ? (
                      <input
                        type="text"
                        placeholder="[ឈ្មោះអ្នកចុះហត្ថលេខា]"
                        value={signatoryName}
                        onChange={(e) => setSignatoryName(e.target.value)}
                        className="text-center border border-slate-300 rounded px-1.5 py-0.5 text-xs w-64 font-moul"
                      />
                    ) : null}
                  </div>
                </div>
              </div>

              {/* Page 2 Footer */}
              <div className="pt-2 flex justify-between items-center text-[9pt] font-times text-slate-500 border-t border-slate-200 print:border-none">
                <span className="font-siemreap text-[8pt] text-slate-400">
                  {fullTeamHeaderTitle} — របាយការណ៍សន្លឹកទិដ្ឋាការ
                </span>
                <span className="font-bold">2</span>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  );
};

export default DailyStickerVisaOperationReport;
