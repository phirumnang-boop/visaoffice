import React, { useState, useEffect, useMemo } from 'react';
import {
  Users,
  Plane,
  MapPin,
  Anchor,
  Play,
  Pause,
  ChevronLeft,
  ChevronRight,
  RotateCcw,
  Search,
  Building2,
  CheckCircle2,
  Sparkles,
} from 'lucide-react';
import { CategoriesState, Officer, StockRecord } from '../types';
import { normalizeTeamName } from '../utils/teamNormalization';
import { OFFICIAL_TEAM_FULL_NAMES } from './TeamCodeSearch';

export type PortCategoryType = 'air' | 'border' | 'port';

interface VisaTeamItem {
  id: string;
  robokName: string;
  fullName: string;
  category: PortCategoryType;
  categoryLabel: string;
  categoryBadgeBg: string;
  categoryBadgeText: string;
  officerCount: number;
  useSticker: boolean;
  useEvisa: boolean;
}

interface VisaTeamsCategorizedTableProps {
  categories?: CategoriesState;
  officers?: Officer[];
  stockRecords?: StockRecord[];
  onNavigate?: (page: string, catType?: any) => void;
}

// Category Configuration & Sequence for the animated rotation cycle
const CATEGORY_SEQUENCE: {
  key: PortCategoryType;
  stepNum: number;
  labelKh: string;
  shortLabel: string;
  icon: typeof Plane;
  accentColor: string;
  borderColor: string;
  bgGradient: string;
  activeTabClass: string;
  badgeClass: string;
  durationMs: number;
}[] = [
  {
    key: 'air',
    stepNum: 1,
    labelKh: 'ច្រកទ្វារអាកាសយានដ្ឋានអន្តរជាតិ (អាកាស)',
    shortLabel: 'អាកាស',
    icon: Plane,
    accentColor: 'text-sky-600',
    borderColor: 'border-sky-300',
    bgGradient: 'from-sky-50 via-blue-50/40 to-white',
    activeTabClass: 'bg-sky-600 text-white border-sky-600 shadow-md ring-2 ring-sky-200',
    badgeClass: 'bg-sky-100 text-sky-900 border-sky-200',
    durationMs: 8000, // 8 seconds for airport teams
  },
  {
    key: 'border',
    stepNum: 2,
    labelKh: 'ច្រកទ្វារព្រំដែនអន្តរជាតិ (ព្រំដែន)',
    shortLabel: 'ព្រំដែន',
    icon: MapPin,
    accentColor: 'text-emerald-700',
    borderColor: 'border-emerald-300',
    bgGradient: 'from-emerald-50 via-teal-50/40 to-white',
    activeTabClass: 'bg-emerald-700 text-white border-emerald-700 shadow-md ring-2 ring-emerald-200',
    badgeClass: 'bg-emerald-100 text-emerald-900 border-emerald-200',
    durationMs: 16000, // 16 seconds for 20 border teams
  },
  {
    key: 'port',
    stepNum: 3,
    labelKh: 'ច្រកទ្វារកំពង់ផែអន្តរជាតិ (កំពង់ផែ)',
    shortLabel: 'កំពង់ផែ',
    icon: Anchor,
    accentColor: 'text-indigo-700',
    borderColor: 'border-indigo-300',
    bgGradient: 'from-indigo-50 via-purple-50/40 to-white',
    activeTabClass: 'bg-indigo-700 text-white border-indigo-700 shadow-md ring-2 ring-indigo-200',
    badgeClass: 'bg-indigo-100 text-indigo-900 border-indigo-200',
    durationMs: 10000, // 10 seconds for 6 seaport teams
  },
];

export const VisaTeamsCategorizedTable: React.FC<VisaTeamsCategorizedTableProps> = ({
  categories,
  officers = [],
  stockRecords = [],
  onNavigate,
}) => {
  // Current active step index: 0 = air, 1 = border, 2 = port
  const [currentCatIndex, setCurrentCatIndex] = useState<number>(0);
  const [isPlaying, setIsPlaying] = useState<boolean>(true);
  const [isHovered, setIsHovered] = useState<boolean>(false);
  const [progress, setProgress] = useState<number>(0); // 0 to 100%
  const [searchQuery, setSearchQuery] = useState('');
  const [isStaticAllMode, setIsStaticAllMode] = useState<boolean>(false);
  const [fadeKey, setFadeKey] = useState<number>(0); // Trigger smooth CSS fade on change

  const activeCategoryConfig = CATEGORY_SEQUENCE[currentCatIndex];

  // 1. Build the complete list of 29 teams
  const teamsList: VisaTeamItem[] = useMemo(() => {
    let rawPairs: { full: string; robok: string; id: string }[] = [];

    const vtrList = categories?.visaTeamsRobok || [];
    const vtList = categories?.visaTeams || [];
    const maxLen = Math.max(vtrList.length, vtList.length);

    if (maxLen > 0) {
      for (let i = 0; i < maxLen; i++) {
        const robok = (vtrList[i]?.name || '').trim();
        let full = (vtList[i]?.name || '').trim();
        const id = vtrList[i]?.id || vtList[i]?.id || `team-${i + 1}`;

        if (!full || /^ក្រុមទី\s*\d+$/i.test(full)) {
          if (robok && OFFICIAL_TEAM_FULL_NAMES[robok]) {
            full = OFFICIAL_TEAM_FULL_NAMES[robok];
          } else if (robok) {
            full = `ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារ${robok.replace(/^(ច្រកទ្វារ|ក្រុមផ្តល់ទិដ្ឋាការ|ប្រចាំ)+/, '').trim()}`;
          }
        }

        if (robok || full) {
          rawPairs.push({
            id,
            robok: robok || full,
            full: full || robok,
          });
        }
      }
    }

    // Baseline fallback if needed
    if (rawPairs.length === 0) {
      const defaultShort = [
        'អាកាស តេជោ',
        'អាកាស សៀមរាប',
        'អាកាស ព្រះសីហនុ',
        'ព្រំដែន ប៉ោយប៉ែត',
        'ព្រំដែន បាវិត',
        'ព្រំដែន ចាំយាម',
        'ព្រំដែន ដូង',
        'ព្រំដែន អូរស្មាច់',
        'ព្រំដែន ព្រំ',
        'ព្រំដែន បន្ទាយចក្រី',
        'ព្រំដែន ជាំ',
        'ព្រំដែន ត្រពាំងផ្លុង',
        'ព្រំដែន ត្រពាំងស្រែ',
        'ព្រំដែន ត្រពាំងរូង',
        'ព្រំដែន ភ្នំដិន',
        'ព្រំដែន កោះរកា',
        'ព្រំដែន ព្រែកចាក',
        'ព្រំដែន អូរយ៉ាដាវ',
        'ព្រំដែន ក្អមសំណ',
        'ព្រំដែន ភ្នំដី',
        'កំពង់ផែ ឧកញ៉ាម៉ុង',
        'កំពង់ផែ ស្ទឹងហាវ',
        'កំពង់ផែ ព្រះសីហនុ',
        'កំពង់ផែ ភ្នំពេញ',
        'ព្រំដែន ព្រែកបាក់',
        'ព្រំដែន ម៉ឺនជ័យ',
        'ព្រំដែន ស្ទឹងបត់',
        'កំពង់ផែ កោះកុង',
        'កំពង់ផែ កំពត',
      ];

      rawPairs = defaultShort.map((shortName, idx) => ({
        id: `def-${idx + 1}`,
        robok: shortName,
        full: OFFICIAL_TEAM_FULL_NAMES[shortName] || `ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារ ${shortName}`,
      }));
    }

    // Load custom team type options if stored
    let savedTypeOpts: Record<number, { sticker?: boolean; evisa?: boolean }> = {};
    try {
      const s = localStorage.getItem('visa_teams_type_options');
      if (s) savedTypeOpts = JSON.parse(s);
    } catch {}

    return rawPairs.map((p, idx) => {
      const combined = `${p.robok} ${p.full}`.toLowerCase();
      let category: PortCategoryType = 'border';
      let categoryLabel = 'ច្រកទ្វារព្រំដែនគោក';
      let categoryBadgeBg = 'bg-emerald-50 border-emerald-200 text-emerald-800';
      let categoryBadgeText = 'ព្រំដែន';

      if (/អាកាស|airport|air/i.test(combined)) {
        category = 'air';
        categoryLabel = 'ច្រកទ្វារអាកាសយានដ្ឋាន';
        categoryBadgeBg = 'bg-sky-50 border-sky-200 text-sky-800';
        categoryBadgeText = 'អាកាស';
      } else if (/កំពង់ផែ|ផែ|seaport|port/i.test(combined)) {
        category = 'port';
        categoryLabel = 'ច្រកទ្វារកំពង់ផែអន្តរជាតិ';
        categoryBadgeBg = 'bg-indigo-50 border-indigo-200 text-indigo-800';
        categoryBadgeText = 'កំពង់ផែ';
      }

      // Count officers in this team
      const normRobok = normalizeTeamName(p.robok);
      const count = (officers || []).filter((of) => {
        if (!of) return false;
        if (of.visaTeamId === p.id) return true;
        const ofWork = (of.officeWork || '').trim();
        const ofStation = (of.station || '').trim();
        const ofDept = (of.department || '').trim();

        if (ofWork && (ofWork === p.robok || ofWork === p.full || normalizeTeamName(ofWork) === normRobok)) return true;
        if (ofStation && (ofStation === p.robok || ofStation === p.full || normalizeTeamName(ofStation) === normRobok)) return true;
        if (ofDept && (ofDept === p.robok || ofDept === p.full || normalizeTeamName(ofDept) === normRobok)) return true;
        return false;
      }).length;

      const opt = savedTypeOpts[idx] || { sticker: true, evisa: true };

      return {
        id: p.id,
        robokName: p.robok,
        fullName: p.full,
        category,
        categoryLabel,
        categoryBadgeBg,
        categoryBadgeText,
        officerCount: count,
        useSticker: opt.sticker !== false,
        useEvisa: opt.evisa !== false,
      };
    });
  }, [categories?.visaTeamsRobok, categories?.visaTeams, officers]);

  // Group teams by category
  const categorizedTeams = useMemo(() => {
    return {
      air: teamsList.filter((t) => t.category === 'air'),
      border: teamsList.filter((t) => t.category === 'border'),
      port: teamsList.filter((t) => t.category === 'port'),
    };
  }, [teamsList]);

  // Current list to display
  const currentCategoryTeams = useMemo(() => {
    const list = categorizedTeams[activeCategoryConfig.key] || [];
    if (!searchQuery.trim()) return list;
    const q = searchQuery.toLowerCase().trim();
    return list.filter(
      (t) => t.robokName.toLowerCase().includes(q) || t.fullName.toLowerCase().includes(q)
    );
  }, [categorizedTeams, activeCategoryConfig.key, searchQuery]);

  // 2. Animated Rotation Engine (ដំបូង អាកាស -> ព្រំដែន -> កំពង់ផែ -> វិលមកអាកាសវិញ)
  useEffect(() => {
    if (isStaticAllMode || !isPlaying || isHovered) {
      return;
    }

    const duration = activeCategoryConfig.durationMs;
    const intervalMs = 100;
    const stepIncrement = (intervalMs / duration) * 100;

    const timer = setInterval(() => {
      setProgress((prev) => {
        const next = prev + stepIncrement;
        if (next >= 100) {
          // Advance to next category in cycle!
          setCurrentCatIndex((c) => (c + 1) % CATEGORY_SEQUENCE.length);
          setFadeKey((k) => k + 1);
          return 0;
        }
        return next;
      });
    }, intervalMs);

    return () => clearInterval(timer);
  }, [isStaticAllMode, isPlaying, isHovered, activeCategoryConfig.durationMs, currentCatIndex]);

  // Manual navigation handlers
  const handleSelectCategory = (idx: number) => {
    setCurrentCatIndex(idx);
    setProgress(0);
    setFadeKey((k) => k + 1);
    setIsStaticAllMode(false);
  };

  const handleNext = () => {
    setCurrentCatIndex((c) => (c + 1) % CATEGORY_SEQUENCE.length);
    setProgress(0);
    setFadeKey((k) => k + 1);
  };

  const handlePrev = () => {
    setCurrentCatIndex((c) => (c - 1 + CATEGORY_SEQUENCE.length) % CATEGORY_SEQUENCE.length);
    setProgress(0);
    setFadeKey((k) => k + 1);
  };

  return (
    <div
      className="bg-white rounded-md border border-gray-200 shadow-xs overflow-hidden mt-4 font-siemreap font-sans"
      onMouseEnter={() => setIsHovered(true)}
      onMouseLeave={() => setIsHovered(false)}
    >
      {/* 1. Header with Motion State & Controls */}
      <div className="px-4 py-3 border-b border-gray-200 bg-white flex flex-col lg:flex-row lg:items-center justify-between gap-3 font-siemreap">
        <div>
          <div className="flex items-center gap-2.5">
            <div className="p-1.5 bg-[#17a2b8]/15 text-[#17a2b8] rounded-md">
              <Sparkles className="w-5 h-5 text-teal-700" />
            </div>
            <div>
              <div className="flex items-center gap-2">
                <h3 className="text-base font-bold text-gray-900 leading-tight font-siemreap">
                  តារាងក្រុមផ្តល់ទិដ្ឋាការ (ចលនាវិលជុំតាមប្រភេទច្រកទ្វារ)
                </h3>
                {isPlaying && !isHovered && !isStaticAllMode && (
                  <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-teal-100 text-teal-800 text-[10.5px] font-semibold animate-pulse font-siemreap">
                    <span className="w-1.5 h-1.5 rounded-full bg-teal-600"></span>
                    <span>កំពុងវិលជុំស្វ័យប្រវត្តិ</span>
                  </span>
                )}
                {isHovered && !isStaticAllMode && isPlaying && (
                  <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-amber-100 text-amber-900 text-[10.5px] font-semibold font-siemreap">
                    <Pause className="w-2.5 h-2.5" />
                    <span>បានផ្អាកបណ្តោះអាសន្ន (Mouse លើតារាង)</span>
                  </span>
                )}
              </div>
              <p className="text-xs text-gray-500 mt-0.5 font-siemreap">
                បង្ហាញវិលជុំជាបន្តបន្ទាប់ ៖ <span className="font-semibold text-sky-700">១. អាកាស</span> ➔{' '}
                <span className="font-semibold text-emerald-700">២. ព្រំដែន</span> ➔{' '}
                <span className="font-semibold text-indigo-700">៣. កំពង់ផែ</span> ➔{' '}
                <span className="font-semibold text-gray-700">វិលមកអាកាសវិញ</span>
              </p>
            </div>
          </div>
        </div>

        {/* Right Action Controls */}
        <div className="flex flex-wrap items-center gap-2 self-start lg:self-auto font-siemreap">
          {/* Search Box */}
          <div className="relative">
            <input
              type="text"
              value={searchQuery}
              onChange={(e) => setSearchQuery(e.target.value)}
              placeholder="ស្វែងរកក្រុម..."
              className="w-40 sm:w-48 pl-7 pr-2.5 py-1.5 text-xs bg-white border border-gray-300 rounded-md focus:border-teal-600 focus:ring-1 focus:ring-teal-600 focus:outline-none transition shadow-2xs font-siemreap"
            />
            <Search className="w-3 h-3 text-gray-400 absolute left-2.5 top-1/2 -translate-y-1/2 pointer-events-none" />
            {searchQuery && (
              <button
                onClick={() => setSearchQuery('')}
                className="absolute right-2 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600 text-xs"
              >
                ×
              </button>
            )}
          </div>

          {/* Prev / Play / Next Controls */}
          <div className="inline-flex items-center rounded-md border border-gray-200 bg-white shadow-2xs p-0.5">
            <button
              onClick={handlePrev}
              title="ថយទៅប្រភេទមុន"
              className="p-1 text-gray-600 hover:text-gray-900 hover:bg-gray-100 rounded transition cursor-pointer"
            >
              <ChevronLeft className="w-4 h-4" />
            </button>
            <button
              onClick={() => setIsPlaying(!isPlaying)}
              title={isPlaying ? 'ផ្អាកចលនា' : 'បន្តដំណើរការចលនាស្វ័យប្រវត្តិ'}
              className="px-2 py-1 flex items-center gap-1 text-xs font-bold text-gray-700 hover:bg-gray-100 rounded transition cursor-pointer font-siemreap"
            >
              {isPlaying ? (
                <>
                  <Pause className="w-3 h-3 text-amber-600" />
                  <span className="hidden sm:inline">ផ្អាក</span>
                </>
              ) : (
                <>
                  <Play className="w-3 h-3 text-teal-600" />
                  <span className="hidden sm:inline">ចាក់</span>
                </>
              )}
            </button>
            <button
              onClick={handleNext}
              title="ទៅប្រភេទបន្ទាប់"
              className="p-1 text-gray-600 hover:text-gray-900 hover:bg-gray-100 rounded transition cursor-pointer"
            >
              <ChevronRight className="w-4 h-4" />
            </button>
          </div>

          {/* Toggle Static All View */}
          <button
            onClick={() => setIsStaticAllMode(!isStaticAllMode)}
            className={`px-2.5 py-1.5 rounded-md text-xs font-semibold border transition cursor-pointer flex items-center gap-1 font-siemreap ${
              isStaticAllMode
                ? 'bg-gray-900 text-white border-gray-900'
                : 'bg-white text-gray-700 border-gray-200 hover:bg-gray-50 shadow-2xs'
            }`}
            title="ប្តូររវាងចលនាវិលជុំ ឬបង្ហាញទាំងអស់ក្នុងពេលតែមួយ"
          >
            <RotateCcw className="w-3 h-3" />
            <span className="hidden sm:inline">
              {isStaticAllMode ? 'វិលជុំស្វ័យប្រវត្តិ' : 'មើលទាំងអស់ (Static)'}
            </span>
          </button>

          {/* Navigate to Category Manager */}
          {onNavigate && (
            <button
              onClick={() => onNavigate('categoryManager', 'visaTeamsData')}
              className="px-2.5 py-1.5 text-xs font-medium text-teal-800 bg-teal-50 hover:bg-teal-100 border border-teal-200 rounded-md transition flex items-center gap-1 cursor-pointer font-siemreap"
              title="គ្រប់គ្រងទិន្នន័យក្រុម"
            >
              <Building2 className="w-3.5 h-3.5 text-teal-700" />
            </button>
          )}
        </div>
      </div>

      {/* 2. Interactive Step Sequence Tabs: [១. អាកាស] ➔ [២. ព្រំដែន] ➔ [៣. កំពង់ផែ] */}
      {!isStaticAllMode && (
        <div className="px-4 py-2.5 bg-gray-50/80 border-b border-gray-200/80 flex flex-wrap items-center justify-between gap-3 font-siemreap">
          <div className="flex flex-wrap items-center gap-2">
            {CATEGORY_SEQUENCE.map((cat, idx) => {
              const isCurrent = currentCatIndex === idx;
              const IconComp = cat.icon;
              const count = categorizedTeams[cat.key].length;

              return (
                <React.Fragment key={cat.key}>
                  <button
                    onClick={() => handleSelectCategory(idx)}
                    className={`inline-flex items-center gap-2 px-3 py-1.5 rounded-md text-xs font-bold transition-all cursor-pointer border font-siemreap ${
                      isCurrent
                        ? cat.activeTabClass
                        : 'bg-white text-gray-700 border-gray-200 hover:bg-gray-100 hover:border-gray-300'
                    }`}
                  >
                    <IconComp className={`w-3.5 h-3.5 ${isCurrent ? 'text-white' : cat.accentColor}`} />
                    <span className="font-siemreap">
                      {cat.stepNum}. {cat.shortLabel}
                    </span>
                    <span
                      className={`text-[11px] px-1.5 py-0.2 rounded-full font-bold ${
                        isCurrent ? 'bg-white/25 text-white' : 'bg-gray-100 text-gray-800'
                      }`}
                    >
                      <span className="font-times">{count}</span> <span className="font-siemreap">ក្រុម</span>
                    </span>
                  </button>

                  {idx < CATEGORY_SEQUENCE.length - 1 && (
                    <span className="text-gray-300 font-bold text-xs hidden sm:inline">➔</span>
                  )}
                </React.Fragment>
              );
            })}

            <span className="text-gray-300 font-bold text-xs hidden sm:inline">➔</span>
            <span
              className="inline-flex items-center gap-1 text-[11px] font-semibold text-gray-500 bg-white px-2 py-1 rounded border border-gray-200 font-siemreap"
              title="វិលមកអាកាសវិញជាស្វ័យប្រវត្តិ"
            >
              <RotateCcw className="w-3 h-3 text-teal-600 animate-spin" style={{ animationDuration: '6s' }} />
              <span className="font-siemreap">វិលជុំវិញ</span>
            </span>
          </div>

          {/* Step Counter & Duration Indicator */}
          <div className="text-xs text-gray-500 flex items-center gap-2 font-siemreap">
            <span>
              ជំហាន ៖ <strong className="text-gray-900 font-times">{currentCatIndex + 1}/3</strong>
            </span>
            <span className="text-gray-300">|</span>
            <span>
              រយៈពេល ៖{' '}
              <strong className="text-gray-800 font-times">
                {activeCategoryConfig.durationMs / 1000}s
              </strong>
            </span>
          </div>
        </div>
      )}

      {/* 3. Smooth Progress Bar */}
      {!isStaticAllMode && isPlaying && (
        <div className="w-full bg-gray-100 h-1 overflow-hidden">
          <div
            className={`h-full transition-all duration-100 ease-linear ${
              activeCategoryConfig.key === 'air'
                ? 'bg-sky-500'
                : activeCategoryConfig.key === 'border'
                ? 'bg-emerald-600'
                : 'bg-indigo-600'
            }`}
            style={{ width: `${Math.min(100, Math.max(0, progress))}%` }}
          />
        </div>
      )}

      {/* 4. Active Category Presentation */}
      <div className="p-4 sm:p-5 font-siemreap">
        {isStaticAllMode ? (
          /* STATIC ALL MODE: 3 tables stacked */
          <div className="space-y-6">
            {CATEGORY_SEQUENCE.map((cat) => {
              const list = categorizedTeams[cat.key];
              const IconComp = cat.icon;
              return (
                <div key={cat.key} className={`bg-white rounded-md border ${cat.borderColor} shadow-2xs overflow-hidden font-siemreap`}>
                  <div className={`px-4 py-2.5 bg-gradient-to-r ${cat.bgGradient} border-b ${cat.borderColor} flex items-center justify-between`}>
                    <div className="flex items-center gap-2 font-bold text-xs text-gray-900 font-siemreap">
                      <IconComp className={`w-4 h-4 ${cat.accentColor}`} />
                      <span className="font-siemreap">
                        {cat.stepNum}. {cat.labelKh}
                      </span>
                    </div>
                    <span className={`px-2 py-0.5 rounded text-[11px] font-bold ${cat.badgeClass}`}>
                      <span className="font-times">{list.length}</span> <span className="font-siemreap">ក្រុម</span>
                    </span>
                  </div>
                  <div className="overflow-x-auto">
                    <table className="w-full text-xs border-collapse font-siemreap">
                      <thead>
                        <tr className="bg-gray-100/90 text-gray-700 font-bold border-b border-gray-200 text-[11.5px] font-siemreap">
                          <th className="py-2 px-3 text-center w-12 border-r border-gray-200 font-siemreap">ល.រ</th>
                          <th className="py-2 px-3.5 text-left w-52 border-r border-gray-200 font-siemreap">ឈ្មោះកាត់/របក</th>
                          <th className="py-2 px-3.5 text-left border-r border-gray-200 font-siemreap">ឈ្មោះផ្លូវការពេញលេញ</th>
                          <th className="py-2 px-3 text-center w-28 border-r border-gray-200 font-siemreap">ប្រភេទប្រើប្រាស់</th>
                          <th className="py-2 px-3 text-center w-24 border-r border-gray-200 font-siemreap">មន្ត្រីប្រចាំក្រុម</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-gray-100 font-siemreap">
                        {list.map((team, tIdx) => (
                          <tr key={team.id || `${team.robokName}-${tIdx}`} className="hover:bg-gray-50 transition-colors font-siemreap">
                            <td className="py-2 px-3 text-center text-gray-500 font-times font-semibold border-r border-gray-100">
                              {tIdx + 1}
                            </td>
                            <td className="py-2 px-3.5 font-bold text-gray-900 border-r border-gray-100 font-siemreap">
                              {team.robokName}
                            </td>
                            <td className="py-2 px-3.5 text-gray-700 border-r border-gray-100 font-siemreap">
                              {team.fullName}
                            </td>
                            <td className="py-2 px-3 text-center border-r border-gray-100 font-siemreap">
                              <span className="text-[10.5px] px-1.5 py-0.5 bg-blue-50 text-blue-700 border border-blue-200 rounded font-medium mr-1 font-siemreap">
                                សន្លឹក
                              </span>
                              <span className="text-[10.5px] px-1.5 py-0.5 bg-amber-50 text-amber-800 border border-amber-200 rounded font-medium font-siemreap">
                                ក្រដាស
                              </span>
                            </td>
                            <td className="py-2 px-3 text-center border-r border-gray-100 font-siemreap">
                              {team.officerCount > 0 ? (
                                <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-800 text-[11px] font-bold">
                                  <span className="font-times">{team.officerCount}</span>
                                  <span className="text-[10px] font-siemreap">នាក់</span>
                                </span>
                              ) : (
                                <span className="text-gray-400 font-times text-xs">-</span>
                              )}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              );
            })}
          </div>
        ) : (
          /* ANIMATED VIEW: Native CSS smooth transition */
          <div
            key={fadeKey}
            className={`bg-white rounded-md border ${activeCategoryConfig.borderColor} shadow-xs overflow-hidden transition-all duration-300 ease-out font-siemreap`}
          >
            {/* Active Category Title Banner */}
            <div
              className={`px-4 sm:px-5 py-3 bg-gradient-to-r ${activeCategoryConfig.bgGradient} border-b ${activeCategoryConfig.borderColor} flex flex-col sm:flex-row sm:items-center justify-between gap-2 font-siemreap`}
            >
              <div className="flex items-center gap-2.5">
                <div
                  className={`p-1.5 rounded-md bg-white border ${activeCategoryConfig.borderColor} shadow-2xs`}
                >
                  {activeCategoryConfig.key === 'air' && <Plane className="w-5 h-5 text-sky-600" />}
                  {activeCategoryConfig.key === 'border' && <MapPin className="w-5 h-5 text-emerald-700" />}
                  {activeCategoryConfig.key === 'port' && <Anchor className="w-5 h-5 text-indigo-700" />}
                </div>
                <div>
                  <div className="flex items-center gap-2">
                    <span className="text-xs font-bold text-gray-500 uppercase tracking-wide font-siemreap">
                      ជំហានទី {activeCategoryConfig.stepNum} នៃ {CATEGORY_SEQUENCE.length} ៖
                    </span>
                    <h4 className="text-sm sm:text-base font-extrabold text-gray-900 leading-tight font-siemreap">
                      {activeCategoryConfig.labelKh}
                    </h4>
                  </div>
                  <p className="text-xs text-gray-600 mt-0.5 font-siemreap">
                    {activeCategoryConfig.key === 'air'
                      ? 'រៀបរាប់ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអាកាសយានដ្ឋានអន្តរជាតិទាំងអស់'
                      : activeCategoryConfig.key === 'border'
                      ? 'រៀបរាប់ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិទាំងអស់'
                      : 'រៀបរាប់ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារកំពង់ផែអន្តរជាតិទាំងអស់'}
                  </p>
                </div>
              </div>

              {/* Team count tag */}
              <div className="flex items-center gap-2 self-start sm:self-auto font-siemreap">
                <span
                  className={`px-3 py-1 rounded-full text-xs font-bold border ${activeCategoryConfig.badgeClass} flex items-center gap-1.5 shadow-2xs font-siemreap`}
                >
                  <CheckCircle2 className="w-3.5 h-3.5" />
                  <span className="font-siemreap">សរុប ៖</span>
                  <strong className="font-times text-sm">
                    {currentCategoryTeams.length}
                  </strong>
                  <span className="font-siemreap">ក្រុម</span>
                </span>
              </div>
            </div>

            {/* Table of Teams for the active category */}
            <div className="overflow-x-auto font-siemreap">
              <table className="w-full text-xs border-collapse font-siemreap">
                <thead>
                  <tr className="bg-gray-100 text-gray-700 font-bold border-b border-gray-200 text-[11.5px] font-siemreap">
                    <th className="py-2.5 px-3 text-center w-12 border-r border-gray-200 font-siemreap">ល.រ</th>
                    <th className="py-2.5 px-3 text-left w-24 border-r border-gray-200 font-siemreap">ប្រភេទច្រក</th>
                    <th className="py-2.5 px-3.5 text-left w-52 border-r border-gray-200 font-siemreap">
                      ឈ្មោះសម្គាល់របាយការណ៍ (ឈ្មោះកាត់/របក)
                    </th>
                    <th className="py-2.5 px-3.5 text-left border-r border-gray-200 font-siemreap">
                      ឈ្មោះផ្លូវការពេញលេញ
                    </th>
                    <th className="py-2.5 px-3 text-center w-28 border-r border-gray-200 font-siemreap">
                      ប្រភេទប្រើប្រាស់
                    </th>
                    <th className="py-2.5 px-3 text-center w-24 border-r border-gray-200 font-siemreap">
                      មន្ត្រីប្រចាំក្រុម
                    </th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-100 font-siemreap">
                  {currentCategoryTeams.length === 0 ? (
                    <tr>
                      <td colSpan={6} className="py-8 text-center text-gray-400 text-xs italic font-siemreap">
                        ពុំមានទិន្នន័យក្រុមត្រូវតាមលក្ខខណ្ឌស្វែងរកនេះទេ
                      </td>
                    </tr>
                  ) : (
                    currentCategoryTeams.map((team, idx) => {
                      return (
                        <tr
                          key={team.id || `${team.robokName}-${idx}`}
                          className={`transition-colors hover:bg-amber-50/50 font-siemreap ${
                            idx % 2 === 0 ? 'bg-white' : 'bg-gray-50/40'
                          }`}
                        >
                          {/* 1. ល.រ */}
                          <td className="py-2.5 px-3 text-center text-gray-500 font-times text-xs font-semibold border-r border-gray-100">
                            {idx + 1}
                          </td>

                          {/* 2. ប្រភេទច្រកទ្វារ */}
                          <td className="py-2.5 px-3 border-r border-gray-100 whitespace-nowrap font-siemreap">
                            <span
                              className={`inline-flex items-center gap-1 px-2 py-0.5 rounded text-[11px] font-bold border font-siemreap ${team.categoryBadgeBg}`}
                            >
                              {team.category === 'air' && <Plane className="w-3 h-3 text-sky-600" />}
                              {team.category === 'border' && <MapPin className="w-3 h-3 text-emerald-600" />}
                              {team.category === 'port' && <Anchor className="w-3 h-3 text-indigo-600" />}
                              <span className="font-siemreap">{team.categoryBadgeText}</span>
                            </span>
                          </td>

                          {/* 3. ឈ្មោះសម្គាល់របាយការណ៍ (ឈ្មោះកាត់/របក) */}
                          <td className="py-2.5 px-3.5 border-r border-gray-100 font-siemreap">
                            <span className="font-bold text-gray-900 text-xs tracking-tight font-siemreap">
                              {team.robokName}
                            </span>
                          </td>

                          {/* 4. ឈ្មោះផ្លូវការពេញលេញ (Khmer OS Siemreap identical to Image 2) */}
                          <td className="py-2.5 px-3.5 border-r border-gray-100 text-xs text-gray-700 leading-snug font-siemreap">
                            {team.fullName}
                          </td>

                          {/* 5. ប្រភេទប្រើប្រាស់ */}
                          <td className="py-2.5 px-3 border-r border-gray-100 text-center whitespace-nowrap font-siemreap">
                            <div className="inline-flex items-center gap-1">
                              <span
                                className={`text-[10.5px] px-1.5 py-0.5 rounded font-medium font-siemreap ${
                                  team.useSticker
                                    ? 'bg-blue-50 text-blue-700 border border-blue-200'
                                    : 'bg-gray-100 text-gray-400'
                                }`}
                                title="សន្លឹកទិដ្ឋាការ (Sticker)"
                              >
                                សន្លឹក
                              </span>
                              <span
                                className={`text-[10.5px] px-1.5 py-0.5 rounded font-medium font-siemreap ${
                                  team.useEvisa
                                    ? 'bg-amber-50 text-amber-800 border border-amber-200'
                                    : 'bg-gray-100 text-gray-400'
                                }`}
                                title="ក្រដាសអនុម័ត (cEA / Approval Paper)"
                              >
                                ក្រដាស
                              </span>
                            </div>
                          </td>

                          {/* 6. ចំនួនមន្ត្រីប្រចាំក្រុម */}
                          <td className="py-2.5 px-3 text-center border-r border-gray-100 font-siemreap">
                            {team.officerCount > 0 ? (
                              <span className="inline-flex items-center gap-1 px-2 py-0.5 rounded-full bg-emerald-100 text-emerald-800 text-[11px] font-bold">
                                <span className="font-times">
                                  {team.officerCount}
                                </span>
                                <span className="text-[10px] font-siemreap">នាក់</span>
                              </span>
                            ) : (
                              <span className="text-gray-400 font-times text-xs">
                                -
                              </span>
                            )}
                          </td>
                        </tr>
                      );
                    })
                  )}
                </tbody>
                {/* Clean, perfectly-formatted Footer matching Image 2 Khmer font */}
                <tfoot className="font-siemreap">
                  <tr className="bg-gray-50/95 font-bold border-t-2 border-gray-200 text-gray-800 text-xs font-siemreap">
                    <td colSpan={2} className="py-2.5 px-3 text-left border-r border-gray-200 font-siemreap font-bold">
                      សរុបក្រុម {activeCategoryConfig.shortLabel} ៖
                    </td>
                    <td
                      colSpan={3}
                      className="py-2.5 px-3.5 text-left border-r border-gray-200 text-teal-800 font-siemreap font-bold"
                    >
                      <span className="font-times font-bold text-[13px]">{currentCategoryTeams.length}</span>{' '}
                      <span className="font-siemreap">ក្រុម (ក្នុងចំណោមច្រកទ្វារអន្តរជាតិសរុប</span>{' '}
                      <span className="font-times font-bold text-[13px]">{teamsList.length}</span>{' '}
                      <span className="font-siemreap">ក្រុម)</span>
                    </td>
                    <td className="py-2.5 px-3 text-center text-gray-900 font-siemreap font-bold">
                      <span className="font-times font-bold text-[13px]">
                        {currentCategoryTeams.reduce((acc, curr) => acc + curr.officerCount, 0)}
                      </span>{' '}
                      <span className="font-siemreap text-xs font-normal">នាក់</span>
                    </td>
                  </tr>
                </tfoot>
              </table>
            </div>
          </div>
        )}

        {/* Informative Footer & Legend */}
        <div className="mt-3.5 pt-2.5 border-t border-gray-100 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-2 text-[11px] text-gray-400 font-siemreap">
          <div className="flex items-center gap-2 font-siemreap">
            <span className="font-semibold text-gray-600 font-siemreap">លំដាប់ចលនា ៖</span>
            <span className="font-siemreap">
              ដំបូងរៀបរាប់ <strong>អាកាស</strong> (<span className="font-times">{categorizedTeams.air.length}</span> ក្រុម) ➔ ពេលបង្ហាញអស់ ដល់{' '}
              <strong>ព្រំដែន</strong> (<span className="font-times">{categorizedTeams.border.length}</span> ក្រុម) ➔ ដល់{' '}
              <strong>កំពង់ផែ</strong> (<span className="font-times">{categorizedTeams.port.length}</span> ក្រុម) ➔ វិលមកអាកាសវិញ។
            </span>
          </div>
          <div className="text-[10.5px] text-gray-400 font-times">
            Total: {teamsList.length} Teams
          </div>
        </div>
      </div>
    </div>
  );
};
