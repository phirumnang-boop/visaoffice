import React, { useState, useMemo, useEffect } from 'react';
import { StockRecord } from '../types';
import { isCeaRecord, calculateAllTeamsStickerStockAtDate, isRecordForTeam } from '../utils/teamStockCalculation';
import {
  normalizeDateToISO,
  normalizeVisaType,
  VISA_TYPES,
  normalizeTeamName,
  matchTeamInList,
  OFFICIAL_29_TEAMS,
} from '../utils/teamNormalization';
import {
  BarChart3,
  Calendar,
  Filter,
  TrendingUp,
  TrendingDown,
  Table as TableIcon,
  ChevronDown,
  RotateCw,
  Archive,
  Users,
} from 'lucide-react';

interface MonthlyVisaUsageComparisonChartProps {
  stockRecords: StockRecord[];
  assignedTeam?: string;
}

// 13 Official Visa Types in exact government sequence
const ALL_13_VISA_TYPES = ['T', 'T1', 'T2', 'T3', 'E', 'E1', 'E2', 'E3', 'D', 'K', 'A', 'B', 'C'] as const;

// Official Baseline for "សន្និធិ ចុងគ្រា ៣០-វិច្ឆិកា-២០១៨" (30-Nov-2018 benchmark figures)
const OFFICIAL_K2_DEC_2018_BASELINE: Record<string, number> = {
  T: 304000,
  T1: 3200,
  T2: 6200,
  T3: 5200,
  E: 330750,
  E1: 2650,
  E2: 4700,
  E3: 4950,
  D: 3600,
  K: 16750,
  A: 6100,
  B: 7250,
  C: 4850,
};

const OFFICIAL_TEAMS_DEC_2018_BASELINE: Record<string, number> = {
  T: 93675,
  T1: 2223,
  T2: 1927,
  T3: 1682,
  E: 37448,
  E1: 2144,
  E2: 2057,
  E3: 1881,
  D: 1349,
  K: 3648,
  A: 2060,
  B: 2512,
  C: 3500,
};

const KHMER_MONTHS = [
  { num: 1, khmer: 'មករា', shortKh: 'មករា', en: 'Jan' },
  { num: 2, khmer: 'កុម្ភៈ', shortKh: 'កុម្ភៈ', en: 'Feb' },
  { num: 3, khmer: 'មីនា', shortKh: 'មីនា', en: 'Mar' },
  { num: 4, khmer: 'មេសា', shortKh: 'មេសា', en: 'Apr' },
  { num: 5, khmer: 'ឧសភា', shortKh: 'ឧសភា', en: 'May' },
  { num: 6, khmer: 'មិថុនា', shortKh: 'មិថុនា', en: 'Jun' },
  { num: 7, khmer: 'កក្កដា', shortKh: 'កក្កដា', en: 'Jul' },
  { num: 8, khmer: 'សីហា', shortKh: 'សីហា', en: 'Aug' },
  { num: 9, khmer: 'កញ្ញា', shortKh: 'កញ្ញា', en: 'Sep' },
  { num: 10, khmer: 'តុលា', shortKh: 'តុលា', en: 'Oct' },
  { num: 11, khmer: 'វិច្ឆិកា', shortKh: 'វិច្ឆិកា', en: 'Nov' },
  { num: 12, khmer: 'ធ្នូ', shortKh: 'ធ្នូ', en: 'Dec' },
];

export const MonthlyVisaUsageComparisonChart: React.FC<MonthlyVisaUsageComparisonChartProps> = ({
  stockRecords = [],
  assignedTeam,
}) => {
  const currentYear = new Date().getFullYear();

  // Selected Type: 'all' (Sticker + cEA), 'sticker_all', 'cea_all', or specific sticker type (T, E, etc.)
  const [selectedType, setSelectedType] = useState<string>('all');

  // Selected Years for comparison
  const [primaryYear, setPrimaryYear] = useState<number>(currentYear);
  const [compareYear, setCompareYear] = useState<number>(currentYear - 1);

  // View mode: 'chart' or 'table'
  const [viewMode, setViewMode] = useState<'chart' | 'table'>('chart');

  // Show numbers / data labels directly on top of bars
  const [showBarValues] = useState<boolean>(true);
  const [numberFormat] = useState<'compact' | 'full'>('full');
  const [textRotation, setTextRotation] = useState<'up' | 'angle' | 'horizontal'>('horizontal');

  // Hovered Month for Tooltip
  const [hoveredMonthIndex, setHoveredMonthIndex] = useState<number | null>(null);

  const formatBarValue = (val: number): string => {
    if (!val || val <= 0) return '';
    if (numberFormat === 'full') {
      return Number(val).toLocaleString('en-US');
    }
    if (val >= 1_000_000) {
      const formatted = (val / 1_000_000).toFixed(1);
      return `${formatted.endsWith('.0') ? formatted.slice(0, -2) : formatted}M`;
    }
    if (val >= 10_000) {
      const formatted = (val / 1_000).toFixed(val >= 100_000 ? 0 : 1);
      return `${formatted.endsWith('.0') ? formatted.slice(0, -2) : formatted}k`;
    }
    if (val >= 1_000) {
      const formatted = (val / 1_000).toFixed(1);
      return `${formatted.endsWith('.0') ? formatted.slice(0, -2) : formatted}k`;
    }
    return Number(val).toLocaleString('en-US');
  };

  // Extract all available years from stockRecords and daily operations
  const availableYears = useMemo(() => {
    const yearSet = new Set<number>([currentYear, currentYear - 1, 2024, 2023, 2022, 2021, 2020, 2019, 2018]);
    stockRecords.forEach((r) => {
      const d = normalizeDateToISO(r.date || (r as any).createdAt || '');
      if (d && d.length >= 4) {
        const y = parseInt(d.slice(0, 4), 10);
        if (y >= 2000 && y <= 2100) yearSet.add(y);
      }
    });

    try {
      const savedDaily = localStorage.getItem('app_daily_team_operations_v5');
      if (savedDaily) {
        const parsed = JSON.parse(savedDaily);
        if (Array.isArray(parsed)) {
          parsed.forEach((d: any) => {
            const dateStr = normalizeDateToISO(d.date || '');
            if (dateStr && dateStr.length >= 4) {
              const y = parseInt(dateStr.slice(0, 4), 10);
              if (y >= 2000 && y <= 2100) yearSet.add(y);
            }
          });
        }
      }
    } catch {}

    return Array.from(yearSet).sort((a, b) => b - a);
  }, [stockRecords, currentYear]);

  // Compute monthly usage map: year -> month (1-12) -> quantity
  const monthlyUsageData = useMemo(() => {
    const yearMonthData: Record<number, number[]> = {};
    const initYear = (y: number) => {
      if (!yearMonthData[y]) {
        yearMonthData[y] = Array(12).fill(0);
      }
    };

    initYear(primaryYear);
    initYear(compareYear);

    // Filter stockRecords
    (stockRecords || []).forEach((r) => {
      if (assignedTeam && !isRecordForTeam(r, assignedTeam)) return;

      const dIso = normalizeDateToISO(r.date || (r as any).createdAt || '');
      if (!dIso || dIso.length < 7) return;

      const [yStr, mStr] = dIso.split('-');
      const y = parseInt(yStr, 10);
      const m = parseInt(mStr, 10);
      if (isNaN(y) || isNaN(m) || m < 1 || m > 12) return;
      if (y !== primaryYear && y !== compareYear) return;

      initYear(y);

      const op = (r.operationType || '').trim().toLowerCase();
      const isUse =
        op === 'useteam' ||
        op === 'use_team' ||
        op === 'used' ||
        op.includes('useteam') ||
        op.includes('ប្រើប្រាស់');

      if (!isUse) return;

      const isCea = isCeaRecord(r) || r.stockType === 'evisa';
      const qty =
        Number(
          (r as any).quantity ||
            (r as any).quantityBundles ||
            r.totalSheets ||
            ((r as any).count ? (r as any).count * 50 : 0)
        ) || 0;

      if (qty <= 0) return;

      const vtNorm = normalizeVisaType(r.visaType);

      if (selectedType === 'all') {
        // Combined Sticker + cEA
        yearMonthData[y][m - 1] += qty;
      } else if (selectedType === 'sticker_all') {
        // Only Sticker (all types)
        if (!isCea) {
          yearMonthData[y][m - 1] += qty;
        }
      } else if (selectedType === 'cea_all') {
        // Only cEA / Approval Paper
        if (isCea) {
          yearMonthData[y][m - 1] += qty;
        }
      } else {
        // Specific sticker type (e.g. 'T', 'E', etc.)
        if (!isCea && vtNorm === selectedType) {
          yearMonthData[y][m - 1] += qty;
        }
      }
    });

    // Merge daily operations from localStorage to include uncommitted daily logs
    try {
      const savedDaily = localStorage.getItem('app_daily_team_operations_v5');
      if (savedDaily) {
        const parsedDaily = JSON.parse(savedDaily);
        if (Array.isArray(parsedDaily)) {
          const existingUseKeySet = new Set<string>();
          (stockRecords || []).forEach((sr) => {
            const srOp = (sr.operationType || '').toLowerCase();
            if (srOp === 'useteam' || srOp.includes('useteam') || srOp.includes('ប្រើ')) {
              const dIso = normalizeDateToISO(sr.date || '');
              const tNorm = normalizeTeamName(sr.visaTeamRobokName || sr.sourceFrom || '');
              const vNorm = normalizeVisaType(sr.visaType);
              if (dIso && tNorm && vNorm) {
                existingUseKeySet.add(`${dIso}_${tNorm}_${vNorm}`);
              }
            }
          });

          parsedDaily.forEach((dRec: any) => {
            const rawTeamName = dRec.teamName || '';
            const normTeam = normalizeTeamName(rawTeamName);
            if (assignedTeam && normTeam !== normalizeTeamName(assignedTeam) && rawTeamName !== assignedTeam) return;

            const recDate = normalizeDateToISO(dRec.date || '');
            if (!dRec.values || !recDate || recDate.length < 7) return;

            const [yStr, mStr] = recDate.split('-');
            const y = parseInt(yStr, 10);
            const m = parseInt(mStr, 10);
            if (isNaN(y) || isNaN(m) || m < 1 || m > 12) return;
            if (y !== primaryYear && y !== compareYear) return;

            initYear(y);

            const isCeaCategory = dRec.categoryType === 'cEA';

            if (isCeaCategory) {
              // cEA record
              const item = dRec.values['cEA'] || dRec.values['CEA'] || dRec.values['ApprovalPaper'];
              const qty = parseInt(item?.quantity || '', 10) || 0;
              if (qty > 0) {
                const lookupKey = `${recDate}_${normTeam}_CEA`;
                if (!existingUseKeySet.has(lookupKey)) {
                  if (selectedType === 'all' || selectedType === 'cea_all') {
                    yearMonthData[y][m - 1] += qty;
                  }
                }
              }
            } else {
              // Sticker daily records
              VISA_TYPES.forEach((vt) => {
                const item = dRec.values[vt];
                const qty = parseInt(item?.quantity || '', 10) || 0;
                if (qty <= 0) return;

                const lookupKey = `${recDate}_${normTeam}_${vt}`;
                if (!existingUseKeySet.has(lookupKey)) {
                  if (selectedType === 'all' || selectedType === 'sticker_all') {
                    yearMonthData[y][m - 1] += qty;
                  } else if (selectedType === vt) {
                    yearMonthData[y][m - 1] += qty;
                  }
                }
              });
            }
          });
        }
      }
    } catch {}

    return yearMonthData;
  }, [stockRecords, primaryYear, compareYear, selectedType, assignedTeam]);

  // Compute Remaining Stock for the Office ONLY (សន្និធិសន្លឹកទិដ្ឋាការនៅសល់ ការិយាល័យ ក២)
  const officeRemainingStockByType = useMemo(() => {
    const hasData = stockRecords && stockRecords.length > 0;
    // 1. Initialize with baseline for Office (K2)
    let baseMap: Record<string, number> = hasData ? { ...OFFICIAL_K2_DEC_2018_BASELINE } : {};
    try {
      const saved = localStorage.getItem('sticker_office_report_baselines');
      if (saved) {
        const parsed = JSON.parse(saved);
        if (parsed && typeof parsed === 'object') {
          const sum = (Object.values(parsed) as any[]).reduce((a: number, b: any) => a + (Number(b) || 0), 0);
          if (sum > 0) {
            baseMap = parsed as Record<string, number>;
          }
        }
      }
    } catch {}

    const k2Stock: Record<string, number> = {};
    ALL_13_VISA_TYPES.forEach((vt) => {
      k2Stock[vt] = baseMap[vt] || 0;
    });

    // 2. Process stock records affecting the Office (ក២)
    (stockRecords || []).forEach((rec) => {
      if (rec.stockType && rec.stockType !== 'sticker') return;
      if (isCeaRecord(rec)) return;

      const vt = normalizeVisaType(rec.visaType);
      if (!ALL_13_VISA_TYPES.includes(vt as any)) return;

      const qty = Number(rec.quantityBundles || rec.totalSheets || (rec as any).quantity || 0);
      const recDate = normalizeDateToISO(rec.date || '');
      const op = (rec.operationType || '').toLowerCase();

      if (recDate && recDate <= '2018-11-30') return;

      const isDamagedTeam =
        op === 'damagedteam' ||
        op === 'voidteam' ||
        op === 'missingteam' ||
        op === 'teamdamaged' ||
        op.includes('ខូចក្រុម') ||
        op.includes('ខ្វះក្រុម') ||
        op.includes('ខូចតាមក្រុម') ||
        op.includes('ខ្វះតាមក្រុម') ||
        ((op.includes('ខូច') || op.includes('មិនបានការ')) &&
          !op.includes('ក២') &&
          !op.includes('ក១') &&
          !op.includes('ការិយាល័យ') &&
          (rec.sourceFrom?.includes('ក្រុម') || (rec as any).visaTeamRobokName));

      const isDamagedK2 =
        !isDamagedTeam &&
        (op === 'damaged' ||
          op === 'damagedk2' ||
          op === 'damagedoffice' ||
          op === 'void' ||
          op === 'voidk2' ||
          op === 'voidoffice' ||
          op === 'missingoffice' ||
          op.includes('ខូចក') ||
          op.includes('មិនបានការក') ||
          op.includes('ខូច') ||
          op.includes('មិនបានការ'));

      const isTestSampleK2 =
        op === 'testprintk2' ||
        op === 'testk2' ||
        op === 'samplek2' ||
        op === 'testoffice' ||
        op === 'sampleoffice' ||
        op === 'test' ||
        op === 'sample' ||
        op === 'testprint' ||
        op.includes('សាកក') ||
        op.includes('សាកល្បង') ||
        op.includes('សាក');

      const isTransferTeam =
        op === 'transferteam' ||
        op === 'transferuseteam' ||
        op === 'transfer' ||
        op.includes('transfer') ||
        op.includes('ផ្ទេរ') ||
        (rec.sourceFrom && rec.sourceFrom.includes('ផ្ទេរ')) ||
        (rec.notes && rec.notes.includes('ផ្ទេរ'));

      // Office (K2) movements ONLY
      if (isTransferTeam) {
        // Team-to-team transfers do not affect Office (K2) stock balance
      } else if (op === 'openk1' || op === 'receivek1' || op === 'k1' || op.includes('បញ្ចូលស្តុក') || op.includes('ក១')) {
        k2Stock[vt] += qty;
      } else if (op === 'returnteam' || op === 'returnoffice' || op.includes('បង្វិល')) {
        k2Stock[vt] += qty;
      } else if (op === 'issueteam' || op.includes('បើកផ្តល់')) {
        k2Stock[vt] -= qty;
      } else if (isDamagedK2) {
        k2Stock[vt] -= qty;
      } else if (isTestSampleK2) {
        k2Stock[vt] -= qty;
      }
      // Note: Team usage (useTeam) and Team damaged (damagedTeam) do NOT reduce Office stock
    });

    const items = ALL_13_VISA_TYPES.map((vt) => {
      const remaining = k2Stock[vt] || 0;
      return {
        type: vt,
        remaining,
      };
    });

    const totalRemaining = items.reduce((acc, item) => acc + item.remaining, 0);

    return {
      items,
      totalRemaining,
    };
  }, [stockRecords]);

  // Selected team filter for Teams Remaining Stock Card: 'all' (All 29 teams) or specific team name
  const [selectedTeamFilter, setSelectedTeamFilter] = useState<string>(assignedTeam || 'all');

  useEffect(() => {
    if (assignedTeam) {
      setSelectedTeamFilter(assignedTeam);
    }
  }, [assignedTeam]);

  // Compute Remaining Stock for Teams (សន្និធិសន្លឹកទិដ្ឋាការនៅសល់ បណ្តាក្រុម)
  const teamsRemainingStockByType = useMemo(() => {
    const hasData = stockRecords && stockRecords.length > 0;
    if (!hasData) {
      return {
        items: ALL_13_VISA_TYPES.map((vt) => ({ type: vt, remaining: 0 })),
        totalRemaining: 0,
      };
    }
    if (selectedTeamFilter !== 'all') {
      const calculated = calculateAllTeamsStickerStockAtDate(
        stockRecords || [],
        '',
        [selectedTeamFilter] // ONLY calculate for the filtered team for extreme 29x speedup!
      );
      const normAssigned = normalizeTeamName(selectedTeamFilter);
      const targetOfficial = matchTeamInList(selectedTeamFilter, Array.from(OFFICIAL_29_TEAMS));
      const matchedKey =
        Object.keys(calculated.remainingAfterMatrix).find((k) => {
          if (k === selectedTeamFilter) return true;
          if (normalizeTeamName(k) === normAssigned) return true;
          const kOfficial = matchTeamInList(k, Array.from(OFFICIAL_29_TEAMS));
          if (kOfficial && targetOfficial && kOfficial === targetOfficial) return true;
          return false;
        }) || selectedTeamFilter;

      const teamStockObj = calculated.remainingAfterMatrix[matchedKey] || {};
      const items = ALL_13_VISA_TYPES.map((vt) => ({
        type: vt,
        remaining: Math.max(0, Number(teamStockObj[vt]) || 0),
      }));
      const totalRemaining = items.reduce((acc, item) => acc + item.remaining, 0);
      return { items, totalRemaining };
    }

    // Default: Total Remaining Stock across all 29 Teams (matching RobokTotalStockWorkReport & Dashboard)
    const openTeams: Record<string, number> = hasData ? { ...OFFICIAL_TEAMS_DEC_2018_BASELINE } : {};
    const issuedK2ToTeams: Record<string, number> = {};
    const teamReturnedToK2: Record<string, number> = {};
    const teamsUsed: Record<string, number> = {};
    const damagedTeam: Record<string, number> = {};

    ALL_13_VISA_TYPES.forEach((vt) => {
      if (!openTeams[vt]) openTeams[vt] = 0;
      issuedK2ToTeams[vt] = 0;
      teamReturnedToK2[vt] = 0;
      teamsUsed[vt] = 0;
      damagedTeam[vt] = 0;
    });

    (stockRecords || []).forEach((rec) => {
      if (rec.stockType && rec.stockType !== 'sticker') return;
      if (isCeaRecord(rec)) return;

      const vt = normalizeVisaType(rec.visaType);
      if (!ALL_13_VISA_TYPES.includes(vt as any)) return;

      const qty = Number(rec.quantityBundles || rec.totalSheets || (rec as any).quantity || 0);
      const recDate = normalizeDateToISO(rec.date || '');
      const op = (rec.operationType || '').toLowerCase();

      if (recDate && recDate <= '2018-11-30') return;

      const isDamagedTeam =
        op === 'damagedteam' ||
        op === 'voidteam' ||
        op === 'missingteam' ||
        op === 'teamdamaged' ||
        op.includes('ខូចក្រុម') ||
        op.includes('ខ្វះក្រុម') ||
        op.includes('ខូចតាមក្រុម') ||
        op.includes('ខ្វះតាមក្រុម') ||
        ((op.includes('ខូច') || op.includes('មិនបានការ')) &&
          !op.includes('ក២') &&
          !op.includes('ក១') &&
          !op.includes('ការិយាល័យ') &&
          (rec.sourceFrom?.includes('ក្រុម') || (rec as any).visaTeamRobokName));

      const isTransferTeam =
        op === 'transferteam' ||
        op === 'transferuseteam' ||
        op === 'transfer' ||
        op.includes('transfer') ||
        op.includes('ផ្ទេរ') ||
        (rec.sourceFrom && rec.sourceFrom.includes('ផ្ទេរ')) ||
        (rec.notes && rec.notes.includes('ផ្ទេរ'));

      if (isTransferTeam) {
        // Ignore team-to-team transfers for Office (K2) stock calculations
      } else if (op === 'returnteam' || op === 'returnoffice' || op.includes('បង្វិល')) {
        teamReturnedToK2[vt] += qty;
      } else if (op === 'issueteam' || op.includes('បើកផ្តល់')) {
        issuedK2ToTeams[vt] += qty;
      } else if (op === 'useteam' || op.includes('ប្រើប្រាស់')) {
        teamsUsed[vt] += qty;
      } else if (isDamagedTeam) {
        damagedTeam[vt] += qty;
      }
    });

    // Process local daily operations
    try {
      const savedDaily = localStorage.getItem('app_daily_team_operations_v5');
      if (savedDaily) {
        const parsedDaily = JSON.parse(savedDaily);
        if (Array.isArray(parsedDaily)) {
          const existingUseKeySet = new Set<string>();
          (stockRecords || []).forEach((sr) => {
            if (sr.operationType === 'useTeam' && !isCeaRecord(sr)) {
              const dIso = normalizeDateToISO(sr.date || '');
              const tNorm = (sr.visaTeamRobokName || sr.sourceFrom || '').trim();
              const vNorm = normalizeVisaType(sr.visaType);
              if (dIso && tNorm && vNorm) {
                existingUseKeySet.add(`${dIso}_${tNorm}_${vNorm}`);
              }
            }
          });

          parsedDaily.forEach((dRec: any) => {
            if (dRec.categoryType === 'cEA' || (dRec.categoryType && dRec.categoryType !== 'Sticker')) return;
            const recDate = normalizeDateToISO(dRec.date || '');
            if (!dRec.values || (recDate && recDate <= '2018-11-30')) return;

            const rawTeamName = dRec.teamName || '';

            ALL_13_VISA_TYPES.forEach((vt) => {
              const item = dRec.values[vt];
              const qty = parseInt(item?.quantity || '', 10) || 0;
              if (qty <= 0) return;

              const lookupKey = `${recDate}_${rawTeamName}_${vt}`;
              if (!existingUseKeySet.has(lookupKey)) {
                teamsUsed[vt] += qty;
              }
            });
          });
        }
      }
    } catch {}

    const items = ALL_13_VISA_TYPES.map((vt) => {
      const remaining = Math.max(
        0,
        (openTeams[vt] || 0) +
          (issuedK2ToTeams[vt] || 0) -
          (teamReturnedToK2[vt] || 0) -
          (teamsUsed[vt] || 0) -
          (damagedTeam[vt] || 0)
      );
      return {
        type: vt,
        remaining,
      };
    });

    const totalRemaining = items.reduce((acc, item) => acc + item.remaining, 0);

    return {
      items,
      totalRemaining,
    };
  }, [stockRecords, selectedTeamFilter]);

  const primaryMonthlyValues = monthlyUsageData[primaryYear] || Array(12).fill(0);
  const compareMonthlyValues = monthlyUsageData[compareYear] || Array(12).fill(0);

  const primaryTotal = primaryMonthlyValues.reduce((a, b) => a + b, 0);
  const compareTotal = compareMonthlyValues.reduce((a, b) => a + b, 0);

  // Growth / difference calculations
  const totalDiff = primaryTotal - compareTotal;
  const growthPercent =
    compareTotal > 0 ? ((totalDiff / compareTotal) * 100).toFixed(1) : primaryTotal > 0 ? '+100' : '0';

  // Peak month for primary year
  const peakMonthIdx = primaryMonthlyValues.reduce(
    (maxI, val, i, arr) => (val > arr[maxI] ? i : maxI),
    0
  );
  const peakMonthName = KHMER_MONTHS[peakMonthIdx]?.khmer || '';
  const peakMonthVal = primaryMonthlyValues[peakMonthIdx] || 0;

  // Max value for chart Y-axis scale
  const maxMonthlyVal = Math.max(
    ...primaryMonthlyValues,
    ...compareMonthlyValues,
    100 // minimum baseline
  );

  // Generate nice step markers for Y-axis (4 or 5 intervals) with headroom for bar labels
  const yAxisTicks = useMemo(() => {
    const bufferedMax = maxMonthlyVal * (textRotation === 'up' ? 1.36 : textRotation === 'angle' ? 1.28 : 1.20);
    const rawStep = bufferedMax / 4;
    const magnitude = Math.pow(10, Math.floor(Math.log10(rawStep || 1)));
    const cleanStep = Math.ceil(rawStep / magnitude) * magnitude || 100;
    const ticks = [];
    for (let i = 0; i <= 4; i++) {
      ticks.push(cleanStep * i);
    }
    if (ticks[ticks.length - 1] < bufferedMax) {
      ticks.push(cleanStep * 5);
    }
    return ticks;
  }, [maxMonthlyVal, textRotation]);

  const yMax = yAxisTicks[yAxisTicks.length - 1] || 100;

  const formatUnit =
    selectedType === 'cea_all'
      ? 'ច្បាប់ / ដុំ'
      : selectedType === 'all'
      ? 'សន្លឹក / ច្បាប់'
      : 'សន្លឹក';

  const typeLabel = useMemo(() => {
    if (selectedType === 'all') return 'Sticker + cEA (បញ្ចូលគ្នាទាំងអស់)';
    if (selectedType === 'sticker_all') return 'សន្លឹកទិដ្ឋាការសរុប (Sticker All)';
    if (selectedType === 'cea_all') return 'ក្រដាសអនុម័ត (cEA / Approval Paper)';
    return `ទិដ្ឋាការប្រភេទ «${selectedType}»`;
  }, [selectedType]);

  return (
    <div className="grid grid-cols-1 md:grid-cols-2 xl:grid-cols-4 gap-4 items-stretch mt-4 font-siemreap">
      {/* 1. Left Card: Office Visa Sticker Remaining Stock Card (កាតសន្និធិសន្លឹកទិដ្ឋាការនៅសល់ ការិយាល័យ) */}
      <div className="col-span-1 bg-white rounded-md border border-gray-200 shadow-xs overflow-hidden flex flex-col justify-between">
        {/* Header */}
        <div className="px-3.5 py-2.5 border-b border-gray-200 bg-white flex items-center justify-between shrink-0">
          <div className="flex items-center gap-2">
            <div className="p-1.5 bg-amber-100 text-amber-800 rounded-md">
              <Archive className="w-4 h-4" />
            </div>
            <div>
              <h4 className="text-sm font-bold text-gray-900 leading-tight whitespace-nowrap">
                សន្និធិ
              </h4>
              <div className="flex items-center gap-1.5 mt-0.5">
                <span className="inline-block px-2 py-0.5 bg-amber-100/80 border border-amber-300 text-amber-900 text-[11px] font-bold rounded whitespace-nowrap">
                  ការិយាល័យ (ក២)
                </span>
                <span className="text-[10.5px] text-gray-400 whitespace-nowrap">១៣ ប្រភេទ</span>
              </div>
            </div>
          </div>
          {selectedType !== 'all' && (
            <button
              onClick={() => setSelectedType('all')}
              className="text-[11px] px-2 py-0.5 bg-gray-100 hover:bg-gray-200 text-gray-700 rounded transition font-medium cursor-pointer"
              title="បង្ហាញគ្រប់ប្រភេទវិញ"
            >
              ទាំងអស់
            </button>
          )}
        </div>

        {/* 13 Types Stock Table */}
        <div className="p-3 flex-1 flex flex-col justify-between gap-2.5">
          <div className="bg-white rounded-md border border-gray-200 shadow-2xs overflow-hidden flex-1 flex flex-col">
            <table className="w-full text-xs border-collapse h-full">
              <thead>
                <tr className="bg-gray-100 text-gray-700 font-bold border-b border-gray-200 text-[11.5px]">
                  <th className="py-1.5 px-2.5 text-center w-8 border-r border-gray-200">ល.រ</th>
                  <th className="py-1.5 px-3 text-left border-r border-gray-200">ប្រភេទ</th>
                  <th className="py-1.5 px-3 text-right font-semibold">សន្និធិសល់</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-100 text-[12px]">
                {officeRemainingStockByType.items.map((item, idx) => {
                  const isSelected = selectedType === item.type;
                  return (
                    <tr
                      key={item.type}
                      onClick={() => setSelectedType(selectedType === item.type ? 'all' : item.type)}
                      className={`cursor-pointer transition-colors ${
                        isSelected
                          ? 'bg-orange-100/75 font-semibold text-orange-950'
                          : idx % 2 === 0
                          ? 'bg-white hover:bg-orange-50/50'
                          : 'bg-gray-50/40 hover:bg-orange-50/50'
                      }`}
                      title={`ចុចដើម្បីចម្រាញ់ក្រាហ្វិកតាមប្រភេទ ${item.type}`}
                    >
                      <td className="py-1.5 px-2.5 text-center text-gray-400 font-['Times_New_Roman',_Times,_serif] border-r border-gray-100 text-[11.5px]">
                        {idx + 1}
                      </td>
                      <td className="py-1.5 px-3 font-bold border-r border-gray-100">
                        <span
                          className={`inline-block px-2 py-0.5 rounded text-[11.5px] font-['Times_New_Roman',_Times,_serif] ${
                            isSelected
                              ? 'bg-orange-600 text-white font-black'
                              : 'bg-gray-100 text-gray-800'
                          }`}
                        >
                          {item.type}
                        </span>
                      </td>
                      <td className="py-1.5 px-3 text-right font-['Times_New_Roman',_Times,_serif] font-bold text-gray-900 text-[12.5px]">
                        {Number(item.remaining).toLocaleString('en-US')}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
              <tfoot>
                <tr className="bg-amber-50/90 font-bold border-t-2 border-amber-200 text-amber-950 text-xs">
                  <td colSpan={2} className="py-1.5 px-3 text-left border-r border-amber-200 font-bold">
                    សរុបរួម ៖
                  </td>
                  <td className="py-1.5 px-3 text-right font-['Times_New_Roman',_Times,_serif] text-sm font-black text-amber-950">
                    {Number(officeRemainingStockByType.totalRemaining).toLocaleString('en-US')}
                  </td>
                </tr>
              </tfoot>
            </table>
          </div>

          <div className="pt-1.5 mt-auto border-t border-gray-100 text-[11px] text-gray-400 leading-snug flex items-center justify-between shrink-0">
            <span>* ស្តុកសន្និធិការិយាល័យ</span>
            <span className="text-[10.5px] text-gray-400">ចុចលើជួរដើម្បីចម្រាញ់</span>
          </div>
        </div>
      </div>

      {/* 2. Middle Card: Teams Visa Sticker Remaining Stock Card (កាតសន្និធិសន្លឹកទិដ្ឋាការនៅសល់ បណ្តាក្រុម) */}
      <div className="col-span-1 bg-white rounded-md border border-gray-200 shadow-xs overflow-hidden flex flex-col justify-between">
        {/* Header */}
        <div className="px-3.5 py-2.5 border-b border-gray-200 bg-white flex items-center justify-between shrink-0 gap-2">
          <div className="flex items-center gap-2 min-w-0">
            <div className="p-1.5 bg-blue-100 text-blue-800 rounded-md shrink-0">
              <Users className="w-4 h-4" />
            </div>
            <div className="min-w-0">
              <h4 className="text-sm font-bold text-gray-900 leading-tight whitespace-nowrap">
                សន្និធិ
              </h4>
              <div className="flex items-center gap-1.5 mt-0.5">
                <span
                  className="inline-block px-2 py-0.5 bg-blue-100/80 border border-blue-300 text-blue-900 text-[11px] font-bold rounded truncate max-w-[155px] whitespace-nowrap"
                  title={selectedTeamFilter === 'all' ? 'បណ្តាក្រុម (សរុប)' : selectedTeamFilter}
                >
                  {selectedTeamFilter === 'all' ? 'បណ្តាក្រុម (សរុប)' : selectedTeamFilter}
                </span>
                <span className="text-[10.5px] text-gray-400 shrink-0 whitespace-nowrap">១៣ ប្រភេទ</span>
              </div>
            </div>
          </div>
          <div className="flex items-center gap-1 shrink-0">
            <select
              value={selectedTeamFilter}
              onChange={(e) => setSelectedTeamFilter(e.target.value)}
              className="text-[11px] py-1 px-2 bg-gray-50 hover:bg-gray-100 border border-gray-200 text-gray-800 font-medium rounded transition cursor-pointer focus:outline-hidden focus:ring-1 focus:ring-blue-500 max-w-[130px] truncate"
              title="ជ្រើសរើសក្រុមជាក់លាក់ ឬ សរុបគ្រប់ក្រុម"
            >
              <option value="all">គ្រប់ក្រុម</option>
              {Array.from(OFFICIAL_29_TEAMS).map((t) => (
                <option key={t} value={t}>
                  {t}
                </option>
              ))}
            </select>
          </div>
        </div>

        {/* 13 Types Stock Table for Teams */}
        <div className="p-3 flex-1 flex flex-col justify-between gap-2.5">
          <div className="bg-white rounded-md border border-gray-200 shadow-2xs overflow-hidden flex-1 flex flex-col">
            <table className="w-full text-xs border-collapse h-full">
              <thead>
                <tr className="bg-gray-100 text-gray-700 font-bold border-b border-gray-200 text-[11.5px]">
                  <th className="py-1.5 px-2.5 text-center w-8 border-r border-gray-200">ល.រ</th>
                  <th className="py-1.5 px-3 text-left border-r border-gray-200">ប្រភេទ</th>
                  <th className="py-1.5 px-3 text-right font-semibold">សន្និធិសល់</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-gray-100 text-[12px]">
                {teamsRemainingStockByType.items.map((item, idx) => {
                  const isSelected = selectedType === item.type;
                  return (
                    <tr
                      key={item.type}
                      onClick={() => setSelectedType(selectedType === item.type ? 'all' : item.type)}
                      className={`cursor-pointer transition-colors ${
                        isSelected
                          ? 'bg-blue-100/75 font-semibold text-blue-950'
                          : idx % 2 === 0
                          ? 'bg-white hover:bg-blue-50/50'
                          : 'bg-gray-50/40 hover:bg-blue-50/50'
                      }`}
                      title={`ចុចដើម្បីចម្រាញ់ក្រាហ្វិកតាមប្រភេទ ${item.type}`}
                    >
                      <td className="py-1.5 px-2.5 text-center text-gray-400 font-['Times_New_Roman',_Times,_serif] border-r border-gray-100 text-[11.5px]">
                        {idx + 1}
                      </td>
                      <td className="py-1.5 px-3 font-bold border-r border-gray-100">
                        <span
                          className={`inline-block px-2 py-0.5 rounded text-[11.5px] font-['Times_New_Roman',_Times,_serif] ${
                            isSelected
                              ? 'bg-blue-600 text-white font-black'
                              : 'bg-gray-100 text-gray-800'
                          }`}
                        >
                          {item.type}
                        </span>
                      </td>
                      <td className="py-1.5 px-3 text-right font-['Times_New_Roman',_Times,_serif] font-bold text-gray-900 text-[12.5px]">
                        {Number(item.remaining).toLocaleString('en-US')}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
              <tfoot>
                <tr className="bg-blue-50/90 font-bold border-t-2 border-blue-200 text-blue-950 text-xs">
                  <td colSpan={2} className="py-1.5 px-3 text-left border-r border-blue-200 font-bold">
                    សរុបរួម ៖
                  </td>
                  <td className="py-1.5 px-3 text-right font-['Times_New_Roman',_Times,_serif] text-sm font-black text-blue-950">
                    {Number(teamsRemainingStockByType.totalRemaining).toLocaleString('en-US')}
                  </td>
                </tr>
              </tfoot>
            </table>
          </div>

          <div className="pt-1.5 mt-auto border-t border-gray-100 text-[11px] text-gray-400 leading-snug flex items-center justify-between shrink-0">
            <span className="truncate max-w-[185px]" title={selectedTeamFilter === 'all' ? '* ស្តុកសន្និធិបណ្តាក្រុម' : `* ${selectedTeamFilter}`}>
              {selectedTeamFilter === 'all' ? '* ស្តុកសន្និធិបណ្តាក្រុម' : `* ${selectedTeamFilter}`}
            </span>
            <span className="text-[10.5px] text-gray-400 shrink-0">ចុចលើជួរដើម្បីចម្រាញ់</span>
          </div>
        </div>
      </div>

      {/* 3. Right Card: Monthly Visa Usage Comparison Chart Card (កាតប្រៀបធៀបការប្រើប្រាស់ទិដ្ឋាការប្រចាំខែ) */}
      <div className="col-span-1 md:col-span-2 xl:col-span-2 min-w-0 bg-white rounded-md border border-gray-200 shadow-xs overflow-hidden flex flex-col justify-between">
        {/* Header with Title and Filtering Controls */}
        <div className="px-4 py-3 border-b border-gray-200 bg-white flex flex-col lg:flex-row lg:items-center justify-between gap-3 shrink-0">
        <div>
          <div className="flex items-center gap-2">
            <div className="p-1.5 bg-orange-100/70 text-orange-700 rounded-md">
              <BarChart3 className="w-5 h-5" />
            </div>
            <h3 className="text-base sm:text-lg font-black text-gray-900 tracking-tight">
              ការប្រៀបធៀបការប្រើប្រាស់ទិដ្ឋាការប្រចាំខែ
            </h3>
          </div>
          <p className="text-xs text-gray-500 mt-0.5">
            ទិន្នន័យប្រើប្រាស់ជាក់ស្តែងប្រចាំខែ ប្រៀបធៀបរវាងឆ្នាំ <span className="font-bold text-gray-800">{primaryYear}</span> និង{' '}
            <span className="font-bold text-gray-800">{compareYear}</span> ({typeLabel})
          </p>
        </div>

        {/* Filters and Controls */}
        <div className="flex flex-wrap items-center gap-2.5">
          {/* Filter 1: Type Selection (All, Sticker, cEA, or individual types) */}
          <div className="flex items-center gap-1.5 bg-white border border-gray-300 rounded-md px-2.5 py-1.5 text-xs shadow-2xs">
            <Filter className="w-3.5 h-3.5 text-gray-400" />
            <span className="text-gray-500 font-medium">ប្រភេទ ៖</span>
            <select
              value={selectedType}
              onChange={(e) => setSelectedType(e.target.value)}
              className="bg-transparent font-bold text-gray-900 border-none outline-hidden cursor-pointer text-xs"
            >
              <optgroup label="បូកសរុបរួម">
                <option value="all">Sticker + cEA (បញ្ចូលគ្នាទាំងអស់)</option>
                <option value="sticker_all">សន្លឹកទិដ្ឋាការ (Sticker ទាំងអស់)</option>
                <option value="cea_all">ក្រដាសអនុម័ត (cEA ទាំងអស់)</option>
              </optgroup>
              <optgroup label="តាមប្រភេទ Sticker នីមួយៗ">
                {VISA_TYPES.map((vt) => (
                  <option key={vt} value={vt}>
                    ប្រភេទ «{vt}»
                  </option>
                ))}
              </optgroup>
            </select>
          </div>

          {/* Filter 2: Primary Year */}
          <div className="flex items-center gap-1.5 bg-white border border-gray-300 rounded-md px-2.5 py-1.5 text-xs shadow-2xs">
            <span className="w-2.5 h-2.5 rounded-2xs bg-[#EA580C] shrink-0" title={`ពណ៌របារឆ្នាំ ${primaryYear}`} />
            <Calendar className="w-3.5 h-3.5 text-orange-600" />
            <span className="text-gray-500 font-medium">ឆ្នាំគោល ៖</span>
            <select
              value={primaryYear}
              onChange={(e) => setPrimaryYear(Number(e.target.value))}
              className="bg-transparent font-bold text-gray-900 border-none outline-hidden cursor-pointer text-xs"
            >
              {availableYears.map((y) => (
                <option key={y} value={y}>
                  ឆ្នាំ {y}
                </option>
              ))}
            </select>
          </div>

          {/* Filter 3: Comparison Year */}
          <div className="flex items-center gap-1.5 bg-white border border-gray-300 rounded-md px-2.5 py-1.5 text-xs shadow-2xs">
            <span className="w-2.5 h-2.5 rounded-2xs bg-[#FB923C] shrink-0" title={`ពណ៌របារឆ្នាំ ${compareYear}`} />
            <span className="text-gray-500 font-medium">ធៀបនឹង ៖</span>
            <select
              value={compareYear}
              onChange={(e) => setCompareYear(Number(e.target.value))}
              className="bg-transparent font-bold text-gray-900 border-none outline-hidden cursor-pointer text-xs"
            >
              {availableYears.map((y) => (
                <option key={y} value={y}>
                  ឆ្នាំ {y}
                </option>
              ))}
            </select>
          </div>

          {/* Rotation control: ផ្តេក (0°) */}
          {viewMode === 'chart' && (
            <button
              onClick={() => {
                if (textRotation === 'horizontal') setTextRotation('angle');
                else if (textRotation === 'angle') setTextRotation('up');
                else setTextRotation('horizontal');
              }}
              className="px-2.5 py-1.5 bg-white hover:bg-gray-50 border border-gray-300 text-gray-700 text-xs font-medium rounded-md shadow-2xs cursor-pointer transition flex items-center gap-1.5"
              title="ប្តូរទិសដៅអក្សរ ៖ ផ្តេក (0°) / ទ្រេត (-45°) / បង្វិលឡើងលើ (-90°)"
            >
              <RotateCw className="w-3.5 h-3.5 text-orange-600" />
              <span className="font-semibold">
                {textRotation === 'horizontal'
                  ? 'ផ្តេក (0°)'
                  : textRotation === 'angle'
                  ? 'ទ្រេត (-45°)'
                  : 'បង្វិលឡើងលើ (-90°)'}
              </span>
            </button>
          )}

          {/* Mode switch: Chart vs Table */}
          <div className="flex rounded-md border border-gray-300 p-0.5 bg-gray-100 text-xs">
            <button
              onClick={() => setViewMode('chart')}
              className={`px-2.5 py-1 rounded font-medium flex items-center gap-1 cursor-pointer transition ${
                viewMode === 'chart'
                  ? 'bg-white text-gray-900 shadow-2xs font-bold'
                  : 'text-gray-600 hover:text-gray-900'
              }`}
              title="មើលជាក្រាហ្វិក"
            >
              <BarChart3 className="w-3.5 h-3.5" />
              <span className="hidden sm:inline">ក្រាហ្វិក</span>
            </button>
            <button
              onClick={() => setViewMode('table')}
              className={`px-2.5 py-1 rounded font-medium flex items-center gap-1 cursor-pointer transition ${
                viewMode === 'table'
                  ? 'bg-white text-gray-900 shadow-2xs font-bold'
                  : 'text-gray-600 hover:text-gray-900'
              }`}
              title="មើលជាតារាងទិន្នន័យ"
            >
              <TableIcon className="w-3.5 h-3.5" />
              <span className="hidden sm:inline">តារាង</span>
            </button>
          </div>
        </div>
      </div>

      {/* Summary KPI Cards Strip */}
      <div className="grid grid-cols-2 md:grid-cols-4 gap-3 p-4 bg-gray-50/30 border-b border-gray-100">
        <div className="p-3 bg-white rounded-md border border-gray-200 shadow-2xs">
          <div className="flex items-center gap-1.5">
            <span className="w-2.5 h-2.5 rounded-2xs bg-[#EA580C] shrink-0" />
            <span className="text-[11px] font-semibold text-gray-500 uppercase">
              សរុបប្រើប្រាស់ ឆ្នាំ {primaryYear}
            </span>
          </div>
          <div className="flex items-baseline gap-1 mt-1">
            <span className="text-lg sm:text-xl font-black text-gray-950 font-['Times_New_Roman',_Times,_serif]">
              {Number(primaryTotal).toLocaleString('en-US')}
            </span>
            <span className="text-[11px] text-gray-500">{formatUnit}</span>
          </div>
        </div>

        <div className="p-3 bg-white rounded-md border border-gray-200 shadow-2xs">
          <div className="flex items-center gap-1.5">
            <span className="w-2.5 h-2.5 rounded-2xs bg-[#FB923C] shrink-0" />
            <span className="text-[11px] font-semibold text-gray-500 uppercase">
              សរុបប្រើប្រាស់ ឆ្នាំ {compareYear}
            </span>
          </div>
          <div className="flex items-baseline gap-1 mt-1">
            <span className="text-lg sm:text-xl font-black text-gray-700 font-['Times_New_Roman',_Times,_serif]">
              {Number(compareTotal).toLocaleString('en-US')}
            </span>
            <span className="text-[11px] text-gray-500">{formatUnit}</span>
          </div>
        </div>

        <div className="p-3 bg-white rounded-md border border-gray-200 shadow-2xs">
          <span className="text-[11px] font-semibold text-gray-500 uppercase">
            កម្រិតប្រៀបធៀប (Change)
          </span>
          <div className="flex items-center gap-1.5 mt-1">
            {totalDiff >= 0 ? (
              <span className="inline-flex items-center gap-0.5 text-xs font-black text-emerald-700 bg-emerald-50 px-1.5 py-0.5 rounded">
                <TrendingUp className="w-3.5 h-3.5" />+{growthPercent}%
              </span>
            ) : (
              <span className="inline-flex items-center gap-0.5 text-xs font-black text-rose-700 bg-rose-50 px-1.5 py-0.5 rounded">
                <TrendingDown className="w-3.5 h-3.5" />
                {growthPercent}%
              </span>
            )}
            <span className="text-xs text-gray-600 font-times font-semibold">
              ({totalDiff >= 0 ? '+' : ''}
              {totalDiff.toLocaleString()})
            </span>
          </div>
        </div>

        <div className="p-3 bg-white rounded-md border border-gray-200 shadow-2xs">
          <span className="text-[11px] font-semibold text-gray-500 uppercase">
            ខែប្រើប្រាស់ច្រើនបំផុត ({primaryYear})
          </span>
          <div className="flex items-baseline gap-1 mt-1">
            <span className="text-sm sm:text-base font-bold text-gray-900">
              {peakMonthName || 'គ្មាន'}
            </span>
            {peakMonthVal > 0 && (
              <span className="text-xs font-semibold text-orange-700 font-times">
                ({peakMonthVal.toLocaleString()})
              </span>
            )}
          </div>
        </div>
      </div>

      {/* Main Content: Chart or Table View */}
      {viewMode === 'chart' ? (
        <div className="p-3.5 sm:p-4 flex-1 flex flex-col justify-between">
          {/* Bar Chart Container */}
          <div className="relative flex-1 flex flex-col justify-between">
            {/* Y-axis label */}
            <div className="text-[11px] text-gray-500 mb-1.5 font-medium shrink-0">
              បរិមាណប្រើប្រាស់ ({formatUnit})
            </div>

            {/* Chart Area: Y-Axis + Bars & Dedicated Bottom X-Axis Row */}
            <div className="relative h-[255px] sm:h-[275px] md:h-[295px] flex">
              {/* Y-Axis Labels Column (aligned with bars height, excluding bottom 28px X-axis row) */}
              <div className="w-11 sm:w-14 h-full flex flex-col justify-between shrink-0 pb-7">
                <div className="flex-1 flex flex-col-reverse justify-between text-right pr-2 text-[10.5px] sm:text-[11px] font-['Times_New_Roman',_Times,_serif] font-medium text-gray-400 select-none">
                  {yAxisTicks.map((tick) => (
                    <span key={tick} className="leading-none">
                      {tick >= 1000 ? `${(tick / 1000).toFixed(tick % 1000 === 0 ? 0 : 1)}k` : Number(tick).toLocaleString('en-US')}
                    </span>
                  ))}
                </div>
              </div>

              {/* Chart Main Body: Bars + Gridlines + Bottom X-Axis Row */}
              <div className="relative flex-1 h-full flex flex-col min-w-0">
                {/* Upper Area: Bars & Horizontal Grid Lines */}
                <div className="relative flex-1 w-full min-h-0">
                  {/* Horizontal Grid lines */}
                  <div className="absolute inset-0 flex flex-col-reverse justify-between pointer-events-none">
                    {yAxisTicks.map((tick) => (
                      <div
                        key={tick}
                        className="w-full border-b border-gray-200/80 border-dashed first:border-solid first:border-gray-300"
                      />
                    ))}
                  </div>

                  {/* 12 Months Bar Columns */}
                  <div className="relative z-10 w-full h-full flex items-end justify-between gap-1 sm:gap-1.5 px-0.5 sm:px-1">
                    {KHMER_MONTHS.map((m, idx) => {
                      const valPrimary = primaryMonthlyValues[idx] || 0;
                      const valCompare = compareMonthlyValues[idx] || 0;

                      const heightPrimaryPercent = yMax > 0 ? (valPrimary / yMax) * 100 : 0;
                      const heightComparePercent = yMax > 0 ? (valCompare / yMax) * 100 : 0;

                      const isHovered = hoveredMonthIndex === idx;

                      return (
                        <div
                          key={m.num}
                          className="flex-1 flex items-end justify-center h-full group relative cursor-pointer"
                          onMouseEnter={() => setHoveredMonthIndex(idx)}
                          onMouseLeave={() => setHoveredMonthIndex(null)}
                        >
                          {/* Hover Overlay Tooltip */}
                          {isHovered && (
                            <div className="absolute bottom-full mb-2 z-30 bg-gray-900/95 text-white text-xs rounded-md shadow-xl py-2 px-3 whitespace-nowrap pointer-events-none transition-all duration-200 min-w-[160px]">
                              <div className="font-bold text-amber-300 border-b border-white/15 pb-1 mb-1.5 flex items-center justify-between gap-3">
                                <span>
                                  {m.khmer} ({m.en})
                                </span>
                                <span className="text-[10px] text-gray-300 font-normal">
                                  {formatUnit}
                                </span>
                              </div>
                              <div className="space-y-1">
                                <div className="flex items-center justify-between gap-3">
                                  <span className="flex items-center gap-1.5 text-gray-300">
                                    <span className="w-2.5 h-2.5 rounded-2xs bg-[#EA580C]" />
                                    <span>ឆ្នាំ {primaryYear} ៖</span>
                                  </span>
                                  <span className="font-['Times_New_Roman',_Times,_serif] font-bold text-white text-sm">
                                    {Number(valPrimary).toLocaleString('en-US')}
                                  </span>
                                </div>
                                <div className="flex items-center justify-between gap-3">
                                  <span className="flex items-center gap-1.5 text-gray-300">
                                    <span className="w-2.5 h-2.5 rounded-2xs bg-[#FB923C]" />
                                    <span>ឆ្នាំ {compareYear} ៖</span>
                                  </span>
                                  <span className="font-['Times_New_Roman',_Times,_serif] font-bold text-white text-sm">
                                    {Number(valCompare).toLocaleString('en-US')}
                                  </span>
                                </div>
                                <div className="pt-1 mt-1 border-t border-white/10 flex items-center justify-between text-[11px]">
                                  <span className="text-gray-400">ប្រៀបធៀប ៖</span>
                                  <span
                                    className={`font-['Times_New_Roman',_Times,_serif] font-bold ${
                                      valPrimary >= valCompare ? 'text-emerald-400' : 'text-rose-400'
                                    }`}
                                  >
                                    {valPrimary >= valCompare ? '+' : ''}
                                    {Number(valPrimary - valCompare).toLocaleString('en-US')} (
                                    {valCompare > 0
                                      ? `${(((valPrimary - valCompare) / valCompare) * 100).toFixed(
                                          1
                                        )}%`
                                      : valPrimary > 0
                                      ? '+100%'
                                      : '0%'}
                                    )
                                  </span>
                                </div>
                              </div>
                            </div>
                          )}

                          {/* Dual Bars Container with Numbers directly above each bar */}
                          <div className="w-full flex items-end justify-center gap-0.5 sm:gap-1.5 h-full relative">
                            {/* Bar 1: Compare Year (Lighter Coral/Peach) */}
                            <div className="relative w-1/2 max-w-[15px] sm:max-w-[22px] h-full flex flex-col justify-end items-center">
                              {showBarValues && valCompare > 0 && (
                                <div
                                  style={{
                                    bottom: `calc(${heightComparePercent}% + 3px)`,
                                  }}
                                  className="absolute left-1/2 -translate-x-1/2 z-20 pointer-events-none select-none flex flex-col items-center justify-end"
                                >
                                  <span
                                    style={{
                                      transform:
                                        textRotation === 'up'
                                          ? 'rotate(-90deg)'
                                          : textRotation === 'angle'
                                          ? 'rotate(-45deg)'
                                          : 'none',
                                      transformOrigin:
                                        textRotation === 'horizontal' ? 'center bottom' : 'left center',
                                    }}
                                    className={`font-['Times_New_Roman',_Times,_serif] font-bold whitespace-nowrap leading-none transition-all duration-200 text-gray-800 ${
                                      textRotation === 'horizontal'
                                        ? 'text-[7.5px] sm:text-[8px] md:text-[8.5px] pb-0.5'
                                        : textRotation === 'angle'
                                        ? 'text-[8px] sm:text-[9px] md:text-[9.5px]'
                                        : 'text-[8px] sm:text-[9px] md:text-[10px] tracking-tight'
                                    }`}
                                    title={`ឆ្នាំ ${compareYear}: ${Number(valCompare).toLocaleString('en-US')}`}
                                  >
                                    {formatBarValue(valCompare)}
                                  </span>
                                </div>
                              )}
                              <div
                                style={{ height: `${Math.max(heightComparePercent, valCompare > 0 ? 2 : 0)}%` }}
                                className={`w-full bg-[#FB923C] rounded-t-sm transition-all duration-300 ${
                                  isHovered ? 'brightness-110' : 'opacity-90'
                                }`}
                              />
                            </div>

                            {/* Bar 2: Primary Year (Deeper Coral/Orange-Red) */}
                            <div className="relative w-1/2 max-w-[15px] sm:max-w-[22px] h-full flex flex-col justify-end items-center">
                              {showBarValues && valPrimary > 0 && (
                                <div
                                  style={{
                                    bottom: `calc(${heightPrimaryPercent}% + 3px)`,
                                  }}
                                  className="absolute left-1/2 -translate-x-1/2 z-20 pointer-events-none select-none flex flex-col items-center justify-end"
                                >
                                  <span
                                    style={{
                                      transform:
                                        textRotation === 'up'
                                          ? 'rotate(-90deg)'
                                          : textRotation === 'angle'
                                          ? 'rotate(-45deg)'
                                          : 'none',
                                      transformOrigin:
                                        textRotation === 'horizontal' ? 'center bottom' : 'left center',
                                    }}
                                    className={`font-['Times_New_Roman',_Times,_serif] font-bold whitespace-nowrap leading-none transition-all duration-200 text-orange-950 ${
                                      textRotation === 'horizontal'
                                        ? 'text-[7.5px] sm:text-[8px] md:text-[8.5px] pb-0.5'
                                        : textRotation === 'angle'
                                        ? 'text-[8px] sm:text-[9px] md:text-[9.5px]'
                                        : 'text-[8px] sm:text-[9px] md:text-[10px] tracking-tight'
                                    }`}
                                    title={`ឆ្នាំ ${primaryYear}: ${Number(valPrimary).toLocaleString('en-US')}`}
                                  >
                                    {formatBarValue(valPrimary)}
                                  </span>
                                </div>
                              )}
                              <div
                                style={{ height: `${Math.max(heightPrimaryPercent, valPrimary > 0 ? 2 : 0)}%` }}
                                className={`w-full bg-[#EA580C] rounded-t-sm transition-all duration-300 shadow-2xs ${
                                  isHovered ? 'brightness-110 ring-2 ring-orange-400/40' : ''
                                }`}
                              />
                            </div>
                          </div>
                        </div>
                      );
                    })}
                  </div>
                </div>

                {/* Dedicated Bottom Row: Month Labels (X-Axis) */}
                <div className="h-7 w-full border-t border-gray-300 flex items-center justify-between gap-1 sm:gap-1.5 px-0.5 sm:px-1 shrink-0 bg-gray-50/40">
                  {KHMER_MONTHS.map((m) => (
                    <div
                      key={m.num}
                      className="flex-1 text-center text-[10px] sm:text-[11px] font-semibold text-gray-700 truncate select-none"
                    >
                      <span className="hidden sm:inline">{m.khmer}</span>
                      <span className="sm:hidden">{m.en}</span>
                    </div>
                  ))}
                </div>
              </div>
            </div>

          </div>
        </div>
      ) : (
        /* Table View */
        <div className="p-4 overflow-x-auto">
          <table className="w-full text-xs text-left border-collapse border border-gray-200">
            <thead>
              <tr className="bg-gray-100 text-gray-700 font-bold border-b border-gray-200 text-[11px]">
                <th className="py-1.5 px-2.5 border-r border-gray-200">ខែ</th>
                <th className="py-1.5 px-2.5 border-r border-gray-200 text-right">
                  ឆ្នាំ {compareYear} ({formatUnit})
                </th>
                <th className="py-1.5 px-2.5 border-r border-gray-200 text-right">
                  ឆ្នាំ {primaryYear} ({formatUnit})
                </th>
                <th className="py-1.5 px-2.5 border-r border-gray-200 text-right">ផលសង</th>
                <th className="py-1.5 px-2.5 text-right">អត្រា (%)</th>
              </tr>
            </thead>
            <tbody className="text-[11.5px]">
              {KHMER_MONTHS.map((m, idx) => {
                const valC = compareMonthlyValues[idx] || 0;
                const valP = primaryMonthlyValues[idx] || 0;
                const diff = valP - valC;
                const pct =
                  valC > 0 ? ((diff / valC) * 100).toFixed(1) : valP > 0 ? '+100' : '0.0';

                return (
                  <tr key={m.num} className="border-b border-gray-200 hover:bg-orange-50/40">
                    <td className="py-1.25 px-2.5 border-r border-gray-200 font-bold text-gray-800">
                      {m.khmer} ({m.en})
                    </td>
                    <td className="py-1.25 px-2.5 border-r border-gray-200 text-right font-['Times_New_Roman',_Times,_serif] text-gray-700">
                      {Number(valC).toLocaleString('en-US')}
                    </td>
                    <td className="py-1.25 px-2.5 border-r border-gray-200 text-right font-['Times_New_Roman',_Times,_serif] font-bold text-gray-900">
                      {Number(valP).toLocaleString('en-US')}
                    </td>
                    <td
                      className={`py-1.25 px-2.5 border-r border-gray-200 text-right font-['Times_New_Roman',_Times,_serif] font-bold ${
                        diff >= 0 ? 'text-emerald-700' : 'text-rose-700'
                      }`}
                    >
                      {diff >= 0
                        ? `+${Number(diff).toLocaleString('en-US')}`
                        : Number(diff).toLocaleString('en-US')}
                    </td>
                    <td
                      className={`py-1.25 px-2.5 text-right font-['Times_New_Roman',_Times,_serif] font-bold ${
                        diff >= 0 ? 'text-emerald-700' : 'text-rose-700'
                      }`}
                    >
                      {diff >= 0 ? `+${pct}%` : `${pct}%`}
                    </td>
                  </tr>
                );
              })}
            </tbody>
            <tfoot>
              <tr className="bg-gray-100/90 font-bold text-gray-900 border-t-2 border-gray-300 text-xs">
                <td className="py-1.5 px-2.5 border-r border-gray-200">សរុបរួម</td>
                <td className="py-1.5 px-2.5 border-r border-gray-200 text-right font-['Times_New_Roman',_Times,_serif]">
                  {Number(compareTotal).toLocaleString('en-US')}
                </td>
                <td className="py-1.5 px-2.5 border-r border-gray-200 text-right font-['Times_New_Roman',_Times,_serif] text-orange-700">
                  {Number(primaryTotal).toLocaleString('en-US')}
                </td>
                <td
                  className={`py-1.5 px-2.5 border-r border-gray-200 text-right font-['Times_New_Roman',_Times,_serif] ${
                    totalDiff >= 0 ? 'text-emerald-700' : 'text-rose-700'
                  }`}
                >
                  {totalDiff >= 0
                    ? `+${Number(totalDiff).toLocaleString('en-US')}`
                    : Number(totalDiff).toLocaleString('en-US')}
                </td>
                <td
                  className={`py-1.5 px-2.5 text-right font-['Times_New_Roman',_Times,_serif] ${
                    totalDiff >= 0 ? 'text-emerald-700' : 'text-rose-700'
                  }`}
                >
                  {totalDiff >= 0 ? `+${growthPercent}%` : `${growthPercent}%`}
                </td>
              </tr>
            </tfoot>
          </table>
        </div>
      )}
      </div>
    </div>
  );
};
