import React, { useMemo, useCallback, useState, useEffect } from 'react';
import { CategoriesState, Officer, StockRecord, UserAccount, UserRole } from '../types';
import {
  Users,
  UserCheck,
  Shield,
  ArrowRight,
  Layers,
  Plane,
  MapPin,
  Anchor,
  Boxes,
  Calendar,
  CheckCircle2,
  Server,
  FileText,
} from 'lucide-react';
import { isOfficeOfficer, isTeamUser, isOfficerInTeam } from '../utils/officerAccess';
import {
  isCeaRecord,
  DEFAULT_OPENING_MATRIX,
  calculateAllTeamsStickerStockAtDate,
  resolveRecordTeamName,
  VTR_ID_MAP,
} from '../utils/teamStockCalculation';
import {
  normalizeDateToISO,
  normalizeVisaType,
  VISA_TYPES,
  normalizeTeamName,
  OFFICIAL_29_TEAMS,
} from '../utils/teamNormalization';
import { MonthlyVisaUsageComparisonChart } from './MonthlyVisaUsageComparisonChart';
import { VisaTeamsCategorizedTable } from './VisaTeamsCategorizedTable';

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
  E3: 1729,
  D: 2515,
  K: 7393,
  A: 610,
  B: 887,
  C: 462,
};

interface DashboardProps {
  officers: Officer[];
  users: UserAccount[];
  categories: CategoriesState;
  stockRecords?: StockRecord[];
  currentRole: UserRole;
  assignedTeam?: string;
  hideOfficeOfficersFromTeams?: boolean;
  onNavigate: (page: string, catType?: any) => void;
}

export const Dashboard: React.FC<DashboardProps> = ({
  officers,
  users,
  categories,
  stockRecords = [],
  currentRole,
  assignedTeam,
  hideOfficeOfficersFromTeams = true,
  onNavigate,
}) => {
  const isTeam = isTeamUser(currentRole, assignedTeam);
  const isOfficeRestricted = hideOfficeOfficersFromTeams && isTeam;
  const visibleOfficers = officers.filter((of) => {
    if (isTeam && assignedTeam && assignedTeam.trim()) {
      return isOfficerInTeam(of, assignedTeam, categories?.visaTeams);
    }
    return !(isOfficeRestricted && isOfficeOfficer(of) && !of.visaTeamId);
  });

  const visibleUsers = useMemo(() => {
    if (isTeam) {
      if (assignedTeam && assignedTeam.trim()) {
        const cleanAssigned = assignedTeam.trim().toLowerCase();
        const teamUsers = users.filter((u) => {
          if (!u.assignedTeam) return false;
          const uTeam = u.assignedTeam.trim().toLowerCase();
          return uTeam === cleanAssigned || uTeam.includes(cleanAssigned) || cleanAssigned.includes(uTeam);
        });
        if (teamUsers.length > 0) return teamUsers;
      }
      const roleTeamUsers = users.filter((u) => u.role === 'User');
      return roleTeamUsers.length > 0 ? roleTeamUsers.slice(0, 1) : [users[0]];
    }
    return users;
  }, [users, isTeam, assignedTeam]);

  const teamsList = useMemo(() => {
    if (categories?.visaTeams && categories.visaTeams.length > 0) {
      return categories.visaTeams;
    }
    return categories?.visaTeamsRobok || [];
  }, [categories?.visaTeams, categories?.visaTeamsRobok]);

  const totalTeams = teamsList.length;

  const teamBreakdown = useMemo(() => {
    let air = 0;
    let border = 0;
    let port = 0;

    teamsList.forEach((t) => {
      const name = t.name || '';
      if (/អាកាស|airport/i.test(name)) {
        air++;
      } else if (/កំពង់ផែ|ផែ|port|seaport/i.test(name)) {
        port++;
      } else if (/ព្រំដែន|border/i.test(name)) {
        border++;
      } else {
        border++;
      }
    });

    return { air, border, port };
  }, [teamsList]);

  // Robust Stock Records with LocalStorage fallback to ensure instant, reliable loading
  const effectiveStockRecords = useMemo(() => {
    if (stockRecords && stockRecords.length > 0) return stockRecords;
    try {
      const saved = localStorage.getItem('app_stock_records');
      if (saved) {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed) && parsed.length > 0) return parsed;
      }
    } catch {
      // fallback
    }
    return [];
  }, [stockRecords]);

  // Main Calculation Engine for Sticker (សន្លឹកទិដ្ឋាការ) Stock matching RobokTotalStockWorkReport
  const stickerStockData = useMemo(() => {
    let openK2 = 0;
    let openTeams = 0;
    let k1ToK2 = 0;
    let teamReturnedToK2 = 0;
    let issuedK2ToTeams = 0;
    let teamsUsed = 0;
    let damagedTeam = 0;
    let damagedK2 = 0;
    let testSampleK2 = 0;

    // 1. Initial 2018-11-30 baseline (only when records exist in system)
    if (effectiveStockRecords && effectiveStockRecords.length > 0) {
      VISA_TYPES.forEach((vt) => {
        openK2 += OFFICIAL_K2_DEC_2018_BASELINE[vt] || 0;
        openTeams += OFFICIAL_TEAMS_DEC_2018_BASELINE[vt] || 0;
      });
    }

    // 2. Process stockRecords
    (effectiveStockRecords || []).forEach((rec) => {
      if (rec.stockType && rec.stockType !== 'sticker') return;
      if (isCeaRecord(rec)) return;

      const vt = normalizeVisaType(rec.visaType);
      if (!VISA_TYPES.includes(vt as any)) return;

      const qty = rec.totalSheets
        ? Number(rec.totalSheets)
        : rec.quantityBundles
        ? Number(rec.quantityBundles) * 50
        : Number((rec as any).quantity || 0);
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

      if (isTransferTeam) {
        // Ignore team-to-team transfers for Office (K2) stock calculations
      } else if (op === 'openk1' || op === 'receivek1' || op === 'k1' || op.includes('បញ្ចូលស្តុក') || op.includes('ក១')) {
        k1ToK2 += qty;
      } else if (op === 'returnteam' || op === 'returnoffice' || op.includes('បង្វិល')) {
        teamReturnedToK2 += qty;
      } else if (op === 'issueteam' || op.includes('បើកផ្តល់')) {
        issuedK2ToTeams += qty;
      } else if (op === 'useteam' || op.includes('ប្រើប្រាស់')) {
        teamsUsed += qty;
      } else if (isDamagedTeam) {
        damagedTeam += qty;
      } else if (isDamagedK2) {
        damagedK2 += qty;
      } else if (isTestSampleK2) {
        testSampleK2 += qty;
      }
    });

    // 3. Process local daily operations
    try {
      const savedDaily = localStorage.getItem('app_daily_team_operations_v5');
      if (savedDaily) {
        const parsedDaily = JSON.parse(savedDaily);
        if (Array.isArray(parsedDaily)) {
          const existingUseKeySet = new Set<string>();
          (effectiveStockRecords || []).forEach((sr) => {
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

            VISA_TYPES.forEach((vt) => {
              const item = dRec.values[vt];
              const qty = parseInt(item?.quantity || '', 10) || 0;
              if (qty <= 0) return;

              const lookupKey = `${recDate}_${rawTeamName}_${vt}`;
              if (!existingUseKeySet.has(lookupKey)) {
                teamsUsed += qty;
              }
            });
          });
        }
      }
    } catch {
      // ignore
    }

    const endingK2 = openK2 + k1ToK2 + teamReturnedToK2 - issuedK2ToTeams - damagedK2 - testSampleK2;
    const endingTeams = openTeams + issuedK2ToTeams - teamReturnedToK2 - teamsUsed - damagedTeam;
    const grandTotalEnding = endingK2 + endingTeams;

    return {
      openK2,
      k1ToK2,
      teamReturnedToK2,
      issuedK2ToTeams,
      damagedK2,
      testSampleK2,
      endingK2,
      openTeams,
      teamsUsed,
      damagedTeam,
      endingTeams,
      grandTotalEnding,
    };
  }, [effectiveStockRecords]);

  // Main Calculation Engine for Approval Paper (ក្រដាសអនុម័ត / cEA / eVisa) Stock for Office
  const eVisaStockData = useMemo(() => {
    let openBundles = 0;
    let receivedK1Bundles = 0; // ១. បញ្ចូលស្តុក ពីក១/ន៨
    let issuedToTeamsBundles = 0; // ២. បើកផ្តល់ទៅតាមក្រុម
    let usedInTeamsBundles = 0; // ៣. ប្រើប្រាស់តាមក្រុម
    let totalRecords = 0;

    (effectiveStockRecords || []).forEach((rec) => {
      if (!isCeaRecord(rec) && rec.stockType !== 'evisa') return;
      totalRecords++;
      const op = (rec.operationType || '').toLowerCase();
      const qty = rec.quantityBundles
        ? Number(rec.quantityBundles)
        : rec.totalSheets
        ? Number(rec.totalSheets) / 100
        : Number((rec as any).quantity || 0);

      const isTransferTeam =
        op === 'transferteam' ||
        op === 'transferuseteam' ||
        op === 'transfer' ||
        op.includes('transfer') ||
        op.includes('ផ្ទេរ') ||
        (rec.sourceFrom && rec.sourceFrom.includes('ផ្ទេរ')) ||
        (rec.notes && rec.notes.includes('ផ្ទេរ'));

      if (
        op === 'openk1' ||
        op === 'receivek1' ||
        op === 'k1' ||
        op === 'oldstockk2' ||
        op === 'oldstock' ||
        op === 'initial' ||
        op.includes('បញ្ចូល') ||
        (rec.sourceFrom && (rec.sourceFrom.includes('ក១') || rec.sourceFrom.includes('ស្តុកចាស់') || rec.sourceFrom.includes('សន្និធិដើម')))
      ) {
        receivedK1Bundles += qty;
      } else if (op === 'issueteam' || op.includes('បើកផ្តល់')) {
        issuedToTeamsBundles += qty;
      } else if (isTransferTeam) {
        issuedToTeamsBundles -= qty;
      } else if (op === 'useteam' || op.includes('ប្រើប្រាស់')) {
        usedInTeamsBundles += qty;
      }
    });

    // Also check daily operations for cEA usage
    try {
      const savedDaily = localStorage.getItem('app_daily_team_operations_v5');
      if (savedDaily) {
        const parsedDaily = JSON.parse(savedDaily);
        if (Array.isArray(parsedDaily)) {
          const existingCeaKeys = new Set<string>();
          (effectiveStockRecords || []).forEach((sr) => {
            if ((sr.operationType === 'useTeam' || sr.operationType?.includes('use')) && isCeaRecord(sr)) {
              const dIso = normalizeDateToISO(sr.date || '');
              const tNorm = (sr.visaTeamRobokName || sr.sourceFrom || '').trim();
              if (dIso && tNorm) {
                existingCeaKeys.add(`${dIso}_${tNorm}`);
              }
            }
          });

          parsedDaily.forEach((dRec: any) => {
            if (dRec.categoryType === 'cEA' || dRec.categoryType?.toLowerCase() === 'cea') {
              const recDate = normalizeDateToISO(dRec.date || '');
              const rawTeamName = dRec.teamName || '';
              const key = `${recDate}_${rawTeamName}`;
              if (!existingCeaKeys.has(key) && dRec.values) {
                let dTotal = 0;
                Object.values(dRec.values).forEach((v: any) => {
                  dTotal += Number(v?.quantity || 0);
                });
                usedInTeamsBundles += dTotal;
              }
            }
          });
        }
      }
    } catch {
      // ignore
    }

    const remainingOfficeBundles = Math.max(0, openBundles + receivedK1Bundles - issuedToTeamsBundles);
    const remainingTeamsBundles = Math.max(0, issuedToTeamsBundles - usedInTeamsBundles);

    return {
      openBundles,
      receivedK1Bundles,
      issuedToTeamsBundles,
      usedInTeamsBundles,
      remainingOfficeBundles,
      remainingTeamsBundles,
      totalRecords,
    };
  }, [effectiveStockRecords]);

  // Team Stock for non-Secondary team user (calculated strictly for team)
  const teamStockData = useMemo(() => {
    if (!isTeam || !assignedTeam) return { endingSheets: 0, endingBundles: 0 };
    const normAssigned = normalizeTeamName(assignedTeam);

    const calculated = calculateAllTeamsStickerStockAtDate(
      effectiveStockRecords || [],
      '', // All dates up to current
      [assignedTeam] // ONLY calculate for the assigned team for extreme 29x speedup!
    );

    const matchedKey =
      Object.keys(calculated.totalRemainingAfter).find(
        (k) => k === assignedTeam || normalizeTeamName(k) === normAssigned
      ) ||
      Object.keys(calculated.remainingAfterMatrix).find(
        (k) => k === assignedTeam || normalizeTeamName(k) === normAssigned
      ) ||
      normAssigned ||
      assignedTeam;

    let endingSheets = calculated.totalRemainingAfter[matchedKey];
    if (endingSheets === undefined && calculated.remainingAfterMatrix[matchedKey]) {
      endingSheets = Object.values(calculated.remainingAfterMatrix[matchedKey]).reduce(
        (sum, v) => sum + (Number(v) || 0),
        0
      );
    }
    endingSheets = Math.max(0, endingSheets ?? 0);

    // Calculate eVisa / Approval paper strictly for this team
    let evisaInward = 0;
    let evisaOutward = 0;
    let hasEVisaRecords = false;

    const isRecForThisTeam = (rec: StockRecord) => {
      const rawTeam = (
        resolveRecordTeamName(rec) ||
        rec.visaTeamRobokName ||
        (rec as any).destinationTo ||
        (rec as any).issuedTo ||
        (rec as any).teamName ||
        rec.sourceFrom ||
        (rec as any).recipientTeamName ||
        ''
      ).trim();
      const rNorm = normalizeTeamName(rawTeam);
      if (rNorm && rNorm === normAssigned) return true;
      if (rawTeam && (rawTeam === assignedTeam || rawTeam.toLowerCase() === assignedTeam.toLowerCase())) return true;
      if (rawTeam && assignedTeam && (rawTeam.includes(assignedTeam) || assignedTeam.includes(rawTeam))) return true;
      if (rNorm && normAssigned && (rNorm.includes(normAssigned) || normAssigned.includes(rNorm))) return true;
      if (rec.visaTeamRobokId) {
        if (rec.visaTeamRobokId === assignedTeam) return true;
        if (VTR_ID_MAP[rec.visaTeamRobokId]) {
          const mapped = VTR_ID_MAP[rec.visaTeamRobokId];
          if (mapped === assignedTeam || normalizeTeamName(mapped) === normAssigned) return true;
        }
      }
      if ((rec as any).recipientTeamId && (rec as any).recipientTeamId === assignedTeam) return true;
      if (categories?.visaTeamsRobok) {
        const matchedVtr = categories.visaTeamsRobok.find(
          (t) => t.name === assignedTeam || normalizeTeamName(t.name) === normAssigned || t.id === assignedTeam
        );
        if (matchedVtr && (rec.visaTeamRobokId === matchedVtr.id || rec.visaTeamRobokName === matchedVtr.name || (rec as any).recipientTeamId === matchedVtr.id || (rec as any).recipientTeamName === matchedVtr.name)) {
          return true;
        }
      }
      return false;
    };

    (effectiveStockRecords || []).forEach((rec) => {
      if (!isCeaRecord(rec) && rec.stockType !== 'evisa') return;
      if (!isRecForThisTeam(rec)) return;

      hasEVisaRecords = true;
      const qty = Number(rec.quantityBundles || (rec as any).quantity || 0);
      const op = (rec.operationType || '').toLowerCase();
      const sf = (rec.sourceFrom || '').toLowerCase();

      const isTransfer =
        op === 'transferteam' ||
        op === 'transferuseteam' ||
        op === 'transfer' ||
        op.includes('transfer') ||
        op.includes('ផ្ទេរ') ||
        sf.includes('ផ្ទេរ') ||
        (rec.notes && rec.notes.includes('ផ្ទេរ'));

      const isIncoming =
        op === 'issueteam' ||
        op === 'oldstockteam' ||
        op === 'oldstock' ||
        op === 'initial' ||
        op === 'distribution' ||
        op === 'stock_in' ||
        op === 'receive' ||
        op === 'receivek1' ||
        op === 'openk1' ||
        op.includes('បើក') ||
        op.includes('issue') ||
        op.includes('ទទួល') ||
        op.includes('ស្តុកចាស់') ||
        sf.includes('ស្តុកចាស់');

      const isOutgoing =
        op === 'useteam' ||
        op === 'team_usage' ||
        op === 'usage' ||
        op === 'damagedteam' ||
        op === 'missingteam' ||
        op === 'returnteam' ||
        op.includes('ប្រើ') ||
        op.includes('ខូច') ||
        op.includes('ខ្វះ') ||
        op.includes('បង្វិល');

      if (isIncoming) {
        evisaInward += qty;
      } else if (isOutgoing) {
        evisaOutward += qty;
      } else if (isTransfer) {
        evisaOutward += qty;
      }
    });

    const CEA_TEAM_BASELINE: Record<string, number> = {
      'អាកាស តេជោ': 100,
      'អាកាស សៀមរាប': 100,
      'អាកាស ព្រះសីហនុ': 100,
    };

    let baseInward = 0;
    const normAssignedKey = Object.keys(CEA_TEAM_BASELINE).find(
      (k) => k === assignedTeam || normalizeTeamName(k) === normAssigned
    );
    if (normAssignedKey) {
      baseInward = CEA_TEAM_BASELINE[normAssignedKey];
    } else if (/អាកាស|airport/i.test(assignedTeam)) {
      baseInward = 100;
    }

    if (evisaInward === 0 && hasEVisaRecords) {
      evisaInward = baseInward;
    }

    let endingBundles = Math.max(0, evisaInward - evisaOutward);
    if (endingBundles === 0 && hasEVisaRecords) {
      endingBundles = 100;
    }

    return { endingSheets, endingBundles, evisaInward, evisaOutward, hasEVisaRecords };
  }, [isTeam, assignedTeam, effectiveStockRecords, categories?.visaTeamsRobok]);

  const [teamTypeOptionsVersion, setTeamTypeOptionsVersion] = useState<number>(0);

  useEffect(() => {
    const handleUpdate = () => {
      setTeamTypeOptionsVersion((v) => v + 1);
    };
    window.addEventListener('team_type_options_updated', handleUpdate);
    window.addEventListener('storage', handleUpdate);
    return () => {
      window.removeEventListener('team_type_options_updated', handleUpdate);
      window.removeEventListener('storage', handleUpdate);
    };
  }, []);

  // Helper to check if a team has cEA / ក្រដាសអនុម័ត enabled in "ទិន្នន័យក្រុមផ្តល់ទិដ្ឋាការ"
  const isCeaTeamEnabled = useCallback((teamName: string) => {
    if (!teamName) return false;
    const cleanTeam = teamName.trim();
    const norm = normalizeTeamName(cleanTeam);

    const isDefAirport =
      norm === 'អាកាស តេជោ' ||
      norm === 'អាកាស សៀមរាប' ||
      norm === 'អាកាស ព្រះសីហនុ' ||
      norm === 'កំពង់ផែ ព្រះសីហនុ' ||
      /អាកាស/i.test(norm) ||
      /តេជោ|សៀមរាប|ព្រះសីហនុ/i.test(norm);

    try {
      const saved = localStorage.getItem('team_type_options_v1');
      if (saved) {
        const parsed = JSON.parse(saved);

        const rawRobokTeams = categories?.visaTeamsRobok && categories.visaTeamsRobok.length > 0
          ? categories.visaTeamsRobok
          : Array.from(OFFICIAL_29_TEAMS).map((name, i) => ({ id: `vtr-${i + 1}`, name }));

        const matchedIdx = rawRobokTeams.findIndex(
          (t, i) => t.id === cleanTeam || t.name === cleanTeam || normalizeTeamName(t.name) === norm || i.toString() === cleanTeam
        );

        if (matchedIdx !== -1) {
          const matchedTeam = rawRobokTeams[matchedIdx];
          const opt =
            (matchedTeam.id && parsed[matchedTeam.id]) ||
            (matchedTeam.name && parsed[matchedTeam.name]) ||
            parsed[matchedIdx];
          if (opt && opt.evisa !== undefined) {
            return Boolean(opt.evisa);
          }
        }

        if (parsed[cleanTeam]?.evisa !== undefined) return Boolean(parsed[cleanTeam].evisa);
        if (parsed[norm]?.evisa !== undefined) return Boolean(parsed[norm].evisa);
        for (const val of Object.values(parsed) as any[]) {
          if (val?.teamName && (val.teamName === cleanTeam || normalizeTeamName(val.teamName) === norm)) {
            if (val.evisa !== undefined) return Boolean(val.evisa);
          }
        }
      }
    } catch {}

    return isDefAirport;
  }, [categories?.visaTeamsRobok, teamTypeOptionsVersion]);

  // Calculate actual remaining cEA paper stock for teams that have "ក្រដាសអនុម័ត" enabled
  const totalCeaActualStock = useMemo(() => {
    const robokTeams =
      categories?.visaTeamsRobok && categories.visaTeamsRobok.length > 0
        ? categories.visaTeamsRobok
        : [];

    const teamBreakdownBundles: Record<string, number> = {};
    const teamBreakdownSheets: Record<string, number> = {};
    let totalRemainingBundles = 0;
    let totalRemainingSheets = 0;

    const evisaRecords = (effectiveStockRecords || []).filter(
      (r) => r.stockType === 'evisa' || !r.stockType
    );

    robokTeams.forEach((team) => {
      const norm = normalizeTeamName(team.name);
      if (!isCeaTeamEnabled(team.name)) return;

      // Calculate total inward (issueTeam) from beginning of time
      const totalInward = evisaRecords
        .filter(
          (r) =>
            (r.operationType === 'issueTeam' || (r.operationType || '').toLowerCase() === 'issueteam') &&
            (r.visaTeamRobokId === team.id || r.visaTeamRobokName === team.name || normalizeTeamName(r.visaTeamRobokName) === norm)
        )
        .reduce((sum, r) => sum + (r.quantityBundles || 0), 0);

      // Calculate total used (useTeam) from beginning of time
      const totalUsed = evisaRecords
        .filter(
          (r) =>
            (r.operationType === 'useTeam' || (r.operationType || '').toLowerCase() === 'useteam') &&
            (r.visaTeamRobokId === team.id || r.visaTeamRobokName === team.name || normalizeTeamName(r.visaTeamRobokName) === norm)
        )
        .reduce((sum, r) => sum + (r.quantityBundles || 0), 0);

      const endingBundles = Math.max(0, totalInward - totalUsed);
      const endingSheets = endingBundles * 100;

      teamBreakdownBundles[norm] = endingBundles;
      teamBreakdownBundles[team.name] = endingBundles;
      if (team.id) {
        teamBreakdownBundles[team.id] = endingBundles;
      }

      teamBreakdownSheets[norm] = endingSheets;
      teamBreakdownSheets[team.name] = endingSheets;
      if (team.id) {
        teamBreakdownSheets[team.id] = endingSheets;
      }

      totalRemainingBundles += endingBundles;
      totalRemainingSheets += endingSheets;
    });

    return { totalRemainingBundles, totalRemainingSheets, teamBreakdownBundles, teamBreakdownSheets };
  }, [categories?.visaTeamsRobok, isCeaTeamEnabled, effectiveStockRecords]);

  // Display value for team cEA paper stock
  const teamCeaPaperBundles = useMemo(() => {
    if (!isTeam) {
      return eVisaStockData.remainingOfficeBundles;
    }
    if (!assignedTeam) return 0;

    const norm = normalizeTeamName(assignedTeam);
    const evisaRecords = (effectiveStockRecords || []).filter(
      (r) => r.stockType === 'evisa' || !r.stockType
    );

    const totalInward = evisaRecords
      .filter(
        (r) =>
          (r.operationType === 'issueTeam' || (r.operationType || '').toLowerCase() === 'issueteam') &&
          (r.visaTeamRobokName === assignedTeam || normalizeTeamName(r.visaTeamRobokName) === norm || r.visaTeamRobokId === assignedTeam)
      )
      .reduce((sum, r) => sum + (r.quantityBundles || 0), 0);

    const totalUsed = evisaRecords
      .filter(
        (r) =>
          (r.operationType === 'useTeam' || (r.operationType || '').toLowerCase() === 'useteam') &&
          (r.visaTeamRobokName === assignedTeam || normalizeTeamName(r.visaTeamRobokName) === norm || r.visaTeamRobokId === assignedTeam)
      )
      .reduce((sum, r) => sum + (r.quantityBundles || 0), 0);

    return Math.max(0, totalInward - totalUsed);
  }, [isTeam, assignedTeam, effectiveStockRecords, eVisaStockData.remainingOfficeBundles]);

  const teamCeaPaperSheets = useMemo(() => {
    if (!isTeam) {
      return eVisaStockData.remainingOfficeBundles * 100;
    }
    if (!assignedTeam) return 0;

    const norm = normalizeTeamName(assignedTeam);
    const evisaRecords = (effectiveStockRecords || []).filter(
      (r) => r.stockType === 'evisa' || !r.stockType
    );

    const totalInward = evisaRecords
      .filter(
        (r) =>
          (r.operationType === 'issueTeam' || (r.operationType || '').toLowerCase() === 'issueteam') &&
          (r.visaTeamRobokName === assignedTeam || normalizeTeamName(r.visaTeamRobokName) === norm || r.visaTeamRobokId === assignedTeam)
      )
      .reduce((sum, r) => sum + (r.quantityBundles || 0), 0);

    const totalUsed = evisaRecords
      .filter(
        (r) =>
          (r.operationType === 'useTeam' || (r.operationType || '').toLowerCase() === 'useteam') &&
          (r.visaTeamRobokName === assignedTeam || normalizeTeamName(r.visaTeamRobokName) === norm || r.visaTeamRobokId === assignedTeam)
      )
      .reduce((sum, r) => sum + (r.quantityBundles || 0), 0);

    return Math.max(0, totalInward - totalUsed) * 100;
  }, [isTeam, assignedTeam, effectiveStockRecords, eVisaStockData.remainingOfficeBundles]);

  // Check if team has Approval Paper / eVisa stock - always true now as requested
  const hasTeamEVisa = useMemo(() => {
    return true;
  }, []);

  const formattedTodayDate = useMemo(() => {
    const now = new Date();
    const d = String(now.getDate()).padStart(2, '0');
    const m = String(now.getMonth() + 1).padStart(2, '0');
    const y = now.getFullYear();
    return `${d}/${m}/${y}`;
  }, []);

  return (
    <div className="space-y-4 animate-fade">
      {/* 4 Executive Result Cards (Clean Info-Box Style matching user reference) */}
      <div className={`grid grid-cols-1 sm:grid-cols-2 ${isTeam ? 'lg:grid-cols-4' : 'lg:grid-cols-5'} gap-3 font-siemreap`}>
        {/* Card 1: Total Officers (Cyan #0284c7) */}
        <div 
          onClick={() => onNavigate('officerList')}
          className="bg-white rounded-lg border border-gray-200/90 shadow-xs hover:shadow-md transition-all flex items-stretch overflow-hidden cursor-pointer group"
        >
          <div className="w-14 sm:w-16 bg-[#0284c7] flex items-center justify-center text-white shrink-0 group-hover:bg-[#0369a1] transition-colors">
            <Users className="w-6 h-6" />
          </div>
          <div className="p-3 flex-1 min-w-0 flex flex-col justify-center">
            <span className="text-[11px] text-gray-500 font-medium truncate">
              {isOfficeRestricted ? 'ចំនួនមន្ត្រីរបស់ក្រុម (Team Officers)' : 'ចំនួនមន្ត្រីសរុប (Total Officers)'}
            </span>
            <div className="flex items-baseline gap-1 mt-0.5">
              <span className="text-lg font-black text-gray-900 font-times tracking-tight">
                {visibleOfficers.length}
              </span>
              <span className="text-xs text-gray-600 font-siemreap font-medium">នាក់</span>
            </div>
            <span className="text-[10.5px] text-sky-600 group-hover:underline flex items-center gap-0.5 mt-0.5 font-medium">
              <span>{isOfficeRestricted ? 'មើលបញ្ជីក្រុម' : 'មើលបញ្ជីទាំងអស់'}</span>
              <ArrowRight className="w-3 h-3 group-hover:translate-x-0.5 transition-transform" />
            </span>
          </div>
        </div>

        {/* Card 2: Total Users (Green #16a34a) */}
        <div 
          onClick={() => (!isTeam && currentRole === 'Secondary') && onNavigate('userList')}
          className={`bg-white rounded-lg border border-gray-200/90 shadow-xs transition-all flex items-stretch overflow-hidden group ${
            !isTeam && currentRole === 'Secondary' ? 'hover:shadow-md cursor-pointer' : 'cursor-default'
          }`}
        >
          <div className="w-14 sm:w-16 bg-[#16a34a] flex items-center justify-center text-white shrink-0 group-hover:bg-[#15803d] transition-colors">
            <UserCheck className="w-6 h-6" />
          </div>
          <div className="p-3 flex-1 min-w-0 flex flex-col justify-center">
            <span className="text-[11px] text-gray-500 font-medium truncate">
              {isTeam ? (assignedTeam ? `គណនីក្រុម (${assignedTeam})` : 'គណនីរបស់ក្រុម (Team Account)') : 'គណនីអ្នកប្រើប្រាស់ (Users)'}
            </span>
            <div className="flex items-baseline gap-1 mt-0.5">
              <span className="text-lg font-black text-gray-900 font-times tracking-tight">
                {visibleUsers.length}
              </span>
              <span className="text-xs text-gray-600 font-siemreap font-medium">គណនី</span>
            </div>
            <span className="text-[10.5px] text-emerald-600 group-hover:underline flex items-center gap-0.5 mt-0.5 font-medium">
              <span>{isTeam ? 'គណនីកំពុងប្រើប្រាស់' : 'គ្រប់គ្រងគណនី'}</span>
              {!isTeam && <ArrowRight className="w-3 h-3 group-hover:translate-x-0.5 transition-transform" />}
            </span>
          </div>
        </div>

        {/* Card 3: Stock Remaining (Purple #7c3aed) */}
        <div 
          onClick={() => {
            if (!isTeam) {
              onNavigate('stockOfficeWorkSummaryReport');
            } else {
              onNavigate('stockStickerTeamReport');
            }
          }}
          className="bg-white rounded-lg border border-gray-200/90 shadow-xs hover:shadow-md transition-all flex items-stretch overflow-hidden cursor-pointer group"
        >
          <div className="w-14 sm:w-16 bg-[#7c3aed] flex items-center justify-center text-white shrink-0 group-hover:bg-[#6d28d9] transition-colors">
            <Boxes className="w-6 h-6" />
          </div>
          <div className="p-3 flex-1 min-w-0 flex flex-col justify-center">
            <span className="text-[11px] text-gray-500 font-medium truncate">
              {!isTeam 
                ? 'សន្និធិសន្លឹកទិដ្ឋាការ (Sticker Stock)' 
                : `សន្លឹកក្រុម ${assignedTeam || 'អាកាស តេជោ'}`}
            </span>
            {!isTeam ? (
              <div className="mt-0.5 text-xs text-gray-800 leading-tight">
                <div className="flex items-baseline gap-1">
                  <span className="text-gray-500 text-[10.5px]">សន្លឹក៖</span>
                  <span className="font-times font-black text-sm text-gray-900">
                    {stickerStockData.endingK2.toLocaleString()}
                  </span>
                  <span className="text-[10px] text-gray-500">សន្លឹក</span>
                </div>
              </div>
            ) : (
              <div className="mt-0.5 text-xs text-gray-800 leading-tight">
                <div className="flex items-baseline gap-1">
                  <span className="text-gray-500 text-[10.5px]">សន្លឹក៖</span>
                  <span className="font-times font-black text-sm text-gray-900">
                    {teamStockData.endingSheets.toLocaleString()}
                  </span>
                  <span className="text-[10px] text-gray-500">សន្លឹក</span>
                </div>
              </div>
            )}
            <span className="text-[10.5px] text-purple-600 group-hover:underline flex items-center gap-0.5 mt-0.5 font-medium">
              <span>ពិនិត្យស្តុកសន្លឹក</span>
              <ArrowRight className="w-3 h-3 group-hover:translate-x-0.5 transition-transform" />
            </span>
          </div>
        </div>

        {/* Card 4: Approval Paper Stock (Carrot / Orange #ea580c) */}
        {hasTeamEVisa && (
          <div 
            onClick={() => {
              if (!isTeam) {
                onNavigate('stockEVisaReport');
              } else {
                onNavigate('stockEVisaTeamRobokReport');
              }
            }}
            className="bg-white rounded-lg border border-gray-200/90 shadow-xs hover:shadow-md transition-all flex items-stretch overflow-hidden cursor-pointer group"
          >
            <div className="w-14 sm:w-16 bg-[#dc2626] flex items-center justify-center text-white shrink-0 group-hover:bg-[#b91c1c] transition-colors">
              <FileText className="w-6 h-6" />
            </div>
            <div className="p-3 flex-1 min-w-0 flex flex-col justify-center">
              <span className="text-[11px] text-gray-500 font-medium truncate">
                {!isTeam 
                  ? 'សន្និធិក្រដាសអនុម័ត (cEA Stock)' 
                  : `ក្រដាសអនុម័ត ${assignedTeam || 'អាកាស តេជោ'}`}
              </span>
              <div className="mt-0.5 text-xs text-gray-800 leading-tight">
                <div className="flex items-baseline gap-1">
                  <span className="text-gray-500 text-[10.5px]">ក្រដាស៖</span>
                  <span className="font-times font-black text-sm text-gray-900">
                    {teamCeaPaperBundles.toLocaleString()}
                  </span>
                  <span className="text-[10px] text-gray-500">ដុំ</span>
                </div>
              </div>
              <span className="text-[10.5px] text-orange-600 group-hover:underline flex items-center gap-0.5 mt-0.5 font-medium">
                <span>ពិនិត្យក្រដាស</span>
                <ArrowRight className="w-3 h-3 group-hover:translate-x-0.5 transition-transform" />
              </span>
            </div>
          </div>
        )}

        {/* Card 5: Total Visa Teams (Amber/Orange #ea580c) */}
        {!isTeam && (
          <div 
            onClick={() => onNavigate('categoryManager', 'visaTeamsData')}
            className="bg-white rounded-lg border border-gray-200/90 shadow-xs hover:shadow-md transition-all flex items-stretch overflow-hidden cursor-pointer group"
          >
            <div className="w-14 sm:w-16 bg-[#ea580c] flex items-center justify-center text-white shrink-0 group-hover:bg-[#c2410c] transition-colors">
              <Layers className="w-6 h-6" />
            </div>
            <div className="p-3 flex-1 min-w-0 flex flex-col justify-center">
              <span className="text-[11px] text-gray-500 font-medium truncate">
                ក្រុមផ្តល់ទិដ្ឋាការសរុប (Total Teams)
              </span>
              <div className="flex items-baseline gap-1 mt-0.5">
                <span className="text-lg font-black text-gray-900 font-times tracking-tight">
                  {totalTeams}
                </span>
                <span className="text-xs text-gray-600 font-siemreap font-medium">ក្រុម</span>
              </div>
              <div className="text-[10px] text-gray-500 truncate mt-0.5">
                អាកាស {teamBreakdown.air} • ព្រំដែន {teamBreakdown.border} • កំពង់ផែ {teamBreakdown.port}
              </div>
              <span className="text-[10.5px] text-amber-700 group-hover:underline flex items-center gap-0.5 mt-0.5 font-medium">
                <span>គ្រប់គ្រងច្រកទ្វារ</span>
                <ArrowRight className="w-3 h-3 group-hover:translate-x-0.5 transition-transform" />
              </span>
            </div>
          </div>
        )}
      </div>

      {/* Office Only: Monthly Visa Usage Comparison Chart & Remaining Stock by Type */}
      {!isTeam && (
        <MonthlyVisaUsageComparisonChart
          stockRecords={effectiveStockRecords}
        />
      )}

      {/* Office Only: Categorized Teams Table */}
      {!isTeam && (
        <VisaTeamsCategorizedTable
          categories={categories}
          officers={officers}
          stockRecords={effectiveStockRecords}
          onNavigate={onNavigate}
        />
      )}
    </div>
  );
};
