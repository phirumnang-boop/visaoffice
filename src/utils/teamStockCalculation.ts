import { StockRecord } from '../types';
import {
  OFFICIAL_29_TEAMS,
  VISA_TYPES,
  normalizeDateToISO,
  normalizeVisaType,
  normalizeTeamName,
  matchTeamInList,
} from './teamNormalization';

// Default Baseline Values (Starts clean at 0 / empty until user records are uploaded or entered)
export const DEFAULT_OPENING_MATRIX: Record<string, Record<string, number>> = {
  'អាកាស តេជោ': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'អាកាស សៀមរាប': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'អាកាស ព្រះសីហនុ': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន ប៉ោយប៉ែត': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន បាវិត': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន ចាំយាម': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន ដូង': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន អូរស្មាច់': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន ព្រំ': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន បន្ទាយចក្រី': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន ជាំ': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន ត្រពាំងស្រែ': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន ត្រពាំងក្រៀល': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន ត្រពាំងផ្លុង': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន ភ្នំដិន': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន កោះរកា': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន ព្រៃវល្លិ៍': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន អូរយ៉ាដាវ': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន ក្អមសំណ': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន ភ្នំដី': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'កំពង់ផែ ឧកញ៉ាម៉ុង': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'កំពង់ផែ ស្ទឹងហាវ': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'កំពង់ផែ ព្រះសីហនុ': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'កំពង់ផែ ភ្នំពេញ': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន ព្រែកចាក': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន ម៉ឺនជ័យ': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'ព្រំដែន ស្ទឹងបត់': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'កំពង់ផែ កោះកុង': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
  'កំពង់ផែ កំពត': { T: 0, T1: 0, T2: 0, T3: 0, E: 0, E1: 0, E2: 0, E3: 0, D: 0, K: 0, A: 0, B: 0, C: 0 },
};

export const isCeaRecord = (rec: StockRecord): boolean => {
  if (rec.stockType === 'evisa' || rec.stockType === 'cea' || rec.stockType === 'cEA') return true;
  const sf = (rec.sourceFrom || '').trim().toLowerCase();
  const id = (rec.id || '').toLowerCase();
  const vt = (rec.visaType || '').trim().toLowerCase();
  const op = (rec.operationType || '').toLowerCase();
  const ct = ((rec as any).categoryType || '').trim().toLowerCase();
  return (
    sf === 'cea' ||
    sf.includes('cea') ||
    ct === 'cea' ||
    ct.includes('cea') ||
    vt === 'cea' ||
    vt.includes('cea') ||
    vt.includes('ក្រដាសអនុម័ត') ||
    vt.includes('evisa') ||
    id.includes('-cea-') ||
    op.includes('cea')
  );
};

export const isOldStockTeamRecord = (rec: StockRecord): boolean => {
  if (rec.stockType && rec.stockType !== 'sticker') return false;
  if (isCeaRecord(rec)) return false;
  const op = (rec.operationType || '').trim().toLowerCase();
  const sf = (rec.sourceFrom || '').trim();

  if (
    op === 'oldstockteam' ||
    op === 'old_stock_team' ||
    op.includes('oldstockteam') ||
    op.includes('ស្តុកចាស់ក្រុម')
  ) {
    return true;
  }

  if (
    sf === 'ស្តុកចាស់ក្រុម' ||
    sf === 'ស្តុកចាស់របស់ក្រុម' ||
    sf.includes('ស្តុកចាស់ក្រុម') ||
    sf.includes('ស្តុកចាស់របស់ក្រុម') ||
    sf.includes('សន្និធិដើមក្រុម') ||
    sf.includes('ស្តុកចាស់តាមក្រុម')
  ) {
    return true;
  }

  const hasTeam = Boolean(
    rec.visaTeamRobokName ||
    rec.visaTeamRobokId ||
    (rec as any).team ||
    (rec as any).teamName ||
    sf.includes('ក្រុម') ||
    sf.includes('ប៉ុស្តិ៍') ||
    sf.includes('ព្រំដែន') ||
    sf.includes('អាកាស') ||
    sf.includes('កំពង់ផែ')
  );

  if (sf.includes('ស្តុកចាស់') && hasTeam && !sf.includes('ក២') && !sf.includes('ក១')) {
    return true;
  }

  return false;
};

export const VTR_ID_MAP: Record<string, string> = {
  'vtr-1': 'អាកាស តេជោ',
  'vtr-2': 'អាកាស សៀមរាប',
  'vtr-3': 'អាកាស ព្រះសីហនុ',
  'vtr-4': 'ព្រំដែន ប៉ោយប៉ែត',
  'vtr-5': 'ព្រំដែន បាវិត',
  'vtr-6': 'ព្រំដែន ចាំយាម',
  'vtr-7': 'ព្រំដែន ដូង',
  'vtr-8': 'ព្រំដែន អូរស្មាច់',
  'vtr-9': 'ព្រំដែន ព្រំ',
  'vtr-10': 'ព្រំដែន បន្ទាយចក្រី',
  'vtr-11': 'ព្រំដែន ជាំ',
  'vtr-12': 'ព្រំដែន ត្រពាំងស្រែ',
  'vtr-13': 'ព្រំដែន ត្រពាំងក្រៀល',
  'vtr-14': 'ព្រំដែន ត្រពាំងផ្លុង',
  'vtr-15': 'ព្រំដែន ភ្នំដិន',
  'vtr-16': 'ព្រំដែន កោះរកា',
  'vtr-17': 'ព្រំដែន ព្រៃវល្លិ៍',
  'vtr-18': 'ព្រំដែន អូរយ៉ាដាវ',
  'vtr-19': 'ព្រំដែន ក្អមសំណ',
  'vtr-20': 'ព្រំដែន ភ្នំដី',
  'vtr-21': 'កំពង់ផែ ឧកញ៉ាម៉ុង',
  'vtr-22': 'កំពង់ផែ ស្ទឹងហាវ',
  'vtr-23': 'កំពង់ផែ ព្រះសីហនុ',
  'vtr-24': 'កំពង់ផែ ភ្នំពេញ',
  'vtr-25': 'ព្រំដែន ព្រែកចាក',
  'vtr-26': 'ព្រំដែន ម៉ឺនជ័យ',
  'vtr-27': 'ព្រំដែន ស្ទឹងបត់',
  'vtr-28': 'កំពង់ផែ កោះកុង',
  'vtr-29': 'កំពង់ផែ កំពត',
};

export const resolveRecordTeamName = (r: StockRecord, fallbackTeams?: string[]): string => {
  let raw = (
    r.visaTeamRobokName ||
    (r as any).issuedTo ||
    (r as any).teamName ||
    (r as any).destinationTo ||
    (r as any).to ||
    (r as any).targetTeam ||
    (r as any).team ||
    (r as any).visaTeam ||
    (r as any).station ||
    (r as any).port ||
    (r as any).recipient ||
    ''
  ).trim();

  if (!raw && r.visaTeamRobokId) {
    raw = r.visaTeamRobokId;
  }

  if (raw && VTR_ID_MAP[raw]) {
    return VTR_ID_MAP[raw];
  }

  if (!raw && r.sourceFrom && isOldStockTeamRecord(r)) {
    const cleaned = r.sourceFrom
      .replace(/ស្តុកចាស់របស់ក្រុម|ស្តុកចាស់ក្រុម|សន្និធិដើមក្រុម|ស្តុកចាស់|សន្និធិដើម|[\-:_]/g, '')
      .trim();
    if (cleaned) {
      raw = cleaned;
    }
  }

  if (raw && VTR_ID_MAP[raw]) {
    return VTR_ID_MAP[raw];
  }

  return raw.trim();
};

// Helper to check if a StockRecord belongs to a specific team (as primary owner or sender)
export const isRecordForTeam = (
  r: StockRecord | any,
  targetTeamName?: string
): boolean => {
  if (!r || isExcludedFromTeamStockFormula(r)) return false;
  if (!targetTeamName || targetTeamName === 'ផ្នែករដ្ឋបាល' || targetTeamName === 'all' || targetTeamName.toLowerCase() === 'all') {
    return true;
  }
  const targetNorm = normalizeTeamName(targetTeamName);
  if (!targetNorm) return true;

  const rawTeam = resolveRecordTeamName(r);
  const normTeam = normalizeTeamName(rawTeam);
  const matchedOfficial = matchTeamInList(rawTeam, Array.from(OFFICIAL_29_TEAMS));
  const targetOfficial = matchTeamInList(targetTeamName, Array.from(OFFICIAL_29_TEAMS));

  if (rawTeam && rawTeam === targetTeamName) return true;
  if (normTeam && normTeam === targetNorm) return true;
  if (matchedOfficial && targetOfficial && matchedOfficial === targetOfficial) return true;
  if (normTeam && targetNorm && (normTeam.includes(targetNorm) || targetNorm.includes(normTeam))) return true;

  if (r.visaTeamRobokId && normalizeTeamName(r.visaTeamRobokId) === targetNorm) return true;
  if (r.visaTeamRobokName && normalizeTeamName(r.visaTeamRobokName) === targetNorm) return true;
  if (isOldStockTeamRecord(r) && r.sourceFrom && normalizeTeamName(r.sourceFrom) === targetNorm) return true;
  if ((r as any).teamName && normalizeTeamName((r as any).teamName) === targetNorm) return true;

  // Only check destination/issuedTo if record is NOT a team transfer to another recipient
  if (!r.recipientTeamName && !r.recipientTeamId) {
    if ((r as any).issuedTo && normalizeTeamName((r as any).issuedTo) === targetNorm) return true;
    if ((r as any).destinationTo && normalizeTeamName((r as any).destinationTo) === targetNorm) return true;
    if ((r as any).targetTeam && normalizeTeamName((r as any).targetTeam) === targetNorm) return true;
  }

  return false;
};

// Helper to check if a StockRecord belongs to a specific team as RECIPIENT of a transfer
export const isRecordForRecipientTeam = (
  r: StockRecord | any,
  targetTeamName?: string
): boolean => {
  if (!targetTeamName || targetTeamName === 'ផ្នែករដ្ឋបាល' || targetTeamName === 'all' || targetTeamName.toLowerCase() === 'all') {
    return false;
  }
  const targetNorm = normalizeTeamName(targetTeamName);
  if (!targetNorm) return false;

  const targetOfficial = matchTeamInList(targetTeamName, Array.from(OFFICIAL_29_TEAMS));

  const recipName = (r.recipientTeamName || (r as any).destinationTo || (r as any).targetTeam || '').trim();
  const recipId = (r.recipientTeamId || '').trim();

  if (recipName) {
    if (recipName === targetTeamName) return true;
    const recipNorm = normalizeTeamName(recipName);
    if (recipNorm && recipNorm === targetNorm) return true;
    const recipOfficial = matchTeamInList(recipName, Array.from(OFFICIAL_29_TEAMS));
    if (recipOfficial && targetOfficial && recipOfficial === targetOfficial) return true;
    if (recipNorm && targetNorm && (recipNorm.includes(targetNorm) || targetNorm.includes(recipNorm))) return true;
  }

  if (recipId) {
    if (recipId === targetTeamName || recipId.toLowerCase() === targetTeamName.toLowerCase()) return true;
    const idNorm = normalizeTeamName(recipId);
    if (idNorm && idNorm === targetNorm) return true;
    const idOfficial = matchTeamInList(recipId, Array.from(OFFICIAL_29_TEAMS));
    if (idOfficial && targetOfficial && idOfficial === targetOfficial) return true;
    if (VTR_ID_MAP && VTR_ID_MAP[recipId]) {
      const mapped = VTR_ID_MAP[recipId];
      if (mapped === targetTeamName || normalizeTeamName(mapped) === targetNorm) return true;
    }
  }

  return false;
};

// Helper to check if a record is a transfer operation
export const isTransferTeamRecord = (r: StockRecord | any): boolean => {
  const op = (r.operationType || '').trim().toLowerCase();
  const src = (r.sourceFrom || '').trim().toLowerCase();
  const notes = ((r as any).notes || '').trim().toLowerCase();
  return (
    op === 'transferteam' ||
    op === 'transferuseteam' ||
    op === 'transfer' ||
    op.includes('transfer') ||
    op.includes('ផ្ទេរ') ||
    src.includes('ផ្ទេរ') ||
    notes.includes('ផ្ទេរ') ||
    Boolean(r.recipientTeamName || r.recipientTeamId)
  );
};

/**
 * Strictly excludes operations from ទិន្នន័យសន្លឹកទិដ្ឋាការ (ការងារស្តុក) that must NEVER be included
 * in ប្រតិបត្តិការប្រចាំថ្ងៃតាមក្រុម or its remaining stock formula:
 * - ប្រមូលគល់សន្លឹកទិដ្ឋាការ (returnStub / stub collection)
 * - ការបញ្ចូលស្តុក ក១ (openK1)
 * - ស្តុកចាស់ ក២ (oldStockK2)
 * - ទិដ្ឋាការសាកក២ (testPrintK2)
 * - ទិដ្ឋាការខូចក២ (damaged / damagedK2)
 * - បង្វិលក១ (returnK1)
 */
export const isExcludedFromTeamStockFormula = (r: StockRecord | any): boolean => {
  if (!r) return true;
  const id = String(r.id || '').trim().toLowerCase();
  const op = String(r.operationType || '').trim().toLowerCase();
  const sf = String(r.sourceFrom || '').trim().toLowerCase();
  const rem = String(r.remarks || '').trim().toLowerCase();
  const notes = String((r as any).notes || '').trim().toLowerCase();

  // 1. Stub collection (ប្រមូលគល់សន្លឹកទិដ្ឋាការ / ការប្រមូលគល់ទិដ្ឋាការ)
  if (
    id.startsWith('stock-stub-') ||
    id.includes('-stub-') ||
    op === 'returnstub' ||
    op === 'collectstub' ||
    op === 'stubcollection' ||
    op.includes('stub') ||
    op.includes('គល់សន្លឹក') ||
    op.includes('ប្រមូលគល់') ||
    op.includes('គល់ទិដ្ឋាការ') ||
    sf.includes('គល់សន្លឹក') ||
    sf.includes('ប្រមូលគល់') ||
    sf.includes('គល់ទិដ្ឋាការ') ||
    rem.includes('គល់សន្លឹក') ||
    rem.includes('ប្រមូលគល់') ||
    rem.includes('គល់ទិដ្ឋាការ') ||
    notes.includes('គល់សន្លឹក') ||
    notes.includes('ប្រមូលគល់') ||
    notes.includes('គល់ទិដ្ឋាការ')
  ) {
    return true;
  }

  // 2. Office K1 / K2 operations (openK1, oldStockK2, returnK1, testPrintK2, damaged K2)
  if (
    op === 'openk1' ||
    op === 'open_k1' ||
    op === 'k1' ||
    op === 'receive' ||
    op === 'received' ||
    op === 'oldstockk2' ||
    op === 'old_stock_k2' ||
    op === 'returnk1' ||
    op === 'return_k1' ||
    op === 'testprintk2' ||
    op === 'test_print_k2' ||
    op === 'testk2' ||
    op === 'testprint' ||
    op === 'damaged' ||
    op === 'damagedk2' ||
    op === 'damaged_k2' ||
    op === 'invalid' ||
    op.includes('openk1') ||
    op.includes('oldstockk2') ||
    op.includes('returnk1') ||
    op.includes('testprint') ||
    op.includes('សាកក២') ||
    op.includes('សាក ក២') ||
    op.includes('បោះពុម្ពសាកល្បង') ||
    op.includes('ខូចក២') ||
    op.includes('ខូច ក២') ||
    op.includes('ខូចរបស់ក២') ||
    op.includes('ខូចរបស់ ក២') ||
    op.includes('បង្វិលក១') ||
    op.includes('បង្វិលទៅក១') ||
    op.includes('បង្វិលទៅ ក១') ||
    op.includes('បើកពីក១') ||
    op.includes('បើកពី ក១') ||
    op.includes('បើកពីk1') ||
    op.includes('បើកពី k1') ||
    sf.includes('សាកក២') ||
    sf.includes('ខូចក២') ||
    sf.includes('បង្វិលក១') ||
    sf.includes('បង្វិលទៅក១') ||
    sf.includes('ស្តុកចាស់ ក២') ||
    sf.includes('ស្តុកចាស់ក២')
  ) {
    return true;
  }

  return false;
};

// Strictly checks if a record is "បង្វិលក២" (ទិដ្ឋាការបង្វិលពីក្រុមមកក២)
export const isStrictTeamReturnK2Record = (r: StockRecord | any): boolean => {
  if (!r || isExcludedFromTeamStockFormula(r) || isCeaRecord(r) || isOldStockTeamRecord(r) || isTransferTeamRecord(r)) {
    return false;
  }
  const op = String(r.operationType || '').trim().toLowerCase();
  const sf = String(r.sourceFrom || '').trim().toLowerCase();
  return (
    op === 'returnteam' ||
    op === 'return_team' ||
    op === 'return_from_team' ||
    op === 'returnoffice' ||
    op.includes('បង្វិលពីក្រុម') ||
    op.includes('បង្វិលទៅក២') ||
    op.includes('បង្វិលក២') ||
    sf.includes('បង្វិលពីក្រុម') ||
    sf.includes('បង្វិលទៅក២') ||
    sf.includes('បង្វិលក២')
  );
};

// Strictly checks if a record is "ខូច/បាត់របស់ក្រុម" (damagedTeam / missingTeam)
export const isStrictTeamDamagedOrMissingRecord = (r: StockRecord | any): boolean => {
  if (!r || isExcludedFromTeamStockFormula(r) || isCeaRecord(r) || isOldStockTeamRecord(r) || isTransferTeamRecord(r)) {
    return false;
  }
  const op = String(r.operationType || '').trim().toLowerCase();
  const sf = String(r.sourceFrom || '').trim().toLowerCase();
  return (
    op === 'damagedteam' ||
    op === 'damaged_team' ||
    op === 'teamdamaged' ||
    op === 'missingteam' ||
    op === 'missing_team' ||
    op === 'voidteam' ||
    op.includes('ខូចក្រុម') ||
    op.includes('ខូចរបស់ក្រុម') ||
    op.includes('ខ្វះក្រុម') ||
    op.includes('ខ្វះរបស់ក្រុម') ||
    op.includes('បាត់ក្រុម') ||
    op.includes('បាត់របស់ក្រុម') ||
    sf.includes('ខូចក្រុម') ||
    sf.includes('ខូចរបស់ក្រុម') ||
    sf.includes('ខ្វះក្រុម') ||
    sf.includes('បាត់ក្រុម')
  );
};

// Strictly checks if a record is "បានប្រើប្រាស់" (useTeam - ការប្រើប្រាស់តាមក្រុម)
export const isStrictTeamUseRecord = (r: StockRecord | any): boolean => {
  if (
    !r ||
    isExcludedFromTeamStockFormula(r) ||
    isCeaRecord(r) ||
    isOldStockTeamRecord(r) ||
    isTransferTeamRecord(r) ||
    isStrictTeamReturnK2Record(r) ||
    isStrictTeamDamagedOrMissingRecord(r)
  ) {
    return false;
  }
  const op = String(r.operationType || '').trim().toLowerCase();
  const sf = String(r.sourceFrom || '').trim().toLowerCase();
  return (
    op === 'useteam' ||
    op === 'use_team' ||
    op === 'used' ||
    op === 'use' ||
    op === 'ប្រើប្រាស់តាមក្រុម' ||
    op === 'ការប្រើប្រាស់តាមក្រុម' ||
    ((op.includes('ប្រើប្រាស់') || sf === 'ប្រើប្រាស់តាមក្រុម' || sf === 'ការប្រើប្រាស់តាមក្រុម') &&
      !op.includes('ផ្ទេរ') &&
      !sf.includes('ផ្ទេរ') &&
      !op.includes('បើក') &&
      !op.includes('បង្វិល') &&
      !op.includes('ខូច') &&
      !op.includes('ខ្វះ') &&
      !op.includes('បាត់'))
  );
};

// Helper to check if a record is strictly an Issue to Team (ប្រតិបត្តិការ: បើកផ្តល់តាមក្រុម / បើកពីក២) in Sticker Visa data
export const isIssueTeamRecord = (r: StockRecord): boolean => {
  if (r.stockType && r.stockType !== 'sticker') return false;
  if (isCeaRecord(r)) return false;
  if (isOldStockTeamRecord(r)) return false;
  if (isExcludedFromTeamStockFormula(r)) return false;
  if (isTransferTeamRecord(r)) return false;
  if (isStrictTeamReturnK2Record(r) || isStrictTeamDamagedOrMissingRecord(r) || isStrictTeamUseRecord(r)) return false;

  const op = (r.operationType || '').trim().toLowerCase();

  // Strictly issuance to teams (header ប្រតិបត្តិការ: បើកផ្តល់តាមក្រុម / បើកពីក២)
  return (
    op === 'issueteam' ||
    op === 'issue_team' ||
    op === 'openteam' ||
    op === 'issued' ||
    op.includes('បើកផ្តល់តាមក្រុម') ||
    op.includes('បើកផ្តល់ទៅក្រុម') ||
    op.includes('បើកផ្តល់') ||
    op.includes('បើកពីក២') ||
    op.includes('ទទួលពីការិយាល័យ') ||
    op.includes('ទទួលពីក២')
  );
};

export interface TeamStockStatus {
  opening: Record<string, number>;
  issued: Record<string, number>;
  usedBefore: Record<string, number>;
  usedToday: Record<string, number>;
  transferred: Record<string, number>;
  damaged: Record<string, number>;
  remainingBeforeToday: Record<string, number>;
  remainingAfterToday: Record<string, number>;
  totals: {
    opening: number;
    issued: number;
    usedBefore: number;
    usedToday: number;
    transferred: number;
    damaged: number;
    remainingBeforeToday: number;
    remainingAfterToday: number;
  };
}

// Caching for performance to avoid expensive JSON.parse in high-frequency rendering loops
let lastSavedDaily1Str = '';
let lastSavedDaily1Parsed: any[] = [];

let lastSavedDaily5Str = '';
let lastSavedDaily5Parsed: any[] = [];

let lastSavedExStr = '';
let lastSavedExParsed: string[] = [];

function getCachedDailyRecords(): any[] {
  if (typeof window === 'undefined' || !window.localStorage) return [];
  const savedDaily1 = localStorage.getItem('app_daily_team_records_v1') || '';
  const savedDaily5 = localStorage.getItem('app_daily_team_operations_v5') || '';
  const allDailyRecords: any[] = [];

  if (savedDaily1) {
    if (savedDaily1 === lastSavedDaily1Str) {
      allDailyRecords.push(...lastSavedDaily1Parsed);
    } else {
      try {
        const p1 = JSON.parse(savedDaily1);
        if (Array.isArray(p1)) {
          lastSavedDaily1Str = savedDaily1;
          lastSavedDaily1Parsed = p1;
          allDailyRecords.push(...p1);
        }
      } catch {}
    }
  }

  if (savedDaily5) {
    if (savedDaily5 === lastSavedDaily5Str) {
      allDailyRecords.push(...lastSavedDaily5Parsed);
    } else {
      try {
        const p5 = JSON.parse(savedDaily5);
        if (Array.isArray(p5)) {
          lastSavedDaily5Str = savedDaily5;
          lastSavedDaily5Parsed = p5;
          allDailyRecords.push(...p5);
        }
      } catch {}
    }
  }

  return allDailyRecords;
}

function getCachedExcludedIds(): string[] {
  if (typeof window === 'undefined' || !window.localStorage) return [];
  const savedEx = localStorage.getItem('app_daily_team_excluded_ids') || '';
  if (!savedEx) return [];
  if (savedEx === lastSavedExStr) {
    return lastSavedExParsed;
  }
  try {
    const parsedEx = JSON.parse(savedEx);
    if (Array.isArray(parsedEx)) {
      lastSavedExStr = savedEx;
      lastSavedExParsed = parsedEx;
      return parsedEx;
    }
  } catch {}
  return [];
}

let lastStockCalcKey = '';
let lastStockCalcResult: any = null;

/**
 * Calculates remaining visa sticker stock for all teams or a specific team up to targetDate.
 */
export function calculateAllTeamsStickerStockAtDate(
  stockRecords: StockRecord[],
  targetDate: string,
  teamsList: string[] = Array.from(OFFICIAL_29_TEAMS),
  customExcludedIds?: string[]
): {
  openingMatrix: Record<string, Record<string, number>>;
  issuedMatrix: Record<string, Record<string, number>>;
  usedBeforeMatrix: Record<string, Record<string, number>>;
  usedTodayMatrix: Record<string, Record<string, number>>;
  transferredMatrix: Record<string, Record<string, number>>;
  damagedMatrix: Record<string, Record<string, number>>;
  remainingBeforeMatrix: Record<string, Record<string, number>>;
  remainingAfterMatrix: Record<string, Record<string, number>>;
  totalRemainingAfter: Record<string, number>;
} {
  const cacheKey = `${stockRecords?.length || 0}_${targetDate}_${teamsList.length}_${teamsList[0] || ''}_${customExcludedIds?.length || 0}`;
  if (cacheKey === lastStockCalcKey && lastStockCalcResult) {
    return lastStockCalcResult;
  }

  const normTargetDate = normalizeDateToISO(targetDate);

  // Load excluded IDs from checklist
  const excludedSet = new Set<string>(customExcludedIds || []);
  const cachedEx = getCachedExcludedIds();
  cachedEx.forEach((id) => excludedSet.add(id));

  // Initialize matrices
  const openMat: Record<string, Record<string, number>> = {};
  const issMat: Record<string, Record<string, number>> = {};
  const useBeforeMat: Record<string, Record<string, number>> = {};
  const useTodayMat: Record<string, Record<string, number>> = {};
  const transMat: Record<string, Record<string, number>> = {};
  const damMat: Record<string, Record<string, number>> = {};
  const remBeforeMat: Record<string, Record<string, number>> = {};
  const remAfterMat: Record<string, Record<string, number>> = {};
  const totalRemAfter: Record<string, number> = {};

  // 1. Map recorded old stock
  const recordedOldStockMap: Record<string, Record<string, number>> = {};
  const teamsWithRecordedOldStock = new Set<string>();

  (stockRecords || []).forEach((rec) => {
    if (excludedSet.has(rec.id)) return;
    if (!isOldStockTeamRecord(rec)) return;
    const rawTeam = resolveRecordTeamName(rec);
    const targetTeam = matchTeamInList(rawTeam, teamsList);
    const normTeam = normalizeTeamName(rawTeam);
    if (!targetTeam && !normTeam && !rawTeam) return;

    const vt = normalizeVisaType(rec.visaType);
    if (!VISA_TYPES.includes(vt as any)) return;

    const qty = parseInt(String(rec.quantityBundles ?? rec.totalSheets ?? (rec as any).quantity ?? 0), 10) || 0;
    const matched = targetTeam || normTeam || rawTeam;
    teamsWithRecordedOldStock.add(matched);
    if (normTeam) teamsWithRecordedOldStock.add(normTeam);

    if (targetTeam) {
      if (!recordedOldStockMap[targetTeam]) recordedOldStockMap[targetTeam] = {};
      recordedOldStockMap[targetTeam][vt] = (recordedOldStockMap[targetTeam][vt] || 0) + qty;
    }
    if (normTeam && normTeam !== targetTeam) {
      if (!recordedOldStockMap[normTeam]) recordedOldStockMap[normTeam] = {};
      recordedOldStockMap[normTeam][vt] = (recordedOldStockMap[normTeam][vt] || 0) + qty;
    }
  });

  // Setup initial baselines
  teamsList.forEach((team) => {
    const normTeam = normalizeTeamName(team);
    openMat[team] = {};
    issMat[team] = {};
    useBeforeMat[team] = {};
    useTodayMat[team] = {};
    transMat[team] = {};
    damMat[team] = {};
    remBeforeMat[team] = {};
    remAfterMat[team] = {};

    VISA_TYPES.forEach((vt) => {
      let initialVal = 0;
      if (recordedOldStockMap[team]?.[vt] !== undefined) {
        initialVal = recordedOldStockMap[team][vt];
      } else if (recordedOldStockMap[normTeam]?.[vt] !== undefined) {
        initialVal = recordedOldStockMap[normTeam][vt];
      } else {
        initialVal = DEFAULT_OPENING_MATRIX[team]?.[vt] ?? DEFAULT_OPENING_MATRIX[normTeam]?.[vt] ?? 0;
      }

      openMat[team][vt] = initialVal;
      issMat[team][vt] = 0;
      useBeforeMat[team][vt] = 0;
      useTodayMat[team][vt] = 0;
      transMat[team][vt] = 0;
      damMat[team][vt] = 0;
      remBeforeMat[team][vt] = initialVal;
      remAfterMat[team][vt] = initialVal;
    });
  });

  // Helper to ensure team matrices exist for any team discovered in stockRecords or daily operations
  const ensureTeamMat = (tName: string) => {
    if (!tName || openMat[tName]) return;
    const nTeam = normalizeTeamName(tName);
    openMat[tName] = {};
    issMat[tName] = {};
    useBeforeMat[tName] = {};
    useTodayMat[tName] = {};
    transMat[tName] = {};
    damMat[tName] = {};
    remBeforeMat[tName] = {};
    remAfterMat[tName] = {};

    VISA_TYPES.forEach((vt) => {
      let initialVal = 0;
      if (recordedOldStockMap[tName]?.[vt] !== undefined) {
        initialVal = recordedOldStockMap[tName][vt];
      } else if (recordedOldStockMap[nTeam]?.[vt] !== undefined) {
        initialVal = recordedOldStockMap[nTeam][vt];
      } else {
        initialVal = DEFAULT_OPENING_MATRIX[tName]?.[vt] ?? DEFAULT_OPENING_MATRIX[nTeam]?.[vt] ?? 0;
      }
      openMat[tName][vt] = initialVal;
      issMat[tName][vt] = 0;
      useBeforeMat[tName][vt] = 0;
      useTodayMat[tName][vt] = 0;
      transMat[tName][vt] = 0;
      damMat[tName][vt] = 0;
      remBeforeMat[tName][vt] = initialVal;
      remAfterMat[tName][vt] = initialVal;
    });
  };

  // 2. Process stock records
  (stockRecords || []).forEach((rec) => {
    if (excludedSet.has(rec.id)) return;
    if (rec.stockType && rec.stockType !== 'sticker') return;
    if (isCeaRecord(rec)) return;
    if (isOldStockTeamRecord(rec)) return;
    if (isExcludedFromTeamStockFormula(rec)) return;

    const rawTeam = resolveRecordTeamName(rec);
    const targetTeam = matchTeamInList(rawTeam, teamsList);
    const normTeam = normalizeTeamName(rawTeam);
    const effectiveTeam = teamsList.find((t) => t === targetTeam || t === rawTeam || normalizeTeamName(t) === normTeam) || targetTeam || normTeam || rawTeam;
    if (!effectiveTeam) return;

    ensureTeamMat(effectiveTeam);
    if (!openMat[effectiveTeam]) return;

    const vt = normalizeVisaType(rec.visaType);
    if (!VISA_TYPES.includes(vt as any)) return;

    const qty = Number(rec.totalSheets || rec.quantityBundles || (rec as any).quantity || 0);
    const recDate = normalizeDateToISO(rec.date || '');

    // Strict cutoff: Baseline is fixed as of 30-Nov-2018.
    // Do NOT count any operations (issued from K2, used, transferred, damaged) that occurred on or before 2018-11-30.
    // Operations (issuance and usage) must strictly be on or after 2018-12-01.
    if (!recDate || recDate <= '2018-11-30' || recDate < '2018-12-01') return;

    const isTransfer = isTransferTeamRecord(rec);
    const isDamagedOrReturn = isStrictTeamReturnK2Record(rec) || isStrictTeamDamagedOrMissingRecord(rec);
    const isUse = isStrictTeamUseRecord(rec);
    const isIssue = isIssueTeamRecord(rec);

    if (!isTransfer && !isDamagedOrReturn && !isUse && !isIssue) return;

    // If transfer:
    // 1. Sender transfers out -> transMat
    // 2. Recipient receives transfer -> issMat (បើកផ្តល់ពីក២: ទទួលពីការិយាល័យ ឬផ្ទេរពីក្រុម)
    if (isTransfer) {
      // 1. Sender
      if (effectiveTeam && transMat[effectiveTeam]) {
        if (normTargetDate && recDate) {
          if (recDate < normTargetDate) {
            transMat[effectiveTeam][vt] = (transMat[effectiveTeam][vt] || 0) + qty;
          } else if (recDate === normTargetDate) {
            transMat[effectiveTeam][vt] = (transMat[effectiveTeam][vt] || 0) + qty;
          }
        } else {
          transMat[effectiveTeam][vt] = (transMat[effectiveTeam][vt] || 0) + qty;
        }
      }

      // 2. Recipient
      const rawRecip = (
        rec.recipientTeamName ||
        (rec as any).destinationTo ||
        (rec as any).targetTeam ||
        (rec as any).recipientTeamId ||
        ''
      ).trim();
      const targetRecip = matchTeamInList(rawRecip, teamsList);
      const normRecip = normalizeTeamName(rawRecip);
      const effectiveRecip =
        teamsList.find((t) => t === targetRecip || t === rawRecip || normalizeTeamName(t) === normRecip) ||
        targetRecip ||
        normRecip ||
        rawRecip;

      if (effectiveRecip) {
        ensureTeamMat(effectiveRecip);
        if (issMat[effectiveRecip]) {
          if (normTargetDate && recDate) {
            if (recDate < normTargetDate) {
              issMat[effectiveRecip][vt] = (issMat[effectiveRecip][vt] || 0) + qty;
            } else if (recDate === normTargetDate) {
              issMat[effectiveRecip][vt] = (issMat[effectiveRecip][vt] || 0) + qty;
            }
          } else {
            issMat[effectiveRecip][vt] = (issMat[effectiveRecip][vt] || 0) + qty;
          }
        }
      }
      return;
    }

    // If targetDate is set, process movements up to targetDate
    if (normTargetDate && recDate) {
      if (recDate < normTargetDate) {
        // Occurred BEFORE targetDate
        if (isIssue) {
          issMat[effectiveTeam][vt] = (issMat[effectiveTeam][vt] || 0) + qty;
        } else if (isUse) {
          useBeforeMat[effectiveTeam][vt] = (useBeforeMat[effectiveTeam][vt] || 0) + qty;
        } else if (isDamagedOrReturn) {
          damMat[effectiveTeam][vt] = (damMat[effectiveTeam][vt] || 0) + qty;
        }
      } else if (recDate === normTargetDate) {
        // Occurred ON targetDate (today)
        if (isIssue) {
          issMat[effectiveTeam][vt] = (issMat[effectiveTeam][vt] || 0) + qty;
        } else if (isUse) {
          useTodayMat[effectiveTeam][vt] = (useTodayMat[effectiveTeam][vt] || 0) + qty;
        } else if (isDamagedOrReturn) {
          damMat[effectiveTeam][vt] = (damMat[effectiveTeam][vt] || 0) + qty;
        }
      }
    } else {
      // If no targetDate, process all
      if (isIssue) {
        issMat[effectiveTeam][vt] = (issMat[effectiveTeam][vt] || 0) + qty;
      } else if (isUse) {
        useTodayMat[effectiveTeam][vt] = (useTodayMat[effectiveTeam][vt] || 0) + qty;
      } else if (isDamagedOrReturn) {
        damMat[effectiveTeam][vt] = (damMat[effectiveTeam][vt] || 0) + qty;
      }
    }
  });

  // 3. Process local daily team operations (if not duplicate in stockRecords)
  try {
    const allDailyRecords = getCachedDailyRecords();

    if (allDailyRecords.length > 0) {
      const existingUseKeySet = new Set<string>();
      (stockRecords || []).forEach((sr) => {
        if (isStrictTeamUseRecord(sr)) {
          const dIso = normalizeDateToISO(sr.date || '');
          const tNorm = normalizeTeamName(resolveRecordTeamName(sr) || sr.visaTeamRobokName || '');
          const vNorm = normalizeVisaType(sr.visaType);
          if (dIso && tNorm && vNorm) {
            existingUseKeySet.add(`${dIso}_${tNorm}_${vNorm}`);
          }
        }
      });

      const processedRecordIds = new Set<string>();

      allDailyRecords.forEach((dRec: any) => {
        if (!dRec || !dRec.id || processedRecordIds.has(dRec.id)) return;
        processedRecordIds.add(dRec.id);

        if (excludedSet.has(dRec.id)) return;
        if (dRec.categoryType === 'cEA' || (dRec.categoryType && dRec.categoryType !== 'Sticker')) return;
        const rawTeamName = dRec.teamName || '';
        const normTeam = normalizeTeamName(rawTeamName);
        const targetTeam = teamsList.find((t) => t === rawTeamName || normalizeTeamName(t) === normTeam) || normTeam || rawTeamName;
        if (!targetTeam) return;

        ensureTeamMat(targetTeam);
        if (!openMat[targetTeam]) return;

        const recDate = normalizeDateToISO(dRec.date || '');
        if (!dRec.values || !recDate || recDate <= '2018-11-30') return;

        VISA_TYPES.forEach((vt) => {
          const item = dRec.values[vt];
          let qty = 0;
          if (item) {
            if (Array.isArray(item.entries) && item.entries.length > 0) {
              qty = item.entries.reduce((sum: number, e: any) => sum + (parseInt(e?.quantity || '', 10) || 0), 0);
            } else {
              qty = parseInt(item.quantity || '', 10) || 0;
            }
          }
          if (qty <= 0) return;

          const lookupKey = `${recDate}_${normTeam}_${vt}`;
          if (!existingUseKeySet.has(lookupKey)) {
            if (normTargetDate && recDate < normTargetDate) {
              useBeforeMat[targetTeam][vt] = (useBeforeMat[targetTeam][vt] || 0) + qty;
            } else if (!normTargetDate || recDate === normTargetDate) {
              useTodayMat[targetTeam][vt] = (useTodayMat[targetTeam][vt] || 0) + qty;
            }
          }
        });
      });
    }
  } catch {}

  // 4. Calculate Remaining
  const allTeamKeys = Array.from(new Set([...teamsList, ...Object.keys(openMat)]));
  allTeamKeys.forEach((team) => {
    let teamRemSum = 0;
    if (!remBeforeMat[team]) remBeforeMat[team] = {};
    if (!remAfterMat[team]) remAfterMat[team] = {};

    VISA_TYPES.forEach((vt) => {
      const op = openMat[team]?.[vt] || 0;
      const iss = issMat[team]?.[vt] || 0;
      const usBef = useBeforeMat[team]?.[vt] || 0;
      const usTod = useTodayMat[team]?.[vt] || 0;
      const tr = transMat[team]?.[vt] || 0;
      const dm = damMat[team]?.[vt] || 0;

      const remBefore = op + iss - usBef - tr - dm;
      const remAfter = remBefore - usTod;

      remBeforeMat[team][vt] = remBefore;
      remAfterMat[team][vt] = remAfter;
      teamRemSum += remAfter;
    });
    totalRemAfter[team] = teamRemSum;
  });

  const res = {
    openingMatrix: openMat,
    issuedMatrix: issMat,
    usedBeforeMatrix: useBeforeMat,
    usedTodayMatrix: useTodayMat,
    transferredMatrix: transMat,
    damagedMatrix: damMat,
    remainingBeforeMatrix: remBeforeMat,
    remainingAfterMatrix: remAfterMat,
    totalRemainingAfter: totalRemAfter,
  };

  lastStockCalcKey = cacheKey;
  lastStockCalcResult = res;
  return res;
}
