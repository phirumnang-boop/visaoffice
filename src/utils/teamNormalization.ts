export const OFFICIAL_29_TEAMS = [
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
  'ព្រំដែន ត្រពាំងស្រែ',
  'ព្រំដែន ត្រពាំងក្រៀល',
  'ព្រំដែន ត្រពាំងផ្លុង',
  'ព្រំដែន ភ្នំដិន',
  'ព្រំដែន កោះរកា',
  'ព្រំដែន ព្រៃវល្លិ៍',
  'ព្រំដែន អូរយ៉ាដាវ',
  'ព្រំដែន ក្អមសំណ',
  'ព្រំដែន ភ្នំដី',
  'កំពង់ផែ ឧកញ៉ាម៉ុង',
  'កំពង់ផែ ស្ទឹងហាវ',
  'កំពង់ផែ ព្រះសីហនុ',
  'កំពង់ផែ ភ្នំពេញ',
  'ព្រំដែន ព្រែកចាក',
  'ព្រំដែន ម៉ឺនជ័យ',
  'ព្រំដែន ស្ទឹងបត់',
  'កំពង់ផែ កោះកុង',
  'កំពង់ផែ កំពត',
] as const;

export const VISA_TYPES = ['T', 'T1', 'T2', 'T3', 'E', 'E1', 'E2', 'E3', 'D', 'K', 'A', 'B', 'C'] as const;

/**
 * Month lookup dictionaries for English and Khmer
 */
const MONTH_LOOKUP: Record<string, string> = {
  // English short & long
  jan: '01', january: '01',
  feb: '02', february: '02',
  mar: '03', march: '03',
  apr: '04', april: '04',
  may: '05',
  jun: '06', june: '06',
  jul: '07', july: '07',
  aug: '08', august: '08',
  sep: '09', sept: '09', september: '09',
  oct: '10', october: '10',
  nov: '11', november: '11',
  dec: '12', december: '12',

  // Khmer months
  មករា: '01',
  កុម្ភៈ: '02', កុម្ភះ: '02', កុម្ភ: '02',
  មីនា: '03', មិនា: '03',
  មេសា: '04',
  ឧសភា: '05',
  មិថុនា: '06',
  កក្កដា: '07', កក្កដ: '07',
  សីហា: '08',
  កញ្ញា: '09',
  តុលា: '10',
  វិច្ឆិកា: '11', វិច្ឆិក: '11',
  ធ្នូ: '12',
};

/**
 * Standardize any date string into ISO YYYY-MM-DD format for accurate comparisons
 */
export function normalizeDateToISO(dStr?: string | number | Date | null): string {
  if (dStr === null || dStr === undefined) return '';
  if (dStr instanceof Date && !isNaN(dStr.getTime())) {
    const y = dStr.getFullYear();
    const m = String(dStr.getMonth() + 1).padStart(2, '0');
    const d = String(dStr.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }

  // If number (e.g. Excel timestamp serial ~ 40000 - 60000)
  if (typeof dStr === 'number') {
    if (dStr > 25000 && dStr < 80000) {
      // Excel epoch starts 1899-12-30
      const excelDate = new Date((dStr - 25569) * 86400 * 1000);
      if (!isNaN(excelDate.getTime())) {
        const y = excelDate.getUTCFullYear();
        const m = String(excelDate.getUTCMonth() + 1).padStart(2, '0');
        const d = String(excelDate.getUTCDate()).padStart(2, '0');
        return `${y}-${m}-${d}`;
      }
    }
  }

  const trimmed = String(dStr).trim();
  if (!trimmed) return '';

  // Convert Khmer digits to Arabic digits
  const khmerDigits = ['០', '១', '២', '៣', '៤', '៥', '៦', '៧', '៨', '៩'];
  let cleanStr = trimmed;
  khmerDigits.forEach((kd, idx) => {
    cleanStr = cleanStr.replaceAll(kd, String(idx));
  });
  cleanStr = cleanStr.trim();

  // If contains 'T' or space for ISO timestamp, take date part first
  const dateOnlyPart = cleanStr.split(/[T\s]/)[0];
  if (/^\d{4}-\d{2}-\d{2}$/.test(dateOnlyPart)) {
    return dateOnlyPart;
  }

  // Check matching YYYY-MM-DD or YYYY/MM/DD
  const ymdMatch = dateOnlyPart.match(/^(\d{4})[-/\.](\d{1,2})[-/\.](\d{1,2})$/);
  if (ymdMatch) {
    return `${ymdMatch[1]}-${ymdMatch[2].padStart(2, '0')}-${ymdMatch[3].padStart(2, '0')}`;
  }

  // Check for month names (e.g. "30-Nov-2018", "17-Aug-2026", "17 Aug 2026", "Nov 30, 2018")
  const wordsAndDigits = cleanStr.split(/[\s,\-\.\/]+/).filter(Boolean);
  if (wordsAndDigits.length >= 3) {
    let year = '';
    let month = '';
    let day = '';

    // Find Year (4 digits, or 2 digits)
    const yIdx = wordsAndDigits.findIndex((w) => /^\d{4}$/.test(w));
    if (yIdx !== -1) {
      year = wordsAndDigits[yIdx];
      const otherParts = wordsAndDigits.filter((_, idx) => idx !== yIdx);

      // Check if one of other parts is a month name
      const mNameIdx = otherParts.findIndex((w) => MONTH_LOOKUP[w.toLowerCase()]);
      if (mNameIdx !== -1) {
        month = MONTH_LOOKUP[otherParts[mNameIdx].toLowerCase()];
        const dayPart = otherParts.find((_, idx) => idx !== mNameIdx && /^\d{1,2}$/.test(_));
        if (dayPart) {
          day = dayPart.padStart(2, '0');
        }
      } else {
        // Both are numbers
        const num1 = parseInt(otherParts[0], 10);
        const num2 = parseInt(otherParts[1], 10);
        if (!isNaN(num1) && !isNaN(num2)) {
          if (yIdx === 0) {
            // YYYY-MM-DD
            month = String(num1).padStart(2, '0');
            day = String(num2).padStart(2, '0');
          } else {
            // DD-MM-YYYY or MM-DD-YYYY
            if (num1 > 12 && num2 <= 12) {
              // DD-MM
              day = String(num1).padStart(2, '0');
              month = String(num2).padStart(2, '0');
            } else if (num2 > 12 && num1 <= 12) {
              // MM-DD
              month = String(num1).padStart(2, '0');
              day = String(num2).padStart(2, '0');
            } else {
              // Default to DD-MM-YYYY
              day = String(num1).padStart(2, '0');
              month = String(num2).padStart(2, '0');
            }
          }
        }
      }

      if (year && month && day) {
        return `${year}-${month}-${day}`;
      }
    }
  }

  // Check matching DD-MM-YYYY
  const dmyMatch = dateOnlyPart.match(/^(\d{1,2})[-/\.](\d{1,2})[-/\.](\d{4})$/);
  if (dmyMatch) {
    const dOrM1 = parseInt(dmyMatch[1], 10);
    const dOrM2 = parseInt(dmyMatch[2], 10);
    if (dOrM1 <= 12 && dOrM2 > 12) {
      // MM-DD-YYYY
      return `${dmyMatch[3]}-${String(dOrM1).padStart(2, '0')}-${String(dOrM2).padStart(2, '0')}`;
    }
    return `${dmyMatch[3]}-${String(dOrM2).padStart(2, '0')}-${String(dOrM1).padStart(2, '0')}`;
  }

  // Fallback for general split
  const parts = dateOnlyPart.split(/[\/\-\.]/).filter(Boolean);
  if (parts.length === 3) {
    if (parts[0].length === 4) {
      return `${parts[0]}-${parts[1].padStart(2, '0')}-${parts[2].padStart(2, '0')}`;
    }
    if (parts[2].length === 4) {
      return `${parts[2]}-${parts[1].padStart(2, '0')}-${parts[0].padStart(2, '0')}`;
    }
  }

  return dateOnlyPart;
}

/**
 * Normalizes visa types into standard 13 sticker codes
 */
export function normalizeVisaType(vRaw?: string): string {
  if (!vRaw) return '';
  const clean = vRaw.trim().toUpperCase().replace(/[-\s_]/g, '');

  if (VISA_TYPES.includes(clean as any)) return clean;

  const parenMatch = vRaw.match(/\(([A-Z0-9]+)\)/i);
  if (parenMatch && VISA_TYPES.includes(parenMatch[1].toUpperCase() as any)) {
    return parenMatch[1].toUpperCase();
  }

  if (clean.includes('T1') || clean.includes('T-1')) return 'T1';
  if (clean.includes('T2') || clean.includes('T-2')) return 'T2';
  if (clean.includes('T3') || clean.includes('T-3')) return 'T3';
  if (clean.startsWith('T') && !clean.startsWith('TR')) return 'T';
  if (clean.includes('E1') || clean.includes('E-1')) return 'E1';
  if (clean.includes('E2') || clean.includes('E-2')) return 'E2';
  if (clean.includes('E3') || clean.includes('E-3')) return 'E3';
  if (clean.startsWith('E')) return 'E';
  if (clean.startsWith('D')) return 'D';
  if (clean.startsWith('K')) return 'K';
  if (clean.startsWith('A')) return 'A';
  if (clean.startsWith('B')) return 'B';
  if (clean.startsWith('C')) return 'C';

  return clean;
}

/**
 * Thoroughly normalizes any gate / airport / seaport / team name to match official 29 teams.
 */
export function normalizeTeamName(name?: string): string {
  if (!name) return '';
  const raw = name.trim();
  if (!raw) return '';

  const cleaned = raw.toLowerCase().replace(/\s+/g, ' ');
  const noSpace = cleaned.replace(/\s+/g, '');

  // 1. ក្អមសំណ / រអមសំណ / លើកដែក / ភូមិសាលា (Kaam Samnor / Leuk Daek)
  if (
    noSpace.includes('ក្អមសំណ') ||
    noSpace.includes('ក្អមសំន') ||
    noSpace.includes('រអមសំណ') ||
    noSpace.includes('រអមសំន') ||
    cleaned.includes('ក្អម') ||
    cleaned.includes('រអម') ||
    cleaned.includes('លើកដែក') ||
    cleaned.includes('ភូមិសាលា')
  ) {
    return 'ព្រំដែន ក្អមសំណ';
  }

  // 2. ព្រែកចាក (Prek Chak)
  if (
    noSpace.includes('ព្រែកចាក') ||
    noSpace.includes('ព្រែកចាក់') ||
    (cleaned.includes('ព្រែក') && cleaned.includes('ចាក'))
  ) {
    return 'ព្រំដែន ព្រែកចាក';
  }

  // 3. ព្រះសីហនុ / សីហនុ (Ports vs Airport)
  const hasSihanouk = cleaned.includes('ព្រះសីហនុ') || cleaned.includes('សីហនុ') || cleaned.includes('កំពង់សោម') || noSpace.includes('kps') || noSpace.includes('pas') || cleaned.includes('កងកេង');
  const isPort = cleaned.includes('កំពង់ផែ') || cleaned.includes('ផែ') || cleaned.includes('port') || cleaned.includes('ស្វយ័ត') || noSpace.includes('pas') || cleaned.includes('កំពង់ផែអន្តរជាតិ');
  const isAir = cleaned.includes('អាកាស') || cleaned.includes('ព្រលាន') || cleaned.includes('airport') || cleaned.includes('កងកេង') || cleaned.includes('អាកាសយានដ្ឋាន');

  if (hasSihanouk) {
    if (isPort && !isAir) return 'កំពង់ផែ ព្រះសីហនុ';
    if (isAir && !isPort) return 'អាកាស ព្រះសីហនុ';
    if (isPort) return 'កំពង់ផែ ព្រះសីហនុ';
    if (isAir) return 'អាកាស ព្រះសីហនុ';
    if (cleaned.includes('កំពង់ផែ') || cleaned.includes('ផែ')) return 'កំពង់ផែ ព្រះសីហនុ';
    if (cleaned.includes('អាកាស') || cleaned.includes('យានដ្ឋាន')) return 'អាកាស ព្រះសីហនុ';
    return 'កំពង់ផែ ព្រះសីហនុ';
  }

  // 4. ភ្នំពេញ (Phnom Penh Port vs Techo Airport)
  const hasPhnomPenh = cleaned.includes('ភ្នំពេញ') || cleaned.includes('ពោធិ៍ចិនតុង') || cleaned.includes('តេជោ') || cleaned.includes('តេជា') || noSpace.includes('ppap') || noSpace.includes('tia') || noSpace.includes('pnh');
  if (hasPhnomPenh) {
    if ((cleaned.includes('កំពង់ផែ') || cleaned.includes('ផែ') || cleaned.includes('ppap') || cleaned.includes('ស្វយ័ត')) && !cleaned.includes('អាកាស') && !cleaned.includes('ព្រលាន')) {
      return 'កំពង់ផែ ភ្នំពេញ';
    }
    if (cleaned.includes('អាកាស') || cleaned.includes('ព្រលាន') || cleaned.includes('តេជោ') || cleaned.includes('តេជា') || cleaned.includes('ពោធិ៍ចិនតុង') || noSpace.includes('tia') || noSpace.includes('pnh')) {
      return 'អាកាស តេជោ';
    }
  }

  // 5. ត្រពាំងស្រែ (Trapeang Sre / Snuol)
  if (noSpace.includes('ត្រពាំងស្រែ') || cleaned.includes('ស្នួល')) {
    return 'ព្រំដែន ត្រពាំងស្រែ';
  }

  // 6. ត្រពាំងក្រៀល / ត្រពាំងរូង (Trapeang Kriel)
  if (noSpace.includes('ត្រពាំងក្រៀល') || noSpace.includes('ត្រពាំងរូង') || cleaned.includes('ស្ទឹងត្រែង')) {
    return 'ព្រំដែន ត្រពាំងក្រៀល';
  }

  // 7. ត្រពាំងផ្លុង (Trapeang Phlong / Ponhea Krek)
  if (noSpace.includes('ត្រពាំងផ្លុង') || cleaned.includes('ពញាក្រែក')) {
    return 'ព្រំដែន ត្រពាំងផ្លុង';
  }

  // 8. បាវិត (Bavet / Svay Rieng)
  if (noSpace.includes('បាវិត') || (cleaned.includes('បាវិត') && cleaned.includes('ស្វាយរៀង'))) {
    return 'ព្រំដែន បាវិត';
  }

  // 9. អូរយ៉ាដាវ / អូយ៉ាដាវ (O Yadav / Ratanakiri)
  if (
    noSpace.includes('អូរយ៉ាដាវ') ||
    noSpace.includes('អូយ៉ាដាវ') ||
    noSpace.includes('អូរយ៉ាដៅ') ||
    noSpace.includes('អូយ៉ាដៅ') ||
    noSpace.includes('យ៉ាដាវ') ||
    cleaned.includes('រតនគិរី')
  ) {
    return 'ព្រំដែន អូរយ៉ាដាវ';
  }

  // 10. ប៉ោយប៉ែត (Poipet)
  if (noSpace.includes('ប៉ោយប៉ែត') || noSpace.includes('ប៉ោយប៉ែត')) {
    return 'ព្រំដែន ប៉ោយប៉ែត';
  }

  // 11. ចាំយាម (Cham Yeam / Koh Kong Border)
  if (noSpace.includes('ចាំយាម') || (cleaned.includes('ចាំ') && cleaned.includes('យាម'))) {
    return 'ព្រំដែន ចាំយាម';
  }

  // 12. ដូង (Doung / Kamrieng / Battambang)
  if (noSpace.includes('ដូង') || cleaned.includes('កំរៀង')) {
    return 'ព្រំដែន ដូង';
  }

  // 13. អូរស្មាច់ / អូស្មាច់ (O Smach)
  if (noSpace.includes('អូរស្មាច់') || noSpace.includes('អូស្មាច់')) {
    return 'ព្រំដែន អូរស្មាច់';
  }

  // 14. បន្ទាយចក្រី (Banteay Chakrei)
  if (noSpace.includes('បន្ទាយចក្រី') || (cleaned.includes('បន្ទាយ') && cleaned.includes('ចក្រី'))) {
    return 'ព្រំដែន បន្ទាយចក្រី';
  }

  // 15. ជាំ (Choam)
  if (cleaned.includes('ជាំ') && !cleaned.includes('ចាំ')) {
    return 'ព្រំដែន ជាំ';
  }

  // 16. ភ្នំដិន (Phnom Den / Kirivong / Takeo)
  if (noSpace.includes('ភ្នំដិន') || cleaned.includes('គិរីវង់')) {
    return 'ព្រំដែន ភ្នំដិន';
  }

  // 17. កោះរកា (Koh Roka / Peam Chor)
  if (noSpace.includes('កោះរកា') || cleaned.includes('ពាមជរ')) {
    return 'ព្រំដែន កោះរកា';
  }

  // 18. ព្រៃវល្លិ៍ (Prey Voa)
  if (noSpace.includes('ព្រៃវល្លិ៍') || noSpace.includes('ព្រៃវល្ល') || noSpace.includes('ព្រៃវល្លិ') || cleaned.includes('កំពង់រោទិ៍')) {
    return 'ព្រំដែន ព្រៃវល្លិ៍';
  }

  // 19. ភ្នំដី (Phnom Dei / Sampov Loun)
  if (noSpace.includes('ភ្នំដី') || cleaned.includes('សំពៅលូន')) {
    return 'ព្រំដែន ភ្នំដី';
  }

  // 20. ឧកញ៉ាម៉ុង (Oknha Mong)
  if (noSpace.includes('ឧកញ៉ាម៉ុង') || cleaned.includes('កែវផុស')) {
    return 'កំពង់ផែ ឧកញ៉ាម៉ុង';
  }

  // 21. ស្ទឹងហាវ (Steung Hav)
  if (noSpace.includes('ស្ទឹងហាវ')) {
    return 'កំពង់ផែ ស្ទឹងហាវ';
  }

  // 22. ម៉ឺនជ័យ (Meun Chey / Kamchay Mear)
  if (noSpace.includes('ម៉ឺនជ័យ') || cleaned.includes('កំចាយមារ')) {
    return 'ព្រំដែន ម៉ឺនជ័យ';
  }

  // 23. ស្ទឹងបត់ (Steung Bot)
  if (noSpace.includes('ស្ទឹងបត់')) {
    return 'ព្រំដែន ស្ទឹងបត់';
  }

  // 24. កោះកុង កំពង់ផែ (Koh Kong Port)
  if (
    !noSpace.includes('ចាំយាម') &&
    (noSpace.includes('កោះកុង') || cleaned.includes('កោះកុង'))
  ) {
    return 'កំពង់ផែ កោះកុង';
  }

  // 25. កំពត កំពង់ផែ (Kampot Port)
  if (
    !noSpace.includes('ព្រែកចាក') &&
    (noSpace.includes('កំពត') || cleaned.includes('កំពត'))
  ) {
    return 'កំពង់ផែ កំពត';
  }

  // 26. អាកាស សៀមរាប (Siem Reap Airport)
  if (cleaned.includes('សៀមរាប') || cleaned.includes('អង្គរ') || noSpace.includes('sai') || noSpace.includes('rep')) {
    return 'អាកាស សៀមរាប';
  }

  // 27. ព្រំដែន ព្រំ (Phrom Border / Pailin)
  // Strip all standard prefixes
  const stripped = cleaned
    .replace(/^(ច្រកទ្វារព្រំដែនអន្តរជាតិ|ច្រកទ្វារព្រំដែន|ច្រកទ្វារអន្តរជាតិ|ច្រកទ្វារ|ព្រំដែន|ច្រក|ក្រុម|ប៉ុស្តិ៍|អន្តរជាតិ|\s)+/g, '')
    .trim();

  if (
    stripped === 'ព្រំ' ||
    cleaned === 'ព្រំ' ||
    cleaned.includes('ព្រំដែន ព្រំ') ||
    cleaned.includes('ព្រំដែនព្រំ') ||
    cleaned.includes('ច្រកព្រំ') ||
    cleaned.includes('ច្រកទ្វារព្រំ') ||
    cleaned.includes('ក្រុមព្រំ') ||
    cleaned.includes('ក្រុម ព្រំ') ||
    cleaned.includes('ផ្សារព្រំ') ||
    cleaned.includes('ប៉ៃលិន')
  ) {
    return 'ព្រំដែន ព្រំ';
  }

  // 28. ព្រែកបាក់ (Prek Bak)
  if (noSpace.includes('ព្រែកបាក់') || cleaned.includes('ព្រែកបាក់')) {
    return 'ព្រំដែន ព្រែកបាក់';
  }

  return raw;
}

/**
 * Match a raw team name into the provided teamsList
 */
export function matchTeamInList(rawName: string, teamsList: string[] = Array.from(OFFICIAL_29_TEAMS)): string {
  if (!rawName) return '';
  const exact = teamsList.find((t) => t === rawName);
  if (exact) return exact;

  const norm = normalizeTeamName(rawName);
  const matched = teamsList.find((t) => normalizeTeamName(t) === norm);
  if (matched) return matched;

  const partial = teamsList.find((t) => {
    const tNorm = normalizeTeamName(t);
    return tNorm.includes(norm) || norm.includes(tNorm);
  });
  if (partial) return partial;

  return norm || rawName;
}

export const OFFICIAL_TEAM_FULL_NAMES: Record<string, string> = {
  'អាកាស តេជោ': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអាកាសយានដ្ឋានអន្តរជាតិតេជោ',
  'អាកាស សៀមរាប': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអាកាសយានដ្ឋានអន្តរជាតិសៀមរាបអង្គរ',
  'អាកាស ព្រះសីហនុ': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអាកាសយានដ្ឋានអន្តរជាតិក្រុងព្រះសីហនុ',
  'ព្រំដែន ប៉ោយប៉ែត': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិប៉ោយប៉ែត',
  'ព្រំដែន បាវិត': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិបាវិត',
  'ព្រំដែន ចាំយាម': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិចាំយាម',
  'ព្រំដែន ដូង': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិដូង',
  'ព្រំដែន អូរស្មាច់': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិអូរស្មាច់',
  'ព្រំដែន ព្រំ': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិព្រំ',
  'ព្រំដែន បន្ទាយចក្រី': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិបន្ទាយចក្រី',
  'ព្រំដែន ជាំ': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិជាំ',
  'ព្រំដែន ត្រពាំងផ្លុង': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិត្រពាំងផ្លុង',
  'ព្រំដែន ត្រពាំងស្រែ': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិត្រពាំងស្រែ',
  'ព្រំដែន ត្រពាំងក្រៀល': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិត្រពាំងក្រៀល',
  'ព្រំដែន ភ្នំដិន': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិភ្នំដិន',
  'ព្រំដែន កោះរកា': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិកោះរកា',
  'ព្រំដែន ព្រៃវល្លិ៍': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិព្រៃវល្លិ៍',
  'ព្រំដែន អូរយ៉ាដាវ': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិអូរយ៉ាដាវ',
  'ព្រំដែន ក្អមសំណ': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិក្អមសំណ',
  'ព្រំដែន ភ្នំដី': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិភ្នំដី',
  'កំពង់ផែ ឧកញ៉ាម៉ុង': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារកំពង់ផែអន្តរជាតិឧកញ៉ាម៉ុង',
  'កំពង់ផែ ស្ទឹងហាវ': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារកំពង់ផែអន្តរជាតិស្ទឹងហាវ',
  'កំពង់ផែ ព្រះសីហនុ': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារកំពង់ផែអន្តរជាតិក្រុងព្រះសីហនុ',
  'កំពង់ផែ ភ្នំពេញ': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារកំពង់ផែអន្តរជាតិភ្នំពេញ',
  'ព្រំដែន ព្រែកចាក': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិព្រែកចាក',
  'ព្រំដែន ម៉ឺនជ័យ': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិម៉ឺនជ័យ',
  'ព្រំដែន ស្ទឹងបត់': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិស្ទឹងបត់',
  'កំពង់ផែ កោះកុង': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារកំពង់ផែអន្តរជាតិកោះកុង',
  'កំពង់ផែ កំពត': 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារកំពង់ផែអន្តរជាតិកំពត',
};

/**
 * Format a short team name into its full official administrative Khmer title
 */
export function formatReportTeamName(teamName?: string): string {
  if (!teamName) return 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារ';
  const t = teamName.trim();
  if (t === 'សរុបបណ្តាក្រុម' || t === 'សរុបបណ្តាក្រុមផ្តល់ទិដ្ឋាការ') {
    return 'សរុបបណ្តាក្រុមផ្តល់ទិដ្ឋាការ';
  }
  const norm = normalizeTeamName(t);

  if (OFFICIAL_TEAM_FULL_NAMES[norm]) {
    return OFFICIAL_TEAM_FULL_NAMES[norm];
  }
  if (OFFICIAL_TEAM_FULL_NAMES[t]) {
    return OFFICIAL_TEAM_FULL_NAMES[t];
  }

  if (t.startsWith('ក្រុមផ្តល់ទិដ្ឋាការ')) {
    return t;
  }

  if (norm.startsWith('អាកាស') || t.startsWith('អាកាស')) {
    const airName = (norm.startsWith('អាកាស') ? norm : t).replace(/^អាកាស\s*/, '').trim();
    if (airName.includes('តេជោ')) return 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអាកាសយានដ្ឋានអន្តរជាតិតេជោ';
    if (airName.includes('សៀមរាប')) return 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអាកាសយានដ្ឋានអន្តរជាតិសៀមរាបអង្គរ';
    if (airName.includes('ព្រះសីហនុ')) return 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអាកាសយានដ្ឋានអន្តរជាតិក្រុងព្រះសីហនុ';
    return `ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអាកាសយានដ្ឋានអន្តរជាតិ${airName}`;
  }

  if (norm.startsWith('កំពង់ផែ') || t.startsWith('កំពង់ផែ')) {
    const portName = (norm.startsWith('កំពង់ផែ') ? norm : t).replace(/^កំពង់ផែ\s*/, '').trim();
    return `ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារកំពង់ផែអន្តរជាតិ${portName}`;
  }

  if (norm.startsWith('ព្រំដែន') || t.startsWith('ព្រំដែន')) {
    const borderName = (norm.startsWith('ព្រំដែន') ? norm : t).replace(/^ព្រំដែន\s*/, '').trim();
    return `ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារព្រំដែនអន្តរជាតិ${borderName}`;
  }

  return `ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារ ${t}`;
}
