import momentkh from '@thyrith/momentkh';

export const KHMER_MONTHS = [
  'មករា', 'កុម្ភៈ', 'មីនា', 'មេសា', 'ឧសភា', 'មិថុនា',
  'កក្កដា', 'សីហា', 'កញ្ញា', 'តុលា', 'វិច្ឆិកា', 'ធ្នូ'
];

export const toKhmerNum = (num: number | string): string => {
  if (num === null || num === undefined || num === '') return '';
  const khmerDigits = ['០', '១', '២', '៣', '៤', '៥', '៦', '៧', '៨', '៩'];
  return String(num).replace(/[0-9]/g, (digit) => khmerDigits[parseInt(digit, 10)]);
};

export const toKhmerDigits = toKhmerNum;

export function formatCurrencyUSD(amount: number | string): string {
  const num = typeof amount === 'number' ? amount : parseFloat(amount);
  if (isNaN(num) || num === null || num === undefined) return '$0.00';
  return `$${num.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 })}`;
}

export function parseDateInput(dateInput: Date | string): Date {
  if (!dateInput) return new Date();
  if (dateInput instanceof Date) return dateInput;
  if (typeof dateInput === 'string') {
    let cleanStr = dateInput.trim();
    const khmerDigits = ['០', '១', '២', '៣', '៤', '៥', '៦', '៧', '៨', '៩'];
    khmerDigits.forEach((kd, idx) => {
      cleanStr = cleanStr.replaceAll(kd, String(idx));
    });

    if (cleanStr.includes('T')) {
      const parsed = new Date(cleanStr);
      if (!isNaN(parsed.getTime())) return parsed;
    }
    const parts = cleanStr.split(/[T\s]/)[0].split(/[-/\.]/).filter(Boolean);
    if (parts.length === 3) {
      if (parts[0].length === 4) {
        return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
      }
      if (parts[2].length === 4) {
        return new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
      }
    }
    const standardDate = new Date(cleanStr);
    if (!isNaN(standardDate.getTime())) return standardDate;
  }
  return new Date(dateInput);
}

/**
 * Calculates traditional Khmer lunar date from Gregorian date
 * Example output: "ថ្ងៃសុក្រ ៩រោច ខែទុតិយាសាឍ ឆ្នាំមមី អដ្ឋស័ក ព.ស. ២៥៧០"
 */
export function getKhmerLunarDate(dateInput: Date | string): string {
  const d = parseDateInput(dateInput);
  if (isNaN(d.getTime())) return '';

  try {
    const res = momentkh.fromDate(d);
    const k = res.khmer;
    const dayStr = toKhmerNum(k.day);
    const beYearStr = toKhmerNum(k.beYear);
    return `ថ្ងៃ${k.dayOfWeekName} ${dayStr}${k.moonPhaseName} ខែ${k.monthName} ឆ្នាំ${k.animalYearName} ${k.sakName} ព.ស${beYearStr}`;
  } catch (err) {
    console.error('Error calculating Khmer lunar date:', err);
    return '';
  }
}

/**
 * Returns Khmer solar date parts (day in Khmer digits, month name in Khmer, year in Khmer digits)
 */
export function getKhmerSolarParts(dateInput: Date | string) {
  const d = parseDateInput(dateInput);
  const day = isNaN(d.getDate()) ? 1 : d.getDate();
  const monthIndex = isNaN(d.getMonth()) ? 0 : d.getMonth();
  const year = isNaN(d.getFullYear()) ? 2026 : d.getFullYear();

  const khmerDay = toKhmerNum(String(day).padStart(2, '0'));
  const khmerMonth = KHMER_MONTHS[monthIndex] || 'មករា';
  const khmerYear = toKhmerNum(year);

  return {
    day,
    month: khmerMonth,
    monthIndex: monthIndex + 1,
    year,
    khmerDay,
    khmerMonth,
    khmerYear,
  };
}

export function formatKhmerDate(dateInput: Date | string): string {
  const parts = getKhmerSolarParts(dateInput);
  return `ថ្ងៃទី${parts.khmerDay} ខែ${parts.khmerMonth} ឆ្នាំ${parts.khmerYear}`;
}

export function formatPhoneNumber(input: string): string {
  if (!input) return '';
  const digits = input.replace(/\D/g, '');
  if (!digits) return '';

  if (digits.length <= 3) {
    return digits;
  }
  if (digits.length <= 6) {
    return `${digits.slice(0, 3)} ${digits.slice(3)}`;
  }
  if (digits.length <= 9) {
    return `${digits.slice(0, 3)} ${digits.slice(3, 6)} ${digits.slice(6)}`;
  }
  return `${digits.slice(0, 4)} ${digits.slice(4, 7)} ${digits.slice(7, 10)}${digits.length > 10 ? ' ' + digits.slice(10) : ''}`;
}
