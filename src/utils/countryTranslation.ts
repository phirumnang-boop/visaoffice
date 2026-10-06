export const COUNTRY_MAPPING: Record<string, string> = {
  // Common nationalities / countries in Khmer to English
  'ស៊ូម៉ាលី': 'Somalia',
  'សូម៉ាលី': 'Somalia',
  'ចិន': 'China',
  'វៀតណាម': 'Vietnam',
  'ថៃ': 'Thailand',
  'អាមេរិក': 'United States',
  'បារាំង': 'France',
  'ជប៉ុន': 'Japan',
  'កូរ៉េខាងត្បូង': 'South Korea',
  'កូរ៉េ': 'Korea',
  'ឥណ្ឌា': 'India',
  'ម៉ាឡេស៊ី': 'Malaysia',
  'សិង្ហបុរី': 'Singapore',
  'ឡាវ': 'Laos',
  'មីយ៉ាន់ម៉ា': 'Myanmar',
  'ភូមា': 'Myanmar',
  'ហ្វីលីពីន': 'Philippines',
  'ឥណ្ឌូនេស៊ី': 'Indonesia',
  'អង់គ្លេស': 'United Kingdom',
  'ចក្រភពអង់គ្លេស': 'United Kingdom',
  'អូស្ត្រាលី': 'Australia',
  'កាណាដា': 'Canada',
  'អាល្លឺម៉ង់': 'Germany',
  'រុស្ស៊ី': 'Russia',
  'បង់ក្លាដែស': 'Bangladesh',
  'ប៉ាគីស្ថាន': 'Pakistan',
  'នេប៉ាល់': 'Nepal',
  'ស្រីលង្កា': 'Sri Lanka',
  'អ៊ីរ៉ង់': 'Iran',
  'អ៊ីរ៉ាក់': 'Iraq',
  'នីហ្សេរីយ៉ា': 'Nigeria',
  'អេហ្ស៊ីប': 'Egypt',
  'អាហ្វ្រិកខាងត្បូង': 'South Africa',
  'ប្រេស៊ីល': 'Brazil',
  'ម៉ិកស៊ិក': 'Mexico',
  'អាហ្សង់ទីន': 'Argentina',
  'ហ្វីហ្ស៊ី': 'Fiji',
  'អ៊ុយក្រែន': 'Ukraine',
  'ប៉ូឡូញ': 'Poland',
  'អេស្ប៉ាញ': 'Spain',
  'អ៊ីតាលី': 'Italy',
  'ហូឡង់': 'Netherlands',
  'ស្វីស': 'Switzerland',
  'ស៊ុយអែត': 'Sweden',
  'ន័រវែស': 'Norway',
  'ដាណermark': 'Denmark',
  'ហ្វាំងឡង់': 'Finland',
  'បែលហ្ស៊ិក': 'Belgium',
  'អូទ្រីស': 'Austria',
  'ក្រិក': 'Greece',
  'ព័រទុយហ្គាល់': 'Portugal',
  'តួកគី': 'Turkey',
  'អារ៉ាប៊ីសាអូឌីត': 'Saudi Arabia',
  'អេមីរ៉ាតអារ៉ាប់រួម': 'United Arab Emirates',
  'កាតា': 'Qatar',
  'គុយវ៉ែត': 'Kuwait',
  'អូម៉ង់': 'Oman',
  'បារ៉ែន': 'Bahrain',
  'ហ្ស៊កដានី': 'Jordan',
  'លីបង់': 'Lebanon',
  'ស៊ីរី': 'Syria',
  'យេម៉ែន': 'Yemen',
  'អ៊ូសបេគីស្ថាន': 'Uzbekistan',
  'កាហ្សាក់ស្ថាន': 'Kazakhstan',
};

export const translateCountryName = (name?: string, toEnglish: boolean = false): string => {
  if (!name || !name.trim()) return '';
  const trimmed = name.trim();
  if (!toEnglish) return trimmed;

  // Check exact match
  if (COUNTRY_MAPPING[trimmed]) {
    return COUNTRY_MAPPING[trimmed];
  }

  // Check case-insensitive or partial match
  for (const [kh, en] of Object.entries(COUNTRY_MAPPING)) {
    if (trimmed.includes(kh)) {
      return trimmed.replace(kh, en);
    }
  }

  return trimmed;
};
