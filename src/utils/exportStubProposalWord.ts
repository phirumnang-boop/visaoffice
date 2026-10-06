import { saveAs } from 'file-saver';

export interface ExportRobokWordOptions {
  fileName: string;
  reportTitle: string;
  customMargins: { top: number; bottom: number; left: number; right: number };
  activeTabStop?: number;
  customTacteingImage?: string | null;
  // Metadata
  ministryHierarchy: string[];
  docNumber: string;
  docRecipientRank?: string;
  docRecipientTitle: string;
  docThroughTitle: string;
  docReferenceText?: string;
  reportPeriodLabel: string;
  introSalutation: string;
  startKhmerDate: { day: string; month: string; year: string };
  endKhmerDate: { day: string; month: string; year: string };
  prevKhmerDate: { day: string; month: string; year: string };
  signKhmerDate: { day: string; month: string; year: string };
  signLunarDate?: string;
  signSolarDate?: string;
  // Mode & Team info
  viewMode?: 'office' | 'team';
  teamDisplayName?: string;
  teamPeriodKhmerTitle?: string;
  teamPeriodShortKhmerTitle?: string;
  teamStickerData?: {
    opening: { values: Record<string, number>; total: number };
    newIssued: { values: Record<string, number>; total: number };
    returnedDamagedMissing: { values: Record<string, number>; total: number };
    used: { values: Record<string, number>; total: number };
    ending: { values: Record<string, number>; total: number };
  };
  // Tables Data for Standard/Office Report
  stickerRows?: Array<{
    visaType: string;
    openK2: number;
    openTeams: number;
    k1ToK2: number;
    teamReturnedToK2: number;
    issuedK2ToTeams: number;
    teamsUsed: number;
    damagedTeam: number;
    damagedK2: number;
    testSampleK2: number;
    endingK2: number;
    endingTeams: number;
    grandTotalEnding: number;
  }>;
  stickerTotalRow?: {
    openK2: number;
    openTeams: number;
    k1ToK2: number;
    teamReturnedToK2: number;
    issuedK2ToTeams: number;
    teamsUsed: number;
    damagedTeam: number;
    damagedK2: number;
    testSampleK2: number;
    endingK2: number;
    endingTeams: number;
    grandTotalEnding: number;
  };
  eVisaData?: {
    openBundles: number;
    receivedK1Bundles: number;
    issuedToTeamsBundles: number;
    endingBundles: number;
  };
  stubProposalRows?: Array<{
    visaType: string;
    booklets: number;
    sheets: number;
    remarks?: string;
  }>;
  // Signatures
  leftSignTitle: string;
  leftSignOfficer?: string;
  rightSignTitle: string;
  rightSignOfficer?: string;
  variant?: 'standard' | 'director';
  proposalType?: 'section' | 'office';
  paragraphShiftY?: number;
  isCalculated?: boolean;
}

const DEFAULT_TACTEING_SVG_DATA =
  'data:image/svg+xml;utf8,' +
  encodeURIComponent(
    `<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 240 24" width="240" height="24"><g stroke="#000000" fill="#000000"><path d="M 2 12 L 76 10.5 L 76 13.5 Z" stroke="#000000" stroke-width="0.8" fill="#000000"/><path d="M 76 12 C 80 4.5, 92 3.5, 96 12 C 92 20.5, 80 19.5, 76 12 Z" fill="none" stroke="#000000" stroke-width="2.8"/><circle cx="86" cy="12" r="3" fill="#000000"/><path d="M 91 12 C 89 5.5, 95 2, 93 1 C 90 1, 88 5.5, 90 12 C 88 18.5, 90 23, 93 23 C 95 22, 89 18.5, 91 12 Z" fill="#000000"/><circle cx="120" cy="12" r="3.2" fill="#000000"/><path d="M 120 7.5 C 117 3.5, 117 0.5, 120 -1 C 123 0.5, 123 3.5, 120 7.5 Z" fill="#000000"/><path d="M 120 16.5 C 117 20.5, 117 23.5, 120 25 C 123 23.5, 123 20.5, 120 16.5 Z" fill="#000000"/><path d="M 115.5 12 C 111.5 9, 108.5 9, 107 12 C 108.5 15, 111.5 15, 115.5 12 Z" fill="#000000"/><path d="M 124.5 12 C 128.5 9, 131.5 9, 133 12 C 131.5 15, 128.5 15, 124.5 12 Z" fill="#000000"/><path d="M 117.8 9.8 C 114.5 6.5, 112 4.5, 111 5.8 C 112.5 7.5, 114.5 9.2, 117.8 9.8 Z" fill="#000000"/><path d="M 122.2 9.8 C 125.5 6.5, 128 4.5, 129 5.8 C 127.5 7.5, 125.5 9.2, 122.2 9.8 Z" fill="#000000"/><path d="M 117.8 14.2 C 114.5 17.5, 112 19.5, 111 18.2 C 112.5 16.5, 114.5 14.8, 117.8 14.2 Z" fill="#000000"/><path d="M 122.2 14.2 C 125.5 17.5, 128 19.5, 129 18.2 C 127.5 16.5, 125.5 14.8, 122.2 14.2 Z" fill="#000000"/><path d="M 164 12 C 160 4.5, 148 3.5, 144 12 C 148 20.5, 160 19.5, 164 12 Z" fill="none" stroke="#000000" stroke-width="2.8"/><circle cx="154" cy="12" r="3" fill="#000000"/><path d="M 149 12 C 151 5.5, 145 2, 147 1 C 150 1, 152 5.5, 150 12 C 152 18.5, 150 23, 147 23 C 145 22, 151 18.5, 149 12 Z" fill="#000000"/><path d="M 238 12 L 164 10.5 L 164 13.5 Z" stroke="#000000" stroke-width="0.8" fill="#000000"/></g></svg>`
  );

const VISA_TYPES_LIST = ['T', 'T1', 'T2', 'T3', 'E', 'E1', 'E2', 'E3', 'D', 'K', 'A', 'B', 'C'] as const;

export function exportRobokToWord(options: ExportRobokWordOptions) {
  const {
    fileName,
    customMargins,
    activeTabStop = 2.25,
    customTacteingImage,
    ministryHierarchy,
    docNumber,
    docRecipientRank,
    docRecipientTitle,
    docThroughTitle,
    docReferenceText,
    reportPeriodLabel,
    introSalutation,
    startKhmerDate,
    endKhmerDate,
    prevKhmerDate,
    signKhmerDate,
    signLunarDate,
    signSolarDate,
    viewMode = 'office',
    teamDisplayName,
    teamPeriodKhmerTitle,
    teamPeriodShortKhmerTitle,
    teamStickerData,
    paragraphShiftY = 16,
    stickerRows = [],
    stickerTotalRow = {
      openK2: 0,
      openTeams: 0,
      k1ToK2: 0,
      teamReturnedToK2: 0,
      issuedK2ToTeams: 0,
      teamsUsed: 0,
      damagedTeam: 0,
      damagedK2: 0,
      testSampleK2: 0,
      endingK2: 0,
      endingTeams: 0,
      grandTotalEnding: 0,
    },
    eVisaData = {
      openBundles: 0,
      receivedK1Bundles: 0,
      issuedToTeamsBundles: 0,
      endingBundles: 0,
    },
    leftSignTitle,
    leftSignOfficer,
    rightSignTitle,
    rightSignOfficer,
    variant = 'standard',
    proposalType,
    stubProposalRows,
    isCalculated,
  } = options;

  const ALL_13_VISA_TYPES = ['T', 'T1', 'T2', 'T3', 'E', 'E1', 'E2', 'E3', 'D', 'K', 'A', 'B', 'C'] as const;
  const DEFAULT_STUB_PROPOSAL_ROWS = ALL_13_VISA_TYPES.map((vt) => ({
    visaType: vt,
    booklets: 0,
    sheets: 0,
    remarks: '',
  }));

  const isPendingCalc = isCalculated === false;
  const effectiveStubRows =
    stubProposalRows && stubProposalRows.length > 0 ? stubProposalRows : DEFAULT_STUB_PROPOSAL_ROWS;
  const stubTotalBooklets = effectiveStubRows.reduce((sum, r) => sum + (Number(r.booklets) || 0), 0);
  const stubTotalSheets = effectiveStubRows.reduce((sum, r) => sum + (Number(r.sheets) || 0), 0);

  const stubProposalRowsHtml = effectiveStubRows
    .map(
      (r, idx) => `
    <tr style="mso-height-rule: exactly; height: 0.45cm; page-break-inside: avoid;">
      <td style="border: 0.5pt solid black; text-align: center; font-family: 'Times New Roman', serif; font-size: 12pt; padding: 2pt 2pt; vertical-align: middle;">${idx + 1}</td>
      <td style="border: 0.5pt solid black; text-align: center; font-weight: bold; font-family: 'Times New Roman', serif; font-size: 12pt; padding: 2pt 2pt; vertical-align: middle;">${r.visaType}</td>
      <td style="border: 0.5pt solid black; text-align: center; font-family: 'Times New Roman', serif; font-size: 12pt; padding: 2pt 2pt; vertical-align: middle;">${isPendingCalc || r.booklets === undefined ? '-' : Number(r.booklets).toLocaleString()}</td>
      <td style="border: 0.5pt solid black; text-align: center; font-family: 'Times New Roman', serif; font-size: 12pt; padding: 2pt 2pt; vertical-align: middle;">${isPendingCalc || r.sheets === undefined ? '-' : Number(r.sheets).toLocaleString()}</td>
      <td style="border: 0.5pt solid black; text-align: center; font-family: 'Khmer OS Siemreap', sans-serif; font-size: 12pt; padding: 2pt 2pt; vertical-align: middle;">${r.remarks || ''}</td>
    </tr>`
    )
    .join('');

  const topCm = typeof customMargins.top === 'number' ? customMargins.top : 0.5;
  const bottomCm = customMargins.bottom || 1.2;
  const leftCm = customMargins.left || 1.3;
  const rightCm = customMargins.right || 1.3;
  const tabStopCm = activeTabStop || 2.25;

  const tacteingSrc = customTacteingImage || DEFAULT_TACTEING_SVG_DATA;

  const rowsHtml = stickerRows
    .map(
      (row, idx) => `
    <tr style="mso-height-rule: exactly; height: 0.35cm; page-break-inside: avoid;">
      <td style="border: 0.5pt solid black; text-align: center; font-family: 'Times New Roman', 'Khmer OS Siemreap', serif; font-size: 8.5pt; padding: 0.5pt 1pt; vertical-align: middle;">${idx + 1}</td>
      <td style="border: 0.5pt solid black; text-align: center; font-weight: bold; font-family: 'Times New Roman', 'Khmer OS Siemreap', serif; font-size: 8.5pt; padding: 0.5pt 1pt; vertical-align: middle;">${row.visaType}</td>
      <td style="border: 0.5pt solid black; text-align: right; font-family: 'Times New Roman', 'Khmer OS Siemreap', serif; font-size: 8.5pt; padding: 0.5pt 2pt; vertical-align: middle;">${row.openK2 ? row.openK2.toLocaleString() : '0'}</td>
      <td style="border: 0.5pt solid black; text-align: right; font-family: 'Times New Roman', 'Khmer OS Siemreap', serif; font-size: 8.5pt; padding: 0.5pt 2pt; vertical-align: middle;">${row.openTeams ? row.openTeams.toLocaleString() : '0'}</td>
      <td style="border: 0.5pt solid black; text-align: center; font-family: 'Times New Roman', 'Khmer OS Siemreap', serif; font-size: 8.5pt; padding: 0.5pt 1pt; vertical-align: middle;">${row.k1ToK2 ? row.k1ToK2.toLocaleString() : '0'}</td>
      <td style="border: 0.5pt solid black; text-align: center; font-family: 'Times New Roman', 'Khmer OS Siemreap', serif; font-size: 8.5pt; padding: 0.5pt 1pt; vertical-align: middle;">${row.teamReturnedToK2 ? row.teamReturnedToK2.toLocaleString() : '0'}</td>
      <td style="border: 0.5pt solid black; text-align: right; font-family: 'Times New Roman', 'Khmer OS Siemreap', serif; font-size: 8.5pt; padding: 0.5pt 2pt; vertical-align: middle;">${row.issuedK2ToTeams ? row.issuedK2ToTeams.toLocaleString() : '0'}</td>
      <td style="border: 0.5pt solid black; text-align: right; font-family: 'Times New Roman', 'Khmer OS Siemreap', serif; font-size: 8.5pt; padding: 0.5pt 2pt; vertical-align: middle;">${row.teamsUsed ? row.teamsUsed.toLocaleString() : '0'}</td>
      <td style="border: 0.5pt solid black; text-align: center; font-family: 'Times New Roman', 'Khmer OS Siemreap', serif; font-size: 8.5pt; padding: 0.5pt 1pt; vertical-align: middle;">${row.damagedTeam ? row.damagedTeam.toLocaleString() : '0'}</td>
      <td style="border: 0.5pt solid black; text-align: center; font-family: 'Times New Roman', 'Khmer OS Siemreap', serif; font-size: 8.5pt; padding: 0.5pt 1pt; vertical-align: middle;">${row.damagedK2 ? row.damagedK2.toLocaleString() : '0'}</td>
      <td style="border: 0.5pt solid black; text-align: center; font-family: 'Times New Roman', 'Khmer OS Siemreap', serif; font-size: 8.5pt; padding: 0.5pt 1pt; vertical-align: middle;">${row.testSampleK2 ? row.testSampleK2.toLocaleString() : '0'}</td>
      <td style="border: 0.5pt solid black; text-align: right; font-family: 'Times New Roman', 'Khmer OS Siemreap', serif; font-size: 8.5pt; padding: 0.5pt 2pt; vertical-align: middle;">${row.endingK2 ? row.endingK2.toLocaleString() : '0'}</td>
      <td style="border: 0.5pt solid black; text-align: right; font-family: 'Times New Roman', 'Khmer OS Siemreap', serif; font-size: 8.5pt; padding: 0.5pt 2pt; vertical-align: middle;">${row.endingTeams ? row.endingTeams.toLocaleString() : '0'}</td>
      <td style="border: 0.5pt solid black; text-align: right; font-weight: bold; font-family: 'Times New Roman', 'Khmer OS Siemreap', serif; font-size: 8.5pt; padding: 0.5pt 2pt; vertical-align: middle;">${row.grandTotalEnding ? row.grandTotalEnding.toLocaleString() : '0'}</td>
    </tr>`
    )
    .join('');

  // 5-Row Table for Individual Team Report
  const isTeamReport = viewMode === 'team' && !!teamStickerData;
  const effectiveTeamName = teamDisplayName || 'កំពង់ផែ កោះកុង';

  let bodyContentHtml = '';

  if (isTeamReport && teamStickerData) {
    const renderTeamNum = (n: number | undefined) => (n !== undefined && n !== null ? n.toLocaleString() : '0');

    bodyContentHtml = `
  <!-- Header Two-Column Layout (Ministry Info on Left, Kingdom on Right) -->
  <table class="layout-table" style="margin-bottom: 4pt; width: 100%;">
    <tr>
      <td style="width: 50%; text-align: center; vertical-align: top; padding-top: 3pt;"></td>
      <td style="width: 50%; text-align: center; vertical-align: top; padding-top: 3pt;">
        <div class="font-moul" style="font-size: 12pt; line-height: 1.35; margin-bottom: 2pt; white-space: nowrap;">ព្រះរាជាណាចក្រកម្ពុជា</div>
      </td>
    </tr>
    <tr>
      <td style="width: 50%; text-align: center; vertical-align: top;">
        <div class="font-moul" style="font-size: 12pt; line-height: 1.35; margin-bottom: 1.5pt; white-space: nowrap;">ក្រសួងមហាផ្ទៃ</div>
        <div class="font-moul" style="font-size: 12pt; line-height: 1.35; margin-bottom: 1.5pt; white-space: nowrap;">អគ្គនាយកដ្ឋានអន្តោប្រវេសន៍</div>
        <div class="font-moul" style="font-size: 12pt; line-height: 1.35; margin-bottom: 1.5pt; white-space: nowrap;">នាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត</div>
        <div class="font-moul" style="font-size: 12pt; line-height: 1.35; margin-bottom: 1.5pt; white-space: nowrap;">ការិយាល័យទិដ្ឋាការចូល</div>
        <div class="font-siemreap" style="font-weight: bold; font-size: 10pt; line-height: 1.35; margin-bottom: 1.5pt; white-space: nowrap;">${effectiveTeamName}</div>
        <div style="text-align: center; margin-top: 2pt; margin-bottom: 1pt;">
          <img src="${tacteingSrc}" width="110" height="12" style="display: inline-block; vertical-align: middle; max-width: 115px; max-height: 14px;" alt="តាក់តែង" />
        </div>
      </td>
      <td style="width: 50%; text-align: center; vertical-align: top;">
        <div class="font-moul" style="font-size: 12pt; line-height: 1.35; margin-bottom: 2pt; white-space: nowrap;">ជាតិ សាសនា ព្រះមហាក្សត្រ</div>
        <div style="text-align: center; margin-top: 2pt;">
          <img src="${tacteingSrc}" width="130" height="14" style="display: inline-block; vertical-align: middle; max-width: 140px; max-height: 16px;" alt="តាក់តែង" />
        </div>
      </td>
    </tr>
  </table>

  <!-- Letter To Salutation -->
  <div style="text-align: center; margin-top: 14pt; margin-bottom: 8pt;">
    <div class="font-moul" style="font-size: 14pt; line-height: 1.3;">សូមគោរពជូន</div>
    <div class="font-moul" style="font-size: 12pt; line-height: 1.3; margin-top: 2pt;">លោកវរសេនីយ៍ទោ នាយការិយាល័យទិដ្ឋាការចូល</div>
  </div>

  <!-- Subject Details -->
  <table class="layout-table" style="margin-top: 6pt; margin-bottom: 2pt; font-size: 12pt; line-height: 1.7;">
    <tr>
      <td style="width: 65pt; white-space: nowrap; vertical-align: top; line-height: 1.7;">
        <span class="font-moul" style="font-size: 12pt;">កម្មវត្ថុ</span>
        <span class="font-siemreap" style="font-weight: bold; float: right; padding-right: 4pt;">៖</span>
      </td>
      <td style="vertical-align: top; font-size: 12pt; line-height: 1.7;" class="font-siemreap text-justify">
        របាយការណ៍ ស្តីពីការប្រើប្រាស់សន្លឹកទិដ្ឋាការស្អិត ${teamPeriodKhmerTitle || reportPeriodLabel} ។
      </td>
    </tr>
  </table>

  <!-- Introductory Body Paragraph dropped down properly -->
  <div class="text-justify font-siemreap" style="font-size: 12pt; line-height: 1.75; text-indent: ${tabStopCm}cm; margin-top: ${paragraphShiftY !== undefined ? paragraphShiftY : 6}pt; margin-bottom: 8pt;">
    តបតាមកម្មវត្ថុខាងលើ ${effectiveTeamName} សូមជម្រាបជូន <span class="font-moul" style="font-size: 12pt;">លោកវរសេនីយ៍ទោ នាយការិយាល័យទិដ្ឋាការចូល</span> មេត្តាជ្រាបថា៖ ការប្រើប្រាស់សន្លឹកទិដ្ឋាការស្អិត របស់${effectiveTeamName} ${teamPeriodShortKhmerTitle || reportPeriodLabel} ដោយគិតចាប់ពីថ្ងៃទី${startKhmerDate.day} ខែ${startKhmerDate.month} ឆ្នាំ${startKhmerDate.year} រហូតដល់ថ្ងៃទី${endKhmerDate.day} ខែ${endKhmerDate.month} ឆ្នាំ${endKhmerDate.year} មានចំនួនដូចខាងក្រោម ៖
  </div>

  <!-- 5-Row Team Stock Table -->
  <div style="margin-top: 2pt; margin-bottom: 4pt; page-break-inside: avoid;">
    <table class="report-table" style="font-size: 8.5pt; text-align: center; line-height: 1.2; width: 100%; border-collapse: collapse; table-layout: fixed;">
      <colgroup>
        <col style="width: 14%;" />
        ${VISA_TYPES_LIST.map((vt) => `<col style="width: ${vt === 'T' || vt === 'E' ? '8%' : vt === 'D' || vt === 'K' ? '7%' : '5.333%'};" />`).join('')}
        <col style="width: 8%;" />
      </colgroup>
      <thead>
        <tr style="background-color: #ffffff; font-weight: normal; mso-height-rule: exactly; height: 0.50cm; mso-yfti-tblheader: yes; page-break-inside: avoid;">
          <th style="width: 14%; font-family: 'Khmer OS Siemreap', sans-serif; font-size: 8pt; font-weight: normal; padding: 2pt 2pt; border: 0.5pt solid black; vertical-align: middle;"></th>
          ${VISA_TYPES_LIST.map((vt) => {
            const w = vt === 'T' || vt === 'E' ? '8%' : vt === 'D' || vt === 'K' ? '7%' : '5.333%';
            return `<th style="width: ${w}; font-family: 'Times New Roman', serif; font-size: 8.5pt; font-weight: bold; padding: 2pt 1pt; border: 0.5pt solid black; vertical-align: middle; white-space: nowrap;">${vt}</th>`;
          }).join('')}
          <th style="width: 8%; font-family: 'Khmer OS Siemreap', sans-serif; font-size: 8.5pt; font-weight: bold; padding: 2pt 1pt; border: 0.5pt solid black; vertical-align: middle; white-space: nowrap;">សរុប</th>
        </tr>
      </thead>
      <tbody>
        <!-- Row 1: ចំនួនសល់ខែចាស់ -->
        <tr style="mso-height-rule: exactly; height: 0.45cm; page-break-inside: avoid;">
          <td style="width: 14%; border: 0.5pt solid black; text-align: left; font-family: 'Khmer OS Siemreap', sans-serif; font-size: 8pt; padding: 2pt 3pt; vertical-align: middle; white-space: nowrap;">ចំនួនសល់ខែចាស់</td>
          ${VISA_TYPES_LIST.map((vt) => {
            const w = vt === 'T' || vt === 'E' ? '8%' : vt === 'D' || vt === 'K' ? '7%' : '5.333%';
            return `<td style="width: ${w}; border: 0.5pt solid black; text-align: center; font-family: 'Times New Roman', serif; font-size: 8.5pt; padding: 2pt 1pt; vertical-align: middle; white-space: nowrap;">${renderTeamNum(teamStickerData.opening.values[vt])}</td>`;
          }).join('')}
          <td style="width: 8%; border: 0.5pt solid black; text-align: center; font-family: 'Times New Roman', serif; font-size: 8.5pt; font-weight: bold; padding: 2pt 1pt; vertical-align: middle; white-space: nowrap;">${renderTeamNum(teamStickerData.opening.total)}</td>
        </tr>

        <!-- Row 2: ចំនួនបើកថ្មី -->
        <tr style="mso-height-rule: exactly; height: 0.45cm; page-break-inside: avoid;">
          <td style="width: 14%; border: 0.5pt solid black; text-align: left; font-family: 'Khmer OS Siemreap', sans-serif; font-size: 8pt; padding: 2pt 3pt; vertical-align: middle; white-space: nowrap;">ចំនួនបើកថ្មី</td>
          ${VISA_TYPES_LIST.map((vt) => {
            const w = vt === 'T' || vt === 'E' ? '8%' : vt === 'D' || vt === 'K' ? '7%' : '5.333%';
            return `<td style="width: ${w}; border: 0.5pt solid black; text-align: center; font-family: 'Times New Roman', serif; font-size: 8.5pt; padding: 2pt 1pt; vertical-align: middle; white-space: nowrap;">${renderTeamNum(teamStickerData.newIssued.values[vt])}</td>`;
          }).join('')}
          <td style="width: 8%; border: 0.5pt solid black; text-align: center; font-family: 'Times New Roman', serif; font-size: 8.5pt; font-weight: bold; padding: 2pt 1pt; vertical-align: middle; white-space: nowrap;">${renderTeamNum(teamStickerData.newIssued.total)}</td>
        </tr>

        <!-- Row 3: ផ្ទេរ/បង្វិល/ខូច/ខ្វះ -->
        <tr style="mso-height-rule: exactly; height: 0.45cm; page-break-inside: avoid;">
          <td style="width: 14%; border: 0.5pt solid black; text-align: left; font-family: 'Khmer OS Siemreap', sans-serif; font-size: 8pt; padding: 2pt 3pt; vertical-align: middle; white-space: nowrap;">ផ្ទេរ/បង្វិល/ខូច/ខ្វះ</td>
          ${VISA_TYPES_LIST.map((vt) => {
            const w = vt === 'T' || vt === 'E' ? '8%' : vt === 'D' || vt === 'K' ? '7%' : '5.333%';
            return `<td style="width: ${w}; border: 0.5pt solid black; text-align: center; font-family: 'Times New Roman', serif; font-size: 8.5pt; padding: 2pt 1pt; vertical-align: middle; white-space: nowrap;">${renderTeamNum(teamStickerData.returnedDamagedMissing.values[vt])}</td>`;
          }).join('')}
          <td style="width: 8%; border: 0.5pt solid black; text-align: center; font-family: 'Times New Roman', serif; font-size: 8.5pt; font-weight: bold; padding: 2pt 1pt; vertical-align: middle; white-space: nowrap;">${renderTeamNum(teamStickerData.returnedDamagedMissing.total)}</td>
        </tr>

        <!-- Row 4: ចំនួនប្រើ -->
        <tr style="mso-height-rule: exactly; height: 0.45cm; page-break-inside: avoid;">
          <td style="width: 14%; border: 0.5pt solid black; text-align: left; font-family: 'Khmer OS Siemreap', sans-serif; font-size: 8pt; padding: 2pt 3pt; vertical-align: middle; white-space: nowrap;">ចំនួនប្រើ</td>
          ${VISA_TYPES_LIST.map((vt) => {
            const w = vt === 'T' || vt === 'E' ? '8%' : vt === 'D' || vt === 'K' ? '7%' : '5.333%';
            return `<td style="width: ${w}; border: 0.5pt solid black; text-align: center; font-family: 'Times New Roman', serif; font-size: 8.5pt; padding: 2pt 1pt; vertical-align: middle; white-space: nowrap;">${renderTeamNum(teamStickerData.used.values[vt])}</td>`;
          }).join('')}
          <td style="width: 8%; border: 0.5pt solid black; text-align: center; font-family: 'Times New Roman', serif; font-size: 8.5pt; font-weight: bold; padding: 2pt 1pt; vertical-align: middle; white-space: nowrap;">${renderTeamNum(teamStickerData.used.total)}</td>
        </tr>

        <!-- Row 5: សល់ខែបន្ទាប់ -->
        <tr style="font-weight: bold; mso-height-rule: exactly; height: 0.45cm; page-break-inside: avoid;">
          <td style="width: 14%; border: 0.5pt solid black; text-align: left; font-family: 'Khmer OS Siemreap', sans-serif; font-size: 8pt; padding: 2pt 3pt; vertical-align: middle; white-space: nowrap;">សល់ខែបន្ទាប់</td>
          ${VISA_TYPES_LIST.map((vt) => {
            const w = vt === 'T' || vt === 'E' ? '8%' : vt === 'D' || vt === 'K' ? '7%' : '5.333%';
            return `<td style="width: ${w}; border: 0.5pt solid black; text-align: center; font-family: 'Times New Roman', serif; font-size: 8.5pt; padding: 2pt 1pt; vertical-align: middle; white-space: nowrap;">${renderTeamNum(teamStickerData.ending.values[vt])}</td>`;
          }).join('')}
          <td style="width: 8%; border: 0.5pt solid black; text-align: center; font-family: 'Times New Roman', serif; font-size: 8.5pt; font-weight: bold; padding: 2pt 1pt; vertical-align: middle; white-space: nowrap;">${renderTeamNum(teamStickerData.ending.total)}</td>
        </tr>
      </tbody>
    </table>
  </div>

  <!-- Concluding Salutations -->
  <div style="font-size: 12pt; line-height: 1.7; margin-top: 12pt; margin-bottom: 6pt; page-break-inside: avoid;" class="font-siemreap">
    <div class="text-justify" style="text-indent: ${tabStopCm}cm; margin-bottom: 4pt; line-height: 1.7;">
      អាស្រ័យដូចបានជម្រាបជូនខាងលើ សូម <span class="font-moul" style="font-size: 12pt;">លោកវរសេនីយ៍ទោ នាយការិយាល័យទិដ្ឋាការចូល</span> មេត្តា ពិនិត្យ និងជ្រាប ជារាយការណ៍ដោយក្តីអនុគ្រោះ ។
    </div>
    <div class="text-justify" style="text-indent: ${tabStopCm}cm; line-height: 1.7;">
      សូម<span class="font-moul" style="font-size: 12pt;">លោកវរសេនីយ៍ទោ</span> មេត្តាទទួលនូវការគោរព អំពីខ្ញុំ ។
    </div>
  </div>

  <!-- Signatures Section -->
  <table class="layout-table" style="margin-top: 3pt; font-size: 12pt; line-height: 1.25; page-break-inside: avoid;">
    <tr>
      <td style="width: 50%;"></td>
      <td style="width: 50%; text-align: center; vertical-align: top;">
        <div style="font-size: 12pt;">
          ${signLunarDate || 'ថ្ងៃ.....................ខែ.............ឆ្នាំ..............សំរឹទ្ធិស័ក ព.ស. ២៥៦....'}
        </div>
        <div style="font-size: 12pt;">
          ${signSolarDate || `${effectiveTeamName ? (effectiveTeamName.includes('កំពង់ផែ') || effectiveTeamName.includes('ព្រំដែន') || effectiveTeamName.includes('អាកាស') ? effectiveTeamName : 'ច្រកទ្វារ' + effectiveTeamName) : ''}, ថ្ងៃទី${signKhmerDate.day} ខែ${signKhmerDate.month} ឆ្នាំ${signKhmerDate.year}`}
        </div>
        <div class="font-moul" style="font-size: 12pt; margin-top: 2pt;">${rightSignTitle || `ប្រធាន${effectiveTeamName}`}</div>
        <div style="height: 42pt;"></div>
        ${rightSignOfficer ? `<div style="font-weight: bold; font-size: 12pt;">${rightSignOfficer}</div>` : ''}
      </td>
    </tr>
  </table>
    `;
  } else {
    bodyContentHtml = `
  <!-- Header Two-Column Layout (Ministry Info on Left, Kingdom on Right) -->
  <table class="layout-table" style="margin-bottom: 1pt;">
    <tr>
      <td style="width: 50%; text-align: center; vertical-align: top; padding-top: 14pt;">
        ${ministryHierarchy
          .map(
            (line) =>
              `<div class="font-moul" style="font-size: 12pt; line-height: 1.2; white-space: nowrap;">${line}</div>`
          )
          .join('')}
        <div style="text-align: center; margin-top: 1pt; margin-bottom: 0.5pt;">
          <!-- Tacteing Ornament Line under Administration in Word -->
          <img src="${tacteingSrc}" width="115" height="12" style="display: inline-block; vertical-align: middle; max-width: 120px; max-height: 14px;" alt="តាក់តែង" />
        </div>
        ${variant === 'director' ? `
        <div class="font-siemreap" style="font-size: 12pt; margin-top: 0.5pt; text-align: center; white-space: nowrap;">
          លេខ ៖ ${docNumber ? docNumber.replace(/\|/g, '') : '...................របក/២៦'}
        </div>` : ''}
      </td>
      <td style="width: 50%; text-align: center; vertical-align: top;">
        <div class="font-moul" style="font-size: 12pt; line-height: 1.2; white-space: nowrap;">ព្រះរាជាណាចក្រកម្ពុជា</div>
        <div class="font-moul" style="font-size: 12pt; line-height: 1.2; white-space: nowrap;">ជាតិ សាសនា ព្រះមហាក្សត្រ</div>
        <div style="text-align: center; margin-top: 1pt;">
          <!-- Tacteing Ornament Line under Kingdom in Word -->
          <img src="${tacteingSrc}" width="135" height="14" style="display: inline-block; vertical-align: middle; max-width: 145px; max-height: 16px;" alt="តាក់តែង" />
        </div>
      </td>
    </tr>
  </table>

  <!-- Letter To Salutation -->
  <div style="text-align: center; margin-top: 3pt; margin-bottom: 2pt;">
    <div class="font-moul" style="font-size: 14pt; line-height: 1.1; margin: 0; padding: 0;">សូមគោរពជូន</div>
    ${variant === 'director' && docRecipientRank ? `<div class="font-moul" style="font-size: 12pt; line-height: 1.1; margin-top: -1pt; padding: 0;">${docRecipientRank}</div>` : ''}
    <div class="font-moul" style="font-size: 12pt; line-height: 1.1; margin-top: -1pt; padding: 0;">${docRecipientTitle}</div>
  </div>

  <!-- Subject & Through Details with Aligned Colons -->
  <table class="layout-table" style="margin-top: 2pt; margin-bottom: 2pt; font-size: 12pt; line-height: 1.25;">
    <tr>
      <td style="width: 75pt; white-space: nowrap; vertical-align: top;">
        <span class="font-moul" style="font-size: 12pt;">តាមរយៈ</span>
        <span class="font-siemreap" style="font-weight: bold; float: right; padding-right: 4pt; font-size: 12pt;">៖</span>
      </td>
      <td style="vertical-align: top; font-size: 12pt;" class="font-siemreap">
        ${(docThroughTitle || '').replace(/លោកនាយរងការិយាល័យទិទទួលបន្ទុក/g, 'លោកនាយរងការិយាល័យទទួលបន្ទុក').replace(/ទិទទួលបន្ទុក/g, 'ទទួលបន្ទុក')}
      </td>
    </tr>
    <tr>
      <td style="width: 75pt; white-space: nowrap; vertical-align: top; padding-top: 1pt;">
        <span class="font-moul" style="font-size: 12pt;">កម្មវត្ថុ</span>
        <span class="font-siemreap" style="font-weight: bold; float: right; padding-right: 4pt; font-size: 12pt;">៖</span>
      </td>
      <td style="vertical-align: top; padding-top: 1pt; font-size: 12pt;" class="font-siemreap text-justify">
        ${proposalType === 'section' || proposalType === 'office'
          ? 'សំណើសុំការអនុញ្ញាតប្រគល់ គល់ទិដ្ឋាការស្អិតដែលបានប្រើប្រាស់រួច ពីបណ្តាក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអន្តរជាតិនានា ជូនទៅការិយាល័យរដ្ឋបាលសរុប។'
          : `របាយការណ៍ស្តីពីការបើក ការផ្តល់សន្លឹកទិដ្ឋាការស្អិត និងក្រដាសអនុម័តផ្តល់ទិដ្ឋាការអេឡិចត្រូនិកនៅពេលមកដល់ នៅតាមបណ្តាក្រុមផ្តល់ទិដ្ឋាការ ${reportPeriodLabel}។`}
      </td>
    </tr>
    ${proposalType === 'section' ? `
    <tr>
      <td style="width: 75pt; white-space: nowrap; vertical-align: top; padding-top: 1pt;">
        <span class="font-moul" style="font-size: 12pt;">យោង</span>
        <span class="font-siemreap" style="font-weight: bold; float: right; padding-right: 4pt; font-size: 12pt;">៖</span>
      </td>
      <td style="vertical-align: top; padding-top: 1pt; font-size: 12pt;" class="font-siemreap text-justify">
        ${docReferenceText || 'លិខិតប្រគល់ទទួលតាមក្រុមផ្តល់ទិដ្ឋាការ ។'}
      </td>
    </tr>` : ''}
  </table>

  <!-- Introductory Body Paragraph dropped down properly -->
  <div class="text-justify font-siemreap" style="font-size: 12pt; line-height: 1.25; text-indent: ${tabStopCm}cm; margin-top: ${paragraphShiftY !== undefined ? paragraphShiftY : 5}pt; margin-bottom: 3pt;">
    ${proposalType === 'section'
      ? `សេចក្តីដូចមានចែងក្នុងកម្មវត្ថុ និងយោងខាងលើ ផ្នែករដ្ឋបាល មានកិត្តិយសសូមជម្រាប ជូន<span class="font-moul" style="font-size: 12pt;">${introSalutation || docRecipientTitle}</span> មេត្តាជ្រាបថា ៖ ផ្នែកស្នើសុំការអនុញ្ញាតប្រគល់ គល់ទិដ្ឋាការដែលប្រើប្រាស់រួច ពីបណ្តាក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអន្តរជាតិនានា ទៅការិយាល័យរដ្ឋបាលសរុប មានចំនួនដូចខាងក្រោម ៖`
      : variant === 'director'
      ? `សេចក្តីដូចមានចែងក្នុងកម្មវត្ថុខាងលើ <span style="font-weight: bold;">ការិយាល័យទិដ្ឋាការចូល</span> មានកិត្តិយសសូមជម្រាបជូន<span class="font-moul" style="font-size: 12pt;">${introSalutation || docRecipientTitle}</span> មេត្តាជ្រាបថា ៖ នៅក្នុងកិច្ចប្រតិបត្តិការ${reportPeriodLabel} ដោយគិតពីថ្ងៃទី${startKhmerDate.day} ខែ${startKhmerDate.month} ឆ្នាំ${startKhmerDate.year} រហូតដល់ថ្ងៃទី${endKhmerDate.day} ខែ${endKhmerDate.month} ឆ្នាំ${endKhmerDate.year} ការិយាល័យសម្រេចលទ្ធផលការបើក ការផ្តល់សន្លឹកទិដ្ឋាការ និងផ្តល់ក្រដាសអនុម័តទិដ្ឋាការអេឡិចត្រូនិកពេលមកដល់ នៅតាមបណ្តាក្រុមផ្តល់ទិដ្ឋាការ មានដូចខាងក្រោម៖`
      : `តបតាមកម្មវត្ថុខាងលើ ផ្នែករដ្ឋបាល មានកិត្តិយសសូមជម្រាបជូន <span class="font-moul" style="font-size: 12pt;">${introSalutation || docRecipientTitle}</span> មេត្តាជ្រាបថា ៖ នៅក្នុងកិច្ចប្រតិបត្តិការ ${reportPeriodLabel} ដោយគិតចាប់ពីថ្ងៃទី${startKhmerDate.day} ខែ${startKhmerDate.month} ឆ្នាំ${startKhmerDate.year} រហូតដល់ថ្ងៃទី${endKhmerDate.day} ខែ${endKhmerDate.month} ឆ្នាំ${endKhmerDate.year} ផ្នែកសម្រេចលទ្ធផលការបើក ការផ្តល់សន្លឹកទិដ្ឋាការ និងផ្តល់ក្រដាសអនុម័តផ្តល់ទិដ្ឋាការអេឡិចត្រូនិកពេលមកដល់ នៅតាមបណ្តាក្រុមផ្តល់ទិដ្ឋាការ មានដូចខាងក្រោម ៖`}
  </div>

  ${proposalType === 'section' ? `
  <!-- Stub Proposal Table matching Picture 2 -->
  <div style="margin-top: 2pt; margin-bottom: 2pt; page-break-inside: avoid;">
    <table class="report-table" style="font-size: 12pt; text-align: center; line-height: 1.25; width: 100%; border-collapse: collapse;">
      <thead>
        <tr style="background-color: #ffffff; font-weight: bold; mso-yfti-tblheader: yes; page-break-inside: avoid;">
          <th rowspan="2" style="width: 10%; font-family: 'Khmer OS Siemreap', sans-serif; font-size: 12pt; font-weight: bold; padding: 2pt 2pt; border: 0.5pt solid black; vertical-align: middle;">ល.រ</th>
          <th rowspan="2" style="width: 18%; font-family: 'Khmer OS Siemreap', sans-serif; font-size: 12pt; font-weight: bold; padding: 2pt 2pt; border: 0.5pt solid black; vertical-align: middle;">ប្រភេទ</th>
          <th colspan="2" style="width: 48%; font-family: 'Khmer OS Siemreap', sans-serif; font-size: 12pt; font-weight: bold; padding: 2pt 2pt; border: 0.5pt solid black; vertical-align: middle;">ឯកតា</th>
          <th rowspan="2" style="width: 24%; font-family: 'Khmer OS Siemreap', sans-serif; font-size: 12pt; font-weight: bold; padding: 2pt 2pt; border: 0.5pt solid black; vertical-align: middle;">ផ្សេង ៗ</th>
        </tr>
        <tr style="background-color: #ffffff; font-weight: bold; mso-yfti-tblheader: yes; page-break-inside: avoid;">
          <th style="width: 24%; font-family: 'Khmer OS Siemreap', sans-serif; font-size: 12pt; font-weight: bold; padding: 2pt 2pt; border: 0.5pt solid black; vertical-align: middle;">ក្បាល</th>
          <th style="width: 24%; font-family: 'Khmer OS Siemreap', sans-serif; font-size: 12pt; font-weight: bold; padding: 2pt 2pt; border: 0.5pt solid black; vertical-align: middle;">សន្លឹក</th>
        </tr>
      </thead>
      <tbody>
        ${stubProposalRowsHtml}
        <tr style="font-weight: bold; background-color: #ffffff; page-break-inside: avoid;">
          <td colspan="2" style="border: 0.5pt solid black; text-align: center; font-family: 'Khmer OS Muol Light', 'Moul', serif; font-size: 12pt; padding: 2pt 2pt; vertical-align: middle;">សរុប</td>
          <td style="border: 0.5pt solid black; text-align: center; font-family: 'Times New Roman', serif; font-size: 12pt; font-weight: bold; padding: 2pt 2pt; vertical-align: middle;">${isPendingCalc ? '-' : stubTotalBooklets.toLocaleString()}</td>
          <td style="border: 0.5pt solid black; text-align: center; font-family: 'Times New Roman', serif; font-size: 12pt; font-weight: bold; padding: 2pt 2pt; vertical-align: middle;">${isPendingCalc ? '-' : stubTotalSheets.toLocaleString()}</td>
          <td style="border: 0.5pt solid black; text-align: center; font-family: 'Khmer OS Siemreap', sans-serif; font-size: 12pt; padding: 2pt 2pt; border: 0.5pt solid black; vertical-align: middle;"></td>
        </tr>
      </tbody>
    </table>
  </div>
  ` : `
  <!-- Section ក. សន្លឹកទិដ្ឋាការស្អិត -->
  <div style="margin-top: 1pt; margin-bottom: 1pt; page-break-inside: avoid;">
    <div class="font-moul" style="font-size: 12pt; margin-bottom: 0.5pt;">
      ក. សន្លឹកទិដ្ឋាការស្អិត
    </div>

    <table class="report-table" style="font-size: 8pt; text-align: center; line-height: 1.15;">
      <tbody>
        <tr style="background-color: #ffffff; font-weight: normal; font-size: 6pt; mso-height-rule: exactly; height: 0.52cm; mso-yfti-tblheader: yes; page-break-inside: avoid;">
          <th rowspan="2" style="width: 18pt; font-family: 'Khmer OS Siemreap', sans-serif; font-size: 6pt; font-weight: normal; padding: 0.5pt 1pt; border: 0.5pt solid black; vertical-align: middle;">ល.រ</th>
          <th rowspan="2" style="width: 28pt; font-family: 'Khmer OS Siemreap', sans-serif; font-size: 6pt; font-weight: normal; padding: 0.5pt 1pt; border: 0.5pt solid black; vertical-align: middle;">ប្រភេទ</th>
          <th colspan="2" style="font-family: 'Khmer OS Siemreap', sans-serif; font-size: 6pt; font-weight: normal; padding: 0.5pt 1pt; border: 0.5pt solid black; vertical-align: middle;">
            សន្និធិ<br>ចុងគ្រា ${prevKhmerDate.day}-${prevKhmerDate.month}-${prevKhmerDate.year}
          </th>
          <th colspan="7" style="font-family: 'Khmer OS Siemreap', sans-serif; font-size: 6pt; font-weight: normal; padding: 0.5pt 1pt; border: 0.5pt solid black; vertical-align: middle;">
            ${startKhmerDate.day}-${startKhmerDate.month}-${startKhmerDate.year} រហូតដល់ ${endKhmerDate.day}-${endKhmerDate.month}-${endKhmerDate.year}
          </th>
          <th colspan="2" style="font-family: 'Khmer OS Siemreap', sans-serif; font-size: 6pt; font-weight: normal; padding: 0.5pt 1pt; border: 0.5pt solid black; vertical-align: middle;">
            សន្និធិ<br>សល់ ${endKhmerDate.day}-${endKhmerDate.month}-${endKhmerDate.year}
          </th>
          <th rowspan="2" style="width: 44pt; font-family: 'Khmer OS Siemreap', sans-serif; font-size: 6pt; font-weight: normal; padding: 0.5pt 1pt; border: 0.5pt solid black; vertical-align: middle;">
            សន្និធិសរុប<br>ក២និងក្រុម
          </th>
        </tr>
        <tr style="background-color: #ffffff; font-size: 6pt; font-weight: normal; mso-height-rule: exactly; height: 0.34cm; line-height: 1.0; mso-yfti-tblheader: yes; page-break-inside: avoid;">
          <th style="font-family: 'Khmer OS Siemreap', sans-serif; font-size: 6pt; font-weight: normal; padding: 0.5pt 1pt; border: 0.5pt solid black; vertical-align: middle;">ក២</th>
          <th style="font-family: 'Khmer OS Siemreap', sans-serif; font-size: 6pt; font-weight: normal; padding: 0.5pt 1pt; border: 0.5pt solid black; vertical-align: middle;">ក្រុម</th>
          <th style="font-family: 'Khmer OS Siemreap', sans-serif; font-size: 6pt; font-weight: normal; padding: 0.5pt 1pt; border: 0.5pt solid black; vertical-align: middle;">ទិដ្ឋាការបើក<br>ពីក១</th>
          <th style="font-family: 'Khmer OS Siemreap', sans-serif; font-size: 6pt; font-weight: normal; padding: 0.5pt 1pt; border: 0.5pt solid black; vertical-align: middle;">ទិដ្ឋាការបង្វិល<br>ពីក្រុម</th>
          <th style="font-family: 'Khmer OS Siemreap', sans-serif; font-size: 6pt; font-weight: normal; padding: 0.5pt 1pt; border: 0.5pt solid black; vertical-align: middle;">ទិដ្ឋាការបើក<br>ផ្តល់ទៅក្រុម</th>
          <th style="font-family: 'Khmer OS Siemreap', sans-serif; font-size: 6pt; font-weight: normal; padding: 0.5pt 1pt; border: 0.5pt solid black; vertical-align: middle;">ក្រុម<br>ប្រើប្រាស់</th>
          <th style="font-family: 'Khmer OS Siemreap', sans-serif; font-size: 6pt; font-weight: normal; padding: 0.5pt 1pt; border: 0.5pt solid black; vertical-align: middle;">ទិដ្ឋាការក្រុម<br>មិនបានការ</th>
          <th style="font-family: 'Khmer OS Siemreap', sans-serif; font-size: 6pt; font-weight: normal; padding: 0.5pt 1pt; border: 0.5pt solid black; vertical-align: middle;">មិនបានការ<br>ក២</th>
          <th style="font-family: 'Khmer OS Siemreap', sans-serif; font-size: 6pt; font-weight: normal; padding: 0.5pt 1pt; border: 0.5pt solid black; vertical-align: middle;">ទិដ្ឋាការសាក<br>ល្បង ក២</th>
          <th style="font-family: 'Khmer OS Siemreap', sans-serif; font-size: 6pt; font-weight: normal; padding: 0.5pt 1pt; border: 0.5pt solid black; vertical-align: middle;">ក២</th>
          <th style="font-family: 'Khmer OS Siemreap', sans-serif; font-size: 6pt; font-weight: normal; padding: 0.5pt 1pt; border: 0.5pt solid black; vertical-align: middle;">ក្រុម</th>
        </tr>
        ${rowsHtml}
        <!-- Grand Total Row -->
        <tr style="font-weight: bold; background-color: #ffffff; mso-height-rule: exactly; height: 0.36cm; page-break-inside: avoid;">
          <td colspan="2" style="border: 0.5pt solid black; text-align: center; font-family: 'Khmer OS Muol Light', 'Moul', serif; font-size: 8pt; padding: 0.5pt 1pt; vertical-align: middle;">សរុប</td>
          <td style="border: 0.5pt solid black; text-align: right; font-family: 'Times New Roman', serif; font-size: 8.5pt; font-weight: bold; padding: 0.5pt 2pt; vertical-align: middle;">${stickerTotalRow.openK2.toLocaleString()}</td>
          <td style="border: 0.5pt solid black; text-align: right; font-family: 'Times New Roman', serif; font-size: 8.5pt; font-weight: bold; padding: 0.5pt 2pt; vertical-align: middle;">${stickerTotalRow.openTeams.toLocaleString()}</td>
          <td style="border: 0.5pt solid black; text-align: center; font-family: 'Times New Roman', serif; font-size: 8.5pt; font-weight: bold; padding: 0.5pt 1pt; vertical-align: middle;">${stickerTotalRow.k1ToK2.toLocaleString()}</td>
          <td style="border: 0.5pt solid black; text-align: center; font-family: 'Times New Roman', serif; font-size: 8.5pt; font-weight: bold; padding: 0.5pt 1pt; vertical-align: middle;">${stickerTotalRow.teamReturnedToK2.toLocaleString()}</td>
          <td style="border: 0.5pt solid black; text-align: right; font-family: 'Times New Roman', serif; font-size: 8.5pt; font-weight: bold; padding: 0.5pt 2pt; vertical-align: middle;">${stickerTotalRow.issuedK2ToTeams.toLocaleString()}</td>
          <td style="border: 0.5pt solid black; text-align: right; font-family: 'Times New Roman', serif; font-size: 8.5pt; font-weight: bold; padding: 0.5pt 2pt; vertical-align: middle;">${stickerTotalRow.teamsUsed.toLocaleString()}</td>
          <td style="border: 0.5pt solid black; text-align: center; font-family: 'Times New Roman', serif; font-size: 8.5pt; font-weight: bold; padding: 0.5pt 1pt; vertical-align: middle;">${stickerTotalRow.damagedTeam.toLocaleString()}</td>
          <td style="border: 0.5pt solid black; text-align: center; font-family: 'Times New Roman', serif; font-size: 8.5pt; font-weight: bold; padding: 0.5pt 1pt; vertical-align: middle;">${stickerTotalRow.damagedK2.toLocaleString()}</td>
          <td style="border: 0.5pt solid black; text-align: center; font-family: 'Times New Roman', serif; font-size: 8.5pt; font-weight: bold; padding: 0.5pt 1pt; vertical-align: middle;">${stickerTotalRow.testSampleK2.toLocaleString()}</td>
          <td style="border: 0.5pt solid black; text-align: right; font-family: 'Times New Roman', serif; font-size: 8.5pt; font-weight: bold; padding: 0.5pt 2pt; vertical-align: middle;">${stickerTotalRow.endingK2.toLocaleString()}</td>
          <td style="border: 0.5pt solid black; text-align: right; font-family: 'Times New Roman', serif; font-size: 8.5pt; font-weight: bold; padding: 0.5pt 2pt; vertical-align: middle;">${stickerTotalRow.endingTeams.toLocaleString()}</td>
          <td style="border: 0.5pt solid black; text-align: right; font-family: 'Times New Roman', serif; font-size: 8.5pt; font-weight: bold; padding: 0.5pt 2pt; vertical-align: middle;">${stickerTotalRow.grandTotalEnding.toLocaleString()}</td>
        </tr>
      </tbody>
    </table>
  </div>

  <!-- Section ខ. ក្រដាសអនុម័តទិដ្ឋាការអេឡិចត្រូនិក -->
  <div style="margin-top: 1pt; margin-bottom: 1pt; page-break-inside: avoid;">
    <div class="font-moul" style="font-size: 12pt; margin-bottom: 0.5pt;">
      ខ. ក្រដាសអនុម័តទិដ្ឋាការអេឡិចត្រូនិក
    </div>

    <table class="report-table" style="font-size: 8.5pt; text-align: center; line-height: 1.15;">
      <tbody>
        <tr style="background-color: #ffffff; font-weight: normal; font-size: 6pt; mso-height-rule: exactly; height: 0.52cm; mso-yfti-tblheader: yes; page-break-inside: avoid;">
          <th style="width: 25%; font-family: 'Khmer OS Siemreap', sans-serif; font-size: 6pt; font-weight: normal; padding: 0.5pt 1pt; border: 0.5pt solid black; vertical-align: middle;">
            សន្និធិ<br>ដើមគ្រា ${prevKhmerDate.day}-${prevKhmerDate.month}-${prevKhmerDate.year}
          </th>
          <th style="width: 25%; font-family: 'Khmer OS Siemreap', sans-serif; font-size: 6pt; font-weight: normal; padding: 0.5pt 1pt; border: 0.5pt solid black; vertical-align: middle;">
            បញ្ចូលស្តុក
          </th>
          <th style="width: 25%; font-family: 'Khmer OS Siemreap', sans-serif; font-size: 6pt; font-weight: normal; padding: 0.5pt 1pt; border: 0.5pt solid black; vertical-align: middle;">
            បើកផ្តល់ទៅតាមក្រុម<br>ផ្តល់ទិដ្ឋាការអេឡិចត្រូនិក
          </th>
          <th style="width: 25%; font-family: 'Khmer OS Siemreap', sans-serif; font-size: 6pt; font-weight: normal; padding: 0.5pt 1pt; border: 0.5pt solid black; vertical-align: middle;">
            សន្និធិនៅសល់<br>ត្រឹម ${endKhmerDate.day}-${endKhmerDate.month}-${endKhmerDate.year}
          </th>
        </tr>
        <tr style="mso-height-rule: exactly; height: 0.35cm; page-break-inside: avoid;">
          <td style="border: 0.5pt solid black; text-align: center; font-family: 'Times New Roman', serif; font-size: 9pt; padding: 0.5pt 2pt; vertical-align: middle;">${eVisaData.openBundles ? eVisaData.openBundles.toLocaleString() : '0'}</td>
          <td style="border: 0.5pt solid black; text-align: center; font-family: 'Times New Roman', serif; font-size: 9pt; padding: 0.5pt 2pt; vertical-align: middle;">${eVisaData.receivedK1Bundles ? eVisaData.receivedK1Bundles.toLocaleString() : '0'}</td>
          <td style="border: 0.5pt solid black; text-align: center; font-family: 'Times New Roman', serif; font-size: 9pt; padding: 0.5pt 2pt; vertical-align: middle;">${eVisaData.issuedToTeamsBundles ? eVisaData.issuedToTeamsBundles.toLocaleString() : '0'}</td>
          <td style="border: 0.5pt solid black; text-align: center; font-family: 'Times New Roman', serif; font-size: 9pt; font-weight: bold; padding: 0.5pt 2pt; vertical-align: middle;">${eVisaData.endingBundles ? eVisaData.endingBundles.toLocaleString() : '0'}</td>
        </tr>
      </tbody>
    </table>
  </div>
  `}

  <!-- Concluding Salutations -->
  <div style="font-size: 12pt; line-height: 1.25; margin-top: 1pt; margin-bottom: 1.5pt; page-break-inside: avoid;" class="font-siemreap">
    <div class="text-justify" style="text-indent: ${tabStopCm}cm; margin-bottom: 1pt;">
      ${variant === 'director'
        ? `អាស្រ័យដូចបានជម្រាបជូនខាងលើ សូម <span class="font-moul" style="font-size: 12pt;">${introSalutation || `${docRecipientRank || 'លោកឧត្តមសេនីយ៍ទោ'} ${docRecipientTitle || 'ប្រធាននាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត'}`}</span> មេត្តា ពិនិត្យ និងជ្រាបជា របាយការណ៍ ដ៏ខ្ពង់ខ្ពស់ ។`
        : proposalType === 'section'
        ? `អាស្រ័យដូចបានជម្រាបជូនខាងលើ សូម<span class="font-moul" style="font-size: 12pt;">លោកវរសេនីយ៍ទោ នាយការិយាល័យទិដ្ឋាការចូល</span> មេត្តា ពិនិត្យ និងសម្រចដោយក្តីអនុគ្រោះ ។`
        : `អាស្រ័យដូចបានជម្រាបជូនខាងលើ សូម <span class="font-moul" style="font-size: 12pt;">${introSalutation || 'លោកវរសេនីយ៍ទោ នាយការិយាល័យទិដ្ឋាការចូល'}</span> មេត្តា ពិនិត្យ និងជ្រាបជារបាយការណ៍ដោយក្តីអនុគ្រោះ ។`}
    </div>
    <div class="text-justify" style="text-indent: ${tabStopCm}cm;">
      សូម <span class="font-moul" style="font-size: 12pt;">${variant === 'director' ? (docRecipientRank || 'លោកឧត្តមសេនីយ៍ទោ') : 'លោកវរសេនីយ៍ទោ'}</span> មេត្តាទទួលនូវការគោរព ពីខ្ញុំបាទ ។
    </div>
  </div>

  <!-- Signatures Two-Column Layout -->
  <table class="layout-table" style="margin-top: 2pt; font-size: ${variant === 'director' ? '10pt' : '9pt'}; line-height: 1.25; page-break-inside: avoid;">
    <tr>
      <!-- Left Signature -->
      <td style="width: 50%; text-align: center; vertical-align: top; padding-left: 10pt; font-size: ${variant === 'director' ? '10pt' : '9pt'};">
        <div style="white-space: nowrap;">បានឃើញ និងគោរពជូន</div>
        <div style="white-space: nowrap;">${variant === 'director' ? `${docRecipientRank ? `${docRecipientRank} ` : 'លោកឧត្តមសេនីយ៍ទោ '}ប្រធាននាយកដ្ឋាន មេត្តាជ្រាប` : (docThroughTitle ? `${docThroughTitle} មេត្តាជ្រាប` : 'លោកនាយការិយាល័យទិដ្ឋាការចូល មេត្តាជ្រាប')}</div>
        <div style="white-space: nowrap;">${variant === 'director' ? 'ជារបាយការណ៍ ដ៏ខ្ពង់ខ្ពស់ ។' : 'ជារបាយការណ៍ ដោយក្តីអនុគ្រោះ។'}</div>
        ${signLunarDate ? `<div style="font-size: ${variant === 'director' ? '10pt' : '8.5pt'}; margin-top: 1pt; white-space: nowrap;">${signLunarDate}</div>` : ''}
        <div style="font-size: ${variant === 'director' ? '10pt' : '8.5pt'}; white-space: nowrap;">
          ${signSolarDate || `${variant === 'director' ? 'រាជធានីភ្នំពេញ' : 'ភ្នំពេញ'}, ថ្ងៃទី${signKhmerDate.day} ខែ${signKhmerDate.month} ឆ្នាំ${signKhmerDate.year}`}
        </div>
        <div class="font-moul" style="font-size: 12pt; margin-top: 2pt;">${leftSignTitle}</div>
        <div style="height: 38pt;"></div>
        ${leftSignOfficer ? `<div style="font-weight: bold; font-size: ${variant === 'director' ? '10pt' : '9.5pt'};">${leftSignOfficer}</div>` : ''}
      </td>

      <!-- Right Signature -->
      <td style="width: 50%; text-align: center; vertical-align: top; font-size: ${variant === 'director' ? '10pt' : '9pt'};">
        ${signLunarDate ? `<div style="font-size: ${variant === 'director' ? '10pt' : '8.5pt'}; white-space: nowrap;">${signLunarDate}</div>` : ''}
        <div style="font-size: ${variant === 'director' ? '10pt' : '8.5pt'}; white-space: nowrap;">
          ${signSolarDate || `${variant === 'director' ? 'រាជធានីភ្នំពេញ' : 'ភ្នំពេញ'}, ថ្ងៃទី${signKhmerDate.day} ខែ${signKhmerDate.month} ឆ្នាំ${signKhmerDate.year}`}
        </div>
        <div class="font-moul" style="font-size: 12pt; margin-top: 2pt;">${rightSignTitle}</div>
        <div style="height: 38pt;"></div>
        ${rightSignOfficer ? `<div style="font-weight: bold; font-size: ${variant === 'director' ? '10pt' : '9.5pt'};">${rightSignOfficer}</div>` : ''}
      </td>
    </tr>
  </table>
    `;
  }

  const wordHtml = `
<html xmlns:o='urn:schemas-microsoft-com:office:office'
      xmlns:w='urn:schemas-microsoft-com:office:word'
      xmlns='http://www.w3.org/TR/REC-html40'>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>${fileName}</title>
<!--[if gte mso 9]>
<xml>
  <w:WordDocument>
    <w:View>Print</w:View>
    <w:Zoom>100</w:Zoom>
    <w:DoNotOptimizeForBrowser/>
    <w:ValidateAgainstSchemas/>
    <w:SaveIfXMLInvalid>false</w:SaveIfXMLInvalid>
    <w:IgnoreMixedContent>false</w:IgnoreMixedContent>
    <w:AlwaysShowPlaceholderText>false</w:AlwaysShowPlaceholderText>
    <w:Compatibility>
      <w:BreakWrappedTables/>
      <w:SnapToGridInCell/>
      <w:WrapTextWithPunct/>
      <w:UseAsianBreakRules/>
      <w:DontGrowAutofit/>
    </w:Compatibility>
  </w:WordDocument>
</xml>
<![endif]-->
<style>
  @page WordSection1 {
    size: 210mm 297mm; /* Standard A4 Paper */
    margin: ${topCm}cm ${rightCm}cm ${bottomCm}cm ${leftCm}cm;
    mso-header-margin: 0cm;
    mso-footer-margin: 0cm;
    mso-paper-source: 0;
  }
  div.WordSection1 {
    page: WordSection1;
    font-family: 'Khmer OS Siemreap', 'Siemreap', 'Khmer OS', 'Segoe UI', Tahoma, sans-serif;
    font-size: 10pt;
    color: #000000;
    line-height: 1.25;
  }
  p, div, td, th {
    font-family: 'Khmer OS Siemreap', 'Siemreap', 'Khmer OS', 'Segoe UI', Tahoma, sans-serif;
    color: #000000;
    margin: 0;
    padding: 0;
  }
  .font-moul {
    font-family: 'Khmer Mool1', 'Khmer Mool 1', 'Khmer Mool', 'Khmer OS Muol Light', 'Khmer OS Muol', 'Khmer OS Mool1', 'Moul', 'Khmer OS', serif !important;
    font-weight: normal;
  }
  .font-siemreap {
    font-family: 'Khmer OS Siemreap', 'Siemreap', 'Khmer OS', 'Segoe UI', sans-serif !important;
  }
  .font-times {
    font-family: 'Times New Roman', 'Khmer OS Siemreap', serif !important;
  }
  table.report-table {
    border-collapse: collapse;
    mso-table-lspace: 0pt;
    mso-table-rspace: 0pt;
    width: 100%;
    margin-top: 1pt;
    margin-bottom: 2pt;
    page-break-inside: avoid;
  }
  table.report-table th, table.report-table td {
    border: 0.5pt solid #000000;
    padding: 0.5pt 1pt;
    vertical-align: middle;
  }
  table.layout-table {
    border-collapse: collapse;
    border: none;
    mso-table-lspace: 0pt;
    mso-table-rspace: 0pt;
    width: 100%;
    page-break-inside: avoid;
  }
  table.layout-table td {
    border: none;
    padding: 0;
    vertical-align: top;
  }
  .text-justify {
    text-align: justify;
    text-justify: inter-cluster;
  }
</style>
</head>
<body style="background-color: #ffffff; margin: 0; padding: 0;">
<div class="WordSection1">
  ${bodyContentHtml}
</div>
</body>
</html>
  `;

  // Create Word document Blob
  const blob = new Blob(['\ufeff' + wordHtml], {
    type: 'application/msword;charset=utf-8',
  });

  saveAs(blob, `${fileName}.doc`);
}

