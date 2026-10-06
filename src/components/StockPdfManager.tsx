import React, { useRef, useState, useEffect } from 'react';
import { CategoriesState, Officer, StockRecord } from '../types';
import { INITIAL_CATEGORIES } from '../data/initialData';
import { Download, X, FileText, Settings, Sparkles, SlidersHorizontal, Check, Plus, Edit3, Ruler } from 'lucide-react';
import html2canvas from 'html2canvas';
import { jsPDF } from 'jspdf';
import { saveAs } from 'file-saver';
import { exportElementToPdf } from '../utils/pdfExportHelper';
import { getKhmerLunarDate, getKhmerSolarParts, toKhmerNum } from '../utils/khmerCalendar';
import { TacteingLine, TacteingType, getSavedTacteingSettings } from './TacteingLink';

interface StockPdfModalProps {
  record: StockRecord;
  allRecords?: StockRecord[];
  categories?: CategoriesState;
  officers?: Officer[];
  userName?: string;
  onClose: () => void;
}

// Helper function to convert OKLAB / OKLCH color strings to standard RGB/RGBA strings for html2canvas
function oklabToRgb(L: number, aVal: number, bVal: number, alpha: number = 1): string {
  // OKLAB -> LMS
  const l_ = L + 0.3963377774 * aVal + 0.2158037573 * bVal;
  const m_ = L - 0.1055613458 * aVal - 0.0638541728 * bVal;
  const s_ = L - 0.0894841775 * aVal - 0.1291986507 * bVal;

  const l3 = l_ * l_ * l_;
  const m3 = m_ * m_ * m_;
  const s3 = s_ * s_ * s_;

  // LMS -> Linear RGB
  const rLin = +4.0767416621 * l3 - 3.3077115913 * m3 + 0.2309699292 * s3;
  const gLin = -1.2684380046 * l3 + 2.6097574011 * m3 - 0.3413193965 * s3;
  const bLin = -0.0041960863 * l3 - 0.7034186147 * m3 + 1.7076147010 * s3;

  // Gamma correction for sRGB
  const gamma = (c: number) => {
    const clamped = Math.max(0, Math.min(1, c));
    return clamped <= 0.0031308
      ? 12.92 * clamped
      : 1.055 * Math.pow(clamped, 1 / 2.4) - 0.055;
  };

  const r = Math.round(gamma(rLin) * 255);
  const g = Math.round(gamma(gLin) * 255);
  const b = Math.round(gamma(bLin) * 255);

  if (alpha < 1) {
    return `rgba(${r}, ${g}, ${b}, ${Number(alpha.toFixed(3))})`;
  }
  return `rgb(${r}, ${g}, ${b})`;
}

function convertModernCssColors(input: string): string {
  if (!input || typeof input !== 'string') return input;

  let result = input;

  // Convert oklch(...)
  if (result.includes('oklch')) {
    const oklchRegex = /oklch\(\s*([-\d.%]+)(?:[\s,]+)([-\d.%]+)(?:[\s,]+)([-\d.deg%radturn]+)(?:\s*[\/,]\s*([-\d.%]+))?\s*\)/gi;
    result = result.replace(oklchRegex, (fullMatch, lStr, cStr, hStr, aStr) => {
      try {
        let L = parseFloat(lStr);
        if (lStr.endsWith('%')) L = L / 100;

        let C = parseFloat(cStr);
        if (cStr.endsWith('%')) C = C / 100;

        let H = parseFloat(hStr.replace(/(deg|rad|turn)/gi, ''));
        if (hStr.toLowerCase().endsWith('rad')) H = (H * 180) / Math.PI;
        else if (hStr.toLowerCase().endsWith('turn')) H = H * 360;

        let alpha = 1;
        if (aStr !== undefined && aStr !== null) {
          alpha = parseFloat(aStr);
          if (aStr.endsWith('%')) alpha = alpha / 100;
        }

        if (isNaN(L) || isNaN(C) || isNaN(H)) return fullMatch;

        const hRad = (H * Math.PI) / 180;
        const aVal = C * Math.cos(hRad);
        const bVal = C * Math.sin(hRad);

        return oklabToRgb(L, aVal, bVal, alpha);
      } catch (e) {
        return fullMatch;
      }
    });
  }

  // Convert oklab(...)
  if (result.includes('oklab')) {
    const oklabRegex = /oklab\(\s*([-\d.%]+)(?:[\s,]+)([-\d.%]+)(?:[\s,]+)([-\d.%]+)(?:\s*[\/,]\s*([-\d.%]+))?\s*\)/gi;
    result = result.replace(oklabRegex, (fullMatch, lStr, aStr, bStr, alphaStr) => {
      try {
        let L = parseFloat(lStr);
        if (lStr.endsWith('%')) L = L / 100;

        let aVal = parseFloat(aStr);
        if (aStr.endsWith('%')) aVal = aVal / 100;

        let bVal = parseFloat(bStr);
        if (bStr.endsWith('%')) bVal = bVal / 100;

        let alpha = 1;
        if (alphaStr !== undefined && alphaStr !== null) {
          alpha = parseFloat(alphaStr);
          if (alphaStr.endsWith('%')) alpha = alpha / 100;
        }

        if (isNaN(L) || isNaN(aVal) || isNaN(bVal)) return fullMatch;

        return oklabToRgb(L, aVal, bVal, alpha);
      } catch (e) {
        return fullMatch;
      }
    });
  }

  return result;
}

export const formatOfficerGiverInfo = (officer: Officer, cat: CategoriesState) => {
  const rankObj = cat.ranks?.find((r) => r.id === officer.rankId);
  const posObj = cat.positions?.find((p) => p.id === officer.positionId);

  const rankName = rankObj ? rankObj.name : '';
  const posName = posObj ? posObj.name : '';

  let posText = '';
  if (posName) {
    posText = posName.startsWith('ជា') ? posName : `ជា${posName}`;
  }

  const parts = [];
  if (rankName) parts.push(rankName);
  if (officer.name) parts.push(officer.name);
  if (posText) parts.push(posText);

  const mainStr = parts.join(' ');
  return mainStr ? `${mainStr} នៃការិយាល័យទិដ្ឋាការចូល។` : officer.name;
};

export const formatOfficerReceiverInfo = (officer: Officer, cat: CategoriesState, teamName: string) => {
  const rankObj = cat.ranks?.find((r) => r.id === officer.rankId);
  const posObj = cat.positions?.find((p) => p.id === officer.positionId);

  const rankName = rankObj ? rankObj.name : 'វរៈសេនីយ៍ឯក';
  const posName = posObj ? posObj.name : 'ប្រធានក្រុមផ្តល់ទិដ្ឋាការ';

  const cleanRole = posName.replace(/^(មានតួនាទីជា|ជា)\s*/, '').replace(/[\s។]+$/, '');
  const rankStr = rankName ? `${rankName} ` : '';
  const nameStr = officer.name || 'រស់ លីហួត';
  const teamStr = teamName || 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអាកាសយានដ្ឋានអន្តរជាតិសៀមរាប';

  if (cleanRole.includes(teamStr)) {
    return `${rankStr}${nameStr} ជា${cleanRole}។`;
  }
  return `${rankStr}${nameStr} ជា${cleanRole} នៃ${teamStr}។`;
};

export const StockPdfModal: React.FC<StockPdfModalProps> = ({
  record,
  allRecords,
  categories,
  officers,
  userName = '',
  onClose,
}) => {
  const documentRef = useRef<HTMLDivElement>(null);
  const [isDownloading, setIsDownloading] = useState(false);
  const [downloadSuccess, setDownloadSuccess] = useState(false);
  const [downloadedBlobUrl, setDownloadedBlobUrl] = useState<string | null>(null);
  const [showEditPanel, setShowEditPanel] = useState(true);
  const [showColoredBoxesInPrint, setShowColoredBoxesInPrint] = useState(true);

  // Categories fallback
  const effectiveCategories: CategoriesState = categories || (() => {
    try {
      const saved = localStorage.getItem('app_categories');
      if (saved) return JSON.parse(saved);
    } catch {
      // fallback
    }
    return INITIAL_CATEGORIES;
  })();

  // Officers fallback
  const effectiveOfficers: Officer[] = officers || (() => {
    try {
      const saved = localStorage.getItem('app_officers');
      if (saved) return JSON.parse(saved);
    } catch {
      // fallback
    }
    return [];
  })();

  const collectorRolesList = effectiveCategories.collectorRoles || INITIAL_CATEGORIES.collectorRoles;

  // Extract officer rank + name for bottom signature display
  const getGiverDisplayName = () => {
    if (!giverFullInfo) return '';

    // 1. If colon exists (e.g., "មាន លាភ : នាយរងផ្នែក...")
    if (giverFullInfo.includes(':')) {
      return giverFullInfo.split(':')[0].trim();
    }

    // 2. Check if any officer in effectiveOfficers has their name in giverFullInfo
    const matchedOfficer = effectiveOfficers.find(
      (off) => off.name && giverFullInfo.includes(off.name)
    );
    if (matchedOfficer) {
      const rankObj = effectiveCategories.ranks?.find((r) => r.id === matchedOfficer.rankId);
      const rankName = rankObj ? rankObj.name : '';
      if (rankName && !matchedOfficer.name.includes(rankName)) {
        return `${rankName} ${matchedOfficer.name}`;
      }
      return matchedOfficer.name;
    }

    // 3. Extract rank + name from giverFullInfo if present before "ជា"
    const beforeJia = giverFullInfo.split(/\s+ជា/)[0].trim();
    if (beforeJia && beforeJia.length > 2 && beforeJia !== giverFullInfo) {
      return beforeJia;
    }

    // 4. Common sample names fallback
    if (giverFullInfo.includes('អ៊ុក រ័ត្នបញ្ញា') && !giverFullInfo.includes('អនុសេនីយ៍ឯក')) {
      return 'អនុសេនីយ៍ឯក អ៊ុក រ័ត្នបញ្ញា';
    }
    if (giverFullInfo.includes('ជ្រេង ថុល') && !giverFullInfo.includes('អនុសេនីយ៍ឯក')) {
      return 'អនុសេនីយ៍ឯក ជ្រេង ថុល';
    }

    return giverFullInfo;
  };

  const getReceiverDisplayName = () => {
    const name = receiverName || 'រស់ លីហួត';
    const rank = receiverRank || 'វរៈសេនីយ៍ឯក';
    if (rank && !name.includes(rank)) {
      return `${rank} ${name}`;
    }
    return name;
  };

  // Document Type State: 'evisa' (ក្រដាសអនុម័តទិដ្ឋាការអេឡិចត្រូនិក) or 'sticker' (សន្លឹកទិដ្ឋាការស្អិត)
  const [docStockType, setDocStockType] = useState<'evisa' | 'sticker'>(() =>
    record.stockType === 'evisa' ? 'evisa' : 'sticker'
  );
  const isEvisa = docStockType === 'evisa';

  const itemTerm = isEvisa
    ? 'ក្រដាសអនុម័តទិដ្ឋាការអេឡិចត្រូនិក'
    : 'សន្លឹកទិដ្ឋាការស្អិត';

  const docTitle = isEvisa
    ? 'លិខិតប្រគល់ទទួលក្រដាសអនុម័តទិដ្ឋាការអេឡិចត្រូនិក'
    : 'លិខិតប្រគល់ទទួលសន្លឹកទិដ្ឋាការស្អិត';

  // eVisa Bundles State (ចំនួនដុំ)
  const initialEvisaBundles = record.quantityBundles || record.totalSheets || 8;
  const [evisaBundles, setEvisaBundles] = useState<number | string>(initialEvisaBundles);

  // Solar Date & Parts in Khmer based on record date
  const initialSolarDate = record.date || new Date().toISOString().split('T')[0];
  const initialParts = getKhmerSolarParts(initialSolarDate);
  const [khmerDay, setKhmerDay] = useState(initialParts.khmerDay);
  const [khmerMonth, setKhmerMonth] = useState(initialParts.khmerMonth);
  const [khmerYear, setKhmerYear] = useState(initialParts.khmerYear);
  const [khmerTime, setKhmerTime] = useState(
    record.time ? toKhmerNum(record.time) : (record.stockType === 'evisa' ? '០៨:០០' : '១១:០០')
  );

  // TacTeing Line Settings State
  const savedTacteing = getSavedTacteingSettings();
  const [tacteingType, setTacteingType] = useState<TacteingType>(savedTacteing.type);
  const [tacteingCustomImage, setTacteingCustomImage] = useState<string | null>(savedTacteing.customImage);

  // Lunar Date State - Automatically calculated from initialSolarDate using momentkh
  const [lunarDate, setLunarDate] = useState<string>(() => getKhmerLunarDate(initialSolarDate));

  // Signature Horizontal Inset (px): 0px means fully shifted to left & right edges
  const [sigShiftInset, setSigShiftInset] = useState<number>(() => {
    try {
      const saved = localStorage.getItem('stock_pdf_sig_inset');
      if (saved !== null) {
        const val = parseInt(saved, 10);
        if (!isNaN(val)) return val;
      }
    } catch {}
    return 0;
  });

  // Vertical spacing between ភាគីទទួល and បានប្រគល់និងទទួល (px)
  const [evisaGapSpacing, setEvisaGapSpacing] = useState<number>(() => {
    try {
      const saved = localStorage.getItem('stock_pdf_evisa_gap');
      if (saved !== null) {
        const val = parseInt(saved, 10);
        if (!isNaN(val)) return val;
      }
    } catch {}
    return 6;
  });

  // Vertical spacing above signatures/dates ថ្ងៃខែ (px)
  const [sigTopSpacing, setSigTopSpacing] = useState<number>(() => {
    try {
      const saved = localStorage.getItem('stock_pdf_sig_top_spacing');
      if (saved !== null) {
        const val = parseInt(saved, 10);
        if (!isNaN(val)) return val;
      }
    } catch {}
    return 8;
  });

  // Document Body Font Size (pt) - Standard is 12pt for Khmer official documents
  const [bodyFontSizePt, setBodyFontSizePt] = useState<number>(() => {
    try {
      const saved = localStorage.getItem('stock_pdf_font_size_pt');
      if (saved !== null) {
        const val = parseFloat(saved);
        if (!isNaN(val) && val >= 8 && val <= 20) return val;
      }
    } catch {}
    return 12;
  });

  // Document Header Font Size (pt) - Standard 12pt as requested
  const [headerFontSizePt, setHeaderFontSizePt] = useState<number>(() => {
    try {
      const saved = localStorage.getItem('stock_pdf_header_font_size_pt');
      if (saved !== null) {
        const val = parseFloat(saved);
        if (!isNaN(val) && val >= 8 && val <= 20) return val;
      }
    } catch {}
    return 12;
  });

  // Document Title Font Size (pt) - Standard 12pt as requested
  const [titleFontSizePt, setTitleFontSizePt] = useState<number>(() => {
    try {
      const saved = localStorage.getItem('stock_pdf_title_font_size_pt');
      if (saved !== null) {
        const val = parseFloat(saved);
        if (!isNaN(val) && val >= 8 && val <= 24) return val;
      }
    } catch {}
    return 12;
  });

  // Document Page Margins (cm) - defaults matching របក.សរុបការងារស្តុក: left=2.8cm, right=1.3cm, top=1cm, bottom=0.5cm
  const [docMarginLeftCm, setDocMarginLeftCm] = useState<number>(() => {
    try {
      const saved = localStorage.getItem('stock_pdf_margin_left_cm');
      if (saved !== null) {
        const val = parseFloat(saved);
        if (!isNaN(val) && val >= 0) return val;
      }
    } catch {}
    return 2.8;
  });

  const [docMarginRightCm, setDocMarginRightCm] = useState<number>(() => {
    try {
      const saved = localStorage.getItem('stock_pdf_margin_right_cm');
      if (saved !== null) {
        const val = parseFloat(saved);
        if (!isNaN(val) && val >= 0) return val;
      }
    } catch {}
    return 1.3;
  });

  const [docMarginTopCm, setDocMarginTopCm] = useState<number>(() => {
    try {
      const saved = localStorage.getItem('stock_pdf_margin_top_cm');
      if (saved !== null) {
        const val = parseFloat(saved);
        if (!isNaN(val) && val >= 0) return val;
      }
    } catch {}
    return 1;
  });

  const [docMarginBottomCm, setDocMarginBottomCm] = useState<number>(() => {
    try {
      const saved = localStorage.getItem('stock_pdf_margin_bottom_cm');
      if (saved !== null) {
        const val = parseFloat(saved);
        if (!isNaN(val) && val >= 0) return val;
      }
    } catch {}
    return 0.5;
  });

  // Active Tab Stop (cm) - defaults to 2.25cm matching របក.សរុបការងារស្តុក
  const [activeTabStop, setActiveTabStop] = useState<number>(() => {
    try {
      const saved = localStorage.getItem('stock_pdf_active_tab_stop');
      if (saved !== null) {
        const val = parseFloat(saved);
        if (!isNaN(val) && val >= 0) return val;
      }
    } catch {}
    return 2.25;
  });

  const [showRuler, setShowRuler] = useState<boolean>(true);
  const [isWordEditMode, setIsWordEditMode] = useState<boolean>(false);

  // 4 Editable Highlighted Points (as requested by user with colored boxes)
  // 1. Red Box Point: Ref Role Title (defaults to empty so "-- ជ្រើសរើសអ្នកស្នើសុំ --" is selected)
  const initialRefRoleTitle = '';

  const [refRoleTitle, setRefRoleTitle] = useState(initialRefRoleTitle);

  // 2. Yellow Box Point 1: Giver Full Info (from Officer Data)
  const sampleGiverInfo = 'អនុសេនីយ៍ឯក អ៊ុក រ័ត្នបញ្ញា ជានាយផ្នែក នៃការិយាល័យទិដ្ឋាការចូល។';
  const initialGiverFullInfo = '';

  const [giverFullInfo, setGiverFullInfo] = useState(initialGiverFullInfo);

  // Helper to find full team name from categories (matching row index in visaTeamsData / visaTeamsRobok)
  const getFullVisaTeamName = () => {
    const vtrList = effectiveCategories.visaTeamsRobok || [];
    const vtList = effectiveCategories.visaTeams || [];

    // 1. Try matching by visaTeamRobokId
    let matchedIdx = -1;
    if (record.visaTeamRobokId) {
      matchedIdx = vtrList.findIndex((item) => item.id === record.visaTeamRobokId);
    }

    // 2. Try matching by visaTeamRobokName text
    if (matchedIdx === -1 && record.visaTeamRobokName) {
      matchedIdx = vtrList.findIndex(
        (item) => item.name.trim().toLowerCase() === record.visaTeamRobokName?.trim().toLowerCase()
      );
    }

    // 3. If row matched in visaTeamsRobok, grab full team name from same row index in visaTeams
    if (matchedIdx !== -1 && vtList[matchedIdx]?.name) {
      return vtList[matchedIdx].name;
    }

    // 4. Try matching record.visaTeamRobokName directly in visaTeams
    if (record.visaTeamRobokName) {
      const directMatch = vtList.find(
        (item) => item.name.trim().toLowerCase() === record.visaTeamRobokName?.trim().toLowerCase()
      );
      if (directMatch) return directMatch.name;
    }

    // 5. Fallback
    return vtList[0]?.name || record.visaTeamRobokName || 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំទ្វារអាកាសយានដ្ឋានអន្តរជាតិភ្នំពេញ';
  };

  // Additional editable fields
  const [teamName, setTeamName] = useState(getFullVisaTeamName);

  // 3. Yellow Box Point 2: Receiver Full Info (Rank, Name, Role)
  const collectorNameRaw = record.collectorName || 'រស់ លីហួត';
  const initialReceiverRank = record.requestedRankName || effectiveCategories.ranks?.[0]?.name || 'វរៈសេនីយ៍ឯក';
  const initialReceiverName = collectorNameRaw.includes(' ') ? collectorNameRaw : collectorNameRaw;
  const initialReceiverRole = (record.collectorRoleName || 'ប្រធានក្រុមផ្តល់ទិដ្ឋាការ').trim().replace(/^(មានតួនាទីជា|ជា)\s*/, '').replace(/[\s។]+$/, '');

  const [receiverRank, setReceiverRank] = useState(initialReceiverRank);
  const [receiverName, setReceiverName] = useState(initialReceiverName);
  const [receiverRole, setReceiverRole] = useState(initialReceiverRole);

  const formatReceiverFullInfo = (rk: string, nm: string, rl: string, tm: string) => {
    const rankStr = rk ? `${rk} ` : '';
    const nameStr = nm || 'រស់ លីហួត';
    const roleStr = (rl || 'ប្រធានក្រុមផ្តល់ទិដ្ឋាការ').trim().replace(/^(មានតួនាទីជា|ជា)\s*/, '').replace(/[\s។]+$/, '');
    const teamStr = tm || 'ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអាកាសយានដ្ឋានអន្តរជាតិសៀមរាប';

    if (roleStr.includes(teamStr)) {
      return `${rankStr}${nameStr} ជា${roleStr}។`;
    }
    return `${rankStr}${nameStr} ជា${roleStr} នៃ${teamStr}។`;
  };

  const [receiverFullInfo, setReceiverFullInfo] = useState(() =>
    formatReceiverFullInfo(initialReceiverRank, initialReceiverName, initialReceiverRole, getFullVisaTeamName())
  );

  // 4. Central Quantity & Table Data States (Supports multiple visa types - 1 ក្បាល = 50 សន្លឹក)
  const calculateInitialQuantitiesForRecord = (rec: StockRecord) => {
    const sNum = parseInt(rec.startSerial || '', 10);
    const eNum = parseInt(rec.endSerial || '', 10);
    if (!isNaN(sNum) && !isNaN(eNum) && eNum >= sNum) {
      const sheets = eNum - sNum + 1;
      const kbal = Math.ceil(sheets / 50);
      return { sheets, kbal };
    }
    const qty = rec.quantityBundles || 10;
    if (qty > 100) {
      return { sheets: qty, kbal: Math.ceil(qty / 50) };
    } else {
      return { sheets: qty * 50, kbal: qty };
    }
  };

  const getAutoGroupedItems = () => {
    if (allRecords && allRecords.length > 0) {
      const teamA = (record.visaTeamRobokName || record.visaTeamRobokId || '').trim().toLowerCase();
      const matching = allRecords.filter((r) => {
        if (r.stockType !== record.stockType) return false;
        if (r.operationType !== record.operationType) return false;
        if (r.date !== record.date) return false;

        const teamB = (r.visaTeamRobokName || r.visaTeamRobokId || '').trim().toLowerCase();
        if (teamA && teamB && teamA !== teamB) return false;

        return true;
      });

      if (matching.length > 0) {
        return matching.map((r, idx) => {
          const init = calculateInitialQuantitiesForRecord(r);
          return {
            id: r.id || String(idx + 1),
            visaType: r.visaType || 'T',
            kbalCount: init.kbal,
            quantitySheets: init.sheets,
            startSerial: r.startSerial || '',
            endSerial: r.endSerial || '',
          };
        });
      }
    }

    const init = calculateInitialQuantitiesForRecord(record);
    return [
      {
        id: record.id || '1',
        visaType: record.visaType || 'T',
        kbalCount: init.kbal,
        quantitySheets: init.sheets,
        startSerial: record.startSerial || '',
        endSerial: record.endSerial || '',
      },
    ];
  };

  const [tableItems, setTableItems] = useState(getAutoGroupedItems);

  useEffect(() => {
    setTableItems(getAutoGroupedItems());
  }, [record, allRecords]);

  const formatTableNum = (val: number | string) => String(val);

  const handleRowChange = (id: string, field: string, value: any) => {
    setTableItems((prev) =>
      prev.map((item) => {
        if (item.id !== id) return item;
        const updated = { ...item, [field]: value };

        if (field === 'kbalCount') {
          const kbal = parseInt(value, 10) || 0;
          updated.kbalCount = kbal;
          updated.quantitySheets = kbal * 50;
          const sNum = parseInt(updated.startSerial, 10);
          if (!isNaN(sNum) && updated.quantitySheets > 0) {
            updated.endSerial = String(sNum + updated.quantitySheets - 1);
          }
        } else if (field === 'quantitySheets') {
          const sheets = parseInt(value, 10) || 0;
          updated.quantitySheets = sheets;
          updated.kbalCount = Math.ceil(sheets / 50);
          const sNum = parseInt(updated.startSerial, 10);
          if (!isNaN(sNum) && sheets > 0) {
            updated.endSerial = String(sNum + sheets - 1);
          }
        } else if (field === 'startSerial') {
          updated.startSerial = value;
          const sNum = parseInt(value, 10);
          const eNum = parseInt(updated.endSerial, 10);
          if (!isNaN(sNum) && !isNaN(eNum) && eNum >= sNum) {
            updated.quantitySheets = eNum - sNum + 1;
            updated.kbalCount = Math.ceil(updated.quantitySheets / 50);
          }
        } else if (field === 'endSerial') {
          updated.endSerial = value;
          const sNum = parseInt(updated.startSerial, 10);
          const eNum = parseInt(value, 10);
          if (!isNaN(sNum) && !isNaN(eNum) && eNum >= sNum) {
            updated.quantitySheets = eNum - sNum + 1;
            updated.kbalCount = Math.ceil(updated.quantitySheets / 50);
          }
        }

        return updated;
      })
    );
  };

  const handleAddRow = () => {
    setTableItems((prev) => {
      const last = prev[prev.length - 1];
      let nextStart = '1903849001';
      if (last && last.endSerial) {
        const eNum = parseInt(last.endSerial, 10);
        if (!isNaN(eNum)) {
          nextStart = String(eNum + 1);
        }
      }
      const nextEnd = String(parseInt(nextStart, 10) + 499); // default 500 sheets = 10 kbal
      const defaultType = prev.some((i) => i.visaType === 'T') ? (prev.some((i) => i.visaType === 'E') ? 'K' : 'E') : 'T';

      return [
        ...prev,
        {
          id: String(Date.now()),
          visaType: defaultType,
          kbalCount: 10,
          quantitySheets: 500,
          startSerial: nextStart,
          endSerial: nextEnd,
        },
      ];
    });
  };

  const handleRemoveRow = (id: string) => {
    setTableItems((prev) => (prev.length > 1 ? prev.filter((i) => i.id !== id) : prev));
  };
  const initialRequester = record.requesterName
    ? (record.requestedRankName && !record.requesterName.includes(record.requestedRankName)
        ? `${record.requestedRankName} ${record.requesterName}`
        : record.requesterName)
    : 'វរសេនីយ៍ទោ ផែន វិបុល';
  const [requesterName, setRequesterName] = useState(initialRequester);
  const [approverTitle, setApproverTitle] = useState(
    record.stockType === 'evisa' ? 'ភប.នាយការិយាល័យទិដ្ឋាការចូល' : 'តប.នាយការិយាល័យទិដ្ឋាការចូល'
  );
  const [approverName, setApproverName] = useState('អនុសេនីយ៍ឯក អ៊ុក រ័ត្នបញ្ញា');

  const getReferenceFullText = () => {
    const req = requesterName || 'វរសេនីយ៍ទោ ផែន វិបុល';
    const cleanRole = (refRoleTitle || '').trim().replace(/^ជា\s*/, '');
    const rolePart = cleanRole ? ` ជា${cleanRole}` : '';
    const cleanTeam = (teamName || '').trim().replace(/^នៃ\s*/, '');
    const teamPart = cleanTeam ? ` នៃ${cleanTeam}` : '';
    return `សំណើរបស់ ${req}${rolePart}${teamPart}។`;
  };

  const getOfficerDisplayName = (off: Officer) => {
    const rankObj = effectiveCategories.ranks?.find((r) => r.id === off.rankId);
    const rankName = rankObj ? rankObj.name : '';
    if (rankName && off.name.startsWith(rankName)) {
      return off.name;
    }
    return rankName ? `${rankName} ${off.name}` : off.name;
  };

  const getTargetPositionNames = (title: string): string[] => {
    if (title.includes('នាយកដ្ឋាន')) {
      return ['ប្រធាននាយកដ្ឋាន', 'អនុប្រធាននាយកដ្ឋាន', 'នាយការិយាល័យ', 'នាយរងការិយាល័យ'];
    }
    if (title === 'តប.នាយការិយាល័យទិដ្ឋាការចូល' || title.includes('នាយការិយាល័យ') || title.includes('ការិយាល័យ')) {
      return ['នាយរងការិយាល័យ', 'អនុប្រធានការិយាល័យ', 'នាយការិយាល័យ', 'ប្រធានការិយាល័យ'];
    }
    if (title === 'ជ.នាយផ្នែក' || title === 'នាយរងផ្នែក') {
      return ['នាយរងផ្នែក'];
    }
    if (title === 'នាយផ្នែក') {
      return ['នាយផ្នែក'];
    }
    return [title];
  };

  const getMatchingOfficersForTitle = (title: string) => {
    if (!effectiveOfficers || effectiveOfficers.length === 0) return [];
    const targetPositions = getTargetPositionNames(title);
    if (!targetPositions || targetPositions.length === 0) return [];

    return effectiveOfficers.filter((off) => {
      const posObj = effectiveCategories.positions?.find((p) => p.id === off.positionId);
      const posName = posObj ? posObj.name : '';
      const posClean = posName.replace(/^ជា/, '').trim().toLowerCase();
      return targetPositions.some((target) => {
        const targetClean = target.replace(/^ជា/, '').trim().toLowerCase();
        return posClean.includes(targetClean) || targetClean.includes(posClean);
      });
    });
  };

  const matchingOfficers = getMatchingOfficersForTitle(approverTitle);
  const approverOfficersToDisplay = matchingOfficers.length > 0 ? matchingOfficers : effectiveOfficers;

  const giverFilteredOfficers = effectiveOfficers.filter((off) => {
    const posObj = effectiveCategories.positions?.find((p) => p.id === off.positionId);
    const posName = posObj ? posObj.name : '';
    const cleanPos = posName.replace(/^ជា/, '').trim().toLowerCase();
    const allowed = ['នាយផ្នែក', 'នាយរងផ្នែក', 'មន្ត្រី'];
    return allowed.some((target) => cleanPos.includes(target.toLowerCase()));
  });

  const giverOfficersToDisplay = giverFilteredOfficers.length > 0 ? giverFilteredOfficers : effectiveOfficers;

  useEffect(() => {
    if (matchingOfficers.length > 0) {
      const isCurrentInMatched = matchingOfficers.some((off) => getOfficerDisplayName(off) === approverName);
      if (!isCurrentInMatched) {
        setApproverName(getOfficerDisplayName(matchingOfficers[0]));
      }
    }
  }, [effectiveOfficers, effectiveCategories, approverTitle]);

  const handleTitleChange = (newTitle: string) => {
    setApproverTitle(newTitle);
    const matched = getMatchingOfficersForTitle(newTitle);
    if (matched.length > 0) {
      setApproverName(getOfficerDisplayName(matched[0]));
    }
  };

  // Helper to render text with officer names in Khmer Mool1 font and ranks (ឋានន្តរស័ក្កិ) in Khmer OS Siemreap font
  const renderTextWithMoulName = (text: string, defaultClassName?: string) => {
    if (!text) return null;

    const baseRanks = [
      'នាយឧត្តមសេនីយ៍',
      'ឧត្តមសេនីយ៍ឯក',
      'ឧត្តមសេនីយ៍ទោ',
      'ឧត្តមសេនីយ៍ត្រី',
      'វរៈសេនីយ៍ឯក',
      'វរេសេនីយ៍ឯក',
      'វរៈសេនីយ៍ទោ',
      'វរេសេនីយ៍ទោ',
      'វរៈសេនីយ៍ត្រី',
      'វរេសេនីយ៍ត្រី',
      'អនុសេនីយ៍ឯក',
      'អនុសេនីយ៍ទោ',
      'អនុសេនីយ៍ត្រី',
      'ព្រិន្ទបាលឯក',
      'ព្រិន្ទបាលទោ',
      'ព្រិន្ទបាលត្រី',
      'នាយចំណង់',
      'លោក',
      'លោកស្រី',
    ];

    if (effectiveCategories?.ranks) {
      effectiveCategories.ranks.forEach((r) => {
        if (r.name && r.name.trim()) baseRanks.push(r.name.trim());
      });
    }

    const ranksArray = Array.from(new Set(baseRanks))
      .filter((r) => r.length > 1)
      .sort((a, b) => b.length - a.length);

    const stripRanks = (str: string) => {
      let s = str || '';
      ranksArray.forEach((rk) => {
        s = s.replace(new RegExp(rk, 'g'), '');
      });
      return s.trim();
    };

    const namesSet = new Set<string>();
    if (receiverName) namesSet.add(stripRanks(receiverName));
    if (requesterName) namesSet.add(stripRanks(requesterName));
    if (approverName) namesSet.add(stripRanks(approverName));
    if (record.collectorName) namesSet.add(stripRanks(record.collectorName));

    effectiveOfficers.forEach((off) => {
      if (off.name) {
        namesSet.add(stripRanks(off.name));
      }
    });

    ['អ៊ុក រ័ត្នបញ្ញា', 'រស់ លីហួត', 'ដែន និមល', 'ជ្រេង ថុល', 'លី រតនៈ'].forEach((n) => namesSet.add(n));

    const namesArray = Array.from(namesSet)
      .filter((n) => n && n.length > 1)
      .sort((a, b) => b.length - a.length);

    const allTokens = Array.from(new Set([...namesArray, ...ranksArray])).sort((a, b) => b.length - a.length);

    if (allTokens.length === 0) return <span className={defaultClassName}>{text}</span>;

    const escapeRegExp = (s: string) => s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
    const pattern = new RegExp(`(${allTokens.map(escapeRegExp).join('|')})`, 'g');

    const parts = text.split(pattern);
    return (
      <span className={defaultClassName}>
        {parts.map((part, idx) => {
          if (namesArray.includes(part)) {
            return (
              <span key={idx} className="font-moul font-normal" style={{ fontSize: `${bodyFontSizePt}pt` }}>
                {part}
              </span>
            );
          }
          if (ranksArray.includes(part)) {
            return (
              <span key={idx} className="font-siemreap font-normal" style={{ fontSize: `${bodyFontSizePt}pt` }}>
                {part}
              </span>
            );
          }
          return <span key={idx}>{part}</span>;
        })}
      </span>
    );
  };

  // Direct PDF Download handle using html2canvas & jsPDF
  const handleDownloadPdf = async () => {
    // 1. Check if all required selections in colored boxes are chosen
    const missingFields: string[] = [];
    if (!refRoleTitle) missingFields.push('១. ប្រអប់ក្រហម (តួនាទីស្នើសុំ)');
    if (!giverFullInfo) missingFields.push('២. ប្រអប់លឿង (ភាគីប្រគល់)');
    if (!receiverRank) missingFields.push('៣. ប្រអប់លឿង (ភាគីទទួល)');

    if (missingFields.length > 0) {
      setShowEditPanel(true);
      alert(
        `សូមជ្រើសរើសព័ត៌មានក្នុងប្រអប់ខាងក្រោមឱ្យបានគ្រប់ជ្រុងជ្រោយជាមុនសិន មុននឹងទាញយក PDF:\n- ${missingFields.join('\n- ')}`
      );
      return;
    }

    // 2. Check if "បង្ហាញប្រអប់ពណ៌" is still checked (user needs to unclick it before pulling PDF)
    if (showColoredBoxesInPrint) {
      setShowEditPanel(true);
      alert('សូមដោះគ្រីច "បង្ហាញប្រអប់ពណ៌" (Show Colored Boxes) ជាមុនសិន មុននឹងទាញយក PDF!');
      return;
    }

    if (!documentRef.current) return;
    try {
      setIsDownloading(true);

      // Ensure all custom fonts (Khmer fonts) are loaded before rendering to canvas
      if (document.fonts && document.fonts.ready) {
        await document.fonts.ready;
      }
      await new Promise((r) => setTimeout(r, 200));

      const element = documentRef.current;
      const teamText = record.visaTeamRobokName || teamName || 'ក្រុមផ្តល់ទិដ្ឋាការ';
      const cleanTeam = teamText.replace(/[/\\?%*:|"<>]/g, '-').trim();
      const dateText = record.date || initialSolarDate || 'ថ្ងៃខែបើកផ្តល់';
      const cleanDate = dateText.replace(/[/\\?%*:|"<>]/g, '-').trim();
      const docPrefix = isEvisa ? 'លិខិតប្រគល់ទទួលក្រដាសអនុម័ត' : 'លិខិតប្រគល់ទទួលសន្លឹកទិដ្ឋាការ';
      const fileName = `${docPrefix}(${cleanTeam}_${cleanDate}).pdf`;

      // 1. Primary: Use browser-native SVG rendering (html-to-image) for mathematically exact layout and cell centering
      try {
        const blobUrl = await exportElementToPdf(element, fileName, {
          pixelRatio: 4,
          orientation: 'portrait',
          fitSinglePage: true,
        });
        setDownloadedBlobUrl(blobUrl);
        setDownloadSuccess(true);
        return;
      } catch (nativeErr) {
        console.warn('Native PDF export failed, using html2canvas fallback:', nativeErr);
      }

      // 2. Fallback: html2canvas with full table cell middle-alignment sanitization
      const canvas = await html2canvas(element, {
        scale: 4, // Ultra high-resolution scale for maximum clarity of Khmer text & borders
        useCORS: true,
        allowTaint: true,
        logging: false,
        backgroundColor: '#ffffff',
        windowWidth: 1200,
        windowHeight: 1600,
        onclone: (clonedDoc) => {
          // Ensure font-smoothing is enabled on cloned document body
          if (clonedDoc.body) {
            clonedDoc.body.style.setProperty('-webkit-font-smoothing', 'antialiased');
            clonedDoc.body.style.textRendering = 'geometricPrecision';
          }

          const win = clonedDoc.defaultView || window;
          const colorProps = [
            'color',
            'background-color',
            'border-color',
            'border-top-color',
            'border-right-color',
            'border-bottom-color',
            'border-left-color',
            'outline-color',
            'fill',
            'stroke',
          ];

          // Replace oklch/oklab in all <style> tags with accurate RGB equivalents
          const styleTags = Array.from(clonedDoc.getElementsByTagName('style'));
          styleTags.forEach((style) => {
            if (style.textContent && (style.textContent.includes('oklch') || style.textContent.includes('oklab'))) {
              style.textContent = convertModernCssColors(style.textContent);
            }
          });

          // Replace oklch/oklab in all element inline styles and computed styles
          const allElements = Array.from(clonedDoc.querySelectorAll('*'));
          allElements.forEach((el) => {
            const htmlEl = el as HTMLElement;
            const styleAttr = htmlEl.getAttribute('style');
            if (styleAttr && (styleAttr.includes('oklch') || styleAttr.includes('oklab'))) {
              htmlEl.setAttribute('style', convertModernCssColors(styleAttr));
            }

            try {
              const computed = win.getComputedStyle(htmlEl);
              colorProps.forEach((prop) => {
                const val = computed.getPropertyValue(prop);
                if (val && (val.includes('oklab') || val.includes('oklch'))) {
                  const converted = convertModernCssColors(val);
                  htmlEl.style.setProperty(prop, converted, 'important');
                }
              });
            } catch (e) {
              // ignore computed style read errors
            }
          });

          // Sanitize table and ensure true middle vertical alignment and horizontal centering matching the live table exactly
          const clonedEl = clonedDoc.getElementById('official-pdf-document');
          const tables = clonedEl ? Array.from(clonedEl.querySelectorAll('table')) : Array.from(clonedDoc.querySelectorAll('table'));
          tables.forEach((table) => {
            // Prevent collapsed border overlapping glitches in html2canvas
            table.style.setProperty('border-collapse', 'separate', 'important');
            table.style.setProperty('border-spacing', '0px', 'important');
            table.style.setProperty('border-top', '1px solid #000000', 'important');
            table.style.setProperty('border-left', '1px solid #000000', 'important');
            table.style.setProperty('border-right', 'none', 'important');
            table.style.setProperty('border-bottom', 'none', 'important');

            const rows = Array.from(table.querySelectorAll('tr'));
            rows.forEach((tr) => {
              (tr as HTMLElement).style.setProperty('vertical-align', 'middle', 'important');
              (tr as HTMLElement).style.setProperty('height', 'auto', 'important');
            });

            const cells = Array.from(table.querySelectorAll('th, td'));
            cells.forEach((c) => {
              const cell = c as HTMLElement;
              cell.style.setProperty('border-right', '1px solid #000000', 'important');
              cell.style.setProperty('border-bottom', '1px solid #000000', 'important');
              cell.style.setProperty('border-top', 'none', 'important');
              cell.style.setProperty('border-left', 'none', 'important');
              cell.style.setProperty('box-sizing', 'border-box', 'important');
              cell.style.setProperty('background-color', '#ffffff', 'important');
              cell.style.setProperty('overflow', 'visible', 'important');
              cell.style.setProperty('vertical-align', 'middle', 'important');
              cell.style.setProperty('text-align', 'center', 'important');
              cell.style.setProperty('padding-top', '5px', 'important');
              cell.style.setProperty('padding-bottom', '5px', 'important');
              cell.style.setProperty('line-height', '1.2', 'important');
            });
          });

          if (clonedEl) {
            let parent = clonedEl.parentElement;
            while (parent && parent !== clonedDoc.body) {
              parent.style.width = 'auto';
              parent.style.maxWidth = 'none';
              parent.style.overflow = 'visible';
              parent = parent.parentElement;
            }

            clonedEl.style.width = '210mm';
            clonedEl.style.minWidth = '210mm';
            clonedEl.style.maxWidth = '210mm';
            clonedEl.style.transform = 'none';
            clonedEl.style.margin = '0 auto';
            clonedEl.style.boxShadow = 'none';
            clonedEl.style.backgroundColor = '#ffffff';
          }
        },
      });

      // Use PNG format for crisp, lossless text and vector rendering without JPEG artifacts
      const imgData = canvas.toDataURL('image/png', 1.0);
      const pdf = new jsPDF({
        orientation: 'portrait',
        unit: 'mm',
        format: 'a4',
        compress: true,
      });

      const imgWidth = 210; // A4 width mm
      const pageHeight = 297; // A4 height mm
      const imgHeight = (canvas.height * imgWidth) / canvas.width;

      // High-quality bicubic rendering without 'FAST' downsampling artifacts
      pdf.addImage(imgData, 'PNG', 0, 0, imgWidth, Math.min(imgHeight, pageHeight), undefined, 'NONE');

      // Trigger direct PDF download with specified filename
      pdf.save(fileName);

      const pdfBlob = pdf.output('blob');
      const blobUrl = URL.createObjectURL(pdfBlob);

      setDownloadedBlobUrl(blobUrl);
      setDownloadSuccess(true);
    } catch (err) {
      console.error('PDF generation error:', err);
      alert('មានបញ្ហាក្នុងការទាញយក PDF។ សូមព្យាយាមម្តងទៀត។');
    } finally {
      setIsDownloading(false);
    }
  };

  return (
    <div className="fixed inset-0 bg-black/70 z-50 flex items-center justify-center p-2 sm:p-4 overflow-y-auto print:p-0 print:bg-white print:static print:block">
      {/* Container */}
      <div className="bg-white rounded-lg shadow-2xl w-full max-w-5xl max-h-[96vh] flex flex-col overflow-hidden border border-gray-300 print:shadow-none print:border-none print:max-h-none print:w-full">
        {/* Modal Top Bar */}
        <div className="bg-[#007bff] text-white px-4 py-3 flex flex-wrap items-center justify-between gap-2 shrink-0 print:hidden">
          <div className="flex items-center gap-3 font-bold text-sm">
            <div className="flex items-center gap-1.5">
              <FileText className="w-5 h-5 text-white" />
              <span>លិខិតប្រគល់ទទួល (PDF — ខ្នាត A4)</span>
            </div>

            {/* Document Format Switcher */}
            <div className="flex items-center bg-blue-900/50 rounded p-0.5 border border-blue-300/40 text-xs">
              <button
                type="button"
                onClick={() => setDocStockType('evisa')}
                className={`px-2.5 py-1 rounded transition cursor-pointer font-bold ${
                  isEvisa ? 'bg-white text-blue-900 shadow-xs' : 'text-blue-100 hover:text-white'
                }`}
                title="ទម្រង់លិខិតប្រគល់ទទួលក្រដាសអនុម័តទិដ្ឋាការអេឡិចត្រូនិក"
              >
                ក្រដាសអនុម័ត (eVisa)
              </button>
              <button
                type="button"
                onClick={() => setDocStockType('sticker')}
                className={`px-2.5 py-1 rounded transition cursor-pointer font-bold ${
                  !isEvisa ? 'bg-white text-blue-900 shadow-xs' : 'text-blue-100 hover:text-white'
                }`}
                title="ទម្រង់លិខិតប្រគល់ទទួលសន្លឹកទិដ្ឋាការស្អិត"
              >
                សន្លឹកទិដ្ឋាការ (Sticker)
              </button>
            </div>
          </div>

          <div className="flex items-center gap-2">
            {/* Word Direct Edit Mode Toggle */}
            <button
              type="button"
              onClick={() => setIsWordEditMode(!isWordEditMode)}
              className={`px-3 py-1.5 rounded text-xs font-bold transition flex items-center gap-1.5 cursor-pointer shadow-2xs ${
                isWordEditMode
                  ? 'bg-amber-400 text-gray-950 ring-2 ring-amber-300 font-extrabold'
                  : 'bg-white/20 hover:bg-white/30 text-white'
              }`}
              title="បើក/បិទ ការកែសម្រួលអត្ថបទផ្ទាល់លើក្រដាស (Word Mode)"
            >
              <Edit3 className="w-4 h-4" />
              <span>{isWordEditMode ? 'Word Edit: បើក' : 'Word Mode'}</span>
            </button>

            {/* Ruler Toggle */}
            <button
              type="button"
              onClick={() => setShowRuler(!showRuler)}
              className={`px-2.5 py-1.5 rounded text-xs font-bold transition flex items-center gap-1.5 cursor-pointer shadow-2xs ${
                showRuler
                  ? 'bg-blue-800 text-white border border-blue-300/40'
                  : 'bg-white/20 hover:bg-white/30 text-white'
              }`}
              title="បង្ហាញ/លាក់ បន្ទាត់រង្វាស់ខ្នាត A4"
            >
              <Ruler className="w-4 h-4" />
              <span className="hidden sm:inline">{showRuler ? 'Ruler' : 'លាក់ Ruler'}</span>
            </button>

            <button
              onClick={() => setShowEditPanel(!showEditPanel)}
              className={`px-3 py-1.5 rounded text-xs font-bold transition flex items-center gap-1.5 cursor-pointer shadow-2xs ${
                showEditPanel
                  ? 'bg-amber-400 text-gray-900'
                  : 'bg-white/20 hover:bg-white/30 text-white'
              }`}
            >
              <SlidersHorizontal className="w-4 h-4" />
              <span>{showEditPanel ? 'លាក់ផ្ទាំងកែប្រែប្រអប់' : 'កែប្រែខ្លឹមសារប្រអប់ពណ៌'}</span>
            </button>

            <button
              onClick={handleDownloadPdf}
              disabled={isDownloading}
              className="bg-emerald-600 hover:bg-emerald-700 text-white px-3.5 py-1.5 rounded text-xs font-bold transition flex items-center gap-1.5 cursor-pointer shadow-2xs disabled:opacity-50"
            >
              <Download className="w-4 h-4" />
              <span>{isDownloading ? 'កំពុងទាញយក...' : 'ទាញយក PDF'}</span>
            </button>

            <button
              onClick={onClose}
              className="p-1 hover:bg-white/20 rounded text-white transition cursor-pointer"
              title="បិទ"
            >
              <X className="w-5 h-5" />
            </button>
          </div>
        </div>

        {/* Download Success Alert Banner */}
        {downloadSuccess && (
          <div className="bg-emerald-50 border-b border-emerald-200 p-3 px-4 flex flex-wrap items-center justify-between gap-3 text-xs text-emerald-900 animate-fade shrink-0 print:hidden">
            <div className="flex items-center gap-2">
              <div className="w-6 h-6 rounded-full bg-emerald-600 text-white flex items-center justify-center font-bold text-xs shrink-0 shadow-2xs">
                <Check className="w-4 h-4" />
              </div>
              <div>
                <span className="font-bold text-xs sm:text-sm">បានបង្កើត និងទាញយកឯកសារ PDF រួចរាល់ដោយជោគជ័យ!</span>
                <span className="block text-[11px] text-emerald-700 mt-0.5">
                  ឯកសារត្រូវបានរក្សាទុក។ ប្រសិនបើកម្មវិធីរុករកមិនបានទាញយកដោយស្វ័យប្រវត្តិទេ អ្នកអាចចុចប៊ូតុងខាងក្រោមដើម្បីបើកមើល ឬបោះពុម្ពបាន។
                </span>
              </div>
            </div>
            <div className="flex items-center gap-2">
              {downloadedBlobUrl && (
                <a
                  href={downloadedBlobUrl}
                  target="_blank"
                  rel="noopener noreferrer"
                  className="bg-emerald-600 hover:bg-emerald-700 text-white px-3 py-1.5 rounded text-xs font-bold transition flex items-center gap-1.5 cursor-pointer shadow-2xs"
                >
                  <FileText className="w-3.5 h-3.5" />
                  <span>បើកមើល PDF ក្នុង Tab ថ្មី</span>
                </a>
              )}
              <button
                type="button"
                onClick={handleDownloadPdf}
                disabled={isDownloading}
                className="bg-emerald-700 hover:bg-emerald-800 text-white px-3 py-1.5 rounded text-xs font-bold transition flex items-center gap-1.5 cursor-pointer shadow-2xs disabled:opacity-50"
              >
                <Download className="w-3.5 h-3.5" />
                <span>{isDownloading ? 'កំពុងទាញយក...' : 'ទាញយកម្តងទៀត'}</span>
              </button>
              <button
                type="button"
                onClick={() => setDownloadSuccess(false)}
                className="p-1 hover:bg-emerald-200/60 rounded text-emerald-700 transition cursor-pointer"
                title="លាក់ដំណឹង"
              >
                <X className="w-4 h-4" />
              </button>
            </div>
          </div>
        )}

        {/* Edit Panel Drawer (Allows editing points highlighted with colored boxes) */}
        {showEditPanel && (
          <div className="bg-amber-50/90 border-b border-amber-200 p-3 sm:p-4 shrink-0 text-xs text-gray-800 print:hidden transition-all">
            <div className="flex items-center justify-between mb-2">
              <h4 className="font-bold text-amber-900 flex items-center gap-1.5 text-xs uppercase tracking-wide">
                <Sparkles className="w-4 h-4 text-amber-600" />
                <span>កែប្រែចំណុចប្រអប់ពណ៌ (Editable Highlighted Fields):</span>
              </h4>
              <label className="flex items-center gap-2 cursor-pointer text-xs font-bold text-gray-700 bg-white/80 px-2.5 py-1 rounded border border-amber-300">
                <input
                  type="checkbox"
                  checked={showColoredBoxesInPrint}
                  onChange={(e) => setShowColoredBoxesInPrint(e.target.checked)}
                  className="rounded text-amber-600 focus:ring-amber-500 cursor-pointer"
                />
                <span>បង្ហាញប្រអប់ពណ៌ (Show Colored Boxes)</span>
              </label>
            </div>

            <div className="grid grid-cols-1 md:grid-cols-3 gap-3">
              {/* Point 1: Red Box (Ref Role Title) */}
              <div className="bg-white p-2.5 rounded border border-red-300 shadow-2xs">
                <label className="block font-bold text-red-700 mb-1 flex items-center gap-1 text-[11px]">
                  <span className="w-2.5 h-2.5 rounded-full bg-red-500 inline-block"></span>
                  <span>១. ប្រអប់ក្រហម (តួនាទីស្នើសុំ)</span>
                </label>
                <select
                  value={refRoleTitle}
                  onChange={(e) => setRefRoleTitle(e.target.value)}
                  className="w-full border border-gray-300 rounded px-2 py-1 text-xs font-semibold text-gray-800 focus:border-red-500 focus:ring-1 focus:ring-red-500 focus:outline-none bg-white cursor-pointer"
                >
                  <option value="">-- ជ្រើសរើសអ្នកស្នើសុំ --</option>
                  {collectorRolesList.map((cr) => (
                    <option key={cr.id} value={cr.name}>
                      {cr.name}
                    </option>
                  ))}
                  {refRoleTitle && !collectorRolesList.some((cr) => cr.name === refRoleTitle) && (
                    <option value={refRoleTitle}>{refRoleTitle}</option>
                  )}
                </select>
              </div>

              {/* Point 2: Yellow Box 1 (Giver Full Info) */}
              <div className="bg-white p-2.5 rounded border border-yellow-400 shadow-2xs">
                <label className="block font-bold text-yellow-800 mb-1 flex items-center gap-1 text-[11px]">
                  <span className="w-2.5 h-2.5 rounded-full bg-yellow-400 inline-block"></span>
                  <span>២. ប្រអប់លឿង (ភាគីប្រគល់)</span>
                </label>
                <select
                  value={giverFullInfo}
                  onChange={(e) => setGiverFullInfo(e.target.value)}
                  className="w-full border border-gray-300 rounded px-2 py-1 text-xs text-gray-800 focus:border-yellow-500 focus:ring-1 focus:ring-yellow-500 focus:outline-none bg-white cursor-pointer"
                >
                  <option value="">-- សូមជ្រើសរើសអ្នកប្រគល់ --</option>
                  {giverOfficersToDisplay.map((off) => {
                    const formatted = formatOfficerGiverInfo(off, effectiveCategories);
                    const posObj = effectiveCategories.positions?.find((p) => p.id === off.positionId);
                    const posLabel = posObj ? ` (${posObj.name})` : '';
                    return (
                      <option key={off.id} value={formatted}>
                        {off.name}{posLabel}
                      </option>
                    );
                  })}
                  {(!giverOfficersToDisplay.length || !giverOfficersToDisplay.some((off) => formatOfficerGiverInfo(off, effectiveCategories) === sampleGiverInfo)) && (
                    <option value={sampleGiverInfo}>អ៊ុក រ័ត្នបញ្ញា</option>
                  )}
                  {giverFullInfo &&
                    giverFullInfo !== sampleGiverInfo &&
                    !giverOfficersToDisplay.some((off) => formatOfficerGiverInfo(off, effectiveCategories) === giverFullInfo) && (
                      <option value={giverFullInfo}>{giverFullInfo}</option>
                    )}
                </select>
              </div>

              {/* Point 3: Yellow Box 2 (Receiver Rank Only) */}
              <div className="bg-white p-2.5 rounded border border-yellow-400 shadow-2xs">
                <label className="block font-bold text-yellow-800 mb-1 flex items-center gap-1 text-[11px]">
                  <span className="w-2.5 h-2.5 rounded-full bg-yellow-400 inline-block"></span>
                  <span>៣. ប្រអប់លឿង (ភាគីទទួល)</span>
                </label>
                <select
                  value={receiverRank}
                  onChange={(e) => {
                    const newRank = e.target.value;
                    setReceiverRank(newRank);
                    setReceiverFullInfo(formatReceiverFullInfo(newRank, receiverName, receiverRole, teamName));
                  }}
                  className="w-full border border-gray-300 rounded px-2 py-1 text-xs text-gray-800 focus:border-yellow-500 focus:ring-1 focus:ring-yellow-500 focus:outline-none bg-white cursor-pointer"
                >
                  <option value="">-- ជ្រើសរើសឋានន្តរ័ក្កិ --</option>
                  {(effectiveCategories.ranks || []).map((rk) => (
                    <option key={rk.id} value={rk.name}>
                      {rk.name}
                    </option>
                  ))}
                  {receiverRank && !(effectiveCategories.ranks || []).some((rk) => rk.name === receiverRank) && (
                    <option value={receiverRank}>{receiverRank}</option>
                  )}
                </select>
              </div>


            </div>

            {/* Additional Header Customization Row */}
            <div className="grid grid-cols-1 sm:grid-cols-2 md:grid-cols-3 lg:grid-cols-5 xl:grid-cols-6 gap-3 mt-3 pt-2 border-t border-amber-200">
              <div>
                <label className="block font-bold text-gray-700 mb-1 text-[11px]">
                  ក្រុមផ្តល់ទិដ្ឋាការ (ឈ្មោះពេញ):
                </label>
                <select
                  value={teamName}
                  onChange={(e) => {
                    const newTeam = e.target.value;
                    setTeamName(newTeam);
                    setReceiverFullInfo(formatReceiverFullInfo(receiverRank, receiverName, receiverRole, newTeam));
                  }}
                  className="w-full border border-gray-300 rounded px-2 py-1 text-xs text-gray-800 focus:border-blue-500 focus:outline-none bg-white cursor-pointer"
                >
                  {(effectiveCategories.visaTeams || []).map((vt) => (
                    <option key={vt.id} value={vt.name}>
                      {vt.name}
                    </option>
                  ))}
                  {teamName && !(effectiveCategories.visaTeams || []).some((vt) => vt.name === teamName) && (
                    <option value={teamName}>{teamName}</option>
                  )}
                </select>
              </div>

              <div>
                <label className="block font-bold text-gray-700 mb-1 text-[11px]">
                  អ្នកស្នើសុំ (ប្រធានក្រុម):
                </label>
                <input
                  type="text"
                  value={requesterName}
                  onChange={(e) => setRequesterName(e.target.value)}
                  placeholder="បញ្ចូលឈ្មោះអ្នកស្នើសុំ..."
                  className="w-full border border-gray-300 rounded px-2 py-1 text-xs text-gray-800 focus:border-blue-500 focus:outline-none bg-white"
                />
              </div>

              {isEvisa && (
                <div>
                  <label className="block font-bold text-amber-900 mb-1 text-[11px] flex items-center gap-1">
                    <span className="w-2 h-2 rounded-full bg-amber-500"></span>
                    <span>ចំនួនក្រដាសបើកជូន (ដុំ):</span>
                  </label>
                  <input
                    type="number"
                    min={1}
                    value={evisaBundles}
                    onChange={(e) => setEvisaBundles(e.target.value)}
                    placeholder="ចំនួនដុំ..."
                    className="w-full border border-amber-400 rounded px-2 py-1 text-xs text-gray-900 focus:border-amber-600 focus:outline-none bg-yellow-50/70 font-bold"
                  />
                </div>
              )}

              <div>
                <label className="block font-bold text-gray-700 mb-1 text-[11px]">
                  តួនាទីអ្នកចុះហត្ថលេខាបាតក្រោម (Approval Title):
                </label>
                <select
                  value={approverTitle}
                  onChange={(e) => handleTitleChange(e.target.value)}
                  className="w-full border border-gray-300 rounded px-2 py-1 text-xs text-gray-800 focus:border-blue-500 focus:outline-none bg-white cursor-pointer mb-1.5"
                >
                  <option value="ភប.នាយការិយាល័យទិដ្ឋាការចូល">ភប.នាយការិយាល័យទិដ្ឋាការចូល</option>
                  <option value="តប.នាយការិយាល័យទិដ្ឋាការចូល">តប.នាយការិយាល័យទិដ្ឋាការចូល</option>
                  <option value="នាយការិយាល័យ">នាយការិយាល័យ</option>
                  <option value="នាយផ្នែក">នាយផ្នែក</option>
                  <option value="ជ.នាយផ្នែក">ជ.នាយផ្នែក</option>
                </select>

                <select
                  value={
                    approverOfficersToDisplay.some((off) => getOfficerDisplayName(off) === approverName)
                      ? approverName
                      : approverName
                      ? 'custom'
                      : ''
                  }
                  onChange={(e) => {
                    if (e.target.value !== 'custom' && e.target.value !== '') {
                      setApproverName(e.target.value);
                    }
                  }}
                  className="w-full border border-gray-300 rounded px-2 py-1 text-xs text-gray-800 focus:border-blue-500 focus:outline-none bg-white cursor-pointer"
                >
                  <option value="">-- ជ្រើសរើសមន្ត្រី --</option>
                  {approverOfficersToDisplay.map((off) => {
                    const dispName = getOfficerDisplayName(off);
                    return (
                      <option key={off.id} value={dispName}>
                        {dispName} {off.officerNumber ? `(អត្តលេខ: ${off.officerNumber})` : ''}
                      </option>
                    );
                  })}
                  <option value="custom">-- ផ្សេងៗ --</option>
                </select>

                {!approverOfficersToDisplay.some((off) => getOfficerDisplayName(off) === approverName) && (
                  <input
                    type="text"
                    value={approverName}
                    onChange={(e) => setApproverName(e.target.value)}
                    placeholder="បញ្ចូលឈ្មោះអ្នកអនុម័ត..."
                    className="w-full border border-gray-300 rounded px-2 py-1 text-xs text-gray-800 focus:border-blue-500 focus:outline-none bg-white mt-1.5"
                  />
                )}
              </div>

              <div>
                <label className="block font-bold text-gray-700 mb-1 text-[11px]">
                  ទំហំក្បាលលិខិត (Header Font Size)៖
                </label>
                <div className="flex items-center gap-1.5">
                  <input
                    type="number"
                    min={9}
                    max={16}
                    step={0.5}
                    value={headerFontSizePt}
                    onChange={(e) => {
                      const val = parseFloat(e.target.value) || 12;
                      setHeaderFontSizePt(val);
                      localStorage.setItem('stock_pdf_header_font_size_pt', String(val));
                    }}
                    className="w-16 border border-gray-300 rounded px-2 py-1 text-xs text-gray-800 focus:border-blue-500 focus:outline-none bg-white font-bold"
                  />
                  <span className="text-[10px] text-gray-500 font-medium">pt (ស្តង់ដារ 12pt)</span>
                </div>
              </div>

              <div>
                <label className="block font-bold text-gray-700 mb-1 text-[11px]">
                  ទំហំអក្សរ (Font Size)៖
                </label>
                <div className="flex items-center gap-1.5">
                  <input
                    type="number"
                    min={9}
                    max={16}
                    step={0.5}
                    value={bodyFontSizePt}
                    onChange={(e) => {
                      const val = parseFloat(e.target.value) || 12;
                      setBodyFontSizePt(val);
                      localStorage.setItem('stock_pdf_font_size_pt', String(val));
                    }}
                    className="w-16 border border-gray-300 rounded px-2 py-1 text-xs text-gray-800 focus:border-blue-500 focus:outline-none bg-white font-bold"
                  />
                  <span className="text-[10px] text-gray-500 font-medium">pt (ស្តង់ដារ 12pt)</span>
                </div>
              </div>

              <div>
                <label className="block font-bold text-gray-700 mb-1 text-[11px]">
                  ទំហំចំណងជើង (Title Font Size)៖
                </label>
                <div className="flex items-center gap-1.5">
                  <input
                    type="number"
                    min={9}
                    max={20}
                    step={0.5}
                    value={titleFontSizePt}
                    onChange={(e) => {
                      const val = parseFloat(e.target.value) || 12;
                      setTitleFontSizePt(val);
                      localStorage.setItem('stock_pdf_title_font_size_pt', String(val));
                    }}
                    className="w-16 border border-gray-300 rounded px-2 py-1 text-xs text-gray-800 focus:border-blue-500 focus:outline-none bg-white font-bold"
                  />
                  <span className="text-[10px] text-gray-500 font-medium">pt (ស្តង់ដារ 12pt)</span>
                </div>
              </div>

              <div>
                <label className="block font-bold text-gray-700 mb-1 text-[11px]">
                  Margin Top (cm)៖
                </label>
                <div className="flex items-center gap-1.5">
                  <input
                    type="number"
                    min={0}
                    max={5}
                    step={0.1}
                    value={docMarginTopCm}
                    onChange={(e) => {
                      const val = parseFloat(e.target.value) >= 0 ? parseFloat(e.target.value) : 1;
                      setDocMarginTopCm(val);
                      localStorage.setItem('stock_pdf_margin_top_cm', String(val));
                    }}
                    className="w-16 border border-gray-300 rounded px-2 py-1 text-xs text-gray-800 focus:border-blue-500 focus:outline-none bg-white font-bold"
                  />
                  <span className="text-[10px] text-gray-500 font-medium">cm (1.0cm)</span>
                </div>
              </div>

              <div>
                <label className="block font-bold text-gray-700 mb-1 text-[11px]">
                  Margin Bottom (cm)៖
                </label>
                <div className="flex items-center gap-1.5">
                  <input
                    type="number"
                    min={0}
                    max={5}
                    step={0.1}
                    value={docMarginBottomCm}
                    onChange={(e) => {
                      const val = parseFloat(e.target.value) >= 0 ? parseFloat(e.target.value) : 0.5;
                      setDocMarginBottomCm(val);
                      localStorage.setItem('stock_pdf_margin_bottom_cm', String(val));
                    }}
                    className="w-16 border border-gray-300 rounded px-2 py-1 text-xs text-gray-800 focus:border-blue-500 focus:outline-none bg-white font-bold"
                  />
                  <span className="text-[10px] text-gray-500 font-medium">cm (0.5cm)</span>
                </div>
              </div>

              <div>
                <label className="block font-bold text-amber-700 mb-1 text-[11px]">
                  Margin Left គែមឆ្វេង (cm)៖
                </label>
                <div className="flex items-center gap-1.5">
                  <input
                    type="number"
                    min={0.5}
                    max={5}
                    step={0.1}
                    value={docMarginLeftCm}
                    onChange={(e) => {
                      const val = parseFloat(e.target.value) >= 0 ? parseFloat(e.target.value) : 2.8;
                      setDocMarginLeftCm(val);
                      localStorage.setItem('stock_pdf_margin_left_cm', String(val));
                    }}
                    className="w-16 border border-amber-400 rounded px-2 py-1 text-xs text-gray-800 focus:border-blue-500 focus:outline-none bg-amber-50 font-bold"
                  />
                  <span className="text-[10px] text-gray-500 font-medium">cm (2.8cm)</span>
                </div>
              </div>

              <div>
                <label className="block font-bold text-gray-700 mb-1 text-[11px]">
                  Margin Right គែមស្តាំ (cm)៖
                </label>
                <div className="flex items-center gap-1.5">
                  <input
                    type="number"
                    min={0.5}
                    max={5}
                    step={0.1}
                    value={docMarginRightCm}
                    onChange={(e) => {
                      const val = parseFloat(e.target.value) >= 0 ? parseFloat(e.target.value) : 1.3;
                      setDocMarginRightCm(val);
                      localStorage.setItem('stock_pdf_margin_right_cm', String(val));
                    }}
                    className="w-16 border border-gray-300 rounded px-2 py-1 text-xs text-gray-800 focus:border-blue-500 focus:outline-none bg-white font-bold"
                  />
                  <span className="text-[10px] text-gray-500 font-medium">cm (1.3cm)</span>
                </div>
              </div>

              <div>
                <label className="block font-bold text-amber-700 mb-1 text-[11px]">
                  Tab Stop ចន្លោះចូលបន្ទាត់ (cm)៖
                </label>
                <div className="flex items-center gap-1.5">
                  <input
                    type="number"
                    min={0.5}
                    max={8}
                    step={0.05}
                    value={activeTabStop}
                    onChange={(e) => {
                      const val = parseFloat(e.target.value) >= 0 ? parseFloat(e.target.value) : 2.25;
                      setActiveTabStop(val);
                      localStorage.setItem('stock_pdf_active_tab_stop', String(val));
                    }}
                    className="w-16 border border-amber-400 rounded px-2 py-1 text-xs text-gray-800 focus:border-blue-500 focus:outline-none bg-amber-50 font-bold"
                  />
                  <span className="text-[10px] text-gray-500 font-medium">cm (2.25cm)</span>
                </div>
              </div>

              {/* Quick Margin Preset Buttons */}
              <div className="sm:col-span-2 lg:col-span-4 flex flex-wrap items-center gap-2 pt-1 border-t border-gray-200">
                <span className="text-[11px] font-bold text-gray-600">កំណត់រហ័ស៖</span>
                <button
                  type="button"
                  onClick={() => {
                    setDocMarginLeftCm(2.8);
                    setDocMarginRightCm(1.3);
                    setDocMarginTopCm(1.0);
                    setDocMarginBottomCm(0.5);
                    setActiveTabStop(2.25);
                    localStorage.setItem('stock_pdf_margin_left_cm', '2.8');
                    localStorage.setItem('stock_pdf_margin_right_cm', '1.3');
                    localStorage.setItem('stock_pdf_margin_top_cm', '1.0');
                    localStorage.setItem('stock_pdf_margin_bottom_cm', '0.5');
                    localStorage.setItem('stock_pdf_active_tab_stop', '2.25');
                  }}
                  className="px-2.5 py-1 bg-blue-50 border border-blue-300 text-blue-800 hover:bg-blue-100 rounded text-[11px] font-bold transition cursor-pointer"
                >
                  ស្តង់ដារ របក.សរុបការងារស្តុក (ឆ្វេង 2.8cm, ស្តាំ 1.3cm, លើ 1cm, ក្រោម 0.5cm)
                </button>
                <button
                  type="button"
                  onClick={() => {
                    setDocMarginLeftCm(2.5);
                    setDocMarginRightCm(1.5);
                    setDocMarginTopCm(1.2);
                    setDocMarginBottomCm(1.2);
                    setActiveTabStop(2.0);
                    localStorage.setItem('stock_pdf_margin_left_cm', '2.5');
                    localStorage.setItem('stock_pdf_margin_right_cm', '1.5');
                    localStorage.setItem('stock_pdf_margin_top_cm', '1.2');
                    localStorage.setItem('stock_pdf_margin_bottom_cm', '1.2');
                    localStorage.setItem('stock_pdf_active_tab_stop', '2.0');
                  }}
                  className="px-2.5 py-1 bg-gray-100 border border-gray-300 text-gray-700 hover:bg-gray-200 rounded text-[11px] font-medium transition cursor-pointer"
                >
                  ស្តង់ដាររដ្ឋបាល (ឆ្វេង 2.5cm, 1.2cm)
                </button>
              </div>

              <div>
                <label className="block font-bold text-gray-700 mb-1 text-[11px]">
                  គម្លាតហត្ថលេខាសងខាង (px)៖
                </label>
                <div className="flex items-center gap-1.5">
                  <input
                    type="number"
                    min={0}
                    max={150}
                    step={5}
                    value={sigShiftInset}
                    onChange={(e) => {
                      const val = Number(e.target.value) || 0;
                      setSigShiftInset(val);
                      localStorage.setItem('stock_pdf_sig_inset', String(val));
                    }}
                    className="w-20 border border-gray-300 rounded px-2 py-1 text-xs text-gray-800 focus:border-blue-500 focus:outline-none bg-white"
                  />
                  <span className="text-[10px] text-gray-500 font-medium">0 = ខិតឆ្វេង/ស្តាំអតិបរមា</span>
                </div>
              </div>

              {isEvisa && (
                <>
                  <div>
                    <label className="block font-bold text-gray-700 mb-1 text-[11px]">
                      គម្លាតលើ-ក្រោមក្រដាស (px)៖
                    </label>
                    <div className="flex items-center gap-1.5">
                      <input
                        type="number"
                        min={0}
                        max={60}
                        step={2}
                        value={evisaGapSpacing}
                        onChange={(e) => {
                          const val = Number(e.target.value) >= 0 ? Number(e.target.value) : 0;
                          setEvisaGapSpacing(val);
                          localStorage.setItem('stock_pdf_evisa_gap', String(val));
                        }}
                        className="w-16 border border-gray-300 rounded px-2 py-1 text-xs text-gray-800 focus:border-blue-500 focus:outline-none bg-white"
                      />
                      <span className="text-[10px] text-gray-500 font-medium">លំនាំដើម: 6</span>
                    </div>
                  </div>

                  <div>
                    <label className="block font-bold text-gray-700 mb-1 text-[11px]">
                      គម្លាតលើថ្ងៃខែ (px)៖
                    </label>
                    <div className="flex items-center gap-1.5">
                      <input
                        type="number"
                        min={0}
                        max={60}
                        step={2}
                        value={sigTopSpacing}
                        onChange={(e) => {
                          const val = Number(e.target.value) >= 0 ? Number(e.target.value) : 0;
                          setSigTopSpacing(val);
                          localStorage.setItem('stock_pdf_sig_top_spacing', String(val));
                        }}
                        className="w-16 border border-gray-300 rounded px-2 py-1 text-xs text-gray-800 focus:border-blue-500 focus:outline-none bg-white"
                      />
                      <span className="text-[10px] text-gray-500 font-medium">លំនាំដើម: 8</span>
                    </div>
                  </div>
                </>
              )}
            </div>
          </div>
        )}

        {/* Modal Main Body / Document Preview Area */}
        <div className="p-4 sm:p-8 overflow-y-auto flex-1 bg-gray-200 print:bg-white print:p-0 print:overflow-visible flex flex-col items-center">
          {/* Top Interactive Word Ruler Bar (A4 exact scale) */}
          {showRuler && (
            <div
              data-ruler-container="true"
              className="print:hidden w-[210mm] h-6 bg-slate-200 border-b border-t border-slate-400 relative select-none cursor-pointer flex items-center shadow-xs overflow-hidden mb-1 rounded-t-sm"
              title="ចុចលើបន្ទាត់ដើម្បីកំណត់ទីតាំង Tab Stop (ចន្លោះចូលបន្ទាត់)"
              onClick={(e) => {
                const rect = e.currentTarget.getBoundingClientRect();
                const clickX_px = e.clientX - rect.left;
                const cmScale = rect.width / 21.0;
                const clickedTotalCm = clickX_px / cmScale;
                const relativeTab = parseFloat((clickedTotalCm - docMarginLeftCm).toFixed(2));
                if (relativeTab >= 0.5 && relativeTab <= 8.0) {
                  setActiveTabStop(relativeTab);
                  localStorage.setItem('stock_pdf_active_tab_stop', String(relativeTab));
                }
              }}
            >
              {/* Shaded Left Margin Area */}
              <div
                className="absolute top-0 bottom-0 left-0 bg-slate-300/80 border-r border-slate-400/80 pointer-events-none"
                style={{ width: `${docMarginLeftCm}cm` }}
              />
              {/* White Printable Area on Ruler */}
              <div
                className="absolute top-0 bottom-0 bg-white pointer-events-none"
                style={{
                  left: `${docMarginLeftCm}cm`,
                  width: `${Math.max(1, 21.0 - docMarginLeftCm - docMarginRightCm)}cm`,
                }}
              />
              {/* Shaded Right Margin Area */}
              <div
                className="absolute top-0 bottom-0 right-0 bg-slate-300/80 border-l border-slate-400/80 pointer-events-none"
                style={{ width: `${docMarginRightCm}cm` }}
              />

              {/* cm Scale Numbers & Ticks */}
              <div className="relative w-full h-full pointer-events-none">
                {Array.from({ length: 22 }).map((_, cmIdx) => {
                  const relativeCm = Math.round(cmIdx - docMarginLeftCm);
                  return (
                    <div
                      key={cmIdx}
                      className="absolute top-0 flex flex-col items-center pointer-events-none"
                      style={{ left: `${cmIdx}cm` }}
                    >
                      <div className="w-[1px] h-3 bg-slate-600" />
                      {/* Half cm tick */}
                      {cmIdx < 21 && (
                        <div
                          className="absolute top-0 w-[1px] h-2 bg-slate-400"
                          style={{ left: '0.5cm' }}
                        />
                      )}
                      {/* Quarter cm ticks */}
                      {cmIdx < 21 && (
                        <>
                          <div
                            className="absolute top-0 w-[1px] h-1.5 bg-slate-300"
                            style={{ left: '0.25cm' }}
                          />
                          <div
                            className="absolute top-0 w-[1px] h-1.5 bg-slate-300"
                            style={{ left: '0.75cm' }}
                          />
                        </>
                      )}
                      {/* Number label */}
                      <span className="text-[8.5px] font-semibold text-slate-700 mt-0.5 -ml-1">
                        {relativeCm >= 0 ? relativeCm : ''}
                      </span>
                    </div>
                  );
                })}
              </div>

              {/* Active Tab Stop Marker (Classic Word Tab Stop L/Marker on Ruler) */}
              <div
                className="absolute top-0 z-20 -ml-1.5 flex flex-col items-center pointer-events-none"
                style={{ left: `${docMarginLeftCm + activeTabStop}cm` }}
                title={`Tab Stop កំណត់នៅ: ${activeTabStop} cm`}
              >
                <div className="w-0 h-0 border-l-[4px] border-l-transparent border-r-[4px] border-r-transparent border-t-[5px] border-t-amber-600" />
                <div className="w-2.5 h-3 bg-amber-500 rounded-[1px] shadow-sm flex items-center justify-center text-[7px] text-black font-extrabold">
                  {activeTabStop}
                </div>
                <div className="w-0 h-0 border-l-[4px] border-l-transparent border-r-[4px] border-r-transparent border-b-[5px] border-b-amber-600" />
              </div>

              {/* Left Margin Indicator Marker */}
              <div
                className="absolute top-0 z-10 -ml-1.5 flex flex-col items-center cursor-pointer group pointer-events-none"
                style={{ left: `${docMarginLeftCm}cm` }}
                title={`គែមឆ្វេង (Left Margin): ${docMarginLeftCm} cm`}
              >
                <div className="w-0 h-0 border-l-[4px] border-l-transparent border-r-[4px] border-r-transparent border-t-[5px] border-t-blue-600" />
                <div className="w-2 h-2.5 bg-blue-600 rounded-[1px] shadow-sm flex items-center justify-center text-[6.5px] text-white font-bold">
                  L
                </div>
                <div className="w-0 h-0 border-l-[4px] border-l-transparent border-r-[4px] border-r-transparent border-b-[5px] border-b-blue-600" />
              </div>

              {/* Right Margin Indicator Marker */}
              <div
                className="absolute top-0 z-10 -mr-1.5 flex flex-col items-center cursor-pointer group pointer-events-none"
                style={{ right: `${docMarginRightCm}cm` }}
                title={`គែមស្តាំ (Right Margin): ${docMarginRightCm} cm`}
              >
                <div className="w-0 h-0 border-l-[4px] border-l-transparent border-r-[4px] border-r-transparent border-t-[5px] border-t-blue-600" />
                <div className="w-2 h-2.5 bg-blue-600 rounded-[1px] shadow-sm flex items-center justify-center text-[6.5px] text-white font-bold">
                  R
                </div>
                <div className="w-0 h-0 border-l-[4px] border-l-transparent border-r-[4px] border-r-transparent border-b-[5px] border-b-blue-600" />
              </div>
            </div>
          )}

          {/* Printable Official Khmer Document Frame - Strictly Formatted to A4 Dimensions with exact margins */}
          <div
            ref={documentRef}
            id="official-pdf-document"
            contentEditable={isWordEditMode}
            suppressContentEditableWarning={true}
            className={`bg-white shadow-lg border border-gray-300 rounded-sm text-black leading-relaxed text-sm w-[210mm] min-h-[297mm] max-w-full print:shadow-none print:border-none print:m-0 print:w-[210mm] print:min-h-[297mm] print:max-w-none flex flex-col justify-start font-siemreap ${
              isWordEditMode
                ? 'outline-none ring-2 ring-blue-400/40 hover:ring-blue-500 cursor-text select-text'
                : 'outline-none cursor-default'
            }`}
            style={{
              boxSizing: 'border-box',
              paddingTop: `${docMarginTopCm}cm`,
              paddingLeft: `${docMarginLeftCm}cm`,
              paddingRight: `${docMarginRightCm}cm`,
              paddingBottom: `${docMarginBottomCm}cm`,
            }}
          >
            <div>
              {/* Cambodian Government Official Header */}
              <div className="flex justify-between items-start mb-3 pt-1 font-siemreap">
                {/* Left Header Block */}
                <div
                  className="text-center font-moul leading-relaxed space-y-0.5 text-black pt-4"
                  style={{ fontSize: `${headerFontSizePt}pt` }}
                >
                  <p className="font-moul" style={{ fontSize: `${headerFontSizePt}pt` }}>ក្រសួងមហាផ្ទៃ</p>
                  <p className="font-moul" style={{ fontSize: `${headerFontSizePt}pt` }}>អគ្គនាយកដ្ឋានអន្តោប្រវេសន៍</p>
                  <p className="font-moul" style={{ fontSize: `${headerFontSizePt}pt` }}>នាយកដ្ឋានជនបរទេសមិនមែនអន្តោប្រវេសន្ត</p>
                  <p className="font-moul" style={{ fontSize: `${headerFontSizePt}pt` }}>ការិយាល័យទិដ្ឋាការចូល</p>
                  <p className="font-moul mb-1" style={{ fontSize: `${headerFontSizePt}pt` }}>ផ្នែករដ្ឋបាល</p>
                  {/* TacTeing Decorative Line - Height 16px with clear separation */}
                  <div className="flex justify-center pt-1">
                    <TacteingLine type={tacteingType} customImage={tacteingCustomImage} width={160} height={16} />
                  </div>
                </div>

                {/* Right Header Block (Kingdom Header) */}
                <div
                  className="text-center font-moul leading-relaxed space-y-0.5 text-black"
                  style={{ fontSize: `${headerFontSizePt}pt` }}
                >
                  <p className="font-moul" style={{ fontSize: `${headerFontSizePt}pt` }}>ព្រះរាជាណាចក្រកម្ពុជា</p>
                  <p className="font-moul mb-1" style={{ fontSize: `${headerFontSizePt}pt` }}>ជាតិ សាសនា ព្រះមហាក្សត្រ</p>
                  {/* TacTeing Decorative Line - Height 16px with clear separation */}
                  <div className="flex justify-center pt-1">
                    <TacteingLine type={tacteingType} customImage={tacteingCustomImage} width={160} height={16} />
                  </div>
                </div>
              </div>

              {/* Document Title (Moul Font) */}
              <div className="text-center my-2.5">
                <h1
                  className="font-moul font-normal text-black leading-normal"
                  style={{ fontSize: `${titleFontSizePt}pt` }}
                >
                  {docTitle}
                </h1>
              </div>

              {/* Intro Paragraph */}
              <div className="mb-2 font-siemreap" style={{ fontSize: `${bodyFontSizePt}pt` }}>
                <p
                  className="text-justify text-black leading-normal"
                  style={{ textAlign: 'justify', fontSize: `${bodyFontSizePt}pt`, textIndent: `${activeTabStop}cm` }}
                >
                  {lunarDate ? lunarDate.replace(/ព\.ស\.\s*/g, 'ព.ស') : ''} ត្រូវនឹងថ្ងៃទី{khmerDay} ខែ{khmerMonth} ឆ្នាំ{khmerYear} វេលាម៉ោង {khmerTime}នាទី {isEvisa ? 'ផ្នែករដ្ឋបាល បានធ្វើជាសាក្សីការិយាល័យទិដ្ឋាការចូល បើកផ្តល់ក្រដាសអនុម័តទិដ្ឋាការអេឡិចត្រូនិក។' : `ផ្នែករដ្ឋបាល បានធ្វើជាសេនាធិការជូនការិយាល័យទិដ្ឋាការចូល បើកផ្តល់${itemTerm}។`}
                </p>
              </div>

              {/* Reference & Parties Details List */}
              <div
                className={`space-y-1.5 text-black leading-normal ${isEvisa ? 'mb-2' : 'mb-2.5'} font-siemreap text-justify`}
                style={{ textAlign: 'justify', fontSize: `${bodyFontSizePt}pt` }}
              >
                {/* Reference Line */}
                <div className="flex items-baseline">
                  <div className="w-[96px] shrink-0 flex items-baseline justify-between pr-2.5 font-moul font-normal text-black" style={{ fontSize: `${bodyFontSizePt}pt` }}>
                    <span>យោង</span>
                    <span className="font-bold">៖</span>
                  </div>
                  <div className="flex-1 text-justify text-black" style={{ textAlign: 'justify' }}>
                    {showColoredBoxesInPrint ? (
                      <>
                        សំណើរបស់{' '}
                        <span className="border border-amber-400 bg-yellow-50/60 px-1.5 py-0.5 rounded-2xs text-gray-950 font-normal">
                          {renderTextWithMoulName(requesterName || 'វរសេនីយ៍ទោ ផែន វិបុល')}
                        </span>{' '}
                        ជា
                        <span
                          className="border border-red-500 bg-red-50/40 px-1.5 py-0.5 rounded-2xs text-red-950 font-normal min-w-[36px] min-h-[22px] text-center"
                          title="ចុចកែប្រែក្នុងផ្ទាំងខាងលើ"
                        >
                          {refRoleTitle}
                        </span>{' '}
                        នៃ
                        <span className="border border-amber-400 bg-yellow-50/60 px-1.5 py-0.5 rounded-2xs text-gray-950 font-normal">
                          {teamName}
                        </span>
                        ។
                      </>
                    ) : (
                      <span>{renderTextWithMoulName(getReferenceFullText())}</span>
                    )}
                  </div>
                </div>

                {/* Giver Line */}
                <div className="flex items-baseline">
                  <div className="w-[96px] shrink-0 flex items-baseline justify-between pr-2.5 font-moul font-normal text-black" style={{ fontSize: `${bodyFontSizePt}pt` }}>
                    <span>ភាគីប្រគល់</span>
                    <span className="font-bold">៖</span>
                  </div>
                  <div className="flex-1 text-justify text-black" style={{ textAlign: 'justify' }}>
                    {showColoredBoxesInPrint ? (
                      <span
                        className="border border-amber-400 bg-yellow-50/60 px-1.5 py-0.5 rounded-2xs text-gray-950 font-normal min-w-[36px] min-h-[22px]"
                        title="ចុចកែប្រែក្នុងផ្ទាំងខាងលើ"
                      >
                        {giverFullInfo ? renderTextWithMoulName(giverFullInfo) : ''}
                      </span>
                    ) : (
                      <span>{giverFullInfo ? renderTextWithMoulName(giverFullInfo) : ''}</span>
                    )}
                  </div>
                </div>

                {/* Receiver Line */}
                <div className="flex items-baseline">
                  <div className="w-[96px] shrink-0 flex items-baseline justify-between pr-2.5 font-moul font-normal text-black" style={{ fontSize: `${bodyFontSizePt}pt` }}>
                    <span>ភាគីទទួល</span>
                    <span className="font-bold">៖</span>
                  </div>
                  <div className="flex-1 text-justify text-black" style={{ textAlign: 'justify' }}>
                    {showColoredBoxesInPrint ? (
                      <>
                        <span className="border border-amber-400 bg-yellow-50/60 px-1.5 py-0.5 rounded-2xs text-gray-950 font-normal">
                          {receiverRank}
                        </span>{' '}
                        <span className="font-moul text-black">{receiverName}</span> ជា
                        <span className="border border-amber-400 bg-yellow-50/60 px-1.5 py-0.5 rounded-2xs text-gray-950 font-normal mx-0.5">
                          {receiverRole}
                        </span>{' '}
                        នៃ
                        <span className="border border-amber-400 bg-yellow-50/60 px-1.5 py-0.5 rounded-2xs text-gray-950 font-normal mx-0.5">
                          {teamName}
                        </span>
                        ។
                      </>
                    ) : (
                      <span>
                        {renderTextWithMoulName(
                          formatReceiverFullInfo(receiverRank, receiverName, receiverRole, teamName)
                        )}
                      </span>
                    )}
                  </div>
                </div>
              </div>

              {isEvisa ? (
                /* eVisa Handover Statement - Aligned horizontally with receiver rank and name */
                <div
                  className="font-siemreap text-black flex items-baseline"
                  style={{
                    marginTop: `${evisaGapSpacing}px`,
                    marginBottom: `${evisaGapSpacing}px`,
                    fontSize: `${bodyFontSizePt}pt`,
                  }}
                >
                  {/* Invisible spacer matching the w-[96px] label width so the text stands equal to វរសេនីយ៍ឯក រស់ លីហួត */}
                  <div className="w-[96px] shrink-0 pr-2.5 invisible select-none" aria-hidden="true">
                    <span className="font-moul">ភាគីទទួល</span>
                    <span>៖</span>
                  </div>
                  <p className="flex-1 font-bold leading-relaxed text-left text-black" style={{ fontSize: `${bodyFontSizePt}pt` }}>
                    បានប្រគល់និងទទួលក្រដាសអនុម័តទិដ្ឋាការអេឡិចត្រូនិក ចំនួន{' '}
                    <span
                      className={`inline-block ${
                        showColoredBoxesInPrint
                          ? 'border border-amber-400 bg-yellow-50/60 px-2 py-0.5 rounded-2xs text-gray-950 font-bold mx-1'
                          : 'font-bold mx-1'
                      }`}
                      title="ចំនួនដុំ (អាចកែប្រែក្នុងផ្ទាំងខាងលើ)"
                    >
                      {toKhmerNum(evisaBundles || 0)}
                    </span>{' '}
                    ដុំ។
                  </p>
                </div>
              ) : (
                /* Sticker Handover Table Statement & Table */
                <div>
                  {/* Handover Statement Header - Aligned horizontally with receiver rank and name */}
                  <div
                    className="font-siemreap text-black flex items-baseline mt-2 mb-2 pt-0.5"
                    style={{ fontSize: `${bodyFontSizePt}pt` }}
                  >
                    {/* Invisible spacer matching the w-[96px] label width so the text stands equal to វរសេនីយ៍ឯក រស់ លីហួត */}
                    <div className="w-[96px] shrink-0 pr-2.5 invisible select-none" aria-hidden="true">
                      <span className="font-moul">ភាគីទទួល</span>
                      <span>៖</span>
                    </div>
                    <p className="flex-1 font-bold leading-normal text-left text-black font-siemreap" style={{ fontSize: `${bodyFontSizePt}pt` }}>
                      បានធ្វើការប្រគល់ និងទទួល{itemTerm} មានចំនួនដូចខាងក្រោម ៖
                    </p>
                  </div>

                  {/* Official Handover Table */}
                  <div className="mt-2 mb-3 relative">
                    <table
                      className="w-full border-collapse border border-black text-center font-siemreap"
                      style={{ fontSize: `${Math.max(10, bodyFontSizePt - 0.5)}pt` }}
                    >
                      <thead>
                        <tr className="bg-white" style={{ verticalAlign: 'middle' }}>
                          <th className="border border-black px-2 py-1.5 font-bold w-[8%] text-center align-middle text-black" style={{ verticalAlign: 'middle', fontSize: `${Math.max(10, bodyFontSizePt - 0.5)}pt` }}>
                            ល.រ
                          </th>
                          <th className="border border-black px-2 py-1.5 font-bold w-[13%] text-center align-middle text-black" style={{ verticalAlign: 'middle', fontSize: `${Math.max(10, bodyFontSizePt - 0.5)}pt` }}>
                            ប្រភេទ
                          </th>
                          <th className="border border-black px-2 py-1.5 font-bold w-[14%] text-center align-middle text-black" style={{ verticalAlign: 'middle', fontSize: `${Math.max(10, bodyFontSizePt - 0.5)}pt` }}>
                            ក្បាល
                          </th>
                          <th className="border border-black px-2 py-1.5 font-bold w-[16%] text-center align-middle text-black" style={{ verticalAlign: 'middle', fontSize: `${Math.max(10, bodyFontSizePt - 0.5)}pt` }}>
                            សន្លឹក
                          </th>
                          <th className="border border-black px-2 py-1.5 font-bold w-[24.5%] text-center align-middle text-black" style={{ verticalAlign: 'middle', fontSize: `${Math.max(10, bodyFontSizePt - 0.5)}pt` }}>
                            ចាប់ពីលេខ
                          </th>
                          <th className="border border-black px-2 py-1.5 font-bold w-[24.5%] text-center align-middle text-black" style={{ verticalAlign: 'middle', fontSize: `${Math.max(10, bodyFontSizePt - 0.5)}pt` }}>
                            ដល់លេខ
                          </th>
                        </tr>
                      </thead>
                      <tbody>
                        {/* Dynamic Visa Item Rows */}
                        {tableItems.map((item, idx) => (
                          <tr key={item.id} className="font-times bg-white" style={{ fontFamily: "'Times New Roman', Times, serif", verticalAlign: 'middle' }}>
                            <td className="border border-black px-2 py-1.5 text-center font-bold font-times align-middle text-black" style={{ fontFamily: "'Times New Roman', Times, serif", verticalAlign: 'middle', fontSize: `${bodyFontSizePt}pt` }}>
                              {formatTableNum(idx + 1)}
                            </td>
                            <td className="border border-black px-2 py-1.5 text-center font-bold font-times align-middle text-black" style={{ fontFamily: "'Times New Roman', Times, serif", verticalAlign: 'middle', fontSize: `${bodyFontSizePt}pt` }}>
                              {item.visaType || 'T'}
                            </td>
                            <td className="border border-black px-2 py-1.5 text-center font-bold font-times align-middle text-black" style={{ fontFamily: "'Times New Roman', Times, serif", verticalAlign: 'middle', fontSize: `${bodyFontSizePt}pt` }}>
                              {showColoredBoxesInPrint ? (
                                <span className="inline-block border border-amber-400 bg-yellow-50/60 px-1 py-0.5 rounded-2xs">
                                  {formatTableNum(item.kbalCount)}
                                </span>
                              ) : (
                                formatTableNum(item.kbalCount)
                              )}
                            </td>
                            <td className="border border-black px-2 py-1.5 text-center font-bold font-times align-middle text-black" style={{ fontFamily: "'Times New Roman', Times, serif", verticalAlign: 'middle', fontSize: `${bodyFontSizePt}pt` }}>
                              {showColoredBoxesInPrint ? (
                                <span className="inline-block border border-amber-400 bg-yellow-50/60 px-1 py-0.5 rounded-2xs">
                                  {formatTableNum(item.quantitySheets)}
                                </span>
                              ) : (
                                formatTableNum(item.quantitySheets)
                              )}
                            </td>
                            <td className="border border-black px-2 py-1.5 text-center font-bold font-times align-middle text-black" style={{ fontFamily: "'Times New Roman', Times, serif", verticalAlign: 'middle', fontSize: `${bodyFontSizePt}pt` }}>
                              {item.startSerial || '-'}
                            </td>
                            <td className="border border-black px-2 py-1.5 text-center font-bold font-times align-middle text-black" style={{ fontFamily: "'Times New Roman', Times, serif", verticalAlign: 'middle', fontSize: `${bodyFontSizePt}pt` }}>
                              {item.endSerial || '-'}
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
            </div>

            {/* Bottom Signatures & Approvals Section */}
            <div
              className="pt-0 font-siemreap"
              style={{
                marginTop: `${isEvisa ? sigTopSpacing : 2}px`,
                fontSize: `${bodyFontSizePt}pt`,
              }}
            >
              {/* Signatures Section (Receiver shifted to left, Giver shifted to right) */}
              <div
                className="flex justify-between items-start mb-8 text-center leading-relaxed"
                style={{
                  paddingLeft: `${sigShiftInset}px`,
                  paddingRight: `${sigShiftInset}px`,
                  fontSize: `${bodyFontSizePt}pt`,
                }}
              >
                {/* Left Column: Receiver */}
                <div className="space-y-0.5 flex flex-col items-center text-center">
                  <p className="whitespace-nowrap" style={{ fontSize: '10pt' }}>{lunarDate ? lunarDate.replace(/ព\.ស\.\s*/g, 'ព.ស') : ''}</p>
                  <p className="whitespace-nowrap" style={{ fontSize: '10pt' }}>រាជធានីភ្នំពេញ, ថ្ងៃទី{khmerDay} ខែ{khmerMonth} ឆ្នាំ{khmerYear}</p>
                  <p className="font-moul font-normal pt-1" style={{ fontSize: `${bodyFontSizePt}pt` }}>អ្នកទទួល</p>
                  <div className="h-20"></div>
                  {getReceiverDisplayName() && (
                    <p className="font-bold" style={{ color: '#c00000', fontSize: `${bodyFontSizePt}pt` }}>
                      {renderTextWithMoulName(getReceiverDisplayName())}
                    </p>
                  )}
                </div>

                {/* Right Column: Giver */}
                <div className="space-y-0.5 flex flex-col items-center text-center">
                  <p className="whitespace-nowrap" style={{ fontSize: '10pt' }}>{lunarDate ? lunarDate.replace(/ព\.ស\.\s*/g, 'ព.ស') : ''}</p>
                  <p className="whitespace-nowrap" style={{ fontSize: '10pt' }}>រាជធានីភ្នំពេញ, ថ្ងៃទី{khmerDay} ខែ{khmerMonth} ឆ្នាំ{khmerYear}</p>
                  <p className="font-moul font-normal pt-1" style={{ fontSize: `${bodyFontSizePt}pt` }}>អ្នកប្រគល់</p>
                  <div className="h-20"></div>
                  {getGiverDisplayName() && (
                    <p className="font-bold" style={{ color: '#c00000', fontSize: `${bodyFontSizePt}pt` }}>
                      {renderTextWithMoulName(getGiverDisplayName())}
                    </p>
                  )}
                </div>
              </div>

              {/* Bottom Center: Approval / Seen */}
              <div className="text-center leading-relaxed flex flex-col items-center" style={{ fontSize: `${bodyFontSizePt}pt` }}>
                <p className="font-moul font-normal" style={{ fontSize: `${bodyFontSizePt}pt` }}>បានឃើញ</p>
                <p className="whitespace-nowrap" style={{ fontSize: '10pt' }}>{lunarDate ? lunarDate.replace(/ព\.ស\.\s*/g, 'ព.ស') : ''}</p>
                <p className="whitespace-nowrap" style={{ fontSize: '10pt' }}>រាជធានីភ្នំពេញ, ថ្ងៃទី{khmerDay} ខែ{khmerMonth} ឆ្នាំ{khmerYear}</p>
                <p className="font-moul font-normal pt-1" style={{ fontSize: `${bodyFontSizePt}pt` }}>
                  <span className={showColoredBoxesInPrint ? 'border border-amber-400 bg-yellow-50/60 px-1 py-0.5 rounded-2xs' : ''}>
                    {approverTitle}
                  </span>
                </p>
                <div className="h-20"></div>
                {approverName && (
                  <p className="font-bold" style={{ color: '#c00000', fontSize: `${bodyFontSizePt}pt` }}>
                    {renderTextWithMoulName(approverName)}
                  </p>
                )}
              </div>
            </div>
          </div>
        </div>

        {/* Modal Bottom Footer Controls */}
        <div className="bg-gray-50 border-t border-gray-200 px-6 py-3 flex flex-wrap items-center justify-between gap-3 shrink-0 print:hidden">
          <span className="text-xs text-gray-500 font-medium">
            * អាចចុច "ទាញយក PDF" ដើម្បីទាញយកឯកសារជាទម្រង់ A4 (PDF)
          </span>
          <div className="flex items-center gap-2">
            <button
              onClick={onClose}
              className="bg-gray-200 hover:bg-gray-300 text-gray-800 px-4 py-1.5 rounded text-xs font-bold transition cursor-pointer"
            >
              បិទ (Close)
            </button>
            <button
              onClick={handleDownloadPdf}
              disabled={isDownloading}
              className="bg-emerald-600 hover:bg-emerald-700 text-white px-5 py-1.5 rounded text-xs font-bold transition flex items-center gap-1.5 cursor-pointer shadow-2xs disabled:opacity-50"
            >
              <Download className="w-4 h-4" />
              <span>{isDownloading ? 'កំពុងទាញយក...' : 'ទាញយក PDF'}</span>
            </button>
          </div>
        </div>
      </div>
    </div>
  );
};

export default StockPdfModal;
