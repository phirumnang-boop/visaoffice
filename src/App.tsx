import React, { 
  useState, 
  useEffect, 
  useRef, 
  useCallback,
  useMemo 
} from 'react';
import { 
  CategoriesState, 
  CategoryType, 
  Officer, 
  StockRecord, 
  UserAccount, 
  UserRole, 
  VisaRecord, 
  RefusalDeportationRecord 
} from './types';
import { apiService } from './services/apiService';
import { idbStorage } from './utils/idbStorage';
import {
  isOfficeOfficer,
  isTeamUser,
  STORAGE_KEY_HIDE_OFFICE_OFFICERS,
  API_KEY_HIDE_OFFICE_OFFICERS,
  DEFAULT_HIDE_OFFICE_OFFICERS_FROM_TEAMS,
  STORAGE_KEY_HIDE_OFFICER_MENU,
  API_KEY_HIDE_OFFICER_MENU,
  DEFAULT_HIDE_OFFICER_MENU_FROM_TEAMS,
  STORAGE_KEY_HIDE_USER_MENU,
  API_KEY_HIDE_USER_MENU,
  DEFAULT_HIDE_USER_MENU_FROM_TEAMS,
  STORAGE_KEY_HIDE_DAILY_OPS,
  API_KEY_HIDE_DAILY_OPS,
  DEFAULT_HIDE_DAILY_OPS_FROM_TEAMS,
  STORAGE_KEY_HIDE_REFUSAL,
  API_KEY_HIDE_REFUSAL,
  DEFAULT_HIDE_REFUSAL_FROM_TEAMS,
  STORAGE_KEY_HIDE_VISA_WORK,
  API_KEY_HIDE_VISA_WORK,
  DEFAULT_HIDE_VISA_WORK_FROM_TEAMS,
  STORAGE_KEY_HIDE_OFFICE_STOCK,
  API_KEY_HIDE_OFFICE_STOCK,
  DEFAULT_HIDE_OFFICE_STOCK_FROM_TEAMS,
  STORAGE_KEY_HIDE_SEARCH_CODE,
  API_KEY_HIDE_SEARCH_CODE,
  DEFAULT_HIDE_SEARCH_CODE_FROM_TEAMS,
  STORAGE_KEY_HIDE_CATEGORIES,
  API_KEY_HIDE_CATEGORIES,
  DEFAULT_HIDE_CATEGORIES_FROM_TEAMS,
} from './utils/officerAccess';
import {
  INITIAL_CATEGORIES,
  INITIAL_OFFICERS,
  INITIAL_USERS,
  INITIAL_VISA_RECORDS,
  INITIAL_STOCK_RECORDS,
} from './data/initialData';
import { Sidebar } from './components/Sidebar';
import { Header } from './components/Header';
import { Dashboard } from './components/Dashboard';
import { OfficerForm } from './components/OfficerForm';
import { OfficerList } from './components/OfficerList';
import { CategoryManager } from './components/CategoryManager';
import { UserManagement } from './components/UserManagement';
import { VisaForm } from './components/VisaForm';
import { VisaList } from './components/VisaList';
import { VisaSummary } from './components/VisaSummary';
import { VisaDateSearch } from './components/VisaDateSearch';
import { StockManager } from './components/StockManager';
import { StickerVisaStatus } from './components/StickerVisaStatus';
import { DailyTeamVisaOperations } from './components/DailyTeamVisaOperations';
import { VisaStubCollectionOperations } from './components/VisaStubCollectionOperations';
import { StickerTeamStockReport } from './components/StickerTeamStockReport';
import { RobokTotalStockWorkReport } from './components/RobokTotalStockWorkReport';
import { RobokDirectorStockWorkReport } from './components/RobokDirectorStockWorkReport';
import { YearlyTeamDistributionReport } from './components/YearlyTeamDistributionReport';
import { YearlyTeamUsageReport } from './components/YearlyTeamUsageReport';
import { YearlyK2ReceivedK1Report } from './components/YearlyK2ReceivedK1Report';
import { YearlyK2IssuedToTeamsReport } from './components/YearlyK2IssuedToTeamsReport';
import { DailyTeamStatisticsReport } from './components/DailyTeamStatisticsReport';
import { TeamVisaIssuanceStatsReport } from './components/TeamVisaIssuanceStatsReport';
import { DailyStickerVisaOperationReport } from './components/DailyStickerVisaOperationReport';
import { BranchChiefTeamVisaStatsReport } from './components/BranchChiefTeamVisaStatsReport';
import { TeamCodeSearch } from './components/TeamCodeSearch';
import { MonthlyVisaUsageComparisonChart } from './components/MonthlyVisaUsageComparisonChart';
import { EVisaTeamRobokReport } from './components/EVisaTeamRobokReport';
import { RefusalDeportationManager } from './components/RefusalDeportationManager';
import { RefusalNationalityData } from './components/RefusalNationalityData';
import { RefusalTotalSummaryReport } from './components/RefusalTotalSummaryReport';
import { useWorkspaceSettings } from './context/WorkspaceSettingsContext';
import { Lock, User, Eye, EyeOff, LogIn, ShieldAlert } from 'lucide-react';

export default function App() {
  const { scaleMode, isSidebarCollapsed, containerWidthMode } = useWorkspaceSettings();
  // Persistence state
  const [categories, setCategories] = useState<CategoriesState>(() => {
    const saved = localStorage.getItem('app_categories');
    if (!saved) return INITIAL_CATEGORIES;
    try {
      const parsed = JSON.parse(saved);
      const loadedCat: CategoriesState = {
        ...INITIAL_CATEGORIES,
        ...parsed,
        visaTeams: parsed.visaTeams || INITIAL_CATEGORIES.visaTeams,
        visaTeamsRobok: parsed.visaTeamsRobok || INITIAL_CATEGORIES.visaTeamsRobok,
        collectorRoles: parsed.collectorRoles || INITIAL_CATEGORIES.collectorRoles,
        teamPositions: parsed.teamPositions || INITIAL_CATEGORIES.teamPositions,
        organizations: parsed.organizations || INITIAL_CATEGORIES.organizations,
        refusalReasons: parsed.refusalReasons || INITIAL_CATEGORIES.refusalReasons,
      };

      // Ensure visaTeams and visaTeamsRobok have the complete 29 official pairs if legacy or mismatched
      const isPlaceholderList =
        !loadedCat.visaTeams ||
        loadedCat.visaTeams.length < 29 ||
        loadedCat.visaTeams.some((vt) => /^ក្រុមទី\s*\d+$/i.test(vt.name.trim())) ||
        !loadedCat.visaTeamsRobok ||
        loadedCat.visaTeamsRobok.length < 29 ||
        loadedCat.visaTeamsRobok.some(
          (vtr) => vtr.name === 'ព្រំដែន ត្រពាំងរូង' || vtr.name === 'ព្រំដែន ព្រែកបាក់'
        );

      if (isPlaceholderList) {
        loadedCat.visaTeams = INITIAL_CATEGORIES.visaTeams;
        loadedCat.visaTeamsRobok = INITIAL_CATEGORIES.visaTeamsRobok;
      }

      // Replace any legacy 'លេខាធិការ' with 'សេនាធិការ'
      Object.keys(loadedCat).forEach((key) => {
        const k = key as keyof CategoriesState;
        if (Array.isArray(loadedCat[k])) {
          loadedCat[k] = loadedCat[k].map((item) => ({
            ...item,
            name: item.name === 'លេខាធិការ' ? 'សេនាធិការ' : item.name,
          }));
        }
      });

      // Automatically upgrade legacy refusal reasons to the official ones matching the user's uploaded image
      if (loadedCat.refusalReasons) {
        loadedCat.refusalReasons = loadedCat.refusalReasons.map((item) => {
          const match = INITIAL_CATEGORIES.refusalReasons.find((r) => r.id === item.id);
          if (match) {
            return { ...item, name: match.name };
          }
          return item;
        });
      }

      return loadedCat;
    } catch {
      return INITIAL_CATEGORIES;
    }
  });

  const [officers, setOfficers] = useState<Officer[]>(() => {
    const saved = localStorage.getItem('app_officers');
    if (saved) {
      try {
        const parsed = JSON.parse(saved);
        if (Array.isArray(parsed)) {
          const dummyIds = new Set(['off-01', 'off-02', 'off-03', 'off-04', 'off-05']);
          const dummyNames = ['ចាន់ សុក្ខា', 'ចាន់ សុភា', 'ម៉េង វណ្ណា', 'វង្ស រតនា', 'រង្សី រតនា', 'កែវ បុប្ផា', 'ហេង សំណាង'];
          const cleaned = parsed.filter(
            (o: Officer) => !dummyIds.has(o.id) && !dummyNames.some((dn) => (o.nameKh || o.name || o.nameEn || '').includes(dn))
          );
          if (cleaned.length > 0) return cleaned;
        }
      } catch {
        // fallback
      }
    }
    return INITIAL_OFFICERS;
  });

  const [users, setUsers] = useState<UserAccount[]>(() => {
    const saved = localStorage.getItem('app_users');
    return saved ? JSON.parse(saved) : INITIAL_USERS;
  });

  const [visaRecords, setVisaRecords] = useState<VisaRecord[]>(() => {
    const saved = localStorage.getItem('app_visa_records');
    return saved ? JSON.parse(saved) : INITIAL_VISA_RECORDS;
  });
  const [stockRecords, setStockRecords] = useState<StockRecord[]>([]);
  const [stickerActualStock, setStickerActualStock] = useState<StockRecord[]>([]);

  const [refusalRecords, setRefusalRecords] = useState<RefusalDeportationRecord[]>(() => {
    try {
      const saved = localStorage.getItem('app_refusal_deportation_records_v1');
      return saved ? JSON.parse(saved) : [];
    } catch {
      return [];
    }
  });

  useEffect(() => {
    try {
      localStorage.setItem('app_refusal_deportation_records_v1', JSON.stringify(refusalRecords));
    } catch {}
  }, [refusalRecords]);

  const [currentUserIndex, setCurrentUserIndex] = useState(0);

  const [sessionUserRole, setSessionUserRole] = useState<UserRole>(() => {
    return (localStorage.getItem('app_session_user_role_v1') as UserRole) || 'Secondary';
  });

  const [isLoggedIn, setIsLoggedIn] = useState<boolean>(() => {
    const saved = localStorage.getItem('app_is_logged_in_v1');
    return saved !== 'false'; // default to true, unless explicitly logged out
  });
  const [loginUsername, setLoginUsername] = useState('');
  const [loginPassword, setLoginPassword] = useState('');
  const [showLoginPassword, setShowLoginPassword] = useState(false);
  const [loginError, setLoginError] = useState('');

  // Active Navigation Page
  const [activePage, setActivePage] = useState<string>('dashboard');
  const [activeCatType, setActiveCatType] = useState<CategoryType>('ranks');

  // Security Access Control: Hide Office Officers from Teams (ទិន្នន័យមន្ត្រី នៅការិយាល័យ មិនអនុញ្ញាតអោយ របស់ក្រុម មើលឃើញឡើយ)
  const [hideOfficeOfficersFromTeams, setHideOfficeOfficersFromTeams] = useState<boolean>(() => {
    const saved = localStorage.getItem(STORAGE_KEY_HIDE_OFFICE_OFFICERS);
    if (saved !== null) {
      return saved === 'true';
    }
    return DEFAULT_HIDE_OFFICE_OFFICERS_FROM_TEAMS;
  });

  // New Security Control: Hide "ព័ត៌មានមន្ត្រី" menu from Teams
  const [hideOfficerMenuFromTeams, setHideOfficerMenuFromTeams] = useState<boolean>(() => {
    const saved = localStorage.getItem(STORAGE_KEY_HIDE_OFFICER_MENU);
    if (saved !== null) {
      return saved === 'true';
    }
    return DEFAULT_HIDE_OFFICER_MENU_FROM_TEAMS;
  });

  // New Security Control: Hide "គណនីអ្នកប្រើប្រាស់" menu from Teams
  const [hideUserMenuFromTeams, setHideUserMenuFromTeams] = useState<boolean>(() => {
    const saved = localStorage.getItem(STORAGE_KEY_HIDE_USER_MENU);
    if (saved !== null) {
      return saved === 'true';
    }
    return DEFAULT_HIDE_USER_MENU_FROM_TEAMS;
  });

  const [hideDailyOpsFromTeams, setHideDailyOpsFromTeams] = useState<boolean>(() => {
    const saved = localStorage.getItem(STORAGE_KEY_HIDE_DAILY_OPS);
    if (saved !== null) {
      return saved === 'true';
    }
    return DEFAULT_HIDE_DAILY_OPS_FROM_TEAMS;
  });

  const [hideRefusalFromTeams, setHideRefusalFromTeams] = useState<boolean>(() => {
    const saved = localStorage.getItem(STORAGE_KEY_HIDE_REFUSAL);
    if (saved !== null) {
      return saved === 'true';
    }
    return DEFAULT_HIDE_REFUSAL_FROM_TEAMS;
  });

  const [hideVisaWorkFromTeams, setHideVisaWorkFromTeams] = useState<boolean>(() => {
    const saved = localStorage.getItem(STORAGE_KEY_HIDE_VISA_WORK);
    if (saved !== null) {
      return saved === 'true';
    }
    return DEFAULT_HIDE_VISA_WORK_FROM_TEAMS;
  });

  const [hideOfficeStockFromTeams, setHideOfficeStockFromTeams] = useState<boolean>(() => {
    const saved = localStorage.getItem(STORAGE_KEY_HIDE_OFFICE_STOCK);
    if (saved !== null) {
      return saved === 'true';
    }
    return DEFAULT_HIDE_OFFICE_STOCK_FROM_TEAMS;
  });

  const [hideSearchCodeFromTeams, setHideSearchCodeFromTeams] = useState<boolean>(() => {
    const saved = localStorage.getItem(STORAGE_KEY_HIDE_SEARCH_CODE);
    if (saved !== null) {
      return saved === 'true';
    }
    return DEFAULT_HIDE_SEARCH_CODE_FROM_TEAMS;
  });

  const [hideCategoriesFromTeams, setHideCategoriesFromTeams] = useState<boolean>(() => {
    const saved = localStorage.getItem(STORAGE_KEY_HIDE_CATEGORIES);
    if (saved !== null) {
      return saved === 'true';
    }
    return DEFAULT_HIDE_CATEGORIES_FROM_TEAMS;
  });

  // Sync settings from backend concurrently
  useEffect(() => {
    Promise.allSettled([
      apiService.getSetting<boolean>(API_KEY_HIDE_OFFICE_OFFICERS),
      apiService.getSetting<boolean>(API_KEY_HIDE_OFFICER_MENU),
      apiService.getSetting<boolean>(API_KEY_HIDE_USER_MENU),
      apiService.getSetting<boolean>(API_KEY_HIDE_DAILY_OPS),
      apiService.getSetting<boolean>(API_KEY_HIDE_REFUSAL),
      apiService.getSetting<boolean>(API_KEY_HIDE_VISA_WORK),
      apiService.getSetting<boolean>(API_KEY_HIDE_OFFICE_STOCK),
      apiService.getSetting<boolean>(API_KEY_HIDE_SEARCH_CODE),
      apiService.getSetting<boolean>(API_KEY_HIDE_CATEGORIES),
    ]).then(([rOffice, rOfficer, rUser, rDaily, rRefusal, rVisaWork, rOfficeStock, rSearchCode, rCategories]) => {
      if (rOffice.status === 'fulfilled' && typeof rOffice.value === 'boolean') {
        setHideOfficeOfficersFromTeams(rOffice.value);
        localStorage.setItem(STORAGE_KEY_HIDE_OFFICE_OFFICERS, String(rOffice.value));
      }
      if (rOfficer.status === 'fulfilled' && typeof rOfficer.value === 'boolean') {
        setHideOfficerMenuFromTeams(rOfficer.value);
        localStorage.setItem(STORAGE_KEY_HIDE_OFFICER_MENU, String(rOfficer.value));
      }
      if (rUser.status === 'fulfilled' && typeof rUser.value === 'boolean') {
        setHideUserMenuFromTeams(rUser.value);
        localStorage.setItem(STORAGE_KEY_HIDE_USER_MENU, String(rUser.value));
      }
      if (rDaily.status === 'fulfilled' && typeof rDaily.value === 'boolean') {
        setHideDailyOpsFromTeams(rDaily.value);
        localStorage.setItem(STORAGE_KEY_HIDE_DAILY_OPS, String(rDaily.value));
      }
      if (rRefusal.status === 'fulfilled' && typeof rRefusal.value === 'boolean') {
        setHideRefusalFromTeams(rRefusal.value);
        localStorage.setItem(STORAGE_KEY_HIDE_REFUSAL, String(rRefusal.value));
      }
      if (rVisaWork.status === 'fulfilled' && typeof rVisaWork.value === 'boolean') {
        setHideVisaWorkFromTeams(rVisaWork.value);
        localStorage.setItem(STORAGE_KEY_HIDE_VISA_WORK, String(rVisaWork.value));
      }
      if (rOfficeStock.status === 'fulfilled' && typeof rOfficeStock.value === 'boolean') {
        setHideOfficeStockFromTeams(rOfficeStock.value);
        localStorage.setItem(STORAGE_KEY_HIDE_OFFICE_STOCK, String(rOfficeStock.value));
      }
      if (rSearchCode.status === 'fulfilled' && typeof rSearchCode.value === 'boolean') {
        setHideSearchCodeFromTeams(rSearchCode.value);
        localStorage.setItem(STORAGE_KEY_HIDE_SEARCH_CODE, String(rSearchCode.value));
      }
      if (rCategories.status === 'fulfilled' && typeof rCategories.value === 'boolean') {
        setHideCategoriesFromTeams(rCategories.value);
        localStorage.setItem(STORAGE_KEY_HIDE_CATEGORIES, String(rCategories.value));
      }
    });
  }, []);

  const handleToggleHideOfficeOfficersFromTeams = async (enabled: boolean) => {
    setHideOfficeOfficersFromTeams(enabled);
    localStorage.setItem(STORAGE_KEY_HIDE_OFFICE_OFFICERS, String(enabled));
    try {
      await apiService.saveSetting(API_KEY_HIDE_OFFICE_OFFICERS, enabled);
    } catch {}
    showToast(
      enabled
        ? 'បានបើកដំណើរការ (SET ON)៖ ទិន្នន័យមន្ត្រីនៅការិយាល័យ មិនអនុញ្ញាតឱ្យរបស់ក្រុមមើលឃើញឡើយ'
        : 'បានបិទការកំណត់រឹតបន្តឹងទិន្នន័យមន្ត្រីការិយាល័យ',
      enabled ? 'success' : 'error'
    );
  };

  const handleToggleHideOfficerMenuFromTeams = async (enabled: boolean) => {
    setHideOfficerMenuFromTeams(enabled);
    localStorage.setItem(STORAGE_KEY_HIDE_OFFICER_MENU, String(enabled));
    try {
      await apiService.saveSetting(API_KEY_HIDE_OFFICER_MENU, enabled);
    } catch {}
    showToast(
      enabled
        ? 'បានបើកដំណើរការ (SET ON)៖ លាក់មឺនុយ «ព័ត៌មានមន្ត្រី» ពីក្រុមការងារ (User)'
        : 'បានអនុញ្ញាតឱ្យក្រុមការងារមើលឃើញមឺនុយ «ព័ត៌មានមន្ត្រី» ធម្មតា',
      enabled ? 'success' : 'error'
    );
  };

  const handleToggleHideUserMenuFromTeams = async (enabled: boolean) => {
    setHideUserMenuFromTeams(enabled);
    localStorage.setItem(STORAGE_KEY_HIDE_USER_MENU, String(enabled));
    try {
      await apiService.saveSetting(API_KEY_HIDE_USER_MENU, enabled);
    } catch {}
    showToast(
      enabled
        ? 'បានបើកដំណើរការ (SET ON)៖ លាក់មឺនុយ «គណនីអ្នកប្រើប្រាស់» ពីក្រុមការងារ (User)'
        : 'បានអនុញ្ញាតឱ្យក្រុមការងារមើលឃើញ និងប្រើប្រាស់មឺនុយ «គណនីអ្នកប្រើប្រាស់»',
      enabled ? 'success' : 'error'
    );
  };

  const handleToggleHideDailyOpsFromTeams = async (enabled: boolean) => {
    setHideDailyOpsFromTeams(enabled);
    localStorage.setItem(STORAGE_KEY_HIDE_DAILY_OPS, String(enabled));
    try {
      await apiService.saveSetting(API_KEY_HIDE_DAILY_OPS, enabled);
    } catch {}
    showToast(
      enabled
        ? 'បានបើកដំណើរការ (SET ON)៖ លាក់មឺនុយ «ប្រតិបត្តិការប្រចាំថ្ងៃតាមក្រុម» ពីក្រុមការងារ'
        : 'បានអនុញ្ញាតឱ្យក្រុមការងារមើលឃើញមឺនុយ «ប្រតិបត្តិការប្រចាំថ្ងៃតាមក្រុម»',
      enabled ? 'success' : 'error'
    );
  };

  const handleToggleHideRefusalFromTeams = async (enabled: boolean) => {
    setHideRefusalFromTeams(enabled);
    localStorage.setItem(STORAGE_KEY_HIDE_REFUSAL, String(enabled));
    try {
      await apiService.saveSetting(API_KEY_HIDE_REFUSAL, enabled);
    } catch {}
    showToast(
      enabled
        ? 'បានបើកដំណើរការ (SET ON)៖ លាក់មឺនុយ «បដិសេធនិងបញ្ជូនចេញ» ពីក្រុមការងារ'
        : 'បានអនុញ្ញាតឱ្យក្រុមការងារមើលឃើញមឺនុយ «បដិសេធនិងបញ្ជូនចេញ»',
      enabled ? 'success' : 'error'
    );
  };

  const handleToggleHideVisaWorkFromTeams = async (enabled: boolean) => {
    setHideVisaWorkFromTeams(enabled);
    localStorage.setItem(STORAGE_KEY_HIDE_VISA_WORK, String(enabled));
    try {
      await apiService.saveSetting(API_KEY_HIDE_VISA_WORK, enabled);
    } catch {}
    showToast(
      enabled
        ? 'បានបើកដំណើរការ (SET ON)៖ លាក់មឺនុយ «ការងារទិដ្ឋាការ» ពីក្រុមការងារ'
        : 'បានអនុញ្ញាតឱ្យក្រុមការងារមើលឃើញមឺនុយ «ការងារទិដ្ឋាការ»',
      enabled ? 'success' : 'error'
    );
  };

  const handleToggleHideOfficeStockFromTeams = async (enabled: boolean) => {
    setHideOfficeStockFromTeams(enabled);
    localStorage.setItem(STORAGE_KEY_HIDE_OFFICE_STOCK, String(enabled));
    try {
      await apiService.saveSetting(API_KEY_HIDE_OFFICE_STOCK, enabled);
    } catch {}
    showToast(
      enabled
        ? 'បានបើកដំណើរការ (SET ON)៖ លាក់មឺនុយ «ការងារស្តុកការិយាល័យ» ពីក្រុមការងារ'
        : 'បានអនុញ្ញាតឱ្យក្រុមការងារមើលឃើញមឺនុយ «ការងារស្តុកការិយាល័យ»',
      enabled ? 'success' : 'error'
    );
  };

  const handleToggleHideSearchCodeFromTeams = async (enabled: boolean) => {
    setHideSearchCodeFromTeams(enabled);
    localStorage.setItem(STORAGE_KEY_HIDE_SEARCH_CODE, String(enabled));
    try {
      await apiService.saveSetting(API_KEY_HIDE_SEARCH_CODE, enabled);
    } catch {}
    showToast(
      enabled
        ? 'បានបើកដំណើរការ (SET ON)៖ លាក់មឺនុយ «ស្វែងរកលេខកូដតាមក្រុម» ពីក្រុមការងារ'
        : 'បានអនុញ្ញាតឱ្យក្រុមការងារមើលឃើញមឺនុយ «ស្វែងរកលេខកូដតាមក្រុម»',
      enabled ? 'success' : 'error'
    );
  };

  const handleToggleHideCategoriesFromTeams = async (enabled: boolean) => {
    setHideCategoriesFromTeams(enabled);
    localStorage.setItem(STORAGE_KEY_HIDE_CATEGORIES, String(enabled));
    try {
      await apiService.saveSetting(API_KEY_HIDE_CATEGORIES, enabled);
    } catch {}
    showToast(
      enabled
        ? 'បានបើកដំណើរការ (SET ON)៖ លាក់មឺនុយ «បញ្ចូលប្រភេទ» ពីក្រុមការងារ'
        : 'បានអនុញ្ញាតឱ្យក្រុមការងារមើលឃើញមឺនុយ «បញ្ចូលប្រភេទ»',
      enabled ? 'success' : 'error'
    );
  };

  // Toast Notification State
  const [toastMessage, setToastMessage] = useState<{
    text: string;
    type: 'success' | 'error';
  } | null>(null);

  // Safe localStorage helper to prevent QuotaExceededError crashes and main-thread JSON.stringify freezes
  const safeStorageSet = (key: string, value: any) => {
    try {
      if (Array.isArray(value) && value.length > 300) {
        // Large arrays (>300 items) are stored in high-capacity IndexedDB + Server; remove oversized localStorage copy to free quota and prevent UI lag
        localStorage.removeItem(key);
        return;
      }
      localStorage.setItem(key, typeof value === 'string' ? value : JSON.stringify(value));
    } catch (e) {
      // LocalStorage limit exceeded, IndexedDB handles large volumes safely
    }
  };

  // Guard ref so we never overwrite IndexedDB with initial empty array on mount
  const isInitialLoadDoneRef = useRef<boolean>(false);

  // Debounced non-blocking background sync helper to prevent UI thread lag on menu switching
  const syncTimerRef = useRef<{ [key: string]: NodeJS.Timeout }>({});

  const debouncedSync = (key: string, fn: () => void, delay = 150) => {
    if (syncTimerRef.current[key]) {
      clearTimeout(syncTimerRef.current[key]);
    }
    syncTimerRef.current[key] = setTimeout(fn, delay);
  };

  // Sync to IndexedDB and LocalStorage safely ONLY after initial load is done
  useEffect(() => {
    if (!isInitialLoadDoneRef.current) return;
    debouncedSync('categories', () => {
      idbStorage.setItem('categories', 'data', categories);
      safeStorageSet('app_categories', categories);
    });
  }, [categories]);

  useEffect(() => {
    if (!isInitialLoadDoneRef.current) return;
    debouncedSync('officers', () => {
      idbStorage.setItem('officers', 'data', officers);
      safeStorageSet('app_officers', officers);
    });
  }, [officers]);

  useEffect(() => {
    if (!isInitialLoadDoneRef.current) return;
    debouncedSync('users', () => {
      idbStorage.setItem('users', 'data', users);
      safeStorageSet('app_users', users);
    });
  }, [users]);

  useEffect(() => {
    if (!isInitialLoadDoneRef.current) return;
    debouncedSync('visa_records', () => {
      idbStorage.setItem('visa_records', 'data', visaRecords);
      safeStorageSet('app_visa_records', visaRecords);
    });
  }, [visaRecords]);

  useEffect(() => {
    if (!isInitialLoadDoneRef.current) return;
    debouncedSync('stock_records', () => {
      idbStorage.setItem('stock_records', 'data', stockRecords);
      safeStorageSet('app_stock_records', stockRecords);
    });
  }, [stockRecords]);

  useEffect(() => {
    if (!isInitialLoadDoneRef.current) return;
    debouncedSync('sticker_actual_stock', () => {
      idbStorage.setItem('sticker_actual_stock', 'data', stickerActualStock);
      safeStorageSet('app_sticker_actual_stock_records', stickerActualStock);
    });
  }, [stickerActualStock]);

  // Cloud SQL Database Connection & Sync state
  const [isCloudConnected, setIsCloudConnected] = useState<boolean>(true);
  const [isCloudSyncing, setIsCloudSyncing] = useState<boolean>(false);

  // Initial High-Capacity IndexedDB & Cloud synchronization on app load
  useEffect(() => {
    let isMounted = true;
    const initializeData = async () => {
      try {
        setIsCloudSyncing(true);

        // 1. First restore from high-capacity IndexedDB
        const [idbStock, idbSticker, idbVisa] = await Promise.all([
          idbStorage.getItem<StockRecord[]>('stock_records', 'data'),
          idbStorage.getItem<StockRecord[]>('sticker_actual_stock', 'data'),
          idbStorage.getItem<VisaRecord[]>('visa_records', 'data'),
        ]);

        // 2. Fetch Server Records (with null as fallback if request fails due to offline/server down)
        const fetchServerCollection = async (url: string) => {
          try {
            const res = await fetch(url);
            if (res.ok) {
              const json = await res.json();
              return Array.isArray(json?.data) ? json.data : (Array.isArray(json) ? json : []);
            }
          } catch (e) {
            console.warn(`Fetch error for ${url}:`, e);
          }
          return null;
        };

        const [cloudStock, cloudStickerActual, cloudVisa, cloudDailyOps] = await Promise.all([
          fetchServerCollection('/api/stock-records'),
          fetchServerCollection('/api/sticker-actual-stock'),
          fetchServerCollection('/api/visa-records'),
          fetchServerCollection('/api/daily-team-operations'),
        ]);

        if (isMounted) {
          // Robust Union-Merge for Stock Records (Both Sticker and EVisa)
          try {
            const stockMap = new Map<string, StockRecord>();

            // 1. Add localStorage fallback items first
            try {
              const saved = localStorage.getItem('app_stock_records');
              if (saved) {
                const parsed = JSON.parse(saved);
                if (Array.isArray(parsed)) {
                  parsed.forEach((r: StockRecord) => {
                    if (r && r.id) stockMap.set(r.id, r);
                  });
                }
              }
            } catch {}

            // 2. Add high-capacity IndexedDB items
            if (Array.isArray(idbStock)) {
              idbStock.forEach((r) => {
                if (r && r.id) stockMap.set(r.id, r);
              });
            }

            // 3. Add Server items & clean deleted items
            if (cloudStock !== null) {
              const cloudIds = new Set(cloudStock.map((r) => r.id));
              // Wipes out deleted records from local cache if we successfully retrieved server database list
              for (const localId of stockMap.keys()) {
                if (!cloudIds.has(localId)) {
                  stockMap.delete(localId);
                }
              }
              // Merge/update newest server records
              cloudStock.forEach((r) => {
                if (r && r.id) {
                  const existing = stockMap.get(r.id);
                  stockMap.set(r.id, {
                    ...r,
                    isImported: existing?.isImported ?? r.isImported,
                    source: existing?.source ?? r.source,
                  });
                }
              });
            }

            const resolvedStock = Array.from(stockMap.values());

            setStockRecords(resolvedStock);
            idbStorage.setItem('stock_records', 'data', resolvedStock).catch(() => {});
            safeStorageSet('app_stock_records', resolvedStock);
          } catch (e) {
            console.error('Error restoring stock records:', e);
          }

          // Dedicated Union-Merge for Sticker Actual Stock
          try {
            const stickerMap = new Map<string, StockRecord>();

            // 1. Add localStorage fallback items
            try {
              const localSaved = localStorage.getItem('app_sticker_actual_stock_records');
              if (localSaved) {
                const parsed = JSON.parse(localSaved);
                if (Array.isArray(parsed)) {
                  parsed.forEach((r: StockRecord) => {
                    if (r && r.id) stickerMap.set(r.id, r);
                  });
                }
              }
            } catch {}

            // 2. Add high-capacity IndexedDB items
            if (Array.isArray(idbSticker)) {
              idbSticker.forEach((r) => {
                if (r && r.id) stickerMap.set(r.id, r);
              });
            }

            // 3. Add Dedicated Server items & clean deleted items
            if (cloudStickerActual !== null) {
              const cloudIds = new Set(cloudStickerActual.map((r) => r.id));
              for (const localId of stickerMap.keys()) {
                if (!cloudIds.has(localId)) {
                  stickerMap.delete(localId);
                }
              }
              cloudStickerActual.forEach((r) => {
                if (r && r.id) {
                  const existing = stickerMap.get(r.id);
                  stickerMap.set(r.id, {
                    ...r,
                    isImported: existing?.isImported ?? r.isImported,
                    source: existing?.source ?? r.source,
                  });
                }
              });
            }

            const resolvedSticker = Array.from(stickerMap.values());
            setStickerActualStock(resolvedSticker);
            idbStorage.setItem('sticker_actual_stock', 'data', resolvedSticker).catch(() => {});
            safeStorageSet('app_sticker_actual_stock_records', resolvedSticker);
          } catch (e) {
            console.error('Error restoring sticker actual stock:', e);
          }

          // Robust Union-Merge for Visa Records
          try {
            const visaMap = new Map<string, VisaRecord>();
            if (Array.isArray(idbVisa)) {
              idbVisa.forEach((r) => {
                if (r && r.id) visaMap.set(r.id, r);
              });
            }
            if (cloudVisa !== null) {
              const cloudIds = new Set(cloudVisa.map((r) => r.id));
              for (const localId of visaMap.keys()) {
                if (!cloudIds.has(localId)) {
                  visaMap.delete(localId);
                }
              }
              cloudVisa.forEach((r) => {
                if (r && r.id) {
                  visaMap.set(r.id, r);
                }
              });
            }

            const resolvedVisa = Array.from(visaMap.values());

            setVisaRecords(resolvedVisa);
            idbStorage.setItem('visa_records', 'data', resolvedVisa).catch(() => {});
            safeStorageSet('app_visa_records', resolvedVisa);
          } catch (e) {
            console.error('Error restoring visa records:', e);
          }

          // Robust Union-Merge for User Accounts
          try {
            const userMap = new Map<string, UserAccount>();
            // 1. Add existing local state/localStorage users
            users.forEach((u) => {
              if (u && u.id) userMap.set(u.id, u);
            });
            // 2. Fetch cloud users
            const cloudUsers = await apiService.getCloudUsers().catch(() => []);
            if (Array.isArray(cloudUsers)) {
              cloudUsers.forEach((u) => {
                if (u && u.id) {
                  userMap.set(u.id, u);
                }
              });
            }
            const resolvedUsers = Array.from(userMap.values());
            setUsers(resolvedUsers);
            idbStorage.setItem('users', 'data', resolvedUsers).catch(() => {});
            safeStorageSet('app_users', resolvedUsers);
          } catch (e) {
            console.error('Error restoring user accounts:', e);
          }

          // Restore and Sync Daily Team Operations
          try {
            if (cloudDailyOps !== null) {
              localStorage.setItem('app_daily_team_operations_v5', JSON.stringify(cloudDailyOps));
              window.dispatchEvent(new Event('storage'));
            }
          } catch (e) {
            console.error('Error syncing daily team operations:', e);
          }

          setIsCloudConnected(true);
        }
      } catch (err) {
        console.warn('Sync initialization notice:', err);
        if (isMounted) setIsCloudConnected(false);
      } finally {
        if (isMounted) {
          isInitialLoadDoneRef.current = true;
          setIsCloudSyncing(false);
        }
      }
    };

    initializeData();
    return () => {
      isMounted = false;
    };
  }, []);

  const handleAddStockRecord = (record: StockRecord) => {
    const taggedRecord: StockRecord = {
      ...record,
      source: record.source || 'form',
    };
    setStockRecords((prev) => {
      const next = [taggedRecord, ...prev];
      idbStorage.setItem('stock_records', 'data', next).catch(() => {});
      if (next.length < 2000) {
        safeStorageSet('app_stock_records', next);
      }
      return next;
    });

    // When entering stock from K1 (receiving stickers from K1), the user explicitly instructed:
    // "ហើយនៅពេលខ្ញុំបញ្ចូលស្តុកពីក១ (សន្លឹកទិដ្ឋាការ) សូមកុំភ្លេច ចូលតារាងស្តុកជាក់ស្តែងសម្រាប់ទុកធ្វើការបើកផ្តល់តាមក្រុម ផង"
    if (taggedRecord.stockType === 'sticker') {
      const isK1Inflow =
        taggedRecord.operationType === 'openK1' ||
        taggedRecord.operationType === 'receive' ||
        taggedRecord.operationType === 'received' ||
        taggedRecord.sourceFrom === 'ក១' ||
        taggedRecord.sourceFrom?.includes('ក១');

      if (isK1Inflow) {
        setStickerActualStock((prev) => {
          const next = [taggedRecord, ...prev];
          idbStorage.setItem('sticker_actual_stock', 'data', next).catch(() => {});
          if (next.length < 2000) {
            safeStorageSet('app_sticker_actual_stock_records', next);
          }
          return next;
        });
        apiService.saveStickerActualStock(taggedRecord);
      }
    }
    // Async save to Server
    apiService.saveStockRecord(taggedRecord);
  };

  const handleBatchImportStockRecords = (newRecords: StockRecord[]) => {
    const taggedRecords = newRecords.map((r) => ({
      ...r,
      isImported: true,
      source: 'import' as const,
    }));
    setStockRecords((prev) => {
      const prevIds = new Set(prev.map((r) => r.id));
      const nonDuplicates = taggedRecords.filter((r) => !prevIds.has(r.id));
      const updated = [...nonDuplicates, ...prev];
      idbStorage.setItem('stock_records', 'data', updated).catch(() => {});
      if (updated.length < 2000) {
        safeStorageSet('app_stock_records', updated);
      }
      return updated;
    });

    // Async bulk save to Server
    apiService.bulkSaveStockRecords(taggedRecords);
  };

  const handleDeleteStockRecord = (id: string) => {
    setStockRecords((prev) => {
      const updated = prev.filter((r) => r.id !== id);
      idbStorage.setItem('stock_records', 'data', updated).catch(() => {});
      if (updated.length < 2000) {
        safeStorageSet('app_stock_records', updated);
      }
      return updated;
    });
    // Async delete from Server
    apiService.deleteStockRecord(id);
  };

  const handleDeleteBatchStockRecords = (ids: string[], stockType?: string, operationType?: string) => {
    const idSet = new Set(ids);
    setStockRecords((prev) => {
      const updated = prev.filter((r) => !idSet.has(r.id));
      idbStorage.setItem('stock_records', 'data', updated).catch(() => {});
      if (updated.length < 2000) {
        safeStorageSet('app_stock_records', updated);
      }
      return updated;
    });
    // Async bulk delete from Server
    apiService.bulkDeleteStockRecords(ids, stockType, operationType);
  };

  const handleClearStickerStock = () => {
    const idsToDelete = stockRecords.filter((r) => r.stockType === 'sticker').map((r) => r.id);
    setStockRecords((prev) => {
      const updated = prev.filter((r) => r.stockType !== 'sticker');
      idbStorage.setItem('stock_records', 'data', updated).catch(() => {});
      if (updated.length < 2000) {
        safeStorageSet('app_stock_records', updated);
      }
      return updated;
    });
    if (idsToDelete.length > 0) {
      apiService.bulkDeleteStockRecords(idsToDelete);
    }
    showToast('បានសម្អាតទិន្នន័យស្តុកសន្លឹកទិដ្ឋាការរួចរាល់! (មិនប៉ះពាល់ស្តុកជាក់ស្តែងសម្រាប់បើកផ្តល់តាមក្រុម)', 'success');
  };

  // Dedicated handlers for "ស្តុកជាក់ស្តែងសម្រាប់បើកផ្តល់តាមក្រុម" (StickerVisaStatus)
  const handleBatchImportStickerActualStock = (newRecords: StockRecord[], overwrite: boolean = false) => {
    const taggedRecords = newRecords.map((r) => ({
      ...r,
      isImported: true,
      source: 'import' as const,
    }));
    setStickerActualStock((prev) => {
      const updated = overwrite ? taggedRecords : (() => {
        const prevIds = new Set(prev.map((r) => r.id));
        const nonDuplicates = taggedRecords.filter((r) => !prevIds.has(r.id));
        return [...nonDuplicates, ...prev];
      })();
      idbStorage.setItem('sticker_actual_stock', 'data', updated).catch(() => {});
      if (updated.length < 2000) {
        safeStorageSet('app_sticker_actual_stock_records', updated);
      }
      return updated;
    });
    // Also mirror to general stockRecords so office reports have the records
    setStockRecords((prev) => {
      const prevIds = new Set(prev.map((r) => r.id));
      const nonDuplicates = taggedRecords.filter((r) => !prevIds.has(r.id));
      if (nonDuplicates.length === 0) return prev;
      const updated = [...nonDuplicates, ...prev];
      idbStorage.setItem('stock_records', 'data', updated).catch(() => {});
      if (updated.length < 2000) {
        safeStorageSet('app_stock_records', updated);
      }
      return updated;
    });
    apiService.bulkSaveStockRecords(taggedRecords);
    apiService.bulkSaveStickerActualStock(taggedRecords, overwrite);
  };

  // When adding stock from K1 (receiving stickers from K1):
  // 1. Adds to stickerActualStock so that actual stock available for teams increases
  // 2. Also adds to general stockRecords so the K1 receipt is officially logged in StockManager & reports
  const handleAddK1StockToActualStock = (record: StockRecord) => {
    const tagged: StockRecord = {
      ...record,
      stockType: 'sticker',
      operationType: 'openK1',
      sourceFrom: record.sourceFrom || 'ក១',
      source: record.source || 'form',
    };
    setStickerActualStock((prev) => {
      const next = [tagged, ...prev];
      idbStorage.setItem('sticker_actual_stock', 'data', next).catch(() => {});
      if (next.length < 2000) {
        safeStorageSet('app_sticker_actual_stock_records', next);
      }
      return next;
    });
    setStockRecords((prev) => {
      const next = [tagged, ...prev];
      idbStorage.setItem('stock_records', 'data', next).catch(() => {});
      if (next.length < 2000) {
        safeStorageSet('app_stock_records', next);
      }
      return next;
    });
    apiService.saveStickerActualStock(tagged);
    apiService.saveStockRecord(tagged);
    showToast(`បានបញ្ចូលស្តុកពី ក១ (${tagged.visaType}: ${tagged.totalSheets || (tagged.quantityBundles ? tagged.quantityBundles * 50 : 0)} សន្លឹក) ចូលក្នុងស្តុកជាក់ស្តែង K2 ជោគជ័យ!`, 'success');
  };

  const handleReturnIssuedStock = (id: string) => {
    setStickerActualStock((prev) => {
      const updated = prev.filter((r) => r.id !== id);
      idbStorage.setItem('sticker_actual_stock', 'data', updated).catch(() => {});
      if (updated.length < 2000) {
        safeStorageSet('app_sticker_actual_stock_records', updated);
      }
      return updated;
    });
    setStockRecords((prev) => {
      const updated = prev.filter((r) => r.id !== id);
      idbStorage.setItem('stock_records', 'data', updated).catch(() => {});
      if (updated.length < 2000) {
        safeStorageSet('app_stock_records', updated);
      }
      return updated;
    });
    apiService.deleteStickerActualStock(id);
    apiService.deleteStockRecord(id);
    showToast('បានបង្វិលទិន្នន័យបើកផ្តល់ត្រឡប់ចូលស្តុកជាក់ស្តែង និងទិន្នន័យ (cEA) រួចរាល់!', 'success');
  };

  const handleReturnIssuedStockBatch = (ids: string[]) => {
    const idSet = new Set(ids);
    setStickerActualStock((prev) => {
      const updated = prev.filter((r) => !idSet.has(r.id));
      idbStorage.setItem('sticker_actual_stock', 'data', updated).catch(() => {});
      if (updated.length < 2000) {
        safeStorageSet('app_sticker_actual_stock_records', updated);
      }
      return updated;
    });
    setStockRecords((prev) => {
      const updated = prev.filter((r) => !idSet.has(r.id));
      idbStorage.setItem('stock_records', 'data', updated).catch(() => {});
      if (updated.length < 2000) {
        safeStorageSet('app_stock_records', updated);
      }
      return updated;
    });
    apiService.bulkDeleteStickerActualStock(ids);
    apiService.bulkDeleteStockRecords(ids);
    showToast(`បានបង្វិលទិន្នន័យបើកផ្តល់ចំនួន ${ids.length} ត្រឡប់ចូលស្តុកជាក់ស្តែង និងទិន្នន័យ (cEA) រួចរាល់!`, 'success');
  };

  const handleReturnIssuedStockByVisaType = (visaType: string) => {
    const targetType = (visaType || '').trim().toUpperCase();
    const idsToReturn = stickerActualStock
      .filter((r) => (r.visaType || '').trim().toUpperCase() === targetType && r.operationType === 'issueTeam')
      .map((r) => r.id);
    if (idsToReturn.length > 0) {
      handleReturnIssuedStockBatch(idsToReturn);
    }
  };

  const handleClearAllActualStock = () => {
    setStickerActualStock([]);
    idbStorage.setItem('sticker_actual_stock', 'data', []).catch(() => {});
    safeStorageSet('app_sticker_actual_stock_records', []);
    apiService.clearAllStickerActualStock();
    showToast('បានលុបទិន្នន័យក្នុងតារាងស្តុកជាក់ស្តែងទាំងអស់រួចរាល់! លោកអ្នកអាចធ្វើការ Import ទិន្នន័យថ្មីបាន (ទិន្នន័យសន្លឹកទិដ្ឋាការនៅរក្សាទុកដដែល)', 'success');
  };

  const handleDeleteActualStockRecord = (id: string) => {
    setStickerActualStock((prev) => {
      const updated = prev.filter((r) => r.id !== id);
      idbStorage.setItem('sticker_actual_stock', 'data', updated).catch(() => {});
      if (updated.length < 2000) {
        safeStorageSet('app_sticker_actual_stock_records', updated);
      }
      return updated;
    });
    apiService.deleteStickerActualStock(id);
    showToast('បានលុបទិន្នន័យចេញពីស្តុកជាក់ស្តែងរួចរាល់! (ទិន្នន័យសន្លឹកទិដ្ឋាការនៅរក្សាទុកដដែល)', 'success');
  };

  const handleDeleteBatchActualStockRecords = (ids: string[]) => {
    const idSet = new Set(ids);
    setStickerActualStock((prev) => {
      const updated = prev.filter((r) => !idSet.has(r.id));
      idbStorage.setItem('sticker_actual_stock', 'data', updated).catch(() => {});
      if (updated.length < 2000) {
        safeStorageSet('app_sticker_actual_stock_records', updated);
      }
      return updated;
    });
    apiService.bulkDeleteStickerActualStock(ids);
    showToast(`បានលុបទិន្នន័យចំនួន ${ids.length} ចេញពីស្តុកជាក់ស្តែងរួចរាល់! (ទិន្នន័យសន្លឹកទិដ្ឋាការនៅរក្សាទុកដដែល)`, 'success');
  };

  const handleDeleteActualStockByRange = (visaType: string, startSerial: string, endSerial: string) => {
    const cleanVt = (visaType || '').trim().toUpperCase();
    const parseSerial = (serial: string | undefined) => {
      if (!serial) return null;
      const str = serial.trim();
      const match = str.match(/^([A-Za-z\s_-]*)(\d+)$/);
      if (!match) return null;
      return { prefix: match[1] || '', num: parseInt(match[2], 10) };
    };

    const pStart = parseSerial(startSerial);
    const pEnd = parseSerial(endSerial);

    let nextList: StockRecord[] = [];
    setStickerActualStock((prev) => {
      const updated = prev.filter((r) => {
        const rVt = (r.visaType || '').trim().toUpperCase();
        if (rVt !== cleanVt) return true;

        if (pStart && pEnd && r.startSerial) {
          const rStart = parseSerial(r.startSerial);
          const rEnd = r.endSerial ? parseSerial(r.endSerial) : rStart;
          if (rStart && rEnd) {
            // Check if this record's range falls within the deleted range
            if (rStart.num >= pStart.num && rEnd.num <= pEnd.num) {
              return false;
            }
          }
        } else if (r.startSerial === startSerial && r.endSerial === endSerial) {
          return false;
        }
        return true;
      });
      nextList = updated;
      idbStorage.setItem('sticker_actual_stock', 'data', updated).catch(() => {});
      if (updated.length < 2000) {
        safeStorageSet('app_sticker_actual_stock_records', updated);
      }
      return updated;
    });
    apiService.bulkSaveStickerActualStock(nextList, true);
    showToast(`បានលុបទិន្នន័យចន្លោះលេខ ${startSerial} - ${endSerial} ចេញពីស្តុកជាក់ស្តែងរួចរាល់! (ទិន្នន័យសន្លឹកទិដ្ឋាការនៅរក្សាទុកដដែល)`, 'success');
  };

  // When issuing to team from actual stock:
  // 1. Adds to stickerActualStock so that interval math deducts available sheets/books
  // 2. Also adds to general stockRecords so the team issuance is officially logged in StockManager & reports
  const handleIssueToTeamFromActualStock = (record: StockRecord) => {
    const tagged: StockRecord = {
      ...record,
      source: record.source || 'form',
    };
    setStickerActualStock((prev) => {
      const next = [tagged, ...prev];
      idbStorage.setItem('sticker_actual_stock', 'data', next).catch(() => {});
      if (next.length < 2000) {
        safeStorageSet('app_sticker_actual_stock_records', next);
      }
      return next;
    });
    setStockRecords((prev) => {
      const next = [tagged, ...prev];
      idbStorage.setItem('stock_records', 'data', next).catch(() => {});
      if (next.length < 2000) {
        safeStorageSet('app_stock_records', next);
      }
      return next;
    });
    apiService.saveStickerActualStock(tagged);
    apiService.saveStockRecord(tagged);
  };

  const handleUpdateStockRecord = (updatedRecord: StockRecord) => {
    setStockRecords((prev) => {
      const updated = prev.map((r) => {
        if (r.id === updatedRecord.id) {
          return {
            ...r,
            ...updatedRecord,
            isImported: r.isImported !== undefined ? r.isImported : updatedRecord.isImported,
          };
        }
        return r;
      });
      idbStorage.setItem('stock_records', 'data', updated).catch(() => {});
      if (updated.length < 2000) {
        safeStorageSet('app_stock_records', updated);
      }
      return updated;
    });
    if (updatedRecord.stockType === 'sticker') {
      setStickerActualStock((prev) => {
        let found = false;
        const updated = prev.map((r) => {
          if (r.id === updatedRecord.id) {
            found = true;
            return {
              ...r,
              ...updatedRecord,
              isImported: r.isImported !== undefined ? r.isImported : updatedRecord.isImported,
            };
          }
          return r;
        });
        if (found) {
          idbStorage.setItem('sticker_actual_stock', 'data', updated).catch(() => {});
          if (updated.length < 2000) {
            safeStorageSet('app_sticker_actual_stock_records', updated);
          }
          apiService.saveStickerActualStock(updatedRecord);
        }
        return updated;
      });
    }
    apiService.saveStockRecord(updatedRecord);
  };

  const showToast = (text: string, type: 'success' | 'error' = 'success') => {
    setToastMessage({ text, type });
    setTimeout(() => {
      setToastMessage(null);
    }, 3200);
  };

  const currentUser = users[currentUserIndex] || users[0] || {
    id: 'usr-default',
    username: 'admin',
    role: 'Secondary',
    permissions: {
      refusal: true,
      userMenu: true,
      officerMenu: true,
      visaWork: true,
      officeStock: true,
      categories: true,
      dailyOps: true,
      searchCode: true,
      userList: true
    }
  };
  const mainScrollRef = useRef<HTMLDivElement>(null);

  // Navigation handler
  const handleNavigate = useCallback((page: string, catType?: CategoryType) => {
    const isUserPageRestricted = hideUserMenuFromTeams;
    if (
      currentUser.role !== 'Secondary' && currentUser.role !== 'User (ការិយាល័យ)' &&
      (
        (page === 'createUser' && isUserPageRestricted) ||
        (page === 'userList' && isUserPageRestricted) ||
        page === 'stockStubProposalSection' ||
        page === 'stockStubProposalOffice'
      )
    ) {
      setActivePage('dashboard');
      return;
    }
    setActivePage(page);
    if (catType) {
      setActiveCatType(catType);
    }
    if (mainScrollRef.current) {
      mainScrollRef.current.scrollTop = 0;
    } else {
      window.scrollTo(0, 0);
    }
  }, [hideUserMenuFromTeams, currentUser.role]);

  // Category handlers
  const handleAddCategory = (type: CategoryType, name: string) => {
    const newItem = {
      id: `${type}-${Date.now()}-${Math.random().toString(36).substring(2, 7)}`,
      name,
      createdAt: new Date().toISOString(),
    };

    setCategories((prev) => ({
      ...prev,
      [type]: [...prev[type], newItem],
    }));
  };

  const handleDeleteCategory = (type: CategoryType, id: string) => {
    setCategories((prev) => ({
      ...prev,
      [type]: prev[type].filter((item) => item.id !== id),
    }));
  };

  const handleUpdateCategory = (type: CategoryType, id: string, newName: string) => {
    setCategories((prev) => ({
      ...prev,
      [type]: prev[type].map((item) => (item.id === id ? { ...item, name: newName } : item)),
    }));
  };

  const handleResetVisaTeams = () => {
    setCategories((prev) => ({
      ...prev,
      visaTeams: INITIAL_CATEGORIES.visaTeams,
      visaTeamsRobok: INITIAL_CATEGORIES.visaTeamsRobok,
    }));
    showToast('បានស្ដារទិន្នន័យក្រុមផ្តល់ទិដ្ឋាការដើម (២៩ ក្រុមផ្លូវការ) រួចរាល់!', 'success');
  };

  // Officer handlers
  const handleAddOfficer = (newOfficerData: Omit<Officer, 'id' | 'createdAt'>) => {
    const isTeam = isTeamUser(currentUser.role, currentUser.assignedTeam);
    const sanitizedData = { ...newOfficerData };
    if (isTeam && hideOfficeOfficersFromTeams && sanitizedData.officeWork) {
      delete (sanitizedData as any).officeWork;
    }

    const newOfficer: Officer = {
      ...sanitizedData,
      id: `of-${Date.now()}-${Math.random().toString(36).substring(2, 7)}`,
      createdAt: new Date().toISOString(),
    };

    setOfficers((prev) => [newOfficer, ...prev]);
  };

  const handleUpdateOfficer = (id: string, updated: Partial<Officer>) => {
    setOfficers((prev) =>
      prev.map((of) => (of.id === id ? { ...of, ...updated } : of))
    );
  };

  const handleDeleteOfficer = (id: string) => {
    setOfficers((prev) => prev.filter((of) => of.id !== id));
  };

  // User account handlers
  const handleAddUser = (newUserData: Omit<UserAccount, 'id' | 'createdAt'>) => {
    const newUser: UserAccount = {
      ...newUserData,
      id: `usr-${Date.now()}-${Math.random().toString(36).substring(2, 7)}`,
      createdAt: new Date().toISOString(),
    };

    setUsers((prev) => [...prev, newUser]);
    apiService.saveCloudUser(newUser); // Real-time Firestore Sync
  };

  const handleUpdateUser = (id: string, updated: Partial<UserAccount>) => {
    setUsers((prev) => {
      const next = prev.map((u) => (u.id === id ? { ...u, ...updated } : u));
      const updatedUser = next.find((u) => u.id === id);
      if (updatedUser) {
        apiService.saveCloudUser(updatedUser); // Real-time Firestore Sync
      }
      return next;
    });
  };

  const handleDeleteUser = (id: string) => {
    setUsers((prev) => prev.filter((u) => u.id !== id));
    apiService.deleteCloudUser(id); // Real-time Firestore Sync
  };

  // Visa record handlers
  const handleAddVisaRecord = (newRecord: Omit<VisaRecord, 'id' | 'createdAt'>) => {
    const record: VisaRecord = {
      ...newRecord,
      id: `visa-${Date.now()}-${Math.random().toString(36).substring(2, 7)}`,
      createdAt: new Date().toISOString(),
      source: 'form',
    };
    setVisaRecords((prev) => {
      const next = [record, ...prev];
      idbStorage.setItem('visa_records', 'data', next);
      return next;
    });
    apiService.saveVisaRecord(record);
  };

  const handleUpdateVisaRecord = (updated: VisaRecord) => {
    setVisaRecords((prev) => {
      const next = prev.map((r) => {
        if (r.id === updated.id) {
          return {
            ...r,
            ...updated,
            isImported: r.isImported !== undefined ? r.isImported : updated.isImported,
          };
        }
        return r;
      });
      idbStorage.setItem('visa_records', 'data', next);
      return next;
    });
    apiService.saveVisaRecord(updated);
  };

  const handleDeleteVisaRecord = (id: string) => {
    setVisaRecords((prev) => {
      const next = prev.filter((r) => r.id !== id);
      idbStorage.setItem('visa_records', 'data', next);
      return next;
    });
    apiService.deleteVisaRecord(id);
  };

  const handleBatchImportVisaRecords = (recordsToImport: Omit<VisaRecord, 'id' | 'createdAt'>[]) => {
    const newRecords: VisaRecord[] = recordsToImport.map((r, idx) => ({
      ...r,
      id: `visa-${Date.now()}-${idx}-${Math.random().toString(36).substring(2, 7)}`,
      createdAt: new Date().toISOString(),
      isImported: true,
      source: 'import' as const,
    }));
    setVisaRecords((prev) => {
      const prevIds = new Set(prev.map((r) => r.id));
      const nonDuplicates = newRecords.filter((r) => !prevIds.has(r.id));
      const updated = [...nonDuplicates, ...prev];
      idbStorage.setItem('visa_records', 'data', updated);
      return updated;
    });
    apiService.bulkSaveVisaRecords(newRecords);
  };

  const handleDeleteVisaRecordsByDateRange = (startDate: string, endDate: string) => {
    let deletedIds: string[] = [];
    setVisaRecords((prev) => {
      const kept = prev.filter((r) => {
        const d = r.applicationDate || r.issueDate || '';
        const match = d >= startDate && d <= endDate;
        if (match) deletedIds.push(r.id);
        return !match;
      });
      idbStorage.setItem('visa_records', 'data', kept);
      return kept;
    });
    if (deletedIds.length > 0) {
      apiService.bulkDeleteVisaRecords(deletedIds);
    }
  };

  // Switch role for quick testing
  const handleSwitchRole = (targetRole: UserRole) => {
    if (targetRole !== 'Secondary' && (activePage === 'createUser' || activePage === 'userList')) {
      setActivePage('dashboard');
    }
    const matchIndex = users.findIndex((u) => u.role === targetRole);
    if (matchIndex !== -1) {
      setCurrentUserIndex(matchIndex);
      showToast(`បានប្តូរទៅកាន់គណនី ${targetRole}`, 'success');
    } else {
      // Temporarily switch current user role
      setUsers((prev) =>
        prev.map((u, idx) => (idx === currentUserIndex ? { ...u, role: targetRole } : u))
      );
      showToast(`បានកំណត់តួនាទីជា ${targetRole}`, 'success');
    }
  };

  // Switch or assign team for quick testing / permissions
  const handleSwitchTeam = (team: string) => {
    setUsers((prev) =>
      prev.map((u, idx) => (idx === currentUserIndex ? { ...u, assignedTeam: team || undefined } : u))
    );
    if (team) {
      showToast(`បានកំណត់ក្រុម "${team}" សម្រាប់អ្នកប្រើប្រាស់`, 'success');
    } else {
      showToast('បានកំណត់ជាគ្រប់ក្រុម', 'success');
    }
  };

  const handleLogout = () => {
    setIsLoggedIn(false);
    localStorage.setItem('app_is_logged_in_v1', 'false');
    localStorage.removeItem('app_session_user_role_v1');
    setSessionUserRole('Secondary');
    setLoginUsername('');
    setLoginPassword('');
    setLoginError('');
    setActivePage('dashboard');
    showToast('បានចាកចេញពីប្រព័ន្ធដោយជោគជ័យ', 'success');
  };

  const handleLoginSubmit = (e: React.FormEvent) => {
    e.preventDefault();
    const cleanUser = loginUsername.trim().toLowerCase();
    const cleanPass = loginPassword.trim();

    if (!cleanUser || !cleanPass) {
      setLoginError('សូមបំពេញអត្តលេខ និងលេខសម្ងាត់!');
      return;
    }

    const matchIndex = users.findIndex(
      (u) =>
        (((u.officerNumber || '').trim().toLowerCase() === cleanUser ||
          u.username.trim().toLowerCase() === cleanUser ||
          (u.email || '').trim().toLowerCase() === cleanUser) &&
          (u.password || '123456').trim() === cleanPass)
    );

    if (matchIndex !== -1) {
      setCurrentUserIndex(matchIndex);
      setIsLoggedIn(true);
      const role = users[matchIndex].role;
      setSessionUserRole(role);
      localStorage.setItem('app_is_logged_in_v1', 'true');
      localStorage.setItem('app_session_user_role_v1', role);
      setLoginError('');
      showToast(`សួស្តី! បានចូលប្រើប្រាស់ជាគណនី ${users[matchIndex].username}`, 'success');
    } else {
      setLoginError('អត្តលេខ ឬលេខសម្ងាត់មិនត្រឹមត្រូវឡើយ!');
    }
  };

  // Determine Page Title
  const getPageTitle = () => {
    switch (activePage) {
      case 'dashboard':
        return 'ទំព័រដើម';
      case 'refusalDeportationManager':
        return 'ការងារបដិសេធនិងបញ្ជូនចេញ (Refusal & Deportation)';
      case 'refusalNationalityData':
        return 'ទិន្នន័យសញ្ជាតិបដិសេធ';
      case 'refusalTotalSummaryReport':
        return 'សរុបទិន្នន័យបដិសេធ';
      case 'createUser':
        return 'បង្កើតគណនីប្រើប្រាស់';
      case 'userList':
        return 'ទិន្នន័យគណនីអ្នកប្រើប្រាស់';
      case 'officerForm':
        return 'ទម្រង់បញ្ចូលទិន្នន័យមន្ត្រី';
      case 'officerList':
        return 'ទិន្នន័យមន្ត្រីទាំងអស់';
      case 'visaForm':
        return 'ទម្រង់ស្នើសុំប្តូរប្រភេទទិដ្ឋាការ';
      case 'visaList':
        return 'ទិន្នន័យទិដ្ឋាការទាំងអស់';
      case 'visaSummary':
        return 'សរុបទិន្នន័យប្តូរទិដ្ឋាការ';
      case 'visaDateSearch':
        return 'ស្វែងរកទិន្នន័យទិដ្ឋាការតាមកាលបរិច្ឆេទ';
      case 'stockSticker':
        return 'ការងារស្តុក — ទម្រង់កត់ត្រាស្តុកសន្លឹក';
      case 'stockTotalSummaryReport':
        return 'ការងារស្តុក — របក.សរុបការងារស្តុក';
      case 'stockOfficeWorkSummaryReport':
        return 'ការងារស្តុក — របក.ស្តុកសរុបរួម (ការិយាល័យ)';
      case 'stockOfficeRollingSummaryReport':
        return 'ការងារស្តុក — របក.សរុបរួមរំកិល';
      case 'stockStubProposalSection':
        return 'ការងារស្តុក — ផ្នែកស្នើសុំប្រគល់គល់ទិដ្ឋាការ (ការិយាល័យ)';
      case 'stockStubProposalOffice':
        return 'ការងារស្តុក — ការិយាល័យស្នើសុំប្រគល់គល់ទិដ្ឋាការ';
      case 'stockStickerDailyTeam':
        return 'ការងារស្តុក — ប្រតិបត្តិការប្រចាំថ្ងៃតាមក្រុម';
      case 'stockStickerStubCollection':
        return 'ការងារស្តុក — ប្រមូលគល់សន្លឹកទិដ្ឋាការ';
      case 'stockStickerData':
        return 'ការងារស្តុក — ទិន្នន័យសន្លឹកទិដ្ឋាការ';
      case 'stockStickerStatus':
        return 'ការងារស្តុក — ស្តុកជាក់ស្តែងសម្រាប់បើកផ្តល់តាមក្រុម';
      case 'stockStickerOfficeReport':
        return 'ការងារស្តុក — របាយការណ៍ស្តុកសន្លឹកទិដ្ឋាការការិយាល័យ';
      case 'stockStickerTeamReport':
        return 'ការងារស្តុក — របាយការណ៍តាមក្រុម (សន្លឹកទិដ្ឋាការស្អិត)';
      case 'stockStickerK2ReceiveFromK1Report':
      case 'stockStickerK2DistributeToTeamsReport':
        return 'ការងារស្តុក — តារាងទិដ្ឋាការបើកពីក១';
      case 'stockStickerYearlyTeamUsageReport':
        return currentUser.role === 'Secondary' || currentUser.role === 'User (ការិយាល័យ)' || currentUser.role === 'Admin'
          ? 'ការងារស្តុក — តារាងប្រើប្រាស់តាមក្រុម (សន្លឹកទិដ្ឋាការស្អិត)'
          : 'ការងារស្តុក — តារាងប្រើប្រាស់របស់ក្រុម (សន្លឹកទិដ្ឋាការស្អិត)';
      case 'stockStickerYearlyTeamDistributionReport':
        return currentUser.role === 'Secondary' || currentUser.role === 'User (ការិយាល័យ)' || currentUser.role === 'Admin'
          ? 'ការងារស្តុក — តារាងផ្តល់ប្រចាំឆ្នាំទៅក្រុម (សន្លឹកទិដ្ឋាការស្អិត)'
          : 'ការងារស្តុក — តារាងទិដ្ឋាការទទួលបានពីក២ (សន្លឹកទិដ្ឋាការស្អិត)';
      case 'stockOfficeDailyTeamStatsReport':
        return 'ការងារស្តុក — ស្ថិតិសរុបប្រចាំថ្ងៃរបស់ក្រុម';
      case 'stockTeamDailyStatsReport':
        return 'ការងារស្តុក — តារាងស្ថិតិប្រចាំថ្ងៃរបស់ក្រុម';
      case 'stockStickerDailyTeamStatsReport':
        return currentUser.role === 'Secondary' || currentUser.role === 'User (ការិយាល័យ)' || currentUser.role === 'Admin'
          ? 'ការងារស្តុក — ស្ថិតិសរុបប្រចាំថ្ងៃរបស់ក្រុម'
          : 'ការងារស្តុក — តារាងស្ថិតិប្រចាំថ្ងៃរបស់ក្រុម';
      case 'teamVisaIssuanceStatsReport':
        return 'ការងារស្តុក — ស្ថិតិផ្តល់ទិដ្ឋាការរបស់ក្រុម';
      case 'singleDayTeamVisaIssuanceReport':
        return 'ការងារស្តុក — ស្ថិតិផ្តល់ប្រចាំថ្ងៃរបស់ក្រុម';
      case 'dailyStickerVisaOperationReport':
        return 'ការងារស្តុក — របាយការណ៍សន្លឹកទិដ្ឋាការ';
      case 'stockEVisa':
        return 'ការងារស្តុក — ទម្រង់កត់ត្រាស្តុកក្រដាសអនុម័ត';
      case 'stockEVisaTeamUse':
        return 'ការងារស្តុក — ទម្រង់ប្រើប្រាស់ក្រដាសតាមក្រុម';
      case 'stockEVisaData':
        return 'ការងារស្តុក — ទិន្នន័យក្រដាសអនុម័ត';
      case 'stockEVisaReport':
        return 'ការងារស្តុក — របាយការណ៍ក្រដាសអនុម័ត (គិតតាមកាលបរិច្ឆេទ)';
      case 'stockEVisaRobokReport':
        return 'ការងារស្តុក — របកក្រដាសអនុម័តតាមបណ្តាក្រុម';
      case 'stockEVisaSingleRobokReport':
        return 'ការងារស្តុក — របកក្រដាសបង្ហាញតែ១ក្រុម';
      case 'stockEVisaTeamRobokReport':
        return 'ការងារស្តុក — របកក្រដាសអនុម័តក្រុម';
      case 'stockHandoverWorkspace':
        return 'ការងារស្តុក — ផ្ទាំងការងារលិខិតប្រគល់ទទួល (PDF)';
      case 'monthlyVisaUsageChart':
        return 'ការងារស្តុក — ក្រាហ្វិកប្រៀបធៀបការប្រើប្រាស់ទិដ្ឋាការ';
      case 'categoryManager': {
        const labels: Record<CategoryType, string> = {
          ranks: 'ឋានន្តរស័ក្កិ',
          positions: 'តួនាទី',
          teamPositions: 'តួនាទីតាមក្រុម',
          visaTeams: 'ក្រុមផ្តល់ទិដ្ឋាការ',
          visaTeamsRobok: 'ក្រុមផ្តល់ទិដ្ឋាការ.របក',
          visaTeamsData: 'ទិន្នន័យក្រុមផ្តល់ទិដ្ឋាការ',
          collectorRoles: 'តួនាទីអ្នកមកបើក',
          organizations: 'ស្ថាប័នដើម',
          refusalReasons: 'ករណីបដិសេធ',
        };
        return `បញ្ចូលប្រភេទ — ${labels[activeCatType]}`;
      }
      default:
        return 'ប្រព័ន្ធគ្រប់គ្រងការិយាល័យទិដ្ឋាការចូល';
    }
  };

  if (!isLoggedIn) {
    return (
      <div className="min-h-screen w-full flex flex-col items-center justify-center bg-[#f4f6f9] text-[#1F2937] p-4 font-sans select-none">
        {/* AdminLTE style Thin Logo Header */}
        <div className="text-center mb-6">
          <h1 className="text-xl sm:text-2xl font-light text-gray-700 leading-tight">
            <span className="font-bold text-[#222]">ប្រព័ន្ធគ្រប់គ្រង</span>ការិយាល័យទិដ្ឋាការចូល
          </h1>
          <p className="text-[10px] text-gray-400 tracking-wider font-semibold uppercase mt-0.5">
            Entry Visa Office Management System
          </p>
        </div>

        {/* AdminLTE style Card */}
        <div className="w-full max-w-sm bg-white rounded-[4px] shadow-md border border-gray-200 border-t-[4px] border-t-[#007bff] p-6 flex flex-col space-y-5">
          <p className="text-xs text-center text-gray-600 font-medium">
            សូមបញ្ចូលគណនី និងលេខសម្ងាត់ដើម្បីចូលប្រើប្រាស់
          </p>

          {loginError && (
            <div className="flex items-center gap-2 bg-rose-50 border border-rose-200 text-rose-700 p-2.5 rounded-[4px] text-xs font-semibold animate-shake">
              <ShieldAlert className="w-4 h-4 shrink-0 text-rose-600" />
              <span>{loginError}</span>
            </div>
          )}

          <form onSubmit={handleLoginSubmit} className="space-y-4">
            {/* Input Group: Username / Officer ID */}
            <div className="space-y-1.5">
              <label className="block text-xs font-bold text-gray-600">
                អត្តលេខ (Officer ID)
              </label>
              <div className="flex">
                <span className="inline-flex items-center px-3 rounded-l-[4px] border border-r-0 border-gray-300 bg-[#e9ecef] text-[#495057] text-xs shrink-0">
                  <User className="w-3.5 h-3.5 text-gray-500" />
                </span>
                <input
                  type="text"
                  required
                  value={loginUsername}
                  onChange={(e) => setLoginUsername(e.target.value)}
                  placeholder="ឧទាហរណ៍៖ 65412"
                  className="flex-1 min-w-0 block w-full px-3 py-2 rounded-r-[4px] border border-gray-300 text-xs focus:border-[#80bdff] focus:ring-1 focus:ring-[#80bdff] focus:outline-none bg-white text-gray-900 transition-all"
                />
              </div>
            </div>

            {/* Input Group: Password */}
            <div className="space-y-1.5">
              <label className="block text-xs font-bold text-gray-600">
                លេខសម្ងាត់ (Password)
              </label>
              <div className="flex relative">
                <span className="inline-flex items-center px-3 rounded-l-[4px] border border-r-0 border-gray-300 bg-[#e9ecef] text-[#495057] text-xs shrink-0">
                  <Lock className="w-3.5 h-3.5 text-gray-500" />
                </span>
                <input
                  type={showLoginPassword ? 'text' : 'password'}
                  required
                  value={loginPassword}
                  onChange={(e) => setLoginPassword(e.target.value)}
                  placeholder="••••••••"
                  className="flex-1 min-w-0 block w-full px-3 pr-10 py-2 rounded-r-[4px] border border-gray-300 text-xs focus:border-[#80bdff] focus:ring-1 focus:ring-[#80bdff] focus:outline-none bg-white text-gray-900 font-mono transition-all text-gray-800"
                />
                <button
                  type="button"
                  onClick={() => setShowLoginPassword(!showLoginPassword)}
                  className="absolute inset-y-0 right-0 flex items-center pr-3 text-gray-400 hover:text-gray-600 cursor-pointer"
                >
                  {showLoginPassword ? <EyeOff className="w-4 h-4 text-gray-500" /> : <Eye className="w-4 h-4 text-gray-500" />}
                </button>
              </div>
            </div>

            {/* Submit Button */}
            <button
              type="submit"
              className="w-full bg-[#007bff] hover:bg-[#0069d9] hover:shadow-sm active:bg-[#005cbf] transition-all py-2 rounded-[4px] text-xs font-bold text-white shadow-xs cursor-pointer flex items-center justify-center gap-1.5"
            >
              <LogIn className="w-3.5 h-3.5" />
              <span>ចូលប្រើប្រាស់ប្រព័ន្ធ</span>
            </button>
          </form>

          {/* Test credentials panel */}
          <div className="border-t border-gray-200 pt-3 mt-1">
            <p className="text-[10px] text-gray-500 font-bold mb-1 text-center uppercase tracking-wider">
              គណនីសម្រាប់សាកល្បង៖
            </p>
            <div className="bg-[#f8f9fa] p-2 rounded-[4px] border border-gray-200 text-[10px] text-gray-600 font-mono space-y-1">
              <div className="flex justify-between border-b border-gray-200/50 pb-1">
                <span className="font-sans font-bold">Secondary (Office):</span>
                <span className="font-bold text-[#007bff]">admin / 123456</span>
              </div>
              <div className="flex justify-between border-b border-gray-200/50 pb-1">
                <span className="font-sans font-bold">User (Office):</span>
                <span className="font-bold text-[#007bff]">65412 / 65412</span>
              </div>
              <div className="flex justify-between">
                <span className="font-sans font-bold">User (Team/ច្រក)៖</span>
                <span className="font-bold text-[#007bff]">user / 123456</span>
              </div>
            </div>
          </div>
        </div>
      </div>
    );
  }

  return (
    <div className="h-screen w-full flex bg-[#f4f6f9] text-[#1F2937] overflow-hidden print:h-auto print:overflow-visible print:bg-white app-layout-container">
      {/* Sidebar Navigation */}
      <Sidebar
        activePage={activePage}
        activeCatType={activeCatType}
        currentRole={currentUser.role}
        userName={currentUser.username}
        assignedTeam={currentUser.assignedTeam}
        categories={categories}
        userPermissions={currentUser.permissions}
        hideOfficeOfficersFromTeams={hideOfficeOfficersFromTeams}
        hideOfficerMenuFromTeams={hideOfficerMenuFromTeams}
        hideUserMenuFromTeams={hideUserMenuFromTeams}
        hideDailyOpsFromTeams={hideDailyOpsFromTeams}
        hideRefusalFromTeams={hideRefusalFromTeams}
        hideVisaWorkFromTeams={hideVisaWorkFromTeams}
        hideOfficeStockFromTeams={hideOfficeStockFromTeams}
        hideSearchCodeFromTeams={hideSearchCodeFromTeams}
        hideCategoriesFromTeams={hideCategoriesFromTeams}
        onNavigate={handleNavigate}
        onLogout={handleLogout}
      />

      {/* Main Content Area */}
      <div className="flex-1 flex flex-col min-w-0 h-full overflow-hidden print:h-auto print:overflow-visible bg-[#f4f6f9]">
        {/* Sticky Header */}
        <div className="shrink-0 z-20">
          <Header
            currentUser={currentUser}
            pageTitle={getPageTitle()}
            categories={categories}
            toastMessage={toastMessage}
            isCloudConnected={isCloudConnected}
            isCloudSyncing={isCloudSyncing}
            onSwitchUserRole={handleSwitchRole}
            onSwitchTeam={handleSwitchTeam}
            onLogout={handleLogout}
            showRoleSwitcher={sessionUserRole === 'Secondary' || sessionUserRole === 'Admin'}
          />
        </div>

        {/* Dynamic Page Rendering - Independent Scrollable Area */}
        <div
          ref={mainScrollRef}
          className="flex-1 overflow-y-auto overflow-x-hidden bg-[#f4f6f9] focus:outline-none custom-main-scroll print:h-auto print:overflow-visible app-main-content-scroll"
        >
          <main
            className={`w-full ${
              activePage === 'stockStickerDailyTeam' || activePage === 'stockStickerStubCollection'
                ? 'p-2 sm:p-3'
                : 'p-3 sm:p-4 md:p-5 lg:p-6'
            }`}
          >
          {activePage === 'dashboard' && (
            <Dashboard
              officers={officers}
              users={users}
              categories={categories}
              stockRecords={stockRecords}
              currentRole={currentUser.role}
              assignedTeam={currentUser.assignedTeam}
              hideOfficeOfficersFromTeams={hideOfficeOfficersFromTeams}
              onNavigate={handleNavigate}
            />
          )}

          {activePage === 'refusalDeportationManager' && (
            <RefusalDeportationManager
              currentRole={currentUser.role}
              assignedTeam={currentUser.assignedTeam}
              userName={currentUser.username}
              categories={categories}
              records={refusalRecords}
              onUpdateRecords={setRefusalRecords}
              onClose={() => handleNavigate('dashboard')}
            />
          )}

          {activePage === 'refusalNationalityData' && (
            <RefusalNationalityData records={refusalRecords} categories={categories} />
          )}

          {activePage === 'refusalTotalSummaryReport' && (
            <RefusalTotalSummaryReport records={refusalRecords} categories={categories} />
          )}

          {activePage === 'monthlyVisaUsageChart' && (
            <div className="w-full space-y-4">
              <MonthlyVisaUsageComparisonChart stockRecords={stockRecords} />
            </div>
          )}

          {activePage === 'createUser' && (currentUser.role === 'Secondary' || currentUser.role === 'User (ការិយាល័យ)' || !hideUserMenuFromTeams) && (
            <UserManagement
              users={users}
              viewMode="create"
              categories={categories}
              hideOfficeOfficersFromTeams={hideOfficeOfficersFromTeams}
              hideOfficerMenuFromTeams={hideOfficerMenuFromTeams}
              hideUserMenuFromTeams={hideUserMenuFromTeams}
              hideDailyOpsFromTeams={hideDailyOpsFromTeams}
              hideRefusalFromTeams={hideRefusalFromTeams}
              hideVisaWorkFromTeams={hideVisaWorkFromTeams}
              hideOfficeStockFromTeams={hideOfficeStockFromTeams}
              hideSearchCodeFromTeams={hideSearchCodeFromTeams}
              hideCategoriesFromTeams={hideCategoriesFromTeams}
              onToggleHideOfficeOfficersFromTeams={handleToggleHideOfficeOfficersFromTeams}
              onToggleHideOfficerMenuFromTeams={handleToggleHideOfficerMenuFromTeams}
              onToggleHideUserMenuFromTeams={handleToggleHideUserMenuFromTeams}
              onToggleHideDailyOpsFromTeams={handleToggleHideDailyOpsFromTeams}
              onToggleHideRefusalFromTeams={handleToggleHideRefusalFromTeams}
              onToggleHideVisaWorkFromTeams={handleToggleHideVisaWorkFromTeams}
              onToggleHideOfficeStockFromTeams={handleToggleHideOfficeStockFromTeams}
              onToggleHideSearchCodeFromTeams={handleToggleHideSearchCodeFromTeams}
              onToggleHideCategoriesFromTeams={handleToggleHideCategoriesFromTeams}
              onAddUser={handleAddUser}
              onUpdateUser={handleUpdateUser}
              onDeleteUser={handleDeleteUser}
              onShowToast={showToast}
            />
          )}

          {activePage === 'userList' && (currentUser.role === 'Secondary' || currentUser.role === 'User (ការិយាល័យ)' || !hideUserMenuFromTeams) && (
            <UserManagement
              users={
                isTeamUser(currentUser.role, currentUser.assignedTeam)
                  ? users.filter((u) => {
                      if (u.isVisible === false || u.hidden) return false;
                      if (currentUser.assignedTeam && u.assignedTeam) {
                        return u.assignedTeam.trim().toLowerCase() === currentUser.assignedTeam.trim().toLowerCase();
                      }
                      return u.id === currentUser.id || u.username === currentUser.username;
                    })
                  : users
              }
              viewMode="list"
              categories={categories}
              hideOfficeOfficersFromTeams={hideOfficeOfficersFromTeams}
              hideOfficerMenuFromTeams={hideOfficerMenuFromTeams}
              hideUserMenuFromTeams={hideUserMenuFromTeams}
              hideDailyOpsFromTeams={hideDailyOpsFromTeams}
              hideRefusalFromTeams={hideRefusalFromTeams}
              hideVisaWorkFromTeams={hideVisaWorkFromTeams}
              hideOfficeStockFromTeams={hideOfficeStockFromTeams}
              hideSearchCodeFromTeams={hideSearchCodeFromTeams}
              hideCategoriesFromTeams={hideCategoriesFromTeams}
              onToggleHideOfficeOfficersFromTeams={handleToggleHideOfficeOfficersFromTeams}
              onToggleHideOfficerMenuFromTeams={handleToggleHideOfficerMenuFromTeams}
              onToggleHideUserMenuFromTeams={handleToggleHideUserMenuFromTeams}
              onToggleHideDailyOpsFromTeams={handleToggleHideDailyOpsFromTeams}
              onToggleHideRefusalFromTeams={handleToggleHideRefusalFromTeams}
              onToggleHideVisaWorkFromTeams={handleToggleHideVisaWorkFromTeams}
              onToggleHideOfficeStockFromTeams={handleToggleHideOfficeStockFromTeams}
              onToggleHideSearchCodeFromTeams={handleToggleHideSearchCodeFromTeams}
              onToggleHideCategoriesFromTeams={handleToggleHideCategoriesFromTeams}
              onAddUser={handleAddUser}
              onUpdateUser={handleUpdateUser}
              onDeleteUser={handleDeleteUser}
              onShowToast={showToast}
            />
          )}

          {activePage === 'officerForm' && (
            <OfficerForm
              categories={categories}
              currentRole={currentUser.role}
              currentUserId={currentUser.id}
              assignedTeam={currentUser.assignedTeam}
              hideOfficeOfficersFromTeams={hideOfficeOfficersFromTeams}
              onAddOfficer={handleAddOfficer}
              onNavigateCategoryManager={(type) => handleNavigate('categoryManager', type)}
              onShowToast={showToast}
            />
          )}

          {activePage === 'officerList' && (
            <OfficerList
              officers={officers}
              categories={categories}
              currentRole={currentUser.role}
              assignedTeam={currentUser.assignedTeam}
              hideOfficeOfficersFromTeams={hideOfficeOfficersFromTeams}
              onUpdateOfficer={handleUpdateOfficer}
              onDeleteOfficer={handleDeleteOfficer}
              onNavigateCategoryManager={(type) => handleNavigate('categoryManager', type)}
              onShowToast={showToast}
            />
          )}

          {activePage === 'visaForm' && (
            <VisaForm
              officers={officers}
              visaTeams={categories.visaTeams}
              organizations={categories.organizations}
              onAddVisaRecord={handleAddVisaRecord}
              onShowToast={showToast}
              onNavigateToList={() => handleNavigate('visaList')}
            />
          )}

          {activePage === 'visaList' && (
            <VisaList
              records={visaRecords}
              officers={officers}
              visaTeams={categories.visaTeams}
              organizations={categories.organizations}
              onUpdateRecord={handleUpdateVisaRecord}
              onDeleteRecord={handleDeleteVisaRecord}
              onBatchImportRecords={handleBatchImportVisaRecords}
              onDeleteRecordsByDateRange={handleDeleteVisaRecordsByDateRange}
              onShowToast={showToast}
              onNavigateToForm={() => handleNavigate('visaForm')}
            />
          )}

          {activePage === 'visaSummary' && (
            <VisaSummary
              records={visaRecords}
              visaTeams={categories.visaTeams}
            />
          )}

          {activePage === 'visaDateSearch' && (
            <VisaDateSearch
              records={visaRecords}
            />
          )}

          {activePage === 'stockStickerStubCollection' && (
            <VisaStubCollectionOperations
              categories={categories}
              stockRecords={stockRecords}
              currentRole={currentUser.role}
              userName={currentUser.username}
              assignedTeam={currentUser.assignedTeam}
              onAddStockRecord={handleAddStockRecord}
              onBatchImportStockRecords={handleBatchImportStockRecords}
              onUpdateStockRecord={handleUpdateStockRecord}
              onDeleteStockRecord={handleDeleteStockRecord}
              onDeleteBatchStockRecords={handleDeleteBatchStockRecords}
              onShowToast={showToast}
            />
          )}

          {activePage === 'stockSearchCodeByTeam' && (
            <TeamCodeSearch
              categories={categories}
              stockRecords={stockRecords}
              visaRecords={visaRecords}
              stickerActualStock={stickerActualStock}
              currentRole={currentUser.role}
              userName={currentUser.username}
              onShowToast={showToast}
            />
          )}

          {activePage === 'stockStickerTeamReport' && (
            <StickerTeamStockReport
              stockRecords={stockRecords}
              categories={categories}
              officers={officers}
              currentRole={currentUser.role}
              userName={currentUser.username}
              onClose={() => handleNavigate('dashboard')}
            />
          )}

          {activePage === 'stockStickerK2ReceiveFromK1Report' && (
            <YearlyK2ReceivedK1Report
              stockRecords={stockRecords}
              categories={categories}
              officers={officers}
              currentRole={currentUser.role}
              userName={currentUser.username}
              assignedTeam={currentUser.assignedTeam}
              onClose={() => handleNavigate('dashboard')}
            />
          )}

          {activePage === 'stockStickerK2DistributeToTeamsReport' && (
            <YearlyK2IssuedToTeamsReport
              stockRecords={stockRecords}
              categories={categories}
              officers={officers}
              currentRole={currentUser.role}
              userName={currentUser.username}
              assignedTeam={currentUser.assignedTeam}
              onClose={() => handleNavigate('dashboard')}
            />
          )}

          {activePage === 'stockStickerYearlyTeamDistributionReport' && (
            <YearlyTeamDistributionReport
              stockRecords={stockRecords}
              categories={categories}
              officers={officers}
              currentRole={currentUser.role}
              userName={currentUser.username}
              assignedTeam={currentUser.assignedTeam}
              onClose={() => handleNavigate('dashboard')}
            />
          )}

          {activePage === 'stockStickerYearlyTeamUsageReport' && (
            <YearlyTeamUsageReport
              stockRecords={stockRecords}
              categories={categories}
              officers={officers}
              currentRole={currentUser.role}
              userName={currentUser.username}
              assignedTeam={currentUser.assignedTeam}
              onClose={() => handleNavigate('dashboard')}
            />
          )}

          {(activePage === 'stockOfficeDailyTeamStatsReport' || (activePage === 'stockStickerDailyTeamStatsReport' && (currentUser.role === 'Secondary' || currentUser.role === 'User (ការិយាល័យ)' || currentUser.role === 'Admin'))) && (
            <DailyTeamStatisticsReport
              mode="office"
              stockRecords={stockRecords}
              categories={categories}
              officers={officers}
              currentRole={currentUser.role}
              userName={currentUser.username}
              assignedTeam={currentUser.assignedTeam}
              onClose={() => handleNavigate('dashboard')}
              onShowToast={showToast}
              onAddStockRecord={handleAddStockRecord}
              onBatchImportStockRecords={handleBatchImportStockRecords}
              onUpdateStockRecord={handleUpdateStockRecord}
              onDeleteStockRecord={handleDeleteStockRecord}
              onDeleteBatchStockRecords={handleDeleteBatchStockRecords}
            />
          )}

          {activePage === 'stockStickerDailyTeam' && (
            <DailyTeamVisaOperations
              categories={categories}
              stockRecords={stockRecords}
              currentRole={currentUser.role}
              userName={currentUser.username}
              assignedTeam={currentUser.assignedTeam}
              onAddStockRecord={handleAddStockRecord}
              onBatchImportStockRecords={handleBatchImportStockRecords}
              onUpdateStockRecord={handleUpdateStockRecord}
              onDeleteStockRecord={handleDeleteStockRecord}
              onDeleteBatchStockRecords={handleDeleteBatchStockRecords}
              onShowToast={showToast}
            />
          )}

          {(activePage === 'stockTeamDailyStatsReport' || (activePage === 'stockStickerDailyTeamStatsReport' && currentUser.role !== 'Secondary' && currentUser.role !== 'User (ការិយាល័យ)' && currentUser.role !== 'Admin')) && (
            <DailyTeamStatisticsReport
              mode="team"
              stockRecords={stockRecords}
              categories={categories}
              officers={officers}
              currentRole={currentUser.role}
              userName={currentUser.username}
              assignedTeam={currentUser.assignedTeam}
              onClose={() => handleNavigate('dashboard')}
              onShowToast={showToast}
              onAddStockRecord={handleAddStockRecord}
              onBatchImportStockRecords={handleBatchImportStockRecords}
              onUpdateStockRecord={handleUpdateStockRecord}
              onDeleteStockRecord={handleDeleteStockRecord}
              onDeleteBatchStockRecords={handleDeleteBatchStockRecords}
            />
          )}

          {activePage === 'teamVisaIssuanceStatsReport' && (
            <TeamVisaIssuanceStatsReport
              stockRecords={stockRecords}
              categories={categories}
              officers={officers}
              currentRole={currentUser.role}
              userName={currentUser.username}
              assignedTeam={currentUser.assignedTeam}
              onClose={() => handleNavigate('dashboard')}
              singleDateMode={false}
              storageKeyPrefix="range_"
            />
          )}

          {activePage === 'singleDayTeamVisaIssuanceReport' && (
            <TeamVisaIssuanceStatsReport
              stockRecords={stockRecords}
              categories={categories}
              officers={officers}
              currentRole={currentUser.role}
              userName={currentUser.username}
              assignedTeam={currentUser.assignedTeam}
              onClose={() => handleNavigate('dashboard')}
              singleDateMode={true}
              storageKeyPrefix="single_day_"
            />
          )}

          {activePage === 'dailyStickerVisaOperationReport' && (
            <DailyStickerVisaOperationReport
              stockRecords={stockRecords}
              categories={categories}
              officers={officers}
              currentRole={currentUser.role}
              userName={currentUser.username}
              assignedTeam={currentUser.assignedTeam}
              onClose={() => handleNavigate('dashboard')}
              onShowToast={showToast}
            />
          )}

          {activePage === 'branchChiefTeamVisaStatsReport' && (
            <BranchChiefTeamVisaStatsReport
              stockRecords={stockRecords}
              categories={categories}
              officers={officers}
              currentRole={currentUser.role}
              userName={currentUser.username}
              assignedTeam={currentUser.assignedTeam}
              onClose={() => handleNavigate('dashboard')}
            />
          )}

          {activePage === 'stockEVisaTeamRobokReport' && (
            <EVisaTeamRobokReport
              stockRecords={stockRecords}
              categories={categories}
              officers={officers}
              currentRole={currentUser.role}
              userName={currentUser.username}
              assignedTeam={currentUser.assignedTeam}
              onClose={() => handleNavigate('dashboard')}
              onShowToast={showToast}
            />
          )}

          {activePage === 'stockTotalSummaryReport' && (
            <RobokTotalStockWorkReport
              stockRecords={stockRecords}
              categories={categories}
              officers={officers}
              currentRole={currentUser.role}
              userName={currentUser.username}
              assignedTeam={currentUser.assignedTeam}
              initialViewMode="team"
              onClose={() => handleNavigate('dashboard')}
            />
          )}

          {activePage === 'stockOfficeWorkSummaryReport' && (
            <RobokTotalStockWorkReport
              stockRecords={stockRecords}
              categories={categories}
              officers={officers}
              currentRole={currentUser.role}
              userName={currentUser.username}
              assignedTeam={currentUser.assignedTeam}
              initialViewMode="office"
              onClose={() => handleNavigate('dashboard')}
            />
          )}

          {activePage === 'stockOfficeRollingSummaryReport' && (
            <RobokTotalStockWorkReport
              stockRecords={stockRecords}
              categories={categories}
              officers={officers}
              currentRole={currentUser.role}
              userName={currentUser.username}
              assignedTeam={currentUser.assignedTeam}
              initialViewMode="office"
              isRolling={true}
              onClose={() => handleNavigate('dashboard')}
            />
          )}

          {activePage === 'stockStubProposalSection' && (currentUser.role === 'Secondary' || currentUser.role === 'User (ការិយាល័យ)') && (
            <RobokTotalStockWorkReport
              stockRecords={stockRecords}
              categories={categories}
              officers={officers}
              currentRole={currentUser.role}
              userName={currentUser.username}
              assignedTeam={currentUser.assignedTeam}
              initialViewMode="office"
              proposalType="section"
              defaultDocThroughTitle="លោកនាយរងការិយាល័យទទួលបន្ទុក"
              onClose={() => handleNavigate('dashboard')}
            />
          )}

          {activePage === 'stockStubProposalOffice' && (currentUser.role === 'Secondary' || currentUser.role === 'User (ការិយាល័យ)') && (
            <RobokTotalStockWorkReport
              stockRecords={stockRecords}
              categories={categories}
              officers={officers}
              currentRole={currentUser.role}
              userName={currentUser.username}
              assignedTeam={currentUser.assignedTeam}
              initialViewMode="office"
              proposalType="office"
              defaultDocThroughTitle="លោកនាយការិយាល័យទិដ្ឋាការចូល"
              onClose={() => handleNavigate('dashboard')}
            />
          )}

          {activePage === 'stockDirectorSummaryReport' && (
            <RobokDirectorStockWorkReport
              stockRecords={stockRecords}
              categories={categories}
              officers={officers}
              currentRole={currentUser.role}
              userName={currentUser.username}
              assignedTeam={currentUser.assignedTeam}
              onClose={() => handleNavigate('dashboard')}
            />
          )}

          {activePage === 'stockStickerStatus' && (
            <StickerVisaStatus
              stockRecords={stickerActualStock}
              categories={categories}
              officers={officers}
              currentRole={currentUser.role}
              userName={currentUser.username}
              onAddStockRecord={handleIssueToTeamFromActualStock}
              onAddK1StockRecord={handleAddK1StockToActualStock}
              onBatchImportStockRecords={handleBatchImportStickerActualStock}
              onReturnIssuedStock={handleReturnIssuedStock}
              onReturnIssuedStockBatch={handleReturnIssuedStockBatch}
              onReturnIssuedStockByVisaType={handleReturnIssuedStockByVisaType}
              onClearAllActualStock={handleClearAllActualStock}
              onDeleteActualStockRecord={handleDeleteActualStockRecord}
              onDeleteBatchActualStockRecords={handleDeleteBatchActualStockRecords}
              onDeleteActualStockByRange={handleDeleteActualStockByRange}
              onShowToast={showToast}
              onNavigateToForm={() => handleNavigate('stockSticker')}
            />
          )}

          {(activePage === 'stockSticker' || activePage === 'stockStickerData' || activePage === 'stockStickerOfficeReport' || activePage === 'stockEVisa' || activePage === 'stockEVisaTeamUse' || activePage === 'stockEVisaData' || activePage === 'stockEVisaReport' || activePage === 'stockEVisaRobokReport' || activePage === 'stockEVisaSingleRobokReport' || activePage === 'stockManager' || activePage === 'stockHandoverWorkspace') && (
            <StockManager
              key={activePage}
              initialStockType={activePage === 'stockSticker' || activePage === 'stockStickerData' || activePage === 'stockStickerOfficeReport' ? 'sticker' : 'evisa'}
              initialOperationType={
                activePage === 'stockEVisaTeamUse'
                  ? 'useTeam'
                  : activePage === 'stockEVisa'
                  ? (currentUser.role === 'Secondary' || currentUser.role === 'User (ការិយាល័យ)' ? 'openK1' : 'useTeam')
                  : undefined
              }
              hideOperationSelector={false}
              initialViewMode={
                activePage === 'stockHandoverWorkspace'
                  ? 'handoverWorkspace'
                  : activePage === 'stockStickerOfficeReport'
                  ? 'stickerOfficeReport'
                  : activePage === 'stockEVisaSingleRobokReport'
                  ? 'singleRobokReport'
                  : activePage === 'stockEVisaRobokReport'
                  ? 'robokReport'
                  : activePage === 'stockEVisaReport'
                  ? 'report'
                  : activePage === 'stockEVisaData' || activePage === 'stockStickerData'
                  ? 'data'
                  : (activePage === 'stockEVisa' || activePage === 'stockEVisaTeamUse' || activePage === 'stockSticker')
                  ? 'form'
                  : 'all'
              }
              categories={categories}
              officers={officers}
              stockRecords={stockRecords}
              actualStockRecords={stickerActualStock}
              users={users}
              currentRole={currentUser.role}
              userName={currentUser.username}
              assignedTeam={currentUser.assignedTeam}
              onAddStockRecord={handleAddStockRecord}
              onBatchImportStockRecords={handleBatchImportStockRecords}
              onUpdateStockRecord={handleUpdateStockRecord}
              onDeleteStockRecord={handleDeleteStockRecord}
              onDeleteBatchStockRecords={handleDeleteBatchStockRecords}
              onNavigate={handleNavigate}
              onShowToast={showToast}
            />
          )}

          {activePage === 'categoryManager' && (
            <CategoryManager
              categories={categories}
              activeCatType={activeCatType}
              currentRole={currentUser.role}
              onAddCategory={handleAddCategory}
              onUpdateCategory={handleUpdateCategory}
              onDeleteCategory={handleDeleteCategory}
              onNavigateToOfficerForm={() => handleNavigate('officerForm')}
              onSelectCatType={(type) => setActiveCatType(type)}
              onShowToast={showToast}
              onResetVisaTeams={handleResetVisaTeams}
            />
          )}
          </main>
        </div>
      </div>
    </div>
  );
}

