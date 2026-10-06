import React, { useState, useEffect, useMemo } from 'react';
import { CategoriesState, CategoryType, UserRole, UserPermissions } from '../types';
import { useWorkspaceSettings } from '../context/WorkspaceSettingsContext';
import {
  LayoutDashboard,
  UserPlus,
  Users,
  FileText,
  Table,
  FolderPlus,
  ChevronDown,
  Award,
  Briefcase,
  Layers,
  Building2,
  Shield,
  ShieldCheck,
  ShieldAlert,
  History,
  Lock,
  Menu,
  X,
  RefreshCw,
  FileSpreadsheet,
  BarChart3,
  CalendarSearch,
  FileCheck,
  UserCheck,
  Boxes,
  Package,
  Calendar,
  BookOpen,
  Search,
  BookmarkCheck,
  PanelLeftClose,
  PanelLeftOpen,
} from 'lucide-react';

interface SidebarProps {
  activePage: string;
  activeCatType?: CategoryType;
  currentRole: UserRole;
  userName: string;
  assignedTeam?: string;
  categories?: CategoriesState;
  userPermissions?: UserPermissions;
  hideOfficeOfficersFromTeams?: boolean;
  hideOfficerMenuFromTeams?: boolean;
  hideUserMenuFromTeams?: boolean;
  hideDailyOpsFromTeams?: boolean;
  hideRefusalFromTeams?: boolean;
  hideVisaWorkFromTeams?: boolean;
  hideOfficeStockFromTeams?: boolean;
  hideSearchCodeFromTeams?: boolean;
  hideCategoriesFromTeams?: boolean;
  onNavigate: (page: string, catType?: CategoryType) => void;
  onLogout: () => void;
}

export const Sidebar: React.FC<SidebarProps> = ({
  activePage,
  activeCatType,
  currentRole,
  userName,
  assignedTeam,
  categories,
  userPermissions,
  hideOfficeOfficersFromTeams = true,
  hideOfficerMenuFromTeams = false,
  hideUserMenuFromTeams = true,
  hideDailyOpsFromTeams = false,
  hideRefusalFromTeams = false,
  hideVisaWorkFromTeams = true,
  hideOfficeStockFromTeams = true,
  hideSearchCodeFromTeams = false,
  hideCategoriesFromTeams = true,
  onNavigate,
  onLogout,
}) => {
  const { isSidebarCollapsed, toggleSidebarCollapse } = useWorkspaceSettings();
  const isSecondary = currentRole === 'Secondary' || currentRole === 'User (ការិយាល័យ)';
  const isSuperAdmin = currentRole === 'Secondary' || currentRole === 'Admin';
  const isTeam = !isSecondary;
  const isOfficeRestricted = hideOfficeOfficersFromTeams && isTeam;

  // Individual or global menu visibility checks
  const isDailyOpsVisible = userPermissions?.dailyOps !== undefined 
    ? userPermissions.dailyOps 
    : (!hideDailyOpsFromTeams || isSuperAdmin);

  const isRefusalVisible = userPermissions?.refusal !== undefined 
    ? userPermissions.refusal 
    : (!hideRefusalFromTeams || isSuperAdmin);

  const isUserMenuVisible = userPermissions?.userMenu !== undefined 
    ? userPermissions.userMenu 
    : (!hideUserMenuFromTeams || isSuperAdmin);

  const isUserListVisible = userPermissions?.userList !== undefined 
    ? userPermissions.userList 
    : (!hideUserMenuFromTeams || isSuperAdmin);

  const isOfficerMenuVisible = userPermissions?.officerMenu !== undefined 
    ? userPermissions.officerMenu 
    : (!hideOfficerMenuFromTeams || isSuperAdmin);

  const isVisaWorkVisible = userPermissions?.visaWork !== undefined 
    ? userPermissions.visaWork 
    : (!hideVisaWorkFromTeams || isSuperAdmin);

  const isOfficeStockVisible = userPermissions?.officeStock !== undefined 
    ? userPermissions.officeStock 
    : (!hideOfficeStockFromTeams || isSuperAdmin);

  const isSearchCodeVisible = userPermissions?.searchCode !== undefined 
    ? userPermissions.searchCode 
    : (!hideSearchCodeFromTeams || isSuperAdmin);

  const isCategoriesVisible = userPermissions?.categories !== undefined 
    ? userPermissions.categories 
    : (!hideCategoriesFromTeams || isSuperAdmin);

  // Listen to team type options updates in real-time
  const [teamOptionsVersion, setTeamOptionsVersion] = useState(0);

  useEffect(() => {
    const handleUpdate = () => {
      setTeamOptionsVersion((v) => v + 1);
    };
    window.addEventListener('team_type_options_updated', handleUpdate);
    return () => window.removeEventListener('team_type_options_updated', handleUpdate);
  }, []);

  // Determine if paper approval (ក្រដាសអនុម័ត / cEA) is enabled for the team
  const hasPaperApproval = useMemo(() => {
    if (isSecondary) return true; // Secondary / Office has full access to all stock tools
    if (!assignedTeam) return true; // If no team assigned, default to true or show all

    try {
      const saved = localStorage.getItem('team_type_options_v1');
      if (!saved) return false; // Default to false if not yet set in team options
      const parsed = JSON.parse(saved);
      const cleanTeam = assignedTeam.trim().toLowerCase();

      // Check direct name match
      for (const [key, val] of Object.entries<any>(parsed)) {
        if (String(key).trim().toLowerCase() === cleanTeam) {
          return Boolean(val?.evisa);
        }
      }

      // Search in visaTeamsRobok
      if (categories?.visaTeamsRobok) {
        const idx = categories.visaTeamsRobok.findIndex(
          (t) => t.name.trim().toLowerCase() === cleanTeam
        );
        if (idx !== -1) {
          if (parsed[idx] && parsed[idx].evisa !== undefined) return Boolean(parsed[idx].evisa);
          const vtrId = categories.visaTeamsRobok[idx]?.id;
          if (vtrId && parsed[vtrId] && parsed[vtrId].evisa !== undefined) return Boolean(parsed[vtrId].evisa);
        }
      }

      // Search in visaTeams (Full name)
      if (categories?.visaTeams) {
        const idx = categories.visaTeams.findIndex(
          (t) => t.name.trim().toLowerCase() === cleanTeam
        );
        if (idx !== -1) {
          if (parsed[idx] && parsed[idx].evisa !== undefined) return Boolean(parsed[idx].evisa);
          const vtId = categories.visaTeams[idx]?.id;
          if (vtId && parsed[vtId] && parsed[vtId].evisa !== undefined) return Boolean(parsed[vtId].evisa);
        }
      }
    } catch (e) {
      console.error('Error reading team_type_options_v1', e);
    }
    return false;
  }, [isSecondary, assignedTeam, categories?.visaTeamsRobok, categories?.visaTeams, teamOptionsVersion]);

  const isStickerPage =
    activePage === 'stockStickerDailyTeam' ||
    activePage === 'stockStickerStubCollection' ||
    activePage === 'stockSticker' ||
    activePage === 'stockStickerData' ||
    activePage === 'stockStickerOfficeReport' ||
    activePage === 'stockStickerTeamReport' ||
    activePage === 'stockStickerStatus';

  const isEVisaPage =
    activePage === 'stockEVisa' ||
    activePage === 'stockEVisaTeamUse' ||
    activePage === 'stockEVisaData' ||
    activePage === 'stockEVisaReport' ||
    activePage === 'stockEVisaRobokReport' ||
    activePage === 'stockEVisaSingleRobokReport' ||
    activePage === 'stockEVisaTeamRobokReport';

  const isK2OperationPage =
    activePage === 'stockStickerK2ReceiveFromK1Report' ||
    activePage === 'stockStickerK2DistributeToTeamsReport';

  const isDailyTeamReportPage =
    activePage === 'stockTeamDailyStatsReport' ||
    activePage === 'stockStickerDailyTeam' ||
    activePage === 'singleDayTeamVisaIssuanceReport' ||
    activePage === 'dailyStickerVisaOperationReport';

  const isTeamOperationPage =
    activePage === 'stockStickerYearlyTeamDistributionReport' ||
    activePage === 'stockStickerYearlyTeamUsageReport' ||
    activePage === 'branchChiefTeamVisaStatsReport';

  const isReportPage =
    activePage === 'monthlyVisaUsageChart' ||
    activePage === 'stockTotalSummaryReport' ||
    activePage === 'stockOfficeWorkSummaryReport' ||
    activePage === 'stockOfficeRollingSummaryReport' ||
    activePage === 'stockOfficeDailyTeamStatsReport' ||
    activePage === 'stockTeamDailyStatsReport' ||
    isK2OperationPage ||
    isTeamOperationPage;

  const isStubProposalPage =
    activePage === 'stockStubProposalSection' ||
    activePage === 'stockStubProposalOffice';

  const isStockPage =
    isStickerPage ||
    isEVisaPage ||
    isReportPage ||
    isStubProposalPage ||
    activePage === 'stockSearchCodeByTeam' ||
    activePage === 'stockDirectorSummaryReport' ||
    activePage === 'stockManager';

  const isAcctPage =
    activePage === 'createUser' ||
    activePage === 'userList';

  const isRefusalPage =
    activePage === 'refusalDeportationManager' ||
    activePage === 'refusalNationalityData' ||
    activePage === 'refusalTotalSummaryReport';

  const [refusalMenuOpen, setRefusalMenuOpen] = useState(isRefusalPage);
  const [acctMenuOpen, setAcctMenuOpen] = useState(isAcctPage);
  const [officerMenuOpen, setOfficerMenuOpen] = useState(
    activePage === 'officerForm' || activePage === 'officerList'
  );
  const [catMenuOpen, setCatMenuOpen] = useState(activePage === 'categoryManager');
  const [stockMenuOpen, setStockMenuOpen] = useState(isStockPage);
  const [stickerSubMenuOpen, setStickerSubMenuOpen] = useState(isStickerPage);
  const [eVisaSubMenuOpen, setEVisaSubMenuOpen] = useState(isEVisaPage);
  const [reportSubMenuOpen, setReportSubMenuOpen] = useState(isReportPage);
  const [stubProposalSubMenuOpen, setStubProposalSubMenuOpen] = useState(isStubProposalPage);
  const [k2OperationSubMenuOpen, setK2OperationSubMenuOpen] = useState(isK2OperationPage);
  const [teamOperationSubMenuOpen, setTeamOperationSubMenuOpen] = useState(isTeamOperationPage);
  const [dailyReportSubMenuOpen, setDailyReportSubMenuOpen] = useState(isDailyTeamReportPage);
  const [visaMenuOpen, setVisaMenuOpen] = useState(
    activePage === 'visaForm' ||
      activePage === 'visaList' ||
      activePage === 'visaSummary' ||
      activePage === 'visaDateSearch'
  );
  const [mobileOpen, setMobileOpen] = useState(false);

  useEffect(() => {
    if (isStickerPage) {
      setStockMenuOpen(true);
      setStickerSubMenuOpen(true);
    } else if (isEVisaPage && hasPaperApproval) {
      setStockMenuOpen(true);
      setEVisaSubMenuOpen(true);
    } else if (isDailyTeamReportPage) {
      setStockMenuOpen(true);
      setDailyReportSubMenuOpen(true);
    } else if (isK2OperationPage) {
      setStockMenuOpen(true);
      setReportSubMenuOpen(true);
      setK2OperationSubMenuOpen(true);
    } else if (isTeamOperationPage) {
      setStockMenuOpen(true);
      setReportSubMenuOpen(true);
      setTeamOperationSubMenuOpen(true);
    } else if (isReportPage) {
      setStockMenuOpen(true);
      setReportSubMenuOpen(true);
    } else if (isStubProposalPage) {
      setStockMenuOpen(true);
      setStubProposalSubMenuOpen(true);
    } else if (isStockPage) {
      setStockMenuOpen(true);
    }
  }, [activePage, isStickerPage, isEVisaPage, isDailyTeamReportPage, isK2OperationPage, isTeamOperationPage, isReportPage, isStubProposalPage, isStockPage, hasPaperApproval]);

  const handleNav = (page: string, catType?: CategoryType) => {
    onNavigate(page, catType);
    setMobileOpen(false);
  };

  return (
    <>
      {/* Mobile menu trigger */}
      <div className="lg:hidden fixed top-3 left-3 z-30">
        <button
          onClick={() => setMobileOpen(!mobileOpen)}
          className="bg-[#1E3A8A] text-[#FDE047] p-2 rounded-lg shadow-md border border-[#C6A15B]/40 focus:outline-none"
        >
          {mobileOpen ? <X className="w-5 h-5" /> : <Menu className="w-5 h-5" />}
        </button>
      </div>

      {/* Backdrop for mobile */}
      {mobileOpen && (
        <div
          onClick={() => setMobileOpen(false)}
          className="lg:hidden fixed inset-0 bg-[#172554]/60 backdrop-blur-xs z-30"
        />
      )}

      {/* Sidebar container */}
      <aside
        className={`print:hidden fixed lg:static top-0 left-0 bottom-0 h-screen lg:h-full bg-[#1E3A8A] text-white flex flex-col shrink-0 z-40 transition-all duration-200 ease-in-out border-r border-[#C6A15B]/20 shadow-xl lg:shadow-none overflow-hidden select-none ${
          mobileOpen ? 'translate-x-0 w-64' : '-translate-x-full lg:translate-x-0'
        } ${isSidebarCollapsed ? 'lg:w-[68px]' : 'lg:w-64'}`}
      >
        {/* Brand Header */}
        <div className="flex items-center justify-between px-3 py-4 border-b border-white/10 shrink-0 select-none">
          <div className="flex items-center gap-3 overflow-hidden">
            <div className="w-10 h-10 rounded-full border-2 border-[#C6A15B] flex items-center justify-center bg-[#172554] shrink-0 shadow-inner">
              <Shield className="w-5 h-5 text-[#FDE047]" />
            </div>
            {!isSidebarCollapsed && (
              <div className="min-w-0">
                <p className="text-sm font-bold text-[#FDE047] leading-tight tracking-normal truncate">
                  ប្រព័ន្ធគ្រប់គ្រង
                </p>
                <p className="text-[11px] font-medium text-white/75 mt-0.5 truncate">ទិន្នន័យមន្ត្រី</p>
              </div>
            )}
          </div>
          {/* Desktop Collapse / Expand Toggle Button */}
          <button
            type="button"
            onClick={toggleSidebarCollapse}
            className="hidden lg:flex items-center justify-center w-8 h-8 rounded-lg text-slate-100 hover:text-[#FDE047] hover:bg-white/10 transition cursor-pointer shrink-0"
            title={isSidebarCollapsed ? 'ពង្រីករបារចំហៀង (Expand Sidebar)' : 'បង្រួមរបារចំហៀង (Collapse Sidebar)'}
          >
            {isSidebarCollapsed ? (
              <PanelLeftOpen className="w-4 h-4 text-[#FDE047]" />
            ) : (
              <PanelLeftClose className="w-4 h-4" />
            )}
          </button>
        </div>

        {/* Navigation items */}
        <nav className="flex-1 py-4 text-sm overflow-y-auto px-2 space-y-1 custom-sidebar-scroll">
          {/* Dashboard */}
          <button
            onClick={() => handleNav('dashboard')}
            className={`w-full text-left flex items-center gap-3 px-4 py-2.5 rounded-lg transition font-medium ${
              activePage === 'dashboard'
                ? 'bg-[#C6A15B]/25 text-[#FDE047] border-r-3 border-[#C6A15B]'
                : 'text-slate-50 font-semibold hover:bg-white/5 hover:text-white'
            }`}
          >
            <LayoutDashboard className="w-4 h-4 text-[#C6A15B]" />
            <span>ទំព័រដើម</span>
          </button>

          {/* Daily Operations by Team (Moved from inside stock) */}
          {isDailyOpsVisible && (
            <button
              onClick={() => handleNav('stockStickerDailyTeam')}
              className={`w-full text-left flex items-center gap-3 px-4 py-2.5 rounded-lg transition font-medium cursor-pointer ${
                activePage === 'stockStickerDailyTeam'
                  ? 'bg-[#C6A15B]/25 text-[#FDE047] border-r-3 border-[#C6A15B]'
                  : 'text-slate-50 font-semibold hover:bg-white/5 hover:text-white'
              }`}
            >
              <Table className="w-4 h-4 text-[#C6A15B]" />
              <span>ប្រតិបត្តិការប្រចាំថ្ងៃតាមក្រុម</span>
            </button>
          )}

          {/* Refusal and Deportation */}
          {isRefusalVisible && (
            <div>
              <button
                onClick={() => setRefusalMenuOpen(!refusalMenuOpen)}
                className="w-full flex items-center justify-between gap-3 px-4 py-2.5 rounded-lg text-slate-50 font-semibold hover:bg-white/5 hover:text-white transition font-medium text-left cursor-pointer"
              >
                <div className="flex items-center gap-3">
                  <ShieldAlert className="w-4 h-4 text-[#C6A15B]" />
                  <span>បដិសេធនិងបញ្ជូនចេញ</span>
                </div>
                <ChevronDown
                  className={`w-4 h-4 transition-transform duration-200 ${
                    refusalMenuOpen ? 'rotate-180 text-[#C6A15B]' : 'text-slate-300'
                  }`}
                />
              </button>

              {refusalMenuOpen && (
                <div className="pl-6 pt-1 pb-2 space-y-1 bg-black/10 rounded-lg my-1">
                  <button
                    onClick={() => handleNav('refusalDeportationManager')}
                    className={`w-full text-left flex items-center gap-2.5 px-3 py-2 rounded-md text-xs transition cursor-pointer ${
                      activePage === 'refusalDeportationManager'
                        ? 'bg-[#C6A15B]/30 text-[#FDE047] font-semibold'
                        : 'text-slate-100 hover:text-white hover:bg-white/5'
                    }`}
                  >
                    <FileCheck className="w-3.5 h-3.5 text-[#C6A15B]" />
                    <span>ទម្រង់បដិសេធ</span>
                  </button>
                  <button
                    onClick={() => handleNav('refusalNationalityData')}
                    className={`w-full text-left flex items-center gap-2.5 px-3 py-2 rounded-md text-xs transition cursor-pointer ${
                      activePage === 'refusalNationalityData'
                        ? 'bg-[#C6A15B]/30 text-[#FDE047] font-semibold'
                        : 'text-slate-100 hover:text-white hover:bg-white/5'
                    }`}
                  >
                    <Users className="w-3.5 h-3.5 text-[#C6A15B]" />
                    <span>ទិន្នន័យសញ្ជាតិបដិសេធ</span>
                  </button>
                  <button
                    onClick={() => handleNav('refusalTotalSummaryReport')}
                    className={`w-full text-left flex items-center gap-2.5 px-3 py-2 rounded-md text-xs transition cursor-pointer ${
                      activePage === 'refusalTotalSummaryReport'
                        ? 'bg-[#C6A15B]/30 text-[#FDE047] font-semibold'
                        : 'text-slate-100 hover:text-white hover:bg-white/5'
                    }`}
                  >
                    <FileText className="w-3.5 h-3.5 text-[#C6A15B]" />
                    <span>សរុបទិន្នន័យបដិសេធ</span>
                  </button>
                </div>
              )}
            </div>
          )}

          {/* Account Management (Secondary Role Only by default, toggleable to let teams see) */}
          {isUserMenuVisible && (
            <div>
              <button
                onClick={() => setAcctMenuOpen(!acctMenuOpen)}
                className="w-full flex items-center justify-between gap-3 px-4 py-2.5 rounded-lg text-slate-50 font-semibold hover:bg-white/5 hover:text-white transition font-medium text-left"
              >
                <div className="flex items-center gap-3">
                  <Users className="w-4 h-4 text-[#C6A15B]" />
                  <span>គណនីអ្នកប្រើប្រាស់</span>
                </div>
                <ChevronDown
                  className={`w-4 h-4 transition-transform duration-200 ${
                    acctMenuOpen ? 'rotate-180 text-[#C6A15B]' : 'text-slate-300'
                  }`}
                />
              </button>

              {acctMenuOpen && (
                <div className="pl-6 pt-1 pb-2 space-y-1 bg-black/10 rounded-lg my-1">
                  <button
                    onClick={() => handleNav('createUser')}
                    className={`w-full text-left flex items-center gap-2.5 px-3 py-2 rounded-md text-xs transition ${
                      activePage === 'createUser'
                        ? 'bg-[#C6A15B]/30 text-[#FDE047] font-semibold'
                        : 'text-slate-100 hover:text-white hover:bg-white/5'
                    }`}
                  >
                    <UserPlus className="w-3.5 h-3.5 text-[#C6A15B]" />
                    <span>បង្កើតគណនីប្រើប្រាស់</span>
                  </button>
                  {isUserListVisible && (
                    <button
                      onClick={() => handleNav('userList')}
                      className={`w-full text-left flex items-center gap-2.5 px-3 py-2 rounded-md text-xs transition ${
                        activePage === 'userList'
                          ? 'bg-[#C6A15B]/30 text-[#FDE047] font-semibold'
                          : 'text-slate-100 hover:text-white hover:bg-white/5'
                      }`}
                    >
                      <Users className="w-3.5 h-3.5 text-[#C6A15B]" />
                      <span>ទិន្នន័យគណនី</span>
                    </button>
                  )}
                </div>
              )}
            </div>
          )}

          <div className="border-t border-white/10 my-2" />
          {/* Officer Submenu (ព័ត៌មានមន្ត្រី) */}
          {isOfficerMenuVisible && (
            <div>
              <button
                onClick={() => setOfficerMenuOpen(!officerMenuOpen)}
                className="w-full flex items-center justify-between gap-3 px-4 py-2.5 rounded-lg text-slate-50 font-semibold hover:bg-white/5 hover:text-white transition font-medium text-left"
              >
                <div className="flex items-center gap-3">
                  <UserCheck className="w-4 h-4 text-[#C6A15B]" />
                  <span>ព័ត៌មានមន្ត្រី</span>
                </div>
                <ChevronDown
                  className={`w-4 h-4 transition-transform duration-200 ${
                    officerMenuOpen ? 'rotate-180 text-[#C6A15B]' : 'text-slate-300'
                  }`}
                />
              </button>

              {officerMenuOpen && (
                <div className="pl-6 pt-1 pb-2 space-y-1 bg-black/10 rounded-lg my-1">
                  <button
                    onClick={() => handleNav('officerForm')}
                    className={`w-full text-left flex items-center gap-2.5 px-3 py-2 rounded-md text-xs transition ${
                      activePage === 'officerForm'
                        ? 'bg-[#C6A15B]/30 text-[#FDE047] font-semibold'
                        : 'text-slate-100 hover:text-white hover:bg-white/5'
                    }`}
                  >
                    <FileText className="w-3.5 h-3.5 text-[#C6A15B]" />
                    <span>ទម្រង់បញ្ចូលមន្ត្រី</span>
                  </button>

                  <button
                    onClick={() => handleNav('officerList')}
                    className={`w-full text-left flex items-center justify-between gap-2 px-3 py-2 rounded-md text-xs transition ${
                      activePage === 'officerList'
                        ? 'bg-[#C6A15B]/30 text-[#FDE047] font-semibold'
                        : 'text-slate-100 hover:text-white hover:bg-white/5'
                    }`}
                  >
                    <div className="flex items-center gap-2.5">
                      <Table className="w-3.5 h-3.5 text-[#C6A15B]" />
                      <span>{isOfficeRestricted ? 'ទិន្នន័យមន្ត្រីរបស់ក្រុម' : 'ទិន្នន័យមន្ត្រី'}</span>
                    </div>
                    {isOfficeRestricted && (
                      <span className="text-[10px] bg-amber-500/20 text-amber-300 px-1.5 py-0.5 rounded border border-amber-500/30">
                        ក្រុម
                      </span>
                    )}
                  </button>
                </div>
              )}
            </div>
          )}

          {/* Visa Operations Submenu (ការងារទិដ្ឋាការ) - Secondary / Office Only by default, toggleable to let teams see */}
          {isVisaWorkVisible && (
            <>
              <div className="border-t border-white/10 my-2" />
              <div>
                <button
                  onClick={() => setVisaMenuOpen(!visaMenuOpen)}
                  className="w-full flex items-center justify-between gap-3 px-4 py-2.5 rounded-lg text-slate-50 font-semibold hover:bg-white/5 hover:text-white transition font-medium text-left cursor-pointer"
                >
                  <div className="flex items-center gap-3">
                    <RefreshCw className="w-4 h-4 text-[#C6A15B]" />
                    <span>ការងារទិដ្ឋាការ</span>
                  </div>
                  <ChevronDown
                    className={`w-4 h-4 transition-transform duration-200 ${
                      visaMenuOpen ? 'rotate-180 text-[#C6A15B]' : 'text-slate-300'
                    }`}
                  />
                </button>

                {visaMenuOpen && (
                  <div className="pl-6 pt-1 pb-2 space-y-1 bg-black/10 rounded-lg my-1">
                    <button
                      onClick={() => handleNav('visaForm')}
                      className={`w-full text-left flex items-center gap-2.5 px-3 py-2 rounded-md text-xs transition ${
                        activePage === 'visaForm'
                          ? 'bg-[#C6A15B]/30 text-[#FDE047] font-semibold'
                          : 'text-slate-100 hover:text-white hover:bg-white/5'
                      }`}
                    >
                      <RefreshCw className="w-3.5 h-3.5 text-[#C6A15B]" />
                      <span>ប្តូរប្រភេទទិដ្ឋាការ</span>
                    </button>

                    <button
                      onClick={() => handleNav('visaList')}
                      className={`w-full text-left flex items-center gap-2.5 px-3 py-2 rounded-md text-xs transition ${
                        activePage === 'visaList'
                          ? 'bg-[#C6A15B]/30 text-[#FDE047] font-semibold'
                          : 'text-slate-100 hover:text-white hover:bg-white/5'
                      }`}
                    >
                      <FileSpreadsheet className="w-3.5 h-3.5 text-[#C6A15B]" />
                      <span>ទិន្នន័យទិដ្ឋាការ</span>
                    </button>

                    <button
                      onClick={() => handleNav('visaSummary')}
                      className={`w-full text-left flex items-center gap-2.5 px-3 py-2 rounded-md text-xs transition ${
                        activePage === 'visaSummary'
                          ? 'bg-[#C6A15B]/30 text-[#FDE047] font-semibold'
                          : 'text-slate-100 hover:text-white hover:bg-white/5'
                      }`}
                    >
                      <BarChart3 className="w-3.5 h-3.5 text-[#C6A15B]" />
                      <span>សរុបទិន្នន័យប្តូរទិដ្ឋាការ</span>
                    </button>

                    <button
                      onClick={() => handleNav('visaDateSearch')}
                      className={`w-full text-left flex items-center gap-2.5 px-3 py-2 rounded-md text-xs transition ${
                        activePage === 'visaDateSearch'
                          ? 'bg-[#C6A15B]/30 text-[#FDE047] font-semibold'
                          : 'text-slate-100 hover:text-white hover:bg-white/5'
                      }`}
                    >
                      <CalendarSearch className="w-3.5 h-3.5 text-[#C6A15B]" />
                      <span>ស្វែងរកតាមកាលបរិច្ឆេទ</span>
                    </button>
                  </div>
                )}
              </div>
            </>
          )}

          <div className="border-t border-white/10 my-2" />

          {/* Stock Work Submenu (ការងារស្តុក) */}
          {isOfficeStockVisible && (
            <div>
              <button
                onClick={() => setStockMenuOpen(!stockMenuOpen)}
              className="w-full flex items-center justify-between gap-3 px-4 py-2.5 rounded-lg text-slate-50 font-semibold hover:bg-white/5 hover:text-white transition font-medium text-left cursor-pointer"
            >
              <div className="flex items-center gap-3">
                <Boxes className="w-4 h-4 text-[#C6A15B]" />
                <span>{isSecondary ? 'ការងារស្តុកការិយាល័យ' : 'ការងារស្តុកក្រុម'}</span>
              </div>
              <ChevronDown
                className={`w-4 h-4 transition-transform duration-200 ${
                  stockMenuOpen ? 'rotate-180 text-[#C6A15B]' : 'text-slate-300'
                }`}
              />
            </button>

            {stockMenuOpen && (
              <div className="pl-6 pt-1 pb-2 space-y-1 bg-black/10 rounded-lg my-1">
                {/* សន្លឹកទិដ្ឋាការស្អិត group */}
                <div>
                  <button
                    onClick={() => {
                      setStickerSubMenuOpen(!stickerSubMenuOpen);
                    }}
                    className={`w-full text-left flex items-center justify-between gap-2 px-3 py-2 rounded-md text-xs transition cursor-pointer ${
                      isStickerPage
                        ? 'bg-[#C6A15B]/30 text-[#FDE047] font-semibold'
                        : 'text-slate-100 hover:text-white hover:bg-white/5'
                    }`}
                  >
                    <div className="flex items-center gap-2.5">
                      <Package className="w-3.5 h-3.5 text-[#C6A15B]" />
                      <span>សន្លឹកទិដ្ឋាការស្អិត</span>
                    </div>
                    <ChevronDown
                      className={`w-3 h-3 transition-transform duration-200 ${
                        stickerSubMenuOpen ? 'rotate-180 text-[#C6A15B]' : 'text-white/40'
                      }`}
                    />
                  </button>

                  {stickerSubMenuOpen && (
                    <div className="pl-5 pt-1 space-y-1">
                      {isSecondary && (
                        <button
                          onClick={() => handleNav('stockStickerStubCollection')}
                          className={`w-full text-left flex items-center gap-2 px-2.5 py-1.5 rounded-md text-[11px] transition cursor-pointer ${
                            activePage === 'stockStickerStubCollection'
                              ? 'bg-[#C6A15B]/40 text-white font-bold'
                              : 'text-slate-200 font-medium hover:text-white hover:bg-white/5'
                          }`}
                        >
                          <BookmarkCheck className="w-3 h-3 text-[#C6A15B]" />
                          <span>ប្រមូលគល់សន្លឹកទិដ្ឋាការ</span>
                        </button>
                      )}

                      <button
                        onClick={() => handleNav('stockSticker')}
                        className={`w-full text-left flex items-center gap-2 px-2.5 py-1.5 rounded-md text-[11px] transition cursor-pointer ${
                          activePage === 'stockSticker'
                            ? 'bg-[#C6A15B]/40 text-white font-bold'
                            : 'text-slate-200 font-medium hover:text-white hover:bg-white/5'
                        }`}
                      >
                        <FileText className="w-3 h-3 text-[#C6A15B]" />
                        <span>ទម្រង់កត់ត្រាស្តុកសន្លឹក</span>
                      </button>

                      <button
                        onClick={() => handleNav('stockStickerData')}
                        className={`w-full text-left flex items-center gap-2 px-2.5 py-1.5 rounded-md text-[11px] transition cursor-pointer ${
                          activePage === 'stockStickerData'
                            ? 'bg-[#C6A15B]/40 text-white font-bold'
                            : 'text-slate-200 font-medium hover:text-white hover:bg-white/5'
                        }`}
                      >
                        <FileSpreadsheet className="w-3 h-3 text-[#C6A15B]" />
                        <span>ទិន្នន័យសន្លឹកទិដ្ឋាការ</span>
                      </button>

                      {isSecondary && (
                        <button
                          onClick={() => handleNav('stockStickerOfficeReport')}
                          className={`w-full text-left flex items-center gap-2 px-2.5 py-1.5 rounded-md text-[11px] transition cursor-pointer ${
                            activePage === 'stockStickerOfficeReport'
                              ? 'bg-[#C6A15B]/40 text-white font-bold'
                              : 'text-slate-200 font-medium hover:text-white hover:bg-white/5'
                          }`}
                        >
                          <Calendar className="w-3 h-3 text-[#C6A15B]" />
                          <span>របាយការណ៍ស្តុកការិយាល័យ</span>
                        </button>
                      )}

                      {isSecondary && (
                        <button
                          onClick={() => handleNav('stockStickerTeamReport')}
                          className={`w-full text-left flex items-center gap-2 px-2.5 py-1.5 rounded-md text-[11px] transition cursor-pointer ${
                            activePage === 'stockStickerTeamReport'
                              ? 'bg-[#C6A15B]/40 text-white font-bold'
                              : 'text-slate-200 font-medium hover:text-white hover:bg-white/5'
                          }`}
                        >
                          <Layers className="w-3 h-3 text-[#C6A15B]" />
                          <span>របាយការណ៍តាមក្រុម</span>
                        </button>
                      )}

                      {isSecondary && (
                        <button
                          onClick={() => handleNav('stockStickerStatus')}
                          className={`w-full text-left flex items-center gap-2 px-2.5 py-1.5 rounded-md text-[11px] transition cursor-pointer ${
                            activePage === 'stockStickerStatus'
                              ? 'bg-[#C6A15B]/40 text-white font-bold'
                              : 'text-slate-200 font-medium hover:text-white hover:bg-white/5'
                          }`}
                        >
                          <Boxes className="w-3 h-3 text-[#C6A15B]" />
                          <span>ស្តុកជាក់ស្តែងសម្រាប់បើកផ្តល់</span>
                        </button>
                      )}
                    </div>
                  )}
                </div>

                {/* ក្រដាសអនុម័ត (Approval Paper Submenu) */}
                {hasPaperApproval && (
                  <div>
                    <button
                      onClick={() => setEVisaSubMenuOpen(!eVisaSubMenuOpen)}
                      className={`w-full text-left flex items-center justify-between gap-2 px-3 py-2 rounded-md text-xs transition cursor-pointer ${
                        isEVisaPage
                          ? 'bg-[#C6A15B]/30 text-[#FDE047] font-semibold'
                          : 'text-slate-100 hover:text-white hover:bg-white/5'
                      }`}
                    >
                      <div className="flex items-center gap-2.5">
                        <FileCheck className="w-3.5 h-3.5 text-[#C6A15B]" />
                        <span>ក្រដាសអនុម័ត</span>
                      </div>
                      <ChevronDown
                        className={`w-3 h-3 transition-transform duration-200 ${
                          eVisaSubMenuOpen ? 'rotate-180 text-[#C6A15B]' : 'text-white/40'
                        }`}
                      />
                    </button>

                    {eVisaSubMenuOpen && (
                      <div className="pl-5 pt-1 space-y-1">
                        {isSecondary && (
                          <button
                            onClick={() => handleNav('stockEVisa')}
                            className={`w-full text-left flex items-center gap-2 px-2.5 py-1.5 rounded-md text-[11px] transition cursor-pointer ${
                              activePage === 'stockEVisa'
                                ? 'bg-[#C6A15B]/40 text-white font-bold'
                                : 'text-slate-200 font-medium hover:text-white hover:bg-white/5'
                            }`}
                          >
                            <FileText className="w-3 h-3 text-[#C6A15B]" />
                            <span>ទម្រង់កត់ត្រាស្តុកក្រដាស</span>
                          </button>
                        )}

                        <button
                          onClick={() => handleNav('stockEVisaTeamUse')}
                          className={`w-full text-left flex items-center gap-2 px-2.5 py-1.5 rounded-md text-[11px] transition cursor-pointer ${
                            activePage === 'stockEVisaTeamUse'
                              ? 'bg-[#C6A15B]/40 text-white font-bold'
                              : 'text-slate-200 font-medium hover:text-white hover:bg-white/5'
                          }`}
                        >
                          <UserCheck className="w-3 h-3 text-[#C6A15B]" />
                          <span>ទម្រង់ប្រើប្រាស់ក្រដាសតាមក្រុម</span>
                        </button>

                        <button
                          onClick={() => handleNav('stockEVisaData')}
                          className={`w-full text-left flex items-center gap-2 px-2.5 py-1.5 rounded-md text-[11px] transition cursor-pointer ${
                            activePage === 'stockEVisaData'
                              ? 'bg-[#C6A15B]/40 text-white font-bold'
                              : 'text-slate-200 font-medium hover:text-white hover:bg-white/5'
                          }`}
                        >
                          <FileSpreadsheet className="w-3 h-3 text-[#C6A15B]" />
                          <span>ទិន្នន័យក្រដាសអនុម័ត</span>
                        </button>

                        {!isSecondary && (
                          <button
                            onClick={() => handleNav('stockEVisaTeamRobokReport')}
                            className={`w-full text-left flex items-center gap-2 px-2.5 py-1.5 rounded-md text-[11px] transition cursor-pointer ${
                              activePage === 'stockEVisaTeamRobokReport'
                                ? 'bg-[#C6A15B]/40 text-white font-bold'
                                : 'text-slate-200 font-medium hover:text-white hover:bg-white/5'
                            }`}
                          >
                            <FileCheck className="w-3 h-3 text-[#C6A15B]" />
                            <span>របកក្រដាសអនុម័តក្រុម</span>
                          </button>
                        )}
                      </div>
                    )}
                  </div>
                )}

                {/* របាយការណ៍ Submenu Group */}
                <div>
                  <button
                    onClick={() => {
                      setReportSubMenuOpen(!reportSubMenuOpen);
                    }}
                    className={`w-full text-left flex items-center justify-between gap-2 px-3 py-2 rounded-md text-xs transition cursor-pointer ${
                      isReportPage
                        ? 'bg-[#C6A15B]/30 text-[#FDE047] font-semibold'
                        : 'text-slate-100 hover:text-white hover:bg-white/5'
                    }`}
                  >
                    <div className="flex items-center gap-2.5">
                      <BarChart3 className="w-3.5 h-3.5 text-[#C6A15B]" />
                      <span>របាយការណ៍</span>
                    </div>
                    <ChevronDown
                      className={`w-3 h-3 transition-transform duration-200 ${
                        reportSubMenuOpen ? 'rotate-180 text-[#C6A15B]' : 'text-white/40'
                      }`}
                    />
                  </button>

                  {reportSubMenuOpen && (
                    <div className="pl-5 pt-1 space-y-1">
                      {/* របក.សរុបស្តុកក្រុម - Only visible for Team user */}
                      {!isSecondary && (
                        <button
                          onClick={() => handleNav('stockTotalSummaryReport')}
                          className={`w-full text-left flex items-center gap-2 px-2.5 py-1.5 rounded-md text-[11px] transition cursor-pointer ${
                            activePage === 'stockTotalSummaryReport'
                              ? 'bg-[#C6A15B]/40 text-white font-bold'
                              : 'text-slate-200 font-medium hover:text-white hover:bg-white/5'
                          }`}
                          title="របក.សរុបស្តុកក្រុម"
                        >
                          <Users className="w-3 h-3 text-[#C6A15B]" />
                          <span>របក.សរុបស្តុកក្រុម</span>
                        </button>
                      )}

                      {/* របក.ស្តុកសរុបរួម - Only visible for Office user (isSecondary) */}
                      {isSecondary && (
                        <button
                          onClick={() => handleNav('stockOfficeWorkSummaryReport')}
                          className={`w-full text-left flex items-center gap-2 px-2.5 py-1.5 rounded-md text-[11px] transition cursor-pointer ${
                            activePage === 'stockOfficeWorkSummaryReport'
                              ? 'bg-[#C6A15B]/40 text-white font-bold'
                              : 'text-slate-200 font-medium hover:text-white hover:bg-white/5'
                          }`}
                          title="របក.ស្តុកសរុបរួម"
                        >
                          <Building2 className="w-3 h-3 text-[#C6A15B]" />
                          <span>របក.ស្តុកសរុបរួម</span>
                        </button>
                      )}

                      {/* របក.សរុបរួមរំកិល - Only visible for Office user (isSecondary) */}
                      {isSecondary && (
                        <button
                          id="menu-stock-office-rolling-summary-report"
                          onClick={() => handleNav('stockOfficeRollingSummaryReport')}
                          className={`w-full text-left flex items-center gap-2 px-2.5 py-1.5 rounded-md text-[11px] transition cursor-pointer ${
                            activePage === 'stockOfficeRollingSummaryReport'
                              ? 'bg-[#C6A15B]/40 text-white font-bold'
                              : 'text-slate-200 font-medium hover:text-white hover:bg-white/5'
                          }`}
                          title="របក.សរុបរួមរំកិល"
                        >
                          <FileSpreadsheet className="w-3 h-3 text-[#C6A15B]" />
                          <span>របក.សរុបរួមរំកិល</span>
                        </button>
                      )}

                      {/* របាយការណ៍ប្រចាំថ្ងៃ Group */}
                      <div>
                        <button
                          onClick={() => setDailyReportSubMenuOpen(!dailyReportSubMenuOpen)}
                          className={`w-full text-left flex items-center justify-between gap-2 px-2.5 py-1.5 rounded-md text-[11px] transition cursor-pointer ${
                            isDailyTeamReportPage
                              ? 'bg-[#C6A15B]/40 text-white font-bold'
                              : 'text-slate-200 font-medium hover:text-white hover:bg-white/5'
                          }`}
                          title="របាយការណ៍ប្រចាំថ្ងៃ"
                        >
                          <div className="flex items-center gap-2">
                            <Calendar className="w-3 h-3 text-[#C6A15B]" />
                            <span>របាយការណ៍ប្រចាំថ្ងៃ</span>
                          </div>
                          <ChevronDown
                            className={`w-3 h-3 transition-transform duration-200 ${
                              dailyReportSubMenuOpen ? 'rotate-180 text-[#C6A15B]' : 'text-white/40'
                            }`}
                          />
                        </button>

                        {dailyReportSubMenuOpen && (
                          <div className="pl-4 pt-1 space-y-1 border-l border-[#C6A15B]/20 ml-2">
                            {/* ១. ស្ថិតិផ្តល់ប្រចាំថ្ងៃរបស់ក្រុម */}
                            <button
                              onClick={() => handleNav('singleDayTeamVisaIssuanceReport')}
                              className={`w-full text-left flex items-center gap-2 px-2 py-1.5 rounded-md text-[10.5px] transition cursor-pointer ${
                                activePage === 'singleDayTeamVisaIssuanceReport'
                                  ? 'bg-[#C6A15B]/40 text-white font-bold'
                                  : 'text-slate-200 font-medium hover:text-white hover:bg-white/5'
                              }`}
                              title="ស្ថិតិផ្តល់ប្រចាំថ្ងៃរបស់ក្រុម"
                            >
                              <FileText className="w-2.5 h-2.5 text-[#C6A15B]" />
                              <span>ស្ថិតិផ្តល់ប្រចាំថ្ងៃរបស់ក្រុម</span>
                            </button>

                            {/* ២. របាយការណ៍សន្លឹកទិដ្ឋាការ (New A4 Report) */}
                            <button
                              onClick={() => handleNav('dailyStickerVisaOperationReport')}
                              className={`w-full text-left flex items-center gap-2 px-2 py-1.5 rounded-md text-[10.5px] transition cursor-pointer ${
                                activePage === 'dailyStickerVisaOperationReport'
                                  ? 'bg-[#C6A15B]/40 text-white font-bold'
                                  : 'text-slate-200 font-medium hover:text-white hover:bg-white/5'
                              }`}
                              title="របាយការណ៍សន្លឹកទិដ្ឋាការ"
                            >
                              <FileText className="w-2.5 h-2.5 text-[#C6A15B]" />
                              <span>របាយការណ៍សន្លឹកទិដ្ឋាការ</span>
                            </button>
                          </div>
                        )}
                      </div>



                      {/* ស្ថិតិសរុបប្រចាំថ្ងៃរបស់ក្រុម - របាយការណ៍ការងារស្តុកការិយាល័យ (Only visible for Office user) */}
                      {isSecondary && (
                        <button
                          onClick={() => handleNav('stockOfficeDailyTeamStatsReport')}
                          className={`w-full text-left flex items-center gap-2 px-2.5 py-1.5 rounded-md text-[11px] transition cursor-pointer ${
                            activePage === 'stockOfficeDailyTeamStatsReport'
                              ? 'bg-[#C6A15B]/40 text-white font-bold'
                              : 'text-slate-200 font-medium hover:text-white hover:bg-white/5'
                          }`}
                          title="ស្ថិតិសរុបប្រចាំថ្ងៃរបស់ក្រុម"
                        >
                          <BarChart3 className="w-3 h-3 text-[#C6A15B]" />
                          <span>ស្ថិតិសរុបប្រចាំថ្ងៃរបស់ក្រុម</span>
                        </button>
                      )}

                      {/* តារាងប្រតិបត្តិការរបស់ក២ Nested Submenu Group - Only visible for Office user (isSecondary) */}
                      {isSecondary && (
                        <div>
                          <button
                            onClick={() => setK2OperationSubMenuOpen(!k2OperationSubMenuOpen)}
                            className={`w-full text-left flex items-center justify-between gap-2 px-2.5 py-1.5 rounded-md text-[11px] transition cursor-pointer ${
                              isK2OperationPage
                                ? 'bg-[#C6A15B]/30 text-[#FDE047] font-bold'
                                : 'text-slate-100 hover:text-white hover:bg-white/5'
                            }`}
                            title="តារាងប្រតិបត្តិការរបស់ក២"
                          >
                            <div className="flex items-center gap-2">
                              <Building2 className="w-3 h-3 text-[#C6A15B]" />
                              <span>តារាងប្រតិបត្តិការរបស់ក២</span>
                            </div>
                            <ChevronDown
                              className={`w-3 h-3 transition-transform duration-200 ${
                                k2OperationSubMenuOpen ? 'rotate-180 text-[#C6A15B]' : 'text-white/40'
                              }`}
                            />
                          </button>

                          {k2OperationSubMenuOpen && (
                            <div className="pl-4 pt-1 space-y-1 border-l border-[#C6A15B]/20 ml-2">
                              {/* តារាងទិដ្ឋាការបើកពីក១ */}
                              <button
                                onClick={() => handleNav('stockStickerK2ReceiveFromK1Report')}
                                className={`w-full text-left flex items-center gap-2 px-2 py-1.5 rounded-md text-[10.5px] transition cursor-pointer ${
                                  activePage === 'stockStickerK2ReceiveFromK1Report'
                                    ? 'bg-[#C6A15B]/40 text-white font-bold'
                                    : 'text-slate-200 font-medium hover:text-white hover:bg-white/5'
                                }`}
                              >
                                <Calendar className="w-2.5 h-2.5 text-[#C6A15B]" />
                                <span>តារាងទិដ្ឋាការបើកពីក១</span>
                              </button>

                              {/* តារាងបើកផ្តល់តាមក្រុម */}
                              <button
                                onClick={() => handleNav('stockStickerK2DistributeToTeamsReport')}
                                className={`w-full text-left flex items-center gap-2 px-2 py-1.5 rounded-md text-[10.5px] transition cursor-pointer ${
                                  activePage === 'stockStickerK2DistributeToTeamsReport'
                                    ? 'bg-[#C6A15B]/40 text-white font-bold'
                                    : 'text-slate-200 font-medium hover:text-white hover:bg-white/5'
                                }`}
                              >
                                <Calendar className="w-2.5 h-2.5 text-[#C6A15B]" />
                                <span>តារាងបើកផ្តល់តាមក្រុម</span>
                              </button>
                            </div>
                          )}
                        </div>
                      )}

                      {/* តារាងប្រតិបត្តិការរបស់ក្រុម Nested Submenu Group */}
                      <div>
                        <button
                          onClick={() => setTeamOperationSubMenuOpen(!teamOperationSubMenuOpen)}
                          className={`w-full text-left flex items-center justify-between gap-2 px-2.5 py-1.5 rounded-md text-[11px] transition cursor-pointer ${
                            isTeamOperationPage
                              ? 'bg-[#C6A15B]/30 text-[#FDE047] font-bold'
                              : 'text-slate-100 hover:text-white hover:bg-white/5'
                          }`}
                          title="តារាងប្រតិបត្តិការរបស់ក្រុម"
                        >
                          <div className="flex items-center gap-2">
                            <Table className="w-3 h-3 text-[#C6A15B]" />
                            <span>តារាងប្រតិបត្តិការរបស់ក្រុម</span>
                          </div>
                          <ChevronDown
                            className={`w-3 h-3 transition-transform duration-200 ${
                              teamOperationSubMenuOpen ? 'rotate-180 text-[#C6A15B]' : 'text-white/40'
                            }`}
                          />
                        </button>

                        {teamOperationSubMenuOpen && (
                          <div className="pl-4 pt-1 space-y-1 border-l border-[#C6A15B]/20 ml-2">
                            {/* តារាងផ្តល់ប្រចាំឆ្នាំទៅក្រុម (ការិយាល័យ) ឬ តារាងទិដ្ឋាការទទួលបានពីក២ (ក្រុម) */}
                            <button
                              onClick={() => handleNav('stockStickerYearlyTeamDistributionReport')}
                              className={`w-full text-left flex items-center gap-2 px-2 py-1.5 rounded-md text-[10.5px] transition cursor-pointer ${
                                activePage === 'stockStickerYearlyTeamDistributionReport'
                                  ? 'bg-[#C6A15B]/40 text-white font-bold'
                                  : 'text-slate-200 font-medium hover:text-white hover:bg-white/5'
                              }`}
                            >
                              <Calendar className="w-2.5 h-2.5 text-[#C6A15B]" />
                              <span>{isSecondary ? 'តារាងផ្តល់ប្រចាំឆ្នាំទៅក្រុម' : 'តារាងទិដ្ឋាការទទួលបានពីក២'}</span>
                            </button>

                            {/* តារាងប្រើប្រាស់តាមក្រុម (ការិយាល័យ) ឬ តារាងប្រើប្រាស់របស់ក្រុម (ក្រុម) */}
                            <button
                              onClick={() => handleNav('stockStickerYearlyTeamUsageReport')}
                              className={`w-full text-left flex items-center gap-2 px-2 py-1.5 rounded-md text-[10.5px] transition cursor-pointer ${
                                activePage === 'stockStickerYearlyTeamUsageReport'
                                  ? 'bg-[#C6A15B]/40 text-white font-bold'
                                  : 'text-slate-200 font-medium hover:text-white hover:bg-white/5'
                              }`}
                            >
                              <Calendar className="w-2.5 h-2.5 text-[#C6A15B]" />
                              <span>{isSecondary ? 'តារាងប្រើប្រាស់តាមក្រុម' : 'តារាងប្រើប្រាស់របស់ក្រុម'}</span>
                            </button>

                            {/* តារាងស្ថិតិប្រចាំថ្ងៃរបស់ក្រុម (សម្រាប់តែការងារស្តុកក្រុម) */}
                            {!isSecondary && (
                              <button
                                onClick={() => handleNav('stockTeamDailyStatsReport')}
                                className={`w-full text-left flex items-center gap-2 px-2 py-1.5 rounded-md text-[10.5px] transition cursor-pointer ${
                                  activePage === 'stockTeamDailyStatsReport'
                                    ? 'bg-[#C6A15B]/40 text-white font-bold'
                                    : 'text-slate-200 font-medium hover:text-white hover:bg-white/5'
                                }`}
                                title="តារាងស្ថិតិប្រចាំថ្ងៃរបស់ក្រុម"
                              >
                                <BarChart3 className="w-2.5 h-2.5 text-[#C6A15B]" />
                                <span>តារាងស្ថិតិប្រចាំថ្ងៃរបស់ក្រុម</span>
                              </button>
                            )}

                            {/* ស្ថិតិផ្តល់ទិដ្ឋាការរបស់ក្រុម (រក្សាទុកតែក្នុងការងារស្តុកតាមក្រុមប៉ុណ្ណោះ) */}
                            {!isSecondary && (
                              <button
                                onClick={() => handleNav('teamVisaIssuanceStatsReport')}
                                className={`w-full text-left flex items-center gap-2 px-2 py-1.5 rounded-md text-[10.5px] transition cursor-pointer ${
                                  activePage === 'teamVisaIssuanceStatsReport'
                                    ? 'bg-[#C6A15B]/40 text-white font-bold'
                                    : 'text-slate-200 font-medium hover:text-white hover:bg-white/5'
                                }`}
                                title="ស្ថិតិផ្តល់ទិដ្ឋាការរបស់ក្រុម"
                              >
                                <FileText className="w-2.5 h-2.5 text-[#C6A15B]" />
                                <span>ស្ថិតិផ្តល់ទិដ្ឋាការរបស់ក្រុម</span>
                              </button>
                            )}

                            {/* ស្ថិតិរបស់ក្រុមប្រធានសាខា (រក្សាទុកតែក្នុងការងារស្តុកតាមក្រុមប៉ុណ្ណោះ) */}
                            {!isSecondary && (
                              <button
                                onClick={() => handleNav('branchChiefTeamVisaStatsReport')}
                                className={`w-full text-left flex items-center gap-2 px-2 py-1.5 rounded-md text-[10.5px] transition cursor-pointer ${
                                  activePage === 'branchChiefTeamVisaStatsReport'
                                    ? 'bg-[#C6A15B]/40 text-white font-bold'
                                    : 'text-slate-200 font-medium hover:text-white hover:bg-white/5'
                                }`}
                                title="ស្ថិតិរបស់ក្រុមប្រធានសាខា"
                              >
                                <FileText className="w-2.5 h-2.5 text-[#C6A15B]" />
                                <span>ស្ថិតិរបស់ក្រុមប្រធានសាខា</span>
                              </button>
                            )}
                          </div>
                        )}
                      </div>
                    </div>
                  )}
                </div>

                {/* សំណើប្រគល់គល់ទិដ្ឋាការ Submenu Group (សម្រាប់តែការងារស្តុកការិយាល័យ - Office តែប៉ុណ្ណោះ) */}
                {isSecondary && (
                  <div>
                    <button
                      onClick={() => setStubProposalSubMenuOpen(!stubProposalSubMenuOpen)}
                      className={`w-full text-left flex items-center justify-between gap-2 px-3 py-2 rounded-md text-xs transition cursor-pointer ${
                        isStubProposalPage
                          ? 'bg-[#C6A15B]/30 text-[#FDE047] font-semibold'
                          : 'text-slate-100 hover:text-white hover:bg-white/5'
                      }`}
                      title="សំណើប្រគល់គល់ទិដ្ឋាការ"
                    >
                      <div className="flex items-center gap-2.5">
                        <BookmarkCheck className="w-3.5 h-3.5 text-[#C6A15B]" />
                        <span>សំណើប្រគល់គល់ទិដ្ឋាការ</span>
                      </div>
                      <ChevronDown
                        className={`w-3 h-3 transition-transform duration-200 ${
                          stubProposalSubMenuOpen ? 'rotate-180 text-[#C6A15B]' : 'text-white/40'
                        }`}
                      />
                    </button>

                    {stubProposalSubMenuOpen && (
                      <div className="pl-5 pt-1 space-y-1">
                        {/* ផ្នែកស្នើសុំប្រគល់គល់ទិដ្ឋាការ */}
                        <button
                          onClick={() => handleNav('stockStubProposalSection')}
                          className={`w-full text-left flex items-center gap-2 px-2.5 py-1.5 rounded-md text-[11px] transition cursor-pointer ${
                            activePage === 'stockStubProposalSection'
                              ? 'bg-[#C6A15B]/40 text-white font-bold'
                              : 'text-slate-200 font-medium hover:text-white hover:bg-white/5'
                          }`}
                          title="ផ្នែកស្នើសុំប្រគល់គល់ទិដ្ឋាការ"
                        >
                          <Users className="w-3 h-3 text-[#C6A15B]" />
                          <span>ផ្នែកស្នើសុំប្រគល់គល់ទិដ្ឋាការ</span>
                        </button>

                        {/* ការិយាល័យស្នើសុំប្រគល់គល់ទិដ្ឋាការ */}
                        <button
                          onClick={() => handleNav('stockStubProposalOffice')}
                          className={`w-full text-left flex items-center gap-2 px-2.5 py-1.5 rounded-md text-[11px] transition cursor-pointer ${
                            activePage === 'stockStubProposalOffice'
                              ? 'bg-[#C6A15B]/40 text-white font-bold'
                              : 'text-slate-200 font-medium hover:text-white hover:bg-white/5'
                          }`}
                          title="ការិយាល័យស្នើសុំប្រគល់គល់ទិដ្ឋាការ"
                        >
                          <Building2 className="w-3 h-3 text-[#C6A15B]" />
                          <span>ការិយាល័យស្នើសុំប្រគល់គល់ទិដ្ឋាការ</span>
                        </button>
                      </div>
                    )}
                  </div>
                )}

                {/* របក.សរុបលោកនាយក */}
                {isSecondary && (
                  <button
                    onClick={() => handleNav('stockDirectorSummaryReport')}
                    className={`w-full text-left flex items-center gap-2.5 px-3 py-2 rounded-md text-xs transition cursor-pointer ${
                      activePage === 'stockDirectorSummaryReport'
                        ? 'bg-[#C6A15B]/30 text-[#FDE047] font-bold'
                        : 'text-slate-100 hover:text-white hover:bg-white/5'
                    }`}
                  >
                    <Award className="w-3.5 h-3.5 text-[#C6A15B]" />
                    <span>របក.សរុបលោកនាយក</span>
                  </button>
                )}
              </div>
            )}
          </div>
          )}

          {/* ស្វែងរកលេខកូដតាមក្រុម (Moved outside) */}
          {isSearchCodeVisible && (
            <button
              onClick={() => handleNav('stockSearchCodeByTeam')}
              className={`w-full text-left flex items-center gap-3 px-4 py-2.5 rounded-lg transition font-medium cursor-pointer ${
                activePage === 'stockSearchCodeByTeam'
                  ? 'bg-[#C6A15B]/25 text-[#FDE047] border-r-3 border-[#C6A15B]'
                  : 'text-slate-50 font-semibold hover:bg-white/5 hover:text-white'
              }`}
            >
              <Search className="w-4 h-4 text-[#C6A15B]" />
              <span>ស្វែងរកលេខកូដតាមក្រុម</span>
            </button>
          )}

          <div className="border-t border-white/10 my-2" />
          {/* Category Management Submenu (បញ្ចូលប្រភេទ) - Admin Only */}
          {isCategoriesVisible && (
            <div>
              <button
                onClick={() => setCatMenuOpen(!catMenuOpen)}
                className="w-full flex items-center justify-between gap-3 px-4 py-2.5 rounded-lg text-slate-50 font-semibold hover:bg-white/5 hover:text-white transition font-medium text-left"
              >
                <div className="flex items-center gap-3">
                  <FolderPlus className="w-4 h-4 text-[#C6A15B]" />
                  <span>បញ្ចូលប្រភេទ</span>
                </div>
                <ChevronDown
                  className={`w-4 h-4 transition-transform duration-200 ${
                    catMenuOpen ? 'rotate-180 text-[#C6A15B]' : 'text-slate-300'
                  }`}
                />
              </button>

              {catMenuOpen && (
                <div className="pl-6 pt-1 pb-2 space-y-1 bg-black/10 rounded-lg my-1">
                  <button
                    onClick={() => handleNav('categoryManager', 'ranks')}
                    className={`w-full text-left flex items-center gap-2.5 px-3 py-2 rounded-md text-xs transition ${
                      activePage === 'categoryManager' && activeCatType === 'ranks'
                        ? 'bg-[#C6A15B]/30 text-[#FDE047] font-semibold'
                        : 'text-slate-100 hover:text-white hover:bg-white/5'
                    }`}
                  >
                    <Award className="w-3.5 h-3.5 text-[#C6A15B]" />
                    <span>ឋានន្តរស័ក្កិ</span>
                  </button>

                  <button
                    onClick={() => handleNav('categoryManager', 'positions')}
                    className={`w-full text-left flex items-center gap-2.5 px-3 py-2 rounded-md text-xs transition ${
                      activePage === 'categoryManager' && activeCatType === 'positions'
                        ? 'bg-[#C6A15B]/30 text-[#FDE047] font-semibold'
                        : 'text-slate-100 hover:text-white hover:bg-white/5'
                    }`}
                  >
                    <Briefcase className="w-3.5 h-3.5 text-[#C6A15B]" />
                    <span>តួនាទី</span>
                  </button>

                  <button
                    onClick={() => handleNav('categoryManager', 'teamPositions')}
                    className={`w-full text-left flex items-center gap-2.5 px-3 py-2 rounded-md text-xs transition ${
                      activePage === 'categoryManager' && activeCatType === 'teamPositions'
                        ? 'bg-[#C6A15B]/30 text-[#FDE047] font-semibold'
                        : 'text-slate-100 hover:text-white hover:bg-white/5'
                    }`}
                  >
                    <Briefcase className="w-3.5 h-3.5 text-emerald-400" />
                    <span>តួនាទីតាមក្រុម</span>
                  </button>

                  <button
                    onClick={() => handleNav('categoryManager', 'visaTeamsData')}
                    className={`w-full text-left flex items-center gap-2.5 px-3 py-2 rounded-md text-xs transition ${
                      activePage === 'categoryManager' && (activeCatType === 'visaTeamsData' || activeCatType === 'visaTeams' || activeCatType === 'visaTeamsRobok')
                        ? 'bg-[#C6A15B]/30 text-[#FDE047] font-semibold'
                        : 'text-slate-100 hover:text-white hover:bg-white/5'
                    }`}
                  >
                    <Table className="w-3.5 h-3.5 text-[#C6A15B]" />
                    <span>ទិន្នន័យក្រុមផ្តល់ទិដ្ឋាការ</span>
                  </button>

                  <button
                    onClick={() => handleNav('categoryManager', 'collectorRoles')}
                    className={`w-full text-left flex items-center gap-2.5 px-3 py-2 rounded-md text-xs transition ${
                      activePage === 'categoryManager' && activeCatType === 'collectorRoles'
                        ? 'bg-[#C6A15B]/30 text-[#FDE047] font-semibold'
                        : 'text-slate-100 hover:text-white hover:bg-white/5'
                    }`}
                  >
                    <UserCheck className="w-3.5 h-3.5 text-[#C6A15B]" />
                    <span>តួនាទីអ្នកមកបើក</span>
                  </button>

                  <button
                    onClick={() => handleNav('categoryManager', 'organizations')}
                    className={`w-full text-left flex items-center gap-2.5 px-3 py-2 rounded-md text-xs transition ${
                      activePage === 'categoryManager' && activeCatType === 'organizations'
                        ? 'bg-[#C6A15B]/30 text-[#FDE047] font-semibold'
                        : 'text-slate-100 hover:text-white hover:bg-white/5'
                    }`}
                  >
                    <Building2 className="w-3.5 h-3.5 text-[#C6A15B]" />
                    <span>ស្ថាប័នដើម</span>
                  </button>

                  <button
                    onClick={() => handleNav('categoryManager', 'refusalReasons')}
                    className={`w-full text-left flex items-center gap-2.5 px-3 py-2 rounded-md text-xs transition ${
                      activePage === 'categoryManager' && activeCatType === 'refusalReasons'
                        ? 'bg-[#C6A15B]/30 text-[#FDE047] font-semibold'
                        : 'text-slate-100 hover:text-white hover:bg-white/5'
                    }`}
                  >
                    <ShieldAlert className="w-3.5 h-3.5 text-[#C6A15B]" />
                    <span>ករណីបដិសេធ</span>
                  </button>
                </div>
              )}
            </div>
          )}
        </nav>

        {/* Footer User Info */}
        <div className={`border-t border-white/10 text-xs bg-[#172554]/50 shrink-0 select-none ${isSidebarCollapsed ? 'p-2 text-center' : 'px-5 py-4'}`}>
          {!isSidebarCollapsed ? (
            <>
              <p className="text-[#FDE047] font-semibold truncate">{userName}</p>
              <p className="text-slate-300 mt-0.5 font-medium">
                តួនាទី៖ {currentRole === 'Secondary' ? 'Secondary (ការិយាល័យ)' : currentRole === 'User (ការិយាល័យ)' ? 'User (ការិយាល័យ)' : 'User (ក្រុម)'}
              </p>
              <button
                onClick={onLogout}
                className="mt-3 w-full text-left text-rose-300/80 hover:text-rose-200 transition text-xs flex items-center gap-1.5 font-medium cursor-pointer"
              >
                <span>→ ចាកចេញពីប្រព័ន្ធ</span>
              </button>
            </>
          ) : (
            <button
              onClick={onLogout}
              className="w-full flex justify-center py-2 text-rose-400 hover:text-rose-200 transition"
              title="ចាកចេញពីប្រព័ន្ធ"
            >
              <span>→</span>
            </button>
          )}
        </div>
      </aside>
    </>
  );
};
