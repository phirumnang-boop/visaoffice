import React from 'react';
import { CategoriesState, UserAccount, UserRole } from '../types';
import { useWorkspaceSettings, WorkspaceScale } from '../context/WorkspaceSettingsContext';
import {
  ShieldCheck,
  User,
  LogOut,
  CheckCircle,
  AlertCircle,
  Cloud,
  RefreshCw,
  Building2,
  ChevronDown,
  Maximize2,
  Sliders,
  Monitor,
} from 'lucide-react';

interface HeaderProps {
  currentUser?: UserAccount | null;
  pageTitle?: string;
  categories?: CategoriesState;
  toastMessage?: { text: string; type: 'success' | 'error' } | null;
  isCloudConnected?: boolean;
  isCloudSyncing?: boolean;
  onSwitchUserRole?: (role: UserRole) => void;
  onSwitchTeam?: (team: string) => void;
  onToggleSidebar?: () => void;
  onOpenAuthModal?: () => void;
  onLogout?: () => void;
  showRoleSwitcher?: boolean;
}

export const Header: React.FC<HeaderProps> = ({
  currentUser = { id: 'usr-01', username: 'admin.immigration', role: 'Secondary' },
  pageTitle = 'ទំព័រដើម (Dashboard)',
  categories,
  toastMessage,
  isCloudConnected = true,
  isCloudSyncing = false,
  onSwitchUserRole,
  onSwitchTeam,
  onLogout,
  showRoleSwitcher = false,
}) => {
  const effectiveUser: UserAccount = currentUser || { id: 'usr-01', username: 'admin.immigration', role: 'Secondary' };
  const { scaleMode, setScaleMode, containerWidthMode, setContainerWidthMode } = useWorkspaceSettings();
  return (
    <header className="print:hidden bg-white border-b border-[#C6A15B]/30 px-4 md:px-8 py-3 md:py-4 flex flex-col md:flex-row md:items-center justify-between gap-3 shadow-xs sticky top-0 z-20">
      <div className="flex items-center justify-between w-full md:w-auto">
        <div>
          <h2 className="font-bold text-[#0B2545] text-sm md:text-base leading-snug">
            {pageTitle}
          </h2>
          <p className="text-[11px] text-gray-500 font-normal">
            ប្រព័ន្ធគ្រប់គ្រងទិន្នន័យ និងព័ត៌មានមន្ត្រី
          </p>
        </div>

        {/* Toast indicator on mobile */}
        {toastMessage && (
          <div className="md:hidden">
            <span
              className={`inline-flex items-center gap-1.5 px-2.5 py-1 rounded-full text-xs font-medium animate-fade ${
                toastMessage.type === 'success'
                  ? 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                  : 'bg-rose-100 text-rose-800 border border-rose-300'
              }`}
            >
              {toastMessage.type === 'success' ? (
                <CheckCircle className="w-3.5 h-3.5" />
              ) : (
                <AlertCircle className="w-3.5 h-3.5" />
              )}
              {toastMessage.text}
            </span>
          </div>
        )}
      </div>

      <div className="flex items-center gap-2.5 self-end md:self-center flex-wrap">
        {/* Cloud SQL Database Sync Indicator */}
        <div
          title={
            isCloudSyncing
              ? 'កំពុងធ្វើសមកាលកម្មទិន្នន័យ...'
              : isCloudConnected
              ? 'បានភ្ជាប់ទិន្នន័យប្រព័ន្ធ'
              : 'ដំណើរការក្នុងទម្រង់ Offline'
          }
          className={`flex items-center gap-1.5 px-2.5 py-1 rounded-full text-[11px] font-semibold border transition ${
            isCloudSyncing
              ? 'bg-blue-50 text-blue-700 border-blue-200'
              : isCloudConnected
              ? 'bg-emerald-50 text-emerald-700 border-emerald-200'
              : 'bg-amber-50 text-amber-700 border-amber-200'
          }`}
        >
          {isCloudSyncing ? (
            <RefreshCw className="w-3.5 h-3.5 animate-spin text-blue-600" />
          ) : (
            <Cloud className="w-3.5 h-3.5 text-emerald-600" />
          )}
          <span className="hidden sm:inline">
            {isCloudSyncing
              ? 'Syncing Cloud...'
              : isCloudConnected
              ? 'System Online'
              : 'Local Mode'}
          </span>
        </div>

        {/* Toast indicator on desktop */}
        {toastMessage && (
          <div className="hidden md:block">
            <span
              className={`inline-flex items-center gap-1.5 px-3 py-1.5 rounded-full text-xs font-medium animate-fade ${
                toastMessage.type === 'success'
                  ? 'bg-emerald-100 text-emerald-800 border border-emerald-300'
                  : 'bg-rose-100 text-rose-800 border border-rose-300'
              }`}
            >
              {toastMessage.type === 'success' ? (
                <CheckCircle className="w-3.5 h-3.5 text-emerald-600" />
              ) : (
                <AlertCircle className="w-3.5 h-3.5 text-rose-600" />
              )}
              {toastMessage.text}
            </span>
          </div>
        )}

        {/* User Badge & Quick Role Switcher */}
        <div className="flex items-center gap-2 bg-[#0B2545]/5 px-3 py-1.5 rounded-lg border border-[#C6A15B]/30 text-xs">
          <div className="w-7 h-7 rounded-full bg-[#0B2545] text-[#E4CD98] flex items-center justify-center font-bold font-display text-[11px]">
            {(effectiveUser.username || 'A').substring(0, 1).toUpperCase()}
          </div>
          <div className="text-left">
            <p className="font-semibold text-[#0B2545] truncate max-w-[120px] sm:max-w-[160px]">
              {effectiveUser.username}
            </p>
            <div className="flex items-center gap-1 flex-wrap">
              <span
                className={`inline-block w-1.5 h-1.5 rounded-full ${
                  effectiveUser.role === 'Secondary' || effectiveUser.role === 'User (ការិយាល័យ)' ? 'bg-amber-500' : 'bg-blue-500'
                }`}
              />
              <span className="text-[10px] text-gray-600 font-medium">
                {effectiveUser.role === 'Secondary'
                  ? 'ការិយាល័យ (គ្រប់គ្រង)'
                  : effectiveUser.role === 'User (ការិយាល័យ)'
                  ? 'User (ការិយាល័យ)'
                  : effectiveUser.assignedTeam
                  ? `ក្រុម (${effectiveUser.assignedTeam})`
                  : 'User (មន្ត្រី)'}
              </span>
            </div>
          </div>

          {/* Role Toggle Switch for Testing */}
          {showRoleSwitcher && onSwitchUserRole && (
            <div className="ml-2 pl-2 border-l border-gray-300 flex items-center gap-1 flex-wrap">
              <button
                title="ប្តូរទៅកាន់ការិយាល័យ Secondary (គ្រប់គ្រងពេញលេញ)"
                onClick={() => onSwitchUserRole('Secondary')}
                className={`px-1.5 py-0.5 rounded text-[10px] transition cursor-pointer ${
                  effectiveUser.role === 'Secondary'
                    ? 'bg-[#0B2545] text-[#E4CD98] font-bold shadow-xs'
                    : 'bg-gray-200 text-gray-700 hover:bg-gray-300'
                }`}
              >
                ការិយាល័យ (Sec)
              </button>
              <button
                title="ប្តូរទៅកាន់ក្រុម User (បញ្ចូលទិន្នន័យតាមក្រុម)"
                onClick={() => onSwitchUserRole('User')}
                className={`px-1.5 py-0.5 rounded text-[10px] transition cursor-pointer ${
                  effectiveUser.role === 'User'
                    ? 'bg-blue-700 text-white font-bold shadow-xs'
                    : 'bg-gray-200 text-gray-700 hover:bg-gray-300'
                }`}
              >
                ក្រុម (User)
              </button>
            </div>
          )}
        </div>

        {onLogout && (
          <button
            onClick={onLogout}
            title="ចាកចេញ"
            className="p-2 text-gray-600 hover:text-rose-600 hover:bg-rose-50 rounded-lg transition"
          >
            <LogOut className="w-4 h-4" />
          </button>
        )}
      </div>
    </header>
  );
};
