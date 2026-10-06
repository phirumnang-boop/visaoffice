import { Officer, UserAccount, UserRole } from '../types';
import { normalizeTeamName } from './teamNormalization';

export const STORAGE_KEY_HIDE_OFFICE_OFFICERS = 'app_hide_office_officers_from_teams';
export const API_KEY_HIDE_OFFICE_OFFICERS = 'hide_office_officers_from_teams';
export const DEFAULT_HIDE_OFFICE_OFFICERS_FROM_TEAMS = true;

export const STORAGE_KEY_HIDE_OFFICER_MENU = 'app_hide_officer_menu_from_teams';
export const API_KEY_HIDE_OFFICER_MENU = 'hide_officer_menu_from_teams';
export const DEFAULT_HIDE_OFFICER_MENU_FROM_TEAMS = false;

export const STORAGE_KEY_HIDE_USER_MENU = 'app_hide_user_menu_from_teams';
export const API_KEY_HIDE_USER_MENU = 'hide_user_menu_from_teams';
export const DEFAULT_HIDE_USER_MENU_FROM_TEAMS = true;

export const STORAGE_KEY_HIDE_USER_LIST = 'app_hide_user_list_from_teams';
export const API_KEY_HIDE_USER_LIST = 'hide_user_list_from_teams';
export const DEFAULT_HIDE_USER_LIST_FROM_TEAMS = true;

export const STORAGE_KEY_HIDE_DAILY_OPS = 'app_hide_daily_ops_from_teams';
export const API_KEY_HIDE_DAILY_OPS = 'hide_daily_ops_from_teams';
export const DEFAULT_HIDE_DAILY_OPS_FROM_TEAMS = false;

export const STORAGE_KEY_HIDE_REFUSAL = 'app_hide_refusal_from_teams';
export const API_KEY_HIDE_REFUSAL = 'hide_refusal_from_teams';
export const DEFAULT_HIDE_REFUSAL_FROM_TEAMS = false;

export const STORAGE_KEY_HIDE_VISA_WORK = 'app_hide_visa_work_from_teams';
export const API_KEY_HIDE_VISA_WORK = 'hide_visa_work_from_teams';
export const DEFAULT_HIDE_VISA_WORK_FROM_TEAMS = true;

export const STORAGE_KEY_HIDE_OFFICE_STOCK = 'app_hide_office_stock_from_teams';
export const API_KEY_HIDE_OFFICE_STOCK = 'hide_office_stock_from_teams';
export const DEFAULT_HIDE_OFFICE_STOCK_FROM_TEAMS = true;

export const STORAGE_KEY_HIDE_SEARCH_CODE = 'app_hide_search_code_from_teams';
export const API_KEY_HIDE_SEARCH_CODE = 'hide_search_code_from_teams';
export const DEFAULT_HIDE_SEARCH_CODE_FROM_TEAMS = false;

export const STORAGE_KEY_HIDE_CATEGORIES = 'app_hide_categories_from_teams';
export const API_KEY_HIDE_CATEGORIES = 'hide_categories_from_teams';
export const DEFAULT_HIDE_CATEGORIES_FROM_TEAMS = true;

/**
 * Checks if an officer works at the office (បំរើការងារនៅការិយាល័យ)
 * Officers assigned to office work (ក១, ក២, ក៣, ក៤, ក៥) or office departments/stations
 */
export const isOfficeOfficer = (officer?: Officer | null): boolean => {
  if (!officer) return false;

  // 1. Check officeWork field (e.g. 'ក១', 'ក២', 'ក៣', 'ក៤', 'ក៥' or any office assignment)
  const officeWork = (officer.officeWork || '').trim();
  if (officeWork && officeWork !== '-' && officeWork !== '') {
    return true;
  }

  // 2. Check department (e.g. 'ការិយាល័យទិដ្ឋាការចូល', 'ផ្នែករដ្ឋបាល', 'ការិយាល័យ...')
  const department = (officer.department || '').trim();
  if (/ការិយាល័យ|HQ|Headquarters|រដ្ឋបាលកណ្តាល/i.test(department)) {
    return true;
  }

  // 3. Check station (e.g. 'Phnom Penh HQ', 'ការិយាល័យ')
  const station = (officer.station || '').trim();
  if (/HQ|Headquarters|ការិយាល័យ/i.test(station)) {
    return true;
  }

  return false;
};

/**
 * Checks if the user is a team user ("របស់ក្រុម")
 * Office admins are 'Secondary' or 'Admin' without specific team lockdown.
 * Team users are 'User', 'Supervisor', or any user assigned to a team.
 */
export const isTeamUser = (role?: UserRole | string, assignedTeam?: string): boolean => {
  if (role === 'Secondary' || role === 'Admin' || role === 'User (ការិយាល័យ)') {
    return false;
  }
  return true;
};

/**
 * Checks if an officer is assigned to a specific team (visaTeamId matches assignedTeam)
 * Supports matching short names (e.g. "អាកាស តេជោ"), full names (e.g. "ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអាកាសយានដ្ឋានអន្តរជាតិតេជោ"), and IDs.
 */
export const isOfficerInTeam = (
  officer?: Officer | null,
  assignedTeam?: string,
  visaTeams?: Array<{ id: string; name: string }>
): boolean => {
  if (!assignedTeam || !assignedTeam.trim()) return true;
  if (!officer || !officer.visaTeamId || !officer.visaTeamId.trim()) return false;

  const ofRaw = officer.visaTeamId.trim();
  const userRaw = assignedTeam.trim();

  // 1. Direct case-insensitive match
  if (ofRaw.toLowerCase() === userRaw.toLowerCase()) return true;

  // 2. Normalized team name match (e.g. "អាកាស តេជោ" vs "ក្រុមផ្តល់ទិដ្ឋាការប្រចាំច្រកទ្វារអាកាសយានដ្ឋានអន្តរជាតិតេជោ")
  const ofNorm = normalizeTeamName(ofRaw);
  const userNorm = normalizeTeamName(userRaw);
  if (ofNorm && userNorm && ofNorm === userNorm) return true;

  // 3. Lookup via category list
  if (visaTeams && visaTeams.length > 0) {
    const ofItem = visaTeams.find(
      (v) =>
        v.id === ofRaw ||
        v.name === ofRaw ||
        (v.name && normalizeTeamName(v.name) === ofNorm) ||
        (v.id && normalizeTeamName(v.id) === ofNorm)
    );
    const userItem = visaTeams.find(
      (v) =>
        v.id === userRaw ||
        v.name === userRaw ||
        (v.name && normalizeTeamName(v.name) === userNorm) ||
        (v.id && normalizeTeamName(v.id) === userNorm)
    );

    if (ofItem && userItem && ofItem.id === userItem.id) return true;
    if (userItem && (userItem.id === ofRaw || userItem.name === ofRaw || normalizeTeamName(userItem.name) === ofNorm)) return true;
    if (ofItem && (ofItem.id === userRaw || ofItem.name === userRaw || normalizeTeamName(ofItem.name) === userNorm)) return true;
  }

  // 4. Substring inclusion
  if (ofNorm && userNorm && (ofNorm.includes(userNorm) || userNorm.includes(ofNorm))) return true;

  return false;
};

/**
 * Filter officers based on access restriction rules:
 * When hideOfficeOfficers is active and viewer is a team user:
 * - If officer has a visaTeamId matching the user's assignedTeam, ALWAYS include them!
 * - Omit office-only officers without a matching visaTeam.
 */
export const filterOfficersForViewer = (
  officers: Officer[],
  options: {
    isTeam: boolean;
    hideOfficeOfficers?: boolean;
    assignedTeam?: string;
    visaTeams?: Array<{ id: string; name: string }>;
  }
): Officer[] => {
  const { isTeam, hideOfficeOfficers = true, assignedTeam, visaTeams } = options;

  if (!isTeam) {
    return officers;
  }

  return officers.filter((of) => {
    // If officer is explicitly assigned to this team via visaTeamId, ALWAYS include them!
    if (assignedTeam && of.visaTeamId) {
      if (isOfficerInTeam(of, assignedTeam, visaTeams)) {
        return true;
      }
    }

    // Strict block: Team users cannot see office-only officers
    if (hideOfficeOfficers && isOfficeOfficer(of) && !of.visaTeamId) {
      return false;
    }

    return true;
  });
};
