export type UserRole = 'Secondary' | 'User' | 'Admin' | 'Supervisor' | 'Stock Manager' | string;

export type CategoryType = 'ranks' | 'positions' | 'teamPositions' | 'visaTeams' | 'visaTeamsRobok' | 'visaTeamsData' | 'collectorRoles' | 'organizations' | 'refusalReasons';

export interface CategoryItem {
  id: string;
  name: string;
  createdAt?: string;
}

export interface Officer {
  id: string;
  name?: string;
  rankId?: string;
  gender?: 'ប' | 'ស' | 'Male' | 'Female' | string;
  officerNumber?: string;
  badgeNumber?: string;
  dob?: string;
  positionId?: string;
  teamPositionId?: string;
  officeWork?: string;
  visaTeamId?: string;
  phone?: string;
  createdAt?: string;
  createdBy?: string;
  nameEn?: string;
  nameKh?: string;
  rank?: string;
  department?: string;
  station?: string;
  email?: string;
  status?: string;
  totalVisasIssued?: number;
}

export interface UserPermissions {
  refusal?: boolean;
  userMenu?: boolean;
  officerMenu?: boolean;
  visaWork?: boolean;
  officeStock?: boolean;
  categories?: boolean;
  dailyOps?: boolean;
  searchCode?: boolean;
  userList?: boolean;
}

export interface UserAccount {
  id: string;
  username: string;
  officerNumber?: string;
  email?: string;
  password?: string;
  role: UserRole;
  assignedTeam?: string; // e.g. "អាកាស តេជោ", "ព្រំដែន ប៉ោយប៉ែត", etc.
  isVisible?: boolean; // toggle to show or hide this specific user account
  hidden?: boolean;
  permissions?: UserPermissions;
  createdAt?: string;
}

export interface CategoriesState {
  ranks: CategoryItem[];
  positions: CategoryItem[];
  teamPositions?: CategoryItem[];
  visaTeams: CategoryItem[];
  visaTeamsRobok: CategoryItem[];
  collectorRoles: CategoryItem[];
  organizations: CategoryItem[];
  refusalReasons: CategoryItem[];
}

export interface VisaRecord {
  id: string;
  passportNumber: string;
  fullName: string;
  nationality: string;
  gender: 'ប្រុស' | 'ស្រី' | 'M' | 'F' | string;
  oldVisaType?: string;
  newVisaType?: string;
  documentNumber?: string;
  issueDate?: string;
  organization?: string;
  extension?: string;
  duration?: string;
  fee: number;
  applicationDate: string;
  officerId?: string;
  officerName?: string;
  visaTeamId?: string;
  status?: 'អនុម័តរួច' | 'រង់ចាំពិនិត្យ' | 'បដិសេធ' | string;
  notes?: string;
  isImported?: boolean;
  source?: 'form' | 'import' | string;
  createdAt?: string;
}

export interface StockRecord {
  id: string;
  stockType: 'sticker' | 'evisa' | string;
  operationType: 'openK1' | 'issueTeam' | 'useTeam' | 'testPrintK2' | 'damaged' | 'returnTeam' | 'transferTeam' | 'damagedTeam' | 'missingTeam' | string;
  date: string;
  time?: string;
  sourceFrom?: string;
  quantityBundles: number;
  visaTeamRobokId?: string;
  visaTeamRobokName?: string;
  recipientTeamId?: string;
  recipientTeamName?: string;
  officeApproved?: boolean;
  deptApprovalDate?: string;
  visaType?: string;
  requestedRankId?: string;
  requestedRankName?: string;
  requesterName?: string;
  collectorName?: string;
  collectorRoleId?: string;
  collectorRoleName?: string;
  startSerial?: string;
  endSerial?: string;
  totalSheets?: number;
  isImported?: boolean;
  source?: 'form' | 'import' | string;
  remarks?: string;
  note?: string;
  createdAt?: string;
  createdBy?: string;
}

// Legacy/Compatibility types
export interface Visa {
  id: string;
  visaNumber?: string;
  passportNumber: string;
  fullName: string;
  fullNameKhmer?: string;
  gender: string;
  nationality: string;
  visaType?: string;
  issueDate?: string;
  expiryDate?: string;
  officerId?: string;
  officerName?: string;
  entryPoint?: string;
  fee: number;
  status?: string;
  createdAt?: string;
  remarks?: string;
}

export interface StockItem {
  id: string;
  code: string;
  name: string;
  nameKh?: string;
  category: string;
  startSerial?: string;
  endSerial?: string;
  totalQuantity: number;
  usedQuantity: number;
  remainingQuantity: number;
  minAlertThreshold?: number;
  unitPrice?: number;
  station?: string;
  lastUpdated?: string;
  status?: string;
}

export interface StockTransaction {
  id: string;
  stockItemId: string;
  type: 'IN' | 'OUT';
  quantity: number;
  date: string;
  officerId?: string;
  officerName?: string;
  referenceNo?: string;
  note?: string;
}

export interface Category {
  id: string;
  code?: string;
  nameEn?: string;
  nameKh?: string;
  fee?: number;
  validityDays?: number;
  entriesAllowed?: string;
  description?: string;
  status?: string;
}

export interface DailyTeamRecord {
  id: string;
  date: string;
  teamName: string;
  teamId?: string;
  stickerUsage?: Record<string, number>;
  evisaUsage?: number;
  notes?: string;
  createdAt?: string;
}

export interface User {
  id: string;
  username: string;
  fullName?: string;
  role: string;
  email?: string;
  department?: string;
  status?: string;
  lastLogin?: string;
}

export interface RefusalDeportationPerson {
  id: string;
  caseNumber?: string;
  fullName: string;
  gender: 'ប្រុស' | 'ស្រី' | string;
  dob: string;
  nationality: string;
  passportNumber: string;
  reason: string;
}

export interface RefusalDeportationRecord {
  id: string;
  gateType?: 'airport' | 'border' | string;
  teamId?: string;
  teamName: string;
  date: string;
  people: RefusalDeportationPerson[];
  incomingFlight?: string;
  originCountry?: string;
  returnFlight?: string;
  destinationCountry?: string;
  comingFromBorderCountry?: string;
  returningToBorderCountry?: string;
  allowedByOffice?: boolean;
  createdAt: string;
  createdBy: string;
}



