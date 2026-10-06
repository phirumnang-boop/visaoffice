import fs from 'fs';
import path from 'path';
import { db } from './index';
import {
  users,
  categories,
  officers,
  visaRecords,
  stockRecords,
  dailyTeamOperations,
  type User,
  type InsertUser,
  type Category,
  type InsertCategory,
  type DbOfficer,
  type InsertOfficer,
  type DbVisaRecord,
  type InsertVisaRecord,
  type DbStockRecord,
  type InsertStockRecord,
  type DbDailyTeamOperation,
  type InsertDailyTeamOperation,
} from './schema';
import { eq, desc, and, inArray } from 'drizzle-orm';

// Persistent Local File Storage for durability when Cloud SQL is unconfigured or offline
interface FileStorageData {
  users: User[];
  categories: Category[];
  officers: DbOfficer[];
  visaRecords: DbVisaRecord[];
  stockRecords: DbStockRecord[];
  stickerActualStock: DbStockRecord[];
  dailyTeamOperations: DbDailyTeamOperation[];
  settings: Record<string, any>;
}

let DATA_DIR = path.join(process.cwd(), 'data');
// Robust fallback if running inside 'src' subdirectory
if (!fs.existsSync(DATA_DIR) && process.cwd().endsWith('src')) {
  const parentData = path.resolve(process.cwd(), '..', 'data');
  DATA_DIR = parentData;
}
const DATA_FILE = path.join(DATA_DIR, 'app_storage.json');

class FilePersistence {
  private cache: FileStorageData;
  private saveTimeout: NodeJS.Timeout | null = null;

  constructor() {
    this.cache = this.loadFromFile();
  }

  private loadFromFile(): FileStorageData {
    try {
      if (!fs.existsSync(DATA_DIR)) {
        fs.mkdirSync(DATA_DIR, { recursive: true });
      }
      if (fs.existsSync(DATA_FILE)) {
        const content = fs.readFileSync(DATA_FILE, 'utf-8');
        const parsed = JSON.parse(content);
        return {
          users: Array.isArray(parsed.users) ? parsed.users : [],
          categories: Array.isArray(parsed.categories) ? parsed.categories : [],
          officers: Array.isArray(parsed.officers) ? parsed.officers : [],
          visaRecords: Array.isArray(parsed.visaRecords) ? parsed.visaRecords : [],
          stockRecords: Array.isArray(parsed.stockRecords) ? parsed.stockRecords : [],
          stickerActualStock: Array.isArray(parsed.stickerActualStock) ? parsed.stickerActualStock : [],
          dailyTeamOperations: Array.isArray(parsed.dailyTeamOperations) ? parsed.dailyTeamOperations : [],
          settings: parsed.settings && typeof parsed.settings === 'object' ? parsed.settings : {},
        };
      }
    } catch (e) {
      console.warn('Notice loading file persistence:', e);
    }
    return {
      users: [],
      categories: [],
      officers: [],
      visaRecords: [],
      stockRecords: [],
      stickerActualStock: [],
      dailyTeamOperations: [],
      settings: {},
    };
  }

  public flush() {
    try {
      if (!fs.existsSync(DATA_DIR)) {
        fs.mkdirSync(DATA_DIR, { recursive: true });
      }
      fs.writeFileSync(DATA_FILE, JSON.stringify(this.cache, null, 2), 'utf-8');
    } catch (e) {
      console.warn('Notice saving file persistence:', e);
    }
  }

  public scheduleFlush() {
    if (this.saveTimeout) clearTimeout(this.saveTimeout);
    this.saveTimeout = setTimeout(() => {
      this.flush();
    }, 200);
  }

  public getData(): FileStorageData {
    return this.cache;
  }
}

const fileStore = new FilePersistence();

export interface IStorage {
  // Users
  getUser(id: string): Promise<User | undefined>;
  getUserByUsername(username: string): Promise<User | undefined>;
  getUsers(): Promise<User[]>;
  createUser(user: InsertUser): Promise<User>;
  updateUser(id: string, user: Partial<InsertUser>): Promise<User | undefined>;
  deleteUser(id: string): Promise<boolean>;

  // Categories
  getCategories(type?: string): Promise<Category[]>;
  createCategory(category: InsertCategory): Promise<Category>;
  updateCategory(id: string, name: string): Promise<Category | undefined>;
  deleteCategory(id: string): Promise<boolean>;

  // Officers
  getOfficers(): Promise<DbOfficer[]>;
  getOfficer(id: string): Promise<DbOfficer | undefined>;
  createOfficer(officer: InsertOfficer): Promise<DbOfficer>;
  updateOfficer(id: string, officer: Partial<InsertOfficer>): Promise<DbOfficer | undefined>;
  deleteOfficer(id: string): Promise<boolean>;

  // Visa Records
  getVisaRecords(): Promise<DbVisaRecord[]>;
  saveVisaRecord(record: InsertVisaRecord): Promise<DbVisaRecord>;
  bulkSaveVisaRecords(records: InsertVisaRecord[]): Promise<void>;
  createVisaRecord(record: InsertVisaRecord): Promise<DbVisaRecord>;
  updateVisaRecord(id: string, record: Partial<InsertVisaRecord>): Promise<DbVisaRecord | undefined>;
  deleteVisaRecord(id: string): Promise<boolean>;
  bulkDeleteVisaRecords(ids: string[]): Promise<boolean>;

  // Stock Records (General Stock - both Sticker & EVisa)
  getStockRecords(): Promise<DbStockRecord[]>;
  saveStockRecord(record: InsertStockRecord): Promise<DbStockRecord>;
  bulkSaveStockRecords(records: InsertStockRecord[]): Promise<void>;
  createStockRecord(record: InsertStockRecord): Promise<DbStockRecord>;
  updateStockRecord(id: string, record: Partial<InsertStockRecord>): Promise<DbStockRecord | undefined>;
  deleteStockRecord(id: string): Promise<boolean>;
  bulkDeleteStockRecords(ids: string[], stockType?: string, operationType?: string): Promise<boolean>;

  // Dedicated Sticker Actual Stock (ស្តុកជាក់ស្តែងសម្រាប់បើកផ្តល់តាមក្រុម)
  getStickerActualStock(): Promise<DbStockRecord[]>;
  saveStickerActualStock(record: InsertStockRecord): Promise<DbStockRecord>;
  bulkSaveStickerActualStock(records: InsertStockRecord[], overwrite?: boolean): Promise<void>;
  deleteStickerActualStock(id: string): Promise<boolean>;
  bulkDeleteStickerActualStock(ids: string[]): Promise<boolean>;
  clearAllStickerActualStock(): Promise<boolean>;

  // Daily Operations
  getDailyTeamOperations(): Promise<DbDailyTeamOperation[]>;
  saveDailyTeamOperation(op: InsertDailyTeamOperation): Promise<DbDailyTeamOperation>;
  bulkSaveDailyTeamOperations(records: InsertDailyTeamOperation[]): Promise<void>;
  deleteDailyTeamOperation(id: string): Promise<boolean>;

  // Settings
  getSetting(key: string): Promise<any>;
  saveSetting(key: string, value: any): Promise<void>;
}

export class DatabaseStorage implements IStorage {
  // User operations
  async getUser(id: string): Promise<User | undefined> {
    if (db) {
      try {
        const [user] = await db.select().from(users).where(eq(users.id, id));
        if (user) return user;
      } catch {}
    }
    return fileStore.getData().users.find((u) => u.id === id);
  }

  async getUserByUsername(username: string): Promise<User | undefined> {
    if (db) {
      try {
        const [user] = await db.select().from(users).where(eq(users.username, username));
        if (user) return user;
      } catch {}
    }
    return fileStore.getData().users.find((u) => u.username === username);
  }

  async getUsers(): Promise<User[]> {
    if (db) {
      try {
        const dbUsers = await db.select().from(users).orderBy(desc(users.createdAt));
        if (dbUsers.length > 0) return dbUsers;
      } catch {}
    }
    return fileStore.getData().users;
  }

  async createUser(insertUser: InsertUser): Promise<User> {
    if (db) {
      try {
        const [user] = await db.insert(users).values(insertUser).returning();
        if (user) return user;
      } catch {}
    }
    const newUser = { ...insertUser, id: insertUser.id || `usr-${Date.now()}` } as User;
    fileStore.getData().users.push(newUser);
    fileStore.scheduleFlush();
    return newUser;
  }

  async updateUser(id: string, updated: Partial<InsertUser>): Promise<User | undefined> {
    if (db) {
      try {
        const [user] = await db.update(users).set(updated).where(eq(users.id, id)).returning();
        if (user) return user;
      } catch {}
    }
    const idx = fileStore.getData().users.findIndex((u) => u.id === id);
    if (idx !== -1) {
      fileStore.getData().users[idx] = { ...fileStore.getData().users[idx], ...updated } as User;
      fileStore.scheduleFlush();
      return fileStore.getData().users[idx];
    }
    return undefined;
  }

  async deleteUser(id: string): Promise<boolean> {
    if (db) {
      try {
        const result = await db.delete(users).where(eq(users.id, id));
        return (result.rowCount ?? 0) > 0;
      } catch {}
    }
    const prevLen = fileStore.getData().users.length;
    fileStore.getData().users = fileStore.getData().users.filter((u) => u.id !== id);
    fileStore.scheduleFlush();
    return fileStore.getData().users.length < prevLen;
  }

  // Category operations
  async getCategories(type?: string): Promise<Category[]> {
    if (db) {
      try {
        if (type) {
          return await db.select().from(categories).where(eq(categories.type, type));
        }
        return await db.select().from(categories);
      } catch {}
    }
    const all = fileStore.getData().categories;
    return type ? all.filter((c) => c.type === type) : all;
  }

  async createCategory(category: InsertCategory): Promise<Category> {
    if (db) {
      try {
        const [item] = await db.insert(categories).values(category).returning();
        if (item) return item;
      } catch {}
    }
    const newCat = { ...category, id: category.id || `cat-${Date.now()}` } as Category;
    fileStore.getData().categories.push(newCat);
    fileStore.scheduleFlush();
    return newCat;
  }

  async updateCategory(id: string, name: string): Promise<Category | undefined> {
    if (db) {
      try {
        const [item] = await db.update(categories).set({ name }).where(eq(categories.id, id)).returning();
        if (item) return item;
      } catch {}
    }
    const cat = fileStore.getData().categories.find((c) => c.id === id);
    if (cat) {
      cat.name = name;
      fileStore.scheduleFlush();
      return cat;
    }
    return undefined;
  }

  async deleteCategory(id: string): Promise<boolean> {
    if (db) {
      try {
        const result = await db.delete(categories).where(eq(categories.id, id));
        return (result.rowCount ?? 0) > 0;
      } catch {}
    }
    const prevLen = fileStore.getData().categories.length;
    fileStore.getData().categories = fileStore.getData().categories.filter((c) => c.id !== id);
    fileStore.scheduleFlush();
    return fileStore.getData().categories.length < prevLen;
  }

  // Officer operations
  async getOfficers(): Promise<DbOfficer[]> {
    if (db) {
      try {
        return await db.select().from(officers);
      } catch {}
    }
    return fileStore.getData().officers;
  }

  async getOfficer(id: string): Promise<DbOfficer | undefined> {
    if (db) {
      try {
        const [officer] = await db.select().from(officers).where(eq(officers.id, id));
        if (officer) return officer;
      } catch {}
    }
    return fileStore.getData().officers.find((o) => o.id === id);
  }

  async createOfficer(officer: InsertOfficer): Promise<DbOfficer> {
    if (db) {
      try {
        const [item] = await db.insert(officers).values(officer).returning();
        if (item) return item;
      } catch {}
    }
    const newOff = { ...officer, id: officer.id || `off-${Date.now()}` } as DbOfficer;
    fileStore.getData().officers.push(newOff);
    fileStore.scheduleFlush();
    return newOff;
  }

  async updateOfficer(id: string, officer: Partial<InsertOfficer>): Promise<DbOfficer | undefined> {
    if (db) {
      try {
        const [item] = await db.update(officers).set(officer).where(eq(officers.id, id)).returning();
        if (item) return item;
      } catch {}
    }
    const off = fileStore.getData().officers.find((o) => o.id === id);
    if (off) {
      Object.assign(off, officer);
      fileStore.scheduleFlush();
      return off;
    }
    return undefined;
  }

  async deleteOfficer(id: string): Promise<boolean> {
    if (db) {
      try {
        const result = await db.delete(officers).where(eq(officers.id, id));
        return (result.rowCount ?? 0) > 0;
      } catch {}
    }
    const prevLen = fileStore.getData().officers.length;
    fileStore.getData().officers = fileStore.getData().officers.filter((o) => o.id !== id);
    fileStore.scheduleFlush();
    return fileStore.getData().officers.length < prevLen;
  }

  // Visa Records
  async getVisaRecords(): Promise<DbVisaRecord[]> {
    if (db) {
      try {
        const records = await db.select().from(visaRecords).orderBy(desc(visaRecords.createdAt));
        if (records.length > 0) return records;
      } catch {}
    }
    return fileStore.getData().visaRecords;
  }

  async saveVisaRecord(record: InsertVisaRecord): Promise<DbVisaRecord> {
    if (db) {
      try {
        const [existing] = await db.select().from(visaRecords).where(eq(visaRecords.id, record.id));
        if (existing) {
          const [updated] = await db.update(visaRecords).set(record).where(eq(visaRecords.id, record.id)).returning();
          return updated;
        } else {
          const [inserted] = await db.insert(visaRecords).values(record).returning();
          return inserted;
        }
      } catch {}
    }
    const items = fileStore.getData().visaRecords;
    const idx = items.findIndex((r) => r.id === record.id);
    const item = { ...record } as DbVisaRecord;
    if (idx !== -1) {
      items[idx] = item;
    } else {
      items.unshift(item);
    }
    fileStore.scheduleFlush();
    return item;
  }

  async bulkSaveVisaRecords(records: InsertVisaRecord[]): Promise<void> {
    if (db) {
      try {
        for (const r of records) {
          await this.saveVisaRecord(r);
        }
        return;
      } catch {}
    }
    const items = fileStore.getData().visaRecords;
    const map = new Map<string, DbVisaRecord>();
    items.forEach((r) => map.set(r.id, r));
    records.forEach((r) => map.set(r.id, r as DbVisaRecord));
    fileStore.getData().visaRecords = Array.from(map.values());
    fileStore.flush();
  }

  async createVisaRecord(record: InsertVisaRecord): Promise<DbVisaRecord> {
    return this.saveVisaRecord(record);
  }

  async updateVisaRecord(id: string, record: Partial<InsertVisaRecord>): Promise<DbVisaRecord | undefined> {
    if (db) {
      try {
        const [item] = await db.update(visaRecords).set(record).where(eq(visaRecords.id, id)).returning();
        if (item) return item;
      } catch {}
    }
    const item = fileStore.getData().visaRecords.find((r) => r.id === id);
    if (item) {
      Object.assign(item, record);
      fileStore.scheduleFlush();
      return item;
    }
    return undefined;
  }

  async deleteVisaRecord(id: string): Promise<boolean> {
    if (db) {
      try {
        const result = await db.delete(visaRecords).where(eq(visaRecords.id, id));
        return (result.rowCount ?? 0) > 0;
      } catch {}
    }
    const prev = fileStore.getData().visaRecords.length;
    fileStore.getData().visaRecords = fileStore.getData().visaRecords.filter((r) => r.id !== id);
    fileStore.scheduleFlush();
    return fileStore.getData().visaRecords.length < prev;
  }

  async bulkDeleteVisaRecords(ids: string[]): Promise<boolean> {
    if (!ids || ids.length === 0) return true;
    if (db) {
      try {
        const result = await db.delete(visaRecords).where(inArray(visaRecords.id, ids));
        return (result.rowCount ?? 0) > 0;
      } catch {}
    }
    const idSet = new Set(ids);
    fileStore.getData().visaRecords = fileStore.getData().visaRecords.filter((r) => !idSet.has(r.id));
    fileStore.flush();
    return true;
  }

  // Stock Records (General Stock - both Sticker & EVisa)
  async getStockRecords(): Promise<DbStockRecord[]> {
    if (db) {
      try {
        const records = await db.select().from(stockRecords).orderBy(desc(stockRecords.createdAt));
        if (records.length > 0) return records;
      } catch {}
    }
    return fileStore.getData().stockRecords;
  }

  async saveStockRecord(record: InsertStockRecord): Promise<DbStockRecord> {
    if (db) {
      try {
        const [existing] = await db.select().from(stockRecords).where(eq(stockRecords.id, record.id));
        if (existing) {
          const [updated] = await db.update(stockRecords).set(record).where(eq(stockRecords.id, record.id)).returning();
          return updated;
        } else {
          const [inserted] = await db.insert(stockRecords).values(record).returning();
          return inserted;
        }
      } catch {}
    }
    const items = fileStore.getData().stockRecords;
    const idx = items.findIndex((r) => r.id === record.id);
    const item = { ...record } as DbStockRecord;
    if (idx !== -1) {
      items[idx] = item;
    } else {
      items.unshift(item);
    }
    fileStore.scheduleFlush();
    return item;
  }

  async bulkSaveStockRecords(records: InsertStockRecord[]): Promise<void> {
    if (db) {
      try {
        for (const r of records) {
          await this.saveStockRecord(r);
        }
        return;
      } catch {}
    }
    const items = fileStore.getData().stockRecords;
    const map = new Map<string, DbStockRecord>();
    items.forEach((r) => map.set(r.id, r));
    records.forEach((r) => map.set(r.id, r as DbStockRecord));
    fileStore.getData().stockRecords = Array.from(map.values());
    fileStore.flush();
  }

  async createStockRecord(record: InsertStockRecord): Promise<DbStockRecord> {
    return this.saveStockRecord(record);
  }

  async updateStockRecord(id: string, record: Partial<InsertStockRecord>): Promise<DbStockRecord | undefined> {
    if (db) {
      try {
        const [item] = await db.update(stockRecords).set(record).where(eq(stockRecords.id, id)).returning();
        if (item) return item;
      } catch {}
    }
    const item = fileStore.getData().stockRecords.find((r) => r.id === id);
    if (item) {
      Object.assign(item, record);
      fileStore.scheduleFlush();
      return item;
    }
    return undefined;
  }

  async deleteStockRecord(id: string): Promise<boolean> {
    if (db) {
      try {
        const result = await db.delete(stockRecords).where(eq(stockRecords.id, id));
        return (result.rowCount ?? 0) > 0;
      } catch {}
    }
    const prev = fileStore.getData().stockRecords.length;
    fileStore.getData().stockRecords = fileStore.getData().stockRecords.filter((r) => r.id !== id);
    fileStore.scheduleFlush();
    return fileStore.getData().stockRecords.length < prev;
  }

  async bulkDeleteStockRecords(ids: string[], stockType?: string, operationType?: string): Promise<boolean> {
    if (db) {
      try {
        if (ids && ids.length > 0) {
          const result = await db.delete(stockRecords).where(inArray(stockRecords.id, ids));
          return (result.rowCount ?? 0) > 0;
        }
        if (stockType && operationType) {
          const result = await db.delete(stockRecords).where(and(eq(stockRecords.stockType, stockType), eq(stockRecords.operationType, operationType)));
          return (result.rowCount ?? 0) > 0;
        }
      } catch {}
    }
    if (ids && ids.length > 0) {
      const idSet = new Set(ids);
      fileStore.getData().stockRecords = fileStore.getData().stockRecords.filter((r) => !idSet.has(r.id));
    } else if (stockType && operationType) {
      fileStore.getData().stockRecords = fileStore.getData().stockRecords.filter(
        (r) => !(r.stockType === stockType && r.operationType === operationType)
      );
    }
    fileStore.flush();
    return true;
  }

  // Dedicated Sticker Actual Stock (ស្តុកជាក់ស្តែងសម្រាប់បើកផ្តល់តាមក្រុម)
  async getStickerActualStock(): Promise<DbStockRecord[]> {
    return fileStore.getData().stickerActualStock;
  }

  async saveStickerActualStock(record: InsertStockRecord): Promise<DbStockRecord> {
    const items = fileStore.getData().stickerActualStock;
    const idx = items.findIndex((r) => r.id === record.id);
    const item = { ...record } as DbStockRecord;
    if (idx !== -1) {
      items[idx] = item;
    } else {
      items.unshift(item);
    }
    fileStore.scheduleFlush();
    return item;
  }

  async bulkSaveStickerActualStock(records: InsertStockRecord[], overwrite: boolean = false): Promise<void> {
    if (overwrite) {
      fileStore.getData().stickerActualStock = records as DbStockRecord[];
    } else {
      const items = fileStore.getData().stickerActualStock;
      const map = new Map<string, DbStockRecord>();
      items.forEach((r) => map.set(r.id, r));
      records.forEach((r) => map.set(r.id, r as DbStockRecord));
      fileStore.getData().stickerActualStock = Array.from(map.values());
    }
    fileStore.flush();
  }

  async deleteStickerActualStock(id: string): Promise<boolean> {
    const prev = fileStore.getData().stickerActualStock.length;
    fileStore.getData().stickerActualStock = fileStore.getData().stickerActualStock.filter((r) => r.id !== id);
    fileStore.scheduleFlush();
    return fileStore.getData().stickerActualStock.length < prev;
  }

  async bulkDeleteStickerActualStock(ids: string[]): Promise<boolean> {
    if (!ids || ids.length === 0) return true;
    const idSet = new Set(ids);
    fileStore.getData().stickerActualStock = fileStore.getData().stickerActualStock.filter((r) => !idSet.has(r.id));
    fileStore.flush();
    return true;
  }

  async clearAllStickerActualStock(): Promise<boolean> {
    fileStore.getData().stickerActualStock = [];
    fileStore.flush();
    return true;
  }

  // Daily Operations
  async getDailyTeamOperations(): Promise<DbDailyTeamOperation[]> {
    if (db) {
      try {
        return await db.select().from(dailyTeamOperations).orderBy(desc(dailyTeamOperations.date));
      } catch {}
    }
    return fileStore.getData().dailyTeamOperations;
  }

  async saveDailyTeamOperation(op: InsertDailyTeamOperation): Promise<DbDailyTeamOperation> {
    if (db) {
      try {
        const [existing] = await db
          .select()
          .from(dailyTeamOperations)
          .where(and(eq(dailyTeamOperations.date, op.date), eq(dailyTeamOperations.teamName, op.teamName)));

        if (existing) {
          const [updated] = await db
            .update(dailyTeamOperations)
            .set(op)
            .where(eq(dailyTeamOperations.id, existing.id))
            .returning();
          return updated;
        } else {
          const [inserted] = await db.insert(dailyTeamOperations).values(op).returning();
          return inserted;
        }
      } catch {}
    }
    const items = fileStore.getData().dailyTeamOperations;
    const idx = items.findIndex((d) => d.date === op.date && d.teamName === op.teamName);
    const item = { ...op } as DbDailyTeamOperation;
    if (idx !== -1) {
      items[idx] = item;
    } else {
      items.unshift(item);
    }
    fileStore.scheduleFlush();
    return item;
  }

  async bulkSaveDailyTeamOperations(records: InsertDailyTeamOperation[]): Promise<void> {
    if (db) {
      try {
        for (const r of records) {
          await this.saveDailyTeamOperation(r);
        }
        return;
      } catch {}
    }
    const items = fileStore.getData().dailyTeamOperations;
    const map = new Map<string, DbDailyTeamOperation>();
    items.forEach((r) => map.set(`${r.date}_${r.teamName}`, r));
    records.forEach((r) => map.set(`${r.date}_${r.teamName}`, r as DbDailyTeamOperation));
    fileStore.getData().dailyTeamOperations = Array.from(map.values());
    fileStore.flush();
  }

  async deleteDailyTeamOperation(id: string): Promise<boolean> {
    if (db) {
      try {
        const result = await db.delete(dailyTeamOperations).where(eq(dailyTeamOperations.id, id));
        return (result.rowCount ?? 0) > 0;
      } catch (err) {
        console.error('Database delete failed:', err);
        throw err;
      }
    }
    const prev = fileStore.getData().dailyTeamOperations.length;
    fileStore.getData().dailyTeamOperations = fileStore.getData().dailyTeamOperations.filter((r) => r.id !== id);
    fileStore.scheduleFlush();
    return fileStore.getData().dailyTeamOperations.length < prev;
  }

  async clearAllDailyTeamOperations(): Promise<void> {
    if (db) {
      try {
        await db.delete(dailyTeamOperations);
        return;
      } catch (err) {
        console.error('Database clear failed:', err);
        throw err;
      }
    }
    fileStore.getData().dailyTeamOperations = [];
    fileStore.flush();
  }

  // Settings
  async getSetting(key: string): Promise<any> {
    if (db) {
      try {
        const [cat] = await db.select().from(categories).where(eq(categories.type, `setting_${key}`));
        if (cat) return JSON.parse(cat.name);
      } catch {}
    }
    return fileStore.getData().settings[key] ?? null;
  }

  async saveSetting(key: string, value: any): Promise<void> {
    if (db) {
      try {
        const settingKey = `setting_${key}`;
        const [existing] = await db.select().from(categories).where(eq(categories.type, settingKey));
        if (existing) {
          await db.update(categories).set({ name: JSON.stringify(value) }).where(eq(categories.id, existing.id));
        } else {
          await db.insert(categories).values({
            id: `setting_${Date.now()}_${Math.random().toString(36).substring(2, 7)}`,
            type: settingKey,
            name: JSON.stringify(value),
          });
        }
      } catch {}
    }
    fileStore.getData().settings[key] = value;
    fileStore.scheduleFlush();
  }
}

export const storage = new DatabaseStorage();
