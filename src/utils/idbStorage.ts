/**
 * IndexedDB High-Capacity Storage Manager
 * Allows persisting 100,000+ stock and visa records without encountering
 * the browser's 5MB localStorage QuotaExceededError.
 */

const DB_NAME = 'VisaStockMasterDB';
const DB_VERSION = 2;
const KNOWN_STORES = [
  'stock_records',
  'sticker_actual_stock',
  'visa_records',
  'categories',
  'officers',
  'users',
  'daily_team_operations',
  'yearly_team_distribution_report_draft_v6',
  'yearly_k2_receive_k1_report_draft_v6',
  'yearly_k2_issued_teams_report_draft_v1',
  'robok_report_latest_draft',
  'robok_director_report_latest_draft',
];

let dbPromise: Promise<IDBDatabase> | null = null;

function getDB(): Promise<IDBDatabase> {
  if (dbPromise) return dbPromise;

  dbPromise = new Promise((resolve, reject) => {
    if (typeof window === 'undefined' || !window.indexedDB) {
      return reject(new Error('IndexedDB is not supported in this environment'));
    }

    const request = window.indexedDB.open(DB_NAME, DB_VERSION);

    request.onupgradeneeded = (event) => {
      const db = (event.target as IDBOpenDBRequest).result;
      KNOWN_STORES.forEach((storeName) => {
        if (!db.objectStoreNames.contains(storeName)) {
          db.createObjectStore(storeName);
        }
      });
    };

    request.onsuccess = (event) => {
      const db = (event.target as IDBOpenDBRequest).result;
      resolve(db);
    };

    request.onerror = (event) => {
      console.warn('IndexedDB open error:', (event.target as IDBOpenDBRequest).error);
      reject((event.target as IDBOpenDBRequest).error);
    };
  });

  return dbPromise;
}

async function ensureStore(storeName: string): Promise<IDBDatabase> {
  const db = await getDB();
  if (db.objectStoreNames.contains(storeName)) {
    return db;
  }

  // If the store does not exist, safely upgrade the database to create it
  try {
    db.close();
  } catch (e) {}
  dbPromise = null;

  const nextVersion = (db.version || DB_VERSION) + 1;

  dbPromise = new Promise((resolve, reject) => {
    const request = window.indexedDB.open(DB_NAME, nextVersion);

    request.onupgradeneeded = (event) => {
      const upgradedDb = (event.target as IDBOpenDBRequest).result;
      if (!upgradedDb.objectStoreNames.contains(storeName)) {
        upgradedDb.createObjectStore(storeName);
      }
      KNOWN_STORES.forEach((s) => {
        if (!upgradedDb.objectStoreNames.contains(s)) {
          upgradedDb.createObjectStore(s);
        }
      });
    };

    request.onsuccess = (event) => {
      const upgradedDb = (event.target as IDBOpenDBRequest).result;
      resolve(upgradedDb);
    };

    request.onerror = (event) => {
      console.warn(`Failed to upgrade IndexedDB for store ${storeName}:`, (event.target as IDBOpenDBRequest).error);
      reject((event.target as IDBOpenDBRequest).error);
    };
  });

  return dbPromise;
}

const pendingWrites = new Map<string, { timer: ReturnType<typeof setTimeout>; value: any; resolvers: Array<(ok: boolean) => void> }>();

async function executeSetItem<T>(storeName: string, key: string, value: T): Promise<boolean> {
  try {
    const db = await ensureStore(storeName);
    if (!db.objectStoreNames.contains(storeName)) {
      return false;
    }
    return new Promise((resolve) => {
      try {
        const tx = db.transaction(storeName, 'readwrite');
        const store = tx.objectStore(storeName);
        const req = store.put(value, key);
        req.onsuccess = () => resolve(true);
        req.onerror = () => {
          console.warn(`IndexedDB setItem error for ${storeName}/${key}:`, req.error);
          resolve(false);
        };
      } catch (txErr) {
        console.warn(`IndexedDB transaction error for ${storeName}:`, txErr);
        resolve(false);
      }
    });
  } catch (e) {
    console.warn(`IndexedDB setItem failed for ${storeName}:`, e);
    return false;
  }
}

export const idbStorage = {
  async setItem<T>(storeName: string, key: string, value: T): Promise<boolean> {
    const isHeavyStore =
      storeName === 'stock_records' ||
      storeName === 'sticker_actual_stock' ||
      storeName === 'visa_records';

    if (!isHeavyStore) {
      return executeSetItem(storeName, key, value);
    }

    const mapKey = `${storeName}:::${key}`;
    return new Promise((resolve) => {
      const existing = pendingWrites.get(mapKey);
      if (existing) {
        clearTimeout(existing.timer);
        existing.value = value;
        existing.resolvers.push(resolve);
      } else {
        pendingWrites.set(mapKey, {
          value,
          resolvers: [resolve],
          timer: setTimeout(() => {}, 0),
        });
      }

      const entry = pendingWrites.get(mapKey)!;
      entry.timer = setTimeout(async () => {
        pendingWrites.delete(mapKey);
        const ok = await executeSetItem(storeName, key, entry.value);
        entry.resolvers.forEach((r) => r(ok));
      }, 600);
    });
  },

  async getItem<T>(storeName: string, key: string): Promise<T | null> {
    try {
      const db = await ensureStore(storeName);
      if (!db.objectStoreNames.contains(storeName)) {
        return null;
      }
      return new Promise((resolve) => {
        try {
          const tx = db.transaction(storeName, 'readonly');
          const store = tx.objectStore(storeName);
          const req = store.get(key);
          req.onsuccess = () => {
            resolve(req.result !== undefined ? (req.result as T) : null);
          };
          req.onerror = () => {
            console.warn(`IndexedDB getItem error for ${storeName}/${key}:`, req.error);
            resolve(null);
          };
        } catch (txErr) {
          console.warn(`IndexedDB transaction error for ${storeName}:`, txErr);
          resolve(null);
        }
      });
    } catch (e) {
      console.warn(`IndexedDB getItem failed for ${storeName}:`, e);
      return null;
    }
  },

  async removeItem(storeName: string, key: string): Promise<boolean> {
    try {
      const db = await ensureStore(storeName);
      if (!db.objectStoreNames.contains(storeName)) {
        return false;
      }
      return new Promise((resolve) => {
        try {
          const tx = db.transaction(storeName, 'readwrite');
          const store = tx.objectStore(storeName);
          const req = store.delete(key);
          req.onsuccess = () => resolve(true);
          req.onerror = () => resolve(false);
        } catch (txErr) {
          resolve(false);
        }
      });
    } catch (e) {
      return false;
    }
  },

  async clear(storeName: string): Promise<boolean> {
    try {
      const db = await ensureStore(storeName);
      if (!db.objectStoreNames.contains(storeName)) {
        return false;
      }
      return new Promise((resolve) => {
        try {
          const tx = db.transaction(storeName, 'readwrite');
          const store = tx.objectStore(storeName);
          const req = store.clear();
          req.onsuccess = () => resolve(true);
          req.onerror = () => resolve(false);
        } catch (txErr) {
          resolve(false);
        }
      });
    } catch (e) {
      return false;
    }
  },
};
