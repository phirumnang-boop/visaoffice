import { StockRecord, VisaRecord, DailyTeamRecord, UserAccount } from '../types';

export const checkIsFirestoreQuotaExceeded = () => false;

export const apiService = {
  // Stock records (ទិន្នន័យសន្លឹកទិដ្ឋាការ)
  async getStockRecords(): Promise<StockRecord[]> {
    try {
      const res = await fetch('/api/stock-records');
      if (!res.ok) throw new Error('Fetch failed');
      const data = await res.json();
      return Array.isArray(data?.data) ? data.data : (Array.isArray(data) ? data : []);
    } catch {
      return [];
    }
  },

  async saveStockRecord(record: StockRecord): Promise<void> {
    try {
      await fetch('/api/stock-records', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(record),
      });
    } catch {}
  },

  async bulkSaveStockRecords(records: StockRecord[]): Promise<void> {
    try {
      await fetch('/api/stock-records/bulk', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ items: records }),
      });
    } catch {}
  },

  async deleteStockRecord(id: string): Promise<void> {
    try {
      await fetch(`/api/stock-records/${id}`, {
        method: 'DELETE',
      });
    } catch {}
  },

  async bulkDeleteStockRecords(ids: string[], stockType?: string, operationType?: string): Promise<void> {
    try {
      await fetch('/api/stock-records/delete-bulk', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ids, stockType, operationType }),
      });
    } catch {}
  },

  // Dedicated Sticker Actual Stock (ស្តុកជាក់ស្តែងសម្រាប់បើកផ្តល់តាមក្រុម)
  async getStickerActualStock(): Promise<StockRecord[]> {
    try {
      const res = await fetch('/api/sticker-actual-stock');
      if (!res.ok) throw new Error('Fetch failed');
      const data = await res.json();
      return Array.isArray(data?.data) ? data.data : (Array.isArray(data) ? data : []);
    } catch {
      return [];
    }
  },

  async saveStickerActualStock(record: StockRecord): Promise<void> {
    try {
      await fetch('/api/sticker-actual-stock', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(record),
      });
    } catch {}
  },

  async bulkSaveStickerActualStock(records: StockRecord[], overwrite: boolean = false): Promise<void> {
    try {
      await fetch('/api/sticker-actual-stock/bulk', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ items: records, overwrite }),
      });
    } catch {}
  },

  async deleteStickerActualStock(id: string): Promise<void> {
    try {
      await fetch(`/api/sticker-actual-stock/${id}`, {
        method: 'DELETE',
      });
    } catch {}
  },

  async bulkDeleteStickerActualStock(ids: string[]): Promise<void> {
    try {
      await fetch('/api/sticker-actual-stock/delete-bulk', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ids }),
      });
    } catch {}
  },

  async clearAllStickerActualStock(): Promise<void> {
    try {
      await fetch('/api/sticker-actual-stock/clear', {
        method: 'POST',
      });
    } catch {}
  },

  // Visa records
  async getVisaRecords(): Promise<VisaRecord[]> {
    try {
      const res = await fetch('/api/visa-records');
      if (!res.ok) throw new Error('Fetch failed');
      const data = await res.json();
      return Array.isArray(data?.data) ? data.data : (Array.isArray(data) ? data : []);
    } catch {
      return [];
    }
  },

  async saveVisaRecord(record: VisaRecord): Promise<void> {
    try {
      await fetch('/api/visa-records', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(record),
      });
    } catch {}
  },

  async bulkSaveVisaRecords(records: VisaRecord[]): Promise<void> {
    try {
      await fetch('/api/visa-records/bulk', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ items: records }),
      });
    } catch {}
  },

  async deleteVisaRecord(id: string): Promise<void> {
    try {
      await fetch(`/api/visa-records/${id}`, {
        method: 'DELETE',
      });
    } catch {}
  },

  async bulkDeleteVisaRecords(ids: string[]): Promise<void> {
    try {
      await fetch('/api/visa-records/delete-bulk', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ ids }),
      });
    } catch {}
  },

  // Daily team operations
  async getDailyTeamOperations(): Promise<DailyTeamRecord[]> {
    try {
      const res = await fetch('/api/daily-team-operations');
      if (!res.ok) throw new Error('Fetch failed');
      const data = await res.json();
      return Array.isArray(data?.data) ? data.data : (Array.isArray(data) ? data : []);
    } catch {
      const saved = localStorage.getItem('app_daily_team_operations_v5');
      if (saved) {
        try {
          return JSON.parse(saved);
        } catch {}
      }
      return [];
    }
  },

  async saveDailyTeamOperation(record: DailyTeamRecord): Promise<void> {
    try {
      await fetch('/api/daily-team-operations', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(record),
      });
    } catch {}
  },

  async bulkSaveDailyTeamOperations(records: DailyTeamRecord[]): Promise<void> {
    try {
      await fetch('/api/daily-team-operations/bulk', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ items: records }),
      });
    } catch {}
  },

  async deleteDailyTeamOperation(id: string): Promise<boolean> {
    const res = await fetch(`/api/daily-team-operations/${id}`, {
      method: 'DELETE',
    });
    return res.ok;
  },

  async clearAllDailyTeamOperations(): Promise<void> {
    await fetch('/api/daily-team-operations/clear', {
      method: 'POST',
    });
  },

  // Settings
  async getSetting<T = any>(key: string): Promise<T | null> {
    try {
      const res = await fetch(`/api/settings/${encodeURIComponent(key)}`);
      if (!res.ok) throw new Error('Fetch setting failed');
      const data = await res.json();
      return (data?.data !== undefined ? data.data : data) as T;
    } catch {
      return null;
    }
  },

  async saveSetting<T = any>(key: string, value: T): Promise<void> {
    try {
      await fetch(`/api/settings/${encodeURIComponent(key)}`, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ value }),
      });
    } catch {}
  },

  // Users
  async getCloudUsers(): Promise<UserAccount[]> {
    try {
      const res = await fetch('/api/users');
      if (!res.ok) throw new Error('Fetch users failed');
      const data = await res.json();
      if (Array.isArray(data?.data)) {
        return data.data;
      }
    } catch {}
    // Fallback to local storage
    try {
      const saved = localStorage.getItem('app_users');
      if (saved) return JSON.parse(saved);
    } catch {}
    return [];
  },

  async saveCloudUser(user: UserAccount): Promise<void> {
    try {
      await fetch('/api/users', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(user),
      });
    } catch {}
  },

  async deleteCloudUser(id: string): Promise<void> {
    try {
      await fetch(`/api/users/${id}`, {
        method: 'DELETE',
      });
    } catch {}
  }
};
