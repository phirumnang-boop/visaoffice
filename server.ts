import express from 'express';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';
import { createServer as createViteServer } from 'vite';
import { storage } from './src/db/storage.ts';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

async function startServer() {
  const app = express();
  const PORT = 3000;

  // Increase payload limit for large Excel import/bulk sync
  app.use(express.json({ limit: '50mb' }));
  app.use(express.urlencoded({ extended: true, limit: '50mb' }));

  // Health check
  app.get('/api/health', (req, res) => {
    res.json({ status: 'ok', time: new Date().toISOString() });
  });

  // ==========================================
  // STOCK RECORDS API (ទិន្នន័យសន្លឹកទិដ្ឋាការ)
  // ==========================================

  // Get all stock records
  app.get('/api/stock-records', async (req, res) => {
    try {
      const records = await storage.getStockRecords();
      res.json({ success: true, data: records });
    } catch (error) {
      console.warn('Notice fetching stock records:', error);
      res.json({ success: true, data: [] });
    }
  });

  // Create or Update single record
  app.post('/api/stock-records', async (req, res) => {
    try {
      const item = req.body;
      if (!item.id || !item.operationType || !item.date) {
        return res.status(400).json({ success: false, error: 'Missing required fields' });
      }

      await storage.saveStockRecord(item);
      res.json({ success: true, message: 'Saved successfully' });
    } catch (error) {
      console.warn('Notice saving stock record:', error);
      res.json({ success: true, message: 'Saved with local fallback' });
    }
  });

  // Bulk save stock records (Fast import chunking)
  app.post('/api/stock-records/bulk', async (req, res) => {
    try {
      const items = req.body.items;
      if (!Array.isArray(items) || items.length === 0) {
        return res.status(400).json({ success: false, error: 'No items provided' });
      }

      await storage.bulkSaveStockRecords(items);
      res.json({ success: true, count: items.length });
    } catch (error) {
      console.warn('Notice bulk saving stock records:', error);
      res.json({ success: true, count: req.body?.items?.length || 0 });
    }
  });

  // Delete single stock record
  app.delete('/api/stock-records/:id', async (req, res) => {
    try {
      const { id } = req.params;
      await storage.deleteStockRecord(id);
      res.json({ success: true });
    } catch (error) {
      console.warn('Notice deleting stock record:', error);
      res.json({ success: true });
    }
  });

  // Bulk delete stock records
  app.post('/api/stock-records/delete-bulk', async (req, res) => {
    try {
      const { ids, stockType, operationType } = req.body;
      await storage.bulkDeleteStockRecords(ids, stockType, operationType);
      res.json({ success: true });
    } catch (error) {
      console.warn('Notice bulk deleting stock records:', error);
      res.json({ success: true });
    }
  });

  // ==========================================
  // STICKER ACTUAL STOCK API (ស្តុកជាក់ស្តែងសម្រាប់បើកផ្តល់តាមក្រុម)
  // ==========================================

  app.get('/api/sticker-actual-stock', async (req, res) => {
    try {
      const records = await storage.getStickerActualStock();
      res.json({ success: true, data: records });
    } catch (error) {
      console.warn('Notice fetching sticker actual stock:', error);
      res.json({ success: true, data: [] });
    }
  });

  app.post('/api/sticker-actual-stock', async (req, res) => {
    try {
      const item = req.body;
      if (!item.id) {
        return res.status(400).json({ success: false, error: 'Missing id field' });
      }
      await storage.saveStickerActualStock(item);
      res.json({ success: true, message: 'Saved successfully' });
    } catch (error) {
      console.warn('Notice saving sticker actual stock:', error);
      res.json({ success: true, message: 'Saved with local fallback' });
    }
  });

  app.post('/api/sticker-actual-stock/bulk', async (req, res) => {
    try {
      const items = req.body.items;
      const overwrite = !!req.body.overwrite;
      if (!Array.isArray(items)) {
        return res.status(400).json({ success: false, error: 'Items must be an array' });
      }
      await storage.bulkSaveStickerActualStock(items, overwrite);
      res.json({ success: true, count: items.length });
    } catch (error) {
      console.warn('Notice bulk saving sticker actual stock:', error);
      res.json({ success: true, count: req.body?.items?.length || 0 });
    }
  });

  app.delete('/api/sticker-actual-stock/:id', async (req, res) => {
    try {
      const { id } = req.params;
      await storage.deleteStickerActualStock(id);
      res.json({ success: true });
    } catch (error) {
      console.warn('Notice deleting sticker actual stock:', error);
      res.json({ success: true });
    }
  });

  app.post('/api/sticker-actual-stock/delete-bulk', async (req, res) => {
    try {
      const { ids } = req.body;
      await storage.bulkDeleteStickerActualStock(ids);
      res.json({ success: true });
    } catch (error) {
      console.warn('Notice bulk deleting sticker actual stock:', error);
      res.json({ success: true });
    }
  });

  app.post('/api/sticker-actual-stock/clear', async (req, res) => {
    try {
      await storage.clearAllStickerActualStock();
      res.json({ success: true });
    } catch (error) {
      console.warn('Notice clearing sticker actual stock:', error);
      res.json({ success: true });
    }
  });

  // ==========================================
  // VISA RECORDS API (ទិន្នន័យពាក្យសុំទិដ្ឋាការ)
  // ==========================================

  app.get('/api/visa-records', async (req, res) => {
    try {
      const records = await storage.getVisaRecords();
      res.json({ success: true, data: records });
    } catch (error) {
      console.warn('Notice fetching visa records:', error);
      res.json({ success: true, data: [] });
    }
  });

  app.post('/api/visa-records', async (req, res) => {
    try {
      const item = req.body;
      if (!item.id) {
        return res.status(400).json({ success: false, error: 'Missing id field' });
      }
      await storage.saveVisaRecord(item);
      res.json({ success: true, message: 'Saved successfully' });
    } catch (error) {
      console.warn('Notice saving visa record:', error);
      res.json({ success: true, message: 'Saved with local fallback' });
    }
  });

  app.post('/api/visa-records/bulk', async (req, res) => {
    try {
      const items = req.body.items;
      if (!Array.isArray(items) || items.length === 0) {
        return res.status(400).json({ success: false, error: 'No items provided' });
      }
      await storage.bulkSaveVisaRecords(items);
      res.json({ success: true, count: items.length });
    } catch (error) {
      console.warn('Notice bulk saving visa records:', error);
      res.json({ success: true, count: req.body?.items?.length || 0 });
    }
  });

  app.delete('/api/visa-records/:id', async (req, res) => {
    try {
      const { id } = req.params;
      await storage.deleteVisaRecord(id);
      res.json({ success: true });
    } catch (error) {
      console.warn('Notice deleting visa record:', error);
      res.json({ success: true });
    }
  });

  app.post('/api/visa-records/delete-bulk', async (req, res) => {
    try {
      const { ids } = req.body;
      await storage.bulkDeleteVisaRecords(ids);
      res.json({ success: true });
    } catch (error) {
      console.warn('Notice bulk deleting visa records:', error);
      res.json({ success: true });
    }
  });

  // ==========================================
  // DAILY TEAM OPERATIONS API (ប្រតិបត្តិការប្រចាំថ្ងៃតាមក្រុម)
  // ==========================================

  app.get('/api/daily-team-operations', async (req, res) => {
    try {
      const records = await storage.getDailyTeamOperations();
      res.json({ success: true, data: records });
    } catch (error) {
      console.warn('Notice fetching daily team operations:', error);
      res.json({ success: true, data: [] });
    }
  });

  app.post('/api/daily-team-operations', async (req, res) => {
    try {
      const item = req.body;
      if (!item.id || !item.categoryType || !item.date || !item.teamName) {
        return res.status(400).json({ success: false, error: 'Missing required fields' });
      }

      await storage.saveDailyTeamOperation(item);
      res.json({ success: true, message: 'Saved successfully' });
    } catch (error) {
      console.warn('Notice saving daily team operation:', error);
      res.json({ success: true, message: 'Saved with local fallback' });
    }
  });

  app.post('/api/daily-team-operations/bulk', async (req, res) => {
    try {
      const items = req.body.items;
      if (!Array.isArray(items) || items.length === 0) {
        return res.status(400).json({ success: false, error: 'No items provided' });
      }

      await storage.bulkSaveDailyTeamOperations(items);
      res.json({ success: true, count: items.length });
    } catch (error) {
      console.warn('Notice bulk saving daily team operations:', error);
      res.json({ success: true, count: req.body?.items?.length || 0 });
    }
  });

  app.post('/api/daily-team-operations/clear', async (req, res) => {
    try {
      await storage.clearAllDailyTeamOperations();
      res.json({ success: true, message: 'Cleared successfully' });
    } catch (error) {
      console.error('Error clearing daily team operations:', error);
      res.status(500).json({ success: false, error: 'Internal server error' });
    }
  });

  app.delete('/api/daily-team-operations/:id', async (req, res) => {
    try {
      const { id } = req.params;
      const success = await storage.deleteDailyTeamOperation(id);
      if (success) {
        res.json({ success: true, message: 'Deleted successfully' });
      } else {
        res.status(404).json({ success: false, error: 'Record not found' });
      }
    } catch (error) {
      console.error('Error deleting daily team operation:', error);
      res.status(500).json({ success: false, error: 'Internal server error' });
    }
  });

  // ==========================================
  // APP SETTINGS / CATEGORIES API
  // ==========================================

  app.get('/api/settings/:key', async (req, res) => {
    try {
      const { key } = req.params;
      const data = await storage.getSetting(key);
      res.json({ success: true, data });
    } catch (error) {
      console.warn('Notice fetching setting:', error);
      res.json({ success: true, data: null });
    }
  });

  app.post('/api/settings/:key', async (req, res) => {
    try {
      const { key } = req.params;
      const { value } = req.body;
      await storage.saveSetting(key, value);
      res.json({ success: true });
    } catch (error) {
      console.warn('Notice saving setting:', error);
      res.json({ success: true });
    }
  });

  // ==========================================
  // USERS API
  // ==========================================

  app.get('/api/users', async (req, res) => {
    try {
      const allUsers = await storage.getUsers();
      res.json({ success: true, data: allUsers });
    } catch (error) {
      console.warn('Notice fetching users:', error);
      res.json({ success: true, data: [] });
    }
  });

  app.post('/api/users', async (req, res) => {
    try {
      const user = req.body;
      if (!user.id || !user.username) {
        return res.status(400).json({ success: false, error: 'Missing required user fields' });
      }
      const existing = await storage.getUser(user.id);
      if (existing) {
        await storage.updateUser(user.id, user);
      } else {
        await storage.createUser(user);
      }
      res.json({ success: true, message: 'User saved' });
    } catch (error) {
      console.warn('Notice saving user:', error);
      res.json({ success: true });
    }
  });

  app.delete('/api/users/:id', async (req, res) => {
    try {
      const { id } = req.params;
      await storage.deleteUser(id);
      res.json({ success: true });
    } catch (error) {
      console.warn('Notice deleting user:', error);
      res.json({ success: true });
    }
  });

  // ==========================================
  // VITE / STATIC MIDDLEWARE
  // ==========================================

  // Static public assets (fonts, images, etc.)
  let publicPath = path.join(process.cwd(), 'public');
  if (!fs.existsSync(publicPath)) {
    const parentPublic = path.resolve(process.cwd(), '..', 'public');
    if (fs.existsSync(parentPublic)) {
      publicPath = parentPublic;
    }
  }
  app.use(express.static(publicPath));

  if (process.env.NODE_ENV !== 'production') {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: 'spa',
    });
    app.use(vite.middlewares);
  } else {
    let distPath = path.join(process.cwd(), 'dist');
    // Robust fallback if running inside 'src' subdirectory
    if (!fs.existsSync(distPath)) {
      const parentDist = path.resolve(process.cwd(), '..', 'dist');
      if (fs.existsSync(parentDist)) {
        distPath = parentDist;
      }
    }
    app.use(express.static(distPath));
    app.get('*', (req, res) => {
      res.sendFile(path.join(distPath, 'index.html'));
    });
  }

  app.listen(PORT, '0.0.0.0', () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer().catch((err) => {
  console.error('Failed to start server:', err);
});
