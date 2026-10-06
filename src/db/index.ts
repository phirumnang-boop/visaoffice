import { drizzle } from 'drizzle-orm/node-postgres';
import { Pool } from 'pg';
import * as schema from './schema.ts';

declare global {
  var _postgresPool: Pool | undefined;
}

export const isSqlConfigured = Boolean(
  process.env.DATABASE_URL ||
  (process.env.SQL_HOST &&
   process.env.SQL_USER &&
   process.env.SQL_PASSWORD &&
   process.env.SQL_DB_NAME)
);

export const createPool = (): Pool | null => {
  if (!isSqlConfigured) {
    return null;
  }

  if (!global._postgresPool) {
    try {
      if (process.env.DATABASE_URL) {
        global._postgresPool = new Pool({
          connectionString: process.env.DATABASE_URL,
          ssl: { rejectUnauthorized: false },
          max: 10,
          connectionTimeoutMillis: 5000,
        });
      } else {
        global._postgresPool = new Pool({
          host: process.env.SQL_HOST,
          user: process.env.SQL_USER,
          password: process.env.SQL_PASSWORD,
          database: process.env.SQL_DB_NAME,
          port: process.env.SQL_PORT ? Number(process.env.SQL_PORT) : 5432,
          max: 10,
          connectionTimeoutMillis: 3000,
        });
      }

      global._postgresPool.on('error', (err) => {
        console.warn('PostgreSQL pool connection note:', err.message);
      });
    } catch (e) {
      console.warn('PostgreSQL pool initialization note:', e);
      return null;
    }
  }
  return global._postgresPool;
};

const pool = createPool();

export const db = pool ? drizzle(pool, { schema }) : null;
