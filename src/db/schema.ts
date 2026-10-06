import { pgTable, text, varchar, integer, timestamp, jsonb, boolean } from 'drizzle-orm/pg-core';
import { sql } from 'drizzle-orm';

// Users Table
export const users = pgTable('users', {
  id: varchar('id', { length: 128 }).primaryKey(),
  username: varchar('username', { length: 255 }).notNull(),
  officerNumber: varchar('officer_number', { length: 100 }),
  email: varchar('email', { length: 255 }),
  password: text('password'),
  role: varchar('role', { length: 50 }).notNull().default('User'),
  createdAt: timestamp('created_at').default(sql`CURRENT_TIMESTAMP`),
});

// Categories Table
export const categories = pgTable('categories', {
  id: varchar('id', { length: 128 }).primaryKey(),
  type: varchar('type', { length: 100 }).notNull(), // 'ranks', 'positions', 'visaTeams', etc.
  name: text('name').notNull(),
  createdAt: timestamp('created_at').default(sql`CURRENT_TIMESTAMP`),
});

// Officers Table
export const officers = pgTable('officers', {
  id: varchar('id', { length: 128 }).primaryKey(),
  name: varchar('name', { length: 255 }).notNull(),
  rankId: varchar('rank_id', { length: 128 }),
  rankName: varchar('rank_name', { length: 255 }),
  gender: varchar('gender', { length: 10 }),
  officerNumber: varchar('officer_number', { length: 100 }),
  dob: varchar('dob', { length: 50 }),
  positionId: varchar('position_id', { length: 128 }),
  positionName: varchar('position_name', { length: 255 }),
  officeWork: varchar('office_work', { length: 100 }),
  visaTeamId: varchar('visa_team_id', { length: 128 }),
  visaTeamName: text('visa_team_name'),
  phone: varchar('phone', { length: 100 }),
  createdAt: timestamp('created_at').default(sql`CURRENT_TIMESTAMP`),
  createdBy: varchar('created_by', { length: 255 }),
});

// Visa Records Table
export const visaRecords = pgTable('visa_records', {
  id: varchar('id', { length: 128 }).primaryKey(),
  passportNumber: varchar('passport_number', { length: 100 }),
  fullName: varchar('full_name', { length: 255 }),
  nationality: varchar('nationality', { length: 100 }),
  gender: varchar('gender', { length: 20 }),
  oldVisaType: varchar('old_visa_type', { length: 50 }),
  newVisaType: varchar('new_visa_type', { length: 50 }),
  documentNumber: varchar('document_number', { length: 100 }),
  issueDate: varchar('issue_date', { length: 50 }),
  organization: text('organization'),
  extension: varchar('extension', { length: 100 }),
  duration: varchar('duration', { length: 100 }),
  fee: integer('fee'),
  applicationDate: varchar('application_date', { length: 50 }),
  officerId: varchar('officer_id', { length: 128 }),
  officerName: varchar('officer_name', { length: 255 }),
  visaTeamId: varchar('visa_team_id', { length: 128 }),
  status: varchar('status', { length: 50 }).default('អនុម័តរួច'),
  notes: text('notes'),
  createdAt: timestamp('created_at').default(sql`CURRENT_TIMESTAMP`),
});

// Stock Records Table
export const stockRecords = pgTable('stock_records', {
  id: varchar('id', { length: 128 }).primaryKey(),
  recordType: varchar('record_type', { length: 20 }), // 'IN' | 'OUT'
  stockType: varchar('stock_type', { length: 50 }), // 'evisa' | 'sticker'
  operationType: varchar('operation_type', { length: 100 }),
  sourceFrom: varchar('source_from', { length: 100 }),
  date: varchar('date', { length: 50 }),
  recordDate: varchar('record_date', { length: 50 }),
  time: varchar('time', { length: 50 }),
  quantity: integer('quantity'),
  quantityBundles: integer('quantity_bundles'),
  totalSheets: integer('total_sheets'),
  visaType: varchar('visa_type', { length: 50 }),
  startSerial: varchar('start_serial', { length: 100 }),
  endSerial: varchar('end_serial', { length: 100 }),
  startNumber: varchar('start_number', { length: 100 }),
  endNumber: varchar('end_number', { length: 100 }),
  visaTeamId: varchar('visa_team_id', { length: 128 }),
  visaTeamName: text('visa_team_name'),
  visaTeamRobokId: varchar('visa_team_robok_id', { length: 128 }),
  visaTeamRobokName: text('visa_team_robok_name'),
  teamId: varchar('team_id', { length: 128 }),
  teamName: text('team_name'),
  officerId: varchar('officer_id', { length: 128 }),
  officerName: varchar('officer_name', { length: 255 }),
  requestedRankId: varchar('requested_rank_id', { length: 128 }),
  requestedRankName: varchar('requested_rank_name', { length: 255 }),
  requesterName: varchar('requester_name', { length: 255 }),
  collectorName: varchar('collector_name', { length: 255 }),
  collectorRole: varchar('collector_role', { length: 255 }),
  collectorRoleId: varchar('collector_role_id', { length: 128 }),
  collectorRoleName: varchar('collector_role_name', { length: 255 }),
  documentNumber: varchar('document_number', { length: 100 }),
  description: text('description'),
  notes: text('notes'),
  isImported: boolean('is_imported').default(false),
  createdAt: timestamp('created_at').default(sql`CURRENT_TIMESTAMP`),
  createdBy: varchar('created_by', { length: 255 }),
});

// Daily Team Operations Table
export const dailyTeamOperations = pgTable('daily_team_operations', {
  id: varchar('id', { length: 128 }).primaryKey(),
  date: varchar('date', { length: 50 }).notNull(),
  teamName: text('team_name').notNull(),
  teamId: varchar('team_id', { length: 128 }),
  stickerUsage: jsonb('sticker_usage'),
  evisaUsage: integer('evisa_usage').default(0),
  notes: text('notes'),
  createdAt: timestamp('created_at').default(sql`CURRENT_TIMESTAMP`),
});

export type User = typeof users.$inferSelect;
export type InsertUser = typeof users.$inferInsert;

export type Category = typeof categories.$inferSelect;
export type InsertCategory = typeof categories.$inferInsert;

export type DbOfficer = typeof officers.$inferSelect;
export type InsertOfficer = typeof officers.$inferInsert;

export type DbVisaRecord = typeof visaRecords.$inferSelect;
export type InsertVisaRecord = typeof visaRecords.$inferInsert;

export type DbStockRecord = typeof stockRecords.$inferSelect;
export type InsertStockRecord = typeof stockRecords.$inferInsert;

export type DbDailyTeamOperation = typeof dailyTeamOperations.$inferSelect;
export type InsertDailyTeamOperation = typeof dailyTeamOperations.$inferInsert;
