/**
 * Utils.ts — Shared utility functions for the Time Tracking Workbook
 * Used by all other scripts for session validation, hierarchy resolution, and helpers.
 */

// ── Constants ──────────────────────────────────────────────────────────────────
const SHEET_PASSWORD: string = "TT#2026!Sec";
const HASH_SALT: string = "LG_TimeTrack_2026";

// ── Color Constants ────────────────────────────────────────────────────────────
const COLORS: Record<string, string> = {
  darkBlue: "#243347",
  orange: "#EE742B",
  blueGreen: "#239A98",
  yellow: "#FDC400",
  softGreen: "#A8D0C0",
  red: "#D9415C",
  blush: "#F5B0A3",
  white: "#FFFFFF",
  grey: "#C5C6C6",
  anthracite: "#575756",
};

// ── Type Definitions ───────────────────────────────────────────────────────────
interface UserRow {
  userID: string;
  fullName: string;
  username: string;
  passwordHash: string;
  role: string;
  managerID: string;
  teamID: string;
  active: boolean;
}

interface SessionData {
  userID: string;
  username: string;
  role: string;
  managerID: string;
  loginTime: string;
}

interface TimeEntryRow {
  entryID: string;
  date: string;
  employeeID: string;
  employeeName: string;
  team: string;
  project: string;
  activity: string;
  hours: number;
  notes: string;
  submittedOn: string;
  lastEditedBy: string;
  lastEditedOn: string;
}

// ── Hash Function ──────────────────────────────────────────────────────────────
function simpleHash(input: string): string {
  const salted: string = HASH_SALT + input + HASH_SALT;
  let h1: number = 0xdeadbeef;
  let h2: number = 0x41c6ce57;
  for (let i: number = 0; i < salted.length; i++) {
    const ch: number = salted.charCodeAt(i);
    h1 = Math.imul(h1 ^ ch, 2654435761);
    h2 = Math.imul(h2 ^ ch, 1597334677);
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  const combined: number = 4294967296 * (2097151 & h2) + (h1 >>> 0);
  return combined.toString(16).padStart(16, "0");
}

// ── Session Management ─────────────────────────────────────────────────────────
function getSession(workbook: ExcelScript.Workbook): SessionData | null {
  const sheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("SESSION");
  if (!sheet) return null;

  const table: ExcelScript.Table | undefined = sheet.getTable("SessionTable");
  if (!table) return null;

  const rows: (string | number | boolean)[][] = table.getRangeBetweenHeaderAndTotal().getValues();
  if (rows.length === 0 || !rows[0][0]) return null;

  return {
    userID: String(rows[0][0]),
    username: String(rows[0][1]),
    role: String(rows[0][2]),
    managerID: String(rows[0][3]),
    loginTime: String(rows[0][4]),
  };
}

function requireSession(workbook: ExcelScript.Workbook): SessionData {
  const session: SessionData | null = getSession(workbook);
  if (!session) {
    throw new Error("No active session. Please log in first.");
  }
  return session;
}

function writeSession(workbook: ExcelScript.Workbook, data: SessionData): void {
  const sheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("SESSION");
  if (!sheet) return;

  const table: ExcelScript.Table | undefined = sheet.getTable("SessionTable");
  if (!table) return;

  const body: ExcelScript.Range = table.getRangeBetweenHeaderAndTotal();
  if (body.getRowCount() > 0) {
    body.delete(ExcelScript.DeleteShiftDirection.up);
  }

  table.addRow(-1, [data.userID, data.username, data.role, data.managerID, data.loginTime]);
}

function clearSession(workbook: ExcelScript.Workbook): void {
  const sheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("SESSION");
  if (!sheet) return;

  const table: ExcelScript.Table | undefined = sheet.getTable("SessionTable");
  if (!table) return;

  const body: ExcelScript.Range = table.getRangeBetweenHeaderAndTotal();
  if (body.getRowCount() > 0) {
    body.delete(ExcelScript.DeleteShiftDirection.up);
  }
}

// ── User Data Helpers ──────────────────────────────────────────────────────────
function getAllUsers(workbook: ExcelScript.Workbook): UserRow[] {
  const sheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("USERS_DB");
  if (!sheet) return [];

  const table: ExcelScript.Table | undefined = sheet.getTable("UsersTable");
  if (!table) return [];

  const rows: (string | number | boolean)[][] = table.getRangeBetweenHeaderAndTotal().getValues();
  return rows.map((r: (string | number | boolean)[]): UserRow => ({
    userID: String(r[0]),
    fullName: String(r[1]),
    username: String(r[2]),
    passwordHash: String(r[3]),
    role: String(r[4]),
    managerID: String(r[5]),
    teamID: String(r[6]),
    active: Boolean(r[7]),
  }));
}

function findUserByUsername(users: UserRow[], username: string): UserRow | undefined {
  const lower: string = username.toLowerCase();
  return users.find((u: UserRow): boolean => u.username.toLowerCase() === lower && u.active);
}

function findUserByID(users: UserRow[], userID: string): UserRow | undefined {
  return users.find((u: UserRow): boolean => u.userID === userID);
}

// ── Hierarchy Resolution ───────────────────────────────────────────────────────
function getSubtreeUserIDs(managerID: string, allUsers: UserRow[]): string[] {
  const result: string[] = [];
  const directReports: UserRow[] = allUsers.filter((u: UserRow): boolean => u.managerID === managerID && u.active);
  for (const report of directReports) {
    result.push(report.userID);
    const subReports: string[] = getSubtreeUserIDs(report.userID, allUsers);
    result.push(...subReports);
  }
  return result;
}

function getRoleLevel(role: string): number {
  const levels: Record<string, number> = {
    Admin: 0,
    Director: 1,
    Manager: 2,
    TeamLead: 3,
    Employee: 4,
  };
  return levels[role] ?? 4;
}

function canUserEditEntry(sessionUserID: string, entryOwnerID: string, allUsers: UserRow[]): boolean {
  const sessionUser: UserRow | undefined = findUserByID(allUsers, sessionUserID);
  if (!sessionUser) return false;

  if (sessionUser.role === "Admin") return true;
  if (sessionUserID === entryOwnerID) return true;

  const subtree: string[] = getSubtreeUserIDs(sessionUserID, allUsers);
  return subtree.includes(entryOwnerID);
}

// ── Date Helpers ───────────────────────────────────────────────────────────────
function getWeekStart(d: Date): Date {
  const result: Date = new Date(d);
  const day: number = result.getDay();
  const diff: number = result.getDate() - day + (day === 0 ? -6 : 1);
  result.setDate(diff);
  result.setHours(0, 0, 0, 0);
  return result;
}

function getWeekEnd(d: Date): Date {
  const start: Date = getWeekStart(d);
  const end: Date = new Date(start);
  end.setDate(start.getDate() + 4);
  return end;
}

function formatDate(d: Date): string {
  const y: number = d.getFullYear();
  const m: string = String(d.getMonth() + 1).padStart(2, "0");
  const day: string = String(d.getDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function parseExcelDate(value: string | number | boolean): Date {
  if (typeof value === "number") {
    const epoch: Date = new Date(1899, 11, 30);
    epoch.setDate(epoch.getDate() + value);
    return epoch;
  }
  return new Date(String(value));
}

function generateEntryID(): string {
  const now: Date = new Date();
  const ts: string = now.getTime().toString(36);
  const rand: string = Math.random().toString(36).substring(2, 6);
  return `TE-${ts}-${rand}`.toUpperCase();
}

// ── Sheet Visibility Helpers ───────────────────────────────────────────────────
function hideAllWorksheets(workbook: ExcelScript.Workbook): void {
  const alwaysVisible: string[] = ["LOGIN", "README"];
  const sheets: ExcelScript.Worksheet[] = workbook.getWorksheets();
  for (const sheet of sheets) {
    const name: string = sheet.getName();
    if (alwaysVisible.includes(name)) {
      sheet.setVisibility(ExcelScript.SheetVisibility.visible);
    } else {
      sheet.setVisibility(ExcelScript.SheetVisibility.hidden);
    }
  }
  workbook.getWorksheet("LOGIN")?.activate();
}

function showSheetsForRole(workbook: ExcelScript.Workbook, role: string): void {
  const roleLevel: number = getRoleLevel(role);

  const visibleSheets: string[] = ["TIME_ENTRY", "DAILY_CHECK"];

  if (roleLevel <= 3) {
    visibleSheets.push("MANAGER_VIEW");
  }
  if (roleLevel <= 2) {
    visibleSheets.push("REPORTS");
  }
  if (roleLevel === 0) {
    visibleSheets.push("ADMIN", "USERS_DB", "PROJECTS_DB", "ACTIVITIES_DB", "TEAMS_DB");
  }

  visibleSheets.push("README");

  for (const name of visibleSheets) {
    const sheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet(name);
    if (sheet) {
      sheet.setVisibility(ExcelScript.SheetVisibility.visible);
    }
  }

  const loginSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("LOGIN");
  if (loginSheet) {
    loginSheet.setVisibility(ExcelScript.SheetVisibility.hidden);
  }

  workbook.getWorksheet("TIME_ENTRY")?.activate();
}

// ── Table Reading Helpers ──────────────────────────────────────────────────────
function getAllTimeEntries(workbook: ExcelScript.Workbook): TimeEntryRow[] {
  const sheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("TIME_ENTRY");
  if (!sheet) return [];

  const table: ExcelScript.Table | undefined = sheet.getTable("TimeEntryTable");
  if (!table) return [];

  const rows: (string | number | boolean)[][] = table.getRangeBetweenHeaderAndTotal().getValues();
  return rows.map((r: (string | number | boolean)[]): TimeEntryRow => ({
    entryID: String(r[0]),
    date: String(r[1]),
    employeeID: String(r[2]),
    employeeName: String(r[3]),
    team: String(r[4]),
    project: String(r[5]),
    activity: String(r[6]),
    hours: Number(r[7]),
    notes: String(r[8]),
    submittedOn: String(r[9]),
    lastEditedBy: String(r[10]),
    lastEditedOn: String(r[11]),
  }));
}

function getTeamName(workbook: ExcelScript.Workbook, teamID: string): string {
  const sheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("TEAMS_DB");
  if (!sheet) return "";

  const table: ExcelScript.Table | undefined = sheet.getTable("TeamsTable");
  if (!table) return "";

  const rows: (string | number | boolean)[][] = table.getRangeBetweenHeaderAndTotal().getValues();
  for (const r of rows) {
    if (String(r[0]) === teamID) return String(r[1]);
  }
  return "";
}
