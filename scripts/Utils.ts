/**
 * Utils.ts — Shared utility functions for the Time Tracking Workbook
 * Used by all other scripts for session validation, hierarchy resolution, and helpers.
 */

// ── Constants ──────────────────────────────────────────────────────────────────
const SHEET_PASSWORD = "TT#2026!Sec";
const HASH_SALT = "LG_TimeTrack_2026";

// ── Color Constants ────────────────────────────────────────────────────────────
const COLORS = {
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
/**
 * Simple deterministic hash for internal use. NOT cryptographic-grade.
 * Uses character codes + salt constant to produce a hex-like string.
 */
function simpleHash(input: string): string {
  const salted = HASH_SALT + input + HASH_SALT;
  let h1 = 0xdeadbeef;
  let h2 = 0x41c6ce57;
  for (let i = 0; i < salted.length; i++) {
    const ch = salted.charCodeAt(i);
    h1 = Math.imul(h1 ^ ch, 2654435761);
    h2 = Math.imul(h2 ^ ch, 1597334677);
  }
  h1 = Math.imul(h1 ^ (h1 >>> 16), 2246822507) ^ Math.imul(h2 ^ (h2 >>> 13), 3266489909);
  h2 = Math.imul(h2 ^ (h2 >>> 16), 2246822507) ^ Math.imul(h1 ^ (h1 >>> 13), 3266489909);
  const combined = 4294967296 * (2097151 & h2) + (h1 >>> 0);
  return combined.toString(16).padStart(16, "0");
}

// ── Session Management ─────────────────────────────────────────────────────────
/**
 * Reads the current session from the hidden SESSION sheet.
 * Returns null if no valid session exists.
 */
function getSession(workbook: ExcelScript.Workbook): SessionData | null {
  const sheet = workbook.getWorksheet("SESSION");
  if (!sheet) return null;

  const table = sheet.getTable("SessionTable");
  if (!table) return null;

  const rows = table.getRangeBetweenHeaderAndTotal().getValues();
  if (rows.length === 0 || !rows[0][0]) return null;

  return {
    userID: String(rows[0][0]),
    username: String(rows[0][1]),
    role: String(rows[0][2]),
    managerID: String(rows[0][3]),
    loginTime: String(rows[0][4]),
  };
}

/**
 * Validates that a session exists; returns the session or throws an alert message.
 */
function requireSession(workbook: ExcelScript.Workbook): SessionData {
  const session = getSession(workbook);
  if (!session) {
    throw new Error("No active session. Please log in first.");
  }
  return session;
}

/**
 * Writes session data to the SESSION sheet.
 */
function writeSession(workbook: ExcelScript.Workbook, data: SessionData): void {
  const sheet = workbook.getWorksheet("SESSION");
  if (!sheet) return;

  const table = sheet.getTable("SessionTable");
  if (!table) return;

  // Clear existing rows
  const body = table.getRangeBetweenHeaderAndTotal();
  if (body.getRowCount() > 0) {
    body.delete(ExcelScript.DeleteShiftDirection.up);
  }

  table.addRow(-1, [data.userID, data.username, data.role, data.managerID, data.loginTime]);
}

/**
 * Clears the session (logout).
 */
function clearSession(workbook: ExcelScript.Workbook): void {
  const sheet = workbook.getWorksheet("SESSION");
  if (!sheet) return;

  const table = sheet.getTable("SessionTable");
  if (!table) return;

  const body = table.getRangeBetweenHeaderAndTotal();
  if (body.getRowCount() > 0) {
    body.delete(ExcelScript.DeleteShiftDirection.up);
  }
}

// ── User Data Helpers ──────────────────────────────────────────────────────────
/**
 * Reads all users from Users_DB and returns them as UserRow objects.
 */
function getAllUsers(workbook: ExcelScript.Workbook): UserRow[] {
  const sheet = workbook.getWorksheet("USERS_DB");
  if (!sheet) return [];

  const table = sheet.getTable("UsersTable");
  if (!table) return [];

  const rows = table.getRangeBetweenHeaderAndTotal().getValues();
  return rows.map((r) => ({
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

/**
 * Finds a user by username (case-insensitive).
 */
function findUserByUsername(users: UserRow[], username: string): UserRow | undefined {
  const lower = username.toLowerCase();
  return users.find((u) => u.username.toLowerCase() === lower && u.active);
}

/**
 * Finds a user by UserID.
 */
function findUserByID(users: UserRow[], userID: string): UserRow | undefined {
  return users.find((u) => u.userID === userID);
}

// ── Hierarchy Resolution ───────────────────────────────────────────────────────
/**
 * Recursively walks the ManagerID chain downward.
 * Returns all UserIDs that report (directly or indirectly) to the given managerID.
 */
function getSubtreeUserIDs(managerID: string, allUsers: UserRow[]): string[] {
  const result: string[] = [];
  const directReports = allUsers.filter((u) => u.managerID === managerID && u.active);
  for (const report of directReports) {
    result.push(report.userID);
    const subReports = getSubtreeUserIDs(report.userID, allUsers);
    result.push(...subReports);
  }
  return result;
}

/**
 * Returns a numeric level for the role.
 * Admin=0, Director=1, Manager=2, TeamLead=3, Employee=4
 */
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

/**
 * Returns true if the session user is an admin, or if entryOwnerID is in their subtree.
 */
function canUserEditEntry(sessionUserID: string, entryOwnerID: string, allUsers: UserRow[]): boolean {
  const sessionUser = findUserByID(allUsers, sessionUserID);
  if (!sessionUser) return false;

  // Admins can edit everything
  if (sessionUser.role === "Admin") return true;

  // Users can always edit their own entries
  if (sessionUserID === entryOwnerID) return true;

  // Managers can edit entries of their subtree
  const subtree = getSubtreeUserIDs(sessionUserID, allUsers);
  return subtree.includes(entryOwnerID);
}

// ── Date Helpers ───────────────────────────────────────────────────────────────
/**
 * Returns the Monday of the week containing the given date.
 */
function getWeekStart(date: Date): Date {
  const d = new Date(date);
  const day = d.getDay();
  const diff = d.getDate() - day + (day === 0 ? -6 : 1);
  d.setDate(diff);
  d.setHours(0, 0, 0, 0);
  return d;
}

/**
 * Returns the Friday of the week containing the given date.
 */
function getWeekEnd(date: Date): Date {
  const start = getWeekStart(date);
  const end = new Date(start);
  end.setDate(start.getDate() + 4);
  return end;
}

/**
 * Formats a Date as YYYY-MM-DD string.
 */
function formatDate(date: Date): string {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

/**
 * Parses a date value from Excel (could be serial number or string).
 */
function parseExcelDate(value: string | number | boolean): Date {
  if (typeof value === "number") {
    // Excel serial date: days since 1899-12-30
    const epoch = new Date(1899, 11, 30);
    epoch.setDate(epoch.getDate() + value);
    return epoch;
  }
  return new Date(String(value));
}

/**
 * Generates a unique entry ID based on timestamp.
 */
function generateEntryID(): string {
  const now = new Date();
  const ts = now.getTime().toString(36);
  const rand = Math.random().toString(36).substring(2, 6);
  return `TE-${ts}-${rand}`.toUpperCase();
}

// ── Sheet Visibility Helpers ───────────────────────────────────────────────────
/**
 * Hides all sheets except LOGIN and README.
 */
function hideAllWorksheets(workbook: ExcelScript.Workbook): void {
  const alwaysVisible = ["LOGIN", "README"];
  const sheets = workbook.getWorksheets();
  for (const sheet of sheets) {
    const name = sheet.getName();
    if (alwaysVisible.includes(name)) {
      sheet.setVisibility(ExcelScript.SheetVisibility.visible);
    } else {
      sheet.setVisibility(ExcelScript.SheetVisibility.hidden);
    }
  }
  // Activate LOGIN sheet
  workbook.getWorksheet("LOGIN")?.activate();
}

/**
 * Shows sheets based on the user's role after login.
 */
function showSheetsForRole(workbook: ExcelScript.Workbook, role: string): void {
  const roleLevel = getRoleLevel(role);

  // Everyone gets these
  const visibleSheets = ["TIME_ENTRY", "DAILY_CHECK"];

  // Managers and above get MANAGER_VIEW
  if (roleLevel <= 3) {
    visibleSheets.push("MANAGER_VIEW");
  }

  // Managers and above get REPORTS
  if (roleLevel <= 2) {
    visibleSheets.push("REPORTS");
  }

  // Admins get ADMIN sheet and DB sheets
  if (roleLevel === 0) {
    visibleSheets.push("ADMIN", "USERS_DB", "PROJECTS_DB", "ACTIVITIES_DB", "TEAMS_DB");
  }

  // README always visible
  visibleSheets.push("README");

  for (const name of visibleSheets) {
    const sheet = workbook.getWorksheet(name);
    if (sheet) {
      sheet.setVisibility(ExcelScript.SheetVisibility.visible);
    }
  }

  // Hide LOGIN sheet after successful login
  const loginSheet = workbook.getWorksheet("LOGIN");
  if (loginSheet) {
    loginSheet.setVisibility(ExcelScript.SheetVisibility.hidden);
  }

  // Activate TIME_ENTRY sheet
  workbook.getWorksheet("TIME_ENTRY")?.activate();
}

// ── Table Reading Helpers ──────────────────────────────────────────────────────
/**
 * Gets all time entries from the TIME_ENTRY sheet.
 */
function getAllTimeEntries(workbook: ExcelScript.Workbook): TimeEntryRow[] {
  const sheet = workbook.getWorksheet("TIME_ENTRY");
  if (!sheet) return [];

  const table = sheet.getTable("TimeEntryTable");
  if (!table) return [];

  const rows = table.getRangeBetweenHeaderAndTotal().getValues();
  return rows.map((r) => ({
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

/**
 * Gets team name by TeamID.
 */
function getTeamName(workbook: ExcelScript.Workbook, teamID: string): string {
  const sheet = workbook.getWorksheet("TEAMS_DB");
  if (!sheet) return "";

  const table = sheet.getTable("TeamsTable");
  if (!table) return "";

  const rows = table.getRangeBetweenHeaderAndTotal().getValues();
  for (const r of rows) {
    if (String(r[0]) === teamID) return String(r[1]);
  }
  return "";
}
