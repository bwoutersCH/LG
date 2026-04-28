/**
 * Login.ts — Handles user authentication and session management.
 * Self-contained Office Script. Entry point: main()
 *
 * Contains two modes — set the ACTION cell (B8) to "login" or "logout"
 * before running, or call via button assignment.
 *
 * For button use, create separate scripts that call runLogin / runLogout directly.
 */

// ── Types ─────────────────────────────────────────────────────────────────────
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

// ── Entry Point ───────────────────────────────────────────────────────────────
function main(workbook: ExcelScript.Workbook): void {
  runLogin(workbook);
}

// ── Login ─────────────────────────────────────────────────────────────────────
function runLogin(workbook: ExcelScript.Workbook): void {
  const loginSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("LOGIN");
  if (!loginSheet) return;

  const msgCell: ExcelScript.Range = loginSheet.getRange("B6");
  msgCell.setValue("");

  const username: string = String(loginSheet.getRange("B3").getValue()).trim();
  const password: string = String(loginSheet.getRange("B4").getValue()).trim();

  if (!username || !password) {
    msgCell.setValue("Please enter both username and password.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }

  const allUsers: UserRow[] = getAllUsers(workbook);
  const user: UserRow | undefined = findUserByUsername(allUsers, username);

  if (!user) {
    msgCell.setValue("Invalid username or password.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    loginSheet.getRange("B4").setValue("");
    return;
  }

  const inputHash: string = simpleHash(password);
  // Normalize both hashes to string for comparison (Excel may store hash as number)
  const storedHash: string = String(user.passwordHash).trim();
  const computedHash: string = String(inputHash).trim();
  if (computedHash !== storedHash) {
    msgCell.setValue(`Invalid username or password. Debug: stored=[${storedHash}] computed=[${computedHash}]`);
    msgCell.getFormat().getFont().setColor("#D9415C");
    loginSheet.getRange("B4").setValue("");
    return;
  }

  if (!user.active) {
    msgCell.setValue("This account has been deactivated. Contact an administrator.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    loginSheet.getRange("B4").setValue("");
    return;
  }

  const sessionData: SessionData = {
    userID: user.userID,
    username: user.username,
    role: user.role,
    managerID: user.managerID,
    loginTime: new Date().toISOString(),
  };
  writeSession(workbook, sessionData);

  loginSheet.getRange("B4").setValue("");

  // First-login check: if the stored hash matches the default password hash,
  // require the user to set a new one before granting access to other sheets.
  if (storedHash === simpleHash("Pass1234")) {
    msgCell.setValue("First login: type a new password in the Password field and click 'Run changePassword'.");
    msgCell.getFormat().getFont().setColor("#FDC400");
    return;
  }

  msgCell.setValue("");
  showSheetsForRole(workbook, user.role);

  const timeSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("TIME_ENTRY");
  if (timeSheet) {
    timeSheet.getRange("B2").setValue(formatDate(new Date()));
  }

  refreshDailyCheck(workbook, sessionData);
}

// ── Change Password ──────────────────────────────────────────────────────────
function changePassword(workbook: ExcelScript.Workbook): void {
  const loginSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("LOGIN");
  if (!loginSheet) return;
  const msgCell: ExcelScript.Range = loginSheet.getRange("B6");

  const session: SessionData | null = getSession(workbook);
  if (!session) {
    msgCell.setValue("Please log in first, then type a new password and run Change Password.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }

  const newPassword: string = String(loginSheet.getRange("B4").getValue()).trim();
  if (!newPassword) {
    msgCell.setValue("Type a new password in the Password field, then run Change Password.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }
  if (newPassword === "Pass1234") {
    msgCell.setValue("New password cannot be the default. Choose a different password.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }
  if (newPassword.length < 6) {
    msgCell.setValue("New password must be at least 6 characters.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }

  const usersSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("USERS_DB");
  if (!usersSheet) {
    msgCell.setValue("USERS_DB not found.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }
  const table: ExcelScript.Table | undefined = usersSheet.getTable("UsersTable");
  if (!table) {
    msgCell.setValue("UsersTable not found.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }

  const body: ExcelScript.Range = table.getRangeBetweenHeaderAndTotal();
  const rows: (string | number | boolean)[][] = body.getValues();
  let userRowIdx: number = -1;
  for (let i: number = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === session.userID) { userRowIdx = i; break; }
  }
  if (userRowIdx === -1) {
    msgCell.setValue("Session user not found in USERS_DB.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }
  body.getCell(userRowIdx, 3).setValue(simpleHash(newPassword));

  loginSheet.getRange("B4").setValue("");
  msgCell.setValue(`Password updated. Welcome ${session.username}.`);
  msgCell.getFormat().getFont().setColor("#239A98");

  showSheetsForRole(workbook, session.role);
  const timeSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("TIME_ENTRY");
  if (timeSheet) {
    timeSheet.getRange("B2").setValue(formatDate(new Date()));
  }
  refreshDailyCheck(workbook, session);
}

// ── Logout ────────────────────────────────────────────────────────────────────
function runLogout(workbook: ExcelScript.Workbook): void {
  clearSession(workbook);
  hideAllWorksheets(workbook);

  const loginSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("LOGIN");
  if (loginSheet) {
    loginSheet.getRange("B3").setValue("");
    loginSheet.getRange("B4").setValue("");
    loginSheet.getRange("B6").setValue("");
    loginSheet.activate();
  }
}

// ── Daily Check Refresh ───────────────────────────────────────────────────────
function refreshDailyCheck(workbook: ExcelScript.Workbook, session: SessionData): void {
  const checkSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("DAILY_CHECK");
  if (!checkSheet) return;

  const today: Date = new Date();
  const weekStart: Date = getWeekStart(today);
  const entries: TimeEntryRow[] = getAllTimeEntries(workbook);

  const days: string[] = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];
  const dayHours: number[] = [0, 0, 0, 0, 0];

  for (const entry of entries) {
    if (entry.employeeID !== session.userID) continue;
    const entryDate: Date = parseExcelDate(entry.date);
    const entryWeekStart: Date = getWeekStart(entryDate);
    if (formatDate(entryWeekStart) !== formatDate(weekStart)) continue;

    const dayOfWeek: number = entryDate.getDay();
    if (dayOfWeek >= 1 && dayOfWeek <= 5) {
      dayHours[dayOfWeek - 1] += entry.hours;
    }
  }

  for (let i: number = 0; i < 5; i++) {
    const dateForDay: Date = new Date(weekStart);
    dateForDay.setDate(weekStart.getDate() + i);

    const row: number = 3 + i;
    checkSheet.getRange(`A${row}`).setValue(days[i]);
    checkSheet.getRange(`B${row}`).setValue(formatDate(dateForDay));
    checkSheet.getRange(`C${row}`).setValue(dayHours[i]);

    const fmt: ExcelScript.RangeFormat = checkSheet.getRange(`A${row}:C${row}`).getFormat();
    if (dayHours[i] < 8 && formatDate(dateForDay) <= formatDate(today)) {
      fmt.getFill().setColor("#FDC400");
      fmt.getFont().setColor("#575756");
    } else {
      fmt.getFill().setColor("#FFFFFF");
      fmt.getFont().setColor("#575756");
    }
  }

  const totalHours: number = dayHours.reduce((a: number, b: number): number => a + b, 0);
  checkSheet.getRange("A8").setValue("TOTAL");
  checkSheet.getRange("C8").setValue(totalHours);
  checkSheet.getRange("A8:C8").getFormat().getFont().setBold(true);

  checkSheet.getRange("A1").setValue(`Weekly Hours — ${session.username} — Week of ${formatDate(weekStart)}`);
}

// ── Shared Utilities (duplicated for self-containment) ────────────────────────
function simpleHash(input: string): string {
  const saltVal: string = "LG_TimeTrack_2026";
  const salted: string = saltVal + input + saltVal;
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

function writeSession(workbook: ExcelScript.Workbook, data: SessionData): void {
  const sheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("SESSION");
  if (!sheet) return;
  const table: ExcelScript.Table | undefined = sheet.getTable("SessionTable");
  if (!table) return;
  // Overwrite the first row instead of delete+add (delete fails inside tables)
  const body: ExcelScript.Range = table.getRangeBetweenHeaderAndTotal();
  if (body.getRowCount() > 0) {
    body.getRow(0).setValues([[data.userID, data.username, data.role, data.managerID, data.loginTime]]);
  } else {
    table.addRow(-1, [data.userID, data.username, data.role, data.managerID, data.loginTime]);
  }
}

function clearSession(workbook: ExcelScript.Workbook): void {
  const sheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("SESSION");
  if (!sheet) return;
  const table: ExcelScript.Table | undefined = sheet.getTable("SessionTable");
  if (!table) return;
  // Clear values instead of deleting rows
  const body: ExcelScript.Range = table.getRangeBetweenHeaderAndTotal();
  if (body.getRowCount() > 0) {
    body.getRow(0).setValues([["", "", "", "", ""]]);
  }
}

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

function getRoleLevel(role: string): number {
  const levels: Record<string, number> = { Admin: 0, Director: 1, Manager: 2, TeamLead: 3, Employee: 4 };
  return levels[role] ?? 4;
}

function getWeekStart(d: Date): Date {
  const result: Date = new Date(d);
  const dayNum: number = result.getDay();
  const diff: number = result.getDate() - dayNum + (dayNum === 0 ? -6 : 1);
  result.setDate(diff);
  result.setHours(0, 0, 0, 0);
  return result;
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

function hideAllWorksheets(workbook: ExcelScript.Workbook): void {
  const alwaysVisible: string[] = ["LOGIN", "README"];
  const sheets: ExcelScript.Worksheet[] = workbook.getWorksheets();
  for (const s of sheets) {
    const nm: string = s.getName();
    if (alwaysVisible.includes(nm)) {
      s.setVisibility(ExcelScript.SheetVisibility.visible);
    } else {
      s.setVisibility(ExcelScript.SheetVisibility.hidden);
    }
  }
  workbook.getWorksheet("LOGIN")?.activate();
}

function showSheetsForRole(workbook: ExcelScript.Workbook, role: string): void {
  const roleLevel: number = getRoleLevel(role);
  const visibleSheets: string[] = ["TIME_ENTRY", "DAILY_CHECK"];
  if (roleLevel <= 3) visibleSheets.push("MANAGER_VIEW");
  if (roleLevel <= 2) visibleSheets.push("REPORTS");
  if (roleLevel === 0) visibleSheets.push("ADMIN", "USERS_DB", "PROJECTS_DB", "ACTIVITIES_DB", "TEAMS_DB");
  visibleSheets.push("README");

  for (const name of visibleSheets) {
    const ws: ExcelScript.Worksheet | undefined = workbook.getWorksheet(name);
    if (ws) ws.setVisibility(ExcelScript.SheetVisibility.visible);
  }
  const loginSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("LOGIN");
  // Keep LOGIN visible so users can navigate back to log out
  workbook.getWorksheet("TIME_ENTRY")?.activate();
}
