/**
 * TimeEntry.ts — Time entry creation, validation, filtering, weekly checks.
 * Self-contained Office Script. Entry point: main() runs addEntry().
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
  addEntry(workbook);
}

// ── Add Entry ─────────────────────────────────────────────────────────────────
function addEntry(workbook: ExcelScript.Workbook): void {
  const session: SessionData = requireSession(workbook);
  const sheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("TIME_ENTRY");
  if (!sheet) return;

  const msgCell: ExcelScript.Range = sheet.getRange("D7");
  msgCell.setValue("");

  const dateValue: string | number | boolean = sheet.getRange("B2").getValue();
  const project: string = String(sheet.getRange("B3").getValue()).trim();
  const activity: string = String(sheet.getRange("B4").getValue()).trim();
  const hoursValue: string | number | boolean = sheet.getRange("B5").getValue();
  const notes: string = String(sheet.getRange("B6").getValue()).trim();

  if (!dateValue || !project || !hoursValue) {
    msgCell.setValue("Please fill in Date, Project, and Hours.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }

  const entryDate: Date = parseExcelDate(dateValue);
  const today: Date = new Date();
  today.setHours(0, 0, 0, 0);

  const diffDays: number = Math.floor((entryDate.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
  if (diffDays < -30 || diffDays > 30) {
    msgCell.setValue("Date must be within 30 days of today.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }

  const hours: number = Number(hoursValue);
  if (!Number.isInteger(hours) || hours < 1 || hours > 16) {
    msgCell.setValue("Hours must be a whole number between 1 and 16.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }

  const allEntries: TimeEntryRow[] = getAllTimeEntries(workbook);
  const dateStr: string = formatDate(entryDate);
  const duplicate: TimeEntryRow | undefined = allEntries.find(
    (e: TimeEntryRow): boolean =>
      e.employeeID === session.userID &&
      e.date === dateStr &&
      e.project === project &&
      e.activity === activity
  );

  if (duplicate) {
    msgCell.setValue(`Duplicate: entry for ${project} / ${activity} on ${dateStr} already exists (${duplicate.entryID}).`);
    msgCell.getFormat().getFont().setColor("#FDC400");
  }

  const allUsers: UserRow[] = getAllUsers(workbook);
  const currentUser: UserRow | undefined = findUserByID(allUsers, session.userID);
  if (!currentUser) {
    msgCell.setValue("User not found in database.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }

  const teamName: string = getTeamName(workbook, currentUser.teamID);
  const now: string = new Date().toISOString();
  const entryID: string = generateEntryID();

  const table: ExcelScript.Table | undefined = sheet.getTable("TimeEntryTable");
  if (!table) {
    msgCell.setValue("TimeEntryTable not found.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }

  table.addRow(-1, [
    entryID, dateStr, session.userID, currentUser.fullName, teamName,
    project, activity, hours, notes, now, session.username, now,
  ]);

  // Excel auto-converts ISO date strings to serial numbers; force text so
  // the Date column stays comparable with existing string-formatted entries.
  const newRowIdx: number = table.getRowCount() - 1;
  const dateCell: ExcelScript.Range = table.getRangeBetweenHeaderAndTotal().getCell(newRowIdx, 1);
  dateCell.setNumberFormat("@");
  dateCell.setValue(dateStr);

  sheet.getRange("B5").setValue("");
  sheet.getRange("B6").setValue("");

  msgCell.setValue(`Entry ${entryID} added successfully (${hours}h on ${dateStr}).`);
  msgCell.getFormat().getFont().setColor("#239A98");

  refreshDailyCheck(workbook, session);
}

// ── Filter My Entries ─────────────────────────────────────────────────────────
function filterMyEntries(workbook: ExcelScript.Workbook): void {
  const session: SessionData = requireSession(workbook);
  const sheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("TIME_ENTRY");
  if (!sheet) return;

  const table: ExcelScript.Table | undefined = sheet.getTable("TimeEntryTable");
  if (!table) return;

  const today: Date = new Date();
  const weekStart: Date = getWeekStart(today);
  const weekEnd: Date = getWeekEnd(today);
  const weekStartStr: string = formatDate(weekStart);
  const weekEndStr: string = formatDate(weekEnd);

  table.getAutoFilter().clearCriteria();
  table.getAutoFilter().apply(
    table.getColumn("EmployeeID").getRange(),
    2,
    { filterOn: ExcelScript.FilterOn.values, values: [session.userID] }
  );

  const msgCell: ExcelScript.Range = sheet.getRange("D7");
  msgCell.setValue(`Showing entries for ${session.username}, week ${weekStartStr} to ${weekEndStr}`);
  msgCell.getFormat().getFont().setColor("#239A98");
}

// ── Clear Filters ─────────────────────────────────────────────────────────────
function clearFilters(workbook: ExcelScript.Workbook): void {
  requireSession(workbook);
  const sheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("TIME_ENTRY");
  if (!sheet) return;
  const table: ExcelScript.Table | undefined = sheet.getTable("TimeEntryTable");
  if (!table) return;
  table.getAutoFilter().clearCriteria();
  const msgCell: ExcelScript.Range = sheet.getRange("D7");
  msgCell.setValue("Filters cleared.");
  msgCell.getFormat().getFont().setColor("#575756");
}

// ── Validate Week ─────────────────────────────────────────────────────────────
function validateWeek(workbook: ExcelScript.Workbook): void {
  const session: SessionData = requireSession(workbook);
  const checkSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("DAILY_CHECK");
  if (!checkSheet) return;

  const today: Date = new Date();
  const weekStart: Date = getWeekStart(today);
  const entries: TimeEntryRow[] = getAllTimeEntries(workbook);

  const dayNames: string[] = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];
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

  const issues: string[] = [];
  for (let i: number = 0; i < 5; i++) {
    const dateForDay: Date = new Date(weekStart);
    dateForDay.setDate(weekStart.getDate() + i);
    if (dateForDay <= today) {
      if (dayHours[i] < 8) {
        issues.push(`${dayNames[i]} (${formatDate(dateForDay)}): ${dayHours[i]}h - needs ${8 - dayHours[i]}h more`);
      }
    }
  }

  const msgCell: ExcelScript.Range = checkSheet.getRange("A10");
  if (issues.length === 0) {
    msgCell.setValue("Week is complete! All days have 8+ hours logged.");
    msgCell.getFormat().getFont().setColor("#239A98");
    msgCell.getFormat().getFont().setBold(true);
  } else {
    const msg: string = "Incomplete days:\n" + issues.join("\n");
    msgCell.setValue(msg);
    msgCell.getFormat().getFont().setColor("#D9415C");
    msgCell.getFormat().getFont().setBold(true);
  }

  refreshDailyCheck(workbook, session);
}

// ── Delete Entry ──────────────────────────────────────────────────────────────
function deleteEntry(workbook: ExcelScript.Workbook): void {
  const session: SessionData = requireSession(workbook);
  const sheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("TIME_ENTRY");
  if (!sheet) return;

  const entryID: string = String(sheet.getRange("F2").getValue()).trim();
  const msgCell: ExcelScript.Range = sheet.getRange("D7");

  if (!entryID) {
    msgCell.setValue("Please enter an Entry ID in cell F2 to delete.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }

  const table: ExcelScript.Table | undefined = sheet.getTable("TimeEntryTable");
  if (!table) return;

  const allUsers: UserRow[] = getAllUsers(workbook);
  const rows: (string | number | boolean)[][] = table.getRangeBetweenHeaderAndTotal().getValues();

  let rowIndex: number = -1;
  for (let i: number = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === entryID) {
      rowIndex = i;
      break;
    }
  }

  if (rowIndex === -1) {
    msgCell.setValue(`Entry ${entryID} not found.`);
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }

  const ownerID: string = String(rows[rowIndex][2]);
  if (!canUserEditEntry(session.userID, ownerID, allUsers)) {
    msgCell.setValue("You do not have permission to delete this entry.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }

  table.deleteRowsAt(rowIndex, 1);
  sheet.getRange("F2").setValue("");
  msgCell.setValue(`Entry ${entryID} deleted.`);
  msgCell.getFormat().getFont().setColor("#239A98");

  refreshDailyCheck(workbook, session);
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
  checkSheet.getRange("A1").setValue(`Weekly Hours - ${session.username} - Week of ${formatDate(weekStart)}`);
}

// ══════════════════════════════════════════════════════════════════════════════
// SHARED UTILITIES (self-contained copies)
// ══════════════════════════════════════════════════════════════════════════════

function getSession(workbook: ExcelScript.Workbook): SessionData | null {
  const sheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("SESSION");
  if (!sheet) return null;
  const table: ExcelScript.Table | undefined = sheet.getTable("SessionTable");
  if (!table) return null;
  const rows: (string | number | boolean)[][] = table.getRangeBetweenHeaderAndTotal().getValues();
  if (rows.length === 0 || !rows[0][0]) return null;
  return {
    userID: String(rows[0][0]), username: String(rows[0][1]),
    role: String(rows[0][2]), managerID: String(rows[0][3]), loginTime: String(rows[0][4]),
  };
}

function requireSession(workbook: ExcelScript.Workbook): SessionData {
  const session: SessionData | null = getSession(workbook);
  if (!session) throw new Error("No active session. Please log in first.");
  return session;
}

function getAllUsers(workbook: ExcelScript.Workbook): UserRow[] {
  const sheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("USERS_DB");
  if (!sheet) return [];
  const table: ExcelScript.Table | undefined = sheet.getTable("UsersTable");
  if (!table) return [];
  const rows: (string | number | boolean)[][] = table.getRangeBetweenHeaderAndTotal().getValues();
  return rows.map((r: (string | number | boolean)[]): UserRow => ({
    userID: String(r[0]), fullName: String(r[1]), username: String(r[2]),
    passwordHash: String(r[3]), role: String(r[4]), managerID: String(r[5]),
    teamID: String(r[6]), active: Boolean(r[7]),
  }));
}

function findUserByID(users: UserRow[], userID: string): UserRow | undefined {
  return users.find((u: UserRow): boolean => u.userID === userID);
}

function getAllTimeEntries(workbook: ExcelScript.Workbook): TimeEntryRow[] {
  const sheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("TIME_ENTRY");
  if (!sheet) return [];
  const table: ExcelScript.Table | undefined = sheet.getTable("TimeEntryTable");
  if (!table) return [];
  const rows: (string | number | boolean)[][] = table.getRangeBetweenHeaderAndTotal().getValues();
  return rows.map((r: (string | number | boolean)[]): TimeEntryRow => ({
    entryID: String(r[0]), date: normalizeDateValue(r[1]), employeeID: String(r[2]),
    employeeName: String(r[3]), team: String(r[4]), project: String(r[5]),
    activity: String(r[6]), hours: Number(r[7]), notes: String(r[8]),
    submittedOn: String(r[9]), lastEditedBy: String(r[10]), lastEditedOn: String(r[11]),
  }));
}

function normalizeDateValue(value: string | number | boolean): string {
  if (typeof value === "number") return formatDate(parseExcelDate(value));
  return String(value);
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

function canUserEditEntry(sessionUserID: string, entryOwnerID: string, allUsers: UserRow[]): boolean {
  const sessionUser: UserRow | undefined = findUserByID(allUsers, sessionUserID);
  if (!sessionUser) return false;
  if (sessionUser.role === "Admin") return true;
  if (sessionUserID === entryOwnerID) return true;
  const subtree: string[] = getSubtreeUserIDs(sessionUserID, allUsers);
  return subtree.includes(entryOwnerID);
}

function getWeekStart(d: Date): Date {
  const result: Date = new Date(d);
  const dayNum: number = result.getDay();
  const diff: number = result.getDate() - dayNum + (dayNum === 0 ? -6 : 1);
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
