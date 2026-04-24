/**
 * FilterManagerView — Auto-generated button script.
 * Entry point: main() calls filterManagerView().
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
  filterManagerView(workbook);
}

// ── Filter Manager View ───────────────────────────────────────────────────────
function filterManagerView(workbook: ExcelScript.Workbook): void {
  const session: SessionData = requireSession(workbook);
  const sheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("MANAGER_VIEW");
  if (!sheet) return;

  const roleLevel: number = getRoleLevel(session.role);
  if (roleLevel > 3) {
    sheet.getRange("A1").setValue("Access denied. Manager role or above required.");
    return;
  }

  const teamFilter: string = String(sheet.getRange("B2").getValue()).trim();
  const employeeFilter: string = String(sheet.getRange("D2").getValue()).trim();
  const dateFrom: string | number | boolean = sheet.getRange("F2").getValue();
  const dateTo: string | number | boolean = sheet.getRange("H2").getValue();

  const allUsers: UserRow[] = getAllUsers(workbook);
  let subtreeIDs: string[];

  if (session.role === "Admin") {
    subtreeIDs = allUsers.filter((u: UserRow): boolean => u.active).map((u: UserRow): string => u.userID);
  } else {
    subtreeIDs = getSubtreeUserIDs(session.userID, allUsers);
  }

  const allEntries: TimeEntryRow[] = getAllTimeEntries(workbook);
  let filtered: TimeEntryRow[] = allEntries.filter((e: TimeEntryRow): boolean => subtreeIDs.includes(e.employeeID));

  if (teamFilter && teamFilter !== "All") {
    filtered = filtered.filter((e: TimeEntryRow): boolean => e.team === teamFilter);
  }
  if (employeeFilter && employeeFilter !== "All") {
    filtered = filtered.filter((e: TimeEntryRow): boolean => e.employeeName === employeeFilter);
  }
  if (dateFrom) {
    const fromDate: Date = parseExcelDate(dateFrom);
    filtered = filtered.filter((e: TimeEntryRow): boolean => parseExcelDate(e.date) >= fromDate);
  }
  if (dateTo) {
    const toDate: Date = parseExcelDate(dateTo);
    filtered = filtered.filter((e: TimeEntryRow): boolean => parseExcelDate(e.date) <= toDate);
  }

  const outputTable: ExcelScript.Table | undefined = sheet.getTable("ManagerViewTable");
  if (!outputTable) return;

  // Clear existing table body content
  const existingCount: number = outputTable.getRowCount();
  if (existingCount > 0) {
    const body: ExcelScript.Range = outputTable.getRangeBetweenHeaderAndTotal();
    body.clear(ExcelScript.ClearApplyTo.contents);
  }

  // Write filtered data: overwrite row 0, addRow for the rest
  for (let idx: number = 0; idx < filtered.length; idx++) {
    const entry: TimeEntryRow = filtered[idx];
    const rowData: (string | number)[] = [
      entry.entryID, entry.date, entry.employeeID, entry.employeeName,
      entry.team, entry.project, entry.activity, entry.hours,
      entry.notes, entry.submittedOn, entry.lastEditedBy, entry.lastEditedOn,
    ];
    if (idx < existingCount) {
      // Overwrite existing row
      const body: ExcelScript.Range = outputTable.getRangeBetweenHeaderAndTotal();
      body.getRow(idx).setValues([rowData]);
    } else {
      // Add new row
      outputTable.addRow(-1, rowData);
    }
  }

  sheet.getRange("A4").setValue(`Showing ${filtered.length} entries for your team.`);
  sheet.getRange("A4").getFormat().getFont().setColor("#239A98");
}

// ── Correct Entry ─────────────────────────────────────────────────────────────
function correctEntry(workbook: ExcelScript.Workbook): void {
  const session: SessionData = requireSession(workbook);
  const sheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("MANAGER_VIEW");
  if (!sheet) return;

  const roleLevel: number = getRoleLevel(session.role);
  if (roleLevel > 3) {
    sheet.getRange("J3").setValue("Access denied.");
    return;
  }

  const entryID: string = String(sheet.getRange("B14").getValue()).trim();
  const correctedHours: number = Number(sheet.getRange("D14").getValue());
  const reason: string = String(sheet.getRange("F14").getValue()).trim();
  const msgCell: ExcelScript.Range = sheet.getRange("B16");

  if (!entryID) { msgCell.setValue("Please enter an Entry ID."); msgCell.getFormat().getFont().setColor("#D9415C"); return; }
  if (!Number.isInteger(correctedHours) || correctedHours < 1 || correctedHours > 16) { msgCell.setValue("Corrected hours must be 1-16."); msgCell.getFormat().getFont().setColor("#D9415C"); return; }
  if (!reason) { msgCell.setValue("Please provide a reason for the correction."); msgCell.getFormat().getFont().setColor("#D9415C"); return; }

  const timeSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("TIME_ENTRY");
  if (!timeSheet) return;
  const table: ExcelScript.Table | undefined = timeSheet.getTable("TimeEntryTable");
  if (!table) return;

  const allUsers: UserRow[] = getAllUsers(workbook);
  const rows: (string | number | boolean)[][] = table.getRangeBetweenHeaderAndTotal().getValues();

  let targetRow: number = -1;
  let ownerID: string = "";
  for (let i: number = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === entryID) { targetRow = i; ownerID = String(rows[i][2]); break; }
  }

  if (targetRow === -1) { msgCell.setValue(`Entry ${entryID} not found.`); msgCell.getFormat().getFont().setColor("#D9415C"); return; }
  if (!canUserEditEntry(session.userID, ownerID, allUsers)) { msgCell.setValue("You do not have permission to edit this entry."); msgCell.getFormat().getFont().setColor("#D9415C"); return; }

  const dataRange: ExcelScript.Range = table.getRangeBetweenHeaderAndTotal();
  const now: string = new Date().toISOString();
  dataRange.getCell(targetRow, 7).setValue(correctedHours);
  const existingNotes: string = String(rows[targetRow][8]);
  const updatedNotes: string = existingNotes ? `${existingNotes} | CORRECTED by ${session.username}: ${reason}` : `CORRECTED by ${session.username}: ${reason}`;
  dataRange.getCell(targetRow, 8).setValue(updatedNotes);
  dataRange.getCell(targetRow, 10).setValue(session.username);
  dataRange.getCell(targetRow, 11).setValue(now);

  sheet.getRange("B14").setValue("");
  sheet.getRange("D14").setValue("");
  sheet.getRange("F14").setValue("");
  msgCell.setValue(`Entry ${entryID} corrected to ${correctedHours}h by ${session.username}.`);
  msgCell.getFormat().getFont().setColor("#239A98");
}

// ── Generate Report ───────────────────────────────────────────────────────────
function generateReport(workbook: ExcelScript.Workbook): void {
  const session: SessionData = requireSession(workbook);
  const sheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("REPORTS");
  if (!sheet) return;

  const roleLevel: number = getRoleLevel(session.role);
  if (roleLevel > 2) { sheet.getRange("A8").setValue("Access denied. Manager role or above required."); return; }

  const period: string = String(sheet.getRange("B2").getValue()).trim();
  const breakdown: string = String(sheet.getRange("B3").getValue()).trim();
  const teamFilter: string = String(sheet.getRange("B4").getValue()).trim();

  const today: Date = new Date();
  let startDate: Date;
  let endDate: Date = today;

  switch (period) {
    case "This week": startDate = getWeekStart(today); endDate = getWeekEnd(today); break;
    case "This month": startDate = new Date(today.getFullYear(), today.getMonth(), 1); endDate = new Date(today.getFullYear(), today.getMonth() + 1, 0); break;
    case "This quarter": { const q: number = Math.floor(today.getMonth() / 3); startDate = new Date(today.getFullYear(), q * 3, 1); endDate = new Date(today.getFullYear(), q * 3 + 3, 0); break; }
    case "YTD": startDate = new Date(today.getFullYear(), 0, 1); break;
    case "Full year": startDate = new Date(today.getFullYear(), 0, 1); endDate = new Date(today.getFullYear(), 11, 31); break;
    default: startDate = getWeekStart(today); endDate = getWeekEnd(today);
  }

  const startStr: string = formatDate(startDate);
  const endStr: string = formatDate(endDate);

  const allUsers: UserRow[] = getAllUsers(workbook);
  let scopeUserIDs: string[];
  if (session.role === "Admin") {
    scopeUserIDs = allUsers.filter((u: UserRow): boolean => u.active).map((u: UserRow): string => u.userID);
  } else {
    scopeUserIDs = getSubtreeUserIDs(session.userID, allUsers);
    scopeUserIDs.push(session.userID);
  }

  const allEntries: TimeEntryRow[] = getAllTimeEntries(workbook);
  let entries: TimeEntryRow[] = allEntries.filter((e: TimeEntryRow): boolean => {
    if (!scopeUserIDs.includes(e.employeeID)) return false;
    const d: string = e.date;
    return d >= startStr && d <= endStr;
  });
  if (teamFilter && teamFilter !== "All") {
    entries = entries.filter((e: TimeEntryRow): boolean => e.team === teamFilter);
  }

  // Clear previous report
  const usedRange: ExcelScript.Range | undefined = sheet.getUsedRange();
  if (usedRange) {
    const lastRow: number = usedRange.getRowCount();
    if (lastRow > 7) sheet.getRange(`A8:H${lastRow + 1}`).clear(ExcelScript.ClearApplyTo.all);
  }

  let outputRow: number = 8;

  // Section 1: Total hours per dimension
  const dimMap: Map<string, number> = new Map<string, number>();
  for (const e of entries) {
    let key: string = "";
    switch (breakdown) {
      case "By Employee": key = e.employeeName; break;
      case "By Project": key = e.project; break;
      case "By Activity": key = e.activity; break;
      case "By Team": key = e.team; break;
      default: key = e.employeeName;
    }
    dimMap.set(key, (dimMap.get(key) || 0) + e.hours);
  }

  sheet.getRange(`A${outputRow}`).setValue(`Total Hours ${breakdown} (${period}: ${startStr} to ${endStr})`);
  sheet.getRange(`A${outputRow}:C${outputRow}`).getFormat().getFont().setBold(true);
  sheet.getRange(`A${outputRow}:C${outputRow}`).getFormat().getFill().setColor("#243347");
  sheet.getRange(`A${outputRow}:C${outputRow}`).getFormat().getFont().setColor("#FFFFFF");
  outputRow++;
  sheet.getRange(`A${outputRow}`).setValue("Dimension");
  sheet.getRange(`B${outputRow}`).setValue("Total Hours");
  sheet.getRange(`A${outputRow}:B${outputRow}`).getFormat().getFont().setBold(true);
  outputRow++;
  dimMap.forEach((val: number, key: string): void => {
    sheet.getRange(`A${outputRow}`).setValue(key);
    sheet.getRange(`B${outputRow}`).setValue(val);
    if (outputRow % 2 === 0) sheet.getRange(`A${outputRow}:B${outputRow}`).getFormat().getFill().setColor("#F5B0A3");
    outputRow++;
  });
  outputRow += 2;

  // Section 2: Hours by project
  const projectMap: Map<string, number> = new Map<string, number>();
  for (const e of entries) { projectMap.set(e.project, (projectMap.get(e.project) || 0) + e.hours); }
  sheet.getRange(`A${outputRow}`).setValue("Hours by Project");
  sheet.getRange(`A${outputRow}:C${outputRow}`).getFormat().getFill().setColor("#239A98");
  sheet.getRange(`A${outputRow}:C${outputRow}`).getFormat().getFont().setColor("#FFFFFF");
  sheet.getRange(`A${outputRow}:C${outputRow}`).getFormat().getFont().setBold(true);
  outputRow++;
  sheet.getRange(`A${outputRow}`).setValue("Project"); sheet.getRange(`B${outputRow}`).setValue("Hours");
  sheet.getRange(`A${outputRow}:B${outputRow}`).getFormat().getFont().setBold(true);
  outputRow++;
  projectMap.forEach((val: number, key: string): void => {
    sheet.getRange(`A${outputRow}`).setValue(key); sheet.getRange(`B${outputRow}`).setValue(val); outputRow++;
  });
  outputRow += 2;

  // Section 3: Hours by activity
  const activityMap: Map<string, number> = new Map<string, number>();
  for (const e of entries) { activityMap.set(e.activity, (activityMap.get(e.activity) || 0) + e.hours); }
  sheet.getRange(`A${outputRow}`).setValue("Hours by Activity");
  sheet.getRange(`A${outputRow}:C${outputRow}`).getFormat().getFill().setColor("#239A98");
  sheet.getRange(`A${outputRow}:C${outputRow}`).getFormat().getFont().setColor("#FFFFFF");
  sheet.getRange(`A${outputRow}:C${outputRow}`).getFormat().getFont().setBold(true);
  outputRow++;
  sheet.getRange(`A${outputRow}`).setValue("Activity"); sheet.getRange(`B${outputRow}`).setValue("Hours");
  sheet.getRange(`A${outputRow}:B${outputRow}`).getFormat().getFont().setBold(true);
  outputRow++;
  activityMap.forEach((val: number, key: string): void => {
    sheet.getRange(`A${outputRow}`).setValue(key); sheet.getRange(`B${outputRow}`).setValue(val); outputRow++;
  });
  outputRow += 2;

  // Section 4: Daily average per employee
  const empDays: Map<string, Set<string>> = new Map<string, Set<string>>();
  const empHours: Map<string, number> = new Map<string, number>();
  for (const e of entries) {
    if (!empDays.has(e.employeeName)) empDays.set(e.employeeName, new Set<string>());
    empDays.get(e.employeeName)!.add(e.date);
    empHours.set(e.employeeName, (empHours.get(e.employeeName) || 0) + e.hours);
  }
  sheet.getRange(`A${outputRow}`).setValue("Daily Average per Employee");
  sheet.getRange(`A${outputRow}:D${outputRow}`).getFormat().getFill().setColor("#239A98");
  sheet.getRange(`A${outputRow}:D${outputRow}`).getFormat().getFont().setColor("#FFFFFF");
  sheet.getRange(`A${outputRow}:D${outputRow}`).getFormat().getFont().setBold(true);
  outputRow++;
  sheet.getRange(`A${outputRow}`).setValue("Employee"); sheet.getRange(`B${outputRow}`).setValue("Total Hours");
  sheet.getRange(`C${outputRow}`).setValue("Days Logged"); sheet.getRange(`D${outputRow}`).setValue("Daily Average");
  sheet.getRange(`A${outputRow}:D${outputRow}`).getFormat().getFont().setBold(true);
  outputRow++;
  empHours.forEach((total: number, name: string): void => {
    const numDays: number = empDays.get(name)?.size || 1;
    const avg: number = Math.round((total / numDays) * 10) / 10;
    sheet.getRange(`A${outputRow}`).setValue(name); sheet.getRange(`B${outputRow}`).setValue(total);
    sheet.getRange(`C${outputRow}`).setValue(numDays); sheet.getRange(`D${outputRow}`).setValue(avg);
    outputRow++;
  });
  outputRow += 2;

  // Section 5: Under 8h day flags
  sheet.getRange(`A${outputRow}`).setValue("Under 8h Day Flags");
  sheet.getRange(`A${outputRow}:C${outputRow}`).getFormat().getFill().setColor("#D9415C");
  sheet.getRange(`A${outputRow}:C${outputRow}`).getFormat().getFont().setColor("#FFFFFF");
  sheet.getRange(`A${outputRow}:C${outputRow}`).getFormat().getFont().setBold(true);
  outputRow++;
  sheet.getRange(`A${outputRow}`).setValue("Employee"); sheet.getRange(`B${outputRow}`).setValue("Days Under 8h");
  sheet.getRange(`A${outputRow}:B${outputRow}`).getFormat().getFont().setBold(true);
  outputRow++;
  const empDayHours: Map<string, Map<string, number>> = new Map<string, Map<string, number>>();
  for (const e of entries) {
    if (!empDayHours.has(e.employeeName)) empDayHours.set(e.employeeName, new Map<string, number>());
    const dayMap: Map<string, number> = empDayHours.get(e.employeeName)!;
    dayMap.set(e.date, (dayMap.get(e.date) || 0) + e.hours);
  }
  empDayHours.forEach((dayMap: Map<string, number>, name: string): void => {
    let underDays: number = 0;
    dayMap.forEach((hrs: number): void => { if (hrs < 8) underDays++; });
    sheet.getRange(`A${outputRow}`).setValue(name); sheet.getRange(`B${outputRow}`).setValue(underDays);
    if (underDays > 0) sheet.getRange(`A${outputRow}:B${outputRow}`).getFormat().getFill().setColor("#FDC400");
    outputRow++;
  });
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
  return { userID: String(rows[0][0]), username: String(rows[0][1]), role: String(rows[0][2]), managerID: String(rows[0][3]), loginTime: String(rows[0][4]) };
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
