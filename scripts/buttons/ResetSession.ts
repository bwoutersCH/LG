/**
 * ResetSession — Auto-generated button script.
 * Entry point: main() calls resetSession().
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
  resetSession(workbook);
}

// ── Save User ─────────────────────────────────────────────────────────────────
function saveUser(workbook: ExcelScript.Workbook): void {
  const session: SessionData = requireSession(workbook);
  if (session.role !== "Admin") {
    const as1: ExcelScript.Worksheet | undefined = workbook.getWorksheet("ADMIN");
    if (as1) { as1.getRange("F3").setValue("Admin access required."); as1.getRange("F3").getFormat().getFont().setColor("#D9415C"); }
    return;
  }

  const adminSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("ADMIN");
  if (!adminSheet) return;

  const msgCell: ExcelScript.Range = adminSheet.getRange("F3");
  msgCell.setValue("");

  const userID: string = String(adminSheet.getRange("B3").getValue()).trim();
  const fullName: string = String(adminSheet.getRange("B4").getValue()).trim();
  const username: string = String(adminSheet.getRange("B5").getValue()).trim();
  const password: string = String(adminSheet.getRange("B6").getValue()).trim();
  const role: string = String(adminSheet.getRange("B7").getValue()).trim();
  const managerID: string = String(adminSheet.getRange("B8").getValue()).trim();
  const teamID: string = String(adminSheet.getRange("B9").getValue()).trim();
  const activeStr: string = String(adminSheet.getRange("B10").getValue()).trim();

  if (!fullName || !username || !role) { msgCell.setValue("FullName, Username, and Role are required."); msgCell.getFormat().getFont().setColor("#D9415C"); return; }

  const validRoles: string[] = ["Admin", "Director", "Manager", "TeamLead", "Employee"];
  if (!validRoles.includes(role)) { msgCell.setValue(`Invalid role. Must be one of: ${validRoles.join(", ")}`); msgCell.getFormat().getFont().setColor("#D9415C"); return; }

  const usersSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("USERS_DB");
  if (!usersSheet) return;
  const table: ExcelScript.Table | undefined = usersSheet.getTable("UsersTable");
  if (!table) return;

  const rows: (string | number | boolean)[][] = table.getRangeBetweenHeaderAndTotal().getValues();
  const active: boolean = activeStr.toUpperCase() !== "FALSE";

  let existingRow: number = -1;
  if (userID) {
    for (let i: number = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === userID) { existingRow = i; break; }
    }
  }

  if (existingRow >= 0) {
    const dataRange: ExcelScript.Range = table.getRangeBetweenHeaderAndTotal();
    dataRange.getCell(existingRow, 1).setValue(fullName);
    dataRange.getCell(existingRow, 2).setValue(username);
    if (password) dataRange.getCell(existingRow, 3).setValue(simpleHash(password));
    dataRange.getCell(existingRow, 4).setValue(role);
    dataRange.getCell(existingRow, 5).setValue(managerID);
    dataRange.getCell(existingRow, 6).setValue(teamID);
    dataRange.getCell(existingRow, 7).setValue(active);
    msgCell.setValue(`User ${userID} (${fullName}) updated.`);
    msgCell.getFormat().getFont().setColor("#239A98");
  } else {
    const maxID: number = rows.reduce((max: number, r: (string | number | boolean)[]): number => {
      const num: number = parseInt(String(r[0]).replace("U", ""), 10);
      return num > max ? num : max;
    }, 0);
    const newID: string = `U${String(maxID + 1).padStart(3, "0")}`;
    if (!password) { msgCell.setValue("Password is required for new users."); msgCell.getFormat().getFont().setColor("#D9415C"); return; }
    table.addRow(-1, [newID, fullName, username, simpleHash(password), role, managerID, teamID, active]);
    msgCell.setValue(`User ${newID} (${fullName}) created.`);
    msgCell.getFormat().getFont().setColor("#239A98");
  }
  adminSheet.getRange("B3:B10").setValue("");
}

// ── Save Project ──────────────────────────────────────────────────────────────
function saveProject(workbook: ExcelScript.Workbook): void {
  const session: SessionData = requireSession(workbook);
  if (session.role !== "Admin") return;
  const adminSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("ADMIN");
  if (!adminSheet) return;
  const msgCell: ExcelScript.Range = adminSheet.getRange("F13");
  msgCell.setValue("");

  const projectID: string = String(adminSheet.getRange("B13").getValue()).trim();
  const projectName: string = String(adminSheet.getRange("B14").getValue()).trim();
  const teamID: string = String(adminSheet.getRange("B15").getValue()).trim();
  const activeStr: string = String(adminSheet.getRange("B16").getValue()).trim();

  if (!projectName) { msgCell.setValue("Project name is required."); msgCell.getFormat().getFont().setColor("#D9415C"); return; }

  const projSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("PROJECTS_DB");
  if (!projSheet) return;
  const table: ExcelScript.Table | undefined = projSheet.getTable("ProjectsTable");
  if (!table) return;
  const rows: (string | number | boolean)[][] = table.getRangeBetweenHeaderAndTotal().getValues();
  const active: boolean = activeStr.toUpperCase() !== "FALSE";

  let existingRow: number = -1;
  if (projectID) { for (let i: number = 0; i < rows.length; i++) { if (String(rows[i][0]) === projectID) { existingRow = i; break; } } }

  if (existingRow >= 0) {
    const dataRange: ExcelScript.Range = table.getRangeBetweenHeaderAndTotal();
    dataRange.getCell(existingRow, 1).setValue(projectName);
    dataRange.getCell(existingRow, 2).setValue(teamID);
    dataRange.getCell(existingRow, 3).setValue(active);
    msgCell.setValue(`Project ${projectID} updated.`); msgCell.getFormat().getFont().setColor("#239A98");
  } else {
    const maxID: number = rows.reduce((max: number, r: (string | number | boolean)[]): number => { const num: number = parseInt(String(r[0]).replace("P", ""), 10); return num > max ? num : max; }, 0);
    const newID: string = `P${String(maxID + 1).padStart(3, "0")}`;
    table.addRow(-1, [newID, projectName, teamID, active]);
    msgCell.setValue(`Project ${newID} created.`); msgCell.getFormat().getFont().setColor("#239A98");
  }
  adminSheet.getRange("B13:B16").setValue("");
}

// ── Save Activity ─────────────────────────────────────────────────────────────
function saveActivity(workbook: ExcelScript.Workbook): void {
  const session: SessionData = requireSession(workbook);
  if (session.role !== "Admin") return;
  const adminSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("ADMIN");
  if (!adminSheet) return;
  const msgCell: ExcelScript.Range = adminSheet.getRange("F18");
  msgCell.setValue("");

  const activityID: string = String(adminSheet.getRange("B18").getValue()).trim();
  const activityName: string = String(adminSheet.getRange("B19").getValue()).trim();
  const activeStr: string = String(adminSheet.getRange("B20").getValue()).trim();

  if (!activityName) { msgCell.setValue("Activity name is required."); msgCell.getFormat().getFont().setColor("#D9415C"); return; }

  const actSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("ACTIVITIES_DB");
  if (!actSheet) return;
  const table: ExcelScript.Table | undefined = actSheet.getTable("ActivitiesTable");
  if (!table) return;
  const rows: (string | number | boolean)[][] = table.getRangeBetweenHeaderAndTotal().getValues();
  const active: boolean = activeStr.toUpperCase() !== "FALSE";

  let existingRow: number = -1;
  if (activityID) { for (let i: number = 0; i < rows.length; i++) { if (String(rows[i][0]) === activityID) { existingRow = i; break; } } }

  if (existingRow >= 0) {
    const dataRange: ExcelScript.Range = table.getRangeBetweenHeaderAndTotal();
    dataRange.getCell(existingRow, 1).setValue(activityName);
    dataRange.getCell(existingRow, 2).setValue(active);
    msgCell.setValue(`Activity ${activityID} updated.`); msgCell.getFormat().getFont().setColor("#239A98");
  } else {
    const maxID: number = rows.reduce((max: number, r: (string | number | boolean)[]): number => { const num: number = parseInt(String(r[0]).replace("A", ""), 10); return num > max ? num : max; }, 0);
    const newID: string = `A${String(maxID + 1).padStart(3, "0")}`;
    table.addRow(-1, [newID, activityName, active]);
    msgCell.setValue(`Activity ${newID} created.`); msgCell.getFormat().getFont().setColor("#239A98");
  }
  adminSheet.getRange("B18:B20").setValue("");
}

// ── Reset Session ─────────────────────────────────────────────────────────────
function resetSession(workbook: ExcelScript.Workbook): void {
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

// ── Export All Data ───────────────────────────────────────────────────────────
function exportAllData(workbook: ExcelScript.Workbook): void {
  const session: SessionData = requireSession(workbook);
  if (session.role !== "Admin") return;

  const existing: ExcelScript.Worksheet | undefined = workbook.getWorksheet("EXPORT");
  if (existing) existing.delete();

  const exportSheet: ExcelScript.Worksheet = workbook.addWorksheet("EXPORT");
  const headers: string[] = ["EntryID", "Date", "EmployeeID", "EmployeeName", "Team", "Project", "Activity", "Hours", "Notes", "SubmittedOn", "LastEditedBy", "LastEditedOn"];

  for (let i: number = 0; i < headers.length; i++) {
    const cell: ExcelScript.Range = exportSheet.getRange(`${String.fromCharCode(65 + i)}1`);
    cell.setValue(headers[i]);
    cell.getFormat().getFont().setBold(true);
    cell.getFormat().getFont().setColor("#FFFFFF");
    cell.getFormat().getFill().setColor("#243347");
  }

  const entries: TimeEntryRow[] = getAllTimeEntries(workbook);
  for (let r: number = 0; r < entries.length; r++) {
    const e: TimeEntryRow = entries[r];
    const row: number = r + 2;
    exportSheet.getRange(`A${row}`).setValue(e.entryID);
    exportSheet.getRange(`B${row}`).setValue(e.date);
    exportSheet.getRange(`C${row}`).setValue(e.employeeID);
    exportSheet.getRange(`D${row}`).setValue(e.employeeName);
    exportSheet.getRange(`E${row}`).setValue(e.team);
    exportSheet.getRange(`F${row}`).setValue(e.project);
    exportSheet.getRange(`G${row}`).setValue(e.activity);
    exportSheet.getRange(`H${row}`).setValue(e.hours);
    exportSheet.getRange(`I${row}`).setValue(e.notes);
    exportSheet.getRange(`J${row}`).setValue(e.submittedOn);
    exportSheet.getRange(`K${row}`).setValue(e.lastEditedBy);
    exportSheet.getRange(`L${row}`).setValue(e.lastEditedOn);
  }

  exportSheet.activate();
  const adminSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("ADMIN");
  if (adminSheet) {
    adminSheet.getRange("F22").setValue(`Exported ${entries.length} entries to EXPORT sheet.`);
    adminSheet.getRange("F22").getFormat().getFont().setColor("#239A98");
  }
}

// ══════════════════════════════════════════════════════════════════════════════
// SHARED UTILITIES (self-contained copies)
// ══════════════════════════════════════════════════════════════════════════════

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
  return { userID: String(rows[0][0]), username: String(rows[0][1]), role: String(rows[0][2]), managerID: String(rows[0][3]), loginTime: String(rows[0][4]) };
}

function requireSession(workbook: ExcelScript.Workbook): SessionData {
  const session: SessionData | null = getSession(workbook);
  if (!session) throw new Error("No active session. Please log in first.");
  return session;
}

function clearSession(workbook: ExcelScript.Workbook): void {
  const sheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("SESSION");
  if (!sheet) return;
  const table: ExcelScript.Table | undefined = sheet.getTable("SessionTable");
  if (!table) return;
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
    userID: String(r[0]), fullName: String(r[1]), username: String(r[2]),
    passwordHash: String(r[3]), role: String(r[4]), managerID: String(r[5]),
    teamID: String(r[6]), active: Boolean(r[7]),
  }));
}

function getAllTimeEntries(workbook: ExcelScript.Workbook): TimeEntryRow[] {
  const sheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("TIME_ENTRY");
  if (!sheet) return [];
  const table: ExcelScript.Table | undefined = sheet.getTable("TimeEntryTable");
  if (!table) return [];
  const rows: (string | number | boolean)[][] = table.getRangeBetweenHeaderAndTotal().getValues();
  return rows.map((r: (string | number | boolean)[]): TimeEntryRow => ({
    entryID: String(r[0]), date: String(r[1]), employeeID: String(r[2]),
    employeeName: String(r[3]), team: String(r[4]), project: String(r[5]),
    activity: String(r[6]), hours: Number(r[7]), notes: String(r[8]),
    submittedOn: String(r[9]), lastEditedBy: String(r[10]), lastEditedOn: String(r[11]),
  }));
}

function hideAllWorksheets(workbook: ExcelScript.Workbook): void {
  const alwaysVisible: string[] = ["LOGIN", "README"];
  const sheets: ExcelScript.Worksheet[] = workbook.getWorksheets();
  for (const s of sheets) {
    const nm: string = s.getName();
    if (alwaysVisible.includes(nm)) s.setVisibility(ExcelScript.SheetVisibility.visible);
    else s.setVisibility(ExcelScript.SheetVisibility.hidden);
  }
  workbook.getWorksheet("LOGIN")?.activate();
}
