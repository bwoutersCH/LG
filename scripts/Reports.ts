/**
 * Reports.ts — Handles MANAGER_VIEW filtering, entry correction, and REPORTS generation.
 * Linked to buttons on the MANAGER_VIEW and REPORTS sheets.
 */

/**
 * Filters the MANAGER_VIEW table based on the manager's subtree and selected filters.
 * Linked to the "Apply filter" button on MANAGER_VIEW sheet.
 */
function filterManagerView(workbook: ExcelScript.Workbook): void {
  const session = requireSession(workbook);
  const sheet = workbook.getWorksheet("MANAGER_VIEW");
  if (!sheet) return;

  const roleLevel = getRoleLevel(session.role);
  if (roleLevel > 3) {
    sheet.getRange("A1").setValue("⚠ Access denied. Manager role or above required.");
    return;
  }

  // Read filter values
  const teamFilter = String(sheet.getRange("B2").getValue()).trim();
  const employeeFilter = String(sheet.getRange("D2").getValue()).trim();
  const dateFrom = sheet.getRange("F2").getValue();
  const dateTo = sheet.getRange("H2").getValue();

  // Get subtree users
  const allUsers = getAllUsers(workbook);
  let subtreeIDs: string[];

  if (session.role === "Admin") {
    subtreeIDs = allUsers.filter((u) => u.active).map((u) => u.userID);
  } else {
    subtreeIDs = getSubtreeUserIDs(session.userID, allUsers);
  }

  // Get all time entries
  const allEntries = getAllTimeEntries(workbook);

  // Filter entries
  let filtered = allEntries.filter((e) => subtreeIDs.includes(e.employeeID));

  if (teamFilter && teamFilter !== "All") {
    filtered = filtered.filter((e) => e.team === teamFilter);
  }

  if (employeeFilter && employeeFilter !== "All") {
    filtered = filtered.filter((e) => e.employeeName === employeeFilter);
  }

  if (dateFrom) {
    const fromDate = parseExcelDate(dateFrom);
    filtered = filtered.filter((e) => parseExcelDate(e.date) >= fromDate);
  }

  if (dateTo) {
    const toDate = parseExcelDate(dateTo);
    filtered = filtered.filter((e) => parseExcelDate(e.date) <= toDate);
  }

  // Write filtered results to the display table
  const outputTable = sheet.getTable("ManagerViewTable");
  if (!outputTable) return;

  // Clear existing data
  const body = outputTable.getRangeBetweenHeaderAndTotal();
  if (body.getRowCount() > 0) {
    body.delete(ExcelScript.DeleteShiftDirection.up);
  }

  // Add filtered rows
  for (const entry of filtered) {
    outputTable.addRow(-1, [
      entry.entryID,
      entry.date,
      entry.employeeID,
      entry.employeeName,
      entry.team,
      entry.project,
      entry.activity,
      entry.hours,
      entry.notes,
      entry.submittedOn,
      entry.lastEditedBy,
      entry.lastEditedOn,
    ]);
  }

  sheet.getRange("A4").setValue(`Showing ${filtered.length} entries for your team.`);
  sheet.getRange("A4").getFormat().getFont().setColor("#239A98");
}

/**
 * Corrects a time entry. Manager enters EntryID, corrected hours, and reason.
 * Linked to the "Save correction" button on MANAGER_VIEW sheet.
 */
function correctEntry(workbook: ExcelScript.Workbook): void {
  const session = requireSession(workbook);
  const sheet = workbook.getWorksheet("MANAGER_VIEW");
  if (!sheet) return;

  const roleLevel = getRoleLevel(session.role);
  if (roleLevel > 3) {
    sheet.getRange("J3").setValue("⚠ Access denied.");
    return;
  }

  // Read correction inputs (located below the filter area)
  const entryID = String(sheet.getRange("B14").getValue()).trim();
  const correctedHours = Number(sheet.getRange("D14").getValue());
  const reason = String(sheet.getRange("F14").getValue()).trim();
  const msgCell = sheet.getRange("B16");

  if (!entryID) {
    msgCell.setValue("⚠ Please enter an Entry ID.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }

  if (!Number.isInteger(correctedHours) || correctedHours < 1 || correctedHours > 16) {
    msgCell.setValue("⚠ Corrected hours must be 1–16.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }

  if (!reason) {
    msgCell.setValue("⚠ Please provide a reason for the correction.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }

  // Find the entry in the main TimeEntryTable
  const timeSheet = workbook.getWorksheet("TIME_ENTRY");
  if (!timeSheet) return;

  const table = timeSheet.getTable("TimeEntryTable");
  if (!table) return;

  const allUsers = getAllUsers(workbook);
  const rows = table.getRangeBetweenHeaderAndTotal().getValues();

  let targetRow = -1;
  let ownerID = "";
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === entryID) {
      targetRow = i;
      ownerID = String(rows[i][2]);
      break;
    }
  }

  if (targetRow === -1) {
    msgCell.setValue(`⚠ Entry ${entryID} not found.`);
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }

  // Check permission
  if (!canUserEditEntry(session.userID, ownerID, allUsers)) {
    msgCell.setValue("⚠ You do not have permission to edit this entry.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }

  // Update the entry
  const dataRange = table.getRangeBetweenHeaderAndTotal();
  const now = new Date().toISOString();

  // Hours column is index 7
  dataRange.getCell(targetRow, 7).setValue(correctedHours);
  // Notes: append correction reason
  const existingNotes = String(rows[targetRow][8]);
  const updatedNotes = existingNotes
    ? `${existingNotes} | CORRECTED by ${session.username}: ${reason}`
    : `CORRECTED by ${session.username}: ${reason}`;
  dataRange.getCell(targetRow, 8).setValue(updatedNotes);
  // LastEditedBy
  dataRange.getCell(targetRow, 10).setValue(session.username);
  // LastEditedOn
  dataRange.getCell(targetRow, 11).setValue(now);

  // Clear correction inputs
  sheet.getRange("B14").setValue("");
  sheet.getRange("D14").setValue("");
  sheet.getRange("F14").setValue("");

  msgCell.setValue(`✓ Entry ${entryID} corrected to ${correctedHours}h by ${session.username}.`);
  msgCell.getFormat().getFont().setColor("#239A98");
}

/**
 * Generates a report based on selected period, breakdown, and team filters.
 * Linked to the "Generate report" button on the REPORTS sheet.
 */
function generateReport(workbook: ExcelScript.Workbook): void {
  const session = requireSession(workbook);
  const sheet = workbook.getWorksheet("REPORTS");
  if (!sheet) return;

  const roleLevel = getRoleLevel(session.role);
  if (roleLevel > 2) {
    sheet.getRange("A8").setValue("⚠ Access denied. Manager role or above required.");
    return;
  }

  // Read filter selections
  const period = String(sheet.getRange("B2").getValue()).trim();
  const breakdown = String(sheet.getRange("B3").getValue()).trim();
  const teamFilter = String(sheet.getRange("B4").getValue()).trim();

  // Determine date range based on period
  const today = new Date();
  let startDate: Date;
  let endDate: Date = today;

  switch (period) {
    case "This week":
      startDate = getWeekStart(today);
      endDate = getWeekEnd(today);
      break;
    case "This month":
      startDate = new Date(today.getFullYear(), today.getMonth(), 1);
      endDate = new Date(today.getFullYear(), today.getMonth() + 1, 0);
      break;
    case "This quarter": {
      const quarter = Math.floor(today.getMonth() / 3);
      startDate = new Date(today.getFullYear(), quarter * 3, 1);
      endDate = new Date(today.getFullYear(), quarter * 3 + 3, 0);
      break;
    }
    case "YTD":
      startDate = new Date(today.getFullYear(), 0, 1);
      break;
    case "Full year":
      startDate = new Date(today.getFullYear(), 0, 1);
      endDate = new Date(today.getFullYear(), 11, 31);
      break;
    default:
      startDate = getWeekStart(today);
      endDate = getWeekEnd(today);
  }

  const startStr = formatDate(startDate);
  const endStr = formatDate(endDate);

  // Get users in scope
  const allUsers = getAllUsers(workbook);
  let scopeUserIDs: string[];

  if (session.role === "Admin") {
    scopeUserIDs = allUsers.filter((u) => u.active).map((u) => u.userID);
  } else {
    scopeUserIDs = getSubtreeUserIDs(session.userID, allUsers);
    scopeUserIDs.push(session.userID);
  }

  // Get entries in scope
  const allEntries = getAllTimeEntries(workbook);
  let entries = allEntries.filter((e) => {
    if (!scopeUserIDs.includes(e.employeeID)) return false;
    const d = e.date;
    return d >= startStr && d <= endStr;
  });

  if (teamFilter && teamFilter !== "All") {
    entries = entries.filter((e) => e.team === teamFilter);
  }

  // ── Clear previous report output (rows 8 onward) ────────────────────────
  const usedRange = sheet.getUsedRange();
  if (usedRange) {
    const lastRow = usedRange.getRowCount();
    if (lastRow > 7) {
      sheet.getRange(`A8:H${lastRow + 1}`).clear(ExcelScript.ClearApplyTo.all);
    }
  }

  let outputRow = 8;

  // ── Section 1: Total hours per selected dimension ────────────────────────
  const dimMap = new Map<string, number>();
  for (const e of entries) {
    let key = "";
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

  for (const [key, val] of dimMap) {
    sheet.getRange(`A${outputRow}`).setValue(key);
    sheet.getRange(`B${outputRow}`).setValue(val);
    if (outputRow % 2 === 0) {
      sheet.getRange(`A${outputRow}:B${outputRow}`).getFormat().getFill().setColor("#F5B0A3");
    }
    outputRow++;
  }

  outputRow += 2;

  // ── Section 2: Hours by project ──────────────────────────────────────────
  const projectMap = new Map<string, number>();
  for (const e of entries) {
    projectMap.set(e.project, (projectMap.get(e.project) || 0) + e.hours);
  }

  sheet.getRange(`A${outputRow}`).setValue("Hours by Project");
  sheet.getRange(`A${outputRow}:C${outputRow}`).getFormat().getFont().setBold(true);
  sheet.getRange(`A${outputRow}:C${outputRow}`).getFormat().getFill().setColor("#239A98");
  sheet.getRange(`A${outputRow}:C${outputRow}`).getFormat().getFont().setColor("#FFFFFF");
  outputRow++;

  sheet.getRange(`A${outputRow}`).setValue("Project");
  sheet.getRange(`B${outputRow}`).setValue("Hours");
  sheet.getRange(`A${outputRow}:B${outputRow}`).getFormat().getFont().setBold(true);
  outputRow++;

  for (const [key, val] of projectMap) {
    sheet.getRange(`A${outputRow}`).setValue(key);
    sheet.getRange(`B${outputRow}`).setValue(val);
    outputRow++;
  }

  outputRow += 2;

  // ── Section 3: Hours by activity ─────────────────────────────────────────
  const activityMap = new Map<string, number>();
  for (const e of entries) {
    activityMap.set(e.activity, (activityMap.get(e.activity) || 0) + e.hours);
  }

  sheet.getRange(`A${outputRow}`).setValue("Hours by Activity");
  sheet.getRange(`A${outputRow}:C${outputRow}`).getFormat().getFont().setBold(true);
  sheet.getRange(`A${outputRow}:C${outputRow}`).getFormat().getFill().setColor("#239A98");
  sheet.getRange(`A${outputRow}:C${outputRow}`).getFormat().getFont().setColor("#FFFFFF");
  outputRow++;

  sheet.getRange(`A${outputRow}`).setValue("Activity");
  sheet.getRange(`B${outputRow}`).setValue("Hours");
  sheet.getRange(`A${outputRow}:B${outputRow}`).getFormat().getFont().setBold(true);
  outputRow++;

  for (const [key, val] of activityMap) {
    sheet.getRange(`A${outputRow}`).setValue(key);
    sheet.getRange(`B${outputRow}`).setValue(val);
    outputRow++;
  }

  outputRow += 2;

  // ── Section 4: Daily average per employee ────────────────────────────────
  const empDays = new Map<string, Set<string>>();
  const empHours = new Map<string, number>();

  for (const e of entries) {
    if (!empDays.has(e.employeeName)) empDays.set(e.employeeName, new Set());
    empDays.get(e.employeeName)!.add(e.date);
    empHours.set(e.employeeName, (empHours.get(e.employeeName) || 0) + e.hours);
  }

  sheet.getRange(`A${outputRow}`).setValue("Daily Average per Employee");
  sheet.getRange(`A${outputRow}:C${outputRow}`).getFormat().getFont().setBold(true);
  sheet.getRange(`A${outputRow}:C${outputRow}`).getFormat().getFill().setColor("#239A98");
  sheet.getRange(`A${outputRow}:C${outputRow}`).getFormat().getFont().setColor("#FFFFFF");
  outputRow++;

  sheet.getRange(`A${outputRow}`).setValue("Employee");
  sheet.getRange(`B${outputRow}`).setValue("Total Hours");
  sheet.getRange(`C${outputRow}`).setValue("Days Logged");
  sheet.getRange(`D${outputRow}`).setValue("Daily Average");
  sheet.getRange(`A${outputRow}:D${outputRow}`).getFormat().getFont().setBold(true);
  outputRow++;

  for (const [name, total] of empHours) {
    const days = empDays.get(name)?.size || 1;
    const avg = Math.round((total / days) * 10) / 10;
    sheet.getRange(`A${outputRow}`).setValue(name);
    sheet.getRange(`B${outputRow}`).setValue(total);
    sheet.getRange(`C${outputRow}`).setValue(days);
    sheet.getRange(`D${outputRow}`).setValue(avg);
    outputRow++;
  }

  outputRow += 2;

  // ── Section 5: Under-threshold flag ──────────────────────────────────────
  sheet.getRange(`A${outputRow}`).setValue("Under 8h Day Flags");
  sheet.getRange(`A${outputRow}:C${outputRow}`).getFormat().getFont().setBold(true);
  sheet.getRange(`A${outputRow}:C${outputRow}`).getFormat().getFill().setColor("#D9415C");
  sheet.getRange(`A${outputRow}:C${outputRow}`).getFormat().getFont().setColor("#FFFFFF");
  outputRow++;

  sheet.getRange(`A${outputRow}`).setValue("Employee");
  sheet.getRange(`B${outputRow}`).setValue("Days Under 8h");
  sheet.getRange(`A${outputRow}:B${outputRow}`).getFormat().getFont().setBold(true);
  outputRow++;

  // Calculate per-employee, per-day totals
  const empDayHours = new Map<string, Map<string, number>>();
  for (const e of entries) {
    if (!empDayHours.has(e.employeeName)) empDayHours.set(e.employeeName, new Map());
    const dayMap = empDayHours.get(e.employeeName)!;
    dayMap.set(e.date, (dayMap.get(e.date) || 0) + e.hours);
  }

  for (const [name, dayMap] of empDayHours) {
    let underDays = 0;
    for (const [, hrs] of dayMap) {
      if (hrs < 8) underDays++;
    }
    sheet.getRange(`A${outputRow}`).setValue(name);
    sheet.getRange(`B${outputRow}`).setValue(underDays);
    if (underDays > 0) {
      sheet.getRange(`A${outputRow}:B${outputRow}`).getFormat().getFill().setColor("#FDC400");
    }
    outputRow++;
  }
}
