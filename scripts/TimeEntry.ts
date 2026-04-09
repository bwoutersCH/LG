/**
 * TimeEntry.ts — Handles time entry creation, validation, filtering, and weekly checks.
 * Linked to buttons on the TIME_ENTRY and DAILY_CHECK sheets.
 */

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

  if (!dateValue || !project || !activity || !hoursValue) {
    msgCell.setValue("Please fill in Date, Project, Activity, and Hours.");
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

  sheet.getRange("B5").setValue("");
  sheet.getRange("B6").setValue("");

  msgCell.setValue(`Entry ${entryID} added successfully (${hours}h on ${dateStr}).`);
  msgCell.getFormat().getFont().setColor("#239A98");

  refreshDailyCheck(workbook, session);
}

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
        issues.push(`${dayNames[i]} (${formatDate(dateForDay)}): ${dayHours[i]}h — needs ${8 - dayHours[i]}h more`);
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

function deleteEntry(workbook: ExcelScript.Workbook): void {
  const session: SessionData = requireSession(workbook);
  const sheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("TIME_ENTRY");
  if (!sheet) return;

  const entryID: string = String(sheet.getRange("E2").getValue()).trim();
  const msgCell: ExcelScript.Range = sheet.getRange("D7");

  if (!entryID) {
    msgCell.setValue("Please enter an Entry ID in cell E2 to delete.");
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
  sheet.getRange("E2").setValue("");
  msgCell.setValue(`Entry ${entryID} deleted.`);
  msgCell.getFormat().getFont().setColor("#239A98");

  refreshDailyCheck(workbook, session);
}
