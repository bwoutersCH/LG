/**
 * TimeEntry.ts — Handles time entry creation, validation, filtering, and weekly checks.
 * Linked to buttons on the TIME_ENTRY and DAILY_CHECK sheets.
 */

/**
 * Adds a new time entry. Linked to the "Add entry" button on TIME_ENTRY sheet.
 * Reads input cells, validates, checks for duplicates, and appends to TimeEntryTable.
 */
function addEntry(workbook: ExcelScript.Workbook): void {
  const session = requireSession(workbook);
  const sheet = workbook.getWorksheet("TIME_ENTRY");
  if (!sheet) return;

  const msgCell = sheet.getRange("D7");
  msgCell.setValue("");

  // Read input values
  const dateValue = sheet.getRange("B2").getValue();
  const project = String(sheet.getRange("B3").getValue()).trim();
  const activity = String(sheet.getRange("B4").getValue()).trim();
  const hoursValue = sheet.getRange("B5").getValue();
  const notes = String(sheet.getRange("B6").getValue()).trim();

  // ── Validate inputs ──────────────────────────────────────────────────────
  if (!dateValue || !project || !activity || !hoursValue) {
    msgCell.setValue("⚠ Please fill in Date, Project, Activity, and Hours.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }

  // Parse and validate date
  const entryDate = parseExcelDate(dateValue);
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const diffDays = Math.floor((entryDate.getTime() - today.getTime()) / (1000 * 60 * 60 * 24));
  if (diffDays < -30 || diffDays > 30) {
    msgCell.setValue("⚠ Date must be within 30 days of today.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }

  // Validate hours (integer, 1–16)
  const hours = Number(hoursValue);
  if (!Number.isInteger(hours) || hours < 1 || hours > 16) {
    msgCell.setValue("⚠ Hours must be a whole number between 1 and 16.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }

  // ── Duplicate check ──────────────────────────────────────────────────────
  const allEntries = getAllTimeEntries(workbook);
  const dateStr = formatDate(entryDate);
  const duplicate = allEntries.find(
    (e) =>
      e.employeeID === session.userID &&
      e.date === dateStr &&
      e.project === project &&
      e.activity === activity
  );

  if (duplicate) {
    msgCell.setValue(`⚠ Duplicate: entry for ${project} / ${activity} on ${dateStr} already exists (${duplicate.entryID}).`);
    msgCell.getFormat().getFont().setColor("#FDC400");
    // Allow user to proceed by adding again — this is just a warning
    // To block, uncomment: return;
  }

  // ── Get employee info ────────────────────────────────────────────────────
  const allUsers = getAllUsers(workbook);
  const currentUser = findUserByID(allUsers, session.userID);
  if (!currentUser) {
    msgCell.setValue("⚠ User not found in database.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }

  const teamName = getTeamName(workbook, currentUser.teamID);
  const now = new Date().toISOString();
  const entryID = generateEntryID();

  // ── Add row to table ─────────────────────────────────────────────────────
  const table = sheet.getTable("TimeEntryTable");
  if (!table) {
    msgCell.setValue("⚠ TimeEntryTable not found.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }

  table.addRow(-1, [
    entryID,
    dateStr,
    session.userID,
    currentUser.fullName,
    teamName,
    project,
    activity,
    hours,
    notes,
    now,
    session.username,
    now,
  ]);

  // ── Clear input fields ───────────────────────────────────────────────────
  sheet.getRange("B5").setValue("");
  sheet.getRange("B6").setValue("");

  msgCell.setValue(`✓ Entry ${entryID} added successfully (${hours}h on ${dateStr}).`);
  msgCell.getFormat().getFont().setColor("#239A98");

  // Refresh daily check
  refreshDailyCheck(workbook, session);
}

/**
 * Filters the TIME_ENTRY table to show only the current user's entries for this week.
 * Linked to the "My entries this week" button.
 */
function filterMyEntries(workbook: ExcelScript.Workbook): void {
  const session = requireSession(workbook);
  const sheet = workbook.getWorksheet("TIME_ENTRY");
  if (!sheet) return;

  const table = sheet.getTable("TimeEntryTable");
  if (!table) return;

  const today = new Date();
  const weekStart = getWeekStart(today);
  const weekEnd = getWeekEnd(today);
  const weekStartStr = formatDate(weekStart);
  const weekEndStr = formatDate(weekEnd);

  // Remove existing filters
  table.getAutoFilter().clearCriteria();

  // Apply filter on EmployeeID column (index 2)
  table.getAutoFilter().apply(
    table.getColumn("EmployeeID").getRange(),
    2,
    { filterOn: ExcelScript.FilterOn.values, values: [session.userID] }
  );

  // Show status message
  const msgCell = sheet.getRange("D7");
  msgCell.setValue(`Showing entries for ${session.username}, week ${weekStartStr} to ${weekEndStr}`);
  msgCell.getFormat().getFont().setColor("#239A98");
}

/**
 * Clears all filters on the TIME_ENTRY table.
 * Linked to a "Show all" button if desired.
 */
function clearFilters(workbook: ExcelScript.Workbook): void {
  requireSession(workbook);
  const sheet = workbook.getWorksheet("TIME_ENTRY");
  if (!sheet) return;

  const table = sheet.getTable("TimeEntryTable");
  if (!table) return;

  table.getAutoFilter().clearCriteria();

  const msgCell = sheet.getRange("D7");
  msgCell.setValue("Filters cleared.");
  msgCell.getFormat().getFont().setColor("#575756");
}

/**
 * Validates the current week for the logged-in user.
 * Checks that every weekday (Mon–Fri) has >= 8 hours logged.
 * Linked to the "Validate week" button on DAILY_CHECK sheet.
 */
function validateWeek(workbook: ExcelScript.Workbook): void {
  const session = requireSession(workbook);
  const checkSheet = workbook.getWorksheet("DAILY_CHECK");
  if (!checkSheet) return;

  const today = new Date();
  const weekStart = getWeekStart(today);
  const entries = getAllTimeEntries(workbook);

  const dayNames = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];
  const dayHours: number[] = [0, 0, 0, 0, 0];

  for (const entry of entries) {
    if (entry.employeeID !== session.userID) continue;
    const entryDate = parseExcelDate(entry.date);
    const entryWeekStart = getWeekStart(entryDate);
    if (formatDate(entryWeekStart) !== formatDate(weekStart)) continue;

    const dayOfWeek = entryDate.getDay();
    if (dayOfWeek >= 1 && dayOfWeek <= 5) {
      dayHours[dayOfWeek - 1] += entry.hours;
    }
  }

  // Build result
  const issues: string[] = [];
  for (let i = 0; i < 5; i++) {
    const dateForDay = new Date(weekStart);
    dateForDay.setDate(weekStart.getDate() + i);

    // Only check days up to today
    if (dateForDay <= today) {
      if (dayHours[i] < 8) {
        issues.push(`${dayNames[i]} (${formatDate(dateForDay)}): ${dayHours[i]}h — needs ${8 - dayHours[i]}h more`);
      }
    }
  }

  const msgCell = checkSheet.getRange("A10");
  if (issues.length === 0) {
    msgCell.setValue("✓ Week is complete! All days have 8+ hours logged.");
    msgCell.getFormat().getFont().setColor("#239A98");
    msgCell.getFormat().getFont().setBold(true);
  } else {
    const msg = "⚠ Incomplete days:\n" + issues.join("\n");
    msgCell.setValue(msg);
    msgCell.getFormat().getFont().setColor("#D9415C");
    msgCell.getFormat().getFont().setBold(true);
  }

  // Refresh the daily check display
  refreshDailyCheck(workbook, session);
}

/**
 * Deletes a time entry by EntryID. Only the owner or a manager in the chain can delete.
 */
function deleteEntry(workbook: ExcelScript.Workbook): void {
  const session = requireSession(workbook);
  const sheet = workbook.getWorksheet("TIME_ENTRY");
  if (!sheet) return;

  const entryID = String(sheet.getRange("E2").getValue()).trim();
  const msgCell = sheet.getRange("D7");

  if (!entryID) {
    msgCell.setValue("⚠ Please enter an Entry ID in cell E2 to delete.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }

  const table = sheet.getTable("TimeEntryTable");
  if (!table) return;

  const allUsers = getAllUsers(workbook);
  const rows = table.getRangeBetweenHeaderAndTotal().getValues();

  let rowIndex = -1;
  for (let i = 0; i < rows.length; i++) {
    if (String(rows[i][0]) === entryID) {
      rowIndex = i;
      break;
    }
  }

  if (rowIndex === -1) {
    msgCell.setValue(`⚠ Entry ${entryID} not found.`);
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }

  const ownerID = String(rows[rowIndex][2]);
  if (!canUserEditEntry(session.userID, ownerID, allUsers)) {
    msgCell.setValue("⚠ You do not have permission to delete this entry.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }

  table.deleteRowsAt(rowIndex, 1);
  sheet.getRange("E2").setValue("");
  msgCell.setValue(`✓ Entry ${entryID} deleted.`);
  msgCell.getFormat().getFont().setColor("#239A98");

  refreshDailyCheck(workbook, session);
}
