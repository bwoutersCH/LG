/**
 * Admin.ts — Admin-only operations: user management, reference data, session reset, data export.
 * Linked to buttons on the ADMIN sheet.
 */

/**
 * Saves a new or updated user to Users_DB.
 * Reads input cells from the ADMIN sheet and writes to UsersTable.
 * Linked to the "Save user" button on ADMIN sheet.
 */
function saveUser(workbook: ExcelScript.Workbook): void {
  const session = requireSession(workbook);

  if (session.role !== "Admin") {
    const adminSheet = workbook.getWorksheet("ADMIN");
    if (adminSheet) {
      adminSheet.getRange("F3").setValue("⚠ Admin access required.");
      adminSheet.getRange("F3").getFormat().getFont().setColor("#D9415C");
    }
    return;
  }

  const adminSheet = workbook.getWorksheet("ADMIN");
  if (!adminSheet) return;

  const msgCell = adminSheet.getRange("F3");
  msgCell.setValue("");

  // Read user input fields (Admin sheet layout: column A = labels, column B = values)
  const userID = String(adminSheet.getRange("B3").getValue()).trim();
  const fullName = String(adminSheet.getRange("B4").getValue()).trim();
  const username = String(adminSheet.getRange("B5").getValue()).trim();
  const password = String(adminSheet.getRange("B6").getValue()).trim();
  const role = String(adminSheet.getRange("B7").getValue()).trim();
  const managerID = String(adminSheet.getRange("B8").getValue()).trim();
  const teamID = String(adminSheet.getRange("B9").getValue()).trim();
  const activeStr = String(adminSheet.getRange("B10").getValue()).trim();

  // Validate required fields
  if (!fullName || !username || !role) {
    msgCell.setValue("⚠ FullName, Username, and Role are required.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }

  // Validate role
  const validRoles = ["Admin", "Director", "Manager", "TeamLead", "Employee"];
  if (!validRoles.includes(role)) {
    msgCell.setValue(`⚠ Invalid role. Must be one of: ${validRoles.join(", ")}`);
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }

  const usersSheet = workbook.getWorksheet("USERS_DB");
  if (!usersSheet) return;

  const table = usersSheet.getTable("UsersTable");
  if (!table) return;

  const rows = table.getRangeBetweenHeaderAndTotal().getValues();
  const active = activeStr.toUpperCase() !== "FALSE";

  // Check if updating existing user
  let existingRow = -1;
  if (userID) {
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === userID) {
        existingRow = i;
        break;
      }
    }
  }

  if (existingRow >= 0) {
    // Update existing user
    const dataRange = table.getRangeBetweenHeaderAndTotal();
    dataRange.getCell(existingRow, 1).setValue(fullName);
    dataRange.getCell(existingRow, 2).setValue(username);
    if (password) {
      dataRange.getCell(existingRow, 3).setValue(simpleHash(password));
    }
    dataRange.getCell(existingRow, 4).setValue(role);
    dataRange.getCell(existingRow, 5).setValue(managerID);
    dataRange.getCell(existingRow, 6).setValue(teamID);
    dataRange.getCell(existingRow, 7).setValue(active);

    msgCell.setValue(`✓ User ${userID} (${fullName}) updated.`);
    msgCell.getFormat().getFont().setColor("#239A98");
  } else {
    // Generate new UserID
    const maxID = rows.reduce((max, r) => {
      const num = parseInt(String(r[0]).replace("U", ""), 10);
      return num > max ? num : max;
    }, 0);
    const newID = `U${String(maxID + 1).padStart(3, "0")}`;

    if (!password) {
      msgCell.setValue("⚠ Password is required for new users.");
      msgCell.getFormat().getFont().setColor("#D9415C");
      return;
    }

    table.addRow(-1, [
      newID,
      fullName,
      username,
      simpleHash(password),
      role,
      managerID,
      teamID,
      active,
    ]);

    msgCell.setValue(`✓ User ${newID} (${fullName}) created.`);
    msgCell.getFormat().getFont().setColor("#239A98");
  }

  // Clear input fields
  adminSheet.getRange("B3:B10").setValue("");
}

/**
 * Saves a new or updated project to Projects_DB.
 * Linked to the "Save project" button on ADMIN sheet.
 */
function saveProject(workbook: ExcelScript.Workbook): void {
  const session = requireSession(workbook);
  if (session.role !== "Admin") return;

  const adminSheet = workbook.getWorksheet("ADMIN");
  if (!adminSheet) return;

  const msgCell = adminSheet.getRange("F13");
  msgCell.setValue("");

  const projectID = String(adminSheet.getRange("B13").getValue()).trim();
  const projectName = String(adminSheet.getRange("B14").getValue()).trim();
  const teamID = String(adminSheet.getRange("B15").getValue()).trim();
  const activeStr = String(adminSheet.getRange("B16").getValue()).trim();

  if (!projectName) {
    msgCell.setValue("⚠ Project name is required.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }

  const projSheet = workbook.getWorksheet("PROJECTS_DB");
  if (!projSheet) return;

  const table = projSheet.getTable("ProjectsTable");
  if (!table) return;

  const rows = table.getRangeBetweenHeaderAndTotal().getValues();
  const active = activeStr.toUpperCase() !== "FALSE";

  let existingRow = -1;
  if (projectID) {
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === projectID) {
        existingRow = i;
        break;
      }
    }
  }

  if (existingRow >= 0) {
    const dataRange = table.getRangeBetweenHeaderAndTotal();
    dataRange.getCell(existingRow, 1).setValue(projectName);
    dataRange.getCell(existingRow, 2).setValue(teamID);
    dataRange.getCell(existingRow, 3).setValue(active);
    msgCell.setValue(`✓ Project ${projectID} updated.`);
    msgCell.getFormat().getFont().setColor("#239A98");
  } else {
    const maxID = rows.reduce((max, r) => {
      const num = parseInt(String(r[0]).replace("P", ""), 10);
      return num > max ? num : max;
    }, 0);
    const newID = `P${String(maxID + 1).padStart(3, "0")}`;

    table.addRow(-1, [newID, projectName, teamID, active]);
    msgCell.setValue(`✓ Project ${newID} created.`);
    msgCell.getFormat().getFont().setColor("#239A98");
  }

  adminSheet.getRange("B13:B16").setValue("");
}

/**
 * Saves a new or updated activity to Activities_DB.
 * Linked to the "Save activity" button on ADMIN sheet.
 */
function saveActivity(workbook: ExcelScript.Workbook): void {
  const session = requireSession(workbook);
  if (session.role !== "Admin") return;

  const adminSheet = workbook.getWorksheet("ADMIN");
  if (!adminSheet) return;

  const msgCell = adminSheet.getRange("F18");
  msgCell.setValue("");

  const activityID = String(adminSheet.getRange("B18").getValue()).trim();
  const activityName = String(adminSheet.getRange("B19").getValue()).trim();
  const activeStr = String(adminSheet.getRange("B20").getValue()).trim();

  if (!activityName) {
    msgCell.setValue("⚠ Activity name is required.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }

  const actSheet = workbook.getWorksheet("ACTIVITIES_DB");
  if (!actSheet) return;

  const table = actSheet.getTable("ActivitiesTable");
  if (!table) return;

  const rows = table.getRangeBetweenHeaderAndTotal().getValues();
  const active = activeStr.toUpperCase() !== "FALSE";

  let existingRow = -1;
  if (activityID) {
    for (let i = 0; i < rows.length; i++) {
      if (String(rows[i][0]) === activityID) {
        existingRow = i;
        break;
      }
    }
  }

  if (existingRow >= 0) {
    const dataRange = table.getRangeBetweenHeaderAndTotal();
    dataRange.getCell(existingRow, 1).setValue(activityName);
    dataRange.getCell(existingRow, 2).setValue(active);
    msgCell.setValue(`✓ Activity ${activityID} updated.`);
    msgCell.getFormat().getFont().setColor("#239A98");
  } else {
    const maxID = rows.reduce((max, r) => {
      const num = parseInt(String(r[0]).replace("A", ""), 10);
      return num > max ? num : max;
    }, 0);
    const newID = `A${String(maxID + 1).padStart(3, "0")}`;

    table.addRow(-1, [newID, activityName, active]);
    msgCell.setValue(`✓ Activity ${newID} created.`);
    msgCell.getFormat().getFont().setColor("#239A98");
  }

  adminSheet.getRange("B18:B20").setValue("");
}

/**
 * Resets the session for testing purposes.
 * Linked to the "Reset session" button on ADMIN sheet.
 */
function resetSession(workbook: ExcelScript.Workbook): void {
  clearSession(workbook);
  hideAllWorksheets(workbook);

  const loginSheet = workbook.getWorksheet("LOGIN");
  if (loginSheet) {
    loginSheet.getRange("B3").setValue("");
    loginSheet.getRange("B4").setValue("");
    loginSheet.getRange("B6").setValue("");
    loginSheet.activate();
  }
}

/**
 * Exports all time entry data to a new sheet called "EXPORT".
 * Linked to the "Export all data" button on ADMIN sheet.
 */
function exportAllData(workbook: ExcelScript.Workbook): void {
  const session = requireSession(workbook);
  if (session.role !== "Admin") return;

  // Delete existing EXPORT sheet if present
  const existing = workbook.getWorksheet("EXPORT");
  if (existing) {
    existing.delete();
  }

  const exportSheet = workbook.addWorksheet("EXPORT");

  // Copy headers
  const headers = [
    "EntryID", "Date", "EmployeeID", "EmployeeName", "Team",
    "Project", "Activity", "Hours", "Notes", "SubmittedOn",
    "LastEditedBy", "LastEditedOn",
  ];

  for (let i = 0; i < headers.length; i++) {
    const cell = exportSheet.getRange(`${String.fromCharCode(65 + i)}1`);
    cell.setValue(headers[i]);
    cell.getFormat().getFont().setBold(true);
    cell.getFormat().getFont().setColor("#FFFFFF");
    cell.getFormat().getFill().setColor("#243347");
  }

  // Copy all time entries
  const entries = getAllTimeEntries(workbook);
  for (let r = 0; r < entries.length; r++) {
    const e = entries[r];
    const row = r + 2;
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

  const adminSheet = workbook.getWorksheet("ADMIN");
  if (adminSheet) {
    adminSheet.getRange("F22").setValue(`✓ Exported ${entries.length} entries to EXPORT sheet.`);
    adminSheet.getRange("F22").getFormat().getFont().setColor("#239A98");
  }
}
