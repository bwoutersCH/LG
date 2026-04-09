/**
 * Login.ts — Handles user authentication and session management.
 * Linked to the "Log in" and "Log out" buttons on the LOGIN sheet.
 */

/**
 * Main login function — linked to the "Log in" button on the LOGIN sheet.
 * Reads username/password from cells, validates against Users_DB,
 * writes session, and shows role-appropriate sheets.
 */
function runLogin(workbook: ExcelScript.Workbook): void {
  const loginSheet = workbook.getWorksheet("LOGIN");
  if (!loginSheet) return;

  const msgCell = loginSheet.getRange("B6");
  msgCell.setValue("");

  // Read input
  const username = String(loginSheet.getRange("B3").getValue()).trim();
  const password = String(loginSheet.getRange("B4").getValue()).trim();

  // Validate inputs
  if (!username || !password) {
    msgCell.setValue("⚠ Please enter both username and password.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    return;
  }

  // Get all users
  const allUsers = getAllUsers(workbook);
  const user = findUserByUsername(allUsers, username);

  if (!user) {
    msgCell.setValue("⚠ Invalid username or password.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    loginSheet.getRange("B4").setValue("");
    return;
  }

  // Check password hash
  const inputHash = simpleHash(password);
  if (inputHash !== user.passwordHash) {
    msgCell.setValue("⚠ Invalid username or password.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    loginSheet.getRange("B4").setValue("");
    return;
  }

  // Check if user is active
  if (!user.active) {
    msgCell.setValue("⚠ This account has been deactivated. Contact an administrator.");
    msgCell.getFormat().getFont().setColor("#D9415C");
    loginSheet.getRange("B4").setValue("");
    return;
  }

  // Write session
  const sessionData: SessionData = {
    userID: user.userID,
    username: user.username,
    role: user.role,
    managerID: user.managerID,
    loginTime: new Date().toISOString(),
  };
  writeSession(workbook, sessionData);

  // Clear password field
  loginSheet.getRange("B4").setValue("");
  msgCell.setValue("");

  // Show appropriate sheets for role
  showSheetsForRole(workbook, user.role);

  // Set default date on TIME_ENTRY sheet
  const timeSheet = workbook.getWorksheet("TIME_ENTRY");
  if (timeSheet) {
    timeSheet.getRange("B2").setValue(formatDate(new Date()));
  }

  // Refresh DAILY_CHECK for the logged-in user
  refreshDailyCheck(workbook, sessionData);
}

/**
 * Logout function — linked to the "Log out" button.
 * Clears session and hides all non-login sheets.
 */
function runLogout(workbook: ExcelScript.Workbook): void {
  clearSession(workbook);
  hideAllWorksheets(workbook);

  // Clear login form
  const loginSheet = workbook.getWorksheet("LOGIN");
  if (loginSheet) {
    loginSheet.getRange("B3").setValue("");
    loginSheet.getRange("B4").setValue("");
    loginSheet.getRange("B6").setValue("");
    loginSheet.activate();
  }
}

/**
 * Refreshes the DAILY_CHECK sheet with data for the current user's week.
 */
function refreshDailyCheck(workbook: ExcelScript.Workbook, session: SessionData): void {
  const checkSheet = workbook.getWorksheet("DAILY_CHECK");
  if (!checkSheet) return;

  const today = new Date();
  const weekStart = getWeekStart(today);
  const entries = getAllTimeEntries(workbook);

  // Get entries for the current user this week
  const days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"];
  const dayHours: number[] = [0, 0, 0, 0, 0];

  for (const entry of entries) {
    if (entry.employeeID !== session.userID) continue;
    const entryDate = parseExcelDate(entry.date);
    const entryWeekStart = getWeekStart(entryDate);
    if (formatDate(entryWeekStart) !== formatDate(weekStart)) continue;

    const dayOfWeek = entryDate.getDay();
    // Monday=1 -> index 0, Friday=5 -> index 4
    if (dayOfWeek >= 1 && dayOfWeek <= 5) {
      dayHours[dayOfWeek - 1] += entry.hours;
    }
  }

  // Write summary starting at row 3
  for (let i = 0; i < 5; i++) {
    const dateForDay = new Date(weekStart);
    dateForDay.setDate(weekStart.getDate() + i);

    const row = 3 + i;
    checkSheet.getRange(`A${row}`).setValue(days[i]);
    checkSheet.getRange(`B${row}`).setValue(formatDate(dateForDay));
    checkSheet.getRange(`C${row}`).setValue(dayHours[i]);

    // Highlight days < 8h in yellow
    const fmt = checkSheet.getRange(`A${row}:C${row}`).getFormat();
    if (dayHours[i] < 8 && formatDate(dateForDay) <= formatDate(today)) {
      fmt.getFill().setColor("#FDC400");
      fmt.getFont().setColor("#575756");
    } else {
      fmt.getFill().setColor("#FFFFFF");
      fmt.getFont().setColor("#575756");
    }
  }

  // Total row
  const totalHours = dayHours.reduce((a, b) => a + b, 0);
  checkSheet.getRange("A8").setValue("TOTAL");
  checkSheet.getRange("C8").setValue(totalHours);
  checkSheet.getRange("A8:C8").getFormat().getFont().setBold(true);

  // Header info
  checkSheet.getRange("A1").setValue(`Weekly Hours — ${session.username} — Week of ${formatDate(weekStart)}`);
}
