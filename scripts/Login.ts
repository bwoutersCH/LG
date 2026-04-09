/**
 * Login.ts — Handles user authentication and session management.
 * Linked to the "Log in" and "Log out" buttons on the LOGIN sheet.
 */

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
  if (inputHash !== user.passwordHash) {
    msgCell.setValue("Invalid username or password.");
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
  msgCell.setValue("");

  showSheetsForRole(workbook, user.role);

  const timeSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("TIME_ENTRY");
  if (timeSheet) {
    timeSheet.getRange("B2").setValue(formatDate(new Date()));
  }

  refreshDailyCheck(workbook, sessionData);
}

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
