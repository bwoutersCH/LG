/**
 * InitializeWorkbook.ts — Populates all DB tables with sample data.
 * Run ONCE after creating the workbook structure.
 * Creates 25 users, 5 projects, 6 activities, 4 teams, and 3 weeks of time entries.
 */

function initializeWorkbook(workbook: ExcelScript.Workbook): void {
  populateTeams(workbook);
  populateActivities(workbook);
  populateProjects(workbook);
  populateUsers(workbook);
  populateTimeEntries(workbook);

  hideAllWorksheets(workbook);

  const readmeSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("README");
  if (readmeSheet) readmeSheet.setVisibility(ExcelScript.SheetVisibility.visible);
}

function populateTeams(workbook: ExcelScript.Workbook): void {
  const sheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("TEAMS_DB");
  if (!sheet) return;
  const table: ExcelScript.Table | undefined = sheet.getTable("TeamsTable");
  if (!table) return;

  const teams: string[][] = [
    ["T001", "Engineering", "", "U002"],
    ["T002", "Design", "", "U003"],
    ["T003", "Backend", "T001", "U004"],
    ["T004", "Frontend", "T001", "U005"],
  ];

  for (const t of teams) {
    table.addRow(-1, t);
  }
}

function populateActivities(workbook: ExcelScript.Workbook): void {
  const sheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("ACTIVITIES_DB");
  if (!sheet) return;
  const table: ExcelScript.Table | undefined = sheet.getTable("ActivitiesTable");
  if (!table) return;

  const activities: (string | boolean)[][] = [
    ["A001", "Development", true],
    ["A002", "Design", true],
    ["A003", "Testing", true],
    ["A004", "Planning", true],
    ["A005", "Client Meeting", true],
    ["A006", "Admin", true],
  ];

  for (const a of activities) {
    table.addRow(-1, a);
  }
}

function populateProjects(workbook: ExcelScript.Workbook): void {
  const sheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("PROJECTS_DB");
  if (!sheet) return;
  const table: ExcelScript.Table | undefined = sheet.getTable("ProjectsTable");
  if (!table) return;

  const projects: (string | boolean)[][] = [
    ["P001", "Website Redesign", "T004", true],
    ["P002", "API Platform", "T003", true],
    ["P003", "Mobile App", "T001", true],
    ["P004", "Data Pipeline", "T003", true],
    ["P005", "Brand Refresh", "T002", true],
  ];

  for (const p of projects) {
    table.addRow(-1, p);
  }
}

function populateUsers(workbook: ExcelScript.Workbook): void {
  const sheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("USERS_DB");
  if (!sheet) return;
  const table: ExcelScript.Table | undefined = sheet.getTable("UsersTable");
  if (!table) return;

  const pwd: string = simpleHash("Pass1234");

  const users: (string | boolean)[][] = [
    ["U001", "Alice Admin", "alice.admin", pwd, "Admin", "", "T001", true],
    ["U002", "Bob Director", "bob.director", pwd, "Director", "U001", "T001", true],
    ["U003", "Carol Director", "carol.director", pwd, "Director", "U001", "T002", true],
    ["U004", "Dave Manager", "dave.manager", pwd, "Manager", "U002", "T003", true],
    ["U005", "Eve Manager", "eve.manager", pwd, "Manager", "U002", "T004", true],
    ["U006", "Frank Manager", "frank.manager", pwd, "Manager", "U003", "T002", true],
    ["U007", "Grace Manager", "grace.manager", pwd, "Manager", "U003", "T002", true],
    ["U008", "Hank TeamLead", "hank.teamlead", pwd, "TeamLead", "U004", "T003", true],
    ["U009", "Ivy TeamLead", "ivy.teamlead", pwd, "TeamLead", "U005", "T004", true],
    ["U010", "Jack TeamLead", "jack.teamlead", pwd, "TeamLead", "U006", "T002", true],
    ["U011", "Kate TeamLead", "kate.teamlead", pwd, "TeamLead", "U007", "T002", true],
    ["U012", "Leo Employee", "leo.employee", pwd, "Employee", "U008", "T003", true],
    ["U013", "Mia Employee", "mia.employee", pwd, "Employee", "U008", "T003", true],
    ["U014", "Noah Employee", "noah.employee", pwd, "Employee", "U008", "T003", true],
    ["U015", "Olivia Employee", "olivia.employee", pwd, "Employee", "U009", "T004", true],
    ["U016", "Pete Employee", "pete.employee", pwd, "Employee", "U009", "T004", true],
    ["U017", "Quinn Employee", "quinn.employee", pwd, "Employee", "U009", "T004", true],
    ["U018", "Ryan Employee", "ryan.employee", pwd, "Employee", "U010", "T002", true],
    ["U019", "Sara Employee", "sara.employee", pwd, "Employee", "U010", "T002", true],
    ["U020", "Tom Employee", "tom.employee", pwd, "Employee", "U010", "T002", true],
    ["U021", "Uma Employee", "uma.employee", pwd, "Employee", "U011", "T002", true],
    ["U022", "Vic Employee", "vic.employee", pwd, "Employee", "U011", "T002", true],
    ["U023", "Wendy Employee", "wendy.employee", pwd, "Employee", "U004", "T003", true],
    ["U024", "Xander Employee", "xander.employee", pwd, "Employee", "U005", "T004", true],
    ["U025", "Yara Employee", "yara.employee", pwd, "Employee", "U006", "T002", true],
  ];

  for (const u of users) {
    table.addRow(-1, u);
  }
}

function populateTimeEntries(workbook: ExcelScript.Workbook): void {
  const sheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("TIME_ENTRY");
  if (!sheet) return;
  const table: ExcelScript.Table | undefined = sheet.getTable("TimeEntryTable");
  if (!table) return;

  const allUsers: UserRow[] = getAllUsers(workbook);
  const projects: string[] = ["Website Redesign", "API Platform", "Mobile App", "Data Pipeline", "Brand Refresh"];
  const activities: string[] = ["Development", "Design", "Testing", "Planning", "Client Meeting", "Admin"];

  const today: Date = new Date();
  const currentWeekStart: Date = getWeekStart(today);

  const startDate: Date = new Date(currentWeekStart);
  startDate.setDate(startDate.getDate() - 14);

  let entryCount: number = 0;

  for (const user of allUsers) {
    if (user.role === "Admin" || user.role === "Director") continue;

    const teamName: string = getTeamName(workbook, user.teamID);

    for (let week: number = 0; week < 3; week++) {
      const weekStart: Date = new Date(startDate);
      weekStart.setDate(weekStart.getDate() + week * 7);

      for (let day: number = 0; day < 5; day++) {
        const entryDate: Date = new Date(weekStart);
        entryDate.setDate(weekStart.getDate() + day);

        if (entryDate > today) continue;

        const dateStr: string = formatDate(entryDate);

        const proj1: string = projects[entryCount % projects.length];
        const act1: string = activities[entryCount % activities.length];
        const hours1: number = 5 + (entryCount % 3);
        const hours2: number = 8 - hours1 + (entryCount % 2);

        entryCount++;
        const id1: string = `TE-INIT-${String(entryCount).padStart(5, "0")}`;
        const ts: string = entryDate.toISOString();

        table.addRow(-1, [
          id1, dateStr, user.userID, user.fullName, teamName,
          proj1, act1, hours1,
          "Sample entry", ts, user.username, ts,
        ]);

        if (hours2 > 0 && hours2 <= 16) {
          entryCount++;
          const id2: string = `TE-INIT-${String(entryCount).padStart(5, "0")}`;
          const proj2: string = projects[(entryCount + 2) % projects.length];
          const act2: string = activities[(entryCount + 1) % activities.length];

          table.addRow(-1, [
            id2, dateStr, user.userID, user.fullName, teamName,
            proj2, act2, hours2,
            "Sample entry", ts, user.username, ts,
          ]);
        }
      }
    }
  }
}
