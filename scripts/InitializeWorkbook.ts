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

  // Hide DB sheets and show only LOGIN + README
  hideAllWorksheets(workbook);

  const readmeSheet = workbook.getWorksheet("README");
  if (readmeSheet) readmeSheet.setVisibility(ExcelScript.SheetVisibility.visible);
}

function populateTeams(workbook: ExcelScript.Workbook): void {
  const sheet = workbook.getWorksheet("TEAMS_DB");
  if (!sheet) return;
  const table = sheet.getTable("TeamsTable");
  if (!table) return;

  const teams = [
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
  const sheet = workbook.getWorksheet("ACTIVITIES_DB");
  if (!sheet) return;
  const table = sheet.getTable("ActivitiesTable");
  if (!table) return;

  const activities = [
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
  const sheet = workbook.getWorksheet("PROJECTS_DB");
  if (!sheet) return;
  const table = sheet.getTable("ProjectsTable");
  if (!table) return;

  const projects = [
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
  const sheet = workbook.getWorksheet("USERS_DB");
  if (!sheet) return;
  const table = sheet.getTable("UsersTable");
  if (!table) return;

  // Password for all sample users is "Pass1234" — hash computed by simpleHash()
  const pwd = simpleHash("Pass1234");

  const users: (string | boolean)[][] = [
    // Admins (1)
    ["U001", "Alice Admin", "alice.admin", pwd, "Admin", "", "T001", true],
    // Directors (2)
    ["U002", "Bob Director", "bob.director", pwd, "Director", "U001", "T001", true],
    ["U003", "Carol Director", "carol.director", pwd, "Director", "U001", "T002", true],
    // Managers (4)
    ["U004", "Dave Manager", "dave.manager", pwd, "Manager", "U002", "T003", true],
    ["U005", "Eve Manager", "eve.manager", pwd, "Manager", "U002", "T004", true],
    ["U006", "Frank Manager", "frank.manager", pwd, "Manager", "U003", "T002", true],
    ["U007", "Grace Manager", "grace.manager", pwd, "Manager", "U003", "T002", true],
    // Team Leads (4)
    ["U008", "Hank TeamLead", "hank.teamlead", pwd, "TeamLead", "U004", "T003", true],
    ["U009", "Ivy TeamLead", "ivy.teamlead", pwd, "TeamLead", "U005", "T004", true],
    ["U010", "Jack TeamLead", "jack.teamlead", pwd, "TeamLead", "U006", "T002", true],
    ["U011", "Kate TeamLead", "kate.teamlead", pwd, "TeamLead", "U007", "T002", true],
    // Employees (14)
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
  const sheet = workbook.getWorksheet("TIME_ENTRY");
  if (!sheet) return;
  const table = sheet.getTable("TimeEntryTable");
  if (!table) return;

  // Get user data for names and teams
  const allUsers = getAllUsers(workbook);
  const projects = ["Website Redesign", "API Platform", "Mobile App", "Data Pipeline", "Brand Refresh"];
  const activities = ["Development", "Design", "Testing", "Planning", "Client Meeting", "Admin"];

  // Generate 3 weeks of sample data ending this week
  const today = new Date();
  const currentWeekStart = getWeekStart(today);

  // Start 2 weeks before current week (3 weeks total)
  const startDate = new Date(currentWeekStart);
  startDate.setDate(startDate.getDate() - 14);

  let entryCount = 0;

  for (const user of allUsers) {
    // Only employees and team leads log time in sample data
    if (user.role === "Admin" || user.role === "Director") continue;

    const teamName = getTeamName(workbook, user.teamID);

    // For each of 3 weeks
    for (let week = 0; week < 3; week++) {
      const weekStart = new Date(startDate);
      weekStart.setDate(weekStart.getDate() + week * 7);

      // For each weekday (Mon–Fri)
      for (let day = 0; day < 5; day++) {
        const entryDate = new Date(weekStart);
        entryDate.setDate(weekStart.getDate() + day);

        // Skip future dates
        if (entryDate > today) continue;

        const dateStr = formatDate(entryDate);

        // Each person logs 1–2 entries per day totaling ~8h
        const proj1 = projects[entryCount % projects.length];
        const act1 = activities[entryCount % activities.length];
        const hours1 = 5 + (entryCount % 3); // 5, 6, or 7 hours
        const hours2 = 8 - hours1 + (entryCount % 2); // fill to 8 or 9

        entryCount++;
        const id1 = `TE-INIT-${String(entryCount).padStart(5, "0")}`;
        const ts = entryDate.toISOString();

        table.addRow(-1, [
          id1, dateStr, user.userID, user.fullName, teamName,
          proj1, act1, hours1,
          "Sample entry", ts, user.username, ts,
        ]);

        if (hours2 > 0 && hours2 <= 16) {
          entryCount++;
          const id2 = `TE-INIT-${String(entryCount).padStart(5, "0")}`;
          const proj2 = projects[(entryCount + 2) % projects.length];
          const act2 = activities[(entryCount + 1) % activities.length];

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
