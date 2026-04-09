/**
 * RefreshDropdowns.ts — Updates all dropdown data validations from DB tables.
 * Run this after adding/editing projects, activities, or teams.
 * Self-contained Office Script. Entry point: main()
 */

function main(workbook: ExcelScript.Workbook): void {
  refreshProjectDropdown(workbook);
  refreshActivityDropdown(workbook);
  refreshTeamDropdowns(workbook);

  // Show confirmation on TIME_ENTRY sheet
  const teSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("TIME_ENTRY");
  if (teSheet) {
    const msgCell: ExcelScript.Range = teSheet.getRange("D7");
    msgCell.setValue("Dropdowns refreshed from database tables.");
    msgCell.getFormat().getFont().setColor("#239A98");
  }
}

function refreshProjectDropdown(workbook: ExcelScript.Workbook): void {
  const dbSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("PROJECTS_DB");
  if (!dbSheet) return;
  const table: ExcelScript.Table | undefined = dbSheet.getTable("ProjectsTable");
  if (!table) return;

  const rows: (string | number | boolean)[][] = table.getRangeBetweenHeaderAndTotal().getValues();
  const names: string[] = [];
  for (const r of rows) {
    // Only include active projects (column 3 = Active)
    if (Boolean(r[3])) {
      names.push(String(r[1]));
    }
  }

  const teSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("TIME_ENTRY");
  if (!teSheet) return;

  // Remove existing validation and set new one
  const cell: ExcelScript.Range = teSheet.getRange("B3");
  cell.getDataValidation().clear();
  const rule: ExcelScript.DataValidationRule = {
    list: {
      inCellDropDown: true,
      source: names.join(",")
    }
  };
  cell.getDataValidation().setRule(rule);
}

function refreshActivityDropdown(workbook: ExcelScript.Workbook): void {
  const dbSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("ACTIVITIES_DB");
  if (!dbSheet) return;
  const table: ExcelScript.Table | undefined = dbSheet.getTable("ActivitiesTable");
  if (!table) return;

  const rows: (string | number | boolean)[][] = table.getRangeBetweenHeaderAndTotal().getValues();
  const names: string[] = [];
  for (const r of rows) {
    if (Boolean(r[2])) {
      names.push(String(r[1]));
    }
  }

  const teSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("TIME_ENTRY");
  if (!teSheet) return;

  const cell: ExcelScript.Range = teSheet.getRange("B4");
  cell.getDataValidation().clear();
  const rule: ExcelScript.DataValidationRule = {
    list: {
      inCellDropDown: true,
      source: names.join(",")
    }
  };
  cell.getDataValidation().setRule(rule);
}

function refreshTeamDropdowns(workbook: ExcelScript.Workbook): void {
  const dbSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("TEAMS_DB");
  if (!dbSheet) return;
  const table: ExcelScript.Table | undefined = dbSheet.getTable("TeamsTable");
  if (!table) return;

  const rows: (string | number | boolean)[][] = table.getRangeBetweenHeaderAndTotal().getValues();
  const names: string[] = ["All"];
  for (const r of rows) {
    names.push(String(r[1]));
  }

  const source: string = names.join(",");
  const rule: ExcelScript.DataValidationRule = {
    list: {
      inCellDropDown: true,
      source: source
    }
  };

  // Update MANAGER_VIEW team dropdown (B2)
  const mvSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("MANAGER_VIEW");
  if (mvSheet) {
    const cell: ExcelScript.Range = mvSheet.getRange("B2");
    cell.getDataValidation().clear();
    cell.getDataValidation().setRule(rule);
  }

  // Update REPORTS team dropdown (B4)
  const rpSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("REPORTS");
  if (rpSheet) {
    const cell: ExcelScript.Range = rpSheet.getRange("B4");
    cell.getDataValidation().clear();
    cell.getDataValidation().setRule(rule);
  }
}
