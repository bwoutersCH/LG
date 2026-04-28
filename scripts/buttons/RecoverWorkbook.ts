/**
 * RecoverWorkbook.ts — One-shot recovery script.
 * Use when sheets are stuck as veryHidden or after experimenting with
 * workbook structure protection. Run once, then re-run Login normally.
 *
 * If the workbook is still structure-protected with a password, this
 * script will report it; remove protection via Review > Protect Workbook
 * (entering the password) before running this again.
 */

function main(workbook: ExcelScript.Workbook): void {
  const messages: string[] = [];

  // Try to unprotect (no-arg). If a password is set, this will throw and we
  // tell the user to remove protection manually.
  try {
    const prot: ExcelScript.WorkbookProtection = workbook.getProtection();
    if (prot.getProtected()) {
      try {
        prot.unprotect();
        messages.push("Removed workbook structure protection.");
      } catch (_e) {
        messages.push("Workbook structure is protected with a password. Go to Review > Protect Workbook, enter the password, then run RecoverWorkbook again.");
      }
    }
  } catch (_e) {
    // Older Office Scripts API may not expose this — ignore.
  }

  // Force every sheet to visible so the user can see what's there.
  let unhidden: number = 0;
  for (const s of workbook.getWorksheets()) {
    try {
      if (s.getVisibility() !== ExcelScript.SheetVisibility.visible) {
        s.setVisibility(ExcelScript.SheetVisibility.visible);
        unhidden++;
      }
    } catch (_e) {
      messages.push(`Could not unhide sheet "${s.getName()}" (likely structure protection).`);
    }
  }
  messages.push(`Unhid ${unhidden} sheet(s).`);

  // Surface the report somewhere the user will see it.
  const loginSheet: ExcelScript.Worksheet | undefined = workbook.getWorksheet("LOGIN");
  if (loginSheet) {
    loginSheet.activate();
    const cell: ExcelScript.Range = loginSheet.getRange("B6");
    cell.setValue(messages.join(" "));
    cell.getFormat().getFont().setColor("#239A98");
  }
  console.log(messages.join("\n"));
}
