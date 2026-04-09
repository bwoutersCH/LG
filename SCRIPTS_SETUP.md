# Office Scripts Setup Guide

## How to Install Scripts

1. Open `TimeTracking.xlsx` in Excel for the Web (SharePoint).
2. Go to **Automate** tab > **New Script**.
3. Copy-paste each `.ts` file's content from the `scripts/` folder.
4. Save each script with the name shown below.
5. Assign scripts to button cells: right-click the orange button cell > **Assign Script**.

---

## Script Reference

| Script File | Function | Trigger (Button/Sheet) | Description |
|---|---|---|---|
| **Utils.ts** | *(shared functions)* | Imported by all scripts | Shared utilities: hash, session management, hierarchy resolution, date helpers |
| **Login.ts** | `runLogin()` | LOGIN sheet > "Log In" button | Validates credentials, writes session, shows role-appropriate sheets |
| **Login.ts** | `runLogout()` | LOGIN sheet > "Log Out" button | Clears session, hides all sheets, returns to LOGIN |
| **Login.ts** | `refreshDailyCheck()` | Called internally after login | Populates DAILY_CHECK with current user's weekly hours |
| **TimeEntry.ts** | `addEntry()` | TIME_ENTRY sheet > "Add Entry" button | Validates input, checks duplicates, appends to TimeEntryTable |
| **TimeEntry.ts** | `filterMyEntries()` | TIME_ENTRY sheet > "My Entries This Week" button | Filters table to show only current user's entries for this week |
| **TimeEntry.ts** | `clearFilters()` | TIME_ENTRY sheet > "Clear Filters" button | Removes all filters from the TimeEntryTable |
| **TimeEntry.ts** | `validateWeek()` | DAILY_CHECK sheet > "Validate Week" button | Checks all weekdays have >= 8h; highlights gaps |
| **TimeEntry.ts** | `deleteEntry()` | TIME_ENTRY sheet > "Delete Entry" button | Deletes entry by ID (permission-checked) |
| **Reports.ts** | `filterManagerView()` | MANAGER_VIEW sheet > "Apply Filter" button | Filters subtree entries by team, employee, date range |
| **Reports.ts** | `correctEntry()` | MANAGER_VIEW sheet > "Save Correction" button | Updates hours on an entry with reason (permission-checked) |
| **Reports.ts** | `generateReport()` | REPORTS sheet > "Generate Report" button | Builds multi-section report by period and breakdown dimension |
| **Admin.ts** | `saveUser()` | ADMIN sheet > "Save User" button | Creates or updates a user in Users_DB |
| **Admin.ts** | `saveProject()` | ADMIN sheet > "Save Project" button | Creates or updates a project in Projects_DB |
| **Admin.ts** | `saveActivity()` | ADMIN sheet > "Save Activity" button | Creates or updates an activity in Activities_DB |
| **Admin.ts** | `resetSession()` | ADMIN sheet > "Reset Session" button | Clears session and returns to LOGIN (for testing) |
| **Admin.ts** | `exportAllData()` | ADMIN sheet > "Export All Data" button | Copies all time entries to a new EXPORT sheet |
| **InitializeWorkbook.ts** | `initializeWorkbook()` | ADMIN sheet > "Initialize Sample Data" button | One-time: populates all DB tables with 25 users, 5 projects, 6 activities, 4 teams |

---

## Important Notes

- **Utils.ts must be loaded first** — all other scripts depend on its functions.
- In Excel for the Web, Office Scripts share a global scope per workbook, so all functions from all scripts are available once loaded.
- **Run InitializeWorkbook.ts only once** — sample data is already pre-loaded in the .xlsx file. Running again will create duplicate entries.
- Button cells in the workbook are styled as orange placeholders. Assign the corresponding script function to each via Automate > Assign Script.
- Sheet protection password: `TT#2026!Sec` (defined in Utils.ts as `SHEET_PASSWORD`).
- Default user password: `Pass1234` for all 25 sample users.
