#!/usr/bin/env python3
"""
Generate individual button scripts for each action.
Each script is a copy of its parent script with main() changed to call the target function.
"""
import os, re

SCRIPTS_DIR = os.path.join(os.path.dirname(os.path.dirname(__file__)), "scripts")
BUTTONS_DIR = os.path.join(SCRIPTS_DIR, "buttons")

# (button_filename, source_script, function_to_call_in_main)
BUTTONS = [
    # LOGIN sheet
    ("Login.ts", "Login.ts", "runLogin"),
    ("Logout.ts", "Login.ts", "runLogout"),
    ("ChangePassword.ts", "Login.ts", "changePassword"),
    # TIME_ENTRY sheet
    ("AddEntry.ts", "TimeEntry.ts", "addEntry"),
    ("FilterMyEntries.ts", "TimeEntry.ts", "filterMyEntries"),
    ("ClearFilters.ts", "TimeEntry.ts", "clearFilters"),
    ("DeleteEntry.ts", "TimeEntry.ts", "deleteEntry"),
    # DAILY_CHECK sheet
    ("ValidateWeek.ts", "TimeEntry.ts", "validateWeek"),
    ("PreviousMonth.ts", "TimeEntry.ts", "previousMonth"),
    ("NextMonth.ts", "TimeEntry.ts", "nextMonth"),
    # MANAGER_VIEW sheet
    ("FilterManagerView.ts", "Reports.ts", "filterManagerView"),
    ("CorrectEntry.ts", "Reports.ts", "correctEntry"),
    # REPORTS sheet
    ("GenerateReport.ts", "Reports.ts", "generateReport"),
    # ADMIN sheet
    ("SaveUser.ts", "Admin.ts", "saveUser"),
    ("SaveProject.ts", "Admin.ts", "saveProject"),
    ("SaveActivity.ts", "Admin.ts", "saveActivity"),
    ("ResetSession.ts", "Admin.ts", "resetSession"),
    ("ExportAllData.ts", "Admin.ts", "exportAllData"),
    # Also include these standalone ones
    ("InitializeWorkbook.ts", "InitializeWorkbook.ts", None),  # already has correct main
    ("RefreshDropdowns.ts", "RefreshDropdowns.ts", None),  # already has correct main
]

def replace_main(source_content, new_function_name, button_name):
    """Replace the main() body to call the target function."""
    # Replace the line inside main that calls the old function
    # Pattern: function main(...) { ... old_call(workbook); ... }
    new_main_body = f"  {new_function_name}(workbook);"

    # Find and replace the main function body
    pattern = r'(function main\(workbook: ExcelScript\.Workbook\): void \{\n)(.*?)(\n\})'

    def replacer(match):
        return match.group(1) + new_main_body + match.group(3)

    result = re.sub(pattern, replacer, source_content, count=1, flags=re.DOTALL)

    # Update the header comment
    old_header_end = result.find("*/")
    if old_header_end > 0:
        header = f"/**\n * {button_name} — Auto-generated button script.\n * Entry point: main() calls {new_function_name}().\n */"
        result = header + result[old_header_end + 2:]

    return result

def main():
    os.makedirs(BUTTONS_DIR, exist_ok=True)

    for button_file, source_file, target_func in BUTTONS:
        source_path = os.path.join(SCRIPTS_DIR, source_file)
        output_path = os.path.join(BUTTONS_DIR, button_file)

        with open(source_path, "r") as f:
            content = f.read()

        if target_func is None:
            # Just copy as-is
            with open(output_path, "w") as f:
                f.write(content)
            print(f"  Copied: {button_file} (from {source_file})")
        else:
            modified = replace_main(content, target_func, button_file.replace(".ts", ""))
            with open(output_path, "w") as f:
                f.write(modified)
            print(f"  Generated: {button_file} (main -> {target_func})")

    print(f"\nDone! {len(BUTTONS)} scripts in {BUTTONS_DIR}")

if __name__ == "__main__":
    main()
