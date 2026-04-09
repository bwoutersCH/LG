"""Sample data for the time tracking workbook."""
from datetime import date, timedelta

# ── Hash function (must match Utils.ts simpleHash) ───────────────────────────
HASH_SALT = "LG_TimeTrack_2026"

def simple_hash(input_str):
    """Python port of the simpleHash function in Utils.ts."""
    salted = HASH_SALT + input_str + HASH_SALT
    h1 = 0xDEADBEEF
    h2 = 0x41C6CE57
    for ch in salted:
        code = ord(ch)
        h1 = _imul(h1 ^ code, 2654435761) & 0xFFFFFFFF
        h2 = _imul(h2 ^ code, 1597334677) & 0xFFFFFFFF
    h1 = (_imul(h1 ^ (h1 >> 16), 2246822507) ^ _imul(h2 ^ (h2 >> 13), 3266489909)) & 0xFFFFFFFF
    h2 = (_imul(h2 ^ (h2 >> 16), 2246822507) ^ _imul(h1 ^ (h1 >> 13), 3266489909)) & 0xFFFFFFFF
    combined = 4294967296 * (2097151 & h2) + (h1 >> 0)
    return format(combined, '016x')

def _imul(a, b):
    """Emulate Math.imul (32-bit integer multiply)."""
    a = a & 0xFFFFFFFF
    b = b & 0xFFFFFFFF
    ah = (a >> 16) & 0xFFFF
    al = a & 0xFFFF
    bh = (b >> 16) & 0xFFFF
    bl = b & 0xFFFF
    result = (al * bl) + (((ah * bl + al * bh) & 0xFFFF) << 16)
    return result & 0xFFFFFFFF

# Default password for all sample users
DEFAULT_PASSWORD = "Pass1234"
DEFAULT_HASH = simple_hash(DEFAULT_PASSWORD)

# ── Teams ────────────────────────────────────────────────────────────────────
TEAMS = [
    ("T001", "Engineering", "", "U002"),
    ("T002", "Design", "", "U003"),
    ("T003", "Backend", "T001", "U004"),
    ("T004", "Frontend", "T001", "U005"),
]
TEAM_HEADERS = ["TeamID", "TeamName", "ParentTeamID", "OwnerUserID"]

TEAM_MAP = {t[0]: t[1] for t in TEAMS}

# ── Activities ───────────────────────────────────────────────────────────────
ACTIVITIES = [
    ("A001", "Development", True),
    ("A002", "Design", True),
    ("A003", "Testing", True),
    ("A004", "Planning", True),
    ("A005", "Client Meeting", True),
    ("A006", "Admin", True),
]
ACTIVITY_HEADERS = ["ActivityID", "ActivityName", "Active"]
ACTIVITY_NAMES = [a[1] for a in ACTIVITIES]

# ── Projects ─────────────────────────────────────────────────────────────────
PROJECTS = [
    ("P001", "Website Redesign", "T004", True),
    ("P002", "API Platform", "T003", True),
    ("P003", "Mobile App", "T001", True),
    ("P004", "Data Pipeline", "T003", True),
    ("P005", "Brand Refresh", "T002", True),
]
PROJECT_HEADERS = ["ProjectID", "ProjectName", "TeamID", "Active"]
PROJECT_NAMES = [p[1] for p in PROJECTS]

# ── Users (25 total) ────────────────────────────────────────────────────────
USERS = [
    # UserID, FullName, Username, PasswordHash, Role, ManagerID, TeamID, Active
    ("U001", "Alice Admin", "alice.admin", DEFAULT_HASH, "Admin", "", "T001", True),
    ("U002", "Bob Director", "bob.director", DEFAULT_HASH, "Director", "U001", "T001", True),
    ("U003", "Carol Director", "carol.director", DEFAULT_HASH, "Director", "U001", "T002", True),
    ("U004", "Dave Manager", "dave.manager", DEFAULT_HASH, "Manager", "U002", "T003", True),
    ("U005", "Eve Manager", "eve.manager", DEFAULT_HASH, "Manager", "U002", "T004", True),
    ("U006", "Frank Manager", "frank.manager", DEFAULT_HASH, "Manager", "U003", "T002", True),
    ("U007", "Grace Manager", "grace.manager", DEFAULT_HASH, "Manager", "U003", "T002", True),
    ("U008", "Hank TeamLead", "hank.teamlead", DEFAULT_HASH, "TeamLead", "U004", "T003", True),
    ("U009", "Ivy TeamLead", "ivy.teamlead", DEFAULT_HASH, "TeamLead", "U005", "T004", True),
    ("U010", "Jack TeamLead", "jack.teamlead", DEFAULT_HASH, "TeamLead", "U006", "T002", True),
    ("U011", "Kate TeamLead", "kate.teamlead", DEFAULT_HASH, "TeamLead", "U007", "T002", True),
    ("U012", "Leo Employee", "leo.employee", DEFAULT_HASH, "Employee", "U008", "T003", True),
    ("U013", "Mia Employee", "mia.employee", DEFAULT_HASH, "Employee", "U008", "T003", True),
    ("U014", "Noah Employee", "noah.employee", DEFAULT_HASH, "Employee", "U008", "T003", True),
    ("U015", "Olivia Employee", "olivia.employee", DEFAULT_HASH, "Employee", "U009", "T004", True),
    ("U016", "Pete Employee", "pete.employee", DEFAULT_HASH, "Employee", "U009", "T004", True),
    ("U017", "Quinn Employee", "quinn.employee", DEFAULT_HASH, "Employee", "U009", "T004", True),
    ("U018", "Ryan Employee", "ryan.employee", DEFAULT_HASH, "Employee", "U010", "T002", True),
    ("U019", "Sara Employee", "sara.employee", DEFAULT_HASH, "Employee", "U010", "T002", True),
    ("U020", "Tom Employee", "tom.employee", DEFAULT_HASH, "Employee", "U010", "T002", True),
    ("U021", "Uma Employee", "uma.employee", DEFAULT_HASH, "Employee", "U011", "T002", True),
    ("U022", "Vic Employee", "vic.employee", DEFAULT_HASH, "Employee", "U011", "T002", True),
    ("U023", "Wendy Employee", "wendy.employee", DEFAULT_HASH, "Employee", "U004", "T003", True),
    ("U024", "Xander Employee", "xander.employee", DEFAULT_HASH, "Employee", "U005", "T004", True),
    ("U025", "Yara Employee", "yara.employee", DEFAULT_HASH, "Employee", "U006", "T002", True),
]
USER_HEADERS = ["UserID", "FullName", "Username", "PasswordHash", "Role", "ManagerID", "TeamID", "Active"]

# ── Time Entry generation ────────────────────────────────────────────────────
TIME_ENTRY_HEADERS = [
    "EntryID", "Date", "EmployeeID", "EmployeeName", "Team",
    "Project", "Activity", "Hours", "Notes", "SubmittedOn",
    "LastEditedBy", "LastEditedOn",
]

def get_monday(d):
    """Get the Monday of the week containing date d."""
    return d - timedelta(days=d.weekday())

def generate_time_entries():
    """Generate 3 weeks of sample time entries for all non-admin/non-director users."""
    entries = []
    today = date.today()
    current_monday = get_monday(today)
    start_monday = current_monday - timedelta(weeks=2)

    counter = 0
    for user in USERS:
        uid, full_name, username, _, role, _, team_id, _ = user
        if role in ("Admin", "Director"):
            continue
        team_name = TEAM_MAP.get(team_id, "")

        for week_offset in range(3):
            week_start = start_monday + timedelta(weeks=week_offset)
            for day_offset in range(5):  # Mon-Fri
                entry_date = week_start + timedelta(days=day_offset)
                if entry_date > today:
                    continue

                date_str = entry_date.isoformat()
                ts = f"{date_str}T09:00:00.000Z"

                proj1 = PROJECT_NAMES[counter % len(PROJECT_NAMES)]
                act1 = ACTIVITY_NAMES[counter % len(ACTIVITY_NAMES)]
                hours1 = 5 + (counter % 3)
                hours2 = 8 - hours1 + (counter % 2)

                counter += 1
                eid1 = f"TE-INIT-{counter:05d}"
                entries.append((
                    eid1, date_str, uid, full_name, team_name,
                    proj1, act1, hours1, "Sample entry", ts, username, ts
                ))

                if 0 < hours2 <= 16:
                    counter += 1
                    eid2 = f"TE-INIT-{counter:05d}"
                    proj2 = PROJECT_NAMES[(counter + 2) % len(PROJECT_NAMES)]
                    act2 = ACTIVITY_NAMES[(counter + 1) % len(ACTIVITY_NAMES)]
                    entries.append((
                        eid2, date_str, uid, full_name, team_name,
                        proj2, act2, hours2, "Sample entry", ts, username, ts
                    ))

    return entries

# ── Session table ────────────────────────────────────────────────────────────
SESSION_HEADERS = ["UserID", "Username", "Role", "ManagerID", "LoginTime"]
