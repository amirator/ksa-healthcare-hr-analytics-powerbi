"""
Builds a complete Power BI Project (PBIP) for Contoso Health KSA - HR Analytics.
Generates: embedded sample data (deflate+base64 M queries), model.bim (TMSL),
report.json (6 pages, ~60 visuals), custom theme, .pbip/.pbir/.pbism/.platform files.
Run:  python build_pbip.py
"""
import base64, json, os, random, uuid, zlib
from datetime import date, timedelta

random.seed(42)
ROOT = os.path.dirname(os.path.abspath(__file__))
NAME = "KSA Healthcare HR Analytics"
TODAY = date(2026, 6, 11)

# ----------------------------------------------------------------------------
# 1. DATA GENERATION
# ----------------------------------------------------------------------------
FACILITIES = [
    ("Contoso Hospital - Sharafiyah",      "Jeddah", "Western", "Hospital",       0.20),
    ("Contoso Medical Center - Aziziyah",  "Jeddah", "Western", "Medical Center", 0.12),
    ("Contoso Medical Center - Sanaiyah",  "Jeddah", "Western", "Medical Center", 0.12),
    ("Contoso Medical Center - Bawadi",    "Jeddah", "Western", "Medical Center", 0.09),
    ("Contoso Hospital - Riyadh",          "Riyadh", "Central", "Hospital",       0.18),
    ("Contoso Medical Center - Batha",     "Riyadh", "Central", "Medical Center", 0.09),
    ("Contoso Medical Center - Dammam",    "Dammam", "Eastern", "Medical Center", 0.10),
    ("Contoso Medical Center - Makkah",    "Makkah", "Western", "Medical Center", 0.10),
]
# name, division, weight, (salary min,max), saudi share, female share, job titles
DEPARTMENTS = [
    ("Physicians",             "Clinical",     0.14, (16000, 42000), 0.15, 0.35,
     ["Consultant", "Specialist Physician", "General Practitioner", "Resident Physician"]),
    ("Nursing",                "Clinical",     0.25, (4500, 9500),   0.08, 0.82,
     ["Staff Nurse", "Charge Nurse", "Nurse Supervisor", "Midwife"]),
    ("Pharmacy",               "Clinical",     0.07, (6000, 13000),  0.30, 0.45,
     ["Pharmacist", "Pharmacy Technician", "Pharmacy Supervisor"]),
    ("Laboratory",             "Clinical",     0.06, (5500, 11000),  0.25, 0.40,
     ["Lab Technologist", "Lab Technician", "Lab Supervisor"]),
    ("Radiology",              "Clinical",     0.05, (6000, 12500),  0.22, 0.30,
     ["Radiology Technologist", "Sonographer", "Radiology Supervisor"]),
    ("Dental",                 "Clinical",     0.05, (9000, 25000),  0.15, 0.40,
     ["Dentist", "Dental Assistant", "Dental Hygienist"]),
    ("Physiotherapy",          "Clinical",     0.04, (5500, 10500),  0.15, 0.45,
     ["Physiotherapist", "Physiotherapy Assistant"]),
    ("Administration",         "Non-Clinical", 0.08, (5000, 16000),  0.65, 0.35,
     ["Admin Officer", "Operations Coordinator", "Branch Manager", "Executive Assistant"]),
    ("Human Resources",        "Non-Clinical", 0.04, (6000, 15000),  0.75, 0.45,
     ["HR Officer", "Recruitment Specialist", "Payroll Officer", "HR Business Partner"]),
    ("Finance",                "Non-Clinical", 0.04, (7000, 17000),  0.55, 0.25,
     ["Accountant", "Financial Analyst", "Insurance Coordinator", "Billing Officer"]),
    ("Information Technology", "Non-Clinical", 0.04, (7000, 16000),  0.45, 0.20,
     ["IT Support Engineer", "Systems Administrator", "HIS Analyst", "Network Engineer"]),
    ("Customer Service",       "Non-Clinical", 0.08, (4000, 7500),   0.70, 0.55,
     ["Patient Relations Officer", "Receptionist", "Call Center Agent"]),
    ("Support Services",       "Non-Clinical", 0.06, (3000, 6000),   0.10, 0.15,
     ["Housekeeping Staff", "Porter", "Maintenance Technician", "Driver"]),
]
NATIONALITIES = [("India", 0.34), ("Philippines", 0.22), ("Egypt", 0.14), ("Pakistan", 0.08),
                 ("Sudan", 0.07), ("Jordan", 0.06), ("Yemen", 0.05), ("Bangladesh", 0.04)]
NAME_POOLS = {
    "Saudi": {
        "m": ["Mohammed", "Abdullah", "Fahad", "Khalid", "Saud", "Faisal", "Turki", "Nasser", "Bandar", "Majed", "Sultan", "Rakan"],
        "f": ["Noura", "Sara", "Reem", "Lama", "Aljohara", "Hessa", "Munira", "Dalal", "Shahad", "Ghada", "Amal", "Jawaher"],
        "l": ["Al-Qahtani", "Al-Otaibi", "Al-Ghamdi", "Al-Harbi", "Al-Zahrani", "Al-Shehri", "Al-Dossari", "Al-Mutairi", "Al-Subaie", "Al-Amri"]},
    "India": {
        "m": ["Rajesh", "Suresh", "Anil", "Vivek", "Pradeep", "Arun", "Sanjay", "Manoj", "Abdul", "Faisal", "Nithin", "Jithin"],
        "f": ["Priya", "Anitha", "Deepa", "Sreeja", "Lakshmi", "Reshma", "Neha", "Asha", "Jisha", "Divya", "Smitha", "Remya"],
        "l": ["Nair", "Menon", "Kumar", "Pillai", "Varghese", "Thomas", "Sharma", "Khan", "Joseph", "Iyer", "Mathew", "Rahman"]},
    "Philippines": {
        "m": ["Jose", "Mark", "John Paul", "Ramon", "Carlo", "Dennis", "Ryan", "Joel", "Allan", "Christian"],
        "f": ["Maria", "Angelica", "Kristine", "Joanna", "Liza", "Cherry", "Grace", "Rowena", "Michelle", "Janine", "Aileen", "Catherine"],
        "l": ["Santos", "Reyes", "Cruz", "Garcia", "Dela Cruz", "Mendoza", "Flores", "Villanueva", "Ramos", "Aquino"]},
    "Egypt": {
        "m": ["Ahmed", "Mostafa", "Hossam", "Tarek", "Karim", "Amr", "Sherif", "Mahmoud", "Wael", "Hany"],
        "f": ["Mona", "Heba", "Dina", "Rania", "Yasmin", "Nesma", "Salma", "Marwa", "Eman", "Aya"],
        "l": ["Hassan", "Ibrahim", "El-Sayed", "Abdelaziz", "Fawzy", "Saleh", "Mansour", "Farouk", "Gamal", "Shawky"]},
    "Pakistan": {
        "m": ["Imran", "Asif", "Bilal", "Usman", "Kashif", "Adnan", "Shahid", "Tariq", "Waqas", "Zeeshan"],
        "f": ["Ayesha", "Fatima", "Sana", "Hira", "Mariam", "Nadia", "Saima", "Rabia", "Zainab", "Bushra"],
        "l": ["Khan", "Malik", "Ahmed", "Hussain", "Sheikh", "Butt", "Qureshi", "Raza", "Iqbal", "Chaudhry"]},
    "Sudan": {
        "m": ["Osman", "Elfatih", "Mubarak", "Siddig", "Babiker", "Hatim", "Muzamil", "Awad"],
        "f": ["Amani", "Ishraga", "Tasneem", "Wifag", "Nahla", "Samia", "Hanan", "Rasha"],
        "l": ["Abdelrahman", "Elamin", "Osman", "Ahmed", "Idris", "Hamid", "Babiker", "Suliman"]},
    "Jordan": {
        "m": ["Omar", "Laith", "Yazan", "Hamzeh", "Mutaz", "Zaid", "Anas", "Qais"],
        "f": ["Rana", "Dana", "Lina", "Tala", "Haneen", "Razan", "Aseel", "Farah"],
        "l": ["Haddad", "Khasawneh", "Momani", "Obeidat", "Shawabkeh", "Zoubi", "Tarawneh", "Nsour"]},
    "Yemen": {
        "m": ["Saleh", "Yahya", "Adel", "Munir", "Fuad", "Sami", "Akram", "Riyadh"],
        "f": ["Arwa", "Bushra", "Sumaya", "Najla", "Wafa", "Iman", "Huda", "Sabreen"],
        "l": ["Al-Hadi", "Al-Sanabani", "Al-Maqtari", "Al-Sharabi", "Al-Absi", "Al-Hammadi", "Saif", "Numan"]},
    "Bangladesh": {
        "m": ["Rafiq", "Kamal", "Shahin", "Jahangir", "Monir", "Habib", "Faruk", "Nazrul"],
        "f": ["Nasrin", "Shirin", "Rina", "Salma", "Taslima", "Ruma", "Khadija", "Shilpi"],
        "l": ["Islam", "Hossain", "Rahman", "Uddin", "Ahmed", "Miah", "Chowdhury", "Sarker"]},
}
TERM_REASONS_VOL = ["Resignation - Better Offer", "Resignation - Relocation", "Resignation - Family Reasons", "End of Contract (Not Renewed by Employee)"]
TERM_REASONS_INVOL = ["End of Contract", "Performance", "Restructuring", "Retirement"]
GRADES = ["G1 - Entry", "G2 - Professional", "G3 - Senior", "G4 - Supervisory", "G5 - Management"]
SOURCES = [("Company Website", 0.18), ("LinkedIn", 0.22), ("Recruitment Agency", 0.20),
           ("Employee Referral", 0.16), ("TAQAT / Hadaf", 0.14), ("Job Fair", 0.10)]
COURSES = [("BLS Certification", "Life Support", 8, 350), ("ACLS Certification", "Life Support", 16, 900),
           ("Infection Prevention & Control", "Clinical", 6, 200), ("Patient Safety Essentials", "Clinical", 6, 200),
           ("CBAHI Accreditation Readiness", "Compliance", 8, 250), ("Saudi Labor Law Essentials", "Compliance", 4, 150),
           ("Fire Safety & Evacuation", "Compliance", 3, 100), ("Leadership in Healthcare", "Leadership", 16, 1500),
           ("Effective Communication with Patients", "Customer Experience", 4, 180), ("Arabic for Healthcare Staff", "Customer Experience", 20, 600),
           ("HIS / Electronic Medical Records", "Technical", 8, 300), ("Medication Safety", "Clinical", 5, 220),
           ("Emergency Triage", "Clinical", 8, 400), ("Excel for HR & Finance", "Technical", 8, 250)]
LEAVE_TYPES = [("Annual", 0.46), ("Sick", 0.28), ("Emergency", 0.10), ("Unpaid", 0.06),
               ("Maternity", 0.05), ("Hajj", 0.05)]

def wchoice(pairs):
    r, acc = random.random(), 0.0
    for v, w in pairs:
        acc += w
        if r <= acc: return v
    return pairs[-1][0]

def rand_date(d1, d2):
    return d1 + timedelta(days=random.randint(0, (d2 - d1).days))

def iso(d): return d.isoformat() if d else None

employees, used_ids = [], set()
N_EMP = 680
fac_pairs = [(f[0], f[4]) for f in FACILITIES]
dep_pairs = [(d[0], d[2]) for d in DEPARTMENTS]
dep_map = {d[0]: d for d in DEPARTMENTS}
fac_map = {f[0]: f for f in FACILITIES}

for i in range(N_EMP):
    dept = wchoice(dep_pairs)
    dname, division, _, (smin, smax), saudi_share, female_share, titles = dep_map[dept]
    facility = wchoice(fac_pairs)
    is_saudi = random.random() < saudi_share
    nationality = "Saudi" if is_saudi else wchoice(NATIONALITIES)
    gender = "Female" if random.random() < female_share else "Male"
    pool = NAME_POOLS[nationality]
    first = random.choice(pool["f" if gender == "Female" else "m"])
    last = random.choice(pool["l"])
    hire = rand_date(date(2014, 1, 1), date(2026, 4, 30))
    # ~26% terminated
    term, reason, ttype = None, None, None
    if random.random() < 0.26:
        earliest = hire + timedelta(days=120)
        if earliest < TODAY:
            term = rand_date(earliest, TODAY)
            if random.random() < 0.68:
                ttype, reason = "Voluntary", random.choice(TERM_REASONS_VOL)
            else:
                ttype, reason = "Involuntary", random.choice(TERM_REASONS_INVOL)
    age = random.randint(24, 58)
    birth = TODAY - timedelta(days=age * 365 + random.randint(0, 364))
    salary = int(round(random.uniform(smin, smax), -2))
    grade_idx = min(4, max(0, int((salary - smin) / max(1, smax - smin) * 4) + (1 if random.random() < .15 else 0)))
    rating = round(min(5.0, max(1.5, random.gauss(3.7, 0.6))), 1)
    etype = wchoice([("Full-time", 0.88), ("Part-time", 0.07), ("Locum", 0.05)])
    tenure_ref = term or TODAY
    tenure_y = (tenure_ref - hire).days / 365.25
    tband = ("< 1 yr" if tenure_y < 1 else "1-3 yrs" if tenure_y < 3 else "3-5 yrs" if tenure_y < 5
             else "5-10 yrs" if tenure_y < 10 else "10+ yrs")
    aband = ("< 25" if age < 25 else "25-34" if age < 35 else "35-44" if age < 45 else "45-54" if age < 55 else "55+")
    eid = f"AMG-{2014 + (hire.year - 2014):04d}-{i+1:04d}"
    employees.append([eid, f"{first} {last}", gender, nationality,
                      "Saudi" if is_saudi else "Non-Saudi", dname, random.choice(titles),
                      facility, iso(hire), iso(term), reason, ttype, str(salary),
                      iso(birth), aband, tband, GRADES[grade_idx], str(rating), etype])

active_emps = [e for e in employees if e[9] is None]

# Recruitment
recruitment = []
for i in range(170):
    dept = wchoice(dep_pairs)
    facility = wchoice(fac_pairs)
    posting = rand_date(date(2024, 1, 1), date(2026, 5, 20))
    status = wchoice([("Filled", 0.62), ("Open", 0.28), ("Cancelled", 0.10)])
    if status == "Open" and posting < date(2025, 9, 1):
        posting = rand_date(date(2025, 9, 1), date(2026, 5, 20))
    days_to_fill, filled = None, None
    apps = random.randint(8, 140)
    offers = max(1, int(apps * random.uniform(0.03, 0.12)))
    accepted = max(0, offers - random.randint(0, max(1, offers // 2)))
    if status == "Filled":
        days_to_fill = random.randint(18, 130)
        filled = posting + timedelta(days=days_to_fill)
        accepted = max(1, accepted)
    elif status == "Cancelled":
        offers, accepted = 0, 0
    title = random.choice(dep_map[dept][6])
    recruitment.append([f"REQ-{posting.year}-{i+1:03d}", title, dept, facility, iso(posting),
                        iso(filled), status, str(days_to_fill) if days_to_fill else None,
                        str(apps), str(offers), str(accepted), wchoice(SOURCES)])

# Training
training = []
for i in range(620):
    emp = random.choice(employees)
    course, cat, hrs, cost = random.choice(COURSES)
    tdate = rand_date(max(date(2025, 1, 2), date.fromisoformat(emp[8])), date(2026, 6, 5))
    status = wchoice([("Completed", 0.86), ("In Progress", 0.08), ("No Show", 0.06)])
    training.append([f"TRN-{i+1:04d}", emp[0], course, cat, iso(tdate), str(hrs),
                     str(cost + random.randint(-30, 60)), status])

# Leave
leave = []
for i in range(760):
    emp = random.choice(employees)
    ltype = wchoice(LEAVE_TYPES)
    if emp[2] == "Male" and ltype == "Maternity":
        ltype = "Annual"
    start = rand_date(max(date(2025, 1, 2), date.fromisoformat(emp[8])), date(2026, 6, 5))
    days = {"Annual": random.randint(5, 22), "Sick": random.randint(1, 7), "Emergency": random.randint(1, 3),
            "Unpaid": random.randint(3, 15), "Maternity": 70, "Hajj": 10}[ltype]
    leave.append([f"LV-{i+1:04d}", emp[0], ltype, iso(start), str(days)])

dim_department = [[d[0], d[1]] for d in DEPARTMENTS]
dim_facility = [[f[0], f[1], f[2], f[3]] for f in FACILITIES]

# ----------------------------------------------------------------------------
# 2. M QUERY BUILDER (Enter-Data style: deflate + base64 JSON rows)
# ----------------------------------------------------------------------------
def m_query(rows, columns, type_changes):
    payload = json.dumps(rows, separators=(",", ":")).encode("utf-8")
    co = zlib.compressobj(9, zlib.DEFLATED, -15)
    b64 = base64.b64encode(co.compress(payload) + co.flush()).decode("ascii")
    col_sig = ", ".join(f'#"{c}" = _t' for c in columns)
    tc = ", ".join('{"%s", %s}' % (c, t) for c, t in type_changes)
    src = f'    Source = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("{b64}", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [{col_sig}])'
    if type_changes:
        lines = ["let", src + ",",
                 f'    #"Changed Type" = Table.TransformColumnTypes(Source,{{{tc}}}, "en-US")',
                 "in", '    #"Changed Type"']
    else:
        lines = ["let", src, "in", "    Source"]
    return lines

def tag(): return str(uuid.uuid4())

def col(name, dtype, fmt=None, summarize="none", hidden=False, sort_by=None, is_key=False, calc=False):
    c = {"name": name, "dataType": dtype, "sourceColumn": name, "lineageTag": tag(),
         "summarizeBy": summarize,
         "annotations": [{"name": "SummarizationSetBy", "value": "Automatic"}]}
    if fmt: c["formatString"] = fmt
    if hidden: c["isHidden"] = True
    if sort_by: c["sortByColumn"] = sort_by
    if is_key: c["isKey"] = True
    if calc:
        c["type"] = "calculatedTableColumn"
        c["isNameInferred"] = True
        c["sourceColumn"] = f"[{name}]"
        c["isDataTypeInferred"] = True
    return c

def m_table(name, rows, cols_spec, hidden=False):
    """cols_spec: list of (colname, m_type, model_dtype, fmt, summarize)"""
    columns = [c[0] for c in cols_spec]
    type_changes = [(c[0], c[1]) for c in cols_spec if c[1] != "type text"]
    t = {"name": name, "lineageTag": tag(),
         "columns": [col(c[0], c[2], fmt=c[3], summarize=c[4]) for c in cols_spec],
         "partitions": [{"name": name, "mode": "import",
                         "source": {"type": "m", "expression": m_query(rows, columns, type_changes)}}],
         "annotations": [{"name": "PBI_ResultType", "value": "Table"}]}
    if hidden: t["isHidden"] = True
    return t

T = "type text"; D = "type date"; I = "Int64.Type"; N = "type number"

tbl_employees = m_table("Employees", employees, [
    ("EmployeeID", T, "string", None, "none"), ("FullName", T, "string", None, "none"),
    ("Gender", T, "string", None, "none"), ("Nationality", T, "string", None, "none"),
    ("NationalityGroup", T, "string", None, "none"), ("Department", T, "string", None, "none"),
    ("JobTitle", T, "string", None, "none"), ("Facility", T, "string", None, "none"),
    ("HireDate", D, "dateTime", "dd-mmm-yyyy", "none"), ("TerminationDate", D, "dateTime", "dd-mmm-yyyy", "none"),
    ("TerminationReason", T, "string", None, "none"), ("TerminationType", T, "string", None, "none"),
    ("MonthlySalarySAR", I, "int64", "#,0", "none"), ("BirthDate", D, "dateTime", "dd-mmm-yyyy", "none"),
    ("AgeBand", T, "string", None, "none"), ("TenureBand", T, "string", None, "none"),
    ("Grade", T, "string", None, "none"), ("PerformanceRating", N, "double", "0.0", "none"),
    ("EmploymentType", T, "string", None, "none")])

tbl_recruitment = m_table("Recruitment", recruitment, [
    ("RequisitionID", T, "string", None, "none"), ("Position", T, "string", None, "none"),
    ("Department", T, "string", None, "none"), ("Facility", T, "string", None, "none"),
    ("PostingDate", D, "dateTime", "dd-mmm-yyyy", "none"), ("FilledDate", D, "dateTime", "dd-mmm-yyyy", "none"),
    ("Status", T, "string", None, "none"), ("DaysToFill", I, "int64", "#,0", "none"),
    ("Applications", I, "int64", "#,0", "sum"), ("OffersMade", I, "int64", "#,0", "sum"),
    ("OffersAccepted", I, "int64", "#,0", "sum"), ("Source", T, "string", None, "none")])

tbl_training = m_table("Training", training, [
    ("TrainingID", T, "string", None, "none"), ("EmployeeID", T, "string", None, "none"),
    ("Course", T, "string", None, "none"), ("Category", T, "string", None, "none"),
    ("Date", D, "dateTime", "dd-mmm-yyyy", "none"), ("Hours", I, "int64", "#,0", "sum"),
    ("CostSAR", I, "int64", "#,0", "sum"), ("CompletionStatus", T, "string", None, "none")])

tbl_leave = m_table("Leave", leave, [
    ("LeaveID", T, "string", None, "none"), ("EmployeeID", T, "string", None, "none"),
    ("LeaveType", T, "string", None, "none"), ("StartDate", D, "dateTime", "dd-mmm-yyyy", "none"),
    ("Days", I, "int64", "#,0", "sum")])

tbl_dim_dept = m_table("DimDepartment", dim_department, [
    ("Department", T, "string", None, "none"), ("Division", T, "string", None, "none")])

tbl_dim_fac = m_table("DimFacility", dim_facility, [
    ("Facility", T, "string", None, "none"), ("City", T, "string", None, "none"),
    ("Region", T, "string", None, "none"), ("FacilityType", T, "string", None, "none")])

# Date table (DAX calculated)
date_cols = [
    col("Date", "dateTime", fmt="dd-mmm-yyyy", is_key=True, calc=True),
    col("Year", "int64", fmt="0", calc=True),
    col("Month Number", "int64", fmt="0", hidden=True, calc=True),
    col("Month", "string", sort_by="Month Number", calc=True),
    col("MonthStart", "dateTime", fmt="mmm yyyy", calc=True),
    col("Quarter", "string", calc=True),
    col("Year-Month", "string", calc=True),
]
tbl_date = {"name": "Date", "lineageTag": tag(), "dataCategory": "Time",
    "columns": date_cols,
    "partitions": [{"name": "Date", "mode": "import", "source": {"type": "calculated", "expression": [
        "ADDCOLUMNS (",
        "    CALENDAR ( DATE ( 2020, 1, 1 ), DATE ( 2026, 6, 30 ) ),",
        '    "Year", YEAR ( [Date] ),',
        '    "Month Number", MONTH ( [Date] ),',
        '    "Month", FORMAT ( [Date], "MMM" ),',
        '    "MonthStart", DATE ( YEAR ( [Date] ), MONTH ( [Date] ), 1 ),',
        '    "Quarter", "Q" & FORMAT ( [Date], "Q" ),',
        '    "Year-Month", FORMAT ( [Date], "YYYY-MM" )',
        ")"]}}],
    "annotations": [{"name": "PBI_ResultType", "value": "Table"}]}

# Measures
def M(name, expr, fmt=None, folder=None):
    m = {"name": name, "expression": expr if isinstance(expr, list) else [expr], "lineageTag": tag()}
    if fmt: m["formatString"] = fmt
    if folder: m["displayFolder"] = folder
    return m

ACTIVE_FILTER = ("FILTER ( Employees, Employees[HireDate] <= _d && "
                 "( ISBLANK ( Employees[TerminationDate] ) || Employees[TerminationDate] > _d ) )")

measures = [
    M("Headcount", [" ", "VAR _d = MAX ( 'Date'[Date] )", "RETURN", f"    COUNTROWS ( {ACTIVE_FILTER} )"], "#,0", "Workforce"),
    M("Saudi Headcount", "CALCULATE ( [Headcount], Employees[NationalityGroup] = \"Saudi\" )", "#,0", "Saudization"),
    M("Non-Saudi Headcount", "[Headcount] - [Saudi Headcount]", "#,0", "Saudization"),
    M("Saudization %", "DIVIDE ( [Saudi Headcount], [Headcount] )", "0.0%", "Saudization"),
    M("Saudization Target", "0.40", "0%", "Saudization"),
    M("Female Headcount", "CALCULATE ( [Headcount], Employees[Gender] = \"Female\" )", "#,0", "Workforce"),
    M("Female Ratio %", "DIVIDE ( [Female Headcount], [Headcount] )", "0.0%", "Workforce"),
    M("Hires", "COUNTROWS ( FILTER ( Employees, Employees[HireDate] >= MIN ( 'Date'[Date] ) && Employees[HireDate] <= MAX ( 'Date'[Date] ) ) )", "#,0", "Attrition"),
    M("Terminations", "COUNTROWS ( FILTER ( Employees, NOT ISBLANK ( Employees[TerminationDate] ) && Employees[TerminationDate] >= MIN ( 'Date'[Date] ) && Employees[TerminationDate] <= MAX ( 'Date'[Date] ) ) )", "#,0", "Attrition"),
    M("Voluntary Terminations", "CALCULATE ( [Terminations], Employees[TerminationType] = \"Voluntary\" )", "#,0", "Attrition"),
    M("Voluntary Share %", "DIVIDE ( [Voluntary Terminations], [Terminations] )", "0.0%", "Attrition"),
    M("Attrition Rate %", "DIVIDE ( [Terminations], AVERAGEX ( VALUES ( 'Date'[MonthStart] ), [Headcount] ) )", "0.0%", "Attrition"),
    M("Net Headcount Change", "[Hires] - [Terminations]", "+#,0;-#,0;0", "Attrition"),
    M("Avg Tenure (Years)", [" ", "VAR _d = MAX ( 'Date'[Date] )", "RETURN", f"    AVERAGEX ( {ACTIVE_FILTER}, DATEDIFF ( Employees[HireDate], _d, DAY ) / 365.25 )"], "0.0", "Workforce"),
    M("Avg Age", [" ", "VAR _d = MAX ( 'Date'[Date] )", "RETURN", f"    AVERAGEX ( {ACTIVE_FILTER}, DATEDIFF ( Employees[BirthDate], _d, DAY ) / 365.25 )"], "0.0", "Workforce"),
    M("Monthly Payroll (SAR)", [" ", "VAR _d = MAX ( 'Date'[Date] )", "RETURN", f"    SUMX ( {ACTIVE_FILTER}, Employees[MonthlySalarySAR] )"], "#,0", "Compensation"),
    M("Avg Monthly Salary (SAR)", [" ", "VAR _d = MAX ( 'Date'[Date] )", "RETURN", f"    AVERAGEX ( {ACTIVE_FILTER}, Employees[MonthlySalarySAR] )"], "#,0", "Compensation"),
    M("Median Salary (SAR)", [" ", "VAR _d = MAX ( 'Date'[Date] )", "RETURN", f"    MEDIANX ( {ACTIVE_FILTER}, Employees[MonthlySalarySAR] )"], "#,0", "Compensation"),
    M("Avg Performance Rating", [" ", "VAR _d = MAX ( 'Date'[Date] )", "RETURN", f"    AVERAGEX ( {ACTIVE_FILTER}, Employees[PerformanceRating] )"], "0.00", "Performance"),
    M("High Performers %", [" ", "VAR _d = MAX ( 'Date'[Date] )", "VAR _act = " + ACTIVE_FILTER, "RETURN", "    DIVIDE ( COUNTROWS ( FILTER ( _act, Employees[PerformanceRating] >= 4.5 ) ), COUNTROWS ( _act ) )"], "0.0%", "Performance"),
    M("Open Positions", "CALCULATE ( COUNTROWS ( Recruitment ), Recruitment[Status] = \"Open\" )", "#,0", "Recruitment"),
    M("Positions Filled", "CALCULATE ( COUNTROWS ( Recruitment ), Recruitment[Status] = \"Filled\" )", "#,0", "Recruitment"),
    M("Avg Time to Fill (Days)", "CALCULATE ( AVERAGE ( Recruitment[DaysToFill] ), Recruitment[Status] = \"Filled\" )", "0.0", "Recruitment"),
    M("Total Applications", "SUM ( Recruitment[Applications] )", "#,0", "Recruitment"),
    M("Offers Made", "SUM ( Recruitment[OffersMade] )", "#,0", "Recruitment"),
    M("Offers Accepted", "SUM ( Recruitment[OffersAccepted] )", "#,0", "Recruitment"),
    M("Offer Acceptance %", "DIVIDE ( [Offers Accepted], [Offers Made] )", "0.0%", "Recruitment"),
    M("Applications per Hire", "DIVIDE ( [Total Applications], [Positions Filled] )", "0.0", "Recruitment"),
    M("Training Hours", "SUM ( Training[Hours] )", "#,0", "Training"),
    M("Training Cost (SAR)", "SUM ( Training[CostSAR] )", "#,0", "Training"),
    M("Trainings Delivered", "COUNTROWS ( Training )", "#,0", "Training"),
    M("Training Completion %", "DIVIDE ( CALCULATE ( COUNTROWS ( Training ), Training[CompletionStatus] = \"Completed\" ), COUNTROWS ( Training ) )", "0.0%", "Training"),
    M("Training Hours per Employee", "DIVIDE ( [Training Hours], [Headcount] )", "0.0", "Training"),
    M("Leave Days", "SUM ( 'Leave'[Days] )", "#,0", "Leave"),
    M("Sick Leave Days", "CALCULATE ( SUM ( 'Leave'[Days] ), 'Leave'[LeaveType] = \"Sick\" )", "#,0", "Leave"),
    M("Absenteeism %", "DIVIDE ( [Sick Leave Days], [Headcount] * COUNTROWS ( VALUES ( 'Date'[MonthStart] ) ) * 22 )", "0.00%", "Leave"),
]

tbl_measures = {"name": "HR Measures", "lineageTag": tag(),
    "columns": [col("Col", "int64", hidden=True)],
    "partitions": [{"name": "HR Measures", "mode": "import",
                    "source": {"type": "m", "expression": [
                        "let", "    Source = #table ( type table [ Col = Int64.Type ], {} )", "in", "    Source"]}}],
    "measures": measures,
    "annotations": [{"name": "PBI_ResultType", "value": "Table"}]}

def rel(ft, fc, tt, tc, active=True):
    r = {"name": str(uuid.uuid4()), "fromTable": ft, "fromColumn": fc, "toTable": tt, "toColumn": tc}
    if not active: r["isActive"] = False
    return r

relationships = [
    rel("Leave", "EmployeeID", "Employees", "EmployeeID"),
    rel("Training", "EmployeeID", "Employees", "EmployeeID"),
    rel("Leave", "StartDate", "Date", "Date"),
    rel("Training", "Date", "Date", "Date"),
    rel("Recruitment", "PostingDate", "Date", "Date"),
    rel("Employees", "Department", "DimDepartment", "Department"),
    rel("Recruitment", "Department", "DimDepartment", "Department"),
    rel("Employees", "Facility", "DimFacility", "Facility"),
    rel("Recruitment", "Facility", "DimFacility", "Facility"),
]

model_bim = {
    "name": str(uuid.uuid4()),
    "compatibilityLevel": 1550,
    "model": {
        "culture": "en-US",
        "dataAccessOptions": {"legacyRedirects": True, "returnErrorValuesAsNull": True},
        "defaultPowerBIDataSourceVersion": "powerBI_V3",
        "sourceQueryCulture": "en-US",
        "tables": [tbl_employees, tbl_recruitment, tbl_training, tbl_leave,
                   tbl_dim_dept, tbl_dim_fac, tbl_date, tbl_measures],
        "relationships": relationships,
        "cultures": [{"name": "en-US", "linguisticMetadata": {"content": {"Language": "en-US", "Version": "1.0.0"}, "contentType": "json"}}],
        "annotations": [
            {"name": "PBI_QueryOrder", "value": json.dumps(["Employees", "Recruitment", "Training", "Leave", "DimDepartment", "DimFacility", "HR Measures"])},
            {"name": "__PBI_TimeIntelligenceEnabled", "value": "0"},
            {"name": "PBIDesktopVersion", "value": "2.130.0.0 (Main)"},
        ],
    },
}

# ----------------------------------------------------------------------------
# 3. THEME
# ----------------------------------------------------------------------------
TEAL = "#0E7C7B"; NAVY = "#0B4F6C"; LIGHT_BG = "#F3F7F7"; ACCENT = "#F4A259"
theme = {
    "name": "Contoso Health KSA",
    "dataColors": [TEAL, NAVY, "#17A398", "#5AB9A8", ACCENT, "#E76F51", "#8AB17D", "#9C89B8", "#C44536", "#3D5A80"],
    "background": "#FFFFFF", "foreground": "#1B2A33", "tableAccent": TEAL,
    "good": "#1E9E6A", "neutral": ACCENT, "bad": "#C44536",
    "textClasses": {
        "title": {"fontFace": "Segoe UI Semibold", "fontSize": 12, "color": NAVY},
        "label": {"fontFace": "Segoe UI", "fontSize": 10, "color": "#3B4A54"},
        "callout": {"fontFace": "Segoe UI Semibold", "fontSize": 28, "color": TEAL},
        "header": {"fontFace": "Segoe UI Semibold", "fontSize": 12, "color": NAVY},
    },
    "visualStyles": {"*": {"*": {
        "background": [{"show": True, "color": {"solid": {"color": "#FFFFFF"}}, "transparency": 0}],
        "border": [{"show": True, "color": {"solid": {"color": "#DCE7E7"}}, "radius": 10}],
        "dropShadow": [{"show": True, "color": {"solid": {"color": "#1B2A33"}}, "position": "Outer", "preset": "BottomRight", "transparency": 92}],
        "title": [{"show": True, "fontColor": {"solid": {"color": NAVY}}, "fontSize": 11, "fontFamily": "Segoe UI Semibold"}],
        "outspacePane": [{"backgroundColor": {"solid": {"color": LIGHT_BG}}}],
    }}},
}

# ----------------------------------------------------------------------------
# 4. REPORT BUILDER
# ----------------------------------------------------------------------------
ALIASES = {"Employees": "e", "Recruitment": "r", "Training": "t", "Leave": "l",
           "DimDepartment": "dd", "DimFacility": "df", "Date": "d", "HR Measures": "m"}

def lit(v):
    if isinstance(v, bool): return {"expr": {"Literal": {"Value": "true" if v else "false"}}}
    if isinstance(v, (int, float)): return {"expr": {"Literal": {"Value": f"{v}D"}}}
    return {"expr": {"Literal": {"Value": f"'{v}'"}}}

def solid(color): return {"solid": {"color": {"expr": {"Literal": {"Value": f"'{color}'"}}}}}

def mref(table, name):  # measure select item
    return {"Measure": {"Expression": {"SourceRef": {"Source": ALIASES[table]}}, "Property": name},
            "Name": f"{table}.{name}", "NativeReferenceName": name}

def cref(table, name):  # column select item
    return {"Column": {"Expression": {"SourceRef": {"Source": ALIASES[table]}}, "Property": name},
            "Name": f"{table}.{name}", "NativeReferenceName": name}

_vnum = [0]
def visual(vtype, x, y, w, h, projections, selects, title=None, objects=None,
           order_by=None, no_title=False):
    _vnum[0] += 1
    z = _vnum[0] * 100
    tables = sorted({s["Name"].split(".")[0] for s in selects})
    sv = {"visualType": vtype,
          "projections": projections,
          "prototypeQuery": {
              "Version": 2,
              "From": [{"Name": ALIASES[t], "Entity": t, "Type": 0} for t in tables],
              "Select": selects},
          "drillFilterOtherVisuals": True}
    if order_by: sv["prototypeQuery"]["OrderBy"] = order_by
    if objects: sv["objects"] = objects
    vc_objects = {}
    if title:
        vc_objects["title"] = [{"properties": {"show": lit(True), "text": lit(title)}}]
    elif no_title:
        vc_objects["title"] = [{"properties": {"show": lit(False)}}]
    if vc_objects: sv["vcObjects"] = vc_objects
    cfg = {"name": f"vis{_vnum[0]:04d}{uuid.uuid4().hex[:8]}",
           "layouts": [{"id": 0, "position": {"x": x, "y": y, "z": z, "width": w, "height": h, "tabOrder": z}}],
           "singleVisual": sv}
    return {"config": json.dumps(cfg), "filters": "[]",
            "x": float(x), "y": float(y), "z": float(z), "width": float(w), "height": float(h)}

def order_asc(table, column):
    return [{"Direction": 1, "Expression": {"Column": {"Expression": {"SourceRef": {"Source": ALIASES[table]}}, "Property": column}}}]

def order_desc_measure(table, measure):
    return [{"Direction": 2, "Expression": {"Measure": {"Expression": {"SourceRef": {"Source": ALIASES[table]}}, "Property": measure}}}]

def card(measure_name, title, x, y, w, h, color=TEAL):
    return visual("card", x, y, w, h,
        {"Values": [{"queryRef": f"HR Measures.{measure_name}"}]},
        [mref("HR Measures", measure_name)], title=title,
        objects={"labels": [{"properties": {"fontSize": lit(24), "color": solid(color)}}],
                 "categoryLabels": [{"properties": {"show": lit(False)}}]})

def cat_chart(vtype, cat, vals, title, x, y, w, h, sort_desc=True, extra_objects=None):
    """cat: (table, column); vals: list of measure names."""
    projections = {"Category": [{"queryRef": f"{cat[0]}.{cat[1]}"}],
                   "Y": [{"queryRef": f"HR Measures.{v}"} for v in vals]}
    selects = [cref(*cat)] + [mref("HR Measures", v) for v in vals]
    ob = order_desc_measure("HR Measures", vals[0]) if sort_desc else order_asc(*cat)
    objects = {"labels": [{"properties": {"show": lit(True)}}]}
    if extra_objects: objects.update(extra_objects)
    return visual(vtype, x, y, w, h, projections, selects, title=title, order_by=ob, objects=objects)

def donut(cat, val, title, x, y, w, h):
    return visual("donutChart", x, y, w, h,
        {"Category": [{"queryRef": f"{cat[0]}.{cat[1]}"}], "Y": [{"queryRef": f"HR Measures.{val}"}]},
        [cref(*cat), mref("HR Measures", val)], title=title,
        objects={"labels": [{"properties": {"show": lit(True), "labelStyle": lit("Data value, percent of total")}}],
                 "legend": [{"properties": {"show": lit(True), "position": lit("Right")}}]},
        order_by=order_desc_measure("HR Measures", val))

def treemap(cat, val, title, x, y, w, h):
    return visual("treemap", x, y, w, h,
        {"Group": [{"queryRef": f"{cat[0]}.{cat[1]}"}], "Values": [{"queryRef": f"HR Measures.{val}"}]},
        [cref(*cat), mref("HR Measures", val)], title=title,
        objects={"labels": [{"properties": {"show": lit(True)}}]},
        order_by=order_desc_measure("HR Measures", val))

def trend(vtype, vals, title, x, y, w, h, series=None):
    projections = {"Category": [{"queryRef": "Date.MonthStart"}],
                   "Y": [{"queryRef": f"HR Measures.{v}"} for v in vals]}
    selects = [cref("Date", "MonthStart")] + [mref("HR Measures", v) for v in vals]
    return visual(vtype, x, y, w, h, projections, selects, title=title,
                  order_by=order_asc("Date", "MonthStart"))

def slicer(table, column, title, x, y, w, h, dropdown=True):
    objects = {"data": [{"properties": {"mode": lit("Dropdown" if dropdown else "Basic")}}],
               "header": [{"properties": {"show": lit(True), "text": lit(title), "fontColor": solid(NAVY)}}],
               "general": [{"properties": {"selfFilterEnabled": lit(True)}}]}
    return visual("slicer", x, y, w, h,
        {"Values": [{"queryRef": f"{table}.{column}"}]},
        [cref(table, column)], objects=objects, no_title=True)

def textbox(runs, x, y, w, h, bg=None):
    _vnum[0] += 1
    z = _vnum[0] * 100
    paragraphs = [{"textRuns": [{"value": v, "textStyle": st} for v, st in runs]}]
    cfg = {"name": f"txt{_vnum[0]:04d}{uuid.uuid4().hex[:8]}",
           "layouts": [{"id": 0, "position": {"x": x, "y": y, "z": z, "width": w, "height": h, "tabOrder": z}}],
           "singleVisual": {"visualType": "textbox", "drillFilterOtherVisuals": True,
                            "objects": {"general": [{"properties": {"paragraphs": paragraphs}}]}}}
    if bg:
        cfg["singleVisual"]["vcObjects"] = {
            "background": [{"properties": {"show": lit(True), "color": solid(bg), "transparency": lit(0)}}],
            "border": [{"properties": {"show": lit(False)}}],
            "title": [{"properties": {"show": lit(False)}}]}
    return {"config": json.dumps(cfg), "filters": "[]",
            "x": float(x), "y": float(y), "z": float(z), "width": float(w), "height": float(h)}

def table_visual(items, title, x, y, w, h):
    """items: list of ('col', table, name) or ('measure', table, name)"""
    projections = {"Values": [{"queryRef": f"{t}.{n}"} for _, t, n in items]}
    selects = [cref(t, n) if k == "col" else mref(t, n) for k, t, n in items]
    return visual("tableEx", x, y, w, h, projections, selects, title=title)

def gauge(val, target, title, x, y, w, h):
    return visual("gauge", x, y, w, h,
        {"Y": [{"queryRef": f"HR Measures.{val}"}], "TargetValue": [{"queryRef": f"HR Measures.{target}"}]},
        [mref("HR Measures", val), mref("HR Measures", target)], title=title,
        objects={"labels": [{"properties": {"show": lit(True)}}]})

def header(title_text, subtitle):
    return textbox(
        [(title_text, {"fontWeight": "bold", "fontSize": "17pt", "color": "#FFFFFF", "fontFamily": "Segoe UI Semibold"}),
         ("   |   " + subtitle, {"fontSize": "11pt", "color": "#CFE8E8", "fontFamily": "Segoe UI"})],
        0, 0, 1280, 52, bg=NAVY)

def section(name, display_name, containers, ordinal):
    return {"config": json.dumps({"objects": {"background": [{"properties": {
                "color": solid(LIGHT_BG), "transparency": lit(0)}}]}}),
            "displayName": display_name, "displayOption": 1, "filters": "[]",
            "height": 720.0, "width": 1280.0, "name": name, "ordinal": ordinal,
            "visualContainers": containers}

PAGES = []
SLICER_ROW = lambda: [
    slicer("DimFacility", "Facility", "Facility", 16, 60, 300, 64),
    slicer("DimDepartment", "Department", "Department", 332, 60, 300, 64),
    slicer("Date", "Year", "Year", 648, 60, 160, 64),
]

# ---- Page 1: Executive Overview
v = []
v.append(header("CONTOSO HEALTH KSA (FICTIONAL)  -  HR EXECUTIVE OVERVIEW", "People analytics across 8 KSA facilities"))
v += SLICER_ROW()
cards = [("Headcount", "Active Headcount", TEAL), ("Hires", "Hires (Period)", NAVY),
         ("Attrition Rate %", "Attrition Rate", "#C44536"), ("Saudization %", "Saudization", "#1E9E6A"),
         ("Monthly Payroll (SAR)", "Monthly Payroll (SAR)", ACCENT)]
for i, (m, t, c) in enumerate(cards):
    v.append(card(m, t, 16 + i * 252, 136, 240, 92, color=c))
v.append(trend("lineChart", ["Headcount"], "Headcount Trend", 16, 240, 624, 224))
v.append(cat_chart("clusteredBarChart", ("DimFacility", "Facility"), ["Headcount"], "Headcount by Facility", 652, 240, 612, 224))
v.append(treemap(("DimDepartment", "Department"), "Headcount", "Headcount by Department", 16, 476, 624, 228))
v.append(donut(("Employees", "Gender"), "Headcount", "Gender Mix", 652, 476, 296, 228))
v.append(cat_chart("clusteredColumnChart", ("Employees", "EmploymentType"), ["Headcount"], "Employment Type", 964, 476, 300, 228))
PAGES.append(("ReportSection_Overview", "Executive Overview", v))

# ---- Page 2: Workforce Demographics
v = []
v.append(header("WORKFORCE DEMOGRAPHICS", "Who powers the workforce - age, gender, nationality & tenure"))
v += SLICER_ROW()
cards = [("Headcount", "Active Headcount", TEAL), ("Avg Age", "Average Age", NAVY),
         ("Female Ratio %", "Female Ratio", "#9C89B8"), ("Avg Tenure (Years)", "Avg Tenure (Years)", "#1E9E6A")]
for i, (m, t, c) in enumerate(cards):
    v.append(card(m, t, 16 + i * 316, 136, 304, 92, color=c))
v.append(cat_chart("clusteredColumnChart", ("Employees", "AgeBand"), ["Headcount"], "Headcount by Age Band", 16, 240, 410, 224, sort_desc=False))
v.append(cat_chart("clusteredBarChart", ("Employees", "Nationality"), ["Headcount"], "Headcount by Nationality", 442, 240, 410, 224))
v.append(cat_chart("clusteredColumnChart", ("Employees", "TenureBand"), ["Headcount"], "Headcount by Tenure Band", 868, 240, 396, 224, sort_desc=False))
v.append(visual("hundredPercentStackedBarChart", 16, 476, 624, 228,
    {"Category": [{"queryRef": "DimDepartment.Department"}],
     "Series": [{"queryRef": "Employees.Gender"}],
     "Y": [{"queryRef": "HR Measures.Headcount"}]},
    [cref("DimDepartment", "Department"), cref("Employees", "Gender"), mref("HR Measures", "Headcount")],
    title="Gender Split by Department", order_by=order_desc_measure("HR Measures", "Headcount")))
v.append(table_visual([("col", "DimDepartment", "Department"),
                       ("measure", "HR Measures", "Headcount"),
                       ("measure", "HR Measures", "Female Ratio %"),
                       ("measure", "HR Measures", "Saudization %"),
                       ("measure", "HR Measures", "Avg Tenure (Years)"),
                       ("measure", "HR Measures", "Avg Monthly Salary (SAR)")],
                      "Department Scorecard", 652, 476, 612, 228))
PAGES.append(("ReportSection_Demographics", "Demographics", v))

# ---- Page 3: Attrition & Retention
v = []
v.append(header("ATTRITION & RETENTION", "Turnover dynamics, reasons and hotspots"))
v += SLICER_ROW()
cards = [("Hires", "Hires", TEAL), ("Terminations", "Terminations", "#C44536"),
         ("Attrition Rate %", "Attrition Rate", "#E76F51"), ("Voluntary Share %", "Voluntary Share", ACCENT),
         ("Net Headcount Change", "Net Change", NAVY)]
for i, (m, t, c) in enumerate(cards):
    v.append(card(m, t, 16 + i * 252, 136, 240, 92, color=c))
v.append(trend("clusteredColumnChart", ["Hires", "Terminations"], "Hires vs Terminations by Month", 16, 240, 624, 224))
v.append(trend("lineChart", ["Attrition Rate %"], "Attrition Rate Trend", 652, 240, 612, 224))
v.append(cat_chart("clusteredBarChart", ("DimDepartment", "Department"), ["Terminations"], "Terminations by Department", 16, 476, 410, 228))
v.append(visual("funnel", 442, 476, 410, 228,
    {"Category": [{"queryRef": "Employees.TerminationReason"}], "Y": [{"queryRef": "HR Measures.Terminations"}]},
    [cref("Employees", "TerminationReason"), mref("HR Measures", "Terminations")],
    title="Termination Reasons", order_by=order_desc_measure("HR Measures", "Terminations"),
    objects={"labels": [{"properties": {"show": lit(True)}}]}))
v.append(cat_chart("clusteredBarChart", ("DimFacility", "Facility"), ["Attrition Rate %"], "Attrition Rate by Facility", 868, 476, 396, 228))
PAGES.append(("ReportSection_Attrition", "Attrition & Retention", v))

# ---- Page 4: Recruitment
v = []
v.append(header("TALENT ACQUISITION", "Pipeline health, speed and sourcing effectiveness"))
v += SLICER_ROW()
cards = [("Open Positions", "Open Positions", "#C44536"), ("Positions Filled", "Positions Filled", TEAL),
         ("Avg Time to Fill (Days)", "Avg Time to Fill (Days)", NAVY), ("Offer Acceptance %", "Offer Acceptance", "#1E9E6A"),
         ("Applications per Hire", "Applications per Hire", ACCENT)]
for i, (m, t, c) in enumerate(cards):
    v.append(card(m, t, 16 + i * 252, 136, 240, 92, color=c))
v.append(visual("funnel", 16, 240, 410, 224,
    {"Category": [{"queryRef": "Recruitment.Source"}], "Y": [{"queryRef": "HR Measures.Total Applications"}]},
    [cref("Recruitment", "Source"), mref("HR Measures", "Total Applications")],
    title="Applications by Source", order_by=order_desc_measure("HR Measures", "Total Applications"),
    objects={"labels": [{"properties": {"show": lit(True)}}]}))
v.append(cat_chart("clusteredBarChart", ("DimDepartment", "Department"), ["Avg Time to Fill (Days)"], "Time to Fill by Department", 442, 240, 410, 224))
v.append(trend("clusteredColumnChart", ["Offers Made", "Offers Accepted"], "Offers Made vs Accepted by Month", 868, 240, 396, 224))
v.append(table_visual([("col", "Recruitment", "RequisitionID"), ("col", "Recruitment", "Position"),
                       ("col", "Recruitment", "Department"), ("col", "Recruitment", "Facility"),
                       ("col", "Recruitment", "Status"), ("col", "Recruitment", "PostingDate"),
                       ("col", "Recruitment", "DaysToFill"), ("col", "Recruitment", "Source")],
                      "Requisition Register", 16, 476, 836, 228))
v.append(donut(("Recruitment", "Status"), "Total Applications", "Pipeline by Status", 868, 476, 396, 228))
PAGES.append(("ReportSection_Recruitment", "Recruitment", v))

# ---- Page 5: Saudization & Nitaqat
v = []
v.append(header("SAUDIZATION & NITAQAT COMPLIANCE", "National workforce localization vs 40% target"))
v += SLICER_ROW()
cards = [("Saudi Headcount", "Saudi Employees", "#1E9E6A"), ("Non-Saudi Headcount", "Non-Saudi Employees", NAVY),
         ("Saudization %", "Saudization Rate", TEAL)]
for i, (m, t, c) in enumerate(cards):
    v.append(card(m, t, 16 + i * 212, 136, 200, 92, color=c))
v.append(gauge("Saudization %", "Saudization Target", "Saudization vs Nitaqat Target (40%)", 652, 136, 612, 200))
v.append(cat_chart("clusteredBarChart", ("DimFacility", "Facility"), ["Saudization %"], "Saudization by Facility", 16, 240, 624, 224))
v.append(cat_chart("clusteredBarChart", ("DimDepartment", "Department"), ["Saudization %"], "Saudization by Department", 16, 476, 624, 228))
v.append(visual("stackedAreaChart", 652, 348, 612, 180,
    {"Category": [{"queryRef": "Date.MonthStart"}],
     "Y": [{"queryRef": "HR Measures.Saudi Headcount"}, {"queryRef": "HR Measures.Non-Saudi Headcount"}]},
    [cref("Date", "MonthStart"), mref("HR Measures", "Saudi Headcount"), mref("HR Measures", "Non-Saudi Headcount")],
    title="Saudi vs Non-Saudi Headcount Trend", order_by=order_asc("Date", "MonthStart")))
v.append(donut(("Employees", "Nationality"), "Headcount", "Nationality Mix", 652, 540, 612, 164))
PAGES.append(("ReportSection_Saudization", "Saudization", v))

# ---- Page 6: Compensation, Training & Leave
v = []
v.append(header("REWARDS, LEARNING & ABSENCE", "Pay equity, capability building and leave behaviour"))
v += SLICER_ROW()
cards = [("Avg Monthly Salary (SAR)", "Avg Salary (SAR)", TEAL), ("Median Salary (SAR)", "Median Salary (SAR)", NAVY),
         ("Training Hours", "Training Hours", "#1E9E6A"), ("Training Completion %", "Training Completion", ACCENT),
         ("Absenteeism %", "Absenteeism", "#C44536")]
for i, (m, t, c) in enumerate(cards):
    v.append(card(m, t, 16 + i * 252, 136, 240, 92, color=c))
v.append(cat_chart("clusteredBarChart", ("DimDepartment", "Department"), ["Monthly Payroll (SAR)"], "Monthly Payroll by Department", 16, 240, 410, 224))
v.append(cat_chart("clusteredColumnChart", ("Employees", "Grade"), ["Avg Monthly Salary (SAR)"], "Avg Salary by Grade", 442, 240, 410, 224, sort_desc=False))
v.append(cat_chart("clusteredColumnChart", ("Training", "Category"), ["Training Hours"], "Training Hours by Category", 868, 240, 396, 224))
v.append(trend("lineChart", ["Training Hours"], "Training Hours by Month", 16, 476, 410, 228))
v.append(donut(("Leave", "LeaveType"), "Leave Days", "Leave Days by Type", 442, 476, 410, 228))
v.append(cat_chart("clusteredBarChart", ("DimDepartment", "Department"), ["Avg Performance Rating"], "Avg Performance Rating by Department", 868, 476, 396, 228))
PAGES.append(("ReportSection_Rewards", "Rewards & Learning", v))

report_config = {
    "version": "5.43",
    "themeCollection": {
        "baseTheme": {"name": "CY24SU06", "version": "5.55", "type": 2},
        "customTheme": {"name": "HealthcareTheme.json", "version": "5.55", "type": 1},
    },
    "activeSectionIndex": 0,
    "defaultDrillFilterOtherVisuals": True,
    "settings": {"useNewFilterPaneExperience": True, "useStylableVisualContainerHeader": True},
    "objects": {"outspacePane": [{"properties": {"expanded": lit(False)}}]},
}

report_json = {
    "config": json.dumps(report_config),
    "layoutOptimization": 0,
    "resourcePackages": [
        {"resourcePackage": {"name": "SharedResources", "type": 2, "disabled": False,
                             "items": [{"name": "CY24SU06", "path": "BaseThemes/CY24SU06.json", "type": 202}]}},
        {"resourcePackage": {"name": "RegisteredResources", "type": 1, "disabled": False,
                             "items": [{"name": "HealthcareTheme.json", "path": "HealthcareTheme.json", "type": 202}]}},
    ],
    "sections": [section(n, d, c, i) for i, (n, d, c) in enumerate(PAGES)],
}

# ----------------------------------------------------------------------------
# 5. WRITE PROJECT FILES
# ----------------------------------------------------------------------------
def write(path, obj, raw=False):
    full = os.path.join(ROOT, path)
    os.makedirs(os.path.dirname(full), exist_ok=True)
    with open(full, "w", encoding="utf-8") as f:
        f.write(obj if raw else json.dumps(obj, indent=2))

write(f"{NAME}.pbip", {
    "$schema": "https://developer.microsoft.com/json-schemas/fabric/pbip/pbipProperties/1.0.0/schema.json",
    "version": "1.0",
    "artifacts": [{"report": {"path": f"{NAME}.Report"}}],
    "settings": {"enableAutoRecovery": True}})

write(f"{NAME}.Report/.platform", {
    "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
    "metadata": {"type": "Report", "displayName": NAME},
    "config": {"version": "2.0", "logicalId": str(uuid.uuid4())}})

write(f"{NAME}.Report/definition.pbir", {
    "version": "1.0",
    "datasetReference": {"byPath": {"path": f"../{NAME}.SemanticModel"}}})

write(f"{NAME}.Report/report.json", report_json)
write(f"{NAME}.Report/StaticResources/RegisteredResources/HealthcareTheme.json", theme)

write(f"{NAME}.SemanticModel/.platform", {
    "$schema": "https://developer.microsoft.com/json-schemas/fabric/gitIntegration/platformProperties/2.0.0/schema.json",
    "metadata": {"type": "SemanticModel", "displayName": NAME},
    "config": {"version": "2.0", "logicalId": str(uuid.uuid4())}})

write(f"{NAME}.SemanticModel/definition.pbism", {"version": "4.0", "settings": {}})
write(f"{NAME}.SemanticModel/model.bim", model_bim)

write(".gitignore", "*.pbi/\n.pbi/\ncache.abf\n", raw=True)

n_active = len(active_emps)
n_saudi = sum(1 for e in active_emps if e[4] == "Saudi")
print(f"OK  employees={len(employees)} active={n_active} saudization={n_saudi/n_active:.1%}")
print(f"    recruitment={len(recruitment)} training={len(training)} leave={len(leave)}")
print(f"    pages={len(PAGES)} visuals={_vnum[0]}")
