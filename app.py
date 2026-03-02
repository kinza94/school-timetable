import streamlit as st
import pandas as pd
import sqlite3
import json
from docx import Document
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
import random

import copy

conn = sqlite3.connect("school.db", check_same_thread=False)
c = conn.cursor()


def clean_name(x):
    return str(x).strip().upper()


def clean(x):
    return str(x).strip().upper()


# ==================================================
# -------- CONSTRAINT ENGINE GLOBAL STATE ----------
# ==================================================

teacher_day_load = {}
teacher_timeline = {}
double_used = {}
styles = getSampleStyleSheet()
styleN = styles["Normal"]
# Create table
c.execute("""
CREATE TABLE IF NOT EXISTS app_data (
    id INTEGER PRIMARY KEY,
    data TEXT
)
""")

conn.commit()

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if "role" not in st.session_state:
    st.session_state.role = None

if not st.session_state.logged_in:

    st.title("Login")

    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

    if st.button("Login"):

        if username == "admin" and password == "Kinz@420":
            st.session_state.logged_in = True
            st.session_state.role = "admin"
            st.success("Logged in as Admin")
            st.rerun()

        elif username in st.session_state.teachers and password == "1234":
            st.session_state.logged_in = True
            st.session_state.role = "viewer"
            st.success(f"Logged in as {username}")
            st.rerun()

        elif username == "head" and password == "9999":
            st.session_state.logged_in = True
            st.session_state.role = "viewer"
            st.success("Logged in as Head")
            st.rerun()

        else:
            st.error("Invalid credentials")

if st.session_state.role == "admin":
    if st.button("Generate Timetable", key="admin_generate"):
        pass

# ==================================================
# ---------------- APP SETUP -----------------------
# ==================================================
st.set_page_config(page_title="School Scheduler Pro", layout="wide")

st.markdown("""
    <style>
        .main {
            background-color: #f4f7fb;
        }
        h1 {
            color: #1f4e79;
        }
        .stButton>button {
            background-color: #4a90e2;
            color: white;
            border-radius: 8px;
            height: 3em;
            width: 100%;
        }
    </style>
""", unsafe_allow_html=True)

st.title("📚 School Timetable Scheduler Pro")

# ==================================================
# ---------------- GLOBAL CONSTANTS ----------------
# ==================================================
DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]

ALL_PERIODS = ["P1", "P2", "P3", "P4", "P5", "Lunch", "P6", "P7", "P8"]


def get_periods(day):
    if day == "Friday":
        return ["P1", "P2", "P3", "P4", "Lunch", "P5", "P6"]
    return ALL_PERIODS


# ==================================================
# ---------------- SESSION STATE -------------------
# ==================================================
# ==================================================
# ---------------- SESSION STATE -------------------
# ==================================================

if "teachers" not in st.session_state:
    st.session_state.teachers = {}

if "sections" not in st.session_state:
    st.session_state.sections = {}

if "class_teachers" not in st.session_state:
    st.session_state.class_teachers = {}

if "subject_config" not in st.session_state:
    st.session_state.subject_config = {}

if "teacher_assignment" not in st.session_state:
    st.session_state.teacher_assignment = {}

if "timetable" not in st.session_state:
    st.session_state.timetable = {}


def save_all_data():
    data = {
        "teachers": st.session_state.get("teachers", {}),
        "sections": st.session_state.get("sections", {}),
        "class_teachers": st.session_state.get("class_teachers", {}),
        "subject_config": st.session_state.get("subject_config", {}),
        "teacher_assignment": st.session_state.get("teacher_assignment", {}),
        "timetable": st.session_state.get("timetable", {}),
    }

    c.execute("SELECT COUNT(*) FROM app_data")
    count = c.fetchone()[0]

    if count == 0:
        c.execute(
            "INSERT INTO app_data (id, data) VALUES (1, ?)",
            (json.dumps(data),)
        )
    else:
        c.execute(
            "UPDATE app_data SET data=? WHERE id=1",
            (json.dumps(data),)
        )

    conn.commit()


def load_all_data():
    c.execute("SELECT data FROM app_data LIMIT 1")
    row = c.fetchone()

    if row:
        data = json.loads(row[0])

        st.session_state.teachers = data.get("teachers", {})
        st.session_state.sections = data.get("sections", {})
        st.session_state.class_teachers = data.get("class_teachers", {})
        st.session_state.subject_config = data.get("subject_config", {})
        st.session_state.teacher_assignment = data.get("teacher_assignment", {})
        st.session_state.timetable = data.get("timetable", {})

    cleaned = {}

    for t in st.session_state.teachers:
        cleaned[clean_name(t)] = {}

    st.session_state.teachers = cleaned
    for sec in st.session_state.timetable:
        for day in DAYS:
            for p in get_periods(day):
                teacher = st.session_state.timetable[sec][day][p]["teacher"]
                st.session_state.timetable[sec][day][p]["teacher"] = clean_name(teacher)


def validate_subject_weekly(section):
    issues = []

    if section not in st.session_state.subject_config:
        return issues

    required_counts = st.session_state.subject_config[section]
    actual_counts = {}

    # Count actual subjects in timetable
    for day in DAYS:
        for period in get_periods(day):
            subject = st.session_state.timetable[section][day][period]["subject"]
            if subject:
                actual_counts[subject] = actual_counts.get(subject, 0) + 1

    # Compare with required
    for subject, required in required_counts.items():
        actual = actual_counts.get(subject, 0)
        if actual != required:
            issues.append(
                f"{subject} should be {required} periods but currently {actual}"
            )

    return issues


def validate_teacher_clashes():
    issues = []

    for day in DAYS:
        for period in get_periods(day):

            teacher_slots = {}

            for section in st.session_state.timetable:

                teacher = st.session_state.timetable[section][day][period]["teacher"]

                if teacher:

                    if teacher not in teacher_slots:
                        teacher_slots[teacher] = []

                    teacher_slots[teacher].append(section)

            # If same teacher in multiple sections
            for teacher, sections in teacher_slots.items():
                if len(sections) > 1:
                    issues.append(
                        f"Clash: {teacher} has {sections} at {day} {period}"
                    )

    return issues


def validate_class_teacher_presence():
    issues = []

    for section, class_teacher in st.session_state.class_teachers.items():

        # 🔒 Safety check: section exists in timetable
        if section not in st.session_state.timetable:
            continue

        found = False

        for day in DAYS:

            # Safety check: day exists
            if day not in st.session_state.timetable[section]:
                continue

            for period in get_periods(day):

                # Safety check: period exists
                if period not in st.session_state.timetable[section][day]:
                    continue

                teacher = st.session_state.timetable[section][day][period]["teacher"]

                if teacher == class_teacher:
                    found = True
                    break

            if found:
                break

        if not found:
            issues.append(
                f"Class teacher {class_teacher} has no period in {section}"
            )

    return issues


if "data_loaded" not in st.session_state:
    load_all_data()
    st.session_state.data_loaded = True


# ==================================================
# AI DETECTION ENGINE
# ============================================
def find_alternative_teacher(section, subject, day, period):
    alternatives = []

    for teacher, sec_data in st.session_state.teacher_assignment.items():

        if section in sec_data and subject in sec_data[section]:

            # Skip original teacher
            if teacher == st.session_state.timetable[section][day][period]["teacher"]:
                continue

            if not teacher_busy(teacher, day, period):
                alternatives.append(teacher)

    return alternatives


def find_swap_option(section, teacher, subject, day, period):
    swaps = []

    for d in DAYS:
        for p in get_periods(d):

            if (d == day and p == period):
                continue

            if st.session_state.timetable[section][d][p]["teacher"] == teacher:

                # Check if swapping avoids clash
                if not teacher_busy(teacher, day, period):
                    swaps.append((d, p))

    return swaps


# ==================================================
# ---------------- CREATE TIMETABLE ----------------
# ==================================================
def can_assign(section, subject, teacher, day, period):
    periods = get_periods(day)
    idx = periods.index(period)

    # ---------- teacher clash ----------
    if teacher_busy(teacher, day, period):
        return False

    # ---------- max 6 periods/day ----------
    if teacher_day_load[teacher][day] >= 6:
        return False

    # ---------- no 4 or 5 consecutive ----------
    timeline = teacher_timeline[teacher][day].copy()
    timeline[idx] = 1

    streak = 0
    for v in timeline:
        streak = streak + 1 if v else 0
        if streak >= 4:
            return False

    # ---------- IX-X double rule ----------
    if is_ix_x_double(subject):

        prev_same = False
        next_same = False

        if idx > 0:
            prev = periods[idx - 1]
            if st.session_state.timetable[section][day][prev]["subject"] == subject:
                prev_same = True

        if idx < len(periods) - 1:
            nxt = periods[idx + 1]
            if st.session_state.timetable[section][day][nxt]["subject"] == subject:
                next_same = True

        if (prev_same or next_same):
            if double_used.get((section, subject), False):
                return False

    return True


# ==================================================
# -------- SAFE ASSIGNMENT (STATE UPDATE) ----------
# ==================================================

def apply_assignment(section, subject, teacher, day, period):
    st.session_state.timetable[section][day][period]["subject"] = subject
    st.session_state.timetable[section][day][period]["teacher"] = teacher

    idx = get_periods(day).index(period)

    teacher_day_load[teacher][day] += 1
    teacher_timeline[teacher][day][idx] = 1

    # mark IX-X double used
    if is_ix_x_double(subject):

        periods = get_periods(day)

        if idx > 0:
            if st.session_state.timetable[section][day][periods[idx - 1]]["subject"] == subject:
                double_used[(section, subject)] = True

        if idx < len(periods) - 1:
            if st.session_state.timetable[section][day][periods[idx + 1]]["subject"] == subject:
                double_used[(section, subject)] = True


# --------------------------------------------
# CREATE EMPTY TIMETABLE
# --------------------------------------------
def create_empty_timetable():
    timetable = {}

    # -------- RESET TRACKERS ONCE --------
    teacher_day_load.clear()
    teacher_timeline.clear()
    double_used.clear()

    for t in st.session_state.teachers:
        teacher_day_load[t] = {}
        teacher_timeline[t] = {}

        for d in DAYS:
            periods = get_periods(d)
            teacher_day_load[t][d] = 0
            teacher_timeline[t][d] = [0] * len(periods)

    # -------- CREATE TABLE --------
    for section in st.session_state.sections:
        timetable[section] = {}
        for day in DAYS:
            timetable[section][day] = {}
            for period in get_periods(day):
                timetable[section][day][period] = {
                    "subject": "",
                    "teacher": ""
                }

    return timetable


# -------------
# ---DOUBLE RULE
def is_ix_x_double(subject):
    s = subject.upper()

    targets = [
        "PHYSICS", "CHEMISTRY", "COMPUTER IX-X",
    ]

    return any(t in s for t in targets)


# TEACHER BUSY CHECK
# ----------------------------
def teacher_busy(teacher, day, period):
    for sec in st.session_state.timetable:
        if st.session_state.timetable[sec][day][period]["teacher"] == teacher:
            return True
    return False


def teacher_daily_load(teacher, day):
    count = 0
    for sec in st.session_state.timetable:
        for period in get_periods(day):
            if st.session_state.timetable[sec][day][period]["teacher"] == teacher:
                count += 1
    return count


def assign_class_teacher_priority():
    for section, class_teacher in st.session_state.class_teachers.items():

        # Check if teacher teaches any subject in this section
        if class_teacher not in st.session_state.teacher_assignment:
            continue

        if section not in st.session_state.teacher_assignment[class_teacher]:
            continue

        subjects = st.session_state.teacher_assignment[class_teacher][section]

        for day in DAYS:

            if "P1" not in get_periods(day):
                continue

            if not subjects:
                continue

            # Pick first subject assigned to class teacher
            subject = subjects[0]

            # Check if subject still needs periods
            required = st.session_state.subject_config[section].get(subject, 0)

            # count already placed periods
            current_count = 0
            for d in DAYS:
                for p in get_periods(d):
                    if st.session_state.timetable[section][d][p]["subject"] == subject:
                        current_count += 1

            if current_count >= required:
                continue

            # Check slot free & teacher not busy
            if (
                    st.session_state.timetable[section][day]["P1"]["subject"] == ""
                    and can_assign(section, subject, class_teacher, day, "P1")
            ):
                apply_assignment(section, subject, class_teacher, day, "P1")


def calculate_fitness():
    score = 1000  # start high

    # Penalize 3 consecutive
    issues = validate_no_three_consecutive()
    score -= len(issues) * 50

    # Penalize imbalance
    dist = validate_teacher_distribution()
    score -= len(dist) * 20

    # Penalize Friday overload
    friday = validate_friday_load()
    score -= len(friday) * 10

    return score


# -------------------------------------------
# BASIC AUTO FILL ENGINE
# --------------------------------------------
def try_swap(section, subject, teacher):
    for day in DAYS:
        for period in get_periods(day):

            if period == "Lunch":
                continue

            current = st.session_state.timetable[section][day][period]

            # Skip if empty
            if current["subject"] == "":
                continue

            other_subject = current["subject"]
            other_teacher = current["teacher"]

            # Can we move this subject somewhere else?
            for d2 in DAYS:
                for p2 in get_periods(d2):

                    if p2 == "Lunch":
                        continue

                    target = st.session_state.timetable[section][d2][p2]

                    if target["subject"] == "" and not teacher_busy(other_teacher, d2, p2):

                        # Check original slot safe for new teacher
                        if not teacher_busy(teacher, day, period):

                            # perform swap safely

                            if not can_assign(section, subject, teacher, day, period):
                                continue

                            # --- clear BOTH slots FIRST ---
                            st.session_state.timetable[section][day][period] = {
                                "subject": "",
                                "teacher": ""
                            }

                            st.session_state.timetable[section][d2][p2] = {
                                "subject": "",
                                "teacher": ""
                            }

                            # --- now assign ---
                            apply_assignment(section, subject, teacher, day, period)
                            apply_assignment(section, other_subject, other_teacher, d2, p2)

                            return True

    return False


def basic_auto_fill():
    sections = list(st.session_state.subject_config.keys())
    random.shuffle(sections)

    for section in sections:

        subjects = st.session_state.subject_config[section]
        maths_double_done = False

        subject_items = list(subjects.items())
        random.shuffle(subject_items)

        for subject, count in subject_items:

            # STEP 1️⃣: Find assigned teacher FIRST
            assigned_teacher = None

            for teacher, sec_data in st.session_state.teacher_assignment.items():
                if section in sec_data and subject in sec_data[section]:
                    assigned_teacher = teacher
                    break

            if not assigned_teacher:
                continue

            # STEP 2️⃣: Daily guarantee for high frequency subjects
            if count >= 5:

                for day in DAYS:

                    already_present = False
                    for p in get_periods(day):
                        if st.session_state.timetable[section][day][p]["subject"] == subject:
                            already_present = True
                            break

                    if already_present:
                        continue

                    valid_periods = []

                    periods = get_periods(day).copy()
                    random.shuffle(periods)

                    for period in periods:
                        real_periods = get_periods(day)
                        real_index = real_periods.index(period)

                        if period == "Lunch":
                            continue

                        if (
                                st.session_state.timetable[section][day][period]["subject"] == ""
                                and can_assign(section, subject, assigned_teacher, day, period)
                        ):
                            valid_periods.append(period)

                    if valid_periods:
                        chosen = random.choice(valid_periods)

                        apply_assignment(section, subject, assigned_teacher, day, chosen)

            # 🔹 Find teacher (still inside subject loop)
            for teacher, sec_data in st.session_state.teacher_assignment.items():
                if section in sec_data and subject in sec_data[section]:
                    assigned_teacher = teacher
                    break

            if not assigned_teacher:
                continue

            # Count already filled slots for this subject
            filled = 0
            for d in DAYS:
                for p in get_periods(d):
                    if st.session_state.timetable[section][d][p]["subject"] == subject:
                        filled += 1

            subject_day_count = {day: 0 for day in DAYS}

            while filled < count:

                valid_slots = []

                days = DAYS.copy()
                random.shuffle(days)

                for day in days:

                    if subject_day_count[day] >= 2:
                        continue

                    periods = get_periods(day).copy()
                    random.shuffle(periods)

                    for period in periods:
                        real_periods = get_periods(day)
                        real_index = real_periods.index(period)

                        if period == "Lunch":
                            continue

                        if st.session_state.timetable[section][day][period]["subject"] != "":
                            continue

                        if teacher_busy(assigned_teacher, day, period):
                            continue

                        if teacher_daily_load(assigned_teacher, day) >= 6:
                            continue
                        # 🔴 HARD BLOCK: prevent ANY 3 consecutive pattern

                        periods = get_periods(day)

                        # ---- HARD STREAK CHECK ----

                        if not can_assign(section, subject, assigned_teacher, day, period):
                            continue
                        valid_slots.append((0, day, period))

                # ===============================
                # PICK BEST SLOT
                # ===============================

                if not valid_slots:

                    swapped = try_swap(section, subject, assigned_teacher)

                    if not swapped:
                        print(f"⚠ Could not place {subject}")
                        filled = count
                        break
                    else:
                        filled += 1
                        continue

                best = random.choice(valid_slots)
                _, best_day, best_period = best

                apply_assignment(section, subject, assigned_teacher, best_day, best_period)

                subject_day_count[best_day] += 1
                filled += 1


def validate_teacher_distribution():
    issues = []

    for teacher in st.session_state.teachers:
        daily_load = []

        for day in DAYS:
            count = 0
            for sec in st.session_state.timetable:
                for period in get_periods(day):
                    if st.session_state.timetable[sec][day][period]["teacher"] == teacher:
                        count += 1
            daily_load.append(count)

        if max(daily_load) - min(daily_load) >= 2:
            issues.append(f"{teacher} workload not balanced: {daily_load}")

    return issues


def validate_friday_load():
    issues = []

    for teacher in st.session_state.teachers:
        count = 0
        for sec in st.session_state.timetable:
            for period in get_periods("Friday"):
                if st.session_state.timetable[sec]["Friday"][period]["teacher"] == teacher:
                    count += 1

        if count > 5:
            issues.append(f"{teacher} heavy Friday ({count} periods)")
        elif count == 5:
            issues.append(f"{teacher} Friday 5 periods (acceptable)")

    return issues


def validate_maths_consecutive():
    issues = []

    for sec in st.session_state.timetable:
        for day in DAYS:

            maths_positions = []

            for period in get_periods(day):
                subject = st.session_state.timetable[sec][day][period]["subject"]
                if subject and "Math" in subject:
                    maths_positions.append(period)

            if len(maths_positions) == 2:
                periods = get_periods(day)
                idx1 = periods.index(maths_positions[0])
                idx2 = periods.index(maths_positions[1])

                if abs(idx1 - idx2) != 1:
                    issues.append(f"{sec} Maths not consecutive on {day}")

    return issues


def validate_double_period_rule():
    issues = []

    for section in st.session_state.timetable:

        if section not in st.session_state.subject_config:
            continue

        required_config = st.session_state.subject_config[section]

        for subject, weekly_required in required_config.items():

            double_days = 0

            for day in DAYS:
                count_today = 0

                for period in get_periods(day):
                    if st.session_state.timetable[section][day][period]["subject"] == subject:
                        count_today += 1

                if count_today >= 2:
                    double_days += 1

            if weekly_required >= 6:
                if double_days != 1:
                    issues.append(f"{section} - {subject} must have exactly 1 double period")
            else:
                if double_days > 0:
                    issues.append(f"{section} - {subject} cannot have double period")

    return issues


# SUGGESTION============================================
def get_free_teachers(day, period):
    free = []

    for teacher in st.session_state.teachers:
        if not teacher_busy(teacher, day, period):
            free.append(teacher)

    return free


def suggest_safe_slots(section, teacher):
    safe = []

    for day in DAYS:
        for period in get_periods(day):

            if period == "Lunch":
                continue

            if (
                    st.session_state.timetable[section][day][period]["subject"] == ""
                    and not teacher_busy(teacher, day, period)
            ):
                safe.append((day, period))

    return safe[:5]


# ===================================
# ------- max 25 periods per week----
# ===================================
def validate_teacher_max_load(max_periods=25):
    issues = []

    for teacher in st.session_state.teachers:

        count = 0

        for sec in st.session_state.timetable:
            for day in DAYS:
                for period in get_periods(day):
                    if st.session_state.timetable[sec][day][period]["teacher"] == teacher:
                        count += 1

        if count > max_periods:
            issues.append(f"{teacher} exceeds {max_periods} periods (currently {count})")

    return issues


# ================================================
# ----------TEACHER REPLACEMENT----------------
# ================================================
def replace_teacher_everywhere(old_teacher, new_teacher):
    for section in st.session_state.timetable:
        for day in DAYS:
            for period in get_periods(day):

                if st.session_state.timetable[section][day][period]["teacher"] == old_teacher:
                    st.session_state.timetable[section][day][period]["teacher"] = new_teacher

    save_all_data()


def export_all_teachers_word():
    file_path = "All_Teachers_Timetable.docx"

    doc = Document()
    doc.add_heading("All Teachers Timetable", level=1)

    df = build_all_teachers_matrix()

    table = doc.add_table(
        rows=df.shape[0] + 1,
        cols=df.shape[1]
    )

    table.style = "Table Grid"

    # Header
    for col_index, col_name in enumerate(df.columns):
        table.cell(0, col_index).text = str(col_name)

    # Data
    for i in range(df.shape[0]):
        for j in range(df.shape[1]):
            table.cell(i + 1, j).text = str(df.iloc[i, j])

    doc.save(file_path)

    return file_path


def export_all_teachers_pdf():
    file_path = "All_Teachers_Timetable.pdf"

    doc = SimpleDocTemplate(file_path, pagesize=A4)
    elements = []

    df = build_all_teachers_matrix()

    data = [df.columns.tolist()] + df.values.tolist()

    from reportlab.lib.pagesizes import landscape

    doc = SimpleDocTemplate(
        file_path,
        pagesize=landscape(A4),  # 🔥 rotate page
        leftMargin=20,
        rightMargin=20,
        topMargin=30,
        bottomMargin=30
    )

    # ---------- COLUMN WIDTH CONTROL ----------
    page_width = landscape(A4)[0] - 40
    num_cols = max(len(data[0]), 1)

    col_width = max(page_width / num_cols, 40)
    col_widths = [col_width] * num_cols

    # ✅ CREATE TABLE (THIS LINE WAS MISSING)
    table = Table(data, colWidths=col_widths, repeatRows=1)

    # ---------- STYLE ----------
    table.setStyle(TableStyle([

        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1f4e79")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),

        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),

        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),

        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('BACKGROUND', (0, 1), (-1, -1), colors.whitesmoke),
    ]))

    elements.append(table)

    doc.build(elements)

    return file_path


def export_teacher_view_word(teacher):
    doc = Document()
    doc.add_heading(f"Teacher Timetable - {teacher}", 0)

    table = doc.add_table(rows=len(ALL_PERIODS) + 1, cols=len(DAYS) + 1)
    table.style = "Table Grid"

    # Header
    table.cell(0, 0).text = "Period"

    for i, day in enumerate(DAYS):
        table.cell(0, i + 1).text = day

    # Fill timetable
    for r, period in enumerate(ALL_PERIODS):

        table.cell(r + 1, 0).text = period

        for c, day in enumerate(DAYS):

            value = ""

            if period in get_periods(day):

                for sec in st.session_state.timetable:
                    entry = st.session_state.timetable[sec][day][period]

                    if clean(entry["teacher"]) == clean(teacher):
                        value = f"{sec}\n{entry['subject']}"
                        break

            table.cell(r + 1, c + 1).text = value

    filename = f"{teacher}_timetable.docx"
    doc.save(filename)

    return filename


def export_teacher_view_pdf(teacher):
    file_path = f"{teacher}_timetable.pdf"

    doc = SimpleDocTemplate(file_path, pagesize=A4)
    elements = []

    styles = getSampleStyleSheet()

    title = Paragraph(
        f"<b>Teacher Timetable — {teacher}</b>",
        styles["Title"]
    )

    elements.append(title)
    elements.append(Spacer(1, 20))

    data = [["Period"] + DAYS]

    for period in ALL_PERIODS:

        row = [period]

        for day in DAYS:

            value = ""

            if period in get_periods(day):

                for sec in st.session_state.timetable:
                    entry = st.session_state.timetable[sec][day][period]

                    if entry["teacher"] == teacher:
                        value = f"{sec}\n{entry['subject']}"
                        break

            row.append(value)

        data.append(row)

    table = Table(data, repeatRows=1)

    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1f4e79")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
    ]))

    elements.append(table)
    doc.build(elements)

    return file_path


def export_class_timetable_pdf(section):

    file_path = f"{section}_timetable.pdf"

    doc = SimpleDocTemplate(
        file_path,
        pagesize=landscape(A4),
        rightMargin=40,
        leftMargin=40,
        topMargin=40,
        bottomMargin=40
    )

    elements = []
    styles = getSampleStyleSheet()

    # ---------- TITLE ----------
    title = Paragraph(
        f"<b>Class Timetable — {section}</b>",
        styles["Title"]
    )

    elements.append(title)
    elements.append(Spacer(1, 20))

    # ---------- TABLE DATA ----------
    data = [["Period"] + DAYS]

    for period in ALL_PERIODS:
        row = [period]

        for day in DAYS:
            if period in get_periods(day):
                cell = st.session_state.timetable[section][day][period]

                text = ""
                if cell["subject"]:
                    text = f"{cell['subject']}\n({cell['teacher']})"
            else:
                text = ""

            row.append(text)

        data.append(row)

    # ---------- COLUMN WIDTH ----------
    page_width = landscape(A4)[0] - 80
    num_cols = len(data[0])
    col_width = page_width / num_cols
    col_widths = [col_width] * num_cols

    # ✅ CREATE TABLE (THIS WAS MISSING)
    table = Table(data, colWidths=col_widths, repeatRows=1)

    # ---------- STYLE ----------
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1f4e79")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('BACKGROUND', (0, 1), (-1, -1), colors.whitesmoke),
    ]))

    elements.append(table)

    doc.build(elements)

    return file_path


def export_all_teachers_pdf():
    file_path = "All_Teachers_Timetable.pdf"

    doc = SimpleDocTemplate(file_path, pagesize=A4)
    elements = []

    df = build_all_teachers_matrix()

    data = []

    # header
    data.append([
        Paragraph("" if col is None else str(col), styleN)
        for col in df.columns
    ])

    # rows
    for row in df.values.tolist():
        clean_row = []

        for cell in row:
            if cell is None:
                text = ""
            else:
                text = str(cell)

            clean_row.append(Paragraph(text, styleN))

        data.append(clean_row)

    table = Table(data, repeatRows=1)

    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1f4e79")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTSIZE', (0, 0), (-1, -1), 7),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER')
    ]))

    elements.append(table)
    doc.build(elements)

    return file_path


# ==================================================
# ---------------- SIDEBAR -------------------------
# ==================================================
if st.session_state.role == "admin":
    menu = st.sidebar.selectbox(
        "Navigation",
        ["Dashboard", "Configuration", "Generate", "Class View", "Teacher View", "Analytics"]
    )
else:
    menu = st.sidebar.selectbox(
        "Navigation",
        ["Class View", "Teacher View", "Analytics"]
    )


def build_all_teachers_matrix():
    columns = []

    # create columns → Monday P1, Monday P2 ...
    for day in DAYS:
        for period in get_periods(day):
            if period != "Lunch":
                columns.append(f"{day[:3]}-{period}")

    data = []

    for teacher in st.session_state.teachers:

        row = {"Teacher": teacher}

        for day in DAYS:
            for period in get_periods(day):

                if period == "Lunch":
                    continue

                key = f"{day[:3]}-{period}"
                value = ""

                for sec in st.session_state.timetable:
                    entry = st.session_state.timetable[sec][day][period]

                    if entry["teacher"] == teacher:
                        value = f"{sec} {entry['subject']}"
                        break

                row[key] = value

        data.append(row)

    df = pd.DataFrame(data)

    return df


# ==================================================
# ---------------- DASHBOARD -----------------------
# ==================================================
if menu == "Dashboard":
    st.subheader("CREATE YOUR TIME TABLE👋")
    st.write("Use sidebar to configure and generate timetable.")

# ==================================================
# ---------------- CONFIGURATION -------------------
# ==================================================
if menu == "Configuration":

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("➕ Add Section")
        sec = st.text_input("Section Name (e.g., 6A)")
        if st.button("Add Section"):
            if sec:
                st.session_state.sections[sec] = {}
                save_all_data()
                st.success("Section Added")

        st.write("Sections:", list(st.session_state.sections.keys()))
        if st.session_state.sections:
            remove_sec = st.selectbox(
                "Remove Section",
                list(st.session_state.sections.keys()),
                key="remove_section_select"
            )

            if st.button("Delete Section"):
                st.session_state.sections.pop(remove_sec, None)
                st.session_state.subject_config.pop(remove_sec, None)
                st.session_state.class_teachers.pop(remove_sec, None)
                st.success("Section Removed")
                save_all_data()

    with col2:
        st.subheader("➕ Add Teacher")
        teacher = st.text_input("Teacher Name")
        if st.button("Add Teacher"):
            if teacher:
                st.session_state.teachers[teacher] = {}
                save_all_data()
                st.success("Teacher Added")

        st.write("Teachers:", list(st.session_state.teachers.keys()))
        if st.session_state.teachers:
            remove_teacher = st.selectbox(
                "Remove Teacher",
                list(st.session_state.teachers.keys()),
                key="remove_teacher_select"
            )

            if st.button("Delete Teacher"):
                st.session_state.teachers.pop(remove_teacher, None)
                st.success("Teacher Removed")
                save_all_data()

    st.subheader("📌 Assign Class Teacher")

    if st.session_state.sections and st.session_state.teachers:
        selected_sec = st.selectbox(
            "Select Section",
            list(st.session_state.sections.keys()),
            key="class_teacher_section"
        )

        selected_teacher = st.selectbox("Select Teacher", list(st.session_state.teachers.keys()))

        if st.button("Assign"):
            st.session_state.class_teachers[selected_sec] = selected_teacher
            st.success("Class Teacher Assigned")

    st.write("Class Teachers:", st.session_state.class_teachers)
    if st.session_state.class_teachers:
        remove_class_teacher = st.selectbox(
            "Remove Class Teacher",
            list(st.session_state.class_teachers.keys()),
            key="remove_class_teacher_select"
        )

        if st.button("Delete Class Teacher"):
            st.session_state.class_teachers.pop(remove_class_teacher, None)
            st.success("Class Teacher Removed")
            save_all_data()

    st.subheader("📚 Configure Subjects for Section")

    if st.session_state.sections:

        selected_section = st.selectbox(
            "Select Section",
            list(st.session_state.sections.keys()),
            key="subject_config_section"
        )

        subject_name = st.text_input("Subject Name (e.g., Maths)")
        weekly_periods = st.number_input("Weekly Periods", min_value=1, max_value=10, step=1)

        if st.button("Add Subject"):
            if selected_section not in st.session_state.subject_config:
                st.session_state.subject_config[selected_section] = {}

            st.session_state.subject_config[selected_section][subject_name] = weekly_periods
            save_all_data()
            st.success("Subject Added Successfully")

        st.write("### Subjects for this Section:")
        st.write(st.session_state.subject_config.get(selected_section, {}))
        subjects_for_section = st.session_state.subject_config.get(selected_section, {})

        if subjects_for_section:
            remove_subject = st.selectbox(
                "Remove Subject",
                list(subjects_for_section.keys()),
                key="remove_subject_select"
            )

            if st.button("Delete Subject"):
                st.session_state.subject_config[selected_section].pop(remove_subject, None)
                st.success("Subject Removed")

            st.subheader("👩‍🏫 Assign Teacher to Section & Subject")

            if st.session_state.teachers and st.session_state.subject_config:

                assign_teacher = st.selectbox(
                    "Select Teacher",
                    list(st.session_state.teachers.keys()),
                    key="assign_teacher_select"
                )

                assign_section = st.selectbox(
                    "Select Section",
                    list(st.session_state.subject_config.keys()),
                    key="assign_section_select"
                )

                subjects_available = list(
                    st.session_state.subject_config.get(assign_section, {}).keys()
                )

                if subjects_available:

                    assign_subject = st.selectbox(
                        "Select Subject",
                        subjects_available,
                        key="assign_subject_select"
                    )

                    # ✅ Assign button MUST be inside this block
                    if st.button("Assign Teacher"):

                        if assign_teacher not in st.session_state.teacher_assignment:
                            st.session_state.teacher_assignment[assign_teacher] = {}

                        if assign_section not in st.session_state.teacher_assignment[assign_teacher]:
                            st.session_state.teacher_assignment[assign_teacher][assign_section] = []

                        if assign_subject not in st.session_state.teacher_assignment[assign_teacher][assign_section]:
                            st.session_state.teacher_assignment[assign_teacher][assign_section].append(assign_subject)

                        st.success("Teacher Assigned Successfully")

            # ===============================
            # CURRENT ASSIGNMENTS
            # ===============================

            st.write("### Current Teacher Assignments")
            st.write(st.session_state.teacher_assignment)

            st.subheader("❌ Remove Teacher Assignment")

            if st.session_state.teacher_assignment:

                remove_teacher = st.selectbox(
                    "Select Teacher",
                    list(st.session_state.teacher_assignment.keys()),
                    key="remove_assign_teacher"
                )

                sections_for_teacher = list(
                    st.session_state.teacher_assignment[remove_teacher].keys()
                )

                if sections_for_teacher:

                    remove_section = st.selectbox(
                        "Select Section",
                        sections_for_teacher,
                        key="remove_assign_section"
                    )

                    subjects_for_teacher = st.session_state.teacher_assignment[remove_teacher][remove_section]

                    if subjects_for_teacher:

                        remove_subject = st.selectbox(
                            "Select Subject",
                            subjects_for_teacher,
                            key="remove_assign_subject"
                        )

                        if st.button("Delete Assignment"):

                            st.session_state.teacher_assignment[remove_teacher][remove_section].remove(remove_subject)

                            # Clean empty structures
                            if not st.session_state.teacher_assignment[remove_teacher][remove_section]:
                                del st.session_state.teacher_assignment[remove_teacher][remove_section]

                            if not st.session_state.teacher_assignment[remove_teacher]:
                                del st.session_state.teacher_assignment[remove_teacher]

                            st.success("Assignment Removed Successfully")

            # ==================================================
            # ---------------- GENERATE ------------------------
            # ==================================================


def validate_no_three_consecutive():
    return []

    if menu == "Generate":

        if st.button("Generate Timetable", key="generate_main"):

            best_score = -999999
            best_timetable = None

            for _ in range(15):
                temp_table = create_empty_timetable()
                st.session_state.timetable = temp_table

                assign_class_teacher_priority()
                basic_auto_fill()

                score = calculate_fitness()

                if score > best_score:
                    best_score = score
                    best_timetable = copy.deepcopy(temp_table)

            st.session_state.timetable = best_timetable
            st.success("Timetable Generated Successfully")
            save_all_data()

        # ✅ Replacement Section MUST stay inside Generate menu
        if st.session_state.timetable:

            st.subheader("🔄 Replace Teacher")

            old_teacher = st.selectbox(
                "Select Teacher to Replace",
                list(st.session_state.teachers.keys()),
                key="replace_old"
            )

            new_teacher = st.selectbox(
                "Replace With",
                list(st.session_state.teachers.keys()),
                key="replace_new"
            )

            if st.button("Replace Teacher"):
                replace_teacher_everywhere(old_teacher, new_teacher)
                save_all_data()
                st.success(f"{old_teacher} replaced with {new_teacher}")

    # ===============================
    # SOFT CONSTRAINT VALIDATIONS
    # ===============================


consecutive_issues = validate_no_three_consecutive()
for issue in consecutive_issues:
    st.warning(issue)

distribution_issues = validate_teacher_distribution()
for issue in distribution_issues:
    st.warning(issue)

friday_issues = validate_friday_load()
for issue in friday_issues:
    st.info(issue)

maths_issues = validate_maths_consecutive()
for issue in maths_issues:
    st.info(issue)


def get_clash_slots():
    clashes = set()

    for day in DAYS:
        for period in get_periods(day):

            teacher_map = {}

            for sec in st.session_state.timetable:
                teacher = st.session_state.timetable[sec][day][period]["teacher"]

                if teacher:
                    teacher_map.setdefault(teacher, []).append(sec)

            for teacher, sections in teacher_map.items():
                if len(sections) > 1:
                    for sec in sections:
                        clashes.add((sec, day, period))

    return clashes


def export_teacher_timetable_pdf(teacher):
    from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import getSampleStyleSheet

    file_path = f"{teacher}_timetable.pdf"

    doc = SimpleDocTemplate(file_path, pagesize=A4)
    elements = []

    styles = getSampleStyleSheet()

    title = Paragraph(
        f"<b>Teacher Timetable — {teacher}</b>",
        styles["Title"]
    )

    elements.append(title)
    elements.append(Spacer(1, 20))

    data = [["Period"] + DAYS]

    for period in ALL_PERIODS:
        row = [period]

        for day in DAYS:

            if period in get_periods(day):

                found = False

                for sec in st.session_state.timetable:
                    if st.session_state.timetable[sec][day][period]["teacher"] == teacher:
                        subject = st.session_state.timetable[sec][day][period]["subject"]
                        row.append(f"{sec}\n{subject}")
                        found = True
                        break

                if not found:
                    row.append("")
            else:
                row.append("")

        data.append(row)

    page_width = A4[0] - 80
    col_width = page_width / (len(DAYS) + 1)
    col_widths = [col_width] * (len(DAYS) + 1)

    table = Table(data, colWidths=col_widths, repeatRows=1)

    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1f4e79")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('BACKGROUND', (0, 1), (-1, -1), colors.whitesmoke),
    ]))

    elements.append(table)
    doc.build(elements)

    return file_path


def export_class_timetable_pdf(section):

    file_path = f"{section}_timetable.pdf"

    doc = SimpleDocTemplate(
        file_path,
        pagesize=landscape(A4),   # landscape use kar rahe ho width calculation ke liye
        rightMargin=40,
        leftMargin=40,
        topMargin=40,
        bottomMargin=40
    )

    elements = []
    styles = getSampleStyleSheet()

    # ---------- TITLE ----------
    title = Paragraph(
        f"<b>Class Timetable — {section}</b>",
        styles["Title"]
    )

    elements.append(title)
    elements.append(Spacer(1, 20))

    # ---------- TABLE DATA ----------
    data = [["Period"] + DAYS]

    for period in ALL_PERIODS:
        row = [period]

        for day in DAYS:
            if period in get_periods(day):
                cell = st.session_state.timetable[section][day][period]

                if cell["subject"]:
                    text = f"{cell['subject']}\n({cell['teacher']})"
                else:
                    text = ""
            else:
                text = ""

            row.append(text)

        data.append(row)

    # ---------- COLUMN WIDTH ----------
    page_width = landscape(A4)[0] - 40
    num_cols = len(data[0])
    col_width = max(page_width / num_cols, 40)
    col_widths = [col_width] * num_cols

    # ✅ CREATE TABLE (THIS WAS MISSING)
    table = Table(data, colWidths=col_widths, repeatRows=1)

    # ---------- STYLE ----------
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor("#1f4e79")),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.white),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, -1), 9),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('BACKGROUND', (0, 1), (-1, -1), colors.whitesmoke),
    ]))

    elements.append(table)

    doc.build(elements)

    return file_path


# ==================================================
# ---------------- CLASS VIEW ----------------------
# ==================================================
if menu == "Class View":
    is_admin = st.session_state.get("role") == "admin"

    if not st.session_state.timetable:
        st.warning("Generate timetable first")
    else:
        sec = st.selectbox("Select Section", list(st.session_state.timetable.keys()))

        df_data = {}

        for day in DAYS:
            row = []
            for p in ALL_PERIODS:
                if p in get_periods(day):
                    subject = st.session_state.timetable.get(sec, {}) \
                        .get(day, {}) \
                        .get(p, {}) \
                        .get("subject", "")
                    teacher = st.session_state.timetable[sec][day][p]["teacher"]
                    value = f"{subject}\n({teacher})" if subject else ""
                    row.append(value)
                else:
                    row.append("")
            df_data[day] = row

        if st.button("⬇ Download PDF"):
            file_path = export_class_timetable_pdf(sec)

            with open(file_path, "rb") as f:
                st.download_button(
                    "Download PDF",
                    f,
                    file_name=file_path,
                    mime="application/pdf"
                )

        # ---- STYLE FIRST ----
        st.markdown("""
        <style>
        [data-testid="stDataFrame"] td {
            white-space: pre-line;
            text-align: center;
            font-size: 14px;
        }
        </style>
        """, unsafe_allow_html=True)

        # ---- DATAFRAME ----
        df = pd.DataFrame(df_data, index=ALL_PERIODS)

        edited_df = st.data_editor(
            df,
            use_container_width=True,
            disabled=not is_admin  # 🔒 role control
        )
        if st.button("⬇ Download Word File"):
            file_path = export_class_timetable_to_word(sec)

            with open(file_path, "rb") as f:
                st.download_button(
                    label="Download Timetable",
                    data=f,
                    file_name=file_path,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

        # ✅ SAVE BUTTON MUST BE HERE

        if st.button("Save Manual Changes"):

            # 🔒 BACKEND SAFETY CHECK
            if st.session_state.get("role") != "admin":
                st.warning("Only admin can modify timetable")
                st.stop()

            for day in DAYS:
                for period in ALL_PERIODS:
                    if period in get_periods(day):

                        cell = edited_df.loc[period, day]

                        if cell:
                            try:
                                if cell:
                                    try:
                                        parts = cell.split("\n")

                                        subject = parts[0].strip()

                                        teacher = ""
                                        if len(parts) > 1:
                                            teacher = parts[1].replace("(", "").replace(")", "").strip()

                                        st.session_state.timetable[sec][day][period]["subject"] = subject
                                        st.session_state.timetable[sec][day][period]["teacher"] = teacher

                                    except Exception as e:
                                        st.warning(f"Format error in {day}-{period}")
                                subject = subject.strip()
                                teacher = teacher.strip().upper()

                                st.session_state.timetable[sec][day][period]["subject"] = subject
                                st.session_state.timetable[sec][day][period]["teacher"] = teacher
                            except:
                                pass
                        else:
                            st.session_state.timetable[sec][day][period]["subject"] = ""
                            st.session_state.timetable[sec][day][period]["teacher"] = ""

            save_all_data()
            st.success("Manual changes saved")
            st.rerun()

    # 🔎 Run validations after manual edit

    for sec_check in st.session_state.timetable:
        subject_issues = validate_subject_weekly(sec_check)
        for issue in subject_issues:
            st.error(issue)

    clash_issues = validate_teacher_clashes()
    for issue in clash_issues:
        st.error(issue)

    class_teacher_issues = validate_class_teacher_presence()
    for issue in class_teacher_issues:
        st.warning(issue)

    max_load_issues = validate_teacher_max_load()
    for issue in max_load_issues:
        st.error(issue)

# ==================================================
# ---------------- TEACHER VIEW --------------------
# ==================================================
if menu == "Teacher View":

    if not st.session_state.timetable:
        st.warning("Generate timetable first")

    else:

        teacher = st.selectbox(
            "Select Teacher",
            list(st.session_state.teachers.keys()),
            key="teacher_view_select"
        )

        df_data = {}

        for day in DAYS:
            row = []

            for p in ALL_PERIODS:

                if p in get_periods(day):

                    found = False

                    # ✅ sec ONLY exists inside loop
                    for sec in st.session_state.timetable:

                        entry = st.session_state.timetable[sec][day][p]

                        if clean_name(entry["teacher"]) == clean_name(teacher):
                            row.append(f"{sec}\n{entry['subject']}")
                            found = True
                            break

                    if not found:
                        row.append("")

                else:
                    row.append("")

            df_data[day] = row

        df = pd.DataFrame(df_data, index=ALL_PERIODS)
        st.dataframe(df)
        st.subheader("⬇ Download Teacher Timetable")

        col1, col2 = st.columns(2)

        with col1:
            file_word = export_teacher_view_word(teacher)
            with open(file_word, "rb") as f:
                st.download_button(
                    "Download Word",
                    f,
                    file_name=file_word,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )

        with col2:
            file_pdf = export_teacher_view_pdf(teacher)
            with open(file_pdf, "rb") as f:
                st.download_button(
                    "Download PDF",
                    f,
                    file_name=file_pdf,
                    mime="application/pdf"
                )
        # ---------------- TOTAL PERIOD COUNT ----------------
        total = 0

        for sec in st.session_state.timetable:
            for day in DAYS:
                for period in get_periods(day):

                    entry = st.session_state.timetable[sec][day][period]

                    if clean(entry["teacher"]) == clean(teacher):
                        total += 1
        st.markdown(f"### 📊 Total Weekly Periods: {total}")
# ==================================================
# ---------------- ANALYTICS -----------------------
# ==================================================
if menu == "Analytics":

    if not st.session_state.timetable:
        st.warning("Generate timetable first")
    else:
        workload = {}

        for teacher in st.session_state.teachers:
            count = 0
            for sec in st.session_state.timetable:
                for day in DAYS:
                    for period in get_periods(day):
                        if st.session_state.timetable[sec][day][period]["teacher"] == teacher:
                            count += 1
            workload[teacher] = count

        df = pd.DataFrame(workload.items(), columns=["Teacher", "Total Periods"])
        st.dataframe(df)


        def workload_color(val):
            if val >= 25:
                return "background-color:#ff6b6b"
            elif val >= 20:
                return "background-color:#ffd166"
            else:
                return "background-color:#90ee90"


        styled_df = df.style.applymap(workload_color, subset=["Total Periods"])

        st.dataframe(styled_df, use_container_width=True)
        st.bar_chart(df.set_index("Teacher"))
st.subheader("📋 All Teachers — One Page View")

df_all = build_all_teachers_matrix()
st.dataframe(df_all, use_container_width=True)

col1, col2 = st.columns(2)

with col1:
    if st.button("⬇ Download Word"):
        file_path = export_all_teachers_word()
        with open(file_path, "rb") as f:
            st.download_button(
                "Download Word File",
                f,
                file_name=file_path,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

with col2:
    if st.button("⬇ Download PDF", key="class_pdf_download"):
        file_path = export_all_teachers_pdf()
        with open(file_path, "rb") as f:
            st.download_button(
                "Download PDF File",
                f,
                file_name=file_path,
                mime="application/pdf"
            )
