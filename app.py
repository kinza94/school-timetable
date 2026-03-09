import streamlit as st

# ✅ FIX 1: set_page_config MUST be the very first Streamlit call
st.set_page_config(page_title="School Scheduler Pro", layout="wide")

import pandas as pd
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import sqlite3
import json
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.shared import Inches
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
import random
import copy
import os

# ==================================================
# ---------------- CONSTANTS -----------------------
# ==================================================

DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]

MON_THU_TIMES = {
    "P1": "8:00-8:50",
    "P2": "8:50-9:30",
    "P3": "9:30-10:10",
    "P4": "10:10-10:50",
    "P5": "10:50-11:30",
    "Lunch": "11:30-12:00",
    "P6": "12:00-12:40",
    "P7": "12:40-1:20",
    "P8": "1:20-2:00",
}

FRIDAY_TIMES = {
    "P1": "8:00-8:40",
    "P2": "8:40-9:15",
    "P3": "9:15-9:50",
    "P4": "9:50-10:25",
    "Lunch": "10:25-10:55",
    "P5": "10:55-11:35",
    "P6": "11:35-12:15",
}

ALL_PERIODS = ["P1", "P2", "P3", "P4", "P5", "Lunch", "P6", "P7", "P8"]

# ✅ FIX 2: Credentials moved to constants (replace with env vars in production)
ADMIN_USERNAME = "admin"
ADMIN_PASSWORD = "Kinz@420"
HEAD_USERNAME = "head"
HEAD_PASSWORD = "9999"
VIEWER_PASSWORD = "1234"

# ==================================================
# ---------------- DB SETUP ------------------------
# ==================================================

conn = sqlite3.connect("school.db", check_same_thread=False)
c = conn.cursor()

c.execute("""
CREATE TABLE IF NOT EXISTS app_data (
    id INTEGER PRIMARY KEY,
    data TEXT
)
""")
conn.commit()

styles_pdf = getSampleStyleSheet()
styleN = styles_pdf["Normal"]

# ==================================================
# -------- CONSTRAINT ENGINE GLOBAL STATE ----------
# ==================================================

teacher_day_load = {}
teacher_timeline = {}
double_used = {}

# ==================================================
# ---------------- HELPERS -------------------------
# ==================================================

# ✅ FIX 3: Removed duplicate clean_name/clean — one function only
def clean(x):
    return str(x).strip().upper()


def get_periods(day):
    if day == "Friday":
        return ["P1", "P2", "P3", "P4", "Lunch", "P5", "P6"]
    return ["P1", "P2", "P3", "P4", "P5", "Lunch", "P6", "P7", "P8"]


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

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

if "role" not in st.session_state:
    st.session_state.role = None

# ==================================================
# ---------------- DB FUNCTIONS --------------------
# ==================================================

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
        c.execute("INSERT INTO app_data (id, data) VALUES (1, ?)", (json.dumps(data),))
    else:
        c.execute("UPDATE app_data SET data=? WHERE id=1", (json.dumps(data),))

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

    # Normalize teacher names to uppercase
    cleaned = {}
    for t in st.session_state.teachers:
        cleaned[clean(t)] = {}
    st.session_state.teachers = cleaned

    for sec in st.session_state.timetable:
        for day in DAYS:
            for p in get_periods(day):
                teacher = st.session_state.timetable[sec][day][p]["teacher"]
                st.session_state.timetable[sec][day][p]["teacher"] = clean(teacher)


# ✅ FIX 4: Load data BEFORE login so teachers dict is ready
if "data_loaded" not in st.session_state:
    load_all_data()
    st.session_state.data_loaded = True

# ==================================================
# ---------------- LOGIN ---------------------------
# ==================================================

if not st.session_state.logged_in:
    st.title("Login")
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")

    if st.button("Login"):
        if username == ADMIN_USERNAME and password == ADMIN_PASSWORD:
            st.session_state.logged_in = True
            st.session_state.role = "admin"
            st.success("Logged in as Admin")
            st.rerun()

        # ✅ FIX 5: teachers is now loaded before this check runs
        elif username in st.session_state.teachers and password == VIEWER_PASSWORD:
            st.session_state.logged_in = True
            st.session_state.role = "viewer"
            st.success(f"Logged in as {username}")
            st.rerun()

        elif username == HEAD_USERNAME and password == HEAD_PASSWORD:
            st.session_state.logged_in = True
            st.session_state.role = "viewer"
            st.success("Logged in as Head")
            st.rerun()

        else:
            st.error("Invalid credentials")

    st.stop()  # Don't render anything else until logged in

# ==================================================
# ---------------- STYLING -------------------------
# ==================================================

st.markdown("""
    <style>
        .main { background-color: #f4f7fb; }
        h1 { color: #1f4e79; }
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
# ---------------- VALIDATION FUNCTIONS ------------
# ==================================================

def validate_subject_weekly(section):
    issues = []
    if section not in st.session_state.subject_config:
        return issues

    required_counts = st.session_state.subject_config[section]
    actual_counts = {}

    for day in DAYS:
        for period in get_periods(day):
            subject = st.session_state.timetable[section][day][period]["subject"]
            if subject:
                actual_counts[subject] = actual_counts.get(subject, 0) + 1

    for subject, required in required_counts.items():
        actual = actual_counts.get(subject, 0)
        if actual != required:
            issues.append(f"{subject} should be {required} periods but currently {actual}")

    return issues


def validate_teacher_clashes():
    issues = []

    for day in DAYS:
        for period in get_periods(day):
            teacher_slots = {}

            for section in st.session_state.timetable:
                teacher = st.session_state.timetable[section][day][period]["teacher"]
                if teacher:
                    teacher_slots.setdefault(teacher, []).append(section)

            for teacher, sections in teacher_slots.items():
                if len(sections) > 1:
                    issues.append(f"Clash: {teacher} has {sections} at {day} {period}")

    return issues


def validate_class_teacher_presence():
    issues = []

    for section, class_teacher in st.session_state.class_teachers.items():
        if section not in st.session_state.timetable:
            continue

        found = False
        for day in DAYS:
            if day not in st.session_state.timetable[section]:
                continue
            for period in get_periods(day):
                if period not in st.session_state.timetable[section][day]:
                    continue
                if st.session_state.timetable[section][day][period]["teacher"] == class_teacher:
                    found = True
                    break
            if found:
                break

        if not found:
            issues.append(f"Class teacher {class_teacher} has no period in {section}")

    return issues


def validate_teacher_distribution():
    issues = []

    for teacher in st.session_state.teachers:
        daily_load = []
        for day in DAYS:
            count = sum(
                1 for sec in st.session_state.timetable
                for period in get_periods(day)
                if st.session_state.timetable[sec][day][period]["teacher"] == teacher
            )
            daily_load.append(count)

        if max(daily_load) - min(daily_load) >= 2:
            issues.append(f"{teacher} workload not balanced: {daily_load}")

    return issues


def validate_friday_load():
    issues = []

    for teacher in st.session_state.teachers:
        count = sum(
            1 for sec in st.session_state.timetable
            for period in get_periods("Friday")
            if st.session_state.timetable[sec]["Friday"][period]["teacher"] == teacher
        )
        if count > 5:
            issues.append(f"{teacher} heavy Friday ({count} periods)")
        elif count == 5:
            issues.append(f"{teacher} Friday 5 periods (acceptable)")

    return issues


def validate_maths_consecutive():
    issues = []

    for sec in st.session_state.timetable:
        for day in DAYS:
            maths_positions = [
                period for period in get_periods(day)
                if "Math" in (st.session_state.timetable[sec][day][period]["subject"] or "")
            ]

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
                count_today = sum(
                    1 for period in get_periods(day)
                    if st.session_state.timetable[section][day][period]["subject"] == subject
                )
                if count_today >= 2:
                    double_days += 1

            if weekly_required >= 6:
                if double_days != 1:
                    issues.append(f"{section} - {subject} must have exactly 1 double period")
            else:
                if double_days > 0:
                    issues.append(f"{section} - {subject} cannot have double period")

    return issues


def validate_teacher_max_load(max_periods=25):
    issues = []

    for teacher in st.session_state.teachers:
        count = count_teacher_periods(teacher)
        if count > max_periods:
            issues.append(f"{teacher} exceeds {max_periods} periods (currently {count})")

    return issues


# ✅ FIX 6: validate_no_three_consecutive is now a real function
def validate_no_three_consecutive():
    issues = []

    for sec in st.session_state.timetable:
        for day in DAYS:
            periods = get_periods(day)
            streak = 0
            prev_teacher = None

            for period in periods:
                if period == "Lunch":
                    streak = 0
                    prev_teacher = None
                    continue

                teacher = st.session_state.timetable[sec][day][period]["teacher"]

                if teacher and teacher == prev_teacher:
                    streak += 1
                    if streak >= 3:
                        issues.append(
                            f"{teacher} has 3+ consecutive periods in {sec} on {day}"
                        )
                else:
                    streak = 1
                    prev_teacher = teacher

    return issues


# ==================================================
# ---------------- HELPER FUNCTIONS ----------------
# ==================================================

# ✅ FIX 7: Extracted repeated period-counting logic into one function
def count_teacher_periods(teacher):
    return sum(
        1 for sec in st.session_state.timetable
        for day in DAYS
        for period in get_periods(day)
        if st.session_state.timetable[sec][day][period]["teacher"] == teacher
    )


def teacher_busy(teacher, day, period):
    for sec in st.session_state.timetable:
        if st.session_state.timetable[sec][day][period]["teacher"] == teacher:
            return True
    return False


def teacher_daily_load(teacher, day):
    return sum(
        1 for sec in st.session_state.timetable
        for period in get_periods(day)
        if st.session_state.timetable[sec][day][period]["teacher"] == teacher
    )


def is_ix_x_double(subject):
    targets = ["PHYSICS", "CHEMISTRY", "COMPUTER IX-X"]
    s = subject.upper()
    return any(t in s for t in targets)


def get_free_teachers(day, period):
    return [t for t in st.session_state.teachers if not teacher_busy(t, day, period)]


def suggest_safe_slots(section, teacher):
    safe = [
        (day, period)
        for day in DAYS
        for period in get_periods(day)
        if period != "Lunch"
        and st.session_state.timetable[section][day][period]["subject"] == ""
        and not teacher_busy(teacher, day, period)
    ]
    return safe[:5]


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


def find_alternative_teacher(section, subject, day, period):
    alternatives = []
    for teacher, sec_data in st.session_state.teacher_assignment.items():
        if section in sec_data and subject in sec_data[section]:
            if teacher == st.session_state.timetable[section][day][period]["teacher"]:
                continue
            if not teacher_busy(teacher, day, period):
                alternatives.append(teacher)
    return alternatives


def find_swap_option(section, teacher, subject, day, period):
    swaps = []
    for d in DAYS:
        for p in get_periods(d):
            if d == day and p == period:
                continue
            if st.session_state.timetable[section][d][p]["teacher"] == teacher:
                if not teacher_busy(teacher, day, period):
                    swaps.append((d, p))
    return swaps


# ==================================================
# ---------------- CONSTRAINT ENGINE ---------------
# ==================================================

def can_assign(section, subject, teacher, day, period):
    periods = get_periods(day)
    idx = periods.index(period)

    if teacher_busy(teacher, day, period):
        return False

    if teacher_day_load[teacher][day] >= 6:
        return False

    # No 4+ consecutive periods
    timeline = teacher_timeline[teacher][day].copy()
    timeline[idx] = 1

    streak = 0
    for v in timeline:
        streak = streak + 1 if v else 0
        if streak >= 4:
            return False

    # IX-X double rule
    if is_ix_x_double(subject):
        prev_same = (
            idx > 0 and
            st.session_state.timetable[section][day][periods[idx - 1]]["subject"] == subject
        )
        next_same = (
            idx < len(periods) - 1 and
            st.session_state.timetable[section][day][periods[idx + 1]]["subject"] == subject
        )
        if (prev_same or next_same) and double_used.get((section, subject), False):
            return False

    return True


def apply_assignment(section, subject, teacher, day, period):
    st.session_state.timetable[section][day][period]["subject"] = subject
    st.session_state.timetable[section][day][period]["teacher"] = teacher

    idx = get_periods(day).index(period)
    teacher_day_load[teacher][day] += 1
    teacher_timeline[teacher][day][idx] = 1

    if is_ix_x_double(subject):
        periods = get_periods(day)
        if idx > 0 and st.session_state.timetable[section][day][periods[idx - 1]]["subject"] == subject:
            double_used[(section, subject)] = True
        if idx < len(periods) - 1 and st.session_state.timetable[section][day][periods[idx + 1]]["subject"] == subject:
            double_used[(section, subject)] = True


def create_empty_timetable():
    timetable = {}

    teacher_day_load.clear()
    teacher_timeline.clear()
    double_used.clear()

    for t in st.session_state.teachers:
        teacher_day_load[t] = {d: 0 for d in DAYS}
        teacher_timeline[t] = {d: [0] * len(get_periods(d)) for d in DAYS}

    for section in st.session_state.sections:
        timetable[section] = {
            day: {
                period: {"subject": "", "teacher": ""}
                for period in get_periods(day)
            }
            for day in DAYS
        }

    return timetable


def assign_class_teacher_priority():
    for section, class_teacher in st.session_state.class_teachers.items():
        if class_teacher not in st.session_state.teacher_assignment:
            continue
        if section not in st.session_state.teacher_assignment[class_teacher]:
            continue

        subjects = st.session_state.teacher_assignment[class_teacher][section]

        for day in DAYS:
            if "P1" not in get_periods(day) or not subjects:
                continue

            subject = subjects[0]
            required = st.session_state.subject_config[section].get(subject, 0)

            current_count = sum(
                1 for d in DAYS for p in get_periods(d)
                if st.session_state.timetable[section][d][p]["subject"] == subject
            )

            if current_count >= required:
                continue

            if (
                st.session_state.timetable[section][day]["P1"]["subject"] == ""
                and can_assign(section, subject, class_teacher, day, "P1")
            ):
                apply_assignment(section, subject, class_teacher, day, "P1")


def calculate_fitness():
    score = 1000
    score -= len(validate_no_three_consecutive()) * 50
    score -= len(validate_teacher_distribution()) * 20
    score -= len(validate_friday_load()) * 10
    return score


def try_swap(section, subject, teacher):
    for day in DAYS:
        for period in get_periods(day):
            if period == "Lunch":
                continue

            current = st.session_state.timetable[section][day][period]
            if current["subject"] == "":
                continue

            other_subject = current["subject"]
            other_teacher = current["teacher"]

            for d2 in DAYS:
                for p2 in get_periods(d2):
                    if p2 == "Lunch":
                        continue

                    target = st.session_state.timetable[section][d2][p2]

                    if (
                        target["subject"] == ""
                        and not teacher_busy(other_teacher, d2, p2)
                        and not teacher_busy(teacher, day, period)
                        and can_assign(section, subject, teacher, day, period)
                    ):
                        # Clear both slots first
                        st.session_state.timetable[section][day][period] = {"subject": "", "teacher": ""}
                        st.session_state.timetable[section][d2][p2] = {"subject": "", "teacher": ""}

                        apply_assignment(section, subject, teacher, day, period)
                        apply_assignment(section, other_subject, other_teacher, d2, p2)
                        return True

    return False


def basic_auto_fill():
    sections = list(st.session_state.subject_config.keys())
    random.shuffle(sections)

    for section in sections:
        subjects = st.session_state.subject_config[section]
        subject_items = list(subjects.items())
        random.shuffle(subject_items)

        for subject, count in subject_items:

            # ✅ FIX 8: Find teacher once, not twice
            assigned_teacher = None
            for teacher, sec_data in st.session_state.teacher_assignment.items():
                if section in sec_data and subject in sec_data[section]:
                    assigned_teacher = teacher
                    break

            if not assigned_teacher:
                continue

            # Daily guarantee for high-frequency subjects
            if count >= 5:
                for day in DAYS:
                    already_present = any(
                        st.session_state.timetable[section][day][p]["subject"] == subject
                        for p in get_periods(day)
                    )
                    if already_present:
                        continue

                    periods = get_periods(day).copy()
                    random.shuffle(periods)
                    valid_periods = [
                        p for p in periods
                        if p != "Lunch"
                        and st.session_state.timetable[section][day][p]["subject"] == ""
                        and can_assign(section, subject, assigned_teacher, day, p)
                    ]

                    if valid_periods:
                        apply_assignment(section, subject, assigned_teacher, day, random.choice(valid_periods))

            # Count already filled
            filled = sum(
                1 for d in DAYS for p in get_periods(d)
                if st.session_state.timetable[section][d][p]["subject"] == subject
            )

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
                        if period == "Lunch":
                            continue
                        if st.session_state.timetable[section][day][period]["subject"] != "":
                            continue
                        if teacher_busy(assigned_teacher, day, period):
                            continue
                        if teacher_daily_load(assigned_teacher, day) >= 6:
                            continue
                        if not can_assign(section, subject, assigned_teacher, day, period):
                            continue
                        valid_slots.append((0, day, period))

                if not valid_slots:
                    swapped = try_swap(section, subject, assigned_teacher)
                    if not swapped:
                        print(f"⚠ Could not place {subject} in {section}")
                        break
                    else:
                        filled += 1
                        continue

                _, best_day, best_period = random.choice(valid_slots)
                apply_assignment(section, subject, assigned_teacher, best_day, best_period)
                subject_day_count[best_day] += 1
                filled += 1


def replace_teacher_everywhere(old_teacher, new_teacher):
    for section in st.session_state.timetable:
        for day in DAYS:
            for period in get_periods(day):
                if st.session_state.timetable[section][day][period]["teacher"] == old_teacher:
                    st.session_state.timetable[section][day][period]["teacher"] = new_teacher
    save_all_data()


# ==================================================
# ---------------- EXPORT FUNCTIONS ----------------
# ==================================================

def export_teacher_view_word(teacher):
    doc = Document()
    doc.add_heading(f"Teacher Timetable — {teacher}", 0)

    total = count_teacher_periods(teacher)
    doc.add_paragraph(f"Total Weekly Periods: {total}")

    table = doc.add_table(rows=len(ALL_PERIODS) + 1, cols=len(DAYS) + 3)
    table.style = "Table Grid"

    headers = ["Period", "Mon-Thu Time", "Mon", "Tue", "Wed", "Thu", "Fri", "Fri Time"]
    for i, h in enumerate(headers):
        table.cell(0, i).text = h

    for r, period in enumerate(ALL_PERIODS):
        table.cell(r + 1, 0).text = period
        table.cell(r + 1, 1).text = MON_THU_TIMES.get(period, "")

        col_index = 2
        for day in DAYS:
            actual_period = period

            if day == "Friday":
                if period == "P5":
                    table.cell(r + 1, col_index).text = "Lunch"
                    col_index += 1
                    continue
                elif period == "Lunch":
                    actual_period = "P5"

            value = ""
            if actual_period in get_periods(day):
                for sec in st.session_state.timetable:
                    entry = st.session_state.timetable[sec][day][actual_period]
                    if clean(entry["teacher"]) == clean(teacher):
                        value = f"{sec}\n{entry['subject']}"
                        break

            table.cell(r + 1, col_index).text = value
            col_index += 1

        table.cell(r + 1, col_index).text = FRIDAY_TIMES.get(period, "")

    filename = f"{teacher}_timetable.docx"
    doc.save(filename)
    return filename


def export_teacher_view_pdf(teacher):
    file_path = f"{teacher}_timetable.pdf"

    doc = SimpleDocTemplate(
        file_path,
        pagesize=landscape(A4),
        leftMargin=20, rightMargin=20, topMargin=30, bottomMargin=30
    )

    elements = []
    styles = getSampleStyleSheet()

    elements.append(Paragraph(f"<b>Teacher Timetable — {teacher}</b>", styles["Title"]))
    elements.append(Spacer(1, 10))

    total = count_teacher_periods(teacher)
    elements.append(Paragraph(f"<b>Total Weekly Periods:</b> {total}", styles["Normal"]))
    elements.append(Spacer(1, 15))

    data = [["Period", "Mon-Thu Time", "Mon", "Tue", "Wed", "Thu", "Fri", "Fri Time"]]

    for period in ALL_PERIODS:
        row = [period, MON_THU_TIMES.get(period, "")]

        for day in DAYS:
            actual_period = period

            if day == "Friday":
                if period == "P5":
                    row.append("Lunch")
                    continue
                elif period == "Lunch":
                    actual_period = "P5"

            value = ""
            if actual_period in get_periods(day):
                for sec in st.session_state.timetable:
                    entry = st.session_state.timetable[sec][day][actual_period]
                    if clean(entry["teacher"]) == clean(teacher):
                        value = f"{sec}\n{entry['subject']}"
                        break

            row.append(value)

        row.append(FRIDAY_TIMES.get(period, ""))
        data.append(row)

    page_width = landscape(A4)[0] - 40
    col_width = page_width / len(data[0])
    table = Table(data, colWidths=[col_width] * len(data[0]), repeatRows=1)

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
        pagesize=landscape(A4),
        rightMargin=40, leftMargin=40, topMargin=40, bottomMargin=40
    )

    elements = []
    styles = getSampleStyleSheet()

    elements.append(Paragraph(f"<b>Class Timetable — {section}</b>", styles["Title"]))
    elements.append(Spacer(1, 10))

    class_teacher = st.session_state.class_teachers.get(section, "Not Assigned")
    elements.append(Paragraph(f"<b>Class Teacher:</b> {class_teacher}", styles["Normal"]))
    elements.append(Spacer(1, 20))

    data = [["Period", "Mon-Thu Time"] + DAYS + ["Fri Time"]]

    for period in ALL_PERIODS:
        row = [period, MON_THU_TIMES.get(period, "")]

        for day in DAYS:
            if period in get_periods(day):
                actual_period = period

                if day == "Friday":
                    if period == "P5":
                        row.append("Lunch")
                        continue
                    elif period == "Lunch":
                        actual_period = "P5"

                cell = st.session_state.timetable[section][day][actual_period]
                text = f"{cell['subject']}\n({cell['teacher']})" if cell["subject"] else ""
            else:
                text = ""

            row.append(text)

        row.append(FRIDAY_TIMES.get(period, ""))
        data.append(row)

    page_width = landscape(A4)[0] - 40
    num_cols = len(data[0])
    col_width = max(page_width / num_cols, 40)
    table = Table(data, colWidths=[col_width] * num_cols, repeatRows=1)

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


# ✅ FIX 9: Added missing export_class_timetable_to_word function
def export_class_timetable_to_word(section):
    doc = Document()

    sec_obj = doc.sections[-1]
    sec_obj.orientation = WD_ORIENT.LANDSCAPE
    sec_obj.page_width, sec_obj.page_height = sec_obj.page_height, sec_obj.page_width

    doc.add_heading(f"Class Timetable — {section}", 0)

    class_teacher = st.session_state.class_teachers.get(section, "Not Assigned")
    doc.add_paragraph(f"Class Teacher: {class_teacher}")

    headers = ["Period", "Mon-Thu Time"] + DAYS + ["Fri Time"]
    table = doc.add_table(rows=len(ALL_PERIODS) + 1, cols=len(headers))
    table.style = "Table Grid"

    for i, h in enumerate(headers):
        table.cell(0, i).text = h

    for r, period in enumerate(ALL_PERIODS):
        table.cell(r + 1, 0).text = period
        table.cell(r + 1, 1).text = MON_THU_TIMES.get(period, "")

        col_index = 2
        for day in DAYS:
            if period in get_periods(day):
                actual_period = period

                if day == "Friday":
                    if period == "P5":
                        table.cell(r + 1, col_index).text = "Lunch"
                        col_index += 1
                        continue
                    elif period == "Lunch":
                        actual_period = "P5"

                cell = st.session_state.timetable[section][day][actual_period]
                text = f"{cell['subject']}\n({cell['teacher']})" if cell["subject"] else ""
            else:
                text = ""

            table.cell(r + 1, col_index).text = text
            col_index += 1

        table.cell(r + 1, col_index).text = FRIDAY_TIMES.get(period, "")

    filename = f"{section}_timetable.docx"
    doc.save(filename)
    return filename


def export_teacher_timetable_pdf(teacher):
    file_path = f"{teacher}_timetable.pdf"
    doc = SimpleDocTemplate(file_path, pagesize=A4)
    elements = []
    styles = getSampleStyleSheet()

    elements.append(Paragraph(f"<b>Teacher Timetable — {teacher}</b>", styles["Title"]))
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
    table = Table(data, colWidths=[col_width] * (len(DAYS) + 1), repeatRows=1)

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


def build_principal_matrix():
    rows = []

    for day in DAYS:
        periods = get_periods(day)
        period_no = 0

        for period in periods:
            if period != "Lunch":
                period_no += 1

            row = {
                "Day": day,
                "P. No.": "" if period == "Lunch" else period_no,
                "Bell Timing": (
                    MON_THU_TIMES.get(period, "") if day != "Friday"
                    else FRIDAY_TIMES.get(period, "")
                ),
            }

            for teacher in st.session_state.teachers:
                value = ""
                for sec in st.session_state.timetable:
                    entry = st.session_state.timetable[sec][day][period]
                    if entry["teacher"] == teacher:
                        value = f"{sec} {entry['subject']}"
                        break
                row[teacher] = value

            rows.append(row)

    df = pd.DataFrame(rows)
    df.loc[df["Day"].duplicated(), "Day"] = ""
    df["P. No."] = df["P. No."].astype(str)
    return df


def export_timetable_word(df):
    file = "school_timetable.docx"
    doc = Document()

    section = doc.sections[-1]
    section.orientation = WD_ORIENT.LANDSCAPE
    section.page_width, section.page_height = section.page_height, section.page_width

    days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]
    rows_per_day = 9

    for i, day in enumerate(days):
        day_df = df.iloc[i * rows_per_day:(i + 1) * rows_per_day]
        doc.add_heading(day, level=1)

        rows = len(day_df) + 1
        cols = len(day_df.columns)

        table = doc.add_table(rows=rows, cols=cols)
        table.style = "Table Grid"

        for j, col in enumerate(day_df.columns):
            table.rows[0].cells[j].text = str(col)

        for r, (_, row) in enumerate(day_df.iterrows(), start=1):
            for col_i, val in enumerate(row):
                table.rows[r].cells[col_i].text = str(val)

        doc.add_paragraph("")

    doc.save(file)
    return file


def export_excel(df):
    file = "School_Timetable.xlsx"

    with pd.ExcelWriter(file, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Timetable", index=False)

        wb = writer.book
        ws = writer.sheets["Timetable"]
        ws.freeze_panes = "A2"

        header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.alignment = Alignment(horizontal="center", vertical="center")
            cell.fill = header_fill

        border = Border(
            left=Side(style="thin"), right=Side(style="thin"),
            top=Side(style="thin"), bottom=Side(style="thin")
        )

        for row in ws.iter_rows():
            for cell in row:
                cell.border = border
                cell.alignment = Alignment(horizontal="center", vertical="center")

        for col in ws.columns:
            max_length = 0
            letter = get_column_letter(col[0].column)
            for cell in col:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except Exception:
                    pass
            ws.column_dimensions[letter].width = max_length + 4

        lunch_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
        for row in range(2, ws.max_row + 1):
            if ws.cell(row=row, column=2).value == "":
                for col in range(1, ws.max_column + 1):
                    ws.cell(row=row, column=col).fill = lunch_fill

    return file


# ==================================================
# ---------------- SIDEBAR -------------------------
# ==================================================

is_admin = st.session_state.get("role") == "admin"

if is_admin:
    menu = st.sidebar.selectbox(
        "Navigation",
        ["Dashboard", "Configuration", "Generate", "Class View", "Teacher View", "Analytics"]
    )
else:
    menu = st.sidebar.selectbox(
        "Navigation",
        ["Class View", "Teacher View", "Analytics"]
    )

if st.sidebar.button("Logout"):
    st.session_state.logged_in = False
    st.session_state.role = None
    st.rerun()

# ==================================================
# ---------------- DASHBOARD -----------------------
# ==================================================

if menu == "Dashboard":
    st.subheader("CREATE YOUR TIMETABLE 👋")
    st.write("Use the sidebar to configure and generate the timetable.")

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
                save_all_data()
                st.success("Section Removed")

    with col2:
        st.subheader("➕ Add Teacher")
        teacher = st.text_input("Teacher Name")
        if st.button("Add Teacher"):
            if teacher:
                st.session_state.teachers[clean(teacher)] = {}
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
                save_all_data()
                st.success("Teacher Removed")

    st.subheader("📌 Assign Class Teacher")

    if st.session_state.sections and st.session_state.teachers:
        selected_sec = st.selectbox(
            "Select Section",
            list(st.session_state.sections.keys()),
            key="class_teacher_section"
        )
        selected_teacher = st.selectbox(
            "Select Teacher",
            list(st.session_state.teachers.keys())
        )

        if st.button("Assign Class Teacher"):
            st.session_state.class_teachers[selected_sec] = selected_teacher
            save_all_data()
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
            save_all_data()
            st.success("Class Teacher Removed")

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
                save_all_data()
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

            subjects_available = list(st.session_state.subject_config.get(assign_section, {}).keys())

            if subjects_available:
                assign_subject = st.selectbox(
                    "Select Subject",
                    subjects_available,
                    key="assign_subject_select"
                )

                if st.button("Assign Teacher"):
                    if assign_teacher not in st.session_state.teacher_assignment:
                        st.session_state.teacher_assignment[assign_teacher] = {}

                    if assign_section not in st.session_state.teacher_assignment[assign_teacher]:
                        st.session_state.teacher_assignment[assign_teacher][assign_section] = []

                    if assign_subject not in st.session_state.teacher_assignment[assign_teacher][assign_section]:
                        st.session_state.teacher_assignment[assign_teacher][assign_section].append(assign_subject)

                    # ✅ FIX 10: Save after assigning teacher
                    save_all_data()
                    st.success("Teacher Assigned Successfully")

        st.write("### Current Teacher Assignments")
        st.write(st.session_state.teacher_assignment)

        st.subheader("❌ Remove Teacher Assignment")

        if st.session_state.teacher_assignment:
            remove_teacher_assign = st.selectbox(
                "Select Teacher",
                list(st.session_state.teacher_assignment.keys()),
                key="remove_assign_teacher"
            )

            sections_for_teacher = list(
                st.session_state.teacher_assignment[remove_teacher_assign].keys()
            )

            if sections_for_teacher:
                remove_section_assign = st.selectbox(
                    "Select Section",
                    sections_for_teacher,
                    key="remove_assign_section"
                )

                subjects_for_teacher = st.session_state.teacher_assignment[remove_teacher_assign][remove_section_assign]

                if subjects_for_teacher:
                    remove_subject_assign = st.selectbox(
                        "Select Subject",
                        subjects_for_teacher,
                        key="remove_assign_subject"
                    )

                    if st.button("Delete Assignment"):
                        st.session_state.teacher_assignment[remove_teacher_assign][remove_section_assign].remove(
                            remove_subject_assign
                        )

                        if not st.session_state.teacher_assignment[remove_teacher_assign][remove_section_assign]:
                            del st.session_state.teacher_assignment[remove_teacher_assign][remove_section_assign]

                        if not st.session_state.teacher_assignment[remove_teacher_assign]:
                            del st.session_state.teacher_assignment[remove_teacher_assign]

                        save_all_data()
                        st.success("Assignment Removed Successfully")

# ==================================================
# ---------------- GENERATE ------------------------
# ==================================================

# ✅ FIX 11: Generate menu is now a proper top-level block (was buried inside validate_no_three_consecutive)
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
        save_all_data()
        st.success(f"Timetable Generated Successfully (fitness score: {best_score})")

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
            st.success(f"{old_teacher} replaced with {new_teacher}")

    # Soft constraint validations (admin only, shown in Generate menu)
    if st.session_state.timetable:
        st.subheader("🔍 Validation Report")

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

        if not any([consecutive_issues, distribution_issues, friday_issues, maths_issues]):
            st.success("✅ No constraint violations found!")

# ==================================================
# ---------------- CLASS VIEW ----------------------
# ==================================================

if menu == "Class View":

    if not st.session_state.timetable:
        st.warning("Generate timetable first")
    else:
        sec = st.selectbox("Select Section", list(st.session_state.timetable.keys()))
        class_teacher = st.session_state.class_teachers.get(sec, "Not Assigned")
        st.markdown(f"### 👩‍🏫 Class Teacher: {class_teacher}")

        df_data = {}

        for day in DAYS:
            row = []
            for p in ALL_PERIODS:
                if p in get_periods(day):
                    subject = st.session_state.timetable.get(sec, {}).get(day, {}).get(p, {}).get("subject", "")
                    teacher = st.session_state.timetable[sec][day][p]["teacher"]
                    value = f"{subject}\n({teacher})" if subject else ""
                    row.append(value)
                else:
                    row.append("")
            df_data[day] = row

        # ---- Style ----
        st.markdown("""
        <style>
        [data-testid="stDataFrame"] td {
            white-space: pre-line;
            text-align: center;
            font-size: 14px;
        }
        </style>
        """, unsafe_allow_html=True)

        display_periods = ["P1", "P2", "P3", "P4", "P5", "Lunch", "P6", "P7", "P8"]
        df = pd.DataFrame(df_data, index=display_periods)
        df.insert(0, "Mon-Thu Time", [MON_THU_TIMES.get(p, "") for p in display_periods])

        # Friday adjustments
        p5 = df.loc["P5", "Friday"]
        df.loc["P5", "Friday"] = "Lunch"
        df.loc["Lunch", "Friday"] = p5

        fri_times = [
            "8:00-8:40", "8:40-9:15", "9:15-9:50", "9:50-10:25",
            "Lunch\n10:25-10:55", "10:55-11:35", "11:35-12:15", "", ""
        ]
        df.insert(6, "Fri Time", fri_times)

        edited_df = st.data_editor(
            df,
            use_container_width=True,
            disabled=not is_admin
        )

        col1, col2 = st.columns(2)

        with col1:
            if st.button("⬇ Download PDF"):
                file_path = export_class_timetable_pdf(sec)
                with open(file_path, "rb") as f:
                    st.download_button(
                        "Download PDF",
                        f,
                        file_name=file_path,
                        mime="application/pdf"
                    )

        with col2:
            if st.button("⬇ Download Word File"):
                file_path = export_class_timetable_to_word(sec)
                with open(file_path, "rb") as f:
                    st.download_button(
                        label="Download Timetable Word",
                        data=f,
                        file_name=file_path,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                    )

        if is_admin and st.button("Save Manual Changes"):
            for day in DAYS:
                for period in ALL_PERIODS:
                    if period in get_periods(day):
                        cell = edited_df.loc[period, day]

                        if cell:
                            try:
                                parts = cell.split("\n")
                                subject = parts[0].strip()
                                teacher = ""
                                if len(parts) > 1:
                                    teacher = parts[1].replace("(", "").replace(")", "").strip().upper()

                                st.session_state.timetable[sec][day][period]["subject"] = subject
                                st.session_state.timetable[sec][day][period]["teacher"] = teacher
                            except Exception as e:
                                st.warning(f"Format error in {day}-{period}: {e}")
                        else:
                            st.session_state.timetable[sec][day][period]["subject"] = ""
                            st.session_state.timetable[sec][day][period]["teacher"] = ""

            save_all_data()
            st.success("Manual changes saved")
            st.rerun()

        # Validations after class view
        for sec_check in st.session_state.timetable:
            for issue in validate_subject_weekly(sec_check):
                st.error(issue)

        for issue in validate_teacher_clashes():
            st.error(issue)

        for issue in validate_class_teacher_presence():
            st.warning(issue)

        for issue in validate_teacher_max_load():
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
                found = False
                actual_p = p

                if day == "Friday":
                    if p == "P5":
                        row.append("Lunch")
                        continue
                    elif p == "Lunch":
                        actual_p = "P5"

                if actual_p in get_periods(day):
                    for sec in st.session_state.timetable:
                        entry = st.session_state.timetable[sec][day][actual_p]
                        if clean(entry["teacher"]) == clean(teacher):
                            row.append(f"{sec}\n{entry['subject']}")
                            found = True
                            break
                    if not found:
                        row.append("")
                else:
                    row.append("")

            df_data[day] = row

        df = pd.DataFrame(df_data, index=ALL_PERIODS)
        df.insert(0, "Mon-Thu Time", [MON_THU_TIMES.get(p, "") for p in ALL_PERIODS])
        df.insert(len(DAYS) + 1, "Fri Time", [FRIDAY_TIMES.get(p, "") for p in ALL_PERIODS])
        st.dataframe(df)

        total = count_teacher_periods(teacher)
        st.markdown(f"### 📊 Total Weekly Periods: {total}")

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

# ==================================================
# ---------------- ANALYTICS -----------------------
# ==================================================

if menu == "Analytics":

    if not st.session_state.timetable:
        st.warning("Generate timetable first")
    else:
        workload = {
            teacher: count_teacher_periods(teacher)
            for teacher in st.session_state.teachers
        }

        df = pd.DataFrame(workload.items(), columns=["Teacher", "Total Periods"])

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

        st.subheader("School Master Timetable")
        df_master = build_principal_matrix()
        st.dataframe(df_master, use_container_width=True)

        # ✅ FIX 12: Export buttons are inside Analytics menu, not at module level
        st.subheader("⬇ Download Master Timetable")
        col1, col2 = st.columns(2)

        with col1:
            file_word = export_timetable_word(df_master)
            with open(file_word, "rb") as f:
                st.download_button(
                    "Download Word Timetable",
                    f,
                    file_name="school_timetable.docx"
                )

        with col2:
            file_excel = export_excel(df_master)
            with open(file_excel, "rb") as f:
                st.download_button(
                    "Download Excel Timetable",
                    f,
                    file_name="School_Timetable.xlsx"
                )
