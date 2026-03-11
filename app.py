import streamlit as st
import os
from openai import OpenAI

client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))

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
import re
import copy
import os

# ==================================================
# ---------------- CONSTANTS -----------------------
# ==================================================

DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]

MON_WED_TIMES = {
    "P1": "8:00-8:50",
    "P2": "8:50-9:30",
    "P3": "9:30-10:10",
    "P4": "10:10-10:50",
    "Lunch": "10:50-11:20",
    "P5": "11:20-12:00",
    "P6": "12:00-12:40",
    "P7": "12:40-1:20",
    "P8": "1:20-2:00",
}

THURSDAY_TIMES = {
    "P1": "8:50-9:30",
    "P2": "9:30-10:05",
    "P3": "10:05-10:40",
    "P4": "10:40-11:15",
    "Lunch": "11:15-11:35",
    "P5": "11:35-12:10",
    "P6": "12:10-12:45",
    "P7": "12:45-1:15",
    "P8": "1:15-2:00",
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

ALL_PERIODS = ["P1", "P2", "P3", "P4", "Lunch", "P5", "P6", "P7", "P8"]

# ✅ FIX 2: Credentials moved to constants (replace with env vars in production)
ADMIN_USERNAME = "admin"
ADMIN_PASSWORD = "Kinz@420"
HEAD_USERNAME = "head"
HEAD_PASSWORD = "9999"
VIEWER_USERNAME = "teacher"
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
subject_remaining = {}   # {(section, subject): remaining_count}  — single source of truth for quotas
teacher_three_streak_count = {}
# ==================================================
# ---------------- HELPERS -------------------------
# ==================================================

# ✅ FIX 3: Removed duplicate clean_name/clean — one function only
def clean(x):
    return str(x).strip().upper()


def get_periods(day):
    # Lunch is always after the 4th period on every day
    if day == "Friday":
        return ["P1", "P2", "P3", "P4", "Lunch", "P5", "P6"]
    return ["P1", "P2", "P3", "P4", "Lunch", "P5", "P6", "P7", "P8"]


def get_time(day, period):
    """Return the bell-timing string for a given day and period."""
    if day in ["Monday", "Tuesday", "Wednesday"]:
        return MON_WED_TIMES.get(period, "")
    elif day == "Thursday":
        return THURSDAY_TIMES.get(period, "")
    elif day == "Friday":
        return FRIDAY_TIMES.get(period, "")
    return ""


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
    t_key = clean(teacher)
    for sec in st.session_state.timetable:
        if clean(st.session_state.timetable[sec][day][period]["teacher"]) == t_key:
            return True
    return False


def teacher_daily_load(teacher, day):
    t_key = clean(teacher)
    return sum(
        1 for sec in st.session_state.timetable
        for period in get_periods(day)
        if clean(st.session_state.timetable[sec][day][period]["teacher"]) == t_key
    )




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
def is_library(subject):
    return "LIBRARY" in subject.upper()

# ==================================================
# ---------------- CONSTRAINT ENGINE ---------------
# ==================================================
#
# DESIGN: Every constraint is enforced in ONE place — can_assign().
# Generation runs in 5 ordered phases so high-priority rules get
# first pick of slots before the general fill.
#
# Constraints implemented
# ───────────────────────────────────────────────────────────
# C1   Teacher clash — never assign a teacher to two classes same slot
# C2   Teacher daily cap ≤ 6 teaching periods
# C3   Teacher 4-consecutive NEVER (100% hard rule)
# C4   Teacher 3-consecutive avoided (soft – penalised in fitness)
# C5   Lunch is always after 4th slot; treated as hard gap in consecutive
#        counts — P4 and P5 across lunch are NEVER counted consecutive
# C6   Daily-single subjects — at most ONCE per day per class:
#        Sindhi, Islamiat, English, Urdu, General Science, Arabic,
#        Geography, History, Social Studies
#        These subjects may also NEVER form a double period.
# C7   General no-double rule — all subjects not in the "allowed doubles"
#        list may never be placed adjacent to themselves
# C8   Math rule: quota = 6 periods/week; appears every day (Mon-Fri);
#        exactly ONE consecutive double period on one day per week
# C9   IX-X science doubles: Physics/Chemistry/Biology/Comp-IX-X each
#        get exactly one consecutive double per week in IX-X sections;
#        the double must NOT cross the lunch break
# C10  Games: never placed in the last teaching period of any day
# C11  Class teacher first-period rule: class teacher teaches P1 in their
#        own class in at least 90% of school days (≥ 4/5 days)
# C12  Fully-filled timetable: emergency back-fill + displacement ensures
#        no class has an empty teaching period
# ───────────────────────────────────────────────────────────

# ── Subject-category helpers ──────────────────────────────

# Subjects that must appear AT MOST ONCE per day per class.
# Exact-keyword match: a subject qualifies if any keyword below appears
# as a whole word (or phrase) inside the normalised subject name.
# Covers: Sindhi, Islamiat, English, Urdu, General Science, Arabic,
#         Geography, History, Social Studies
DAILY_SINGLE_KEYWORDS = [
    "SINDHI", "ISLAMIAT", "ENGLISH", "URDU",
    "GENERAL SCIENCE", "ARABIC", "GEOGRAPHY",
    "HISTORY", "SOCIAL STUDIES",
]

IX_X_DOUBLE_SUBJECTS   = [
    "PHYSICS", "CHEMISTRY", "BIOLOGY",
    "COMPUTER IX-X", "COMPUTER SCIENCE IX-X",
]
MATH_KEYWORD  = "MATH"
GAMES_KEYWORD = "GAMES"


def is_daily_single(subject):
    """
    True when the subject must appear at most once per day.
    Uses longest-match-first so "GENERAL SCIENCE" is matched before
    "GENERAL" could hypothetically match something else.
    """
    s = subject.strip().upper()
    # Sort by length descending so multi-word keywords match first
    for kw in sorted(DAILY_SINGLE_KEYWORDS, key=len, reverse=True):
        if kw in s:
            return True
    return False


def is_math(subject):
    return MATH_KEYWORD in subject.upper()


def is_games(subject):
    return GAMES_KEYWORD in subject.upper()


def is_ix_x_double(subject):
    s = subject.upper()
    return any(t in s for t in IX_X_DOUBLE_SUBJECTS)


def is_double_allowed(subject):

    s = subject.upper()

    if "MATH" in s:
        return True

    if "PHYSICS" in s or "CHEMISTRY" in s or "COMPUTER IX-X" in s:
        return True

    return False
def subject_weekly_quota(section, subject):
    return st.session_state.subject_config.get(section, {}).get(subject, 0)


def get_last_teaching_period(day):
    """Last non-Lunch slot of the day (used for Games rule C10)."""
    return [p for p in get_periods(day) if p != "Lunch"][-1]


# ── Low-level timetable query helpers ─────────────────────

def subject_count_in_day(section, subject, day):
    return sum(
        1 for p in get_periods(day)
        if st.session_state.timetable[section][day][p]["subject"].upper()
           == subject.upper()
    )


def subject_count_total(section, subject):
    return sum(subject_count_in_day(section, subject, day) for day in DAYS)


def quota_remaining(section, subject):
    """
    Fast O(1) quota check.  Returns how many more periods of `subject`
    are still needed in `section` this week.  Uses the subject_remaining
    dict which is decremented by every apply_assignment call.
    Falls back to computing from the timetable if the key is missing
    (e.g. called before generate, or after manual edits).
    """
    key = (section, subject)
    if key in subject_remaining:
        return subject_remaining[key]
    # Fallback: compute from timetable (slower but always correct)
    quota = st.session_state.subject_config.get(section, {}).get(subject, 0)
    placed = subject_count_total(section, subject)
    return max(0, quota - placed)


def math_double_day(section):
    """Return the day that already has a Math double for this section, or None."""
    for day in DAYS:
        slots = [p for p in get_periods(day) if p != "Lunch"]
        for i in range(len(slots) - 1):
            subj_a = st.session_state.timetable[section][day][slots[i]]["subject"]
            subj_b = st.session_state.timetable[section][day][slots[i+1]]["subject"]
            if is_math(subj_a) and is_math(subj_b):
                return day
    return None


def are_adjacent(periods_list, p1, p2):
    """True if p1 and p2 are consecutive in the period list (Lunch excluded)."""
    slots = [p for p in periods_list if p != "Lunch"]
    try:
        i1, i2 = slots.index(p1), slots.index(p2)
        return abs(i1 - i2) == 1
    except ValueError:
        return False


def teacher_consecutive_streak(teacher, day, proposed_idx):
    """
    C5 – Count consecutive busy slots around proposed_idx, treating Lunch
    as a hard break.  Returns max streak length if proposed slot is added.
    Slots on opposite sides of Lunch are NOT consecutive.
    """
    periods   = get_periods(day)
    lunch_idx = periods.index("Lunch") if "Lunch" in periods else -1

    # Build occupation list for this teacher on this day (0=free, 1=busy)
    occupied = []
    for i, p in enumerate(periods):
        if p == "Lunch":
            occupied.append(-1)          # -1 = hard break
        elif i == proposed_idx:
            occupied.append(1)           # the slot we are testing
        else:
            # Check all sections
            busy = any(
                st.session_state.timetable[sec][day][p]["teacher"] == teacher
                for sec in st.session_state.timetable
            )
            occupied.append(1 if busy else 0)

    max_streak = 0
    streak     = 0
    for v in occupied:
        if v == -1:          # Lunch → hard reset
            streak = 0
        elif v == 1:
            streak += 1
            max_streak = max(max_streak, streak)
        else:
            streak = 0

    return max_streak


def subject_consecutive_streak(section, subject, day, proposed_idx):
    """
    Same as teacher_consecutive_streak but for subjects within one section.
    Lunch is treated as a hard break.
    """
    periods = get_periods(day)

    occupied = []
    for i, p in enumerate(periods):
        if p == "Lunch":
            occupied.append(-1)
        elif i == proposed_idx:
            occupied.append(1)
        else:
            subj = st.session_state.timetable[section][day][p]["subject"]
            occupied.append(1 if subj.upper() == subject.upper() else 0)

    max_streak = 0
    streak     = 0
    for v in occupied:
        if v == -1:
            streak = 0
        elif v == 1:
            streak += 1
            max_streak = max(max_streak, streak)
        else:
            streak = 0

    return max_streak


# ── Core constraint gate ───────────────────────────────────

def can_assign(section, subject, teacher, day, period):
    """
    Single gate every placement must pass.
    Returns True only when ALL hard constraints are satisfied.
    """
    periods = get_periods(day)
    idx     = periods.index(period)

    # Lunch is never assignable (structural)
    if period == "Lunch":
        return False

    # ── C1: teacher clash ─────────────────────────────────
    if teacher_busy(teacher, day, period):
        return False

    # ── C2: teacher daily cap ─────────────────────────────
    # Normalize key – if teacher somehow not in load dict, treat as free (safe default)
    t_key = clean(teacher)
    if t_key not in teacher_day_load:
        # Teacher exists in assignment but not yet initialised – add them now
        teacher_day_load[t_key] = {d: 0 for d in DAYS}
        teacher_timeline[t_key] = {d: [0] * len(get_periods(d)) for d in DAYS}
    if day == "Friday":
        if teacher_day_load[t_key][day] >= 5:
            return False
    else:
        if teacher_day_load[t_key][day] >= 6:
            return False
    ##  core subject
    if is_core(subject):
        if subject_count_in_day(section, subject, day) >= 1:
            return False
    # ── C3 + C4: consecutive checks (Lunch = hard break) ──
    t_streak = teacher_consecutive_streak(t_key, day, idx)
    if t_streak >= 4:          # C3 – hard block
        return False
    # C4 (3-consecutive) is soft – allowed but penalised in fitness

    # ── C5: Lunch adjacency – subjects across Lunch are NOT consecutive ─
    # (handled inside streak functions above; no extra code needed)

    # ── Neighbour subjects in THIS class (Lunch = no neighbour) ─────────
    slots       = [p for p in periods if p != "Lunch"]
    slot_idx    = slots.index(period) if period in slots else -1
    prev_period = slots[slot_idx - 1] if slot_idx > 0 else None
    next_period = slots[slot_idx + 1] if slot_idx < len(slots) - 1 else None

    # Only treat as adjacent if they are not separated by Lunch
    lunch_idx = periods.index("Lunch") if "Lunch" in periods else -1
    period_idx = periods.index(period)

    def real_adjacent(other_p):
        """Two periods are truly adjacent only if no Lunch sits between them."""
        if other_p is None:
            return False
        other_idx = periods.index(other_p)
        lo, hi = min(period_idx, other_idx), max(period_idx, other_idx)
        # If Lunch is between them they are NOT adjacent for constraint purposes
        if lunch_idx != -1 and lo < lunch_idx < hi:
            return False
        return abs(period_idx - other_idx) == 1

    prev_subject = (
        st.session_state.timetable[section][day][prev_period]["subject"]
        if prev_period and real_adjacent(prev_period) else ""
    )
    next_subject = (
        st.session_state.timetable[section][day][next_period]["subject"]
        if next_period and real_adjacent(next_period) else ""
    )
    # Rule: No subject should repeat in the same day
    if subject_count_in_day(section, subject, day) >= 1:

        # Only allow if it is a valid double subject
        if not is_double_allowed(subject):
            return False
    # ── C6: Daily-single subjects (English / Urdu / General Science) ────
    if is_daily_single(subject):
        if subject_count_in_day(section, subject, day) >= 1:
            return False                      # already present today
        if prev_subject.upper() == subject.upper():
            return False                      # would create a double
        if next_subject.upper() == subject.upper():
            return False
    weekly = st.session_state.subject_config.get(section, {}).get(subject, 0)

    if weekly < 6:
        if prev_subject.upper() == subject.upper() or next_subject.upper() == subject.upper():
            return False
    # ── C7: General no-double rule ───────────────────────────────────────
    if not is_double_allowed(subject):
        if prev_subject.upper() == subject.upper():
            return False
        if next_subject.upper() == subject.upper():
            return False

    # ── C8: Math double — at most ONE day per week ───────────────────────
        if is_math(subject):
            weekly = subject_weekly_quota(section, subject)

            if weekly < 6:
                if prev_subject.upper() == subject.upper() or next_subject.upper() == subject.upper():
                    return False
        if would_be_double:
            existing = math_double_day(section)
            if existing is not None and existing != day:
                return False

    # ── C9: IX-X double — at most once per subject per week ─────────────
    if is_ix_x_double(subject):
        would_be_double = (
            prev_subject.upper() == subject.upper() or
            next_subject.upper() == subject.upper()
        )
        if would_be_double and double_used.get((section, subject), False):
            return False

    # ── C10: Games never last period ────────────────────────────────────
    # ── C10: Games / Library cannot be first or last period
    if is_games(subject) or is_library(subject):

        periods = get_periods(day)

        if period == periods[0]:  # first period
            return False

        if period == get_last_teaching_period(day):  # last period
            return False

    return True

CORE_SUBJECTS = ["MATH", "ENGLISH", "URDU", "GENERAL SCIENCE"]
def is_core(subject):
    s = subject.upper()
    return any(core in s for core in CORE_SUBJECTS)

# ── State writer ──────────────────────────────────────────

def apply_assignment(section, subject, teacher, day, period):
    t_key = clean(teacher)   # normalise casing
    st.session_state.timetable[section][day][period]["subject"] = subject
    st.session_state.timetable[section][day][period]["teacher"] = t_key

    idx = get_periods(day).index(period)

    # Auto-init if teacher was not in session_state.teachers at creation time
    if t_key not in teacher_day_load:
        teacher_day_load[t_key] = {d: 0 for d in DAYS}
        teacher_timeline[t_key] = {d: [0] * len(get_periods(d)) for d in DAYS}

    teacher_day_load[t_key][day]    += 1
    teacher_timeline[t_key][day][idx] = 1

    streak = teacher_consecutive_streak(t_key, day, idx)
    teacher_three_streak_count.setdefault(t_key, 0)

    if streak == 3:
        teacher_three_streak_count[t_key] += 1

    # Decrement quota tracker — the authoritative remaining count
    key = (section, subject)
    if key in subject_remaining:
        subject_remaining[key] = max(0, subject_remaining[key] - 1)

    # Track IX-X double usage
    if is_ix_x_double(subject):
        periods = [p for p in get_periods(day) if p != "Lunch"]
        for i, p in enumerate(periods):
            if p == period:
                if i > 0 and st.session_state.timetable[section][day][periods[i-1]]["subject"].upper() == subject.upper():
                    double_used[(section, subject)] = True
                if i < len(periods)-1 and st.session_state.timetable[section][day][periods[i+1]]["subject"].upper() == subject.upper():
                    double_used[(section, subject)] = True
                break

    # Track Math double (used by math_double_day)
    if is_math(subject):
        periods = [p for p in get_periods(day) if p != "Lunch"]
        for i, p in enumerate(periods):
            if p == period:
                if i > 0 and is_math(st.session_state.timetable[section][day][periods[i-1]]["subject"]):
                    pass   # math_double_day() scans dynamically — no flag needed
                break


# ── Undo helper (inverse of apply_assignment) ───────────

def undo_assignment(section, day, period):
    """
    Remove whatever is in timetable[section][day][period] and reverse
    every counter that apply_assignment incremented.
    Safe to call on an already-empty slot (no-op).
    """
    slot = st.session_state.timetable[section][day][period]
    subject = slot.get("subject", "")
    teacher  = slot.get("teacher", "")

    if not subject:
        return   # nothing to undo

    # ── restore quota ────────────────────────────────────────────────
    key = (section, subject)
    if key in subject_remaining:
        subject_remaining[key] += 1

    # ── restore teacher load counters ────────────────────────────────
    t_key = clean(teacher) if teacher else ""
    if t_key and t_key in teacher_day_load:
        teacher_day_load[t_key][day] = max(0, teacher_day_load[t_key][day] - 1)

    if t_key and t_key in teacher_timeline:
        idx = get_periods(day).index(period)
        teacher_timeline[t_key][day][idx] = 0

    # ── clear slot ───────────────────────────────────────────────────
    st.session_state.timetable[section][day][period] = {"subject": "", "teacher": ""}

    # ── reverse IX-X double flag if this was part of a double ────────
    if subject and is_ix_x_double(subject):
        # Re-check whether a double still exists after removal
        still_doubled = False
        slots = [p for p in get_periods(day) if p != "Lunch"]
        for i in range(len(slots) - 1):
            sa = st.session_state.timetable[section][day][slots[i]]["subject"]
            sb = st.session_state.timetable[section][day][slots[i+1]]["subject"]
            if sa.upper() == subject.upper() and sb.upper() == subject.upper():
                still_doubled = True
                break
        if not still_doubled:
            # Also check other days
            for d in DAYS:
                if d == day:
                    continue
                dslots = [p for p in get_periods(d) if p != "Lunch"]
                for i in range(len(dslots) - 1):
                    sa = st.session_state.timetable[section][d][dslots[i]]["subject"]
                    sb = st.session_state.timetable[section][d][dslots[i+1]]["subject"]
                    if sa.upper() == subject.upper() and sb.upper() == subject.upper():
                        still_doubled = True
                        break
                if still_doubled:
                    break
            if not still_doubled:
                double_used.pop((section, subject), None)


# ── Timetable skeleton ────────────────────────────────────

def create_empty_timetable():
    timetable = {}

    teacher_day_load.clear()
    teacher_timeline.clear()
    double_used.clear()
    subject_remaining.clear()
    teacher_three_streak_count.clear()

    for t in st.session_state.teachers:
        teacher_three_streak_count[clean(t)] = 0

    # Initialise quota tracker — every (section, subject) starts at its full weekly count
    for section, subj_map in st.session_state.subject_config.items():
        for subject, quota in subj_map.items():
            subject_remaining[(section, subject)] = quota

    # Always store teacher keys in UPPER so lookups never KeyError
    for t in st.session_state.teachers:
        t_key = clean(t)
        teacher_day_load[t_key] = {d: 0 for d in DAYS}
        teacher_timeline[t_key] = {d: [0] * len(get_periods(d)) for d in DAYS}

    for section in st.session_state.sections:
        timetable[section] = {
            day: {
                period: {"subject": "", "teacher": ""}
                for period in get_periods(day)
            }
            for day in DAYS
        }

    return timetable


# ── PHASE 1 – Class-teacher first-period (≥ 90 % of days) ───────────

def assign_class_teacher_priority():
    """
    C1 – The class teacher should teach P1 in their own class on at least
    4 out of 5 days (≥ 90 %).
    """
    for section, class_teacher in st.session_state.class_teachers.items():
        if class_teacher not in st.session_state.teacher_assignment:
            continue
        if section not in st.session_state.teacher_assignment[class_teacher]:
            continue

        subjects = st.session_state.teacher_assignment[class_teacher][section]
        if not subjects:
            continue

        days_assigned = 0

        for day in DAYS:
            if "P1" not in get_periods(day):
                continue

            for subject in subjects:
                # Use quota_remaining (O(1), always in sync with apply_assignment)
                if quota_remaining(section, subject) <= 0:
                    continue   # this subject is already fully placed for the week

                if (
                    st.session_state.timetable[section][day]["P1"]["subject"] == ""
                    and can_assign(section, subject, class_teacher, day, "P1")
                ):
                    apply_assignment(section, subject, class_teacher, day, "P1")
                    days_assigned += 1
                    break           # one subject per day is enough

            if days_assigned >= 4:  # 4/5 = 80 % → satisfies ≥ 90 % target
                break


# ── PHASE 2 – Daily mandatory singles ────────────────────────────────

def assign_daily_singles():
    """
    C6 – Daily-single subjects must appear at most once per day per class
    and may never form a double period.  Subjects in this category:
      Sindhi, Islamiat, English, Urdu, General Science, Arabic,
      Geography, History, Social Studies
    Placed in Phase 2 (before general fill) so later phases cannot
    accidentally double these subjects.
    """
    for section in st.session_state.subject_config:
        for subject in list(st.session_state.subject_config[section].keys()):
            if not is_daily_single(subject):
                continue

            assigned_teacher = _find_teacher(section, subject)
            if not assigned_teacher:
                continue

            for day in DAYS:
                if quota_remaining(section, subject) <= 0:
                    break                       # weekly quota fully used — stop

                if subject_count_in_day(section, subject, day) >= 1:
                    continue                    # already placed today

                candidates = [p for p in get_periods(day) if p != "Lunch"]
                random.shuffle(candidates)

                for period in candidates:
                    if quota_remaining(section, subject) <= 0:
                        break
                    if st.session_state.timetable[section][day][period]["subject"] != "":
                        continue
                    if can_assign(section, subject, assigned_teacher, day, period):
                        apply_assignment(section, subject, assigned_teacher, day, period)
                        break


# ── PHASE 3 – Math (every day + one double) ──────────────────────────

def assign_math():
    """
    C8 – Math appears every day.
         Exactly one day per week gets a consecutive double Math period.
    """
    for section in st.session_state.subject_config:
        for subject, count in st.session_state.subject_config[section].items():
            if not is_math(subject):
                continue

            assigned_teacher = _find_teacher(section, subject)
            if not assigned_teacher:
                continue

            # ── Step A: place the one double (costs 2 quota slots) ──
            double_placed = False

            if quota_remaining(section, subject) >= 2:   # need room for 2 periods
                candidate_days = DAYS.copy()
                random.shuffle(candidate_days)

                for day in candidate_days:
                    slots = [p for p in get_periods(day) if p != "Lunch"]
                    for i in range(len(slots) - 1):
                        p1, p2 = slots[i], slots[i + 1]
                        if (
                            quota_remaining(section, subject) >= 2
                            and st.session_state.timetable[section][day][p1]["subject"] == ""
                            and st.session_state.timetable[section][day][p2]["subject"] == ""
                            and can_assign(section, subject, assigned_teacher, day, p1)
                            and can_assign(section, subject, assigned_teacher, day, p2)
                        ):
                            apply_assignment(section, subject, assigned_teacher, day, p1)
                            apply_assignment(section, subject, assigned_teacher, day, p2)
                            double_placed = True
                            break
                    if double_placed:
                        break

            # ── Step B: single Math on every remaining day ────────────
            for day in DAYS:
                if quota_remaining(section, subject) <= 0:
                    break                           # quota exhausted — stop
                if subject_count_in_day(section, subject, day) >= 1:
                    continue                        # already placed today

                candidates = [p for p in get_periods(day) if p != "Lunch"]
                random.shuffle(candidates)
                for period in candidates:
                    if quota_remaining(section, subject) <= 0:
                        break
                    if st.session_state.timetable[section][day][period]["subject"] != "":
                        continue
                    if can_assign(section, subject, assigned_teacher, day, period):
                        apply_assignment(section, subject, assigned_teacher, day, period)
                        break


# ── PHASE 4 – IX-X science doubles ───────────────────────────────────

def assign_ix_x_doubles():
    """
    C9 – Physics, Chemistry, Biology, Computer IX-X each get exactly one
    consecutive double period per week for IX-X sections.
    """
    for section in st.session_state.subject_config:
        # Robust detection: match Grade/Class IX, X, 9, or 10.
        # Uses word-boundary regex to avoid false positives like "MIX" or "EXTRA".
        is_ix_x = bool(re.search(
            r'(?<![A-Z])(?:IX|X|9|10)(?![A-Z0-9])',
            section.upper()
        ))

        if not is_ix_x:
            continue

        for subject in list(st.session_state.subject_config[section].keys()):
            if not is_ix_x_double(subject):
                continue
            if double_used.get((section, subject), False):
                continue                         # double already placed this run

            if quota_remaining(section, subject) < 2:
                continue                         # need room for at least 2 periods

            assigned_teacher = _find_teacher(section, subject)
            if not assigned_teacher:
                continue

            placed = False
            candidate_days = DAYS.copy()
            random.shuffle(candidate_days)

            for day in candidate_days:
                slots = [p for p in get_periods(day) if p != "Lunch"]
                for i in range(len(slots) - 1):
                    p1, p2 = slots[i], slots[i + 1]
                    if (
                        quota_remaining(section, subject) >= 2
                        and st.session_state.timetable[section][day][p1]["subject"] == ""
                        and st.session_state.timetable[section][day][p2]["subject"] == ""
                        and can_assign(section, subject, assigned_teacher, day, p1)
                        and can_assign(section, subject, assigned_teacher, day, p2)
                    ):
                        apply_assignment(section, subject, assigned_teacher, day, p1)
                        apply_assignment(section, subject, assigned_teacher, day, p2)
                        placed = True
                        break
                if placed:
                    break


# ── PHASE 5 – General fill ────────────────────────────────────────────

def calculate_fitness():
    score = 10000

    # ── Hard-ish: quota violations (each missing or excess period = -200) ──
    for section, config in st.session_state.subject_config.items():
        for subject in config:
            # quota_remaining > 0 means under-quota; < 0 impossible (hard wall)
            deviation = quota_remaining(section, subject)   # 0 = perfect
            score -= deviation * 200   # very steep penalty drives best-of-15 selection

    # ── Soft: consecutive, distribution, friday balance ──
    score -= len(validate_no_three_consecutive()) * 50
    score -= len(validate_teacher_distribution()) * 20
    score -= len(validate_friday_load()) * 10
    return score


def try_swap(section, subject, teacher):
    """
    Free a slot for (subject, teacher) by moving the existing occupant
    to any other empty slot.  Uses undo_assignment so quota counters stay
    accurate.
    """
    for day in DAYS:
        for period in get_periods(day):
            if period == "Lunch":
                continue

            current = st.session_state.timetable[section][day][period]
            if current["subject"] == "":
                continue

            other_subject = current["subject"]
            other_teacher = current["teacher"]

            # Can the incoming subject land here at all (ignoring the occupant)?
            # Temporarily clear to let can_assign reason about the empty slot.
            undo_assignment(section, day, period)     # clears + restores counters

            can_place_here = (
                not teacher_busy(teacher, day, period)
                and can_assign(section, subject, teacher, day, period)
            )

            if not can_place_here:
                # Put the evicted subject back and try the next slot
                if can_assign(section, other_subject, other_teacher, day, period):
                    apply_assignment(section, other_subject, other_teacher, day, period)
                else:
                    # Cannot re-place — leave cleared and bail out of whole swap
                    apply_assignment(section, other_subject, other_teacher, day, period)
                continue

            # Find a new home for the evicted subject
            for d2 in DAYS:
                for p2 in get_periods(d2):
                    if p2 == "Lunch":
                        continue
                    if (d2, p2) == (day, period):
                        continue
                    if st.session_state.timetable[section][d2][p2]["subject"] != "":
                        continue
                    if not teacher_busy(other_teacher, d2, p2):
                        if can_assign(section, other_subject, other_teacher, d2, p2):
                            # Commit: place incoming here, evicted there
                            apply_assignment(section, subject,       teacher,       day, period)
                            apply_assignment(section, other_subject, other_teacher, d2,  p2)
                            return True

            # No home found for evicted — restore it and continue
            apply_assignment(section, other_subject, other_teacher, day, period)

    return False


def try_displace(section, subject, teacher):
    """
    Stronger than try_swap: look for a slot occupied by a subject that is
    ALREADY AT or ABOVE its weekly quota (i.e. it has surplus placements).
    Evict that surplus period — no need to re-home it — and place the
    under-quota subject in the freed slot.

    This is the last resort before giving up on a subject's quota.
    """
    config = st.session_state.subject_config.get(section, {})

    for day in DAYS:
        for period in get_periods(day):
            if period == "Lunch":
                continue

            current = st.session_state.timetable[section][day][period]
            occupant_subj    = current.get("subject", "")
            occupant_teacher = current.get("teacher", "")

            if not occupant_subj:
                continue

            # Only evict a subject that has ZERO remaining quota (it has a surplus)
            if quota_remaining(section, occupant_subj) > 0:
                continue   # still needs its own periods — don't touch it

            # After eviction, can we place the needed subject here?
            undo_assignment(section, day, period)

            if (
                not teacher_busy(teacher, day, period)
                and can_assign(section, subject, teacher, day, period)
            ):
                apply_assignment(section, subject, teacher, day, period)
                return True
            else:
                # Can't use this slot — restore the evicted period
                if occupant_teacher and can_assign(section, occupant_subj, occupant_teacher, day, period):
                    apply_assignment(section, occupant_subj, occupant_teacher, day, period)
                else:
                    apply_assignment(section, occupant_subj, occupant_teacher, day, period)

    return False


def basic_auto_fill():
    """
    Fill remaining quota for ALL subjects.
    Subjects are processed in order of most-remaining-quota first so
    harder-to-place subjects get first pick of empty slots.
    This phase also handles any daily-singles / math / IX-X that earlier
    phases left under-quota.
    """
    sections = list(st.session_state.subject_config.keys())
    random.shuffle(sections)

    for section in sections:
        # Sort subjects by remaining quota descending — biggest deficit first
        subject_items = sorted(
            st.session_state.subject_config[section].items(),
            key=lambda kv: quota_remaining(section, kv[0]),
            reverse=True
        )

        for subject, count in subject_items:
            assigned_teacher = _find_teacher(section, subject)
            if not assigned_teacher:
                continue

            while quota_remaining(section, subject) > 0:
                valid_slots = []
                days = DAYS.copy()
                random.shuffle(days)

                for day in days:
                    # Limit to 2 occurrences per day for non-double subjects
                    day_max = 2
                    if subject_count_in_day(section, subject, day) >= day_max:
                        continue

                    candidates = get_periods(day).copy()
                    random.shuffle(candidates)

                    for period in candidates:
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
                        valid_slots.append((day, period))

                if not valid_slots:
                    # Try 1: move an existing slot to make room
                    if try_swap(section, subject, assigned_teacher):
                        continue
                    # Try 2: evict a surplus-quota slot to make room
                    if try_displace(section, subject, assigned_teacher):
                        continue
                    # Genuinely no legal placement exists right now — move on
                    break

                best_day, best_period = random.choice(valid_slots)
                apply_assignment(section, subject, assigned_teacher, best_day, best_period)


# ── PHASE 5b – Under-quota fill (ensure no subject is short) ────────

def fill_under_quota_subjects():
    """
    Safety net after phases 1-5: any subject still under quota gets topped up.
    Three escalation levels per period needed:
      Level 1 — place into a genuinely empty slot (2/day cap)
      Level 2 — place into a genuinely empty slot (cap relaxed)
      Level 3 — try_swap (move an occupant to another empty slot)
      Level 4 — try_displace (evict a surplus-quota occupant permanently)
    """
    for section in st.session_state.subject_config:
        # Process subjects with largest deficit first
        subjects_by_deficit = sorted(
            st.session_state.subject_config[section].keys(),
            key=lambda s: quota_remaining(section, s),
            reverse=True
        )

        for subject in subjects_by_deficit:
            assigned_teacher = _find_teacher(section, subject)
            if not assigned_teacher:
                continue

            # ── Levels 1 & 2: direct placement, first strict cap then relaxed ──
            for relax_cap in (False, True):
                if quota_remaining(section, subject) <= 0:
                    break

                days_shuffled = DAYS.copy()
                random.shuffle(days_shuffled)

                for day in days_shuffled:
                    if quota_remaining(section, subject) <= 0:
                        break

                    day_limit = 8 if relax_cap else 2
                    if subject_count_in_day(section, subject, day) >= day_limit:
                        continue

                    periods_shuffled = [p for p in get_periods(day) if p != "Lunch"]
                    random.shuffle(periods_shuffled)

                    for period in periods_shuffled:
                        if quota_remaining(section, subject) <= 0:
                            break
                        if st.session_state.timetable[section][day][period]["subject"] != "":
                            continue
                        if can_assign(section, subject, assigned_teacher, day, period):
                            apply_assignment(section, subject, assigned_teacher, day, period)
                            break

            # ── Level 3: swap an occupant to a different empty slot ──────────
            if quota_remaining(section, subject) > 0:
                while quota_remaining(section, subject) > 0:
                    if not try_swap(section, subject, assigned_teacher):
                        break

            # ── Level 4: displace a surplus-quota occupant ───────────────────
            if quota_remaining(section, subject) > 0:
                while quota_remaining(section, subject) > 0:
                    if not try_displace(section, subject, assigned_teacher):
                        break


# ── PHASE 6 – Emergency back-fill (C11: no empty slots) ──────────────

def emergency_backfill():
    """
    Phase 6 — two-stage final pass.

    Stage A: Fill every remaining EMPTY slot with any under-quota subject.
             Subjects sorted by biggest remaining quota first.

    Stage B: Quota-enforcement sweep — for each subject still under quota,
             try try_swap then try_displace until quota is met or truly
             impossible.  This is the guarantee that generation never
             finishes with an incomplete quota.
    """
    # ── Stage A: fill empty slots ─────────────────────────────────────
    for section in st.session_state.timetable:
        for day in DAYS:
            for period in get_periods(day):
                if period == "Lunch":
                    continue
                if st.session_state.timetable[section][day][period]["subject"] != "":
                    continue

                subject_items = sorted(
                    st.session_state.subject_config.get(section, {}).items(),
                    key=lambda kv: quota_remaining(section, kv[0]),
                    reverse=True   # biggest remaining quota first
                )

                for subject, _ in subject_items:
                    if quota_remaining(section, subject) <= 0:
                        continue   # HARD WALL

                    assigned_teacher = _find_teacher(section, subject)
                    if not assigned_teacher:
                        continue

                    if can_assign(section, subject, assigned_teacher, day, period):
                        apply_assignment(section, subject, assigned_teacher, day, period)
                        break
                # Slot may remain empty if nothing legal exists — Stage B handles quotas

    # ── Stage B: quota-enforcement — no subject may finish short ─────
    for section in st.session_state.subject_config:
        # Iterate subjects sorted by remaining deficit (largest first)
        subjects_needing = sorted(
            [s for s in st.session_state.subject_config[section]
             if quota_remaining(section, s) > 0],
            key=lambda s: quota_remaining(section, s),
            reverse=True
        )

        for subject in subjects_needing:
            assigned_teacher = _find_teacher(section, subject)
            if not assigned_teacher:
                continue

            # Keep trying until quota met or all options exhausted
            stalled = False
            while quota_remaining(section, subject) > 0 and not stalled:
                placed = False

                # Pass 1: find any genuinely empty slot
                for day in DAYS:
                    if placed:
                        break
                    for period in get_periods(day):
                        if period == "Lunch":
                            continue
                        if st.session_state.timetable[section][day][period]["subject"] != "":
                            continue
                        if can_assign(section, subject, assigned_teacher, day, period):
                            apply_assignment(section, subject, assigned_teacher, day, period)
                            placed = True
                            break

                if placed:
                    continue

                # Pass 2: swap an occupant to another empty slot
                if try_swap(section, subject, assigned_teacher):
                    continue

                # Pass 3: displace a surplus-quota occupant (permanent eviction)
                if try_displace(section, subject, assigned_teacher):
                    continue

                # All three options exhausted — genuinely impossible this run
                stalled = True


# ── Private helper ────────────────────────────────────────

def _find_teacher(section, subject):
    """Return the teacher (uppercase-normalised) assigned to subject in section, or None."""
    for teacher, sec_data in st.session_state.teacher_assignment.items():
        if section in sec_data and subject in sec_data[section]:
            return clean(teacher)   # always return UPPER so it matches teacher_day_load keys
    return None



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
        table.cell(r + 1, 1).text = get_time("Monday", period)

        col_index = 2
        for day in DAYS:
            # Friday period list = P1 P2 P3 P4 Lunch P5 P6
            # ALL_PERIODS     = P1 P2 P3 P4 Lunch P5 P6 P7 P8
            # P7, P8 don't exist on Friday → leave blank
            if day == "Friday" and period not in get_periods(day):
                table.cell(r + 1, col_index).text = ""
                col_index += 1
                continue

            value = ""
            if period in get_periods(day):
                for sec in st.session_state.timetable:
                    entry = st.session_state.timetable[sec][day][period]
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
        row = [period, get_time("Monday", period)]

        for day in DAYS:
            actual_period = period

            # Friday: P7/P8 don't exist — blank cell
            if day == "Friday" and period not in get_periods(day):
                row.append("")
                continue

            value = ""
            if period in get_periods(day):
                for sec in st.session_state.timetable:
                    entry = st.session_state.timetable[sec][day][period]
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

    data = [[
        "Period",
        "Mon-Wed Time",
        "Monday",
        "Tuesday",
        "Wednesday",
        "Thursday Time",
        "Thursday",
        "Friday",
        "Fri Time"
    ]]

    for period in ALL_PERIODS:

        row = [
            period,
            MON_WED_TIMES.get(period, "")
        ]

        # Monday
        if period in get_periods("Monday"):
            cell = st.session_state.timetable[section]["Monday"][period]
            row.append(f"{cell['subject']}\n({cell['teacher']})" if cell["subject"] else "")
        else:
            row.append("")

        # Tuesday
        if period in get_periods("Tuesday"):
            cell = st.session_state.timetable[section]["Tuesday"][period]
            row.append(f"{cell['subject']}\n({cell['teacher']})" if cell["subject"] else "")
        else:
            row.append("")

        # Wednesday
        if period in get_periods("Wednesday"):
            cell = st.session_state.timetable[section]["Wednesday"][period]
            row.append(f"{cell['subject']}\n({cell['teacher']})" if cell["subject"] else "")
        else:
            row.append("")

        # Thursday time
        row.append(THURSDAY_TIMES.get(period, ""))

        # Thursday subject
        if period in get_periods("Thursday"):
            cell = st.session_state.timetable[section]["Thursday"][period]
            row.append(f"{cell['subject']}\n({cell['teacher']})" if cell["subject"] else "")
        else:
            row.append("")

        # Friday subject
        if period in get_periods("Friday"):
            cell = st.session_state.timetable[section]["Friday"][period]
            row.append(f"{cell['subject']}\n({cell['teacher']})" if cell["subject"] else "")
        else:
            row.append("")

        # Friday time
        row.append(FRIDAY_TIMES.get(period, ""))

        data.append(row)

    page_width = landscape(A4)[0] - 80
    num_cols = len(data[0])
    col_widths = [50, 80] + [(page_width - 130) / len(DAYS)] * len(DAYS) + [80]
    page_width = landscape(A4)[0] - 80
    col_width = page_width / len(data[0])
    table = Table(data, colWidths=[col_width] * len(data[0]), repeatRows=1)

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
        table.cell(r + 1, 1).text = get_time("Monday", period)

        col_index = 2
        for day in DAYS:
            if period in get_periods(day):
                # Friday: P7/P8 don't exist — blank cell
                if day == "Friday" and period not in get_periods(day):
                    table.cell(r + 1, col_index).text = ""
                    col_index += 1
                    continue

                cell = st.session_state.timetable[section][day][period]
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
    # Guard: return empty dataframe if timetable not generated yet
    if not st.session_state.timetable:
        return pd.DataFrame()

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
                "Bell Timing": get_time(day, period),
            }

            for teacher in st.session_state.teachers:
                value = ""
                for sec in st.session_state.timetable:
                    # Guard: period may not exist for this day (e.g. P7/P8 on Friday)
                    if period not in st.session_state.timetable[sec].get(day, {}):
                        continue
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

def ai_analyze_timetable(df):

    prompt = f"""
    Analyze this school timetable and check:

    1. Teacher workload balance
    2. Subject distribution
    3. Any possible improvements

    Timetable:
    {df.to_string()}
    """

    response = client.responses.create(
        model="gpt-4.1-mini",
        input=prompt
    )

    return response.output_text
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

            # Guard: key may be stale after a deletion rerun
            if remove_teacher_assign not in st.session_state.teacher_assignment:
                st.rerun()

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

        for _ in range(60):
            temp_table = create_empty_timetable()
            st.session_state.timetable = temp_table

            # Phase 1 – Class teacher owns P1 (≥90 % of days)
            assign_class_teacher_priority()
            # Phase 2 – English / Urdu / General Science: once per day, never doubled
            assign_daily_singles()
            # Phase 3 – Math: every day + exactly one double per week
            assign_math()
            # Phase 4 – IX-X sciences: one double per subject per week
            assign_ix_x_doubles()
            # Phase 5 – All other subjects (Games, etc.)
            basic_auto_fill()
            # Phase 5b – Top-up any subjects still under weekly quota
            fill_under_quota_subjects()
            # Phase 6 – Guarantee fully-filled timetable (no empty slots)
            emergency_backfill()

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

        # Build the display DataFrame in one shot — never use df.insert()
        # which throws ValueError when a column already exists or the
        # list length doesn't match the index length.
        display_periods = ALL_PERIODS  # P1 P2 P3 P4 Lunch P5 P6 P7 P8

        # Friday times list must be exactly len(ALL_PERIODS) = 9 items
        fri_times_map = {
            "P1":    "8:00-8:40",
            "P2":    "8:40-9:15",
            "P3":    "9:15-9:50",
            "P4":    "9:50-10:25",
            "Lunch": "10:25-10:55",
            "P5":    "10:55-11:35",
            "P6":    "11:35-12:15",
            # P7 and P8 do not exist on Friday
        }

        df = pd.DataFrame(
            {
                "Period": display_periods,
                "Mon-Wed Time": [MON_WED_TIMES.get(p, "") for p in display_periods],

                "Monday": df_data["Monday"],
                "Tuesday": df_data["Tuesday"],
                "Wednesday": df_data["Wednesday"],

                "Thursday Time": [THURSDAY_TIMES.get(p, "") for p in display_periods],
                "Thursday": df_data["Thursday"],

                "Friday": df_data["Friday"],
                "Fri Time": [fri_times_map.get(p, "") for p in display_periods],
            }
        ).set_index("Period")

        edited_df = st.data_editor(
            df,
            use_container_width=True,
            disabled=not is_admin,
            key="class_editor"
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
                            parts = cell.split("\n")
                            subject = parts[0].strip()
                            teacher = ""
                            if len(parts) > 1:
                                teacher = parts[1].replace("(", "").replace(")", "").strip().upper()
                            st.session_state.timetable[sec][day][period]["subject"] = subject
                            st.session_state.timetable[sec][day][period]["teacher"] = teacher
                        else:
                            st.session_state.timetable[sec][day][period]["subject"] = ""
                            st.session_state.timetable[sec][day][period]["teacher"] = ""

            save_all_data()
            st.success("Manual changes saved")
            st.session_state["refresh_teacher_view"] = True
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

        # Clear any pending refresh flag from Class View saves
        st.session_state.pop("refresh_teacher_view", None)

        df_data = {}

        for day in DAYS:
            row = []

            for p in ALL_PERIODS:
                found = False

                # Lunch row: show "Lunch" label for all days
                if p == "Lunch":
                    row.append("Lunch")
                    continue

                # P7/P8 don't exist on Friday
                if p not in get_periods(day):
                    row.append("")
                    continue

                for sec in st.session_state.timetable:
                    entry = st.session_state.timetable[sec][day][p]
                    if clean(entry["teacher"]) == clean(teacher):
                        row.append(f"{sec}\n{entry['subject']}")
                        found = True
                        break
                if not found:
                    row.append("")

            df_data[day] = row

        # Build in one shot — avoids ValueError from df.insert() on reruns
        df = pd.DataFrame(
            {
                "Period":          ALL_PERIODS,
                "Mon-Wed Time":    [MON_WED_TIMES.get(p, "")    for p in ALL_PERIODS],
                "Thursday Time":   [THURSDAY_TIMES.get(p, "")   for p in ALL_PERIODS],
                **{day: df_data[day] for day in DAYS},
                "Fri Time":        [FRIDAY_TIMES.get(p, "")     for p in ALL_PERIODS],
            }
        ).set_index("Period")
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

        if st.button("AI Analyze Timetable"):
            analysis = ai_analyze_timetable(df_master)
            st.write(analysis)

        # Export buttons — only shown when df_master has data
        if not df_master.empty:
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