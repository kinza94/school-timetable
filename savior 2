import streamlit as st
st.set_page_config(page_title="School Scheduler Pro", layout="wide", page_icon="📚")

import pandas as pd
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import sqlite3, json, random, re, copy, os
try:
    from openai import OpenAI as _OpenAI
    _openai_key = os.getenv("OPENAI_API_KEY", "")
    openai_client = _OpenAI(api_key=_openai_key) if _openai_key else None
except ImportError:
    openai_client = None
from docx import Document
from docx.enum.section import WD_ORIENT
from docx.shared import Inches
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.styles import getSampleStyleSheet

# ══════════════════════════════════════════════════════════════════════════════
# CONSTANTS
# ══════════════════════════════════════════════════════════════════════════════

DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]

MON_WED_TIMES = {
    "P1": "8:00-8:50",   "P2": "8:50-9:30",   "P3": "9:30-10:10",
    "P4": "10:10-10:50", "Lunch": "10:50-11:20",
    "P5": "11:20-12:00", "P6": "12:00-12:40",
    "P7": "12:40-1:20",  "P8": "1:20-2:00",
}
THURSDAY_TIMES = {
    "P1": "8:50-9:30",   "P2": "9:30-10:05",  "P3": "10:05-10:40",
    "P4": "10:40-11:15", "Lunch": "11:15-11:35",
    "P5": "11:35-12:10", "P6": "12:10-12:45",
    "P7": "12:45-1:15",  "P8": "1:15-2:00",
}
FRIDAY_TIMES = {
    "P1": "8:00-8:40",  "P2": "8:40-9:15",   "P3": "9:15-9:50",
    "P4": "9:50-10:25", "Lunch": "10:25-10:55",
    "P5": "10:55-11:35","P6": "11:35-12:15",
}
ALL_PERIODS   = ["P1","P2","P3","P4","Lunch","P5","P6","P7","P8"]
BRAND_BLUE    = "#1f4e79"

# Credentials
ADMIN_USERNAME  = "admin";  ADMIN_PASSWORD  = "Kinz@420"
HEAD_USERNAME   = "head";   HEAD_PASSWORD   = "9999"
VIEWER_PASSWORD = "1234"

# ══════════════════════════════════════════════════════════════════════════════
# DATABASE
# ══════════════════════════════════════════════════════════════════════════════

conn = sqlite3.connect("school.db", check_same_thread=False)
cur  = conn.cursor()
cur.execute("""CREATE TABLE IF NOT EXISTS app_data(id INTEGER PRIMARY KEY, data TEXT)""")
conn.commit()

# ══════════════════════════════════════════════════════════════════════════════
# CONSTRAINT ENGINE GLOBAL STATE  (reset each generation run)
# ══════════════════════════════════════════════════════════════════════════════

teacher_day_load  = {}   # {t_key: {day: int}}
teacher_timeline  = {}   # {t_key: {day: [0|1, ...]}}
double_used       = {}   # {(section, subject): bool}
subject_remaining = {}   # {(section, subject): int}  — single source of truth
teacher_three_streak_count = {}  # {t_key: count of 3-consecutive occurrences — used for fitness}

# ══════════════════════════════════════════════════════════════════════════════
# CORE HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def clean(x):
    return str(x).strip().upper()

def get_periods(day):
    if day == "Friday":
        return ["P1","P2","P3","P4","Lunch","P5","P6"]
    return ["P1","P2","P3","P4","Lunch","P5","P6","P7","P8"]

def get_time(day, period):
    if day in ("Monday","Tuesday","Wednesday"): return MON_WED_TIMES.get(period,"")
    if day == "Thursday":                       return THURSDAY_TIMES.get(period,"")
    if day == "Friday":                         return FRIDAY_TIMES.get(period,"")
    return ""

def slot(section, day, period):
    """Shortcut: return the timetable cell dict."""
    return st.session_state.timetable[section][day][period]

# ══════════════════════════════════════════════════════════════════════════════
# SESSION STATE
# ══════════════════════════════════════════════════════════════════════════════

_DEFAULTS = {
    "teachers":{}, "sections":{}, "class_teachers":{},
    "subject_config":{}, "teacher_assignment":{}, "timetable":{},
    "logged_in":False, "role":None,
}
for _k, _v in _DEFAULTS.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v

# ══════════════════════════════════════════════════════════════════════════════
# PERSISTENCE
# ══════════════════════════════════════════════════════════════════════════════

def save_all_data():
    data = json.dumps({k: st.session_state[k] for k in
        ("teachers","sections","class_teachers","subject_config",
         "teacher_assignment","timetable")})
    cur.execute("SELECT COUNT(*) FROM app_data")
    if cur.fetchone()[0] == 0:
        cur.execute("INSERT INTO app_data(id,data) VALUES(1,?)", (data,))
    else:
        cur.execute("UPDATE app_data SET data=? WHERE id=1", (data,))
    conn.commit()

def load_all_data():
    cur.execute("SELECT data FROM app_data LIMIT 1")
    row = cur.fetchone()
    if row:
        data = json.loads(row[0])
        for key in ("teachers","sections","class_teachers",
                    "subject_config","teacher_assignment","timetable"):
            st.session_state[key] = data.get(key, {})
    st.session_state.teachers = {clean(t):{} for t in st.session_state.teachers}
    for sec in st.session_state.timetable:
        for day in DAYS:
            for p in get_periods(day):
                cell = st.session_state.timetable[sec][day][p]
                cell["teacher"] = clean(cell.get("teacher",""))

if "data_loaded" not in st.session_state:
    load_all_data()
    st.session_state.data_loaded = True

# ══════════════════════════════════════════════════════════════════════════════
# LOGIN
# ══════════════════════════════════════════════════════════════════════════════

if not st.session_state.logged_in:
    st.markdown("<h1 style='text-align:center;color:#1f4e79;'>DHACSS PHASE IV CAMPUS</h1>", unsafe_allow_html=True)
    st.markdown("<h3 style='text-align:center;'>Academic Scheduling Engine</h3>", unsafe_allow_html=True)

    st.title("🔐 Login")
    _, col, _ = st.columns([1,2,1])
    with col:
        username  = st.text_input("Username")
        password  = st.text_input("Password", type="password")
        login_btn = st.button("Login", use_container_width=True)
    if login_btn:
        if   username == ADMIN_USERNAME and password == ADMIN_PASSWORD:  role = "admin"
        elif username == HEAD_USERNAME  and password == HEAD_PASSWORD:   role = "viewer"
        elif clean(username) in st.session_state.teachers and password == VIEWER_PASSWORD:
            role = "viewer"
        else:
            role = None
        if role:
            st.session_state.logged_in = True
            st.session_state.role = role
            st.rerun()
        else:
            st.error("Invalid credentials.")
    st.stop()

# ══════════════════════════════════════════════════════════════════════════════
# STYLING
# ══════════════════════════════════════════════════════════════════════════════

st.markdown("""
<style>
.main{background:#f4f7fb}
h1{color:#1f4e79}
.stButton>button{
    background:#4a90e2;color:white;border-radius:8px;
    height:3em;width:100%;font-weight:600}
.stButton>button:hover{background:#357abd}
[data-testid="stDataFrame"] td{white-space:pre-line;text-align:center;font-size:14px}
</style>""", unsafe_allow_html=True)

st.title("📚 School Timetable Scheduler Pro")

# ══════════════════════════════════════════════════════════════════════════════
# VALIDATION
# ══════════════════════════════════════════════════════════════════════════════

def validate_subject_weekly(section):
    issues, actual = [], {}
    for day in DAYS:
        for p in get_periods(day):
            s = slot(section, day, p)["subject"]
            if s: actual[s] = actual.get(s, 0) + 1
    for subj, req in st.session_state.subject_config.get(section, {}).items():
        got = actual.get(subj, 0)
        if got != req:
            issues.append(f"{subj}: needs {req} periods, has {got}")
    return issues

def validate_teacher_clashes():
    issues = []
    for day in DAYS:
        for p in get_periods(day):
            seen = {}
            for sec in st.session_state.timetable:
                t = slot(sec, day, p)["teacher"]
                if t: seen.setdefault(t, []).append(sec)
            for t, secs in seen.items():
                if len(secs) > 1:
                    issues.append(f"Clash: {t} in {secs} at {day} {p}")
    return issues

def validate_class_teacher_presence():
    issues = []
    for sec, ct in st.session_state.class_teachers.items():
        if sec not in st.session_state.timetable: continue
        found = any(slot(sec, day, p)["teacher"] == ct
                    for day in DAYS if day in st.session_state.timetable[sec]
                    for p in get_periods(day) if p in st.session_state.timetable[sec][day])
        if not found:
            issues.append(f"Class teacher {ct} has no period in {sec}")
    return issues

def validate_teacher_distribution():
    issues = []
    for t in st.session_state.teachers:
        loads = [sum(1 for sec in st.session_state.timetable for p in get_periods(day)
                     if slot(sec, day, p)["teacher"] == t) for day in DAYS]
        if max(loads) - min(loads) >= 2:
            issues.append(f"{t} daily load imbalanced: {loads}")
    return issues

def validate_friday_load():
    issues = []
    for t in st.session_state.teachers:
        n = sum(1 for sec in st.session_state.timetable for p in get_periods("Friday")
                if slot(sec, "Friday", p)["teacher"] == t)
        if n > 5: issues.append(f"{t}: heavy Friday ({n} periods)")
        elif n == 5: issues.append(f"{t}: Friday 5 periods (acceptable)")
    return issues

def validate_maths_consecutive():
    issues = []
    for sec in st.session_state.timetable:
        for day in DAYS:
            mps = [p for p in get_periods(day) if "Math" in (slot(sec,day,p)["subject"] or "")]
            if len(mps) == 2:
                ps = get_periods(day)
                if abs(ps.index(mps[0]) - ps.index(mps[1])) != 1:
                    issues.append(f"{sec}: Maths not consecutive on {day}")
    return issues

def validate_teacher_max_load(max_p=25):
    return [f"{t} exceeds {max_p} periods (has {count_teacher_periods(t)})"
            for t in st.session_state.teachers if count_teacher_periods(t) > max_p]

def validate_no_three_consecutive():
    issues = []
    for sec in st.session_state.timetable:
        for day in DAYS:
            streak, prev_t = 0, None
            for p in get_periods(day):
                if p == "Lunch": streak, prev_t = 0, None; continue
                t = slot(sec, day, p)["teacher"]
                if t and t == prev_t:
                    streak += 1
                    if streak >= 3: issues.append(f"{t}: 3+ consecutive in {sec} on {day}")
                else:
                    streak, prev_t = 1, t
    return issues

# ══════════════════════════════════════════════════════════════════════════════
# TIMETABLE QUERY HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def count_teacher_periods(teacher):
    return sum(1 for sec in st.session_state.timetable for day in DAYS
               for p in get_periods(day) if slot(sec, day, p)["teacher"] == teacher)

def teacher_busy(teacher, day, period):
    t = clean(teacher)
    return any(clean(slot(sec, day, period)["teacher"]) == t
               for sec in st.session_state.timetable)

def teacher_daily_load(teacher, day):
    t = clean(teacher)
    return sum(1 for sec in st.session_state.timetable for p in get_periods(day)
               if clean(slot(sec, day, p)["teacher"]) == t)

def subject_count_in_day(section, subject, day):
    return sum(1 for p in get_periods(day)
               if slot(section, day, p)["subject"].upper() == subject.upper())

def subject_count_total(section, subject):
    return sum(subject_count_in_day(section, subject, day) for day in DAYS)

def quota_remaining(section, subject):
    """O(1) quota check using subject_remaining dict; falls back to timetable scan."""
    key = (section, subject)
    if key in subject_remaining:
        return subject_remaining[key]
    quota = st.session_state.subject_config.get(section, {}).get(subject, 0)
    return max(0, quota - subject_count_total(section, subject))

def math_double_day(section):
    for day in DAYS:
        ps = [p for p in get_periods(day) if p != "Lunch"]
        for i in range(len(ps)-1):
            if is_math(slot(section,day,ps[i])["subject"]) and \
               is_math(slot(section,day,ps[i+1])["subject"]):
                return day
    return None

def get_last_teaching_period(day):
    return [p for p in get_periods(day) if p != "Lunch"][-1]

# ══════════════════════════════════════════════════════════════════════════════
# SUBJECT CATEGORY HELPERS
# ══════════════════════════════════════════════════════════════════════════════

DAILY_SINGLE_KEYWORDS = [
    "SINDHI","ISLAMIAT","ENGLISH","URDU","GENERAL SCIENCE",
    "ARABIC","GEOGRAPHY","HISTORY","SOCIAL STUDIES",
]
# Per spec: only Physics, Chemistry, Computer IX-X may form doubles.
# Biology is intentionally excluded — it follows the no-double rule like other subjects.
IX_X_DOUBLE_SUBJECTS = [
    "PHYSICS","CHEMISTRY","COMPUTER IX-X","COMPUTER SCIENCE IX-X",
]
MATH_KEYWORD  = "MATH"
GAMES_KEYWORD = "GAMES"
_DS_SORTED = sorted(DAILY_SINGLE_KEYWORDS, key=len, reverse=True)

def is_daily_single(subject):
    s = subject.strip().upper()
    return any(kw in s for kw in _DS_SORTED)

def is_math(subject):    return MATH_KEYWORD  in subject.upper()
def is_games(subject):   return GAMES_KEYWORD in subject.upper()
def is_ix_x_double(subject):
    s = subject.upper()
    return any(t in s for t in IX_X_DOUBLE_SUBJECTS)
def is_double_allowed(subject):
    return is_math(subject) or is_ix_x_double(subject)

# ══════════════════════════════════════════════════════════════════════════════
# CONSTRAINT ENGINE
# Constraints: C1 teacher clash  C2 daily cap≤6  C3 4-consecutive hard
#              C4 3-consecutive soft  C5 lunch=break  C6 daily-single
#              C7 no-double  C8 math daily+1double  C9 IX-X doubles
#              C10 games≠last  C11 class-teacher P1≥4/5  C12 full fill
# ══════════════════════════════════════════════════════════════════════════════

def _init_teacher(t_key):
    if t_key not in teacher_day_load:
        teacher_day_load[t_key] = {d:0 for d in DAYS}
        teacher_timeline[t_key] = {d:[0]*len(get_periods(d)) for d in DAYS}

def teacher_consecutive_streak(teacher, day, proposed_idx):
    periods  = get_periods(day)
    occupied = []
    for i, p in enumerate(periods):
        if p == "Lunch": occupied.append(-1); continue
        if i == proposed_idx: occupied.append(1); continue
        occupied.append(1 if any(slot(sec,day,p)["teacher"]==teacher
                                  for sec in st.session_state.timetable) else 0)
    max_s = s = 0
    for v in occupied:
        if v == -1: s = 0
        elif v == 1: s += 1; max_s = max(max_s, s)
        else: s = 0
    return max_s

def is_library(subject):
    """Library/free-period — never first or last teaching slot."""
    return "LIBRARY" in subject.upper()

CORE_SUBJECTS = ["MATH", "ENGLISH", "URDU", "GENERAL SCIENCE"]

def is_core(subject):
    """Core academic subjects that should not repeat on same day (except math double)."""
    return any(c in subject.upper() for c in CORE_SUBJECTS)

def can_assign(section, subject, teacher, day, period):
    """
    Single constraint gate — returns True only when ALL hard rules are satisfied.

    C1  Teacher clash: teacher cannot be in two sections simultaneously.
    C2  Daily teaching cap: ≤6 periods (≤5 on Friday).
    C3  4-consecutive hard block (Lunch = hard break).
    C4  3-consecutive soft — allowed but penalised in fitness.
    C5  Lunch is a hard break in consecutive counting.
    C6  Daily-single subjects (English/Urdu/Sindhi etc.): at most once per day,
        never consecutive.
    C7  General no-double: non-double-allowed subjects never placed adjacently.
    C7b General no-same-day-repeat: ANY subject that doesn't allow doubles must
        not appear more than once on the same day regardless of adjacency.
    C8  Math: only ONE double per week; the double must be consecutive and must
        not cross Lunch.
    C9  IX-X doubles (Physics/Chemistry/Computer): only one per subject per week;
        must be consecutive, must not cross Lunch.
    C10 Games / Library: never placed in the first or last teaching period.
    """
    periods = get_periods(day)
    idx     = periods.index(period)
    if period == "Lunch": return False

    # ── C1: teacher clash ─────────────────────────────────────────────────────
    if teacher_busy(teacher, day, period): return False
    # weekly load limit
    if count_teacher_periods(teacher) >= 25:
        return False

    # ── C2: daily cap ─────────────────────────────────────────────────────────
    t_key = clean(teacher)
    _init_teacher(t_key)
    if teacher_day_load[t_key][day] >= (5 if day == "Friday" else 6): return False

    # ── C3: hard 4-consecutive block ──────────────────────────────────────────
    if teacher_consecutive_streak(t_key, day, idx) >= 4: return False

    # ── Build real-adjacency helpers (Lunch = hard break) ────────────────────
    slots_list = [p for p in periods if p != "Lunch"]
    si         = slots_list.index(period) if period in slots_list else -1
    prev_p     = slots_list[si - 1] if si > 0 else None
    next_p     = slots_list[si + 1] if si < len(slots_list) - 1 else None
    lunch_i    = periods.index("Lunch") if "Lunch" in periods else -1
    period_i   = periods.index(period)

    def real_adj(other_p):
        """True only when the two periods are immediately adjacent with no Lunch between."""
        if other_p is None: return False
        oi  = periods.index(other_p)
        lo, hi = min(period_i, oi), max(period_i, oi)
        if lunch_i != -1 and lo < lunch_i < hi: return False
        return abs(period_i - oi) == 1

    prev_s = slot(section, day, prev_p)["subject"] if prev_p and real_adj(prev_p) else ""
    next_s = slot(section, day, next_p)["subject"] if next_p and real_adj(next_p) else ""
    su     = subject.upper()

    # ── C6: daily-single subjects ─────────────────────────────────────────────
    if is_daily_single(subject):
        if subject_count_in_day(section, subject, day) >= 1: return False
        if prev_s.upper() == su or next_s.upper() == su:     return False

    # ── C7b: general no-same-day-repeat for non-double subjects ──────────────
    # Any subject that isn't allowed to double must appear at most once per day.
    if not is_double_allowed(subject):
        if subject_count_in_day(section, subject, day) >= 1: return False

    # ── C7: no consecutive identical subjects for non-double subjects ─────────
    if not is_double_allowed(subject):
        if prev_s.upper() == su or next_s.upper() == su: return False

    # ── C8: Math — at most ONE double per week, consecutive, no cross-Lunch ───
    if is_math(subject):
        # How many times Math already appears today?
        today_count = subject_count_in_day(section, subject, day)
        would_be_double = (prev_s.upper() == su or next_s.upper() == su)

        if would_be_double:
            # Doubles must not cross Lunch (real_adj already enforces this via prev_s/next_s)
            # Block if a Math double is already established on a DIFFERENT day this week
            ex = math_double_day(section)
            if ex is not None and ex != day: return False
            # Block if today already HAS a complete double (2 Math) — no triples
            if today_count >= 2: return False
        else:
            # Non-double placement: Math must not appear more than once per day
            # UNLESS this day is the designated double-day and only one Math is there yet.
            ex = math_double_day(section)
            if today_count >= 1:
                # Allow a second Math only if the day is the double-day and both
                # this slot and an adjacent empty slot can form the double.
                # Here we are NOT adjacent to existing Math → block.
                return False

    # ── C9: IX-X doubles — at most one per subject per week ──────────────────
    if is_ix_x_double(subject):
        would_be_double = (prev_s.upper() == su or next_s.upper() == su)
        if would_be_double and double_used.get((section, subject), False):
            return False
        # Prevent a non-adjacent second occurrence (which is not a double, just a repeat)
        if not would_be_double and subject_count_in_day(section, subject, day) >= 1:
            return False

    # ── C10: Games / Library never first or last teaching period ─────────────
    if is_games(subject) or is_library(subject):
        teaching = [p for p in get_periods(day) if p != "Lunch"]
        if period in (teaching[0], teaching[-1]): return False

    return True

def apply_assignment(section, subject, teacher, day, period):
    t_key = clean(teacher)
    st.session_state.timetable[section][day][period] = {"subject":subject,"teacher":t_key}
    idx = get_periods(day).index(period)
    _init_teacher(t_key)
    teacher_day_load[t_key][day]     += 1
    teacher_timeline[t_key][day][idx] = 1
    # Track 3-consecutive streaks for fitness scoring (C4 soft penalty)
    streak = teacher_consecutive_streak(t_key, day, idx)
    teacher_three_streak_count.setdefault(t_key, 0)
    if streak == 3:
        teacher_three_streak_count[t_key] += 1
    key = (section, subject)
    if key in subject_remaining:
        subject_remaining[key] = max(0, subject_remaining[key]-1)
    if is_ix_x_double(subject):
        ps = [p for p in get_periods(day) if p != "Lunch"]
        for i, p in enumerate(ps):
            if p == period:
                if (i>0 and slot(section,day,ps[i-1])["subject"].upper()==subject.upper()) or \
                   (i<len(ps)-1 and slot(section,day,ps[i+1])["subject"].upper()==subject.upper()):
                    double_used[(section,subject)] = True
                break

def undo_assignment(section, day, period):
    cell    = st.session_state.timetable[section][day][period]
    subject = cell.get("subject","")
    teacher = cell.get("teacher","")
    if not subject: return
    key = (section, subject)
    if key in subject_remaining: subject_remaining[key] += 1
    t_key = clean(teacher) if teacher else ""
    if t_key and t_key in teacher_day_load:
        teacher_day_load[t_key][day] = max(0, teacher_day_load[t_key][day]-1)
    if t_key and t_key in teacher_timeline:
        teacher_timeline[t_key][day][get_periods(day).index(period)] = 0
    st.session_state.timetable[section][day][period] = {"subject":"","teacher":""}
    if is_ix_x_double(subject):
        still = any(
            slot(section,d,ps[i])["subject"].upper()==subject.upper() and
            slot(section,d,ps[i+1])["subject"].upper()==subject.upper()
            for d in DAYS
            for ps in [[p for p in get_periods(d) if p!="Lunch"]]
            for i in range(len(ps)-1)
        )
        if not still: double_used.pop((section,subject), None)

# ══════════════════════════════════════════════════════════════════════════════
# SCHEDULER HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _find_teacher(section, subject):
    for t, sec_data in st.session_state.teacher_assignment.items():
        if section in sec_data and subject in sec_data[section]:
            return clean(t)
    return None

def _is_ix_x_section(section):
    return bool(re.search(r'(?<![A-Z])(?:IX|X|9|10)(?![A-Z0-9])', section.upper()))

def create_empty_timetable():
    teacher_day_load.clear(); teacher_timeline.clear()
    double_used.clear();      subject_remaining.clear()
    teacher_three_streak_count.clear()
    for t in st.session_state.teachers:
        teacher_three_streak_count[clean(t)] = 0
    for sec, sm in st.session_state.subject_config.items():
        for subj, quota in sm.items():
            subject_remaining[(sec, subj)] = quota
    for t in st.session_state.teachers:
        _init_teacher(clean(t))
    return {sec:{day:{p:{"subject":"","teacher":""}
                      for p in get_periods(day)} for day in DAYS}
            for sec in st.session_state.sections}

# ── Phase 1: class teacher P1 (≥4/5 days) ────────────────────────────────────
def assign_class_teacher_priority():
    for sec, ct in st.session_state.class_teachers.items():
        subjects = st.session_state.teacher_assignment.get(ct,{}).get(sec,[])
        if not subjects: continue
        assigned = 0
        for day in DAYS:
            if "P1" not in get_periods(day): continue
            for subj in subjects:
                if quota_remaining(sec,subj)<=0: continue
                if slot(sec,day,"P1")["subject"]=="" and can_assign(sec,subj,ct,day,"P1"):
                    apply_assignment(sec,subj,ct,day,"P1")
                    assigned += 1; break
            if assigned >= 4: break

# ── Phase 2: daily-single subjects ───────────────────────────────────────────
def assign_daily_singles():
    """
    Place subjects that must appear exactly once per day.
    Two-pass design:
      Pass 1 — one placement per day, strictly in order Mon→Fri.
               Each day is only considered if the subject hasn't been placed
               there yet, preventing any day from getting two occurrences.
      Pass 2 — if quota still remains after Pass 1 (subject has fewer than 5
               weekly periods), place remaining on days that still have none.
    This guarantees both the once-per-day rule and balanced weekly distribution.
    """
    for sec in st.session_state.subject_config:
        for subj in list(st.session_state.subject_config[sec]):
            if not is_daily_single(subj): continue
            teacher = _find_teacher(sec, subj)
            if not teacher: continue

            weekly_quota = st.session_state.subject_config[sec][subj]

            # Pass 1: try to place exactly one per day (Mon→Fri)
            # Shuffle periods within each day for variety but keep day order fixed
            days_to_try = DAYS.copy()
            random.shuffle(days_to_try)

            for day in days_to_try:
                if quota_remaining(sec, subj) <= 0: break
                if subject_count_in_day(sec, subj, day) >= 1: continue  # already placed today

                cands = [p for p in get_periods(day) if p != "Lunch"]
                random.shuffle(cands)
                for p in cands:
                    if quota_remaining(sec, subj) <= 0: break
                    if slot(sec, day, p)["subject"] != "": continue
                    if can_assign(sec, subj, teacher, day, p):
                        apply_assignment(sec, subj, teacher, day, p)
                        break

# ── Phase 3: Math (daily + one double) ───────────────────────────────────────
def assign_math():
    """
    Place Math for each section:
      Step A — place the ONE double (two consecutive periods on one random day).
               If no consecutive pair is available, falls through to singles only.
      Step B — place one Math on each remaining day (no doubles allowed after Step A).

    Invariants maintained:
      • At most 2 Math periods on the double-day (never 3+).
      • All other days have exactly 1 Math.
      • Weekly total = quota (typically 6 for Mon-Fri + double day counted twice).
    """
    for sec in st.session_state.subject_config:
        for subj in st.session_state.subject_config[sec]:
            if not is_math(subj): continue
            teacher = _find_teacher(sec, subj)
            if not teacher: continue

            double_day = None  # track which day got the double this week

            # ── Step A: place the one allowed double ──────────────────────────
            if quota_remaining(sec, subj) >= 2:
                days_shuffled = DAYS.copy(); random.shuffle(days_shuffled)
                for day in days_shuffled:
                    if quota_remaining(sec, subj) < 2: break
                    ps = [p for p in get_periods(day) if p != "Lunch"]
                    for i in range(len(ps) - 1):
                        p1, p2 = ps[i], ps[i + 1]
                        if (slot(sec, day, p1)["subject"] == ""
                                and slot(sec, day, p2)["subject"] == ""
                                and can_assign(sec, subj, teacher, day, p1)
                                and can_assign(sec, subj, teacher, day, p2)):
                            apply_assignment(sec, subj, teacher, day, p1)
                            apply_assignment(sec, subj, teacher, day, p2)
                            double_day = day
                            break
                    if double_day: break

            # ── Step B: one single Math per remaining day ─────────────────────
            for day in DAYS:
                if quota_remaining(sec, subj) <= 0: break
                # Skip the double-day — it already has 2 Math periods
                if day == double_day: continue
                if subject_count_in_day(sec, subj, day) >= 1: continue

                cands = [p for p in get_periods(day) if p != "Lunch"]
                random.shuffle(cands)
                for p in cands:
                    if quota_remaining(sec, subj) <= 0: break
                    if slot(sec, day, p)["subject"] == "" and can_assign(sec, subj, teacher, day, p):
                        apply_assignment(sec, subj, teacher, day, p)
                        break

# ── Phase 4: IX-X science doubles ────────────────────────────────────────────
def assign_ix_x_doubles():
    """
    Place IX-X science subjects (Physics, Chemistry, Computer IX-X):
      Step A — place the required ONE consecutive double this week.
      Step B — place any remaining single occurrences on other days,
               ensuring one-per-day (no same-day repeats).
    """
    for sec in st.session_state.subject_config:
        if not _is_ix_x_section(sec): continue
        for subj in list(st.session_state.subject_config[sec]):
            if not is_ix_x_double(subj): continue
            teacher = _find_teacher(sec, subj)
            if not teacher: continue

            double_day = None

            # ── Step A: place exactly one double ─────────────────────────────
            if not double_used.get((sec, subj), False) and quota_remaining(sec, subj) >= 2:
                days_shuffled = DAYS.copy(); random.shuffle(days_shuffled)
                for day in days_shuffled:
                    if quota_remaining(sec, subj) < 2: break
                    ps = [p for p in get_periods(day) if p != "Lunch"]
                    for i in range(len(ps) - 1):
                        p1, p2 = ps[i], ps[i + 1]
                        if (slot(sec, day, p1)["subject"] == ""
                                and slot(sec, day, p2)["subject"] == ""
                                and can_assign(sec, subj, teacher, day, p1)
                                and can_assign(sec, subj, teacher, day, p2)):
                            apply_assignment(sec, subj, teacher, day, p1)
                            apply_assignment(sec, subj, teacher, day, p2)
                            double_day = day
                            break
                    if double_day: break

            # ── Step B: fill remaining quota as singles on other days ─────────
            for day in DAYS:
                if quota_remaining(sec, subj) <= 0: break
                if day == double_day: continue          # already has the double
                if subject_count_in_day(sec, subj, day) >= 1: continue  # already placed today

                cands = [p for p in get_periods(day) if p != "Lunch"]
                random.shuffle(cands)
                for p in cands:
                    if quota_remaining(sec, subj) <= 0: break
                    if slot(sec, day, p)["subject"] == "" and can_assign(sec, subj, teacher, day, p):
                        apply_assignment(sec, subj, teacher, day, p)
                        break

# ── Swap / displace helpers ───────────────────────────────────────────────────
def try_swap(section, subject, teacher):
    for day in DAYS:
        for period in get_periods(day):
            if period=="Lunch": continue
            curr = st.session_state.timetable[section][day][period]
            if not curr["subject"]: continue
            os2, ot = curr["subject"], curr["teacher"]
            undo_assignment(section, day, period)
            if not teacher_busy(teacher,day,period) and can_assign(section,subject,teacher,day,period):
                for d2 in DAYS:
                    for p2 in get_periods(d2):
                        if p2=="Lunch" or (d2,p2)==(day,period): continue
                        if st.session_state.timetable[section][d2][p2]["subject"]: continue
                        if not teacher_busy(ot,d2,p2) and can_assign(section,os2,ot,d2,p2):
                            apply_assignment(section,subject,teacher,day,period)
                            apply_assignment(section,os2,ot,d2,p2)
                            return True
            apply_assignment(section,os2,ot,day,period)
    return False

def try_displace(section, subject, teacher):
    """
    Free a slot by permanently evicting an occupant whose subject is already
    at or over its weekly quota (surplus period — safe to remove).
    The freed slot is then given to the under-quota `subject`.

    On failure to use the freed slot, the occupant is UNCONDITIONALLY restored
    so no slot is accidentally left empty.
    """
    for day in DAYS:
        for period in get_periods(day):
            if period == "Lunch": continue
            curr = st.session_state.timetable[section][day][period]
            os2, ot = curr.get("subject",""), curr.get("teacher","")
            if not os2 or quota_remaining(section, os2) > 0: continue  # occupant still needed

            undo_assignment(section, day, period)

            if (not teacher_busy(teacher, day, period)
                    and can_assign(section, subject, teacher, day, period)):
                apply_assignment(section, subject, teacher, day, period)
                return True

            # Cannot use this slot — UNCONDITIONALLY restore the evicted period
            # so the slot is never left empty (even if can_assign fails edge cases).
            apply_assignment(section, os2, ot, day, period)

    return False

# ── Phase 5: general fill ─────────────────────────────────────────────────────
def basic_auto_fill():
    """
    Fill remaining quota for all subjects.

    Key improvements:
    • Subjects processed largest-deficit-first so hard-to-place subjects win slots.
    • Day selection uses a two-tier preference:
        Tier 1 — days where the subject hasn't appeared yet (spread-first).
        Tier 2 — days where the subject appears once (only for double-allowed subjects).
      This naturally distributes subjects across the week rather than clustering.
    • Math: daily cap is strictly 2 (double-day) or 1 (all other days) — never 3+.
    • Non-double subjects: hard cap of 1 per day enforced (reinforces C7b).
    """
    secs = list(st.session_state.subject_config.keys()); random.shuffle(secs)

    for sec in secs:
        items = sorted(
            st.session_state.subject_config[sec].items(),
            key=lambda kv: quota_remaining(sec, kv[0]),
            reverse=True,
        )

        for subj, _ in items:
            teacher = _find_teacher(sec, subj)
            if not teacher: continue

            while quota_remaining(sec, subj) > 0:
                # ── Day preference: zero-occurrence days first ─────────────
                zero_days = [d for d in DAYS if subject_count_in_day(sec, subj, d) == 0]
                one_days  = [d for d in DAYS if subject_count_in_day(sec, subj, d) == 1
                             and is_double_allowed(subj)]
                random.shuffle(zero_days); random.shuffle(one_days)
                ordered_days = zero_days + one_days

                # Math hard cap: double-day max 2, others max 1
                def day_cap(d):
                    if is_math(subj):
                        return 2 if d == math_double_day(sec) else 1
                    if is_double_allowed(subj):
                        return 2
                    return 1    # all other subjects: strictly once per day

                valid = []
                for day in ordered_days:
                    if subject_count_in_day(sec, subj, day) >= day_cap(day): continue
                    ps = get_periods(day).copy(); random.shuffle(ps)
                    for p in ps:
                        if p == "Lunch" or slot(sec, day, p)["subject"]: continue
                        if teacher_busy(teacher, day, p): continue
                        if teacher_daily_load(teacher, day) >= (5 if day == "Friday" else 6): continue
                        if can_assign(sec, subj, teacher, day, p):
                            valid.append((day, p))

                if not valid:
                    if try_swap(sec, subj, teacher): continue
                    if try_displace(sec, subj, teacher): continue
                    break   # genuinely impossible this run

                day, p = random.choice(valid)
                apply_assignment(sec, subj, teacher, day, p)

# ── Phase 5b: under-quota top-up ─────────────────────────────────────────────
def fill_under_quota_subjects():
    """
    Safety-net pass after Phases 1-5: top up any subject still below weekly quota.

    Four escalation levels:
      L1  — empty slot, strict daily cap (1 per day for non-doubles; 2 for doubles).
      L2  — empty slot, relaxed cap (doubles only) — non-doubles still capped at 1/day.
      L3  — try_swap: move an occupant to another empty slot.
      L4  — try_displace: evict a surplus-quota occupant permanently.

    Math-specific: cap is always ≤2/day (never relaxed), because placing a 3rd Math
    on any day would violate the double-period-only rule.
    """
    for sec in st.session_state.subject_config:
        for subj in sorted(
            st.session_state.subject_config[sec],
            key=lambda s: quota_remaining(sec, s),
            reverse=True,
        ):
            teacher = _find_teacher(sec, subj)
            if not teacher: continue

            def strict_day_cap(d):
                """Per-day max for this subject in relaxed mode."""
                if is_math(subj):
                    return 2  # never exceed double-pair
                if is_double_allowed(subj):
                    return 2
                return 1   # non-doubles: one per day even when relaxing

            # ── Levels 1 & 2: direct placement ───────────────────────────────
            for relax in (False, True):
                if quota_remaining(sec, subj) <= 0: break
                days_shuffled = DAYS.copy(); random.shuffle(days_shuffled)
                for day in days_shuffled:
                    if quota_remaining(sec, subj) <= 0: break
                    cap = strict_day_cap(day) if relax else 1
                    if is_double_allowed(subj) and not relax: cap = 2
                    # Non-doubles never exceed 1/day regardless of relax flag
                    if not is_double_allowed(subj): cap = 1
                    if subject_count_in_day(sec, subj, day) >= cap: continue

                    ps = [p for p in get_periods(day) if p != "Lunch"]
                    random.shuffle(ps)
                    for p in ps:
                        if quota_remaining(sec, subj) <= 0: break
                        if slot(sec, day, p)["subject"]: continue
                        if can_assign(sec, subj, teacher, day, p):
                            apply_assignment(sec, subj, teacher, day, p)
                            break

            # ── Level 3: swap ─────────────────────────────────────────────────
            while quota_remaining(sec, subj) > 0:
                if not try_swap(sec, subj, teacher): break

            # ── Level 4: displace ─────────────────────────────────────────────
            while quota_remaining(sec, subj) > 0:
                if not try_displace(sec, subj, teacher): break

# ── Phase 6: emergency backfill ──────────────────────────────────────────────
def emergency_backfill():
    for sec in st.session_state.timetable:
        for day in DAYS:
            for period in get_periods(day):
                if period=="Lunch" or slot(sec,day,period)["subject"]: continue
                for subj,_ in sorted(st.session_state.subject_config.get(sec,{}).items(),
                                     key=lambda kv: quota_remaining(sec,kv[0]), reverse=True):
                    if quota_remaining(sec,subj)<=0: continue
                    teacher=_find_teacher(sec,subj)
                    if teacher and can_assign(sec,subj,teacher,day,period):
                        apply_assignment(sec,subj,teacher,day,period); break
    for sec in st.session_state.subject_config:
        for subj in sorted([s for s in st.session_state.subject_config[sec]
                            if quota_remaining(sec,s)>0],
                           key=lambda s: quota_remaining(sec,s), reverse=True):
            teacher=_find_teacher(sec,subj)
            if not teacher: continue
            stalled=False
            while quota_remaining(sec,subj)>0 and not stalled:
                placed=False
                for day in DAYS:
                    if placed: break
                    for period in get_periods(day):
                        if period=="Lunch" or slot(sec,day,period)["subject"]: continue
                        if can_assign(sec,subj,teacher,day,period):
                            apply_assignment(sec,subj,teacher,day,period); placed=True; break
                if placed: continue
                if try_swap(sec,subj,teacher): continue
                if try_displace(sec,subj,teacher): continue
                stalled=True

# ── Phase 7: Hard guarantee — class teacher gets ≥1 period in their class ────
def ensure_class_teacher_presence():
    """
    Rule 10 hard guarantee: after all phases, every class teacher must have
    at least one period teaching in their own class.

    Strategy (in order of preference):
      1. Find any empty slot the class teacher can fill with one of their subjects.
      2. If no empty slot, find a slot occupied by ANOTHER teacher teaching a
         subject the class teacher also teaches — swap the slot directly.
      3. If still not present, log (can't fix without breaking other guarantees).
    """
    for sec, ct in st.session_state.class_teachers.items():
        if sec not in st.session_state.timetable: continue
        # Already has a period? Done.
        if any(slot(sec, day, p)["teacher"] == ct
               for day in DAYS for p in get_periods(day) if p != "Lunch"):
            continue

        subjects = st.session_state.teacher_assignment.get(ct, {}).get(sec, [])
        if not subjects: continue

        placed = False

        # ── Strategy 1: empty slot ────────────────────────────────────────────
        for day in DAYS:
            if placed: break
            for p in get_periods(day):
                if p == "Lunch" or slot(sec, day, p)["subject"]: continue
                for subj in subjects:
                    if quota_remaining(sec, subj) <= 0: continue
                    if can_assign(sec, subj, ct, day, p):
                        apply_assignment(sec, subj, ct, day, p)
                        placed = True; break
                if placed: break

        if placed: continue

        # ── Strategy 2: steal a slot from another teacher ─────────────────────
        for day in DAYS:
            if placed: break
            for p in get_periods(day):
                if p == "Lunch": continue
                cell = slot(sec, day, p)
                if cell["teacher"] == ct: break          # already there
                if cell["subject"] not in subjects: continue  # CT doesn't teach this subject
                subj    = cell["subject"]
                old_t   = cell["teacher"]
                if teacher_busy(ct, day, p): continue    # CT busy elsewhere
                # Temporarily check: if we replace old_t with ct, is everything OK?
                undo_assignment(sec, day, p)
                if can_assign(sec, subj, ct, day, p):
                    apply_assignment(sec, subj, ct, day, p)
                    placed = True; break
                else:
                    # Restore original
                    if old_t and can_assign(sec, subj, old_t, day, p):
                        apply_assignment(sec, subj, old_t, day, p)
                    elif old_t:
                        apply_assignment(sec, subj, old_t, day, p)  # force-restore
            if placed: break


# ── Phase 7b: force_fill — absolute last resort for empty slots ───────────────
def force_fill():
    """
    Absolute last resort called after all normal phases.

    Goal: guarantee ZERO empty teaching slots. Accepts any legal placement first,
    then falls back to constraint-relaxed placement if necessary.

    Strategy (in order):
      Pass 1 — for each empty slot, find any subject with remaining quota that
               passes ALL hard constraints (can_assign). Uses any teacher assigned
               to that subject-section pair.
      Pass 2 — for each STILL-empty slot, relax C6/C7b (same-day repetition)
               but keep C1 (teacher clash), C2 (daily cap), C3 (4-consecutive).
               This handles edge cases where constraints have backed the scheduler
               into a corner with no legal moves.
      Pass 3 — for each STILL-empty slot, place the subject with the highest
               remaining quota regardless of same-day rules, subject only to C1/C2.
               This is the nuclear option: the result may violate soft rules, but
               the timetable will have no empty cells.

    No slot is ever left empty after this function completes (assuming at least
    one subject-teacher assignment exists for the section).
    """
    # ── Pass 1: try every subject through the full can_assign gate ────────────
    for sec in st.session_state.timetable:
        for day in DAYS:
            for period in get_periods(day):
                if period == "Lunch" or slot(sec, day, period)["subject"]: continue
                candidates = sorted(
                    [(subj, quota_remaining(sec, subj))
                     for subj in st.session_state.subject_config.get(sec, {})
                     if quota_remaining(sec, subj) > 0],
                    key=lambda x: -x[1]
                )
                for subj, _ in candidates:
                    t = _find_teacher(sec, subj)
                    if t and can_assign(sec, subj, t, day, period):
                        apply_assignment(sec, subj, t, day, period)
                        break

    # ── Pass 2: relax same-day rules, keep teacher/consecutive constraints ────
    for sec in st.session_state.timetable:
        for day in DAYS:
            for period in get_periods(day):
                if period == "Lunch" or slot(sec, day, period)["subject"]: continue
                periods_list = get_periods(day)
                idx = periods_list.index(period)
                t_key_check = None   # determined per subject below
                for subj in sorted(
                    st.session_state.subject_config.get(sec, {}),
                    key=lambda s: -quota_remaining(sec, s)
                ):
                    t = _find_teacher(sec, subj)
                    if not t: continue
                    tk = clean(t)
                    _init_teacher(tk)
                    # Enforce only C1 (clash), C2 (daily cap), C3 (4-consecutive)
                    if teacher_busy(t, day, period): continue
                    if teacher_day_load[tk][day] >= (5 if day == "Friday" else 6): continue
                    if teacher_consecutive_streak(tk, day, idx) >= 4: continue
                    if is_games(subj) or is_library(subj):
                        teaching = [p for p in get_periods(day) if p != "Lunch"]
                        if period in (teaching[0], teaching[-1]): continue
                    apply_assignment(sec, subj, t, day, period)
                    break

    # ── Pass 3: nuclear option — place by teacher availability only (C1+C2) ──
    for sec in st.session_state.timetable:
        for day in DAYS:
            for period in get_periods(day):
                if period == "Lunch" or slot(sec, day, period)["subject"]: continue
                placed = False
                for subj in st.session_state.subject_config.get(sec, {}):
                    if placed: break
                    t = _find_teacher(sec, subj)
                    if not t: continue
                    tk = clean(t)
                    _init_teacher(tk)
                    if teacher_busy(t, day, period): continue
                    if teacher_day_load[tk][day] >= (5 if day == "Friday" else 6): continue
                    apply_assignment(sec, subj, t, day, period)
                    placed = True
                if not placed:
                    # Truly no teacher available — place ANY assigned teacher
                    # regardless of load (extreme edge case with tiny teacher pool)
                    for subj in st.session_state.subject_config.get(sec, {}):
                        t = _find_teacher(sec, subj)
                        if t and not teacher_busy(t, day, period):
                            apply_assignment(sec, subj, t, day, period)
                            break


def calculate_fitness():
    score = 10_000
    for sec,cfg in st.session_state.subject_config.items():
        for subj in cfg: score -= quota_remaining(sec,subj)*200
    score -= sum(teacher_three_streak_count.values())*30  # C4 soft penalty
    score -= len(validate_no_three_consecutive())*50
    score -= len(validate_teacher_distribution())*20
    score -= len(validate_friday_load())*10
    return score

# ══════════════════════════════════════════════════════════════════════════════
# MISC
# ══════════════════════════════════════════════════════════════════════════════

def replace_teacher_everywhere(old, new):
    for sec in st.session_state.timetable:
        for day in DAYS:
            for p in get_periods(day):
                if st.session_state.timetable[sec][day][p]["teacher"]==old:
                    st.session_state.timetable[sec][day][p]["teacher"]=new
    save_all_data()

def _pdf_style(lunch_rows=None):
    cmds = [
        ("BACKGROUND",(0,0),(-1,0), colors.HexColor(BRAND_BLUE)),
        ("TEXTCOLOR",(0,0),(-1,0),  colors.white),
        ("FONTNAME",(0,0),(-1,0),   "Helvetica-Bold"),
        ("ALIGN",(0,0),(-1,-1),     "CENTER"),
        ("VALIGN",(0,0),(-1,-1),    "MIDDLE"),
        ("FONTSIZE",(0,0),(-1,-1),  9),
        ("GRID",(0,0),(-1,-1),      0.5, colors.grey),
        ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white, colors.HexColor("#eaf0fb")]),
    ]
    if lunch_rows:
        for r in lunch_rows:
            cmds.append(("BACKGROUND",(0,r),(-1,r),colors.HexColor("#fff2cc")))
    return TableStyle(cmds)

# ══════════════════════════════════════════════════════════════════════════════
# EXPORTS
# ══════════════════════════════════════════════════════════════════════════════

def _word_landscape(doc):
    s = doc.sections[-1]
    s.orientation = WD_ORIENT.LANDSCAPE
    s.page_width, s.page_height = s.page_height, s.page_width

def export_teacher_view_word(teacher):
    doc = Document(); _word_landscape(doc)
    doc.add_heading(f"Teacher Timetable — {teacher}", 0)
    doc.add_paragraph(f"Total Weekly Periods: {count_teacher_periods(teacher)}")
    headers = ["Period","Mon-Wed Time","Mon","Tue","Wed","Thu","Fri","Fri Time"]
    table   = doc.add_table(rows=len(ALL_PERIODS)+1, cols=len(headers))
    table.style = "Table Grid"
    for i,h in enumerate(headers):
        c=table.cell(0,i); c.text=h; c.paragraphs[0].runs[0].bold=True
    for r,period in enumerate(ALL_PERIODS):
        table.cell(r+1,0).text = period
        table.cell(r+1,1).text = get_time("Monday",period)
        col=2
        for day in DAYS:
            if day=="Friday" and period not in get_periods(day):
                table.cell(r+1,col).text=""
            elif period in get_periods(day):
                val=""
                for sec in st.session_state.timetable:
                    e=slot(sec,day,period)
                    if clean(e["teacher"])==clean(teacher): val=f"{sec}\n{e['subject']}"; break
                table.cell(r+1,col).text=val
            col+=1
        table.cell(r+1,col).text=FRIDAY_TIMES.get(period,"")
    fn=f"{teacher}_timetable.docx"; doc.save(fn); return fn

def export_teacher_view_pdf(teacher):
    path=f"{teacher}_timetable.pdf"
    doc=SimpleDocTemplate(path,pagesize=landscape(A4),
                          leftMargin=20,rightMargin=20,topMargin=30,bottomMargin=30)
    styles=getSampleStyleSheet()
    elems=[Paragraph(f"<b>Teacher Timetable — {teacher}</b>",styles["Title"]),Spacer(1,8),
           Paragraph(f"<b>Total Weekly Periods:</b> {count_teacher_periods(teacher)}",styles["Normal"]),
           Spacer(1,14)]
    data=[["Period","Mon-Wed Time","Mon","Tue","Wed","Thu","Fri","Fri Time"]]
    lunch_rows=[]
    for ri,period in enumerate(ALL_PERIODS,1):
        if period=="Lunch": lunch_rows.append(ri)
        row=[period,get_time("Monday",period)]
        for day in DAYS:
            if day=="Friday" and period not in get_periods(day): row.append(""); continue
            val=""
            if period in get_periods(day):
                for sec in st.session_state.timetable:
                    e=slot(sec,day,period)
                    if clean(e["teacher"])==clean(teacher): val=f"{sec}\n{e['subject']}"; break
            row.append(val)
        row.append(FRIDAY_TIMES.get(period,"")); data.append(row)
    pw=landscape(A4)[0]-40; cw=pw/len(data[0])
    tbl=Table(data,colWidths=[cw]*len(data[0]),repeatRows=1)
    tbl.setStyle(_pdf_style(lunch_rows)); elems.append(tbl); doc.build(elems); return path

def export_class_timetable_pdf(section):
    path=f"{section}_timetable.pdf"
    doc=SimpleDocTemplate(path,pagesize=landscape(A4),
                          leftMargin=40,rightMargin=40,topMargin=40,bottomMargin=40)
    styles=getSampleStyleSheet()
    ct=st.session_state.class_teachers.get(section,"Not Assigned")
    elems=[Paragraph(f"<b>Class Timetable — {section}</b>",styles["Title"]),Spacer(1,8),
           Paragraph(f"<b>Class Teacher:</b> {ct}",styles["Normal"]),Spacer(1,18)]
    data=[["Period","Mon-Wed","Mon","Tue","Wed","Thu Time","Thursday","Friday","Fri Time"]]
    lunch_rows=[]
    for ri,period in enumerate(ALL_PERIODS,1):
        if period=="Lunch": lunch_rows.append(ri)
        row=[period, MON_WED_TIMES.get(period,"")]
        for day in ["Monday","Tuesday","Wednesday"]:
            c=slot(section,day,period) if period in get_periods(day) else {}
            row.append(f"{c['subject']}\n({c['teacher']})" if c.get("subject") else "")
        row.append(THURSDAY_TIMES.get(period,""))
        c=slot(section,"Thursday",period) if period in get_periods("Thursday") else {}
        row.append(f"{c['subject']}\n({c['teacher']})" if c.get("subject") else "")
        c=slot(section,"Friday",period) if period in get_periods("Friday") else {}
        row.append(f"{c['subject']}\n({c['teacher']})" if c.get("subject") else "")
        row.append(FRIDAY_TIMES.get(period,"")); data.append(row)
    cw=[36,58,78,78,78,58,78,78,58]
    tbl=Table(data,colWidths=cw,repeatRows=1)
    tbl.setStyle(_pdf_style(lunch_rows)); elems.append(tbl); doc.build(elems); return path

def export_class_timetable_to_word(section):
    doc=Document(); _word_landscape(doc)
    ct=st.session_state.class_teachers.get(section,"Not Assigned")
    doc.add_heading(f"Class Timetable — {section}",0)
    doc.add_paragraph(f"Class Teacher: {ct}")
    headers=["Period","Mon-Wed","Monday","Tuesday","Wednesday","Thu Time","Thursday","Friday","Fri Time"]
    table=doc.add_table(rows=len(ALL_PERIODS)+1,cols=len(headers))
    table.style="Table Grid"
    for i,h in enumerate(headers):
        c=table.cell(0,i); c.text=h; c.paragraphs[0].runs[0].bold=True
    for r,period in enumerate(ALL_PERIODS):
        table.cell(r+1,0).text=period
        table.cell(r+1,1).text=MON_WED_TIMES.get(period,"")
        col=2
        for day in ["Monday","Tuesday","Wednesday"]:
            c=slot(section,day,period) if period in get_periods(day) else {}
            table.cell(r+1,col).text=(f"{c['subject']} ({c['teacher']})" if c.get("subject") else ""); col+=1
        table.cell(r+1,col).text=THURSDAY_TIMES.get(period,""); col+=1
        c=slot(section,"Thursday",period) if period in get_periods("Thursday") else {}
        table.cell(r+1,col).text=(f"{c['subject']} ({c['teacher']})" if c.get("subject") else ""); col+=1
        c=slot(section,"Friday",period) if period in get_periods("Friday") else {}
        table.cell(r+1,col).text=(f"{c['subject']} ({c['teacher']})" if c.get("subject") else ""); col+=1
        table.cell(r+1,col).text=FRIDAY_TIMES.get(period,"")
    fn=f"{section}_timetable.docx"; doc.save(fn); return fn

def build_principal_matrix():
    if not st.session_state.timetable: return pd.DataFrame()
    rows=[]
    for day in DAYS:
        pno=0
        for period in get_periods(day):
            if period!="Lunch": pno+=1
            row={"Day":day,"P. No.":"" if period=="Lunch" else pno,"Bell Timing":get_time(day,period)}
            for t in st.session_state.teachers:
                val=""
                for sec in st.session_state.timetable:
                    if period not in st.session_state.timetable[sec].get(day,{}): continue
                    e=slot(sec,day,period)
                    if e["teacher"]==t: val=f"{sec} {e['subject']}"; break
                row[t]=val
            rows.append(row)
    df=pd.DataFrame(rows)
    df.loc[df["Day"].duplicated(),"Day"]=""
    df["P. No."]=df["P. No."].astype(str)
    return df

def export_timetable_word(df):
    file="school_timetable.docx"; doc=Document(); _word_landscape(doc)
    for i,day in enumerate(DAYS):
        day_df=df.iloc[i*9:(i+1)*9]
        doc.add_heading(day,level=1)
        table=doc.add_table(rows=len(day_df)+1,cols=len(day_df.columns))
        table.style="Table Grid"
        for j,col in enumerate(day_df.columns): table.rows[0].cells[j].text=str(col)
        for r,(_,row) in enumerate(day_df.iterrows(),1):
            for ci,val in enumerate(row): table.rows[r].cells[ci].text=str(val)
        doc.add_paragraph("")
    doc.save(file); return file

def export_excel(df):
    file="School_Timetable.xlsx"
    with pd.ExcelWriter(file,engine="openpyxl") as writer:
        df.to_excel(writer,sheet_name="Timetable",index=False)
        ws=writer.sheets["Timetable"]
        ws.freeze_panes="A2"
        hf=PatternFill("solid",fgColor="1F4E79")
        for cell in ws[1]:
            cell.font=Font(bold=True,color="FFFFFF")
            cell.alignment=Alignment(horizontal="center",vertical="center")
            cell.fill=hf
        thin=Side(style="thin"); border=Border(left=thin,right=thin,top=thin,bottom=thin)
        alt=PatternFill("solid",fgColor="EAF0FB")
        lf =PatternFill("solid",fgColor="FFF2CC")
        for r,row in enumerate(ws.iter_rows(min_row=2),2):
            is_lunch = str(ws.cell(r,2).value)==""
            for cell in row:
                cell.border=border
                cell.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True)
                cell.fill = lf if is_lunch else (alt if r%2==0 else cell.fill)
        for col in ws.columns:
            ml=max((len(str(c.value or "")) for c in col),default=0)
            ws.column_dimensions[get_column_letter(col[0].column)].width=min(ml+4,30)
    return file

# ══════════════════════════════════════════════════════════════════════════════
# AI ANALYSIS
# ══════════════════════════════════════════════════════════════════════════════

def ai_analyze_timetable(df) -> str:
    """Send the principal matrix to GPT-4o-mini for expert timetable analysis."""
    if openai_client is None:
        return (
            "⚠️ OpenAI API key not configured. "
            "Set the OPENAI_API_KEY environment variable to enable AI analysis."
        )
    prompt = (
        "You are a school timetabling expert. Analyze the school timetable below "
        "and provide concise, actionable feedback on:\n"
        "1. Teacher workload balance (flag overloaded or underloaded teachers)\n"
        "2. Subject distribution across the week (spread vs clustering)\n"
        "3. Any scheduling constraint violations you can detect\n"
        "4. Top 3 specific improvement suggestions\n\n"
        f"Timetable:\n{df.to_string()}"
    )
    try:
        response = openai_client.chat.completions.create(
            model="gpt-4o-mini",
            messages=[{"role": "user", "content": prompt}],
            max_tokens=900,
            temperature=0.3,
        )
        return response.choices[0].message.content.strip()
    except Exception as exc:
        return f"⚠️ AI analysis failed: {exc}"


# ══════════════════════════════════════════════════════════════════════════════
# NAVIGATION
# ══════════════════════════════════════════════════════════════════════════════

is_admin = st.session_state.role == "admin"
with st.sidebar:
    st.markdown(f"**👤 Role:** {st.session_state.role.title()}")
    st.divider()
    pages = ["Dashboard","Configuration","Generate","Class View","Teacher View","Analytics"] \
            if is_admin else ["Class View","Teacher View","Analytics"]
    menu = st.selectbox("Navigation", pages)
    st.divider()
    if st.button("🚪 Logout", use_container_width=True):
        st.session_state.logged_in=False; st.session_state.role=None; st.rerun()

# ══════════════════════════════════════════════════════════════════════════════
# DASHBOARD
# ══════════════════════════════════════════════════════════════════════════════

if menu == "Dashboard":
    st.subheader("Welcome! 👋")
    tt = st.session_state.timetable
    c1,c2,c3,c4 = st.columns(4)
    c1.metric("Sections",  len(st.session_state.sections))
    c2.metric("Teachers",  len(st.session_state.teachers))
    filled = sum(1 for sec in tt for day in DAYS for p in get_periods(day)
                 if p!="Lunch" and slot(sec,day,p)["subject"]) if tt else 0
    c3.metric("Slots Filled", filled)
    clashes = len(validate_teacher_clashes()) if tt else 0
    c4.metric("Clashes", clashes)
    st.info("Configure your school in **Configuration**, then click **Generate** to build the timetable.")

# ══════════════════════════════════════════════════════════════════════════════
# CONFIGURATION
# ══════════════════════════════════════════════════════════════════════════════

if menu == "Configuration":
    t_sec, t_teach, t_ct, t_subj, t_assign = st.tabs(
        ["Sections","Teachers","Class Teachers","Subjects","Assignments"])

    with t_sec:
        col1,col2 = st.columns(2)
        with col1:
            st.subheader("➕ Add")
            ns = st.text_input("Section name", key="ns")
            if st.button("Add Section", key="add_sec"):
                if ns.strip():
                    st.session_state.sections[ns.strip()]={};  save_all_data(); st.rerun()
                else: st.warning("Enter a section name.")
        with col2:
            st.subheader("🗑️ Remove")
            if st.session_state.sections:
                rs=st.selectbox("Select",list(st.session_state.sections),key="rs")
                if st.button("Delete Section", key="del_sec"):
                    for d in (st.session_state.sections, st.session_state.subject_config,
                              st.session_state.class_teachers): d.pop(rs,None)
                    save_all_data(); st.success(f"'{rs}' removed."); st.rerun()
        st.divider()
        st.write("**Sections:**", list(st.session_state.sections))

    with t_teach:
        col1,col2 = st.columns(2)
        with col1:
            st.subheader("➕ Add")
            nt = st.text_input("Teacher name", key="nt")
            if st.button("Add Teacher", key="add_t"):
                if nt.strip():
                    st.session_state.teachers[clean(nt)]={};  save_all_data(); st.rerun()
                else: st.warning("Enter a teacher name.")
        with col2:
            st.subheader("🗑️ Remove")
            if st.session_state.teachers:
                rt=st.selectbox("Select",list(st.session_state.teachers),key="rt")
                if st.button("Delete Teacher", key="del_t"):
                    st.session_state.teachers.pop(rt,None); save_all_data()
                    st.success(f"'{rt}' removed."); st.rerun()
        st.divider()
        st.write("**Teachers:**", list(st.session_state.teachers))

    with t_ct:
        if st.session_state.sections and st.session_state.teachers:
            ss=st.selectbox("Section",list(st.session_state.sections),key="ct_s")
            st2=st.selectbox("Teacher",list(st.session_state.teachers),key="ct_t")
            if st.button("Assign Class Teacher"):
                st.session_state.class_teachers[ss]=st2; save_all_data()
                st.success(f"{st2} assigned to {ss}."); st.rerun()
        else:
            st.info("Add sections and teachers first.")
        if st.session_state.class_teachers:
            st.divider()
            ct_df=pd.DataFrame(st.session_state.class_teachers.items(),columns=["Section","Class Teacher"])
            st.dataframe(ct_df,use_container_width=True,hide_index=True)
            rm_ct=st.selectbox("Remove CT for section",list(st.session_state.class_teachers),key="rm_ct")
            if st.button("Remove CT"):
                st.session_state.class_teachers.pop(rm_ct,None); save_all_data()
                st.success("Removed."); st.rerun()

    with t_subj:
        if not st.session_state.sections:
            st.info("Add sections first.")
        else:
            sel_sec=st.selectbox("Section",list(st.session_state.sections),key="ss_sec")
            col1,col2=st.columns(2)
            with col1:
                sn=st.text_input("Subject name",key="sn")
                wc=st.number_input("Weekly periods",min_value=1,max_value=10,step=1,key="wc")
                if st.button("Add Subject"):
                    if sn.strip():
                        st.session_state.subject_config.setdefault(sel_sec,{})[sn.strip()]=int(wc)
                        save_all_data(); st.success("Added."); st.rerun()
            with col2:
                cur_subjs=st.session_state.subject_config.get(sel_sec,{})
                if cur_subjs:
                    st.dataframe(pd.DataFrame(cur_subjs.items(),
                                              columns=["Subject","Periods/Week"]),
                                 hide_index=True,use_container_width=True)
                    rm_s=st.selectbox("Remove",list(cur_subjs),key="rm_s")
                    if st.button("Delete Subject"):
                        st.session_state.subject_config[sel_sec].pop(rm_s,None)
                        save_all_data(); st.success("Removed."); st.rerun()

    with t_assign:
        if not (st.session_state.teachers and st.session_state.subject_config):
            st.info("Add teachers and subjects first.")
        else:
            col1,col2=st.columns(2)
            with col1:
                st.subheader("➕ Assign")
                at=st.selectbox("Teacher",list(st.session_state.teachers),key="at")
                as_=st.selectbox("Section",list(st.session_state.subject_config),key="as_")
                ab_list=list(st.session_state.subject_config.get(as_,{}))
                if ab_list:
                    ab=st.selectbox("Subject",ab_list,key="ab")
                    if st.button("Assign"):
                        ta=st.session_state.teacher_assignment
                        ta.setdefault(at,{}).setdefault(as_,[])
                        if ab not in ta[at][as_]: ta[at][as_].append(ab)
                        save_all_data(); st.success("Assigned."); st.rerun()
            with col2:
                st.subheader("❌ Remove")
                ta=st.session_state.teacher_assignment
                if ta:
                    rm_at=st.selectbox("Teacher",list(ta),key="rm_at")
                    if rm_at in ta:
                        rm_as=st.selectbox("Section",list(ta[rm_at]),key="rm_as") if ta[rm_at] else None
                        if rm_as and rm_as in ta.get(rm_at,{}):
                            rm_ab=st.selectbox("Subject",ta[rm_at][rm_as],key="rm_ab")
                            if st.button("Remove Assignment"):
                                ta[rm_at][rm_as].remove(rm_ab)
                                if not ta[rm_at][rm_as]: del ta[rm_at][rm_as]
                                if not ta[rm_at]: del ta[rm_at]
                                save_all_data(); st.success("Removed."); st.rerun()
            st.divider()
            rows=[{"Teacher":t,"Section":s,"Subject":sub}
                  for t,secs in st.session_state.teacher_assignment.items()
                  for s,subs in secs.items() for sub in subs]
            if rows:
                st.dataframe(pd.DataFrame(rows),hide_index=True,use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# GENERATE
# ══════════════════════════════════════════════════════════════════════════════

if menu == "Generate":
    st.subheader("⚙️ Generate Timetable")
    RUNS = st.slider("Generation attempts (more = better quality)",
                     min_value=5, max_value=30, value=15, step=5)

    if st.button("🚀 Generate Timetable", key="gen_btn"):
        best_score, best_tt = -999_999, None
        progress = st.progress(0, text="Generating…")
        for run in range(RUNS):
            st.session_state.timetable = create_empty_timetable()
            assign_class_teacher_priority()   # Phase 1: class teacher P1 ≥4/5 days
            assign_daily_singles()            # Phase 2: English/Urdu/Sindhi etc. once/day
            assign_math()                     # Phase 3: Math every day + one double/week
            assign_ix_x_doubles()             # Phase 4: Physics/Chemistry/Computer doubles
            basic_auto_fill()                 # Phase 5: general fill (spread-first)
            fill_under_quota_subjects()       # Phase 5b: quota top-up
            emergency_backfill()              # Phase 6: fill empty slots + quota enforcement
            ensure_class_teacher_presence()   # Phase 7: guarantee CT has ≥1 period
            force_fill()                      # Phase 7b: absolute last resort — no empty slots
            score = calculate_fitness()
            if score > best_score:
                best_score = score
                best_tt    = copy.deepcopy(st.session_state.timetable)
            progress.progress((run+1)/RUNS, text=f"Run {run+1}/{RUNS} — best: {best_score}")
        st.session_state.timetable = best_tt
        save_all_data(); progress.empty()
        st.success(f"✅ Done! Best fitness score: {best_score}")
        unmet = [(sec, subj, quota_remaining(sec, subj))
                 for sec, cfg in st.session_state.subject_config.items()
                 for subj in cfg if quota_remaining(sec, subj) > 0]
        if unmet:
            st.warning(f"⚠️ {len(unmet)} subject quota(s) could not be fully satisfied.")
            for sec, subj, rem in unmet[:10]:
                st.caption(f"  {sec} – {subj}: {rem} period(s) short")
        else:
            st.success("✅ All subject quotas fully met!")

    if st.session_state.timetable:
        st.divider(); st.subheader("🔄 Replace Teacher")
        c1,c2,c3 = st.columns([2,2,1])
        with c1: old_t=st.selectbox("Replace",list(st.session_state.teachers),key="rep_old")
        with c2: new_t=st.selectbox("With",   list(st.session_state.teachers),key="rep_new")
        with c3:
            st.write(""); st.write("")
            if st.button("Apply"): replace_teacher_everywhere(old_t,new_t); st.success(f"{old_t} → {new_t}")

        st.divider(); st.subheader("🔍 Validation Report")
        checks = {
            "3+ consecutive periods":   (validate_no_three_consecutive(), st.warning),
            "Unbalanced teacher loads": (validate_teacher_distribution(), st.warning),
            "Heavy Friday":             (validate_friday_load(),          st.info),
            "Maths non-consecutive":    (validate_maths_consecutive(),    st.info),
        }
        all_ok=True
        for label,(issues,fn) in checks.items():
            if issues:
                all_ok=False
                with st.expander(f"⚠️ {label} ({len(issues)})"):
                    for iss in issues: fn(iss)
        if all_ok: st.success("✅ No constraint violations found!")

# ══════════════════════════════════════════════════════════════════════════════
# CLASS VIEW
# ══════════════════════════════════════════════════════════════════════════════

if menu == "Class View":
    if not st.session_state.timetable:
        st.warning("Generate a timetable first.")
    else:
        sec = st.selectbox("Select Section", sorted(st.session_state.timetable.keys()), key="cv_sec")
        ct  = st.session_state.class_teachers.get(sec,"Not Assigned")
        c1,c2,c3 = st.columns(3)
        c1.metric("Class Teacher", ct)
        c2.metric("Subjects", len(st.session_state.subject_config.get(sec,{})))
        c3.metric("Quota Gaps", sum(1 for s in st.session_state.subject_config.get(sec,{})
                                    if quota_remaining(sec,s)>0))

        df_data = {day:[
            (f"{slot(sec,day,p)['subject']}\n({slot(sec,day,p)['teacher']})"
             if slot(sec,day,p)["subject"] else "")
            if p in get_periods(day) else ""
            for p in ALL_PERIODS] for day in DAYS}

        display_periods = ALL_PERIODS
        fri_times_map = {p: FRIDAY_TIMES.get(p, "") for p in ALL_PERIODS}

        df = pd.DataFrame(
            {
                "Period":        display_periods,
                "Mon-Wed Time":  [MON_WED_TIMES.get(p, "")  for p in display_periods],
                "Monday":        df_data["Monday"],
                "Tuesday":       df_data["Tuesday"],
                "Wednesday":     df_data["Wednesday"],
                "Thu Time":      [THURSDAY_TIMES.get(p, "") for p in display_periods],
                "Thursday":      df_data["Thursday"],
                "Friday":        df_data["Friday"],
                "Fri Time":      [fri_times_map.get(p, "")  for p in display_periods],
            }
        ).set_index("Period")
        edited_df = st.data_editor(df, use_container_width=True,
                                   disabled=not is_admin, key="class_editor")

        col1,col2 = st.columns(2)
        with col1:
            if st.button("⬇ Download PDF", key="cv_pdf"):
                path=export_class_timetable_pdf(sec)
                with open(path,"rb") as f:
                    st.download_button("Download Class PDF",f,file_name=path,mime="application/pdf")
        with col2:
            if st.button("⬇ Download Word", key="cv_word"):
                path=export_class_timetable_to_word(sec)
                with open(path,"rb") as f:
                    st.download_button("Download Class Word",f,file_name=path,
                                       mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")

        if is_admin and st.button("💾 Save Manual Changes"):
            for day in DAYS:
                for period in ALL_PERIODS:
                    if period not in get_periods(day): continue
                    # edited_df index=Period, columns include explicit day names
                    val = edited_df.loc[period, day] if day in edited_df.columns else ""
                    if val:
                        parts=str(val).split("\n"); subj=parts[0].strip()
                        teacher=(parts[1].replace("(","").replace(")","").strip().upper()
                                 if len(parts)>1 else "")
                    else:
                        subj=teacher=""
                    st.session_state.timetable[sec][day][period]={"subject":subj,"teacher":teacher}
            save_all_data(); st.success("Changes saved."); st.rerun()

        with st.expander("🔍 Validation", expanded=False):
            any_iss=False
            for iss in validate_subject_weekly(sec):    st.error(iss);   any_iss=True
            for iss in validate_teacher_clashes():      st.error(iss);   any_iss=True
            for iss in validate_class_teacher_presence():st.warning(iss);any_iss=True
            for iss in validate_teacher_max_load():     st.error(iss);   any_iss=True
            if not any_iss: st.success("No issues found.")

# ══════════════════════════════════════════════════════════════════════════════
# TEACHER VIEW
# ══════════════════════════════════════════════════════════════════════════════

if menu == "Teacher View":
    if not st.session_state.timetable:
        st.warning("Generate a timetable first.")
    else:
        st.session_state.pop("refresh_teacher_view", None)
        teacher = st.selectbox("Select Teacher", sorted(st.session_state.teachers.keys()),
                               key="tv_teacher")
        total = count_teacher_periods(teacher)
        c1,c2,c3 = st.columns(3)
        c1.metric("Total Weekly Periods", total)
        c2.metric("Avg / Day", f"{total/5:.1f}")
        busy = sum(1 for day in DAYS
                   if any(slot(sec,day,p)["teacher"]==teacher
                          for sec in st.session_state.timetable
                          for p in get_periods(day) if p!="Lunch"))
        c3.metric("Days Teaching", busy)

        df_data={}
        for day in DAYS:
            row=[]
            for p in ALL_PERIODS:
                if p=="Lunch":            row.append("Lunch"); continue
                if p not in get_periods(day): row.append(""); continue
                val=""
                for sec in st.session_state.timetable:
                    e=slot(sec,day,p)
                    if clean(e["teacher"])==clean(teacher): val=f"{sec}\n{e['subject']}"; break
                row.append(val)
            df_data[day]=row

        df=pd.DataFrame({
            "Period":        ALL_PERIODS,
            "Mon-Wed Time":  [MON_WED_TIMES.get(p,"")   for p in ALL_PERIODS],
            "Monday":        df_data["Monday"],
            "Tuesday":       df_data["Tuesday"],
            "Wednesday":     df_data["Wednesday"],
            "Thu Time":      [THURSDAY_TIMES.get(p,"")  for p in ALL_PERIODS],
            "Thursday":      df_data["Thursday"],
            "Friday":        df_data["Friday"],
            "Fri Time":      [FRIDAY_TIMES.get(p,"")    for p in ALL_PERIODS],
        }).set_index("Period")
        st.dataframe(df, use_container_width=True)

        col1,col2=st.columns(2)
        with col1:
            path=export_teacher_view_word(teacher)
            with open(path,"rb") as f:
                st.download_button("⬇ Download Word",f,file_name=path,
                                   mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
        with col2:
            path=export_teacher_view_pdf(teacher)
            with open(path,"rb") as f:
                st.download_button("⬇ Download PDF",f,file_name=path,mime="application/pdf")

# ══════════════════════════════════════════════════════════════════════════════
# ANALYTICS
# ══════════════════════════════════════════════════════════════════════════════

if menu == "Analytics":
    if not st.session_state.timetable:
        st.warning("Generate a timetable first.")
    else:
        st.subheader("👩‍🏫 Teacher Workload")
        wl={t:count_teacher_periods(t) for t in st.session_state.teachers}
        wl_df=pd.DataFrame(wl.items(),columns=["Teacher","Total Periods"]) \
               .sort_values("Total Periods",ascending=False)

        def _wl(v):
            if v>=25: return "background-color:#ff6b6b"
            if v>=20: return "background-color:#ffd166"
            return "background-color:#90ee90"

        st.dataframe(wl_df.style.applymap(_wl,subset=["Total Periods"]),
                     use_container_width=True,hide_index=True)
        st.bar_chart(wl_df.set_index("Teacher"))

        st.subheader("📊 Quota Completeness")
        qrows=[]
        for sec,cfg in st.session_state.subject_config.items():
            for subj,target in cfg.items():
                placed=target-quota_remaining(sec,subj)
                qrows.append({"Section":sec,"Subject":subj,"Target":target,
                               "Placed":placed,"Remaining":max(0,target-placed)})
        if qrows:
            def _qc(v): return "background-color:#ff6b6b" if v>0 else "background-color:#90ee90"
            st.dataframe(pd.DataFrame(qrows).style.applymap(_qc,subset=["Remaining"]),
                         use_container_width=True,hide_index=True)

        st.subheader("📋 School Master Timetable")
        df_master=build_principal_matrix()
        st.dataframe(df_master, use_container_width=True)
        # AI Analysis button
        if not df_master.empty:
            if st.button("🤖 AI Analyze Timetable"):
                with st.spinner("Analyzing with GPT-4o-mini…"):
                    result = ai_analyze_timetable(df_master)
                st.markdown("### 🤖 AI Analysis")
                st.markdown(result)

        if not df_master.empty:
            col1,col2=st.columns(2)
            with col1:
                path=export_timetable_word(df_master)
                with open(path,"rb") as f:
                    st.download_button("⬇ Download Word",f,file_name="school_timetable.docx")
            with col2:
                path=export_excel(df_master)
                with open(path,"rb") as f:
                    st.download_button("⬇ Download Excel",f,file_name="School_Timetable.xlsx")
