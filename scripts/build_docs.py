#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Build a competency-based training program package from one program.json.

Outputs (written next to program.json unless an output folder is given):
  01_Training_Program.docx / .md   full program document: sections I-VI,
                                   every module with its 7 fixed sections
  00_Program_Overview.docx / .md   one-page tree: program > module > lesson > topics

Usage:
  python build_docs.py path/to/program.json [output_folder]

Time unit: hours by default. Set "time_unit": "minutes" at the top of
program.json for a short video course (micro-course): every time value is then
read as whole minutes of learner time and the tables print minutes. Any lesson
may carry "video_minutes", the length of its video, which is printed next to
the lesson time; it is the field a lesson writer reads for the video length.

The script adds hours up from lessons to modules to the program, builds every
time-allocation table itself, and refuses to write files when a declared total
does not match the computed one. It also lints learning objectives and prints
warnings (missing objective group, unmeasurable opening verb).

Requires: python-docx  (pip install -r requirements.txt, into the project's .venv when there is one)
"""
import io
import json
import os
import re
import sys

if len(sys.argv) < 2 or sys.argv[1] in ("-h", "--help"):
    # Help works even before python-docx is installed.
    print(__doc__)
    sys.exit(0 if len(sys.argv) > 1 else 2)

try:
    from docx import Document
    from docx.enum.section import WD_ORIENT
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Inches, Pt
except ImportError:
    _req = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "requirements.txt")
    sys.exit(
        "ERROR: python-docx is not installed for this Python:\n"
        f"  {sys.executable}\n"
        "If the project has a .venv, rerun this script with .venv/Scripts/python (Windows)\n"
        "or .venv/bin/python (macOS, Linux). Otherwise install it into this Python:\n"
        f'  "{sys.executable}" -m pip install -r "{_req}"'
    )

if hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

FONT = "Times New Roman"
MONO = "Consolas"
BODY_PT = 13
GROUPS = ["knowledge", "skills", "autonomy"]

# --------------------------------------------------------------------------- labels
# One entry per output language. Only "en" ships. To add a language, add an entry
# here and in MINUTE_LABELS (and WEAK_VERBS) with the same keys.
LABELS = {
    "en": {
        "knowledge": "Knowledge",
        "skills": "Skills",
        "autonomy": "Autonomy and responsibility",
        "s1": "I. TARGET LEARNERS",
        "s2": "II. PROGRAM OBJECTIVES",
        "s2_1": "1. General objective",
        "s2_2": "2. Specific objectives",
        "s3": "III. PROGRAM DURATION AND TIME UNITS",
        "s3_1": "1. Program duration: {n} hours.",
        "s3_2": "Time units: one theory hour is 45 minutes; one practice, integrated or discussion hour is 60 minutes; one assessment hour is 60 minutes.",
        "s4": "IV. MODULE LIST AND TIME ALLOCATION",
        "s5": "V. DETAILED SYLLABUS",
        "s6": "VI. PROGRAM IMPLEMENTATION GUIDELINES",
        "col_code": "Module code",
        "col_module": "Module title",
        "col_no": "No.",
        "col_lessons": "Lessons in the module",
        "col_time": "Time (hours)",
        "col_total": "Total",
        "col_theory": "Theory",
        "col_practice": "Practice/ Discussion",
        "col_assess": "Assessment",
        "required": "REQUIRED MODULES",
        "elective": "ELECTIVE MODULES",
        "grand_total": "Total",
        "grand_total_ab": "Total (A+B)",
        "sum": "Total",
        "exam_row": "Module assessment",
        "module": "MODULE",
        "duration": "Duration: ",
        "duration_rest": "{t} hours (Theory: {a} hours; Practice, discussion: {b} hours; Assessment: {c} hours)",
        "m1": "1. POSITION AND NATURE OF THE MODULE",
        "position": "Position",
        "nature": "Nature",
        "m2": "2. MODULE OBJECTIVES",
        "m2_lead": "On completing this module, learners are able to:",
        "m3": "3. MODULE CONTENT",
        "m3_1": "3.1. Content overview and time allocation",
        "m3_2": "3.2. Detailed content",
        "lesson": "Lesson",
        "lesson_time": "Time: {n} hours",
        "lesson_time_video": "Time: {n} hours, of which video: {v} minutes",
        "obj": "* Objectives:",
        "obj_lead": "On completing this lesson, learners are able to:",
        "content": "* Content:",
        "m4": "4. CONDITIONS FOR DELIVERING THE MODULE",
        "m4_1": "4.1. Learning space",
        "m4_2": "4.2. Equipment",
        "m4_3": "4.3. Learning materials, tools and supplies",
        "m4_4": "4.4. Other resources",
        "m5": "5. ASSESSMENT CONTENT AND METHODS",
        "m5_1": "5.1. Content",
        "m5_2": "5.2. Assessment method",
        "m6": "6. MODULE IMPLEMENTATION GUIDELINES",
        "m6_1": "6.1. Scope of application",
        "m6_2": "6.2. Teaching and learning methods",
        "instructor": "Instructor",
        "learners": "Learners",
        "m6_3": "6.3. Key points to emphasize",
        "m7": "7. REFERENCES",
        "overview_title": "TRAINING PROGRAM OVERVIEW",
        "tree_root": "Program: {t} ({n} hours)",
        "tree_mod": "{c} - {t} ({n} hours: Theory {a}, Practice {b}, Assessment {x})",
        "tree_lesson": "{t} ({n} hours)",
        "tree_lesson_video": "{t} ({n} hours, video {v} minutes)",
        "file_program": "01_Training_Program",
        "file_overview": "00_Program_Overview",
    },
}

# Overrides applied on top of LABELS when "time_unit" is "minutes" (micro-courses).
MINUTE_LABELS = {
    "en": {
        "s3_1": "1. Program duration: {n} minutes of learner time (about {h} hours).",
        "s3_2": "Time units: every time in this program is in minutes of learner time. Theory is time spent watching the lesson videos; practice is time spent on pause tasks, exercises and assignments; assessment is time spent on quizzes and graded tests.",
        "col_time": "Time (minutes)",
        "col_total": "Total (minutes)",
        "duration_rest": "{t} minutes (Theory: {a} minutes; Practice, discussion: {b} minutes; Assessment: {c} minutes)",
        "lesson_time": "Time: {n} minutes",
        "lesson_time_video": "Time: {n} minutes, of which video: {v} minutes",
        "tree_root": "Program: {t} ({n} minutes)",
        "tree_mod": "{c} - {t} ({n} minutes: Theory {a}, Practice {b}, Assessment {x})",
        "tree_lesson": "{t} ({n} minutes)",
        "tree_lesson_video": "{t} ({n} minutes, video {v} minutes)",
    },
}

# Keys that carry a time value, by unit. Using the other unit's key is an error,
# so a value can never be read in the wrong unit.
UNIT_KEYS = {
    "hours": {"exam": "exam_hours", "total": "total_hours"},
    "minutes": {"exam": "exam_minutes", "total": "total_minutes"},
}


def labels_for(lang, unit):
    T = dict(LABELS[lang])
    if unit == "minutes":
        T.update(MINUTE_LABELS[lang])
    T["_lang"] = lang
    T["_unit"] = unit
    return T


# Opening verbs that cannot be observed or measured. Used by the objective linter.
WEAK_VERBS = {
    "en": ["understand", "know", "learn", "be aware", "be familiar", "appreciate", "grasp", "realize", "realise", "comprehend", "master"],
}


class BuildError(Exception):
    pass


def as_list(v):
    if v is None:
        return []
    if isinstance(v, str):
        return [v]
    return list(v)


# --------------------------------------------------------------------------- hours
def hours(v, where, errors, unit="hours"):
    """Time values are whole numbers. A fraction would be silently lost, so it is an error."""
    if v in (None, ""):
        return 0
    try:
        f = float(v)
    except (TypeError, ValueError):
        errors.append(f"{where}: {unit} value {v!r} is not a number")
        return 0
    if f != int(f) or f < 0:
        errors.append(f"{where}: {unit} must be whole, non-negative numbers, got {v!r}")
    return int(f)


def validate(prog):
    """Shape errors that would otherwise surface as a traceback. Returns a list of messages."""
    errs = []
    if not isinstance(prog, dict):
        return ["program.json must be a JSON object ({ ... }) at the top level"]
    mods = prog.get("modules")
    if mods is not None and not isinstance(mods, list):
        return ["'modules' must be a list ([ ... ]) of module objects"]
    so = prog.get("specific_objectives")
    if so is not None and not isinstance(so, dict):
        errs.append("Program: 'specific_objectives' must be an object with knowledge, skills, autonomy")
    for n, m in enumerate(mods or [], 1):
        if not isinstance(m, dict):
            errs.append(f"Module {n}: must be an object ({{ ... }}), got {type(m).__name__}")
            continue
        w = str(m.get("code") or f"Module {n}")
        for key in ("code", "title"):
            if not m.get(key):
                errs.append(f"{w}: module has no '{key}'")
        les_list = m.get("lessons")
        if les_list is None:
            m["lessons"] = []
        if les_list is not None and not isinstance(les_list, list):
            errs.append(f"{w}: 'lessons' must be a list ([ ... ])")
            continue
        for i, les in enumerate(les_list or [], 1):
            if not isinstance(les, dict):
                errs.append(f"{w} lesson {i}: must be an object ({{ ... }})")
            elif not les.get("title"):
                errs.append(f"{w} lesson {i}: lesson has no 'title'")
        for key in ("conditions", "assessment", "guidelines"):
            if m.get(key) is not None and not isinstance(m[key], dict):
                errs.append(f"{w}: '{key}' must be an object ({{ ... }})")
    return errs


def compute(prog, unit="hours"):
    """Attach _theory/_practice/_assess/_total to lessons, modules and program."""
    errors = []
    keys = UNIT_KEYS[unit]
    other_unit = "minutes" if unit == "hours" else "hours"

    def wrong_unit(obj, where):
        for kind, k in UNIT_KEYS[other_unit].items():
            if k in obj:
                errors.append(f"{where}: '{k}' does not fit \"time_unit\": \"{unit}\"; use '{keys[kind]}'")

    wrong_unit(prog, "Program")
    p = [0, 0, 0]
    for m in prog.get("modules", []):
        wrong_unit(m, str(m.get("code")))
        mt = [0, 0, 0]
        for i, les in enumerate(m.get("lessons", []), 1):
            w = f"{m.get('code')} lesson {i}"
            a = hours(les.get("theory"), w, errors, unit)
            b = hours(les.get("practice"), w, errors, unit)
            c = hours(les.get("assessment"), w, errors, unit)
            les["_no"] = i
            les["_theory"], les["_practice"], les["_assess"], les["_total"] = a, b, c, a + b + c
            les["_video"] = None
            if les.get("video_minutes") not in (None, ""):
                les["_video"] = hours(les["video_minutes"], f"{w} video_minutes", errors, "minutes")
            if les.get("total") not in (None, "") and hours(les["total"], f"{w} total", errors, unit) != les["_total"]:
                errors.append(f"{m.get('code')} lesson {i}: declared total {les['total']} but theory+practice+assessment = {les['_total']}")
            mt[0] += a
            mt[1] += b
            mt[2] += c
        exam = hours(m.get(keys["exam"]), f"{m.get('code')} {keys['exam']}", errors, unit)
        m["_exam"] = exam
        mt[2] += exam
        m["_theory"], m["_practice"], m["_assess"] = mt
        m["_total"] = sum(mt)
        tk = keys["total"]
        if m.get(tk) not in (None, "") and hours(m[tk], f"{m.get('code')} {tk}", errors, unit) != m["_total"]:
            errors.append(f"{m.get('code')}: declared {tk} {m[tk]} but lessons + exam add up to {m['_total']}")
        for k in range(3):
            p[k] += mt[k]
    prog["_theory"], prog["_practice"], prog["_assess"] = p
    prog["_total"] = sum(p)
    tk = keys["total"]
    if prog.get(tk) not in (None, "") and hours(prog[tk], f"Program {tk}", errors, unit) != prog["_total"]:
        errors.append(f"Program: declared {tk} {prog[tk]} but modules add up to {prog['_total']}")

    def grp(flag):
        ms = [m for m in prog.get("modules", []) if bool(m.get("required", True)) == flag]
        return ms, sum(m["_theory"] for m in ms), sum(m["_practice"] for m in ms), sum(m["_assess"] for m in ms), sum(m["_total"] for m in ms)

    prog["_A"] = grp(True)
    prog["_B"] = grp(False)
    if errors:
        raise BuildError("\n".join(errors))


def lint(prog, lang, unit="hours"):
    """Return a list of warnings about learning objectives. Never blocks the build."""
    warns = []
    weak = WEAK_VERBS.get(lang, []) + (WEAK_VERBS["en"] if lang != "en" else [])

    def check(where, obj, need_all=True):
        obj = obj or {}
        for g in GROUPS:
            items = as_list(obj.get(g))
            if not items and need_all:
                warns.append(f"{where}: no '{g}' objective")
            for s in items:
                low = s.strip().lower()
                if low.startswith("to "):
                    low = low[3:].lstrip()
                for w in weak:
                    # "understand", "understands", then a space, comma, colon or the end
                    if re.match(re.escape(w) + "s?(?![a-z])", low):
                        warns.append(f"{where}: objective starts with unmeasurable verb '{w}': {s[:70]}")
                        break

    check("Program", prog.get("specific_objectives"))
    for m in prog.get("modules", []):
        check(m.get("code", "?"), m.get("objectives"))
        for les in m.get("lessons", []):
            w = f"{m.get('code', '?')} lesson {les['_no']}"
            check(w, les.get("objectives"))
            if unit == "minutes":
                if les["_video"] is None:
                    warns.append(f"{w}: no 'video_minutes'; the lesson writer reads the video length from it")
                elif les["_video"] > les["_theory"]:
                    warns.append(f"{w}: video_minutes {les['_video']} is more than its theory minutes {les['_theory']} (theory is video watching time)")
                if les["_theory"] and not les["_practice"]:
                    warns.append(f"{w}: no practice minutes; count the pause tasks, exercises or assignment as practice")
        limit = 240 if unit == "minutes" else 4
        if m["_total"] and m["_practice"] == 0 and m["_theory"] > limit:
            warns.append(f"{m.get('code')}: {m['_theory']} theory {unit} and no practice")
    return warns


def f2(n, blank_zero=True, pad=True):
    n = int(n or 0)
    if n == 0 and blank_zero:
        return ""
    return f"{n:02d}" if pad else str(n)


def fz(n):
    n = int(n or 0)
    return "0" if n == 0 else f"{n:02d}"


def singular(text):
    """English only: '1 hours' -> '1 hour', '1 minutes' -> '1 minute'."""
    return re.sub(r"(?<![\w.])(0?1) (hour|minute)s\b", r"\1 \2", text)


def approx_hours(minutes):
    """35 -> '0.6', 120 -> '2', 125 -> '2.1'."""
    h = round(minutes / 60, 1)
    return str(int(h)) if h == int(h) else str(h)


def lesson_time(les, T, n):
    if les.get("_video") is not None:
        return T["lesson_time_video"].format(n=n, v=les["_video"])
    return T["lesson_time"].format(n=n)


def lesson_title(les, T):
    t = les["title"].strip()
    if re.match(r"^(Lesson|Unit)\s*\d+\s*:", t, re.I):
        return t
    return f"{T['lesson']} {les['_no']}: {t}"


def overview_items(les):
    if les.get("overview"):
        return as_list(les["overview"])
    out = []
    for line in as_list(les.get("content")):
        mm = re.match(r"^(\d+)\.\s+(.*)$", line.strip())
        if mm and not re.match(r"^(practice|discussion|exercise)\b", mm.group(2).strip(), re.I):
            out.append(mm.group(2).strip())
    return out


def objective_lines(obj, T):
    lines = []
    for g in GROUPS:
        v = as_list((obj or {}).get(g))
        if not v:
            continue
        if len(v) == 1:
            lines.append(f"- {T[g]}: {v[0]}")
        else:
            lines.append(f"- {T[g]}:")
            lines.extend(f"+ {x}" for x in v)
    return lines


_f2 = f2


# --------------------------------------------------------------------------- overview tree
def build_tree(prog, T):
    short = re.sub(r"^\s*(training\s+)?program(me)?\s*:\s*", "", prog["program_title"], flags=re.I)
    lines = [T["tree_root"].format(t=short, n=prog["_total"])]
    mods = prog.get("modules", [])
    for mi, m in enumerate(mods):
        last_m = mi == len(mods) - 1
        br = "`-- " if last_m else "|-- "
        if m["_total"] == 0:
            lines.append(f"{br}{m['code']} - {m['title']}")
        else:
            lines.append(br + T["tree_mod"].format(c=m["code"], t=m["title"], n=m["_total"], a=m["_theory"], b=m["_practice"], x=m["_assess"]))
        pre1 = "    " if last_m else "|   "
        lessons = m.get("lessons", [])
        for li, les in enumerate(lessons):
            last_l = li == len(lessons) - 1
            br = "`-- " if last_l else "|-- "
            if les.get("_video") is not None:
                lines.append(pre1 + br + T["tree_lesson_video"].format(t=lesson_title(les, T), n=les["_total"], v=les["_video"]))
            else:
                lines.append(pre1 + br + T["tree_lesson"].format(t=lesson_title(les, T), n=les["_total"]))
            pre2 = pre1 + ("    " if last_l else "|   ")
            items = overview_items(les)
            for ii, it in enumerate(items):
                br = "`-- " if ii == len(items) - 1 else "|-- "
                lines.append(f"{pre2}{br}{it}")
    return [singular(x) for x in lines] if T["_lang"] == "en" else lines


# --------------------------------------------------------------------------- neutral block model
# Blocks: ("title", t) ("note", t) ("h1", t, centered) ("h2", t) ("h3", t) ("p", t)
#         ("modhead", code_line, title, bold_part, rest) ("lesson", title, time)
#         ("label", t) ("outline", [lines]) ("table", [h0, h1], rows, widths)
def build_blocks(prog, T):
    pad = T["_unit"] == "hours"

    def f2(n, blank_zero=True):
        return _f2(n, blank_zero, pad)

    B = []
    if prog.get("subtitle"):
        B.append(("pretitle", prog["subtitle"]))
    B.append(("title", prog["program_title"]))
    if prog.get("basis_note"):
        B.append(("note", prog["basis_note"]))

    B.append(("h1", T["s1"], False))
    B += [("p", x) for x in as_list(prog.get("target_learners"))]

    B.append(("h1", T["s2"], False))
    B.append(("h2", T["s2_1"]))
    B += [("p", x) for x in as_list(prog.get("general_objective"))]
    B.append(("h2", T["s2_2"]))
    so = prog.get("specific_objectives") or {}
    for g in GROUPS:
        v = as_list(so.get(g))
        if v:
            B.append(("p", f"- {T[g]}:"))
            B += [("p", f"+ {x}") for x in v]

    B.append(("h1", T["s3"], False))
    B.append(("p", T["s3_1"].format(n=prog["_total"], h=approx_hours(prog["_total"]))))
    B.append(("p", "2. " + prog.get("time_unit_note", T["s3_2"])))

    B.append(("h1", T["s4"], True))
    rows = []
    A, Bg = prog["_A"], prog["_B"]

    def mod_row(m):
        return [m["code"], m["title"], f2(m["_total"]), f2(m["_theory"]), f2(m["_practice"]), f2(m["_assess"]), False]

    if A[0]:
        rows.append(["A", T["required"], str(A[4]), f2(A[1]), f2(A[2]), f2(A[3]), True])
        rows += [mod_row(m) for m in A[0]]
    if Bg[0]:
        label = T["elective"] + (f" ({prog['elective_note']})" if prog.get("elective_note") else "")
        rows.append(["B", label, str(Bg[4]), "", "", "", True])
        rows += [mod_row(m) for m in Bg[0]]
    rows.append(["", T["grand_total_ab"] if Bg[0] else T["grand_total"], str(prog["_total"]),
                 f2(prog["_theory"]), f2(prog["_practice"]), f2(prog["_assess"]), True])
    B.append(("table", [T["col_code"], T["col_module"]], rows, [1.8, 6.0, 1.5, 1.7, 2.5, 2.5]))

    B.append(("h1", T["s5"], True))
    for m in prog.get("modules", []):
        rest = T["duration_rest"].format(t=m["_total"], a=m["_theory"], b=m["_practice"], c=m["_assess"])
        B.append(("modhead", f"{T['module']} {m['code']}", m["title"].upper(), T["duration"], rest))

        B.append(("h3", T["m1"]))
        B.append(("p", f"- {T['position']}: {m.get('position', '')}"))
        B.append(("p", f"- {T['nature']}: {m.get('nature', '')}"))

        B.append(("h3", T["m2"]))
        B.append(("p", T["m2_lead"]))
        B += [("p", x) for x in objective_lines(m.get("objectives"), T)]

        B.append(("h3", T["m3"]))
        B.append(("h2", T["m3_1"]))
        rows = []
        for les in m.get("lessons", []):
            rows.append([str(les["_no"]), lesson_title(les, T), f2(les["_total"], False), f2(les["_theory"]),
                         f2(les["_practice"]), f2(les["_assess"]), False])
        if m["_exam"]:
            rows.append([str(len(m.get("lessons", [])) + 1), T["exam_row"], f2(m["_exam"], False), "", "", f2(m["_exam"]), False])
        rows.append(["", T["sum"], f2(m["_total"]), f2(m["_theory"]), f2(m["_practice"]), f2(m["_assess"]), True])
        B.append(("table", [T["col_no"], T["col_lessons"]], rows, [1.2, 6.6, 1.5, 1.7, 2.5, 2.5]))

        if m.get("lessons"):
            B.append(("h2", T["m3_2"]))
            for les in m["lessons"]:
                B.append(("lesson", lesson_title(les, T), lesson_time(les, T, f2(les["_total"], False))))
                B.append(("label", T["obj"]))
                B.append(("p", T["obj_lead"]))
                B += [("p", x) for x in objective_lines(les.get("objectives"), T)]
                B.append(("label", T["content"]))
                B.append(("outline", as_list(les.get("content"))))

        c = m.get("conditions") or {}
        B.append(("h3", T["m4"]))
        B.append(("p", f"{T['m4_1']}: {c.get('learning_space', '')}"))
        B.append(("p", f"{T['m4_2']}: {c.get('equipment', '')}"))
        B.append(("p", f"{T['m4_3']}: {c.get('materials', '')}"))
        B.append(("p", f"{T['m4_4']}: {c.get('other_resources', '')}"))

        a = m.get("assessment") or {}
        B.append(("h3", T["m5"]))
        B.append(("p", T["m5_1"]))
        B += [("p", x) for x in objective_lines(a.get("content"), T)]
        B.append(("p", f"{T['m5_2']}: {a.get('method', '')}"))

        g = m.get("guidelines") or {}
        B.append(("h3", T["m6"]))
        B.append(("p", f"{T['m6_1']}: {g.get('scope', '')}"))
        B.append(("p", T["m6_2"]))
        B.append(("p", f"- {T['instructor']}: {g.get('instructor', '')}"))
        B.append(("p", f"- {T['learners']}: {g.get('learners', '')}"))
        B.append(("p", T["m6_3"]))
        B += [("p", f"- {x}") for x in as_list(g.get("key_points"))]

        B.append(("h3", T["m7"]))
        B += [("p", f"[{i}] {t}") for i, t in enumerate(as_list(m.get("references")), 1)]

    B.append(("h1", T["s6"], False))

    def rec(sections, level):
        for s in sections:
            if not as_list(s.get("items")) and not s.get("sections"):
                continue
            if s.get("heading"):
                B.append(("h2", s["heading"]) if level == 0 else ("p", s["heading"]))
            for x in as_list(s.get("items")):
                B.append(("p", x if x.startswith("-") else f"- {x}"))
            if s.get("sections"):
                rec(s["sections"], level + 1)

    rec(as_list(prog.get("implementation_guidelines")), 0)
    return B


# --------------------------------------------------------------------------- Markdown renderer
def bold_md(s):
    s = str(s)
    return f"**{s}**" if s else ""


def render_program_md(blocks, T):
    L = []
    def gap():
        if L and L[-1] != "":
            L.append("")

    for b in blocks:
        k = b[0]
        if k == "pretitle":
            gap()
            L.append(f"*{b[1]}*")
        elif k == "title":
            gap()
            L.append(f"# {b[1]}")
        elif k == "note":
            gap()
            L.append(f"*{b[1]}*")
        elif k == "h1":
            gap()
            L.append(f"## {b[1]}")
        elif k == "h2":
            gap()
            L.append(f"#### {b[1]}")
        elif k == "h3":
            gap()
            L.append(f"#### {b[1]}")
        elif k == "modhead":
            gap()
            L.append(f"### {b[1]}: {b[2]}")
            L.append("")
            L.append(f"**{b[3].strip()}** {b[4]}")
        elif k == "lesson":
            gap()
            L.append(f"**{b[1]}** - *{b[2]}*")
        elif k == "label":
            gap()
            L.append("\\" + b[1] if b[1].startswith("*") else b[1])
        elif k == "outline":
            gap()
            for line in b[1]:
                s = line.strip()
                indent = "&nbsp;&nbsp;&nbsp;&nbsp;" if re.match(r"^\d+\.\d", s) else ""
                s = re.sub(r"^(\d+)\.(\s)", r"\1\\.\2", s)
                L.append(f"{indent}{s}  ")
        elif k == "p":
            t = b[1]
            if t.startswith("- "):
                if L and not (L[-1].startswith("- ") or L[-1].startswith("  - ")):
                    gap()
                L.append(t)
            elif t.startswith("+ "):
                L.append("  - " + t[2:])
            else:
                gap()
                L.append(t)
        elif k == "table":
            gap()
            head = [b[1][0], b[1][1], T["col_total"], T["col_theory"], T["col_practice"], T["col_assess"]]
            L.append("| " + " | ".join(head) + " |")
            L.append("|" + "---|" * 6)
            for r in b[2]:
                cells = [bold_md(c) if r[-1] else str(c) for c in r[:-1]]
                L.append("| " + " | ".join(cells) + " |")
    return "\n".join(L).rstrip() + "\n"


def render_overview_md(prog, T):
    return f"# {T['overview_title']}\n\n## {prog['program_title']}\n\n```\n" + "\n".join(build_tree(prog, T)) + "\n```\n"


# --------------------------------------------------------------------------- DOCX renderer
def set_run_font(run, name=FONT, size=BODY_PT, bold=None, italic=None):
    run.font.name = name
    run.font.size = Pt(size)
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    for a in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        rfonts.set(qn(a), name)
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic


def para(doc, text="", align=None, bold=None, italic=None, size=BODY_PT, first=0.0,
         before=0, after=4, name=FONT, keep_next=False):
    p = doc.add_paragraph()
    if text:
        set_run_font(p.add_run(text), name, size, bold, italic)
    pf = p.paragraph_format
    pf.space_before = Pt(before)
    pf.space_after = Pt(after)
    pf.first_line_indent = Cm(first) if first else None
    pf.keep_with_next = keep_next
    p.alignment = {"center": WD_ALIGN_PARAGRAPH.CENTER, "right": WD_ALIGN_PARAGRAPH.RIGHT,
                   "justify": WD_ALIGN_PARAGRAPH.JUSTIFY}.get(align)
    return p


def cell_text(cell, text, bold=False, align="center", size=12):
    cell.text = ""
    p = cell.paragraphs[0]
    set_run_font(p.add_run(str(text)), FONT, size, bold)
    p.paragraph_format.space_after = Pt(0)
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER if align == "center" else WD_ALIGN_PARAGRAPH.LEFT


def set_col_widths(table, widths_cm):
    """Word honors widths only with a fixed layout and widths set on columns and cells."""
    table.autofit = False
    tblPr = table._tbl.tblPr
    layout = tblPr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tblPr.append(layout)
    layout.set(qn("w:type"), "fixed")
    for i, w in enumerate(widths_cm):
        table.columns[i].width = Cm(w)
    for row in table.rows:
        for i, w in enumerate(widths_cm):
            row.cells[i].width = Cm(w)


def time_table(doc, first_header, rows, widths, T):
    t = doc.add_table(rows=2, cols=6)
    t.style = "Table Grid"
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    h0, h1 = t.rows[0].cells, t.rows[1].cells
    cell_text(h0[0], first_header[0], True)
    cell_text(h0[1], first_header[1], True)
    h0[0].merge(h1[0])
    h0[1].merge(h1[1])
    cell_text(h0[2], T["col_time"], True)
    h0[2].merge(h0[5])
    for i, key in enumerate(["col_total", "col_theory", "col_practice", "col_assess"]):
        cell_text(h1[2 + i], T[key], True)
    for r in rows:
        cells = t.add_row().cells
        for i, v in enumerate(r[:-1]):
            cell_text(cells[i], v, r[-1], "left" if i == 1 else "center")
    set_col_widths(t, widths)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def setup_page(doc, landscape=False):
    s = doc.sections[0]
    if landscape:
        s.orientation = WD_ORIENT.LANDSCAPE
        s.page_width, s.page_height = Cm(29.7), Cm(21)
        for side in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
            setattr(s, side, Inches(0.5))
    else:
        s.page_width, s.page_height = Cm(21), Cm(29.7)
        s.top_margin = s.bottom_margin = Cm(2)
        s.left_margin, s.right_margin = Cm(3), Cm(2)
    st = doc.styles["Normal"]
    st.font.name = FONT
    st.font.size = Pt(BODY_PT)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), FONT)


def render_program_docx(blocks, T, path):
    doc = Document()
    setup_page(doc)
    for b in blocks:
        k = b[0]
        if k in ("title", "pretitle"):
            para(doc, b[1], align="center", bold=True, size=14, after=2)
        elif k == "note":
            para(doc, b[1], align="center", italic=True, after=10)
        elif k == "h1":
            para(doc, b[1], align="center" if b[2] else None, bold=True, before=10, keep_next=True)
        elif k == "h2":
            para(doc, b[1], bold=True, before=2, keep_next=True)
        elif k == "h3":
            para(doc, b[1], bold=True, before=8, keep_next=True)
        elif k == "modhead":
            para(doc, "")
            para(doc, b[1], align="center", bold=True, after=0, keep_next=True)
            para(doc, b[2], align="center", after=0, keep_next=True)
            p = para(doc, "", align="center")
            set_run_font(p.add_run(b[3]), bold=True)
            set_run_font(p.add_run(b[4]))
        elif k == "lesson":
            para(doc, b[1], bold=True, before=6, after=0, keep_next=True)
            para(doc, b[2], align="right", italic=True, after=2)
        elif k == "label":
            para(doc, b[1], bold=True, first=1.0, after=2)
        elif k == "outline":
            for line in b[1]:
                para(doc, line, align="justify", first=1.0)
        elif k == "p":
            para(doc, b[1], align="justify", first=1.0)
        elif k == "table":
            time_table(doc, b[1], b[2], b[3], T)
    doc.save(path)


def render_overview_docx(prog, T, path):
    doc = Document()
    setup_page(doc, landscape=True)
    para(doc, T["overview_title"], align="center", bold=True, size=16, after=2)
    para(doc, prog["program_title"], align="center", bold=True, size=14, after=8)
    for line in build_tree(prog, T):
        para(doc, line, size=8, name=MONO, after=0)
    doc.save(path)


# --------------------------------------------------------------------------- main
def main():
    if len(sys.argv) < 2 or sys.argv[1] in ("-h", "--help"):
        print(__doc__)
        sys.exit(0 if len(sys.argv) > 1 else 2)
    src = sys.argv[1]
    outdir = sys.argv[2] if len(sys.argv) > 2 else os.path.dirname(os.path.abspath(src))
    try:
        # utf-8-sig: Windows PowerShell 5.1 and Notepad save UTF-8 with a BOM
        with open(src, encoding="utf-8-sig") as f:
            prog = json.load(f)
    except FileNotFoundError:
        print(f"ERROR: file not found: {src}\nGive the path to program.json, for example programs/my-course/program.json")
        sys.exit(1)
    except json.JSONDecodeError as e:
        print(f"ERROR: {src} is not valid JSON: {e}\nFix the file at that line (a missing comma or quote is the usual cause) and rerun.")
        sys.exit(1)
    shape = validate(prog)
    if shape:
        print("PROGRAM.JSON IS NOT VALID. No files written.\n" + "\n".join(shape))
        sys.exit(1)
    os.makedirs(outdir, exist_ok=True)
    lang = prog.get("language", "en")
    if lang not in LABELS:
        print(f"Language '{lang}' is not available. Available: {', '.join(LABELS)}. Building in English ('en').")
        lang = "en"
    unit = prog.get("time_unit", "hours")
    if unit not in UNIT_KEYS:
        print(f"ERROR: \"time_unit\" must be \"hours\" or \"minutes\", got {unit!r}.")
        sys.exit(1)
    T = labels_for(lang, unit)
    for key in ("program_title", "modules"):
        if not prog.get(key):
            print(f"ERROR: program.json has no '{key}'.")
            sys.exit(1)
    try:
        compute(prog, unit)
    except BuildError as e:
        print(f"TIME VALUES ARE NOT VALID ({unit} do not add up or are not whole numbers). No files written.\n" + str(e))
        sys.exit(1)

    blocks = build_blocks(prog, T)
    if lang == "en":
        blocks = [tuple(singular(x) if isinstance(x, str) else x for x in b) for b in blocks]
    targets = [
        (T["file_program"] + ".md", lambda p: open(p, "w", encoding="utf-8").write(render_program_md(blocks, T))),
        (T["file_overview"] + ".md", lambda p: open(p, "w", encoding="utf-8").write(render_overview_md(prog, T))),
        (T["file_program"] + ".docx", lambda p: render_program_docx(blocks, T, p)),
        (T["file_overview"] + ".docx", lambda p: render_overview_docx(prog, T, p)),
    ]
    for name, fn in targets:
        p = os.path.join(outdir, name)
        fn(p)
        print("Wrote", p)

    short = "h" if unit == "hours" else "min"
    print(f"\nTotal: {prog['_total']} {unit} (theory {prog['_theory']}, practice {prog['_practice']}, assessment {prog['_assess']})")
    for m in prog["modules"]:
        vid = sum(les["_video"] or 0 for les in m.get("lessons", []))
        print(f"  {m['code']} {m['title']}: {m['_total']} {short} (theory {m['_theory']}, practice {m['_practice']}, "
              f"assessment {m['_assess']}), {len(m.get('lessons') or [])} lesson{'' if len(m.get('lessons') or []) == 1 else 's'}" + (f", video {vid} min" if vid else ""))
    warns = lint(prog, lang, unit)
    if warns:
        print(f"\n{len(warns)} warning(s):")
        for w in warns:
            print("  - " + w)
    else:
        print("\nObjective lint: clean.")


if __name__ == "__main__":
    main()
