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

The script adds hours up from lessons to modules to the program, builds every
time-allocation table itself, and refuses to write files when a declared total
does not match the computed one. It also lints learning objectives and prints
warnings (missing objective group, unmeasurable opening verb).

Requires: python-docx  (pip install python-docx)
"""
import io
import json
import os
import re
import sys

from docx import Document
from docx.enum.section import WD_ORIENT
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Inches, Pt

if hasattr(sys.stdout, "buffer"):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

FONT = "Times New Roman"
MONO = "Consolas"
BODY_PT = 13
GROUPS = ["knowledge", "skills", "autonomy"]

# --------------------------------------------------------------------------- labels
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
        "file_program": "01_Training_Program",
        "file_overview": "00_Program_Overview",
    },
    "vi": {
        "knowledge": "Kiến thức",
        "skills": "Kỹ năng",
        "autonomy": "Năng lực tự chủ và trách nhiệm",
        "s1": "I. ĐỐI TƯỢNG BỒI DƯỠNG",
        "s2": "II. MỤC TIÊU BỒI DƯỠNG",
        "s2_1": "1. Mục tiêu chung",
        "s2_2": "2. Mục tiêu cụ thể",
        "s3": "III. THỜI GIAN BỒI DƯỠNG VÀ ĐƠN VỊ THỜI GIAN TRONG CHƯƠNG TRÌNH",
        "s3_1": "1. Thời gian bồi dưỡng: {n} giờ.",
        "s3_2": "Đơn vị thời gian của giờ học: Một giờ học lý thuyết là 45 phút; một giờ học thực hành/tích hợp/thảo luận là 60 phút; một giờ thi, kiểm tra là 60 phút.",
        "s4": "IV. DANH MỤC CÁC MÔ-ĐUN VÀ PHÂN BỔ THỜI GIAN",
        "s5": "V. CHƯƠNG TRÌNH CHI TIẾT",
        "s6": "VI. HƯỚNG DẪN THỰC HIỆN CHƯƠNG TRÌNH",
        "col_code": "Mã mô-đun",
        "col_module": "Tên mô-đun",
        "col_no": "TT",
        "col_lessons": "Tên các bài trong mô-đun",
        "col_time": "Thời gian (giờ)",
        "col_total": "Tổng số",
        "col_theory": "Lý thuyết",
        "col_practice": "Thực hành/ Thảo luận",
        "col_assess": "Thi/ kiểm tra",
        "required": "MÔ-ĐUN BẮT BUỘC",
        "elective": "MÔ-ĐUN TỰ CHỌN",
        "grand_total": "Tổng cộng",
        "grand_total_ab": "Tổng cộng (A+B)",
        "sum": "Cộng",
        "exam_row": "Thi, kiểm tra",
        "module": "MÔ-ĐUN",
        "duration": "Thời gian thực hiện: ",
        "duration_rest": "{t} giờ (Lý thuyết: {a} giờ; Thực hành, thảo luận: {b} giờ; Thi/ kiểm tra: {c} giờ)",
        "m1": "1. VỊ TRÍ, TÍNH CHẤT CỦA MÔ-ĐUN",
        "position": "Vị trí",
        "nature": "Tính chất",
        "m2": "2. MỤC TIÊU MÔ-ĐUN",
        "m2_lead": "Hoàn thành mô-đun này, người học có khả năng:",
        "m3": "3. NỘI DUNG MÔ-ĐUN",
        "m3_1": "3.1. Nội dung tổng quát và phân bổ thời gian",
        "m3_2": "3.2. Nội dung chi tiết của mô-đun",
        "lesson": "Bài",
        "lesson_time": "Thời gian: {n} giờ",
        "obj": "* Mục tiêu:",
        "obj_lead": "Hoàn thành bài học này, người học có khả năng:",
        "content": "* Nội dung:",
        "m4": "4. ĐIỀU KIỆN THỰC HIỆN MÔ-ĐUN",
        "m4_1": "4.1. Phòng học chuyên môn hóa",
        "m4_2": "4.2. Trang thiết bị máy móc",
        "m4_3": "4.3. Học liệu, dụng cụ, vật liệu",
        "m4_4": "4.4. Nguồn lực khác",
        "m5": "5. NỘI DUNG VÀ PHƯƠNG PHÁP ĐÁNH GIÁ",
        "m5_1": "5.1. Nội dung",
        "m5_2": "5.2. Phương pháp đánh giá",
        "m6": "6. HƯỚNG DẪN THỰC HIỆN MÔ-ĐUN",
        "m6_1": "6.1. Phạm vi áp dụng mô-đun",
        "m6_2": "6.2. Hướng dẫn về phương pháp giảng dạy, học tập mô-đun",
        "instructor": "Giảng viên",
        "learners": "Người học",
        "m6_3": "6.3. Những trọng tâm chương trình cần chú ý",
        "m7": "7. TÀI LIỆU THAM KHẢO",
        "overview_title": "TỔNG QUAN CHƯƠNG TRÌNH ĐÀO TẠO",
        "tree_root": "Chương trình: {t} ({n} giờ)",
        "tree_mod": "{c} - {t} ({n} giờ: LT {a}, TH {b}, Thi/KT {x})",
        "tree_lesson": "{t} ({n} giờ)",
        "file_program": "01_Chuong_trinh_dao_tao",
        "file_overview": "00_Tong_quan_chuong_trinh",
    },
}

# Opening verbs that cannot be observed or measured. Used by the objective linter.
WEAK_VERBS = {
    "en": ["understand", "know", "learn", "be aware", "be familiar", "appreciate", "grasp", "realize", "realise", "comprehend", "master"],
    "vi": ["hiểu", "biết", "nắm", "nhận thức", "làm quen"],
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
def hours(v, where, errors):
    """Hours are whole numbers. A fraction would be silently lost, so it is an error."""
    if v in (None, ""):
        return 0
    try:
        f = float(v)
    except (TypeError, ValueError):
        errors.append(f"{where}: hours value {v!r} is not a number")
        return 0
    if f != int(f) or f < 0:
        errors.append(f"{where}: hours must be whole, non-negative numbers, got {v!r}")
    return int(f)


def compute(prog):
    """Attach _theory/_practice/_assess/_total to lessons, modules and program."""
    errors = []
    p = [0, 0, 0]
    for m in prog.get("modules", []):
        mt = [0, 0, 0]
        for i, les in enumerate(m.get("lessons", []), 1):
            w = f"{m.get('code')} lesson {i}"
            a = hours(les.get("theory"), w, errors)
            b = hours(les.get("practice"), w, errors)
            c = hours(les.get("assessment"), w, errors)
            les["_no"] = i
            les["_theory"], les["_practice"], les["_assess"], les["_total"] = a, b, c, a + b + c
            if "total" in les and int(les["total"]) != les["_total"]:
                errors.append(f"{m.get('code')} lesson {i}: declared total {les['total']} but theory+practice+assessment = {les['_total']}")
            mt[0] += a
            mt[1] += b
            mt[2] += c
        exam = hours(m.get("exam_hours"), f"{m.get('code')} exam_hours", errors)
        m["_exam"] = exam
        mt[2] += exam
        m["_theory"], m["_practice"], m["_assess"] = mt
        m["_total"] = sum(mt)
        if "total_hours" in m and int(m["total_hours"]) != m["_total"]:
            errors.append(f"{m.get('code')}: declared total_hours {m['total_hours']} but lessons + exam add up to {m['_total']}")
        for k in range(3):
            p[k] += mt[k]
    prog["_theory"], prog["_practice"], prog["_assess"] = p
    prog["_total"] = sum(p)
    if "total_hours" in prog and int(prog["total_hours"]) != prog["_total"]:
        errors.append(f"Program: declared total_hours {prog['total_hours']} but modules add up to {prog['_total']}")

    def grp(flag):
        ms = [m for m in prog.get("modules", []) if bool(m.get("required", True)) == flag]
        return ms, sum(m["_theory"] for m in ms), sum(m["_practice"] for m in ms), sum(m["_assess"] for m in ms), sum(m["_total"] for m in ms)

    prog["_A"] = grp(True)
    prog["_B"] = grp(False)
    if errors:
        raise BuildError("\n".join(errors))


def lint(prog, lang):
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
                for w in weak:
                    if low.startswith(w + " ") or low.startswith(w + ","):
                        warns.append(f"{where}: objective starts with unmeasurable verb '{w}': {s[:70]}")
                        break

    check("Program", prog.get("specific_objectives"))
    for m in prog.get("modules", []):
        check(m.get("code", "?"), m.get("objectives"))
        for les in m.get("lessons", []):
            check(f"{m.get('code', '?')} lesson {les['_no']}", les.get("objectives"))
        if m["_total"] and m["_practice"] == 0 and m["_theory"] > 4:
            warns.append(f"{m.get('code')}: {m['_theory']} theory hours and no practice")
    return warns


def f2(n, blank_zero=True):
    n = int(n or 0)
    if n == 0 and blank_zero:
        return ""
    return f"{n:02d}"


def fz(n):
    n = int(n or 0)
    return "0" if n == 0 else f"{n:02d}"


def singular(text):
    """English only: '1 hours' -> '1 hour'."""
    return re.sub(r"\b(0?1) hours\b", r"\1 hour", text)


def lesson_title(les, T):
    t = les["title"].strip()
    if re.match(r"^(Lesson|Bài|Unit)\s*\d+\s*:", t, re.I):
        return t
    return f"{T['lesson']} {les['_no']}: {t}"


def overview_items(les):
    if les.get("overview"):
        return as_list(les["overview"])
    out = []
    for line in as_list(les.get("content")):
        mm = re.match(r"^(\d+)\.\s+(.*)$", line.strip())
        if mm and not re.match(r"^(practice|discussion|exercise|thực hành|thảo luận|luyện tập)\b", mm.group(2).strip(), re.I):
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
            lines.append(pre1 + br + T["tree_lesson"].format(t=lesson_title(les, T), n=les["_total"]))
            pre2 = pre1 + ("    " if last_l else "|   ")
            items = overview_items(les)
            for ii, it in enumerate(items):
                br = "`-- " if ii == len(items) - 1 else "|-- "
                lines.append(f"{pre2}{br}{it}")
    return [singular(x) for x in lines] if T is LABELS["en"] else lines


# --------------------------------------------------------------------------- neutral block model
# Blocks: ("title", t) ("note", t) ("h1", t, centered) ("h2", t) ("h3", t) ("p", t)
#         ("modhead", code_line, title, bold_part, rest) ("lesson", title, time)
#         ("label", t) ("outline", [lines]) ("table", [h0, h1], rows, widths)
def build_blocks(prog, T):
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
    so = prog.get("specific_objectives", {})
    for g in GROUPS:
        v = as_list(so.get(g))
        if v:
            B.append(("p", f"- {T[g]}:"))
            B += [("p", f"+ {x}") for x in v]

    B.append(("h1", T["s3"], False))
    B.append(("p", T["s3_1"].format(n=prog["_total"])))
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
                B.append(("lesson", lesson_title(les, T), T["lesson_time"].format(n=f2(les["_total"], False))))
                B.append(("label", T["obj"]))
                B.append(("p", T["obj_lead"]))
                B += [("p", x) for x in objective_lines(les.get("objectives"), T)]
                B.append(("label", T["content"]))
                B.append(("outline", as_list(les.get("content"))))

        c = m.get("conditions", {})
        B.append(("h3", T["m4"]))
        B.append(("p", f"{T['m4_1']}: {c.get('learning_space', '')}"))
        B.append(("p", f"{T['m4_2']}: {c.get('equipment', '')}"))
        B.append(("p", f"{T['m4_3']}: {c.get('materials', '')}"))
        B.append(("p", f"{T['m4_4']}: {c.get('other_resources', '')}"))

        a = m.get("assessment", {})
        B.append(("h3", T["m5"]))
        B.append(("p", T["m5_1"]))
        B += [("p", x) for x in objective_lines(a.get("content"), T)]
        B.append(("p", f"{T['m5_2']}: {a.get('method', '')}"))

        g = m.get("guidelines", {})
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
    os.makedirs(outdir, exist_ok=True)
    with open(src, encoding="utf-8") as f:
        prog = json.load(f)
    lang = prog.get("language", "en")
    if lang not in LABELS:
        print(f"Unknown language '{lang}'. Available: {', '.join(LABELS)}. Falling back to 'en'.")
        lang = "en"
    T = LABELS[lang]
    for key in ("program_title", "modules"):
        if not prog.get(key):
            print(f"ERROR: program.json has no '{key}'.")
            sys.exit(1)
    try:
        compute(prog)
    except BuildError as e:
        print("HOURS DO NOT ADD UP. No files written.\n" + str(e))
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

    print(f"\nTotal: {prog['_total']} hours (theory {prog['_theory']}, practice {prog['_practice']}, assessment {prog['_assess']})")
    for m in prog["modules"]:
        print(f"  {m['code']} {m['title']}: {m['_total']} h (theory {m['_theory']}, practice {m['_practice']}, "
              f"assessment {m['_assess']}), {len(m.get('lessons', []))} lessons")
    warns = lint(prog, lang)
    if warns:
        print(f"\n{len(warns)} warning(s) on learning objectives:")
        for w in warns:
            print("  - " + w)
    else:
        print("\nObjective lint: clean.")


if __name__ == "__main__":
    main()
