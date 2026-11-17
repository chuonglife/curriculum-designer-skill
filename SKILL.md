---
name: curriculum-designer
description: Design a complete competency-based training program or course curriculum (modules, lessons, measurable learning objectives in Knowledge / Skills / Autonomy-and-responsibility, theory/practice/assessment hour allocation, assessment plan, delivery guidelines) and export it as Word (.docx) and Markdown files. Use this skill whenever the user wants to create a course, a training program, a syllabus, a curriculum, a bootcamp or workshop outline, an onboarding or corporate training plan, a module outline, learning outcomes, or says things like "build a course on X", "write the program for my Udemy course", "turn this topic into a 40-hour training", "I need a syllabus", even when they never say the word "curriculum". Also use it to restructure or audit an existing course outline against competency-based design rules.
---

# Curriculum Designer

Turn a course idea into a complete, defensible training program: who it is for, what learners can do at the end, which modules and lessons get them there, how many hours each takes, how each is assessed. The method is competency-based training (CBT): every module produces one usable competency, and every objective is something an assessor could watch a learner do.

The document structure follows a real national standard for vocational teacher training (Appendix I of Circular 06/2022/TT-BLDTBXH, Vietnam). It is used here because it is strict in the right places. It forces objectives, content, hours, delivery conditions and assessment to line up at program, module and lesson level, so gaps show immediately.

## What you produce

Everything goes in `programs/<slug>/` in the user's working folder:

| File | What it is |
|---|---|
| `program.json` | The single source of truth. Edit this, rerun the script. |
| `01_Training_Program.docx` + `.md` | Full program document: sections I to VI, each module with its 7 fixed sections |
| `00_Program_Overview.docx` + `.md` | One-page tree: program, modules, lessons, topics, with hours |

You write the JSON. `scripts/build_docs.py` does the arithmetic, the tables and the formatting, so the numbers cannot drift between the summary table and the module pages.

## Workflow

### Step 0. Pin down the brief

From the request, write down five things. When something is missing, make a reasonable assumption, state it in your final report, and keep going. Ask only when two readings would produce two very different programs (8 hours or 80 hours, beginners or experts).

1. **Program title.**
2. **Target learners**: job role, starting level, what they already can do.
3. **Total hours** of learner time. Typical ranges: workshop 4 to 16, online course 20 to 60, professional certificate 40 to 300. A micro-course of a few short videos (say 5 videos of 7 minutes) is counted in minutes instead, see Step 1.
4. **Delivery mode**: self-paced online, live online, face to face, blended.
5. **Source material**: any outline, slides, textbook, job description or market research the user already has. Read all of it before designing. If the user gave a draft outline, keep their module names and order unless a design rule is clearly broken, and say what you changed.

### Step 1. Analyze the work, then derive modules

Read `references/design-principles.md` first. It holds the 7-step development process, the mini DACUM method, the objective formula, the verb tables, the hour-allocation rules and the 20-point review checklist.

Do a small occupational analysis before naming any module. List the **duties** a graduate must be able to carry out, and under each duty the **tasks** it consists of. Then map:

- one complete duty = one **module**
- one task, or one block of knowledge a task needs = one **lesson**

Starting from the work instead of from a table of contents is what separates a competency-based program from a topic list. A topic list answers "what will be covered". A duty and task map answers "what will they be able to do", and that is the question the learner and the employer care about.

Rules of thumb:

- 4 to 8 modules, 2 to 5 lessons each. A module is a complete unit: finishing it gives the learner one competency they can use right away. A micro-course may have as few as 2 modules.
- Lessons do not overlap. A later lesson uses the product of an earlier one.
- The first module is usually orientation and core concepts. The last is usually an integrated product or a final assessment.
- In skill modules, practice hours are at least equal to theory hours, and usually double. A pure knowledge module (regulations, concepts) may be theory-heavy.
- Each module carries 0 to 2 assessment hours, entered as `exam_hours`.
- Time units, unless the user's institution uses others: one theory hour is 45 minutes, one practice or discussion hour is 60 minutes, one assessment hour is 60 minutes. For self-paced video courses, theory is video watching time, practice is time spent following along and doing assignments, and total hours means learner time, not video length.
- **Micro-courses: count in minutes.** When the user asks for a short video course (a few videos of 3 to 15 minutes, or under about 8 hours of learner time in all) or gives video lengths in minutes, set `"time_unit": "minutes"`. Every time value is then whole minutes: theory is the lesson's video length, practice the pause tasks and assignment, assessment the quiz. Put `video_minutes` on every lesson, because whoever writes the lesson from `program.json` (a person or the Lesson Content Writer skill) takes the video length from it. The 2 to 16 hour lesson length and the 4 to 8 module count do not apply. Do not fake a micro-course with one-hour lessons and a `time_unit_note`.

### Step 2. Write objectives at three levels

Program, module and lesson objectives all use the same three groups: **Knowledge**, **Skills**, **Autonomy and responsibility**.

- **Program**: a general objective of 2 to 3 sentences describing the graduate, then specific objectives as a list, one competency per line.
- **Module**: 1 to 3 statements per group.
- **Lesson**: one statement per group, as concrete as you can make it.

Every statement follows the formula **action verb + object + conditions or tools + performance standard**. Verb and object are always present. Conditions and standard are expected at lesson level. The lead-in is always "On completing this ..., learners are able to:".

Pick verbs from the tables in the principles file. Do not open an objective with "understand", "know", "learn", "be aware of", "be familiar with" or "appreciate". Nobody can observe understanding. Ask what the learner would do if they understood, and write that. The build script flags these verbs.

### Step 3. Write lesson content

Number the outline as `1.`, `1.1.`, `1.2.`, `2.` and so on, one string per line in the `content` array. The last top-level item of a skill lesson is normally a practice or discussion item that names what learners actually do. Read `references/sample-module.md` to calibrate density and tone. It is a real module from the source standard.

### Step 4. Fill the remaining module sections

Keep the order fixed (details in `references/program-framework.md`):

4. Conditions for delivery: learning space, equipment, materials, other resources
5. Assessment content and methods. The assessment content must mirror the module objectives group by group
6. Implementation guidelines: scope, what the instructor does, what learners do, key points
7. References. List only sources you are sure exist. A standard, a regulation, the platform's official documentation or the user's own materials are safer than a half-remembered book. Never invent a citation

Then section VI for the whole program: admission, delivery, assessment and certification, requirements for the provider.

### Step 5. Write `program.json` and build

Read `references/json-schema.md` for the compact list of keys, then use `templates/program.example.json` as the worked example (it is long; read one module of it, not all eight). Keep every key present and use an empty list or empty string where something does not apply. Time values are whole numbers in the program's `time_unit` (hours unless it is a micro-course). An objective group accepts either one string or a list of strings. Set `"language": "en"` (or `"vi"` for Vietnamese headings). Then run:

```bash
python <skill-folder>/scripts/build_docs.py programs/<slug>/program.json
```

**Which Python.** If the user's project folder has a `.venv`, run the script with that interpreter: `.venv/Scripts/python` on Windows, `.venv/bin/python` on macOS and Linux. Otherwise use `python` (or `python3` where `python` is missing). The script needs `python-docx`; if it is missing, install `<skill-folder>/requirements.txt` into the same interpreter with `<that python> -m pip install -r <skill-folder>/requirements.txt`.

The script sums the time values from lessons up to the program, builds all tables, and stops without writing anything if a declared total disagrees with the computed total. It then prints objective warnings. Fix the JSON and rerun until the totals add up and the warnings are either gone or deliberately accepted.

### Step 6. Review before you hand over

Go through the 20-point checklist at the end of `references/design-principles.md`. At minimum confirm:

- The module set is sufficient to reach the program objectives, and each module traces back to a duty.
- The lessons of a module are sufficient to reach that module's objectives.
- Assessment content (section 5) tests what the objectives (section 2) promise.
- Lesson objectives have performance, conditions and standard. Lesson content matches lesson objectives.
- One term for one concept throughout.

Report to the user: the four file paths, the module table with hours, the assumptions you made, and anything in the checklist you could not satisfy.

## Principles

- **The frame is fixed, the content is yours.** Do not add, drop or rename sections. A fixed frame is what makes two programs comparable and makes missing pieces visible.
- **Hours must add up.** Never type an estimated total. Enter lesson hours and let the script sum.
- **Write for the stated learners.** Examples, practice tasks and deliverables should be things that specific audience would use at work next week.
- **Plain, formal register.** Short sentences that start with a verb. No sales language. A training program is a specification, and marketing copy belongs on the course landing page.
- **No invented references.**

## Reference files

- `references/program-framework.md`: the full document frame, section by section, with notes on what goes where.
- `references/design-principles.md`: curriculum development theory condensed into working rules, verb tables and the review checklist.
- `references/sample-module.md`: one complete module from the source standard, for calibration.
- `references/overview-format.md`: format of the overview tree.
- `references/json-schema.md`: every key of `program.json` on one page.
- `templates/program.example.json`: complete data structure, a real 45-hour course.
- `scripts/build_docs.py`: builds the four files from the JSON.
