# program.json on one page

`OBJ` below means an objective block with three keys: `knowledge`, `skills`, `autonomy` (printed as "Autonomy and responsibility"). Each key takes **one string or a list of strings**. One string prints inline. A list prints as sub-items.

All hours are **whole numbers**. The script rejects fractions. If a module needs half an hour of assessment, fold it into a lesson's practice time or round the module to full hours.

```
language                 "en" | "vi"                 heading language, default "en"
slug                     string                      folder name, informational
subtitle                 string, may be ""           line printed above the title
program_title            string   REQUIRED           printed as given, usually capitals
basis_note               string, may be ""           italic line under the title
total_hours              integer, optional           cross-check only; build fails on mismatch
target_learners          string | [string]           section I, one paragraph each
general_objective        string | [string]           section II.1
specific_objectives      OBJ                         section II.2, always printed as lists
time_unit_note           string, optional            replaces the default 45/60/60 minute sentence
elective_note            string, may be ""           e.g. "choose one of three"
modules                  [module]  REQUIRED
implementation_guidelines [section]                  section VI

module:
  code                   string   e.g. "M01"
  title                  string
  required               boolean, default true       false puts the module in group B (electives)
  total_hours            integer, optional           cross-check only
  position               string                      section 1
  nature                 string                      section 1
  objectives             OBJ                         section 2
  lessons                [lesson], may be []
  exam_hours             integer, default 0          the "Module assessment" row
  conditions             { learning_space, equipment, materials, other_resources }   strings, section 4
  assessment             { content: OBJ, method: string }                              section 5
  guidelines             { scope, instructor, learners: strings; key_points: [string] } section 6
  references             [string], may be []         section 7, numbered automatically

lesson:
  title                  string                      "Lesson n:" is added automatically
  theory                 integer hours
  practice               integer hours
  assessment             integer hours, usually 0    only for a test inside the lesson
  objectives             OBJ
  content                [string]                    outline lines: "1. ...", "1.1. ...", "2. ..."
  overview               [string], optional          topics shown in the overview tree; when absent the
                                                     tree uses the top-level content lines, skipping
                                                     lines that start with Practice or Discussion

section (implementation_guidelines):
  heading                string
  items                  [string]                    printed as dash items
  sections               [section], optional         nested one level: "1.1. Admission" and so on
```

A section with no items and no nested sections is not printed, so leave out parts that do not apply (for example credit transfer).

Keys that start with `_` are ignored. Use them for notes to yourself.
