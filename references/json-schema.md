# program.json on one page

`OBJ` below means an objective block with three keys: `knowledge`, `skills`, `autonomy` (printed as "Autonomy and responsibility"). Each key takes **one string or a list of strings**. One string prints inline. A list prints as sub-items.

All time values are **whole numbers** in the unit set by `time_unit`: hours by default, minutes for a micro-course. The script rejects fractions. If a module needs half an hour of assessment, fold it into a lesson's practice time or round the module to full hours, or switch the program to minutes.

```
language                 "en" | "vi"                 heading language, default "en"
time_unit                "hours" | "minutes"         default "hours"; see "Micro-courses" below
slug                     string                      folder name, informational
subtitle                 string, may be ""           line printed above the title
program_title            string   REQUIRED           printed as given, usually capitals
basis_note               string, may be ""           italic line under the title
total_hours              integer, optional           cross-check only; build fails on mismatch
                                                     (total_minutes in minutes mode)
target_learners          string | [string]           section I, one paragraph each
general_objective        string | [string]           section II.1
specific_objectives      OBJ                         section II.2, always printed as lists
time_unit_note           string, optional            replaces the default 45/60/60 minute sentence
elective_note            string, may be ""           e.g. "choose one of three"
modules                  [module]  REQUIRED
implementation_guidelines [section]                  section VI

module:
  code                   string   REQUIRED           e.g. "M01"
  title                  string   REQUIRED
  required               boolean, default true       false puts the module in group B (electives)
  total_hours            integer, optional           cross-check only (total_minutes in minutes mode)
  position               string                      section 1
  nature                 string                      section 1
  objectives             OBJ                         section 2
  lessons                [lesson], may be []
  exam_hours             integer, default 0          the "Module assessment" row (exam_minutes in minutes mode)
  conditions             { learning_space, equipment, materials, other_resources }   strings, section 4
  assessment             { content: OBJ, method: string }                              section 5
  guidelines             { scope, instructor, learners: strings; key_points: [string] } section 6
  references             [string], may be []         section 7, numbered automatically

lesson:
  title                  string   REQUIRED           "Lesson n:" is added automatically, unless the
                                                     title already starts with "Lesson n:" or "Unit n:"
  total                  integer, optional           cross-check only; build fails when it differs from
                                                     theory + practice + assessment
  theory                 integer, in time_unit
  practice               integer, in time_unit
  assessment             integer, in time_unit       usually 0 in hours mode; the quiz in minutes mode
  video_minutes          integer, optional           length of the lesson's video(s) in minutes, in either
                                                     mode; printed after the lesson time; the field a
                                                     lesson writer reads for the video length
  objectives             OBJ
  content                [string]                    outline lines: "1. ...", "1.1. ...", "2. ..."
  overview               [string], optional          topics shown in the overview tree; when absent the
                                                     tree uses the top-level content lines, skipping
                                                     lines that start with Practice, Exercise or
                                                     Discussion

section (implementation_guidelines):
  heading                string
  items                  [string]                    printed as dash items
  sections               [section], optional         nested one level: "1.1. Admission" and so on
```

A missing REQUIRED key, a `modules` or `lessons` value that is not a list, or a file that is not valid JSON stops the build with a message that names the place; nothing is written. A file saved with a UTF-8 byte order mark (Windows PowerShell 5.1 and Notepad do that) is read normally.

A section with no items and no nested sections is not printed, so leave out parts that do not apply (for example credit transfer).

Keys that start with `_` are ignored. Use them for notes to yourself.

## Micro-courses: `"time_unit": "minutes"`

Use it when the course is a handful of short videos and whole hours would distort it (a 35-minute course printed as "5 hours"). Everything else in the file stays the same.

- `theory`, `practice`, `assessment` of each lesson are whole minutes of learner time. Theory is the video watching time, practice the pause tasks, exercises and assignment, assessment the quiz.
- `video_minutes` on every lesson. The build warns when it is missing or larger than the lesson's theory minutes.
- The module keys become `exam_minutes` and `total_minutes`; the program key becomes `total_minutes`. An hours key in minutes mode (or the reverse) stops the build, so a number is never read in the wrong unit.
- Tables print "Time (minutes)", lesson headings print "Time: 27 minutes, of which video: 7 minutes", and section III states the total in minutes with the hour equivalent. `time_unit_note` still overrides the time-unit sentence if you need your own wording.

A fragment that shows only the keys that change (it will not build on its own; `program_title`, module titles and the rest are left out):

```json
{
  "language": "en",
  "time_unit": "minutes",
  "total_minutes": 135,
  "modules": [
    {
      "code": "M01",
      "exam_minutes": 0,
      "lessons": [
        { "title": "Vague versus measurable objectives",
          "video_minutes": 7, "theory": 7, "practice": 15, "assessment": 5 }
      ]
    }
  ]
}
```
