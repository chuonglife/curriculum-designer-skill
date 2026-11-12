# Program overview format

The overview is a one-page map of the whole program. It lets a reader check coverage and hours at a glance before reading the full document.

## Layout

- A4 landscape, 0.5 inch margins.
- Line 1: `TRAINING PROGRAM OVERVIEW`, Times New Roman 16, bold, centered.
- Line 2: the program title, Times New Roman 14, bold, centered.
- The tree: Consolas 8, one paragraph per line, indentation done with characters.

## Tree structure (plain ASCII)

```
Program: PROGRAM TITLE (45 hours)
|-- M01 - Module title (5 hours: Theory 3, Practice 2, Assessment 0)
|   |-- Lesson 1: Lesson title (2 hours)
|   |   |-- Topic 1
|   |   |-- Topic 2
|   |   `-- Last topic
|   `-- Lesson 2: Lesson title (3 hours)
|       |-- Topic
|       `-- Last topic
|-- M02 - ...
`-- M0n - Last module (...)
    `-- Lesson 1: ...
        `-- ...
```

Rules:

- A branch that is not the last of its level uses `|--`. The last branch uses `` `-- ``.
- The vertical bar `|   ` continues as long as the parent still has branches below. Once the parent is a last branch, it becomes four spaces.
- Topics under a lesson are the lesson's top-level outline items (`1.`, `2.`, ...) with the number removed. Sub-items such as `1.1.` are left out. If a lesson has an `overview` list in the JSON, that list is used instead.
- Module hours are the sum of its lessons plus `exam_hours`.
- A module with zero hours and no lessons prints as a single line without the hours.

The `.md` version is the same tree inside a code block under the two title lines.
