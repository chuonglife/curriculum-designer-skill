# Program document framework

Source structure: Appendix I, Circular 06/2022/TT-BLDTBXH (Vietnam, Ministry of Labour, Invalids and Social Affairs), the national pedagogical training program for vocational teachers. Headings are translated. The order and the numbering are kept.

`build_docs.py` produces exactly this order from `program.json`. Text in parentheses is guidance and does not appear in the output.

```
PROGRAM TITLE (bold, centered)
(optional note in italics: legal basis, version, issuing body)

I. TARGET LEARNERS
    One or two paragraphs: who should take this, entry requirements.

II. PROGRAM OBJECTIVES
    1. General objective
        2 to 3 sentences describing the graduate.
    2. Specific objectives
        - Knowledge:
            + Describe ...
        - Skills:
            + Design ...
        - Autonomy and responsibility:
            + Take initiative in ...

III. PROGRAM DURATION AND TIME UNITS
    1. Program duration: N hours.              (computed)
    2. Time units: one theory hour is 45 minutes; one practice, integrated or
       discussion hour is 60 minutes; one assessment hour is 60 minutes.
       (With "time_unit": "minutes" both lines are in minutes of learner time
       and the default sentence explains that theory is video watching time.)

IV. MODULE LIST AND TIME ALLOCATION                (table, computed)
    Module code | Module title | Total | Theory | Practice/Discussion | Assessment
    Row group A: REQUIRED MODULES with subtotal
    Row group B: ELECTIVE MODULES, only if any module has "required": false
    Last row: Total

V. DETAILED SYLLABUS
    (repeated for every module)

    MODULE M01
    MODULE TITLE IN CAPITALS
    Duration: 40 hours (Theory: 20 hours; Practice, discussion: 18 hours; Assessment: 2 hours)

    1. POSITION AND NATURE OF THE MODULE
        - Position: where the module sits in the program, what comes before it.
        - Nature: required or elective; foundation knowledge or integrated
          knowledge, skill and attitude; what competency it builds.

    2. MODULE OBJECTIVES
        On completing this module, learners are able to:
        - Knowledge: ...
        - Skills: ...
        - Autonomy and responsibility: ...

    3. MODULE CONTENT
        3.1. Content overview and time allocation     (table, computed)
            No. | Lessons in the module | Total | Theory | Practice/Discussion | Assessment
            One row per lesson, one "Module assessment" row if exam_hours > 0, a total row.
        3.2. Detailed content
            (repeated for every lesson)
            Lesson 1: Title                                        Time: 04 hours
            * Objectives:
            On completing this lesson, learners are able to:
            - Knowledge: ...
            - Skills: ...
            - Autonomy and responsibility: ...
            * Content:
            1. Main heading
            1.1. Sub-heading
            1.2. Sub-heading
            2. Main heading
            n. Practice, discussion: what learners do

    4. CONDITIONS FOR DELIVERING THE MODULE
        4.1. Learning space: ...
        4.2. Equipment: ...
        4.3. Learning materials, tools and supplies: ...
        4.4. Other resources: ...

    5. ASSESSMENT CONTENT AND METHODS
        5.1. Content
            - Knowledge: ...
            - Skills: ...
            - Autonomy and responsibility: ...
        5.2. Assessment method: ...

    6. MODULE IMPLEMENTATION GUIDELINES
        6.1. Scope of application: ...
        6.2. Teaching and learning methods
            - Instructor: ...
            - Learners: ...
        6.3. Key points to emphasize
            - ...

    7. REFERENCES
        [1] Author (year), Title, Publisher or source.

VI. PROGRAM IMPLEMENTATION GUIDELINES
    1. Admission, delivery, assessment and certification
        1.1. Admission
        1.2. Delivery          (face to face, online, blended; what a fully online
                                run requires: regulations, an LMS, infrastructure)
        1.3. Assessment and certification
        1.4. Credit transfer and recognition of prior learning (if any)
    2. Requirements for the training provider
```

## Notes on filling it in

- Module codes: `M01`, `M02`, and so on. Electives continue the numbering.
- Hours in tables print with two digits (`02`, `09`, `40`). Zero prints as an empty cell.
- The "Module assessment" row comes from `exam_hours`. Use a lesson's own `assessment` hours only when a test sits inside that lesson.
- Section 5.2 in the source standard reads: "Learning results are assessed through an end-of-module test in one of these forms: essay, oral exam, objective test, reflective report." Adapt the forms to the delivery mode: quiz, submitted assignment, product scored with a rubric, skill demonstration.
- The source standard's default rules for section VI 1.3, useful as a starting point: learners sit the module assessment after attending at least 80% of hours and completing the assigned work; scoring is on a 10-point scale; below 5.0 the module is retaken; overall grades are Excellent 9.0 to 10, Very good 8.0 to under 9.0, Good 7.0 to under 8.0, Pass 5.0 to under 7.0. Replace these with the user's institution rules when known.
- A module may have no lessons (for example a final capstone assessed by rubric). It then shows as a single line in the overview tree.
