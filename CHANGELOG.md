# Changelog

## v1.1.4 - 2026-11-19

- **Output follows the language the user writes in.** This replaces the English-default wording of the previous release: whatever language the request is in, the lesson text comes back in it. Only fixed labels that ship in English stay English.

## v1.1.3 - 2026-11-19

- **English is the default output language,** stated in SKILL.md: another language only when the user asks for it or the material is in it. The docs no longer show a Vietnamese language code as an example, so nobody reads it as a supported option.

## v1.1.2 - 2026-11-19

The skill is for an international audience. This release removes the Vietnamese output mode and the Vietnamese legal framing.

### Removed

- `build_docs.py`: the Vietnamese output mode, with its headings, file names, weak verbs in the objective lint and words in two pattern checks. `LABELS` keeps the same structure with `"en"` only, so another language can still be added. A `program.json` with any other language code builds in English and says so on the console.
- README and SKILL.md: the promise of Vietnamese headings and file names, including the FAQ entry about them.
- `templates/program.example.json` and `examples/udemy-course-packaging/`: the citations of Vietnamese regulations. `basis_note` is now empty, the module references no longer list Circular 06/2022/TT-BLDTBXH or Decree 13/2023/ND-CP, and "Other resources" in M01 names the Udemy Teaching Center guidelines.
- `references/sample-module.md`: the two Vietnamese regulations in the sample's reference list.

### Changed

- README, SKILL.md and the reference files present the document as the skill's own competency-based training (CBT) frame. They no longer say it follows a national standard. Circular 06/2022/TT-BLDTBXH stays in "Sources and credits" as the source the section order was adapted from.
- SKILL.md: never cite a national regulation, law or standard in `basis_note` unless the user names one for their own program.

### Unchanged

- English output. The Markdown of `examples/support-agent-onboarding/` is identical to v1.1.1. The Markdown of `examples/udemy-course-packaging/` and of `templates/program.example.json` differs only in the lines that cited Vietnamese regulations: the basis line under the title, M01 "Other resources", and the reference entries in M01, M02, M03, M05 and M07 (with M01's remaining reference renumbered).

## v1.1.1 - 2026-11-19

Fixes from an audit as a new user on a clean machine: fresh clone, fresh venv, every README command run as written.

### Fixed

- README: the "run the builder yourself" command pointed at `scripts/build_docs.py`, a path that exists in no folder the user is in. It now gives the full path from the project folder, in bash and PowerShell.
- README: the bash blocks used `python`, which current macOS, Debian and Ubuntu do not have. They use `python3` now, with the Git Bash, `python3-venv` and `externally-managed-environment` cases spelled out.
- README: the Claude app install named a menu path that no longer exists and left out that code execution must be on.
- `build_docs.py`: a missing `python-docx` printed a traceback. It now names the Python it ran in and the exact install command, and `--help` works before anything is installed.
- `build_docs.py`: a `program.json` saved with a UTF-8 byte order mark (Windows PowerShell 5.1, Notepad) crashed. It is read normally.
- `build_docs.py`: a module without `code` or `title`, a lesson without `title`, `null` where an object belongs, a total that is not a number, `modules` that is not a list, invalid JSON and a wrong file path all gave a Python traceback. Each now stops the build with one line that names the module and lesson, and nothing is written.
- SKILL.md: Claude ran the build with the system `python` even when the project had a `.venv`, because the command block came before the rule. The rule now comes first and the block says `<python>`.
- Objective lint: "Understands ...", "To understand ..." and "Understand: ..." were not flagged.

### Added

- README: a "Before you start" table (Claude Code, git, Python, with download links and check commands), a build-script check after install, a note for projects that are git repositories, and two FAQ entries (Vietnamese file names, Word headings).
- `references/json-schema.md`: module `code`, module `title` and lesson `title` marked REQUIRED; the lesson `total` cross-check documented; the micro-course sample labelled as a fragment.
- SKILL.md: leave `basis_note` empty unless the program has its own basis.

### Changed

- The error banner reads "TIME VALUES ARE NOT VALID" (it also covers fractions and non-numbers), the build summary says "1 lesson", and the warnings header no longer says "on learning objectives" for video and practice warnings.

### Unchanged

- The Markdown and Word text output of `templates/program.example.json` and both examples is identical to v1.1.0.

## v1.1.0 - 2026-11-17

Fixes from a trial run as a new user on Windows 11 (Git Bash and PowerShell, Python 3.12, venv).

### Added

- Micro-course mode. Set `"time_unit": "minutes"` in `program.json` and every time value is read as whole minutes of learner time. Tables print "Time (minutes)", section III states the total in minutes with the hour equivalent, and the default time-unit sentence explains that theory is video watching time. Module and program keys become `exam_minutes` and `total_minutes`; an hours key in minutes mode (or the reverse) stops the build so a number is never read in the wrong unit.
- `video_minutes` on any lesson, in either mode. It is printed next to the lesson time ("Time: 27 minutes, of which video: 7 minutes") and in the overview tree, so whoever writes the lesson from `program.json` sees the video length. In minutes mode the build warns when it is missing, larger than the lesson's theory minutes, or when a lesson has theory but no practice.
- SKILL.md: when to use minutes mode, and that a micro-course may have 2 modules and is exempt from the 2 to 16 hour lesson length.
- SKILL.md: which Python to run. Use the project's `.venv` when there is one (`.venv/Scripts/python` on Windows, `.venv/bin/python` elsewhere), otherwise `python`.
- README: PowerShell commands next to every bash command, for both the personal and the one-project install; venv setup in bash and PowerShell; how to check the skill is loaded and how to update it.
- README FAQ entry on short courses.

### Changed

- README installs the dependency with `pip install -r requirements.txt` instead of `pip install python-docx`.
- `references/json-schema.md`, `references/design-principles.md` and `references/program-framework.md` document the minutes mode.

### Unchanged

- Programs without `time_unit` build exactly as before. The Markdown and Word text output of `templates/program.example.json` and both examples is identical to v1.0.0.

## v1.0.0 - 2026-11-12

First public release.
