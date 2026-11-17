# Changelog

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
