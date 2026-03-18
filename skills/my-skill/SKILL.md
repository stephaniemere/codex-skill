[SKILL.md](https://github.com/user-attachments/files/26091622/SKILL.md)
---
name: xlsx-sync-update-plain-text
description: Apply user-requested edits to school sync .xlsx files (especially Classes/Class Enrolments changes by email, class name, and IDs), then output a values-only workbook with no formulas and generate validation/comparison reports. Use when a user provides spreadsheet-sync change instructions and needs an updated workbook plus report artifacts.
---

# xlsx-sync-update-plain-text

Use this workflow to process spreadsheet-sync update requests end-to-end.

## Workflow

1. Identify input workbook and output folder.
2. Parse IDs needed by the user instructions:
- Teacher IDs from `Teachers.email`
- Student IDs from `Students.email`
- Class IDs from `Classes.name`
3. Apply requested adds/removes to `Classes` and `Class Enrolments`.
4. Normalize `Class Enrolments.role` values to lowercase using `scripts/normalize_class_enrolments_role_case.py`.
5. Convert workbook to values-only (remove all formulas) using `scripts/strip_formulas_xlsx.py`.
6. Validate and compare with the bundled xlsx audit skill scripts.
7. Return paths for:
- updated workbook
- `validation_report.md`
- `comparison_report.md`

## Required Output Rules

- Keep all requested data changes.
- Ensure `Class Enrolments.role` values are lowercase (for example `student`/`teacher`).
- Ensure formula count is zero in the final workbook.
- Create both validation and comparison reports for every run.
- Report any assumptions (for example class-name variants like `Physics 11S2` vs `Physics Y11S2`).

## Commands

Normalize `Class Enrolments.role` (lowercase):

```bash
python3 scripts/normalize_class_enrolments_role_case.py \
  --input "/abs/path/input.xlsx" \
  --output "/abs/path/output_role_normalized.xlsx"
```

Values-only conversion:

```bash
python3 scripts/strip_formulas_xlsx.py \
  --input "/abs/path/input.xlsx" \
  --output "/abs/path/output_values_only.xlsx"
```

Validation:

```bash
python3 /Users/stephaniemeredith/.codex/skills/xlsx-sync-audit/scripts/xlsx_sync_audit.py validate \
  --workbook "/abs/path/output_values_only.xlsx" \
  --out "/abs/path/output_dir"
```

Comparison:

```bash
python3 /Users/stephaniemeredith/.codex/skills/xlsx-sync-audit/scripts/xlsx_sync_audit.py compare \
  --current "/abs/path/original.xlsx" \
  --new "/abs/path/output_values_only.xlsx" \
  --out "/abs/path/output_dir"
```

## Notes

- Prefer case-insensitive email matching.
- For class-name matching, normalize by lowercasing and removing non-alphanumeric characters before matching.
- If workbook opens with repair prompts, regenerate from the original source and reapply edits; do not keep patching a corrupted output as the new source.
