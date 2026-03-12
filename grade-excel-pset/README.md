# grade-excel-pset

Grade individual student Excel problem-set submissions using an automated scoring script + 4 independent Claude Code agents, then reconcile into student-facing feedback reports (DOCX).

## Quick Start

```bash
# 1. Install Python dependencies
pip install openpyxl python-docx

# 2. Run the scoring script against an answer key
python scripts/grade-pset.py --key path/to/KEY.xlsx --submissions path/to/submissions/

# 3. Invoke the skill in Claude Code
/grade-pset <assignment-dir>
```

The assignment directory must contain:
- An answer key `.xlsx` file (filename containing "KEY")
- A `submissions/` folder with student `.xlsx` files

## How It Works

**Phase 0 — Automated Scoring:** `grade-pset.py` compares each student's Excel answers against the key with 0.1% tolerance, detects common error patterns (sign errors, annualization mistakes, population vs. sample stats), and outputs `scoring-report.json`.

**Phase 1 — 4-Agent Ensemble:** Four Claude Code agents independently review each student's work:
- **Agent A** — traces the exact formula error (mechanical cause)
- **Agent B** — diagnoses the conceptual misunderstanding (why it matters)
- **Agent C** — catches completeness issues and verifies the script's work
- **Agent D** — advocates for the student, finds legitimate reasons for credit the script missed

**Phase 2 — Reconciliation:** Agents' assessments are reconciled into a single feedback report per student with score-based tone calibration.

**Phase 2.5 — DOCX Reports:** `md2docx.py` converts markdown reports to formatted Word documents for distribution.

## Files

| File | Purpose |
|------|---------|
| `grade-excel-pset.md` | The skill file — Claude Code reads this to run the workflow |
| `scripts/grade-pset.py` | Automated scoring script (compares submissions to answer key) |
| `scripts/generate-pset-report.py` | Quick-mode report generator (single-agent, no ensemble) |
| `scripts/md2docx.py` | Converts markdown feedback reports to DOCX |
| `references/feedback-report-format.md` | Defines report structure and tone bands |

## Requirements

- Python 3.11+
- `openpyxl` — reads Excel files
- `python-docx` — generates Word documents
- Claude Code with Opus model access (for the 4-agent ensemble)

## Customization

- **Answer key format:** The scoring script auto-detects grading columns by looking for "Absolute Difference" headers. Structure your key with tolerance formulas and it will find them.
- **Tone bands:** Feedback tone scales with score (90%+ encouraging, 80-89% positive, 70-79% constructive, 60-69% direct, <60% blunt). Edit `references/feedback-report-format.md` to adjust.
- **Agent roles:** Each agent's lens is defined in the skill file. Add or modify agents by editing `grade-excel-pset.md`.
