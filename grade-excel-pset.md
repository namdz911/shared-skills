---
skill: grade-excel-pset
version: 1.0
created: 2026-03-10
updated: 2026-03-12
updated_by: user
domain: teaching/global
depends_on:
  scripts: [scripts/grade-pset.py, scripts/generate-pset-report.py, scripts/md2docx.py]
  skills: []
  references:
    - references/global/feedback-report-format.md
  optional_references: []
used_by_agents: [teaching]
improvement_history:
  - 0.1 — initial design: 3-phase architecture (scoring script + 3-agent ensemble + reconciliation/feedback)
  - 0.2 — added DOCX report generation (Phase 0.5) with cell-referenced narrative error explanations
  - 0.3 — added Agent D (Student Advocate); updated reconciliation to 4-agent ensemble; fixed A01 Q3 key row offset
  - 0.4 — codified "no correct answers" rule; added Professor Rulings section; added DOCX conversion step (Phase 2.5); documented full A01 4-agent production run
  - 1.0 — promoted to production; strengthened no-correct-answers rule to cover all sections (including praise); mandatory Professor Flags; fixed tone band table (5-tier, pset-specific); rewritten orchestration to match proven architecture (single Opus agent per student, parallel waves)
---

# Grade Excel Problem Set

Grade individual student Excel problem-set submissions using an automated scoring script + 4 independent Claude Code agents, then reconcile their assessments into a student-facing feedback report.

**Usage:** `/grade-pset <assignment-dir>`
Example: `/grade-pset data/fina-363/semesters/spring-2026/submissions/assignment-01`

The assignment directory must contain:
- An answer key xlsx file (filename containing "KEY")
- A `submissions/` folder with student xlsx files

---

## Why This Skill Exists

The `grade-excel-ensemble` skill handles case-based submissions with rubric weighting, partial credit, error propagation, and modeling quality. This skill handles **problem sets** — assignments where answers are numerical values compared against an answer key with tolerance. Different grading model, different feedback needs, different skill.

The multi-agent approach adds value even for binary right/wrong scoring because:
- Agent A traces the exact formula error (mechanical cause)
- Agent B diagnoses the conceptual misunderstanding (why it matters)
- Agent C catches completeness issues and verifies the script's work
- Agent D advocates for the student — finds legitimate reasons for credit the script missed

---

## Phase 0 — Automated Scoring (Python Script)

Run `scripts/grade-pset.py` once across all students. This produces `scoring-report.json`.

### Script Behavior

1. **Auto-detect grading columns** in the key:
   - Scan each sheet for header cells containing "Absolute Difference"
   - From that column, work backwards to identify the key answer column and student answer column (the two operands of the ABS formula)
   - Scan rows for cells with the tolerance formula to identify graded items
   - Extract question text from earlier columns for each item

2. **Extract correct answers** from the key (data_only=True read)

3. **For each student submission:**
   - Open xlsx twice: data_only=True for values, data_only=False for formulas
   - Map student answer cells to key answer cells by sheet name + row
   - Compare with 0.1% relative tolerance: `ABS(student - correct) < ABS(correct) * 0.001`
   - For wrong answers, run pattern matching:

   | Pattern | Detection | Label |
   |---------|-----------|-------|
   | Forgot to annualize | Student ~ correct monthly value | `gave_monthly_not_annual` |
   | Annualization error | Student ~ correct x wrong annualization | `annualization_error` |
   | Population vs sample | Student ~ VAR.P/STDEV.P result | `used_population_stat` |
   | Missing answer | Cell is blank or zero | `missing_answer` |
   | Sign error | Student ~ -1 x correct | `sign_error` |
   | Wrong data range | Student value consistent with truncated range | `wrong_range` |
   | Hardcoded value | Formula cell contains a literal number | `hardcoded` |

4. **Output: `scoring-report.json`**

```json
{
  "assignment": "FINA363_Excel_Assignment_01",
  "key_file": "FINA363_Excel_Assignment_01_KEY.xlsx",
  "tolerance": 0.001,
  "grading_map": [
    {
      "item_id": "Q1-1",
      "sheet": "Q1",
      "row": 7,
      "question": "What is the average monthly excess market return?",
      "correct_value": 0.006501,
      "key_col": "N",
      "student_col": "O"
    }
  ],
  "qualitative_items": [
    {
      "item_id": "Q2-qual-1",
      "sheet": "Q2",
      "row": 20,
      "question": "Plot the combination of potential portfolios...",
      "max_points": 1
    }
  ],
  "students": [
    {
      "name": "Cole Barnett",
      "file": "barnettcole_..._Cole Barnett.xlsx",
      "scores": {
        "Q1": {"right": 18, "total": 21},
        "Q2": {"right": 14, "total": 16},
        "Q3": {"right": 38, "total": 42}
      },
      "total_numerical": 70,
      "total_possible": 79,
      "items": [
        {
          "item_id": "Q1-3",
          "correct": false,
          "student_value": 0.0065,
          "correct_value": 0.0809,
          "student_formula": "=AVERAGE(E5:E1168)",
          "pattern_match": "gave_monthly_not_annual",
          "pattern_detail": "Student answer matches monthly value; likely forgot to annualize"
        }
      ]
    }
  ]
}
```

### Special Handling for Regression Assignments (A02)

Assignment 02 requires students to create regression output sheets via Data Analysis. The script must:
- Check if expected output sheets exist in the student file (e.g., "BKE CAPM Output", "NNI CAPM Output")
- If sheet exists, read the regression coefficients from the expected cells
- If sheet is missing, mark all items dependent on that sheet as `missing_regression_output`
- Note: regression output cell positions may vary if the student configured Data Analysis differently; if expected cell is empty, scan nearby cells for the coefficient

---

## Phase 0.5 — Quick Mode: DOCX Performance Reports (Python Script)

> **Skip this section if running the full ensemble (Phase 1-2).** Use only when time-constrained or for low-stakes assignments.

Run `scripts/generate-pset-report.py` to generate one `.docx` per student from the scoring report.

```bash
python3 scripts/generate-pset-report.py \
    --input scoring-report.json \
    --outdir reports/

# Single student:
python3 scripts/generate-pset-report.py \
    --input scoring-report.json \
    --outdir reports/ \
    --student "Cole Barnett"
```

### Report Structure

1. **Header** — assignment name, student name, overall score (color-coded)
2. **Score Summary Table** — per-sheet breakdown (Section / Correct / Total / Score)
3. **Where You Lost Points** — errors grouped by pattern, narrative explanations
4. **Next Steps** — tone-calibrated closing guidance

### Error Explanations

- Errors are **grouped by detected pattern** (e.g., "Used Population Stat -- 2 pts", "Annualization Error -- 4 pts"), with unmatched errors under "Other Errors"
- Each error is explained in **narrative form**, referencing the student's actual cell and citing formulas inline where they help explain the issue
- **CRITICAL: No correct answers, correct formulas, correct values, or answer key data anywhere in the report.** See Professor Rulings SS4 and Important Rules SS10.
- Example: "In cell L32 on Q1, your formula `=_xlfn.STDEV.P(E905:E1168)*SQRT(12)` uses STDEV.P -- replace it with STDEV.S, which divides by (n - 1) as appropriate for sample data."

### Tone Calibration

| Score Range | Tone |
|-------------|------|
| 90-100% | Encouraging -- "Excellent work" |
| 80-89% | Positive -- "Good work overall" |
| 70-79% | Constructive -- "Solid effort, review items below" |
| 60-69% | Direct -- "Several areas to work on" |
| Below 60% | Blunt -- "Needs significant improvement" |

### Supported Error Patterns

| Pattern | Narrative Template |
|---------|-------------------|
| `missing_answer` | "The cell is blank. You need to enter a formula for..." |
| `sign_error` | "Your answer has the wrong sign. Check your formula..." |
| `gave_monthly_not_annual` | "You reported a monthly value instead of annualized..." |
| `annualization_error` | "Your formula multiplies by 12 instead of compounding..." |
| `used_population_stat` | "Your formula uses STDEV.P -- replace with STDEV.S..." |
| `wrong_range` | "The data range in your formula appears truncated..." |
| `hardcoded` | "You entered a hardcoded number instead of a formula..." |
| `wrong_item_value` | "Your answer appears to match a different item's value..." |
| (no pattern) | "Your answer is incorrect. Your formula is [formula]..." |

---

## Phase 1 — Multi-Agent Ensemble (Per Student)

For each student, the agent executes 4 roles sequentially. Each role receives:
- The student's entry from `scoring-report.json` (scores + wrong items + pattern matches)
- The student's formula dump (all formulas extracted by openpyxl)
- The answer key correct values
- The question text for each item

### Agent A — Formula Tracer

Focus: What is the mechanical cause of each wrong answer?

- For each wrong item, trace the student's formula chain
- Classify the error:
  - Wrong Excel function (e.g., STDEV.P vs STDEV.S, VAR.P vs VAR.S)
  - Wrong cell range (e.g., included header row, truncated data)
  - Wrong cell reference (e.g., referenced column C instead of column E)
  - Hardcoded value where formula expected
  - Missing formula (blank cell)
  - Circular reference
  - Arithmetic error in manual formula
- Note formula quality on correct items (hardcoded correct answer = flag)
- For regression items: check if Data Analysis was run correctly

Output: Per-item formula diagnosis with error classification.

### Agent B — Conceptual Evaluator

Focus: What concept did the student misunderstand, and what did they do well?

- For each wrong item, diagnose the conceptual gap:
  - Doesn't understand annualization
  - Confuses excess return with total return
  - Doesn't understand portfolio variance formula
  - Misapplies CAPM / factor model
  - Solver setup error (wrong objective, wrong constraints)
- Group related errors by root concept
- Grade qualitative items (1 point each):
  - Charts: labeled axes, title, correct data series, clean formatting
  - Text answers: demonstrates understanding, addresses the question, correct reasoning
- Identify 3-6 positive observations (specific, not generic)

Output: Conceptual diagnosis + qualitative scores (with justification) + positive observations.

### Agent C — Completeness & Cross-Check

Focus: Did the script miss anything? Is the submission complete?

- Verify every graded item has a non-blank answer
- Cross-check 5 randomly selected scoring results against own reading
- For A02: verify regression output sheets exist, are correctly named
- Flag anomalies:
  - Correct value but hardcoded (no formula) — script flags this too, but Agent C double-checks
  - Answers that look copied between students (unusual precision match)
  - Items where student value is very close to threshold (within 2x tolerance)
- Check formatting: are percentage answers in decimal or percentage format?
- Grade qualitative items independently from Agent B

Output: Completeness report + scoring verification + qualitative scores + flags.

### Agent D — Student Advocate

Focus: Find legitimate reasons the student deserves credit the script didn't give.

- For each wrong item, check:
  - **Alternative valid approach:** Does the student's method produce a defensible answer? (e.g., different but valid annualization convention, arithmetic vs geometric mean)
  - **Rounding tolerance edge cases:** Is the student's answer within 1-2% of correct but just outside the 0.1% threshold? If so, did intermediate rounding cause the drift?
  - **Valid interpretation:** Could the question be read differently? Does the student's answer match a reasonable alternative reading?
  - **Cascade unfairness:** If one root formula error causes 5+ downstream misses, recommend the professor consider cascade forgiveness (penalize once at root, forgive downstream)
  - **Hardcoded but close:** If a value is hardcoded but within 1% of correct, the student likely computed correctly elsewhere and typed the result — note partial understanding
  - **Near-miss understanding:** Student's approach is conceptually correct but has a minor input error (wrong cell ref by 1 row, used column C instead of D)
- For each advocacy case, provide:
  - The item(s) affected
  - Why the student may deserve credit
  - Confidence level (high/medium/low)
  - Recommended action: full credit, partial credit, or just a softer explanation in the feedback
- Do NOT advocate for genuinely wrong work — missing answers, fundamentally wrong functions (STDEV.P vs STDEV.S), or blank cells are not advocacy candidates
- The goal is fairness, not leniency — advocate only where a reasonable professor might agree

Output: Per-item advocacy notes + recommended credit adjustments + cascade forgiveness recommendations.

---

## Phase 2 — Reconciliation + Feedback Report

### Step 1: Merge Agent Results

| Aspect | Rule |
|--------|------|
| Numerical scores | Script is authoritative. If Agent C found a discrepancy, flag for professor review. |
| Error diagnosis | Merge all agents. For formula issues: prefer Agent A's diagnosis. For conceptual issues: prefer Agent B's. |
| Qualitative scores | Average of Agent B and Agent C scores (each 0 or 1). If they disagree (0.5 average), flag for professor. |
| Positives | Union of all agents' positive observations, deduplicated. |
| Flags | Union of all flags from all agents. |
| Student advocacy | Agent D's recommendations are presented as flags for professor review. If Agent D recommends credit with high confidence AND the error is minor (within 2% of correct), auto-flag as "advocacy: consider credit." If Agent D identifies cascade unfairness (1 root -> 5+ items), present cascade forgiveness recommendation with the root cause and downstream item count. |

### Step 2: Present Flags to Professor

For each flagged item, present:
- The flag description
- Each agent's assessment
- Point impact
- Ask for a decision

### Step 3: Generate Feedback Report

Per `references/global/feedback-report-format.md` with these adaptations:

**Sections used:** Header, What You Did Well, Where You Lost Points, Score Summary Table, Closing Guidance.

**Section skipped:** Modeling Quality (not applicable to problem sets).

**"Where You Lost Points" adaptation:**
- Group by root concept, not by item number
- Each error group shows: which items, what the student did, why it's wrong, and what concept to review
- Include Excel function explanations (undergraduate audience)
- **CRITICAL: Never show correct answers, correct formulas, or correct values in any section of the report.** This applies to error explanations, praise sections, score commentary, and closing guidance. Explain *why* the student's approach is wrong by referencing their actual cells and formulas. Guide them toward the right concept without giving the answer. For example: "Your formula omits the mean return -- parametric VaR includes both the expected return and the volatility component" rather than citing the correct formula or its result. When praising correct work, describe what the student did conceptually (e.g., "Your Solver weights matched the optimal solution") rather than citing specific numerical values from the answer key (e.g., NOT "Your weights of 55.6%, 31.9%, 12.5% matched the key"). See Professor Rulings SS4.

**Score Summary Table:**

Use the numerical total as the denominator (e.g., 79 for A01). Qualitative items are scored separately and shown in the summary table but NOT included in the headline score.

```
| Section | Right | Total | Pct |
|---------|-------|-------|-----|
| Q1      | 18    | 21    | 86% |
| Q2      | 14    | 16    | 88% |
| Q3      | 38    | 42    | 90% |
| Qualitative | 2 | 3     | 67% |
| **Total** | **70** | **79** | **89%** |
```

The Total row reflects numerical items only. Qualitative items appear in the table for transparency but do not affect the headline score or percentage.

**Tone calibration:**

| Score Range | Tone |
|-------------|------|
| 90-100% | Encouraging -- "Excellent work" |
| 80-89% | Positive -- "Good work overall" |
| 70-79% | Constructive -- "Solid effort, review items below" |
| 60-69% | Direct -- "Several areas to work on" |
| Below 60% | Blunt -- "Needs significant improvement" |

**Professor Flags section:** Every non-perfect report MUST have a `## Professor Flags` section at the end of the markdown report. This section contains Agent D advocacy cases, scoring discrepancies, and any items requiring professor decision. If no flags exist, include the section with "No flags identified." This section is stripped from student-facing DOCX output.

### Step 4: Convert to DOCX

Run `scripts/md2docx.py` to convert all markdown reports to student-facing DOCX files:

```bash
python3 scripts/md2docx.py --indir reports/ --outdir docx/
```

The converter:
- Parses markdown structure (headings, bullets, tables, inline formatting)
- Applies score-based color coding (green >=90%, blue >=80%, goldenrod >=70%, orange >=60%, red <60%)
- **Strips the `## Professor Flags` section** — students never see advocacy cases or internal flags
- Uses Calibri 11pt, Light Grid Accent 1 tables

### Step 5: Log Results

Log grading results via `life log teaching` with:
- Number of students graded
- Score distribution summary
- Common error patterns across the class
- Agent agreement rates
- Qualitative scoring consistency
- Any scoring corrections found by Agent C

---

## Orchestration

### Per-Student Flow (Full Ensemble)

1. Read student data bundle from scoring-report.json
2. If perfect scorer: generate template report, skip to next student
3. Launch single Opus agent that executes all 4 roles sequentially:
   a. Agent A — Formula Tracer: trace formula chains, classify errors
   b. Agent B — Conceptual Evaluator: diagnose concepts, grade qualitative items, identify positives
   c. Agent C — Completeness & Cross-Check: verify scoring, find anomalies, grade qualitative items
   d. Agent D — Student Advocate: find legitimate credit cases, cascade forgiveness
   e. Reconcile all 4 assessments into a single feedback report
4. Write markdown report (with Professor Flags section)
5. Move to next student

### Batch Flow

1. Run grade-pset.py -> scoring-report.json (all students, ~30 seconds)
2. Generate template reports for perfect scorers (script, no agent)
3. Extract data bundles for non-perfect students
4. Dispatch agents in parallel waves (~4 students per wave)
5. After all complete: run scripts/md2docx.py to convert all markdown -> DOCX
6. Generate class summary statistics

### Quick Mode (Single-Agent Sonnet)

For low-stakes or time-constrained assignments, replace the full ensemble with a single Sonnet agent per student that performs all 4 roles in a condensed pass. Quality is good but not as deep as the full Opus ensemble.

---

## Professor Rulings

Decisions made during calibration that apply to all future runs. Agents must follow these unless the professor overrides for a specific student.

1. **Error propagation (cascade forgiveness):** Penalize once at the root error. Downstream items with correct formula structure get full credit. Example: if a wrong tangency weight cascades to wrong portfolio return, variance, volatility, and Sharpe ratio -- penalize the weight, credit the rest.

2. **Chart scoring without visual access:** Agents cannot evaluate charts. Default to 1/1 if the chart object exists in the workbook with axes, title, and data series present. Flag for professor review only if the chart appears to use fundamentally wrong data (e.g., references empty columns). Professor grades charts manually during final review.

3. **Date boundary edge cases:** When a student uses AVERAGEIFS with a date filter that excludes one boundary month due to calendar quirks (e.g., `>=1/31/2010` misses Jan 29), this is a data-handling edge case, not a conceptual error. Flag for professor -- default recommendation is full credit if the formula structure and annualization logic are correct.

4. **No correct answers in student reports:** Feedback reports must never show correct formulas, correct numerical answers, or answer key values **in any section** -- including "What You Did Well," "Where You Lost Points," score commentary, and closing guidance. When explaining errors, reference the student's actual cells and formulas and guide toward the right approach. When praising correct work, describe what the student did conceptually ("Your Solver weights matched the optimal solution") without citing specific values from the answer key. This prevents answer sharing between students.

5. **Qualitative text answers:** Grade 1/1 if the answer demonstrates understanding of the key concept, addresses the question, and uses correct reasoning. Minor spelling errors or informal language do not affect the score. Grade 0/1 only if the answer is missing, fundamentally wrong, or fails to address the question.

---

## Important Rules

1. **Never auto-post grades** — always present to professor first
2. **Script scoring is authoritative** for numerical items — agents diagnose, not re-score
3. **Flag uncertainty** — if any agent is unsure, flag for review
4. **Read everything first** — agents must read the scoring report and formulas before diagnosing
5. **Each agent diagnoses independently** — no agent sees another's work until reconciliation
6. **Qualitative items need consensus** — both Agent B and C must assess
7. **Agent D advocates, professor decides** — advocacy recommendations are always flagged for review, never auto-applied
8. **Log everything** — every session feeds calibration
9. **Feedback is the final output** — students never see agent reports or scoring JSON
10. **Explain the Excel** — this is undergraduate; explain which function to use and why
11. **Never show correct answers** — reports reference student cells and formulas, explain why they're wrong, and guide toward the right concept. Never include correct values, correct formulas, or answer key data **in any section of the report, including praise**. When complimenting correct work, describe the concept or approach, not specific numerical values. See Professor Rulings SS4.
