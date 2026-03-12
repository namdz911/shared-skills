# Student Feedback Report Format

Standard template for student-facing feedback on graded Excel case submissions. Used by the grading framework (single-agent or ensemble) to produce consistent, actionable feedback.

---

## Report Structure

Every feedback report follows this sequence:

```
1. Header (case name, team/student, score)
2. What You Did Well (positives)
3. Where You Lost Points (errors, grouped by root cause)
4. Modeling Quality Feedback (if applicable)
5. Score Summary Table
6. Closing Guidance
```

---

## Section 1: Header

```markdown
# [Case Name] — [Team/Student Name] Feedback Report

**Score: XX.X / [max]**
```

One line. Score is the final reconciled (or single-agent) score scaled to the assignment's point value.

---

## Section 2: What You Did Well

List 3-6 specific positives. Each positive should:
- Name the specific thing done well (not generic praise)
- Reference the relevant section or concept
- Explain *why* it matters when non-obvious

**Sources for positives:**
- Perfect scores on rubric sections (especially conceptual sections)
- Correct handling of common error points (e.g., depreciation base, terminal value components)
- Sophisticated or creative approaches that demonstrate understanding
- Strong qualitative reasoning in text answers
- Knowledge-execution disconnect items where the *knowledge* was correct (credit the understanding)

**Tone:** Direct and specific. "You correctly identified all eight items for inclusion/exclusion" — not "Great job on Q1!"

---

## Section 3: Where You Lost Points

Group deductions by **root cause**, not by item number. This mirrors the error propagation analysis from grading. For each root cause:

### Structure per Error Group

```markdown
### [N]. [Error Name] (-X.X pts)

**Items affected:** [list]

[1-2 paragraphs explaining:]
- What the correct approach is
- What the student did instead
- Why the student's approach is wrong (the conceptual explanation)
- If error propagation applied, explain that deductions are concentrated here
  and downstream items were not separately penalized

[If relevant: what the student did RIGHT within the error — e.g., correct OCF
despite wrong NWC, correct formula structure despite wrong inputs]
```

### Cell References
When explaining errors, reference specific cells from the student's work file (e.g., "In cell H6, your formula computes `=H5*55%`..."). This helps students locate the exact issue in their workbook and verify the feedback against their own file. Do not reference cells from the answer key.

### Ordering
- Order by point impact (largest deduction first)
- Group related errors under one heading (e.g., all NWC-driven CF errors are one group, not six separate entries)

### Learning Points (No Deduction)
If an item warrants feedback but no points were deducted (e.g., offsetting errors, algebraically equivalent methods), include it as:

```markdown
### [N]. [Topic] — A Learning Point (no points deducted)

[Explain what happened, why no penalty, and what the student should understand
for future cases]
```

### What NOT to Include
- Internal grading mechanics (tier classifications, agent disagreements, reconciliation details)
- References to "the grading framework" or "Agent A/B/C"
- Tolerance bands or partial credit percentages
- Cell references from the answer key (student cell references are OK)

---

## Section 4: Modeling Quality Feedback

Include only when the assignment has a modeling quality component (Section 9 of grading framework).

```markdown
## Modeling Quality (X / 10)

[1-2 sentence overall assessment]

**Areas for improvement:**

- **[Dimension]:** [Specific observation and actionable suggestion]
- **[Dimension]:** [Specific observation and actionable suggestion]
- **[Dimension]:** [Specific observation and actionable suggestion]
```

**Rules:**
- List only dimensions rated Weak or Poor — don't enumerate all six
- Each bullet must include a concrete, actionable suggestion (not just "needs improvement")
- If the student scored 8+ (Good or Excellent), keep this section brief and focus on what was done well

---

## Section 5: Score Summary Table

```markdown
## Summary

| Section | Score | Max | Pct |
|---------|-------|-----|-----|
| [Section name] | X.XX | X.XX | XX% |
| ... | ... | ... | ... |
| Modeling Quality | X.XX | 10.00 | XX% |
| **Total** | **XX.X** | **XX.X** | **XX%** |
```

- Group GradingMap items into logical sections (e.g., Q1: Conceptual, Q2: Initial Investment, Q3: Cash Flows, Q4: Metrics, Q5: Recommendation)
- Section groupings should match the case structure, not the GradingMap item numbers
- Include percentages so the student can see relative performance across sections

---

## Section 6: Closing Guidance

One short paragraph (2-3 sentences) summarizing:
- The student's strongest area
- The 1-2 most impactful things to focus on for improvement
- Forward-looking (next assignment, future cases), not backward-looking

---

## Tone and Style Guidelines

### Tone Calibration by Performance

The overall tone adjusts based on submission quality. Be objective throughout, but shift emphasis:

| Score Range | Tone | Emphasis |
|-------------|------|----------|
| **85-100%** | Encouraging, detailed | Lead with strengths. Errors are refinements, not fundamental gaps. Modeling feedback focuses on moving from good to excellent. |
| **70-84%** | Objective, constructive | Balance positives and errors. Explain root causes clearly. Acknowledge what the student understands even where execution failed. The Team 1 report (84%) is the reference tone for this range. |
| **50-69%** | Direct, matter-of-fact | Shorter positives section — only cite genuinely correct work, don't stretch for compliments. Error explanations are clear and instructive. Focus on the 2-3 most impactful things to fix. |
| **Below 50%** | Straightforward, blunt | Do not sugarcoat. State plainly what is missing or wrong. If large sections are blank, unanswered, or fundamentally broken, say so directly. Skip the positives section entirely if there is nothing substantive to credit — a forced compliment on a poor submission is patronizing, not encouraging. Focus feedback on the most basic requirements the student needs to meet. |

**The principle:** Respect the student's intelligence at every level. A strong student deserves specific, nuanced feedback. A struggling student deserves honest, clear feedback about what went wrong — not vague encouragement that obscures the gap between their work and the standard.

### Style Rules

1. **Direct and specific.** Name the exact error, the exact cell, the exact concept. Avoid vague feedback like "be more careful" or "check your work."

2. **Explain the why.** Don't just say "NPV is wrong." Explain why the Excel NPV function behaves differently from the textbook formula, and show the correct usage.

3. **Credit understanding even when execution fails.** If a student's report says "11% WACC" but their model uses 10%, acknowledge they understand the concept — the error is implementation, not comprehension. (Applies to 50%+ submissions. Below 50%, focus on fundamentals.)

4. **No jargon about the grading process.** The student should not see references to "error propagation rules," "partial credit tiers," or "agents." These are internal grading mechanics. The student sees: "You lost X points because Y."

5. **Positive first, negative second.** The "What You Did Well" section comes before "Where You Lost Points." This is intentional — students read more carefully when they don't feel attacked from the start. Exception: below 50%, skip or minimize the positives section.

6. **Actionable over evaluative.** "Create a dedicated assumptions section and reference those cells with absolute references" is better than "Your model lacks input separation."

---

## Audience Calibration

The same template is used for all course levels. What changes:

| Aspect | Undergraduate | MBA |
|--------|--------------|-----|
| Conceptual explanations | Include — students may be encountering concepts for the first time | Brief — assume baseline knowledge |
| Excel mechanics | Explain (e.g., how NPV() function works) | Reference only (e.g., "NPV() treats first argument as Period 1") |
| Modeling quality feedback | Focus on structure and auditability | Add expectations around scenario analysis, sensitivity |
| Qualitative answer depth | Note if basic explanation is sufficient | Note if deeper strategic reasoning is expected |

The course level comes from `course-config.md`. When in doubt, err toward more explanation — over-explaining costs the student nothing, under-explaining leaves them confused.
