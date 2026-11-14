# SUPP report pipeline: fixes from the EAB Fall 2026 QA Error Log

Updated 4 October 2026. Four files in this folder replace the previous versions:

| File | What it is |
|---|---|
| `Executive Report with Appendices 9 7 2026.py` | Pipeline for 8th-12th grade schools |
| `Executive Report with Appendices With 7th Grade.py` | Pipeline for 7th-12th grade schools |
| `Fall Executive Report with Appendices Template.docx` | Word template for the 8th-12th script |
| `Fall Executive Report with Appendices Template With 7th Grade.docx` | Word template for the 7th grade script |

Each script looks for its template in the same folder, so keep the four files together. Nothing changed in how you run them (same setup questions, same inputs).

## New outputs on every run

Three small files are saved next to the report:

- `<report>_QA_metrics.csv`: every narrative percentage with its numerator, denominator and base. Use it to trace any number in the body.
- `<report>_QA_flags.csv`: students whose answers disagree with each other, and any respondents excluded as invalid.
- `<report>_QA_lint.txt`: problems found in the rendered report (leftover placeholders, author notes, broken text, "majority" under 50%, and so on). If this list is not empty the script saves the report and then stops with a "BUILD CHECK FAILED" message, so a bad report is never mistaken for a finished one.

Settings at the top of each script (search for `QA FIXES`):

- `BODY_GROUP_GRADES = ['10th', '11th', '12th']`: the grades used for the higher-risk / lower-risk / non-drinker groups in the report body. Appendix C always uses all grades. Every body heading now says "(drinking groups: grades 10-12)".
- `MIN_CELL_N = 10` and `MIN_STUDENT_GAP = 1`: small-sample guards.
- `EXCLUDE_ALL_DRUG_REPORTERS` and `EXCLUDE_LIFETIME_CONTRADICTIONS`: the R1 automatic exclusion rules (both on).
- `FAIL_ON_LINT_ERRORS = True`: set to False if you ever want the run to finish despite lint errors.

## Fix by bug

**B1. Two drinker counts.** One set of flags per student (`drank_lifetime`, `drank_12m`, `drank_30d`, `pot_lifetime`, `pot_12m`, `pot_30d`, `vape_12m`, `vape_30d`) is built once after cleaning and used for every count, table and sentence. Drinking follows the follow-up items (47 and 60 for EAB), not the substance list. Any past-year cannabis user counts as a lifetime user, so past-year use can no longer exceed lifetime use. Students whose two answers disagree are listed in `_QA_flags.csv`. The lower-risk group now includes drinkers who said "Never" to typical frequency (R4), so the three groups add up to all drinkers.

**B2. Wrong denominators.** A `pct(n, d, name, base)` helper computes every narrative percentage from counts and records its base. Blank answers are no longer counted as a response ("drinking nan" is gone), so questions asked only of drinkers divide by drinkers who answered. "% of available measures" counts only substances with a U.S. norm. "1 in N" uses rounding, not truncation (25 of 222 is 1 in 9). "Once a month" bullets now say "once a month or more" and use the same base for actual and perceived.

**B3. Hidden grade filter.** One setting (`BODY_GROUP_GRADES`) drives every body group. Headings for Consequences, Policies, Enforcement, Stressors and the three K10 tables say which grades they use. Appendix C says it covers all grades and that the body uses grades 10-12. The "Using substances to cope" K10 column uses the same grades as its neighbours.

**B4. Rounding.** A round-half-up `round()` replaces Python's round-half-even throughout (6.25 shows as 6.3). Combined figures (agree + strongly agree, "1 or more times", "1-2 times a year or less") are recomputed from counts instead of adding rounded row percentages. U.S. norm tables show one decimal (2.9%), not whole numbers with ".0".

**B5. Fixed template wording.** Direction words are generated from the data: `trend_phrase()` writes "rises from", "falls from", or "rises to a peak in 10th grade, then falls", and `range_phrase()` quotes the real lowest and highest grade. "Majority" is only used above 50%, "more than double" only above a 2:1 ratio, and "exceeds U.S. norms" sentences in the Alcohol, Cannabis and Vaping intros, the Culture of Use highlight, the respect-for-choices paragraph, the academic-achievement bullets and the subgroup "clustering" claim are all written from the numbers. A post-render check also flips any remaining "rises from X to Y" where Y is lower than X.

**B6. Row and column mapping.** Consequence labels are keyed by item ID (a dictionary), so the table can no longer shift. The Policy bullets have their own placeholders (`pol_*`) instead of reusing the Enforcement ones. Cannabis normative bullets use cumulative "1-2 times a year or more" / "once a month or more" counts for both actual and perceived. The stressor tables are built by column name with each group's own denominator (and the cannabis columns are now filled in). "Drunk 3 or more times in the last 30 days" reads the 30-day item, not the 12-month one. "Typically have 5 or more drinks" uses the how-many-drinks item.

**B7. Template artifacts.** The author notes ("Do these manually...") are replaced by generated highlight bullets. "14 5.9% students" reads "14 students (5.9%)". The Survey Administration paragraph gives both surveys collected and surveys analyzed. Consequence bullets are in the third person. Yes/No questions always show both rows and read "Yes"/"No" (never "True"). U.S. norm tables keep their Total row even when the MTF file has no matching row. The lint step fails the build on placeholders, doubled tokens and broken punctuation.

**B8. Definitions and footnotes.** The three group definitions are text constants in the code and are printed into the report body and Appendix C from the same place. Footnote 2 (drinker-only base) and footnote 4 (yearly or more frequent cannabis users, grades 10-12) describe what the tables now do; the footnote 3 example is computed for the school; the duplicate footnote under "Younger students' perceptions" now says what it refers to.

**B9. Small-sample guard.** `norm_verdict()` only calls a result below or above U.S. norms when the gap is larger than both the 5% margin and one student's worth of percentage points; "well above" also needs at least 10 students. Groups and bases under 10 students are logged as warnings and listed in `_QA_metrics.csv`.

**R1-R6 (raw data).** R1: respondents who report every listed hard drug, or whose past-year use contradicts a "never used" lifetime answer, are flagged and excluded before analysis (they appear in `_QA_flags.csv`). R2 follows from R1. R3: past-year cannabis users count as lifetime users. R4: the five "Yes but Never" drinkers are lower-risk. R5: the two swapped Appendix titles are corrected in both templates. R6: "drove or rode" counts students, not answers.

## Other template changes

- "Almost", "Nearly" and "Approximately" dropped before exact figures.
- The broken "Have you used hallucinogens other than LSD..." title in Appendix B is completed and the duplicate inhalants table removed.
- The 8th grade gender table header no longer has a stray empty paragraph.
- The "Definitions" box is labelled DEFINITIONS, not HIGHLIGHTS.
- The sentence "the majority of alcohol use at all grade levels is occurring at home or a friend's home" is only printed when it is true; otherwise it reports how many grades it holds for.

## Tested

Both scripts were run end to end on synthetic survey exports (230 students in grades 8-12 and 260 in grades 7-12) with a synthetic MTF file. Both produced a report with no lint errors. Please run the next real school through them and read `_QA_lint.txt` and `_QA_flags.csv` before sending the report.
