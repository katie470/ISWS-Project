# SUPP report pipeline: November 2026 fixes

Updated 14 November 2026. Source: "SUPP Executive Report: fixes for the Python code and the template", a list of problems found when one school's report was corrected by hand. Every fix is a general rule: the scripts and templates work out the numbers, wording, grade lists and page layout from whatever survey data is loaded, for any school. These files replace the October 2026 versions:

| File | What it is |
|---|---|
| `Executive Report with Appendices 9 7 2026.py` | Pipeline for 8th-12th grade schools |
| `Executive Report with Appendices With 7th Grade.py` | Pipeline for 7th-12th grade schools |
| `Fall Executive Report with Appendices Template.docx` | Word template for the 8th-12th script |
| `Fall Executive Report with Appendices Template With 7th Grade.docx` | Word template for the 7th grade script |
| `fix_templates.py` | The script that applied the template fixes (T1-T20). Run it again on any new copy of a template: `python3 fix_templates.py IN.docx OUT.docx` |

Keep each script next to its template. The setup questions are unchanged, so the SUPP Report Maker form keeps working without any change. The one new option (the report date) is set without a new question; see C25 below.

## New outputs on every run

Saved next to the report, alongside the October files (`_QA_metrics.csv`, `_QA_flags.csv`, `_QA_lint.txt`, build record):

- `<report>_exclusion_log.csv` (C2): one row per excluded response with its response ID (the Typeform `#` or Token), grade, sex and the rule(s) it triggered. The report can be rebuilt from the raw CSV plus this log.
- `_QA_lint.txt` now starts with **PART 4 CHECKS**: the ten checks from the fix list, each PASS or FAIL. Any FAIL stops the run with "BUILD CHECK FAILED" after the file is saved, like the other lint errors.

## Settings (top of each script, search for `NOV 2026 FIXES`)

- `EXCLUDE_FIRST_DRINK_AGE`, `EXCLUDE_ALCOHOL_12M_CONTRADICTIONS`, `EXCLUDE_MOSTLY_BLANK` (all on) and `MIN_ANSWERED_ITEMS = 40`: the new exclusion rules.
- `SHOW_N_BELOW = 10`: a percentage with a smaller base is printed with its count, e.g. "50.0% (1 of 2)".
- `MIN_GROUP_FOR_GRADES_CLAIM = 20`: below this, the grades section makes no claim in either direction.
- Report date: today's date unless the environment variable `SUPP_REPORT_DATE` is set (any format).
- Chart sizes: `CHART_SIZE_MTF`, `CHART_SIZE_PAIR`, `CHART_SIZE_K10`, `CHART_FONT_PT`.

## Python fixes

**C1 Norms file.** One reader (`read_mtf`) for the startup check and the analysis. It accepts comma, semicolon or tab files and files with no header row (previously a file without a header lost its first data row, e.g. the Alcohol lifetime Total row, because it was read as the header). The run stops at startup if any substance and timeline lacks a Total, Male or Female row.

**C2/C3 Exclusions.** All rules run in one place and are named:

| Code | Rule |
|---|---|
| R1a ALL_DRUGS | reported every listed hard drug in the 12-month or 30-day list |
| R1b CANNABIS_AGE | age of first cannabis use given but "Never" for frequency (or the reverse). This used to be removed silently before everything else |
| R1c FIRST_DRINK_AGE | first-drink age older than the student's age, or a first-drink age with free text saying they do not drink |
| R1d ALCOHOL_12M | lifetime drinker, "No" to a drink in the past 12 months, but alcohol ticked in the 12-month list |
| R1e MOSTLY_BLANK | fewer than `MIN_ANSWERED_ITEMS` questions answered |
| R1f LIFETIME_NEVER | past-year use contradicts their own "never used" answer |
| R1g SPEEDER | under `MIN_COMPLETION_MINUTES` (raw Typeform downloads only) |
| R1h TIME_ORDER | past-30-day hard-drug use with no past-12-month use |

The Reliability section prints how many surveys each rule removed.

**C4 One result per statistic.** Every named percentage is stored once with its numerator and denominator (`RESULTS`). The Culture of Use "agree" figures under the first two tables now divide by the students who answered the question, not by every student in those grades, so they match the same figure elsewhere in the report.

**C5** "Use less than the U.S. average on X% of substances" uses the same one-student-margin classifier as the highlight box, counts only substances with a U.S. norm, and prints the count, e.g. "45.5% (5 of 11)".

**C6** "Double or greater" requires the higher-risk value to be above zero (0.0% vs 0.0% no longer counts).

**C7** A 30-day "Yes" counts as a 12-month "Yes" for every substance, in one place, and the methodology says so.

**C8 Kessler (K10).** Cut-offs written down in the code and printed in the methodology: K10 10-19 well, 20-24 mild, 25-29 moderate, 30+ severe; depression subscale (6 items) 20+ high; anxiety subscale (4 items) 13+ high. Missing items no longer count as 0: one skipped item is filled from the student's other items on that scale; more than one and the student is left out of that table. Each K10 table now divides by the students scored on that scale (the anxiety and full K10 columns used the depression count).

**C9** The alcohol-vs-norms sentence lists the measures that are actually below norms (a missing norms row used to drop a measure silently; C1 now stops the run instead).

**C10** Grade lists ("10th and 12th graders", the comparison titles, "grades 9, 10, 11 and 12") are built from grades that have respondents.

**C11** Executive Summary higher-risk breakdown names the count for every grade with higher-risk drinkers, e.g. "1 is in 10th grade, 3 are in 11th grade, and 3 are in 12th grade."

**C12** "Risk factors cluster" examples are printed only for substances where higher-risk drinkers exceed both other groups by more than one student.

**C13** Circumstances of Use: when fewer students answered the location question than drank in the past 30 days, the lead-in says how many answered ("Of the N students who drank within the past 30 days and answered this question").

**C14** Culture of Use paragraph: "mixed" only when the figures are mixed; no figure repeated from the sentence before; the light-drinker claim is tied to a number (who disagree vs who did not drink in the past 30 days).

**C15** Grades section: below 20 higher-risk drinkers or cannabis users, the neutral sentence is printed; "in keeping with other research" only when the data show it; a "more likely" claim needs more than one student's difference.

**C16** Percentages with a base under 10 carry their count wherever they are printed in text.

**C17** Footnote under "Greatest Number of Whole Drinks" when it and the higher-risk count disagree, with the counts filled in.

**C18** "Typically never" in the typical-frequency tables.

**C19** Quoted wording comes from the survey file ("drinking is okay", not "drinking is alright").

**C20** The missing bullet on "...% of 12th graders who drank within the past 12 months..." restored.

**C21-C23 Charts.** Drawn at final size (U.S. norm charts 5.4 x 3.6 in, two-pair pages 4.7 x 2.9 in) with 9 pt text and inserted at 100%. X-axis shows "9th", not "9th (N)". Groups read "Higher-risk" / "Lower-risk".

**C24** A grade with no respondents has no column, bars or table: 8th-grade columns (school and U.S.) are removed and the 8th-grade substance and gender sections are not printed.

**C25** Cover and appendix covers print "Survey date: ..." and "Report date: ...".

## Template fixes (applied by `fix_templates.py`)

- **T1** 1,928 empty spacer paragraphs removed; spacing comes from styles. The cover keeps its position with space-before on the school name.
- **T2** Heading 1 style: page break before + keep with next (no break when a heading directly follows another, e.g. KEY DATA ANALYSIS / Alcohol Use). Vaping Nicotine also starts a page.
- **T3/T4/T6/T7** (also enforced on every generated report): rows never split, header rows repeat, labels and notes stay with their table, a chart under a table stays with it, chart paragraphs are not keep-with-next, tables over 10 rows start a new page, and a page break never leaves a title alone at the bottom of a page.
- **T5** 20 hard page breaks replaced by "page break before".
- **T8** Every table spans the text width (6.5 in) with no indent; Part 3 widths: agree/disagree 1.25 in first column, parental 1.5 in, substance tables 1.85 in, past-30-days frequency 2.0 in, consequences 4.7 in, K10 1.35 in.
- **T9** One left-aligned, single-spaced "SUPP Table Label" style for question text and table titles.
- **T10** One centered comparison title for alcohol, cannabis and vaping: "Comparison between U.S. Norms and {School} {grades} Graders".
- **T11** "Cannabis Use", "Frequency of Cannabis Use", "Younger Students' Perceptions".
- **T12** Header cells: one paragraph each, no spacing.
- **T13** "Higher-risk drinkers" / "Lower-risk drinkers" in all headers.
- **T14** Each statistic / "Yet," pair is kept together.
- **T15/T16** TOC limited to levels 1-2 (sub-headings moved to a non-TOC style); TOC styles Calibri 12, regular, 1.15 spacing; Word refreshes the TOC when the report is opened.
- **T17-T19** "individual students", the Conclusion no longer points to recommendations in the Executive Summary, "Prevention Ed" throughout (the cover said "Prevention Education Solutions").
- **2.2 layouts** One page per U.S.-norm measure; Culture of Use and school attitudes two pairs per page (one when the page also has intro text or bullets); gender tables one grade per page, each labelled; Kessler on three pages; stressor questions evenly spaced.

## Tested

Both scripts were run end to end with their templates on synthetic survey exports of different shapes, and every Part 4 check passed:

- 8th-12th script: a 9th-12th school (491 responses) and an 8th-12th school (411)
- 7th grade script: a 7th-12th school (431) and a 9th-12th school (491, so both the 7th and 8th grade columns drop out)
- A norms file missing its Alcohol lifetime Total row stops at startup; a norms file with no header row is read correctly.

Before sending the next real report, open it in Word (let it update the table of contents when asked), and read `_QA_lint.txt` (Part 4 checks first) and `_exclusion_log.csv`.
