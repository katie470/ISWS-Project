# Bangkok Patana (BP) specialised SUPP report, Nov 26 update

Keep these three files in the same folder:
- `BPSUPPcode Nov 26 update.py`: the BP report script (British year groups, gambling section, no Kessler or stressor sections)
- `BPtemplate Nov 26 update.docx`: its template
- `Asia Europe benchmarks 2026.csv`: Asia and Europe international-school averages the report compares against

## Comparisons
Every comparison that used U.S. (MTF) norms now compares Bangkok Patana with the Asia and
Europe international-school averages: the comparison tables, charts, highlight boxes and
bullets, for Year 9, 11 and 13 (U.S. 8th, 10th and 12th grade equivalents). No student
counts are printed.

The averages come from Prevention Ed SUPP surveys, each cleaned with the executive-report
rules, combined per region and weighted by students:
- Asia: ISM (Fall 2026), ISB (Fall 2026), AES (Fall 2026), IST (Fall 2025), ISKL, UNIS Hanoi
- Europe: ASH (one school)

At the benchmark question, press Enter to use the file next to the script.

## Rebuilding
- `tools/clean_school_for_benchmarks.py SURVEY.csv OUT.pkl`: cleans one school with the main SUPP rules.
- `tools/bp_regional_template.py IN.docx OUT.docx`: converts U.S. comparison tables and wording in the BP template to Asia and Europe.
