# Bangkok Patana: Asia and Europe comparison (work in progress, paused 18 Nov 2026)

Goal: the BP SUPP report compares Bangkok Patana with aggregate Asia and Europe
international-school averages instead of U.S. (MTF) norms, in every comparison
table, chart, highlight box and bullet. Year 9, 11 and 13 (= 8th, 10th, 12th).
No student counts are shown anywhere.

Done
- Asia and Europe benchmarks built (`Asia Europe benchmarks 2026.csv`): each school
  cleaned with the executive-report rules (`clean_school_for_benchmarks.py`), then
  combined and weighted by students. Asia: ISM Fall 2026, ISB Fall 2026, AES Fall 2026,
  IST Fall 2025, ISKL, UNIS Hanoi. Europe: ASH. Same format as the MTF file plus a
  Region column.
- `regional_comparison_block.py`: the new comparison step (tables, charts, bullets,
  highlights) written, not yet inserted into the BP script or tested.

Still to do
1. BP script: load the benchmark file at the MTF prompt (default next to the script),
   set BENCH / BENCH_NOTE, insert the block before `ReportListofVariables`.
2. BP template: rebuild the 8 measure tables (Asia, Europe, school per year group), the
   6 substance tables and 3 gender tables; replace "U.S. norms" wording; remove the MTF
   footnote rows; rewrite footnote 1; "less than the Asia average on X and the Europe
   average on Y" sentences.
3. Run with the real BP data, check, and deliver.
