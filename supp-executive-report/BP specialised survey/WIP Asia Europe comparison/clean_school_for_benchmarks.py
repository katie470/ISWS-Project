"""Clean one SUPP export with the main SUPP script's rules and save the cleaned
student-level flags (no free text) for benchmark building."""
import builtins, contextlib, io, os, sys, tempfile
import pandas as pd
PIPE = '/home/user/supp-report-maker/scripts/Executive Report with Appendices With 7th Grade.py'
CUT = 'def risk_categorizer_pot(df):'
csv_path, out_path = sys.argv[1], sys.argv[2]
src = open(PIPE, encoding='utf-8').read(); src = src[:src.index(CUT)]
work = tempfile.mkdtemp()
mtf = os.path.join(work, 'mtf.csv'); open(mtf, 'w').write('Type,Year\nTotal,2024\n')
answers = iter(['7th,8th,9th,10th,11th,12th', '', 'date', 'School', csv_path, mtf, '2024', '', os.path.join(work, 'x.docx')] + [''] * 50)
builtins.input = lambda *a, **k: next(answers)
g = {'__file__': PIPE, '__name__': '__main__', 'SUPP_CLEANING_ONLY': True}
log = io.StringIO(); os.chdir(work)
try:
    with contextlib.redirect_stdout(log):
        exec(compile(src, PIPE, 'exec'), g)
except SystemExit:
    sys.stderr.write(log.getvalue()[-2500:]); raise
c = g['clean']
keep = ['grade_level', 'sex', 'drank_lifetime', 'drank_12m', 'drank_30d', 'pot_lifetime', 'pot_12m', 'pot_30d',
        'vape_12m', 'vape_30d'] + [x for x in c.columns if x.startswith('use_') and x[-3:] in ('_12', '_30')]
c[keep].to_pickle(out_path)
print(f"{os.path.basename(csv_path)}: raw {g['N']}, removed {g['removed_N']}, analyzed {len(c)}, grades {c['grade_level'].value_counts().to_dict()}")
