##############################################################################
### REGIONAL COMPARISON (Nov 26 update): every comparison that used U.S. (MTF)
### norms now compares the school with the Asia and Europe international-school
### averages from the benchmark file (Prevention Ed SUPP data, weighted by
### students). Builds the comparison tables, charts, highlight boxes and
### bullets; overrides the U.S.-based values passed to the template.
##############################################################################
REGIONS = [r for r in ['Asia', 'Europe'] if r in BENCH]
_CMP_GRADES = [g for g in ['8th', '10th', '12th'] if g in _grades_seq]
_TL_WORD = {'lifetime': 'lifetime', 'past-12-months': 'past-12-month', 'past-30-days': 'past-30-day'}
_VERDICT_WORD = {'below': 'below', 'onpar': 'on par with', 'above': 'above', 'sigabove': 'well above'}

def _bench_val(region, var, tl, grade, typ='Total'):
    _b = BENCH[region]
    _r = _b[(_b['variable'] == var) & (_b['timeline'] == tl) & (_b['Type'] == typ)]
    if len(_r) != 1:
        return None
    v = pd.to_numeric(_r[grade].iloc[0], errors='coerce')
    return None if pd.isna(v) else float(v)

def _fmt_cell(v):
    try:
        if v is None or v == '' or pd.isna(v):
            return '-'
        return f"{float(str(v).replace('%', '')):.1f}%"
    except (TypeError, ValueError):
        return '-'

def _school_total(tbl, grade, typ='Total'):
    try:
        v = tbl.loc[tbl['sex'] == typ, grade].iloc[0]
        v = pd.to_numeric(str(v).replace('%', ''), errors='coerce')
        return None if pd.isna(v) else float(v)
    except Exception:
        return None

# ---------- 1. measure tables: label | per grade: Asia, Europe, school ----------
_MEASURES = {
    'alcohollifetumetable': ('Alcohol', 'lifetime', everdrank_grade_sex),
    'alcoholl12monthtstable': ('Alcohol', 'past-12-months', alcohol_past_12monthbysex),
    'alcoholl30daysstable': ('Alcohol', 'past-30-days', alcohol_past_30daysbysex),
    'pot_lifetime_MTF': ('Marijuana', 'lifetime', pot_lifetime),
    'pot_12_MTF': ('Marijuana', 'past-12-months', POTPAST12),
    'pot_30_MTF': ('Marijuana', 'past-30-days', POTPAST30),
    'VAPE_12_MTF': ('Vaping Nicotine', 'past-12-months', VAPEPAST12),
    'VAPE_30_MTF': ('Vaping Nicotine', 'past-30-days', VAPEPAST30),
}
CMP_TABLES = {}
for _key, (_var, _tl, _tbl) in _MEASURES.items():
    _rows = []
    for _typ in ['Total', 'Male', 'Female']:
        _row = [_typ]
        for _g in ['8th', '10th', '12th']:
            for _reg in ['Asia', 'Europe']:
                _row.append(_fmt_cell(_bench_val(_reg, _var, _tl, _g, _typ) if _reg in BENCH else None))
            _row.append(_fmt_cell(_school_total(_tbl, _g, _typ) if _g in _CMP_GRADES else None))
        _rows.append(_row)
    CMP_TABLES[_key] = pd.DataFrame(_rows)
    QA_CONTEXT[_key] = CMP_TABLES[_key]

# ---------- 2. charts: Total only, year group x (Asia, Europe, school) ----------
_CMP_PALETTE = {'Asia': '#8DB4E2', 'Europe': '#C9D2DA', str(School): '#1A63CD'}
def regional_chart(key, save):
    _var, _tl, _tbl = _MEASURES[key]
    _pts = []
    for _g in _CMP_GRADES:
        for _reg in REGIONS:
            _v = _bench_val(_reg, _var, _tl, _g)
            if _v is not None:
                _pts.append({'Grade': _grade_label(_g), 'Source': _reg, 'value': _v})
        _s = _school_total(_tbl, _g)
        if _s is not None:
            _pts.append({'Grade': _grade_label(_g), 'Source': str(School), 'value': _s})
    _df = pd.DataFrame(_pts)
    if _df.empty:
        return
    _draw_bars(_df, 'Grade', 'value', 'Source', _CMP_PALETTE, save, CHART_SIZE_MTF, _grade_axis(),
               hue_order=[h for h in REGIONS + [str(School)] if h in set(_df['Source'])])
for _key, _png in [('alcohollifetumetable', 'alcol_melt_chart.png'), ('alcoholl12monthtstable', 'alcohol12monthschart.png'),
                   ('alcoholl30daysstable', 'alcohol30monthschart.png'), ('pot_lifetime_MTF', 'lifetimepotchart.png'),
                   ('pot_12_MTF', 'annualpotchart.png'), ('pot_30_MTF', 'thirtypotchart.png'),
                   ('VAPE_12_MTF', 'annualvapechart.png'), ('VAPE_30_MTF', 'thirtyvapechart.png')]:
    regional_chart(_key, os.path.join(chart_dir, _png))

# ---------- 3. substance tables per grade: substance | Asia | Europe | school ----------
_SUB_PLAIN = {'Electronic Vaporizers**': 'Electronic Vaporizers'}
def _bench_sub_table(grade, tl, typ='Total'):
    out = {}
    for _reg in REGIONS:
        _b = fix_MTF_naming(BENCH[_reg][(BENCH[_reg]['timeline'] == tl) & (BENCH[_reg]['Type'] == typ)][[grade, 'variable']])
        out[_reg] = dict(zip(_b['variable'], pd.to_numeric(_b[grade], errors='coerce')))
    return out
CMP_SUB = {}
for _g in ['8th', '10th', '12th']:
    for _tl, _suf in [('past-12-months', '12'), ('past-30-days', '30')]:
        _school = substances_DF[(substances_DF['grade_level'] == _g) & (substances_DF['timeline'] == _tl)]
        _bm = _bench_sub_table(_g, _tl)
        _rows = []
        for _, _r in _school.iterrows():
            _sv = pd.to_numeric(str(_r['percent']).replace('%', ''), errors='coerce')
            _rows.append((_SUB_PLAIN.get(_r['substance'], _r['substance']), _bm.get('Asia', {}).get(_r['substance']),
                          _bm.get('Europe', {}).get(_r['substance']), _sv))
        _rows.sort(key=lambda t: -(t[3] if t[3] is not None and not pd.isna(t[3]) else -1))
        _df = pd.DataFrame([[a, _fmt_cell(b), _fmt_cell(c), _fmt_cell(d)] for a, b, c, d in _rows])
        CMP_SUB[(_g, _tl)] = _rows
        QA_CONTEXT[f'ALLSUBSTANCE{_g}_{_suf}_MTF'] = _df if _g in _CMP_GRADES else pd.DataFrame()

# ---------- 4. gender tables: substance | Males: Asia, Europe, school | Females: ... ----------
for _g in ['8th', '10th', '12th']:
    _sx = globals().get(f'sub12months_sex{_g}')
    _rows = []
    if isinstance(_sx, pd.DataFrame) and len(_sx):
        _bm_m = _bench_sub_table(_g, 'past-12-months', 'Male'); _bm_f = _bench_sub_table(_g, 'past-12-months', 'Female')
        for _, _r in _sx.iterrows():
            _sub = _r['substance']
            _rows.append([_SUB_PLAIN.get(_sub, _sub),
                          _fmt_cell(_bm_m.get('Asia', {}).get(_sub)), _fmt_cell(_bm_m.get('Europe', {}).get(_sub)),
                          _fmt_cell(_r.get(f'{_g} Male')),
                          _fmt_cell(_bm_f.get('Asia', {}).get(_sub)), _fmt_cell(_bm_f.get('Europe', {}).get(_sub)),
                          _fmt_cell(_r.get(f'{_g} Female'))])
        _rows.sort(key=lambda r: -float(r[3].rstrip('%')) if r[3] != '-' else 1)
    QA_CONTEXT[f'sub12_by_sex_MTFcombo{_g}'] = pd.DataFrame(_rows)

# ---------- 5. sentences ----------
def _verdicts(var, tls, grade, school_by_tl, region):
    out = {}
    for _tl in tls:
        _b = _bench_val(region, var, _tl, grade); _s = school_by_tl.get(_tl)
        if _b is None or _s is None:
            continue
        out[_tl] = norm_verdict(_s, _b, grade_size(grade))
    return out

def _region_clause(vmap, region):
    groups = {}
    for _tl, _v in vmap.items():
        groups.setdefault(_v, []).append(_TL_WORD[_tl])
    parts = [f"{_VERDICT_WORD[v]} the {region} average for {_join_and(tls)} use" for v, tls in groups.items()]
    return _join_and(parts)

def measure_bullet(var, tls, grade, school_by_tl):
    """'is below the Asia average for lifetime and past-12-month use, and above the Europe average for ...'"""
    clauses = [c for c in (_region_clause(_verdicts(var, tls, grade, school_by_tl, r), r) for r in REGIONS) if c]
    return ('is ' + '; it is '.join(clauses) + '.') if clauses else ''

_ALC = {'lifetime': everdrank_grade_sex, 'past-12-months': alcohol_past_12monthbysex, 'past-30-days': alcohol_past_30daysbysex}
_POT = {'lifetime': pot_lifetime, 'past-12-months': POTPAST12, 'past-30-days': POTPAST30}
_VAP = {'past-12-months': VAPEPAST12, 'past-30-days': VAPEPAST30}
def _school_by_tl(tbls, grade):
    return {tl: _school_total(t, grade) for tl, t in tbls.items()}
for _i, _g in enumerate(['8th', '10th', '12th'], 1):
    _ok = _g in _CMP_GRADES
    QA_CONTEXT[f'albullet{_i}'] = measure_bullet('Alcohol', list(_ALC), _g, _school_by_tl(_ALC, _g)) if _ok else ''
    QA_CONTEXT[f'maribullet{_i}'] = measure_bullet('Marijuana', list(_POT), _g, _school_by_tl(_POT, _g)) if _ok else ''
QA_CONTEXT['albullet1withare'] = ('are' + QA_CONTEXT['albullet1'][2:]) if QA_CONTEXT['albullet1'].startswith('is ') else QA_CONTEXT['albullet1']
QA_CONTEXT['vapebullets'] = [f"Vaping nicotine by {School} {_g} graders {measure_bullet('Vaping Nicotine', list(_VAP), _g, _school_by_tl(_VAP, _g))}"
                             for _g in _CMP_GRADES if measure_bullet('Vaping Nicotine', list(_VAP), _g, _school_by_tl(_VAP, _g))]
QA_CONTEXT['vapebullets'] += [''] * max(0, 6 - len(QA_CONTEXT['vapebullets']))

def _compare_intro2(noun, tbls, var):
    parts = []
    for _g in _CMP_GRADES:
        _b = measure_bullet(var, list(tbls), _g, _school_by_tl(tbls, _g))
        if _b:
            parts.append(f"{noun} by {_g} graders {_b[:-1]}")
    return ('Compared with the Asia and Europe international school averages, ' + '; '.join(
        p[0].lower() + p[1:] for p in parts) + '.') if parts else ''
QA_CONTEXT['alcohol_compare_intro'] = _compare_intro2('Alcohol use', _ALC, 'Alcohol')
QA_CONTEXT['cannabis_compare_intro'] = _compare_intro2('Cannabis use', _POT, 'Marijuana')
QA_CONTEXT['vape_compare_intro'] = _compare_intro2('Vaping nicotine', _VAP, 'Vaping Nicotine')

_life_parts = []
for _g in _CMP_GRADES:
    for _reg in REGIONS:
        _v = _verdicts('Alcohol', ['lifetime'], _g, _school_by_tl(_ALC, _g), _reg).get('lifetime')
        if _v:
            _life_parts.append(f"{_VERDICT_WORD[_v]} the {_reg} average for {_g} graders")
QA_CONTEXT['exec_alcohol_trend_sentence'] = (
    f"Lifetime alcohol use {QA_CONTEXT['lifetime_trend']}" + (f", and is {_join_and(_life_parts)}." if _life_parts else '.')
) if QA_CONTEXT.get('lifetime_trend') else ''

# substance tables: highlight boxes and the "less than the average on X% of substances" sentences
def _sub_summary(grade, tl):
    rows = CMP_SUB.get((grade, tl), [])
    out = {}
    for _ri, _reg in enumerate(REGIONS):
        _pairs = [(s, b, c) for s, *bm, c in [(r[0], r[1], r[2], r[3]) for r in rows]
                  for b in [bm[_ri]] if b is not None and not pd.isna(b) and c is not None and not pd.isna(c)]
        _v = [(s, norm_verdict(c, b, grade_size(grade))) for s, b, c in _pairs]
        out[_reg] = _v
    return out
_window = {'past-12-months': 'the past 12 months', 'past-30-days': 'the past 30 days'}
for _tl, _key in [('past-12-months', 'substance12_highlights'), ('past-30-days', 'substance30_highlights')]:
    _hl = []
    for _g in ['8th', '10th', '12th']:
        if _g not in _CMP_GRADES:
            _hl.append(''); continue
        _sm = _sub_summary(_g, _tl)
        _bits = []
        for _reg, _v in _sm.items():
            if not _v:
                continue
            _below = sum(v == 'below' for _, v in _v); _above = [s.lower() for s, v in _v if v in ('above', 'sigabove')]
            _bits.append(f"lower use than the {_reg} average for {_below} of {len(_v)} substances"
                         + (f" and higher use for {_join_and(_above)}" if _above else ''))
        _hl.append(f"Compared with international schools, {School} {_g} graders report {'; and '.join(_bits)} in {_window[_tl]}." if _bits else '')
    QA_CONTEXT[_key] = _hl
for _key, _g, _tl in [('eighthbelowmeasure12', '8th', 'past-12-months'), ('tenthbelowmeasure', '10th', 'past-12-months'),
                      ('twelvebelowmeasure', '12th', 'past-12-months'), ('eighthbelowmeasure30', '8th', 'past-30-days'),
                      ('tenthbelowmeasure_30', '10th', 'past-30-days'), ('twelvebelowmeasure_30', '12th', 'past-30-days')]:
    _sm = _sub_summary(_g, _tl) if _g in _CMP_GRADES else {}
    for _reg, _suffix in [('Asia', ''), ('Europe', '_eu')]:
        _v = _sm.get(_reg, [])
        if _v:
            _b = sum(v == 'below' for _, v in _v)
            _p = pct(_b, len(_v), f'{_g} {_tl}: substances below the {_reg} average', f'{_g} substances')
            QA_CONTEXT[_key + _suffix] = f"{_p._num()}% ({_b} of {len(_v)})"
        else:
            QA_CONTEXT[_key + _suffix] = '-'

# gender bullets
def _sex_bullets(grade):
    out = []
    _sx = QA_CONTEXT.get(f'sub12_by_sex_MTFcombo{grade}')
    if grade not in _CMP_GRADES or not isinstance(_sx, pd.DataFrame) or _sx.empty:
        return ['', '']
    for _sex, _cols, _word in [('Male', (1, 2, 3), 'males'), ('Female', (4, 5, 6), 'females')]:
        _n = grade_size(grade, _sex)
        _bits = []
        for _ri, _reg in enumerate(['Asia', 'Europe']):
            if _reg not in REGIONS:
                continue
            _vs = []
            for _, _r in _sx.iterrows():
                _b = pd.to_numeric(str(_r[_cols[_ri]]).rstrip('%'), errors='coerce')
                _s = pd.to_numeric(str(_r[_cols[2]]).rstrip('%'), errors='coerce')
                if pd.isna(_b) or pd.isna(_s):
                    continue
                _vs.append((_r[0], norm_verdict(_s, _b, _n)))
            if _vs:
                _above = [s.lower() for s, v in _vs if v in ('above', 'sigabove')]
                _bits.append(f"below the {_reg} average for {sum(v == 'below' for _, v in _vs)} of {len(_vs)} substances"
                             + (f" and above it for {_join_and(_above)}" if _above else ''))
        out.append(f"Past-12-month substance use by {School} {grade} grade {_word} is {'; it is '.join(_bits)}." if _bits else '')
    return out
QA_CONTEXT['outputstringssexbullets'] = _sex_bullets('8th')
QA_CONTEXT['outputstringssexbullets10th'] = _sex_bullets('10th')
QA_CONTEXT['outputstringssexbullets12th'] = _sex_bullets('12th')
QA_CONTEXT['cmp_grades_title'] = f"Comparison between Asia and Europe International School Averages and {School} {_join_and(_CMP_GRADES)} Graders"
QA_CONTEXT['benchmark_note'] = BENCH_NOTE
