#!/usr/bin/env python
# coding: utf-8

# In[1]:


##############################
### Libraries
### BP specialised survey version, Nov 26 update (Nov 2026 report fixes applied;
### no Kessler or student stressor sections; gambling section unchanged)
##############################

import pandas as pd
import itertools
from importlib import reload
import typing_extensions
#reload(typing_extensions)
import datetime
import numpy as np
import warnings
warnings.filterwarnings('ignore', category=FutureWarning)
warnings.filterwarnings('ignore', category=DeprecationWarning)
from varname.helpers import Wrapper
#from typing_extensions import Self
from docxtpl import DocxTemplate
import pymsgbox
from docx.oxml import OxmlElement, ns
#import camelot
import seaborn as sns
import re
import os
import csv
import matplotlib
matplotlib.rcParams['figure.max_open_warning'] = 0  # suppress "too many figures" warning





pd.set_option('display.max_columns', None)
pd.options.display.float_format = '{:,.1f}'.format

##############################
### CORRECTED: Utility functions for data quality (added per review recommendations)
##############################

def safe_round(value, decimals=1):
    """Round a numeric value safely, handling edge cases."""
    if isinstance(value, str):
        try:
            value = float(value.replace('%', ''))
        except (ValueError, AttributeError):
            return value
    if pd.isna(value):
        return 0.0
    return round(float(value), decimals)

def cap_percentage(value, cap=100.0):
    """Cap a percentage sum at a maximum value (default 100.0).
    Fixes floating-point summation errors like 100.1%."""
    val = safe_round(value, 1)
    if isinstance(val, (int, float)):
        return min(val, cap)
    return val

def safe_perc_sum(series, cap=100.0):
    """Sum percentage strings (e.g., '30.5%') safely, round and cap result."""
    vals = series.str.split('%').str[0].astype(float)
    return cap_percentage(round(vals.sum(), 1), cap)

def fuzzy_location_match(df, col, target):
    """Match a location string flexibly, handling apostrophe variants."""
    # Try exact match first
    mask = df[col] == target
    if mask.any():
        return mask
    # Try without apostrophes
    target_clean = target.replace("'", "").replace("\u2019", "").lower()
    mask = df[col].str.replace("'", "", regex=False).str.replace("\u2019", "", regex=False).str.lower() == target_clean
    if mask.any():
        return mask
    # Try contains as fallback
    mask = df[col].str.lower().str.contains(target_clean[:15], na=False)
    return mask

class _ItemableValue:
    """Wrapper that makes any scalar value support .item() for pandas compatibility.
    When corrected functions return plain float/int/str instead of pandas Series,
    downstream code calling .item() will still work."""
    def __init__(self, val):
        self._val = val
    def item(self):
        return self._val
    # Make it behave like the wrapped value for arithmetic, comparison, formatting
    def __float__(self):
        return float(self._val) if self._val != '' else 0.0
    def __int__(self):
        return int(float(self._val)) if self._val != '' else 0
    def __str__(self):
        return str(self._val)
    def __repr__(self):
        return repr(self._val)
    def __eq__(self, other):
        return self._val == other
    def __ne__(self, other):
        return self._val != other
    def __lt__(self, other):
        return float(self) < float(other)
    def __le__(self, other):
        return float(self) <= float(other)
    def __gt__(self, other):
        return float(self) > float(other)
    def __ge__(self, other):
        return float(self) >= float(other)
    def __add__(self, other):
        return float(self) + float(other)
    def __radd__(self, other):
        return float(other) + float(self)
    def __sub__(self, other):
        return float(self) - float(other)
    def __rsub__(self, other):
        return float(other) - float(self)
    def __mul__(self, other):
        return float(self) * float(other)
    def __rmul__(self, other):
        return float(other) * float(self)
    def __truediv__(self, other):
        return float(self) / float(other)
    def __rtruediv__(self, other):
        return float(other) / float(self)
    def __round__(self, ndigits=0):
        # B4: half-up (the module-level round() is defined later in the script)
        return float(_Decimal(repr(float(self))).quantize(_Decimal(1).scaleb(-(ndigits or 0)), rounding=_ROUND_HALF_UP))
    def __bool__(self):
        return bool(self._val) and self._val != ''
    def __format__(self, format_spec):
        return format(self._val, format_spec)

def _safe_item(val):
    """Extract scalar from pandas object, or return val if already a scalar/string."""
    if isinstance(val, _ItemableValue):
        return val._val
    if isinstance(val, str):
        return val
    if isinstance(val, (int, float)):
        return val
    try:
        return val.item()
    except (AttributeError, ValueError):
        return val

def _wrap_scalar(val):
    """Wrap a plain scalar (float, int, str) in _ItemableValue so .item() works on it.
    If val is already a pandas type with .item(), return as-is."""
    if isinstance(val, (str, int, float)):
        return _ItemableValue(val)
    return val

import logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')
logger = logging.getLogger('SUPP_Report')

##############################
### QA FIXES (EAB Fall SUPP 2026 QA Error Log, Oct 2026)
### Report-wide settings. Bug codes (B1-B9, R1-R6) refer to the QA log.
##############################

# B3/B8: ONE definition of the drinking-group grades used by every body
# (narrative) statistic. Appendix C always shows all grades.
BODY_GROUP_GRADES = ['10th', '11th', '12th']

# B9: never claim a difference smaller than this many students, and flag
# any group/cell with fewer than MIN_CELL_N students.
MIN_STUDENT_GAP = 1
MIN_CELL_N = 10

# R1: drop respondents who report every listed hard drug, or whose past-year
# answers contradict their own "never used" lifetime answers.
EXCLUDE_ALL_DRUG_REPORTERS = True
EXCLUDE_LIFETIME_CONTRADICTIONS = True

# Completion-time check (raw Typeform downloads only — the Google Sheet export
# has no Start/Submit timestamps, so the check is skipped for those files).
MIN_COMPLETION_MINUTES = 5   # responses submitted faster than this are unreliable
EXCLUDE_SPEEDERS = True      # remove them before analysis (always listed in _QA_flags.csv)

# R5: time-order consistency. A student who used a substance in the past 30 days
# must also count as a past-12-month user. Two rules (both listed in _QA_flags.csv):
#  - A student who reports past-30-day use of a HARD DRUG (heroin, meth, cocaine,
#    etc.) but says they did NOT use it in the past 12 months is treated as an
#    untruthful/inconsistent survey and removed (EXCLUDE_TIME_ORDER_CONTRADICTIONS).
#  - For every other substance (alcohol, tobacco, vapes, cannabis, cold medicine),
#    a past-30-day "Yes" with a past-12-month "No"/blank is corrected to a
#    past-12-month "Yes", so every table uses the same answer.
EXCLUDE_TIME_ORDER_CONTRADICTIONS = True
# Stop with an error before rendering if any grade/sex/drinking group would still
# show past-30-day use above past-12-month use, or if two tables would disagree.
FAIL_ON_TIME_ORDER_ERRORS = True

# B7: stop with an error (after saving the file) if the rendered report still
# contains template placeholders, author notes or broken text.
FAIL_ON_LINT_ERRORS = True

##############################
### NOV 2026 FIXES (from the Nov 2026 report review; general rules for every school)
### Item codes C1-C25 / T1-T20 refer to "SUPP report code and template fixes".
##############################
# C3: named exclusion rules. Every excluded respondent is written to
# <report>_exclusion_log.csv with its response ID and the rule(s) it triggered,
# so the report can be rebuilt from the raw CSV plus that log.
#   R1a ALL_DRUGS        reported every listed hard drug (12-month or 30-day block)
#   R1b CANNABIS_AGE     gave an age of first cannabis use but "Never" for cannabis
#                        frequency (or the reverse)
#   R1c FIRST_DRINK_AGE  age of first drink is older than the student's own age, or
#                        gave a first-drink age but wrote that they do not drink
#   R1d ALCOHOL_12M      lifetime drinker who says "No" to a drink in the past 12
#                        months but ticks alcohol in the 12-month substance list
#   R1e MOSTLY_BLANK     answered fewer than MIN_ANSWERED_ITEMS questions
#   R1f LIFETIME_NEVER   past-year use contradicts their own "never used" answer
#   R1g SPEEDER          finished in under MIN_COMPLETION_MINUTES (raw Typeform only)
#   R1h TIME_ORDER       past-30-day hard-drug use with no past-12-month use
EXCLUDE_FIRST_DRINK_AGE = True
EXCLUDE_ALCOHOL_12M_CONTRADICTIONS = True
EXCLUDE_MOSTLY_BLANK = True
MIN_ANSWERED_ITEMS = 40          # of the ~170 survey questions

# C16: a percentage whose denominator is under this many students is always
# printed with its count, e.g. "50.0% (1 of 2)".
SHOW_N_BELOW = 10
# C15: below this many students in a group, no direction is claimed for the
# grades/achievement comparison.
MIN_GROUP_FOR_GRADES_CLAIM = 20

# C25: two explicit dates. The survey date comes from the setup questions. The
# report date is today unless SUPP_REPORT_DATE is set (any format), so the
# setup questions (and the SUPP Report Maker form) are unchanged.
REPORT_DATE = os.environ.get('SUPP_REPORT_DATE', '').strip() or \
    datetime.date.today().strftime('%B %-d, %Y' if os.name != 'nt' else '%B %d, %Y')

def _grade_num(g):
    try:
        return int(re.sub(r'\D', '', str(g)))
    except ValueError:
        return 99

def grade_range_label(glist, style='ordinal'):
    """['10th','11th','12th'] -> '10th-12th' (ordinal) or '10-12' (plain)."""
    gl = sorted(glist, key=_grade_num)
    if not gl:
        return ''
    if style == 'plain':
        return f"{_grade_num(gl[0])}-{_grade_num(gl[-1])}" if len(gl) > 1 else str(_grade_num(gl[0]))
    return f"{gl[0]}-{gl[-1]}" if len(gl) > 1 else gl[0]

BODY_GROUP_LABEL = grade_range_label(BODY_GROUP_GRADES)                 # '10th-12th'
BODY_GROUP_LABEL_PLAIN = grade_range_label(BODY_GROUP_GRADES, 'plain')  # '10-12'

# ---- B4: one rounding rule everywhere: round half up, 1 decimal ----
# Python's built-in round() and pandas .round() send halves to even
# (6.25 -> 6.2). This shadows round() for the whole script so every
# percentage rounds half up (6.25 -> 6.3, 31.25 -> 31.3).
import builtins as _builtins
from decimal import Decimal as _Decimal, ROUND_HALF_UP as _ROUND_HALF_UP, InvalidOperation as _InvalidOperation

def _rhu_scalar(x, ndigits):
    if x is None or isinstance(x, (bool, np.bool_)) or isinstance(x, str):
        return x
    try:
        if pd.isna(x):
            return x
    except (TypeError, ValueError):
        pass
    try:
        d = _Decimal(repr(float(x)))
        q = _Decimal(1).scaleb(-(ndigits or 0))
        r = d.quantize(q, rounding=_ROUND_HALF_UP)
    except (TypeError, ValueError, _InvalidOperation, OverflowError):
        return x
    if ndigits is None:
        return int(r)
    return np.float64(r)   # np.float64 keeps .item() working

def round(x, ndigits=None):
    """Round-half-up replacement for round(); works on scalars, Series, DataFrames."""
    if isinstance(x, pd.Series):
        return x.apply(lambda v: _rhu_scalar(v, 0 if ndigits is None else ndigits))
    if isinstance(x, pd.DataFrame):
        out = x.copy()
        for c in out.columns:
            if pd.api.types.is_numeric_dtype(out[c]):
                out[c] = out[c].apply(lambda v: _rhu_scalar(v, 0 if ndigits is None else ndigits))
        return out
    if isinstance(x, np.ndarray):
        return np.vectorize(lambda v: _rhu_scalar(v, 0 if ndigits is None else ndigits), otypes=[float])(x)
    if hasattr(x, '_val') and hasattr(x, 'item'):   # _ItemableValue
        return _rhu_scalar(float(x), ndigits)
    if isinstance(x, (int, float, np.integer, np.floating)):
        return _rhu_scalar(x, ndigits)
    return _builtins.round(x, ndigits) if ndigits is not None else _builtins.round(x)

# ---- B2: every narrative percentage is computed from counts, with its base ----
METRIC_LOG = []
# C4: the shared results object. Every named statistic is stored ONCE here
# (value, numerator, denominator, base). Narrative sentences read from it and
# the Part 4 consistency check compares it against the tables.
RESULTS = {}

class _Pct(float):
    """A percentage that remembers its numerator and denominator (C4/C16).
    Behaves like a float in arithmetic. When the denominator is under
    SHOW_N_BELOW, printing it (f-string or template) leaves an invisible marker
    that the post-render step turns into "50.0% (1 of 2)"."""
    def __new__(cls, value, n, d):
        obj = float.__new__(cls, value)
        obj.n, obj.d = int(round(n)), int(round(d))
        return obj
    @property
    def small(self):
        return 0 < self.d < SHOW_N_BELOW
    def _num(self):
        return f"{float(self):.1f}"
    def marked(self):
        """Report text: the number plus an invisible marker when the base is small."""
        if self.small:
            return f"{self._num()}⁣N{self.n}D{self.d}⁣"
        return self._num()
    def __str__(self):
        # str() stays a plain number so code that parses it back still works;
        # f-strings (report sentences) and the template get the marked form.
        return self._num()
    def __format__(self, spec):
        return self.marked() if not spec else format(float(self), spec)
    def __repr__(self):
        return self._num()
    def __reduce__(self):
        return (float, (float(self),))

_SMALL_N_MARKER = re.compile('(\\d+(?:\\.\\d+)?)⁣N(\\d+)D(\\d+)⁣(%?)')

def expand_small_n(text):
    """'50.0<marker>%' -> '50.0% (1 of 2)'. Used on rendered paragraphs."""
    return _SMALL_N_MARKER.sub(lambda m: f"{m.group(1)}{m.group(4)} ({m.group(2)} of {m.group(3)})", text)

def pct(n, d, name=None, base=''):
    """n/d*100 rounded half up to 1 decimal; '' when d is 0. Logged with its base."""
    try:
        n = float(n); d = float(d)
    except (TypeError, ValueError):
        return ''
    val = _Pct(round(n / d * 100, 1), n, d) if d > 0 else ''
    if name:
        METRIC_LOG.append({'metric': name, 'numerator': n, 'denominator': d,
                           'percent': float(val) if val != '' else '', 'base': base,
                           'small_base': bool(d and d < MIN_CELL_N)})
        RESULTS[name] = {'value': float(val) if val != '' else None, 'n': n, 'd': d, 'base': base}
        if d and d < MIN_CELL_N:
            logger.warning(f"B9 SMALL BASE: {name} = {float(val) if val != '' else ''}% uses only {int(d)} students ({base})")
    return val

def fmt_pct(v):
    """'12.5%' for numbers, '-' for blanks; '50.0% (1 of 2)' for small bases (C16)."""
    if v == '' or v is None:
        return '-'
    try:
        if pd.isna(v):
            return '-'
    except (TypeError, ValueError):
        pass
    if isinstance(v, _Pct) and v.small:
        return f"{v._num()}% ({v.n} of {v.d})"
    return f"{round(float(v), 1)}%"

def count_pct_from_table(df, mask, col):
    """B4: combined percentage of the rows in `mask`, recomputed from the table's
    count column (e.g. '12th (N)' for '12th (%)') instead of adding rounded
    row percentages. Returns '' if the column/base is missing."""
    ncol = col.replace('(%)', '(N)') if '(%)' in col else col
    if ncol not in df.columns:
        return ''
    counts = pd.to_numeric(df[ncol], errors='coerce').fillna(0)
    total = counts.sum()
    if total == 0:
        return ''
    return pct(counts[mask].sum(), total)   # C4/C16: keeps its numerator and denominator

def share_word(p):
    """B5: 'a majority' only when over 50%."""
    try:
        p = float(p)
    except (TypeError, ValueError):
        return 'some'
    if p >= 75:
        return 'the vast majority'
    if p > 50:
        return 'a majority'
    if p >= 45:
        return 'about half'
    return 'a minority'

def _join_and(items):
    items = [str(i) for i in items if str(i)]
    if len(items) <= 1:
        return ''.join(items)
    if len(items) == 2:
        return ' and '.join(items)
    return ', '.join(items[:-1]) + ', and ' + items[-1]

def trend_phrase(series, unit='grade'):
    """B5: describe a by-grade series from the numbers themselves.
    series: list of (grade, value) in grade order. Returns e.g.
    'rises from 14.3% in 8th grade to 71.4% in 12th grade',
    'falls from ...', 'stays at ...', or
    'rises from 2.5% in 7th grade to a peak of 33.3% in 10th grade, then falls to 21.4% in 12th grade'."""
    pts = []
    for g, v in series:
        try:
            if v == '' or v is None or pd.isna(v):
                continue
            pts.append((g, v if isinstance(v, _Pct) else float(v)))   # C16: keep the count
        except (TypeError, ValueError):
            continue
    if len(pts) < 2:
        return ''
    (g0, v0), (g1, v1) = pts[0], pts[-1]
    vals = [v for _, v in pts]
    pk_i = int(np.argmax(vals)); lo_i = int(np.argmin(vals))
    if 0 < pk_i < len(pts) - 1 and vals[pk_i] > max(v0, v1):
        gp, vp = pts[pk_i]
        verb2 = 'falls' if v1 < vp else 'stays'
        return (f"{'rises' if vp > v0 else 'changes'} from {fmt_pct(v0)} in {g0} {unit} to a peak of "
                f"{fmt_pct(vp)} in {gp} {unit}, then {verb2} to {fmt_pct(v1)} in {g1} {unit}")
    if 0 < lo_i < len(pts) - 1 and vals[lo_i] < min(v0, v1):
        gl_, vl = pts[lo_i]
        return (f"falls from {fmt_pct(v0)} in {g0} {unit} to a low of {fmt_pct(vl)} in {gl_} {unit}, "
                f"then rises to {fmt_pct(v1)} in {g1} {unit}")
    if v1 > v0:
        verb = 'rises'
    elif v1 < v0:
        verb = 'falls'
    else:
        return f"stays at {fmt_pct(v0)} from {g0} {unit} to {g1} {unit}"
    return f"{verb} from {fmt_pct(v0)} in {g0} {unit} to {fmt_pct(v1)} in {g1} {unit}"

def range_phrase(series):
    """B5: 'ranging from X% of Ath graders to Y% of Bth graders' using the real min and max."""
    pts = []
    for g, v in series:
        try:
            if v == '' or v is None or pd.isna(v):
                continue
            pts.append((g, v if isinstance(v, _Pct) else float(v)))   # C16: keep the count
        except (TypeError, ValueError):
            continue
    if not pts:
        return ''
    lo = min(pts, key=lambda t: t[1]); hi = max(pts, key=lambda t: t[1])
    if lo[1] == hi[1]:
        return f"{fmt_pct(lo[1])} in every grade"
    return f"ranging from {fmt_pct(lo[1])} of {lo[0]} graders to {fmt_pct(hi[1])} of {hi[0]} graders"

def norm_verdict(school, mtf, base_n, rel_margin=0.05, sig_margin=0.075):
    """B9: compare a school % with a U.S. norm. A gap is only claimed when it is
    bigger than both the relative margin and MIN_STUDENT_GAP students' worth
    of percentage points (one student = 100/base_n)."""
    try:
        s = float(school); m = float(mtf)
    except (TypeError, ValueError):
        return None
    if pd.isna(s) or pd.isna(m):
        return None
    one = (100.0 / base_n) if base_n else 0.0
    gap_floor = one * MIN_STUDENT_GAP
    if abs(s - m) <= max(m * rel_margin, gap_floor):
        return 'onpar'
    if s < m:
        return 'below'
    if (s - m) > max(m * sig_margin, 2 * gap_floor) and base_n and base_n >= MIN_CELL_N:
        return 'sigabove'
    return 'above'

##############################
### Ask some questions
##############################
print('Going to ask some questions now')

## grades as in 9th 10th 11th 12th etc
print("\n--- SETUP QUESTIONS ---\n")
response_grade_level = input("""What grade level system does this school use?
If they use '9th, 10th, 11th, 12th', enter 'default'
Otherwise enter as a list with items separated by commas without spaces.
Only enter grade levels you want included in the report.
e.g. 8th,9th,10th,11th,12th
> """).strip()

if response_grade_level == 'default':
    grades = ['9th','10th', '11th', '12th']
else:
    grades = response_grade_level.split(',')

response_grading_scale = input("""\nWhat grading scale does this school use for "How were your grades overall"?
Enter 'letters' for A-F letter grades, or 'numbers' for IB-style 1-7.
[Press Enter to auto-detect from the survey data]
> """).strip().lower()

date_survey = input("\nWhat date(s) did students take the survey? Enter in any format\n> ").strip()

School = input("\nEnter the school's name how you wish it to appear in the report\n> ").strip()

def clean_path(p):
    """Remove quotes and unescape characters added by macOS drag-and-drop."""
    import re as _re
    p = p.strip().strip('"').strip("'")
    # Unescape any backslash-escaped character (spaces, parens, brackets, etc.)
    p = _re.sub(r'\\(.)', r'\1', p)
    return p

surveryoutput = clean_path(input("\nWhere on your computer is the Typeform export of the survey results in CSV form?\ne.g. /Users/Heather/Scripts/PreventionEd/EAB_clean.csv\n(Tip: drag the file from Finder into this window)\n> "))

# Nov 26 update: this school is compared with the Asia and Europe international-school
# averages (not U.S. MTF norms). The benchmark file has the MTF layout plus a Region column.
_default_bench = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'Asia Europe benchmarks 2026.csv')
MTFfilepath = clean_path(input(f"\nWhere is the Asia and Europe benchmark file (CSV)?\n[Press Enter for default: {_default_bench}]\n(Tip: drag the file from Finder into this window)\n> ")) or _default_bench
MTFyear = input("\nWhich benchmark year? [Press Enter for the latest]\n> ").strip()

# Default to the appendices template that lives next to this script
_default_template = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'BPtemplate Nov 26 update.docx')
_docx_input = input(f"\nWhere is the report + appendices template?\n[Press Enter for default: {_default_template}]\n(Tip: drag the file from Finder into this window)\n> ").strip()
docxpath = clean_path(_docx_input) if _docx_input else _default_template

# Reports save into the season folder by default; the QA files land next to the report.
_default_report_dir = '/Users/katiegreeley/Desktop/26 27 SUPP REPORTS'
_default_report_name = f"{School} SUPP Executive Report {datetime.date.today().year}.docx"
_save_input = clean_path(input(f"\nWhere do you want to save the report and appendices file to?\n[Press Enter for default: {os.path.join(_default_report_dir, _default_report_name)}]\n(Or type a filename, or drag a folder from Finder to save there)\n> "))
if not _save_input:
    savepath = os.path.join(_default_report_dir, _default_report_name)
elif os.path.isdir(_save_input) or _save_input.endswith(os.sep):
    # A folder was given (typed or dragged from Finder): use the default filename inside it
    savepath = os.path.join(_save_input, _default_report_name)
elif os.sep not in _save_input:
    # Just a filename: save it in the season folder
    savepath = os.path.join(_default_report_dir, _save_input)
else:
    savepath = _save_input
os.makedirs(os.path.dirname(savepath), exist_ok=True)

# Create a directory for chart images next to the output file
import tempfile as _tempfile
chart_dir = os.path.join(os.path.dirname(savepath), 'supp_charts_tmp')
os.makedirs(chart_dir, exist_ok=True)

##############################
### STARTUP VALIDATION
### Checks all inputs before doing any processing.
### Gives clear error messages so you know what to fix.
##############################
import warnings
warnings.filterwarnings('ignore', category=FutureWarning)
try:
    warnings.filterwarnings('ignore', category=pd.errors.SettingWithCopyWarning)
except AttributeError:
    pass  # newer pandas versions removed this warning class

print("\n--- VALIDATING INPUTS ---")
_errors = []
_warnings = []

# 1. Check survey CSV exists and is readable
if not os.path.isfile(surveryoutput):
    _errors.append(f"Survey CSV not found: {surveryoutput}")
else:
    try:
        with open(surveryoutput, 'r', encoding='latin-1', errors='replace') as _f:
            _sample = ''.join(_f.readline() for _ in range(10))
        try:
            _dialect = csv.Sniffer().sniff(_sample, delimiters=',;\t')
            _sep = _dialect.delimiter
        except csv.Error:
            _first = _sample.split('\n')[0] or ''
            _sep = ';' if _first.count(';') > max(_first.count(','), _first.count('\t')) else ('\t' if _first.count('\t') > _first.count(',') else ',')
        _test_csv = pd.read_csv(surveryoutput, encoding='latin-1', sep=_sep, engine='python', nrows=5)
        print(f"  ✅ Survey CSV found ({len(_test_csv.columns)} columns detected, delimiter={repr(_sep)})")
        _test_upper = {str(c).strip().upper() for c in _test_csv.columns}
        if '#' in _test_upper and 'SUBMIT DATE (UTC)' in _test_upper:
            print("  ✅ Survey CSV: raw Typeform download — completion-time check enabled")
        else:
            print("  ✅ Survey CSV: Google Sheet export — no timestamps, completion-time check skipped")
    except Exception as _e:
        _errors.append(f"Survey CSV could not be read: {_e}")

# 2. Check MTF CSV exists, is readable and complete (C1)
_MTF_COLS = ['Type', '8th', '10th', '12th', 'variable', 'timeline', 'Year']
_MTF_TYPES = ['Total', 'Male', 'Female']
_MTF_TIMELINES = ['lifetime', 'past-12-months', 'past-30-days']

def read_mtf(path):
    """C1: read the MTF norms file without ever losing its first data row.
    Accepts comma, semicolon or tab separated files, with or without a header
    row. A file whose first line is data (e.g. 'Total;16.8;29.6;48.6;Alcohol;
    lifetime;2025') used to be read with that line as the header, which silently
    dropped the Alcohol lifetime Total row."""
    with open(path, 'r', encoding='utf-8-sig', errors='replace') as _f:
        _lines = [l for l in _f.read().splitlines() if l.strip()]
    if not _lines:
        raise ValueError('the MTF file is empty')
    _first = _lines[0]
    _sep = max([',', ';', '\t'], key=_first.count)
    _cells = [c.strip().strip('"') for c in _first.split(_sep)]
    _looks_like_data = bool(_cells) and (_cells[0] in _MTF_TYPES or
                                         any(re.fullmatch(r'-?\d+(\.\d+)?', c) for c in _cells))
    if _looks_like_data:
        _df = pd.read_csv(path, sep=_sep, header=None, encoding='utf-8-sig', dtype=str)
        if _df.shape[1] != len(_MTF_COLS):
            raise ValueError(f"the MTF file has no header row and {_df.shape[1]} columns; expected "
                             f"{len(_MTF_COLS)} columns in this order: {', '.join(_MTF_COLS)}")
        _df.columns = _MTF_COLS
        logger.warning("C1: the MTF file has no header row; its first line was read as data (not as a header).")
    else:
        _df = pd.read_csv(path, sep=_sep, encoding='utf-8-sig', dtype=str)
        _df.columns = [str(c).strip() for c in _df.columns]
        if 'Type' not in _df.columns:
            _df = _df.rename(columns={_df.columns[0]: 'Type'})
    for _c in ['Type', 'variable', 'timeline']:
        if _c in _df.columns:
            _df[_c] = _df[_c].astype(str).str.strip()
    if 'Year' in _df.columns:
        _df['Year'] = pd.to_numeric(_df['Year'], errors='coerce').astype('Int64')
    for _g in ['8th', '10th', '12th']:
        if _g in _df.columns:
            # '-' (or a garbled dash) means "not measured by MTF"; keep it as text.
            _df[_g] = _df[_g].astype(str).str.strip().replace({'nan': '-', '': '-'})
            _num = pd.to_numeric(_df[_g], errors='coerce')
            _df[_g] = _num.where(_num.notna(), '-')
    return _df

def validate_mtf(df, year):
    """C1 / Part 4 check 1: every substance x timeline must have Total, Male
    and Female rows for the chosen year. Returns a list of problems."""
    problems = []
    missing_cols = [c for c in _MTF_COLS if c not in df.columns]
    if missing_cols:
        return [f"missing column(s) {missing_cols}; expected {', '.join(_MTF_COLS)}"]
    d = df[df['Year'].astype(str) == str(year)]
    if d.empty:
        return [f"no rows for year {year}"]
    for (sub, tl), grp in d.groupby(['variable', 'timeline']):
        have = set(grp['Type'])
        lack = [t for t in _MTF_TYPES if t not in have]
        if lack:
            problems.append(f"{sub} / {tl}: no {', '.join(lack)} row")
    for sub, grp in d.groupby('variable'):
        lack_tl = [tl for tl in _MTF_TIMELINES if tl not in set(grp['timeline'])
                   and not (tl == 'lifetime' and sub not in ('Alcohol', 'Marijuana'))]
        if lack_tl and sub in ('Alcohol', 'Marijuana'):
            problems.append(f"{sub}: no {', '.join(lack_tl)} rows")
    return problems

if not os.path.isfile(MTFfilepath):
    _errors.append(f"MTF CSV not found: {MTFfilepath}")
else:
    try:
        _test_mtf = read_mtf(MTFfilepath)
        if 'Year' not in _test_mtf.columns:
            _errors.append("MTF CSV is missing a 'Year' column — check the file format.")
        else:
            _avail = sorted(int(y) for y in _test_mtf['Year'].dropna().unique())
            if not MTFyear:
                MTFyear = str(max(_avail))          # Enter = the latest benchmark year
            elif MTFyear not in [str(y) for y in _avail]:
                _warnings.append(f"Benchmark year '{MTFyear}' not in file. Available: {_avail}. Will use {max(_avail)}.")
                MTFyear = str(max(_avail))
            # The Two-Year Trend Report runs this script only to clean the survey
            # data and passes a placeholder norms file; norms are not used there.
            if globals().get('SUPP_CLEANING_ONLY'):
                _mtf_problems = []
            elif 'Region' not in _test_mtf.columns:
                _mtf_problems = ["this is not the Asia and Europe benchmark file (no Region column). Use "
                                 "'Asia Europe benchmarks 2026.csv' (press Enter for the default)."]
            else:
                _mtf_problems = [f"{_r}: {_p}" for _r in ['Asia', 'Europe']
                                 for _p in (validate_mtf(_test_mtf[_test_mtf['Region'] == _r].drop(columns='Region'), MTFyear)
                                            if (_test_mtf['Region'] == _r).any() else ['no rows for this region'])]
            if _mtf_problems:
                _errors.append("MTF CSV is incomplete (every substance and timeline needs Total, Male and Female rows):\n      "
                               + "\n      ".join(_mtf_problems))
            else:
                print(f"  ✅ Asia and Europe benchmark file found and complete (year {MTFyear})")
    except Exception as _e:
        _errors.append(f"MTF CSV could not be read: {_e}")

# 3. Check docx template exists
if not os.path.isfile(docxpath):
    _errors.append(f"Template .docx not found: {docxpath}")
else:
    print(f"  ✅ Template .docx found")

# 4. Check save path directory exists and is writable
_save_dir = os.path.dirname(savepath)
if not _save_dir:
    _save_dir = '.'
if not os.path.isdir(_save_dir):
    _errors.append(f"Output directory does not exist: {_save_dir}")
elif not os.access(_save_dir, os.W_OK):
    _errors.append(f"Output directory is not writable: {_save_dir}")
else:
    if not savepath.endswith('.docx'):
        _warnings.append(f"Save path doesn't end in .docx — adding it automatically.")
        savepath = savepath + '.docx'
    print(f"  ✅ Output path OK: {savepath}")

# 5. Check grades are valid
_valid_grade_formats = [g for g in grades if re.match(r'^\d+(st|nd|rd|th)$', g.strip())]
if not grades:
    _errors.append("No grades specified. Enter grades like: 8th,9th,10th,11th,12th")
elif len(_valid_grade_formats) < len(grades):
    _warnings.append(f"Some grades look unusual: {grades}. Expected format like '8th','9th','10th'.")
else:
    print(f"  ✅ Grades: {grades}")

# 6. Print warnings (non-fatal)
for _w in _warnings:
    print(f"  ⚠️  WARNING: {_w}")

# 7. Abort on errors
if _errors:
    print("\n❌ ERRORS FOUND — please fix these before running:\n")
    for _err in _errors:
        print(f"   • {_err}")
    raise SystemExit("Fix the above errors and run the script again.")

print("--- ALL INPUTS VALID — starting analysis ---\n")


# In[2]:


# # MTFfilepath = r"/Users/jessie/vs-code/prevEd/MTF_stats.csv"
# MTFfilepath = r"/Users/jessie/vs-code/prevEd/MTF_stats_2024 - MTF_stats_2024.csv"
# # MTFfilepath = r"/Users/jessie/vs-code/prevEd/MTF_stats_2024 - dummy.csv"
# MTFyear = '2023'
# docxpath= r"/Users/jessie/vs-code/prevEd/Template Report and Appendices_fix2.docx"
# savepath =  "/Users/jessie/vs-code/prevEd/test_output/ISMoutput1.docx"
# surveryoutput = r"/Users/jessie/vs-code/prevEd/ISM SUPP Survey Fall 2025 Data - ISM Prevention Ed SUPP Fall 2025.csv"
# grades = [ '9th','10th', '11th', '12th']
# date_survey = '1/1/2025'
# School = 'International School Manila'


# In[ ]:





# In[3]:


# expectedcols = ['#',
#  'What is your sex? ',
#  'How old are you? ',
#  'What grade are you in?',
#  'How were your grades overall during the last school year?',
#  'If you drink, how old were you when you had your first *whole drink* of alcohol?',
#  'Where did you first use alcohol?',
#  'Other',
#  'How often do you *typically* have at least one whole drink of alcohol?',
#  'How many *whole drinks of alcohol*, if any, do you typically have when you drink?',
#  'Have you had at least 1 whole drink of alcohol during the *PAST 12 MONTHS*?',
#  'Felt the need or tried to cut down on your drinking or other drug use ',
#  'Thought you had a drinking or other drug problem ',
#  'Felt guilty about your drinking or other drug use',
#  'Needed a drink or other drug first thing in the morning',
#  'Used alcohol or other drugs while you were by yourself',
#  'Get drunk',
#  'Do something you later felt sorry for',
#  'Have a fight, argument, or bad feelings with a friend/person I am dating',
#  'Have a fight, argument, or bad feelings with a family member',
#  "Had a sexual experience that you later wished didn't happen",
#  'Miss or perform poorly at school, work, or an athletic event',
#  'Have trouble concentrating on homework, class work, or other important tasks',
#  'Feel really sick, vomit, or have a hangover',
#  'Drive a car or other vehicle after drinking',
#  'Ride as a passenger in a car or other vehicle after the driver had been drinking',
#  'Get in trouble at school or home',
#  'Hang out with people you think are a bad influence',
#  'Forget what happened while you were drinking (Black out) ',
#  'Pass out because of drinking',
#  'Get hurt or injured',
#  'Have something to drink right before or during school hours',
#  "Try other substances you wouldn't have tried otherwise",
#  'Drank more than you intended',
#  'During the *PAST* *30 DAYS*, how many days did you have at least one whole drink of alcohol? ',
#  'What was the greatest number of whole drinks of alcohol you consumed in a row in the *PAST 30 DAYS*? ',
#  'During the *PAST 30 DAYS, *how many days did you have five or more whole drinks of alcohol in a row? ',
#  'During the *PAST 30 DAYS*, how many times have you been or felt drunk from drinking alcohol? ',
#  'Where do you do most of your drinking?',
#  'Other.1',
#  'If you have ever used marijuana (cannabis, pot, weed, hash, edibles), how old were you the first time you used it? ',
#  'How often do you use marijuana (cannabis, pot, weed, hash, edibles)? ',
#  'Alcohol (at least one whole drink of wine, beer, malt beverage, wine cooler, liquor, mixed  drink) ',
#  'Cigarettes',
#  'Electronic vaporizers (e-cigarettes, e-cigs, vape, vape pens, Puff Bars, Elf Bar, JUUL, etc.) ',
#  'Smokeless tobacco (chew, snuff, plug, dip, ZYN) ',
#  'Marijuana (cannabis, weed, hash, edibles, vaped cannabis, etc.)',
#  'Ecstasy (Molly, E, MDMA) ',
#  'Heroin',
#  'Narcotics NOT prescribed for you (OxyContin, Vicodin, Dilaudid) ',
#  'Amphetamines NOT prescribed for you (Ritalin, Concerta, Dexedrine, Adderall, diet pills, etc) ',
#  'Sedatives or Sleeping Pills NOT prescribed for you  (Valium, Ativan, Xanax, Klonopin, Librium, Rohypnol, GHB, etc.)',
#  'Methamphetamine (speed, crystal meth, ice, etc.) ',
#  'Cocaine',
#  'LSD (acid)',
#  'Hallucinogens other than LSD (psilocybin, mushrooms, shrooms peyote, mescaline, etc.)',
#  'Inhalants (poppers, nitrous, glue, gas, paint thinner, etc.) ',
#  'Steroids NOT prescribed for you ',
#  'Over-the-counter cold medicines used for non-medicinal purposes (NyQuil, cough syrup, etc.) ',
#  'Alcohol (at least one whole drink of wine, beer, malt beverage, wine cooler, liquor, mixed  drink) .1',
#  'Cigarettes.1',
#  'Electronic vaporizers (e-cigarettes, e-cigs, vape, vape pens, Puff Bars, Elf Bar, JUUL etc.) ',
#  'Smokeless tobacco (chew, snuff, plug, dip, ZYN) .1',
#  'Marijuana (cannabis, weed, hash, edibles, vaped cannabis, etc.).1',
#  'Ecstasy (Molly, E, MDMA) .1',
#  'Heroin.1',
#  'Narcotics NOT prescribed for you (OxyContin, Vicodin, Dilaudid) .1',
#  'Amphetamines NOT prescribed for you (Ritalin, Concerta, Dexedrine, Adderall, diet pills, etc) .1',
#  'Sedatives or Sleeping Pills NOT prescribed for you  (Valium, Ativan, Xanax, Klonopin, Librium, Rohypnol, GHB, etc.).1',
#  'Methamphetamine (speed, crystal meth, ice, etc.) .1',
#  'Cocaine.1',
#  'LSD (acid).1',
#  'Hallucinogens other than LSD (psilocybin, mushrooms, shrooms peyote, mescaline, etc.).1',
#  'Inhalants (poppers, nitrous, glue, gas, paint thinner, etc.) .1',
#  'Steroids NOT prescribed for you .1',
#  'Over-the-counter cold medicines used for non-medicinal purposes (NyQuil, cough syrup, etc.) .1',
#  'Talking to someone because you were concerned about their drinking or other drug use',
#  'Taking care of someone who drank too much',
#  'Taking care of someone who consumed too much marijuana (cannabis, weed, pot, etc.)',
#  'Taking someone for emergency medical care',
#  'Getting in trouble',
#  'Feeling insulted or humiliated',
#  'Engaging in a verbal argument',
#  'Being pushed, hit, assaulted, or involved in a physical fight',
#  "Had a sexual experience that you later wished didn't happen.1",
#  'Witnessing someone vomiting in front of you',
#  'Feeling unsafe',
#  'Having difficulty sleeping',
#  'Having difficulty doing homework',
#  'Have you ever used alcohol, marijuana, or other drugs or medications not prescribed to you before coming to, or during a school dance, school sporting event, or another school-sponsored event? \nPlease *DO NOT* include tobacco/nicotine.',
#  'Which of the following statements best represents *YOUR OWN* attitude towards teenagers drinking alcohol?',
#  'Had 1 or 2 drinks',
#  'Had 3 to 4 drinks',
#  'Had 5 or more drinks once or twice each weekend',
#  'Got drunk',
#  'Try one or two drinks of an alcoholic beverage (beer, wine, liquor) ',
#  'Have once or two drinks nearly every day',
#  'Have three or four drinks nearly every day ',
#  'Have five or more drinks once or twice each weekend',
#  'Once',
#  'Occasionally',
#  'Regularly',
#  'The social atmosphere at this school encourages alcohol use',
#  'Students at this school feel pressured to drink alcohol',
#  'It is easy to make friends at this school without drinking alcohol',
#  'Students assume that everyone at this school drinks alcohol',
#  'Students at this school will let you make your own decisions about drinking alcohol',
#  'Students at this school respect you if you do not drink alcohol',
#  'Students at this school who abstain from alcohol do not fit in',
#  "Most students at this school don't drink or are light drinkers",
#  'Students at this school think that heavy or binge drinking is okay',
#  'With a friend',
#  'With a parent or guardian',
#  'With a teacher or other adult at school',
#  'Students in your grade',
#  'Seniors (12th graders)',
#  'Students in your grade.1',
#  'Seniors (12th graders).1',
#  'During the *PAST 30 DAYS*, how often do you estimate that students at your school consumed at least one whole drink of alcohol?',
#  'Students in your grade.2',
#  'Seniors (12th graders).2',
#  'Students in your grade.3',
#  'Seniors (12th graders).3',
#  'Students in your grade.4',
#  'Seniors (12th graders).4',
#  'What percentage of students in your school do you think have ever used alcohol, marijuana, or other drugs not prescribed for them (do NOT include tobacco/nicotine) before coming to, or during a school dance, school sporting event, or other school-sponsored events? \n\n_Select the number that most closely represents your estimate. _',
#  'Alcohol/other drug policies at your school ',
#  'Enforcement of alcohol/other drug policies at your school ',
#  'Alcohol/other drug policies at your school .1',
#  'Enforcement of alcohol/other drug policies at your school .1',
#  'I feel valued as a person at my school. ',
#  'Academics are important to me. ',
#  'I feel that teachers and school staff care about me as a student. ',
#  'My school is making a sincere effort to address alcohol and other drug use problems.',
#  'Most students care about what happens at this school. ',
#  'Academics are important to other students at this school. ',
#  'I have at least one adult I trust at my school.',
#  'Which of the following statements best describes how you felt overall about going to school during the *PAST 12 MONTHS*? ',
#  'My parent/guardian show interest in my schoolwork. ',
#  'The rules in my family are clear. ',
#  "My parent/guardian wants me to call if I'm going to be late getting home.",
#  'My parent/guardian would know if I did not come home on time. ',
#  'My family has clear rules about alcohol and drug misuse. ',
#  'When I am not at home,  my parent/guardian knows how to contact me in an emergency.',
#  'Have you ever used alcohol, marijuana, or other drugs not prescribed for you at your home *WITH* your parent or guardian knowing? (Not including a few sips of wine with family or for religious purposes.) ',
#  'Have you ever used alcohol, marijuana, or other drugs not prescribed for you at your home *WITHOUT* your parent or guardian knowing? ',
#  'Please indicate how often your family eats at least one meal together. \n',
#  'Do you have a friend that you think has a problem with alcohol or other drug use? \n',
#  'Would you be comfortable talking to your parent/guardian about* your own *use/abuse of alcohol or other drugs? ',
#  'Would you be comfortable talking to your parent/guardian *about a friend* of yours who you think has an alcohol or other drug problem? ',
#  'My classmates feel pressured to use alcohol when they participate in non-school social events and activities with their peers.  ',
#  'When my classmates host parties or other social gatherings where alcohol or  other drugs are used, parents are present and visible.',
#  'Academic',
#  'Parental pressure',
#  'College expectations',
#  'Being overlooked',
#  'My own expectations',
#  'Boredom',
#  'Fitting in socially',
#  'Alcohol (at least one whole drink of wine, beer, malt beverage, wine cooler, liquor, mixed drink) ',
#  'Cigarettes.2',
#  'Electronic vaporizers (e-cigarettes, e-cigs, vape, vape pens, Puff Bars, Elf Bars, JUUL etc.)',
#  'Smokeless tobacco (chew, snuff, plug, dip, ZYN) .2',
#  'Marijuana (cannabis, weed, hash, edibles, vaped cannabis, etc.).2',
#  'Ecstasy (Molly, E, MDMA) .2',
#  'Heroin.2',
#  'Narcotics NOT prescribed for you (OxyContin, Vicodin, Dilaudid) .2',
#  'Amphetamines NOT prescribed for you (Ritalin, Concerta, Dexedrine, Adderall, diet pills, etc.) ',
#  'Sedatives or Sleeping Pills NOT prescribed for you  (Valium, Ativan, Xanax, Klonopin, Librium, Rohypnol, GHB, etc.).2',
#  'Methamphetamine (speed, crystal meth, ice, etc.) .2',
#  'Cocaine.2',
#  'LSD (acid) ',
#  'Hallucinogens other than LSD (psilocybin, mushrooms, shrooms peyote, mescaline, etc.).2',
#  'Inhalants (poppers, nitrous, glue, gas, paint thinner, etc.) .2',
#  'Steroids NOT prescribed for you .2',
#  'Over-the-counter cold medicines used for non-medicinal purposes (NyQuil, cough syrup, etc.) .2',
#  'In the *past 4 weeks*, about how often did you feel tired out for no good reason?',
#  'In the *past 4 weeks*, about how often did you feel nervous?',
#  'In the *past 4 weeks*, about how often did you feel so nervous that nothing could calm you down?',
#  'In the *past 4 weeks*, about how often did you feel hopeless?',
#  'In the *past 4 weeks*, about how often did you feel restless or fidgety?',
#  'In the *past 4 weeks*, about how often did you feel so restless you could not sit still?',
#  'In the *past 4 weeks*, about how often did you feel depressed?',
#  'In the *past 4 weeks*, about how often did you feel that everything was an effort?',
#  'In the *past 4 weeks*, about how often did you feel so sad that nothing could cheer you up?',
#  'In the *past 4 weeks*, about how often did you feel worthless?',
#  'Response Type',
#  'Start Date (UTC)',
#  'Stage Date (UTC)',
#  'Submit Date (UTC)',
#  'Network ID',
#  'Tags']
expectedcols = [
 'What is your sex? ',
 'How old are you? ',
 'What grade are you in?',
 'How were your grades overall during the last school year?',
 'If you drink, how old were you when you had your first *whole drink* of alcohol?',
 'Where did you first use alcohol?',
 'Other',
 'How often do you *typically* have at least one whole drink of alcohol?',
 'How many *whole drinks of alcohol*, if any, do you typically have when you drink?',
 'Have you had at least 1 whole drink of alcohol during the *PAST 12 MONTHS*?',
 'Felt the need or tried to cut down on your drinking or other drug use ',
 'Thought you had a drinking or other drug problem ',
 'Felt guilty about your drinking or other drug use',
 'Needed a drink or other drug first thing in the morning',
 'Used alcohol or other drugs while you were by yourself',
 'Get drunk',
 'Do something you later felt sorry for',
 'Have a fight, argument, or bad feelings with a friend/person I am dating',
 'Have a fight, argument, or bad feelings with a family member',
 "Had a sexual experience that you later wished didn't happen",
 'Miss or perform poorly at school, work, or an athletic event',
 'Have trouble concentrating on homework, class work, or other important tasks',
 'Feel really sick, vomit, or have a hangover',
 'Drive a car or other vehicle after drinking',
 'Ride as a passenger in a car or other vehicle after the driver had been drinking',
 'Get in trouble at school or home',
 'Hang out with people you think are a bad influence',
 'Forget what happened while you were drinking (Black out) ',
 'Pass out because of drinking',
 'Get hurt or injured',
 'Have something to drink right before or during school hours',
 "Try other substances you wouldn't have tried otherwise",
 'Drank more than you intended',
 'During the *PAST* *30 DAYS*, how many days did you have at least one whole drink of alcohol? ',
 'What was the greatest number of whole drinks of alcohol you consumed in a row in the *PAST 30 DAYS*? ',
 'During the *PAST 30 DAYS, *how many days did you have five or more whole drinks of alcohol in a row? ',
 'During the *PAST 30 DAYS*, how many times have you been or felt drunk from drinking alcohol? ',
 'Where do you do most of your drinking?',
 'Other.1',
 'If you have ever used marijuana (cannabis, pot, weed, hash, edibles), how old were you the first time you used it? ',
 'How often do you use marijuana (cannabis, pot, weed, hash, edibles)? ',
 'Alcohol (at least one whole drink of wine, beer, malt beverage, wine cooler, liquor, mixed  drink) ',
 'Cigarettes',
 'Electronic vaporizers (e-cigarettes, e-cigs, vape, vape pens, Puff Bars, Elf Bar, JUUL, etc.) ',
 'Smokeless tobacco (chew, snuff, plug, dip, ZYN) ',
 'Marijuana (cannabis, weed, hash, edibles, vaped cannabis, etc.)',
 'Ecstasy (Molly, E, MDMA) ',
 'Heroin',
 'Narcotics NOT prescribed for you (OxyContin, Vicodin, Dilaudid) ',
 'Amphetamines NOT prescribed for you (Ritalin, Concerta, Dexedrine, Adderall, diet pills, etc) ',
 'Sedatives or Sleeping Pills NOT prescribed for you  (Valium, Ativan, Xanax, Klonopin, Librium, Rohypnol, GHB, etc.)',
 'Methamphetamine (speed, crystal meth, ice, etc.) ',
 'Cocaine',
 'LSD (acid)',
 'Hallucinogens other than LSD (psilocybin, mushrooms, shrooms peyote, mescaline, etc.)',
 'Inhalants (poppers, nitrous, glue, gas, paint thinner, etc.) ',
 'Steroids NOT prescribed for you ',
 'Over-the-counter cold medicines used for non-medicinal purposes (NyQuil, cough syrup, etc.) ',
 'Alcohol (at least one whole drink of wine, beer, malt beverage, wine cooler, liquor, mixed  drink) .1',
 'Cigarettes.1',
 'Electronic vaporizers (e-cigarettes, e-cigs, vape, vape pens, Puff Bars, Elf Bar, JUUL etc.) ',
 'Smokeless tobacco (chew, snuff, plug, dip, ZYN) .1',
 'Marijuana (cannabis, weed, hash, edibles, vaped cannabis, etc.).1',
 'Ecstasy (Molly, E, MDMA) .1',
 'Heroin.1',
 'Narcotics NOT prescribed for you (OxyContin, Vicodin, Dilaudid) .1',
 'Amphetamines NOT prescribed for you (Ritalin, Concerta, Dexedrine, Adderall, diet pills, etc) .1',
 'Sedatives or Sleeping Pills NOT prescribed for you  (Valium, Ativan, Xanax, Klonopin, Librium, Rohypnol, GHB, etc.).1',
 'Methamphetamine (speed, crystal meth, ice, etc.) .1',
 'Cocaine.1',
 'LSD (acid).1',
 'Hallucinogens other than LSD (psilocybin, mushrooms, shrooms peyote, mescaline, etc.).1',
 'Inhalants (poppers, nitrous, glue, gas, paint thinner, etc.) .1',
 'Steroids NOT prescribed for you .1',
 'Over-the-counter cold medicines used for non-medicinal purposes (NyQuil, cough syrup, etc.) .1',
 'Talking to someone because you were concerned about their drinking or other drug use',
 'Taking care of someone who drank too much',
 'Taking care of someone who consumed too much marijuana (cannabis, weed, pot, etc.)',
 'Taking someone for emergency medical care',
 'Getting in trouble',
 'Feeling insulted or humiliated',
 'Engaging in a verbal argument',
 'Being pushed, hit, assaulted, or involved in a physical fight',
 "Had a sexual experience that you later wished didn't happen.1",
 'Witnessing someone vomiting in front of you',
 'Feeling unsafe',
 'Having difficulty sleeping',
 'Having difficulty doing homework',
 'Have you ever used alcohol, marijuana, or other drugs or medications not prescribed to you before coming to, or during a school dance, school sporting event, or another school-sponsored event? \nPlease *DO NOT* include tobacco/nicotine.',
 'Which of the following statements best represents *YOUR OWN* attitude towards teenagers drinking alcohol?',
 'Had 1 or 2 drinks',
 'Had 3 to 4 drinks',
 'Had 5 or more drinks once or twice each weekend',
 'Got drunk',
 'Try one or two drinks of an alcoholic beverage (beer, wine, liquor) ',
 'Have once or two drinks nearly every day',
 'Have three or four drinks nearly every day ',
 'Have five or more drinks once or twice each weekend',
 'Once',
 'Occasionally',
 'Regularly',
 'The social atmosphere at this school encourages alcohol use',
 'Students at this school feel pressured to drink alcohol',
 'It is easy to make friends at this school without drinking alcohol',
 'Students assume that everyone at this school drinks alcohol',
 'Students at this school will let you make your own decisions about drinking alcohol',
 'Students at this school respect you if you do not drink alcohol',
 'Students at this school who abstain from alcohol do not fit in',
 "Most students at this school don't drink or are light drinkers",
 'Students at this school think that heavy or binge drinking is okay',
 'With a friend',
 'With a parent or guardian',
 'With a teacher or other adult at school',
 'Students in your grade',
 'Seniors (12th graders)',
 'Students in your grade.1',
 'Seniors (12th graders).1',
 'During the *PAST 30 DAYS*, how often do you estimate that students at your school consumed at least one whole drink of alcohol?',
 'Students in your grade.2',
 'Seniors (12th graders).2',
 'Students in your grade.3',
 'Seniors (12th graders).3',
 'Students in your grade.4',
 'Seniors (12th graders).4',
 'What percentage of students in your school do you think have ever used alcohol, marijuana, or other drugs not prescribed for them (do NOT include tobacco/nicotine) before coming to, or during a school dance, school sporting event, or other school-sponsored events? \n\n_Select the number that most closely represents your estimate. _',
 'Alcohol/other drug policies at your school ',
 'Enforcement of alcohol/other drug policies at your school ',
 'Alcohol/other drug policies at your school .1',
 'Enforcement of alcohol/other drug policies at your school .1',
 'I feel valued as a person at my school. ',
 'Academics are important to me. ',
 'I feel that teachers and school staff care about me as a student. ',
 'My school is making a sincere effort to address alcohol and other drug use problems.',
 'Most students care about what happens at this school. ',
 'Academics are important to other students at this school. ',
 'I have at least one adult I trust at my school.',
 'Which of the following statements best describes how you felt overall about going to school during the *PAST 12 MONTHS*? ',
 'My parent/guardian show interest in my schoolwork. ',
 'The rules in my family are clear. ',
 "My parent/guardian wants me to call if I'm going to be late getting home.",
 'My parent/guardian would know if I did not come home on time. ',
 'My family has clear rules about alcohol and drug misuse. ',
 'When I am not at home,  my parent/guardian knows how to contact me in an emergency.',
 'Have you ever used alcohol, marijuana, or other drugs not prescribed for you at your home *WITH* your parent or guardian knowing? (Not including a few sips of wine with family or for religious purposes.) ',
 'Have you ever used alcohol, marijuana, or other drugs not prescribed for you at your home *WITHOUT* your parent or guardian knowing? ',
 'Please indicate how often your family eats at least one meal together. \n',
 'Do you have a friend that you think has a problem with alcohol or other drug use? \n',
 'Would you be comfortable talking to your parent/guardian about* your own *use/abuse of alcohol or other drugs? ',
 'Would you be comfortable talking to your parent/guardian *about a friend* of yours who you think has an alcohol or other drug problem? ',
 'My classmates feel pressured to use alcohol when they participate in non-school social events and activities with their peers.  ',
 'When my classmates host parties or other social gatherings where alcohol or  other drugs are used, parents are present and visible.',
 'Academic',
 'Parental pressure',
 'College expectations',
 'Being overlooked',
 'My own expectations',
 'Boredom',
 'Fitting in socially',
 'Alcohol (at least one whole drink of wine, beer, malt beverage, wine cooler, liquor, mixed drink) ',
 'Cigarettes.2',
 'Electronic vaporizers (e-cigarettes, e-cigs, vape, vape pens, Puff Bars, Elf Bars, JUUL etc.)',
 'Smokeless tobacco (chew, snuff, plug, dip, ZYN) .2',
 'Marijuana (cannabis, weed, hash, edibles, vaped cannabis, etc.).2',
 'Ecstasy (Molly, E, MDMA) .2',
 'Heroin.2',
 'Narcotics NOT prescribed for you (OxyContin, Vicodin, Dilaudid) .2',
 'Amphetamines NOT prescribed for you (Ritalin, Concerta, Dexedrine, Adderall, diet pills, etc.) ',
 'Sedatives or Sleeping Pills NOT prescribed for you  (Valium, Ativan, Xanax, Klonopin, Librium, Rohypnol, GHB, etc.).2',
 'Methamphetamine (speed, crystal meth, ice, etc.) .2',
 'Cocaine.2',
 'LSD (acid) ',
 'Hallucinogens other than LSD (psilocybin, mushrooms, shrooms peyote, mescaline, etc.).2',
 'Inhalants (poppers, nitrous, glue, gas, paint thinner, etc.) .2',
 'Steroids NOT prescribed for you .2',
 'Over-the-counter cold medicines used for non-medicinal purposes (NyQuil, cough syrup, etc.) .2',
 'In the *past 4 weeks*, about how often did you feel tired out for no good reason?',
 'In the *past 4 weeks*, about how often did you feel nervous?',
 'In the *past 4 weeks*, about how often did you feel so nervous that nothing could calm you down?',
 'In the *past 4 weeks*, about how often did you feel hopeless?',
 'In the *past 4 weeks*, about how often did you feel restless or fidgety?',
 'In the *past 4 weeks*, about how often did you feel so restless you could not sit still?',
 'In the *past 4 weeks*, about how often did you feel depressed?',
 'In the *past 4 weeks*, about how often did you feel that everything was an effort?',
 'In the *past 4 weeks*, about how often did you feel so sad that nothing could cheer you up?',
 'In the *past 4 weeks*, about how often did you feel worthless?',
 'Submitted At', 'Token']

# In[ ]:





# In[ ]:





# In[ ]:





# In[4]:


#grades = ['7th', '8th', '9th','10th', '11th', '12th']
grades_n = [x+' (N)' for x in grades ]
grades_perc = [x+' (%)' for x in grades ]

has_8th = '8th' in grades

_FULL_GRADE_COLS = []
for _g in ['8th', '9th', '10th', '11th', '12th']:
    _FULL_GRADE_COLS += [f'{_g} (N)', f'{_g} (%)']

# Grade (N) only list — used by total_agree_disagree and similar summary functions
_FULL_GRADES_N = [f'{_g} (N)' for _g in ['8th', '9th', '10th', '11th', '12th']]

def _pad_grade_cols(df, label_col):
    """Ensure df always has all 5 grade column pairs in the correct positional order,
    padding missing grades so template row[N] indices are stable.
    The template expects 13 columns: label + 5 grade pairs (N,%) + 2 trailing pads.
    (N) columns are padded with 0 so arithmetic division-by-zero guards fire correctly.
    (%) columns are padded with '' so downstream early-exit checks return '' cleanly."""
    for col in _FULL_GRADE_COLS:
        if col not in df.columns:
            if col.endswith(' (N)'):
                df[col] = 0
            else:
                df[col] = ''
    # Add 2 trailing pad columns so template row[11] and row[12] don't error
    df['_pad11'] = ''
    df['_pad12'] = ''
    return df[[label_col] + _FULL_GRADE_COLS + ['_pad11', '_pad12']]


# collist = [ "id","sex","age","grade_level","grades_overall","age_at_first_drink","where_first_use_alcohol","where_first_use_alcohol_text","how_often_1_drink","how_many_drinks","one_drink_past_12","felt_or_tried_to_cut","drinking_or_drug_problem","felt_guilty_about_your_drinking","needed_a_drink_first_thing","used_alcohol_by_yourself","get_drunk","do_something_you_later_felt_sorry_for","friend_fight_argument_bad_feelsing","family_fight_argument_bad_feelsing","sexual_experience","perform_poorly_school_work","trouble_concentrating","sick_vomit_hangover","drive","ride","trouble_school_home","hang_out_bad_influence","forget_what_happened","pass_out","get_hurt","drink_before_school","try_other_substances","drank_more","one_drink_past_30_days","drinks_in_a_row_30_days","days_5_or_more_drinks","drunk_30_days","where_do_you_drink","where_do_you_drink_other","age_first_time_pot","how_often_pot","use_alcohol_12","use_cigarettes_12","use_vapes_12","use_smokeless_tobacco_12","use_pot_12","use_ecstasy_12","use_heroin_12","use_narcotics_12","use_amphetamines_12","use_sedatives_12","use_methamphetamine_12","use_cocaine_12","use_lsd_12","use_hallucinogen_12","use_inhalants_12","use_steroids_12","use_cold_medicine_12","use_alcohol_30","use_cigarettes_30","use_vapes_30","use_smokeless_tobacco_30","use_pot_30","use_ecstasy_30","use_heroin_30","use_narcotics_30","use_amphetamines_30","use_sedatives_30","use_methamphetamine_30","use_cocaine_30","use_lsd_30","use_hallucinogen_30","use_inhalants_30","use_steroids_30","use_cold_medicine_30","taking_concerned_other","taking_care_drank_too_much_other","taking_care_consumed_too_much_other","emergency_medical_care_other","in_trouble_other","feeling_insulted_other","verbal_argument_other","pushed_hit_other","sexual_experience_other","witness_vomit_other","unsafe_other","sleep_difficulty_other","homework_difficulty_other","school_dance_use","attitude","one_2_friend","three_4_friend","five_friend","got_drunk_friend","harm_1_2_drinks","harm_1_2_daily","harm_3_4_daily","harm_5","pot_friend_once","pot_friend_occasion","pot_friend_regular","social_atmosphere","pressured","easy_to_make_friends","students_assume","own_decisions","respect","fit_in","light_drinkers","binge_okay","discussion_friend","discussion_guardian","discussion_teacher","drink_how_often_your_grade","drink_how_often_seniors","how_many_your_grade","how_many_seniors","one_drink_30_days_your_school","one_drink_30_days_your_grade","one_drink_30_days_seniors","how_many_30_days_your_grade","how_many_30_days_your_seniors","marijuana_your_grade","marijuana_seniors","percent_ever_used","thoughts_alcohol_policies","thoughts_enforcement","thoughts_other_alcohol_policies","thoughts_other_enforcement","valued","academics_to_me","teachers","sincere_effort","care","academics_to_others","adult_trust","school_12","parent_interest","clear_rules","call","come_home","clear_rules_alcohol","not_home","use_parent_knowing","use_parent_not_knowing","meals","friend_problem","talking_parents","talking_parents_about_friend","classmates_pressured","parents_present_visible","academic_stress","parent_stress","college_stress","overlooked_stress","expectations_stress","boredom_stress","fitting_in_stress","coping_alcohol","coping_cigarettes","coping_vape","coping_smokeless_tobacco","coping_marijuana","coping_molly","coping_heroin","coping_narcotics","coping_amphetamines","coping_sedatives","coping_meth","coping_cocaine","coping_lsd","coping_hallucinogens","coping_inhalant","coping_steroid","coping_cold_med","tired","nervous","no_calm","hopless","restless","no_sit_still","depressed","effort","cheer","worthless"]
collist = [ "sex","age","grade_level","grades_overall","age_at_first_drink","where_first_use_alcohol",
           "how_often_1_drink","how_many_drinks","one_drink_past_12",
           "felt_or_tried_to_cut","drinking_or_drug_problem","felt_guilty_about_your_drinking",
           "needed_a_drink_first_thing","used_alcohol_by_yourself","get_drunk","do_something_you_later_felt_sorry_for",
           "friend_fight_argument_bad_feelsing","family_fight_argument_bad_feelsing","sexual_experience",
           "perform_poorly_school_work","trouble_concentrating","sick_vomit_hangover","drive","ride","trouble_school_home",
           "hang_out_bad_influence","forget_what_happened","pass_out","get_hurt","drink_before_school","try_other_substances",
           "drank_more","one_drink_past_30_days","drinks_in_a_row_30_days","days_5_or_more_drinks","drunk_30_days",
           "where_do_you_drink","age_first_time_pot","how_often_pot","use_alcohol_12",
           "use_cigarettes_12","use_vapes_12","use_smokeless_tobacco_12","use_pot_12","use_ecstasy_12","use_heroin_12",
           "use_narcotics_12","use_amphetamines_12","use_sedatives_12","use_methamphetamine_12","use_cocaine_12",
           "use_lsd_12","use_hallucinogen_12","use_inhalants_12","use_steroids_12","use_cold_medicine_12","use_alcohol_30",
           "use_cigarettes_30","use_vapes_30","use_smokeless_tobacco_30","use_pot_30","use_ecstasy_30","use_heroin_30",
           "use_narcotics_30","use_amphetamines_30","use_sedatives_30","use_methamphetamine_30","use_cocaine_30","use_lsd_30",
           "use_hallucinogen_30","use_inhalants_30","use_steroids_30","use_cold_medicine_30","taking_concerned_other",
           "taking_care_drank_too_much_other","taking_care_consumed_too_much_other","emergency_medical_care_other",
           "in_trouble_other","feeling_insulted_other","verbal_argument_other","pushed_hit_other","sexual_experience_other",
           "witness_vomit_other","unsafe_other","sleep_difficulty_other","homework_difficulty_other","school_dance_use",
           "attitude","one_2_friend","three_4_friend","five_friend","got_drunk_friend","harm_1_2_drinks","harm_1_2_daily",
           "harm_3_4_daily","harm_5","pot_friend_once","pot_friend_occasion","pot_friend_regular","social_atmosphere",
           "pressured","easy_to_make_friends","students_assume","own_decisions","respect","fit_in","light_drinkers",
           "binge_okay","discussion_friend","discussion_guardian","discussion_teacher","drink_how_often_your_grade",
           "drink_how_often_seniors","how_many_your_grade","how_many_seniors","one_drink_30_days_your_school",
           "one_drink_30_days_your_grade","one_drink_30_days_seniors","how_many_30_days_your_grade",
           "how_many_30_days_your_seniors","marijuana_your_grade","marijuana_seniors","percent_ever_used",
           "thoughts_alcohol_policies","thoughts_enforcement","thoughts_other_alcohol_policies","thoughts_other_enforcement",
           "valued","academics_to_me","teachers","sincere_effort","care","academics_to_others","adult_trust","school_12",
           "parent_interest","clear_rules","call","come_home","clear_rules_alcohol","not_home","use_parent_knowing",
           "use_parent_not_knowing","meals","friend_problem","talking_parents","talking_parents_about_friend",
           "classmates_pressured","parents_present_visible"]


### to change the columns to this later if there are the right number
#no adult trust


# In[5]:


gradenums = [x.split('th')[0] for x in grades ]


# In[6]:


prettygrades = f'{", ".join(gradenums[:-1])} and {gradenums[-1]}'


# In[7]:


##############################
### Files
##############################

def _detect_csv_delimiter(filepath, encoding='latin-1', sample_lines=10):
    """Detect CSV delimiter from file (comma, semicolon, or tab)."""
    with open(filepath, 'r', encoding=encoding, errors='replace') as f:
        sample = ''.join(f.readline() for _ in range(sample_lines))
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=',;\t')
        return dialect.delimiter
    except csv.Error:
        # Fallback: use whichever delimiter appears most in first line (header)
        first_line = sample.split('\n')[0] if sample else ''
        if first_line.count(';') > first_line.count(',') and first_line.count(';') > first_line.count('\t'):
            return ';'
        if first_line.count('\t') > first_line.count(',') and first_line.count('\t') > first_line.count(';'):
            return '\t'
        return ','

_csv_delimiter = _detect_csv_delimiter(surveryoutput)
logger.info(f"Survey CSV delimiter detected: {repr(_csv_delimiter)}")
# engine='python' tolerates inconsistent field counts (fills with NaN); on_bad_lines skips/warns on malformed rows
try:
    clean = pd.read_csv(surveryoutput, encoding='latin-1', sep=_csv_delimiter, engine='python', on_bad_lines='warn')
except TypeError:
    # pandas < 1.3
    clean = pd.read_csv(surveryoutput, encoding='latin-1', sep=_csv_delimiter, engine='python', error_bad_lines=False, warn_bad_lines=True)

# CORRECTED: Normalize all fancy/curly apostrophes to standard apostrophe.
# The CSV contains \u2019 (right single quote) encoded as UTF-8 bytes that become
# multi-byte mojibake (\xc3\xa2\xc2\x80\xc2\x99 or \xe2\x80\x99) under latin-1 reading.
# This affects column names AND cell values. Normalize them all to standard "'".
_apostrophe_variants = [
    '\u2019',       # right single quotation mark (if read as utf-8)
    '\u2018',       # left single quotation mark
    '\xc2\x92',     # windows-1252 right single quote read as latin-1
    '\xe2\x80\x99', # UTF-8 bytes for \u2019 read as latin-1 (â€™)
    '\xe2\x80\x98', # UTF-8 bytes for \u2018 read as latin-1 (â€˜)
]
for _apo in _apostrophe_variants:
    clean.columns = [c.replace(_apo, "'") for c in clean.columns]
# Also normalize cell values in string columns
for _col in clean.select_dtypes(include='object').columns:
    for _apo in _apostrophe_variants:
        clean[_col] = clean[_col].str.replace(_apo, "'", regex=False)
logger.info("Apostrophe normalization complete for column names and cell values.")

# Accept a raw Typeform download as well as the Google Sheet export.
# Raw downloads have '#' first and 7 metadata columns at the end, and no
# 'Submitted At'/'Token'; the question columns are identical in both.
_TF_META = {'#', 'RESPONSE TYPE', 'START DATE (UTC)', 'STAGE DATE (UTC)',
            'SUBMIT DATE (UTC)', 'NETWORK ID', 'TAGS', 'ENDING'}
_tf_cols = {str(c).strip().upper(): c for c in clean.columns}
_completion_minutes = None
if '#' in _tf_cols and 'SUBMIT DATE (UTC)' in _tf_cols:
    logger.info("Raw Typeform export detected - converting to the Google Sheet layout.")
    if 'START DATE (UTC)' in _tf_cols:
        _tf_start = pd.to_datetime(clean[_tf_cols['START DATE (UTC)']], errors='coerce')
        _tf_submit = pd.to_datetime(clean[_tf_cols['SUBMIT DATE (UTC)']], errors='coerce')
        _completion_minutes = (_tf_submit - _tf_start).dt.total_seconds() / 60
    clean['Submitted At'] = clean[_tf_cols['SUBMIT DATE (UTC)']]
    clean['Token'] = clean[_tf_cols['#']]   # response number stands in for the Token id
    clean = clean.drop(columns=[c for c in clean.columns
                                if str(c).strip().upper() in _TF_META])
    clean = clean[[c for c in clean.columns if c not in ('Submitted At', 'Token')]
                  + ['Submitted At', 'Token']]

MTF = read_mtf(MTFfilepath)   # C1: same reader as the startup check (never drops the first row)
# Nov 26 update: one norms table per region. MTF holds Asia so the existing
# calculations still run; every comparison shown in the report uses BENCH.
BENCH = {}
if 'Region' in MTF.columns:
    for _r in ['Asia', 'Europe']:
        _b = MTF[MTF['Region'] == _r].drop(columns='Region').reset_index(drop=True)
        if len(_b):
            BENCH[_r] = _b
    MTF = BENCH['Asia'].copy()
BENCH_NOTE = ('Asia and Europe averages are from Prevention Ed SUPP surveys of international schools in each region, '
              'combined across schools and weighted by the number of students.')
# Filter to requested MTF year (use closest available if exact not found)
available_years = sorted(int(y) for y in MTF['Year'].dropna().unique())
if MTFyear and MTFyear in [str(y) for y in available_years]:
    MTF = MTF[MTF['Year'].astype(str) == MTFyear]
else:
    closest = max(available_years)
    print(f"MTF year {MTFyear} not found. Using {closest} (latest available).")
    MTF = MTF[MTF['Year'] == closest]
# (already renamed and filtered above — no further action needed)



# In[9]:


if "Other" in clean.columns:
    clean['Other'].value_counts(dropna=False)
else:
    print("column 'Other' not found")


# In[ ]:





# In[8]:


expectedcols = [x.strip().replace('*', '').upper() for x in expectedcols]
cleancols = [x.strip().replace('*', '').upper() for x in clean.columns]


mismatchedcols = list(set(cleancols) - set(expectedcols))
print(f"{len(mismatchedcols)} missing or unmatched columns found compared to expected")
if len(mismatchedcols) > 0:
    print(f"  Unmatched columns: {mismatchedcols}")


# In[ ]:





# In[ ]:





# In[9]:



# clean.columns = collist + ['Response Type', 'Start Date (UTC)',  'Stage Date','Submit Date (UTC)', 'Network ID', 'Tags']
# Typeform "Other (please specify)" text boxes arrive as columns named 'Other',
# 'Other.1', ... They are not in collist, so they must be dropped before the
# positional rename below, or every later column is shifted by one.
_other_text_cols = [c for c in clean.columns if re.fullmatch(r'Other(\.\d+)?', str(c).strip())]
# R1c needs the "where did you first use alcohol: Other" free text (a student who
# gives a first-drink age but writes "I don't drink"). Keep it aside.
_first_use_other_text = clean[_other_text_cols[0]].copy() if _other_text_cols else None
if _other_text_cols:
    logger.info(f"Dropping {len(_other_text_cols)} free-text 'Other' column(s): {_other_text_cols}")
    clean = clean.drop(columns=_other_text_cols)

# Newer surveys add gambling and loot box questions. Find them by question text
# (not position), set them aside, and add them back after the rename.
_extra_q = {}
for _c in list(clean.columns):
    _lc = str(_c).lower()
    if 'bet or gamble' in _lc:
        _extra_q['gambling_30_days'] = clean.pop(_c)
    elif 'loot box' in _lc:
        _extra_q['lootbox_30_days'] = clean.pop(_c)

# Older surveys (e.g. 2023) did not ask some questions added later. Insert an
# empty column where a missing one belongs, so the positional rename lines up.
# Add to this list if another question is added to the survey.
_ADDED_LATER = {'adult_trust': 'adult i trust at my school'}
for _name, _words in _ADDED_LATER.items():
    if _name in collist and not any(_words in str(_c).lower() for _c in clean.columns):
        clean.insert(collist.index(_name), f'__missing_{_name}', np.nan)
        logger.warning(f"This survey did not ask the '{_name}' question; it is left blank.")

_raw_question_headers = list(clean.columns[:len(collist)])
try:
    n_extra = len(clean.columns) - len(collist)
    if n_extra == 2:
        clean.columns = collist + ['Submitted At', 'Token']
    elif n_extra == 3:
        clean.columns = collist + ['Submitted At', 'Token', 'agenum']
    else:
        raise ValueError(f"{n_extra} columns beyond collist (expected 2: Submitted At, Token)")
    # C19: the exact question text, from the survey file itself, keyed by
    # internal name. Quoted item wording in the report comes from here, never
    # from typed boilerplate ("drinking is okay", not "drinking is alright").
    QUESTION_TEXT = {k: re.sub(r'\s+', ' ', str(v).replace('*', '')).strip()
                     for k, v in zip(collist, _raw_question_headers)}
    for _name, _vals in _extra_q.items():
        clean[_name] = _vals
    clean['first_use_other_text'] = _first_use_other_text.values if _first_use_other_text is not None else np.nan
except ValueError as e:
    raise ValueError(f"Column count mismatch: CSV has {len(clean.columns)} columns but collist+extras has {len(collist)+2}. Check collist. Original error: {e}")

# Completion time from the raw Typeform timestamps (NaN for Google Sheet exports).
# Attached by index before any row drops, so alignment holds.
clean['completion_minutes'] = _completion_minutes if _completion_minutes is not None else np.nan


# In[ ]:





# In[10]:






##############################
### Remap British "Year" grades to US-style grade labels
##############################
_YEAR_TO_GRADE = {
    'Year 7': '7th', 'Year 8': '8th', 'Year 9': '8th', 'Year 10': '9th',
    'Year 11': '10th', 'Year 12': '11th', 'Year 13': '12th',
}
if clean['grade_level'].dropna().str.startswith('Year').any():
    clean['_was_year_system'] = True
    clean['grade_level'] = clean['grade_level'].map(lambda x: _YEAR_TO_GRADE.get(str(x).strip(), x))
    grades = [_YEAR_TO_GRADE.get(g, g) for g in grades]
    grades_n = [x + ' (N)' for x in grades]
    grades_perc = [x + ' (%)' for x in grades]
    gradenums = [x.replace('th', '') for x in grades]
    prettygrades = f'{", ".join(gradenums[:-1])} and {gradenums[-1]}'
    logger.info(f"Remapped British year groups to: {grades}")

_has_stress = 'academic_stress' in clean.columns
_has_coping = 'coping_alcohol' in clean.columns

##############################
### Process and clean
##############################


# clean['Start Date (UTC)'] = pd.to_datetime(clean['Start Date (UTC)'])
# clean['Submit Date (UTC)']= pd.to_datetime(clean['Submit Date (UTC)'])


# clean['Minutes to complete'] = (pd.to_datetime(clean['Submit Date (UTC)'] ) - pd.to_datetime(clean['Start Date (UTC)'] ) ).dt.total_seconds() / 60

# print(f"{len(clean[clean['Minutes to complete']<=3])} people completed the survey in less than 3 minutes and have been filtered out")

# print(f"Leaving a n={len(clean[clean['Minutes to complete']>3])}")

# clean = clean[clean['Minutes to complete']>3]



# # In[ ]:





# # In[ ]:





# # In[11]:





# def create_flag(s):
#     if (s['how_often_1_drink'] != 'Never') & (s['one_drink_past_12']=='No'):
#         return 'flag 1'
#     elif (s['how_often_1_drink'] == 'Never') & (s['one_drink_past_12']=='Yes'):
#         return 'flag 2'
#     elif (s['one_drink_past_30_days']!= "I did not drink alcohol in the past 30 days") & (s['drinks_in_a_row_30_days']== "0 drinks"):
#         return 'flag 3'
    
    
# clean['Flags'] = clean.apply(create_flag, axis=1)

# removed_N = len(clean[clean['Flags'].notna()])

# print('Checking whether people answered how often they have one drink and how often they drank in the last year align.')

# print(f"This many flags found: {removed_N}" )


# N = len(clean[clean['Flags'].isna()])
# print(f"Leaving a n={N}")
# clean =clean[clean['Flags'].isna()]






# ################# fix the apostrophes 


# fixcols = ['grades_overall',
#            'where_do_you_drink',
#            'attitude',
#            'thoughts_alcohol_policies',
# 'thoughts_enforcement',
# 'thoughts_other_alcohol_policies',
# 'thoughts_other_enforcement']

# for col in fixcols:
#     clean[col] = clean[col].str.replace("'", "'")

    
    
# clean['agenum'] = clean['age'].str.split(" ").str[0]

# clean['agenum'] = clean['agenum'].astype(float)

# N_removed_perc = round((removed_N/N)*100,1) if N > 0 else 0


# # In[12]:



# gradesincluded = clean['grade_level'].value_counts().reset_index()['index'].tolist()

# gradesmissing = list(set(gradesincluded) ^ set(['7th', '8th', '9th', '10th', '11th', '12th']))
# gradesmissing


# # In[13]:


# addemptyrows = pd.DataFrame({'grade_level': gradesmissing})


# # In[14]:


# clean =pd.concat([clean, addemptyrows])


N = len(clean)
# print(f"Leaving a n={N}")
# clean =clean[clean['Flags'].isna()]

# C2/C3: the cannabis age-vs-frequency contradiction (rule R1b) used to be removed
# here silently. All exclusion rules now run together in the R1 block below and
# every excluded response is written to the exclusion log.
removed_N = 0

# Reconcile the entered grade list with the grades actually present in the
# survey data. Grades present in the data but missing from the entered list
# would otherwise be silently excluded from every table (and can crash the
# 8th-grade code paths); grades entered with no respondents render as '-'.
_SUPPORTED_GRADES = ['7th', '8th', '9th', '10th', '11th', '12th']
_grades_in_survey = [g for g in _SUPPORTED_GRADES if g in set(clean['grade_level'].dropna())]
_missing_from_input = [g for g in _grades_in_survey if g not in grades]
if _missing_from_input:
    print(f"⚠️  WARNING: the survey data contains {_missing_from_input} respondents who were "
          f"not in the grade list you entered. Adding them so their answers are included.")
    grades = [g for g in _SUPPORTED_GRADES if g in grades or g in _missing_from_input] \
             + [g for g in grades if g not in _SUPPORTED_GRADES]
    grades_n = [x + ' (N)' for x in grades]
    grades_perc = [x + ' (%)' for x in grades]
    has_8th = '8th' in grades
    gradenums = [x.split('th')[0] for x in grades]
    prettygrades = f'{", ".join(gradenums[:-1])} and {gradenums[-1]}'
_unsupported_grades = sorted(set(clean['grade_level'].dropna()) - set(_SUPPORTED_GRADES))
if _unsupported_grades:
    print(f"⚠️  WARNING: {_unsupported_grades} respondents found in the data — this report "
          f"only supports grades 8th-12th, so they are excluded from grade-level tables.")
_empty_grades = [g for g in grades if g not in _grades_in_survey and g in _SUPPORTED_GRADES]
if _empty_grades:
    print(f"⚠️  NOTE: no survey responses for {_empty_grades} — their table columns will show '-'.")



# print("There was no start and end time in this data, just a submitted time so we cannot calculate how fast the survery was taken" )



fixcols = ['grades_overall',
           'where_do_you_drink',
           'attitude',
           'thoughts_alcohol_policies',
'thoughts_enforcement',
'thoughts_other_alcohol_policies',
'thoughts_other_enforcement']

for col in fixcols:
    if col in clean.columns:
        # B2/B7: astype(str) turned blank answers into the text 'nan', which was
        # then counted as an answer ("do most of their drinking nan") and put
        # every student into the denominator. Only convert real answers.
        _notna = clean[col].notna()
        clean.loc[_notna, col] = clean.loc[_notna, col].astype(str).str.replace("'", "'")

# Typographic apostrophes/quotes arrive as UTF-8 byte sequences read as
# latin-1 (e.g. "Mostly A\xe2\x80\x99s" shows as "Mostly Aâ\x80\x99s"), which
# breaks answer-option matching and renders garbled text in tables. Normalize
# them to plain ASCII in every text column of the survey data.
_mojibake_map = {
    'â\x80\x99': "'",
    'â\x80\x98': "'",
    'â\x80\x9c': '"',
    'â\x80\x9d': '"',
    '’': "'",
    '‘': "'",
}
for _col in clean.columns:
    try:
        _s = clean[_col]
        for _bad, _good in _mojibake_map.items():
            _s = _s.str.replace(_bad, _good, regex=False)
        # Trailing/leading whitespace in Typeform answer labels (e.g.
        # "2 days ", "Most of the time ") silently zeroes those options in
        # every table keyed on an answer list — strip it everywhere.
        _s = _s.str.strip()
        clean[_col] = _s
    except (AttributeError, TypeError):
        pass  # non-string column

# B2: blank / '-' / 'nan' cells are "did not answer", never a response option.
# (Typeform writes '-' or '' for skipped follow-up questions.)
for _col in clean.columns:
    if clean[_col].dtype == object:
        clean[_col] = clean[_col].replace({'': np.nan, '-': np.nan, 'nan': np.nan, 'NaN': np.nan, 'None': np.nan})

# B7: Yes/No questions exported as True/False (or 1/0) printed "True" in the
# appendix and dropped rows. Normalize them all to 'Yes' / 'No'.
_BINARY_COLS = ['school_dance_use', 'use_parent_knowing', 'use_parent_not_knowing',
                'friend_problem', 'talking_parents', 'talking_parents_about_friend']
_YES_VALUES = {True, 1, '1', 'True', 'TRUE', 'true', 'Yes', 'yes', 'YES'}
_NO_VALUES = {False, 0, '0', 'False', 'FALSE', 'false', 'No', 'no', 'NO'}
for _col in _BINARY_COLS:
    if _col in clean.columns:
        clean[_col] = clean[_col].apply(
            lambda v: 'Yes' if v in _YES_VALUES else ('No' if v in _NO_VALUES else (np.nan if pd.isna(v) else v)))

##############################
### R1: automatic exclusion of invalid respondents (before any analysis)
##############################
_HARD_DRUGS = ['ecstasy', 'heroin', 'narcotics', 'amphetamines', 'sedatives', 'methamphetamine',
               'cocaine', 'lsd', 'hallucinogen', 'inhalants', 'steroids']
_hd12 = [f'use_{d}_12' for d in _HARD_DRUGS if f'use_{d}_12' in clean.columns]
_hd30 = [f'use_{d}_30' for d in _HARD_DRUGS if f'use_{d}_30' in clean.columns]
_never_drink = clean['age_at_first_drink'].eq('I have never had a whole drink of alcohol')
_never_pot = clean['age_first_time_pot'].eq('I have never used marijuana at all')
_drank30_followup = clean['one_drink_past_30_days'].notna() & ~clean['one_drink_past_30_days'].astype(str).str.startswith('I did not')

_flag_all_drugs = (clean[_hd12].eq('Yes').all(axis=1) | clean[_hd30].eq('Yes').all(axis=1)) if _hd12 else pd.Series(False, index=clean.index)
_flag_contradiction = (
    (_never_pot & (clean['use_pot_12'].eq('Yes') | clean['use_pot_30'].eq('Yes'))) |
    (_never_drink & (clean['one_drink_past_12'].eq('Yes') | _drank30_followup))
)
_flag_speeder = clean['completion_minutes'].notna() & (clean['completion_minutes'] < MIN_COMPLETION_MINUTES)
# R5: past-30-day hard-drug use with no past-12-month use of the same drug
_flag_time_order = pd.Series(False, index=clean.index)
for _d in _HARD_DRUGS:
    _c12, _c30 = f'use_{_d}_12', f'use_{_d}_30'
    if _c12 in clean.columns and _c30 in clean.columns:
        _flag_time_order |= clean[_c30].eq('Yes') & ~clean[_c12].eq('Yes')
# R1b (C3): age of first cannabis use given, but "Never" for frequency (or the reverse)
_flag_cannabis_age = (
    (_never_pot & (clean['how_often_pot'] != 'Never')) |
    ((clean['how_often_pot'] == 'Never') & ~_never_pot)
)
# R1c (C3): age of first drink is older than the student is now (e.g. "19 years or
# older" at 17), or a first-drink age was given with free text saying they do not drink.
def _band_low(v):
    _m = re.match(r'\s*(\d+)', str(v))
    return float(_m.group(1)) if _m else np.nan
_age_now = pd.to_numeric(clean['age'].astype(str).str.extract(r'(\d+)')[0], errors='coerce')
_first_low = clean['age_at_first_drink'].map(lambda v: np.nan if pd.isna(v) else _band_low(v))
_no_drink_text = clean['first_use_other_text'].astype(str).str.contains(
    r"(?i)\b(?:don'?t|do not|never|not)\s+(?:drink|had|have|tried)|\bno alcohol\b|\bnever drank\b", na=False)
_flag_first_drink_age = (~_never_drink & clean['age_at_first_drink'].notna() &
                         ((_first_low > _age_now) | _no_drink_text))
# R1d (C3): lifetime drinker who says "No" to a drink in the past 12 months but
# ticks alcohol in the 12-month substance list.
_flag_alcohol_12m = (~_never_drink & clean['age_at_first_drink'].notna() &
                     clean['one_drink_past_12'].eq('No') & clean['use_alcohol_12'].eq('Yes'))
# R1e (C3): mostly blank. Counts answered survey questions (not metadata).
_answered = clean[[c for c in collist if c in clean.columns]].notna().sum(axis=1)
_flag_mostly_blank = _answered < MIN_ANSWERED_ITEMS

# Named rules, in the order they are reported (C2/C3).
EXCLUSION_RULES = [
    ('R1a', 'ALL_DRUGS', 'reported every listed hard drug in the 12-month or 30-day list',
     _flag_all_drugs, EXCLUDE_ALL_DRUG_REPORTERS),
    ('R1b', 'CANNABIS_AGE', 'gave an age of first cannabis use but "Never" for cannabis frequency (or the reverse)',
     _flag_cannabis_age, True),
    ('R1c', 'FIRST_DRINK_AGE', 'age of first drink is older than their age, or wrote that they do not drink',
     _flag_first_drink_age, EXCLUDE_FIRST_DRINK_AGE),
    ('R1d', 'ALCOHOL_12M', 'lifetime drinker with no drink in the past 12 months who ticked alcohol in the 12-month list',
     _flag_alcohol_12m, EXCLUDE_ALCOHOL_12M_CONTRADICTIONS),
    ('R1e', 'MOSTLY_BLANK', f'answered fewer than {MIN_ANSWERED_ITEMS} questions',
     _flag_mostly_blank, EXCLUDE_MOSTLY_BLANK),
    ('R1f', 'LIFETIME_NEVER', 'past-year use contradicts their own "never used" lifetime answer',
     _flag_contradiction, EXCLUDE_LIFETIME_CONTRADICTIONS),
    ('R1g', 'SPEEDER', f'completed the survey in under {MIN_COMPLETION_MINUTES} minutes',
     _flag_speeder, EXCLUDE_SPEEDERS),
    ('R1h', 'TIME_ORDER', 'past-30-day hard-drug use but no past-12-month use of the same drug',
     _flag_time_order, EXCLUDE_TIME_ORDER_CONTRADICTIONS),
]
_exclude_mask = pd.Series(False, index=clean.index)
_any_flag = pd.Series(False, index=clean.index)
for _code, _name, _desc, _mask, _on in EXCLUSION_RULES:
    _any_flag |= _mask
    if _on:
        _exclude_mask |= _mask
INVALID_RESPONDENTS = clean[_any_flag].copy()
_inv_idx = INVALID_RESPONDENTS.index
def _rules_for(i, on_only=False):
    return [(c, n, d) for c, n, d, m, on in EXCLUSION_RULES if m[i] and (on or not on_only)]
INVALID_RESPONDENTS['reason'] = [
    '; '.join(f'{c} {n}: {d}' for c, n, d in _rules_for(i)) for i in _inv_idx]
INVALID_RESPONDENTS['excluded'] = _exclude_mask[_inv_idx]
# C2: the exclusion log. One row per excluded response, with its response ID.
EXCLUSION_LOG = pd.DataFrame([{
    'response_id': clean.at[i, 'Token'],
    'grade': clean.at[i, 'grade_level'], 'sex': clean.at[i, 'sex'],
    'rule_codes': '; '.join(c for c, _, _ in _rules_for(i, True)),
    'rules': '; '.join(f'{n}: {d}' for _, n, d in _rules_for(i, True)),
} for i in clean.index[_exclude_mask]], columns=['response_id', 'grade', 'sex', 'rule_codes', 'rules'])
EXCLUSION_COUNTS = {f'{c} {n}': int((m & _exclude_mask).sum()) for c, n, d, m, on in EXCLUSION_RULES if on}
if _exclude_mask.any():
    print(f"R1: excluding {int(_exclude_mask.sum())} respondent(s) as invalid: "
          + ', '.join(f'{k} = {v}' for k, v in EXCLUSION_COUNTS.items() if v)
          + " (a response can trigger more than one rule). See <report>_exclusion_log.csv.")
    removed_N += int(_exclude_mask.sum())
    clean = clean[~_exclude_mask].reset_index(drop=True)
# Text for the report (Reliability section): how many responses each rule removed.
_rule_bits = [f"{d} ({v})" for (c, n, d, m, on), v in
              zip([r for r in EXCLUSION_RULES if r[4]], EXCLUSION_COUNTS.values()) if v]
EXCLUSION_SUMMARY = ((f"The {int(_exclude_mask.sum())} excluded surveys were identified by these checks "
                      f"(a survey can fail more than one): " + '; '.join(_rule_bits) + '.')
                     if _exclude_mask.any() else 'No surveys failed the consistency checks.')
# R5: for all remaining students, a past-30-day "Yes" implies a past-12-month "Yes".
# Recode the 12-month answer so the 12-month tables, the 30-day tables and the
# derived flags (pot_12m, vape_12m, ...) all count the same students.
_TIME_RECODE_FLAGS = []
for _c30 in [c for c in clean.columns if re.fullmatch(r'use_[a-z_]+_30', c)]:
    _c12 = _c30[:-3] + '_12'
    if _c12 not in clean.columns:
        continue
    _fix = clean[_c30].eq('Yes') & ~clean[_c12].eq('Yes')
    for _i in clean.index[_fix]:
        _TIME_RECODE_FLAGS.append({'row': int(_i), 'grade': clean.at[_i, 'grade_level'], 'sex': clean.at[_i, 'sex'],
                                   'issue': f'{_c30} = Yes but {_c12} = {clean.at[_i, _c12]!r}; past-12-month answer set to Yes'})
    clean.loc[_fix, _c12] = 'Yes'
if _TIME_RECODE_FLAGS:
    print(f"R5: corrected {len(_TIME_RECODE_FLAGS)} past-12-month answer(s) where the student reported past-30-day use "
          f"(see _QA_flags.csv).")
if clean['completion_minutes'].isna().all():
    print("Completion-time check skipped: input has no Start/Submit timestamps (Google Sheet export).")

##############################
### B1: ONE canonical flag per behavior, used everywhere
### Drinking follows the drinking follow-up items (Q9 past 12 months,
### Q30-day days item). The substance-list items (use_alcohol_12/30) are
### only used to FLAG students whose two answers disagree.
##############################
clean['drank_30d'] = clean['one_drink_past_30_days'].notna() & ~clean['one_drink_past_30_days'].astype(str).str.startswith('I did not')
clean['drank_12m'] = clean['one_drink_past_12'].eq('Yes') | clean['drank_30d']
clean['drank_lifetime'] = (clean['age_at_first_drink'].notna() &
                           ~clean['age_at_first_drink'].eq('I have never had a whole drink of alcohol')) | clean['drank_12m']
# Cannabis: any past-year or past-30-day user counts as a lifetime user.
clean['pot_12m'] = clean['use_pot_12'].eq('Yes') | clean['use_pot_30'].eq('Yes')
clean['pot_30d'] = clean['use_pot_30'].eq('Yes')
clean['pot_lifetime'] = (clean['age_first_time_pot'].notna() &
                         ~clean['age_first_time_pot'].eq('I have never used marijuana at all')) | clean['pot_12m'] | \
                        (clean['how_often_pot'].notna() & ~clean['how_often_pot'].eq('Never'))
clean['vape_12m'] = clean['use_vapes_12'].eq('Yes') | clean['use_vapes_30'].eq('Yes')
clean['vape_30d'] = clean['use_vapes_30'].eq('Yes')

QA_FLAGS = []
QA_FLAGS.extend(_TIME_RECODE_FLAGS)
def _qa_flag(mask, reason):
    for _i in clean.index[mask]:
        QA_FLAGS.append({'row': int(_i), 'grade': clean.at[_i, 'grade_level'], 'sex': clean.at[_i, 'sex'], 'issue': reason})
_qa_flag(clean['drank_30d'] != clean['use_alcohol_30'].eq('Yes'),
         'past-30-day drinking: follow-up item and substance list disagree (follow-up used)')
_qa_flag(clean['drank_12m'] != clean['use_alcohol_12'].eq('Yes'),
         'past-12-month drinking: follow-up item and substance list disagree (follow-up used)')
_qa_flag(clean['drank_30d'] & clean['drinks_in_a_row_30_days'].eq('0 drinks'),
         'past-30-day drinker reports a maximum of 0 drinks in a row')
_qa_flag(clean['days_5_or_more_drinks'].notna() & ~clean['days_5_or_more_drinks'].astype(str).str.startswith('I did not')
         & clean['drinks_in_a_row_30_days'].isin(['0 drinks', '1-2 drinks', '3-4 drinks']),
         'reports 5+ drinks in a row on 1+ days but a smaller greatest-number-in-a-row')
_qa_flag(clean['use_pot_12'].eq('Yes') & clean['age_first_time_pot'].eq('I have never used marijuana at all'),
         'past-year cannabis use but "never used marijuana" (counted as lifetime user)')
print(f"B1: drinkers past 30 days = {int(clean['drank_30d'].sum())} (substance list says {int(clean['use_alcohol_30'].eq('Yes').sum())}); "
      f"past 12 months = {int(clean['drank_12m'].sum())} (substance list says {int(clean['use_alcohol_12'].eq('Yes').sum())}). "
      f"{len(QA_FLAGS)} data-quality flags written to the _QA_flags.csv file.")

if clean['age'].dtype != 'int64':
    
    clean['agenum'] = clean['age'].str.split(" ").str[0]
    clean['agenum'] = clean['agenum'].astype(float)
else: 
    clean['agenum'] = clean['age']

##because the timing flag isn't there, manually setting to zero
# removed_N = 0

N_removed_perc = round((removed_N/N)*100,1) if N > 0 else 0
# B7: report both numbers: surveys collected (N) and surveys analyzed.
N_analyzed = len(clean)
grades_in_data_sorted = [g for g in ['7th', '8th', '9th', '10th', '11th', '12th'] if g in set(clean['grade_level'].dropna())]
lowest_grade = grades_in_data_sorted[0] if grades_in_data_sorted else ''
_body_grades_present = [g for g in BODY_GROUP_GRADES if g in grades_in_data_sorted]

def grade_size(g, sex=None):
    m = clean['grade_level'].eq(g)
    if sex is not None:
        m &= clean['sex'].eq(sex)
    return int(m.sum())

def flag_rate(flag_col, grade=None, sex=None, among=None, name=None, base=None):
    """B1/B2: % of students (in a grade/sex, optionally among a sub-population)
    with a canonical flag set. Denominator = all students in that group."""
    m = pd.Series(True, index=clean.index)
    if grade is not None:
        m &= clean['grade_level'].eq(grade)
    if sex is not None:
        m &= clean['sex'].eq(sex)
    if among is not None:
        m &= clean[among].astype(bool)
    n = int((clean[flag_col].astype(bool) & m).sum()); d = int(m.sum())
    return pct(n, d, name=name, base=base or f"{grade or 'all'} {sex or ''} students {('who ' + among) if among else ''}".strip())


############################## ############################## ##############################
### Appendices 
############################## ############################## ##############################

##############################
### Functions
##############################

############################## ############################## ##############################
### Appendix A - basic count and percent
############################## ############################## ##############################

def create_count_and_percent(df, col, colname, startswithnum, agefix):
    df2 = df.groupby(col).count()[['Token']].reset_index()
    df2.columns = [colname, 'Count (N)']
    total = df2['Count (N)'].sum()
    if total == 0:
        df2['Percentage (%)'] = '0.0%'
    else:
        percenter = round((df2['Count (N)'] / total) * 100, 1).astype(str) + '%'
        df2['Percentage (%)'] = percenter
    
    if startswithnum =='Yes':
        ## for grades or ages, sort by the number that starts the string
        df2['Numeric'] =pd.to_numeric( df2[colname].str.extract(r'(^\d*)', expand=True)[0], errors='coerce')
        df2 =df2.sort_values('Numeric')
        df2.drop('Numeric', axis=1, inplace=True)
        
        return df2
    
    if agefix =='Yes':
        df2['agenum'] = df2[colname].astype(str).str.split(" ").str[0]
        df2['agenum'] = df2['agenum'].astype(float)
        
        
        stats = [min(df2[colname]), max(df2[colname]), df2[colname].mean(), df2[colname].std()]
        statnames = ['Minage', 'Maxage', 'Meanage', 'Stdage']

        ### why does the index exclude the 19 year old? the 19 year old is also in the R code

        ## looks like that 19 year old is in 9th grade? 

        age_stats =  pd.DataFrame(
            {'Stat': statnames,
             'Values': stats
            })
        
        return age_stats, df2
        
    if (startswithnum !='Yes') & (agefix !='Yes'):
        return df2
    
    
# def fix_missing_zero_values_intable(ogdf, newdf, colog, newcol, listofall):
#     missingvals = list(set(ogdf[colog].unique().tolist())^ set(listofall))
#     missingdf = pd.DataFrame({newcol: missingvals,
#              'Count (N)': 0,
#               'Percentage (%)': '0%'
#              })
#     return pd.concat([newdf, missingdf]).sort_values([newcol]).reset_index(drop=True).dropna(axis=0)


def create_count_and_percent_0check(df, col, colname, startswithnum, agefix, listofall):
    df2 = df.groupby(col).count()[['Token']].reset_index()
    df2.columns = [colname, 'Count (N)']
    total = df2['Count (N)'].sum()
    if total == 0:
        df2['Percentage (%)'] = '0.0%'
    else:
        percenter = round((df2['Count (N)'] / total) * 100, 1).astype(str) + '%'
        df2['Percentage (%)'] = percenter
    
    if startswithnum =='Yes':
        ## for grades or ages, sort by the number that starts the string
        df2['Numeric'] =pd.to_numeric( df2[colname].str.extract(r'(^\d*)', expand=True)[0], errors='coerce')
        df2 =df2.sort_values('Numeric')
        df2.drop('Numeric', axis=1, inplace=True)
        
        return df2
    
    if agefix =='Yes':
        df2['agenum'] = df2[colname].astype(str).str.split(" ").str[0]
        df2['agenum'] = df2['agenum'].astype(float)
        
        
        stats = [min(df2[colname]), max(df2[colname]), df2[colname].mean(), df2[colname].std()]
        statnames = ['Minage', 'Maxage', 'Meanage', 'Stdage']

        ### why does the index exclude the 19 year old? the 19 year old is also in the R code

        ## looks like that 19 year old is in 9th grade? 

        age_stats =  pd.DataFrame(
            {'Stat': statnames,
             'Values': stats
            })
        
        return age_stats, df2
        
    if (startswithnum !='Yes') & (agefix !='Yes'):
        missingvals = list(set(df[col].unique().tolist())^ set(listofall))
        missingdf = pd.DataFrame({colname: missingvals,
             'Count (N)': 0,
              'Percentage (%)': '0.0%'
             })
        return pd.concat([df2, missingdf]).sort_values([colname]).reset_index(drop=True).dropna(axis=0)


def sort_by_customlist(df, col, orderlist):
    # Restrict to rows in orderlist, then reindex so order matches orderlist (avoids KeyError if data has extra values)
    df = df[df[col].isin(orderlist)].set_index(col)
    df = df.reindex(orderlist).reset_index()
    return df

def multi_questioner(listofQ, sortorderlist, colname):
    Answers = []
    for Q in listofQ:
        print(Q)
        newdf = create_count_and_percent(clean, Q, colname, 'No', 'No')
        missingvals = list(set(newdf[colname].unique().tolist())^ set(sortorderlist))

        if len(missingvals)>0:
            missingdf = pd.DataFrame({colname: missingvals,
             'Count (N)': 0,
              'Percentage (%)': '0.0%'
             })
            newdf = pd.concat([newdf, missingdf]).reset_index(drop=True).dropna(axis=0)

        newdf = sort_by_customlist(newdf, colname, sortorderlist)
        Answers.append(newdf)
    return Answers


def binary_Qer(Q):
    # B7: always both rows (Yes and No), keyed by answer, never by position.
    newdf = create_count_and_percent_0check(clean, Q, 'Response', 'No', 'No', YN_order2)
    newdf = sort_by_customlist(newdf, 'Response', YN_order2)
    newdf['Count (N)'] = newdf['Count (N)'].fillna(0)
    newdf['Percentage (%)'] = newdf['Percentage (%)'].fillna('0.0%')
    return newdf



def binary_multiQer(listofQ, colname):
    Answers = []
    for Q in listofQ:
        print(Q)
        newdf = binary_Qer(Q)
        Answers.append(newdf)
    return Answers



############################## ############################## ##############################
### Appendix B - Results by Grade Level
############################## ############################## ##############################


def by_grader_count_and_percenter(xrows, ycolumns, answerslist, sortorderlist, newcolname):
    #totals = clean[[xrows] + ['id']].groupby(xrows).count().reset_index()
    #totals.rename(columns={'id':'total'}, inplace=True)
    
    countdf =clean[[xrows] + [ycolumns] +['Token']].groupby([xrows] + [ycolumns]).count().reset_index().sort_values(ycolumns).pivot(index=xrows, columns=ycolumns, values='Token').reset_index()

    missingcols =list(set( answerslist)^ set(countdf.columns))
    #print(missingcol)
    if len(missingcols)>1:
        #print(missingcols)
        for m in missingcols:
            if m in answerslist:
                #print('missing column fixed')
                countdf[m] = np.nan
                #print(countdf)
    
    countdf = countdf[[xrows] + answerslist  ]

    #countdf = pd.merge(countdf, totals, on=xrows)
    #print(countdf)
    
    
    output = pd.DataFrame()
    output[xrows] = countdf[xrows]
    
    for col in answerslist:
        col_sum = countdf[col].sum()
        output[f"{col} (N)"] = np.where(countdf[col].isna(), 0, countdf[col])
        output[f"{col} (%)"] = np.where(
            countdf[col].isna() | (col_sum == 0), '0.0%',
            round(countdf[col] / col_sum * 100, 1).astype(str) + '%'
        )
       
    #print(output['grades_overall'])
    
    missingvals = list(set(clean[xrows].unique().tolist())^ set(sortorderlist))
    
    if len(missingvals)>0:
        #print(missingvals)
        missingdf = pd.DataFrame()
        missingdf[xrows] = missingvals
        
        for thing in answerslist:
            missingdf[f"{thing} (N)"] = [0] * len(missingvals)
            missingdf[f"{thing} (%)"] = ['0.0%']  * len(missingvals)
            
        COMBO =pd.concat([output, missingdf]).reset_index(drop=True)
        COMBO= COMBO.set_index(xrows).loc[sortorderlist].reset_index()
        COMBO.rename(columns={ COMBO.columns[0]: newcolname }, inplace = True)
        if ycolumns == 'grade_level':
            return _pad_grade_cols(COMBO, newcolname)
        return COMBO

    if len(missingvals) ==0:
        output = output.set_index(xrows).loc[sortorderlist].reset_index()
        output.rename(columns={ output.columns[0]: newcolname }, inplace = True)
        if ycolumns == 'grade_level':
            return _pad_grade_cols(output, newcolname)
        return output


def multi_by_grader(listofQ, ycolumns, answerslist, sortorderlist, newcolname):
    Answers = []
    for Q in listofQ:
        #print(Q)
        newdf = by_grader_count_and_percenter(Q,ycolumns, answerslist, sortorderlist, newcolname )
        Answers.append(newdf)
    return Answers


def binary_grade_fixer(df):
    col = df.columns[0]
    df[col] = df[col].apply(lambda x: 'Yes' if x in _YES_VALUES else ('No' if x in _NO_VALUES else x))

    
############################## ############################## ##############################
### Appendix C Results by Drinking Behavior
############################## ############################## ##############################


# def risk_categorizer(df):
#     if (df['days_5_or_more_drinks'] != 'I did not consume 5 or more whole drinks of alcohol in a row over the past 30 days') & ( pd.notna(df['days_5_or_more_drinks'])) :
#         return 'High-risk'
#     if (df['how_often_1_drink']== "1-2 times per year") & (df['one_drink_past_12']=="Yes") & (df['days_5_or_more_drinks'] =="I did not consume 5 or more whole drinks of alcohol in a row over the past 30 days"):
#         return 'Low-risk'
#     if (df['age_at_first_drink']=='I have never had a whole drink of alcohol') or (df['one_drink_past_12']=="No" ):
#         return "Non-drinker"
#     else:
#         return 'Missing'
    
# clean['drinker_risk'] = clean.apply(risk_categorizer, axis=1)



# Group definitions (B1, B8, R4). The SAME text is printed in the report body
# and Appendix C (see HIGHRISK_DEFINITION etc. below), so code and wording
# cannot drift apart again.
#   Higher-risk: 5+ drinks in a row on 1 or more days in the past 30 days.
#   Lower-risk:  drank in the past 12 months (canonical drank_12m flag), but
#                not 5+ in a row in the past 30 days. (R4: students who said
#                "Yes" to the past 12 months but "Never" to typical frequency
#                used to fall into no group; they are lower-risk now.)
#   Non-drinker: did not drink in the past 12 months (includes never drank).
_NO_BINGE = 'I did not consume 5 or more whole drinks of alcohol in a row over the past 30 days'

def risk_categorizer(df):
    if pd.notna(df['days_5_or_more_drinks']) and df['days_5_or_more_drinks'] != _NO_BINGE:
        return 'High-risk'
    if df['drank_12m']:
        return 'Low-risk'
    if (df['age_at_first_drink'] == 'I have never had a whole drink of alcohol') or (df['one_drink_past_12'] == 'No'):
        return 'Non-drinker'
    return 'Missing'

clean['drinker_risk'] = clean.apply(risk_categorizer, axis=1)

HIGHRISK_DEFINITION = ('A “higher-risk drinker” is someone who reported consuming 5 or more drinks of alcohol '
                       'in a row on 1 or more days within the past 30 days.')
LOWRISK_DEFINITION = ('A “lower-risk drinker” is someone who reported having at least one whole drink of alcohol '
                      'within the past 12 months, but did not consume 5 or more drinks in a row on any day within the past 30 days.')
NONDRINKER_DEFINITION = ('A “non-drinker” is someone who reported not having a whole drink of alcohol within the past 12 months, '
                         'including students who have never had a whole drink.')
BODY_GROUP_NOTE = f'drinking groups: grades {BODY_GROUP_LABEL_PLAIN}'
APPENDIX_C_GROUP_NOTE = ('Groups in this appendix include students in all grades surveyed. Sections of the report body '
                         f'that compare drinking groups use students in grades {BODY_GROUP_LABEL_PLAIN} only.')
_missing_group = int((clean['drinker_risk'] == 'Missing').sum())
if _missing_group:
    logger.warning(f"R4: {_missing_group} students fit no drinking group (missing answers) and are left out of the groups.")

def risk_categorizer_subsetbygrade(df):
    # B3: body groups = BODY_GROUP_GRADES (one setting at the top of the script)
    if df['grade_level'] in BODY_GROUP_GRADES:
        return df['drinker_risk']
    return 'Missing'

clean['drinker_risk_10_12'] = clean.apply(risk_categorizer_subsetbygrade, axis=1)


##############################
### BUILD RECORD: the cleaned counts this report is built on. After the report
### is saved, they are written next to it (<report>_build_record.json) together
### with a copy of this script (<report>_SUPP_script_used.py). The SUPP Two-Year
### Trend Report uses both to clean the data with exactly the same rules later
### and stops if its numbers do not match this report.
##############################
_RECORD_FLAGS = ['drank_lifetime', 'drank_12m', 'drank_30d', 'pot_lifetime', 'pot_12m', 'pot_30d',
                 'vape_12m', 'vape_30d']
BUILD_RECORD = {
    'raw_n': int(N), 'removed_n': int(removed_N), 'analyzed_n': int(len(clean)),
    'counts': {str(_g): {'n': int(len(_sub)), **{_f: int(_sub[_f].astype(bool).sum())
                                                 for _f in _RECORD_FLAGS if _f in _sub.columns}}
               for _g, _sub in [('All', clean)] + [(g, clean[clean['grade_level'] == g])
                                                   for g in clean['grade_level'].dropna().unique()]},
}


def risk_categorizer_pot(df):
    # B8: "yearly or more frequent" = typically uses cannabis 1-2 times a year or more
    if df['grade_level'] in BODY_GROUP_GRADES:
        if pd.notna(df['how_often_pot']) and (df['how_often_pot']!='Never'):
            return 'Yearly or More Frequent Marijuana Users'
    return None

clean['yearlypot'] = clean.apply(risk_categorizer_pot, axis=1)


def risk_categorizer_pot2(df):
    if df['grade_level'] in BODY_GROUP_GRADES:
        if df['pot_30d']:
            return 'Cannabis use in last 30 days'
        if not df['pot_lifetime']:
            return 'Never used cannabis'
    return None

clean['pot_risk30_never'] = clean.apply(risk_categorizer_pot2, axis=1)

############################## ############################## ##############################
### Report
############################## ############################## ##############################

def new_percenter(df, condition, gradecol):
    """Calculate 'ever used' percentage as 100 - 'never used' percentage.
    CORRECTED: Returns rounded scalar wrapped in _ItemableValue for .item() compat;
    returns 'N/A' instead of blank for missing grades."""
    if gradecol in df.columns:
        ## how many have ever drank? 100 - %who have not, etc
        col_data = df[condition][gradecol]
        if col_data.eq('').all() or col_data.empty:
            return _wrap_scalar('N/A')
        try:
            val = 100 - (col_data.str.split('%').str[0]).astype(float)
            # Return a rounded scalar, not a Series
            if len(val) == 1:
                return _wrap_scalar(round(val.item(), 1))
            return round(val, 1)
        except (ValueError, TypeError):
            return _wrap_scalar('N/A')
    else:
        return _wrap_scalar('N/A')


    #### updated to handling missing 8th grade
def by_grader_count_and_sex_percenter(xrows, ycolumns,  sortorderlist, keepgrades, newrow):
    #totals = clean[[xrows] + ['id']].groupby(xrows).count().reset_index()
    #totals.rename(columns={'id':'total'}, inplace=True)
    
    countdf =clean[[xrows] + ycolumns +['Token']].groupby([xrows] + ycolumns).count().reset_index().sort_values(ycolumns).pivot(index=xrows, columns=ycolumns, values='Token').reset_index()
    #print(countdf)
    hasnt_df = countdf[(countdf[xrows].str.contains("never")) | (countdf[xrows].str.contains("Never")) ].sum().reset_index()
    has_df = countdf[~((countdf[xrows].str.contains("never")) | (countdf[xrows].str.contains("Never")))].sum().reset_index()
    hasnt_df=hasnt_df[1:]
    has_df=has_df[1:]
    has_df.rename(columns={0:'Yes'}, inplace=True)
    hasnt_df.rename(columns={0:'No'}, inplace=True)
    combo = pd.merge(has_df, hasnt_df, on=ycolumns )
    denom = combo['Yes'] + combo['No']
    combo['Perc Ever'] = np.where(denom == 0, 0, (combo['Yes'] / denom) * 100)
    if '8th' in combo['grade_level'].values:
        combo = combo[ycolumns + ['Perc Ever']].pivot(index=ycolumns[1], columns=ycolumns[0], values ='Perc Ever').reset_index()[['sex']+sortorderlist]
        combo = combo[keepgrades]
        combo.loc[-1]  = newrow
        combo= combo.set_index('sex').loc[['Total', 'Male', 'Female']].reset_index()

    else:
        combo = combo[ycolumns + ['Perc Ever']].pivot(index=ycolumns[1], columns=ycolumns[0], values ='Perc Ever').reset_index()[['sex']+['10th', '12th']]
        combo['8th'] = ''
        combo = combo[keepgrades]
        combo.loc[-1]  = newrow
        combo= combo.set_index('sex').loc[['Total', 'Male', 'Female']].reset_index()

    return combo



def ever_by_age_yes(df, gradecol):
    """CORRECTED: Return _ItemableValue for .item() compat.
    BUG FIX: Changed != to == so it returns the 'Yes' percentage
    (who DID drink), matching the template text about rising alcohol use."""
    if gradecol in df.columns:
        col_data = df[df['Amount']=='Yes'][gradecol]
        if col_data.eq('').all():
            return _wrap_scalar('')
        val = (col_data.str.split('%').str[0]).astype(float).sum()
        return _wrap_scalar(val)
    else:
        return _wrap_scalar('')
    
def percent_of_binary_where_total_elsewhere(yesval, totalval, grade):
    """CORRECTED: Returns _ItemableValue wrapper so .item() calls work."""
    if f'{grade} (N)' in yesval.columns:
        yesdf = yesval[f'{grade} (N)']
        if yesdf.eq('').all():
            return _wrap_scalar('')
        total = totalval[f'{grade} (N)'].sum()
        if total == '' or total == 0:
            return _wrap_scalar('')
        yes_perc = round((yesdf/ total)*100,1)
        return yes_perc
    else:
        return _wrap_scalar('')


def total_agree_disagree(df):
    """Compute agree/disagree percentages per grade for culture-of-use tables.

    DENOMINATOR NOTE (Category F): This function uses per-grade N (count) columns
    as denominators, NOT the total student count. This means each grade's percentage
    is calculated as (agree_count_for_grade / total_responses_for_grade) * 100.
    This is intentional: it shows the percentage within each grade group, allowing
    valid cross-grade comparisons even when grade sizes differ. The resulting
    percentages may not match what you'd get using total N=286 as denominator.
    """
    dislist= []
    agglist =[]
    # Use _FULL_GRADES_N (8th-12th) so all 5 grades are included even if the
    # school only has 9th-12th.  Columns padded by _pad_grade_cols will have 0
    # values, producing NaN percentages that fillna(0) handles below.
    use_grades_n = [g for g in _FULL_GRADES_N if g in df.columns]
    dis_df = df[(df['Response'].str.contains('Disagree')) |(df['Response'].str.contains('disagree'))][['Response'] + use_grades_n].sum()
    agg_df = df[(df['Response'].str.contains('Agree')) | (df['Response'].str.contains('Strongly agree'))][['Response'] + use_grades_n].sum()
    total = df[['Response'] + use_grades_n].sum()

    newdf = pd.DataFrame({'disagree': dis_df, 'agree': agg_df, 'total':total})
    newdf = newdf[1:]
    newdf['perc dis'] = ((newdf['disagree']/ newdf['total'])*100).fillna(0).pipe(round, 1)
    newdf['perc agg'] = ((newdf['agree']/ newdf['total'])*100).fillna(0).pipe(round, 1)
    newdf =newdf[['perc dis','perc agg']].reset_index()
    newdf.columns =['Grade', 'Total Disagree', 'Total Agree']
    newdf =newdf.T
    newdf.columns = newdf.iloc[0]
    newdf = newdf[1:].reset_index()
    newdf =newdf.set_index('index').loc[['Total Agree', 'Total Disagree']].reset_index()
    # Template reads [0,1,2,3,4,5] = label + one value per grade (8th-12th).
    # Ensure all 5 grade columns exist; add any missing with empty string.
    for g in _FULL_GRADES_N:
        if g not in newdf.columns:
            newdf[g] = 0.0
    return newdf[['index'] + _FULL_GRADES_N]


def risk_by_subset_pot_andall_1012th(xrows,  answerslist, sortorderlist, newcolname):
    
    al_risk = by_grader_count_and_percenter(xrows, 'drinker_risk_10_12', answerslist, sortorderlist, newcolname )
    al_risk =al_risk[[newcolname,'High-risk (%)', 'Low-risk (%)', 'Non-drinker (%)']]
    
    yearlypottotal = clean[clean['yearlypot'].notna()][xrows].count()
    ypot_risk = clean[clean['yearlypot'].notna()][xrows].value_counts().reset_index()

    ypot_risk[xrows] = ypot_risk[xrows] #.astype(int)
    if yearlypottotal == 0:
        ypot_risk['perc'] = '0.0%'
    else:
        ypot_risk['perc'] = (ypot_risk['count'] / yearlypottotal) * 100
        ypot_risk['perc'] = round(ypot_risk['perc'].astype(float), 1).astype(str) + '%'
    ypot_risk.columns =[newcolname, 'count', 'Yearly or More Frequent Marijuana Use']
    al_pot_risk =pd.merge(al_risk, ypot_risk, on=newcolname,  how='left').drop('count', axis=1).fillna('0.0%')
    
   
    all1012grades = clean[clean['grade_level'].isin(BODY_GROUP_GRADES)][xrows].value_counts().reset_index()
    all1012total = clean[clean['grade_level'].isin(BODY_GROUP_GRADES)][xrows].count().sum()
    if all1012total == 0:
        all1012grades['perc'] = '0.0%'
    else:
        all1012grades['perc'] = (all1012grades['count'] / all1012total) * 100
        all1012grades['perc'] = round(all1012grades['perc'].astype(float), 1).astype(str) + '%'
    all1012grades.columns =[newcolname, 'count', '10-12th Graders']
    
    all_risk = pd.merge(al_pot_risk, all1012grades, on=newcolname,  how='left').drop('count', axis=1).fillna('0.0%')
    return all_risk[[newcolname, 'Yearly or More Frequent Marijuana Use','High-risk (%)', 'Low-risk (%)', 'Non-drinker (%)',
        '10-12th Graders']]

def percent_of_binary_where_total_elsewhere_bysex(yesval, totalval, grade):
    if f'{grade} (N)' not in yesval.columns or f'{grade} (N)' not in totalval.columns:
        return ''
    if yesval[f'{grade} (N)'].eq('').all():
        return ''
    yesdf = yesval[f'{grade} (N)'].sum()
    total = totalval[f'{grade} (N)'].sum()
    if total == 0 or total == '':
        return ''
    yes_perc = round((yesdf/ total)*100,1)
    return yes_perc



def by_grader_count_and_sex_percenter_yn(xrows, keepgrades, newrow):
    noyes = clean[[xrows, 'grade_level', 'sex', 'Token']].groupby([xrows,'grade_level', 'sex']).count().reset_index()
    noyes =noyes[noyes[xrows]=='Yes']
    noyes =noyes[['grade_level', 'sex', 'Token']]
    noyes.columns = ['grade_level', 'sex', 'countyes']
    totals= clean[[xrows, 'grade_level', 'sex', 'Token']].groupby(['grade_level', 'sex']).count().reset_index()
    totals =totals[['grade_level', 'sex', 'Token']]
    totals.columns = ['grade_level', 'sex', 'totalbysex']
    perctable = pd.merge(noyes.reset_index(), totals.reset_index(), on=['grade_level', 'sex'], how='right')
    perctable['perc'] = round((perctable['countyes'] / perctable['totalbysex'])*100, 1)
    perctable =perctable.pivot(index='sex', columns='grade_level', values='perc').reset_index()
    if '8th' in perctable.columns:
        perctable =perctable[keepgrades]
        perctable.loc[-1]  = newrow
        perctable= perctable.set_index('sex').loc[['Total', 'Male', 'Female']].reset_index()
    else:
        perctable['8th'] =''
        perctable =perctable[keepgrades]
        perctable.loc[-1]  = newrow
        perctable= perctable.set_index('sex').loc[['Total', 'Male', 'Female']].reset_index()
    return perctable




def create_high_low_10_non_bullets(df, colname):
    yes_row = df[df[colname] == 'Yes']
    if len(yes_row) == 0:
        return '0.0%', '0.0%', '0.0%', '0.0%'
    high_risk_perc = yes_row['High-risk (%)'].iloc[0]
    ten12_risk_perc = yes_row['10-12th Graders'].iloc[0]
    low_risk_perc = yes_row['Low-risk (%)'].iloc[0]
    norisk = yes_row['Non-drinker (%)'].iloc[0]
    return high_risk_perc, ten12_risk_perc, low_risk_perc, norisk


def fix_MTF_naming(df):
    return df.replace({
        'Vapes': 'Electronic Vaporizers**',
        'Vaping Nicotine': 'Electronic Vaporizers**',
        'Hallucinogen': 'Hallucinogens',
        'Lsd': 'LSD',
        'Cold': 'Over the Counter Cold Medicine',
        'Pot': 'Marijuana',
        'Methamphetamine': 'Meth',
    })

def total_aharm_no_harm_nograde_PERC(df):
    dislist= []
    agglist =[]
    df = df.copy()
    # B4: combine from COUNTS, not from already-rounded row percentages.
    # (Output column keeps its old name so the rest of the script is unchanged.)
    if 'Count (N)' in df.columns:
        df['Percentage (%)'] = pd.to_numeric(df['Count (N)'], errors='coerce').fillna(0)
    else:
        df['Percentage (%)'] =pd.to_numeric( df['Percentage (%)'].astype(str).str.split('%').str[0], errors='coerce')
    dis_df = df[(df['Response']=='No risk') |(df['Response'].str.lower()=='slight risk')][['Response','Percentage (%)']].sum()
    agg_df = df[(df['Response']=='Moderate risk') | (df['Response']=='Great risk')][['Response','Percentage (%)']].sum()
    total = df[['Response','Percentage (%)']].sum()
    
    newdf = pd.DataFrame({'disagree': dis_df, 'agree': agg_df, 'total':total})
    newdf = newdf[1:]
    newdf['perc dis'] = ((newdf['disagree']/ newdf['total'])*100).fillna(0).pipe(round, 1)
    newdf['perc agg'] = ((newdf['agree']/ newdf['total'])*100).fillna(0).pipe(round, 1)
    newdf =newdf[['perc dis','perc agg']].reset_index()
    newdf.columns =['Grade', 'Total low risk', 'Total mod/great']
    newdf =newdf.T
    newdf.columns = newdf.iloc[0]
    newdf = newdf[1:].reset_index()
    newdf =newdf.set_index('index').loc[['Total low risk', 'Total mod/great']].reset_index()
    return newdf



def total_agree_disagree_nograde(df):
    dislist= []
    agglist =[]
    dis_df = df[(df['Response'].str.contains('Disagree')) |(df['Response'].str.contains('disagree'))][['Response','Count (N)']].sum()
    agg_df = df[(df['Response'].str.contains('Agree')) | (df['Response'].str.contains('Strongly agree'))][['Response','Count (N)']].sum()
    total = df[['Response','Count (N)']].sum()
    
    newdf = pd.DataFrame({'disagree': dis_df, 'agree': agg_df, 'total':total})
    newdf = newdf[1:]
    newdf['perc dis'] = ((newdf['disagree']/ newdf['total'])*100).fillna(0).pipe(round, 1)
    newdf['perc agg'] = ((newdf['agree']/ newdf['total'])*100).fillna(0).pipe(round, 1)
    newdf =newdf[['perc dis','perc agg']].reset_index()
    newdf.columns =['Grade', 'Total Disagree', 'Total Agree']
    newdf =newdf.T
    newdf.columns = newdf.iloc[0]
    newdf = newdf[1:].reset_index()
    newdf =newdf.set_index('index').loc[['Total Agree', 'Total Disagree']].reset_index()
    return newdf#[['index','7th (N)', '8th (N)','9th (N)', '10th (N)', '11th (N)', '12th (N)']]




def total_agree_disagree_nograde_PERC(df):
    dislist= []
    agglist =[]
    df = df.copy()
    # B4: combine from COUNTS, not from already-rounded row percentages.
    # (Output column keeps its old name so the rest of the script is unchanged.)
    if 'Count (N)' in df.columns:
        df['Percentage (%)'] = pd.to_numeric(df['Count (N)'], errors='coerce').fillna(0)
    else:
        df['Percentage (%)'] =pd.to_numeric( df['Percentage (%)'].astype(str).str.split('%').str[0], errors='coerce')
    dis_df = df[(df['Response'].str.contains('Disagree')) |(df['Response'].str.contains('disagree'))][['Response','Percentage (%)']].sum()
    agg_df = df[(df['Response'].str.contains('Agree')) | (df['Response'].str.contains('Strongly agree'))][['Response','Percentage (%)']].sum()
    total = df[['Response','Percentage (%)']].sum()
    
    newdf = pd.DataFrame({'disagree': dis_df, 'agree': agg_df, 'total':total})
    newdf = newdf[1:]
    newdf['perc dis'] = ((newdf['disagree']/ newdf['total'])*100).fillna(0).pipe(round, 1)
    newdf['perc agg'] = ((newdf['agree']/ newdf['total'])*100).fillna(0).pipe(round, 1)
    newdf =newdf[['perc dis','perc agg']].reset_index()
    newdf.columns =['Grade', 'Total Disagree', 'Total Agree']
    newdf =newdf.T
    newdf.columns = newdf.iloc[0]
    newdf = newdf[1:].reset_index()
    newdf =newdf.set_index('index').loc[['Total Agree', 'Total Disagree']].reset_index()
    return newdf


def total_agree_disagree_perc(df, percentindicator):
    df = df.copy()
    if percentindicator=='yes':
        for col in grades_perc:
            # B4: use the grade's COUNT column so combined agree/disagree
            # figures are computed from counts, not summed rounded percentages.
            _ncol = col.replace('(%)', '(N)')
            if _ncol in df.columns:
                df[col] = pd.to_numeric(df[_ncol], errors='coerce').fillna(0)
            else:
                df[col] = pd.to_numeric(df[col].astype(str).str.split('%').str[0], errors='coerce')
    dislist= []
    agglist =[]
    dis_df = df[(df['Response'].str.contains('Disagree')) |(df['Response'].str.contains('disagree'))][['Response']+grades_perc].sum()
    agg_df = df[(df['Response'].str.contains('Agree')) | (df['Response'].str.contains('Strongly agree'))][['Response'] + grades_perc].sum()
    

    
    total = df[['Response'] + grades_perc].sum()
 
    newdf = pd.DataFrame({'disagree': dis_df, 'agree': agg_df, 'total':total})
    newdf = newdf[1:]


    newdf['perc dis'] = ((newdf['disagree']/ newdf['total'])*100).fillna(0).pipe(round, 1)
    newdf['perc agg'] = ((newdf['agree']/ newdf['total'])*100).fillna(0).pipe(round, 1)
    newdf =newdf[['perc dis','perc agg']].reset_index()
    newdf.columns =['Grade', 'Total Disagree', 'Total Agree']
    newdf =newdf.T
    newdf.columns = newdf.iloc[0]
    newdf = newdf[1:].reset_index()
    newdf =newdf.set_index('index').loc[['Total Agree', 'Total Disagree']].reset_index()
    return newdf

def perceived_pot(grade):
    thing = pot_how_often_your_grade_grade[pot_how_often_your_grade_grade['Frequency']=='Never'][f'{grade} (%)'].item()
    return thing

### add_pot_30_to_df removed (only used by K10 tables, which this school opted out of)

def create_bullet_point(df, surveryval, grade, per_timeline=None):
    """Build the 'is below/on par with/exceeds U.S. norms for ...' bullet.

    per_timeline: optional dict {timeline: school value}. When given, each MTF
    timeline row is compared against the school's value for THAT timeline.
    Without it, the single surveryval (the school's lifetime rate) is compared
    against every timeline's norm, which makes e.g. a school's lifetime rate
    look like it "exceeds" the much smaller past-30-days norm.
    """
    # Ensure surveryval is a plain Python float (not a Series or DataFrame)
    if hasattr(surveryval, 'iloc'):
        try:
            surveryval = float(surveryval.iloc[0])
        except Exception:
            surveryval = 0.0
    else:
        try:
            surveryval = float(surveryval)
        except Exception:
            surveryval = 0.0

    below = []
    sigabove = []
    above = []
    onpar = []


    for index, row in df.iterrows():
        _mtfval = pd.to_numeric(row[grade], errors='coerce')
        if pd.isna(_mtfval):
            continue
        _sv = surveryval
        if per_timeline is not None:
            _sv = per_timeline.get(row['timeline'])
            if _sv is None:
                continue
        # B9: a gap must exceed the margin AND one student's worth of
        # percentage points before it is called below/above U.S. norms.
        _v = norm_verdict(_sv, _mtfval, grade_size(grade))
        if _v == 'onpar':
            onpar.append(row['timeline'])
        elif _v == 'below':
            below.append(row['timeline'])
        elif _v == 'sigabove':
            sigabove.append(row['timeline'])
        elif _v == 'above':
            above.append(row['timeline'])
           
    outputstring = []    

    if (len(below)==1) | (len(below)==2):
        outputstring.append(f"is below U.S. norms for {' and '.join(below)}")
    if (len(below)>=3) :
        outputstring.append(f"is below U.S. norms for {', '.join(below[:2])}, and {below[2]}")  
        
    if (len(onpar)==1) | (len(onpar)==2):
        outputstring.append(f"is on par with U.S. norms for {' and '.join(onpar)}")
    if (len(onpar)>=3) :
        outputstring.append(f"is on par with U.S. norms for {', '.join(onpar[:2])}, and {onpar[2]}")  
        
    if (len(above)==1) | (len(above)==2):
        outputstring.append(f"is above U.S. norms for {' and '.join(above)}")
    if (len(above)>=3) :
        outputstring.append(f"is above U.S. norms for {', '.join(above[:2])}, and {above[2]}")  
    
    if (len(sigabove)==1) | (len(sigabove)==2):
        outputstring.append(f"is well above U.S. norms for {' and '.join(sigabove)}")
    if (len(sigabove)>=3) :
        outputstring.append(f"is well above U.S. norms for {', '.join(sigabove[:2])}, and {sigabove[2]}") 
    if not outputstring:
        return ''
    # B7: "is below ... for X, is on par ... for Y, and is above ... for Z measures."
    return _join_and(outputstring) + ' measures.'




##############################
### C21-C24: charts are drawn at their FINAL size with the fonts wanted and
### inserted at 100% (never shrunk in Word, which shrank the axis and legend
### text too). Sizes from Part 3 of the Nov 2026 fix list.
##############################
CHART_SIZE_MTF = (5.4, 3.6)     # U.S. norm comparison charts (one page per measure)
CHART_SIZE_PAIR = (4.7, 2.9)    # culture of use / school attitudes (two pairs per page)
CHART_SIZE_K10 = (5.4, 3.6)     # Kessler charts (one table + chart per page)
CHART_FONT_PT = 9
GROUP_DISPLAY = {'High-risk (%)': 'Higher-risk', 'Low-risk (%)': 'Lower-risk', 'Non-drinker (%)': 'Non-drinkers',
                 'High-risk': 'Higher-risk', 'Low-risk': 'Lower-risk', 'Non-drinker': 'Non-drinkers',
                 'Cannabis use in last 30 days': 'Cannabis in past 30 days',
                 'Never used cannabis': 'Never used cannabis',
                 'Using substances to cope': 'Using substances to cope'}

_YR_LABEL = {'7th': 'Yr 8', '8th': 'Yr 9', '9th': 'Yr 10', '10th': 'Yr 11', '11th': 'Yr 12', '12th': 'Yr 13'}

def _grade_label(g):
    """Chart label for a grade: 'Yr 11' at a year-group school, otherwise '10th'."""
    g = str(g).strip()
    if '_was_year_system' not in clean.columns:
        return g
    for k, v in _YR_LABEL.items():
        if g == k or g.startswith(k + ' '):
            return v + g[len(k):]
    return g

def _grade_axis():
    return 'Year Group' if '_was_year_system' in clean.columns else 'Grade'

def _draw_bars(df_melt, x, y, hue, palette, save, size, xlabel, ylabel='Percentage (%)', rotate=0, order=None, hue_order=None,
               legend_cols=4, legend_y=None):
    import matplotlib.pyplot as _plt
    with _plt.rc_context({'font.size': CHART_FONT_PT, 'axes.labelsize': CHART_FONT_PT,
                          'xtick.labelsize': CHART_FONT_PT - 1, 'ytick.labelsize': CHART_FONT_PT - 1,
                          'legend.fontsize': CHART_FONT_PT - 1}):
        fig, ax = _plt.subplots(figsize=size)
        sns.barplot(data=df_melt, x=x, y=y, hue=hue, palette=palette, ax=ax, order=order, hue_order=hue_order,
                    errorbar=None)
        ax.set_xlabel(xlabel); ax.set_ylabel(ylabel)
        ax.set_ylim(0, max(10, min(100, (pd.to_numeric(df_melt[y], errors='coerce').max() or 0) * 1.15)))
        for _s in ('top', 'right'):
            ax.spines[_s].set_visible(False)
        if rotate:
            _plt.setp(ax.get_xticklabels(), rotation=rotate, ha='right')
        _n_hue = df_melt[hue].nunique()
        ax.legend(title=None, frameon=False, ncol=min(_n_hue, legend_cols), loc='upper center',
                  bbox_to_anchor=(0.5, legend_y if legend_y is not None else (-0.22 if not rotate else -0.42)))
        fig.tight_layout()
        fig.patch.set_facecolor('white')
        fig.savefig(save, dpi=300, facecolor='white')   # no bbox 'tight': keeps the exact size
        _plt.close(fig)

def bar_chart_it_up(df, save):
    palette={    'Total':'#1A63CD', 'Male':'#C9D2DA', 'Female':'#267EF2'}
    _all_grade_cols = ['8th', '8th MTF', '10th', '10th MTF', '12th', '12th MTF']
    _chart_cols = [c for c in _all_grade_cols if c in df.columns]
    # C24: no 8th graders took the survey -> no 8th-grade school bar AND no 8th MTF bar.
    if '8th' not in clean['grade_level'].values:
        _chart_cols = [c for c in _chart_cols if not c.startswith('8th')]
    df_melt =df.melt(id_vars='Type', value_vars=_chart_cols)
    df_melt['value'] = pd.to_numeric(df_melt['value'], errors='coerce')
    df_melt['variable'] = df_melt['variable'].map(_grade_label).str.replace(' MTF', ' U.S.', regex=False)
    _draw_bars(df_melt, 'variable', 'value', 'Type', palette, save, CHART_SIZE_MTF, _grade_axis(),
               hue_order=[t for t in ['Total', 'Male', 'Female'] if t in set(df_melt['Type'])])

def bar_chart_it_up_2(df, save):
    palette={    'Total Agree':'#1A63CD', 'Total Disagree':'#C9D2DA'}
    # Chart only grades with actual respondents — an entered-but-absent grade
    # would otherwise show fabricated 0%-agree/0%-disagree bars.
    _data_grades = set(clean['grade_level'].dropna())
    _chart_grades_n = [g for g in grades_n if g.split(' (')[0] in _data_grades]
    df_melt =df.melt(id_vars ='index', value_vars=_chart_grades_n, var_name='GradeCol')
    df_melt['value'] = pd.to_numeric(df_melt['value'], errors='coerce')
    # C22: x-axis labels are the grade only ("9th"), never "9th (N)".
    df_melt['Grade'] = df_melt['GradeCol'].str.replace(r'\s*\(N\)', '', regex=True).map(_grade_label)
    df_melt['Answer'] = df_melt['index'].replace({'Total Agree': 'Agree', 'Total Disagree': 'Disagree'})
    _draw_bars(df_melt, 'Grade', 'value', 'Answer', {'Agree': palette['Total Agree'], 'Disagree': palette['Total Disagree']},
               save, CHART_SIZE_PAIR, _grade_axis())

def bar_chart_it_up_3(df, save):
    palette={    'Liked a lot':'#1A63CD', 'Kind of liked':'#C9D2DA', 'Did not like':'#267EF2', 'Hated': '#838383'}
    # Chart only grades with actual respondents (see bar_chart_it_up_2).
    _data_grades = set(clean['grade_level'].dropna())
    vv = [g + ' (%)' for g in ['7th', '8th', '9th', '10th', '11th', '12th'] if g in _data_grades]
    df_melt =df[['Response'] + vv].melt(id_vars ='Response', value_vars=vv)
    df_melt['value'] = pd.to_numeric(df_melt['value'].astype(str).str.replace('%', '', regex=False), errors='coerce')
    df_melt['Grade'] =df_melt['variable'].str.replace(r'\s*\(%\)', '', regex=True).map(_grade_label)
    _draw_bars(df_melt, 'Grade', 'value', 'Response', palette, save, CHART_SIZE_PAIR, _grade_axis(),
               hue_order=[r for r in palette if r in set(df_melt['Response'])])

def bar_chart_it_up_4(df, save, palette):
    _cols = ['High-risk (%)', 'Low-risk (%)', 'Non-drinker (%)',
             'Cannabis use in last 30 days', 'Never used cannabis', 'Using substances to cope']
    df_melt =df[['Response'] + _cols].melt(id_vars ='Response', value_vars=_cols)
    df_melt['value'] = pd.to_numeric(df_melt['value'], errors='coerce')
    # C23: "Higher-risk" / "Lower-risk" everywhere, matching the text and table headers.
    import textwrap as _tw
    df_melt['variable'] = df_melt['variable'].map(lambda v: _tw.fill(GROUP_DISPLAY.get(v, v), 13))
    _resp = list(dict.fromkeys(df_melt['Response']))
    _draw_bars(df_melt, 'variable', 'value', 'Response', dict(zip(_resp, palette)), save, CHART_SIZE_K10, 'Group',
               hue_order=_resp, legend_cols=2, legend_y=-0.3)


def ever_by_age(df, gradecol):
    """CORRECTED: Round result, return _ItemableValue for .item() compat."""
    if gradecol in df.columns:
        col_data = df[df['Age']!='I have never had a whole drink of alcohol'][gradecol]
        if col_data.eq('').all() or col_data.empty:
            return _wrap_scalar('N/A')
        try:
            val = (col_data.str.split('%').str[0]).astype(float).sum()
            return _wrap_scalar(round(val, 1))
        except (ValueError, TypeError):
            return _wrap_scalar('N/A')
    else:
        return _wrap_scalar('N/A')

def never_by_age(df, gradecol):
    """CORRECTED: Return _ItemableValue for .item() compat."""
    if gradecol in df.columns:
        val = (df[df['Age']=='I have never had a whole drink of alcohol'][gradecol]).item()
        if val == '':
            return _wrap_scalar('N/A')
        return val
    else:
        return _wrap_scalar('N/A')


# In[16]:




# def risk_categorizer_subsetbygrade(df):
#     if ((df['grade_level']=='10th')| (df['grade_level']=='11th')| (df['grade_level']=='12th') ):
#         if (df['days_5_or_more_drinks'] != 'I did not consume 5 or more whole drinks of alcohol in a row over the past 30 days') & ( pd.notna(df['days_5_or_more_drinks'])) :
#             return 'High-risk'
#         if (df['how_often_1_drink']== "1-2 times per year") & (df['one_drink_past_12']=="Yes") &  (df['days_5_or_more_drinks'] =="I did not consume 5 or more whole drinks of alcohol in a row over the past 30 days"):
#             return 'Low-risk'
#         if (df['age_at_first_drink']=='I have never had a whole drink of alcohol') or (df['one_drink_past_12']=="No" ):
#             return "Non-drinker"
#         else:
#             return 'Missing'
#     else:
#         return 'Missing'
    
# clean['drinker_risk_10_12'] = clean.apply(risk_categorizer_subsetbygrade, axis=1)

#clean[['days_5_or_more_drinks', 'how_often_1_drink', 'one_drink_past_12',  'age_at_first_drink', 'drinker_risk', 'drinker_risk2']]


# In[ ]:





# In[17]:


### MTF uses - to mean we didn't measure it so replace the nan with "0" in table 

### the appendices are rounded appropriately

### drugs capitalized in tables but lowercase in bullets other than lsd which should be LSD


# In[18]:


#clean[clean['drinker_risk']=='Missing'][['days_5_or_more_drinks', 'how_often_1_drink', 'one_drink_past_12',  'age_at_first_drink', 'drinker_risk']].value_counts()


# In[19]:


##############################
### Answer lists - used to check that all answers are shown even if no one chose one or more options
### also used to make sure the rows are sorted into the right order
##############################

### Grading scale for 'How were your grades overall' (letter grades vs IB 1-7).
### Set by the setup prompt; auto-detected from the data when left blank.
### Hardcoding one scale renders the academics tables as all zeros for schools
### on the other scale.
_number_scale = ["Mostly 1's", "Mostly 2's", "Mostly 3's", "Mostly 4's", "Mostly 5's", "Mostly 6's", "Mostly 7's"]
_letter_scale = ["Mostly A's", "Mostly B's", "Mostly C's", "Mostly D's", "Mostly F's"]
# Five-step wording used by some surveys (no letters or numbers), highest first.
_word_scale = ['Mostly the highest grades', 'Mostly upper-middle grades', 'Mostly middle grades',
               'Mostly lower-middle grades', 'Mostly the lowest grades']
_grades_in_data = set(clean['grades_overall'].dropna().astype(str))
if _grades_in_data & set(_word_scale):
    listofageoptions = _word_scale          # detected from the answers themselves
elif response_grading_scale.startswith('l'):
    listofageoptions = _letter_scale
elif response_grading_scale.startswith('n'):
    listofageoptions = _number_scale
else:
    listofageoptions = _letter_scale if _grades_in_data & set(_letter_scale) else _number_scale
_unmatched_grades = _grades_in_data - set(listofageoptions)
if _unmatched_grades:
    logger.warning(f"Grades overall: answers not on the {listofageoptions[0]}... scale and left out of "
                   f"the academics tables: {sorted(_unmatched_grades)}")
print(f"Grading scale for academics tables: {listofageoptions[0]} ... {listofageoptions[-1]}")




drinkinglist = ['High-risk','Low-risk','Non-drinker' ]



alllocations = ["At a friend's house", 'At home', 'At a club or bar', 'At camp', 'At school', 'Other']

sort_order1 = ['Never',
'1-2 times per year',
'Once a month',
'Twice a month',
'Once a week',
'Twice a week',
'Daily']



amountlist = ["0 drinks",
"1-2 drinks",
"3-4 drinks",
"5 or more drinks"]

amountlist2 = [
"0 drinks",
"1-2 drinks",
"3-4 drinks",
"5 or more drinks"]

sort_order2 = [ 'Never','1 time', '2 times', '3 times or more']

sort_order3 = ['I did not drink alcohol during the past 30 days',
               '1 day',
 '2 days',
 '3 to 5 days',
 '6 to 9 days',
 '10 or more days']


freq_sort = ['0 times', '1 time', '2 times', '3 times', '4 times', '5 times or more']


sort_order4 = ['I did not consume 5 or more whole drinks of alcohol in a row over the past 30 days', '1 day',
 '2 days',
 '3 to 5 days',
               '6 to 9 days',
 '10 or more days']

sort_order5 = ['Never','1 to 2 times per year',
'Once a month',
'Twice a month',
               'Once a week',
'Twice a week',
'Daily']

perc_order = ['0%',
 '1-10%',
 '11-20%',
 '21-30%',
 '31-40%',
 '41-50%',
 '51-60%',
 '61-70%',
 '71-80%',
 '81-90%',
 '91-100%']

YN_order2 = [ 'Yes', 'No']

sort_order6 = ['Drinking is never a good thing to do.',
 'Drinking is okay, but a person should not get drunk.',
 "Occasionally getting drunk is okay as long as it DOESN'T interfere with academics or other responsibilities",
 'Occasionally getting drunk is okay EVEN IF IT DOES interfere with academics or other responsibilities',
              'Frequently getting drunk is okay if an individual wants to do that']


sort_order7 =['Strongly disapprove', 
'Disapprove', 'Neither disapprove nor approve', 'Approve' ,'Strongly approve']


sort_order8 =[
"No risk" ,
"Slight risk" ,
"Moderate risk" ,
"Great risk"]

sort_order9 =['Strongly disapprove', 
'Disapprove',  'Approve' ,'Strongly approve']

sort_order10 =['Strongly disagree', 
'Disagree',  'Agree' ,'Strongly agree']

sort_order11 =['Never', 'Once', 'More than once']

sort_order12 = ['0 days' ,
'1 day' ,
'2 days' ,
'3 to 5 days', 
'6 to 9 days', 
'10 or more days']


sort_order13 = ['Too lenient', 
'About right', 
'Too strict', 
"Don't know"]

sort_order14 = ['Liked a lot', 'Kind of liked',  'Did not like', 'Hated' ]

sort_order15 = [ 'Never','1-2 days a week', '3-5 days a week', '6-7 days a week']

# 'Never' is a real answer option (and usually the most chosen) — omitting
# it dropped those students from every classmates-pressured table.
sort_order16 =['Never', 'Rarely'  ,'Sometimes'  ,'Most of the time' ,'Always']


### sort_order17 removed (only used by K10 Kessler tables)

pot_age = [
    '8 years old or younger',
'9 to 10 years old',
'11 to 12 years old',
 '13 to 14 years old',
 '15 to 16 years old',
 '17 to 18 years old',
'19 years or older',
'I have never used marijuana at all']

drinkslistnodash = clean['how_many_seniors'].dropna().unique().tolist()


sort_order18 = ['Never',              
'1 time',              
'2 times',
'3 or more times' ]


# In[20]:


##############################
### Questions lists - Questions with the same answer list like all YN questions so can batch process
##############################



never123Qs = ['felt_or_tried_to_cut',
 'drinking_or_drug_problem',
 'felt_guilty_about_your_drinking',
 'needed_a_drink_first_thing',
 'used_alcohol_by_yourself',
 'get_drunk',
 'do_something_you_later_felt_sorry_for',
 'friend_fight_argument_bad_feelsing',
 'family_fight_argument_bad_feelsing',
 'sexual_experience',
 'perform_poorly_school_work',
 'trouble_concentrating',
 'sick_vomit_hangover',
 'drive',
 'ride',
 'trouble_school_home',
 'hang_out_bad_influence',
 'forget_what_happened',
 'pass_out',
 'get_hurt',
 'drink_before_school',
 'try_other_substances',
 'drank_more']


YNQs =['use_alcohol_12',
 'use_cigarettes_12',
 'use_vapes_12',
 'use_smokeless_tobacco_12',
 'use_pot_12',
 'use_ecstasy_12',
 'use_heroin_12',
 'use_narcotics_12',
 'use_amphetamines_12',
 'use_sedatives_12',
 'use_methamphetamine_12',
 'use_cocaine_12',
 'use_lsd_12',
 'use_hallucinogen_12',
 'use_inhalants_12',
 'use_steroids_12',
 'use_cold_medicine_12',
 'use_alcohol_30',
 'use_cigarettes_30',
 'use_vapes_30',
 'use_smokeless_tobacco_30',
 'use_pot_30',
 'use_ecstasy_30',
 'use_heroin_30',
 'use_narcotics_30',
 'use_amphetamines_30',
 'use_sedatives_30',
 'use_methamphetamine_30',
 'use_cocaine_30',
 'use_lsd_30',
 'use_hallucinogen_30',
 'use_inhalants_30',
 'use_steroids_30',
 'use_cold_medicine_30']


moreNver123Q = ['taking_concerned_other',
 'taking_care_drank_too_much_other',
 'taking_care_consumed_too_much_other',
 'emergency_medical_care_other',
 'in_trouble_other',
 'feeling_insulted_other',
 'verbal_argument_other',
 'pushed_hit_other',
 'sexual_experience_other',
 'witness_vomit_other',
 'unsafe_other',
 'sleep_difficulty_other',
 'homework_difficulty_other']



ApproveDisapproveQ = [ 'one_2_friend',
 'three_4_friend',
 'five_friend',
 'got_drunk_friend']


RiskQs = [ 'harm_1_2_drinks',
 'harm_1_2_daily',
 'harm_3_4_daily',
 'harm_5']

ApproveDisapprove2_q = ['pot_friend_once',
 'pot_friend_occasion',
 'pot_friend_regular']

AgreeDisagree_q = [ 'social_atmosphere',
 'pressured',
 'easy_to_make_friends',
 'students_assume',
 'own_decisions',
 'respect',
 'fit_in',
 'light_drinkers',
 'binge_okay']

Never_once_qs = [ 'discussion_friend',
 'discussion_guardian',
 'discussion_teacher']


freq_qs = [ 'drink_how_often_your_grade',
 'drink_how_often_seniors']

one_drink_days_qs = [ 'one_drink_30_days_your_school',
 'one_drink_30_days_your_grade',
 'one_drink_30_days_seniors']

pot_freq_q = [ 'marijuana_your_grade',
 'marijuana_seniors']



too_lenient_qs = ['thoughts_alcohol_policies',
 'thoughts_enforcement',
 'thoughts_other_alcohol_policies',
 'thoughts_other_enforcement']

AgreeDisagree2_q = [ 'valued',
 'academics_to_me',
 'teachers',
 'sincere_effort',
 'care',
 'academics_to_others']

AgreeDisagree3_q = [ 'parent_interest',
 'clear_rules',
 'call',
 'come_home',
 'clear_rules_alcohol',
 'not_home']

binarylist = ['friend_problem',
 'talking_parents',
 'talking_parents_about_friend']

Rarely_Qs = ['classmates_pressured',
 'parents_present_visible']

AgreeDisagree4_q = [ 'academic_stress',
 'parent_stress',
 'college_stress',
 'overlooked_stress',
 'expectations_stress',
 'boredom_stress',
 'fitting_in_stress']

YNQ2s =[ 'coping_alcohol',
 'coping_cigarettes',
 'coping_vape',
 'coping_smokeless_tobacco',
 'coping_marijuana',
 'coping_molly',
 'coping_heroin',
 'coping_narcotics',
 'coping_amphetamines',
 'coping_sedatives',
 'coping_meth',
 'coping_cocaine',
 'coping_lsd',
 'coping_hallucinogens',
 'coping_inhalant',
 'coping_steroid',
 'coping_cold_med']

### Kessler_Qs removed (K10 not on this survey)

drinks0Qs = ['how_many_30_days_your_grade', 'how_many_30_days_your_seniors']


# In[ ]:





# In[21]:


# def new_percenter(df, condition, gradecol):
#     if gradecol in df.columns:
#     ## how many have ever drank? 100 - %who have not, etc 
#         val = 100 - (df[condition][gradecol].str.split('%').str[0]).astype(float)
#         return val
#     else:
#         return ''


    
    


# In[22]:


#(age_at_first_grade[age_at_first_grade['Age']!='I have never had a whole drink of alcohol']['12th (%)'].str.split('%').str[0]).astype(float).sum()

#never_by_age(age_at_first_grade, '7th (%)')


# In[23]:


##############################
### Questions
##############################

## What is your sex? 
sexdf = create_count_and_percent(clean, 'sex', 'Sex', 'No', 'No')


sex_by_grade = by_grader_count_and_percenter('sex','grade_level', grades, ['Female', 'Male', 'Prefer not to say'], 'Sex' )

sex_by_risk = by_grader_count_and_percenter('sex','drinker_risk', drinkinglist, ['Female', 'Male', 'Prefer not to say'], 'Sex' )


#What grade are you in?

gradedf = create_count_and_percent(clean, 'grade_level', 'Grade Level', 'Yes', 'No')

grade_by_risk = by_grader_count_and_percenter('grade_level','drinker_risk', drinkinglist, grades, 'Grade' )

############### variable for report

grade_by_risk10_12 = by_grader_count_and_percenter('grade_level','drinker_risk_10_12', drinkinglist, grades, 'Frequency')
highrisk_count_10_12 = int(grade_by_risk10_12['High-risk (N)'].sum())
lowrisk_count_10_12 =int(grade_by_risk10_12['Low-risk (N)'].sum())
norisk_count_10_12 =int(grade_by_risk10_12['Non-drinker (N)'].sum())





avg_age_highrisk = round(clean[clean['drinker_risk_10_12']=='High-risk']['agenum'].mean(),1)
avg_age_lowrisk = round(clean[clean['drinker_risk_10_12']=='Low-risk']['agenum'].mean(),1)
avg_age_norisk = round(clean[clean['drinker_risk_10_12']=='Non-drinker']['agenum'].mean(),1)



#If you drink, how old were you when you had your first whole drink of alcohol?

age_at_first_drinkdf = create_count_and_percent(clean, 'age_at_first_drink', 'Age At First Drink', 'Yes', 'No')
age_at_first_drink_output = clean.groupby(['age_at_first_drink', 'grade_level']).count()['Token'].reset_index()
age_at_first_drink_output.columns = ['age_at_first_drink', 'grade_level', 'count']

age_at_first_drink_output = age_at_first_drink_output.pivot(index='age_at_first_drink', columns='grade_level', values='count').reset_index().fillna(0)

age_at_first_grade = by_grader_count_and_percenter('age_at_first_drink','grade_level', grades, age_at_first_drinkdf['Age At First Drink'].reset_index(drop=True).tolist(), 'Age')
age_at_first_by_risk = by_grader_count_and_percenter('age_at_first_drink','drinker_risk', drinkinglist, age_at_first_drinkdf['Age At First Drink'].reset_index(drop=True).tolist(), 'Age')


########### Report variable

## 24 {{School}} students 2 in 9th grade,  7 in 10th grade, 4 in 11th grade, and 11 in 12th grade could be considered higher-risk drinkers, having reported consuming 5 or more drinks of alcohol on 1 or more days within the past 30 days.
totalhighrisk = int(grade_by_risk['High-risk (N)'].sum())
# CORRECTED: Log validation - compare all-grade vs 10-12th-only risk counts
logger.info(f"Higher-risk drinkers (all grades via drinker_risk): {totalhighrisk}")
logger.info(f"Higher-risk drinkers (10-12th only via drinker_risk_10_12): {highrisk_count_10_12}")
if totalhighrisk != highrisk_count_10_12:
    logger.warning(f"Higher-risk count discrepancy: all-grade={totalhighrisk} vs 10-12th={highrisk_count_10_12}. "
                   "The exec summary uses the all-grade count. Verify this is intentional.")

highriskb1 =[]

for index, row in grade_by_risk.iterrows():
    if row['High-risk (N)'] > 0:
        highriskb1.append(f"{int(row['High-risk (N)'])} in {row['Grade']} grade")
    
if len(highriskb1) == 0:
    highriskbullet1 = ''
elif len(highriskb1) == 1:
    highriskbullet1 = highriskb1[0]
else:
    highriskbullet1 = ', '.join(highriskb1[:-1]) + ', and ' + highriskb1[-1]

# Build dynamic comparison text for exec summary
# Replaces hardcoded template text like "nearly half are in 12th grade (11) as in 11th grade (4)"
highrisk_by_grade = {}
for index, row in grade_by_risk.iterrows():
    if row['High-risk (N)'] > 0:
        highrisk_by_grade[row['Grade']] = int(row['High-risk (N)'])

def _count_by_grade_sentence(counts):
    """C11: name the count for EVERY grade in grade order, e.g. "1 is in 10th grade,
    3 are in 11th grade, and 3 are in 12th grade." No "top two" comparison."""
    _parts = [f"{n} {'is' if n == 1 else 'are'} in {g} grade"
              for g, n in sorted(counts.items(), key=lambda kv: _grade_num(kv[0]))]
    if not _parts:
        return ''
    if len(_parts) == 1:
        return f"All of these students are in {sorted(counts)[0]} grade."
    return 'Of these students, ' + _join_and(_parts) + '.'

highrisk_comparison = _count_by_grade_sentence(highrisk_by_grade)

totalstudents = age_at_first_drinkdf['Count (N)'].sum()
perc_high_riskall = round((totalhighrisk/totalstudents)*100,1)
# B2: 25/222 is about 1 in 9 (int() truncated it to 1 in 8)
perc_high_riskallone_in = round(totalstudents / totalhighrisk) if totalhighrisk else ''

# How were your grades overall during the last school year?

grade_overalldf = create_count_and_percent_0check(clean, 'grades_overall', 'Grades Overall', 'No', 'No', listofageoptions)
grades_overall_bygrade = by_grader_count_and_percenter('grades_overall','grade_level', grades, listofageoptions, "Grades")
grades_overall_by_risk = by_grader_count_and_percenter('grades_overall','drinker_risk', drinkinglist, listofageoptions, 'Grades' )
all_risk_options_grades=risk_by_subset_pot_andall_1012th('grades_overall', drinkinglist, listofageoptions, 'Grades' )
all_risk_options_grades['Yearly or More Frequent Marijuana Use'] = round(all_risk_options_grades['Yearly or More Frequent Marijuana Use'].str.split('%').str[0].astype(float), 1).astype(str) +'%'

### report variable
lowriskdrinkers =age_at_first_by_risk['Low-risk (N)'].sum()
noniskdrinkers =age_at_first_by_risk['Non-drinker (N)'].sum()

percent8thgraders_ever_drank = new_percenter( age_at_first_grade, (age_at_first_grade['Age']== 'I have never had a whole drink of alcohol'), '8th (%)')
percent10thgraders_ever_drank = new_percenter( age_at_first_grade, (age_at_first_grade['Age']== 'I have never had a whole drink of alcohol'), '10th (%)')
percent12thgraders_ever_drank = new_percenter( age_at_first_grade, (age_at_first_grade['Age']== 'I have never had a whole drink of alcohol'), '12th (%)')

percent_never_drank_allgrades = (age_at_first_drinkdf[age_at_first_drinkdf['Age At First Drink']=='I have never had a whole drink of alcohol']['Percentage (%)'].str.split('%').str[0]).astype(float)
percent_never_drank_11th = (age_at_first_grade[age_at_first_grade['Age']=='I have never had a whole drink of alcohol']['11th (%)'].str.split('%').str[0]).astype(float)
percent_never_drank_12th= (age_at_first_grade[age_at_first_grade['Age']=='I have never had a whole drink of alcohol']['12th (%)'].str.split('%').str[0]).astype(float)


lifetime7th = ever_by_age(age_at_first_grade, '7th (%)')
lifetime12th = ever_by_age(age_at_first_grade, '12th (%)')


percent8thgraders_Never_drank =never_by_age(age_at_first_grade,'8th (%)')
percent9thgraders_Never_drank =never_by_age(age_at_first_grade,'9th (%)')
percent10thgraders_Never_drank =never_by_age(age_at_first_grade,'10th (%)')
percent11thgraders_Never_drank =never_by_age(age_at_first_grade,'11th (%)')
percent12thgraders_Never_drank =never_by_age(age_at_first_grade,'12th (%)')

#percent12thgraders_Never_drank = round(percent12thgraders_Never_drank,1)
#percent11thgraders_Never_drank = round(percent11thgraders_Never_drank,1)
percent10thgraders_ever_drank = round(percent10thgraders_ever_drank,1)
percent12thgraders_ever_drank = round(percent12thgraders_ever_drank,1)
percent_never_drank_11th = round(percent_never_drank_11th, 1)
percent_never_drank_12th = round(percent_never_drank_12th, 1)

def _to_val(v):
    """Safely extract scalar from Series, float, or return 0 for empty string."""
    if isinstance(v, str):
        return 0
    try:
        return float(v.iloc[0]) if hasattr(v, 'iloc') else float(v)
    except Exception:
        return 0

_8th_val = _to_val(percent8thgraders_ever_drank) if has_8th else None
if '8th' in clean['grade_level'].values:
    newrow = ['Total', _to_val(percent8thgraders_ever_drank), _to_val(percent10thgraders_ever_drank), _to_val(percent12thgraders_ever_drank)]
    everdrank_grade_sex = by_grader_count_and_sex_percenter('age_at_first_drink', ['grade_level', 'sex'], ['8th', '10th', '12th'],
                                      ['sex', '8th', '10th', '12th'], newrow)
else:
    newrow = ['Total', _to_val(percent8thgraders_ever_drank), _to_val(percent10thgraders_ever_drank), _to_val(percent12thgraders_ever_drank)]
    everdrank_grade_sex = by_grader_count_and_sex_percenter('age_at_first_drink', ['grade_level', 'sex'], ['8th', '10th', '12th'],
                                      ['sex', '8th', '10th', '12th'], newrow)


# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[24]:


#How old are you? 
aggregates, agedf =  create_count_and_percent(clean, 'agenum', 'Age', 'No', 'Yes')
agedf['Age'] = agedf['Age'].astype(int)
aggregates['Values'] = round(aggregates['Values'], 1)

# Where did you first use alcohol?
# Note: Only students who reported ever drinking a whole drink of alcohol were asked this question
where_firstdf= create_count_and_percent_0check(clean, 'where_first_use_alcohol', 'Location', 'No', 'No', alllocations)
where_firstdf =where_firstdf.sort_values('Count (N)', ascending= False)
where_first_grade = by_grader_count_and_percenter('where_first_use_alcohol','grade_level', grades, alllocations, 'Location')
where_first_by_risk = by_grader_count_and_percenter('where_first_use_alcohol','drinker_risk', drinkinglist, alllocations, 'Location')


######################################################################
########## report variable

totalsbygrade =age_at_first_grade[[k for k in age_at_first_grade.columns if '(N)' in k]].sum().reset_index()
# strip() matters: '8th (N)'.split('(')[0] is '8th ' with a trailing space,
# which silently broke every == comparison against plain grade names.
totalsbygrade['grade_level'] = totalsbygrade['index'].str.split('(').str[0].str.strip()
totalsbygrade = totalsbygrade[[0, 'grade_level']]
totalsbygrade.columns = ['totalingrade', 'grade_level']

# Select 10th-12th by name: totalsbygrade always has all five grade rows
# (padded), so positional [3:] selected only 11th-12th.
onestudentin1012outof100 =round((1/int(totalsbygrade[totalsbygrade['grade_level'].isin(['10th', '11th', '12th'])]['totalingrade'].sum()))*100, 1)

totalsbygrade_andsex =sex_by_grade[['Sex'] +[k for k in age_at_first_grade.columns if '(N)' in k]].groupby('Sex').sum().T.reset_index()
totalsbygrade_andsex['grade_level'] = totalsbygrade_andsex['index'].str.split('(').str[0].str.strip()
totalsbygrade_andsex = totalsbygrade_andsex[['Male','Female' ,'grade_level']]

totalsbygrade_andsex =totalsbygrade_andsex.melt(id_vars=['grade_level'], value_vars=['Male', 'Female'] )
totalsbygrade_andsex.columns =['grade_level', 'sex', 'total']

# CORRECTED: Use fuzzy matching for location string to handle apostrophe variants
# Original had hidden \x19 char that didn't match CSV data
_friends_mask = fuzzy_location_match(where_firstdf, 'Location', "At a friend's house")
if _friends_mask.any():
    _friends_perc = where_firstdf[_friends_mask]['Percentage (%)'].str.split('%').str[0].astype(float)
    first_drink_at_friends = round(_friends_perc.iloc[0], 1)
    logger.info(f"first_drink_at_friends = {first_drink_at_friends}%")
else:
    first_drink_at_friends = 0.0
    logger.warning("Could not find 'At a friend\\'s house' in where_firstdf locations. Check CSV values.")

# How often do you have at least one whole drink of alcohol?
# Note: Only students who reported ever drinking a whole drink of alcohol were asked this question

often_one_drink = create_count_and_percent_0check(clean, 'how_often_1_drink', 'Frequency', 'No', 'No', sort_order1)
often_one_drink = sort_by_customlist(often_one_drink, 'Frequency', sort_order1)

often_one_grade = by_grader_count_and_percenter('how_often_1_drink','grade_level', grades, sort_order1, 'Frequency')
often_one_by_risk = by_grader_count_and_percenter('how_often_1_drink','drinker_risk', drinkinglist, sort_order1, 'Frequency')

########## report variable
often_one_by_risk10_12 = by_grader_count_and_percenter('how_often_1_drink','drinker_risk_10_12', drinkinglist, sort_order1, 'Frequency')
perc_1012_highrisk_twiceormore = (often_one_by_risk10_12[3:]['High-risk (N)'].sum()/ often_one_by_risk10_12['High-risk (N)'].sum()) * 100
perc_1012_highrisk_twiceormore = round(perc_1012_highrisk_twiceormore, 1)

# CORRECTED: Round percentage sum
# B4: from counts (base = lifetime drinkers who answered how often they drink)
_ood_counts = pd.to_numeric(often_one_drink['Count (N)'], errors='coerce').fillna(0)
drinklessthan12timesperyear = pct(_ood_counts[often_one_drink['Frequency'].isin(['Never', '1-2 times per year'])].sum(), _ood_counts.sum(),
                                  'drink 1-2 times per year or less', 'lifetime drinkers who answered how often they drink')
# B2: "once a month OR MORE", actual on the same base as the perceived figure
# (all students in the grade, counting never-drinkers as "Never").
def _actual_freq_share(grade, rows):
    _ncol = f'{grade} (N)'
    if _ncol not in often_one_grade.columns:
        return ''
    _n = pd.to_numeric(often_one_grade.loc[often_one_grade['Frequency'].isin(rows), _ncol], errors='coerce').fillna(0).sum()
    return pct(_n, grade_size(grade), f'{grade} actual drinking {rows[0]} or more', f'all {grade} graders')
_monthly_plus = ['Once a month', 'Twice a month', 'Once a week', 'Twice a week', 'Daily']
howoften10th_onceamonthperc = fmt_pct(_actual_freq_share('10th', _monthly_plus))
eleventhmorethantwice = _actual_freq_share('11th', ['Twice a month', 'Once a week', 'Twice a week', 'Daily'])
twlthtwiceaweek_perc = round(often_one_grade[often_one_grade['Frequency']=='Twice a week']['12th (N)']/  totalsbygrade[totalsbygrade['grade_level'].str.contains('12th')]['totalingrade'].item() * 100, 1)
twlthtwiceaweek_perc = twlthtwiceaweek_perc.item()

# How many whole drinks of alcohol, if any, do you usually have when you drink?
howmany_df = create_count_and_percent_0check(clean, 'how_many_drinks', 'Amount', 'No', 'No', amountlist)
howmany_grade = by_grader_count_and_percenter('how_many_drinks','grade_level', grades, amountlist, 'Amount')
howmany_by_risk = by_grader_count_and_percenter('how_many_drinks','drinker_risk', drinkinglist,amountlist, 'Amount')
howmany_by_risk10_12 =by_grader_count_and_percenter('how_many_drinks','drinker_risk_10_12', drinkinglist,amountlist, 'Amount')
high_risk_morethan5 = fmt_pct(count_pct_from_table(howmany_by_risk10_12, howmany_by_risk10_12['Amount'].str.contains('5'), 'High-risk (%)'))

############# report variables
less_than2_peryear = drinklessthan12timesperyear   # B4/B5: template now says "1-2 times per year or less"
howmanyhighrisk_perc = count_pct_from_table(howmany_by_risk, howmany_by_risk['Amount']=='3-4 drinks', 'High-risk (%)')
# B6: "typically have 5 or more drinks" = how_many_drinks, NOT the 5+-in-a-row
# item (the template used howmanyhighrisk5_perc = 23/45 here by mistake).
howmanyhighrisk5typ_perc = count_pct_from_table(howmany_by_risk, howmany_by_risk['Amount']=='5 or more drinks', 'High-risk (%)')
#howmanyhighrisk5_perc = howmany_by_risk[howmany_by_risk['Amount']=='5 or more drinks']['High-risk (%)'].str.split('%').str[0].astype(float).sum()
#howmanyhighrisk5_count = howmany_by_risk[howmany_by_risk['Amount']=='5 or more drinks']['High-risk (N)'].astype(float).sum()
#howmanyhighrisk5_count = int(howmanyhighrisk5_count)



# Have you had at least 1 whole drink of alcohol during the PAST 12 MONTHS?
one_drink12df = create_count_and_percent_0check(clean, 'one_drink_past_12', 'Response', 'No', 'No', YN_order2)
onewhole_grade = by_grader_count_and_percenter('one_drink_past_12','grade_level', grades, YN_order2, 'Amount')
onewhole_by_risk = by_grader_count_and_percenter('one_drink_past_12','drinker_risk', drinkinglist, YN_order2, 'Amount')

########## report variable, what percent have ever drank 12 months



# Use 8th grade as lowest when no 7th graders
if '7th' in clean['grade_level'].values:
    twelve_7th = ever_by_age_yes(onewhole_grade,'7th (%)')
elif '8th' in clean['grade_level'].values:
    twelve_7th = ever_by_age_yes(onewhole_grade,'8th (%)')
else:
    twelve_7th = _wrap_scalar('')
twelve_12th = ever_by_age_yes(onewhole_grade,'12th (%)')





total8th_last12_perc =percent_of_binary_where_total_elsewhere( onewhole_grade[onewhole_grade['Amount']=='Yes'], age_at_first_grade, '8th')
tota10th_last12_perc =percent_of_binary_where_total_elsewhere( onewhole_grade[onewhole_grade['Amount']=='Yes'], age_at_first_grade, '10th')
total12th_last12_perc =percent_of_binary_where_total_elsewhere( onewhole_grade[onewhole_grade['Amount']=='Yes'], age_at_first_grade, '12th')

total8th_last12_perc =percent_of_binary_where_total_elsewhere( onewhole_grade[onewhole_grade['Amount']=='Yes'], age_at_first_grade, '8th')
tota10th_last12_perc =percent_of_binary_where_total_elsewhere( onewhole_grade[onewhole_grade['Amount']=='Yes'], age_at_first_grade, '10th')
total12th_last12_perc =percent_of_binary_where_total_elsewhere( onewhole_grade[onewhole_grade['Amount']=='Yes'], age_at_first_grade, '12th')

noyes = clean[['one_drink_past_12', 'grade_level', 'sex', 'Token']].groupby(['one_drink_past_12','grade_level', 'sex']).count().reset_index()
# Keep only affirmative answers. This question's options are Yes/No, so 'No'
# must be excluded too — filtering only "I did not" kept the No rows and made
# the by-sex past-12-months table show lifetime numbers instead.
noyes =noyes[~noyes['one_drink_past_12'].str.contains('I did not')]
noyes =noyes[noyes['one_drink_past_12'] != 'No']
noyes =noyes[['grade_level', 'sex', 'Token']]
noyes.columns = ['grade_level', 'sex', 'countyes']
noyes=noyes.groupby(['grade_level', 'sex']).sum()
perctable = pd.merge(noyes, totalsbygrade_andsex, on=['grade_level', 'sex'], how='left')
perctable['perc'] = round((perctable['countyes'] / perctable['total'])*100, 1)
perctable =perctable.pivot(index='sex', columns='grade_level', values='perc').reset_index()
# A (grade, sex) cell with zero drinkers is absent from the pivot and would
# render as '-'; for grades that took the survey it must show 0.0% instead.
for _g in [g for g in ['8th', '9th', '10th', '11th', '12th'] if g in clean['grade_level'].values]:
    if _g not in perctable.columns:
        perctable[_g] = 0.0
_gradecols = [c for c in perctable.columns if c != 'sex']
perctable[_gradecols] = perctable[_gradecols].fillna(0.0)
# ...but only for cohorts that exist: a grade with zero respondents of a
# given sex must stay blank ('-'), not claim 0% use.
_cohort_sizes = clean.groupby(['grade_level', 'sex']).size()
for _ri in perctable.index:
    for _g in _gradecols:
        if (_g, perctable.loc[_ri, 'sex']) not in _cohort_sizes.index:
            perctable.loc[_ri, _g] = float('nan')

# CORRECTED: Use _safe_item() - .item() fails on plain strings (e.g. '' for missing grades)
if '8th' in perctable.columns:
    perctable =perctable[ ['sex', '8th', '10th', '12th']]
    perctable.loc[-1]  = ['Total', _safe_item(total8th_last12_perc), _safe_item(tota10th_last12_perc), _safe_item(total12th_last12_perc)]
    perctable= perctable.set_index('sex').loc[['Total', 'Male', 'Female']].reset_index()
    alcohol_past_12monthbysex =perctable
else:
    perctable =perctable[ ['sex', '10th', '12th']]
    perctable.loc[-1]  = ['Total', _safe_item(tota10th_last12_perc), _safe_item(total12th_last12_perc)]
    perctable= perctable.set_index('sex').loc[['Total', 'Male', 'Female']].reset_index()
    alcohol_past_12monthbysex =perctable
    alcohol_past_12monthbysex['8th'] = ''
    alcohol_past_12monthbysex = alcohol_past_12monthbysex[ ['sex', '8th', '10th', '12th']]


# In[ ]:





# In[ ]:





# In[25]:


#clean['do_something_you_later_felt_sorry_for'].value_counts().reset_index()['index'].tolist()


# In[26]:


for col in never123Qs:
    clean[col] = clean[col].str.rstrip()


# In[27]:


neverQ_answers = multi_questioner(never123Qs, sort_order2, 'Response' )
neverQ_answers_grade = multi_by_grader( never123Qs,'grade_level', grades, sort_order2, 'Response' )
neverQ_answers_by_risk = multi_by_grader( never123Qs,'drinker_risk', drinkinglist, sort_order2, 'Response' )


# In[ ]:





# In[28]:


neverQ_answers_by_risk_subs = []

# B6: consequence labels are keyed by ITEM ID, never by list position (the old
# positional list put "Drank more than you intended" five rows too early, so
# the last five rows of the Consequences table showed the wrong values).
# B7: third-person wording only ("they", not "you").
CONSEQUENCE_LABELS = {
    'felt_or_tried_to_cut': "Felt need or tried to cut down on drinking/other drug use",
    'drinking_or_drug_problem': "Thought they had a drinking/drug problem",
    'felt_guilty_about_your_drinking': "Felt guilty about their drinking/drug use",
    'needed_a_drink_first_thing': "Needed a drink or other drug first thing in the morning",
    'used_alcohol_by_yourself': "Used alcohol or other drugs by themselves",
    'get_drunk': "Got drunk",
    'do_something_you_later_felt_sorry_for': "Did something they later felt sorry for",
    'friend_fight_argument_bad_feelsing': "Had a fight/argument/bad feelings with a friend/boyfriend/girlfriend",
    'family_fight_argument_bad_feelsing': "Had a fight/argument/bad feelings with a family member",
    'sexual_experience': "Had a sexual experience that they later wished didn't happen",
    'perform_poorly_school_work': "Missed or performed poorly at school/work/athletic event",
    'trouble_concentrating': "Had trouble concentrating on homework/class work/other tasks",
    'sick_vomit_hangover': "Felt really sick, vomited, or had a hangover",
    'drive': "Drove a vehicle after drinking",
    'ride': "Rode as a passenger in a vehicle after the driver had been drinking",
    'trouble_school_home': "Got in trouble at home/school",
    'hang_out_bad_influence': "Got involved with people they think are a bad influence",
    'forget_what_happened': "Forgot what happened while they were drinking",
    'pass_out': "Passed out because of drinking",
    'get_hurt': "Got hurt or injured",
    'drink_before_school': "Had something to drink right before/during school hours",
    'try_other_substances': "Tried other substances they wouldn't have tried otherwise",
    'drank_more': "Drank more than they intended",
}

def group_share(Q, positive, group_col='drinker_risk_10_12', group='High-risk', name=None):
    """B2/B4: % of a group (who answered Q) whose answer is positive, from counts."""
    sub = clean[(clean[group_col] == group) & clean[Q].notna()]
    n = int(positive(sub[Q]).sum())
    return pct(n, len(sub), name, f'{group} ({group_col}) who answered {Q}')

_not_never = lambda ser: ~ser.astype(str).eq('Never')

for Q in never123Qs:
    print(Q)
    df =risk_by_subset_pot_andall_1012th(Q,drinkinglist, sort_order2, 'Response' )
    df = df[df['Response']!='Never']
    # B4: the Total row ("1 or more times") is computed from COUNTS, not by
    # adding the already-rounded row percentages (43.4 -> 43.5, 8.6 -> 8.7 ...).
    TOTALS = {
        'Yearly or More Frequent Marijuana Use': pct(((clean['yearlypot'].notna()) & clean[Q].notna() & _not_never(clean[Q])).sum(),
                                                     ((clean['yearlypot'].notna()) & clean[Q].notna()).sum()),
        'High-risk (%)': group_share(Q, _not_never, group='High-risk', name=f'consequence {Q}: higher-risk'),
        'Low-risk (%)': group_share(Q, _not_never, group='Low-risk', name=f'consequence {Q}: lower-risk'),
        'Non-drinker (%)': group_share(Q, _not_never, group='Non-drinker'),
        '10-12th Graders': pct((clean['grade_level'].isin(BODY_GROUP_GRADES) & clean[Q].notna() & _not_never(clean[Q])).sum(),
                               (clean['grade_level'].isin(BODY_GROUP_GRADES) & clean[Q].notna()).sum()),
    }
    TOTALS = {k: (0.0 if v == '' else v) for k, v in TOTALS.items()}
    df = pd.concat([df, pd.DataFrame({'Response': 'Total', **TOTALS}, index=[4])], ignore_index=True)
    df['Q'] = Q
    neverQ_answers_by_risk_subs.append(df)

highlowrisk_consequences=pd.concat(neverQ_answers_by_risk_subs)

highlowrisk_consequence_sub =highlowrisk_consequences[highlowrisk_consequences['Response']=='Total'][['Q', 'High-risk (%)', 'Low-risk (%)']].copy()
highlowrisk_consequence_sub['Qtext'] = highlowrisk_consequence_sub['Q'].map(CONSEQUENCE_LABELS)
_unlabeled = highlowrisk_consequence_sub[highlowrisk_consequence_sub['Qtext'].isna()]['Q'].tolist()
if _unlabeled:
    raise ValueError(f"B6: no Consequences label for item(s) {_unlabeled}. Add them to CONSEQUENCE_LABELS.")

len_highlowrisk_consequence_sub = len(highlowrisk_consequence_sub)

for col in ["High-risk (%)","Low-risk (%)"]:
    highlowrisk_consequence_sub[col] = pd.to_numeric(highlowrisk_consequence_sub[col]).apply(lambda v: f"{round(float(v), 1)}%")

highlowrisk_consequences =pd.merge( highlowrisk_consequence_sub[['Q', 'Qtext']], highlowrisk_consequences, on='Q')

highriskconsquencestotals = highlowrisk_consequences[highlowrisk_consequences['Response']=='Total']

highriskbulletsssss = []
for index, row in highriskconsquencestotals.iterrows():
    highriskbulletsssss.append(f'{round(float(row["High-risk (%)"]),1)}% of "higher-risk" drinkers {row["Qtext"][0].lower() + row["Qtext"][1:]}')

lowriskbulletsssss = []
for index, row in highriskconsquencestotals.iterrows():
    lowriskbulletsssss.append(f'{round(float(row["Low-risk (%)"]),1)}% of "lower-risk" drinkers {row["Qtext"][0].lower() + row["Qtext"][1:]}')


highlowrisk_consequence_sub['High-risk (%)']= pd.to_numeric(highlowrisk_consequence_sub['High-risk (%)'].str.split('%').str[0])
highlowrisk_consequence_sub['Low-risk (%)']= pd.to_numeric(highlowrisk_consequence_sub['Low-risk (%)'].str.split('%').str[0])


higherthan50perc_consequences = len(highlowrisk_consequence_sub[highlowrisk_consequence_sub['High-risk (%)']>=50])
# "Most frequently reported" = highest percentage: sort DESCENDING on the
# numeric value (default ascending picked the least-reported consequence).
_lr_sorted = highlowrisk_consequence_sub.sort_values(
    'Low-risk (%)', ascending=False,
    key=lambda _s: pd.to_numeric(_s.astype(str).str.replace('%', '', regex=False), errors='coerce'))
mostrpopular_lowriskconsequence =_lr_sorted.head(1)['Qtext'].item()
mostrpopular_lowriskconsequenceperc = _lr_sorted.head(1)['Low-risk (%)'].item()
# C6: counted from the table values: higher-risk >= 2 x lower-risk AND higher-risk > 0
# (0.0% vs 0.0% used to count as "double or greater").
highrisk_twice_risk =highlowrisk_consequence_sub[(highlowrisk_consequence_sub['High-risk (%)']>=(highlowrisk_consequence_sub['Low-risk (%)'])*2)
                                                 & (highlowrisk_consequence_sub['High-risk (%)'] > 0)]
len_highrisk_twice_risk= len(highrisk_twice_risk)


# In[ ]:





# In[29]:


high_risk_drunk10_12 = by_grader_count_and_percenter('get_drunk','drinker_risk_10_12', drinkinglist,sort_order2, 'Response')
# B6: "drunk or felt drunk 3 or more times in the last 30 days" is the
# drunk_30_days item, not the 12-month "Get drunk" consequence item.
high_risk_morethan3drunk = fmt_pct(group_share('drunk_30_days', lambda ser: ser.isin(['3 times', '4 times', '5 times or more']),
                                               name='higher-risk drunk 3+ times in 30 days'))
high_risk_tocut10_12 = by_grader_count_and_percenter('felt_or_tried_to_cut','drinker_risk_10_12', drinkinglist,sort_order2, 'Response')
high_risk10_12_cutdown = group_share('felt_or_tried_to_cut', _not_never, name='higher-risk tried to cut down')   # B4
high_risk_guiltyt10_12 = by_grader_count_and_percenter('felt_guilty_about_your_drinking','drinker_risk_10_12', drinkinglist,sort_order2, 'Response')
high_risk10_12_guilty = group_share('felt_guilty_about_your_drinking', _not_never, name='higher-risk felt guilty')   # B4

#26  students who used alcohol in the past 12 months, reported driving or being a passenger in a car after the driver had been using alcohol. 

drive = neverQ_answers[13]
ride = neverQ_answers[14]
# R6: count STUDENTS who drove OR rode, not answers (a student who did both
# was counted twice: 3 + 16 = 19, but the unique count was 16).
ride_drive = int(((clean['drive'].notna() & _not_never(clean['drive'])) |
                  (clean['ride'].notna() & _not_never(clean['ride']))).sum())
later_sorry = neverQ_answers_by_risk[6]
later_sorry_perc = group_share('do_something_you_later_felt_sorry_for', _not_never, group_col='drinker_risk',
                               name='exec: higher-risk (all grades) did something they felt sorry for')   # B4


## Of students who used alcohol in the past 12 months, 50.7% drank more than they intended on at least 1 occasion. 
morethanintended = (clean[clean['drank_more']!='Never']['drank_more'].value_counts().sum() /clean[ (clean['one_drink_past_12']!='No')]['one_drink_past_12'].value_counts())*100
morethanintended = round(float(morethanintended.iloc[0]), 1)

#During the PAST 30 DAYS, how many days did you have at least one whole drink of alcohol? 

one_drink_30df = create_count_and_percent_0check(clean, 'one_drink_past_30_days', 'Response', 'No', 'No', sort_order3)
one_drink_30df = sort_by_customlist(one_drink_30df, 'Response', sort_order3)


daysatleast1_grade = by_grader_count_and_percenter('one_drink_past_30_days','grade_level', grades, sort_order3, 'Response')
daysatleast1_by_risk = by_grader_count_and_percenter('one_drink_past_30_days','drinker_risk', drinkinglist,  sort_order3, 'Response')


# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[30]:


dranklast30_6ormore =(one_drink_30df[4:]['Percentage (%)'].str.split('%').str[0]).astype(float).sum()
dranklast30_6ormore = round(dranklast30_6ormore,1)

dranklast30_6ormoresum =int(one_drink_30df[4:]['Count (N)'].sum())


if '8th' in clean['grade_level'].values:
    total8th_last30_perc =percent_of_binary_where_total_elsewhere_bysex( daysatleast1_grade[~daysatleast1_grade['Response'].str.contains("I did not")], age_at_first_grade, '8th')
else:
    total8th_last30_perc =''
    
tota10th_last30_perc =percent_of_binary_where_total_elsewhere_bysex(  daysatleast1_grade[~daysatleast1_grade['Response'].str.contains("I did not")], age_at_first_grade, '10th')
total12th_last30_perc =percent_of_binary_where_total_elsewhere_bysex(  daysatleast1_grade[~daysatleast1_grade['Response'].str.contains("I did not")], age_at_first_grade, '12th')




noyes = clean[['one_drink_past_30_days', 'grade_level', 'sex', 'Token']].groupby(['one_drink_past_30_days','grade_level', 'sex']).count().reset_index()
noyes =noyes[~noyes['one_drink_past_30_days'].str.contains('I did not')]
noyes =noyes[['grade_level', 'sex', 'Token']]
noyes.columns = ['grade_level', 'sex', 'countyes']
noyes=noyes.groupby(['grade_level', 'sex']).sum()
perctable = pd.merge(noyes, totalsbygrade_andsex, on=['grade_level', 'sex'], how='left')
perctable['perc'] = round((perctable['countyes'] / perctable['total'])*100, 1)
perctable =perctable.pivot(index='sex', columns='grade_level', values='perc').reset_index()
# A (grade, sex) cell with zero drinkers is absent from the pivot and would
# render as '-'; for grades that took the survey it must show 0.0% instead.
for _g in [g for g in ['8th', '9th', '10th', '11th', '12th'] if g in clean['grade_level'].values]:
    if _g not in perctable.columns:
        perctable[_g] = 0.0
_gradecols = [c for c in perctable.columns if c != 'sex']
perctable[_gradecols] = perctable[_gradecols].fillna(0.0)
# ...but only for cohorts that exist: a grade with zero respondents of a
# given sex must stay blank ('-'), not claim 0% use.
_cohort_sizes = clean.groupby(['grade_level', 'sex']).size()
for _ri in perctable.index:
    for _g in _gradecols:
        if (_g, perctable.loc[_ri, 'sex']) not in _cohort_sizes.index:
            perctable.loc[_ri, _g] = float('nan')
if '8th' in perctable.columns:
    perctable =perctable[ ['sex', '8th', '10th', '12th']]
    perctable.loc[-1]  = ['Total', _safe_item(total8th_last30_perc), _safe_item(tota10th_last30_perc), _safe_item(total12th_last30_perc)]
    perctable= perctable.set_index('sex').loc[['Total', 'Male', 'Female']].reset_index()
    alcohol_past_30daysbysex =perctable

else:
    perctable =perctable[ ['sex', '10th', '12th']]
    perctable.loc[-1]  = ['Total',  _safe_item(tota10th_last30_perc), _safe_item(total12th_last30_perc)]
    perctable= perctable.set_index('sex').loc[['Total', 'Male', 'Female']].reset_index()
    alcohol_past_30daysbysex =perctable
    alcohol_past_30daysbysex['8th'] = ''
    alcohol_past_30daysbysex = alcohol_past_30daysbysex[ ['sex', '8th', '10th', '12th']]



######################################################################
########## report variable
noalcoholpaste30 =one_drink_30df[one_drink_30df['Response'].str.contains('I did not')]['Count (N)'].item() + age_at_first_drinkdf[age_at_first_drinkdf['Age At First Drink'].str.contains("I have never")]['Count (N)'].item()
noalcoholpaste30perc =round((noalcoholpaste30/totalstudents) *100,1)


if '7th' in clean['grade_level'].values:
    thirty_7th = (daysatleast1_grade[daysatleast1_grade['Response']!='I did not drink alcohol during the past 30 days']['7th (%)'].str.split('%').str[0]).astype(float).sum()
elif '8th' in clean['grade_level'].values and '8th (%)' in daysatleast1_grade.columns:
    # Use 8th grade as lowest when no 7th graders
    thirty_7th = round((daysatleast1_grade[daysatleast1_grade['Response']!='I did not drink alcohol during the past 30 days']['8th (%)'].str.split('%').str[0]).astype(float).sum(), 1)
else:
    thirty_7th =''
    
thirty_12th = (daysatleast1_grade[daysatleast1_grade['Response']!='I did not drink alcohol during the past 30 days']['12th (%)'].str.split('%').str[0]).astype(float).sum()
thirty_12th= round(thirty_12th, 1)

twormoredays =round((daysatleast1_grade[(daysatleast1_grade['Response']!= 'I did not drink alcohol during the past 30 days') & (daysatleast1_grade['Response']!= "1 day")]['12th (%)'].str.split('%').str[0]).astype(float).sum(),1)


#What was the greatest number of whole drinks of alcohol you consumed in a row in the PAST 30 DAYS? 
greatest30df = create_count_and_percent_0check(clean, 'drinks_in_a_row_30_days', 'Amount', 'No', 'No', amountlist2)
greatest30_grade = by_grader_count_and_percenter('drinks_in_a_row_30_days','grade_level', grades, amountlist2, 'Amount')
greatest30_by_risk = by_grader_count_and_percenter('drinks_in_a_row_30_days','drinker_risk', drinkinglist,  amountlist2, 'Amount')


#During the PAST 30 DAYS, on how many days did you have five or more whole drinks of alcohol in a row? 
fiveinrowdf = create_count_and_percent_0check(clean, 'days_5_or_more_drinks', 'Response', 'No', 'No', sort_order4)
fiveinrowdf = sort_by_customlist(fiveinrowdf, 'Response', sort_order4)
days_5_or_more_grade = by_grader_count_and_percenter('days_5_or_more_drinks','grade_level', grades, sort_order4, 'Response')
days_5_or_more_by_risk = by_grader_count_and_percenter('days_5_or_more_drinks','drinker_risk', drinkinglist, sort_order4, 'Response')

##### Of the students who reported drinking in the last 30 days, 54.5% (24 students) reported consuming 5 or more drinks in a row on at least 1 of those occasions. 
######################################################################
########## report variable,
five_ormoreperc = (fiveinrowdf[~fiveinrowdf['Response'].str.contains('I did not')]['Percentage (%)'].str.split('%').str[0]).astype(float).sum()
five_ormoreperc = round(five_ormoreperc, 1)

five_ormorecount = (fiveinrowdf[~fiveinrowdf['Response'].str.contains('I did not')]['Count (N)']).astype(float).sum()
five_ormorecount = int(five_ormorecount)

## Of the 24 students who drank 5 or more drinks in a row within the past 30 days,  41.7% (10 students) report having done so on 2 or more days. 
fve_ormore_two_ormore_days_count = (fiveinrowdf[(~fiveinrowdf['Response'].str.contains('I did not')) & ( fiveinrowdf['Response']!='1 day')]['Count (N)']).astype(float).sum()
fve_ormore_two_ormore_days_count= int(fve_ormore_two_ormore_days_count)

fve_ormore_two_ormore_days_perc = (fve_ormore_two_ormore_days_count / five_ormorecount) *100
fve_ormore_two_ormore_days_perc = round(fve_ormore_two_ormore_days_perc, 1)





# In[ ]:





# In[ ]:





# In[ ]:





# In[31]:


#During the PAST 30 DAYS, how many times have you been or felt drunk from drinking alcohol? 
drunk30df = create_count_and_percent_0check(clean, 'drunk_30_days', 'Frequency', 'No', 'No', freq_sort)
drunk_30_grade = by_grader_count_and_percenter('drunk_30_days','grade_level', grades, freq_sort, 'Frequency')
drunk_30_by_risk = by_grader_count_and_percenter('drunk_30_days','drinker_risk', drinkinglist, freq_sort, 'Frequency')

#Where do you do most of your drinking? 
wheremostdf = create_count_and_percent_0check(clean, 'where_do_you_drink', 'Location', 'No', 'No', alllocations)
wheremostdf = wheremostdf.sort_values('Count (N)', ascending=False)


where_do_you_grade = by_grader_count_and_percenter('where_do_you_drink','grade_level', grades, alllocations, 'Location')
where_do_you_by_risk = by_grader_count_and_percenter('where_do_you_drink','drinker_risk', drinkinglist, alllocations, 'Location')


twelfthbar_clubperc = where_do_you_grade[where_do_you_grade['Location'].str.contains('club')]['12th (%)'].item()
eleventhbar_clubperc = where_do_you_grade[where_do_you_grade['Location'].str.contains('club')]['11th (%)'].item()
eleventhbar_clubcount = where_do_you_grade[where_do_you_grade['Location'].str.contains('club')]['11th (N)'].item()
eleventhbar_clubcount = int(eleventhbar_clubcount)

wheredrinkingstrings = []

for location in wheremostdf['Location']:
    # Skip NaN or empty location names
    if pd.isna(location) or str(location).strip() == '':
        continue
    row = wheremostdf[wheremostdf['Location'] == location]
    if len(row) == 0:
        perc = 0
    else:
        perc_val = row['Percentage (%)'].str.split('%').str[0].iloc[0]
        perc = pd.to_numeric(perc_val, errors='coerce')
        if pd.isna(perc):
            perc = 0
    if perc ==0:
        outputstring = f'No students reported doing most of their drinking {location.lower()}.'
        wheredrinkingstrings.append(outputstring)
    if (perc <25) & (perc >0) :
        outputstring = f'A small percentage ({perc}%) do most of their drinking {location.lower()}.'
        wheredrinkingstrings.append(outputstring)        
        
    if (perc <45) & (perc>=25):
        outputstring = f'A minority ({perc}%) do most of their drinking {location.lower()}.'
        wheredrinkingstrings.append(outputstring)
    if (perc >= 45) & (perc <= 55):
        outputstring = f'About half ({perc}%) do most of their drinking {location.lower()}.'
        wheredrinkingstrings.append(outputstring)
    if(perc > 55):
        # CORRECTED: Added missing opening parenthesis before {perc}
        outputstring = f'A majority ({perc}%) do most of their drinking {location.lower()}.'
        wheredrinkingstrings.append(outputstring)
    
    
# If you have ever used marijuana (cannabis, pot, weed, hash, edibles), how old were you the first time you did it? 
agefirstpot = create_count_and_percent_0check(clean, 'age_first_time_pot', 'Age', 'No', 'No', pot_age)
age_first_time_pot_grade = by_grader_count_and_percenter('age_first_time_pot','grade_level', grades, pot_age, 'Age')
age_first_time_pot_by_risk = by_grader_count_and_percenter('age_first_time_pot','drinker_risk', drinkinglist, pot_age, 'Age')


# In[32]:


#How often do you use marijuana (cannabis, pot, weed, hash, edibles)?
howoftenpot = create_count_and_percent_0check(clean, 'how_often_pot', 'Frequency', 'No', 'No', sort_order5)
howoftenpot=sort_by_customlist(howoftenpot, 'Frequency', sort_order5)

how_often_pot_grade = by_grader_count_and_percenter('how_often_pot','grade_level', grades, sort_order5, 'Frequency')
how_often_pot_by_risk = by_grader_count_and_percenter('how_often_pot','drinker_risk', drinkinglist, sort_order5, 'Frequency')

neverpot_all= howoftenpot[howoftenpot['Frequency']=='Never']['Percentage (%)'].item()
if '8th' in clean['grade_level'].values:
    neverpot8th = how_often_pot_grade[how_often_pot_grade['Frequency']=='Never']['8th (%)'].item()
else:
    neverpot8th = ''
neverpot9th = how_often_pot_grade[how_often_pot_grade['Frequency']=='Never']['9th (%)'].item()
neverpot10th = how_often_pot_grade[how_often_pot_grade['Frequency']=='Never']['10th (%)'].item()

oneor2pot10th = how_often_pot_grade[how_often_pot_grade['Frequency']=='1 to 2 times per year']['10th (%)'].item()
oneor2pot11th = how_often_pot_grade[how_often_pot_grade['Frequency']=='1 to 2 times per year']['11th (%)'].item()
oneor2pot12th = how_often_pot_grade[how_often_pot_grade['Frequency']=='1 to 2 times per year']['12th (%)'].item()


neverpot11th = how_often_pot_grade[how_often_pot_grade['Frequency']=='Never']['11th (%)'].item()
neverpot12th = how_often_pot_grade[how_often_pot_grade['Frequency']=='Never']['12th (%)'].item()
twelthpot_moreathan12times =round((how_often_pot_grade[1:]['12th (%)'].str.split('%').str[0]).astype(float).sum(),1)

twelthpot_morethanonceamonth = round((how_often_pot_grade[2:]['12th (%)'].str.split('%').str[0]).astype(float).sum(), 1)

percent8thgraders_ever_pot = new_percenter( age_first_time_pot_grade, (age_first_time_pot_grade['Age'].str.contains('I have never')), '8th (%)')
percent10thgraders_ever_pot = new_percenter( age_first_time_pot_grade, (age_first_time_pot_grade['Age'].str.contains('I have never')), '10th (%)')
percent12thgraders_ever_pot = new_percenter( age_first_time_pot_grade, (age_first_time_pot_grade['Age'].str.contains('I have never')), '12th (%)')

percent_never_pot_allgrades = (agefirstpot[agefirstpot['Age'].str.contains('I have never')]['Percentage (%)'].str.split('%').str[0]).astype(float)
percent_all_typically =( howoftenpot[howoftenpot['Frequency']=='Never']['Percentage (%)'].str.split('%').str[0]).astype(float)

if '8th' in clean['grade_level'].values:
    newrow = ['Total', percent8thgraders_ever_pot.item(), percent10thgraders_ever_pot.item(), percent12thgraders_ever_pot.item()]
    pot_lifetime = by_grader_count_and_sex_percenter('how_often_pot', ['grade_level', 'sex'],
                                                 ['8th', '10th', '12th'],
                                                 ['sex', '8th', '10th', '12th'], newrow)
else:
    newrow = ['Total', percent8thgraders_ever_pot, percent10thgraders_ever_pot.item(), percent12thgraders_ever_pot.item()]
    pot_lifetime = by_grader_count_and_sex_percenter('how_often_pot', ['grade_level', 'sex'],
                                                 ['8th', '10th', '12th'],
                                                 ['sex', '8th', '10th', '12th'], newrow)
    


# In[ ]:





# In[ ]:





# In[33]:


YNQ_answers = multi_questioner(YNQs, YN_order2, 'Response' )
YNQ_answers_grade= multi_by_grader( YNQs, 'grade_level', grades, YN_order2, 'Response' )
YNQ_answers_by_risk= multi_by_grader( YNQs,'drinker_risk', drinkinglist, YN_order2, 'Response' )

# Vaping: percentage who did NOT vape in past 12 months (YNQs index 2 = 'use_vapes_12')
nevervape_all = YNQ_answers[2][YNQ_answers[2]['Response']=='No']['Percentage (%)'].item()

# In[34]:


# YNQs index 1 is 'use_cigarettes_12' (index 2 is vapes — using it here made
# the cigarette abstention stat show the vaping number).
cigarettes12 = YNQ_answers[1]
notcig12monthall = cigarettes12[cigarettes12['Response']=='No']['Percentage (%)'].item()

cigarettes_12_risk10_12=risk_by_subset_pot_andall_1012th('use_cigarettes_12', drinkinglist, YN_order2, 'Response' )
cigarettes_12_risk10_12 =create_high_low_10_non_bullets(cigarettes_12_risk10_12, 'Response')
cigarettes_30_risk10_12=risk_by_subset_pot_andall_1012th('use_cigarettes_30', drinkinglist, YN_order2, 'Response' )
cigarettes_30_risk10_12 =create_high_low_10_non_bullets(cigarettes_30_risk10_12, 'Response')
pot_last_12_risksubset=risk_by_subset_pot_andall_1012th('use_pot_12', drinkinglist, YN_order2, 'Response' )
pot_risk_12_bullets =create_high_low_10_non_bullets(pot_last_12_risksubset, 'Response')
pot_last_30_risksubset=risk_by_subset_pot_andall_1012th('use_pot_30', drinkinglist, YN_order2, 'Response' )
pot_last_30_risksubset =create_high_low_10_non_bullets(pot_last_30_risksubset, 'Response')
OTC_last_12_risksubset=risk_by_subset_pot_andall_1012th('use_cold_medicine_12', drinkinglist, YN_order2, 'Response' )
OTC_risk_12_bullets =create_high_low_10_non_bullets(OTC_last_12_risksubset, 'Response')
VAPE_last_12_risksubset=risk_by_subset_pot_andall_1012th('use_vapes_12', drinkinglist, YN_order2, 'Response' )
VAPE_risk_12_bullets =create_high_low_10_non_bullets(VAPE_last_12_risksubset, 'Response')


# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[35]:


drank30and12 = len(clean[ (clean['use_alcohol_12']=='Yes') &  (clean['use_alcohol_30']=='Yes')])


drank12 =len(clean[clean['use_alcohol_12']=='Yes'])

drank12_12 =len(clean[(clean['use_alcohol_12']=='Yes') & (clean['grade_level']=='12th')])


drank30and12_perc = round((drank30and12/ drank12)*100,1)

drank30and12_12 = len(clean[ (clean['use_alcohol_12']=='Yes') &  (clean['use_alcohol_30']=='Yes')& (clean['grade_level']=='12th')])
drank30and12_perc_12 =round((drank30and12_12/ drank12_12)*100,1)



use_alcohol_30df = YNQ_answers[17]
past30_yes_drankcount =use_alcohol_30df[use_alcohol_30df['Response']=='Yes']['Count (N)'].item()

use_alcohol_30_by_gradedf =YNQ_answers_grade[17]
past30_yes_drankcount101112 = use_alcohol_30_by_gradedf[use_alcohol_30_by_gradedf['Response']=='Yes'][['10th (N)', '11th (N)', '12th (N)']].sum(axis=1, numeric_only=True)
past30_yes_drankcount101112 = round(past30_yes_drankcount101112.item())


use_alcohol_30_by_riskdf =YNQ_answers_by_risk[17]
highrisk_yes_last_drink_30 = use_alcohol_30_by_riskdf[use_alcohol_30_by_riskdf['Response']=='Yes']['High-risk (N)'].item()

percdrank30days_didntdrinkathome = len(clean[(clean['use_alcohol_30']=='Yes') & (clean['where_first_use_alcohol']!='At home')]) /past30_yes_drankcount
percdrank30days_didntdrinkathome =round(percdrank30days_didntdrinkathome*100, 1)


# Match numerator and denominator populations: count higher-risk drinkers
# among 10th-12th graders only (drinker_risk covers all grades, but the
# denominator is the 10th-12th past-30-day drinker count).
highrisk_yes_last_drink_30_1012 = len(clean[(clean['drinker_risk_10_12'] == 'High-risk') & (clean['use_alcohol_30'] == 'Yes')])
percent_yes_last30_drank = round((highrisk_yes_last_drink_30_1012 /past30_yes_drankcount101112)*100,1)

if percdrank30days_didntdrinkathome>= 55:
    percdrank30days_didntdrinkathome_string = f'the majority ({percdrank30days_didntdrinkathome}%)'
    
if (percdrank30days_didntdrinkathome< 55) &  (percdrank30days_didntdrinkathome>= 45):
    percdrank30days_didntdrinkathome_string = f'about half of ({percdrank30days_didntdrinkathome}%)'
    
if (percdrank30days_didntdrinkathome< 45) :
    percdrank30days_didntdrinkathome_string = f'a minority ({percdrank30days_didntdrinkathome}%)'
    
######################################################################
########## report variable, for all vice need per ever used, 12 and 30


total8th_last12VAPE_perc =percent_of_binary_where_total_elsewhere( YNQ_answers_grade[2][YNQ_answers_grade[2]['Response']=='Yes'], age_at_first_grade, '8th')
tota10th_last12VAPE_perc =percent_of_binary_where_total_elsewhere( YNQ_answers_grade[2][YNQ_answers_grade[2]['Response']=='Yes'], age_at_first_grade, '10th')
total12th_last12VAPE_perc =percent_of_binary_where_total_elsewhere( YNQ_answers_grade[2][YNQ_answers_grade[2]['Response']=='Yes'], age_at_first_grade, '12th')




if '8th' in clean['grade_level'].values:
    newrow =  ['Total', _safe_item(total8th_last12VAPE_perc), _safe_item(tota10th_last12VAPE_perc), _safe_item(total12th_last12VAPE_perc)]
else:
    newrow =  ['Total', total8th_last12VAPE_perc, _safe_item(tota10th_last12VAPE_perc), _safe_item(total12th_last12VAPE_perc)]

VAPEPAST12 =by_grader_count_and_sex_percenter_yn('use_vapes_12',
                                                 ['sex','8th', '10th', '12th'],
                                                 newrow)






    


total8th_last30VAPE_perc =percent_of_binary_where_total_elsewhere( YNQ_answers_grade[19][YNQ_answers_grade[19]['Response']=='Yes'], age_at_first_grade, '8th')
tota10th_last30VAPE_perc =percent_of_binary_where_total_elsewhere( YNQ_answers_grade[19][YNQ_answers_grade[19]['Response']=='Yes'], age_at_first_grade, '10th')
total12th_last30VAPE_perc =percent_of_binary_where_total_elsewhere( YNQ_answers_grade[19][YNQ_answers_grade[19]['Response']=='Yes'], age_at_first_grade, '12th')

if '8th' in clean['grade_level'].values:
    newrow =  ['Total', _to_val(total8th_last30VAPE_perc), _to_val(tota10th_last30VAPE_perc), _to_val(total12th_last30VAPE_perc)]
else:
    newrow =  ['Total', _to_val(total8th_last30VAPE_perc), _to_val(tota10th_last30VAPE_perc), _to_val(total12th_last30VAPE_perc)]

VAPEPAST30 =by_grader_count_and_sex_percenter_yn('use_vapes_30', ['sex','8th', '10th', '12th'], newrow)


total8th_last12POT_perc =percent_of_binary_where_total_elsewhere( YNQ_answers_grade[4][YNQ_answers_grade[4]['Response']=='Yes'], age_at_first_grade, '8th')
tota10th_last12POT_perc =percent_of_binary_where_total_elsewhere( YNQ_answers_grade[4][YNQ_answers_grade[4]['Response']=='Yes'], age_at_first_grade, '10th')
total12th_last12POT_perc =percent_of_binary_where_total_elsewhere( YNQ_answers_grade[4][YNQ_answers_grade[4]['Response']=='Yes'], age_at_first_grade, '12th')


if '8th' in clean['grade_level'].values:
    newrow =  ['Total', _to_val(total8th_last12POT_perc), _to_val(tota10th_last12POT_perc), _to_val(total12th_last12POT_perc)]
else:
    newrow =  ['Total', total8th_last12POT_perc, _safe_item(tota10th_last12POT_perc), _safe_item(total12th_last12POT_perc)]

POTPAST12 =by_grader_count_and_sex_percenter_yn('use_pot_12', ['sex','8th', '10th', '12th'], newrow)


total8th_last30POT_perc =percent_of_binary_where_total_elsewhere( YNQ_answers_grade[21][YNQ_answers_grade[21]['Response']=='Yes'], age_at_first_grade, '8th')
tota10th_last30POT_perc =percent_of_binary_where_total_elsewhere( YNQ_answers_grade[21][YNQ_answers_grade[21]['Response']=='Yes'], age_at_first_grade, '10th')
total30th_last30POT_perc =percent_of_binary_where_total_elsewhere( YNQ_answers_grade[21][YNQ_answers_grade[21]['Response']=='Yes'], age_at_first_grade, '12th')



if '8th' in clean['grade_level'].values:
    newrow =  ['Total', _safe_item(total8th_last30POT_perc), _safe_item(tota10th_last30POT_perc), _safe_item(total30th_last30POT_perc)]
else:
    newrow =  ['Total', total8th_last30POT_perc, _safe_item(tota10th_last30POT_perc), _safe_item(total30th_last30POT_perc)]

POTPAST30 =by_grader_count_and_sex_percenter_yn('use_pot_30', ['sex','8th', '10th', '12th'], newrow)

days_5_or_more_by_risk1012 = by_grader_count_and_percenter('days_5_or_more_drinks','drinker_risk_10_12', drinkinglist, sort_order4, 'Response')




howmanyhighrisk5_count = round(days_5_or_more_by_risk1012[~days_5_or_more_by_risk1012['Response'].str.contains('I did not')]['High-risk (N)'].sum())
howmanyhighrisk5_perc = round((howmanyhighrisk5_count / past30_yes_drankcount101112)*100, 1)


# In[ ]:





# In[ ]:





# In[36]:


YNQ_answers_grade_Q_perc_all = []
YNQ_answers_grade_Q_perc_allqs = []
timelinesssss = []
gradesssss =[]
substancesss = []
for q, Q in zip(YNQs, YNQ_answers_grade):
    
    if '8th' in clean['grade_level'].values:

        for grade in ['8th', '10th', '12th']:

            percent =Q[Q['Response']=='Yes'][f"{grade} (%)"].str.split('%').str[0] 

            YNQ_answers_grade_Q_perc_allqs.append(q)

            if (re.findall(r'\d+', q)) ==['12']:
                timelinesssss.append('past-12-months')
            elif (re.findall(r'\d+', q)) ==['30']:
                timelinesssss.append('past-30-days')

            YNQ_answers_grade_Q_perc_all.append(percent.item())
            gradesssss.append(grade)
            substancesss.append(q.split("_")[1].capitalize())
            
    else:
        for grade in ['10th', '12th']:

            percent =Q[Q['Response']=='Yes'][f"{grade} (%)"].str.split('%').str[0] 

            YNQ_answers_grade_Q_perc_allqs.append(q)

            if (re.findall(r'\d+', q)) ==['12']:
                timelinesssss.append('past-12-months')
            elif (re.findall(r'\d+', q)) ==['30']:
                timelinesssss.append('past-30-days')

            YNQ_answers_grade_Q_perc_all.append(percent.item())
            gradesssss.append(grade)
            substancesss.append(q.split("_")[1].capitalize())


# In[37]:


all_the_sub_dfs_again = []

for Q in YNQs:
    print(Q)
    df =risk_by_subset_pot_andall_1012th(Q, drinkinglist, YN_order2, 'Response' )
    df = df[df['Response']=='Yes']
    df['Q'] = Q
    df['Substance'] = Q.split("_")[1].capitalize()
    all_the_sub_dfs_again.append(df)
    


all_sub_risks_all = pd.concat(all_the_sub_dfs_again)

all_sub_risks_all = fix_MTF_naming(all_sub_risks_all)

all_sub_risks_all= all_sub_risks_all[all_sub_risks_all['Substance']!='Alcohol']

all_sub_risks_all['timeline'] = all_sub_risks_all['Q'].str[-2:]
all_sub_risks_all_output =all_sub_risks_all[all_sub_risks_all['timeline']=='12'][['Substance','High-risk (%)', '10-12th Graders','Low-risk (%)', 'Non-drinker (%)']]
all_sub_risks_all_output30 =all_sub_risks_all[all_sub_risks_all['timeline']=='30'][['Substance','High-risk (%)', '10-12th Graders','Low-risk (%)', 'Non-drinker (%)']]

highriskampusage =all_sub_risks_all_output[all_sub_risks_all_output['Substance']=='Amphetamines']['High-risk (%)'].item()

substances_DF =pd.DataFrame({'percent':YNQ_answers_grade_Q_perc_all,
            'Question': YNQ_answers_grade_Q_perc_allqs,
            'timeline': timelinesssss,
             'grade_level':gradesssss,
             'substance':substancesss} )

substances_DF=fix_MTF_naming(substances_DF)


substances_DF= substances_DF[substances_DF['substance']!='Alcohol']
#substances_DF = substances_DF.pivot(columns ='grade_level', index=['substance', 'timeline'], values='percent').reset_index()


ALLSUBSTANCE8th_12 = substances_DF[(substances_DF['grade_level']=='8th') &(substances_DF['timeline']=='past-12-months') ][['substance', 'percent']].sort_values('percent', ascending=False)
ALLSUBSTANCE10th_12 = substances_DF[(substances_DF['grade_level']=='10th') &(substances_DF['timeline']=='past-12-months') ][['substance', 'percent']].sort_values('percent', ascending=False)
ALLSUBSTANCE12th_12 = substances_DF[(substances_DF['grade_level']=='12th') &(substances_DF['timeline']=='past-12-months') ][['substance', 'percent']].sort_values('percent', ascending=False)


ALLSUBSTANCE8th_30 = substances_DF[(substances_DF['grade_level']=='8th') &(substances_DF['timeline']=='past-30-days') ][['substance', 'percent']].sort_values('percent', ascending=False)
ALLSUBSTANCE10th_30 = substances_DF[(substances_DF['grade_level']=='10th') &(substances_DF['timeline']=='past-30-days') ][['substance', 'percent']].sort_values('percent', ascending=False)
ALLSUBSTANCE12th_30 = substances_DF[(substances_DF['grade_level']=='12th') &(substances_DF['timeline']=='past-30-days') ][['substance', 'percent']].sort_values('percent', ascending=False)

substancedfsss= []

for Q in YNQs:
    print(Q)
    for grade in ['8th', '10th', '12th']:
        noyes = clean[[Q,'grade_level', 'sex', 'Token']].groupby([Q,'grade_level', 'sex']).count().reset_index()
        noyes =noyes[noyes[Q]=='Yes']
        noyes =noyes[['grade_level', 'sex', 'Token']]
        noyes.columns = ['grade_level', 'sex', 'countyes']
        totals= clean[[Q, 'grade_level', 'sex', 'Token']].groupby(['grade_level', 'sex']).count().reset_index()
        totals =totals[['grade_level', 'sex', 'Token']]
        totals.columns = ['grade_level', 'sex', 'totalbysex']
        perctable = pd.merge(noyes.reset_index(), totals.reset_index(), on=['grade_level', 'sex'], how='right')
        perctable['perc'] = round((perctable['countyes'] / perctable['totalbysex'])*100, 1)
        #perctable =perctable.pivot(index='sex', columns='grade_level', values='perc').reset_index()
        perctable = perctable[perctable['grade_level']==grade]
        perctable['substance'] = Q.split("_")[1].capitalize()
        if (re.findall(r'\d+', Q)) ==['12']:
            timeliness= 'past-12-months'
        elif (re.findall(r'\d+', Q)) ==['30']:
            timeliness ='past-30-days'
        perctable['timeliness'] = timeliness
        substancedfsss.append(perctable[['substance', 'timeliness','grade_level', 'sex', 'perc']])
     
    
substancesbysexandgrader =pd.concat(substancedfsss)
substancesbysexandgrader = fix_MTF_naming(substancesbysexandgrader)
substancesbysexandgrader= substancesbysexandgrader[(substancesbysexandgrader['substance']!='Alcohol')]
sub12months_sex8th = substancesbysexandgrader[(substancesbysexandgrader['grade_level']=='8th') & (substancesbysexandgrader['timeliness']=='past-12-months')].pivot(index=['substance'], columns=['grade_level', 'sex'], values='perc').reset_index().fillna(0)
sub12months_sex10th = substancesbysexandgrader[(substancesbysexandgrader['grade_level']=='10th') & (substancesbysexandgrader['timeliness']=='past-12-months')].pivot(index=['substance'], columns=['grade_level', 'sex'], values='perc').reset_index().fillna(0)
sub12months_sex12th = substancesbysexandgrader[(substancesbysexandgrader['grade_level']=='12th') & (substancesbysexandgrader['timeliness']=='past-12-months')].pivot(index=['substance'], columns=['grade_level', 'sex'], values='perc').reset_index().fillna(0)
# In[ ]:
def rename_sex_cols(df, grade):
    """Rename substance-by-sex pivot columns, handling 2 or 3 sex categories."""
    n = len(df.columns) - 1  # subtract 'substance' col
    if n == 3:
        df.columns = ['substance', f'{grade} Female', f'{grade} Male', f'{grade} Prefer Not to Say']
    elif n == 2:
        df.columns = ['substance', f'{grade} Female', f'{grade} Male']
    else:
        df.columns = ['substance'] + [f'{grade} Sex{i}' for i in range(n)]
    return df

if '8th' in clean['grade_level'].values:
    sub12months_sex8th = rename_sex_cols(sub12months_sex8th, '8th')
sub12months_sex10th = rename_sex_cols(sub12months_sex10th, '10th')
sub12months_sex12th = rename_sex_cols(sub12months_sex12th, '12th')




# In[38]:


neverQ2_answers = multi_questioner(moreNver123Q, sort_order18, 'Response' )
neverQ2_answers_grade= multi_by_grader( moreNver123Q, 'grade_level', grades, sort_order18, 'Response' )
neverQ2_answers_by_risk= multi_by_grader( moreNver123Q,'drinker_risk', drinkinglist, sort_order18, 'Response' )


# In[39]:


tookcare = neverQ2_answers_by_risk[1]
vom = neverQ2_answers_by_risk[9]
dranktoomuch = risk_by_subset_pot_andall_1012th('taking_care_drank_too_much_other', drinkinglist, sort_order18, 'Response')
vom2 = risk_by_subset_pot_andall_1012th('witness_vomit_other', drinkinglist, sort_order18, 'Response')
# CORRECTED: Cap percentage sums at 100.0 to prevent floating-point overshoot (e.g., 100.1%)
# B4: from counts. Executive Summary bullets use the all-grade groups (they sit
# under the all-grade higher-risk count); the "Consequences to Others" section
# uses the body groups (BODY_GROUP_GRADES), so it gets its own variables (B3).
tookcare_perc = group_share('taking_care_drank_too_much_other', _not_never, group_col='drinker_risk', name='exec: higher-risk took care of someone')
vom_perc  = group_share('witness_vomit_other', _not_never, group_col='drinker_risk', name='exec: higher-risk saw someone vomit')
tookcare_perc_1012 = group_share('taking_care_drank_too_much_other', _not_never, name='higher-risk 10-12 took care of someone')
vom_perc_1012 = group_share('witness_vomit_other', _not_never, name='higher-risk 10-12 saw someone vomit')


def body_grade_share(Q, positive, name=None):
    sub = clean[clean['grade_level'].isin(BODY_GROUP_GRADES) & clean[Q].notna()]
    return pct(int(positive(sub[Q]).sum()), len(sub), name, f'{BODY_GROUP_LABEL} graders who answered {Q}')
dranktoomuch_perc1012 = body_grade_share('taking_care_drank_too_much_other', _not_never)
vom_10_12_perc = body_grade_share('witness_vomit_other', _not_never)
concerned_other  = risk_by_subset_pot_andall_1012th('taking_concerned_other', drinkinglist, sort_order18, 'Response')
concerned_other_perc = group_share('taking_concerned_other', _not_never, name='higher-risk 10-12 talked to someone')
concerned_other_perc10_12 = body_grade_share('taking_concerned_other', _not_never)


# Have you ever used alcohol, marijuana, or other drugs not prescribed to you  
# (do NOT include tobacco/nicotine) before coming to, or during, a school dance, school 
# sporting event, or another school-sponsored event? 

## this one was binary ?? 
schooldance= binary_Qer('school_dance_use')

school_dance_use_grade = by_grader_count_and_percenter('school_dance_use','grade_level', grades, YN_order2, 'Response')
binary_grade_fixer(school_dance_use_grade)

school_dance_use_by_risk = by_grader_count_and_percenter('school_dance_use','drinker_risk', drinkinglist, YN_order2, 'Response')
binary_grade_fixer(school_dance_use_by_risk)

## 94.6% of {{School}} students have never used alcohol, marijuana, or other drugs not prescribed to them before coming to, or during, a school dance, school sporting event, or other school-sponsored event. 
danceperc = schooldance[schooldance['Response']=='No']['Percentage (%)']
danceperc = danceperc.item()

## Which of the following statements best represents YOUR OWN attitude towards teenagers drinking alcohol? 
attitude = create_count_and_percent_0check(clean, 'attitude', 'Response', 'No', 'No', sort_order6)
attitude = sort_by_customlist(attitude, 'Response', sort_order6)

attitude_grade = by_grader_count_and_percenter('attitude','grade_level', grades, sort_order6, 'Response')
attitude_by_risk = by_grader_count_and_percenter('attitude','drinker_risk', drinkinglist, sort_order6, 'Response')

# CORRECTED: Round all percentage sums to 1 decimal place and cap at 100%
neveragoodthingtodo_perc = cap_percentage(attitude_by_risk[:2]['Non-drinker (%)'].str.split('%').str[0].astype(float).sum())
neveragoodthingtodo_perc_high = cap_percentage(attitude_by_risk[:2]['High-risk (%)'].str.split('%').str[0].astype(float).sum())
nondrink_okay_attitude = cap_percentage(attitude_by_risk.iloc[3:]['Non-drinker (%)'].str.split('%').str[0].astype(float).sum())
highrisk_okay_attitude = cap_percentage(attitude_by_risk.iloc[3:]['High-risk (%)'].str.split('%').str[0].astype(float).sum())


# In[ ]:





# In[40]:


ApproveDisapprove_answers = multi_questioner(ApproveDisapproveQ, sort_order7, 'Response' )
ApproveDisapprove_answers_grade= multi_by_grader( ApproveDisapproveQ, 'grade_level', grades, sort_order7, 'Response' )
ApproveDisapprove_answers_by_risk= multi_by_grader( ApproveDisapproveQ,'drinker_risk', drinkinglist, sort_order7, 'Response' )

##How much do you think people risk harming themselves (physically or in other ways)
#if they try one or two drinks of an alcoholic beverage (beer, wine, liquor)? 
RiskQ_answers = multi_questioner(RiskQs, sort_order8, 'Response' )
RiskQ_answers_grade= multi_by_grader( RiskQs, 'grade_level', grades, sort_order8, 'Response' )
RiskQ_answers_by_risk= multi_by_grader( RiskQs,'drinker_risk', drinkinglist, sort_order8, 'Response' )

risk_harm_1_2= total_aharm_no_harm_nograde_PERC(RiskQ_answers[1])
risk_harm_daily_1_2perc = risk_harm_1_2[risk_harm_1_2['index']=='Total mod/great']['Percentage (%)'].item()
risk_harm_5_weekend= total_aharm_no_harm_nograde_PERC(RiskQ_answers[3])
risk_harm_5_weekendperc = risk_harm_5_weekend[risk_harm_5_weekend['index']=='Total mod/great']['Percentage (%)'].item()


# In[ ]:





# In[41]:


ApproveDisapprove2_answers = multi_questioner(ApproveDisapprove2_q, sort_order9, 'Response' )
ApproveDisapprove2_answers_grade= multi_by_grader( ApproveDisapprove2_q, 'grade_level', grades, sort_order9, 'Response' )
ApproveDisapprove2_answers_by_risk= multi_by_grader( ApproveDisapprove2_q,'drinker_risk', drinkinglist, sort_order9, 'Response' )

AgreeDisagree_answers = multi_questioner(AgreeDisagree_q, sort_order10, 'Response' )
AgreeDisagree_answers_grade= multi_by_grader( AgreeDisagree_q, 'grade_level', grades, sort_order10, 'Response' )
AgreeDisagree_answers_by_risk= multi_by_grader( AgreeDisagree_q,'drinker_risk', drinkinglist, sort_order10, 'Response' )


social_atmo_df= AgreeDisagree_answers[0]
pressured= AgreeDisagree_answers[1]


social_atmo_df_totals = total_agree_disagree_nograde(social_atmo_df)
social_atmo_df_yes_perc = round(social_atmo_df_totals[social_atmo_df_totals['index']=='Total Agree']['Count (N)'].item(),1)

social_atmo_7th_yes = total_agree_disagree_perc(AgreeDisagree_answers_grade[0], 'yes')

if '7th' in clean['grade_level'].values:
    social_atmo_7th_yes_perc = social_atmo_7th_yes[social_atmo_7th_yes['index']=='Total Agree']['7th (%)'].item()
elif '8th' in clean['grade_level'].values and '8th (%)' in social_atmo_7th_yes.columns:
    # Use 8th grade as lowest when no 7th graders
    social_atmo_7th_yes_perc = social_atmo_7th_yes[social_atmo_7th_yes['index']=='Total Agree']['8th (%)'].item()
else:
    social_atmo_7th_yes_perc =''

social_atmo_12th_yes_perc = social_atmo_7th_yes[social_atmo_7th_yes['index']=='Total Agree']['12th (%)'].item()

social_atmo_bygrade = AgreeDisagree_answers_grade[0]
if '7th (N)' in social_atmo_bygrade.columns and '8th (N)' in social_atmo_bygrade.columns:
    seventhand8thsocial_atmocountagree = social_atmo_bygrade[(social_atmo_bygrade['Response']=='Agree') | (social_atmo_bygrade['Response']=='Strongly agree')][['7th (N)', '8th (N)']].sum().sum()
else:
    seventhand8thsocial_atmocountagree = np.nan
    
ninto12thsocial_atmocountagree = social_atmo_bygrade[(social_atmo_bygrade['Response']=='Agree') | (social_atmo_bygrade['Response']=='Strongly agree')][['9th (N)', '10th (N)', '11th (N)', '12th (N)']].sum().sum()
_denom_sa = totalsbygrade[totalsbygrade['grade_level'].isin(['7th', '8th'])]['totalingrade'].sum()
seventhand8thsocial_atmocountagree_perc = round((seventhand8thsocial_atmocountagree / _denom_sa)*100, 1) if _denom_sa > 0 else np.nan
# Denominator must cover the same 9th-12th population as the numerator
# (positional [2:] picked the wrong rows once totalsbygrade was padded).
ninto12thsocial_atmocountagree_perc = round((ninto12thsocial_atmocountagree /totalsbygrade[totalsbygrade['grade_level'].isin(['9th', '10th', '11th', '12th'])]['totalingrade'].sum())*100, 1)


socialatmo =total_agree_disagree(AgreeDisagree_answers_grade[0])


disagree_social_atmo = round(100-social_atmo_df_yes_perc, 1)


# In[ ]:





# In[ ]:





# In[42]:



bg= AgreeDisagree_answers[8]

bingedrinking = total_agree_disagree_nograde_PERC(bg)
bingedrinkperc = bingedrinking[bingedrinking['index']=='Total Disagree']['Percentage (%)'].item()

# B5: wording from the number (the old thresholds overlapped and called
# anything under 65% "About half").
bingedrinkperc = round(bingedrinkperc, 1)
bingedrinkstring = f'{share_word(bingedrinkperc).capitalize()} ({bingedrinkperc}%)'
    
if '7th' in clean['grade_level'].values:
    totalagree_bingedrinking_7th = AgreeDisagree_answers_grade[8][2:]['7th (%)'].str.split('%').str[0].astype(float).sum()
elif '8th' in clean['grade_level'].values and '8th (%)' in AgreeDisagree_answers_grade[8].columns:
    # Use 8th grade as lowest when no 7th graders
    totalagree_bingedrinking_7th = round(AgreeDisagree_answers_grade[8][2:]['8th (%)'].str.split('%').str[0].astype(float).sum(), 1)
else:
    totalagree_bingedrinking_7th = ''
    
    
totalagree_bingedrinking_12th = AgreeDisagree_answers_grade[8][2:]['12th (%)'].str.split('%').str[0].astype(float).sum()

pressured_df_totals = total_agree_disagree_nograde(pressured)
pressured_df_yes_perc = pressured_df_totals[pressured_df_totals['index']=='Total Agree']['Count (N)'].item()
pressured_df_yes_perc= round(pressured_df_yes_perc,1)
pressured_7th_yes = total_agree_disagree_perc(AgreeDisagree_answers_grade[1], 'yes')

if '7th' in clean['grade_level'].values:
    pressured_7th_yes_perc = pressured_7th_yes[pressured_7th_yes['index']=='Total Agree']['7th (%)'].item()
elif '8th' in clean['grade_level'].values and '8th (%)' in pressured_7th_yes.columns:
    # Use 8th grade as lowest when no 7th graders
    pressured_7th_yes_perc = pressured_7th_yes[pressured_7th_yes['index']=='Total Agree']['8th (%)'].item()
else:
    pressured_7th_yes_perc=''
    
pressured_12th_yes_perc = pressured_7th_yes[pressured_7th_yes['index']=='Total Agree']['12th (%)'].item()

pressured_table =total_agree_disagree(AgreeDisagree_answers_grade[1])


pressured_bygrade = AgreeDisagree_answers_grade[1]
if '7th (N)' in pressured_bygrade.columns and '8th (N)' in pressured_bygrade.columns:
    seventhand8thpressured_bygradecountagree = pressured_bygrade[(pressured_bygrade['Response']=='Agree') | (pressured_bygrade['Response']=='Strongly agree')][['7th (N)', '8th (N)']].sum().sum()
else:
    seventhand8thpressured_bygradecountagree = np.nan
    
ninto12thpressured_bygradecountagree = pressured_bygrade[(pressured_bygrade['Response']=='Agree') | (pressured_bygrade['Response']=='Strongly agree')][['9th (N)', '10th (N)', '11th (N)', '12th (N)']].sum().sum()
_denom_pr = totalsbygrade[totalsbygrade['grade_level'].isin(['7th', '8th'])]['totalingrade'].sum()
seventhand8thpressured_bygradecountagree_perc = round((seventhand8thpressured_bygradecountagree / _denom_pr)*100, 1) if _denom_pr > 0 else np.nan
# Same 9th-12th denominator fix as the social-atmosphere stat above.
ninto12thpressured_bygradecountagree_perc = round((ninto12thpressured_bygradecountagree /totalsbygrade[totalsbygrade['grade_level'].isin(['9th', '10th', '11th', '12th'])]['totalingrade'].sum())*100, 1)


agree_easy_perc = total_agree_disagree_nograde(AgreeDisagree_answers[2])
agree_easy_perc_val =agree_easy_perc[agree_easy_perc['index']=='Total Agree']['Count (N)'].item()
agree_easy_perc_val = round(agree_easy_perc_val,1)


own_decision_table =total_agree_disagree( AgreeDisagree_answers_grade[4])
respect_table =total_agree_disagree( AgreeDisagree_answers_grade[5])
respect_table_nograde = total_agree_disagree_nograde_PERC( AgreeDisagree_answers[5])
respect_yes_perc = respect_table_nograde[respect_table_nograde['index']=='Total Agree']['Percentage (%)'].item()
fitin_table_nograde = total_agree_disagree_nograde_PERC( AgreeDisagree_answers[6])
fitin_no_perc = fitin_table_nograde[fitin_table_nograde['index']=='Total Disagree']['Percentage (%)'].item()


fitinrespectstring = share_word((respect_yes_perc + fitin_no_perc) / 2).capitalize()   # B5


# In[ ]:





# In[43]:


NeverOnce_answers = multi_questioner(Never_once_qs, sort_order11, 'Response' )
NeverOnce_answers_grade= multi_by_grader( Never_once_qs, 'grade_level', grades, sort_order11, 'Response' )
NeverOnce_answers_by_risk= multi_by_grader( Never_once_qs,'drinker_risk', drinkinglist, sort_order11, 'Response' )

freq_answers = multi_questioner(freq_qs, sort_order5, 'Frequency' )
freq_answers_grade= multi_by_grader( freq_qs, 'grade_level', grades, sort_order5, 'Frequency' )
freq_answers_by_risk= multi_by_grader( freq_qs,'drinker_risk', drinkinglist, sort_order5, 'Frequency' )

drink_how_often_your_grade_grade = freq_answers_grade[0]

if '8th' in clean['grade_level'].values:
    eighth_perceived_neverdrink_perc  = drink_how_often_your_grade_grade[drink_how_often_your_grade_grade['Frequency']=='Never']['8th (%)'].item()
else:
    eighth_perceived_neverdrink_perc = ''
    
ninth_perceived_neverdrink_perc  = drink_how_often_your_grade_grade[drink_how_often_your_grade_grade['Frequency']=='Never']['9th (%)'].item()
tenth_perceived_neverdrink_perc  = drink_how_often_your_grade_grade[drink_how_often_your_grade_grade['Frequency']=='Never']['10th (%)'].item()
eleven_perceived_neverdrink_perc  = drink_how_often_your_grade_grade[drink_how_often_your_grade_grade['Frequency']=='Never']['11th (%)'].item()
twelve_perceived_neverdrink_perc  = drink_how_often_your_grade_grade[drink_how_often_your_grade_grade['Frequency']=='Never']['12th (%)'].item()


# B2/B6: perceived figures cover the same answer range as the actual figures
# ("once a month or more"), keyed by answer label, from counts.
_dhy = drink_how_often_your_grade_grade
tenth_perceived_oncemonth_perc = fmt_pct(count_pct_from_table(_dhy, _dhy['Frequency'].isin(['Once a month', 'Twice a month', 'Once a week', 'Twice a week', 'Daily']), '10th (%)'))
eleventhmorethantwice_perceived = count_pct_from_table(_dhy, _dhy['Frequency'].isin(['Twice a month', 'Once a week', 'Twice a week', 'Daily']), '11th (%)')
twelvethtwiceaweek_perceived = count_pct_from_table(_dhy, _dhy['Frequency'].eq('Twice a week'), '12th (%)')


#How many whole drinks of alcohol, if any, do you think students in your grade usually have on average when they drink?  

howmanygrade = create_count_and_percent_0check(clean, 'how_many_your_grade', 'Amount', 'Yes', 'No', drinkslistnodash)
how_many_your_grade_grade = by_grader_count_and_percenter('how_many_your_grade','grade_level', grades, drinkslistnodash, 'Amount')
how_many_your_grade_by_risk = by_grader_count_and_percenter('how_many_your_grade','drinker_risk', drinkinglist, drinkslistnodash, 'Amount')



# In[44]:


#How many whole drinks of alcohol, if any, do you think 12th graders usually have on average when they drink?  

howmany12 = create_count_and_percent_0check(clean, 'how_many_seniors', 'Amount', 'Yes', 'No', drinkslistnodash)
how_many_seniors_grade = by_grader_count_and_percenter('how_many_seniors','grade_level', grades, drinkslistnodash, 'Amount')
how_many_seniors_by_risk = by_grader_count_and_percenter('how_many_seniors','drinker_risk', drinkinglist, drinkslistnodash, 'Amount')

one_drink_days_answers = multi_questioner(one_drink_days_qs, sort_order12, 'Response' )
one_drink_days_answers_grade= multi_by_grader( one_drink_days_qs, 'grade_level', grades, sort_order12, 'Response' )
one_drink_days_answers_by_risk= multi_by_grader( one_drink_days_qs,'drinker_risk', drinkinglist, sort_order12, 'Response' )

#What is the greatest number of whole drinks of alcohol you think students in your grade who drink consumed in a row in the PAST 30 DAYS? 
howmany30grade = create_count_and_percent_0check(clean, 'how_many_30_days_your_grade', 'Amount', 'Yes', 'No', drinkslistnodash)

# What is the greatest number of whole drinks of alcohol you think 12th graders who drink consumed in a row in the PAST 30 DAYS?
howmany30grade12 = create_count_and_percent_0check(clean, 'how_many_30_days_your_seniors', 'Amount', 'Yes', 'No', drinkslistnodash)
how_many_30_days_answers_grade= multi_by_grader( drinks0Qs,'grade_level', grades, drinkslistnodash, 'Amount' )
how_many_30_days_answers_by_risk= multi_by_grader( drinks0Qs,'drinker_risk', drinkinglist, drinkslistnodash, 'Amount' )



pot_freq_answers = multi_questioner(pot_freq_q, sort_order5, 'Frequency' )
pot_freq_q_answers_grade= multi_by_grader( pot_freq_q,'grade_level', grades, sort_order5, 'Frequency' )
pot_freq_q_answers_by_risk= multi_by_grader( pot_freq_q,'drinker_risk', drinkinglist, sort_order5, 'Frequency' )


# In[ ]:





# In[45]:


pot_how_often_your_grade_grade = pot_freq_q_answers_grade[0]
if '8th' in clean['grade_level'].values:
    eighth_perceived_neverpot_perc  = perceived_pot('8th')
else: 
    eighth_perceived_neverpot_perc = ''
    
ninth_perceived_neverpot_perc = perceived_pot('9th')
tenth_perceived_neverpot_perc = perceived_pot('10th')
eleventh_perceived_neverpot_perc = perceived_pot('11th')
twelth_perceived_neverpot_perc = perceived_pot('12th')

# CORRECTED: Cap percentage sums at 100%
twelthpot_morethanonceamonth_perceived = cap_percentage((pot_how_often_your_grade_grade[2:]['12th (%)'].str.split('%').str[0]).astype(float).sum())
twelthpot_more12year_perceived = round((pot_how_often_your_grade_grade[pot_how_often_your_grade_grade['Frequency']=="1 to 2 times per year"]['12th (%)'].str.split('%').str[0]).astype(float).sum(), 1)
#What percentage of students in your school do you think have ever used alcohol, marijuana, or other drugs not prescribed for them (do NOT include tobacco/nicotine) before coming to, or during, a school dance, school sporting event, or other school-sponsored events? 

percent_ever_used = create_count_and_percent_0check(clean, 'percent_ever_used', 'Amount', 'Yes', 'No', perc_order)
percent_ever_used_grade = by_grader_count_and_percenter('percent_ever_used','grade_level', grades, perc_order, 'Amount')
percent_ever_used_by_risk = by_grader_count_and_percenter('percent_ever_used','drinker_risk', drinkinglist, perc_order, 'Amount')


# In[46]:


too_lenient_answers = multi_questioner(too_lenient_qs, sort_order13, 'Response' )
too_lenient_answers_grade= multi_by_grader( too_lenient_qs,'grade_level', grades, sort_order13, 'Response' )
too_lenient_answers_by_risk= multi_by_grader( too_lenient_qs,'drinker_risk', drinkinglist, sort_order13, 'Response' )


subsetpolicies_1012 =risk_by_subset_pot_andall_1012th('thoughts_alcohol_policies', drinkinglist, sort_order13, 'Response')

allstudents_policies = too_lenient_answers[0]
allstudents_policies = allstudents_policies[['Response', 'Percentage (%)']]
allstudents_policies.columns = ['Response', 'All Students']

allstudents_policies =pd.merge(subsetpolicies_1012, allstudents_policies, on='Response')
# B6: POLICY values get their own names. The old code reused the same four
# names for Enforcement below, so both sets of bullets printed Enforcement.
pol_highrisktoostrict = allstudents_policies[allstudents_policies['Response']=='Too strict']['High-risk (%)'].item()
pol_norisktoostrict = allstudents_policies[allstudents_policies['Response']=='Too strict']['Non-drinker (%)'].item()
pol_highrisktoolenient = allstudents_policies[allstudents_policies['Response']=='Too lenient']['High-risk (%)'].item()
pol_norisktoolenient = allstudents_policies[allstudents_policies['Response']=='Too lenient']['Non-drinker (%)'].item()

subsetenforcement_1012 =risk_by_subset_pot_andall_1012th('thoughts_enforcement', drinkinglist, sort_order13, 'Response')

allstudents_enforcement = too_lenient_answers[1]
allstudents_enforcement = allstudents_enforcement[['Response', 'Percentage (%)']]
allstudents_enforcement.columns = ['Response', 'All Students']

allstudents_enforcement =pd.merge(subsetenforcement_1012, allstudents_enforcement, on='Response')
highrisktoostrict = allstudents_enforcement[allstudents_enforcement['Response']=='Too strict']['High-risk (%)'].item()
norisktoostrict = allstudents_enforcement[allstudents_enforcement['Response']=='Too strict']['Non-drinker (%)'].item()

highrisktoolenient = allstudents_enforcement[allstudents_enforcement['Response']=='Too lenient']['High-risk (%)'].item()
norisktoolenient = allstudents_enforcement[allstudents_enforcement['Response']=='Too lenient']['Non-drinker (%)'].item()


# In[ ]:





# In[47]:


AgreeDisagree2_answers = multi_questioner(AgreeDisagree2_q, sort_order10, 'Response' )
AgreeDisagree2_answers_grade= multi_by_grader( AgreeDisagree2_q,'grade_level', grades, sort_order10, 'Response' )
AgreeDisagree2_answers_by_risk= multi_by_grader( AgreeDisagree2_q,'drinker_risk', drinkinglist, sort_order10, 'Response' )

abstain_table =total_agree_disagree( AgreeDisagree_answers_grade[6])
easytomake_table=total_agree_disagree(AgreeDisagree_answers_grade[2])
assume_table=total_agree_disagree(AgreeDisagree_answers_grade[3])

light_table=total_agree_disagree(AgreeDisagree_answers_grade[7])
heavy_table=total_agree_disagree(AgreeDisagree_answers_grade[8])



academics_to_me_all = total_agree_disagree_nograde_PERC(AgreeDisagree2_answers[1])
acade_import_all =academics_to_me_all[academics_to_me_all['index']=='Total Agree']['Percentage (%)'].item()

valued_perc = total_agree_disagree_nograde(AgreeDisagree2_answers[0])
valued_perc_val =valued_perc[valued_perc['index']=='Total Agree']['Count (N)'].item()
valued_perc_val = round(valued_perc_val, 1)

teachers_perc = total_agree_disagree_nograde(AgreeDisagree2_answers[2])
teachers_perc_val =round(teachers_perc[teachers_perc['index']=='Total Agree']['Count (N)'].item(),1)

sincere_df= total_agree_disagree_nograde_PERC(AgreeDisagree2_answers[3])
sincere_perc_val =round(sincere_df[sincere_df['index']=='Total Agree']['Percentage (%)'].item(),1)

care_table = total_agree_disagree_nograde_PERC(AgreeDisagree2_answers[4])
care_perc_val =care_table[care_table['index']=='Total Agree']['Percentage (%)'].item()

#Which of the following statements best describes how you felt overall about going to school during the PAST 12 MONTHS? 

school12 = create_count_and_percent_0check(clean, 'school_12', 'Response', 'No', 'No', sort_order14)
school12 = sort_by_customlist(school12, 'Response', sort_order14)
school12_grade = by_grader_count_and_percenter('school_12','grade_level', grades, sort_order14, 'Response')
school12_by_risk = by_grader_count_and_percenter('school_12','drinker_risk', drinkinglist, sort_order14, 'Response')

# CORRECTED: Round percentage sums
schoollikedatlotorkindofliked = round(school12[:2]['Percentage (%)'].str.split('%').str[0].astype(float).sum(), 1)

if '7th' in clean['grade_level'].values:
    schoollikedatlotorkindofliked7 = round(school12_grade[:2]['7th (%)'].str.split('%').str[0].astype(float).sum(), 1)
else:
    schoollikedatlotorkindofliked7 = ''

schoollikedatlotorkindofliked12 = round(school12_grade[:2]['12th (%)'].str.split('%').str[0].astype(float).sum(), 1)

AgreeDisagree3_answers = multi_questioner(AgreeDisagree3_q, sort_order10, 'Response' )
AgreeDisagree3_answers_grade= multi_by_grader( AgreeDisagree3_q,'grade_level', grades, sort_order10, 'Response' )
AgreeDisagree3_answers_by_risk= multi_by_grader( AgreeDisagree3_q,'drinker_risk', drinkinglist, sort_order10, 'Response' )


# In[48]:



    
#Have you ever used alcohol, marijuana, or other drugs not prescribed for you at your home WITH your parents or guardians knowing? (Not including a few sips of wine with family or for religious purposes.)  
# also binary


use_knowing = binary_Qer('use_parent_knowing')

use_parent_knowing_grade = by_grader_count_and_percenter('use_parent_knowing','grade_level', grades, YN_order2, 'Response')
binary_grade_fixer(use_parent_knowing_grade)

use_parent_knowing_by_year = by_grader_count_and_percenter('use_parent_knowing','drinker_risk', drinkinglist, YN_order2, 'Response')
binary_grade_fixer(use_parent_knowing_by_year)

use_knowing_yes_perc = use_knowing[use_knowing['Response']=='Yes']['Percentage (%)'].item()
use_knowing_10th = use_parent_knowing_grade[use_parent_knowing_grade['Response']=='Yes']['10th (%)'].item()
use_knowing_11th = use_parent_knowing_grade[use_parent_knowing_grade['Response']=='Yes']['11th (%)'].item()
use_knowing_12th = use_parent_knowing_grade[use_parent_knowing_grade['Response']=='Yes']['12th (%)'].item()

use_not_knowing = binary_Qer('use_parent_not_knowing')
use_parent_not_knowing_grade = by_grader_count_and_percenter('use_parent_not_knowing','grade_level', grades, YN_order2, 'Response')
binary_grade_fixer(use_parent_not_knowing_grade)
use_parent_not_knowing_by_risk = by_grader_count_and_percenter('use_parent_not_knowing','drinker_risk', drinkinglist, YN_order2, 'Response')
binary_grade_fixer(use_parent_not_knowing_by_risk)

use_not_knowing_yes_perc = use_not_knowing[use_not_knowing['Response']=='Yes']['Percentage (%)'].item()
use_not_knowing_10th = use_parent_not_knowing_grade[use_parent_not_knowing_grade['Response']=='Yes']['10th (%)'].item()
use_not_knowing_11th = use_parent_not_knowing_grade[use_parent_not_knowing_grade['Response']=='Yes']['11th (%)'].item()
use_not_knowing_12th = use_parent_not_knowing_grade[use_parent_knowing_grade['Response']=='Yes']['12th (%)'].item()

#Please indicate how often your family eats at least one meal together 

meals = create_count_and_percent_0check(clean, 'meals', 'Response', 'No', 'No', sort_order15)
meals =sort_by_customlist(meals, 'Response', sort_order15)

meals_grade = by_grader_count_and_percenter('meals','grade_level', grades, sort_order15, 'Response')
meals_by_risk = by_grader_count_and_percenter('meals','drinker_risk', drinkinglist, sort_order15, 'Response')

binary_answers =binary_multiQer(binarylist, 'Response')
binary_answers_grade= multi_by_grader( binarylist,'grade_level', grades, YN_order2, 'Response' )
binary_answers_by_risk= multi_by_grader( binarylist,'drinker_risk', drinkinglist, YN_order2, 'Response' )

Rarely_answers = multi_questioner(Rarely_Qs, sort_order16, 'Response' )
Rarely_answers_grade= multi_by_grader( Rarely_Qs,'grade_level', grades, sort_order16, 'Response' )
Rarely_answers_by_risk= multi_by_grader( Rarely_Qs,'drinker_risk', drinkinglist, sort_order16, 'Response' )

if _has_stress:
    AgreeDisagree4_answers = multi_questioner(AgreeDisagree4_q, sort_order10, 'Response' )
    AgreeDisagree4_answers_grade= multi_by_grader( AgreeDisagree4_q,'grade_level', grades, sort_order10, 'Response' )
    AgreeDisagree4_answers_by_risk= multi_by_grader( AgreeDisagree4_q,'drinker_risk', drinkinglist, sort_order10, 'Response' )
else:
    AgreeDisagree4_answers = []
    AgreeDisagree4_answers_grade = []
    AgreeDisagree4_answers_by_risk = []


valued_table =total_agree_disagree(AgreeDisagree2_answers_grade[0])

aca_table =total_agree_disagree(AgreeDisagree2_answers_grade[1])
staff_table =total_agree_disagree(AgreeDisagree2_answers_grade[2])
sincere_table =total_agree_disagree(AgreeDisagree2_answers_grade[3])
care_table =total_agree_disagree(AgreeDisagree2_answers_grade[4])

aca2_table =total_agree_disagree(  AgreeDisagree2_answers_grade[5])

if _has_stress:
    aca3_table = total_agree_disagree( AgreeDisagree4_answers_grade[0])
    parent_table = total_agree_disagree( AgreeDisagree4_answers_grade[1])
    over_table = total_agree_disagree( AgreeDisagree4_answers_grade[3])
    college_table = total_agree_disagree( AgreeDisagree4_answers_grade[2])
    exp_table = total_agree_disagree( AgreeDisagree4_answers_grade[4])
    bored_table = total_agree_disagree( AgreeDisagree4_answers_grade[5])
    fitin_table = total_agree_disagree( AgreeDisagree4_answers_grade[6])
else:
    _empty_agree = pd.DataFrame()
    aca3_table = _empty_agree
    parent_table = _empty_agree
    over_table = _empty_agree
    college_table = _empty_agree
    exp_table = _empty_agree
    bored_table = _empty_agree
    fitin_table = _empty_agree

totalagree_parents = []

for Q in AgreeDisagree3_answers:
    df = total_agree_disagree_nograde(Q)
    totalagree = df[df['index']=='Total Agree']['Count (N)'].item()
    totalagree_parents.append(totalagree)
    
    
for Q in ([use_knowing, use_not_knowing] + binary_answers):
    totalagree = Q[Q['Response']=='Yes']['Percentage (%)'].str.split('%').str[0].astype(float).item()
    totalagree_parents.append(totalagree)
    
    
totalagree_parents.append(meals[meals['Response']!='Never']['Percentage (%)'].str.split('%').str[0].astype(float).sum().item())


if sum(totalagree_parents)/ len(totalagree_parents) >= 75:
    familybulletpoint = 'The vast majority'
if (sum(totalagree_parents)/ len(totalagree_parents) >= 55) & (sum(totalagree_parents)/ len(totalagree_parents) < 75):
    familybulletpoint = 'The majority'
if (sum(totalagree_parents)/ len(totalagree_parents) >= 45) & (sum(totalagree_parents)/ len(totalagree_parents) < 55):
    familybulletpoint = 'About half'
if (sum(totalagree_parents)/ len(totalagree_parents) < 45):
    familybulletpoint = 'The minority'


# In[86]:


use_knowing


# In[87]:


use_not_knowing


# In[88]:


use_parent_not_knowing_grade


# In[49]:


if _has_stress:
    academic_stress = risk_by_subset_pot_andall_1012th('academic_stress',drinkinglist, sort_order10, 'Response' )
else:
    academic_stress = pd.DataFrame()
    logger.warning("Stress questions not on this survey; stress tables will be empty.")

def stressor_group_table(Q):
    """B6/B2/B4: Agree / Disagree by group, built by COLUMN NAME (the old
    groupby().sum() table was read by position and every value sat one column
    to the right), with each group's own denominator (the cannabis columns
    used to divide by both cannabis groups combined), from counts.
    Column order matches the template: [0] label, [2] higher-risk,
    [3] lower-risk, [4] non-drinkers, [5] never used cannabis,
    [6] cannabis in last 30 days (same layout as the K10 tables)."""
    groups = [
        ('Yearly or More Frequent Marijuana Use', clean['yearlypot'].notna()),
        ('High-risk (%)', clean['drinker_risk_10_12'].eq('High-risk')),
        ('Low-risk (%)', clean['drinker_risk_10_12'].eq('Low-risk')),
        ('Non-drinker (%)', clean['drinker_risk_10_12'].eq('Non-drinker')),
        ('Never used cannabis', clean['pot_risk30_never'].eq('Never used cannabis')),
        ('Cannabis use in last 30 days', clean['pot_risk30_never'].eq('Cannabis use in last 30 days')),
    ]
    rows = []
    for label, test in [('Agree', lambda ser: ser.isin(['Agree', 'Strongly agree'])),
                        ('Disagree', lambda ser: ser.isin(['Disagree', 'Strongly disagree']))]:
        row = {'agreedisagree': label}
        for col, gmask in groups:
            sub = clean.loc[gmask & clean[Q].notna(), Q]
            row[col] = pct(int(test(sub).sum()), len(sub), f'stressor {Q} {label}: {col}', f'{col} ({BODY_GROUP_LABEL}) who answered')
            if row[col] == '':
                row[col] = '-'
        rows.append(row)
    return pd.DataFrame(rows)[['agreedisagree'] + [g for g, _ in groups]]

if _has_stress:
    academic_stress_outputTABLE = stressor_group_table('academic_stress')
    parent_stress_outputTABLE = stressor_group_table('parent_stress')
else:
    academic_stress_outputTABLE = pd.DataFrame()
    parent_stress_outputTABLE = pd.DataFrame()


# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[51]:


if _has_coping:
    YNQ2_answers = multi_questioner(YNQ2s, YN_order2, 'Response' )
    YNQ2_answers_grade= multi_by_grader( YNQ2s,'grade_level', grades, YN_order2, 'Response' )
    YNQ2_answers_by_risk= multi_by_grader( YNQ2s,'drinker_risk', drinkinglist, YN_order2, 'Response' )
    coping_anx =YNQ2_answers_by_risk[0]
    coping_anx_perc = group_share('coping_alcohol', lambda ser: ser.eq('Yes'), group_col='drinker_risk', name='exec: higher-risk use alcohol to cope')
else:
    YNQ2_answers = []
    YNQ2_answers_grade = []
    YNQ2_answers_by_risk = []
    coping_anx = pd.DataFrame()
    coping_anx_perc = '0.0'
    logger.warning("Coping questions not on this survey; coping tables will be empty.")


### K10/Kessler scoring removed (not on this survey)


# In[54]:


if _has_coping:
    def coping(df):
        for Q in YNQ2s:
            if df[Q] == 'Yes':
                return 1
        return 0
    clean['coping'] = clean.apply(coping, axis=1)
else:
    clean['coping'] = 0


# In[55]:





### K10 analysis tables, coping K10 breakdowns, and formatting all removed (not on this survey)


# In[ ]:

##############################
### Gambling and Loot Box Analysis
##############################

_gambling_sort = ['0 days', '1-2 days', '3-5 days', '6-9 days', '10-19 days', '20-30 days']

_has_gambling = 'gambling_30_days' in clean.columns and clean['gambling_30_days'].notna().any()
_has_lootbox = 'lootbox_30_days' in clean.columns and clean['lootbox_30_days'].notna().any()

if _has_gambling:
    clean['gambler'] = clean['gambling_30_days'].apply(
        lambda x: 'Yes' if pd.notna(x) and str(x).strip() not in ('0 days', '', 'nan') else 'No')
    _actual_gambling_vals = [v for v in _gambling_sort if v in clean['gambling_30_days'].unique()]
    gambling_all = create_count_and_percent_0check(clean, 'gambling_30_days', 'Response', 'No', 'No', _actual_gambling_vals)
    gambling_all = sort_by_customlist(gambling_all, 'Response', _actual_gambling_vals)
    gambling_grade = by_grader_count_and_percenter('gambling_30_days', 'grade_level', grades, _actual_gambling_vals, 'Response')
    gambling_by_risk = by_grader_count_and_percenter('gambling_30_days', 'drinker_risk', drinkinglist, _actual_gambling_vals, 'Response')
    gambling_any_perc = round(float(group_share('gambler', lambda s: s.eq('Yes'), name='any gambling past 30 days')), 1)
else:
    gambling_all = pd.DataFrame(columns=['Response', 'Count (N)', 'Percentage (%)'])
    gambling_grade = pd.DataFrame()
    gambling_by_risk = pd.DataFrame()
    gambling_any_perc = '0.0'
    clean['gambler'] = 'No'

if _has_lootbox:
    clean['lootboxer'] = clean['lootbox_30_days'].apply(
        lambda x: 'Yes' if pd.notna(x) and str(x).strip() not in ('0 days', '', 'nan') else 'No')
    _actual_lootbox_vals = [v for v in _gambling_sort if v in clean['lootbox_30_days'].unique()]
    lootbox_all = create_count_and_percent_0check(clean, 'lootbox_30_days', 'Response', 'No', 'No', _actual_lootbox_vals)
    lootbox_all = sort_by_customlist(lootbox_all, 'Response', _actual_lootbox_vals)
    lootbox_grade = by_grader_count_and_percenter('lootbox_30_days', 'grade_level', grades, _actual_lootbox_vals, 'Response')
    lootbox_by_risk = by_grader_count_and_percenter('lootbox_30_days', 'drinker_risk', drinkinglist, _actual_lootbox_vals, 'Response')
    lootbox_any_perc = round(float(group_share('lootboxer', lambda s: s.eq('Yes'), name='any loot box past 30 days')), 1)
else:
    lootbox_all = pd.DataFrame(columns=['Response', 'Count (N)', 'Percentage (%)'])
    lootbox_grade = pd.DataFrame()
    lootbox_by_risk = pd.DataFrame()
    lootbox_any_perc = '0.0'
    clean['lootboxer'] = 'No'

# Correlation: gambling and loot box by drinking risk
gambling_by_risk_10_12 = by_grader_count_and_percenter(
    'gambler', 'drinker_risk_10_12', drinkinglist, ['Yes', 'No'], 'Response') if _has_gambling else pd.DataFrame()
lootbox_by_risk_10_12 = by_grader_count_and_percenter(
    'lootboxer', 'drinker_risk_10_12', drinkinglist, ['Yes', 'No'], 'Response') if _has_lootbox else pd.DataFrame()

# Cross-tab: loot box use among gamblers vs non-gamblers
if _has_gambling and _has_lootbox:
    _gamblers = clean[clean['gambler'] == 'Yes']
    _non_gamblers = clean[clean['gambler'] == 'No']
    _n_gamblers = len(_gamblers)
    _n_non_gamblers = len(_non_gamblers)
    lootbox_among_gamblers = round(float(pct(len(_gamblers[_gamblers['lootboxer'] == 'Yes']), _n_gamblers)), 1) if _n_gamblers > 0 else 0.0
    lootbox_among_non_gamblers = round(float(pct(len(_non_gamblers[_non_gamblers['lootboxer'] == 'Yes']), _n_non_gamblers)), 1) if _n_non_gamblers > 0 else 0.0
    gambling_lootbox_cross = pd.DataFrame({
        'Group': ['Students who gambled', 'Students who did not gamble'],
        'Also used loot boxes (%)': [lootbox_among_gamblers, lootbox_among_non_gamblers],
        'N': [_n_gamblers, _n_non_gamblers]
    })
else:
    lootbox_among_gamblers = '0.0'
    lootbox_among_non_gamblers = '0.0'
    gambling_lootbox_cross = pd.DataFrame()

# Gambling among drinkers vs non-drinkers
if _has_gambling:
    _drinkers_30 = clean[clean['drank_30d'] == True] if 'drank_30d' in clean.columns else clean[clean['one_drink_past_30_days'] == 'Yes']
    _non_drinkers_30 = clean[~clean.index.isin(_drinkers_30.index)]
    _n_drinkers = len(_drinkers_30)
    _n_non_drinkers = len(_non_drinkers_30)
    gambling_among_drinkers = round(float(pct(len(_drinkers_30[_drinkers_30['gambler'] == 'Yes']), _n_drinkers)), 1) if _n_drinkers > 0 else 0.0
    gambling_among_non_drinkers = round(float(pct(len(_non_drinkers_30[_non_drinkers_30['gambler'] == 'Yes']), _n_non_drinkers)), 1) if _n_non_drinkers > 0 else 0.0
    gambling_drinking_cross = pd.DataFrame({
        'Group': ['Drank in past 30 days', 'Did not drink in past 30 days'],
        'Also gambled (%)': [gambling_among_drinkers, gambling_among_non_drinkers],
        'N': [_n_drinkers, _n_non_drinkers]
    })
else:
    gambling_among_drinkers = '0.0'
    gambling_among_non_drinkers = '0.0'
    gambling_drinking_cross = pd.DataFrame()

logger.info(f"Gambling: {gambling_any_perc}% any past 30 days; Loot box: {lootbox_any_perc}% any past 30 days")


# In[ ]:





# In[ ]:





# In[ ]:





# In[60]:


############################## ############################## ##############################
### Executive Report
############################## ############################## ##############################


#######################################
### Read in Stats page I made from the PDFS
#########################################



percent8thgraders_ever_drank = new_percenter( age_at_first_grade, (age_at_first_grade['Age']== 'I have never had a whole drink of alcohol'), '8th (%)')
percent10thgraders_ever_drank = new_percenter( age_at_first_grade, (age_at_first_grade['Age']== 'I have never had a whole drink of alcohol'), '10th (%)')
percent12thgraders_ever_drank = new_percenter( age_at_first_grade, (age_at_first_grade['Age']== 'I have never had a whole drink of alcohol'), '12th (%)')

percent_never_drank_allgrades = (age_at_first_drinkdf[age_at_first_drinkdf['Age At First Drink']=='I have never had a whole drink of alcohol']['Percentage (%)'].str.split('%').str[0]).astype(float).item()
percent_never_drank_11th = (age_at_first_grade[age_at_first_grade['Age']=='I have never had a whole drink of alcohol']['11th (%)'].str.split('%').str[0]).astype(float).item()
percent_never_drank_12th= (age_at_first_grade[age_at_first_grade['Age']=='I have never had a whole drink of alcohol']['12th (%)'].str.split('%').str[0]).astype(float).item()

if '7th' in clean['grade_level'].values:
    lifetime7th = (age_at_first_grade[age_at_first_grade['Age']!='I have never had a whole drink of alcohol']['7th (%)'].str.split('%').str[0]).astype(float).sum()
elif '8th' in clean['grade_level'].values and '8th (%)' in age_at_first_grade.columns:
    # Use 8th grade as the lowest grade when no 7th graders exist
    lifetime7th = (age_at_first_grade[age_at_first_grade['Age']!='I have never had a whole drink of alcohol']['8th (%)'].str.split('%').str[0]).astype(float).sum()
    lifetime7th = round(lifetime7th, 1)
else:
    lifetime7th =''

lifetime12th = (age_at_first_grade[age_at_first_grade['Age']!='I have never had a whole drink of alcohol']['12th (%)'].str.split('%').str[0]).astype(float).sum()
lifetime12th = round(lifetime12th,1)

##############################
### B1/B2/B4: U.S. norm comparison tables rebuilt from the canonical flags.
### Base = ALL students in the grade (and sex), the same base MTF uses, so
### past-year use can never exceed lifetime use, and the Total row always exists.
##############################
def sex_rate_table(flag_col, grade_cols=('8th', '10th', '12th')):
    rows = []
    for sex_label, sex_val in [('Total', None), ('Male', 'Male'), ('Female', 'Female')]:
        row = {'sex': sex_label}
        for g in grade_cols:
            d = grade_size(g, sex_val)
            row[g] = flag_rate(flag_col, g, sex_val, name=f'{flag_col} {g} {sex_label}') if d > 0 else np.nan
            if row[g] == '':
                row[g] = np.nan
        rows.append(row)
    return pd.DataFrame(rows)[['sex'] + list(grade_cols)]

everdrank_grade_sex = sex_rate_table('drank_lifetime')
alcohol_past_12monthbysex = sex_rate_table('drank_12m')
alcohol_past_30daysbysex = sex_rate_table('drank_30d')
pot_lifetime = sex_rate_table('pot_lifetime')
POTPAST12 = sex_rate_table('pot_12m')
POTPAST30 = sex_rate_table('pot_30d')
VAPEPAST12 = sex_rate_table('vape_12m')
VAPEPAST30 = sex_rate_table('vape_30d')

def _total_of(tbl, g):
    v = tbl.loc[tbl['sex'] == 'Total', g].iloc[0]
    return _wrap_scalar('N/A') if pd.isna(v) else _wrap_scalar(float(v))

percent8thgraders_ever_drank = _total_of(everdrank_grade_sex, '8th')
percent10thgraders_ever_drank = _total_of(everdrank_grade_sex, '10th')
percent12thgraders_ever_drank = _total_of(everdrank_grade_sex, '12th')
percent8thgraders_ever_pot = _total_of(pot_lifetime, '8th')
percent10thgraders_ever_pot = _total_of(pot_lifetime, '10th')
percent12thgraders_ever_pot = _total_of(pot_lifetime, '12th')

clean['never_drank'] = ~clean['drank_lifetime']
percent_never_drank_allgrades = flag_rate('never_drank', name='never had a whole drink, all students')
percent_never_drank_11th = flag_rate('never_drank', '11th') if grade_size('11th') else ''
percent_never_drank_12th = flag_rate('never_drank', '12th') if grade_size('12th') else ''


# School values per MTF timeline, so each highlight bullet compares
# lifetime-to-lifetime, past-12-months-to-past-12-months, and
# past-30-days-to-past-30-days (previously the school's LIFETIME rate was
# compared against every timeline's norm, so a school could "significantly
# exceed" the past-30-days norm while actually being below it).
def _bullet_school_vals(lifetime_val, df12, df30, gradecol):
    def _flt(v):
        try:
            if hasattr(v, 'item'):
                v = v.item()
            elif hasattr(v, 'iloc'):
                v = v.iloc[0]
            return float(str(v).replace('%', ''))
        except Exception:
            return None
    def _total_cell(_df, _col):
        try:
            return _flt(_df.loc[_df['sex'] == 'Total', _col].iloc[0])
        except Exception:
            return None
    return {'lifetime': _flt(lifetime_val),
            'past-12-months': _total_cell(df12, gradecol),
            'past-30-days': _total_cell(df30, gradecol)}

if '8th' in clean['grade_level'].values:
    albullet1 = create_bullet_point(MTF[(MTF['variable']=='Alcohol') & (MTF['Type']=='Total')], percent8thgraders_ever_drank, '8th',
        per_timeline=_bullet_school_vals(percent8thgraders_ever_drank, alcohol_past_12monthbysex, alcohol_past_30daysbysex, '8th'))
else:
    albullet1 = 'is'
albullet2 = create_bullet_point(MTF[(MTF['variable']=='Alcohol') & (MTF['Type']=='Total')], percent10thgraders_ever_drank, '10th',
    per_timeline=_bullet_school_vals(percent10thgraders_ever_drank, alcohol_past_12monthbysex, alcohol_past_30daysbysex, '10th'))
albullet3 = create_bullet_point(MTF[(MTF['variable']=='Alcohol') & (MTF['Type']=='Total')], percent12thgraders_ever_drank, '12th',
    per_timeline=_bullet_school_vals(percent12thgraders_ever_drank, alcohol_past_12monthbysex, alcohol_past_30daysbysex, '12th'))

albullet1withare = ('are' + albullet1.split('is', 1)[1]) if 'is' in albullet1 else albullet1.replace('is ', 'are ', 1)


if '8th' in clean['grade_level'].values:
    maribullet1 = create_bullet_point(MTF[(MTF['variable']=='Marijuana') & (MTF['Type']=='Total')], percent8thgraders_ever_pot, '8th',
        per_timeline=_bullet_school_vals(percent8thgraders_ever_pot, POTPAST12, POTPAST30, '8th'))
else:
    maribullet1 =''

maribullet2 =create_bullet_point(MTF[(MTF['variable']=='Marijuana') & (MTF['Type']=='Total')], percent10thgraders_ever_pot, '10th',
    per_timeline=_bullet_school_vals(percent10thgraders_ever_pot, POTPAST12, POTPAST30, '10th'))
maribullet3 = create_bullet_point(MTF[(MTF['variable']=='Marijuana') & (MTF['Type']=='Total')], percent12thgraders_ever_pot, '12th',
    per_timeline=_bullet_school_vals(percent12thgraders_ever_pot, POTPAST12, POTPAST30, '12th'))


if '8th' in clean['grade_level'].values:
    ALLSUBSTANCE8th_30_MTF = pd.merge( ALLSUBSTANCE8th_30, fix_MTF_naming(MTF[(MTF['timeline']=='past-30-days') & (MTF['Type']=='Total')][['8th', 'variable']]), left_on='substance', right_on='variable')

    ALLSUBSTANCE8th_30_MTF =ALLSUBSTANCE8th_30_MTF[['substance', '8th',  'percent']]
    ALLSUBSTANCE8th_30_MTF.columns = ['substance',  'MTF', 'school']
    ALLSUBSTANCE8th_30_MTF['school'] =ALLSUBSTANCE8th_30_MTF['school'].astype(float)
    ALLSUBSTANCE8th_30_MTF = ALLSUBSTANCE8th_30_MTF.sort_values('school', ascending=False)
else:
    ALLSUBSTANCE8th_30_MTF =pd.DataFrame()


ALLSUBSTANCE10th_30_MTF = pd.merge( ALLSUBSTANCE10th_30, fix_MTF_naming(MTF[(MTF['timeline']=='past-30-days') & (MTF['Type']=='Total')][['10th', 'variable']]), left_on='substance', right_on='variable')

ALLSUBSTANCE10th_30_MTF =ALLSUBSTANCE10th_30_MTF[['substance', '10th',  'percent']]
ALLSUBSTANCE10th_30_MTF.columns = ['substance',  'MTF', 'school']
ALLSUBSTANCE10th_30_MTF['school'] =ALLSUBSTANCE10th_30_MTF['school'].astype(float)
ALLSUBSTANCE10th_30_MTF = ALLSUBSTANCE10th_30_MTF.sort_values('school', ascending=False)


ALLSUBSTANCE12th_30_MTF = pd.merge( ALLSUBSTANCE12th_30, fix_MTF_naming(MTF[(MTF['timeline']=='past-30-days') & (MTF['Type']=='Total')][['12th', 'variable']]), left_on='substance', right_on='variable')

ALLSUBSTANCE12th_30_MTF =ALLSUBSTANCE12th_30_MTF[['substance', '12th',  'percent']]
ALLSUBSTANCE12th_30_MTF.columns = ['substance',  'MTF', 'school']
ALLSUBSTANCE12th_30_MTF['school'] =ALLSUBSTANCE12th_30_MTF['school'].astype(float)
ALLSUBSTANCE12th_30_MTF = ALLSUBSTANCE12th_30_MTF.sort_values('school', ascending=False)

#8th-Grade Substance Use Over the Past 12 Months Compared to U.S. Norms 
if '8th' in clean['grade_level'].values:
    ALLSUBSTANCE8th_12_MTF = pd.merge( ALLSUBSTANCE8th_12, fix_MTF_naming(MTF[(MTF['timeline']=='past-12-months') & (MTF['Type']=='Total')][['8th', 'variable']]), left_on='substance', right_on='variable')

    ALLSUBSTANCE8th_12_MTF =ALLSUBSTANCE8th_12_MTF[['substance', '8th',  'percent']]
    ALLSUBSTANCE8th_12_MTF.columns = ['substance',  'MTF', 'school']
    ALLSUBSTANCE8th_12_MTF['school'] =ALLSUBSTANCE8th_12_MTF['school'].astype(float)
    ALLSUBSTANCE8th_12_MTF = ALLSUBSTANCE8th_12_MTF.sort_values('school', ascending=False)
    ALLSUBSTANCE8th_12_MTF =ALLSUBSTANCE8th_12_MTF.drop_duplicates()
else:
    ALLSUBSTANCE8th_12_MTF = pd.DataFrame()

ALLSUBSTANCE10th_12_MTF = pd.merge( ALLSUBSTANCE10th_12, fix_MTF_naming(MTF[(MTF['timeline']=='past-12-months') & (MTF['Type']=='Total')][['10th', 'variable']]), left_on='substance', right_on='variable')

ALLSUBSTANCE10th_12_MTF =ALLSUBSTANCE10th_12_MTF[['substance', '10th',  'percent']]
ALLSUBSTANCE10th_12_MTF.columns = ['substance',  'MTF', 'school']
ALLSUBSTANCE10th_12_MTF['school'] =ALLSUBSTANCE10th_12_MTF['school'].astype(float)
ALLSUBSTANCE10th_12_MTF = ALLSUBSTANCE10th_12_MTF.sort_values('school', ascending=False)
ALLSUBSTANCE10th_12_MTF =ALLSUBSTANCE10th_12_MTF.drop_duplicates()

ALLSUBSTANCE12th_12_MTF = pd.merge( ALLSUBSTANCE12th_12, fix_MTF_naming(MTF[(MTF['timeline']=='past-12-months') & (MTF['Type']=='Total')][['12th', 'variable']]), left_on='substance', right_on='variable')

ALLSUBSTANCE12th_12_MTF =ALLSUBSTANCE12th_12_MTF[['substance', '12th',  'percent']]
ALLSUBSTANCE12th_12_MTF.columns = ['substance',  'MTF', 'school']
ALLSUBSTANCE12th_12_MTF['school'] =ALLSUBSTANCE12th_12_MTF['school'].astype(float)
ALLSUBSTANCE12th_12_MTF = ALLSUBSTANCE12th_12_MTF.sort_values('school', ascending=False)
ALLSUBSTANCE12th_12_MTF =ALLSUBSTANCE12th_12_MTF.drop_duplicates()

if '8th' in clean['grade_level'].values:
    femsub12 = fix_MTF_naming(MTF[(MTF['timeline']=='past-12-months') & (MTF['Type']=='Female')][['8th', 'variable']])

    femsub12.columns = ['8th MTF Females', 'substance']


    mascsub12 = fix_MTF_naming(MTF[(MTF['timeline']=='past-12-months') & (MTF['Type']=='Male')][['8th', 'variable']])

    mascsub12.columns = ['8th MTF Males', 'substance']

    sub12_by_sex_MTF8th = pd.merge(femsub12,mascsub12, on ='substance')
    sub12_by_sex_MTFcombo8th =pd.merge(sub12_by_sex_MTF8th, sub12months_sex8th, on='substance')
    sub12_by_sex_MTFcombo8th =sub12_by_sex_MTFcombo8th[['substance', '8th MTF Males', '8th Male', '8th MTF Females', '8th Female']]

    sub12_by_sex_MTFcombo8th =sub12_by_sex_MTFcombo8th.drop_duplicates().sort_values('8th MTF Males', ascending=False, key=lambda _s: pd.to_numeric(_s, errors='coerce'))
else:
    sub12_by_sex_MTFcombo8th = pd.DataFrame(columns = ['substance', '8th MTF Males', '8th Survey Males', '8th MTF Females', '8th Survey Females'])


femsub12_10 = fix_MTF_naming(MTF[(MTF['timeline']=='past-12-months') & (MTF['Type']=='Female')][['10th', 'variable']])
femsub12_10.columns = ['10th MTF Females', 'substance']
mascsub12_10 = fix_MTF_naming(MTF[(MTF['timeline']=='past-12-months') & (MTF['Type']=='Male')][['10th', 'variable']])
mascsub12_10.columns = ['10th MTF Males', 'substance']
sub12_by_sex_MTF10th = pd.merge(femsub12_10,mascsub12_10, on ='substance')

sub12_by_sex_MTFcombo10th =pd.merge(sub12_by_sex_MTF10th, sub12months_sex10th, on='substance')
# sub12_by_sex_MTFcombo10th =sub12_by_sex_MTFcombo10th[['substance', '10th MTF Males', (      '10th',   'Male'), '10th MTF Females',  (      '10th', 'Female') ]]
sub12_by_sex_MTFcombo10th =sub12_by_sex_MTFcombo10th[['substance', '10th MTF Males', '10th Male', '10th MTF Females',  '10th Female' ]]
sub12_by_sex_MTFcombo10th =sub12_by_sex_MTFcombo10th.drop_duplicates().sort_values('10th MTF Males', ascending=False, key=lambda _s: pd.to_numeric(_s, errors='coerce'))



femsub12_12 = fix_MTF_naming(MTF[(MTF['timeline']=='past-12-months') & (MTF['Type']=='Female')][['12th', 'variable']])
femsub12_12.columns = ['12th MTF Females', 'substance']
mascsub12_12 = fix_MTF_naming(MTF[(MTF['timeline']=='past-12-months') & (MTF['Type']=='Male')][['12th', 'variable']])
mascsub12_12.columns = ['12th MTF Males', 'substance']
sub12_by_sex_MTF12th = pd.merge(femsub12_12,mascsub12_12, on ='substance')
sub12_by_sex_MTFcombo12th =pd.merge(sub12_by_sex_MTF12th, sub12months_sex12th, on='substance')
# sub12_by_sex_MTFcombo12th =sub12_by_sex_MTFcombo12th[['substance', '12th MTF Males', (      '12th',   'Male'), '12th MTF Females',  (      '12th', 'Female') ]]
sub12_by_sex_MTFcombo12th =sub12_by_sex_MTFcombo12th[['substance', '12th MTF Males', '12th Male', '12th MTF Females',  '12th Female' ]]
sub12_by_sex_MTFcombo12th =sub12_by_sex_MTFcombo12th.drop_duplicates().sort_values('12th MTF Males', ascending=False, key=lambda _s: pd.to_numeric(_s, errors='coerce'))


# In[ ]:




    


# In[ ]:





# In[ ]:





# In[ ]:





# In[61]:


################### COME BACK TO THIIIIISSSSSSSSS
sub12_by_sex_MTFcombo8th.columns = ['substance', '8th MTF Males', '8th Survey Males', '8th MTF Females', '8th Survey Females']
#sub12_by_sex_MTFcombo8th['substance'] = sub12_by_sex_MTFcombo8th['substance'].str.lower()
for col in ['8th MTF Males', '8th Survey Males', '8th MTF Females', '8th Survey Females']:
    sub12_by_sex_MTFcombo8th[col] = pd.to_numeric(sub12_by_sex_MTFcombo8th[col], errors='coerce')

sub12_by_sex_MTFcombo8th[['8th Survey Males', '8th Survey Females']] = sub12_by_sex_MTFcombo8th[['8th Survey Males', '8th Survey Females']].fillna(0.0)
    
sub12_by_sex_MTFcombo10th.columns = ['substance', '10th MTF Males', '10th Survey Males', '10th MTF Females', '10th Survey Females']
#sub12_by_sex_MTFcombo10th['substance'] = sub12_by_sex_MTFcombo10th['substance'].str.lower()

for col in ['10th MTF Males', '10th Survey Males', '10th MTF Females', '10th Survey Females']:
    sub12_by_sex_MTFcombo10th[col] = pd.to_numeric(sub12_by_sex_MTFcombo10th[col], errors='coerce')
    
sub12_by_sex_MTFcombo10th[['10th Survey Males', '10th Survey Females']] = sub12_by_sex_MTFcombo10th[['10th Survey Males', '10th Survey Females']].fillna(0.0)

sub12_by_sex_MTFcombo12th.columns = ['substance', '12th MTF Males', '12th Survey Males', '12th MTF Females', '12th Survey Females']
#sub12_by_sex_MTFcombo12th['substance']= sub12_by_sex_MTFcombo12th['substance'].str.lower()

for col in [ '12th MTF Males', '12th Survey Males', '12th MTF Females', '12th Survey Females']:
    sub12_by_sex_MTFcombo12th[col] = pd.to_numeric(sub12_by_sex_MTFcombo12th[col], errors='coerce')
    
sub12_by_sex_MTFcombo12th[['12th Survey Males', '12th Survey Females']] = sub12_by_sex_MTFcombo12th[['12th Survey Males', '12th Survey Females']].fillna(0.0)

def substance_grade_bullets(df, grade):
    outputstringssexbullets = []
    for sex in ['Males', 'Females']:
        below = []
        sigabove = []
        above = []
        onpar = []
        cats_exist = []
        for index, row in df.iterrows():
            if pd.isna(row[f'{grade} MTF '+ sex]):
                continue



            # B9: same small-sample guard as the other U.S. norm comparisons,
            # using the number of students of this sex in this grade.
            _sex_label = 'Male' if sex == 'Males' else 'Female'
            _v = norm_verdict(pd.to_numeric(row[f'{grade} Survey ' + sex], errors='coerce'),
                              pd.to_numeric(row[f'{grade} MTF ' + sex], errors='coerce'),
                              grade_size(grade, _sex_label))
            if _v == 'onpar':
                onpar.append(row['substance'])
            elif _v == 'below':
                below.append(row['substance'])
            elif _v == 'sigabove':
                sigabove.append(row['substance'])
            elif _v == 'above':
                above.append(row['substance'])

        # B7: grammatical sentence ("is significantly exceeds with" before)
        outputstring1 = f"Substance use by {grade} grade {sex.lower()} "
        if len(below) > 0:
            cats_exist.append('is below U.S. norms for ' + _join_and(below))
        if len(onpar)>0:
            cats_exist.append('is on par with U.S. norms for ' + _join_and(onpar))
        if len(above)>0:
            cats_exist.append('is above U.S. norms for ' + _join_and(above))
        if len(sigabove)>0:
            cats_exist.append('is well above U.S. norms for ' + _join_and(sigabove))
        if not cats_exist or grade not in set(clean['grade_level'].dropna()):
            outputstringssexbullets.append('')
            continue
        outputstringfinal = outputstring1 + '; it '.join(cats_exist) + '.'
        # B7: plain substance names in sentences (no table footnote markers)
        outputstringfinal = (outputstringfinal.replace('Electronic Vaporizers**', 'electronic vaporizers')
                             .replace('Over the Counter Cold Medicine', 'over-the-counter cold medicine')
                             .replace('Smokeless', 'smokeless tobacco').replace('Meth', 'methamphetamine')
                             .replace('Marijuana', 'cannabis').replace('Cigarettes', 'cigarettes').replace('Ecstasy', 'ecstasy')
                             .replace('Heroin', 'heroin').replace('Narcotics', 'narcotics').replace('Amphetamines', 'amphetamines')
                             .replace('Sedatives', 'sedatives').replace('Cocaine', 'cocaine').replace('Hallucinogens', 'hallucinogens')
                             .replace('Inhalants', 'inhalants').replace('Steroids', 'steroids'))
        outputstringssexbullets.append(outputstringfinal)
    return outputstringssexbullets


# In[62]:


outputstringssexbullets = substance_grade_bullets(sub12_by_sex_MTFcombo8th, '8th')
outputstringssexbullets10th = substance_grade_bullets(sub12_by_sex_MTFcombo10th, '10th')
outputstringssexbullets12th = substance_grade_bullets(sub12_by_sex_MTFcombo12th, '12th')


# In[ ]:





# In[ ]:





# In[ ]:





# In[63]:


MTF_total = MTF[MTF['Type']=='Total']
MTF_total.columns = ['Type','8th', '10th', '12th', 'substance', 'timeline','Year' ]
MTFtotal =MTF_total.melt(id_vars=[ 'substance','timeline'], value_vars=['8th', '10th', '12th'])
MTFtotal.columns = [ 'substance','timeline', 'grade_level', 'MTFperc']

substanceMTFcombo =pd.merge(substances_DF, MTFtotal, on=['substance', 'timeline', 'grade_level'], how='outer')
substanceMTFcombo.columns = ['schoolperc', 'question', 'timeline', 'grade', 'substance', 'MTFperc']
substanceMTFcombo = substanceMTFcombo.drop_duplicates()
substanceMTFcombo['schoolperc'] = pd.to_numeric(substanceMTFcombo['schoolperc'], errors='coerce')
substanceMTFcombo['MTFperc'] = pd.to_numeric(substanceMTFcombo['MTFperc'], errors='coerce')


allsubstancebullets = []

for substance in substanceMTFcombo['substance'].unique():
    #print(substance)
    for grade in substanceMTFcombo['grade'].unique():
        #print(grade)

        test= substanceMTFcombo[(substanceMTFcombo['substance']==substance) & (substanceMTFcombo['grade']==grade)]
        below = []
        sigabove = []
        above = []
        onpar = []
        for index, row in test.iterrows():

            error = pd.to_numeric(row['MTFperc']) * .05
            error2 = pd.to_numeric(row['MTFperc']) * .075

            if (row['schoolperc'] <= (row['MTFperc']+ error)) & ( (row['MTFperc']- error)<=(row['schoolperc'] )):
                onpar.append(row['timeline'])
            if(  (row['MTFperc'] - error) > (row['schoolperc'])) | (row['MTFperc']==0):
                below.append(row['timeline'])
            elif row['schoolperc'] > (row['MTFperc']+ error2) :
                sigabove.append(row['timeline'])
            elif (row['schoolperc'] > (row['MTFperc']+ error)) & ( (row['MTFperc']+ error2) <=row['schoolperc']) :
                above.append(row['timeline'])

        outputstring = []    

        if (len(below)==1) | (len(below)==2):
            outputstring.append(f"is below U.S. norms for {' and '.join(below)}")
        if (len(below)>=3) :
            outputstring.append(f"is below U.S. norms for {', '.join(below[:2])}, and {below[2]}")  

        if (len(onpar)==1) | (len(onpar)==2):
            outputstring.append(f"is on par with U.S. norms for {' and '.join(onpar)}")
        if (len(onpar)>=3) :
            outputstring.append(f"is on par with U.S. norms for {', '.join(onpar[:2])}, and {onpar[2]}")  

        if (len(above)==1) | (len(above)==2):
            outputstring.append(f"exceeds U.S. norms for {' and '.join(above)}")
        if (len(above)>=3) :
            outputstring.append(f"exceeds U.S. norms for {', '.join(above[:2])}, and {above[2]}")  

        if (len(sigabove)==1) | (len(sigabove)==2):
            outputstring.append(f"significantly exceeds U.S. norms for {' and '.join(sigabove)}")
        if (len(sigabove)>=3) :
            outputstring.append(f"significantly exceeds U.S. norms for {', '.join(sigabove[:2])}, and {sigabove[2]}") 

        if len(below + above + onpar +sigabove)>0:
            _subname = (str(substance).replace('**', '').replace('Over the Counter Cold Medicine', 'Over-the-counter cold medicine')
                        .replace('Smokeless', 'Smokeless tobacco').replace('Meth', 'Methamphetamine'))
            allsubstancebullets.append(f"{_subname} use by {School} {grade} graders "+ ' and '.join(outputstring) +' measures.')
            
            
mostpopular_substances =substances_DF[~substances_DF['substance'].isin(['Alcohol', 'Nicotine', 'Cigarettes',
                                                'Marijuana', 'Electronic Vaporizers**', 'Smokeless'])].sort_values('percent', ascending=False).reset_index()['substance'][:6].unique().tolist()


_mtf_grade_cols = ['Type', '8th', '10th', '12th']
alcohollifetumetable = pd.merge(everdrank_grade_sex, MTF[(MTF['variable']=='Alcohol') & (MTF['timeline']=='lifetime')][_mtf_grade_cols], left_on='sex', right_on='Type', suffixes = ["", " MTF"], how="left")
alcoholl12monthtstable = pd.merge( alcohol_past_12monthbysex, MTF[(MTF['variable']=='Alcohol') & (MTF['timeline']=='past-12-months')][_mtf_grade_cols], left_on='sex', right_on='Type', suffixes = ["", " MTF"], how="left")
alcoholl30daysstable =  pd.merge(alcohol_past_30daysbysex, MTF[(MTF['variable']=='Alcohol') & (MTF['timeline']=='past-30-days')][_mtf_grade_cols], left_on='sex', right_on='Type', suffixes = ["", " MTF"], how="left")


pot_lifetime_MTF =pd.merge(pot_lifetime, MTF[(MTF['variable']=='Marijuana') & (MTF['timeline']=='lifetime')][_mtf_grade_cols], left_on='sex', right_on='Type', suffixes = ["", " MTF"], how="left").fillna(0.0)
pot_12_MTF =pd.merge(POTPAST12, MTF[(MTF['variable']=='Marijuana') & (MTF['timeline']=='past-12-months')][_mtf_grade_cols], left_on='sex', right_on='Type', suffixes = ["", " MTF"], how="left").fillna(0.0)
pot_30_MTF =pd.merge(POTPAST30, MTF[(MTF['variable']=='Marijuana') & (MTF['timeline']=='past-30-days')][_mtf_grade_cols], left_on='sex', right_on='Type', suffixes = ["", " MTF"], how="left").fillna(0.0)


# MTF files name the vaping variable differently by year
_vape_mtf_var = 'Any Vaping' if (MTF['variable'] == 'Any Vaping').any() else 'Vaping Nicotine'
VAPE_12_MTF =pd.merge(VAPEPAST12, MTF[(MTF['variable']==_vape_mtf_var) & (MTF['timeline']=='past-12-months')][_mtf_grade_cols], left_on='sex', right_on='Type', suffixes = ["", " MTF"], how="left").fillna(0.0)
VAPE_30_MTF =pd.merge(VAPEPAST30, MTF[(MTF['variable']==_vape_mtf_var) & (MTF['timeline']=='past-30-days')][_mtf_grade_cols], left_on='sex', right_on='Type', suffixes = ["", " MTF"], how="left").fillna(0.0)


# In[ ]:





# In[ ]:





# In[ ]:





# In[ ]:





# In[64]:


# B7: keep the school's Total row even when the MTF file has no matching row
# (the lifetime alcohol table lost its Total row this way).
for _t in [alcohollifetumetable, alcoholl12monthtstable, alcoholl30daysstable, pot_lifetime_MTF,
           pot_12_MTF, pot_30_MTF, VAPE_12_MTF, VAPE_30_MTF]:
    _t['Type'] = _t['sex']

_numeric_grade_cols = ['8th', '10th', '12th', '8th MTF', '10th MTF', '12th MTF']
for col in _numeric_grade_cols:
    alcohollifetumetable[col] = pd.to_numeric(alcohollifetumetable[col], errors='coerce').pipe(round, 1)   # B4: one decimal, not whole numbers shown as '.0'
    alcoholl12monthtstable[col]= pd.to_numeric(alcoholl12monthtstable[col], errors='coerce').pipe(round, 1)   # B4: one decimal, not whole numbers shown as '.0'
    alcoholl30daysstable[col] = pd.to_numeric(alcoholl30daysstable[col], errors='coerce').pipe(round, 1)   # B4: one decimal, not whole numbers shown as '.0'
    pot_lifetime_MTF[col]= pd.to_numeric(pot_lifetime_MTF[col], errors='coerce').pipe(round, 1)   # B4: one decimal, not whole numbers shown as '.0'
    pot_12_MTF[col]= pd.to_numeric(pot_12_MTF[col], errors='coerce').pipe(round, 1)   # B4: one decimal, not whole numbers shown as '.0'
    pot_30_MTF[col]= pd.to_numeric(pot_30_MTF[col], errors='coerce').pipe(round, 1)   # B4: one decimal, not whole numbers shown as '.0'
    VAPE_12_MTF[col]= pd.to_numeric(VAPE_12_MTF[col], errors='coerce').pipe(round, 1)   # B4: one decimal, not whole numbers shown as '.0'
    VAPE_30_MTF[col]= pd.to_numeric(VAPE_30_MTF[col], errors='coerce').pipe(round, 1)   # B4: one decimal, not whole numbers shown as '.0'


# In[ ]:





# In[65]:


vapebullets = []
gradeslist_vape_bullets = []


for grade in [g for g in ['8th', '10th', '12th'] if g in grades]:
    print(VAPE_30_MTF)


    _sel30 = VAPE_30_MTF[(VAPE_30_MTF['Type']=='Total')][grade]
    _sel12 = VAPE_12_MTF[(VAPE_12_MTF['Type']=='Total')][grade]
    if len(_sel30) != 1 or len(_sel12) != 1:
        gradeslist_vape_bullets.append([])
        continue
    surverystat1 = _sel30.item()
    surverystat2 = _sel12.item()
    gradelist =[]
    for stat, timeline in zip([surverystat1, surverystat2], ['past-30-days', 'past-12-months']):
        _mtfsel = MTF[(MTF['variable']==_vape_mtf_var) & (MTF['Type']=='Total') & (MTF['timeline']==timeline)][grade]
        MTFvar = pd.to_numeric(_mtfsel.iloc[0], errors='coerce') if len(_mtfsel) == 1 else float('nan')
        if pd.isna(MTFvar):
            continue

        # B9: small-sample guard
        _vv = norm_verdict(stat, MTFvar, grade_size(grade))
        if _vv == 'onpar':
            gradeoutputstring = f'Vaping nicotine by {School} {grade} graders is on par with U.S. norms for {timeline} usage.'
            vapebullets.append(gradeoutputstring)
            gradelist.append('on par')
        elif _vv == 'below':
            gradeoutputstring = f'Vaping nicotine by {School} {grade} graders is below U.S. norms for {timeline} usage.'
            vapebullets.append(gradeoutputstring)
            gradelist.append('below')
        elif _vv == 'sigabove':
            gradeoutputstring = f'Vaping nicotine by {School} {grade} graders is well above U.S. norms for {timeline} usage.'
            vapebullets.append(gradeoutputstring)
            gradelist.append('significantly exceeds')
        # No upper-bound condition: 'significantly exceeds' already caught
        # values > norm+7.5%; the old second condition made this unreachable.
        elif _vv == 'above':
            gradeoutputstring = f'Vaping nicotine by {School} {grade} graders is above U.S. norms for {timeline} usage.'
            vapebullets.append(gradeoutputstring)
            gradelist.append('exceeds')
    gradeslist_vape_bullets.append(gradelist)


# In[ ]:





# In[66]:


# Verdict word for the vaping comparison intro. With an 8th grade this
# combines the 8th and 10th verdicts (template says "8th and 10th graders");
# without one, only the 10th verdicts apply — the intro's "8th and" is
# removed post-render, and combining [0]+[1] here would wrongly mix in 12th.
vapebullet8th10th = ''
_vape_bullet_grades = [g for g in ['8th', '10th', '12th'] if g in grades]
if '8th' in _vape_bullet_grades and len(gradeslist_vape_bullets) >= 2:
    combined = gradeslist_vape_bullets[0] + gradeslist_vape_bullets[1]
    _vape_majority = 2   # more than half of the 4 grade-timeline verdicts
elif len(gradeslist_vape_bullets) >= 1:
    combined = gradeslist_vape_bullets[0]
    _vape_majority = 1   # more than half of the 2 timeline verdicts
else:
    combined = []
    _vape_majority = 0
if combined.count('below') > _vape_majority:
    vapebullet8th10th = 'below'
elif combined.count('on par') > _vape_majority:
    vapebullet8th10th = 'on par with'
elif combined.count('exceeds') + combined.count('significantly exceeds') > _vape_majority:
    vapebullet8th10th = 'exceeds'
elif combined:
    # No majority verdict — avoid rendering "is  U.S. norms" mid-sentence.
    vapebullet8th10th = 'mixed relative to'


# In[67]:


tenthcountofabove = 0
for index, row in ALLSUBSTANCE10th_12_MTF[ALLSUBSTANCE10th_12_MTF['MTF']!=''].iterrows():
    if pd.to_numeric(row['school'],errors= 'coerce') <pd.to_numeric(row['MTF'], errors='coerce'):
        tenthcountofabove+=1
        
tenthcountofabove12 = 0
for index, row in ALLSUBSTANCE12th_12_MTF[ALLSUBSTANCE12th_12_MTF['MTF']!=''].iterrows():
    if pd.to_numeric(row['school'],errors= 'coerce') <pd.to_numeric(row['MTF'], errors='coerce'):
        tenthcountofabove12+=1
        
tenthcountofabove8_30 = 0
if len(ALLSUBSTANCE8th_30_MTF) >0:
    for index, row in ALLSUBSTANCE8th_30_MTF[ALLSUBSTANCE8th_30_MTF['MTF']!=''].iterrows():
        if pd.to_numeric(row['school'],errors= 'coerce') <pd.to_numeric(row['MTF'], errors='coerce'):
            tenthcountofabove8_30+=1
        
        
tenthcountofabove_30 = 0
for index, row in ALLSUBSTANCE10th_30_MTF[ALLSUBSTANCE10th_30_MTF['MTF']!=''].iterrows():
    if pd.to_numeric(row['school'],errors= 'coerce') <pd.to_numeric(row['MTF'], errors='coerce'):
        tenthcountofabove_30+=1
        
        
tenthcountofabove12_30 = 0
for index, row in ALLSUBSTANCE12th_30_MTF[ALLSUBSTANCE12th_30_MTF['MTF']!=''].iterrows():
    if pd.to_numeric(row['school'],errors= 'coerce') <pd.to_numeric(row['MTF'], errors='coerce'):
        tenthcountofabove12_30+=1


# In[ ]:





# In[68]:


# B2: '% of available measures' only counts substances that HAVE a U.S. norm
_denom10_12 = len(ALLSUBSTANCE10th_12_MTF[pd.to_numeric(ALLSUBSTANCE10th_12_MTF['MTF'], errors='coerce').notna()])
tenthbelowmeasure = round((tenthcountofabove / _denom10_12) * 100, 1) if _denom10_12 > 0 else 0
_denom12_12 = len(ALLSUBSTANCE12th_12_MTF[pd.to_numeric(ALLSUBSTANCE12th_12_MTF['MTF'], errors='coerce').notna()])
twelvebelowmeasure = round((tenthcountofabove12 / _denom12_12) * 100, 1) if _denom12_12 > 0 else 0

if len(ALLSUBSTANCE8th_30_MTF)>0:
    _denom8_30 = len(ALLSUBSTANCE8th_30_MTF[pd.to_numeric(ALLSUBSTANCE8th_30_MTF['MTF'], errors='coerce').notna()])
    eighthbelowmeasure30 = round((tenthcountofabove8_30 / _denom8_30) * 100, 1) if _denom8_30 > 0 else ''
else:
    eighthbelowmeasure30 = ''

_denom10_30 = len(ALLSUBSTANCE10th_30_MTF[pd.to_numeric(ALLSUBSTANCE10th_30_MTF['MTF'], errors='coerce').notna()])
tenthbelowmeasure_30 = round((tenthcountofabove_30 / _denom10_30) * 100, 1) if _denom10_30 > 0 else 0
_denom12_30 = len(ALLSUBSTANCE12th_30_MTF[pd.to_numeric(ALLSUBSTANCE12th_30_MTF['MTF'], errors='coerce').notna()])
twelvebelowmeasure_30 = round((tenthcountofabove12_30 / _denom12_30) * 100, 1) if _denom12_30 > 0 else 0

_tot10 = totalsbygrade[totalsbygrade['grade_level'].str.contains('10th')]['totalingrade']
one10thgraderis = round((1 / _tot10.item()) * 100, 1) if len(_tot10) > 0 and _tot10.item() != 0 else 0
_tot12 = totalsbygrade[totalsbygrade['grade_level'].str.contains('12th')]['totalingrade']
one12thgraderis = round((1 / _tot12.item()) * 100, 1) if len(_tot12) > 0 and _tot12.item() != 0 else 0

# 8th-grade equivalents, used post-render to replace the template's hardcoded
# "(One 8th grader equals 2.7%)" margins and the 12-month "all available
# measures" claim (the template has no placeholders for these).
_tot8 = totalsbygrade[totalsbygrade['grade_level'].str.contains('8th')]['totalingrade']
one8thgraderis = round((1 / _tot8.item()) * 100, 1) if len(_tot8) > 0 and _tot8.item() != 0 else ''
eighthbelowmeasure12 = ''
if '8th' in clean['grade_level'].values and len(ALLSUBSTANCE8th_12_MTF) > 0:
    _mtf8_12 = ALLSUBSTANCE8th_12_MTF[pd.to_numeric(ALLSUBSTANCE8th_12_MTF['MTF'], errors='coerce').notna()]
    _below8_12 = sum(1 for _idx, _row in _mtf8_12.iterrows()
                     if pd.to_numeric(_row['school'], errors='coerce') < pd.to_numeric(_row['MTF'], errors='coerce'))
    if len(_mtf8_12) > 0:
        eighthbelowmeasure12 = round((_below8_12 / len(_mtf8_12)) * 100, 1)





#pot_how_often_your_grade_grade #perceived pot
# The template reads these tables at fixed positions row[1]..row[10]
# (Perceived 8th, Actual 8th, ..., Perceived 12th, Actual 12th), so they must
# ALWAYS be built at full width. When the school has no 8th grade the 8th
# columns are filled with '-' — building a narrower table here would shift
# every grade's data under the wrong header in the report.
_has_8th_data = '8th' in clean['grade_level'].values
_perc_cols = ['Frequency', '8th (%)', '9th (%)', '10th (%)', '11th (%)', '12th (%)']
_perc_names = ['Frequency', 'Perceived 8th', 'Perceived 9th', 'Perceived 10th', 'Perceived 11th', 'Perceived 12th']
_actual_names = ['Frequency', 'Actual 8th', 'Actual 9th', 'Actual 10th', 'Actual 11th', 'Actual 12th']
_ordered_cols = ['Frequency',
                 'Perceived 8th','Actual 8th',
                 'Perceived 9th', 'Actual 9th','Perceived 10th', 'Actual 10th',
                 'Perceived 11th', 'Actual 11th',
                 'Perceived 12th' ,'Actual 12th']

def _full_width_percs(df, names):
    """Select the five grade (%) columns, padding a missing/empty 8th with '-'."""
    out = df.copy()
    if '8th (%)' not in out.columns:
        out['8th (%)'] = '-'
    out = out[_perc_cols].copy()
    out.columns = names
    if not _has_8th_data:
        out[names[1]] = '-'
    return out

perceived_only_percsdf = _full_width_percs(freq_answers_grade[0], _perc_names)

# The "Actual" columns must match the template footnote: "'Actual norm'
# figures incorporate 'never' responses to Q5 with all responses from Q7."
# how_often_1_drink (Q7) is only asked of students who have ever had a whole
# drink, so its raw percentages are among drinkers only. Recompute over ALL
# students in each grade, adding the "I have never had a whole drink" (Q5)
# students to the 'Never' row — the cannabis version of this table already
# works on this all-students basis.
_actual_freq_labels = ['Never',
 '1 to 2 times per year',
 'Once a month',
 'Twice a month',
 'Once a week',
 'Twice a week',
 'Daily']
howoften_justpercs_df = pd.DataFrame({'Frequency': _actual_freq_labels})
for _g in ['8th', '9th', '10th', '11th', '12th']:
    _ncol = f'{_g} (N)'
    if _ncol not in often_one_grade.columns or _ncol not in age_at_first_grade.columns:
        howoften_justpercs_df[f'{_g} (%)'] = '-'
        continue
    _freq_counts = pd.to_numeric(often_one_grade[_ncol], errors='coerce').fillna(0).tolist()
    _never_drank = pd.to_numeric(
        age_at_first_grade.loc[age_at_first_grade['Age'] == 'I have never had a whole drink of alcohol', _ncol],
        errors='coerce').fillna(0).sum()
    _freq_counts[0] += _never_drank
    _tot = pd.to_numeric(age_at_first_grade[_ncol], errors='coerce').fillna(0).sum()
    if _tot > 0:
        howoften_justpercs_df[f'{_g} (%)'] = [f'{round(_c / _tot * 100, 1)}%' for _c in _freq_counts]
    else:
        howoften_justpercs_df[f'{_g} (%)'] = '-'
howoften_justpercs_df = _full_width_percs(howoften_justpercs_df, _actual_names)
perceved_actual_table = pd.merge(howoften_justpercs_df, perceived_only_percsdf, on='Frequency')
perceved_actual_table = perceved_actual_table[_ordered_cols]

perceived_only_percsdf_SENIORS = _full_width_percs(freq_answers_grade[1], _perc_names)
perceived_SENIOR_versus_senioractual = pd.merge(perceived_only_percsdf_SENIORS, howoften_justpercs_df, on='Frequency')

perceived_only_percs_POTdf = _full_width_percs(pot_how_often_your_grade_grade, _perc_names)
how_often_pot_grade_perc = _full_width_percs(how_often_pot_grade, _actual_names)
perceved_actual_table_POT = pd.merge(how_often_pot_grade_perc, perceived_only_percs_POTdf, on='Frequency')
perceved_actual_table_POT = perceved_actual_table_POT[_ordered_cols]



# In[ ]:





# In[ ]:





# In[ ]:





# In[69]:


allyes_parent = []
for Q in AgreeDisagree3_answers:
    df = (total_agree_disagree_nograde(Q))
    value = df[df['index']=='Total Agree']['Count (N)']
    allyes_parent.append(value)
    
parent_involve_perc = float((sum(allyes_parent)/ len(AgreeDisagree3_answers)).iloc[0])

if parent_involve_perc>= 75:
    parentstring= 'high levels'
if (parent_involve_perc<75) & (parent_involve_perc>=55):
    parentstring=  'moderately high levels'
if (parent_involve_perc>35) & (parent_involve_perc<55):
    parentstring= 'moderate levels'
if (parent_involve_perc<35) :
    parentstring= 'low levels'


# In[ ]:





# In[70]:


# #######################################
# ### CREATE REPORT INSTANCE SO CAN ADD IMAGES TO IT 
# #########################################
doc = DocxTemplate(docxpath)




# In[71]:




sns.set(font="Georgia")


# In[72]:


# #######################################
# ### ADD ALL THE VIZES
# #########################################
import seaborn as sns
import matplotlib.pyplot as plt
from docxtpl import InlineImage
from docx.shared import Mm, Inches


bar_chart_it_up(alcohollifetumetable, os.path.join(chart_dir, 'alcol_melt_chart.png'))
alcol_melt_chart =  InlineImage(doc, os.path.join(chart_dir, 'alcol_melt_chart.png'), width=Inches(CHART_SIZE_MTF[0]))

bar_chart_it_up(alcoholl12monthtstable, os.path.join(chart_dir, 'alcohol12monthschart.png'))
alcohol12monthschart =  InlineImage(doc, os.path.join(chart_dir, 'alcohol12monthschart.png'), width=Inches(CHART_SIZE_MTF[0]))

bar_chart_it_up(alcoholl30daysstable, os.path.join(chart_dir, 'alcohol30monthschart.png'))
alcohol30monthschart =  InlineImage(doc, os.path.join(chart_dir, 'alcohol30monthschart.png'), width=Inches(CHART_SIZE_MTF[0]))

bar_chart_it_up(pot_lifetime_MTF, os.path.join(chart_dir, 'lifetimepotchart.png'))
lifetimepotchart =  InlineImage(doc, os.path.join(chart_dir, 'lifetimepotchart.png'), width=Inches(CHART_SIZE_MTF[0]))

bar_chart_it_up(pot_12_MTF, os.path.join(chart_dir, 'annualpotchart.png'))
annualpotchart =  InlineImage(doc, os.path.join(chart_dir, 'annualpotchart.png'), width=Inches(CHART_SIZE_MTF[0]))

bar_chart_it_up(pot_30_MTF, os.path.join(chart_dir, 'thirtypotchart.png'))
thirtypotchart =  InlineImage(doc, os.path.join(chart_dir, 'thirtypotchart.png'), width=Inches(CHART_SIZE_MTF[0]))

bar_chart_it_up(VAPE_12_MTF, os.path.join(chart_dir, 'annualvapechart.png'))
annualvapechart =  InlineImage(doc, os.path.join(chart_dir, 'annualvapechart.png'), width=Inches(CHART_SIZE_MTF[0]))

bar_chart_it_up(VAPE_30_MTF, os.path.join(chart_dir, 'thirtyvapechart.png'))
thirtyvapechart =  InlineImage(doc, os.path.join(chart_dir, 'thirtyvapechart.png'), width=Inches(CHART_SIZE_MTF[0]))

bar_chart_it_up_2(socialatmo, os.path.join(chart_dir, 'socialatmochart.png'))
socialatmochart =  InlineImage(doc, os.path.join(chart_dir, 'socialatmochart.png'), width=Inches(CHART_SIZE_PAIR[0]))


bar_chart_it_up_2(pressured_table, os.path.join(chart_dir, 'pressure_chart.png'))
pressure_chart =  InlineImage(doc, os.path.join(chart_dir, 'pressure_chart.png'), width=Inches(CHART_SIZE_PAIR[0]))

bar_chart_it_up_2(own_decision_table, os.path.join(chart_dir, 'owddesc_chart.png'))
owddesc_chart =  InlineImage(doc, os.path.join(chart_dir, 'owddesc_chart.png'), width=Inches(CHART_SIZE_PAIR[0]))

bar_chart_it_up_2(respect_table, os.path.join(chart_dir, 'respect_chart.png'))
respect_chart =  InlineImage(doc, os.path.join(chart_dir, 'respect_chart.png'), width=Inches(CHART_SIZE_PAIR[0]))


bar_chart_it_up_2(abstain_table, os.path.join(chart_dir, 'abstain_chart.png'))
abstain_chart =  InlineImage(doc, os.path.join(chart_dir, 'abstain_chart.png'), width=Inches(CHART_SIZE_PAIR[0]))

bar_chart_it_up_2(easytomake_table, os.path.join(chart_dir, 'easytomake_table_chart.png'))
easytomake_table_chart =  InlineImage(doc, os.path.join(chart_dir, 'easytomake_table_chart.png'), width=Inches(CHART_SIZE_PAIR[0]))


bar_chart_it_up_2(assume_table, os.path.join(chart_dir, 'assume_tablechart.png'))
assume_tablechart =  InlineImage(doc, os.path.join(chart_dir, 'assume_tablechart.png'), width=Inches(CHART_SIZE_PAIR[0]))

bar_chart_it_up_2(light_table, os.path.join(chart_dir, 'light_tablechart.png'))
light_tablechart =  InlineImage(doc, os.path.join(chart_dir, 'light_tablechart.png'), width=Inches(CHART_SIZE_PAIR[0]))

bar_chart_it_up_2(heavy_table, os.path.join(chart_dir, 'heavy_table_chart.png'))
heavy_table_chart =  InlineImage(doc, os.path.join(chart_dir, 'heavy_table_chart.png'), width=Inches(CHART_SIZE_PAIR[0]))   

bar_chart_it_up_2(valued_table, os.path.join(chart_dir, 'valued_table_chart.png'))
valued_table_chart =  InlineImage(doc, os.path.join(chart_dir, 'valued_table_chart.png'), width=Inches(CHART_SIZE_PAIR[0])) 

bar_chart_it_up_2(aca_table, os.path.join(chart_dir, 'aca_table_chart.png'))
aca_table_chart  =  InlineImage(doc, os.path.join(chart_dir, 'aca_table_chart.png'), width=Inches(CHART_SIZE_PAIR[0]))                              

bar_chart_it_up_2(staff_table, os.path.join(chart_dir, 'staff_tablechart.png'))
staff_tablechart =  InlineImage(doc, os.path.join(chart_dir, 'staff_tablechart.png'), width=Inches(CHART_SIZE_PAIR[0]))

bar_chart_it_up_2(sincere_table, os.path.join(chart_dir, 'sincere_tablechart.png'))
sincere_tablechart =  InlineImage(doc, os.path.join(chart_dir, 'sincere_tablechart.png'), width=Inches(CHART_SIZE_PAIR[0]))

bar_chart_it_up_2(care_table, os.path.join(chart_dir, 'care_tablechart.png'))
care_tablechart =  InlineImage(doc, os.path.join(chart_dir, 'care_tablechart.png'), width=Inches(CHART_SIZE_PAIR[0]))

bar_chart_it_up_2(aca2_table, os.path.join(chart_dir, 'aca2_table_chart.png'))
aca2_table_chart =  InlineImage(doc, os.path.join(chart_dir, 'aca2_table_chart.png'), width=Inches(CHART_SIZE_PAIR[0]))

school12_attitudes_chart =bar_chart_it_up_3(school12_grade, os.path.join(chart_dir, 'school12_attitudes_chart.png'))
school12_attitudes_chart =  InlineImage(doc, os.path.join(chart_dir, 'school12_attitudes_chart.png'), width=Inches(CHART_SIZE_PAIR[0]))

### K10 charts removed (not on this survey)

# Gambling and loot box charts
if _has_gambling and not gambling_all.empty:
    _gamb_chart_df = gambling_all.copy()
    _gamb_chart_df['pct'] = _gamb_chart_df['Percentage (%)'].str.replace('%', '').astype(float)
    fig, ax = plt.subplots(figsize=(7, 4))
    ax.barh(_gamb_chart_df['Response'], _gamb_chart_df['pct'], color='#1A63CD')
    ax.set_xlabel('Percentage (%)', fontname='Georgia')
    ax.set_title('Gambling in the Past 30 Days', fontname='Georgia', fontsize=13, fontweight='bold')
    ax.invert_yaxis()
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.tick_params(left=False)
    plt.gcf().patch.set_facecolor('white')
    plt.savefig(os.path.join(chart_dir, 'gambling_chart.png'), bbox_inches='tight', dpi=300, facecolor='white')
    plt.close()
    gambling_chart = InlineImage(doc, os.path.join(chart_dir, 'gambling_chart.png'), width=Mm(150))
else:
    gambling_chart = ''

if _has_lootbox and not lootbox_all.empty:
    _lb_chart_df = lootbox_all.copy()
    _lb_chart_df['pct'] = _lb_chart_df['Percentage (%)'].str.replace('%', '').astype(float)
    fig, ax = plt.subplots(figsize=(7, 4))
    ax.barh(_lb_chart_df['Response'], _lb_chart_df['pct'], color='#267EF2')
    ax.set_xlabel('Percentage (%)', fontname='Georgia')
    ax.set_title('Loot Box / Randomized Purchases in the Past 30 Days', fontname='Georgia', fontsize=13, fontweight='bold')
    ax.invert_yaxis()
    for spine in ax.spines.values():
        spine.set_visible(False)
    ax.tick_params(left=False)
    plt.gcf().patch.set_facecolor('white')
    plt.savefig(os.path.join(chart_dir, 'lootbox_chart.png'), bbox_inches='tight', dpi=300, facecolor='white')
    plt.close()
    lootbox_chart = InlineImage(doc, os.path.join(chart_dir, 'lootbox_chart.png'), width=Mm(150))
else:
    lootbox_chart = ''

if _has_gambling and not gambling_drinking_cross.empty:
    fig, ax = plt.subplots(figsize=(6, 3.5))
    bars = ax.bar(gambling_drinking_cross['Group'],
                  gambling_drinking_cross['Also gambled (%)'].astype(str).str.replace('%', '').astype(float),
                  color=['#1A63CD', '#C9D2DA'], width=0.5)
    ax.set_ylabel('Percentage (%)', fontname='Georgia')
    ax.set_title('Gambling Rate by Drinking Status', fontname='Georgia', fontsize=13, fontweight='bold')
    for bar in bars:
        h = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2., h + 0.5, f'{h}%', ha='center', fontsize=10, fontname='Georgia')
    for spine in ['top', 'right']:
        ax.spines[spine].set_visible(False)
    plt.gcf().patch.set_facecolor('white')
    plt.savefig(os.path.join(chart_dir, 'gambling_drinking_chart.png'), bbox_inches='tight', dpi=300, facecolor='white')
    plt.close()
    gambling_drinking_chart = InlineImage(doc, os.path.join(chart_dir, 'gambling_drinking_chart.png'), width=Mm(150))
else:
    gambling_drinking_chart = ''

if _has_gambling and _has_lootbox and not gambling_lootbox_cross.empty:
    fig, ax = plt.subplots(figsize=(6, 3.5))
    bars = ax.bar(gambling_lootbox_cross['Group'],
                  gambling_lootbox_cross['Also used loot boxes (%)'].astype(str).str.replace('%', '').astype(float),
                  color=['#1A63CD', '#C9D2DA'], width=0.5)
    ax.set_ylabel('Percentage (%)', fontname='Georgia')
    ax.set_title('Loot Box Use by Gambling Status', fontname='Georgia', fontsize=13, fontweight='bold')
    for bar in bars:
        h = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2., h + 0.5, f'{h}%', ha='center', fontsize=10, fontname='Georgia')
    for spine in ['top', 'right']:
        ax.spines[spine].set_visible(False)
    plt.gcf().patch.set_facecolor('white')
    plt.savefig(os.path.join(chart_dir, 'gambling_lootbox_chart.png'), bbox_inches='tight', dpi=300, facecolor='white')
    plt.close()
    gambling_lootbox_chart = InlineImage(doc, os.path.join(chart_dir, 'gambling_lootbox_chart.png'), width=Mm(150))
else:
    gambling_lootbox_chart = ''


# In[73]:


def bar_chart_it_up_5(df, save, palette):
    palette=palette

    
    
    df_melt =df[['Response','High-risk (%)',
       'Low-risk (%)', 'Non-drinker (%)', 
       'Cannabis use in last 30 days','Never used cannabis', 
       'Using substances to cope']].melt(id_vars ='Response', value_vars=['High-risk (%)',
       'Low-risk (%)', 'Non-drinker (%)', 
       'Cannabis use in last 30 days','Never used cannabis', 
       'Using substances to cope'])


    ax = sns.catplot(
        x="variable",       # x variable name
        y="value",       # y variable name
        hue="Response",  # group variable name
        data=df_melt,     # dataframe to plot
        kind="bar",
        palette = palette, aspect=13/10)

    ax.set_xticklabels(rotation=90)
    ax.set(xlabel='Group', ylabel='Percentage (%)')

    #plt.show()
    
    plt.gcf().patch.set_facecolor('white')
    plt.savefig(save, bbox_inches='tight', dpi=300, facecolor='white')
    plt.close()


# In[74]:




# In[75]:


####################
## Appendix C
###################

# --- 8th-grade scalars referenced by the March-2026 template (the script
# previously computed only the old 7th-grade variants, so these template
# slots rendered empty). Definitions ported from the executive script. ---
lifetime8th = ever_by_age(age_at_first_grade, '8th (%)')
twelve_8th = ever_by_age_yes(onewhole_grade,'8th (%)')

if '8th (%)' in daysatleast1_grade.columns:
    _days30_8th = daysatleast1_grade[daysatleast1_grade['Response']!='I did not drink alcohol during the past 30 days']['8th (%)']
else:
    _days30_8th = pd.Series(dtype=object)
if _days30_8th.eq('').all() or _days30_8th.empty:
    thirty_8th = 'N/A'
else:
    thirty_8th = round((_days30_8th.str.split('%').str[0]).astype(float).sum(), 1)

social_atmo_8th_yes = total_agree_disagree_perc(AgreeDisagree_answers_grade[0], 'yes')
if '8th (%)' in social_atmo_8th_yes.columns:
    social_atmo_8th_yes_perc = social_atmo_8th_yes[social_atmo_8th_yes['index']=='Total Agree']['8th (%)'].item()
else:
    social_atmo_8th_yes_perc = ''

if '8th (N)' in social_atmo_bygrade.columns:
    eighthsocial_atmocountagree = social_atmo_bygrade[(social_atmo_bygrade['Response']=='Agree') | (social_atmo_bygrade['Response']=='Strongly agree')][['8th (N)']].sum().sum()
else:
    eighthsocial_atmocountagree = 0
_denom_sa = totalsbygrade[totalsbygrade['grade_level']=='8th']['totalingrade'].sum()
if _denom_sa > 0:
    eighthsocial_atmocountagree_perc = round((eighthsocial_atmocountagree / _denom_sa)*100, 1)
else:
    # Fall back to the percentage already extracted from the grade table
    _fallback = social_atmo_8th_yes_perc
    try:
        eighthsocial_atmocountagree_perc = round(float(str(_fallback).replace('%','')), 1)
    except (ValueError, TypeError):
        eighthsocial_atmocountagree_perc = 0.0

pressured_8th_yes = total_agree_disagree_perc(AgreeDisagree_answers_grade[1], 'yes')
if '8th (%)' in pressured_8th_yes.columns:
    pressured_8th_yes_perc = pressured_8th_yes[pressured_8th_yes['index']=='Total Agree']['8th (%)'].item()
else:
    pressured_8th_yes_perc = ''

if '8th (N)' in pressured_bygrade.columns:
    eighthpressured_bygradecountagree = pressured_bygrade[(pressured_bygrade['Response']=='Agree') | (pressured_bygrade['Response']=='Strongly agree')][['8th (N)']].sum().sum()
else:
    eighthpressured_bygradecountagree = 0
_denom_pr = totalsbygrade[totalsbygrade['grade_level']=='8th']['totalingrade'].sum()
if _denom_pr > 0:
    eighthpressured_bygradecountagree_perc = round((eighthpressured_bygradecountagree / _denom_pr)*100, 1)
else:
    # Fall back to the percentage already extracted from the grade table
    _fallback_pr = pressured_8th_yes_perc
    try:
        eighthpressured_bygradecountagree_perc = round(float(str(_fallback_pr).replace('%','')), 1)
    except (ValueError, TypeError):
        eighthpressured_bygradecountagree_perc = 0.0

if '8th (%)' in AgreeDisagree_answers_grade[8].columns:
    _binge_8th = AgreeDisagree_answers_grade[8][2:]['8th (%)']
else:
    _binge_8th = pd.Series(dtype=object)
if _binge_8th.eq('').all() or _binge_8th.empty:
    totalagree_bingedrinking_8th = 'N/A'
else:
    totalagree_bingedrinking_8th = round(_binge_8th.str.split('%').str[0].astype(float).sum(), 1)

# When the school has no 8th grade, the 8th-grade substance tables built above
# are empty DataFrames, which renders as tables with headers but no rows.
# Rebuild them here (after all numeric/bullet logic has run on the empty
# versions) as display tables showing the U.S. MTF norms with '-' in the
# school columns, so the report shows the norms alongside a dash.
if '8th' not in clean['grade_level'].values:
    _school_subs = pd.DataFrame({'substance': substances_DF['substance'].unique()})

    def _mtf_8th_display(timeline):
        _m = fix_MTF_naming(MTF[(MTF['timeline'] == timeline) & (MTF['Type'] == 'Total')][['8th', 'variable']])
        _t = pd.merge(_school_subs, _m, left_on='substance', right_on='variable')
        _t = _t[['substance', '8th']].drop_duplicates()
        _t.columns = ['substance', 'MTF']
        _t['school'] = '-'
        return _t.sort_values('MTF', ascending=False, key=lambda _s: pd.to_numeric(_s, errors='coerce'))

    ALLSUBSTANCE8th_12_MTF = _mtf_8th_display('past-12-months')
    ALLSUBSTANCE8th_30_MTF = _mtf_8th_display('past-30-days')

    _fem8 = fix_MTF_naming(MTF[(MTF['timeline'] == 'past-12-months') & (MTF['Type'] == 'Female')][['8th', 'variable']])
    _fem8.columns = ['8th MTF Females', 'substance']
    _masc8 = fix_MTF_naming(MTF[(MTF['timeline'] == 'past-12-months') & (MTF['Type'] == 'Male')][['8th', 'variable']])
    _masc8.columns = ['8th MTF Males', 'substance']
    sub12_by_sex_MTFcombo8th = pd.merge(_masc8, _fem8, on='substance')
    sub12_by_sex_MTFcombo8th = pd.merge(_school_subs, sub12_by_sex_MTFcombo8th, on='substance')
    sub12_by_sex_MTFcombo8th['8th Survey Males'] = '-'
    sub12_by_sex_MTFcombo8th['8th Survey Females'] = '-'
    sub12_by_sex_MTFcombo8th = sub12_by_sex_MTFcombo8th[['substance', '8th MTF Males', '8th Survey Males', '8th MTF Females', '8th Survey Females']]
    sub12_by_sex_MTFcombo8th = sub12_by_sex_MTFcombo8th.drop_duplicates().sort_values('8th MTF Males', ascending=False, key=lambda _s: pd.to_numeric(_s, errors='coerce'))

##############################################################################
### QA FIXES: narrative values recomputed from the canonical flags and counts
### (B1 drinker counts, B2 denominators, B4 rounding, B5 wording from data,
### B6 label mapping). Everything below is passed to the template through
### QA_CONTEXT (merged into the render context) or overrides an older value.
##############################################################################
QA_CONTEXT = {}
_grades_seq = grades_in_data_sorted

# ---------- Survey administration (B7) ----------
QA_CONTEXT['N_analyzed'] = N_analyzed

# ---------- Exec Summary: alcohol ----------
_n30 = int(clean['drank_30d'].sum()); _n12 = int(clean['drank_12m'].sum())
noalcoholpaste30 = N_analyzed - _n30
noalcoholpaste30perc = pct(noalcoholpaste30, N_analyzed, 'no alcohol in past 30 days', 'all analyzed students')

_lifetime_by_grade = [(g, flag_rate('drank_lifetime', g, name=f'lifetime alcohol {g}')) for g in _grades_seq]
_twelve_by_grade = [(g, flag_rate('drank_12m', g, name=f'past-12-month alcohol {g}')) for g in _grades_seq]
# "In students who report any lifetime use of alcohol, past-30-day use ..." -> base = lifetime drinkers
_thirty_by_grade = [(g, flag_rate('drank_30d', g, among='drank_lifetime', name=f'past-30-day alcohol among lifetime drinkers {g}'))
                    for g in _grades_seq]
QA_CONTEXT['lifetime_trend'] = trend_phrase(_lifetime_by_grade)
QA_CONTEXT['twelve_trend'] = trend_phrase(_twelve_by_grade)
QA_CONTEXT['thirty_trend'] = trend_phrase(_thirty_by_grade)
_lt = dict(_lifetime_by_grade); _tw = dict(_twelve_by_grade); _th = dict(_thirty_by_grade)
lifetime8th, lifetime12th = _lt.get('8th', ''), _lt.get('12th', '')
twelve_8th, twelve_12th = _tw.get('8th', ''), _tw.get('12th', '')
thirty_8th, thirty_12th = _th.get('8th', ''), _th.get('12th', '')

_d30 = clean[clean['drank_30d']]
twormoredays = pct(int((_d30['one_drink_past_30_days'] != '1 day').sum()), len(_d30),
                   'drank on 2+ of the past 30 days', 'past-30-day drinkers')
_binge = clean['days_5_or_more_drinks'].notna() & (clean['days_5_or_more_drinks'] != _NO_BINGE)
five_ormorecount = int(_binge.sum())
five_ormoreperc = pct(five_ormorecount, int((clean['drank_30d'] | _binge).sum()), '5+ drinks in a row (past-30-day drinkers)', 'past-30-day drinkers')
fve_ormore_two_ormore_days_count = int((_binge & (clean['days_5_or_more_drinks'] != '1 day')).sum())
fve_ormore_two_ormore_days_perc = pct(fve_ormore_two_ormore_days_count, five_ormorecount, '5+ in a row on 2+ days', 'students with 5+ in a row')
morethanintended = pct(int((clean['drank_12m'] & clean['drank_more'].notna() & _not_never(clean['drank_more'])).sum()), _n12,
                       'drank more than intended', 'past-12-month drinkers')
QA_CONTEXT['howmanyhighrisk5typ_perc'] = howmanyhighrisk5typ_perc

# ---------- Key Data: alcohol highlights ----------
past30_yes_drankcount = _n30
_body = clean['grade_level'].isin(BODY_GROUP_GRADES)
past30_yes_drankcount101112 = int((clean['drank_30d'] & _body).sum())
howmanyhighrisk5_count = int((clean['drinker_risk_10_12'] == 'High-risk').sum())
howmanyhighrisk5_perc = pct(howmanyhighrisk5_count, int(((clean['drank_30d'] | _binge) & _body).sum()),
                            f'higher-risk among {BODY_GROUP_LABEL} past-30-day drinkers', f'{BODY_GROUP_LABEL} past-30-day drinkers')
percent_yes_last30_drank = howmanyhighrisk5_perc
_where = clean.loc[clean['drank_30d'] & clean['where_do_you_drink'].notna(), 'where_do_you_drink']
percdrank30days_didntdrinkathome = pct(int((_where != 'At home').sum()), len(_where),
                                       'most drinking outside own home', 'past-30-day drinkers who answered where')
percdrank30days_didntdrinkathome_string = f"{share_word(percdrank30days_didntdrinkathome)} ({percdrank30days_didntdrinkathome}%)"
drank30and12_perc = pct(_n30, _n12, 'drank in past 30 days (of past-12-month drinkers)', 'past-12-month drinkers')
drank30and12_perc_12 = pct(int((clean['drank_30d'] & clean['grade_level'].eq('12th')).sum()),
                           int((clean['drank_12m'] & clean['grade_level'].eq('12th')).sum()), '12th: past 30 days of past-12-month drinkers')
_six_plus = _d30['one_drink_past_30_days'].isin(['6 to 9 days', '10 or more days'])
dranklast30_6ormoresum = int(_six_plus.sum())
dranklast30_6ormore = pct(dranklast30_6ormoresum, len(_d30), 'drank on 6+ days', 'past-30-day drinkers')

# Footnote [3] example, computed for this school (12th grade)
_d30_12 = _d30[_d30['grade_level'] == '12th']
_fn3_n = int(((_d30_12['drinks_in_a_row_30_days']) == '5 or more drinks').sum())
QA_CONTEXT['fn3_perc'] = pct(_fn3_n, len(_d30_12)) if len(_d30_12) else '-'
QA_CONTEXT['fn3_count'] = _fn3_n
QA_CONTEXT['fn3_body_perc'] = pct(_fn3_n, grade_size('12th')) if grade_size('12th') else '-'

# ---------- Circumstances of use (B2: respondents only, no blank row) ----------
def _where_share(p):
    """Finer bands than share_word: 'A minority (0.9%)' reads oddly."""
    try:
        p = float(p)
    except (TypeError, ValueError):
        return 'Some'
    if p < 5:
        return 'Very few'
    if p < 25:
        return 'A small percentage'
    return share_word(p).capitalize()
_where_counts = _where.value_counts()
wheredrinkingstrings = []
for _loc, _cnt in _where_counts.items():
    _p = pct(_cnt, len(_where))
    wheredrinkingstrings.append(f"{_where_share(_p)} ({_p}%) do most of their drinking {str(_loc)[0].lower() + str(_loc)[1:]}.")
wheredrinkingstrings = [w.replace('drinking other.', 'drinking somewhere else.') for w in wheredrinkingstrings]
wheredrinkingstrings += [''] * max(0, 5 - len(wheredrinkingstrings))

# "Location of alcohol use varies by grade level ..." built from the data
_club_by_grade = []
for g in _grades_seq:
    _wg = clean.loc[clean['drank_30d'] & clean['grade_level'].eq(g) & clean['where_do_you_drink'].notna(), 'where_do_you_drink']
    if len(_wg):
        _club_by_grade.append((g, pct(int(_wg.eq('At a club or bar').sum()), len(_wg)), int(_wg.eq('At a club or bar').sum()), len(_wg)))
if _club_by_grade:
    _hi = max(_club_by_grade, key=lambda t: t[1]); _lo = min(_club_by_grade, key=lambda t: t[1])
    if _hi[1] == _lo[1]:
        QA_CONTEXT['bar_club_sentence'] = (f"Of students who have had alcohol within the past 30 days, {_hi[1]}% reported doing most of their "
                                           f"drinking at a bar or club, with no difference between grades.")
    else:
        QA_CONTEXT['bar_club_sentence'] = (f"Of students who have had alcohol within the past 30 days, {_hi[1]}% of {_hi[0]} graders "
                                           f"({_hi[2]} of {_hi[3]} students) reported doing most of their drinking at a bar or club, compared with "
                                           f"{_lo[1]}% of {_lo[0]} graders ({_lo[2]} of {_lo[3]} students).")
else:
    QA_CONTEXT['bar_club_sentence'] = ''
_hf_by_grade = []
for g in _grades_seq:
    _wg = clean.loc[clean['drank_30d'] & clean['grade_level'].eq(g) & clean['where_do_you_drink'].notna(), 'where_do_you_drink']
    if len(_wg):
        _hf_by_grade.append((g, pct(int(_wg.isin(['At home', "At a friend's house"]).sum()), len(_wg))))
if _hf_by_grade and all(v > 50 for _, v in _hf_by_grade):
    QA_CONTEXT['home_friend_sentence'] = ("Still, the majority of alcohol use at every grade level occurs at a student’s own home "
                                          "or at a friend’s home.")
elif _hf_by_grade:
    _maj = [g for g, v in _hf_by_grade if v > 50]
    QA_CONTEXT['home_friend_sentence'] = (f"Most alcohol use occurs at a student’s own home or at a friend’s home in "
                                          f"{len(_maj)} of {len(_hf_by_grade)} grades ({range_phrase(_hf_by_grade)}).")
else:
    QA_CONTEXT['home_friend_sentence'] = ''

# ---------- Cannabis ----------
clean['never_pot'] = ~clean['pot_lifetime']
percent_never_pot_allgrades = pd.Series([flag_rate('never_pot', name='never used cannabis, all students')])
neverpot_all = fmt_pct(percent_never_pot_allgrades.item())

# Normative cannabis (B6): cumulative "1-2 times a year OR MORE" and
# "once a month or more", actual vs perceived on matching answer ranges.
_pot_any = ['1 to 2 times per year', 'Once a month', 'Twice a month', 'Once a week', 'Twice a week', 'Daily']
_pot_month = ['Once a month', 'Twice a month', 'Once a week', 'Twice a week', 'Daily']
_hop = how_often_pot_grade; _pp = pot_how_often_your_grade_grade
for g in ['8th', '9th', '10th', '11th', '12th', '7th']:
    QA_CONTEXT[f'pot_yearplus_actual_{g}'] = fmt_pct(count_pct_from_table(_hop, _hop['Frequency'].isin(_pot_any), f'{g} (%)'))
    QA_CONTEXT[f'pot_yearplus_perceived_{g}'] = fmt_pct(count_pct_from_table(_pp, _pp['Frequency'].isin(_pot_any), f'{g} (%)'))
    QA_CONTEXT[f'pot_monthplus_actual_{g}'] = fmt_pct(count_pct_from_table(_hop, _hop['Frequency'].isin(_pot_month), f'{g} (%)'))
    QA_CONTEXT[f'pot_monthplus_perceived_{g}'] = fmt_pct(count_pct_from_table(_pp, _pp['Frequency'].isin(_pot_month), f'{g} (%)'))
twelthpot_moreathan12times = count_pct_from_table(_hop, _hop['Frequency'].isin(_pot_any), '12th (%)')
twelthpot_more12year_perceived = count_pct_from_table(_pp, _pp['Frequency'].isin(_pot_any), '12th (%)')
twelthpot_morethanonceamonth = count_pct_from_table(_hop, _hop['Frequency'].isin(_pot_month), '12th (%)')
twelthpot_morethanonceamonth_perceived = count_pct_from_table(_pp, _pp['Frequency'].isin(_pot_month), '12th (%)')

# ---------- U.S. norm comparison intros (B5/B7: built from the data) ----------
def _strip_measures(b):
    b = (b or '').strip()
    return b[:-len(' measures.')] + ' measures' if b.endswith(' measures.') else b.rstrip('.')

def _by_grade_flag(flag):
    return [(g, flag_rate(flag, g)) for g in _grades_seq]

def _age_sentence(series, noun):
    tp = trend_phrase(series)
    return f"{noun} {tp}." if tp else ''

_cmp_grades = [g for g in ['8th', '10th', '12th'] if g in _grades_seq]
def _compare_intro(noun, bullets):
    parts = [f"use by {g} graders {_strip_measures(b)}" for g, b in zip(['8th', '10th', '12th'], bullets)
             if g in _cmp_grades and b]
    if not parts:
        return ''
    s_ = '; '.join(parts)
    return f"Compared with U.S. norms, {noun} {s_}."
QA_CONTEXT['alcohol_compare_intro'] = _compare_intro('alcohol', [albullet1 if '8th' in _cmp_grades else '', albullet2, albullet3])
QA_CONTEXT['cannabis_compare_intro'] = _compare_intro('cannabis', [maribullet1, maribullet2, maribullet3])
QA_CONTEXT['alcohol_age_sentence'] = _age_sentence(_lifetime_by_grade, 'At {School}, lifetime alcohol use'.replace('{School}', str(School)))
QA_CONTEXT['cannabis_age_sentence'] = _age_sentence(_by_grade_flag('pot_lifetime'), 'Lifetime cannabis use')
QA_CONTEXT['vape_age_sentence'] = _age_sentence(_by_grade_flag('vape_12m'), 'Past-12-month vaping')
_vape_verdicts = {}
for g in _cmp_grades:
    _vals = []
    for flag, tl in [('vape_12m', 'past-12-months'), ('vape_30d', 'past-30-days')]:
        _m = MTF[(MTF['variable'] == _vape_mtf_var) & (MTF['Type'] == 'Total') & (MTF['timeline'] == tl)][g]
        if len(_m) == 1:
            _vals.append(norm_verdict(flag_rate(flag, g), pd.to_numeric(_m.iloc[0], errors='coerce'), grade_size(g)))
    _vals = [v for v in _vals if v]
    if _vals:
        _vape_verdicts[g] = 'below' if all(v == 'below' for v in _vals) else (
            'on par with' if all(v == 'onpar' for v in _vals) else (
            'above' if all(v in ('above', 'sigabove') for v in _vals) else 'mixed relative to'))
QA_CONTEXT['vape_compare_intro'] = ("Compared with U.S. norms, vaping nicotine by " +
    '; by '.join(f"{g} graders is {v} U.S. norms" for g, v in _vape_verdicts.items()) + '.') if _vape_verdicts else ''

def _verdict_lifetime(g):
    _m = MTF[(MTF['variable'] == 'Alcohol') & (MTF['Type'] == 'Total') & (MTF['timeline'] == 'lifetime')]
    if g not in _m.columns or len(_m) != 1:
        return None
    return norm_verdict(_lt.get(g), pd.to_numeric(_m[g].iloc[0], errors='coerce'), grade_size(g))
_vwords = {'below': 'below', 'onpar': 'on par with', 'above': 'above', 'sigabove': 'well above'}
_exec_parts = [f"{_vwords[v]} U.S. norms for {g} graders" for g in _cmp_grades for v in [_verdict_lifetime(g)] if v]
QA_CONTEXT['exec_alcohol_trend_sentence'] = (
    f"Lifetime alcohol use {QA_CONTEXT['lifetime_trend']}" +
    (f", and is {_join_and(_exec_parts)}." if _exec_parts else '.')) if QA_CONTEXT['lifetime_trend'] else ''

# Substance-comparison highlights (B7: replaces the "Do these manually" note)
def _substance_highlight(df, g, period):
    _m = df[pd.to_numeric(df['MTF'], errors='coerce').notna()] if len(df) else df
    if g not in _grades_seq or len(_m) == 0:
        return ''
    _v = [norm_verdict(pd.to_numeric(r['school'], errors='coerce'), pd.to_numeric(r['MTF'], errors='coerce'), grade_size(g))
          for _, r in _m.iterrows()]
    _below = sum(v == 'below' for v in _v); _par = sum(v == 'onpar' for v in _v)
    _above = [r['substance'] for (_, r), v in zip(_m.iterrows(), _v) if v in ('above', 'sigabove')]
    # B5/B7: plain-language summary — no "measures"/"on par" jargon,
    # no zero-count clauses, "all N" when one verdict covers everything.
    n = len(_v)
    window = 'the past 12 months' if '12' in period else 'the past 30 days'
    lead = f"Compared with U.S. averages, {School} {g} graders report "
    if _below == n:
        return lead + f"lower use of all {n} substances in {window}."
    if _par == n:
        return lead + f"about the same use of all {n} substances in {window}; none are higher."
    if len(_above) == n:
        return lead + f"higher use of all {n} substances in {window} ({_join_and(_above)})."
    clauses = []
    if _below:
        clauses.append(f"lower use of {_below} of the {n} substances in {window}")
    if _par:
        _par_label = f"the other {_par}" if _below and _below + _par == n else f"{_par}"
        clauses.append(f"about the same use of {_par_label}" if _below
                       else f"about the same use of {_par} of the {n} substances in {window}")
    if not clauses:
        return lead + (f"higher use of {_join_and(_above)} in {window}." if _above
                       else f"use close to U.S. averages across the {n} substances in {window}.")
    txt = lead + ' and '.join(clauses)
    txt += (f". Use of {_join_and(_above)} is higher than the U.S. average." if _above
            else ". None are higher than U.S. averages.")
    return txt
QA_CONTEXT['substance12_highlights'] = [_substance_highlight(d, g, 'past-12-month') for d, g in
    [(ALLSUBSTANCE8th_12_MTF, '8th'), (ALLSUBSTANCE10th_12_MTF, '10th'), (ALLSUBSTANCE12th_12_MTF, '12th')]]
QA_CONTEXT['substance30_highlights'] = [_substance_highlight(d, g, 'past-30-day') for d, g in
    [(ALLSUBSTANCE8th_30_MTF, '8th'), (ALLSUBSTANCE10th_30_MTF, '10th'), (ALLSUBSTANCE12th_30_MTF, '12th')]]
QA_CONTEXT['eighthbelowmeasure12'] = eighthbelowmeasure12
QA_CONTEXT['one8thgraderis'] = one8thgraderis

# ---------- School climate / culture of use (B4, B5) ----------
def _agree_by_grade(df):
    return [(g, count_pct_from_table(df, df['Response'].isin(['Agree', 'Strongly agree']), f'{g} (%)')) for g in _grades_seq]
_sa = _agree_by_grade(AgreeDisagree_answers_grade[0]); _pr = _agree_by_grade(AgreeDisagree_answers_grade[1])
_sa_d = dict(_sa); _pr_d = dict(_pr)
social_atmo_8th_yes_perc = _sa_d.get('8th', ''); social_atmo_12th_yes_perc = _sa_d.get('12th', '')
pressured_8th_yes_perc = _pr_d.get('8th', ''); pressured_12th_yes_perc = _pr_d.get('12th', '')
QA_CONTEXT['social_atmo_range'] = range_phrase(_sa)
QA_CONTEXT['pressured_range'] = range_phrase(_pr)
QA_CONTEXT['lowest_grade'] = lowest_grade
QA_CONTEXT['social_atmo_low_perc'] = _sa_d.get(lowest_grade, '')
QA_CONTEXT['pressured_low_perc'] = _pr_d.get(lowest_grade, '')
QA_CONTEXT['culture_highlight_text'] = (
    f"The share of students who believe that “the social atmosphere at this school encourages alcohol use” "
    f"{trend_phrase(_sa)}. Agreement that students feel pressured to drink {trend_phrase(_pr)}.")
QA_CONTEXT['culture_trend_paragraph'] = (
    f"Agreement that the social atmosphere at {School} encourages alcohol use {trend_phrase(_sa)}. "
    f"Agreement that {School} students feel pressured to use alcohol {trend_phrase(_pr)}.")
disagree_social_atmo = round(100 - float(social_atmo_df_yes_perc), 1)
QA_CONTEXT['disagree_social_atmo_word'] = share_word(disagree_social_atmo)
QA_CONTEXT['disagree_pressured'] = round(100 - float(pressured_df_yes_perc), 1)
QA_CONTEXT['disagree_pressured_word'] = share_word(QA_CONTEXT['disagree_pressured'])
_od = _agree_by_grade(AgreeDisagree_answers_grade[4])
_od_all = total_agree_disagree_nograde(AgreeDisagree_answers[4])
_od_all = round(float(_od_all[_od_all['index'] == 'Total Agree']['Count (N)'].item()), 1)
_fit_all = round(float(fitin_no_perc), 1)
QA_CONTEXT['tolerance_sentence'] = (
    f"Responses to the following questions show attitudes towards abstaining from alcohol. Overall, {_od_all}% of students agree that "
    f"{School} students let you make your own decisions about drinking alcohol ({range_phrase(_od)}), and {_fit_all}% disagree "
    f"that students who abstain from alcohol do not fit in.")
# Respect-for-choices paragraph (B5): each claim checked against its number
def _agree_all(df):
    _c = pd.to_numeric(df['Count (N)'], errors='coerce').fillna(0)
    return pct(_c[df['Response'].isin(['Agree', 'Strongly agree'])].sum(), _c.sum())
_easy = _agree_all(AgreeDisagree_answers[2]); _own = _agree_all(AgreeDisagree_answers[4])
_resp = _agree_all(AgreeDisagree_answers[5]); _fit_dis = round(100 - _agree_all(AgreeDisagree_answers[6]), 1)
QA_CONTEXT['respect_paragraph'] = (
    f"Attitudes toward individual choices about alcohol at {School} are mixed. {share_word(_easy).capitalize()} of students ({_easy}%) "
    f"believe that “it is easy to make friends without drinking,” {share_word(_own)} ({_own}%) agree that “students will let you make your own "
    f"decisions about drinking,” {share_word(_resp)} ({_resp}%) feel their peers respect students who do not drink alcohol, and "
    f"{share_word(_fit_dis)} ({_fit_dis}%) disagree that “students at this school who abstain from alcohol do not fit in.”"
    + (" These positive sentiments suggest that the school climate would respond well to efforts to encourage and support non-use."
       if min(_easy, _own, _resp, _fit_dis) > 50 else
       " Strengthening support for students who choose not to drink is an area where prevention efforts could focus."))
_binge_series = [(g, count_pct_from_table(AgreeDisagree_answers_grade[8], AgreeDisagree_answers_grade[8]['Response'].isin(['Agree', 'Strongly agree']), f'{g} (%)')) for g in _grades_seq]
QA_CONTEXT['binge_trend_sentence'] = (f"Agreement that “students at this school think heavy or binge drinking is okay” {trend_phrase(_binge_series)}."
                                      if trend_phrase(_binge_series) else '')

# Academic achievement sentences (B5/B9): from the grades-by-group table
def _grade_share(label):
    _r = all_risk_options_grades[all_risk_options_grades['Grades'] == label]
    if not len(_r):
        return {}
    _r = _r.iloc[0]
    return {k: float(str(_r[c]).replace('%', '')) for k, c in
            [('higher-risk', 'High-risk (%)'), ('lower-risk', 'Low-risk (%)'), ('non', 'Non-drinker (%)'),
             ('cannabis', 'Yearly or More Frequent Marijuana Use')]}
_top_label = listofageoptions[-1] if listofageoptions == _number_scale else listofageoptions[0]   # Mostly 7's / Mostly A's
_mid_label = listofageoptions[3] if listofageoptions == _number_scale else listofageoptions[2]   # Mostly 4's / Mostly C's
def _academic_sentence(label, more_or_less):
    sh = _grade_share(label)
    if not sh:
        return ''
    _lbl = label.replace("'", "’")
    if more_or_less == 'less':
        ok = sh['higher-risk'] < min(sh['non'], sh['lower-risk']) and sh['cannabis'] < min(sh['non'], sh['lower-risk'])
        return (f'"Higher-risk" drinkers ({sh["higher-risk"]}%) and cannabis users ({sh["cannabis"]}%) were less likely than “non-drinkers” '
                f'({sh["non"]}%) and "lower-risk" drinkers ({sh["lower-risk"]}%) to have received “{_lbl}” during the last school year.'
                if ok else
                f'The share of students who received “{_lbl}” during the last school year was {sh["higher-risk"]}% for "higher-risk" drinkers, '
                f'{sh["cannabis"]}% for cannabis users, {sh["lower-risk"]}% for "lower-risk" drinkers and {sh["non"]}% for “non-drinkers.”')
    ok = sh['higher-risk'] > sh['non'] and sh['cannabis'] > sh['non']
    return (f'"Higher-risk" drinkers ({sh["higher-risk"]}%) and cannabis users ({sh["cannabis"]}%) were more likely than “non-drinkers” '
            f'({sh["non"]}%) to have received “{_lbl}” during the last school year.'
            if ok else
            f'The share who received “{_lbl}” was {sh["higher-risk"]}% for "higher-risk" drinkers, {sh["cannabis"]}% for cannabis users '
            f'and {sh["non"]}% for “non-drinkers.”')
QA_CONTEXT['academic_sentence_top'] = _academic_sentence(_top_label, 'less')
QA_CONTEXT['academic_sentence_mid'] = _academic_sentence(_mid_label, 'more')
_n_cannabis_group = int(clean['yearlypot'].notna().sum())
QA_CONTEXT['academic_caveat'] = (f" Note that the cannabis-user group has only {_n_cannabis_group} students, so small differences should not be over-interpreted."
                                 if _n_cannabis_group < MIN_CELL_N else '')

_att = attitude_grade
_att_mask = _att['Response'].isin(sort_order6[:2])
QA_CONTEXT['attitude_young_perc'] = count_pct_from_table(_att, _att_mask, f'{lowest_grade} (%)') if lowest_grade else ''
QA_CONTEXT['attitude_12_perc'] = count_pct_from_table(_att, _att_mask, '12th (%)')
_bg = AgreeDisagree_answers_grade[8]
QA_CONTEXT['binge_low_perc'] = count_pct_from_table(_bg, _bg['Response'].isin(['Agree', 'Strongly agree']), f'{lowest_grade} (%)') if lowest_grade else ''
totalagree_bingedrinking_12th = count_pct_from_table(_bg, _bg['Response'].isin(['Agree', 'Strongly agree']), '12th (%)')
totalagree_bingedrinking_8th = count_pct_from_table(_bg, _bg['Response'].isin(['Agree', 'Strongly agree']), '8th (%)')

# ---------- Higher-risk section: subgroup clustering (B5/B9) ----------
_clust_yes = 0; _clust_tot = 0
for _q in [q for q in YNQs if q.endswith('_12') and q != 'use_alcohol_12']:
    _hr = group_share(_q, lambda ser: ser.eq('Yes'))
    _others = [group_share(_q, lambda ser: ser.eq('Yes'), group=gname) for gname in ['Low-risk', 'Non-drinker']]
    _hr_n = int(((clean['drinker_risk_10_12'] == 'High-risk')).sum())
    if _hr == '' or any(o == '' for o in _others):
        continue
    _clust_tot += 1
    _gap = 100.0 / _hr_n * MIN_STUDENT_GAP if _hr_n else 100
    if all(_hr - o > _gap for o in _others):
        _clust_yes += 1
QA_CONTEXT['clustering_sentence'] = (
    f"For {_clust_yes} of {_clust_tot} substances, a larger share of “higher-risk” drinkers than of “lower-risk” drinkers "
    f"and “non-drinkers” used the substance in the past 12 months (by more than one student’s worth of difference). For example:")

# ---------- Consequences (B5) ----------
_cs = highlowrisk_consequence_sub
_hr_more = int((_cs['High-risk (%)'] > _cs['Low-risk (%)']).sum()); _lr_more = int((_cs['Low-risk (%)'] > _cs['High-risk (%)']).sum())
QA_CONTEXT['consequence_comparison_sentence'] = (
    f"\"Higher-risk\" drinkers reported a higher rate than “lower-risk” drinkers for {_hr_more} of {len(_cs)} consequences"
    + (f", and a lower rate for {_lr_more}." if _lr_more else '.'))

# ---------- Consequences to others / attitudes / policies: body groups (B3) ----------
QA_CONTEXT['tookcare_perc_1012'] = tookcare_perc_1012
QA_CONTEXT['vom_perc_1012'] = vom_perc_1012
attitude_by_risk1012 = by_grader_count_and_percenter('attitude', 'drinker_risk_10_12', drinkinglist, sort_order6, 'Response')
_a = attitude_by_risk1012
for _grp, _key in [('Non-drinker', 'non'), ('High-risk', 'high')]:
    QA_CONTEXT[f'att_never_good_{_key}'] = count_pct_from_table(_a, _a['Response'].isin(sort_order6[:2]), f'{_grp} (%)')
    QA_CONTEXT[f'att_ok_noninterfere_{_key}'] = fmt_pct(count_pct_from_table(_a, _a['Response'].eq(sort_order6[2]), f'{_grp} (%)'))
    QA_CONTEXT[f'att_ok_any_{_key}'] = count_pct_from_table(_a, _a['Response'].isin(sort_order6[3:]), f'{_grp} (%)')
for _k in ['pol_highrisktoostrict', 'pol_norisktoostrict', 'pol_highrisktoolenient', 'pol_norisktoolenient']:
    QA_CONTEXT[_k] = globals()[_k]
_pol = too_lenient_answers[0]; _enf = too_lenient_answers[1]
def _about_right(df):
    _c = pd.to_numeric(df['Count (N)'], errors='coerce').fillna(0)
    return pct(_c[df['Response'] == 'About right'].sum(), _c.sum())
_p_ar, _e_ar = _about_right(_pol), _about_right(_enf)
QA_CONTEXT['policies_about_right_sentence'] = (
    f"{_p_ar}% of students find the school’s current policies “about right,” and {_e_ar}% say the same of enforcement.")
QA_CONTEXT['BODY_GROUP_NOTE'] = BODY_GROUP_NOTE
QA_CONTEXT['BODY_GROUP_LABEL'] = BODY_GROUP_LABEL
QA_CONTEXT['HIGHRISK_DEFINITION'] = HIGHRISK_DEFINITION
QA_CONTEXT['LOWRISK_DEFINITION'] = LOWRISK_DEFINITION
QA_CONTEXT['NONDRINKER_DEFINITION'] = NONDRINKER_DEFINITION
QA_CONTEXT['APPENDIX_C_GROUP_NOTE'] = APPENDIX_C_GROUP_NOTE

# ---------- Normative alcohol gap sentence (B5) ----------
_perc_never = [(g, count_pct_from_table(drink_how_often_your_grade_grade, drink_how_often_your_grade_grade['Frequency'].eq('Never'), f'{g} (%)'))
               for g in _grades_seq]
_perc_any = [(g, (round(100 - v, 1) if v != '' else '')) for g, v in _perc_never]
def _delta(series):
    pts = [(g, v) for g, v in series if v != '']
    return (pts[0], pts[-1], round(pts[-1][1] - pts[0][1], 1)) if len(pts) >= 2 else None
_da = _delta(_lifetime_by_grade); _dp = _delta(_perc_any)
if _da and _dp:
    QA_CONTEXT['normative_alcohol_gap_sentence'] = (
        f"Note that the share of students who report ever having a whole drink {trend_phrase(_lifetime_by_grade)} "
        f"({'+' if _da[2] >= 0 else ''}{_da[2]} points), while the share who assume their classmates drink at least sometimes "
        f"{trend_phrase(_perc_any)} ({'+' if _dp[2] >= 0 else ''}{_dp[2]} points). Even in the youngest grade, students "
        f"assume far more drinking than actually occurs.")
else:
    QA_CONTEXT['normative_alcohol_gap_sentence'] = ''

# ---------- small-sample report (B9) ----------
for _grp in ['High-risk', 'Low-risk', 'Non-drinker']:
    _n = int((clean['drinker_risk_10_12'] == _grp).sum())
    if 0 < _n < MIN_CELL_N:
        logger.warning(f"B9 SMALL GROUP: {_grp} ({BODY_GROUP_LABEL}) has only {_n} students")
for _grp in ['Cannabis use in last 30 days', 'Never used cannabis']:
    _n = int((clean['pot_risk30_never'] == _grp).sum())
    if 0 < _n < MIN_CELL_N:
        logger.warning(f"B9 SMALL GROUP: '{_grp}' ({BODY_GROUP_LABEL}) has only {_n} students; treat its column with caution")
_yp = int(clean['yearlypot'].notna().sum())
if 0 < _yp < MIN_CELL_N:
    logger.warning(f"B9 SMALL GROUP: yearly-or-more cannabis users ({BODY_GROUP_LABEL}) has only {_yp} students")


##############################################################################
### NOV 2026 FIXES: narrative rules C4, C5, C10, C12-C15, C17, C19, C25
### and the methodology text for C7/C8. Everything is passed to the template
### through QA_CONTEXT, which overrides older values of the same name.
##############################################################################
def _q(key, lower_first=True):
    """C19: quoted item wording from the survey file itself."""
    t = QUESTION_TEXT.get(key, '').strip().rstrip('.').strip()
    return (t[0].lower() + t[1:]) if (t and lower_first) else t

# C24: template sections that only make sense with 8th graders
QA_CONTEXT['has_8th'] = '8th' in _grades_seq

# ---------- C25: survey date and report date, both labelled ----------
QA_CONTEXT['report_date'] = REPORT_DATE
QA_CONTEXT['date_today'] = REPORT_DATE
QA_CONTEXT['date_survey'] = date_survey

# ---------- C10: grade lists only include grades that have respondents ----------
_cmp_present = [g for g in ['7th', '8th', '10th', '12th'] if g in _grades_seq and g != '7th']
QA_CONTEXT['cmp_grades_text'] = _join_and(_cmp_present)
QA_CONTEXT['cmp_grades_title'] = f"Comparison between U.S. Norms and {School} {_join_and(_cmp_present)} Graders"
_gnums = [g[:-2] for g in _grades_seq]
QA_CONTEXT['prettygrades'] = (f'{", ".join(_gnums[:-1])} and {_gnums[-1]}' if len(_gnums) > 1 else
                              (_gnums[0] if _gnums else prettygrades))

# ---------- C4: Culture of Use overall agree, ONE computation, answered base ----------
def _agree_share(df_grade, glist, name):
    _ncols = [f'{g} (N)' for g in glist if f'{g} (N)' in df_grade.columns]
    _c = df_grade[_ncols].apply(pd.to_numeric, errors='coerce').fillna(0)
    _agree = _c[df_grade['Response'].isin(['Agree', 'Strongly agree']).values].sum().sum()
    return pct(_agree, _c.sum().sum(), name, f'students in {", ".join(glist)} who answered')
_upper = [g for g in ['9th', '10th', '11th', '12th'] if g in _grades_seq]
_lower = [g for g in ['7th', '8th'] if g in _grades_seq]
# C10: "7th-8th graders" / "9th-12th graders" are built from the grades present
QA_CONTEXT['has_lower_grades'] = bool(_lower)
QA_CONTEXT['lower_grades_label'] = grade_range_label(_lower) if _lower else ''
QA_CONTEXT['upper_grades_label'] = grade_range_label(_upper) if _upper else ''
ninto12thsocial_atmocountagree_perc = _agree_share(AgreeDisagree_answers_grade[0], _upper, 'culture: social atmosphere agree, 9th-12th')
ninto12thpressured_bygradecountagree_perc = _agree_share(AgreeDisagree_answers_grade[1], _upper, 'culture: pressured agree, 9th-12th')
QA_CONTEXT['ninto12thsocial_atmocountagree_perc'] = ninto12thsocial_atmocountagree_perc
QA_CONTEXT['ninto12thpressured_bygradecountagree_perc'] = ninto12thpressured_bygradecountagree_perc
if _lower:
    QA_CONTEXT['seventhand8thsocial_atmocountagree_perc'] = _agree_share(AgreeDisagree_answers_grade[0], _lower, 'culture: social atmosphere agree, 7th-8th')
    QA_CONTEXT['seventhand8thpressured_bygradecountagree_perc'] = _agree_share(AgreeDisagree_answers_grade[1], _lower, 'culture: pressured agree, 7th-8th')
QA_CONTEXT['social_atmo_df_yes_perc'] = _agree_share(AgreeDisagree_answers_grade[0], _grades_seq, 'culture: social atmosphere agree, all grades')
QA_CONTEXT['pressured_df_yes_perc'] = _agree_share(AgreeDisagree_answers_grade[1], _grades_seq, 'culture: pressured agree, all grades')

# ---------- C5: "use less than the U.S. average on X% of substances" ----------
# One classifier (norm_verdict, one-student margin) feeds BOTH the highlight
# box and this sentence. Only substances with a U.S. norm count.
def _below_share(df, g, period):
    if g not in _grades_seq or not len(df):
        return '-'
    _m = df[pd.to_numeric(df['MTF'], errors='coerce').notna()]
    _v = [norm_verdict(pd.to_numeric(r['school'], errors='coerce'), pd.to_numeric(r['MTF'], errors='coerce'), grade_size(g))
          for _, r in _m.iterrows()]
    if not _v:
        return '-'
    _b = sum(v == 'below' for v in _v)
    _p = pct(_b, len(_v), f'{g} {period}: substances below U.S. norms', f'{g} substances with a U.S. norm')
    return f"{_p._num()}% ({_b} of {len(_v)})"
for _key, _df, _g, _per in [('eighthbelowmeasure12', ALLSUBSTANCE8th_12_MTF, '8th', '12-month'),
                            ('tenthbelowmeasure', ALLSUBSTANCE10th_12_MTF, '10th', '12-month'),
                            ('twelvebelowmeasure', ALLSUBSTANCE12th_12_MTF, '12th', '12-month'),
                            ('eighthbelowmeasure30', ALLSUBSTANCE8th_30_MTF, '8th', '30-day'),
                            ('tenthbelowmeasure_30', ALLSUBSTANCE10th_30_MTF, '10th', '30-day'),
                            ('twelvebelowmeasure_30', ALLSUBSTANCE12th_30_MTF, '12th', '30-day')]:
    QA_CONTEXT[_key] = _below_share(_df, _g, _per)

# ---------- C12: "risk factors cluster" examples only where the gap is real ----------
_SUB_WORDS = {'cigarettes': 'cigarettes', 'vapes': 'electronic vaporizers', 'smokeless': 'smokeless tobacco',
              'pot': 'cannabis', 'ecstasy': 'ecstasy', 'heroin': 'heroin', 'narcotics': 'narcotics not prescribed to them',
              'amphetamines': 'amphetamines not prescribed to them', 'sedatives': 'sedatives not prescribed to them',
              'methamphetamine': 'methamphetamine', 'cocaine': 'cocaine', 'lsd': 'LSD', 'hallucinogen': 'other hallucinogens',
              'inhalants': 'inhalants', 'steroids': 'steroids not prescribed to them', 'cold': 'over-the-counter drugs'}
_hr_n = int((clean['drinker_risk_10_12'] == 'High-risk').sum())
_one_hr = (100.0 / _hr_n * MIN_STUDENT_GAP) if _hr_n else 100.0
_cluster = []
for _qq in [q for q in YNQs if q.endswith('_12') and q != 'use_alcohol_12']:
    _yes = lambda ser: ser.eq('Yes')
    _h = group_share(_qq, _yes, group='High-risk')
    _l = group_share(_qq, _yes, group='Low-risk')
    _nd = group_share(_qq, _yes, group='Non-drinker')
    _all = pct(int((clean['grade_level'].isin(BODY_GROUP_GRADES) & clean[_qq].eq('Yes')).sum()),
               int((clean['grade_level'].isin(BODY_GROUP_GRADES) & clean[_qq].notna()).sum()))
    if '' in (_h, _l, _nd):
        continue
    _gap = float(_h) - max(float(_l), float(_nd))
    if _gap > _one_hr:     # more than one higher-risk student's worth of difference
        _cluster.append((_gap, _qq.split('_')[1], _h, _l, _nd, _all))
_cluster.sort(key=lambda t: -t[0])
QA_CONTEXT['cluster_examples'] = [
    f"{fmt_pct(h)} of “higher-risk” drinkers have used {_SUB_WORDS.get(s, s)} in the past 12 months, compared with "
    f"{fmt_pct(l)} of “lower-risk” drinkers, {fmt_pct(nd)} of “non-drinkers” and {fmt_pct(a)} of all {BODY_GROUP_LABEL} graders."
    for _, s, h, l, nd, a in _cluster[:4]]
_n_subs = len([q for q in YNQs if q.endswith('_12') and q != 'use_alcohol_12'])
if not _hr_n:
    QA_CONTEXT['clustering_sentence'] = "There are no “higher-risk” drinkers in these grades, so no comparison is made."
elif _cluster:
    QA_CONTEXT['clustering_sentence'] = (
        f"These findings show how risk factors can “cluster.” For {len(_cluster)} of {_n_subs} substances, a larger share "
        f"of “higher-risk” drinkers than of “lower-risk” drinkers and “non-drinkers” used the substance in the past 12 "
        f"months, by more than one student’s worth of difference. For example:")
else:
    QA_CONTEXT['clustering_sentence'] = (
        f"In the {School} results, no substance was used in the past 12 months by a clearly larger share of “higher-risk” "
        f"drinkers than of “lower-risk” drinkers and “non-drinkers” (a difference of more than one student).")

# ---------- C13: say how many answered when it differs from the group size ----------
_n_where = len(_where)
QA_CONTEXT['circumstances_lead'] = (
    f"Of the {_n_where} students who drank within the past 30 days and answered this question:" if _n_where != _n30 else
    f"Of the {_n30} students who have had alcohol within the past 30 days:")

# ---------- C14: Culture of Use paragraph, framing word from the data ----------
_easy = _agree_all(AgreeDisagree_answers[2]); _own = _agree_all(AgreeDisagree_answers[4])
_resp = _agree_all(AgreeDisagree_answers[5]); _fit_dis = round(100 - float(_agree_all(AgreeDisagree_answers[6])), 1)
_supportive = [float(v) for v in (_easy, _own, _resp, _fit_dis) if v != '']
if _supportive and min(_supportive) > 50:
    _frame = 'largely supportive'
elif _supportive and max(_supportive) < 50:
    _frame = 'largely unsupportive'
else:
    _frame = 'mixed'
_light_dis = pct(*[pd.to_numeric(AgreeDisagree_answers[7]['Count (N)'], errors='coerce').fillna(0)[
                       AgreeDisagree_answers[7]['Response'].isin(['Disagree', 'Strongly disagree']).values].sum(),
                   pd.to_numeric(AgreeDisagree_answers[7]['Count (N)'], errors='coerce').fillna(0).sum()],
                 'culture: disagree most students do not drink or are light drinkers', 'students who answered')
_respect_txt = (f"Overall, attitudes toward students who choose not to drink are {_frame}: {_easy}% of students agree that "
                f"“{_q('easy_to_make_friends')},” and {_resp}% agree that “{_q('respect')}.”")
if _light_dis != '' and float(_light_dis) > 0:
    _respect_txt += (f" At the same time, {_light_dis}% of students disagree that “{_q('light_drinkers')},” although "
                     f"{noalcoholpaste30perc}% of {School} students did not drink alcohol in the past 30 days.")
_respect_txt += (" Because most students already support the choice not to drink, efforts that encourage and support "
                 "non-use build on attitudes students already hold." if _frame.startswith('largely supportive') else
                 " Strengthening support for students who choose not to drink is an area where prevention efforts could focus.")
QA_CONTEXT['respect_paragraph'] = _respect_txt
QA_CONTEXT['tolerance_sentence'] = (
    f"Responses to the following questions show attitudes towards abstaining from alcohol. Overall, {_own}% of students agree that "
    f"“{_q('own_decisions')}” ({range_phrase(_od)}), and {_fit_dis}% disagree that “{_q('fit_in')}.”")

# ---------- C19: answer wording quoted from the survey, not paraphrased ----------
QA_CONTEXT['att_never_good_quote'] = (f"“{sort_order6[0].rstrip('.').lower()},” or that “{sort_order6[1].rstrip('.').lower()}.”")

# ---------- C15: grades closing sentence is conditional ----------
_n_cannabis_group = int(clean['yearlypot'].notna().sum())
_ng = {k: int((clean['drinker_risk_10_12'] == v).sum()) for k, v in
       [('higher-risk', 'High-risk'), ('lower-risk', 'Low-risk'), ('non', 'Non-drinker')]}
_ng['cannabis'] = _n_cannabis_group
def _gap_ok(a_key, b_key, sh):
    """More than one student's worth of difference in BOTH groups compared."""
    one = max(100.0 / _ng[a_key] if _ng[a_key] else 100.0, 100.0 / _ng[b_key] if _ng[b_key] else 100.0)
    return abs(sh[a_key] - sh[b_key]) > one * MIN_STUDENT_GAP
def _academic_sentence2(label, direction):
    sh = _grade_share(label)
    if not sh:
        return ''
    _lbl = label.replace("'", "’")
    if direction == 'less':
        ok = all(sh[a] < min(sh['non'], sh['lower-risk']) and _gap_ok(a, 'non', sh) and _gap_ok(a, 'lower-risk', sh)
                 for a in ('higher-risk', 'cannabis'))
        if ok:
            return (f'"Higher-risk" drinkers ({sh["higher-risk"]}%) and cannabis users ({sh["cannabis"]}%) were less likely than '
                    f'“non-drinkers” ({sh["non"]}%) and "lower-risk" drinkers ({sh["lower-risk"]}%) to have received “{_lbl}” '
                    f'during the last school year.')
    else:
        ok = all(sh[a] > sh['non'] and _gap_ok(a, 'non', sh) for a in ('higher-risk', 'cannabis'))
        if ok:
            return (f'"Higher-risk" drinkers ({sh["higher-risk"]}%) and cannabis users ({sh["cannabis"]}%) were more likely than '
                    f'“non-drinkers” ({sh["non"]}%) to have received “{_lbl}” during the last school year.')
        return ''   # no claim about the middle grade band unless the gap is real
    return (f'The share of students who received “{_lbl}” during the last school year was {sh["higher-risk"]}% of '
            f'"higher-risk" drinkers ({_ng["higher-risk"]} students), {sh["cannabis"]}% of cannabis users ({_ng["cannabis"]}), '
            f'{sh["lower-risk"]}% of "lower-risk" drinkers ({_ng["lower-risk"]}) and {sh["non"]}% of “non-drinkers” ({_ng["non"]}).')
QA_CONTEXT['academic_sentence_top'] = _academic_sentence2(_top_label, 'less')
QA_CONTEXT['academic_sentence_mid'] = _academic_sentence2(_mid_label, 'more')
QA_CONTEXT['academic_caveat'] = ''
_sh_top = _grade_share(_top_label)
if min(_ng['higher-risk'], _ng['cannabis']) < MIN_GROUP_FOR_GRADES_CLAIM or not _sh_top:
    QA_CONTEXT['grades_closing_sentence'] = (
        f"We cannot conclude from this data that a causal relationship exists between alcohol or cannabis use and grades. "
        f"With only {_ng['higher-risk']} “higher-risk” drinkers and {_ng['cannabis']} cannabis users, the {School} results "
        f"do not show a clear pattern in either direction.")
elif all(_sh_top[a] < _sh_top['non'] and _gap_ok(a, 'non', _sh_top) for a in ('higher-risk', 'cannabis')):
    QA_CONTEXT['grades_closing_sentence'] = (
        "We cannot conclude from this data that a causal relationship exists between alcohol or cannabis use and grades. "
        "The findings are, however, in keeping with other research that posits a negative association between "
        "drinking/drug use and academic performance.")
else:
    QA_CONTEXT['grades_closing_sentence'] = (
        f"We cannot conclude from this data that a causal relationship exists between alcohol or cannabis use and grades. "
        f"In the {School} results, “higher-risk” drinkers and cannabis users were not less likely than “non-drinkers” to "
        f"report the highest grades, so these data do not show the negative association reported in other research.")

# ---------- C17: footnote when the two "higher-risk" items disagree ----------
_m_greatest5 = int((clean['drinks_in_a_row_30_days'] == '5 or more drinks').sum())
_k_zero = int((clean['drank_30d'] & clean['drinks_in_a_row_30_days'].eq('0 drinks')).sum())
QA_CONTEXT['greatest_footnote'] = (
    (f"This table and the “higher-risk” count come from two separate questions, so they do not match exactly. "
     f"“Higher-risk” drinkers are the {totalhighrisk} students who reported having 5 or more drinks on 1 or more days in "
     f"the past 30 days. On this question, {_m_greatest5} students gave 5 or more drinks as their greatest number in a row, "
     f"and {_k_zero} students who reported drinking in the past 30 days answered “0 drinks.” Responses are shown as "
     f"students gave them.") if (_m_greatest5 != totalhighrisk or _k_zero) else '')

# ---------- C2/C3/C7/C8: methodology text printed in the report ----------
QA_CONTEXT['exclusion_summary'] = EXCLUSION_SUMMARY
QA_CONTEXT['time_rule_text'] = (
    "For every substance, a student who reported use in the past 30 days is also counted as having used it in the "
    "past 12 months, so past-30-day figures can never be higher than past-12-month figures.")
QA_CONTEXT['k10_method_text'] = ''   # no Kessler (K10) questions on this survey


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
_SUB_PLAIN = {'Electronic Vaporizers**': 'Electronic Vaporizers', 'Smokeless': 'Smokeless Tobacco',
              'Over the Counter Cold Medicine': 'Over-the-Counter Cold Medicine', 'Meth': 'Methamphetamine'}
def _sub_word(name):
    """Substance name for a sentence: lower case, LSD kept."""
    n = _SUB_PLAIN.get(name, name)
    return n if n == 'LSD' else n.lower()

def _count_clause(n_below, n_total, region):
    if n_below == 0:
        return f"not below the {region} average for any of the {n_total} substances"
    return f"below the {region} average for {n_below} of {n_total} substances"
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
    """{region: [(substance, verdict), ...]} for one grade and time window."""
    rows = CMP_SUB.get((grade, tl), [])
    out = {}
    for _reg in REGIONS:
        _ri = 1 if _reg == 'Asia' else 2          # rows are (substance, Asia, Europe, school)
        _v = []
        for _r in rows:
            _b, _c = _r[_ri], _r[3]
            if _b is None or _c is None or pd.isna(_b) or pd.isna(_c):
                continue
            _v.append((_r[0], norm_verdict(_c, _b, grade_size(grade))))
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
            _below = sum(v == 'below' for _, v in _v); _above = [_sub_word(s) for s, v in _v if v in ('above', 'sigabove')]
            _bits.append((f"lower use than the {_reg} average for {_below} of {len(_v)} substances" if _below else
                          f"no lower use than the {_reg} average for any of the {len(_v)} substances")
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
                _above = [_sub_word(s) for s, v in _vs if v in ('above', 'sigabove')]
                _bits.append(_count_clause(sum(v == 'below' for _, v in _vs), len(_vs), _reg)
                             + (f" and above it for {_join_and(_above)}" if _above else ''))
        out.append(f"Past-12-month substance use by {School} {grade} grade {_word} is {'; it is '.join(_bits)}." if _bits else '')
    return out
QA_CONTEXT['outputstringssexbullets'] = _sex_bullets('8th')
QA_CONTEXT['outputstringssexbullets10th'] = _sex_bullets('10th')
QA_CONTEXT['outputstringssexbullets12th'] = _sex_bullets('12th')
QA_CONTEXT['cmp_grades_title'] = f"Comparison between Asia and Europe International School Averages and {School} {_join_and(_CMP_GRADES)} Graders"
QA_CONTEXT['benchmark_note'] = BENCH_NOTE


ReportListofVariables =[lifetime8th,
twelve_8th,
thirty_8th,
social_atmo_8th_yes_perc,
pressured_8th_yes_perc,
eighthsocial_atmocountagree_perc,
eighthpressured_bygradecountagree_perc,
totalagree_bingedrinking_8th,
sexdf,
sex_by_grade,
sex_by_risk,
gradedf,
grade_by_risk,
grade_by_risk10_12,
highrisk_count_10_12,
lowrisk_count_10_12,
norisk_count_10_12,
avg_age_highrisk,
avg_age_lowrisk,
avg_age_norisk,
age_at_first_drinkdf,
age_at_first_drink_output,
age_at_first_grade,
age_at_first_by_risk,
totalhighrisk,
highriskbullet1,
highrisk_comparison,
totalstudents,
perc_high_riskall,
perc_high_riskallone_in,
grade_overalldf,
grades_overall_bygrade,
grades_overall_by_risk,
all_risk_options_grades,
lowriskdrinkers,
noniskdrinkers,
percent8thgraders_ever_drank,
percent10thgraders_ever_drank,
percent12thgraders_ever_drank,
percent_never_drank_allgrades,
percent_never_drank_11th,
percent_never_drank_12th,
lifetime7th,
lifetime12th,
percent8thgraders_Never_drank,
percent9thgraders_Never_drank,
percent10thgraders_Never_drank,
percent11thgraders_Never_drank,
percent12thgraders_Never_drank,
everdrank_grade_sex,
aggregates,
agedf,
where_firstdf,
where_first_grade,
where_first_by_risk,
totalsbygrade,
onestudentin1012outof100,
totalsbygrade_andsex,
first_drink_at_friends,
often_one_drink,
often_one_grade,
often_one_by_risk,
often_one_by_risk10_12,
perc_1012_highrisk_twiceormore,
drinklessthan12timesperyear,
howoften10th_onceamonthperc,
eleventhmorethantwice,
twlthtwiceaweek_perc,
howmany_df,
howmany_grade,
howmany_by_risk,
howmany_by_risk10_12,
high_risk_morethan5,
less_than2_peryear,
howmanyhighrisk_perc,
howmanyhighrisk5_perc,
one_drink12df,
onewhole_grade,
onewhole_by_risk,
twelve_7th,
twelve_12th,
total8th_last12_perc,
tota10th_last12_perc,
total12th_last12_perc,
total8th_last12_perc,
tota10th_last12_perc,
total12th_last12_perc,
alcohol_past_12monthbysex,
neverQ_answers,
neverQ_answers_grade,
neverQ_answers_by_risk,
highriskbulletsssss,
lowriskbulletsssss,
higherthan50perc_consequences,
mostrpopular_lowriskconsequence,
mostrpopular_lowriskconsequenceperc,
highrisk_twice_risk,
high_risk_drunk10_12,
high_risk_morethan3drunk,
high_risk_tocut10_12,
high_risk10_12_cutdown,
high_risk_guiltyt10_12,
high_risk10_12_guilty,
ride_drive,
later_sorry,
later_sorry_perc,
morethanintended,
one_drink_30df,
daysatleast1_grade,
daysatleast1_by_risk,
dranklast30_6ormore,
total8th_last30_perc,
tota10th_last30_perc,
total12th_last30_perc,
alcohol_past_30daysbysex,
noalcoholpaste30,
noalcoholpaste30perc,
thirty_7th,
thirty_12th,
twormoredays,
greatest30df,
greatest30_grade,
greatest30_by_risk,
fiveinrowdf,
days_5_or_more_grade,
days_5_or_more_by_risk,
five_ormoreperc,
five_ormorecount,
fve_ormore_two_ormore_days_count,
fve_ormore_two_ormore_days_perc,
drunk30df,
drunk_30_grade,
drunk_30_by_risk,
wheremostdf,
where_do_you_grade,
where_do_you_by_risk,
wheredrinkingstrings,
agefirstpot,
age_first_time_pot_grade,
age_first_time_pot_by_risk,
howoftenpot,
how_often_pot_grade,
how_often_pot_by_risk,
neverpot_all,
neverpot8th,
neverpot9th,
neverpot10th,
neverpot11th,
neverpot12th,
twelthpot_moreathan12times,
twelthpot_morethanonceamonth,
percent8thgraders_ever_pot,
percent10thgraders_ever_pot,
percent12thgraders_ever_pot,
percent_never_pot_allgrades,
percent_all_typically,
pot_lifetime,
YNQ_answers,
YNQ_answers_grade,
YNQ_answers_by_risk,
cigarettes12,
notcig12monthall,
cigarettes_12_risk10_12,
cigarettes_30_risk10_12,
pot_last_12_risksubset,
pot_risk_12_bullets,
pot_last_30_risksubset,
OTC_last_12_risksubset,
OTC_risk_12_bullets,
VAPE_last_12_risksubset,
VAPE_risk_12_bullets,
drank30and12,
drank12,
drank30and12_perc,
use_alcohol_30df,
past30_yes_drankcount,
use_alcohol_30_by_gradedf,
past30_yes_drankcount101112,
use_alcohol_30_by_riskdf,
highrisk_yes_last_drink_30,
percdrank30days_didntdrinkathome,
percdrank30days_didntdrinkathome_string,
percent_yes_last30_drank,
total8th_last12VAPE_perc,
tota10th_last12VAPE_perc,
total12th_last12VAPE_perc,
VAPEPAST12,
total8th_last30VAPE_perc,
tota10th_last30VAPE_perc,
total12th_last30VAPE_perc,
VAPEPAST30,
total8th_last12POT_perc,
tota10th_last12POT_perc,
total12th_last12POT_perc,
POTPAST12,
total8th_last30POT_perc,
tota10th_last30POT_perc,
total30th_last30POT_perc,
POTPAST30,
all_sub_risks_all,
all_sub_risks_all_output,
all_sub_risks_all_output30,
highriskampusage,
substances_DF,
ALLSUBSTANCE8th_12,
ALLSUBSTANCE10th_12,
ALLSUBSTANCE12th_12,
ALLSUBSTANCE8th_30,
ALLSUBSTANCE10th_30,
ALLSUBSTANCE12th_30,
substancesbysexandgrader,
sub12months_sex8th,
sub12months_sex10th,
sub12months_sex12th,
neverQ2_answers,
neverQ2_answers_grade,
neverQ2_answers_by_risk,
tookcare,
vom,
dranktoomuch,
vom2,
tookcare_perc,
vom_perc,
dranktoomuch_perc1012,
vom_10_12_perc,
concerned_other,
concerned_other_perc,
schooldance,
school_dance_use_grade,
school_dance_use_by_risk,
danceperc,
attitude,
attitude_grade,
attitude_by_risk,
neveragoodthingtodo_perc,
neveragoodthingtodo_perc_high,
nondrink_okay_attitude,
highrisk_okay_attitude,
ApproveDisapprove_answers,
ApproveDisapprove_answers_grade,
ApproveDisapprove_answers_by_risk,
RiskQ_answers,
RiskQ_answers_grade,
RiskQ_answers_by_risk,
risk_harm_1_2,
risk_harm_daily_1_2perc,
risk_harm_5_weekend,
risk_harm_5_weekendperc,
ApproveDisapprove2_answers,
ApproveDisapprove2_answers_grade,
ApproveDisapprove2_answers_by_risk,
AgreeDisagree_answers,
AgreeDisagree_answers_grade,
AgreeDisagree_answers_by_risk,
social_atmo_df,
pressured,
social_atmo_df_totals,
social_atmo_df_yes_perc,
social_atmo_7th_yes,
social_atmo_7th_yes_perc,
social_atmo_12th_yes_perc,
social_atmo_bygrade,
seventhand8thsocial_atmocountagree,
ninto12thsocial_atmocountagree,
seventhand8thsocial_atmocountagree_perc,
ninto12thsocial_atmocountagree_perc,
socialatmo,
bingedrinkstring,
totalagree_bingedrinking_7th,
totalagree_bingedrinking_12th,
pressured_df_totals,
pressured_df_yes_perc,
pressured_7th_yes,
pressured_7th_yes_perc,
pressured_12th_yes_perc,
pressured_table,
pressured_bygrade,
seventhand8thpressured_bygradecountagree,
ninto12thpressured_bygradecountagree,
seventhand8thpressured_bygradecountagree_perc,
ninto12thpressured_bygradecountagree_perc,
agree_easy_perc,
agree_easy_perc_val,
own_decision_table,
respect_table,
respect_table_nograde,
respect_yes_perc,
fitin_table_nograde,
fitin_no_perc,
fitinrespectstring,
abstain_table,
easytomake_table,
assume_table,
light_table,
heavy_table,
valued_table,
aca_table,
staff_table,
sincere_table,
care_table,
aca2_table,
aca3_table,
parent_table,
over_table,
college_table,
exp_table,
bored_table,
fitin_table,
NeverOnce_answers,
NeverOnce_answers_grade,
NeverOnce_answers_by_risk,
freq_answers,
freq_answers_grade,
freq_answers_by_risk,
drink_how_often_your_grade_grade,
eighth_perceived_neverdrink_perc,
ninth_perceived_neverdrink_perc,
tenth_perceived_neverdrink_perc,
eleven_perceived_neverdrink_perc,
twelve_perceived_neverdrink_perc,
tenth_perceived_oncemonth_perc,
eleventhmorethantwice_perceived,
twelvethtwiceaweek_perceived,
howmanygrade,
how_many_your_grade_grade,
how_many_your_grade_by_risk,
howmany12,
how_many_seniors_grade,
how_many_seniors_by_risk,
one_drink_days_answers,
one_drink_days_answers_grade,
one_drink_days_answers_by_risk,
howmany30grade,
howmany30grade12,
how_many_30_days_answers_grade,
how_many_30_days_answers_by_risk,
pot_freq_answers,
pot_freq_q_answers_grade,
pot_freq_q_answers_by_risk,
pot_how_often_your_grade_grade,
eighth_perceived_neverpot_perc,
ninth_perceived_neverpot_perc,
tenth_perceived_neverpot_perc,
eleventh_perceived_neverpot_perc,
twelth_perceived_neverpot_perc,
twelthpot_morethanonceamonth_perceived,
twelthpot_more12year_perceived,
percent_ever_used,
percent_ever_used_grade,
percent_ever_used_by_risk,
too_lenient_answers,
too_lenient_answers_grade,
too_lenient_answers_by_risk,
subsetpolicies_1012,
allstudents_policies,
highrisktoostrict,
norisktoostrict,
highrisktoolenient,
norisktoolenient,
subsetenforcement_1012,
allstudents_enforcement,
highrisktoostrict,
norisktoostrict,
highrisktoolenient,
norisktoolenient,
AgreeDisagree2_answers,
AgreeDisagree2_answers_grade,
AgreeDisagree2_answers_by_risk,
academics_to_me_all,
acade_import_all,
valued_perc,
valued_perc_val,
teachers_perc,
teachers_perc_val,
sincere_df,
sincere_perc_val,
care_table,
care_perc_val,
school12,
school12_grade,
school12_by_risk,
schoollikedatlotorkindofliked,
schoollikedatlotorkindofliked7,
schoollikedatlotorkindofliked12,
AgreeDisagree3_answers,
AgreeDisagree3_answers_grade,
AgreeDisagree3_answers_by_risk,
use_knowing,
use_parent_knowing_grade,
use_parent_knowing_by_year,
use_knowing_yes_perc,
use_knowing_10th,
use_knowing_11th,
use_knowing_12th,
use_not_knowing,
use_parent_not_knowing_grade,
use_parent_not_knowing_by_risk,
use_not_knowing_yes_perc,
use_not_knowing_10th,
use_not_knowing_11th,
use_not_knowing_12th,
meals,
meals_grade,
meals_by_risk,
binary_answers,
binary_answers_grade,
binary_answers_by_risk,
Rarely_answers,
Rarely_answers_grade,
Rarely_answers_by_risk,
AgreeDisagree4_answers,
AgreeDisagree4_answers_grade,
AgreeDisagree4_answers_by_risk,
totalagree_parents,
familybulletpoint,
academic_stress,
academic_stress_outputTABLE,
parent_stress_outputTABLE,
YNQ2_answers,
YNQ2_answers_grade,
YNQ2_answers_by_risk,
coping_anx,
coping_anx_perc,
gambling_all,
gambling_grade,
gambling_by_risk,
gambling_any_perc,
lootbox_all,
lootbox_grade,
lootbox_by_risk,
lootbox_any_perc,
gambling_drinking_cross,
gambling_lootbox_cross,
gambling_among_drinkers,
gambling_among_non_drinkers,
lootbox_among_gamblers,
lootbox_among_non_gamblers,
gambling_chart,
lootbox_chart,
gambling_drinking_chart,
gambling_lootbox_chart,
percent8thgraders_ever_drank,
percent10thgraders_ever_drank,
percent12thgraders_ever_drank,
percent_never_drank_allgrades,
percent_never_drank_11th,
percent_never_drank_12th,
lifetime7th,
lifetime12th,
albullet1,
albullet2,
albullet3,
maribullet1,
maribullet2,
maribullet3,
ALLSUBSTANCE8th_30_MTF,
ALLSUBSTANCE10th_30_MTF,
ALLSUBSTANCE12th_30_MTF,
ALLSUBSTANCE8th_12_MTF,
ALLSUBSTANCE10th_12_MTF,
ALLSUBSTANCE12th_12_MTF,
sub12_by_sex_MTFcombo8th,
sub12_by_sex_MTFcombo10th,
sub12_by_sex_MTFcombo12th,
allsubstancebullets,
mostpopular_substances,
alcohollifetumetable,
alcoholl12monthtstable,
alcoholl30daysstable,
pot_lifetime_MTF,
pot_12_MTF,
pot_30_MTF,
VAPE_12_MTF,
VAPE_30_MTF,
tenthbelowmeasure,
twelvebelowmeasure,
eighthbelowmeasure30,
tenthbelowmeasure_30,
twelvebelowmeasure_30,
one10thgraderis,
one12thgraderis,
perceved_actual_table,
perceived_SENIOR_versus_senioractual,
perceved_actual_table_POT,
parentstring,
alcol_melt_chart,
alcohol12monthschart,
alcohol30monthschart,
lifetimepotchart,
annualpotchart,
thirtypotchart,
annualvapechart,
thirtyvapechart,
pressure_chart,
socialatmochart,
owddesc_chart,
respect_chart,
abstain_chart,
easytomake_table_chart,
assume_tablechart,
heavy_table_chart,
light_tablechart,
valued_table_chart,
valued_table_chart,
staff_tablechart,
sincere_tablechart,
care_tablechart,
aca2_table_chart,
school12_attitudes_chart,
highlowrisk_consequence_sub,
eleventhbar_clubperc,
eleventhbar_clubcount,
twelfthbar_clubperc,
vapebullets,
vapebullet8th10th,
nevervape_all,
School,
date_survey,
N,
grades,
N_removed_perc,
len_highrisk_twice_risk,
len_highlowrisk_consequence_sub,
concerned_other_perc10_12,
outputstringssexbullets,
outputstringssexbullets10th,
outputstringssexbullets12th,
dranklast30_6ormoresum,
howmanyhighrisk5_count,
drank30and12_perc_12,
removed_N,
albullet1withare,
prettygrades,
disagree_social_atmo,
oneor2pot10th ,
oneor2pot11th,
oneor2pot12th,
aca_table_chart

]




ReportListofVariablesSTRINGS = ["lifetime8th",
"twelve_8th",
"thirty_8th",
"social_atmo_8th_yes_perc",
"pressured_8th_yes_perc",
"eighthsocial_atmocountagree_perc",
"eighthpressured_bygradecountagree_perc",
"totalagree_bingedrinking_8th",
"sexdf",
"sex_by_grade",
"sex_by_risk",
"gradedf",
"grade_by_risk",
"grade_by_risk10_12",
"highrisk_count_10_12",
"lowrisk_count_10_12",
"norisk_count_10_12",
"avg_age_highrisk",
"avg_age_lowrisk",
"avg_age_norisk",
"age_at_first_drinkdf",
"age_at_first_drink_output",
"age_at_first_grade",
"age_at_first_by_risk",
"totalhighrisk",
"highriskbullet1",
"highrisk_comparison",
"totalstudents",
"perc_high_riskall",
"perc_high_riskallone_in",
"grade_overalldf",
"grades_overall_bygrade",
"grades_overall_by_risk",
"all_risk_options_grades",
"lowriskdrinkers",
"noniskdrinkers",
"percent8thgraders_ever_drank",
"percent10thgraders_ever_drank",
"percent12thgraders_ever_drank",
"percent_never_drank_allgrades",
"percent_never_drank_11th",
"percent_never_drank_12th",
"lifetime7th",
"lifetime12th",
"percent8thgraders_Never_drank",
"percent9thgraders_Never_drank",
"percent10thgraders_Never_drank",
"percent11thgraders_Never_drank",
"percent12thgraders_Never_drank",
"everdrank_grade_sex",
"aggregates",
"agedf",
"where_firstdf",
"where_first_grade",
"where_first_by_risk",
"totalsbygrade",
"onestudentin1012outof100",
"totalsbygrade_andsex",
"first_drink_at_friends",
"often_one_drink",
"often_one_grade",
"often_one_by_risk",
"often_one_by_risk10_12",
"perc_1012_highrisk_twiceormore",
"drinklessthan12timesperyear",
"howoften10th_onceamonthperc",
"eleventhmorethantwice",
"twlthtwiceaweek_perc",
"howmany_df",
"howmany_grade",
"howmany_by_risk",
"howmany_by_risk10_12",
"high_risk_morethan5",
"less_than2_peryear",
"howmanyhighrisk_perc",
"howmanyhighrisk5_perc",
"one_drink12df",
"onewhole_grade",
"onewhole_by_risk",
"twelve_7th",
"twelve_12th",
"total8th_last12_perc",
"tota10th_last12_perc",
"total12th_last12_perc",
"total8th_last12_perc",
"tota10th_last12_perc",
"total12th_last12_perc",
"alcohol_past_12monthbysex",
"neverQ_answers",
"neverQ_answers_grade",
"neverQ_answers_by_risk",
"highriskbulletsssss",
"lowriskbulletsssss",
"higherthan50perc_consequences",
"mostrpopular_lowriskconsequence",
"mostrpopular_lowriskconsequenceperc",
"highrisk_twice_risk",
"high_risk_drunk10_12",
"high_risk_morethan3drunk",
"high_risk_tocut10_12",
"high_risk10_12_cutdown",
"high_risk_guiltyt10_12",
"high_risk10_12_guilty",
"ride_drive",
"later_sorry",
"later_sorry_perc",
"morethanintended",
"one_drink_30df",
"daysatleast1_grade",
"daysatleast1_by_risk",
"dranklast30_6ormore",
"total8th_last30_perc",
"tota10th_last30_perc",
"total12th_last30_perc",
"alcohol_past_30daysbysex",
"noalcoholpaste30",
"noalcoholpaste30perc",
"thirty_7th",
"thirty_12th",
"twormoredays",
"greatest30df",
"greatest30_grade",
"greatest30_by_risk",
"fiveinrowdf",
"days_5_or_more_grade",
"days_5_or_more_by_risk",
"five_ormoreperc",
"five_ormorecount",
"fve_ormore_two_ormore_days_count",
"fve_ormore_two_ormore_days_perc",
"drunk30df",
"drunk_30_grade",
"drunk_30_by_risk",
"wheremostdf",
"where_do_you_grade",
"where_do_you_by_risk",
"wheredrinkingstrings",
"agefirstpot",
"age_first_time_pot_grade",
"age_first_time_pot_by_risk",
"howoftenpot",
"how_often_pot_grade",
"how_often_pot_by_risk",
"neverpot_all",
"neverpot8th",
"neverpot9th",
"neverpot10th",
"neverpot11th",
"neverpot12th",
"twelthpot_moreathan12times",
"twelthpot_morethanonceamonth",
"percent8thgraders_ever_pot",
"percent10thgraders_ever_pot",
"percent12thgraders_ever_pot",
"percent_never_pot_allgrades",
"percent_all_typically",
"pot_lifetime",
"YNQ_answers",
"YNQ_answers_grade",
"YNQ_answers_by_risk",
"cigarettes12",
"notcig12monthall",
"cigarettes_12_risk10_12",
"cigarettes_30_risk10_12",
"pot_last_12_risksubset",
"pot_risk_12_bullets",
"pot_last_30_risksubset",
"OTC_last_12_risksubset",
"OTC_risk_12_bullets",
"VAPE_last_12_risksubset",
"VAPE_risk_12_bullets",
"drank30and12",
"drank12",
"drank30and12_perc",
"use_alcohol_30df",
"past30_yes_drankcount",
"use_alcohol_30_by_gradedf",
"past30_yes_drankcount101112",
"use_alcohol_30_by_riskdf",
"highrisk_yes_last_drink_30",
"percdrank30days_didntdrinkathome",
"percdrank30days_didntdrinkathome_string",
"percent_yes_last30_drank",
"total8th_last12VAPE_perc",
"tota10th_last12VAPE_perc",
"total12th_last12VAPE_perc",
"VAPEPAST12",
"total8th_last30VAPE_perc",
"tota10th_last30VAPE_perc",
"total12th_last30VAPE_perc",
"VAPEPAST30",
"total8th_last12POT_perc",
"tota10th_last12POT_perc",
"total12th_last12POT_perc",
"POTPAST12",
"total8th_last30POT_perc",
"tota10th_last30POT_perc",
"total30th_last30POT_perc",
"POTPAST30",
"all_sub_risks_all",
"all_sub_risks_all_output",
"all_sub_risks_all_output30",
"highriskampusage",
"substances_DF",
"ALLSUBSTANCE8th_12",
"ALLSUBSTANCE10th_12",
"ALLSUBSTANCE12th_12",
"ALLSUBSTANCE8th_30",
"ALLSUBSTANCE10th_30",
"ALLSUBSTANCE12th_30",
"substancesbysexandgrader",
"sub12months_sex8th",
"sub12months_sex10th",
"sub12months_sex12th",
"neverQ2_answers",
"neverQ2_answers_grade",
"neverQ2_answers_by_risk",
"tookcare",
"vom",
"dranktoomuch",
"vom2",
"tookcare_perc",
"vom_perc",
"dranktoomuch_perc1012",
"vom_10_12_perc",
"concerned_other",
"concerned_other_perc",
"schooldance",
"school_dance_use_grade",
"school_dance_use_by_risk",
"danceperc",
"attitude",
"attitude_grade",
"attitude_by_risk",
"neveragoodthingtodo_perc",
"neveragoodthingtodo_perc_high",
"nondrink_okay_attitude",
"highrisk_okay_attitude",
"ApproveDisapprove_answers",
"ApproveDisapprove_answers_grade",
"ApproveDisapprove_answers_by_risk",
"RiskQ_answers",
"RiskQ_answers_grade",
"RiskQ_answers_by_risk",
"risk_harm_1_2",
"risk_harm_daily_1_2perc",
"risk_harm_5_weekend",
"risk_harm_5_weekendperc",
"ApproveDisapprove2_answers",
"ApproveDisapprove2_answers_grade",
"ApproveDisapprove2_answers_by_risk",
"AgreeDisagree_answers",
"AgreeDisagree_answers_grade",
"AgreeDisagree_answers_by_risk",
"social_atmo_df",
"pressured",
"social_atmo_df_totals",
"social_atmo_df_yes_perc",
"social_atmo_7th_yes",
"social_atmo_7th_yes_perc",
"social_atmo_12th_yes_perc",
"social_atmo_bygrade",
"seventhand8thsocial_atmocountagree",
"ninto12thsocial_atmocountagree",
"seventhand8thsocial_atmocountagree_perc",
"ninto12thsocial_atmocountagree_perc",
"socialatmo",
"bingedrinkstring",
"totalagree_bingedrinking_7th",
"totalagree_bingedrinking_12th",
"pressured_df_totals",
"pressured_df_yes_perc",
"pressured_7th_yes",
"pressured_7th_yes_perc",
"pressured_12th_yes_perc",
"pressured_table",
"pressured_bygrade",
"seventhand8thpressured_bygradecountagree",
"ninto12thpressured_bygradecountagree",
"seventhand8thpressured_bygradecountagree_perc",
"ninto12thpressured_bygradecountagree_perc",
"agree_easy_perc",
"agree_easy_perc_val",
"own_decision_table",
"respect_table",
"respect_table_nograde",
"respect_yes_perc",
"fitin_table_nograde",
"fitin_no_perc",
"fitinrespectstring",
"abstain_table",
"easytomake_table",
"assume_table",
"light_table",
"heavy_table",
"valued_table",
"aca_table",
"staff_table",
"sincere_table",
"care_table",
"aca2_table",
"aca3_table",
"parent_table",
"over_table",
"college_table",
"exp_table",
"bored_table",
"fitin_table",
"NeverOnce_answers",
"NeverOnce_answers_grade",
"NeverOnce_answers_by_risk",
"freq_answers",
"freq_answers_grade",
"freq_answers_by_risk",
"drink_how_often_your_grade_grade",
"eighth_perceived_neverdrink_perc",
"ninth_perceived_neverdrink_perc",
"tenth_perceived_neverdrink_perc",
"eleven_perceived_neverdrink_perc",
"twelve_perceived_neverdrink_perc",
"tenth_perceived_oncemonth_perc",
"eleventhmorethantwice_perceived",
"twelvethtwiceaweek_perceived",
"howmanygrade",
"how_many_your_grade_grade",
"how_many_your_grade_by_risk",
"howmany12",
"how_many_seniors_grade",
"how_many_seniors_by_risk",
"one_drink_days_answers",
"one_drink_days_answers_grade",
"one_drink_days_answers_by_risk",
"howmany30grade",
"howmany30grade12",
"how_many_30_days_answers_grade",
"how_many_30_days_answers_by_risk",
"pot_freq_answers",
"pot_freq_q_answers_grade",
"pot_freq_q_answers_by_risk",
"pot_how_often_your_grade_grade",
"eighth_perceived_neverpot_perc",
"ninth_perceived_neverpot_perc",
"tenth_perceived_neverpot_perc",
"eleventh_perceived_neverpot_perc",
"twelth_perceived_neverpot_perc",
"twelthpot_morethanonceamonth_perceived",
"twelthpot_more12year_perceived",
"percent_ever_used",
"percent_ever_used_grade",
"percent_ever_used_by_risk",
"too_lenient_answers",
"too_lenient_answers_grade",
"too_lenient_answers_by_risk",
"subsetpolicies_1012",
"allstudents_policies",
"highrisktoostrict",
"norisktoostrict",
"highrisktoolenient",
"norisktoolenient",
"subsetenforcement_1012",
"allstudents_enforcement",
"highrisktoostrict",
"norisktoostrict",
"highrisktoolenient",
"norisktoolenient",
"AgreeDisagree2_answers",
"AgreeDisagree2_answers_grade",
"AgreeDisagree2_answers_by_risk",
"academics_to_me_all",
"acade_import_all",
"valued_perc",
"valued_perc_val",
"teachers_perc",
"teachers_perc_val",
"sincere_df",
"sincere_perc_val",
"care_table",
"care_perc_val",
"school12",
"school12_grade",
"school12_by_risk",
"schoollikedatlotorkindofliked",
"schoollikedatlotorkindofliked7",
"schoollikedatlotorkindofliked12",
"AgreeDisagree3_answers",
"AgreeDisagree3_answers_grade",
"AgreeDisagree3_answers_by_risk",
"use_knowing",
"use_parent_knowing_grade",
"use_parent_knowing_by_year",
"use_knowing_yes_perc",
"use_knowing_10th",
"use_knowing_11th",
"use_knowing_12th",
"use_not_knowing",
"use_parent_not_knowing_grade",
"use_parent_not_knowing_by_risk",
"use_not_knowing_yes_perc",
"use_not_knowing_10th",
"use_not_knowing_11th",
"use_not_knowing_12th",
"meals",
"meals_grade",
"meals_by_risk",
"binary_answers",
"binary_answers_grade",
"binary_answers_by_risk",
"Rarely_answers",
"Rarely_answers_grade",
"Rarely_answers_by_risk",
"AgreeDisagree4_answers",
"AgreeDisagree4_answers_grade",
"AgreeDisagree4_answers_by_risk",
"totalagree_parents",
"familybulletpoint",
"academic_stress",
"academic_stress_outputTABLE",
"parent_stress_outputTABLE",
"YNQ2_answers",
"YNQ2_answers_grade",
"YNQ2_answers_by_risk",
"coping_anx",
"coping_anx_perc",
"gambling_all",
"gambling_grade",
"gambling_by_risk",
"gambling_any_perc",
"lootbox_all",
"lootbox_grade",
"lootbox_by_risk",
"lootbox_any_perc",
"gambling_drinking_cross",
"gambling_lootbox_cross",
"gambling_among_drinkers",
"gambling_among_non_drinkers",
"lootbox_among_gamblers",
"lootbox_among_non_gamblers",
"gambling_chart",
"lootbox_chart",
"gambling_drinking_chart",
"gambling_lootbox_chart",
"percent8thgraders_ever_drank",
"percent10thgraders_ever_drank",
"percent12thgraders_ever_drank",
"percent_never_drank_allgrades",
"percent_never_drank_11th",
"percent_never_drank_12th",
"lifetime7th",
"lifetime12th",
"albullet1",
"albullet2",
"albullet3",
"maribullet1",
"maribullet2",
"maribullet3",
"ALLSUBSTANCE8th_30_MTF",
"ALLSUBSTANCE10th_30_MTF",
"ALLSUBSTANCE12th_30_MTF",
"ALLSUBSTANCE8th_12_MTF",
"ALLSUBSTANCE10th_12_MTF",
"ALLSUBSTANCE12th_12_MTF",
"sub12_by_sex_MTFcombo8th",
"sub12_by_sex_MTFcombo10th",
"sub12_by_sex_MTFcombo12th",
"allsubstancebullets",
"mostpopular_substances",
"alcohollifetumetable",
"alcoholl12monthtstable",
"alcoholl30daysstable",
"pot_lifetime_MTF",
"pot_12_MTF",
"pot_30_MTF",
"VAPE_12_MTF",
"VAPE_30_MTF",
"tenthbelowmeasure",
"twelvebelowmeasure",
"eighthbelowmeasure30",
"tenthbelowmeasure_30",
"twelvebelowmeasure_30",
"one10thgraderis",
"one12thgraderis",
"perceved_actual_table",
"perceived_SENIOR_versus_senioractual",
"perceved_actual_table_POT",
"parentstring",
"alcol_melt_chart",
"alcohol12monthschart",
"alcohol30monthschart",
"lifetimepotchart",
"annualpotchart",
"thirtypotchart",
"annualvapechart",
"thirtyvapechart",
"pressure_chart",
"socialatmochart",
"owddesc_chart",
"respect_chart",
"abstain_chart",
"easytomake_table_chart",
"assume_tablechart",
"heavy_table_chart",
"light_tablechart",
"valued_table_chart",
"valued_table_chart",
"staff_tablechart",
"sincere_tablechart",
"care_tablechart",
"aca2_table_chart",
"school12_attitudes_chart",
"highlowrisk_consequence_sub",
"eleventhbar_clubperc",
"eleventhbar_clubcount",
"twelfthbar_clubperc",
"vapebullets",
"vapebullet8th10th",
"nevervape_all",
"School",
"date_survey",
"N",
"grades",
"N_removed_perc",
"len_highrisk_twice_risk",                            
"len_highlowrisk_consequence_sub",
"concerned_other_perc10_12",
"outputstringssexbullets",
"outputstringssexbullets10th",
"outputstringssexbullets12th",
"dranklast30_6ormoresum",
"howmanyhighrisk5_count",
"drank30and12_perc_12",
"removed_N",
"albullet1withare",
"prettygrades",
"disagree_social_atmo",
"oneor2pot10th" ,
"oneor2pot11th",
"oneor2pot12th",
"aca_table_chart"

]

assert len(ReportListofVariables) == len(ReportListofVariablesSTRINGS), (
    f"Report list length mismatch: ReportListofVariables has {len(ReportListofVariables)} items, "
    f"ReportListofVariablesSTRINGS has {len(ReportListofVariablesSTRINGS)} items."
)
context_zipped_Report = {ReportListofVariablesSTRINGS[i]: ReportListofVariables[i] for i in range(len(ReportListofVariablesSTRINGS))}


# In[76]:


print('Calculations and whatnot are done. Putting the doc together now. This can easily take 5-10 minutes. Go get a coffee or something')


# In[77]:


#######################################
### Print The Thing! Report
#########################################


## merge the bigggg zipped dict and the small variables
context = {**{
           'date_today': datetime.datetime.now().strftime('%b %d %Y'),
           'has_8th': has_8th,
          }, **  context_zipped_Report }
# QA fixes: new data-driven wording and corrected values (see QA_CONTEXT above)
context.update(QA_CONTEXT)

# FIX for newer pandas: template uses row[0], row[1], etc. (positional integer access)
# Older pandas fell back to positional access on string-labeled Series, newer pandas does not.
# Make copies with integer column names so row[0] works in Jinja2 templates,
# without modifying the original DataFrames (which are still used for charts below).
def _int_cols_copy(df):
    """Return a copy of the DataFrame with integer column names."""
    c = df.copy()
    c.columns = list(range(len(c.columns)))
    return c

# Skip DataFrames that are accessed by column name in the template (not by integer position)
_skip_int_rename = {'attitude_by_risk', 'aggregates'}  # accessed by column name in the template

# When the school has no 8th grade, every school-data 8th-grade column must
# render as '-' (never 0%, 0.0 or an empty cell). MTF norm columns (which
# contain the national comparison values) are left untouched. This is done on
# copies so the original DataFrames (already used for charts) are unaffected.
_no_8th_graders = '8th' not in clean['grade_level'].values

def _dash_8th_columns(df):
    out = df.copy()
    for _col in out.columns:
        if isinstance(_col, str) and '8th' in _col and 'MTF' not in _col:
            out[_col] = '-'
    return out

if _no_8th_graders:
    for _k, _v in list(context.items()):
        if isinstance(_v, pd.DataFrame):
            context[_k] = _dash_8th_columns(_v)
        elif isinstance(_v, list):
            context[_k] = [
                _dash_8th_columns(item) if isinstance(item, pd.DataFrame) else item
                for item in _v
            ]

for _k, _v in list(context.items()):
    if _k in _skip_int_rename:
        continue
    if isinstance(_v, pd.DataFrame):
        context[_k] = _int_cols_copy(_v)
    elif isinstance(_v, list):
        context[_k] = [
            _int_cols_copy(item) if isinstance(item, pd.DataFrame) else item
            for item in _v
        ]

### okay the key was using iterrows in the template


### okay the key was using iterrows in the template

# CORRECTED: Pre-render validation - check for common data issues
logger.info("=== Pre-render Validation ===")
_validation_warnings = 0
for varname, varval in context.items():
    # Check scalar numeric values > 100 that look like percentages
    if isinstance(varval, (int, float)) and 'perc' in varname.lower():
        if varval > 100:
            logger.warning(f"VALIDATION: {varname} = {varval} (exceeds 100%)")
            _validation_warnings += 1
    # Check string percentage values > 100
    if isinstance(varval, str) and '%' in varval:
        try:
            numval = float(varval.replace('%', ''))
            if numval > 100:
                logger.warning(f"VALIDATION: {varname} = {varval} (exceeds 100%)")
                _validation_warnings += 1
        except ValueError:
            pass
if _validation_warnings == 0:
    logger.info("Pre-render validation passed: no percentage values exceed 100%.")
else:
    logger.warning(f"Pre-render validation found {_validation_warnings} potential issue(s). Review warnings above.")

##############################
### R5 CHECK 2: time-order and cross-table consistency, right before rendering
##############################
def _check_time_order():
    errors = []
    pairs = []
    for _c30 in [c for c in clean.columns if re.fullmatch(r'use_[a-z_]+_30', c)]:
        _c12 = _c30[:-3] + '_12'
        if _c12 in clean.columns:
            pairs.append((_c30[4:-3], clean[_c12].eq('Yes'), clean[_c30].eq('Yes')))
    pairs += [('alcohol (follow-up items)', clean['drank_12m'], clean['drank_30d']),
              ('cannabis (derived)', clean['pot_12m'], clean['pot_30d']),
              ('vapes (derived)', clean['vape_12m'], clean['vape_30d'])]
    groups = {'all students': pd.Series(True, index=clean.index)}
    for _g in sorted(set(clean['grade_level'].dropna())):
        _gm = clean['grade_level'].eq(_g)
        groups[_g] = _gm
        for _sx in sorted(set(clean['sex'].dropna())):
            groups[f'{_g} {_sx}'] = _gm & clean['sex'].eq(_sx)
    for _grp in sorted(set(clean['drinker_risk_10_12'].dropna()) - {'Missing'}):
        groups[f'{_grp} (grades 10-12)'] = clean['drinker_risk_10_12'].eq(_grp)
    groups['10th-12th graders'] = clean['grade_level'].isin(BODY_GROUP_GRADES)
    for _name, _m in groups.items():
        _n = int(_m.sum())
        if _n == 0:
            continue
        for _sub, _u12, _u30 in pairs:
            _p12 = round(100 * (_u12 & _m).sum() / _n, 1)
            _p30 = round(100 * (_u30 & _m).sum() / _n, 1)
            if _p30 > _p12:
                errors.append(f'{_name} / {_sub}: past-30-day {_p30}% is higher than past-12-month {_p12}%')
    # Cross-table: the substance-list answers and the derived flags feed different
    # tables (e.g. Cannabis section vs. substance comparison), so they must match.
    for _label, _raw, _derived in [('cannabis 12-month', clean['use_pot_12'].eq('Yes'), clean['pot_12m']),
                                   ('cannabis 30-day', clean['use_pot_30'].eq('Yes'), clean['pot_30d']),
                                   ('vapes 12-month', clean['use_vapes_12'].eq('Yes'), clean['vape_12m']),
                                   ('vapes 30-day', clean['use_vapes_30'].eq('Yes'), clean['vape_30d'])]:
        _diff = int((_raw != _derived).sum())
        if _diff:
            errors.append(f'{_label}: substance-list answer and derived flag disagree for {_diff} student(s), '
                          f'so two tables would show different numbers')
    return errors

_time_errors = _check_time_order()
if _time_errors:
    for _e in _time_errors:
        logger.error('R5 CHECK 2: ' + _e)
    if FAIL_ON_TIME_ORDER_ERRORS:
        raise ValueError('R5 CHECK 2: report not built. Impossible or inconsistent values:\n  '
                         + '\n  '.join(_time_errors))
else:
    print('R5 CHECK 2 passed: no group shows past-30-day use above past-12-month use, and cross-table counts match.')

### If British year groups were remapped, convert grade labels back to Year X for display
_GRADE_TO_YEAR = {}
if '_YEAR_TO_GRADE' in dir() and '_was_year_system' in clean.columns:
    _GRADE_TO_YEAR = {_YEAR_TO_GRADE[y]: y for y in ['Year 9', 'Year 10', 'Year 11', 'Year 12', 'Year 13']
                      if y in _YEAR_TO_GRADE}
if _GRADE_TO_YEAR:
    _YEAR_SUBS = []
    for _us, _yr in sorted(_GRADE_TO_YEAR.items(), key=lambda x: -len(x[0])):
        _YEAR_SUBS.extend([
            (f'{_us}-Grade', _yr), (f'{_us}-grade', _yr),
            (f'{_us} Graders', f'{_yr} Students'), (f'{_us} graders', f'{_yr} students'),
            (f'{_us} Grader', f'{_yr} Student'), (f'{_us} grader', f'{_yr} student'),
            (f'{_us} grade', _yr), (f'{_us} Grade', _yr),
            (f'{_us} (', f'{_yr} ('), (f'{_us}:', f'{_yr}:'),
            (f'{_us} ', f'{_yr} '), (f'{_us},', f'{_yr},'), (f'{_us}.', f'{_yr}.'),
        ])
    _YEAR_SUBS.extend([
        ('grades 8, 9, 10, 11 and 12', 'years 9, 10, 11, 12 and 13'),
        ('grades 8, 10 and 12', 'years 9, 11 and 13'),
        ('grade-level', 'year-level'), ('Grade-Level', 'Year-Level'),
        ('grade level', 'year group'), ('Grade Level', 'Year Group'),
    ])

    def _relabel_str(s):
        if not isinstance(s, str):
            return s
        for _old, _new in _YEAR_SUBS:
            s = s.replace(_old, _new)
        for _us, _yr in _GRADE_TO_YEAR.items():
            if s == _us:
                s = _yr
        s = re.sub(r'in (Year \d+) (Year \d+)', r'in \1', s)
        s = re.sub(r'(Year \d+) (Year \d+)', r'\1', s)
        return s

    def _relabel(obj):
        if isinstance(obj, str):
            return _relabel_str(obj)
        if isinstance(obj, pd.DataFrame):
            obj = obj.copy()
            # Keep numbered columns as numbers: the template reads cells by
            # position (row[1], row[5], ...). Turning 0, 1, 2 into '0', '1', '2'
            # left every one of those table cells empty.
            obj.columns = [_relabel_str(c) if isinstance(c, str) else c for c in obj.columns]
            for col in obj.select_dtypes(include='object').columns:
                obj[col] = obj[col].apply(lambda v: _relabel_str(v) if isinstance(v, str) else v)
            return obj
        if isinstance(obj, list):
            return [_relabel(x) for x in obj]
        return obj

    for _k in list(context.keys()):
        context[_k] = _relabel(context[_k])
    print(f"Relabeled grade references to Year system: {_GRADE_TO_YEAR}")

# C18: in the typical-frequency tables, "Never" includes lifetime drinkers who
# typically never drink, so it is labelled "Typically never" (it is larger than
# the "never had a whole drink" figure quoted above it).
for _k in ['often_one_drink', 'often_one_grade', 'often_one_by_risk', 'often_one_by_risk10_12']:
    _v = context.get(_k)
    if isinstance(_v, pd.DataFrame) and len(_v.columns):
        _v = _v.copy()
        _v[_v.columns[0]] = _v[_v.columns[0]].replace({'Never': 'Typically never'})
        context[_k] = _v

# C16: percentages with a small base carry an invisible marker into the
# document; the step right after rendering turns it into "50.0% (1 of 2)".
for _k, _v in list(context.items()):
    if isinstance(_v, _Pct):
        context[_k] = _v.marked()
    elif isinstance(_v, list) and any(isinstance(_i, _Pct) for _i in _v):
        context[_k] = [_i.marked() if isinstance(_i, _Pct) else _i for _i in _v]
    elif isinstance(_v, tuple) and any(isinstance(_i, _Pct) for _i in _v):
        context[_k] = tuple(_i.marked() if isinstance(_i, _Pct) else _i for _i in _v)

try:
    doc.render(context)
    print("doc.render() succeeded")
except Exception as _render_err:
    import traceback
    print("ERROR during doc.render():")
    traceback.print_exc()
    raise

### this actually adds the page numbers

# POST-RENDER: Year-system relabeling on the rendered doc (catches code-generated strings)
if _GRADE_TO_YEAR:
    _post_replacements = [
        ('8th graders', 'Year 9 students'), ('10th graders', 'Year 11 students'), ('12th graders', 'Year 13 students'),
        ('8th Graders', 'Year 9 Students'), ('10th Graders', 'Year 11 Students'), ('12th Graders', 'Year 13 Students'),
        ('9th graders', 'Year 10 students'), ('11th graders', 'Year 12 students'),
        ('9th Graders', 'Year 10 Students'), ('11th Graders', 'Year 12 Students'),
        ('8th-Grade', 'Year 9'), ('10th-Grade', 'Year 11'), ('12th-Grade', 'Year 13'),
        ('9th-Grade', 'Year 10'), ('11th-Grade', 'Year 12'),
        ('8th-grade', 'Year 9'), ('10th-grade', 'Year 11'), ('12th-grade', 'Year 13'),
        ('9th-grade', 'Year 10'), ('11th-grade', 'Year 12'),
        ('8th grade', 'Year 9'), ('9th grade', 'Year 10'), ('10th grade', 'Year 11'),
        ('11th grade', 'Year 12'), ('12th grade', 'Year 13'),
        ('8th Grade', 'Year 9'), ('9th Grade', 'Year 10'), ('10th Grade', 'Year 11'),
        ('11th Grade', 'Year 12'), ('12th Grade', 'Year 13'),
        ('in Year 9 Year 9', 'in Year 9'), ('in Year 10 Year 10', 'in Year 10'),
        ('in Year 11 Year 11', 'in Year 11'), ('in Year 12 Year 12', 'in Year 12'),
        ('in Year 13 Year 13', 'in Year 13'),
        ('grades 8, 9, 10, 11 and 12', 'years 9, 10, 11, 12 and 13'),
        ('grades 8, 10 and 12', 'years 9, 11 and 13'),
        ('grades 8, 10, and 12', 'years 9, 11, and 13'),
        ('10th-12th graders', 'Year 11-Year 13 students'),
        ('10th-Year 13 students', 'Year 11-Year 13 students'),
        ('Results by Grade Level', 'Results by Year Group'),
        ('grade level', 'year group'), ('Grade Level', 'Year Group'),
        ('students in your grade', 'students in your year'),
        ('Students in your grade', 'Students in your year'),
        ('What grade are you in?', 'What year group are you in?'),
        ('grade students', 'students'),
        ('8th, 10th and 12th graders', 'Year 9, Year 11 and Year 13 students'),
        ('8th, 10th, and 12th graders', 'Year 9, Year 11, and Year 13 students'),
        ('8th, 10th and 12th Graders', 'Year 9, Year 11 and Year 13 Students'),
        ('8th, 10th, and 12th Graders', 'Year 9, Year 11, and Year 13 Students'),
        ('Seniors (12th graders)', 'Year 13 students'),
        ('seniors (12th graders)', 'Year 13 students'),
        ('Seniors (Year 13 students)', 'Year 13 students'),
        ('grades 8, 10 and 12 students', 'Year 9, Year 11 and Year 13 students'),
        ('grades 10-12', 'Years 11-13'),
        ('grades 10-Year 13', 'Years 11-13'),
        ('in their grade', 'in their year'),
        ('in their Grade', 'in their Year'),
        ('in your grade', 'in your year'),
        ('In Your Grade', 'In Your Year'),
        ('grade to', 'year to'),
        ('grade students', 'students'),
        ('surveyed in grades', 'surveyed in years'),
        ('students in grades', 'students in years'),
        ('all students in grades', 'all students in years'),
        ('grade to grade', 'year to year'),
        ('certain grades', 'certain years'),
        ('all grades surveyed', 'all years surveyed'),
        ('all grades say', 'all years say'),
        ('Across all grades', 'Across all years'),
        ('of 5 grades', 'of 5 years'),
        ('older grades', 'older years'),
        ('younger grades', 'younger years'),
        ('advanced grades', 'advanced years'),
        ('youngest grade', 'youngest year'),
        ('grade-', 'year-'),
        ('Year 9 graders', 'Year 9 students'),
        ('Year 10 graders', 'Year 10 students'),
        ('Year 11 graders', 'Year 11 students'),
        ('Year 12 graders', 'Year 12 students'),
        ('Year 13 graders', 'Year 13 students'),
        ('in that grade', 'in that year'),
        ('at grades', 'at years'),
        ('looking at grades', 'looking at years'),
        ('Grade 8', 'Year 9'), ('Grade 9', 'Year 10'), ('Grade 10', 'Year 11'),
        ('Grade 11', 'Year 12'), ('Grade 12', 'Year 13'),
        ('grade 8', 'Year 9'), ('grade 9', 'Year 10'), ('grade 10', 'Year 11'),
        ('grade 11', 'Year 12'), ('grade 12', 'Year 13'),
    ]
    # Handle bare table cell headers
    _BARE_CELL_SUBS = {
        'Grade': 'Year Group',
        '8th': 'Year 9', '9th': 'Year 10', '10th': 'Year 11',
        '11th': 'Year 12', '12th': 'Year 13',
        '8th MTF': '8th MTF', '10th MTF': '10th MTF', '12th MTF': '12th MTF',
    }
    def _fix_bare_grade_headers(_p):
        t = _p.text.strip()
        if t in _BARE_CELL_SUBS and _p.runs:
            _p.runs[0].text = _BARE_CELL_SUBS[t]
            for _r in _p.runs[1:]:
                _r.text = ''
            return t != _BARE_CELL_SUBS[t]
        return False
    def _year_replace_para(_p):
        """Replace grade references in a paragraph, handling cross-run text."""
        full = _p.text
        new_full = full
        for _old, _repl in _post_replacements:
            new_full = new_full.replace(_old, _repl)
        if new_full != full and _p.runs:
            _p.runs[0].text = new_full
            for _r in _p.runs[1:]:
                _r.text = ''
            return 1
        return 0

    _post_count = 0
    for _p in doc.paragraphs:
        _post_count += _year_replace_para(_p)
    for _tbl in doc.tables:
        for _row in _tbl.rows:
            for _cell in _row.cells:
                for _p in _cell.paragraphs:
                    _post_count += _year_replace_para(_p)
                    if _fix_bare_grade_headers(_p):
                        _post_count += 1
    print(f"Post-render Year relabeling: fixed {_post_count} paragraphs")

# POST-RENDER FIX: comprehensive formatting cleanup
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn as _qn

from docx.text.paragraph import Paragraph as _DocxParagraph

def _iter_all_paragraphs(_doc):
    """Every paragraph in the body, including table cells (nested too)."""
    for _pel in _doc.element.body.iter(_qn('w:p')):
        yield _DocxParagraph(_pel, _doc)

# C16: expand small-base markers into "50.0% (1 of 2)", run by run, so the
# count lands after the % sign even when % is in the next run.
_MARK_ONLY = re.compile('⁣N(\\d+)D(\\d+)⁣')
for _p in _iter_all_paragraphs(doc):
    if '⁣' not in _p.text:
        continue
    _runs = _p.runs
    for _ri, _r in enumerate(_runs):
        while True:
            _m = _MARK_ONLY.search(_r.text)
            if not _m:
                break
            _suffix = f" ({_m.group(1)} of {_m.group(2)})"
            _before, _after = _r.text[:_m.start()], _r.text[_m.end():]
            if _after.startswith('%'):
                _r.text = _before + '%' + _suffix + _after[1:]
            elif not _after and _ri + 1 < len(_runs) and _runs[_ri + 1].text.startswith('%'):
                _r.text = _before
                _runs[_ri + 1].text = '%' + _suffix + _runs[_ri + 1].text[1:]
            else:
                _r.text = _before + _suffix + _after
    if '⁣' in _p.text:      # marker split across runs: rebuild the paragraph text
        _set_txt = expand_small_n(_p.text).replace('⁣', '')
        if _p.runs:
            _p.runs[0].text = _set_txt
            for _r in _p.runs[1:]:
                _r.text = ''

_nan_pattern = re.compile(r'^nan$', re.IGNORECASE)
_bare_number = re.compile(r'^-?\d+\.?\d*$')  # matches "17.0", "3", "100.0" etc.
# Garbled em-dash from MTF CSV (UTF-8 bytes read as latin-1): represents missing data
_garbled_dash = re.compile(r'^[â€"\x80-\x94\u2014—]+$')

# Prevention Ed brand colors for table headers
_BRAND_NAVY = '1F3864'
_BRAND_BLUE = '1A63CD'
_BRAND_LIGHT_BLUE = '267EF2'
# Map old template colors (uppercased) to brand colors
_HEADER_COLOR_MAP = {
    '2B5797': _BRAND_NAVY,
    '8DB4E2': _BRAND_LIGHT_BLUE,
}

# --- TABLE FORMATTING ---
for _tbl in doc.tables:
    # Determine how many header rows this table has (rows where first cell is empty
    # and data rows follow, or single-row highlight boxes)
    _num_header_rows = 1
    if len(_tbl.rows) >= 2:
        # Check if row 1 is also a header (e.g., "MTF" / school name sub-header)
        _row1_texts = [c.text.strip() for c in _tbl.rows[1].cells]
        if all(t == '' or t in ('MTF', '') or not _bare_number.match(t.replace('%','')) for t in _row1_texts):
            # Row 1 looks like a sub-header if it has text but no numeric data
            if any(t and not _bare_number.match(t.replace('%','').replace('-','')) for t in _row1_texts[1:]):
                _num_header_rows = 2

    # Columns whose header says Count/N hold student counts, not percentages —
    # never append '%' to them (appendix tables have Count (N) columns).
    _count_cols = set()
    for _hr in range(min(2, len(_tbl.rows))):
        for _ci2, _c2 in enumerate(_tbl.rows[_hr].cells):
            _h = _c2.text.strip()
            if _h in ('N', 'Count') or '(N)' in _h:
                _count_cols.add(_ci2)

    for _ri, _row in enumerate(_tbl.rows):
        _is_header = _ri < _num_header_rows
        # Merged cells appear once per grid column they span; when data-row
        # spans don't line up with header spans, judge each cell by its
        # RIGHTMOST grid position so Count columns are detected reliably.
        _row_cells = _row.cells
        _tc_positions = {}
        for _ci0, _c0 in enumerate(_row_cells):
            _tc_positions.setdefault(id(_c0), set()).add(_ci0)
        _row_label = _row_cells[0].text.strip() if _row_cells else ''
        _is_stat_row = _row_label in ('Minimum age', 'Maximum age', 'Mean age', 'Standard deviation', 'Aggregates')
        for _ci, _cell in enumerate(_row_cells):
            _eff_ci = max(_tc_positions[id(_cell)])
            # --- Fix 4: Update header cell colors to brand palette ---
            if _is_header:
                _tcPr = _cell._tc.find(_qn('w:tcPr'))
                if _tcPr is not None:
                    _shd = _tcPr.find(_qn('w:shd'))
                    if _shd is not None:
                        _fill = _shd.get(_qn('w:fill'))
                        if _fill and _fill.upper() in _HEADER_COLOR_MAP:
                            _shd.set(_qn('w:fill'), _HEADER_COLOR_MAP[_fill.upper()])

            for _para in _cell.paragraphs:
                _txt = _para.text.strip()
                # Replace nan/NaN with '-'
                if _nan_pattern.match(_txt):
                    for _run in _para.runs:
                        _run.text = ''
                    if _para.runs:
                        _para.runs[0].text = '-'
                # Replace garbled em-dash characters (â, â€", etc.) with '-'
                elif _garbled_dash.match(_txt) or ('â' in _txt and len(_txt) <= 5):
                    for _run in _para.runs:
                        _run.text = ''
                    if _para.runs:
                        _para.runs[0].text = '-'
                # Add '%' to bare numbers in data cells (skip headers, skip
                # col 0 labels, skip Count/N columns)
                elif not _is_header and not _is_stat_row and _eff_ci > 0 and _eff_ci not in _count_cols and _bare_number.match(_txt):
                    for _run in _para.runs:
                        if _bare_number.match(_run.text.strip()):
                            _run.text = _run.text.strip() + '%'
                            break
                # Count columns: display whole numbers ("16", not "16.0")
                elif not _is_header and _eff_ci in _count_cols and _bare_number.match(_txt) and _txt.endswith('.0'):
                    for _run in _para.runs:
                        if _bare_number.match(_run.text.strip()):
                            _run.text = _run.text.strip()[:-2]
                            break
                # Center all table cell paragraphs
                _para.alignment = WD_ALIGN_PARAGRAPH.CENTER

                # --- Fix 2: Un-bold data cells, keep headers bold ---
                for _run in _para.runs:
                    if _is_header:
                        _run.font.bold = True
                    else:
                        _run.font.bold = False

                    # --- Fix 5: Normalize font size in data cells ---
                    if not _is_header and _run.font.size is not None:
                        _run.font.size = None  # inherit from style (consistent sizing)

# --- Fix 1: Replace nan/nan% in paragraph text (not just tables) ---
for _para in doc.paragraphs:
    _full_text = _para.text
    if 'nan' in _full_text.lower():
        for _run in _para.runs:
            if 'nan%' in _run.text.lower():
                _run.text = _run.text.replace('nan%', '0.0%').replace('NaN%', '0.0%').replace('Nan%', '0.0%')
            elif 'nan' in _run.text.lower() and '%' not in _run.text:
                _run.text = re.sub(r'\bnan\b', '0.0', _run.text, flags=re.IGNORECASE)

# --- Fix 7: No-8th-grade narrative cleanup ---
# When the school has no 8th grade, bullet sentences about 8th-grade data
# render dangling (e.g. "of 8th graders report never using marijuana.").
# Remove those bullets, and rewrite the "ranging from X% of 8th graders" /
# "rise from X% in 8th grade" sentences to start from 9th grade instead.
if _no_8th_graders:
    def _num_or_none(fn):
        try:
            return str(round(float(str(fn()).replace('%', '')), 1))
        except Exception:
            return None

    _sa9 = _num_or_none(lambda: socialatmo.loc[socialatmo['index'] == 'Total Agree', '9th (N)'].item())
    _pr9 = _num_or_none(lambda: pressured_table.loc[pressured_table['index'] == 'Total Agree', '9th (N)'].item())
    _lifetime9 = _num_or_none(lambda: (age_at_first_grade[age_at_first_grade['Age'] != 'I have never had a whole drink of alcohol']['9th (%)'].str.split('%').str[0]).astype(float).sum())
    _twelve9 = _num_or_none(lambda: (onewhole_grade[onewhole_grade['Amount'] == 'Yes']['9th (%)'].str.split('%').str[0]).astype(float).sum())
    _thirty9 = _num_or_none(lambda: (daysatleast1_grade[daysatleast1_grade['Response'] != 'I did not drink alcohol during the past 30 days']['9th (%)'].str.split('%').str[0]).astype(float).sum())

    _binge9 = _num_or_none(lambda: AgreeDisagree_answers_grade[8][2:]['9th (%)'].str.split('%').str[0].astype(float).sum())

    _range_re = re.compile(r'ranging from\s*(?:N/A|nan|[\d.]*)\s*% of 8th graders to')
    _example_re = re.compile(r'For example,\s*(?:N/A|nan|[\d.]*)\s*% of 8th graders agree')
    _rise_re = re.compile(r'(rises?) from\s*(?:N/A|nan|[\d.]*)\s*% in 8th grade to')
    _from8_re = re.compile(r'from\s*(?:N/A|nan|[\d.]*)\s*% in 8th grade to')
    _believe8_re = re.compile(r'(?:N/A|nan|[\d.]*)\s*% of 8th graders believe that')

    def _rewrite_8th_sentences(_p):
        _t = _p.text
        _new = _t
        if _range_re.search(_new):
            _val = _pr9 if 'pressured' in _new else _sa9
            if _val is not None:
                _new = _range_re.sub('ranging from %s%% of 9th graders to' % _val, _new)
        if _example_re.search(_new) and _sa9 is not None:
            _new = _example_re.sub('For example, %s%% of 9th graders agree' % _sa9, _new)
        _m = _rise_re.search(_new)
        if _m:
            _val = None
            # Check the 30-day/12-month sentences first: the past-30-day
            # sentence also contains the word "lifetime" in its opening clause.
            if 'past-30-day' in _new:
                _val = _thirty9
            elif 'past-12-months' in _new:
                _val = _twelve9
            elif 'lifetime' in _new:
                _val = _lifetime9
            if _val is not None:
                _new = _rise_re.sub('%s from %s%% in 9th grade to' % (_m.group(1), _val), _new)
        # "...increases with age from % in 8th grade to X% in 12th grade"
        if _from8_re.search(_new):
            _val = _pr9 if 'pressured' in _new else (_sa9 if 'social atmosphere' in _new else None)
            if _val is not None:
                _new = _from8_re.sub('from %s%% in 9th grade to' % _val, _new)
        # Culture-of-use intro: binge-drinking belief and attitude sentences
        if _believe8_re.search(_new) and _binge9 is not None:
            _new = _believe8_re.sub('%s%% of 9th graders believe that' % _binge9, _new)
        _new = _new.replace('Over half of 8th graders also believe', 'Over half of 9th graders also believe')
        # Executive summary boilerplate written for 8th-12th schools
        _new = _new.replace('use from 8th grade to 12th grade, with levels that start below U.S. norms for 8th graders',
                            'use from 9th grade to 12th grade, with levels that start below U.S. norms in the younger grades')
        # Comparison-section intro paragraphs (static template text with no
        # 8th-grade guard): drop the dangling/false 8th-grade clauses.
        _new = re.sub(r'While alcohol use by 8th graders at .{1,80}? are\s*,\s*use by 10th graders',
                      'Use by 10th graders', _new)
        _new = re.sub(r'cannabis use by (.{1,80}?) 8th\s+and 10th graders', r'cannabis use by \1 10th graders', _new)
        _new = re.sub(r'vaping by (.{1,80}?) 8th and 10th graders', r'vaping by \1 10th graders', _new)
        # 8th-grade substance-table intros: hardcoded below-norms claim and
        # "(One 8th grader equals X%)" margin make no sense with no 8th data.
        _new = re.sub(
            r'The comparisons, allowing for a margin of error, show that (?:substance use by 8th graders is below\s+U\.S\. norms for (?:all available measures|[\d.]*\s*% of\s*available measures)|8th graders use less than the U\.S\. average on [\d.]*\s*% of the substances compared)\.?\s*\(One 8th grader equals\s*[\d.]*\s*%(?: in these figures\.)?\)\.?',
            'No 8th graders took the survey, so no comparison is shown; the U.S. norms appear in the table for reference.',
            _new)
        if _new != _t:
            for _run in _p.runs:
                _run.text = ''
            if _p.runs:
                _p.runs[0].text = _new
            else:
                _p.add_run(_new)

    _delete_8th_res = [
        re.compile(r'Alcohol use by .{0,120} 8th graders is\s*$'),
        re.compile(r'Cannabis use by .{0,120} 8th graders\s*$'),
        re.compile(r'^\s*(?:•\s*)?(?:Yet, only\s*)?(?:N/A|nan|[\d.]+\s*%|%)?\s*of 8th grade(?:rs| students)\b'),
    ]

    def _is_dangling_8th(_txt):
        return any(_r.search(_txt) for _r in _delete_8th_res)

    def _remove_paragraph(_p):
        _el = _p._element
        _el.getparent().remove(_el)

    for _p in list(doc.paragraphs):
        _rewrite_8th_sentences(_p)
        if _is_dangling_8th(_p.text):
            _remove_paragraph(_p)
    for _tbl in doc.tables:
        for _row in _tbl.rows:
            for _cell in _row.cells:
                for _p in list(_cell.paragraphs):
                    if _is_dangling_8th(_p.text) and len(_cell.paragraphs) > 1:
                        _remove_paragraph(_p)

# --- Fix 9: 8th-grade schools: the template hardcodes another school's
# "(One 8th grader equals X%)" margins, and the 12-month 8th intro hardcodes
# "below U.S. norms for all available measures" (no placeholders exist for
# either). Replace both with values computed from this school's data.
if not _no_8th_graders:
    for _para in doc.paragraphs:
        _t0 = _para.text
        if 'One 8th grader equals' not in _t0:
            continue
        _t1 = _t0
        if one8thgraderis != '':
            _t1 = re.sub(r'\(One 8th grader equals [\d.]+%', '(One 8th grader equals %s%%' % one8thgraderis, _t1)
        if eighthbelowmeasure12 != '' and 'past year' in _t1:
            _t1 = _t1.replace('below U.S. norms for all available measures',
                              'below U.S. norms for %s%% of available measures' % eighthbelowmeasure12)
        if _t1 != _t0:
            for _run in _para.runs:
                _run.text = ''
            if _para.runs:
                _para.runs[0].text = _t1
            else:
                _para.add_run(_t1)

# --- Fix 10: punctuation artifacts where bullets ending in "measures." are
# spliced mid-sentence in the comparison intros ("measures. and use by",
# "measures., use by", trailing "measures..").
for _para in doc.paragraphs:
    _t0 = _para.text
    if _t0.startswith('The following tables compare') and 'measures.' in _t0:
        _t1 = (_t0.replace('measures. and use by', 'measures, and use by')
                  .replace('measures., use by', 'measures, use by')
                  .replace('measures..', 'measures.'))
        if _t1 != _t0:
            for _run in _para.runs:
                _run.text = ''
            if _para.runs:
                _para.runs[0].text = _t1
            else:
                _para.add_run(_t1)

# --- Fix 11: drop empty bullet paragraphs. Padded placeholder slots (e.g.
# wheredrinkingstrings[4] at a school with 4 locations, or the 8th-grade
# highlight at a 9th-12th school) render as a bullet with no text.
def _is_empty_bullet(_p):
    if _p.text.strip():
        return False
    _pPr = _p._p.find(_qn('w:pPr'))
    return _pPr is not None and _pPr.find(_qn('w:numPr')) is not None
for _para in list(doc.paragraphs):
    if _is_empty_bullet(_para):
        _para._element.getparent().remove(_para._element)
for _tbl in doc.tables:
    for _row in _tbl.rows:
        for _cell in _row.cells:
            for _para in list(_cell.paragraphs):
                if _is_empty_bullet(_para) and len(_cell.paragraphs) > 1:
                    _para._element.getparent().remove(_para._element)

# --- Fix 8: Letter-grade schools: the template's academics bullets are
# written for the IB 1-7 scale ("Mostly 7's" = top marks, "Mostly 4's" =
# middling). Swap in the letter equivalents (A's / C's) when the school
# grades with letters.
if listofageoptions == _letter_scale:
    for _p in doc.paragraphs:
        if 'Mostly 7' in _p.text or 'Mostly 4' in _p.text:
            _new = (_p.text.replace('Mostly 7’s', 'Mostly A’s').replace("Mostly 7's", "Mostly A's")
                          .replace('Mostly 4’s', 'Mostly C’s').replace("Mostly 4's", "Mostly C's"))
            if _new != _p.text:
                for _run in _p.runs:
                    _run.text = ''
                if _p.runs:
                    _p.runs[0].text = _new
                else:
                    _p.add_run(_new)

# --- Fix 3: Normalize line spacing (2.0 → 1.3 for bullet lists) ---
for _para in doc.paragraphs:
    if _para.paragraph_format.line_spacing == 2.0:
        _para.paragraph_format.line_spacing = 1.3

# --- Fix 6: Equalize column widths on 6-column grade tables with uneven distribution ---
from docx.shared import Inches as _Inches
for _tbl in doc.tables:
    if len(_tbl.columns) == 6 and len(_tbl.rows) >= 2:
        _row0 = [c.text.strip() for c in _tbl.rows[0].cells]
        if 'Grade 12' in _row0:
            _widths = [c.width for c in _tbl.rows[0].cells]
            _grade_widths = _widths[1:]
            # Only fix if the widths are noticeably uneven
            if max(_grade_widths) > min(_grade_widths) * 1.3:
                _label_w = _Inches(1.5)
                _grade_w = _Inches(0.9)
                for _row in _tbl.rows:
                    _row.cells[0].width = _label_w
                    for _ci in range(1, 6):
                        _row.cells[_ci].width = _grade_w

# --- Fix 7: Enforce Calibri font globally on all paragraph runs ---
for _para in doc.paragraphs:
    for _run in _para.runs:
        _run.font.name = 'Calibri'

# --- Fix 8: Apply brand colors to headings ---
_HEADING_BRAND = {
    'Heading': RGBColor(0x1F, 0x38, 0x64),   # Navy (old template style name)
    'Heading 1': RGBColor(0x1F, 0x38, 0x64), # Navy (section headings, built-in Heading 1)
    'Heading 2': RGBColor(0x1F, 0x38, 0x64),  # Navy
    'Heading 3': RGBColor(0x40, 0x40, 0x40),  # Dark gray
    'Heading 4': RGBColor(0x40, 0x40, 0x40),  # Dark gray
    'Title': RGBColor(0x1F, 0x38, 0x64),      # Navy
}
for _para in doc.paragraphs:
    _style_name = _para.style.name if _para.style else None
    if _style_name in _HEADING_BRAND:
        for _run in _para.runs:
            _run.font.color.rgb = _HEADING_BRAND[_style_name]
            _run.font.name = 'Calibri'

# --- Fix 9: Normalize table borders (thin light gray on all cells) ---
def _set_cell_borders(cell, color='D9D9D9', sz='4'):
    """Set uniform thin borders on a table cell."""
    tc = cell._tc
    tcPr = tc.find(_qn('w:tcPr'))
    if tcPr is None:
        tcPr = OxmlElement('w:tcPr')
        tc.insert(0, tcPr)
    tcBorders = tcPr.find(_qn('w:tcBorders'))
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    for edge in ('top', 'left', 'bottom', 'right'):
        elem = OxmlElement(f'w:{edge}')
        elem.set(_qn('w:val'), 'single')
        elem.set(_qn('w:sz'), sz)
        elem.set(_qn('w:space'), '0')
        elem.set(_qn('w:color'), color)
        # Remove existing if present
        existing = tcBorders.find(_qn(f'w:{edge}'))
        if existing is not None:
            tcBorders.remove(existing)
        tcBorders.append(elem)

def _set_cell_valign(cell, align='center'):
    """Set vertical alignment on a table cell."""
    tc = cell._tc
    tcPr = tc.find(_qn('w:tcPr'))
    if tcPr is None:
        tcPr = OxmlElement('w:tcPr')
        tc.insert(0, tcPr)
    vAlign = tcPr.find(_qn('w:vAlign'))
    if vAlign is None:
        vAlign = OxmlElement('w:vAlign')
        tcPr.append(vAlign)
    vAlign.set(_qn('w:val'), align)

for _tbl in doc.tables:
    _is_highlight = len(_tbl.rows) == 1 and len(_tbl.columns) == 1
    for _row in _tbl.rows:
        for _cell in _row.cells:
            if not _is_highlight:
                # Fix 9: Borders on multi-cell tables only
                _set_cell_borders(_cell)
            # Fix 10: Vertical center alignment on all cells
            _set_cell_valign(_cell)

# --- Fix 11: OPTION C — Subtle background with dot bullets ---
# Restyle all highlight boxes (1x1 tables) to Option C design:
# - Light gray background, subtle border
# - "HIGHLIGHTS" header in small uppercase muted text
# - Dot bullets with blue stat numbers, normal-weight body text

_OC_BG = 'F0F2F5'           # Light cool-gray card background
_OC_BORDER = 'CDD2D9'       # Subtle border color
_OC_STAT_BLUE = RGBColor(0x18, 0x5F, 0xA5)   # #185FA5 for key stats
_OC_BODY_COLOR = RGBColor(0x2D, 0x2D, 0x2D)   # Dark body text
_OC_MUTED = RGBColor(0x1F, 0x38, 0x64)        # Brand Navy #1F3864 for "HIGHLIGHTS" label
_OC_BULLET_CHAR = '\u2022'                     # Filled circle bullet

# Regex to find stat-worthy tokens — must be specific enough to avoid mid-word numbers
# Uses word boundaries and lookahead/lookbehind to avoid "past-12-months" false positives
_stat_pattern = re.compile(
    r'('
    r'\d[\d,]*\.?\d*%'                              # percentages: "14.3%", "68.0%"
    r'|\d[\d,]+\s+students'                          # "50 students"
    r'|\d[\d,]+\s+\d+(?:st|nd|rd|th)-\d+(?:st|nd|rd|th)\s+graders'  # "49 10th-12th graders"
    r'|\d[\d,]+\s+(?:10th|11th|12th|8th|9th)\s'      # "49 10th " (number + grade)
    r'|\b\d+(?:st|nd|rd|th)-\d+(?:st|nd|rd|th)\s+graders'  # "10th-12th graders"
    r'|\(\d[\d,]*\.?\d*%\)'                          # "(14.3%)" with parens
    r'|\b\d[\d,]+\s*\(\d[\d,]*\.?\d*%\)'             # "7 (14.3%)"
    r')'
)

# Words/phrases that indicate the line is a title, not content — skip these
_HIGHLIGHT_TITLE_WORDS = {'highlights', 'highlight', 'key findings', 'key data'}

for _tbl in doc.tables:
    if not (len(_tbl.rows) == 1 and len(_tbl.columns) == 1):
        continue

    _cell = _tbl.rows[0].cells[0]

    # Check if this is actually a highlight box (has a colored fill)
    _tcPr = _cell._tc.find(_qn('w:tcPr'))
    _is_highlight_box = False
    if _tcPr is not None:
        _shd = _tcPr.find(_qn('w:shd'))
        if _shd is not None:
            _fill = _shd.get(_qn('w:fill'))
            if _fill and _fill.upper() in ('2B5797', '1A63CD', _BRAND_NAVY):
                _is_highlight_box = True

    if not _is_highlight_box:
        continue

    # --- Collect existing text content (preserving paragraph breaks as bullet items) ---
    _lines = []
    for _para in _cell.paragraphs:
        _text = _para.text.strip()
        if _text and _text.lower() not in _HIGHLIGHT_TITLE_WORDS:
            _lines.append(_text)

    # --- Clear existing cell content ---
    for _para in list(_cell.paragraphs):
        _p_elem = _para._element
        _p_elem.getparent().remove(_p_elem)

    # --- Set cell background to light gray ---
    if _tcPr is None:
        _tcPr = OxmlElement('w:tcPr')
        _cell._tc.insert(0, _tcPr)
    _shd = _tcPr.find(_qn('w:shd'))
    if _shd is None:
        _shd = OxmlElement('w:shd')
        _tcPr.append(_shd)
    _shd.set(_qn('w:val'), 'clear')
    _shd.set(_qn('w:color'), 'auto')
    _shd.set(_qn('w:fill'), _OC_BG)

    # --- Set subtle border on cell ---
    _tcBorders = _tcPr.find(_qn('w:tcBorders'))
    if _tcBorders is None:
        _tcBorders = OxmlElement('w:tcBorders')
        _tcPr.append(_tcBorders)
    for _edge in ('top', 'left', 'bottom', 'right'):
        _bdr = OxmlElement(f'w:{_edge}')
        _bdr.set(_qn('w:val'), 'single')
        _bdr.set(_qn('w:sz'), '4')   # 0.5pt
        _bdr.set(_qn('w:space'), '0')
        _bdr.set(_qn('w:color'), _OC_BORDER)
        _existing = _tcBorders.find(_qn(f'w:{_edge}'))
        if _existing is not None:
            _tcBorders.remove(_existing)
        _tcBorders.append(_bdr)

    # --- Set cell padding (~1.25rem vertical, ~1.5rem horizontal) ---
    _tcMar = _tcPr.find(_qn('w:tcMar'))
    if _tcMar is None:
        _tcMar = OxmlElement('w:tcMar')
        _tcPr.append(_tcMar)
    for _side, _val in [('top', '240'), ('bottom', '240'), ('start', '300'), ('end', '300')]:
        _m = OxmlElement(f'w:{_side}')
        _m.set(_qn('w:w'), _val)
        _m.set(_qn('w:type'), 'dxa')
        _existing = _tcMar.find(_qn(f'w:{_side}'))
        if _existing is not None:
            _tcMar.remove(_existing)
        _tcMar.append(_m)

    # --- Add "HIGHLIGHTS" (or "DEFINITIONS") header paragraph ---
    _box_title = 'HIGHLIGHTS'
    if _lines and _lines[0].strip().lower() == 'definitions':
        _box_title = 'DEFINITIONS'
        _lines = _lines[1:]
    _header_p = OxmlElement('w:p')
    _cell._tc.append(_header_p)
    _header_para = _cell.paragraphs[-1]
    _header_run = _header_para.add_run(_box_title)
    _header_run.font.name = 'Calibri'
    _header_run.font.size = Pt(16)   # "HIGHLIGHTS" label size (was 8.5)
    _header_run.font.color.rgb = _OC_MUTED
    _header_run.font.bold = True
    _header_para.paragraph_format.space_after = Pt(8)
    _header_para.paragraph_format.space_before = Pt(0)
    # Add letter-spacing via character spacing
    _header_run.font._element.set(_qn('w:spacing'), '40')
    _rPr = _header_run._element.find(_qn('w:rPr'))
    if _rPr is None:
        _rPr = OxmlElement('w:rPr')
        _header_run._element.insert(0, _rPr)
    _spacing_elem = OxmlElement('w:spacing')
    _spacing_elem.set(_qn('w:val'), '40')
    _rPr.append(_spacing_elem)

    # --- Thin separator line below header ---
    _sep_p = OxmlElement('w:p')
    _cell._tc.append(_sep_p)
    _sep_para = _cell.paragraphs[-1]
    _sep_para.paragraph_format.space_before = Pt(0)
    _sep_para.paragraph_format.space_after = Pt(6)
    # Add bottom border to paragraph as separator
    _pPr = _sep_para._element.find(_qn('w:pPr'))
    if _pPr is None:
        _pPr = OxmlElement('w:pPr')
        _sep_para._element.insert(0, _pPr)
    _pBdr = OxmlElement('w:pBdr')
    _pPr.append(_pBdr)
    _bottom_bdr = OxmlElement('w:bottom')
    _bottom_bdr.set(_qn('w:val'), 'single')
    _bottom_bdr.set(_qn('w:sz'), '4')
    _bottom_bdr.set(_qn('w:space'), '1')
    _bottom_bdr.set(_qn('w:color'), 'DEE1E6')
    _pBdr.append(_bottom_bdr)

    # --- Add each line as a bullet paragraph ---
    for _li, _line in enumerate(_lines):
        _bp = OxmlElement('w:p')
        _cell._tc.append(_bp)
        _bullet_para = _cell.paragraphs[-1]
        _bullet_para.paragraph_format.space_before = Pt(5)
        _bullet_para.paragraph_format.space_after = Pt(5)
        _bullet_para.paragraph_format.line_spacing = 1.35
        # Hanging indent: bullet at 0, text indented 240 twips
        _bpPr = _bullet_para._element.find(_qn('w:pPr'))
        if _bpPr is None:
            _bpPr = OxmlElement('w:pPr')
            _bullet_para._element.insert(0, _bpPr)
        _ind = OxmlElement('w:ind')
        _ind.set(_qn('w:left'), '360')
        _ind.set(_qn('w:hanging'), '240')
        _bpPr.append(_ind)

        # Add bullet character
        _dot_run = _bullet_para.add_run(f'{_OC_BULLET_CHAR}  ')
        _dot_run.font.name = 'Calibri'
        _dot_run.font.size = Pt(10)
        _dot_run.font.color.rgb = _OC_STAT_BLUE
        _dot_run.font.bold = False

        # Split line into stat vs non-stat segments and style accordingly
        _parts = _stat_pattern.split(_line)
        for _part in _parts:
            if not _part:
                continue
            _is_stat = bool(_stat_pattern.fullmatch(_part))
            _run = _bullet_para.add_run(_part)
            _run.font.name = 'Calibri'
            _run.font.size = Pt(10.5)
            if _is_stat:
                _run.font.color.rgb = _OC_STAT_BLUE
                _run.font.bold = True
            else:
                _run.font.color.rgb = _OC_BODY_COLOR
                _run.font.bold = False

# --- Fix 12: Enforce Calibri in all table cells ---
for _tbl in doc.tables:
    for _row in _tbl.rows:
        for _cell in _row.cells:
            for _para in _cell.paragraphs:
                for _run in _para.runs:
                    _run.font.name = 'Calibri'

print("Post-render formatting applied: all fixes including typography, brand colors, borders, alignment")

##############################
### C24: a grade with no respondents gets no column. When no 8th graders took
### the survey, every "Grade 8" / "8th" column (school AND U.S. norm) is removed
### from every table and the table width is shared out again (the 8th-grade-only
### tables and their text are removed in the template with {%p if has_8th %}).
##############################
from docx.oxml import OxmlElement as _OxmlElement

def _tc_span(tc):
    _gs = tc.find(_qn('w:tcPr') + '/' + _qn('w:gridSpan'))
    return int(_gs.get(_qn('w:val'))) if _gs is not None else 1

def _set_tc_width(tc, twips):
    _pr = tc.find(_qn('w:tcPr'))
    if _pr is None:
        _pr = _OxmlElement('w:tcPr'); tc.insert(0, _pr)
    _w = _pr.find(_qn('w:tcW'))
    if _w is None:
        _w = _OxmlElement('w:tcW'); _pr.insert(0, _w)
    _w.set(_qn('w:w'), str(int(twips))); _w.set(_qn('w:type'), 'dxa')

def drop_grade_columns(tbl_el, header_re):
    """Remove the grid columns whose header cell matches header_re. Returns True if changed."""
    _rows = tbl_el.findall(_qn('w:tr'))
    _grid = tbl_el.find(_qn('w:tblGrid'))
    if not _rows or _grid is None:
        return False
    _cols = _grid.findall(_qn('w:gridCol'))
    _drop = set()
    for _hdr in _rows[:2]:               # one or two header rows
        _pos = 0
        for _tc in _hdr.findall(_qn('w:tc')):
            _span = _tc_span(_tc)
            _txt = ''.join(t.text or '' for t in _tc.iter(_qn('w:t'))).strip()
            if header_re.search(_txt):
                _drop.update(range(_pos, _pos + _span))
            _pos += _span
    if not _drop or len(_drop) >= len(_cols) - 1:
        return False
    _widths = [int(c.get(_qn('w:w')) or 0) for c in _cols]
    _total = sum(_widths)
    for _tr in _rows:
        _pos = 0
        for _tc in list(_tr.findall(_qn('w:tc'))):
            _span = _tc_span(_tc)
            _cells = set(range(_pos, _pos + _span))
            if _cells <= _drop:
                _tr.remove(_tc)
            elif _cells & _drop:          # merged cell straddling the dropped columns: shrink it
                _gs = _tc.find(_qn('w:tcPr') + '/' + _qn('w:gridSpan'))
                _gs.set(_qn('w:val'), str(_span - len(_cells & _drop)))
            _pos += _span
    for _i in sorted(_drop, reverse=True):
        _grid.remove(_cols[_i])
    # Share the width out again: label column unchanged, the rest equal (T8).
    _keep = [w for i, w in enumerate(_widths) if i not in _drop]
    _first = _keep[0]
    _rest = (_total - _first) / max(1, len(_keep) - 1)
    _new = [_first] + [_rest] * (len(_keep) - 1)
    for _gc, _w in zip(_grid.findall(_qn('w:gridCol')), _new):
        _gc.set(_qn('w:w'), str(int(_w)))
    for _tr in _rows:
        _pos = 0
        for _tc in _tr.findall(_qn('w:tc')):
            _span = _tc_span(_tc)
            _set_tc_width(_tc, sum(_new[_pos:_pos + _span]))
            _pos += _span
    return True

def _grade_header_re(n):
    return re.compile(rf'^(?:Grade {n}\b|{n}th(?:\s|$)|{n}th Grade|Perceived\s*{n}th|Actual\s*{n}th|{n}th Graders)', re.I)
for _absent_g in ['7th', '8th']:
    if _absent_g not in clean['grade_level'].values:
        _hre = _grade_header_re(_absent_g[:-2])
        _n_dropped = sum(drop_grade_columns(_t, _hre) for _t in list(doc.element.body.iter(_qn('w:tbl'))))
        if _n_dropped:
            print(f"C24: removed the {_absent_g}-grade columns from {_n_dropped} table(s) (no {_absent_g} graders took the survey).")

##############################
### T1, T3, T4, T6, T7: pagination rules enforced on the RENDERED report, so
### rows added by the template loops follow them too.
###  - every table: rows never split across pages, header row repeats, rows kept
###    together; the label above it is kept with it (T3)
###  - a note under a table ("*Annual use not measured...") stays with it (T7)
###  - a chart directly under a table stays with the table, but chart paragraphs
###    themselves are never "keep with next" (T4)
###  - a table longer than LONG_TABLE_ROWS rows starts on a new page (T6)
###  - empty spacer paragraphs are removed (T1)
##############################
LONG_TABLE_ROWS = 10

def _ppr(p_el):
    _pp = p_el.find(_qn('w:pPr'))
    if _pp is None:
        _pp = _OxmlElement('w:pPr'); p_el.insert(0, _pp)
    return _pp

def _set_flag(p_el, tag, on=True):
    _pp = _ppr(p_el)
    _x = _pp.find(_qn(tag))
    if on and _x is None:
        _x = _OxmlElement(tag)
        # keep schema order loosely: put flags right after pStyle
        _st = _pp.find(_qn('w:pStyle'))
        (_st.addnext(_x) if _st is not None else _pp.insert(0, _x))
    elif on:
        _x.set(_qn('w:val'), '1')
    elif _x is not None:
        _x.set(_qn('w:val'), '0')

def _has_drawing(el):
    return el.find('.//' + _qn('w:drawing')) is not None or el.find('.//' + _qn('w:pict')) is not None

def _ptext(el):
    return ''.join(t.text or '' for t in el.iter(_qn('w:t'))).strip()

def _is_spacer(el):
    if el.tag != _qn('w:p') or _ptext(el) or _has_drawing(el):
        return False
    if el.find('.//' + _qn('w:sectPr')) is not None or el.find('.//' + _qn('w:br')) is not None:
        return False
    if el.find('.//' + _qn('w:fldChar')) is not None or el.find('.//' + _qn('w:instrText')) is not None:
        return False
    if el.find('.//' + _qn('w:bookmarkStart')) is not None:
        return False
    return True

def _style_id(p_el):
    _st = p_el.find(_qn('w:pPr') + '/' + _qn('w:pStyle'))
    return _st.get(_qn('w:val')) if _st is not None else ''

_body = doc.element.body
# T1: spacer paragraphs in the body (not table cells) go; spacing comes from styles.
_spacers = [el for el in list(_body) if _is_spacer(el)]
for el in _spacers:
    _body.remove(el)
for _tbl in _body.iter(_qn('w:tbl')):
    _trs = _tbl.findall(_qn('w:tr'))
    for _i, _tr in enumerate(_trs):
        _trpr = _tr.find(_qn('w:trPr'))
        if _trpr is None:
            _trpr = _OxmlElement('w:trPr'); _tr.insert(0, _trpr)
        if _trpr.find(_qn('w:cantSplit')) is None:
            _trpr.append(_OxmlElement('w:cantSplit'))
        if _i == 0 and len(_trs) > 2 and _trpr.find(_qn('w:tblHeader')) is None:
            _trpr.append(_OxmlElement('w:tblHeader'))
        if _i < len(_trs) - 1:
            for _p in _tr.iter(_qn('w:p')):
                _set_flag(_p, 'w:keepNext')
    if _tbl.getparent() is not _body:
        continue          # nested (e.g. highlight boxes): only row rules
    # label paragraph(s) right above the table stay with it
    _prev = _tbl.getprevious()
    _label = None
    if _prev is not None and _prev.tag == _qn('w:p') and _ptext(_prev) and not _has_drawing(_prev):
        _set_flag(_prev, 'w:keepNext'); _label = _prev
        _pp2 = _prev.getprevious()
        if _pp2 is not None and _pp2.tag == _qn('w:p') and _ptext(_pp2) and len(_ptext(_pp2)) < 200 \
                and not _style_id(_pp2).lower().startswith('heading'):
            _set_flag(_pp2, 'w:keepNext')     # question text + table title
    _nxt = _tbl.getnext()
    _last = _trs[-1] if _trs else None
    if _nxt is not None and _nxt.tag == _qn('w:p') and _last is not None and \
            (_ptext(_nxt).startswith('*') or _has_drawing(_nxt)):
        for _p in _last.iter(_qn('w:p')):
            _set_flag(_p, 'w:keepNext')       # T7 notes / T3 chart under the table
        if _ptext(_nxt).startswith('*'):
            _n2 = _nxt.getnext()             # second note line ("**Identified as...")
            if _n2 is not None and _n2.tag == _qn('w:p') and _ptext(_n2).startswith('*'):
                _set_flag(_nxt, 'w:keepNext')
    # T6: long tables start on a new page (the break goes on the block's first line)
    # ...unless its section heading (Heading 1, which always starts a page) is
    # a few lines above with no other table in between: the approved layouts put
    # heading + bullets + table on one page (higher-risk subgroups p. A-C).
    _sec_start = False
    _w = _label.getprevious() if _label is not None else None
    for _k in range(10):
        if _w is None or _w.tag == _qn('w:tbl'):
            break
        _sid = _style_id(_w)
        if _sid and any(s.style_id == _sid and s.name.lower() == 'heading 1' for s in doc.styles if getattr(s, 'name', None)):
            _sec_start = True
            break
        _w = _w.getprevious()
    if len(_trs) > LONG_TABLE_ROWS + 1 and _label is not None and not _sec_start:
        _start = _label
        _p0 = _label.getprevious()
        while _p0 is not None and _p0.tag == _qn('w:p') and _ptext(_p0) and \
                (((_style_id(_p0).lower().startswith('heading') or len(_ptext(_p0)) < 200) and
                  _p0.find(_qn('w:pPr') + '/' + _qn('w:keepNext')) is not None)
                 or _ptext(_p0).startswith('The following table')):
            if _ptext(_p0).startswith('The following table'):
                _set_flag(_p0, 'w:keepNext')
            _start = _p0; _p0 = _p0.getprevious()
        _set_flag(_start, 'w:pageBreakBefore')
# No title left alone at the bottom of a page: a page break on a paragraph
# moves up to the short title/label lines directly above it (e.g. when the
# 8th-grade gender table is not printed, the 10th-grade table's break moves up
# to the "Substance Use by Gender" title). A Heading 1 above already starts a
# page, so the extra break is dropped.
_H1_IDS = {s.style_id for s in doc.styles if getattr(s, 'name', '') and s.name.lower() == 'heading 1'}
_TITLE_STYLES = {'SUPPTableLabel', 'SUPPSubheading', 'SUPPComparisonTitle'}
for _p in list(_body.findall(_qn('w:p'))):
    _pb = _p.find(_qn('w:pPr') + '/' + _qn('w:pageBreakBefore'))
    if _pb is None or _pb.get(_qn('w:val')) in ('0', 'false'):
        continue
    _top = _p
    _q = _p.getprevious()
    while _q is not None and _q.tag == _qn('w:p') and _ptext(_q) and \
            (len(_ptext(_q)) < 200 or _ptext(_q).startswith('The following table')) and \
            (_style_id(_q) in _TITLE_STYLES or _style_id(_q) in _H1_IDS or
             _style_id(_q).lower().startswith('heading') or _ptext(_q).startswith('The following table')):
        _top = _q
        if _style_id(_q) in _H1_IDS or _q.find(_qn('w:pPr') + '/' + _qn('w:pageBreakBefore')) is not None:
            break                         # this line already starts the page
        _q = _q.getprevious()
    if _top is not _p:
        _pb.getparent().remove(_pb)
        if _style_id(_top) not in _H1_IDS:
            _set_flag(_top, 'w:pageBreakBefore')
        _q = _top                         # keep the moved lines with the table
        while _q is not None and _q is not _p:
            _set_flag(_q, 'w:keepNext'); _q = _q.getnext()
# T4: chart paragraphs are never "keep with next"
for _p in _body.iter(_qn('w:p')):
    if _has_drawing(_p):
        _kn = _p.find(_qn('w:pPr') + '/' + _qn('w:keepNext'))
        if _kn is not None:
            _kn.getparent().remove(_kn)
print(f"T1/T3/T6: removed {len(_spacers)} spacer paragraph(s); table pagination rules applied.")

##############################
### B5 + B7: post-render QA lint.
### 1) Auto-corrects direction words that contradict the numbers in the same
###    sentence ("rises from 100.0% ... to 80.0%").
### 2) Flags (and by default stops the run on) leftover template text, author
###    notes, doubled tokens, broken punctuation, "majority" under 50%,
###    "more than double" that is not, duplicated bullets, second-person
###    wording in third-person bullets.
##############################
def _all_report_paragraphs(_doc):
    for _p in _doc.paragraphs:
        yield _p
    for _t in _doc.tables:
        for _r in _t.rows:
            for _c in _r.cells:
                for _p in _c.paragraphs:
                    yield _p

def _set_para_text(_p, _new):
    if _p.runs:
        _p.runs[0].text = _new
        for _r in _p.runs[1:]:
            _r.text = ''
    else:
        _p.add_run(_new)

_dir_re = re.compile(r'\b(rises|rise|increases|increase)\s+from\s+([\d.]+)%\s+(?:in|of)\s+(\w+)\s+grade(?:rs)?\s+to\s+([\d.]+)%')
LINT_ERRORS, LINT_WARNINGS = [], []
_recent_paras = []
_LINT_PATTERNS = [
    (r'\{\{|\}\}|\{%|%\}', 'unfilled template placeholder'),
    (r'(?i)do these manually|make these manually', 'author note left in report'),
    (r'\b\d+\s+\d+(\.\d+)?%\s+students\b', 'doubled number token (e.g. "14 5.9% students")'),
    (r'is significantly exceeds|exceeds with U\.S\.|\bis exceeds\b', 'broken comparison wording'),
    (r'measures\.\s+and\b|measures\.,|\.\.(?!\.)', 'broken punctuation'),
    (r'(?i)\bnan\b|N/A%|\bNone%', 'missing value printed'),
    (r'%\s*%', 'double percent sign'),
]
for _p in _all_report_paragraphs(doc):
    _t = _p.text
    if not _t.strip():
        continue
    # 1) direction words
    _m = _dir_re.search(_t)
    if _m:
        _a, _b = float(_m.group(2)), float(_m.group(4))
        if _b < _a:
            _verb = {'rises': 'falls', 'rise': 'fall', 'increases': 'decreases', 'increase': 'decrease'}[_m.group(1)]
            _new = _t[:_m.start(1)] + _verb + _t[_m.end(1):]
            _set_para_text(_p, _new)
            LINT_WARNINGS.append(f'auto-corrected "{_m.group(1)}" to "{_verb}": {_new[:140]}')
            _t = _new
    # 2) lint checks
    for _pat, _why in _LINT_PATTERNS:
        if re.search(_pat, _t):
            LINT_ERRORS.append(f'{_why}: "{_t[:160]}"')
    for _mm in re.finditer(r'(?i)\bmajority\s*\(?\s*([\d.]+)%', _t):
        if float(_mm.group(1)) <= 50:
            LINT_ERRORS.append(f'"majority" used for {_mm.group(1)}%: "{_t[:160]}"')
    _md = re.search(r'more than double.*?\((\d+)\).*?\((\d+)\)', _t)
    if _md and int(_md.group(1)) <= 2 * int(_md.group(2)):
        LINT_ERRORS.append(f'"more than double" is not true: "{_t[:160]}"')
    if re.match(r'^\s*[\d.]+%\s+of\b', _t) and re.search(r"\byou(r)?\b", _t):
        LINT_WARNINGS.append(f'second-person wording in a third-person bullet: "{_t[:160]}"')
    # duplicated bullets close together (the Exec Summary legitimately repeats
    # some sentences later in the report, so only nearby repeats count)
    _key = _t.strip()
    if len(_key) > 25 and re.search(r'\d', _key) and not _key.startswith(('Note:', '*')) and _key in _recent_paras:
        LINT_WARNINGS.append(f'possible duplicated bullet (check it is two different questions): "{_t[:160]}"')
    _recent_paras.append(_key)
    if len(_recent_paras) > 6:
        _recent_paras.pop(0)

##############################
### Part 4: checks run on every generated report (Nov 2026 fix list). Problems are
### lint ERRORS, so the run stops with "BUILD CHECK FAILED" after saving.
##############################
PART4 = []
def _p4(num, ok, msg):
    PART4.append(f"{num}. {'PASS' if ok else 'FAIL'}: {msg}")
    if not ok:
        LINT_ERRORS.append(f'Part 4 check {num}: {msg}')

# 1. norms file complete (checked at startup; the run stops there if not)
_p4(1, not validate_mtf(MTF, MTFyear), 'norms file has Total, Male and Female for every substance x timeline')
# 2. exclusion log + totals match Appendix A
_appA_grade = int(pd.to_numeric(gradedf['Count (N)'], errors='coerce').sum())
_appA_sex = int(pd.to_numeric(sexdf['Count (N)'], errors='coerce').sum())
_p4(2, len(EXCLUSION_LOG) == removed_N and N - removed_N == N_analyzed == _appA_grade == _appA_sex,
    f'exclusion log has {len(EXCLUSION_LOG)} rows for {removed_N} exclusions; {N} collected - {removed_N} = '
    f'{N_analyzed} analyzed; Appendix A totals by grade {_appA_grade}, by sex {_appA_sex}')
# 3. the same statistic has the same value in the sentence and in its table
_c3 = []
for _lbl, _tbl, _series in [('social atmosphere', socialatmo, _sa), ('pressured', pressured_table, _pr)]:
    _row = _tbl[_tbl['index'] == 'Total Agree']
    for _g, _v in _series:
        if f'{_g} (N)' in _row.columns and _v != '':
            _tv = round(float(_row[f'{_g} (N)'].iloc[0]), 1)
            if abs(_tv - float(_v)) > 0.05:
                _c3.append(f'{_lbl} {_g}: sentence {float(_v)} vs table {_tv}')
_recount_double = int(((_cs['High-risk (%)'] >= 2 * _cs['Low-risk (%)']) & (_cs['High-risk (%)'] > 0)).sum())
if _recount_double != len_highrisk_twice_risk:
    _c3.append(f'"double or greater": sentence {len_highrisk_twice_risk} vs table {_recount_double}')
if sum(highrisk_by_grade.values()) != totalhighrisk:
    _c3.append(f'higher-risk count {totalhighrisk} vs grade breakdown {sum(highrisk_by_grade.values())}')
_ov = float(QA_CONTEXT['social_atmo_df_yes_perc']) if QA_CONTEXT['social_atmo_df_yes_perc'] != '' else None
if _ov is not None and abs(_ov - float(social_atmo_df_yes_perc)) > 0.05:
    _c3.append(f'social atmosphere overall: {_ov} vs {social_atmo_df_yes_perc}')
_p4(3, not _c3, 'numbers in sentences match their tables' + (': ' + '; '.join(_c3) if _c3 else ''))
# 4. no sentence names a grade with no respondents
_absent = [g for g in ['7th', '8th', '9th', '10th', '11th', '12th'] if g not in _grades_seq]
_grade_word_re = re.compile(r'\b(' + '|'.join(_absent) + r')[ -](?:grade|graders)\b') if _absent else None
_c4 = []
if _grade_word_re:
    for _pel in _body.findall(_qn('w:p')):
        _t = _ptext(_pel)
        if _grade_word_re.search(_t) and 'Seniors (12th graders)' not in _t:
            _c4.append(_t[:120])
_p4(4, not _c4, f'no sentence names a grade with zero respondents ({", ".join(_absent) or "none absent"})'
    + (': ' + ' | '.join(_c4[:5]) if _c4 else ''))
# 5. percentages with a base under SHOW_N_BELOW carry their n
_left = [p.text[:80] for p in _iter_all_paragraphs(doc) if '⁣' in p.text]
_small = sorted({k for k, v in RESULTS.items() if v['d'] and v['d'] < SHOW_N_BELOW})
_p4(5, not _left, f'small-base percentages printed with their n ({len(_small)} small-base statistics computed)')
for _k in _small[:40]:
    LINT_WARNINGS.append(f'small base (shown with n where printed): {_k} = {RESULTS[_k]["value"]}% '
                         f'({int(RESULTS[_k]["n"])} of {int(RESULTS[_k]["d"])})')
# 6. "Higher-risk" / "Lower-risk" everywhere (charts are labelled by GROUP_DISPLAY)
_c6 = [p.text[:100] for p in _iter_all_paragraphs(doc) if re.search(r'\b(High|Low)-risk\b', p.text)]
_p4(6, not _c6, '"Higher-risk" / "Lower-risk" used everywhere, including chart labels'
    + (': ' + ' | '.join(_c6[:5]) if _c6 else ''))
# 7. no empty paragraphs in the body (tables kept whole by the rules above)
_c7 = sum(1 for el in _body if _is_spacer(el))
_p4(7, _c7 == 0, f'no empty paragraphs in the body ({_c7} found); table rows cannot split; labels kept with tables')
# 8. TOC: levels 1-2 only, refreshed when the file is opened (after pagination is final)
_toc_instr = ' '.join(t.text or '' for t in _body.iter(_qn('w:instrText')) if 'TOC' in (t.text or ''))
_p4(8, (not _toc_instr) or ('1-2' in _toc_instr),
    f'table of contents limited to heading levels 1-2 ({_toc_instr.strip()[:40] or "no TOC field"})')
# 9. organization name
_c9 = [p.text[:100] for p in _iter_all_paragraphs(doc) if re.search(r'\bFCD\b|Prevention Education Solutions', p.text)]
_p4(9, not _c9, 'organization name is "Prevention Ed" throughout' + (': ' + ' | '.join(_c9[:3]) if _c9 else ''))
# 10. both dates printed and labelled
_all_text = '\n'.join(p.text for p in _iter_all_paragraphs(doc))
_p4(10, ('Survey date' in _all_text and str(date_survey) in _all_text and 'Report date' in _all_text
         and REPORT_DATE in _all_text),
    f'survey date ({date_survey}) and report date ({REPORT_DATE}) both printed with labels')
print('PART 4 CHECKS\n  ' + '\n  '.join(PART4))
# Paragraphs shaded green in the template need a manual review in every report:
# list them so they are not missed.
_green = []
for _pel in _body.iter(_qn('w:p')):
    if any((_s.get(_qn('w:fill')) or '').lower() == '00ff00' for _s in _pel.iter(_qn('w:shd'))) and _ptext(_pel).strip():
        _green.append(_ptext(_pel).strip())
for _g in _green:
    LINT_WARNINGS.append(f'MANUAL REVIEW (green text): "{_g[:150]}"')
if _green:
    print(f"{len(_green)} green paragraph(s) need a manual review (listed in _QA_lint.txt).")

# Word refreshes the table of contents (and page numbers) when the report is
# opened, i.e. after pagination is final (T16).
_settings = doc.settings.element
if _settings.find(_qn('w:updateFields')) is None:
    _uf = _OxmlElement('w:updateFields'); _uf.set(_qn('w:val'), 'true')
    _settings.append(_uf)

##############################
### Year groups: when the survey used British year groups (Year 9 to Year 13,
### mapped to 8th-12th for the analysis), every grade mention in the finished
### report is written as a year group, so the report never mixes "10th grade",
### "Year 11" and half-replaced labels such as "10th-Year 13". The protected
### gambling section is left exactly as written.
##############################
YEAR_GROUP_PROTECTED_TITLES = ['Gambling and In-Game Purchasing']

def _yr(n):
    n = int(n)
    return f"Year {n + 1}" if 8 <= n <= 12 else f"{n}th grade"

_YEAR_RULES = [
    # repair labels half-converted earlier ("12th grade" + "r" -> "Year 13r")
    (re.compile(r'\b(Year \d{1,2})rs\b'), lambda m: m.group(1) + ' students'),
    (re.compile(r'\b(Year \d{1,2})r\b'), lambda m: m.group(1) + ' student'),
    # ranges first
    (re.compile(r'\b(\d{1,2})th\s*-\s*Year (\d{1,2})\b'), lambda m: f"Year {int(m.group(1)) + 1}-{m.group(2)}"),
    (re.compile(r'\b(\d{1,2})th\s*-\s*(\d{1,2})th[ -](?:grade students|graders|grade)\b', re.I),
     lambda m: f"Year {int(m.group(1)) + 1}-{int(m.group(2)) + 1} students"),
    (re.compile(r'\bgrades (\d{1,2})\s*-\s*(\d{1,2})\b', re.I),
     lambda m: f"Years {int(m.group(1)) + 1}-{int(m.group(2)) + 1}"),
    (re.compile(r'\bgrades ((?:\d{1,2}, )*\d{1,2}(?:,)? and \d{1,2})\b', re.I),
     lambda m: 'Years ' + re.sub(r'\d{1,2}', lambda d: str(int(d.group(0)) + 1), m.group(1))),
    # lists: "10th and 12th Graders", "8th, 10th, and 12th grade students"
    (re.compile(r'\b((?:\d{1,2}th(?:,\s*and\s+|,\s*|\s+and\s+))+\d{1,2}th)\s+(Graders|graders|grade students|Grade Students)\b'),
     lambda m: re.sub(r'(\d{1,2})th', lambda d: _yr(d.group(1)), m.group(1)) + (' Students' if m.group(2)[0].isupper() else ' students')),
    # single grades
    (re.compile(r'\b(\d{1,2})th[ -](?:grade students|graders)\b', re.I), lambda m: f"{_yr(m.group(1))} students"),
    (re.compile(r'\b(\d{1,2})th[ -]grader\b', re.I), lambda m: f"{_yr(m.group(1))} student"),
    (re.compile(r'\bin (\d{1,2})th grade\b', re.I), lambda m: f"in {_yr(m.group(1))}"),
    (re.compile(r'\b(\d{1,2})th[ -]grade\b', re.I), lambda m: _yr(m.group(1))),
    (re.compile(r'\bGrade (\d{1,2})\b'), lambda m: f"Year {int(m.group(1)) + 1}" if 8 <= int(m.group(1)) <= 12 else m.group(0)),
    # table headers that are just "10th", "10th Males", "10th MTF"
    (re.compile(r'^(\d{1,2})th(?=\s|$)'), lambda m: f"Year {int(m.group(1)) + 1}" if 8 <= int(m.group(1)) <= 12 else m.group(0)),
]

def _to_year_groups(text):
    for rx, fn in _YEAR_RULES:
        text = rx.sub(fn, text)
    return text

if '_was_year_system' in clean.columns:
    _skip = set()
    _els = list(_body)
    for _title in YEAR_GROUP_PROTECTED_TITLES:
        _i = next((k for k, e in enumerate(_els) if e.tag == _qn('w:p') and _ptext(e).strip() == _title), None)
        if _i is None:
            continue
        _st = _style_id(_els[_i])
        for e in _els[_i:]:
            if e is not _els[_i] and e.tag == _qn('w:p') and _ptext(e).strip() and \
                    _style_id(e).lower().startswith('heading') and _style_id(e) != '':
                break
            _skip.add(e)
    _n_year = 0
    for _top in _els:
        if _top in _skip:
            continue
        for _pel in ([_top] if _top.tag == _qn('w:p') else list(_top.iter(_qn('w:p')))):
            _ts = [t for t in _pel.iter(_qn('w:t'))]
            _full = ''.join(t.text or '' for t in _ts)
            if not re.search(r'\d{1,2}th|Grade \d|grades \d|Year \d{1,2}rs?\b', _full, re.I):
                continue
            _target = _to_year_groups(_full)
            if _target == _full:
                continue
            # run by run keeps bold/italic words; only if that gives the same
            # sentence as converting the whole paragraph (a label split over
            # two runs, e.g. "10th grade" + "r", is converted as a whole)
            _per_run = [_to_year_groups(t.text or '') for t in _ts]
            if ''.join(_per_run) == _target:
                for t, v in zip(_ts, _per_run):
                    t.text = v
            else:
                _ts[0].text = _target
                for t in _ts[1:]:
                    t.text = ''
            _n_year += 1
    print(f"Year groups: {_n_year} grade label(s) written as year groups.")

# Save the QA side files next to the report
_qa_base = os.path.splitext(savepath)[0]
try:
    EXCLUSION_LOG.to_csv(_qa_base + '_exclusion_log.csv', index=False)   # C2
    print(f"Exclusion log saved: {_qa_base}_exclusion_log.csv ({len(EXCLUSION_LOG)} responses)")
except Exception as _ex_err:
    LINT_ERRORS.append(f'Part 4 check 2: could not write the exclusion log: {_ex_err}')
try:
    pd.DataFrame(METRIC_LOG).to_csv(_qa_base + '_QA_metrics.csv', index=False)
    _flags_df = pd.DataFrame(QA_FLAGS)
    if len(INVALID_RESPONDENTS):
        _flags_df = pd.concat([_flags_df, pd.DataFrame({'row': INVALID_RESPONDENTS.index, 'grade': INVALID_RESPONDENTS['grade_level'],
                                                        'sex': INVALID_RESPONDENTS['sex'],
                                                        'issue': INVALID_RESPONDENTS['reason'] + np.where(INVALID_RESPONDENTS['excluded'], ' (EXCLUDED)', ' (kept)')})])
    _flags_df.to_csv(_qa_base + '_QA_flags.csv', index=False)
    with open(_qa_base + '_QA_lint.txt', 'w') as _f:
        _f.write('PART 4 CHECKS\n' + '\n'.join(PART4) + '\n\nERRORS\n' + '\n'.join(LINT_ERRORS or ['none']) + '\n\nWARNINGS / AUTO-FIXES\n' + '\n'.join(LINT_WARNINGS or ['none']) + '\n')
    print(f"QA files saved: {_qa_base}_QA_metrics.csv, _QA_flags.csv, _QA_lint.txt")
except Exception as _qa_err:
    print(f"Could not write QA side files: {_qa_err}")
for _w in LINT_WARNINGS:
    logger.warning('LINT: ' + _w)
for _e in LINT_ERRORS:
    logger.error('LINT: ' + _e)

try:
    doc.save(savepath)
    print(f"SUCCESS: Report saved to {savepath}")
except Exception as _save_err:
    import traceback
    print("ERROR during doc.save():")
    traceback.print_exc()
    raise

# Build record + copy of this script, so later trend reports reproduce these numbers exactly
try:
    import json as _json, shutil as _shutil, hashlib as _hashlib
    _script_copy = _qa_base + '_SUPP_script_used.py'
    _shutil.copyfile(os.path.abspath(__file__), _script_copy)
    with open(surveryoutput, 'rb') as _f:
        _csv_hash = _hashlib.sha256(_f.read()).hexdigest()
    BUILD_RECORD.update({
        'school': School, 'survey_date': date_survey, 'report': os.path.abspath(savepath),
        'survey_csv': os.path.abspath(surveryoutput), 'survey_csv_sha256': _csv_hash,
        'script_used': os.path.abspath(_script_copy), 'grade_system_answer': response_grade_level,
        'built': datetime.datetime.now().isoformat(timespec='minutes'),
    })
    with open(_qa_base + '_build_record.json', 'w') as _f:
        _json.dump(BUILD_RECORD, _f, indent=2)
    print(f"Build record saved: {_qa_base}_build_record.json (keep it with the report)")
except Exception as _rec_err:
    print(f"Could not write the build record: {_rec_err}")

if LINT_ERRORS and FAIL_ON_LINT_ERRORS:
    raise SystemExit(f"B7 BUILD CHECK FAILED: {len(LINT_ERRORS)} problem(s) in the rendered report "
                     f"(listed above and in {_qa_base}_QA_lint.txt). The file was saved so you can inspect it, "
                     f"but fix the template/code before sending it to the school.")


# In[78]:


########## K charts say value and variable in the x and y axis, fix that


# In[79]:


#print(f'The file is done and saved at {savepath}. Woohoo!')


# In[80]:


# import winsound
# frequency = 2500  # Set Frequency To 2500 Hertz
# duration = 200  # Set Duration To 1000 ms == 1 second
# winsound.Beep(frequency, duration)


# In[81]:


AgreeDisagree2_answers_grade[4]


# In[82]:


aca_table_chart 


# In[83]:


bar_chart_it_up_2(aca_table, os.path.join(chart_dir, 'aca_table_chart.png'))


# In[84]:


bar_chart_it_up_2(valued_table, os.path.join(chart_dir, 'valued_table_chart.png'))


# In[85]:


often_one_grade


# In[ ]:




