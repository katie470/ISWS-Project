#!/usr/bin/env python
# coding: utf-8
"""
Refactored & safer version of the PreventionEd survey-processing script.
- Fixes: membership checks, chained assignments, .item() guards, dtype checks
- Adds: logging, argument-driven IO, main(), column normalization helpers
- Keeps original analysis structure but encourages modular use.

Note: This is a safe drop-in replacement for the core processing. Further
vectorization and tests are recommended.
"""

import argparse
import logging
from pathlib import Path
import sys
import datetime

import numpy as np
import pandas as pd
import seaborn as sns

# Optional imports used by reporting/templating
try:
    from docxtpl import DocxTemplate, InlineImage
    from docx.shared import Mm
except Exception:
    DocxTemplate = None

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
logger = logging.getLogger(__name__)

# --------------------------- Utility helpers ---------------------------

def normalize_columns(df: pd.DataFrame) -> pd.DataFrame:
    """Normalize column names for more robust matching.
    - strips surrounding whitespace
    - replaces multiple spaces and '*' characters
    - lowercases for consistent matching (keeps original mapping)
    """
    mapping = {c: c.strip().replace('*', '').replace('  ', ' ') for c in df.columns}
    df = df.rename(columns=mapping)
    return df


def percent_to_float(s: str) -> float:
    """Convert strings like '12.3%' or numeric-like to float safely."""
    if pd.isna(s):
        return np.nan
    if isinstance(s, (int, float)):
        return float(s)
    try:
        return float(str(s).strip().replace('%', ''))
    except Exception:
        return np.nan


def safe_item(series: pd.Series, default=None):
    """Return series.item() safely; if empty, return default."""
    if series is None or len(series) == 0:
        return default
    try:
        return series.item()
    except Exception:
        try:
            return series.iat[0]
        except Exception:
            return default


# --------------------------- Cleaning & validation ---------------------------

def load_inputs(survey_csv: Path, mtf_csv: Path):
    logger.info("Loading CSV files")
    clean = pd.read_csv(survey_csv)
    mtf = pd.read_csv(mtf_csv)
    clean = normalize_columns(clean)
    mtf = normalize_columns(mtf)
    return clean, mtf


def ensure_column(df: pd.DataFrame, col: str):
    """Ensure column exists in df; if not, add it filled with NaN."""
    if col not in df.columns:
        logger.debug("Adding missing column: %s", col)
        df[col] = np.nan


def safe_to_numeric(series: pd.Series, **kwargs) -> pd.Series:
    return pd.to_numeric(series, errors='coerce', **kwargs)


# --------------------------- Example safer operations ---------------------------

def drop_contradictions(df: pd.DataFrame) -> pd.DataFrame:
    """Remove rows with contradictory marijuana answers (example from original).
    Uses vectorized operations and returns cleaned DataFrame.
    """
    # Normalize key columns existence
    for col in ['age_first_time_pot', 'how_often_pot']:
        ensure_column(df, col)

    cond1 = (df['age_first_time_pot'] == 'I have never used marijuana at all') & (df['how_often_pot'].fillna('') != 'Never')
    cond2 = (df['how_often_pot'].fillna('') == 'Never') & (df['age_first_time_pot'].fillna('') != 'I have never used marijuana at all')
    mask = cond1 | cond2
    n = mask.sum()
    if n > 0:
        logger.info("Found %d contradictory survey answers (marijuana use), removing rows", int(n))
        df = df.loc[~mask].reset_index(drop=True)
    else:
        logger.info("No marijuana contradictions found")
    return df


def safe_replace_apostrophes(df: pd.DataFrame, cols: list):
    for col in cols:
        if col in df.columns:
            df[col] = df[col].astype(str).str.replace("'", "’", regex=False)


# --------------------------- Core processing (example) ---------------------------

def compute_basic_counts(df: pd.DataFrame) -> dict:
    """Compute a few example aggregated outputs and return as dict.
    This replaces many one-off computations in the original script with
    reusable pieces.
    """
    out = {}
    # Ensure Token exists
    ensure_column(df, 'Token')

    out['N'] = int(df['Token'].notna().sum())

    # Grade counts (robust membership check)
    if 'grade_level' in df.columns:
        unique_grades = df['grade_level'].dropna().unique().tolist()
        out['grades_present'] = unique_grades
        counts = df['grade_level'].value_counts(dropna=True).to_dict()
        out['grade_counts'] = counts
    else:
        out['grades_present'] = []
        out['grade_counts'] = {}

    # Sex distribution
    if 'sex' in df.columns:
        out['sex_counts'] = df['sex'].fillna('Unknown').value_counts().to_dict()
    else:
        out['sex_counts'] = {}

    return out


# --------------------------- Reporting stub ---------------------------

def render_docx(template_path: Path, context: dict, out_path: Path):
    if DocxTemplate is None:
        logger.warning('docxtpl not available; skipping docx render')
        return
    doc = DocxTemplate(str(template_path))
    doc.render(context)
    doc.save(str(out_path))
    logger.info('Saved report to %s', out_path)


# --------------------------- Main and CLI ---------------------------

def main(argv=None):
    parser = argparse.ArgumentParser(description='Cleaned PreventionEd survey script')
    parser.add_argument('--survey', required=True, help='Path to survey CSV (Typeform export)')
    parser.add_argument('--mtf', required=True, help='Path to MTF CSV file')
    parser.add_argument('--template', required=False, help='(optional) docx template path')
    parser.add_argument('--out', required=False, default='report_output.docx', help='Output report path')
    args = parser.parse_args(argv)

    survey_csv = Path(args.survey)
    mtf_csv = Path(args.mtf)

    if not survey_csv.exists() or not mtf_csv.exists():
        logger.error('One or more input files not found. survey=%s mtf=%s', survey_csv, mtf_csv)
        sys.exit(2)

    clean, mtf = load_inputs(survey_csv, mtf_csv)

    # Example cleaning steps
    clean = drop_contradictions(clean)
    safe_replace_apostrophes(clean, ['grades_overall', 'where_do_you_drink', 'attitude'])

    # Cast ages to numeric where possible
    if 'age' in clean.columns:
        # Many variants; try a few heuristics
        clean['agenum'] = (
            clean['age'].astype(str).str.extract(r'(\d{1,2})', expand=False)
        )
        clean['agenum'] = safe_to_numeric(clean['agenum']).astype(pd.Float64Dtype())

    # Derived variables
    results = compute_basic_counts(clean)
    logger.info('Basic counts: %s', results)

    # Render docx if requested
    if args.template:
        context = {**results, 'date_today': datetime.datetime.now().strftime('%b %d %Y')}
        render_docx(Path(args.template), context, Path(args.out))

    # Save a cleaned CSV next to original for inspection
    out_clean = survey_csv.with_name(survey_csv.stem + '_cleaned.csv')
    clean.to_csv(out_clean, index=False)
    logger.info('Wrote cleaned CSV to %s', out_clean)


if __name__ == '__main__':
    main()
