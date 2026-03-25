#!/usr/bin/env python3
"""
Pre-process SUPP survey CSV + MTF CSV into a JSON blob for the dashboard.
Usage:  python3 supp_preprocess.py > dashboard_data.json
"""

import pandas as pd
import numpy as np
import json, sys, os

# ── paths ──────────────────────────────────────────────────────────────
SURVEY = "/Users/katiegreeley/Desktop/2026 SUPP files/ISM SUPP Survey Fall 2025 Data  - ISM Prevention Ed SUPP Fall 2025.csv"
MTF    = "/Users/katiegreeley/Desktop/2026 SUPP files/MTF2024.csv"

# ── column position map (0-indexed from raw CSV header row) ───────────
# Demographics
C_SEX   = 0
C_AGE   = 1
C_GRADE = 2
C_GPA   = 3

# Alcohol onset / use
C_FIRST_DRINK_AGE = 4
C_FIRST_DRINK_LOC = 5
C_DRINK_FREQ      = 6
C_DRINKS_PER      = 7
C_PAST12_ALCOHOL  = 8

# Consequences of USE (9-31) — 23 items
CONS_START, CONS_END = 9, 31  # inclusive

# Past-30-day alcohol detail
C_30D_DAYS     = 32
C_30D_MAX      = 33
C_30D_BINGE    = 34
C_30D_DRUNK    = 35
C_DRINK_LOC    = 36

# Marijuana
C_MJ_AGE  = 37
C_MJ_FREQ = 38

# Lifetime substance use (cols 39-55 = 17 substances)
LIFE_START, LIFE_END = 39, 55
# Past-30-day substance use (cols 56-72 = same 17 substances)
P30_START, P30_END   = 56, 72

SUBSTANCE_NAMES = [
    "Alcohol", "Cigarettes", "Vaping/E-cigs", "Smokeless Tobacco/ZYN",
    "Marijuana", "Ecstasy/MDMA", "Heroin", "Narcotics (non-Rx)",
    "Amphetamines (non-Rx)", "Sedatives (non-Rx)", "Methamphetamine",
    "Cocaine", "LSD", "Hallucinogens (other)", "Inhalants",
    "Steroids (non-Rx)", "OTC Cold Medicine"
]

# Secondhand effects (73-85) — 13 items
SEC_START, SEC_END = 73, 85

# School event use
C_SCHOOL_EVENT = 86

# Attitude toward teen drinking
C_ATTITUDE = 87

# Peer disapproval items (88-91), risk items (92-95), pot-friend (96-98)
# School climate (99-107) — 9 Likert items
CLIMATE_START, CLIMATE_END = 99, 107
CLIMATE_LABELS = [
    "Social atmosphere encourages alcohol use",
    "Students feel pressured to drink",
    "Easy to make friends without drinking",
    "Students assume everyone drinks",
    "Students let you make own decisions",
    "Students respect non-drinkers",
    "Abstainers don't fit in",
    "Most students are light/non-drinkers",
    "Heavy/binge drinking is okay"
]

# School belonging (127-133) — 7 items
BELONG_START, BELONG_END = 127, 133
BELONG_LABELS = [
    "I feel valued as a person",
    "Academics are important to me",
    "Teachers/staff care about me",
    "School addressing substance problems",
    "Most students care about school",
    "Academics important to others",
    "I have a trusted adult at school"
]

# Family (135-140) — 6 Likert items
FAMILY_START, FAMILY_END = 135, 140
FAMILY_LABELS = [
    "Parents show interest in schoolwork",
    "Clear family rules",
    "Parent wants call if late",
    "Parent knows if not home on time",
    "Clear rules about substances",
    "Parent can contact me in emergency"
]

# Home use with/without parent
C_USE_WITH_PARENT    = 141
C_USE_WITHOUT_PARENT = 142
C_FAMILY_MEALS       = 143

# Stress items (149-155)
STRESS_START, STRESS_END = 149, 155
STRESS_LABELS = [
    "Academic", "Parental pressure", "College expectations",
    "Being overlooked", "Own expectations", "Boredom", "Fitting in socially"
]

# Coping substances (156-172) — same 17 substances
COPE_START, COPE_END = 156, 172

# Kessler K10 (173-182) — 10 items
K10_START, K10_END = 173, 182
K10_LABELS = [
    "Tired for no reason", "Nervous", "Nothing could calm you",
    "Hopeless", "Restless/fidgety", "Couldn't sit still",
    "Depressed", "Everything was an effort",
    "So sad nothing could cheer up", "Worthless"
]

K10_MAP = {
    "None of the time": 1, "A little of the time": 2,
    "Some of the time": 3, "Most of the time": 4, "All of the time": 5
}

# ── helpers ────────────────────────────────────────────────────────────
def pct(num, denom):
    """Return rounded percentage or 0."""
    if denom == 0:
        return 0.0
    return round(num / denom * 100, 1)

def likert_agree_pct(series):
    """% answering Agree or Strongly agree (but NOT Disagree)."""
    s = series.dropna().astype(str).str.strip()
    n = len(s)
    if n == 0:
        return 0.0
    # Match 'Agree' or 'Strongly agree' but exclude 'Disagree'/'Strongly disagree'
    agree = s.isin(["Agree", "Strongly agree"]).sum()
    return round(agree / n * 100, 1)

def yes_pct(series):
    """% answering Yes or TRUE (case-insensitive)."""
    s = series.dropna().astype(str).str.strip().str.lower()
    n = len(s)
    if n == 0:
        return 0.0
    y = s.isin(["yes", "true"]).sum()
    return round(y / n * 100, 1)

def consequence_pct(series):
    """% who experienced at least once (not 'Never' and not blank)."""
    s = series.dropna().astype(str).str.strip()
    s = s[s != ""]
    n = len(s)
    if n == 0:
        return 0.0
    experienced = s[~s.str.contains("Never|never", regex=True)]
    return round(len(experienced) / n * 100, 1)

# ── load data ──────────────────────────────────────────────────────────
def load():
    raw = pd.read_csv(SURVEY, header=None)
    headers = raw.iloc[0].tolist()
    data = raw.iloc[1:].reset_index(drop=True)       # skip header row
    # Some CSVs have sub-header rows; filter to rows that have an actual grade value
    grade_col = data.iloc[:, C_GRADE].astype(str).str.strip()
    valid_grades = grade_col.isin(["9th", "10th", "11th", "12th",
                                   "7th", "8th", "9th grade", "10th grade",
                                   "11th grade", "12th grade"])
    data = data[valid_grades].reset_index(drop=True)
    # Normalise grade to simple form
    data.iloc[:, C_GRADE] = data.iloc[:, C_GRADE].astype(str).str.strip().str.replace(" grade", "", regex=False)
    mtf = pd.read_csv(MTF)
    return data, headers, mtf

# ── compute sections ──────────────────────────────────────────────────
def demographics(df):
    N = len(df)
    grades = df.iloc[:, C_GRADE].value_counts().to_dict()
    sex = df.iloc[:, C_SEX].astype(str).str.strip().value_counts().to_dict()
    ages = df.iloc[:, C_AGE].dropna().astype(str).str.strip()
    age_counts = {}
    for a in ages:
        try:
            ai = int(float(a))
            age_counts[ai] = age_counts.get(ai, 0) + 1
        except:
            pass

    grade_by_sex = {}
    for g in ["9th", "10th", "11th", "12th"]:
        sub = df[df.iloc[:, C_GRADE] == g]
        grade_by_sex[g] = sub.iloc[:, C_SEX].astype(str).str.strip().value_counts().to_dict()

    return {
        "total": N,
        "grades": {g: int(grades.get(g, 0)) for g in ["9th", "10th", "11th", "12th"]},
        "sex": {k: int(v) for k, v in sex.items() if k in ["Male", "Female"]},
        "ages": {str(k): int(v) for k, v in sorted(age_counts.items())},
        "grade_by_sex": grade_by_sex
    }

def substance_prevalence(df):
    N = len(df)
    result = {"substances": SUBSTANCE_NAMES, "lifetime": {}, "past_30_days": {}}

    for period, (start, end) in [("lifetime", (LIFE_START, LIFE_END)),
                                  ("past_30_days", (P30_START, P30_END))]:
        all_pcts = []
        by_grade = {g: [] for g in ["9th", "10th", "11th", "12th"]}
        for c in range(start, end + 1):
            all_pcts.append(yes_pct(df.iloc[:, c]))
            for g in by_grade:
                sub = df[df.iloc[:, C_GRADE] == g]
                by_grade[g].append(yes_pct(sub.iloc[:, c]))
        result[period]["all"] = all_pcts
        result[period].update(by_grade)

    return result

def alcohol_detail(df):
    N = len(df)
    past12 = yes_pct(df.iloc[:, C_PAST12_ALCOHOL])
    # Past 30 days: "I did not drink..." counts as no
    p30 = df.iloc[:, C_30D_DAYS].dropna().astype(str).str.strip()
    p30_any = pct(len(p30[~p30.str.contains("did not|0 days", case=False, regex=True, na=False)]), len(p30)) if len(p30) > 0 else 0
    # Binge
    binge = df.iloc[:, C_30D_BINGE].dropna().astype(str).str.strip()
    binge_any = pct(len(binge[~binge.str.contains("did not|0 days|I did not", case=False, regex=True, na=False)]), len(binge)) if len(binge) > 0 else 0

    by_grade = {}
    for g in ["9th", "10th", "11th", "12th"]:
        sub = df[df.iloc[:, C_GRADE] == g]
        n_g = len(sub)
        p12 = yes_pct(sub.iloc[:, C_PAST12_ALCOHOL])
        s30 = sub.iloc[:, C_30D_DAYS].dropna().astype(str).str.strip()
        s30_any = pct(len(s30[~s30.str.contains("did not|0 days", case=False, regex=True, na=False)]), len(s30)) if len(s30) > 0 else 0
        sb = sub.iloc[:, C_30D_BINGE].dropna().astype(str).str.strip()
        sb_any = pct(len(sb[~sb.str.contains("did not|0 days|I did not", case=False, regex=True, na=False)]), len(sb)) if len(sb) > 0 else 0
        by_grade[g] = {"past_12mo": p12, "past_30d": s30_any, "binge": sb_any}

    return {"past_12mo_rate": past12, "past_30d_any": p30_any, "binge_rate": binge_any, "by_grade": by_grade}

def mtf_comparison(df, mtf):
    """Compare school 10th/12th rates to MTF 'Total' rates for key substances."""
    # Map SUPP substance columns to MTF variable names
    # Map: (display_label, mtf_variable_name, supp_lifetime_col_index)
    sub_map = {
        "Alcohol": ("Alcohol", LIFE_START),
        "Marijuana": ("Marijuana", LIFE_START + 4),
        "Vaping": ("Vaping Nicotine", LIFE_START + 2),
        "Cigarettes": ("Cigarettes", LIFE_START + 1),
    }

    results = []
    for label, (mtf_var, life_col) in sub_map.items():
        p30_col = life_col + (P30_START - LIFE_START)
        # MTF has "Total" rows for past-12-months and past-30-days but not always lifetime
        # Use past-12-months and past-30-days for comparison
        for timeline, col, mtf_tl in [("past_12_months", life_col, "past-12-months"),
                                       ("past_30_days", p30_col, "past-30-days")]:
            # For past_12_months school data: use lifetime col as proxy (lifetime >= 12-month)
            # (The survey's substance Yes/No cols don't distinguish 12-month from lifetime
            #  except for alcohol which has a dedicated past-12-month question)
            school_10 = yes_pct(df[df.iloc[:, C_GRADE] == "10th"].iloc[:, col])
            school_12 = yes_pct(df[df.iloc[:, C_GRADE] == "12th"].iloc[:, col])

            mtf_row = mtf[(mtf["variable"] == mtf_var) & (mtf["timeline"] == mtf_tl) & (mtf["Type"] == "Total")]
            nat_10 = 0
            nat_12 = 0
            if len(mtf_row) > 0:
                try: nat_10 = float(mtf_row["10th"].values[0])
                except: nat_10 = 0
                try: nat_12 = float(mtf_row["12th"].values[0])
                except: nat_12 = 0

            results.append({
                "substance": label,
                "timeline": timeline,
                "school_10th": school_10,
                "school_12th": school_12,
                "national_10th": nat_10,
                "national_12th": nat_12
            })

    return results

def consequences_data(df):
    """% who experienced each consequence at least once (among all respondents)."""
    labels = [
        "Tried to cut down", "Thought had a problem", "Felt guilty",
        "Needed a drink first thing", "Used alone", "Got drunk",
        "Did something regretted", "Fight with friend/partner",
        "Fight with family", "Unwanted sexual experience",
        "Missed school/work", "Trouble concentrating",
        "Vomited/hangover", "Drove after drinking",
        "Rode with impaired driver", "Trouble at school/home",
        "Hung out with bad influences", "Blacked out",
        "Passed out", "Got hurt/injured",
        "Drank before school", "Tried other substances",
        "Drank more than intended"
    ]
    pcts = []
    for c in range(CONS_START, CONS_END + 1):
        pcts.append(consequence_pct(df.iloc[:, c]))

    # Sort by pct descending
    paired = sorted(zip(labels, pcts), key=lambda x: -x[1])
    return {"labels": [p[0] for p in paired], "pcts": [p[1] for p in paired]}

def school_climate(df):
    pcts = []
    for c in range(CLIMATE_START, CLIMATE_END + 1):
        pcts.append(likert_agree_pct(df.iloc[:, c]))
    # positive_score: invert negatively worded items (0,1,3,6,8)
    neg_indices = [0, 1, 3, 6, 8]
    pos_score = []
    for i, p in enumerate(pcts):
        pos_score.append(round(100 - p, 1) if i in neg_indices else p)
    return {"labels": CLIMATE_LABELS, "pct_agree": pcts, "positive_score": pos_score}

def belonging(df):
    pcts = []
    for c in range(BELONG_START, BELONG_END + 1):
        pcts.append(likert_agree_pct(df.iloc[:, c]))
    return {"labels": BELONG_LABELS, "pct_agree": pcts}

def family(df):
    pcts = []
    for c in range(FAMILY_START, FAMILY_END + 1):
        pcts.append(likert_agree_pct(df.iloc[:, c]))
    with_parent = yes_pct(df.iloc[:, C_USE_WITH_PARENT])
    without_parent = yes_pct(df.iloc[:, C_USE_WITHOUT_PARENT])
    return {
        "labels": FAMILY_LABELS,
        "pct_agree": pcts,
        "used_with_parent_pct": with_parent,
        "used_without_parent_pct": without_parent
    }

def stress_data(df):
    pcts = []
    for c in range(STRESS_START, STRESS_END + 1):
        pcts.append(likert_agree_pct(df.iloc[:, c]))
    return {"labels": STRESS_LABELS, "pcts": pcts}

def mental_health(df):
    """Kessler K10 scoring."""
    scores = []
    for _, row in df.iterrows():
        total = 0
        valid = True
        for c in range(K10_START, K10_END + 1):
            val = str(row.iloc[c]).strip() if pd.notna(row.iloc[c]) else ""
            if val in K10_MAP:
                total += K10_MAP[val]
            else:
                valid = False
                break
        if valid:
            scores.append(total)

    scores = np.array(scores)
    n = len(scores)
    cats = {
        "Well (10-15)": int(((scores >= 10) & (scores <= 15)).sum()),
        "Mild (16-21)": int(((scores >= 16) & (scores <= 21)).sum()),
        "Moderate (22-29)": int(((scores >= 22) & (scores <= 29)).sum()),
        "Severe (30-50)": int(((scores >= 30) & (scores <= 50)).sum()),
    }
    cats_pct = {k: round(v / n * 100, 1) if n > 0 else 0 for k, v in cats.items()}

    # Individual items: % answering "Some of the time" or worse
    item_pcts = []
    for c in range(K10_START, K10_END + 1):
        col = df.iloc[:, c].dropna().astype(str).str.strip()
        elevated = col.isin(["Some of the time", "Most of the time", "All of the time"])
        item_pcts.append(round(elevated.sum() / len(col) * 100, 1) if len(col) > 0 else 0)

    return {
        "n_scored": n,
        "mean_score": round(float(scores.mean()), 1) if n > 0 else 0,
        "categories": cats,
        "categories_pct": cats_pct,
        "items": {"labels": K10_LABELS, "pct_elevated": item_pcts}
    }

def headline_stats(df, demo, subst, alc, mh):
    return {
        "total": demo["total"],
        "pct_ever_alcohol": subst["lifetime"]["all"][0],
        "pct_ever_vaped": subst["lifetime"]["all"][2],
        "pct_ever_marijuana": subst["lifetime"]["all"][4],
        "past_12mo_alcohol": alc["past_12mo_rate"],
        "past_30d_alcohol": alc["past_30d_any"],
        "pct_moderate_severe_k10": round(
            mh["categories_pct"].get("Moderate (22-29)", 0) +
            mh["categories_pct"].get("Severe (30-50)", 0), 1
        )
    }

# ── main ───────────────────────────────────────────────────────────────
def main():
    df, headers, mtf = load()
    demo = demographics(df)
    subst = substance_prevalence(df)
    alc = alcohol_detail(df)
    mtf_comp = mtf_comparison(df, mtf)
    cons = consequences_data(df)
    clim = school_climate(df)
    bel = belonging(df)
    fam = family(df)
    strs = stress_data(df)
    mh = mental_health(df)
    head = headline_stats(df, demo, subst, alc, mh)

    output = {
        "meta": {"school": "ISM", "survey_date": "Fall 2025"},
        "headline_stats": head,
        "demographics": demo,
        "substance_prevalence": subst,
        "alcohol_detail": alc,
        "mtf_comparison": mtf_comp,
        "consequences": cons,
        "school_climate": clim,
        "belonging": bel,
        "family": fam,
        "stress": strs,
        "mental_health": mh
    }
    print(json.dumps(output, indent=2))

if __name__ == "__main__":
    main()
