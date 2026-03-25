"""
SUPP Survey Data Processor

Core data processing functions for the Student Substance Use and Peer Perception Survey.
Handles CSV ingestion, cleaning, categorization, and statistical analysis.
"""

import warnings
from collections import OrderedDict

import pandas as pd

warnings.filterwarnings("ignore", category=FutureWarning)
warnings.filterwarnings("ignore", category=DeprecationWarning)


# ---------------------------------------------------------------------------
# Typeform → Internal column mapping
# ---------------------------------------------------------------------------
# When Typeform CSV is loaded, columns are renamed from verbose survey text
# to short internal names using this ordered list. The CSV columns (after
# stripping the Typeform metadata columns) map 1:1 positionally to collist.

COLLIST = [
    "sex", "age", "grade_level", "grades_overall", "age_at_first_drink",
    "where_first_use_alcohol", "how_often_1_drink", "how_many_drinks",
    "one_drink_past_12", "felt_or_tried_to_cut", "drinking_or_drug_problem",
    "felt_guilty_about_your_drinking", "needed_a_drink_first_thing",
    "used_alcohol_by_yourself", "get_drunk", "do_something_you_later_felt_sorry_for",
    "friend_fight_argument_bad_feelsing", "family_fight_argument_bad_feelsing",
    "sexual_experience", "perform_poorly_school_work", "trouble_concentrating",
    "sick_vomit_hangover", "drive", "ride", "trouble_school_home",
    "hang_out_bad_influence", "forget_what_happened", "pass_out", "get_hurt",
    "drink_before_school", "try_other_substances", "drank_more",
    "one_drink_past_30_days", "drinks_in_a_row_30_days", "days_5_or_more_drinks",
    "drunk_30_days", "where_do_you_drink", "age_first_time_pot", "how_often_pot",
    # Past 12 months substance use
    "use_alcohol_12", "use_cigarettes_12", "use_vapes_12", "use_smokeless_tobacco_12",
    "use_pot_12", "use_ecstasy_12", "use_heroin_12", "use_narcotics_12",
    "use_amphetamines_12", "use_sedatives_12", "use_methamphetamine_12",
    "use_cocaine_12", "use_lsd_12", "use_hallucinogen_12", "use_inhalants_12",
    "use_steroids_12", "use_cold_medicine_12",
    # Past 30 days substance use
    "use_alcohol_30", "use_cigarettes_30", "use_vapes_30", "use_smokeless_tobacco_30",
    "use_pot_30", "use_ecstasy_30", "use_heroin_30", "use_narcotics_30",
    "use_amphetamines_30", "use_sedatives_30", "use_methamphetamine_30",
    "use_cocaine_30", "use_lsd_30", "use_hallucinogen_30", "use_inhalants_30",
    "use_steroids_30", "use_cold_medicine_30",
    # Negative consequences from others' use
    "taking_concerned_other", "taking_care_drank_too_much_other",
    "taking_care_consumed_too_much_other", "emergency_medical_care_other",
    "in_trouble_other", "feeling_insulted_other", "verbal_argument_other",
    "pushed_hit_other", "sexual_experience_other", "witness_vomit_other",
    "unsafe_other", "sleep_difficulty_other", "homework_difficulty_other",
    # School events, attitudes, norms
    "school_dance_use", "attitude",
    "one_2_friend", "three_4_friend", "five_friend", "got_drunk_friend",
    "harm_1_2_drinks", "harm_1_2_daily", "harm_3_4_daily", "harm_5",
    "pot_friend_once", "pot_friend_occasion", "pot_friend_regular",
    # School environment (agree/disagree)
    "social_atmosphere", "pressured", "easy_to_make_friends", "students_assume",
    "own_decisions", "respect", "fit_in", "light_drinkers", "binge_okay",
    # Comfort discussing
    "discussion_friend", "discussion_guardian", "discussion_teacher",
    # Perceived norms
    "drink_how_often_your_grade", "drink_how_often_seniors",
    "how_many_your_grade", "how_many_seniors",
    "one_drink_30_days_your_school", "one_drink_30_days_your_grade",
    "one_drink_30_days_seniors", "how_many_30_days_your_grade",
    "how_many_30_days_your_seniors", "marijuana_your_grade", "marijuana_seniors",
    "percent_ever_used",
    # Policies
    "thoughts_alcohol_policies", "thoughts_enforcement",
    "thoughts_other_alcohol_policies", "thoughts_other_enforcement",
    # School connectedness
    "valued", "academics_to_me", "teachers", "sincere_effort", "care",
    "academics_to_others", "adult_trust", "school_12",
    # Family factors
    "parent_interest", "clear_rules", "call", "come_home",
    "clear_rules_alcohol", "not_home", "use_parent_knowing",
    "use_parent_not_knowing", "meals",
    # Friend & peer factors
    "friend_problem", "talking_parents", "talking_parents_about_friend",
    "classmates_pressured", "parents_present_visible",
    # Stress / reasons for use
    "academic_stress", "parent_stress", "college_stress", "overlooked_stress",
    "expectations_stress", "boredom_stress", "fitting_in_stress",
    # Coping substances
    "coping_alcohol", "coping_cigarettes", "coping_vape", "coping_smokeless_tobacco",
    "coping_marijuana", "coping_molly", "coping_heroin", "coping_narcotics",
    "coping_amphetamines", "coping_sedatives", "coping_meth", "coping_cocaine",
    "coping_lsd", "coping_hallucinogens", "coping_inhalant", "coping_steroid",
    "coping_cold_med",
    # K10 mental health
    "tired", "nervous", "no_calm", "hopless", "restless", "no_sit_still",
    "depressed", "effort", "cheer", "worthless",
]

# K10 columns (by internal name) for scoring
K10_COLUMNS = [
    "tired", "nervous", "no_calm", "hopless", "restless", "no_sit_still",
    "depressed", "effort", "cheer", "worthless",
]


# ---------------------------------------------------------------------------
# Response sort orders (exact values from Typeform responses)
# ---------------------------------------------------------------------------
# These lists define the correct display order for table rows.
# Tables MUST use these to sort responses; do not use alphabetical sort.

SORT_ORDERS = {
    "drinking_risk": ["High-risk", "Low-risk", "Non-drinker"],

    "drinking_frequency": [  # sort_order1
        "Never", "1-2 times per year", "Once a month",
        "Twice a month", "Once a week", "Twice a week", "Daily",
    ],

    "drink_amount": [  # amountlist
        "0 drinks", "1-2 drinks", "3-4 drinks", "5 or more drinks",
    ],

    "drink_amount_no_zero": [  # amountlist2
        "1-2 drinks", "3-4 drinks", "5 or more drinks",
    ],

    "consequence_frequency": [  # sort_order2
        "Never", "1 time", "2 times", "3 times or more",
    ],

    "past_30_days_drinking": [  # sort_order3
        "I did not drink alcohol during the past 30 days",
        "1 day", "2 days", "3 to 5 days", "6 to 9 days", "10 or more days",
    ],

    "times_frequency": [  # freq_sort
        "0 times", "1 time", "2 times", "3 times", "4 times", "5 times or more",
    ],

    "binge_30_days": [  # sort_order4
        "I did not consume 5 or more whole drinks of alcohol in a row over the past 30 days",
        "1 day", "2 days", "3 to 5 days", "6 to 9 days", "10 or more days",
    ],

    "marijuana_frequency": [  # sort_order5
        "Never", "1 to 2 times per year", "Once a month",
        "Twice a month", "Once a week", "Twice a week", "Daily",
    ],

    "percentage_ranges": [  # perc_order
        "0%", "1-10%", "11-20%", "21-30%", "31-40%", "41-50%",
        "51-60%", "61-70%", "71-80%", "81-90%", "91-100%",
    ],

    "yes_no": ["Yes", "No"],

    "drinking_attitude": [  # sort_order6
        "Drinking is never a good thing to do.",
        "Drinking is okay, but a person should not get drunk.",
        "Occasionally getting drunk is okay as long as it DOESN'T interfere with academics or other responsibilities",
        "Occasionally getting drunk is okay EVEN IF IT DOES interfere with academics or other responsibilities",
        "Frequently getting drunk is okay if an individual wants to do that",
    ],

    "disapprove_approve": [  # sort_order7
        "Strongly disapprove", "Disapprove",
        "Neither disapprove nor approve", "Approve", "Strongly approve",
    ],

    "risk_perception": [  # sort_order8
        "No risk", "Slight risk", "Moderate risk", "Great risk",
    ],

    "disapprove_approve_short": [  # sort_order9
        "Strongly disapprove", "Disapprove", "Approve", "Strongly approve",
    ],

    "agree_disagree": [  # sort_order10
        "Strongly disagree", "Disagree", "Agree", "Strongly agree",
    ],

    "never_once_more": [  # sort_order11
        "Never", "Once", "More than once",
    ],

    "days_frequency": [  # sort_order12
        "0 days", "1 day", "2 days", "3 to 5 days", "6 to 9 days", "10 or more days",
    ],

    "policy_judgment": [  # sort_order13
        "Too lenient", "About right", "Too strict", "Don't know",
    ],

    "school_feeling": [  # sort_order14
        "Liked a lot", "Kind of liked", "Did not like", "Hated",
    ],

    "weekly_frequency": [  # sort_order15
        "Never", "1-2 days a week", "3-5 days a week", "6-7 days a week",
    ],

    "adverb_frequency": [  # sort_order16
        "Rarely", "Sometimes", "Most of the time", "Always",
    ],

    "time_frequency": [  # sort_order17
        "All of the time", "Most of the time",
        "A little of the time", "Some of the time", "None of the time",
    ],

    "limited_times": [  # sort_order18
        "Never", "1 time", "2 times", "3 or more times",
    ],

    "marijuana_age": [  # pot_age
        "8 years old or younger", "9 to 10 years old", "11 to 12 years old",
        "13 to 14 years old", "15 to 16 years old", "17 to 18 years old",
        "19 years or older", "I have never used marijuana at all",
    ],

    "drink_locations": [
        "At a friend's house", "At home", "At a club or bar",
        "At camp", "At school", "Other",
    ],

    "academic_grades": [
        "Mostly 1's", "Mostly 2's", "Mostly 3's", "Mostly 4's",
        "Mostly 5's", "Mostly 6's", "Mostly 7's",
    ],
}


# ---------------------------------------------------------------------------
# MTF variable name mapping (CSV 'variable' column → display name)
# ---------------------------------------------------------------------------
# These are the exact values in the MTF CSV 'variable' column.

MTF_VARIABLES = [
    "Alcohol", "Amphetamines", "Cigarettes", "Cocaine", "Ecstasy",
    "Hallucinogens", "Heroin", "Inhalants", "LSD", "Marijuana",
    "Methamphetamine", "Narcotics", "Steroids", "Vaping Nicotine", "Any Vaping",
]

# Map internal SUPP substance column names to MTF variable names
SUBSTANCE_TO_MTF = {
    "use_alcohol_12": ("Alcohol", "past-12-months"),
    "use_alcohol_30": ("Alcohol", "past-30-days"),
    "use_cigarettes_12": ("Cigarettes", "past-12-months"),
    "use_cigarettes_30": ("Cigarettes", "past-30-days"),
    "use_vapes_12": ("Vaping Nicotine", "past-12-months"),
    "use_vapes_30": ("Vaping Nicotine", "past-30-days"),
    "use_smokeless_tobacco_12": ("Smokeless tobacco", "past-12-months"),
    "use_smokeless_tobacco_30": ("Smokeless tobacco", "past-30-days"),
    "use_pot_12": ("Marijuana", "past-12-months"),
    "use_pot_30": ("Marijuana", "past-30-days"),
    "use_ecstasy_12": ("Ecstasy", "past-12-months"),
    "use_ecstasy_30": ("Ecstasy", "past-30-days"),
    "use_heroin_12": ("Heroin", "past-12-months"),
    "use_heroin_30": ("Heroin", "past-30-days"),
    "use_narcotics_12": ("Narcotics", "past-12-months"),
    "use_narcotics_30": ("Narcotics", "past-30-days"),
    "use_amphetamines_12": ("Amphetamines", "past-12-months"),
    "use_amphetamines_30": ("Amphetamines", "past-30-days"),
    "use_sedatives_12": ("Sedatives", "past-12-months"),
    "use_sedatives_30": ("Sedatives", "past-30-days"),
    "use_methamphetamine_12": ("Methamphetamine", "past-12-months"),
    "use_methamphetamine_30": ("Methamphetamine", "past-30-days"),
    "use_cocaine_12": ("Cocaine", "past-12-months"),
    "use_cocaine_30": ("Cocaine", "past-30-days"),
    "use_lsd_12": ("LSD", "past-12-months"),
    "use_lsd_30": ("LSD", "past-30-days"),
    "use_hallucinogen_12": ("Hallucinogens", "past-12-months"),
    "use_hallucinogen_30": ("Hallucinogens", "past-30-days"),
    "use_inhalants_12": ("Inhalants", "past-12-months"),
    "use_inhalants_30": ("Inhalants", "past-30-days"),
    "use_steroids_12": ("Steroids", "past-12-months"),
    "use_steroids_30": ("Steroids", "past-30-days"),
    "use_cold_medicine_12": ("OTC cold medicine", "past-12-months"),
    "use_cold_medicine_30": ("OTC cold medicine", "past-30-days"),
}

# Chart color palette (matches production script)
CHART_COLORS = {
    "primary_blue": "#1A63CD",
    "light_gray": "#C9D2DA",
    "accent_blue": "#267EF2",
    "dark_gray": "#838383",
    "total": "#1A63CD",
    "male": "#C9D2DA",
    "female": "#267EF2",
    "agree": "#1A63CD",
    "disagree": "#C9D2DA",
}


# ---------------------------------------------------------------------------
# Utility helpers
# ---------------------------------------------------------------------------

def safe_round(value, decimals=1):
    """Safely round a value, handling strings and NaN."""
    try:
        return round(float(value), decimals)
    except (ValueError, TypeError):
        return 0.0


def cap_percentage(value, cap=100.0):
    """Cap a percentage at a maximum value to prevent rounding artifacts."""
    return min(safe_round(value), cap)


def safe_item(x):
    """Extract a scalar from a pandas Series or return the value as-is."""
    if hasattr(x, "iloc"):
        return x.iloc[0] if len(x) > 0 else 0
    return x


# ---------------------------------------------------------------------------
# Data loading & cleaning
# ---------------------------------------------------------------------------

def load_survey_csv(filepath, encoding="latin-1"):
    """
    Load a Typeform CSV export of SUPP survey data.

    Handles apostrophe normalization (Typeform exports contain curly quotes
    that become mojibake under latin-1 encoding).

    Args:
        filepath: Path to the CSV file
        encoding: File encoding (default latin-1 for Typeform exports)

    Returns:
        pandas DataFrame with raw survey data
    """
    df = pd.read_csv(filepath, encoding=encoding)
    df.columns = df.columns.str.strip()

    # Normalize curly/smart apostrophes in column names and string values
    apostrophe_variants = [
        "\u2019", "\u2018",  # Unicode right/left single quotes
        "\xc2\x92",          # Windows-1252 right quote as latin-1
        "\xe2\x80\x99",      # UTF-8 bytes for \u2019 read as latin-1
        "\xe2\x80\x98",      # UTF-8 bytes for \u2018 read as latin-1
    ]
    for apo in apostrophe_variants:
        df.columns = [c.replace(apo, "'") for c in df.columns]
    for col in df.select_dtypes(include="object").columns:
        # Skip if duplicate column name returns a DataFrame instead of Series
        if isinstance(df[col], pd.DataFrame):
            continue
        for apo in apostrophe_variants:
            df[col] = df[col].str.replace(apo, "'", regex=False)

    return df


def rename_typeform_columns(df):
    """
    Rename verbose Typeform column headers to short internal names.

    The Typeform CSV has ~160+ columns with full question text as headers.
    This function renames them positionally using COLLIST, plus the trailing
    metadata columns ('Submitted At', 'Token').

    Args:
        df: DataFrame loaded from Typeform CSV

    Returns:
        DataFrame with renamed columns
    """
    n_extra = len(df.columns) - len(COLLIST)
    if n_extra == 2:
        df.columns = COLLIST + ["Submitted At", "Token"]
    elif n_extra == 3:
        df.columns = COLLIST + ["Submitted At", "Token", "agenum"]
    else:
        # Best effort: use COLLIST for the first N columns
        df.columns = COLLIST + list(df.columns[len(COLLIST):])
    return df


def load_mtf_data(filepath, year=2024):
    """
    Load MTF (Monitoring the Future) benchmark data.

    Args:
        filepath: Path to MTF CSV file
        year: Year to filter for (default 2024)

    Returns:
        pandas DataFrame with MTF benchmarks
    """
    mtf = pd.read_csv(filepath)
    mtf.columns = mtf.columns.str.strip()
    if "Year" in mtf.columns:
        mtf = mtf[mtf["Year"] == year]
    return mtf


def clean_survey_data(df, grade_list=None):
    """
    Clean and validate survey data. Remove contradictory responses.

    Args:
        df: Raw survey DataFrame
        grade_list: List of grade levels to include (e.g., ["9th", "10th", "11th", "12th"])

    Returns:
        Tuple of (cleaned DataFrame, number removed, removal percentage)
    """
    initial_count = len(df)

    # Filter to selected grades if provided
    if grade_list:
        grade_col = _find_grade_column(df)
        if grade_col:
            df = df[df[grade_col].isin(grade_list)].copy()

    # Remove contradictory responses
    df, removed = _remove_contradictions(df)

    removal_pct = safe_round((removed / initial_count * 100) if initial_count > 0 else 0)
    return df, removed, removal_pct


def _find_grade_column(df):
    """Find the grade level column in the DataFrame."""
    for col in df.columns:
        col_lower = col.lower()
        if "grade" in col_lower and "academic" not in col_lower:
            return col
    return None


def _find_sex_column(df):
    """Find the sex/gender column in the DataFrame."""
    for col in df.columns:
        col_lower = col.lower()
        if col_lower in ("sex", "gender") or "what is your sex" in col_lower:
            return col
    return None


def _remove_contradictions(df):
    """
    Remove rows with contradictory marijuana responses.

    Production logic: flag rows where age_first_time_pot says "never used"
    but how_often_pot is not "Never", or vice versa.

    Returns:
        Tuple of (cleaned DataFrame, count removed)
    """
    initial = len(df)

    # Only apply if columns exist (they will after rename_typeform_columns)
    pot_age_col = "age_first_time_pot" if "age_first_time_pot" in df.columns else None
    pot_freq_col = "how_often_pot" if "how_often_pot" in df.columns else None

    if pot_age_col and pot_freq_col:
        contradiction_mask = (
            ((df[pot_age_col] == "I have never used marijuana at all")
             & (df[pot_freq_col] != "Never"))
            | ((df[pot_freq_col] == "Never")
               & (df[pot_age_col] != "I have never used marijuana at all"))
        )
        df = df[~contradiction_mask].reset_index(drop=True)

    removed = initial - len(df)
    return df, removed


# ---------------------------------------------------------------------------
# Demographic analysis
# ---------------------------------------------------------------------------

def analyze_demographics(df):
    """
    Generate demographic summary tables.

    Returns:
        dict with keys: sex, grade, age_stats, academic_performance
    """
    results = {}

    # Sex distribution
    sex_col = _find_sex_column(df)
    if sex_col:
        results["sex"] = create_count_and_percent(df, sex_col)

    # Grade distribution
    grade_col = _find_grade_column(df)
    if grade_col:
        results["grade"] = create_count_and_percent(df, grade_col)

    # Age statistics
    age_col = _find_age_column(df)
    if age_col:
        ages = pd.to_numeric(df[age_col], errors="coerce").dropna()
        results["age_stats"] = {
            "min": int(ages.min()) if len(ages) > 0 else 0,
            "max": int(ages.max()) if len(ages) > 0 else 0,
            "mean": safe_round(ages.mean()),
            "std": safe_round(ages.std()),
            "n": len(ages),
        }

    return results


def _find_age_column(df):
    """Find the age column in the DataFrame."""
    for col in df.columns:
        if "age" in col.lower() and "first" not in col.lower():
            return col
    return None


# ---------------------------------------------------------------------------
# Count & percent tables
# ---------------------------------------------------------------------------

def create_count_and_percent(df, column, sort_numeric=False):
    """
    Create a count and percentage table for a single column.

    Args:
        df: DataFrame
        column: Column name to analyze
        sort_numeric: Whether to sort values numerically

    Returns:
        DataFrame with columns: Response, Count, Percent
    """
    counts = df[column].value_counts(dropna=False)
    total = counts.sum()

    result = pd.DataFrame({
        "Response": counts.index,
        "Count": counts.values,
        "Percent": [cap_percentage(c / total * 100) for c in counts.values],
    })

    if sort_numeric:
        result["_sort"] = pd.to_numeric(
            result["Response"].astype(str).str.extract(r"(\d+)")[0],
            errors="coerce",
        )
        result = result.sort_values("_sort").drop(columns=["_sort"])

    result = result.reset_index(drop=True)

    # Add totals row
    totals = pd.DataFrame([{
        "Response": "Total",
        "Count": total,
        "Percent": cap_percentage(result["Percent"].sum()),
    }])
    result = pd.concat([result, totals], ignore_index=True)

    return result


def create_by_grade_table(df, column, grade_col=None):
    """
    Create a crosstab of responses by grade level.

    Returns:
        DataFrame with grade columns showing Count (%) format
    """
    if grade_col is None:
        grade_col = _find_grade_column(df)
    if grade_col is None:
        return pd.DataFrame()

    ct = pd.crosstab(df[column], df[grade_col], margins=True, margins_name="Total")
    pct = pd.crosstab(df[column], df[grade_col], normalize="columns", margins=False) * 100

    return ct, pct


def sort_by_custom_list(df, column, sort_order):
    """
    Sort a DataFrame by a custom order list for the given column.

    Args:
        df: DataFrame to sort
        column: Column name to sort by
        sort_order: List of values in desired order

    Returns:
        Sorted DataFrame. Values not in sort_order are appended at the end.
    """
    if column not in df.columns:
        return df
    order_map = {val: i for i, val in enumerate(sort_order)}
    df = df.copy()
    df["_sort_key"] = df[column].map(order_map).fillna(len(sort_order))
    df = df.sort_values("_sort_key").drop(columns=["_sort_key"]).reset_index(drop=True)
    return df


# Alias for backward compatibility
by_grade_count_and_percent = create_by_grade_table


# ---------------------------------------------------------------------------
# Substance use analysis
# ---------------------------------------------------------------------------

def calculate_substance_prevalence(df, substance_columns):
    """
    Calculate Yes/No prevalence for multiple substance columns.

    Args:
        df: Survey DataFrame
        substance_columns: List of column names OR dict mapping label to column name

    Returns:
        DataFrame with columns: Substance, Yes_count, No_count, Yes_pct, No_pct
    """
    results = []
    total_n = len(df)

    # Accept both list and dict
    if isinstance(substance_columns, list):
        items = [(col, col) for col in substance_columns]
    else:
        items = substance_columns.items()

    for label, col in items:
        if col not in df.columns:
            continue
        yes_count = (df[col].astype(str).str.strip().str.lower() == "yes").sum()
        no_count = total_n - yes_count
        yes_pct = safe_round(yes_count / total_n * 100) if total_n > 0 else 0
        no_pct = safe_round(no_count / total_n * 100) if total_n > 0 else 0
        results.append({
            "Substance": label,
            "Yes_count": int(yes_count),
            "No_count": int(no_count),
            "Yes_pct": yes_pct,
            "No_pct": no_pct,
        })

    result_df = pd.DataFrame(results)
    if not result_df.empty:
        result_df = result_df.sort_values("Yes_pct", ascending=False).reset_index(drop=True)
    return result_df


def categorize_drinking_risk(df):
    """
    Categorize respondents into drinking risk levels using the exact SUPP logic.

    Uses internal column names (after Typeform mapping):
    - days_5_or_more_drinks: Past 30 days binge drinking
    - how_often_1_drink: Typical drinking frequency
    - one_drink_past_12: Had a drink in past 12 months
    - age_at_first_drink: Age at first drink (or "I have never had a whole drink of alcohol")

    Categories:
        - Non-drinker: Never had a drink OR did not drink in past 12 months
        - Low-risk: Drinks but no 5+ drinks in a row in past 30 days
        - High-risk: Had 5+ drinks in a row in past 30 days
        - Missing: Cannot determine from responses

    Returns:
        Series with risk category for each respondent
    """
    no_binge = "I did not consume 5 or more whole drinks of alcohol in a row over the past 30 days"

    def classify(row):
        binge = row.get("days_5_or_more_drinks", "")
        freq = row.get("how_often_1_drink", "")
        past_12 = row.get("one_drink_past_12", "")
        first_drink = row.get("age_at_first_drink", "")

        # High-risk: had 5+ drinks in a row in past 30 days
        if str(binge) != no_binge and pd.notna(binge) and str(binge).strip() != "":
            return "High-risk"

        # Low-risk: drinks (not "Never"), had a drink in past 12 months,
        # but no binge drinking
        if (str(freq) != "Never" and str(past_12) == "Yes"
                and (str(binge) == no_binge or pd.isna(binge))):
            return "Low-risk"

        # Non-drinker: never had a whole drink OR did not drink in past 12 months
        if (str(first_drink) == "I have never had a whole drink of alcohol"
                or str(past_12) == "No"):
            return "Non-drinker"

        return "Missing"

    df = df.copy()
    df["drinker_risk"] = df.apply(classify, axis=1)
    return df


# ---------------------------------------------------------------------------
# MTF comparison
# ---------------------------------------------------------------------------

def compare_to_mtf(school_pct, mtf_df, substance, timeline, grade, gender="Total"):
    """
    Compare a school's prevalence rate to MTF national norms.

    Uses the exact SUPP production logic with 5% and 7.5% error margins:
    - "below": school rate < (MTF - 5% of MTF), or MTF is 0
    - "on par": school rate within +/- 5% of MTF value
    - "exceeds": school rate > (MTF + 5%) but <= (MTF + 7.5%)
    - "significantly exceeds": school rate > (MTF + 7.5%)

    Args:
        school_pct: School's percentage for this measure
        mtf_df: MTF benchmark DataFrame
        substance: Substance name (must match MTF 'variable' column)
        timeline: "lifetime", "past-12-months", or "past-30-days"
        grade: "8th", "10th", or "12th"
        gender: "Total", "Male", or "Female"

    Returns:
        dict with mtf_pct, difference, and comparison label
    """
    mask = (
        (mtf_df["variable"].str.lower() == substance.lower())
        & (mtf_df["timeline"].str.lower() == timeline.lower())
        & (mtf_df["Type"].str.lower() == gender.lower())
    )
    matched = mtf_df[mask]

    if matched.empty or grade not in matched.columns:
        return {"mtf_pct": None, "difference": None, "comparison": "no data"}

    mtf_val = safe_item(matched[grade])

    # Handle mojibake "â" (em-dash) that indicates missing data in MTF CSV
    if isinstance(mtf_val, str) and mtf_val.strip() in ("â", "—", "-", ""):
        return {"mtf_pct": None, "difference": None, "comparison": "no data"}

    try:
        mtf_pct = safe_round(float(mtf_val))
    except (ValueError, TypeError):
        return {"mtf_pct": None, "difference": None, "comparison": "no data"}

    # Production comparison logic: 5% and 7.5% error margins relative to MTF value
    error_5 = mtf_pct * 0.05
    error_7_5 = mtf_pct * 0.075

    if (school_pct <= mtf_pct + error_5) and (mtf_pct - error_5 <= school_pct):
        label = "on par"
    elif (mtf_pct - error_5 > school_pct) or (mtf_pct == 0):
        label = "below"
    elif school_pct > (mtf_pct + error_7_5):
        label = "significantly exceeds"
    elif school_pct > (mtf_pct + error_5):
        label = "exceeds"
    else:
        label = "on par"

    diff = safe_round(school_pct - mtf_pct)
    return {"mtf_pct": mtf_pct, "difference": diff, "comparison": label}


def generate_mtf_bullets(school_data, mtf_df, substance, grades=None):
    """
    Generate natural language comparison bullets for a substance.

    Uses exact SUPP production phrasing with 4 comparison levels:
    - "is below U.S. norms for ..."
    - "is on par with U.S. norms for ..."
    - "exceeds U.S. norms for ..."
    - "significantly exceeds U.S. norms for ..."

    Args:
        school_data: dict of {grade: {timeline: pct}} for the school
        mtf_df: MTF benchmark DataFrame
        substance: Substance name
        grades: List of grades to compare (default: ["8th", "10th", "12th"])

    Returns:
        List of bullet point strings
    """
    if grades is None:
        grades = ["8th", "10th", "12th"]

    bullets = []
    for grade in grades:
        if grade not in school_data:
            continue

        below = []
        on_par = []
        exceeds = []
        sig_exceeds = []

        for timeline in ["lifetime", "past-12-months", "past-30-days"]:
            school_pct = school_data[grade].get(timeline, 0)
            result = compare_to_mtf(school_pct, mtf_df, substance, timeline, grade)
            if result["comparison"] == "no data":
                continue
            if result["comparison"] == "below":
                below.append(timeline)
            elif result["comparison"] == "on par":
                on_par.append(timeline)
            elif result["comparison"] == "exceeds":
                exceeds.append(timeline)
            elif result["comparison"] == "significantly exceeds":
                sig_exceeds.append(timeline)

        parts = []
        if below:
            parts.append(f"is below U.S. norms for {_join_timelines(below)}")
        if on_par:
            parts.append(f"is on par with U.S. norms for {_join_timelines(on_par)}")
        if exceeds:
            parts.append(f"exceeds U.S. norms for {_join_timelines(exceeds)}")
        if sig_exceeds:
            parts.append(f"significantly exceeds U.S. norms for {_join_timelines(sig_exceeds)}")

        if parts:
            bullets.append(f"{grade} grade {substance.lower()} use {' and '.join(parts)} measures.")

    return bullets


def _join_timelines(timelines):
    """Join timeline names with commas and 'and'."""
    clean = [t.replace("-", " ") for t in timelines]
    if len(clean) == 1:
        return clean[0]
    if len(clean) == 2:
        return f"{clean[0]} and {clean[1]}"
    return ", ".join(clean[:-1]) + f", and {clean[-1]}"


# ---------------------------------------------------------------------------
# K10 Mental Health Scoring
# ---------------------------------------------------------------------------

K10_SCORING = {
    # Standard Kessler wording (as used in Typeform SUPP exports)
    "none of the time": 1,
    "a little of the time": 2,
    "some of the time": 3,
    "most of the time": 4,
    "all of the time": 5,
    # Alternate wording (some surveys use these)
    "never": 1,
    "rarely": 2,
    "sometimes": 3,
    "often": 4,
    "always": 5,
}

K10_ITEMS = [
    "tired out for no good reason",
    "nervous",
    "so nervous that nothing could calm you down",
    "hopeless",
    "restless or fidgety",
    "so restless you could not sit still",
    "depressed",
    "everything was an effort",
    "so sad that nothing could cheer you up",
    "worthless",
]


def score_k10(df, k10_columns):
    """
    Score the Kessler K10 psychological distress scale.

    Args:
        df: Survey DataFrame
        k10_columns: List of 10 column names for K10 items

    Returns:
        Series of K10 total scores (10-50)
    """
    scores = pd.DataFrame()
    for col in k10_columns:
        if col in df.columns:
            scores[col] = df[col].astype(str).str.strip().str.lower().map(K10_SCORING)

    return scores.sum(axis=1)


def categorize_k10(score):
    """
    Categorize a K10 score into distress levels.

    Returns:
        str: "Low", "Moderate", "High", or "Very High"
    """
    if score <= 15:
        return "Low"
    if score <= 21:
        return "Moderate"
    if score <= 29:
        return "High"
    return "Very High"


def analyze_k10(df, k10_columns):
    """
    Full K10 analysis with distribution and summary statistics.

    Returns:
        dict with scores, categories, and summary stats
    """
    scores = score_k10(df, k10_columns)
    categories = scores.apply(categorize_k10)

    cat_counts = categories.value_counts()
    total = len(categories)

    distribution = OrderedDict()
    for level in ["Low", "Moderate", "High", "Very High"]:
        n = int(cat_counts.get(level, 0))
        distribution[level] = {
            "n": n,
            "pct": cap_percentage(n / total * 100) if total > 0 else 0,
        }

    return {
        "scores": scores,
        "categories": categories,
        "distribution": distribution,
        "mean": safe_round(scores.mean()),
        "median": safe_round(scores.median()),
        "std": safe_round(scores.std()),
    }


# ---------------------------------------------------------------------------
# Perception vs Reality
# ---------------------------------------------------------------------------

def compare_perception_reality(perceived_pct, actual_pct):
    """
    Compare perceived peer use against actual reported use.

    Returns:
        dict with perceived, actual, gap, and interpretation
    """
    gap = safe_round(perceived_pct - actual_pct)
    if gap > 10:
        interpretation = "Students significantly overestimate peer use"
    elif gap > 5:
        interpretation = "Students moderately overestimate peer use"
    elif gap > -5:
        interpretation = "Perceptions roughly match reality"
    elif gap > -10:
        interpretation = "Students moderately underestimate peer use"
    else:
        interpretation = "Students significantly underestimate peer use"

    return {
        "perceived_pct": perceived_pct,
        "actual_pct": actual_pct,
        "gap": gap,
        "interpretation": interpretation,
    }


# ---------------------------------------------------------------------------
# Agree/Disagree analysis
# ---------------------------------------------------------------------------

AGREE_LABELS = ["Strongly Agree", "Agree"]
DISAGREE_LABELS = ["Strongly Disagree", "Disagree"]
NEUTRAL_LABEL = "Neither Agree nor Disagree"


def analyze_agree_disagree(df, column):
    """
    Analyze an agree/disagree Likert-scale question.

    Returns:
        dict with individual response percentages and totals
    """
    counts = df[column].value_counts()
    total = counts.sum()

    result = {}
    for label in ["Strongly Agree", "Agree", NEUTRAL_LABEL, "Disagree", "Strongly Disagree"]:
        n = int(counts.get(label, 0))
        result[label] = {"n": n, "pct": cap_percentage(n / total * 100) if total > 0 else 0}

    agree_n = sum(counts.get(lbl, 0) for lbl in AGREE_LABELS)
    disagree_n = sum(counts.get(lbl, 0) for lbl in DISAGREE_LABELS)

    result["Total Agree"] = {
        "n": int(agree_n),
        "pct": cap_percentage(agree_n / total * 100) if total > 0 else 0,
    }
    result["Total Disagree"] = {
        "n": int(disagree_n),
        "pct": cap_percentage(disagree_n / total * 100) if total > 0 else 0,
    }

    return result


# ---------------------------------------------------------------------------
# Report context builder
# ---------------------------------------------------------------------------

def build_report_context(df, mtf_df, school_name, survey_date, grades):
    """
    Build the full context dictionary for report template rendering.

    Args:
        df: Cleaned survey DataFrame
        mtf_df: MTF benchmark DataFrame
        school_name: Name of the school
        survey_date: Date(s) of survey administration
        grades: List of grade levels included

    Returns:
        dict with all template variables
    """
    context = {
        "School": school_name,
        "date_survey": survey_date,
        "N": len(df),
        "grades": grades,
        "prettygrades": _format_grade_list(grades),
    }

    # Demographics
    context["demographics"] = analyze_demographics(df)

    return context


def _format_grade_list(grades):
    """Format grade list for display (e.g., '9th, 10th, 11th, and 12th')."""
    if len(grades) == 1:
        return grades[0]
    if len(grades) == 2:
        return f"{grades[0]} and {grades[1]}"
    return ", ".join(grades[:-1]) + f", and {grades[-1]}"
