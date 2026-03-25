"""
SUPP Report Generation — Prompt Templates

Contains the system prompt and user prompt for Claude's code execution + docx skill
to analyze SUPP survey data and produce a professional Word report.
"""

SYSTEM_PROMPT = """\
You are an expert analyst for the Student Substance Use and Peer Perception (SUPP) Survey, \
developed by Prevention Education Solutions. You process Typeform CSV exports and produce \
comprehensive Word reports with national benchmark comparisons. You have access to \
supp_processor.py which contains all the analysis functions you need.
"""

REPORT_PROMPT = """\
# Task

Analyze the attached SUPP survey data for **{school_name}** and generate a professional \
Word document report. The survey was administered **{survey_date}** to grades **{grades}**.

# Files Provided

1. **survey_data.csv** — Typeform CSV export with student responses
2. **MTF2024.csv** — Monitoring the Future 2024 national benchmark data
3. **supp_processor.py** — Analysis module with all processing functions

# Step 1: Run the Analysis Pipeline

```python
import sys
sys.path.insert(0, '.')
import supp_processor as sp
import pandas as pd
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker

# Load and process
df = sp.load_survey_csv('survey_data.csv')
df = sp.rename_typeform_columns(df)
df, removed, removal_pct = sp.clean_survey_data(df)
N = len(df)

# Demographics
demo = sp.analyze_demographics(df)
grade_df = demo.get('grade')
sex_df = demo.get('sex')
age_stats = demo.get('age_stats', {{}})

# Risk categorization
df = sp.categorize_drinking_risk(df)
risk_counts = df['drinker_risk'].value_counts()

# Substance prevalence (past 12 months and past 30 days)
sub12_cols = [c for c in df.columns if c.endswith('_12')]
sub30_cols = [c for c in df.columns if c.endswith('_30')]
prev_12 = sp.calculate_substance_prevalence(df, sub12_cols)
prev_30 = sp.calculate_substance_prevalence(df, sub30_cols)

# K10 Mental Health
k10 = sp.analyze_k10(df, sp.K10_COLUMNS)

# MTF comparisons
mtf_df = sp.load_mtf_data('MTF2024.csv')

# Frequency tables for key questions
freq_tables = {{}}
for col in ['how_often_1_drink', 'how_many_drinks', 'days_5_or_more_drinks',
            'how_often_pot', 'age_at_first_drink', 'age_first_time_pot']:
    if col in df.columns:
        freq_tables[col] = sp.create_count_and_percent(df, col)

# School climate (agree/disagree)
climate_cols = ['social_atmosphere', 'pressured', 'easy_to_make_friends',
                'students_assume', 'own_decisions', 'respect', 'fit_in',
                'light_drinkers', 'binge_okay']
climate_data = {{}}
for col in climate_cols:
    if col in df.columns:
        climate_data[col] = sp.analyze_agree_disagree(df, col)
```

# Step 2: Generate Charts

Generate these charts using matplotlib. Save each as PNG.

**Color palette:**
- School data: `#1A63CD` (primary blue)
- MTF national: `#C9D2DA` (light gray)
- Accent: `#267EF2`
- Use white gridlines, clean axis labels, Georgia or sans-serif font

**Charts to generate:**

1. **Substance Prevalence (Past 12 Months)** — Horizontal bar chart of top substances
2. **Substance Prevalence (Past 30 Days)** — Same format for 30-day window
3. **MTF Comparison: Alcohol** — Grouped bar chart: school vs MTF by grade (8th, 10th, 12th) for lifetime/12mo/30day
4. **MTF Comparison: Marijuana** — Same format
5. **MTF Comparison: Vaping** — Same format
6. **K10 Distress Distribution** — Bar chart of Low/Moderate/High/Very High percentages
7. **Drinking Risk Categories** — Pie or bar chart of Non-drinker/Low-risk/High-risk
8. **School Climate** — Stacked horizontal bar for agree/disagree items

For MTF charts: use `sp.compare_to_mtf(school_pct, mtf_df, substance, timeline, grade)` to get MTF values.
Map substance columns to MTF variable names:
- use_alcohol_12/30 → "Alcohol"
- use_pot_12/30 → "Marijuana"
- use_vapes_12/30 → "Vaping Nicotine"
- use_cigarettes_12/30 → "Cigarettes"
- use_cocaine_12/30 → "Cocaine"
- use_ecstasy_12/30 → "Ecstasy"

MTF timelines: "past-12-months", "past-30-days", "lifetime"
MTF grades: "8th", "10th", "12th"

# Step 3: Generate the Word Report

Use the **docx skill** to create a professional report with this structure:

## Document Formatting
- **Title font:** Georgia 14pt bold, dark blue (#1A63CD)
- **Heading 1:** Georgia 12pt bold
- **Body text:** Georgia 10pt
- **Table headers:** White bold text on dark blue (#1A63CD) background
- **Table alt rows:** Light lavender (#ebebf8) shading
- **Page size:** Letter (8.5 x 11)
- **Margins:** 1 inch

## Report Sections

### Cover Page
- Title: "Student Substance Use and Peer Perception Survey"
- Subtitle: "Report for {school_name}"
- Survey date: {survey_date}
- "Prepared by Prevention Education Solutions"

### 1. Introduction & Methodology
- Brief overview of the SUPP survey purpose
- Sample size (N), grades surveyed, data cleaning note (removed X responses, Y%)
- Demographic breakdown table (grade and sex)

### 2. Executive Summary
Write 6-8 bullet points summarizing the KEY findings. Include:
- Overall alcohol use rate (past 12 months)
- Overall marijuana and vaping rates
- Risk category breakdown (% non-drinker, low-risk, high-risk)
- K10 mental health headline (mean score, % in High/Very High distress)
- Most notable MTF comparison (where school significantly exceeds or is notably below)
- Perception gap finding

### 3. Alcohol Use Analysis
- Past 12-month and 30-day prevalence
- Drinking frequency table (how_often_1_drink)
- Quantity table (how_many_drinks)
- Binge drinking (days_5_or_more_drinks)
- Age at first drink distribution
- Risk category distribution chart
- Include the substance prevalence chart (12mo)

### 4. Marijuana & Cannabis
- Lifetime, 12-month, 30-day prevalence
- Frequency of use (how_often_pot)
- Age at first use

### 5. Vaping & Other Substances
- Vaping prevalence (12mo and 30-day)
- Other substance table (all use_*_12 columns)
- Include the substance prevalence chart (30-day)

### 6. National Benchmark Comparisons (MTF)
- Explain the MTF comparison methodology (relative error margins: 5% and 7.5%)
- Alcohol comparison chart and table by grade
- Marijuana comparison chart and table by grade
- Vaping comparison chart and table by grade
- Summary: where school is above/below/on par with national norms
- Use the 4-category system: Below, On par, Exceeds, Significantly exceeds

### 7. School Climate & Environment
- Stacked bar chart for agree/disagree items
- Table with all 9 climate statements showing % agree vs disagree
- Key findings narrative

### 8. Perceived Norms
- Compare what students THINK peers use vs actual reported use
- Use perception columns if available (drink_how_often_your_grade, etc.)
- Highlight the perception gap

### 9. K10 Mental Health
- K10 methodology explanation (10 items, scored 1-5, range 10-50)
- Score distribution chart
- Category breakdown table (Low/Moderate/High/Very High)
- Mean, median, std dev
- If possible, cross-tabulate K10 by risk category

### 10. Family & Peer Factors
- Parental monitoring items (if available)
- Peer influence items
- Consequences from others' use

### 11. Appendix A — Full Response Tables
- Count and percent table for every survey question
- Use sp.create_count_and_percent(df, column) for each column

### 12. Appendix B — Results by Grade
- For key questions, show by-grade breakdown using sp.create_by_grade_table()

### 13. Appendix C — Results by Risk Category
- Filter df by drinker_risk and show key prevalence rates for each group

# Important Notes

- Round all percentages to 1 decimal place
- Include sample sizes (n=) alongside percentages
- Cap percentages at 100%
- Flag sections with fewer than 5 respondents
- Use professional, objective language throughout
- Embed all charts as images in the document
- Generate the document using the docx skill
"""


def build_prompt(school_name, survey_date, grades):
    """Build the formatted prompt for report generation."""
    if isinstance(grades, list):
        grades_str = ", ".join(grades)
    else:
        grades_str = str(grades)

    return REPORT_PROMPT.format(
        school_name=school_name,
        survey_date=survey_date,
        grades=grades_str,
    )
