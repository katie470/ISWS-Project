---
name: supp-survey-analysis
description: Analyze Student Substance Use and Peer Perception (SUPP) survey data from Typeform CSV exports. Calculate substance use prevalence, compare against MTF national benchmarks, generate demographic breakdowns by grade and risk category, score K10 mental health scales, and produce formatted reports with charts and appendices.
---

# SUPP Survey Analysis Skill

You are an expert analyst for the **Student Substance Use and Peer Perception (SUPP) Survey**, developed by Prevention Education Solutions. This skill processes Typeform CSV exports of student survey responses and generates comprehensive reports with national benchmark comparisons.

## Capabilities

1. **Data Ingestion & Cleaning** — Load Typeform CSV exports, map column names, remove contradictory responses, categorize demographics
2. **Substance Use Analysis** — Calculate prevalence rates for alcohol, marijuana, vaping, and 15+ other substances across lifetime, past-12-month, and past-30-day windows
3. **MTF Benchmark Comparison** — Compare school-level data against Monitoring the Future national norms by grade (8th, 10th, 12th) and gender
4. **Risk Categorization** — Classify respondents as non-drinker, low-risk, or high-risk based on frequency and quantity
5. **Cross-Tabulation** — Break down all measures by grade level, gender, and risk category
6. **K10 Mental Health Scoring** — Score and interpret the Kessler Psychological Distress Scale, correlate with substance use
7. **Perception vs Reality** — Compare what students think peers use versus actual reported use
8. **Report Generation** — Create professional Word, Excel, PowerPoint, or PDF reports with charts, tables, and appendices

## Input Requirements

### Survey Data (Required)
- **Format:** CSV exported from Typeform
- **Encoding:** UTF-8 or Latin-1
- **Content:** One row per student response with all survey sections

### MTF Benchmark Data (Required for comparisons)
- **Format:** CSV with columns: Type, 8th, 10th, 12th, variable, timeline, Year
- **Type values:** Total, Male, Female
- **Timeline values:** lifetime, past-12-months, past-30-days

### Configuration (User provides)
- School name
- Survey date(s)
- Grade levels included (default: 9th, 10th, 11th, 12th)
- MTF year for comparison (e.g., 2024)

## Survey Sections

1. **Demographics** — Sex, age, grade level, academic performance
2. **Alcohol Behaviors** — Age at first drink, frequency, quantity, 22 consequences
3. **Marijuana Behaviors** — Lifetime, annual, 30-day use; age at first use
4. **Other Substances** — 15 substance categories (lifetime use)
5. **Substance Use at School Events**
6. **Perceived Norms** — Estimated peer alcohol/drug use by grade
7. **School Environment** — 9 agree/disagree culture statements
8. **Alcohol/Drug Policies** — Policy and enforcement perceptions
9. **School Connectedness** — 8 belonging/engagement measures
10. **Family Factors** — Parental monitoring, family rules, substance use at home
11. **Friend & Peer Factors** — Peer influence, negative consequences from others' use
12. **Stress & Mental Health** — K10 Kessler scale (10 items)
13. **Reasons for Use** — Academic stress, social pressure, boredom, fitting in

## Output Formats

### Tables
- **Count & Percent** — Response counts with percentages, totals
- **By Grade** — Crosstab showing (N) and (%) per grade level
- **By Risk Category** — Filtered to non-drinker, low-risk, high-risk subsets
- **MTF Comparison** — School rates vs national norms with above/below indicators

### Charts
- Bar charts comparing school vs MTF rates by grade
- Stacked bar charts for agree/disagree distributions
- K10 score distributions by risk category

### Reports
- **Executive Summary** — Key findings, top-line statistics
- **Full Report** — All sections with analysis, charts, tables
- **Appendix A** — Basic counts & percentages by question
- **Appendix B** — Results by grade level
- **Appendix C** — Results by drinking risk category

## Key Calculations

### Risk Categories
- **Non-drinker:** Never consumed alcohol (age_at_first_drink contains "never" OR one_drink_past_12 == "No")
- **Low-risk drinker:** Drinks but not at high frequency/quantity
- **High-risk drinker:** days_5_or_more_drinks is NOT "None" AND NOT "0 days", OR how_often_1_drink is "Once a week or more" / "Every day or nearly every day" / "Twice a week"

### MTF Comparison Logic (4-Category System)
Uses **relative error margins** of the MTF value (not fixed percentage points):
- `error_5` = MTF_value × 0.05
- `error_7_5` = MTF_value × 0.075

| Comparison | Condition |
|------------|-----------|
| **Below** | school_pct < (MTF - error_5), OR MTF is 0 |
| **On par** | school_pct within MTF ± error_5 |
| **Exceeds** | school_pct > (MTF + error_5) AND ≤ (MTF + error_7_5) |
| **Significantly exceeds** | school_pct > (MTF + error_7_5) |

### K10 Scoring
Sum 10 items scored 1-5 (Never=1 to Always=5):
- **10-15:** Low distress
- **16-21:** Moderate distress
- **22-29:** High distress
- **30-50:** Very high distress

### Data Quality
- Flag and remove contradictory responses (e.g., "never used" + specific frequency)
- Report removal rate as percentage of total responses
- Cap calculated percentages at 100%

## Formatting Standards

### Tables
- Header row: White bold text on dark blue background
- Data rows: Alternating row shading (light lavender #ebebf8)
- Borders: Soft blue-lavender (#c0c0ee)
- Font: Georgia 9pt for data, 10pt for headers
- Percentages: 1 decimal place with % symbol

### Headings
- H1: Georgia 14pt bold, blue bottom border
- H2: Georgia 12pt bold
- H3: Georgia 11pt bold

### Charts
- School data in primary blue (#3333cc)
- MTF data in gray (#999999)
- Clear labels, gridlines, and legends

## Best Practices

1. Always validate data before analysis — check for expected columns, handle missing values
2. Use `safe_round()` for all percentage calculations to handle edge cases
3. Cap percentage sums at 100% to prevent rounding artifacts
4. When comparing to MTF, note that MTF only covers 8th, 10th, and 12th grades
5. Include sample sizes (N) alongside all percentages
6. Flag any sections with fewer than 5 respondents as potentially non-representative
7. Generate natural language bullet points for key findings, not just raw numbers

## Template Report Generation

The SUPP report uses a **docxtpl** Word template with 294 template variables, 609 dynamic
tables, and 27 embedded charts. See `TEMPLATE_VARS.md` for the complete catalog.

### Template Variable Categories
| Category | Count | Examples |
|----------|-------|---------|
| Config | 7 | `School`, `date_survey`, `N`, `prettygrades` |
| Charts | 27 | `alcol_melt_chart`, `K10_ALL_CHART`, `socialatmochart` |
| Bullets | 81 | `albullet1..3`, `maribullet1..3`, `highriskbulletsssss[0..22]` |
| Percentages | 53 | `percent_never_drank_allgrades`, `pressured_df_yes_perc` |
| Counts | 13 | `five_ormorecount`, `past30_yes_drankcount` |
| Perceived norms | 15 | `eighth_perceived_neverdrink_perc` |
| Strings | 9 | `bingedrinkstring`, `wheredrinkingstrings[0..4]` |
| Other scalars | 86 | `lifetime8th`, `totalhighrisk`, `avg_age_highrisk` |

### Dynamic Table DataFrames
Each survey question produces up to 3 DataFrame variants:
- **Base** (`answers[i]`) — All students, count & percent columns
- **By grade** (`answers_grade[i]`) — Columns: Response, then (N)/(%) per grade
- **By risk** (`answers_by_risk[i]`) — Columns: Response, then (N)/(%) per risk category

Major DataFrame families:
- `YNQ_answers` (34 items) — Yes/No substance use questions
- `neverQ_answers` (23 items) — Frequency response questions
- `AgreeDisagree_answers` (9 items) — School environment
- `Kessler_As` (10 items) — K10 mental health
- `ALLSUBSTANCE{grade}_{timeline}_MTF` — MTF comparison tables

### Report Sections
1. Cover → Introduction → Methodology
2. Executive Summary (key findings bullets)
3. Key Data Analysis: Alcohol, Cannabis, Vaping
4. MTF National Benchmark Comparisons (by grade × timeline × gender)
5. School Climate & Environment
6. Perceived Norms (perception vs reality)
7. Higher-Risk Drinkers Analysis
8. Attitudes & Policies
9. School Connectedness, Family Factors, Peer Factors
10. K10 Mental Health & Stress/Coping
11. Appendix A (all counts), B (by grade), C (by risk)

### Chart Specifications
- **Colors:** School data `#1A63CD`, MTF data `#C9D2DA`, Accent `#267EF2`
- **Type:** Horizontal bar charts for MTF comparisons, stacked bars for agree/disagree
- **Library:** matplotlib with seaborn styling
- **Output:** InlineImage objects for docxtpl embedding

## Usage with Anthropic Skills

Combine this skill with Anthropic document skills for full report generation:
- `xlsx` — Export analysis tables and charts to Excel workbooks
- `pptx` — Create executive summary presentations for school administrators
- `pdf` — Generate formatted PDF reports for distribution
- `docx` — Create Word documents matching the SUPP template format
