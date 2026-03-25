# SUPP Survey Reference

## Typeform Column Mapping

The Typeform CSV export uses verbose question text as column headers. After loading,
columns are renamed positionally to short internal names via `COLLIST` in `supp_processor.py`.
The trailing columns are always `Submitted At` and `Token`.

**Key internal column names for analysis:**
- `grade_level` — grade (e.g., "9th", "10th", "11th", "12th")
- `sex` — "Male" or "Female"
- `age_at_first_drink` — age or "I have never had a whole drink of alcohol"
- `how_often_1_drink` — drinking frequency (see sort_order1)
- `one_drink_past_12` — "Yes" or "No"
- `days_5_or_more_drinks` — binge indicator (see sort_order4)
- `age_first_time_pot` — age or "I have never used marijuana at all"
- `how_often_pot` — marijuana frequency (see sort_order5)
- `use_*_12` — past-12-months substance use (Yes/No)
- `use_*_30` — past-30-days substance use (Yes/No)
- `drinker_risk` — computed: "High-risk", "Low-risk", "Non-drinker", "Missing"
- K10 items: `tired`, `nervous`, `no_calm`, `hopless`, `restless`, `no_sit_still`, `depressed`, `effort`, `cheer`, `worthless`

## Survey Question Map

### Section 1: Demographics
| Internal Name | Question |
|----------|----------|
| sex | What is your sex? (Male / Female) |
| age | How old are you? |
| grade_level | What grade are you in? |
| grades_overall | How were your grades overall during the last school year? |

### Section 2: Alcohol Behaviors
| Internal Name | Question |
|----------|----------|
| age_at_first_drink | If you drink, how old were you when you had your first whole drink of alcohol? |
| where_first_use_alcohol | Where did you first use alcohol? |
| how_often_1_drink | How often do you typically have at least one whole drink of alcohol? |
| how_many_drinks | How many whole drinks of alcohol do you typically have when you drink? |
| one_drink_past_12 | Have you had at least 1 whole drink of alcohol during the PAST 12 MONTHS? |

### Section 2a: Alcohol Consequences (22 items)
"How many times has the following happened to you because of YOUR OWN alcohol or drug use?"

| Internal Name | Consequence |
|----------|------------|
| felt_or_tried_to_cut | Felt the need or tried to cut down |
| drinking_or_drug_problem | Thought you had a drinking or other drug problem |
| felt_guilty_about_your_drinking | Felt guilty about your drinking or other drug use |
| needed_a_drink_first_thing | Needed a drink or other drug first thing in the morning |
| used_alcohol_by_yourself | Used alcohol or other drugs while by yourself |
| get_drunk | Get drunk |
| do_something_you_later_felt_sorry_for | Do something you later felt sorry for |
| friend_fight_argument_bad_feelsing | Fight/argument/bad feelings with friend/dating partner |
| family_fight_argument_bad_feelsing | Fight/argument/bad feelings with family member |
| sexual_experience | Sexual experience you later wished didn't happen |
| perform_poorly_school_work | Miss or perform poorly at school, work, or athletic event |
| trouble_concentrating | Trouble concentrating on homework, class work, etc. |
| sick_vomit_hangover | Feel really sick, vomit, or have a hangover |
| drive | Drive a car or other vehicle after drinking |
| ride | Ride as passenger after driver had been drinking |
| trouble_school_home | Get in trouble at school or home |
| hang_out_bad_influence | Hang out with people who are a bad influence |
| forget_what_happened | Forget what happened while drinking (Black out) |
| pass_out | Pass out because of drinking |
| get_hurt | Get hurt or injured |
| drink_before_school | Have something to drink before or during school hours |
| try_other_substances | Try other substances you wouldn't have tried otherwise |
| drank_more | Drank more than you intended |

### Section 2b: Past 30 Days Alcohol
| Internal Name | Question |
|----------|----------|
| one_drink_past_30_days | Days with at least one whole drink in past 30 days |
| drinks_in_a_row_30_days | Greatest number of whole drinks consumed in a row in past 30 days |
| days_5_or_more_drinks | Days with 5+ whole drinks in a row in past 30 days |
| drunk_30_days | Times been or felt drunk in past 30 days |
| where_do_you_drink | Where do you do most of your drinking? |

### Section 3: Marijuana Behaviors
| Internal Name | Question |
|----------|----------|
| age_first_time_pot | How old were you the first time you used marijuana? |
| how_often_pot | How often do you use marijuana? |

### Section 4: Substance Use (Past 12 Months & Past 30 Days)
Each substance has two columns: `use_[substance]_12` and `use_[substance]_30` (Yes/No).

| Suffix | Substance |
|--------|-----------|
| alcohol | Alcohol (at least one whole drink) |
| cigarettes | Cigarettes |
| vapes | Electronic vaporizers (e-cigs, vape pens, JUUL, etc.) |
| smokeless_tobacco | Smokeless tobacco (chew, snuff, plug, dip, ZYN) |
| pot | Marijuana (cannabis, weed, hash, edibles, etc.) |
| ecstasy | Ecstasy (Molly, E, MDMA) |
| heroin | Heroin |
| narcotics | Narcotics NOT prescribed (OxyContin, Vicodin, Dilaudid) |
| amphetamines | Amphetamines NOT prescribed (Ritalin, Adderall, diet pills) |
| sedatives | Sedatives/Sleeping Pills NOT prescribed (Valium, Xanax, etc.) |
| methamphetamine | Methamphetamine (speed, crystal meth, ice) |
| cocaine | Cocaine |
| lsd | LSD (acid) |
| hallucinogen | Hallucinogens other than LSD (psilocybin, mushrooms, peyote) |
| inhalants | Inhalants (poppers, nitrous, glue, paint thinner) |
| steroids | Steroids NOT prescribed |
| cold_medicine | Over-the-counter cold medicines (non-medicinal use) |

### Section 5: Substance Use at School Events
| Internal Name | Question |
|----------|----------|
| school_dance_use | Have you ever used alcohol, marijuana, or other drugs before or during a school event? |

### Section 6: Attitudes & Perceived Norms
| Internal Name | Question |
|----------|----------|
| attitude | Which statement best represents YOUR OWN attitude towards teenagers drinking? |
| one_2_friend / three_4_friend / five_friend / got_drunk_friend | Friend drinking frequency norms |
| harm_1_2_drinks / harm_1_2_daily / harm_3_4_daily / harm_5 | Perceived harm levels |
| pot_friend_once / pot_friend_occasion / pot_friend_regular | Friend marijuana norms |

### Section 7: School Environment (Agree/Disagree)
| Internal Name | Statement |
|----------|-----------|
| social_atmosphere | The social atmosphere at this school encourages alcohol use |
| pressured | Students at this school feel pressured to drink alcohol |
| easy_to_make_friends | It is easy to make friends without drinking alcohol |
| students_assume | Students assume that everyone drinks alcohol |
| own_decisions | Students will let you make your own decisions about drinking |
| respect | Students respect you if you do not drink alcohol |
| fit_in | Students who abstain from alcohol do not fit in |
| light_drinkers | Most students don't drink or are light drinkers |
| binge_okay | Students think that heavy or binge drinking is okay |

### Section 8: Perceived Norms (Quantitative)
| Internal Name | Question |
|----------|----------|
| drink_how_often_your_grade / drink_how_often_seniors | How often do students drink? |
| how_many_your_grade / how_many_seniors | How many drinks do students have? |
| one_drink_30_days_your_school | Past 30 days frequency estimate |
| one_drink_30_days_your_grade / one_drink_30_days_seniors | By grade estimate |
| marijuana_your_grade / marijuana_seniors | Marijuana use estimates |
| percent_ever_used | % who used at school events estimate |

### Section 9: Policies
| Internal Name | Statement |
|----------|-----------|
| thoughts_alcohol_policies / thoughts_enforcement | Policies are... (Too lenient / About right / Too strict / Don't know) |
| thoughts_other_alcohol_policies / thoughts_other_enforcement | Enforcement of policies |

### Section 10: School Connectedness
| Internal Name | Statement |
|----------|-----------|
| valued | I feel valued as a person at my school |
| academics_to_me | Academics are important to me |
| teachers | I feel that teachers and school staff care about me |
| sincere_effort | My school is making a sincere effort to address substance use |
| care | Most students care about what happens at this school |
| academics_to_others | Academics are important to other students |
| adult_trust | I have at least one adult I trust at my school |
| school_12 | How you felt about going to school in the past 12 months |

### Section 11: Family Factors
| Internal Name | Statement |
|----------|-----------|
| parent_interest | My parent/guardian shows interest in my schoolwork |
| clear_rules | The rules in my family are clear |
| call | My parent/guardian wants me to call if I'm going to be late |
| come_home | My parent/guardian would know if I did not come home on time |
| clear_rules_alcohol | My family has clear rules about alcohol and drug misuse |
| not_home | My parent/guardian knows how to contact me when away |
| use_parent_knowing | Used alcohol/drugs WITH parent knowing |
| use_parent_not_knowing | Used alcohol/drugs WITHOUT parent knowing |
| meals | How often does your family eat at least one meal together? |

### Section 12: Friend & Peer Factors
| Internal Name | Statement |
|----------|-----------|
| friend_problem | Do you have a friend with an alcohol/drug problem? |
| talking_parents | Comfortable talking to parent about YOUR OWN use? |
| talking_parents_about_friend | Comfortable talking to parent about A FRIEND's problem? |
| classmates_pressured | Classmates feel pressured to use at non-school events |
| parents_present_visible | When classmates host parties, parents are present |

### Section 12a: Negative Peer Consequences
"Because of SOMEONE ELSE'S drinking/drug use..."

| Internal Name | Consequence |
|----------|------------|
| taking_concerned_other | Talked to someone about their use |
| taking_care_drank_too_much_other | Took care of someone who drank too much |
| taking_care_consumed_too_much_other | Took care of someone who used too much marijuana |
| emergency_medical_care_other | Took someone for emergency medical care |
| in_trouble_other | Got in trouble |
| feeling_insulted_other | Felt insulted or humiliated |
| verbal_argument_other | Engaged in a verbal argument |
| pushed_hit_other | Was pushed, hit, assaulted, or in a fight |
| sexual_experience_other | Sexual experience you wished didn't happen |
| witness_vomit_other | Witnessed someone vomiting |
| unsafe_other | Felt unsafe |
| sleep_difficulty_other | Had difficulty sleeping |
| homework_difficulty_other | Had difficulty doing homework |

### Section 13: K10 Mental Health (Past 4 Weeks)
"During the past 4 weeks, how often did you feel..."

| Internal Name | Item |
|----------|------------|
| tired | Tired out for no good reason |
| nervous | Nervous |
| no_calm | So nervous that nothing could calm you down |
| hopless | Hopeless |
| restless | Restless or fidgety |
| no_sit_still | So restless you could not sit still |
| depressed | Depressed |
| effort | That everything was an effort |
| cheer | So sad that nothing could cheer you up |
| worthless | Worthless |

Response scale: Never (1), Rarely (2), Sometimes (3), Often (4), Always (5)

Score ranges: 10-15 Low, 16-21 Moderate, 22-29 High, 30-50 Very High

### Section 14: Stress & Coping
| Internal Name | Stressor |
|----------|----------|
| academic_stress | Academic pressure |
| parent_stress | Parental pressure |
| college_stress | College expectations |
| overlooked_stress | Being overlooked |
| expectations_stress | My own expectations |
| boredom_stress | Boredom |
| fitting_in_stress | Fitting in socially |

Coping substance columns: `coping_alcohol`, `coping_cigarettes`, `coping_vape`,
`coping_smokeless_tobacco`, `coping_marijuana`, `coping_molly`, `coping_heroin`,
`coping_narcotics`, `coping_amphetamines`, `coping_sedatives`, `coping_meth`,
`coping_cocaine`, `coping_lsd`, `coping_hallucinogens`, `coping_inhalant`,
`coping_steroid`, `coping_cold_med`

## MTF Benchmark Data

### CSV Structure
- **Columns:** `Type, 8th, 10th, 12th, variable, timeline, Year`
- **Type values:** Total, Male, Female
- **Timeline values:** lifetime, past-12-months, past-30-days
- **Missing data:** Represented as `â` (mojibake from em-dash `—` read as latin-1)

### MTF Variable Names (exact values in CSV)
| variable | Timelines in 2024 data |
|----------|----------------------|
| Alcohol | lifetime, past-12-months, past-30-days |
| Amphetamines | lifetime, past-12-months, past-30-days |
| Cigarettes | lifetime, past-30-days |
| Cocaine | lifetime, past-12-months, past-30-days |
| Ecstasy | lifetime, past-12-months, past-30-days |
| Hallucinogens | lifetime, past-12-months, past-30-days |
| Heroin | lifetime, past-12-months, past-30-days |
| Inhalants | lifetime, past-12-months, past-30-days |
| LSD | lifetime, past-12-months, past-30-days |
| Marijuana | lifetime, past-12-months, past-30-days |
| Methamphetamine | lifetime, past-12-months, past-30-days |
| Narcotics | lifetime, past-12-months, past-30-days |
| Steroids | lifetime, past-12-months, past-30-days |
| Vaping Nicotine | lifetime, past-12-months, past-30-days |
| Any Vaping | lifetime, past-12-months, past-30-days |

**Note:** Narcotics data only available for 12th grade (8th and 10th show `â`).
Smokeless tobacco and Sedatives are NOT in the MTF 2024 data.

### MTF Comparison Algorithm
Uses **relative error margins** (not fixed percentage-point differences):
- `error_5` = MTF_value × 0.05
- `error_7_5` = MTF_value × 0.075

| Comparison | Condition |
|------------|-----------|
| **Below** | school_pct < (MTF - error_5), OR MTF is 0 |
| **On par** | school_pct within MTF ± error_5 |
| **Exceeds** | school_pct > (MTF + error_5) AND ≤ (MTF + error_7_5) |
| **Significantly exceeds** | school_pct > (MTF + error_7_5) |

Grades available: 8th, 10th, 12th
Gender breakdowns: Total, Male, Female

## Word Template Structure

The production report uses a **docxtpl** Word template (`2026 SUPP Template Report - Reformatted_v2.docx`).

**Template Stats:**
- 3,189 paragraphs | 618 tables (609 dynamic, 9 static)
- 294 unique `{{variable}}` placeholders
- 599 Jinja2 control blocks (`{%tr for ... %}` table row loops)
- 27 embedded chart placeholders (InlineImage objects)

**Variable categories:** Config (7), Charts (27), Bullets (81), Percentages (53),
Counts (13), Perceived Norms (15), Strings (9), Age (3), Other Scalars (86).

**Dynamic tables** use `{%tr for index, row in DATAFRAME.iterrows() %}` with 3 variants
per question: base (all students), by_grade, and by_risk.

**Full catalog:** See `TEMPLATE_VARS.md` for the complete variable list, DataFrame
families, and report section structure.

### Chart Color Palette
| Color | Hex | Usage |
|-------|-----|-------|
| Primary Blue | `#1A63CD` | School data bars |
| Light Gray | `#C9D2DA` | MTF comparison bars |
| Accent Blue | `#267EF2` | Highlights |
| Table Header | Dark blue | White bold text |
| Alt Row | `#ebebf8` | Light lavender shading |
| Border | `#c0c0ee` | Soft blue-lavender |
