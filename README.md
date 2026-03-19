# UK Education Attainment Gap Analysis

Does ethnicity, gender, or deprivation predict your A-Level results in England? This analysis uses official UK Government data (Department for Education, 2007-2025) to find out.

## Why This Project

My MSc dissertation at Aston University examined ethnicity and academic outcomes using university-internal data. This is the redesign — using publicly available government datasets covering millions of students across England, with proper statistical analysis and machine learning.

I'm also a PSW graduate visa holder. The data shows that immigrant-background students often outperform White British students — yet the visa system pushes skilled graduates out. This analysis is both academic and personal.

## What I Found

### 1. The Ethnicity Gap Is Real and Persistent
A-Level attainment varies significantly by ethnicity. Chinese and Indian students consistently achieve the highest rates. Black Caribbean students face the largest gap.

### 2. Deprivation Amplifies Ethnicity Effects
Disadvantaged students in ALL ethnic groups perform worse — but the deprivation penalty is not equal. Some groups are doubly penalised by both ethnicity and poverty.

### 3. Gender Gap Varies by Ethnicity
Females outperform males across most ethnic groups, but the magnitude of the gender gap differs substantially between communities.

### 4. Immigration Background Is Not a Disadvantage
Indian, Chinese, and Black African students — communities with strong immigration ties — often outperform White British students. The UK education system benefits from immigrant talent.

### 5. 10-Year Policy Impact Is Mixed
Some attainment gaps have narrowed since 2007, but structural inequalities remain deeply entrenched despite government widening participation initiatives.

### 6. Predictive Model
Gradient Boosting and Random Forest models predict above/below average attainment using ethnicity, gender, and deprivation as features.

## Data Sources

| Dataset | Source | Records | Period |
|---|---|---|---|
| A-Level attainment by ethnicity, sex, disadvantage | DfE | 7,200 rows | 2020-2025 |
| Level 2/3 attainment by characteristics | DfE | 200,000+ rows | 2007-2024 |
| Disadvantage gap index (KS2) | DfE | Supplementary | 2007-2025 |

All data from [explore-education-statistics.service.gov.uk](https://explore-education-statistics.service.gov.uk)

## Charts Generated

| Chart | Shows |
|---|---|
| Ethnicity attainment trend | A-Level rates by major ethnic group (2020-2025) |
| Deprivation x ethnicity gap | Does poverty hit all groups equally? |
| Gender x ethnicity gap | Male vs female attainment by ethnic group |
| Immigration attainment | Immigrant-background communities vs White British |
| 10-year policy trend | Level 2 attainment by ethnicity (2007-2024) |
| Feature importance | What predicts attainment? (ML model) |

## Tools Used

| Tool | Purpose |
|---|---|
| Python | Analysis and ML |
| pandas | Data manipulation |
| matplotlib + seaborn | Visualisation |
| scikit-learn | Random Forest, Gradient Boosting |
| Jupyter | Notebook-based analysis |

## Project Structure

```
uk-education-attainment/
├── notebooks/
│   └── 01_education_attainment_analysis.ipynb    # Full analysis
├── data/
│   └── raw/                                       # DfE CSV downloads
│       ├── alevel_ethnicity.csv                   # A-Level data
│       ├── level23_ethnicity.csv                  # Level 2/3 data
│       ├── level23_characteristics.csv            # Characteristics breakdown
│       └── ks2_disadvantage_gap.csv               # KS2 gap index
├── charts/                                        # Generated visualisations
└── README.md
```

## How to Run

```bash
pip install pandas matplotlib seaborn scikit-learn jupyter
cd notebooks
jupyter notebook 01_education_attainment_analysis.ipynb
```

## Connection to My MSc Dissertation

My original dissertation (Aston University, 2024) used internal university data to study ethnicity effects on academic performance. This project extends that work to the national level using government open data, adds deprivation and gender as variables, includes a 10-year longitudinal view, and frames the findings in the context of UK immigration policy.
