"""
Regenerate all charts for the UK Education Attainment Gap Analysis.
Fixes: broken chart 05 (age type mismatch), axis labels, consistent styling.
"""
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import numpy as np
from sklearn.ensemble import GradientBoostingClassifier
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import LabelEncoder
import warnings
warnings.filterwarnings('ignore')
import matplotlib
matplotlib.rcParams['font.family'] = 'sans-serif'

# --- Consistent style ---
BLUE = '#1e3a5f'
RED = '#b91c1c'
GREEN = '#15803d'
ORANGE = '#d97706'
PURPLE = '#7c3aed'
TEAL = '#0d9488'
PINK = '#be185d'
BROWN = '#92400e'

plt.rcParams.update({
    'figure.facecolor': 'white',
    'axes.facecolor': 'white',
    'axes.edgecolor': '#e2e8f0',
    'axes.grid': True,
    'grid.color': '#f1f5f9',
    'grid.linewidth': 0.8,
    'font.size': 11,
    'axes.titlesize': 14,
    'axes.titleweight': 'bold',
    'axes.labelsize': 11,
})

# ============================================================
# LOAD DATA
# ============================================================
alevel = pd.read_csv('data/raw/alevel_ethnicity.csv')
alevel['year'] = alevel['time_period'].astype(str).str[:4].astype(int)

chars = pd.read_csv('data/raw/level23_characteristics.csv')
chars['year'] = chars['time_period'].astype(str).str[:4].astype(int)

major_groups = ['White', 'Asian or Asian British', 'Black or Black British',
                'Mixed Dual background', 'Any other ethnic group']

year_labels = {202021: '2020/21', 202122: '2021/22', 202223: '2022/23',
               202324: '2023/24', 202425: '2024/25'}

def get_attainment(df):
    """Calculate attainment rate from student counts."""
    num = pd.to_numeric(df.get('number_of_students_level3', 0), errors='coerce')
    den = pd.to_numeric(df.get('number_of_students_potential', 1), errors='coerce')
    return (num / den * 100).clip(0, 100)

# ============================================================
# CHART 1: Ethnicity Attainment Trend
# ============================================================
filtered = alevel[
    (alevel['sex'] == 'All') & (alevel['disadvantage'] == 'All') &
    (alevel['ethnicity_minor'].isin(major_groups)) & (alevel['sen'] == 'All')
].copy()
filtered['attainment_rate'] = get_attainment(filtered)
trend = filtered.groupby(['time_period', 'ethnicity_minor'])['attainment_rate'].mean().reset_index()

colors_map = {
    'White': BLUE, 'Asian or Asian British': ORANGE,
    'Black or Black British': GREEN, 'Mixed Dual background': PURPLE,
    'Any other ethnic group': RED
}

fig, ax = plt.subplots(figsize=(12, 7))
for eth in major_groups:
    data = trend[trend['ethnicity_minor'] == eth].sort_values('time_period')
    if not data.empty and data['attainment_rate'].sum() > 0:
        labels = [year_labels.get(tp, str(tp)) for tp in data['time_period']]
        ax.plot(labels, data['attainment_rate'], marker='o', linewidth=2.5,
                label=eth, color=colors_map.get(eth, 'grey'), markersize=7)

ax.set_xlabel('Academic Year')
ax.set_ylabel('Level 3 Attainment Rate (%)')
ax.set_title('A-Level Attainment Rate by Ethnicity (2020\u20132025)\nSource: Department for Education')
ax.legend(loc='lower right', frameon=True, fancybox=False, edgecolor='#e2e8f0')
ax.set_axisbelow(True)
plt.xticks(rotation=0)
plt.tight_layout()
plt.savefig('charts/01_ethnicity_attainment_trend.png', dpi=200, bbox_inches='tight')
plt.close()
print('  01 done')

# ============================================================
# CHART 2: Deprivation x Ethnicity Gap
# ============================================================
latest_year = alevel['time_period'].max()
dep = alevel[
    (alevel['time_period'] == latest_year) & (alevel['sex'] == 'All') &
    (alevel['sen'] == 'All') &
    (alevel['disadvantage'].isin(['Disadvantaged', 'Non-Disadvantaged'])) &
    (alevel['ethnicity_minor'].isin(major_groups))
].copy()
dep['attainment_rate'] = get_attainment(dep)

pivot = dep.pivot_table(index='ethnicity_minor', columns='disadvantage',
                        values='attainment_rate', aggfunc='mean')
if not pivot.empty and 'Non-Disadvantaged' in pivot.columns and 'Disadvantaged' in pivot.columns:
    pivot['gap'] = pivot['Non-Disadvantaged'] - pivot['Disadvantaged']
    pivot = pivot.sort_values('gap', ascending=False)

    fig, ax = plt.subplots(figsize=(12, 6))
    y_pos = range(len(pivot))
    ax.barh([i + 0.15 for i in y_pos], pivot['Non-Disadvantaged'], 0.3,
            label='Non-Disadvantaged', color=GREEN)
    ax.barh([i - 0.15 for i in y_pos], pivot['Disadvantaged'], 0.3,
            label='Disadvantaged', color=RED)
    ax.set_yticks(y_pos)
    ax.set_yticklabels(pivot.index)
    ax.set_xlabel('Attainment Rate (%)')
    ax.set_title(f'Deprivation Gap by Ethnicity ({year_labels.get(latest_year, latest_year)})\nDoes poverty hit all groups equally?')
    ax.legend(frameon=True, fancybox=False, edgecolor='#e2e8f0')
    ax.set_axisbelow(True)

    for i, (idx, row) in enumerate(pivot.iterrows()):
        gap = row.get('gap', 0)
        if pd.notna(gap) and gap > 0:
            x_pos = max(row.get('Non-Disadvantaged', 0), row.get('Disadvantaged', 0))
            ax.text(x_pos + 0.5, i, f'Gap: {gap:.1f}pp', va='center',
                    fontweight='bold', color='#334155', fontsize=9)

    plt.tight_layout()
    plt.savefig('charts/02_deprivation_ethnicity_gap.png', dpi=200, bbox_inches='tight')
    plt.close()
    print('  02 done')

# ============================================================
# CHART 3: Gender x Ethnicity Gap
# ============================================================
gender = alevel[
    (alevel['time_period'] == latest_year) & (alevel['sex'].isin(['Male', 'Female'])) &
    (alevel['disadvantage'] == 'All') & (alevel['sen'] == 'All') &
    (alevel['ethnicity_minor'].isin(major_groups))
].copy()
gender['attainment_rate'] = get_attainment(gender)

gender_pivot = gender.pivot_table(index='ethnicity_minor', columns='sex',
                                  values='attainment_rate', aggfunc='mean')
if not gender_pivot.empty:
    gender_pivot['gender_gap'] = gender_pivot.get('Female', 0) - gender_pivot.get('Male', 0)
    gender_pivot = gender_pivot.sort_values('gender_gap', ascending=True)

    fig, ax = plt.subplots(figsize=(12, 6))
    y_pos = range(len(gender_pivot))
    ax.barh([i + 0.15 for i in y_pos], gender_pivot['Female'], 0.3, label='Female', color=PINK)
    ax.barh([i - 0.15 for i in y_pos], gender_pivot['Male'], 0.3, label='Male', color=BLUE)
    ax.set_yticks(y_pos)
    ax.set_yticklabels(gender_pivot.index)  # Clean label, not "ethnicity_minor"
    ax.set_xlabel('Attainment Rate (%)')
    ax.set_title(f'Gender Gap in A-Level Attainment by Ethnicity ({year_labels.get(latest_year, latest_year)})')
    ax.legend(frameon=True, fancybox=False, edgecolor='#e2e8f0')
    ax.set_axisbelow(True)
    plt.tight_layout()
    plt.savefig('charts/03_gender_ethnicity_gap.png', dpi=200, bbox_inches='tight')
    plt.close()
    print('  03 done')

# ============================================================
# CHART 4: Immigration Background Comparison
# ============================================================
immigrant_groups = [
    'Asian or Asian British - Indian', 'Asian or Asian British - Pakistani',
    'Asian or Asian British - Bangladeshi', 'Asian or Asian British - Chinese',
    'Black or Black British - Black African', 'Black or Black British - Black Caribbean',
    'White', 'White - White British'
]
available = [g for g in immigrant_groups if g in alevel['ethnicity_minor'].unique()]

imm = alevel[
    (alevel['sex'] == 'All') & (alevel['disadvantage'] == 'All') &
    (alevel['sen'] == 'All') & (alevel['ethnicity_minor'].isin(available))
].copy()
imm['attainment_rate'] = get_attainment(imm)
imm_trend = imm.groupby(['time_period', 'ethnicity_minor'])['attainment_rate'].mean().reset_index()

imm_colors = {
    'Asian or Asian British - Indian': BLUE,
    'Asian or Asian British - Pakistani': ORANGE,
    'Asian or Asian British - Bangladeshi': TEAL,
    'Asian or Asian British - Chinese': RED,
    'Black or Black British - Black African': GREEN,
    'Black or Black British - Black Caribbean': BROWN,
    'White': '#94a3b8',
    'White - White British': PURPLE
}

fig, ax = plt.subplots(figsize=(13, 7))
for group in available:
    data = imm_trend[imm_trend['ethnicity_minor'] == group].sort_values('time_period')
    if not data.empty and data['attainment_rate'].sum() > 0:
        label = (group.replace('Asian or Asian British - ', '')
                      .replace('Black or Black British - ', '')
                      .replace('White - ', ''))
        labels = [year_labels.get(tp, str(tp)) for tp in data['time_period']]
        ax.plot(labels, data['attainment_rate'], marker='o', linewidth=2,
                label=label, color=imm_colors.get(group, 'grey'), markersize=6)

ax.set_xlabel('Academic Year')
ax.set_ylabel('Level 3 Attainment Rate (%)')
ax.set_title('A-Level Attainment: Immigrant-Background Communities vs White British (2020\u20132025)\nSource: Department for Education')
ax.legend(loc='best', frameon=True, fancybox=False, edgecolor='#e2e8f0', fontsize=9)
ax.set_axisbelow(True)
plt.tight_layout()
plt.savefig('charts/04_immigration_attainment.png', dpi=200, bbox_inches='tight')
plt.close()
print('  04 done')

# ============================================================
# CHART 5: 10-Year Policy Trend (FIXED - age type mismatch)
# ============================================================
eth_major = chars[
    (chars['characteristic_group'] == 'Ethnicity Major') &
    (chars['number_or_percentage'] == 'Percentage') &
    (chars['age'].astype(str) == '19')  # FIX: compare as strings
].copy()

eth_major['Level_2'] = pd.to_numeric(eth_major['Level_2'], errors='coerce')

fig, ax = plt.subplots(figsize=(13, 7))
plotted = False
for ethnicity in eth_major['characteristic'].unique():
    if ethnicity not in ['Attainment gap', 'Total', 'Unclassified']:
        data = eth_major[eth_major['characteristic'] == ethnicity].sort_values('year')
        data = data.dropna(subset=['Level_2'])
        if len(data) > 2:
            ax.plot(data['year'], data['Level_2'], marker='o', linewidth=2,
                    label=ethnicity, markersize=5)
            plotted = True

if plotted:
    ax.set_xlabel('Year')
    ax.set_ylabel('Level 2 Attainment by Age 19 (%)')
    ax.set_title('10+ Year Trend: Level 2 Attainment by Ethnicity (2004\u20132025)\nHas the government\'s widening participation push worked?')
    ax.legend(loc='lower right', frameon=True, fancybox=False, edgecolor='#e2e8f0')
    ax.set_axisbelow(True)
else:
    ax.text(0.5, 0.5, 'No data available for this view', ha='center', va='center',
            transform=ax.transAxes, fontsize=14, color='#94a3b8')
    ax.set_title('10+ Year Trend: Level 2 Attainment by Ethnicity')

plt.tight_layout()
plt.savefig('charts/05_10year_policy_trend.png', dpi=200, bbox_inches='tight')
plt.close()
print('  05 done')

# ============================================================
# CHART 6: Feature Importance (ML)
# ============================================================
ml_data = alevel[
    (alevel['sex'].isin(['Male', 'Female'])) &
    (alevel['disadvantage'].isin(['Disadvantaged', 'Non-Disadvantaged'])) &
    (alevel['sen'] == 'All') &
    (~alevel['ethnicity_minor'].isin(['All', 'Unknown ethnicity']))
].copy()
ml_data['attainment_rate'] = get_attainment(ml_data)
ml_data = ml_data.dropna(subset=['attainment_rate'])
ml_data = ml_data[ml_data['attainment_rate'] > 0]

if len(ml_data) > 50:
    median_rate = ml_data['attainment_rate'].median()
    ml_data['above_average'] = (ml_data['attainment_rate'] >= median_rate).astype(int)

    le_eth = LabelEncoder()
    le_sex = LabelEncoder()
    le_dis = LabelEncoder()
    ml_data['eth_encoded'] = le_eth.fit_transform(ml_data['ethnicity_minor'])
    ml_data['sex_encoded'] = le_sex.fit_transform(ml_data['sex'])
    ml_data['dis_encoded'] = le_dis.fit_transform(ml_data['disadvantage'])

    X = ml_data[['eth_encoded', 'sex_encoded', 'dis_encoded', 'year']]
    y = ml_data['above_average']
    X_train, X_test, y_train, y_test = train_test_split(X, y, test_size=0.25, random_state=42, stratify=y)

    gb = GradientBoostingClassifier(n_estimators=100, random_state=42, max_depth=5)
    gb.fit(X_train, y_train)
    accuracy = gb.score(X_test, y_test)

    importance = pd.DataFrame({
        'Feature': ['Ethnicity', 'Gender', 'Deprivation', 'Year'],
        'Importance': gb.feature_importances_
    }).sort_values('Importance', ascending=True)

    fig, ax = plt.subplots(figsize=(10, 5))
    bar_colors = [BLUE if imp < importance['Importance'].max() else RED
                  for imp in importance['Importance']]
    # Use consistent blue, highlight top in dark
    ax.barh(importance['Feature'], importance['Importance'], color=BLUE, height=0.5)
    ax.set_xlabel('Feature Importance')
    ax.set_title(f'What Predicts A-Level Attainment?\nGradient Boosting Feature Importance (Accuracy: {accuracy:.1%})')
    ax.set_axisbelow(True)

    for i, (_, row) in enumerate(importance.iterrows()):
        ax.text(row['Importance'] + 0.005, i, f"{row['Importance']:.1%}",
                va='center', fontweight='bold', fontsize=11)

    plt.tight_layout()
    plt.savefig('charts/06_feature_importance.png', dpi=200, bbox_inches='tight')
    plt.close()
    print(f'  06 done (model accuracy: {accuracy:.1%})')

print('\nAll education charts regenerated.')
