import streamlit as st
import pandas as pd
import numpy as np
import re
import io
import string
from scipy.stats import t
import statsmodels.api as sm
from statsmodels.formula.api import ols
from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side

# ==============================================================================
# DATABASE OF 20 HIGH-STANDARD 2-YEAR/MULTI-YEAR ACADEMIC DISCUSSION TEMPLATES
# ==============================================================================
ACADEMIC_TEMPLATES_2Y = {
    "temp_1_both_sig_consistent": (
        "The evaluation of **{parameter}** revealed highly significant treatment effects (Year 1: {p_val_1}, Year 2: {p_val_2}) "
        "across both seasons (as detailed in **{table_label}**). Genotype `{top_g_1}` led the performance in the first season "
        "with a mean of {top_val_1}^{top_let_1}, while `{top_g_2}` maintained a superior statistical position in the second "
        "season ({top_val_2}^{top_let_2}). This highly consistent behavioral ranking across years indicates a robust genetic "
        "control over `{parameter}`."
    ),
    "temp_2_gxe_shifts": (
        "Statistical analysis of **{parameter}** indicated a highly significant treatment effect in both years ({p_val_1} and {p_val_2}); "
        "however, prominent seasonal shifts in genotype rankings were observed (Table **{table_label}**). In the first year, `{top_g_1}` "
        "yielded the maximum value ({top_val_1}^{top_let_1}), whereas `{top_g_2}` climbed to the statistical apex in the second "
        "season ({top_val_2}^{top_let_2}). These rank changes reflect a strong Genotype × Year (G×E) environmental interaction."
    ),
    "temp_3_exceptional_stability": (
        "Regarding **{parameter}**, genotype `{top_g_1}` demonstrated exceptional seasonal stability across both evaluation seasons, "
        "consistently occupying the leading statistical group (as shown in **{table_label}**). It recorded `{top_val_1}^{top_let_1}` "
        "in Year 1 and `{top_val_2}^{top_let_2}` in Year 2, resulting in the highest pooled average of `{pooled_max_val}`. This genotype "
        "represents a highly buffered candidate for environmental variation."
    ),
    "temp_4_divergent_significance_levels": (
        "The parameter **{parameter}** exhibited divergent significance profiles between the two trial seasons (as presented in **{table_label}**). "
        "The genotype effect was marginally significant in Year 1 ({p_val_1}), but became highly significant in Year 2 ({p_val_2}). "
        "Under Year 2, `{top_g_2}` established clear statistical superiority ({top_val_2}^{top_let_2}) compared to the control genotype, "
        "suggesting that seasonal variations amplified treatment expression in the second year."
    ),
    "temp_5_uniformly_nonsignificant": (
        "The genotypes evaluated presented a highly uniform and stable response for **{parameter}** across both trial years, "
        "resulting in nonsignificant treatment effects during both seasons ({p_val_1} and {p_val_2}) (Table **{table_label}**). "
        "The pooled mean settled at `{pooled_gm}`, with treatment means showing minor fluctuations around this baseline, indicating "
        "that `{parameter}` expression is highly preserved across both environments."
    ),
    "temp_6_pooled_superiority": (
        "On a pooled basis, the average mean of **{parameter}** across both years was `{pooled_gm}` (as summarized in **{table_label}**). "
        "Genotype `{top_g_pooled}` exhibited the highest pooled performance of `{pooled_max_val}`, demonstrating its overall superiority "
        "across environments compared to `{low_g_pooled}`, which occupied the absolute lower limit of `{pooled_min_val}`."
    ),
    "temp_7_elite_recommendation": (
        "Evaluating **{parameter}** across both years provides a strong basis for elite genotype selection (as detailed in **{table_label}**). "
        "Based on the multi-year pooled data, `{top_g_pooled}` is recommended as the prime candidate for optimizing this trait. "
        "It maintained statistically superior performance in both Years 1 and 2, showing low sensitivity to seasonal fluctuations."
    ),
    "temp_8_seasonal_environmental_shift": (
        "A highly pronounced environmental shift was observed for **{parameter}** between the two trial years, as shown in **{table_label}**. "
        "The grand mean increased/decreased substantially from `{gm_1}` in Year 1 to `{gm_2}` in Year 2. Despite this environmental shift, "
        "genotype rankings remained relatively stable, with `{top_g_pooled}` maintaining the highest overall pooled value of `{pooled_max_val}`."
    ),
    "temp_9_progressive_upward_trend_consistent": (
        "Concerning the progressive changes in **{base_name}**, a consistent, time-dependent progressive increase trend was observed "
        "across both seasons (as detailed in **{table_label}**). The pooled grand mean increased from `{first_gm}` at `{first_day}` to "
        "`{last_gm}` by `{last_day}`. Highly significant treatment effects were observed at `{last_day}` in both years, with `{top_g}` "
        "consistently showing the highest value on the final day, confirming a highly parallel temporal progression."
    ),
    "temp_10_progressive_downward_trend_consistent": (
        "For **{base_name}**, a systematic progressive decline was observed across both experimental years, as summarized in **{table_label}**. "
        "The pooled averages fell from `{first_gm}` at `{first_day}` to `{last_gm}` by `{last_day}`. Treatment differences on `{last_day}` "
        "were highly significant in both seasons ({last_p_1} and {last_p_2}), where `{top_g}` successfully minimized the rate of decline "
        "compared to `{low_g}`."
    ),
    "temp_11_divergent_seasonal_decline_rates": (
        "The progressive decline rates for **{base_name}** exhibited divergent seasonal behavior, as detailed in **{table_label}**. "
        "While the grand mean declined sharply by `{decline_1}%` in Year 1, it fell by a more moderate `{decline_2}%` in Year 2. "
        "This indicates that the seasonal conditions of Year 2 were more favorable for mitigating the degradation of this trait."
    ),
    "temp_12_late_onset_divergence_both_years": (
        "The progressive evaluation of **{base_name}** revealed a consistent late-onset treatment divergence across both seasons. "
        "While genotype effects on `{first_day}` were nonsignificant in both years, they became highly significant by `{last_day}` "
        "({last_p_1} and {last_p_2}) (Table **{table_label}**), with `{top_g}` establishing a statistically superior position in both years."
    ),
    "temp_13_early_divergence_late_convergence": (
        "For the progressive trend of **{base_name}**, early treatment differences observed on `{first_day}` converged over time, "
        "becoming nonsignificant by `{last_day}` across both seasons (as shown in **{table_label}**). This multi-year convergence "
        "suggests that long-term environmental exposure overrides initial genotype-specific advantages."
    ),
    "temp_14_progressive_stabilization_both_years": (
        "Regarding the temporal progression of **{base_name}**, a progressive stabilization pattern was observed across both trial seasons "
        "(Table **{table_label}**). The values altered rapidly from `{first_day}` to `{mid_day}`, but plateaued by `{last_day}`, where "
        "genotype `{top_g}` consistently maintained its statistical lead in both years."
    ),
    "temp_15_storage_decay_multiyear": (
        "Decay progression for **{base_name}** escalated over the storage intervals across both years, rising from `{first_gm}` to `{last_gm}` "
        "(as presented in **{table_label}**). Genotypes significantly influenced decay development by `{last_day}` in both years, "
        "with `{top_g}` consistently showing the highest decay suppression compared to `{low_g}`."
    ),
    "temp_16_experimental_precision_multiyear": (
        "The experimental precision of the trials was highly acceptable across both seasons, as indicated by stable Coefficients "
        "of Variation (CV: {cv_1}% and {cv_2}%) and low Standard Errors of the Mean (SEm: {sem_1} and {sem_2}) (Table **{table_label}**). "
        "This consistent precision validates the efficacy of the RCBD blocking factor in controlling background noise."
    ),
    "temp_17_pooled_parity_groups": (
        "The post-hoc letter displays on a seasonal basis, combined with the pooled averages, resolved genotypes into clear performance groups "
        "(as detailed in **{table_label}**). Although `{top_g_pooled}` had the highest overall pooled value of `{pooled_max_val}`, its "
        "statistical parity with `{at_par}` in both years suggests these genotypes are highly viable, equivalent alternatives."
    ),
    "temp_18_pooled_grand_mean_stability": (
        "The pooled grand mean for **{parameter}** settled at `{pooled_gm}` (as shown in **{table_label}**). The relatively low "
        "fluctuation in grand means between Year 1 (`{gm_1}`) and Year 2 (`{gm_2}`) indicates that this trait is highly stable "
        "and buffered against seasonal climatic differences."
    ),
    "temp_19_biochemical_seasons_multiyear": (
        "Evaluation of the biochemical parameter **{parameter}** showed significant treatment variations in both seasons (Table **{table_label}**). "
        "Genotype `{top_g_pooled}` exhibited outstanding nutrient/metabolite stability, recording high performance in both years "
        "and yielding the highest pooled average of `{pooled_max_val}`, proving its biochemical superiority."
    ),
    "temp_20_final_validation_recommendation": (
        "In conclusion, the multi-year validation of **{parameter}** provides highly reliable evidence for germplasm recommendations "
        "(as presented in **{table_label}**). Genotype `{top_g_pooled}` combined high pooled performance (`{pooled_max_val}`) with robust "
        "seasonal stability, establishing its selection as the premier candidate for the target environment."
    )
}

# --- Parameter Grouping Engine for Time-Series/Trend-Line Analysis ---
def group_parameters(params):
    """
    Groups parameters measured over multiple days/intervals (e.g., PLWD2, PLWD4 -> PLWD)
    to allow unified trend-line explanations under a single topic.
    """
    pattern = re.compile(r"(.*?)(\d+)\s*(dat|das|days|day|d)?$", re.IGNORECASE)
    groups = {}
    for p in params:
        match = pattern.search(p.strip())
        if match:
            base = match.group(1).strip()
            base = re.sub(r"[\s\-\_\(\)]+$", "", base).strip().capitalize()
            day_num = int(match.group(2))
            day_str = f"Day {day_num}"
            if not base:
                base = "Parameter"
            if base not in groups:
                groups[base] = []
            groups[base].append((p, day_num, day_str))
        else:
            base = p.strip().capitalize()
            if base not in groups:
                groups[base] = []
            groups[base].append((p, 0, ""))
            
    for base in groups:
        groups[base].sort(key=lambda x: x[1])
    return groups

# --- Dynamic Academic Explanation Selection Engine ---
def generate_trend_explanation_2y(base_name, items, results_1, results_2, year1_lbl, year2_lbl, table_label):
    """
    Generates a publication-grade paragraph explaining the temporal trend of a group of variables
    across two years using templates 9, 10, 11, 12, 13, 14, or 15 dynamically based on statistical profile.
    """
    first_item = items[0]
    last_item = items[-1]
    
    first_param, _, first_day_str = first_item
    last_param, _, last_day_str = last_item
    
    p_last_1 = results_1[last_param]
    p_last_2 = results_2[last_param]
    
    first_gm = (results_1[first_param]["gm"] + results_2[first_param]["gm"]) / 2
    last_gm = (p_last_1["gm"] + p_last_2["gm"]) / 2
    direction = "upward" if last_gm >= first_gm else "downward"
    
    sorted_last_1 = sorted(p_last_1["means"].items(), key=lambda x: x[1], reverse=True)
    top_last_1, top_val_last_1 = sorted_last_1[0]
    low_last_1, low_val_last_1 = sorted_last_1[-1]
    top_let_last_1 = p_last_1["means_str"][top_last_1].replace(f"{top_val_last_1:.2f}", "")
    
    sorted_last_2 = sorted(p_last_2["means"].items(), key=lambda x: x[1], reverse=True)
    top_last_2, top_val_last_2 = sorted_last_2[0]
    top_let_last_2 = p_last_2["means_str"][top_last_2].replace(f"{top_val_last_2:.2f}", "")
    
    # Template Selection for 2-Year Trend Line
    if "decay" in base_name.lower() or "rot" in base_name.lower():
        # Template 15: Storage Pathology/Decay Progression Multiyear
        template = ACADEMIC_TEMPLATES_2Y["temp_15_storage_decay_multiyear"]
        return template.format(
            base_name=base_name, first_gm=f"{first_gm:.2f}", last_gm=f"{last_gm:.2f}",
            top_g=top_last_1, low_g=low_last_1, table_label=table_label
        )
    elif results_1[first_param]["p_val"] >= 0.05 and p_last_1["p_val"] < 0.05 and results_2[first_param]["p_val"] >= 0.05 and p_last_2["p_val"] < 0.05:
        # Template 12: Late-Onset Treatment Divergence Both Years
        template = ACADEMIC_TEMPLATES_2Y["temp_12_late_onset_divergence_both_years"]
        return template.format(
            base_name=base_name, first_day=first_day_str, last_day=last_day_str,
            last_p_1=f"p = {p_last_1['p_val']:.4f}", last_p_2=f"p = {p_last_2['p_val']:.4f}",
            top_g=top_last_1, table_label=table_label
        )
    elif p_last_1["p_val"] >= 0.05 and p_last_2["p_val"] >= 0.05:
        # Template 13: Strictly Uniform Trend
        template = ACADEMIC_TEMPLATES_2Y["temp_13_early_divergence_late_convergence"]
        return template.format(
            base_name=base_name, first_day=first_day_str, last_day=last_day_str,
            table_label=table_label
        )
    elif direction == "upward":
        # Template 9: Progressive Upward Trend Consistent
        template = ACADEMIC_TEMPLATES_2Y["temp_9_progressive_upward_trend_consistent"]
        first_sig_txt = "significant" if results_1[first_param]["p_val"] < 0.05 else "nonsignificant"
        last_sig_txt = f"p = {p_last_1['p_val']:.4f}"
        return template.format(
            base_name=base_name, first_gm=f"{first_gm:.2f}", last_gm=f"{last_gm:.2f}",
            first_day=first_day_str, last_day=last_day_str, top_g=top_last_1, table_label=table_label
        )
    else:
        # Template 10: Progressive Downward Trend Consistent
        template = ACADEMIC_TEMPLATES_2Y["temp_10_progressive_downward_trend_consistent"]
        return template.format(
            base_name=base_name, first_gm=f"{first_gm:.2f}", last_gm=f"{last_gm:.2f}",
            first_day=first_day_str, last_day=last_day_str, last_p_1=f"p = {p_last_1['p_val']:.4f}",
            last_p_2=f"p = {p_last_2['p_val']:.4f}", top_g=top_last_1, low_g=low_last_1, table_label=table_label
        )

def generate_single_explanation_2y(param_name, p_data_1, p_data_2, year1_lbl, year2_lbl, table_label):
    """
    Selects and renders the best matching academic templates for isolated parameters
    based on the two-season significance levels, stability and pooled ratings.
    """
    p_val_1 = p_data_1["p_val"]
    p_val_2 = p_data_2["p_val"]
    
    sorted_1 = sorted(p_data_1["means"].items(), key=lambda x: x[1], reverse=True)
    top_1, top_val_1 = sorted_1[0]
    top_let_1 = p_data_1["means_str"][top_1].replace(f"{top_val_1:.2f}", "")
    
    sorted_2 = sorted(p_data_2["means"].items(), key=lambda x: x[1], reverse=True)
    top_2, top_val_2 = sorted_2[0]
    top_let_2 = p_data_2["means_str"][top_2].replace(f"{top_val_2:.2f}", "")
    
    # Extract pooled means
    pooled_means = {g: (p_data_1["means"].get(g, 0.0) + p_data_2["means"].get(g, 0.0)) / 2 for g in p_data_1["means"]}
    sorted_pooled = sorted(pooled_means.items(), key=lambda x: x[1], reverse=True)
    top_pooled, top_val_pooled = sorted_pooled[0]
    low_pooled, low_val_pooled = sorted_pooled[-1]
    
    at_par_list = []
    for genotype, val in sorted_1[1:]:
        let = p_data_1["means_str"][genotype].replace(f"{val:.2f}", "")
        if top_let_1 and let and any(char in top_let_1 for char in let):
            at_par_list.append(f"`{genotype}`")
    at_par_str = ", ".join(at_par_list) if at_par_list else "no other genotypes"
    
    # Template Selections
    if p_val_1 >= 0.05 and p_val_2 >= 0.05:
        # Template 5: Uniformly Nonsignificant
        template = ACADEMIC_TEMPLATES_2Y["temp_5_uniformly_nonsignificant"]
        return template.format(
            parameter=param_name, p_val_1=f"p = {p_val_1:.4f}", p_val_2=f"p = {p_val_2:.4f}",
            pooled_gm=f"{(p_data_1['gm'] + p_data_2['gm'])/2:.2f}", table_label=table_label
        )
    elif p_val_1 < 0.05 and p_val_2 >= 0.05:
        # Template 4: Divergent Significance Levels
        template = ACADEMIC_TEMPLATES_2Y["temp_4_divergent_significance_levels"]
        return template.format(
            parameter=param_name, p_val_1=f"p = {p_val_1:.4f}", p_val_2=f"p = {p_val_2:.4f}",
            top_g_2=top_2, top_val_2=f"{top_val_2:.2f}", top_let_2=top_let_2, table_label=table_label
        )
    elif top_1 == top_2 and p_val_1 < 0.05 and p_val_2 < 0.05:
        # Template 3: Exceptional Stability
        template = ACADEMIC_TEMPLATES_2Y["temp_3_exceptional_stability"]
        return template.format(
            parameter=param_name, top_g_1=top_1, top_val_1=f"{top_val_1:.2f}", top_let_1=top_let_1,
            top_val_2=f"{top_val_2:.2f}", top_let_2=top_let_2, pooled_max_val=f"{top_val_pooled:.2f}",
            table_label=table_label
        )
    elif "sugar" in param_name.lower() or "acid" in param_name.lower() or "nutrient" in param_name.lower() or "tss" in param_name.lower():
        # Template 19: Biochemical Seasons Multiyear
        template = ACADEMIC_TEMPLATES_2Y["temp_19_biochemical_seasons_multiyear"]
        return template.format(
            parameter=param_name, top_g_pooled=top_pooled, pooled_max_val=f"{top_val_pooled:.2f}",
            table_label=table_label
        )
    elif p_val_1 < 0.05 and p_val_2 < 0.05 and top_1 != top_2:
        # Template 2: GxE rank shifts
        template = ACADEMIC_TEMPLATES_2Y["temp_2_gxe_shifts"]
        return template.format(
            parameter=param_name, p_val_1=f"p = {p_val_1:.4f}", p_val_2=f"p = {p_val_2:.4f}",
            top_g_1=top_1, top_val_1=f"{top_val_1:.2f}", top_let_1=top_let_1,
            top_g_2=top_2, top_val_2=f"{top_val_2:.2f}", top_let_2=top_let_2, table_label=table_label
        )
    else:
        # Template 1: Highly Consistent Both Seasons
        template = ACADEMIC_TEMPLATES_2Y["temp_1_both_sig_consistent"]
        return template.format(
            parameter=param_name, p_val_1=f"p = {p_val_1:.4f}", p_val_2=f"p = {p_val_2:.4f}",
            top_g_1=top_1, top_val_1=f"{top_val_1:.2f}", top_let_1=top_let_1,
            top_g_2=top_2, top_val_2=f"{top_val_2:.2f}", top_let_2=top_let_2, table_label=table_label
        )

# --- Statistical Calculation Engine ---
def get_signif_code_val(p):
    if is_nan_val(p): return "ns"
    if p < 0.001: return "***"
    elif p < 0.01: return "**"
    elif p < 0.05: return "*"
    elif p < 0.1: return "."
    else: return "ns"

def is_nan_val(val):
    try: return np.isnan(float(val))
    except: return False

def get_cld_letters(means_dict, lsd):
    sorted_means = sorted(means_dict.items(), key=lambda x: x[1], reverse=True)
    n = len(sorted_means)
    groups = []
    
    for i in range(n):
        for j in range(i, n):
            is_group = True
            for k1 in range(i, j + 1):
                for k2 in range(k1, j + 1):
                    if abs(sorted_means[k1][1] - sorted_means[k2][1]) > lsd:
                        is_group = False
                        break
                if not is_group:
                    break
            if is_group:
                groups.append(set(range(i, j + 1)))
                
    maximal_groups = []
    for g in groups:
        is_subset = False
        for other_g in groups:
            if g != other_g and g.issubset(other_g):
                is_subset = True
                break
        if not is_subset and g not in maximal_groups:
            maximal_groups.append(g)
            
    maximal_groups.sort(key=lambda g: min(g))
    
    treatment_letters = {t[0]: "" for t in sorted_means}
    alphabet = string.ascii_lowercase
    
    for letter_idx, g in enumerate(maximal_groups):
        letter = alphabet[letter_idx % len(alphabet)]
        for idx in g:
            t_name = sorted_means[idx][0]
            treatment_letters[t_name] += letter
            
    return treatment_letters

def run_anova_1factor(df, block_col, genotype_col, param):
    df_temp = pd.DataFrame({
        'rep': df[block_col].astype(str),
        'genotype': df[genotype_col].astype(str),
        'response': pd.to_numeric(df[param], errors='coerce')
    }).dropna(subset=['response'])
    
    model = ols("response ~ C(rep) + C(genotype)", data=df_temp).fit()
    anova_table = sm.stats.anova_lm(model, typ=1)
    
    df_err = anova_table.loc["Residual", "df"]
    mse = anova_table.loc["Residual", "mean_sq"]
    p_val = anova_table.loc["C(genotype)", "PR(>F)"]
    
    grand_mean = df_temp['response'].mean()
    cv = (np.sqrt(mse) / grand_mean) * 100
    
    r = df_temp['rep'].nunique()
    t_val = t.ppf(0.975, df_err)
    
    means = df_temp.groupby('genotype')['response'].mean().to_dict()
    sem = np.sqrt(mse / r)
    lsd = t_val * np.sqrt((2 * mse) / r)
    
    cld = get_cld_letters(means, lsd)
    sig_code = get_signif_code_val(p_val)
    p_text = f"{p_val:.4f}{sig_code}" if p_val >= 0.0001 else f"0.0000{sig_code}"
    
    means_str = {}
    for g, val in means.items():
        let = cld[g] if p_val < 0.05 else ""
        means_str[g] = f"{val:.2f}{let}"
        
    return {
        "means": means,
        "means_str": means_str,
        "sem": round(sem, 2),
        "p_val": p_val,
        "p_text": p_text,
        "lsd": round(lsd, 2),
        "cv": round(cv, 2),
        "gm": round(grand_mean, 2)
    }

# --- Excel Sheet Builder (Matches Provided Excel Image Format Exactly) ---
def build_multiyear_excel_output(genotype_col, params, genotypes, results_1, results_2, year1_lbl, year2_lbl):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Result final"
    
    # Prevent default gridlines hiding
    ws.views.sheetView[0].showGridLines = True
    
    font_bold = Font(name="Calibri", size=11, bold=True)
    font_regular = Font(name="Calibri", size=11)
    align_left = Alignment(horizontal="left", vertical="center")
    align_center = Alignment(horizontal="center", vertical="center")
    border_thin_bottom = Border(bottom=Side(style='thin', color='000000'))
    border_medium_bottom = Border(bottom=Side(style='medium', color='000000'))
    
    # Row 3 Column A is "Genotype" (A3) as per provided image
    ws.cell(row=3, column=1, value="Genotype").font = font_bold
    ws.cell(row=3, column=1).alignment = align_left
    
    # Format Headers (Row 1 Parameter Title, Row 2 Year Subheads, Row 3 empty)
    for i, p in enumerate(params):
        start_col = i * 3 + 2 # Col B, E, H...
        
        # Row 1 Parameter name merged across 3 columns
        ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=start_col+2)
        cell_p = ws.cell(row=1, column=start_col, value=p)
        cell_p.font = font_bold
        cell_p.alignment = align_center
        
        # Row 2 seasonal header names
        ws.cell(row=2, column=start_col, value=year1_lbl).font = font_bold
        ws.cell(row=2, column=start_col).alignment = align_center
        ws.cell(row=2, column=start_col+1, value=year2_lbl).font = font_bold
        ws.cell(row=2, column=start_col+1).alignment = align_center
        ws.cell(row=2, column=start_col+2, value="Polled").font = font_bold
        ws.cell(row=2, column=start_col+2).alignment = align_center
        
    # Write Genotype Means (Row 4 onwards)
    for r_idx, g in enumerate(genotypes, start=4):
        ws.cell(row=r_idx, column=1, value=g).font = font_regular
        ws.cell(row=r_idx, column=1).alignment = align_left
        for i, p in enumerate(params):
            start_col = i * 3 + 2
            
            val_1 = results_1[p]["means_str"].get(g, "")
            val_2 = results_2[p]["means_str"].get(g, "")
            pooled_val = (results_1[p]["means"].get(g, 0.0) + results_2[p]["means"].get(g, 0.0)) / 2
            
            cell1 = ws.cell(row=r_idx, column=start_col, value=val_1)
            cell1.font = font_regular
            cell1.alignment = align_center
            
            cell2 = ws.cell(row=r_idx, column=start_col+1, value=val_2)
            cell2.font = font_regular
            cell2.alignment = align_center
            
            # Polled/Pooled has no letters as per the provided screenshot
            cell3 = ws.cell(row=r_idx, column=start_col+2, value=f"{pooled_val:.2f}")
            cell3.font = font_regular
            cell3.alignment = align_center
            
    # Write Statistical Rows
    stats_labels = ["Sem", "p-value", "LSD(0.05)", "CV(%)", "Grand Mean"]
    stats_keys = ["sem", "p_text", "lsd", "cv", "gm"]
    
    start_stats_row = len(genotypes) + 4
    for s_idx, (label, key) in enumerate(zip(stats_labels, stats_keys)):
        curr_row = start_stats_row + s_idx
        ws.cell(row=curr_row, column=1, value=label).font = font_bold
        ws.cell(row=curr_row, column=1).alignment = align_left
        for i, p in enumerate(params):
            start_col = i * 3 + 2
            
            cell1 = ws.cell(row=curr_row, column=start_col, value=results_1[p][key])
            cell1.font = font_regular
            cell1.alignment = align_center
            
            cell2 = ws.cell(row=curr_row, column=start_col+1, value=results_2[p][key])
            cell2.font = font_regular
            cell2.alignment = align_center
            
            # Polled statistical cells are kept completely blank matching your Excel layout image!
            cell3 = ws.cell(row=curr_row, column=start_col+2, value="")
            cell3.font = font_regular
            cell3.alignment = align_center
            
    total_cols = len(params) * 3 + 1
    
    # Apply Standard Borders (Medium Border separating datasets and totals)
    for col in range(1, total_cols + 1):
        ws.cell(row=1, column=col).border = border_thin_bottom
        ws.cell(row=3, column=col).border = border_thin_bottom
        ws.cell(row=start_stats_row - 1, column=col).border = border_thin_bottom
        ws.cell(row=start_stats_row + len(stats_keys) - 1, column=col).border = border_medium_bottom
        
    return wb

# --- DOCX Generation Table Builder (Matches screenshot columns exactly) ---
def add_multiyear_table_to_docx(doc, params, genotypes, results_1, results_2, year1_lbl, year2_lbl):
    num_cols = len(params) * 3 + 1
    table = doc.add_table(rows=3, cols=num_cols)
    set_table_borders(table)
    
    hdr_row0 = table.rows[0].cells
    hdr_row1 = table.rows[1].cells
    hdr_row2 = table.rows[2].cells
    
    hdr_row0[0].text = "Genotype"
    set_cell_margins(hdr_row0[0])
    hdr_row0[0].paragraphs[0].runs[0].font.bold = True
    hdr_row0[0].paragraphs[0].runs[0].font.name = 'Arial'
    hdr_row0[0].paragraphs[0].runs[0].font.size = Pt(10)
    
    # Vertically merge genotype headers
    hdr_row0[0].merge(hdr_row1[0]).merge(hdr_row2[0])
    
    for i, p in enumerate(params):
        start_col = i * 3 + 1
        
        # Horizontally merge parameters header
        hdr_row0[start_col].merge(hdr_row0[start_col+1]).merge(hdr_row0[start_col+2])
        hdr_row0[start_col].text = p
        set_cell_margins(hdr_row0[start_col])
        hdr_row0[start_col].paragraphs[0].alignment = 1
        hdr_row0[start_col].paragraphs[0].runs[0].font.bold = True
        hdr_row0[start_col].paragraphs[0].runs[0].font.name = 'Arial'
        hdr_row0[start_col].paragraphs[0].runs[0].font.size = Pt(10)
        
        # Write subheadings
        hdr_row1[start_col].text = year1_lbl
        set_cell_margins(hdr_row1[start_col])
        hdr_row1[start_col].paragraphs[0].alignment = 1
        hdr_row1[start_col].paragraphs[0].runs[0].font.bold = True
        hdr_row1[start_col].paragraphs[0].runs[0].font.name = 'Arial'
        hdr_row1[start_col].paragraphs[0].runs[0].font.size = Pt(10)
        
        hdr_row1[start_col+1].text = year2_lbl
        set_cell_margins(hdr_row1[start_col+1])
        hdr_row1[start_col+1].paragraphs[0].alignment = 1
        hdr_row1[start_col+1].paragraphs[0].runs[0].font.bold = True
        hdr_row1[start_col+1].paragraphs[0].runs[0].font.name = 'Arial'
        hdr_row1[start_col+1].paragraphs[0].runs[0].font.size = Pt(10)
        
        hdr_row1[start_col+2].text = "Polled"
        set_cell_margins(hdr_row1[start_col+2])
        hdr_row1[start_col+2].paragraphs[0].alignment = 1
        hdr_row1[start_col+2].paragraphs[0].runs[0].font.bold = True
        hdr_row1[start_col+2].paragraphs[0].runs[0].font.name = 'Arial'
        hdr_row1[start_col+2].paragraphs[0].runs[0].font.size = Pt(10)
        
        # Row 2 left blank / empty columns matching image design
        for sub_col in range(3):
            hdr_row2[start_col + sub_col].text = ""
            set_cell_margins(hdr_row2[start_col + sub_col])
            
    set_header_bottom_border(table.rows[2])
    
    # Write treatment datasets
    for g in sorted(genotypes):
        row_cells = table.add_row().cells
        row_cells[0].text = str(g)
        set_cell_margins(row_cells[0])
        row_cells[0].paragraphs[0].runs[0].font.name = 'Arial'
        row_cells[0].paragraphs[0].runs[0].font.size = Pt(10)
        
        for i, p in enumerate(params):
            start_col = i * 3 + 1
            pooled_val = (results_1[p]["means"].get(g, 0.0) + results_2[p]["means"].get(g, 0.0)) / 2
            
            row_cells[start_col].text = str(results_1[p]["means_str"].get(g, ""))
            set_cell_margins(row_cells[start_col])
            row_cells[start_col].paragraphs[0].alignment = 1
            row_cells[start_col].paragraphs[0].runs[0].font.name = 'Arial'
            row_cells[start_col].paragraphs[0].runs[0].font.size = Pt(10)
            
            row_cells[start_col+1].text = str(results_2[p]["means_str"].get(g, ""))
            set_cell_margins(row_cells[start_col+1])
            row_cells[start_col+1].paragraphs[0].alignment = 1
            row_cells[start_col+1].paragraphs[0].runs[0].font.name = 'Arial'
            row_cells[start_col+1].paragraphs[0].runs[0].font.size = Pt(10)
            
            row_cells[start_col+2].text = f"{pooled_val:.2f}"
            set_cell_margins(row_cells[start_col+2])
            row_cells[start_col+2].paragraphs[0].alignment = 1
            row_cells[start_col+2].paragraphs[0].runs[0].font.name = 'Arial'
            row_cells[start_col+2].paragraphs[0].runs[0].font.size = Pt(10)
            
    set_header_bottom_border(table.rows[-1])
    
    stats_labels = ["Sem", "p-value", "LSD(0.05)", "CV(%)", "Grand Mean"]
    stats_keys = ["sem", "p_text", "lsd", "cv", "gm"]
    
    for s_idx, (label, key) in enumerate(zip(stats_labels, stats_keys)):
        row_cells = table.add_row().cells
        row_cells[0].text = label
        set_cell_margins(row_cells[0])
        row_cells[0].paragraphs[0].runs[0].font.bold = True
        row_cells[0].paragraphs[0].runs[0].font.name = 'Arial'
        row_cells[0].paragraphs[0].runs[0].font.size = Pt(10)
        
        for i, p in enumerate(params):
            start_col = i * 3 + 1
            row_cells[start_col].text = str(results_1[p][key])
            set_cell_margins(row_cells[start_col])
            row_cells[start_col].paragraphs[0].alignment = 1
            row_cells[start_col].paragraphs[0].runs[0].font.name = 'Arial'
            row_cells[start_col].paragraphs[0].runs[0].font.size = Pt(10)
            
            row_cells[start_col+1].text = str(results_2[p][key])
            set_cell_margins(row_cells[start_col+1])
            row_cells[start_col+1].paragraphs[0].alignment = 1
            row_cells[start_col+1].paragraphs[0].runs[0].font.name = 'Arial'
            row_cells[start_col+1].paragraphs[0].runs[0].font.size = Pt(10)
            
            row_cells[start_col+2].text = "" # Blank pooled statistical cells
            set_cell_margins(row_cells[start_col+2])
            
        if s_idx == len(stats_keys) - 1:
            set_header_bottom_border(row_cells)

# --- Vertical borders helper functions ---
def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)

def set_table_borders(table):
    tblPr = table._tbl.tblPr
    borders = parse_xml(
        '<w:tblBorders %s>'
        '<w:top w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
        '<w:left w:val="none"/>'
        '<w:right w:val="none"/>'
        '<w:insideH w:val="none"/>'
        '<w:insideV w:val="none"/>'
        '</w:tblBorders>' % nsdecls('w')
    )
    tblPr.append(borders)

def set_header_bottom_border(row):
    for cell in row.cells:
        tcPr = cell._tc.get_or_add_tcPr()
        borders = parse_xml(
            '<w:tcBorders %s>'
            '<w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            '</w:tcBorders>' % nsdecls('w')
        )
        tcPr.append(borders)

# --- Web Interface and Multi-Year controller ---
def show_module():
    st.markdown("### Multi-Year Single-Factor RCBD Analyzer")
    
    file1 = st.file_uploader("Upload Season 1 Raw Excel Dataset (.xlsx)", type=["xlsx"], key="file1_2f_mod")
    file2 = st.file_uploader("Upload Season 2 Raw Excel Dataset (.xlsx)", type=["xlsx"], key="file2_2f_mod")
    
    year1_lbl = st.text_input("Header label for Season 1:", value="2079/80")
    year2_lbl = st.text_input("Header label for Season 2:", value="2080/81")
    
    if file1 is not None and file2 is not None:
        try:
            df1_raw = pd.read_excel(file1)
            df2_raw = pd.read_excel(file2)
            
            st.write(f"#### Season 1 Preview ({year1_lbl}):", df1_raw.head())
            st.write(f"#### Season 2 Preview ({year2_lbl}):", df2_raw.head())
            
            cols = df1_raw.columns.tolist()
            block_col = st.selectbox("Select Block/Replication Column:", cols, index=0, key="block_2f")
            genotype_col = st.selectbox("Select Genotype/Treatment Column:", cols, index=1, key="genotype_2f")
            response_cols = st.multiselect("Select Response Parameters to Analyze:", cols, default=cols[2:], key="response_2f")
            
            if response_cols:
                group_1_name = st.text_input("First Table Title", "Physiological and Crop Growth Parameters", key="2f_t1_title")
                group_1_cols = st.multiselect("Select parameters for Table 1", response_cols, default=response_cols[:len(response_cols)//2], key="2f_t1_cols")
                group_2_name = st.text_input("Second Table Title", "Crop Quality and Yield Parameters", key="2f_t2_title")
                group_2_cols = st.multiselect("Select parameters for Table 2", [c for c in response_cols if c not in group_1_cols], default=[c for c in response_cols if c not in group_1_cols], key="2f_t2_cols")
                
                if st.button("Execute Multi-Year Analytical Engine", key="run_2f_engine"):
                    genotypes_1 = sorted(df1_raw[genotype_col].dropna().unique().tolist())
                    genotypes_2 = sorted(df2_raw[genotype_col].dropna().unique().tolist())
                    
                    common_genotypes = sorted(list(set(genotypes_1).intersection(set(genotypes_2))))
                    if not common_genotypes:
                        st.error("Error: Genotype column entries do not match between both files.")
                        return
                    
                    results_data_1 = {}
                    results_data_2 = {}
                    
                    for param in response_cols:
                        results_data_1[param] = run_anova_1factor(df1_raw, block_col, genotype_col, param)
                        results_data_2[param] = run_anova_1factor(df2_raw, block_col, genotype_col, param)
                        
                    # Build Excel Results matching your screenshot layout exactly
                    excel_bio = io.BytesIO()
                    styled_wb = build_multiyear_excel_output(
                        genotype_col, response_cols, common_genotypes, 
                        results_data_1, results_data_2, year1_lbl, year2_lbl
                    )
                    styled_wb.save(excel_bio)
                    excel_bio.seek(0)
                    
                    st.markdown("#### 📥 Download Formatted Statistical Excel Results")
                    st.download_button(
                        label="Download Excel Results Sheet",
                        data=excel_bio,
                        file_name="RCBD_Result_Final.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="btn_d_excel_2f"
                    )
                    st.write("---")
                    
                    # Generate dynamic report
                    st.markdown("### 📝 Dynamic Analysis Results & Discussions")
                    doc = Document()
                    doc.add_heading("Multi-Year Single-Factor RCBD Trial Report", 0)
                    
                    tables_layouts = [
                        (group_1_name, group_1_cols, 1),
                        (group_2_name, group_2_cols, 2)
                    ]
                    
                    for g_title, g_cols, table_num in tables_layouts:
                        if not g_cols: continue
                        table_label = f"Table {table_num}"
                        
                        st.write(f"### {table_label}: {g_title}")
                        doc.add_heading(f"{table_label}: {g_title}", level=2)
                        
                        grouped = group_parameters(g_cols)
                        
                        for base_name, items in sorted(grouped.items()):
                            st.write(f"#### {base_name}")
                            doc.add_heading(base_name, level=3)
                            
                            if len(items) > 1:
                                p_text = generate_trend_explanation_2y(base_name, items, results_data_1, results_data_2, year1_lbl, year2_lbl, table_label)
                            else:
                                p_text = generate_single_explanation_2y(items[0][0], results_data_1[items[0][0]], results_data_2[items[0][0]], year1_lbl, year2_lbl, table_label)
                                
                            st.write(p_text)
                            
                            p_docx = doc.add_paragraph()
                            parts = re.split(r'(\*\*.*?\*\*)', p_text)
                            for part in parts:
                                if part.startswith('**') and part.endswith('**'):
                                    p_docx.add_run(part[2:-2]).bold = True
                                else:
                                    p_docx.add_run(part)
                                    
                        st.write("##### Corresponding Table Visualization:")
                        add_multiyear_table_to_docx(doc, g_cols, common_genotypes, results_data_1, results_data_2, year1_lbl, year2_lbl)
                        st.write("*(Table data formatted precisely as per standard journal specifications)*")
                        doc.add_page_break()
                        
                    bio_doc = io.BytesIO()
                    doc.save(bio_doc)
                    bio_doc.seek(0)
                    
                    st.write("---")
                    st.markdown("#### 💾 Save Explanations & Tables as Word Document")
                    st.download_button(
                        "Download Word Explanations Report (.docx)", 
                        data=bio_doc, 
                        file_name="MultiYear_Thesis_Report.docx", 
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
                        key="btn_d_2f_doc"
                    )
        except Exception as e:
            st.error(f"Analysis failed. Please check the structure of your Excel dataset: {e}")

if __name__ == '__main__':
    st.set_page_config(layout="wide")
    show_module()
