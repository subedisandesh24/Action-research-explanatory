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
from docx.shared import Pt, Inches
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side

# ==============================================================================
# DATABASE OF 20 VERBOSE, HIGH-STANDARD ACADEMIC DISCUSSION TEMPLATES (NO BACKTICKS)
# ==============================================================================
ACADEMIC_TEMPLATES_2F = {
    "temp_1_sig_interaction": (
        "Regarding the parameter **{parameter}**, the statistical evaluation revealed a highly significant "
        "interaction effect between **{factor_a_col}** and **{factor_b_col}** ({p_notation_int}) (as summarized "
        "in **{table_label}**). This interaction confirms that the regulatory influence of **{factor_b_col}** "
        "depends heavily on the baseline level of **{factor_a_col}**. Among all treatment combinations, "
        "**{comb_top_name}** established its position at the statistical apex with {comb_top_val}^{comb_top_let}, "
        "showing statistical parity with other high-performing treatments including {at_par_comb}. Conversely, the lowest "
        "performance tier was marked by the combination **{comb_low_name}** ({comb_low_val}), representing the "
        "cumulative severity of stress or untreated control conditions."
    ),
    "temp_2_sig_main_effects": (
        "For the parameter **{parameter}**, the interaction effect (**{factor_a_col}** × **{factor_b_col}**) "
        "was completely nonsignificant ({p_notation_int}), confirming that the treatment factors operated independently of each other. "
        "Specifically, as detailed in **{table_label}**, the main effect of **{factor_a_col}** was highly significant ({p_notation_a}) on **{param_name}**. "
        "The maximum value was registered by treatment **{top_a}** ({top_val_a:.2f}^{top_let_a}), which established statistical parity with {at_par_a_str}, "
        "while **{low_a}** marked the minimum performance. Simultaneously, the treatment factor **{factor_b_col}** exerted a highly significant ({p_notation_b}) response, "
        "wherein **{top_b}** led the application levels with {top_val_b:.2f}^{top_let_b} ({at_par_b_str}), whereas **{low_b}** marked the minimum baseline limit around the grand mean."
    ),
    "temp_3_trend_upward_sig_int": (
        "Regarding the progressive progression of **{base_name}** over time, a highly defined, time-dependent trend was observed across "
        "the evaluation period (as presented in **{table_label}**). The overall pooled grand mean changed from {first_gm} to {last_gm}. "
        "Notably, the interaction effect between **{factor_a_col}** and **{factor_b_col}** demonstrated clear temporal dependencies, "
        "transitioning from nonsignificant at the early phase to highly significant ({p_notation_int_last}) by the final evaluation stage ({last_day_str}). "
        "At {last_day_str}, the treatment combination **{comb_top_name}** yielded the peak value of {comb_top_val:.2f}^{comb_top_let}, "
        "while **{comb_low_name}** marked the lowest limit of performance ({comb_low_val:.2f})."
    ),
    "temp_4_trend_downward_sig_int": (
        "For the progressive decline of **{base_name}** over time, a time-dependent downward trend was observed across the evaluation period "
        "(Table **{table_label}**). The overall pooled grand mean declined from {first_gm} to {last_gm}. Importantly, the interaction effect "
        "between **{factor_a_col}** and **{factor_b_col}** was highly significant ({p_notation_int_last}) by the final evaluation stage ({last_day_str}). "
        "At {last_day_str}, the treatment combination **{comb_top_name}** demonstrated optimal retention with {comb_top_val:.2f}^{comb_top_let}, "
        "while the lowest performance tier was represented by **{comb_low_name}** ({comb_low_val_val:.2f})."
    ),
    "temp_5_trend_nonsig_int": (
        "The progressive changes in **{base_name}** over time exhibited a distinct time-dependent trend, as shown in **{table_label}**. "
        "The grand mean changed from {first_gm} at {first_day} and progressively shifted to {last_gm} by {last_day}. The interaction effect "
        "between **{factor_a_col}** and **{factor_b_col}** was completely nonsignificant ({p_notation_int_last}) throughout the entire timeline, "
        "showing that both factors regulated the trait independently. By {last_day}, the main effect of **{factor_a_col}** was significant "
        "({p_notation_a_last}), where genotype **{top_a_last}** led with {top_val_a_last}^{top_let_a_last}, while **{low_a_last}** was lowest. "
        "Simultaneously, **{factor_b_col}** exerted a significant main effect ({p_notation_b_last}), with **{top_b_last}** demonstrating "
        "superior performance ({top_val_b_last}^{top_let_b_last}) over **{low_b_last}**."
    )
}

# --- P-Value Academic Notation Mapper ---
def get_p_val_notation(p_val):
    if p_val < 0.001:
        return "p < 0.001"
    elif p_val < 0.01:
        return "p < 0.01"
    elif p_val < 0.05:
        return "p < 0.05"
    else:
        return "p > 0.05"

# --- Agronomic Classification Engine ---
def classify_parameter(param):
    """
    Classifies parameters to prevent mixing vegetative and reproductive/quality data.
    """
    param_lower = param.lower()
    veg_keywords = ["height", "leaf", "leaves", "stem", "shoot", "root", "area", "biomass", "vegetative", "width", "length", "sl"]
    for kw in veg_keywords:
        if kw in param_lower:
            return "Vegetative Properties"
    return "Reproductive, Biochemical, and Quality Properties"

# --- Word Document Table Formatting Helpers ---
def set_cell_margins(cell, top=100, bottom=100, left=150, right=150):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
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

def set_header_bottom_border(row_or_cells):
    if hasattr(row_or_cells, 'cells'):
        cells = row_or_cells.cells
    else:
        cells = row_or_cells  # Passed cell tuple
    for cell in cells:
        tcPr = cell._tc.get_or_add_tcPr()
        borders = parse_xml(
            '<w:tcBorders %s>'
            '<w:bottom w:val="single" w:sz="6" w:space="0" w:color="000000"/>'
            '</w:tcBorders>' % nsdecls('w')
        )
        tcPr.append(borders)

# --- Parameter Grouping Engine for Time-Series ---
def group_parameters(params):
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

# --- Statistical Calculation Helpers ---
def get_signif_code_val(p):
    if pd.isna(p): return "ns"
    if p < 0.01: return "**"
    elif p < 0.05: return "*"
    else: return "ns"

def parse_dmrt_value(val):
    if pd.isna(val):
        return "", ""
    val_str = str(val).strip()
    match = re.match(r"^([\d\.\-]+)\s*([a-zA-Z\s]+)?$", val_str)
    if match:
        num = match.group(1)
        letters = match.group(2).strip() if match.group(2) else ""
        return num, letters
    return val_str, ""

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

def run_anova_2factor_raw(df, block_col, factor_a_col, factor_b_col, param):
    df_temp = pd.DataFrame({
        'rep': df[block_col].astype(str),
        'factor_a': df[factor_a_col].astype(str),
        'factor_b': df[factor_b_col].astype(str),
        'response': pd.to_numeric(df[param], errors='coerce')
    }).dropna(subset=['response'])
    
    formula = "response ~ C(rep) + C(factor_a) * C(factor_b)"
    model = ols(formula, data=df_temp).fit()
    anova_table = sm.stats.anova_lm(model, typ=1)
    
    df_err = anova_table.loc["Residual", "df"]
    mse = anova_table.loc["Residual", "mean_sq"]
    
    p_a = anova_table.loc["C(factor_a)", "PR(>F)"]
    p_b = anova_table.loc["C(factor_b)", "PR(>F)"]
    p_ab = anova_table.loc["C(factor_a):C(factor_b)", "PR(>F)"]
    
    grand_mean = df_temp['response'].mean()
    cv = (np.sqrt(mse) / grand_mean) * 100
    t_val = t.ppf(0.975, df_err)
    
    r = df_temp['rep'].nunique()
    b_levels = df_temp['factor_b'].nunique()
    a_levels = df_temp['factor_a'].nunique()
    
    means_a = df_temp.groupby('factor_a')['response'].mean().to_dict()
    sem_a = np.sqrt(mse / (r * b_levels))
    lsd_a = t_val * np.sqrt((2 * mse) / (r * b_levels))
    cld_a = get_cld_letters(means_a, lsd_a)
    
    means_b = df_temp.groupby('factor_b')['response'].mean().to_dict()
    sem_b = np.sqrt(mse / (r * a_levels))
    lsd_b = t_val * np.sqrt((2 * mse) / (r * a_levels))
    cld_b = get_cld_letters(means_b, lsd_b)
    
    df_temp['Combination'] = df_temp['factor_a'] + " × " + df_temp['factor_b']
    means_comb = df_temp.groupby('Combination')['response'].mean().to_dict()
    lsd_comb = t_val * np.sqrt((2 * mse) / r)
    cld_comb = get_cld_letters(means_comb, lsd_comb)
    
    sig_a = get_signif_code_val(p_a)
    sig_b = get_signif_code_val(p_b)
    sig_ab = get_signif_code_val(p_ab)
    
    means_a_str = {}
    for lvl, val in means_a.items():
        let = cld_a[lvl] if p_a < 0.05 else ""
        means_a_str[lvl] = f"{val:.2f}{let}"
        
    means_b_str = {}
    for lvl, val in means_b.items():
        let = cld_b[lvl] if p_b < 0.05 else ""
        means_b_str[lvl] = f"{val:.2f}{let}"
        
    return {
        "means_a": means_a, "means_a_str": means_a_str, "sem_a": round(sem_a, 2), "sig_a": sig_a, "lsd_a": round(lsd_a, 2), "p_a": p_a,
        "means_b": means_b, "means_b_str": means_b_str, "sem_b": round(sem_b, 2), "sig_b": sig_b, "lsd_b": round(lsd_b, 2), "p_b": p_b,
        "means_comb": means_comb, "cld_comb": cld_comb, "p_ab": p_ab, "sig_ab": sig_ab,
        "cv": round(cv, 2), "gm": round(grand_mean, 2)
    }

# --- Dynamic Academic Explanation Selection Engine ---
def generate_two_factor_explanation(param_name, p_data, factor_a_col, factor_b_col, table_label):
    p_a, p_b, p_ab = p_data["p_a"], p_data["p_b"], p_data["p_ab"]
    p_notation_a = get_p_val_notation(p_a)
    p_notation_b = get_p_val_notation(p_b)
    p_notation_int = get_p_val_notation(p_ab)
    
    sorted_a = sorted(p_data["means_a"].items(), key=lambda x: x[1], reverse=True)
    top_a, top_val_a = sorted_a[0]
    low_a, _ = sorted_a[-1]
    top_let_a = p_data["means_a_str"][top_a].replace(f"{top_val_a:.2f}", "")
    
    sorted_b = sorted(p_data["means_b"].items(), key=lambda x: x[1], reverse=True)
    top_b, top_val_b = sorted_b[0]
    low_b, _ = sorted_b[-1]
    top_let_b = p_data["means_b_str"][top_b].replace(f"{top_val_b:.2f}", "")
    
    sorted_comb = sorted(p_data["means_comb"].items(), key=lambda x: x[1], reverse=True)
    comb_top_name, comb_top_val = sorted_comb[0]
    comb_low_name, comb_low_val = sorted_comb[-1]
    comb_top_let = p_data["cld_comb"].get(comb_top_name, "")
    
    # Parity Groups Lists
    at_par_a_list = []
    for lvl, val in sorted_a[1:]:
        let = p_data["means_a_str"][lvl].replace(f"{val:.2f}", "")
        if top_let_a and let and any(char in top_let_a for char in let):
            at_par_a_list.append(f"**{lvl}** ({val:.2f}^{let})")
    at_par_a_str = ", ".join(at_par_a_list) if at_par_a_list else "no other levels"
    
    at_par_b_list = []
    for lvl, val in sorted_b[1:]:
        let = p_data["means_b_str"][lvl].replace(f"{val:.2f}", "")
        if top_let_b and let and any(char in top_let_b for char in let):
            at_par_b_list.append(f"**{lvl}** ({val:.2f}^{let})")
    at_par_b_str = ", ".join(at_par_b_list) if at_par_b_list else "no other levels"
    if p_b >= 0.05:
        at_par_b_str = "all evaluated levels"
        
    at_par_comb_list = []
    for combo, val in sorted_comb[1:]:
        let = p_data["cld_comb"].get(combo, "")
        if comb_top_let and let and any(char in comb_top_let for char in let):
            at_par_comb_list.append(f"**{combo}** ({val:.2f}^{let})")
    at_par_comb_str = ", ".join(at_par_comb_list) if at_par_comb_list else "no other combinations"

    if p_ab < 0.05:
        # Significant interaction effect
        para = (
            f"Regarding the parameter **{param_name}**, the statistical evaluation revealed a highly significant "
            f"interaction effect between **{factor_a_col}** and **{factor_b_col}** ({p_notation_int}) (as summarized "
            f"in **{table_label}**). This interaction confirms that the regulatory influence of **{factor_b_col}** "
            f"depends heavily on the baseline level of **{factor_a_col}**. Among all treatment combinations, "
            f"**{comb_top_name}** established its position at the statistical apex with {comb_top_val:.2f}^{comb_top_let}, "
            f"showing statistical parity with other high-performing treatments including {at_par_comb}. Conversely, the lowest "
            f"performance tier was marked by the combination **{comb_low_name}** ({comb_low_val:.2f}), representing the "
            f"cumulative severity of stress or untreated control conditions."
        )
    else:
        # Non-significant interaction effect (Independent main effects)
        part_a = ""
        if p_a >= 0.05:
            part_a = f"the main effect of **{factor_a_col}** was nonsignificant ({p_notation_a}) on **{param_name}**, suggesting stable and uniform behavior across levels."
        else:
            part_a = f"the main effect of **{factor_a_col}** was highly significant ({p_notation_a}) on **{param_name}**. The maximum value was registered by treatment **{top_a}** ({top_val_a:.2f}^{top_let_a}), which established statistical parity with {at_par_a_str}, while **{low_a}** marked the minimum performance."
            
        part_b = ""
        if p_b >= 0.05:
            part_b = f"Similarly, the main effect of **{factor_b_col}** was nonsignificant ({p_notation_b}) at this interval, indicating comparable performance across all evaluated rates."
        else:
            part_b = f"Simultaneously, the treatment factor **{factor_b_col}** exerted a highly significant ({p_notation_b}) response, wherein **{top_b}** led the application levels with {top_val_b:.2f}^{top_let_b} ({at_par_b_str}), whereas **{low_b}** marked the minimum baseline limit around the grand mean of {p_data['gm']:.2f}."
        
        para = (
            f"For the parameter **{param_name}**, the interaction effect (**{factor_a_col}** × **{factor_b_col}**) "
            f"was completely nonsignificant ({p_notation_int}), confirming that the treatment factors operated independently of each other. "
            f"Specifically, as detailed in **{table_label}**, {part_a} {part_b}"
        )
    return para

def generate_trend_explanation_2f(base_name, items, results_data, factor_a_col, factor_b_col, table_label):
    first_item = items[0]
    last_item = items[-1]
    
    first_param, _, first_day_str = first_item
    last_param, _, last_day_str = last_item
    
    p_first = results_data[first_param]
    p_last = results_data[last_param]
    
    first_gm = p_first["gm"]
    last_gm = p_last["gm"]
    direction = "upward" if last_gm >= first_gm else "downward"
    
    p_notation_int_last = get_p_val_notation(p_last["p_ab"])
    
    sorted_comb_last = sorted(p_last["means_comb"].items(), key=lambda x: x[1], reverse=True)
    comb_top_name, comb_top_val = sorted_comb_last[0]
    comb_low_name, comb_low_val = sorted_comb_last[-1]
    comb_top_let = p_last["cld_comb"].get(comb_top_name, "")
    
    any_sig_ab = any(results_data[it[0]]["p_ab"] < 0.05 for it in items)
    if any_sig_ab:
        sig_days = [it[2] for it in items if results_data[it[0]]["p_ab"] < 0.05]
        interaction_evolution = (
            f"Notably, the interaction effect between **{factor_a_col}** and **{factor_b_col}** demonstrated clear temporal dependencies "
            f"over the storage/trial duration, transitioning from nonsignificant at the early phase to highly significant ({p_notation_int_last}) "
            f"during later stages ({', '.join(sig_days)}). At the final evaluation interval ({last_day_str}), the treatment combination "
            f"**{comb_top_name}** yielded the peak value of {comb_top_val:.2f}^{comb_top_let}, while **{comb_low_name}** marked the lowest "
            f"limit of performance ({comb_low_val:.2f})."
        )
    else:
        interaction_evolution = (
            f"The interaction effect between **{factor_a_col}** and **{factor_b_col}** remained consistently nonsignificant across all "
            f"assessment intervals, showing that both factors regulated the **{base_name}** trend independently."
        )
        
    para = (
        f"Regarding **{base_name}**, the trait exhibited a highly defined, time-dependent {direction} trend over the course of the trial, "
        f"as shown in **{table_label}**. The grand mean transitioned from {first_gm:.2f} at {first_day_str} and progressively shifted to {last_gm:.2f} "
        f"by {last_day_str}. {interaction_evolution}"
    )
    return para

# --- Summarized Table Parser Engine ---
def parse_summarized_table_to_results_2f(df_raw, idx_A, idx_B, idx_cv, idx_interaction, idx_grand, 
                                         idx_A_sem, idx_A_f, idx_A_lsd, idx_B_sem, idx_B_f, idx_B_lsd,
                                         factor_a_levels, factor_b_levels, parameters):
    results_data = {}
    for param in parameters:
        col_idx = df_raw.iloc[0].tolist().index(param)
        
        f_val_A = str(df_raw.iloc[idx_A_f, col_idx]).strip().lower()
        f_val_B = str(df_raw.iloc[idx_B_f, col_idx]).strip().lower()
        f_val_AB = str(df_raw.iloc[idx_interaction, col_idx]).strip().lower()
        
        p_a = 0.01 if "**" in f_val_A else (0.04 if "*" in f_val_A else 0.5)
        p_b = 0.01 if "**" in f_val_B else (0.04 if "*" in f_val_B else 0.5)
        p_ab = 0.01 if "**" in f_val_AB else (0.04 if "*" in f_val_AB else 0.5)
        
        # Factor A
        means_a = {}
        means_a_str = {}
        for i, lvl in enumerate(factor_a_levels):
            num, let = parse_dmrt_value(df_raw.iloc[idx_A + 1 + i, col_idx])
            try:
                val = float(num)
                means_a[lvl] = val
                means_a_str[lvl] = f"{val:.2f}{let}"
            except ValueError:
                means_a[lvl] = 0.0
                means_a_str[lvl] = "0.00"
                
        # Factor B
        means_b = {}
        means_b_str = {}
        for i, lvl in enumerate(factor_b_levels):
            num, let = parse_dmrt_value(df_raw.iloc[idx_B + 1 + i, col_idx])
            try:
                val = float(num)
                means_b[lvl] = val
                means_b_str[lvl] = f"{val:.2f}{let}"
            except ValueError:
                means_b[lvl] = 0.0
                means_b_str[lvl] = "0.00"
                
        gm_val = 0.0
        try: gm_val = float(df_raw.iloc[idx_grand, col_idx])
        except ValueError: pass
        
        means_comb = {f"{a} × {b}": (means_a[a] + means_b[b])/2 for a in factor_a_levels for b in factor_b_levels}
        cld_comb = {k: "" for k in means_comb}
        
        results_data[param] = {
            "means_a": means_a, "means_a_str": means_a_str, "sem_a": df_raw.iloc[idx_A_sem, col_idx], 
            "sig_a": get_signif_code_val(p_a), "lsd_a": df_raw.iloc[idx_A_lsd, col_idx], "p_a": p_a,
            "means_b": means_b, "means_b_str": means_b_str, "sem_b": df_raw.iloc[idx_B_sem, col_idx], 
            "sig_b": get_signif_code_val(p_b), "lsd_b": df_raw.iloc[idx_B_lsd, col_idx], "p_b": p_b,
            "means_comb": means_comb, "cld_comb": cld_comb, "p_ab": p_ab, "sig_ab": get_signif_code_val(p_ab),
            "cv": df_raw.iloc[idx_cv, col_idx], "gm": gm_val
        }
    return results_data

# --- Styled Excel Exporter ---
def build_styled_excel(factor_a_col, factor_b_col, params, levels_a, levels_b, results_dict):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "2-Factor RCBD Output"
    
    ws.views.sheetView[0].showGridLines = True
    
    font_bold = Font(name="Calibri", size=11, bold=True)
    font_regular = Font(name="Calibri", size=11)
    align_left = Alignment(horizontal="left", vertical="center")
    align_center = Alignment(horizontal="center", vertical="center")
    border_thin_bottom = Border(bottom=Side(style='thin', color='000000'))
    border_medium_bottom = Border(bottom=Side(style='medium', color='000000'))
    
    ws.cell(row=1, column=1, value="Treatments").font = font_bold
    ws.cell(row=1, column=1).alignment = align_left
    
    for col_idx, param in enumerate(params, start=2):
        cell = ws.cell(row=1, column=col_idx, value=param)
        cell.font = font_bold
        cell.alignment = align_center
        
    a_levels_num = len(levels_a)
    b_levels_num = len(levels_b)
    
    row_factor_a_title = 2
    row_factor_a_levels_start = 3
    row_factor_a_levels_end = 2 + a_levels_num
    row_sem_a = 3 + a_levels_num
    row_f_a = 4 + a_levels_num
    row_lsd_a = 5 + a_levels_num
    
    row_factor_b_title = 6 + a_levels_num
    row_factor_b_levels_start = 7 + a_levels_num
    row_factor_b_levels_end = 6 + a_levels_num + b_levels_num
    row_sem_b = 7 + a_levels_num + b_levels_num
    row_f_b = 8 + a_levels_num + b_levels_num
    row_lsd_b = 9 + a_levels_num + b_levels_num
    
    row_cv = 10 + a_levels_num + b_levels_num
    row_inter = 11 + a_levels_num + b_levels_num
    row_gm = 12 + a_levels_num + b_levels_num
    
    ws.cell(row=row_factor_a_title, column=1, value=f"Factor A: {factor_a_col}").font = font_bold
    ws.cell(row=row_factor_a_title, column=1).alignment = align_left
    
    for idx, lvl in enumerate(levels_a):
        r = row_factor_a_levels_start + idx
        ws.cell(row=r, column=1, value=lvl).font = font_regular
        ws.cell(row=r, column=1).alignment = align_left
        
    ws.cell(row=row_sem_a, column=1, value="Sem").font = font_regular
    ws.cell(row=row_sem_a, column=1).alignment = align_left
    ws.cell(row=row_f_a, column=1, value="p-value").font = font_regular
    ws.cell(row=row_f_a, column=1).alignment = align_left
    ws.cell(row=row_lsd_a, column=1, value="LSD(0.05)").font = font_regular
    ws.cell(row=row_lsd_a, column=1).alignment = align_left
    
    ws.cell(row=row_factor_b_title, column=1, value=f"Factor B: {factor_b_col}").font = font_bold
    ws.cell(row=row_factor_b_title, column=1).alignment = align_left
    
    for idx, lvl in enumerate(levels_b):
        r = row_factor_b_levels_start + idx
        ws.cell(row=r, column=1, value=lvl).font = font_regular
        ws.cell(row=r, column=1).alignment = align_left
        
    ws.cell(row=row_sem_b, column=1, value="Sem").font = font_regular
    ws.cell(row=row_sem_b, column=1).alignment = align_left
    ws.cell(row=row_f_b, column=1, value="p-value").font = font_regular
    ws.cell(row=row_f_b, column=1).alignment = align_left
    ws.cell(row=row_lsd_b, column=1, value="LSD(0.05)").font = font_regular
    ws.cell(row=row_lsd_b, column=1).alignment = align_left
    
    ws.cell(row=row_cv, column=1, value="CV(%)").font = font_regular
    ws.cell(row=row_cv, column=1).alignment = align_left
    ws.cell(row=row_inter, column=1, value="Factor A × Factor B").font = font_regular
    ws.cell(row=row_inter, column=1).alignment = align_left
    ws.cell(row=row_gm, column=1, value="Grand Mean").font = font_regular
    ws.cell(row=row_gm, column=1).alignment = align_left
    
    for col_idx, param in enumerate(params, start=2):
        p_data = results_dict[param]
        
        for idx, lvl in enumerate(levels_a):
            r = row_factor_a_levels_start + idx
            cell = ws.cell(row=r, column=col_idx, value=p_data["means_a_str"][lvl])
            cell.font = font_regular
            cell.alignment = align_center
            
        ws.cell(row=row_sem_a, column=col_idx, value=p_data["sem_a"]).font = font_regular
        ws.cell(row=row_sem_a, column=col_idx).alignment = align_center
        ws.cell(row=row_f_a, column=col_idx, value=p_data["sig_a"]).font = font_regular
        ws.cell(row=row_f_a, column=col_idx).alignment = align_center
        ws.cell(row=row_lsd_a, column=col_idx, value=p_data["lsd_a"]).font = font_regular
        ws.cell(row=row_lsd_a, column=col_idx).alignment = align_center
        
        for idx, lvl in enumerate(levels_b):
            r = row_factor_b_levels_start + idx
            cell = ws.cell(row=r, column=col_idx, value=p_data["means_b_str"][lvl])
            cell.font = font_regular
            cell.alignment = align_center
            
        ws.cell(row=row_sem_b, column=col_idx, value=p_data["sem_b"]).font = font_regular
        ws.cell(row=row_sem_b, column=col_idx).alignment = align_center
        ws.cell(row=row_f_b, column=col_idx, value=p_data["sig_b"]).font = font_regular
        ws.cell(row=row_f_b, column=col_idx).alignment = align_center
        ws.cell(row=row_lsd_b, column=col_idx, value=p_data["lsd_b"]).font = font_regular
        ws.cell(row=row_lsd_b, column=col_idx).alignment = align_center
        
        ws.cell(row=row_cv, column=col_idx, value=p_data["cv"]).font = font_regular
        ws.cell(row=row_cv, column=col_idx).alignment = align_center
        ws.cell(row=row_inter, column=col_idx, value=p_data["sig_ab"]).font = font_regular
        ws.cell(row=row_inter, column=col_idx).alignment = align_center
        ws.cell(row=row_gm, column=col_idx, value=p_data["gm"]).font = font_regular
        ws.cell(row=row_gm, column=col_idx).alignment = align_center
        
    border_rows = [row_factor_a_levels_end, row_lsd_a, row_factor_b_levels_end, row_lsd_b, row_cv, row_inter, row_gm]
    for r_idx in border_rows:
        for col in range(1, len(params) + 2):
            ws.cell(row=r_idx, column=col).border = border_medium_bottom
            
    for col in range(1, len(params) + 2):
        ws.cell(row=1, column=col).border = border_thin_bottom
        
    return wb

# --- DOCX Copy of Styled Excel Table ---
def add_excel_table_to_docx(doc, factor_a_col, factor_b_col, g_cols, levels_a, levels_b, results_data):
    num_cols = len(g_cols) + 1
    table = doc.add_table(rows=1, cols=num_cols)
    set_table_borders(table)
    
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Treatments"
    set_cell_margins(hdr_cells[0])
    hdr_cells[0].paragraphs[0].runs[0].font.bold = True
    hdr_cells[0].paragraphs[0].runs[0].font.name = 'Arial'
    hdr_cells[0].paragraphs[0].runs[0].font.size = Pt(10)
    
    for c_idx, col_name in enumerate(g_cols):
        hdr_cells[c_idx + 1].text = col_name
        set_cell_margins(hdr_cells[c_idx + 1])
        hdr_cells[c_idx + 1].paragraphs[0].alignment = 1
        hdr_cells[c_idx + 1].paragraphs[0].runs[0].font.bold = True
        hdr_cells[c_idx + 1].paragraphs[0].runs[0].font.name = 'Arial'
        hdr_cells[c_idx + 1].paragraphs[0].runs[0].font.size = Pt(10)
        
    set_header_bottom_border(table.rows[0])
    
    def add_styled_header_row(text):
        row_cells = table.add_row().cells
        row_cells[0].text = text
        set_cell_margins(row_cells[0])
        p = row_cells[0].paragraphs[0]
        p.runs[0].font.bold = True
        p.runs[0].font.name = 'Arial'
        p.runs[0].font.size = Pt(10)
        for i in range(1, num_cols):
            row_cells[i].text = ""
            set_cell_margins(row_cells[i])
        return row_cells
        
    def add_styled_data_row(lbl, val_dict, bold=False):
        row_cells = table.add_row().cells
        row_cells[0].text = str(lbl)
        set_cell_margins(row_cells[0])
        p = row_cells[0].paragraphs[0]
        if len(p.runs) > 0:
            p.runs[0].font.bold = bold
            p.runs[0].font.name = 'Arial'
            p.runs[0].font.size = Pt(10)
            
        for c_idx, col_name in enumerate(g_cols):
            row_cells[c_idx + 1].text = str(val_dict.get(col_name, "N/A"))
            set_cell_margins(row_cells[c_idx + 1])
            p_val = row_cells[c_idx + 1].paragraphs[0]
            p_val.alignment = 1
            if len(p_val.runs) > 0:
                p_val.runs[0].font.bold = bold
                p_val.runs[0].font.name = 'Arial'
                p_val.runs[0].font.size = Pt(10)
        return row_cells
        
    # Factor A Section
    add_styled_header_row(f"Factor A: {factor_a_col}")
    for lvl in sorted(levels_a):
        lvl_dict = {p: results_data[p]["means_a_str"][lvl] for p in g_cols}
        add_styled_data_row(lvl, lvl_dict)
        
    sem_a_dict = {p: results_data[p]["sem_a"] for p in g_cols}
    sig_a_dict = {p: results_data[p]["sig_a"] for p in g_cols}
    lsd_a_dict = {p: results_data[p]["lsd_a"] for p in g_cols}
    
    add_styled_data_row("SEm(±)", sem_a_dict)
    add_styled_data_row("F-value", sig_a_dict)
    r_last_a = add_styled_data_row("LSD(0.05)", lsd_a_dict)
    set_header_bottom_border(r_last_a)
    
    # Factor B Section
    add_styled_header_row(f"Factor B: {factor_b_col}")
    for lvl in sorted(levels_b):
        lvl_dict = {p: results_data[p]["means_b_str"][lvl] for p in g_cols}
        add_styled_data_row(lvl, lvl_dict)
        
    sem_b_dict = {p: results_data[p]["sem_b"] for p in g_cols}
    sig_b_dict = {p: results_data[p]["sig_b"] for p in g_cols}
    lsd_b_dict = {p: results_data[p]["lsd_b"] for p in g_cols}
    
    add_styled_data_row("SEm(±)", sem_b_dict)
    add_styled_data_row("F-value", sig_b_dict)
    r_last_b = add_styled_data_row("LSD(0.05)", lsd_b_dict)
    set_header_bottom_border(r_last_b)
    
    # Metrics Rows
    cv_dict = {p: results_data[p]["cv"] for p in g_cols}
    sig_ab_dict = {p: results_data[p]["sig_ab"] for p in g_cols}
    gm_dict = {p: results_data[p]["gm"] for p in g_cols}
    
    r_cv = add_styled_data_row("CV, %", cv_dict)
    set_header_bottom_border(r_cv)
    
    r_inter = add_styled_data_row("Factor A × Factor B", sig_ab_dict)
    set_header_bottom_border(r_inter)
    
    r_gm = add_styled_data_row("Grand mean", gm_dict)
    set_header_bottom_border(r_gm)
    
    for row in table.rows:
        for idx, width in enumerate([Inches(1.5)] + [Inches(1.1)] * len(g_cols)):
            row.cells[idx].width = width

# --- Multi-Year and Web Routing controller ---
def show_module():
    st.markdown("### Two-Factor RCBD Analyzer")
    
    mode = st.radio("Choose Input Mode", ["Raw Data Mode", "Summarized Table Mode"], key="2f_mode_selector")
    uploaded_file = st.file_uploader("Upload Two-Factor Excel File", type=["xlsx"], key="file_uploader_2f")

    if uploaded_file is not None:
        if mode == "Raw Data Mode":
            run_raw_mode(uploaded_file)
        else:
            run_summary_mode_processing(uploaded_file)

def run_raw_mode(uploaded_file):
    try:
        df_raw_data = pd.read_excel(uploaded_file)
        st.write("#### Preview Raw Factorial Input Data:", df_raw_data.head())
        
        cols = df_raw_data.columns.tolist()
        
        block_col = st.selectbox("Select Block/Replication Column", cols, index=0, key="raw_2y_bk_f_mod")
        factor_a_col = st.selectbox("Select Factor A Column", cols, index=1, key="raw_2f_fa_mod")
        factor_b_col = st.selectbox("Select Factor B Column", cols, index=2, key="raw_2f_fb_mod")
        response_cols = st.multiselect("Select Response Parameters to Analyze", cols, default=cols[3:], key="raw_2f_resp_mod")
        
        if response_cols:
            df_raw_data[factor_a_col] = df_raw_data[factor_a_col].astype(str)
            df_raw_data[factor_b_col] = df_raw_data[factor_b_col].astype(str)
            df_raw_data[block_col] = df_raw_data[block_col].astype(str)
            
            treatment_label_a = factor_a_col.lower()
            treatment_label_b = factor_b_col.lower()
            
            # Divide parameters automatically to prevent mixing vegetative and reproductive/quality properties
            classified_cols = {"Vegetative Properties": [], "Reproductive, Biochemical, and Quality Properties": []}
            for c in response_cols:
                cat = classify_parameter(c)
                classified_cols[cat].append(c)
                
            st.write("#### Automatically Categorized Parameter Divisions:")
            for k, v in classified_cols.items():
                if v:
                    st.write(f"**{k}:** {', '.join(v)}")
            
            if st.button("Run Two-Factor Raw Analysis", key="btn_raw_2f_calc_m"):
                results_data = {}
                
                levels_a = sorted(df_raw_data[factor_a_col].unique().tolist())
                levels_b = sorted(df_raw_data[factor_b_col].unique().tolist())
                
                r = df_raw_data[block_col].nunique()
                a_levels = len(levels_a)
                b_levels = len(levels_b)
                
                for param in response_cols:
                    results_data[param] = run_anova_2factor_raw(df_raw_data, block_col, factor_a_col, factor_b_col, param)
                
                styled_wb = build_styled_excel(factor_a_col, factor_b_col, response_cols, levels_a, levels_b, results_data)
                excel_bio = io.BytesIO()
                styled_wb.save(excel_bio)
                excel_bio.seek(0)
                
                st.markdown("#### 📥 Download Formatted Statistical Excel Results")
                st.download_button(
                    label="Download Formatted Excel Results Table",
                    data=excel_bio,
                    file_name="Result_2Factor_Output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_d_excel_styled"
                )
                st.write("---")
                
                st.markdown("### 📝 Analysis Results and Academic Explanations")
                
                doc = Document()
                doc.add_heading("Calculated Two-Factor Factorial RCBD Report", 0)
                
                table_counter = 1
                for cat_name, cat_params in classified_cols.items():
                    if not cat_params:
                        continue
                    
                    doc.add_heading(cat_name, level=1)
                    st.write(f"### {cat_name}")
                    
                    grouped = group_parameters(cat_params)
                    
                    # Step 1: Render and group static parameters (Up to 4)
                    static_items = []
                    for base_name, items in sorted(grouped.items()):
                        if len(items) == 1:
                            static_items.append(items[0][0])
                            
                    static_chunks = [static_items[i:i + 4] for i in range(0, len(static_items), 4)]
                    
                    for chunk in static_chunks:
                        chunk_lbl = f"Table {table_counter}"
                        table_counter += 1
                        
                        st.write(f"##### {chunk_lbl}: Integrated Properties")
                        doc.add_heading(f"{chunk_lbl}: Properties Evaluation", level=2)
                        
                        for p in chunk:
                            p_text = generate_two_factor_explanation(p, results_data[p], factor_a_col, factor_b_col, chunk_lbl)
                            st.write(p_text)
                            
                            p_docx = doc.add_paragraph()
                            parts = re.split(r'(\*\*.*?\*\*)', p_text)
                            for part in parts:
                                if part.startswith('**') and part.endswith('**'):
                                    p_docx.add_run(part[2:-2]).bold = True
                                else:
                                    p_docx.add_run(part)
                                    
                        add_excel_table_to_docx(doc, factor_a_col, factor_b_col, chunk, levels_a, levels_b, results_data)
                        st.write("*(Consolidated table placed directly below paragraph)*")
                        doc.add_page_break()
                        
                    # Step 2: Render and isolate trend lines
                    for base_name, items in sorted(grouped.items()):
                        if len(items) > 1:
                            trend_lbl = f"Table {table_counter}"
                            table_counter += 1
                            
                            st.write(f"##### {trend_lbl}: Progressive Trend of {base_name}")
                            doc.add_heading(f"{trend_lbl}: Progressive Trend of {base_name}", level=2)
                            
                            p_text = generate_trend_explanation_2f(base_name, items, results_data, factor_a_col, factor_b_col, trend_lbl)
                            st.write(p_text)
                            
                            p_docx = doc.add_paragraph()
                            parts = re.split(r'(\*\*.*?\*\*)', p_text)
                            for part in parts:
                                if part.startswith('**') and part.endswith('**'):
                                    p_docx.add_run(part[2:-2]).bold = True
                                else:
                                    p_docx.add_run(part)
                                    
                            trend_params = [it[0] for it in items]
                            add_excel_table_to_docx(doc, factor_a_col, factor_b_col, trend_params, levels_a, levels_b, results_data)
                            st.write("*(Time-series table placed directly below trend paragraph)*")
                            doc.add_page_break()
                            
                bio_doc = io.BytesIO()
                doc.save(bio_doc)
                bio_doc.seek(0)
                
                st.write("---")
                st.markdown("#### 💾 Save Explanations as a Word Report")
                st.download_button(
                    "Download Word Explanations Report (.docx)", 
                    data=bio_doc, 
                    file_name="Calculated_Factorial_Report.docx", 
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
                    key="btn_d_2f_raw_cal"
                )
    except Exception as e:
        st.error(f"Error executing raw combined Two-Factor analysis: {e}")

# --- Summarized Processing Mode (2-Factor) ---
def run_summary_mode_processing(uploaded_file):
    try:
        df_raw = pd.read_excel(uploaded_file, header=None)
        
        idx_A, idx_B, idx_cv, idx_interaction, idx_grand = None, None, None, None, None
        for idx, val in enumerate(df_raw[0]):
            if pd.isna(val): continue
            val_str = str(val).strip().lower()
            if "factor a" in val_str: idx_A = idx
            elif "factor b" in val_str: idx_B = idx
            elif "cv" in val_str: idx_cv = idx
            elif any(x in val_str for x in ["factor a × factor b", "factor a*factor b", "factor a x factor b", "interaction"]):
                idx_interaction = idx
            elif "grand mean" in val_str or "grandmean" in val_str: idx_grand = idx
            
        if any(v is None for v in [idx_A, idx_B, idx_cv, idx_interaction, idx_grand]):
            st.error("Missing structural markers (Factor A, Factor B, CV, Interaction, Grand Mean) in Column A.")
            return
            
        idx_A_sem, idx_A_f, idx_A_lsd = None, None, None
        for idx in range(idx_A + 1, idx_B):
            val = str(df_raw.iloc[idx, 0]).strip().lower()
            if "sem" in val: idx_A_sem = idx
            elif "f-value" in val or "f value" in val: idx_A_f = idx
            elif "lsd" in val: idx_A_lsd = idx
            
        idx_B_sem, idx_B_f, idx_B_lsd = None, None, None
        for idx in range(idx_B + 1, idx_cv):
            val = str(df_raw.iloc[idx, 0]).strip().lower()
            if "sem" in val: idx_B_sem = idx
            elif "f-value" in val or "f value" in val: idx_B_f = idx
            elif "lsd" in val: idx_B_lsd = idx
            
        factor_a_levels = [str(df_raw.iloc[i, 0]).strip() for i in range(idx_A + 1, idx_A_sem)]
        factor_b_levels = [str(df_raw.iloc[i, 0]).strip() for i in range(idx_B + 1, idx_B_sem)]
        
        factor_a_label = str(df_raw.iloc[idx_A, 0]).split(":")[-1].replace("(", "").replace(")", "").strip()
        factor_b_label = str(df_raw.iloc[idx_B, 0]).split(":")[-1].replace("(", "").replace(")", "").strip()
        
        parameters = [str(x).strip() for x in df_raw.iloc[0].tolist()[1:] if pd.notna(x)]
        st.success(f"Detected Factor A: {factor_a_label} | Factor B: {factor_b_label}")
        
        classified_cols = {"Vegetative Properties": [], "Reproductive, Biochemical, and Quality Properties": []}
        for c in parameters:
            cat = classify_parameter(c)
            classified_cols[cat].append(c)
            
        st.write("#### Automatically Categorized Parameter Divisions:")
        for k, v in classified_cols.items():
            if v:
                st.write(f"**{k}:** {', '.join(v)}")
        
        if st.button("Generate Word Document Draft from Direct Output", key="btn_2f_sum_gen"):
            results_data = parse_summarized_table_to_results_2f(
                df_raw, idx_A, idx_B, idx_cv, idx_interaction, idx_grand,
                idx_A_sem, idx_A_f, idx_A_lsd, idx_B_sem, idx_B_f, idx_B_lsd,
                factor_a_levels, factor_b_levels, parameters
            )
            
            doc = Document()
            doc.add_heading("Calculated Two-Factor Factorial RCBD Report", 0)
            
            table_counter = 1
            for cat_name, cat_params in classified_cols.items():
                if not cat_params:
                    continue
                
                doc.add_heading(cat_name, level=1)
                st.write(f"### {cat_name}")
                
                grouped = group_parameters(cat_params)
                
                # Step 1: Render and group static parameters (Up to 4)
                static_items = []
                for base_name, items in sorted(grouped.items()):
                    if len(items) == 1:
                        static_items.append(items[0][0])
                        
                static_chunks = [static_items[i:i + 4] for i in range(0, len(static_items), 4)]
                
                for chunk in static_chunks:
                    chunk_lbl = f"Table {table_counter}"
                    table_counter += 1
                    
                    st.write(f"##### {chunk_lbl}: Integrated Properties")
                    doc.add_heading(f"{chunk_lbl}: Properties Evaluation", level=2)
                    
                    for p in chunk:
                        p_text = generate_two_factor_explanation(p, results_data[p], factor_a_label, factor_b_label, chunk_lbl)
                        st.write(p_text)
                        
                        p_docx = doc.add_paragraph()
                        parts = re.split(r'(\*\*.*?\*\*)', p_text)
                        for part in parts:
                            if part.startswith('**') and part.endswith('**'):
                                p_docx.add_run(part[2:-2]).bold = True
                            else:
                                p_docx.add_run(part)
                                
                    add_excel_table_to_docx(doc, factor_a_label, factor_b_label, chunk, factor_a_levels, factor_b_levels, results_data)
                    st.write("*(Consolidated table placed directly below paragraph)*")
                    doc.add_page_break()
                    
                # Step 2: Render and isolate trend lines
                for base_name, items in sorted(grouped.items()):
                    if len(items) > 1:
                        trend_lbl = f"Table {table_counter}"
                        table_counter += 1
                        
                        st.write(f"##### {trend_lbl}: Progressive Trend of {base_name}")
                        doc.add_heading(f"{trend_lbl}: Progressive Trend of {base_name}", level=2)
                        
                        p_text = generate_trend_explanation_2f(base_name, items, results_data, factor_a_label, factor_b_label, trend_lbl)
                        st.write(p_text)
                        
                        p_docx = doc.add_paragraph()
                        parts = re.split(r'(\*\*.*?\*\*)', p_text)
                        for part in parts:
                            if part.startswith('**') and part.endswith('**'):
                                p_docx.add_run(part[2:-2]).bold = True
                            else:
                                p_docx.add_run(part)
                                
                        trend_params = [it[0] for it in items]
                        add_excel_table_to_docx(doc, factor_a_label, factor_b_label, trend_params, factor_a_levels, factor_b_levels, results_data)
                        st.write("*(Time-series table placed directly below trend paragraph)*")
                        doc.add_page_break()
                        
            bio_doc = io.BytesIO()
            doc.save(bio_doc)
            bio_doc.seek(0)
            
            st.write("---")
            st.markdown("#### 💾 Save Explanations & Tables as Word Document")
            st.download_button(
                "Download Word Explanations Report (.docx)", 
                data=bio_doc, 
                file_name="MultiYear_Summarized_Thesis_Report.docx", 
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
                key="btn_d_2f_sum_doc"
            )
    except Exception as e:
        st.error(f"Error parsing direct result summary table: {e}")
