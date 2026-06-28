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
# DATABASE OF 20 HIGH-STANDARD ACADEMIC THESIS & Q1 JOURNAL TEMPLATES
# ==============================================================================
ACADEMIC_TEMPLATES = {
    "temp_1_sig_yield": (
        "The evaluation of **{parameter}** demonstrated that genotypes exerted a highly significant influence "
        "({p_val}) on this trait, as presented in **{table_label}**. The highest mean performance was recorded "
        "under `{top_g}` ({top_val}^{top_let}), which was statistically at par with {at_par}. This highlights "
        "the substantial genetic variation present among the evaluated treatments."
    ),
    "temp_2_sig_quality": (
        "With respect to **{parameter}**, a significant treatment effect was observed ({p_val}) (as shown in "
        "**{table_label}**). Genotype `{top_g}` ({top_val}^{top_let}) proved statistically superior, sharing "
        "statistical parity only with {at_par}. This variation suggests that structural retention and tissue quality "
        "were uniquely preserved under this genotype."
    ),
    "temp_3_strict_superiority": (
        "For the parameter **{parameter}**, the treatment factor induced a highly significant response ({p_val}). "
        "Genotype `{top_g}` recorded `{top_val}^{top_let}`, establishing clear statistical superiority over all "
        "other treatments (as summarized in **{table_label}**). The closest competitor was `{second_g}` (`{second_val}`), "
        "indicating that `{top_g}` possesses a highly efficient phenotypic expression for this trait."
    ),
    "temp_4_homogeneous": (
        "The statistical analysis of **{parameter}** revealed a nonsignificant treatment effect ({p_val}), as detailed "
        "in **{table_label}**. The treatment means remained highly stable around the grand mean of `{grand_mean}`, "
        "suggesting that the evaluated genotypes possess comparable physiological and chemical stability for this trait."
    ),
    "temp_5_marginal_sig": (
        "The trait **{parameter}** was significantly influenced by genotype variations ({p_val}) (as shown in **{table_label}**). "
        "Genotype `{top_g}` led the performance with `{top_val}^{top_let}`, and shared statistical letters with {at_par}. "
        "The lower baseline limit was defined by `{low_g}` (`{low_val}`), indicating a moderate range of phenotypic expression."
    ),
    "temp_6_precision_verification": (
        "Regarding **{parameter}**, the analysis of variance revealed a significant treatment effect ({p_val}). The overall "
        "precision of the trial was verified by a low Coefficient of Variation (CV) of `{cv}%` and a Standard Error of "
        "the Mean (SEm) of `{sem}` (Table **{table_label}**), demonstrating that background spatial noise was managed under "
        "the block partitions."
    ),
    "temp_7_broad_parity": (
        "For **{parameter}**, the genotype effect was significant ({p_val}) (as summarized in **{table_label}**). "
        "Genotype `{top_g}` ({top_val}^{top_let}) occupied the top statistical tier but did not differ significantly "
        "from a broad range of genotypes, including {at_par}. This suggests a wide genetic buffer for this trait."
    ),
    "temp_8_extreme_contrast": (
        "The response of **{parameter}** was highly dependent on the treatment genotypes ({p_val}) (Table **{table_label}**). "
        "A sharp contrast was observed between the leading genotype `{top_g}` ({top_val}^{top_let}) and the lowest-performing "
        "genotype `{low_g}` ({low_val}), representing a highly pronounced phenotypic difference."
    ),
    "temp_9_progressive_upward_trend": (
        "Concerning the progressive changes in **{base_name}**, a distinct time-dependent upward trend was observed, "
        "as summarized in **{table_label}**. The grand mean of the trial increased from `{first_gm}` at `{first_day}` "
        "to `{last_gm}` by `{last_day}`. The genotype effect was `{first_sig}` early on, but developed into a highly "
        "significant (`{last_sig}`) effect by `{last_day}`, with `{top_g}` leading at `{top_val}^{top_let}`."
    ),
    "temp_10_progressive_decline_trend": (
        "For **{base_name}**, a systematic, progressive decline was observed across the storage period, as detailed in "
        "**{table_label}**. The pooled averages fell from `{first_gm}` at `{first_day}` to `{last_gm}` by `{last_day}`. "
        "Highly significant genotype differences ({last_p}) were recorded on `{last_day}`, with `{top_g}` ({top_val}^{top_let}) "
        "demonstrating optimal structural retention compared to `{low_g}`."
    ),
    "temp_11_late_onset_divergence": (
        "The progressive evaluation of **{base_name}** revealed a late-onset treatment divergence. While genotype "
        "differences were nonsignificant on `{first_day}` ({first_p}), they became highly significant on `{last_day}` "
        "({last_p}) (as presented in **{table_label}**), where `{top_g}` ({top_val}^{top_let}) established its statistical lead."
    ),
    "temp_12_early_onset_convergence": (
        "For **{base_name}**, the initial treatment differences observed at `{first_day}` ({first_p}) converged over time, "
        "becoming nonsignificant on `{last_day}` ({last_p}) (as shown in **{table_label}**). This indicates that long-term "
        "exposure or storage leveled out genotype-specific variations."
    ),
    "temp_13_uniform_trend": (
        "Although **{base_name}** changed progressively from `{first_gm}` to `{last_gm}` over the course of the trial (Table "
        "**{table_label}**), the genotype main effect remained consistently nonsignificant ({last_p}) across all intervals, "
        "confirming highly uniform behavioral patterns."
    ),
    "temp_14_stabilization_trend": (
        "Regarding **{base_name}**, a progressive stabilization pattern was observed in the latter half of the trial, "
        "as shown in **{table_label}**. The values changed sharply from `{first_day}` to `{mid_day}`, but stabilized by "
        "`{last_day}`, where genotype `{top_g}` ({top_val}^{top_let}) maintained its statistical lead."
    ),
    "temp_15_decay_progression": (
        "Decay progression for **{base_name}** rose over the evaluation intervals, escalating from `{first_gm}` to `{last_gm}` "
        "(Table **{table_label}**). Genotypes significantly influenced decay development by `{last_day}` ({last_p}), with "
        "`{top_g}` successfully suppressing decay loss ({top_val}^{top_let}) compared to the control `{low_g}`."
    ),
    "temp_16_performance_tier": (
        "The post-hoc grouping for **{parameter}** clearly differentiated the genotypes into distinct performance tiers, "
        "as detailed in **{table_label}**. Genotype `{top_g}` ({top_val}^{top_let}) led the elite tier, making it a promising "
        "candidate for future breeding or selection programs."
    ),
    "temp_17_stress_conservation": (
        "Under the trial conditions, **{parameter}** was preserved best under genotype `{top_g}` ({top_val}^{top_let}) (Table "
        "**{table_label}**), demonstrating superior cellular buffering or metabolic stability. This genotype was statistically "
        "at par with {at_par}."
    ),
    "temp_18_buffered_expression": (
        "Trait expression for **{parameter}** was highly stable, with nonsignificant genotype effects ({p_val}), as detailed "
        "in **{table_label}**. The minimal variance among treatments is supported by a very low CV of `{cv}%`, proving "
        "that this trait is highly buffered against genetic differences."
    ),
    "temp_19_biochemical_nutrient": (
        "Genotype `{top_g}` ({top_val}^{top_let}) exhibited the highest levels of **{parameter}** (as shown in **{table_label}**), "
        "indicating a highly active metabolic pathway for this compound. It shared statistical parity only with {at_par}."
    ),
    "temp_20_comprehensive_candidate": (
        "In conclusion, genotype `{top_g}` proved to be the most promising candidate for the optimization of **{parameter}** "
        "(as presented in **{table_label}**). It combined high statistical performance ({top_val}^{top_let}) with strong "
        "experimental precision, establishing its superiority."
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
def generate_single_trend_explanation(base_name, items, results_data, table_label):
    """
    Generates a Q1-standard paragraph explaining the temporal trend of a group of variables
    using templates 9, 10, 11, 12, 13, 14, or 15 dynamically based on statistical profile.
    """
    first_item = items[0]
    last_item = items[-1]
    
    first_param, _, first_day_str = first_item
    last_param, _, last_day_str = last_item
    
    p_first = results_data[first_param]
    p_last = results_data[last_param]
    
    first_gm = p_first["gm"]
    last_gm = p_last["gm"]
    
    # Determine basic directions
    direction = "upward" if last_gm >= first_gm else "downward"
    
    sorted_last = sorted(p_last["means"].items(), key=lambda x: x[1], reverse=True)
    top_last, top_val_last = sorted_last[0]
    low_last, low_val_last = sorted_last[-1]
    top_let_last = p_last["means_str"][top_last].replace(f"{top_val_last:.2f}", "")
    
    at_par_last_list = []
    for genotype, val in sorted_last[1:]:
        let = p_last["means_str"][genotype].replace(f"{val:.2f}", "")
        if top_let_last and let and any(char in top_let_last for char in let):
            at_par_last_list.append(f"`{genotype}` ({val:.2f}^{let})")
    at_par_last_str = ", ".join(at_par_last_list) if at_par_last_list else "no other genotypes"
    
    # Selection of Trend Templates
    if direction == "upward" and p_first["p_val"] >= 0.05 and p_last["p_val"] < 0.05:
        # Template 11: Late-Onset Divergence (Upward)
        template = ACADEMIC_TEMPLATES["temp_11_late_onset_divergence"]
        return template.format(
            base_name=base_name, first_day=first_day_str, last_day=last_day_str,
            first_p=f"p = {p_first['p_val']:.4f}", last_p=f"p = {p_last['p_val']:.4f}",
            top_g=top_last, top_val=f"{top_val_last:.2f}", top_let=top_let_last, table_label=table_label
        )
    elif direction == "downward" and p_first["p_val"] < 0.05 and p_last["p_val"] >= 0.05:
        # Template 12: Early-Onset with Late Convergence
        template = ACADEMIC_TEMPLATES["temp_12_early_onset_convergence"]
        return template.format(
            base_name=base_name, first_day=first_day_str, last_day=last_day_str,
            first_p=f"p = {p_first['p_val']:.4f}", last_p=f"p = {p_last['p_val']:.4f}",
            table_label=table_label
        )
    elif p_first["p_val"] >= 0.05 and p_last["p_val"] >= 0.05:
        # Template 13: Strictly Uniform Trend
        template = ACADEMIC_TEMPLATES["temp_13_uniform_trend"]
        return template.format(
            base_name=base_name, first_gm=f"{first_gm:.2f}", last_gm=f"{last_gm:.2f}",
            last_p=f"p = {p_last['p_val']:.4f}", table_label=table_label
        )
    elif "decay" in base_name.lower() or "rot" in base_name.lower():
        # Template 15: Decay Progression
        template = ACADEMIC_TEMPLATES["temp_15_decay_progression"]
        return template.format(
            base_name=base_name, first_gm=f"{first_gm:.2f}", last_gm=f"{last_gm:.2f}",
            last_p=f"p = {p_last['p_val']:.4f}", top_g=top_last, top_val=f"{top_val_last:.2f}",
            top_let=top_let_last, low_g=low_last, table_label=table_label
        )
    elif direction == "upward":
        # Template 9: Progressive Upward Trend
        template = ACADEMIC_TEMPLATES["temp_9_progressive_upward_trend"]
        first_sig_txt = "significant" if p_first["p_val"] < 0.05 else "nonsignificant"
        last_sig_txt = f"p = {p_last['p_val']:.4f}"
        return template.format(
            base_name=base_name, first_gm=f"{first_gm:.2f}", last_gm=f"{last_gm:.2f}",
            first_day=first_day_str, last_day=last_day_str, first_sig=first_sig_txt,
            last_sig=last_sig_txt, top_g=top_last, top_val=f"{top_val_last:.2f}",
            top_let=top_let_last, table_label=table_label
        )
    else:
        # Template 10: Progressive Decline Trend
        template = ACADEMIC_TEMPLATES["temp_10_progressive_decline_trend"]
        return template.format(
            base_name=base_name, first_gm=f"{first_gm:.2f}", last_gm=f"{last_gm:.2f}",
            first_day=first_day_str, last_day=last_day_str, last_p=f"p = {p_last['p_val']:.4f}",
            top_g=top_last, top_val=f"{top_val_last:.2f}", top_let=top_let_last, low_g=low_last,
            table_label=table_label
        )

def generate_single_explanation(param_name, p_data, table_label):
    """
    Selects and renders the best matching academic templates for isolated parameters
    based on the statistical significance and treatment CLD profile.
    """
    p_val = p_data["p_val"]
    
    sorted_means = sorted(p_data["means"].items(), key=lambda x: x[1], reverse=True)
    top_g, top_val_g = sorted_means[0]
    low_g, low_val_g = sorted_means[-1]
    top_let_g = p_data["means_str"][top_g].replace(f"{top_val_g:.2f}", "")
    
    at_par_list = []
    for genotype, val in sorted_means[1:]:
        let = p_data["means_str"][genotype].replace(f"{val:.2f}", "")
        if top_let_g and let and any(char in top_let_g for char in let):
            at_par_list.append(f"`{genotype}` ({val:.2f}^{let})")
    at_par_str = ", ".join(at_par_list) if at_par_list else "no other genotypes"
    
    # Selection logic mapping back to the 20 templates
    if p_val >= 0.05:
        # Template 4: Highly Homogeneous / Non-Significant Response
        template = ACADEMIC_TEMPLATES["temp_4_homogeneous"]
        return template.format(
            parameter=param_name, p_val=f"p = {p_val:.4f}", grand_mean=f"{p_data['gm']:.2f}", table_label=table_label
        )
    elif len(sorted_means) >= 3 and not at_par_list:
        # Template 3: Genotype Leader with No At-Par Competitors
        second_g, second_val = sorted_means[1]
        template = ACADEMIC_TEMPLATES["temp_3_strict_superiority"]
        return template.format(
            parameter=param_name, p_val=f"p = {p_val:.4f}", top_g=top_g, top_val=f"{top_val_g:.2f}",
            top_let=top_let_g, second_g=second_g, second_val=f"{second_val:.2f}", table_label=table_label
        )
    elif "yield" in param_name.lower() or "weight" in param_name.lower() or "output" in param_name.lower():
        # Template 1: Highly Significant Genotype Effect (Yield focus)
        template = ACADEMIC_TEMPLATES["temp_1_sig_yield"]
        return template.format(
            parameter=param_name, p_val=f"p = {p_val:.4f}", top_g=top_g, top_val=f"{top_val_g:.2f}",
            top_let=top_let_g, at_par=at_par_str, table_label=table_label
        )
    elif any(x in param_name.lower() for x in ["ascorbic", "acid", "tss", "sugar", "nutrient", "antioxidant"]):
        # Template 19: Biochemical Activity / Nutrient Stability
        template = ACADEMIC_TEMPLATES["temp_19_biochemical_nutrient"]
        return template.format(
            parameter=param_name, top_g=top_g, top_val=f"{top_val_g:.2f}", top_let=top_let_g,
            at_par=at_par_str, table_label=table_label
        )
    elif len(at_par_list) >= 3:
        # Template 7: Genotype with Broad Parity
        template = ACADEMIC_TEMPLATES["temp_7_broad_parity"]
        return template.format(
            parameter=param_name, p_val=f"p = {p_val:.4f}", top_g=top_g, top_val=f"{top_val_g:.2f}",
            top_let=top_let_g, at_par=at_par_str, table_label=table_label
        )
    else:
        # Template 2: Significant Genotype Effect (Quality focus)
        template = ACADEMIC_TEMPLATES["temp_2_sig_quality"]
        return template.format(
            parameter=param_name, p_val=f"p = {p_val:.4f}", top_g=top_g, top_val=f"{top_val_g:.2f}",
            top_let=top_let_g, at_par=at_par_str, table_label=table_label
        )

# --- Statistical Computation Engine ---
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

# --- Excel Sheet Builder ---
def build_single_year_excel(genotype_col, params, genotypes, results_data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Analysis Output"
    
    font_bold = Font(name="Arial", size=10, bold=True)
    font_regular = Font(name="Arial", size=10)
    align_left = Alignment(horizontal="left", vertical="center")
    align_center = Alignment(horizontal="center", vertical="center")
    border_thin_bottom = Border(bottom=Side(style='thin', color='000000'))
    border_medium_bottom = Border(bottom=Side(style='medium', color='000000'))
    
    ws.cell(row=1, column=1, value="Genotype").font = font_bold
    ws.cell(row=1, column=1).alignment = align_left
    
    for c_idx, p in enumerate(params, start=2):
        cell = ws.cell(row=1, column=c_idx, value=p.upper())
        cell.font = font_bold
        cell.alignment = align_center
        
    for r_idx, g in enumerate(genotypes, start=2):
        ws.cell(row=r_idx, column=1, value=g).font = font_regular
        ws.cell(row=r_idx, column=1).alignment = align_left
        for c_idx, p in enumerate(params, start=2):
            cell = ws.cell(row=r_idx, column=c_idx, value=results_data[p]["means_str"].get(g, ""))
            cell.font = font_regular
            cell.alignment = align_center
            
    stats_labels = ["SEM", "P-value", "LSD (0.05)", "CV (%)", "Grand Mean"]
    stats_keys = ["sem", "p_text", "lsd", "cv", "gm"]
    
    start_stats_row = len(genotypes) + 2
    for s_idx, (label, key) in enumerate(zip(stats_labels, stats_keys)):
        curr_row = start_stats_row + s_idx
        ws.cell(row=curr_row, column=1, value=label).font = font_bold
        ws.cell(row=curr_row, column=1).alignment = align_left
        for c_idx, p in enumerate(params, start=2):
            cell = ws.cell(row=curr_row, column=c_idx, value=results_data[p][key])
            cell.font = font_regular
            cell.alignment = align_center
            
    for col in range(1, len(params) + 2):
        ws.cell(row=1, column=col).border = border_thin_bottom
        ws.cell(row=start_stats_row, column=col).border = border_thin_bottom
        ws.cell(row=start_stats_row + len(stats_keys) - 1, column=col).border = border_medium_bottom
        
    return wb

# --- DOCX Layout Formatting Functions ---
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

def add_single_table_to_docx(doc, params, genotypes, results_data):
    num_cols = len(params) + 1
    table = doc.add_table(rows=1, cols=num_cols)
    set_table_borders(table)
    
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = "Genotype"
    set_cell_margins(hdr_cells[0])
    hdr_cells[0].paragraphs[0].runs[0].font.bold = True
    hdr_cells[0].paragraphs[0].runs[0].font.name = 'Arial'
    hdr_cells[0].paragraphs[0].runs[0].font.size = Pt(10)
    
    for c_idx, p in enumerate(params):
        hdr_cells[c_idx + 1].text = p.upper()
        set_cell_margins(hdr_cells[c_idx + 1])
        hdr_cells[c_idx + 1].paragraphs[0].alignment = 1
        hdr_cells[c_idx + 1].paragraphs[0].runs[0].font.bold = True
        hdr_cells[c_idx + 1].paragraphs[0].runs[0].font.name = 'Arial'
        hdr_cells[c_idx + 1].paragraphs[0].runs[0].font.size = Pt(10)
        
    set_header_bottom_border(table.rows[0])
    
    for g in sorted(genotypes):
        row_cells = table.add_row().cells
        row_cells[0].text = str(g)
        set_cell_margins(row_cells[0])
        row_cells[0].paragraphs[0].runs[0].font.name = 'Arial'
        row_cells[0].paragraphs[0].runs[0].font.size = Pt(10)
        
        for c_idx, p in enumerate(params):
            row_cells[c_idx + 1].text = str(results_data[p]["means_str"].get(g, ""))
            set_cell_margins(row_cells[c_idx + 1])
            p_val = row_cells[c_idx + 1].paragraphs[0]
            p_val.alignment = 1
            if p_val.runs:
                p_val.runs[0].font.name = 'Arial'
                p_val.runs[0].font.size = Pt(10)
                
    set_header_bottom_border(table.rows[-1])
    
    stats_labels = ["SEM", "P-value", "LSD (0.05)", "CV (%)", "Grand Mean"]
    stats_keys = ["sem", "p_text", "lsd", "cv", "gm"]
    
    for s_idx, (label, key) in enumerate(zip(stats_labels, stats_keys)):
        row_cells = table.add_row().cells
        row_cells[0].text = label
        set_cell_margins(row_cells[0])
        row_cells[0].paragraphs[0].runs[0].font.bold = True
        row_cells[0].paragraphs[0].runs[0].font.name = 'Arial'
        row_cells[0].paragraphs[0].runs[0].font.size = Pt(10)
        
        for c_idx, p in enumerate(params):
            row_cells[c_idx + 1].text = str(results_data[p][key])
            set_cell_margins(row_cells[c_idx + 1])
            p_val = row_cells[c_idx + 1].paragraphs[0]
            p_val.alignment = 1
            if p_val.runs:
                p_val.runs[0].font.name = 'Arial'
                p_val.runs[0].font.size = Pt(10)
                
        if s_idx == len(stats_keys) - 1:
            set_header_bottom_border(row_cells)

# --- Parent Layout Web App ---
def show_module():
    st.markdown("### Single-Year Single-Factor RCBD Analyzer")
    
    file1 = st.file_uploader("Upload Raw Data File (.xlsx)", type=["xlsx"], key="file1_1f_sy")
    
    if file1 is not None:
        try:
            df1_raw = pd.read_excel(file1)
            st.write("#### Data Preview:", df1_raw.head())
            
            cols = df1_raw.columns.tolist()
            block_col = st.selectbox("Select Block/Replication Column:", cols, index=0, key="block_1f_sy")
            genotype_col = st.selectbox("Select Genotype/Treatment Column:", cols, index=1, key="genotype_1f_sy")
            response_cols = st.multiselect("Select Response Parameters to Analyze:", cols, default=cols[2:], key="response_1f_sy")
            
            if response_cols:
                group_1_name = st.text_input("First Table Title", "Physiological and Crop Growth Parameters", key="1f_sy_t1_title")
                group_1_cols = st.multiselect("Select parameters for Table 1", response_cols, default=response_cols[:len(response_cols)//2], key="1f_sy_t1_cols")
                group_2_name = st.text_input("Second Table Title", "Crop Quality and Yield Parameters", key="1f_sy_t2_title")
                group_2_cols = st.multiselect("Select parameters for Table 2", [c for c in response_cols if c not in group_1_cols], default=[c for c in response_cols if c not in group_1_cols], key="1f_sy_t2_cols")
                
                if st.button("Execute Statistical Calculations", key="run_1f_sy_engine"):
                    genotypes = sorted(df1_raw[genotype_col].dropna().unique().tolist())
                    
                    results_data = {}
                    for param in response_cols:
                        results_data[param] = run_anova_1factor(df1_raw, block_col, genotype_col, param)
                        
                    # Build Excel Results
                    excel_bio = io.BytesIO()
                    styled_wb = build_single_year_excel(genotype_col, response_cols, genotypes, results_data)
                    styled_wb.save(excel_bio)
                    excel_bio.seek(0)
                    
                    st.markdown("#### 📥 Download Formatted Statistical Excel Results")
                    st.download_button(
                        label="Download Excel Results Sheet",
                        data=excel_bio,
                        file_name="RCBD_SingleYear_Output.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="btn_d_excel_sy"
                    )
                    st.write("---")
                    
                    # Generate on-screen discussions and docx
                    st.markdown("### 📝 Dynamic Analysis Results & Discussions")
                    doc = Document()
                    doc.add_heading("Single-Factor RCBD Comprehensive Trial Report", 0)
                    
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
                            
                            # Temporal/trend-line parameter selection logic vs static single parameter
                            if len(items) > 1:
                                p_text = generate_single_trend_explanation(base_name, items, results_data, table_label)
                            else:
                                p_text = generate_single_explanation(items[0][0], results_data[items[0][0]], table_label)
                                
                            st.write(p_text)
                            
                            # Append formatted paragraph to docx
                            p_docx = doc.add_paragraph()
                            parts = re.split(r'(\*\*.*?\*\*)', p_text)
                            for part in parts:
                                if part.startswith('**') and part.endswith('**'):
                                    p_docx.add_run(part[2:-2]).bold = True
                                else:
                                    p_docx.add_run(part)
                                    
                        st.write("##### Corresponding Table Visualization:")
                        add_single_table_to_docx(doc, g_cols, genotypes, results_data)
                        st.write("*(Table data formatted as per standard journal specifications)*")
                        doc.add_page_break()
                        
                    bio_doc = io.BytesIO()
                    doc.save(bio_doc)
                    bio_doc.seek(0)
                    
                    st.write("---")
                    st.markdown("#### 💾 Save Explanations & Tables as Word Document")
                    st.download_button(
                        "Download Word Explanations Report (.docx)", 
                        data=bio_doc, 
                        file_name="SingleYear_Thesis_Report.docx", 
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
                        key="btn_d_sy_doc"
                    )
        except Exception as e:
            st.error(f"Analysis failed. Please check the structure of your Excel dataset: {e}")

if __name__ == '__main__':
    st.set_page_config(layout="wide")
    show_module()
