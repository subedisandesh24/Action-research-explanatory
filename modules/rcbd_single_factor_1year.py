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

# --- Dynamic Academic Explanation Generators for Single Factor ---
def generate_single_trend_explanation(base_name, items, results_data, table_label):
    """
    Generates a publication-grade paragraph explaining the temporal trend of a group of variables
    (e.g., weight loss over 10 days) for single factor design under one topic.
    """
    first_item = items[0]
    last_item = items[-1]
    
    first_param, _, first_day_str = first_item
    last_param, _, last_day_str = last_item
    
    p_first = results_data[first_param]
    p_last = results_data[last_param]
    
    first_gm = p_first["gm"]
    last_gm = p_last["gm"]
    
    direction = "progressive increase" if last_gm >= first_gm else "progressive decrease"
    trend_verb = "increased" if last_gm >= first_gm else "decreased"
    
    sig_first_text = "statistically significant" if p_first["p_val"] < 0.05 else "nonsignificant"
    sig_last_text = "significant" if p_last["p_val"] < 0.05 else "nonsignificant"
    
    sorted_last = sorted(p_last["means"].items(), key=lambda x: x[1], reverse=True)
    top_last, top_val_last = sorted_last[0]
    low_last, low_val_last = sorted_last[-1]
    top_let_last = p_last["means_str"][top_last].replace(f"{top_val_last:.2f}", "")
    
    at_par_last_list = []
    for genotype, val in sorted_last[1:]:
        let = p_last["means_str"][genotype].replace(f"{val:.2f}", "")
        if top_let_last and let and any(char in top_let_last for char in let):
            at_par_last_list.append(f"{genotype} ({val:.2f}^{let})")
    at_par_last_str = f"statistically at par with {', '.join(at_par_last_list)}" if at_par_last_list else "distinctly superior to other genotypes"
    
    para = (
        f"With respect to **{base_name}**, the values displayed a continuous {direction} trend "
        f"as storage/duration progressed (as detailed in **{table_label}**). The grand mean {trend_verb} from "
        f"{first_gm:.2f} at {first_day_str} to {last_gm:.2f} by {last_day_str}. "
        f"The treatment factor exerted a {sig_first_text} effect at the initial observation point, which evolved into "
        f"a highly {sig_last_text} factor by the final evaluation stage ({last_day_str}). At {last_day_str}, the highest "
        f"mean was recorded under `{top_last}` ({top_val_last:.2f}^{top_let_last}), which was {at_par_last_str}. "
        f"Conversely, the minimum performance baseline was observed in `{low_last}` ({low_val_last:.2f})."
    )
    return para

def generate_multiyear_trend_explanation(base_name, items, results_data_1, results_data_2, year1_label, year2_label, table_label):
    """
    Generates an academic paragraph explaining a time-series parameter across two consecutive seasons.
    """
    first_item = items[0]
    last_item = items[-1]
    first_param, _, first_day_str = first_item
    last_param, _, last_day_str = last_item
    
    p_last_1 = results_data_1[last_param]
    p_last_2 = results_data_2[last_param]
    
    first_gm = (results_data_1[first_param]["gm"] + results_data_2[first_param]["gm"]) / 2
    last_gm = (p_last_1["gm"] + p_last_2["gm"]) / 2
    direction = "progressive increase" if last_gm >= first_gm else "progressive decrease"
    trend_verb = "increased" if last_gm >= first_gm else "decreased"
    
    sig_1 = "significant" if p_last_1["p_val"] < 0.05 else "nonsignificant"
    sig_2 = "significant" if p_last_2["p_val"] < 0.05 else "nonsignificant"
    
    sorted_1 = sorted(p_last_1["means"].items(), key=lambda x: x[1], reverse=True)
    top_1, top_val_1 = sorted_1[0]
    top_let_1 = p_last_1["means_str"][top_1].replace(f"{top_val_1:.2f}", "")
    
    sorted_2 = sorted(p_last_2["means"].items(), key=lambda x: x[1], reverse=True)
    top_2, top_val_2 = sorted_2[0]
    top_let_2 = p_last_2["means_str"][top_2].replace(f"{top_val_2:.2f}", "")
    
    para = (
        f"Evaluating the temporal development of **{base_name}** across two consecutive seasons "
        f"({year1_label} and {year2_label}) revealed a highly consistent, time-dependent {direction} "
        f"profile (as shown in **{table_label}**). The overall pooled grand mean {trend_verb} from {first_gm:.2f} at {first_day_str} "
        f"to {last_gm:.2f} by the final interval ({last_day_str}). In the first season ({year1_label}), treatment effects at "
        f"{last_day_str} were highly {sig_1}, with `{top_1}` establishing the highest value of {top_val_1:.2f}^{top_let_1}. "
        f"This pattern was closely mirrored in the second season ({year2_label}), where treatment effects at {last_day_str} "
        f"remained highly {sig_2}, with `{top_2}` occupying the leading statistical group ({top_val_2:.2f}^{top_let_2})."
    )
    return para

def generate_single_explanation(param_name, p_data, table_label):
    """
    Generates Q1-standard paragraph explaining single-interval parameters (e.g., Shelf Life, pH).
    """
    p_val = p_data["p_val"]
    sig_text = "significant" if p_val < 0.05 else "nonsignificant"
    p_not = f"(p < 0.05)" if p_val < 0.05 else "(p > 0.05)"
    if p_val < 0.01: p_not = "(p < 0.01)"
    
    sorted_means = sorted(p_data["means"].items(), key=lambda x: x[1], reverse=True)
    top_g, top_val_g = sorted_means[0]
    low_g, low_val_g = sorted_means[-1]
    top_let_g = p_data["means_str"][top_g].replace(f"{top_val_g:.2f}", "")
    
    at_par_list = []
    for genotype, val in sorted_means[1:]:
        let = p_data["means_str"][genotype].replace(f"{val:.2f}", "")
        if top_let_g and let and any(char in top_let_g for char in let):
            at_par_list.append(f"{genotype} ({val:.2f}^{let})")
    at_par_str = f"statistically at par with {', '.join(at_par_list)}" if at_par_list else "distinctly superior to alternative genotypes"
    
    if p_val > 0.05:
        para = (
            f"The statistical analysis regarding **{param_name}** indicated a **nonsignificant** "
            f"treatment effect {p_not} (as summarized in **{table_label}**). The genotype values remained "
            f"highly stable and uniform, clustering closely around the grand mean of {p_data['gm']:.2f}. "
            f"This suggests that the trial treatments did not induce phenotypic stratification for this specific trait."
        )
    else:
        para = (
            f"Statistical evaluation of the **{param_name}** data revealed a highly **significant** "
            f"treatment effect {p_not} (as shown in **{table_label}**). Genotype `{top_g}` registered "
            f"the maximum value of {top_val_g:.2f}^{top_let_g}, which was {at_par_str}. "
            f"Conversely, the minimum performance value was restricted to genotype `{low_g}` ({low_val_g:.2f}). "
            f"The trial was characterized by a Coefficient of Variation (CV) of {p_data['cv']:.2f}%."
        )
    return para

def generate_multiyear_explanation(param_name, p_data_1, p_data_2, year1_label, year2_label, table_label):
    """
    Generates an explanation paragraph comparing a single-interval parameter across two consecutive years.
    """
    sig_1 = "significant" if p_data_1["p_val"] < 0.05 else "nonsignificant"
    sig_2 = "significant" if p_data_2["p_val"] < 0.05 else "nonsignificant"
    
    sorted_1 = sorted(p_data_1["means"].items(), key=lambda x: x[1], reverse=True)
    top_1, top_val_1 = sorted_1[0]
    top_let_1 = p_data_1["means_str"][top_1].replace(f"{top_val_1:.2f}", "")
    
    sorted_2 = sorted(p_data_2["means"].items(), key=lambda x: x[1], reverse=True)
    top_2, top_val_2 = sorted_2[0]
    top_let_2 = p_data_2["means_str"][top_2].replace(f"{top_val_2:.2f}", "")
    
    pooled_gm = (p_data_1["gm"] + p_data_2["gm"]) / 2
    
    para = (
        f"Evaluation of **{param_name}** across consecutive seasons revealed that genotype effects "
        f"were highly {sig_1} during {year1_label} and highly {sig_2} during {year2_label} "
        f"(as presented in **{table_label}**). In {year1_label}, `{top_1}` achieved the maximum value of "
        f"{top_val_1:.2f}^{top_let_1}, while in {year2_label}, `{top_2}` maintained its superior performance "
        f"with {top_val_2:.2f}^{top_let_2}. Across both seasons, the pooled grand mean settled at {pooled_gm:.2f}, "
        f"demonstrating a reliable and repeatable response pattern under varying environmental environments."
    )
    return para

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

# --- Styled Excel Builders ---
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
            
    # Borders
    for col in range(1, len(params) + 2):
        ws.cell(row=1, column=col).border = border_thin_bottom
        ws.cell(row=start_stats_row, column=col).border = border_thin_bottom
        ws.cell(row=start_stats_row + len(stats_keys) - 1, column=col).border = border_medium_bottom
        
    return wb

def build_multiyear_excel(genotype_col, params, genotypes, results_1, results_2, year1_lbl, year2_lbl):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Analysis Output"
    
    font_bold = Font(name="Arial", size=10, bold=True)
    font_regular = Font(name="Arial", size=10)
    align_left = Alignment(horizontal="left", vertical="center")
    align_center = Alignment(horizontal="center", vertical="center")
    border_thin_bottom = Border(bottom=Side(style='thin', color='000000'))
    border_medium_bottom = Border(bottom=Side(style='medium', color='000000'))
    
    ws.cell(row=3, column=1, value="Genotype").font = font_bold
    ws.cell(row=3, column=1).alignment = align_left
    
    # Header Columns Setup
    for i, p in enumerate(params):
        start_col = i * 3 + 2
        ws.merge_cells(start_row=1, start_column=start_col, end_row=1, end_column=start_col+2)
        cell_p = ws.cell(row=1, column=start_col, value=p.upper())
        cell_p.font = font_bold
        cell_p.alignment = align_center
        
        ws.cell(row=2, column=start_col, value=year1_lbl).font = font_bold
        ws.cell(row=2, column=start_col).alignment = align_center
        ws.cell(row=2, column=start_col+1, value=year2_lbl).font = font_bold
        ws.cell(row=2, column=start_col+1).alignment = align_center
        ws.cell(row=2, column=start_col+2, value="Pooled").font = font_bold
        ws.cell(row=2, column=start_col+2).alignment = align_center
        
        ws.cell(row=3, column=start_col, value="Mean").font = font_bold
        ws.cell(row=3, column=start_col).alignment = align_center
        ws.cell(row=3, column=start_col+1, value="Mean").font = font_bold
        ws.cell(row=3, column=start_col+1).alignment = align_center
        ws.cell(row=3, column=start_col+2, value="Mean").font = font_bold
        ws.cell(row=3, column=start_col+2).alignment = align_center
        
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
            
            cell3 = ws.cell(row=r_idx, column=start_col+2, value=f"{pooled_val:.2f}")
            cell3.font = font_regular
            cell3.alignment = align_center
            
    stats_labels = ["SEM", "P-value", "LSD (0.05)", "CV (%)", "Grand Mean"]
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
            
            # Pooled Stats Placeholder Column left empty as per agricultural layouts
            cell3 = ws.cell(row=curr_row, column=start_col+2, value="")
            cell3.font = font_regular
            cell3.alignment = align_center
            
    total_cols = len(params) * 3 + 1
    for col in range(1, total_cols + 1):
        ws.cell(row=1, column=col).border = border_thin_bottom
        ws.cell(row=3, column=col).border = border_thin_bottom
        ws.cell(row=start_stats_row, column=col).border = border_thin_bottom
        ws.cell(row=start_stats_row + len(stats_keys) - 1, column=col).border = border_medium_bottom
        
    return wb

# --- DOCX Exporter Formatting Functions ---
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
    
    # Data rows
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

def add_multi_table_to_docx(doc, params, genotypes, results_1, results_2, year1_lbl, year2_lbl):
    num_cols = len(params) * 3 + 1
    table = doc.add_table(rows=3, cols=num_cols)
    set_table_borders(table)
    
    # Merge cells dynamically to replicate multi-year Excel layouts in Word
    hdr_row0 = table.rows[0].cells
    hdr_row1 = table.rows[1].cells
    hdr_row2 = table.rows[2].cells
    
    hdr_row0[0].text = "Genotype"
    set_cell_margins(hdr_row0[0])
    hdr_row0[0].paragraphs[0].runs[0].font.bold = True
    hdr_row0[0].paragraphs[0].runs[0].font.name = 'Arial'
    hdr_row0[0].paragraphs[0].runs[0].font.size = Pt(10)
    
    # Merge vertical headers for Treatment Column
    hdr_row0[0].merge(hdr_row1[0]).merge(hdr_row2[0])
    
    for i, p in enumerate(params):
        start_col = i * 3 + 1
        
        # Merge horizontal headers for Parameters
        hdr_row0[start_col].merge(hdr_row0[start_col+1]).merge(hdr_row0[start_col+2])
        hdr_row0[start_col].text = p.upper()
        set_cell_margins(hdr_row0[start_col])
        hdr_row0[start_col].paragraphs[0].alignment = 1
        hdr_row0[start_col].paragraphs[0].runs[0].font.bold = True
        hdr_row0[start_col].paragraphs[0].runs[0].font.name = 'Arial'
        hdr_row0[start_col].paragraphs[0].runs[0].font.size = Pt(10)
        
        # Subheaders (Yearly labels and pooled tags)
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
        
        hdr_row1[start_col+2].text = "Pooled"
        set_cell_margins(hdr_row1[start_col+2])
        hdr_row1[start_col+2].paragraphs[0].alignment = 1
        hdr_row1[start_col+2].paragraphs[0].runs[0].font.bold = True
        hdr_row1[start_col+2].paragraphs[0].runs[0].font.name = 'Arial'
        hdr_row1[start_col+2].paragraphs[0].runs[0].font.size = Pt(10)
        
        # Data categories row (Mean markers)
        for sub_col in range(3):
            hdr_row2[start_col + sub_col].text = "Mean"
            set_cell_margins(hdr_row2[start_col + sub_col])
            hdr_row2[start_col + sub_col].paragraphs[0].alignment = 1
            hdr_row2[start_col + sub_col].paragraphs[0].runs[0].font.bold = True
            hdr_row2[start_col + sub_col].paragraphs[0].runs[0].font.name = 'Arial'
            hdr_row2[start_col + sub_col].paragraphs[0].runs[0].font.size = Pt(10)
            
    set_header_bottom_border(table.rows[2])
    
    # Fill Data rows
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
    
    stats_labels = ["SEM", "P-value", "LSD (0.05)", "CV (%)", "Grand Mean"]
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
            
            row_cells[start_col+2].text = ""
            set_cell_margins(row_cells[start_col+2])
            
        if s_idx == len(stats_keys) - 1:
            set_header_bottom_border(row_cells)

# --- Streamlit Controller ---
def run_one_factor_analysis():
    st.markdown("### Single-Factor RCBD Analyzer")
    years_option = st.radio("Select Trial Duration:", ["1 Year", "2 Years / Multi-Year Pooling"], key="1f_duration")
    
    file1 = st.file_uploader("Upload Trial File 1 (.xlsx)", type=["xlsx"], key="file1_1f")
    
    file2 = None
    year1_lbl = "Year 1"
    year2_lbl = "Year 2"
    
    if years_option == "2 Years / Multi-Year Pooling":
        file2 = st.file_uploader("Upload Trial File 2 (.xlsx)", type=["xlsx"], key="file2_1f")
        year1_lbl = st.text_input("Label for Year 1 (Optional):", value="Year 1")
        year2_lbl = st.text_input("Label for Year 2 (Optional):", value="Year 2")
        
    if file1 is not None:
        try:
            df1_raw = pd.read_excel(file1)
            st.write("#### Trial 1 Preview:", df1_raw.head())
            
            cols = df1_raw.columns.tolist()
            block_col = st.selectbox("Select Block/Replication Column:", cols, index=0, key="block_1f")
            genotype_col = st.selectbox("Select Genotype/Treatment Column:", cols, index=1, key="genotype_1f")
            response_cols = st.multiselect("Select Parameters to Analyze:", cols, default=cols[2:], key="response_1f")
            
            if response_cols:
                # Custom layout subdivisions
                group_1_name = st.text_input("First Table Title", "Physiological and Morphological Properties", key="1f_t1_title")
                group_1_cols = st.multiselect("Select parameters for Table 1", response_cols, default=response_cols[:len(response_cols)//2], key="1f_t1_cols")
                group_2_name = st.text_input("Second Table Title", "Biochemical and Quality Properties", key="1f_t2_title")
                group_2_cols = st.multiselect("Select parameters for Table 2", [c for c in response_cols if c not in group_1_cols], default=[c for c in response_cols if c not in group_1_cols], key="1f_t2_cols")
                
                if st.button("Run Single-Factor Statistical Engine", key="run_1f_engine"):
                    genotypes_1 = sorted(df1_raw[genotype_col].dropna().unique().tolist())
                    
                    results_data_1 = {}
                    for param in response_cols:
                        results_data_1[param] = run_anova_1factor(df1_raw, block_col, genotype_col, param)
                        
                    results_data_2 = {}
                    genotypes_2 = []
                    common_genotypes = genotypes_1
                    
                    if file2 is not None:
                        df2_raw = pd.read_excel(file2)
                        genotypes_2 = sorted(df2_raw[genotype_col].dropna().unique().tolist())
                        common_genotypes = sorted(list(set(genotypes_1).intersection(set(genotypes_2))))
                        
                        if not common_genotypes:
                            st.error("Error: No matching treatment/genotype tags found between both excel sheets.")
                            return
                            
                        for param in response_cols:
                            results_data_2[param] = run_anova_1factor(df2_raw, block_col, genotype_col, param)
                            
                    # 1. EXCEL WORKBOOK EXPORT
                    excel_bio = io.BytesIO()
                    if file2 is None:
                        styled_wb = build_single_year_excel(genotype_col, response_cols, common_genotypes, results_data_1)
                    else:
                        styled_wb = build_multiyear_excel(genotype_col, response_cols, common_genotypes, results_data_1, results_data_2, year1_lbl, year2_lbl)
                        
                    styled_wb.save(excel_bio)
                    excel_bio.seek(0)
                    
                    st.markdown("#### 📥 Download Formatted Statistical Excel Results")
                    st.download_button(
                        label="Download Excel Results Sheet",
                        data=excel_bio,
                        file_name="RCBD_1Factor_Output.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="btn_d_excel_1f"
                    )
                    st.write("---")
                    
                    # 2. DISCUSSIONS & ACADEMIC DOCUMENT BUILDER
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
                        
                        # Generate text discussions
                        for base_name, items in sorted(grouped.items()):
                            st.write(f"#### {base_name}")
                            doc.add_heading(base_name, level=3)
                            
                            if len(items) > 1:
                                if file2 is None:
                                    p_text = generate_single_trend_explanation(base_name, items, results_data_1, table_label)
                                else:
                                    p_text = generate_multiyear_trend_explanation(base_name, items, results_data_1, results_data_2, year1_lbl, year2_lbl, table_label)
                            else:
                                if file2 is None:
                                    p_text = generate_single_explanation(items[0][0], results_data_1[items[0][0]], table_label)
                                else:
                                    p_text = generate_multiyear_explanation(items[0][0], results_data_1[items[0][0]], results_data_2[items[0][0]], year1_lbl, year2_lbl, table_label)
                                    
                            st.write(p_text)
                            
                            p_docx = doc.add_paragraph()
                            parts = re.split(r'(\*\*.*?\*\*)', p_text)
                            for part in parts:
                                if part.startswith('**') and part.endswith('**'):
                                    p_docx.add_run(part[2:-2]).bold = True
                                else:
                                    p_docx.add_run(part)
                                    
                        st.write("##### Corresponding Table Visualization:")
                        if file2 is None:
                            add_single_table_to_docx(doc, g_cols, common_genotypes, results_data_1)
                        else:
                            add_multi_table_to_docx(doc, g_cols, common_genotypes, results_data_1, results_data_2, year1_lbl, year2_lbl)
                            
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
                        file_name="SingleFactor_Thesis_Report.docx", 
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
                        key="btn_d_1f_doc"
                    )
        except Exception as e:
            st.error(f"Analysis failed. Please check the structure of your Excel dataset: {e}")

# --- Parent Entrypoint ---
if __name__ == '__main__':
    st.set_page_config(layout="wide")
    run_one_factor_analysis()
