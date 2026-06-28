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
            # Clean up trailing non-alphanumeric punctuation
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
            
    # Sort chronologically by day/interval
    for base in groups:
        groups[base].sort(key=lambda x: x[1])
    return groups

# --- Dynamic Academic Explanation Generators ---
def generate_trend_explanation(base_name, items, results_data, factor_a_col, factor_b_col, table_label):
    """
    Generates a Q1-standard paragraph explaining the temporal trend of a group of variables
    (e.g., physiological weight loss over 10 days) as a single cohesive topic.
    """
    first_item = items[0]
    last_item = items[-1]
    
    first_param, _, first_day_str = first_item
    last_param, _, last_day_str = last_item
    
    p_first = results_data[first_param]
    p_last = results_data[last_param]
    
    first_gm = p_first["gm"]
    last_gm = p_last["gm"]
    
    # Identify trend direction
    direction = "progressive increase" if last_gm >= first_gm else "progressive decrease"
    trend_verb = "increased" if last_gm >= first_gm else "decreased"
    
    # Factor A first/last evaluations
    sig_a_first_text = "statistically significant" if p_first["p_a"] < 0.05 else "nonsignificant"
    sig_a_last_text = "significant" if p_last["p_a"] < 0.05 else "nonsignificant"
    
    sorted_a_last = sorted(p_last["means_a"].items(), key=lambda x: x[1], reverse=True)
    top_a_last, top_val_a_last = sorted_a_last[0]
    low_a_last, low_val_a_last = sorted_a_last[-1]
    top_let_a_last = p_last["means_a_str"][top_a_last].replace(f"{top_val_a_last:.2f}", "")
    
    # Factor B first/last evaluations
    sig_b_first_text = "statistically significant" if p_first["p_b"] < 0.05 else "nonsignificant"
    sig_b_last_text = "significant" if p_last["p_b"] < 0.05 else "nonsignificant"
    
    sorted_b_last = sorted(p_last["means_b"].items(), key=lambda x: x[1], reverse=True)
    top_b_last, top_val_b_last = sorted_b_last[0]
    low_b_last, low_val_b_last = sorted_b_last[-1]
    top_let_b_last = p_last["means_b_str"][top_b_last].replace(f"{top_val_b_last:.2f}", "")
    
    # Evaluate interaction evolution across the entire timeline
    interaction_evolution = ""
    any_sig_ab = any(results_data[it[0]]["p_ab"] < 0.05 for it in items)
    if any_sig_ab:
        sig_days = [it[2] for it in items if results_data[it[0]]["p_ab"] < 0.05]
        sorted_comb_last = sorted(p_last["means_comb"].items(), key=lambda x: x[1], reverse=True)
        comb_top_name, comb_top_val = sorted_comb_last[0]
        comb_low_name, comb_low_val = sorted_comb_last[-1]
        comb_top_let = p_last["cld_comb"].get(comb_top_name, "")
        
        interaction_evolution = (
            f"Notably, the interaction effect between `{factor_a_col}` and `{factor_b_col}` demonstrated distinct temporal dependencies, "
            f"transitioning from nonsignificant at the early phase to highly significant during the later stages of observation ({', '.join(sig_days)}). "
            f"At the final evaluation interval ({last_day_str}), the treatment combination `{comb_top_name}` yielded the peak value of {comb_top_val:.2f}^{comb_top_let}, "
            f"while `{comb_low_name}` marked the lowest limit of performance ({comb_low_val:.2f})."
        )
    else:
        interaction_evolution = (
            f"The interaction effect between `{factor_a_col}` and `{factor_b_col}` remained consistently nonsignificant across all "
            f"assessment intervals, demonstrating that both treatment factors regulated the {base_name} trend independently."
        )
        
    para = (
        f"Regarding **{base_name}**, the trait exhibited a clear, time-dependent {direction} trend over the course of the trial "
        f"(as shown in **{table_label}**). The grand mean {trend_verb} from {first_gm:.2f} at {first_day_str} and progressively shifted to {last_gm:.2f} "
        f"by {last_day_str}. The main effect of `{factor_a_col}` was {sig_a_first_text} initially, but developed into a highly {sig_a_last_text} factor "
        f"as storage/duration extended. By {last_day_str}, treatment `{top_a_last}` attained the maximum average value of {top_val_a_last:.2f}^{top_let_a_last}, "
        f"whereas `{low_a_last}` was restricted to the lowest tier ({low_val_a_last:.2f}). Simultaneously, `{factor_b_col}` demonstrated a "
        f"{sig_b_first_text} response early on, progressing to a {sig_b_last_text} effect by the final day of evaluation. At {last_day_str}, the `{top_b_last}` "
        f"treatment was superior ({top_val_b_last:.2f}^{top_let_b_last}) compared to `{low_b_last}` ({low_val_b_last:.2f}). {interaction_evolution}"
    )
    return para

def generate_single_explanation(param_name, p_data, factor_a_col, factor_b_col, table_label):
    """
    Generates academic explanations for single-interval variables (e.g., Shelf Life, pH, Firmness)
    matching all potential significance profiles.
    """
    p_a, p_b, p_ab = p_data["p_a"], p_data["p_b"], p_data["p_ab"]
    sig_a_text = "significant" if p_a < 0.05 else "nonsignificant"
    p_not_a = f"(p < 0.05)" if p_a < 0.05 else "(p > 0.05)"
    if p_a < 0.01: p_not_a = "(p < 0.01)"
    
    sig_b_text = "significant" if p_b < 0.05 else "nonsignificant"
    p_not_b = f"(p < 0.05)" if p_b < 0.05 else "(p > 0.05)"
    if p_b < 0.01: p_not_b = "(p < 0.01)"
    
    sig_ab_text = "significant" if p_ab < 0.05 else "nonsignificant"
    p_not_ab = f"(p < 0.05)" if p_ab < 0.05 else "(p > 0.05)"
    if p_ab < 0.01: p_not_ab = "(p < 0.01)"
    
    sorted_a = sorted(p_data["means_a"].items(), key=lambda x: x[1], reverse=True)
    top_a, top_val_a = sorted_a[0]
    low_a, low_val_a = sorted_a[-1]
    top_let_a = p_data["means_a_str"][top_a].replace(f"{top_val_a:.2f}", "")
    
    sorted_b = sorted(p_data["means_b"].items(), key=lambda x: x[1], reverse=True)
    top_b, top_val_b = sorted_b[0]
    low_b, low_val_b = sorted_b[-1]
    top_let_b = p_data["means_b_str"][top_b].replace(f"{top_val_b:.2f}", "")
    
    sorted_comb = sorted(p_data["means_comb"].items(), key=lambda x: x[1], reverse=True)
    comb_top_name, comb_top_val = sorted_comb[0]
    comb_low_name, comb_low_val = sorted_comb[-1]
    comb_top_let = p_data["cld_comb"].get(comb_top_name, "")
    
    # Build at-par lists
    at_par_a_list = []
    for lvl, val in sorted_a[1:]:
        let = p_data["means_a_str"][lvl].replace(f"{val:.2f}", "")
        if top_let_a and let and any(char in top_let_a for char in let):
            at_par_a_list.append(f"{lvl} ({val:.2f}^{let})")
    at_par_a_str = f"statistically at par with {', '.join(at_par_a_list)}" if at_par_a_list else "distinctly superior to all other levels"
    
    at_par_b_list = []
    for lvl, val in sorted_b[1:]:
        let = p_data["means_b_str"][lvl].replace(f"{val:.2f}", "")
        if top_let_b and let and any(char in top_let_b for char in let):
            at_par_b_list.append(f"{lvl} ({val:.2f}^{let})")
    at_par_b_str = f"statistically comparable to {', '.join(at_par_b_list)}" if at_par_b_list else "uniquely superior to alternative levels"
    
    if p_ab < 0.05:
        para = (
            f"Regarding the parameter **{param_name}**, the results summarized in **{table_label}** revealed a highly **significant** "
            f"interaction effect between `{factor_a_col}` and `{factor_b_col}` {p_not_ab}. This combined response confirms that "
            f"the regulatory behavior of `{factor_b_col}` depends heavily on the specific application level of `{factor_a_col}`. "
            f"Among all treatment combinations, `{comb_top_name}` yielded the maximum value of {comb_top_val:.2f}^{comb_top_let}, "
            f"which was significantly greater than the lowest-performing combination `{comb_low_name}` ({comb_low_val:.2f}). "
            f"The main effect of `{factor_a_col}` was {sig_a_text} {p_not_a} and `{factor_b_col}` was {sig_b_text} {p_not_b}, but these "
            f"individual responses are best interpreted through their interactive behavior."
        )
    elif p_a > 0.05 and p_b > 0.05:
        para = (
            f"The analysis of variance for **{param_name}** indicated that both `{factor_a_col}` {p_not_a} and `{factor_b_col}` {p_not_b} "
            f"had **nonsignificant** main effects on this trait, as shown in **{table_label}**. The observed treatment means fluctuated "
            f"within a very narrow margin around the grand mean of {p_data['gm']:.2f}, representing a highly stable and uniform trait expression. "
            f"Furthermore, the interaction effect (`{factor_a_col}` × `{factor_b_col}`) was strictly **nonsignificant** {p_not_ab}, "
            f"reflecting mutual independence and lack of combined phenotypic stratification across the trial treatments."
        )
    else:
        part_a_text = ""
        if p_a < 0.05:
            part_a_text = (
                f"the main effect of `{factor_a_col}` was **significant** {p_not_a}, with treatment `{top_a}` "
                f"recording the highest performance of {top_val_a:.2f}^{top_let_a}, which was {at_par_a_str}, while `{low_a}` "
                f"({low_val_a:.2f}) produced the lowest performance"
            )
        else:
            part_a_text = f"the main effect of `{factor_a_col}` was **nonsignificant** {p_not_a}, showing uniform behavior across levels"
            
        part_b_text = ""
        if p_b < 0.05:
            part_b_text = (
                f"the treatment factor `{factor_b_col}` exerted a highly **significant** {p_not_b} response, "
                f"wherein `{top_b}` led with {top_val_b:.2f}^{top_let_b} ({at_par_b_str}), whereas `{low_b}` "
                f"({low_val_b:.2f}) marked the minimum baseline limit"
            )
        else:
            part_b_text = f"the main effect of `{factor_b_col}` was **nonsignificant** {p_not_b}, with all levels remaining statistically at par"
            
        para = (
            f"For the parameter **{param_name}**, the interaction effect (`{factor_a_col}` × `{factor_b_col}`) was completely "
            f"**nonsignificant** {p_not_ab}, confirming that the treatment factors operated independently. "
            f"Specifically, as detailed in **{table_label}**, {part_a_text}. Simultaneously, {part_b_text} around the grand mean of {p_data['gm']:.2f}."
        )
    return para

# --- Formatting Helpers for Excel and Word Output ---
def get_signif_code_python(p):
    if pd.isna(p): return "ns"
    if p < 0.01: return "**"
    elif p < 0.05: return "*"
    else: return "ns"

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

def parse_dmrt_value(val):
    if pd.isna(val):
        return "", ""
    val_str = str(val).strip()
    match = re.match(r"^([\d\.\-]+)\s*([a-z]+)?$", val_str)
    if match:
        num = match.group(1)
        letters = match.group(2) if match.group(2) else ""
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

# --- Styled Excel Exporter ---
def build_styled_excel(factor_a_col, factor_b_col, params, levels_a, levels_b, results_dict):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "2-Factor RCBD Output"
    
    font_arial_bold = Font(name="Arial", size=10, bold=True)
    font_arial_regular = Font(name="Arial", size=10)
    align_left = Alignment(horizontal="left", vertical="center")
    align_center = Alignment(horizontal="center", vertical="center")
    
    border_thin_bottom = Border(bottom=Side(style='thin', color='000000'))
    border_medium_bottom = Border(bottom=Side(style='medium', color='000000'))
    
    ws.cell(row=1, column=1, value="Treatments").font = font_arial_bold
    ws.cell(row=1, column=1).alignment = align_left
    
    for col_idx, param in enumerate(params, start=2):
        cell = ws.cell(row=1, column=col_idx, value=param)
        cell.font = font_arial_bold
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
    
    ws.cell(row=row_factor_a_title, column=1, value=f"Factor A: {factor_a_col}").font = font_arial_bold
    ws.cell(row=row_factor_a_title, column=1).alignment = align_left
    
    for idx, lvl in enumerate(levels_a):
        r = row_factor_a_levels_start + idx
        ws.cell(row=r, column=1, value=lvl).font = font_arial_regular
        ws.cell(row=r, column=1).alignment = align_left
        
    ws.cell(row=row_sem_a, column=1, value="SEm(±)").font = font_arial_regular
    ws.cell(row=row_sem_a, column=1).alignment = align_left
    ws.cell(row=row_f_a, column=1, value="F-value").font = font_arial_regular
    ws.cell(row=row_f_a, column=1).alignment = align_left
    ws.cell(row=row_lsd_a, column=1, value="LSD(0.05)").font = font_arial_regular
    ws.cell(row=row_lsd_a, column=1).alignment = align_left
    
    ws.cell(row=row_factor_b_title, column=1, value=f"Factor B: {factor_b_col}").font = font_arial_bold
    ws.cell(row=row_factor_b_title, column=1).alignment = align_left
    
    for idx, lvl in enumerate(levels_b):
        r = row_factor_b_levels_start + idx
        ws.cell(row=r, column=1, value=lvl).font = font_arial_regular
        ws.cell(row=r, column=1).alignment = align_left
        
    ws.cell(row=row_sem_b, column=1, value="SEm(±)").font = font_arial_regular
    ws.cell(row=row_sem_b, column=1).alignment = align_left
    ws.cell(row=row_f_b, column=1, value="F-value").font = font_arial_regular
    ws.cell(row=row_f_b, column=1).alignment = align_left
    ws.cell(row=row_lsd_b, column=1, value="LSD(0.05)").font = font_arial_regular
    ws.cell(row=row_lsd_b, column=1).alignment = align_left
    
    ws.cell(row=row_cv, column=1, value="CV, %").font = font_arial_regular
    ws.cell(row=row_cv, column=1).alignment = align_left
    ws.cell(row=row_inter, column=1, value="Factor A × Factor B").font = font_arial_regular
    ws.cell(row=row_inter, column=1).alignment = align_left
    ws.cell(row=row_gm, column=1, value="Grand mean").font = font_arial_regular
    ws.cell(row=row_gm, column=1).alignment = align_left
    
    for col_idx, param in enumerate(params, start=2):
        p_data = results_dict[param]
        
        for idx, lvl in enumerate(levels_a):
            r = row_factor_a_levels_start + idx
            cell = ws.cell(row=r, column=col_idx, value=p_data["means_a_str"][lvl])
            cell.font = font_arial_regular
            cell.alignment = align_center
            
        ws.cell(row=row_sem_a, column=col_idx, value=p_data["sem_a"]).font = font_arial_regular
        ws.cell(row=row_sem_a, column=col_idx).alignment = align_center
        ws.cell(row=row_f_a, column=col_idx, value=p_data["sig_a"]).font = font_arial_regular
        ws.cell(row=row_f_a, column=col_idx).alignment = align_center
        ws.cell(row=row_lsd_a, column=col_idx, value=p_data["lsd_a"]).font = font_arial_regular
        ws.cell(row=row_lsd_a, column=col_idx).alignment = align_center
        
        for idx, lvl in enumerate(levels_b):
            r = row_factor_b_levels_start + idx
            cell = ws.cell(row=r, column=col_idx, value=p_data["means_b_str"][lvl])
            cell.font = font_arial_regular
            cell.alignment = align_center
            
        ws.cell(row=row_sem_b, column=col_idx, value=p_data["sem_b"]).font = font_arial_regular
        ws.cell(row=row_sem_b, column=col_idx).alignment = align_center
        ws.cell(row=row_f_b, column=col_idx, value=p_data["sig_b"]).font = font_arial_regular
        ws.cell(row=row_f_b, column=col_idx).alignment = align_center
        ws.cell(row=row_lsd_b, column=col_idx, value=p_data["lsd_b"]).font = font_arial_regular
        ws.cell(row=row_lsd_b, column=col_idx).alignment = align_center
        
        ws.cell(row=row_cv, column=col_idx, value=p_data["cv"]).font = font_arial_regular
        ws.cell(row=row_cv, column=col_idx).alignment = align_center
        ws.cell(row=row_inter, column=col_idx, value=p_data["sig_ab"]).font = font_arial_regular
        ws.cell(row=row_inter, column=col_idx).alignment = align_center
        ws.cell(row=row_gm, column=col_idx, value=p_data["gm"]).font = font_arial_regular
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
        p.runs[0].font.italic = True
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
        
    # Factor A Block
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
    
    # Factor B Block
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
    
    # Statistical Metrics Block
    cv_dict = {p: results_data[p]["cv"] for p in g_cols}
    sig_ab_dict = {p: results_data[p]["sig_ab"] for p in g_cols}
    gm_dict = {p: results_data[p]["gm"] for p in g_cols}
    
    r_cv = add_styled_data_row("CV, %", cv_dict)
    set_header_bottom_border(r_cv)
    
    r_inter = add_styled_data_row("Factor A × Factor B", sig_ab_dict)
    set_header_bottom_border(r_inter)
    
    r_gm = add_styled_data_row("Grand mean", gm_dict)
    set_header_bottom_border(r_gm)

# --- Summarized Table Parser Engine ---
def parse_summarized_table_to_results(df_raw, idx_A, idx_B, idx_cv, idx_interaction, idx_grand,
                                      idx_A_sem, idx_A_f, idx_A_lsd, idx_B_sem, idx_B_f, idx_B_lsd,
                                      factor_a_levels, factor_b_levels, parameters):
    results_data = {}
    for param in parameters:
        col_idx = df_raw.iloc[0].tolist().index(param)
        
        f_val_A = str(df_raw.iloc[idx_A_f, col_idx]).strip().lower()
        f_val_B = str(df_raw.iloc[idx_B_f, col_idx]).strip().lower()
        f_val_AB = str(df_raw.iloc[idx_interaction, col_idx]).strip().lower()
        
        # P-values mapping from significance stars
        p_a = 0.01 if "**" in f_val_A else (0.04 if "*" in f_val_A else 0.5)
        p_b = 0.01 if "**" in f_val_B else (0.04 if "*" in f_val_B else 0.5)
        p_ab = 0.01 if "**" in f_val_AB else (0.04 if "*" in f_val_AB else 0.5)
        
        # Parse Factor A Means
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
                means_a_str[lvl] = f"0.00"
                
        # Parse Factor B Means
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
                means_b_str[lvl] = f"0.00"
                
        gm_val = 0.0
        try: gm_val = float(df_raw.iloc[idx_grand, col_idx])
        except ValueError: pass
        
        # Placeholder structures since interaction details aren't stored element-wise in summarized tables
        means_comb = {f"{a} × {b}": (means_a[a] + means_b[b])/2 for a in factor_a_levels for b in factor_b_levels}
        cld_comb = {k: "" for k in means_comb}
        
        results_data[param] = {
            "means_a": means_a, "means_a_str": means_a_str, "sem_a": df_raw.iloc[idx_A_sem, col_idx], 
            "sig_a": df_val_or_blank(df_raw.iloc[idx_A_f, col_idx]), "lsd_a": df_raw.iloc[idx_A_lsd, col_idx], "p_a": p_a,
            "means_b": means_b, "means_b_str": means_b_str, "sem_b": df_raw.iloc[idx_B_sem, col_idx], 
            "sig_b": df_val_or_blank(df_raw.iloc[idx_B_f, col_idx]), "lsd_b": df_raw.iloc[idx_B_lsd, col_idx], "p_b": p_b,
            "means_comb": means_comb, "cld_comb": cld_comb, "p_ab": p_ab, "sig_ab": df_val_or_blank(df_raw.iloc[idx_interaction, col_idx]),
            "cv": df_raw.iloc[idx_cv, col_idx], "gm": gm_val
        }
    return results_data

def df_val_or_blank(val):
    return "" if pd.isna(val) else str(val).strip()

# --- Main App Logic ---
def show_module():
    st.markdown("## RCBD Two-Factor Factorial Analytical Engine")
    mode = st.radio("Choose Input Mode", ["Raw Data Mode", "Summarized Table Mode"], key="2f_mode")
    uploaded_file = st.file_uploader("Upload Two-Factor Excel File", type=["xlsx"], key="file_uploader_2f")

    if uploaded_file is not None:
        if mode == "Raw Data Mode":
            run_raw_mode(uploaded_file)
        else:
            run_summary_mode(uploaded_file)

# --- Raw Data Analyzer Mode ---
def run_raw_mode(uploaded_file):
    try:
        df_raw_data = pd.read_excel(uploaded_file)
        st.write("#### Raw Data Preview:", df_raw_data.head())
        
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
            
            group_1_name = st.text_input("First Table Title", f"Physiological and Physical Properties", key="raw_2f_g1_name_m")
            group_1_cols = st.multiselect("Select parameters for Table 1", response_cols, default=response_cols[:len(response_cols)//2], key="raw_2f_g1_cols_m")
            group_2_name = st.text_input("Second Table Title", f"Quality and Biochemical Properties", key="raw_2f_g2_name_m")
            group_2_cols = st.multiselect("Select parameters for Table 2", [c for c in response_cols if c not in group_1_cols], default=[c for c in response_cols if c not in group_1_cols], key="raw_2f_g2_cols_m")
            
            if st.button("Execute Two-Factor Analysis", key="btn_raw_2f_calc_m"):
                results_data = {}
                
                levels_a = sorted(df_raw_data[factor_a_col].unique().tolist())
                levels_b = sorted(df_raw_data[factor_b_col].unique().tolist())
                
                r = df_raw_data[block_col].nunique()
                a_levels = len(levels_a)
                b_levels = len(levels_b)
                
                for param in response_cols:
                    df_temp = pd.DataFrame({
                        'rep': df_raw_data[block_col].astype(str),
                        'factor_a': df_raw_data[factor_a_col].astype(str),
                        'factor_b': df_raw_data[factor_b_col].astype(str),
                        'response': pd.to_numeric(df_raw_data[param], errors='coerce')
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
                    
                    # Factor A Computations
                    means_a = df_temp.groupby('factor_a')['response'].mean().to_dict()
                    sem_a = np.sqrt(mse / (r * b_levels))
                    lsd_a = t_val * np.sqrt((2 * mse) / (r * b_levels))
                    cld_a = get_cld_letters(means_a, lsd_a)
                    
                    # Factor B Computations
                    means_b = df_temp.groupby('factor_b')['response'].mean().to_dict()
                    sem_b = np.sqrt(mse / (r * a_levels))
                    lsd_b = t_val * np.sqrt((2 * mse) / (r * a_levels))
                    cld_b = get_cld_letters(means_b, lsd_b)
                    
                    # Interaction Combinations
                    df_temp['Combination'] = df_temp['factor_a'] + " × " + df_temp['factor_b']
                    means_comb = df_temp.groupby('Combination')['response'].mean().to_dict()
                    lsd_comb = t_val * np.sqrt((2 * mse) / r)
                    cld_comb = get_cld_letters(means_comb, lsd_comb)
                    
                    sig_a = get_signif_code_python(p_a)
                    sig_b = get_signif_code_python(p_b)
                    sig_ab = get_signif_code_python(p_ab)
                    
                    means_a_str = {}
                    for lvl, val in means_a.items():
                        letter = cld_a[lvl] if sig_a != "ns" else ""
                        means_a_str[lvl] = f"{val:.2f}{letter}"
                        
                    means_b_str = {}
                    for lvl, val in means_b.items():
                        letter = cld_b[lvl] if sig_b != "ns" else ""
                        means_b_str[lvl] = f"{val:.2f}{letter}"
                        
                    results_data[param] = {
                        "means_a": means_a, "means_a_str": means_a_str, "sem_a": round(sem_a, 2), "sig_a": sig_a, "lsd_a": round(lsd_a, 2), "p_a": p_a,
                        "means_b": means_b, "means_b_str": means_b_str, "sem_b": round(sem_b, 2), "sig_b": sig_b, "lsd_b": round(lsd_b, 2), "p_b": p_b,
                        "means_comb": means_comb, "cld_comb": cld_comb, "p_ab": p_ab, "sig_ab": sig_ab,
                        "cv": round(cv, 2), "gm": round(grand_mean, 2)
                    }
                
                # Create Styled Excel Workbook
                styled_wb = build_styled_excel(factor_a_col, factor_b_col, response_cols, levels_a, levels_b, results_data)
                excel_bio = io.BytesIO()
                styled_wb.save(excel_bio)
                excel_bio.seek(0)
                
                st.markdown("#### 📥 Download Formatted Statistical Excel Results")
                st.download_button(
                    label="Download Excel Results Sheet",
                    data=excel_bio,
                    file_name="Result_2Factor_Output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="btn_d_excel_styled"
                )
                st.write("---")
                
                # Generate explanations and structured docx layout
                doc = Document()
                doc.add_heading("Two-Factor Factorial RCBD Comprehensive Report", 0)
                
                groups_layouts = [
                    (group_1_name, group_1_cols, 1),
                    (group_2_name, group_2_cols, 2)
                ]
                
                st.markdown("### 📝 Dynamic Analysis Results & Discussions")
                
                for g_title, g_cols, table_num in groups_layouts:
                    if not g_cols: continue
                    table_label = f"Table {table_num}"
                    
                    st.write(f"### {table_label}: {g_title}")
                    doc.add_heading(f"{table_label}: {g_title}", level=2)
                    
                    # Group items inside this table to identify trend lines
                    grouped = group_parameters(g_cols)
                    
                    # Print explanations first
                    for base_name, items in sorted(grouped.items()):
                        st.write(f"#### {base_name}")
                        doc.add_heading(base_name, level=3)
                        
                        if len(items) > 1:
                            # Multi-day/Time-series parameter (Trend-line)
                            p_text = generate_trend_explanation(base_name, items, results_data, factor_a_col, factor_b_col, table_label)
                        else:
                            # Isolated singular parameter
                            p_text = generate_single_explanation(items[0][0], results_data[items[0][0]], factor_a_col, factor_b_col, table_label)
                        
                        st.write(p_text)
                        
                        p_docx = doc.add_paragraph()
                        parts = re.split(r'(\*\*.*?\*\*)', p_text)
                        for part in parts:
                            if part.startswith('**') and part.endswith('**'):
                                p_docx.add_run(part[2:-2]).bold = True
                            else:
                                p_docx.add_run(part)
                                
                    # Place formatted table directly under explanations
                    st.write("##### Corresponding Table Visualization:")
                    add_excel_table_to_docx(doc, factor_a_col, factor_b_col, g_cols, levels_a, levels_b, results_data)
                    st.write("*(Table data formatted as per standard Q1 guidelines)*")
                    doc.add_page_break()
                    
                bio_doc = io.BytesIO()
                doc.save(bio_doc)
                bio_doc.seek(0)
                
                st.write("---")
                st.markdown("#### 💾 Save Explanations & Tables as Word Document")
                st.download_button(
                    "Download Word Explanations Report (.docx)", 
                    data=bio_doc, 
                    file_name="TwoFactorial_Thesis_Report.docx", 
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", 
                    key="btn_d_2f_raw_cal"
                )
    except Exception as e:
        st.error(f"Error executing raw combined Two-Factor analysis: {e}")

# --- Summarized Table Mode ---
def run_summary_mode(uploaded_file):
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
            
        # Parse statistics for Factor A
        idx_A_sem, idx_A_f, idx_A_lsd = None, None, None
        for idx in range(idx_A + 1, idx_B):
            val = str(df_raw.iloc[idx, 0]).strip().lower()
            if "sem" in val: idx_A_sem = idx
            elif "f-value" in val or "f value" in val: idx_A_f = idx
            elif "lsd" in val: idx_A_lsd = idx
            
        # Parse statistics for Factor B
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
        
        group_1_name = st.text_input("First Table Title", "Physiological and Physical Attributes", key="2f_g1_name")
        group_1_cols = st.multiselect("Select parameters for Table 1", parameters, default=parameters[:len(parameters)//2], key="2f_g1_cols")
        group_2_name = st.text_input("Second Table Title", "Quality and Chemical Attributes", key="2f_g2_name")
        group_2_cols = st.multiselect("Select parameters for Table 2", [c for c in parameters if c not in group_1_cols], default=[c for c in parameters if c not in group_1_cols], key="2f_g2_cols")
        
        if st.button("Generate Word Document Draft from Summary", key="btn_2f_summary"):
            results_data = parse_summarized_table_to_results(
                df_raw, idx_A, idx_B, idx_cv, idx_interaction, idx_grand,
                idx_A_sem, idx_A_f, idx_A_lsd, idx_B_sem, idx_B_f, idx_B_lsd,
                factor_a_levels, factor_b_levels, parameters
            )
            
            doc = Document()
            doc.add_heading("Two-Factor Factorial RCBD Comprehensive Report", 0)
            
            groups_layouts = [
                (group_1_name, group_1_cols, 1),
                (group_2_name, group_2_cols, 2)
            ]
            
            st.markdown("### 📝 Analysis Results and Academic Explanations")
            
            for g_title, g_cols, table_num in groups_layouts:
                if not g_cols: continue
                table_label = f"Table {table_num}"
                
                st.write(f"### {table_label}: {g_title}")
                doc.add_heading(f"{table_label}: {g_title}", level=2)
                
                grouped = group_parameters(g_cols)
                
                for base_name, items in sorted(grouped.items()):
                    st.write(f"#### {base_name}")
                    doc.add_heading(base_name, level=3)
                    
                    if len(items) > 1:
                        p_text = generate_trend_explanation(base_name, items, results_data, factor_a_label, factor_b_label, table_label)
                    else:
                        p_text = generate_single_explanation(items[0][0], results_data[items[0][0]], factor_a_label, factor_b_label, table_label)
                    
                    st.write(p_text)
                    
                    p_docx = doc.add_paragraph()
                    parts = re.split(r'(\*\*.*?\*\*)', p_text)
                    for part in parts:
                        if part.startswith('**') and part.endswith('**'):
                            p_docx.add_run(part[2:-2]).bold = True
                        else:
                            p_docx.add_run(part)
                            
                st.write("##### Corresponding Table Visualization:")
                add_excel_table_to_docx(doc, factor_a_label, factor_b_label, g_cols, factor_a_levels, factor_b_levels, results_data)
                st.write("*(Table formatted as per standard publication design)*")
                doc.add_page_break()
                
            bio = io.BytesIO()
            doc.save(bio)
            bio.seek(0)
            st.success("🎉 Word Report generated from Summarized Two-Factor data!")
            st.download_button("Download Report (.docx)", data=bio, file_name="Summarized_Factorial_Report.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="btn_d_2f_sum_mod")
    except Exception as e:
        st.error(f"Error parsing Summarized Factorial Table: {e}")

# --- Entrypoint ---
if __name__ == '__main__':
    show_module()
