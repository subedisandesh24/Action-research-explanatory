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

# --- 20 Academic Thesis-Style Factorial Templates ---
TEMPLATES_MAIN_EFFECTS = [
    "The effect of `{factor_a}` and `{factor_b}` on **{parameter}** revealed that the main effect of `{factor_a}` was `{sig_a}` `{p_a}`. Specifically, `{top_a}` recorded a mean of `{top_val_a}^{top_let_a}`, which was significantly superior to `{low_a}` (`{low_val_a}^{low_let_a}`) and statistically at par with `{at_par_a}`. Regarding `{factor_b}`, a `{sig_b}` `{p_b}` response was observed, with `{top_b}` (`{top_val_b}^{top_let_b}`) achieving the highest performance while `{at_par_b}` remained statistically equivalent around the grand mean of `{grand_mean}`. The interaction between `{factor_a}` and `{factor_b}` was **nonsignificant** `{p_interaction}`, showing that the two treatment factors influenced this parameter independently.",
    "An analysis of the **{parameter}** data showed that the treatment factor `{factor_a}` had a `{sig_a}` `{p_a}` main effect. Genotype/treatment `{top_a}` attained the highest performance of `{top_val_a}^{top_let_a}`, which was statistically at par with `{at_par_a}`, whereas `{low_a}` (`{low_val_a}^{low_let_a}`) recorded the minimum value. For `{factor_b}`, the main effect was `{sig_b}` `{p_b}`, with `{top_b}` (`{top_val_b}^{top_let_b}`) outperforming the other levels, while `{at_par_b}` remained statistically comparable. No significant interaction effect (`{factor_a}` × `{factor_b}`) was registered `{p_interaction}`, confirming independent regulatory behavior.",
    "For **{parameter}**, the main effect of `{factor_a}` was `{sig_a}` `{p_a}`. The maximum value was observed under `{top_a}` (`{top_val_a}^{top_let_a}`), showing statistical parity with `{at_par_a}`. The lowest values were restricted to `{low_a}` (`{low_val_a}^{low_let_a}`). Simultaneously, `{factor_b}` exerted a `{sig_b}` `{p_b}` main effect, where `{top_b}` (`{top_val_b}^{top_let_b}`) registered the highest mean value, and `{at_par_b}` co-occupied a similar statistical tier. The interaction effect between the two factors was **nonsignificant** `{p_interaction}`, indicating a lack of combined dependency.",
    "Regarding **{parameter}**, the treatment levels of `{factor_a}` induced a `{sig_a}` `{p_a}` main effect. Treatment `{top_a}` achieved the maximum value of `{top_val_a}^{top_let_a}` which was statistically at par with `{at_par_a}`. Conversely, `{low_a}` (`{low_val_a}^{low_let_a}`) resulted in the lowest value. For `{factor_b}`, a `{sig_b}` `{p_b}` main effect was observed where `{top_b}` (`{top_val_b}^{top_let_b}`) was superior, and `{at_par_b}` remained statistically similar. Crucially, the `{factor_a}` × `{factor_b}` interaction was **nonsignificant** `{p_interaction}`, showing that the performance of `{factor_a}` does not vary across different levels of `{factor_b}`.",
    "The parameter **{parameter}** was significantly influenced by the main effect of `{factor_a}` `{p_a}`. The highest performance was recorded under `{top_a}` (`{top_val_a}^{top_let_a}`), which was statistically at par with `{at_par_a}`, whereas `{low_a}` (`{low_val_a}^{low_let_a}`) produced the lowest values. The main effect of `{factor_b}` was `{sig_b}` `{p_b}`, showing that `{top_b}` (`{top_val_b}^{top_let_b}`) led the application rates while `{at_par_b}` behaved statistically at par relative to the grand mean of `{grand_mean}`. The interaction effect between `{factor_a}` and `{factor_b}` was verified as **nonsignificant** `{p_interaction}`.",
    "Statistical evaluation of **{parameter}** indicated that the main effect of `{factor_a}` was `{sig_a}` `{p_a}`. Treatment `{top_a}` recorded a mean of `{top_val_a}^{top_let_a}`, representing a significantly higher value compared to `{low_a}` (`{low_val_a}^{low_let_a}`), and remained at par with `{at_par_a}`. The factor `{factor_b}` showed a `{sig_b}` `{p_b}` main effect, wherein `{top_b}` (`{top_val_b}^{top_let_b}`) achieved maximum efficiency, and `{at_par_b}` performed statistically equivalent. Furthermore, a **nonsignificant** `{p_interaction}` interaction was confirmed, pointing to mutually independent actions of the factors.",
    "Concerning **{parameter}**, the main effect of `{factor_a}` was `{sig_a}` `{p_a}`. The maximum value of `{top_val_a}^{top_let_a}` was registered by `{top_a}`, which was statistically at par with `{at_par_a}` and significantly superior to `{low_a}` (`{low_val_a}^{low_let_a}`). For `{factor_b}`, a `{sig_b}` `{p_b}` main effect occurred, where `{top_b}` (`{top_val_b}^{top_let_b}`) achieved the highest mean, while `{at_par_b}` shared statistical parity. The interaction between `{factor_a}` and `{factor_b}` remained **nonsignificant** `{p_interaction}`, reflecting independent performance patterns under the trial conditions.",
    "The analysis of variance for **{parameter}** revealed that the main effect of `{factor_a}` was `{sig_a}` `{p_a}`. The highest tier was occupied by `{top_a}` (`{top_val_a}^{top_let_a}`), which shared statistical letters with `{at_par_a}`. Meanwhile, the main effect of `{factor_b}` was `{sig_b}` `{p_b}`, with `{top_b}` (`{top_val_b}^{top_let_b}`) recording the maximum values and `{at_par_b}` showing statistical equivalence. No significant interaction effect (`{factor_a}` × `{factor_b}`) was observed `{p_interaction}`, which denotes that the response of `{factor_a}` was consistent across all levels of `{factor_b}`.",
    "For **{parameter}**, the main effect of `{factor_a}` was `{sig_a}` `{p_a}`. Genotype `{top_a}` stood out as the leading treatment with a mean of `{top_val_a}^{top_let_a}`, remaining statistically at par with `{at_par_a}`. In comparison, the main effect of `{factor_b}` was `{sig_b}` `{p_b}`, which meant `{top_b}` (`{top_val_b}^{top_let_b}`) led the treatments while `{at_par_b}` remained statistically similar. Finally, the additive action of `{factor_a}` and `{factor_b}` was confirmed by a **nonsignificant** `{p_interaction}` interaction effect.",
    "The parameter **{parameter}** was significantly affected by the main effect of `{factor_a}` `{p_a}`. The highest mean was registered by `{top_a}` (`{top_val_a}^{top_let_a}`), which was statistically at par with `{at_par_a}`, whereas `{low_a}` (`{low_val_a}^{low_let_a}`) marked the lower limit. Conversely, `{factor_b}` exerted a `{sig_b}` `{p_b}` main effect, meaning `{top_b}` (`{top_val_b}^{top_let_b}`) and `{at_par_b}` remained statistically equivalent around the grand mean of `{grand_mean}`. The statistical interaction between `{factor_a}` and `{factor_b}` was **nonsignificant** `{p_interaction}`."
]

TEMPLATES_BOTH_NS = [
    "For the parameter **{parameter}**, the main effect of `{factor_a}` was **nonsignificant** `(p > 0.05)`, indicating that all levels performed statistically equivalent. Similarly, the main effect of `{factor_b}` was **nonsignificant** `(p > 0.05)`, leaving all treatment levels statistically at par around the grand mean of `{grand_mean}`. Finally, the interaction effect (`{factor_a}` × `{factor_b}`) was **nonsignificant** `(p > 0.05)`, demonstrating a complete lack of phenotypic stratification across the entire trial.",
    "Regarding **{parameter}**, statistical analysis showed that `{factor_a}` had a **nonsignificant** `(p > 0.05)` main effect, with `{top_a}` (`{top_val_a}`) and `{low_a}` (`{low_val_a}`) remaining statistically equivalent. This flat trend was mirrored in `{factor_b}`, where the main effect was also **nonsignificant** `(p > 0.05)`, leaving all levels statistically at par. In conclusion, the interaction effect between `{factor_a}` and `{factor_b}` was **nonsignificant** `(p > 0.05)`, reflecting highly stable and uniform trait expression.",
    "The main effect of `{factor_a}` on **{parameter}** was **nonsignificant** `(p > 0.05)`, indicating that all treatment levels possessed equivalent performance potential. The main effect of `{factor_b}` was similarly **nonsignificant** `(p > 0.05)`, leaving `{top_b}` (`{top_val_b}`) and `{low_b}` (`{low_val_b}`) statistically at par. Ultimately, the interaction of `{factor_a}` × `{factor_b}` was verified as **nonsignificant** `(p > 0.05)`, confirming the mutual independence of both factors under the experimental conditions.",
    "Analysis of variance for **{parameter}** indicated that neither the individual factor `{factor_a}` nor `{factor_b}` had a significant main effect (p > 0.05). The mean values remained statistically uniform across all treatments, with a grand mean of `{grand_mean}`. The interaction effect (`{factor_a}` × `{factor_b}`) was also confirmed as **nonsignificant** `(p > 0.05)`, pointing to a completely homogeneous response pattern.",
    "The main effect of `{factor_a}` on **{parameter}** was **nonsignificant** `(p > 0.05)`, leaving `{at_par_b}` statistically equivalent. For `{factor_b}`, the main effect was also **nonsignificant** `(p > 0.05)`, meaning all treatment levels performed statistically at par. At last, the interaction effect between `{factor_a}` and `{factor_b}` was **nonsignificant** `(p > 0.05)`, indicating a lack of combined dependency and showing that the trait was highly stable."
]

TEMPLATES_SIGNIFICANT_INTERACTION = [
    "For **{parameter}**, the main effect of `{factor_a}` was `{sig_a}` `{p_a}` and `{factor_b}` was `{sig_b}` `{p_b}`. However, the interaction effect between `{factor_a}` and `{factor_b}` was highly **significant** `{p_interaction}`, confirming that the factors do not act independently. Among all treatment combinations, `{combination_top}` registered the highest mean value, remaining statistically at par with `{at_par_comb}`. Conversely, `{combination_low}` marked the lowest limit of performance, representing the severity of stress or low-performance conditions.",
    "The analysis of variance for **{parameter}** revealed a highly **significant** `{p_interaction}` interaction between `{factor_a}` and `{factor_b}`, meaning the regulatory effect of `{factor_b}` was differently expressed depending on the level of `{factor_a}`. Upon evaluating the treatment combinations, `{combination_top}` stood at the statistical apex, showing statistical parity with `{at_par_comb}`. In contrast, `{combination_low}` was significantly inferior, representing the lowest performance tier among all combinations.",
    "Regarding **{parameter}**, while both `{factor_a}` `{p_a}` and `{factor_b}` `{p_b}` main effects were significant, they were mediated by a highly **significant** `{p_interaction}` interaction. This interaction became statistically significant as the combinations responded differentially. Among all twelve treatment combinations, `{combination_top}` achieved superior performance, remaining statistically at par with `{at_par_comb}`. Conversely, `{combination_low}` showed the lowest value, representing the most severely affected or lowest-performing combination.",
    "The parameter **{parameter}** showed a highly **significant** interaction `{p_interaction}` between `{factor_a}` and `{factor_b}`, highlighting that the combination of these factors is crucial for trait manifestation. Analysis of treatment combinations showed that `{combination_top}` resulted in the highest value, which was statistically at par with `{at_par_comb}`. The most severely restricted combination was `{combination_low}`, which recorded the lowest mean and was statistically inferior to most other combinations.",
    "A highly **significant** `{p_interaction}` interaction was observed between `{factor_a}` and `{factor_b}` for **{parameter}**. Evaluating the treatment combinations revealed that `{combination_top}` was the optimal combination, remaining statistically at par with `{at_par_comb}`. This combination was significantly superior to the other groups, whereas `{combination_low}` represented the absolute minimum value recorded, indicating the severity of the untreated or control conditions."
]

# --- Helper Functions for Style and Parsing ---
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

# --- Parameter Grouping Engine for Time-Series Analysis ---
def group_parameters(params):
    pattern = re.compile(r"\b(?:at\s+)?(\d+)\s*(dat|das|days|day|d)\b", re.IGNORECASE)
    groups = {}
    
    for p in params:
        match = pattern.search(p)
        if match:
            day_num = int(match.group(1))
            day_str = match.group(0)
            
            # Extract parent category name and strip hanging punctuation
            base = pattern.sub("", p)
            base = re.sub(r"\s*[\(\)\-\,\s]+\s*at\s*$", "", base, flags=re.IGNORECASE)
            base = re.sub(r"\s*[\(\)\-\,\s]+\s*$", "", base)
            base = base.strip().capitalize()
            
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

# --- DOCX Copy of the exact styled Excel table ---
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
        hdr_cells[c_idx + 1].paragraphs[0].alignment = 1  # Center alignment
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
            p_val.alignment = 1  # Center alignment
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

# --- Main Streamlit Function ---
def show_module():
    st.markdown("### RCBD Two-Factor Factorial Analyzer")
    mode = st.radio("Choose Input Mode", ["Summarized Table Mode", "Raw Data Mode"], key="2f_mode")
    uploaded_file = st.file_uploader("Upload Two-Factor Excel File", type=["xlsx"], key="file_uploader_2f")

    if uploaded_file is not None:
        if mode == "Summarized Table Mode":
            run_summary_mode(uploaded_file)
        else:
            run_raw_mode(uploaded_file)

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
            elif "factor a × factor b" in val_str or "factor a*factor b" in val_str or "factor a x factor b" in val_str or "interaction" in val_str:
                idx_interaction = idx
            elif "grand mean" in val_str or "grandmean" in val_str: idx_grand = idx
                
        if any(v is None for v in [idx_A, idx_B, idx_cv, idx_interaction, idx_grand]):
            st.error("Missing structural markers (Factor A, Factor B, CV, Interaction, Grand Mean) in Column A.")
            return
            
        # Locate SEM, F-value, LSD for Factor A
        idx_A_sem, idx_A_f, idx_A_lsd = None, None, None
        for idx in range(idx_A + 1, idx_B):
            val = str(df_raw.iloc[idx, 0]).strip().lower()
            if "sem" in val: idx_A_sem = idx
            elif "f-value" in val or "f value" in val: idx_A_f = idx
            elif "lsd" in val: idx_A_lsd = idx
            
        # Locate SEM, F-value, LSD for Factor B
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
        st.success(f"Successfully detected Factors. A: {factor_a_label}, B: {factor_b_label}")
        
        group_1_name = st.text_input("First Table Title", "Effect of different factors on morphological attributes", key="2f_g1_name")
        group_1_cols = st.multiselect("Select parameters for Table 1", parameters, default=parameters[:len(parameters)//2], key="2f_g1_cols")
        group_2_name = st.text_input("Second Table Title", "Effect of different factors on yield and yield components", key="2f_g2_name")
        group_2_cols = st.multiselect("Select parameters for Table 2", [c for c in parameters if c not in group_1_cols], default=[c for c in parameters if c not in group_1_cols], key="2f_g2_cols")
        
        if st.button("Generate Word Document Draft", key="btn_2f_summary"):
            doc = Document()
            doc.add_heading("Summarized Factorial RCBD Report Draft", 0)
            
            groups = [(group_1_name, group_1_cols, 1), (group_2_name, group_2_cols, 2)]
            main_idx, both_ns_idx = 0, 0
            
            for g_title, g_cols, table_num in groups:
                if not g_cols: continue
                doc.add_heading(f"Table {table_num}: {g_title}", level=1)
                text_paragraphs = []
                
                for param in g_cols:
                    col_idx = df_raw.iloc[0].tolist().index(param)
                    
                    # Significance parsing
                    f_val_A = str(df_raw.iloc[idx_A_f, col_idx]).strip().lower()
                    f_val_B = str(df_raw.iloc[idx_B_f, col_idx]).strip().lower()
                    f_val_AB = str(df_raw.iloc[idx_interaction, col_idx]).strip().lower()
                    
                    sig_a, p_a = "significant", "(p < 0.05)"
                    if "ns" in f_val_A: sig_a, p_a = "nonsignificant", "(p > 0.05)"
                    elif "**" in f_val_A: p_a = "(p < 0.01)"
                    
                    sig_b, p_b = "significant", "(p < 0.05)"
                    if "ns" in f_val_B: sig_b, p_b = "nonsignificant", "(p > 0.05)"
                    elif "**" in f_val_B: p_b = "(p < 0.01)"
                    
                    sig_interaction, p_interaction = "significant", "(p < 0.05)"
                    if "ns" in f_val_AB: sig_interaction, p_interaction = "nonsignificant", "(p > 0.05)"
                    elif "**" in f_val_AB: p_interaction = "(p < 0.01)"
                    
                    # Factor A Parsing
                    a_means = []
                    for i, lvl in enumerate(factor_a_levels):
                        num, let = parse_dmrt_value(df_raw.iloc[idx_A + 1 + i, col_idx])
                        try: a_means.append((float(num), let, lvl))
                        except ValueError: pass
                    a_means.sort(reverse=True, key=lambda x: x[0])
                    top_val_a, top_let_a, top_a = a_means[0]
                    low_val_a, low_let_a, low_a = a_means[-1]
                    
                    at_par_a_list = []
                    for val, let, lvl in a_means[1:]:
                        if top_let_a and let and any(char in top_let_a for char in let):
                            at_par_a_list.append(f"{lvl} ({val:.2f}^{let})")
                    at_par_a_str = ", ".join(at_par_a_list) if at_par_a_list else "no other treatment levels"
                    
                    # Factor B Parsing
                    b_means = []
                    for i, lvl in enumerate(factor_b_levels):
                        num, let = parse_dmrt_value(df_raw.iloc[idx_B + 1 + i, col_idx])
                        try: b_means.append((float(num), let, lvl))
                        except ValueError: pass
                    b_means.sort(reverse=True, key=lambda x: x[0])
                    top_val_b, top_let_b, top_b = b_means[0]
                    low_val_b, low_let_b, low_b = b_means[-1]
                    
                    at_par_b_list = []
                    for val, let, lvl in b_means[1:]:
                        if top_let_b and let and any(char in top_let_b for char in let):
                            at_par_b_list.append(f"{lvl} ({val:.2f}^{let})")
                    at_par_b_str = ", ".join(at_par_b_list) if at_par_b_list else "no other treatment levels"
                    if sig_b == "nonsignificant":
                        at_par_b_str = "all application rates"
                        
                    grand_mean = str(df_raw.iloc[idx_grand, col_idx])
                    
                    # Template selection
                    if sig_interaction == "nonsignificant" and sig_a == "nonsignificant" and sig_b == "nonsignificant":
                        template = TEMPLATES_BOTH_NS[both_ns_idx % len(TEMPLATES_BOTH_NS)]
                        both_ns_idx += 1
                    elif sig_interaction == "nonsignificant":
                        template = TEMPLATES_MAIN_EFFECTS[main_idx % len(TEMPLATES_MAIN_EFFECTS)]
                        main_idx += 1
                    else:
                        template = TEMPLATES_SIGNIFICANT_INTERACTION[0]
                        
                    desc = template.format(
                        parameter=param, factor_a=factor_a_label, factor_b=factor_b_label,
                        sig_a=sig_a, p_a=p_a, sig_b=sig_b, p_b=p_b,
                        p_interaction=p_interaction, sig_interaction=sig_interaction,
                        top_a=top_a, top_val_a=f"{top_val_a:.2f}", top_let_a=top_let_a, at_par_a=at_par_a_str,
                        low_a=low_a, low_val_a=f"{low_val_a:.2f}", low_let_a=low_let_a,
                        top_b=top_b, top_val_b=f"{top_val_b:.2f}", top_let_b=top_let_b, at_par_b=at_par_b_str,
                        low_b=low_b, low_val_b=f"{low_val_b:.2f}", low_let_b=low_let_b,
                        grand_mean=grand_mean, combination_top="N/A", combination_low="N/A", at_par_comb="N/A"
                    )
                    text_paragraphs.append(desc)
                    
                for para in text_paragraphs:
                    p = doc.add_paragraph()
                    parts = re.split(r'(\*\*.*?\*\*)', para)
                    for part in parts:
                        if part.startswith('**') and part.endswith('**'):
                            p.add_run(part[2:-2]).bold = True
                        else:
                            p.add_run(part)
                            
                # Rebuild table inside Word Document
                table = doc.add_table(rows=1, cols=len(g_cols) + 1)
                set_table_borders(table)
                hdr_cells = table.rows[0].cells
                hdr_cells[0].text = "Treatments"
                set_cell_margins(hdr_cells[0])
                for c_idx, col_name in enumerate(g_cols):
                    hdr_cells[c_idx + 1].text = col_name
                    set_cell_margins(hdr_cells[c_idx + 1])
                set_header_bottom_border(table.rows[0])
                
                def append_row(target_idx, italic=False):
                    row_cells = table.add_row().cells
                    row_cells[0].text = str(df_raw.iloc[target_idx, 0])
                    set_cell_margins(row_cells[0])
                    if italic: row_cells[0].paragraphs[0].runs[0].italic = True
                    for c_idx, col_name in enumerate(g_cols):
                        c_raw_idx = df_raw.iloc[0].tolist().index(col_name)
                        row_cells[c_idx + 1].text = str(df_raw.iloc[target_idx, c_raw_idx])
                        set_cell_margins(row_cells[c_idx + 1])
                        
                # Add Factor A blocks
                append_row(idx_A, italic=True)
                for i in range(len(factor_a_levels)): append_row(idx_A + 1 + i)
                append_row(idx_A_sem)
                append_row(idx_A_f)
                append_row(idx_A_lsd)
                
                # Add Factor B blocks
                append_row(idx_B, italic=True)
                for i in range(len(factor_b_levels)): append_row(idx_B + 1 + i)
                append_row(idx_B_sem)
                append_row(idx_B_f)
                append_row(idx_B_lsd)
                
                # Add stats summary
                append_row(idx_cv)
                append_row(idx_interaction)
                append_row(idx_grand)
                doc.add_page_break()
                
            bio = io.BytesIO()
            doc.save(bio)
            bio.seek(0)
            st.success("🎉 Word Report generated from Summarized Two-Factor data!")
            st.download_button("Download Report (.docx)", data=bio, file_name="Summarized_Factorial_Report.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="btn_d_2f_sum_mod")
    except Exception as e:
        st.error(f"Error parsing Summarized Factorial Table: {e}")

# --- Raw Mode ---
def run_raw_mode(uploaded_file):
    try:
        df_raw_data = pd.read_excel(uploaded_file)
        st.write("#### Preview Raw Factorial Input Data:", df_raw_data.head())
        
        cols = df_raw_data.columns.tolist()
        
        # Following the standard R-script data sequence
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
            
            group_1_name = st.text_input("First Table Title", f"Effect of {treatment_label_a} and {treatment_label_b} on growth attributes", key="raw_2f_g1_name_m")
            group_1_cols = st.multiselect("Select parameters for Table 1", response_cols, default=response_cols[:len(response_cols)//2], key="raw_2f_g1_cols_m")
            group_2_name = st.text_input("Second Table Title", f"Effect of {treatment_label_a} and {treatment_label_b} on yield attributes", key="raw_2f_g2_name_m")
            group_2_cols = st.multiselect("Select parameters for Table 2", [c for c in response_cols if c not in group_1_cols], default=[c for c in response_cols if c not in group_1_cols], key="raw_2f_g2_cols_m")
            
            if st.button("Run Two-Factor Raw Analysis", key="btn_raw_2f_calc_m"):
                results_data = {}
                
                levels_a = sorted(df_raw_data[factor_a_col].unique().tolist())
                levels_b = sorted(df_raw_data[factor_b_col].unique().tolist())
                
                r = df_raw_data[block_col].nunique()
                a_levels = len(levels_a)
                b_levels = len(levels_b)
                
                # Dynamic statistical computations utilizing temporary standardized column mappings
                # to prevent Patsy parser failures on special characters
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
                    
                    # Factor A Main Effects
                    means_a = df_temp.groupby('factor_a')['response'].mean().to_dict()
                    sem_a = np.sqrt(mse / (r * b_levels))
                    lsd_a = t_val * np.sqrt((2 * mse) / (r * b_levels))
                    cld_a = get_cld_letters(means_a, lsd_a)
                    
                    # Factor B Main Effects
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
                
                # 1. DOWNLOAD FOR STYLED EXCEL WORKBOOK FIRST
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
                
                # 2. INTERPRETING RESULTS VIA CHRONOLOGICAL GROUPED ANALYSIS
                st.markdown("### 📝 Analysis Results and Academic Explanations")
                
                doc = Document()
                doc.add_heading("Calculated Two-Factor Factorial RCBD Report", 0)
                
                grouped_params = group_parameters(response_cols)
                
                # Display chronological parameter groupings
                for base_name, items in grouped_params.items():
                    day_strs = [d_str for _, _, d_str in items if d_str]
                    if day_strs:
                        if len(day_strs) == 1:
                            group_title = f"{base_name} at {day_strs[0]}"
                        elif len(day_strs) == 2:
                            group_title = f"{base_name} at {day_strs[0]} and {day_strs[1]}"
                        else:
                            group_title = f"{base_name} at {', '.join(day_strs[:-1])}, and {day_strs[-1]}"
                    else:
                        group_title = base_name
                        
                    st.write(f"### {group_title}")
                    doc.add_heading(group_title, level=2)
                    
                    paragraphs = []
                    for idx, (param_name, day_num, day_str) in enumerate(items):
                        p_data = results_data[param_name]
                        p_a, p_b, p_ab = p_data["p_a"], p_data["p_b"], p_data["p_ab"]
                        
                        sig_a_text = "significant" if p_a < 0.05 else "nonsignificant"
                        p_notation_a = f"(p < 0.05)" if p_a < 0.05 else "(p > 0.05)"
                        if p_a < 0.01: p_notation_a = "(p < 0.01)"
                        if p_a < 0.001: p_notation_a = "(p < 0.001)"
                        
                        sig_b_text = "significant" if p_b < 0.05 else "nonsignificant"
                        p_notation_b = f"(p < 0.05)" if p_b < 0.05 else "(p > 0.05)"
                        if p_b < 0.01: p_notation_b = "(p < 0.01)"
                        if p_b < 0.001: p_notation_b = "(p < 0.001)"
                        
                        sig_int = "significant" if p_ab < 0.05 else "nonsignificant"
                        p_notation_int = f"(p < 0.05)" if p_ab < 0.05 else "(p > 0.05)"
                        if p_ab < 0.01: p_notation_int = "(p < 0.01)"
                        if p_ab < 0.001: p_notation_int = "(p < 0.001)"
                        
                        sorted_a = sorted(p_data["means_a"].items(), key=lambda x: x[1], reverse=True)
                        top_a, top_val_a = sorted_a[0]
                        low_a, low_val_a = sorted_a[-1]
                        top_let_a = get_cld_letters(p_data["means_a"], p_data["lsd_a"]).get(top_a, "")
                        
                        sorted_b = sorted(p_data["means_b"].items(), key=lambda x: x[1], reverse=True)
                        top_b, top_val_b = sorted_b[0]
                        low_b, low_val_b = sorted_b[-1]
                        top_let_b = get_cld_letters(p_data["means_b"], p_data["lsd_b"]).get(top_b, "")
                        
                        sorted_comb = sorted(p_data["means_comb"].items(), key=lambda x: x[1], reverse=True)
                        comb_top_name, comb_top_val = sorted_comb[0]
                        comb_low_name, comb_low_val = sorted_comb[-1]
                        comb_top_let = p_data["cld_comb"].get(comb_top_name, "")
                        
                        at_par_a_list = []
                        for lvl, val in sorted_a[1:]:
                            let = get_cld_letters(p_data["means_a"], p_data["lsd_a"]).get(lvl, "")
                            if top_let_a and let and any(char in top_let_a for char in let):
                                at_par_a_list.append(f"{lvl} ({val:.2f}^{let})")
                        at_par_a_str = ", ".join(at_par_a_list) if at_par_a_list else "no other levels"
                        
                        at_par_b_list = []
                        for lvl, val in sorted_b[1:]:
                            let = get_cld_letters(p_data["means_b"], p_data["lsd_b"]).get(lvl, "")
                            if top_let_b and let and any(char in top_let_b for char in let):
                                at_par_b_list.append(f"{lvl} ({val:.2f}^{let})")
                        at_par_b_str = ", ".join(at_par_b_list) if at_par_b_list else "no other levels"
                        if sig_b_text == "nonsignificant":
                            at_par_b_str = "all evaluated levels"
                            
                        comb_top_let = p_data["cld_comb"].get(comb_top_name, "")
                        at_par_comb_list = []
                        for combo, val in sorted_comb[1:]:
                            let = p_data["cld_comb"].get(combo, "")
                            if comb_top_let and let and any(char in comb_top_let for char in let):
                                at_par_comb_list.append(f"{combo} ({val:.2f}^{let})")
                        at_par_comb_str = ", ".join(at_par_comb_list) if at_par_comb_list else "no other combinations"
                        
                        if idx == 0:
                            prefix = f"Initially, at {day_str if day_str else 'the beginning of the assessment'},"
                        elif idx == len(items) - 1:
                            prefix = f"Finally, at the later stage of {day_str},"
                        else:
                            prefix = f"As evaluation progressed to {day_str},"
                            
                        if p_a > 0.05:
                            part_a = f"{prefix} the main effect of `{factor_a_col}` was **nonsignificant** {p_notation_a} on **{param_name}**, suggesting that the variable did not induce notable variations, with values remaining near a mean of {top_val_a:.2f}."
                        else:
                            part_a = f"{prefix} the main effect of `{factor_a_col}` was **significant** {p_notation_a} on **{param_name}**. The maximum value was registered by `{top_a}` ({top_val_a:.2f}^{top_let_a}), which established statistical parity with `{at_par_a_str}`, while `{low_a}` ({low_val_a:.2f}) registered the minimum performance."
                            
                        if p_b > 0.05:
                            part_b = f"Similarly, the main effect of `{factor_b_col}` was **nonsignificant** {p_notation_b} at this interval, indicating that all treatment rates performed comparably with each other."
                        else:
                            part_b = f"Simultaneously, `{factor_b_col}` exerted a **significant** {p_notation_b} main effect. The superior tier was led by `{top_b}` ({top_val_b:.2f}^{top_let_b}), showing statistical parity with `{at_par_b_str}`, whereas `{low_b}` ({low_val_b:.2f}) marked the lowest performance."
                            
                        if p_ab > 0.05:
                            part_ab = f"The interaction effect between `{factor_a_col}` and `{factor_b_col}` was **nonsignificant** {p_notation_int}, showing that these two variables influenced **{param_name}** independently."
                        else:
                            part_ab = f"Importantly, a highly **significant** interaction {p_notation_int} was observed, confirming that the regulatory response of `{factor_b_col}` was modified by `{factor_a_col}`. Among all combinations, `{comb_top_name}` stood at the statistical apex with `{comb_top_val:.2f}^{comb_top_let}`, showing statistical parity with `{at_par_comb_str}`, while `{comb_low_name}` recorded the lowest mean performance."
                            
                        paragraphs.append(f"{part_a} {part_b} {part_ab}")
                        
                    for para in paragraphs:
                        st.write(para)
                        p_docx = doc.add_paragraph()
                        parts = re.split(r'(\*\*.*?\*\*)', para)
                        for part in parts:
                            if part.startswith('**') and part.endswith('**'):
                                p_docx.add_run(part[2:-2]).bold = True
                            else:
                                p_docx.add_run(part)
                    st.write("")
                    
                # 3. APPEND THE EXACT STYLED TABLES TO DOCX
                doc.add_page_break()
                groups_layouts = [
                    (group_1_name, group_1_cols, 1),
                    (group_2_name, group_2_cols, 2)
                ]
                
                for g_title, g_cols, table_num in groups_layouts:
                    if not g_cols: continue
                    doc.add_heading(f"Table {table_num}: {g_title}", level=1)
                    add_excel_table_to_docx(doc, factor_a_col, factor_b_col, g_cols, levels_a, levels_b, results_data)
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
