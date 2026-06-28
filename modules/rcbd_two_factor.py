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
from docx.shared import Inches, Pt
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

# --- 20 Q1 Factorial Templates ---
TEMPLATES_MAIN_EFFECTS = [
    "The main effect of `{factor_a}` on **{parameter}** was `{sig_a}` `{p_a}`. Genotype `{top_a}` (`{top_val_a}^{top_let_a}`) recorded the maximum value, which was statistically at par with `{at_par_a}`, whereas `{low_a}` (`{low_val_a}^{low_let_a}`) marked the lowest performance. In contrast, `{factor_b}` exerted a `{sig_b}` `{p_b}` main effect, indicating that `{at_par_b}` performed statistically equivalent and at par around the grand mean of `{grand_mean}`. Finally, the interaction effect between `{factor_a}` and `{factor_b}` was **nonsignificant** `{p_interaction}`, confirming that the two treatment factors influenced the trait independently [3].",
    "Regarding **{parameter}**, `{factor_a}` displayed a `{sig_a}` `{p_a}` main effect. The maximum performance was registered by `{top_a}` (`{top_val_a}^{top_let_a}`), which shared statistical parity with `{at_par_a}`, outclassing the lowest-performing entry `{low_a}` (`{low_val_a}^{low_let_a}`). For `{factor_b}`, the main effect was **nonsignificant** `{p_b}`, leaving `{at_par_b}` statistically equivalent and at par with each other. Ultimately, the interaction effect (`{factor_a}` × `{factor_b}`) was **nonsignificant** `{p_interaction}`, demonstrating a lack of synergistic dependency between the factors [3].",
    "Analysis of variance showed that `{factor_a}` had a `{sig_a}` `{p_a}` main effect on **{parameter}**. Genotype `{top_a}` achieved the apex value of `{top_val_a}^{top_let_a}`, remaining statistically at par with `{at_par_a}`, while the lowest boundary was marked by `{low_a}` (`{low_val_a}^{low_let_a}`). Conversely, the main effect of `{factor_b}` was **nonsignificant** `{p_b}`, confirming that `{at_par_b}` behaved statistically at par relative to the grand mean of `{grand_mean}`. In conclusion, the `{factor_a}` × `{factor_b}` interaction was verified as **nonsignificant** `{p_interaction}`, indicating that the two treatment factors act independently [3].",
    "Trait characterization of **{parameter}** demonstrated that `{factor_a}` exerted a `{sig_a}` `{p_a}` main effect. `{top_a}` (`{top_val_a}^{top_let_a}`) led the performance, showing statistical parity with `{at_par_a}`, while `{low_a}` (`{low_val_a}^{low_let_a}`) was significantly inferior. Meanwhile, `{factor_b}` main effects were completely **nonsignificant** `{p_b}`, meaning `{at_par_b}` remained statistically at par with each other. At last, the statistical interaction between `{factor_a}` and `{factor_b}` was **nonsignificant** `{p_interaction}`, pointing to a lack of combined dependency.",
    "A `{sig_a}` `{p_a}` genotypic influence was noted on **{parameter}** under the main effect of `{factor_a}`. `{top_a}` (`{top_val_a}^{top_let_a}`) established the upper threshold, remaining statistically at par with `{at_par_a}`, whereas `{low_a}` (`{low_val_a}^{low_let_a}`) represented the minimum. For `{factor_b}`, the main effect was **nonsignificant** `{p_b}`, indicating that `{at_par_b}` performed statistically equivalent and at par with each other. To conclude, the interaction effect (`{factor_a}` × `{factor_b}`) was **nonsignificant** `{p_interaction}`.",
    "Evaluating **{parameter}** revealed that `{factor_a}` significantly modified the trait `{p_a}`. The germplasm partitioned such that `{top_a}` (`{top_val_a}^{top_let_a}`) occupied the highest statistical tier, sharing statistical letters with `{at_par_a}`. The main effect of `{factor_b}` was **nonsignificant** `{p_b}`, leaving `{at_par_b}` statistically equivalent and at par around the grand mean of `{grand_mean}`. At the end, the interaction effect (`{factor_a}` × `{factor_b}`) was confirmed as **nonsignificant** `{p_interaction}`.",
    "Regarding **{parameter}**, `{factor_a}` main effects were `{sig_a}` `{p_a}`. `{top_a}` (`{top_val_a}^{top_let_a}`) stood out as the leading treatment, showing statistical parity with `{at_par_a}`. However, `{factor_b}` main effects proved to be **nonsignificant** `{p_b}`, leaving `{at_par_b}` statistically at par under the trial environment. In conclusion, the interaction `{factor_a}` × `{factor_b}` was **nonsignificant** `{p_interaction}`, showing that the factors act independently [3].",
    "The main effect of `{factor_a}` on **{parameter}** was `{sig_a}` `{p_a}`. `{top_a}` (`{top_val_a}^{top_let_a}`) achieved the maximum mean, showing statistical parity with `{at_par_a}`, while `{low_a}` (`{low_val_a}^{low_let_a}`) marked the lower limit. Conversely, the main effect of `{factor_b}` was **nonsignificant** `{p_b}`, which meant `{at_par_b}` co-occupied the same statistical tier. Finally, the additive action of `{factor_a}` and `{factor_b}` was confirmed by a **nonsignificant** `{p_interaction}` interaction [3].",
    "For the parameter **{parameter}**, the main effect of `{factor_a}` was `{sig_a}` `{p_a}`. Genotype `{top_a}` (`{top_val_a}^{top_let_a}`) led the performance, showing statistical parity with `{at_par_a}`. In contrast, the main effect of `{factor_b}` was completely **nonsignificant** `{p_b}`, confirming that `{at_par_b}` behaved statistically at par relative to the grand mean of `{grand_mean}`. Ultimately, the interaction of `{factor_a}` × `{factor_b}` was verified as **nonsignificant** `{p_interaction}`.",
    "Trait characterization of **{parameter}** demonstrated `{sig_a}` `{p_a}` variation under `{factor_a}`. `{top_a}` (`{top_val_a}^{top_let_a}`) registered the maximum mean value, showing statistical parity with `{at_par_a}`. However, `{factor_b}` exerted a `{sig_b}` `{p_b}` main effect, meaning `{at_par_b}` remained statistically equivalent. At last, the interaction effect between `{factor_a}` and `{factor_b}` was **nonsignificant** `{p_interaction}`, confirming independent factor action."
]

TEMPLATES_BOTH_NS = [
    "For **{parameter}**, the main effect of `{factor_a}` was **nonsignificant** `(p > 0.05)`, indicating that `{at_par_b}` performed statistically equivalent. Similarly, the main effect of `{factor_b}` was **nonsignificant** `(p > 0.05)`, leaving all treatments statistically at par. Finally, the interaction effect (`{factor_a}` × `{factor_b}`) was **nonsignificant** `(p > 0.05)`, confirming a complete lack of phenotypic stratification across the entire trial.",
    "The main effect of `{factor_a}` on **{parameter}** was **nonsignificant** `(p > 0.05)`, leaving all treatment levels statistically at par. This flat trend was mirrored in `{factor_b}`, where the main effect was also **nonsignificant** `(p > 0.05)`, with `{top_b}` (`{top_val_b}`) and `{low_b}` (`{low_val_b}`) remaining statistically equivalent. In conclusion, the interaction effect between `{factor_a}` and `{factor_b}` was **nonsignificant** `(p > 0.05)`, demonstrating highly stable trait expression.",
    "Regarding **{parameter}**, the main effect of `{factor_a}` proved to be **nonsignificant** `(p > 0.05)`, with `{at_par_b}` co-occupying the same statistical tier. The main effect of `{factor_b}` was also **nonsignificant** `(p > 0.05)`, indicating that all levels performed statistically at par around the grand mean of `{grand_mean}`. Ultimately, the interaction of `{factor_a}` × `{factor_b}` was verified as **nonsignificant** `(p > 0.05)`, confirming the mutual independence of the factors.",
    "Analysis of variance for **{parameter}** showed that `{factor_a}` had a **nonsignificant** `(p > 0.05)` main effect, indicating that `{at_par_b}` possessed equivalent performance potential. The main effect of `{factor_b}` was similarly **nonsignificant** `(p > 0.05)`, leaving all treatments statistically at par. At the end, the interaction effect (`{factor_a}` × `{factor_b}`) was confirmed as **nonsignificant** `(p > 0.05)`, pointing to a homogenous response pattern.",
    "The main effect of `{factor_a}` on **{parameter}** was **nonsignificant** `(p > 0.05)`, leaving `{at_par_b}` statistically equivalent. For `{factor_b}`, the main effect was also **nonsignificant** `(p > 0.05)`, with all treatment levels performing statistically at par relative to the grand mean of `{grand_mean}`. At last, the interaction effect between `{factor_a}` and `{factor_b}` was **nonsignificant** `(p > 0.05)`, indicating a lack of combined dependency."
]

TEMPLATES_SIGNIFICANT_INTERACTION = [
    "The main effect of `{factor_a}` on **{parameter}** was `{sig_a}` `{p_a}`, while `{factor_b}` main effects were `{sig_b}` `{p_b}`. However, evaluating main effects independently was not sufficient as the interaction `{factor_a}` × `{factor_b}` was highly **significant** `{p_interaction}` [2]. Consequently, simple effects comparisons were analyzed, revealing that `{combination_top}` registered the maximum mean, remaining statistically at par with `{at_par_comb}`, while `{combination_low}` marked the lowest limit of performance.",
    "For **{parameter}**, `{factor_a}` showed a `{sig_a}` `{p_a}` main effect, whereas `{factor_b}` showed a `{sig_b}` `{p_b}` main effect. Ultimately, the interaction of `{factor_a}` × `{factor_b}` was **significant** `{p_interaction}`, showing that the response of `{factor_a}` depended heavily on `{factor_b}`. Because the factors did not act independently, treatment combinations were evaluated, where `{combination_top}` stood at the statistical apex, showing statistical parity with `{at_par_comb}`, while `{combination_low}` was significantly inferior.",
    "Regarding **{parameter}**, the main effect of `{factor_a}` was `{sig_a}` `{p_a}` and the main effect of `{factor_b}` was `{sig_b}` `{p_b}`. In conclusion, the interaction effect between the two factors was **significant** `{p_interaction}` [2]. This dependency is highlighted by the superior performance of the `{combination_top}` combination, which shared statistical letters with `{at_par_comb}`, whereas the lowest trait manifestation was restricted to `{combination_low}`.",
    "Analysis of variance indicated that the main effects of `{factor_a}` (`{p_a}`) and `{factor_b}` (`{p_b}`) significantly modified **{parameter}**. At last, the interaction effect (`{factor_a}` × `{factor_b}`) proved to be highly **significant** `{p_interaction}`. Simple effect comparisons showed that the combination `{combination_top}` achieved superior performance, remaining statistically at par with `{at_par_comb}`, while `{combination_low}` represented the absolute minimum.",
    "The main effect of `{factor_a}` on **{parameter}** was `{sig_a}` `{p_a}`, and `{factor_b}` had a `{sig_b}` `{p_b}` main effect. Ultimately, the ANOVA demonstrated a highly **significant** `{p_interaction}` interaction [2]. Evaluating the treatment combinations revealed that `{combination_top}` was the optimal treatment combination, remaining statistically at par with `{at_par_comb}`, while the lowest value was recorded in `{combination_low}`."
]

# --- Cell Margins & Border Styling ---
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

# --- Main Entry Point ---
def show_module():
    st.markdown("### RCBD Two-Factor Factorial Analyzer")
    mode = st.radio("Choose Input Mode", ["Summarized Table Mode", "Raw Data Mode"], key="2f_mode")
    uploaded_file = st.file_uploader("Upload Two-Factor Excel File", type=["xlsx"], key="file_uploader_2f")

    if uploaded_file is not None:
        if mode == "Summarized Table Mode":
            run_summary_mode(uploaded_file)
        else:
            run_raw_mode(uploaded_file)

# --- Summary Table Mode ---
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
                            
                # Rebuild table
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
        factor_a_col = st.selectbox("Select Factor A Column", cols, index=0, key="raw_2f_fa_mod")
        factor_b_col = st.selectbox("Select Factor B Column", cols, index=1, key="raw_2f_fb_mod")
        block_col = st.selectbox("Select Block/Replication Column", cols, index=2, key="raw_2y_bk_f_mod")
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
            
            if st.button("Generate Two-Factor Word Report from Raw Data", key="btn_raw_2f_calc_m"):
                doc = Document()
                doc.add_heading("Calculated Two-Factor Factorial RCBD Report", 0)
                
                groups = [(group_1_name, group_1_cols, 1), (group_2_name, group_2_cols, 2)]
                main_idx, both_ns_idx, int_idx = 0, 0, 0
                
                for g_title, g_cols, table_num in groups:
                    if not g_cols: continue
                    doc.add_heading(f"Table {table_num}: {g_title}", level=1)
                    text_paragraphs = []
                    
                    calculated_means_for_table = {}
                    calculated_stats_for_table = {k: {} for k in ["sem_a", "lsd_a", "sem_b", "lsd_b", "cv", "interaction", "grand mean"]}
                    
                    for param in g_cols:
                        formula = f"Q('{param}') ~ C(Q('{block_col}')) + C(Q('{factor_a_col}')) + C(Q('{factor_b_col}')) + C(Q('{factor_a_col}')):C(Q('{factor_b_col}'))"
                        model = ols(formula, data=df_raw_data).fit()
                        anova_table = sm.stats.anova_lm(model, typ=1)
                        
                        df_err = anova_table.loc["Residual", "df"]
                        mse = anova_table.loc["Residual", "mean_sq"]
                        
                        p_a = anova_table.loc[f"C(Q('{factor_a_col}'))", "PR(>F)"]
                        p_b = anova_table.loc[f"C(Q('{factor_b_col}'))", "PR(>F)"]
                        p_ab = anova_table.loc[f"C(Q('{factor_a_col}')):C(Q('{factor_b_col}'))", "PR(>F)"]
                        
                        r = df_raw_data[block_col].nunique()
                        a_levels = df_raw_data[factor_a_col].nunique()
                        b_levels = df_raw_data[factor_b_col].nunique()
                        
                        grand_mean = df_raw_data[param].mean()
                        cv = (np.sqrt(mse) / grand_mean) * 100
                        t_val = t.ppf(0.975, df_err)
                        
                        # Factor A Main Effects
                        means_a = df_raw_data.groupby(factor_a_col)[param].mean().to_dict()
                        sem_a = np.sqrt(mse / (r * b_levels))
                        lsd_a = t_val * np.sqrt((2 * mse) / (r * b_levels))
                        cld_a = get_cld_letters(means_a, lsd_a)
                        
                        # Factor B Main Effects
                        means_b = df_raw_data.groupby(factor_b_col)[param].mean().to_dict()
                        sem_b = np.sqrt(mse / (r * a_levels))
                        lsd_b = t_val * np.sqrt((2 * mse) / (r * a_levels))
                        cld_b = get_cld_letters(means_b, lsd_b)
                        
                        # Combinations (For significant interaction)
                        df_raw_data['Combination'] = df_raw_data[factor_a_col] + " × " + df_raw_data[factor_b_col]
                        means_comb = df_raw_data.groupby('Combination')[param].mean().to_dict()
                        sem_comb = np.sqrt(mse / r)
                        lsd_comb = t_val * np.sqrt((2 * mse) / r)
                        cld_comb = get_cld_letters(means_comb, lsd_comb)
                        
                        for t_name, val in means_a.items():
                            if t_name not in calculated_means_for_table:
                                calculated_means_for_table[t_name] = {}
                            calculated_means_for_table[t_name][param] = f"{val:.2f}{cld_a.get(t_name, '')}"
                            
                        for t_name, val in means_b.items():
                            if t_name not in calculated_means_for_table:
                                calculated_means_for_table[t_name] = {}
                            calculated_means_for_table[t_name][param] = f"{val:.2f}{cld_b.get(t_name, '')}"
                            
                        calculated_stats_for_table["sem_a"][param] = f"{sem_a:.2f}"
                        calculated_stats_for_table["lsd_a"][param] = f"{lsd_a:.2f}"
                        calculated_stats_for_table["sem_b"][param] = f"{sem_b:.2f}"
                        calculated_stats_for_table["lsd_b"][param] = f"{lsd_b:.2f}"
                        calculated_stats_for_table["cv"][param] = f"{cv:.2f}"
                        calculated_stats_for_table["grand mean"][param] = f"{grand_mean:.2f}"
                        
                        # Mapping Significance Texts
                        sig_int = "significant"
                        p_notation_int = "(p < 0.05)"
                        if p_ab > 0.05:
                            sig_int, p_notation_int = "nonsignificant", "(p > 0.05)"
                            calculated_stats_for_table["interaction"][param] = "ns"
                        elif p_ab < 0.001:
                            p_notation_int = "(p < 0.001)"
                            calculated_stats_for_table["interaction"][param] = "***"
                        elif p_ab < 0.01:
                            p_notation_int = "(p < 0.01)"
                            calculated_stats_for_table["interaction"][param] = "**"
                        else:
                            calculated_stats_for_table["interaction"][param] = "*"
                            
                        sig_a_text, p_notation_a = "significant", "(p < 0.05)"
                        if p_a > 0.05: sig_a_text, p_notation_a = "nonsignificant", "(p > 0.05)"
                        elif p_a < 0.001: p_notation_a = "(p < 0.001)"
                        elif p_a < 0.01: p_notation_a = "(p < 0.01)"
                        
                        sig_b_text, p_notation_b = "significant", "(p < 0.05)"
                        if p_b > 0.05: sig_b_text, p_notation_b = "nonsignificant", "(p > 0.05)"
                        elif p_b < 0.001: p_notation_b = "(p < 0.001)"
                        elif p_b < 0.01: p_notation_b = "(p < 0.01)"
                        
                        sorted_a = sorted(means_a.items(), key=lambda x: x[1], reverse=True)
                        top_a, top_val_a = sorted_a[0]
                        top_let_a = cld_a[top_a]
                        low_a, low_val_a = sorted_a[-1]
                        
                        sorted_b = sorted(means_b.items(), key=lambda x: x[1], reverse=True)
                        top_b, top_val_b = sorted_b[0]
                        low_b, low_val_b = sorted_b[-1]
                        
                        sorted_comb = sorted(means_comb.items(), key=lambda x: x[1], reverse=True)
                        comb_top_name, comb_top_val = sorted_comb[0]
                        comb_low_name, comb_low_val = sorted_comb[-1]
                        
                        at_par_a_list = []
                        for lvl, val in sorted_a[1:]:
                            let = cld_a[lvl]
                            if top_let_a and let and any(char in top_let_a for char in let):
                                at_par_a_list.append(f"{lvl} ({val:.2f}^{let})")
                        at_par_a_str = ", ".join(at_par_a_list) if at_par_a_list else "no other treatment levels"
                        
                        at_par_b_list = []
                        top_let_b = cld_b.get(top_b, "")
                        for lvl, val in sorted_b[1:]:
                            let = cld_b.get(lvl, "")
                            if top_let_b and let and any(char in top_let_b for char in let):
                                at_par_b_list.append(f"{lvl} ({val:.2f}^{let})")
                        at_par_b_str = ", ".join(at_par_b_list) if at_par_b_list else "no other treatment levels"
                        if sig_b_text == "nonsignificant":
                            at_par_b_str = "all application rates"
                            
                        comb_top_let = cld_comb.get(comb_top_name, "")
                        at_par_comb_list = []
                        for combo, val in sorted_comb[1:]:
                            let = cld_comb.get(combo, "")
                            if comb_top_let and let and any(char in comb_top_let for char in let):
                                at_par_comb_list.append(f"{combo} ({val:.2f}^{let})")
                        at_par_comb_str = ", ".join(at_par_comb_list) if at_par_comb_list else "no other combinations"
                        
                        if sig_int == "significant":
                            template = TEMPLATES_SIGNIFICANT_INTERACTION[int_idx % len(TEMPLATES_SIGNIFICANT_INTERACTION)]
                            int_idx += 1
                        elif sig_a_text == "nonsignificant" and sig_b_text == "nonsignificant":
                            template = TEMPLATES_BOTH_NS[both_ns_idx % len(TEMPLATES_BOTH_NS)]
                            both_ns_idx += 1
                        else:
                            template = TEMPLATES_MAIN_EFFECTS[main_idx % len(TEMPLATES_MAIN_EFFECTS)]
                            main_idx += 1
                            
                        desc = template.format(
                            parameter=param, factor_a=factor_a_col, factor_b=factor_b_col,
                            sig_a=sig_a_text, p_a=p_notation_a, sig_b=sig_b_text, p_b=p_notation_b,
                            p_interaction=p_notation_int, sig_interaction=sig_int,
                            top_a=top_a, top_val_a=f"{top_val_a:.2f}", top_let_a=top_let_a, at_par_a=at_par_a_str,
                            low_a=low_a, low_val_a=f"{low_val_a:.2f}", low_let_a=cld_a.get(low_a, ""),
                            top_b=top_b, top_val_b=f"{top_val_b:.2f}", top_let_b=cld_b.get(top_b, ""), at_par_b=at_par_b_str,
                            low_b=low_b, low_val_b=f"{low_val_b:.2f}", low_let_b=cld_b.get(low_b, ""),
                            grand_mean=f"{grand_mean:.2f}",
                            combination_top=comb_top_name, combination_low=comb_low_name, at_par_comb=at_par_comb_str
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
                                
                    # Reconstruct Table
                    table = doc.add_table(rows=1, cols=len(g_cols) + 1)
                    set_table_borders(table)
                    hdr_cells = table.rows[0].cells
                    hdr_cells[0].text = "Treatments"
                    set_cell_margins(hdr_cells[0])
                    for c_idx, col_name in enumerate(g_cols):
                        hdr_cells[c_idx + 1].text = col_name
                        set_cell_margins(hdr_cells[c_idx + 1])
                    set_header_bottom_border(table.rows[0])
                    
                    def add_header_row(text):
                        row_cells = table.add_row().cells
                        row_cells[0].text = text
                        set_cell_margins(row_cells[0])
                        row_cells[0].paragraphs[0].runs[0].italic = True
                        for i in range(len(g_cols)):
                            row_cells[i+1].text = ""
                            set_cell_margins(row_cells[i+1])
                            
                    def add_data_row(lbl, val_dict):
                        row_cells = table.add_row().cells
                        row_cells[0].text = str(lbl)
                        set_cell_margins(row_cells[0])
                        for c_idx, col_name in enumerate(g_cols):
                            row_cells[c_idx + 1].text = str(val_dict.get(col_name, "N/A"))
                            set_cell_margins(row_cells[c_idx + 1])
                            
                    # Add Factor A blocks
                    add_header_row(f"Factor A: ({factor_a_col})")
                    for lvl in sorted(means_a.keys()):
                        add_data_row(lvl, calculated_means_for_table[lvl])
                    add_data_row("SEm(±)", calculated_stats_for_table["sem_a"])
                    add_data_row("LSD(0.05)", calculated_stats_for_table["lsd_a"])
                    
                    # Add Factor B blocks
                    add_header_row(f"Factor B: ({factor_b_col})")
                    for lvl in sorted(means_b.keys()):
                        add_data_row(lvl, calculated_means_for_table[lvl])
                    add_data_row("SEm(±)", calculated_stats_for_table["sem_b"])
                    add_data_row("LSD(0.05)", calculated_stats_for_table["lsd_b"])
                    
                    # Add summary rows
                    add_data_row("CV, %", calculated_stats_for_table["cv"])
                    add_data_row("Factor A × Factor B", calculated_stats_for_table["interaction"])
                    add_data_row("Grand Mean", calculated_stats_for_table["grand mean"])
                    
                    doc.add_page_break()
                    
                bio = io.BytesIO()
                doc.save(bio)
                bio.seek(0)
                st.success("🎉 Word Report successfully built from Combined Raw Two-Factor data!")
                st.download_button("Download Report (.docx)", data=bio, file_name="Calculated_Factorial_Report.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="btn_d_2f_raw_cal")
    except Exception as e:
        st.error(f"Error executing raw combined Two-Factor analysis: {e}")
