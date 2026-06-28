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

# --- 20 Multi-Year Long-Form Q1 Templates ---
COMBINED_TEMPLATES = [
    "The genotypic performance for **{parameter}** was found to be `{sig_y1}` `{p_y1}` in `{year1}` and `{sig_y2}` `{p_y2}` in `{year2}`. During `{year1}`, the highest value was registered by `{top_geno_y1}` (`{top_val_y1}^{top_let_y1}`), which was statistically at par with `{at_par_y1}`. In `{year2}`, however, `{top_geno_y2}` (`{top_val_y2}^{top_let_y2}`) stood out as the leading treatment, showing statistical parity with `{at_par_y2}`. Over both seasons combined, the pooled data indicated that `{pooled_top_geno}` maintained the highest cumulative mean of `{pooled_top_val}`, which was substantially superior compared to the lowest pooled mean of `{pooled_low_val}` recorded by `{pooled_low_geno}`.",
    "Regarding **{parameter}**, analysis of variance showed a `{sig_y1}` `{p_y1}` treatment effect in `{year1}` and a `{sig_y2}` `{p_y2}` effect in `{year2}`. The crop cycle of `{year1}` was dominated by `{top_geno_y1}` (`{top_val_y1}^{top_let_y1}`), which was statistically at par with `{at_par_y1}`. A seasonal shift in vigor occurred in `{year2}` where `{top_geno_y2}` (`{top_val_y2}^{top_let_y2}`) recorded the maximum expression, remaining at par with `{at_par_y2}`. The two-year pooled evaluation established `{pooled_top_geno}` as the top performer with a combined mean of `{pooled_top_val}`, whereas `{pooled_low_geno}` was confined to the lowest pooled boundary at `{pooled_low_val}`.",
    "The evaluated `{treatment_plural}` exhibited `{sig_y1}` `{p_y1}` variations for **{parameter}** during `{year1}`, and the trend remained `{sig_y2}` `{p_y2}` during `{year2}`. For `{year1}`, `{top_geno_y1}` (`{top_val_y1}^{top_let_y1}`) led the trial, showing statistical parity with `{at_par_y1}`. In the subsequent `{year2}` cycle, `{top_geno_y2}` (`{top_val_y2}^{top_let_y2}`) achieved the absolute peak, with no statistical difference observed between it and `{at_par_y2}`. When looking at the pooled means across both years, `{pooled_top_geno}` (`{pooled_top_val}`) proved to be the most robust `{treatment_singular}`, outclassing `{pooled_low_geno}` which recorded the lowest combined mean of `{pooled_low_val}`.",
    "Trait characterization of **{parameter}** revealed `{sig_y1}` `{p_y1}` difference among `{treatment_plural}` in `{year1}` and `{sig_y2}` `{p_y2}` in `{year2}`. The highest treatment mean during `{year1}` was obtained from `{top_geno_y1}` (`{top_val_y1}^{top_let_y1}`), which was statistically at par with `{at_par_y1}`. In `{year2}`, `{top_geno_y2}` (`{top_val_y2}^{top_let_y2}`) occupied the peak group, sharing statistical parity with `{at_par_y2}`. The combined response (pooled column) revealed that `{pooled_top_geno}` maintained a dominant cumulative mean of `{pooled_top_val}`, whereas `{pooled_low_geno}` showed the lowest collective average of `{pooled_low_val}`.",
    "A `{sig_y1}` `{p_y1}` genotypic influence was noted on **{parameter}** in `{year1}`, which remained `{sig_y2}` `{p_y2}` in `{year2}`. The `{year1}` season highlighted `{top_geno_y1}` (`{top_val_y1}^{top_let_y1}`) as the premier treatment, showing statistical parity with `{at_par_y1}`. In contrast, `{top_geno_y2}` (`{top_val_y2}^{top_let_y2}`) excelled in `{year2}`, remaining statistically at par with `{at_par_y2}`. The pooled data across the two years demonstrated that `{pooled_top_geno}` emerged as the outstanding treatment with an average of `{pooled_top_val}`, which was significantly higher than the lowest pooled value of `{pooled_low_val}` registered by `{pooled_low_geno}`.",
    "Significant genotypic variation was observed for **{parameter}**, which was `{sig_y1}` `{p_y1}` in `{year1}` and `{sig_y2}` `{p_y2}` in `{year2}`. For `{year1}`, the response was led by `{top_geno_y1}` (`{top_val_y1}^{top_let_y1}`), which remained statistically equivalent to `{at_par_y1}`. In `{year2}`, `{top_geno_y2}` (`{top_val_y2}^{top_let_y2}`) co-occupied the top statistical tier along with `{at_par_y2}`. Over both seasons, `{pooled_top_geno}` maintained the highest pooled performance of `{pooled_top_val}`, while `{pooled_low_geno}` fell to the lowest pooled boundary of `{pooled_low_val}`.",
    "The phenotypic expression of **{parameter}** was `{sig_y1}` `{p_y1}` in `{year1}` and `{sig_y2}` `{p_y2}` in `{year2}`. In `{year1}`, the maximum mean was attained by `{top_geno_y1}` (`{top_val_y1}^{top_let_y1}`), showing no significant difference from `{at_par_y1}`. During the `{year2}` crop cycle, `{top_geno_y2}` (`{top_val_y2}^{top_let_y2}`) emerged as the top performer, remaining statistically at par with `{at_par_y2}`. Evaluating the pooled averages over the two-year period, `{pooled_top_geno}` (`{pooled_top_val}`) established the maximum combined value, whereas the minimum combined value was recorded in `{pooled_low_geno}` (`{pooled_low_val}`).",
    "In terms of **{parameter}**, the treatments exhibited `{sig_y1}` `{p_y1}` and `{sig_y2}` `{p_y2}` differences during the `{year1}` and `{year2}` seasons, respectively. The first year favored `{top_geno_y1}` (`{top_val_y1}^{top_let_y1}`), which shared statistical parity with `{at_par_y1}`. The second year was dominated by `{top_geno_y2}` (`{top_val_y2}^{top_let_y2}`), which was statistically at par with `{at_par_y2}`. When the seasons were pooled, `{pooled_top_geno}` was verified as the most productive `{treatment_singular}` with a combined average of `{pooled_top_val}`, while `{pooled_low_geno}` stayed at the lower limit of `{pooled_low_val}`.",
    "Germplasm evaluation for **{parameter}** revealed a `{sig_y1}` `{p_y1}` effect in `{year1}` and a `{sig_y2}` `{p_y2}` effect in `{year2}`. During the `{year1}` evaluation, `{top_geno_y1}` (`{top_val_y1}^{top_let_y1}`) secured the highest rank, staying statistically at par with `{at_par_y1}`. For `{year2}`, `{top_geno_y2}` (`{top_val_y2}^{top_let_y2}`) stood out in the upper statistical tier, showing parity with `{at_par_y2}`. The pooled data across the years indicated that `{pooled_top_geno}` (`{pooled_top_val}`) occupied the highest overall position, significantly outperforming `{pooled_low_geno}` which recorded the lowest pooled mean of `{pooled_low_val}`.",
    "The treatment influence on **{parameter}** was verified as `{sig_y1}` `{p_y1}` in `{year1}` and `{sig_y2}` `{p_y2}` in `{year2}`. `{top_geno_y1}` (`{top_val_y1}^{top_let_y1}`) registered the maximum value in `{year1}`, remaining statistically equivalent to `{at_par_y1}`. During `{year2}`, the highest mean belonged to `{top_geno_y2}` (`{top_val_y2}^{top_let_y2}`), which was statistically at par with `{at_par_y2}`. The combined multi-year pooled average confirmed `{pooled_top_geno}` as the leading treatment at `{pooled_top_val}`, while `{pooled_low_geno}` remained significantly inferior at `{pooled_low_val}`.",
    "Analyzing **{parameter}** chronologically showed a `{sig_y1}` `{p_y1}` genotypic variation in `{year1}` and `{sig_y2}` `{p_y2}` variation in `{year2}`. In the `{year1}` crop cycle, `{top_geno_y1}` (`{top_val_y1}^{top_let_y1}`) attained the peak value, sharing statistical parity with `{at_par_y1}`. During the `{year2}` cycle, `{top_geno_y2}` (`{top_val_y2}^{top_let_y2}`) secured the highest statistical grouping alongside `{at_par_y2}`. The combined pooled means across the two trials highlighted `{pooled_top_geno}` as the superior treatment with a pooled mean of `{pooled_top_val}`, outperforming the lowest combined average of `{pooled_low_val}` recorded by `{pooled_low_geno}`.",
    "The magnitude of **{parameter}** expression was `{sig_y1}` `{p_y1}` during the `{year1}` season and `{sig_y2}` `{p_y2}` during the `{year2}` season. The highest value in `{year1}` was observed in `{top_geno_y1}` (`{top_val_y1}^{top_let_y1}`), which was statistically at par with `{at_par_y1}`. In `{year2}`, `{top_geno_y2}` (`{top_val_y2}^{top_let_y2}`) achieved the leading position, remaining at par with `{at_par_y2}`. The pooled response across both seasons established `{pooled_top_geno}` as the outstanding treatment at `{pooled_top_val}`, which was significantly higher than the lowest pooled value of `{pooled_low_val}` noted in `{pooled_low_geno}`.",
    "Adaptation trials evaluating **{parameter}** demonstrated `{sig_y1}` `{p_y1}` differences in `{year1}` and `{sig_y2}` `{p_y2}` differences in `{year2}`. `{top_geno_y1}` (`{top_val_y1}^{top_let_y1}`) co-occupied the top statistical tier in `{year1}` with `{at_par_y1}`. During the `{year2}` trial, `{top_geno_y2}` (`{top_val_y2}^{top_let_y2}`) emerged at the apex, showing statistical parity with `{at_par_y2}`. The multi-year pooled analysis confirmed `{pooled_top_geno}` (`{pooled_top_val}`) as the most consistent treatment, whereas the lowest pooled performance was observed in `{pooled_low_geno}` (`{pooled_low_val}`).",
    "Evaluating the agronomic performance of **{parameter}** revealed a `{sig_y1}` `{p_y1}` response in `{year1}` and a `{sig_y2}` `{p_y2}` response in `{year2}`. During the `{year1}` crop season, `{top_geno_y1}` (`{top_val_y1}^{top_let_y1}`) took the lead, showing statistical parity with `{at_par_y1}`. In `{year2}`, `{top_geno_y2}` (`{top_val_y2}^{top_let_y2}`) attained the top position, remaining statistically equivalent to `{at_par_y2}`. Over both seasons, `{pooled_top_geno}` demonstrated the highest pooled mean of `{pooled_top_val}`, representing a substantial increase over `{pooled_low_geno}` which remained at the pooled bottom (`{pooled_low_val}`).",
    "Across distinct trial cycles, **{parameter}** was significantly influenced by the treatments, with `{sig_y1}` `{p_y1}` effects in `{year1}` and `{sig_y2}` `{p_y2}` effects in `{year2}`. `{top_geno_y1}` (`{top_val_y1}^{top_let_y1}`) registered the maximum value in `{year1}`, remaining statistically at par with `{at_par_y1}`. In `{year2}`, `{top_geno_y2}` (`{top_val_y2}^{top_let_y2}`) occupied the highest statistical group, remaining at par with `{at_par_y2}`. The combined pooled means across both seasons confirmed `{pooled_top_geno}` as the outstanding treatment with `{pooled_top_val}`, while `{pooled_low_geno}` was confined to the lowest pooled boundary at `{pooled_low_val}`.",
    "Regarding **{parameter}**, mean discrepancies among `{treatment_plural}` were found to be `{sig_y1}` `{p_y1}` in `{year1}` and `{sig_y2}` `{p_y2}` in `{year2}`. `{top_geno_y1}` (`{top_val_y1}^{top_let_y1}`) secured the maximum value in `{year1}`, showing statistical parity with `{at_par_y1}`. In `{year2}`, `{top_geno_y2}` (`{top_val_y2}^{top_let_y2}`) emerged at the top, remaining statistically equivalent to `{at_par_y2}`. The pooled data across the years indicated that `{pooled_top_geno}` (`{pooled_top_val}`) occupied the highest overall position, significantly outperforming the lowest pooled mean of `{pooled_low_val}` recorded by `{pooled_low_geno}`.",
    "The treatment influence on **{parameter}** was confirmed as `{sig_y1}` `{p_y1}` in `{year1}` and `{sig_y2}` `{p_y2}` in `{year2}`. `{top_geno_y1}` (`{top_val_y1}^{top_let_y1}`) attained the peak value in `{year1}`, sharing statistical parity with `{at_par_y1}`. During the `{year2}` cycle, the highest mean belonged to `{top_geno_y2}` (`{top_val_y2}^{top_let_y2}`), which was statistically at par with `{at_par_y2}`. The combined multi-year pooled average confirmed `{pooled_top_geno}` as the leading treatment at `{pooled_top_val}`, while `{pooled_low_geno}` remained significantly inferior at `{pooled_low_val}`.",
    "The response pattern of `{treatment_plural}` for **{parameter}** was `{sig_y1}` `{p_y1}` in `{year1}` and `{sig_y2}` `{p_y2}` in `{year2}`. In `{year1}`, the highest value was registered by `{top_geno_y1}` (`{top_val_y1}^{top_let_y1}`), which was statistically at par with `{at_par_y1}`. In `{year2}`, `{top_geno_y2}` (`{top_val_y2}^{top_let_y2}`) stood out as the leading treatment, showing statistical parity with `{at_par_y2}`. Over both seasons combined, the pooled data indicated that `{pooled_top_geno}` maintained the highest cumulative mean of `{pooled_top_val}`, which was substantially superior compared to the lowest pooled mean of `{pooled_low_val}` recorded by `{pooled_low_geno}`.",
    "Statistical analysis confirmed that the genotypic effect on **{parameter}** was `{sig_y1}` `{p_y1}` in `{year1}` and `{sig_y2}` `{p_y2}` in `{year2}`. `{top_geno_y1}` (`{top_val_y1}^{top_let_y1}`) registered the maximum value in `{year1}`, remaining statistically equivalent to `{at_par_y1}`. During `{year2}`, the highest mean belonged to `{top_geno_y2}` (`{top_val_y2}^{top_let_y2}`), which was statistically at par with `{at_par_y2}`. The combined multi-year pooled average confirmed `{pooled_top_geno}` as the leading treatment at `{pooled_top_val}`, while `{pooled_low_geno}` remained significantly inferior at `{pooled_low_val}`.",
    "Analysis of variance for **{parameter}** yielded a `{sig_y1}` `{p_y1}` result in `{year1}` and a `{sig_y2}` `{p_y2}` result in `{year2}`. The highest treatment mean during `{year1}` was obtained from `{top_geno_y1}` (`{top_val_y1}^{top_let_y1}`), which was statistically at par with `{at_par_y1}`. In `{year2}`, `{top_geno_y2}` (`{top_val_y2}^{top_let_y2}`) occupied the peak group, sharing statistical parity with `{at_par_y2}`. The combined response (pooled column) revealed that `{pooled_top_geno}` maintained a dominant cumulative mean of `{pooled_top_val}`, whereas `{pooled_low_geno}` showed the lowest collective average of `{pooled_low_val}`."
]

# --- 10 Multi-Year Non-Significant (NS) Templates ---
COMBINED_NS_TEMPLATES = [
    "Regarding **{parameter}**, treatment differences were **significant** `{p_value_y1}` in `{year1}`, but became **nonsignificant** `{p_value_y2}` during `{year2}`. During the `{year1}` season, `{top_geno_y1}` (`{top_val_y1}^{top_let_y1}`) led the trial, but in `{year2}`, the lack of significance indicated that all `{treatment_plural}` performed statistically at par with each other. The pooled mean over both seasons was `{pooled_val_top}` for `{pooled_top_geno}`.",
    "For **{parameter}**, the genotypic effect was **nonsignificant** `{p_value_y1}` in `{year1}`, but was highly **significant** `{p_value_y2}` in `{year2}`. While all tested entries performed statistically at par during the `{year1}` crop cycle with no statistical separation, the second year `{year2}` allowed for clear differentiation where `{top_geno_y2}` (`{top_val_y2}^{top_let_y2}`) stood out. The combined pooled average across both seasons was recorded at `{pooled_val_top}` for `{pooled_top_geno}`.",
    "The analysis of variance for **{parameter}** proved to be **nonsignificant** `{p_value_y1}` in `{year1}` and remained **nonsignificant** `{p_value_y2}` in `{year2}`. The treatments behaved statistically at par across both crop cycles, indicating consistent trait stability. The cumulative pooled mean across seasons reached a maximum of `{pooled_val_top}` for `{pooled_top_geno}` and a minimum of `{pooled_val_low}` for `{pooled_low_geno}`.",
    "Trait expression of **{parameter}** was **significant** `{p_value_y1}` in `{year1}` but transitioned to a **nonsignificant** `{p_value_y2}` response in `{year2}`. Although `{top_geno_y1}` (`{top_val_y1}^{top_let_y1}`) led `{year1}`, the identical performance of all `{treatment_plural}` in `{year2}` indicates they were statistically at par during that season. Across both years, the pooled data indicated that `{pooled_top_geno}` maintained the highest average of `{pooled_val_top}`.",
    "With respect to **{parameter}**, differences were **nonsignificant** `{p_value_y1}` in the `{year1}` trial but was **significant** `{p_value_y2}` in the `{year2}` trial. The uniform response in `{year1}` left all `{treatment_plural}` statistically at par, while `{year2}` exhibited distinct groupings where `{top_geno_y2}` (`{top_val_y2}^{top_let_y2}`) was dominant. The combined pooled means confirmed `{pooled_top_geno}` as the outstanding entry at `{pooled_val_top}`.",
    "A **significant** `{p_value_y1}` treatment influence was noted on **{parameter}** in `{year1}`, but this effect became **nonsignificant** `{p_value_y2}` in `{year2}`. The second year `{year2}` showed a flat trend where all `{treatment_plural}` remained statistically equivalent, in contrast to the clear stratification observed in `{year1}`. Over both seasons, the pooled value reached `{pooled_val_top}` in the highest-yielding entry `{pooled_top_geno}`.",
    "The treatment effect on **{parameter}** was found to be **nonsignificant** `{p_value_y1}` in the first crop cycle `{year1}`, but was **significant** `{p_value_y2}` in the second cycle `{year2}`. This delayed expression meant all `{treatment_plural}` performed statistically at par during `{year1}`, but separated significantly in `{year2}`. The overall pooled average across the two cycles was `{pooled_val_top}` for `{pooled_top_geno}` compared to `{pooled_val_low}` for `{pooled_low_geno}`.",
    "For **{parameter}**, genotypic variations were **significant** `{p_value_y1}` in `{year1}` but were verified as **nonsignificant** `{p_value_y2}` in `{year2}` due to seasonal environmental dampening. The complete lack of significance in `{year2}` left all `{treatment_plural}` statistically at par, whereas `{year1}` allowed for clear statistical separation. The combined pooled average across both seasons was `{pooled_val_top}` for `{pooled_top_geno}`.",
    "Statistical evaluation of **{parameter}** confirmed that treatment effects were **nonsignificant** `{p_value_y1}` in `{year1}` and **nonsignificant** `{p_value_y2}` in `{year2}`. The absence of significant differences in both crop years indicates that all `{treatment_plural}` co-occupied the same statistical tier throughout the study. Over the combined seasons, `{pooled_top_geno}` achieved the highest pooled value of `{pooled_val_top}`.",
    "Regarding **{parameter}**, the agronomic response was **nonsignificant** `{p_value_y1}` during the `{year1}` cycle but became **significant** `{p_value_y2}` during `{year2}`. The lack of phenotypic stratification in `{year1}` meant all `{treatment_plural}` performed statistically at par, whereas `{year2}` displayed distinct performance hierarchies. The pooled average over both seasons concluded at `{pooled_val_top}` for `{pooled_top_geno}`."
]

# --- APA Style Word Document Formatting Helpers ---
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
    st.markdown("### RCBD Single Factor - 2-Year Combined Analysis")
    mode = st.radio("Choose Input Mode", ["Summarized Table Mode", "Raw Data Mode"], key="2y_mode")
    uploaded_file = st.file_uploader("Upload 2-Year Combined Excel File", type=["xlsx"], key="file_uploader_2y")

    if uploaded_file is not None:
        if mode == "Summarized Table Mode":
            run_summary_mode(uploaded_file)
        else:
            run_raw_mode(uploaded_file)

# --- Summarized Mode ---
def run_summary_mode(uploaded_file):
    try:
        df_raw = pd.read_excel(uploaded_file, header=None)
        
        anchor_row_idx = None
        anchor_keywords = ["genotype", "treatment", "variety", "fertilizer", "sowing", "pesticide", "spacing", "cultivar", "treatments"]
        
        for idx, val in enumerate(df_raw[0]):
            if pd.notna(val) and str(val).strip().lower() in anchor_keywords:
                anchor_row_idx = idx
                break
                
        if anchor_row_idx is None:
            st.error("Could not find the Treatment/Genotype column. Ensure Column A has a header like 'Genotype' or 'Treatment'.")
            return
            
        # Extract Parameter Row (Row 1 of table, index anchor_row_idx - 2)
        param_row = df_raw.iloc[anchor_row_idx - 2].tolist()
        # Forward fill parameter row to handle merged cells
        filled_params = []
        current_param = ""
        for p in param_row:
            if pd.notna(p) and str(p).strip() != "":
                current_param = str(p).strip()
            filled_params.append(current_param)
            
        # Extract Year Row (Row 2 of table, index anchor_row_idx - 1)
        year_row = [str(x).strip() if pd.notna(x) else "" for x in df_raw.iloc[anchor_row_idx - 1]]
        
        # Read df starting from anchor_row_idx
        df_cleaned = pd.read_excel(uploaded_file, skiprows=anchor_row_idx)
        df_cleaned.columns = [str(c).strip() for c in df_cleaned.columns]
        treatment_col = df_cleaned.columns[0]
        
        treatment_label = treatment_col.lower()
        treatment_label_plural = treatment_label[:-1] + "ies" if treatment_label.endswith('y') else treatment_label + "s"
        
        stats_keywords = ["sem", "p-value", "lsd", "cv", "grand mean"]
        stats_rows = {}
        treatment_end_idx = len(df_cleaned)
        
        for idx, val in enumerate(df_cleaned[treatment_col]):
            if pd.notna(val):
                val_clean = str(val).strip().lower()
                for key in stats_keywords:
                    if key in val_clean:
                        stats_rows[key] = idx
                        treatment_end_idx = min(treatment_end_idx, idx)
                        
        df_treatments = df_cleaned.iloc[:treatment_end_idx].copy()
        stats_data = {key: df_cleaned.iloc[row_idx] for key, row_idx in stats_rows.items()}
        
        st.write("#### Preview Treatments:", df_treatments)
        
        # Map parameters from columns in steps of 3
        parameters_unique = []
        for p in filled_params[1:]:
            if p not in parameters_unique and p != "":
                parameters_unique.append(p)
                
        group_1_name = st.text_input("First Table Title", "Vegetative growth and morphological traits of genotypes over two seasons", key="2y_g1_name")
        group_1_cols = st.multiselect("Select parameters for Table 1", parameters_unique, default=parameters_unique[:len(parameters_unique)//2], key="2y_g1_cols")
        group_2_name = st.text_input("Second Table Title", "Yield and yield attributes of genotypes over two seasons", key="2y_g2_name")
        group_2_cols = st.multiselect("Select parameters for Table 2", [c for c in parameters_unique if c not in group_1_cols], default=[c for c in parameters_unique if c not in group_1_cols], key="2y_g2_cols")
        
        if st.button("Generate Word Report (Combined Summary Mode)", key="2y_btn_summary"):
            doc = Document()
            doc.add_heading("Combined Two-Year RCBD Report Draft", 0)
            
            groups = [(group_1_name, group_1_cols, 1), (group_2_name, group_2_cols, 2)]
            comb_idx, ns_idx = 0, 0
            
            for g_title, g_cols, table_num in groups:
                if not g_cols:
                    continue
                doc.add_heading(f"Table {table_num}: {g_title}", level=1)
                text_paragraphs = []
                
                for param in g_cols:
                    # Find indices matching this parameter in the original raw dataframe
                    col_indices = [idx for idx, name in enumerate(filled_params) if name == param]
                    if len(col_indices) < 3:
                        continue
                    
                    y1_idx = col_indices[0]
                    y2_idx = col_indices[1]
                    p_idx = col_indices[2] # Polled
                    
                    y1_name = year_row[y1_idx]
                    y2_name = year_row[y2_idx]
                    
                    # Extract statistics
                    p_val_y1_raw = str(stats_data.get("p-value", {}).iloc[y1_idx])
                    p_val_y2_raw = str(stats_data.get("p-value", {}).iloc[y2_idx])
                    
                    # Year 1 Significance
                    sig_y1 = "significant"
                    p_notation_y1 = "(p < 0.05)"
                    if "NS" in p_val_y1_raw.upper():
                        sig_y1, p_notation_y1 = "nonsignificant", "(p > 0.05)"
                    elif "***" in p_val_y1_raw: p_notation_y1 = "(p < 0.001)"
                    elif "**" in p_val_y1_raw: p_notation_y1 = "(p < 0.01)"
                    
                    # Year 2 Significance
                    sig_y2 = "significant"
                    p_notation_y2 = "(p < 0.05)"
                    if "NS" in p_val_y2_raw.upper():
                        sig_y2, p_notation_y2 = "nonsignificant", "(p > 0.05)"
                    elif "***" in p_val_y2_raw: p_notation_y2 = "(p < 0.001)"
                    elif "**" in p_val_y2_raw: p_notation_y2 = "(p < 0.01)"
                    
                    # Extract values for Year 1, Year 2, and Polled
                    y1_series = df_treatments.iloc[:, y1_idx]
                    y2_series = df_treatments.iloc[:, y2_idx]
                    p_series = df_treatments.iloc[:, p_idx]
                    
                    y1_parsed = [parse_dmrt_value(x) for x in y1_series]
                    y2_parsed = [parse_dmrt_value(x) for x in y2_series]
                    
                    # Parse Year 1 Numeric values for Sorting
                    num_y1 = []
                    for i, (num_str, let) in enumerate(y1_parsed):
                        try: num_y1.append((float(num_str), let, df_treatments.iloc[i, 0]))
                        except ValueError: pass
                    num_y1.sort(reverse=True, key=lambda x: x[0])
                    
                    # Parse Year 2 Numeric values
                    num_y2 = []
                    for i, (num_str, let) in enumerate(y2_parsed):
                        try: num_y2.append((float(num_str), let, df_treatments.iloc[i, 0]))
                        except ValueError: pass
                    num_y2.sort(reverse=True, key=lambda x: x[0])
                    
                    # Parse Polled numeric values
                    num_p = []
                    for i, val in enumerate(p_series):
                        try: num_p.append((float(val), df_treatments.iloc[i, 0]))
                        except ValueError: pass
                    num_p.sort(reverse=True, key=lambda x: x[0])
                    
                    if not num_y1 or not num_y2 or not num_p:
                        continue
                        
                    top_val_y1, top_let_y1, top_geno_y1 = num_y1[0]
                    low_val_y1, low_let_y1, low_geno_y1 = num_y1[-1]
                    top_val_y2, top_let_y2, top_geno_y2 = num_y2[0]
                    low_val_y2, low_let_y2, low_geno_y2 = num_y2[-1]
                    pooled_top_val, pooled_top_geno = num_p[0]
                    pooled_low_val, pooled_low_geno = num_p[-1]
                    other_geno_y1, other_val_y1, other_let_y1 = num_y1[1] if len(num_y1) > 1 else (top_val_y1, top_let_y1, top_geno_y1)
                    other_geno_y2, other_val_y2, other_let_y2 = num_y2[1] if len(num_y2) > 1 else (top_val_y2, top_let_y2, top_geno_y2)
                    
                    # Gather at-par genotypes for both years
                    at_par_y1_list = []
                    for val, let, geno in num_y1[1:]:
                        if top_let_y1 and let and any(char in top_let_y1 for char in let):
                            at_par_y1_list.append(f"{geno} ({val:.2f}^{let})")
                    at_par_y1_str = ", ".join(at_par_y1_list) if at_par_y1_list else "no other treatments"
                    
                    at_par_y2_list = []
                    for val, let, geno in num_y2[1:]:
                        if top_let_y2 and let and any(char in top_let_y2 for char in let):
                            at_par_y2_list.append(f"{geno} ({val:.2f}^{let})")
                    at_par_y2_str = ", ".join(at_par_y2_list) if at_par_y2_list else "no other treatments"
                    
                    # Template selection based on significance
                    if sig_y1 == "nonsignificant" or sig_y2 == "nonsignificant":
                        template = COMBINED_NS_TEMPLATES[ns_idx % len(COMBINED_NS_TEMPLATES)]
                        ns_idx += 1
                    else:
                        template = COMBINED_TEMPLATES[comb_idx % len(COMBINED_TEMPLATES)]
                        comb_idx += 1
                        
                    desc = template.format(
                        parameter=param, year1=y1_name, year2=y2_name,
                        treatment_singular=treatment_label, treatment_plural=treatment_label_plural,
                        sig_y1=sig_y1, p_y1=p_notation_y1,
                        sig_y2=sig_y2, p_y2=p_notation_y2,
                        top_geno_y1=top_geno_y1, top_val_y1=f"{top_val_y1:.2f}", top_let_y1=top_let_y1,
                        at_par_y1=at_par_y1_str,
                        low_geno_y1=low_geno_y1, low_val_y1=f"{low_val_y1:.2f}", low_let_y1=low_let_y1,
                        top_geno_y2=top_geno_y2, top_val_y2=f"{top_val_y2:.2f}", top_let_y2=top_let_y2,
                        at_par_y2=at_par_y2_str,
                        low_geno_y2=low_geno_y2, low_val_y2=f"{low_val_y2:.2f}", low_let_y2=low_let_y2,
                        other_geno_y1=other_geno_y1, other_val_y1=f"{other_val_y1:.2f}", other_let_y1=other_let_y1,
                        other_geno_y2=other_geno_y2, other_val_y2=f"{other_val_y2:.2f}", other_let_y2=other_let_y2,
                        pooled_top_geno=pooled_top_geno, pooled_top_val=f"{pooled_top_val:.2f}",
                        pooled_low_geno=pooled_low_geno, pooled_low_val=f"{pooled_low_val:.2f}"
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
                            
                # Create Multi-level Table (merged header for parameters)
                table = doc.add_table(rows=2, cols=len(g_cols)*3 + 1)
                set_table_borders(table)
                
                # Header row 1 (Parameter names merged over 3 columns)
                table.rows[0].cells[0].text = str(treatment_col)
                set_cell_margins(table.rows[0].cells[0])
                for c_idx, col_name in enumerate(g_cols):
                    start_idx = 1 + c_idx * 3
                    table.rows[0].cells[start_idx].text = col_name
                    set_cell_margins(table.rows[0].cells[start_idx])
                    # Merge cells programmatically
                    table.rows[0].cells[start_idx].merge(table.rows[0].cells[start_idx + 2])
                    
                # Header row 2 (Year 1, Year 2, Polled names)
                table.rows[1].cells[0].text = ""
                set_cell_margins(table.rows[1].cells[0])
                for c_idx, col_name in enumerate(g_cols):
                    col_indices = [idx for idx, name in enumerate(filled_params) if name == col_name]
                    start_idx = 1 + c_idx * 3
                    table.rows[1].cells[start_idx].text = year_row[col_indices[0]]
                    set_cell_margins(table.rows[1].cells[start_idx])
                    table.rows[1].cells[start_idx + 1].text = year_row[col_indices[1]]
                    set_cell_margins(table.rows[1].cells[start_idx + 1])
                    table.rows[1].cells[start_idx + 2].text = year_row[col_indices[2]]
                    set_cell_margins(table.rows[1].cells[start_idx + 2])
                    
                set_header_bottom_border(table.rows[1])
                
                # Fill treatment rows
                for r_idx, t_row in df_treatments.iterrows():
                    row_cells = table.add_row().cells
                    row_cells[0].text = str(t_row[treatment_col])
                    set_cell_margins(row_cells[0])
                    for c_idx, col_name in enumerate(g_cols):
                        col_indices = [idx for idx, name in enumerate(filled_params) if name == col_name]
                        start_idx = 1 + c_idx * 3
                        row_cells[start_idx].text = str(t_row.iloc[col_indices[0]])
                        set_cell_margins(row_cells[start_idx])
                        row_cells[start_idx + 1].text = str(t_row.iloc[col_indices[1]])
                        set_cell_margins(row_cells[start_idx + 1])
                        row_cells[start_idx + 2].text = str(t_row.iloc[col_indices[2]])
                        set_cell_margins(row_cells[start_idx + 2])
                        
                # Add stats rows
                for stat_key in ["sem", "lsd", "cv", "grand mean"]:
                    if stat_key in stats_data:
                        row_cells = table.add_row().cells
                        label_map = {"sem": "SEm (±)", "lsd": "LSD (0.05)", "cv": "CV (%)", "grand mean": "Grand Mean"}
                        row_cells[0].text = label_map.get(stat_key, stat_key.title())
                        set_cell_margins(row_cells[0])
                        for c_idx, col_name in enumerate(g_cols):
                            col_indices = [idx for idx, name in enumerate(filled_params) if name == col_name]
                            start_idx = 1 + c_idx * 3
                            
                            row_cells[start_idx].text = str(stats_data[stat_key].iloc[col_indices[0]])
                            set_cell_margins(row_cells[start_idx])
                            row_cells[start_idx + 1].text = str(stats_data[stat_key].iloc[col_indices[1]])
                            set_cell_margins(row_cells[start_idx + 1])
                            row_cells[start_idx + 2].text = str(stats_data[stat_key].iloc[col_indices[2]]) if len(col_indices) > 2 else "N/A"
                            set_cell_margins(row_cells[start_idx + 2])
                            
                doc.add_page_break()
                
            bio = io.BytesIO()
            doc.save(bio)
            bio.seek(0)
            st.success("🎉 Word Report generated from Summarized 2-Year data!")
            st.download_button("Download Report (.docx)", data=bio, file_name="Summarized_2Year_Report.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="btn_d_2y")
    except Exception as e:
        st.error(f"Error parsing summarized combined table: {e}")

# --- Raw Combined 2-Year Analysis Mode ---
def run_raw_mode(uploaded_file):
    try:
        df_raw_data = pd.read_excel(uploaded_file)
        st.write("#### Preview Raw Combined Input Data:", df_raw_data.head())
        
        cols = df_raw_data.columns.tolist()
        treatment_col = st.selectbox("Select Treatment Column", cols, index=0, key="raw_2y_tr")
        year_col = st.selectbox("Select Year Column", cols, index=1, key="raw_2y_yr")
        block_col = st.selectbox("Select Block/Replication Column", cols, index=2, key="raw_2y_bk")
        response_cols = st.multiselect("Select Response Parameters to Analyze", cols, default=cols[3:], key="raw_2y_resp")
        
        if response_cols:
            df_raw_data[treatment_col] = df_raw_data[treatment_col].astype(str)
            df_raw_data[year_col] = df_raw_data[year_col].astype(str)
            df_raw_data[block_col] = df_raw_data[block_col].astype(str)
            
            # Setup nested block
            df_raw_data['Block_Nested'] = df_raw_data[year_col] + "_" + df_raw_data[block_col]
            
            years_unique = sorted(df_raw_data[year_col].unique().tolist())
            if len(years_unique) != 2:
                st.error("Combined analysis module requires data spanning exactly 2 distinct years.")
                return
                
            y1_name, y2_name = years_unique[0], years_unique[1]
            
            treatment_label = treatment_col.lower()
            treatment_label_plural = treatment_label[:-1] + "ies" if treatment_label.endswith('y') else treatment_label + "s"
            
            group_1_name = st.text_input("First Table Title", "Vegetative growth and morphological traits of genotypes over two seasons", key="raw_2y_g1_name")
            group_1_cols = st.multiselect("Select parameters for Table 1", response_cols, default=response_cols[:len(response_cols)//2], key="raw_2y_g1_cols")
            group_2_name = st.text_input("Second Table Title", "Yield and yield attributes of genotypes over two seasons", key="raw_2y_g2_name")
            group_2_cols = st.multiselect("Select parameters for Table 2", [c for c in response_cols if c not in group_1_cols], default=[c for c in response_cols if c not in group_1_cols], key="raw_2y_g2_cols")
            
            if st.button("Generate Combined Word Report from Raw Data", key="btn_raw_2y_calc"):
                doc = Document()
                doc.add_heading("Calculated Combined 2-Year RCBD Report", 0)
                
                groups = [(group_1_name, group_1_cols, 1), (group_2_name, group_2_cols, 2)]
                comb_idx, ns_idx = 0, 0
                
                for g_title, g_cols, table_num in groups:
                    if not g_cols:
                        continue
                    doc.add_heading(f"Table {table_num}: {g_title}", level=1)
                    text_paragraphs = []
                    
                    # Storage structures for table construction
                    calculated_means_for_table = {}  # {Treatment: {col_year: "val_letters"}}
                    calculated_stats_for_table = {k: {} for k in ["sem", "lsd", "cv", "grand mean"]} # {stat_type: {col_year: val}}
                    
                    for param in g_cols:
                        # 1. Combined ANOVA to verify interaction significance
                        formula_combined = f"Q('{param}') ~ C(Q('{year_col}')) + C(Block_Nested) + C(Q('{treatment_col}')) + C(Q('{treatment_col}')):C(Q('{year_col}'))"
                        model_comb = ols(formula_combined, data=df_raw_data).fit()
                        anova_comb = sm.stats.anova_lm(model_comb, typ=1)
                        
                        p_tr_year = anova_comb.loc[f"C(Q('{treatment_col}')):C(Q('{year_col}'))", "PR(>F)"]
                        
                        # 2. Year-wise stats calculations (Run ANOVA separately per year)
                        df_y1 = df_raw_data[df_raw_data[year_col] == y1_name]
                        df_y2 = df_raw_data[df_raw_data[year_col] == y2_name]
                        
                        # Year 1 Stats
                        formula_y1 = f"Q('{param}') ~ C(Q('{block_col}')) + C(Q('{treatment_col}'))"
                        model_y1 = ols(formula_y1, data=df_y1).fit()
                        anova_y1 = sm.stats.anova_lm(model_y1, typ=1)
                        p_val_y1 = anova_y1.loc[f"C(Q('{treatment_col}'))", "PR(>F)"]
                        df_err_y1 = anova_y1.loc["Residual", "df"]
                        mse_y1 = anova_y1.loc["Residual", "mean_sq"]
                        r1 = df_y1[block_col].nunique()
                        grand_mean_y1 = df_y1[param].mean()
                        sem_y1 = np.sqrt(mse_y1 / r1)
                        cv_y1 = (np.sqrt(mse_y1) / grand_mean_y1) * 100
                        t_val_y1 = t.ppf(0.975, df_err_y1)
                        lsd_y1 = t_val_y1 * np.sqrt((2 * mse_y1) / r1)
                        
                        means_y1 = df_y1.groupby(treatment_col)[param].mean().to_dict()
                        cld_y1 = get_cld_letters(means_y1, lsd_y1)
                        
                        # Year 2 Stats
                        formula_y2 = f"Q('{param}') ~ C(Q('{block_col}')) + C(Q('{treatment_col}'))"
                        model_y2 = ols(formula_y2, data=df_y2).fit()
                        anova_y2 = sm.stats.anova_lm(model_y2, typ=1)
                        p_val_y2 = anova_y2.loc[f"C(Q('{treatment_col}'))", "PR(>F)"]
                        df_err_y2 = anova_y2.loc["Residual", "df"]
                        mse_y2 = anova_y2.loc["Residual", "mean_sq"]
                        r2 = df_y2[block_col].nunique()
                        grand_mean_y2 = df_y2[param].mean()
                        sem_y2 = np.sqrt(mse_y2 / r2)
                        cv_y2 = (np.sqrt(mse_y2) / grand_mean_y2) * 100
                        t_val_y2 = t.ppf(0.975, df_err_y2)
                        lsd_y2 = t_val_y2 * np.sqrt((2 * mse_y2) / r2)
                        
                        means_y2 = df_y2.groupby(treatment_col)[param].mean().to_dict()
                        cld_y2 = get_cld_letters(means_y2, lsd_y2)
                        
                        # Combined/Pooled Raw Averages
                        means_pooled = df_raw_data.groupby(treatment_col)[param].mean().to_dict()
                        
                        # Populate storage structure for table
                        all_treatments = sorted(df_raw_data[treatment_col].unique().tolist())
                        for t_name in all_treatments:
                            if t_name not in calculated_means_for_table:
                                calculated_means_for_table[t_name] = {}
                            
                            calculated_means_for_table[t_name][f"{param}_{y1_name}"] = f"{means_y1.get(t_name, 0.0):.2f}{cld_y1.get(t_name, '')}"
                            calculated_means_for_table[t_name][f"{param}_{y2_name}"] = f"{means_y2.get(t_name, 0.0):.2f}{cld_y2.get(t_name, '')}"
                            calculated_means_for_table[t_name][f"{param}_Polled"] = f"{means_pooled.get(t_name, 0.0):.2f}"
                            
                        # Save stats rows
                        calculated_stats_for_table["sem"][f"{param}_{y1_name}"] = f"{sem_y1:.2f}"
                        calculated_stats_for_table["sem"][f"{param}_{y2_name}"] = f"{sem_y2:.2f}"
                        calculated_stats_for_table["sem"][f"{param}_Polled"] = ""
                        
                        calculated_stats_for_table["lsd"][f"{param}_{y1_name}"] = f"{lsd_y1:.2f}"
                        calculated_stats_for_table["lsd"][f"{param}_{y2_name}"] = f"{lsd_y2:.2f}"
                        calculated_stats_for_table["lsd"][f"{param}_Polled"] = ""
                        
                        calculated_stats_for_table["cv"][f"{param}_{y1_name}"] = f"{cv_y1:.2f}"
                        calculated_stats_for_table["cv"][f"{param}_{y2_name}"] = f"{cv_y2:.2f}"
                        calculated_stats_for_table["cv"][f"{param}_Polled"] = ""
                        
                        calculated_stats_for_table["grand mean"][f"{param}_{y1_name}"] = f"{grand_mean_y1:.2f}"
                        calculated_stats_for_table["grand mean"][f"{param}_{y2_name}"] = f"{grand_mean_y2:.2f}"
                        calculated_stats_for_table["grand mean"][f"{param}_Polled"] = f"{df_raw_data[param].mean():.2f}"
                        
                        # Map Significance texts
                        sig_y1, p_notation_y1 = "significant", "(p < 0.05)"
                        if p_val_y1 > 0.05: sig_y1, p_notation_y1 = "nonsignificant", "(p > 0.05)"
                        elif p_val_y1 < 0.001: p_notation_y1 = "(p < 0.001)"
                        elif p_val_y1 < 0.01: p_notation_y1 = "(p < 0.01)"
                        
                        sig_y2, p_notation_y2 = "significant", "(p < 0.05)"
                        if p_val_y2 > 0.05: sig_y2, p_notation_y2 = "nonsignificant", "(p > 0.05)"
                        elif p_val_y2 < 0.001: p_notation_y2 = "(p < 0.001)"
                        elif p_val_y2 < 0.01: p_notation_y2 = "(p < 0.01)"
                        
                        # Sort arrays
                        sorted_y1 = sorted(means_y1.items(), key=lambda x: x[1], reverse=True)
                        top_val_y1, top_geno_y1 = sorted_y1[0]
                        top_let_y1 = cld_y1[top_geno_y1]
                        low_val_y1, low_geno_y1 = sorted_y1[-1]
                        other_geno_y1, other_val_y1 = sorted_y1[1] if len(sorted_y1) > 1 else (top_geno_y1, top_val_y1)
                        
                        sorted_y2 = sorted(means_y2.items(), key=lambda x: x[1], reverse=True)
                        top_val_y2, top_geno_y2 = sorted_y2[0]
                        top_let_y2 = cld_y2[top_geno_y2]
                        low_val_y2, low_geno_y2 = sorted_y2[-1]
                        other_geno_y2, other_val_y2 = sorted_y2[1] if len(sorted_y2) > 1 else (top_geno_y2, top_val_y2)
                        
                        sorted_p = sorted(means_pooled.items(), key=lambda x: x[1], reverse=True)
                        pooled_top_val, pooled_top_geno = sorted_p[0]
                        pooled_low_val, pooled_low_geno = sorted_p[-1]
                        
                        # Generate "at par" strings
                        at_par_y1_list = []
                        for geno, val in sorted_y1[1:]:
                            let = cld_y1[geno]
                            if top_let_y1 and let and any(char in top_let_y1 for char in let):
                                at_par_y1_list.append(f"{geno} ({val:.2f}^{let})")
                        at_par_y1_str = ", ".join(at_par_y1_list) if at_par_y1_list else "no other treatments"
                        
                        at_par_y2_list = []
                        for geno, val in sorted_y2[1:]:
                            let = cld_y2[geno]
                            if top_let_y2 and let and any(char in top_let_y2 for char in let):
                                at_par_y2_list.append(f"{geno} ({val:.2f}^{let})")
                        at_par_y2_str = ", ".join(at_par_y2_list) if at_par_y2_list else "no other treatments"
                        
                        if sig_y1 == "nonsignificant" or sig_y2 == "nonsignificant":
                            template = COMBINED_NS_TEMPLATES[ns_idx % len(COMBINED_NS_TEMPLATES)]
                            ns_idx += 1
                        else:
                            template = COMBINED_TEMPLATES[comb_idx % len(COMBINED_TEMPLATES)]
                            comb_idx += 1
                            
                        desc = template.format(
                            parameter=param, year1=y1_name, year2=y2_name,
                            treatment_singular=treatment_label, treatment_plural=treatment_label_plural,
                            sig_y1=sig_y1, p_y1=p_notation_y1,
                            sig_y2=sig_y2, p_y2=p_notation_y2,
                            top_geno_y1=top_geno_y1, top_val_y1=f"{top_val_y1:.2f}", top_let_y1=top_let_y1,
                            at_par_y1=at_par_y1_str,
                            low_geno_y1=low_geno_y1, low_val_y1=f"{low_val_y1:.2f}", low_let_y1=cld_y1.get(low_geno_y1, ""),
                            top_geno_y2=top_geno_y2, top_val_y2=f"{top_val_y2:.2f}", top_let_y2=top_let_y2,
                            at_par_y2=at_par_y2_str,
                            low_geno_y2=low_geno_y2, low_val_y2=f"{low_val_y2:.2f}", low_let_y2=cld_y2.get(low_geno_y2, ""),
                            other_geno_y1=other_geno_y1, other_val_y1=f"{other_val_y1:.2f}", other_let_y1=cld_y1.get(other_geno_y1, ""),
                            other_geno_y2=other_geno_y2, other_val_y2=f"{other_val_y2:.2f}", other_let_y2=cld_y2.get(other_geno_y2, ""),
                            pooled_top_geno=pooled_top_geno, pooled_top_val=f"{pooled_top_val:.2f}",
                            pooled_low_geno=pooled_low_geno, pooled_low_val=f"{pooled_low_val:.2f}"
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
                                
                    # Word Table Generation
                    table = doc.add_table(rows=2, cols=len(g_cols)*3 + 1)
                    set_table_borders(table)
                    
                    # Row 1 Parameters merged
                    table.rows[0].cells[0].text = str(treatment_col)
                    set_cell_margins(table.rows[0].cells[0])
                    for c_idx, col_name in enumerate(g_cols):
                        start_idx = 1 + c_idx * 3
                        table.rows[0].cells[start_idx].text = col_name
                        set_cell_margins(table.rows[0].cells[start_idx])
                        table.rows[0].cells[start_idx].merge(table.rows[0].cells[start_idx+2])
                        
                    # Row 2 Years
                    table.rows[1].cells[0].text = ""
                    set_cell_margins(table.rows[1].cells[0])
                    for c_idx, col_name in enumerate(g_cols):
                        start_idx = 1 + c_idx * 3
                        table.rows[1].cells[start_idx].text = y1_name
                        set_cell_margins(table.rows[1].cells[start_idx])
                        table.rows[1].cells[start_idx+1].text = y2_name
                        set_cell_margins(table.rows[1].cells[start_idx+1])
                        table.rows[1].cells[start_idx+2].text = "Polled"
                        set_cell_margins(table.rows[1].cells[start_idx+2])
                        
                    set_header_bottom_border(table.rows[1])
                    
                    # Fill treatment means
                    for t_name in sorted(calculated_means_for_table.keys()):
                        row_cells = table.add_row().cells
                        row_cells[0].text = str(t_name)
                        set_cell_margins(row_cells[0])
                        for c_idx, col_name in enumerate(g_cols):
                            start_idx = 1 + c_idx * 3
                            row_cells[start_idx].text = calculated_means_for_table[t_name].get(f"{col_name}_{y1_name}", "N/A")
                            set_cell_margins(row_cells[start_idx])
                            row_cells[start_idx+1].text = calculated_means_for_table[t_name].get(f"{col_name}_{y2_name}", "N/A")
                            set_cell_margins(row_cells[start_idx+1])
                            row_cells[start_idx+2].text = calculated_means_for_table[t_name].get(f"{col_name}_Polled", "N/A")
                            set_cell_margins(row_cells[start_idx+2])
                            
                    # Add stats
                    for stat_key in ["sem", "lsd", "cv", "grand mean"]:
                        row_cells = table.add_row().cells
                        label_map = {"sem": "SEm (±)", "lsd": "LSD (0.05)", "cv": "CV (%)", "grand mean": "Grand Mean"}
                        row_cells[0].text = label_map.get(stat_key, stat_key.title())
                        set_cell_margins(row_cells[0])
                        for c_idx, col_name in enumerate(g_cols):
                            start_idx = 1 + c_idx * 3
                            row_cells[start_idx].text = calculated_stats_for_table[stat_key].get(f"{col_name}_{y1_name}", "N/A")
                            set_cell_margins(row_cells[start_idx])
                            row_cells[start_idx+1].text = calculated_stats_for_table[stat_key].get(f"{col_name}_{y2_name}", "N/A")
                            set_cell_margins(row_cells[start_idx+1])
                            row_cells[start_idx+2].text = calculated_stats_for_table[stat_key].get(f"{col_name}_Polled", "N/A")
                            set_cell_margins(row_cells[start_idx+2])
                            
                    doc.add_page_break()
                    
                bio = io.BytesIO()
                doc.save(bio)
                bio.seek(0)
                st.success("🎉 Word Report successfully built from combined raw data!")
                st.download_button("Download Calculated Combined Report (.docx)", data=bio, file_name="Calculated_Combined_Report.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document", key="btn_d_rcbd_raw_2y")
    except Exception as e:
        st.error(f"Error executing raw combined RCBD analysis: {e}")
