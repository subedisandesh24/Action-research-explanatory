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

# --- 15 Significant and 5 Nonsignificant Templates ---
SIGNIFICANT_TEMPLATES = [
    "The analysis of variance for **{parameter}** revealed a `{significance}` `{p_value}` treatment effect during the trial. The maximum value was registered by `{top_geno}` (`{top_val}^{top_let}`), which was statistically at par with `{at_par_genotypes}`, whereas `{low_geno}` (`{low_val}^{low_let}`) marked the lowest performance.",
    "Regarding **{parameter}**, `{significance}` `{p_value}` phenotypic differentiation was observed among the evaluated `{treatment_plural}`. Genotype `{top_geno}` (`{top_val}^{top_let}`) dominated the ranking, sharing statistical parity with `{at_par_genotypes}`. On the other end of the performance spectrum, `{low_geno}` (`{low_val}^{low_let}`) occupied the lowest statistical tier, demonstrating a pronounced gap compared to the top-performing group.",
    "A `{significance}` `{p_value}` genotypic influence was noted on the agronomic performance of **{parameter}**. `{top_geno}` achieved the apex value of `{top_val}^{top_let}`, remaining statistically equivalent to `{at_par_genotypes}`, while `{low_geno}` (`{low_val}^{low_let}`) was significantly inferior.",
    "The performance hierarchy for **{parameter}** proved to be `{significance}` `{p_value}` across the germplasm. `{top_geno}` (`{top_val}^{top_let}`) stood at the peak of the statistical distribution, showing no significant difference from `{at_par_genotypes}`. This group of high-performing entries outclassed `{low_geno}` (`{low_val}^{low_let}`), which represented the lower limit of performance relative to the grand mean of `{grand_mean}`.",
    "Mean separation of **{parameter}** data revealed `{significance}` `{p_value}` variation among the evaluated `{treatment_plural}`. The upper boundary was occupied by `{top_geno}` (`{top_val}^{top_let}`), which co-occupied the top statistical tier along with `{at_par_genotypes}`. In contrast, `{low_geno}` (`{low_val}^{low_let}`) was isolated at the minimum threshold, indicating a significant reduction in trait expression.",
    "The evaluated `{treatment_plural}` clustered into distinct performance tiers for **{parameter}**, exhibiting a `{significance}` `{p_value}` variance. The highest cluster was led by `{top_geno}` (`{top_val}^{top_let}`), showing statistical parity with `{at_par_genotypes}`. This superior group maintained a substantial margin over the lowest-ranked `{treatment_singular}`, `{low_geno}` (`{low_val}^{low_let}`).",
    "Trait characterization for **{parameter}** indicated a `{significance}` `{p_value}` treatment influence during the trial. `{top_geno}` (`{top_val}^{top_let}`) emerged as the outstanding performer, remaining statistically at par with `{at_par_genotypes}`. The lowest phenotypic expression was observed in `{low_geno}` (`{low_val}^{low_let}`), which failed to cross the general grand mean of `{grand_mean}`.",
    "Regarding **{parameter}**, the treatments exhibited a `{significance}` `{p_value}` response pattern. Genotype `{top_geno}` (`{top_val}^{top_let}`) registered the maximum mean value, showing statistical parity with `{at_par_genotypes}`. This group of high-performing treatments displayed a significant advantage over `{low_geno}` (`{low_val}^{low_let}`), which represented the absolute minimum in this trial.",
    "The statistical analysis of **{parameter}** showed `{significance}` `{p_value}` stratification among the evaluated `{treatment_plural}`. The highest performing bracket was defined by `{top_geno}` (`{top_val}^{top_let}`), which was statistically equivalent to `{at_par_genotypes}`. The baseline of this parameter was occupied by `{low_geno}` (`{low_val}^{low_let}`), highlighting a significant gap between the extremes.",
    "ANOVA results for **{parameter}** were verified as `{significance}` `{p_value}`. DMRT grouping assigned the supreme letter '{top_let}' to `{top_geno}` (`{top_val}^{top_let}`), which shared statistical letters with `{at_par_genotypes}`. The lowest performer, `{low_geno}` (`{low_val}^{low_let}`), was significantly outclassed by this leading group.",
    "For **{parameter}**, the genotypic effect was `{significance}` `{p_value}`. `{top_geno}` (`{top_val}^{top_let}`) established the upper vigor threshold, remaining statistically at par with `{at_par_genotypes}`. On the other hand, `{low_geno}` (`{low_val}^{low_let}`) represented the minimum limit of performance, demonstrating a significant drop.",
    "The magnitude of **{parameter}** manifestation was `{significance}` `{p_value}` across the evaluated `{treatment_plural}`. The highest value was observed in `{top_geno}` (`{top_val}^{top_let}`), which was statistically at par with `{at_par_genotypes}`. The lowest performance was restricted to `{low_geno}` (`{low_val}^{low_let}`), which fell significantly short of the trial's grand mean of `{grand_mean}`.",
    "A `{significance}` `{p_value}` treatment influence was recorded on **{parameter}** during the cropping cycle. `{top_geno}` (`{top_val}^{top_let}`) outperformed `{low_geno}` (`{low_val}^{low_let}`) by a statistically significant margin, with `{top_geno}` remaining statistically at par with `{at_par_genotypes}`.",
    "The evaluated `{treatment_plural}` displayed `{significance}` `{p_value}` differences in phenotypic vigor for **{parameter}**. Genotype `{top_geno}` (`{top_val}^{top_let}`) led the rankings, showing no significant difference from `{at_par_genotypes}`. The baseline response was recorded in `{low_geno}` (`{low_val}^{low_let}`), which marked the lowest limit of performance.",
    "The performance profile for **{parameter}** proved to be `{significance}` `{p_value}`. `{top_geno}` (`{top_val}^{top_let}`) stood out as the leading treatment, showing statistical parity with `{at_par_genotypes}`. Compared to this superior group, `{low_geno}` (`{low_val}^{low_let}`) exhibited a significant reduction in value."
]

NONSIGNIFICANT_TEMPLATES = [
    "For **{parameter}**, the genotypic influence was found to be **nonsignificant** `{p_value}`. No statistical differences were detected among the evaluated `{treatment_plural}`, indicating that all treatments performed statistically at par under the experimental conditions, with the mean values remaining relatively close to the trial's grand mean of `{grand_mean}`.",
    "Regarding **{parameter}**, the treatment effect was **nonsignificant** `{p_value}`, demonstrating highly stable and uniform trait expression across the germplasm. All evaluated `{treatment_plural}` remained statistically equivalent, with no individual treatment separating out as superior or inferior.",
    "Statistical analysis of **{parameter}** confirmed that treatment effects were **nonsignificant** `{p_value}`. This indicates that all tested `{treatment_plural}` co-occupied the same statistical tier and possessed equivalent performance potential under the single-year trial environment.",
    "The agronomic response for **{parameter}** exhibited a **nonsignificant** `{p_value}` trend during the cropping cycle. No phenotypic stratification occurred among the `{treatment_plural}`, leaving all entries statistically at par with a flat distribution of values around the grand mean of `{grand_mean}`.",
    "Analysis of variance for **{parameter}** yielded a **nonsignificant** `{p_value}` result, pointing to highly homogenous data. All evaluated `{treatment_plural}` behaved identically from a statistical standpoint, with no significant differences observed between the highest and lowest recorded values."
]

# --- Formatting Helpers ---
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

# --- Main Module UI/Execution Entry ---
def show_module():
    st.markdown("### RCBD Single Factor - One Year Analysis")
    mode = st.radio("Choose Input Mode", ["Summarized Table Mode", "Raw Data Mode"])

    uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"], key="file_uploader_1y")

    if uploaded_file is not None:
        if mode == "Summarized Table Mode":
            run_summary_mode(uploaded_file)
        else:
            run_raw_mode(uploaded_file)

# --- Execution Core for Summary Table ---
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
        all_parameters = [col for col in df_cleaned.columns if col != treatment_col]
        
        group_1_name = st.text_input("First Table Title", "Vegetative and Morphological Traits")
        group_1_cols = st.multiselect("Select parameters for Table 1", all_parameters, default=all_parameters[:len(all_parameters)//2])
        group_2_name = st.text_input("Second Table Title", "Yield and Yield Components")
        group_2_cols = st.multiselect("Select parameters for Table 2", [c for c in all_parameters if c not in group_1_cols], default=[c for c in all_parameters if c not in group_1_cols])
        
        if st.button("Generate Word Report (Summary Mode)"):
            doc = Document()
            doc.add_heading("Summarized RCBD Report Draft", 0)
            
            groups = [(group_1_name, group_1_cols, 1), (group_2_name, group_2_cols, 2)]
            sig_idx, ns_idx = 0, 0
            
            for g_title, g_cols, table_num in groups:
                if not g_cols:
