import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import RGBColor, Pt
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import io
import os
import re

# Page config MUST be first
st.set_page_config(
    page_title="Gene Coverage Analyzer",
    page_icon="🧬",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Minimal CSS - only essential styles
st.markdown("""
    <style>
    .stButton>button {
        width: 100%;
    }
    </style>
""", unsafe_allow_html=True)

# Initialize session state ONCE
if 'coverage_data' not in st.session_state:
    st.session_state.coverage_data = None
if 'panel_genes' not in st.session_state:
    st.session_state.panel_genes = []
if 'filtered_data' not in st.session_state:
    st.session_state.filtered_data = None
if 'processed_csv' not in st.session_state:
    st.session_state.processed_csv = None
if 'file_basename' not in st.session_state:
    st.session_state.file_basename = "output"
if 'last_processed_file' not in st.session_state:
    st.session_state.last_processed_file = None

@st.cache_data
def extract_gene_name(name):
    """Extract the first gene name from comma or semicolon-separated names."""
    if isinstance(name, str):
        if ',' in name:
            return name.split(',')[0]
        elif ';' in name:
            return name.split(';')[0]
        return name
    return None

def parse_gene_list(text):
    """Parse gene list with support for newlines, commas, and spaces. Remove duplicates."""
    if not text.strip():
        return []
    
    # Replace newlines and commas with spaces
    text = text.replace('\n', ' ').replace(',', ' ').replace('\t', ' ')
    
    # Split by whitespace and filter empty strings
    genes = [gene.strip() for gene in text.split() if gene.strip()]
    
    # Remove duplicates while preserving order
    seen = set()
    unique_genes = []
    for gene in genes:
        if gene not in seen:
            seen.add(gene)
            unique_genes.append(gene)
    
    return unique_genes

@st.cache_data(show_spinner=False)
def preprocess_excel_cached(file_bytes, file_name, skip_rows=1):
    """Cached preprocessing to avoid re-computation"""
    try:
        data = pd.read_excel(io.BytesIO(file_bytes), skiprows=skip_rows)
        
        if 'Gene Name' not in data.columns:
            raise ValueError("Column 'Gene Name' not found")
        
        data['Gene_Name'] = data['Gene Name'].apply(extract_gene_name)
        
        result = []
        identifiers = data[['Gene Names', 'Aliases', 'Gene_Name']].fillna('')
        unique_ids = identifiers.drop_duplicates()
        
        for _, row in unique_ids.iterrows():
            gene = row['Gene Names']
            alias = row['Aliases']
            name = row['Gene_Name']
            
            if gene:
                gene_data = data[data['Gene Names'] == gene]
            elif alias:
                gene_data = data[data['Aliases'] == alias]
            elif name:
                gene_data = data[data['Name'] == name]
            else:
                continue
            
            total_counted_bases = gene_data['Counted Bases'].sum()
            if total_counted_bases == 0:
                continue
            
            mean_depth = (gene_data['Mean Depth'] * gene_data['Counted Bases']).sum() / total_counted_bases
            mean_1x = (gene_data['% 1x'] * gene_data['Counted Bases']).sum() / total_counted_bases
            
            summary = {
                'Region': 'total',
                'Ref Name': gene if gene else None,
                'Aliases': alias if alias else None,
                'Gene_Name': name if name else None,
                'Name': gene_data['Name'].iloc[0] if 'Name' in gene_data.columns else None,
                'Gene IDs': gene_data['Gene IDs'].iloc[0] if 'Gene IDs' in gene_data.columns else None,
                'Counted Bases': total_counted_bases,
                'Mean Depth': mean_depth,
                'Min Depth': gene_data['Min Depth'].min(),
                'Max Depth': gene_data['Max Depth'].max(),
                '% 1x': mean_1x,
            }
            result.append(summary)
        
        coverage_df = pd.DataFrame(result)
        base_name = os.path.splitext(file_name)[0]
        
        return coverage_df, base_name
    except Exception as e:
        raise Exception(f"Error: {str(e)}")

def create_word_document(df_grouped, output_filename):
    """Create Word document with gene coverage table"""
    doc = Document()
    doc.add_heading('Appendix 1: Gene Coverage', 1)
    doc.add_heading('Indication Based Analysis:', 2)
    
    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_after = Pt(0)
    
    chunk_size = 4
    chunks = [df_grouped[i:i+chunk_size] for i in range(0, len(df_grouped), chunk_size)]
    
    table = doc.add_table(rows=1, cols=chunk_size*2)
    table.alignment = 1
    hdr_cells = table.rows[0].cells
    
    for i in range(chunk_size):
        gene_hdr_cell = hdr_cells[i*2]
        gene_hdr_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        gene_hdr_para = gene_hdr_cell.paragraphs[0]
        gene_hdr_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        gene_hdr_para.paragraph_format.space_after = Pt(0)
        gene_hdr_run = gene_hdr_para.add_run("Gene Name")
        gene_hdr_run.font.name = 'Calibri'
        gene_hdr_run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Calibri')
        gene_hdr_run.font.size = Pt(8)
        
        perc_hdr_cell = hdr_cells[i*2+1]
        perc_hdr_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        perc_hdr_para = perc_hdr_cell.paragraphs[0]
        perc_hdr_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        perc_hdr_para.paragraph_format.space_after = Pt(0)
        perc_hdr_run = perc_hdr_para.add_run("Percentage of coding region covered")
        perc_hdr_run.font.name = 'Calibri'
        perc_hdr_run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Calibri')
        perc_hdr_run.font.size = Pt(8)
    
    for chunk in chunks:
        row_cells = table.add_row().cells
        for i in range(chunk_size):
            if i < len(chunk):
                row = chunk.iloc[i]
                gene = str(row['Gene_ID'])
                percent = row['Perc_1x']
                
                gene_cell = row_cells[i*2]
                gene_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                gene_para = gene_cell.paragraphs[0]
                gene_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                gene_para.paragraph_format.space_after = Pt(0)
                gene_run = gene_para.add_run(gene)
                gene_run.font.name = 'Calibri'
                gene_run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Calibri')
                gene_run.font.size = Pt(8)
                gene_run.italic = True
                if percent < 90:
                    gene_run.font.color.rgb = RGBColor(255, 0, 0)
                
                perc_cell = row_cells[i*2+1]
                perc_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                perc_para = perc_cell.paragraphs[0]
                perc_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                perc_para.paragraph_format.space_after = Pt(0)
                perc_run = perc_para.add_run(str(percent))
                perc_run.font.name = 'Calibri'
                perc_run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Calibri')
                perc_run.font.size = Pt(8)
                if percent < 90:
                    perc_run.font.color.rgb = RGBColor(255, 0, 0)
            else:
                for empty_cell in [row_cells[i*2], row_cells[i*2+1]]:
                    empty_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                    para = empty_cell.paragraphs[0]
                    para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                    para.paragraph_format.space_after = Pt(0)
                    run = para.add_run("–")
                    run.font.name = 'Calibri'
                    run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Calibri')
                    run.font.size = Pt(8)
                    run.font.color.rgb = RGBColor(255, 255, 255)
    
    tbl = table._tbl
    tblPr = tbl.tblPr
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.append(tblPr)
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '000000')
        tblBorders.append(border)
    tblPr.append(tblBorders)
    
    return doc

# Main UI
st.title("🧬 Gene Coverage Analyzer")
st.caption("Process raw Excel data or upload pre-processed CSV to generate Word report")

# Step 1
input_type = st.radio(
    "**Step 1: Choose Input Type**",
    ["📊 Pre-processed CSV", "📁 Raw Excel File"],
    horizontal=True
)

# Step 2
col1, col2 = st.columns(2)

with col1:
    if "Raw Excel" in input_type:
        st.subheader("Step 2A: Upload Raw Excel")
        raw_excel = st.file_uploader("Choose Excel file", type=['xlsx', 'xls'], key='raw_excel')
        
        if raw_excel:
            file_id = f"{raw_excel.name}_{raw_excel.size}"
            
            if st.session_state.last_processed_file != file_id:
                progress = st.progress(0, "Reading file...")
                
                try:
                    progress.progress(30, "Processing...")
                    file_bytes = raw_excel.read()
                    coverage_df, basename = preprocess_excel_cached(file_bytes, raw_excel.name)
                    
                    progress.progress(80, "Generating CSV...")
                    st.session_state.coverage_data = coverage_df
                    st.session_state.file_basename = basename
                    st.session_state.processed_csv = coverage_df.to_csv(index=False)
                    st.session_state.last_processed_file = file_id
                    
                    progress.progress(100, "Complete!")
                    progress.empty()
                except Exception as e:
                    progress.empty()
                    st.error(f"❌ {str(e)}")
            
            if st.session_state.processed_csv:
                st.success(f"✅ {len(st.session_state.coverage_data)} records")
                st.download_button(
                    "⬇️ Download CSV",
                    data=st.session_state.processed_csv,
                    file_name=f"{st.session_state.file_basename}_coverage.csv",
                    mime="text/csv",
                    use_container_width=True
                )
    else:
        st.subheader("Step 2A: Upload CSV")
        coverage_file = st.file_uploader("Choose CSV file", type=['csv'], key='coverage')
        
        if coverage_file:
            st.session_state.file_basename = os.path.splitext(coverage_file.name)[0]
            df = pd.read_csv(coverage_file)
            df.columns = df.columns.str.strip()
            
            coverage_col = next((col for col in ['% 1x', '%1x', '% 1x '] if col in df.columns), None)
            
            if coverage_col:
                st.session_state.coverage_data = df
                st.success(f"✅ {len(df)} records")
            else:
                st.error("❌ No coverage column found")

with col2:
    st.subheader("Step 2B: Panel Genes")
    tab1, tab2 = st.tabs(["📁 Excel", "📝 Paste"])
    
    with tab1:
        panel_file = st.file_uploader("Choose Excel", type=['xlsx', 'xls'], key='panel')
        if panel_file:
            panel_df = pd.read_excel(panel_file)
            if 'GENE' in panel_df.columns:
                genes = panel_df['GENE'].dropna().unique().tolist()
                st.session_state.panel_genes = genes
                st.success(f"✅ {len(genes)} genes")
            else:
                st.error("❌ No GENE column")
    
    with tab2:
        gene_text = st.text_area(
            "Paste genes (newline, comma, or space separated)",
            height=150,
            placeholder="BRCA1 BRCA2 TP53\nEGFR, KRAS\nTP53"
        )
        
        if st.button("Load Genes", use_container_width=True):
            genes = parse_gene_list(gene_text)
            if genes:
                st.session_state.panel_genes = genes
                duplicates_removed = len(gene_text.split()) - len(genes)
                if duplicates_removed > 0:
                    st.success(f"✅ {len(genes)} unique genes ({duplicates_removed} duplicates removed)")
                else:
                    st.success(f"✅ {len(genes)} genes")
            else:
                st.warning("⚠️ No genes found")

# Process data
if st.session_state.coverage_data is not None and st.session_state.panel_genes:
    df = st.session_state.coverage_data
    panel_genes = st.session_state.panel_genes
    
    coverage_col = next((col for col in ['% 1x', '%1x', '% 1x '] if col in df.columns), None)
    if not coverage_col:
        st.error("Coverage column not found")
        st.stop()
    
    df_filtered = df[df['Gene_Name'].isin(panel_genes) | df['Ref Name'].isin(panel_genes)].copy()
    df_filtered['Gene_ID'] = df_filtered['Gene_Name'].fillna(df_filtered['Ref Name'])
    df_filtered = df_filtered[~df_filtered['Gene_ID'].fillna('').str.startswith("Intron:")]
    df_filtered['Perc_1x'] = pd.to_numeric(df_filtered[coverage_col], errors='coerce')
    
    df_grouped = df_filtered.groupby('Gene_ID', as_index=False)['Perc_1x'].mean()
    df_grouped['Perc_1x'] = df_grouped['Perc_1x'].round(2)
    df_grouped = df_grouped.sort_values('Gene_ID')
    
    st.session_state.filtered_data = df_grouped
    
    st.divider()
    
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Genes", len(df_grouped))
    with col2:
        st.metric("Low Coverage", len(df_grouped[df_grouped['Perc_1x'] < 90]))
    with col3:
        st.metric("Avg Coverage", f"{df_grouped['Perc_1x'].mean():.2f}%")
    
    st.divider()
    
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("📄 Generate Word", type="primary", use_container_width=True):
            with st.spinner("Creating..."):
                try:
                    word_filename = f"{st.session_state.file_basename}_report.docx"
                    doc = create_word_document(df_grouped, word_filename)
                    doc_bytes = io.BytesIO()
                    doc.save(doc_bytes)
                    doc_bytes.seek(0)
                    st.session_state.doc_bytes = doc_bytes
                    st.session_state.word_filename = word_filename
                    st.success("✅ Ready!")
                except Exception as e:
                    st.error(f"❌ {str(e)}")
        
        if 'doc_bytes' in st.session_state:
            st.download_button(
                f"⬇️ {st.session_state.word_filename}",
                data=st.session_state.doc_bytes,
                file_name=st.session_state.word_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                type="primary"
            )
    
    with col2:
        csv_filename = f"{st.session_state.file_basename}_filtered.csv"
        st.download_button(
            f"📊 {csv_filename}",
            data=df_grouped.to_csv(index=False),
            file_name=csv_filename,
            mime="text/csv",
            use_container_width=True
        )
    
    with st.expander("👁️ View Genes"):
        st.dataframe(df_grouped, use_container_width=True, height=300)

st.divider()
st.caption("All processing happens on the server. Data is not stored permanently.")
