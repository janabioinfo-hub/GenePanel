import streamlit as st
import pandas as pd
from docx import Document
from docx.shared import RGBColor, Pt, Inches
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
import io
import os

def extract_gene_name(name):
    """Extract the first gene name from comma or semicolon-separated names."""
    if isinstance(name, str):
        if ',' in name:
            return name.split(',')[0]
        elif ';' in name:
            return name.split(';')[0]
        else:
            return name
    else:
        return None

def preprocess_excel(file_data, file_name, skip_rows=1):
    """Preprocesses raw Excel file to calculate coverage statistics."""
    try:
        # Load Excel data
        data = pd.read_excel(file_data, skiprows=skip_rows)
        
        if 'Gene Name' not in data.columns:
            raise ValueError("The column 'Gene Name' is not found in the file.")
        
        # Extract gene names
        data['Gene_Name'] = data['Gene Name'].apply(extract_gene_name)
        
        # Calculate coverage statistics
        result = []
        identifiers = data[['Gene Names', 'Aliases', 'Gene_Name']].fillna('')
        unique_ids = identifiers.drop_duplicates()
        
        total_genes = len(unique_ids)
        
        for idx, row in enumerate(unique_ids.iterrows()):
            _, row = row  # Unpack the tuple
            gene = row['Gene Names']
            alias = row['Aliases']
            name = row['Gene_Name']
            
            # Select matching rows
            if gene:
                gene_data = data[data['Gene Names'] == gene]
            elif alias:
                gene_data = data[data['Aliases'] == alias]
            elif name:
                gene_data = data[data['Name'] == name]
            else:
                continue
            
            # Total bases
            total_counted_bases = gene_data['Counted Bases'].sum()
            
            if total_counted_bases == 0:
                continue
            
            # Weighted mean depth
            mean_depth = (
                (gene_data['Mean Depth'] * gene_data['Counted Bases']).sum()
                / total_counted_bases
            )
            
            # Weighted %1x
            mean_1x = (
                (gene_data['% 1x'] * gene_data['Counted Bases']).sum()
                / total_counted_bases
            )
            
            min_depth = gene_data['Min Depth'].min()
            max_depth = gene_data['Max Depth'].max()
            
            summary = {
                'Region': 'total',
                'Ref Name': gene if gene else None,
                'Aliases': alias if alias else None,
                'Gene_Name': name if name else None,
                'Name': gene_data['Name'].iloc[0] if 'Name' in gene_data.columns else None,
                'Gene IDs': gene_data['Gene IDs'].iloc[0] if 'Gene IDs' in gene_data.columns else None,
                'Counted Bases': total_counted_bases,
                'Mean Depth': mean_depth,
                'Min Depth': min_depth,
                'Max Depth': max_depth,
                '% 1x': mean_1x,
            }
            
            result.append(pd.DataFrame([summary]))
        
        coverage_df = pd.concat(result, ignore_index=True)
        
        # Store the base filename without extension
        base_name = os.path.splitext(file_name)[0]
        
        return coverage_df, base_name
        
    except Exception as e:
        raise Exception(f"Error preprocessing Excel file: {str(e)}")

def create_word_document(df_grouped, output_filename):
    """Create Word document with gene coverage table"""
    doc = Document()
    doc.add_heading('Appendix 1: Gene Coverage', 1)
    doc.add_heading('Indication Based Analysis:', 2)
    
    # Add spacing paragraph with 0pt after
    spacer = doc.add_paragraph()
    spacer_format = spacer.paragraph_format
    spacer_format.space_after = Pt(0)
    
    chunk_size = 4
    chunks = [df_grouped[i:i+chunk_size] for i in range(0, len(df_grouped), chunk_size)]
    
    table = doc.add_table(rows=1, cols=chunk_size*2)
    table.alignment = 1
    hdr_cells = table.rows[0].cells
    
    # Header row
    for i in range(chunk_size):
        # Gene Name header
        gene_hdr_cell = hdr_cells[i*2]
        gene_hdr_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        gene_hdr_para = gene_hdr_cell.paragraphs[0]
        gene_hdr_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        gene_hdr_para.paragraph_format.space_after = Pt(0)
        gene_hdr_run = gene_hdr_para.add_run("Gene Name")
        gene_hdr_run.font.name = 'Calibri'
        gene_hdr_run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Calibri')
        gene_hdr_run.font.size = Pt(8)
        
        # Coverage header
        perc_hdr_cell = hdr_cells[i*2+1]
        perc_hdr_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        perc_hdr_para = perc_hdr_cell.paragraphs[0]
        perc_hdr_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        perc_hdr_para.paragraph_format.space_after = Pt(0)
        perc_hdr_run = perc_hdr_para.add_run("Percentage of coding region covered")
        perc_hdr_run.font.name = 'Calibri'
        perc_hdr_run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Calibri')
        perc_hdr_run.font.size = Pt(8)
    
    # Data rows
    for chunk in chunks:
        row_cells = table.add_row().cells
        for i in range(chunk_size):
            if i < len(chunk):
                row = chunk.iloc[i]
                gene = str(row['Gene_ID'])
                percent = row['Perc_1x']
                percent_str = str(percent)
                
                # Gene cell
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
                
                # Percentage cell
                perc_cell = row_cells[i*2+1]
                perc_cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
                perc_para = perc_cell.paragraphs[0]
                perc_para.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                perc_para.paragraph_format.space_after = Pt(0)
                perc_run = perc_para.add_run(percent_str)
                perc_run.font.name = 'Calibri'
                perc_run._element.rPr.rFonts.set(qn('w:eastAsia'), 'Calibri')
                perc_run.font.size = Pt(8)
                if percent < 90:
                    perc_run.font.color.rgb = RGBColor(255, 0, 0)
            else:
                # Empty cells
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
    
    # Add table borders
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

st.set_page_config(
    page_title="Gene Coverage Analyzer",
    page_icon="🧬",
    layout="wide"
)

# Custom CSS
st.markdown("""
    <style>
    .main {
        padding: 2rem;
    }
    .stButton>button {
        width: 100%;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        color: white;
        border: none;
        padding: 0.75rem 1.5rem;
        font-size: 1.1rem;
        border-radius: 8px;
        font-weight: 600;
    }
    .stButton>button:hover {
        background: linear-gradient(135deg, #764ba2 0%, #667eea 100%);
    }
    .upload-box {
        background: #f8f9fa;
        padding: 1.5rem;
        border-radius: 10px;
        border: 2px dashed #dee2e6;
    }
    .success-box {
        background: #d4edda;
        color: #155724;
        padding: 1rem;
        border-radius: 5px;
        border: 1px solid #c3e6cb;
        margin: 1rem 0;
    }
    .gene-preview {
        background: white;
        padding: 1rem;
        border-radius: 8px;
        border: 1px solid #dee2e6;
        max-height: 300px;
        overflow-y: auto;
    }
    .gene-badge {
        display: inline-block;
        padding: 0.3rem 0.6rem;
        margin: 0.2rem;
        background: #e9ecef;
        border-radius: 4px;
        font-family: monospace;
        font-size: 0.9rem;
    }
    .gene-badge-low {
        background: #f8d7da;
        color: #721c24;
        font-weight: bold;
    }
    </style>
""", unsafe_allow_html=True)

# Title
st.title("🧬 Gene Coverage Analyzer")
st.markdown("Process raw Excel data or upload pre-processed CSV to generate Word report")

# Initialize session state
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

# Step 1: Choose input type
st.markdown("### Step 1: Choose Input Type")
input_type = st.radio(
    "Select your input format:",
    ["📊 Pre-processed CSV (Skip preprocessing)", "📁 Raw Excel File (Run preprocessing first)"],
    help="Choose whether you already have a processed CSV or need to process raw Excel data first"
)

# Create two columns for upload
col1, col2 = st.columns(2)

with col1:
    if "Raw Excel" in input_type:
        st.markdown("### Step 2A: Upload Raw Excel File")
        raw_excel = st.file_uploader(
            "Choose Raw Excel file",
            type=['xlsx', 'xls'],
            key='raw_excel',
            help="Upload the raw coverage Excel file for preprocessing"
        )
        
        if raw_excel:
            # Create progress indicators
            progress_bar = st.progress(0)
            status_text = st.empty()
            
            try:
                # Step 1: Reading Excel file
                status_text.text("📂 Reading Excel file...")
                progress_bar.progress(20)
                
                coverage_df, basename = preprocess_excel(raw_excel, raw_excel.name)
                
                # Step 2: Processing complete
                status_text.text("⚙️ Calculating coverage statistics...")
                progress_bar.progress(70)
                
                st.session_state.coverage_data = coverage_df
                st.session_state.file_basename = basename
                
                # Step 3: Generating CSV
                status_text.text("📄 Generating CSV output...")
                progress_bar.progress(90)
                
                csv_buffer = io.StringIO()
                coverage_df.to_csv(csv_buffer, index=False)
                st.session_state.processed_csv = csv_buffer.getvalue()
                
                # Complete
                progress_bar.progress(100)
                status_text.text("✅ Processing complete!")
                
                # Clear progress indicators after a moment
                import time
                time.sleep(0.5)
                progress_bar.empty()
                status_text.empty()
                
                st.success(f"✅ Processed {len(coverage_df)} gene records from {raw_excel.name}")
                
                # Download button for processed CSV
                st.download_button(
                    label="⬇️ Download Processed CSV",
                    data=st.session_state.processed_csv,
                    file_name=f"{basename}_total_coverage_statistics.csv",
                    mime="text/csv",
                    use_container_width=True
                )
                
            except Exception as e:
                progress_bar.empty()
                status_text.empty()
                st.error(f"❌ Error processing Excel: {str(e)}")
    else:
        st.markdown("### Step 2A: Upload Coverage CSV")
        coverage_file = st.file_uploader(
            "Choose CSV file",
            type=['csv'],
            key='coverage',
            help="Must contain columns: Gene_Name, Ref Name, and % 1x (or %1x)"
        )
        
        if coverage_file:
            try:
                # Extract filename without extension
                st.session_state.file_basename = os.path.splitext(coverage_file.name)[0]
                
                df = pd.read_csv(coverage_file)
                df.columns = df.columns.str.strip()
                
                # Find coverage column
                if '% 1x' in df.columns:
                    coverage_col = '% 1x'
                elif '%1x' in df.columns:
                    coverage_col = '%1x'
                elif '% 1x ' in df.columns:
                    coverage_col = '% 1x '
                else:
                    st.error("❌ No '% 1x' or '%1x' column found in CSV")
                    coverage_col = None
                
                if coverage_col:
                    st.session_state.coverage_data = df
                    st.success(f"✅ Loaded {len(df)} coverage records from {coverage_file.name}")
            except Exception as e:
                st.error(f"❌ Error reading CSV: {str(e)}")

with col2:
    st.markdown("### Step 2B: Upload Panel Excel OR Paste Gene List")
    
    # Tab selection
    tab1, tab2 = st.tabs(["📁 Upload Excel", "📝 Paste Genes"])
    
    with tab1:
        panel_file = st.file_uploader(
            "Choose Excel file",
            type=['xlsx', 'xls'],
            key='panel',
            help="Must contain a 'GENE' column"
        )
        
        if panel_file:
            try:
                panel_df = pd.read_excel(panel_file)
                if 'GENE' in panel_df.columns:
                    genes = panel_df['GENE'].dropna().unique().tolist()
                    st.session_state.panel_genes = genes
                    st.success(f"✅ Loaded {len(genes)} genes from panel")
                else:
                    st.error("❌ No 'GENE' column found in Excel file")
            except Exception as e:
                st.error(f"❌ Error reading Excel: {str(e)}")
    
    with tab2:
        gene_text = st.text_area(
            "Paste gene names (one per line)",
            height=200,
            placeholder="BRCA1\nBRCA2\nTP53\nEGFR\nKRAS",
            help="Enter one gene name per line"
        )
        
        if st.button("Load Gene List"):
            if gene_text.strip():
                genes = [gene.strip() for gene in gene_text.split('\n') if gene.strip()]
                genes = list(set(genes))  # Remove duplicates
                st.session_state.panel_genes = genes
                st.success(f"✅ Loaded {len(genes)} genes from list")
            else:
                st.warning("⚠️ Please paste gene names first")

# Process data if both are available
if st.session_state.coverage_data is not None and st.session_state.panel_genes:
    df = st.session_state.coverage_data
    panel_genes = st.session_state.panel_genes
    
    # Find coverage column
    if '% 1x' in df.columns:
        coverage_col = '% 1x'
    elif '%1x' in df.columns:
        coverage_col = '%1x'
    elif '% 1x ' in df.columns:
        coverage_col = '% 1x '
    else:
        st.error("Coverage column not found")
        st.stop()
    
    # Filter data
    df_filtered = df[
        df['Gene_Name'].isin(panel_genes) | 
        df['Ref Name'].isin(panel_genes)
    ].copy()
    
    # Use Gene_Name if present, otherwise Ref Name
    df_filtered['Gene_ID'] = df_filtered['Gene_Name']
    df_filtered.loc[df_filtered['Gene_ID'].isna(), 'Gene_ID'] = df_filtered['Ref Name']
    
    # Remove rows starting with "Intron:"
    df_filtered = df_filtered[~df_filtered['Gene_ID'].fillna('').str.startswith("Intron:")]
    
    # Convert coverage to numeric
    df_filtered['Perc_1x'] = pd.to_numeric(df_filtered[coverage_col], errors='coerce')
    
    # Group by Gene_ID and take average
    df_grouped = df_filtered.groupby('Gene_ID', as_index=False)['Perc_1x'].mean()
    df_grouped['Perc_1x'] = df_grouped['Perc_1x'].round(2)
    df_grouped = df_grouped.sort_values('Gene_ID')
    
    st.session_state.filtered_data = df_grouped
    
    # Display preview
    st.markdown("---")
    
    # Summary stats
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Total Genes", len(df_grouped))
    with col2:
        low_coverage = len(df_grouped[df_grouped['Perc_1x'] < 90])
        st.metric("Low Coverage (<90%)", low_coverage)
    with col3:
        avg_coverage = df_grouped['Perc_1x'].mean()
        st.metric("Average Coverage", f"{avg_coverage:.2f}%")
    
    st.markdown("---")
    st.markdown("### 📥 Download Results")
    
    # Action buttons - prominently placed
    col1, col2 = st.columns(2)
    
    with col1:
        if st.button("📄 Generate Word Document", type="primary", use_container_width=True, key="gen_word"):
            with st.spinner("Creating Word document..."):
                try:
                    # Generate filename with basename
                    word_filename = f"{st.session_state.file_basename}_gene_coverage_report.docx"
                    
                    # Generate document
                    doc = create_word_document(df_grouped, word_filename)
                    
                    # Save to bytes
                    doc_bytes = io.BytesIO()
                    doc.save(doc_bytes)
                    doc_bytes.seek(0)
                    
                    # Store in session state for download
                    st.session_state.doc_bytes = doc_bytes
                    st.session_state.word_filename = word_filename
                    st.success("✅ Document ready! Click below to download.")
                except Exception as e:
                    st.error(f"❌ Error: {str(e)}")
        
        # Show download button if document is ready
        if 'doc_bytes' in st.session_state:
            st.download_button(
                label=f"⬇️ Download {st.session_state.word_filename}",
                data=st.session_state.doc_bytes,
                file_name=st.session_state.word_filename,
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                use_container_width=True,
                type="primary"
            )
    
    with col2:
        # CSV download - always available
        csv_filename = f"{st.session_state.file_basename}_gene_coverage_filtered.csv"
        csv = df_grouped.to_csv(index=False)
        st.download_button(
            label=f"📊 Download {csv_filename}",
            data=csv,
            file_name=csv_filename,
            mime="text/csv",
            use_container_width=True
        )
    
    st.markdown("---")
    
    with st.expander("👁️ View All Filtered Genes", expanded=False):
        # Create preview with badges
        preview_html = '<div class="gene-preview">'
        for _, row in df_grouped.iterrows():
            badge_class = "gene-badge-low" if row['Perc_1x'] < 90 else "gene-badge"
            preview_html += f'<span class="{badge_class}"><i>{row["Gene_ID"]}</i> ({row["Perc_1x"]}%)</span>'
        preview_html += '</div>'
        
        st.markdown(preview_html, unsafe_allow_html=True)
        st.info("ℹ️ Genes with coverage below 90% are highlighted in red")

# Footer
st.markdown("---")
st.markdown(
    "<p style='text-align: center; color: #6c757d;'>"
    "All processing happens on the server. Your data is not stored permanently."
    "</p>",
    unsafe_allow_html=True
)
