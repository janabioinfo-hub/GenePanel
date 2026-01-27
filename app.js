// Global variables
let coverageData = null;
let panelGenes = [];
let filteredData = [];

// Wait for all libraries to load
let librariesLoaded = false;

window.addEventListener('load', function() {
    // Give libraries time to initialize
    setTimeout(function() {
        if (typeof docx !== 'undefined' && typeof Papa !== 'undefined' && 
            typeof XLSX !== 'undefined' && typeof saveAs !== 'undefined') {
            librariesLoaded = true;
            console.log('All libraries loaded successfully');
        } else {
            console.error('Library loading status:', {
                docx: typeof docx !== 'undefined',
                Papa: typeof Papa !== 'undefined', 
                XLSX: typeof XLSX !== 'undefined',
                saveAs: typeof saveAs !== 'undefined'
            });
            alert('Some libraries failed to load. Please refresh the page.');
        }
    }, 1000);
});

// File upload handlers
document.addEventListener('DOMContentLoaded', function() {
    document.getElementById('coverageFile').addEventListener('change', handleCoverageUpload);
    document.getElementById('panelFile').addEventListener('change', handlePanelUpload);
});

function handleCoverageUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const statusDiv = document.getElementById('coverageStatus');
    statusDiv.textContent = 'Reading CSV file...';
    statusDiv.className = 'status';
    statusDiv.style.display = 'block';

    Papa.parse(file, {
        header: true,
        skipEmptyLines: true,
        complete: function(results) {
            try {
                // Trim column names
                const data = results.data.map(row => {
                    const trimmedRow = {};
                    Object.keys(row).forEach(key => {
                        trimmedRow[key.trim()] = row[key];
                    });
                    return trimmedRow;
                });

                // Find the coverage column
                let coverageCol = null;
                if (data[0].hasOwnProperty('% 1x')) {
                    coverageCol = '% 1x';
                } else if (data[0].hasOwnProperty('%1x')) {
                    coverageCol = '%1x';
                } else {
                    throw new Error("No '% 1x' or '%1x' column found in CSV");
                }

                // Normalize the data
                coverageData = data.map(row => ({
                    'Gene_Name': row['Gene_Name'] || '',
                    'Ref Name': row['Ref Name'] || '',
                    'Perc_1x': parseFloat(row[coverageCol]) || 0
                }));

                statusDiv.textContent = `✓ Loaded ${coverageData.length} coverage records`;
                statusDiv.className = 'status success';
                checkReadyToProcess();
            } catch (error) {
                statusDiv.textContent = `✗ Error: ${error.message}`;
                statusDiv.className = 'status error';
            }
        },
        error: function(error) {
            statusDiv.textContent = `✗ Error parsing CSV: ${error.message}`;
            statusDiv.className = 'status error';
        }
    });
}

function handlePanelUpload(event) {
    const file = event.target.files[0];
    if (!file) return;

    const statusDiv = document.getElementById('panelStatus');
    statusDiv.textContent = 'Reading Excel file...';
    statusDiv.className = 'status';
    statusDiv.style.display = 'block';

    const reader = new FileReader();
    reader.onload = function(e) {
        try {
            const data = new Uint8Array(e.target.result);
            const workbook = XLSX.read(data, { type: 'array' });
            
            // Read the first sheet
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            const jsonData = XLSX.utils.sheet_to_json(firstSheet);

            // Extract GENE column
            panelGenes = [];
            jsonData.forEach(row => {
                if (row.GENE) {
                    panelGenes.push(row.GENE.toString().trim());
                }
            });

            // Remove duplicates
            panelGenes = [...new Set(panelGenes)];

            statusDiv.textContent = `✓ Loaded ${panelGenes.length} genes from panel`;
            statusDiv.className = 'status success';
            checkReadyToProcess();
        } catch (error) {
            statusDiv.textContent = `✗ Error: ${error.message}`;
            statusDiv.className = 'status error';
        }
    };
    reader.readAsArrayBuffer(file);
}

function switchTab(tabName) {
    // Update tab buttons
    document.querySelectorAll('.tab-button').forEach(btn => {
        btn.classList.remove('active');
    });
    event.target.classList.add('active');

    // Update tab content
    document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.remove('active');
    });
    
    if (tabName === 'upload') {
        document.getElementById('uploadTab').classList.add('active');
    } else {
        document.getElementById('pasteTab').classList.add('active');
    }
}

function loadGeneList() {
    const geneListText = document.getElementById('geneList').value;
    if (!geneListText.trim()) {
        alert('Please paste gene names first');
        return;
    }

    // Parse gene list (one per line)
    panelGenes = geneListText.split('\n')
        .map(gene => gene.trim())
        .filter(gene => gene.length > 0);

    // Remove duplicates
    panelGenes = [...new Set(panelGenes)];

    const statusDiv = document.getElementById('panelStatus');
    statusDiv.textContent = `✓ Loaded ${panelGenes.length} genes from list`;
    statusDiv.className = 'status success';
    statusDiv.style.display = 'block';

    checkReadyToProcess();
}

function checkReadyToProcess() {
    const processBtn = document.getElementById('processBtn');
    if (coverageData && panelGenes.length > 0) {
        processBtn.disabled = false;
        // Auto-process to show preview
        filterData();
    } else {
        processBtn.disabled = true;
    }
}

function filterData() {
    if (!coverageData || panelGenes.length === 0) return;

    // Filter where either Gene_Name or Ref Name is in panel
    const filtered = coverageData.filter(row => {
        return panelGenes.includes(row.Gene_Name) || panelGenes.includes(row['Ref Name']);
    });

    // Use Gene_Name if present, otherwise Ref Name
    filtered.forEach(row => {
        row.Gene_ID = row.Gene_Name || row['Ref Name'];
    });

    // Remove rows starting with "Intron:"
    const noIntrons = filtered.filter(row => {
        return !row.Gene_ID.startsWith('Intron:');
    });

    // Group by Gene_ID and average coverage
    const grouped = {};
    noIntrons.forEach(row => {
        if (!grouped[row.Gene_ID]) {
            grouped[row.Gene_ID] = [];
        }
        grouped[row.Gene_ID].push(row.Perc_1x);
    });

    filteredData = Object.keys(grouped).map(geneId => {
        const coverages = grouped[geneId];
        const avgCoverage = coverages.reduce((a, b) => a + b, 0) / coverages.length;
        return {
            Gene_ID: geneId,
            Perc_1x: Math.round(avgCoverage * 100) / 100
        };
    });

    // Sort by gene name
    filteredData.sort((a, b) => a.Gene_ID.localeCompare(b.Gene_ID));

    showPreview();
}

function showPreview() {
    const previewSection = document.getElementById('previewSection');
    const genePreview = document.getElementById('genePreview');
    
    previewSection.style.display = 'block';
    
    genePreview.innerHTML = filteredData.map(row => {
        const className = row.Perc_1x < 90 ? 'gene-item low-coverage' : 'gene-item';
        return `<span class="${className}">${row.Gene_ID} (${row.Perc_1x}%)</span>`;
    }).join('');

    document.getElementById('downloadCsvBtn').disabled = false;
}

async function processData() {
    if (!filteredData || filteredData.length === 0) {
        alert('No data to process');
        return;
    }

    if (!librariesLoaded) {
        alert('Libraries are still loading. Please wait a moment and try again.');
        return;
    }

    // Show progress
    const progressSection = document.getElementById('progressSection');
    const progressFill = document.getElementById('progressFill');
    const progressText = document.getElementById('progressText');
    
    progressSection.style.display = 'block';
    progressFill.style.width = '30%';
    progressText.textContent = 'Creating Word document...';

    try {
        await generateWordDocument();
        
        progressFill.style.width = '100%';
        progressText.textContent = '✓ Document generated successfully!';
        
        setTimeout(() => {
            progressSection.style.display = 'none';
            progressFill.style.width = '0%';
        }, 2000);
    } catch (error) {
        console.error('Document generation error:', error);
        progressText.textContent = `✗ Error: ${error.message}`;
        progressFill.style.width = '0%';
        
        // Show more detailed error
        setTimeout(() => {
            alert('Error generating document. Please:\n1. Refresh the page\n2. Try again\n3. Check browser console (F12) for details');
        }, 1000);
    }
}

async function generateWordDocument() {
    // Check if docx library is available
    if (typeof docx === 'undefined') {
        throw new Error('Word library not loaded. Please refresh the page.');
    }

    const { Document, Packer, Paragraph, Table, TableCell, TableRow, TextRun, 
            WidthType, AlignmentType, VerticalAlign, BorderStyle, HeadingLevel } = docx;

    // Create document sections
    const sections = [{
        properties: {},
        children: [
            new Paragraph({
                text: "Appendix 1: Gene Coverage",
                heading: HeadingLevel.HEADING_1,
            }),
            new Paragraph({
                text: "Indication Based Analysis:",
                heading: HeadingLevel.HEADING_2,
            }),
            new Paragraph({ text: "" }), // Spacer
            createCoverageTable(filteredData)
        ]
    }];

    // Create document
    const doc = new Document({ sections });

    // Generate blob and download
    const blob = await Packer.toBlob(doc);
    saveAs(blob, "gene_coverage_report.docx");
}

function createCoverageTable(data) {
    const { Table, TableCell, TableRow, TextRun, Paragraph, 
            WidthType, AlignmentType, VerticalAlign, BorderStyle } = docx;

    const chunkSize = 4;
    const chunks = [];
    for (let i = 0; i < data.length; i += chunkSize) {
        chunks.push(data.slice(i, i + chunkSize));
    }

    const borderStyle = {
        style: BorderStyle.SINGLE,
        size: 4,
        color: "000000"
    };

    const borders = {
        top: borderStyle,
        bottom: borderStyle,
        left: borderStyle,
        right: borderStyle,
        insideHorizontal: borderStyle,
        insideVertical: borderStyle
    };

    // Create header row
    const headerCells = [];
    for (let i = 0; i < chunkSize; i++) {
        headerCells.push(
            new TableCell({
                children: [new Paragraph({
                    children: [new TextRun({
                        text: "Gene Name",
                        font: "Calibri",
                        size: 16
                    })],
                    alignment: AlignmentType.CENTER
                })],
                verticalAlign: VerticalAlign.CENTER,
                width: { size: 25, type: WidthType.PERCENTAGE }
            }),
            new TableCell({
                children: [new Paragraph({
                    children: [new TextRun({
                        text: "Percentage of coding region covered",
                        font: "Calibri",
                        size: 16
                    })],
                    alignment: AlignmentType.CENTER
                })],
                verticalAlign: VerticalAlign.CENTER,
                width: { size: 25, type: WidthType.PERCENTAGE }
            })
        );
    }

    const rows = [new TableRow({ children: headerCells })];

    // Create data rows
    chunks.forEach(chunk => {
        const rowCells = [];
        for (let i = 0; i < chunkSize; i++) {
            if (i < chunk.length) {
                const row = chunk[i];
                const isLowCoverage = row.Perc_1x < 90;

                rowCells.push(
                    new TableCell({
                        children: [new Paragraph({
                            children: [new TextRun({
                                text: row.Gene_ID,
                                font: "Calibri",
                                size: 16,
                                italics: true,
                                color: isLowCoverage ? "FF0000" : "000000"
                            })],
                            alignment: AlignmentType.CENTER
                        })],
                        verticalAlign: VerticalAlign.CENTER
                    }),
                    new TableCell({
                        children: [new Paragraph({
                            children: [new TextRun({
                                text: row.Perc_1x.toString(),
                                font: "Calibri",
                                size: 16,
                                color: isLowCoverage ? "FF0000" : "000000"
                            })],
                            alignment: AlignmentType.CENTER
                        })],
                        verticalAlign: VerticalAlign.CENTER
                    })
                );
            } else {
                // Empty cells
                rowCells.push(
                    new TableCell({
                        children: [new Paragraph({
                            children: [new TextRun({
                                text: "–",
                                font: "Calibri",
                                size: 16,
                                color: "FFFFFF"
                            })],
                            alignment: AlignmentType.CENTER
                        })],
                        verticalAlign: VerticalAlign.CENTER
                    }),
                    new TableCell({
                        children: [new Paragraph({
                            children: [new TextRun({
                                text: "–",
                                font: "Calibri",
                                size: 16,
                                color: "FFFFFF"
                            })],
                            alignment: AlignmentType.CENTER
                        })],
                        verticalAlign: VerticalAlign.CENTER
                    })
                );
            }
        }
        rows.push(new TableRow({ children: rowCells }));
    });

    return new Table({
        rows: rows,
        width: { size: 100, type: WidthType.PERCENTAGE },
        borders: borders,
        alignment: AlignmentType.CENTER
    });
}

function downloadCSV() {
    if (!filteredData || filteredData.length === 0) {
        alert('No data to download');
        return;
    }

    // Convert to CSV
    const headers = ['Gene_ID', 'Perc_1x'];
    const csv = [
        headers.join(','),
        ...filteredData.map(row => `${row.Gene_ID},${row.Perc_1x}`)
    ].join('\n');

    // Download
    const blob = new Blob([csv], { type: 'text/csv' });
    saveAs(blob, 'gene_coverage_filtered.csv');
}
