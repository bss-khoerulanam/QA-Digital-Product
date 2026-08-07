/**
 * QRIS UAT Document Generator - Web Version
 * Bank Sahabat Sampoerna (BSS)
 * 
 * Generates UAT Result and Lampiran 7C documents from UAT Script Excel
 */

// Global error handler - menampilkan error ke user, bukan cuma console
window.onerror = function(msg, url, line, col, error) {
    console.error('Error:', msg, 'at', url, ':', line);
    alert('Terjadi error: ' + msg + '\n\nLine: ' + line + '\nSilakan buka Console (F12) untuk detail.');
    return false;
};

window.addEventListener('unhandledrejection', function(event) {
    console.error('Unhandled promise rejection:', event.reason);
    alert('Terjadi error async: ' + (event.reason?.message || event.reason) + '\n\nSilakan buka Console (F12) untuk detail.');
});

// =============================================================================
// GLOBAL STATE
// =============================================================================

let parsedScenarios = [];
let metadata = {
    nama_penyedia: "Bank Sahabat Sampoerna",
    nama_layanan: "API QR MPM",
    nama_pengguna: "",
    tanggal_pengujian: ""
};
let skippedSet = new Set();

// =============================================================================
// FILE UPLOAD & DRAG/DROP
// =============================================================================

// PENTING: Prevent browser default drag behavior di seluruh halaman
// Tanpa ini, drag file ke browser akan membuka file di tab baru
document.addEventListener('dragover', (e) => {
    e.preventDefault();
    e.stopPropagation();
});
document.addEventListener('drop', (e) => {
    e.preventDefault();
    e.stopPropagation();
});

const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const fileInfo = document.getElementById('fileInfo');

dropZone.addEventListener('click', () => fileInput.click());
dropZone.addEventListener('dragover', (e) => {
    e.preventDefault();
    e.stopPropagation();
    dropZone.classList.add('dragover');
});
dropZone.addEventListener('dragleave', (e) => {
    e.preventDefault();
    e.stopPropagation();
    dropZone.classList.remove('dragover');
});
dropZone.addEventListener('drop', (e) => {
    e.preventDefault();
    e.stopPropagation();
    dropZone.classList.remove('dragover');
    const file = e.dataTransfer.files[0];
    if (file) handleFile(file);
});
fileInput.addEventListener('change', (e) => {
    const file = e.target.files[0];
    if (file) handleFile(file);
});

function handleFile(file) {
    if (!file.name.match(/\.xlsx?$/i)) {
        alert('Format file harus .xlsx atau .xls');
        return;
    }

    dropZone.classList.add('has-file');
    dropZone.querySelector('.drop-zone-icon').textContent = '✅';
    dropZone.querySelector('.drop-zone-text').innerHTML = `<strong>${file.name}</strong><br><small>File berhasil dimuat</small>`;

    fileInfo.textContent = `File: ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
    fileInfo.classList.add('show');

    const reader = new FileReader();
    reader.onload = (e) => {
        parseExcel(e.target.result);
    };
    reader.readAsArrayBuffer(file);
}

// =============================================================================
// EXCEL PARSER
// =============================================================================

function parseExcel(data) {
    const workbook = XLSX.read(data, { type: 'array' });
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

    // Extract metadata from header rows (first 6 rows)
    for (let i = 0; i < Math.min(6, jsonData.length); i++) {
        const row = jsonData[i];
        if (!row) continue;
        for (let j = 0; j < row.length; j++) {
            const cellValue = String(row[j] || '');
            if (cellValue.includes('Nama Penyedia')) {
                metadata.nama_penyedia = extractValue(cellValue, row, j);
            } else if (cellValue.includes('Nama Layanan')) {
                metadata.nama_layanan = extractValue(cellValue, row, j);
            } else if (cellValue.includes('Nama Pengguna')) {
                metadata.nama_pengguna = extractValue(cellValue, row, j);
            } else if (cellValue.includes('Tanggal')) {
                metadata.tanggal_pengujian = extractValue(cellValue, row, j);
            }
        }
    }

    // Find header row
    let headerRowIdx = -1;
    for (let i = 0; i < jsonData.length; i++) {
        const row = jsonData[i];
        if (!row) continue;
        const rowLower = row.map(c => String(c || '').toLowerCase().trim());
        if (rowLower.includes('no') && rowLower.includes('service')) {
            headerRowIdx = i;
            break;
        }
    }

    if (headerRowIdx === -1) {
        alert('ERROR: Tidak dapat menemukan header row (No, Service, Scenario...) di Excel.');
        return;
    }

    // Parse scenarios
    parsedScenarios = [];
    for (let i = headerRowIdx + 1; i < jsonData.length; i++) {
        const row = jsonData[i];
        if (!row || !row[0] || String(row[0]).trim() === '') continue;

        const scenario = {
            no: String(row[0] || '').trim(),
            service: String(row[1] || '').trim(),
            scenario: String(row[2] || '').trim(),
            expected_result: String(row[3] || '').trim(),
            request: String(row[4] || '').trim(),
            response: String(row[5] || '').trim(),
            result: String(row[6] || '').trim(),
            notes: String(row[7] || '').trim()
        };
        parsedScenarios.push(scenario);
    }

    // Auto-detect skipped scenarios
    skippedSet = new Set();
    parsedScenarios.forEach(s => {
        if (s.notes && (s.notes.toLowerCase().includes('tidak dites') || s.notes.toLowerCase().includes('tidak dilakukan'))) {
            skippedSet.add(s.no);
        }
    });

    // Update UI
    document.getElementById('configCard').style.display = 'block';
    document.getElementById('namaPengguna').value = metadata.nama_pengguna || '';
    document.getElementById('tanggalPengujian').value = metadata.tanggal_pengujian || '';
    document.getElementById('skippedScenarios').value = Array.from(skippedSet).join(', ');

    // Validate and show results
    validateAndShowResults();
}

function extractValue(cellValue, row, colIdx) {
    if (cellValue.includes(':')) {
        const parts = cellValue.split(':');
        if (parts.length > 1 && parts.slice(1).join(':').trim()) {
            return parts.slice(1).join(':').trim();
        }
    }
    // Try next column
    if (row[colIdx + 1]) {
        return String(row[colIdx + 1]).trim();
    }
    return '';
}

// =============================================================================
// VALIDATION LOGIC
// =============================================================================

function extractExpectedCode(expectedText) {
    if (!expectedText) return null;
    const patterns = [
        /[Rr]esponse\s*[Cc]ode[:\s]*(\d{7})/,
        /[Ee]rror\s*[Cc]ode[:\s]*(\d+xx\d+)/,
        /[Ee]rror\s*[Cc]ode[:\s]*(\d{7})/,
        /(\d{7})/,
        /(\d{3}xx\d{2})/
    ];
    for (const pattern of patterns) {
        const match = expectedText.match(pattern);
        if (match) return match[1];
    }
    return null;
}

function extractResponseCode(responseText) {
    if (!responseText) return null;
    const patterns = [
        /"responseCode"\s*:\s*"(\d+)"/,
        /"httpCode"\s*:\s*(\d+)/,
        /"responseCode"\s*:\s*(\d+)/,
        /responseCode.*?(\d{7})/,
        /[Cc]ode.*?(\d{7})/
    ];
    for (const pattern of patterns) {
        const match = responseText.match(pattern);
        if (match) return match[1];
    }
    return null;
}

function validateScenario(scenario) {
    // If already filled by mitra
    if (scenario.result && ['PASS', 'NOT PASS', 'N/A'].includes(scenario.result.toUpperCase())) {
        return scenario.result.toUpperCase();
    }

    // If no response provided
    if (!scenario.response || scenario.response === '' || scenario.response === 'Response Body:') {
        return 'NOT TESTED';
    }

    const expectedCode = extractExpectedCode(scenario.expected_result);
    const actualCode = extractResponseCode(scenario.response);

    if (!expectedCode || !actualCode) return 'NOT TESTED';

    // Handle xx pattern
    if (expectedCode.includes('xx')) {
        const regexPattern = expectedCode.replace('xx', '\\d{2}');
        if (new RegExp('^' + regexPattern + '$').test(actualCode)) {
            return 'PASS';
        } else {
            return 'NOT PASS';
        }
    }

    // Direct comparison
    return expectedCode === actualCode ? 'PASS' : 'NOT PASS';
}

function isScenarioSkipped(scenario) {
    if ((!scenario.request || scenario.request === '') && (!scenario.response || scenario.response === '')) {
        if (scenario.notes && scenario.notes.toLowerCase().includes('tidak')) return true;
        if (skippedSet.has(scenario.no)) return true;
    }
    return false;
}

// =============================================================================
// RESULTS DISPLAY
// =============================================================================

function validateAndShowResults() {
    const resultsCard = document.getElementById('resultsCard');
    const generateCard = document.getElementById('generateCard');
    const resultsBody = document.getElementById('resultsBody');
    const summaryCards = document.getElementById('summaryCards');

    resultsCard.style.display = 'block';
    generateCard.style.display = 'block';

    let passed = 0, failed = 0, na = 0, notTested = 0;

    resultsBody.innerHTML = '';
    parsedScenarios.forEach(s => {
        const isSkipped = isScenarioSkipped(s);
        let result;
        if (isSkipped) {
            result = 'N/A';
            na++;
        } else {
            result = validateScenario(s);
            if (result === 'PASS') passed++;
            else if (result === 'NOT PASS') failed++;
            else notTested++;
        }

        const expectedCode = extractExpectedCode(s.expected_result) || '-';
        const actualCode = extractResponseCode(s.response) || '-';

        let badgeClass = 'badge-pending';
        if (result === 'PASS') badgeClass = 'badge-pass';
        else if (result === 'NOT PASS') badgeClass = 'badge-fail';
        else if (result === 'N/A') badgeClass = 'badge-na';

        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${s.no}</td>
            <td>${s.service}</td>
            <td>${s.scenario.substring(0, 50)}${s.scenario.length > 50 ? '...' : ''}</td>
            <td><code>${expectedCode}</code></td>
            <td><code>${actualCode}</code></td>
            <td><span class="badge ${badgeClass}">${result}</span></td>
            <td>${s.notes ? s.notes.substring(0, 30) + '...' : ''}</td>
        `;
        resultsBody.appendChild(row);
    });

    const total = parsedScenarios.length;
    summaryCards.innerHTML = `
        <div class="summary-card total"><div class="number">${total}</div><div class="label">Total Skenario</div></div>
        <div class="summary-card pass"><div class="number">${passed}</div><div class="label">PASS</div></div>
        <div class="summary-card fail"><div class="number">${failed}</div><div class="label">NOT PASS</div></div>
        <div class="summary-card na"><div class="number">${na + notTested}</div><div class="label">N/A / Belum Diisi</div></div>
    `;

    const passRate = total > 0 ? ((passed / (total - na - notTested)) * 100) || 0 : 0;
    document.getElementById('progressFill').style.width = `${Math.min(passRate, 100)}%`;
}

// =============================================================================
// DOCX GENERATION - UAT RESULT
// =============================================================================

async function generateUATResult() {
    addLog('Generating UAT Result...', 'info');
    updateMetadataFromUI();

    const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, 
            Table, TableRow, TableCell, WidthType, BorderStyle, 
            ShadingType, PageBreak } = docx;

    const children = [];

    // Title page
    children.push(new Paragraph({ text: '' }));
    children.push(new Paragraph({ text: '' }));
    children.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: 'UAT Result', bold: true, size: 48, font: 'Calibri' })]
    }));
    children.push(new Paragraph({ text: '' }));
    children.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: 'Penambahan Layanan QRIS Merchant Aggregator', bold: true, size: 32, font: 'Calibri' })]
    }));
    children.push(new Paragraph({ text: '' }));
    children.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [
            new TextRun({ text: `Nama Penyedia Layanan: ${metadata.nama_penyedia}`, size: 22 }),
            new TextRun({ text: `\nNama Pengguna Layanan: ${metadata.nama_pengguna}`, size: 22, break: 1 }),
            new TextRun({ text: `\nTanggal Pengujian: ${metadata.tanggal_pengujian}`, size: 22, break: 1 }),
        ]
    }));
    children.push(new Paragraph({ children: [new PageBreak()] }));

    // Table of Contents
    children.push(new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: 'Daftar Isi', bold: true })]
    }));

    const tocItems = [
        '1  Balance Services',
        '2  API Transaction History List',
        '3  QR MPM',
        '4  Pengecekan Mutasi Dan Jurnal',
        '5  Generate QR SNAP',
        '6  Refund Payment',
        '7  Query Payment',
        '8  Inquiry & Report'
    ];
    tocItems.forEach(item => {
        children.push(new Paragraph({
            indent: { left: 720 },
            children: [new TextRun({ text: item, size: 20 })]
        }));
    });
    children.push(new Paragraph({ children: [new PageBreak()] }));

    // Section 1 & 2: Skipped
    children.push(new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: '1 Balance Services', bold: true })]
    }));
    children.push(new Paragraph({
        children: [new TextRun({ text: 'Tidak dites karena tidak sesuai dengan kondisi produk.', italics: true })]
    }));
    children.push(new Paragraph({ text: '' }));

    children.push(new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: '2 API Transaction History List', bold: true })]
    }));
    children.push(new Paragraph({
        children: [new TextRun({ text: 'Tidak dites karena tidak sesuai dengan kondisi produk.', italics: true })]
    }));
    children.push(new Paragraph({ children: [new PageBreak()] }));

    // Section 3: QR MPM
    children.push(new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: '3 QR MPM', bold: true })]
    }));

    let subIdx = 1;
    parsedScenarios.forEach(s => {
        if (!s.no.startsWith('18.')) return;

        children.push(new Paragraph({
            heading: HeadingLevel.HEADING_2,
            children: [new TextRun({ text: `3.${subIdx} ${s.no} ${s.scenario}`, bold: true })]
        }));

        const isSkipped = isScenarioSkipped(s);
        if (isSkipped) {
            children.push(new Paragraph({
                children: [new TextRun({ text: 'Tidak dites karena tidak sesuai dengan kondisi produk.', italics: true })]
            }));
        } else {
            // Expected Result
            children.push(new Paragraph({
                children: [
                    new TextRun({ text: 'Expected Result: ', bold: true, size: 20 }),
                    new TextRun({ text: s.expected_result, size: 20 })
                ]
            }));

            // Request
            children.push(new Paragraph({
                children: [new TextRun({ text: 'Request:', bold: true, size: 20 })]
            }));
            if (s.request) {
                children.push(new Paragraph({
                    indent: { left: 720 },
                    children: [new TextRun({ text: s.request, font: 'Consolas', size: 16 })]
                }));
            }

            // Response
            children.push(new Paragraph({
                children: [new TextRun({ text: 'Response:', bold: true, size: 20 })]
            }));
            if (s.response) {
                children.push(new Paragraph({
                    indent: { left: 720 },
                    children: [new TextRun({ text: s.response, font: 'Consolas', size: 16 })]
                }));
            }

            // Result
            const result = validateScenario(s);
            const resultColor = result === 'PASS' ? '008000' : (result === 'NOT PASS' ? 'FF0000' : '666666');
            children.push(new Paragraph({
                children: [
                    new TextRun({ text: 'Result: ', bold: true, size: 20 }),
                    new TextRun({ text: result, bold: true, size: 20, color: resultColor })
                ]
            }));

            // Notes
            if (s.notes) {
                children.push(new Paragraph({
                    children: [
                        new TextRun({ text: 'Notes: ', bold: true, size: 20 }),
                        new TextRun({ text: s.notes, size: 20 })
                    ]
                }));
            }
        }
        children.push(new Paragraph({ text: '' }));
        subIdx++;
    });

    // Section 4: Pengecekan Mutasi
    children.push(new Paragraph({ children: [new PageBreak()] }));
    children.push(new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: '4 Pengecekan Mutasi Dan Jurnal', bold: true })]
    }));
    children.push(new Paragraph({
        children: [new TextRun({ text: 'Hasil pengecekan mutasi dan jurnal akan dilampirkan terpisah.' })]
    }));

    // Summary section
    children.push(new Paragraph({ children: [new PageBreak()] }));
    children.push(new Paragraph({
        heading: HeadingLevel.HEADING_1,
        children: [new TextRun({ text: 'Ringkasan Hasil Pengujian', bold: true })]
    }));

    let passed = 0, failed = 0, naCount = 0, notTested = 0;
    parsedScenarios.forEach(s => {
        if (isScenarioSkipped(s)) { naCount++; return; }
        const r = validateScenario(s);
        if (r === 'PASS') passed++;
        else if (r === 'NOT PASS') failed++;
        else notTested++;
    });

    const summaryData = [
        ['Kategori', 'Jumlah'],
        ['Total Skenario', String(parsedScenarios.length)],
        ['PASS', String(passed)],
        ['NOT PASS', String(failed)],
        ['Tidak Diuji (N/A)', String(naCount)],
        ['Belum Diisi', String(notTested)]
    ];

    const summaryRows = summaryData.map((row, idx) => {
        return new TableRow({
            children: row.map(cellText => new TableCell({
                children: [new Paragraph({
                    children: [new TextRun({
                        text: cellText,
                        bold: idx === 0,
                        size: 20,
                        color: idx === 0 ? 'FFFFFF' : '000000'
                    })]
                })],
                shading: idx === 0 ? { type: ShadingType.SOLID, color: '4472C4' } : undefined,
                width: { size: 4000, type: WidthType.DXA }
            }))
        });
    });

    children.push(new Table({ rows: summaryRows }));

    // Create document
    const doc = new Document({
        sections: [{ children: children }]
    });

    const blob = await Packer.toBlob(doc);
    const timestamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    saveAs(blob, `UAT_Result_QRIS_Merchant_Aggregator_${timestamp}.docx`);
    addLog('UAT Result berhasil di-generate!', 'success');
}

// =============================================================================
// DOCX GENERATION - LAMPIRAN 7C
// =============================================================================

async function generateLampiran7C() {
    addLog('Generating Lampiran 7C...', 'info');
    updateMetadataFromUI();

    const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
            Table, TableRow, TableCell, WidthType, BorderStyle,
            ShadingType, PageOrientation } = docx;

    const children = [];

    // Header
    children.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: 'Lampiran 7.C', bold: true, size: 28 })]
    }));
    children.push(new Paragraph({
        alignment: AlignmentType.CENTER,
        children: [new TextRun({ text: 'Skenario dan Hasil Uji Fungsionalitas', bold: true, size: 24 })]
    }));
    children.push(new Paragraph({ text: '' }));

    // Metadata
    const metaItems = [
        ['Nama Penyedia Layanan', metadata.nama_penyedia],
        ['Nama Pengguna Layanan', metadata.nama_pengguna],
        ['Nama Layanan API', metadata.nama_layanan],
        ['Tanggal Pengujian', metadata.tanggal_pengujian]
    ];
    metaItems.forEach(([label, value]) => {
        children.push(new Paragraph({
            children: [
                new TextRun({ text: `${label}: `, bold: true, size: 20 }),
                new TextRun({ text: value || '', size: 20 })
            ]
        }));
    });
    children.push(new Paragraph({ text: '' }));

    // Scenario table
    const headerTexts = ['No', 'Service', 'Scenario', 'Expected Result', 'Request', 'Response', 'Result', 'Notes'];

    const headerRow = new TableRow({
        children: headerTexts.map(h => new TableCell({
            children: [new Paragraph({
                alignment: AlignmentType.CENTER,
                children: [new TextRun({ text: h, bold: true, size: 16, color: 'FFFFFF' })]
            })],
            shading: { type: ShadingType.SOLID, color: '4472C4' },
            width: { size: h === 'No' ? 600 : (h === 'Result' ? 800 : 1500), type: WidthType.DXA }
        }))
    });

    const dataRows = parsedScenarios.map(s => {
        const isSkipped = isScenarioSkipped(s);
        let resultText, notesText;
        if (isSkipped) {
            resultText = 'N/A';
            notesText = s.notes || 'Tidak dites karena tidak sesuai dengan kondisi produk.';
        } else {
            resultText = validateScenario(s);
            notesText = s.notes || '';
        }

        const resultColor = resultText === 'PASS' ? '008000' : (resultText === 'NOT PASS' ? 'FF0000' : '666666');

        const cellData = [s.no, s.service, s.scenario, s.expected_result, s.request || '', s.response || '', resultText, notesText];

        return new TableRow({
            children: cellData.map((text, idx) => new TableCell({
                children: [new Paragraph({
                    alignment: idx === 6 ? AlignmentType.CENTER : AlignmentType.LEFT,
                    children: [new TextRun({
                        text: text || '',
                        size: 14,
                        bold: idx === 6,
                        color: idx === 6 ? resultColor : '000000',
                        font: (idx === 4 || idx === 5) ? 'Consolas' : 'Calibri'
                    })]
                })],
                width: { size: idx === 0 ? 600 : (idx === 6 ? 800 : 1500), type: WidthType.DXA }
            }))
        });
    });

    children.push(new Table({
        rows: [headerRow, ...dataRows],
        width: { size: 100, type: WidthType.PERCENTAGE }
    }));

    // Footer notes
    children.push(new Paragraph({ text: '' }));
    const footerNotes = [
        'Lampiran Skenario hasil uji fungsional sekurangnya 1 Pengguna Layanan atas 1 sub API unverified, dengan ketentuan sebagai berikut:',
        '',
        'a. Pada kolom request diisi dengan request yang dilakukan Pengguna layanan, sedangkan pada kolom response diisi dengan respon yang diberikan Penyedia. Sementara pada kolom result diisi dengan hasil PASS atau NOT PASS yang harus sesuai dengan expected result.',
        '',
        'b. Pengisian pada dokumen skenario hasil uji fungsional tidak dilakukan dengan cara screen capture, melainkan dilakukan dengan cara copy paste payload request dan response dari log API server ke kolom tabel skenario hasil uji fungsional.',
        '',
        'c. Seluruh skenario diujikan dan tidak boleh dihapus atau diubah. Dalam hal terdapat skenario yang tidak diujikan dapat dikosongkan pengisiannya, namun diberikan catatan pada kolom Notes yang akan kami review lebih lanjut apakah skenario diperkenankan untuk tidak diujikan.',
        '',
        'd. Dalam hal terdapat penambahan skenario pengujian, maka penambahan tersebut dilakukan pada baris paling bawah, sehingga tidak mengubah susunan atau urutan template skenario.'
    ];
    footerNotes.forEach(note => {
        children.push(new Paragraph({
            indent: { left: 720 },
            children: [new TextRun({ text: note, size: 16, italics: true })]
        }));
    });

    // Create document (landscape)
    const doc = new Document({
        sections: [{
            properties: {
                page: {
                    size: { orientation: PageOrientation.LANDSCAPE }
                }
            },
            children: children
        }]
    });

    const blob = await Packer.toBlob(doc);
    const timestamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
    saveAs(blob, `Lampiran_7C_QRIS_${timestamp}.docx`);
    addLog('Lampiran 7C berhasil di-generate!', 'success');
}

// =============================================================================
// UTILITIES
// =============================================================================

async function generateBoth() {
    await generateUATResult();
    await generateLampiran7C();
    addLog('Kedua dokumen berhasil di-generate!', 'success');
}

function updateMetadataFromUI() {
    metadata.nama_penyedia = document.getElementById('namaPenyedia').value || metadata.nama_penyedia;
    metadata.nama_pengguna = document.getElementById('namaPengguna').value || metadata.nama_pengguna;
    metadata.nama_layanan = document.getElementById('namaLayanan').value || metadata.nama_layanan;
    metadata.tanggal_pengujian = document.getElementById('tanggalPengujian').value || metadata.tanggal_pengujian;

    // Update skipped scenarios from textarea
    const skippedInput = document.getElementById('skippedScenarios').value;
    if (skippedInput.trim()) {
        const additional = skippedInput.split(',').map(s => s.trim()).filter(s => s);
        additional.forEach(s => skippedSet.add(s));
    }
}

function addLog(message, type) {
    const logArea = document.getElementById('logArea');
    logArea.style.display = 'block';
    const timestamp = new Date().toLocaleTimeString('id-ID');
    const className = type === 'success' ? 'log-success' : (type === 'error' ? 'log-error' : 'log-info');
    logArea.innerHTML += `<div class="${className}">[${timestamp}] ${message}</div>`;
    logArea.scrollTop = logArea.scrollHeight;
}
