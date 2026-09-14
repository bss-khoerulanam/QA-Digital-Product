/**
 * QRIS Document Generator - Web Version
 * Bank Sahabat Sampoerna (BSS)
 *
 * Generates UAT Result and Lampiran 7C documents from UAT Script Excel.
 *
 * Parsing logic mirrors the reference Python tool (generate_uat_docs.py):
 *   - Flexible header-row detection (Kirimo Indonesian layout + old English layout)
 *   - NAME-based column mapping (no hardcoded positional columns)
 *   - Remarks column parsed into URL / Headers / Request Body / Response
 *   - PASS / NOT PASS / N/A / NOT TESTED classification per ASPI rules
 *
 * The parsing functions are also exported via module.exports (guarded) so they
 * can be unit-tested headless under Node/Bun. In the browser (<script>) the
 * guard is skipped and everything stays on the global scope.
 */

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
let currentFileData = null; // ArrayBuffer of the loaded file, awaiting "Proses File"
let currentFileName = "";

// Statuses considered "not tested" (mirror SKIP_STATUSES in Python)
const SKIP_STATUSES = ["tidak dites", "tidak ditest"];

// =============================================================================
// REMARKS PARSER - Split Remarks into URL / Headers / Request Body / Response
// Mirrors RemarksParser in generate_uat_docs.py
// =============================================================================

const RemarksParser = {
    parse(remarksText) {
        const result = {
            url: "",
            headers: "",
            request_body: "",
            response: "",
            full_request: "",
            full_response: ""
        };

        if (!remarksText || String(remarksText).trim() === "") {
            return result;
        }

        const text = String(remarksText).trim();

        const [requestPart, responsePart] = RemarksParser._splitRequestResponse(text);

        if (requestPart) {
            const { url, headers, body } = RemarksParser._parseRequestSection(requestPart);
            result.url = url;
            result.headers = headers;
            result.request_body = body;
            result.full_request = RemarksParser._formatRequestOutput(url, headers, body);
        }

        if (responsePart) {
            result.response = responsePart.trim();
            result.full_response = responsePart.trim();
        }

        return result;
    },

    _splitRequestResponse(text) {
        // Pattern 1: a standalone "Response" line acting as request|response
        // separator. Covers BOTH the classic "Response:" label (with colon) AND
        // the notification format used by e.g. qr-mpm-notify (Kirimo row 61),
        // where the marker is just the word "Response" on its own line WITHOUT a
        // colon and WITHOUT an "HTTP/1.1 ..." status line, directly followed by
        // the JSON response body ({ "responseCode": "2005200", ... }).
        //
        // Only a line whose trimmed content is exactly "Response" or "Response:"
        // is treated as the separator, so the word "response" inside a sentence
        // or a JSON key (e.g. "responseCode") is never mistaken for one.
        const responseLine = text.match(/^[ \t]*Response:?[ \t]*$/m);
        if (responseLine) {
            const start = responseLine.index;
            const end = responseLine.index + responseLine[0].length;
            return [text.slice(0, start).trim(), text.slice(end).trim()];
        }

        // Pattern 2: request JSON body immediately followed by an HTTP status line
        const jsonThenHttp = text.match(/\}\s*\n\s*\n*(HTTP\/\d\.\d\s+\d+)/);
        if (jsonThenHttp) {
            const splitPos = jsonThenHttp.index + jsonThenHttp[0].indexOf(jsonThenHttp[1]);
            return [text.slice(0, splitPos).trim(), text.slice(splitPos).trim()];
        }

        // Pattern 3: any HTTP response status line
        const httpResponse = text.match(/\n(HTTP\/\d\.\d\s+\d+\s*\n)/);
        if (httpResponse) {
            const start = httpResponse.index + 1; // keep after the leading \n
            return [text.slice(0, start).trim(), text.slice(start).trim()];
        }

        // No response section (e.g. notification scenario) - all request
        return [text, ""];
    },

    _parseRequestSection(requestText) {
        let url = "";
        let headers = "";
        let body = "";

        const lines = requestText.split('\n');

        // Skip a leading "URL:" / "Request:" label line
        let startIdx = 0;
        if (lines.length && /^(URL|Request)\s*:\s*$/i.test(lines[0].trim())) {
            startIdx = 1;
        }

        // Find the HTTP method line: "POST /path HTTP/1.1"
        let httpMethodIdx = -1;
        for (let i = startIdx; i < lines.length; i++) {
            if (/^(POST|GET|PUT|DELETE|PATCH)\s+\//.test(lines[i].trim())) {
                httpMethodIdx = i;
                break;
            }
        }

        if (httpMethodIdx >= 0) {
            const urlLine = lines[httpMethodIdx].trim();

            // Find Host header to build the full URL
            let host = "";
            for (let i = httpMethodIdx + 1; i < lines.length; i++) {
                if (lines[i].trim().toLowerCase().startsWith('host:')) {
                    host = lines[i].trim().split(':').slice(1).join(':').trim();
                    break;
                }
            }

            const methodMatch = urlLine.match(/(POST|GET|PUT|DELETE|PATCH)\s+(\S+)/);
            if (methodMatch && host) {
                url = `URL:\n${methodMatch[1]} https://${host}${methodMatch[2]}`;
            } else {
                url = `URL:\n${urlLine}`;
            }

            // Collect header lines after the method line until blank line or JSON body
            const headerLines = [];
            let bodyStartIdx = -1;
            for (let i = httpMethodIdx + 1; i < lines.length; i++) {
                const line = lines[i].trim();
                if (line.toLowerCase().startsWith('host:')) continue; // already in URL
                if (line === '') { bodyStartIdx = i + 1; break; }
                if (line.startsWith('{')) { bodyStartIdx = i; break; }
                if (/^[\w-]+[\w-]*\s*:/.test(line)) headerLines.push(line);
            }

            if (bodyStartIdx >= 0 && bodyStartIdx < lines.length) {
                let bodyText = lines.slice(bodyStartIdx).join('\n').trim();
                if (bodyText) {
                    const jsonStart = bodyText.indexOf('{');
                    if (jsonStart >= 0) {
                        bodyText = bodyText.slice(jsonStart);
                        body = RemarksParser._extractJsonObject(bodyText);
                    }
                }
            }

            headers = headerLines.join('\n').trim();
        } else {
            // No HTTP method line - notification format:
            //   "URL: https://..." then headers then a JSON body
            const headerLines = [];
            let bodyStart = -1;
            for (let i = startIdx; i < lines.length; i++) {
                const line = lines[i].trim();
                if (line.startsWith('{')) { bodyStart = i; break; }
                if (line === '') continue;
                if (/^url\s*:\s*\S/i.test(line)) {
                    const urlValue = line.split(':').slice(1).join(':').trim();
                    url = `URL:\n${urlValue}`;
                } else if (/^[\w-]+[\w-]*\s*:/.test(line)) {
                    headerLines.push(line);
                }
            }
            headers = headerLines.join('\n').trim();
            if (bodyStart >= 0) {
                const bodyText = lines.slice(bodyStart).join('\n').trim();
                body = RemarksParser._extractJsonObject(bodyText);
            }
        }

        return { url, headers, body };
    },

    _extractJsonObject(text) {
        // Return the first balanced {...} JSON object, else the whole text
        let braceCount = 0;
        let jsonEnd = -1;
        for (let i = 0; i < text.length; i++) {
            if (text[i] === '{') braceCount++;
            else if (text[i] === '}') {
                braceCount--;
                if (braceCount === 0) { jsonEnd = i + 1; break; }
            }
        }
        return jsonEnd > 0 ? text.slice(0, jsonEnd) : text;
    },

    _formatRequestOutput(url, headers, body) {
        const parts = [];
        if (url) parts.push(url);
        if (headers) parts.push(`\n${headers}`);
        if (body) parts.push(`\nRequest Body:\n${body}`);
        return parts.join('\n').trim();
    }
};

// =============================================================================
// RESPONSE PARSER - Extract responseCode from full HTTP response
// Mirrors ResponseParser.extract_response_code in generate_uat_docs.py
// =============================================================================

const ResponseParser = {
    extractResponseCode(responseText) {
        if (!responseText) return null;
        const patterns = [
            /"responseCode"\s*:\s*"(\d+)"/,
            /"responseCode"\s*:\s*(\d+)/,
            /responseCode["\s:]*(\d{7})/
        ];
        for (const pattern of patterns) {
            const match = responseText.match(pattern);
            if (match) return match[1];
        }
        return null;
    }
};

// =============================================================================
// SCENARIO NAME CLEANER - Mirrors clean_scenario_name in generate_uat_docs.py
// =============================================================================

function cleanScenarioName(langkahTes) {
    if (!langkahTes) return "";
    let cleaned = String(langkahTes);
    // Remove leading number prefix like "18,1 " or "3.5 "
    cleaned = cleaned.replace(/^\d+[,.]\d+\s*/, '');
    // Remove marker/comment artifacts (a line of ===== and everything after it)
    cleaned = cleaned.replace(/\n[\s\S]*?======[\s\S]*/, '');
    return cleaned.trim();
}

function extractAspiNumber(langkahTes) {
    if (!langkahTes) return "";
    const match = String(langkahTes).match(/(\d+)[,.](\d+)/);
    return match ? `${match[1]}.${match[2]}` : "";
}

// =============================================================================
// VALIDATION LOGIC - Mirrors ResultValidator in generate_uat_docs.py
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

// Kept for backward compatibility with the results table display.
function extractResponseCode(responseText) {
    return ResponseParser.extractResponseCode(responseText);
}

function escapeRegExp(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

/**
 * Validate a scenario -> 'PASS' | 'NOT PASS' | 'N/A' | 'NOT TESTED'.
 *
 * ASPI status rules (mirror ResultValidator.validate in Python):
 *   - PASS      : actual responseCode matches expected (incl. 'xx' wildcard).
 *   - NOT PASS  : actual responseCode present but does not match expected.
 *   - N/A       : scenario 'Tidak dites'/'Tidak ditest' (skipped), OR a
 *                 functional 'Berhasil' scenario whose expected result has no
 *                 extractable code and carries no Remarks payload, OR a
 *                 notification-format 'Berhasil' scenario with a request
 *                 payload but no Response section to validate.
 *   - NOT TESTED: expected an API code and had response data but the code
 *                 could not be extracted, or no response at all while not
 *                 skipped/functional.
 */
function validateScenario(scenario) {
    if (scenario.is_skipped) return "N/A";

    const expectedCode = extractExpectedCode(scenario.expected_result || "");
    const actualCode = ResponseParser.extractResponseCode(scenario.response || "");
    const hasResponse = !!(scenario.response && String(scenario.response).trim());
    const hasRemarksPayload = !!(
        (scenario.request && String(scenario.request).trim()) || hasResponse
    );
    const isBerhasil = String(scenario.hasil_aktual || "").toLowerCase() === "berhasil";

    // No response payload at all.
    if (!hasResponse) {
        if (isBerhasil && expectedCode === null && !hasRemarksPayload) return "N/A";
        if (isBerhasil && hasRemarksPayload) return "N/A";
        return "NOT TESTED";
    }

    // Response payload present but codes missing.
    if (!expectedCode || !actualCode) {
        if (expectedCode === null && isBerhasil) return "N/A";
        return "NOT TESTED";
    }

    // xx wildcard - anchor fully so 401xx01 matches exactly 4014701 (7 digits).
    if (expectedCode.includes("xx")) {
        const pattern = "^" + escapeRegExp(expectedCode).replace(/xx/g, "\\d{2}") + "$";
        return new RegExp(pattern).test(actualCode) ? "PASS" : "NOT PASS";
    }

    return expectedCode === actualCode ? "PASS" : "NOT PASS";
}

// =============================================================================
// EXCEL PARSER - flexible header detection + name-based column mapping
// Mirrors UATScriptParser in generate_uat_docs.py
// =============================================================================

function firstLine(v) {
    return v ? String(v).split('\n')[0].trim() : '';
}

function pickWorksheet(workbook) {
    // Prefer a sheet named like "UAT Script" (matches Python selection logic).
    for (const name of workbook.SheetNames) {
        const lower = name.toLowerCase();
        if (lower.includes('uat') && lower.includes('script')) return name;
        if (lower.trim() === 'uat script') return name;
        if (lower.includes('script') && !lower.includes('error')) return name;
    }
    return workbook.SheetNames[0];
}

/**
 * Detect header row + build a column map by header NAME.
 * Returns { headerRowIdx, colMap } or { headerRowIdx: -1 } if not found.
 */
function findHeaderRow(rows) {
    const maxScan = Math.min(50, rows.length);

    // Primary: Indonesian (Kirimo) layout.
    for (let i = 0; i < maxScan; i++) {
        const row = rows[i];
        if (!row) continue;
        const firstLines = row.map(c => firstLine(c).toLowerCase());

        const hasLangkah = firstLines.some(fl => fl.includes('langkah tes') || fl === 'langkah tes');
        const hasKategori = firstLines.some(fl => fl.includes('kategori'));
        const hasNamaModul = firstLines.some(fl => fl.includes('nama modul'));
        const hasHasil = firstLines.some(fl => fl.includes('hasil') && fl.includes('diharapkan'));
        const hasNomorKasus = firstLines.some(fl => fl.includes('nomor kasus') || fl.includes('kasus tes'));

        if (hasLangkah && (hasKategori || hasNamaModul || hasHasil || hasNomorKasus)) {
            const colMap = {};
            firstLines.forEach((fl, idx) => {
                if (fl === 'kategori' || (fl.startsWith('kategori') && !fl.includes('nama'))) colMap.kategori = idx;
                else if (fl.includes('nama modul')) colMap.nama_modul = idx;
                else if (fl.includes('nomor skenario')) colMap.nomor_skenario = idx;
                else if (fl.includes('nomor kasus') || fl.includes('kasus tes')) colMap.nomor_kasus_tes = idx;
                else if (fl.includes('langkah')) colMap.langkah_tes = idx;
                else if (fl.includes('hasil') && fl.includes('diharapkan')) colMap.hasil_diharapkan = idx;
                else if (fl.includes('hasil aktual')) colMap.hasil_aktual = idx;
                else if (fl.includes('remark')) colMap.remarks = idx;
                else if (fl.includes('tanggal') && fl.includes('pelaksanaan')) colMap.tanggal = idx;
                else if (fl.includes('jenis') && fl.includes('script')) colMap.jenis_script = idx;
                else if (fl.includes('pelaksana')) colMap.pelaksana = idx;
            });
            return { headerRowIdx: i, colMap, layout: 'id' };
        }
    }

    // Secondary: old English layout
    //   No | Service | Scenario | Expected Result | Request | Response | Result | Notes
    for (let i = 0; i < maxScan; i++) {
        const row = rows[i];
        if (!row) continue;
        const firstLines = row.map(c => firstLine(c).toLowerCase());

        const hasScenario = firstLines.some(fl => fl === 'scenario');
        const hasExpected = firstLines.some(fl => fl.includes('expected') && fl.includes('result'));
        const hasService = firstLines.some(fl => fl === 'service');

        if (hasScenario && hasExpected && hasService) {
            const colMap = {};
            firstLines.forEach((fl, idx) => {
                if (fl === 'service') colMap.nama_modul = idx;
                else if (fl === 'scenario') colMap.langkah_tes = idx;
                else if (fl.includes('expected') && fl.includes('result')) colMap.hasil_diharapkan = idx;
                else if (fl === 'request') colMap.request = idx;
                else if (fl === 'response') colMap.remarks = idx;
                else if (fl === 'result') colMap.hasil_aktual = idx;
                else if (fl === 'notes') colMap.notes = idx;
                else if (fl === 'no') colMap.nomor_kasus_tes = idx;
            });
            // If no explicit Response column, use Notes as remarks source.
            if (colMap.remarks === undefined && colMap.notes !== undefined) {
                colMap.remarks = colMap.notes;
            }
            return { headerRowIdx: i, colMap, layout: 'en' };
        }
    }

    return { headerRowIdx: -1, colMap: {}, layout: null };
}

const SECTION_KEYWORDS = [
    "Balance Services", "API Transaction History",
    "QR MPM", "PENGECEKAN MUTASI", "Generate QR SNAP",
    "Refund Payment", "Query Payment", "Inquiry"
];

/**
 * Parse a workbook's data array into structured scenarios.
 * Pure function so it can be unit-tested headless. Returns
 * { scenarios, metadata, headerRowIdx, layout, error }.
 */
function parseWorkbookData(rows) {
    const meta = {
        nama_penyedia: "Bank Sahabat Sampoerna",
        nama_layanan: "API QR MPM",
        nama_pengguna: "",
        tanggal_pengujian: ""
    };

    // --- Metadata extraction (mitra name in <...>, Tanggal, Nomor Referensi) ---
    for (let i = 0; i < Math.min(20, rows.length); i++) {
        const row = rows[i];
        if (!row) continue;
        for (let j = 0; j < row.length; j++) {
            const val = String(row[j] || '').trim();
            if (!val) continue;
            if (val.includes('<') && val.includes('>')) {
                const m = val.match(/<(.+?)>/);
                if (m) {
                    const mitra = m[1].match(/((?:PT|CV)\s+[\w\s]+(?:\([^)]+\))?)/);
                    if (mitra) meta.nama_pengguna = mitra[0].trim();
                }
            }
            if (val.includes('Nama Penyedia')) {
                const v = extractInlineValue(val, row, j);
                if (v) meta.nama_penyedia = v;
            } else if (val.includes('Nama Layanan')) {
                const v = extractInlineValue(val, row, j);
                if (v) meta.nama_layanan = v;
            } else if (val.includes('Nama Pengguna')) {
                const v = extractInlineValue(val, row, j);
                if (v) meta.nama_pengguna = v;
            } else if (val.includes('Tanggal') && !val.includes('Pelaksanaan')) {
                const v = extractInlineValue(val, row, j);
                if (v) meta.tanggal_pengujian = v;
            }
        }
    }

    const { headerRowIdx, colMap, layout } = findHeaderRow(rows);
    if (headerRowIdx === -1) {
        return { scenarios: [], metadata: meta, headerRowIdx: -1, layout: null, error: 'no-header' };
    }

    const get = (row, key, dflt) => {
        const idx = colMap[key];
        if (idx === undefined) return "";
        const v = row[idx];
        return v === undefined || v === null ? "" : String(v).trim();
    };

    const scenarios = [];
    let currentSection = "";

    for (let i = headerRowIdx + 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row) continue;

        // Skip completely empty rows
        if (row.every(v => v === undefined || v === null || String(v).trim() === "")) continue;

        const kategori = get(row, 'kategori', '');
        const namaModul = get(row, 'nama_modul', '');
        const langkahTes = get(row, 'langkah_tes', '');

        // Section header detection (row naming a section, no test step)
        let isSectionHeader = false;
        for (const kw of SECTION_KEYWORDS) {
            const kwl = kw.toLowerCase();
            if (kategori.toLowerCase().includes(kwl) || namaModul.toLowerCase().includes(kwl)) {
                if (!langkahTes) {
                    currentSection = kategori || namaModul;
                    isSectionHeader = true;
                    break;
                }
            }
        }
        if (isSectionHeader) continue;

        const nomorKasus = get(row, 'nomor_kasus_tes', '');
        if (!nomorKasus && !langkahTes) continue;

        const hasilAktual = get(row, 'hasil_aktual', '');
        const remarksRaw = get(row, 'remarks', '');
        const parsed = RemarksParser.parse(remarksRaw);

        const scenario = {
            section: currentSection,
            kategori: kategori,
            nama_modul: namaModul,
            nomor_skenario: get(row, 'nomor_skenario', ''),
            nomor_kasus_tes: nomorKasus,
            langkah_tes: langkahTes,
            aspi_no: extractAspiNumber(langkahTes),
            scenario_name: cleanScenarioName(langkahTes),
            expected_result: get(row, 'hasil_diharapkan', ''),
            hasil_aktual: hasilAktual,
            remarks_raw: remarksRaw,
            url: parsed.url,
            headers: parsed.headers,
            request_body: parsed.request_body,
            request: parsed.full_request,
            response: parsed.full_response,
            is_skipped: SKIP_STATUSES.includes(hasilAktual.toLowerCase())
        };
        scenarios.push(scenario);
    }

    return { scenarios, metadata: meta, headerRowIdx, layout, error: null };
}

function extractInlineValue(cellValue, row, colIdx) {
    if (cellValue.includes(':')) {
        const parts = cellValue.split(':');
        const rest = parts.slice(1).join(':').trim();
        if (rest) return rest;
    }
    if (row[colIdx + 1]) return String(row[colIdx + 1]).trim();
    return '';
}

// =============================================================================
// BROWSER-ONLY CODE (DOM, docx generation). Guarded so headless tests skip it.
// =============================================================================

const IS_BROWSER = typeof window !== 'undefined' && typeof document !== 'undefined';

if (IS_BROWSER) {

    // ---- Global error handlers ---------------------------------------------
    window.onerror = function (msg, url, line) {
        console.error('Error:', msg, 'at', url, ':', line);
        alert('Terjadi error: ' + msg + '\n\nLine: ' + line + '\nSilakan buka Console (F12) untuk detail.');
        return false;
    };
    window.addEventListener('unhandledrejection', function (event) {
        console.error('Unhandled promise rejection:', event.reason);
        alert('Terjadi error async: ' + ((event.reason && event.reason.message) || event.reason) + '\n\nSilakan buka Console (F12) untuk detail.');
    });

    // ---- Prevent browser default drag behavior -----------------------------
    document.addEventListener('dragover', (e) => { e.preventDefault(); e.stopPropagation(); });
    document.addEventListener('drop', (e) => { e.preventDefault(); e.stopPropagation(); });

    const dropZone = document.getElementById('dropZone');
    const fileInput = document.getElementById('fileInput');
    const fileInfo = document.getElementById('fileInfo');

    dropZone.addEventListener('click', () => fileInput.click());
    dropZone.addEventListener('dragover', (e) => {
        e.preventDefault(); e.stopPropagation();
        dropZone.classList.add('dragover');
    });
    dropZone.addEventListener('dragleave', (e) => {
        e.preventDefault(); e.stopPropagation();
        dropZone.classList.remove('dragover');
    });
    dropZone.addEventListener('drop', (e) => {
        e.preventDefault(); e.stopPropagation();
        dropZone.classList.remove('dragover');
        const file = e.dataTransfer.files[0];
        if (file) handleFile(file);
    });
    fileInput.addEventListener('change', (e) => {
        const file = e.target.files[0];
        if (file) handleFile(file);
    });

    // ---- File handling -----------------------------------------------------
    // Loading a file no longer auto-parses; it stores the data and reveals the
    // "Proses File" / "Hapus File" buttons so the user triggers processing.
    function handleFile(file) {
        if (!file.name.match(/\.xlsx?$/i)) {
            alert('Format file harus .xlsx atau .xls');
            return;
        }

        currentFileName = file.name;
        dropZone.classList.add('has-file');
        dropZone.querySelector('.drop-zone-icon').textContent = '✅';
        dropZone.querySelector('.drop-zone-text').innerHTML =
            `<strong>${file.name}</strong><br><small>File siap diproses. Klik "Proses File".</small>`;

        fileInfo.textContent = `File: ${file.name} (${(file.size / 1024).toFixed(1)} KB)`;
        fileInfo.classList.add('show');

        const reader = new FileReader();
        reader.onload = (e) => {
            currentFileData = e.target.result;
            showFileActions(true);
        };
        reader.readAsArrayBuffer(file);
    };

    // ---- Proses File button ------------------------------------------------
    function processFile() {
        if (!currentFileData) {
            alert('Belum ada file yang dimuat. Silakan upload file UAT Script Excel terlebih dahulu.');
            return;
        }
        parseExcel(currentFileData);
    };

    // ---- Hapus File button (reset everything) ------------------------------
    function resetFile() {
        currentFileData = null;
        currentFileName = "";
        parsedScenarios = [];
        skippedSet = new Set();
        metadata = {
            nama_penyedia: "Bank Sahabat Sampoerna",
            nama_layanan: "API QR MPM",
            nama_pengguna: "",
            tanggal_pengujian: ""
        };

        if (fileInput) fileInput.value = '';

        // Reset drop zone to its initial look
        dropZone.classList.remove('has-file', 'dragover');
        dropZone.querySelector('.drop-zone-icon').textContent = '📄';
        dropZone.querySelector('.drop-zone-text').innerHTML =
            '<strong>Drag & Drop</strong> file UAT Script Excel (.xlsx) di sini<br><small>atau klik untuk memilih file</small>';

        if (fileInfo) { fileInfo.textContent = ''; fileInfo.classList.remove('show'); }

        // Hide action buttons and the downstream cards
        showFileActions(false);
        const configCard = document.getElementById('configCard');
        const resultsCard = document.getElementById('resultsCard');
        const generateCard = document.getElementById('generateCard');
        if (configCard) configCard.style.display = 'none';
        if (resultsCard) resultsCard.style.display = 'none';
        if (generateCard) generateCard.style.display = 'none';

        const logArea = document.getElementById('logArea');
        if (logArea) { logArea.innerHTML = ''; logArea.style.display = 'none'; }
    };

    function showFileActions(show) {
        const actions = document.getElementById('fileActions');
        if (actions) actions.style.display = show ? 'flex' : 'none';
    }

    // ---- Excel parsing (browser entry point) -------------------------------
    function parseExcel(data) {
        const workbook = XLSX.read(data, { type: 'array' });
        const sheetName = pickWorksheet(workbook);
        const worksheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

        const result = parseWorkbookData(rows);

        if (result.error === 'no-header') {
            alert('ERROR: Tidak dapat menemukan header row di Excel.\n\n' +
                  'Format yang didukung:\n' +
                  '- Kirimo (Indonesia): Kategori, Nama Modul, Langkah Tes, Hasil yang diharapkan, Remarks, ...\n' +
                  '- Lama (English): No, Service, Scenario, Expected Result, Request, Response, Result, Notes');
            return;
        }

        parsedScenarios = result.scenarios;

        // Merge detected metadata (do not overwrite non-empty defaults with blanks)
        if (result.metadata.nama_pengguna) metadata.nama_pengguna = result.metadata.nama_pengguna;
        if (result.metadata.tanggal_pengujian) metadata.tanggal_pengujian = result.metadata.tanggal_pengujian;
        if (result.metadata.nama_penyedia) metadata.nama_penyedia = result.metadata.nama_penyedia;
        if (result.metadata.nama_layanan) metadata.nama_layanan = result.metadata.nama_layanan;

        // Auto-detect skipped scenarios (is_skipped from hasil_aktual)
        skippedSet = new Set();
        parsedScenarios.forEach(s => {
            if (s.is_skipped) skippedSet.add(s.nomor_kasus_tes || s.aspi_no);
        });

        // Populate the (now-hidden) config inputs so their .value stays in sync
        // with the auto-detected metadata. The "Konfigurasi (Opsional)" card is
        // intentionally NOT shown (per user request); we keep the elements in
        // the DOM only so updateMetadataFromUI() can still read them safely.
        // NOTE: configCard is deliberately never set to display:block here.
        setInputValue('namaPengguna', metadata.nama_pengguna || '');
        setInputValue('tanggalPengujian', metadata.tanggal_pengujian || '');
        setInputValue('skippedScenarios', Array.from(skippedSet).filter(Boolean).join(', '));

        addLog(`File diproses: ${parsedScenarios.length} skenario ditemukan (header baris ${result.headerRowIdx + 1}, layout ${result.layout}).`, 'info');

        validateAndShowResults();
    };

    // ---- Results display ---------------------------------------------------
    // The "Hasil Validasi" card (resultsCard) is intentionally NOT shown (per
    // user request). We still run the full validation + summary computation so
    // the generated documents stay correct, and we still reveal the "Generate
    // Dokumen" card so the user can produce the .docx files.
    function validateAndShowResults() {
        const resultsCard = document.getElementById('resultsCard');
        const generateCard = document.getElementById('generateCard');
        const resultsBody = document.getElementById('resultsBody');
        const summaryCards = document.getElementById('summaryCards');

        // Keep resultsCard hidden; only the Generate card is revealed.
        if (resultsCard) resultsCard.style.display = 'none';
        if (generateCard) generateCard.style.display = 'block';

        let passed = 0, failed = 0, na = 0, notTested = 0;
        if (resultsBody) resultsBody.innerHTML = '';

        parsedScenarios.forEach(s => {
            const result = validateScenario(s);
            if (result === 'PASS') passed++;
            else if (result === 'NOT PASS') failed++;
            else if (result === 'N/A') na++;
            else notTested++;

            const expectedCode = extractExpectedCode(s.expected_result) || '-';
            const actualCode = ResponseParser.extractResponseCode(s.response) || '-';

            let badgeClass = 'badge-pending';
            if (result === 'PASS') badgeClass = 'badge-pass';
            else if (result === 'NOT PASS') badgeClass = 'badge-fail';
            else if (result === 'N/A') badgeClass = 'badge-na';

            // The results table is hidden, but keep populating it (when present)
            // so nothing breaks if the card is ever re-enabled.
            if (resultsBody) {
                const name = s.scenario_name || s.langkah_tes || '';
                const noLabel = s.aspi_no || s.nomor_kasus_tes || '';
                const row = document.createElement('tr');
                row.innerHTML = `
                    <td>${escapeHtml(noLabel)}</td>
                    <td>${escapeHtml(s.nama_modul)}</td>
                    <td>${escapeHtml(name.substring(0, 50))}${name.length > 50 ? '...' : ''}</td>
                    <td><code>${escapeHtml(expectedCode)}</code></td>
                    <td><code>${escapeHtml(actualCode)}</code></td>
                    <td><span class="badge ${badgeClass}">${result}</span></td>
                    <td>${escapeHtml((s.hasil_aktual || '').substring(0, 30))}</td>
                `;
                resultsBody.appendChild(row);
            }
        });

        const total = parsedScenarios.length;
        if (summaryCards) {
            summaryCards.innerHTML = `
                <div class="summary-card total"><div class="number">${total}</div><div class="label">Total Skenario</div></div>
                <div class="summary-card pass"><div class="number">${passed}</div><div class="label">PASS</div></div>
                <div class="summary-card fail"><div class="number">${failed}</div><div class="label">NOT PASS</div></div>
                <div class="summary-card na"><div class="number">${na + notTested}</div><div class="label">N/A / Belum Diisi</div></div>
            `;
        }

        const denom = total - na - notTested;
        const passRate = denom > 0 ? (passed / denom) * 100 : 0;
        const progressFill = document.getElementById('progressFill');
        if (progressFill) progressFill.style.width = `${Math.min(passRate, 100)}%`;

        addLog(`Validasi selesai: PASS ${passed}, NOT PASS ${failed}, N/A/Belum Diisi ${na + notTested} dari ${total} skenario.`, 'info');
    };

    function escapeHtml(str) {
        return String(str == null ? '' : str)
            .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    }

    // Safe getters/setters for the (hidden) config inputs. The Konfigurasi and
    // Hasil Validasi cards are hidden per user request; these helpers make the
    // code tolerant whether the inputs are merely hidden or removed entirely.
    function setInputValue(id, value) {
        const el = document.getElementById(id);
        if (el) el.value = value;
    }
    function getInputValue(id) {
        const el = document.getElementById(id);
        return el ? el.value : '';
    }

    // ---- Helper: build docx paragraphs from a multi-line monospace block ----
    // docx.js does NOT render "\n" inside a single TextRun, so each line must be
    // its own paragraph (with a mono font). This is what makes URL / Headers /
    // Body Request appear on separate lines in the Word document.
    function monospaceParagraphs(text, sizeHalfPt, indentLeft) {
        const { Paragraph, TextRun } = docx;
        const lines = String(text || '').split('\n');
        return lines.map(line => new Paragraph({
            indent: indentLeft ? { left: indentLeft } : undefined,
            spacing: { before: 0, after: 0 },
            children: [new TextRun({ text: line, font: 'Consolas', size: sizeHalfPt })]
        }));
    }

    // Build a single table cell whose content is split across lines/paragraphs.
    function multiLineCell(text, opts) {
        const { Paragraph, TextRun, WidthType, AlignmentType } = docx;
        opts = opts || {};
        const lines = String(text == null ? '' : text).split('\n');
        const paragraphs = lines.map(line => new Paragraph({
            alignment: opts.alignment || AlignmentType.LEFT,
            spacing: { before: 0, after: 0 },
            children: [new TextRun({
                text: line,
                size: opts.size || 14,
                bold: !!opts.bold,
                color: opts.color || '000000',
                font: opts.font || 'Calibri'
            })]
        }));
        const cellOpts = { children: paragraphs.length ? paragraphs : [new Paragraph({ text: '' })] };
        if (opts.width !== undefined) cellOpts.width = { size: opts.width, type: WidthType.DXA };
        return new docx.TableCell(cellOpts);
    }

    // =========================================================================
    // DOCX GENERATION - UAT RESULT
    // =========================================================================
    async function generateUATResult() {
        addLog('Generating UAT Result...', 'info');
        updateMetadataFromUI();

        const { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType,
                Table, TableRow, TableCell, WidthType, ShadingType, PageBreak } = docx;

        const children = [];

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
                new TextRun({ text: `Nama Pengguna Layanan: ${metadata.nama_pengguna}`, size: 22, break: 1 }),
                new TextRun({ text: `Tanggal Pengujian: ${metadata.tanggal_pengujian}`, size: 22, break: 1 })
            ]
        }));
        children.push(new Paragraph({ children: [new PageBreak()] }));

        // Detail per scenario
        children.push(new Paragraph({
            heading: HeadingLevel.HEADING_1,
            children: [new TextRun({ text: 'Detail Hasil Pengujian', bold: true })]
        }));

        parsedScenarios.forEach((s, idx) => {
            const name = s.scenario_name || s.langkah_tes || '';
            const noLabel = s.aspi_no || s.nomor_kasus_tes || String(idx + 1);
            children.push(new Paragraph({
                heading: HeadingLevel.HEADING_2,
                children: [new TextRun({ text: `${noLabel} ${name}`, bold: true })]
            }));

            if (s.is_skipped) {
                children.push(new Paragraph({
                    children: [new TextRun({ text: 'Tidak dites karena tidak sesuai dengan kondisi produk.', italics: true })]
                }));
            } else {
                children.push(new Paragraph({
                    children: [
                        new TextRun({ text: 'Expected Result: ', bold: true, size: 20 }),
                        new TextRun({ text: s.expected_result, size: 20 })
                    ]
                }));

                // Request (URL + Headers + Body Request), each line its own paragraph
                children.push(new Paragraph({
                    children: [new TextRun({ text: 'Request:', bold: true, size: 20 })]
                }));
                if (s.request) {
                    monospaceParagraphs(s.request, 16, 720).forEach(p => children.push(p));
                }

                // Response
                children.push(new Paragraph({
                    children: [new TextRun({ text: 'Response:', bold: true, size: 20 })]
                }));
                if (s.response) {
                    monospaceParagraphs(s.response, 16, 720).forEach(p => children.push(p));
                }
            }
            children.push(new Paragraph({ text: '' }));
        });

        // Summary
        children.push(new Paragraph({ children: [new PageBreak()] }));
        children.push(new Paragraph({
            heading: HeadingLevel.HEADING_1,
            children: [new TextRun({ text: 'Ringkasan Hasil Pengujian', bold: true })]
        }));

        let passed = 0, failed = 0, naCount = 0, notTested = 0;
        parsedScenarios.forEach(s => {
            const r = validateScenario(s);
            if (r === 'PASS') passed++;
            else if (r === 'NOT PASS') failed++;
            else if (r === 'N/A') naCount++;
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
        const summaryRows = summaryData.map((row, i) => new TableRow({
            children: row.map(cellText => new TableCell({
                children: [new Paragraph({
                    children: [new TextRun({
                        text: cellText, bold: i === 0, size: 20,
                        color: i === 0 ? 'FFFFFF' : '000000'
                    })]
                })],
                shading: i === 0 ? { type: ShadingType.SOLID, color: '4472C4' } : undefined,
                width: { size: 4000, type: WidthType.DXA }
            }))
        }));
        children.push(new Table({ rows: summaryRows }));

        const doc = new Document({ sections: [{ children: children }] });
        const blob = await Packer.toBlob(doc);
        const timestamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
        saveAs(blob, `UAT_Result_QRIS_Merchant_Aggregator_${timestamp}.docx`);
        addLog('UAT Result berhasil di-generate!', 'success');
    };

    // =========================================================================
    // DOCX GENERATION - LAMPIRAN 7C
    // =========================================================================
    async function generateLampiran7C() {
        addLog('Generating Lampiran 7C...', 'info');
        updateMetadataFromUI();

        const { Document, Packer, Paragraph, TextRun, AlignmentType,
                Table, TableRow, TableCell, WidthType, ShadingType, PageOrientation,
                TableLayoutType } = docx;

        const children = [];

        // Proportional column widths (twips) for the 8 columns
        // [No, Service, Scenario, Expected Result, Request, Response, Result, Notes].
        // Request/Response are the widest; No/Result are narrow. The total (15340)
        // fits comfortably inside the A3 landscape text area emitted below
        // (page 23811 twips wide, 680 twip side margins => 22451 usable, ~7111 twips
        // of slack) so the fixed tblGrid is honored by Word without collapse.
        const COL_WIDTHS = [520, 1350, 2150, 2000, 3350, 3350, 720, 1900];
        const TABLE_WIDTH = COL_WIDTHS.reduce((a, b) => a + b, 0); // 15340 twips

        children.push(new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: 'Lampiran 7.C', bold: true, size: 28 })]
        }));
        children.push(new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [new TextRun({ text: 'Skenario dan Hasil Uji Fungsionalitas', bold: true, size: 24 })]
        }));
        children.push(new Paragraph({ text: '' }));

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

        const headerTexts = ['No', 'Service', 'Scenario', 'Expected Result', 'Request', 'Response', 'Result', 'Notes'];
        const headerRow = new TableRow({
            children: headerTexts.map((h, i) => new TableCell({
                children: [new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [new TextRun({ text: h, bold: true, size: 16, color: 'FFFFFF' })]
                })],
                shading: { type: ShadingType.SOLID, color: '4472C4' },
                width: { size: COL_WIDTHS[i], type: WidthType.DXA }
            }))
        });

        const dataRows = parsedScenarios.map((s, idx) => {
            let resultText, notesText;
            if (s.is_skipped) {
                resultText = 'N/A';
                notesText = s.hasil_aktual || 'Tidak dites karena tidak sesuai dengan kondisi produk.';
            } else {
                resultText = validateScenario(s);
                notesText = s.hasil_aktual || '';
            }
            const resultColor = resultText === 'PASS' ? '008000' : (resultText === 'NOT PASS' ? 'FF0000' : '666666');

            const noLabel = s.aspi_no || s.nomor_kasus_tes || String(idx + 1);
            const name = s.scenario_name || s.langkah_tes || '';

            return new TableRow({
                children: [
                    multiLineCell(noLabel, { width: COL_WIDTHS[0] }),
                    multiLineCell(s.nama_modul, { width: COL_WIDTHS[1] }),
                    multiLineCell(name, { width: COL_WIDTHS[2] }),
                    multiLineCell(s.expected_result, { width: COL_WIDTHS[3] }),
                    multiLineCell(s.request || '', { width: COL_WIDTHS[4], font: 'Consolas' }),
                    multiLineCell(s.response || '', { width: COL_WIDTHS[5], font: 'Consolas' }),
                    multiLineCell(resultText, { width: COL_WIDTHS[6], bold: true, color: resultColor, alignment: AlignmentType.CENTER }),
                    multiLineCell(notesText, { width: COL_WIDTHS[7] })
                ]
            });
        });

        // ONE consistent width strategy: explicit DXA table width + explicit
        // columnWidths (so docx.js emits a matching tblGrid) + fixed layout so
        // Word honors the grid instead of autofitting to a degenerate ~100-twip grid.
        children.push(new Table({
            rows: [headerRow, ...dataRows],
            width: { size: TABLE_WIDTH, type: WidthType.DXA },
            columnWidths: COL_WIDTHS,
            layout: TableLayoutType.FIXED
        }));

        children.push(new Paragraph({ text: '' }));
        const footerNotes = [
            'Lampiran Skenario hasil uji fungsional sekurangnya 1 Pengguna Layanan atas 1 sub API unverified, dengan ketentuan sebagai berikut:',
            '',
            'a. Pada kolom request diisi dengan request yang dilakukan Pengguna layanan, sedangkan pada kolom response diisi dengan respon yang diberikan Penyedia. Sementara pada kolom result diisi dengan hasil PASS atau NOT PASS yang harus sesuai dengan expected result.',
            '',
            'b. Pengisian pada dokumen skenario hasil uji fungsional tidak dilakukan dengan cara screen capture, melainkan dilakukan dengan cara copy paste payload request dan response dari log API server ke kolom tabel skenario hasil uji fungsional.',
            '',
            'c. Seluruh skenario diujikan dan tidak boleh dihapus atau diubah. Dalam hal terdapat skenario yang tidak diujikan dapat dikosongkan pengisiannya, namun diberikan catatan pada kolom Notes yang akan kami review lebih lanjut apakah skenario diperkenankan untuk tidak diujikan.',
            'd. Dalam hal terdapat penambahan skenario pengujian, maka penambahan tersebut dilakukan pada baris paling bawah, sehingga tidak mengubah susunan atau urutan template skenario.'
        ];
        footerNotes.forEach(note => {
            children.push(new Paragraph({
                indent: { left: 720 },
                children: [new TextRun({ text: note, size: 16, italics: true })]
            }));
        });

        const doc = new Document({
            sections: [{
                properties: {
                    page: {
                        // A3 landscape. NOTE: under PageOrientation.LANDSCAPE docx.js
                        // swaps the pair so the LARGER value becomes the emitted page
                        // width (w:w) and the smaller becomes the height (w:h). To get
                        // the wide A3 long side (23811) as the actual page width we must
                        // pass width=16838 / height=23811 here; docx.js then emits
                        // <w:pgSz w:w="23811" w:h="16838" w:orient="landscape"/>.
                        // Usable width = 23811 - (680 + 680) = 22451 twips, so the
                        // 15340-twip table has ~7111 twips (~12.5 cm) of headroom.
                        size: {
                            orientation: PageOrientation.LANDSCAPE,
                            width: 16838,
                            height: 23811
                        },
                        margin: { top: 720, right: 680, bottom: 720, left: 680 }
                    }
                },
                children: children
            }]
        });

        const blob = await Packer.toBlob(doc);
        const timestamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
        saveAs(blob, `Lampiran_7C_QRIS_${timestamp}.docx`);
        addLog('Lampiran 7C berhasil di-generate!', 'success');
    };

    // ---- Utilities ---------------------------------------------------------
    async function generateBoth() {
        await generateUATResult();
        await generateLampiran7C();
        addLog('Kedua dokumen berhasil di-generate!', 'success');
    };

    function updateMetadataFromUI() {
        // The config inputs live inside the hidden "Konfigurasi (Opsional)" card.
        // Read them defensively and fall back to the auto-detected metadata (from
        // the Excel) whenever an input is missing or left blank.
        metadata.nama_penyedia = getInputValue('namaPenyedia') || metadata.nama_penyedia;
        metadata.nama_pengguna = getInputValue('namaPengguna') || metadata.nama_pengguna;
        metadata.nama_layanan = getInputValue('namaLayanan') || metadata.nama_layanan;
        metadata.tanggal_pengujian = getInputValue('tanggalPengujian') || metadata.tanggal_pengujian;

        const skippedInput = getInputValue('skippedScenarios');
        if (skippedInput.trim()) {
            const additional = skippedInput.split(',').map(s => s.trim()).filter(Boolean);
            additional.forEach(no => {
                skippedSet.add(no);
                // Reflect the manual skip onto matching scenarios so validation follows
                parsedScenarios.forEach(s => {
                    if ((s.nomor_kasus_tes === no) || (s.aspi_no === no)) s.is_skipped = true;
                });
            });
        }
    };

    function addLog(message, type) {
        const logArea = document.getElementById('logArea');
        logArea.style.display = 'block';
        const timestamp = new Date().toLocaleTimeString('id-ID');
        const className = type === 'success' ? 'log-success' : (type === 'error' ? 'log-error' : 'log-info');
        logArea.innerHTML += `<div class="${className}">[${timestamp}] ${message}</div>`;
        logArea.scrollTop = logArea.scrollHeight;
    };

    // ---- Expose functions on window ----------------------------------------
    // The generate buttons in index.html use inline onclick="generateUATResult()"
    // etc., which resolve against the global (window) scope, so expose them.
    window.handleFile = handleFile;
    window.processFile = processFile;
    window.resetFile = resetFile;
    window.parseExcel = parseExcel;
    window.validateAndShowResults = validateAndShowResults;
    window.generateUATResult = generateUATResult;
    window.generateLampiran7C = generateLampiran7C;
    window.generateBoth = generateBoth;
    window.updateMetadataFromUI = updateMetadataFromUI;
    window.addLog = addLog;

    // ---- Wire the index-page buttons (Proses File / Hapus File) ------------
    const btnProses = document.getElementById('btnProses');
    const btnHapus = document.getElementById('btnHapus');
    if (btnProses) btnProses.addEventListener('click', processFile);
    if (btnHapus) btnHapus.addEventListener('click', resetFile);
}

// =============================================================================
// HEADLESS EXPORT (guarded) - lets bun/node import the pure parsing functions
// without touching the DOM. Browsers load this file via <script>, where
// module is undefined, so this block is skipped.
// =============================================================================
if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        RemarksParser,
        ResponseParser,
        cleanScenarioName,
        extractAspiNumber,
        extractExpectedCode,
        extractResponseCode,
        validateScenario,
        findHeaderRow,
        pickWorksheet,
        parseWorkbookData,
        SKIP_STATUSES
    };
}
