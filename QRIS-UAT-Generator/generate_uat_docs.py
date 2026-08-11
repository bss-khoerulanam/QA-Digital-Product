#!/usr/bin/env python3
"""
QRIS Merchant Aggregator - UAT Document Generator
Bank Sahabat Sampoerna (BSS)

Tool ini membaca UAT Script Excel yang sudah diisi oleh mitra,
lalu menghasilkan:
1. UAT Result (.docx) - Dokumen hasil pengujian detail
2. Lampiran 7C (.docx) - Berita Acara untuk ASPI Portal

Format input: UAT Script Excel dari mitra dengan kolom:
- Kategori, Nama Modul, Nomor Skenario, Nomor Kasus Tes,
  Langkah Tes, Hasil yang diharapkan, Hasil Aktual, Remarks

Kolom Remarks berisi data request & response (URL, Headers, Body, Response)

Author: IT QA BSS
"""

import re
import sys
import os
from datetime import datetime

from openpyxl import load_workbook
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import nsdecls
from docx.oxml import parse_xml


# =============================================================================
# CONFIGURATION
# =============================================================================

SKIP_REASON = "Tidak dites karena tidak sesuai dengan kondisi produk."

# Status yang dianggap "tidak dites"
SKIP_STATUSES = ["tidak dites", "tidak ditest"]



# =============================================================================
# REMARKS PARSER - Memisahkan URL, Headers, Request Body, Response
# =============================================================================

class RemarksParser:
    """
    Parse kolom Remarks dari UAT Script mitra.
    
    Format yang diharapkan di kolom Remarks:
    
    URL:
    POST /snap-qris/v1.1/qr/qr-mpm-generate HTTP/1.1
    Host: ob-sandbox.banksampoerna.co.id
    ...
    Headers:
    Authorization: Bearer xxx
    Content-Type: application/json
    ...
    Request Body:
    {"field": "value", ...}
    
    Response:
    HTTP/1.1 200
    Content-Type: application/json
    ...
    {"responseCode":"2004700",...}
    """

    @staticmethod
    def parse(remarks_text):
        """
        Parse remarks text into structured components.
        Returns dict with keys: url, headers, request_body, response
        """
        result = {
            "url": "",
            "headers": "",
            "request_body": "",
            "response": "",
            "full_request": "",  # Combined URL + Headers + Body for Lampiran 7C
            "full_response": "", # Full response for Lampiran 7C
        }

        if not remarks_text or remarks_text.strip() == "":
            return result

        text = remarks_text.strip()

        # Detect format variants from mitra
        # Format 1: "URL:\n..." or "Request:\n..."
        # Format 2: starts directly with "POST /..." or "GET /..."
        # Format 3: Headers first (e.g., notification format)

        # Try to split into Request and Response sections
        request_part, response_part = RemarksParser._split_request_response(text)

        # Parse request part
        if request_part:
            url, headers, body = RemarksParser._parse_request_section(request_part)
            result["url"] = url
            result["headers"] = headers
            result["request_body"] = body
            result["full_request"] = RemarksParser._format_request_output(url, headers, body)

        # Parse response part
        if response_part:
            result["response"] = response_part.strip()
            result["full_response"] = response_part.strip()

        return result

    @staticmethod
    def _split_request_response(text):
        """Split text into request and response sections."""
        # Look for "Response:" or "Response:\n" separator
        # Also handle "HTTP/1.1 XXX" as response start after a blank line
        
        # Pattern 1: Explicit "Response:" label
        response_markers = [
            r'\nResponse:\s*\n',
            r'\nResponse:\s*$',
            r'^Response:\s*\n',
        ]
        
        for pattern in response_markers:
            match = re.search(pattern, text, re.MULTILINE)
            if match:
                request_part = text[:match.start()].strip()
                response_part = text[match.end():].strip()
                return request_part, response_part

        # Pattern 2: Look for HTTP response line after request body (JSON followed by HTTP/1.1)
        # Find the boundary between request JSON and response HTTP status
        json_then_http = re.search(
            r'(\})\s*\n\s*\n*(HTTP/\d\.\d\s+\d+)',
            text, re.MULTILINE
        )
        if json_then_http:
            split_pos = json_then_http.start(2)
            request_part = text[:split_pos].strip()
            response_part = text[split_pos:].strip()
            return request_part, response_part

        # Pattern 3: No clear response section - treat entire text as request
        # Check if text contains HTTP response pattern anywhere
        http_response = re.search(r'\n(HTTP/\d\.\d\s+\d+\s*\n)', text)
        if http_response:
            request_part = text[:http_response.start()].strip()
            response_part = text[http_response.start():].strip()
            return request_part, response_part

        # No response found - everything is request (e.g., notification scenario)
        return text, ""

    @staticmethod
    def _parse_request_section(request_text):
        """Parse request section into URL, Headers, and Body."""
        url = ""
        headers = ""
        body = ""

        lines = request_text.split('\n')
        
        # Remove leading labels like "URL:", "Request:"
        start_idx = 0
        if lines and re.match(r'^(URL|Request)\s*:\s*$', lines[0].strip(), re.IGNORECASE):
            start_idx = 1
        
        # Find HTTP method line (POST /path HTTP/1.1 or GET /path HTTP/1.1)
        http_method_idx = -1
        for i in range(start_idx, len(lines)):
            if re.match(r'^(POST|GET|PUT|DELETE|PATCH)\s+/', lines[i].strip()):
                http_method_idx = i
                break

        if http_method_idx >= 0:
            # URL is the HTTP method line
            url_line = lines[http_method_idx].strip()
            
            # Find Host header to construct full URL
            host = ""
            for i in range(http_method_idx + 1, len(lines)):
                if lines[i].strip().lower().startswith('host:'):
                    host = lines[i].strip().split(':', 1)[1].strip()
                    break

            # Extract method and path from "POST /path HTTP/1.1"
            method_match = re.match(r'(POST|GET|PUT|DELETE|PATCH)\s+(\S+)', url_line)
            if method_match and host:
                method = method_match.group(1)
                path = method_match.group(2)
                url = f"URL:\n{method} https://{host}{path}"
            else:
                url = f"URL:\n{url_line}"

            # Collect ALL lines after HTTP method line as headers until we hit
            # an empty line or a JSON body (starts with {)
            header_lines = []
            body_start_idx = -1
            
            for i in range(http_method_idx + 1, len(lines)):
                line = lines[i].strip()
                
                # Skip Host line (already included in URL)
                if line.lower().startswith('host:'):
                    continue
                
                # Skip HTTP version info (e.g., "User-Agent: Go-http-client/1.1")
                # These are still headers, include them
                
                # Empty line = separator between headers and body
                if line == '':
                    body_start_idx = i + 1
                    break
                
                # JSON body starts
                if line.startswith('{'):
                    body_start_idx = i
                    break
                
                # Otherwise it's a header line (Key: Value or Key-Name: Value)
                if re.match(r'^[\w-]+[\w-]*\s*:', line):
                    header_lines.append(line)
            
            # Extract body (everything from body_start_idx to end)
            if body_start_idx >= 0 and body_start_idx < len(lines):
                body_text = '\n'.join(lines[body_start_idx:]).strip()
                # Find the JSON object in body
                if body_text:
                    json_start = body_text.find('{')
                    if json_start >= 0:
                        body_text = body_text[json_start:]
                        # Find end of JSON
                        brace_count = 0
                        json_end = -1
                        for i, ch in enumerate(body_text):
                            if ch == '{':
                                brace_count += 1
                            elif ch == '}':
                                brace_count -= 1
                                if brace_count == 0:
                                    json_end = i + 1
                                    break
                        if json_end > 0:
                            body = body_text[:json_end]
                        else:
                            body = body_text
            
            headers = '\n'.join(header_lines).strip()

        else:
            # No HTTP method line found - might be notification format
            # (starts directly with headers like Authorization:, Content-Type:, etc.)
            header_lines = []
            body_start = -1
            
            for i in range(start_idx, len(lines)):
                line = lines[i].strip()
                if line.startswith('{'):
                    body_start = i
                    break
                elif line == '':
                    # Empty line might mean body follows
                    continue
                elif re.match(r'^[\w-]+[\w-]*\s*:', line):
                    header_lines.append(line)

            headers = '\n'.join(header_lines).strip()
            if body_start >= 0:
                body_text = '\n'.join(lines[body_start:]).strip()
                # Find end of JSON
                brace_count = 0
                json_end = -1
                for i, ch in enumerate(body_text):
                    if ch == '{':
                        brace_count += 1
                    elif ch == '}':
                        brace_count -= 1
                        if brace_count == 0:
                            json_end = i + 1
                            break
                if json_end > 0:
                    body = body_text[:json_end]
                else:
                    body = body_text

        return url, headers, body

    @staticmethod
    def _format_request_output(url, headers, body):
        """Format request for Lampiran 7C output column."""
        parts = []
        if url:
            parts.append(url)
        if headers:
            parts.append(f"\nHeaders:\n{headers}")
        if body:
            parts.append(f"\nRequest Body:\n{body}")
        return '\n'.join(parts).strip()



# =============================================================================
# RESPONSE PARSER - Extract response body/JSON from full HTTP response
# =============================================================================

class ResponseParser:
    """Parse full HTTP response to extract just the response body."""

    @staticmethod
    def extract_response_body(response_text):
        """
        Extract JSON response body from full HTTP response.
        Input: Full HTTP response including status line and headers
        Output: Just the JSON body
        """
        if not response_text:
            return ""

        # If it's already just JSON, return as-is
        stripped = response_text.strip()
        if stripped.startswith('{'):
            return stripped

        # Find JSON body in the response (last JSON object)
        lines = stripped.split('\n')
        json_start = -1
        for i, line in enumerate(lines):
            if line.strip().startswith('{'):
                json_start = i
                break

        if json_start >= 0:
            json_text = '\n'.join(lines[json_start:]).strip()
            # Find end of JSON
            brace_count = 0
            json_end = -1
            for i, ch in enumerate(json_text):
                if ch == '{':
                    brace_count += 1
                elif ch == '}':
                    brace_count -= 1
                    if brace_count == 0:
                        json_end = i + 1
                        break
            if json_end > 0:
                return json_text[:json_end]
            return json_text

        return response_text

    @staticmethod
    def extract_response_code(response_text):
        """Extract responseCode from response text."""
        if not response_text:
            return None
        patterns = [
            r'"responseCode"\s*:\s*"(\d+)"',
            r'"responseCode"\s*:\s*(\d+)',
            r'responseCode.*?(\d{7})',
        ]
        for pattern in patterns:
            match = re.search(pattern, response_text)
            if match:
                return match.group(1)
        return None

    @staticmethod
    def format_response_output(response_text):
        """Format response for Lampiran 7C output - show full response."""
        if not response_text:
            return ""
        return response_text.strip()



# =============================================================================
# EXCEL PARSER - Handles real mitra UAT Script format
# =============================================================================

class UATScriptParser:
    """
    Parser untuk membaca UAT Script Excel yang diisi mitra.
    
    Kolom yang diharapkan:
    - Kategori
    - Nama Modul
    - Nomor Skenario
    - Nomor Kasus Tes
    - Langkah Tes
    - Hasil yang diharapkan
    - Hasil Aktual
    - Remarks (berisi URL, Headers, Body, Response)
    - Tanggal Pelaksanaan (opsional)
    - Jenis Script (opsional)
    - Pelaksana (opsional)
    """

    # Column mapping - will be detected dynamically
    COL_KATEGORI = 0
    COL_NAMA_MODUL = 1
    COL_NOMOR_SKENARIO = 2
    COL_NOMOR_KASUS_TES = 3
    COL_LANGKAH_TES = 4
    COL_HASIL_DIHARAPKAN = 5
    COL_HASIL_AKTUAL = 6
    COL_REMARKS = 7
    COL_TANGGAL = 8
    COL_JENIS_SCRIPT = 9
    COL_PELAKSANA = 10

    def __init__(self, excel_path):
        self.excel_path = excel_path
        self.wb = load_workbook(excel_path, data_only=True)
        self.scenarios = []
        self.metadata = {
            "nama_penyedia": "Bank Sahabat Sampoerna",
            "nama_layanan": "API QR MPM",
            "nama_pengguna": "",
            "tanggal_pengujian": "",
            "nomor_referensi": "",
        }

    def parse(self):
        """Parse the UAT Script Excel file."""
        # Find the main sheet (first sheet or 'UAT Script')
        ws = self.wb.active
        print(f"      -> Sheets tersedia: {self.wb.sheetnames}")
        
        for name in self.wb.sheetnames:
            if 'uat' in name.lower() and 'script' in name.lower():
                ws = self.wb[name]
                break
            elif name.lower().strip() == 'uat script':
                ws = self.wb[name]
                break
            elif 'script' in name.lower() and 'error' not in name.lower():
                ws = self.wb[name]
                break

        print(f"      -> Menggunakan sheet: '{ws.title}' (max_row={ws.max_row}, max_col={ws.max_column})")

        # Extract metadata from header area
        self._extract_metadata(ws)

        # Find the header row with column names
        header_row, col_map = self._find_header_row(ws)

        if header_row is None:
            print("ERROR: Tidak dapat menemukan header row di Excel.")
            print("       Mencari kolom: Kategori, Nama Modul, Langkah Tes, dll.")
            print("       Tips: Pastikan sheet yang benar dipilih dan header row ada.")
            # Debug: show first 30 rows content
            print("       DEBUG - Isi row 1-30:")
            for r in range(1, min(31, ws.max_row + 1)):
                vals = []
                for c in range(1, min(12, ws.max_column + 1)):
                    v = ws.cell(row=r, column=c).value
                    if v:
                        vals.append(f"C{c}:{str(v)[:30]}")
                if vals:
                    print(f"         Row {r}: {vals}")
            return [], self.metadata

        # Parse scenario rows
        current_section = ""
        for row_idx in range(header_row + 1, ws.max_row + 1):
            row_data = []
            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                row_data.append(cell.value)

            # Skip completely empty rows
            if all(v is None or str(v).strip() == "" for v in row_data):
                continue

            # Detect section headers (e.g., "Balance Services", "QR MPM", etc.)
            first_col = str(row_data[col_map.get('kategori', 0)] or '').strip()
            second_col = str(row_data[col_map.get('nama_modul', 1)] or '').strip()
            
            # Section header detection - row with bold section name
            section_keywords = [
                "Balance Services", "API Transaction History",
                "QR MPM", "PENGECEKAN MUTASI", "Generate QR SNAP",
                "Refund Payment", "Query Payment", "Inquiry"
            ]
            
            is_section_header = False
            for kw in section_keywords:
                if kw.lower() in first_col.lower() or kw.lower() in second_col.lower():
                    # Check if this row has no test case data
                    langkah = row_data[col_map.get('langkah_tes', 4)] if col_map.get('langkah_tes', 4) < len(row_data) else None
                    if not langkah or str(langkah).strip() == "":
                        current_section = first_col or second_col
                        is_section_header = True
                        break

            if is_section_header:
                continue

            # Parse scenario row
            nomor_kasus = str(row_data[col_map.get('nomor_kasus_tes', 3)] or '').strip()
            langkah_tes = str(row_data[col_map.get('langkah_tes', 4)] or '').strip()

            # Skip rows without test case number or step
            if not nomor_kasus and not langkah_tes:
                continue

            # Get actual result and remarks
            hasil_aktual = str(row_data[col_map.get('hasil_aktual', 6)] or '').strip()
            remarks_raw = str(row_data[col_map.get('remarks', 7)] or '').strip()

            # Determine ASPI scenario number from Langkah Tes
            # Format: "18,1 Access Token Invalid" -> extract "18.1"
            aspi_no = self._extract_aspi_number(langkah_tes)

            # Parse remarks to extract request/response
            parsed_remarks = RemarksParser.parse(remarks_raw)

            scenario = {
                "section": current_section,
                "kategori": first_col,
                "nama_modul": str(row_data[col_map.get('nama_modul', 1)] or '').strip(),
                "nomor_skenario": str(row_data[col_map.get('nomor_skenario', 2)] or '').strip(),
                "nomor_kasus_tes": nomor_kasus,
                "langkah_tes": langkah_tes,
                "aspi_no": aspi_no,
                "expected_result": str(row_data[col_map.get('hasil_diharapkan', 5)] or '').strip(),
                "hasil_aktual": hasil_aktual,
                "remarks_raw": remarks_raw,
                # Parsed request/response
                "url": parsed_remarks["url"],
                "headers": parsed_remarks["headers"],
                "request_body": parsed_remarks["request_body"],
                "request": parsed_remarks["full_request"],
                "response": parsed_remarks["full_response"],
                # For validation
                "is_skipped": hasil_aktual.lower() in SKIP_STATUSES,
                "is_tested": hasil_aktual.lower() == "berhasil" or 
                            (remarks_raw != "" and hasil_aktual.lower() not in SKIP_STATUSES),
            }

            self.scenarios.append(scenario)

        print(f"      -> Parsed {len(self.scenarios)} skenario")
        return self.scenarios, self.metadata

    def _extract_metadata(self, ws):
        """Extract metadata from header rows."""
        for row_idx in range(1, min(20, ws.max_row + 1)):
            for col_idx in range(1, min(15, ws.max_column + 1)):
                cell = ws.cell(row=row_idx, column=col_idx)
                if cell.value and isinstance(cell.value, str):
                    val = cell.value.strip()
                    
                    # Look for project name in angle brackets
                    if '<' in val and '>' in val:
                        match = re.search(r'<(.+?)>', val)
                        if match:
                            project_name = match.group(1)
                            # Extract mitra name from project name
                            # e.g., "Penambahan Layanan QRIS Merchant Aggregator PT Sender Integrasi Digital (Kirimo)"
                            mitra_match = re.search(r'((?:PT|CV)\s+[\w\s]+(?:\([^)]+\))?)', project_name)
                            if mitra_match:
                                self.metadata["nama_pengguna"] = mitra_match.group(0).strip()
                    
                    if "Nomor Referensi" in val:
                        next_cell = ws.cell(row=row_idx, column=col_idx + 1)
                        if next_cell.value:
                            self.metadata["nomor_referensi"] = str(next_cell.value).strip()
                    
                    if "Tanggal" in val and "Pelaksanaan" not in val:
                        next_cell = ws.cell(row=row_idx, column=col_idx + 1)
                        if next_cell.value:
                            self.metadata["tanggal_pengujian"] = str(next_cell.value).strip()

    def _find_header_row(self, ws):
        """Find the header row and create column mapping."""
        col_map = {}
        
        # Scan up to row 50 to find header (file may have lots of metadata at top)
        for row_idx in range(1, min(50, ws.max_row + 1)):
            row_values = []
            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                row_values.append(str(cell.value or '').strip().lower())

            # Check if this row contains expected column headers
            # Use first line of each cell (cells may contain comment artifacts after \n)
            first_lines = [v.split('\n')[0].strip() if v else '' for v in row_values]
            
            has_kategori = any('kategori' in fl for fl in first_lines)
            has_nama_modul = any('nama modul' in fl for fl in first_lines)
            has_langkah = any('langkah tes' in fl or fl == 'langkah tes' for fl in first_lines)
            has_hasil = any('hasil' in fl and 'diharapkan' in fl for fl in first_lines)
            has_nomor_kasus = any('nomor kasus' in fl or 'kasus tes' in fl for fl in first_lines)

            # Primary detection: must have "Langkah Tes" + at least one other
            if has_langkah and (has_kategori or has_nama_modul or has_hasil or has_nomor_kasus):
                # Map columns - check first line of cell value for header name
                for idx, v in enumerate(row_values):
                    # Get just the first line (before any newline/comment artifacts)
                    first_line = v.split('\n')[0].strip() if v else ''
                    
                    if first_line == 'kategori' or (first_line.startswith('kategori') and 'nama' not in first_line):
                        col_map['kategori'] = idx
                    elif 'nama modul' in first_line:
                        col_map['nama_modul'] = idx
                    elif 'nomor skenario' in first_line:
                        col_map['nomor_skenario'] = idx
                    elif 'nomor kasus' in first_line or 'kasus tes' in first_line:
                        col_map['nomor_kasus_tes'] = idx
                    elif 'langkah' in first_line:
                        col_map['langkah_tes'] = idx
                    elif 'hasil' in first_line and 'diharapkan' in first_line:
                        col_map['hasil_diharapkan'] = idx
                    elif 'hasil aktual' in first_line:
                        col_map['hasil_aktual'] = idx
                    elif 'remark' in first_line:
                        col_map['remarks'] = idx
                    elif 'tanggal' in first_line and 'pelaksanaan' in first_line:
                        col_map['tanggal'] = idx
                    elif 'jenis' in first_line and 'script' in first_line:
                        col_map['jenis_script'] = idx
                    elif 'pelaksana' in first_line:
                        col_map['pelaksana'] = idx

                print(f"      -> Header row ditemukan di baris {row_idx}")
                print(f"      -> Kolom terdeteksi: {list(col_map.keys())}")
                return row_idx, col_map

        # Fallback - try simpler detection (look for "Langkah Tes" anywhere)
        for row_idx in range(1, min(50, ws.max_row + 1)):
            for col_idx in range(1, ws.max_column + 1):
                cell = ws.cell(row=row_idx, column=col_idx)
                val = str(cell.value or '').strip().lower()
                if 'langkah tes' in val:
                    # Found it - build col_map from this row
                    for c in range(1, ws.max_column + 1):
                        cv = str(ws.cell(row=row_idx, column=c).value or '').strip().lower()
                        if 'kategori' in cv:
                            col_map['kategori'] = c - 1
                        elif 'nama modul' in cv:
                            col_map['nama_modul'] = c - 1
                        elif 'nomor skenario' in cv:
                            col_map['nomor_skenario'] = c - 1
                        elif 'nomor kasus' in cv or 'kasus tes' in cv:
                            col_map['nomor_kasus_tes'] = c - 1
                        elif 'langkah' in cv:
                            col_map['langkah_tes'] = c - 1
                        elif 'hasil' in cv and 'diharapkan' in cv:
                            col_map['hasil_diharapkan'] = c - 1
                        elif 'hasil aktual' in cv:
                            col_map['hasil_aktual'] = c - 1
                        elif 'remark' in cv:
                            col_map['remarks'] = c - 1
                    print(f"      -> Header row ditemukan (fallback) di baris {row_idx}")
                    return row_idx, col_map

        return None, {}

    def _extract_aspi_number(self, langkah_tes):
        """Extract ASPI scenario number from Langkah Tes field."""
        if not langkah_tes:
            return ""
        # Match patterns like "18,1" or "18.1" at the beginning
        match = re.match(r'(\d+)[,.](\d+)', langkah_tes)
        if match:
            return f"{match.group(1)}.{match.group(2)}"
        return ""



# =============================================================================
# PASS/FAIL VALIDATOR
# =============================================================================

class ResultValidator:
    """Validasi hasil berdasarkan response code vs expected result."""

    @staticmethod
    def extract_expected_code(expected_text):
        """Extract expected response/error code from expected result text."""
        if not expected_text:
            return None
        patterns = [
            r'[Rr]esponse\s*[Cc]ode[:\s]*(\d{7})',
            r'[Ee]rror\s*[Cc]ode[:\s]*(\d+xx\d+)',
            r'[Ee]rror\s*[Cc]ode[:\s]*(\d{7})',
            r'(\d{7})',
            r'(\d{3}xx\d{2})',
        ]
        for pattern in patterns:
            match = re.search(pattern, expected_text)
            if match:
                return match.group(1)
        return None

    @staticmethod
    def validate(scenario):
        """
        Validate if response matches expected result.
        Kriteria: responseCode di response HARUS sama dengan expected result.
        Returns: 'PASS', 'NOT PASS', 'N/A', atau 'NOT TESTED'
        """
        # If skipped
        if scenario.get("is_skipped"):
            return "N/A"

        # If no response data
        if not scenario.get("response") or scenario["response"].strip() == "":
            # Check if hasil_aktual says "Berhasil"
            if scenario.get("hasil_aktual", "").lower() == "berhasil":
                return "NOT TESTED"  # Has result but no evidence
            return "NOT TESTED"

        # Extract codes
        expected_code = ResultValidator.extract_expected_code(scenario.get("expected_result", ""))
        actual_code = ResponseParser.extract_response_code(scenario.get("response", ""))

        if not expected_code or not actual_code:
            # Can't validate - check hasil_aktual
            if scenario.get("hasil_aktual", "").lower() == "berhasil":
                return "PASS"
            return "NOT TESTED"

        # Handle xx pattern (e.g., 401xx01)
        if "xx" in expected_code:
            pattern = expected_code.replace("xx", r"\d{2}")
            if re.match(pattern, actual_code):
                return "PASS"
            else:
                return "NOT PASS"

        # Direct comparison
        if expected_code == actual_code:
            return "PASS"
        else:
            return "NOT PASS"



# =============================================================================
# DOCUMENT STYLING HELPERS
# =============================================================================

def set_cell_shading(cell, color):
    """Set background color of a table cell."""
    shading_elm = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{color}"/>')
    cell._tc.get_or_add_tcPr().append(shading_elm)


def set_cell_border(cell):
    """Set border for a table cell."""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        f'<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        f'<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        f'<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(tcBorders)


def create_table_with_borders(doc, rows, cols):
    """Create a table with all borders."""
    table = doc.add_table(rows=rows, cols=cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for row in table.rows:
        for cell in row.cells:
            set_cell_border(cell)
    return table


def add_cell_text(cell, text, font_name='Calibri', font_size=Pt(8),
                  bold=False, color=None, alignment=None):
    """Add formatted text to a table cell."""
    cell.text = ""
    p = cell.paragraphs[0]
    if alignment:
        p.alignment = alignment
    run = p.add_run(str(text) if text else "")
    run.font.name = font_name
    run.font.size = font_size
    run.bold = bold
    if color:
        run.font.color.rgb = color
    return run



# =============================================================================
# LAMPIRAN 7C (BERITA ACARA) DOCUMENT GENERATOR
# =============================================================================

class Lampiran7CGenerator:
    """
    Generate Lampiran 7C - Berita Acara for ASPI Portal (.docx).
    
    Format output sesuai template ASPI:
    Kolom: No | Service | Scenario | Expected Result | Request | Response | Result | Notes
    
    Kolom Request berisi:
        URL:
        [method] [full_url]
        
        Headers:
        [header lines]
        
        Request Body:
        [JSON body]
    
    Kolom Response berisi:
        Response:
        [full HTTP response with headers and body]
    """

    def __init__(self, scenarios, metadata):
        self.scenarios = scenarios
        self.metadata = metadata
        self.doc = Document()
        self.validator = ResultValidator()

    def generate(self, output_path):
        """Generate the Lampiran 7C document."""
        self._set_styles()
        self._add_header()
        self._add_scenario_table()
        self._add_footer_notes()

        self.doc.save(output_path)
        print(f"[OK] Lampiran 7C saved: {output_path}")

    def _set_styles(self):
        """Set default document styles."""
        style = self.doc.styles['Normal']
        font = style.font
        font.name = 'Calibri'
        font.size = Pt(9)

        # Set page to landscape for table readability
        section = self.doc.sections[0]
        section.orientation = WD_ORIENT.LANDSCAPE
        section.page_width = Cm(42.0)  # A3 width for wide table
        section.page_height = Cm(29.7)
        section.left_margin = Cm(1.0)
        section.right_margin = Cm(1.0)
        section.top_margin = Cm(1.5)
        section.bottom_margin = Cm(1.5)

    def _add_header(self):
        """Add document header with metadata."""
        title = self.doc.add_paragraph()
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = title.add_run("Lampiran 7.C")
        run.bold = True
        run.font.size = Pt(14)

        subtitle = self.doc.add_paragraph()
        subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = subtitle.add_run("Skenario dan Hasil Uji Fungsionalitas")
        run.bold = True
        run.font.size = Pt(12)

        self.doc.add_paragraph()

        # Metadata
        meta_items = [
            ("Nama Penyedia Layanan", self.metadata.get("nama_penyedia", "Bank Sahabat Sampoerna")),
            ("Nama Pengguna Layanan", self.metadata.get("nama_pengguna", "")),
            ("Nama Layanan API", self.metadata.get("nama_layanan", "API QR MPM")),
            ("Tanggal Pengujian", self.metadata.get("tanggal_pengujian", "")),
        ]

        for label, value in meta_items:
            p = self.doc.add_paragraph()
            p.add_run(f"{label}: ").bold = True
            p.add_run(value or "")

        self.doc.add_paragraph()

    def _add_scenario_table(self):
        """Add the main scenario table following ASPI format."""
        # Filter only QR MPM scenarios (ASPI no starting with 18.)
        qr_scenarios = [s for s in self.scenarios if s.get("aspi_no", "").startswith("18.")]
        
        if not qr_scenarios:
            # If no ASPI numbers detected, use all scenarios from QR MPM section
            qr_scenarios = [s for s in self.scenarios 
                          if "qr" in s.get("section", "").lower() or 
                             "qr" in s.get("nama_modul", "").lower() or
                             s.get("aspi_no", "").startswith("18.")]

        num_rows = len(qr_scenarios) + 1  # +1 for header
        table = create_table_with_borders(self.doc, num_rows, 8)

        # Set column widths
        col_widths = [Cm(1.0), Cm(2.5), Cm(4.0), Cm(3.5), Cm(11.0), Cm(11.0), Cm(1.5), Cm(4.0)]
        for i, width in enumerate(col_widths):
            table.columns[i].width = width

        # Header row
        headers = ["No", "Service", "Scenario", "Expected Result",
                   "Request", "Response", "Result", "Notes"]
        for i, header in enumerate(headers):
            cell = table.rows[0].cells[i]
            add_cell_text(cell, header, font_size=Pt(9), bold=True,
                         color=RGBColor(255, 255, 255),
                         alignment=WD_ALIGN_PARAGRAPH.CENTER)
            set_cell_shading(cell, "4472C4")

        # Data rows
        for row_idx, scenario in enumerate(qr_scenarios, 1):
            row = table.rows[row_idx]
            
            result = self.validator.validate(scenario)
            
            # Determine notes
            if scenario["is_skipped"]:
                notes = scenario.get("remarks_raw", "") or SKIP_REASON
                if not notes or notes == "None":
                    notes = SKIP_REASON
            else:
                notes = ""

            # Format request column - extract from parsed remarks
            request_output = scenario.get("request", "")
            
            # Format response column
            response_output = scenario.get("response", "")

            # Extract scenario name from langkah_tes
            scenario_name = self._extract_scenario_name(scenario.get("langkah_tes", ""))
            
            # Determine service name
            service_name = scenario.get("nama_modul", "Any Service")

            # Row number (ASPI numbering)
            row_no = scenario.get("aspi_no", scenario.get("nomor_kasus_tes", ""))

            # Fill cells
            # Col 0: No
            add_cell_text(row.cells[0], row_no, font_size=Pt(8),
                         alignment=WD_ALIGN_PARAGRAPH.CENTER)
            # Col 1: Service
            add_cell_text(row.cells[1], service_name, font_size=Pt(8))
            # Col 2: Scenario
            add_cell_text(row.cells[2], scenario_name, font_size=Pt(8))
            # Col 3: Expected Result
            add_cell_text(row.cells[3], scenario.get("expected_result", ""), font_size=Pt(8))
            # Col 4: Request (URL + Headers + Body)
            add_cell_text(row.cells[4], request_output, 
                         font_name='Consolas', font_size=Pt(7))
            # Col 5: Response
            add_cell_text(row.cells[5], response_output, 
                         font_name='Consolas', font_size=Pt(7))
            # Col 6: Result
            result_color = None
            if result == "PASS":
                result_color = RGBColor(0, 128, 0)
            elif result == "NOT PASS":
                result_color = RGBColor(255, 0, 0)
            add_cell_text(row.cells[6], result, font_size=Pt(8),
                         bold=True, color=result_color,
                         alignment=WD_ALIGN_PARAGRAPH.CENTER)
            # Col 7: Notes
            add_cell_text(row.cells[7], notes, font_size=Pt(7))

    def _extract_scenario_name(self, langkah_tes):
        """Extract scenario name from Langkah Tes, removing ASPI number prefix."""
        if not langkah_tes:
            return ""
        # Remove leading number pattern like "18,1 " or "3,5 "
        cleaned = re.sub(r'^\d+[,.]\d+\s*', '', langkah_tes)
        # Remove comment artifacts
        cleaned = re.sub(r'\n.*?======.*', '', cleaned, flags=re.DOTALL)
        return cleaned.strip()

    def _add_footer_notes(self):
        """Add footer notes as per ASPI requirements."""
        self.doc.add_paragraph()
        notes = [
            "Lampiran Skenario hasil uji fungsional sekurangnya 1 Pengguna Layanan atas 1 sub API unverified, dengan ketentuan sebagai berikut:",
            "",
            "a. Pada kolom request diisi dengan request yang dilakukan Pengguna layanan, sedangkan pada kolom response diisi dengan respon yang diberikan Penyedia. Sementara pada kolom result diisi dengan hasil PASS atau NOT PASS yang harus sesuai dengan expected result.",
            "",
            "b. Pengisian pada dokumen skenario hasil uji fungsional tidak dilakukan dengan cara screen capture, melainkan dilakukan dengan cara copy paste payload request dan response dari log API server ke kolom tabel skenario hasil uji fungsional.",
            "",
            "c. Seluruh skenario diujikan dan tidak boleh dihapus atau diubah. Dalam hal terdapat skenario yang tidak diujikan dapat dikosongkan pengisiannya, namun diberikan catatan pada kolom Notes yang akan kami review lebih lanjut apakah skenario diperkenankan untuk tidak diujikan.",
            "",
            "d. Dalam hal terdapat penambahan skenario pengujian, maka penambahan tersebut dilakukan pada baris paling bawah, sehingga tidak mengubah susunan atau urutan template skenario.",
        ]
        for note in notes:
            p = self.doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(1)
            run = p.add_run(note)
            run.font.size = Pt(8)
            run.italic = True



# =============================================================================
# UAT RESULT DOCUMENT GENERATOR
# =============================================================================

class UATResultGenerator:
    """Generate UAT Result document (.docx)."""

    def __init__(self, scenarios, metadata):
        self.scenarios = scenarios
        self.metadata = metadata
        self.doc = Document()
        self.validator = ResultValidator()

    def generate(self, output_path):
        """Generate the UAT Result document."""
        self._set_styles()
        self._add_title_page()
        self._add_table_of_contents()
        self._add_skipped_sections()
        self._add_qr_mpm_section()
        self._add_additional_sections()
        self._add_summary()

        self.doc.save(output_path)
        print(f"[OK] UAT Result saved: {output_path}")

    def _set_styles(self):
        """Set default document styles."""
        style = self.doc.styles['Normal']
        font = style.font
        font.name = 'Calibri'
        font.size = Pt(10)

    def _add_title_page(self):
        """Add title page."""
        for _ in range(3):
            self.doc.add_paragraph()

        title = self.doc.add_paragraph()
        title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = title.add_run("UAT Result")
        run.bold = True
        run.font.size = Pt(24)

        self.doc.add_paragraph()

        subtitle = self.doc.add_paragraph()
        subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
        run = subtitle.add_run("Penambahan Layanan QRIS Merchant Aggregator")
        run.bold = True
        run.font.size = Pt(16)

        self.doc.add_paragraph()

        info = self.doc.add_paragraph()
        info.alignment = WD_ALIGN_PARAGRAPH.CENTER
        info.add_run(f"Nama Penyedia Layanan: {self.metadata.get('nama_penyedia', '')}").font.size = Pt(11)
        info.add_run("\n")
        info.add_run(f"Nama Pengguna Layanan: {self.metadata.get('nama_pengguna', '')}").font.size = Pt(11)
        info.add_run("\n")
        info.add_run(f"Tanggal Pengujian: {self.metadata.get('tanggal_pengujian', '')}").font.size = Pt(11)

        self.doc.add_page_break()

    def _add_table_of_contents(self):
        """Add table of contents."""
        self.doc.add_heading("Daftar Isi", level=1)

        toc_items = [
            "1  Balance Services",
            "2  API Transaction History List",
            "3  QR MPM",
            "4  Pengecekan Mutasi Dan Jurnal",
            "5  Generate QR SNAP",
            "6  Refund Payment",
            "7  Query Payment",
            "8  Inquiry & Report",
        ]
        for item in toc_items:
            p = self.doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(1)
            p.add_run(item).font.size = Pt(10)

        self.doc.add_page_break()

    def _add_skipped_sections(self):
        """Add Balance Services and Transaction History (all skipped)."""
        self.doc.add_heading("1 Balance Services", level=1)
        p = self.doc.add_paragraph()
        p.add_run(SKIP_REASON).italic = True
        self.doc.add_paragraph()

        self.doc.add_heading("2 API Transaction History List", level=1)
        p = self.doc.add_paragraph()
        p.add_run(SKIP_REASON).italic = True
        self.doc.add_page_break()

    def _add_qr_mpm_section(self):
        """Add QR MPM section with test results."""
        self.doc.add_heading("3 QR MPM", level=1)

        # Get QR MPM scenarios
        qr_scenarios = [s for s in self.scenarios if s.get("aspi_no", "").startswith("18.")]
        
        if not qr_scenarios:
            qr_scenarios = [s for s in self.scenarios 
                          if "qr" in s.get("section", "").lower() or
                             "qr" in s.get("nama_modul", "").lower()]

        for idx, scenario in enumerate(qr_scenarios, 1):
            scenario_name = re.sub(r'^\d+[,.]\d+\s*', '', scenario.get("langkah_tes", ""))
            scenario_name = re.sub(r'\n.*?======.*', '', scenario_name, flags=re.DOTALL).strip()
            
            aspi_no = scenario.get("aspi_no", "")
            self.doc.add_heading(f"3.{idx} {aspi_no} {scenario_name}", level=2)

            if scenario["is_skipped"]:
                p = self.doc.add_paragraph()
                p.add_run(SKIP_REASON).italic = True
            else:
                self._add_scenario_detail(scenario)

            self.doc.add_paragraph()

    def _add_scenario_detail(self, scenario):
        """Add detailed scenario with request/response."""
        # Expected Result
        p = self.doc.add_paragraph()
        p.add_run("Expected Result: ").bold = True
        p.add_run(scenario.get("expected_result", ""))

        # Request (URL + Headers + Body)
        p = self.doc.add_paragraph()
        p.add_run("Request:").bold = True
        
        request_text = scenario.get("request", "")
        if request_text:
            p = self.doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.5)
            run = p.add_run(request_text)
            run.font.name = 'Consolas'
            run.font.size = Pt(8)

        # Response
        p = self.doc.add_paragraph()
        p.add_run("Response:").bold = True
        
        response_text = scenario.get("response", "")
        if response_text:
            p = self.doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(0.5)
            run = p.add_run(response_text)
            run.font.name = 'Consolas'
            run.font.size = Pt(8)

        # Result
        result = self.validator.validate(scenario)
        p = self.doc.add_paragraph()
        p.add_run("Result: ").bold = True
        result_run = p.add_run(result)
        result_run.bold = True
        if result == "PASS":
            result_run.font.color.rgb = RGBColor(0, 128, 0)
        elif result == "NOT PASS":
            result_run.font.color.rgb = RGBColor(255, 0, 0)

    def _add_additional_sections(self):
        """Add remaining sections (Pengecekan Mutasi, etc.)."""
        self.doc.add_page_break()
        
        sections = [
            ("4", "Pengecekan Mutasi Dan Jurnal"),
            ("5", "Generate QR SNAP"),
            ("6", "Refund Payment"),
            ("7", "Query Payment"),
            ("8", "Inquiry & Report"),
        ]

        for num, title in sections:
            self.doc.add_heading(f"{num} {title}", level=1)
            
            # Find scenarios for this section
            section_scenarios = [s for s in self.scenarios 
                               if title.lower() in s.get("section", "").lower() or
                                  title.lower() in s.get("nama_modul", "").lower()]
            
            if section_scenarios:
                for s in section_scenarios:
                    p = self.doc.add_paragraph()
                    langkah = s.get("langkah_tes", "")
                    hasil = s.get("hasil_aktual", "")
                    p.add_run(f"- {langkah}: ").bold = True
                    p.add_run(hasil)
            else:
                p = self.doc.add_paragraph()
                p.add_run(SKIP_REASON).italic = True
            
            self.doc.add_paragraph()

    def _add_summary(self):
        """Add summary table."""
        self.doc.add_page_break()
        self.doc.add_heading("Ringkasan Hasil Pengujian", level=1)

        total = 0
        passed = 0
        failed = 0
        skipped = 0
        not_tested = 0

        for s in self.scenarios:
            total += 1
            result = self.validator.validate(s)
            if result == "N/A":
                skipped += 1
            elif result == "PASS":
                passed += 1
            elif result == "NOT PASS":
                failed += 1
            else:
                not_tested += 1

        # Summary table
        table = create_table_with_borders(self.doc, 6, 2)
        table.columns[0].width = Cm(8)
        table.columns[1].width = Cm(4)

        data = [
            ("Kategori", "Jumlah"),
            ("Total Skenario", str(total)),
            ("PASS", str(passed)),
            ("NOT PASS", str(failed)),
            ("Tidak Diuji (N/A)", str(skipped)),
            ("Belum Diisi", str(not_tested)),
        ]

        for i, (label, value) in enumerate(data):
            add_cell_text(table.rows[i].cells[0], label, font_size=Pt(10),
                         bold=(i == 0),
                         color=RGBColor(255, 255, 255) if i == 0 else None)
            add_cell_text(table.rows[i].cells[1], value, font_size=Pt(10),
                         bold=(i == 0),
                         color=RGBColor(255, 255, 255) if i == 0 else None,
                         alignment=WD_ALIGN_PARAGRAPH.CENTER)
            if i == 0:
                set_cell_shading(table.rows[i].cells[0], "4472C4")
                set_cell_shading(table.rows[i].cells[1], "4472C4")



# =============================================================================
# MAIN EXECUTION
# =============================================================================

def main():
    """Main function to generate UAT documents."""
    print("=" * 60)
    print("  QRIS Merchant Aggregator - UAT Document Generator")
    print("  Bank Sahabat Sampoerna (BSS)")
    print("=" * 60)
    print()

    # Input file
    if len(sys.argv) > 1:
        input_excel = sys.argv[1]
    else:
        # Auto-find .xlsx file in current directory
        xlsx_files = [f for f in os.listdir('.') if f.endswith('.xlsx') and not f.startswith('~')]
        if xlsx_files:
            input_excel = xlsx_files[0]
            print(f"[INFO] Auto-detected Excel file: {input_excel}")
        else:
            print("ERROR: File Excel (.xlsx) tidak ditemukan.")
            print(f"Usage: python {sys.argv[0]} <path_to_uat_script.xlsx>")
            sys.exit(1)

    if not os.path.exists(input_excel):
        print(f"ERROR: File tidak ditemukan: {input_excel}")
        sys.exit(1)

    # Output paths
    timestamp = datetime.now().strftime("%Y%m%d")
    output_dir = "output"
    os.makedirs(output_dir, exist_ok=True)

    uat_result_path = os.path.join(output_dir, f"UAT_Result_QRIS_Merchant_Aggregator_{timestamp}.docx")
    lampiran_7c_path = os.path.join(output_dir, f"Lampiran_7C_QRIS_{timestamp}.docx")

    # Parse Excel
    print(f"[1/3] Membaca UAT Script: {input_excel}")
    parser = UATScriptParser(input_excel)
    scenarios, metadata = parser.parse()
    
    if not scenarios:
        print("ERROR: Tidak ada skenario yang ditemukan. Periksa format Excel.")
        sys.exit(1)

    print(f"      -> Nama Pengguna: {metadata.get('nama_pengguna', '(belum diisi)')}")
    print(f"      -> Tanggal: {metadata.get('tanggal_pengujian', '(belum diisi)')}")
    
    # Show parsing summary
    tested = sum(1 for s in scenarios if s["is_tested"])
    skipped = sum(1 for s in scenarios if s["is_skipped"])
    print(f"      -> Tested: {tested}, Skipped: {skipped}, Total: {len(scenarios)}")
    print()

    # Generate UAT Result
    print(f"[2/3] Generating UAT Result...")
    uat_gen = UATResultGenerator(scenarios, metadata)
    uat_gen.generate(uat_result_path)
    print()

    # Generate Lampiran 7C
    print(f"[3/3] Generating Lampiran 7C (Berita Acara)...")
    lamp_gen = Lampiran7CGenerator(scenarios, metadata)
    lamp_gen.generate(lampiran_7c_path)
    print()

    # Summary
    print("=" * 60)
    print("  SELESAI!")
    print("=" * 60)
    print(f"  Output files:")
    print(f"    1. {uat_result_path}")
    print(f"    2. {lampiran_7c_path}")
    print()

    # Validation summary
    validator = ResultValidator()
    passed = sum(1 for s in scenarios if validator.validate(s) == "PASS")
    failed = sum(1 for s in scenarios if validator.validate(s) == "NOT PASS")
    na = sum(1 for s in scenarios if validator.validate(s) == "N/A")
    not_tested = sum(1 for s in scenarios if validator.validate(s) == "NOT TESTED")

    print(f"  Hasil Validasi:")
    print(f"    PASS       : {passed}")
    print(f"    NOT PASS   : {failed}")
    print(f"    N/A        : {na}")
    print(f"    Belum Diisi: {not_tested}")
    print()

    if failed > 0:
        print("  [WARNING] Ada skenario NOT PASS! Review kembali sebelum submit ke ASPI.")
    elif not_tested > 0:
        print("  [INFO] Masih ada skenario yang belum diisi data request/response.")
    else:
        print("  [OK] Semua skenario yang diuji PASS. Siap submit ke ASPI Portal.")


if __name__ == "__main__":
    main()
