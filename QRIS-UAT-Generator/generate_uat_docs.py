#!/usr/bin/env python3
"""
QRIS Merchant Aggregator - UAT Document Generator
Bank Sahabat Sampoerna (BSS)

Tool ini membaca UAT Script Excel yang sudah diisi oleh mitra,
lalu menghasilkan:
1. UAT Result (.docx) - Dokumen hasil pengujian detail
2. Lampiran 7C (.docx) - Berita Acara untuk ASPI Portal

Author: IT QA BSS
"""

import re
import sys
import os
from datetime import datetime
from copy import deepcopy

from openpyxl import load_workbook
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_ORIENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml


# =============================================================================
# CONFIGURATION
# =============================================================================

# Skenario yang TIDAK digunakan (highlight kuning) - bisa dikustomisasi
# Berdasarkan produk QRIS Merchant Aggregator BSS
SKIPPED_SCENARIOS_DEFAULT = {
    # Balance Services (semua di-skip untuk QRIS)
    "3.1", "3.2", "3.3", "3.4", "3.5", "3.6", "3.7", "3.8", "3.9", "3.10", "3.11",
    # Transaction History (semua di-skip untuk QRIS)
    "4.1", "4.2", "4.3", "4.4", "4.5", "4.6", "4.7", "4.8",
}

SKIP_REASON = "Tidak dites karena tidak sesuai dengan kondisi produk."



# =============================================================================
# EXCEL PARSER
# =============================================================================

class UATScriptParser:
    """Parser untuk membaca UAT Script Excel yang diisi mitra."""

    def __init__(self, excel_path):
        self.excel_path = excel_path
        self.wb = load_workbook(excel_path, data_only=True)
        self.scenarios = []
        self.metadata = {
            "nama_penyedia": "Bank Sahabat Sampoerna",
            "nama_layanan": "API QR MPM",
            "nama_pengguna": "",
            "tanggal_pengujian": "",
        }

    def parse(self):
        """Parse the UAT Script Excel file."""
        ws = self.wb.active

        # Try to extract metadata from header rows
        for row in ws.iter_rows(min_row=1, max_row=6, values_only=False):
            for cell in row:
                if cell.value and isinstance(cell.value, str):
                    if "Nama Penyedia" in cell.value:
                        # Get value from next cell or same cell after colon
                        val = self._extract_value(cell, row, ws)
                        if val:
                            self.metadata["nama_penyedia"] = val
                    elif "Nama Layanan" in cell.value:
                        val = self._extract_value(cell, row, ws)
                        if val:
                            self.metadata["nama_layanan"] = val
                    elif "Nama Pengguna" in cell.value:
                        val = self._extract_value(cell, row, ws)
                        if val:
                            self.metadata["nama_pengguna"] = val
                    elif "Tanggal" in cell.value:
                        val = self._extract_value(cell, row, ws)
                        if val:
                            self.metadata["tanggal_pengujian"] = str(val)

        # Find header row (contains "No", "Service", "Scenario", etc.)
        header_row = None
        for row_idx, row in enumerate(ws.iter_rows(min_row=1, values_only=True), 1):
            row_values = [str(c).strip().lower() if c else "" for c in row]
            if "no" in row_values and "service" in row_values:
                header_row = row_idx
                break

        if not header_row:
            print("ERROR: Tidak dapat menemukan header row di Excel.")
            return

        # Parse scenario rows
        for row in ws.iter_rows(min_row=header_row + 1, values_only=True):
            no = row[0] if len(row) > 0 else None
            if no is None or str(no).strip() == "":
                continue

            scenario = {
                "no": str(row[0]).strip() if row[0] else "",
                "service": str(row[1]).strip() if len(row) > 1 and row[1] else "",
                "scenario": str(row[2]).strip() if len(row) > 2 and row[2] else "",
                "expected_result": str(row[3]).strip() if len(row) > 3 and row[3] else "",
                "request": str(row[4]).strip() if len(row) > 4 and row[4] else "",
                "response": str(row[5]).strip() if len(row) > 5 and row[5] else "",
                "result": str(row[6]).strip() if len(row) > 6 and row[6] else "",
                "notes": str(row[7]).strip() if len(row) > 7 and row[7] else "",
            }
            self.scenarios.append(scenario)

        return self.scenarios, self.metadata


    def _extract_value(self, cell, row, ws):
        """Extract value from cell - handles 'Label: Value' or adjacent cell."""
        val = cell.value
        if ":" in str(val):
            parts = str(val).split(":", 1)
            if len(parts) > 1 and parts[1].strip():
                return parts[1].strip()
        # Try next column
        col_idx = cell.column
        next_cell = ws.cell(row=cell.row, column=col_idx + 1)
        if next_cell.value:
            return str(next_cell.value).strip()
        return None


# =============================================================================
# PASS/FAIL VALIDATOR
# =============================================================================

class ResultValidator:
    """Validasi hasil berdasarkan response code vs expected result."""

    @staticmethod
    def extract_response_code(response_text):
        """Extract response/error code from response text."""
        if not response_text:
            return None
        # Match patterns like "responseCode": "2004700" or responseCode: 2004700
        patterns = [
            r'"responseCode"\s*:\s*"(\d+)"',
            r'"httpCode"\s*:\s*(\d+)',
            r'"responseCode"\s*:\s*(\d+)',
            r'responseCode.*?(\d{7})',
            r'[Cc]ode.*?(\d{7})',
        ]
        for pattern in patterns:
            match = re.search(pattern, response_text)
            if match:
                return match.group(1)
        return None

    @staticmethod
    def extract_expected_code(expected_text):
        """Extract expected response/error code from expected result text."""
        if not expected_text:
            return None
        # Match patterns like "Response Code: 2004700" or "Error Code: 4044708"
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
        Returns: 'PASS', 'NOT PASS', 'N/A', or 'NOT TESTED'
        """
        # If already filled by mitra
        if scenario.get("result") and scenario["result"].upper() in ["PASS", "NOT PASS", "N/A"]:
            return scenario["result"].upper()

        # If no response provided, it's not tested
        if not scenario.get("response") or scenario["response"] in ["None", "", "Response Body:"]:
            return "NOT TESTED"

        expected_code = ResultValidator.extract_expected_code(scenario.get("expected_result", ""))
        actual_code = ResultValidator.extract_response_code(scenario.get("response", ""))

        if not expected_code or not actual_code:
            return "NOT TESTED"

        # Handle xx pattern (e.g., 401xx01)
        if "xx" in expected_code:
            # Convert pattern like 401xx01 to regex
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


def set_cell_border(cell, **kwargs):
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


def add_formatted_text(paragraph, text, bold=False, size=Pt(10)):
    """Add formatted run to paragraph."""
    run = paragraph.add_run(text)
    run.bold = bold
    run.font.size = size
    return run


def create_table_with_borders(doc, rows, cols):
    """Create a table with all borders."""
    table = doc.add_table(rows=rows, cols=cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER

    # Set borders for all cells
    for row in table.rows:
        for cell in row.cells:
            set_cell_border(cell)
    return table



# =============================================================================
# UAT RESULT DOCUMENT GENERATOR
# =============================================================================

class UATResultGenerator:
    """Generate UAT Result document (.docx)."""

    def __init__(self, scenarios, metadata, skipped_scenarios=None):
        self.scenarios = scenarios
        self.metadata = metadata
        self.skipped = skipped_scenarios or set()
        self.doc = Document()
        self.validator = ResultValidator()

    def generate(self, output_path):
        """Generate the UAT Result document."""
        self._set_styles()
        self._add_title_page()
        self._add_table_of_contents()
        self._add_qr_mpm_section()
        self._add_pengecekan_mutasi_section()
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
        # Add some spacing
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

        # Metadata
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
        heading = self.doc.add_heading("Daftar Isi", level=1)

        # Group scenarios by service
        toc_items = []
        current_section = ""
        section_num = 0

        # Section 1 & 2: Balance & Transaction History (all skipped)
        toc_items.append(("1", "Balance Services", []))
        toc_items.append(("2", "API Transaction History List", []))

        # Section 3: QR MPM
        qr_scenarios = []
        for s in self.scenarios:
            if s["no"].startswith("18."):
                qr_scenarios.append(s)
        toc_items.append(("3", "QR MPM", qr_scenarios))

        # Section 4+: Additional sections
        toc_items.append(("4", "Pengecekan Mutasi Dan Jurnal", []))
        toc_items.append(("5", "Generate QR SNAP", []))
        toc_items.append(("6", "Refund Payment", []))
        toc_items.append(("7", "Query Payment", []))
        toc_items.append(("8", "Inquiry & Report", []))

        for section_id, section_name, _ in toc_items:
            p = self.doc.add_paragraph()
            p.paragraph_format.left_indent = Cm(1)
            p.add_run(f"{section_id}  {section_name}").font.size = Pt(10)

        self.doc.add_page_break()

    def _add_section_header(self, level, number, title):
        """Add a section header."""
        heading = self.doc.add_heading(f"{number} {title}", level=level)
        return heading


    def _add_skipped_section(self, section_num, section_title, scenarios_info):
        """Add a section that is entirely skipped."""
        self._add_section_header(1, str(section_num), section_title)
        for idx, (sub_no, scenario_name) in enumerate(scenarios_info, 1):
            self._add_section_header(2, f"{section_num}.{idx}", f"{sub_no} {scenario_name}")
            p = self.doc.add_paragraph()
            p.add_run(SKIP_REASON).italic = True

    def _add_qr_mpm_section(self):
        """Add QR MPM section with actual test results."""
        self._add_section_header(1, "3", "QR MPM")

        # Define which scenarios are used (not highlighted yellow)
        # Based on the UAT Result template, these are tested:
        # 18.1-18.7: General + Generate QR (tested)
        # 18.8-18.14: Decode QR, Payment Redirect, Apply OTT, Payment H2H (skipped based on mode)
        # 18.15-18.17: Query Payment (tested)
        # 18.18-18.19: Payment Notification (tested)

        sub_idx = 1
        for scenario in self.scenarios:
            no = scenario["no"]
            if not no.startswith("18."):
                continue

            self._add_section_header(2, f"3.{sub_idx}", f"{no} {scenario['scenario']}")

            # Check if this scenario is in the skip list
            is_skipped = self._is_scenario_skipped(scenario)

            if is_skipped:
                p = self.doc.add_paragraph()
                p.add_run(SKIP_REASON).italic = True
            else:
                self._add_scenario_detail(scenario)

            sub_idx += 1

    def _is_scenario_skipped(self, scenario):
        """Determine if a scenario should be marked as skipped."""
        no = scenario["no"]
        # If result is empty and no request/response filled
        if (not scenario.get("request") or scenario["request"] in ["None", "", "URL Endpoint:\nHeader Request:\nRequest Body:"]) \
           and (not scenario.get("response") or scenario["response"] in ["None", "", "Response Body:"]):
            # Check notes
            if scenario.get("notes") and "tidak" in scenario["notes"].lower():
                return True
            # Check if in predefined skip list
            if no in self.skipped:
                return True
        return False


    def _add_scenario_detail(self, scenario):
        """Add detailed scenario result with request/response."""
        # Expected Result
        p = self.doc.add_paragraph()
        p.add_run("Expected Result: ").bold = True
        p.add_run(scenario.get("expected_result", ""))

        # Request
        p = self.doc.add_paragraph()
        p.add_run("Request:").bold = True
        self.doc.add_paragraph()
        req_text = scenario.get("request", "")
        if req_text and req_text not in ["None", ""]:
            # Parse and display request
            self._add_code_block(req_text)

        # Response
        p = self.doc.add_paragraph()
        p.add_run("Response:").bold = True
        self.doc.add_paragraph()
        resp_text = scenario.get("response", "")
        if resp_text and resp_text not in ["None", ""]:
            self._add_code_block(resp_text)

        # Result (PASS/NOT PASS)
        result = self.validator.validate(scenario)
        p = self.doc.add_paragraph()
        p.add_run("Result: ").bold = True
        result_run = p.add_run(result)
        if result == "PASS":
            result_run.font.color.rgb = RGBColor(0, 128, 0)  # Green
        elif result == "NOT PASS":
            result_run.font.color.rgb = RGBColor(255, 0, 0)  # Red
        result_run.bold = True

        # Notes
        if scenario.get("notes") and scenario["notes"] not in ["None", ""]:
            p = self.doc.add_paragraph()
            p.add_run("Notes: ").bold = True
            p.add_run(scenario["notes"])

        self.doc.add_paragraph()  # spacing

    def _add_code_block(self, text):
        """Add a code block style text."""
        p = self.doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(1)
        run = p.add_run(text)
        run.font.name = 'Consolas'
        run.font.size = Pt(8)


    def _add_pengecekan_mutasi_section(self):
        """Add Pengecekan Mutasi Dan Jurnal section."""
        self._add_section_header(1, "4", "Pengecekan Mutasi Dan Jurnal")
        p = self.doc.add_paragraph()
        p.add_run("Hasil pengecekan mutasi dan jurnal akan dilampirkan terpisah.")

    def _add_summary(self):
        """Add summary of test results."""
        self.doc.add_page_break()
        self._add_section_header(1, "", "Ringkasan Hasil Pengujian")

        total = 0
        passed = 0
        failed = 0
        skipped = 0
        not_tested = 0

        for s in self.scenarios:
            total += 1
            result = self.validator.validate(s)
            if self._is_scenario_skipped(s):
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
            table.rows[i].cells[0].text = label
            table.rows[i].cells[1].text = value
            if i == 0:
                set_cell_shading(table.rows[i].cells[0], "4472C4")
                set_cell_shading(table.rows[i].cells[1], "4472C4")
                for cell in table.rows[i].cells:
                    for paragraph in cell.paragraphs:
                        for run in paragraph.runs:
                            run.font.color.rgb = RGBColor(255, 255, 255)
                            run.bold = True



# =============================================================================
# LAMPIRAN 7C (BERITA ACARA) DOCUMENT GENERATOR
# =============================================================================

class Lampiran7CGenerator:
    """Generate Lampiran 7C - Berita Acara for ASPI Portal (.docx)."""

    def __init__(self, scenarios, metadata, skipped_scenarios=None):
        self.scenarios = scenarios
        self.metadata = metadata
        self.skipped = skipped_scenarios or set()
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
        section.page_width = Cm(29.7)
        section.page_height = Cm(21.0)
        section.left_margin = Cm(1.5)
        section.right_margin = Cm(1.5)
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

        # Metadata info
        meta_items = [
            ("Nama Penyedia Layanan", self.metadata.get("nama_penyedia", "Bank Sahabat Sampoerna")),
            ("Nama Pengguna Layanan", self.metadata.get("nama_pengguna", "")),
            ("Nama Layanan API", self.metadata.get("nama_layanan", "API QR MPM")),
            ("Tanggal Pengujian", self.metadata.get("tanggal_pengujian", "")),
        ]

        for label, value in meta_items:
            p = self.doc.add_paragraph()
            p.add_run(f"{label}: ").bold = True
            p.add_run(value)

        self.doc.add_paragraph()


    def _add_scenario_table(self):
        """Add the main scenario table following ASPI format."""
        # Header row + data rows
        num_rows = len(self.scenarios) + 1
        table = create_table_with_borders(self.doc, num_rows, 8)

        # Set column widths
        col_widths = [Cm(1.2), Cm(2.5), Cm(3.5), Cm(3.5), Cm(5.5), Cm(5.5), Cm(1.5), Cm(3.5)]
        for i, width in enumerate(col_widths):
            table.columns[i].width = width

        # Header row
        headers = ["No", "Service", "Scenario", "Expected Result", "Request", "Response", "Result", "Notes"]
        for i, header in enumerate(headers):
            cell = table.rows[0].cells[i]
            cell.text = ""
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = p.add_run(header)
            run.bold = True
            run.font.size = Pt(9)
            set_cell_shading(cell, "4472C4")
            run.font.color.rgb = RGBColor(255, 255, 255)

        # Data rows
        for row_idx, scenario in enumerate(self.scenarios, 1):
            row = table.rows[row_idx]

            # Determine result
            is_skipped = self._is_scenario_skipped(scenario)
            if is_skipped:
                result_text = "N/A"
                notes_text = scenario.get("notes", "") or SKIP_REASON
            else:
                result_text = self.validator.validate(scenario)
                notes_text = scenario.get("notes", "")

            # Fill cells
            cell_data = [
                scenario.get("no", ""),
                scenario.get("service", ""),
                scenario.get("scenario", ""),
                scenario.get("expected_result", ""),
                self._format_request(scenario.get("request", "")),
                self._format_response(scenario.get("response", "")),
                result_text,
                notes_text,
            ]

            for col_idx, data in enumerate(cell_data):
                cell = row.cells[col_idx]
                cell.text = ""
                p = cell.paragraphs[0]
                run = p.add_run(str(data) if data and data != "None" else "")
                run.font.size = Pt(8)

                # Color coding for result
                if col_idx == 6:  # Result column
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    run.bold = True
                    if result_text == "PASS":
                        run.font.color.rgb = RGBColor(0, 128, 0)
                    elif result_text == "NOT PASS":
                        run.font.color.rgb = RGBColor(255, 0, 0)

    def _is_scenario_skipped(self, scenario):
        """Determine if a scenario should be marked as skipped."""
        no = scenario["no"]
        if (not scenario.get("request") or scenario["request"] in ["None", "", "URL Endpoint:\nHeader Request:\nRequest Body:"]) \
           and (not scenario.get("response") or scenario["response"] in ["None", "", "Response Body:"]):
            if scenario.get("notes") and "tidak" in scenario["notes"].lower():
                return True
            if no in self.skipped:
                return True
        return False

    def _format_request(self, request_text):
        """Format request text for table cell."""
        if not request_text or request_text in ["None", ""]:
            return ""
        return request_text

    def _format_response(self, response_text):
        """Format response text for table cell."""
        if not response_text or response_text in ["None", ""]:
            return ""
        return response_text


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
# MAIN EXECUTION
# =============================================================================

def main():
    """Main function to generate UAT documents."""
    print("=" * 60)
    print("  QRIS Merchant Aggregator - UAT Document Generator")
    print("  Bank Sahabat Sampoerna (BSS)")
    print("=" * 60)
    print()

    # Default input/output paths
    if len(sys.argv) > 1:
        input_excel = sys.argv[1]
    else:
        input_excel = "UAT_Script_QRIS.xlsx"

    if not os.path.exists(input_excel):
        print(f"ERROR: File tidak ditemukan: {input_excel}")
        print(f"Usage: python {sys.argv[0]} <path_to_uat_script.xlsx>")
        print()
        print("Pastikan file UAT Script Excel yang sudah diisi mitra tersedia.")
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
    print(f"      -> {len(scenarios)} skenario ditemukan")
    print(f"      -> Nama Pengguna: {metadata.get('nama_pengguna', '(belum diisi)')}")
    print(f"      -> Tanggal: {metadata.get('tanggal_pengujian', '(belum diisi)')}")
    print()

    # Define skipped scenarios based on product
    # These are scenarios highlighted yellow (not applicable for this product)
    skipped = set()
    for s in scenarios:
        no = s["no"]
        # Auto-detect skipped: if notes contain skip reason or no data filled
        if s.get("notes") and ("tidak dites" in s["notes"].lower() or "tidak dilakukan" in s["notes"].lower()):
            skipped.add(no)

    # Generate UAT Result
    print(f"[2/3] Generating UAT Result...")
    uat_gen = UATResultGenerator(scenarios, metadata, skipped)
    uat_gen.generate(uat_result_path)
    print()

    # Generate Lampiran 7C
    print(f"[3/3] Generating Lampiran 7C (Berita Acara)...")
    lamp_gen = Lampiran7CGenerator(scenarios, metadata, skipped)
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
    passed = sum(1 for s in scenarios if validator.validate(s) == "PASS" and s["no"] not in skipped)
    failed = sum(1 for s in scenarios if validator.validate(s) == "NOT PASS" and s["no"] not in skipped)
    na = len(skipped)
    not_tested = len(scenarios) - passed - failed - na

    print(f"  Hasil Validasi:")
    print(f"    PASS      : {passed}")
    print(f"    NOT PASS  : {failed}")
    print(f"    N/A       : {na}")
    print(f"    Belum Diisi: {not_tested}")
    print()

    if failed > 0:
        print("  [WARNING] Ada skenario NOT PASS! Review kembali sebelum submit ke ASPI.")
    else:
        print("  [OK] Semua skenario yang diuji PASS. Siap submit ke ASPI Portal.")


if __name__ == "__main__":
    main()
