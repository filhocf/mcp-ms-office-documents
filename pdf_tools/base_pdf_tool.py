"""PDF generation from Markdown and DOCX conversion using fpdf2."""

import io
import re
from pathlib import Path
from fpdf import FPDF
from upload_tools.main import upload_file


class MarkdownPDF(FPDF):
    """Custom PDF class with Markdown rendering support."""

    def __init__(self):
        super().__init__()
        self.add_page()
        self.set_auto_page_break(auto=True, margin=15)
        self.set_font("Helvetica", size=11)

    def render_markdown(self, markdown_text: str):
        """Parse and render Markdown content to PDF."""
        lines = markdown_text.split("\n")
        i = 0
        while i < len(lines):
            line = lines[i]

            # Headings
            if line.startswith("# "):
                self.set_font("Helvetica", "B", 20)
                self.cell(0, 12, line[2:].strip(), new_x="LMARGIN", new_y="NEXT")
                self.ln(4)
            elif line.startswith("## "):
                self.set_font("Helvetica", "B", 16)
                self.cell(0, 10, line[3:].strip(), new_x="LMARGIN", new_y="NEXT")
                self.ln(3)
            elif line.startswith("### "):
                self.set_font("Helvetica", "B", 13)
                self.cell(0, 9, line[4:].strip(), new_x="LMARGIN", new_y="NEXT")
                self.ln(2)
            elif line.startswith("#### "):
                self.set_font("Helvetica", "B", 11)
                self.cell(0, 8, line[5:].strip(), new_x="LMARGIN", new_y="NEXT")
                self.ln(2)
            # Bullet lists
            elif line.strip().startswith("- ") or line.strip().startswith("* "):
                self.set_font("Helvetica", size=11)
                text = line.strip()[2:]
                self.cell(10)
                self.cell(0, 7, f"- {text}", new_x="LMARGIN", new_y="NEXT")
            # Numbered lists
            elif re.match(r"^\d+\.\s", line.strip()):
                self.set_font("Helvetica", size=11)
                self.cell(10)
                self.cell(0, 7, line.strip(), new_x="LMARGIN", new_y="NEXT")
            # Horizontal rule
            elif line.strip() in ("---", "***", "___"):
                self.ln(3)
                self.line(self.l_margin, self.get_y(), self.w - self.r_margin, self.get_y())
                self.ln(3)
            # Empty line
            elif line.strip() == "":
                self.ln(4)
            # Regular paragraph
            else:
                self.set_font("Helvetica", size=11)
                # Handle bold/italic inline
                clean_text = re.sub(r"\*\*(.+?)\*\*", r"\1", line)
                clean_text = re.sub(r"\*(.+?)\*", r"\1", clean_text)
                clean_text = re.sub(r"`(.+?)`", r"\1", clean_text)
                self.multi_cell(0, 7, clean_text)

            i += 1


def markdown_to_pdf(markdown_content: str, file_name: str | None = None) -> str:
    """Convert Markdown text to PDF and upload."""
    pdf = MarkdownPDF()
    pdf.render_markdown(markdown_content)

    buffer = io.BytesIO()
    pdf.output(buffer)
    buffer.seek(0)

    result = upload_file(buffer, "pdf", file_name=file_name)
    return result


def docx_to_pdf(file_path: str, file_name: str | None = None) -> str:
    """Convert a DOCX file to PDF by extracting text and rendering."""
    from read_tools.docx_reader import get_docx_paragraphs

    paragraphs = get_docx_paragraphs(file_path)

    pdf = MarkdownPDF()
    for para in paragraphs:
        style = para["style"]
        text = para["text"]
        if not text.strip():
            pdf.ln(4)
            continue

        if "Heading 1" in style:
            pdf.set_font("Helvetica", "B", 20)
            pdf.cell(0, 12, text, new_x="LMARGIN", new_y="NEXT")
            pdf.ln(4)
        elif "Heading 2" in style:
            pdf.set_font("Helvetica", "B", 16)
            pdf.cell(0, 10, text, new_x="LMARGIN", new_y="NEXT")
            pdf.ln(3)
        elif "Heading 3" in style:
            pdf.set_font("Helvetica", "B", 13)
            pdf.cell(0, 9, text, new_x="LMARGIN", new_y="NEXT")
            pdf.ln(2)
        elif "List" in style:
            pdf.set_font("Helvetica", size=11)
            pdf.cell(10)
            pdf.cell(0, 7, f"- {text}", new_x="LMARGIN", new_y="NEXT")
        else:
            pdf.set_font("Helvetica", size=11)
            pdf.multi_cell(0, 7, text)

    buffer = io.BytesIO()
    pdf.output(buffer)
    buffer.seek(0)

    if not file_name:
        file_name = Path(file_path).stem

    result = upload_file(buffer, "pdf", file_name=file_name)
    return result
