from fastmcp import FastMCP
from fastmcp.exceptions import ToolError
from pydantic import BaseModel, Field
from typing import Annotated, List, Dict, Optional, Literal
from xlsx_tools import markdown_to_excel
from docx_tools import markdown_to_word
from docx_tools.dynamic_docx_tools import register_docx_template_tools_from_yaml
from pptx_tools import create_presentation
from email_tools import create_eml
from email_tools.dynamic_email_tools import register_email_template_tools_from_yaml
from pathlib import Path
from config import get_config
from xml_tools import create_xml_file
from middleware import ApiKeyAuthMiddleware

import logging
mcp = FastMCP("MCP Office Documents")

# Initialize config and logging
config = get_config()
logger = logging.getLogger(__name__)

# Register API key auth middleware (only when API_KEY is set)
if config.api_key:
    mcp.add_middleware(ApiKeyAuthMiddleware(expected_key=config.api_key))
    logger.info("[auth] API key authentication enabled")
else:
    logger.info("[auth] No API_KEY set – authentication disabled")

# Look for dynamic email templates in production and local locations.
# Production (container): /app/config/email_templates.yaml
# Local development: <project_root>/config/email_templates.yaml
APP_CONFIG_PATH = Path("/app/config") / "email_templates.yaml"
LOCAL_CONFIG_PATH = Path(__file__).resolve().parent / "config" / "email_templates.yaml"

# Prefer the production path when present, otherwise fall back to local config.
_primary_yaml = None
for candidate in (APP_CONFIG_PATH, LOCAL_CONFIG_PATH):
    if candidate.exists():
        _primary_yaml = candidate
        logger.info("[dynamic-email] Found email templates file: %s", candidate)
        break

if _primary_yaml:
    try:
        register_email_template_tools_from_yaml(mcp, _primary_yaml)
    except Exception as e:
        logger.exception("[dynamic-email] Failed to register email templates from %s: %s", _primary_yaml, e)
else:
    logger.info(
        "[dynamic-email] No dynamic email templates file found at /app/config/email_templates.yaml or config/email_templates.yaml - skipping"
    )

# Look for dynamic DOCX templates in production and local locations.
# Production (container): /app/config/docx_templates.yaml
# Local development: <project_root>/config/docx_templates.yaml
APP_DOCX_CONFIG_PATH = Path("/app/config") / "docx_templates.yaml"
LOCAL_DOCX_CONFIG_PATH = Path(__file__).resolve().parent / "config" / "docx_templates.yaml"

_docx_yaml = None
for candidate in (APP_DOCX_CONFIG_PATH, LOCAL_DOCX_CONFIG_PATH):
    if candidate.exists():
        _docx_yaml = candidate
        logger.info("[dynamic-docx] Found DOCX templates file: %s", candidate)
        break

if _docx_yaml:
    try:
        register_docx_template_tools_from_yaml(mcp, _docx_yaml)
    except Exception as e:
        logger.exception("[dynamic-docx] Failed to register DOCX templates from %s: %s", _docx_yaml, e)
else:
    logger.info(
        "[dynamic-docx] No dynamic DOCX templates file found at /app/config/docx_templates.yaml or config/docx_templates.yaml - skipping"
    )

class PowerPointSlide(BaseModel):
    """PowerPoint slide - can be title, section, content, or table slide based on slide_type."""
    slide_type: Literal["title", "section", "content", "table"] = Field(description="Type of slide: 'title' for presentation opening, 'section' for dividers, 'content' for slide with bullet points, 'table' for slide with a table")
    slide_title: str = Field(description="Title text for the slide")

    # Optional fields based on slide type
    author: Optional[str] = Field(default="", description="Author name for title slides - appears in subtitle placeholder. Leave empty for section/content/table slides.")
    slide_text: Optional[List[Dict]] = Field(
        default=None,
        description="Array of bullet points for content slides. Each bullet point must have 'text' (string) and 'indentation_level' (integer 1-5). Leave empty/null for title, section, and table slides."
    )
    table_data: Optional[List[List[str]]] = Field(
        default=None,
        description="Table data for table slides. A list of rows where each row is a list of cell values (strings). The first row is treated as the header row. Leave empty/null for title, section, and content slides."
    )

@mcp.tool(
    name="create_excel_from_markdown",
    description="Converts markdown content with tables and formulas to Excel (.xlsx) format. Use '## Sheet: Name' headings to create multiple sheets.",
    tags={"excel", "spreadsheet", "data"},
    annotations={"title": "Markdown to Excel Converter"}
)
async def create_excel_document(
    markdown_content: Annotated[str, Field(description="Markdown content containing tables, headers, and formulas. Use '## Sheet: Sheet Name' to create multiple worksheets. Use T1.B[0] for cross-table references and B[0] for current row references. Use SheetName!T1.B[0] for cross-sheet references (resolves to SheetName!B2 in Excel). ALWAYS use [0], [1], [2] notation, NEVER use absolute row numbers like B2, B3. Do NOT count table header as first row, first row has index [0]. Supports cell formatting: **bold**, *italic*.")],
    file_name: Annotated[Optional[str], Field(description="Custom filename for the output file (without extension). If not provided, a unique identifier will be used.", default=None)] = None,
) -> str:
    """
    Converts markdown to Excel with advanced formula support.
    """

    logger.info("Converting markdown to Excel document")

    try:
        result = markdown_to_excel(markdown_content, file_name=file_name)
        logger.info("Excel document uploaded successfully")
        return result
    except Exception as e:
        logger.error(f"Error creating Excel document: {e}", exc_info=True)
        raise ToolError(f"Error creating Excel document: {e}")

@mcp.tool(
    name="create_word_from_markdown",
    description="Converts markdown content to a professionally formatted Word (.docx) document. Supports headings, lists, tables, images, block quotes, page breaks, horizontal lines, text alignment, and rich inline formatting.",
    tags={"word", "document", "text", "legal", "contract"},
    annotations={"title": "Markdown to Word Converter"}
)
async def create_word_document(
    markdown_content: Annotated[str, Field(description=(
        "Markdown content for the document body. Separate all block elements with blank lines.\n"
        "\n"
        "BLOCK ELEMENTS (each on its own line):\n"
        "- Headings: # H1, ## H2, ### H3, #### H4, ##### H5, ###### H6\n"
        "- Unordered lists: - item (or * or +); nest with 3-space indent\n"
        "- Ordered lists: 1. item, 2. item; nest with 3-space indent\n"
        "- Tables: | H1 | H2 |\\n|---|---|\\n| C1 | C2 | (cells support inline formatting)\n"
        "- Block quotes: > text (supports inline formatting)\n"
        "- Page break: --- (three+ dashes alone on a line — starts new page)\n"
        "- Horizontal line: *** (three+ asterisks alone on a line — visual separator)\n"
        "- Images: ![alt text](url)\n"
        "\n"
        "INLINE FORMATTING (usable in paragraphs, headings, lists, tables, quotes):\n"
        "- **bold**, *italic*, ***bold italic***\n"
        "- ~~strikethrough~~, __underline__ (double underscore — NOT bold)\n"
        "- `code` (Courier New font)\n"
        "- [link text](https://url)\n"
        "- Nesting: **bold with *italic* inside**, *italic with **bold** inside*\n"
        "- Combinations: **~~bold strikethrough~~**, **__bold underline__**, *~~italic strikethrough~~*\n"
        "- Escaped literals: \\* \\** \\` to render *, **, ` without formatting\n"
        "\n"
        "TEXT ALIGNMENT (HTML tags):\n"
        "- <center>text</center> or multi-line: <center>\\nline1\\nline2\\n</center>\n"
        "- <div align=\"right|center|justify|left\">text</div> (single or multi-line)\n"
        "\n"
        "LINE BREAKS: End a line with two trailing spaces for a soft break within the same paragraph.\n"
        "\n"
        "CONVENTIONS:\n"
        "- Do NOT confuse --- (page break) with *** (horizontal line).\n"
        "- Do NOT confuse __underline__ with bold; always use **bold** for bold.\n"
        "- LEGAL CONTRACTS: use numbered lists (1., 2., 3.) for clauses, nested lists for provisions; use headings only for the contract title.\n"
        "- Other documents: use headings (# ## ###) to organize sections.\n"
    ))],
    title: Annotated[Optional[str], Field(description="Document title (shown in file properties)", default=None)] = None,
    author: Annotated[Optional[str], Field(description="Document author name (shown in file properties)", default=None)] = None,
    subject: Annotated[Optional[str], Field(description="Document subject/description (shown in file properties)", default=None)] = None,
    header_text: Annotated[Optional[str], Field(description="Text for document header (top of every page). Use {page} for auto page number, {pages} for total pages.", default=None)] = None,
    footer_text: Annotated[Optional[str], Field(description="Text for document footer (bottom of every page). Use {page} for auto page number, {pages} for total pages.", default=None)] = None,
    include_toc: Annotated[Optional[bool], Field(description="If true, inserts a Table of Contents at the beginning of the document. The TOC updates automatically when opened in Word.", default=False)] = False,
    file_name: Annotated[Optional[str], Field(description="Custom filename for the output file (without extension). If not provided, a unique identifier will be used.", default=None)] = None,
) -> str:
    """
    Converts markdown to professionally formatted Word document.

    """

    logger.info("Converting markdown to Word document")

    try:
        result = markdown_to_word(
            markdown_content,
            title=title,
            author=author,
            subject=subject,
            header_text=header_text,
            footer_text=footer_text,
            include_toc=include_toc or False,
            file_name=file_name,
        )
        logger.info("Word document uploaded successfully")
        return result
    except Exception as e:
        logger.error(f"Error creating Word document: {e}", exc_info=True)
        raise ToolError(f"Error creating Word document: {e}")

@mcp.tool(
    name="create_powerpoint_presentation",
    description="Creates PowerPoint presentations from structured slides.",
    tags={"powerpoint", "presentation", "slides"},
    annotations={"title": "PowerPoint Presentation Creator"}
)
async def create_powerpoint_presentation(
    slides: Annotated[List[dict], Field(
        description="""List of slide objects. Each slide requires 'slide_type' (str) and type-specific fields:

- title: {slide_type: "title", slide_title: str, author?: str}
- section: {slide_type: "section", slide_title: str}
- content: {slide_type: "content", slide_title: str, slide_text: [{text: str, indentation_level: int (1-3)}]}
- table: {slide_type: "table", slide_title: str, table_data: [[str]] (first row = header), header_color?: str (hex), alternate_rows?: bool}
- image: {slide_type: "image", slide_title?: str, image_url: str, image_caption?: str}
- two_column: {slide_type: "two_column", slide_title: str, left_column: [{text: str, indentation_level: int}], right_column: [{text: str, indentation_level: int}], left_heading?: str, right_heading?: str}
- chart: {slide_type: "chart", slide_title: str, chart_type: str (bar|column|line|pie|doughnut|stacked_bar|area), chart_data: {categories: [str], series: [{name: str, values: [number]}]}, has_legend?: bool, legend_position?: str}
- quote: {slide_type: "quote", slide_title?: str, quote_text: str, quote_author?: str}

All slides support optional 'speaker_notes': str field."""
    )],
    format: Annotated[Literal["4:3", "16:9"], Field(
        default="16:9",
        description="Aspect ratio: '16:9' (widescreen) or '4:3' (traditional)"
    )] = "16:9",
    file_name: Annotated[Optional[str], Field(description="Custom filename for the output file (without extension). If not provided, a unique identifier will be used.", default=None)] = None,
) -> str:
    """Creates PowerPoint presentations with structured slide models and professional templates."""

    logger.info(f"Creating PowerPoint presentation with {len(slides)} slides in {format} format")

    try:
        result = create_presentation(slides, format, file_name=file_name)
        logger.info(f"PowerPoint presentation created: {result}")
        return result
    except Exception as e:
        logger.error(f"Error creating PowerPoint presentation: {e}", exc_info=True)
        raise ToolError(f"Error creating PowerPoint presentation: {e}")

@mcp.tool(
    name="create_email_draft",
    description="Creates an email draft in EML format with HTML content using preset professional styling.",
    tags={"email", "eml", "communication"},
    annotations={"title": "Email Draft Creator"}
)
async def create_email_draft(
    content: Annotated[str, Field(description="BODY CONTENT ONLY - Do NOT include HTML structure tags like <html>, <head>, <body>, or <style>. Do NOT include any CSS styling. Use <p> for greetings and for signatures, never headers. Use <h2> for section headers (will be bold), <h3> for subsection headers (will be underlined). HTML tags allowed: <p>, <h2>, <h3>, <ul>, <li>, <strong>, <em>, <div>.")],
    subject: Annotated[str, Field(description="Email subject line")],
    to: Annotated[Optional[List[str]], Field(description="List of recipient email addresses", default=None)],
    cc: Annotated[Optional[List[str]], Field(description="List of CC recipient email addresses", default=None)],
    bcc: Annotated[Optional[List[str]], Field(description="List of BCC recipient email addresses", default=None)],
    priority: Annotated[str, Field(description="Email priority: 'low', 'normal', or 'high'", default="normal")],
    language: Annotated[str, Field(description="Language code for proofreading in Outlook (e.g., 'cs-CZ' for Czech, 'en-US' for English, 'de-DE' for German, 'sk-SK' for Slovak)", default="cs-CZ")],
    file_name: Annotated[Optional[str], Field(description="Custom filename for the output file (without extension). If not provided, a unique identifier will be used.", default=None)] = None,
) -> str:
    """
    Creates professional email drafts in EML format with preset styling and language settings.
    """

    logger.info(f"Creating email draft with subject: {subject}")

    try:
        result = create_eml(
            to=to,
            cc=cc,
            bcc=bcc,
            re=subject,
            content=content,
            priority=priority,
            language=language,
            file_name=file_name,
        )
        logger.info(f"Email draft created: {result}")
        return result
    except Exception as e:
        logger.error(f"Error creating email draft: {e}", exc_info=True)
        raise ToolError(f"Error creating email draft: {e}")

@mcp.tool(
    name="create_xml_file",
    description="Creates an XML file from provided XML content.",
    tags={"xml", "data", "configuration"},
    annotations={"title": "XML File Creator"}
)
async def create_xml_document(
    xml_content: Annotated[str, Field(description="Complete, well-formed XML content. Must be valid XML with proper opening and closing tags.")],
    file_name: Annotated[Optional[str], Field(description="Custom filename for the output file (without extension). If not provided, a unique identifier will be used.", default=None)] = None,
) -> str:
    """
    Creates an XML file from provided XML content.
    Validates that the XML is well-formed before saving.
    """

    logger.info("Creating XML file")

    try:
        result = create_xml_file(xml_content, file_name=file_name)
        logger.info("XML file created successfully.")
        return result
    except Exception as e:
        logger.error(f"Error creating XML file: {e}", exc_info=True)
        raise ToolError(f"Error creating XML file: {e}")

# === Read Tools ===
from read_tools import (
    read_docx, get_docx_info, get_docx_paragraphs, get_docx_tables,
    read_xlsx, get_xlsx_info, get_xlsx_sheets,
    read_pptx, get_pptx_info, get_pptx_slides,
)


@mcp.tool(name="read_document", description="Extract all text from an existing document (DOCX, XLSX, or PPTX)", tags=["read"])
async def tool_read_document(
    file_path: Annotated[str, Field(description="Absolute path to the document file (.docx, .xlsx, or .pptx)")],
) -> str:
    """Read and extract text from an existing Office document."""
    path = file_path.lower()
    try:
        if path.endswith(".docx"):
            return read_docx(file_path)
        elif path.endswith(".xlsx"):
            return read_xlsx(file_path)
        elif path.endswith(".pptx"):
            return read_pptx(file_path)
        else:
            raise ToolError(f"Unsupported file format. Supported: .docx, .xlsx, .pptx")
    except ToolError:
        raise
    except Exception as e:
        raise ToolError(f"Error reading document: {e}")


@mcp.tool(name="get_document_info", description="Get metadata and statistics from an existing document", tags=["read"])
async def tool_get_document_info(
    file_path: Annotated[str, Field(description="Absolute path to the document file (.docx, .xlsx, or .pptx)")],
) -> dict:
    """Get document metadata: author, dates, page/slide/sheet count, etc."""
    path = file_path.lower()
    try:
        if path.endswith(".docx"):
            return get_docx_info(file_path)
        elif path.endswith(".xlsx"):
            return get_xlsx_info(file_path)
        elif path.endswith(".pptx"):
            return get_pptx_info(file_path)
        else:
            raise ToolError(f"Unsupported file format. Supported: .docx, .xlsx, .pptx")
    except ToolError:
        raise
    except Exception as e:
        raise ToolError(f"Error getting document info: {e}")


@mcp.tool(name="get_docx_paragraphs", description="List all paragraphs from a DOCX file with index and style", tags=["read", "docx"])
async def tool_get_docx_paragraphs(
    file_path: Annotated[str, Field(description="Absolute path to the .docx file")],
) -> list[dict]:
    """Get paragraphs with index, style name, and text content."""
    try:
        return get_docx_paragraphs(file_path)
    except Exception as e:
        raise ToolError(f"Error reading paragraphs: {e}")


@mcp.tool(name="get_docx_tables", description="Extract tables from a DOCX file with content", tags=["read", "docx"])
async def tool_get_docx_tables(
    file_path: Annotated[str, Field(description="Absolute path to the .docx file")],
) -> list[dict]:
    """Get tables with row/column count and cell data."""
    try:
        return get_docx_tables(file_path)
    except Exception as e:
        raise ToolError(f"Error reading tables: {e}")


@mcp.tool(name="get_xlsx_sheets", description="Get sheet content from an XLSX file", tags=["read", "xlsx"])
async def tool_get_xlsx_sheets(
    file_path: Annotated[str, Field(description="Absolute path to the .xlsx file")],
    sheet_name: Annotated[Optional[str], Field(description="Specific sheet name (omit for all sheets)", default=None)] = None,
    max_rows: Annotated[int, Field(description="Maximum rows to return per sheet", default=100)] = 100,
) -> list[dict]:
    """Get sheet data as rows. Optionally filter by sheet name."""
    try:
        return get_xlsx_sheets(file_path, sheet_name=sheet_name, max_rows=max_rows)
    except Exception as e:
        raise ToolError(f"Error reading sheets: {e}")


@mcp.tool(name="get_pptx_slides", description="List all slides from a PPTX file with content", tags=["read", "pptx"])
async def tool_get_pptx_slides(
    file_path: Annotated[str, Field(description="Absolute path to the .pptx file")],
) -> list[dict]:
    """Get slides with index, layout, shape count, and text content."""
    try:
        return get_pptx_slides(file_path)
    except Exception as e:
        raise ToolError(f"Error reading slides: {e}")


# === Advanced DOCX Tools ===
from docx_tools.advanced import (
    add_image_to_docx, add_header_footer, set_page_margins,
    add_bullet_list, add_numbered_list, merge_table_cells,
)


@mcp.tool(name="add_image_to_docx", description="Add an image to a DOCX file", tags=["edit", "docx", "advanced"])
async def tool_add_image(
    file_path: Annotated[str, Field(description="Absolute path to the .docx file")],
    image_path: Annotated[str, Field(description="Absolute path to the image file")],
    width_inches: Annotated[float, Field(description="Image width in inches", default=4.0)] = 4.0,
    output_path: Annotated[Optional[str], Field(description="Save to different path", default=None)] = None,
) -> str:
    try:
        return add_image_to_docx(file_path, image_path, width_inches, output_path=output_path)
    except Exception as e:
        raise ToolError(f"Error adding image: {e}")


@mcp.tool(name="add_header_footer", description="Add header and/or footer to a DOCX file", tags=["edit", "docx", "advanced"])
async def tool_add_header_footer(
    file_path: Annotated[str, Field(description="Absolute path to the .docx file")],
    header_text: Annotated[str, Field(description="Header text (empty to skip)", default="")] = "",
    footer_text: Annotated[str, Field(description="Footer text (empty to skip)", default="")] = "",
    output_path: Annotated[Optional[str], Field(description="Save to different path", default=None)] = None,
) -> str:
    try:
        return add_header_footer(file_path, header_text, footer_text, output_path)
    except Exception as e:
        raise ToolError(f"Error adding header/footer: {e}")


@mcp.tool(name="set_page_margins", description="Set page margins for a DOCX file (in cm)", tags=["edit", "docx", "advanced"])
async def tool_set_page_margins(
    file_path: Annotated[str, Field(description="Absolute path to the .docx file")],
    top_cm: Annotated[float, Field(description="Top margin in cm", default=2.54)] = 2.54,
    bottom_cm: Annotated[float, Field(description="Bottom margin in cm", default=2.54)] = 2.54,
    left_cm: Annotated[float, Field(description="Left margin in cm", default=2.54)] = 2.54,
    right_cm: Annotated[float, Field(description="Right margin in cm", default=2.54)] = 2.54,
    output_path: Annotated[Optional[str], Field(description="Save to different path", default=None)] = None,
) -> str:
    try:
        return set_page_margins(file_path, top_cm, bottom_cm, left_cm, right_cm, output_path)
    except Exception as e:
        raise ToolError(f"Error setting margins: {e}")


@mcp.tool(name="add_bullet_list", description="Add a bullet list to a DOCX file", tags=["edit", "docx", "advanced"])
async def tool_add_bullet_list(
    file_path: Annotated[str, Field(description="Absolute path to the .docx file")],
    items: Annotated[List[str], Field(description="List items")],
    output_path: Annotated[Optional[str], Field(description="Save to different path", default=None)] = None,
) -> str:
    try:
        return add_bullet_list(file_path, items, output_path)
    except Exception as e:
        raise ToolError(f"Error adding list: {e}")


@mcp.tool(name="add_numbered_list", description="Add a numbered list to a DOCX file", tags=["edit", "docx", "advanced"])
async def tool_add_numbered_list(
    file_path: Annotated[str, Field(description="Absolute path to the .docx file")],
    items: Annotated[List[str], Field(description="List items")],
    output_path: Annotated[Optional[str], Field(description="Save to different path", default=None)] = None,
) -> str:
    try:
        return add_numbered_list(file_path, items, output_path)
    except Exception as e:
        raise ToolError(f"Error adding list: {e}")


@mcp.tool(name="merge_table_cells", description="Merge cells in a DOCX table", tags=["edit", "docx", "advanced"])
async def tool_merge_table_cells(
    file_path: Annotated[str, Field(description="Absolute path to the .docx file")],
    table_index: Annotated[int, Field(description="Table index (0-based)")],
    start_row: Annotated[int, Field(description="Start row (0-based)")],
    start_col: Annotated[int, Field(description="Start column (0-based)")],
    end_row: Annotated[int, Field(description="End row (0-based)")],
    end_col: Annotated[int, Field(description="End column (0-based)")],
    output_path: Annotated[Optional[str], Field(description="Save to different path", default=None)] = None,
) -> str:
    try:
        return merge_table_cells(file_path, table_index, start_row, start_col, end_row, end_col, output_path)
    except Exception as e:
        raise ToolError(f"Error merging cells: {e}")


# === PDF Tools ===
from pdf_tools import markdown_to_pdf, docx_to_pdf


@mcp.tool(name="create_pdf_from_markdown", description="Generate a PDF document from Markdown content", tags=["create", "pdf"])
async def tool_create_pdf_from_markdown(
    markdown_content: Annotated[str, Field(description="Markdown text to convert to PDF")],
    file_name: Annotated[Optional[str], Field(description="Output file name (without extension)", default=None)] = None,
) -> str:
    """Create a PDF from Markdown with headings, lists, paragraphs."""
    try:
        result = markdown_to_pdf(markdown_content, file_name=file_name)
        logger.info("PDF created from Markdown successfully.")
        return result
    except Exception as e:
        logger.error(f"Error creating PDF: {e}", exc_info=True)
        raise ToolError(f"Error creating PDF: {e}")


@mcp.tool(name="convert_docx_to_pdf", description="Convert an existing DOCX file to PDF", tags=["convert", "pdf"])
async def tool_convert_docx_to_pdf(
    file_path: Annotated[str, Field(description="Absolute path to the .docx file to convert")],
    file_name: Annotated[Optional[str], Field(description="Output file name (without extension)", default=None)] = None,
) -> str:
    """Convert a Word document to PDF preserving structure."""
    try:
        result = docx_to_pdf(file_path, file_name=file_name)
        logger.info("DOCX converted to PDF successfully.")
        return result
    except Exception as e:
        logger.error(f"Error converting to PDF: {e}", exc_info=True)
        raise ToolError(f"Error converting to PDF: {e}")


# === Edit Tools ===
from edit_tools import (
    edit_docx_paragraph, delete_docx_paragraph, search_replace_docx, insert_docx_paragraph,
    edit_xlsx_cell, insert_xlsx_row, delete_xlsx_row,
    edit_pptx_slide_text, delete_pptx_slide, reorder_pptx_slides,
)


@mcp.tool(name="edit_docx_paragraph", description="Edit a paragraph in an existing DOCX file by index", tags=["edit", "docx"])
async def tool_edit_docx_paragraph(
    file_path: Annotated[str, Field(description="Absolute path to the .docx file")],
    index: Annotated[int, Field(description="Paragraph index (0-based)")],
    new_text: Annotated[str, Field(description="New text for the paragraph")],
    output_path: Annotated[Optional[str], Field(description="Save to different path (omit to overwrite)", default=None)] = None,
) -> str:
    try:
        return edit_docx_paragraph(file_path, index, new_text, output_path)
    except Exception as e:
        raise ToolError(f"Error editing paragraph: {e}")


@mcp.tool(name="delete_docx_paragraph", description="Delete a paragraph from a DOCX file by index", tags=["edit", "docx"])
async def tool_delete_docx_paragraph(
    file_path: Annotated[str, Field(description="Absolute path to the .docx file")],
    index: Annotated[int, Field(description="Paragraph index to delete (0-based)")],
    output_path: Annotated[Optional[str], Field(description="Save to different path (omit to overwrite)", default=None)] = None,
) -> str:
    try:
        return delete_docx_paragraph(file_path, index, output_path)
    except Exception as e:
        raise ToolError(f"Error deleting paragraph: {e}")


@mcp.tool(name="search_replace_docx", description="Search and replace text in a DOCX file", tags=["edit", "docx"])
async def tool_search_replace_docx(
    file_path: Annotated[str, Field(description="Absolute path to the .docx file")],
    search_text: Annotated[str, Field(description="Text to search for")],
    replace_text: Annotated[str, Field(description="Replacement text")],
    output_path: Annotated[Optional[str], Field(description="Save to different path (omit to overwrite)", default=None)] = None,
) -> str:
    try:
        return search_replace_docx(file_path, search_text, replace_text, output_path)
    except Exception as e:
        raise ToolError(f"Error in search/replace: {e}")


@mcp.tool(name="insert_docx_paragraph", description="Insert a new paragraph at a specific position in a DOCX file", tags=["edit", "docx"])
async def tool_insert_docx_paragraph(
    file_path: Annotated[str, Field(description="Absolute path to the .docx file")],
    index: Annotated[int, Field(description="Position to insert at (0-based)")],
    text: Annotated[str, Field(description="Text for the new paragraph")],
    style: Annotated[str, Field(description="Paragraph style (e.g. 'Normal', 'Heading 1')", default="Normal")] = "Normal",
    output_path: Annotated[Optional[str], Field(description="Save to different path (omit to overwrite)", default=None)] = None,
) -> str:
    try:
        return insert_docx_paragraph(file_path, index, text, style, output_path)
    except Exception as e:
        raise ToolError(f"Error inserting paragraph: {e}")


@mcp.tool(name="edit_xlsx_cell", description="Edit a cell value in an XLSX file", tags=["edit", "xlsx"])
async def tool_edit_xlsx_cell(
    file_path: Annotated[str, Field(description="Absolute path to the .xlsx file")],
    sheet_name: Annotated[str, Field(description="Sheet name")],
    cell: Annotated[str, Field(description="Cell reference (e.g. 'A1', 'B2')")],
    value: Annotated[str, Field(description="New cell value")],
    output_path: Annotated[Optional[str], Field(description="Save to different path (omit to overwrite)", default=None)] = None,
) -> str:
    try:
        return edit_xlsx_cell(file_path, sheet_name, cell, value, output_path)
    except Exception as e:
        raise ToolError(f"Error editing cell: {e}")


@mcp.tool(name="insert_xlsx_row", description="Insert a new row in an XLSX sheet", tags=["edit", "xlsx"])
async def tool_insert_xlsx_row(
    file_path: Annotated[str, Field(description="Absolute path to the .xlsx file")],
    sheet_name: Annotated[str, Field(description="Sheet name")],
    row_index: Annotated[int, Field(description="Row position to insert at (1-based)")],
    values: Annotated[List[str], Field(description="List of cell values for the new row")],
    output_path: Annotated[Optional[str], Field(description="Save to different path (omit to overwrite)", default=None)] = None,
) -> str:
    try:
        return insert_xlsx_row(file_path, sheet_name, row_index, values, output_path)
    except Exception as e:
        raise ToolError(f"Error inserting row: {e}")


@mcp.tool(name="delete_xlsx_row", description="Delete a row from an XLSX sheet", tags=["edit", "xlsx"])
async def tool_delete_xlsx_row(
    file_path: Annotated[str, Field(description="Absolute path to the .xlsx file")],
    sheet_name: Annotated[str, Field(description="Sheet name")],
    row_index: Annotated[int, Field(description="Row to delete (1-based)")],
    output_path: Annotated[Optional[str], Field(description="Save to different path (omit to overwrite)", default=None)] = None,
) -> str:
    try:
        return delete_xlsx_row(file_path, sheet_name, row_index, output_path)
    except Exception as e:
        raise ToolError(f"Error deleting row: {e}")


@mcp.tool(name="edit_pptx_slide_text", description="Replace text in a specific PPTX slide", tags=["edit", "pptx"])
async def tool_edit_pptx_slide_text(
    file_path: Annotated[str, Field(description="Absolute path to the .pptx file")],
    slide_index: Annotated[int, Field(description="Slide index (0-based)")],
    old_text: Annotated[str, Field(description="Text to find")],
    new_text: Annotated[str, Field(description="Replacement text")],
    output_path: Annotated[Optional[str], Field(description="Save to different path (omit to overwrite)", default=None)] = None,
) -> str:
    try:
        return edit_pptx_slide_text(file_path, slide_index, old_text, new_text, output_path)
    except Exception as e:
        raise ToolError(f"Error editing slide: {e}")


@mcp.tool(name="delete_pptx_slide", description="Delete a slide from a PPTX file", tags=["edit", "pptx"])
async def tool_delete_pptx_slide(
    file_path: Annotated[str, Field(description="Absolute path to the .pptx file")],
    slide_index: Annotated[int, Field(description="Slide index to delete (0-based)")],
    output_path: Annotated[Optional[str], Field(description="Save to different path (omit to overwrite)", default=None)] = None,
) -> str:
    try:
        return delete_pptx_slide(file_path, slide_index, output_path)
    except Exception as e:
        raise ToolError(f"Error deleting slide: {e}")


@mcp.tool(name="reorder_pptx_slides", description="Reorder slides in a PPTX file", tags=["edit", "pptx"])
async def tool_reorder_pptx_slides(
    file_path: Annotated[str, Field(description="Absolute path to the .pptx file")],
    new_order: Annotated[List[int], Field(description="New slide order as list of indices (e.g. [2, 0, 1])")],
    output_path: Annotated[Optional[str], Field(description="Save to different path (omit to overwrite)", default=None)] = None,
) -> str:
    try:
        return reorder_pptx_slides(file_path, new_order, output_path)
    except Exception as e:
        raise ToolError(f"Error reordering slides: {e}")


if __name__ == "__main__":
    mcp.run(
        transport="streamable-http",
        host="0.0.0.0",
        port=8958,
        log_level=config.logging.mcp_level_str,
        path="/mcp"
    )
