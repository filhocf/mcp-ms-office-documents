"""Tools for editing existing Office documents (DOCX, XLSX, PPTX)."""

from edit_tools.docx_editor import edit_docx_paragraph, delete_docx_paragraph, search_replace_docx, insert_docx_paragraph
from edit_tools.xlsx_editor import edit_xlsx_cell, insert_xlsx_row, delete_xlsx_row
from edit_tools.pptx_editor import edit_pptx_slide_text, delete_pptx_slide, reorder_pptx_slides

__all__ = [
    "edit_docx_paragraph", "delete_docx_paragraph", "search_replace_docx", "insert_docx_paragraph",
    "edit_xlsx_cell", "insert_xlsx_row", "delete_xlsx_row",
    "edit_pptx_slide_text", "delete_pptx_slide", "reorder_pptx_slides",
]
