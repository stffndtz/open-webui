"""
title: NDA Redline Tool 2.0
author: 42frontiers
author_url: https://42frontiers.com
version: 1.0.0
license: MIT
description: A tool for analyzing and redlining NDA documents. Reads DOCX files, allows block-based editing, and generates tracked-changes (redline) documents.
requirements: python-docx, python-redlines
"""

import os
import re
import uuid
import logging
import datetime
from collections import deque
from io import BytesIO
from typing import Dict, List, Any, Optional, Literal

from pydantic import BaseModel, Field
from pydantic.fields import FieldInfo

import docx
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from python_redlines.engines import XmlPowerToolsEngine

# Open WebUI imports - these are available when running inside Open WebUI
try:
    from open_webui.models.files import Files, FileForm
    from open_webui.storage.provider import Storage
    OPENWEBUI_AVAILABLE = True
except ImportError:
    OPENWEBUI_AVAILABLE = False
    Files = None
    FileForm = None
    Storage = None

# Configure logging
log = logging.getLogger(__name__)


# =============================================================================
# WORD EDITOR - Internal document manipulation engine
# =============================================================================


class _WordEditor:
    """
    Internal document editor that manages DOCX file manipulation.

    Provides block-based access to document content where each paragraph
    and table is assigned a unique ID (e.g., 'block_p_0', 'block_t_1').
    This allows precise, targeted edits while preserving document formatting.

    Not exposed directly to the model - used internally by the Tools class.
    """

    def __init__(self, file_path: str):
        """
        Load a DOCX document and create the block-to-ID mapping.

        Args:
            file_path: Absolute path to the .docx file

        Raises:
            ValueError: If file_path is empty
            FileNotFoundError: If file doesn't exist
        """
        if not file_path or not file_path.strip():
            raise ValueError("File path cannot be empty")

        self.file_path = file_path
        self.document = docx.Document(file_path)

        # Bidirectional mapping between block IDs and document elements
        self._id_to_element: Dict[str, Any] = {}
        self._element_to_id: Dict[Any, str] = {}
        self._next_id = 0

        # Style cache for consistent formatting when inserting new blocks
        self._saved_styles: Dict[str, dict] = {}

        self._build_block_map()

    def _build_block_map(self) -> None:
        """Scan document body and assign unique IDs to each paragraph and table."""
        for child in self.document.element.body:
            if isinstance(child, CT_P):
                para = Paragraph(child, self.document)
                block_id = f"block_p_{self._next_id}"
                self._next_id += 1
                self._id_to_element[block_id] = para
                self._element_to_id[para] = block_id
            elif isinstance(child, CT_Tbl):
                table = Table(child, self.document)
                block_id = f"block_t_{self._next_id}"
                self._next_id += 1
                self._id_to_element[block_id] = table
                self._element_to_id[table] = block_id

    def _rebuild_block_map(self) -> None:
        """Rebuild the block map after structural changes (insert/delete)."""
        self._id_to_element.clear()
        self._element_to_id.clear()
        self._next_id = 0
        self._build_block_map()

    def _classify_table(self, table_data: List[List[str]], row_count: int, col_count: int) -> Dict[str, Any]:
        """
        Classify a table as metadata (OK to ignore) or content (problematic).

        Metadata tables: Small tables with addresses, signatures, names, dates
        Content tables: Large tables with NDA clauses, often bilingual

        Returns:
            dict with 'is_metadata', 'reason', and 'confidence'
        """
        # Flatten all text for keyword analysis
        all_text = " ".join(" ".join(row) for row in table_data).lower()

        # NDA content keywords (if found in large tables, it's content)
        content_keywords = [
            "vertraulich", "confidential", "geheimhaltung", "non-disclosure",
            "offenlegung", "disclosure", "verpflichtung", "obligation",
            "laufzeit", "duration", "term", "geltungsbereich", "scope",
            "präambel", "preamble", "vertragspartei", "party", "parties",
            "schadensersatz", "damages", "haftung", "liability",
            "definition", "zweck", "purpose", "rückgabe", "return",
            "vernichtung", "destruction", "gerichtsstand", "jurisdiction"
        ]

        # Metadata keywords (expected in small signature/address tables)
        metadata_keywords = [
            "name", "address", "adresse", "signature", "unterschrift",
            "date", "datum", "tel", "fax", "email", "phone",
            "company", "firma", "street", "straße", "city", "stadt",
            "zip", "plz", "country", "land"
        ]

        content_matches = sum(1 for kw in content_keywords if kw in all_text)
        metadata_matches = sum(1 for kw in metadata_keywords if kw in all_text)

        # Classification rules:
        # 1. Small tables (≤5 rows) are almost always metadata
        if row_count <= 5:
            return {
                "is_metadata": True,
                "reason": f"Small table ({row_count} rows) - likely addresses/signatures",
                "confidence": "high"
            }

        # 2. Large tables (>10 rows) with content keywords are NDA content
        if row_count > 10 and content_matches >= 2:
            return {
                "is_metadata": False,
                "reason": f"Large table ({row_count} rows) with NDA content keywords",
                "confidence": "high"
            }

        # 3. 2-3 column tables with substantial text in multiple rows = bilingual
        if col_count in (2, 3) and row_count > 5:
            # Check if rows have substantial text (>50 chars average)
            avg_text_length = sum(len(cell) for row in table_data for cell in row) / max(1, row_count * col_count)
            if avg_text_length > 50:
                return {
                    "is_metadata": False,
                    "reason": f"Bilingual table format ({col_count} columns, {row_count} rows with substantial text)",
                    "confidence": "high"
                }

        # 4. Medium tables (6-10 rows) - check keywords
        if content_matches > metadata_matches:
            return {
                "is_metadata": False,
                "reason": f"Medium table with more content keywords ({content_matches}) than metadata ({metadata_matches})",
                "confidence": "medium"
            }

        # Default: assume metadata if small/medium and no clear content indicators
        return {
            "is_metadata": True,
            "reason": f"Table ({row_count} rows) without clear NDA content indicators",
            "confidence": "medium"
        }

    def get_structured_content(self) -> Dict[str, Any]:
        """
        Return document content as a structured dict with block IDs.

        Returns:
            dict with 'file_path' and 'blocks' list, where each block has:
                - id: unique block identifier
                - type: 'paragraph', 'heading', 'list_item', or 'table'
                - text: content (for paragraphs) or data (for tables)
                - level: heading level (only for headings)
        """
        blocks = []

        for element, block_id in self._element_to_id.items():
            block_info = {"id": block_id}

            if isinstance(element, Paragraph):
                # Detect paragraph type from style
                style_name = element.style.name if element.style else ""

                if style_name.startswith("Heading"):
                    block_info["type"] = "heading"
                    try:
                        block_info["level"] = int(style_name[-1])
                    except (ValueError, IndexError):
                        block_info["level"] = 1
                elif style_name.startswith("List"):
                    block_info["type"] = "list_item"
                else:
                    block_info["type"] = "paragraph"

                block_info["text"] = element.text

            elif isinstance(element, Table):
                block_info["type"] = "table"
                row_count = len(element.rows)
                col_count = len(element.columns) if element.rows else 0
                block_info["rows"] = row_count
                block_info["columns"] = col_count

                # Include table data for review
                table_data = [
                    [cell.text for cell in row.cells]
                    for row in element.rows
                ]
                block_info["data"] = table_data

                # Classify the table
                classification = self._classify_table(table_data, row_count, col_count)
                block_info["is_metadata"] = classification["is_metadata"]
                block_info["classification_reason"] = classification["reason"]

                if classification["is_metadata"]:
                    block_info["note"] = "Metadata table (addresses/signatures) - can be ignored"
                else:
                    block_info["editable"] = False
                    block_info["warning"] = (
                        "⚠️ CONTENT TABLE - Contains NDA clauses that cannot be edited programmatically. "
                        "Document processing cannot continue."
                    )

                # Add preview of first row as potential headers
                if element.rows:
                    first_row = [cell.text[:50] for cell in element.rows[0].cells]
                    block_info["header_preview"] = first_row

            blocks.append(block_info)

        # Build document summary
        table_blocks = [b for b in blocks if b["type"] == "table"]
        paragraph_blocks = [b for b in blocks if b["type"] in ("paragraph", "heading", "list_item")]

        # Separate metadata tables from content tables
        metadata_tables = [t for t in table_blocks if t.get("is_metadata", False)]
        content_tables = [t for t in table_blocks if not t.get("is_metadata", True)]

        summary = {
            "total_blocks": len(blocks),
            "paragraphs": len(paragraph_blocks),
            "tables": len(table_blocks),
        }

        # Determine if document can be processed
        if content_tables:
            # Content tables found - STOP
            summary["can_proceed"] = False
            summary["stop_reason"] = (
                f"❌ CANNOT PROCESS: This document contains {len(content_tables)} table(s) with NDA content "
                f"that cannot be edited programmatically. Tables: {[t['id'] for t in content_tables]}"
            )
            summary["content_table_ids"] = [t["id"] for t in content_tables]
            if metadata_tables:
                summary["metadata_table_ids"] = [t["id"] for t in metadata_tables]
        elif metadata_tables:
            # Only metadata tables - OK to proceed
            summary["can_proceed"] = True
            summary["table_note"] = (
                f"✅ Document has {len(metadata_tables)} metadata table(s) (addresses/signatures) - "
                f"these can be ignored. Proceed with compliance analysis on paragraph content."
            )
            summary["metadata_table_ids"] = [t["id"] for t in metadata_tables]
        else:
            # No tables at all - definitely OK
            summary["can_proceed"] = True

        return {
            "file_path": self.file_path,
            "summary": summary,
            "blocks": blocks
        }

    def edit_block(self, block_id: str, new_text: str) -> Dict[str, str]:
        """
        Replace the text content of a paragraph block while preserving formatting.

        Args:
            block_id: The ID of the block to edit (e.g., 'block_p_5')
            new_text: New text content to replace existing text

        Returns:
            dict with 'status' and 'message'

        Raises:
            ValueError: If block_id not found
            TypeError: If block is not a paragraph
        """
        if block_id not in self._id_to_element:
            raise ValueError(f"Block '{block_id}' not found in document")

        element = self._id_to_element[block_id]

        if not isinstance(element, Paragraph):
            raise TypeError(
                f"Block '{block_id}' is a table. Table cell editing is not supported due to "
                "technical limitations with tracked changes. Please note the table content in "
                "your compliance overview and recommend manual review of table contents."
            )

        # Capture existing formatting from first run
        formatting = self._extract_run_formatting(element)
        paragraph_style = element.style
        paragraph_alignment = element.alignment

        # Clear and set new text
        element.text = ""
        new_run = element.add_run(new_text)

        # Restore formatting
        self._apply_run_formatting(new_run, formatting)
        element.style = paragraph_style
        element.alignment = paragraph_alignment

        return {
            "status": "success",
            "message": f"Block '{block_id}' updated successfully"
        }

    def find_replace_in_block(
        self,
        block_id: str,
        find_text: str,
        replace_text: str,
        replace_all: bool = False
    ) -> Dict[str, Any]:
        """
        Find and replace specific text within a paragraph block.

        This is the preferred method for surgical edits - it only changes
        the exact text specified, preserving ALL existing formatting.
        The replacement text inherits the formatting of the text it replaces.

        Args:
            block_id: The ID of the block to edit (e.g., 'block_p_5')
            find_text: Exact text to find within the paragraph
            replace_text: Text to replace it with
            replace_all: If True, replace all occurrences; if False, only first

        Returns:
            dict with 'status', 'message', 'occurrences_found', 'occurrences_replaced'

        Raises:
            ValueError: If block_id not found or find_text not in paragraph
            TypeError: If block is not a paragraph
        """
        if block_id not in self._id_to_element:
            raise ValueError(f"Block '{block_id}' not found in document")

        element = self._id_to_element[block_id]

        if not isinstance(element, Paragraph):
            raise TypeError(
                f"Block '{block_id}' is a table. Use find_replace only on paragraphs."
            )

        # Get current text
        current_text = element.text

        # Check if find_text exists
        occurrences = current_text.count(find_text)
        if occurrences == 0:
            raise ValueError(
                f"Text '{find_text[:50]}{'...' if len(find_text) > 50 else ''}' "
                f"not found in block '{block_id}'. "
                f"Current text: '{current_text[:100]}{'...' if len(current_text) > 100 else ''}'"
            )

        # Perform in-place replacement within runs to preserve all formatting
        # Each run maintains its own formatting, so we replace text within each run
        replaced_count = 0
        max_replacements = occurrences if replace_all else 1

        for run in element.runs:
            if replaced_count >= max_replacements:
                break

            run_text = run.text
            if find_text in run_text:
                # Count how many we can replace in this run
                run_occurrences = run_text.count(find_text)
                can_replace = min(run_occurrences, max_replacements - replaced_count)

                if replace_all or can_replace > 0:
                    # Replace within this run (preserves run's formatting automatically)
                    if replace_all:
                        run.text = run_text.replace(find_text, replace_text)
                        replaced_count += run_occurrences
                    else:
                        run.text = run_text.replace(find_text, replace_text, 1)
                        replaced_count += 1

        # Handle case where find_text spans multiple runs (fallback to full replacement)
        if replaced_count == 0 and find_text in current_text:
            # Text spans runs - need to do full paragraph replacement
            # This is less ideal but necessary for cross-run matches
            new_text = current_text.replace(find_text, replace_text, 1 if not replace_all else -1)
            replaced_count = 1 if not replace_all else occurrences

            # Preserve formatting from first run
            formatting = self._extract_run_formatting(element)
            paragraph_style = element.style
            paragraph_alignment = element.alignment

            # Clear and rebuild
            element.text = ""
            new_run = element.add_run(new_text)

            # Restore formatting (but don't force underline off - preserve original)
            if formatting.get("name"):
                new_run.font.name = formatting["name"]
            if formatting.get("size"):
                new_run.font.size = formatting["size"]
            if formatting.get("bold") is not None:
                new_run.font.bold = formatting["bold"]
            if formatting.get("italic") is not None:
                new_run.font.italic = formatting["italic"]
            if formatting.get("underline") is not None:
                new_run.font.underline = formatting["underline"]
            if formatting.get("color"):
                new_run.font.color.rgb = formatting["color"]

            element.style = paragraph_style
            element.alignment = paragraph_alignment

        return {
            "status": "success",
            "message": f"Replaced '{find_text[:30]}{'...' if len(find_text) > 30 else ''}' with '{replace_text[:30]}{'...' if len(replace_text) > 30 else ''}' in block '{block_id}'",
            "occurrences_found": occurrences,
            "occurrences_replaced": replaced_count
        }

    def insert_block(
        self,
        reference_block_id: str,
        position: str,
        text: str
    ) -> Dict[str, str]:
        """
        Insert a new paragraph before or after a reference block.

        The new paragraph inherits formatting from the reference paragraph,
        ensuring consistent styling within the same section.

        Args:
            reference_block_id: ID of the block to insert relative to
            position: 'before' or 'after'
            text: Content for the new paragraph

        Returns:
            dict with 'status' and 'message'

        Raises:
            ValueError: If reference block not found or invalid position
        """
        if reference_block_id not in self._id_to_element:
            raise ValueError(f"Reference block '{reference_block_id}' not found")

        if position not in ("before", "after"):
            raise ValueError("Position must be 'before' or 'after'")

        ref_element = self._id_to_element[reference_block_id]
        ref_xml = ref_element._element

        # Create new paragraph
        new_para = self.document.add_paragraph(text)
        new_para_xml = new_para._p

        # Copy styles from the reference paragraph (the one we're inserting next to)
        # This ensures the new paragraph matches its logical sibling
        if isinstance(ref_element, Paragraph):
            self._copy_paragraph_formatting(ref_element, new_para)

        # Position the new paragraph
        if position == "after":
            ref_xml.addnext(new_para_xml)
        else:
            ref_xml.addprevious(new_para_xml)

        # Rebuild mapping to include new element
        self._rebuild_block_map()

        return {
            "status": "success",
            "message": f"New paragraph inserted {position} '{reference_block_id}'"
        }

    def delete_block(self, block_id: str) -> Dict[str, str]:
        """
        Remove a block from the document.

        Args:
            block_id: ID of the block to delete

        Returns:
            dict with 'status' and 'message'

        Raises:
            ValueError: If block not found
        """
        if block_id not in self._id_to_element:
            raise ValueError(f"Block '{block_id}' not found")

        element = self._id_to_element[block_id]
        element_xml = element._element
        parent = element_xml.getparent()
        parent.remove(element_xml)

        # Clean up mappings
        del self._id_to_element[block_id]
        del self._element_to_id[element]

        return {
            "status": "success",
            "message": f"Block '{block_id}' deleted"
        }

    def save(self, output_path: str) -> None:
        """Save the document to the specified path."""
        self.document.save(output_path)

    def _extract_run_formatting(self, paragraph: Paragraph) -> dict:
        """Extract font formatting from a paragraph's first run."""
        if not paragraph.runs:
            return {}

        run = paragraph.runs[0]
        return {
            "name": run.font.name,
            "size": run.font.size,
            "bold": run.font.bold,
            "italic": run.font.italic,
            "underline": run.font.underline,
            "color": run.font.color.rgb if run.font.color.rgb else None,
        }

    def _apply_run_formatting(self, run, formatting: dict) -> None:
        """Apply saved formatting to a run."""
        if formatting.get("name"):
            run.font.name = formatting["name"]
        if formatting.get("size"):
            run.font.size = formatting["size"]
        if formatting.get("bold") is not None:
            run.font.bold = formatting["bold"]
        if formatting.get("italic") is not None:
            run.font.italic = formatting["italic"]
        # Always explicitly set underline - use False if not set in source
        run.font.underline = formatting.get("underline") or False
        if formatting.get("color"):
            run.font.color.rgb = formatting["color"]

    def _find_next_paragraph(self, ref_element: Paragraph) -> Optional[Paragraph]:
        """Find the next paragraph after the reference element."""
        found_ref = False
        for child in self.document.element.body:
            if isinstance(child, CT_P):
                para = Paragraph(child, self.document)
                if found_ref:
                    return para
                if para._element == ref_element._element:
                    found_ref = True
        return None

    def _copy_paragraph_formatting(
        self,
        source: Paragraph,
        target: Paragraph
    ) -> None:
        """Copy all formatting from source paragraph to target."""
        # Copy paragraph style (includes most formatting)
        target.style = source.style
        target.alignment = source.alignment

        # Copy paragraph format properties (spacing, indentation, line spacing)
        source_fmt = source.paragraph_format
        target_fmt = target.paragraph_format

        if source_fmt.space_before is not None:
            target_fmt.space_before = source_fmt.space_before
        if source_fmt.space_after is not None:
            target_fmt.space_after = source_fmt.space_after
        if source_fmt.line_spacing is not None:
            target_fmt.line_spacing = source_fmt.line_spacing
        if source_fmt.line_spacing_rule is not None:
            target_fmt.line_spacing_rule = source_fmt.line_spacing_rule
        if source_fmt.first_line_indent is not None:
            target_fmt.first_line_indent = source_fmt.first_line_indent
        if source_fmt.left_indent is not None:
            target_fmt.left_indent = source_fmt.left_indent
        if source_fmt.right_indent is not None:
            target_fmt.right_indent = source_fmt.right_indent

        # Copy run-level formatting (font properties)
        if source.runs and target.runs:
            formatting = self._extract_run_formatting(source)
            self._apply_run_formatting(target.runs[0], formatting)

    # -------------------------------------------------------------------------
    # Numbering Support Methods
    # -------------------------------------------------------------------------

    def _get_paragraph_numbering(self, paragraph: Paragraph) -> Optional[Dict[str, Any]]:
        """
        Get numbering properties from a paragraph.

        Returns:
            dict with 'numId', 'ilvl' if numbered, None otherwise
        """
        pPr = paragraph._element.pPr
        if pPr is None:
            return None

        numPr = pPr.numPr
        if numPr is None:
            return None

        numId = numPr.numId.val if numPr.numId is not None else None
        ilvl = numPr.ilvl.val if numPr.ilvl is not None else 0

        if numId is None:
            return None

        return {
            'numId': numId,
            'ilvl': ilvl
        }

    def _apply_numbering_to_paragraph(
        self,
        paragraph: Paragraph,
        numId: int,
        ilvl: int = 0
    ) -> None:
        """
        Apply numbering properties to a paragraph.

        Args:
            paragraph: The paragraph to number
            numId: The numbering ID (from the document's numbering definitions)
            ilvl: The indentation level (0 = top level, 1 = sub-item like "a.", etc.)
        """
        from docx.oxml import parse_xml
        from docx.oxml.ns import nsdecls

        pPr = paragraph._element.get_or_add_pPr()

        # Remove existing numPr if present
        existing_numPr = pPr.numPr
        if existing_numPr is not None:
            pPr.remove(existing_numPr)

        # Create new numPr element
        numPr_xml = f'''
        <w:numPr {nsdecls('w')}>
            <w:ilvl w:val="{ilvl}"/>
            <w:numId w:val="{numId}"/>
        </w:numPr>
        '''
        numPr = parse_xml(numPr_xml)
        pPr.insert(0, numPr)

    def _find_surrounding_numbering(self, reference_block_id: str) -> Optional[Dict[str, Any]]:
        """
        Find numbering from paragraphs surrounding the reference block.

        Checks the reference paragraph first, then adjacent paragraphs.

        Returns:
            dict with 'numId', 'ilvl' if numbering found, None otherwise
        """
        ref_element = self._id_to_element.get(reference_block_id)
        if not isinstance(ref_element, Paragraph):
            return None

        # Check reference paragraph first
        numbering = self._get_paragraph_numbering(ref_element)
        if numbering:
            return numbering

        # Check next paragraph
        next_para = self._find_next_paragraph(ref_element)
        if next_para:
            numbering = self._get_paragraph_numbering(next_para)
            if numbering:
                return numbering

        # Check previous paragraph
        prev_para = self._find_previous_paragraph(ref_element)
        if prev_para:
            numbering = self._get_paragraph_numbering(prev_para)
            if numbering:
                return numbering

        return None

    def _find_previous_paragraph(self, ref_element: Paragraph) -> Optional[Paragraph]:
        """Find the paragraph before the reference element."""
        prev_para = None
        for child in self.document.element.body:
            if isinstance(child, CT_P):
                para = Paragraph(child, self.document)
                if para._element == ref_element._element:
                    return prev_para
                prev_para = para
        return None

    # -------------------------------------------------------------------------
    # Enhanced Insert with Style Control
    # -------------------------------------------------------------------------

    def insert_block_with_style(
        self,
        reference_block_id: str,
        position: str,
        text: str,
        style_from_block_id: Optional[str] = None,
        inherit_numbering: bool = False,
        as_sub_item: bool = False
    ) -> Dict[str, Any]:
        """
        Insert a new paragraph with explicit style and numbering control.

        Args:
            reference_block_id: ID of the block to insert relative to
            position: 'before' or 'after'
            text: Content for the new paragraph
            style_from_block_id: Optional block ID to copy styles from.
                                 If not provided, inherits from next paragraph.
            inherit_numbering: If True and surrounding paragraphs are numbered,
                               apply the same numbering to the new paragraph.
            as_sub_item: If True and inherit_numbering is True, insert as a
                         sub-item (ilvl=1) which typically displays as "a.", "b.", etc.
                         Use this when inserting between numbered items like "3." and "4."
                         to get "3a." numbering.

        Returns:
            dict with 'status', 'message', 'style_source', and optionally 'numbering'
        """
        if reference_block_id not in self._id_to_element:
            raise ValueError(f"Reference block '{reference_block_id}' not found")

        if position not in ("before", "after"):
            raise ValueError("Position must be 'before' or 'after'")

        ref_element = self._id_to_element[reference_block_id]
        ref_xml = ref_element._element

        # Create new paragraph
        new_para = self.document.add_paragraph(text)
        new_para_xml = new_para._p

        # Determine style source
        # When inserting "after", use the reference paragraph's style (the one we're appending to)
        # When inserting "before", use the reference paragraph's style (the one we're prepending to)
        # This ensures the new paragraph matches its logical sibling, not an unrelated next section
        style_source_id = None
        if style_from_block_id:
            if style_from_block_id not in self._id_to_element:
                raise ValueError(f"Style source block '{style_from_block_id}' not found")
            style_source = self._id_to_element[style_from_block_id]
            if isinstance(style_source, Paragraph):
                self._copy_paragraph_formatting(style_source, new_para)
                style_source_id = style_from_block_id
        elif isinstance(ref_element, Paragraph):
            # Always use the reference element's style - this is the paragraph we're inserting next to
            style_source = ref_element
            self._copy_paragraph_formatting(style_source, new_para)
            style_source_id = self._element_to_id.get(style_source, "unknown")

        # Handle numbering
        numbering_info = None
        if inherit_numbering:
            surrounding_numbering = self._find_surrounding_numbering(reference_block_id)
            if surrounding_numbering:
                numId = surrounding_numbering['numId']
                # If as_sub_item, use ilvl=1 (sub-level), otherwise match the surrounding level
                ilvl = 1 if as_sub_item else surrounding_numbering['ilvl']
                self._apply_numbering_to_paragraph(new_para, numId, ilvl)
                numbering_info = {
                    'applied': True,
                    'numId': numId,
                    'ilvl': ilvl,
                    'as_sub_item': as_sub_item
                }

        # Position the new paragraph
        if position == "after":
            ref_xml.addnext(new_para_xml)
        else:
            ref_xml.addprevious(new_para_xml)

        # Rebuild mapping
        self._rebuild_block_map()

        result = {
            "status": "success",
            "message": f"New paragraph inserted {position} '{reference_block_id}'",
            "style_source": style_source_id
        }

        if numbering_info:
            result["numbering"] = numbering_info
            if as_sub_item:
                result["message"] += " (as sub-item with numbering)"
            else:
                result["message"] += " (with numbering)"

        return result

    # -------------------------------------------------------------------------
    # Change Tracking for Preview
    # -------------------------------------------------------------------------

    def get_changes_summary(self, original_path: str) -> Dict[str, Any]:
        """
        Compare current document state with original and return changes summary.

        Args:
            original_path: Path to the original document

        Returns:
            dict with lists of added, modified, and deleted blocks
        """
        original_doc = docx.Document(original_path)

        # Build original content map
        original_content = {}
        idx = 0
        for child in original_doc.element.body:
            if isinstance(child, CT_P):
                para = Paragraph(child, original_doc)
                original_content[f"block_p_{idx}"] = {
                    "type": "paragraph",
                    "text": para.text
                }
                idx += 1
            elif isinstance(child, CT_Tbl):
                table = Table(child, original_doc)
                original_content[f"block_t_{idx}"] = {
                    "type": "table",
                    "data": [[cell.text for cell in row.cells] for row in table.rows]
                }
                idx += 1

        # Get current content
        current_content = self.get_structured_content()

        changes = {
            "modified": [],
            "added": [],
            "deleted": [],
            "unchanged": 0
        }

        # Compare blocks
        current_blocks = {b["id"]: b for b in current_content["blocks"]}

        for block_id, original in original_content.items():
            if block_id in current_blocks:
                current = current_blocks[block_id]
                if original["type"] == "paragraph" and current.get("text") != original["text"]:
                    changes["modified"].append({
                        "block_id": block_id,
                        "type": "paragraph",
                        "original": original["text"][:100] + "..." if len(original["text"]) > 100 else original["text"],
                        "current": current.get("text", "")[:100] + "..." if len(current.get("text", "")) > 100 else current.get("text", "")
                    })
                elif original["type"] == "table":
                    # Simple table comparison
                    orig_flat = str(original["data"])
                    curr_flat = str(current.get("data", []))
                    if orig_flat != curr_flat:
                        changes["modified"].append({
                            "block_id": block_id,
                            "type": "table",
                            "note": "Table content changed"
                        })
                    else:
                        changes["unchanged"] += 1
                else:
                    changes["unchanged"] += 1
            else:
                changes["deleted"].append({
                    "block_id": block_id,
                    "type": original["type"],
                    "content_preview": original.get("text", str(original.get("data", "")))[:50]
                })

        # Check for added blocks
        for block_id, current in current_blocks.items():
            if block_id not in original_content:
                changes["added"].append({
                    "block_id": block_id,
                    "type": current.get("type", "unknown"),
                    "content_preview": current.get("text", str(current.get("data", "")))[:50]
                })

        return changes


# =============================================================================
# TOOLS CLASS - Public interface for the LLM
# =============================================================================


class Tools:
    """
    NDA Redline Tool for Open WebUI.

    Provides tools for reading, editing, and generating redlined versions
    of NDA documents in DOCX format. Designed for use with a legal review
    system prompt that guides the model through compliance checking.

    Workflow:
        1. read_document() - Load and parse the NDA
        2. edit_block() / insert_block() / delete_block() - Make changes
        3. get_edited_document() - Validate changes
        4. generate_redline_document() - Create tracked-changes output
    """

    class Valves(BaseModel):
        """
        Admin-configurable settings for the NDA Redline Tool.
        Configure these in the Open WebUI admin panel.
        """
        BASE_URL: str = Field(
            default="https://dev.42frontiers.com",
            description="Base URL for file download links. Should match your Open WebUI deployment URL."
        )
        DEFAULT_REDLINE_AUTHOR: str = Field(
            default="Legal Review",
            description="Default author name for tracked changes in redline documents."
        )
        UPLOAD_DIR: str = Field(
            default="/app/backend/data/uploads",
            description="Directory where Open WebUI stores uploaded files."
        )
        LOG_LEVEL: Literal["DEBUG", "INFO", "NONE"] = Field(
            default="DEBUG",
            description="Logging level for tool operations. DEBUG shows all tool outputs in a citation, "
                        "INFO shows summaries only, NONE disables logging citation."
        )

    def __init__(self):
        """Initialize the NDA Redline Tool."""
        # Editor cache - maintains document state across multiple tool calls
        self._editors: Dict[str, _WordEditor] = {}

        # Session log buffer for collecting tool outputs (emitted as citation)
        self._session_logs: deque = deque(maxlen=500)

        # Open WebUI tool configuration
        self.valves = self.Valves()
        self.file_handler = False  # We handle files ourselves
        self.citation = False  # We emit custom citations

    # -------------------------------------------------------------------------
    # Internal Helper Methods (not exposed to model)
    # -------------------------------------------------------------------------

    def _clean_field_value(self, value: Any, default: Any = "") -> Any:
        """Extract actual value from Pydantic FieldInfo if needed."""
        if isinstance(value, FieldInfo):
            return value.default if value.default is not None else default
        return value if value is not None else default

    async def _get_docx_path(
        self,
        files: Optional[List[Dict]],
        file_name: Optional[str] = None
    ) -> str:
        """
        Resolve the path to a DOCX file from uploaded files or direct path.

        Args:
            files: List of file info dicts from __files__
            file_name: Optional specific filename or path

        Returns:
            Absolute path to the DOCX file

        Raises:
            ValueError: If no DOCX file found
            FileNotFoundError: If specified file doesn't exist
        """
        log.info(f"_get_docx_path called with files={files}, file_name={file_name}")

        # Check uploaded files first
        if files:
            log.info(f"Processing {len(files)} files")
            for file_info in files:
                log.info(f"File info: {file_info}")
                name = file_info.get("name", "")
                if name.lower().endswith(".docx"):
                    file_id = file_info.get("id", "")
                    if not file_id:
                        raise ValueError(f"File '{name}' is missing its ID")
                    path = f"{self.valves.UPLOAD_DIR}/{file_id}_{name}"
                    log.info(f"Resolved path: {path}")
                    if os.path.exists(path):
                        log.info(f"File exists at {path}")
                        return path
                    else:
                        log.warning(f"File NOT found at {path}")
                        # Try alternative path patterns
                        alt_paths = [
                            f"/app/backend/data/uploads/{file_id}_{name}",
                            f"/app/backend/data/cache/uploads/{file_id}_{name}",
                        ]
                        for alt_path in alt_paths:
                            if os.path.exists(alt_path):
                                log.info(f"File found at alternative path: {alt_path}")
                                return alt_path
                        raise FileNotFoundError(f"DOCX file not found at {path}")

        # Check if file_name is a direct path
        file_name = self._clean_field_value(file_name, "")
        if file_name:
            if os.path.exists(file_name):
                return file_name
            raise FileNotFoundError(f"File not found: {file_name}")

        raise ValueError(
            "No DOCX file found. Please upload a .docx file or provide a valid file path."
        )

    async def _get_editor(
        self,
        files: Optional[List[Dict]],
        file_name: Optional[str] = None
    ) -> _WordEditor:
        """Get or create an editor instance for the specified document."""
        path = await self._get_docx_path(files, file_name)

        if path not in self._editors:
            self._editors[path] = _WordEditor(path)

        return self._editors[path]

    def _get_user_id(self, files: Optional[List[Dict]]) -> str:
        """Extract user ID from file metadata, with fallback."""
        if files and len(files) > 0:
            file_data = files[0].get("file", {})
            if "user_id" in file_data:
                return file_data["user_id"]
        return "system"

    def _strip_uuid_prefix(self, filename: str) -> str:
        """
        Strip UUID prefix from filename if present.

        Open WebUI stores files as: {uuid}_{original_filename}
        UUID format: xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx (36 chars)

        Example:
            Input:  "d41c52b3-16de-4bbc-a64e-000c9d7990da_250425_NDA.docx"
            Output: "250425_NDA.docx"
        """
        # Pattern: UUID (8-4-4-4-12 hex chars) followed by underscore
        uuid_pattern = r'^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}_'
        return re.sub(uuid_pattern, '', filename, flags=re.IGNORECASE)

    async def _upload_to_storage(
        self,
        user_id: str,
        content: bytes,
        filename: str,
        content_type: str
    ) -> Dict[str, Any]:
        """Upload a file to Open WebUI's storage system."""
        file_id = str(uuid.uuid4())
        storage_filename = f"{file_id}_{filename}"

        file_obj = BytesIO(content)
        file_obj.seek(0)

        contents, file_path = Storage.upload_file(file_obj, storage_filename, {})

        file_item = Files.insert_new_file(
            user_id,
            FileForm(
                id=file_id,
                filename=filename,
                path=file_path,
                meta={
                    "name": filename,
                    "content_type": content_type,
                    "size": len(contents),
                    "data": {},
                },
            ),
        )

        return {
            "id": file_id,
            "filename": filename,
            "path": file_path,
            "size": len(content),
            "download_url": f"{self.valves.BASE_URL}/api/v1/files/{file_id}/content",
        }

    async def _emit_status(
        self,
        emitter,
        description: str,
        done: bool = False
    ) -> None:
        """Emit a status event to the UI."""
        if emitter:
            await emitter({
                "type": "status",
                "data": {
                    "description": description,
                    "done": done,
                    "hidden": done,  # Hide completed statuses
                }
            })

    async def _emit_progress(
        self,
        emitter,
        current_step: int,
        total_steps: int,
        step_description: str
    ) -> None:
        """Emit a progress event with step tracking."""
        if emitter:
            progress_pct = int((current_step / total_steps) * 100)
            await emitter({
                "type": "status",
                "data": {
                    "description": f"[{current_step}/{total_steps}] {step_description} ({progress_pct}%)",
                    "done": current_step >= total_steps,
                    "hidden": False,
                }
            })

    def _log_tool_call(self, tool_name: str, result: Any, level: str = "DEBUG") -> None:
        """
        Log a tool call result to the session buffer.

        Args:
            tool_name: Name of the tool being called
            result: The result dict/object to log
            level: Log level (DEBUG or INFO)
        """
        if self.valves.LOG_LEVEL == "NONE":
            return
        if self.valves.LOG_LEVEL == "INFO" and level == "DEBUG":
            return

        timestamp = datetime.datetime.now().strftime("%H:%M:%S")

        # Format the result for logging
        if isinstance(result, dict):
            # For large results, create a summary
            if "blocks" in result and len(result.get("blocks", [])) > 5:
                summary = {
                    "file_path": result.get("file_path"),
                    "summary": result.get("summary"),
                    "blocks_count": len(result.get("blocks", [])),
                    "first_5_blocks": result.get("blocks", [])[:5],
                }
                result_str = str(summary)
            else:
                result_str = str(result)
        else:
            result_str = str(result)

        # Truncate very long results
        if len(result_str) > 2000:
            result_str = result_str[:2000] + "... [truncated]"

        log_entry = f"[{timestamp}] {tool_name}: {result_str}"
        self._session_logs.append(log_entry)

    async def _emit_citation(
        self,
        emitter,
        document: str,
        source_name: str
    ) -> None:
        """
        Emit collected logs as a citation block.

        Args:
            emitter: The event emitter
            document: The log content to include
            source_name: Name for the citation source
        """
        if emitter is None or not document:
            return

        await emitter({
            "type": "citation",
            "data": {
                "document": [document],
                "metadata": [{
                    "date_accessed": datetime.datetime.now().isoformat(),
                    "source": source_name,
                }],
                "source": {"name": source_name},
            }
        })

    def _get_session_logs_text(self) -> str:
        """Get all session logs as formatted text."""
        if not self._session_logs:
            return ""
        return "\n".join(self._session_logs)

    def _clear_session_logs(self) -> None:
        """Clear the session log buffer."""
        self._session_logs.clear()

    # -------------------------------------------------------------------------
    # Public Tool Methods (exposed to model)
    # -------------------------------------------------------------------------

    async def debug_info(
        self,
        __files__=None,
        __event_emitter__=None,
        __user__=None,
    ) -> Dict[str, Any]:
        """
        Debug tool to show what files and context are available.
        Use this to troubleshoot file access issues.
        """
        return {
            "status": "debug",
            "files_received": __files__,
            "files_count": len(__files__) if __files__ else 0,
            "user_info": __user__,
            "openwebui_available": OPENWEBUI_AVAILABLE,
            "valves": {
                "BASE_URL": self.valves.BASE_URL,
                "UPLOAD_DIR": self.valves.UPLOAD_DIR,
            }
        }

    async def read_document(
        self,
        file_name: str = Field(
            default="",
            description="Name of the DOCX file to read. Optional if a file was uploaded."
        ),
        __files__=None,
        __event_emitter__=None,
    ) -> Dict[str, Any]:
        """
        Read a DOCX document and return its structured content with block IDs.

        This is the first step in the NDA review workflow. The document is loaded
        into memory and each paragraph/table is assigned a unique block ID that
        can be used for subsequent edit operations.

        Returns:
            A dictionary containing:
            - file_path: Path to the loaded document
            - blocks: List of content blocks, each with:
                - id: Unique block identifier (e.g., 'block_p_0')
                - type: 'paragraph', 'heading', 'list_item', or 'table'
                - text: Content for paragraphs
                - data: 2D array for tables
                - level: Heading level (1-9) for headings

        Example response:
            {
                "file_path": "/path/to/nda.docx",
                "blocks": [
                    {"id": "block_p_0", "type": "heading", "level": 1, "text": "NON-DISCLOSURE AGREEMENT"},
                    {"id": "block_p_1", "type": "paragraph", "text": "This Agreement is entered into..."},
                    ...
                ]
            }
        """
        log.info(f"read_document called with file_name={file_name}, __files__={__files__}")
        await self._emit_status(__event_emitter__, "Loading document...")

        try:
            file_name = self._clean_field_value(file_name, "")
            editor = await self._get_editor(__files__, file_name)
            result = editor.get_structured_content()

            block_count = len(result.get("blocks", []))
            await self._emit_status(
                __event_emitter__,
                f"Document loaded: {block_count} blocks found",
                done=True
            )

            # Log the result for citation
            self._log_tool_call("read_document", result)

            return result

        except Exception as e:
            log.exception(f"Error in read_document: {e}")
            await self._emit_status(__event_emitter__, f"Error: {str(e)}", done=True)
            return {"error": str(e), "status": "failed"}

    # NOTE: edit_block is intentionally hidden (prefixed with _) to prevent model from using it.
    # The model should use find_replace for ALL text modifications to ensure surgical edits.
    async def _edit_block_internal(
        self,
        block_id: str,
        new_text: str,
        file_name: str = "",
        __files__=None,
        __event_emitter__=None,
    ) -> str:
        """
        Internal method to edit entire paragraph. Not exposed to model.
        Use find_replace instead for all text modifications.
        """
        await self._emit_status(__event_emitter__, f"Editing block {block_id}...")

        try:
            file_name = self._clean_field_value(file_name, "")
            editor = await self._get_editor(__files__, file_name)
            result = editor.edit_block(block_id, new_text)

            await self._emit_status(
                __event_emitter__,
                f"Block {block_id} updated",
                done=True
            )

            self._log_tool_call("edit_block", {"block_id": block_id, "result": result})
            return f"✓ Updated {block_id}"

        except Exception as e:
            await self._emit_status(__event_emitter__, f"Error: {str(e)}", done=True)
            return f"✗ Error: {str(e)}"

    async def find_replace(
        self,
        block_id: str = Field(
            ...,
            description="The ID of the paragraph block to edit (e.g., 'block_p_5')."
        ),
        find_text: str = Field(
            ...,
            description="The EXACT text to find within the paragraph. Must match precisely."
        ),
        replace_text: str = Field(
            ...,
            description="The text to replace it with."
        ),
        replace_all: bool = Field(
            default=False,
            description="If true, replace ALL occurrences. If false (default), only replace the first."
        ),
        file_name: str = Field(
            default="",
            description="Name of the DOCX file. Optional if a file was uploaded."
        ),
        __files__=None,
        __event_emitter__=None,
    ) -> Dict[str, Any]:
        """
        Find and replace specific text within a paragraph. THIS IS THE PREFERRED EDITING METHOD.

        Use this tool for surgical edits - it only changes the exact text you specify,
        preserving all other content in the paragraph. This prevents accidental
        overwrites and ensures minimal changes.

        Examples:
        - Replace "immediately" with "without undue delay"
        - Replace "3 years" with "2 years"
        - Replace "shall ensure" with "shall instruct the third party"

        The find_text must match EXACTLY (case-sensitive). If the text is not found,
        an error is returned showing what text is actually in the block.

        Returns:
            On success: {"status": "success", "message": "...", "occurrences_found": N, "occurrences_replaced": N}
            On failure: {"error": "...", "status": "failed"}
        """
        await self._emit_status(
            __event_emitter__,
            f"Finding and replacing text in {block_id}..."
        )

        try:
            file_name = self._clean_field_value(file_name, "")
            editor = await self._get_editor(__files__, file_name)
            result = editor.find_replace_in_block(block_id, find_text, replace_text, replace_all)

            await self._emit_status(
                __event_emitter__,
                f"Replaced {result['occurrences_replaced']} occurrence(s)",
                done=True
            )

            # Log the result for citation
            self._log_tool_call("find_replace", {"block_id": block_id, "find": find_text, "replace": replace_text, **result})

            # Return minimal string to hide "View Result" dropdown
            return f"✓ Replaced '{find_text}' → '{replace_text}' in {block_id}"

        except Exception as e:
            await self._emit_status(__event_emitter__, f"Error: {str(e)}", done=True)
            # Return error as string
            return f"✗ Error: {str(e)}"

    async def insert_block(
        self,
        relative_to_block_id: str = Field(
            ...,
            description="The ID of the block to insert relative to."
        ),
        position: str = Field(
            ...,
            description="Where to insert: 'before' or 'after' the reference block."
        ),
        text: str = Field(
            ...,
            description="Text content for the new paragraph."
        ),
        style_from_block_id: str = Field(
            default="",
            description="Optional: Block ID to copy formatting from. If empty, inherits from adjacent paragraph."
        ),
        inherit_numbering: bool = Field(
            default=False,
            description="If true and surrounding paragraphs are numbered, apply numbering to the new paragraph."
        ),
        as_sub_item: bool = Field(
            default=False,
            description="If true (and inherit_numbering=true), insert as sub-item (e.g., '3a.' between '3.' and '4.'). "
                        "Use this when inserting the Fund Disclosure Clause into a numbered list."
        ),
        file_name: str = Field(
            default="",
            description="Name of the DOCX file. Optional if a file was uploaded."
        ),
        __files__=None,
        __event_emitter__=None,
    ) -> Dict[str, Any]:
        """
        Insert a new paragraph before or after a reference block.

        The new paragraph can inherit formatting from:
        1. A specific block (if style_from_block_id is provided)
        2. The next paragraph after the insertion point (default behavior)

        Numbering support:
        - Set inherit_numbering=true to continue the surrounding numbered list
        - Set as_sub_item=true to insert as a sub-numbered item (e.g., "3a." between "3." and "4.")

        Use this for:
        - Adding the Fund Disclosure Clause after a confidentiality heading
        - Inserting new required terms with consistent styling
        - Adding clauses into numbered lists with proper sub-numbering

        Returns:
            {"status": "success", "message": "...", "style_source": "block_id"} on success
            {"error": "...", "status": "failed"} on failure
        """
        await self._emit_status(
            __event_emitter__,
            f"Inserting paragraph {position} {relative_to_block_id}..."
        )

        try:
            file_name = self._clean_field_value(file_name, "")
            style_from = self._clean_field_value(style_from_block_id, "")
            inherit_num = self._clean_field_value(inherit_numbering, False)
            as_sub = self._clean_field_value(as_sub_item, False)
            editor = await self._get_editor(__files__, file_name)

            # Use enhanced method with style and numbering control
            result = editor.insert_block_with_style(
                relative_to_block_id,
                position,
                text,
                style_from if style_from else None,
                inherit_numbering=inherit_num,
                as_sub_item=as_sub
            )

            style_info = f" (styled from {result.get('style_source', 'default')})"
            if result.get('numbering'):
                numbering = result['numbering']
                if numbering.get('as_sub_item'):
                    style_info += f", numbered as sub-item (level {numbering['ilvl']})"
                else:
                    style_info += f", numbered (level {numbering['ilvl']})"

            await self._emit_status(__event_emitter__, f"Paragraph inserted{style_info}", done=True)

            # Log the result for citation
            text_preview = text[:100] + "..." if len(text) > 100 else text
            self._log_tool_call("insert_block", {"ref_block": relative_to_block_id, "position": position, "text": text_preview, **result})

            # Return minimal string to hide "View Result" dropdown
            return f"✓ Inserted paragraph {position} {relative_to_block_id}"

        except Exception as e:
            await self._emit_status(__event_emitter__, f"Error: {str(e)}", done=True)
            return f"✗ Error: {str(e)}"

    async def delete_block(
        self,
        block_id: str = Field(
            ...,
            description="The ID of the block to delete."
        ),
        file_name: str = Field(
            default="",
            description="Name of the DOCX file. Optional if a file was uploaded."
        ),
        __files__=None,
        __event_emitter__=None,
    ) -> Dict[str, Any]:
        """
        Delete a block from the document.

        Removes the specified paragraph or table from the document.
        Changes are held in memory until generate_redline_document is called.

        Use sparingly - prefer editing existing blocks when possible.

        Returns:
            {"status": "success", "message": "..."} on success
            {"error": "...", "status": "failed"} on failure
        """
        await self._emit_status(__event_emitter__, f"Deleting block {block_id}...")

        try:
            file_name = self._clean_field_value(file_name, "")
            editor = await self._get_editor(__files__, file_name)
            result = editor.delete_block(block_id)

            await self._emit_status(__event_emitter__, f"Block deleted", done=True)

            # Log the result for citation
            self._log_tool_call("delete_block", {"block_id": block_id, **result})

            # Return minimal string to hide "View Result" dropdown
            return f"✓ Deleted {block_id}"

        except Exception as e:
            await self._emit_status(__event_emitter__, f"Error: {str(e)}", done=True)
            return f"✗ Error: {str(e)}"

    async def get_edited_document(
        self,
        file_name: str = Field(
            default="",
            description="Name of the DOCX file. Optional if a file was uploaded."
        ),
        __files__=None,
        __event_emitter__=None,
    ) -> Dict[str, Any]:
        """
        Get the current state of the edited document for validation.

        Returns the same structured format as read_document, but reflecting
        all edits made so far. Use this to verify changes before generating
        the final redline document.

        Returns:
            Same format as read_document with current (edited) content.
        """
        await self._emit_status(__event_emitter__, "Getting current document state...")

        try:
            file_name = self._clean_field_value(file_name, "")
            editor = await self._get_editor(__files__, file_name)
            result = editor.get_structured_content()

            await self._emit_status(
                __event_emitter__,
                "Document state retrieved",
                done=True
            )

            # Log the result for citation
            self._log_tool_call("get_edited_document", result)

            return result

        except Exception as e:
            await self._emit_status(__event_emitter__, f"Error: {str(e)}", done=True)
            return {"error": str(e), "status": "failed"}

    async def generate_redline_document(
        self,
        output_name: str = Field(
            ...,
            description="Name for the output file (e.g., 'nda_reviewed.docx')."
        ),
        file_name: str = Field(
            default="",
            description="Name of the source DOCX file. Optional if a file was uploaded."
        ),
        redline_author: str = Field(
            default="",
            description="Author name for tracked changes. Leave empty to use default."
        ),
        __files__=None,
        __event_emitter__=None,
    ) -> Dict[str, Any]:
        """
        Generate and upload a redline document showing tracked changes.

        This is the final step in the NDA review workflow. It:
        1. Saves the modified document
        2. Creates a redline (tracked changes) version comparing original to modified
        3. Uploads the redline document to Open WebUI
        4. Returns a download link

        The redline document shows all insertions, deletions, and modifications
        as tracked changes that can be reviewed in Microsoft Word.

        Returns:
            On success:
            {
                "status": "success",
                "message": "Documents generated successfully",
                "redline_download_url": "https://...",
                "redline_filename": "nda_NH_markup.docx"
            }

            On failure:
            {"error": "...", "status": "failed"}
        """
        total_steps = 5

        try:
            # Step 1: Initialize
            await self._emit_progress(__event_emitter__, 1, total_steps, "Preparing document...")

            file_name = self._clean_field_value(file_name, "")
            redline_author = self._clean_field_value(
                redline_author,
                self.valves.DEFAULT_REDLINE_AUTHOR
            )

            editor = await self._get_editor(__files__, file_name)
            original_path = editor.file_path

            # Ensure output has .docx extension
            if not output_name.lower().endswith(".docx"):
                output_name += ".docx"

            # Step 2: Save modified document
            await self._emit_progress(__event_emitter__, 2, total_steps, "Saving modified document...")

            modified_path = os.path.abspath(output_name)
            editor.save(modified_path)

            # Step 3: Generate redline
            await self._emit_progress(__event_emitter__, 3, total_steps, "Creating tracked changes (this may take a moment)...")

            # Get base name and strip UUID prefix added by Open WebUI storage
            raw_base_name = os.path.splitext(os.path.basename(original_path))[0]
            base_name = self._strip_uuid_prefix(raw_base_name)
            redline_filename = f"{base_name}_NH_markup.docx"
            redline_path = os.path.join(os.path.dirname(modified_path), redline_filename)

            engine = XmlPowerToolsEngine()
            redline_bytes = engine.run_redline(redline_author, original_path, modified_path)

            with open(redline_path, "wb") as f:
                f.write(redline_bytes[0])

            # Step 4: Upload to storage
            await self._emit_progress(__event_emitter__, 4, total_steps, "Uploading redline document...")

            # Upload to Open WebUI
            user_id = self._get_user_id(__files__)
            with open(redline_path, "rb") as f:
                redline_content = f.read()

            upload_result = await self._upload_to_storage(
                user_id,
                redline_content,
                redline_filename,
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )

            # Emit file event so it appears in chat
            if __event_emitter__:
                await __event_emitter__({
                    "type": "files",
                    "data": {
                        "files": [{
                            "id": upload_result["id"],
                            "type": "file",
                            "name": redline_filename,
                            "url": upload_result["download_url"],
                        }]
                    }
                })

                # Also emit as citation/source
                await __event_emitter__({
                    "type": "citation",
                    "data": {
                        "source": {
                            "name": redline_filename,
                            "url": upload_result["download_url"],
                        },
                        "document": [f"Redline document with tracked changes by {redline_author}"],
                        "metadata": [{"source": redline_filename}],
                    }
                })

            # Step 5: Complete
            await self._emit_progress(__event_emitter__, 5, total_steps, f"Complete! Redline ready: {redline_filename}")

            # Clean up temp files
            try:
                os.remove(modified_path)
                os.remove(redline_path)
            except OSError:
                pass  # Best effort cleanup

            # Emit session logs as citation (if logging enabled)
            if self.valves.LOG_LEVEL != "NONE":
                logs_text = self._get_session_logs_text()
                if logs_text:
                    await self._emit_citation(
                        __event_emitter__,
                        logs_text,
                        "Tool Execution Log"
                    )
                self._clear_session_logs()

            # Return with download URL so model can provide clickable link to user
            return f"✓ Redline ready: [{redline_filename}]({upload_result['download_url']})"

        except Exception as e:
            log.exception("Error generating redline document")
            await self._emit_status(__event_emitter__, f"Error: {str(e)}", done=True)
            return f"✗ Error: {str(e)}"
