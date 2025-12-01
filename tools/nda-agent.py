import os
import tempfile
import shutil
import base64
import uuid
from typing import Dict, List, Any, Optional, Callable, Awaitable
import traceback
import asyncio
import time
import hashlib

from open_webui.models.files import Files, FileForm
from open_webui.storage.provider import Storage
from io import BytesIO

from pydantic import Field, BaseModel
from pydantic.fields import FieldInfo  # <-- Import FieldInfo

import docx
from docx import Document
from docx.document import Document as DocumentType
from docx.text.paragraph import Paragraph
from docx.table import Table
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from python_redlines.engines import XmlPowerToolsEngine

import logging

# Set up logging
log = logging.getLogger(__name__)

OPENWEBUI_INTERNAL_AVAILABLE = True
BASE_URL = "https://dev.42frontiers.com"


class WordEditor:
    def __init__(self, file_path: str, __event_emitter__=None):
        """Loads a document and creates the initial block-to-ID map."""
        assert (
            file_path.strip()
        ), f"WordEditor.__init__ received empty file_path: '{file_path}'"
        log.debug(f"WordEditor.__init__: file_path='{file_path}'")
        self.file_path = file_path
        self.document = docx.Document(file_path)

        # Core data structures for robust editing
        self.id_to_element = {}  # Maps "block_p_001" -> paragraph object
        self.element_to_id = {}  # Maps paragraph object -> "block_p_001"
        self._next_id = 0
        
        # Style cache for reusing styles across operations
        self._saved_styles = {}

        self._map_document_blocks()

    def _generate_id(self, element_type: str) -> str:
        """Generates a new unique ID for a block."""
        new_id = f"block_{element_type}_{self._next_id}"
        self._next_id += 1
        return new_id

    def _map_document_blocks(self):
        """
        Iterates through the document body to assign a unique ID to each
        paragraph and table. This is the most critical part.
        """
        # We iterate through the body's children to maintain order
        for child in self.document.element.body:
            if isinstance(child, CT_P):
                # It's a paragraph
                para = Paragraph(child, self.document)
                new_id = self._generate_id("p")
                self.id_to_element[new_id] = para
                self.element_to_id[para] = new_id
            elif isinstance(child, CT_Tbl):
                # It's a table
                table = Table(child, self.document)
                new_id = self._generate_id("t")
                self.id_to_element[new_id] = table
                self.element_to_id[table] = new_id

    def read_document(self, event_emitter=None) -> dict:
        """Returns a structured representation of the document."""

        blocks = []
        for element, block_id in self.element_to_id.items():
            block_info = {"id": block_id}
            if isinstance(element, Paragraph):
                block_info["type"] = "paragraph"
                # Heuristic to detect headings and lists
                if (
                    element.style
                    and element.style.name
                    and element.style.name.startswith("Heading")
                ):
                    block_info["type"] = "heading"
                    try:
                        block_info["level"] = int(element.style.name[-1])
                    except ValueError:
                        block_info["level"] = 1  # Default level
                elif (
                    element.style
                    and element.style.name
                    and element.style.name.startswith("List Paragraph")
                ):
                    block_info["type"] = "list_item"

                block_info["text"] = element.text
            elif isinstance(element, Table):
                block_info["type"] = "table"
                # For simplicity, we'll represent table data as text for now
                # A more advanced version would map cell IDs.
                table_data = []
                for row in element.rows:
                    table_data.append([cell.text for cell in row.cells])
                block_info["data"] = table_data

            blocks.append(block_info)

        # Add document structure analysis for debugging
        log.debug("="*60)
        log.debug("DEBUG: DOCUMENT STRUCTURE ANALYSIS")
        log.debug("="*60)
        
        all_elements = []
        for child in self.document.element.body:
            if isinstance(child, CT_P):
                para = Paragraph(child, self.document)
                all_elements.append(para)
            elif isinstance(child, CT_Tbl):
                table = Table(child, self.document)
                all_elements.append(table)
        
        for i, element in enumerate(all_elements):
            if isinstance(element, Paragraph):
                style_name = element.style.name if element.style else "No style"
                text_preview = element.text[:50] + "..." if len(element.text) > 50 else element.text
                block_id = self.element_to_id.get(element, "unknown")
                log.debug(f"Index {i:2d}: Block {block_id:12s} | Style: {style_name:15s} | Text: {text_preview}")
            else:
                block_id = self.element_to_id.get(element, "unknown")
                log.debug(f"Index {i:2d}: Block {block_id:12s} | Type: Table")
        
        log.debug("="*60)
        log.debug("END DOCUMENT STRUCTURE ANALYSIS")
        log.debug("="*60)

        return {"file_path": self.file_path, "blocks": blocks}

    def edit_block(self, block_id: str, new_text: str):
        """Replaces the entire text content of a given block_id while preserving existing formatting."""
        if block_id not in self.id_to_element:
            raise ValueError(f"Block with id '{block_id}' not found.")

        element = self.id_to_element[block_id]

        if isinstance(element, Paragraph):
            # Preserve existing formatting by analyzing the first run's properties
            existing_runs = element.runs
            if existing_runs:
                # Get formatting from the first run to use as template
                first_run = existing_runs[0]
                font_name = first_run.font.name
                font_size = first_run.font.size
                font_bold = first_run.font.bold
                font_italic = first_run.font.italic
                font_underline = first_run.font.underline
                font_color = first_run.font.color.rgb if first_run.font.color.rgb else None
                
                # Also preserve paragraph-level formatting
                paragraph_alignment = element.alignment
                paragraph_style = element.style
            else:
                # No existing runs, use default formatting
                font_name = None
                font_size = None
                font_bold = None
                font_italic = None
                font_underline = None
                font_color = None
                paragraph_alignment = None
                paragraph_style = None
            
            # Clear existing runs
            element.text = ""
            
            # Add new run with preserved formatting
            new_run = element.add_run(new_text)
            
            # Apply preserved formatting to the new run
            if font_name:
                new_run.font.name = font_name
            if font_size:
                new_run.font.size = font_size
            if font_bold is not None:
                new_run.font.bold = font_bold
            if font_italic is not None:
                new_run.font.italic = font_italic
            if font_underline is not None:
                new_run.font.underline = font_underline
            if font_color:
                new_run.font.color.rgb = font_color
            
            # Preserve paragraph-level formatting
            if paragraph_alignment is not None:
                element.alignment = paragraph_alignment
            if paragraph_style is not None:
                element.style = paragraph_style
            
            return {"status": "success", "message": f"Block '{block_id}' updated with preserved formatting."}
        else:
            raise TypeError("`edit_block` only supports paragraph elements.")

    def insert_block(self, relative_to_block_id: str, position: str, new_block: dict):
        """Inserts a new block before or after a reference block, copying styles from the reference block."""
        if relative_to_block_id not in self.id_to_element:
            raise ValueError(f"Reference block '{relative_to_block_id}' not found.")
        if position not in ["before", "after"]:
            raise ValueError("Position must be 'before' or 'after'.")

        relative_element = self.id_to_element[relative_to_block_id]
        relative_element_xml = relative_element._element

        # Create new paragraph element
        if new_block["type"] == "paragraph":
            # Create new paragraph with default styling first
            new_paragraph = self.document.add_paragraph(new_block.get("text", ""))
            new_paragraph_xml = new_paragraph._p
            
            # Copy styles from the next paragraph after the reference block
            if isinstance(relative_element, Paragraph):
                log.debug(f"insert_block - relative_element is a paragraph: '{relative_element.text[:30]}...'")
                ref_style = relative_element.style.name if relative_element.style else "No style"
                log.debug(f"insert_block - relative_element style: '{ref_style}'")
                
                # Find the next paragraph after the reference element
                style_reference = self._find_next_paragraph(relative_element)
                if style_reference:
                    log.debug(f"insert_block - Found next paragraph for styling: '{style_reference.text[:30]}...'")
                    style_ref_style = style_reference.style.name if style_reference.style else "No style"
                    log.debug(f"insert_block - Next paragraph style: '{style_ref_style}'")
                    self._copy_paragraph_styles(style_reference, new_paragraph)
                else:
                    log.debug(f"insert_block - No next paragraph found, using relative_element")
                    # Fallback to reference element if no next paragraph found
                    self._copy_paragraph_styles(relative_element, new_paragraph)
            
            # Move the new paragraph to the correct position
            if position == "after":
                relative_element_xml.addnext(new_paragraph_xml)
            else:  # 'before'
                relative_element_xml.addprevious(new_paragraph_xml)

            # IMPORTANT: We need to rebuild the map to get the new element
            # This is a key challenge. A simple solution is to remap.
            self._remap_document()
            # Determine what styling was used for the message and provide detailed info
            log.debug("="*60)
            log.debug("DEBUG: INSERT BLOCK STYLING ANALYSIS")
            log.debug("="*60)
            
            if isinstance(relative_element, Paragraph):
                ref_style = relative_element.style.name if relative_element.style else "No style"
                ref_text = relative_element.text[:50] + "..." if len(relative_element.text) > 50 else relative_element.text
                ref_block_id = self.element_to_id.get(relative_element, "unknown")
                
                # Debug: Check if the element is in the mapping
                log.debug(f"DEBUG: Element mapping check:")
                log.debug(f"  relative_element object: {relative_element}")
                log.debug(f"  relative_element._element: {relative_element._element}")
                log.debug(f"  element_to_id keys count: {len(self.element_to_id)}")
                log.debug(f"  id_to_element keys count: {len(self.id_to_element)}")
                
                # Try to find the element by comparing _element objects
                found_block_id = None
                for element, block_id in self.element_to_id.items():
                    if hasattr(element, '_element') and element._element == relative_element._element:
                        found_block_id = block_id
                        break
                
                if found_block_id:
                    log.debug(f"  Found matching element by _element: {found_block_id}")
                    ref_block_id = found_block_id
                else:
                    log.debug(f"  No matching element found by _element comparison")
                
                log.debug(f"Reference Element:")
                log.debug(f"  Block ID: {ref_block_id}")
                log.debug(f"  Style: {ref_style}")
                log.debug(f"  Text: {ref_text}")
                
                style_reference = self._find_next_paragraph(relative_element)
                if style_reference:
                    style_ref_style = style_reference.style.name if style_reference.style else "No style"
                    style_ref_text = style_reference.text[:50] + "..." if len(style_reference.text) > 50 else style_reference.text
                    style_ref_block_id = self.element_to_id.get(style_reference, "unknown")
                    
                    # Debug: Check if the style reference element is in the mapping
                    log.debug(f"DEBUG: Style reference mapping check:")
                    log.debug(f"  style_reference object: {style_reference}")
                    log.debug(f"  style_reference._element: {style_reference._element}")
                    
                    # Try to find the style reference element by comparing _element objects
                    found_style_block_id = None
                    for element, block_id in self.element_to_id.items():
                        if hasattr(element, '_element') and element._element == style_reference._element:
                            found_style_block_id = block_id
                            break
                    
                    if found_style_block_id:
                        log.debug(f"  Found matching style reference by _element: {found_style_block_id}")
                        style_ref_block_id = found_style_block_id
                    else:
                        log.debug(f"  No matching style reference found by _element comparison")
                    
                    log.debug(f"Style Source (Next Paragraph):")
                    log.debug(f"  Block ID: {style_ref_block_id}")
                    log.debug(f"  Style: {style_ref_style}")
                    log.debug(f"  Text: {style_ref_text}")
                    
                    style_info = f"Block inserted with styling from next paragraph (block {style_ref_block_id})."
                else:
                    log.debug(f"Style Source (Reference Element - No Next Paragraph Found):")
                    log.debug(f"  Block ID: {ref_block_id}")
                    log.debug(f"  Style: {ref_style}")
                    log.debug(f"  Text: {ref_text}")
                    
                    style_info = f"Block inserted with styling from reference element (block {ref_block_id})."
            else:
                log.debug(f"Reference Element: Not a paragraph (type: {type(relative_element)})")
                log.debug(f"Style Source: Default styling")
                style_info = f"Block inserted with default styling."
            
            log.debug("="*60)
            log.debug("END INSERT BLOCK STYLING ANALYSIS")
            log.debug("="*60)
            
            return {"status": "success", "message": style_info}
        else:
            # Support for tables and headings can be added here
            raise NotImplementedError("Only paragraph insertion is supported for now.")

    def insert_block_with_saved_style(self, relative_to_block_id: str, position: str, new_block: dict, style_name: str):
        """Inserts a new block using a previously saved style instead of copying from reference block."""
        if relative_to_block_id not in self.id_to_element:
            raise ValueError(f"Reference block '{relative_to_block_id}' not found.")
        if position not in ["before", "after"]:
            raise ValueError("Position must be 'before' or 'after'.")
        if style_name not in self._saved_styles:
            raise ValueError(f"No saved styles found for '{style_name}'.")

        relative_element = self.id_to_element[relative_to_block_id]
        relative_element_xml = relative_element._element

        # Create new paragraph element
        if new_block["type"] == "paragraph":
            # Create new paragraph with default styling first
            new_paragraph = self.document.add_paragraph(new_block.get("text", ""))
            new_paragraph_xml = new_paragraph._p
            
            # Apply saved styles instead of copying from reference
            self.apply_saved_styles(new_paragraph, style_name)
            
            # Move the new paragraph to the correct position
            if position == "after":
                relative_element_xml.addnext(new_paragraph_xml)
            else:  # 'before'
                relative_element_xml.addprevious(new_paragraph_xml)

            # IMPORTANT: We need to rebuild the map to get the new element
            self._remap_document()
            
            # Add debugging info for saved style usage
            log.debug("="*60)
            log.debug("DEBUG: INSERT BLOCK WITH SAVED STYLE ANALYSIS")
            log.debug("="*60)
            
            if isinstance(relative_element, Paragraph):
                ref_style = relative_element.style.name if relative_element.style else "No style"
                ref_text = relative_element.text[:50] + "..." if len(relative_element.text) > 50 else relative_element.text
                ref_block_id = self.element_to_id.get(relative_element, "unknown")
                
                log.debug(f"Reference Element:")
                log.debug(f"  Block ID: {ref_block_id}")
                log.debug(f"  Style: {ref_style}")
                log.debug(f"  Text: {ref_text}")
            
            log.debug(f"Style Source: Saved Style '{style_name}'")
            if style_name in self._saved_styles:
                saved_style = self._saved_styles[style_name]
                log.debug(f"  Saved Style Details:")
                log.debug(f"    Paragraph Style: {saved_style.get('paragraph_style', 'None')}")
                log.debug(f"    Font Name: {saved_style.get('font_name', 'None')}")
                log.debug(f"    Font Size: {saved_style.get('font_size', 'None')}")
                log.debug(f"    Bold: {saved_style.get('font_bold', 'None')}")
                log.debug(f"    Italic: {saved_style.get('font_italic', 'None')}")
            else:
                log.debug(f"  WARNING: Saved style '{style_name}' not found!")
            
            log.debug("="*60)
            log.debug("END INSERT BLOCK WITH SAVED STYLE ANALYSIS")
            log.debug("="*60)
            
            return {"status": "success", "message": f"Block inserted with saved style '{style_name}'."}
        else:
            # Support for tables and headings can be added here
            raise NotImplementedError("Only paragraph insertion is supported for now.")

    def _find_next_paragraph(self, reference_element):
        """Find the next paragraph after the reference element for styling."""
        log.debug(f"_find_next_paragraph called for element: {reference_element.text[:50]}...")
        
        # Get all elements in document order
        all_elements = []
        for child in self.document.element.body:
            if isinstance(child, CT_P):
                para = Paragraph(child, self.document)
                all_elements.append(para)
            elif isinstance(child, CT_Tbl):
                table = Table(child, self.document)
                all_elements.append(table)
        
        log.debug(f"Found {len(all_elements)} total elements in document")
        
        # Find the reference element's position
        try:
            ref_index = all_elements.index(reference_element)
            log.debug(f"Reference element found at index {ref_index}")
        except ValueError:
            log.debug("Reference element not found in document elements")
            return None
        
        # Debug: Show what the reference element is
        ref_style = reference_element.style.name if reference_element.style else "No style"
        log.debug(f"Reference element style: '{ref_style}'")
        
        # Look for the next paragraph after the reference element
        log.debug(f"Looking for next paragraph after index {ref_index}")
        for i in range(ref_index + 1, len(all_elements)):
            element = all_elements[i]
            if isinstance(element, Paragraph):
                element_style = element.style.name if element.style else "No style"
                log.debug(f"Found paragraph at index {i}: style='{element_style}', text='{element.text[:30]}...'")
                log.debug(f"Using this paragraph as style reference")
                return element
        
        log.debug("No next paragraph found, returning None")
        return None

    def _copy_paragraph_styles(self, source_paragraph: Paragraph, target_paragraph: Paragraph):
        """Helper method to copy all formatting from source paragraph to target paragraph."""
        # Copy paragraph-level formatting
        target_paragraph.style = source_paragraph.style
        target_paragraph.alignment = source_paragraph.alignment
        
        # Copy run-level formatting from the first run of the source paragraph
        if source_paragraph.runs and target_paragraph.runs:
            source_run = source_paragraph.runs[0]
            target_run = target_paragraph.runs[0]
            
            # Copy font properties
            if source_run.font.name:
                target_run.font.name = source_run.font.name
            if source_run.font.size:
                target_run.font.size = source_run.font.size
            if source_run.font.bold is not None:
                target_run.font.bold = source_run.font.bold
            if source_run.font.italic is not None:
                target_run.font.italic = source_run.font.italic
            if source_run.font.underline is not None:
                target_run.font.underline = source_run.font.underline
            if source_run.font.color.rgb:
                target_run.font.color.rgb = source_run.font.color.rgb

    def save_paragraph_styles(self, block_id: str, style_name: str = "default"):
        """Save the styles from a paragraph for later reuse."""
        if block_id not in self.id_to_element:
            raise ValueError(f"Block with id '{block_id}' not found.")
        
        element = self.id_to_element[block_id]
        if not isinstance(element, Paragraph):
            raise TypeError("Can only save styles from paragraph elements.")
        
        # Extract and save style information
        style_info = {
            "paragraph_style": element.style,
            "alignment": element.alignment,
        }
        
        # Extract run-level formatting from the first run
        if element.runs:
            first_run = element.runs[0]
            style_info.update({
                "font_name": first_run.font.name,
                "font_size": first_run.font.size,
                "font_bold": first_run.font.bold,
                "font_italic": first_run.font.italic,
                "font_underline": first_run.font.underline,
                "font_color": first_run.font.color.rgb if first_run.font.color.rgb else None,
            })
        
        self._saved_styles[style_name] = style_info
        return {"status": "success", "message": f"Styles saved as '{style_name}'."}

    def apply_saved_styles(self, target_paragraph: Paragraph, style_name: str = "default"):
        """Apply previously saved styles to a paragraph."""
        if style_name not in self._saved_styles:
            raise ValueError(f"No saved styles found for '{style_name}'.")
        
        style_info = self._saved_styles[style_name]
        
        # Apply paragraph-level formatting
        if style_info.get("paragraph_style"):
            target_paragraph.style = style_info["paragraph_style"]
        if style_info.get("alignment") is not None:
            target_paragraph.alignment = style_info["alignment"]
        
        # Apply run-level formatting
        if target_paragraph.runs:
            target_run = target_paragraph.runs[0]
            
            if style_info.get("font_name"):
                target_run.font.name = style_info["font_name"]
            if style_info.get("font_size"):
                target_run.font.size = style_info["font_size"]
            if style_info.get("font_bold") is not None:
                target_run.font.bold = style_info["font_bold"]
            if style_info.get("font_italic") is not None:
                target_run.font.italic = style_info["font_italic"]
            if style_info.get("font_underline") is not None:
                target_run.font.underline = style_info["font_underline"]
            if style_info.get("font_color"):
                target_run.font.color.rgb = style_info["font_color"]

    def get_styling_reference_info(self, block_id: str):
        """Get information about what paragraph would be used for styling when inserting after this block."""
        if block_id not in self.id_to_element:
            raise ValueError(f"Block with id '{block_id}' not found.")
        
        element = self.id_to_element[block_id]
        if not isinstance(element, Paragraph):
            return {"error": "Can only get styling reference for paragraph elements."}
        
        style_reference = self._find_next_paragraph(element)
        
        if style_reference:
            # Get the block ID of the style reference
            style_block_id = self.element_to_id.get(style_reference, "unknown")
            return {
                "status": "success",
                "reference_block_id": style_block_id,
                "reference_text": style_reference.text[:50] + "..." if len(style_reference.text) > 50 else style_reference.text,
                "reference_style": style_reference.style.name if style_reference.style else "Normal",
                "message": f"Would use styling from block '{style_block_id}' (next paragraph)"
            }
        else:
            return {
                "status": "success", 
                "reference_block_id": block_id,
                "reference_text": element.text[:50] + "..." if len(element.text) > 50 else element.text,
                "reference_style": element.style.name if element.style else "Normal",
                "message": f"Would use styling from reference block '{block_id}' (no next paragraph found)"
            }

    def debug_document_structure(self):
        """Debug method to show the complete document structure with styles."""
        log.debug("Document Structure Analysis")
        log.debug("=" * 50)
        
        all_elements = []
        for child in self.document.element.body:
            if isinstance(child, CT_P):
                para = Paragraph(child, self.document)
                all_elements.append(para)
            elif isinstance(child, CT_Tbl):
                table = Table(child, self.document)
                all_elements.append(table)
        
        for i, element in enumerate(all_elements):
            if isinstance(element, Paragraph):
                style_name = element.style.name if element.style else "No style"
                text_preview = element.text[:50] + "..." if len(element.text) > 50 else element.text
                block_id = self.element_to_id.get(element, "unknown")
                log.debug(f"Index {i:2d}: Block {block_id:12s} | Style: {style_name:15s} | Text: {text_preview}")
            else:
                block_id = self.element_to_id.get(element, "unknown")
                log.debug(f"Index {i:2d}: Block {block_id:12s} | Type: Table")
        
        log.debug("=" * 50)
        return {"status": "success", "message": "Document structure logged to debug console"}

    def _remap_document(self):
        """Re-runs the mapping logic. Necessary after structural changes."""

        self.id_to_element.clear()
        self.element_to_id.clear()
        self._next_id = 0
        self._map_document_blocks()

    def delete_block(self, block_id: str):
        """Deletes a block from the document."""
        if block_id not in self.id_to_element:
            raise ValueError(f"Block with id '{block_id}' not found.")

        element = self.id_to_element[block_id]
        element_xml = element._element
        parent = element_xml.getparent()
        parent.remove(element_xml)

        # Remove from our maps
        del self.id_to_element[block_id]
        del self.element_to_id[element]

        return {"status": "success", "message": f"Block '{block_id}' deleted."}

    def save_document(self, output_path: str):
        """Saves the document to a new file."""
        self.document.save(output_path)
        return {"status": "success", "message": f"Document saved to {output_path}."}

    def get_block_by_id(self, block_id: str):
        """Returns a block by its ID."""
        if block_id not in self.id_to_element:
            raise ValueError(f"Block with id '{block_id}' not found.")
        return self.id_to_element[block_id]

    def get_full_text(self) -> str:
        """Returns the full text content of the document as a plain string."""
        full_text = ""
        for paragraph in self.document.paragraphs:
            full_text += paragraph.text + "\n"
        return full_text.strip()


class Tools:
    def __init__(self):
        """Initialize the DOCX Editor Tool."""
        # Cache for WordEditor instances to support stateful, multi-step editing sessions.

        log.debug("Tools.__init__ called")

        self.editors: Dict[str, "WordEditor"] = {}
        self.file_handler = False
        self.citation = False

        pass

    async def _get_docx_file_path(
        self,
        __files__: Optional[List[Dict]] = None,
        file_name: Optional[str] = None,
        __event_emitter__: Callable[[dict], Any] = None,
    ) -> str:
        """Internal use only! Helper method to get the path to a DOCX file from uploaded files or local path."""

        if __event_emitter__:
            await __event_emitter__(
                {
                    "type": "status",
                    "data": {
                        "description": f"Retrieving File",
                        "done": False,
                        "hidden": False,
                    },
                }
            )

        assert __files__ is None or isinstance(
            __files__, list
        ), f"__files__ must be None or list, got {type(__files__)}"
        log.debug(f"_get_docx_file_path: __files__={__files__}, file_name='{file_name}'")

        # No specific file_name, look for any .docx file
        for file in __files__:
            if file.get("name", "").lower().endswith(".docx"):
                file_id = file.get("id", "")
                file_name = file.get("name", "")
                if not file_id:
                    raise ValueError(f"File '{file_name}' missing ID in __files__")
                path = f"/app/backend/data/uploads/{file_id}_{file_name}"
                assert (
                    path.strip()
                ), f"Generated empty path: file_id='{file_id}', file_name='{file_name}'"
                log.debug(f"returning path 2: '{path}'")
                if __event_emitter__:
                    await __event_emitter__(
                        {
                            "type": "status",
                            "data": {
                                "description": f"Using {path}",
                                "done": False,
                                "hidden": False,
                            },
                        }
                    )
                return path

        if file_name and file_name.strip():
            if (
                "/" in file_name
                or "\\" in file_name
                or file_name.lower().endswith(".docx")
            ):
                if os.path.exists(file_name):
                    assert file_name.strip(), f"Local file_name is empty: '{file_name}'"
                    log.debug(f"returning path 3: '{file_name}'")
                    return file_name
                raise FileNotFoundError(
                    f"File '{file_name}' not found at the specified path."
                )

        raise ValueError(
            "No file specified. Please upload a .docx file or provide a valid file path."
        )

    async def _get_editor(
        self,
        __files__: Optional[List[Dict]] = None,
        file_name: Optional[str] = "",
        __event_emitter__: Callable[[dict], Any] = None,
    ) -> "WordEditor":
        """Internal use only! Get editor instance."""

        if __event_emitter__:
            await __event_emitter__(
                {
                    "type": "status",
                    "data": {
                        "description": f"Get virtual Word Editor",
                        "done": False,
                        "hidden": False,
                    },
                }
            )

        try:
            source_file_path = await self._get_docx_file_path(__files__, file_name)
            log.debug(f"_get_editor: source_file_path='{source_file_path}'")

            if source_file_path in self.editors:
                log.debug("Using cached editor")
                return self.editors[source_file_path]

            log.debug("Creating new WordEditor...")
            editor = WordEditor(source_file_path)
            self.editors[source_file_path] = editor
            log.debug("WordEditor created successfully")
            return editor

        except Exception as e:
            log.error(f"WordEditor creation failed: {e}")
            raise Exception(f"Document loading failed: {str(e)}")

    async def read_document(
        self,
        file_name: str = Field(
            default="",
            description="Name of the DOCX file to read (optional if file is uploaded). This will load the document for editing.",
        ),
        __files__: Optional[List[Dict]] = None,
        __event_emitter__: Callable[[dict], Any] = None,
    ) -> Dict[str, Any]:
        """
        Read a DOCX document and return its structured content with block IDs. Loads the document into memory for subsequent edits.
        """

        if __event_emitter__:
            await __event_emitter__(
                {
                    "type": "status",
                    "data": {
                        "description": f"Reading Document {file_name}",
                        "done": False,
                        "hidden": False,
                    },
                }
            )

        try:
            log.debug(f"read_document: file_name='{file_name}'")

            # Handle FieldInfo objects
            clean_file_name = file_name
            if isinstance(file_name, FieldInfo):
                clean_file_name = file_name.default

            editor = await self._get_editor(__files__, clean_file_name)
            return editor.read_document()
        except Exception as e:
            log.error(f"Error in read_document: {e}")
            return {"error": traceback.format_exc(), "status": "failed"}

    async def edit_block(
        self,
        block_id: str = Field(
            ..., description="The ID of the block to edit (e.g., 'block_p_001')."
        ),
        new_text: str = Field(..., description="The new text content for the block."),
        file_name: str = Field(
            default="",
            description="Name of the DOCX file to edit (optional if file is uploaded). Required to identify which document to act on.",
        ),
        __files__: Optional[List[Dict]] = None,
        __event_emitter__: Callable[[dict], Any] = None,
    ) -> Dict[str, Any]:
        """
        Edit the text content of a specific block IN MEMORY. The document is NOT saved automatically. Use 'save_document' to persist changes.
        """
        if __event_emitter__:
            await __event_emitter__(
                {
                    "type": "status",
                    "data": {
                        "description": f"Editing Block {block_id}",
                        "done": False,
                        "hidden": False,
                    },
                }
            )

        try:
            # Handle FieldInfo objects
            clean_file_name = file_name
            if isinstance(file_name, FieldInfo):
                clean_file_name = file_name.default

            editor = await self._get_editor(__files__, clean_file_name)
            result = editor.edit_block(block_id, new_text)
            result["message"] += " Change is in memory. Use 'save_document' to persist."
            return result
        except Exception as e:
            return {"error": traceback.format_exc(), "status": "failed"}

    async def insert_block(
        self,
        relative_to_block_id: str = Field(
            ..., description="The ID of the reference block."
        ),
        position: str = Field(
            ...,
            description="Position relative to reference block: 'before' or 'after'.",
        ),
        text: str = Field(..., description="Text content for the new paragraph block."),
        file_name: str = Field(
            default="",
            description="Name of the DOCX file to edit (optional if file is uploaded). Required to identify which document to act on.",
        ),
        use_saved_style: str = Field(
            default="",
            description="Optional: Use a previously saved style instead of copying from reference block. Leave empty to copy from reference block.",
        ),
        __files__: Optional[List[Dict]] = None,
        __event_emitter__: Callable[[dict], Any] = None,
    ) -> Dict[str, Any]:
        """
        Insert a new paragraph block before or after a reference block IN MEMORY. 
        By default, copies styles from the reference block. Optionally use a saved style.
        Use 'save_document' to persist changes.
        """

        if __event_emitter__:
            await __event_emitter__(
                {
                    "type": "status",
                    "data": {
                        "description": f"Inserting Block with content {text}",
                        "done": False,
                        "hidden": False,
                    },
                }
            )
        try:
            # Handle FieldInfo objects
            clean_file_name = file_name
            if isinstance(file_name, FieldInfo):
                clean_file_name = file_name.default

            clean_use_saved_style = use_saved_style
            if isinstance(use_saved_style, FieldInfo):
                clean_use_saved_style = use_saved_style.default

            editor = await self._get_editor(__files__, clean_file_name)
            new_block = {"type": "paragraph", "text": text}
            
            # If a saved style is specified, use it instead of copying from reference
            if clean_use_saved_style and clean_use_saved_style.strip():
                result = editor.insert_block_with_saved_style(
                    relative_to_block_id, position, new_block, clean_use_saved_style
                )
            else:
                result = editor.insert_block(relative_to_block_id, position, new_block)
            
            result["message"] += " Change is in memory. Use 'save_document' to persist."
            return result
        except Exception as e:
            return {"error": traceback.format_exc(), "status": "failed"}

    async def delete_block(
        self,
        block_id: str = Field(..., description="The ID of the block to delete."),
        file_name: str = Field(
            default="",
            description="Name of the DOCX file to edit (optional if file is uploaded). Required to identify which document to act on.",
        ),
        __files__: Optional[List[Dict]] = None,
        __event_emitter__: Callable[[dict], Any] = None,
    ) -> Dict[str, Any]:
        """
        Delete a specific block from the document IN MEMORY. Use 'save_document' to persist changes.
        """

        if __event_emitter__:
            await __event_emitter__(
                {
                    "type": "status",
                    "data": {
                        "description": f"Deleting Block {block_id}",
                        "done": False,
                        "hidden": False,
                    },
                }
            )
        try:
            # Handle FieldInfo objects
            clean_file_name = file_name
            if isinstance(file_name, FieldInfo):
                clean_file_name = file_name.default

            editor = await self._get_editor(__files__, clean_file_name)
            result = editor.delete_block(block_id)
            result["message"] += " Change is in memory. Use 'save_document' to persist."
            return result
        except Exception as e:
            return {"error": traceback.format_exc(), "status": "failed"}

    async def save_paragraph_styles(
        self,
        block_id: str = Field(..., description="The ID of the block to save styles from."),
        style_name: str = Field(
            default="default",
            description="Name to save the styles under for later reuse.",
        ),
        file_name: str = Field(
            default="",
            description="Name of the DOCX file (optional if file is uploaded).",
        ),
        __files__: Optional[List[Dict]] = None,
        __event_emitter__: Callable[[dict], Any] = None,
    ) -> Dict[str, Any]:
        """
        Save the formatting styles from a specific paragraph for later reuse when inserting new paragraphs.
        """
        if __event_emitter__:
            await __event_emitter__(
                {
                    "type": "status",
                    "data": {
                        "description": f"Saving styles from block {block_id}",
                        "done": False,
                        "hidden": False,
                    },
                }
            )

        try:
            # Handle FieldInfo objects
            clean_file_name = file_name
            if isinstance(file_name, FieldInfo):
                clean_file_name = file_name.default

            clean_style_name = style_name
            if isinstance(style_name, FieldInfo):
                clean_style_name = style_name.default

            editor = await self._get_editor(__files__, clean_file_name)
            result = editor.save_paragraph_styles(block_id, clean_style_name)
            return result
        except Exception as e:
            return {"error": traceback.format_exc(), "status": "failed"}

    async def get_styling_reference_info(
        self,
        block_id: str = Field(..., description="The ID of the block to check styling reference for."),
        file_name: str = Field(
            default="",
            description="Name of the DOCX file (optional if file is uploaded).",
        ),
        __files__: Optional[List[Dict]] = None,
        __event_emitter__: Callable[[dict], Any] = None,
    ) -> Dict[str, Any]:
        """
        Get information about what paragraph would be used for styling when inserting after the specified block.
        This helps you understand which paragraph's formatting will be copied to new paragraphs.
        """
        if __event_emitter__:
            await __event_emitter__(
                {
                    "type": "status",
                    "data": {
                        "description": f"Checking styling reference for block {block_id}",
                        "done": False,
                        "hidden": False,
                    },
                }
            )

        try:
            # Handle FieldInfo objects
            clean_file_name = file_name
            if isinstance(file_name, FieldInfo):
                clean_file_name = file_name.default

            editor = await self._get_editor(__files__, clean_file_name)
            result = editor.get_styling_reference_info(block_id)
            return result
        except Exception as e:
            return {"error": traceback.format_exc(), "status": "failed"}

    async def debug_document_structure(
        self,
        file_name: str = Field(
            default="",
            description="Name of the DOCX file (optional if file is uploaded).",
        ),
        __files__: Optional[List[Dict]] = None,
        __event_emitter__: Callable[[dict], Any] = None,
    ) -> Dict[str, Any]:
        """
        Debug method to show the complete document structure with all paragraphs, their styles, and block IDs.
        This helps understand why certain styling references are being used.
        """
        if __event_emitter__:
            await __event_emitter__(
                {
                    "type": "status",
                    "data": {
                        "description": "Analyzing document structure...",
                        "done": False,
                        "hidden": False,
                    },
                }
            )

        try:
            # Handle FieldInfo objects
            clean_file_name = file_name
            if isinstance(file_name, FieldInfo):
                clean_file_name = file_name.default

            editor = await self._get_editor(__files__, clean_file_name)
            result = editor.debug_document_structure()
            return result
        except Exception as e:
            return {"error": traceback.format_exc(), "status": "failed"}

    def get_user_id_from_files(self, __files__: Optional[List[Dict]] = None) -> str:
        """
        Extract user ID from the files parameter.
        Parameters:
        - __files__: The files parameter passed to tool functions

        Returns:
        - str: User ID or default
        """
        if __files__ and len(__files__) > 0:
            first_file = __files__[0]
            if "file" in first_file and "user_id" in first_file["file"]:
                user_id = first_file["file"]["user_id"]
                log.debug(f"get_user_id_from_files: Found user_id: {user_id}")
                return user_id

        # Fallback - use a default user ID (this shouldn't happen in normal usage)
        default_user = "system"
        log.debug(f"get_user_id_from_files: Using default user_id: {default_user}")
        return default_user

    async def upload_file_direct(
        self, user_id: str, file_content: bytes, filename: str, content_type: str
    ) -> Dict[str, Any]:
        """
        Upload a file using OpenWebUI's internal storage system.

        Parameters:
        - user_id (str): User ID for file ownership
        - file_content (bytes): Raw file content
        - filename (str): Name of the file
        - content_type (str): MIME type of the file

        Returns:
        - dict: File information with proper OpenWebUI file ID
        """
        try:

            log.debug(f"upload_file_direct: Uploading {filename} for user {user_id}")
            log.debug(f"upload_file_direct: File size: {len(file_content)} bytes")
            log.debug(f"upload_file_direct: Content type: {content_type}")

            # Generate ID and prepare file
            file_id = str(uuid.uuid4())
            storage_filename = f"{file_id}_{filename}"

            log.debug(f"upload_file_direct: Generated file_id: {file_id}")
            log.debug(f"upload_file_direct: Storage filename: {storage_filename}")

            # Upload to storage
            file_obj = BytesIO(file_content)
            file_obj.seek(0)
            contents, file_path = Storage.upload_file(file_obj, storage_filename, {})

            log.debug(f"upload_file_direct: Storage upload successful, path: {file_path}")

            # Create database record
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

            log.debug(f"upload_file_direct: Database record created successfully")

            # Return in consistent format
            result = {
                "id": file_id,
                "filename": filename,
                "path": file_path,
                "size": len(file_content),
                "content_type": content_type,
                "status": "uploaded_successfully",
                "message": f"File {filename} uploaded successfully to OpenWebUI",
                "openwebui_registered": True,
                "download_url": f"{BASE_URL}/api/v1/files/{file_id}/content",
                "access": file_item.access_control,
                "ididid": file_item.id,
            }

            log.debug(f"upload_file_direct: Success! File available at /api/v1/files/{file_id}")
            return result

        except Exception as e:
            error_msg = f"Direct upload failed for {filename}: {str(e)}"
            log.error(f"upload_file_direct: {error_msg}")
            logging.error(error_msg)
            return {"error": error_msg, "status": "failed"}

    def upload_file(
        self,
        file_path: str,
        user_id: str = None,
        __files__: Optional[List[Dict]] = None,
    ) -> Dict[str, Any]:
        """
        Main upload function that reads a file and uploads it using OpenWebUI's internal system.

        Parameters:
        - file_path (str): Path to the file to upload
        - user_id (str): User ID (optional, will be extracted from __files__ if not provided)
        - __files__ (Optional[List[Dict]]): Files context for user ID extraction

        Returns:
        - dict: Upload result
        """
        try:
            log.debug(f"upload_file: Starting upload of {file_path}")

            # Get user ID if not provided
            if not user_id:
                user_id = self.get_user_id_from_files(__files__)

            # Read file content
            with open(file_path, "rb") as f:
                file_content = f.read()

            filename = os.path.basename(file_path)

            # Determine content type
            content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            if filename.lower().endswith(".docx"):
                content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            elif filename.lower().endswith(".pdf"):
                content_type = "application/pdf"
            elif filename.lower().endswith(".txt"):
                content_type = "text/plain"

            # Upload using OpenWebUI's internal system
            return self.upload_file_direct(
                user_id, file_content, filename, content_type
            )

        except FileNotFoundError:
            error_msg = f"File not found: {file_path}"
            log.error(f"upload_file: {error_msg}")
            return {"error": error_msg, "status": "failed"}
        except Exception as e:
            error_msg = f"Upload failed for {file_path}: {str(e)}"
            log.error(f"upload_file: {error_msg}")
            return {"error": error_msg, "status": "failed"}

    def _place_file_in_openwebui_system(self, file_path: str) -> Dict[str, Any]:
        """
        Place the file directly in OpenWebUI's file system using the same pattern as uploaded files.
        """
        try:

            # Generate a unique file ID like OpenWebUI does
            file_id = str(uuid.uuid4())
            filename = os.path.basename(file_path)

            # Calculate file hash
            with open(file_path, "rb") as f:
                file_content = f.read()
                file_hash = hashlib.sha256(file_content).hexdigest()

            # Determine OpenWebUI's upload directory from the logs we've seen
            # The pattern from logs is: /app/backend/data/uploads/{file_id}_{filename}
            upload_dir = "/app/backend/data/uploads"
            target_filename = f"{file_id}_{filename}"
            target_path = os.path.join(upload_dir, target_filename)

            log.debug(f"_place_file_in_openwebui_system: Copying to {target_path}")

            # Ensure upload directory exists
            os.makedirs(upload_dir, exist_ok=True)

            # Copy file to OpenWebUI's upload location
            shutil.copy2(file_path, target_path)

            log.debug(f"_place_file_in_openwebui_system: File copied successfully")
            log.debug(f"_place_file_in_openwebui_system: Target file size: {os.path.getsize(target_path)} bytes")

            # Return metadata in the same format as a successful upload
            result = {
                "id": file_id,
                "filename": filename,
                "path": target_path,
                "size": os.path.getsize(target_path),
                "hash": file_hash,
                "status": "uploaded_internally",
                "message": f"File placed in OpenWebUI system at {target_path}",
                "internal_upload": True,
            }

            log.debug(f"_place_file_in_openwebui_system: Success! File available at {target_path}")
            return result

        except Exception as e:
            error_msg = f"Failed to place file in OpenWebUI system: {str(e)}"
            log.error(f"_place_file_in_openwebui_system: {error_msg}")
            return {"error": error_msg}

    def get_file_content_from_path(self, file_path: str) -> bytes:
        """Read file content as bytes from a local file path."""
        with open(file_path, "rb") as file:
            return file.read()

    async def generate_redline_document(
        self,
        __event_emitter__: Callable[[dict[str, Any]], Awaitable[None]] | None,
        output_name: str = Field(
            ..., description="Name for the output file, e.g., 'edited_report.docx'."
        ),
        file_name: str = Field(
            default="",
            description="Name of the source DOCX file being edited (optional if file is uploaded). Specifies which in-memory document to save.",
        ),
        redline_author: str = Field(
            default="Editor",
            description="The author name to use for the tracked changes in the redline document. Only used if 'create_redline' is True.",
        ),
        __files__: Optional[List[Dict]] = None,
    ) -> Dict[str, Any]:
        """
        Saves the current in-memory state of a document and optionally uploads it to the server.
        Can optionally create a 'redline' version with tracked changes against the original.
        """

        try:
            log.debug(f"save_document: Starting...")
            if __event_emitter__:
                await __event_emitter__(
                    {
                        "type": "status",
                        "data": {
                            "description": f"Generating Document {output_name}...",
                            "done": False,
                        },
                    }
                )
            # Handle FieldInfo objects
            clean_file_name = file_name
            if isinstance(file_name, FieldInfo):
                clean_file_name = file_name.default

            clean_redline_author = redline_author
            if isinstance(redline_author, FieldInfo):
                clean_redline_author = redline_author.default

            log.debug(f"Getting editor for file: {clean_file_name}")
            editor = await self._get_editor(__files__, clean_file_name)
            log.debug(f"Got editor successfully")

            original_path = editor.file_path
            log.debug(f"Original path: {original_path}")

            content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            if clean_file_name.lower().endswith(".docx"):
                content_type = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            elif clean_file_name.lower().endswith(".pdf"):
                content_type = "application/pdf"
            elif clean_file_name.lower().endswith(".txt"):
                content_type = "text/plain"

            # if not output_name.lower().endswith(".docx"):
            #     output_name += ".docx"

            # Create temporary file path
            modified_path = os.path.abspath(output_name)
            log.debug(f"Saving document to: {modified_path}")

            # Save the modified document to temporary location
            log.debug(f"About to call editor.save_document...")
            editor.save_document(modified_path)
            log.debug(f"Document saved successfully to {modified_path}")
            log.debug(f"Saved file size: {os.path.getsize(modified_path)} bytes")

            result = {
                "status": "success",
                "message": f"Document saved successfully as {output_name}.",
                "local_path": modified_path,
            }

            log.debug(f"Creating redline document...")
            if __event_emitter__:
                await __event_emitter__(
                    {
                        "type": "status",
                        "data": {
                            "description": "Generating redline comparison document...",
                            "done": False,
                        },
                    }
                )

            # Derive redline filename from the output name
            name_part, ext = os.path.splitext(os.path.basename(modified_path))
            redline_filename = f"{name_part}_NH_markup{ext}"
            redline_output_path = os.path.join(
                os.path.dirname(modified_path), redline_filename
            )

            log.debug(f"About to create redline with XmlPowerToolsEngine...")
            wrapper = XmlPowerToolsEngine()
            # Use the cleaned author_name variable
            redline_bytes_tuple = wrapper.run_redline(
                clean_redline_author, original_path, modified_path
            )
            log.debug(f"Redline created successfully")

            # Save redline document to temporary location
            log.debug(f"Saving redline to {redline_output_path}...")
            with open(redline_output_path, "wb") as f:
                f.write(redline_bytes_tuple[0])
            log.debug(f"Redline saved. Size: {os.path.getsize(redline_output_path)} bytes")

            result["redline_local_path"] = redline_output_path

            if __event_emitter__:
                await __event_emitter__(
                    {
                        "type": "status",
                        "data": {
                            "description": f"Uploading redline document {redline_filename}...",
                            "done": False,
                        },
                    }
                )

            # Upload the redline document
            log.debug(f"About to upload redline...")

            user_id = self.get_user_id_from_files(__files__)
            # self, user_id: str, file_content: bytes, filename: str, content_type: str

            redline_file_bytes = self.get_file_content_from_path(redline_output_path)

            redline_upload_result = await self.upload_file_direct(
                user_id, redline_file_bytes, redline_filename, content_type
            )
            log.debug(f"Redline upload result: {redline_upload_result}")

            # file = Files.get_file_by_id(redline_upload_result["id"])

            file_id = redline_upload_result.get("id")
            file_object = Files.get_file_by_id(file_id)

            contents = file_object.data.get("content", "")

            if file_object:
                context = {
                    "documents": [[file_object.data.get("content", "")]],
                    "metadatas": [
                        [
                            {
                                "file_id": redline_upload_result.get("id"),
                                "name": file_object.filename,
                                "source": file_object.filename,
                            }
                        ]
                    ],
                }

                event_redline = {
                    "type": "files",
                    "data": {
                        "files": [
                            {
                                "name": file_object.filename,
                                "url": f"{BASE_URL}/api/v1/files/{file_id}/content",
                            }
                        ]
                    },
                }

                if __event_emitter__:
                    await __event_emitter__(event_redline)

            if redline_upload_result.get("error"):
                result["redline_upload_error"] = redline_upload_result.get("error")
                result["message"] += f" Redline created locally (upload failed)."
            else:
                result["redline_upload_response"] = redline_upload_result
                result[
                    "message"
                ] += f" Redline document uploaded successfully as {redline_filename}."

                log.debug(f"Redline OpenWebUI download URL: {redline_upload_result}")

            if __event_emitter__:

                # Step 3: Emit files event to frontend UI
                await __event_emitter__(
                    {
                        "type": "files",
                        "data": {
                            "files": [
                                {
                                    "id": file_id,
                                    "type": "file",
                                    "name": redline_filename,
                                    "url": f"{BASE_URL}/api/v1/files/{file_id}/content",
                                }
                            ]
                        },
                    }
                )

                await __event_emitter__(
                    {
                        "type": "source",
                        "data": {
                            "source": {
                                "name": redline_filename,
                                "url": f"{BASE_URL}/api/v1/files/{file_id}/content",
                            },
                            "document": [contents],
                            "metadata": [{"source": redline_filename}],
                        },
                    }
                )

                await __event_emitter__(
                    {
                        "type": "status",
                        "data": {
                            "description": "Redline document processed.",
                            "done": True,
                            "hidden": False,
                        },
                    }
                )

            return result

        except Exception as e:
            log.error(f"Error in save_document: {e}")
            if __event_emitter__:
                await __event_emitter__(
                    {
                        "type": "status",
                        "data": {
                            "description": f"An error occurred during save: {e}",
                            "done": True,
                        },
                    }
                )
            return {"error": traceback.format_exc(), "status": "failed"}
