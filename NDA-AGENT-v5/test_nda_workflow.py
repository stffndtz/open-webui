#!/usr/bin/env python3
"""
Local test harness for nda-agent-3.0.py

Simulates how Open WebUI would call the NDA tool, allowing you to:
1. Load a document and see what the AI sees
2. Make tool calls (find_replace_in_cell, insert_block, etc.)
3. Generate the redline document

Usage:
    source .venv/bin/activate
    python test_nda_workflow.py /path/to/document.docx

Interactive commands:
    read        - Read/re-read the document
    cells       - Show all cells with their IDs and content
    search X    - Search for text X in all cells
    replace     - Execute a find_replace_in_cell
    insert      - Execute an insert_block
    redline     - Generate the redline document
    prompt      - Show the system prompt
    quit        - Exit
"""

import sys
import os
import json
import re
import shutil
import subprocess
from typing import Dict, List, Any, Optional

# Add current directory to path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P


class LocalWordEditor:
    """
    Simplified local version of _WordEditor for testing.
    Replicates the key functionality without Open WebUI dependencies.
    """

    GERMAN_INDICATORS = ['der', 'die', 'das', 'und', 'ist', 'wird', 'sich', 'für', 'auf', 'mit', 'des', 'den']
    ENGLISH_INDICATORS = ['the', 'and', 'is', 'will', 'for', 'shall', 'with', 'any', 'this', 'that', 'of', 'to']

    def __init__(self, file_path: str):
        self.file_path = file_path
        self.document = Document(file_path)
        self._tables: List[Table] = []
        self._table_metadata: Dict[int, dict] = {}
        self._id_to_element: Dict[str, Any] = {}
        self._element_to_id: Dict[Any, str] = {}
        self._id_to_cell: Dict[str, tuple] = {}
        self._cell_to_id: Dict[tuple, str] = {}
        self._next_id = 0

        self._build_block_map()
        self._classify_tables()
        self._map_table_cells()

    def _build_block_map(self):
        """Build mapping of block IDs to document elements."""
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
                self._tables.append(table)

    def _classify_tables(self):
        """Classify tables as content or metadata."""
        for idx, table in enumerate(self._tables):
            rows = len(table.rows)
            cols = len(table.columns) if table.rows else 0

            metadata = {
                "type": "metadata",
                "rows": rows,
                "columns": cols,
                "language_columns": {},
            }

            if rows > 10 and cols >= 2:
                col_text = {}
                for row in table.rows[:15]:
                    for ci, cell in enumerate(row.cells):
                        col_text[ci] = col_text.get(ci, 0) + len(cell.text)

                text_cols = [i for i, length in col_text.items() if length > 500]

                if len(text_cols) >= 2:
                    metadata["type"] = "content"
                    for col in text_cols:
                        text_sample = ""
                        for row in table.rows[:10]:
                            if col < len(row.cells):
                                text_sample += " " + row.cells[col].text.lower() + " "

                        de_score = sum(1 for w in self.GERMAN_INDICATORS if f" {w} " in text_sample)
                        en_score = sum(1 for w in self.ENGLISH_INDICATORS if f" {w} " in text_sample)
                        metadata["language_columns"][col] = "de" if de_score > en_score else "en"

            self._table_metadata[idx] = metadata

    def _map_table_cells(self):
        """Map cells in content tables."""
        for table_idx, metadata in self._table_metadata.items():
            if metadata["type"] != "content":
                continue

            table = self._tables[table_idx]
            lang_cols = metadata.get("language_columns", {})

            for row_idx, row in enumerate(table.rows):
                seen_in_row = set()
                for col_idx, cell in enumerate(row.cells):
                    tc_id = id(cell._tc)
                    if tc_id in seen_in_row:
                        continue
                    seen_in_row.add(tc_id)

                    if col_idx not in lang_cols:
                        continue

                    cell_id = f"cell_{table_idx}_{row_idx}_{col_idx}"
                    self._id_to_cell[cell_id] = (table, row_idx, col_idx)
                    self._cell_to_id[(table_idx, row_idx, col_idx)] = cell_id

    def read_document(self) -> Dict[str, Any]:
        """Read and return document structure."""
        blocks = []
        for block_id, element in self._id_to_element.items():
            if isinstance(element, Paragraph):
                blocks.append({
                    "id": block_id,
                    "type": "paragraph",
                    "text": element.text,
                })
            elif isinstance(element, Table):
                table_idx = self._tables.index(element)
                meta = self._table_metadata.get(table_idx, {})
                blocks.append({
                    "id": block_id,
                    "type": "table",
                    "metadata": {
                        "rows": len(element.rows),
                        "columns": len(element.columns) if element.rows else 0,
                        "classification": meta,
                    }
                })

        cells = []
        for cell_id, (table, row_idx, col_idx) in self._id_to_cell.items():
            cell = table.rows[row_idx].cells[col_idx]
            table_idx = self._tables.index(table)
            lang = self._table_metadata[table_idx]["language_columns"].get(col_idx)
            cells.append({
                "id": cell_id,
                "table": table_idx,
                "row": row_idx,
                "column": col_idx,
                "language": lang,
                "text": cell.text,
            })

        content_tables = [i for i, m in self._table_metadata.items() if m["type"] == "content"]
        detected_languages = set()
        for idx in content_tables:
            detected_languages.update(self._table_metadata[idx].get("language_columns", {}).values())

        is_bilingual = "de" in detected_languages and "en" in detected_languages
        document_type = "bilingual_table" if (content_tables and is_bilingual) else (
            "table_based" if content_tables else "paragraph_based"
        )

        return {
            "file_path": self.file_path,
            "summary": {
                "document_type": document_type,
                "detected_languages": list(detected_languages),
                "is_bilingual": is_bilingual,
                "total_blocks": len(blocks),
                "paragraphs": len([b for b in blocks if b["type"] == "paragraph"]),
                "tables": len(self._tables),
                "content_tables": content_tables,
                "total_cells": len(cells),
            },
            "blocks": blocks,
            "cells": cells,
        }

    def find_replace_in_cell(self, cell_id: str, find_text: str, replace_text: str) -> Dict[str, str]:
        """Find and replace text in a cell."""
        if cell_id not in self._id_to_cell:
            return {"status": "error", "message": f"Cell '{cell_id}' not found"}

        table, row_idx, col_idx = self._id_to_cell[cell_id]
        cell = table.rows[row_idx].cells[col_idx]

        replaced = False
        for para in cell.paragraphs:
            if find_text in para.text:
                full_text = para.text.replace(find_text, replace_text)
                if para.runs:
                    first_run = para.runs[0]
                    saved_format = {
                        "font_name": first_run.font.name,
                        "font_size": first_run.font.size,
                        "bold": first_run.font.bold,
                        "italic": first_run.font.italic,
                    }
                    para.clear()
                    new_run = para.add_run(full_text)
                    if saved_format["font_name"]:
                        new_run.font.name = saved_format["font_name"]
                    if saved_format["font_size"]:
                        new_run.font.size = saved_format["font_size"]
                    if saved_format["bold"] is not None:
                        new_run.font.bold = saved_format["bold"]
                    if saved_format["italic"] is not None:
                        new_run.font.italic = saved_format["italic"]
                    replaced = True

        if not replaced:
            return {"status": "warning", "message": f"Text '{find_text}' not found in cell '{cell_id}'"}

        return {"status": "success", "message": f"Replaced '{find_text}' with '{replace_text}' in {cell_id}"}

    def find_replace(self, block_id: str, find_text: str, replace_text: str) -> Dict[str, str]:
        """Find and replace text in a paragraph block."""
        if block_id not in self._id_to_element:
            return {"status": "error", "message": f"Block '{block_id}' not found"}

        element = self._id_to_element[block_id]
        if not isinstance(element, Paragraph):
            return {"status": "error", "message": f"Block '{block_id}' is not a paragraph"}

        if find_text not in element.text:
            return {"status": "warning", "message": f"Text '{find_text}' not found in block '{block_id}'"}

        full_text = element.text.replace(find_text, replace_text)
        if element.runs:
            first_run = element.runs[0]
            saved_format = {
                "font_name": first_run.font.name,
                "font_size": first_run.font.size,
                "bold": first_run.font.bold,
                "italic": first_run.font.italic,
            }
            element.clear()
            new_run = element.add_run(full_text)
            if saved_format["font_name"]:
                new_run.font.name = saved_format["font_name"]
            if saved_format["font_size"]:
                new_run.font.size = saved_format["font_size"]

        return {"status": "success", "message": f"Replaced '{find_text}' with '{replace_text}' in {block_id}"}

    def insert_block(self, reference_block_id: str, position: str, text: str) -> Dict[str, str]:
        """Insert a new paragraph before or after a reference block."""
        # Handle cell IDs - convert to table block ID
        if reference_block_id.startswith("cell_"):
            parts = reference_block_id.split("_")
            if len(parts) >= 4:
                table_idx = int(parts[1])
                if table_idx < len(self._tables):
                    table = self._tables[table_idx]
                    table_block_id = self._element_to_id.get(table)
                    if table_block_id:
                        reference_block_id = table_block_id

        if reference_block_id not in self._id_to_element:
            return {"status": "error", "message": f"Reference block '{reference_block_id}' not found"}

        ref_element = self._id_to_element[reference_block_id]
        ref_xml = ref_element._element if hasattr(ref_element, '_element') else ref_element._tbl

        new_para = self.document.add_paragraph(text)
        new_para_xml = new_para._p

        if position == "after":
            ref_xml.addnext(new_para_xml)
        else:
            ref_xml.addprevious(new_para_xml)

        return {"status": "success", "message": f"Inserted paragraph {position} {reference_block_id}"}

    def save(self, output_path: str):
        """Save the document."""
        self.document.save(output_path)
        return {"status": "success", "message": f"Saved to {output_path}"}


def print_header(text: str):
    """Print a formatted header."""
    print(f"\n{'='*70}")
    print(f"  {text}")
    print('='*70)


def print_summary(doc_result: dict):
    """Print document summary."""
    s = doc_result['summary']
    print(f"\nDocument Type:      {s['document_type']}")
    print(f"Detected Languages: {s['detected_languages']}")
    print(f"Is Bilingual:       {s['is_bilingual']}")
    print(f"Total Blocks:       {s['total_blocks']}")
    print(f"Paragraphs:         {s['paragraphs']}")
    print(f"Tables:             {s['tables']}")
    print(f"Content Tables:     {s['content_tables']}")
    print(f"Total Cells:        {s['total_cells']}")


def print_cells(doc_result: dict, filter_text: str = None):
    """Print cells, optionally filtered by text."""
    cells = doc_result.get('cells', [])
    if not cells:
        print("No cells found (not a table-based document)")
        return

    # Group by row
    by_row = {}
    for cell in cells:
        key = (cell['table'], cell['row'])
        if key not in by_row:
            by_row[key] = []
        by_row[key].append(cell)

    count = 0
    for (table_idx, row_idx), row_cells in sorted(by_row.items()):
        # Filter if needed
        if filter_text:
            if not any(filter_text.lower() in c['text'].lower() for c in row_cells):
                continue

        print(f"\n--- Table {table_idx}, Row {row_idx} ---")
        for cell in sorted(row_cells, key=lambda x: x['column']):
            text_preview = cell['text'][:80].replace('\n', ' ').strip()
            if len(cell['text']) > 80:
                text_preview += "..."
            highlight = " <-- MATCH" if filter_text and filter_text.lower() in cell['text'].lower() else ""
            print(f"  [{cell['language']}] {cell['id']}: {text_preview}{highlight}")
        count += 1

        if count >= 20 and not filter_text:
            print(f"\n... and {len(by_row) - count} more rows. Use 'search <text>' to filter.")
            break


def load_prompt():
    """Load the system prompt."""
    prompt_path = os.path.join(os.path.dirname(__file__), 'nda-agent-prompt-3.0.md')
    if os.path.exists(prompt_path):
        with open(prompt_path, 'r') as f:
            return f.read()
    return "Prompt file not found"


def interactive_mode(editor: LocalWordEditor, doc_result: dict):
    """Run interactive command loop."""
    print("\nEntering interactive mode. Type 'help' for commands.")

    while True:
        try:
            cmd = input("\n> ").strip()
        except (EOFError, KeyboardInterrupt):
            print("\nExiting...")
            break

        if not cmd:
            continue

        parts = cmd.split(maxsplit=1)
        command = parts[0].lower()
        args = parts[1] if len(parts) > 1 else ""

        if command in ('quit', 'exit', 'q'):
            print("Exiting...")
            break

        elif command == 'help':
            print("""
Commands:
  read              - Re-read the document and show summary
  cells             - Show all cells (first 20 rows)
  search <text>     - Search for text in cells
  compliance        - Check compliance status for "immediately" requirement
  replace           - Execute find_replace_in_cell (interactive)
  replace_para      - Execute find_replace for paragraphs (interactive)
  insert            - Execute insert_block (interactive)
  save <path>       - Save edited document
  redline           - Generate redline (requires original + edited)
  prompt            - Show the system prompt
  blocks            - Show paragraph blocks
  help              - Show this help
  quit              - Exit
""")

        elif command == 'read':
            doc_result = editor.read_document()
            print_summary(doc_result)

        elif command == 'cells':
            print_cells(doc_result)

        elif command == 'search':
            if not args:
                print("Usage: search <text>")
            else:
                print(f"Searching for '{args}'...")
                print_cells(doc_result, args)

        elif command == 'compliance':
            # Special compliance check for "immediately" requirement
            print("\n=== Compliance Check: Requirement #8 (immediately/unverzüglich) ===\n")
            cells = doc_result.get('cells', [])
            by_row = {}
            for cell in cells:
                text_lower = cell['text'].lower()
                has_issue = 'immediately' in text_lower or 'unverzüglich' in text_lower
                has_compliant = 'without undue delay' in text_lower or 'ohne schuldhaftes zögern' in text_lower
                if has_issue or has_compliant:
                    key = (cell['table'], cell['row'])
                    if key not in by_row:
                        by_row[key] = []
                    status = "❌ NEEDS EDIT" if has_issue and not has_compliant else "✅ COMPLIANT"
                    by_row[key].append((cell, status))

            for (table, row), cells_info in sorted(by_row.items()):
                print(f"Row {row}:")
                for cell, status in cells_info:
                    text = cell['text'][:50].replace('\n', ' ')
                    print(f"  [{cell['language']}] {cell['id']}: {status}")
                    print(f"      {text}...")
                print()

        elif command == 'replace':
            print("Find and replace in cell:")
            cell_id = input("  Cell ID (e.g., cell_3_38_0): ").strip()
            find_text = input("  Find text: ").strip()
            replace_text = input("  Replace with: ").strip()

            if cell_id and find_text and replace_text:
                result = editor.find_replace_in_cell(cell_id, find_text, replace_text)
                print(f"  Result: {result['status']} - {result['message']}")
                # Update doc_result
                doc_result = editor.read_document()
            else:
                print("  Cancelled (all fields required)")

        elif command == 'replace_para':
            print("Find and replace in paragraph:")
            block_id = input("  Block ID (e.g., block_p_15): ").strip()
            find_text = input("  Find text: ").strip()
            replace_text = input("  Replace with: ").strip()

            if block_id and find_text and replace_text:
                result = editor.find_replace(block_id, find_text, replace_text)
                print(f"  Result: {result['status']} - {result['message']}")
                doc_result = editor.read_document()
            else:
                print("  Cancelled (all fields required)")

        elif command == 'insert':
            print("Insert block:")
            ref_id = input("  Reference block/cell ID: ").strip()
            position = input("  Position (before/after): ").strip().lower()
            print("  Enter text (end with empty line):")
            lines = []
            while True:
                line = input("  ")
                if not line:
                    break
                lines.append(line)
            text = "\n".join(lines)

            if ref_id and position in ('before', 'after') and text:
                result = editor.insert_block(ref_id, position, text)
                print(f"  Result: {result['status']} - {result['message']}")
                doc_result = editor.read_document()
            else:
                print("  Cancelled (invalid input)")

        elif command == 'save':
            if not args:
                # Default path
                base = os.path.splitext(editor.file_path)[0]
                args = f"{base}_edited.docx"
            result = editor.save(args)
            print(f"  {result['message']}")

        elif command == 'blocks':
            blocks = doc_result.get('blocks', [])
            para_blocks = [b for b in blocks if b['type'] == 'paragraph']
            print(f"\nParagraph blocks ({len(para_blocks)} total):")
            for i, block in enumerate(para_blocks[:30]):
                text_preview = block['text'][:60].replace('\n', ' ').strip()
                if len(block['text']) > 60:
                    text_preview += "..."
                print(f"  {block['id']}: {text_preview}")
            if len(para_blocks) > 30:
                print(f"  ... and {len(para_blocks) - 30} more")

        elif command == 'prompt':
            prompt = load_prompt()
            print("\n" + "-"*70)
            print(prompt[:2000])
            if len(prompt) > 2000:
                print(f"\n... ({len(prompt) - 2000} more characters)")
            print("-"*70)

        elif command == 'redline':
            print("To generate a redline, you need:")
            print("  1. Save the edited document first: save /path/to/edited.docx")
            print("  2. Use the XmlPowerToolsEngine separately")
            print("\nAlternatively, test in Open WebUI with the full tool.")

        else:
            print(f"Unknown command: {command}. Type 'help' for available commands.")


def main():
    if len(sys.argv) < 2:
        print("NDA Tool Local Test Harness")
        print("="*40)
        print("\nUsage: python test_nda_workflow.py <document.docx>")
        print("\nThis script simulates Open WebUI's interaction with the NDA tool.")
        print("You can read documents, make edits, and test the workflow locally.")

        test_dir = '/Users/sd/Downloads/nda-table-multilang/'
        if os.path.exists(test_dir):
            print(f"\nAvailable test documents:")
            for f in sorted(os.listdir(test_dir)):
                if f.endswith('.docx') and not f.startswith('~') and 'redline' not in f.lower():
                    print(f"  {test_dir}{f}")
        sys.exit(1)

    filepath = sys.argv[1]

    if not os.path.exists(filepath):
        print(f"Error: File not found: {filepath}")
        sys.exit(1)

    print_header(f"NDA Tool Test: {os.path.basename(filepath)}")

    # Load document
    print("\nLoading document...")
    editor = LocalWordEditor(filepath)
    doc_result = editor.read_document()

    # Print summary
    print_summary(doc_result)

    # If bilingual table, show sample cells
    if doc_result['summary']['document_type'] in ('bilingual_table', 'table_based'):
        print("\n--- Sample Content (first 5 rows) ---")
        print_cells(doc_result)

    # Enter interactive mode
    interactive_mode(editor, doc_result)


if __name__ == '__main__':
    main()
