#!/usr/bin/env python3
"""
Local test script for nda-agent-3.0.py

Usage:
    source .venv/bin/activate
    python test_nda_tool.py /path/to/document.docx

This script tests the core functionality of the NDA tool without needing Open WebUI.
"""

import sys
import os

# Simplified test that doesn't require importing the full module
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph
from docx.oxml.table import CT_Tbl
from docx.oxml.text.paragraph import CT_P
from typing import Dict, List, Any, Optional

# Language detection keywords
GERMAN_INDICATORS = ['der', 'die', 'das', 'und', 'ist', 'wird', 'sich', 'für', 'auf', 'mit', 'des', 'den']
ENGLISH_INDICATORS = ['the', 'and', 'is', 'will', 'for', 'shall', 'with', 'any', 'this', 'that', 'of', 'to']


def classify_and_detect_languages(doc: Document) -> Dict[str, Any]:
    """Classify document type and detect languages."""
    tables = []
    table_metadata = {}

    # Build table list
    for child in doc.element.body:
        if isinstance(child, CT_Tbl):
            table = Table(child, doc)
            tables.append(table)

    # Classify tables
    content_table_indices = []
    for idx, table in enumerate(tables):
        rows = len(table.rows)
        cols = len(table.columns) if table.rows else 0

        metadata = {
            "type": "metadata",
            "rows": rows,
            "columns": cols,
            "language_columns": {},
        }

        # Content table heuristics
        if rows > 10 and cols >= 2:
            col_text = {}
            for row in table.rows[:15]:
                for ci, cell in enumerate(row.cells):
                    col_text[ci] = col_text.get(ci, 0) + len(cell.text)

            text_cols = [i for i, length in col_text.items() if length > 500]

            if len(text_cols) >= 2:
                metadata["type"] = "content"
                content_table_indices.append(idx)

                # Detect languages
                for col in text_cols:
                    text_sample = ""
                    for row in table.rows[:10]:
                        if col < len(row.cells):
                            text_sample += " " + row.cells[col].text.lower() + " "

                    de_score = sum(1 for w in GERMAN_INDICATORS if f" {w} " in text_sample)
                    en_score = sum(1 for w in ENGLISH_INDICATORS if f" {w} " in text_sample)

                    metadata["language_columns"][col] = "de" if de_score > en_score else "en"

        table_metadata[idx] = metadata

    # Determine document type
    detected_languages = set()
    for idx in content_table_indices:
        lang_cols = table_metadata[idx].get("language_columns", {})
        detected_languages.update(lang_cols.values())

    is_bilingual = "de" in detected_languages and "en" in detected_languages

    if content_table_indices:
        document_type = "bilingual_table" if is_bilingual else "table_based"
    else:
        document_type = "paragraph_based"

    return {
        "tables": tables,
        "table_metadata": table_metadata,
        "content_table_indices": content_table_indices,
        "document_type": document_type,
        "detected_languages": list(detected_languages),
        "is_bilingual": is_bilingual,
    }


def test_document(filepath: str):
    """Test the NDA tool with a document."""
    print(f"\n{'='*70}")
    print(f"Testing: {os.path.basename(filepath)}")
    print('='*70)

    doc = Document(filepath)

    # Classify document
    print("\n1. Document Classification")
    print("-" * 40)
    result = classify_and_detect_languages(doc)

    print(f"   Document type:      {result['document_type']}")
    print(f"   Detected languages: {result['detected_languages']}")
    print(f"   Is bilingual:       {result['is_bilingual']}")
    print(f"   Total tables:       {len(result['tables'])}")
    print(f"   Content tables:     {result['content_table_indices']}")

    # Show table details
    print("\n2. Table Details")
    print("-" * 40)
    for idx, meta in result['table_metadata'].items():
        table_type = meta['type'].upper()
        lang_info = ""
        if meta['language_columns']:
            lang_info = f" | Languages: {meta['language_columns']}"
        print(f"   Table {idx}: {meta['rows']:3d} rows, {meta['columns']} cols - {table_type}{lang_info}")

    # If bilingual, show sample cells
    if result['document_type'] in ('bilingual_table', 'table_based'):
        print("\n3. Sample Cells from Content Table(s)")
        print("-" * 40)

        for table_idx in result['content_table_indices']:
            table = result['tables'][table_idx]
            lang_cols = result['table_metadata'][table_idx]['language_columns']

            print(f"\n   Content Table {table_idx}:")

            # Show first 5 non-empty rows
            shown = 0
            for row_idx, row in enumerate(table.rows):
                if shown >= 5:
                    break

                # Check if row has content
                has_content = False
                for col_idx in lang_cols.keys():
                    if col_idx < len(row.cells) and row.cells[col_idx].text.strip():
                        has_content = True
                        break

                if not has_content:
                    continue

                print(f"\n   Row {row_idx}:")
                for col_idx, lang in sorted(lang_cols.items()):
                    if col_idx < len(row.cells):
                        cell = row.cells[col_idx]
                        cell_id = f"cell_{table_idx}_{row_idx}_{col_idx}"
                        text_preview = cell.text[:50].replace('\n', ' ').strip()
                        if len(cell.text) > 50:
                            text_preview += "..."
                        print(f"      [{lang}] {cell_id}: {text_preview}")

                shown += 1

        # Search for "immediately" / "unverzüglich"
        print("\n4. Searching for 'immediately' / 'unverzüglich'")
        print("-" * 40)

        found = []
        for table_idx in result['content_table_indices']:
            table = result['tables'][table_idx]
            lang_cols = result['table_metadata'][table_idx]['language_columns']

            for row_idx, row in enumerate(table.rows):
                for col_idx, lang in lang_cols.items():
                    if col_idx < len(row.cells):
                        cell_text = row.cells[col_idx].text.lower()
                        cell_id = f"cell_{table_idx}_{row_idx}_{col_idx}"

                        if 'immediately' in cell_text:
                            found.append((cell_id, lang, 'immediately', row_idx))
                        elif 'unverzüglich' in cell_text:
                            found.append((cell_id, lang, 'unverzüglich', row_idx))

        if found:
            # Group by row to show paired cells
            by_row = {}
            for cell_id, lang, term, row_idx in found:
                if row_idx not in by_row:
                    by_row[row_idx] = []
                by_row[row_idx].append((cell_id, lang, term))

            for row_idx, cells in sorted(by_row.items()):
                print(f"\n   Row {row_idx}:")
                for cell_id, lang, term in cells:
                    print(f"      [{lang}] {cell_id}: found '{term}'")

                # Check if both languages present
                langs_found = [c[1] for c in cells]
                if 'de' in langs_found and 'en' in langs_found:
                    print("      ✓ Both languages found - can create paired edit")
                else:
                    print(f"      ⚠ Only {langs_found} found - check other column")
        else:
            print("   No occurrences found")

    print("\n" + "="*70)
    print("TEST COMPLETE - Document can be processed")
    print("="*70 + "\n")


def main():
    if len(sys.argv) < 2:
        print("Usage: python test_nda_tool.py <document.docx>")
        print("\nThis script tests the NDA tool functionality locally.")
        print("It reads a document and shows the parsed structure,")
        print("helping verify classification and cell detection work correctly.")

        # List available test documents
        test_dir = '/Users/sd/Downloads/nda-table-multilang/'
        if os.path.exists(test_dir):
            print(f"\nAvailable test documents in {test_dir}:")
            for f in os.listdir(test_dir):
                if f.endswith('.docx') and not f.startswith('~'):
                    print(f"  - {f}")
        sys.exit(1)

    filepath = sys.argv[1]

    if not os.path.exists(filepath):
        print(f"Error: File not found: {filepath}")
        sys.exit(1)

    if not filepath.lower().endswith('.docx'):
        print(f"Error: File must be a .docx file")
        sys.exit(1)

    test_document(filepath)


if __name__ == '__main__':
    main()
