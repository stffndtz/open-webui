#!/usr/bin/env python3
"""
Test script for NDA Agent with GPT-4o

This script simulates the Open WebUI environment and tests the NDA analysis
workflow with GPT-4o, allowing you to verify prompt effectiveness and
tool call behavior.

Usage:
    # Set your API key
    export OPENAI_API_KEY="sk-..."

    # Run with a test document
    python test_gpt4o_nda.py /path/to/nda.docx

    # Run with verbose output
    python test_gpt4o_nda.py /path/to/nda.docx --verbose

    # Run multiple times to check consistency
    python test_gpt4o_nda.py /path/to/nda.docx --runs 3
"""

import os
import sys
import json
import argparse
import logging
from typing import Dict, List, Any, Optional
from dataclasses import dataclass, field
from datetime import datetime

# Check for OpenAI
try:
    from openai import OpenAI
except ImportError:
    print("Error: openai package not installed. Run: pip install openai")
    sys.exit(1)

# Check for python-docx
try:
    from docx import Document
    from docx.table import Table
    from docx.oxml.table import CT_Tbl
except ImportError:
    print("Error: python-docx not installed. Run: pip install python-docx")
    sys.exit(1)

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(message)s')
log = logging.getLogger(__name__)


# =============================================================================
# Document Parser (simplified from nda-agent-3.0.py)
# =============================================================================

GERMAN_INDICATORS = ['der', 'die', 'das', 'und', 'ist', 'wird', 'sich', 'für', 'auf', 'mit', 'des', 'den']
ENGLISH_INDICATORS = ['the', 'and', 'is', 'will', 'for', 'shall', 'with', 'any', 'this', 'that', 'of', 'to']


def parse_document(filepath: str) -> Dict[str, Any]:
    """Parse a DOCX document and return structured content."""
    doc = Document(filepath)

    # Build table list and classify
    tables = []
    table_metadata = {}
    content_table_indices = []

    for child in doc.element.body:
        if isinstance(child, CT_Tbl):
            table = Table(child, doc)
            tables.append(table)

    # Classify tables
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

    # Build cells array for content tables
    cells = []
    for table_idx in content_table_indices:
        table = tables[table_idx]
        lang_cols = table_metadata[table_idx].get("language_columns", {})

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
                cells.append({
                    "id": cell_id,
                    "table": table_idx,
                    "row": row_idx,
                    "column": col_idx,
                    "language": lang_cols.get(col_idx),
                    "text": cell.text,
                })

    # Build blocks array (paragraphs)
    blocks = []
    block_idx = 0
    for child in doc.element.body:
        from docx.oxml.text.paragraph import CT_P
        if isinstance(child, CT_P):
            from docx.text.paragraph import Paragraph
            para = Paragraph(child, doc)
            block_id = f"block_p_{block_idx}"
            blocks.append({
                "id": block_id,
                "type": "paragraph",
                "text": para.text,
            })
            block_idx += 1
        elif isinstance(child, CT_Tbl):
            table = Table(child, doc)
            block_id = f"block_t_{block_idx}"
            table_idx_in_list = tables.index(table) if table in tables else -1
            blocks.append({
                "id": block_id,
                "type": "table",
                "rows": len(table.rows),
                "columns": len(table.columns) if table.rows else 0,
                "is_content_table": table_idx_in_list in content_table_indices,
            })
            block_idx += 1

    # Pre-flag cells for compliance keywords
    flagged_cells = flag_compliance_cells(cells)

    return {
        "file_path": filepath,
        "summary": {
            "document_type": document_type,
            "detected_languages": list(detected_languages) if detected_languages else None,
            "is_bilingual": is_bilingual,
            "total_blocks": len(blocks),
            "total_cells": len(cells),
            "content_tables": content_table_indices,
        },
        "blocks": blocks,
        "cells": cells,
        "flagged_cells": flagged_cells,
    }


def flag_compliance_cells(cells: List[Dict]) -> Dict[str, List[Dict]]:
    """Pre-scan cells for compliance-relevant keywords."""
    keywords = {
        "req_8_immediately": {
            "de": {"non_compliant": ["unverzüglich"], "compliant": ["ohne schuldhaftes zögern"]},
            "en": {"non_compliant": ["immediately"], "compliant": ["without undue delay"]},
            "fr": {"non_compliant": ["immédiatement"], "compliant": ["sans délai injustifié"]},
            "es": {"non_compliant": ["inmediatamente"], "compliant": ["sin demora indebida"]},
        },
    }

    flagged = {"req_8_immediately": []}

    for cell in cells:
        cell_text = cell.get("text", "").lower()
        cell_lang = cell.get("language", "")
        cell_id = cell.get("id", "")

        for req_key, lang_keywords in keywords.items():
            lang_kw = lang_keywords.get(cell_lang, {})
            non_compliant = lang_kw.get("non_compliant", [])
            compliant = lang_kw.get("compliant", [])

            for term in non_compliant:
                if term in cell_text:
                    is_also_compliant = any(c in cell_text for c in compliant)
                    flagged[req_key].append({
                        "cell_id": cell_id,
                        "language": cell_lang,
                        "found_term": term,
                        "status": "NEEDS_EDIT" if not is_also_compliant else "MIXED",
                    })
                    break

            if not any(term in cell_text for term in non_compliant):
                for term in compliant:
                    if term in cell_text:
                        flagged[req_key].append({
                            "cell_id": cell_id,
                            "language": cell_lang,
                            "found_term": term,
                            "status": "COMPLIANT",
                        })
                        break

    return flagged


# =============================================================================
# Tool Definitions for GPT-4o
# =============================================================================

TOOLS = [
    {
        "type": "function",
        "function": {
            "name": "read_document",
            "description": "Read a DOCX document and return its structured content with block IDs and cell IDs for editing.",
            "parameters": {
                "type": "object",
                "properties": {
                    "file_name": {
                        "type": "string",
                        "description": "Name of the DOCX file to read (optional if already uploaded)"
                    }
                },
                "required": []
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "find_replace",
            "description": "Find and replace specific text within a paragraph block. Use for paragraph-based documents.",
            "parameters": {
                "type": "object",
                "properties": {
                    "block_id": {
                        "type": "string",
                        "description": "The ID of the paragraph block to edit (e.g., 'block_p_5')"
                    },
                    "find_text": {
                        "type": "string",
                        "description": "The EXACT text to find within the paragraph"
                    },
                    "replace_text": {
                        "type": "string",
                        "description": "The text to replace it with"
                    }
                },
                "required": ["block_id", "find_text", "replace_text"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "find_replace_in_cell",
            "description": "Find and replace text within a table cell. Use for bilingual/table-based documents.",
            "parameters": {
                "type": "object",
                "properties": {
                    "cell_id": {
                        "type": "string",
                        "description": "The ID of the cell to edit (e.g., 'cell_0_5_0' for table 0, row 5, column 0)"
                    },
                    "find_text": {
                        "type": "string",
                        "description": "The EXACT text to find within the cell"
                    },
                    "replace_text": {
                        "type": "string",
                        "description": "The text to replace it with"
                    }
                },
                "required": ["cell_id", "find_text", "replace_text"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "insert_block",
            "description": "Insert a new paragraph before or after a reference block or table.",
            "parameters": {
                "type": "object",
                "properties": {
                    "relative_to_block_id": {
                        "type": "string",
                        "description": "The ID of the block to insert relative to"
                    },
                    "position": {
                        "type": "string",
                        "enum": ["before", "after"],
                        "description": "Where to insert: 'before' or 'after' the reference block"
                    },
                    "text": {
                        "type": "string",
                        "description": "Text content for the new paragraph"
                    },
                    "inherit_numbering": {
                        "type": "boolean",
                        "description": "If true, inherit numbering from surrounding paragraphs"
                    }
                },
                "required": ["relative_to_block_id", "position", "text"]
            }
        }
    },
    {
        "type": "function",
        "function": {
            "name": "generate_redline_document",
            "description": "Generate and upload a redline document showing tracked changes.",
            "parameters": {
                "type": "object",
                "properties": {
                    "output_name": {
                        "type": "string",
                        "description": "Name for the output file"
                    }
                },
                "required": ["output_name"]
            }
        }
    }
]


# =============================================================================
# Test Runner
# =============================================================================

@dataclass
class TestResult:
    """Results from a single test run."""
    run_id: int
    document_path: str
    document_type: str
    total_messages: int
    tool_calls: List[Dict[str, Any]] = field(default_factory=list)
    compliance_overview: Optional[str] = None
    planned_edits: List[Dict[str, Any]] = field(default_factory=list)
    errors: List[str] = field(default_factory=list)
    raw_response: Optional[str] = None
    duration_seconds: float = 0.0


def load_system_prompt(prompt_path: str) -> str:
    """Load the NDA agent system prompt."""
    with open(prompt_path, 'r') as f:
        return f.read()


def run_nda_analysis(
    client: OpenAI,
    document_path: str,
    system_prompt: str,
    model: str = "gpt-4o",
    temperature: float = 0,
    max_tokens: int = 4096,
    seed: int = 42,
    verbose: bool = False,
) -> TestResult:
    """Run a single NDA analysis test."""
    import time
    start_time = time.time()

    result = TestResult(
        run_id=seed,
        document_path=document_path,
        document_type="unknown",
        total_messages=0,
    )

    # Parse document
    try:
        doc_content = parse_document(document_path)
        result.document_type = doc_content["summary"]["document_type"]
    except Exception as e:
        result.errors.append(f"Failed to parse document: {e}")
        return result

    # Prepare messages
    messages = [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": f"Please analyze the NDA document I've uploaded and create a redline version with all necessary compliance edits."},
    ]

    # Simulate the read_document response being available
    # In real Open WebUI, the model would call read_document first
    # For testing, we'll provide the document content directly

    if verbose:
        log.info(f"\n{'='*70}")
        log.info(f"Document: {os.path.basename(document_path)}")
        log.info(f"Type: {result.document_type}")
        log.info(f"Cells: {doc_content['summary']['total_cells']}")
        log.info(f"{'='*70}\n")

    # Make API call
    try:
        response = client.chat.completions.create(
            model=model,
            messages=messages,
            tools=TOOLS,
            tool_choice="auto",
            temperature=temperature,
            max_tokens=max_tokens,
            seed=seed,
        )

        result.total_messages = 1

        # Process response
        assistant_message = response.choices[0].message

        if verbose:
            log.info("Assistant Response:")
            log.info("-" * 40)

        # Check for tool calls
        if assistant_message.tool_calls:
            for tool_call in assistant_message.tool_calls:
                tc_info = {
                    "id": tool_call.id,
                    "name": tool_call.function.name,
                    "arguments": json.loads(tool_call.function.arguments),
                }
                result.tool_calls.append(tc_info)

                if verbose:
                    log.info(f"Tool Call: {tc_info['name']}")
                    log.info(f"  Args: {json.dumps(tc_info['arguments'], indent=2)}")

        # Check for content
        if assistant_message.content:
            result.raw_response = assistant_message.content

            # Extract compliance overview if present
            if "## Compliance Overview" in assistant_message.content:
                start = assistant_message.content.find("## Compliance Overview")
                end = assistant_message.content.find("##", start + 5)
                if end == -1:
                    end = len(assistant_message.content)
                result.compliance_overview = assistant_message.content[start:end].strip()

            if verbose:
                log.info(f"\nContent:\n{assistant_message.content[:2000]}")
                if len(assistant_message.content) > 2000:
                    log.info("... [truncated]")

        # If model called read_document, continue the conversation
        if assistant_message.tool_calls:
            for tool_call in assistant_message.tool_calls:
                if tool_call.function.name == "read_document":
                    # Provide the document content
                    messages.append(assistant_message)
                    messages.append({
                        "role": "tool",
                        "tool_call_id": tool_call.id,
                        "content": json.dumps(doc_content),
                    })

                    # Continue conversation
                    response2 = client.chat.completions.create(
                        model=model,
                        messages=messages,
                        tools=TOOLS,
                        tool_choice="auto",
                        temperature=temperature,
                        max_tokens=max_tokens,
                        seed=seed,
                    )

                    result.total_messages += 1
                    msg2 = response2.choices[0].message

                    # Process second response
                    if msg2.tool_calls:
                        for tc in msg2.tool_calls:
                            tc_info = {
                                "id": tc.id,
                                "name": tc.function.name,
                                "arguments": json.loads(tc.function.arguments),
                            }
                            result.tool_calls.append(tc_info)

                            # Track planned edits
                            if tc.function.name in ("find_replace", "find_replace_in_cell"):
                                result.planned_edits.append(tc_info)

                            if verbose:
                                log.info(f"\nTool Call: {tc_info['name']}")
                                log.info(f"  Args: {json.dumps(tc_info['arguments'], indent=2)}")

                    if msg2.content:
                        result.raw_response = msg2.content

                        if "## Compliance Overview" in msg2.content:
                            start = msg2.content.find("## Compliance Overview")
                            end = msg2.content.find("##", start + 5)
                            if end == -1:
                                end = len(msg2.content)
                            result.compliance_overview = msg2.content[start:end].strip()

                        if verbose:
                            log.info(f"\nContent:\n{msg2.content[:3000]}")
                            if len(msg2.content) > 3000:
                                log.info("... [truncated]")

                    break

    except Exception as e:
        result.errors.append(f"API call failed: {e}")

    result.duration_seconds = time.time() - start_time
    return result


def validate_edits(result: TestResult, doc_content: Dict[str, Any]) -> Dict[str, Any]:
    """Validate planned edits against actual document content."""
    validation = {
        "total_edits": len(result.planned_edits),
        "valid_edits": 0,
        "invalid_edits": 0,
        "details": [],
    }

    # Build cell text lookup
    cell_texts = {c["id"]: c["text"] for c in doc_content.get("cells", [])}

    # Build block text lookup for paragraph-based documents
    block_texts = {b["id"]: b.get("text", "") for b in doc_content.get("blocks", []) if b.get("type") == "paragraph"}

    def normalize(t):
        return t.replace('\u2019', "'").replace('\u2018', "'").replace('\u201c', '"').replace('\u201d', '"')

    for edit in result.planned_edits:
        args = edit["arguments"]

        if edit["name"] == "find_replace_in_cell":
            cell_id = args.get("cell_id")
            find_text = args.get("find_text", "")

            if cell_id not in cell_texts:
                validation["invalid_edits"] += 1
                validation["details"].append({
                    "id": cell_id,
                    "status": "INVALID",
                    "reason": f"Cell ID not found",
                })
            elif find_text not in cell_texts[cell_id]:
                if normalize(find_text) in normalize(cell_texts[cell_id]):
                    validation["valid_edits"] += 1
                    validation["details"].append({
                        "id": cell_id,
                        "status": "VALID (normalized)",
                        "find_text": find_text[:50],
                    })
                else:
                    validation["invalid_edits"] += 1
                    validation["details"].append({
                        "id": cell_id,
                        "status": "INVALID",
                        "reason": f"Text not found: '{find_text[:50]}'",
                        "preview": cell_texts[cell_id][:100],
                    })
            else:
                validation["valid_edits"] += 1
                validation["details"].append({
                    "id": cell_id,
                    "status": "VALID",
                    "find_text": find_text[:50],
                })

        elif edit["name"] == "find_replace":
            block_id = args.get("block_id")
            find_text = args.get("find_text", "")

            if block_id not in block_texts:
                validation["invalid_edits"] += 1
                validation["details"].append({
                    "id": block_id,
                    "status": "INVALID",
                    "reason": f"Block ID not found",
                })
            elif find_text not in block_texts[block_id]:
                if normalize(find_text) in normalize(block_texts[block_id]):
                    validation["valid_edits"] += 1
                    validation["details"].append({
                        "id": block_id,
                        "status": "VALID (normalized)",
                        "find_text": find_text[:50],
                    })
                else:
                    validation["invalid_edits"] += 1
                    validation["details"].append({
                        "id": block_id,
                        "status": "INVALID",
                        "reason": f"Text not found: '{find_text[:50]}'",
                        "preview": block_texts[block_id][:100],
                    })
            else:
                validation["valid_edits"] += 1
                validation["details"].append({
                    "id": block_id,
                    "status": "VALID",
                    "find_text": find_text[:50],
                })

    return validation


def main():
    parser = argparse.ArgumentParser(description="Test NDA Agent with GPT-4o")
    parser.add_argument("document", help="Path to DOCX document to analyze")
    parser.add_argument("--model", default="gpt-4o", help="Model to use (default: gpt-4o)")
    parser.add_argument("--temperature", type=float, default=0, help="Temperature (default: 0)")
    parser.add_argument("--max-tokens", type=int, default=4096, help="Max tokens (default: 4096)")
    parser.add_argument("--runs", type=int, default=1, help="Number of runs for consistency check")
    parser.add_argument("--verbose", "-v", action="store_true", help="Verbose output")
    parser.add_argument("--prompt", default="nda-agent-prompt-3.0.md", help="Path to system prompt")
    parser.add_argument("--validate", action="store_true", help="Validate edits against document")
    args = parser.parse_args()

    # Check API key
    api_key = os.environ.get("OPENAI_API_KEY")
    if not api_key:
        print("Error: OPENAI_API_KEY environment variable not set")
        print("Run: export OPENAI_API_KEY='sk-...'")
        sys.exit(1)

    # Check document exists
    if not os.path.exists(args.document):
        print(f"Error: Document not found: {args.document}")
        sys.exit(1)

    # Load system prompt
    prompt_path = args.prompt
    if not os.path.isabs(prompt_path):
        prompt_path = os.path.join(os.path.dirname(__file__), prompt_path)

    if not os.path.exists(prompt_path):
        print(f"Error: System prompt not found: {prompt_path}")
        sys.exit(1)

    system_prompt = load_system_prompt(prompt_path)

    # Initialize client
    client = OpenAI(api_key=api_key)

    # Parse document for validation
    doc_content = parse_document(args.document)

    print(f"\n{'='*70}")
    print(f"NDA Agent GPT-4o Test")
    print(f"{'='*70}")
    print(f"Document: {os.path.basename(args.document)}")
    print(f"Type: {doc_content['summary']['document_type']}")
    print(f"Languages: {doc_content['summary']['detected_languages']}")
    print(f"Cells: {doc_content['summary']['total_cells']}")
    print(f"Model: {args.model}")
    print(f"Temperature: {args.temperature}")
    print(f"Runs: {args.runs}")
    print(f"{'='*70}\n")

    # Run tests
    results = []
    for run in range(args.runs):
        seed = 42 + run
        print(f"\n--- Run {run + 1}/{args.runs} (seed={seed}) ---")

        result = run_nda_analysis(
            client=client,
            document_path=args.document,
            system_prompt=system_prompt,
            model=args.model,
            temperature=args.temperature,
            max_tokens=args.max_tokens,
            seed=seed,
            verbose=args.verbose,
        )
        results.append(result)

        print(f"Duration: {result.duration_seconds:.2f}s")
        print(f"Tool calls: {len(result.tool_calls)}")
        print(f"Planned edits: {len(result.planned_edits)}")

        if result.errors:
            print(f"Errors: {result.errors}")

        # Validate if requested
        if args.validate and result.planned_edits:
            validation = validate_edits(result, doc_content)
            print(f"\nValidation:")
            print(f"  Valid edits: {validation['valid_edits']}/{validation['total_edits']}")
            print(f"  Invalid edits: {validation['invalid_edits']}")
            for detail in validation["details"]:
                status = detail["status"]
                cell_id = detail.get("cell_id", "N/A")
                if "INVALID" in status:
                    print(f"    {cell_id}: {status} - {detail.get('reason', '')}")

    # Summary for multiple runs
    if args.runs > 1:
        print(f"\n{'='*70}")
        print("Consistency Analysis")
        print(f"{'='*70}")

        # Compare tool calls across runs
        all_edits = [set(json.dumps(e) for e in r.planned_edits) for r in results]

        if all(e == all_edits[0] for e in all_edits):
            print("All runs produced identical edit plans")
        else:
            print("Edit plans differed across runs!")
            for i, edits in enumerate(all_edits):
                print(f"  Run {i+1}: {len(edits)} edits")

            # Find common edits
            common = all_edits[0]
            for e in all_edits[1:]:
                common &= e
            print(f"  Common across all runs: {len(common)}")

    print(f"\n{'='*70}")
    print("Test Complete")
    print(f"{'='*70}\n")


if __name__ == "__main__":
    main()
