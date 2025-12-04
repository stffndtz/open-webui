"""
title: NORD Fund Report Analyzer
author: 42frontiers
author_url: https://42frontiers.com
version: 2.1.0
license: MIT
description: Analyzes two consecutive quarterly fund reports and produces a standardized performance summary following NORD's specification. Upload Q1 and Q2 PDF reports for comparative analysis. Uses per-document extraction to handle large files without token limits.
"""

import os
import re
import json
import logging
import asyncio
import time
from datetime import datetime
from typing import Dict, List, Any, Optional, Callable, Awaitable, Tuple

from pydantic import BaseModel, Field

log = logging.getLogger(__name__)

# Module-level flags (checked by Open WebUI)
file_handler = True  # This tool handles file uploads directly, skip RAG
citation = False


# ─────────────────────────────────────────────────────────────────────────────
# Module-level helper functions (not exposed as tools)
# ─────────────────────────────────────────────────────────────────────────────

def _parse_quarter_from_filename(filename: str) -> tuple[int, int]:
    """
    Extract quarter and year from filename.
    Returns (year, quarter) for sorting.
    """
    # Pattern: 2025-Q1, Q1 2025, Q1-2025, etc.
    patterns = [
        r'(\d{4})-Q(\d)',  # 2025-Q1
        r'Q(\d)\s*[-_]?\s*(\d{4})',  # Q1 2025 or Q1-2025
        r'(\d{4}).*Q(\d)',  # 2025 anything Q1
    ]

    for i, pattern in enumerate(patterns):
        match = re.search(pattern, filename, re.IGNORECASE)
        if match:
            groups = match.groups()
            if i == 1:  # Q1 2025 format
                return (int(groups[1]), int(groups[0]))
            else:
                return (int(groups[0]), int(groups[1]))

    return (0, 0)  # Unknown


def _extract_fund_name(filename: str) -> str:
    """Extract fund name from filename."""
    # Remove extension
    name = os.path.splitext(filename)[0]

    # Remove common prefixes/suffixes
    name = re.sub(r'\d{4}-Q\d\s*[-_]?\s*', '', name)
    name = re.sub(r'Q\d\s*[-_]?\s*\d{4}\s*[-_]?\s*', '', name)
    name = re.sub(r'NORD\s*(KB\s*)?(IV|V|VI)?\s*[-_]?\s*', '', name, flags=re.IGNORECASE)
    name = re.sub(r'[-_]?\s*Quarterly\s*Report', '', name, flags=re.IGNORECASE)

    return name.strip(' -_') or "Unknown Fund"


def _clean_content(content: str) -> str:
    """
    Clean and filter extracted content to reduce size while keeping key financial data.
    Removes disclaimers, headers, footers, and other noise.
    """
    # Remove common boilerplate patterns
    patterns_to_remove = [
        # Disclaimers and legal text
        r'(?i)disclaimer.*?(?=\n\n|\n#|\n\*\*|$)',
        r'(?i)this (quarterly )?report.*?confidential.*?(?=\n\n|\n#|$)',
        r'(?i)by accepting this.*?(?=\n\n|\n#|$)',
        r'(?i)strictly private and confidential\.?',
        r'(?i)source:\s*\w+\s*analysis\.?',

        # Page markers and navigation
        # r'<!-- Page(Break|Header|Footer|Number).*?-->',
        # r'<figure>\s*</figure>',

        # Repeated headers/footers
        r'(?i)@?\s*\w+\s*capital\s*partners.*?confidential\.?',
        r'(?i)note\s*\(\d+\):\s*[^\n]{0,200}',  # Footnotes (keep short ones)

        # Empty table cells and formatting artifacts
        r'\|\s*\|\s*\|',
        r'\n{3,}',  # Multiple blank lines
    ]

    cleaned = content
    for pattern in patterns_to_remove:
        cleaned = re.sub(pattern, '\n', cleaned, flags=re.MULTILINE | re.DOTALL)

    # Remove lines that are just symbols/whitespace
    lines = cleaned.split('\n')
    filtered_lines = []
    for line in lines:
        stripped = line.strip()
        # Keep lines with actual content (letters/numbers)
        if re.search(r'[a-zA-Z0-9]{2,}', stripped):
            filtered_lines.append(line)
        elif stripped in ['---', '|', '']:  # Keep separators
            filtered_lines.append(line)

    cleaned = '\n'.join(filtered_lines)

    # Collapse multiple blank lines
    cleaned = re.sub(r'\n{3,}', '\n\n', cleaned)

    return cleaned.strip()


def _get_file_content_from_db(file_id: str) -> tuple[str, str]:
    """
    Get file content from Open WebUI's database.
    Returns (filename, content) tuple.

    The content was already extracted during upload using the configured
    document loader (e.g., Azure Document Intelligence for OCR).
    """
    try:
        from open_webui.models.files import Files

        file_obj = Files.get_file_by_id(file_id)
        if file_obj:
            content = file_obj.data.get("content", "") if file_obj.data else ""
            filename = file_obj.filename or "unknown.pdf"
            return (filename, content)
        else:
            raise ValueError(f"File with ID '{file_id}' not found in database")
    except ImportError:
        raise ImportError("Cannot import Open WebUI Files model. Is this running inside Open WebUI?")


def _get_pdf_files_from_metadata(
    files: Optional[List[Dict]]
) -> List[tuple[str, str, str, tuple[int, int]]]:
    """
    Get PDF files from the __files__ metadata.
    Returns list of (filename, file_id, content, (year, quarter)) tuples.

    The content is retrieved from Open WebUI's database where it was
    already extracted during file upload.
    """
    if not files:
        raise ValueError("No files uploaded. Please upload 2 quarterly fund report PDFs.")

    pdf_files = []
    for file_info in files:
        # Get file name - could be in different places depending on context
        name = file_info.get("name", "")
        file_id = file_info.get("id", "")

        # Check if it's a PDF
        if not name.lower().endswith(".pdf"):
            continue

        if not file_id:
            raise ValueError(f"File '{name}' is missing its ID")

        # Try to get content from the file_info first (if already populated)
        content = ""
        file_data = file_info.get("file", {})
        if file_data:
            content = file_data.get("data", {}).get("content", "")

        # If content not in file_info, fetch from database
        if not content:
            db_filename, content = _get_file_content_from_db(file_id)
            if not name:
                name = db_filename

        if not content:
            raise ValueError(f"No content extracted for file '{name}'. The file may not have been processed yet.")

        quarter_info = _parse_quarter_from_filename(name)
        pdf_files.append((name, file_id, content, quarter_info))

    if len(pdf_files) < 2:
        raise ValueError(
            f"Found {len(pdf_files)} PDF(s), need exactly 2. "
            "Please upload both Q1 and Q2 quarterly reports."
        )

    if len(pdf_files) > 2:
        log.warning(f"Found {len(pdf_files)} PDFs, using first 2 by quarter order")

    # Sort by (year, quarter) to get chronological order
    pdf_files.sort(key=lambda x: x[3])

    return pdf_files[:2]


async def _emit_status(emitter, description: str, done: bool = False) -> None:
    """Emit a status event to the UI."""
    if emitter:
        await emitter({
            "type": "status",
            "data": {
                "description": description,
                "done": done,
                "hidden": done,
            }
        })


async def _emit_citation(emitter, title: str, content: str) -> None:
    """Emit a citation event to show extracted data in the UI."""
    if emitter:
        await emitter({
            "type": "citation",
            "data": {
                "document": [content],
                "metadata": [{"source": title}],
                "source": {"name": title},
            }
        })


# ─────────────────────────────────────────────────────────────────────────────
# Expandable Status Indicator (adapted from manifold_pipe.py)
# ─────────────────────────────────────────────────────────────────────────────

class ProgressTracker:
    """
    Simple progress tracker for OpenWebUI tools.

    Emits status events with a single-line description showing current step.
    OpenWebUI tool status events expect a simple description string.

    Usage:
        progress = ProgressTracker(event_emitter=__event_emitter__)
        await progress.update("Extracting metadata...")
        await progress.update("Extracting fund metrics...")
        await progress.finish()
    """

    def __init__(
        self,
        event_emitter: Optional[Callable[[Dict[str, Any]], Awaitable[None]]] = None,
    ) -> None:
        self._event_emitter = event_emitter
        self._started = time.perf_counter()
        self._current_step = ""
        self._done: bool = False

    async def update(self, description: str) -> None:
        """Update the status with a new description."""
        if self._done:
            return
        self._current_step = description
        await self._emit_status()

    async def finish(self) -> None:
        """Mark as done with elapsed time."""
        if self._done:
            return
        elapsed = time.perf_counter() - self._started
        self._current_step = f"✓ Completed in {elapsed:.1f}s"
        self._done = True
        await self._emit_status(done=True)

    async def error(self, error_msg: str) -> None:
        """Mark as failed with error message."""
        if self._done:
            return
        elapsed = time.perf_counter() - self._started
        self._current_step = f"✗ Error after {elapsed:.1f}s: {error_msg}"
        self._done = True
        await self._emit_status(done=True)

    async def _emit_status(self, done: bool = False) -> None:
        """Emit a status event to OpenWebUI."""
        if not self._event_emitter:
            return

        await self._event_emitter({
            "type": "status",
            "data": {
                "description": self._current_step,
                "done": done,
            }
        })


# ─────────────────────────────────────────────────────────────────────────────
# Per-Document Extraction Prompts (NEW - processes one document at a time)
# ─────────────────────────────────────────────────────────────────────────────

PROMPT_EXTRACT_METADATA = """Extract metadata from this single quarterly fund report.

Look for:
1. Fund name (full official name)
2. Quarter and year of the report (e.g., "Q2 2025")
3. Reporting currency (EUR, USD, etc.)
4. List of portfolio company names mentioned

Respond with ONLY valid JSON:
```json
{
  "fund_name": "Full fund name",
  "quarter": "Q2 2025",
  "currency": "EUR",
  "companies": ["Company A", "Company B", ...]
}
```

Use null for any field not found."""

PROMPT_EXTRACT_FUND_METRICS = """You are an advanced data extraction and vision processing expert. Your primary task is to extract specific fund-level performance metrics from quarterly fund reports. Use advanced text and vision capabilities to identify potential metric locations, including tables, charts, or images, which may contain the required data. If necessary, include relevant calculations based on surrounding data.

### Metrics to extract:
- Gross MoIC (Multiple on Invested Capital)
- Net MoIC
- Gross IRR (Internal Rate of Return)
- Net IRR
- DPI (Distributions to Paid-In Capital)
- Total commitments (monetary amount, €m)
- Invested capital (monetary amount, €m)
- Distributions to date (monetary amount, €m)

### Instructions:
- Look for data across all sections of the document, particularly in fund summary and performance-related sections. Do not prioritize any specific page.
- Use vision-based processing for information that might be represented in images, diagrams, or complex tables.
- Validate extracted data against similar or relevant entries in the report to ensure accuracy.
- If any metrics are unavailable or irretrievable, return `null`.

### Output Format:
Respond with **ONLY** valid JSON:
```json
{
  "gross_moic": null,
  "net_moic": null,
  "gross_irr": null,
  "net_irr": null,
  "dpi": null,
  "commitments": null,
  "invested": null,
  "distributions": null,
  "extraction_notes": null
}
```
Example JSON Response:
```json
{
  "gross_moic": 2.0,
  "net_moic": 1.5,
  "gross_irr": 20.0,
  "net_irr": 15.5,
  "dpi": 0.4,
  "commitments": 200.0,
  "invested": 150.0,
  "distributions": 60.0,
  "extraction_notes": "Data sourced from multiple tables and performance sections; DPI calculated from dependent metrics."
}
```

Further Considerations:

- Format MoIC as decimals (e.g., 2.0), IRR/DPI as percentages (e.g., 20.0), and monetary values in millions (e.g., 200.0).
- Extract notes specifying the section or area where data was found, and provide relevant insights (e.g., tricky formatting or probable inferences).
"""

PROMPT_EXTRACT_DEVELOPMENTS = """Extract all developments (transactions) from this quarterly fund report.

Look for:
1. New platform investments (new portfolio companies added this quarter)
2. Add-on acquisitions (bolt-ons by existing portfolio companies)
3. Exits/Divestments (full or partial exits, refinancings, dividend recaps)
4. Distributions to LPs

Respond with ONLY valid JSON:
```json
{
  "new_platforms": [
    {
      "company": "Company Name",
      "sector": "Description",
      "sales": 10.5,
      "ebitda": 2.1,
      "margin": 20.0,
      "ev_ebitda": 8.5,
      "details": "Additional context"
    }
  ],
  "addons": [
    {
      "acquirer": "Portfolio Company",
      "target": "Target Name",
      "sector": "Description",
      "rationale": "Strategic rationale"
    }
  ],
  "exits": [
    {
      "company": "Company Name",
      "type": "Full Exit|Partial Exit|Refinancing|Dividend Recap",
      "distribution": 0.6,
      "moic": 1.9,
      "irr": 18.0,
      "details": "Additional context"
    }
  ]
}
```

Use empty arrays if no developments. Monetary values in millions. Use null for unavailable metrics."""

PROMPT_EXTRACT_COMPANY_METRICS = """Extract detailed financial metrics for each portfolio company from this single quarterly fund report.

For each company mentioned, extract (where available):
- Sales/Revenue (prefer LTM - Last Twelve Months)
- EBITDA (prefer LTM)
- EBITDA margin
- Net Debt
- Leverage (Net Debt / EBITDA)
- Valuation Multiple (EV/EBITDA)
- Gross MoIC for this company
- Entry metrics (if shown: entry sales, entry EBITDA, entry multiple)
- Key performance commentary/quotes
- **ATTENTION FLAGS**: Look for ANY negative indicators or risk factors mentioned:
  - Management changes (CEO/CFO departure, new leadership)
  - Restructuring activities
  - Legal issues (lawsuits, disputes, settlements)
  - Regulatory changes or compliance issues
  - Customer/contract losses
  - Supply chain disruptions
  - Market/competitive pressures
  - Covenant breaches or refinancing needs
  - Any other concerns mentioned by management

Respond with ONLY valid JSON:
```json
{
  "companies": [
    {
      "name": "Company Name",
      "data_period": "LTM Jun 2025",
      "sales": 15.0,
      "ebitda": 3.0,
      "margin": 20.0,
      "net_debt": 2.5,
      "leverage": 0.8,
      "multiple": 6.0,
      "gross_moic": 1.5,
      "entry_sales": 10.0,
      "entry_ebitda": 2.0,
      "entry_multiple": 5.0,
      "commentary": "Key quote about performance",
      "attention_flags": ["Management change: New CEO appointed", "Restructuring: Cost reduction program"]
    }
  ]
}
```

IMPORTANT for attention_flags:
- Use an EMPTY array [] if no negative indicators found
- Be specific: include the type of flag and brief detail
- Extract VERBATIM quotes where possible
- Flag ANY decline mentioned (sales, EBITDA, margin, customer loss, etc.)

Use null for unavailable metrics. All monetary values in millions. Percentages as numbers (20.0 not "20%")."""


# ─────────────────────────────────────────────────────────────────────────────
# Legacy Multi-Step Extraction Prompts (kept for reference/fallback)
# ─────────────────────────────────────────────────────────────────────────────

STEP1_VALIDATION_PROMPT = """You are analyzing two quarterly fund reports. Extract the metadata.

IMPORTANT: You are provided with TWO reports labeled as "REPORT 1 (t=0)" and "REPORT 2 (t=1)". Both reports exist - extract information from both.

From both reports, extract:
1. Fund name (from either report)
2. Quarter and year for each report (t=0 is the earlier quarter, t=1 is the later quarter)
3. Reporting currency for each report
4. List of portfolio company names mentioned across both reports

Respond with ONLY valid JSON in this exact format:
```json
{
  "valid": true,
  "validation_error": null,
  "fund_name": "Full fund name",
  "t0_quarter": "Q1 2025",
  "t1_quarter": "Q2 2025",
  "t0_currency": "EUR",
  "t1_currency": "EUR",
  "currency_consistent": true,
  "companies": ["Company A", "Company B", ...]
}
```

Set valid=true unless:
- The two reports are clearly for DIFFERENT funds (different fund names)
- The reports are for the SAME quarter (not consecutive)

Minor issues like slight name variations or missing data should NOT cause valid=false."""

STEP2_FUND_METRICS_PROMPT = """Extract fund-level performance metrics from both quarterly reports.

Look for these metrics in fund summary/performance sections:
- Gross MoIC (Multiple on Invested Capital)
- Net MoIC
- Gross IRR (Internal Rate of Return)
- Net IRR
- DPI (Distributions to Paid-In)

Respond with ONLY valid JSON:
```json
{
  "t0": {
    "gross_moic": 1.3,
    "net_moic": 1.1,
    "gross_irr": 11.0,
    "net_irr": 4.1,
    "dpi": 0.0
  },
  "t1": {
    "gross_moic": 1.3,
    "net_moic": 1.1,
    "gross_irr": 10.0,
    "net_irr": 3.8,
    "dpi": 0.0
  },
  "notes": "Any relevant notes about data availability"
}
```

Use null for any metric not found. Values should be numbers (MoIC as decimal like 1.3, IRR/DPI as percentage like 11.0)."""

STEP3_DEVELOPMENTS_PROMPT = """Extract all developments (transactions) from the quarterly reports.

Look for:
1. New platform investments (new portfolio companies)
2. Add-on acquisitions (bolt-ons by existing portfolio companies)
3. Exits/Divestments (full or partial exits, refinancings, dividend recaps)
4. Distributions to LPs

For each development, extract available details: company name, target name, sector, size metrics, valuation, rationale.

Respond with ONLY valid JSON:
```json
{
  "new_platforms": [
    {
      "company": "Company Name",
      "sector": "Description",
      "sales": 10.5,
      "ebitda": 2.1,
      "margin": 20.0,
      "ev_ebitda": 8.5,
      "details": "Additional context",
      "page": "p.X"
    }
  ],
  "addons": [
    {
      "acquirer": "Portfolio Company",
      "target": "Target Name",
      "sector": "Description",
      "sales": null,
      "ebitda": null,
      "rationale": "Strategic rationale quote",
      "page": "p.X"
    }
  ],
  "exits": [
    {
      "company": "Company Name",
      "type": "Full Exit|Partial Exit|Refinancing|Dividend Recap",
      "distribution": 0.6,
      "moic": 1.9,
      "irr": 18.0,
      "details": "Additional context",
      "page": "p.X"
    }
  ]
}
```

Use empty arrays if no developments of that type. Monetary values in millions (local currency). Use null for unavailable metrics."""

STEP4_COMPANY_METRICS_PROMPT = """Extract detailed financial metrics for each portfolio company from both reports.

For each company, extract (where available):
- Sales/Revenue (LTM preferred, or quarterly)
- EBITDA (LTM preferred, or quarterly)
- EBITDA margin
- Net Debt
- Leverage (Net Debt / EBITDA)
- Valuation Multiple (EV/EBITDA)
- Gross MoIC
- Entry metrics (if shown)
- FYE (Fiscal Year End)
- Key performance commentary/quotes

Respond with ONLY valid JSON:
```json
{
  "companies": [
    {
      "name": "Company Name",
      "fye": "Dec",
      "entry": {
        "sales": 10.0,
        "ebitda": 2.0,
        "margin": 20.0,
        "net_debt": 3.0,
        "leverage": 1.5,
        "multiple": 6.0
      },
      "t0": {
        "sales": 15.0,
        "ebitda": 3.0,
        "margin": 20.0,
        "net_debt": 2.5,
        "leverage": 0.8,
        "multiple": 6.0,
        "gross_moic": 1.5,
        "data_type": "LTM|Quarterly"
      },
      "t1": {
        "sales": 16.0,
        "ebitda": 3.5,
        "margin": 21.9,
        "net_debt": 2.0,
        "leverage": 0.6,
        "multiple": 6.0,
        "gross_moic": 1.7,
        "data_type": "LTM|Quarterly"
      },
      "commentary": {
        "performance_drivers": "Key quote about performance",
        "outlook": "Forward-looking statements",
        "concerns": "Any issues mentioned",
        "page_refs": ["t=0; p.X", "t=1; p.Y"]
      }
    }
  ]
}
```

Use null for unavailable metrics. All monetary values in millions. Percentages as numbers (20.0 not "20%")."""

STEP5_FINAL_REPORT_PROMPT = """You are a private equity fund performance analyst generating the final QUARTERLY FUND PERFORMANCE ANALYSIS report.

You have been provided with structured data extracted from the quarterly reports:
- Validation & metadata
- Fund-level metrics
- Developments (investments, exits)
- Company-level metrics

---

## DATA CONVENTIONS

| Type | Format | Example |
|------|--------|---------|
| Money | €X.Xm or $X.Xm (millions, 1 decimal) | €12.5m |
| Multiples | X.Xx suffix | 7.25x |
| Percentages | X.X% (1 decimal) | 14.7% |
| QoQ Change | (t=1 - t=0) / |t=0| × 100% | +15.3% |
| pp Change | t=1% - t=0% | +2.5 pp |

---

## RATING METHODOLOGY (CRITICAL - Apply Strictly)

**Use LTM (Last Twelve Months) figures for all QoQ calculations.**

| Rating | Criteria |
|--------|----------|
| 🟢 **Positive** | LTM EBITDA ≥ +10% QoQ **AND** Gross MoIC increased (t=1 > t=0, not equal) |
| 🔴 **Negative** | LTM EBITDA ≤ -10% QoQ **OR** LTM Sales ≤ -15% QoQ **OR** EBITDA Margin ≤ -15pp QoQ |
| ⚪ **Neutral** | Everything else (includes mixed signals, unchanged MoIC, or missing data) |

**CRITICAL Rules:**
- 🟢 requires BOTH conditions: EBITDA ≥ +10% **AND** MoIC actually increased
- If EBITDA is +15% but MoIC is unchanged (1.1x → 1.1x) → ⚪ Neutral
- If EBITDA is +12% but MoIC decreased → ⚪ Neutral
- If MoIC increased but EBITDA < +10% → ⚪ Neutral

**Negative EBITDA handling:**
- When EBITDA is negative in both periods, measure improvement/deterioration:
  - Loss reduced (less negative): EBITDA -€200k → -€150k = +25% improvement
  - Loss increased (more negative): EBITDA -€150k → -€200k = -33% deterioration
- Formula: ((t=1) - (t=0)) / |t=0| × 100 (use absolute value of t=0 in denominator)

---

## OUTPUT FORMAT (Output directly - no code fences)

# QUARTERLY FUND PERFORMANCE ANALYSIS

**Fund:** [Name from metadata]
**Period:** [t=0 Quarter] vs [t=1 Quarter]
**Currency:** [EUR/USD]
**Analysis Date:** [Date from parameters]

---

## 1. DEVELOPMENTS

### New Investments - Platforms
[If none: "None this quarter"]

**[Company Name]**
- Activity: [sector/description]
- Size: Sales €X.Xm, EBITDA €X.Xm (margin X.X%)
- Multiple: X.Xx EV/EBITDA

### New Investments - Add-ons
[If none: "None this quarter"]

**[Portfolio Co] acquired [Target]**
- Activity: [description]
- Size: Sales €X.Xm, EBITDA €X.Xm
- Rationale: [brief quote with citation]

### Exits / Distributions
[If none: "None this quarter"]

**[Company] - [Full Exit / Partial / Recap]**
- Distribution: €X.Xm
- DPI impact: X.Xx → X.Xx
- Return: X.Xx MoIC, XX.X% IRR

---

## 2. FUND PERFORMANCE

| Metric | [t=0] | [t=1] | Δ QoQ |
|--------|-------|-------|-------|
| Gross MoIC | X.Xx | X.Xx | +X.X% |
| Net MoIC | X.Xx | X.Xx | +X.X% |
| Gross IRR | XX.X% | XX.X% | +X.X pp |
| Net IRR | XX.X% | XX.X% | +X.X pp |
| DPI | X.Xx | X.Xx | +X.X% |

**Commentary:** [2-3 sentences on key drivers. Link DPI changes to distributions above.]

---

## 3. COMPANY PERFORMANCE

### Summary Table

| Company | Rating | Sales Δ | EBITDA Δ | MoIC Δ | ⚠️ Attention Flags |
|---------|--------|---------|----------|--------|-------------------|
| [Name] | 🟢 | +X.X% | +X.X% | X.Xx→X.Xx | — |
| [Name] | ⚪ | +X.X% | -X.X% | X.Xx→X.Xx | Margin pressure |
| [Name] | 🔴 | -X.X% | -X.X% | X.Xx→X.Xx | Mgmt change; Restructuring |

**Attention Flag Legend:**
- Use "—" if no flags
- Abbreviate: "Mgmt change" (management), "Restructuring", "Legal", "Regulatory", "Customer loss", "Cost pressure", "Covenant", etc.
- Multiple flags separated by semicolon

**IMPORTANT: Every company in this table MUST have a detailed section below.**

### [Company Name] [RATING EMOJI]

**Metrics:**
| | Entry | LTM t=0 | LTM t=1 | Δ QoQ |
|---|-------|---------|---------|-------|
| Sales | €X.Xm | €X.Xm | €X.Xm | +X.X% |
| EBITDA | €X.Xm | €X.Xm | €X.Xm | +X.X% |
| Margin | X.X% | X.X% | X.X% | +X.X pp |
| Net Debt | €X.Xm | €X.Xm | €X.Xm | |
| Leverage | X.Xx | X.Xx | X.Xx | |
| Multiple | X.Xx | X.Xx | X.Xx | |
| Gross MoIC | - | X.Xx | X.Xx | +X.X% |

**Checklist**:
- EBITDA Δ: ___% → ≥+10%? [🔥/🗑️]
- MoIC t=0: ___ → MoIC t=1: ___ → t=1 > t=0? [🔥/🗑️]
- Sales Δ: ___% → ≤-15%? [🔥/🗑️]
- Margin Δ: ___ pp → ≤-15pp? [🔥/🗑️]

**Decision Process**:
1. **🟢 Positive Rating**: Line 1 = [🔥] AND Line 2 = [🔥].
2. **🔴 Negative Rating**: Line 3 = [🔥] OR Line 4 = [🔥] OR EBITDA ≤ -10%.
3. **⚪ Neutral Rating**: All other cases; e.g., mixed signals or unchanged MoIC (e.g., 1.1x → 1.1x).

**Final Rating:** [Apply decision above - if Line 2 = [🗑️], rating CANNOT be 🟢]

**Drivers/Concerns:** [1-2 sentences with citation]

**Management Response (if 🔴):** "[quote]" [Report: t=1; p.X]

---

## 4. RANKINGS

### 💸 TOP 2

**1. [Company]** [RATING]
- EBITDA: +X.X% QoQ
- MoIC: X.Xx → X.Xx
- Key driver: [brief]

**2. [Company]** [RATING]
- EBITDA: +X.X% QoQ
- MoIC: X.Xx → X.Xx
- Key driver: [brief]

### 😱 BOTTOM 2

**1. [Company]** [RATING]
- EBITDA: -X.X% QoQ
- MoIC: X.Xx → X.Xx
- Challenge: [brief]

**2. [Company]** [RATING]
- EBITDA: -X.X% QoQ
- MoIC: X.Xx → X.Xx
- Challenge: [brief]

---

## 5. DATA QUALITY

**Missing Data:**
- [List any N/A items]

**Currency Notes:**
- [Any FX or reporting currency issues]

**OCR Issues:**
- [Any data quality concerns]

---

## FINAL VALIDATION (Do Before Outputting)

Before outputting, verify:
1. **Calculations match:** Summary table Δ% = Detailed section Δ%
2. **Rating-Checklist consistency:** For EACH company, Final Rating matches checklist logic
   - If checklist Line 2 (MoIC increased) = NO → Final Rating MUST NOT be 🟢
   - If checklist Line 1 (EBITDA ≥+10%) = NO → Final Rating MUST NOT be 🟢
3. **Negative EBITDA check:** If EBITDA was negative in t=0, verify % change direction is correct
4. **Completeness:** Every company in summary has detailed section
5. **No fabrication:** All numbers from provided data"""


async def _call_llm_azure_openai(
    endpoint: str,
    api_key: str,
    deployment: str,
    api_version: str,
    system_prompt: str,
    user_message: str,
) -> str:
    """Call Azure OpenAI API directly."""
    import aiohttp

    url = f"{endpoint.rstrip('/')}/openai/deployments/{deployment}/chat/completions?api-version={api_version}"

    payload = {
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_message},
        ],
        "temperature": 0.1,
        "max_tokens": 4096,
    }

    headers = {
        "Content-Type": "application/json",
        "api-key": api_key,
    }

    log.info(f"Azure OpenAI: Calling {deployment} at {endpoint}")

    async with aiohttp.ClientSession() as session:
        async with session.post(url, json=payload, headers=headers) as response:
            if response.status != 200:
                error_text = await response.text()
                raise RuntimeError(f"Azure OpenAI API error {response.status}: {error_text}")

            data = await response.json()

    if "choices" in data and len(data["choices"]) > 0:
        content = data["choices"][0].get("message", {}).get("content", "")
        if content:
            return content

    raise RuntimeError(f"Unexpected Azure OpenAI response: {data}")


async def _call_llm_openai(
    api_key: str,
    model: str,
    system_prompt: str,
    user_message: str,
) -> str:
    """Call OpenAI API directly."""
    import aiohttp

    url = "https://api.openai.com/v1/chat/completions"

    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_message},
        ],
        "temperature": 0.1,
        "max_tokens": 4096,
    }

    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}",
    }

    log.info(f"OpenAI: Calling {model}")

    async with aiohttp.ClientSession() as session:
        async with session.post(url, json=payload, headers=headers) as response:
            if response.status != 200:
                error_text = await response.text()
                raise RuntimeError(f"OpenAI API error {response.status}: {error_text}")

            data = await response.json()

    if "choices" in data and len(data["choices"]) > 0:
        content = data["choices"][0].get("message", {}).get("content", "")
        if content:
            return content

    raise RuntimeError(f"Unexpected OpenAI response: {data}")


async def _call_llm_anthropic(
    api_key: str,
    model: str,
    system_prompt: str,
    user_message: str,
) -> str:
    """Call Anthropic API directly."""
    import aiohttp

    url = "https://api.anthropic.com/v1/messages"

    payload = {
        "model": model,
        "max_tokens": 4096,
        "system": system_prompt,
        "messages": [
            {"role": "user", "content": user_message},
        ],
    }

    headers = {
        "Content-Type": "application/json",
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
    }

    log.info(f"Anthropic: Calling {model}")

    async with aiohttp.ClientSession() as session:
        async with session.post(url, json=payload, headers=headers) as response:
            if response.status != 200:
                error_text = await response.text()
                raise RuntimeError(f"Anthropic API error {response.status}: {error_text}")

            data = await response.json()

    if "content" in data and len(data["content"]) > 0:
        content = data["content"][0].get("text", "")
        if content:
            return content

    raise RuntimeError(f"Unexpected Anthropic response: {data}")


def _parse_json_response(response: str) -> dict:
    """Extract and parse JSON from LLM response, handling markdown code blocks."""
    # Try to find JSON in code blocks first
    json_match = re.search(r'```(?:json)?\s*([\s\S]*?)\s*```', response)
    if json_match:
        json_str = json_match.group(1)
    else:
        # Try to find raw JSON
        json_str = response.strip()

    try:
        return json.loads(json_str)
    except json.JSONDecodeError as e:
        log.warning(f"Failed to parse JSON response: {e}")
        log.debug(f"Response was: {response[:500]}...")
        return {"error": str(e), "raw_response": response}


# ─────────────────────────────────────────────────────────────────────────────
# Tools class (only public methods are exposed)
# ─────────────────────────────────────────────────────────────────────────────

class Tools:
    """
    NORD Fund Report Analyzer for Open WebUI.

    Extracts text from two consecutive quarterly fund report PDFs and prepares
    them for analysis. The model then applies the system prompt to generate
    a standardized performance summary.

    This tool reuses the content already extracted by Open WebUI during file
    upload (using Document Intelligence, Tika, or other configured loaders),
    so it works with both text-based and scanned PDFs automatically.

    Workflow:
        1. User uploads Q1 and Q2 PDF reports
        2. Open WebUI extracts content during upload (OCR if needed)
        3. Tool retrieves the already-extracted content
        4. Returns structured content for model to analyze
    """

    class Valves(BaseModel):
        """Admin-configurable settings."""
        ENABLE_MULTI_STEP: bool = Field(
            default=True,
            description="Enable multi-step extraction for improved accuracy. When disabled, uses single-pass analysis."
        )
        # --- LLM Provider Configuration ---
        LLM_PROVIDER: str = Field(
            default="azure_openai",
            description="LLM provider for multi-step extraction: 'azure_openai', 'openai', or 'anthropic'"
        )
        AZURE_OPENAI_ENDPOINT: str = Field(
            default="",
            description="Azure OpenAI endpoint URL (e.g., https://your-resource.openai.azure.com)"
        )
        AZURE_OPENAI_API_KEY: str = Field(
            default="",
            description="Azure OpenAI API key"
        )
        AZURE_OPENAI_DEPLOYMENT: str = Field(
            default="gpt-4o",
            description="Azure OpenAI deployment name (e.g., 'gpt-4o')"
        )
        AZURE_OPENAI_API_VERSION: str = Field(
            default="2024-08-01-preview",
            description="Azure OpenAI API version"
        )
        OPENAI_API_KEY: str = Field(
            default="",
            description="OpenAI API key (if using 'openai' provider)"
        )
        OPENAI_MODEL: str = Field(
            default="gpt-4o",
            description="OpenAI model name (e.g., 'gpt-4o', 'gpt-4-turbo')"
        )
        ANTHROPIC_API_KEY: str = Field(
            default="",
            description="Anthropic API key (if using 'anthropic' provider)"
        )
        ANTHROPIC_MODEL: str = Field(
            default="claude-sonnet-4-20250514",
            description="Anthropic model name (e.g., 'claude-sonnet-4-20250514')"
        )
        # --- Content Settings ---
        MAX_CHARS_PER_PDF: int = Field(
            default=50000,
            description="Maximum characters to use per PDF after cleaning. Multi-step mode can handle larger content (~12K tokens per PDF)."
        )
        MAX_CHARS_PER_PDF_SINGLE_PASS: int = Field(
            default=35000,
            description="Maximum characters per PDF in single-pass mode (~9K tokens per PDF)."
        )

    def __init__(self):
        """Initialize the Fund Report Analyzer."""
        self.valves = self.Valves()
        self.file_handler = True  # Accept file uploads
        self.citation = False

    async def _call_llm(self, system_prompt: str, user_message: str) -> str:
        """
        Call the configured LLM provider directly.

        Uses the valve settings to determine which provider and credentials to use.
        """
        provider = self.valves.LLM_PROVIDER.lower()

        if provider == "azure_openai":
            if not self.valves.AZURE_OPENAI_ENDPOINT or not self.valves.AZURE_OPENAI_API_KEY:
                raise ValueError("Azure OpenAI endpoint and API key must be configured in valves")
            return await _call_llm_azure_openai(
                endpoint=self.valves.AZURE_OPENAI_ENDPOINT,
                api_key=self.valves.AZURE_OPENAI_API_KEY,
                deployment=self.valves.AZURE_OPENAI_DEPLOYMENT,
                api_version=self.valves.AZURE_OPENAI_API_VERSION,
                system_prompt=system_prompt,
                user_message=user_message,
            )
        elif provider == "openai":
            if not self.valves.OPENAI_API_KEY:
                raise ValueError("OpenAI API key must be configured in valves")
            return await _call_llm_openai(
                api_key=self.valves.OPENAI_API_KEY,
                model=self.valves.OPENAI_MODEL,
                system_prompt=system_prompt,
                user_message=user_message,
            )
        elif provider == "anthropic":
            if not self.valves.ANTHROPIC_API_KEY:
                raise ValueError("Anthropic API key must be configured in valves")
            return await _call_llm_anthropic(
                api_key=self.valves.ANTHROPIC_API_KEY,
                model=self.valves.ANTHROPIC_MODEL,
                system_prompt=system_prompt,
                user_message=user_message,
            )
        else:
            raise ValueError(f"Unknown LLM provider: {provider}. Use 'azure_openai', 'openai', or 'anthropic'")

    async def _run_multi_step_analysis(
        self,
        q1_text: str,
        q2_text: str,
        q1_name: str,
        q2_name: str,
        q1_label: str,
        q2_label: str,
        __event_emitter__=None,
    ) -> str:
        """
        Run multi-step extraction and analysis pipeline.

        Steps:
        1. Validation & metadata extraction
        2. Fund-level metrics extraction
        3. Developments extraction (investments, exits)
        4. Company-level metrics extraction
        5. Final report generation

        Returns:
            Complete formatted analysis report.
        """
        report_content = f"""## REPORT 1 (t=0 - {q1_label}):
**Source:** {q1_name}

{q1_text}

---

## REPORT 2 (t=1 - {q2_label}):
**Source:** {q2_name}

{q2_text}"""

        results = {}

        # Step 1: Validation
        await _emit_status(__event_emitter__, "Step 1/5: Validating reports...")
        log.info("Multi-step: Running validation...")
        try:
            validation_response = await self._call_llm(
                system_prompt=STEP1_VALIDATION_PROMPT,
                user_message=report_content,
            )
            results["validation"] = _parse_json_response(validation_response)
            log.info(f"Validation result: {results['validation']}")

            # Log validation issues but don't stop - continue with analysis
            if results["validation"].get("valid") is False:
                error_msg = results["validation"].get("validation_error", "Unknown validation error")
                log.warning(f"Validation flagged issue (continuing anyway): {error_msg}")
                # Don't return early - proceed with analysis

        except Exception as e:
            log.exception(f"Validation step failed: {e}")
            results["validation"] = {"error": str(e)}
            # Continue with analysis even if validation fails

        # Step 2: Fund metrics
        await _emit_status(__event_emitter__, "Step 2/5: Extracting fund metrics...")
        log.info("Multi-step: Extracting fund metrics...")
        try:
            fund_response = await self._call_llm(
                system_prompt=STEP2_FUND_METRICS_PROMPT,
                user_message=report_content,
            )
            results["fund_metrics"] = _parse_json_response(fund_response)
            log.info(f"Fund metrics: {results['fund_metrics']}")
        except Exception as e:
            log.exception(f"Fund metrics step failed: {e}")
            results["fund_metrics"] = {"error": str(e)}

        # Step 3: Developments
        await _emit_status(__event_emitter__, "Step 3/5: Extracting developments...")
        log.info("Multi-step: Extracting developments...")
        try:
            dev_response = await self._call_llm(
                system_prompt=STEP3_DEVELOPMENTS_PROMPT,
                user_message=report_content,
            )
            results["developments"] = _parse_json_response(dev_response)
            log.info(f"Developments: {results['developments']}")
        except Exception as e:
            log.exception(f"Developments step failed: {e}")
            results["developments"] = {"error": str(e)}

        # Step 4: Company metrics
        await _emit_status(__event_emitter__, "Step 4/5: Extracting company metrics...")
        log.info("Multi-step: Extracting company metrics...")
        try:
            company_response = await self._call_llm(
                system_prompt=STEP4_COMPANY_METRICS_PROMPT,
                user_message=report_content,
            )
            results["company_metrics"] = _parse_json_response(company_response)
            log.info(f"Company metrics: found {len(results['company_metrics'].get('companies', []))} companies")
        except Exception as e:
            log.exception(f"Company metrics step failed: {e}")
            results["company_metrics"] = {"error": str(e)}

        # Step 5: Final report generation
        await _emit_status(__event_emitter__, "Step 5/5: Generating final report...")
        log.info("Multi-step: Generating final report...")

        # Prepare structured data for final step
        structured_data = f"""## Extracted Data from Quarterly Reports

### Metadata & Validation
```json
{json.dumps(results.get('validation', {}), indent=2)}
```

### Fund-Level Metrics
```json
{json.dumps(results.get('fund_metrics', {}), indent=2)}
```

### Developments (Investments, Exits)
```json
{json.dumps(results.get('developments', {}), indent=2)}
```

### Company-Level Metrics
```json
{json.dumps(results.get('company_metrics', {}), indent=2)}
```

### Analysis Parameters
- **t=0 (Earlier Quarter):** {q1_label}
- **t=1 (Later Quarter):** {q2_label}
- **Analysis Date:** {datetime.now().strftime("%Y-%m-%d")}
- **Currency:** {results.get('validation', {}).get('t0_currency', 'EUR')}
"""

        try:
            final_response = await self._call_llm(
                system_prompt=STEP5_FINAL_REPORT_PROMPT,
                user_message=structured_data,
            )
            await _emit_status(__event_emitter__, "Analysis complete", done=True)
            return final_response

        except Exception as e:
            log.exception(f"Final report generation failed: {e}")
            await _emit_status(__event_emitter__, "Error in final report generation", done=True)

            # Fall back to returning the structured data
            return f"""# Multi-Step Analysis Results

The final report generation encountered an error, but here is the extracted data:

{structured_data}

**Error:** {str(e)}"""

    async def _run_multi_step_analysis_v2(
        self,
        q1_text: str,
        q2_text: str,
        q1_name: str,
        q2_name: str,
        q1_label: str,
        q2_label: str,
        __event_emitter__=None,
    ) -> str:
        """
        Run per-document multi-step extraction and analysis pipeline.

        This version processes each document individually to avoid token limits
        with large PDFs, then merges results for the final report.

        Steps:
        1. Extract metadata from Q1 and Q2 (parallel)
        2. Validate metadata (same fund, consecutive quarters)
        3. Extract fund metrics from Q1 and Q2 (parallel)
        4. Extract developments from Q2 only
        5. Extract company metrics from Q1 and Q2 (parallel)
        6. Generate final report

        Returns:
            Complete formatted analysis report.
        """
        progress = ProgressTracker(event_emitter=__event_emitter__)
        results = {}

        # ─────────────────────────────────────────────────────────────────────
        # Step 1: Extract metadata from both documents (parallel)
        # ─────────────────────────────────────────────────────────────────────
        await progress.update("Step 1/6: Extracting metadata from Q1 and Q2...")

        async def extract_metadata(doc_text: str, doc_name: str, label: str) -> dict:
            """Extract metadata from a single document."""
            user_content = f"""## Document: {doc_name} ({label})

{doc_text}"""
            try:
                response = await self._call_llm(
                    system_prompt=PROMPT_EXTRACT_METADATA,
                    user_message=user_content,
                )
                return _parse_json_response(response)
            except Exception as e:
                log.exception(f"Metadata extraction failed for {doc_name}: {e}")
                return {"error": str(e)}

        # Run Q1 and Q2 metadata extraction in parallel
        log.info("Extracting metadata from Q1 and Q2 in parallel...")
        q1_meta_task = extract_metadata(q1_text, q1_name, q1_label)
        q2_meta_task = extract_metadata(q2_text, q2_name, q2_label)
        q1_meta, q2_meta = await asyncio.gather(q1_meta_task, q2_meta_task)

        results["q1_metadata"] = q1_meta
        results["q2_metadata"] = q2_meta

        # Update status with metadata info
        q1_fund = q1_meta.get("fund_name", "Unknown")
        q1_qtr = q1_meta.get("quarter", q1_label)
        q2_fund = q2_meta.get("fund_name", "Unknown")
        q2_qtr = q2_meta.get("quarter", q2_label)

        await progress.update(f"Step 1/6: Metadata extracted - {q1_fund} ({q1_qtr} → {q2_qtr})")

        # Emit citation for metadata
        await _emit_citation(
            __event_emitter__,
            "Extracted Metadata",
            f"**Q1 ({q1_label}):**\n```json\n{json.dumps(q1_meta, indent=2)}\n```\n\n"
            f"**Q2 ({q2_label}):**\n```json\n{json.dumps(q2_meta, indent=2)}\n```"
        )

        # ─────────────────────────────────────────────────────────────────────
        # Step 2: Validate metadata
        # ─────────────────────────────────────────────────────────────────────
        await progress.update("Step 2/6: Validating reports...")

        # Check if same fund (fuzzy match on fund name)
        fund_match = (
            q1_fund.lower().replace(" ", "") == q2_fund.lower().replace(" ", "") or
            q1_fund in q2_fund or q2_fund in q1_fund
        )

        validation_result = {
            "fund_name": q2_fund or q1_fund,
            "t0_quarter": q1_qtr,
            "t1_quarter": q2_qtr,
            "currency": q2_meta.get("currency") or q1_meta.get("currency", "EUR"),
            "same_fund": fund_match,
            "companies": list(set(
                (q1_meta.get("companies") or []) + (q2_meta.get("companies") or [])
            )),
        }
        results["validation"] = validation_result

        validation_status = "✓ Validated" if fund_match else "⚠ Fund mismatch"
        await progress.update(f"Step 2/6: {validation_status} - {q1_qtr} → {q2_qtr}")

        # ─────────────────────────────────────────────────────────────────────
        # Step 3: Extract fund metrics from both documents (parallel)
        # ─────────────────────────────────────────────────────────────────────
        await progress.update("Step 3/6: Extracting fund metrics...")

        async def extract_fund_metrics(doc_text: str, doc_name: str, label: str) -> dict:
            """Extract fund metrics from a single document."""
            user_content = f"""## Document: {doc_name} ({label})

{doc_text}"""
            try:
                response = await self._call_llm(
                    system_prompt=PROMPT_EXTRACT_FUND_METRICS,
                    user_message=user_content,
                )
                return _parse_json_response(response)
            except Exception as e:
                log.exception(f"Fund metrics extraction failed for {doc_name}: {e}")
                return {"error": str(e)}

        # Run Q1 and Q2 fund metrics extraction in parallel
        log.info("Extracting fund metrics from Q1 and Q2 in parallel...")
        q1_fund_task = extract_fund_metrics(q1_text, q1_name, q1_label)
        q2_fund_task = extract_fund_metrics(q2_text, q2_name, q2_label)
        q1_fund_metrics, q2_fund_metrics = await asyncio.gather(q1_fund_task, q2_fund_task)

        results["fund_metrics"] = {
            "t0": q1_fund_metrics,
            "t1": q2_fund_metrics,
        }

        # Update status with fund metrics summary
        q1_moic = q1_fund_metrics.get("gross_moic")
        q2_moic = q2_fund_metrics.get("gross_moic")
        q1_moic_str = f"{q1_moic}x" if q1_moic else "N/A"
        q2_moic_str = f"{q2_moic}x" if q2_moic else "N/A"
        await progress.update(f"Step 3/6: Fund metrics extracted - MoIC: {q1_moic_str} → {q2_moic_str}")

        # Emit citation for fund metrics
        await _emit_citation(
            __event_emitter__,
            "Fund Metrics",
            f"**Q1 ({q1_label}):**\n```json\n{json.dumps(q1_fund_metrics, indent=2)}\n```\n\n"
            f"**Q2 ({q2_label}):**\n```json\n{json.dumps(q2_fund_metrics, indent=2)}\n```"
        )

        # ─────────────────────────────────────────────────────────────────────
        # Step 4: Extract developments from Q2 only
        # ─────────────────────────────────────────────────────────────────────
        await progress.update("Step 4/6: Extracting developments...")

        user_content = f"""## Document: {q2_name} ({q2_label})

{q2_text}"""
        try:
            log.info("Extracting developments from Q2...")
            dev_response = await self._call_llm(
                system_prompt=PROMPT_EXTRACT_DEVELOPMENTS,
                user_message=user_content,
            )
            results["developments"] = _parse_json_response(dev_response)
        except Exception as e:
            log.exception(f"Developments extraction failed: {e}")
            results["developments"] = {"error": str(e)}

        # Update status with developments summary
        devs = results["developments"]
        new_platforms = len(devs.get("new_platforms", []))
        addons = len(devs.get("addons", []))
        exits = len(devs.get("exits", []))
        dev_summary = []
        if new_platforms:
            dev_summary.append(f"{new_platforms} new investment(s)")
        if addons:
            dev_summary.append(f"{addons} add-on(s)")
        if exits:
            dev_summary.append(f"{exits} exit(s)")
        dev_status = ", ".join(dev_summary) if dev_summary else "No developments"
        await progress.update(f"Step 4/6: Developments - {dev_status}")

        # Emit citation for developments
        await _emit_citation(
            __event_emitter__,
            "Developments",
            f"```json\n{json.dumps(devs, indent=2)}\n```"
        )

        # ─────────────────────────────────────────────────────────────────────
        # Step 5: Extract company metrics from both documents (parallel)
        # ─────────────────────────────────────────────────────────────────────
        await progress.update("Step 5/6: Extracting company metrics...")

        async def extract_company_metrics(doc_text: str, doc_name: str, label: str) -> dict:
            """Extract company metrics from a single document."""
            user_content = f"""## Document: {doc_name} ({label})

{doc_text}"""
            try:
                response = await self._call_llm(
                    system_prompt=PROMPT_EXTRACT_COMPANY_METRICS,
                    user_message=user_content,
                )
                return _parse_json_response(response)
            except Exception as e:
                log.exception(f"Company metrics extraction failed for {doc_name}: {e}")
                return {"error": str(e)}

        # Run Q1 and Q2 company metrics extraction in parallel
        log.info("Extracting company metrics from Q1 and Q2 in parallel...")
        q1_company_task = extract_company_metrics(q1_text, q1_name, q1_label)
        q2_company_task = extract_company_metrics(q2_text, q2_name, q2_label)
        q1_companies, q2_companies = await asyncio.gather(q1_company_task, q2_company_task)

        # Merge company metrics from Q1 and Q2
        q1_company_list = q1_companies.get("companies", [])
        q2_company_list = q2_companies.get("companies", [])

        # Create a merged view: company name -> {entry, t0, t1, commentary}
        merged_companies = {}
        for co in q1_company_list:
            name = co.get("name", "Unknown")
            merged_companies[name] = {
                "name": name,
                "entry": {k: v for k, v in co.items() if k.startswith("entry_")},
                "t0": {k: v for k, v in co.items() if k not in ["name", "commentary"] and not k.startswith("entry_")},
                "t1": None,
                "commentary": {"t0": co.get("commentary")},
            }

        for co in q2_company_list:
            name = co.get("name", "Unknown")
            if name in merged_companies:
                merged_companies[name]["t1"] = {
                    k: v for k, v in co.items()
                    if k not in ["name", "commentary"] and not k.startswith("entry_")
                }
                merged_companies[name]["commentary"]["t1"] = co.get("commentary")
                # Update entry if not present in Q1
                if not merged_companies[name]["entry"]:
                    merged_companies[name]["entry"] = {k: v for k, v in co.items() if k.startswith("entry_")}
            else:
                merged_companies[name] = {
                    "name": name,
                    "entry": {k: v for k, v in co.items() if k.startswith("entry_")},
                    "t0": None,
                    "t1": {k: v for k, v in co.items() if k not in ["name", "commentary"] and not k.startswith("entry_")},
                    "commentary": {"t1": co.get("commentary")},
                }

        results["company_metrics"] = {
            "companies": list(merged_companies.values()),
            "q1_raw": q1_company_list,
            "q2_raw": q2_company_list,
        }

        # Update status with company count
        await progress.update(f"Step 5/6: Company metrics - {len(merged_companies)} companies extracted")

        # Emit citation for company metrics
        await _emit_citation(
            __event_emitter__,
            "Company Metrics",
            f"**Merged Companies ({len(merged_companies)}):**\n```json\n"
            f"{json.dumps(list(merged_companies.values()), indent=2)}\n```"
        )

        # ─────────────────────────────────────────────────────────────────────
        # Step 6: Generate final report
        # ─────────────────────────────────────────────────────────────────────
        await progress.update("Step 6/6: Generating final report...")

        # Prepare structured data for final step
        structured_data = f"""## Extracted Data from Quarterly Reports

### Metadata
- **Fund Name:** {validation_result['fund_name']}
- **Currency:** {validation_result['currency']}
- **t=0 (Earlier Quarter):** {validation_result['t0_quarter']}
- **t=1 (Later Quarter):** {validation_result['t1_quarter']}
- **Analysis Date:** {datetime.now().strftime("%Y-%m-%d")}

### Fund-Level Metrics
```json
{json.dumps(results['fund_metrics'], indent=2)}
```

### Developments (from {q2_label})
```json
{json.dumps(results['developments'], indent=2)}
```

### Company-Level Metrics (Merged Q1 + Q2)
```json
{json.dumps(results['company_metrics']['companies'], indent=2)}
```
"""

        try:
            log.info("Generating final report...")
            final_response = await self._call_llm(
                system_prompt=STEP5_FINAL_REPORT_PROMPT,
                user_message=structured_data,
            )
            await progress.finish()
            return final_response

        except Exception as e:
            log.exception(f"Final report generation failed: {e}")
            await progress.error(str(e))

            # Fall back to returning the structured data
            return f"""# Multi-Step Analysis Results

The final report generation encountered an error, but here is the extracted data:

{structured_data}

**Error:** {str(e)}"""

    async def analyze_reports(
        self,
        __files__: Optional[List[Dict]] = None,
        __event_emitter__=None,
        __user__: Optional[dict] = None,
        __request__: Any = None,
        __model__: Optional[str] = None,
    ) -> str:
        """
        Extract and prepare two quarterly fund reports for analysis.

        Upload exactly 2 PDF files: the earlier quarter (t=0) and later quarter (t=1).
        Files are automatically sorted by quarter based on filename patterns.

        Supports two modes:
        - Multi-step (default): Runs 5 extraction steps for improved accuracy
        - Single-pass: Returns content for the chat model to analyze

        Returns:
            In multi-step mode: Complete formatted analysis report
            In single-pass mode: Extracted content for model to analyze
        """
        try:
            # Get PDF files with their already-extracted content
            pdf_files = _get_pdf_files_from_metadata(__files__)
            q1_name, q1_id, q1_content, q1_quarter = pdf_files[0]
            q2_name, q2_id, q2_content, q2_quarter = pdf_files[1]

            # Extract fund name
            fund_name = _extract_fund_name(q1_name)

            # Format quarter labels
            q1_label = f"Q{q1_quarter[1]} {q1_quarter[0]}" if q1_quarter != (0, 0) else "Earlier Quarter"
            q2_label = f"Q{q2_quarter[1]} {q2_quarter[0]}" if q2_quarter != (0, 0) else "Later Quarter"

            # Determine max chars based on mode
            if self.valves.ENABLE_MULTI_STEP:
                max_chars = self.valves.MAX_CHARS_PER_PDF
            else:
                max_chars = self.valves.MAX_CHARS_PER_PDF_SINGLE_PASS

            # Clean and truncate content
            q1_cleaned = _clean_content(q1_content)
            q1_text = q1_cleaned[:max_chars]
            if len(q1_cleaned) > max_chars:
                q1_text += f"\n[Truncated at {max_chars} characters]"
            log.info(f"Q1: {len(q1_content)} chars -> {len(q1_cleaned)} cleaned -> {len(q1_text)} final")

            q2_cleaned = _clean_content(q2_content)
            q2_text = q2_cleaned[:max_chars]
            if len(q2_cleaned) > max_chars:
                q2_text += f"\n[Truncated at {max_chars} characters]"
            log.info(f"Q2: {len(q2_content)} chars -> {len(q2_cleaned)} cleaned -> {len(q2_text)} final")

            # Check if multi-step mode is enabled
            if self.valves.ENABLE_MULTI_STEP:
                provider = self.valves.LLM_PROVIDER.lower()
                log.info(f"Running per-document multi-step analysis (v2) with provider: {provider}")
                return await self._run_multi_step_analysis_v2(
                    q1_text=q1_text,
                    q2_text=q2_text,
                    q1_name=q1_name,
                    q2_name=q2_name,
                    q1_label=q1_label,
                    q2_label=q2_label,
                    __event_emitter__=__event_emitter__,
                )

            # Single-pass mode: return content for model to analyze
            return f"""Analyze these two consecutive quarterly fund reports and generate the standardized performance analysis following NORD's specification.

**Fund:** {fund_name}
**Period:** t=0 ({q1_label}) vs t=1 ({q2_label})
**Analysis Date:** {datetime.now().strftime("%Y-%m-%d")}

---

## REPORT 1 (t=0 - {q1_label}):
**Source:** {q1_name}

{q1_text}

---

## REPORT 2 (t=1 - {q2_label}):
**Source:** {q2_name}

{q2_text}

---

Please analyze both reports and produce the standardized QUARTERLY FUND PERFORMANCE ANALYSIS output."""

        except Exception as e:
            log.exception(f"Error in analyze_reports: {e}")
            return f"Error analyzing reports: {str(e)}"
