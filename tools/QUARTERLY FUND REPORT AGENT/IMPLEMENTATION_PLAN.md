# Fund Report Tool - Per-Document Extraction Implementation Plan

## Problem Statement

The current multi-step extraction sends BOTH documents to each extraction step (Steps 1-4), which can cause token limit issues with large PDFs. The solution is to process each document individually, then combine the results for the final report.

## Proposed Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                    CURRENT FLOW (problematic)                   │
├─────────────────────────────────────────────────────────────────┤
│  Step 1: Validation      → Both docs (Q1+Q2)                   │
│  Step 2: Fund Metrics    → Both docs (Q1+Q2)                   │
│  Step 3: Developments    → Both docs (Q1+Q2)                   │
│  Step 4: Company Metrics → Both docs (Q1+Q2)                   │
│  Step 5: Final Report    → Combined JSON                       │
└─────────────────────────────────────────────────────────────────┘

┌─────────────────────────────────────────────────────────────────┐
│                    NEW FLOW (per-document)                      │
├─────────────────────────────────────────────────────────────────┤
│  Step 1: Metadata Q1     → Q1 only → { fund, quarter, cos }    │
│  Step 2: Metadata Q2     → Q2 only → { fund, quarter, cos }    │
│  Step 3: Fund Metrics Q1 → Q1 only → { moic, irr, dpi }        │
│  Step 4: Fund Metrics Q2 → Q2 only → { moic, irr, dpi }        │
│  Step 5: Developments    → Q2 only → { new, exits } (changes)  │
│  Step 6: Companies Q1    → Q1 only → [ company metrics... ]    │
│  Step 7: Companies Q2    → Q2 only → [ company metrics... ]    │
│  Step 8: Final Report    → All JSON combined                   │
└─────────────────────────────────────────────────────────────────┘
```

## Key Features

### 1. ExpandableStatusIndicator Integration
- Real-time progress log using collapsible `<details type="status">` block
- Shows step-by-step progress with bullet points
- Displays timing information
- Collapses when done

### 2. Citation Events for Extracted Data
- Emit extracted JSON as citations so user can inspect raw data
- Use `{"type": "citation", "data": {...}}` events
- Group by extraction step (metadata, fund metrics, companies)

### 3. Per-Document Processing
- Each document processed independently to avoid token limits
- Results merged intelligently for final report
- Validation happens after both metadata extractions

## Implementation TODOs

### Phase 1: Status Indicator ✅
- [x] Create `ExpandableStatusIndicator` class (adapted from manifold_pipe.py)
- [x] Replace simple `_emit_status()` with status indicator pattern
- [x] Add timing information to each step

### Phase 2: Refactor Extraction Steps ✅
- [x] Create `PROMPT_EXTRACT_METADATA` - single document metadata extraction
- [x] Create `PROMPT_EXTRACT_FUND_METRICS` - single document fund metrics
- [x] Create `PROMPT_EXTRACT_DEVELOPMENTS` - Q2 only (changes happen in later quarter)
- [x] Create `PROMPT_EXTRACT_COMPANY_METRICS` - single document company metrics
- [x] Reuse `STEP5_FINAL_REPORT_PROMPT` (already compatible with new structure)

### Phase 3: New Pipeline Flow ✅
- [x] Implement `_run_multi_step_analysis_v2()` with per-document extraction
- [x] Add parallel execution where possible (Q1 and Q2 metadata simultaneously)
- [x] Add validation logic to compare Q1/Q2 metadata (same fund, consecutive quarters)
- [x] Merge company metrics from Q1 and Q2 into unified company objects

### Phase 4: Citations for Extracted Data ✅
- [x] Emit citation for metadata extraction (combined Q1/Q2)
- [x] Emit citation for fund metrics (combined Q1/Q2)
- [x] Emit citation for developments
- [x] Emit citation for company metrics (merged)

### Phase 5: Error Handling & Fallbacks ✅
- [x] Handle partial failures (one document extraction fails)
- [x] Provide meaningful error messages in status indicator
- [x] Fall back to raw data display if final report generation fails

## Status Indicator Design

```
▾ Generating Fund Report
  - **Extracting Q1 Metadata**
    - Fund: Holland Capital Fund IV
    - Quarter: Q1 2025
  - **Extracting Q2 Metadata**
    - Fund: Holland Capital Fund IV
    - Quarter: Q2 2025
  - **Validating Reports**
    - ✓ Same fund
    - ✓ Consecutive quarters
  - **Extracting Fund Metrics**
    - Q1: Gross MoIC 1.3x, IRR N/A
    - Q2: Gross MoIC 1.3x, IRR N/A
  - **Extracting Developments**
    - 1 exit (Magnus Black)
  - **Extracting Company Metrics (Q1)**
    - Magnus Energy, AMP Groep
  - **Extracting Company Metrics (Q2)**
    - Magnus Energy, AMP Groep
  - **Generating Final Report**
  - Finished in 45.2s
```

## New Prompts Design

### STEP_METADATA_PROMPT (per document)
Extract from ONE document:
- Fund name
- Quarter/Year
- Currency
- List of portfolio companies mentioned

### STEP_FUND_METRICS_PROMPT (per document)
Extract from ONE document:
- Gross/Net MoIC
- Gross/Net IRR
- DPI
- Any notes about data availability

### STEP_DEVELOPMENTS_PROMPT (Q2 only)
Extract from Q2 document:
- New platform investments
- Add-on acquisitions
- Exits/distributions
(Developments are typically reported in the quarter they occur)

### STEP_COMPANY_METRICS_PROMPT (per document)
Extract from ONE document for each company:
- Sales (LTM)
- EBITDA (LTM)
- Margin
- Net Debt
- Leverage
- Multiple
- Gross MoIC
- Commentary/quotes

### STEP_FINAL_REPORT_PROMPT
Receives structured JSON with:
```json
{
  "metadata": {
    "fund_name": "...",
    "t0_quarter": "Q1 2025",
    "t1_quarter": "Q2 2025",
    "currency": "EUR"
  },
  "fund_metrics": {
    "t0": { ... },
    "t1": { ... }
  },
  "developments": { ... },
  "companies": [
    {
      "name": "Company A",
      "entry": { ... },
      "t0": { ... },
      "t1": { ... },
      "commentary": { ... }
    }
  ]
}
```

## Estimated Token Usage Comparison

### Current (both docs per step)
- Step 1-4: ~25K tokens input each × 4 = 100K tokens
- Step 5: ~5K tokens
- **Total: ~105K tokens per analysis**

### New (per-document)
- Metadata: ~12K × 2 = 24K tokens
- Fund Metrics: ~12K × 2 = 24K tokens
- Developments: ~12K × 1 = 12K tokens
- Companies: ~12K × 2 = 24K tokens
- Final: ~10K tokens
- **Total: ~94K tokens per analysis**

Slightly more efficient, but more importantly: **no single call exceeds 15K input tokens**.

## Questions for Approval

1. Should developments be extracted from Q2 only, or should we compare Q1 and Q2 to find changes?
2. For the status indicator - should we use the full ExpandableStatusIndicator class, or a simplified version?
3. Should citations show the full JSON or summarized data?
4. Should we support parallel extraction (Q1 + Q2 metadata at same time) to reduce latency?
