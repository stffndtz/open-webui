"""
Test script for the TableAwareMarkdownSplitter.

Run with: python -m open_webui.utils.test_splitter
"""

from open_webui.utils.splitters import TableAwareMarkdownSplitter


SAMPLE_FUND_REPORT = """
# Holland Capital Fund IV - Q2 2025 Report

## Executive Summary

The fund continued to perform well in Q2 2025, with stable valuations across the portfolio.

## Fund Performance

| Metric | Q1 2025 | Q2 2025 | Δ QoQ |
|--------|---------|---------|-------|
| Gross MoIC | 1.3x | 1.3x | +0.0% |
| Net MoIC | N/A | N/A | N/A |
| Gross IRR | N/A | N/A | N/A |
| Net IRR | N/A | N/A | N/A |
| DPI | 0.0% | 0.0% | 0.0% |

Commentary: Portfolio showed stable overall valuations with minimal aggregate movement.

## Portfolio Companies

### Magnus Energy

Magnus Energy is a leading provider of energy consulting services.

**Financial Metrics:**

| Metric | Entry | LTM Q1 | LTM Q2 | Δ QoQ |
|--------|-------|--------|--------|-------|
| Sales | €20.7m | €18.6m | €17.6m | -5.4% |
| EBITDA | €3.4m | €4.1m | €3.8m | -7.3% |
| Margin | 16.1% | 21.9% | 21.3% | -0.6 pp |
| Net Debt | €6.8m | €5.6m | €5.8m | +3.6% |
| Leverage | 2.0x | 1.4x | 1.5x | |
| Multiple | 6.2x | 6.2x | 6.2x | |
| Gross MoIC | - | 2.1x | 2.1x | +0.0% |

Strong performance driven by higher consultants utilization.

### AMP Groep

AMP Groep provides industrial automation services.

**Financial Metrics:**

| Metric | Entry | LTM Q1 | LTM Q2 | Δ QoQ |
|--------|-------|--------|--------|-------|
| Sales | €9.5m | €48.9m | €47.0m | -3.9% |
| EBITDA | €1.7m | €4.1m | €2.8m | -31.7% |
| Margin | 18.4% | 8.4% | 6.0% | -2.4 pp |
| Net Debt | €2.2m | €9.2m | €8.4m | -8.7% |
| Leverage | 1.3x | 2.2x | 3.0x | |

Revenue below budget with significant EBITDA decline due to customer onboarding delays.

## Data Quality Notes

All figures reported consistently in EUR. No OCR issues detected.
"""


def test_table_aware_splitter():
    """Test the TableAwareMarkdownSplitter with sample fund report data."""
    print("=" * 60)
    print("Testing TableAwareMarkdownSplitter")
    print("=" * 60)

    splitter = TableAwareMarkdownSplitter(
        chunk_size=500,
        chunk_overlap=50,
    )

    chunks = splitter.split_text(SAMPLE_FUND_REPORT)

    print(f"\nTotal chunks created: {len(chunks)}")
    print(f"Chunks with tables: {sum(1 for c in chunks if c.metadata.get('has_table'))}")
    print(f"Chunks without tables: {sum(1 for c in chunks if not c.metadata.get('has_table'))}")

    print("\n" + "-" * 60)
    print("CHUNK DETAILS:")
    print("-" * 60)

    for i, chunk in enumerate(chunks, 1):
        has_table = chunk.metadata.get('has_table', False)
        headings = chunk.metadata.get('headings', [])
        content_preview = chunk.page_content[:150].replace('\n', ' ')

        print(f"\n[Chunk {i}] {'📊 TABLE' if has_table else '📝 TEXT'}")
        print(f"  Headings: {' > '.join(headings) if headings else '(none)'}")
        print(f"  Size: {len(chunk.page_content)} chars")
        print(f"  Preview: {content_preview}...")

    # Verify table preservation
    print("\n" + "=" * 60)
    print("VALIDATION:")
    print("=" * 60)

    table_chunks = [c for c in chunks if c.metadata.get('has_table')]
    for i, tc in enumerate(table_chunks, 1):
        # Check if table has header row and separator
        has_separator = '|---' in tc.page_content or '| ---' in tc.page_content
        has_data_rows = tc.page_content.count('|') > 10

        print(f"\nTable Chunk {i}:")
        print(f"  ✓ Has separator row: {has_separator}")
        print(f"  ✓ Has data rows: {has_data_rows}")

        # Check for section context
        headings = tc.metadata.get('headings', [])
        if headings:
            print(f"  ✓ Section context: {' > '.join(headings)}")

    print("\n" + "=" * 60)
    print("Test completed!")
    print("=" * 60)


if __name__ == "__main__":
    test_table_aware_splitter()
