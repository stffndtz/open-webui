"""
Custom text splitters for Open WebUI.

This module provides specialized text splitting strategies, particularly
for financial documents with tables that need to be preserved as single chunks.
"""

import re
import logging
from typing import List, Optional, Tuple
from langchain_core.documents import Document

log = logging.getLogger(__name__)


class TableAwareMarkdownSplitter:
    """
    A markdown text splitter that preserves tables as complete units.

    This splitter:
    1. Splits by markdown headers (##, ###, etc.)
    2. Detects tables within sections and keeps them intact
    3. Only applies character-based splitting to non-table content
    4. Adds section context (header hierarchy) to each chunk's metadata

    Ideal for financial documents where tables contain critical data that
    loses meaning when fragmented.
    """

    # Regex to detect markdown tables (header row + separator row + data rows)
    TABLE_PATTERN = re.compile(
        r'(\|[^\n]+\|\n'                     # Header row: | col1 | col2 |
        r'\|[-:\s|]+\|\n'                    # Separator: |---|---|
        r'(?:\|[^\n]+\|\n?)+)',              # Data rows: | val1 | val2 |
        re.MULTILINE
    )

    # Regex to detect markdown headers
    HEADER_PATTERN = re.compile(r'^(#{1,6})\s+(.+)$', re.MULTILINE)

    def __init__(
        self,
        chunk_size: int = 1000,
        chunk_overlap: int = 100,
        min_table_chunk_size: int = 200,
        max_table_chunk_size: int = 8000,
        preserve_table_context_lines: int = 3,
    ):
        """
        Initialize the table-aware splitter.

        Args:
            chunk_size: Target size for non-table content chunks
            chunk_overlap: Overlap between non-table chunks
            min_table_chunk_size: Minimum size for a table to be kept whole
            max_table_chunk_size: Maximum size for a table chunk (larger tables split by rows)
            preserve_table_context_lines: Number of lines before table to include for context
        """
        self.chunk_size = chunk_size
        self.chunk_overlap = chunk_overlap
        self.min_table_chunk_size = min_table_chunk_size
        self.max_table_chunk_size = max_table_chunk_size
        self.preserve_table_context_lines = preserve_table_context_lines

    def _extract_tables(self, text: str) -> List[Tuple[int, int, str]]:
        """
        Extract all tables from text with their positions.

        Returns:
            List of (start_pos, end_pos, table_text) tuples
        """
        tables = []
        for match in self.TABLE_PATTERN.finditer(text):
            tables.append((match.start(), match.end(), match.group(1)))
        return tables

    def _get_context_before_table(self, text: str, table_start: int) -> str:
        """
        Get context lines before a table for better understanding.
        """
        before_text = text[:table_start]
        lines = before_text.split('\n')

        # Get last N non-empty lines
        context_lines = []
        for line in reversed(lines):
            if line.strip():
                context_lines.insert(0, line)
                if len(context_lines) >= self.preserve_table_context_lines:
                    break

        return '\n'.join(context_lines)

    def _get_current_headers(self, text: str, position: int) -> List[str]:
        """
        Get the header hierarchy at a given position in the text.
        """
        text_before = text[:position]
        headers = {}

        for match in self.HEADER_PATTERN.finditer(text_before):
            level = len(match.group(1))
            header_text = match.group(2).strip()

            # Clear lower-level headers when we hit a higher-level one
            headers = {k: v for k, v in headers.items() if k < level}
            headers[level] = header_text

        # Return headers in order
        return [headers[k] for k in sorted(headers.keys())]

    def _split_prose(self, text: str, headers: List[str]) -> List[Document]:
        """
        Split non-table text using recursive character splitting.
        """
        if not text.strip():
            return []

        chunks = []

        # Simple recursive split by paragraphs, then sentences, then characters
        if len(text) <= self.chunk_size:
            chunks.append(Document(
                page_content=text.strip(),
                metadata={"headings": headers, "has_table": False}
            ))
        else:
            # Split by double newlines (paragraphs)
            paragraphs = re.split(r'\n\n+', text)
            current_chunk = ""

            for para in paragraphs:
                if len(current_chunk) + len(para) + 2 <= self.chunk_size:
                    current_chunk += ("\n\n" if current_chunk else "") + para
                else:
                    if current_chunk.strip():
                        chunks.append(Document(
                            page_content=current_chunk.strip(),
                            metadata={"headings": headers, "has_table": False}
                        ))

                    # Handle paragraphs larger than chunk_size
                    if len(para) > self.chunk_size:
                        # Split by sentences
                        sentences = re.split(r'(?<=[.!?])\s+', para)
                        current_chunk = ""
                        for sentence in sentences:
                            if len(current_chunk) + len(sentence) + 1 <= self.chunk_size:
                                current_chunk += (" " if current_chunk else "") + sentence
                            else:
                                if current_chunk.strip():
                                    chunks.append(Document(
                                        page_content=current_chunk.strip(),
                                        metadata={"headings": headers, "has_table": False}
                                    ))
                                current_chunk = sentence
                    else:
                        current_chunk = para

            if current_chunk.strip():
                chunks.append(Document(
                    page_content=current_chunk.strip(),
                    metadata={"headings": headers, "has_table": False}
                ))

        return chunks

    def _split_large_table(self, table_text: str, context: str, headers: List[str]) -> List[Document]:
        """
        Split a large table while preserving header rows.
        """
        lines = table_text.strip().split('\n')
        if len(lines) < 3:
            return [Document(
                page_content=(context + "\n\n" + table_text).strip(),
                metadata={"headings": headers, "has_table": True}
            )]

        # First two lines are header and separator
        header_lines = '\n'.join(lines[:2])
        data_lines = lines[2:]

        chunks = []
        current_rows = []
        current_size = len(context) + len(header_lines) + 4  # +4 for newlines

        for row in data_lines:
            row_size = len(row) + 1
            if current_size + row_size > self.max_table_chunk_size and current_rows:
                # Create chunk with current rows
                table_chunk = header_lines + '\n' + '\n'.join(current_rows)
                chunk_content = (context + "\n\n" + table_chunk).strip() if context else table_chunk
                chunks.append(Document(
                    page_content=chunk_content,
                    metadata={"headings": headers, "has_table": True, "table_continued": len(chunks) > 0}
                ))
                current_rows = [row]
                current_size = len(context) + len(header_lines) + row_size + 4
            else:
                current_rows.append(row)
                current_size += row_size

        # Add remaining rows
        if current_rows:
            table_chunk = header_lines + '\n' + '\n'.join(current_rows)
            chunk_content = (context + "\n\n" + table_chunk).strip() if context else table_chunk
            chunks.append(Document(
                page_content=chunk_content,
                metadata={"headings": headers, "has_table": True, "table_continued": len(chunks) > 0}
            ))

        return chunks

    def split_text(self, text: str) -> List[Document]:
        """
        Split text while preserving tables.

        Args:
            text: The markdown text to split

        Returns:
            List of Document objects with content and metadata
        """
        if not text or not text.strip():
            return []

        tables = self._extract_tables(text)
        chunks = []
        current_pos = 0

        for table_start, table_end, table_text in tables:
            # Process text before table
            if table_start > current_pos:
                prose_text = text[current_pos:table_start]
                headers = self._get_current_headers(text, table_start)
                chunks.extend(self._split_prose(prose_text, headers))

            # Get context for table
            context = self._get_context_before_table(text, table_start)
            headers = self._get_current_headers(text, table_start)

            # Handle table
            if len(table_text) > self.max_table_chunk_size:
                # Split large table by rows, preserving headers
                chunks.extend(self._split_large_table(table_text, context, headers))
            else:
                # Keep table as single chunk with context
                chunk_content = table_text.strip()
                if context and context not in chunk_content:
                    chunk_content = context + "\n\n" + chunk_content

                chunks.append(Document(
                    page_content=chunk_content,
                    metadata={"headings": headers, "has_table": True}
                ))

            current_pos = table_end

        # Process remaining text after last table
        if current_pos < len(text):
            remaining_text = text[current_pos:]
            headers = self._get_current_headers(text, len(text))
            chunks.extend(self._split_prose(remaining_text, headers))

        # If no tables were found, just split as prose
        if not tables:
            chunks = self._split_prose(text, [])

        log.info(f"TableAwareMarkdownSplitter: Split into {len(chunks)} chunks "
                 f"({sum(1 for c in chunks if c.metadata.get('has_table'))} with tables)")

        return chunks

    def split_documents(self, documents: List[Document]) -> List[Document]:
        """
        Split a list of documents while preserving tables.

        Args:
            documents: List of Document objects to split

        Returns:
            List of split Document objects
        """
        all_chunks = []

        for doc in documents:
            chunks = self.split_text(doc.page_content)

            # Preserve original document metadata
            for chunk in chunks:
                chunk.metadata = {**doc.metadata, **chunk.metadata}

            all_chunks.extend(chunks)

        return all_chunks


def create_table_aware_splitter(
    chunk_size: int = 1000,
    chunk_overlap: int = 100,
) -> TableAwareMarkdownSplitter:
    """
    Factory function to create a configured table-aware splitter.

    Args:
        chunk_size: Target size for non-table chunks
        chunk_overlap: Overlap between chunks

    Returns:
        Configured TableAwareMarkdownSplitter instance
    """
    return TableAwareMarkdownSplitter(
        chunk_size=chunk_size,
        chunk_overlap=chunk_overlap,
    )
