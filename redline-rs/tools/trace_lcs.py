#!/usr/bin/env python3
"""
Trace LCS algorithm for debugging redline comparison.

Extracts text from two DOCX paragraphs and traces through the LCS algorithm
step-by-step to understand how edits are detected and grouped.

Usage:
    # Compare same section from two documents
    ./trace_lcs.py doc1.docx doc2.docx "section 3.1"

    # Compare with custom text
    ./trace_lcs.py --text "the quick brown fox" "the slow brown dog"

    # Show full LCS matrix
    ./trace_lcs.py --matrix doc1.docx doc2.docx "section 3.1"

    # Character-level diff (instead of word-level)
    ./trace_lcs.py --chars doc1.docx doc2.docx "section 3.1"
"""

import argparse
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from dataclasses import dataclass
from enum import Enum
from typing import List, Optional, Tuple

# OOXML namespace
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'


class EditOp(Enum):
    EQUAL = "="
    DELETE = "-"
    INSERT = "+"


@dataclass
class Edit:
    op: EditOp
    tokens: List[str]
    pos1: Optional[int] = None  # Position in sequence 1
    pos2: Optional[int] = None  # Position in sequence 2


def extract_xml_from_docx(docx_path: str, xml_file: str = 'word/document.xml') -> Optional[bytes]:
    """Extract an XML file from a DOCX archive."""
    try:
        with zipfile.ZipFile(docx_path, 'r') as zf:
            if xml_file in zf.namelist():
                return zf.read(xml_file)
    except (zipfile.BadZipFile, FileNotFoundError) as e:
        print(f"Error: {e}", file=sys.stderr)
    return None


def get_paragraph_text(para: ET.Element) -> str:
    """Extract all text content from a paragraph element."""
    texts = []
    for t in para.iter(f'{W}t'):
        if t.text:
            texts.append(t.text)
    return ''.join(texts)


def find_section_text(docx_path: str, section_id: str) -> Optional[str]:
    """Find a section and return its text content."""
    xml_content = extract_xml_from_docx(docx_path)
    if xml_content is None:
        return None

    root = ET.fromstring(xml_content)
    paragraphs = list(root.iter(f'{W}p'))

    result_texts = []
    found = False
    section_pattern = re.compile(
        r'^[\s\u00A0]*'
        r'(\d+(?:\.\d+)*\.?|\([a-zA-Z]\)|\([ivxlcdm]+\)|[a-zA-Z]\.|[ivxlcdmIVXLCDM]+\.)'
        r'[\s\u00A0]+'
    )

    for para in paragraphs:
        text = get_paragraph_text(para)

        if not found:
            text_normalized = text.lstrip().lower()
            section_normalized = section_id.lower()
            if text_normalized.startswith(section_normalized):
                rest = text_normalized[len(section_normalized):]
                if not rest or rest[0] in ' \t\n\u00a0':
                    found = True
                    result_texts.append(text)
                    continue
        else:
            if section_pattern.match(text):
                break
            result_texts.append(text)

    return ' '.join(result_texts) if result_texts else None


def tokenize_words(text: str) -> List[str]:
    """
    Tokenize text into words, preserving whitespace as separate tokens.
    This mimics the ComparisonUnitWord approach.
    """
    tokens = []
    current = ""
    for char in text:
        if char in ' \t\n\r\u00a0':
            if current:
                tokens.append(current)
                current = ""
            tokens.append(char)
        else:
            current += char
    if current:
        tokens.append(current)
    return tokens


def tokenize_chars(text: str) -> List[str]:
    """Tokenize text into individual characters."""
    return list(text)


def compute_lcs_matrix(seq1: List[str], seq2: List[str]) -> List[List[int]]:
    """
    Compute the LCS length matrix.

    matrix[i][j] = length of LCS of seq1[:i] and seq2[:j]
    """
    m, n = len(seq1), len(seq2)
    matrix = [[0] * (n + 1) for _ in range(m + 1)]

    for i in range(1, m + 1):
        for j in range(1, n + 1):
            if seq1[i - 1] == seq2[j - 1]:
                matrix[i][j] = matrix[i - 1][j - 1] + 1
            else:
                matrix[i][j] = max(matrix[i - 1][j], matrix[i][j - 1])

    return matrix


def backtrack_lcs(
    matrix: List[List[int]],
    seq1: List[str],
    seq2: List[str],
    trace: bool = False
) -> List[Edit]:
    """
    Backtrack through the LCS matrix to produce an edit script.

    Returns a list of Edit operations.
    """
    edits = []
    i, j = len(seq1), len(seq2)

    if trace:
        print("\n=== BACKTRACK TRACE ===")
        print(f"Starting at matrix[{i}][{j}] = {matrix[i][j]}")

    while i > 0 or j > 0:
        if i > 0 and j > 0 and seq1[i - 1] == seq2[j - 1]:
            # Match - diagonal move
            if trace:
                print(f"  [{i},{j}] MATCH: {repr(seq1[i-1])}")
            edits.append(Edit(EditOp.EQUAL, [seq1[i - 1]], i - 1, j - 1))
            i -= 1
            j -= 1
        elif j > 0 and (i == 0 or matrix[i][j - 1] >= matrix[i - 1][j]):
            # Insert - move left
            if trace:
                print(f"  [{i},{j}] INSERT: {repr(seq2[j-1])}")
            edits.append(Edit(EditOp.INSERT, [seq2[j - 1]], None, j - 1))
            j -= 1
        elif i > 0:
            # Delete - move up
            if trace:
                print(f"  [{i},{j}] DELETE: {repr(seq1[i-1])}")
            edits.append(Edit(EditOp.DELETE, [seq1[i - 1]], i - 1, None))
            i -= 1

    edits.reverse()
    return edits


def coalesce_edits(edits: List[Edit]) -> List[Edit]:
    """
    Coalesce consecutive edits of the same type into groups.
    This mimics how revisions get grouped in the output.
    """
    if not edits:
        return []

    coalesced = []
    current = Edit(edits[0].op, list(edits[0].tokens), edits[0].pos1, edits[0].pos2)

    for edit in edits[1:]:
        if edit.op == current.op:
            # Same operation - extend current group
            current.tokens.extend(edit.tokens)
        else:
            # Different operation - start new group
            coalesced.append(current)
            current = Edit(edit.op, list(edit.tokens), edit.pos1, edit.pos2)

    coalesced.append(current)
    return coalesced


def print_matrix(matrix: List[List[int]], seq1: List[str], seq2: List[str]):
    """Print the LCS matrix in a readable format."""
    print("\n=== LCS MATRIX ===")

    # Truncate tokens for display
    def trunc(s, n=8):
        s = repr(s)[1:-1]  # Remove quotes
        return s[:n] if len(s) <= n else s[:n-2] + ".."

    # Header row
    header = "        " + " ".join(f"{trunc(t):>8}" for t in seq2)
    print(header)
    print("    " + "-" * (len(header) - 4))

    # Matrix rows
    for i, row in enumerate(matrix):
        if i == 0:
            label = "    "
        else:
            label = f"{trunc(seq1[i-1]):>4}"
        values = " ".join(f"{v:>8}" for v in row)
        print(f"{label}|{values}")


def print_edit_script(edits: List[Edit], title: str = "EDIT SCRIPT"):
    """Print the edit script in a readable format."""
    print(f"\n=== {title} ===")
    for edit in edits:
        op_char = edit.op.value
        text = ''.join(edit.tokens)
        # Escape special characters for display
        text_display = repr(text)[1:-1]
        print(f"  {op_char} {text_display}")


def print_coalescing_analysis(raw_edits: List[Edit], coalesced: List[Edit]):
    """Analyze how edits get coalesced into revision groups."""
    print("\n=== COALESCING ANALYSIS ===")
    print(f"Raw operations: {len(raw_edits)}")
    print(f"Coalesced groups: {len(coalesced)}")

    print("\nRevision groups that would be created:")
    for i, edit in enumerate(coalesced):
        text = ''.join(edit.tokens)
        text_display = repr(text)[1:-1]
        if edit.op == EditOp.DELETE:
            print(f"  <w:del> #{i}: {text_display}")
        elif edit.op == EditOp.INSERT:
            print(f"  <w:ins> #{i}: {text_display}")
        else:
            print(f"  (equal) #{i}: {text_display[:50]}{'...' if len(text_display) > 50 else ''}")

    # Count transitions
    transitions = 0
    for i in range(1, len(coalesced)):
        if coalesced[i].op != coalesced[i-1].op:
            transitions += 1
    print(f"\nOperation transitions: {transitions}")


def print_side_by_side(seq1: List[str], seq2: List[str], edits: List[Edit]):
    """Print side-by-side comparison showing alignment."""
    print("\n=== SIDE-BY-SIDE ALIGNMENT ===")
    print(f"{'Original':<40} | {'Modified':<40} | Op")
    print("-" * 85)

    for edit in edits:
        text = ''.join(edit.tokens)
        text_display = repr(text)[1:-1][:35]

        if edit.op == EditOp.EQUAL:
            print(f"{text_display:<40} | {text_display:<40} | =")
        elif edit.op == EditOp.DELETE:
            print(f"{text_display:<40} | {'':40} | -")
        elif edit.op == EditOp.INSERT:
            print(f"{'':40} | {text_display:<40} | +")


def main():
    parser = argparse.ArgumentParser(
        description='Trace LCS algorithm for debugging redline comparison',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__
    )
    parser.add_argument('input1', help='First DOCX file or text (with --text)')
    parser.add_argument('input2', help='Second DOCX file or text (with --text)')
    parser.add_argument('query', nargs='?', help='Section query (e.g., "section 3.1")')
    parser.add_argument('--text', action='store_true', help='Treat inputs as raw text instead of files')
    parser.add_argument('--matrix', action='store_true', help='Show full LCS matrix')
    parser.add_argument('--chars', action='store_true', help='Character-level diff instead of word-level')
    parser.add_argument('--trace', action='store_true', help='Show backtrack trace')
    parser.add_argument('--no-coalesce', action='store_true', help='Show raw edits without coalescing')

    args = parser.parse_args()

    # Get input texts
    if args.text:
        text1, text2 = args.input1, args.input2
    else:
        if not args.query:
            parser.error("Query is required when comparing DOCX files")

        # Parse query
        query = args.query
        if query.lower().startswith('section '):
            section_id = query[8:].strip()
        else:
            section_id = query

        text1 = find_section_text(args.input1, section_id)
        text2 = find_section_text(args.input2, section_id)

        if text1 is None:
            print(f"Error: Could not find section '{section_id}' in {args.input1}", file=sys.stderr)
            sys.exit(1)
        if text2 is None:
            print(f"Error: Could not find section '{section_id}' in {args.input2}", file=sys.stderr)
            sys.exit(1)

    # Tokenize
    if args.chars:
        tokens1 = tokenize_chars(text1)
        tokens2 = tokenize_chars(text2)
        print(f"Character-level comparison: {len(tokens1)} vs {len(tokens2)} chars")
    else:
        tokens1 = tokenize_words(text1)
        tokens2 = tokenize_words(text2)
        print(f"Word-level comparison: {len(tokens1)} vs {len(tokens2)} tokens")

    print(f"\n=== INPUT ===")
    print(f"Text 1: {repr(text1[:100])}{'...' if len(text1) > 100 else ''}")
    print(f"Text 2: {repr(text2[:100])}{'...' if len(text2) > 100 else ''}")

    # Compute LCS matrix
    matrix = compute_lcs_matrix(tokens1, tokens2)
    lcs_length = matrix[len(tokens1)][len(tokens2)]
    print(f"\nLCS length: {lcs_length}")

    if args.matrix and len(tokens1) <= 20 and len(tokens2) <= 20:
        print_matrix(matrix, tokens1, tokens2)
    elif args.matrix:
        print(f"\n(Matrix too large to display: {len(tokens1)+1}x{len(tokens2)+1})")

    # Backtrack to get edit script
    raw_edits = backtrack_lcs(matrix, tokens1, tokens2, trace=args.trace)

    if args.no_coalesce:
        print_edit_script(raw_edits, "RAW EDIT SCRIPT")
    else:
        coalesced = coalesce_edits(raw_edits)
        print_edit_script(coalesced, "COALESCED EDIT SCRIPT")
        print_coalescing_analysis(raw_edits, coalesced)

    # Side-by-side view for coalesced
    if not args.no_coalesce:
        coalesced = coalesce_edits(raw_edits)
        if len(coalesced) <= 30:
            print_side_by_side(tokens1, tokens2, coalesced)


if __name__ == '__main__':
    main()
