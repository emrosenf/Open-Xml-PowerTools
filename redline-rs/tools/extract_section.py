#!/usr/bin/env python3
"""
Extract sections from DOCX document.xml for debugging and comparison.

Usage:
    # Extract a numbered section (until next numbered section)
    ./extract_section.py doc.docx "section 3.1"
    ./extract_section.py doc.docx "section (b)"

    # Extract paragraph(s) starting with specific text
    ./extract_section.py doc.docx "paragraph 'The quick brown fox'"
    ./extract_section.py doc.docx "para 'Rent Commencement'"

    # Extract a footnote by number
    ./extract_section.py doc.docx "footnote 5"

    # Extract an endnote by number
    ./extract_section.py doc.docx "endnote 3"

    # Extract raw element containing text (single element, no continuation)
    ./extract_section.py doc.docx "element 'some unique text'"

    # Compare two docx files at a section
    ./extract_section.py gold.docx rust.docx "section 3.1"
    ./extract_section.py --diff gold.docx rust.docx "section 3.1"
"""

import argparse
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Optional, List
from xml.dom import minidom

# OOXML namespaces
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
}

# Register namespaces to preserve prefixes in output
for prefix, uri in NAMESPACES.items():
    ET.register_namespace(prefix, uri)

# Namespace prefix for finding elements
W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'

# Pattern to detect section/subsection numbering
SECTION_PATTERN = re.compile(
    r'^[\s\u00A0]*'  # optional leading whitespace (including NBSP)
    r'('
    r'\d+(?:\.\d+)*\.?'      # 1, 1.1, 1.1.1, etc.
    r'|'
    r'\([a-zA-Z]\)'          # (a), (b), (A), (B)
    r'|'
    r'\([ivxlcdm]+\)'        # (i), (ii), (iii), (iv) - roman numerals
    r'|'
    r'[a-zA-Z]\.'            # a., b., A., B.
    r'|'
    r'[ivxlcdmIVXLCDM]+\.'   # i., ii., iii., I., II.
    r')'
    r'[\s\u00A0]+'           # required whitespace after number
)


def extract_xml_from_docx(docx_path: str, xml_file: str = 'word/document.xml') -> Optional[bytes]:
    """Extract an XML file from a DOCX archive."""
    try:
        with zipfile.ZipFile(docx_path, 'r') as zf:
            if xml_file in zf.namelist():
                return zf.read(xml_file)
            # Try without 'word/' prefix for some files
            alt_path = xml_file.replace('word/', '')
            if alt_path in zf.namelist():
                return zf.read(alt_path)
    except zipfile.BadZipFile:
        print(f"Error: {docx_path} is not a valid DOCX file", file=sys.stderr)
    except FileNotFoundError:
        print(f"Error: File not found: {docx_path}", file=sys.stderr)
    return None


def get_paragraph_text(para: ET.Element) -> str:
    """Extract all text content from a paragraph element."""
    texts = []
    for t in para.iter(f'{W}t'):
        if t.text:
            texts.append(t.text)
    return ''.join(texts)


def get_element_text(elem: ET.Element) -> str:
    """Extract all text content from any element."""
    texts = []
    for t in elem.iter(f'{W}t'):
        if t.text:
            texts.append(t.text)
    return ''.join(texts)


def is_section_start(text: str) -> bool:
    """Check if text starts with a section/subsection number."""
    return bool(SECTION_PATTERN.match(text))


def get_section_level(text: str) -> tuple:
    """
    Get the hierarchical level and type of a section number.
    Returns (level, type) where type is 'numeric', 'alpha', 'roman'.
    """
    match = SECTION_PATTERN.match(text)
    if not match:
        return (0, '')

    marker = match.group(1).strip()

    # Numeric: 1, 1.1, 1.1.1
    if re.match(r'^\d+(?:\.\d+)*\.?$', marker):
        level = marker.rstrip('.').count('.') + 1
        return (level, 'numeric')

    # Parenthetical alpha: (a), (b)
    if re.match(r'^\([a-zA-Z]\)$', marker):
        return (10, 'alpha_paren')  # High level = subsection

    # Parenthetical roman: (i), (ii)
    if re.match(r'^\([ivxlcdm]+\)$', marker, re.IGNORECASE):
        return (11, 'roman_paren')

    # Dotted alpha: a., b.
    if re.match(r'^[a-zA-Z]\.$', marker):
        return (10, 'alpha_dot')

    # Dotted roman: i., ii.
    if re.match(r'^[ivxlcdmIVXLCDM]+\.$', marker):
        return (11, 'roman_dot')

    return (0, '')


def find_paragraphs(root: ET.Element) -> List[ET.Element]:
    """Find all paragraph elements in the document."""
    return list(root.iter(f'{W}p'))


def find_section(root: ET.Element, section_id: str) -> List[ET.Element]:
    """
    Find a section by its number (e.g., "3.1", "(b)") and return all elements
    until the next section at the same or higher level.
    """
    paragraphs = find_paragraphs(root)

    result = []
    found = False
    target_level = None
    target_type = None

    for para in paragraphs:
        text = get_paragraph_text(para)

        if not found:
            # Look for the section start
            # Normalize both for comparison
            text_normalized = text.lstrip().lower()
            section_normalized = section_id.lower()

            # Check if this paragraph starts with our section ID
            if text_normalized.startswith(section_normalized):
                # Verify it's followed by whitespace or end
                rest = text_normalized[len(section_normalized):]
                if not rest or rest[0] in ' \t\n\u00a0':
                    found = True
                    target_level, target_type = get_section_level(text)
                    result.append(para)
                    continue
        else:
            # Check if we've hit the next section
            if is_section_start(text):
                level, stype = get_section_level(text)
                # Stop if we hit same level or higher (lower number)
                # For same type, compare levels; for different types, be conservative
                if stype == target_type and level <= target_level:
                    break
                # If different type but looks like a major section, stop
                if stype == 'numeric' and level <= 2:
                    break

            result.append(para)

    return result


def find_paragraph_by_text(root: ET.Element, search_text: str) -> List[ET.Element]:
    """
    Find paragraph(s) that begin with the specified text.
    Returns the matching paragraph and continues until next section.
    """
    paragraphs = find_paragraphs(root)

    result = []
    found = False
    search_lower = search_text.lower()

    for para in paragraphs:
        text = get_paragraph_text(para)

        if not found:
            if text.lower().lstrip().startswith(search_lower):
                found = True
                result.append(para)
                continue
        else:
            # Stop at next section
            if is_section_start(text):
                break
            result.append(para)

    return result


def find_element_by_text(root: ET.Element, search_text: str) -> List[ET.Element]:
    """
    Find the most specific (deepest) element containing the specified text.
    Returns just that element (no continuation).
    """
    search_lower = search_text.lower()

    # Find all matching elements
    matches = []
    for elem in root.iter():
        text = get_element_text(elem)
        if search_lower in text.lower():
            matches.append(elem)

    if not matches:
        return []

    # Return the deepest (last in document order that matches)
    # since iter() goes parent -> children, the last match is typically most specific
    # But we want the smallest element, so find one with fewest descendants
    best = min(matches, key=lambda e: len(list(e.iter())))
    return [best]


def find_footnote(root: ET.Element, note_id: str) -> List[ET.Element]:
    """Find a footnote by its ID number."""
    for footnote in root.iter(f'{W}footnote'):
        fid = footnote.get(f'{W}id')
        if fid == note_id:
            return [footnote]
    return []


def find_endnote(root: ET.Element, note_id: str) -> List[ET.Element]:
    """Find an endnote by its ID number."""
    for endnote in root.iter(f'{W}endnote'):
        eid = endnote.get(f'{W}id')
        if eid == note_id:
            return [endnote]
    return []


def pretty_print_xml(elements: List[ET.Element]) -> str:
    """Pretty print a list of XML elements."""
    if not elements:
        return ""

    # Create a wrapper to hold multiple elements
    wrapper = ET.Element("extracted")
    for elem in elements:
        # Deep copy
        wrapper.append(elem)

    # Convert to string and pretty print
    rough_string = ET.tostring(wrapper, encoding='unicode')

    # Use minidom for prettier output
    try:
        dom = minidom.parseString(rough_string.encode('utf-8'))
        pretty = dom.toprettyxml(indent="  ")
    except Exception:
        # Fallback if minidom fails
        return rough_string

    # Remove the XML declaration and wrapper
    lines = pretty.split('\n')
    result_lines = []
    in_wrapper = False
    for line in lines:
        if '<?xml' in line:
            continue
        if '<extracted' in line and '>' in line:
            in_wrapper = True
            continue
        if '</extracted>' in line:
            break
        if in_wrapper and line.strip():
            # Remove one level of indentation (from wrapper)
            if line.startswith('  '):
                result_lines.append(line[2:])
            else:
                result_lines.append(line)

    return '\n'.join(result_lines)


def parse_query(query: str) -> tuple:
    """
    Parse the search query into (type, value).

    Examples:
        "section 3.1" -> ("section", "3.1")
        "paragraph 'The quick'" -> ("paragraph", "The quick")
        "footnote 5" -> ("footnote", "5")
        "element 'some text'" -> ("element", "some text")
    """
    query = query.strip()

    # Section query
    match = re.match(r'^section\s+(.+)$', query, re.IGNORECASE)
    if match:
        return ('section', match.group(1).strip())

    # Paragraph query (with quoted text)
    match = re.match(r'^para(?:graph)?\s+[\'"](.+)[\'"]$', query, re.IGNORECASE)
    if match:
        return ('paragraph', match.group(1))

    # Footnote query
    match = re.match(r'^footnote\s+(\d+)$', query, re.IGNORECASE)
    if match:
        return ('footnote', match.group(1))

    # Endnote query
    match = re.match(r'^endnote\s+(\d+)$', query, re.IGNORECASE)
    if match:
        return ('endnote', match.group(1))

    # Element query (with quoted text)
    match = re.match(r'^element\s+[\'"](.+)[\'"]$', query, re.IGNORECASE)
    if match:
        return ('element', match.group(1))

    # Default: try as section
    return ('section', query)


def extract_from_docx(docx_path: str, query: str) -> Optional[str]:
    """Main extraction function."""
    query_type, query_value = parse_query(query)

    # Determine which XML file to use
    if query_type == 'footnote':
        xml_file = 'word/footnotes.xml'
    elif query_type == 'endnote':
        xml_file = 'word/endnotes.xml'
    else:
        xml_file = 'word/document.xml'

    xml_content = extract_xml_from_docx(docx_path, xml_file)
    if xml_content is None:
        return None

    root = ET.fromstring(xml_content)

    # Find the elements based on query type
    if query_type == 'section':
        elements = find_section(root, query_value)
    elif query_type == 'paragraph':
        elements = find_paragraph_by_text(root, query_value)
    elif query_type == 'footnote':
        elements = find_footnote(root, query_value)
    elif query_type == 'endnote':
        elements = find_endnote(root, query_value)
    elif query_type == 'element':
        elements = find_element_by_text(root, query_value)
    else:
        print(f"Unknown query type: {query_type}", file=sys.stderr)
        return None

    if not elements:
        print(f"No match found for: {query}", file=sys.stderr)
        return None

    return pretty_print_xml(elements)


def main():
    parser = argparse.ArgumentParser(
        description='Extract sections from DOCX files for comparison',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__
    )
    parser.add_argument('docx_file', help='Path to DOCX file (or first file for comparison)')
    parser.add_argument('second', nargs='?', help='Query string OR second DOCX file for comparison')
    parser.add_argument('query', nargs='?', help='Query string (when comparing two files)')
    parser.add_argument('-o', '--output', help='Output file (default: stdout)')
    parser.add_argument('--diff', action='store_true', help='Show unified diff between two files')

    args = parser.parse_args()

    # Determine if we're comparing two files or extracting from one
    if args.query:
        # Two-file comparison mode
        docx1 = args.docx_file
        docx2 = args.second
        query = args.query

        result1 = extract_from_docx(docx1, query)
        result2 = extract_from_docx(docx2, query)

        if result1 is None or result2 is None:
            sys.exit(1)

        if args.diff:
            import difflib
            diff = difflib.unified_diff(
                result1.splitlines(keepends=True),
                result2.splitlines(keepends=True),
                fromfile=docx1,
                tofile=docx2
            )
            output = ''.join(diff)
        else:
            output = f"=== {docx1} ===\n{result1}\n\n=== {docx2} ===\n{result2}"
    else:
        # Single file extraction mode
        query = args.second
        if not query:
            parser.error("Query string is required")

        result = extract_from_docx(args.docx_file, query)
        if result is None:
            sys.exit(1)
        output = result

    # Output
    if args.output:
        Path(args.output).write_text(output)
        print(f"Written to {args.output}")
    else:
        print(output)


if __name__ == '__main__':
    main()
