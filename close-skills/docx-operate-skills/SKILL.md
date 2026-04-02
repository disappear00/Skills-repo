---
name: "docx-operate-skills"
description: "Provides DOCX document manipulation capabilities including section replacement, content clearing, hyperlink management, and heading operations. Invoke when user needs to modify Word documents programmatically."
---

# DOCX Operate Skills

A lightweight skill for manipulating Microsoft Word (DOCX) documents programmatically. Provides essential utilities for section replacement, content management, and hyperlink operations.

## Module Structure

```
docx-operate-skills/
├── SKILL.md                    # This documentation
├── docx_operations.py          # Core module with all functions
├── scripts/                    # Example usage scripts
│   ├── example_section_replacement.py
│   ├── example_content_insertion.py
│   ├── example_hyperlink_management.py
│   └── example_report_aggregation.py
└── tests/                      # Unit tests
    └── test_docx_operations.py
```

## Quick Start

```python
from docx_operations import (
    replace_section_by_title,
    clear_section_content,
    add_hyperlink_to_heading,
    get_subtitles_under_heading,
    insert_heading_after_subtitles,
)

# Replace section content
replace_section_by_title("source.docx", "target.docx", "Introduction")

# Clear section content
clear_section_content("doc.docx", "Draft Content")

# Add hyperlinks
add_hyperlink_to_heading("doc.docx", "see Appendix", "Appendix A")

# Get subtitles
subtitles = get_subtitles_under_heading("doc.docx", "Main Section")

# Insert new heading
insert_heading_after_subtitles("doc.docx", "Summary", "Results", level=2)
```

## Core Functions

### `replace_section_by_title`

Replace content under a heading in a target document with content from a source document.

**Parameters:**
- `source_file` (str): Path to the source Word document
- `target_file` (str): Path to the target Word document (modified in place)
- `title` (str): Heading text whose content should be replaced
- `title_mapping` (Dict[str, str], optional): Mapping of source title to target title
- `use_last_match` (bool, optional): Use last matching heading when duplicates exist (default: False)

**Example:**
```python
from docx_operations import replace_section_by_title

# Basic replacement
replace_section_by_title("source.docx", "target.docx", "Introduction")

# With title mapping
replace_section_by_title(
    source_file="source.docx",
    target_file="target.docx",
    title="Introduction",
    title_mapping={"Introduction": "Overview"}
)

# Replace last matched heading
replace_section_by_title(
    source_file="source.docx",
    target_file="target.docx",
    title="Conclusion",
    use_last_match=True
)
```

### `clear_section_content`

Remove all content under a specified heading while keeping the heading itself.

**Parameters:**
- `file_path` (str): Path to the DOCX file
- `title` (str): Heading text whose content should be removed
- `keep_subtitles` (List[str], optional): Subheadings whose content should be preserved

**Example:**
```python
from docx_operations import clear_section_content

# Clear all content
clear_section_content("report.docx", "Draft Content")

# Preserve specific subheadings
clear_section_content(
    file_path="report.docx",
    title="Results",
    keep_subtitles=["Important Findings", "Key Data"]
)
```

### `add_hyperlink_to_heading`

Add internal hyperlinks from text occurrences to a target heading.

**Parameters:**
- `file_path` (str): Path to the DOCX file
- `phrase` (str): Text to convert to hyperlinks
- `target_heading` (str): Target heading text to link to

**Returns:**
- `int`: Number of hyperlinks added

**Example:**
```python
from docx_operations import add_hyperlink_to_heading

count = add_hyperlink_to_heading(
    file_path="document.docx",
    phrase="see Appendix A",
    target_heading="Appendix A"
)
print(f"Added {count} hyperlinks")
```

### `get_subtitles_under_heading`

Retrieve all direct child headings under a specified parent heading.

**Parameters:**
- `file_path` (str): Path to the DOCX file
- `parent_title` (str): Parent heading text

**Returns:**
- `List[Tuple[str, str]]`: List of (original_text, normalized_text) tuples

**Example:**
```python
from docx_operations import get_subtitles_under_heading

subtitles = get_subtitles_under_heading(
    file_path="document.docx",
    parent_title="Main Section"
)
for original, normalized in subtitles:
    print(f"Original: {original}, Normalized: {normalized}")
```

### `insert_heading_after_subtitles`

Insert a new heading after all subheadings of a target heading.

**Parameters:**
- `file_path` (str): Path to the DOCX file
- `heading_text` (str): Text for the new heading
- `target_title` (str): Parent heading to insert under
- `level` (int, optional): Heading level for the new heading (default: 2)

**Example:**
```python
from docx_operations import insert_heading_after_subtitles

insert_heading_after_subtitles(
    file_path="document.docx",
    heading_text="Summary",
    target_title="Results",
    level=2
)
```

## Utility Functions

### `is_heading_rank`

Get heading level from a paragraph's style name.

```python
from docx_operations import is_heading_rank
from docx import Document

doc = Document("document.docx")
for para in doc.paragraphs:
    level = is_heading_rank(para)
    if level:
        print(f"[H{level}] {para.text}")
```

### `normalize_heading_text`

Normalize heading text by stripping leading numbering.

```python
from docx_operations import normalize_heading_text

print(normalize_heading_text("1.2.3 Introduction"))  # "Introduction"
print(normalize_heading_text("一、引言"))  # "引言"
```

## Important Notes

1. **In-place Modification**: Most functions modify documents in place. Always backup important documents before processing.

2. **Heading Style Support**: The skill supports both English ("Heading X") and Chinese ("标题 X") heading styles.

3. **Title Mapping**: Use `title_mapping` parameter to map source titles to different target titles when structure differs between documents.

4. **Heading Normalization**: Heading text matching automatically strips numbering prefixes (e.g., "1. Introduction" matches "Introduction").

## Error Handling

All functions raise `ValueError` for:
- Missing or empty required parameters
- Headings not found in documents
- Invalid file paths

## Dependencies

- `python-docx`: For DOCX file manipulation
- Standard library: `copy`, `hashlib`, `re`

## Example Scripts

See the `scripts/` directory for complete usage examples:
- `example_section_replacement.py`: Section replacement workflows
- `example_content_insertion.py`: Content management operations
- `example_hyperlink_management.py`: Hyperlink creation examples
- `example_report_aggregation.py`: Document structure analysis

## Running Tests

```bash
cd docx-operate-skills
python -m pytest tests/ -v
```
