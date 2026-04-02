"""
Example: Section Replacement Workflows

This script demonstrates how to use the section replacement function
to update content in Word documents.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from docx_operations import replace_section_by_title


def example_basic_replacement():
    """
    Basic example: Replace content under a heading in target document
    with content from source document.
    """
    source_file = "source_document.docx"
    target_file = "target_document.docx"
    title = "Introduction"

    try:
        replace_section_by_title(
            source_file=source_file,
            target_file=target_file,
            title=title,
        )
        print(f"Successfully replaced content under '{title}'")
    except ValueError as e:
        print(f"Error: {e}")


def example_replacement_with_title_mapping():
    """
    Example with title mapping: Use when source and target documents
    have different heading names for the same content.
    """
    source_file = "source_document.docx"
    target_file = "target_document.docx"

    title_mapping = {
        "Organ Weights": "Organ Weights and Coefficients",
        "Hematology": "Blood Analysis",
        "Clinical Chemistry": "Biochemical Markers",
    }

    for source_title, target_title in title_mapping.items():
        try:
            replace_section_by_title(
                source_file=source_file,
                target_file=target_file,
                title=source_title,
                title_mapping={source_title: target_title},
            )
            print(f"Replaced '{source_title}' -> '{target_title}'")
        except ValueError as e:
            print(f"Skipped '{source_title}': {e}")


def example_replace_last_matched_heading():
    """
    Use use_last_match=True when the target document has multiple
    headings with the same text and you want to replace the LAST one.
    """
    source_file = "updated_content.docx"
    target_file = "report.docx"
    title = "Conclusion"

    try:
        replace_section_by_title(
            source_file=source_file,
            target_file=target_file,
            title=title,
            use_last_match=True,
        )
        print(f"Replaced content under the last '{title}' heading")
    except ValueError as e:
        print(f"Error: {e}")


def example_batch_replacement():
    """
    Batch replacement: Update multiple sections in a single workflow.
    """
    source_file = "new_data.docx"
    target_file = "report.docx"

    sections_to_replace = [
        "Materials and Methods",
        "Results",
        "Discussion",
        "Conclusion",
    ]

    for section in sections_to_replace:
        try:
            replace_section_by_title(
                source_file=source_file,
                target_file=target_file,
                title=section,
            )
            print(f"[OK] Replaced: {section}")
        except ValueError as e:
            print(f"[SKIP] {section}: {e}")


if __name__ == "__main__":
    print("=" * 60)
    print("Section Replacement Examples")
    print("=" * 60)

    print("\n1. Basic Replacement:")
    print("   See example_basic_replacement()")

    print("\n2. Replacement with Title Mapping:")
    print("   See example_replacement_with_title_mapping()")

    print("\n3. Replace Last Matched Heading:")
    print("   See example_replace_last_matched_heading()")

    print("\n4. Batch Replacement:")
    print("   See example_batch_replacement()")

    print("\n" + "=" * 60)
    print("Note: Ensure document files exist before running examples.")
    print("=" * 60)
