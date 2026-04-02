"""
Example: Hyperlink Management Workflows

This script demonstrates how to create and manage internal hyperlinks
in Word documents, linking text occurrences to target headings.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from docx_operations import add_hyperlink_to_heading


def example_basic_hyperlink():
    """
    Basic example: Add hyperlinks from text occurrences to a target heading.
    """
    file_path = "document.docx"
    phrase = "see Appendix A"
    target_heading = "Appendix A"

    try:
        count = add_hyperlink_to_heading(
            file_path=file_path,
            phrase=phrase,
            target_heading=target_heading,
        )
        print(f"Added {count} hyperlink(s) for '{phrase}' -> '{target_heading}'")
    except ValueError as e:
        print(f"Error: {e}")


def example_multiple_hyperlinks():
    """
    Add multiple hyperlinks to different sections in a document.
    """
    file_path = "report.docx"

    hyperlink_mappings = [
        ("see Methods", "Methods"),
        ("see Results", "Results"),
        ("see Discussion", "Discussion"),
        ("see Appendix A", "Appendix A"),
        ("see Appendix B", "Appendix B"),
        ("see References", "References"),
    ]

    total_links = 0
    for phrase, target_heading in hyperlink_mappings:
        try:
            count = add_hyperlink_to_heading(
                file_path=file_path,
                phrase=phrase,
                target_heading=target_heading,
            )
            total_links += count
            print(f"[OK] '{phrase}' -> '{target_heading}': {count} link(s)")
        except ValueError as e:
            print(f"[SKIP] '{phrase}' -> '{target_heading}': {e}")

    print(f"\nTotal hyperlinks added: {total_links}")


def example_cross_reference_hyperlinks():
    """
    Create cross-reference style hyperlinks throughout a document.
    """
    file_path = "technical_report.docx"

    cross_references = {
        "Table 1": "Table 1. Experimental Parameters",
        "Table 2": "Table 2. Measurement Results",
        "Figure 1": "Figure 1. System Architecture",
        "Figure 2": "Figure 2. Data Flow Diagram",
    }

    for phrase, target_heading in cross_references.items():
        try:
            count = add_hyperlink_to_heading(
                file_path=file_path,
                phrase=phrase,
                target_heading=target_heading,
            )
            if count > 0:
                print(f"Linked {count} occurrence(s) of '{phrase}'")
        except ValueError as e:
            print(f"Could not link '{phrase}': {e}")


def example_section_navigation():
    """
    Add navigation links at the beginning of a document.
    """
    file_path = "long_document.docx"

    sections = [
        "Introduction",
        "Background",
        "Methodology",
        "Implementation",
        "Results",
        "Discussion",
        "Conclusion",
        "References",
        "Appendix",
    ]

    print("Adding navigation hyperlinks...")
    for section in sections:
        try:
            count = add_hyperlink_to_heading(
                file_path=file_path,
                phrase=section,
                target_heading=section,
            )
            print(f"  {section}: {count} link(s)")
        except ValueError as e:
            print(f"  {section}: Not found")


if __name__ == "__main__":
    print("=" * 60)
    print("Hyperlink Management Examples")
    print("=" * 60)

    print("\n1. Basic Hyperlink:")
    print("   See example_basic_hyperlink()")

    print("\n2. Multiple Hyperlinks:")
    print("   See example_multiple_hyperlinks()")

    print("\n3. Cross-Reference Hyperlinks:")
    print("   See example_cross_reference_hyperlinks()")

    print("\n4. Section Navigation:")
    print("   See example_section_navigation()")

    print("\n" + "=" * 60)
    print("Note: Ensure document files exist before running examples.")
    print("      Target headings must exist in the document for links to work.")
    print("=" * 60)
