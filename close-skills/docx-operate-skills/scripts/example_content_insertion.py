"""
Example: Content Management Workflows

This script demonstrates content management operations including
clearing section content and inserting new headings.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from docx_operations import (
    clear_section_content,
    insert_heading_after_subtitles,
    get_subtitles_under_heading,
)


def example_clear_section_content():
    """
    Clear content under a heading while preserving the heading itself.
    """
    file_path = "draft_report.docx"
    title = "Draft Notes"

    try:
        clear_section_content(
            file_path=file_path,
            title=title,
        )
        print(f"Cleared content under '{title}'")
    except ValueError as e:
        print(f"Error: {e}")


def example_clear_with_exclusions():
    """
    Clear section content but preserve specific subheadings and their content.
    """
    file_path = "report.docx"
    title = "Results"
    keep_subtitles = ["Important Findings", "Key Data"]

    try:
        clear_section_content(
            file_path=file_path,
            title=title,
            keep_subtitles=keep_subtitles,
        )
        print(f"Cleared content, preserved: {keep_subtitles}")
    except ValueError as e:
        print(f"Error: {e}")


def example_insert_heading():
    """
    Insert a new heading after all existing subheadings.
    """
    file_path = "document.docx"
    heading_text = "Summary"
    target_title = "Results"
    level = 2

    try:
        insert_heading_after_subtitles(
            file_path=file_path,
            heading_text=heading_text,
            target_title=target_title,
            level=level,
        )
        print(f"Inserted heading '{heading_text}' under '{target_title}'")
    except ValueError as e:
        print(f"Error: {e}")


def example_list_subtitles():
    """
    List all subtitles under a parent heading.
    """
    file_path = "template.docx"
    parent_title = "Results"

    try:
        subtitles = get_subtitles_under_heading(
            file_path=file_path,
            parent_title=parent_title,
        )
        print(f"Subtitles under '{parent_title}':")
        for original, normalized in subtitles:
            print(f"  - {original} (normalized: {normalized})")
    except ValueError as e:
        print(f"Error: {e}")


def example_workflow_clear_and_insert():
    """
    Complete workflow: Clear old content and insert new heading.
    """
    file_path = "report.docx"
    
    print("Step 1: Clear old content")
    try:
        clear_section_content(
            file_path=file_path,
            title="Old Section",
        )
        print("  Cleared successfully")
    except ValueError as e:
        print(f"  Error: {e}")
        return

    print("Step 2: Insert new heading")
    try:
        insert_heading_after_subtitles(
            file_path=file_path,
            heading_text="New Section",
            target_title="Main Chapter",
            level=2,
        )
        print("  Inserted successfully")
    except ValueError as e:
        print(f"  Error: {e}")


if __name__ == "__main__":
    print("=" * 60)
    print("Content Management Examples")
    print("=" * 60)

    print("\n1. Clear Section Content:")
    print("   See example_clear_section_content()")

    print("\n2. Clear with Exclusions:")
    print("   See example_clear_with_exclusions()")

    print("\n3. Insert Heading:")
    print("   See example_insert_heading()")

    print("\n4. List Subtitles:")
    print("   See example_list_subtitles()")

    print("\n5. Complete Workflow:")
    print("   See example_workflow_clear_and_insert()")

    print("\n" + "=" * 60)
    print("Note: Ensure document files exist before running examples.")
    print("=" * 60)
