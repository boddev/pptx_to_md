#!/usr/bin/env python3
"""
Quick start example for PowerPoint to Markdown converter.

This script demonstrates how to use the converter programmatically.
"""

from pptx_to_md import convert_pptx_to_markdown
import os


def main():
    """Example usage of the converter."""
    print("PowerPoint to Markdown Converter - Example Usage")
    print("=" * 50)
    
    # Check if test file exists
    test_file = "test_presentation.pptx"
    if not os.path.exists(test_file):
        print(f"Creating test presentation: {test_file}")
        # Import and run the test creator
        from create_test_pptx import create_test_presentation
        create_test_presentation()
    
    # Convert the presentation
    print(f"\nConverting {test_file} to Markdown...")
    try:
        output_file = convert_pptx_to_markdown(test_file, "example_output.md")
        print(f"Success! Created: {output_file}")
        
        # Show some stats
        with open(output_file, 'r', encoding='utf-8') as f:
            content = f.read()
            lines = len(content.splitlines())
            words = len(content.split())
            
        print(f"\nOutput Statistics:")
        print(f"- Lines: {lines}")
        print(f"- Words: {words}")
        print(f"- Characters: {len(content)}")
        
        print(f"\nFirst few lines of output:")
        print("-" * 30)
        with open(output_file, 'r', encoding='utf-8') as f:
            for i, line in enumerate(f):
                if i < 10:  # Show first 10 lines
                    print(line.rstrip())
                else:
                    break
                    
    except Exception as e:
        print(f"Error: {e}")


if __name__ == "__main__":
    main()
