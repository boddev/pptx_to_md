#!/usr/bin/env python3
"""
Batch PowerPoint to Markdown Converter

This script processes all PowerPoint files (.pptx) in a specified directory
and converts them to Markdown format using the pptx_to_md converter.

Usage:
    python batch_convert.py <input_folder> [output_folder]
    
Examples:
    python batch_convert.py "C:/My Presentations"
    python batch_convert.py "C:/My Presentations" "C:/Output"
    python batch_convert.py . ./output
"""

import argparse
import os
import sys
from pathlib import Path
import subprocess
from datetime import datetime


def find_powerpoint_files(input_folder):
    """
    Find all PowerPoint files in the input folder.
    
    Args:
        input_folder (str): Path to the input directory
        
    Returns:
        list: List of PowerPoint file paths
    """
    input_path = Path(input_folder)
    
    if not input_path.exists():
        raise FileNotFoundError(f"Input folder not found: {input_folder}")
    
    if not input_path.is_dir():
        raise ValueError(f"Input path is not a directory: {input_folder}")
    
    # Find all .pptx files (case insensitive)
    pptx_files = []
    for ext in ['*.pptx', '*.PPTX']:
        pptx_files.extend(input_path.glob(ext))
    
    # Also search subdirectories
    for ext in ['**/*.pptx', '**/*.PPTX']:
        pptx_files.extend(input_path.glob(ext))
    
    # Remove duplicates and sort
    pptx_files = sorted(list(set(pptx_files)))
    
    return pptx_files


def get_python_executable():
    """
    Get the correct Python executable (considering virtual environment).
    
    Returns:
        str: Path to Python executable
    """
    # Check if we're in a virtual environment
    venv_path = os.environ.get('VIRTUAL_ENV')
    if venv_path:
        # We're in a virtual environment, use the venv python
        if os.name == 'nt':  # Windows
            python_exe = os.path.join(venv_path, 'Scripts', 'python.exe')
        else:  # Unix/Linux/Mac
            python_exe = os.path.join(venv_path, 'bin', 'python')
        
        if os.path.exists(python_exe):
            return python_exe
    
    # Fall back to system python
    return sys.executable


def convert_presentation(pptx_file, output_folder=None, converter_script="pptx_to_md.py"):
    """
    Convert a single PowerPoint presentation to Markdown.
    
    Args:
        pptx_file (Path): Path to the PowerPoint file
        output_folder (str, optional): Output directory for the Markdown file
        converter_script (str): Path to the converter script
        
    Returns:
        tuple: (success: bool, output_file: str, error_message: str)
    """
    try:
        # Determine output file path
        if output_folder:
            output_path = Path(output_folder)
            output_path.mkdir(parents=True, exist_ok=True)
            output_file = output_path / f"{pptx_file.stem}.md"
        else:
            # Place in same directory as input file
            output_file = pptx_file.parent / f"{pptx_file.stem}.md"
        
        # Get the correct Python executable
        python_exe = get_python_executable()
        
        # Run the converter
        cmd = [
            python_exe, 
            converter_script, 
            str(pptx_file), 
            str(output_file)
        ]
        
        print(f"  Converting: {pptx_file.name}")
        print(f"  Output: {output_file}")
        
        result = subprocess.run(
            cmd, 
            capture_output=True, 
            text=True, 
            cwd=os.getcwd(),
            env=os.environ.copy()  # Preserve environment variables
        )
        
        if result.returncode == 0:
            return True, str(output_file), ""
        else:
            error_msg = result.stderr or result.stdout or "Unknown error"
            return False, str(output_file), error_msg
            
    except Exception as e:
        return False, "", str(e)


def batch_convert(input_folder, output_folder=None, converter_script="pptx_to_md.py"):
    """
    Convert all PowerPoint presentations in a folder to Markdown.
    
    Args:
        input_folder (str): Path to input directory
        output_folder (str, optional): Path to output directory
        converter_script (str): Path to the converter script
        
    Returns:
        dict: Summary of conversion results
    """
    print("=" * 60)
    print("📊 BATCH POWERPOINT TO MARKDOWN CONVERTER")
    print("=" * 60)
    print(f"Input folder: {input_folder}")
    if output_folder:
        print(f"Output folder: {output_folder}")
    else:
        print("Output: Same directory as input files")
    print("-" * 60)
    
    # Check if converter script exists
    if not os.path.exists(converter_script):
        raise FileNotFoundError(f"Converter script not found: {converter_script}")
    
    # Find all PowerPoint files
    print("🔍 Searching for PowerPoint files...")
    pptx_files = find_powerpoint_files(input_folder)
    
    if not pptx_files:
        print("❌ No PowerPoint files found in the specified directory.")
        return {"total": 0, "success": 0, "failed": 0, "files": []}
    
    print(f"📁 Found {len(pptx_files)} PowerPoint file(s)")
    print()
    
    # Convert each file
    results = {
        "total": len(pptx_files),
        "success": 0,
        "failed": 0,
        "files": []
    }
    
    for i, pptx_file in enumerate(pptx_files, 1):
        print(f"[{i}/{len(pptx_files)}] Processing: {pptx_file.name}")
        
        success, output_file, error_msg = convert_presentation(
            pptx_file, output_folder, converter_script
        )
        
        file_result = {
            "input": str(pptx_file),
            "output": output_file,
            "success": success,
            "error": error_msg
        }
        results["files"].append(file_result)
        
        if success:
            results["success"] += 1
            print(f"  ✅ Success!")
        else:
            results["failed"] += 1
            print(f"  ❌ Failed: {error_msg}")
        
        print()
    
    return results


def print_summary(results):
    """Print a summary of the batch conversion results."""
    print("=" * 60)
    print("📋 CONVERSION SUMMARY")
    print("=" * 60)
    print(f"Total files processed: {results['total']}")
    print(f"✅ Successful: {results['success']}")
    print(f"❌ Failed: {results['failed']}")
    
    if results['failed'] > 0:
        print("\n🚨 Failed conversions:")
        for file_result in results["files"]:
            if not file_result["success"]:
                print(f"  • {Path(file_result['input']).name}: {file_result['error']}")
    
    if results['success'] > 0:
        print(f"\n🎉 {results['success']} file(s) converted successfully!")
        print("\nGenerated files:")
        for file_result in results["files"]:
            if file_result["success"]:
                print(f"  • {file_result['output']}")


def main():
    """Main function to handle command line arguments and run batch conversion."""
    parser = argparse.ArgumentParser(
        description="Batch convert PowerPoint presentations to Markdown format",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
    python batch_convert.py "C:/My Presentations"
    python batch_convert.py "C:/Lectures" "C:/Output"
    python batch_convert.py . ./markdown_output
    python batch_convert.py ~/Documents/PowerPoints
        """
    )
    
    parser.add_argument(
        'input_folder',
        help='Path to the folder containing PowerPoint files (.pptx)'
    )
    
    parser.add_argument(
        'output_folder',
        nargs='?',
        help='Path to the output folder for Markdown files (optional, defaults to same folder as input)'
    )
    
    parser.add_argument(
        '--converter',
        default='pptx_to_md.py',
        help='Path to the PowerPoint to Markdown converter script (default: pptx_to_md.py)'
    )
    
    parser.add_argument(
        '--recursive', '-r',
        action='store_true',
        help='Search subdirectories recursively (default behavior)'
    )
    
    parser.add_argument(
        '--version',
        action='version',
        version='Batch PowerPoint to Markdown Converter 1.0'
    )
    
    args = parser.parse_args()
    
    try:
        # Run batch conversion
        start_time = datetime.now()
        results = batch_convert(
            args.input_folder, 
            args.output_folder, 
            args.converter
        )
        end_time = datetime.now()
        
        # Print summary
        print_summary(results)
        
        duration = end_time - start_time
        print(f"\nTotal processing time: {duration.total_seconds():.1f} seconds")
        
        # Exit with appropriate code
        if results['failed'] > 0:
            sys.exit(1)  # Some files failed
        else:
            sys.exit(0)  # All successful
            
    except FileNotFoundError as e:
        print(f"❌ Error: {e}")
        sys.exit(1)
        
    except ValueError as e:
        print(f"❌ Error: {e}")
        sys.exit(1)
        
    except KeyboardInterrupt:
        print("\n\n⚠️ Conversion interrupted by user")
        sys.exit(1)
        
    except Exception as e:
        print(f"❌ Unexpected error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
