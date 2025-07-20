#!/usr/bin/env python3
"""
Comprehensive test script for the PowerPoint to Markdown converter.
This script runs various tests to ensure the converter works correctly.
"""

import os
import sys
import subprocess
from pathlib import Path


def run_command(command, description):
    """Run a command and return success status."""
    print(f"\n🧪 Testing: {description}")
    print(f"Command: {command}")
    print("-" * 50)
    
    try:
        result = subprocess.run(command, shell=True, capture_output=True, text=True)
        
        if result.returncode == 0:
            print("✅ SUCCESS")
            if result.stdout:
                print("Output:", result.stdout)
            return True
        else:
            print("❌ FAILED")
            if result.stderr:
                print("Error:", result.stderr)
            if result.stdout:
                print("Output:", result.stdout)
            return False
            
    except Exception as e:
        print(f"❌ EXCEPTION: {e}")
        return False


def test_converter():
    """Run comprehensive tests for the converter."""
    
    print("=" * 60)
    print("🚀 POWERPOINT TO MARKDOWN CONVERTER - TEST SUITE")
    print("=" * 60)
    
    # Test 1: Create test presentation
    success1 = run_command(
        "python create_test_pptx.py",
        "Creating test PowerPoint presentation"
    )
    
    # Test 2: Basic conversion
    success2 = run_command(
        "python pptx_to_md.py test_presentation.pptx",
        "Basic conversion (auto-named output)"
    )
    
    # Test 3: Custom output filename
    success3 = run_command(
        "python pptx_to_md.py test_presentation.pptx custom_output.md",
        "Conversion with custom output filename"
    )
    
    # Test 4: Help command
    success4 = run_command(
        "python pptx_to_md.py --help",
        "Help command"
    )
    
    # Test 5: Version command
    success5 = run_command(
        "python pptx_to_md.py --version",
        "Version command"
    )
    
    # Test 6: Error handling - non-existent file
    print(f"\n🧪 Testing: Error handling for non-existent file")
    print(f"Command: python pptx_to_md.py nonexistent.pptx")
    print("-" * 50)
    result = subprocess.run(
        "python pptx_to_md.py nonexistent.pptx", 
        shell=True, 
        capture_output=True, 
        text=True
    )
    success6 = result.returncode != 0  # Should fail with non-zero exit code
    if success6:
        print("✅ SUCCESS - Correctly handled missing file")
        print("Error output:", result.stderr if result.stderr else result.stdout)
    else:
        print("❌ FAILED - Should have failed for missing file")
    
    # Test 7: Check output files exist
    test_files = ["test_presentation.md", "custom_output.md"]
    success7 = True
    for file in test_files:
        if os.path.exists(file):
            print(f"✅ Output file created: {file}")
            # Check file size
            size = os.path.getsize(file)
            print(f"   File size: {size} bytes")
            if size > 0:
                print(f"   ✅ File has content")
            else:
                print(f"   ❌ File is empty")
                success7 = False
        else:
            print(f"❌ Output file missing: {file}")
            success7 = False
    
    # Summary
    print("\n" + "=" * 60)
    print("📊 TEST SUMMARY")
    print("=" * 60)
    
    tests = [
        ("Create test presentation", success1),
        ("Basic conversion", success2),
        ("Custom output filename", success3),
        ("Help command", success4),
        ("Version command", success5),
        ("Error handling", success6),
        ("Output files check", success7),
    ]
    
    passed = sum(1 for _, success in tests if success)
    total = len(tests)
    
    for test_name, success in tests:
        status = "✅ PASS" if success else "❌ FAIL"
        print(f"{status}: {test_name}")
    
    print(f"\n🎯 RESULTS: {passed}/{total} tests passed")
    
    if passed == total:
        print("🎉 ALL TESTS PASSED! The converter is working correctly.")
        return True
    else:
        print("⚠️  Some tests failed. Please review the output above.")
        return False


if __name__ == "__main__":
    success = test_converter()
    sys.exit(0 if success else 1)
