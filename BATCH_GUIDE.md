# Batch PowerPoint to Markdown Converter Guide

This guide covers the batch processing capabilities of the PowerPoint to Markdown converter.

## 📁 Batch Converter Features

- **Multiple File Processing**: Converts all .pptx files in a directory
- **Recursive Search**: Automatically finds presentations in subdirectories
- **Flexible Output**: Choose same directory or separate output folder
- **Progress Tracking**: Real-time conversion status and progress
- **Error Handling**: Detailed error reporting for failed conversions
- **Summary Report**: Complete conversion statistics
- **Virtual Environment Support**: Automatically uses correct Python environment

## 🚀 Quick Start

### Windows Batch File (Easiest)
```batch
# Convert all presentations in a folder
batch_convert.bat "C:\My Presentations"

# Convert with custom output folder
batch_convert.bat "C:\Lectures" "C:\Markdown Output"
```

### Python Script
```bash
# Convert all .pptx files in current directory
python batch_convert.py .

# Convert specific folder
python batch_convert.py "path/to/presentations"

# Convert with custom output folder
python batch_convert.py "input_folder" "output_folder"
```

## 📋 Command Line Options

```bash
python batch_convert.py [-h] [--converter CONVERTER] [--recursive] [--version]
                        input_folder [output_folder]
```

### Positional Arguments:
- `input_folder`: Path to folder containing PowerPoint files (.pptx)
- `output_folder`: (Optional) Path to output folder for Markdown files

### Optional Arguments:
- `-h, --help`: Show help message and exit
- `--converter CONVERTER`: Path to the converter script (default: pptx_to_md.py)
- `--recursive, -r`: Search subdirectories recursively (default behavior)
- `--version`: Show version number and exit

## 📊 Example Usage

### Example 1: Basic Conversion
```bash
python batch_convert.py "C:\Course Materials"
```
**Result**: All .pptx files in the folder are converted to .md files in the same directory.

### Example 2: Organized Output
```bash
python batch_convert.py "C:\Lectures" "C:\Markdown Lectures"
```
**Result**: All presentations converted and saved to a separate markdown folder.

### Example 3: Current Directory
```bash
python batch_convert.py . ./output
```
**Result**: Convert all presentations in current directory to an 'output' subfolder.

## 📈 Sample Output

```
============================================================
📊 BATCH POWERPOINT TO MARKDOWN CONVERTER
============================================================
Input folder: C:\Lectures
Output folder: C:\Markdown Output
------------------------------------------------------------
🔍 Searching for PowerPoint files...
📁 Found 5 PowerPoint file(s)

[1/5] Processing: week1_introduction.pptx
  Converting: week1_introduction.pptx
  Output: C:\Markdown Output\week1_introduction.md
  ✅ Success!

[2/5] Processing: week2_variables.pptx
  Converting: week2_variables.pptx
  Output: C:\Markdown Output\week2_variables.md
  ✅ Success!

[3/5] Processing: week3_functions.pptx
  Converting: week3_functions.pptx
  Output: C:\Markdown Output\week3_functions.md
  ✅ Success!

[4/5] Processing: week4_classes.pptx
  Converting: week4_classes.pptx
  Output: C:\Markdown Output\week4_classes.md
  ❌ Failed: Permission denied

[5/5] Processing: final_review.pptx
  Converting: final_review.pptx
  Output: C:\Markdown Output\final_review.md
  ✅ Success!

============================================================
📋 CONVERSION SUMMARY
============================================================
Total files processed: 5
✅ Successful: 4
❌ Failed: 1

🚨 Failed conversions:
  • week4_classes.pptx: Permission denied

🎉 4 file(s) converted successfully!

Generated files:
  • C:\Markdown Output\week1_introduction.md
  • C:\Markdown Output\week2_variables.md
  • C:\Markdown Output\week3_functions.md
  • C:\Markdown Output\final_review.md

Total processing time: 12.3 seconds
```

## 🔧 Technical Details

### File Discovery
- Searches for files with `.pptx` and `.PPTX` extensions
- Recursively searches all subdirectories
- Removes duplicates and sorts files alphabetically

### Virtual Environment Support
- Automatically detects if running in a virtual environment
- Uses the correct Python executable for subprocess calls
- Preserves all environment variables

### Error Handling
- Continues processing even if individual files fail
- Captures detailed error messages
- Provides summary of all failures at the end

### Output Organization
- Preserves original filenames (changes extension to .md)
- Creates output directories if they don't exist
- Avoids overwriting without warning

## 🛠️ Troubleshooting

### Common Issues

**"No PowerPoint files found"**
- Check the input folder path
- Ensure files have .pptx extension
- Verify folder exists and is accessible

**"Permission denied" errors**
- Close PowerPoint if presentations are open
- Run command prompt as administrator (Windows)
- Check file/folder permissions

**"ModuleNotFoundError: No module named 'pptx'"**
- Activate the virtual environment first
- Install dependencies: `pip install -r requirements.txt`
- Use the batch_convert.bat file which handles this automatically

**Conversion takes a long time**
- Large presentations take longer to process
- Complex slides with many elements are slower
- Consider processing smaller batches

### Getting Help
```bash
python batch_convert.py --help
```

## 📝 Integration Examples

### Automated Course Material Processing
```batch
@echo off
echo Converting all course materials...
python batch_convert.py "C:\Course Materials\Week 1" "C:\Website\markdown\week1"
python batch_convert.py "C:\Course Materials\Week 2" "C:\Website\markdown\week2"
echo Conversion complete!
```

### PowerShell Script
```powershell
$courses = @("Python Basics", "Advanced Python", "Data Science")
foreach ($course in $courses) {
    python batch_convert.py "C:\Lectures\$course" "C:\Website\$course"
}
```

This batch converter makes it easy to process entire collections of PowerPoint presentations efficiently!
