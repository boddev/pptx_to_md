# PowerPoint to Markdown Converter - Quick Start Guide

## 🚀 Quick Setup

### Option 1: Automated Setup (Windows)
```bash
# Run the setup script
setup.bat
```

### Option 2: Manual Setup
```bash
# 1. Create virtual environment
python -m venv pptx_env

# 2. Activate virtual environment
pptx_env\Scripts\activate

# 3. Install dependencies
pip install -r requirements.txt
```

## 📝 Basic Usage

### Command Line
```bash
# Convert with auto-generated output filename
python pptx_to_md.py presentation.pptx

# Convert with custom output filename
python pptx_to_md.py presentation.pptx my_output.md

# Get help
python pptx_to_md.py --help
```

### Programmatic Usage
```python
from pptx_to_md import convert_pptx_to_markdown

# Convert presentation
output_file = convert_pptx_to_markdown("presentation.pptx", "output.md")
print(f"Converted to: {output_file}")
```

## 🧪 Testing

### Create Test Presentation
```bash
python create_test_pptx.py
```

### Run Comprehensive Tests
```bash
python test_suite.py
```

### Quick Example
```bash
python example_usage.py
```

## 📁 Project Structure

```
pptx_to_md/
├── pptx_to_md.py           # Main converter script
├── requirements.txt        # Python dependencies
├── setup.bat              # Windows setup script
├── create_test_pptx.py     # Test presentation generator
├── test_suite.py           # Comprehensive test suite
├── example_usage.py        # Usage example
├── README.md               # Full documentation
├── QUICK_START.md          # This file
└── pptx_env/               # Virtual environment (after setup)
```

## ✨ Features

- ✅ Extracts all text from PowerPoint slides
- ✅ Preserves bullet point formatting
- ✅ Converts tables to Markdown format
- ✅ Handles grouped shapes and text boxes
- ✅ Organizes content by slide number
- ✅ Command-line and programmatic interfaces
- ✅ Error handling and validation
- ✅ Virtual environment support

## 🔧 Requirements

- Python 3.6+
- python-pptx library (automatically installed)
- Windows PowerShell (for batch scripts)

## 💡 Tips

1. **Always use virtual environments** to avoid package conflicts
2. **Test with sample presentations** before using on important files
3. **Check output files** to ensure formatting meets your needs
4. **Use custom output names** to avoid overwriting existing files

## 🆘 Troubleshooting

### Common Issues

1. **Import Error**: Make sure virtual environment is activated
   ```bash
   pptx_env\Scripts\activate
   ```

2. **File Not Found**: Check file path and ensure .pptx extension
   ```bash
   # Wrong
   python pptx_to_md.py presentation
   
   # Correct
   python pptx_to_md.py presentation.pptx
   ```

3. **Permission Denied**: Ensure write permissions in output directory

4. **Unicode Errors**: File will still be created successfully; only display issue

### Getting Help

```bash
# Show help
python pptx_to_md.py --help

# Show version
python pptx_to_md.py --version
```

---

**Ready to convert your first presentation?**

1. Run: `setup.bat` (or manual setup)
2. Test: `python create_test_pptx.py`
3. Convert: `python pptx_to_md.py test_presentation.pptx`
4. Check: Open `test_presentation.md` to see the results!
