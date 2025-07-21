# PowerPoint to Markdown Converter

A Python scr### Specify Output File
```bash
python pptx_to_md.py presentation.pptx output.md
```

### Batch Processing
Convert multiple PowerPoint presentations at once:

```bash
# Convert all .pptx files in a folder (outputs to same folder)
python batch_convert.py "path/to/presentations"

# Convert with custom output folder
python batch_convert.py "input_folder" "output_folder"

# Windows batch script (auto-activates virtual environment)
batch_convert.bat "C:\My Presentations"
```

The batch converter will:
- Find all .pptx files in the specified folder (including subdirectories)
- Convert each presentation using the same features as the single-file converter
- Provide a summary of successful and failed conversions
- Show processing time and detailed error messages if any conversion failsthat converts PowerPoint presentations (.pptx) to Markdown (.md) format, preserving text formatting, tables, and bullet points.

## Features

- Extracts text from all slides in a PowerPoint presentation
- **Automatically generates table of contents** with clickable links to each slide
- **Formats Python interpreter examples** as proper code blocks with syntax highlighting
- **Batch processing** support for converting multiple presentations at once
- Preserves bullet point formatting and indentation
- Converts tables to Markdown table format
- Organizes content by slide number with anchor links
- Handles grouped shapes and various text containers
- Notes presence of images
- **Collapsible TOC** for presentations with >20 slides
- Command-line interface with flexible output naming

## Installation

### Using Virtual Environment (Recommended)

1. Create a virtual environment:
```bash
python -m venv pptx_env
```

2. Activate the virtual environment:
```bash
# On Windows
pptx_env\Scripts\activate

# On macOS/Linux
source pptx_env/bin/activate
```

3. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

### Basic Usage
```bash
python pptx_to_md.py presentation.pptx
```

This will create a file named `presentation.md` in the same directory.

### Specify Output File
```bash
python pptx_to_md.py presentation.pptx my_output.md
```

### Help
```bash
python pptx_to_md.py --help
```

## Output Format

The generated Markdown file includes:
- Document header with presentation name and slide count
- Each slide as a separate section (## Slide N)
- Bullet points formatted as Markdown lists
- Tables converted to Markdown table format
- Image placeholders where images are present
- Slide separators for easy reading

## Example Output

```markdown
# My Presentation

Converted from PowerPoint presentation: `presentation.pptx`
Total slides: 3

## Slide 1

# Welcome to Our Product

- Key feature 1
- Key feature 2
- Key feature 3

---

## Slide 2

| Feature | Benefit | Cost |
| --- | --- | --- |
| Feature A | Saves time | $100 |
| Feature B | Increases efficiency | $200 |

---

## Slide 3

*[Image present]*

Thank you for your attention!
```

## Requirements

- Python 3.6+
- python-pptx library

## Troubleshooting

### Common Issues

1. **ModuleNotFoundError**: Make sure you've installed the requirements:
   ```bash
   pip install -r requirements.txt
   ```

2. **File not found**: Ensure the PowerPoint file path is correct and the file exists.

3. **Permission errors**: Make sure you have write permissions in the output directory.

## Development

To contribute to this project:

1. Fork the repository
2. Create a virtual environment and install dependencies
3. Make your changes
4. Test with various PowerPoint files
5. Submit a pull request

## License

This project is open source and available under the MIT License.
