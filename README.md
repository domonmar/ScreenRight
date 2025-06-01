# ScreenRight - Screenplay Formatter

## Overview

ScreenRight is a Python script designed to format screenplays written in Microsoft Word (.docx files) according to industry standards. It ensures proper indentation, spacing, and text styling, transforming your screenplay into a professionally formatted document.

## Features

- **Customizable Parameters**: Reads formatting rules (margins, font, font size, and line spacing) from a `parameters.txt` file.
- **Paragraph Type Detection**: Automatically detects and formats different types of paragraphs (e.g., scene headings, character names, dialogue, action, parenthetical).
- **Page Numbering**: Adds page numbers to the document, starting from the second page.
- **Section Break Removal**: Removes unnecessary section breaks to ensure consistent formatting.
- **Dynamic Paragraph Formatting**: Applies specific indentation and alignment rules based on paragraph type.

## Requirements

- Python 3.x
- `python-docx` library (install with `pip install python-docx`)

## Customization

To modify formatting rules, edit the `parameters.txt` file, which contains key-value pairs such as:

```txt
Font: Courier
Font Size: 12
Line Spacing: 22
Character Indent Left: 4.2
```

## License

This project is released under the MIT License.
