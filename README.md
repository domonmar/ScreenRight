# ScreenRight - Screenplay Formatter

## Overview

ScreenRight is a Python script designed to format screenplays written in Microsoft Word (.docx files) according to industry standards. It ensures proper indentation, spacing, and text styling, transforming your screenplay into a professionally formatted document.

## Features

- **Customizable Parameters**: Formatting rules can be customized through the GUI. Default settings are bundled with the application, and user preferences are saved locally.
- **Paragraph Type Detection**: Automatically detects and formats different types of paragraphs (e.g., scene headings, character names, dialogue, action, parenthetical).
- **Page Numbering**: Adds page numbers to the document, starting from the second page.
- **Section Break Removal**: Removes unnecessary section breaks to ensure consistent formatting.
- **Dynamic Paragraph Formatting**: Applies specific indentation and alignment rules based on paragraph type.

## Requirements

- Python 3.x
- `python-docx` library (install with `pip install python-docx`)

## Customization

Formatting parameters can be customized through the GUI's "Parameters" tab. The application includes sensible defaults and automatically saves your preferences to your user directory. You can also reset to defaults using the "Reset to Defaults" button.

Available parameters include:
- Font settings (name, size, line spacing)
- Indentation settings for different paragraph types (character, action, scene, dialogue, parenthetical)
- Formatting start marker

## License

This project is released under the MIT License.
