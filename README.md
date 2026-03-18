# WhatsApp to Excel Parser

A Python utility to convert exported WhatsApp chat files (.txt) into formatted Excel (.xlsx) spreadsheets. It handles multiline messages, extracts attachment filenames, and applies styles for better readability.

## Features

- Converts plain text WhatsApp exports to Excel format.
- Deducting formats: Supports both `[YYYY-MM-DD, HH:MM:SS]` and `YYYY/MM/DD, HH:MM` timestamp formats.
- Extracts attachment names from message text.
- Formats Excel headers and columns for readability.
- Multi-chat processing: Can process multiple chat files in one run.

## Prerequisites

- Python 3.x
- Virtual environment (recommended)
- `openpyxl` library

## Installation

1.  **Clone the repository**:
    ```bash
    git clone https://github.com/ebendl/whatsapp-to-excel.git
    cd whatsapp-to-excel
    ```
2.  **Set up the environment**:
    ```bash
    python3 -m venv .venv
    source .venv/bin/activate
    pip install -r requirements.txt
    ```

## Usage

1.  Export your WhatsApp chat as a `.txt` file (without media).
2.  Update the `chats` list in `parse_whatsapp.py` with the paths to your text files and desired output locations.
3.  Run the script:
    ```bash
    python3 parse_whatsapp.py
    ```

## License

MIT
