# PDF to Excel Converter — Bengali Book Data Extractor

A Python script that parses a structured Bengali PDF book and extracts **chapters**, **sections**, and **hadiths** into a formatted Excel workbook with multiple tabs.

---

## Features

- Extracts **chapters** marked with `*` in the PDF
- Extracts **sections** marked with `**` in the PDF
- Extracts **hadiths** starting with Bengali-numbered brackets e.g. `[১]`, `[২]`
- Strips the leading bracket number from hadith text (the `id` column serves as the number)
- Outputs a clean `.xlsx` file with 4 tabs: `chapter`, `secttion`, `hadith`, `summary`
- All cells use **Kalpurush** font for correct Bengali rendering
- Alternating row colors and styled headers for readability

---

## Requirements

### Python Version
Python 3.7 or higher

### Dependencies

Install all required packages with:

```bash
pip install pymupdf openpyxl
```

| Package | Purpose |
|---|---|
| `pymupdf` | Reading and parsing the PDF file |
| `openpyxl` | Creating and formatting the Excel output |

### Font

This script sets **Kalpurush** as the font for all Excel cells. Make sure Kalpurush is installed on your system, otherwise Bengali text may not render correctly in Excel.

- Download Kalpurush: [https://www.omicronlab.com](https://www.omicronlab.com)

---

## Usage

### 1. Clone or download the script

Place `pdf_to_excel.py` in any folder on your machine.

### 2. Update the file paths

Open `pdf_to_excel.py` and update these two lines at the top:

```python
PDF_PATH = "/path/to/your/input.pdf"
OUT_PATH = "/path/to/your/output.xlsx"
```

### 3. Run the script

```bash
python pdf_to_excel.py
```

You will see output like:

```
Extracting data from PDF …
  Chapters : 12
  Sections : 252
  Hadiths  : 665

Saved → /path/to/your/output.xlsx
```

---

## PDF Format Requirements

For the script to work correctly, the source PDF must follow this structure:

| Element | Marker | Example |
|---|---|---|
| Chapter | Starts with `*` (single asterisk) | `*অধ্যায়: পিতা-মাতার সাথে সদ্ব্যবহার` |
| Section | Starts with `**` (double asterisk) | `**মাথয়র সাথে সদাচরণ` |
| Hadith | Starts with `[Bengali number]` | `[১] আমর ইবনু শাইবানি...` |

> **Note:** Chapter names are automatically cleaned — the `অধ্যায়:` prefix is removed so only the title is stored.

---

## Output Excel Structure

The generated `.xlsx` file contains **4 tabs**:

### Tab 1 — `chapter`
| Column | Description |
|---|---|
| `id` | Auto-incremented chapter number (1, 2, 3 …) |
| `name` | Chapter title (cleaned, without the `*` or `অধ্যায়:` prefix) |

### Tab 2 — `secttion`
| Column | Description |
|---|---|
| `id` | Auto-incremented section number (1, 2, 3 …) |
| `name` | Section title (cleaned, without the `**` prefix) |

### Tab 3 — `hadith`
| Column | Description |
|---|---|
| `id` | Auto-incremented hadith number (1, 2, 3 …) |
| `hadith` | Full hadith text — multi-line hadiths are joined into one cell. The `[number]` bracket is removed from the text. |

### Tab 4 — `summary`
A quick overview showing the total count of chapters, sections, and hadiths extracted.

---

## Project Structure

```
your-folder/
│
├── pdf_to_excel.py     # Main converter script
├── README.md           # This file
├── input.pdf           # Your source PDF (Bengali book)
└── output.xlsx         # Generated Excel file
```

---

## Troubleshooting

**Bengali text appears as boxes or gibberish in Excel**
→ Install the Kalpurush font on your system and restart Excel.

**`ModuleNotFoundError: No module named 'fitz'`**
→ Run `pip install pymupdf` — the Python package name is `pymupdf` but it imports as `fitz`.

**Chapters / sections not detected**
→ Make sure your PDF uses exactly `*` for chapters and `**` for sections at the start of the line with no leading spaces.

**Hadiths are missing or incomplete**
→ Check that each hadith in the PDF starts with a Bengali number in square brackets like `[১]`. Any text before the next `[number]` or `*` marker is treated as part of the same hadith.

---

## License

This project is free to use and modify for personal or educational purposes.
