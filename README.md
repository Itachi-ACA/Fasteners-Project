# IndoTech Hardware Automation Platform

A desktop application for automating hardware/fastener documentation. It extracts bolt and stud specifications from engineering PDFs and generates professional Excel reports and PDF summaries — streamlining procurement and inventory workflows.

---

## Features

- **PDF Parsing** — Extracts hardware specifications (sizes, materials, quantities, nuts, washers) from work order and BOM PDFs using `pdfplumber`
- **Excel Report Generation** — Creates formatted Excel workbooks with a company logo, header information, summary tables, and signature blocks
- **PDF Export** — Converts generated Excel files to PDF using Windows COM automation
- **Weight Calculation** — Calculates total weights using configurable, coating-specific weight tables
- **Material Coating Support** — Handles HDG, MAGNI, SS, HT, BLACK, and GALV coatings
- **Auto Rounding** — Odd lengths (X25, X35, X45…) are automatically rounded up to the nearest 10
- **Generation History** — Records work orders, customers, weights, and timestamps
- **Activity Log** — Color-coded system event log
- **Drag & Drop** — Supports dragging PDFs directly onto the application window
- **Dark / Light Theme** — Configurable UI theme
- **Settings** — Configurable approver names, format numbers, output folder, and more

---

## Requirements

- **OS**: Windows (required for Excel COM automation and PDF export)
- **Python**: 3.8 or higher
- **Microsoft Excel**: Must be installed for PDF export functionality

### Python Dependencies

Install the required packages using `pip`:

```bash
pip install pdfplumber openpyxl pandas pillow pywin32 tkinterdnd2
```

| Package | Purpose |
|---|---|
| `pdfplumber` | Extract text from PDF files |
| `openpyxl` | Create and format Excel workbooks |
| `pandas` | Data manipulation |
| `Pillow` | Image handling (logo, splash screen) |
| `pywin32` | Windows COM automation for Excel → PDF export |
| `tkinterdnd2` | Drag-and-drop support in the GUI |

---

## Installation

1. **Clone the repository**

   ```bash
   git clone https://github.com/Itachi-ACA/Fasteners-Project.git
   cd Fasteners-Project
   ```

2. **Create and activate a virtual environment** (recommended)

   ```bash
   python -m venv venv
   venv\Scripts\activate
   ```

3. **Install dependencies**

   ```bash
   pip install pdfplumber openpyxl pandas pillow pywin32 tkinterdnd2
   ```

4. **Run the application**

   ```bash
   python hardware_generator.py
   ```

---

## How to Use

1. Launch the application by running `hardware_generator.py`.
2. Click **"Select Hardware PDF"** (or drag and drop a PDF onto the window).
3. Choose your **output folder** for the generated files.
4. Click **"Generate"** — the application will:
   - Parse hardware specifications from the PDF
   - Build a summary with quantities and weights
   - Create a formatted Excel file
   - Export a PDF copy
5. The generated files will open automatically (configurable in Settings).

---

## Configuration

Application settings are stored in `config.json` in the project root and are managed through the built-in **Settings** page. Available options include:

| Setting | Description |
|---|---|
| `weight_table_path` | Path to an external Excel file containing coating-specific weight data |
| `last_output_folder` | Default output directory for generated files |
| `theme` | UI theme: `"dark"` or `"light"` |
| `open_after_gen` | Automatically open generated files after export |
| `autosave_folder` | Reuse the last output folder without prompting |
| `prepared_by` | Name/initials for the "Prepared By" field in documents |
| `checked_by` | Name/initials for the "Checked By" field in documents |
| `approved_by` | Name/initials for the "Approved By" field in documents |
| `format_number` | Document format number (e.g. `F/DSN/54`) |

---

## Weight Table

- By default, only **HDG** (Hot-Dip Galvanized) coating weights are active.
- **MAGNI, BLACK, SS, HT, GALV** weights will show `N/A` until values are provided.

### Adding Custom Coating Weights

You can upload a custom Excel weight table via the **Weight Table Manager** tab, or add entries directly in the `WEIGHT_TABLE` dictionary in `hardware_generator.py`:

```python
"M16X50-MAGNI": 27,
"M20X80-SS": 45,
```

The key format is `"<SIZE>-<COATING>"` and the value is the weight in grams.

---

## Rounding Rule

Bolt lengths with odd values (X25, X35, X45…) are automatically rounded **up** to the nearest 10 (X30, X40, X50…). The rounded size is used in the output document.

---

## Project Structure

```
Fasteners-Project/
├── hardware_generator.py   # Main application
├── config.json             # User configuration (auto-generated)
├── README.md               # This file
├── .gitignore
├── img/
│   ├── app.ico             # Application icon
│   └── logo.png            # Company logo (used in UI and Excel exports)
└── sound/
    └── startup.wav         # Startup sound effect
```

---

## Troubleshooting

| Problem | Solution |
|---|---|
| Application fails to generate | Close any already-open Excel files and try again |
| PDF not parsed correctly | Ensure the PDF contains selectable text (not a scanned image) |
| PDF export fails | Verify that Microsoft Excel is installed on the machine |
| Weight shows N/A | Add the coating weight entry to the weight table |
| Drag & drop not working | Ensure `tkinterdnd2` is installed correctly |

---

## Notes

- This is an **internal IndoTech tool** intended for use by the engineering and procurement teams.
- The application is Windows-only due to its reliance on Microsoft Excel COM automation for PDF generation.
- Material coatings for **M6 and below** are automatically forced to **SS** (Stainless Steel).

---

*IndoTech Internal Software*
