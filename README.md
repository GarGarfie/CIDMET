# CIDMET — Cross-database Identification and De-duplication Matching Export Tool

[![en](https://img.shields.io/badge/lang-English-blue.svg)](README.md) [![zh](https://img.shields.io/badge/lang-中文-red.svg)](README_zh.md) [![ru](https://img.shields.io/badge/lang-Русский-green.svg)](README_ru.md)

---

## Introduction

**CIDMET** is a desktop application for bibliometric researchers. It takes your local BibTeX reference library and automatically matches entries against exported data from three major academic databases — **Web of Science (WoS)**, **Scopus**, and **Engineering Village (EI/Compendex)** — then extracts matched subsets in each database's native format and provides a merged, deduplicated export.

This solves a common pain point in bibliometric analysis: when you need database-specific export files (for tools like VOSviewer, CiteSpace, or Bibliometrix) that contain only the references in your study, rather than an entire search result set.

## Features

- **Multi-database support** — WoS (TXT / XLS), Scopus (CSV / TXT, English & Chinese), EI (CSV / TXT)
- **Three-tier matching strategy**
  - DOI exact match (100% confidence)
  - Title exact match with Unicode normalization (99% confidence)
  - Fuzzy title match with author & year validation (configurable threshold, default 90%)
- **Format-preserving subset export** — output files are immediately usable by bibliometric software
- **Merged export with author format conversion** — automatically converts author name formats to match WoS / Scopus / EI conventions
- **Automatic deduplication** — detects and lets you review entries matched by multiple database records
- **Drag-and-drop GUI** — built with PySide6 (Qt), with progress tracking and tabbed output
- **Bilingual Scopus support** — handles both English and Chinese Scopus exports

## Matching Strategy

| Tier | Method | Confidence | Description |
|------|--------|------------|-------------|
| 1 | DOI exact match | 100% | Normalized DOI comparison (case-insensitive, prefix-stripped) |
| 2 | Title exact match | 99% | NFKD-normalized, case-insensitive, special characters removed |
| 3 | Fuzzy title match | Configurable (default 90%) | RapidFuzz similarity + first-author surname validation (≥80%) + year validation |

## Supported Formats

| Database | Import Formats | Subset Output | Merged Output |
|----------|---------------|---------------|---------------|
| Web of Science | TXT (tagged), XLS | TXT, XLS | TXT |
| Scopus | CSV, TXT (EN/CN) | CSV, TXT | CSV |
| Engineering Village | CSV, TXT | CSV, TXT | CSV |

## Installation

**Requirements:** Python 3.9+

```bash
# Clone the repository
git clone https://github.com/GarGarfie/CIDMET.git
cd CIDMET

# Install dependencies
pip install -r requirements.txt
```

### Dependencies

| Package | Purpose |
|---------|---------|
| PySide6 ≥ 6.5 | GUI framework (Qt for Python) |
| bibtexparser ≥ 1.4, < 2.0 | BibTeX file parsing |
| rapidfuzz ≥ 3.0 | Fuzzy string matching |
| chardet ≥ 5.0 | Character encoding detection |
| xlrd ≥ 2.0 | Read Excel .xls files |
| xlwt ≥ 1.3 | Write Excel .xls files |
| openpyxl ≥ 3.1 | Write Excel .xlsx files |

## Usage

```bash
python main.py
```

### Workflow

1. **Select BibTeX file** — choose your target reference library (`.bib`)
2. **Add database files** — drag-and-drop or browse for WoS / Scopus / EI export files
3. **Set output directory** — choose where subset and merged files will be saved
4. **Adjust fuzzy threshold** (optional) — slider from 50% to 100%, default 90%
5. **Click "Run Matching"** — the tool processes files and performs three-tier matching
6. **Review results** — check the Results tab for statistics, match details, and unmatched entries
7. **Export merged file** (optional) — select a target format template and export combined records

## Project Structure

```
CIDMET/
├── main.py              # Application entry point
├── gui_app.py           # PySide6 GUI (MainWindow, drag-drop, progress, tabs)
├── parsers.py           # Database format parsers (WoS/Scopus/EI × TXT/CSV/XLS)
├── matcher.py           # Three-tier matching engine
├── writers.py           # Format-preserving subset writers & merged export
├── utils.py             # Encoding detection, DOI/title normalization, helpers
├── draw_flowchart.py    # Data flow diagram generator
├── requirements.txt     # Python dependencies
└── fileTemplate/        # Example database export files for reference
```

## Author Format Conversion (Merged Export)

When merging matched records from different databases, CIDMET automatically converts author names to the target format:

| Target | Short Form | Full Form |
|--------|-----------|-----------|
| WoS | `Gu, S; Wu, YQ` | `Gu, Sheng; Wu, Yanqi` |
| Scopus | `Gu, S.; Wu, Y.Q.` | `Gu, Sheng; Wu, Yanqi` |
| EI | `Gu, Sheng (1); Wu, Yanqi (1)` | — |

## License

This project is licensed under the [MIT License](LICENSE).

## Citation

If you use CIDMET in your research, please consider citing it.
