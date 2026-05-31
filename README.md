# ⚽ Macauslot Soccer Odds Updater

[中文说明](README.zh-CN.md)

Fetch Macauslot soccer odds from JSON feeds and incrementally update an Excel workbook using the expected sample layout. The project provides both a command-line scraper and a Tkinter desktop interface.

## 📦 Requirements

- Python 3.10 or later
- Network access to `https://www.macauslot.com`

### macOS Setup

Install Python 3.13 and Tk 9 with Homebrew:

```bash
brew install python-tk@3.13
```

Create a project virtual environment and install dependencies:

```bash
cd /Users/niuhui/PycharmProjects/macauslotCrawler
/opt/homebrew/bin/python3.13 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
```

For later runs, activate the existing virtual environment:

```bash
cd /Users/niuhui/PycharmProjects/macauslotCrawler
source .venv/bin/activate
```

Exit the virtual environment when needed:

```bash
deactivate
```

### General Installation

If Python and Tk are already configured, install the dependency:

```bash
python3 -m pip install -r requirements.txt
```

## 🖥️ Desktop Interface

Launch the GUI:

```bash
cd /Users/niuhui/PycharmProjects/macauslotCrawler
source .venv/bin/activate
python macauslot_odds_gui.py
```

The interface supports:

- Immediate updates for full-time three-in-one odds
- Scheduled updates with a configurable interval and random delay
- HTTP/HTTPS proxy configuration
- Optional SSL certificate verification bypass for trusted proxy environments
- Live scraping logs

The default output file is:

```text
outputs/清洗后的表格.xlsx
```

## ⌨️ Command Line

Update the default Excel workbook:

```bash
python scrape_macauslot_odds.py
```

Export raw fields for all known betting types:

```bash
python scrape_macauslot_odds.py --format raw
```

Use a proxy:

```bash
python scrape_macauslot_odds.py --proxy-host 127.0.0.1 --proxy-port 7890
```

Show all available options:

```bash
python scrape_macauslot_odds.py --help
```

## 📊 Optional Excel Template

The program attempts to load:

```text
assets/sample_template.xlsx
```

This optional file is used to reuse spreadsheet styles and column widths. If it is not present, the program falls back to built-in formatting and continues to generate the Excel workbook.

## 📁 Project Files

```text
scrape_macauslot_odds.py  Scraping, data normalization, and Excel update logic
macauslot_odds_gui.py     Tkinter desktop interface and scheduled execution
requirements.txt          Python dependency list
README.zh-CN.md           Chinese documentation
```
