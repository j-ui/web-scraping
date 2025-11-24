
# Web Scraping Tool (Excel + GUI)

Simple Tkinter-based scraper.  
Input: `.xlsx` with a column named `URL`  
Output: `output_webscraping.xlsx` saved in the **same folder** as the input.

## Features
- Extracts: title, description, keywords, robots, canonical
- Extracts Open Graph: og:title, og:description, og:image
- Extracts headers: last-modified, server, content-type
- Extracts: H1s, internal/external/total links, word count
- Extracts all `<meta>` tags as JSON
- Handles non-200 responses, timeouts, invalid URLs safely

## Setup (Mac)
```bash
brew install python
git clone https://github.com/j-ui/web-scraping.git
cd web-scraping
python3 -m venv .venv
source .venv/bin/activate
pip install requests beautifulsoup4 lxml pandas openpyxl
```

## Run 
python scraper.py

Select your .xlsx → click Start web scraping.