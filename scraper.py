"""
Simple generic web scraping tool with a Tkinter GUI.

What it does:
- Lets you choose an input Excel file with a 'URL' column.
- Scrapes for each URL:
    - HTTP status, final URL, Last-Modified, Content-Type, Server
    - <title>, meta description, meta keywords, robots, canonical URL
    - Open Graph (og:title, og:description, og:image)
    - Twitter (twitter:title, twitter:description, twitter:image)
    - H1 headings, link counts (internal/external/total), word count
    - All meta tags as a JSON string
- Automatically saves output as 'output_webscraping.xlsx'
  in the same folder as the input file.
- Shows a popup when scraping completes or fails.

Required:
    pip install requests beautifulsoup4 lxml pandas openpyxl
"""

import json
import os
import re
from typing import Dict, List, Optional, Tuple
from urllib.parse import urlparse

import pandas as pd
import requests
import tkinter as tk
from tkinter import filedialog, messagebox
from bs4 import BeautifulSoup


USER_AGENT = (
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
    "AppleWebKit/537.36 (KHTML, like Gecko) "
    "Chrome/122.0.0.0 Safari/537.36"
)
REQUEST_TIMEOUT_SECONDS = 12


def build_headers() -> Dict[str, str]:
    # Force encodings that Python handles well; avoid zstd
    return {
        "User-Agent": USER_AGENT,
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.9",
        "Accept-Encoding": "gzip, deflate, br",
        "Connection": "close",
    }


def normalize_url(url: str) -> str:
    url = url.strip()
    if not url:
        return url
    parsed = urlparse(url)
    if not parsed.scheme:
        return "https://" + url
    return url


def is_html_response(response: requests.Response) -> bool:
    content_type = response.headers.get("Content-Type", "")
    return "text/html" in content_type.lower()


def extract_meta(
    soup: BeautifulSoup,
    name: Optional[str] = None,
    prop: Optional[str] = None,
) -> Optional[str]:
    if name:
        tag = soup.find("meta", attrs={"name": name})
    elif prop:
        tag = soup.find("meta", attrs={"property": prop})
    else:
        return None

    if tag and tag.get("content"):
        return tag["content"].strip()
    return None


def count_links(soup: BeautifulSoup, base_netloc: str) -> Tuple[int, int, int]:
    internal = 0
    external = 0
    total = 0

    for anchor in soup.find_all("a", href=True):
        href = anchor["href"].strip()
        if not href:
            continue
        total += 1

        parsed = urlparse(href)
        if not parsed.netloc:
            internal += 1
        elif parsed.netloc == base_netloc:
            internal += 1
        else:
            external += 1

    return internal, external, total


def word_count_from_soup(soup: BeautifulSoup) -> int:
    for tag in soup(["script", "style", "noscript"]):
        tag.decompose()

    text = soup.get_text(separator=" ")
    text = re.sub(r"\s+", " ", text).strip()
    if not text:
        return 0
    return len(re.findall(r"\w+", text))


def scrape_url(url: str) -> Dict[str, object]:
    """Return a flat dict with scraped info for one URL."""
    result: Dict[str, object] = {
        "url": url,
        "final_url": None,
        "status_code": None,
        "error": None,
        "page_title": None,
        "meta_description": None,
        "meta_keywords": None,
        "robots_meta": None,
        "canonical_url": None,
        "og_title": None,
        "og_description": None,
        "og_image": None,
        "twitter_title": None,
        "twitter_description": None,
        "twitter_image": None,
        "last_modified": None,
        "content_type": None,
        "server": None,
        "h1_headings": None,
        "internal_link_count": None,
        "external_link_count": None,
        "total_link_count": None,
        "word_count": None,
        "all_meta_json": None,
    }

    normalized = normalize_url(url)
    if not normalized:
        result["error"] = "Empty URL"
        return result

    try:
        response = requests.get(
            normalized,
            headers=build_headers(),
            timeout=REQUEST_TIMEOUT_SECONDS,
            allow_redirects=True,
        )
    except requests.exceptions.RequestException as exc:
        result["error"] = f"Request failed: {exc.__class__.__name__}"
        return result

    result["status_code"] = response.status_code
    result["final_url"] = response.url
    result["last_modified"] = response.headers.get("Last-Modified")
    result["content_type"] = response.headers.get("Content-Type")
    result["server"] = response.headers.get("Server")

    if response.status_code != 200:
        result["error"] = f"Non-200 status: {response.status_code}"
        return result

    if not is_html_response(response):
        result["error"] = "Non-HTML response"
        return result

    # Safely decode HTML, even if compression/encoding is weird
    try:
        encoding = response.encoding or "utf-8"
        html_text = response.content.decode(encoding, errors="replace")
    except Exception as exc:
        result["error"] = f"Decode error: {exc.__class__.__name__}"
        return result

    try:
        soup = BeautifulSoup(html_text, "lxml")
    except Exception as exc:
        result["error"] = f"HTML parse error: {exc.__class__.__name__}"
        return result


    # <title>
    title_tag = soup.find("title")
    if title_tag:
        title_text = title_tag.get_text(strip=True)
        if title_text:
            result["page_title"] = title_text

    # canonical
    canonical_tag = soup.find("link", rel=lambda v: v and "canonical" in v.lower())
    if canonical_tag and canonical_tag.get("href"):
        result["canonical_url"] = canonical_tag["href"].strip()

    # basic meta
    result["meta_description"] = extract_meta(soup, name="description")
    result["meta_keywords"] = extract_meta(soup, name="keywords")
    result["robots_meta"] = extract_meta(soup, name="robots")

    # Open Graph
    result["og_title"] = extract_meta(soup, prop="og:title")
    result["og_description"] = extract_meta(soup, prop="og:description")
    result["og_image"] = extract_meta(soup, prop="og:image")

    # Twitter
    result["twitter_title"] = extract_meta(soup, name="twitter:title")
    result["twitter_description"] = extract_meta(soup, name="twitter:description")
    result["twitter_image"] = extract_meta(soup, name="twitter:image")

    # H1 headings
    h1_items: List[str] = []
    for h1 in soup.find_all("h1"):
        h1_text = h1.get_text(strip=True)
        if h1_text:
            h1_items.append(h1_text)
    if h1_items:
        result["h1_headings"] = " | ".join(h1_items)

    # link counts
    base_netloc = urlparse(result["final_url"] or normalized).netloc
    internal, external, total = count_links(soup, base_netloc)
    result["internal_link_count"] = internal
    result["external_link_count"] = external
    result["total_link_count"] = total

    # word count
    result["word_count"] = word_count_from_soup(soup)

    # all meta tags as JSON
    meta_list: List[Dict[str, str]] = []
    for tag in soup.find_all("meta"):
        content = tag.get("content")
        name = tag.get("name")
        prop = tag.get("property")
        if not content:
            continue
        entry: Dict[str, str] = {"content": content}
        if name:
            entry["name"] = name
        if prop:
            entry["property"] = prop
        meta_list.append(entry)

    try:
        result["all_meta_json"] = json.dumps(meta_list, ensure_ascii=False)
    except Exception:
        result["all_meta_json"] = json.dumps(meta_list, default=str)

    return result


class WebScraperApp:
    """Tkinter GUI wrapper (only place we use self.)."""

    def __init__(self) -> None:
        self.root = tk.Tk()
        self.root.title("Generic Web Scraper")
        self.root.geometry("650x220")
        self.root.columnconfigure(0, weight=1)

        self.input_path: Optional[str] = None
        self.output_path: Optional[str] = None

        self._build_ui()

    def _build_ui(self) -> None:
        self.select_input_button = tk.Button(
            self.root,
            text="Select input Excel file (with 'URL' column)",
            command=self.select_input_file,
        )
        self.select_input_button.grid(row=0, column=0, padx=10, pady=(20, 5), sticky="ew")

        self.input_label = tk.Label(self.root, text="No input file selected", anchor="w")
        self.input_label.grid(row=1, column=0, padx=10, pady=5, sticky="ew")

        # Output label (no button, just info)
        self.output_label = tk.Label(
            self.root,
            text="Output: will be created as 'output_webscraping.xlsx' next to input file",
            anchor="w",
        )
        self.output_label.grid(row=2, column=0, padx=10, pady=5, sticky="ew")

        self.status_label = tk.Label(self.root, text="Status: Idle", anchor="w")
        self.status_label.grid(row=3, column=0, padx=10, pady=5, sticky="ew")

        self.start_button = tk.Button(
            self.root,
            text="Start web scraping",
            command=self.start_scraping,
        )
        self.start_button.grid(row=4, column=0, padx=10, pady=(5, 20), sticky="ew")

    def select_input_file(self) -> None:
        file_path = filedialog.askopenfilename(
            title="Select input Excel file",
            filetypes=(("Excel files", "*.xlsx"),),
        )
        if not file_path:
            return

        self.input_path = file_path
        self.input_label.config(text=f"Input: {file_path}")

        input_folder = os.path.dirname(file_path)
        self.output_path = os.path.join(input_folder, "output_webscraping.xlsx")
        self.output_label.config(text=f"Output: {self.output_path}")

    def start_scraping(self) -> None:
        if not self.input_path:
            messagebox.showwarning(
                "Missing input",
                "Please select an input Excel file with a 'URL' column.",
            )
            return

        if not self.output_path:
            # Safety; should always be set when input is chosen
            input_folder = os.path.dirname(self.input_path)
            self.output_path = os.path.join(input_folder, "output_webscraping.xlsx")
            self.output_label.config(text=f"Output: {self.output_path}")

        try:
            self.status_label.config(text="Status: Reading input...")
            self.root.update_idletasks()

            df = pd.read_excel(self.input_path)

            if "URL" not in df.columns:
                raise ValueError("Input Excel must contain a 'URL' column.")

            results: List[Dict[str, object]] = []
            total_rows = len(df)

            for idx, row in df.iterrows():
                url_value = str(row.get("URL", "")).strip()
                if not url_value or url_value.lower() in ("nan", "none"):
                    result = {
                        "url": url_value,
                        "final_url": None,
                        "status_code": None,
                        "error": "Missing URL in input row",
                    }
                else:
                    result = scrape_url(url_value)

                results.append(result)

                self.status_label.config(
                    text=f"Status: Processed {idx + 1}/{total_rows} URLs..."
                )
                self.root.update_idletasks()

            self.status_label.config(text="Status: Writing output...")
            self.root.update_idletasks()

            out_df = pd.DataFrame(results)
            out_df.to_excel(self.output_path, index=False)

            self.status_label.config(text="Status: Completed")
            messagebox.showinfo(
                "Web scraping completed",
                f"Web scraping completed.\nResults saved to:\n{self.output_path}",
            )
        except Exception as exc:
            self.status_label.config(text="Status: Failed")
            messagebox.showerror("Web scraping failed", f"An error occurred:\n{exc}")

    def run(self) -> None:
        self.root.mainloop()


def main() -> None:
    app = WebScraperApp()
    app.run()


if __name__ == "__main__":
    main()
