# Automated Corporate Lead Enrichment & Resilient Web-Crawling Pipeline
### Systems & Execution Operations Guide (Ubuntu Linux Terminal)

This repository provides an enterprise-grade, multi-stage data pipeline designed to ingest organization registries, asynchronously enrich target records with verified web resources, politely crawl remote career portals with anti-bot resistance, and output deduplicated, professionally formatted analytics dashboards into Microsoft Excel workbook sheets.

---

## 📂 System Architecture & Components

The pipeline consists of decoupled components operating sequentially to isolate heavy search queries from intensive portal parsing:
```text
job_scraper_pipeline/
├── enrich_targets.py          # Stage 1: Async target directory filter & domain discovery
├── dummy_sponsors.csv         # Mock target organization list for automated dry-runs
├── README.md                  # Comprehensive operations guide
├── scrapy.cfg                 # Central Scrapy project settings
└── job_spider_project/        # Stage 2 & 3: Resilient crawler & styled Excel pipeline
    ├── __init__.py
    ├── items.py               # Job scraping data schema model definition
    ├── middlewares.py         # Custom User-Agent rotation middleware
    ├── pipelines.py           # openpyxl excel reporter (deduplication, formatting, cell-styling)
    ├── settings.py            # Autothrottle, delays, concurrent requests configuration
    └── spiders/
        ├── __init__.py
        └── job_spider.py      # Core JobHunterSpider crawling engine
```

---

## 🛠️ Step 1: System Environment Setup

Open your terminal and install the core Python execution packages:

```bash
# 1. Update package list and install virtual environment dependencies
sudo apt update
sudo apt install python3-pip python3-venv -y

# 2. Navigate to your active project workspace directory
cd path/to/job_scraper_pipeline

# 3. Create a clean Python Virtual Environment (venv)
python3 -m venv venv

# 4. Activate the virtual environment
source venv/bin/activate
```

---

## 📦 Step 2: Install Python Dependencies

With your virtual environment activated, upgrade pip and install the required libraries:

```bash
# Install core dependencies: Pandas, BeautifulSoup4, HTTPX, Scrapy and OpenPyXL
pip install --upgrade pip
pip install pandas openpyxl scrapy httpx beautifulsoup4
```

*Verification check: Ensure all libraries are installed successfully by running `pip list`.*

---

## 🚀 Step 3: Run Stage 1 — Pre-Scrape Async Target Enrichment

The enrichment utility takes a local organization directory (supporting `.csv`, `.xls`, and `.xlsx` formats), filters for target sectors (e.g., tech-focused entities) based on keywords, and runs a high-performance **asynchronous lookup** to resolve homepages and career portal URLs.

### 🏃‍♂️ Running the Enrichment Script

#### A. Execute a Dry-Run Verification (using mock data)
Verify the setup using the included mock directory, limiting the lookup to 5 companies:
```bash
python3 enrich_targets.py --file dummy_sponsors.csv --limit 5 --output curated_targets.json
```

#### B. Execute on Full Target Directories
To run on your actual full target company listing file (e.g., `companies.csv`):
```bash
python3 enrich_targets.py --file companies.csv --limit 50 --output curated_targets.json
```
*(Tip: Lookups for very large directories can hit search engine rate-limits. We recommend using the `--limit` option to process targets in batches, or registering a SerpAPI key to execute high-volume lookups).*

#### C. Production Execution with SerpAPI (Optional)
If you have a Google Search API key (SerpAPI) to handle large enterprise runs:
```bash
export SERPAPI_API_KEY="your_serp_api_key_here"
python3 enrich_targets.py --file companies.csv --limit 200 --output curated_targets.json
```

*Expected Output:* The script generates a consolidated tracking sheet `curated_targets.json`. It will look like this:
```json
[
    {
        "company_name": "Google UK Limited",
        "base_url": "https://careers.google.com",
        "career_page_url": "https://careers.google.com/jobs/",
        "date_added": "2026-05-31T12:00:00Z"
    }
]
```

*Resilience Feature:* The script is fully **state-aware**. If interrupted, running it again will automatically skip previously resolved companies, preserving your search quotas.

---

## 🕷️ Step 4: Run Stage 2 & 3 — Polite Scrapy Crawler & Excel Pipeline

Once `curated_targets.json` is generated, execute the Scrapy crawler. The spider politely scans the resolved career portals, extracts active tech opportunities, filters roles based on custom restriction keywords, and outputs a highly polished Excel spreadsheet.

### 🏃‍♂️ Run the Scrapy Crawl command:
```bash
# Execute the spider from the project root directory
scrapy crawl JobHunter
```

### 🔍 Engineering Details of the Scrape Phase:
1. **Dynamic Target Allocation:** `JobHunterSpider` starts, loading seed portals from the enriched `curated_targets.json`.
2. **Anti-Bot & Polite Crawling:** The crawler rotates browser User-Agents for each destination, enforces a `2.0` second safety delay, and leverages Scrapy's dynamic **AutoThrottle** system to adjust speed in real time according to remote server load.
3. **Keyword Matching & Pattern Detection:** It parses listing pages to identify technology keywords, extracts salary packages using robust regular expressions, and checks for custom exclusion keywords (e.g., checking if roles have specific localization requirements or restrictions).
4. **Soft-Error Isolation:** If a domain returns a `403 Forbidden`, `404 Not Found`, or times out, it is handled gracefully as a soft error. The failure is recorded in `scraping_errors.log` and the crawler moves to the next company in the queue without crashing.
5. **Excel Export Pipeline:** The items are passed into the `ExcelExportPipeline`, where they are filtered for duplicates by URL, stripped of redundant white spaces, and sorted alphabetically by Company Name.
6. **Excel Generation:** The pipeline writes records into a highly-styled corporate sheet named `curated_job_opportunities.xlsx`.

---

## 📊 Step 5: Inspecting Outputs & Diagnostics

### 1. The Curated Spreadsheet (`curated_job_opportunities.xlsx`)
This file is generated in your workspace root. You can open it in Microsoft Excel, LibreOffice, or Google Sheets. It features professional styling:
* **Interactive Headers:** Formatted in elegant **Steel Navy (`#1F4E78`)** with bold white text and active auto-filtering enabled across all columns.
* **Smart Zebra-Striping:** Interchanging light-blue rows (`#F2F6FA`) for enhanced scanner readability.
* **Exclusion Flags:** Cells in the "Exclusions Found?" column are dynamically color-coded: **Light Red (`#FFC7CE`)** with bold red text if a restriction keyword was triggered, and **Light Green (`#C6EFCE`)** with dark green text if clean.
* **Auto-Adjusted Widths:** Columns are automatically expanded dynamically based on content length to prevent clipping.

### 2. Error Diagnostics (`scraping_errors.log`)
Any connection issues, timeouts, or access limits encountered are isolated here for diagnostic review:
```bash
cat scraping_errors.log
```
*Example Entry:*
```text
[2026-05-31 12:05:10] Company: 'SoftTech Solutions Ltd'
  Requested URL: https://softtech.example/careers
  Error Type: HttpError
  Details: 403 Forbidden
------------------------------------------------------------
```

---

## 💡 Customizations & Custom Extensions

* **Modifying Crawling Speeds:** To change scrape concurrency or speed, adjust parameters in `job_spider_project/settings.py` (e.g. lowering `DOWNLOAD_DELAY = 1.0` or raising concurrency settings).
* **Adding Custom Exclusion Keywords:** Customize the exclusion filter list by modifying the `EXCLUSION_KEYWORDS` array inside `job_spider_project/spiders/job_spider.py` (e.g. adding specific tools, frameworks, or experience tiers).
* **Resetting Cache:** Simply delete `curated_targets.json` to trigger target discovery from scratch, or delete the `.xlsx` file before running the spider to produce a completely fresh workbook.
