# UK Gov Visa Sponsorship Tech Job Scraping Pipeline
### Operations & Execution Guide (Ubuntu Linux Terminal)

This guide provides step-by-step instructions on setting up, verifying, and executing the complete, decoupled scraping pipeline inside your **Ubuntu Linux Terminal**.

---

## 📂 Project Architecture

The pipeline consists of two decoupled scripts running sequentially to isolate intensive network searches from direct scraping:
```text
job_scraper_pipeline/
├── enrich_targets.py          # Stage 1: Filter UK Gov sponsors & discover domains (Async)
├── dummy_sponsors.csv         # Mock sponsor data for automated dry-runs
├── README.md                  # This documentation guide
├── scrapy.cfg                 # Scrapy project configuration
└── job_spider_project/        # Stage 2 & 3: Scraping & Excel Export Pipeline
    ├── __init__.py
    ├── items.py               # Job scraping data schema definition
    ├── middlewares.py         # Custom User-Agent rotation middleware
    ├── pipelines.py           # openpyxl pipeline (cleaning, deduplication, styling)
    ├── settings.py            # Autothrottle, delays, concurrency configuration
    └── spiders/
        ├── __init__.py
        └── job_spider.py      # Core JobHunterSpider spider
```

---

## 🛠️ Step 1: System Environment Setup

Open your Ubuntu terminal and install the required Python environment packages:

```bash
# 1. Update package list and install virtual environment dependencies
sudo apt update
sudo apt install python3-pip python3-venv -y

# 2. Navigate to your project directory (or create a dedicated working space)
# Recommendation: Set your active workspace to this directory
cd path/to/job_scraper_pipeline

# 3. Create a clean Python Virtual Environment (venv)
python3 -m venv venv

# 4. Activate the virtual environment
source venv/bin/activate
```

---

## 📦 Step 2: Install Python Dependencies

With your virtual environment activated (`(venv)` should be prefixed in your terminal prompt), run the following command to install the required packages:

```bash
# Install core dependencies: Pandas, BeautifulSoup4, HTTPX, Scrapy and OpenPyXL
pip install --upgrade pip
pip install pandas openpyxl scrapy httpx beautifulsoup4
```

*Verification check: Ensure all libraries are installed successfully with `pip list`.*

---

## 🚀 Step 3: Run Stage 1 — Pre-Scrape Enrichment

This script reads your local UK Gov sponsorship register (supporting `.csv`, `.xls`, and `.xlsx` files), filters out non-tech organizations based on target keywords, and runs an asynchronous lookup to resolve their careers URLs.

### 📋 Where to get the official UK Gov Visa Sponsorship list?
You can download the live official visa sponsorship dataset directly from the UK Gov website in CSV format:
* **Download URL:** [UK Gov Register of Licensed Sponsors: Workers](https://www.gov.uk/government/publications/register-of-licensed-sponsors-workers)

### 🏃‍♂️ Running the Enrichment Script

#### A. Run a Dry-Run Verification (using mock data)
To verify everything is working perfectly, run the script against the included mock dataset and limit the lookup to 5 companies:
```bash
python3 enrich_targets.py --file dummy_sponsors.csv --limit 5 --output curated_targets.json
```

#### B. Run on the Actual UK Gov Sponsorship File
To run on your actual full sponsorship file (e.g. `sponsors.csv`) loaded into the directory:
```bash
python3 enrich_targets.py --file sponsors.csv --limit 50 --output curated_targets.json
```
*(Tip: Performing organic lookups for 50,000+ companies will hit search limits. We strongly recommend using the `--limit` option to scrape in batches, or registering a **SerpAPI** key to handle unlimited searches).*

#### C. Production Running with SerpAPI Key (Optional)
If you have a Google Search API key (SerpAPI) to prevent any DuckDuckGo throttling blocks during massive runs:
```bash
export SERPAPI_API_KEY="your_serp_api_key_here"
python3 enrich_targets.py --file sponsors.csv --limit 200 --output curated_targets.json
```

*Output:* This script produces `curated_targets.json`. You can verify its content with `cat curated_targets.json`. It will look like this:
```json
[
    {
        "company_name": "Google UK Limited",
        "base_url": "https://careers.google.com",
        "career_page_url": "https://careers.google.com/jobs/",
        "date_added": "2026-05-31T11:50:00Z"
    }
]
```

*Resilience Note:* The script saves state. If interrupted, running it again will **skip** already resolved companies, protecting your search limits.

---

## 🕷️ Step 4: Run Stage 2 & 3 — Scrapy Job crawling & Excel Export

Once `curated_targets.json` is generated, run the Scrapy spider to scan these portals, identify tech job opportunities, filter out visa exclusions, and output an elegantly styled Excel workbook.

### 🏃‍♂️ Run the Scrapy Crawl command:
```bash
# Execute the spider from the project root directory
scrapy crawl JobHunter
```

### 🔍 What happens during the scrape?
1. **Dynamic Target Crawling:** `JobHunterSpider` starts and loads target companies from `curated_targets.json`.
2. **Polite Crawling:** The crawler rotates browser User-Agents for each site, enforces a `2.0` second basic politeness delay, and dynamically activates the **AutoThrottle** system based on target servers load.
3. **Keyword Matching & Exclusion Detection:** It visits job detail links, extracts title and location, extracts salary through regex patterns, and checks if visa exclusions like `"no sponsorship"`, `"must possess local visa"`, or `"indigenous only"` appear in the text.
4. **Resilient Failure Recovery:** If a domain returns a `403 Forbidden`, `404 Not Found`, or times out, it is handled as a soft error. The system records it inside `scraping_errors.log` and proceeds to the next target without interrupting the crawler queue.
5. **Stage 3 Pipeline:** Items flow into `ExcelExportPipeline`, where duplicates are filtered by URL, whitespaces are trimmed, and records are sorted by Company Name.
6. **Excel Generation:** The pipeline writes records into a highly curated spreadsheet named `curated_sponsored_jobs.xlsx`.

---

## 📊 Step 5: Inspecting Outputs & Diagnostics

### 1. The Curated Spreadsheet (`curated_sponsored_jobs.xlsx`)
This file is generated in your current working directory. You can copy it to your local machine and open it in Excel, LibreOffice, or Google Sheets. It is pre-styled with:
* **Auto-Filters:** Enabled on all headers so you can easily filter by company, location, or visa exclusion.
* **Bold Headers:** Styled in elegant **Steel Navy (`#1F4E78`)** with white text.
* **Smart Zebra-Striping:** Interchanging light-blue rows for perfect reading scanner visibility.
* **Dynamic Column Widths:** Automatically expanded to prevent data-clipping.
* **Exclusion Flags:** Cells in "Visa Exclusions Found?" are colored **Light Red (`#FFC7CE`)** with bold red text if a restriction keyword was detected, or **Light Green (`#C6EFCE`)** if clean.

### 2. Error Diagnostics (`scraping_errors.log`)
If any website blocked the crawler or timed out, details are logged here. Run this command to inspect:
```bash
cat scraping_errors.log
```
*Example entry:*
```text
[2026-05-31 11:52:10] Company: 'SoftTech Solutions Ltd'
  Requested URL: https://softtech.example/careers
  Error Type: HttpError
  Details: 403 Forbidden
------------------------------------------------------------
```

---

## 💡 Troubleshooting & Production Customizations

* **Changing Delay Speed:** If you need faster scraping and are running against targets without strict firewalls, you can decrease the delay in `job_spider_project/settings.py` (e.g. `DOWNLOAD_DELAY = 1.0`).
* **Custom Job Keywords:** You can expand the target role matching rules by editing the `JOB_KEYWORDS` array inside `job_spider_project/spiders/job_spider.py` (e.g., adding "Security analyst", "System Administrator").
* **Resetting Scraped Cache:** Delete `curated_targets.json` to start the pre-scrape discovery phase from scratch, or delete the `.xlsx` file before running the spider to generate fresh listings.
