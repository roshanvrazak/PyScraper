#!/usr/bin/env python3
"""
UK Gov Visa Sponsorship Pre-Scrape Enrichment Script (enrich_targets.py)
-----------------------------------------------------------------------
Author: Antigravity AI
Description:
    Reads a UK Gov visa sponsorship file (.csv or .xlsx), filters tech-focused
    organizations, resolves their official homepage and careers URL using
    a robust, asynchronous search engine fallback, and saves the output to a
    state-aware tracking JSON file.
"""

import os
import sys
import json
import asyncio
import argparse
from datetime import datetime
from urllib.parse import urlparse
import logging
import random
import re
from typing import List, Dict, Any, Optional

# Third-party imports (soft-check & import)
try:
    import pandas as pd
except ImportError:
    print("[ERROR] pandas is required. Run: pip install pandas", file=sys.stderr)
    sys.exit(1)

try:
    import httpx
except ImportError:
    print("[ERROR] httpx is required for async requests. Run: pip install httpx", file=sys.stderr)
    sys.exit(1)

try:
    from bs4 import BeautifulSoup
except ImportError:
    print("[ERROR] beautifulsoup4 is required. Run: pip install beautifulsoup4", file=sys.stderr)
    sys.exit(1)

# Setup Logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.StreamHandler(sys.stdout),
        logging.FileHandler("enrich_targets.log", encoding="utf-8")
    ]
)
logger = logging.getLogger("enrich_targets")

# Tech filtering keywords
TECH_KEYWORDS = ["Software", "Tech", "Information Technology", "Computer", "AI", "Data", "Consulting"]

# Desktop User-Agent pool to mimic real browsers
USER_AGENTS = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:121.0) Gecko/20100101 Firefox/121.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.2 Safari/605.1.15",
]

def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Filter UK Gov Visa Sponsorship file and enrich it with career URLs."
    )
    parser.add_argument(
        "--file", "-f",
        default="dummy_sponsors.csv",
        help="Path to the UK Gov visa sponsorship file (.csv or .xlsx)."
    )
    parser.add_argument(
        "--output", "-o",
        default="curated_targets.json",
        help="Path to save the output JSON tracking file."
    )
    parser.add_argument(
        "--limit", "-l",
        type=int,
        default=None,
        help="Limit search lookups to N filtered tech companies (helpful for debugging/dry-runs)."
    )
    parser.add_argument(
        "--concurrency", "-c",
        type=int,
        default=2,
        help="Max concurrent search requests (keep it low to prevent rate limiting)."
    )
    parser.add_argument(
        "--delay", "-d",
        type=float,
        default=1.5,
        help="Base delay (in seconds) between requests to prevent bot detection."
    )
    return parser.parse_args()


def load_and_filter_companies(file_path: str) -> List[str]:
    """Loads CSV/Excel file, finds the company/organisation column, and filters based on tech keywords."""
    if not os.path.exists(file_path):
        logger.error(f"Sponsorship file not found: {file_path}")
        sys.exit(1)
        
    logger.info(f"Reading file: {file_path}")
    ext = os.path.splitext(file_path)[1].lower()
    
    try:
        if ext == ".csv":
            df = pd.read_csv(file_path)
        elif ext in [".xlsx", ".xls"]:
            df = pd.read_excel(file_path)
        else:
            logger.error(f"Unsupported file format '{ext}'. Must be .csv, .xls, or .xlsx")
            sys.exit(1)
    except Exception as e:
        logger.error(f"Failed to read file: {e}")
        sys.exit(1)
        
    logger.info(f"Loaded {len(df)} total rows from file.")
    
    # Resilient Column detection
    name_col = None
    possible_cols = ["organisation name", "organization name", "company name", "company", "organisation", "organization", "name"]
    for col in df.columns:
        if str(col).lower().strip() in possible_cols:
            name_col = col
            break
            
    if not name_col:
        # Fallback to the first column
        name_col = df.columns[0]
        logger.warning(f"Could not explicitly detect company name column. Falling back to first column: '{name_col}'")
    else:
        logger.info(f"Detected company name column: '{name_col}'")
        
    # Standardize names and drop duplicates/empty rows
    df[name_col] = df[name_col].astype(str).str.strip()
    df = df[df[name_col] != ""].drop_duplicates(subset=[name_col])
    
    # Filter by tech-focused keywords (case-insensitive)
    regex_pattern = "|".join([re.escape(kw) for kw in TECH_KEYWORDS])
    tech_mask = df[name_col].str.contains(regex_pattern, case=False, na=False)
    filtered_df = df[tech_mask]
    
    filtered_companies = filtered_df[name_col].tolist()
    logger.info(f"Filtered down to {len(filtered_companies)} tech-focused companies matching keywords: {TECH_KEYWORDS}")
    return filtered_companies


async def search_company_urls_fallback(client: httpx.AsyncClient, company_name: str, delay: float) -> Optional[str]:
    """
    Asynchronous fallback scraper querying DuckDuckGo HTML interface.
    Extracts the first external non-advertisement result.
    """
    # Introduce random jitter to prevent uniform request timings
    await asyncio.sleep(delay + random.uniform(0.5, 2.0))
    
    query = f"{company_name} UK careers"
    url = f"https://html.duckduckgo.com/html/?q={httpx.QueryParams({'q': query})}"
    
    headers = {
        "User-Agent": random.choice(USER_AGENTS),
        "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
        "Accept-Language": "en-US,en;q=0.5",
        "Referer": "https://duckduckgo.com/",
    }
    
    max_retries = 3
    backoff = 2.0
    
    for attempt in range(1, max_retries + 1):
        try:
            response = await client.get(url, headers=headers, timeout=15.0)
            
            if response.status_code == 200:
                soup = BeautifulSoup(response.text, "html.parser")
                # DuckDuckGo HTML results are inside links with class result__url or similar
                links = soup.find_all("a", class_="result__url")
                
                valid_urls = []
                for link in links:
                    href = link.get("href", "")
                    if href:
                        # Extract the actual URL (DuckDuckGo sometimes encodes it or embeds in custom link redirect)
                        # Typical format is: /l/?kh=-1&uddg=https%3A%2F%2Fcareers.google.com%2F
                        if "/l/?uddg=" in href:
                            parsed_qs = httpx.QueryParams(href.split("?")[-1])
                            actual_url = parsed_qs.get("uddg")
                            if actual_url:
                                href = actual_url
                        
                        # Filter out internal duckduckgo or ad domains
                        domain = urlparse(href).netloc.lower()
                        if domain and "duckduckgo.com" not in domain:
                            valid_urls.append(href)
                            
                if valid_urls:
                    # Return the top non-ad search result
                    return valid_urls[0]
                else:
                    # Let's try parsing broad anchors inside results if result__url wasn't found
                    anchors = soup.find_all("a", class_="result__snippet")
                    for anchor in anchors:
                        href = anchor.get("href", "")
                        if href and "/l/?uddg=" in href:
                            parsed_qs = httpx.QueryParams(href.split("?")[-1])
                            actual_url = parsed_qs.get("uddg")
                            if actual_url:
                                href = actual_url
                        domain = urlparse(href).netloc.lower()
                        if domain and "duckduckgo.com" not in domain:
                            return href
                            
                logger.warning(f"No valid external links found on DDG for company: '{company_name}'")
                return None
                
            elif response.status_code in [429, 503]:
                logger.warning(f"DuckDuckGo throttled search for '{company_name}' (Status {response.status_code}). Backoff {backoff}s before retry.")
                await asyncio.sleep(backoff)
                backoff *= 2
            else:
                logger.error(f"Failed to query DDG for '{company_name}' (Status {response.status_code})")
                return None
                
        except (httpx.RequestError, asyncio.TimeoutError) as e:
            logger.warning(f"Request error querying DDG for '{company_name}' on attempt {attempt}: {e}")
            if attempt == max_retries:
                logger.error(f"Failed to resolve '{company_name}' after {max_retries} attempts.")
                return None
            await asyncio.sleep(backoff)
            backoff *= 2
            
    return None


async def search_company_urls_serp(client: httpx.AsyncClient, company_name: str, api_key: str) -> Optional[str]:
    """Production alternative using SerpAPI (Google Search API) if provided."""
    url = "https://serpapi.com/search"
    params = {
        "engine": "google",
        "q": f"{company_name} UK careers",
        "api_key": api_key,
        "num": 3
    }
    try:
        response = await client.get(url, params=params, timeout=10.0)
        if response.status_code == 200:
            data = response.json()
            organic_results = data.get("organic_results", [])
            if organic_results:
                return organic_results[0].get("link")
        else:
            logger.error(f"SerpAPI returned error {response.status_code}: {response.text}")
    except Exception as e:
        logger.error(f"SerpAPI connection failed for '{company_name}': {e}")
    return None


async def process_company(
    client: httpx.AsyncClient,
    company: str,
    semaphore: asyncio.Semaphore,
    delay: float,
    serpapi_key: Optional[str]
) -> Optional[Dict[str, Any]]:
    """Coordinates search lookup and URL normalization for a single company."""
    async with semaphore:
        logger.info(f"Discovering career pages for: '{company}'")
        
        career_url = None
        if serpapi_key:
            career_url = await search_company_urls_serp(client, company, serpapi_key)
            
        if not career_url:
            career_url = await search_company_urls_fallback(client, company, delay)
            
        if not career_url:
            logger.warning(f"Skipping '{company}': Career URL could not be resolved.")
            return None
            
        # Standardize the base URL from career URL
        parsed_url = urlparse(career_url)
        base_url = f"{parsed_url.scheme}://{parsed_url.netloc}"
        
        logger.info(f"Resolved! '{company}' -> Base: {base_url} | Careers: {career_url}")
        
        return {
            "company_name": company,
            "base_url": base_url,
            "career_page_url": career_url,
            "date_added": datetime.utcnow().isoformat() + "Z"
        }


async def enrich_pipeline(args: argparse.Namespace):
    # Load state/already scraped targets to enable resuming
    existing_targets = {}
    if os.path.exists(args.output):
        try:
            with open(args.output, "r", encoding="utf-8") as f:
                data = json.load(f)
                # Map company_name -> dict entry
                existing_targets = {item["company_name"]: item for item in data if "company_name" in item}
            logger.info(f"Resuming pipeline. Loaded {len(existing_targets)} already-enriched targets from '{args.output}'")
        except Exception as e:
            logger.warning(f"Could not load existing file '{args.output}' for resume: {e}. Starting fresh.")

    # Filter tech-focused sponsors
    companies = load_and_filter_companies(args.file)
    
    # Subtract already enriched items to skip them
    companies_to_process = [c for c in companies if c not in existing_targets]
    logger.info(f"{len(companies) - len(companies_to_process)} companies already resolved. {len(companies_to_process)} need lookup.")
    
    if args.limit:
        companies_to_process = companies_to_process[:args.limit]
        logger.info(f"Applying lookup limit of {args.limit}. Active batch: {len(companies_to_process)}")
        
    if not companies_to_process:
        logger.info("No new companies to resolve. Exiting.")
        # Ensure we write out what we already had
        with open(args.output, "w", encoding="utf-8") as f:
            json.dump(list(existing_targets.values()), f, indent=4)
        return

    # Check for SerpAPI credentials in environment variables
    serpapi_key = os.environ.get("SERPAPI_API_KEY")
    if serpapi_key:
        logger.info("Found SERPAPI_API_KEY in environment variables. Using SerpAPI Google engine.")
    else:
        logger.info("No SERPAPI_API_KEY detected. Utilizing resilient DuckDuckGo Lite/HTML fallback.")

    semaphore = asyncio.Semaphore(args.concurrency)
    
    async with httpx.AsyncClient(follow_redirects=True) as client:
        tasks = [
            process_company(client, company, semaphore, args.delay, serpapi_key)
            for company in companies_to_process
        ]
        
        # Gather search results
        results = await asyncio.gather(*tasks)
        
        # Filter successful enrichments and merge with existing
        new_enrichments = [r for r in results if r is not None]
        for item in new_enrichments:
            existing_targets[item["company_name"]] = item
            
    # Save the consolidated file
    with open(args.output, "w", encoding="utf-8") as f:
        json.dump(list(existing_targets.values()), f, indent=4)
        
    logger.info(f"Pipeline enrichment complete! Successfully saved {len(existing_targets)} targets to '{args.output}'")


def main():
    args = parse_arguments()
    try:
        asyncio.run(enrich_pipeline(args))
    except KeyboardInterrupt:
        logger.info("\nEnrichment pipeline interrupted by user. Saved records remain intact.")


if __name__ == "__main__":
    main()
