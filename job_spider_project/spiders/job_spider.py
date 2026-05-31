# -*- coding: utf-8 -*-
"""
Corporate Lead Tech Job Scraper Spider (job_spider.py)
-------------------------------------------------------
Author: Antigravity AI
Description:
    Core Scrapy spider that reads targets from curated_targets.json, crawls
    them politely, extracts tech roles, parses details (salary, location),
    checks for custom exclusion keywords, and gracefully records soft errors.
"""

import os
import json
import re
from datetime import datetime
import scrapy
from urllib.parse import urlparse
from job_spider_project.items import JobItem

# Job role target keywords (case-insensitive)
JOB_KEYWORDS = [
    "software engineer", "developer", "data engineer", "devops",
    "cloud architect", "computer science", "backend", "frontend"
]

# Exclusion search keywords for role/location restriction warning flags
EXCLUSION_KEYWORDS = ["visa", "sponsor", "indigenous"]

# Common UK Tech Cities for location fallback scanner
UK_CITIES = [
    "london", "cambridge", "manchester", "bristol", "edinburgh",
    "birmingham", "leeds", "glasgow", "reading", "oxford", "belfast"
]


class JobHunterSpider(scrapy.Spider):
    name = "JobHunter"
    
    def __init__(self, targets_file="curated_targets.json", *args, **kwargs):
        super(JobHunterSpider, self).__init__(*args, **kwargs)
        self.targets_file = targets_file
        self.error_log_file = "scraping_errors.log"
        
    def start_requests(self):
        """Reads target list from JSON file and schedules initial crawls."""
        if not os.path.exists(self.targets_file):
            self.logger.error(f"Curated targets file not found: {self.targets_file}. Please run enrich_targets.py first.")
            return
            
        try:
            with open(self.targets_file, "r", encoding="utf-8") as f:
                targets = json.load(f)
        except Exception as e:
            self.logger.error(f"Failed to parse targets file: {e}")
            return
            
        self.logger.info(f"Loaded {len(targets)} targets from '{self.targets_file}'. Starting job search...")
        
        for target in targets:
            company_name = target.get("company_name")
            career_url = target.get("career_page_url")
            base_url = target.get("base_url")
            
            if not career_url or not company_name:
                continue
                
            # Schedule request with errback for resilient soft-error logging
            yield scrapy.Request(
                url=career_url,
                callback=self.parse_career_page,
                errback=self.handle_error,
                meta={
                    "company_name": company_name,
                    "base_url": base_url,
                    "career_page_url": career_url,
                    "download_timeout": 15.0
                }
            )

    def parse_career_page(self, response):
        """Scans the main career landing page for links matching tech job roles."""
        company_name = response.meta["company_name"]
        self.logger.info(f"Scanning career portal for company: '{company_name}' at {response.url}")
        
        # Track spawned requests per company to prevent getting trapped in huge listings
        job_links_found = 0
        max_links_per_company = 15
        
        # Select all anchor elements
        anchors = response.xpath("//a")
        for anchor in anchors:
            text = anchor.xpath("string(.)").get("").strip()
            href = anchor.xpath("@href").get("").strip()
            
            if not href or href.startswith(("#", "javascript:", "mailto:", "tel:")):
                continue
                
            # Check if anchor text contains any target job keywords
            text_lower = text.lower()
            is_tech_job = any(keyword in text_lower for keyword in JOB_KEYWORDS)
            
            if is_tech_job:
                # Basic cleaning of extra whitespace in title
                job_title = re.sub(r'\s+', ' ', text).strip()
                
                # Exclude links that are clearly page navigations or filters
                if len(job_title) > 80 or len(job_title) < 5:
                    continue
                    
                resolved_url = response.urljoin(href)
                job_links_found += 1
                
                if job_links_found > max_links_per_company:
                    self.logger.warning(f"Reached limit of {max_links_per_company} job listings followed for '{company_name}'. Skipping remainder.")
                    break
                    
                self.logger.info(f"Found tech job matching: '{job_title}' -> Scheduling detail parse.")
                
                yield scrapy.Request(
                    url=resolved_url,
                    callback=self.parse_job_detail,
                    errback=self.handle_error,
                    meta={
                        "company_name": company_name,
                        "job_title": job_title,
                        "application_url": resolved_url,
                        "download_timeout": 12.0
                    }
                )
                
        # Fallback: If no links were resolved, scan page content directly to see if it's a single static board
        if job_links_found == 0:
            self.logger.warning(f"No discrete job detail links discovered on career page for '{company_name}'. Checking body content...")
            yield from self.parse_static_content(response)

    def parse_job_detail(self, response):
        """Extracts job posting details from job description landing pages."""
        company_name = response.meta["company_name"]
        job_title = response.meta["job_title"]
        application_url = response.meta["application_url"]
        
        # 1. Extract body text for keyword scans
        body_elements = response.xpath("//body//text()").getall()
        body_text = " ".join([text.strip() for text in body_elements if text.strip()])
        body_text_lower = body_text.lower()
        
        # 2. Flag custom exclusion keywords (e.g. restriction markers)
        exclusion_found = any(word in body_text_lower for word in EXCLUSION_KEYWORDS)
        
        # 3. Parse Salary Range via resilient Regex
        # Matches patterns like £45,000 - £55,000, £50k - £60k, £70k per annum, $80,000, etc.
        salary_pattern = re.compile(
            r'(?:£|\$|€)\s*\d+(?:,\d+)*(?:\s*(?:k|K|thousand))?(?:\s*(?:-|to)\s*(?:£|\$|€)?\s*\d+(?:,\d+)*(?:\s*(?:k|K|thousand))?)?',
            re.IGNORECASE
        )
        salaries = salary_pattern.findall(body_text)
        
        # Filter matching numbers to find salary range (excluding trivial digits)
        salary_range = "Not Specified"
        for s in salaries:
            clean_s = s.strip()
            # Eliminate short values like single currencies or zip matches
            if len(clean_s) > 2:
                # Add contextual suffixes if present in nearby text
                context_idx = body_text.find(clean_s)
                suffix = ""
                if context_idx != -1:
                    near_text = body_text[context_idx:context_idx+50].lower()
                    if "annum" in near_text or "year" in near_text or "/yr" in near_text:
                        suffix = " per annum"
                    elif "hour" in near_text or "/hr" in near_text:
                        suffix = " per hour"
                salary_range = clean_s + suffix
                break
                
        # 4. Parse Location
        location = "Not Specified"
        
        # A. Check schema meta or standard class indicators
        loc_xpath = (
            "//*[@itemprop='jobLocation']//*[contains(@class, 'locality') or contains(@class, 'address')]/text() | "
            "//*[contains(@class, 'location') or contains(@id, 'location')]/text() | "
            "//meta[@name='location']/@content"
        )
        loc_candidates = response.xpath(loc_xpath).getall()
        for loc in loc_candidates:
            clean_loc = loc.strip()
            if clean_loc and len(clean_loc) < 40 and not any(x in clean_loc.lower() for x in ["css", "theme", "logo"]):
                location = clean_loc
                break
                
        # B. Fallback: Scan text using Lexicon for key UK cities or Remote keywords
        if location == "Not Specified":
            if "remote" in body_text_lower:
                location = "Remote (UK)"
            elif "hybrid" in body_text_lower:
                location = "Hybrid (UK)"
            else:
                for city in UK_CITIES:
                    if re.search(r'\b' + re.escape(city) + r'\b', body_text_lower):
                        location = city.title()
                        break
                        
        if location == "Not Specified":
            location = "United Kingdom"  # Default generic fallback
            
        # Instantiate item
        item = JobItem()
        item["company_name"] = company_name
        item["job_title"] = job_title
        item["location"] = location
        item["salary_range"] = salary_range
        item["application_url"] = application_url
        item["exclusion_keyword_found"] = exclusion_found
        item["date_scraped"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        yield item

    def parse_static_content(self, response):
        """Fallback parses raw page text blocks for direct job listing listings."""
        company_name = response.meta["company_name"]
        
        # Scrape and search headings and bold lines
        elements = response.xpath("//h1 | //h2 | //h3 | //h4 | //strong | //b | //li")
        for elem in elements:
            text = elem.xpath("string(.)").get("").strip()
            text_lower = text.lower()
            
            # Match keywords
            if any(keyword in text_lower for keyword in JOB_KEYWORDS) and len(text) < 70 and len(text) > 6:
                job_title = re.sub(r'\s+', ' ', text).strip()
                
                # Build Item on the spot (since no detail link is available, use career URL)
                item = JobItem()
                item["company_name"] = company_name
                item["job_title"] = job_title
                item["location"] = "United Kingdom"
                item["salary_range"] = "Not Specified"
                item["application_url"] = response.url
                
                # Quick scan of parent container text for exclusion warnings
                parent_text = "".join(elem.xpath("..//text()").getall()).lower()
                item["exclusion_keyword_found"] = any(w in parent_text for w in EXCLUSION_KEYWORDS)
                item["date_scraped"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                
                yield item

    def handle_error(self, failure):
        """Soft Error Handler. Logs error code, timeouts, DNS failures and continues."""
        req = failure.request
        company_name = req.meta.get("company_name", "Unknown Company")
        error_details = str(failure.value)
        
        # Assemble highly detailed entry for scraping_errors.log
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        log_entry = (
            f"[{timestamp}] Company: '{company_name}'\n"
            f"  Requested URL: {req.url}\n"
            f"  Error Type: {failure.type.__name__}\n"
            f"  Details: {error_details}\n"
            f"{'-'*60}\n"
        )
        
        try:
            with open(self.error_log_file, "a", encoding="utf-8") as f:
                f.write(log_entry)
            self.logger.warning(f"Soft error recorded for company '{company_name}' at {req.url}. Saved to {self.error_log_file}")
        except Exception as e:
            self.logger.critical(f"Failed to write to {self.error_log_file}: {e}")
