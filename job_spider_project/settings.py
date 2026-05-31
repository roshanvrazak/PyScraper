# -*- coding: utf-8 -*-
"""
Scrapy Settings
---------------
Production-grade configurations focusing on extreme resilience, politeness,
and anti-bot mitigation.
"""

BOT_NAME = 'job_spider_project'

SPIDER_MODULES = ['job_spider_project.spiders']
NEWSPIDER_MODULE = 'job_spider_project.spiders'

# --- Resilience & Anti-Bot Rules ---

# 1. Obey Robots.txt (Disabled for career portals which lock crawl policies aggressively)
ROBOTSTXT_OBEY = False

# 2. Politeness Delay to prevent hitting rate limits
DOWNLOAD_DELAY = 2.0

# 3. Dynamic AutoThrottle adjusts speed based on target server load
AUTOTHROTTLE_ENABLED = True
AUTOTHROTTLE_START_DELAY = 2.0
AUTOTHROTTLE_MAX_DELAY = 10.0
AUTOTHROTTLE_TARGET_CONCURRENCY = 1.0  # Max 1 concurrent request per target domain
AUTOTHROTTLE_DEBUG = False

# 4. Limit concurrency to prevent commercial firewalls triggering blocks
CONCURRENT_REQUESTS = 8
CONCURRENT_REQUESTS_PER_DOMAIN = 1
CONCURRENT_REQUESTS_PER_IP = 1

# 5. Disable cookies to avoid tracking and state footprinting
COOKIES_ENABLED = False

# 6. Default Headers
DEFAULT_REQUEST_HEADERS = {
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8',
    'Accept-Language': 'en-GB,en-US;q=0.9,en;q=0.8',
    'Connection': 'keep-alive',
    'Upgrade-Insecure-Requests': '1',
}

# 7. User-Agent rotation pool to simulate modern web browsers
USER_AGENT_POOL = [
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:121.0) Gecko/20100101 Firefox/121.0",
    "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.2 Safari/605.1.15",
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/121.0.0.0 Safari/537.36",
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
]

# Downloader Middlewares registration
DOWNLOADER_MIDDLEWARES = {
    'job_spider_project.middlewares.RotateUserAgentMiddleware': 400,
    'scrapy.downloadermiddlewares.useragent.UserAgentMiddleware': None,
}

# Pipelines configuration
ITEM_PIPELINES = {
    'job_spider_project.pipelines.ExcelExportPipeline': 300,
}

# Request Timeouts and Retry Settings
DOWNLOAD_TIMEOUT = 15.0
RETRY_ENABLED = True
RETRY_TIMES = 2
RETRY_HTTP_CODES = [500, 502, 503, 504, 408, 429]

# Logging configuration
LOG_LEVEL = 'INFO'
