# -*- coding: utf-8 -*-
"""
Scrapy Item Definitions
-----------------------
Defines the JobItem model representing a single Tech Job posting.
"""

import scrapy

class JobItem(scrapy.Item):
    company_name = scrapy.Field()
    job_title = scrapy.Field()
    location = scrapy.Field()
    salary_range = scrapy.Field()
    application_url = scrapy.Field()
    sponsorship_keyword_found = scrapy.Field()
    date_scraped = scrapy.Field()
