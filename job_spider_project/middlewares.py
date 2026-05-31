# -*- coding: utf-8 -*-
"""
Scrapy Custom Downloader Middlewares
-----------------------------------
Implements user-agent rotation to emulate genuine browser sessions.
"""

import random
import logging
from scrapy import signals

logger = logging.getLogger(__name__)

class RotateUserAgentMiddleware:
    """Downloader Middleware that rotates User-Agents from a configured pool."""
    
    def __init__(self, user_agents):
        self.user_agents = user_agents

    @classmethod
    def from_crawler(cls, crawler):
        # Fetch the UA pool from settings
        user_agents = crawler.settings.get('USER_AGENT_POOL', [])
        middleware = cls(user_agents)
        crawler.signals.connect(middleware.spider_opened, signal=signals.spider_opened)
        return middleware

    def spider_opened(self, spider):
        spider.log(f"Registered RotateUserAgentMiddleware with {len(self.user_agents)} agents.")

    def process_request(self, request, spider):
        if self.user_agents:
            selected_ua = random.choice(self.user_agents)
            request.headers['User-Agent'] = selected_ua
            # Debug log to verify rotation is functional
            logger.debug(f"Assigned User-Agent: {selected_ua} to request: {request.url}")
