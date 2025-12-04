"""
title: Domain Tracking Tool
author: 42frontiers
author_url: https://42frontiers.com
version: 1.7.0
license: MIT
description: Monitors newly registered domains to identify potential new PE funds, VC firms, or investment vehicles. Downloads NRD feeds and scans for keyword matches.
requirements: tldextract, textdistance, requests, pandas, aiohttp
"""

# Version identifier for debugging - if you don't see this in logs, old code is cached!
_TOOL_VERSION = "1.7.0-threadpool"

import asyncio
import base64
import concurrent.futures
import datetime
import csv
import json
import logging
import time
import zipfile
from collections import deque
from io import BytesIO
from pathlib import Path
from typing import Dict, List, Any, Optional, Literal

import aiohttp
import requests
import tldextract
import textdistance

from pydantic import BaseModel, Field
from pydantic.fields import FieldInfo

# Configure logging
log = logging.getLogger(__name__)

# Tool directory for file exports
TOOL_DIR = Path(__file__).parent


# =============================================================================
# PROGRESS TRACKER - Status indicator for OpenWebUI
# =============================================================================


class ProgressTracker:
    """
    Simple progress tracker for OpenWebUI tools.

    Emits status events with a single-line description showing current step.
    Status indicator is expandable in the OpenWebUI interface.

    Usage:
        progress = ProgressTracker(event_emitter=__event_emitter__)
        await progress.update("Downloading NRD feeds...")
        await progress.update("Scanning domains...")
        await progress.finish()
    """

    def __init__(self, event_emitter=None) -> None:
        self._event_emitter = event_emitter
        self._started = time.perf_counter()
        self._current_step = ""
        self._done: bool = False

    async def update(self, description: str) -> None:
        """Update the status with a new description."""
        if self._done:
            return
        self._current_step = description
        await self._emit_status()

    async def finish(self) -> None:
        """Mark as done with elapsed time."""
        if self._done:
            return
        elapsed = time.perf_counter() - self._started
        self._current_step = f"✓ Completed in {elapsed:.1f}s"
        self._done = True
        await self._emit_status(done=True)

    async def error(self, error_msg: str) -> None:
        """Mark as failed with error message."""
        if self._done:
            return
        elapsed = time.perf_counter() - self._started
        self._current_step = f"✗ Error after {elapsed:.1f}s: {error_msg}"
        self._done = True
        await self._emit_status(done=True)

    async def _emit_status(self, done: bool = False) -> None:
        """Emit a status event to OpenWebUI."""
        if not self._event_emitter:
            return

        await self._event_emitter({
            "type": "status",
            "data": {
                "description": self._current_step,
                "done": done,
            }
        })


# =============================================================================
# NRD FEED DOWNLOADER - Download newly registered domain feeds
# =============================================================================


class _NRDFeedDownloader:
    """
    Downloads Newly Registered Domain (NRD) feeds from multiple sources using async aiohttp.

    Sources:
    - WhoisDS: Daily NRD feed (~70k domains/day)
    - GitHub hagezi: Weekly NRD feeds that must be combined for longer periods
      - nrd7: days 1-7 (~2.6M domains)
      - nrd14-8: days 8-14 (~2.4M domains)
      - nrd21-15: days 15-21 (~2.6M domains)
      - nrd28-22: days 22-28 (~2.3M domains)

    To get ~28 days of domains, we combine all four feeds (~10M domains total).
    Downloads all feeds IN PARALLEL using asyncio for faster performance.
    """

    WHOISDS_URL_TEMPLATE = "https://whoisds.com/whois-database/newly-registered-domains/{}/nrd"

    # Weekly segment URLs from hagezi - try jsdelivr CDN first (faster), fallback to raw GitHub
    # File naming: nrd7 (days 1-7), nrd14-8 (days 8-14), nrd21-15 (days 15-21), etc.
    GITHUB_NRD_URLS = {
        "nrd7": [
            "https://cdn.jsdelivr.net/gh/hagezi/dns-blocklists@latest/domains/nrd7.txt",
            "https://raw.githubusercontent.com/hagezi/dns-blocklists/main/domains/nrd7.txt",
        ],
        "nrd14-8": [
            "https://cdn.jsdelivr.net/gh/hagezi/dns-blocklists@latest/domains/nrd14-8.txt",
            "https://raw.githubusercontent.com/hagezi/dns-blocklists/main/domains/nrd14-8.txt",
        ],
        "nrd21-15": [
            "https://cdn.jsdelivr.net/gh/hagezi/dns-blocklists@latest/domains/nrd21-15.txt",
            "https://raw.githubusercontent.com/hagezi/dns-blocklists/main/domains/nrd21-15.txt",
        ],
        "nrd28-22": [
            "https://cdn.jsdelivr.net/gh/hagezi/dns-blocklists@latest/domains/nrd28-22.txt",
            "https://raw.githubusercontent.com/hagezi/dns-blocklists/main/domains/nrd28-22.txt",
        ],
    }

    def __init__(self):
        self._headers = {
            "Accept-Encoding": "gzip, deflate",
            "User-Agent": "Mozilla/5.0 (compatible; DomainTracker/1.0)"
        }

    async def download_whoisds_async(self) -> List[str]:
        """Download domains from WhoisDS NRD feed (async)."""
        try:
            # WhoisDS uses base64 encoded date for the URL
            yesterday = (datetime.datetime.today() - datetime.timedelta(days=1)).strftime('%Y-%m-%d')
            encoded_date = base64.b64encode(f"{yesterday}.zip".encode('ascii')).decode('ascii')
            url = self.WHOISDS_URL_TEMPLATE.format(encoded_date)

            async with aiohttp.ClientSession() as session:
                async with session.get(url, timeout=aiohttp.ClientTimeout(total=60)) as response:
                    response.raise_for_status()
                    content = await response.read()

            # Extract domains from zip file
            domains = []
            with zipfile.ZipFile(BytesIO(content)) as zf:
                for filename in zf.namelist():
                    if filename.endswith('.txt'):
                        file_content = zf.read(filename).decode('utf-8-sig')
                        for line in file_content.splitlines():
                            domain = line.strip().lower()
                            if domain:
                                domains.append(domain)

            log.info(f"Downloaded {len(domains)} domains from WhoisDS")
            return domains

        except Exception as e:
            log.warning(f"Failed to download WhoisDS feed: {e}")
            return []

    async def _download_single_nrd_feed_async(
        self,
        feed_name: str,
        urls: List[str],
        session: aiohttp.ClientSession
    ) -> set:
        """
        Download a single NRD feed and return domains as a set (async).

        Tries multiple URLs in order (jsdelivr CDN first, then raw GitHub as fallback).
        Uses gzip compression via Accept-Encoding header to reduce transfer size.
        """
        last_error = None
        log.info(f"[DEBUG] _download_single_nrd_feed_async called for {feed_name} with {len(urls)} URLs")

        for idx, url in enumerate(urls):
            try:
                source = "jsdelivr" if "jsdelivr" in url else "raw.github"
                log.info(f"[DEBUG] Attempting {feed_name} URL {idx+1}/{len(urls)}: {source}")
                log.info(f"[DEBUG] Full URL: {url}")

                async with session.get(
                    url,
                    timeout=aiohttp.ClientTimeout(total=180),
                    headers=self._headers
                ) as response:
                    log.info(f"[DEBUG] {feed_name} response status: {response.status}")
                    log.info(f"[DEBUG] {feed_name} response headers: {dict(response.headers)}")

                    if response.status >= 400:
                        log.warning(f"[DEBUG] {feed_name} got HTTP {response.status}, trying next URL")
                        continue

                    # Read and parse domains
                    content_encoding = response.headers.get("Content-Encoding", "none")
                    content_length = response.headers.get("Content-Length", "unknown")
                    log.info(f"[DEBUG] {feed_name}: encoding={content_encoding}, content-length={content_length}")

                    text = await response.text()
                    log.info(f"[DEBUG] {feed_name}: received {len(text):,} chars of text")

                    # Show first 200 chars for debugging
                    preview = text[:200].replace('\n', '\\n')
                    log.info(f"[DEBUG] {feed_name} preview: {preview}")

                    domains = set()
                    line_count = 0
                    for line in text.splitlines():
                        line_count += 1
                        line = line.strip().lower()
                        if line and not line.startswith('#'):
                            domains.add(line)

                    log.info(f"[DEBUG] {feed_name}: parsed {line_count:,} lines -> {len(domains):,} domains")
                    return domains

            except asyncio.TimeoutError:
                last_error = f"Timeout downloading {feed_name} from {url}"
                log.warning(f"[DEBUG] {last_error}")
                continue
            except aiohttp.ClientError as e:
                last_error = e
                log.warning(f"[DEBUG] aiohttp.ClientError for {feed_name} from {url}: {type(e).__name__}: {e}")
                continue
            except Exception as e:
                last_error = e
                log.warning(f"[DEBUG] Exception for {feed_name} from {url}: {type(e).__name__}: {e}")
                import traceback
                log.warning(f"[DEBUG] Traceback: {traceback.format_exc()}")
                continue

        log.error(f"[DEBUG] All URLs failed for {feed_name}. Last error: {last_error}")
        return set()

    async def download_github_nrd_async(self, days: int = 7) -> List[str]:
        """
        Download domains from GitHub hagezi NRD feeds IN PARALLEL.

        The hagezi feeds are split into weekly segments:
        - nrd7: days 1-7 (yesterday to 7 days ago)
        - nrd14-8: days 8-14
        - nrd21-15: days 15-21
        - nrd28-22: days 22-28

        Args:
            days: Number of days to fetch (7, 14, 21, or 28/30)
                  7 = nrd7 only (~2.6M domains)
                  14 = nrd7 + nrd14-8 (~5M domains)
                  21 = nrd7 + nrd14-8 + nrd21-15 (~7.6M domains)
                  28/30 = all four feeds (~10M domains)

        Returns:
            List of unique domain names
        """
        # Determine which feeds to download based on requested days
        feeds_to_download = []
        if days >= 7:
            feeds_to_download.append("nrd7")
        if days >= 14:
            feeds_to_download.append("nrd14-8")
        if days >= 21:
            feeds_to_download.append("nrd21-15")
        if days >= 28:
            feeds_to_download.append("nrd28-22")

        if not feeds_to_download:
            feeds_to_download = ["nrd7"]  # Default to 7 days

        log.info(f"[DEBUG] download_github_nrd_async called with days={days}")
        log.info(f"[DEBUG] Downloading {len(feeds_to_download)} NRD feeds in parallel: {feeds_to_download}")

        all_domains = set()

        # Use a single session for all parallel downloads
        connector = aiohttp.TCPConnector(limit=4)
        async with aiohttp.ClientSession(connector=connector) as session:
            # Create tasks for all feeds to download in parallel
            tasks = []
            for feed_name in feeds_to_download:
                urls = self.GITHUB_NRD_URLS.get(feed_name, [])
                if urls:
                    task = self._download_single_nrd_feed_async(feed_name, urls, session)
                    tasks.append(task)

            # Download all feeds in parallel
            results = await asyncio.gather(*tasks, return_exceptions=True)

            # Combine results
            for i, result in enumerate(results):
                if isinstance(result, Exception):
                    log.error(f"Feed {feeds_to_download[i]} failed: {result}")
                elif isinstance(result, set):
                    all_domains.update(result)

        log.info(f"Total unique domains from {len(feeds_to_download)} feeds: {len(all_domains):,}")
        return list(all_domains)

    # Sync wrappers for backward compatibility
    def download_whoisds(self) -> List[str]:
        """Sync wrapper for WhoisDS download."""
        try:
            loop = asyncio.get_event_loop()
            if loop.is_running():
                # We're already in an async context, can't use run_until_complete
                # Fall back to sync requests
                return self._download_whoisds_sync()
            return loop.run_until_complete(self.download_whoisds_async())
        except RuntimeError:
            return self._download_whoisds_sync()

    def _download_whoisds_sync(self) -> List[str]:
        """Sync fallback for WhoisDS download."""
        try:
            yesterday = (datetime.datetime.today() - datetime.timedelta(days=1)).strftime('%Y-%m-%d')
            encoded_date = base64.b64encode(f"{yesterday}.zip".encode('ascii')).decode('ascii')
            url = self.WHOISDS_URL_TEMPLATE.format(encoded_date)

            response = requests.get(url, timeout=60)
            response.raise_for_status()

            domains = []
            with zipfile.ZipFile(BytesIO(response.content)) as zf:
                for filename in zf.namelist():
                    if filename.endswith('.txt'):
                        content = zf.read(filename).decode('utf-8-sig')
                        for line in content.splitlines():
                            domain = line.strip().lower()
                            if domain:
                                domains.append(domain)

            log.info(f"Downloaded {len(domains)} domains from WhoisDS")
            return domains
        except Exception as e:
            log.warning(f"Failed to download WhoisDS feed: {e}")
            return []

    def download_github_nrd(self, days: int = 7) -> List[str]:
        """Sync wrapper for GitHub NRD download."""
        try:
            loop = asyncio.get_event_loop()
            if loop.is_running():
                # We're already in an async context, can't use run_until_complete
                # Fall back to sync version
                return self._download_github_nrd_sync(days)
            return loop.run_until_complete(self.download_github_nrd_async(days))
        except RuntimeError:
            return self._download_github_nrd_sync(days)

    def _download_github_nrd_sync(self, days: int = 7) -> List[str]:
        """Sync fallback for GitHub NRD download."""
        feeds_to_download = []
        if days >= 7:
            feeds_to_download.append("nrd7")
        if days >= 14:
            feeds_to_download.append("nrd14-8")
        if days >= 21:
            feeds_to_download.append("nrd21-15")
        if days >= 28:
            feeds_to_download.append("nrd28-22")

        if not feeds_to_download:
            feeds_to_download = ["nrd7"]

        all_domains = set()

        for feed_name in feeds_to_download:
            urls = self.GITHUB_NRD_URLS.get(feed_name, [])
            for url in urls:
                try:
                    source = "jsdelivr" if "jsdelivr" in url else "raw.github"
                    log.info(f"Downloading {feed_name} from {source}...")

                    response = requests.get(url, timeout=180, headers=self._headers)
                    response.raise_for_status()

                    domains = set()
                    for line in response.text.splitlines():
                        line = line.strip().lower()
                        if line and not line.startswith('#'):
                            domains.add(line)

                    log.info(f"Downloaded {len(domains):,} domains from {feed_name}")
                    all_domains.update(domains)
                    break  # Success, don't try fallback URLs
                except Exception as e:
                    log.warning(f"Failed to download {feed_name} from {url}: {e}")
                    continue

        log.info(f"Total unique domains from {len(feeds_to_download)} feeds: {len(all_domains):,}")
        return list(all_domains)


# =============================================================================
# WEBSITE SCANNER - Check if domains have active websites with keywords
# =============================================================================


class _WebsiteScanner:
    """
    Async website scanner using aiohttp for concurrent requests.

    Scans websites behind domains to check for active content and keyword presence.
    Uses asyncio to scan multiple websites concurrently (non-blocking).

    Returns status:
    - "keyword_found": Active website with PE topic keywords found
    - "excluded": Active website but contains VC/M&A/excluded keywords (NOT PE)
    - "no_keywords": Active website but no topic keywords found
    - "parked": Domain is parked or shows placeholder content
    - "no_website": No website accessible (timeout, DNS error, etc.)
    - "error": Scan failed
    """

    # Common parked domain indicators
    PARKED_INDICATORS = [
        "domain is for sale",
        "buy this domain",
        "this domain may be for sale",
        "domain parking",
        "parked free",
        "godaddy",
        "hugedomains",
        "sedo.com",
        "dan.com",
        "afternic",
        "undeveloped.com",
        "this page is parked",
        "coming soon",
        "under construction",
        "website coming soon",
        "page not found",
        "default web page",
        "apache2 default page",
        "nginx welcome",
        "plesk default page",
        "cpanel",
        "website is under maintenance",
    ]

    def __init__(
        self,
        topic_keywords: List[str],
        blacklist_keywords: Optional[List[str]] = None,
        timeout: int = 10,
        max_concurrent: int = 10
    ):
        self.topic_keywords = [k.lower() for k in topic_keywords]
        self.blacklist_keywords = [k.lower() for k in (blacklist_keywords or [])]
        self.timeout = timeout
        self.max_concurrent = max_concurrent
        self._headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
            "Accept-Language": "en-US,en;q=0.5",
        }

    def _analyze_content(self, content: str) -> Dict[str, Any]:
        """Analyze page content for keywords (sync helper)."""
        content = content.lower()
        content_length = len(content)

        # Very short content is likely parked/empty
        if content_length < 500:
            return {
                "status": "parked",
                "keywords_found": [],
                "blacklist_found": [],
                "explanation": "Minimal content (likely placeholder)"
            }

        # Check for parked domain indicators
        for indicator in self.PARKED_INDICATORS:
            if indicator in content:
                return {
                    "status": "parked",
                    "keywords_found": [],
                    "blacklist_found": [],
                    "explanation": f"Parked domain ({indicator[:20]}...)"
                }

        # Check for blacklist keywords FIRST (VC, M&A, etc.)
        found_blacklist = [kw for kw in self.blacklist_keywords if kw in content]

        # Check for topic keywords (PE indicators)
        found_keywords = [kw for kw in self.topic_keywords if kw in content]

        # If blacklist keywords found, mark as excluded
        if found_blacklist:
            return {
                "status": "excluded",
                "keywords_found": found_keywords[:5],
                "blacklist_found": found_blacklist[:5],
                "explanation": f"Excluded: {', '.join(found_blacklist[:3])}"
            }

        if found_keywords:
            return {
                "status": "keyword_found",
                "keywords_found": found_keywords[:5],
                "blacklist_found": [],
                "explanation": f"Keywords: {', '.join(found_keywords[:3])}"
            }
        else:
            return {
                "status": "no_keywords",
                "keywords_found": [],
                "blacklist_found": [],
                "explanation": "Active site, no fund keywords"
            }

    async def scan_website_async(self, domain: str, session: aiohttp.ClientSession) -> Dict[str, Any]:
        """
        Async scan a single domain's website.

        Returns:
            Dict with keys: status, keywords_found, blacklist_found, explanation
        """
        urls_to_try = [f"https://{domain}", f"http://{domain}"]

        for url in urls_to_try:
            try:
                async with session.get(
                    url,
                    timeout=aiohttp.ClientTimeout(total=self.timeout),
                    allow_redirects=True,
                    ssl=False  # Some new domains have invalid certs
                ) as response:
                    if response.status >= 400:
                        continue

                    # Read response text
                    content = await response.text()
                    return self._analyze_content(content)

            except asyncio.TimeoutError:
                continue
            except aiohttp.ClientSSLError:
                continue
            except aiohttp.ClientConnectorError:
                continue
            except aiohttp.ClientError:
                continue
            except Exception as e:
                log.debug(f"Error scanning {domain}: {e}")
                continue

        return {
            "status": "no_website",
            "keywords_found": [],
            "blacklist_found": [],
            "explanation": "No website accessible"
        }

    async def scan_multiple_async(
        self,
        domains: List[str],
        progress_callback: Optional[callable] = None
    ) -> Dict[str, Dict]:
        """
        Scan multiple domains concurrently using asyncio.

        Args:
            domains: List of domains to scan
            progress_callback: Optional async callback(completed, total, current_domain)

        Returns:
            Dict mapping domain -> scan result
        """
        results = {}
        semaphore = asyncio.Semaphore(self.max_concurrent)

        async def scan_with_semaphore(domain: str, session: aiohttp.ClientSession) -> tuple:
            async with semaphore:
                result = await self.scan_website_async(domain, session)
                return domain, result

        connector = aiohttp.TCPConnector(limit=self.max_concurrent, ssl=False)
        async with aiohttp.ClientSession(headers=self._headers, connector=connector) as session:
            tasks = [scan_with_semaphore(domain, session) for domain in domains]

            # Process results as they complete
            completed = 0
            for coro in asyncio.as_completed(tasks):
                domain, result = await coro
                results[domain] = result
                completed += 1

                if progress_callback and completed % 5 == 0:
                    await progress_callback(completed, len(domains), domain)

        return results

    # Keep sync method for backward compatibility
    def scan_website(self, domain: str) -> Dict[str, Any]:
        """Sync wrapper for single domain scan (backward compatibility)."""
        import requests as sync_requests

        urls_to_try = [f"https://{domain}", f"http://{domain}"]

        for url in urls_to_try:
            try:
                response = sync_requests.get(
                    url,
                    timeout=self.timeout,
                    allow_redirects=True,
                    verify=False,
                    headers=self._headers
                )

                if response.status_code >= 400:
                    continue

                return self._analyze_content(response.text)

            except Exception:
                continue

        return {
            "status": "no_website",
            "keywords_found": [],
            "blacklist_found": [],
            "explanation": "No website accessible"
        }


# =============================================================================
# DOMAIN SCANNER - Keyword matching and similarity detection
# =============================================================================


class _DomainScanner:
    """
    Scans domains for keyword matches using a two-tier matching system.

    Matching Strategy:
    1. COMPOUND KEYWORDS (always matched): PE-specific compound terms like
       "capitalpartners", "privateequity", "buyoutfund" - these rarely appear
       in unrelated domains.

    2. SIMPLE KEYWORDS (TLD-restricted): Common words like "capital", "fund",
       "partners" are ONLY matched on PE-specific TLDs (.fund, .capital, etc.)
       to avoid millions of false positives.

    This dramatically reduces false positives while still catching legitimate
    new PE fund domains.
    """

    def __init__(
        self,
        compound_keywords: List[str],
        simple_keywords: List[str],
        blacklist: List[str],
        domain_blacklist: List[str],
        high_priority_tlds: List[str],
        similarity_mode: str = "close"
    ):
        self.compound_keywords = [k.lower() for k in compound_keywords]
        self.simple_keywords = [k.lower() for k in simple_keywords]
        self.blacklist = [b.lower() for b in blacklist]
        self.domain_blacklist = [d.lower() for d in domain_blacklist]
        self.high_priority_tlds = [t.lower().lstrip('.') for t in high_priority_tlds]
        self.similarity_mode = similarity_mode

        # Results storage
        self.results: List[Dict] = []
        self.scan_stats: Dict = {}

        # TLD extractor (lazy init)
        self._tld_extract = None

    def _get_tld_extract(self):
        """Get or create TLD extractor."""
        if self._tld_extract is None:
            self._tld_extract = tldextract.TLDExtract(include_psl_private_domains=True)
            self._tld_extract("google.com")  # Warm up cache
        return self._tld_extract

    def _get_thresholds(self) -> Dict:
        """Get similarity thresholds based on mode."""
        if self.similarity_mode == "medium":
            return {"jaccard": 0.50, "jaro_winkler": 0.85, "damerau_max": 2}
        elif self.similarity_mode == "wide":
            return {"jaccard": 0.45, "jaro_winkler": 0.80, "damerau_max": 3}
        else:  # close (default)
            return {"jaccard": 0.60, "jaro_winkler": 0.90, "damerau_max": 1}

    def _is_blacklisted(self, domain: str) -> bool:
        """Check if domain contains blacklisted keywords."""
        domain_lower = domain.lower()
        return any(bl in domain_lower for bl in self.blacklist)

    def _is_domain_blacklisted(self, domain_name: str) -> bool:
        """
        Check if domain name contains words from the domain blacklist.
        This filters out false positives like 'adventure' matching 'venture'.
        """
        return any(bl in domain_name for bl in self.domain_blacklist)

    def _extract_domain_name(self, domain: str) -> str:
        """Extract the domain name without TLD."""
        ext = self._get_tld_extract()
        result = ext(domain)
        return result.domain.lower()

    def _check_jaccard(self, keyword: str, domain_name: str, threshold: float) -> bool:
        """Check Jaccard similarity using bigrams."""
        if len(keyword) < 2 or len(domain_name) < 2:
            return False

        # Create bigrams
        kw_bigrams = set(keyword[i:i+2] for i in range(len(keyword)-1))
        dom_bigrams = set(domain_name[i:i+2] for i in range(len(domain_name)-1))

        if not kw_bigrams or not dom_bigrams:
            return False

        intersection = len(kw_bigrams & dom_bigrams)
        union = len(kw_bigrams | dom_bigrams)

        similarity = intersection / union if union > 0 else 0
        return similarity >= threshold

    def _check_damerau(self, keyword: str, domain_name: str, max_distance: int) -> bool:
        """Check Damerau-Levenshtein distance."""
        # Only check if lengths are reasonably close
        len_diff = abs(len(keyword) - len(domain_name))
        if len_diff > max_distance + 2:
            return False

        distance = textdistance.damerau_levenshtein(keyword, domain_name)

        # Scale max distance by keyword length
        if len(keyword) <= 5:
            allowed = min(max_distance, 1)
        elif len(keyword) <= 8:
            allowed = max_distance
        else:
            allowed = max_distance + 1

        return distance <= allowed

    def _check_jaro_winkler(self, keyword: str, domain_name: str, threshold: float) -> bool:
        """Check Jaro-Winkler similarity."""
        similarity = textdistance.jaro_winkler(keyword, domain_name)
        return similarity >= threshold

    def _scan_single_domain(
        self,
        domain: str,
        thresholds: Dict,
        today: str,
        priority_tlds: List[str]
    ) -> Optional[Dict]:
        """
        Scan a single domain for keyword matches.
        Returns match result dict or None if no match.
        """
        # Skip blacklisted (content blacklist)
        if self._is_blacklisted(domain):
            return None

        domain_name = self._extract_domain_name(domain)

        # Skip domains containing false positive words
        if self._is_domain_blacklisted(domain_name):
            return None

        ext = self._get_tld_extract()(domain)
        tld = ext.suffix.lower()
        is_high_priority_tld = tld in self.high_priority_tlds

        match_type = None
        matched_keyword = None

        # TIER 1: Check compound keywords (always match)
        for keyword in self.compound_keywords:
            if len(keyword) < 5:
                continue

            if keyword in domain_name:
                match_type = "Compound Match"
                matched_keyword = keyword
                break

            # Similarity check for compound keywords (typosquatting detection)
            if len(keyword) >= 10:
                if self._check_jaccard(keyword, domain_name, thresholds["jaccard"]):
                    match_type = "Similarity (Compound)"
                    matched_keyword = keyword
                    break

        # TIER 2: Check simple keywords ONLY on high-priority TLDs
        if not match_type and is_high_priority_tld:
            for keyword in self.simple_keywords:
                if len(keyword) < 4:
                    continue

                if keyword in domain_name:
                    match_type = "Simple Match (Priority TLD)"
                    matched_keyword = keyword
                    break

        if match_type and matched_keyword:
            return {
                "domain": domain,
                "keyword": matched_keyword,
                "detection": match_type,
                "date": today,
                "priority_tld": tld in priority_tlds,
                "tld": tld
            }

        return None

    def _scan_batch_sync(
        self,
        domains: List[str],
        thresholds: Dict,
        today: str,
        priority_tlds_list: List[str]
    ) -> List[Dict]:
        """
        Scan a batch of domains synchronously (runs in thread pool).
        Returns list of matches.
        """
        results = []
        for domain in domains:
            result = self._scan_single_domain(domain, thresholds, today, priority_tlds_list)
            if result:
                results.append(result)
        return results

    async def scan(
        self,
        domains: List[str],
        priority_tlds: Optional[List[str]] = None,
        progress_callback: Optional[callable] = None
    ) -> List[Dict]:
        """
        Scan domains for keyword matches using two-tier matching.

        Runs CPU-intensive scanning in a thread pool to avoid blocking the event loop.
        Processes domains in batches and reports progress between batches.

        Tier 1 (COMPOUND KEYWORDS): Always matched - these are PE-specific
        compound terms like "capitalpartners", "privateequity" that rarely
        appear in unrelated domains.

        Tier 2 (SIMPLE KEYWORDS): Only matched on high-priority TLDs like
        .fund, .capital, .investments to avoid false positives from common
        words like "capital" or "fund".

        Args:
            domains: List of domain names to scan
            priority_tlds: TLDs to flag as high priority in results
            progress_callback: Optional async callback(scanned, total, matches) for progress

        Returns:
            List of match results
        """
        thresholds = self._get_thresholds()
        today = str(datetime.date.today())
        priority_tlds_list = [t.lower().lstrip('.') for t in (priority_tlds or [])]

        results = []
        total = len(domains)

        # Process in batches using thread pool
        # Batch size of 100k balances progress updates with thread overhead
        batch_size = 100_000
        loop = asyncio.get_event_loop()

        # Use thread pool executor for CPU-bound work
        with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:
            for batch_start in range(0, total, batch_size):
                batch_end = min(batch_start + batch_size, total)
                batch = domains[batch_start:batch_end]

                # Run batch scan in thread pool (non-blocking)
                batch_results = await loop.run_in_executor(
                    executor,
                    self._scan_batch_sync,
                    batch,
                    thresholds,
                    today,
                    priority_tlds_list
                )

                results.extend(batch_results)

                # Report progress after each batch
                if progress_callback:
                    pct = int(100 * batch_end / total)
                    await progress_callback(batch_end, total, len(results))

        # Dedupe by domain
        seen = set()
        unique_results = []
        for r in results:
            if r["domain"] not in seen:
                seen.add(r["domain"])
                unique_results.append(r)

        # Sort: Compound matches first, then priority TLDs, then by domain
        unique_results.sort(key=lambda x: (
            x["detection"] != "Compound Match",
            not x["priority_tld"],
            x["domain"]
        ))

        self.results = unique_results
        self.scan_stats = {
            "scan_date": today,
            "total_domains_scanned": len(domains),
            "compound_keywords": self.compound_keywords,
            "simple_keywords": self.simple_keywords,
            "high_priority_tlds": self.high_priority_tlds,
            "matches_found": len(unique_results),
            "similarity_mode": self.similarity_mode
        }

        return unique_results


# =============================================================================
# TOOLS CLASS - Public interface for the LLM
# =============================================================================


class Tools:
    """
    Domain Tracking Tool for Open WebUI.

    Monitors newly registered domains to identify potential new PE funds,
    VC firms, or investment vehicles being formed.

    Workflow:
        1. scan_new_domains() - Download feeds and scan for matches
        2. get_results_summary() - Get formatted markdown summary
        3. export_results_csv() - Export results to CSV file
    """

    class Valves(BaseModel):
        """Admin-configurable settings."""
        # ----- Compound Keywords (REQUIRED - high specificity) -----
        # These are PE-specific compound terms that rarely appear in unrelated domains
        # The scanner will ONLY match these exact compound patterns
        COMPOUND_KEYWORDS: str = Field(
            default="capitalpartners,equitypartners,privateequity,buyoutfund,buyoutpartners,venturecapital,investmentpartners,equityfund,capitalfund,holdingpartners,beteiligungskapital,kapitalbeteiligung,equityholding,privatedebt,growthequity,growthcapital,midmarketfund,middlemarketfund",
            description="Compound PE-specific terms (high confidence, always matched)"
        )
        # ----- Simple Keywords (only matched with priority TLDs) -----
        # These common words only count as matches on fund-specific TLDs (.fund, .capital, etc.)
        SIMPLE_KEYWORDS: str = Field(
            default="buyout,equity,capital,partners,fund,invest,holdings,beteiligung,kapital,fonds",
            description="Simple keywords - ONLY matched on priority TLDs (.fund, .capital, etc.)"
        )
        # ----- Legacy KEYWORDS field (backward compatibility) -----
        KEYWORDS: str = Field(
            default="",
            description="DEPRECATED: Use COMPOUND_KEYWORDS and SIMPLE_KEYWORDS instead"
        )
        # ----- Topic Keywords (for website content scanning) -----
        TOPIC_KEYWORDS: str = Field(
            default="investments,investment,partners,equity,majority stake,minority stake,EBITDA,portfolio company,acquisition,acquisitions,midmarket,mittelstand,beteiligung,kapital,fonds,vermögen,holdings,management,advisors,fund,funds,financial,strategic,principals,associates,assets,buyout,leveraged,platform investment",
            description="Keywords to search for in website content (indicates active PE fund)"
        )
        # ----- Website Blacklist (VC/M&A indicators on website content) -----
        WEBSITE_BLACKLIST: str = Field(
            default="venture capital,venture fund,seed funding,seed stage,series a,series b,early stage,startup,startups,start-up,angel investor,angel investment,accelerator,incubator,pre-seed,m&a advisory,m&a boutique,investment bank,corporate finance advisory,sell-side,buy-side mandate,deal origination,transaction advisory,fairness opinion,valuation services,due diligence services,restructuring advisory,real estate,property investment,proptech,infrastructure fund,credit fund,hedge fund,quantitative,algorithmic trading,crypto,blockchain,defi,web3,nft",
            description="Keywords that indicate VC/M&A/non-PE when found on website (excludes domain from results)"
        )
        # ----- Blacklist (Exclusions per client requirements - for DOMAIN names) -----
        BLACKLIST: str = Field(
            default="venture capital,real estate,infrastructure,growth capital,blockchain,hedge fund,private credit,private debt,search fund,mezzanine,fixed income,commodities,natural resources,forestry,music rights,consulting,M&A advisor,investment banking,VC,crypto,seed,series a,proptech,fintech,cleantech,accelerator,incubator,bitcoin,nft,casino,loan,forex,trading,bet,gambling,poker,slots,lottery,binary,mlm,pyramid,liquid assets",
            description="Keywords to exclude from domain name matching (VC, real estate, infrastructure, etc.)"
        )
        # ----- Domain Blacklist (filter out common false positive domains) -----
        DOMAIN_BLACKLIST: str = Field(
            default="adventure,adventures,adventurous,adventurer,accounting,accountant,accountants,tourism,tourist,travel,traveling,plumbing,plumber,photography,photographer,portfolio,portfolios,painting,painter,cleaning,cleaner,catering,caterer,marketing,marketer,consulting,consultant,coaching,fishing,hunting,camping,gaming,streaming,blogging,vlogging,crafting,staffing,hosting,renting,booking,shipping,trading,farming,gardening,landscaping,roofing,flooring,fencing,welding,moving,storage,laundry,bakery,brewery,grocery,jewelry,pottery,dentistry,pharmacy,surgery,therapy,recovery,delivery,discovery,advisory,inventory,mandatory,documentary,commentary,elementary,supplementary,parliamentary,revolutionary,evolutionary,extraordinary,contemporary,anniversary,customary,visionary,missionary,imaginary,ordinary,secondary,primary,literary,military,culinary,veterinary,preliminary,disciplinary,dictionary,stationary,sanctuary,monastery,infirmary,obituary,ventures,advisors,boutique,banking,realty,properties,mortgage,crypto,defi,token,nft,coin,blockchain",
            description="Domain words that trigger false positives (e.g., 'adventure' contains 'venture', 'ventures' is usually VC)"
        )
        SIMILARITY_MODE: Literal["close", "medium", "wide"] = Field(
            default="close",
            description="Detection sensitivity: 'close' (precise), 'medium', or 'wide' (more matches)"
        )
        # ----- TLD Configuration -----
        # HIGH_PRIORITY_TLDS: Simple keywords ONLY match on these TLDs
        HIGH_PRIORITY_TLDS: str = Field(
            default=".fund,.capital,.investments,.partners,.ventures",
            description="TLDs where simple keywords are matched (PE-specific TLDs)"
        )
        # PRIORITY_TLDS: All TLDs to flag in results (for sorting)
        PRIORITY_TLDS: str = Field(
            default=".fund,.capital,.investments,.partners,.ventures,.com,.de,.eu,.co,.io,.at,.ch,.nl,.be,.uk",
            description="All priority TLDs for sorting results (European focus)"
        )
        # ----- Data Source Configuration -----
        NRD_SOURCE: Literal["whoisds_only", "github_nrd7", "github_nrd30", "both"] = Field(
            default="github_nrd30",
            description="NRD source: 'whoisds_only' (~70k/day), 'github_nrd7' (~500k/week), 'github_nrd30' (~2M/30days), 'both'"
        )
        INCLUDE_TOPIC_KEYWORDS_IN_SCAN: bool = Field(
            default=False,
            description="Include topic_keywords in domain scanning"
        )
        LOG_LEVEL: Literal["DEBUG", "INFO", "NONE"] = Field(
            default="INFO",
            description="Logging level"
        )

    def __init__(self):
        """Initialize the Domain Tracking Tool."""
        self._scanner: Optional[_DomainScanner] = None
        self._downloader: Optional[_NRDFeedDownloader] = None
        self._session_logs: deque = deque(maxlen=500)
        self.valves = self.Valves()
        self.file_handler = False
        self.citation = False

    def _parse_list(self, value: str) -> List[str]:
        """Parse comma-separated string into list."""
        if isinstance(value, FieldInfo):
            value = value.default or ""
        return [x.strip().lower() for x in value.split(",") if x.strip()]

    def _log(self, message: str) -> None:
        """Add message to session log."""
        timestamp = datetime.datetime.now().strftime("%H:%M:%S")
        self._session_logs.append(f"[{timestamp}] {message}")
        log.info(message)

    def _get_scanner(self) -> _DomainScanner:
        """Create a domain scanner with current settings."""
        # Parse compound keywords (high specificity, always matched)
        compound_keywords = self._parse_list(self.valves.COMPOUND_KEYWORDS)

        # Parse simple keywords (only matched on priority TLDs)
        simple_keywords = self._parse_list(self.valves.SIMPLE_KEYWORDS)

        # Legacy support: if old KEYWORDS field has values, add to simple keywords
        legacy_keywords = self._parse_list(self.valves.KEYWORDS)
        if legacy_keywords:
            simple_keywords.extend(legacy_keywords)

        blacklist = self._parse_list(self.valves.BLACKLIST)
        domain_blacklist = self._parse_list(self.valves.DOMAIN_BLACKLIST)
        high_priority_tlds = self._parse_list(self.valves.HIGH_PRIORITY_TLDS)

        similarity_mode = self.valves.SIMILARITY_MODE
        if isinstance(similarity_mode, FieldInfo):
            similarity_mode = similarity_mode.default or "close"

        self._scanner = _DomainScanner(
            compound_keywords=list(set(compound_keywords)),
            simple_keywords=list(set(simple_keywords)),
            blacklist=blacklist,
            domain_blacklist=domain_blacklist,
            high_priority_tlds=high_priority_tlds,
            similarity_mode=similarity_mode
        )
        return self._scanner

    def _get_downloader(self) -> _NRDFeedDownloader:
        """Get or create NRD feed downloader."""
        if self._downloader is None:
            self._downloader = _NRDFeedDownloader()
        return self._downloader

    async def scan_new_domains(
        self,
        additional_compound_keywords: Optional[str] = None,
        __event_emitter__=None
    ) -> str:
        """
        Scan newly registered domains for fund-related keywords.

        Uses two-tier matching:
        - COMPOUND KEYWORDS: Always matched (e.g., "capitalpartners", "privateequity")
        - SIMPLE KEYWORDS: Only matched on priority TLDs (.fund, .capital, etc.)

        Args:
            additional_compound_keywords: Optional comma-separated compound keywords to add

        Returns:
            JSON with scan results
        """
        progress = ProgressTracker(event_emitter=__event_emitter__)

        try:
            # Log version to confirm new code is running
            log.info(f"[VERSION] Domain Tracking Tool {_TOOL_VERSION} starting scan")
            await progress.update(f"Initializing domain scanner (v{_TOOL_VERSION})...")

            scanner = self._get_scanner()
            downloader = self._get_downloader()

            # Add extra compound keywords if provided
            if additional_compound_keywords:
                extra = self._parse_list(additional_compound_keywords)
                scanner.compound_keywords.extend(extra)
                scanner.compound_keywords = list(set(scanner.compound_keywords))
                self._log(f"Added {len(extra)} additional compound keywords")

            # Get NRD source setting
            nrd_source = self.valves.NRD_SOURCE
            if isinstance(nrd_source, FieldInfo):
                nrd_source = nrd_source.default or "github_nrd30"

            all_domains = []

            if nrd_source in ("whoisds_only", "both"):
                await progress.update("Downloading NRD feeds from WhoisDS...")
                whoisds_domains = await downloader.download_whoisds_async()
                all_domains.extend(whoisds_domains)
                await progress.update(f"WhoisDS: {len(whoisds_domains):,} domains")

            if nrd_source == "github_nrd7":
                await progress.update("Downloading NRD feed from GitHub (nrd-7: last 7 days)...")
                github_domains = await downloader.download_github_nrd_async(days=7)
                all_domains.extend(github_domains)
                await progress.update(f"GitHub nrd-7: {len(github_domains):,} domains")

            if nrd_source in ("github_nrd30", "both"):
                await progress.update("Downloading NRD feeds from GitHub (4 weekly feeds in parallel = ~28 days)...")
                github_domains = await downloader.download_github_nrd_async(days=30)
                all_domains.extend(github_domains)
                await progress.update(f"GitHub NRD (28 days): {len(github_domains):,} domains")

            all_domains = list(set(all_domains))
            total_keywords = len(scanner.compound_keywords) + len(scanner.simple_keywords)
            await progress.update(
                f"Scanning {len(all_domains):,} domains against {total_keywords} keywords..."
            )

            # Progress callback for domain scanning
            async def on_scan_progress(scanned: int, total: int, matches: int):
                pct = int(100 * scanned / total)
                await progress.update(
                    f"Scanning domains: {pct}% ({scanned:,}/{total:,}) - {matches} matches so far"
                )

            priority_tlds = self._parse_list(self.valves.PRIORITY_TLDS)
            results = await scanner.scan(
                domains=all_domains,
                priority_tlds=priority_tlds,
                progress_callback=on_scan_progress
            )

            # Filter for website scanning:
            # - Compound Match (always high confidence)
            # - Simple Match on Priority TLD (medium confidence)
            high_medium = [
                r for r in results
                if r["detection"] in ("Compound Match", "Simple Match (Priority TLD)")
            ]

            if high_medium:
                await progress.update(
                    f"Found {len(results)} matches. Scanning {len(high_medium)} websites..."
                )

                # Get topic keywords and blacklist for website content scanning
                topic_keywords = self._parse_list(self.valves.TOPIC_KEYWORDS)
                website_blacklist = self._parse_list(self.valves.WEBSITE_BLACKLIST)
                website_scanner = _WebsiteScanner(
                    topic_keywords=topic_keywords,
                    blacklist_keywords=website_blacklist,
                    timeout=8,
                    max_concurrent=15  # Scan 15 websites concurrently
                )

                # Progress callback - update every 10 completions for cleaner output
                async def on_progress(completed: int, total: int, domain: str):
                    if completed % 10 == 0 or completed == total:
                        pct = int(100 * completed / total)
                        await progress.update(f"Scanning websites: {pct}% ({completed}/{total})")

                # Get domains to scan
                domains_to_scan = [r["domain"] for r in high_medium]

                # Scan all websites concurrently (non-blocking)
                website_results = await website_scanner.scan_multiple_async(
                    domains_to_scan,
                    progress_callback=on_progress
                )

                # Update results with website scan data
                excluded_count = 0
                for result in high_medium:
                    domain = result["domain"]
                    website_result = website_results.get(domain, {
                        "status": "error",
                        "keywords_found": [],
                        "blacklist_found": [],
                        "explanation": "Scan failed"
                    })
                    result["website_status"] = website_result["status"]
                    result["website_explanation"] = website_result["explanation"]
                    result["website_keywords"] = website_result["keywords_found"]
                    result["website_blacklist"] = website_result.get("blacklist_found", [])
                    if website_result["status"] == "excluded":
                        excluded_count += 1

                await progress.update(
                    f"Website scan complete. {len(high_medium) - excluded_count} PE-relevant, {excluded_count} excluded"
                )

            await progress.finish()

            return json.dumps({
                "status": "success",
                "stats": scanner.scan_stats,
                "results": results[:50]
            }, indent=2, ensure_ascii=False)

        except Exception as e:
            log.exception("Error during domain scan")
            await progress.error(str(e))
            return json.dumps({"status": "error", "error": str(e)})

    async def get_results_summary(
        self,
        __event_emitter__=None
    ) -> str:
        """
        Get markdown summary of last scan results.

        Returns:
            Markdown-formatted report
        """
        if self._scanner is None or not self._scanner.results:
            return "No scan results available. Run `scan_new_domains()` first."

        stats = self._scanner.scan_stats
        results = self._scanner.results

        # Group by detection type and website status
        compound = [r for r in results if r["detection"] == "Compound Match"]
        simple_priority = [r for r in results if r["detection"] == "Simple Match (Priority TLD)"]
        similarity = [r for r in results if "Similarity" in r["detection"]]

        # Separate excluded domains (VC/M&A detected on website)
        excluded = [r for r in results if r.get("website_status") == "excluded"]
        pe_relevant = [r for r in results if r.get("website_status") != "excluded"]

        # Filter compound/simple to exclude VC/M&A
        compound_clean = [r for r in compound if r.get("website_status") != "excluded"]
        simple_clean = [r for r in simple_priority if r.get("website_status") != "excluded"]

        lines = [
            "## Fund Domain Intelligence Report\n",
            f"**Scan Date:** {stats['scan_date']}",
            f"**Domains Scanned:** {stats['total_domains_scanned']:,}",
            f"**Matches Found:** {stats['matches_found']}",
            f"**PE Relevant:** {len(pe_relevant)} | **Excluded (VC/M&A):** {len(excluded)}",
            f"**High-Priority TLDs:** {', '.join(stats.get('high_priority_tlds', []))}\n",
        ]

        # Helper to format website status with icon
        def _format_website_status(r: Dict) -> str:
            status = r.get("website_status", "")
            explanation = r.get("website_explanation", "")
            if status == "keyword_found":
                return f"✅ {explanation}"
            elif status == "excluded":
                return f"🚫 {explanation}"
            elif status == "no_keywords":
                return f"🔍 {explanation}"
            elif status == "parked":
                return f"🅿️ {explanation}"
            elif status == "no_website":
                return f"❌ {explanation}"
            else:
                return "—"

        if compound_clean:
            lines.append("### High Confidence (Compound PE Keywords)")
            lines.append("| Domain | Keyword | TLD | Website Status |")
            lines.append("|--------|---------|-----|----------------|")
            for r in compound_clean[:20]:
                ws = _format_website_status(r)
                lines.append(f"| {r['domain']} | {r['keyword']} | .{r.get('tld', '')} | {ws} |")
            lines.append("")

        if simple_clean:
            lines.append("### Medium Confidence (Simple Keyword + Priority TLD)")
            lines.append("| Domain | Keyword | TLD | Website Status |")
            lines.append("|--------|---------|-----|----------------|")
            for r in simple_clean[:20]:
                ws = _format_website_status(r)
                lines.append(f"| {r['domain']} | {r['keyword']} | .{r.get('tld', '')} | {ws} |")
            lines.append("")

        if similarity:
            lines.append("### Similarity Matches (Review Recommended)")
            lines.append("| Domain | Keyword | Detection |")
            lines.append("|--------|---------|-----------|")
            for r in similarity[:10]:
                lines.append(f"| {r['domain']} | {r['keyword']} | {r['detection']} |")
            lines.append("")

        # Show excluded domains in a separate section
        if excluded:
            lines.append("### Excluded (VC/M&A/Non-PE Detected)")
            lines.append("| Domain | Matched Keyword | Excluded Because |")
            lines.append("|--------|-----------------|------------------|")
            for r in excluded[:15]:
                blacklist_kw = ", ".join(r.get("website_blacklist", [])[:2]) or "N/A"
                lines.append(f"| {r['domain']} | {r['keyword']} | {blacklist_kw} |")
            lines.append("")

        if not any([compound_clean, simple_clean, similarity]):
            lines.append("*No PE-relevant domains found.*")

        return "\n".join(lines)

    async def export_results_csv(
        self,
        __event_emitter__=None
    ) -> str:
        """
        Export scan results to CSV file.

        Returns:
            Path to exported file
        """
        if self._scanner is None or not self._scanner.results:
            return "No results to export. Run `scan_new_domains()` first."

        output_path = TOOL_DIR / f"fund_domains_{datetime.date.today().strftime('%Y_%m_%d')}.csv"

        fieldnames = [
            "domain", "keyword", "detection", "date", "priority_tld",
            "website_status", "website_explanation", "website_keywords"
        ]

        with open(output_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction="ignore")
            writer.writeheader()
            for r in self._scanner.results:
                row = r.copy()
                # Convert list to comma-separated string for CSV
                if "website_keywords" in row and isinstance(row["website_keywords"], list):
                    row["website_keywords"] = ", ".join(row["website_keywords"])
                writer.writerow(row)

        return f"Exported {len(self._scanner.results)} results to: {output_path}"

    async def get_configured_keywords(
        self,
        __event_emitter__=None
    ) -> str:
        """
        Get current configuration.

        Returns:
            JSON with current settings
        """
        return json.dumps({
            "compound_keywords": self._parse_list(self.valves.COMPOUND_KEYWORDS),
            "simple_keywords": self._parse_list(self.valves.SIMPLE_KEYWORDS),
            "topic_keywords": self._parse_list(self.valves.TOPIC_KEYWORDS),
            "website_blacklist": self._parse_list(self.valves.WEBSITE_BLACKLIST),
            "blacklist": self._parse_list(self.valves.BLACKLIST),
            "domain_blacklist": self._parse_list(self.valves.DOMAIN_BLACKLIST),
            "high_priority_tlds": self._parse_list(self.valves.HIGH_PRIORITY_TLDS),
            "priority_tlds": self._parse_list(self.valves.PRIORITY_TLDS),
            "nrd_source": self.valves.NRD_SOURCE,
            "similarity_mode": self.valves.SIMILARITY_MODE
        }, indent=2)
