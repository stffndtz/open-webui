"""
title: LinkedIn Tracker Tool
author: 42frontiers
author_url: https://42frontiers.com
version: 1.3.0
license: MIT
description: Monitors LinkedIn for job changes of tracked people and detects new PE/VC company pages. Uses PhantomBuster for all LinkedIn scraping (profiles, people search, company search).
requirements: requests, aiohttp
"""

import json
import logging
import asyncio
import aiohttp
import time
import concurrent.futures
from datetime import datetime
from typing import Dict, List, Optional, Literal, Union, Callable, Any, Awaitable
from collections import deque

from pydantic import BaseModel, Field
from pydantic.fields import FieldInfo

# Configure logging
log = logging.getLogger(__name__)


# =============================================================================
# PROGRESS TRACKER - Expandable status indicator for OpenWebUI
# =============================================================================


class ProgressTracker:
    """
    Simple progress tracker for OpenWebUI tools.

    Emits status events with a single-line description showing current step.
    OpenWebUI tool status events expect a simple description string.

    Usage:
        progress = ProgressTracker(event_emitter=__event_emitter__)
        await progress.update("Extracting metadata...")
        await progress.update("Extracting fund metrics...")
        await progress.finish()
    """

    def __init__(
        self,
        event_emitter: Optional[Callable[[Dict[str, Any]], Awaitable[None]]] = None,
    ) -> None:
        self._event_emitter = event_emitter
        self._started = time.perf_counter()
        self._current_step = ""
        self._done: bool = False

    async def update(self, description: str) -> None:
        """Update the status with a new description."""
        if self._done:
            return
        self._current_step = description
        log.info(description)
        await self._emit_status()

    async def finish(self) -> None:
        """Mark as done with elapsed time."""
        if self._done:
            return
        elapsed = time.perf_counter() - self._started
        self._current_step = f"✓ Completed in {elapsed:.1f}s"
        self._done = True
        log.info(self._current_step)
        await self._emit_status(done=True)

    async def error(self, error_msg: str) -> None:
        """Mark as failed with error message."""
        if self._done:
            return
        elapsed = time.perf_counter() - self._started
        self._current_step = f"✗ Error after {elapsed:.1f}s: {error_msg}"
        self._done = True
        log.error(self._current_step)
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
# LLM CALLING FUNCTIONS - Direct API calls to OpenAI/Anthropic/Azure
# =============================================================================


async def _call_llm_azure_openai(
    endpoint: str,
    api_key: str,
    deployment: str,
    api_version: str,
    system_prompt: str,
    user_message: str,
) -> str:
    """Call Azure OpenAI API directly."""
    url = f"{endpoint.rstrip('/')}/openai/deployments/{deployment}/chat/completions?api-version={api_version}"

    payload = {
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_message},
        ],
        "temperature": 0.1,
        "max_tokens": 4096,
    }

    headers = {
        "Content-Type": "application/json",
        "api-key": api_key,
    }

    log.info(f"Azure OpenAI: Calling {deployment}")

    timeout = aiohttp.ClientTimeout(total=120)  # 2 minute timeout
    async with aiohttp.ClientSession(timeout=timeout) as session:
        async with session.post(url, json=payload, headers=headers) as response:
            if response.status != 200:
                error_text = await response.text()
                raise RuntimeError(f"Azure OpenAI API error {response.status}: {error_text}")

            data = await response.json()

    if "choices" in data and len(data["choices"]) > 0:
        content = data["choices"][0].get("message", {}).get("content", "")
        if content:
            return content

    raise RuntimeError(f"Unexpected Azure OpenAI response: {data}")


async def _call_llm_openai(
    api_key: str,
    model: str,
    system_prompt: str,
    user_message: str,
) -> str:
    """Call OpenAI API directly."""
    url = "https://api.openai.com/v1/chat/completions"

    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_message},
        ],
        "temperature": 0.1,
        "max_tokens": 4096,
    }

    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {api_key}",
    }

    log.info(f"OpenAI: Calling {model}")

    timeout = aiohttp.ClientTimeout(total=120)  # 2 minute timeout
    async with aiohttp.ClientSession(timeout=timeout) as session:
        async with session.post(url, json=payload, headers=headers) as response:
            if response.status != 200:
                error_text = await response.text()
                raise RuntimeError(f"OpenAI API error {response.status}: {error_text}")

            data = await response.json()

    if "choices" in data and len(data["choices"]) > 0:
        content = data["choices"][0].get("message", {}).get("content", "")
        if content:
            return content

    raise RuntimeError(f"Unexpected OpenAI response: {data}")


async def _call_llm_anthropic(
    api_key: str,
    model: str,
    system_prompt: str,
    user_message: str,
) -> str:
    """Call Anthropic API directly."""
    url = "https://api.anthropic.com/v1/messages"

    payload = {
        "model": model,
        "max_tokens": 4096,
        "system": system_prompt,
        "messages": [
            {"role": "user", "content": user_message},
        ],
    }

    headers = {
        "Content-Type": "application/json",
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
    }

    log.info(f"Anthropic: Calling {model}")

    timeout = aiohttp.ClientTimeout(total=120)  # 2 minute timeout
    async with aiohttp.ClientSession(timeout=timeout) as session:
        async with session.post(url, json=payload, headers=headers) as response:
            if response.status != 200:
                error_text = await response.text()
                raise RuntimeError(f"Anthropic API error {response.status}: {error_text}")

            data = await response.json()

    if "content" in data and len(data["content"]) > 0:
        content = data["content"][0].get("text", "")
        if content:
            return content

    raise RuntimeError(f"Unexpected Anthropic response: {data}")


def _parse_json_response(response: str) -> dict:
    """Extract and parse JSON from LLM response, handling markdown code blocks."""
    import re
    # Try to find JSON in code blocks first
    json_match = re.search(r'```(?:json)?\s*([\s\S]*?)\s*```', response)
    if json_match:
        json_str = json_match.group(1)
    else:
        # Try to find raw JSON
        json_str = response.strip()

    try:
        return json.loads(json_str)
    except json.JSONDecodeError as e:
        log.warning(f"Failed to parse JSON response: {e}")
        log.debug(f"Response was: {response[:500]}...")
        return {"error": str(e), "raw_response": response}


# =============================================================================
# LLM PROMPTS FOR MANAGER EVALUATION
# =============================================================================

PROMPT_EVALUATE_PE_MANAGER = """You are a Private Equity expert evaluating whether a LinkedIn profile represents a PE professional that matches our investment criteria.

## Investment Criteria (INCLUDE)
We are looking for professionals at firms that do:
- Buyout (LBO, MBO, MBI)
- Buy and Build strategies
- Control investments
- Traditional Private Equity

## Exclusion Criteria (EXCLUDE)
Exclude professionals at firms focused on:
- Venture Capital (VC, seed, Series A/B/C)
- Real Estate / Property
- Infrastructure
- Growth Capital / Growth Equity
- Hedge Funds
- Private Credit / Private Debt
- Search Funds
- Mezzanine
- Natural Resources / Forestry / Commodities
- Music Rights / Royalties
- Consulting / Advisory only (no fund)
- Investment Banking (no fund)
- Accelerators / Incubators

## Evaluation Task
For each profile, determine:
1. Is this person at a PE firm that matches our criteria? (true/false)
2. What is your confidence? (high/medium/low)
3. Brief reason for your decision

Respond with ONLY valid JSON:
```json
{
  "evaluations": [
    {
      "linkedin_url": "profile URL",
      "name": "Person Name",
      "include": true,
      "confidence": "high",
      "reason": "Partner at buyout-focused PE firm"
    }
  ]
}
```

Important:
- If the profile mentions "venture", "VC", "seed", "growth equity", "real estate", etc. → EXCLUDE
- If unclear whether it's buyout vs VC, mark as include=true with confidence="low"
- Focus on the FIRM's investment strategy, not just the person's title
"""

# =============================================================================
# PHANTOMBUSTER CLIENT - LinkedIn scraping via PhantomBuster API
# =============================================================================


class _PhantomBusterClient:
    """
    Client for PhantomBuster API to run LinkedIn scraping phantoms.
    """

    BASE_URL = "https://api.phantombuster.com/api/v2"

    def __init__(self, api_key: str):
        self.api_key = api_key
        self.headers = {
            "X-Phantombuster-Key": api_key,
            "Content-Type": "application/json"
        }

    async def launch_agent(self, agent_id: str, argument: Dict = None) -> Dict:
        """Launch a PhantomBuster agent/phantom."""
        async with aiohttp.ClientSession() as session:
            payload = {"id": agent_id}
            if argument:
                payload["argument"] = argument

            async with session.post(
                f"{self.BASE_URL}/agents/launch",
                headers=self.headers,
                json=payload
            ) as response:
                return await response.json()

    async def get_agent_status(self, agent_id: str) -> Dict:
        """Get current status of an agent."""
        async with aiohttp.ClientSession() as session:
            async with session.get(
                f"{self.BASE_URL}/agents/fetch",
                headers=self.headers,
                params={"id": agent_id}
            ) as response:
                return await response.json()

    async def get_agent_output(self, agent_id: str) -> Dict:
        """Get the output of a completed agent run."""
        async with aiohttp.ClientSession() as session:
            async with session.get(
                f"{self.BASE_URL}/agents/fetch-output",
                headers=self.headers,
                params={"id": agent_id}
            ) as response:
                return await response.json()

    async def wait_for_completion(self, agent_id: str, timeout: int = 300) -> Dict:
        """Wait for an agent to complete and return its output."""
        start = datetime.now()
        while (datetime.now() - start).seconds < timeout:
            status = await self.get_agent_status(agent_id)
            if status.get("status") == "finished":
                return await self.get_agent_output(agent_id)
            elif status.get("status") == "error":
                return {"error": status.get("error", "Unknown error")}
            await asyncio.sleep(5)
        return {"error": "Timeout waiting for agent completion"}

    async def get_leads_from_list(
        self,
        list_id: str,
        limit: int = 100,
        offset: int = 0,
        order_by: str = "createdAt",
        order_direction: str = "desc"
    ) -> Dict:
        """
        Fetch leads from a PhantomBuster list using the Org Storage API.

        API: POST /org-storage/leads/by-list/{listId}
        Docs: https://hub.phantombuster.com/reference/post_org-storage-leads-by-list-listid

        Args:
            list_id: The PhantomBuster list ID (e.g., "3600558061474680")
            limit: Number of leads to fetch (default 100, max 1000)
            offset: Pagination offset
            order_by: Field to order by (createdAt, updatedAt, etc.)
            order_direction: asc or desc

        Returns:
            Dict with leads data (normalized to {"leads": [...]} format)
        """
        async with aiohttp.ClientSession() as session:
            payload = {
                "limit": min(limit, 1000),
                "offset": offset,
                "orderBy": order_by,
                "orderDirection": order_direction
            }

            async with session.post(
                f"{self.BASE_URL}/org-storage/leads/by-list/{list_id}",
                headers=self.headers,
                json=payload
            ) as response:
                if response.status != 200:
                    error_text = await response.text()
                    return {"error": f"API error {response.status}: {error_text}"}

                data = await response.json()

                # Normalize response: API may return list directly or dict with "leads" key
                if isinstance(data, list):
                    return {"leads": data}
                elif isinstance(data, dict):
                    # If it's a dict but doesn't have "leads", check for other possible keys
                    if "leads" in data:
                        return data
                    elif "data" in data:
                        return {"leads": data["data"]}
                    elif "results" in data:
                        return {"leads": data["results"]}
                    else:
                        # Return as-is, let caller handle it
                        return data
                else:
                    return {"error": f"Unexpected response type: {type(data)}"}

    async def get_all_leads_from_list(
        self,
        list_id: str,
        max_leads: int = 10000
    ) -> Union[List[Dict], Dict]:
        """
        Fetch all leads from a list with pagination and deduplication.

        Args:
            list_id: The PhantomBuster list ID
            max_leads: Maximum number of leads to fetch (safety limit)

        Returns:
            List of all leads (deduplicated), or Dict with "error" key if API error
        """
        all_leads = []
        seen_ids = set()  # Track unique leads to avoid duplicates
        offset = 0
        batch_size = 500
        api_error = None
        duplicates_skipped = 0

        while len(all_leads) < max_leads:
            result = await self.get_leads_from_list(
                list_id=list_id,
                limit=batch_size,
                offset=offset
            )

            if isinstance(result, dict) and "error" in result:
                api_error = result.get("error")
                log.error(f"API error fetching leads: {api_error}")
                break

            # Extract leads from normalized response
            leads = result.get("leads", []) if isinstance(result, dict) else []

            # Debug: log what we got on first batch
            if offset == 0:
                log.info(f"First batch: got {len(leads)} leads, result type: {type(result)}, keys: {list(result.keys()) if isinstance(result, dict) else 'N/A'}")
                if leads and isinstance(leads[0], dict):
                    log.info(f"First lead keys: {list(leads[0].keys())[:10]}")

            if not leads:
                break

            # Deduplicate leads based on unique identifier
            for lead in leads:
                if not isinstance(lead, dict):
                    continue

                # Create unique ID from available fields (prioritize URN, then URL, then name+company)
                lead_id = (
                    lead.get("linkedinProfileUrn") or
                    lead.get("linkedinProfileUrl") or
                    lead.get("profileUrl") or
                    f"{lead.get('firstName', '')}{lead.get('lastName', '')}{lead.get('companyName', '')}"
                )

                if lead_id and lead_id not in seen_ids:
                    seen_ids.add(lead_id)
                    all_leads.append(lead)
                elif lead_id:
                    duplicates_skipped += 1

            offset += len(leads)

            # If we got fewer than requested, we've reached the end
            if len(leads) < batch_size:
                break

        if duplicates_skipped > 0:
            log.info(f"Deduplicated: skipped {duplicates_skipped} duplicate leads")

        # If we got no leads and there was an error, return the error
        if not all_leads and api_error:
            return {"error": api_error}

        return all_leads




# =============================================================================
# TOOLS CLASS - Public interface for the LLM
# =============================================================================


class Tools:
    """
    LinkedIn Tracker Tool for Open WebUI.

    Monitors LinkedIn for:
    1. Job changes of tracked senior PE/VC staff (PhantomBuster Profile Scraper)
    2. New PE/VC company pages (PhantomBuster Companies Search)
    3. New people discovery via keyword search (PhantomBuster Search Export)

    Uses PhantomBuster exclusively for all LinkedIn scraping operations.
    """

    class Valves(BaseModel):
        """
        Admin-configurable settings for the LinkedIn Tracker Tool.
        All LinkedIn scraping is done via PhantomBuster.
        """
        PHANTOMBUSTER_API_KEY: str = Field(
            default="",
            description="PhantomBuster API key for LinkedIn scraping"
        )
        # ----- Leads List Configuration -----
        PHANTOMBUSTER_MANAGERS_LIST_ID: str = Field(
            default="",
            description="PhantomBuster list ID for new managers (e.g., '3600558061474680')"
        )
        PHANTOMBUSTER_TRACKED_EMPLOYEES_LIST_ID: str = Field(
            default="",
            description="PhantomBuster list ID for tracked company employees (e.g., '4453381301313414')"
        )
        PHANTOMBUSTER_COMPANIES_LIST_ID: str = Field(
            default="",
            description="PhantomBuster list ID for PE/VC companies (for new company detection)"
        )
        # ----- Company Detection Configuration -----
        COMPANY_NAME_KEYWORDS: str = Field(
            default="partners,capital,equity,fund,invest,holding,beteiligung,kapital",
            description="Keywords to match in company names (comma-separated)"
        )
        COMPANY_DETECTION_DAYS: int = Field(
            default=7,
            description="Number of days to look back for newly added companies"
        )
        # ----- People Search Configuration -----
        PEOPLE_SEARCH_QUERIES: str = Field(
            default="partner private equity europe,principal buyout fund europe,director private equity DACH,managing partner capital partners,investment director LBO,partner MBO MBI,director buy and build strategy,investment professional control investments,partner equity partners europe,principal portfolio company",
            description="Search queries for finding PE professionals (comma-separated)"
        )
        PEOPLE_TITLE_KEYWORDS: str = Field(
            default="partner,principal,director,managing director,investment director,investment professional,investment manager,general partner,managing partner,senior principal",
            description="Job titles to prioritize when filtering people results (comma-separated)"
        )
        PEOPLE_GEOGRAPHIC_FOCUS: str = Field(
            default="europe,germany,uk,france,netherlands,switzerland,nordics,DACH,benelux,austria,sweden,denmark,norway,finland",
            description="Geographic regions to focus on (comma-separated)"
        )
        # ----- Company Search Configuration -----
        COMPANY_SEARCH_QUERIES: str = Field(
            default="private equity europe,buyout fund,capital partners,equity partners,LBO fund,control investments,buy and build",
            description="Search queries for finding new PE/VC companies (comma-separated)"
        )
        PE_KEYWORDS: str = Field(
            default="buyout,mbo,mbi,buy and build,private equity,capital partners,equity partners,lbo,control investments,portfolio,investments,partners,equity,beteiligung,kapital",
            description="Keywords to identify PE/VC companies (comma-separated)"
        )
        PE_INDUSTRIES: str = Field(
            default="Venture Capital & Private Equity,Investment Management,Financial Services",
            description="LinkedIn industries to filter for PE/VC companies"
        )
        # ----- Blacklist (Exclusions) -----
        BLACKLIST_KEYWORDS: str = Field(
            default="venture capital,real estate,infrastructure,growth capital,blockchain,hedge fund,private credit,private debt,search fund,mezzanine,fixed income,commodities,natural resources,forestry,music rights,consulting,M&A advisor,investment banking,VC,crypto,seed,series a,proptech,fintech,cleantech,accelerator,incubator",
            description="Keywords to exclude from results (comma-separated)"
        )
        LOG_LEVEL: Literal["DEBUG", "INFO", "NONE"] = Field(
            default="INFO",
            description="Logging level for tool operations"
        )
        # ----- LLM Configuration for Manager Evaluation -----
        LLM_PROVIDER: str = Field(
            default="azure_openai",
            description="LLM provider for manager evaluation: 'azure_openai', 'openai', or 'anthropic'"
        )
        AZURE_OPENAI_ENDPOINT: str = Field(
            default="",
            description="Azure OpenAI endpoint URL (e.g., https://your-resource.openai.azure.com)"
        )
        AZURE_OPENAI_API_KEY: str = Field(
            default="",
            description="Azure OpenAI API key"
        )
        AZURE_OPENAI_DEPLOYMENT: str = Field(
            default="gpt-4o",
            description="Azure OpenAI deployment name (e.g., 'gpt-4o')"
        )
        AZURE_OPENAI_API_VERSION: str = Field(
            default="2024-08-01-preview",
            description="Azure OpenAI API version"
        )
        OPENAI_API_KEY: str = Field(
            default="",
            description="OpenAI API key (if using 'openai' provider)"
        )
        OPENAI_MODEL: str = Field(
            default="gpt-4o",
            description="OpenAI model name (e.g., 'gpt-4o', 'gpt-4-turbo')"
        )
        ANTHROPIC_API_KEY: str = Field(
            default="",
            description="Anthropic API key (if using 'anthropic' provider)"
        )
        ANTHROPIC_MODEL: str = Field(
            default="claude-sonnet-4-20250514",
            description="Anthropic model name (e.g., 'claude-sonnet-4-20250514')"
        )
        LLM_BATCH_SIZE: int = Field(
            default=10,
            description="Number of profiles to evaluate per LLM call (to manage token limits)"
        )

    def __init__(self):
        """Initialize the LinkedIn Tracker Tool."""
        self.valves = self.Valves()
        self._phantombuster: Optional[_PhantomBusterClient] = None
        self._session_logs: deque = deque(maxlen=500)
        self.file_handler = False
        self.citation = False

    # -------------------------------------------------------------------------
    # Internal Helpers
    # -------------------------------------------------------------------------

    def _get_phantombuster(self) -> Optional[_PhantomBusterClient]:
        """Get PhantomBuster client if configured."""
        api_key = self.valves.PHANTOMBUSTER_API_KEY
        if isinstance(api_key, FieldInfo):
            api_key = api_key.default
        if api_key:
            return _PhantomBusterClient(api_key)
        return None

    def _get_valve_str(self, value) -> str:
        """Get string value from valve, handling FieldInfo case."""
        if isinstance(value, FieldInfo):
            return (value.default or "").strip()
        return str(value).strip() if value else ""

    def _parse_list(self, value: str) -> List[str]:
        """Parse comma-separated string into list."""
        if isinstance(value, FieldInfo):
            value = value.default or ""
        return [x.strip().lower() for x in value.split(",") if x.strip()]

    def _log(self, message: str) -> None:
        """Add message to session log."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self._session_logs.append(f"[{timestamp}] {message}")
        log.info(message)

    async def _emit_status(
        self,
        emitter,
        message: str,
        done: bool = False
    ) -> None:
        """Emit status update to UI."""
        if emitter:
            await emitter({
                "type": "status",
                "data": {"description": message, "done": done}
            })
        self._log(message)

    async def _run_in_executor(self, func, *args):
        """
        Run a CPU-bound function in a thread pool to avoid blocking the event loop.
        This prevents the OpenWebUI app from freezing during intensive data processing.
        """
        loop = asyncio.get_event_loop()
        with concurrent.futures.ThreadPoolExecutor(max_workers=1) as executor:
            return await loop.run_in_executor(executor, func, *args)

    async def _call_llm(self, system_prompt: str, user_message: str) -> str:
        """
        Call the configured LLM provider directly.
        Uses the valve settings to determine which provider and credentials to use.
        """
        provider = self.valves.LLM_PROVIDER
        if isinstance(provider, FieldInfo):
            provider = provider.default or "azure_openai"
        provider = provider.lower()

        if provider == "azure_openai":
            endpoint = self.valves.AZURE_OPENAI_ENDPOINT
            api_key = self.valves.AZURE_OPENAI_API_KEY
            deployment = self.valves.AZURE_OPENAI_DEPLOYMENT
            api_version = self.valves.AZURE_OPENAI_API_VERSION

            if isinstance(endpoint, FieldInfo):
                endpoint = endpoint.default
            if isinstance(api_key, FieldInfo):
                api_key = api_key.default
            if isinstance(deployment, FieldInfo):
                deployment = deployment.default
            if isinstance(api_version, FieldInfo):
                api_version = api_version.default

            if not endpoint or not api_key:
                raise ValueError("Azure OpenAI endpoint and API key must be configured in valves")

            return await _call_llm_azure_openai(
                endpoint=endpoint,
                api_key=api_key,
                deployment=deployment,
                api_version=api_version,
                system_prompt=system_prompt,
                user_message=user_message,
            )
        elif provider == "openai":
            api_key = self.valves.OPENAI_API_KEY
            model = self.valves.OPENAI_MODEL

            if isinstance(api_key, FieldInfo):
                api_key = api_key.default
            if isinstance(model, FieldInfo):
                model = model.default

            if not api_key:
                raise ValueError("OpenAI API key must be configured in valves")

            return await _call_llm_openai(
                api_key=api_key,
                model=model,
                system_prompt=system_prompt,
                user_message=user_message,
            )
        elif provider == "anthropic":
            api_key = self.valves.ANTHROPIC_API_KEY
            model = self.valves.ANTHROPIC_MODEL

            if isinstance(api_key, FieldInfo):
                api_key = api_key.default
            if isinstance(model, FieldInfo):
                model = model.default

            if not api_key:
                raise ValueError("Anthropic API key must be configured in valves")

            return await _call_llm_anthropic(
                api_key=api_key,
                model=model,
                system_prompt=system_prompt,
                user_message=user_message,
            )
        else:
            raise ValueError(f"Unknown LLM provider: {provider}. Use 'azure_openai', 'openai', or 'anthropic'")

    # -------------------------------------------------------------------------
    # Non-blocking Processing Helpers (v1.3.0)
    # -------------------------------------------------------------------------

    def _filter_profiles_sync(
        self,
        profiles: List[Dict],
        title_keywords: List[str],
        blacklist_keywords: List[str]
    ) -> tuple:
        """
        Filter profiles by keywords (CPU-bound, runs in thread pool).
        Returns (filtered_profiles, excluded_by_blacklist, excluded_by_title).
        """
        filtered_profiles = []
        excluded_by_title = []
        excluded_by_blacklist = []

        for p in profiles:
            # Use both linkedinJobTitle and linkedinHeadline for filtering
            job_title_lower = p.get("job_title", "").lower()
            headline_lower = p.get("headline", "").lower()
            company_lower = p.get("company", "").lower()
            # Combined text for blacklist checking
            combined = f"{job_title_lower} {headline_lower} {company_lower}"

            # Check blacklist first (exclude VC, real estate, etc.)
            if any(bl.lower() in combined for bl in blacklist_keywords):
                excluded_by_blacklist.append(p)
                continue

            # Check if EITHER job_title OR headline contains any priority keywords
            title_match = any(kw.lower() in job_title_lower for kw in title_keywords)
            headline_match = any(kw.lower() in headline_lower for kw in title_keywords)

            if title_match or headline_match:
                filtered_profiles.append(p)
            else:
                excluded_by_title.append(p)

        return (filtered_profiles, excluded_by_blacklist, excluded_by_title)

    def _process_job_movements_batch_sync(
        self,
        leads_batch: List[Dict],
        clevel_titles: List[str],
        title_hierarchy: Dict[str, int],
        new_hire_indicators: List[str],
        promotion_indicators: List[str],
        months_back: int
    ) -> Dict:
        """
        Process a batch of leads for job movements (CPU-bound, runs in thread pool).
        Returns categorized movements.
        """
        c_level_movements = []
        new_hires = []
        promotions = []
        other_movements = []
        profiles_with_dates = 0
        profiles_without_dates = 0
        profiles_in_timeframe = 0
        skipped_no_url = 0

        def get_title_level(title_str: str) -> int:
            """Get seniority level from title string."""
            if not title_str:
                return 0
            title_lower = title_str.lower()
            max_level = 0
            for title_key, level in title_hierarchy.items():
                if title_key in title_lower:
                    max_level = max(max_level, level)
            return max_level

        def is_promotion(current_title: str, previous_title: str, same_company: bool) -> tuple:
            """Detect if there was a promotion based on title comparison."""
            if not same_company or not current_title or not previous_title:
                return (False, None)

            current_level = get_title_level(current_title)
            previous_level = get_title_level(previous_title)

            if current_level > previous_level and previous_level > 0:
                return (True, f"Promoted from '{previous_title}' to '{current_title}'")
            return (False, None)

        for lead in leads_batch:
            if not isinstance(lead, dict):
                continue

            # Extract lead data - check for nested 'data' field or use lead directly
            raw_data = lead.get("data")
            if isinstance(raw_data, dict):
                data = raw_data
            elif isinstance(raw_data, str):
                try:
                    data = json.loads(raw_data)
                except (json.JSONDecodeError, TypeError):
                    data = lead
            else:
                data = lead

            def get_field(*keys, default=""):
                for key in keys:
                    val = data.get(key) or lead.get(key)
                    if val:
                        return val
                return default

            name = get_field("fullName", "name") or \
                   f"{get_field('firstName')} {get_field('lastName')}".strip()
            linkedin_url = get_field(
                "linkedinProfileUrl", "linkedInProfileUrl",
                "linkedInUrl", "profileUrl", "url", "linkedin_url",
                "linkedinProfileUrn"
            )

            if linkedin_url and linkedin_url.startswith("urn:"):
                slug = get_field("publicIdentifier", "linkedinSlug", "vmid")
                if slug:
                    linkedin_url = f"https://www.linkedin.com/in/{slug}"

            company = get_field("companyName", "company")
            job_title = get_field("linkedinJobTitle", "title", "jobTitle")
            headline = get_field("linkedinHeadline", "headline", "occupation")
            title = headline or job_title
            location = get_field("location")
            description = get_field("description", "summary")

            job_date_range = get_field("linkedinJobDateRange", "jobDateRange")
            previous_job_date_range = get_field("linkedinPreviousJobDateRange", "previousJobDateRange")
            previous_company = get_field("previousCompanyName", "previousCompany")
            previous_title = get_field("linkedinPreviousJobTitle", "previousJobTitle")
            created_at = get_field("createdAt")
            updated_at = get_field("updatedAt")

            if not linkedin_url:
                skipped_no_url += 1
                continue

            # Parse job start date
            start_year, start_month, is_current = self._parse_job_date_range(job_date_range)
            is_recent_job_start = self._is_recent_start(job_date_range, months_back)
            had_job_change = (
                previous_company and
                previous_company.lower() != company.lower() and
                is_recent_job_start
            )

            combined_text = f"{title} {description}".lower()
            job_title_lower = job_title.lower() if job_title else ""
            title_lower = title.lower() if title else ""

            if start_year and start_month:
                start_date_str = f"{start_month:02d}/{start_year}"
            elif start_year:
                start_date_str = str(start_year)
            else:
                start_date_str = ""

            movement_data = {
                "name": name,
                "linkedin_url": linkedin_url,
                "company": company,
                "title": title,
                "job_title": job_title,
                "location": location,
                "job_date_range": job_date_range,
                "job_start_date": start_date_str,
                "is_current_job": is_current,
                "previous_company": previous_company,
                "previous_title": previous_title,
                "previous_job_date_range": previous_job_date_range,
                "created_at": created_at,
                "is_recent_job_start": is_recent_job_start,
                "had_job_change": had_job_change,
                "movement_type": [],
                "is_clevel": False
            }

            # Check for C-Level titles
            for ct in clevel_titles:
                if ct in job_title_lower or ct in title_lower:
                    movement_data["is_clevel"] = True
                    movement_data["clevel_title"] = ct
                    break

            is_job_change = had_job_change
            is_recent_start = is_recent_job_start

            same_company = (
                previous_company and
                company and
                (previous_company.lower() == company.lower() or
                 previous_company.lower() in company.lower() or
                 company.lower() in previous_company.lower())
            )
            is_title_promotion, promotion_reason = is_promotion(
                job_title or title,
                previous_title,
                same_company and is_recent_start
            )

            is_new_hire_text = any(ind in combined_text for ind in new_hire_indicators)
            is_promotion_text = any(ind in combined_text for ind in promotion_indicators)
            is_any_promotion = is_title_promotion or is_promotion_text

            if is_job_change:
                movement_data["movement_type"].append("job_change")
                movement_data["movement_reason"] = f"Changed from {previous_company} to {company}"
            if is_title_promotion:
                movement_data["movement_type"].append("promotion")
                movement_data["promotion_reason"] = promotion_reason
                if not movement_data.get("movement_reason"):
                    movement_data["movement_reason"] = promotion_reason
            if is_recent_start and not is_job_change and not is_title_promotion:
                movement_data["movement_type"].append("recent_start")
                movement_data["movement_reason"] = f"Started current role: {job_date_range}"
            if is_new_hire_text:
                movement_data["movement_type"].append("new_hire")
            if is_promotion_text and not is_title_promotion:
                movement_data["movement_type"].append("promotion_text")

            has_date_info = bool(job_date_range and start_year)

            if has_date_info:
                profiles_with_dates += 1
                if is_recent_job_start:
                    profiles_in_timeframe += 1
            else:
                profiles_without_dates += 1

            has_recent_movement = (
                is_job_change or
                is_recent_start or
                is_any_promotion or
                is_new_hire_text
            )
            has_current_job_no_date = not has_date_info and company and job_title

            # Categorize
            if movement_data["is_clevel"]:
                if has_recent_movement or has_current_job_no_date:
                    c_level_movements.append(movement_data)
            if is_job_change or is_new_hire_text:
                new_hires.append(movement_data)
            if is_any_promotion:
                promotions.append(movement_data)

            if has_recent_movement:
                if not movement_data["is_clevel"] and not is_job_change and not is_new_hire_text and not is_any_promotion:
                    other_movements.append(movement_data)
            elif has_current_job_no_date:
                movement_data["movement_type"].append("no_date_info")
                movement_data["movement_reason"] = "No date range available - included for review"
                other_movements.append(movement_data)

        return {
            "c_level_movements": c_level_movements,
            "new_hires": new_hires,
            "promotions": promotions,
            "other_movements": other_movements,
            "profiles_with_dates": profiles_with_dates,
            "profiles_without_dates": profiles_without_dates,
            "profiles_in_timeframe": profiles_in_timeframe,
            "skipped_no_url": skipped_no_url
        }

    def _aggregate_companies_sync(
        self,
        leads: List[Dict],
        name_keywords: List[str],
        pe_industries: List[str],
        blacklist: List[str],
        cutoff_date,
        days_back: int
    ) -> Dict:
        """
        Aggregate companies from employee profiles (CPU-bound, runs in thread pool).
        Returns categorized companies.
        """
        from datetime import timedelta

        companies_map: Dict[str, Dict] = {}

        for lead in leads:
            if not isinstance(lead, dict):
                continue

            def get_field(*keys, default=""):
                for key in keys:
                    val = lead.get(key)
                    if val:
                        return val
                return default

            company_name = get_field(
                "companyName", "company", "organizationName",
                "linkedinCompanyName", "currentCompany"
            )
            company_url = get_field(
                "linkedinCompanyUrl", "companyUrl", "companyLinkedinUrl",
                "organizationUrl", "currentCompanyUrl"
            )
            industry = get_field(
                "companyIndustry", "industry", "linkedinIndustry"
            )
            location = get_field(
                "location", "companyLocation", "headquarters"
            )
            created_at = get_field("createdAt", "addedAt", "created")

            employee_name = get_field("fullName", "name") or \
                f"{get_field('firstName')} {get_field('lastName')}".strip()
            employee_title = get_field(
                "linkedinJobTitle", "title", "jobTitle", "position"
            )
            employee_url = get_field(
                "linkedinProfileUrl", "linkedInProfileUrl",
                "linkedInUrl", "profileUrl", "url",
                "linkedinProfileUrn"
            )

            if not company_name:
                continue

            company_key = company_name.lower().strip()

            if company_key not in companies_map:
                companies_map[company_key] = {
                    "name": company_name,
                    "linkedin_url": company_url,
                    "industry": industry,
                    "location": location,
                    "employees": [],
                    "earliest_seen": created_at,
                    "latest_seen": created_at
                }

            if company_url and not companies_map[company_key]["linkedin_url"]:
                companies_map[company_key]["linkedin_url"] = company_url

            if industry and not companies_map[company_key]["industry"]:
                companies_map[company_key]["industry"] = industry

            companies_map[company_key]["employees"].append({
                "name": employee_name,
                "title": employee_title,
                "linkedin_url": employee_url,
                "added_at": created_at
            })

            if created_at:
                current_earliest = companies_map[company_key]["earliest_seen"]
                current_latest = companies_map[company_key]["latest_seen"]
                if not current_earliest or created_at < current_earliest:
                    companies_map[company_key]["earliest_seen"] = created_at
                if not current_latest or created_at > current_latest:
                    companies_map[company_key]["latest_seen"] = created_at

        # Now categorize companies
        new_companies = []
        existing_companies = []
        filtered_out = []

        for company_key, company_data in companies_map.items():
            company_name = company_data["name"]
            company_name_lower = company_name.lower()

            is_blacklisted = any(bl in company_name_lower for bl in blacklist)
            if is_blacklisted:
                filtered_out.append({
                    "name": company_name,
                    "reason": "Blacklisted keyword",
                    "employee_count": len(company_data["employees"])
                })
                continue

            has_name_keyword = any(kw in company_name_lower for kw in name_keywords)
            industry = company_data.get("industry") or ""
            industry_lower = industry.lower()
            has_pe_industry = any(ind.lower() in industry_lower for ind in pe_industries)

            matches_criteria = has_name_keyword or has_pe_industry

            if not matches_criteria:
                filtered_out.append({
                    "name": company_name,
                    "reason": "No matching keywords or industry",
                    "employee_count": len(company_data["employees"])
                })
                continue

            is_new = False
            earliest_seen = company_data.get("earliest_seen")

            if earliest_seen:
                try:
                    if isinstance(earliest_seen, str):
                        earliest_clean = earliest_seen.replace("Z", "+00:00")
                        if "." in earliest_clean:
                            parts = earliest_clean.split(".")
                            if len(parts[1]) > 6:
                                earliest_clean = parts[0] + "." + parts[1][:6] + parts[1][-6:]
                        earliest_dt = datetime.fromisoformat(earliest_clean.replace("+00:00", ""))
                        is_new = earliest_dt >= cutoff_date
                except (ValueError, TypeError):
                    is_new = True

            result_data = {
                "name": company_name,
                "linkedin_url": company_data.get("linkedin_url"),
                "industry": industry,
                "location": company_data.get("location"),
                "employee_count": len(company_data["employees"]),
                "first_seen": earliest_seen,
                "last_seen": company_data.get("latest_seen"),
                "sample_employees": [
                    {"name": e["name"], "title": e["title"]}
                    for e in company_data["employees"][:5]
                ],
                "matched_criteria": {
                    "name_keyword": has_name_keyword,
                    "pe_industry": has_pe_industry
                }
            }

            if is_new:
                new_companies.append(result_data)
            else:
                existing_companies.append(result_data)

        return {
            "companies_map": companies_map,
            "new_companies": new_companies,
            "existing_companies": existing_companies,
            "filtered_out": filtered_out
        }

    # -------------------------------------------------------------------------
    # Leads List Processing (PhantomBuster Org Storage API)
    # -------------------------------------------------------------------------

    async def filter_managers_from_list(
        self,
        __event_emitter__=None
    ) -> str:
        """
        Fetch and filter PE managers from a PhantomBuster leads list using LLM evaluation.

        Pulls leads from a PhantomBuster list and uses an LLM to intelligently evaluate
        each profile to determine if they match PE criteria:
        - Include: Buyout, MBO, MBI, LBO, Buy and Build, Private Equity
        - Exclude: VC, real estate, infrastructure, growth capital, credit, hedge funds, etc.

        Returns:
            JSON with filtered managers categorized by relevance
        """
        pb = self._get_phantombuster()
        if not pb:
            return json.dumps({
                "status": "error",
                "message": "PhantomBuster not configured. Set PHANTOMBUSTER_API_KEY in Valves."
            })

        # Get list ID from valve
        list_id = self._get_valve_str(self.valves.PHANTOMBUSTER_MANAGERS_LIST_ID)
        if not list_id:
            return json.dumps({
                "status": "error",
                "message": "No list ID configured. Set PHANTOMBUSTER_MANAGERS_LIST_ID in Valves."
            })

        progress = ProgressTracker(event_emitter=__event_emitter__)
        await progress.update(f"Fetching leads from list {list_id}...")

        try:
            # Fetch all leads from the list
            leads_result = await pb.get_all_leads_from_list(list_id)

            # Check for API error
            if isinstance(leads_result, dict) and "error" in leads_result:
                return json.dumps({
                    "status": "error",
                    "message": f"PhantomBuster API error: {leads_result['error']}",
                    "list_id": list_id
                })

            leads = leads_result if isinstance(leads_result, list) else []

            if not leads:
                return json.dumps({
                    "status": "error",
                    "message": "No leads found in list. Check that the list ID is correct and contains data.",
                    "list_id": list_id
                })

            await progress.update(f"Processing {len(leads)} leads...")

            # Debug: Log first lead structure to understand API response format
            if leads and isinstance(leads[0], dict):
                first_lead = leads[0]
                log.info(f"First lead keys: {list(first_lead.keys())[:15]}")
                # Check if data is nested
                if "data" in first_lead and isinstance(first_lead["data"], dict):
                    log.info(f"Nested data keys: {list(first_lead['data'].keys())[:15]}")

            # Extract profile data from leads
            # PhantomBuster field names (from CSV export):
            # linkedinProfileUrl, fullName, firstName, lastName, companyName,
            # linkedinJobTitle, linkedinHeadline, location, linkedinJobDateRange
            profiles = []
            skipped_count = 0
            for lead in leads:
                if not isinstance(lead, dict):
                    continue

                # PhantomBuster returns data at top level (not nested)
                # Try lead directly first, then check for nested 'data' field
                def get_field(*keys, default=""):
                    for key in keys:
                        # Try lead first (most common case)
                        val = lead.get(key)
                        if val:
                            return val
                    return default

                linkedin_url = get_field(
                    "linkedinProfileUrl", "linkedInProfileUrl",
                    "linkedInUrl", "profileUrl", "url",
                    "linkedinProfileUrn"  # PhantomBuster may return URN instead of URL
                )

                # If we got a URN, try to construct URL from slug or publicIdentifier
                if linkedin_url and linkedin_url.startswith("urn:"):
                    # Try to get the actual profile URL or slug
                    slug = get_field("publicIdentifier", "linkedinSlug", "vmid")
                    if slug:
                        linkedin_url = f"https://www.linkedin.com/in/{slug}"
                    # If no slug available, keep the URN as identifier (still valid for tracking)
                name = get_field("fullName", "name") or \
                       f"{lead.get('firstName', '')} {lead.get('lastName', '')}".strip()
                company = get_field("companyName", "company")
                job_title = get_field("linkedinJobTitle", "title", "jobTitle")
                headline = get_field("linkedinHeadline", "headline", "occupation")
                title = headline or job_title  # Combined for display
                location = get_field("location")
                description = get_field("description", "summary")

                if not linkedin_url or not name:
                    skipped_count += 1
                    continue

                profiles.append({
                    "linkedin_url": linkedin_url,
                    "name": name,
                    "company": company,
                    "title": title,
                    "job_title": job_title,  # linkedinJobTitle
                    "headline": headline,     # linkedinHeadline
                    "location": location,
                    "description": description
                })

            if not profiles:
                # Debug: show what the API returned
                sample_leads = []
                for i, lead in enumerate(leads[:3]):
                    if isinstance(lead, dict):
                        sample_leads.append({
                            "index": i,
                            "top_level_keys": list(lead.keys())[:20],
                            "has_linkedinProfileUrl": "linkedinProfileUrl" in lead,
                            "linkedinProfileUrl_value": str(lead.get("linkedinProfileUrl", ""))[:100],
                            "has_fullName": "fullName" in lead,
                            "fullName_value": str(lead.get("fullName", ""))[:100],
                            "sample_values": {k: str(v)[:80] for k, v in list(lead.items())[:6]}
                        })
                return json.dumps({
                    "status": "error",
                    "message": f"No valid profiles found. Skipped {skipped_count} leads (missing URL or name).",
                    "total_leads_received": len(leads),
                    "skipped_count": skipped_count,
                    "_debug_sample_leads": sample_leads
                }, indent=2)

            # Pre-filter by job title keywords (runs in thread pool to avoid blocking)
            # Uses PEOPLE_TITLE_KEYWORDS to only evaluate relevant profiles
            title_keywords = self._parse_list(self.valves.PEOPLE_TITLE_KEYWORDS)
            blacklist_keywords = self._parse_list(self.valves.BLACKLIST_KEYWORDS)

            await progress.update(f"Filtering {len(profiles)} profiles by keywords...")

            # Run CPU-bound filtering in thread pool to avoid blocking event loop
            filtered_profiles, excluded_by_blacklist, excluded_by_title = await self._run_in_executor(
                self._filter_profiles_sync,
                profiles,
                title_keywords,
                blacklist_keywords
            )

            await progress.update(
                f"Title filter: {len(filtered_profiles)} relevant, {len(excluded_by_blacklist)} blacklisted, {len(excluded_by_title)} other titles"
            )

            if not filtered_profiles:
                return json.dumps({
                    "status": "success",
                    "message": "No profiles matched title filter criteria",
                    "total_leads": len(leads),
                    "total_profiles": len(profiles),
                    "excluded_by_blacklist": len(excluded_by_blacklist),
                    "excluded_by_title": len(excluded_by_title),
                    "title_keywords_used": title_keywords[:10],
                    "sample_excluded_titles": [p["title"][:50] for p in excluded_by_title[:10]]
                }, indent=2)

            # Use filtered profiles for LLM evaluation
            profiles = filtered_profiles

            # Get batch size from valves
            batch_size = self.valves.LLM_BATCH_SIZE
            if isinstance(batch_size, FieldInfo):
                batch_size = batch_size.default or 10

            # Process profiles in batches using LLM
            all_evaluations = []
            num_batches = (len(profiles) + batch_size - 1) // batch_size

            for i in range(0, len(profiles), batch_size):
                batch = profiles[i:i + batch_size]
                batch_num = (i // batch_size) + 1

                await progress.update(
                    f"Evaluating batch {batch_num}/{num_batches} ({len(batch)} profiles)..."
                )

                # Format profiles for LLM
                profiles_text = "\n\n".join([
                    f"Profile {j+1}:\n"
                    f"- LinkedIn URL: {p['linkedin_url']}\n"
                    f"- Name: {p['name']}\n"
                    f"- Company: {p['company']}\n"
                    f"- Title/Headline: {p['title']}\n"
                    f"- Location: {p['location']}\n"
                    f"- Description: {p['description'][:500] if p['description'] else 'N/A'}"
                    for j, p in enumerate(batch)
                ])

                try:
                    response = await self._call_llm(
                        system_prompt=PROMPT_EVALUATE_PE_MANAGER,
                        user_message=f"Evaluate these {len(batch)} profiles:\n\n{profiles_text}"
                    )

                    result = _parse_json_response(response)
                    evaluations = result.get("evaluations", [])

                    # Match evaluations back to profiles
                    for eval_item in evaluations:
                        # Find matching profile
                        for p in batch:
                            if eval_item.get("linkedin_url") == p["linkedin_url"] or \
                               eval_item.get("name", "").lower() == p["name"].lower():
                                eval_item.update({
                                    "company": p["company"],
                                    "title": p["title"],
                                    "location": p["location"]
                                })
                                all_evaluations.append(eval_item)
                                break

                except Exception as e:
                    log.exception(f"LLM evaluation failed for batch {batch_num}: {e}")
                    # On error, mark batch as uncertain
                    for p in batch:
                        all_evaluations.append({
                            "linkedin_url": p["linkedin_url"],
                            "name": p["name"],
                            "company": p["company"],
                            "title": p["title"],
                            "location": p["location"],
                            "include": None,
                            "confidence": "error",
                            "reason": f"LLM evaluation failed: {str(e)}"
                        })

            # Categorize results
            relevant_managers = []
            filtered_out = []
            uncertain = []

            for eval_item in all_evaluations:
                if eval_item.get("include") is True:
                    confidence = eval_item.get("confidence", "medium")
                    eval_item["relevance"] = "high" if confidence == "high" else "medium"
                    relevant_managers.append(eval_item)

                elif eval_item.get("include") is False:
                    filtered_out.append(eval_item)

                else:
                    # None or error
                    uncertain.append(eval_item)

            # Categorize by confidence
            high_relevance = [m for m in relevant_managers if m.get("relevance") == "high"]
            medium_relevance = [m for m in relevant_managers if m.get("relevance") == "medium"]

            await progress.finish()

            # Build result
            result = {
                "status": "success",
                "source": "phantombuster_leads_list",
                "list_id": list_id,
                "evaluation_method": "llm",
                "total_leads": len(leads),
                "total_evaluated": len(all_evaluations),
                "summary": {
                    "high_relevance": len(high_relevance),
                    "medium_relevance": len(medium_relevance),
                    "filtered_out": len(filtered_out),
                    "uncertain": len(uncertain)
                },
                "high_relevance_managers": high_relevance[:50],
                "medium_relevance_managers": medium_relevance[:50],
                "filtered_out_sample": filtered_out[:20],
                "uncertain_sample": uncertain[:20]
            }

            # Emit citation with full results
            await self._emit_citation(
                __event_emitter__,
                f"PE Managers Filter Results - List {list_id}",
                json.dumps(result, indent=2)
            )

            return json.dumps(result, indent=2)

        except Exception as e:
            log.exception("Error filtering managers from list")
            await progress.error(str(e))
            return json.dumps({"status": "error", "message": str(e)})

    async def get_managers_report(
        self,
        __event_emitter__=None
    ) -> str:
        """
        Generate a markdown report of relevant PE managers from a PhantomBuster list.

        Filters managers based on PE criteria and formats as a readable report.

        Returns:
            Markdown-formatted report of relevant managers
        """
        # Get filtered data first
        result = await self.filter_managers_from_list(
            __event_emitter__=__event_emitter__
        )

        data = json.loads(result)

        if data.get("status") != "success":
            return f"Error: {data.get('message', 'Unknown error')}"

        summary = data.get("summary", {})
        high = data.get("high_relevance_managers", [])
        medium = data.get("medium_relevance_managers", [])
        filtered = data.get("filtered_out_sample", [])

        lines = [
            "## PE Manager Intelligence Report",
            "",
            f"**Source:** PhantomBuster List {data.get('list_id')}",
            f"**Total Leads Processed:** {data.get('total_leads', 0)}",
            "",
            "### Summary",
            f"- **High Relevance:** {summary.get('high_relevance', 0)} managers",
            f"- **Medium Relevance:** {summary.get('medium_relevance', 0)} managers",
            f"- **Filtered Out (VC/RE/etc):** {summary.get('filtered_out', 0)}",
            f"- **Uncertain/Low Signal:** {summary.get('uncertain', 0)}",
            ""
        ]

        if high:
            lines.append("### High Relevance (Buyout/PE Focus)")
            lines.append("| Name | Company | Title | Keywords | Location |")
            lines.append("|------|---------|-------|----------|----------|")
            for m in high[:25]:
                keywords = ", ".join(m.get("matched_keywords", [])[:3])
                lines.append(
                    f"| [{m['name']}]({m['linkedin_url']}) | {m['company'][:30]} | "
                    f"{m['title'][:40]} | {keywords} | {m.get('location', '')[:20]} |"
                )
            if len(high) > 25:
                lines.append(f"*...and {len(high) - 25} more*")
            lines.append("")

        if medium:
            lines.append("### Medium Relevance (Some PE Signals)")
            lines.append("| Name | Company | Title |")
            lines.append("|------|---------|-------|")
            for m in medium[:15]:
                lines.append(
                    f"| [{m['name']}]({m['linkedin_url']}) | {m['company'][:30]} | {m['title'][:50]} |"
                )
            if len(medium) > 15:
                lines.append(f"*...and {len(medium) - 15} more*")
            lines.append("")

        if filtered:
            lines.append("### Filtered Out (Sample)")
            lines.append("| Name | Company | Reason |")
            lines.append("|------|---------|--------|")
            for m in filtered[:10]:
                lines.append(
                    f"| {m['name']} | {m['company'][:30]} | {m.get('reason', 'N/A')} |"
                )
            lines.append("")

        lines.append("### Criteria Used")
        lines.append("**Include:** " + ", ".join(self._parse_list(self.valves.PE_KEYWORDS)[:10]))
        lines.append("")
        lines.append("**Exclude:** " + ", ".join(self._parse_list(self.valves.BLACKLIST_KEYWORDS)[:10]) + "...")

        return "\n".join(lines)

    async def _emit_citation(
        self,
        emitter,
        title: str,
        content: str
    ) -> None:
        """Emit a citation event to attach data to the chat."""
        if emitter:
            await emitter({
                "type": "citation",
                "data": {
                    "document": [content],
                    "metadata": [{"source": title}],
                    "source": {"name": title},
                }
            })

    async def debug_api_response(
        self,
        list_name: str = "managers",
        __event_emitter__=None
    ) -> str:
        """
        Debug function to see raw API response structure from PhantomBuster.
        Use this to understand what fields are available in the API response.

        Args:
            list_name: Name of the list to debug. Available lists:
                      - "managers" (PHANTOMBUSTER_MANAGERS_LIST_ID)
                      - "employees" (PHANTOMBUSTER_TRACKED_EMPLOYEES_LIST_ID)
                      - "companies" (PHANTOMBUSTER_COMPANIES_LIST_ID)

        Returns:
            JSON showing the structure of API response
        """
        pb = self._get_phantombuster()
        if not pb:
            return json.dumps({"error": "PhantomBuster not configured"})

        # Map list names to valve IDs
        list_mapping = {
            "managers": ("PHANTOMBUSTER_MANAGERS_LIST_ID", self.valves.PHANTOMBUSTER_MANAGERS_LIST_ID),
            "employees": ("PHANTOMBUSTER_TRACKED_EMPLOYEES_LIST_ID", self.valves.PHANTOMBUSTER_TRACKED_EMPLOYEES_LIST_ID),
            "companies": ("PHANTOMBUSTER_COMPANIES_LIST_ID", self.valves.PHANTOMBUSTER_COMPANIES_LIST_ID),
        }

        list_name_lower = list_name.lower().strip()
        if list_name_lower not in list_mapping:
            available = ", ".join(f'"{k}"' for k in list_mapping.keys())
            return json.dumps({"error": f"Unknown list name '{list_name}'. Available lists: {available}"})

        valve_name, valve_value = list_mapping[list_name_lower]
        list_id = self._get_valve_str(valve_value)
        if not list_id:
            return json.dumps({"error": f"{valve_name} not configured in Valves"})

        await self._emit_status(__event_emitter__, f"Fetching raw API response for {list_name} list ({list_id})...")

        # Fetch just a few leads to see structure
        result = await pb.get_leads_from_list(list_id, limit=3)

        # Analyze the response structure
        debug_info = {
            "result_type": type(result).__name__,
            "result_keys": list(result.keys()) if isinstance(result, dict) else None,
        }

        if isinstance(result, dict):
            if "error" in result:
                debug_info["error"] = result["error"]
            elif "leads" in result:
                leads = result["leads"]
                debug_info["leads_count"] = len(leads)
                if leads and isinstance(leads[0], dict):
                    first_lead = leads[0]
                    debug_info["first_lead_keys"] = list(first_lead.keys())
                    debug_info["first_lead_sample"] = {
                        k: str(v)[:100] for k, v in list(first_lead.items())[:10]
                    }
                    # Check nested data
                    if "data" in first_lead:
                        data_field = first_lead["data"]
                        debug_info["data_field_type"] = type(data_field).__name__
                        if isinstance(data_field, dict):
                            debug_info["data_field_keys"] = list(data_field.keys())
                            debug_info["data_field_sample"] = {
                                k: str(v)[:100] for k, v in list(data_field.items())[:10]
                            }
            else:
                # Unknown structure
                debug_info["raw_keys"] = list(result.keys())
                debug_info["raw_sample"] = {
                    k: str(v)[:100] for k, v in list(result.items())[:5]
                }

        await self._emit_status(__event_emitter__, "Debug complete", done=True)
        return json.dumps(debug_info, indent=2)

    async def get_list_as_table(
        self,
        list_name: str = "managers",
        columns: str = None,
        limit: int = 100,
        __event_emitter__=None
    ) -> str:
        """
        Fetch leads from a PhantomBuster list and return as a markdown table.

        This is a generic function to view the contents of a list without filtering.
        Useful for exploring data before applying filters.
        The full data is also attached as a citation to the chat.

        Args:
            list_name: Name of the list to fetch. Available lists:
                      - "managers" (PHANTOMBUSTER_MANAGERS_LIST_ID)
                      - "employees" (PHANTOMBUSTER_TRACKED_EMPLOYEES_LIST_ID)
                      - "companies" (PHANTOMBUSTER_COMPANIES_LIST_ID)
            columns: Comma-separated column names to display. Defaults to common fields.
                     Available: fullName, companyName, linkedinJobTitle, linkedinHeadline,
                     location, linkedinProfileUrl, linkedinJobDateRange, companyIndustry
            limit: Maximum number of rows to return (default 100)

        Returns:
            Markdown-formatted table of leads
        """
        pb = self._get_phantombuster()
        if not pb:
            return "Error: PhantomBuster not configured. Set PHANTOMBUSTER_API_KEY in Valves."

        # Map list names to valve IDs
        list_mapping = {
            "managers": ("PHANTOMBUSTER_MANAGERS_LIST_ID", self.valves.PHANTOMBUSTER_MANAGERS_LIST_ID),
            "employees": ("PHANTOMBUSTER_TRACKED_EMPLOYEES_LIST_ID", self.valves.PHANTOMBUSTER_TRACKED_EMPLOYEES_LIST_ID),
            "companies": ("PHANTOMBUSTER_COMPANIES_LIST_ID", self.valves.PHANTOMBUSTER_COMPANIES_LIST_ID),
        }

        list_name_lower = list_name.lower().strip()
        if list_name_lower not in list_mapping:
            available = ", ".join(f'"{k}"' for k in list_mapping.keys())
            return f"Error: Unknown list name '{list_name}'. Available lists: {available}"

        valve_name, valve_value = list_mapping[list_name_lower]
        list_id = self._get_valve_str(valve_value)
        if not list_id:
            return f"Error: {valve_name} not configured in Valves."

        await self._emit_status(__event_emitter__, f"Fetching leads from {list_name} list ({list_id})...")

        try:
            # Fetch leads from the list
            leads_result = await pb.get_all_leads_from_list(list_id, max_leads=limit)

            # Handle API error response
            if isinstance(leads_result, dict) and "error" in leads_result:
                return f"Error: PhantomBuster API error: {leads_result['error']}"

            # Ensure leads is a list
            leads = leads_result if isinstance(leads_result, list) else []

            if not leads:
                return f"No leads found in list {list_id}"

            await self._emit_status(__event_emitter__, f"Found {len(leads)} leads. Formatting table...")

            # Define default columns if not specified
            if columns:
                col_list = [c.strip() for c in columns.split(",")]
            else:
                col_list = ["fullName", "companyName", "linkedinJobTitle", "location"]

            # Column display names (friendlier headers)
            col_display = {
                "fullName": "Name",
                "firstName": "First Name",
                "lastName": "Last Name",
                "companyName": "Company",
                "linkedinJobTitle": "Job Title",
                "linkedinHeadline": "Headline",
                "location": "Location",
                "linkedinProfileUrl": "LinkedIn URL",
                "linkedinJobDateRange": "Job Date",
                "companyIndustry": "Industry",
                "previousCompanyName": "Previous Company",
                "linkedinPreviousJobTitle": "Previous Title",
                "connectionDegree": "Connection"
            }

            # Build table header
            headers = [col_display.get(c, c) for c in col_list]
            lines = [
                f"## Leads from List {list_id}",
                "",
                f"**Total Leads:** {len(leads)}",
                f"**Showing:** {min(len(leads), limit)}",
                "",
                "| " + " | ".join(headers) + " |",
                "| " + " | ".join(["---"] * len(headers)) + " |"
            ]

            # Build full data for citation (all columns, not truncated)
            citation_rows = []
            for lead in leads[:limit]:
                if not isinstance(lead, dict):
                    continue

                # Extract data - handle nested 'data' field
                data = lead.get("data", {}) if isinstance(lead.get("data"), dict) else lead

                row_values = []
                for col in col_list:
                    value = data.get(col) or lead.get(col) or ""
                    # Truncate long values and escape pipes for display table
                    if isinstance(value, str):
                        value = value.replace("|", "\\|")[:50]
                    row_values.append(str(value))

                lines.append("| " + " | ".join(row_values) + " |")

                # Full row data for citation (not truncated)
                citation_rows.append({
                    col: str(data.get(col) or lead.get(col) or "")
                    for col in col_list
                })

            # Emit citation with full data as JSON
            citation_content = json.dumps({
                "list_id": list_id,
                "total_leads": len(leads),
                "columns": col_list,
                "data": citation_rows
            }, indent=2)

            await self._emit_citation(
                __event_emitter__,
                f"PhantomBuster List {list_id}",
                citation_content
            )

            await self._emit_status(__event_emitter__, f"Table ready with {min(len(leads), limit)} rows", done=True)

            return "\n".join(lines)

        except Exception as e:
            log.exception("Error fetching list as table")
            return f"Error: {str(e)}"

    # -------------------------------------------------------------------------
    # Job Movements Analysis (C-Level, Hires, Promotions)
    # -------------------------------------------------------------------------

    async def analyze_job_movements(
        self,
        months_back: int = 6,
        __event_emitter__=None
    ) -> str:
        """
        Analyze job movements from a tracked company employees list.

        Detects new hires, promotions, and C-level movements.
        Uses linkedinJobDateRange to detect recent job changes within the specified timeframe.

        Args:
            months_back: How many months back to consider as "recent" movement (default 6)

        Returns:
            JSON with categorized job movements
        """
        pb = self._get_phantombuster()
        if not pb:
            return json.dumps({
                "status": "error",
                "message": "PhantomBuster not configured. Set PHANTOMBUSTER_API_KEY in Valves."
            })

        # Get list ID from valve
        list_id = self._get_valve_str(self.valves.PHANTOMBUSTER_TRACKED_EMPLOYEES_LIST_ID)
        if not list_id:
            return json.dumps({
                "status": "error",
                "message": "No list ID configured. Set PHANTOMBUSTER_TRACKED_EMPLOYEES_LIST_ID in Valves."
            })

        await self._emit_status(__event_emitter__, f"Fetching job movements from list {list_id}...")

        try:
            # Fetch all leads from the list
            leads_result = await pb.get_all_leads_from_list(list_id)

            # Check for API error
            if isinstance(leads_result, dict) and "error" in leads_result:
                return json.dumps({
                    "status": "error",
                    "message": f"PhantomBuster API error: {leads_result['error']}",
                    "list_id": list_id
                })

            leads = leads_result if isinstance(leads_result, list) else []

            if not leads:
                return json.dumps({
                    "status": "error",
                    "message": "No leads found in list. Check that the list ID is correct and contains data.",
                    "list_id": list_id
                })

            # Debug: capture sample of lead structures to understand API response format
            sample_leads = []
            for i, lead in enumerate(leads[:3]):
                if isinstance(lead, dict):
                    sample_leads.append({
                        "index": i,
                        "keys": list(lead.keys()) if lead else [],
                        "data_type": type(lead.get("data")).__name__ if "data" in lead else "N/A",
                        "data_keys": list(lead["data"].keys()) if isinstance(lead.get("data"), dict) else None,
                        "sample_values": {k: str(v)[:50] for k, v in list(lead.items())[:5]}
                    })

            await self._emit_status(__event_emitter__, f"Analyzing {len(leads)} profiles for job movements...")

            # C-Level and senior titles to flag
            clevel_titles = [
                "ceo", "cfo", "coo", "cto", "cio", "cmo", "chro", "cpo",
                "chief executive", "chief financial", "chief operating",
                "chief technology", "chief investment", "chief marketing",
                "managing director", "managing partner", "general partner",
                "president", "chairman", "vice president", "vp",
                "head of", "director", "partner", "principal"
            ]

            # Title hierarchy for promotion detection (higher number = more senior)
            title_hierarchy = {
                "ceo": 100, "chief executive": 100,
                "cfo": 95, "chief financial": 95,
                "coo": 95, "chief operating": 95,
                "cto": 95, "chief technology": 95,
                "cio": 95, "chief investment": 95,
                "cmo": 95, "chief marketing": 95,
                "chro": 95, "cpo": 95,
                "president": 90, "chairman": 90,
                "managing director": 85, "managing partner": 85,
                "general partner": 80, "senior partner": 78,
                "partner": 75,
                "vice president": 70, "vp": 70,
                "senior vice president": 72, "svp": 72,
                "executive vice president": 74, "evp": 74,
                "director": 65, "senior director": 68,
                "head of": 63, "senior principal": 62,
                "principal": 60, "senior manager": 55, "manager": 50,
                "senior associate": 45, "associate": 40,
                "senior analyst": 35, "analyst": 30,
                "assistant": 20, "intern": 10,
            }

            # Movement indicators in titles/descriptions
            new_hire_indicators = [
                "new role", "just joined", "excited to announce", "happy to share",
                "new position", "new chapter", "started as", "joining",
                "thrilled to join", "pleased to announce"
            ]

            promotion_indicators = [
                "promoted to", "promotion", "new role as", "elevated to",
                "appointed", "named as", "taking on", "stepping into"
            ]

            # Run CPU-bound job movements processing in thread pool to avoid blocking
            await self._emit_status(__event_emitter__, f"Processing {len(leads)} leads (non-blocking)...")

            batch_result = await self._run_in_executor(
                self._process_job_movements_batch_sync,
                leads,
                clevel_titles,
                title_hierarchy,
                new_hire_indicators,
                promotion_indicators,
                months_back
            )

            # Extract results from batch processing
            c_level_movements = batch_result["c_level_movements"]
            new_hires = batch_result["new_hires"]
            promotions = batch_result["promotions"]
            other_movements = batch_result["other_movements"]
            profiles_with_dates = batch_result["profiles_with_dates"]
            profiles_without_dates = batch_result["profiles_without_dates"]
            profiles_in_timeframe = batch_result["profiles_in_timeframe"]
            skipped_no_url = batch_result["skipped_no_url"]

            # Sort by job date range (most recent first)
            def sort_by_job_date(x):
                date_range = x.get("job_date_range", "")
                year, month, _ = self._parse_job_date_range(date_range)
                if year and month:
                    return (year * 12 + month)
                elif year:
                    return (year * 12)
                return 0

            for lst in [c_level_movements, new_hires, promotions, other_movements]:
                lst.sort(key=sort_by_job_date, reverse=True)

            # Count job changes specifically
            job_changes = [m for m in new_hires if "job_change" in m.get("movement_type", [])]

            processed_count = profiles_with_dates + profiles_without_dates
            await self._emit_status(
                __event_emitter__,
                f"Analyzed {len(leads)} profiles: {processed_count} processed, {skipped_no_url} skipped (no URL). "
                f"{profiles_with_dates} with dates ({profiles_in_timeframe} in timeframe), "
                f"{profiles_without_dates} without dates. Found {len(c_level_movements)} C-level, "
                f"{len(job_changes)} job changes, {len(promotions)} promotions.",
                done=True
            )

            return json.dumps({
                "status": "success",
                "source": "phantombuster_leads_list",
                "list_id": list_id,
                "total_leads": len(leads),
                "processed_profiles": processed_count,
                "skipped_no_url": skipped_no_url,
                "months_back": months_back,
                "date_coverage": {
                    "profiles_with_dates": profiles_with_dates,
                    "profiles_without_dates": profiles_without_dates,
                    "profiles_in_timeframe": profiles_in_timeframe
                },
                "summary": {
                    "c_level_movements": len(c_level_movements),
                    "job_changes": len(job_changes),
                    "new_hires": len(new_hires),
                    "promotions": len(promotions),
                    "other_movements": len(other_movements)
                },
                "c_level_movements": c_level_movements[:30],
                "job_changes": job_changes[:30],
                "new_hires": new_hires[:30],
                "promotions": promotions[:30],
                "other_movements": other_movements[:50],
                "_debug_sample_leads": sample_leads  # Debug: shows API response structure
            }, indent=2)

        except Exception as e:
            log.exception("Error analyzing job movements")
            return json.dumps({"status": "error", "message": str(e)})

    def _parse_job_date_range(self, date_range: str) -> tuple:
        """
        Parse LinkedIn job date range string to extract start year/month.

        Examples:
            "Jan 2025 - Present" -> (2025, 1, True)  # year, month, is_current
            "2024 - Present" -> (2024, None, True)
            "Mar 2023 - Dec 2024" -> (2023, 3, False)
        """
        if not date_range:
            return (None, None, False)

        is_current = "present" in date_range.lower()

        # Month mapping
        months = {
            "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
            "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12
        }

        # Try to extract start date (before the dash)
        parts = date_range.split("-")[0].strip()

        year = None
        month = None

        # Look for year (4 digits)
        import re
        year_match = re.search(r"(\d{4})", parts)
        if year_match:
            year = int(year_match.group(1))

        # Look for month (3-letter abbreviation)
        parts_lower = parts.lower()
        for mon_name, mon_num in months.items():
            if mon_name in parts_lower:
                month = mon_num
                break

        return (year, month, is_current)

    def _is_recent_start(self, date_range: str, months_back: int = 12) -> bool:
        """
        Check if job started within the last N months.

        Examples:
            "Jan 2025 - Present" with months_back=6 and today=Dec 2025:
            -> Started 11 months ago -> False (outside 6 months)

            "Jul 2025 - Present" with months_back=6 and today=Dec 2025:
            -> Started 5 months ago -> True (within 6 months)

            "2024 - Present" with months_back=18 and today=Dec 2025:
            -> Started ~12-24 months ago -> True (assume mid-year for year-only)
        """
        year, month, is_current = self._parse_job_date_range(date_range)

        if not year:
            return False

        now = datetime.now()
        current_year = now.year
        current_month = now.month

        # Calculate months ago
        if month:
            # Exact month known
            job_months = year * 12 + month
            current_months = current_year * 12 + current_month
            months_ago = current_months - job_months
        else:
            # Only have year - assume mid-year (June) for a reasonable estimate
            # This gives benefit of doubt for year-only dates
            job_months = year * 12 + 6  # June
            current_months = current_year * 12 + current_month
            months_ago = current_months - job_months

        # Job started within the timeframe (0 to months_back months ago)
        # Also allow slightly negative (future dates due to timezone issues)
        return -1 <= months_ago <= months_back

    # NOTE: CSV functions removed - using API functions instead

    async def get_job_movements_report(
        self,
        __event_emitter__=None
    ) -> str:
        """
        Generate a markdown report of job movements.

        Shows new hires, promotions, and C-level changes in a readable format.

        Returns:
            Markdown-formatted job movements report
        """
        result = await self.analyze_job_movements(
            __event_emitter__=__event_emitter__
        )

        data = json.loads(result)

        if data.get("status") != "success":
            return f"Error: {data.get('message', 'Unknown error')}"

        summary = data.get("summary", {})
        c_level = data.get("c_level_movements", [])
        new_hires = data.get("new_hires", [])
        promotions = data.get("promotions", [])

        lines = [
            "## Job Movements Intelligence Report",
            "",
            f"**Source:** PhantomBuster List {data.get('list_id')}",
            f"**Total Profiles Analyzed:** {data.get('total_leads', 0)}",
            ""
        ]

        lines.extend([
            "### Summary",
            f"- **C-Level Movements:** {summary.get('c_level_movements', 0)}",
            f"- **New Hires:** {summary.get('new_hires', 0)}",
            f"- **Promotions:** {summary.get('promotions', 0)}",
            ""
        ])

        # C-Level movements
        if c_level:
            lines.append("### C-Level & Senior Leadership Movements")
            lines.append("| Name | Company | Title | Start Date |")
            lines.append("|------|---------|-------|------------|")
            for m in c_level[:15]:
                start_date = m.get('job_start_date', '') or m.get('job_date_range', '')[:20] or '—'
                lines.append(
                    f"| [{m['name']}]({m['linkedin_url']}) | {m['company'][:30]} | {m['title'][:50]} | {start_date} |"
                )
            if len(c_level) > 15:
                lines.append(f"*...and {len(c_level) - 15} more*")
            lines.append("")

        # New hires
        if new_hires:
            lines.append("### Recent New Hires")
            lines.append("| Name | Company | Title | Start Date |")
            lines.append("|------|---------|-------|------------|")
            for m in new_hires[:15]:
                start_date = m.get('job_start_date', '') or m.get('job_date_range', '')[:20] or '—'
                lines.append(
                    f"| [{m['name']}]({m['linkedin_url']}) | {m['company'][:30]} | {m['title'][:50]} | {start_date} |"
                )
            if len(new_hires) > 15:
                lines.append(f"*...and {len(new_hires) - 15} more*")
            lines.append("")

        # Promotions
        if promotions:
            lines.append("### Recent Promotions")
            lines.append("| Name | Company | Title | Start Date |")
            lines.append("|------|---------|-------|------------|")
            for m in promotions[:15]:
                start_date = m.get('job_start_date', '') or m.get('job_date_range', '')[:20] or '—'
                lines.append(
                    f"| [{m['name']}]({m['linkedin_url']}) | {m['company'][:30]} | {m['title'][:50]} | {start_date} |"
                )
            if len(promotions) > 15:
                lines.append(f"*...and {len(promotions) - 15} more*")
            lines.append("")

        if not any([c_level, new_hires, promotions]):
            lines.append("*No significant job movements detected.*")

        return "\n".join(lines)

    # -------------------------------------------------------------------------
    # New Company Detection (via employee list aggregation)
    # -------------------------------------------------------------------------

    async def detect_new_companies(
        self,
        days_back: int = None,
        __event_emitter__=None
    ) -> str:
        """
        Detect newly discovered PE/VC companies by analyzing the employees list.

        Extracts unique companies from employee profiles and identifies "new" companies
        based on when employees at that company first appeared in the list. A company
        is considered "new" if ALL employees at that company were added within the
        detection window (meaning the company itself was just discovered).

        Filters by:
        - Company name keywords (e.g., "Partners", "Capital", "Equity")
        - Industry (PE/VC related)
        - Excludes blacklisted terms

        Args:
            days_back: Number of days to look back (default from COMPANY_DETECTION_DAYS valve)

        Returns:
            JSON with newly detected companies
        """
        pb = self._get_phantombuster()
        if not pb:
            return json.dumps({
                "status": "error",
                "message": "PhantomBuster not configured. Set PHANTOMBUSTER_API_KEY in Valves."
            })

        # Use employees list (fallback to companies list if configured)
        list_id = self._get_valve_str(self.valves.PHANTOMBUSTER_TRACKED_EMPLOYEES_LIST_ID)
        source_type = "employees"

        # If no employees list, try companies list
        if not list_id:
            list_id = self._get_valve_str(self.valves.PHANTOMBUSTER_COMPANIES_LIST_ID)
            source_type = "companies"

        if not list_id:
            return json.dumps({
                "status": "error",
                "message": "No list configured. Set PHANTOMBUSTER_TRACKED_EMPLOYEES_LIST_ID or PHANTOMBUSTER_COMPANIES_LIST_ID in Valves."
            })

        # Get detection window
        if days_back is None:
            days_back = self.valves.COMPANY_DETECTION_DAYS
            if isinstance(days_back, FieldInfo):
                days_back = days_back.default or 7

        await self._emit_status(__event_emitter__, f"Fetching {source_type} from list {list_id}...")

        try:
            # Fetch all leads from the list
            leads_result = await pb.get_all_leads_from_list(list_id)

            # Check for API error
            if isinstance(leads_result, dict) and "error" in leads_result:
                return json.dumps({
                    "status": "error",
                    "message": f"PhantomBuster API error: {leads_result['error']}",
                    "list_id": list_id
                })

            leads = leads_result if isinstance(leads_result, list) else []

            if not leads:
                return json.dumps({
                    "status": "error",
                    "message": f"No {source_type} found in list. Check that the list ID is correct.",
                    "list_id": list_id
                })

            await self._emit_status(
                __event_emitter__,
                f"Extracting companies from {len(leads)} {source_type}..."
            )

            # Calculate cutoff date
            from datetime import timedelta
            cutoff_date = datetime.now() - timedelta(days=days_back)
            cutoff_timestamp = cutoff_date.isoformat()

            # Get filter keywords
            name_keywords = self._parse_list(self.valves.COMPANY_NAME_KEYWORDS)
            pe_industries = self._parse_list(self.valves.PE_INDUSTRIES)
            blacklist = self._parse_list(self.valves.BLACKLIST_KEYWORDS)

            # Run CPU-bound company aggregation in thread pool to avoid blocking
            await self._emit_status(
                __event_emitter__,
                f"Aggregating companies from {len(leads)} {source_type} (non-blocking)..."
            )

            aggregation_result = await self._run_in_executor(
                self._aggregate_companies_sync,
                leads,
                name_keywords,
                pe_industries,
                blacklist,
                cutoff_date,
                days_back
            )

            # Extract results
            companies_map = aggregation_result["companies_map"]
            new_companies = aggregation_result["new_companies"]
            existing_companies = aggregation_result["existing_companies"]
            filtered_out = aggregation_result["filtered_out"]

            await self._emit_status(
                __event_emitter__,
                f"Found {len(companies_map)} unique companies."
            )

            # Sort new companies by first_seen (most recent first)
            new_companies.sort(
                key=lambda x: x.get("first_seen") or "",
                reverse=True
            )

            await self._emit_status(
                __event_emitter__,
                f"Found {len(new_companies)} new companies in the last {days_back} days. "
                f"{len(existing_companies)} existing, {len(filtered_out)} filtered out.",
                done=True
            )

            result = {
                "status": "success",
                "source": f"phantombuster_{source_type}_list",
                "list_id": list_id,
                "detection_window_days": days_back,
                "cutoff_date": cutoff_timestamp,
                "total_profiles": len(leads),
                "total_companies": len(companies_map),
                "summary": {
                    "new_companies": len(new_companies),
                    "existing_companies": len(existing_companies),
                    "filtered_out": len(filtered_out)
                },
                "new_companies": new_companies[:100],
                "filtered_out_sample": filtered_out[:20]
            }

            # Emit citation with full results
            await self._emit_citation(
                __event_emitter__,
                f"New PE Companies - Last {days_back} days",
                json.dumps(result, indent=2)
            )

            return json.dumps(result, indent=2)

        except Exception as e:
            log.exception("Error detecting new companies")
            return json.dumps({"status": "error", "message": str(e)})

    async def get_new_companies_report(
        self,
        days_back: int = None,
        __event_emitter__=None
    ) -> str:
        """
        Generate a markdown report of newly detected PE/VC companies.

        Extracts unique companies from the employees list and identifies which
        ones were first discovered within the detection window.

        Args:
            days_back: Number of days to look back (default from valve)

        Returns:
            Markdown-formatted report of new companies
        """
        result = await self.detect_new_companies(
            days_back=days_back,
            __event_emitter__=__event_emitter__
        )

        data = json.loads(result)

        if data.get("status") != "success":
            return f"Error: {data.get('message', 'Unknown error')}"

        summary = data.get("summary", {})
        new_companies = data.get("new_companies", [])
        days = data.get("detection_window_days", 7)

        lines = [
            "## New PE/VC Companies Report",
            "",
            f"**Source:** {data.get('source', 'PhantomBuster')} (List {data.get('list_id')})",
            f"**Detection Window:** Last {days} days",
            f"**Cutoff Date:** {data.get('cutoff_date', 'N/A')[:10]}",
            f"**Total Profiles Analyzed:** {data.get('total_profiles', 0):,}",
            f"**Unique Companies Found:** {data.get('total_companies', 0):,}",
            ""
        ]

        lines.extend([
            "### Summary",
            f"- **New Companies:** {summary.get('new_companies', 0)}",
            f"- **Existing Companies:** {summary.get('existing_companies', 0)}",
            f"- **Filtered Out:** {summary.get('filtered_out', 0)}",
            ""
        ])

        if new_companies:
            lines.append("### Newly Detected Companies")
            lines.append("| Company | Industry | Tracked Employees | First Seen |")
            lines.append("|---------|----------|-------------------|------------|")

            for c in new_companies[:30]:
                name = c.get("name", "Unknown")
                url = c.get("linkedin_url", "")
                industry = (c.get("industry") or "—")[:25]
                emp_count = c.get("employee_count") or 0
                first_seen = (c.get("first_seen") or "")[:10] or "—"

                if url:
                    name_cell = f"[{name[:35]}]({url})"
                else:
                    name_cell = name[:35]

                lines.append(f"| {name_cell} | {industry} | {emp_count} | {first_seen} |")

            if len(new_companies) > 30:
                lines.append(f"*...and {len(new_companies) - 30} more*")
            lines.append("")

            # Show sample employees for top companies
            if new_companies[:5]:
                lines.append("### Sample Employees at New Companies")
                for c in new_companies[:5]:
                    company_name = c.get("name", "Unknown")
                    sample_emps = c.get("sample_employees", [])
                    if sample_emps:
                        emp_list = ", ".join([
                            f"{e.get('name', '?')} ({e.get('title', '?')[:30]})"
                            for e in sample_emps[:3]
                        ])
                        lines.append(f"- **{company_name}**: {emp_list}")
                lines.append("")

        else:
            lines.append("### No New Companies")
            lines.append(f"*No new PE/VC companies detected in the last {days} days.*")
            lines.append("")

        # Add criteria info
        lines.extend([
            "### Detection Criteria",
            f"**Name Keywords:** {', '.join(self._parse_list(self.valves.COMPANY_NAME_KEYWORDS)[:8])}...",
            f"**Industries:** {', '.join(self._parse_list(self.valves.PE_INDUSTRIES)[:3])}",
            ""
        ])

        return "\n".join(lines)
