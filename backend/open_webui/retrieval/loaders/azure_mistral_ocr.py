import requests
import aiohttp
import asyncio
import logging
import os
import sys
import time
import base64
from typing import List, Dict, Any
from contextlib import asynccontextmanager

from langchain_core.documents import Document
from open_webui.env import SRC_LOG_LEVELS, GLOBAL_LOG_LEVEL

logging.basicConfig(stream=sys.stdout, level=GLOBAL_LOG_LEVEL)
log = logging.getLogger(__name__)
log.setLevel(SRC_LOG_LEVELS["RAG"])


class AzureMistralOCRLoader:
    """
    Azure AI Foundry Mistral OCR loader for document processing.
    Loads documents by processing them through Azure AI Foundry with Mistral OCR models.
    
    This loader works with Azure AI Foundry endpoints that have Mistral models deployed,
    specifically designed for OCR and document analysis tasks.
    """

    def __init__(
        self,
        api_key: str,
        endpoint_url: str,
        model_name: str,
        file_path: str,
        timeout: int = 300,  # 5 minutes default
        max_retries: int = 3,
        enable_debug_logging: bool = False,
    ):
        """
        Initializes the Azure Mistral OCR loader.

        Args:
            api_key: Your Azure API key.
            endpoint_url: The Azure AI Foundry endpoint URL.
            model_name: The name of the deployed Mistral model (e.g., "mistral-document-ai-2505").
            file_path: The local path to the PDF file to process.
            timeout: Request timeout in seconds.
            max_retries: Maximum number of retry attempts.
            enable_debug_logging: Enable detailed debug logs.
        """
        if not api_key:
            raise ValueError("Azure API key cannot be empty.")
        if not endpoint_url:
            raise ValueError("Azure endpoint URL cannot be empty.")
        if not model_name:
            raise ValueError("Model name cannot be empty.")
        if not os.path.exists(file_path):
            raise FileNotFoundError(f"File not found at {file_path}")

        self.api_key = api_key
        self.endpoint_url = endpoint_url.rstrip("/")
        self.model_name = model_name
        self.file_path = file_path
        self.timeout = timeout
        self.max_retries = max_retries
        self.debug = enable_debug_logging

        # Pre-compute file info
        self.file_name = os.path.basename(file_path)
        self.file_size = os.path.getsize(file_path)

        # Headers for Azure API
        self.headers = {
            "Authorization": f"Bearer {self.api_key}",
            "Content-Type": "application/json",
            "Accept": "application/json",
            "User-Agent": "OpenWebUI-AzureMistralOCRLoader/1.0",
        }

    def _debug_log(self, message: str, *args) -> None:
        """Conditional debug logging for performance."""
        if self.debug:
            log.debug(message, *args)

    def _handle_response(self, response: requests.Response) -> Dict[str, Any]:
        """Checks response status and returns JSON content."""
        try:
            response.raise_for_status()
            if response.status_code == 204 or not response.content:
                return {}
            return response.json()
        except requests.exceptions.HTTPError as http_err:
            log.error(f"HTTP error occurred: {http_err} - Response: {response.text}")
            raise
        except requests.exceptions.RequestException as req_err:
            log.error(f"Request exception occurred: {req_err}")
            raise
        except ValueError as json_err:
            log.error(f"JSON decode error: {json_err} - Response: {response.text}")
            raise

    async def _handle_response_async(
        self, response: aiohttp.ClientResponse
    ) -> Dict[str, Any]:
        """Async version of response handling."""
        try:
            response.raise_for_status()
            
            content_type = response.headers.get("content-type", "")
            if "application/json" not in content_type:
                if response.status == 204:
                    return {}
                text = await response.text()
                raise ValueError(
                    f"Unexpected content type: {content_type}, body: {text[:200]}..."
                )

            return await response.json()

        except aiohttp.ClientResponseError as e:
            error_text = await response.text() if response else "No response"
            log.error(f"HTTP {e.status}: {e.message} - Response: {error_text[:500]}")
            raise
        except aiohttp.ClientError as e:
            log.error(f"Client error: {e}")
            raise
        except Exception as e:
            log.error(f"Unexpected error processing response: {e}")
            raise

    def _is_retryable_error(self, error: Exception) -> bool:
        """Determines if an error is retryable based on its type and status code."""
        if isinstance(error, requests.exceptions.ConnectionError):
            return True
        if isinstance(error, requests.exceptions.Timeout):
            return True
        if isinstance(error, requests.exceptions.HTTPError):
            if hasattr(error, "response") and error.response is not None:
                status_code = error.response.status_code
                return status_code >= 500 or status_code == 429
            return False
        if isinstance(
            error, (aiohttp.ClientConnectionError, aiohttp.ServerTimeoutError)
        ):
            return True
        if isinstance(error, aiohttp.ClientResponseError):
            return error.status >= 500 or error.status == 429
        return False

    def _retry_request_sync(self, request_func, *args, **kwargs):
        """Synchronous retry logic with intelligent error classification."""
        for attempt in range(self.max_retries):
            try:
                return request_func(*args, **kwargs)
            except Exception as e:
                if attempt == self.max_retries - 1 or not self._is_retryable_error(e):
                    raise

                wait_time = min((2**attempt) + 0.5, 30)
                log.warning(
                    f"Retryable error (attempt {attempt + 1}/{self.max_retries}): {e}. "
                    f"Retrying in {wait_time}s..."
                )
                time.sleep(wait_time)

    async def _retry_request_async(self, request_func, *args, **kwargs):
        """Async retry logic with intelligent error classification."""
        for attempt in range(self.max_retries):
            try:
                return await request_func(*args, **kwargs)
            except Exception as e:
                if attempt == self.max_retries - 1 or not self._is_retryable_error(e):
                    raise

                wait_time = min((2**attempt) + 0.5, 30)
                log.warning(
                    f"Retryable error (attempt {attempt + 1}/{self.max_retries}): {e}. "
                    f"Retrying in {wait_time}s..."
                )
                await asyncio.sleep(wait_time)

    def _encode_file_to_base64(self) -> str:
        """Encodes the PDF file to base64 for Azure API."""
        log.info(f"Encoding file to base64: {self.file_name}")
        try:
            with open(self.file_path, "rb") as f:
                file_content = f.read()
                base64_content = base64.b64encode(file_content).decode("utf-8")
                log.info(f"File encoded successfully. Size: {len(base64_content)} characters")
                return base64_content
        except Exception as e:
            log.error(f"Failed to encode file to base64: {e}")
            raise

    def _process_ocr(self, base64_content: str) -> Dict[str, Any]:
        """Sends the base64 content to the Azure OCR endpoint for processing."""
        log.info("Processing OCR via Azure AI Foundry Mistral API")
        
        # Construct the full endpoint URL
        ocr_url = f"{self.endpoint_url}/providers/mistral/azure/ocr"
        
        payload = {
            "model": self.model_name,
            "document": {
                "type": "document_url",
                "document_url": f"data:application/pdf;base64,{base64_content}"
            },
            "bbox_annotation_format": {
                "type": "json_schema",
                "json_schema": {
                    "schema": {
                        "properties": {
                            "document_type": {"title": "Document_Type", "description": "The type of the image.", "type": "string"},
                            "short_description": {"title": "Short_Description", "description": "A description in English describing the image.", "type": "string"},
                            "summary": {"title": "Summary", "description": "Summarize the image.", "type": "string"}
                        },
                        "required": ["document_type", "short_description", "summary"],
                        "title": "BBOXAnnotation",
                        "type": "object",
                        "additionalProperties": False
                    },
                    "name": "image_annotation",
                    "strict": True
                }
            },
            "document_annotation_format": {
                "type": "json_schema",
                "json_schema": {
                    "schema": {
                        "properties": {
                            "summary": {"title": "Summary", "type": "string"},
                            "authors": {"title": "Authors", "type": "string"},
                            "language": {"title": "Language", "type": "string", "description": "The language of the document in ISO 639-1 code format (e.g., 'en', 'fr')."},
                            "chapter_titles": {"title": "Chapter_Titles", "type": "string"},
                            "urls": {"title": "urls", "type": "string"}
                        },
                        "required": ["summary", "language", "chapter_titles", "urls"],
                        "title": "DocumentAnnotation",
                        "type": "object",
                        "additionalProperties": False
                    },
                    "name": "document_annotation",
                    "strict": True
                }
            },
            "include_image_base64": True
        }

        def ocr_request():
            response = requests.post(
                ocr_url, 
                headers=self.headers, 
                json=payload, 
                timeout=self.timeout
            )
            return self._handle_response(response)

        try:
            ocr_response = self._retry_request_sync(ocr_request)
            log.info("OCR processing completed.")
            # Convert response to JSON format
            # response_dict = json.loads(ocr_response.model_dump_json())
            # print(json.dumps(response_dict, indent=4))

            self._debug_log("OCR response: %s", ocr_response)
            return ocr_response
        except Exception as e:
            log.error(f"Failed during OCR processing: {e}")
            raise

    async def _process_ocr_async(
        self, session: aiohttp.ClientSession, base64_content: str
    ) -> Dict[str, Any]:
        """Async OCR processing."""
        ocr_url = f"{self.endpoint_url}/providers/mistral/azure/ocr"
        
        payload = {
            "model": self.model_name,
            "document": {
                "type": "document_url",
                "document_url": f"data:application/pdf;base64,{base64_content}"
            },
            "include_image_base64": True
        }

        async def ocr_request():
            log.info("Starting OCR processing via Azure AI Foundry Mistral API")
            start_time = time.time()

            async with session.post(
                ocr_url,
                json=payload,
                headers=self.headers,
                timeout=aiohttp.ClientTimeout(total=self.timeout),
            ) as response:
                ocr_response = await self._handle_response_async(response)

            processing_time = time.time() - start_time
            log.info(f"OCR processing completed in {processing_time:.2f}s")

            return ocr_response

        return await self._retry_request_async(ocr_request)

    @asynccontextmanager
    async def _get_session(self):
        """Context manager for HTTP session with optimized settings."""
        connector = aiohttp.TCPConnector(
            limit=20,
            limit_per_host=10,
            ttl_dns_cache=600,
            use_dns_cache=True,
            keepalive_timeout=60,
            enable_cleanup_closed=True,
            force_close=False,
            resolver=aiohttp.AsyncResolver(),
        )

        timeout = aiohttp.ClientTimeout(
            total=self.timeout,
            connect=30,
            sock_read=60,
        )

        async with aiohttp.ClientSession(
            connector=connector,
            timeout=timeout,
            headers={"User-Agent": "OpenWebUI-AzureMistralOCRLoader/1.0"},
            raise_for_status=False,
            trust_env=True,
        ) as session:
            yield session

    def _process_results(self, ocr_response: Dict[str, Any]) -> List[Document]:
        """Process OCR results into Document objects."""
        # Azure Mistral OCR response format may differ from standard Mistral
        # We need to handle the response structure based on the actual API response
        
        # Check for different possible response structures
        pages_data = None
        if "pages" in ocr_response:
            pages_data = ocr_response.get("pages")
        elif "content" in ocr_response:
            # If it's a single content response, wrap it in a pages structure
            pages_data = [{"markdown": ocr_response.get("content"), "index": 0}]
        elif "text" in ocr_response:
            # Alternative text field
            pages_data = [{"markdown": ocr_response.get("text"), "index": 0}]
        else:
            # Try to extract any text content from the response
            log.warning(f"Unexpected response structure: {list(ocr_response.keys())}")
            # Look for any field that might contain text content
            for key, value in ocr_response.items():
                if isinstance(value, str) and len(value) > 50:  # Likely text content
                    pages_data = [{"markdown": value, "index": 0}]
                    break

        if not pages_data:
            log.warning("No pages or content found in OCR response.")
            return [
                Document(
                    page_content="No text content found",
                    metadata={"error": "no_pages", "file_name": self.file_name},
                )
            ]

        documents = []
        total_pages = len(pages_data)
        skipped_pages = 0

        # Process pages
        for page_data in pages_data:
            page_content = page_data.get("markdown") or page_data.get("content") or page_data.get("text")
            page_index = page_data.get("index", 0)

            if page_content is None:
                skipped_pages += 1
                self._debug_log(
                    f"Skipping page due to missing content. Data keys: {list(page_data.keys())}"
                )
                continue

            # Clean up content
            if isinstance(page_content, str):
                cleaned_content = page_content.strip()
            else:
                cleaned_content = str(page_content).strip()

            if not cleaned_content:
                skipped_pages += 1
                self._debug_log(f"Skipping empty page {page_index}")
                continue

            # Create document with metadata
            documents.append(
                Document(
                    page_content=cleaned_content,
                    metadata={
                        "page": page_index,
                        "page_label": page_index + 1,
                        "total_pages": total_pages,
                        "file_name": self.file_name,
                        "file_size": self.file_size,
                        "processing_engine": "azure-mistral-ocr",
                        "model_name": self.model_name,
                        "endpoint_url": self.endpoint_url,
                        "content_length": len(cleaned_content),
                    },
                )
            )

        if skipped_pages > 0:
            log.info(
                f"Processed {len(documents)} pages, skipped {skipped_pages} empty/invalid pages"
            )

        if not documents:
            log.warning("OCR response contained pages, but none had valid content.")
            return [
                Document(
                    page_content="No valid text content found in document",
                    metadata={
                        "error": "no_valid_pages",
                        "total_pages": total_pages,
                        "file_name": self.file_name,
                    },
                )
            ]

        return documents

    def load(self) -> List[Document]:
        """
        Executes the Azure Mistral OCR workflow: encode file, process OCR.
        Synchronous version for backward compatibility.

        Returns:
            A list of Document objects, one for each page processed.
        """
        start_time = time.time()

        try:
            # 1. Encode file to base64
            base64_content = self._encode_file_to_base64()

            # 2. Process OCR
            ocr_response = self._process_ocr(base64_content)

            # 3. Process results
            documents = self._process_results(ocr_response)

            total_time = time.time() - start_time
            log.info(
                f"Azure Mistral OCR workflow completed in {total_time:.2f}s, produced {len(documents)} documents"
            )

            return documents

        except Exception as e:
            total_time = time.time() - start_time
            log.error(
                f"An error occurred during the loading process after {total_time:.2f}s: {e}"
            )
            return [
                Document(
                    page_content=f"Error during processing: {e}",
                    metadata={
                        "error": "processing_failed",
                        "file_name": self.file_name,
                    },
                )
            ]

    async def load_async(self) -> List[Document]:
        """
        Asynchronous Azure Mistral OCR workflow execution.

        Returns:
            A list of Document objects, one for each page processed.
        """
        start_time = time.time()

        try:
            # 1. Encode file to base64
            base64_content = self._encode_file_to_base64()

            # 2. Process OCR
            async with self._get_session() as session:
                ocr_response = await self._process_ocr_async(session, base64_content)

            # 3. Process results
            documents = self._process_results(ocr_response)

            total_time = time.time() - start_time
            log.info(
                f"Async Azure Mistral OCR workflow completed in {total_time:.2f}s, produced {len(documents)} documents"
            )

            return documents

        except Exception as e:
            total_time = time.time() - start_time
            log.error(f"Async Azure Mistral OCR workflow failed after {total_time:.2f}s: {e}")
            return [
                Document(
                    page_content=f"Error during OCR processing: {e}",
                    metadata={
                        "error": "processing_failed",
                        "file_name": self.file_name,
                    },
                )
            ]

    @staticmethod
    async def load_multiple_async(
        loaders: List["AzureMistralOCRLoader"],
        max_concurrent: int = 5,
    ) -> List[List[Document]]:
        """
        Process multiple files concurrently with controlled concurrency.

        Args:
            loaders: List of AzureMistralOCRLoader instances
            max_concurrent: Maximum number of concurrent requests

        Returns:
            List of document lists, one for each loader
        """
        if not loaders:
            return []

        log.info(
            f"Starting concurrent processing of {len(loaders)} files with max {max_concurrent} concurrent"
        )
        start_time = time.time()

        # Use semaphore to control concurrency
        semaphore = asyncio.Semaphore(max_concurrent)

        async def process_with_semaphore(loader: "AzureMistralOCRLoader") -> List[Document]:
            async with semaphore:
                return await loader.load_async()

        # Process all files with controlled concurrency
        tasks = [process_with_semaphore(loader) for loader in loaders]
        results = await asyncio.gather(*tasks, return_exceptions=True)

        # Handle any exceptions in results
        processed_results = []
        for i, result in enumerate(results):
            if isinstance(result, Exception):
                log.error(f"File {i} failed: {result}")
                processed_results.append(
                    [
                        Document(
                            page_content=f"Error processing file: {result}",
                            metadata={
                                "error": "batch_processing_failed",
                                "file_index": i,
                            },
                        )
                    ]
                )
            else:
                processed_results.append(result)

        # Log comprehensive batch processing statistics
        total_time = time.time() - start_time
        total_docs = sum(len(docs) for docs in processed_results)
        success_count = sum(
            1 for result in results if not isinstance(result, Exception)
        )
        failure_count = len(results) - success_count

        log.info(
            f"Batch processing completed in {total_time:.2f}s: "
            f"{success_count} files succeeded, {failure_count} files failed, "
            f"produced {total_docs} total documents"
        )

        return processed_results
