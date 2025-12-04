import requests
import aiohttp
import asyncio
import logging
import os
import sys
import time
import base64
import json
from typing import List, Dict, Any, Optional, Union
from contextlib import asynccontextmanager
from io import BytesIO

from langchain_core.documents import Document
from open_webui.env import SRC_LOG_LEVELS, GLOBAL_LOG_LEVEL

logging.basicConfig(stream=sys.stdout, level=GLOBAL_LOG_LEVEL)
log = logging.getLogger(__name__)
log.setLevel(SRC_LOG_LEVELS["RAG"])


class AzureMistralOCRLoader:
    """
    Azure AI Foundry Mistral OCR loader for document processing.
    Loads documents by sending them to Azure AI Foundry Mistral OCR endpoint.
    Supports custom metadata extraction using JSON schemas for both bbox and document annotations.
    """

    def __init__(
        self,
        api_key: str,
        endpoint_url: str,
        model_name: str,
        file_path: str,
        include_image_base64: bool = True,
        document_annotation_format: Optional[Dict[str, Any]] = None,
        bbox_annotation_format: Optional[Dict[str, Any]] = None,
        custom_metadata_schema: Optional[Dict[str, Any]] = None,
    ):
        """
        Initialize the Azure Mistral OCR loader.

        Args:
            api_key: Azure AI Foundry API key
            endpoint_url: Azure AI Foundry endpoint URL
            model_name: Model name deployed in Azure AI Foundry
            file_path: Path to the document file
            include_image_base64: Whether to include base64 images in response
            document_annotation_format: Optional document annotation format schema
            bbox_annotation_format: Optional bbox annotation format schema
            custom_metadata_schema: Optional custom metadata extraction schema
        """
        self.api_key = api_key
        self.endpoint_url = endpoint_url
        self.model_name = model_name
        self.file_path = file_path
        self.include_image_base64 = include_image_base64
        self.document_annotation_format = document_annotation_format
        self.bbox_annotation_format = bbox_annotation_format
        
        # Use custom schema if provided, otherwise use default schemas
        if custom_metadata_schema:
            self.custom_metadata_schema = custom_metadata_schema
        else:
            # Default document annotation schema
            self.document_annotation_format = self.create_document_annotation_schema({
                "language": {"title": "Language", "type": "string"},
                "chapter_titles": {"title": "Chapter_Titles", "type": "string"},
                "urls": {"title": "urls", "type": "string"},
                "summary": {"title": "Summary", "type": "string"},
                "key_topics": {"title": "Key_Topics", "type": "array", "items": {"type": "string"}}
            })
            
            # Default bbox annotation schema
            self.bbox_annotation_format = self.create_bbox_annotation_schema({
                "image_type": {"title": "Image_Type", "type": "string"},
                "short_description": {"title": "Short_Description", "type": "string"},
                "summary": {"title": "Summary", "type": "string"}
            })

        # Validate required parameters
        if not self.api_key:
            raise ValueError("Azure Mistral OCR API key is required")
        if not self.endpoint_url:
            raise ValueError("Azure Mistral OCR endpoint URL is required")
        if not self.model_name:
            raise ValueError("Azure Mistral OCR model name is required")

        log.info(f"Initialized Azure Mistral OCR loader for {file_path}")
        log.info(f"Using document annotation schema: {list(self.document_annotation_format['json_schema']['schema']['properties'].keys())}")
        log.info(f"Using bbox annotation schema: {list(self.bbox_annotation_format['json_schema']['schema']['properties'].keys())}")

    def _encode_pdf_as_base64(self, pdf_stream: BytesIO) -> str:
        """Encode PDF content as base64 string"""
        try:
            pdf_stream.seek(0)
            pdf_bytes = pdf_stream.read()
            base64_encoded = base64.b64encode(pdf_bytes).decode("utf-8")
            return base64_encoded
        except Exception as e:
            log.error(f"Error encoding PDF as base64: {e}")
            raise

    def _create_payload(self, pdf_base64: str) -> Dict[str, Any]:
        """Create the payload for the Azure AI Foundry API"""
        payload = {
            "model": self.model_name,
            "document": {
                "type": "document_url",
                "document_url": f"data:application/pdf;base64,{pdf_base64}"
            },
            "include_image_base64": self.include_image_base64
        }

        # Add document annotation format if provided
        if self.document_annotation_format:
            payload["document_annotation_format"] = self.document_annotation_format

        # Add bbox annotation format if provided
        if self.bbox_annotation_format:
            payload["bbox_annotation_format"] = self.bbox_annotation_format

        return payload

    def _make_request(self, payload: Dict[str, Any]) -> Dict[str, Any]:
        """Make HTTP request to Azure AI Foundry endpoint"""
        headers = {
            "Content-Type": "application/json",
            "Authorization": f"Bearer {self.api_key}",
        }

        try:
            response = requests.post(
                self.endpoint_url,
                headers=headers,
                data=json.dumps(payload),
                timeout=300  # 5 minute timeout
            )
            response.raise_for_status()
            return response.json()
        except requests.exceptions.HTTPError as e:
            log.error(f"HTTP Error occurred: {e}")
            if hasattr(e, "response") and hasattr(e.response, "text"):
                log.error(f"Error details: {e.response.text}")
            raise
        except requests.exceptions.RequestException as e:
            log.error(f"Request error: {e}")
            raise
        except Exception as e:
            log.error(f"Unexpected error: {e}")
            raise

    def _extract_text_and_metadata(self, response: Dict[str, Any]) -> tuple[str, Dict[str, Any]]:
        """Extract text content and metadata from the API response"""
        try:
            extracted_text = ""
            extracted_metadata = {}

            # Try to extract text from various possible response structures
            if "text" in response:
                extracted_text = response["text"]
            elif "content" in response:
                extracted_text = response["content"]
            elif "result" in response:
                if isinstance(response["result"], str):
                    extracted_text = response["result"]
                elif isinstance(response["result"], dict) and "text" in response["result"]:
                    extracted_text = response["result"]["text"]
            elif "data" in response:
                if isinstance(response["data"], str):
                    extracted_text = response["data"]
                elif isinstance(response["data"], dict) and "text" in response["data"]:
                    extracted_text = response["data"]["text"]
            else:
                # If no text field found, return the entire response as JSON
                extracted_text = json.dumps(response, indent=2)

            # Extract document annotations
            if "document_annotation" in response:
                extracted_metadata["document_annotation"] = response["document_annotation"]
            elif "document_annotations" in response:
                extracted_metadata["document_annotation"] = response["document_annotations"]

            # Extract bbox annotations
            if "bbox_annotation" in response:
                extracted_metadata["bbox_annotation"] = response["bbox_annotation"]
            elif "bbox_annotations" in response:
                extracted_metadata["bbox_annotation"] = response["bbox_annotations"]

            # Extract other structured metadata if available
            if "metadata" in response:
                extracted_metadata.update(response["metadata"])
            elif "annotations" in response:
                extracted_metadata.update(response["annotations"])
            elif "structured_data" in response:
                extracted_metadata.update(response["structured_data"])
            elif "extracted_fields" in response:
                extracted_metadata.update(response["extracted_fields"])

            return extracted_text, extracted_metadata

        except Exception as e:
            log.error(f"Error extracting text and metadata from response: {e}")
            return json.dumps(response, indent=2), {}

    def load(self) -> List[Document]:
        """
        Load documents from the file using Azure Mistral OCR.

        Returns:
            List of Document objects containing the extracted text and metadata
        """
        try:
            log.info(f"Starting Azure Mistral OCR processing for {self.file_path}")

            # Read the PDF file
            with open(self.file_path, "rb") as f:
                pdf_bytes = BytesIO(f.read())

            # Encode PDF as base64
            log.info("Encoding PDF as base64...")
            base64_data = self._encode_pdf_as_base64(pdf_bytes)

            # Create payload
            payload = self._create_payload(base64_data)

            # Make API request
            log.info("Sending request to Azure AI Foundry...")
            response = self._make_request(payload)

            # Extract text and metadata from response
            extracted_text, extracted_metadata = self._extract_text_and_metadata(response)

            # Create document with enhanced metadata
            document_metadata = {
                "source": self.file_path,
                "loader": "azure_mistral_ocr",
                "model": self.model_name,
                "endpoint": self.endpoint_url,
                "include_image_base64": self.include_image_base64,
                "response_keys": list(response.keys()) if isinstance(response, dict) else []
            }

            # Add extracted metadata to document metadata
            if extracted_metadata:
                document_metadata.update(extracted_metadata)
                log.info(f"Extracted metadata fields: {list(extracted_metadata.keys())}")

            document = Document(
                page_content=extracted_text,
                metadata=document_metadata
            )

            log.info("Azure Mistral OCR processing completed successfully")
            return [document]

        except Exception as e:
            log.error(f"Error in Azure Mistral OCR processing: {e}")
            raise

    async def aload(self) -> List[Document]:
        """
        Async version of load method.

        Returns:
            List of Document objects containing the extracted text and metadata
        """
        try:
            log.info(f"Starting async Azure Mistral OCR processing for {self.file_path}")

            # Read the PDF file
            with open(self.file_path, "rb") as f:
                pdf_bytes = BytesIO(f.read())

            # Encode PDF as base64
            log.info("Encoding PDF as base64...")
            base64_data = self._encode_pdf_as_base64(pdf_bytes)

            # Create payload
            payload = self._create_payload(base64_data)

            # Make async API request
            log.info("Sending async request to Azure AI Foundry...")
            async with aiohttp.ClientSession() as session:
                headers = {
                    "Content-Type": "application/json",
                    "Authorization": f"Bearer {self.api_key}",
                }

                async with session.post(
                    self.endpoint_url,
                    headers=headers,
                    json=payload,
                    timeout=aiohttp.ClientTimeout(total=300)
                ) as response:
                    response.raise_for_status()
                    response_data = await response.json()

            # Extract text and metadata from response
            extracted_text, extracted_metadata = self._extract_text_and_metadata(response_data)

            # Create document with enhanced metadata
            document_metadata = {
                "source": self.file_path,
                "loader": "azure_mistral_ocr",
                "model": self.model_name,
                "endpoint": self.endpoint_url,
                "include_image_base64": self.include_image_base64,
                "response_keys": list(response_data.keys()) if isinstance(response_data, dict) else []
            }

            # Add extracted metadata to document metadata
            if extracted_metadata:
                document_metadata.update(extracted_metadata)
                log.info(f"Extracted metadata fields: {list(extracted_metadata.keys())}")

            document = Document(
                page_content=extracted_text,
                metadata=document_metadata
            )

            log.info("Async Azure Mistral OCR processing completed successfully")
            return [document]

        except Exception as e:
            log.error(f"Error in async Azure Mistral OCR processing: {e}")
            raise

    @staticmethod
    def create_document_annotation_schema(
        fields: Dict[str, Dict[str, Any]],
        schema_name: str = "document_annotation",
        strict: bool = True
    ) -> Dict[str, Any]:
        """
        Create a JSON schema for document annotation.
        
        Args:
            fields: Dictionary of field definitions with their types and descriptions
            schema_name: Name of the schema
            strict: Whether to use strict mode
            
        Returns:
            JSON schema for document annotation
        """
        return {
            "type": "json_schema",
            "json_schema": {
                "schema": {
                    "properties": fields,
                    "required": list(fields.keys()),
                    "title": "DocumentAnnotation",
                    "type": "object",
                    "additionalProperties": False
                },
                "name": schema_name,
                "strict": strict
            }
        }

    @staticmethod
    def create_bbox_annotation_schema(
        fields: Dict[str, Dict[str, Any]],
        schema_name: str = "bbox_annotation",
        strict: bool = True
    ) -> Dict[str, Any]:
        """
        Create a JSON schema for bbox annotation.
        
        Args:
            fields: Dictionary of field definitions with their types and descriptions
            schema_name: Name of the schema
            strict: Whether to use strict mode
            
        Returns:
            JSON schema for bbox annotation
        """
        return {
            "type": "json_schema",
            "json_schema": {
                "schema": {
                    "properties": fields,
                    "required": list(fields.keys()),
                    "title": "BBOXAnnotation",
                    "type": "object",
                    "additionalProperties": False
                },
                "name": schema_name,
                "strict": strict
            }
        }