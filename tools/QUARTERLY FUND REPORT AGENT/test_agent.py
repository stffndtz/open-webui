#!/usr/bin/env python3
"""
Quarterly Fund Report Agent - Local Test Script

This script allows testing the fund report analysis prompt locally
using OpenAI or Anthropic APIs with PDF extraction.

Usage:
    python test_agent.py --fund "HCF IV" --provider openai
    python test_agent.py --fund "HCF IV" --provider anthropic
    python test_agent.py --list-funds

API Keys:
    Set via environment variables or .env file:
    - OPENAI_API_KEY=sk-...
    - ANTHROPIC_API_KEY=sk-ant-...

    Or pass directly:
    python test_agent.py --fund "HCF IV" --api-key "sk-..."
"""

import os
import sys
import json
import argparse
import base64
from pathlib import Path
from datetime import datetime

# Load .env file if it exists
ENV_FILE = Path(__file__).parent / ".env"
if ENV_FILE.exists():
    with open(ENV_FILE) as f:
        for line in f:
            line = line.strip()
            if line and not line.startswith("#") and "=" in line:
                key, value = line.split("=", 1)
                os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))

# Optional: Install with pip install openai anthropic pymupdf
try:
    import fitz  # PyMuPDF for PDF text extraction
    HAS_PYMUPDF = True
except ImportError:
    HAS_PYMUPDF = False
    print("Warning: PyMuPDF not installed. Install with: pip install pymupdf")

try:
    from openai import OpenAI, AzureOpenAI
    HAS_OPENAI = True
except ImportError:
    HAS_OPENAI = False

try:
    import anthropic
    HAS_ANTHROPIC = True
except ImportError:
    HAS_ANTHROPIC = False

# Configuration
REPORTS_DIR = Path(__file__).parent / "reports"
PROMPT_FILE = Path(__file__).parent / "prompt.md"
OUTPUT_DIR = Path(__file__).parent / "outputs"

# Fund mappings (name -> file patterns)
FUND_MAPPINGS = {
    "HCF IV": {
        "q1": "2025-Q1 - NORD KB IV - HCF IV - Quarterly Report.pdf",
        "q2": "2025-Q2 - NORD KB IV - HCF IV - Quarterly Report.pdf",
    },
    "Henko II": {
        "q1": "2025-Q1 - MC VIII - Henko II - Quarterly Report.pdf",
        "q2": "2025-Q2 - MC VIII - Henko II - Quarterly Report.pdf",
    },
    "Aliter II": {
        "q1": "2025-Q1 - MC VIII_NH138 - Aliter II - Quarterly Report.pdf",
        "q2": "2025-Q2 - MC VIII_NH138 - Aliter II - Quarterly Report.pdf",
    },
    "Bragnum I": {
        "q1": "2025-Q1 - NORD KB IV - Bragnum I - Quarterly Report.pdf",
        "q2": "2025-Q2 - NORD KB IV - Bragnum I - Quarterly Report.pdf",
    },
    "Motive I": {
        "q1": "2025-Q1 - NORD KB IV - Motive I - Investor Update.pdf",
        "q2": "2025-Q2 - NORD KB IV - Motive I - Investor Update.pdf",
    },
    "Progressio III": {
        "q1": "2025-Q1 - NORD KB IV - Progressio III - Quarterly Report.pdf",
        "q2": "2025-Q2 - NORD KB IV - Progressio III - Quarterly Report.pdf",
    },
    "Palero III": {
        "q1": "2025-Q1 - NORD KB MC V - Palero III - Quarterly Report.pdf",
        "q2": "2025-Q2 - NORD KB MC V - Palero III - Quarterly Report.pdf",
    },
    "FA R Evolution": {
        "q1": "2025-Q1 - MC VIII - FA R Evolution - Report.pdf",
        "q2": "2025-Q2 - MC VIII - FA R Evolution - Report.pdf",
    },
}


def extract_text_from_pdf(pdf_path: Path) -> str:
    """Extract text from PDF using PyMuPDF."""
    if not HAS_PYMUPDF:
        raise ImportError("PyMuPDF required. Install with: pip install pymupdf")

    doc = fitz.open(pdf_path)
    text_parts = []

    for page_num, page in enumerate(doc, 1):
        text = page.get_text()
        text_parts.append(f"\n--- Page {page_num} ---\n{text}")

    doc.close()
    return "\n".join(text_parts)


def extract_images_from_pdf(pdf_path: Path, max_pages: int = 5) -> list[dict]:
    """Extract page images from PDF for vision models."""
    if not HAS_PYMUPDF:
        raise ImportError("PyMuPDF required. Install with: pip install pymupdf")

    doc = fitz.open(pdf_path)
    images = []

    for page_num in range(min(len(doc), max_pages)):
        page = doc[page_num]
        # Render page as image
        mat = fitz.Matrix(2, 2)  # 2x zoom for better quality
        pix = page.get_pixmap(matrix=mat)
        img_data = pix.tobytes("png")
        b64_data = base64.b64encode(img_data).decode("utf-8")
        images.append({
            "page": page_num + 1,
            "base64": b64_data,
            "media_type": "image/png"
        })

    doc.close()
    return images


def load_system_prompt(prompt_file: Path = PROMPT_FILE) -> str:
    """Load the system prompt from file and inject current date."""
    with open(prompt_file, "r") as f:
        prompt = f.read()
    # Inject current date
    current_date = datetime.now().strftime("%Y-%m-%d")
    prompt = prompt.replace("{{CURRENT_DATE}}", current_date)
    return prompt


def create_user_message(fund_name: str, q1_text: str, q2_text: str) -> str:
    """Create the user message with both reports."""
    return f"""Analyze these two consecutive quarterly fund reports and generate the standardized performance analysis following NORD's specification.

t=0 = Q1/2025, t=1 = Q2/2025; Fund: {fund_name}

## REPORT 1 (t=0 - Q1 2025):
{q1_text}

## REPORT 2 (t=1 - Q2 2025):
{q2_text}
"""


def get_openai_client(api_key: str = None):
    """Get OpenAI or Azure OpenAI client based on environment."""
    if not HAS_OPENAI:
        raise ImportError("OpenAI library required. Install with: pip install openai")

    # Check for Azure OpenAI configuration
    azure_endpoint = os.environ.get("AZURE_OPENAI_ENDPOINT")
    azure_key = api_key or os.environ.get("AZURE_OPENAI_API_KEY")

    if azure_endpoint and azure_key:
        return AzureOpenAI(
            azure_endpoint=azure_endpoint,
            api_key=azure_key,
            api_version=os.environ.get("AZURE_OPENAI_API_VERSION", "2024-08-01-preview")
        ), True  # is_azure=True
    elif api_key:
        return OpenAI(api_key=api_key), False
    else:
        return OpenAI(), False


def run_openai(system_prompt: str, user_message: str, model: str = "gpt-4o", api_key: str = None) -> str:
    """Run analysis using OpenAI API (supports Azure OpenAI)."""
    client, is_azure = get_openai_client(api_key)

    # For Azure, model is the deployment name
    deployment = os.environ.get("AZURE_OPENAI_DEPLOYMENT", model) if is_azure else model

    response = client.chat.completions.create(
        model=deployment,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": user_message}
        ],
        temperature=0.2,
        max_tokens=8000,
    )

    return response.choices[0].message.content


def run_openai_vision(
    system_prompt: str,
    fund_name: str,
    q1_images: list[dict],
    q2_images: list[dict],
    model: str = "gpt-4o",
    api_key: str = None
) -> str:
    """Run analysis using OpenAI Vision API with page images (supports Azure OpenAI)."""
    client, is_azure = get_openai_client(api_key)

    # For Azure, model is the deployment name
    deployment = os.environ.get("AZURE_OPENAI_DEPLOYMENT", model) if is_azure else model

    # Build content with images
    content = [
        {
            "type": "text",
            "text": f"Analyze these two consecutive quarterly fund reports for {fund_name} (Q1 2025 = t=0, Q2 2025 = t=1). First I'll show Q1 pages, then Q2 pages."
        }
    ]

    # Add Q1 images
    content.append({"type": "text", "text": "\n--- Q1 2025 REPORT (t=0) ---"})
    for img in q1_images:
        content.append({
            "type": "image_url",
            "image_url": {
                "url": f"data:{img['media_type']};base64,{img['base64']}",
                "detail": "high"
            }
        })

    # Add Q2 images
    content.append({"type": "text", "text": "\n--- Q2 2025 REPORT (t=1) ---"})
    for img in q2_images:
        content.append({
            "type": "image_url",
            "image_url": {
                "url": f"data:{img['media_type']};base64,{img['base64']}",
                "detail": "high"
            }
        })

    response = client.chat.completions.create(
        model=deployment,
        messages=[
            {"role": "system", "content": system_prompt},
            {"role": "user", "content": content}
        ],
        temperature=0.2,
        max_tokens=8000,
    )

    return response.choices[0].message.content


def run_anthropic(system_prompt: str, user_message: str, model: str = "claude-sonnet-4-20250514", api_key: str = None) -> str:
    """Run analysis using Anthropic API."""
    if not HAS_ANTHROPIC:
        raise ImportError("Anthropic library required. Install with: pip install anthropic")

    client = anthropic.Anthropic(api_key=api_key) if api_key else anthropic.Anthropic()

    response = client.messages.create(
        model=model,
        max_tokens=8000,
        system=system_prompt,
        messages=[
            {"role": "user", "content": user_message}
        ]
    )

    return response.content[0].text


def run_anthropic_vision(
    system_prompt: str,
    fund_name: str,
    q1_images: list[dict],
    q2_images: list[dict],
    model: str = "claude-sonnet-4-20250514"
) -> str:
    """Run analysis using Anthropic Vision API with page images."""
    if not HAS_ANTHROPIC:
        raise ImportError("Anthropic library required. Install with: pip install anthropic")

    client = anthropic.Anthropic()

    # Build content with images
    content = [
        {
            "type": "text",
            "text": f"Analyze these two consecutive quarterly fund reports for {fund_name} (Q1 2025 = t=0, Q2 2025 = t=1). First I'll show Q1 pages, then Q2 pages."
        }
    ]

    # Add Q1 images
    content.append({"type": "text", "text": "\n--- Q1 2025 REPORT (t=0) ---"})
    for img in q1_images:
        content.append({
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": img["media_type"],
                "data": img["base64"]
            }
        })

    # Add Q2 images
    content.append({"type": "text", "text": "\n--- Q2 2025 REPORT (t=1) ---"})
    for img in q2_images:
        content.append({
            "type": "image",
            "source": {
                "type": "base64",
                "media_type": img["media_type"],
                "data": img["base64"]
            }
        })

    response = client.messages.create(
        model=model,
        max_tokens=8000,
        system=system_prompt,
        messages=[
            {"role": "user", "content": content}
        ]
    )

    return response.content[0].text


def save_output(fund_name: str, provider: str, content: str, mode: str = "text"):
    """Save the analysis output to a file."""
    OUTPUT_DIR.mkdir(exist_ok=True)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    safe_name = fund_name.replace(" ", "_")
    filename = f"{safe_name}_{provider}_{mode}_{timestamp}.md"
    output_path = OUTPUT_DIR / filename

    with open(output_path, "w") as f:
        f.write(f"# Fund Analysis: {fund_name}\n")
        f.write(f"**Provider**: {provider}\n")
        f.write(f"**Mode**: {mode}\n")
        f.write(f"**Generated**: {datetime.now().isoformat()}\n\n")
        f.write("---\n\n")
        f.write(content)

    return output_path


def list_funds():
    """List available funds for testing."""
    print("\nAvailable funds for testing:")
    print("-" * 40)
    for fund_name, files in FUND_MAPPINGS.items():
        q1_exists = (REPORTS_DIR / files["q1"]).exists()
        q2_exists = (REPORTS_DIR / files["q2"]).exists()
        status = "OK" if (q1_exists and q2_exists) else "MISSING"
        print(f"  {fund_name}: [{status}]")
        if not q1_exists:
            print(f"    - Missing Q1: {files['q1']}")
        if not q2_exists:
            print(f"    - Missing Q2: {files['q2']}")
    print()


def main():
    parser = argparse.ArgumentParser(description="Test the Quarterly Fund Report Agent")
    parser.add_argument("--fund", type=str, help="Fund name to analyze")
    parser.add_argument("--provider", type=str, choices=["openai", "anthropic"], default="openai",
                        help="API provider to use")
    parser.add_argument("--model", type=str, help="Model to use (default: gpt-4o or claude-sonnet-4-20250514)")
    parser.add_argument("--mode", type=str, choices=["text", "vision"], default="text",
                        help="Analysis mode: text (OCR) or vision (images)")
    parser.add_argument("--max-pages", type=int, default=10,
                        help="Max pages to include in vision mode")
    parser.add_argument("--list-funds", action="store_true", help="List available funds")
    parser.add_argument("--prompt", type=str, help="Custom prompt file path")
    parser.add_argument("--dry-run", action="store_true", help="Extract text only, don't call API")

    args = parser.parse_args()

    if args.list_funds:
        list_funds()
        return

    if not args.fund:
        parser.print_help()
        print("\nError: --fund is required. Use --list-funds to see available funds.")
        sys.exit(1)

    if args.fund not in FUND_MAPPINGS:
        print(f"Error: Unknown fund '{args.fund}'. Use --list-funds to see available funds.")
        sys.exit(1)

    # Load files
    fund_files = FUND_MAPPINGS[args.fund]
    q1_path = REPORTS_DIR / fund_files["q1"]
    q2_path = REPORTS_DIR / fund_files["q2"]

    if not q1_path.exists() or not q2_path.exists():
        print(f"Error: Report files not found for {args.fund}")
        sys.exit(1)

    # Load prompt
    prompt_file = Path(args.prompt) if args.prompt else PROMPT_FILE
    system_prompt = load_system_prompt(prompt_file)

    print(f"\nAnalyzing fund: {args.fund}")
    print(f"Provider: {args.provider}")
    print(f"Mode: {args.mode}")
    print(f"Q1 Report: {q1_path.name}")
    print(f"Q2 Report: {q2_path.name}")
    print()

    if args.mode == "text":
        # Extract text from PDFs
        print("Extracting text from Q1 report...")
        q1_text = extract_text_from_pdf(q1_path)
        print(f"  Extracted {len(q1_text):,} characters")

        print("Extracting text from Q2 report...")
        q2_text = extract_text_from_pdf(q2_path)
        print(f"  Extracted {len(q2_text):,} characters")

        if args.dry_run:
            print("\n--- Q1 TEXT (first 2000 chars) ---")
            print(q1_text[:2000])
            print("\n--- Q2 TEXT (first 2000 chars) ---")
            print(q2_text[:2000])
            return

        # Create user message
        user_message = create_user_message(args.fund, q1_text, q2_text)

        # Run analysis
        print("\nRunning analysis...")
        model = args.model
        if args.provider == "openai":
            model = model or "gpt-4o"
            result = run_openai(system_prompt, user_message, model)
        else:
            model = model or "claude-sonnet-4-20250514"
            result = run_anthropic(system_prompt, user_message, model)

    else:  # vision mode
        print(f"Extracting page images from Q1 report (max {args.max_pages} pages)...")
        q1_images = extract_images_from_pdf(q1_path, args.max_pages)
        print(f"  Extracted {len(q1_images)} pages")

        print(f"Extracting page images from Q2 report (max {args.max_pages} pages)...")
        q2_images = extract_images_from_pdf(q2_path, args.max_pages)
        print(f"  Extracted {len(q2_images)} pages")

        if args.dry_run:
            print(f"\nWould send {len(q1_images) + len(q2_images)} images to {args.provider}")
            return

        # Run vision analysis
        print("\nRunning vision analysis...")
        model = args.model
        if args.provider == "openai":
            model = model or "gpt-4o"
            result = run_openai_vision(system_prompt, args.fund, q1_images, q2_images, model)
        else:
            model = model or "claude-sonnet-4-20250514"
            result = run_anthropic_vision(system_prompt, args.fund, q1_images, q2_images, model)

    # Save and display result
    output_path = save_output(args.fund, args.provider, result, args.mode)
    print(f"\nOutput saved to: {output_path}")
    print("\n" + "=" * 60)
    print("ANALYSIS RESULT:")
    print("=" * 60 + "\n")
    print(result)


if __name__ == "__main__":
    main()
