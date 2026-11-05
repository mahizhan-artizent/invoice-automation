import os
import time
import json
import logging
from pathlib import Path
from typing import Dict, Any, Optional, Tuple, Union
from google import genai
from google.genai import types
from google.genai.errors import ServerError, ClientError

from core.config import (
    GEMINI_API_KEY,
    GEMINI_PRIMARY_MODEL,
    GEMINI_FALLBACK_MODEL,
    GEMINI_SECONDARY_FALLBACK
)

logger = logging.getLogger(__name__)

class GeminiDocumentService:
    """
    Dedicated Gemini document intelligence service using the official google-genai SDK.
    Supports multimodal input (PDFs, images), dynamic schema enforcement,
    and automatic failover if high-tier models hit demand spikes (503).
    """

    def __init__(self, api_key: Optional[str] = None, client: Optional[Any] = None):
        self.api_key = api_key or GEMINI_API_KEY
        if client:
            self.client = client
        else:
            if not self.api_key:
                raise ValueError("GEMINI_API_KEY is not set. Please provide it in .env or constructor.")
            self.client = genai.Client(api_key=self.api_key)

    def extract_invoice_data(
        self,
        document_path: Union[str, Path],
        prompt: str,
        folder_context: str = "",
        model_name: Optional[str] = None
    ) -> Tuple[Dict[str, Any], str, float]:
        """
        Processes a document multimodal payload and returns structured extraction JSON.
        
        Returns:
            (extracted_dict, model_used, execution_seconds)
        """
        doc_path = Path(document_path)
        if not doc_path.exists():
            raise FileNotFoundError(f"Invoice file not found: {doc_path}")

        suffix = doc_path.suffix.lower()
        mime_types = {
            ".pdf": "application/pdf",
            ".png": "image/png",
            ".jpg": "image/jpeg",
            ".jpeg": "image/jpeg",
            ".tiff": "image/tiff",
            ".webp": "image/webp"
        }
        mime_type = mime_types.get(suffix, "application/pdf")

        with open(doc_path, "rb") as f:
            doc_bytes = f.read()

        primary = model_name or GEMINI_PRIMARY_MODEL
        models_to_try = [
            primary,
            GEMINI_FALLBACK_MODEL,
            GEMINI_SECONDARY_FALLBACK
        ]
        # Remove duplicate model names while preserving order
        seen = set()
        model_cascade = []
        for m in models_to_try:
            if m and m not in seen:
                seen.add(m)
                model_cascade.append(m)

        last_error = None
        start_time = time.time()

        for current_model in model_cascade:
            for attempt in range(2):
                try:
                    logger.info(f"Extracting with model '{current_model}' (attempt {attempt + 1})...")
                    
                    doc_part = types.Part.from_bytes(data=doc_bytes, mime_type=mime_type)
                    
                    response = self.client.models.generate_content(
                        model=current_model,
                        contents=[
                            doc_part,
                            prompt
                        ],
                        config=types.GenerateContentConfig(
                            response_mime_type="application/json",
                            temperature=0.1
                        )
                    )

                    elapsed = time.time() - start_time
                    raw_text = response.text.strip()
                    
                    # Clean markdown code blocks if present
                    if raw_text.startswith("```json"):
                        raw_text = raw_text[7:]
                    elif raw_text.startswith("```"):
                        raw_text = raw_text[3:]
                    if raw_text.endswith("```"):
                        raw_text = raw_text[:-3]
                    raw_text = raw_text.strip()

                    parsed_json = json.loads(raw_text)
                    return parsed_json, current_model, round(elapsed, 2)

                except Exception as e:
                    last_error = e
                    err_str = str(e).lower()
                    logger.warning(f"Model {current_model} failed (attempt {attempt + 1}): {e}")
                    if "503" in err_str or "unavailable" in err_str or "high demand" in err_str:
                        # Temporary high demand spike: sleep briefly then try next model
                        time.sleep(1.5)
                        break
                    elif "429" in err_str or "quota" in err_str:
                        time.sleep(2.0)
                        break
                    else:
                        # Non-recoverable or fatal error on this attempt
                        time.sleep(1.0)

        raise RuntimeError(f"All Gemini models in cascade failed to process {doc_path.name}. Last error: {last_error}")

