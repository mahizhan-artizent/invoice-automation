import os
from pathlib import Path
from dotenv import load_dotenv

# Load .env file from project root
BASE_DIR = Path(__file__).resolve().parent.parent
load_dotenv(BASE_DIR / ".env")

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY", "")
GEMINI_PRIMARY_MODEL = os.getenv("GEMINI_PRIMARY_MODEL", "gemini-3.7-flash")
GEMINI_FALLBACK_MODEL = os.getenv("GEMINI_FALLBACK_MODEL", "gemini-2.5-flash")
GEMINI_SECONDARY_FALLBACK = os.getenv("GEMINI_SECONDARY_FALLBACK", "gemini-2.5-flash-lite")

EXCEL_TEMPLATE_PATH = os.getenv("EXCEL_TEMPLATE_PATH", str(BASE_DIR / "Book3.xlsx"))
EXCEL_SHEET_NAME = os.getenv("EXCEL_SHEET_NAME", "AP BLR")
SQLITE_DB_PATH = os.getenv("SQLITE_DB_PATH", str(BASE_DIR / "invoice_automation.db"))
INVOICE_SAMPLES_DIR = os.getenv("INVOICE_SAMPLES_DIR", str(BASE_DIR / "samples inv"))
OUTPUT_DIR = str(BASE_DIR / "output")

FASTAPI_PORT = int(os.getenv("FASTAPI_PORT", 8000))
STREAMLIT_PORT = int(os.getenv("STREAMLIT_PORT", 8501))

# Microsoft Graph API Configuration (for Outlook mailbox integration)
GRAPH_CLIENT_ID = os.getenv("GRAPH_CLIENT_ID", "")
GRAPH_CLIENT_SECRET = os.getenv("GRAPH_CLIENT_SECRET", "")
GRAPH_TENANT_ID = os.getenv("GRAPH_TENANT_ID", "")
GRAPH_USER_MAILBOX = os.getenv("GRAPH_USER_MAILBOX", "")

