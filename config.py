"""
config.py
---------
Central configuration for API keys and default settings.
You can place your GEMINI_API_KEY or ANTHROPIC_API_KEY here or in a .env file,
and the system will automatically use it without asking for it in the frontend.
"""
import os
from pathlib import Path

try:
    from dotenv import load_dotenv
    env_path = Path(__file__).parent / ".env"
    load_dotenv(dotenv_path=env_path)
except ImportError:
    pass

# ==============================================================================
# 🔑 PASTE YOUR API KEYS HERE
# ==============================================================================

# Google Gemini API Key (Must start with "AIzaSy...")
# Get a free key in 30 seconds from: https://aistudio.google.com/apikey
GEMINI_API_KEY = ""

# Anthropic Claude API Key (Must start with "sk-ant-...")
# Get from: https://console.anthropic.com
ANTHROPIC_API_KEY = ""

# ==============================================================================
# Helper functions to retrieve and validate keys
# ==============================================================================

def get_gemini_key():
    # 1. Direct key in config.py
    if GEMINI_API_KEY and GEMINI_API_KEY.strip() and not GEMINI_API_KEY.startswith("AQ."):
        return GEMINI_API_KEY.strip()
    # 2. Environment variable or .env
    env_key = os.environ.get("GEMINI_API_KEY", "").strip()
    if env_key and not env_key.startswith("AQ."):
        return env_key
    return ""


def get_anthropic_key():
    # 1. Direct key in config.py
    if ANTHROPIC_API_KEY and ANTHROPIC_API_KEY.strip():
        return ANTHROPIC_API_KEY.strip()
    # 2. Environment variable or .env
    env_key = os.environ.get("ANTHROPIC_API_KEY", "").strip()
    if env_key:
        return env_key
    return ""


# Default Vision Models
DEFAULT_GEMINI_MODEL = "gemini-3.6-flash"
DEFAULT_CLAUDE_MODEL = "claude-3-7-sonnet-latest"
FALLBACK_GEMINI_MODEL = "gemini-3.5-flash-lite"

# Default Processing Settings
DEFAULT_PAGES_PER_PARTICIPANT = 2
DEFAULT_SHEET_NAME = "Mod-1 Participant's Feedback"
