import io
import os
import re
import json
import time
import copy
import math
import zipfile
from pathlib import Path

import streamlit as st
from PIL import Image, ImageDraw, ImageFont

import openpyxl
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter, range_boundaries
from openpyxl.cell.cell import MergedCell

from google import genai
from google.genai import types


# ============================================================
# STREAMLIT CONFIG
# ============================================================

st.set_page_config(
    page_title="Gemini AI - Dịch Excel Song Ngữ",
    page_icon="🌐",
    layout="wide",
)


# ============================================================
# CONSTANTS
# ============================================================

PRIMARY_MODEL = "gemini-2.5-flash"
FALLBACK_MODEL = "gemini-2.5-flash-lite"

SUPPORTED_EXTENSIONS = [
    "xlsx",
    "xlsm",
    "png",
    "jpg",
    "jpeg",
]

IMAGE_EXTENSIONS = {"png", "jpg", "jpeg"}
EXCEL_EXTENSIONS = {"xlsx", "xlsm"}

MAX_TEXT_BATCH = 40
MAX_RETRIES = 3


# ============================================================
# UI
# ============================================================

st.title("🌐 Dịch File Song Ngữ Trung ↔ Việt bằng Gemini AI")

st.caption(
    "Excel: giữ cấu trúc/định dạng tối đa và thêm bản dịch bên dưới. "
    "Ảnh: OCR bằng Gemini và thêm tiếng Việt bên dưới tiếng Trung."
)

st.markdown("---")


# ============================================================
# SIDEBAR
# ============================================================

st.sidebar.header("⚙️ Cấu hình")

direction = st.sidebar.selectbox(
    "Chọn hướng dịch:",
    [
        "Trung → Việt",
        "Việt → Trung",
    ],
)

st.sidebar.markdown("---")

st.sidebar.subheader("🤖 Gemini AI")

api_key_input = st.sidebar.text_input(
    "Gemini API Key",
    type="password",
    placeholder="AIza...",
    help="Có thể để trống nếu đã cấu hình GEMINI_API_KEY trong Streamlit Secrets.",
)

if direction == "Trung → Việt":
    st.sidebar.success(
        "Kết quả: tiếng Trung ở trên, tiếng Việt ở dưới."
    )
else:
    st.sidebar.success(
        "Kết quả cuối: tiếng Trung ở trên, tiếng Việt ở dưới."
    )


# ============================================================
# GET API KEY
# ============================================================

def get_api_key():
    """
    Ưu tiên:
    1. API key nhập trên sidebar
    2. st.secrets["GEMINI_API_KEY"]
    3. biến môi trường GEMINI_API_KEY
    """
    if api_key_input and api_key_input.strip():
        return api_key_input.strip()

    try:
        secret_key = st.secrets.get("GEMINI_API_KEY")
        if secret_key:
            return str(secret_key).strip()
    except Exception:
        pass

    env_key = os.getenv("GEMINI_API_KEY")

    if env_key:
        return env_key.strip()

    return ""


API_KEY = get_api_key()


# ============================================================
# GEMINI CLIENT
# ============================================================

@st.cache_resource(show_spinner=False)
def create_gemini_client(api_key):
    if not api_key:
        return None
    return genai.Client(api_key=api_key)


client = create_gemini_client(API_KEY)


# ============================================================
# HELPERS
# ============================================================

def clean_text(value):
    if value is None:
        return ""
    return str(value).strip()


def is_formula(value):
    if not isinstance(value, str):
        return False
    return value.startswith("=")


def is_number(value):
    if value is None:
        return False

    if isinstance(value, (int, float)):
        return True

    text = str(value).strip()

    if not text:
        return False

    patterns = [
        r"^-?\d+$",
        r"^-?\d+\.\d+$",
        r"^-?\d+,\d+$",
        r"^-?\d{1,3}(,\d{3})+(\.\d+)?$",
        r"^-?\d{1,3}(\.\d{3})+(,\d+)?$",
        r"^-?\d+(\.\d+)?%$",
    ]

    return any(re.match(p, text) for p in patterns)


def looks_like_code_or_id(text):
    """
    Không dịch các mã thuần kỹ thuật.
    """
    if not text:
        return False

    text = str(text).strip()

    if len(text) <= 2:
        return False

    if re.fullmatch(r"[A-Z]{2,10}[-_/]?\d{2,}", text):
        return True

    if re.fullmatch(r"[A-Z0-9_-]{5,}", text):
        return True

    return False


def should_translate_cell(value):
    """
    Chỉ dịch text thực sự.
    Không dịch: công thức, số, ô rỗng, mã kỹ thuật rõ ràng.
    """
    if value is None:
        return False

    if isinstance(value, bool):
        return False

    if isinstance(value, (int, float)):
        return False

    text = str(value).strip()

    if not text:
        return False

    if is_formula(text):
        return False

    if is_number(text):
        return False

    if looks_like_code_or_id(text):
        return False

    return True


# ============================================================
# GEMINI JSON CALL
# ============================================================

def extract_json(text):
    """
    Gemini đôi khi trả về markdown JSON. Hàm này cố lấy JSON ra an toàn.
    """
    if not text:
        raise ValueError("Gemini không trả về dữ liệu.")

    text = text.strip()

    try:
        return json.loads(text)
    except Exception:
        pass

    text = re.sub(r"^```json\s*", "", text, flags=re.IGNORECASE)
    text = re.sub(r"\s*
