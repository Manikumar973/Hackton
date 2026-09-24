import os
import re
import urllib.parse
import smtplib
from email.message import EmailMessage
from io import BytesIO
from datetime import date, datetime, timedelta

import streamlit as st
import pandas as pd
import numpy as np
import cv2
import pytesseract

from PIL import Image
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment


# ============================================================
# OPTIONAL QR LIBRARY
# ============================================================

try:
    from pyzbar.pyzbar import decode as pyzbar_decode
    PYZBAR_AVAILABLE = True
except Exception:
    PYZBAR_AVAILABLE = False


# ============================================================
# PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Hackathon Image to Excel Extractor",
    page_icon="🏆",
    layout="wide"
)


# ============================================================
# PROFESSIONAL UI CSS
# ============================================================

st.markdown(
    """
    <style>

    .stApp {
        background: #f6f8fb;
    }

    .block-container {
        max-width: 1400px;
        padding-top: 1.5rem;
        padding-bottom: 3rem;
    }

    .hero {
        background: linear-gradient(
            135deg,
            #111827 0%,
            #1e3a8a 55%,
            #2563eb 100%
        );
        padding: 34px 38px;
        border-radius: 22px;
        margin-bottom: 25px;
        box-shadow: 0 12px 30px rgba(15, 23, 42, 0.18);
        color: white;
    }

    .hero-title {
        font-size: 38px;
        font-weight: 800;
        margin-bottom: 8px;
        letter-spacing: -0.7px;
    }

    .hero-subtitle {
        font-size: 16px;
        color: #dbeafe;
        margin-bottom: 0;
        line-height: 1.6;
    }

    .hero-badge {
        display: inline-block;
        background: rgba(255,255,255,0.14);
        border: 1px solid rgba(255,255,255,0.22);
        padding: 7px 14px;
        border-radius: 999px;
        font-size: 13px;
        margin-bottom: 15px;
    }

    .section-title {
        font-size: 24px;
        font-weight: 800;
        color: #111827;
        margin-top: 22px;
        margin-bottom: 6px;
    }

    .section-subtitle {
        color: #64748b;
        font-size: 14px;
        margin-bottom: 18px;
    }

    .metric-card {
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 16px;
        padding: 20px;
        box-shadow: 0 4px 14px rgba(15, 23, 42, 0.06);
        min-height: 125px;
    }

    .metric-icon {
        font-size: 26px;
        margin-bottom: 8px;
    }

    .metric-label {
        color: #64748b;
        font-size: 13px;
        font-weight: 600;
    }

    .metric-value {
        color: #111827;
        font-size: 29px;
        font-weight: 800;
        margin-top: 3px;
    }

    .info-card {
        background: white;
        border: 1px solid #e5e7eb;
        border-radius: 16px;
        padding: 20px;
        margin: 12px 0;
        box-shadow: 0 4px 14px rgba(15, 23, 42, 0.05);
    }

    .alert-card {
        background: #fff7ed;
        border: 1px solid #fed7aa;
        border-left: 5px solid #f97316;
        border-radius: 14px;
        padding: 18px;
        margin: 12px 0;
    }

    .success-card {
        background: #f0fdf4;
        border: 1px solid #bbf7d0;
        border-left: 5px solid #22c55e;
        border-radius: 14px;
        padding: 18px;
        margin: 12px 0;
    }

    .danger-card {
        background: #fef2f2;
        border: 1px solid #fecaca;
        border-left: 5px solid #ef4444;
        border-radius: 14px;
        padding: 18px;
        margin: 12px 0;
    }

    .neutral-card {
        background: #eff6ff;
        border: 1px solid #bfdbfe;
        border-left: 5px solid #2563eb;
        border-radius: 14px;
        padding: 18px;
        margin: 12px 0;
    }

    .stButton > button {
        border-radius: 10px;
        font-weight: 700;
        min-height: 42px;
    }

    .stDownloadButton > button {
        border-radius: 10px;
        font-weight: 700;
        min-height: 42px;
    }

    [data-testid="stFileUploader"] {
        background: white;
        border: 2px dashed #93c5fd;
        border-radius: 16px;
        padding: 10px;
    }

    [data-testid="stDataEditor"] {
        border-radius: 14px;
        overflow: hidden;
        border: 1px solid #e5e7eb;
    }

    .streamlit-expanderHeader {
        font-weight: 700;
    }

    [data-testid="stSidebar"] {
        background: #f8fafc;
        border-right: 1px solid #e5e7eb;
    }

    .footer {
        text-align: center;
        color: #94a3b8;
        font-size: 13px;
        padding: 28px 0 5px 0;
    }

    </style>
    """,
    unsafe_allow_html=True
)


# ============================================================
# TESSERACT CONFIGURATION
# ============================================================

TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

if os.path.exists(TESSERACT_PATH):
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH


# ============================================================
# HERO
# ============================================================

st.markdown(
    """
    <div class="hero">
        <div class="hero-badge">🏆 OCR • QR • Excel • Deadline Alerts</div>
        <div class="hero-title">
            Hackathon Image-to-Excel Extractor
        </div>
        <div class="hero-subtitle">
            Upload hackathon posters and automatically extract
            hackathon names, themes, deadlines, application links
            and QR codes into an editable Excel-ready table.
        </div>
    </div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# SESSION STATE
# ============================================================

if "email_last_success" not in st.session_state:
    st.session_state.email_last_success = False

if "email_last_message" not in st.session_state:
    st.session_state.email_last_message = ""

if "email_last_time" not in st.session_state:
    st.session_state.email_last_time = ""

if "extracted_hackathon_df" not in st.session_state:
    st.session_state.extracted_hackathon_df = pd.DataFrame()


# ============================================================
# SIDEBAR - EMAIL CONFIGURATION
# ============================================================

with st.sidebar:

    st.markdown("## 📧 Email Alerts")

    st.caption(
        "Configure Gmail SMTP to receive hackathon deadline alerts."
    )

    email_sender = st.text_input(
        "Sender Gmail",
        value=os.getenv("EMAIL_SENDER", ""),
        placeholder="yourgmail@gmail.com"
    )

    email_app_password = st.text_input(
        "Gmail App Password",
        value=os.getenv("EMAIL_APP_PASSWORD", ""),
        type="password",
        placeholder="Gmail App Password"
    )

    email_recipient = st.text_input(
        "Recipient Email",
        value=os.getenv("EMAIL_RECIPIENT", ""),
        placeholder="recipient@example.com"
    )

    st.markdown("---")

    st.markdown("### 🔧 System Status")

    if os.path.exists(TESSERACT_PATH):
        st.success("✅ Tesseract OCR Ready")
    else:
        st.warning("⚠️ Tesseract path not found")

    if PYZBAR_AVAILABLE:
        st.success("✅ Pyzbar QR Ready")
    else:
        st.info("ℹ️ OpenCV QR detection active")

    st.markdown("---")

    st.caption(
        "Email credentials are used only for sending alerts."
    )


# ============================================================
# IMAGE → OPENCV
# ============================================================

def pil_to_cv(image):
    return cv2.cvtColor(
        np.array(image.convert("RGB")),
        cv2.COLOR_RGB2BGR
    )


# ============================================================
# OCR
# ============================================================

def run_ocr(image):

    cv_img = pil_to_cv(image)

    results = []

    try:

        # Original
        original = cv_img.copy()

        # Resize
        enlarged = cv2.resize(
            original,
            None,
            fx=2,
            fy=2,
            interpolation=cv2.INTER_CUBIC
        )

        # Grayscale
        gray = cv2.cvtColor(
            enlarged,
            cv2.COLOR_BGR2GRAY
        )

        # OTSU
        _, otsu = cv2.threshold(
            gray,
            0,
            255,
            cv2.THRESH_BINARY + cv2.THRESH_OTSU
        )

        # Adaptive threshold
        adaptive = cv2.adaptiveThreshold(
            gray,
            255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY,
            31,
            11
        )

        images = [
            original,
            enlarged,
            gray,
            otsu,
            adaptive
        ]

        for img in images:

            try:
                text = pytesseract.image_to_string(
                    img,
                    config="--psm 6"
                )

                if text:
                    results.append(text)

            except Exception:
                continue

    except Exception:
        return ""

    unique_lines = []

    seen = set()

    for text in results:

        for line in text.splitlines():

            line = line.strip()

            if not line:
                continue

            normalized = re.sub(
                r"\s+",
                " ",
                line
            ).strip()

            key = normalized.lower()

            if key not in seen:

                seen.add(key)

                unique_lines.append(
                    normalized
                )

    return "\n".join(unique_lines)


# ============================================================
# QR CODE EXTRACTION
# ============================================================

def extract_qr_codes(image):

    cv_img = pil_to_cv(image)

    decoded_values = []

    detector = cv2.QRCodeDetector()

    try:

        images = [
            cv_img
        ]

        enlarged_2 = cv2.resize(
            cv_img,
            None,
            fx=2,
            fy=2,
            interpolation=cv2.INTER_CUBIC
        )

        enlarged_3 = cv2.resize(
            cv_img,
            None,
            fx=3,
            fy=3,
            interpolation=cv2.INTER_CUBIC
        )

        gray = cv2.cvtColor(
            enlarged_2,
            cv2.COLOR_BGR2GRAY
        )

        _, otsu = cv2.threshold(
            gray,
            0,
            255,
            cv2.THRESH_BINARY + cv2.THRESH_OTSU
        )

        adaptive = cv2.adaptiveThreshold(
            gray,
            255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY,
            31,
            11
        )

        images.extend(
            [
                enlarged_2,
                enlarged_3,
                gray,
                otsu,
                adaptive
            ]
        )

        for img in images:

            try:

                data, points, _ = detector.detectAndDecode(img)

                if data:
                    decoded_values.append(data)

            except Exception:
                pass

            try:

                result = detector.detectAndDecodeMulti(img)

                if len(result) >= 4:

                    retval, decoded_info, _, _ = result

                    if retval and decoded_info:

                        for value in decoded_info:

                            if value:
                                decoded_values.append(value)

            except Exception:
                pass

    except Exception:
        pass

    # Pyzbar
    if PYZBAR_AVAILABLE:

        try:

            for img in [
                cv_img,
                enlarged_2,
                enlarged_3
            ]:

                try:

                    decoded = pyzbar_decode(img)

                    for item in decoded:

                        try:
                            value = item.data.decode(
                                "utf-8",
                                errors="ignore"
                            )

                            if value:
                                decoded_values.append(value)

                        except Exception:
                            pass

                except Exception:
                    pass

        except Exception:
            pass

    # Remove duplicates
    final_values = []

    seen = set()

    for value in decoded_values:

        value = str(value).strip()

        if not value:
            continue

        key = value.lower()

        if key not in seen:

            seen.add(key)
            final_values.append(value)

    return final_values


# ============================================================
# URL FUNCTIONS
# ============================================================

def clean_url(url):

    if not url:
        return ""

    url = str(url).strip()

    url = url.strip(
        " \t\r\n.,;:()[]{}<>\"'"
    )

    if url.startswith("www."):
        url = "https://" + url

    return url


def is_valid_url(url):

    if not url:
        return False

    url = clean_url(url)

    if not url:
        return False

    lower = url.lower()

    if "example.com" in lower:
        return False

    try:

        parsed = urllib.parse.urlparse(url)

        return (
            parsed.scheme in ["http", "https"]
            and bool(parsed.netloc)
        )

    except Exception:
        return False


def get_first_real_link(link):

    if not link:
        return ""

    for value in str(link).splitlines():

        value = clean_url(value)

        if is_valid_url(value):
            return value

    return ""


def get_preferred_application_link(
    link="",
    qr=""
):

    qr_values = []

    if qr:

        if isinstance(qr, list):
            qr_values = qr
        else:
            qr_values = str(qr).splitlines()

    # QR preferred
    for value in qr_values:

        value = clean_url(value)

        if is_valid_url(value):
            return value

    # OCR URL fallback
    return get_first_real_link(link)


def extract_urls(text):

    if not text:
        return []

    patterns = [
        r'https?://[^\s<>"\'\]\)]+',
        r'www\.[^\s<>"\'\]\)]+',
        r'\b[a-zA-Z0-9.-]+\.(?:com|org|net|in|io|co|ai|dev|tech)(?:/[^\s<>"\'\]\)]*)?'
    ]

    found = []

    for pattern in patterns:

        try:
            matches = re.findall(
                pattern,
                text,
                flags=re.IGNORECASE
            )

            found.extend(matches)

        except Exception:
            pass

    result = []

    seen = set()

    for url in found:

        url = clean_url(url)

        if not is_valid_url(url):
            continue

        key = url.lower()

        if key not in seen:

            seen.add(key)
            result.append(url)

    return result


# ============================================================
# DATE FUNCTIONS
# ============================================================

MONTHS = {
    "jan": 1,
    "january": 1,
    "feb": 2,
    "february": 2,
    "mar": 3,
    "march": 3,
    "apr": 4,
    "april": 4,
    "may": 5,
    "jun": 6,
    "june": 6,
    "jul": 7,
    "july": 7,
    "aug": 8,
    "august": 8,
    "sep": 9,
    "sept": 9,
    "september": 9,
    "oct": 10,
    "october": 10,
    "nov": 11,
    "november": 11,
    "dec": 12,
    "december": 12
}


def parse_date(date_text):

    if not date_text:
        return None

    text = str(date_text).strip()

    formats = [
        "%d/%m/%Y",
        "%d-%m-%Y",
        "%d.%m.%Y",
        "%d/%m/%y",
        "%d-%m-%y",
        "%d.%m.%y",
        "%d %b %Y",
        "%d %B %Y",
        "%b %d %Y",
        "%B %d %Y",
        "%d %b, %Y",
        "%d %B, %Y",
        "%b %d, %Y",
        "%B %d, %Y"
    ]

    for fmt in formats:

        try:

            parsed = datetime.strptime(
                text,
                fmt
            ).date()

            return parsed

        except Exception:
            pass

    return None


def extract_dates(text):

    if not text:
        return []

    candidates = []

    patterns = [

        r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b",

        r"\b\d{1,2}\.\d{1,2}\.\d{2,4}\b",

        r"\b\d{1,2}\s+(?:Jan|January|Feb|February|Mar|March|Apr|April|May|Jun|June|Jul|July|Aug|August|Sep|Sept|September|Oct|October|Nov|November|Dec|December)[,]?\s+\d{4}\b",

        r"\b(?:Jan|January|Feb|February|Mar|March|Apr|April|May|Jun|June|Jul|July|Aug|August|Sep|Sept|September|Oct|October|Nov|November|Dec|December)\s+\d{1,2}[,]?\s+\d{4}\b"
    ]

    for pattern in patterns:

        try:

            matches = re.findall(
                pattern,
                text,
                flags=re.IGNORECASE
            )

            candidates.extend(matches)

        except Exception:
            pass

    result = []

    seen = set()

    for value in candidates:

        parsed = parse_date(value)

        if parsed:

            key = parsed.isoformat()

            if key not in seen:

                seen.add(key)
                result.append(parsed)

    return result


def find_last_date(text):

    if not text:
        return ""

    all_dates = extract_dates(text)

    if not all_dates:
        return ""

    deadline_keywords = [
        "last date",
        "deadline",
        "registration closes",
        "registration close",
        "apply by",
        "application closes",
        "applications close",
        "submission deadline",
        "submission closes",
        "closing date",
        "ends on",
        "ends"
    ]

    lines = text.splitlines()

    nearby_dates = []

    for index, line in enumerate(lines):

        line_lower = line.lower()

        if any(
            keyword in line_lower
            for keyword in deadline_keywords
        ):

            nearby_text = "\n".join(
                lines[
                    max(0, index - 1):
                    min(len(lines), index + 3)
                ]
            )

            dates = extract_dates(
                nearby_text
            )

            nearby_dates.extend(dates)

    if nearby_dates:

        return max(nearby_dates).strftime(
            "%d/%m/%Y"
        )

    # Fallback: latest date
    return max(all_dates).strftime(
        "%d/%m/%Y"
    )


# ============================================================
# HACKATHON NAME
# ============================================================

def _clean_ocr_title(value):

    if not value:
        return ""

    value = str(value).strip()

    value = re.sub(
        r"\s+",
        " ",
        value
    )

    value = value.strip(
        " :-|_.,;#*"
    )

    return value


def _is_bad_title_candidate(value):

    if not value:
        return True

    value = value.strip()

    if len(value) < 3:
        return True

    lower = value.lower()

    bad_words = [
        "deadline",
        "last date",
        "register now",
        "registration",
        "apply now",
        "click here",
        "scan here",
        "scan to",
        "visit",
        "www.",
        "http://",
        "https://",
        "organized by",
        "organised by",
        "contact",
        "email",
        "phone",
        "follow us",
        "terms",
        "conditions"
    ]

    if any(
        word in lower
        for word in bad_words
    ):
        return True

    if len(value) > 150:
        return True

    if re.fullmatch(
        r"[\d\s./:-]+",
        value
    ):
        return True

    return False


def find_hackathon_name_from_text(text):

    if not text:
        return "Hackathon"

    lines = [
        _clean_ocr_title(line)
        for line in text.splitlines()
    ]

    lines = [
        line
        for line in lines
        if line and not _is_bad_title_candidate(line)
    ]

    keywords = [
        "hackathon",
        "challenge",
        "innovation",
        "ideathon",
        "codefest",
        "coding",
        "competition",
        "contest",
        "summit"
    ]

    # First search for lines containing strong keywords
    for line in lines:

        lower = line.lower()

        if any(
            keyword in lower
            for keyword in keywords
        ):

            return line

    # Otherwise choose a good-looking short title
    candidates = []

    for line in lines:

        if 3 <= len(line) <= 100:

            words = line.split()

            if len(words) >= 2:

                candidates.append(line)

    if candidates:

        return candidates[0]

    return "Hackathon"


def find_hackathon_name_from_image(
    image,
    text=""
):

    # OCR based title
    return find_hackathon_name_from_text(text)


def find_hackathon_name(
    image,
    text=""
):

    return find_hackathon_name_from_image(
        image,
        text
    )


# ============================================================
# THEME
# ============================================================

def find_theme(text):

    if not text:
        return ""

    lower = text.lower()

    theme_map = {
        "Agriculture": [
            "agriculture",
            "farming",
            "farmer",
            "agritech",
            "agri"
        ],

        "Technology": [
            "technology",
            "software",
            "digital",
            "technology"
        ],

        "Artificial Intelligence": [
            "artificial intelligence",
            "machine learning",
            "deep learning",
            "generative ai",
            " ai ",
            "ml "
        ],

        "Healthcare": [
            "healthcare",
            "health care",
            "medical",
            "hospital",
            "medicine"
        ],

        "Education": [
            "education",
            "learning",
            "student",
            "school",
            "college"
        ],

        "Industrial": [
            "industrial",
            "manufacturing",
            "factory",
            "industry"
        ],

        "Finance": [
            "finance",
            "fintech",
            "banking",
            "financial"
        ],

        "Environment": [
            "environment",
            "climate",
            "sustainability",
            "green energy",
            "renewable"
        ],

        "Cybersecurity": [
            "cybersecurity",
            "cyber security",
            "security",
            "ethical hacking"
        ],

        "Blockchain": [
            "blockchain",
            "web3",
            "cryptocurrency"
        ]
    }

    for theme, keywords in theme_map.items():

        for keyword in keywords:

            if keyword in lower:

                return theme

    return ""


# ============================================================
# LINK COMBINATION
# ============================================================

def combine_links(urls):

    if not urls:
        return ""

    valid = []

    for url in urls:

        url = clean_url(url)

        if is_valid_url(url):
            valid.append(url)

    return "\n".join(valid)


# ============================================================
# IMAGE PROCESSING
# ============================================================

def process_image(uploaded_file):

    image = Image.open(
        uploaded_file
    ).convert("RGB")

    ocr_text = run_ocr(image)

    qr_values = extract_qr_codes(
        image
    )

    urls = extract_urls(
        ocr_text
    )

    # Add QR URLs
    for qr in qr_values:

        qr_clean = clean_url(qr)

        if is_valid_url(qr_clean):
            urls.append(qr_clean)

    # Remove duplicate URLs
    unique_urls = []

    seen = set()

    for url in urls:

        url = clean_url(url)

        if not is_valid_url(url):
            continue

        key = url.lower()

        if key not in seen:

            seen.add(key)
            unique_urls.append(url)

    link = combine_links(
        unique_urls
    )

    qr_text = "\n".join(
        qr_values
    )

    hackathon_name = find_hackathon_name(
        image,
        ocr_text
    )

    theme = find_theme(
        ocr_text
    )

    last_date = find_last_date(
        ocr_text
    )

    application_link = get_preferred_application_link(
        link,
        qr_values
    )

    return {
        "Hackathon Name": hackathon_name,
        "Theme / Category": theme,
        "Who Applied": "",
        "Status": "",
        "Last Date": last_date,
        "Link": application_link,
        "QR": qr_text,
        "_OCR": ocr_text,
        "_QR_LIST": qr_values
    }


# ============================================================
# EXCEL DATE CONVERSION
# ============================================================

def convert_excel_date(value):

    if pd.isna(value):
        return None

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    parsed = parse_date(
        str(value)
    )

    return parsed


# ============================================================
# ALERT MESSAGE
# ============================================================

def create_alert_message(
    hackathon_name,
    last_date,
    link,
    days_remaining=None
):

    if days_remaining is None:

        parsed = parse_date(
            last_date
        )

        if parsed:

            days_remaining = (
                parsed - date.today()
            ).days

    if days_remaining is None:
        days_remaining = ""

    message = (
        "🚨 Hackathon Deadline Alert\n\n"
        f"Hackathon - {hackathon_name}\n"
        f"Last Date - {last_date}\n"
        f"Days Remaining - {days_remaining}\n"
    )

    if link and is_valid_url(link):

        message += (
            f"Apply Link - {link}\n"
        )

    return message


# ============================================================
# EMAIL CONTENT
# ============================================================

def create_email_content(
    hackathon_name,
    last_date,
    link,
    days_remaining=None,
    status=None
):

    if days_remaining is None:

        parsed = parse_date(
            last_date
        )

        if parsed:

            days_remaining = (
                parsed - date.today()
            ).days

    if days_remaining is None:
        days_remaining = ""

    status_text = status or ""

    apply_button = ""

    if link and is_valid_url(link):

        safe_link = (
            link.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
        )

        apply_button = f"""
        <p style="margin-top:20px;">
            <a href="{safe_link}"
               style="
               display:inline-block;
               padding:12px 20px;
               background:#2563eb;
               color:white;
               text-decoration:none;
               border-radius:8px;
               font-weight:bold;">
                Apply Now
            </a>
        </p>
        """

    html = f"""
    <html>
    <body style="
        font-family:Arial,Helvetica,sans-serif;
        background:#f6f8fb;
        padding:25px;
    ">

        <div style="
            max-width:650px;
            margin:auto;
            background:white;
            padding:30px;
            border-radius:14px;
            border:1px solid #e5e7eb;
        ">

            <h2 style="color:#1e3a8a;">
                🚨 Hackathon Deadline Alert
            </h2>

            <hr>

            <p>
                <strong>Hackathon Name:</strong>
                {hackathon_name}
            </p>

            <p>
                <strong>Last Date:</strong>
                {last_date}
            </p>

            <p>
                <strong>Days Remaining:</strong>
                {days_remaining}
            </p>

            <p>
                <strong>Status:</strong>
                {status_text}
            </p>

            {apply_button}

            <hr>

            <p style="color:#64748b;font-size:13px;">
                This alert was generated by the
                Hackathon Image-to-Excel Extractor.
            </p>

        </div>

    </body>
    </html>
    """

    plain_text = (
        "🚨 Hackathon Deadline Alert\n\n"
        f"Hackathon Name: {hackathon_name}\n"
        f"Last Date: {last_date}\n"
        f"Days Remaining: {days_remaining}\n"
        f"Status: {status_text}\n"
    )

    if link and is_valid_url(link):

        plain_text += (
            f"Apply Link: {link}\n"
        )

    return html, plain_text


# ============================================================
# EMAIL CONFIGURATION
# ============================================================

def get_email_config():

    sender = os.getenv(
        "EMAIL_SENDER",
        ""
    )

    password = os.getenv(
        "EMAIL_APP_PASSWORD",
        ""
    )

    recipient = os.getenv(
        "EMAIL_RECIPIENT",
        ""
    )

    try:

        if hasattr(st, "secrets"):

            if "EMAIL_SENDER" in st.secrets:
                sender = st.secrets[
                    "EMAIL_SENDER"
                ]

            if "EMAIL_APP_PASSWORD" in st.secrets:
                password = st.secrets[
                    "EMAIL_APP_PASSWORD"
                ]

            if "EMAIL_RECIPIENT" in st.secrets:
                recipient = st.secrets[
                    "EMAIL_RECIPIENT"
                ]

    except Exception:
        pass

    # Sidebar values have priority
    if "email_sender" in locals():
        pass

    return sender, password, recipient


# ============================================================
# SEND EMAIL
# ============================================================

def send_email_alert(
    hackathon_name,
    last_date,
    link,
    days_remaining=None,
    status=None,
    sender_override=None,
    password_override=None,
    recipient_override=None
):

    sender = (
        sender_override
        if sender_override
        else os.getenv(
            "EMAIL_SENDER",
            ""
        )
    )

    password = (
        password_override
        if password_override
        else os.getenv(
            "EMAIL_APP_PASSWORD",
            ""
        )
    )

    recipient = (
        recipient_override
        if recipient_override
        else os.getenv(
            "EMAIL_RECIPIENT",
            ""
        )
    )

    try:

        if hasattr(st, "secrets"):

            if not sender:

                sender = st.secrets.get(
                    "EMAIL_SENDER",
                    ""
                )

            if not password:

                password = st.secrets.get(
                    "EMAIL_APP_PASSWORD",
                    ""
                )

            if not recipient:

                recipient = st.secrets.get(
                    "EMAIL_RECIPIENT",
                    ""
                )

    except Exception:
        pass

    if not sender:
        return False, "Sender email is missing."

    if not password:
        return False, "Gmail App Password is missing."

    if not recipient:
        return False, "Recipient email is missing."

    html_body, plain_body = create_email_content(
        hackathon_name=hackathon_name,
        last_date=last_date,
        link=link,
        days_remaining=days_remaining,
        status=status
    )

    msg = EmailMessage()

    msg["Subject"] = (
        f"🚨 Hackathon Deadline Alert - "
        f"{hackathon_name}"
    )

    msg["From"] = sender
    msg["To"] = recipient

    msg.set_content(
        plain_body
    )

    msg.add_alternative(
        html_body,
        subtype="html"
    )

    try:

        with smtplib.SMTP_SSL(
            "smtp.gmail.com",
            465
        ) as smtp:

            smtp.login(
                sender,
                password
            )

            smtp.send_message(
                msg
            )

        return True, (
            f"Email sent successfully to {recipient}"
        )

    except Exception as exc:

        return False, (
            f"Email failed: {exc}"
        )


# ============================================================
# CURRENT EXTRACTED EMAIL DETAILS
# ============================================================

def get_current_extracted_email_details():

    df = st.session_state.get(
        "extracted_hackathon_df"
    )

    if df is None:
        return None

    if df.empty:
        return None

    row = df.iloc[0]

    hackathon_name = str(
        row.get(
            "Hackathon Name",
            ""
        )
    ).strip()

    last_date = str(
        row.get(
            "Last Date",
            ""
        )
    ).strip()

    link = str(
        row.get(
            "Link",
            ""
        )
    ).strip()

    status = str(
        row.get(
            "Status",
            ""
        )
    ).strip()

    if status.lower() == "nan":
        status = ""

    parsed_date = parse_date(
        last_date
    )

    days_remaining = None

    if parsed_date:

        days_remaining = (
            parsed_date - date.today()
        ).days

    if not status:

        if days_remaining is None:
            status = "Unknown"

        elif days_remaining < 0:
            status = "Expired"

        elif days_remaining == 0:
            status = "Deadline Today"

        elif days_remaining <= 3:
            status = "Urgent"

        else:
            status = "Upcoming"

    return {
        "Hackathon Name": hackathon_name,
        "Last Date": last_date,
        "Link": link,
        "Days Remaining": days_remaining,
        "Status": status
    }


# ============================================================
# EMAIL BUTTON
# ============================================================

def show_email_button(
    alert,
    key
):

    if st.button(
        "📧 Send Email Alert",
        key=key,
        use_container_width=True
    ):

        success, message = send_email_alert(
            hackathon_name=alert.get(
                "Hackathon Name",
                ""
            ),
            last_date=alert.get(
                "Last Date",
                ""
            ),
            link=alert.get(
                "Link",
                ""
            ),
            days_remaining=alert.get(
                "Days Remaining"
            ),
            status=alert.get(
                "Status",
                ""
            )
        )

        st.session_state.email_last_success = success
        st.session_state.email_last_message = message
        st.session_state.email_last_time = (
            datetime.now().strftime(
                "%d/%m/%Y %H:%M:%S"
            )
        )

        if success:
            st.success(
                message
            )
        else:
            st.error(
                message
            )


# ============================================================
# DEADLINE ALERT CHECK
# ============================================================

def check_deadline_alerts(
    dataframe,
    check_date,
    alert_days=3
):

    alerts = []

    if dataframe is None:
        return alerts

    if dataframe.empty:
        return alerts

    for _, row in dataframe.iterrows():

        hackathon_name = str(
            row.get(
                "Hackathon Name",
                ""
            )
        ).strip()

        last_date_text = str(
            row.get(
                "Last Date",
                ""
            )
        ).strip()

        link = str(
            row.get(
                "Link",
                ""
            )
        ).strip()

        status_value = str(
            row.get(
                "Status",
                ""
            )
        ).strip()

        if status_value.lower() == "nan":
            status_value = ""

        if not last_date_text:
            continue

        deadline = parse_date(
            last_date_text
        )

        if deadline is None:
            continue

        remaining = (
            deadline - check_date
        ).days

        if status_value:

            status = status_value

        elif remaining < 0:

            status = "Expired"

        elif remaining == 0:

            status = "Deadline Today"

        elif remaining <= alert_days:

            status = "Urgent"

        else:

            status = "Upcoming"

        if (
            0 <= remaining <= alert_days
        ):

            alerts.append(
                {
                    "Hackathon Name":
                        hackathon_name,

                    "Last Date":
                        deadline.strftime(
                            "%d/%m/%Y"
                        ),

                    "Link":
                        link,

                    "Days Remaining":
                        remaining,

                    "Status":
                        status
                }
            )

    return alerts


# ============================================================
# WHATSAPP
# ============================================================

def show_whatsapp_button(
    message,
    key
):

    encoded = urllib.parse.quote(
        message
    )

    whatsapp_url = (
        "https://wa.me/?text="
        + encoded
    )

    st.link_button(
        "💬 Send via WhatsApp",
        whatsapp_url,
        use_container_width=True
    )


# ============================================================
# CREATE EXCEL
# ============================================================

def create_excel(dataframe):

    output = BytesIO()

    workbook = Workbook()

    worksheet = workbook.active

    worksheet.title = "Hackathons"

    columns = [
        "S.No",
        "Hackathon Name",
        "Theme / Category",
        "Who Applied",
        "Status",
        "Last Date",
        "Link",
        "QR"
    ]

    for col_index, column in enumerate(
        columns,
        start=1
    ):

        cell = worksheet.cell(
            row=1,
            column=col_index,
            value=column
        )

        cell.font = Font(
            bold=True
        )

        cell.fill = PatternFill(
            "solid",
            fgColor="1E3A8A"
        )

        cell.font = Font(
            color="FFFFFF",
            bold=True
        )

        cell.alignment = Alignment(
            horizontal="center",
            vertical="center"
        )

    for row_index, (_, row) in enumerate(
        dataframe.iterrows(),
        start=2
    ):

        for col_index, column in enumerate(
            columns,
            start=1
        ):

            value = row.get(
                column,
                ""
            )

            if pd.isna(value):
                value = ""

            if column == "Last Date":

                parsed = convert_excel_date(
                    value
                )

                if parsed:

                    cell = worksheet.cell(
                        row=row_index,
                        column=col_index,
                        value=parsed
                    )

                    cell.number_format = (
                        "dd/mm/yyyy"
                    )

                    continue

            cell = worksheet.cell(
                row=row_index,
                column=col_index,
                value=str(value)
            )

            if column in [
                "Link",
                "QR"
            ]:

                url = get_first_real_link(
                    str(value)
                )

                if url:

                    cell.value = url
                    cell.hyperlink = url
                    cell.font = Font(
                        color="0563C1",
                        underline="single"
                    )

    widths = {
        "A": 8,
        "B": 35,
        "C": 25,
        "D": 20,
        "E": 20,
        "F": 18,
        "G": 55,
        "H": 55
    }

    for column, width in widths.items():

        worksheet.column_dimensions[
            column
        ].width = width

    worksheet.freeze_panes = "A2"

    workbook.save(
        output
    )

    output.seek(0)

    return output.getvalue()


# ============================================================
# UPLOAD SECTION
# ============================================================

st.markdown(
    '<div class="section-title">📤 Upload Hackathon Posters</div>',
    unsafe_allow_html=True
)

st.markdown(
    '<div class="section-subtitle">'
    'Upload one or multiple hackathon poster images. '
    'The application will automatically process OCR, URLs, '
    'deadlines and QR codes.'
    '</div>',
    unsafe_allow_html=True
)

uploaded_files = st.file_uploader(
    "Choose hackathon images",
    type=[
        "png",
        "jpg",
        "jpeg",
        "webp"
    ],
    accept_multiple_files=True
)


# ============================================================
# UPLOAD METRICS
# ============================================================

file_count = (
    len(uploaded_files)
    if uploaded_files
    else 0
)

ocr_status = (
    "Ready"
    if os.path.exists(
        TESSERACT_PATH
    )
    else "Check"
)

qr_status = (
    "Ready"
    if PYZBAR_AVAILABLE
    else "OpenCV"
)

m1, m2, m3 = st.columns(3)

with m1:

    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-icon">🖼️</div>
            <div class="metric-label">FILES SELECTED</div>
            <div class="metric-value">
                {file_count}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

with m2:

    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-icon">🔎</div>
            <div class="metric-label">OCR ENGINE</div>
            <div class="metric-value">
                {ocr_status}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

with m3:

    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-icon">📱</div>
            <div class="metric-label">QR DETECTION</div>
            <div class="metric-value">
                {qr_status}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


# ============================================================
# EXTRACTION BUTTON
# ============================================================

if uploaded_files:

    st.markdown(
        "<br>",
        unsafe_allow_html=True
    )

    if st.button(
        "🚀 Extract Hackathon Details",
        type="primary",
        use_container_width=True
    ):

        results = []

        progress = st.progress(
            0,
            text="Starting extraction..."
        )

        total = len(
            uploaded_files
        )

        for index, uploaded_file in enumerate(
            uploaded_files,
            start=1
        ):

            progress.progress(
                (index - 1) / total,
                text=(
                    f"Processing "
                    f"{uploaded_file.name} "
                    f"({index}/{total})"
                )
            )

            try:

                result = process_image(
                    uploaded_file
                )

                results.append(
                    result
                )

            except Exception as exc:

                st.error(
                    f"Error processing "
                    f"{uploaded_file.name}: "
                    f"{exc}"
                )

        progress.progress(
            1.0,
            text="Extraction completed."
        )

        if results:

            # Deduplicate by Hackathon Name
            unique_results = []

            seen_names = set()

            for result in results:

                name = str(
                    result.get(
                        "Hackathon Name",
                        ""
                    )
                ).strip()

                key = name.lower()

                if key not in seen_names:

                    seen_names.add(key)

                    unique_results.append(
                        result
                    )

            display_rows = []

            for index, result in enumerate(
                unique_results,
                start=1
            ):

                display_rows.append(
                    {
                        "S.No": index,
                        "Hackathon Name":
                            result.get(
                                "Hackathon Name",
                                ""
                            ),
                        "Theme / Category":
                            result.get(
                                "Theme / Category",
                                ""
                            ),
                        "Who Applied":
                            result.get(
                                "Who Applied",
                                ""
                            ),
                        "Status":
                            result.get(
                                "Status",
                                ""
                            ),
                        "Last Date":
                            result.get(
                                "Last Date",
                                ""
                            ),
                        "Link":
                            result.get(
                                "Link",
                                ""
                            ),
                        "QR":
                            result.get(
                                "QR",
                                ""
                            )
                    }
                )

            extracted_df = pd.DataFrame(
                display_rows
            )

            st.session_state[
                "extracted_hackathon_df"
            ] = extracted_df

            st.success(
                f"✅ Successfully extracted "
                f"{len(extracted_df)} hackathon(s)."
            )


# ============================================================
# EXTRACTED DATA
# ============================================================

df = st.session_state.get(
    "extracted_hackathon_df"
)

if df is not None and not df.empty:

    st.markdown(
        '<div class="section-title">📊 Extracted Hackathon Data</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="section-subtitle">'
        'Review and edit the extracted information below. '
        'Changes made here are also used by the email alert system.'
        '</div>',
        unsafe_allow_html=True
    )

    edited_df = st.data_editor(
        df,
        use_container_width=True,
        hide_index=True,
        num_rows="dynamic",
        column_config={
            "S.No": st.column_config.NumberColumn(
                "S.No",
                disabled=True
            ),

            "Hackathon Name":
                st.column_config.TextColumn(
                    "Hackathon Name"
                ),

            "Theme / Category":
                st.column_config.TextColumn(
                    "Theme / Category"
                ),

            "Who Applied":
                st.column_config.TextColumn(
                    "Who Applied"
                ),

            "Status":
                st.column_config.TextColumn(
                    "Status"
                ),

            "Last Date":
                st.column_config.TextColumn(
                    "Last Date"
                ),

            "Link":
                st.column_config.LinkColumn(
                    "Link",
                    display_text="Open Link"
                ),

            "QR":
                st.column_config.TextColumn(
                    "QR"
                )
        },
        key="hackathon_data_editor"
    )

    st.session_state[
        "extracted_hackathon_df"
    ] = edited_df.copy()


    # ========================================================
    # DASHBOARD METRICS
    # ========================================================

    total_hackathons = len(
        edited_df
    )

    deadlines_found = 0
    links_found = 0
    urgent_alerts = 0

    for _, row in edited_df.iterrows():

        last_date_value = str(
            row.get(
                "Last Date",
                ""
            )
        ).strip()

        if parse_date(
            last_date_value
        ):
            deadlines_found += 1

        link_value = str(
            row.get(
                "Link",
                ""
            )
        ).strip()

        if is_valid_url(
            get_first_real_link(
                link_value
            )
        ):
            links_found += 1

        parsed = parse_date(
            last_date_value
        )

        if parsed:

            remaining = (
                parsed - date.today()
            ).days

            if 0 <= remaining <= 3:
                urgent_alerts += 1


    st.markdown(
        '<div class="section-title">📈 Dashboard</div>',
        unsafe_allow_html=True
    )

    d1, d2, d3, d4 = st.columns(4)

    with d1:

        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-icon">🏆</div>
                <div class="metric-label">
                    TOTAL HACKATHONS
                </div>
                <div class="metric-value">
                    {total_hackathons}
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with d2:

        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-icon">📅</div>
                <div class="metric-label">
                    DEADLINES FOUND
                </div>
                <div class="metric-value">
                    {deadlines_found}
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with d3:

        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-icon">🔗</div>
                <div class="metric-label">
                    APPLICATION LINKS
                </div>
                <div class="metric-value">
                    {links_found}
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )

    with d4:

        st.markdown(
            f"""
            <div class="metric-card">
                <div class="metric-icon">🚨</div>
                <div class="metric-label">
                    3-DAY ALERTS
                </div>
                <div class="metric-value">
                    {urgent_alerts}
                </div>
            </div>
            """,
            unsafe_allow_html=True
        )


    # ========================================================
    # DEADLINE ALERT SYSTEM
    # ========================================================

    st.markdown(
        '<div class="section-title">🚨 Deadline Alerts</div>',
        unsafe_allow_html=True
    )

    alerts = check_deadline_alerts(
        edited_df,
        date.today(),
        alert_days=3
    )

    if alerts:

        for alert_index, alert in enumerate(
            alerts
        ):

            remaining = alert[
                "Days Remaining"
            ]

            if remaining == 0:

                css_class = "danger-card"

                icon = "🔴"

                title = "Deadline Today"

            elif remaining == 1:

                css_class = "danger-card"

                icon = "🔴"

                title = "1 Day Remaining"

            elif remaining <= 3:

                css_class = "alert-card"

                icon = "🟠"

                title = (
                    f"{remaining} Days Remaining"
                )

            else:

                css_class = "success-card"

                icon = "🟢"

                title = "Upcoming"

            st.markdown(
                f"""
                <div class="{css_class}">
                    <h4>
                        {icon}
                        {alert['Hackathon Name']}
                    </h4>

                    <p>
                        <strong>
                            {title}
                        </strong>
                    </p>

                    <p>
                        📅 Last Date:
                        <strong>
                            {alert['Last Date']}
                        </strong>
                    </p>

                    <p>
                        ⏳ Days Remaining:
                        <strong>
                            {remaining}
                        </strong>
                    </p>

                    <p>
                        📌 Status:
                        <strong>
                            {alert['Status']}
                        </strong>
                    </p>
                </div>
                """,
                unsafe_allow_html=True
            )

            col_a, col_b = st.columns(2)

            with col_a:

                show_whatsapp_button(
                    create_alert_message(
                        alert[
                            "Hackathon Name"
                        ],
                        alert[
                            "Last Date"
                        ],
                        alert[
                            "Link"
                        ],
                        alert[
                            "Days Remaining"
                        ]
                    ),
                    key=(
                        f"whatsapp_alert_"
                        f"{alert_index}"
                    )
                )

            with col_b:

                show_email_button(
                    alert,
                    key=(
                        f"email_alert_"
                        f"{alert_index}"
                    )
                )

    else:

        st.markdown(
            """
            <div class="success-card">
                <strong>✅ No urgent deadline alerts.</strong>
                <br>
                There are currently no hackathons
                with deadlines within the next 3 days.
            </div>
            """,
            unsafe_allow_html=True
        )


    # ========================================================
    # EMAIL RESULT
    # ========================================================

    if st.session_state.email_last_message:

        if st.session_state.email_last_success:

            st.success(
                "📧 "
                + st.session_state.email_last_message
                + (
                    " | "
                    + st.session_state.email_last_time
                )
            )

        else:

            st.error(
                st.session_state.email_last_message
            )


    # ========================================================
    # EXCEL EXPORT
    # ========================================================

    st.markdown(
        '<div class="section-title">📥 Export Results</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        '<div class="section-subtitle">'
        'Download the extracted and edited hackathon '
        'information as an Excel file.'
        '</div>',
        unsafe_allow_html=True
    )

    excel_data = create_excel(
        edited_df
    )

    st.download_button(
        label="📊 Download Excel File",
        data=excel_data,
        file_name=(
            "Hackathon_Extracted_Data.xlsx"
        ),
        mime=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
        type="primary",
        use_container_width=True
    )


    # ========================================================
    # OCR / QR DETAILS
    # ========================================================

    st.markdown(
        '<div class="section-title">🔍 Extraction Details</div>',
        unsafe_allow_html=True
    )

    col1, col2 = st.columns(2)

    with col1:

        with st.expander(
            "📝 View OCR Text"
        ):

            for index, (_, row) in enumerate(
                edited_df.iterrows()
            ):

                st.markdown(
                    f"### {index + 1}. "
                    f"{row.get('Hackathon Name', '')}"
                )

                # Find matching result from uploaded data
                name = str(
                    row.get(
                        "Hackathon Name",
                        ""
                    )
                ).strip()

                found_text = ""

                if uploaded_files:

                    for uploaded_file in uploaded_files:

                        try:

                            temp_result = process_image(
                                uploaded_file
                            )

                            if str(
                                temp_result.get(
                                    "Hackathon Name",
                                    ""
                                )
                            ).strip().lower() == name.lower():

                                found_text = (
                                    temp_result.get(
                                        "_OCR",
                                        ""
                                    )
                                )

                                break

                        except Exception:
                            pass

                if found_text:

                    st.text_area(
                        "OCR",
                        found_text,
                        height=220,
                        key=f"ocr_{index}"
                    )

                else:

                    st.info(
                        "OCR text is not available "
                        "for this row."
                    )

    with col2:

        with st.expander(
            "📱 View QR Codes"
        ):

            for index, (_, row) in enumerate(
                edited_df.iterrows()
            ):

                qr_value = str(
                    row.get(
                        "QR",
                        ""
                    )
                ).strip()

                st.markdown(
                    f"### {index + 1}. "
                    f"{row.get('Hackathon Name', '')}"
                )

                if qr_value:

                    st.code(
                        qr_value
                    )

                else:

                    st.info(
                        "No QR code detected."
                    )


# ============================================================
# TESTING SECTION
# ============================================================

st.markdown(
    '<div class="section-title">🧪 Deadline Alert Testing</div>',
    unsafe_allow_html=True
)

st.markdown(
    """
    <div class="neutral-card">
        <strong>Testing Mode</strong><br>
        Use this section to test the 3-day deadline
        alert system with a selected date.
    </div>
    """,
    unsafe_allow_html=True
)


# ============================================================
# CURRENT EXTRACTED DATA FOR TEST
# ============================================================

current_details = (
    get_current_extracted_email_details()
)


if current_details:

    extracted_deadline = parse_date(
        current_details[
            "Last Date"
        ]
    )

    extracted_hackathon_name = (
        current_details[
            "Hackathon Name"
        ]
    )

    extracted_link = (
        current_details[
            "Link"
        ]
    )

else:

    # Fallback testing data
    extracted_deadline = date(
        2026,
        9,
        30
    )

    extracted_hackathon_name = (
        "GENESIS Investment 2.0"
    )

    extracted_link = ""


if extracted_deadline:

    three_day_start = (
        extracted_deadline
        - timedelta(days=3)
    )

    actual_remaining = (
        extracted_deadline
        - date.today()
    ).days

    if actual_remaining < 0:

        current_status = "Expired"

    elif actual_remaining == 0:

        current_status = "Deadline Today"

    elif actual_remaining <= 3:

        current_status = "Urgent"

    else:

        current_status = "Upcoming"

else:

    three_day_start = None
    actual_remaining = None
    current_status = "Unknown"


t1, t2, t3, t4 = st.columns(4)

with t1:

    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-icon">🏆</div>
            <div class="metric-label">
                HACKATHON
            </div>
            <div class="metric-value"
                 style="font-size:18px;">
                {extracted_hackathon_name}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

with t2:

    deadline_display = (
        extracted_deadline.strftime(
            "%d/%m/%Y"
        )
        if extracted_deadline
        else "Not Found"
    )

    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-icon">📅</div>
            <div class="metric-label">
                DEADLINE
            </div>
            <div class="metric-value"
                 style="font-size:22px;">
                {deadline_display}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

with t3:

    remaining_display = (
        str(actual_remaining)
        if actual_remaining is not None
        else "N/A"
    )

    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-icon">⏳</div>
            <div class="metric-label">
                DAYS REMAINING
            </div>
            <div class="metric-value">
                {remaining_display}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

with t4:

    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-icon">📌</div>
            <div class="metric-label">
                STATUS
            </div>
            <div class="metric-value"
                 style="font-size:20px;">
                {current_status}
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )


# ============================================================
# MANUAL TEST DATE
# ============================================================

test_date = st.date_input(
    "Select Test Date",
    value=date.today(),
    key="manual_test_date"
)


if st.button(
    "🧪 Run Deadline Alert Test",
    type="primary",
    use_container_width=True
):

    if extracted_deadline:

        test_df = pd.DataFrame(
            [
                {
                    "S.No": 1,
                    "Hackathon Name":
                        extracted_hackathon_name,
                    "Theme / Category": "",
                    "Who Applied": "",
                    "Status": "",
                    "Last Date":
                        extracted_deadline.strftime(
                            "%d/%m/%Y"
                        ),
                    "Link":
                        extracted_link,
                    "QR": ""
                }
            ]
        )

        test_alerts = check_deadline_alerts(
            test_df,
            test_date,
            alert_days=3
        )

        remaining = (
            extracted_deadline
            - test_date
        ).days

        st.markdown(
            f"""
            <div class="info-card">
                <h4>🧪 Test Result</h4>

                <p>
                    <strong>Hackathon:</strong>
                    {extracted_hackathon_name}
                </p>

                <p>
                    <strong>Deadline:</strong>
                    {extracted_deadline.strftime('%d/%m/%Y')}
                </p>

                <p>
                    <strong>Test Date:</strong>
                    {test_date.strftime('%d/%m/%Y')}
                </p>

                <p>
                    <strong>Days Remaining:</strong>
                    {remaining}
                </p>
            </div>
            """,
            unsafe_allow_html=True
        )

        if test_alerts:

            st.success(
                "🚨 Alert triggered successfully."
            )

            for test_index, alert in enumerate(
                test_alerts
            ):

                st.markdown(
                    f"""
                    <div class="alert-card">

                    <h4>
                        🚨 {alert['Hackathon Name']}
                    </h4>

                    <p>
                        Last Date:
                        <strong>
                            {alert['Last Date']}
                        </strong>
                    </p>

                    <p>
                        Days Remaining:
                        <strong>
                            {alert['Days Remaining']}
                        </strong>
                    </p>

                    <p>
                        Status:
                        <strong>
                            {alert['Status']}
                        </strong>
                    </p>

                    </div>
                    """,
                    unsafe_allow_html=True
                )

                test_col1, test_col2 = st.columns(2)

                with test_col1:

                    show_whatsapp_button(
                        create_alert_message(
                            alert[
                                "Hackathon Name"
                            ],
                            alert[
                                "Last Date"
                            ],
                            alert[
                                "Link"
                            ],
                            alert[
                                "Days Remaining"
                            ]
                        ),
                        key=(
                            f"test_whatsapp_"
                            f"{test_index}"
                        )
                    )

                with test_col2:

                    show_email_button(
                        alert,
                        key=(
                            f"test_email_"
                            f"{test_index}"
                        )
                    )

        else:

            st.info(
                "ℹ️ No alert triggered for this "
                "test date because it is outside "
                "the 3-day alert window."
            )

    else:

        st.warning(
            "⚠️ No valid deadline available for testing."
        )


# ============================================================
# TESTING INFORMATION
# ============================================================

with st.expander(
    "ℹ️ How the 3-Day Alert System Works"
):

    if extracted_deadline:

        st.markdown(
            f"""
            ### 📅 Deadline Information

            **Hackathon:**  
            {extracted_hackathon_name}

            **Last Date:**  
            {extracted_deadline.strftime('%d/%m/%Y')}

            **3-Day Alert Starts:**  
            {three_day_start.strftime('%d/%m/%Y')}

            **Today's Date:**  
            {date.today().strftime('%d/%m/%Y')}

            **Current Days Remaining:**  
            {actual_remaining}

            ### Alert Schedule

            | Days Remaining | Alert |
            |---:|---|
            | 4 days | No 3-day alert |
            | 3 days | 🚨 Alert |
            | 2 days | 🚨 Alert |
            | 1 day | 🚨 Alert |
            | 0 days | 🔴 Deadline Today |
            | Expired | ❌ Expired |

            The system calculates the remaining days from
            the actual deadline and the selected/current date.
            """
        )

    else:

        st.info(
            "No deadline is currently available."
        )


# ============================================================
# EMAIL TEST - CURRENT EXTRACTED DETAILS
# ============================================================

if current_details:

    st.markdown(
        '<div class="section-title">📧 Current Extracted Email Details</div>',
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div class="section-subtitle">
        The email below uses the current extracted/editable
        Hackathon Name, Last Date, Status and Link.
        </div>
        """,
        unsafe_allow_html=True
    )

    email_preview_col1, email_preview_col2 = st.columns(2)

    with email_preview_col1:

        st.markdown(
            f"""
            <div class="info-card">

            <p>
                <strong>Hackathon Name:</strong><br>
                {current_details['Hackathon Name']}
            </p>

            <p>
                <strong>Last Date:</strong><br>
                {current_details['Last Date']}
            </p>

            <p>
                <strong>Days Remaining:</strong><br>
                {current_details['Days Remaining']}
            </p>

            <p>
                <strong>Status:</strong><br>
                {current_details['Status']}
            </p>

            </div>
            """,
            unsafe_allow_html=True
        )

    with email_preview_col2:

        link_preview = (
            current_details["Link"]
            if current_details["Link"]
            else "No application link"
        )

        st.markdown(
            f"""
            <div class="info-card">

            <p>
                <strong>Application Link:</strong>
            </p>

            <p>
                {link_preview}
            </p>

            <p style="color:#64748b;">
                Email alerts use these current values,
                including values edited in the table above.
            </p>

            </div>
            """,
            unsafe_allow_html=True
        )

    if st.button(
        "📧 Send Email Using Current Extracted Details",
        use_container_width=True,
        key="send_current_extracted_email"
    ):

        success, message = send_email_alert(
            hackathon_name=current_details[
                "Hackathon Name"
            ],
            last_date=current_details[
                "Last Date"
            ],
            link=current_details[
                "Link"
            ],
            days_remaining=current_details[
                "Days Remaining"
            ],
            status=current_details[
                "Status"
            ],
            sender_override=email_sender,
            password_override=email_app_password,
            recipient_override=email_recipient
        )

        st.session_state.email_last_success = success
        st.session_state.email_last_message = message
        st.session_state.email_last_time = (
            datetime.now().strftime(
                "%d/%m/%Y %H:%M:%S"
            )
        )

        if success:
            st.success(
                message
            )
        else:
            st.error(
                message
            )


# ============================================================
# FOOTER
# ============================================================

st.markdown(
    """
    <div class="footer">
        🏆 Hackathon Image-to-Excel Extractor
        &nbsp;•&nbsp;
        OCR + QR + Excel + Deadline Alerts
        <br>
        Built with Streamlit, Python, OpenCV and Tesseract OCR
    </div>
    """,
    unsafe_allow_html=True
)