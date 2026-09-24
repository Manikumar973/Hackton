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
# OPTIONAL QR DECODER
# ============================================================

try:
    from pyzbar.pyzbar import decode as pyzbar_decode
    PYZBAR_AVAILABLE = True
except Exception:
    PYZBAR_AVAILABLE = False


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Hackathon Image to Excel Extractor",
    page_icon="🏆",
    layout="wide"
)


# ============================================================
# TESSERACT CONFIGURATION
# ============================================================

TESSERACT_PATH = r"C:\Program Files\Tesseract-OCR\tesseract.exe"

if os.path.exists(TESSERACT_PATH):
    pytesseract.pytesseract.tesseract_cmd = TESSERACT_PATH


# ============================================================
# TITLE
# ============================================================

st.title("🏆 Hackathon Image-to-Excel Extractor")

st.write(
    "Upload hackathon posters/images and automatically extract "
    "hackathon name, theme, last date, links and QR codes."
)


# ============================================================
# IMAGE → OPENCV
# ============================================================

def pil_to_cv(image):
    rgb = np.array(image.convert("RGB"))
    return cv2.cvtColor(rgb, cv2.COLOR_RGB2BGR)


# ============================================================
# OCR
# ============================================================

def run_ocr(image):

    cv_image = pil_to_cv(image)

    enlarged = cv2.resize(
        cv_image,
        None,
        fx=2,
        fy=2,
        interpolation=cv2.INTER_CUBIC
    )

    texts = []

    try:
        text1 = pytesseract.image_to_string(enlarged)
        if text1.strip():
            texts.append(text1)
    except Exception:
        pass

    try:
        gray = cv2.cvtColor(
            enlarged,
            cv2.COLOR_BGR2GRAY
        )

        text2 = pytesseract.image_to_string(gray)

        if text2.strip():
            texts.append(text2)
    except Exception:
        pass

    try:
        gray = cv2.cvtColor(
            enlarged,
            cv2.COLOR_BGR2GRAY
        )

        _, thresh = cv2.threshold(
            gray,
            0,
            255,
            cv2.THRESH_BINARY + cv2.THRESH_OTSU
        )

        text3 = pytesseract.image_to_string(thresh)

        if text3.strip():
            texts.append(text3)
    except Exception:
        pass

    try:
        gray = cv2.cvtColor(
            enlarged,
            cv2.COLOR_BGR2GRAY
        )

        adaptive = cv2.adaptiveThreshold(
            gray,
            255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY,
            31,
            11
        )

        text4 = pytesseract.image_to_string(adaptive)

        if text4.strip():
            texts.append(text4)
    except Exception:
        pass

    final_lines = []

    for text in texts:
        for line in text.splitlines():

            line = line.strip()

            if line and line not in final_lines:
                final_lines.append(line)

    return "\n".join(final_lines)


# ============================================================
# QR CODE DETECTION
# ============================================================

def extract_qr_codes(image):

    cv_image = pil_to_cv(image)

    results = []

    def add_result(data):

        if data is None:
            return

        data = str(data).strip()

        if data and data not in results:
            results.append(data)

    images_to_check = [
        cv_image
    ]

    try:
        enlarged_2x = cv2.resize(
            cv_image,
            None,
            fx=2,
            fy=2,
            interpolation=cv2.INTER_CUBIC
        )

        images_to_check.append(enlarged_2x)
    except Exception:
        pass

    try:
        enlarged_3x = cv2.resize(
            cv_image,
            None,
            fx=3,
            fy=3,
            interpolation=cv2.INTER_CUBIC
        )

        images_to_check.append(enlarged_3x)
    except Exception:
        pass

    try:
        gray = cv2.cvtColor(
            enlarged_2x,
            cv2.COLOR_BGR2GRAY
        )

        images_to_check.append(gray)

        _, otsu = cv2.threshold(
            gray,
            0,
            255,
            cv2.THRESH_BINARY + cv2.THRESH_OTSU
        )

        images_to_check.append(otsu)

        adaptive = cv2.adaptiveThreshold(
            gray,
            255,
            cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
            cv2.THRESH_BINARY,
            31,
            11
        )

        images_to_check.append(adaptive)

    except Exception:
        pass

    # --------------------------------------------------------
    # OpenCV QR Detection
    # --------------------------------------------------------

    for current_image in images_to_check:

        try:
            detector = cv2.QRCodeDetector()

            data, points, _ = detector.detectAndDecode(
                current_image
            )

            if data:
                add_result(data)

        except Exception:
            pass

        try:
            detector_multi = cv2.QRCodeDetector()

            result = detector_multi.detectAndDecodeMulti(
                current_image
            )

            if len(result) == 4:

                success, decoded_info, points, straight_qrcode = result

                if success:

                    for data in decoded_info:
                        add_result(data)

        except Exception:
            pass

    # --------------------------------------------------------
    # pyzbar
    # --------------------------------------------------------

    if PYZBAR_AVAILABLE:

        for current_image in images_to_check:

            try:
                decoded_objects = pyzbar_decode(
                    current_image
                )

                for obj in decoded_objects:

                    try:
                        data = obj.data.decode(
                            "utf-8",
                            errors="ignore"
                        )

                        add_result(data)

                    except Exception:
                        pass

            except Exception:
                pass

    return results


# ============================================================
# CLEAN URL
# ============================================================

def clean_url(url):

    url = str(url).strip()

    url = url.rstrip(
        ".,;:)]}\"'"
    )

    if url.startswith("www."):
        url = "https://" + url

    return url


# ============================================================
# VALID APPLICATION URL
# ============================================================

def is_valid_url(url):

    if not url:
        return False

    url = str(url).strip()

    if not url:
        return False

    lower = url.lower()

    if lower in [
        "https://example.com",
        "http://example.com",
        "https://example.com/apply",
        "http://example.com/apply"
    ]:
        return False

    return (
        lower.startswith("http://")
        or lower.startswith("https://")
    )


# ============================================================
# GET FIRST REAL URL
# ============================================================

def get_first_real_link(link):

    if link is None:
        return ""

    for item in str(link).splitlines():

        item = clean_url(item)

        if is_valid_url(item):
            return item

    return ""


# ============================================================
# PREFERRED APPLICATION LINK
# ============================================================

def get_preferred_application_link(link="", qr=""):
    """Prefer the real URL decoded from the QR code.
    Fall back to an OCR-extracted URL when no QR URL exists.
    Never return example.com placeholders.
    """
    qr_link = get_first_real_link(qr)
    if qr_link and "example.com" not in qr_link.lower():
        return qr_link

    normal_link = get_first_real_link(link)
    if normal_link and "example.com" not in normal_link.lower():
        return normal_link

    return ""


# ============================================================
# URL EXTRACTION
# ============================================================

def extract_urls(text):

    urls = []

    pattern1 = r'https?://[^\s<>"\']+'

    for match in re.findall(
        pattern1,
        text,
        flags=re.IGNORECASE
    ):

        url = clean_url(match)

        if is_valid_url(url) and url not in urls:
            urls.append(url)

    pattern2 = (
        r'www\.[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}'
        r'[^\s<>"\']*'
    )

    for match in re.findall(
        pattern2,
        text,
        flags=re.IGNORECASE
    ):

        url = clean_url(match)

        if is_valid_url(url) and url not in urls:
            urls.append(url)

    pattern3 = (
        r'\b[a-zA-Z0-9.-]+\.'
        r'(?:com|org|net|in|io|ai|co|tech|dev)'
        r'(?:/[^\s<>"\']*)?'
    )

    for match in re.findall(
        pattern3,
        text,
        flags=re.IGNORECASE
    ):

        url = clean_url(match)

        if not url.startswith("http"):
            url = "https://" + url

        if is_valid_url(url) and url not in urls:
            urls.append(url)

    return urls


# ============================================================
# DATE EXTRACTION
# ============================================================

def extract_dates(text):

    dates = []

    patterns = [

        r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b",

        r"\b\d{1,2}\.\d{1,2}\.\d{2,4}\b",

        r"\b\d{1,2}\s+"
        r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)"
        r"[a-z]*\s+\d{2,4}\b",

        r"\b"
        r"(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)"
        r"[a-z]*\s+\d{1,2},?\s+\d{2,4}\b"
    ]

    for pattern in patterns:

        matches = re.findall(
            pattern,
            text,
            flags=re.IGNORECASE
        )

        for match in matches:

            if match not in dates:
                dates.append(match)

    return dates


# ============================================================
# PARSE DATE
# ============================================================

def parse_date(date_text):

    if date_text is None:
        return None

    date_text = str(date_text).strip()

    formats = [

        "%d/%m/%Y",
        "%d/%m/%y",

        "%d-%m-%Y",
        "%d-%m-%y",

        "%d.%m.%Y",
        "%d.%m.%y",

        "%d %b %Y",
        "%d %B %Y",

        "%b %d %Y",
        "%B %d %Y",

        "%b %d, %Y",
        "%B %d, %Y"
    ]

    for fmt in formats:

        try:
            return datetime.strptime(
                date_text,
                fmt
            ).date()

        except Exception:
            pass

    return None


# ============================================================
# FIND LAST DATE
# ============================================================

def find_last_date(text):

    if not text:
        return None

    dates = extract_dates(text)

    parsed_dates = []

    for date_text in dates:

        parsed = parse_date(date_text)

        if parsed:
            parsed_dates.append(parsed)

    if not parsed_dates:
        return None

    deadline_keywords = [

        "last date",
        "last day",
        "end date",
        "ending date",
        "deadline",
        "application deadline",
        "registration deadline",
        "submission deadline",
        "closing date",
        "closing",
        "registration closes",
        "applications close",
        "application closes",
        "submission closes",
        "apply before",
        "apply by",
        "submit by",
        "ends on",
        "ends",
        "close on",
        "closes on"
    ]

    lines = text.splitlines()

    # --------------------------------------------------------
    # Keyword + nearby lines
    # --------------------------------------------------------

    for index, line in enumerate(lines):

        lower_line = line.lower()

        for keyword in deadline_keywords:

            keyword_pos = lower_line.find(keyword)

            if keyword_pos == -1:
                continue

            start_index = max(
                0,
                index - 2
            )

            end_index = min(
                len(lines),
                index + 3
            )

            nearby_text = "\n".join(
                lines[start_index:end_index]
            )

            nearby_dates = extract_dates(
                nearby_text
            )

            candidates = []

            for date_text in nearby_dates:

                parsed = parse_date(
                    date_text
                )

                if not parsed:
                    continue

                date_pos = nearby_text.lower().find(
                    date_text.lower()
                )

                if date_pos >= 0:
                    distance = abs(
                        date_pos - keyword_pos
                    )
                else:
                    distance = 999999

                candidates.append(
                    (
                        distance,
                        parsed
                    )
                )

            if candidates:

                candidates.sort(
                    key=lambda item: (
                        item[0],
                        item[1]
                    )
                )

                return candidates[0][1]

    # --------------------------------------------------------
    # Keyword followed by date
    # --------------------------------------------------------

    lower_text = text.lower()

    for keyword in deadline_keywords:

        search_start = 0

        while True:

            keyword_pos = lower_text.find(
                keyword,
                search_start
            )

            if keyword_pos == -1:
                break

            after_keyword = text[
                keyword_pos:
                keyword_pos + 180
            ]

            nearby_dates = extract_dates(
                after_keyword
            )

            if nearby_dates:

                parsed = parse_date(
                    nearby_dates[0]
                )

                if parsed:
                    return parsed

            search_start = (
                keyword_pos
                + len(keyword)
            )

    # --------------------------------------------------------
    # Fallback: latest valid date
    # --------------------------------------------------------

    return max(parsed_dates)


# ============================================================
# HACKATHON NAME
# ============================================================

def _clean_ocr_title(text):

    text = re.sub(
        r"\s+",
        " ",
        str(text)
    ).strip(" |-_•·")

    return text


def _is_bad_title_candidate(text):

    if not text:
        return True

    cleaned = _clean_ocr_title(text)

    lower = cleaned.lower()

    if len(cleaned) < 3 or len(cleaned) > 150:
        return True

    if re.fullmatch(
        r"[\d\s./,:;|()\-–—]+",
        cleaned
    ):
        return True

    if re.search(
        r"(?:https?://|www\.|@[a-z0-9.-]+\.)",
        lower
    ):
        return True

    rejected = {

        "hackathon",
        "hack",
        "challenge",
        "coding",
        "codefest",
        "register",
        "register now",
        "apply",
        "apply now",
        "apply here",
        "join now",
        "learn more",
        "click here",
        "last date",
        "last day",
        "end date",
        "deadline",
        "registration",
        "registration deadline",
        "closing date",
        "closing",
        "submit now",
        "submit"
    }

    if lower in rejected:
        return True

    if any(
        phrase in lower
        for phrase in [
            "register now",
            "apply now",
            "apply here",
            "join now",
            "click here",
            "last date",
            "end date",
            "deadline",
            "closing date",
            "registration deadline",
            "submit by"
        ]
    ):
        return True

    return False


def find_hackathon_name_from_text(text):

    lines = [

        _clean_ocr_title(line)

        for line in text.splitlines()

        if _clean_ocr_title(line)
    ]

    keywords = [

        "hackathon",
        "hack",
        "challenge",
        "innovation",
        "genesis",
        "codefest",
        "coding"
    ]

    for line in lines:

        lower_line = line.lower()

        if _is_bad_title_candidate(line):
            continue

        if any(
            keyword in lower_line
            for keyword in keywords
        ):
            return line

    for line in lines:

        if _is_bad_title_candidate(line):
            continue

        if re.search(
            r"\b\d{1,2}[/-]\d{1,2}[/-]\d{2,4}\b",
            line
        ):
            continue

        return line

    return "Unknown Hackathon"


def find_hackathon_name_from_image(
    image,
    fallback_text=""
):

    try:

        cv_image = pil_to_cv(image)

        variants = [

            cv_image,

            cv2.resize(
                cv_image,
                None,
                fx=2,
                fy=2,
                interpolation=cv2.INTER_CUBIC
            )
        ]

        candidates = []

        for variant_index, current_image in enumerate(
            variants
        ):

            try:

                data = pytesseract.image_to_data(
                    current_image,
                    config="--psm 11",
                    output_type=pytesseract.Output.DICT
                )

            except Exception:
                continue

            groups = {}

            count = len(
                data.get(
                    "text",
                    []
                )
            )

            for i in range(count):

                raw_word = str(
                    data["text"][i]
                ).strip()

                if not raw_word:
                    continue

                try:
                    confidence = float(
                        data["conf"][i]
                    )
                except Exception:
                    confidence = 0.0

                if confidence < 20:
                    continue

                block = data.get(
                    "block_num",
                    [0] * count
                )[i]

                par = data.get(
                    "par_num",
                    [0] * count
                )[i]

                line_num = data.get(
                    "line_num",
                    [0] * count
                )[i]

                key = (
                    block,
                    par,
                    line_num
                )

                try:

                    x = int(
                        data["left"][i]
                    )

                    y = int(
                        data["top"][i]
                    )

                    width = int(
                        data["width"][i]
                    )

                    height = int(
                        data["height"][i]
                    )

                except Exception:
                    continue

                groups.setdefault(
                    key,
                    []
                ).append({

                    "text": raw_word,

                    "confidence": confidence,

                    "x": x,

                    "y": y,

                    "width": width,

                    "height": height
                })

            for words in groups.values():

                words.sort(
                    key=lambda item: item["x"]
                )

                candidate_text = _clean_ocr_title(
                    " ".join(
                        item["text"]
                        for item in words
                    )
                )

                if _is_bad_title_candidate(
                    candidate_text
                ):
                    continue

                heights = [
                    item["height"]
                    for item in words
                ]

                confidences = [
                    item["confidence"]
                    for item in words
                ]

                avg_height = (
                    sum(heights)
                    / len(heights)
                )

                max_height = max(
                    heights
                )

                avg_confidence = (
                    sum(confidences)
                    / len(confidences)
                )

                lower = candidate_text.lower()

                keyword_bonus = 0

                if "hackathon" in lower:
                    keyword_bonus += 20

                elif any(
                    word in lower
                    for word in [
                        "challenge",
                        "codefest",
                        "innovation"
                    ]
                ):
                    keyword_bonus += 8

                min_y = min(
                    item["y"]
                    for item in words
                )

                image_height = max(
                    current_image.shape[0],
                    1
                )

                position_bonus = max(
                    0,
                    10 - int(
                        20
                        * min_y
                        / image_height
                    )
                )

                word_count_bonus = min(
                    len(words),
                    6
                )

                score = (
                    avg_height * 4.0
                    + max_height * 1.5
                    + avg_confidence * 0.12
                    + keyword_bonus
                    + position_bonus
                    + word_count_bonus
                )

                candidates.append({

                    "text": candidate_text,

                    "score": score,

                    "avg_height": avg_height,

                    "max_height": max_height,

                    "confidence": avg_confidence,

                    "variant": variant_index
                })

        if candidates:

            candidates.sort(
                key=lambda item: (
                    item["score"],
                    item["avg_height"],
                    item["max_height"],
                    item["confidence"]
                ),
                reverse=True
            )

            return candidates[0]["text"]

    except Exception:
        pass

    return find_hackathon_name_from_text(
        fallback_text
    )


def find_hackathon_name(text):

    return find_hackathon_name_from_text(
        text
    )


# ============================================================
# THEME / CATEGORY
# ============================================================

def find_theme(text):

    lower_text = text.lower()

    themes = {

        "Agriculture": [
            "agriculture",
            "farming",
            "farm",
            "agri",
            "crop"
        ],

        "Technology": [
            "technology",
            "tech",
            "software",
            "digital"
        ],

        "AI": [
            "artificial intelligence",
            "machine learning",
            " ai ",
            "ml "
        ],

        "Healthcare": [
            "healthcare",
            "health",
            "medical",
            "hospital"
        ],

        "Education": [
            "education",
            "learning",
            "student",
            "school"
        ],

        "Industrial": [
            "industrial",
            "manufacturing",
            "industry"
        ],

        "Finance": [
            "finance",
            "fintech",
            "banking",
            "investment"
        ],

        "Environment": [
            "environment",
            "climate",
            "green",
            "sustainability"
        ],

        "Cybersecurity": [
            "cybersecurity",
            "cyber security",
            "security"
        ],

        "Blockchain": [
            "blockchain",
            "web3",
            "crypto"
        ]
    }

    for theme, keywords in themes.items():

        for keyword in keywords:

            if keyword in lower_text:
                return theme

    return "Unknown"


# ============================================================
# COMBINE LINKS
# ============================================================

def combine_links(urls):

    if not urls:
        return ""

    valid_urls = []

    for url in urls:

        url = clean_url(url)

        if is_valid_url(url):
            if url not in valid_urls:
                valid_urls.append(url)

    return "\n".join(valid_urls)


# ============================================================
# PROCESS IMAGE
# ============================================================

def process_image(uploaded_file):

    image = Image.open(
        uploaded_file
    )

    ocr_text = run_ocr(
        image
    )

    qr_codes = extract_qr_codes(
        image
    )

    urls = extract_urls(
        ocr_text
    )

    for qr in qr_codes:

        if qr.lower().startswith(
            (
                "http://",
                "https://",
                "www."
            )
        ):

            qr_url = clean_url(qr)

            if is_valid_url(qr_url):
                if qr_url not in urls:
                    urls.append(qr_url)

    hackathon_name = find_hackathon_name_from_image(
        image,
        ocr_text
    )

    theme = find_theme(
        ocr_text
    )

    last_date = find_last_date(
        ocr_text
    )

    qr_link = get_first_real_link(
        combine_links(qr_codes)
    )

    application_link = get_preferred_application_link(
        combine_links(urls),
        qr_link
    )

    return {

        "Hackathon Name":
            hackathon_name,

        "Theme / Category":
            theme,

        "Who Applied":
            "",

        "Status":
            "",

        "Last Date":
            last_date,

        "Link":
            application_link,

        "QR":
            qr_link,

        "_OCR":
            ocr_text,

        "_QR_LIST":
            qr_codes
    }


# ============================================================
# DATE CONVERSION
# ============================================================

def convert_excel_date(value):

    if value is None:
        return None

    if isinstance(value, datetime):
        return value.date()

    if isinstance(value, date):
        return value

    if isinstance(value, pd.Timestamp):
        return value.date()

    if isinstance(value, str):

        value = value.strip()

        if not value:
            return None

        parsed = parse_date(value)

        if parsed:
            return parsed

    return None


# ============================================================
# WHATSAPP ALERT MESSAGE
# ============================================================

def create_alert_message(
    hackathon_name,
    last_date,
    link,
    days_remaining=None
):

    last_date = convert_excel_date(
        last_date
    )

    if not last_date:
        return (
            "🚨 Hackathon Deadline Alert\n\n"
            f"Hackathon - {hackathon_name}\n"
            "Last Date - Date not available\n"
        )

    date_text = (
        f"{last_date.day:02d}/"
        f"{last_date.month:02d}/"
        f"{last_date.year}"
    )

    first_link = get_first_real_link(
        link
    )

    if not first_link:
        first_link = "No application link found"

    message = (
        "🚨 Hackathon Deadline Alert\n\n"
        f"Hackathon - {hackathon_name}\n"
        f"Last Date - *{date_text}*\n"
    )

    if days_remaining is not None:

        if days_remaining == 0:
            message += "Days Remaining - *Today*\n"

        elif days_remaining == 1:
            message += "Days Remaining - *1 day*\n"

        elif days_remaining > 1:
            message += (
                f"Days Remaining - "
                f"*{days_remaining} days*\n"
            )

    message += (
        f"Apply Here - {first_link}"
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

    last_date = convert_excel_date(
        last_date
    )

    if not last_date:
        last_date = date.today()

    date_text = (
        f"{last_date.day:02d}/"
        f"{last_date.month:02d}/"
        f"{last_date.year}"
    )

    first_link = get_first_real_link(
        link
    )

    if first_link and "example.com" in first_link.lower():
        first_link = ""

    if days_remaining is not None:

        if days_remaining == 0:
            days_text = "Deadline is today."

        elif days_remaining == 1:
            days_text = "Only 1 day remaining."

        elif days_remaining > 1:
            days_text = (
                f"{days_remaining} days remaining."
            )

        else:
            days_text = (
                f"Deadline expired "
                f"{abs(days_remaining)} days ago."
            )

    else:
        days_text = ""

    status_html = ""

    if status:
        status_html = (
            f'<p style="margin:7px 0;">'
            f'<strong>Status:</strong> {status}'
            f'</p>'
        )

    if first_link:

        safe_link = (
            first_link
            .replace("&", "&amp;")
            .replace('"', "%22")
        )

        apply_html = f"""
        <p style="margin:24px 0;">
            <a href="{safe_link}"
               target="_blank"
               style="background:#dc2626;
                      color:#ffffff;
                      padding:12px 22px;
                      text-decoration:none;
                      border-radius:6px;
                      font-weight:bold;
                      display:inline-block;">
                Apply Now
            </a>
        </p>

        <p style="font-size:13px;color:#666;">
            Application Link:
            <a href="{safe_link}" target="_blank">
                {first_link}
            </a>
        </p>
        """

    else:

        apply_html = """
        <p style="color:#b91c1c;
                  font-weight:bold;">
            ⚠️ No real application link was found
            in the extracted data.
        </p>
        """

    days_html = ""

    if days_text:

        days_html = (
            f'<p style="margin:7px 0;'
            f'color:#dc2626;'
            f'font-weight:bold;">'
            f'{days_text}'
            f'</p>'
        )

    html = f"""
    <html>

      <body style="
          margin:0;
          padding:0;
          background:#f4f6f8;
          font-family:Arial,Helvetica,sans-serif;
          color:#222;">

        <div style="
            max-width:600px;
            margin:30px auto;
            background:#ffffff;
            border-radius:10px;
            overflow:hidden;
            box-shadow:0 2px 10px
                       rgba(0,0,0,0.08);">

          <div style="
              background:#dc2626;
              color:#ffffff;
              padding:18px 24px;">

            <h2 style="
                margin:0;
                font-size:21px;">

              🚨 Hackathon Deadline Alert

            </h2>

          </div>

          <div style="padding:24px;">

            <p style="
                font-size:16px;
                margin-top:0;">

              A hackathon deadline is approaching.

            </p>

            <div style="
                background:#f8fafc;
                border-left:4px solid #dc2626;
                padding:16px;
                margin:20px 0;">

              <p style="margin:7px 0;">

                <strong>Hackathon:</strong>
                {hackathon_name}

              </p>

              <p style="margin:7px 0;">

                <strong>Last Date:</strong>

                <span style="
                    color:#dc2626;
                    font-weight:bold;">

                    {date_text}

                </span>

              </p>

              {status_html}

              {days_html}

            </div>

            {apply_html}

            <p style="
                margin-top:24px;
                font-size:14px;
                color:#555;">

              ⚠️ Please complete your application
              before the deadline.

            </p>

            <p style="margin-bottom:0;">

              Best Regards,<br>

              <strong>Hackathon Team</strong>

            </p>

          </div>

        </div>

      </body>

    </html>
    """

    plain_link = (
        first_link
        if first_link
        else "No application link found"
    )

    plain_text = (
        "🚨 Hackathon Deadline Alert\n\n"
        f"Hackathon - {hackathon_name}\n"
        f"Last Date - {date_text}\n"
        f"Days Remaining - "
        f"{days_remaining if days_remaining is not None else 'N/A'}\n"
        f"Apply Here - {plain_link}\n"
    )

    if status:
        plain_text += (
            f"Status - {status}\n"
        )

    plain_text += (
        "\n⚠️ Please complete your application "
        "before the deadline.\n\n"
        "Best Regards,\n"
        "Hackathon Team"
    )

    return html, plain_text


# ============================================================
# EMAIL CONFIGURATION
# ============================================================

EMAIL_SENDER = os.getenv(
    "EMAIL_SENDER",
    ""
)

EMAIL_APP_PASSWORD = os.getenv(
    "EMAIL_APP_PASSWORD",
    ""
)

EMAIL_RECIPIENT = os.getenv(
    "EMAIL_RECIPIENT",
    ""
)

try:

    EMAIL_SENDER = st.secrets.get(
        "EMAIL_SENDER",
        EMAIL_SENDER
    )

    EMAIL_APP_PASSWORD = st.secrets.get(
        "EMAIL_APP_PASSWORD",
        EMAIL_APP_PASSWORD
    )

    EMAIL_RECIPIENT = st.secrets.get(
        "EMAIL_RECIPIENT",
        EMAIL_RECIPIENT
    )

except Exception:
    pass


# ============================================================
# SEND EMAIL
# ============================================================

def send_email_alert(
    message=None,
    recipient=None,
    subject="🚨 Hackathon Deadline Alert",
    hackathon_name=None,
    last_date=None,
    link=None,
    days_remaining=None,
    status=None
):

    sender = str(
        EMAIL_SENDER
    ).strip()

    app_password = str(
        EMAIL_APP_PASSWORD
    ).strip().replace(
        " ",
        ""
    )

    to_email = str(
        recipient or EMAIL_RECIPIENT
    ).strip()

    if not sender or not app_password or not to_email:

        return False, (
            "Email is not configured. "
            "Set EMAIL_SENDER, "
            "EMAIL_APP_PASSWORD and "
            "EMAIL_RECIPIENT."
        )

    try:

        email = EmailMessage()

        email["From"] = sender

        email["To"] = to_email

        email["Subject"] = subject

        # ----------------------------------------------------
        # Always build email from the actual alert data
        # ----------------------------------------------------

        if hackathon_name and last_date:

            html_body, plain_body = (
                create_email_content(
                    hackathon_name,
                    last_date,
                    link,
                    days_remaining,
                    status
                )
            )

        else:

            plain_body = (
                message
                or
                "Hackathon deadline alert."
            )

            html_body = f"""
            <html>
              <body style="
                  font-family:Arial,sans-serif;">

                <pre style="
                    font-family:Arial,sans-serif;
                    white-space:pre-wrap;">

{plain_body}

                </pre>

              </body>
            </html>
            """

        email.set_content(
            plain_body
        )

        email.add_alternative(
            html_body,
            subtype="html"
        )

        with smtplib.SMTP_SSL(
            "smtp.gmail.com",
            465,
            timeout=30
        ) as smtp:

            smtp.login(
                sender,
                app_password
            )

            smtp.send_message(
                email
            )

        return True, (
            f"Email sent successfully "
            f"to {to_email}"
        )

    except smtplib.SMTPAuthenticationError:

        return False, (
            "Gmail authentication failed. "
            "Check EMAIL_SENDER and make sure "
            "EMAIL_APP_PASSWORD is a valid "
            "Google App Password."
        )

    except smtplib.SMTPException as error:

        return False, (
            f"Gmail SMTP error: {error}"
        )

    except Exception as error:

        return False, (
            f"Email sending failed: {error}"
        )


# ============================================================
# EMAIL BUTTON
# ============================================================

def show_email_button(
    alert,
    key
):

    if st.button(
        "📧 Send Email Alert",
        key=key
    ):

        success, result = send_email_alert(

            subject=(
                "🚨 Hackathon Deadline Alert - "
                f"{alert['Hackathon Name']}"
            ),

            hackathon_name=(
                alert["Hackathon Name"]
            ),

            last_date=(
                alert["Last Date"]
            ),

            link=(
                alert["Link"]
            ),

            days_remaining=(
                alert["Days Remaining"]
            ),

            status=(
                alert["Status"]
            )
        )

        st.session_state[
            "email_last_success"
        ] = success

        st.session_state[
            "email_last_message"
        ] = result

        st.session_state[
            "email_last_time"
        ] = datetime.now().strftime(
            "%d-%m-%Y %I:%M:%S %p"
        )

    if (
        st.session_state.get(
            "email_last_success"
        )
        is True
    ):

        st.success(
            "✅ Email sent successfully. "
            + st.session_state.get(
                "email_last_message",
                ""
            )
        )

    elif (
        st.session_state.get(
            "email_last_success"
        )
        is False
    ):

        st.error(
            "❌ Email failed. "
            + st.session_state.get(
                "email_last_message",
                ""
            )
        )


# ============================================================
# DEADLINE ALERT SYSTEM
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

        last_date = convert_excel_date(
            row.get(
                "Last Date"
            )
        )

        link = get_preferred_application_link(
            str(row.get("Link", "")).strip(),
            str(row.get("QR", "")).strip()
        )

        if not hackathon_name:
            continue

        if not last_date:
            continue

        remaining_days = (
            last_date - check_date
        ).days

        if remaining_days < 0:

            status = "Expired ❌"

        elif remaining_days == 0:

            status = "Deadline Today 🔴"

        elif remaining_days == 1:

            status = "1 Day Remaining 🟠"

        elif remaining_days == 2:

            status = "2 Days Remaining 🟡"

        elif remaining_days == 3:

            status = "3 Days Remaining 🔔"

        else:

            status = "Normal"

        # Use the Status value from the extracted/editable Excel data
        # when it is available. If Status is empty, calculate a live
        # deadline status from the remaining days.
        extracted_status = str(
            row.get(
                "Status",
                ""
            )
        ).strip()

        if extracted_status:
            status = extracted_status

        if 0 <= remaining_days <= alert_days:

            message = create_alert_message(
                hackathon_name,
                last_date,
                link,
                remaining_days
            )

            alerts.append({

                "Hackathon Name":
                    hackathon_name,

                "Last Date":
                    last_date,

                "Days Remaining":
                    remaining_days,

                "Status":
                    status,

                "Message":
                    message,

                "Link":
                    link
            })

    return alerts


# ============================================================
# WHATSAPP BUTTON
# ============================================================

def show_whatsapp_button(
    message,
    key
):

    whatsapp_url = (
        "https://wa.me/?text="
        + urllib.parse.quote(
            message
        )
    )

    st.link_button(
        "📱 Open WhatsApp",
        whatsapp_url,
        key=key
    )


# ============================================================
# EXCEL CREATION
# ============================================================

def create_excel(dataframe):

    wb = Workbook()

    ws = wb.active

    ws.title = "Hackathons"

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

        cell = ws.cell(
            row=1,
            column=col_index,
            value=column
        )

        cell.font = Font(
            bold=True
        )

        cell.fill = PatternFill(
            "solid",
            fgColor="D9EAF7"
        )

        cell.alignment = Alignment(
            horizontal="center"
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

            cell = ws.cell(
                row=row_index,
                column=col_index,
                value=value
            )

            if column == "Last Date":

                parsed_date = convert_excel_date(
                    value
                )

                if parsed_date:

                    cell.value = parsed_date

                    cell.number_format = (
                        "dd/mm/yyyy"
                    )

            if column in [
                "Link",
                "QR"
            ]:

                text = str(
                    value
                ).strip()

                if text:

                    first_url = (
                        text.splitlines()[0]
                    )

                    if is_valid_url(
                        first_url
                    ):

                        cell.hyperlink = (
                            first_url
                        )

                        cell.font = Font(
                            color="0563C1",
                            underline="single"
                        )

    widths = {

        "A": 8,
        "B": 35,
        "C": 22,
        "D": 18,
        "E": 18,
        "F": 16,
        "G": 55,
        "H": 55
    }

    for column, width in widths.items():

        ws.column_dimensions[
            column
        ].width = width

    output = BytesIO()

    wb.save(output)

    output.seek(0)

    return output


# ============================================================
# SESSION STATE
# ============================================================

if "email_last_success" not in st.session_state:
    st.session_state[
        "email_last_success"
    ] = None

if "email_last_message" not in st.session_state:
    st.session_state[
        "email_last_message"
    ] = ""

if "email_last_time" not in st.session_state:
    st.session_state[
        "email_last_time"
    ] = ""

if "extracted_hackathon_df" not in st.session_state:
    st.session_state[
        "extracted_hackathon_df"
    ] = None


# ============================================================
# GET CURRENT EXTRACTED DETAILS FOR EMAIL TEST
# ============================================================

def get_current_extracted_email_details():
    """
    Return the first valid extracted hackathon row for the
    email test button. The values come from the current
    extracted/editable dataframe, not hard-coded test values.
    """

    dataframe = st.session_state.get(
        "extracted_hackathon_df"
    )

    if dataframe is None or dataframe.empty:
        return {
            "Hackathon Name": "Hackathon",
            "Last Date": date.today(),
            "Link": "",
            "Status": "No extracted data",
            "Days Remaining": 0
        }

    for _, row in dataframe.iterrows():

        last_date = convert_excel_date(
            row.get("Last Date")
        )

        if not last_date:
            continue

        hackathon_name = str(
            row.get(
                "Hackathon Name",
                "Hackathon"
            )
        ).strip() or "Hackathon"

        link = get_preferred_application_link(
            str(row.get("Link", "")).strip(),
            str(row.get("QR", "")).strip()
        )

        days_remaining = (
            last_date - date.today()
        ).days

        extracted_status = str(
            row.get("Status", "")
        ).strip()

        if extracted_status:
            status = extracted_status
        elif days_remaining < 0:
            status = "Expired ❌"
        elif days_remaining == 0:
            status = "Deadline Today 🔴"
        elif days_remaining == 1:
            status = "1 Day Remaining 🟠"
        elif days_remaining == 2:
            status = "2 Days Remaining 🟡"
        elif days_remaining == 3:
            status = "3 Days Remaining 🔔"
        else:
            status = "Normal"

        return {
            "Hackathon Name": hackathon_name,
            "Last Date": last_date,
            "Link": link,
            "Status": status,
            "Days Remaining": days_remaining
        }

    return {
        "Hackathon Name": "Hackathon",
        "Last Date": date.today(),
        "Link": "",
        "Status": "No valid deadline extracted",
        "Days Remaining": 0
    }


# ============================================================
# EMAIL SETTINGS
# ============================================================

with st.expander(
    "📧 Email Alert Configuration",
    expanded=True
):

    st.write(
        "Configure Gmail SMTP once, test it, "
        "and then use **Send Email Alert** "
        "beside any deadline alert."
    )

    st.caption(
        "For Gmail, use a Google App Password "
        "instead of your normal Gmail password."
    )

    configured = bool(
        EMAIL_SENDER
        and EMAIL_RECIPIENT
        and EMAIL_APP_PASSWORD
    )

    if configured:

        st.success(
            f"Configuration detected: "
            f"📤 {EMAIL_SENDER} → "
            f"📥 {EMAIL_RECIPIENT}"
        )

        if (
            st.session_state.get(
                "email_last_success"
            )
            is True
        ):

            st.success(
                "🟢 EMAIL STATUS: WORKING\n\n"
                f"Last successful send: "
                f"{st.session_state.get('email_last_time', 'N/A')}"
            )

        elif (
            st.session_state.get(
                "email_last_success"
            )
            is False
        ):

            st.error(
                "🔴 EMAIL STATUS: FAILED\n\n"
                f"Last attempt: "
                f"{st.session_state.get('email_last_time', 'N/A')}"
            )

        else:

            st.warning(
                "🟡 EMAIL STATUS: NOT TESTED YET"
            )

        if st.button(
            "📨 Send Test Email",
            key="test_email_button"
        ):

            # IMPORTANT: Use the currently extracted hackathon
            # details instead of hard-coded GENESIS/date values.
            test_details = get_current_extracted_email_details()

            success, result = send_email_alert(

                subject=(
                    "🚨 Hackathon Deadline Alert - "
                    f"{test_details['Hackathon Name']}"
                ),

                hackathon_name=(
                    test_details["Hackathon Name"]
                ),

                last_date=(
                    test_details["Last Date"]
                ),

                link=(
                    test_details["Link"]
                ),

                days_remaining=(
                    test_details["Days Remaining"]
                ),

                status=(
                    test_details["Status"]
                )
            )

            st.session_state[
                "email_last_success"
            ] = success

            st.session_state[
                "email_last_message"
            ] = result

            st.session_state[
                "email_last_time"
            ] = datetime.now().strftime(
                "%d-%m-%Y %I:%M:%S %p"
            )

            if success:

                st.success(
                    "🟢 EMAIL TEST SUCCESSFUL — "
                    "check the recipient inbox."
                )

            else:

                st.error(
                    f"🔴 EMAIL TEST FAILED — "
                    f"{result}"
                )

    else:

        st.error(
            "Email is not configured. Set "
            "EMAIL_SENDER, EMAIL_APP_PASSWORD "
            "and EMAIL_RECIPIENT."
        )


# ============================================================
# IMAGE UPLOAD
# ============================================================

st.subheader(
    "📸 Upload Hackathon Images"
)

uploaded_files = st.file_uploader(

    "Upload one or more images",

    type=[
        "png",
        "jpg",
        "jpeg",
        "webp"
    ],

    accept_multiple_files=True
)


# ============================================================
# EXTRACT BUTTON
# ============================================================

if uploaded_files:

    if st.button(
        "🚀 Extract Hackathon Details",
        type="primary"
    ):

        results = []

        progress = st.progress(0)

        total = len(
            uploaded_files
        )

        for index, uploaded_file in enumerate(
            uploaded_files
        ):

            try:

                result = process_image(
                    uploaded_file
                )

                result["_File"] = (
                    uploaded_file.name
                )

                results.append(
                    result
                )

            except Exception as error:

                st.error(
                    f"Error processing "
                    f"{uploaded_file.name}: "
                    f"{error}"
                )

            progress.progress(
                (index + 1) / total
            )

        if results:

            df = pd.DataFrame(
                results
            )

            df = df.drop_duplicates(
                subset=[
                    "Hackathon Name"
                ],
                keep="first"
            )

            df.insert(
                0,
                "S.No",
                range(
                    1,
                    len(df) + 1
                )
            )

            display_columns = [

                "S.No",
                "Hackathon Name",
                "Theme / Category",
                "Who Applied",
                "Status",
                "Last Date",
                "Link",
                "QR"
            ]

            df = df[
                display_columns
            ]

            st.session_state[
                "extracted_hackathon_df"
            ] = df.copy()

            st.subheader(
                "📊 Extracted Hackathon Data"
            )

            edited_df = st.data_editor(

                df,

                use_container_width=True,

                num_rows="dynamic",

                key="hackathon_editor",

                column_config={

                    "Link":
                        st.column_config.LinkColumn(
                            "Link",
                            help="Open hackathon link",
                            display_text="Open Link"
                        ),

                    "QR":
                        st.column_config.TextColumn(
                            "QR"
                        ),

                    "Last Date":
                        st.column_config.DateColumn(
                            "Last Date",
                            format="DD/MM/YYYY"
                        )
                },

                disabled=[
                    "S.No"
                ]
            )

            # ------------------------------------------------
            # Save edited data
            # ------------------------------------------------

            st.session_state[
                "extracted_hackathon_df"
            ] = edited_df.copy()

            # ------------------------------------------------
            # Deadline Alert
            # ------------------------------------------------

            st.subheader(
                "🚨 Deadline Alert System"
            )

            today = date.today()

            st.write(
                "Checking deadlines from today: "
                f"**{today.strftime('%d/%m/%Y')}**"
            )

            alerts = check_deadline_alerts(
                edited_df,
                today,
                alert_days=3
            )

            if alerts:

                for alert_index, alert in enumerate(
                    alerts
                ):

                    st.warning(

                        f"**{alert['Hackathon Name']}**\n\n"

                        f"Last Date: "
                        f"**{alert['Last Date'].strftime('%d/%m/%Y')}**\n\n"

                        f"Days Remaining: "
                        f"**{alert['Days Remaining']}**\n\n"

                        f"Status: "
                        f"**{alert['Status']}**"
                    )

                    st.code(
                        alert["Message"],
                        language="text"
                    )

                    show_whatsapp_button(
                        alert["Message"],
                        f"whatsapp_{alert_index}"
                    )

                    show_email_button(
                        alert,
                        f"email_{alert_index}"
                    )

            else:

                st.success(
                    "No hackathon deadlines are "
                    "within the last 3 days."
                )

            # ------------------------------------------------
            # Excel
            # ------------------------------------------------

            st.subheader(
                "📥 Download Excel"
            )

            excel_file = create_excel(
                edited_df
            )

            st.download_button(

                label="📊 Download Hackathon Excel",

                data=excel_file,

                file_name=(
                    "hackathon_details.xlsx"
                ),

                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "spreadsheetml.sheet"
                )
            )

            # ------------------------------------------------
            # OCR / QR
            # ------------------------------------------------

            st.subheader(
                "🔍 OCR / QR Details"
            )

            for result in results:

                with st.expander(
                    f"📄 {result.get('_File', 'Image')}"
                ):

                    st.write(
                        "### OCR Text"
                    )

                    st.text(
                        result.get(
                            "_OCR",
                            ""
                        )
                    )

                    st.write(
                        "### QR Codes"
                    )

                    qr_list = result.get(
                        "_QR_LIST",
                        []
                    )

                    if qr_list:

                        for qr in qr_list:
                            st.code(qr)

                    else:

                        st.info(
                            "No QR code detected "
                            "or QR could not be decoded."
                        )


# ============================================================
# TEST ALERT SYSTEM
# ============================================================

st.divider()

st.header(
    "🧪 Test Deadline Alert System"
)

st.write(
    "Test the alert system using the same "
    "deadline extracted from your uploaded image."
)


# ============================================================
# GET EXTRACTED DATA
# ============================================================

extracted_df = st.session_state.get(
    "extracted_hackathon_df"
)


# ============================================================
# DEFAULT VALUES
# ============================================================

extracted_deadline = None

extracted_hackathon_name = (
    "Hackathon"
)

extracted_link = ""


# ============================================================
# GET DATA FROM EXCEL / EXTRACTED DATA
# ============================================================

if (
    extracted_df is not None
    and not extracted_df.empty
):

    # --------------------------------------------------------
    # First valid row with a date
    # --------------------------------------------------------

    for _, row in extracted_df.iterrows():

        possible_date = convert_excel_date(
            row.get(
                "Last Date"
            )
        )

        if possible_date:

            extracted_deadline = (
                possible_date
            )

            extracted_hackathon_name = str(
                row.get(
                    "Hackathon Name",
                    "Unknown Hackathon"
                )
            ).strip()

            extracted_link = get_preferred_application_link(
                str(row.get("Link", "")).strip(),
                str(row.get("QR", "")).strip()
            )

            break


# ============================================================
# FALLBACK TEST DATA
# ============================================================
#
# IMPORTANT:
# There is NO example.com application link now.
#
# If there is no uploaded image, only the deadline and
# hackathon name are used for testing.
# ============================================================

if extracted_deadline is None:

    extracted_deadline = date(
        2026,
        9,
        30
    )

    extracted_hackathon_name = (
        "GENESIS Investment 2.0"
    )

    extracted_link = ""


# ============================================================
# CALCULATE ALERT DATE
# ============================================================

first_alert_date = (
    extracted_deadline
    - timedelta(days=3)
)


# ============================================================
# CALCULATE ACTUAL TODAY REMAINING DAYS
# ============================================================

actual_today = date.today()

actual_remaining_days = (
    extracted_deadline
    - actual_today
).days


# ============================================================
# SHOW DEADLINE INFORMATION
# ============================================================

st.info(

    f"📅 **Deadline:** "
    f"**{extracted_deadline.strftime('%d/%m/%Y')}**\n\n"

    f"📆 **Today:** "
    f"**{actual_today.strftime('%d/%m/%Y')}**\n\n"

    f"⏳ **Actual Days Remaining:** "
    f"**{actual_remaining_days}**\n\n"

    f"🔔 **3-Day Alert Starts:** "
    f"**{first_alert_date.strftime('%d/%m/%Y')}**"
)


# ============================================================
# ACTUAL CURRENT STATUS
# ============================================================

if actual_remaining_days > 3:

    st.success(
        f"ℹ️ No alert yet.\n\n"
        f"Deadline: "
        f"**{extracted_deadline.strftime('%d/%m/%Y')}**\n\n"
        f"Days Remaining: "
        f"**{actual_remaining_days}**\n\n"
        f"3-day alert starts on: "
        f"**{first_alert_date.strftime('%d/%m/%Y')}**"
    )

elif actual_remaining_days == 3:

    st.warning(
        "🔔 **3 Days Remaining — Alert period started.**"
    )

elif actual_remaining_days == 2:

    st.warning(
        "🟡 **2 Days Remaining**"
    )

elif actual_remaining_days == 1:

    st.warning(
        "🟠 **1 Day Remaining**"
    )

elif actual_remaining_days == 0:

    st.error(
        "🔴 **Deadline is TODAY.**"
    )

else:

    st.error(
        f"❌ Deadline expired "
        f"{abs(actual_remaining_days)} days ago."
    )


# ============================================================
# SELECT TEST DATE
# ============================================================

st.subheader(
    "🧪 Manual Alert Testing"
)

test_date = st.date_input(

    "Select test date",

    value=first_alert_date,

    key="test_date"
)


# ============================================================
# SHOW TEST DATE CALCULATION
# ============================================================

test_remaining_days = (
    extracted_deadline
    - test_date
).days

st.write(
    f"Test Date: "
    f"**{test_date.strftime('%d/%m/%Y')}**"
)

st.write(
    f"Test Days Remaining: "
    f"**{test_remaining_days}**"
)


# ============================================================
# RUN ALERT TEST
# ============================================================

if st.button(
    "🧪 Run Alert Test"
):

    test_df = pd.DataFrame([

        {

            "S.No":
                1,

            "Hackathon Name":
                extracted_hackathon_name,

            "Theme / Category":
                "Technology",

            "Who Applied":
                "",

            "Status":
                "",

            "Last Date":
                extracted_deadline,

            "Link":
                extracted_link,

            "QR":
                extracted_link
        }
    ])

    st.subheader(
        "📋 Test Data"
    )

    st.data_editor(

        test_df,

        use_container_width=True,

        key="test_hackathon_editor",

        column_config={

            "Last Date":
                st.column_config.DateColumn(
                    "Last Date",
                    format="DD/MM/YYYY"
                ),

            "Link":
                st.column_config.LinkColumn(
                    "Link",
                    display_text="Open Link"
                )
        },

        disabled=[
            "S.No"
        ]
    )

    # --------------------------------------------------------
    # Alert calculation
    # --------------------------------------------------------

    test_alerts = check_deadline_alerts(
        test_df,
        test_date,
        alert_days=3
    )

    st.subheader(
        "🚨 Test Result"
    )

    if test_alerts:

        for alert_index, alert in enumerate(
            test_alerts
        ):

            st.warning(

                f"**{alert['Hackathon Name']}**\n\n"

                f"Last Date: "
                f"**{alert['Last Date'].strftime('%d/%m/%Y')}**\n\n"

                f"Days Remaining: "
                f"**{alert['Days Remaining']}**\n\n"

                f"Status: "
                f"**{alert['Status']}**"
            )

            st.write(
                "### WhatsApp Message"
            )

            st.code(
                alert["Message"],
                language="text"
            )

            show_whatsapp_button(
                alert["Message"],
                f"test_whatsapp_{alert_index}"
            )

            show_email_button(
                alert,
                f"test_email_{alert_index}"
            )

            st.success(
                "✅ Alert system is working."
            )

            st.info(
                "The WhatsApp button opens WhatsApp "
                "with the message prepared. You must "
                "select the group/contact and press "
                "Send manually."
            )

    else:

        if test_remaining_days > 3:

            st.info(

                f"ℹ️ No alert yet.\n\n"

                f"Deadline: "
                f"**{extracted_deadline.strftime('%d/%m/%Y')}**\n\n"

                f"Days Remaining: "
                f"**{test_remaining_days}**\n\n"

                f"3-day alert starts on: "
                f"**{first_alert_date.strftime('%d/%m/%Y')}**"
            )

        elif test_remaining_days < 0:

            st.error(

                f"❌ Deadline expired on "
                f"**{extracted_deadline.strftime('%d/%m/%Y')}**."
            )

        else:

            st.info(
                "No alert generated for this test date."
            )


# ============================================================
# TESTING INFORMATION
# ============================================================

with st.expander(
    "ℹ️ How to test the last 3 days"
):

    date_4_days = (
        extracted_deadline
        - timedelta(days=4)
    )

    date_3_days = (
        extracted_deadline
        - timedelta(days=3)
    )

    date_2_days = (
        extracted_deadline
        - timedelta(days=2)
    )

    date_1_day = (
        extracted_deadline
        - timedelta(days=1)
    )

    date_expired = (
        extracted_deadline
        + timedelta(days=1)
    )

    st.write(

        f"""
The system uses the **same deadline extracted
from the uploaded image**.

### Hackathon

**{extracted_hackathon_name}**

### Current Deadline

**{extracted_deadline.strftime('%d/%m/%Y')}**

### Current Actual Days Remaining

**{actual_remaining_days}**

### Alert Schedule

- **{date_4_days.strftime('%d/%m/%Y')}** → 4 Days Remaining → No Alert
- **{date_3_days.strftime('%d/%m/%Y')}** → 3 Days Remaining 🔔
- **{date_2_days.strftime('%d/%m/%Y')}** → 2 Days Remaining 🟡
- **{date_1_day.strftime('%d/%m/%Y')}** → 1 Day Remaining 🟠
- **{extracted_deadline.strftime('%d/%m/%Y')}** → Deadline Today 🔴
- **{date_expired.strftime('%d/%m/%Y')}** → Expired ❌
"""
    )