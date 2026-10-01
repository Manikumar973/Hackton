
# ============================================================
# HACKATHON IMAGE-TO-EXCEL EXTRACTOR
# OCR • QR • Excel • Deadline Alerts
#
# Run:
#     streamlit run app.py
# ============================================================

import re
import smtplib
from email.message import EmailMessage
from io import BytesIO
from datetime import date

import streamlit as st
import pandas as pd


# ============================================================
# OPTIONAL IMPORTS
# ============================================================

try:
    import cv2
    import numpy as np
    import pytesseract

    CV2_OK = True
except Exception:
    CV2_OK = False


try:
    from PIL import Image

    PIL_OK = True
except Exception:
    PIL_OK = False


try:
    from pyzbar import pyzbar

    PYZBAR_OK = True
except Exception:
    PYZBAR_OK = False


try:
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment

    OPENPYXL_OK = True
except Exception:
    OPENPYXL_OK = False


# ============================================================
# PAGE CONFIG
# ============================================================

st.set_page_config(
    page_title="Hackathon Image-to-Excel Extractor",
    page_icon="🏆",
    layout="wide",
)


# ============================================================
# PAGE CSS
# ============================================================

st.markdown(
    """
    <style>

    .hero {
        background: linear-gradient(135deg, #1e3a8a 0%, #3b82f6 100%);
        border-radius: 16px;
        padding: 32px;
        color: white;
        margin-bottom: 24px;
    }

    .hero-badge {
        display: inline-block;
        background: rgba(255,255,255,0.2);
        padding: 6px 14px;
        border-radius: 999px;
        font-size: 0.85rem;
        margin-bottom: 14px;
    }

    .hero-title {
        font-size: 2.1rem;
        font-weight: 800;
        margin-bottom: 10px;
    }

    .hero-subtitle {
        font-size: 1rem;
        opacity: 0.9;
        max-width: 700px;
    }

    .section-title {
        font-size: 1.4rem;
        font-weight: 700;
        margin: 28px 0 12px 0;
    }

    .card {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 16px;
    }

    .hackathon-name {
        font-size: 1.35rem;
        font-weight: 800;
        margin-top: 8px;
        margin-bottom: 4px;
    }

    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# SESSION STATE
# ============================================================

for key, default in [
    ("results", []),
    ("email_sender_config", "shankerbidalert@gmail.com"),
    ("email_app_password_config", ""),
    ("email_recipient_config", "manikumardasari007@gmail.com"),
    ("email_last_success", False),
    ("email_last_message", ""),
]:
    if key not in st.session_state:
        st.session_state[key] = default


# ============================================================
# HELPER FUNCTIONS
# ============================================================

def clean_text(text):
    """Clean OCR text."""

    if not text:
        return ""

    text = text.replace("\x00", " ")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n+", "\n", text)

    return text.strip()


# ============================================================
# DEADLINE PARSER
# ============================================================

def parse_deadline(text):
    """
    Find the first plausible deadline date in OCR text.
    Returns date or None.
    """

    if not text:
        return None

    months = {
        "jan": 1,
        "feb": 2,
        "mar": 3,
        "apr": 4,
        "may": 5,
        "jun": 6,
        "jul": 7,
        "aug": 8,
        "sep": 9,
        "oct": 10,
        "nov": 11,
        "dec": 12,
    }

    patterns = [
        # 25 September 2026
        r"(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]{3,9})\s*,?\s*(\d{4})?",

        # 2026-09-25
        r"(\d{4})-(\d{1,2})-(\d{1,2})",

        # 25/09/2026 or 25-09-2026
        r"(\d{1,2})[/\-](\d{1,2})[/\-](\d{2,4})",
    ]

    for pattern_index, pattern in enumerate(patterns):

        for match in re.finditer(
            pattern,
            text,
            re.IGNORECASE,
        ):

            groups = match.groups()

            try:

                # --------------------------------------------
                # 25 September 2026
                # --------------------------------------------

                if pattern_index == 0:

                    day = int(groups[0])

                    month = months.get(
                        groups[1].lower()[:3]
                    )

                    year = (
                        int(groups[2])
                        if groups[2]
                        else date.today().year
                    )

                    if month is None:
                        continue

                    return date(
                        year,
                        month,
                        day,
                    )

                # --------------------------------------------
                # 2026-09-25
                # --------------------------------------------

                elif pattern_index == 1:

                    return date(
                        int(groups[0]),
                        int(groups[1]),
                        int(groups[2]),
                    )

                # --------------------------------------------
                # 25/09/2026
                # --------------------------------------------

                else:

                    first = int(groups[0])
                    second = int(groups[1])
                    year = int(groups[2])

                    if year < 100:
                        year += 2000

                    # DD/MM/YYYY
                    if 1 <= second <= 12 and 1 <= first <= 31:

                        try:
                            return date(
                                year,
                                second,
                                first,
                            )
                        except ValueError:
                            pass

                    # MM/DD/YYYY
                    if 1 <= first <= 12 and 1 <= second <= 31:

                        try:
                            return date(
                                year,
                                first,
                                second,
                            )
                        except ValueError:
                            pass

            except (ValueError, TypeError):
                continue

    return None


# ============================================================
# QR EXTRACTION
# ============================================================

def extract_qr_codes(image_bytes):
    """
    Return list of QR data strings found in image.
    """

    codes = []

    if not PIL_OK:
        return codes

    try:
        img = Image.open(
            BytesIO(image_bytes)
        )
    except Exception:
        return codes

    # --------------------------------------------------------
    # First method: pyzbar
    # --------------------------------------------------------

    if PYZBAR_OK:

        try:

            decoded = pyzbar.decode(img)

            for item in decoded:

                data = item.data.decode(
                    "utf-8",
                    errors="ignore",
                ).strip()

                if data and data not in codes:
                    codes.append(data)

        except Exception:
            pass

    # --------------------------------------------------------
    # Second method: OpenCV
    # --------------------------------------------------------

    if not codes and CV2_OK:

        try:

            arr = np.array(
                img.convert("RGB")
            )

            arr = cv2.cvtColor(
                arr,
                cv2.COLOR_RGB2BGR,
            )

            detector = cv2.QRCodeDetector()

            data, _, _ = detector.detectAndDecode(
                arr
            )

            if data:

                data = data.strip()

                if data:
                    codes.append(data)

        except Exception:
            pass

    return codes


# ============================================================
# OCR
# ============================================================

def extract_ocr(image_bytes):
    """
    Run OCR using multiple preprocessing methods.
    """

    if not (PIL_OK and CV2_OK):
        return ""

    try:

        img = Image.open(
            BytesIO(image_bytes)
        ).convert("RGB")

        arr = np.array(img)

        arr = cv2.cvtColor(
            arr,
            cv2.COLOR_RGB2BGR,
        )

        # Resize image
        arr = cv2.resize(
            arr,
            None,
            fx=1.7,
            fy=1.7,
            interpolation=cv2.INTER_CUBIC,
        )

        gray = cv2.cvtColor(
            arr,
            cv2.COLOR_BGR2GRAY,
        )

        # Slight denoise
        gray = cv2.GaussianBlur(
            gray,
            (3, 3),
            0,
        )

        # Threshold
        threshold = cv2.threshold(
            gray,
            0,
            255,
            cv2.THRESH_BINARY + cv2.THRESH_OTSU,
        )[1]

        texts = []

        # ----------------------------------------------------
        # OCR original grayscale
        # ----------------------------------------------------

        try:

            text1 = pytesseract.image_to_string(
                gray,
                config="--psm 6",
            )

            if text1:
                texts.append(text1)

        except Exception:
            pass

        # ----------------------------------------------------
        # OCR threshold image
        # ----------------------------------------------------

        try:

            text2 = pytesseract.image_to_string(
                threshold,
                config="--psm 6",
            )

            if text2:
                texts.append(text2)

        except Exception:
            pass

        # ----------------------------------------------------
        # OCR sparse text
        # ----------------------------------------------------

        try:

            text3 = pytesseract.image_to_string(
                threshold,
                config="--psm 11",
            )

            if text3:
                texts.append(text3)

        except Exception:
            pass

        # Combine OCR results
        combined = "\n".join(texts)

        return clean_text(combined)

    except Exception:
        return ""


# ============================================================
# URL CLEANING
# ============================================================

def clean_url(url):
    """
    Clean URL extracted from OCR/QR.
    """

    if not url:
        return ""

    url = str(url).strip()

    url = url.rstrip(
        ".,;:)]}>\"'"
    )

    # Add https if www.
    if url.startswith("www."):
        url = "https://" + url

    return url


# ============================================================
# FIND APPLICATION URL
# ============================================================

def extract_application_link(text, qr_codes):
    """
    Prefer visible website URL.
    Otherwise use QR URL.
    """

    if text:

        url_matches = re.findall(
            r"(https?://[^\s<>\"]+|www\.[^\s<>\"]+)",
            text,
            flags=re.IGNORECASE,
        )

        if url_matches:

            for url in url_matches:

                cleaned = clean_url(url)

                if cleaned:
                    return cleaned

    # QR fallback
    for qr in qr_codes:

        qr_clean = clean_url(qr)

        if qr_clean.startswith(
            ("http://", "https://")
        ):
            return qr_clean

    return ""


# ============================================================
# HACKATHON NAME EXTRACTION
# ============================================================

def extract_hackathon_name(text, source_name):
    """
    Try multiple methods to identify the hackathon name.
    Always returns a non-empty value when possible.
    """

    lines = [
        re.sub(
            r"\s+",
            " ",
            line,
        ).strip()
        for line in text.splitlines()
        if line.strip()
    ]

    ignored_words = [
        "register",
        "registration",
        "apply now",
        "scan",
        "deadline",
        "last date",
        "theme",
        "organized by",
        "organised by",
        "visit",
        "http://",
        "https://",
        "www.",
        "contact",
        "email",
        "phone",
        "date",
        "time",
        "venue",
        "prize",
    ]

    # --------------------------------------------------------
    # METHOD 1
    # Search lines containing HACKATHON
    # --------------------------------------------------------

    for line in lines:

        lower = line.lower()

        if "hackathon" in lower:

            cleaned = re.sub(
                r"^[^A-Za-z0-9]+",
                "",
                line,
            ).strip()

            if len(cleaned) >= 4:
                return cleaned.upper()

    # --------------------------------------------------------
    # METHOD 2
    # Look for strong title-like lines
    # --------------------------------------------------------

    candidates = []

    for line in lines:

        lower = line.lower()

        if len(line) < 4:
            continue

        if len(line) > 100:
            continue

        if any(
            word in lower
            for word in ignored_words
        ):
            continue

        alpha_count = sum(
            char.isalpha()
            for char in line
        )

        if alpha_count < 4:
            continue

        candidates.append(line)

    if candidates:
        return candidates[0].upper()

    # --------------------------------------------------------
    # METHOD 3
    # Filename fallback
    # --------------------------------------------------------

    filename = source_name

    filename = re.sub(
        r"\.(png|jpg|jpeg|webp)$",
        "",
        filename,
        flags=re.IGNORECASE,
    )

    filename = re.sub(
        r"[_\-]+",
        " ",
        filename,
    )

    filename = filename.strip()

    if filename:
        return filename.upper()

    return "HACKATHON"


# ============================================================
# PARSE POSTER
# ============================================================

def parse_poster(
    text,
    qr_codes,
    source_name,
):

    text = clean_text(text)

    lines = [
        line.strip()
        for line in text.splitlines()
        if line.strip()
    ]

    # Hackathon name
    name = extract_hackathon_name(
        text,
        source_name,
    )

    # Theme
    theme = ""

    for line in lines:

        if "theme" in line.lower():

            theme = line

            theme = re.sub(
                r"^\s*theme\s*[:\-]?\s*",
                "",
                theme,
                flags=re.IGNORECASE,
            )

            theme = theme.strip()

            if theme:
                break

    # Deadline
    deadline = parse_deadline(text)

    # Application link
    link = extract_application_link(
        text,
        qr_codes,
    )

    # QR value
    qr_value = ""

    if qr_codes:

        # Prefer QR URL
        for qr in qr_codes:

            cleaned_qr = clean_url(qr)

            if cleaned_qr.startswith(
                ("http://", "https://")
            ):

                qr_value = cleaned_qr
                break

        # Otherwise keep first QR text
        if not qr_value:
            qr_value = qr_codes[0]

    return {
        "Hackathon Name": name,
        "Theme": theme,
        "Deadline": (
            deadline.strftime("%d-%m-%Y")
            if deadline
            else ""
        ),
        "_deadline_date": deadline,
        "Application Link": link,
        "QR Codes": qr_value,
        "Source Image": source_name,
        "Full OCR Text": text[:2000],
    }


# ============================================================
# DEADLINE ALERT CHECK
# ============================================================

def check_deadline_alerts(
    df,
    today,
    alert_days=3,
):

    alerts = []

    if (
        df is None
        or df.empty
        or "_deadline_date" not in df.columns
    ):
        return alerts

    for idx, row in df.iterrows():

        d = row.get("_deadline_date")

        if isinstance(d, date):

            days_left = (
                d - today
            ).days

            if 0 <= days_left <= alert_days:

                alerts.append(
                    (
                        idx,
                        d,
                        days_left,
                    )
                )

    return alerts


# ============================================================
# SEND EMAIL
# ============================================================

def send_email(
    sender,
    app_password,
    recipient,
    subject,
    body,
):

    msg = EmailMessage()

    msg["From"] = sender
    msg["To"] = recipient
    msg["Subject"] = subject

    msg.set_content(body)

    with smtplib.SMTP_SSL(
        "smtp.gmail.com",
        465,
    ) as server:

        server.login(
            sender,
            app_password,
        )

        server.send_message(msg)

    return True


# ============================================================
# PAGE HEADER
# ============================================================

st.markdown(
    """
    <div class="hero">

        <div class="hero-badge">
            🏆 OCR • QR • Excel • Deadline Alerts
        </div>

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
    unsafe_allow_html=True,
)


# ============================================================
# UPLOAD
# ============================================================

st.markdown(
    '<div class="section-title">📤 Upload Posters</div>',
    unsafe_allow_html=True,
)

uploaded_files = st.file_uploader(
    "Choose one or more hackathon poster images",
    type=[
        "png",
        "jpg",
        "jpeg",
        "webp",
    ],
    accept_multiple_files=True,
)


if uploaded_files:

    if not CV2_OK:

        st.error(
            "OpenCV / pytesseract is not installed. "
            "Run: pip install opencv-python pytesseract"
        )

    else:

        progress = st.progress(0)

        for i, file in enumerate(uploaded_files):

            progress.progress(
                (i + 1) / len(uploaded_files),
                text=f"Processing {file.name}...",
            )

            image_bytes = file.getvalue()

            # QR
            qr_codes = extract_qr_codes(
                image_bytes
            )

            # OCR
            text = extract_ocr(
                image_bytes
            )

            # Parse
            result = parse_poster(
                text,
                qr_codes,
                file.name,
            )

            # Avoid exact duplicate image
            existing_sources = [
                item.get("Source Image")
                for item in st.session_state.results
            ]

            if file.name not in existing_sources:

                st.session_state.results.append(
                    result
                )

        progress.empty()

        st.success(
            f"Processed {len(uploaded_files)} image(s)."
        )


# ============================================================
# EXTRACTED DATA
# ============================================================

st.markdown(
    '<div class="section-title">📋 Extracted Data (editable)</div>',
    unsafe_allow_html=True,
)

edited = pd.DataFrame()


if st.session_state.results:

    df = pd.DataFrame(
        st.session_state.results
    )

    # Make sure required columns exist
    required_columns = [
        "Hackathon Name",
        "Theme",
        "Deadline",
        "Application Link",
        "QR Codes",
        "Source Image",
        "Full OCR Text",
    ]

    for column in required_columns:

        if column not in df.columns:
            df[column] = ""

    # Internal deadline column
    if "_deadline_date" not in df.columns:

        df["_deadline_date"] = [
            parse_deadline(value)
            for value in df["Deadline"]
        ]

    # --------------------------------------------------------
    # Force hackathon names uppercase
    # --------------------------------------------------------

    df["Hackathon Name"] = (
        df["Hackathon Name"]
        .fillna("")
        .astype(str)
        .str.strip()
        .str.upper()
    )

    # --------------------------------------------------------
    # Display columns
    # --------------------------------------------------------

    display_columns = [
        "Hackathon Name",
        "Theme",
        "Deadline",
        "Application Link",
        "QR Codes",
        "Source Image",
        "Full OCR Text",
    ]

    display_df = df[
        display_columns
    ].copy()

    # --------------------------------------------------------
    # Editable table
    #
    # IMPORTANT:
    # width is an integer for compatibility with
    # older Streamlit versions.
    # --------------------------------------------------------

    edited = st.data_editor(
        display_df,
        width=1000,
        height=450,
        key="edited_table",
        num_rows="dynamic",

        column_config={

            "Hackathon Name": st.column_config.TextColumn(
                "HACKATHON NAME",
                width="large",
                required=True,
            ),

            "Theme": st.column_config.TextColumn(
                "THEME",
                width="medium",
            ),

            "Deadline": st.column_config.TextColumn(
                "DEADLINE",
                width="medium",
            ),

            "Application Link": st.column_config.LinkColumn(
                "APPLICATION LINK",
                width="large",
                help="Click the link to open the application website.",
                validate=r"^https?://.*$",
            ),

            "QR Codes": st.column_config.LinkColumn(
                "QR CODE / URL",
                width="large",
                help="Click the QR URL to open it.",
                validate=r"^https?://.*$",
            ),

            "Source Image": st.column_config.TextColumn(
                "SOURCE IMAGE",
                width="medium",
            ),

            "Full OCR Text": st.column_config.TextColumn(
                "FULL OCR TEXT",
                width="large",
            ),
        },
    )

    # --------------------------------------------------------
    # Rebuild internal deadline dates after editing
    # --------------------------------------------------------

    edited = edited.copy()

    edited["_deadline_date"] = [
        parse_deadline(value)
        for value in edited.get(
            "Deadline",
            [""] * len(edited),
        )
    ]

    # --------------------------------------------------------
    # Force names uppercase after editing
    # --------------------------------------------------------

    if "Hackathon Name" in edited.columns:

        edited["Hackathon Name"] = (
            edited["Hackathon Name"]
            .fillna("")
            .astype(str)
            .str.strip()
            .str.upper()
        )

    # --------------------------------------------------------
    # SAVE TABLE CHANGES BACK TO SESSION
    # --------------------------------------------------------

    st.session_state.results = (
        edited.to_dict("records")
    )

    # --------------------------------------------------------
    # SHOW HACKATHON NAMES
    # --------------------------------------------------------

    st.markdown(
        "### 🏆 Detected Hackathon Names"
    )

    for _, row in edited.iterrows():

        hackathon_name = str(
            row.get(
                "Hackathon Name",
                "",
            )
        ).strip()

        if hackathon_name:

            st.markdown(
                f"""
                <div class="hackathon-name">
                    🏆 {hackathon_name}
                </div>
                """,
                unsafe_allow_html=True,
            )

else:

    st.info(
        "Upload posters above to see extracted data here."
    )


# ============================================================
# EXCEL EXPORT
# ============================================================

if edited is not None and not edited.empty:

    st.markdown(
        '<div class="section-title">📥 Download Excel</div>',
        unsafe_allow_html=True,
    )

    if OPENPYXL_OK:

        wb = Workbook()

        ws = wb.active

        ws.title = "Hackathons"

        out_df = edited.drop(
            columns=["_deadline_date"],
            errors="ignore",
        )

        headers = list(
            out_df.columns
        )

        # ----------------------------------------------------
        # Header style
        # ----------------------------------------------------

        header_fill = PatternFill(
            start_color="1E3A8A",
            end_color="1E3A8A",
            fill_type="solid",
        )

        header_font = Font(
            bold=True,
            color="FFFFFF",
        )

        for col_index, header in enumerate(
            headers,
            start=1,
        ):

            cell = ws.cell(
                row=1,
                column=col_index,
                value=header,
            )

            cell.fill = header_fill
            cell.font = header_font

            cell.alignment = Alignment(
                horizontal="center"
            )

        # ----------------------------------------------------
        # Data
        # ----------------------------------------------------

        for row_index, row in enumerate(
            out_df.itertuples(index=False),
            start=2,
        ):

            for col_index, value in enumerate(
                row,
                start=1,
            ):

                cell = ws.cell(
                    row=row_index,
                    column=col_index,
                    value=(
                        ""
                        if pd.isna(value)
                        else str(value)
                    ),
                )

                header_name = headers[
                    col_index - 1
                ]

                # Make URL columns clickable
                if header_name in [
                    "Application Link",
                    "QR Codes",
                ]:

                    url = str(
                        value
                    ).strip()

                    if url.startswith(
                        ("http://", "https://")
                    ):

                        cell.hyperlink = url
                        cell.style = "Hyperlink"

        # ----------------------------------------------------
        # Column widths
        # ----------------------------------------------------

        for col_index, header in enumerate(
            headers,
            start=1,
        ):

            max_length = len(
                str(header)
            )

            for row_cells in ws.iter_rows(
                min_col=col_index,
                max_col=col_index,
            ):

                for item in row_cells:

                    if item.value:

                        max_length = max(
                            max_length,
                            len(str(item.value)),
                        )

            # Convert column number to Excel letters
            def excel_column_name(number):
                result = ""

                while number:
                    number, remainder = divmod(
                        number - 1,
                        26,
                    )

                    result = (
                        chr(65 + remainder)
                        + result
                    )

                return result

            column_letter = excel_column_name(
                col_index
            )

            ws.column_dimensions[
                column_letter
            ].width = min(
                max(max_length + 2, 15),
                45,
            )

        # ----------------------------------------------------
        # Save Excel
        # ----------------------------------------------------

        buf = BytesIO()

        wb.save(buf)

        buf.seek(0)

        st.download_button(
            "⬇️ Download Hackathons.xlsx",
            data=buf,
            file_name="Hackathons.xlsx",
            mime=(
                "application/"
                "vnd.openxmlformats-officedocument."
                "spreadsheetml.sheet"
            ),
            width=1000,
        )

    else:

        st.download_button(
            "⬇️ Download CSV",
            data=(
                edited
                .drop(
                    columns=["_deadline_date"],
                    errors="ignore",
                )
                .to_csv(index=False)
                .encode()
            ),
            file_name="Hackathons.csv",
            mime="text/csv",
        )


# ============================================================
# DEADLINE ALERTS
# ============================================================

st.markdown(
    '<div class="section-title">🚨 Deadline Alerts</div>',
    unsafe_allow_html=True,
)


if edited is not None and not edited.empty:

    alerts = check_deadline_alerts(
        edited,
        date.today(),
        alert_days=3,
    )

    if alerts:

        for idx, deadline, days_left in alerts:

            name = (
                edited.loc[
                    idx,
                    "Hackathon Name",
                ]
                if "Hackathon Name"
                in edited.columns
                else "Unknown"
            )

            if days_left == 0:

                st.error(
                    f"⏰ {name} — "
                    f"deadline is TODAY "
                    f"({deadline.strftime('%d-%m-%Y')})!"
                )

            else:

                st.warning(
                    f"⏰ {name} — "
                    f"deadline in {days_left} day(s) "
                    f"({deadline.strftime('%d-%m-%Y')})"
                )

    else:

        st.success(
            "✅ No deadlines in the next 3 days."
        )

else:

    st.info(
        "No data yet — upload posters to check deadlines."
    )


# ============================================================
# EMAIL ALERTS
# ============================================================

st.markdown(
    '<div class="section-title">📧 Email Alerts</div>',
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="card">

    Configure Gmail SMTP to receive hackathon deadline alerts.

    <br>

    <small>
    Use a Gmail <b>App Password</b>
    (Google Account → Security → 2-Step Verification
    → App passwords), not your normal Gmail password.
    </small>

    </div>
    """,
    unsafe_allow_html=True,
)


ec1, ec2, ec3 = st.columns(3)


with ec1:

    email_sender = st.text_input(
        "Sender Gmail",
        value=st.session_state.email_sender_config,
        key="email_sender_config",
    )


with ec2:

    email_app_password = st.text_input(
        "Gmail App Password",
        type="password",
        value=st.session_state.email_app_password_config,
        key="email_app_password_config",
        placeholder="Gmail App Password",
    )


with ec3:

    email_recipient = st.text_input(
        "Recipient Email",
        value=st.session_state.email_recipient_config,
        key="email_recipient_config",
    )


# ============================================================
# SEND EMAIL BUTTON
# ============================================================

if st.button(
    "📤 Send Deadline Alert Email",
    type="primary",
):

    if (
        not email_sender
        or not email_app_password
        or not email_recipient
    ):

        st.error(
            "Fill in all three email fields first."
        )

    else:

        try:

            lines = []

            if (
                edited is not None
                and not edited.empty
            ):

                for (
                    idx,
                    deadline,
                    days_left,
                ) in check_deadline_alerts(
                    edited,
                    date.today(),
                    alert_days=3,
                ):

                    name = (
                        edited.loc[
                            idx,
                            "Hackathon Name",
                        ]
                        if "Hackathon Name"
                        in edited.columns
                        else "Unknown"
                    )

                    lines.append(
                        f"• {name} — "
                        f"deadline "
                        f"{deadline.strftime('%d-%m-%Y')} "
                        f"({days_left} day(s) left)"
                    )

            if not lines:

                lines.append(
                    "No deadlines in the next 3 days."
                )

            body = (
                "Hackathon Deadline Alert\n\n"
                + "\n".join(lines)
                + f"\n\nSent on {date.today()}"
            )

            send_email(
                email_sender,
                email_app_password,
                email_recipient,
                "⏰ Hackathon Deadline Alert",
                body,
            )

            st.session_state.email_last_success = True

            st.session_state.email_last_message = (
                "Email sent successfully!"
            )

            st.success(
                "✅ Email sent successfully!"
            )

        except Exception as e:

            st.session_state.email_last_success = False

            st.session_state.email_last_message = str(e)

            st.error(
                f"❌ Email failed: {e}"
            )

            st.caption(
                "Hint: make sure 2-Step Verification "
                "is ON in your Gmail account and you "
                "are using an App Password."
            )
