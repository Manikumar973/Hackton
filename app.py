# ============================================================
# HACKATHON IMAGE-TO-EXCEL EXTRACTOR
# OCR • QR • Excel • Deadline Alerts
# Run:  streamlit run app.py
# ============================================================

import re
import smtplib
from email.message import EmailMessage
from io import BytesIO
from datetime import date, datetime, timedelta

import streamlit as st
import pandas as pd

# ---------------- optional imports (safe) ----------------
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

st.set_page_config(
    page_title="Hackathon Image-to-Excel Extractor",
    page_icon="🏆",
    layout="wide",
)

# ---------------- page CSS ----------------
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
    </style>
    """,
    unsafe_allow_html=True,
)

# ---------------- session state ----------------
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

def parse_deadline(text):
    """Find the first plausible deadline date in text. Returns a date or None."""
    if not text:
        return None

    months = {
        "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
        "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12,
    }

    patterns = [
        r"(\d{1,2})(?:st|nd|rd|th)?\s+([A-Za-z]{3,9})\s*(\d{4})?",
        r"(\d{4})-(\d{1,2})-(\d{1,2})",
        r"(\d{1,2})[/\-](\d{1,2})[/\-](\d{4})",
    ]

    for pat in patterns:
        for m in re.finditer(pat, text, re.IGNORECASE):
            g = m.groups()
            try:
                if pat == patterns[0]:
                    day = int(g[0])
                    mon = months.get(g[1].lower()[:3])
                    year = int(g[2]) if g[2] else date.today().year
                    if mon is None:
                        continue
                    return date(year, mon, day)
                elif pat == patterns[1]:
                    return date(int(g[0]), int(g[1]), int(g[2]))
                else:
                    d, mo, y = int(g[0]), int(g[1]), int(g[2])
                    if y < 100:
                        y += 2000
                    if 1 <= mo <= 12 and 1 <= d <= 31:
                        return date(y, mo, d)
                    if 1 <= d <= 12 and 1 <= mo <= 31:
                        return date(y, d, mo)
            except (ValueError, TypeError):
                continue
    return None


def extract_qr_codes(image_bytes):
    """Return list of QR data strings found in the image."""
    codes = []
    if not PIL_OK:
        return codes

    img = Image.open(BytesIO(image_bytes))

    if PYZBAR_OK:
        try:
            for d in pyzbar.decode(img):
                data = d.data.decode("utf-8", errors="ignore").strip()
                if data:
                    codes.append(data)
        except Exception:
            pass

    if not codes and CV2_OK:
        try:
            arr = np.array(img.convert("RGB"))
            arr = cv2.cvtColor(arr, cv2.COLOR_RGB2BGR)
            detector = cv2.QRCodeDetector()
            data, _, _ = detector.detectAndDecode(arr)
            if data:
                codes.append(data)
        except Exception:
            pass

    return codes


def extract_ocr(image_bytes):
    """Run OCR on the image. Returns text or ''."""
    if not (PIL_OK and CV2_OK):
        return ""
    try:
        img = Image.open(BytesIO(image_bytes)).convert("RGB")
        arr = np.array(img)
        arr = cv2.cvtColor(arr, cv2.COLOR_RGB2BGR)
        arr = cv2.resize(arr, None, fx=1.5, fy=1.5, interpolation=cv2.INTER_CUBIC)
        gray = cv2.cvtColor(arr, cv2.COLOR_BGR2GRAY)
        return pytesseract.image_to_string(gray)
    except Exception:
        return ""


def parse_poster(text, qr_codes, source_name):
    """Pull structured fields out of OCR text + QR codes."""
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]

    name = ""
    for ln in lines:
        if "hackathon" in ln.lower() or len(ln) > 10:
            name = ln
            break
    if not name and lines:
        name = lines[0]

    theme = ""
    for ln in lines:
        if "theme" in ln.lower():
            theme = ln
            break

    deadline = parse_deadline(text)

    link = ""
    m = re.search(r"https?://\S+|www\.\S+", text)
    if m:
        link = m.group(0)
    elif qr_codes:
        link = qr_codes[0]

    return {
        "Hackathon Name": name,
        "Theme": theme,
        "Deadline": deadline.strftime("%d-%m-%Y") if deadline else "",
        "_deadline_date": deadline,
        "Application Link": link,
        "QR Codes": "; ".join(qr_codes),
        "Source Image": source_name,
        "Full OCR Text": text[:2000],
    }


def check_deadline_alerts(df, today, alert_days=3):
    """Return list of (row_index, deadline_date, days_left) for upcoming deadlines."""
    alerts = []
    if df is None or df.empty or "_deadline_date" not in df.columns:
        return alerts
    for idx, row in df.iterrows():
        d = row.get("_deadline_date")
        if isinstance(d, date):
            days_left = (d - today).days
            if 0 <= days_left <= alert_days:
                alerts.append((idx, d, days_left))
    return alerts


def send_email(sender, app_password, recipient, subject, body):
    """Send an email via Gmail SMTP using an App Password."""
    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = recipient
    msg["Subject"] = subject
    msg.set_content(body)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(sender, app_password)
        server.send_message(msg)
    return True


# ============================================================
# PAGE
# ============================================================

st.markdown(
    """
    <div class="hero">
        <div class="hero-badge">🏆 OCR • QR • Excel • Deadline Alerts</div>
        <div class="hero-title">Hackathon Image-to-Excel Extractor</div>
        <div class="hero-subtitle">
            Upload hackathon posters and automatically extract hackathon names,
            themes, deadlines, application links and QR codes into an editable
            Excel-ready table.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

# ---------------- upload ----------------
st.markdown('<div class="section-title">📤 Upload Posters</div>', unsafe_allow_html=True)

uploaded_files = st.file_uploader(
    "Choose one or more hackathon poster images",
    type=["png", "jpg", "jpeg", "webp"],
    accept_multiple_files=True,
)

if uploaded_files:
    if not CV2_OK:
        st.error("OpenCV / pytesseract is not installed. Run: pip install opencv-python pytesseract")
    else:
        progress = st.progress(0)
        for i, f in enumerate(uploaded_files):
            progress.progress((i + 1) / len(uploaded_files), text=f"Processing {f.name}...")
            qr_codes = extract_qr_codes(f.getvalue())
            text = extract_ocr(f.getvalue())
            st.session_state.results.append(parse_poster(text, qr_codes, f.name))
        progress.empty()
        st.success(f"Processed {len(uploaded_files)} image(s).")

# ---------------- editable table ----------------
st.markdown('<div class="section-title">📋 Extracted Data (editable)</div>', unsafe_allow_html=True)

if st.session_state.results:
    df = pd.DataFrame(st.session_state.results)
    display_df = df.drop(columns=["_deadline_date"], errors="ignore")

    edited = st.data_editor(
        display_df,
        width='stretch',
        height=400,
        key="edited_table",
    )

    if "_deadline_date" in df.columns:
        edited = edited.copy()
        edited["_deadline_date"] = [
            parse_deadline(v) for v in edited.get("Deadline", [""] * len(edited))
        ]
else:
    st.info("Upload posters above to see extracted data here.")
    edited = pd.DataFrame()

# ---------------- excel export ----------------
if edited is not None and not edited.empty:
    st.markdown('<div class="section-title">📥 Download Excel</div>', unsafe_allow_html=True)

    if OPENPYXL_OK:
        wb = Workbook()
        ws = wb.active
        ws.title = "Hackathons"

        out_df = edited.drop(columns=["_deadline_date"], errors="ignore")
        headers = list(out_df.columns)

        header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
        header_font = Font(bold=True, color="FFFFFF")
        for c, h in enumerate(headers, start=1):
            cell = ws.cell(row=1, column=c, value=h)
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center")

        for r, row in enumerate(out_df.itertuples(index=False), start=2):
            for c, v in enumerate(row, start=1):
                ws.cell(row=r, column=c, value="" if pd.isna(v) else str(v))

        for c, h in enumerate(headers, start=1):
            width = max(len(str(h)), 15)
            ws.column_dimensions[chr(64 + c) if c <= 26 else "A"].width = min(width + 2, 45)

        buf = BytesIO()
        wb.save(buf)
        buf.seek(0)
        st.download_button(
            "⬇️ Download Hackathons.xlsx",
            data=buf,
            file_name="Hackathons.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            width='stretch',
        )
    else:
        st.download_button(
            "⬇️ Download (CSV)",
            data=edited.drop(columns=["_deadline_date"], errors="ignore").to_csv(index=False).encode(),
            file_name="Hackathons.csv",
            mime="text/csv",
        )

# ---------------- deadline alerts ----------------
st.markdown('<div class="section-title">🚨 Deadline Alerts</div>', unsafe_allow_html=True)

if edited is not None and not edited.empty:
    alerts = check_deadline_alerts(edited, date.today(), alert_days=3)
    if alerts:
        for idx, d, days_left in alerts:
            name = edited.loc[idx, "Hackathon Name"] if "Hackathon Name" in edited.columns else "Unknown"
            if days_left == 0:
                st.error(f"⏰ {name} — deadline is TODAY ({d.strftime('%d-%m-%Y')})!")
            else:
                st.warning(f"⏰ {name} — deadline in {days_left} day(s) ({d.strftime('%d-%m-%Y')})")
    else:
        st.success("✅ No deadlines in the next 3 days.")
else:
    st.info("No data yet — upload posters to check deadlines.")

# ---------------- email alerts ----------------
st.markdown('<div class="section-title">📧 Email Alerts</div>', unsafe_allow_html=True)

st.markdown(
    """
    <div class="card">
    Configure Gmail SMTP to receive hackathon deadline alerts.
    <br><small>Use a Gmail <b>App Password</b> (Google Account → Security → 2-Step
    Verification → App passwords), not your normal Gmail password.</small>
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

# NOTE: Do NOT write st.session_state.email_app_password_config = email_app_password here.
# The widget already syncs its value to session state automatically.

if st.button("📤 Send Deadline Alert Email", type="primary"):
    if not email_sender or not email_app_password or not email_recipient:
        st.error("Fill in all three email fields first.")
    else:
        try:
            lines = []
            if edited is not None and not edited.empty:
                for idx, d, days_left in check_deadline_alerts(edited, date.today(), alert_days=3):
                    name = edited.loc[idx, "Hackathon Name"] if "Hackathon Name" in edited.columns else "Unknown"
                    lines.append(f"• {name} — deadline {d.strftime('%d-%m-%Y')} ({days_left} day(s) left)")
            if not lines:
                lines.append("No deadlines in the next 3 days.")

            body = "Hackathon Deadline Alert\n\n" + "\n".join(lines) + f"\n\nSent on {date.today()}"
            send_email(email_sender, email_app_password, email_recipient,
                       "⏰ Hackathon Deadline Alert", body)
            st.session_state.email_last_success = True
            st.session_state.email_last_message = "Email sent successfully!"
            st.success("✅ Email sent successfully!")
        except Exception as e:
            st.session_state.email_last_success = False
            st.session_state.email_last_message = str(e)
            st.error(f"❌ Email failed: {e}")
            st.caption("Hint: make sure 2-Step Verification is ON in your Gmail account "
                       "and you are using an App Password (16 characters).")