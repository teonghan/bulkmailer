"""
Streamlit Bulk-Mailer (Gmail API, no yagmail)
────────────────────────────────────────────
• Excel mail-merge — tags {Name}, {Salutation}, … in subject & body
• Quill editors for header / body / footer
• Optional header / footer images
• Per-row PDF (URL or local path)
• Auth once via Google OAuth; token stored at ~/.credentials/gmail_token.pickle
"""

import base64, mimetypes, pathlib, pickle, re
from io import BytesIO
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase

from email import encoders          # ‼️ add at top with other imports

import pandas as pd
import requests
import streamlit as st
from streamlit_quill import st_quill
from jinja2 import Environment, Undefined, select_autoescape
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build

import time
import gdown


# ─── OAuth / Gmail helpers (from your Nexus Mailer) ───────────────────────────
SCOPES = ["https://www.googleapis.com/auth/gmail.send"]
TOKEN_PKL = pathlib.Path.home() / ".credentials" / "gmail_token.pickle"
CLIENT_SECRET = pathlib.Path("client_secret.json")  # must exist

def get_gmail_service():
    TOKEN_PKL.parent.mkdir(parents=True, exist_ok=True)
    creds = None
    if TOKEN_PKL.exists():
        creds = pickle.loads(TOKEN_PKL.read_bytes())
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
    if not creds or not creds.valid:
        flow = InstalledAppFlow.from_client_secrets_file(CLIENT_SECRET, SCOPES)
        creds = flow.run_local_server(port=0)
        TOKEN_PKL.write_bytes(pickle.dumps(creds))
    return build("gmail", "v1", credentials=creds), creds

def build_message(sender, to, subject, html_body, attachment=None):
    """Return Gmail API message dict (handles one attachment or none)."""
    if attachment:
        msg = MIMEMultipart()
        msg.attach(MIMEText(html_body, "html"))
        fname, data = attachment
        ctype, _ = mimetypes.guess_type(fname)
        maintype, subtype = (ctype or "application/octet-stream").split("/", 1)
        part = MIMEBase(maintype, subtype)
        part.set_payload(data)
        part.add_header("Content-Disposition", "attachment", filename=fname)
        part.add_header("Content-Type", ctype or "application/octet-stream")
        msg.attach(part)
    else:
        msg = MIMEText(html_body, "html")

    msg["To"] = to
    msg["From"] = sender
    msg["Subject"] = subject
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    return {"raw": raw}

def send_via_gmail(service, message):
    return service.users().messages().send(userId="me", body=message).execute()

def create_message_with_attachment(sender, to, subject, html_body, file_tuple):
    filename, payload = file_tuple
    message = MIMEMultipart()
    message["To"] = to
    message["From"] = sender
    message["Subject"] = subject
    message.attach(MIMEText(html_body, "html"))

    ctype, _ = mimetypes.guess_type(filename)
    maintype, subtype = (ctype or "application/pdf").split("/", 1)

    part = MIMEBase(maintype, subtype)
    if isinstance(payload, str):
        payload = payload.encode()
    part.set_payload(payload)

    encoders.encode_base64(part)                                # ← NEW
    part.add_header("Content-Type", ctype or "application/pdf")
    part.add_header("Content-Disposition", "attachment", filename=filename)

    message.attach(part)
    raw = base64.urlsafe_b64encode(message.as_bytes()).decode()
    return {"raw": raw}


# ─── Jinja env (single-brace tags) ─────────────────────────────────────────────
class SilentUndef(Undefined):
    def _fail_with_undefined_error(self, *a, **kw):
        return ""
jinja_env = Environment(
    variable_start_string="{",
    variable_end_string="}",
    undefined=SilentUndef,
    autoescape=select_autoescape(enabled_extensions=("html",)),
)

def clean_quill(html: str) -> str:
    html = html.replace("&#123;", "{").replace("&#125;", "}")
    html = re.sub(r"<\w+[^>]*>({\w+})<\/\w+>", r"\1", html)
    html = re.sub(r"{ *([^} ]*?) *}", r"{\1}", html)
    html = re.sub(r"<p><br></p>", "", html)
    return html

def fix_inline_img_widths(html: str, width: int) -> str:
    """Force every <img ...> in Quill HTML to the chosen width."""
    return re.sub(
        r'<img([^>]*?)>',
        lambda m: (
            f'<img{m.group(1)} style="width:{width}px;max-width:100%;'
            'display:block;margin:0 auto;" />'
        ),
        html,
        flags=re.IGNORECASE,
    )

def inline_p_spacing(html: str,
                     margin: str = "0 0 0 0",
                     lh: str = "1.4") -> str:
    """Add spacing *and justify alignment* to every <p …> tag."""
    pattern = re.compile(r"<p\b([^>]*)>", flags=re.IGNORECASE)

    def repl(match):
        attrs = match.group(1)
        if "margin" in attrs or "line-height" in attrs:
            return f"<p{attrs}>"
        style = (
            f'style="margin:{margin};line-height:{lh};'
            'text-align:justify;"'
        )
        return f"<p{attrs} {style}>"

    return pattern.sub(repl, html)

def to_img_tag(file, width, br_after=False):
    if not file:
        return ""
    b64 = base64.b64encode(file.read()).decode()
    tag = (
        f'<img src="data:image/png;base64,{b64}" '
        f'style="width:{width}px;max-width:100%;display:block;margin:0 auto;" />'
    )
    return tag + ("<br>" if br_after else "")


# ─── Fun re-branding! ─────────────────────────────────────────────
APP_NAME = "Mail-MagiK ✨📧"

# Page config + fun intro
st.set_page_config(page_title=APP_NAME, layout="centered",
                   initial_sidebar_state="expanded")

st.title(APP_NAME)

st.markdown(
    """
    **Welcome, daring communicator!**

    • 🪄 Just sprinkle your Excel list into the sidebar cauldron.\n
    • 🪄 Conjure a charming subject and body with `{Tags}` for names, links, emoji—whatever.\n
    • 🪄 Pick a width, images, Google-Drive links…\n
    • 🪄 Press **Preview** to gaze into the crystal ball, then **Send bulk emails** to unleash the owls—I mean, SMTP-owls—one by one (1-second pause so Gmail stays happy).\n

    _May your inbox be ever spellbound!_
    """
)

with st.sidebar:
    xlsx = st.file_uploader("Excel file", ["xlsx"])
    sender = st.text_input("Sender Gmail address")
    subj_tpl = st.text_input("Email Subject (tags OK)", "Hello {Name}")
    hdr_img = st.file_uploader("Header image", ["png","jpg","jpeg"])
    ftr_img = st.file_uploader("Footer image", ["png","jpg","jpeg"], key="ftr_img")
    # NEW — choose a common width
    img_width = st.number_input(
        "Width for images & text (px)",
        min_value=200, max_value=1200, value=600, step=20
    )
    st.info("First send will open Google consent screen; token is cached.")

# ─── Excel / tags ─────────────────────────────────────────────────────────────
df, TAGS = None, []
if xlsx:
    df = pd.read_excel(xlsx)
    df.columns = df.columns.str.strip()
    TAGS = df.columns.tolist()
    st.sidebar.markdown("Tags: " + " ".join(f"`{{{t}}}`" for t in TAGS))

# ─── Editors ──────────────────────────────────────────────────────────────────
st.markdown("### Header"); header_html = st_quill(html=True, key="hdr")
st.markdown("### Body");   body_html   = st_quill(html=True, key="bdy")
st.markdown("### Footer"); footer_html = st_quill(html=True, key="ftr")

preview_btn, send_btn = st.columns(2)
preview_click = preview_btn.button("Preview first email")
send_click    = send_btn.button("Send bulk emails")

# clean editor output
header_html = inline_p_spacing(
    fix_inline_img_widths(clean_quill(header_html), img_width)
)
body_html   = inline_p_spacing(
    fix_inline_img_widths(clean_quill(body_html), img_width)
)
footer_html = inline_p_spacing(
    fix_inline_img_widths(clean_quill(footer_html), img_width)
)

hdr_tag = to_img_tag(hdr_img, img_width, br_after=True)
ftr_tag = "<br>" + to_img_tag(ftr_img, img_width) if ftr_img else ""

# after you compute img_width
style_tag = (
    "{% raw %}"
    "<style>"
    f".mail-preview p {{margin:0 0 0em 0;line-height:1.4;}}"
    "</style>"
    "{% endraw %}"
)

wrapper_start = (
    f'<div class="mail-preview" '
    f'style="max-width:{img_width}px;margin:0 auto;">'
)

body_template = (
    style_tag +
    wrapper_start +
    hdr_tag +
    header_html + body_html + footer_html +
    ftr_tag +
    '</div>'
)

subj_template = jinja_env.from_string(subj_tpl)

# ─── Preview ──────────────────────────────────────────────────────────────────
if preview_click and df is not None:
    sample = df.iloc[0].to_dict()
    st.markdown(f"**Subject:** {subj_template.render(**sample)}")
    st.markdown(jinja_env.from_string(body_template).render(**sample),
                unsafe_allow_html=True)

# ─── Bulk send ────────────────────────────────────────────────────────────────
if send_click:
    if df is None or not sender:
        st.error("Please provide Excel file **and** sender address.")
    else:
        service, creds = get_gmail_service()
        logs = []
        st.info("Sending…")
        for _, row in df.iterrows():
            data = row.to_dict()
            html = jinja_env.from_string(body_template).render(**data)
            subj = subj_template.render(**data)

            # optional one attachment
            attach = None
            pdf = str(data.get("PDF Link","")).strip()
            if pdf:
                try:
                    if pdf.lower().startswith("http"):
                        tmp_path = gdown.download(pdf, quiet=True)
                        fname = pathlib.Path(tmp_path).name
                        if not fname.lower().endswith(".pdf"):
                            fname += ".pdf"                       # force .pdf so Gmail knows it
                        filebytes = pathlib.Path(tmp_path).read_bytes()
                        attach = (fname, filebytes)
                    else:
                        filebytes = pathlib.Path(pdf).read_bytes()
                        attach = (pathlib.Path(pdf).name, filebytes)
                except Exception as e:
                    logs.append({"email": data.get("Email",""),
                                 "status": f"PDF err: {e}"})
                    continue

            message = (create_message_with_attachment(sender,
                                                      data.get("Email",""),
                                                      subj, html, attach)
                        if attach else
                        build_message(sender, data.get("Email",""), subj, html))

            try:
                send_via_gmail(service, message)
                logs.append({"email": data.get("Email",""), "status": "Sent"})
            except Exception as e:
                logs.append({"email": data.get("Email",""), "status": f"Err: {e}"})

            time.sleep(1)  # pause before next email

        st.success("Done")
        st.write(pd.DataFrame(logs))
