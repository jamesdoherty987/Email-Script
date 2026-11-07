import base64
import csv
import logging
import os
import re
import sys
import time
from datetime import datetime
from pathlib import Path

import resend
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from dotenv import load_dotenv


# ============================================================
# SETTINGS
# ============================================================

BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / ".env")

CSV_FILE = BASE_DIR / "sponsors.csv"
LOGO_FILE = BASE_DIR / "logo.png"
OUTPUT_DIR = BASE_DIR / "generated_proposals"
LOG_FILE = BASE_DIR / "outreach_log.csv"
DEBUG_LOG_FILE = BASE_DIR / "outreach_debug.log"

YOUR_EMAIL = os.getenv("YOUR_EMAIL", "jkdoherty123@gmail.com")
YOUR_NAME = os.getenv("YOUR_NAME", "James Doherty")
YOUR_PHONE = os.getenv("YOUR_PHONE", "085 263 5954")
YOUR_WEBSITE = os.getenv("YOUR_WEBSITE", "https://www.studentpadel.ie")

# Resend "from" must be a verified domain address (or onboarding@resend.dev for tests).
FROM_EMAIL = os.getenv(
    "FROM_EMAIL",
    f"Student Padel League <{YOUR_EMAIL}>",
)
RESEND_API_KEY = os.getenv("RESEND_API_KEY", "").strip()

EVENT_NAME = "Student Padel League - Limerick Open"
EVENT_DATE = "12 November 2026"
EVENT_VENUE = "Padel LK, Limerick"

# SEND_EMAILS=true actually sends via Resend. false = build proposals only.
SEND_EMAILS = os.getenv("SEND_EMAILS", "false").strip().lower() in {
    "1",
    "true",
    "yes",
    "on",
}

# Pause between Resend calls to reduce rate-limit risk on bulk sends.
SEND_DELAY_SECONDS = float(os.getenv("SEND_DELAY_SECONDS", "1.5"))

OUTPUT_DIR.mkdir(exist_ok=True)


# ============================================================
# LOGGING
# ============================================================

def setup_logging():
    logger = logging.getLogger("outreach")
    logger.setLevel(logging.DEBUG)
    logger.handlers.clear()

    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)-7s | %(message)s",
        datefmt="%H:%M:%S",
    )

    console = logging.StreamHandler(sys.stdout)
    console.setLevel(logging.DEBUG)
    console.setFormatter(formatter)
    logger.addHandler(console)

    file_handler = logging.FileHandler(DEBUG_LOG_FILE, encoding="utf-8")
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    return logger


log = setup_logging()


# ============================================================
# HELPERS
# ============================================================

def clean_filename(name):
    name = re.sub(r'[<>:"/\\|?*]', "", name)
    name = re.sub(r"\s+", "_", name.strip())
    return name[:100] or "sponsor"


def value(row, key, default=""):
    return (row.get(key) or default).strip()


EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


def is_valid_email(address):
    """Require a real email address, not a category/location placeholder."""
    address = (address or "").strip()
    if not address or " " in address:
        return False
    return bool(EMAIL_RE.match(address))


def normalize_company(name):
    return re.sub(r"\s+", " ", (name or "").strip().lower())


def load_already_sent():
    """Load emails/companies already marked SENT so we never email twice.

    BOUNCED/ERROR rows do not block retries (e.g. after swapping to a new address).
    """
    sent_emails = set()
    sent_companies = set()
    if not LOG_FILE.exists():
        return sent_emails, sent_companies

    with LOG_FILE.open("r", newline="", encoding="utf-8-sig") as file:
        for row in csv.DictReader(file):
            status = (row.get("status") or "").strip().upper()
            if status != "SENT":
                continue
            email = (row.get("email") or "").strip().lower()
            company = normalize_company(row.get("company") or "")
            if email:
                sent_emails.add(email)
            if company:
                sent_companies.add(company)

    return sent_emails, sent_companies


def append_log_rows(rows):
    columns = ["date", "company", "email", "status", "details"]
    write_header = not LOG_FILE.exists() or LOG_FILE.stat().st_size == 0
    with LOG_FILE.open("a", newline="", encoding="utf-8") as file:
        writer = csv.DictWriter(file, fieldnames=columns)
        if write_header:
            writer.writeheader()
        writer.writerows(rows)


def env_bool_note():
    return "SEND" if SEND_EMAILS else "DRY-RUN (proposals only, no Resend calls)"


def add_black_heading(doc, text, level=1):
    heading = doc.add_heading(text, level=level)
    for run in heading.runs:
        run.font.color.rgb = RGBColor(0, 0, 0)
        run.font.name = "Arial"
    return heading


# ============================================================
# CREATE PERSONALIZED WORD PROPOSAL
# ============================================================

def create_proposal(company, sponsorship_amount, sponsorship_type):
    log.debug("Creating proposal for %s", company)
    doc = Document()

    section = doc.sections[0]
    section.top_margin = Inches(0.65)
    section.bottom_margin = Inches(0.65)
    section.left_margin = Inches(0.8)
    section.right_margin = Inches(0.8)

    black = RGBColor(0, 0, 0)

    normal = doc.styles["Normal"]
    normal.font.name = "Arial"
    normal.font.size = Pt(10)
    normal.font.color.rgb = black
    normal.paragraph_format.space_after = Pt(5)

    for level in range(1, 10):
        style_name = f"Heading {level}"
        if style_name in doc.styles:
            heading_style = doc.styles[style_name]
            heading_style.font.color.rgb = black
            heading_style.font.name = "Arial"

    if LOGO_FILE.exists():
        logo_paragraph = doc.add_paragraph()
        logo_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        logo_paragraph.paragraph_format.space_after = Pt(8)
        logo_paragraph.add_run().add_picture(str(LOGO_FILE), width=Inches(1.6))
        log.debug("Attached logo: %s", LOGO_FILE.name)
    else:
        log.warning("logo.png not found — proposal will have no logo")

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title.paragraph_format.space_after = Pt(2)
    run = title.add_run(EVENT_NAME)
    run.bold = True
    run.font.name = "Arial"
    run.font.size = Pt(21)
    run.font.color.rgb = black

    subtitle = doc.add_paragraph()
    subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = subtitle.add_run("SPONSORSHIP PROPOSAL")
    run.bold = True
    run.font.size = Pt(13)
    run.font.color.rgb = black

    event_line = doc.add_paragraph()
    event_line.alignment = WD_ALIGN_PARAGRAPH.CENTER
    event_line.add_run(f"{EVENT_DATE}  |  {EVENT_VENUE}")

    doc.add_paragraph()

    add_black_heading(doc, "About the League")
    doc.add_paragraph(
        "The Limerick Student Padel League is a student competition "
        "bringing together up to 48 teams (96 university students) "
        "for a day of competitive padel in Limerick."
    )
    doc.add_paragraph(
        "Our goal is to establish the league as a recurring student "
        "event and expand to other universities and locations across Ireland."
    )

    add_black_heading(doc, "Proposed Partnership")
    doc.add_paragraph(
        f"We would love to welcome {company} as an official sponsor "
        "of the Student Padel League."
    )
    doc.add_paragraph(
        f"Proposed sponsorship: {sponsorship_amount}. "
        "We are flexible and open to discussing a package that suits "
        "your business and objectives."
    )
    doc.add_paragraph(
        f"Partnership format: {sponsorship_type}. "
        "We would also be happy to discuss a combination of financial "
        "support, products, vouchers, memberships or prizes."
    )

    add_black_heading(doc, "What Your Business Receives")
    benefits = [
        "Logo placement on our website and relevant social media content",
        "Recognition as an official Student Padel League sponsor",
        "Branding opportunities on event banners and promotional material",
        "Social media promotion and tagging where appropriate",
        "Opportunity to provide prizes, vouchers or products to participants",
        "Exposure to up to 96 university students in Limerick",
        "Opportunity to promote relevant student offers to participants",
    ]
    for benefit in benefits:
        doc.add_paragraph(benefit, style="List Bullet")

    add_black_heading(doc, "Long-Term Partnership")
    doc.add_paragraph(
        "The league is intended to grow beyond its first event, with "
        "plans to expand into other universities and cities across Ireland."
    )
    doc.add_paragraph(
        f"We would be keen to explore a longer-term partnership with "
        f"{company} as the league grows and reaches more students."
    )

    add_black_heading(doc, "Get in Touch")
    doc.add_paragraph(
        f"{YOUR_NAME}\n"
        "Student Padel League\n"
        f"Email: {YOUR_EMAIL}\n"
        f"Phone: {YOUR_PHONE}\n"
        f"Website: {YOUR_WEBSITE}"
    )

    for section in doc.sections:
        footer = section.footer.paragraphs[0]
        footer.alignment = WD_ALIGN_PARAGRAPH.CENTER
        footer.add_run(
            "Student Padel League  |  Limerick Open"
        ).font.size = Pt(8)

    output_path = OUTPUT_DIR / (
        clean_filename(company) + "_Sponsorship_Proposal.docx"
    )
    doc.save(output_path)
    log.info("Proposal saved: %s (%s bytes)", output_path.name, output_path.stat().st_size)
    return output_path


# ============================================================
# EMAIL CONTENT
# ============================================================

def build_email_content(row):
    company = value(row, "company", "your business")
    contact_name = value(row, "contact_name")
    amount = value(row, "sponsorship_amount", "a flexible package")
    sponsorship_type = value(row, "sponsorship_type", "sponsorship support")

    greeting = f"Hi {contact_name}," if contact_name else f"Hello {company} team,"
    subject = "Sponsorship opportunity | Student Padel League, Limerick"

    body = f"""{greeting}

I'm getting in touch on behalf of the Student Padel League about our
Limerick Open at Padel LK, planned for {EVENT_DATE}.

We're building a student padel competition bringing together up to
48 teams (96 university students). We're looking for a small number
of local and relevant business partners to support the event and help
us grow the league across Ireland.

I thought {company} could be a good fit for a potential partnership.

Our initial proposal is {amount}, with {sponsorship_type}. We're flexible
and happy to discuss a package that makes sense for your business,
including prizes, vouchers or products where appropriate.

I've attached a proposal with further details and potential benefits.

Would you be open to a quick conversation about this?

Thanks for your time.

Kind regards,
{YOUR_NAME}
Student Padel League
{YOUR_EMAIL}
{YOUR_PHONE}
{YOUR_WEBSITE}
"""
    return subject, body


# ============================================================
# SEND VIA RESEND
# ============================================================

def send_via_resend(to_email, subject, body, proposal_path):
    if not RESEND_API_KEY:
        raise RuntimeError(
            "RESEND_API_KEY is missing. Add it to .env in the project folder."
        )

    resend.api_key = RESEND_API_KEY
    attachment_b64 = base64.b64encode(proposal_path.read_bytes()).decode("utf-8")

    payload = {
        "from": FROM_EMAIL,
        "to": [to_email],
        "reply_to": YOUR_EMAIL,
        "subject": subject,
        "text": body,
        "attachments": [
            {
                "filename": proposal_path.name,
                "content": attachment_b64,
            }
        ],
    }

    log.debug(
        "Resend payload: from=%s to=%s subject=%r attachment=%s (%s bytes)",
        FROM_EMAIL,
        to_email,
        subject,
        proposal_path.name,
        proposal_path.stat().st_size,
    )

    result = resend.Emails.send(payload)
    email_id = result.get("id", "") if isinstance(result, dict) else getattr(result, "id", "")
    log.info("Resend accepted email id=%s", email_id or result)
    return email_id or str(result)


# ============================================================
# MAIN
# ============================================================

def main():
    log.info("Starting sponsor outreach")
    log.info("Mode: %s", env_bool_note())
    log.info("CSV: %s", CSV_FILE)
    log.info("From: %s | Reply-To: %s", FROM_EMAIL, YOUR_EMAIL)
    log.info(
        "Resend API key: %s",
        f"set (...{RESEND_API_KEY[-4:]})" if RESEND_API_KEY else "MISSING",
    )

    if not CSV_FILE.exists():
        raise FileNotFoundError(f"Missing spreadsheet: {CSV_FILE}")

    if SEND_EMAILS and not RESEND_API_KEY:
        raise RuntimeError(
            "SEND_EMAILS is true but RESEND_API_KEY is not set in .env"
        )

    with CSV_FILE.open("r", newline="", encoding="utf-8-sig") as file:
        sponsors = list(csv.DictReader(file))

    if not sponsors:
        log.warning("No sponsors found in sponsors.csv")
        return

    log.info("Loaded %s sponsor row(s)", len(sponsors))
    log.debug("CSV columns: %s", list(sponsors[0].keys()))

    already_sent_emails, already_sent_companies = load_already_sent()
    log.info(
        "Already sent: %s email(s), %s company(ies) — these will be skipped",
        len(already_sent_emails),
        len(already_sent_companies),
    )

    log_rows = []
    skip_already = 0
    skip_blank = 0
    skip_invalid = 0
    to_send = []
    for row in sponsors:
        company = value(row, "company")
        email_address = value(row, "email")
        if (
            company
            and is_valid_email(email_address)
            and email_address.lower() not in already_sent_emails
            and normalize_company(company) not in already_sent_companies
        ):
            to_send.append((company, email_address))

    log.info("Ready to process: %s new contact(s)", len(to_send))
    for company, email_address in to_send:
        log.info("  -> %s <%s>", company, email_address)

    for index, row in enumerate(sponsors, start=1):
        company = value(row, "company")
        email_address = value(row, "email")
        log.info("[%s/%s] Processing %s <%s>", index, len(sponsors), company or "?", email_address or "?")

        if not company:
            log.warning("[%s] Skipping — company is blank", index)
            skip_blank += 1
            continue

        if not email_address:
            log.warning("[%s] Skipping %s — email is blank", index, company)
            skip_blank += 1
            continue

        if not is_valid_email(email_address):
            log.warning(
                "[%s] Skipping %s — invalid email %r "
                "(CSV columns may be misaligned; expected email@domain.com)",
                index,
                company,
                email_address,
            )
            skip_invalid += 1
            continue

        company_key = normalize_company(company)
        email_key = email_address.lower()
        if email_key in already_sent_emails or company_key in already_sent_companies:
            log.info(
                "[%s] Skipping %s — already sent previously",
                index,
                company,
            )
            skip_already += 1
            continue

        try:
            amount = value(row, "sponsorship_amount", "Flexible sponsorship package")
            sponsorship_type = value(
                row, "sponsorship_type", "Cash, products or prizes"
            )

            proposal_path = create_proposal(company, amount, sponsorship_type)
            subject, body = build_email_content(row)
            log.debug("Subject: %s", subject)
            log.debug("Body preview: %s...", body[:120].replace("\n", " "))

            if SEND_EMAILS:
                message_id = send_via_resend(
                    email_address, subject, body, proposal_path
                )
                status = "SENT"
                details = message_id
                already_sent_emails.add(email_key)
                already_sent_companies.add(company_key)
                log.info("[%s] SENT to %s", company, email_address)
                if SEND_DELAY_SECONDS > 0:
                    time.sleep(SEND_DELAY_SECONDS)
            else:
                status = "DRY-RUN - PROPOSAL ONLY"
                details = proposal_path.name
                log.info(
                    "[%s] Dry-run complete (set SEND_EMAILS=true in .env to send)",
                    company,
                )

            log_rows.append({
                "date": datetime.now().isoformat(timespec="seconds"),
                "company": company,
                "email": email_address,
                "status": status,
                "details": details,
            })

        except Exception as error:
            log.error("[%s] ERROR for %s: %s", index, company, error)
            log.debug("Full traceback for %s", company, exc_info=True)
            log_rows.append({
                "date": datetime.now().isoformat(timespec="seconds"),
                "company": company,
                "email": email_address,
                "status": "ERROR",
                "details": str(error),
            })

    # Persist only real outcomes (sent/errors). Dry-runs stay out of the send log.
    persist_rows = [
        row for row in log_rows
        if row["status"] == "SENT" or row["status"] == "ERROR"
    ]
    if persist_rows:
        append_log_rows(persist_rows)

    sent = sum(1 for r in log_rows if r["status"] == "SENT")
    errors = sum(1 for r in log_rows if r["status"] == "ERROR")
    dry = sum(1 for r in log_rows if r["status"].startswith("DRY-RUN"))

    log.info(
        "Finished. sent=%s dry_run=%s errors=%s already_sent_skipped=%s blank_skipped=%s invalid_skipped=%s",
        sent,
        dry,
        errors,
        skip_already,
        skip_blank,
        skip_invalid,
    )
    log.info("Proposals: %s", OUTPUT_DIR)
    log.info("Outreach log: %s", LOG_FILE)
    log.info("Debug log: %s", DEBUG_LOG_FILE)


if __name__ == "__main__":
    try:
        main()
    except Exception:
        log.exception("Fatal error — script stopped")
        sys.exit(1)
