#!/usr/bin/env python3
"""
Unified Student Calendar — Live Sync Backend Server
Supports real live syncing for:
  1. Canvas (Canvas LMS REST API v1 + iCal .ics feed)
  2. Gmail (Google Gmail API + Google Calendar API + IMAP SSL + iCal .ics feed)
  3. Outlook (Microsoft Graph API Calendar/Mail + Office365 IMAP SSL + iCal .ics feed)
  4. Teams (Microsoft Graph API Online Meetings/Calendar/Teams + Exchange .ics feed)
"""

import datetime
import email
from email.header import decode_header
import http.server
import imaplib
import json
import os
import re
import ssl
import urllib.error
import urllib.parse
import urllib.request
from pathlib import Path

PORT = int(os.environ.get("PORT", "9090"))
BASE_DIR = Path(__file__).resolve().parent
SSL_CTX = ssl.create_default_context()


def classify_student_category(title: str, description: str = "", default_cat: str = "duedates", integration_id: str = "") -> str:
    """Auto-classify synced items into the 5 student streams:
    duedates, clubs, mentorship, email, sports.
    Canvas platforms can be anything except sports.
    """
    text = f"{title} {description}".lower()
    is_canvas_platform = integration_id == "canvas" or integration_id.startswith("canvas")

    if not is_canvas_platform and any(w in text for w in ["sport", "track", "varsity", "conditioning", "coach", "game", "meet", "athletic", "gym", "soccer", "basketball", "swim", "tennis"]):
        return "sports"
    if any(w in text for w in ["mentor", "advising", "advisor", "1-on-1", "1:1", "tutor", "office hours", "counselor", "checkin", "check-in", "career"]):
        return "mentorship"
    if any(w in text for w in ["club", "society", "robotics", "mun", "model un", "debate", "senate", "student gov", "chapter", "tsa", "fbla", "deca", "nhs", "key club"]):
        return "clubs"
    if any(w in text for w in ["due", "assignment", "quiz", "exam", "test", "homework", "lab", "problem set", "module", "essay", "project", "submission", "pacing"]):
        return "duedates"
    return default_cat


def dedupe_canvas_events(events: list) -> list:
    """Deduplicate Canvas assignments that appear as duplicate pairs on the same date
    (e.g., 'Unit 4 Quiz' and 'Unit 4 Quiz [Course Name]'). Keep the shorter title for each pair.
    """
    def norm(t: str) -> str:
        return re.sub(r"\s+", " ", re.sub(r"[^a-z0-9\s]", " ", t.lower())).strip()

    kept = []
    for ev in sorted(events, key=lambda e: len(e.get("title", ""))):
        ev_norm = norm(ev.get("title", ""))
        is_dup = False
        for k in kept:
            if k.get("date") != ev.get("date"):
                continue
            k_norm = norm(k.get("title", ""))
            if k_norm and ev_norm and (k_norm in ev_norm or ev_norm in k_norm):
                is_dup = True
                break
        if not is_dup:
            kept.append(ev)
    return kept


def is_canvas_excluded(title: str, description: str = "") -> bool:
    """Return True if a Canvas event should be auto-filtered out
    (Q/A sessions, OpenStax, or AP Daily Videos except AP Daily Video Review).
    """
    text = f"{title} {description}".lower()
    if re.search(r"\bq\s*[/&]\s*a\b|\bquestion\s*(?:and|&)\s*answer\b", text):
        return True
    if "openstax" in text:
        return True
    if "daily video" in text and "daily video review" not in text:
        return True
    return False


def http_get_json(url: str, headers: dict = None) -> object:
    req = urllib.request.Request(url, headers=headers or {})
    req.add_header("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    try:
        with urllib.request.urlopen(req, context=SSL_CTX, timeout=20) as resp:
            return json.loads(resp.read().decode("utf-8", errors="replace"))
    except urllib.error.HTTPError as e:
        err_msg = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {e.code} ({e.reason}): {err_msg[:200]}") from e
    except urllib.error.URLError as e:
        raise RuntimeError(f"Network error accessing {url}: {e.reason}") from e


def http_get_text(url: str, headers: dict = None) -> str:
    # Handle webcal:// protocol scheme commonly copied from calendar feeds
    clean_url = url.strip()
    if clean_url.startswith("webcal://"):
        clean_url = "https://" + clean_url[9:]
    elif clean_url.startswith("http://"):
        # Upgrade to https if possible
        clean_url = "https://" + clean_url[7:]

    req = urllib.request.Request(clean_url, headers=headers or {})
    req.add_header("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    try:
        with urllib.request.urlopen(req, context=SSL_CTX, timeout=25) as resp:
            return resp.read().decode("utf-8", errors="replace")
    except urllib.error.HTTPError as e:
        err_msg = e.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"HTTP {e.code} ({e.reason}) fetching feed: {err_msg[:200]}") from e
    except urllib.error.URLError as e:
        raise RuntimeError(f"Network error fetching calendar feed from {clean_url}: {e.reason}") from e


def unfold_ics(ics_text: str) -> str:
    return re.sub(r"\r?\n[ \t]", "", ics_text)


def parse_ics_datetime(raw_val: str):
    """Parse iCalendar DTSTART/DTEND value into (YYYY-MM-DD, HH:MM)."""
    val = raw_val.strip()
    m = re.search(r"(\d{4})(\d{2})(\d{2})(?:T(\d{2})(\d{2})(\d{2})?(Z)?)?", val)
    if not m:
        return None, "12:00"
    yyyy, mm, dd, hh, mins, _sec, is_utc = m.groups()
    if hh and mins:
        if is_utc:
            try:
                dt_utc = datetime.datetime(int(yyyy), int(mm), int(dd), int(hh), int(mins), tzinfo=datetime.timezone.utc)
                local_dt = dt_utc.astimezone()
                return local_dt.strftime("%Y-%m-%d"), local_dt.strftime("%H:%M")
            except Exception:
                pass
        return f"{yyyy}-{mm}-{dd}", f"{hh}:{mins}"
    return f"{yyyy}-{mm}-{dd}", "23:59"


def parse_ics_feed(ics_text: str, integration_id: str, default_category: str = "duedates") -> list:
    """Parse standard RFC 5545 iCalendar (.ics) text into unified calendar event objects."""
    unfolded = unfold_ics(ics_text)
    blocks = unfolded.split("BEGIN:VEVENT")[1:]
    events = []

    for idx, block in enumerate(blocks):
        end_idx = block.find("END:VEVENT")
        if end_idx != -1:
            block = block[:end_idx]

        summary_m = re.search(r"^SUMMARY(?:;[^:]*)?:(.*)$", block, re.MULTILINE | re.IGNORECASE)
        dtstart_m = re.search(r"^DTSTART(?:;[^:]*)?:(.*)$", block, re.MULTILINE | re.IGNORECASE)
        dtend_m = re.search(r"^DTEND(?:;[^:]*)?:(.*)$", block, re.MULTILINE | re.IGNORECASE)
        loc_m = re.search(r"^LOCATION(?:;[^:]*)?:(.*)$", block, re.MULTILINE | re.IGNORECASE)
        desc_m = re.search(r"^DESCRIPTION(?:;[^:]*)?:(.*)$", block, re.MULTILINE | re.IGNORECASE)
        uid_m = re.search(r"^UID(?:;[^:]*)?:(.*)$", block, re.MULTILINE | re.IGNORECASE)

        raw_date = dtstart_m.group(1) if dtstart_m else (dtend_m.group(1) if dtend_m else "")
        if not summary_m or not raw_date:
            continue

        title = summary_m.group(1).strip().replace("\\,", ",").replace("\\;", ";").replace("\\n", " ")
        date_str, time_str = parse_ics_datetime(raw_date)
        if not date_str:
            continue

        loc = loc_m.group(1).strip().replace("\\,", ",") if loc_m else ""
        desc = desc_m.group(1).strip().replace("\\n", " ")[:200] if desc_m else ""
        uid = uid_m.group(1).strip() if uid_m else f"{integration_id}-{idx}-{date_str}"

        if (integration_id == "canvas" or integration_id.startswith("canvas")) and is_canvas_excluded(title, f"{loc} {desc}"):
            continue

        category = classify_student_category(title, f"{loc} {desc}", default_category, integration_id)
        events.append({
            "id": f"sync-{integration_id}-{abs(hash(uid)) % 1000000}",
            "title": title,
            "category": category,
            "integration": integration_id,
            "date": date_str,
            "time": time_str,
            "location": loc or integration_id.upper(),
            "completed": False
        })

    return dedupe_canvas_events(events)


def parse_iso_datetime(iso_str: str):
    """Parse ISO-8601 timestamp into local (YYYY-MM-DD, HH:MM)."""
    if not iso_str:
        return None, "23:59"
    clean = iso_str.strip()
    if len(clean) == 10 and re.match(r"^\d{4}-\d{2}-\d{2}$", clean):
        return clean, "23:59"
    try:
        if clean.endswith("Z"):
            clean = clean[:-1] + "+00:00"
        dt = datetime.datetime.fromisoformat(clean)
        if dt.tzinfo is not None:
            dt = dt.astimezone()
        return dt.strftime("%Y-%m-%d"), dt.strftime("%H:%M")
    except Exception:
        m = re.match(r"^(\d{4}-\d{2}-\d{2})T(\d{2}:\d{2})", iso_str)
        if m:
            return m.group(1), m.group(2)
    return None, "23:59"


# =========================================================
# 1. CANVAS SYNC
# =========================================================
def sync_canvas(integration_id: str, config: dict) -> list:
    """Syncs Canvas via:
    - iCal Feed URL (.ics) if provided (either in icsUrl, domain, or token field)
    - Canvas REST API v1 (/api/v1/planner/items + /api/v1/users/self/upcoming_events) via Access Token
    """
    ics_url = (config.get("icsUrl") or "").strip()
    api_token = (config.get("token") or "").strip()
    base_url = (config.get("domain") or "").strip()

    events = []
    seen_ids = set()
    errors = []

    # Detect if user pasted the .ics link into icsUrl, token, or domain
    potential_ics = None
    for candidate in [ics_url, api_token, base_url]:
        if candidate.startswith("webcal://") or candidate.startswith("http://") or candidate.startswith("https://"):
            if ".ics" in candidate.lower() or "feed" in candidate.lower():
                potential_ics = candidate
                break

    if potential_ics:
        try:
            ics_text = http_get_text(potential_ics)
            feed_events = parse_ics_feed(ics_text, integration_id, default_category="duedates")
            for ev in feed_events:
                if ev["id"] not in seen_ids:
                    seen_ids.add(ev["id"])
                    events.append(ev)
        except Exception as e:
            errors.append(f"iCal feed error: {e}")

    # If api_token is provided and isn't just an ics URL
    if api_token and not (api_token.startswith("http") or api_token.startswith("webcal")):
        target_domain = base_url if (base_url and ".ics" not in base_url.lower() and "feed" not in base_url.lower()) else ""
        if not target_domain:
            target_domain = "canvas.instructure.com"
        if not target_domain.startswith("http"):
            target_domain = "https://" + target_domain
        target_domain = target_domain.rstrip("/")

        headers = {"Authorization": f"Bearer {api_token}"}

        # 1. Fetch Planner Items
        start_date = (datetime.date.today() - datetime.timedelta(days=60)).isoformat()
        end_date = (datetime.date.today() + datetime.timedelta(days=120)).isoformat()
        planner_url = f"{target_domain}/api/v1/planner/items?start_date={start_date}&end_date={end_date}&per_page=100"

        try:
            items = http_get_json(planner_url, headers=headers)
            if isinstance(items, list):
                for item in items:
                    plannable = item.get("plannable") or {}
                    title = plannable.get("title") or item.get("context_name") or "Canvas Item"
                    if is_canvas_excluded(title, context_name):
                        continue
                    raw_dt = item.get("plannable_date") or plannable.get("due_at") or plannable.get("todo_date")
                    date_str, time_str = parse_iso_datetime(raw_dt)
                    if not date_str:
                        continue
                    context_name = item.get("context_name") or "Canvas"
                    submitted = bool((item.get("submissions") or {}).get("submitted"))
                    uid = f"{integration_id}-planner-{item.get('plannable_id') or title}"
                    if uid in seen_ids:
                        continue
                    seen_ids.add(uid)
                    if integration_id == "vva" and is_vva_excluded(title, context_name):
                        continue
                    events.append({
                        "id": uid,
                        "title": title,
                        "category": classify_student_category(title, context_name, "duedates", integration_id),
                        "integration": integration_id,
                        "date": date_str,
                        "time": time_str,
                        "location": context_name,
                        "completed": submitted
                    })
        except Exception as e:
            errors.append(f"Planner API: {e}")

        # 2. Fetch Upcoming Events endpoint
        upcoming_url = f"{target_domain}/api/v1/users/self/upcoming_events"
        try:
            upcoming = http_get_json(upcoming_url, headers=headers)
            if isinstance(upcoming, list):
                for ev in upcoming:
                    title = ev.get("title") or "Course Event"
                    if is_canvas_excluded(title, ev.get("description") or ""):
                        continue
                    assignment = ev.get("assignment") or {}
                    raw_dt = assignment.get("due_at") or ev.get("start_at") or ev.get("end_at")
                    date_str, time_str = parse_iso_datetime(raw_dt)
                    if not date_str:
                        continue
                    uid = f"{integration_id}-up-{ev.get('id') or title}"
                    if uid in seen_ids:
                        continue
                    seen_ids.add(uid)
                    if integration_id == "vva" and is_vva_excluded(title, ev.get("description") or ""):
                        continue
                    events.append({
                        "id": uid,
                        "title": title,
                        "category": classify_student_category(title, ev.get("description") or "", "duedates", integration_id),
                        "integration": integration_id,
                        "date": date_str,
                        "time": time_str,
                        "location": ev.get("context_name") or integration_id.upper(),
                        "completed": False
                    })
        except Exception as e:
            errors.append(f"Upcoming API: {e}")

    if not events:
        if errors:
            raise RuntimeError(" | ".join(errors))
        if not potential_ics and not api_token:
            raise ValueError(f"{integration_id.upper()} requires either a Calendar Feed (.ics) URL or a Canvas API Access Token.")
        raise RuntimeError(f"Connected to {integration_id.upper()}, but found 0 events within the date range.")

    return dedupe_canvas_events(events)


# =========================================================
# EMAIL DATE EXTRACTOR (Used by Gmail & Outlook)
# =========================================================
DATE_PATTERNS = [
    # YYYY-MM-DD
    re.compile(r"\b(202\d-\d{2}-\d{2})\b"),
    # Month DD, YYYY or Month DD
    re.compile(
        r"\b(Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|"
        r"Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?)"
        r"\.?\s+(\d{1,2})(?:st|nd|rd|th)?(?:,?\s+(202\d))?\b",
        re.IGNORECASE
    ),
    # MM/DD/YYYY
    re.compile(r"\b(\d{1,2})/(\d{1,2})/(202\d)\b")
]

MONTH_MAP = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "may": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "oct": 10, "nov": 11, "dec": 12
}


def extract_date_and_time_from_text(text: str, fallback_date: str = None):
    """Extract a mentioned deadline/event date and time from an email subject or body."""
    time_match = re.search(r"\b(\d{1,2}):(\d{2})\s*(am|pm)?\b", text, re.IGNORECASE)
    extracted_time = "15:00"
    if time_match:
        hh = int(time_match.group(1))
        mm = int(time_match.group(2))
        ampm = (time_match.group(3) or "").lower()
        if ampm == "pm" and hh < 12:
            hh += 12
        elif ampm == "am" and hh == 12:
            hh = 0
        if 0 <= hh < 24 and 0 <= mm < 60:
            extracted_time = f"{hh:02d}:{mm:02d}"

    # Check YYYY-MM-DD
    m1 = DATE_PATTERNS[0].search(text)
    if m1:
        return m1.group(1), extracted_time

    # Check Month DD, YYYY
    m2 = DATE_PATTERNS[1].search(text)
    if m2:
        mon_str, day_str, year_str = m2.groups()
        mon = MONTH_MAP.get(mon_str[:3].lower(), 1)
        day = int(day_str)
        year = int(year_str) if year_str else datetime.date.today().year
        try:
            return datetime.date(year, mon, day).isoformat(), extracted_time
        except ValueError:
            pass

    # Check MM/DD/YYYY
    m3 = DATE_PATTERNS[2].search(text)
    if m3:
        mon, day, year = map(int, m3.groups())
        try:
            return datetime.date(year, mon, day).isoformat(), extracted_time
        except ValueError:
            pass

    return fallback_date, extracted_time


def decode_mime_header(raw_header: str) -> str:
    if not raw_header:
        return ""
    parts = decode_header(raw_header)
    decoded = []
    for content, enc in parts:
        if isinstance(content, bytes):
            decoded.append(content.decode(enc or "utf-8", errors="replace"))
        else:
            decoded.append(str(content))
    return "".join(decoded).strip()


def sync_imap_mailbox(imap_host: str, username: str, password: str, integration_id: str) -> list:
    """Connects via IMAP4_SSL to Gmail (imap.gmail.com) or Outlook (outlook.office365.com)
    and extracts calendar invites (.ics attachments) + dates mentioned in recent emails.
    """
    events = []
    mail = imaplib.IMAP4_SSL(imap_host, 993)
    try:
        mail.login(username, password)
        mail.select("INBOX", readonly=True)
        status, data = mail.search(None, "ALL")
        if status != "OK" or not data or not data[0]:
            return dedupe_canvas_events(events)

        msg_ids = data[0].split()[-25:]  # Inspect latest 25 emails
        for mid in reversed(msg_ids):
            res, msg_data = mail.fetch(mid, "(RFC822)")
            if res != "OK" or not msg_data or not msg_data[0]:
                continue
            raw_email = msg_data[0][1]
            msg = email.message_from_bytes(raw_email)
            subject = decode_mime_header(msg.get("Subject", ""))
            sender = decode_mime_header(msg.get("From", ""))

            body_text = ""
            for part in msg.walk():
                ctype = part.get_content_type()
                if ctype == "text/calendar":
                    ics_payload = part.get_payload(decode=True)
                    if ics_payload:
                        events.extend(parse_ics_feed(ics_payload.decode("utf-8", errors="replace"), integration_id, "email"))
                elif ctype == "text/plain" and not body_text:
                    payload = part.get_payload(decode=True)
                    if payload:
                        body_text = payload.decode("utf-8", errors="replace")[:600]

            combined = f"{subject} {body_text}"
            if any(k in combined.lower() for k in ["due", "deadline", "meeting", "rsvp", "club", "practice", "mentor", "schedule", "invite", "exam"]):
                dt_str, tm_str = extract_date_and_time_from_text(combined)
                if dt_str:
                    events.append({
                        "id": f"imap-{integration_id}-{mid.decode()}",
                        "title": subject or "Email Event",
                        "category": classify_student_category(subject, body_text, "email"),
                        "integration": integration_id,
                        "date": dt_str,
                        "time": tm_str,
                        "location": f"Email from {sender[:32]}",
                        "completed": False
                    })
    finally:
        try:
            mail.logout()
        except Exception:
            pass

    return dedupe_canvas_events(events)


# =========================================================
# 3. GMAIL SYNC (Google APIs + IMAP + iCal)
# =========================================================
def sync_gmail(config: dict, integration_id: str = "gmail") -> list:
    ics_url = (config.get("icsUrl") or "").strip()
    token = (config.get("token") or "").strip()
    username = (config.get("username") or "").strip()
    password = (config.get("password") or "").strip()

    events = []

    # 1. Google Calendar Secret iCal (.ics) URL if provided
    if ics_url.startswith("http"):
        ics_text = http_get_text(ics_url)
        events.extend(parse_ics_feed(ics_text, "gmail", default_category="email"))

    # 2. Google OAuth2 Bearer Token (Google Calendar API + Gmail REST API)
    if token and not token.startswith("••"):
        headers = {"Authorization": f"Bearer {token}"}
        time_min = (datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(days=14)).isoformat()
        cal_url = f"https://www.googleapis.com/calendar/v3/calendars/primary/events?timeMin={urllib.parse.quote(time_min)}&singleEvents=true&orderBy=startTime&maxResults=50"
        try:
            cal_data = http_get_json(cal_url, headers=headers)
            for item in cal_data.get("items", []):
                title = item.get("summary") or "Google Calendar Event"
                start_obj = item.get("start") or {}
                raw_dt = start_obj.get("dateTime") or start_obj.get("date")
                date_str, time_str = parse_iso_datetime(raw_dt)
                if not date_str:
                    continue
                events.append({
                    "id": f"gcal-{item.get('id')}",
                    "title": title,
                    "category": classify_student_category(title, item.get("description") or "", "email"),
                    "integration": integration_id,
                    "date": date_str,
                    "time": time_str,
                    "location": item.get("location") or "Google Calendar",
                    "completed": False
                })
        except Exception:
            pass

        # Also query Gmail messages for date extraction
        gmail_q = urllib.parse.quote("deadline OR due OR RSVP OR meeting OR club OR practice")
        msg_list_url = f"https://gmail.googleapis.com/gmail/v1/users/me/messages?q={gmail_q}&maxResults=12"
        try:
            msg_list = http_get_json(msg_list_url, headers=headers)
            for m_ref in msg_list.get("messages", []):
                m_detail = http_get_json(f"https://gmail.googleapis.com/gmail/v1/users/me/messages/{m_ref['id']}?format=metadata&metadataHeaders=Subject&metadataHeaders=From", headers=headers)
                snippet = m_detail.get("snippet") or ""
                hdrs = {h["name"].lower(): h["value"] for h in (m_detail.get("payload", {}).get("headers") or [])}
                subj = hdrs.get("subject") or snippet[:50]
                dt_str, tm_str = extract_date_and_time_from_text(f"{subj} {snippet}")
                if dt_str:
                    events.append({
                        "id": f"gmail-msg-{m_ref['id']}",
                        "title": subj,
                        "category": classify_student_category(subj, snippet, "email"),
                        "integration": integration_id,
                        "date": dt_str,
                        "time": tm_str,
                        "location": "Gmail Inbox",
                        "completed": False
                    })
        except Exception:
            pass

    # 3. Direct IMAP Login (imap.gmail.com) with Gmail + App Password
    if username and password and not password.startswith("••") and not events:
        events.extend(sync_imap_mailbox("imap.gmail.com", username, password, "gmail"))

    if not events and not ics_url and not token and not password:
        raise ValueError("Gmail requires a Google OAuth token, a Gmail App Password (IMAP), or a Google Calendar .ics URL.")

    return dedupe_canvas_events(events)


# =========================================================
# 4 & 5. OUTLOOK & MICROSOFT TEAMS SYNC (Microsoft Graph + IMAP + iCal)
# =========================================================
def sync_outlook_or_teams(provider_type: str, config: dict, integration_id: str = None) -> list:
    integration_id = integration_id or provider_type
    ics_url = (config.get("icsUrl") or "").strip()
    token = (config.get("token") or "").strip()
    username = (config.get("username") or "").strip()
    password = (config.get("password") or "").strip()

    default_cat = "mentorship" if integration_id == "teams" else "email"
    events = []

    # 1. Published Outlook / Exchange / Teams iCal (.ics) Feed URL
    if ics_url.startswith("http"):
        ics_text = http_get_text(ics_url)
        events.extend(parse_ics_feed(ics_text, integration_id, default_category=default_cat))

    # 2. Microsoft Graph API Bearer Token
    if token and not token.startswith("••"):
        headers = {
            "Authorization": f"Bearer {token}",
            "Prefer": 'outlook.timezone="UTC"'
        }
        start_dt = (datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(days=14)).strftime("%Y-%m-%dT%H:%M:%SZ")
        end_dt = (datetime.datetime.now(datetime.timezone.utc) + datetime.timedelta(days=60)).strftime("%Y-%m-%dT%H:%M:%SZ")
        cal_url = f"https://graph.microsoft.com/v1.0/me/calendarView?startDateTime={start_dt}&endDateTime={end_dt}&$top=50"

        try:
            cal_data = http_get_json(cal_url, headers=headers)
            for item in cal_data.get("value", []):
                is_teams = bool(item.get("isOnlineMeeting")) or "teams" in str(item.get("onlineMeetingProvider") or "").lower()
                # If syncing specifically Teams, prioritize online meetings or Teams items
                if integration_id == "teams" and not is_teams and "team" not in (item.get("subject") or "").lower():
                    continue

                title = item.get("subject") or "Microsoft Event"
                start_raw = (item.get("start") or {}).get("dateTime")
                if start_raw and not start_raw.endswith("Z"):
                    start_raw += "Z"
                date_str, time_str = parse_iso_datetime(start_raw)
                if not date_str:
                    continue

                loc = ((item.get("location") or {}).get("displayName")) or ("Microsoft Teams" if is_teams else "Outlook Calendar")
                body_preview = item.get("bodyPreview") or ""
                events.append({
                    "id": f"msgraph-{integration_id}-{item.get('id', title)[:16]}",
                    "title": title,
                    "category": classify_student_category(title, f"{loc} {body_preview}", default_cat),
                    "integration": integration_id,
                    "date": date_str,
                    "time": time_str,
                    "location": loc,
                    "completed": False
                })
        except Exception:
            pass

        # If Outlook, also inspect recent school emails via Graph /me/messages
        if integration_id == "outlook":
            mail_url = "https://graph.microsoft.com/v1.0/me/messages?$top=15&$select=id,subject,bodyPreview,from"
            try:
                mail_data = http_get_json(mail_url, headers=headers)
                for msg in mail_data.get("value", []):
                    subj = msg.get("subject") or ""
                    preview = msg.get("bodyPreview") or ""
                    dt_str, tm_str = extract_date_and_time_from_text(f"{subj} {preview}")
                    if dt_str:
                        events.append({
                            "id": f"msmail-{msg.get('id', subj)[:14]}",
                            "title": subj,
                            "category": classify_student_category(subj, preview, "email"),
                            "integration": integration_id,
                            "date": dt_str,
                            "time": tm_str,
                            "location": "Outlook Inbox",
                            "completed": False
                        })
            except Exception:
                pass

    # 3. Direct IMAP Login (outlook.office365.com)
    if username and password and not password.startswith("••") and not events:
        events.extend(sync_imap_mailbox("outlook.office365.com", username, password, integration_id))

    if not events and not ics_url and not token and not password:
        raise ValueError(f"{integration_id.upper()} requires a Microsoft Graph OAuth token, an Outlook/Exchange .ics URL, or IMAP credentials.")

    return dedupe_canvas_events(events)


# =========================================================
# HTTP REQUEST HANDLER
# =========================================================
class CalendarServerHandler(http.server.SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=str(BASE_DIR), **kwargs)

    def _send_json(self, status_code: int, payload: dict):
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status_code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.send_header("Access-Control-Allow-Origin", "*")
        self.end_headers()
        self.wfile.write(body)

    def do_OPTIONS(self):
        self.send_response(204)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "GET, POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type, Authorization")
        self.end_headers()

    def do_GET(self):
        if self.path == "/api/health":
            self._send_json(200, {
                "status": "ok",
                "integrations": ["canvas", "gmail", "outlook", "teams"]
            })
            return
        super().do_GET()

    def do_POST(self):
        if self.path == "/api/sync":
            content_len = int(self.headers.get("Content-Length", "0"))
            raw_body = self.rfile.read(content_len).decode("utf-8", errors="replace")
            try:
                body = json.loads(raw_body) if raw_body else {}
            except json.JSONDecodeError:
                self._send_json(400, {"ok": False, "error": "Invalid JSON body"})
                return

            integrations_cfg = body.get("integrations") or {}
            synced_events = []
            statuses = {}

            for int_id, cfg in integrations_cfg.items():
                if not isinstance(cfg, dict):
                    continue
                if not cfg.get("connected"):
                    statuses[int_id] = {"ok": False, "skipped": True, "count": 0, "message": "Disabled"}
                    continue

                provider_type = cfg.get("type") or int_id
                try:
                    if provider_type == "canvas" or provider_type.startswith("canvas"):
                        evs = sync_canvas(int_id, cfg)
                    elif provider_type == "gmail":
                        evs = sync_gmail(cfg, integration_id=int_id)
                    elif provider_type in ("outlook", "teams"):
                        evs = sync_outlook_or_teams(provider_type, cfg, integration_id=int_id)
                    else:
                        evs = []

                    synced_events.extend(evs)
                    statuses[int_id] = {
                        "ok": True,
                        "count": len(evs),
                        "message": f"Synced {len(evs)} live events"
                    }
                except Exception as exc:
                    statuses[int_id] = {
                        "ok": False,
                        "count": 0,
                        "message": str(exc)
                    }

            self._send_json(200, {
                "ok": True,
                "events": synced_events,
                "statuses": statuses
            })
            return

        self._send_json(404, {"ok": False, "error": "Unknown endpoint"})


if __name__ == "__main__":
    http.server.ThreadingHTTPServer.allow_reuse_address = True
    server = http.server.ThreadingHTTPServer(("0.0.0.0", PORT), CalendarServerHandler)
    print(f"Unified Student Calendar server running at http://localhost:{PORT}")
    server.serve_forever()
