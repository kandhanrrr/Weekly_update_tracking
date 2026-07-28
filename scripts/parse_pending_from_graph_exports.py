import argparse
import configparser
import hashlib
import json
import re
import smtplib
from collections import Counter
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Any


KEYWORDS = re.compile(
    r"\b(AR|Action item|Action items|Follow-up|Follow up|ETA|Due|Pending|Need to|Please update|Owner)\b",
    re.IGNORECASE,
)
ACTION_HINTS = re.compile(
    r"\b(update|updates|change|changes|review|confirm|check|please|need to|what is update|follow up|follow-up)\b",
    re.IGNORECASE,
)
TEAMS_ACTIONABLE = re.compile(
    r"\b(what is update|need to|can (you|she|he)\b|please\b|follow up|follow-up|action item|send email|check with|confirm)\b",
    re.IGNORECASE,
)
TEAMS_NON_ACTION = re.compile(
    r"\b(sure|ok|yes|will join|not joining|already updated|systemeventmessage|meeting has reached its midpoint|next steps for the remaining time)\b",
    re.IGNORECASE,
)
OUTLOOK_ACTION_HINTS = re.compile(
    r"\b(action required|request for|please|need to|follow-up|follow up|review|what is update|can you|pending)\b",
    re.IGNORECASE,
)
DIRECT_REQUEST = re.compile(
    r"\b(can you|could you|please|kindly|would you|need you to|need to|request|share|provide|send|follow up|follow-up|confirm|review|update)\b",
    re.IGNORECASE,
)
CLOSED = re.compile(
    r"\b(Done|Closed|Dropped|N/A|NA|Completed|Review Done/Closed)\b",
    re.IGNORECASE,
)
NOISE = re.compile(
    r"\b(Canceled:|Cancelled:|wants to access|sent \d+ messages to your chat|no-reply@sharepointonline\.com)\b",
    re.IGNORECASE,
)
IGNORE_SUBJECT = re.compile(
    r"(^\s*AR\s*-\s*Central RTL models\b)|(^\s*cancel(?:ed|led):)",
    re.IGNORECASE,
)
NOISE_SENDER = re.compile(
    r"(no-reply@teams\.mail\.microsoft|the\.download@intel\.com|AI\.for\.Everyday\.Engineering\.Communications@intel\.com|hello@medibuddy\.in|noreply@everbridge\.net|Intel\.Access\.Governance\.System@intel\.com)",
    re.IGNORECASE,
)
OUTLOOK_NOISE_SUBJECT = re.compile(
    r"\b(MM:|Weekly Review|Internal Sync|Office Hour|Heads up notification|The Download|INTEL TEST Notification|Pharmacy Refund Initiated|Meeting Minutes|meeting minutes|your access has been|access has been removed)\b"
    r"|^(Automatic reply|Auto:|Out of Office|OOO:|Autosvar|Automatisch antwoord|Abwesenheitsnotiz|Risposta automatica)\b",
    re.IGNORECASE,
)
OWNER_RX = re.compile(r"(?im)Owner\s*[:\-]\s*([^\r\n,;]+)")
ETA_WW = re.compile(r"\bWW\s*(\d{1,2})(?:[\./]?(\d))?(?:'\d{2})?\b", re.IGNORECASE)
ETA_BY = re.compile(
    r"\bby\b.{0,50}?(WW\s*\d{1,2}(?:[\./]?\d)?(?:'\d{2})?|\d{1,2}[\-/]\d{1,2}[\-/]\d{2,4}|(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\.?\s+\d{1,2})",
    re.IGNORECASE,
)
ETA_Q = re.compile(r"\b(Q[1-4](?:[/\-]Q?[1-4])?(?:\s*['\'']?\d{2,4})?)\b", re.IGNORECASE)
GREETING_NAME = re.compile(
    r"(?:^|\n)\s*(?:Hi|Hello|Dear)\s+([A-Z][a-z]+(?:\s+[A-Z][a-z]+)?)\s*[,\n<]",
    re.MULTILINE,
)
SELF_FOLLOWUP = re.compile(
    r"\b(can you|could you|please share|please send|give (?:the|me)|share (?:the|with|me)|send me|by EOD|by tomorrow|atleast share|share the (?:pre|post|data|excel|report|aqua)|update on|any update|have you|is it done|follow.?up)\b",
    re.IGNORECASE,
)
# Reply convincingly resolves the ask (shared, done, sent, etc.)
REPLY_COMPLETE = re.compile(
    r"\b(done|completed|shared|sent|sending|sharing(?: now| the| it)?|here (is|are)|please find|attached|uploaded|updated|forwarding|as requested|"
    r"i(?:'ve| have) (?:sent|shared|uploaded|updated|completed|done)|just (?:shared|sent)|will share|will send)\b",
    re.IGNORECASE,
)
# Reply acknowledges but does not complete the ask
REPLY_PENDING = re.compile(
    r"\b(will (?:check|do|review|get back|look into|follow)|let me (?:check|look|verify|do)|noted|on it|working on|i'?ll|will get back|ok|sure|ok sure)\b",
    re.IGNORECASE,
)


_HEADER_LINE_RX = re.compile(
    r"(?m)^\s*(?:From|To|Cc|Bcc|Sent|Subject|Date):\s*.*$",
    re.IGNORECASE,
)


def _strip_email_headers(text: str) -> str:
    """Remove embedded email header lines (To:, From:, Cc:, Sent: …) so recipient
    names in the To/Cc fields are not mistaken for greeting targets."""
    return _HEADER_LINE_RX.sub("", text).strip()


def _parse_greeting_name(text: str) -> str:
    m = GREETING_NAME.search(text)
    if m:
        return m.group(1).strip()
    return ""


@dataclass
class Row:
    source: str
    date: datetime
    task: str
    owner: str
    status: str
    eta: str
    sender: str
    summary: str = ""
    assigned_to: str = ""  # who must complete the task


def _load_identity() -> tuple[str, set[str], str, str, re.Pattern]:
    """Read identity from config.ini and build dynamic name-matching regex.

    Returns
    -------
    self_email          : lower-cased email address
    self_aliases        : set of lowercase name/email aliases for self-detection
    first_name          : display first name (e.g. "Kandhan")
    last_name           : display last name  (e.g. "Rajakumar")
    addressed_to_self_rx: compiled regex that matches when an email/message is
                          directly addressed to the user by name
    """
    config = configparser.ConfigParser()
    config.read(Path(__file__).resolve().parents[1] / "config.ini", encoding="utf-8")
    self_email = config.get("EMAIL", "smtp_user", fallback="").strip().lower()

    # Read explicit identity; fall back to deriving from email local-part.
    first_name = config.get("IDENTITY", "first_name", fallback="").strip()
    last_name = config.get("IDENTITY", "last_name", fallback="").strip()
    if not first_name and self_email:
        parts = self_email.split("@", 1)[0].split(".")
        first_name = parts[0].capitalize() if parts else ""
        last_name = parts[1].capitalize() if len(parts) > 1 else ""

    aliases: set[str] = set()
    if self_email:
        aliases.add(self_email)
        local = self_email.split("@", 1)[0].replace(".", " ")
        aliases.add(local)
        aliases.add(local.replace(" ", ""))
    if first_name:
        aliases.add(first_name.lower())
    if first_name and last_name:
        aliases.add(f"{first_name.lower()} {last_name.lower()}")
        aliases.add(f"{last_name.lower()} {first_name.lower()}")

    # Build the "addressed to self" regex dynamically.
    patterns: list[str] = []
    if first_name:
        fn = re.escape(first_name)
        patterns += [
            rf"(?:Hi|Hello|Dear)\s+{fn}\b",
            rf"@\s*{fn}\b",
            rf"\b{fn}[,\s]+please\b",
        ]
    if first_name and last_name:
        fn = re.escape(first_name)
        ln = re.escape(last_name)
        patterns.append(rf"\b{ln},?\s+{fn}\b")
    addressed_to_self_rx = (
        re.compile("|".join(patterns), re.IGNORECASE)
        if patterns
        else re.compile(r"(?!)")  # never-match fallback
    )

    return (
        self_email,
        {a.strip() for a in aliases if a.strip()},
        first_name or "Me",
        last_name or "",
        addressed_to_self_rx,
    )


def _contains_self_mention(text: str, self_aliases: set[str]) -> bool:
    lowered = text.lower()
    return any(
        alias and re.search(rf"\b{re.escape(alias)}\b", lowered)
        for alias in self_aliases
    )


def _looks_like_direct_request(text: str) -> bool:
    return bool(DIRECT_REQUEST.search(text))


def _extract_eta(text: str) -> str:
    """Try to pull a concrete deadline from free text; return TBD when none found."""
    m = ETA_BY.search(text)
    if m:
        return m.group(1).strip() if m.lastindex and m.group(1) else m.group(0).strip()
    m = ETA_WW.search(text)
    if m:
        ww = m.group(1)
        patch = m.group(2) or ""
        return f"WW{ww}{'.' + patch if patch else ''}"
    m = ETA_Q.search(text)
    if m:
        return m.group(1).strip()
    return "TBD"


_HTML_ENTITIES = [
    ("&nbsp;", " "), ("&amp;", "&"), ("&lt;", "<"), ("&gt;", ">"),
    ("&quot;", '"'), ("&#39;", "'"), ("&apos;", "'"),
]


def _clean_html(text: str) -> str:
    """Strip HTML tags and decode common entities, collapse whitespace."""
    t = re.sub(r"<[^>]+>", " ", text or "")
    for ent, ch in _HTML_ENTITIES:
        t = t.replace(ent, ch)
    return re.sub(r"\s+", " ", t).strip()


# Patterns to strip before building the professional context summary
_SUMMARY_STRIP = re.compile(
    r"(?:^|\n)\s*(?:"
    r"(?:Hi|Hello|Dear)\s+[\w ,]+[,\n]"  # greetings
    r"|(?:Thanks?(?:\s+and)?\s+)?Regards[,\s]+[\w ]+"  # sign-offs
    r"|From:\s+.*"  # forwarded-from header and everything after
    r"|Sent:\s+.*"  # forwarded-sent line
    r"|To:\s+.*"  # To: line
    r"|Cc:\s+.*"  # Cc: line
    r"|_{4,}.*"  # separator lines
    r"|Microsoft Teams meeting.*"  # meeting boilerplate
    r")",
    re.IGNORECASE | re.MULTILINE | re.DOTALL,
)


def _make_summary(preview: str, max_sentences: int = 2) -> str:
    """Return a professional 2-line context summary stripped of greetings, sign-offs, and forward headers."""
    # Strip HTML tags and entities
    text = _clean_html(preview or "")
    # Remove greetings at start
    text = re.sub(
        r"^\s*(?:Hi|Hello|Dear)\s+[\w ,]+[,\.\n]",
        "",
        text,
        flags=re.IGNORECASE,
    ).strip()
    # Remove trailing sign-offs / forward headers
    text = re.sub(
        r"(?:Thanks?(?:\s+(?:and\s+)?Regards)?|Best\s+Regards?|Regards?)[,\s]+[\w ]+$",
        "",
        text,
        flags=re.IGNORECASE,
    ).strip()
    text = re.sub(
        r"(?:From|Sent|To|Cc):\s+.*$", "", text, flags=re.IGNORECASE
    ).strip()
    text = re.sub(r"_{4,}.*$", "", text, flags=re.DOTALL).strip()
    text = re.sub(
        r"Microsoft Teams meeting.*$",
        "",
        text,
        flags=re.IGNORECASE | re.DOTALL,
    ).strip()
    # Split on sentence boundaries
    parts = re.split(r"(?<=[.!?])\s+", text)
    sentences: list[str] = []
    for p in parts:
        p = p.strip()
        if not p or len(p) < 12:
            continue
        sentences.append(p)
        if len(sentences) >= max_sentences:
            break
    result = " ".join(sentences) if sentences else text[:240]
    # Cap at 240 chars for table readability
    if len(result) > 240:
        result = result[:237] + "..."
    return result


def _parse_iso(date_text: str) -> datetime:
    value = date_text.strip().replace("Z", "+00:00")
    # Normalize fractional seconds to 6 digits for Python 3.10 compatibility.
    m = re.match(r"^(.*T\d{2}:\d{2}:\d{2})(\.(\d+))?([+-]\d{2}:\d{2})$", value)
    if m:
        base = m.group(1)
        frac = m.group(3) or ""
        tz = m.group(4)
        if frac:
            frac = (frac + "000000")[:6]
            value = f"{base}.{frac}{tz}"
        else:
            value = f"{base}{tz}"
    dt = datetime.fromisoformat(value)
    if dt.tzinfo is None:
        return dt.replace(tzinfo=timezone.utc)
    return dt.astimezone(timezone.utc)


def _extract_outlook_rows(
    payload: dict[str, Any],
    since_utc: datetime,
    exclude_owners: set[str],
    strict: bool,
    self_email: str,
    self_aliases: set[str],
    first_name: str = "Me",
    last_name: str = "",
    addressed_to_self_rx: re.Pattern | None = None,
    search_mode: bool = False,
) -> list[Row]:
    if addressed_to_self_rx is None:
        addressed_to_self_rx = re.compile(r"(?!)")  # never-match fallback
    rows: list[Row] = []
    for e in payload.get("emails", []):
        date_text = str(e.get("receivedDateTime", "")).strip()
        if not date_text:
            continue
        date = _parse_iso(date_text)
        if date < since_utc:
            continue
        text = f"{e.get('subject', '')}\n{e.get('preview', '')}"
        subject = str(e.get("subject", "")).strip()
        sender = str(e.get("from", "")) or str(e.get("sender", ""))
        if self_email and sender.strip().lower() == self_email:
            continue
        if IGNORE_SUBJECT.search(subject):
            continue
        if OUTLOOK_NOISE_SUBJECT.search(subject):
            continue
        if sender.lower().startswith("modelcleaner@"):
            continue
        if NOISE_SENDER.search(sender):
            continue
        if "@intel.com" not in sender.lower() and "@ecsmtp." not in sender.lower():
            continue
        action_text = subject + " " + str(e.get("preview", ""))
        body_text = str(e.get("preview", ""))  # available in both search_mode and normal path
        if search_mode:
            # Email was retrieved via keyword search — it already matches AR criteria.
            # Only verify it is a genuine request (not an FYI or closed status).
            directly_named, in_to = True, True  # treat as explicitly addressed
            if not _looks_like_direct_request(action_text):
                continue
        else:
            to_recipients = {
                str(item).strip().lower()
                for item in (e.get("to", []) or [])
            }
            cc_recipients = {
                str(item).strip().lower()
                for item in (e.get("cc", []) or [])
            }

            directly_named = bool(addressed_to_self_rx.search(body_text))
            in_to = bool(self_email and self_email in to_recipients)
            in_cc_only = (
                bool(self_email and self_email in cc_recipients)
                and not in_to
                and not directly_named
            )

            # If Kandhan is only CC'd and not explicitly named → informational, skip.
            if in_cc_only:
                continue

            # Keep only if addressed to the user: in TO field, or explicitly named in body.
            addressed_to_self = (
                in_to
                or directly_named
                or _contains_self_mention(text, self_aliases)
            )
            if not addressed_to_self:
                continue
            # Must also be a genuine request, not just an FYI.
            if not _looks_like_direct_request(action_text):
                continue
        if strict:
            if not OUTLOOK_ACTION_HINTS.search(action_text):
                continue
        elif not (KEYWORDS.search(text) or OUTLOOK_ACTION_HINTS.search(action_text)):
            continue
        if CLOSED.search(text):
            continue
        if NOISE.search(text):
            continue
        # Ownership: determine from the actual greeting in the body, not just TO-field membership.
        # Rule: "Hi Kandhan" / @Kandhan / ++Kandhan / "Hi All" → Kandhan
        #        "Hi PK" / "Hi Thomas" → that person
        #        No greeting → TBD
        # Strip embedded email headers (To:, From:, Cc:…) first so recipient names
        # in the header list don't pollute the greeting/addressee check.
        body_clean = _strip_email_headers(body_text)
        _greeting = _parse_greeting_name(body_clean)
        _addressed_all = bool(re.search(r"(?:Hi|Hello|Dear)\s+All\b", body_clean, re.IGNORECASE))
        _explicit_self = bool(addressed_to_self_rx.search(body_clean))
        _name_keys = tuple(k for k in (first_name.lower(), last_name.lower()) if k)
        _greeting_is_self = bool(
            _greeting and _name_keys and any(k in _greeting.lower() for k in _name_keys)
        )
        if _explicit_self or _greeting_is_self:
            assigned_to = first_name
        elif _addressed_all:
            assigned_to = first_name
        elif _greeting:
            assigned_to = _greeting  # e.g. "PK", "Thomas", "Harsh"
        else:
            assigned_to = "TBD"
        owner = "TBD"
        m = OWNER_RX.search(text)
        if m:
            owner = m.group(1).strip()
        else:
            s = str(e.get("subject", ""))
            m2 = re.match(r"\s*([A-Za-z]+,\s*[A-Za-z][A-Za-z0-9\-]*)", s)
            if m2:
                owner = m2.group(1)
        if owner.strip().lower() in exclude_owners:
            continue
        preview = str(e.get("preview", ""))
        rows.append(
            Row(
                source="Outlook",
                date=date,
                task=subject or "(no subject)",
                owner=owner,
                assigned_to=assigned_to,
                status="Open",
                eta=_extract_eta(preview),
                sender=sender,
                summary=_make_summary(preview),
            )
        )
    return rows


def _extract_teams_rows(
    payload: dict[str, Any],
    since_utc: datetime,
    exclude_owners: set[str],
    strict: bool,
    self_aliases: set[str],
    first_name: str = "Me",
) -> list[Row]:
    rows: list[Row] = []
    chat_partner = str(payload.get("chatPartnerName", "")).strip()
    # Determine non-self participants so self-sent messages can be assigned correctly.
    other_names: list[str] = []
    for m in payload.get("messages", []):
        dn = str(m.get("fromDisplayName", "")).strip()
        if dn and not _contains_self_mention(dn, self_aliases) and dn not in other_names:
            other_names.append(dn)
    default_assignee = chat_partner or (other_names[0] if len(other_names) == 1 else "TBD")

    # Pre-compute (timestamp, cleaned_text) of Kandhan's own messages in this chat
    # so we can evaluate whether his reply convincingly resolves any incoming AR.
    self_replies: list[tuple[datetime, str]] = []
    for m in payload.get("messages", []):
        dn = str(m.get("fromDisplayName", "")).strip()
        if _contains_self_mention(dn, self_aliases):
            raw_ts = str(m.get("createdDateTime", m.get("lastModifiedDateTime", ""))).strip()
            if raw_ts:
                try:
                    ts = _parse_iso(raw_ts)
                    body = re.sub(r"<[^>]+>", " ", str(m.get("body", "")))
                    cleaned_reply = re.sub(r"\s+", " ", body).strip()
                    self_replies.append((ts, cleaned_reply))
                except ValueError:
                    pass

    def _convincingly_replied(ar_time: datetime) -> bool:
        """Return True if Kandhan's first reply after ar_time looks like a completion.

        - No reply at all → False (still pending)
        - Reply with completion words (done/shared/sent/etc.) → True (resolved)
        - Reply with only acknowledgement (noted/will check/ok) → False (still pending)
        """
        later = [(t, txt) for t, txt in self_replies if t > ar_time]
        if not later:
            return False
        # Evaluate the first (earliest) reply
        first_reply_txt = min(later, key=lambda x: x[0])[1]
        if REPLY_COMPLETE.search(first_reply_txt):
            return True
        return False

    for m in payload.get("messages", []):
        date_text = str(m.get("createdDateTime", m.get("lastModifiedDateTime", ""))).strip()
        if not date_text:
            continue
        date = _parse_iso(date_text)
        if date < since_utc:
            continue
        body = str(m.get("body", ""))
        text = body
        if str(m.get("messageType", "")).lower() != "message":
            continue
        if not m.get("fromDisplayName"):
            continue
        cleaned = _clean_html(text)
        if len(cleaned) < 16:
            continue
        if TEAMS_NON_ACTION.search(cleaned):
            continue
        sender_name = str(m.get("fromDisplayName", "")).strip()
        is_from_self = bool(sender_name and _contains_self_mention(sender_name, self_aliases))
        if is_from_self:
            # Include only if Kandhan is explicitly following up / requesting something.
            if not SELF_FOLLOWUP.search(cleaned):
                continue
            # Responsible owner is the chat partner, or parsed from greeting.
            greeting = _parse_greeting_name(cleaned)
            assigned_to = greeting or default_assignee
            owner = assigned_to
        else:
            # If Kandhan's reply convincingly resolved this ask, skip it.
            if _convincingly_replied(date):
                continue
            assigned_to = first_name
            owner = sender_name or "TBD"
        # Accept if the user is named, OR if it is a direct request in a group channel.
        is_addressed = _contains_self_mention(cleaned, self_aliases) or _looks_like_direct_request(cleaned) or is_from_self
        if not is_addressed:
            continue
        if strict:
            if not TEAMS_ACTIONABLE.search(cleaned):
                continue
        elif not (
            KEYWORDS.search(cleaned)
            or ACTION_HINTS.search(cleaned)
            or TEAMS_ACTIONABLE.search(cleaned)
        ):
            continue
        if CLOSED.search(cleaned):
            continue
        owner_final = owner
        mo = OWNER_RX.search(cleaned)
        if mo:
            owner_final = mo.group(1).strip()
        if not is_from_self:
            sender_name_val = str(m.get("fromDisplayName", "")).strip()
            if sender_name_val and _contains_self_mention(sender_name_val, self_aliases):
                continue
            if sender_name_val:
                owner_final = sender_name_val
        if owner_final.strip().lower() in exclude_owners and not is_from_self:
            continue
        task = cleaned
        if len(task) > 120:
            task = task[:117] + "..."
        rows.append(
            Row(
                source="Teams",
                date=date,
                task=task or "(empty message)",
                owner=owner_final,
                assigned_to=assigned_to,
                status="Open",
                eta=_extract_eta(cleaned),
                sender=sender_name or str(m.get("from", "")),
                summary=_make_summary(cleaned),
            )
        )
    return rows


def _load_json(path: Path) -> dict[str, Any]:
    return json.loads(path.read_text(encoding="utf-8"))


def _load_smtp_config() -> dict[str, Any]:
    """Read SMTP settings and [OWNERS] email map from config.ini."""
    config = configparser.ConfigParser()
    config.read(Path(__file__).resolve().parents[1] / "config.ini", encoding="utf-8")
    owner_map: dict[str, str] = {}
    if config.has_section("OWNERS"):
        for name, email in config.items("OWNERS"):
            if email.strip():
                owner_map[name.strip().lower()] = email.strip()
    return {
        "host": config.get("EMAIL", "smtp_host", fallback=""),
        "port": int(config.get("EMAIL", "smtp_port", fallback="25")),
        "user": config.get("EMAIL", "smtp_user", fallback=""),
        "pass": config.get("EMAIL", "smtp_pass", fallback=""),
        "owners": owner_map,
    }


def _resolve_owner_email(owner: str, owner_map: dict[str, str]) -> str | None:
    """Best-effort match of a display name (e.g. 'Anand, RishiX') to an email."""
    key = owner.strip().lower()
    if key in owner_map:
        return owner_map[key]
    # Try each token (last-name or first-name match)
    tokens = re.split(r"[,\s]+", key)
    for tok in tokens:
        if len(tok) < 3:
            continue
        for k, v in owner_map.items():
            if tok in k:
                return v
    return None


def _send_ar_reminders(
    all_rows: list["Row"],
    min_age_days: int,
    smtp_cfg: dict[str, Any],
    from_name: str,
    self_email: str,
    dry_run: bool = True,
) -> None:
    """Send one digest reminder email per responsible owner for ARs older than
    *min_age_days* days that are still pending.  Also sends a full summary
    to the user themselves."""
    cutoff = datetime.now(timezone.utc) - timedelta(days=min_age_days)
    stale = [r for r in all_rows if r.date <= cutoff]
    if not stale:
        print(f"\n[Reminders] No ARs older than {min_age_days} days — nothing to remind.")
        return

    owner_map = smtp_cfg.get("owners", {})

    # Group by responsible owner
    by_owner: dict[str, list["Row"]] = {}
    for r in stale:
        owner_key = (r.assigned_to or r.owner or "TBD").strip()
        by_owner.setdefault(owner_key, []).append(r)

    def _fmt_rows(rows: list["Row"]) -> str:
        lines = []
        for i, r in enumerate(rows, 1):
            src = f"[{r.source}]"
            lines.append(
                f"  {i}. {r.date.strftime('%Y-%m-%d')} {src}\n"
                f"     From   : {r.sender}\n"
                f"     Subject: {r.task}\n"
                f"     Summary: {r.summary or r.task}\n"
            )
        return "\n".join(lines)

    sent_count = 0
    for owner, rows in sorted(by_owner.items()):
        if owner in ("TBD", ""):
            continue  # can't send without a name
        recipient = _resolve_owner_email(owner, owner_map)
        if not recipient:
            print(f"[Reminders] No email found for '{owner}' — skipping.")
            continue

        subject = (
            f"[AR Reminder] {len(rows)} pending action item(s) — please update"
        )
        body = (
            f"Hi {owner.split(',')[-1].strip()},\n\n"
            f"The following action item(s) have been pending for over {min_age_days} day(s)"
            " and still need your attention:\n\n"
            + _fmt_rows(rows)
            + f"\nPlease reply with a status update, share the deliverable, or let "
            f"{from_name} know if more time is needed.\n\n"
            f"Regards,\n{from_name} (via AR Tracker)"
        )

        print(f"\n{'[DRY-RUN] ' if dry_run else ''}REMINDER EMAIL")
        print(f"  To     : {recipient}")
        print(f"  Subject: {subject}")
        print(f"  Items  : {len(rows)}")
        if not dry_run:
            _smtp_send(smtp_cfg, [recipient], subject, body)
        sent_count += 1

    # Summary digest to self
    if self_email and stale:
        all_text = ""
        for src in ("Outlook", "Teams"):
            src_rows = [r for r in stale if r.source == src]
            if src_rows:
                all_text += f"\n--- {src} ({len(src_rows)}) ---\n" + _fmt_rows(src_rows)
        subject_self = f"[AR Summary] {len(stale)} pending ARs older than {min_age_days} days"
        body_self = (
            f"Hi {from_name},\n\n"
            f"Here is your pending AR digest for items older than {min_age_days} days:\n"
            + all_text
            + f"\nReminders sent to {sent_count} owner(s).\n\n"
            f"Regards,\nAR Tracker"
        )
        print(f"\n{'[DRY-RUN] ' if dry_run else ''}SELF DIGEST EMAIL")
        print(f"  To     : {self_email}")
        print(f"  Subject: {subject_self}")
        print(f"  ARs    : {len(stale)} total, {sent_count} reminders sent")
        if not dry_run:
            _smtp_send(smtp_cfg, [self_email], subject_self, body_self)


def _smtp_send(smtp_cfg: dict[str, Any], to: list[str], subject: str, body: str) -> None:
    msg = MIMEMultipart("alternative")
    msg["From"] = smtp_cfg["user"]
    msg["To"] = ", ".join(to)
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain"))
    try:
        with smtplib.SMTP(smtp_cfg["host"], smtp_cfg["port"]) as srv:
            if smtp_cfg["port"] != 25:
                srv.starttls()
            if smtp_cfg.get("pass") and smtp_cfg["port"] != 25:
                srv.login(smtp_cfg["user"], smtp_cfg["pass"])
            srv.sendmail(smtp_cfg["user"], to, msg.as_string())
        print("  [SENT]")
    except Exception as exc:
        print(f"  [ERROR] {exc}")


def _md_table(headers: list[str], data: list[list[str]]) -> str:
    out = []
    out.append("| " + " | ".join(headers) + " |")
    out.append("|" + "|".join(["---"] * len(headers)) + "|")
    for row in data:
        out.append("| " + " | ".join(row) + " |")
    return "\n".join(out)


def main() -> int:
    parser = argparse.ArgumentParser(description="Parse pending AR tasks from Graph export JSON files.")
    parser.add_argument("--outlook-json", default=None, help="Path to Outlook inbox Graph export JSON (optional if --outlook-search-json is provided)")
    parser.add_argument("--outlook-search-json", help="Path to a Graph email search result JSON; emails here are treated as pre-filtered ARs (no recipient check)")
    parser.add_argument("--teams-json", help="Path to a Teams export JSON file or a ';' separated list of paths")
    parser.add_argument("--lookback-days", type=int, default=7)
    parser.add_argument("--top", type=int, default=15)
    parser.add_argument("--exclude-owner", action="append", default=[], help="Owner/display name to exclude; can be passed multiple times")
    parser.add_argument("--strict", action="store_true", help="Only keep high-confidence actionable asks")
    parser.add_argument("--send-reminders", action="store_true", help="Enable reminder emails for ARs older than --reminder-min-age-days (requires --send to dispatch)")
    parser.add_argument("--send", action="store_true", help="Confirm dispatch: actually send reminder emails. Without this flag no emails are ever sent.")
    parser.add_argument("--reminder-min-age-days", type=int, default=5, help="Only remind about ARs at least this many days old (default: 5 = prev WW)")
    parser.add_argument("--output-json", default=None, help="Save extracted AR rows as JSON for upsert into Excel tracker")
    args = parser.parse_args()

    since = datetime.now(timezone.utc) - timedelta(days=args.lookback_days)

    exclude_owners = {o.strip().lower() for o in args.exclude_owner if o.strip()}
    self_email, self_aliases, first_name, last_name, addressed_to_self_rx = _load_identity()

    outlook_rows: list = []
    if args.outlook_json:
        outlook_payload = json.loads(Path(args.outlook_json).read_text(encoding="utf-8"))
        outlook_rows = _extract_outlook_rows(
            outlook_payload,
            since,
            exclude_owners,
            args.strict,
            self_email,
            self_aliases,
            first_name=first_name,
            last_name=last_name,
            addressed_to_self_rx=addressed_to_self_rx,
        )
    if args.outlook_search_json:
        for search_path_str in args.outlook_search_json.split(";"):
            search_path_str = search_path_str.strip()
            if not search_path_str:
                continue
            search_payload = json.loads(Path(search_path_str).read_text(encoding="utf-8"))
            outlook_rows += _extract_outlook_rows(
                search_payload,
                since,
                exclude_owners,
                args.strict,
                self_email,
                self_aliases,
                first_name=first_name,
                last_name=last_name,
                addressed_to_self_rx=addressed_to_self_rx,
                search_mode=True,
            )

    teams_rows: list[Row] = []
    teams_scanned = 0
    if args.teams_json:
        team_paths = [Path(p.strip()) for p in args.teams_json.split(";") if p.strip()]
        for p in team_paths:
            if not p.exists():
                continue
            teams_payload = _load_json(p)
            teams_scanned += int(teams_payload.get("count", 0))
            teams_rows.extend(
                _extract_teams_rows(
                    teams_payload,
                    since,
                    exclude_owners,
                    args.strict,
                    self_aliases,
                    first_name=first_name,
                )
            )

    # De-duplicate likely repeats across multiple searches/files.
    seen = set()
    deduped_rows: list[Row] = []
    for r in sorted(outlook_rows + teams_rows, key=lambda x: x.date, reverse=True):
        key = (r.source, r.date.strftime("%Y-%m-%d %H:%M"), r.task[:80], r.sender)
        if key in seen:
            continue
        seen.add(key)
        deduped_rows.append(r)
    all_rows = deduped_rows

    print("# Source scan summary")
    print(
        _md_table(
            ["Metric", "Value"],
            [
                ["Outlook scanned", str(outlook_payload.get("count", 0)) if args.outlook_json else "n/a (search only)"],
                ["Teams scanned", str(teams_scanned)],
                ["Outlook pending extracted", str(len(outlook_rows))],
                ["Teams pending extracted", str(len(teams_rows))],
                ["Total pending extracted", str(len(all_rows))],
            ],
        )
    )
    print()

    print("# Pending tasks by owner")
    by_owner = Counter([r.owner for r in all_rows])
    owner_rows = [[owner, str(count)] for owner, count in by_owner.most_common()]
    if not owner_rows:
        owner_rows = [["(none)", "0"]]
    print(_md_table(["Owner", "Pending Count"], owner_rows))
    print()

    COLS = ["Date", "From", "Task / Subject", "Short Summary", "Responsible Owner", "ETA"]

    def _rows_to_table(rows: list[Row]) -> list[list[str]]:
        return [
            [
                r.date.strftime("%Y-%m-%d"),
                r.sender.replace("|", "/"),
                r.task.replace("|", "/"),
                r.summary.replace("|", "/") if r.summary else "\u2014",
                r.assigned_to or r.owner,
                r.eta,
            ]
            for r in rows
        ]

    outlook_sorted = sorted(
        [r for r in all_rows if r.source == "Outlook"], key=lambda x: x.date, reverse=True
    )[: args.top]
    teams_sorted = sorted(
        [r for r in all_rows if r.source == "Teams"], key=lambda x: x.date, reverse=True
    )[: args.top]

    print(f"# Outlook Action Items ({len(outlook_sorted)})")
    ot_rows = _rows_to_table(outlook_sorted)
    if not ot_rows:
        ot_rows = [["-", "No pending Outlook tasks", "—", "-", "—", "-"]]
    print(_md_table(COLS, ot_rows))
    print()

    print(f"# Teams Action Items ({len(teams_sorted)})")
    tm_rows = _rows_to_table(teams_sorted)
    if not tm_rows:
        tm_rows = [["-", "No pending Teams tasks", "—", "-", "—", "-"]]
    print(_md_table(COLS, tm_rows))
    print()

    with_eta = [r for r in all_rows if r.eta != "TBD"]
    without_eta = [r for r in all_rows if r.eta == "TBD"]
    print("# Reminder outcome summary")
    print(
        _md_table(
            ["Metric", "Value"],
            [
                ["Tasks with parsed ETA", str(len(with_eta))],
                ["Tasks with TBD ETA (needs manual entry)", str(len(without_eta))],
            ],
        )
    )

    if args.output_json:
        def _row_key(r: Row) -> str:
            raw = f"{r.source}|{r.date.strftime('%Y-%m-%d')}|{r.task[:80]}"
            return hashlib.sha1(raw.encode()).hexdigest()[:12]

        json_rows = [
            {
                "source_key": _row_key(r),
                "source": r.source,
                "date": r.date.strftime("%Y-%m-%d"),
                "task": r.task,
                "owner": r.assigned_to or r.owner,
                "status": r.status or "Open",
                "eta": r.eta,
                "sender": r.sender,
                "summary": r.summary,
            }
            for r in all_rows
        ]
        out_path = Path(args.output_json)
        out_path.parent.mkdir(parents=True, exist_ok=True)
        out_path.write_text(json.dumps(json_rows, indent=2, ensure_ascii=False), encoding="utf-8")
        print(f"\n[output-json] {len(json_rows)} row(s) saved to: {out_path}")

    if args.send_reminders:
        if not args.send:
            # Safety gate: never dispatch emails without explicit --send
            cutoff = datetime.now(timezone.utc) - timedelta(days=args.reminder_min_age_days)
            stale_count = sum(1 for r in all_rows if r.date <= cutoff)
            print(
                f"\n[Reminders] {stale_count} AR(s) are older than {args.reminder_min_age_days} days"
                " and eligible for reminders."
            )
            print(
                "[Reminders] No emails sent. To dispatch, run with:"
                " --send-reminders --send"
            )
        else:
            smtp_cfg = _load_smtp_config()
            _send_ar_reminders(
                all_rows,
                min_age_days=args.reminder_min_age_days,
                smtp_cfg=smtp_cfg,
                from_name=first_name,
                self_email=self_email,
                dry_run=False,
            )

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
