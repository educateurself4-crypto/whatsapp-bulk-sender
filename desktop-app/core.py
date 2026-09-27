"""
core.py - business logic for the WhatsApp Bulk Sender desktop app.
No GUI code in here, so it can be unit-tested (see test_core.py).

Responsibilities:
  * read Excel / CSV files
  * normalise + validate phone numbers
  * parse WhatsApp templates (and flag the ones this app cannot fill in)
  * talk to the n8n webhooks
  * run a campaign in small chunks and wait for n8n to confirm each chunk
"""
from __future__ import annotations

import csv
import datetime as dt
import re
import threading
import time
from dataclasses import dataclass, field
from typing import Callable, Optional

import requests

# --------------------------------------------------------------------------
# Excel / CSV reading
# --------------------------------------------------------------------------


def _cell_to_text(v) -> str:
    """Convert any spreadsheet cell value to clean text."""
    if v is None:
        return ""
    if isinstance(v, bool):
        return str(v)
    if isinstance(v, float):
        # Excel stores 9876543210 as a float sometimes; avoid '9876543210.0'
        if v.is_integer():
            return str(int(v))
        return ("%f" % v).rstrip("0").rstrip(".")
    if isinstance(v, int):
        return str(v)
    if isinstance(v, (dt.datetime, dt.date)):
        return v.strftime("%d %b %Y")
    return str(v).strip()


def _dedupe_headers(raw_headers: list[str]) -> list[str]:
    seen: dict[str, int] = {}
    out = []
    for i, h in enumerate(raw_headers, start=1):
        name = h.strip() or f"Column {i}"
        if name in seen:
            seen[name] += 1
            name = f"{name} ({seen[name]})"
        else:
            seen[name] = 1
        out.append(name)
    return out


def read_table(path: str) -> tuple[list[str], list[list[str]]]:
    """Read the first sheet of .xlsx/.xlsm or a .csv. First row = headers."""
    lower = path.lower()
    if lower.endswith((".xlsx", ".xlsm")):
        from openpyxl import load_workbook

        wb = load_workbook(path, read_only=True, data_only=True)
        ws = wb.worksheets[0]
        raw_rows = [[_cell_to_text(c) for c in row] for row in ws.iter_rows(values_only=True)]
        wb.close()
    elif lower.endswith(".csv"):
        with open(path, newline="", encoding="utf-8-sig") as f:
            raw_rows = [[_cell_to_text(c) for c in row] for row in csv.reader(f)]
    else:
        raise ValueError("Unsupported file type. Use .xlsx, .xlsm or .csv")

    # drop fully blank rows
    raw_rows = [r for r in raw_rows if any(c != "" for c in r)]
    if len(raw_rows) < 2:
        raise ValueError("The file needs a header row and at least one data row.")

    headers = _dedupe_headers(raw_rows[0])
    width = len(headers)
    rows = [(r + [""] * width)[:width] for r in raw_rows[1:]]
    return headers, rows


# --------------------------------------------------------------------------
# Phone numbers
# --------------------------------------------------------------------------


def normalize_phone(raw: str, default_cc: str = "91") -> tuple[Optional[str], str]:
    """
    Return (number, error). number is digits only in international format
    without '+', e.g. '919876543210' - the format the WhatsApp Cloud API expects.

    Rules:
      * '+' or '00' prefix -> treated as already international
      * 10 digits          -> default country code is added
      * leading 0 (trunk)  -> removed
      * otherwise 11-15 digits are treated as already international
    This checks FORMAT only. It cannot know if the number is on WhatsApp.
    """
    s = (raw or "").strip()
    if not s:
        return None, "empty phone number"

    international = s.startswith("+") or s.startswith("00")
    digits = re.sub(r"\D", "", s)
    if s.startswith("00"):
        digits = digits[2:]
    if not digits:
        return None, "no digits in phone number"

    if not international:
        digits = digits.lstrip("0")
        if len(digits) == 10:
            digits = default_cc + digits

    if not (8 <= len(digits) <= 15):
        return None, f"invalid length ({len(digits)} digits)"

    if digits.startswith("91"):  # India-specific check
        national = digits[2:]
        if len(national) != 10 or national[0] not in "6789":
            return None, "invalid Indian mobile number"
    return digits, ""


# --------------------------------------------------------------------------
# Templates
# --------------------------------------------------------------------------

_PLACEHOLDER = re.compile(r"\{\{\s*(\w+)\s*\}\}")


@dataclass
class Template:
    name: str
    language: str
    category: str
    body: str
    var_count: int
    supported: bool
    reason: str = ""

    @property
    def label(self) -> str:
        flag = "" if self.supported else "  [not supported by this app]"
        return f"{self.name}  ({self.language}, {self.category}){flag}"


def parse_template(t: dict) -> Template:
    """Turn one entry of Meta's message_templates response into a Template."""
    body = ""
    supported = True
    reason = ""

    for comp in t.get("components", []):
        ctype = (comp.get("type") or "").upper()
        if ctype == "BODY":
            body = comp.get("text", "")
        elif ctype == "FOOTER":
            pass
        elif ctype == "HEADER":
            fmt = (comp.get("format") or "").upper()
            if fmt != "TEXT" or _PLACEHOLDER.search(comp.get("text", "")):
                supported, reason = False, f"header ({fmt or 'variable'}) needs extra parameters"
        elif ctype == "BUTTONS":
            for b in comp.get("buttons", []):
                btype = (b.get("type") or "").upper()
                if btype in ("QUICK_REPLY", "PHONE_NUMBER"):
                    continue
                if btype == "URL" and "{{" not in b.get("url", ""):
                    continue
                supported, reason = False, f"button type {btype} needs extra parameters"
        else:
            supported, reason = False, f"component {ctype} is not supported"

    names = _PLACEHOLDER.findall(body)
    var_count = 0
    if names:
        if all(n.isdigit() for n in names):
            var_count = max(int(n) for n in names)
        else:
            supported, reason = False, "named variables are not supported"

    return Template(
        name=t.get("name", ""),
        language=t.get("language", ""),
        category=t.get("category", ""),
        body=body,
        var_count=var_count,
        supported=supported,
        reason=reason,
    )


# --------------------------------------------------------------------------
# Contacts
# --------------------------------------------------------------------------


def clean_param(value: str) -> str:
    """WhatsApp rejects newlines/tabs and runs of 5+ spaces in body parameters."""
    v = re.sub(r"[\r\n\t]+", " ", value)
    v = re.sub(r" {2,}", " ", v)
    return v.strip()


@dataclass
class ValidationReport:
    valid: list[dict] = field(default_factory=list)  # {"phone": str, "params": [str]}
    invalid: list[tuple[int, str, str]] = field(default_factory=list)  # (row no., raw phone, reason)
    duplicates: int = 0
    opted_out: int = 0


def build_contacts(
    rows: list[list[str]],
    phone_idx: int,
    var_idx: list[int],
    default_cc: str = "91",
    opted_out: Optional[set[str]] = None,
) -> ValidationReport:
    """
    rows      - table rows (without the header)
    phone_idx - index of the phone column
    var_idx   - column index for {{1}}, {{2}}, ... in that order
    """
    opted_out = opted_out or set()
    rep = ValidationReport()
    seen: set[str] = set()

    for n, row in enumerate(rows, start=2):  # +2 => matches the row number in Excel
        raw_phone = row[phone_idx]
        phone, err = normalize_phone(raw_phone, default_cc)
        if err:
            rep.invalid.append((n, raw_phone, err))
            continue

        params = [clean_param(row[i]) for i in var_idx]
        empty = [k + 1 for k, p in enumerate(params) if p == ""]
        if empty:
            rep.invalid.append((n, raw_phone, "empty value for {{%s}}" % "}}, {{".join(map(str, empty))))
            continue

        if phone in opted_out:
            rep.opted_out += 1
            continue
        if phone in seen:
            rep.duplicates += 1
            continue
        seen.add(phone)
        rep.valid.append({"phone": phone, "params": params})
    return rep


# --------------------------------------------------------------------------
# n8n client
# --------------------------------------------------------------------------


class N8nError(Exception):
    pass


class N8nClient:
    def __init__(self, base_url: str, api_key: str, timeout: int = 60):
        self.base = base_url.rstrip("/")
        self.headers = {"x-api-key": api_key, "Content-Type": "application/json"}
        self.timeout = timeout

    def _post(self, path: str, payload: dict) -> dict:
        url = f"{self.base}/webhook/{path}"
        try:
            r = requests.post(url, json=payload, headers=self.headers, timeout=self.timeout)
        except requests.RequestException as e:
            raise N8nError(f"Cannot reach n8n: {e}") from e
        if r.status_code in (401, 403):
            raise N8nError("n8n rejected the API key (check the key in Settings).")
        if r.status_code == 404:
            raise N8nError(
                f"Webhook '{path}' not found. Is the workflow ACTIVE and is the URL the "
                "production one (/webhook/, not /webhook-test/)?"
            )
        if r.status_code >= 400:
            raise N8nError(f"n8n error {r.status_code}: {r.text[:300]}")
        if not r.content:
            return {}
        try:
            return r.json()
        except ValueError:
            return {}

    def list_templates(self) -> list[dict]:
        return self._post("wa-list-templates", {}).get("templates", [])

    def send_chunk(self, campaign_id: str, template: Template, contacts: list[dict]) -> None:
        self._post(
            "wa-send-campaign",
            {
                "campaignId": campaign_id,
                "templateName": template.name,
                "language": template.language,
                "contacts": contacts,
            },
        )

    def campaign_status(self, campaign_id: str) -> dict:
        d = self._post("wa-campaign-status", {"campaignId": campaign_id})
        for k in ("processed", "accepted", "send_failed", "delivered", "read", "delivery_failed"):
            d[k] = int(d.get(k, 0) or 0)
        d.setdefault("failures", [])
        return d

    def optouts(self) -> set[str]:
        return set(self._post("wa-optouts", {}).get("phones", []))

    def import_contacts_to_sheet(
        self, contacts: list[dict], sheet_tab: str = "Contacts"
    ) -> int:
        """
        Push a list of contact dicts to Google Sheets via the wa-import-contacts webhook.
        Each dict should have at least a 'phone' key plus any extra column keys.
        Returns the number of rows imported as confirmed by n8n.
        """
        resp = self._post(
            "wa-import-contacts",
            {"contacts": contacts, "sheetTab": sheet_tab},
        )
        return int(resp.get("imported", len(contacts)))


# --------------------------------------------------------------------------
# Campaign runner
# --------------------------------------------------------------------------

ProgressFn = Callable[[str, int, int], None]  # (message, done, total)


def make_campaign_id() -> str:
    return "c_" + dt.datetime.now().strftime("%Y%m%d_%H%M%S")


def run_campaign(
    client: N8nClient,
    campaign_id: str,
    template: Template,
    contacts: list[dict],
    cancel: threading.Event,
    on_progress: ProgressFn,
    chunk_size: int = 50,
    poll_seconds: float = 6.0,
) -> int:
    """
    Send contacts in chunks. After each chunk we wait until n8n has recorded
    a result row for every message of that chunk before sending the next one.
    Returns number of contacts handed to n8n.
    """
    total = len(contacts)
    submitted = 0
    for start in range(0, total, chunk_size):
        if cancel.is_set():
            on_progress("Stopped by user. Chunks already handed to n8n will still finish.", submitted, total)
            break
        chunk = contacts[start : start + chunk_size]
        client.send_chunk(campaign_id, template, chunk)
        submitted += len(chunk)
        on_progress(f"Chunk handed to n8n ({submitted}/{total}). Waiting for results...", submitted, total)

        deadline = time.time() + len(chunk) * 4 + 240
        while True:
            st = client.campaign_status(campaign_id)
            if st["processed"] >= submitted:
                on_progress(
                    f"Processed {st['processed']}/{total}  "
                    f"(accepted {st['accepted']}, failed at send {st['send_failed']})",
                    st["processed"],
                    total,
                )
                break
            if time.time() > deadline:
                raise N8nError(
                    "Timed out waiting for n8n. Check the 'Executions' tab in n8n for the "
                    "failed 'WA - Send Campaign' run. Do NOT resend until you have checked "
                    "the Messages sheet, or people will get duplicates."
                )
            if cancel.wait(poll_seconds):
                break
    return submitted
