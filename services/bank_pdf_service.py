"""Parse SWIFT MT103 bank transfer PDFs to extract payment info."""
import re
import datetime
import io
from typing import Optional


def _extract_text(file_bytes: bytes) -> str:
    try:
        import pdfplumber
        with pdfplumber.open(io.BytesIO(file_bytes)) as pdf:
            pages = [page.extract_text() or "" for page in pdf.pages]
        return "\n".join(pages)
    except Exception as e:
        raise RuntimeError(f"Could not read PDF: {e}")


def _parse_swift_amount(s: str) -> Optional[float]:
    """Parse SWIFT amount: '10716,' → 10716.0, '10704.28' → 10704.28, '10716,28' → 10716.28"""
    s = s.strip().rstrip(",")
    if not s:
        return None
    try:
        # European format: period as thousands separator, comma as decimal
        if "," in s and "." in s:
            # e.g. "10.716,28" → remove dots, replace comma with dot
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:
            # e.g. "10716,28" → replace comma with dot
            s = s.replace(",", ".")
        return float(s)
    except ValueError:
        return None


def parse_swift_pdf(file_bytes: bytes) -> dict:
    """
    Parse a SWIFT MT103 bank transfer PDF.

    Returns dict with:
        invoice_number   : int | None        — first detected invoice number (back-compat)
        invoice_numbers  : list[int]          — every invoice number found in the remittance info
        payment_date     : datetime.date | None
        instructed_amount: float | None  — :33B field (what customer sent)
        received_amount  : float | None  — :32A field (what arrived after fees)
        raw_text         : str
    """
    text = _extract_text(file_bytes)
    result: dict = {
        "invoice_number": None,
        "invoice_numbers": [],
        "payment_date": None,
        "instructed_amount": None,
        "received_amount": None,
        "raw_text": text,
    }

    # ── Invoice number(s) ───────────────────────────────────────────────────────
    # The remittance info can reference a single invoice ("...INFORMATION :8665")
    # or several, separated by slashes/commas when one transfer settles multiple
    # invoices at once ("...INFORMATION :8688/8689/8690/8691/8692/8693/8694").
    # Capture the whole remittance value (up to end of line), then split it.
    m = re.search(
        r':70[:\s]+REMITTANCE\s+INFORMATION\s*:?\s*([\d/,\s]+)',
        text, re.IGNORECASE,
    )
    if m:
        numbers = [int(tok) for tok in re.findall(r'\d+', m.group(1))]
        if numbers:
            result["invoice_numbers"] = numbers
            result["invoice_number"] = numbers[0]

    # ── Payment date ───────────────────────────────────────────────────────────
    # Matches: "VALUE FOR CUSTOMER: 22/01/26"  (DD/MM/YY)
    m = re.search(
        r'VALUE\s+FOR\s+CUSTOMER\s*:\s*(\d{2}/\d{2}/\d{2})',
        text, re.IGNORECASE,
    )
    if m:
        try:
            d, mo, y = m.group(1).split("/")
            result["payment_date"] = datetime.date(2000 + int(y), int(mo), int(d))
        except (ValueError, OverflowError):
            pass

    # ── Instructed amount :33B ─────────────────────────────────────────────────
    # Matches: ":33B:CURR.INSTRUCTED AMT. :EUR10716,"
    m = re.search(
        r':33B[:\s]+CURR\.INSTRUCTED\s+AMT\.\s*:EUR([\d.,]+)',
        text, re.IGNORECASE,
    )
    if m:
        result["instructed_amount"] = _parse_swift_amount(m.group(1))

    # ── Received amount :32A ───────────────────────────────────────────────────
    # Matches ":32A:AMNT.(COL/ACP/ACK) 260122EUR10704.28" as well as the
    # colon-separated variant ":32A:AMNT.(COL/ACP/ACK) :260901EUR38548,4"
    m = re.search(
        r':32A[:\s]+AMNT\.\s*\(COL/ACP/ACK\)\s*:?\s*\d{6}EUR([\d.,]+)',
        text, re.IGNORECASE,
    )
    if m:
        result["received_amount"] = _parse_swift_amount(m.group(1))

    return result
