"""Read-only invoice selection and ZIP generation for saved bank payments."""

from collections.abc import Callable, Iterable
from dataclasses import dataclass, field
import io
import base64
import hashlib
import json
from pathlib import Path
import zipfile

from models.invoice import Invoice


def saved_payment_invoice_numbers(payment: dict, allocations: Iterable[dict]) -> list[int]:
    """Include every linked invoice, not just the payment's first invoice."""
    numbers: set[int] = set()
    for row in [payment, *allocations]:
        value = row.get("invoice_number")
        if value in (None, ""):
            continue
        for token in str(value).split(","):
            number = float(token.strip())
            if not number.is_integer() or number < 0:
                raise ValueError(f"Invalid saved invoice number: {token!r}")
            if number:
                numbers.add(int(number))
    return sorted(numbers)


@dataclass
class InvoiceZip:
    data: bytes = b""
    included: list[int] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)


def build_invoice_zip(
    invoice_numbers: Iterable[int],
    invoices: Iterable[Invoice],
    *,
    render_pdf: Callable[..., bytes],
) -> InvoiceZip:
    """Generate download-only PDFs without reserving numbers or writing records."""
    result = InvoiceZip()
    rows_by_number: dict[int, list[Invoice]] = {}
    for row in invoices:
        if row.invoice_number:
            rows_by_number.setdefault(int(row.invoice_number), []).append(row)

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for number in sorted(set(invoice_numbers)):
            rows = rows_by_number.get(number, [])
            if not rows:
                result.errors.append(f"Invoice #{number}: no saved invoice rows found.")
                continue
            try:
                pdf = render_pdf(invoice_rows=rows, invoice_number=number)
                if not pdf:
                    raise ValueError("PDF generator returned an empty document")
            except Exception as exc:
                result.errors.append(f"Invoice #{number}: PDF generation failed: {exc}")
                continue
            archive.writestr(f"CC_inv_{number}.pdf", pdf)
            result.included.append(number)
    if result.included:
        result.data = buffer.getvalue()
    return result


@dataclass
class ArchivedInvoiceZip:
    data: bytes = b""
    linked: list[int] = field(default_factory=list)
    originals: list[int] = field(default_factory=list)
    missing: list[int] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)
    unresolved: int = 0
    original_pdf_count: int = 0


def read_archived_invoice_file(path: str) -> bytes:
    """Read only within the application's existing sent-invoice archive."""
    from config.settings import get_data_paths

    root = (get_data_paths()["output_dir"] / "sent_invoices").resolve()
    target = Path(path).resolve()
    if not target.is_relative_to(root):
        raise ValueError("Saved archive path is outside the sent-invoice archive.")
    return target.read_bytes()


def build_saved_payment_originals_zip(
    payment: dict,
    allocations: list[dict],
    invoice_rows: list[dict],
    sent_rows: list[dict],
    *,
    download_pdf: Callable[[str, str], bytes],
    read_archive: Callable[[str], bytes] = read_archived_invoice_file,
) -> ArchivedInvoiceZip:
    """Use saved allocation identity and authoritative references; never generate a PDF."""
    result = ArchivedInvoiceZip()
    if payment.get("id") in (None, ""):
        raise ValueError("An immutable saved payment ID is required.")
    rows_by_id = {str(row["id"]): row for row in invoice_rows}
    linked: set[int] = set()
    linked_row_ids: dict[int, set[str]] = {}
    for allocation in allocations:
        if str(allocation.get("payment_id")) != str(payment["id"]):
            raise ValueError("Allocation does not belong to the selected payment.")
        row_id = allocation.get("invoice_row_id")
        numbers = saved_payment_invoice_numbers({}, [allocation])
        if row_id not in (None, ""):
            row = rows_by_id.get(str(row_id))
            resolved = saved_payment_invoice_numbers(row or {}, [])
            if not resolved:
                result.errors.append(f"Allocation {allocation.get('id')}: invoice row {row_id} is missing.")
                result.unresolved += 1
            elif numbers and numbers != resolved:
                raise ValueError(f"Allocation {allocation.get('id')}: invoice number conflicts with row {row_id}.")
            else:
                numbers = resolved
                linked_row_ids.setdefault(numbers[0], set()).add(str(row_id))
        elif not numbers:
            result.errors.append(f"Allocation {allocation.get('id')}: no saved invoice link.")
            result.unresolved += 1
        linked.update(numbers)
    if not allocations:
        # Legacy saved payments may have only an invoice-number link.
        linked.update(saved_payment_invoice_numbers(payment, []))
    result.linked = sorted(linked)

    report: dict = {
        "payment_id": payment["id"], "regenerated_fallbacks": 0,
        "unresolved_allocations": result.unresolved, "invoices": [],
        "errors": result.errors,
    }
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for number in result.linked:
            rows = [row for row in invoice_rows
                    if saved_payment_invoice_numbers(row, []) == [number]]
            sources = [
                ("invoices", row) for row in rows
            ] + [
                (row.get("source_table", "sent_invoices"), row) for row in sent_rows
                if saved_payment_invoice_numbers(row, []) == [number]
            ]
            documents: list[dict] = []
            references: set[tuple[str, ...]] = set()
            digests: set[str] = set()
            for table, source in sources:
                bucket = str(source.get("pdf_storage_bucket") or "").strip()
                path = str(source.get("pdf_storage_path") or "").strip()
                local = str(source.get("pdf_archive_path") or "").strip()
                candidates = []
                if bucket and path:
                    candidates.append(("storage", bucket, path))
                elif bucket or path:
                    result.errors.append(f"Invoice #{number}: incomplete {table} storage reference.")
                if local:
                    candidates.append(("archive", local))
                encoded = source.get("archived_pdf_base64")
                if encoded:
                    candidates.append(("base64", str(encoded)))
                for reference in candidates:
                    if reference in references:
                        continue
                    references.add(reference)
                    try:
                        if reference[0] == "storage":
                            pdf = download_pdf(reference[1], reference[2])
                        elif reference[0] == "base64":
                            pdf = base64.b64decode(reference[1], validate=True)
                        else:
                            pdf = read_archive(reference[1])
                        if not isinstance(pdf, bytes) or not pdf.startswith(b"%PDF-"):
                            raise ValueError("Archived object is not a PDF.")
                    except Exception as exc:
                        result.errors.append(
                            f"Invoice #{number}: {table} record {source.get('id')} "
                            f"archive retrieval failed: {exc}"
                        )
                        continue
                    digest = hashlib.sha256(pdf).hexdigest()
                    if digest in digests:
                        continue
                    digests.add(digest)
                    name = f"originals/CC_inv_{number}_{len(documents) + 1}.pdf"
                    archive.writestr(name, pdf)
                    documents.append({
                        "file": name, "sha256": digest, "source_table": table,
                        "source_id": source.get("id"), "type": "archived_original",
                    })
            if documents:
                result.originals.append(number)
                result.original_pdf_count += len(documents)
            else:
                result.missing.append(number)
            report["invoices"].append({
                "invoice_number": number,
                "allocation_invoice_row_ids": sorted(linked_row_ids.get(number, set())),
                "related_invoice_row_ids": [row["id"] for row in rows],
                "status": "original_found" if documents else "missing_original_pdf",
                "documents": documents,
            })
        archive.writestr("invoice_download_report.json", json.dumps(report, indent=2))
    if result.originals:
        result.data = buffer.getvalue()
    return result
