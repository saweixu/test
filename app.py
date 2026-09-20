from __future__ import annotations

import argparse
import io
import json
import re
import zipfile
from dataclasses import asdict, dataclass
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Iterable


CONTAINER_RE = re.compile(r"\b([A-Z]{4}\s?\d{7})\b")
BL_TOKEN_RE = re.compile(r"\b(?:[A-Z]{2,5}\d{5,12}|\d{8,12})\b")
VERIFY_TYPE = "A verifier"


@dataclass
class InvoiceResult:
    source_file: str
    shipping_line: str
    container_number: str
    bl_number: str
    invoice_number: str
    document_type: str
    currency: str
    amount: str
    proposed_name: str
    confidence: str
    notes: str


def normalize_spaces(text: str) -> str:
    return re.sub(r"\s+", " ", text or "").strip()


def extract_pdf_text_from_bytes(data: bytes) -> str:
    try:
        import pdfplumber

        with pdfplumber.open(io.BytesIO(data)) as pdf:
            return "\n".join(page.extract_text() or "" for page in pdf.pages)
    except Exception:
        try:
            from pypdf import PdfReader

            reader = PdfReader(io.BytesIO(data))
            return "\n".join(page.extract_text() or "" for page in reader.pages)
        except Exception:
            return ""


def extract_pdf_text_from_path(path: Path) -> str:
    return extract_pdf_text_from_bytes(path.read_bytes())


def is_valid_container(number: str) -> bool:
    number = re.sub(r"\s+", "", number.upper())
    if not re.fullmatch(r"[A-Z]{4}\d{7}", number):
        return False

    values = {
        "A": 10,
        "B": 12,
        "C": 13,
        "D": 14,
        "E": 15,
        "F": 16,
        "G": 17,
        "H": 18,
        "I": 19,
        "J": 20,
        "K": 21,
        "L": 23,
        "M": 24,
        "N": 25,
        "O": 26,
        "P": 27,
        "Q": 28,
        "R": 29,
        "S": 30,
        "T": 31,
        "U": 32,
        "V": 34,
        "W": 35,
        "X": 36,
        "Y": 37,
        "Z": 38,
    }
    total = 0
    for idx, char in enumerate(number[:10]):
        value = int(char) if char.isdigit() else values.get(char, -1)
        if value < 0:
            return False
        total += value * (2**idx)

    expected = total % 11
    if expected == 10:
        expected = 0
    return expected == int(number[-1])


def dedupe_keep_order(items: Iterable[str]) -> list[str]:
    seen: set[str] = set()
    result: list[str] = []
    for item in items:
        if item and item not in seen:
            seen.add(item)
            result.append(item)
    return result


def extract_containers(text: str) -> list[str]:
    candidates = [
        re.sub(r"\s+", "", match.group(1).upper()) for match in CONTAINER_RE.finditer(text.upper())
    ]
    candidates = dedupe_keep_order(candidates)
    valid = [candidate for candidate in candidates if is_valid_container(candidate)]
    return valid or candidates


def extract_shipping_line(text: str) -> str:
    upper = text.upper()
    if (
        "SHARER" in upper
        or "SHARER-LOGISTICS" in upper
        or "ACCOUNT HOLDER: SHARER" in upper
        or re.search(r"\bINVOICE\s+NUMBER\s*:\s*SA\d+", upper)
    ):
        return "Sharer"
    if "CMA - CGM" in upper or "CMA CGM" in upper:
        return "CMA CGM"
    if "ORIENT OVERSEAS" in upper or "OOCL" in upper:
        return "OOCL"
    if "COSCO SHIPPING" in upper:
        return "COSCO"
    return ""


def clean_invoice_number(value: str) -> str:
    value = normalize_spaces(value)
    value = re.sub(r"[^A-Za-z0-9 -]", "", value)
    return value.replace(" ", "").strip("-")


def extract_invoice_number(text: str, file_name: str) -> str:
    token_with_digit = r"[A-Z]*\d[\dA-Z-]*"
    patterns = [
        rf"\bINVOICE\s+NUMBER\s*:?\s*({token_with_digit}(?:\s+{token_with_digit})*)",
        rf"\bINVOICE\s+NO\.?\s*:?\s*({token_with_digit}(?:\s+{token_with_digit})*)",
        r"\b(BEDIC\d{5,})\b",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if match:
            return clean_invoice_number(match.group(1))

    file_match = re.search(r"\b(BEDIC\d{5,}|INV\d{6,}[A-Z]*|\d{8,})\b", file_name, re.IGNORECASE)
    if file_match:
        return clean_invoice_number(file_match.group(1))
    return ""


def extract_bl_number(text: str) -> str:
    cma_match = re.search(r"\bBill\s+of\s+Lading\s*:\s*([A-Z0-9-]+)", text, re.IGNORECASE)
    if cma_match:
        return cma_match.group(1).strip()

    sharer_match = re.search(
        r"OCEAN\s+BILL\s+OF\s+LADING\s+HOUSE\s+BILL\s+OF\s+LADING.*?\b(\d{8,12})\b\s+[A-Z]{2,}",
        text,
        flags=re.IGNORECASE | re.DOTALL,
    )
    if sharer_match:
        return sharer_match.group(1).strip()

    general_segment = re.search(
        r"BILL\s+OF\s+LADING\s+NO\.?(.*?)(?:PLACE\s+OF\s+RECEIPT|SHIP\s+TO|VESSEL|REFERENCE|DESCRIPTION|$)",
        text,
        flags=re.IGNORECASE | re.DOTALL,
    )
    if general_segment:
        tokens = BL_TOKEN_RE.findall(general_segment.group(1).upper())
        tokens = [token for token in tokens if not token.startswith("BE")]
        numeric_tokens = [token for token in tokens if token.isdigit()]
        if numeric_tokens:
            return numeric_tokens[-1].strip()
        if tokens:
            return tokens[-1].strip()

    fallback = re.search(r"\bB/?L\s*(?:NO\.?|NUMBER)?\s*:?\s*([A-Z0-9-]+)", text, re.IGNORECASE)
    if fallback:
        return fallback.group(1).strip()
    return ""


def classify_document_type(text: str) -> str:
    upper = text.upper()
    types: list[str] = []

    if any(
        term in upper
        for term in [
            "STORAGE CHARGE",
            "TERMINAL FULL STORAGE",
            "FULL STORAGE",
            "STORAGE AT",
            "STORAGE FEE",
        ]
    ):
        types.append("Storage")

    is_detention = "DETENTION" in upper and (
        "DETENTION IMPORT CHARGE" in upper
        or "IMPORT DETENTION" in upper
        or ("GATE OUT FULL" in upper and "GATE IN EMPTY" in upper)
    )
    if is_detention:
        types.append("Detention")
    elif "DEMURRAGE/DETENTION" in upper or "DEM/DET" in upper:
        types.extend(["Demurrage", "Detention"])
    elif "DEMURRAGE" in upper:
        types.append("Demurrage")

    if any(term in upper for term in ["DEST TRML HANDLG", "TERMINAL HANDLING", "DTHC", "TRML HANDLG"]):
        types.append("THC")

    if not types and "SECURE RELEASE FEE" in upper:
        types.append("Release")

    preferred_order = ["Storage", "Demurrage", "Detention", "THC", "Release"]
    ordered = [item for item in preferred_order if item in set(types)]
    return "-".join(ordered) if ordered else VERIFY_TYPE


def parse_amount(value: str) -> str:
    value = value.strip().replace(" ", "")
    if "," in value and "." in value:
        if value.rfind(",") > value.rfind("."):
            value = value.replace(".", "").replace(",", ".")
        else:
            value = value.replace(",", "")
    elif "," in value:
        value = value.replace(".", "").replace(",", ".")

    try:
        return f"{Decimal(value).quantize(Decimal('0.01'))}"
    except InvalidOperation:
        return ""


def extract_amount(text: str) -> tuple[str, str]:
    patterns = [
        r"AMOUNT\s+DUE\s+([A-Z]{3})\s+([0-9][0-9.,]*)",
        r"AMOUNT\s+DUE\s*:?\s*([0-9][0-9.,]*)\s+([A-Z]{3})",
        r"TOTAL\s+AMOUNT\s+TO\s+BE\s+PAID\s+AMOUNT\s+([A-Z]{3})\s+([0-9][0-9.,]*)",
        r"TOTAL\s+AMOUNT\s+DUE\s*:?\s*([0-9][0-9.,]*)\s+([A-Z]{3})",
        r"TOTAL\s+AMOUNT\s*:?\s*([0-9][0-9.,]*)\s+([A-Z]{3})",
        r"TOTAL\s+([A-Z]{3})\s+([0-9][0-9.,]*)",
        r"TOTAL\s+INCLUDING\s+TAX\s+([0-9][0-9.,]*)",
        r"TOTAL\s+EXCLUDING\s+TAX\s+([0-9][0-9.,]*)",
    ]
    for pattern in patterns:
        match = re.search(pattern, text, flags=re.IGNORECASE)
        if not match:
            continue

        groups = match.groups()
        if len(groups) == 2 and re.fullmatch(r"[A-Z]{3}", groups[0], flags=re.IGNORECASE):
            return groups[0].upper(), parse_amount(groups[1])
        if len(groups) == 2:
            return groups[1].upper(), parse_amount(groups[0])
        return guess_currency(text), parse_amount(groups[0])

    return guess_currency(text), ""


def guess_currency(text: str) -> str:
    match = re.search(r"\b(EUR|USD|GBP)\b", text.upper())
    return match.group(1) if match else ""


def safe_filename_part(value: str) -> str:
    value = normalize_spaces(value).replace(" ", "-")
    value = re.sub(r"[^A-Za-z0-9._+-]+", "-", value)
    value = re.sub(r"-+", "-", value).strip(".-")
    return value or "Unknown"


def build_proposed_name(containers: list[str], document_type: str, file_name: str) -> str:
    if containers:
        container_part = "_".join(containers[:3])
        if len(containers) > 3:
            container_part += f"_plus{len(containers) - 3}"
    else:
        container_part = Path(file_name).stem

    return f"{safe_filename_part(container_part)}-{safe_filename_part(document_type)}.pdf"


def confidence_label(result: dict[str, str]) -> str:
    score = 0
    score += 25 if result["container_number"] else 0
    score += 20 if result["bl_number"] else 0
    score += 20 if result["invoice_number"] else 0
    score += 20 if result["amount"] else 0
    score += 15 if result["document_type"] and result["document_type"] != VERIFY_TYPE else 0

    if score >= 85 and result["document_type"] != VERIFY_TYPE:
        return "High"
    if score >= 60:
        return "Medium"
    return "Low"


def parse_invoice(text: str, file_name: str) -> InvoiceResult:
    if not normalize_spaces(text):
        return InvoiceResult(
            source_file=file_name,
            shipping_line="",
            container_number="",
            bl_number="",
            invoice_number=extract_invoice_number("", file_name),
            document_type=VERIFY_TYPE,
            currency="",
            amount="",
            proposed_name=f"{safe_filename_part(Path(file_name).stem)}-Unreadable.pdf",
            confidence="Low",
            notes="Aucun texte extrait; OCR probablement necessaire.",
        )

    containers = extract_containers(text)
    currency, amount = extract_amount(text)
    payload = {
        "source_file": file_name,
        "shipping_line": extract_shipping_line(text),
        "container_number": "; ".join(containers),
        "bl_number": extract_bl_number(text),
        "invoice_number": extract_invoice_number(text, file_name),
        "document_type": classify_document_type(text),
        "currency": currency,
        "amount": amount,
    }
    payload["proposed_name"] = build_proposed_name(containers, payload["document_type"], file_name)
    payload["confidence"] = confidence_label(payload)

    notes: list[str] = []
    if not payload["container_number"]:
        notes.append("Conteneur non trouve")
    if not payload["bl_number"]:
        notes.append("BL non trouve")
    if not payload["invoice_number"]:
        notes.append("Facture non trouvee")
    if not payload["amount"]:
        notes.append("Montant non trouve")
    if payload["document_type"] == VERIFY_TYPE:
        notes.append("Type a verifier")

    return InvoiceResult(**payload, notes="; ".join(notes))


def make_unique_filename(name: str, used: set[str]) -> str:
    name = safe_filename_part(Path(name).stem) + ".pdf"
    if name not in used:
        used.add(name)
        return name

    stem = Path(name).stem
    suffix = 2
    while True:
        candidate = f"{stem}-{suffix}.pdf"
        if candidate not in used:
            used.add(candidate)
            return candidate
        suffix += 1


def records_from_folder(folder: Path) -> list[dict[str, str]]:
    records: list[dict[str, str]] = []
    for path in sorted(folder.glob("*.pdf")):
        text = extract_pdf_text_from_path(path)
        records.append(asdict(parse_invoice(text, path.name)))
    return records


def build_zip(files: list[tuple[str, bytes]], rows: list[dict[str, str]]) -> bytes:
    used: set[str] = set()
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for idx, row in enumerate(rows):
            proposed = row.get("proposed_name") or row.get("nom_propose") or files[idx][0]
            final_name = make_unique_filename(str(proposed), used)
            archive.writestr(final_name, files[idx][1])
    buffer.seek(0)
    return buffer.getvalue()


def run_streamlit() -> None:
    import pandas as pd
    import streamlit as st

    st.set_page_config(page_title="Factures shipping", layout="wide")
    st.title("Factures shipping")

    uploads = st.file_uploader("PDF factures", type=["pdf"], accept_multiple_files=True)
    if not uploads:
        st.stop()

    files: list[tuple[str, bytes]] = []
    rows: list[dict[str, str]] = []

    for upload in uploads:
        data = upload.getvalue()
        files.append((upload.name, data))
        text = extract_pdf_text_from_bytes(data)
        rows.append(asdict(parse_invoice(text, upload.name)))

    df = pd.DataFrame(rows)
    column_order = [
        "source_file",
        "shipping_line",
        "container_number",
        "bl_number",
        "invoice_number",
        "document_type",
        "currency",
        "amount",
        "proposed_name",
        "confidence",
        "notes",
    ]
    df = df[column_order]

    edited_df = st.data_editor(
        df,
        use_container_width=True,
        hide_index=True,
        num_rows="fixed",
        column_config={
            "source_file": "Fichier original",
            "shipping_line": "Emetteur facture",
            "container_number": "Numero conteneur",
            "bl_number": "Numero BL",
            "invoice_number": "Numero facture",
            "document_type": "Type",
            "currency": "Devise",
            "amount": "Montant",
            "proposed_name": "Nom propose",
            "confidence": "Confiance",
            "notes": "Notes",
        },
    )

    export_df = edited_df.copy()
    csv_bytes = export_df.to_csv(index=False).encode("utf-8-sig")

    excel_buffer = io.BytesIO()
    with pd.ExcelWriter(excel_buffer, engine="openpyxl") as writer:
        export_df.to_excel(writer, index=False, sheet_name="Factures")
    excel_bytes = excel_buffer.getvalue()

    zip_bytes = build_zip(files, export_df.to_dict(orient="records"))

    col1, col2, col3 = st.columns(3)
    with col1:
        st.download_button("Telecharger Excel", excel_bytes, "factures_shipping.xlsx")
    with col2:
        st.download_button("Telecharger CSV", csv_bytes, "factures_shipping.csv")
    with col3:
        st.download_button("Telecharger PDF renommes", zip_bytes, "factures_renommees.zip")


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("--test-folder", type=Path, help="Parse all PDF files in a folder and print JSON.")
    args, _ = parser.parse_known_args()

    if args.test_folder:
        print(json.dumps(records_from_folder(args.test_folder), ensure_ascii=False, indent=2))
    else:
        run_streamlit()


if __name__ == "__main__":
    main()
