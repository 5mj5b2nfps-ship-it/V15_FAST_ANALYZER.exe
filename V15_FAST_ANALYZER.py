#!/usr/bin/env python3
"""
V15_FAST_ANALYZER - read-only, attachment-list-only PO approval cockpit.

Reads (from --base-dir):
  Pending_PO_Details.txt
  Pending_PO_Attachments.txt       (optional; names only, files are never opened)
  Fast_Run_Status.txt              (optional master-stage status)
  EXPORT_*.xlsx                    (latest historical SAP export)

Writes:
  PO_Approval_Cockpit_FAST.xlsx
  V15_FAST_Analysis.json
  V15_FAST_Run_Log.txt

The program does not connect to SAP and does not open or parse attachment files.
Required dependency: openpyxl
"""
from __future__ import annotations

import argparse
import csv
import json
import logging
import math
import re
import statistics
import sys
import traceback
import unicodedata
from collections import Counter, defaultdict
from dataclasses import asdict, dataclass, field
from datetime import date, datetime
from difflib import SequenceMatcher
from pathlib import Path
from typing import Any, Iterable, Optional

try:
    from openpyxl import Workbook, load_workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.table import Table, TableStyleInfo
except ImportError as exc:
    raise SystemExit("Missing required package: openpyxl") from exc

VERSION = "15-FAST-1.0"
DEFAULT_CURRENCY = "OMR"
DECISION_RANK = {"APPROVE": 0, "CHECK": 1, "REVIEW": 2, "HOLD": 3, "DATA ERROR": 4}


@dataclass
class Config:
    history_ok_pct: float = 3.0
    history_review_pct: float = 5.0
    history_justification_pct: float = 10.0
    family_check_pct: float = 10.0
    family_review_pct: float = 25.0
    family_justification_pct: float = 50.0
    fuzzy_history_min: float = 0.74
    family_similarity_min: float = 0.68
    max_history_matches: int = 20


@dataclass
class POItem:
    po_number: str
    po_date: Optional[datetime]
    supplier: str
    item: str
    description: str
    material: str
    material_group: str
    plant: str
    quantity: Optional[float]
    uom: str
    net_price: Optional[float]
    currency: str
    price_unit: float
    price_uom: str

    @property
    def unit_price(self) -> Optional[float]:
        if self.net_price is None:
            return None
        return self.net_price / (self.price_unit or 1.0)

    @property
    def line_value(self) -> Optional[float]:
        if self.net_price is None or self.quantity is None:
            return None
        return self.net_price * self.quantity / (self.price_unit or 1.0)


@dataclass
class AttachmentSummary:
    po_number: str
    scan_status: str = "UNKNOWN"
    count: int = 0
    cs_named: bool = False
    quote_named: bool = False
    rfq_named: bool = False
    scope_named: bool = False
    contract_named: bool = False
    technical_named: bool = False
    names: list[str] = field(default_factory=list)


@dataclass
class HistoryMatch:
    po_number: str
    item: str
    source_po: str
    source_item: str
    source_date: str
    source_supplier: str
    source_description: str
    source_material: str
    source_material_group: str
    source_plant: str
    source_unit_price: Optional[float]
    source_currency: str
    source_uom: str
    tier: str
    score: float
    supplier_match: bool
    plant_match: bool


@dataclass
class AnalysisRow:
    po_number: str
    po_date: str
    supplier: str
    item: str
    description: str
    material: str
    material_group: str
    plant: str
    quantity: Optional[float]
    uom: str
    net_price: Optional[float]
    price_unit: float
    unit_price: Optional[float]
    currency: str
    line_value: Optional[float]
    attachment_scan: str
    attachment_count: int
    attachment_types: str
    cs_named: str
    quote_named: str
    history_tier: str = "NONE"
    history_count: int = 0
    previous_price: Optional[float] = None
    previous_date: str = ""
    median_last3: Optional[float] = None
    vs_previous_pct: Optional[float] = None
    vs_median_pct: Optional[float] = None
    history_match_score: Optional[float] = None
    family_count: int = 0
    family_median: Optional[float] = None
    vs_family_pct: Optional[float] = None
    description_assessment: str = "NO COMPARABLE FAMILY"
    benchmark_strength: str = "NONE"
    justification_required: str = "NO"
    decision: str = "REVIEW"
    reasons: list[str] = field(default_factory=list)


# ------------------------------- helpers ---------------------------------

def text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def normalize(value: Any) -> str:
    s = unicodedata.normalize("NFKD", text(value)).encode("ascii", "ignore").decode("ascii")
    s = s.lower().replace("&", " and ")
    s = re.sub(r"[^a-z0-9]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def compact(value: Any) -> str:
    return re.sub(r"[^a-z0-9]", "", normalize(value))


def token_set(value: Any) -> set[str]:
    stop = {"the", "and", "for", "with", "from", "unit", "item", "nr", "nos", "no", "each"}
    return {t for t in normalize(value).split() if len(t) > 1 and t not in stop}


def identity_tokens(value: Any) -> set[str]:
    result = set()
    for raw in re.findall(r"[A-Za-z0-9][A-Za-z0-9._/\-]{3,}", text(value)):
        t = compact(raw)
        if len(t) >= 5 and re.search(r"[a-z]", t) and re.search(r"\d", t):
            result.add(t)
    return result


def supplier_id(value: Any) -> str:
    m = re.search(r"(?<!\d)(\d{10})(?!\d)", text(value))
    return m.group(1) if m else ""


def normalize_supplier(value: Any) -> str:
    s = normalize(value)
    s = re.sub(r"\b\d{10}\b", " ", s)
    stop = {"llc", "ltd", "limited", "spc", "saoc", "saog", "co", "company", "inc", "gmbh", "trading", "services", "service", "enterprise", "enterprises"}
    return " ".join(t for t in s.split() if t not in stop)


def supplier_equal(a: Any, b: Any) -> bool:
    ia, ib = supplier_id(a), supplier_id(b)
    if ia and ib:
        return ia == ib
    na, nb = normalize_supplier(a), normalize_supplier(b)
    if not na or not nb:
        return False
    return na == nb or SequenceMatcher(None, na, nb).ratio() >= 0.90


def similarity(a: Any, b: Any) -> float:
    na, nb = normalize(a), normalize(b)
    if not na or not nb:
        return 0.0
    seq = SequenceMatcher(None, na, nb).ratio()
    ta, tb = token_set(na), token_set(nb)
    jaccard = len(ta & tb) / len(ta | tb) if ta and tb else 0.0
    containment = len(ta & tb) / min(len(ta), len(tb)) if ta and tb else 0.0
    return max(seq, 0.55 * jaccard + 0.45 * containment)


def to_float(value: Any) -> Optional[float]:
    if value is None or value == "":
        return None
    if isinstance(value, (int, float)):
        v = float(value)
        return v if math.isfinite(v) else None
    s = text(value).replace("\u00a0", " ").strip()
    s = re.sub(r"[^0-9,().+\-]", "", s)
    if not s:
        return None
    negative = s.startswith("(") and s.endswith(")")
    s = s.strip("()")
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        s = "".join(parts) if len(parts[-1]) == 3 else ".".join(parts)
    try:
        v = float(s)
        return -v if negative else v
    except ValueError:
        return None


def parse_date(value: Any) -> Optional[datetime]:
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return value
    if isinstance(value, date):
        return datetime.combine(value, datetime.min.time())
    s = text(value)
    for fmt in ("%Y-%m-%d", "%d.%m.%Y", "%d/%m/%Y", "%m/%d/%Y", "%Y%m%d", "%d-%m-%Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    return None


def pct_variance(current: Optional[float], reference: Optional[float]) -> Optional[float]:
    if current is None or reference is None or reference == 0:
        return None
    return (current - reference) / reference * 100.0


def family_key(description: str) -> str:
    size_words = {"small", "medium", "large", "xl", "xxl", "xxxl", "xs", "s", "m", "l", "size", "gray", "grey", "black", "white", "blue", "red", "green"}
    tokens = [t for t in normalize(description).split() if t not in size_words and not re.fullmatch(r"\d+(?:\.\d+)?", t)]
    return " ".join(tokens)


def service_item(item: POItem) -> bool:
    uoms = {"AU", "ACT", "JOB", "LOT", "LS", "SRV", "SERVICE", "MON", "MONTH"}
    words = {"service", "repair", "maintenance", "installation", "commissioning", "consultancy", "calibration", "rental", "fixing", "transportation"}
    return normalize(item.uom).upper() in uoms or bool(token_set(item.description) & words)


def read_delimited(path: Path) -> list[dict[str, str]]:
    if not path.exists():
        return []
    raw = path.read_bytes()
    decoded = None
    for enc in ("utf-8-sig", "utf-16", "cp1252", "latin-1"):
        try:
            decoded = raw.decode(enc)
            break
        except UnicodeDecodeError:
            continue
    if decoded is None:
        raise ValueError(f"Cannot decode {path}")
    lines = [line for line in decoded.splitlines() if line.strip()]
    if not lines:
        return []
    delimiter = "\t" if "\t" in lines[0] else ","
    return [dict(row) for row in csv.DictReader(lines, delimiter=delimiter)]


def get_field(row: dict[str, Any], *aliases: str) -> Any:
    normalized = {normalize(k): v for k, v in row.items()}
    for alias in aliases:
        key = normalize(alias)
        if key in normalized:
            return normalized[key]
    return ""


# ------------------------------- inputs ----------------------------------

def load_po_details(path: Path) -> list[POItem]:
    result: list[POItem] = []
    for row in read_delimited(path):
        po = text(get_field(row, "PO Number", "PO", "Purchasing Document", "EBELN"))
        if not re.fullmatch(r"\d{10}", po):
            continue
        result.append(POItem(
            po_number=po,
            po_date=parse_date(get_field(row, "PO Date", "Document Date", "BEDAT")),
            supplier=text(get_field(row, "Supplier", "Supplier/Supplying Plant", "Vendor", "Vendor Name", "LIFNR")),
            item=text(get_field(row, "Item", "PO Item", "EBELP")),
            description=text(get_field(row, "Description", "Short Text", "Material Description", "TXZ01")),
            material=text(get_field(row, "Material", "Material Code", "MATNR")),
            material_group=text(get_field(row, "Material Group", "MATKL")),
            plant=text(get_field(row, "Plant", "WERKS")),
            quantity=to_float(get_field(row, "Quantity", "Order Quantity", "MENGE")),
            uom=text(get_field(row, "UoM", "Order Unit", "MEINS")),
            net_price=to_float(get_field(row, "Net Price", "Price", "NETPR")),
            currency=text(get_field(row, "Currency", "WAERS")) or DEFAULT_CURRENCY,
            price_unit=to_float(get_field(row, "Price Unit", "PEINH")) or 1.0,
            price_uom=text(get_field(row, "Price UoM", "Price Unit UoM", "BPRME")),
        ))
    return result


def classify_attachment(name: str) -> str:
    n = normalize(name)
    padded = f" {n} "
    if any(x in n for x in ("contract", "agreement", "framework")):
        return "CONTRACT"
    if any(x in n for x in ("comparison", "comparative statement", "commercial evaluation")) or " cs " in padded or n.startswith("cs "):
        return "CS"
    if any(x in n for x in ("quotation", "quote", "offer")):
        return "QUOTE"
    if " rfq " in padded or n.startswith("rfq "):
        return "RFQ"
    if any(x in n for x in ("scope", "boq", "specification", "drawing")):
        return "SCOPE"
    if any(x in n for x in ("technical", "oem", "single source", "sole source")):
        return "TECHNICAL"
    if n.startswith(("re ", "fw ", "fwd ")) or Path(name).suffix.lower() in {".msg", ".eml"}:
        return "EMAIL"
    return "OTHER"


def load_run_status(path: Path) -> dict[str, dict[str, str]]:
    result: dict[str, dict[str, str]] = {}
    if not path.exists():
        return result
    for row in read_delimited(path):
        po = text(get_field(row, "PO Number", "PO"))
        if po:
            result[po] = row
    return result


def load_attachments(path: Path, po_numbers: Iterable[str], run_status: dict[str, dict[str, str]]) -> tuple[dict[str, AttachmentSummary], list[dict[str, str]]]:
    summaries = {po: AttachmentSummary(po_number=po) for po in po_numbers}
    inventory_rows: list[dict[str, str]] = []
    if path.exists():
        for row in read_delimited(path):
            po = text(get_field(row, "PO Number", "PO"))
            if po not in summaries:
                continue
            desc = text(get_field(row, "Description", "Attachment", "File Name", "Title"))
            if not desc:
                continue
            category = classify_attachment(desc)
            inventory_rows.append({"po_number": po, "row": text(get_field(row, "Attachment Row", "Row")), "description": desc, "category": category})
            s = summaries[po]
            s.names.append(desc)
            if category == "CS": s.cs_named = True
            elif category == "QUOTE": s.quote_named = True
            elif category == "RFQ": s.rfq_named = True
            elif category == "SCOPE": s.scope_named = True
            elif category == "CONTRACT": s.contract_named = True
            elif category == "TECHNICAL": s.technical_named = True
    for po, s in summaries.items():
        s.count = len(s.names)
        status_row = run_status.get(po, {})
        scan = text(get_field(status_row, "Attachment Scan Status", "Attachment Status", "Scan Status")).upper()
        if scan:
            s.scan_status = scan
        elif s.count > 0:
            s.scan_status = "PRESENT"
        else:
            s.scan_status = "UNKNOWN"
        if s.scan_status == "PRESENT" and s.count == 0:
            s.scan_status = "UNKNOWN"
        if s.count > 0 and s.scan_status in {"UNKNOWN", "NONE"}:
            s.scan_status = "PRESENT"
    return summaries, inventory_rows


HISTORY_ALIASES = {
    "po_number": ["Purchasing Document", "PO Number", "PO", "EBELN"],
    "item": ["Item", "PO Item", "EBELP"],
    "date": ["Document Date", "PO Date", "BEDAT"],
    "supplier": ["Supplier/Supplying Plant", "Supplier", "Vendor", "Vendor Name", "LIFNR"],
    "description": ["Short Text", "Description", "Material Description", "TXZ01"],
    "material": ["Material", "Material Code", "MATNR"],
    "material_group": ["Material Group", "MATKL"],
    "plant": ["Plant", "WERKS"],
    "price": ["Net Price", "Price", "Order Price", "NETPR"],
    "currency": ["Currency", "Document Currency", "WAERS"],
    "price_unit": ["Price Unit", "PEINH"],
    "uom": ["Order Unit", "UoM", "Price UoM", "MEINS", "BPRME"],
}


def detect_header_row(ws, max_rows: int = 30) -> tuple[int, dict[str, int]]:
    for r in range(1, min(ws.max_row, max_rows) + 1):
        values = [text(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)]
        normalized = {normalize(v): i + 1 for i, v in enumerate(values) if v}
        mapping: dict[str, int] = {}
        for field_name, aliases in HISTORY_ALIASES.items():
            for alias in aliases:
                if normalize(alias) in normalized:
                    mapping[field_name] = normalized[normalize(alias)]
                    break
        if "description" in mapping and "price" in mapping:
            return r, mapping
    return 0, {}


def load_history(path: Optional[Path], sheet_name: str = "") -> list[dict[str, Any]]:
    if path is None or not path.exists():
        return []
    wb = load_workbook(path, read_only=True, data_only=True)
    sheets = [wb[sheet_name]] if sheet_name and sheet_name in wb.sheetnames else wb.worksheets
    result: list[dict[str, Any]] = []
    for ws in sheets:
        header_row, columns = detect_header_row(ws)
        if not header_row:
            continue
        # Streaming tuples are dramatically faster than repeated ws.cell calls
        # on large SAP exports opened in read_only mode.
        for values in ws.iter_rows(min_row=header_row + 1, values_only=True):
            def value(name: str) -> Any:
                col = columns.get(name)
                return values[col - 1] if col and col <= len(values) else None
            description = text(value("description"))
            price = to_float(value("price"))
            if not description or price is None:
                continue
            price_unit = to_float(value("price_unit")) or 1.0
            result.append({
                "po_number": text(value("po_number")), "item": text(value("item")),
                "date": parse_date(value("date")), "supplier": text(value("supplier")),
                "description": description, "material": text(value("material")),
                "material_group": text(value("material_group")), "plant": text(value("plant")),
                "unit_price": price / price_unit, "currency": text(value("currency")) or DEFAULT_CURRENCY,
                "uom": text(value("uom")),
            })
    return result


# -------------------------- matching and analysis -------------------------

def history_candidate(item: POItem, row: dict[str, Any], config: Config) -> Optional[tuple[int, float, str]]:
    if row["po_number"] == item.po_number:
        return None
    if item.po_date and row["date"] and row["date"] > item.po_date:
        return None
    if normalize(row["currency"]) != normalize(item.currency):
        return None
    if item.uom and row["uom"] and normalize(item.uom) != normalize(row["uom"]):
        return None
    material_match = bool(item.material and row["material"] and compact(item.material) == compact(row["material"]))
    ids_now, ids_hist = identity_tokens(item.description), identity_tokens(row["description"])
    id_match = bool(ids_now and ids_hist and ids_now & ids_hist)
    if ids_now and ids_hist and not id_match and not material_match:
        return None
    supplier_match = supplier_equal(item.supplier, row["supplier"])
    plant_match = bool(item.plant and row["plant"] and normalize(item.plant) == normalize(row["plant"]))
    group_match = bool(item.material_group and row["material_group"] and compact(item.material_group) == compact(row["material_group"]))
    desc_score = similarity(item.description, row["description"])
    if material_match and supplier_match and plant_match:
        return 1, min(1.0, 0.97 + 0.03 * desc_score), "MATERIAL+SUPPLIER+PLANT"
    if material_match and supplier_match:
        return 2, min(1.0, 0.94 + 0.04 * desc_score), "MATERIAL+SUPPLIER"
    if material_match:
        return 3, min(1.0, 0.90 + 0.08 * desc_score), "MATERIAL"
    if id_match and group_match:
        score = min(1.0, 0.84 + 0.12 * desc_score + (0.03 if supplier_match else 0))
        return 4, score, "MODEL+MATERIAL GROUP"
    if group_match and desc_score >= config.fuzzy_history_min:
        score = min(1.0, desc_score + (0.05 if supplier_match else 0) + (0.02 if plant_match else 0))
        return 5, score, "DESCRIPTION+MATERIAL GROUP"
    if desc_score >= max(0.86, config.fuzzy_history_min + 0.08):
        score = min(1.0, desc_score + (0.04 if supplier_match else 0))
        return 6, score, "DESCRIPTION"
    return None


def select_history(item: POItem, history: list[dict[str, Any]], config: Config) -> tuple[list[HistoryMatch], dict[str, Any]]:
    candidates: list[tuple[int, float, dict[str, Any], str]] = []
    for row in history:
        candidate = history_candidate(item, row, config)
        if candidate:
            rank, score, tier = candidate
            candidates.append((rank, score, row, tier))
    if not candidates:
        return [], {"count": 0}
    best_rank = min(x[0] for x in candidates)
    selected = [x for x in candidates if x[0] == best_rank]
    selected.sort(key=lambda x: ((x[2]["date"] or datetime.min), x[1]), reverse=True)
    selected = selected[:config.max_history_matches]
    matches: list[HistoryMatch] = []
    for _, score, row, tier in selected:
        matches.append(HistoryMatch(
            po_number=item.po_number, item=item.item, source_po=row["po_number"], source_item=row["item"],
            source_date=row["date"].strftime("%Y-%m-%d") if row["date"] else "",
            source_supplier=row["supplier"], source_description=row["description"], source_material=row["material"],
            source_material_group=row["material_group"], source_plant=row["plant"], source_unit_price=row["unit_price"],
            source_currency=row["currency"], source_uom=row["uom"], tier=tier, score=score,
            supplier_match=supplier_equal(item.supplier, row["supplier"]),
            plant_match=bool(item.plant and row["plant"] and normalize(item.plant) == normalize(row["plant"])),
        ))
    prices = [m.source_unit_price for m in matches if m.source_unit_price is not None]
    last3 = prices[:3]
    latest = matches[0]
    stats = {
        "count": len(prices), "tier": latest.tier, "previous_price": latest.source_unit_price,
        "previous_date": latest.source_date, "median_last3": statistics.median(last3) if last3 else None,
        "score": latest.score,
    }
    return matches, stats


def family_comparables(item: POItem, items: list[POItem], history: list[dict[str, Any]], config: Config) -> list[float]:
    values: list[float] = []
    base_key = family_key(item.description)
    base_ids = identity_tokens(item.description)
    for other in items:
        if other.po_number != item.po_number or other.item == item.item or other.unit_price is None:
            continue
        if normalize(other.currency) != normalize(item.currency):
            continue
        if item.uom and other.uom and normalize(item.uom) != normalize(other.uom):
            continue
        other_ids = identity_tokens(other.description)
        if base_ids and other_ids and not (base_ids & other_ids):
            continue
        group_ok = not item.material_group or not other.material_group or compact(item.material_group) == compact(other.material_group)
        if group_ok and similarity(base_key, family_key(other.description)) >= config.family_similarity_min:
            values.append(other.unit_price)
    if len(values) < 2:
        for row in history:
            if row["unit_price"] is None or normalize(row["currency"]) != normalize(item.currency):
                continue
            if item.uom and row["uom"] and normalize(item.uom) != normalize(row["uom"]):
                continue
            row_ids = identity_tokens(row["description"])
            if base_ids and row_ids and not (base_ids & row_ids):
                continue
            group_ok = not item.material_group or not row["material_group"] or compact(item.material_group) == compact(row["material_group"])
            if group_ok and similarity(base_key, family_key(row["description"])) >= max(config.family_similarity_min, 0.74):
                values.append(row["unit_price"])
            if len(values) >= 20:
                break
    return values


def attachment_type_text(a: AttachmentSummary) -> str:
    types = []
    if a.cs_named: types.append("CS")
    if a.quote_named: types.append("QUOTE")
    if a.rfq_named: types.append("RFQ")
    if a.scope_named: types.append("SCOPE")
    if a.contract_named: types.append("CONTRACT")
    if a.technical_named: types.append("TECHNICAL")
    return ", ".join(types) if types else ("OTHER" if a.count else "NONE")


def analyze_item(item: POItem, all_items: list[POItem], attachment: AttachmentSummary, history: list[dict[str, Any]], config: Config, run_status: dict[str, dict[str, str]]) -> tuple[AnalysisRow, list[HistoryMatch]]:
    matches, hs = select_history(item, history, config)
    family = family_comparables(item, all_items, history, config)
    family_median = statistics.median(family) if family else None
    row = AnalysisRow(
        po_number=item.po_number, po_date=item.po_date.strftime("%Y-%m-%d") if item.po_date else "",
        supplier=item.supplier, item=item.item, description=item.description, material=item.material,
        material_group=item.material_group, plant=item.plant, quantity=item.quantity, uom=item.uom,
        net_price=item.net_price, price_unit=item.price_unit, unit_price=item.unit_price,
        currency=item.currency, line_value=item.line_value, attachment_scan=attachment.scan_status,
        attachment_count=attachment.count, attachment_types=attachment_type_text(attachment),
        cs_named="YES" if attachment.cs_named else "NO", quote_named="YES" if attachment.quote_named else "NO",
        history_tier=hs.get("tier", "NONE"), history_count=hs.get("count", 0),
        previous_price=hs.get("previous_price"), previous_date=hs.get("previous_date", ""),
        median_last3=hs.get("median_last3"), history_match_score=hs.get("score"),
        family_count=len(family), family_median=family_median,
    )
    row.vs_previous_pct = pct_variance(item.unit_price, row.previous_price)
    row.vs_median_pct = pct_variance(item.unit_price, row.median_last3)
    row.vs_family_pct = pct_variance(item.unit_price, family_median)
    if row.family_count >= 2 and row.vs_family_pct is not None:
        v = abs(row.vs_family_pct)
        if v <= config.family_check_pct:
            row.description_assessment = "CONSISTENT"
        elif v <= config.family_review_pct:
            row.description_assessment = "CHECK"
        else:
            row.description_assessment = "ANOMALY"
    elif service_item(item):
        row.description_assessment = "SERVICE - NOT PRICEABLE FROM TEXT"

    status = run_status.get(item.po_number, {})
    read_state = text(get_field(status, "Read Status", "PO Read Status")).upper()
    attach_state = text(get_field(status, "Attachment Scan Status", "Attachment Status")).upper()
    reasons: list[str] = []
    severity = 0
    justify = False
    if read_state and read_state != "OK":
        row.decision = "DATA ERROR"
        row.reasons = [f"PO reader status: {read_state}"]
        return row, matches
    if attach_state == "SCAN_FAILED" or attachment.scan_status == "SCAN_FAILED":
        reasons.append("Attachment scan failed; attachment evidence is unknown")
        severity = max(severity, 2)
    elif attachment.scan_status == "NONE":
        reasons.append("No attachment listed in SAP")
        severity = max(severity, 2 if row.history_count else 3)
    elif attachment.scan_status == "UNKNOWN":
        reasons.append("Attachment status is unknown")
        severity = max(severity, 2)

    if row.vs_median_pct is not None:
        v = row.vs_median_pct
        if v > config.history_justification_pct:
            reasons.append(f"Current price is {v:+.1f}% vs last-3 historical median; justification required")
            severity = max(severity, 2); justify = True
        elif v > config.history_review_pct:
            reasons.append(f"Current price is {v:+.1f}% vs last-3 historical median")
            severity = max(severity, 2)
        elif v > config.history_ok_pct:
            reasons.append(f"Current price is {v:+.1f}% vs last-3 historical median")
            severity = max(severity, 1)
        elif v < -config.history_review_pct:
            reasons.append(f"Current price is {v:+.1f}% below last-3 historical median")
        else:
            reasons.append(f"Historical variance {v:+.1f}% is within tolerance")
    else:
        reasons.append("No reliable historical benchmark")
        severity = max(severity, 2 if attachment.count > 0 else 3)

    if row.vs_family_pct is not None and row.family_count >= 2:
        v = abs(row.vs_family_pct)
        if v > config.family_justification_pct:
            reasons.append(f"Description-family price anomaly: {row.vs_family_pct:+.1f}% vs family median; justification required")
            severity = max(severity, 2); justify = True
        elif v > config.family_review_pct:
            reasons.append(f"Description-family price anomaly: {row.vs_family_pct:+.1f}% vs family median")
            severity = max(severity, 2)
        elif v > config.family_check_pct:
            reasons.append(f"Description-family variance: {row.vs_family_pct:+.1f}%")
            severity = max(severity, 1)
    elif service_item(item):
        if attachment.scope_named or attachment.contract_named or attachment.rfq_named:
            reasons.append("Service price cannot be validated from description; named scope/contract/RFQ attachment exists")
            severity = max(severity, 2 if row.history_count == 0 else 1)
        else:
            reasons.append("Service price cannot be validated from description and no named scope/contract/RFQ is listed")
            severity = max(severity, 3 if attachment.scan_status == "NONE" else 2)

    row.benchmark_strength = "STRONG" if row.history_tier in {"MATERIAL+SUPPLIER+PLANT", "MATERIAL+SUPPLIER"} and row.history_count >= 2 else ("MEDIUM" if row.history_count else "NONE")
    row.justification_required = "YES" if justify else "NO"
    row.decision = {0: "APPROVE", 1: "CHECK", 2: "REVIEW", 3: "HOLD"}[severity]
    if severity == 0 and attachment.scan_status == "PRESENT":
        reasons.insert(0, "Attachment presence confirmed; attachment content was not opened in FAST mode")
    row.reasons = list(dict.fromkeys(reasons))
    return row, matches


def worst_decision(values: Iterable[str]) -> str:
    return max(values, key=lambda x: DECISION_RANK.get(x, 2), default="REVIEW")


# ------------------------------- workbook --------------------------------
COLORS = {"navy": "17365D", "blue": "1F4E78", "light_blue": "D9EAF7", "green": "C6EFCE", "green_text": "006100", "yellow": "FFF2CC", "yellow_text": "9C6500", "orange": "FCE4D6", "red": "FFC7CE", "red_text": "9C0006", "gray": "E7E6E6", "white": "FFFFFF"}
THIN = Side(style="thin", color="B7B7B7")


def safe(value: Any) -> Any:
    if isinstance(value, (dict, list, tuple, set)):
        return json.dumps(value, ensure_ascii=False)
    return value


def write_sheet(ws, headers: list[str], rows: list[list[Any]], table_name: str, widths: Optional[dict[str, float]] = None) -> None:
    ws.append(headers)
    for row in rows:
        ws.append([safe(x) for x in row])
    for cell in ws[1]:
        cell.fill = PatternFill("solid", fgColor=COLORS["navy"])
        cell.font = Font(color=COLORS["white"], bold=True)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=THIN)
    ws.freeze_panes = "A2"
    # Structured tables already provide AutoFilter. Do not also set the
    # worksheet-level AutoFilter over the same range (Excel overlap warning).
    if ws.max_row >= 2:
        ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
        table = Table(displayName=table_name, ref=ref)
        table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True, showFirstColumn=False, showLastColumn=False, showColumnStripes=False)
        ws.add_table(table)
    for c, header in enumerate(headers, 1):
        max_len = len(header)
        for cell in ws.iter_cols(min_col=c, max_col=c, min_row=2, max_row=ws.max_row):
            value = cell[0].value
            max_len = max(max_len, min(60, len(text(value))))
            cell[0].alignment = Alignment(vertical="top", wrap_text=True)
        width = min(max(max_len + 2, 10), 38)
        if widths and header in widths:
            width = widths[header]
        ws.column_dimensions[get_column_letter(c)].width = width
    ws.row_dimensions[1].height = 34


def apply_decision_style(ws, column: int) -> None:
    fill_map = {"APPROVE": ("green", "green_text"), "CHECK": ("yellow", "yellow_text"), "REVIEW": ("orange", "yellow_text"), "HOLD": ("red", "red_text"), "DATA ERROR": ("red", "red_text")}
    for r in range(2, ws.max_row + 1):
        cell = ws.cell(r, column)
        key = text(cell.value).upper()
        if key in fill_map:
            bg, fg = fill_map[key]
            cell.fill = PatternFill("solid", fgColor=COLORS[bg])
            cell.font = Font(color=COLORS[fg], bold=True)
            cell.alignment = Alignment(horizontal="center", vertical="center")


def dashboard_rows(analysis: list[AnalysisRow]) -> list[list[Any]]:
    by_po: dict[str, list[AnalysisRow]] = defaultdict(list)
    for row in analysis:
        by_po[row.po_number].append(row)
    rows = []
    for po, items in sorted(by_po.items()):
        currencies = sorted({x.currency for x in items if x.currency})
        same_currency = len(currencies) == 1
        total = sum(x.line_value or 0 for x in items) if same_currency else None
        attach_count = max((x.attachment_count for x in items), default=0)
        attach_scan = next((x.attachment_scan for x in items if x.attachment_scan), "UNKNOWN")
        history_cov = sum(1 for x in items if x.history_count > 0)
        anomalies = sum(1 for x in items if x.description_assessment == "ANOMALY")
        decision = worst_decision(x.decision for x in items)
        reasons = []
        for x in items:
            if x.decision == decision:
                reasons.extend(x.reasons)
        rows.append([po, "; ".join(sorted({x.supplier for x in items if x.supplier})), len(items), total, currencies[0] if same_currency and currencies else "MIXED", attach_scan, attach_count, f"{history_cov}/{len(items)}", anomalies, "YES" if any(x.justification_required == "YES" for x in items) else "NO", decision, "; ".join(list(dict.fromkeys(reasons))[:4])])
    return rows


def create_workbook(path: Path, analysis: list[AnalysisRow], inventory: list[dict[str, str]], matches: list[HistoryMatch], run_status_rows: list[dict[str, str]], config: Config) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Dashboard"
    dash_headers = ["PO", "Supplier", "Items", "PO Value", "Currency", "Attachment Scan", "Attachment Count", "History Coverage", "Family Anomalies", "Justification Required", "Decision", "Key Reason"]
    write_sheet(ws, dash_headers, dashboard_rows(analysis), "tblFastDashboard", {"Supplier": 30, "Key Reason": 70})
    apply_decision_style(ws, dash_headers.index("Decision") + 1)
    for r in range(2, ws.max_row + 1): ws.cell(r, 4).number_format = "#,##0.00"

    headers = ["PO", "PO Date", "Supplier", "Item", "Description", "Material", "Material Group", "Plant", "Quantity", "UoM", "Net Price", "Price Unit", "Unit Price", "Currency", "Line Value", "Attachment Scan", "Attachment Count", "Attachment Types", "CS Named", "Quote Named", "History Tier", "History Count", "Previous Price", "Previous Date", "Median Last 3", "Vs Previous %", "Vs Median %", "History Match %", "Family Count", "Family Median", "Vs Family %", "Description Assessment", "Benchmark Strength", "Justification Required", "Decision", "Reasons"]
    detail_rows = []
    for x in analysis:
        detail_rows.append([x.po_number, x.po_date, x.supplier, x.item, x.description, x.material, x.material_group, x.plant, x.quantity, x.uom, x.net_price, x.price_unit, x.unit_price, x.currency, x.line_value, x.attachment_scan, x.attachment_count, x.attachment_types, x.cs_named, x.quote_named, x.history_tier, x.history_count, x.previous_price, x.previous_date, x.median_last3, x.vs_previous_pct, x.vs_median_pct, x.history_match_score * 100 if x.history_match_score is not None else None, x.family_count, x.family_median, x.vs_family_pct, x.description_assessment, x.benchmark_strength, x.justification_required, x.decision, "; ".join(x.reasons)])
    wd = wb.create_sheet("PO Detail")
    write_sheet(wd, headers, detail_rows, "tblFastDetail", {"Description": 42, "Supplier": 30, "Reasons": 70})
    apply_decision_style(wd, headers.index("Decision") + 1)
    for name in ["Net Price", "Price Unit", "Unit Price", "Line Value", "Previous Price", "Median Last 3", "Family Median"]:
        c = headers.index(name) + 1
        for r in range(2, wd.max_row + 1): wd.cell(r, c).number_format = "#,##0.00"
    for name in ["Vs Previous %", "Vs Median %", "History Match %", "Vs Family %"]:
        c = headers.index(name) + 1
        for r in range(2, wd.max_row + 1): wd.cell(r, c).number_format = "0.0"

    wi = wb.create_sheet("Attachment Inventory")
    ih = ["PO", "Row", "Description", "Category"]
    write_sheet(wi, ih, [[r["po_number"], r["row"], r["description"], r["category"]] for r in inventory], "tblFastAttachments", {"Description": 70})

    wh = wb.create_sheet("Historical Matches")
    hh = ["Current PO", "Current Item", "Source PO", "Source Item", "Source Date", "Source Supplier", "Source Description", "Source Material", "Material Group", "Plant", "Unit Price", "Currency", "UoM", "Match Tier", "Match %", "Supplier Match", "Plant Match"]
    hrows = [[m.po_number, m.item, m.source_po, m.source_item, m.source_date, m.source_supplier, m.source_description, m.source_material, m.source_material_group, m.source_plant, m.source_unit_price, m.source_currency, m.source_uom, m.tier, m.score * 100, "YES" if m.supplier_match else "NO", "YES" if m.plant_match else "NO"] for m in matches]
    write_sheet(wh, hh, hrows, "tblFastHistory", {"Source Supplier": 30, "Source Description": 45})

    we = wb.create_sheet("Exceptions")
    eh = ["PO", "Item", "Decision", "Description", "Justification Required", "Reason"]
    erows = []
    for x in analysis:
        if x.decision != "APPROVE":
            for reason in x.reasons:
                erows.append([x.po_number, x.item, x.decision, x.description, x.justification_required, reason])
    write_sheet(we, eh, erows, "tblFastExceptions", {"Description": 42, "Reason": 70})
    apply_decision_style(we, 3)

    wr = wb.create_sheet("Run Status")
    if run_status_rows:
        rh = list(run_status_rows[0].keys())
        write_sheet(wr, rh, [[r.get(h, "") for h in rh] for r in run_status_rows], "tblFastRunStatus", {"Detail": 70})
    else:
        write_sheet(wr, ["Message"], [["Fast_Run_Status.txt was not provided"]], "tblFastRunStatus")

    wp = wb.create_sheet("Parameters")
    params = [["Analyzer Version", VERSION], ["Mode", "FAST - attachment names only; attachment files are not opened"], ["History OK threshold", config.history_ok_pct], ["History review threshold", config.history_review_pct], ["History justification threshold", config.history_justification_pct], ["Family check threshold", config.family_check_pct], ["Family review threshold", config.family_review_pct], ["Family justification threshold", config.family_justification_pct], ["Important", "Attachment presence does not prove attachment quality or commercial correctness"]]
    write_sheet(wp, ["Parameter", "Value"], params, "tblFastParameters", {"Parameter": 32, "Value": 70})
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.forceFullCalc = True
    path.parent.mkdir(parents=True, exist_ok=True)
    wb.save(path)


# ------------------------------- CLI -------------------------------------
def default_base_dir() -> str:
    return str(Path(sys.executable).resolve().parent) if getattr(sys, "frozen", False) else "."


def parse_args(argv: Optional[list[str]] = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description="V15 FAST read-only PO analyzer")
    p.add_argument("--base-dir", default=default_base_dir())
    p.add_argument("--details", default="Pending_PO_Details.txt")
    p.add_argument("--attachments", default="Pending_PO_Attachments.txt")
    p.add_argument("--run-status", default="Fast_Run_Status.txt")
    p.add_argument("--history", default="")
    p.add_argument("--history-sheet", default="")
    p.add_argument("--output", default="PO_Approval_Cockpit_FAST.xlsx")
    p.add_argument("--json-output", default="V15_FAST_Analysis.json")
    p.add_argument("--log", default="V15_FAST_Run_Log.txt")
    p.add_argument("--history-ok-pct", type=float, default=3.0)
    p.add_argument("--history-review-pct", type=float, default=5.0)
    p.add_argument("--history-justification-pct", type=float, default=10.0)
    p.add_argument("--verbose", action="store_true")
    return p.parse_args(argv)


def resolve(base: Path, value: str) -> Path:
    p = Path(value).expanduser()
    return p.resolve() if p.is_absolute() else (base / p).resolve()


def find_history(base: Path, explicit: str) -> Optional[Path]:
    if explicit:
        return Path(explicit).expanduser().resolve()
    found = sorted(base.glob("EXPORT_*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    return found[0] if found else None


def main(argv: Optional[list[str]] = None) -> int:
    args = parse_args(argv)
    base = Path(args.base_dir).expanduser().resolve()
    details = resolve(base, args.details)
    attachments = resolve(base, args.attachments)
    run_status_path = resolve(base, args.run_status)
    output = resolve(base, args.output)
    json_output = resolve(base, args.json_output)
    log_path = resolve(base, args.log)
    log_path.parent.mkdir(parents=True, exist_ok=True)
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO, format="%(asctime)s\t%(levelname)s\t%(message)s", handlers=[logging.FileHandler(log_path, mode="w", encoding="utf-8"), logging.StreamHandler(sys.stdout)], force=True)
    config = Config(history_ok_pct=args.history_ok_pct, history_review_pct=args.history_review_pct, history_justification_pct=args.history_justification_pct)
    started = datetime.now()
    try:
        if not details.exists():
            raise FileNotFoundError(f"Required input not found: {details}")
        items = load_po_details(details)
        if not items:
            raise ValueError("No usable PO item rows were found")
        run_status = load_run_status(run_status_path)
        run_status_rows = read_delimited(run_status_path) if run_status_path.exists() else []
        attachment_map, inventory = load_attachments(attachments, {x.po_number for x in items}, run_status)
        history_path = find_history(base, args.history)
        history = load_history(history_path, args.history_sheet) if history_path else []
        logging.info("Loaded %s PO items, %s attachment names, %s history rows", len(items), len(inventory), len(history))
        analysis: list[AnalysisRow] = []
        all_matches: list[HistoryMatch] = []
        for item in items:
            result, matches = analyze_item(item, items, attachment_map[item.po_number], history, config, run_status)
            analysis.append(result)
            all_matches.extend(matches)
        create_workbook(output, analysis, inventory, all_matches, run_status_rows, config)
        payload = {"version": VERSION, "generated_at": datetime.now().isoformat(timespec="seconds"), "mode": "FAST_ATTACHMENT_LIST_ONLY", "config": asdict(config), "inputs": {"details": str(details), "attachments": str(attachments), "run_status": str(run_status_path), "history": str(history_path) if history_path else ""}, "summary": {"po_count": len({x.po_number for x in items}), "item_count": len(items), "attachment_name_count": len(inventory), "history_row_count": len(history), "decisions": dict(Counter(x.decision for x in analysis))}, "analysis": [asdict(x) for x in analysis]}
        json_output.write_text(json.dumps(payload, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
        elapsed = (datetime.now() - started).total_seconds()
        logging.info("Created %s in %.1fs", output.name, elapsed)
        print(f"\nV15 FAST ANALYSIS COMPLETE\nOutput: {output}\nJSON:   {json_output}\nLog:    {log_path}")
        return 0
    except Exception as exc:
        logging.error("V15 FAST failed: %s", exc)
        logging.debug(traceback.format_exc())
        print(f"\nV15 FAST ANALYSIS FAILED\n{exc}\nSee log: {log_path}", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
