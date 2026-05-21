"""
Indian bank statement -> grouped Deposits / Withdrawals workbook.

Architecture:
  detect_profile(path)  -> Profile (icici | axis | hdfc | sbi | generic)
  Profile.parse(path)   -> DataFrame[txn_date, remarks, withdrawal, deposit, balance, counterparty]
  consolidate(...)      -> merge near-duplicate counterparties
  write_workbook(...)   -> 2-sheet xlsx grouped by counterparty, threshold >=2

CLI:
  python sort_statement.py <statement_file> [output.xlsx]

Adding a new bank: write a new Profile subclass with detect() + parse_columns()
+ extract_counterparty() and register it in PROFILES.
"""

from __future__ import annotations

import csv
import io
import re
import sys
from collections import Counter, defaultdict
from pathlib import Path
from typing import Optional

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


# ============================================================
# Shared normalization / consolidation helpers
# ============================================================

_STOPWORDS = {
    "LTD", "LIMITED", "PVT", "PRIVATE", "CO", "COMPANY", "INC", "CORP",
    "INDIA", "THE", "OF", "AND", "PAYT", "BANK",
    "MR", "MRS", "MS", "DR", "SHRI", "SMT",
    "PAYMENT", "PAYMENTS", "TRANSFER", "ONLINE", "UPI",
}

# Generic free-text notes that aren't real counterparties.
_GENERIC_NOTES = {
    "UPISCANQR", "UPISENDMONEY", "UPI", "UPI TRANSACTION", "UPI PAYMENT",
    "PAYMENT", "COLLECT TRANSAC", "PAY VIA RAZORPA", "PAY TO BHARATPE",
    "SOME NOTE", "GENERATING DYNA", "CCA", "ONLINE", "TRANSFER",
    "OTHERS", "OTHPOS", "POS", "ATM CASH", "ATM WDL", "ATMWDL",
    "PAY TO", "PAY", "PAYTO", "PAYVIA", "COLLEC", "REFUND",
    "SELF TRANSFER", "FD THROUGH NET", "INB", "P2M", "P2A",
}

# Gateway VPAs are aggregators, not merchants.
_GATEWAY_VPA = {
    "PAYTM", "RAZORPAY", "BHARATPE", "PHONEPE", "GPAY", "GOOGLEPAY",
    "OKAXIS", "OKHDFCBANK", "OKSBI", "OKICICI", "AMAZONPAY",
}

# Merchant family aliases — collapse rail/gateway variants.
_MERCHANT_ALIASES = [
    (("ZOMATO", "PAYZOMATO", "ZOMATO ONLINE", "ZOMATO1PAYTM"), "ZOMATO"),
    (("SWIGGY", "BUNDL TECHNOLOGIES"), "SWIGGY"),
    (("BIGBASKET",), "BIGBASKET"),
    (("TATAPAY",), "TATA PAY (utilities)"),
    (("AMAZONPAY", "AMAZON PAY"), "AMAZON PAY"),
    (("DOMINOS",), "DOMINOS"),
    (("UBER",), "UBER"),
    (("DUNZO",), "DUNZO"),
    (("FRANKLIN", "FRANKLINTEMP"), "FRANKLIN TEMPELTON MUTUAL FUND"),
    (("INVESTEASYRMF",), "ICICI INVESTEASY"),
    (("APOLLO",), "APOLLO SPECIALTY HOSPITALS PVT LTD"),
    (("KIDS CLINIC", "KIDSCLINIC"), "KIDS CLINIC INDIA LIMITED"),
    (("AMERCIAN E", "AMEX", "AMERICAN EXPRESS", "AMERICANEX"), "AMERICAN EXPRESS"),
    (("ICICI BANK CREDIT CA",), "ICICI CREDIT CARD"),
    (("HDFCVI",), "HDFC CREDIT CARD"),
    (("BBMP",), "BBMP"),
    (("RAZPGROWW",), "GROWW (Razorpay)"),
    (("IMPERIAL HOSPIT",), "IMPERIAL HOSPITAL"),
    (("HUZUR TREA", "STATE HUZUR"), "STATE HUZUR TREASURY"),
    (("KOTAK CUSTODY",), "KOTAK CUSTODY (RTGS/NEFT)"),
]


def _norm(s) -> str:
    if not s:
        return ""
    s = re.sub(r"\s+", " ", str(s).strip().upper())
    s = re.sub(r"\s+\d+$", "", s)
    s = s.rstrip(" -_/.,;:")
    s = s.lstrip(" -_/.,;:")
    return s


def _apply_alias(key: str) -> str:
    if not key:
        return key
    up = key.upper()
    for patterns, canonical in _MERCHANT_ALIASES:
        for pat in patterns:
            if pat in up:
                return canonical
    return key


def _clean_vpa(v: str) -> Optional[str]:
    if "@" not in v:
        return None
    h = v.split("@")[0]
    if not h or re.fullmatch(r"\d{8,}", h):
        return None
    h = re.sub(r"\.\d+$", "", h)
    h = re.sub(r"\d+$", "", h)
    n = _norm(h)
    return n if n and len(n) >= 3 else None


def _compressed(s) -> str:
    return re.sub(r"[^A-Z0-9]", "", str(s).upper())


def _tokens(key) -> set[str]:
    if not key:
        return set()
    return {t for t in re.split(r"[^A-Z0-9]+", str(key))
            if t and t not in _STOPWORDS and len(t) > 2}


def consolidate(keys: list[str]) -> dict[str, str]:
    """Merge near-duplicate counterparties. Returns {raw: canonical}."""
    clean = sorted({str(k) for k in keys if k and isinstance(k, str) and k.strip()})
    if not clean:
        return {}

    groups: dict[str, list[str]] = defaultdict(list)
    for k in clean:
        toks = sorted(_tokens(k), key=lambda t: (-len(t), t))
        anchor = toks[0] if toks else _compressed(k)[:8] or k
        groups[anchor].append(k)

    group_keys = list(groups.keys())
    parent = {g: g for g in group_keys}

    def find(x):
        while parent[x] != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    def union(a, b):
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[ra] = rb

    reps = {g: _compressed(max(members, key=len)) for g, members in groups.items()}

    for i, g1 in enumerate(group_keys):
        r1 = reps[g1]
        if len(r1) < 6:
            continue
        for g2 in group_keys[i + 1:]:
            r2 = reps[g2]
            if len(r2) < 6:
                continue
            short, long_ = (r1, r2) if len(r1) <= len(r2) else (r2, r1)
            if len(short) >= 8 and short in long_:
                union(g1, g2)
                continue
            common = 0
            for c1, c2 in zip(r1, r2):
                if c1 != c2:
                    break
                common += 1
            if common >= 10:
                union(g1, g2)

    merged: dict[str, list[str]] = defaultdict(list)
    for g, members in groups.items():
        merged[find(g)].extend(members)

    return {m: max(members, key=len) for _, members in merged.items() for m in members}


# ============================================================
# Generic counterparty extractor (last-resort fallback for unknown formats)
# ============================================================

def extract_generic(remarks: str) -> Optional[str]:
    """When no bank profile recognised the remark, fall back to this.
    Strategy: pull the longest run of uppercase words (likely a name/merchant)
    or the longest non-numeric token in the remark."""
    r = str(remarks).strip()
    if not r:
        return None
    # Try uppercase-word runs first (e.g. 'MR KUMAR PRABHAS', 'KIDS CLINIC INDIA')
    runs = re.findall(r"\b([A-Z][A-Z &]{2,}(?:\s+[A-Z][A-Z &]*){0,5})\b", r)
    runs = [s.strip() for s in runs if s.strip() and _norm(s) not in _GENERIC_NOTES]
    if runs:
        return _norm(max(runs, key=len))
    # Fall back: longest alpha token
    tokens = re.findall(r"[A-Za-z]{4,}", r)
    tokens = [t for t in tokens if _norm(t) not in _GENERIC_NOTES]
    if tokens:
        return _norm(max(tokens, key=len))
    return None


# ============================================================
# Profiles
# ============================================================

class Profile:
    name = "base"

    @classmethod
    def detect(cls, path: Path, raw_text: str, raw_df: Optional[pd.DataFrame]) -> bool:
        return False

    @classmethod
    def parse(cls, path: Path, header_row: Optional[int] = None) -> pd.DataFrame:
        raise NotImplementedError

    @classmethod
    def extract(cls, remarks: str) -> Optional[str]:
        return extract_generic(remarks)


def _load_raw(path: Path):
    """Return (raw_text_first_8k, raw_df_or_None). Used for detection and parsing.

    Tries (in order): xlrd (.xls BIFF), openpyxl (.xlsx), pd.read_html
    (many bank exports are HTML wrapped in a .xls extension)."""
    raw_text = ""
    try:
        with open(path, "rb") as f:
            raw_text = f.read(8192).decode("utf-8", errors="replace")
    except Exception:
        pass
    raw_df = None
    for engine in ("xlrd", "openpyxl"):
        try:
            raw_df = pd.read_excel(path, sheet_name=0, engine=engine, header=None)
            break
        except Exception:
            continue
    if raw_df is None:
        try:
            tables = pd.read_html(str(path), header=None)
            if tables:
                raw_df = max(tables, key=lambda t: t.shape[0]).reset_index(drop=True)
                raw_df.columns = range(raw_df.shape[1])
        except Exception:
            pass
    return raw_text, raw_df


# ============================================================
# Shared header-driven row parser
# ============================================================

_DATE_TOKENS = ("txn date", "tran date", "transaction date", "value date", "date")
_REMARKS_TOKENS = ("transaction remarks", "narration", "particulars",
                   "description", "remarks")
_DEBIT_TOKENS = ("withdrawal amt", "withdrawal", "debit amt", "debit", "dr")
_CREDIT_TOKENS = ("deposit amt", "deposit", "credit amt", "credit", "cr")
_BALANCE_TOKENS = ("closing balance", "balance", "bal")


def _header_match(needle: str, cell: str) -> bool:
    """Word-boundary needle match. Prevents 'cr' from hitting 'description'."""
    if not cell:
        return False
    return re.search(rf"\b{re.escape(needle)}", cell) is not None


def _find_header_row_any(df: pd.DataFrame) -> Optional[int]:
    """Find the row that looks like a transaction header.

    Strong match: date + remarks + amount tokens all present.
    Weak match (fallback): date + amount tokens present (remarks column is
    inferred later as the widest non-date/non-amount text column).
    """
    weak_match: Optional[int] = None
    for i in range(min(60, len(df))):
        row_lc = [str(v).lower().strip() for v in df.iloc[i].tolist()]
        has_date = any(any(_header_match(t, v) for t in _DATE_TOKENS) for v in row_lc)
        has_rem = any(any(_header_match(t, v) for t in _REMARKS_TOKENS) for v in row_lc)
        has_amt = any(any(_header_match(t, v) for t in _DEBIT_TOKENS + _CREDIT_TOKENS)
                      for v in row_lc)
        if has_date and has_rem and has_amt:
            return i
        if has_date and has_amt and weak_match is None:
            weak_match = i
    return weak_match


def _pick_col(header: list[str], needles: tuple[str, ...],
              avoid: Optional[set[int]] = None) -> Optional[int]:
    """Return the column index whose lowercase header matches one of the needles.

    Iterates needles outer-loop first so the most specific token (e.g. 'txn date')
    wins over a less specific one (e.g. 'date' matching 'Value Date')."""
    avoid = avoid or set()
    for n in needles:
        for idx, h in enumerate(header):
            if idx in avoid:
                continue
            if _header_match(n, h):
                return idx
    return None


class HeaderNotFoundError(RuntimeError):
    """Raised when the header row can't be auto-detected.

    Carries a structured preview of the first rows so the UI can ask the user
    to pick the header row manually.
    """

    def __init__(self, preview: list[dict]):
        self.preview = preview
        super().__init__(
            "Couldn't locate transaction header row. Expected a row containing "
            "a date column (e.g. 'Date', 'Txn Date') and an amount column "
            "(e.g. 'Debit', 'Credit', 'Withdrawal')."
        )


def _classify_columns(body: pd.DataFrame) -> dict[int, str]:
    """Tag each column as 'date', 'amount', 'text', or 'empty' from its data.

    Used as a fallback when header tokens don't match (e.g. non-English headers
    like 'Tarikh' / 'Money Out') so the user-confirmed header row still yields a
    usable parse."""
    classes: dict[int, str] = {}
    for idx in range(body.shape[1]):
        col = body.iloc[:, idx]
        non_null = col[col.notna()].astype(str).str.strip()
        non_null = non_null[non_null.ne("") & non_null.str.lower().ne("nan")]
        if len(non_null) == 0:
            classes[idx] = "empty"
            continue
        cleaned = non_null.str.replace(",", "", regex=False) \
            .str.replace("₹", "", regex=False).str.strip().str.strip('"').str.strip()
        num_ratio = pd.to_numeric(cleaned, errors="coerce").notna().mean()
        # Suppress noisy dateutil warnings while sniffing
        import warnings
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            dt_ratio = pd.to_datetime(non_null, errors="coerce", dayfirst=True).notna().mean()
        if dt_ratio >= 0.7 and dt_ratio >= num_ratio:
            classes[idx] = "date"
        elif num_ratio >= 0.7:
            classes[idx] = "amount"
        else:
            classes[idx] = "text"
    return classes


def _build_preview(raw: pd.DataFrame, limit: int = 15) -> list[dict]:
    preview = []
    for i in range(min(limit, len(raw))):
        cells = [str(v) for v in raw.iloc[i].tolist() if str(v).strip() and str(v) != "nan"]
        if cells:
            preview.append({"index": i, "text": " | ".join(cells)[:200]})
    return preview


def _parse_from_df(
    raw: pd.DataFrame,
    extract_fn,
    header_row: Optional[int] = None,
) -> pd.DataFrame:
    """Parse a transaction DataFrame by locating columns from header text.

    Used by every bank profile so detection format (xls / xlsx / html) doesn't
    matter — we only care about header names. Pass ``header_row`` to skip
    auto-detection and use that row index as the header."""
    if header_row is not None:
        if header_row < 0 or header_row >= len(raw):
            raise RuntimeError(
                f"Header row {header_row} is out of range (file has {len(raw)} rows)."
            )
        hdr = header_row
    else:
        hdr = _find_header_row_any(raw)
        if hdr is None:
            raise HeaderNotFoundError(_build_preview(raw))
    header = [str(v).lower().strip() for v in raw.iloc[hdr].tolist()]
    body = raw.iloc[hdr + 1:]

    col_date = _pick_col(header, _DATE_TOKENS)
    col_wd = _pick_col(header, _DEBIT_TOKENS)
    # Don't let deposit column collide with the debit column
    col_dep = _pick_col(header, _CREDIT_TOKENS, avoid={col_wd} if col_wd is not None else None)
    col_bal = _pick_col(header, _BALANCE_TOKENS)
    col_rem = _pick_col(header, _REMARKS_TOKENS)

    # Data-based inference for whatever the header tokens didn't resolve.
    # Lets the parser handle non-English / non-standard headers (e.g. 'Money Out',
    # 'Tarikh') once the user has confirmed the header row.
    classified = _classify_columns(body)
    used = {c for c in (col_date, col_wd, col_dep, col_bal, col_rem) if c is not None}

    if col_date is None:
        for idx, kind in classified.items():
            if idx not in used and kind == "date":
                col_date = idx
                used.add(idx)
                break

    if col_wd is None or col_dep is None:
        amount_cols = [idx for idx, kind in classified.items()
                       if idx not in used and kind == "amount"]
        # Header text hints — even if not exact matches — disambiguate debit vs credit
        debit_hints = ("out", "debit", "dr", "withdraw", "paid", "spent")
        credit_hints = ("in", "credit", "cr", "deposit", "received", "in ", "earned")

        def _hint_side(idx):
            h = header[idx] if idx < len(header) else ""
            if any(t in h for t in debit_hints):
                return "wd"
            if any(t in h for t in credit_hints):
                return "dep"
            return None

        unassigned = []
        for idx in amount_cols:
            side = _hint_side(idx)
            if side == "wd" and col_wd is None:
                col_wd = idx
            elif side == "dep" and col_dep is None:
                col_dep = idx
            else:
                unassigned.append(idx)

        # Fall back to column order convention: debit before credit
        for idx in unassigned:
            if col_wd is None:
                col_wd = idx
            elif col_dep is None:
                col_dep = idx
        used = {c for c in (col_date, col_wd, col_dep, col_bal, col_rem) if c is not None}

    if col_rem is None:
        # Widest text column that isn't date/amount/balance.
        best_idx, best_width = None, 0
        for idx in range(body.shape[1]):
            if idx in used:
                continue
            col_vals = body.iloc[:, idx].astype(str)
            width = col_vals.str.len().mean() if len(col_vals) else 0
            if width > best_width:
                best_width, best_idx = width, idx
        col_rem = best_idx

    df = raw.iloc[hdr + 1:].copy().reset_index(drop=True)

    def col(i):
        return df.iloc[:, i] if i is not None else pd.Series([None] * len(df))

    def numify(series):
        # Strip currency symbols / commas / quotes before to_numeric
        cleaned = series.astype(str).str.replace(",", "", regex=False) \
            .str.replace("₹", "", regex=False).str.strip().str.strip('"').str.strip()
        return pd.to_numeric(cleaned, errors="coerce").fillna(0)

    out = pd.DataFrame({
        "txn_date": col(col_date),
        "remarks": col(col_rem),
        "withdrawal": numify(col(col_wd)),
        "deposit": numify(col(col_dep)),
        "balance": pd.to_numeric(
            col(col_bal).astype(str).str.replace(",", "", regex=False).str.strip(),
            errors="coerce",
        ),
    })
    out = out.dropna(subset=["remarks"])
    # Drop separator rows (asterisks, header repeats, etc.)
    out = out[~out["remarks"].astype(str).str.match(r"^\s*[\*\-=_]{3,}\s*$")]
    out = out[(out["withdrawal"] > 0) | (out["deposit"] > 0)].reset_index(drop=True)
    out["counterparty"] = out["remarks"].map(extract_fn)
    return out


def _find_header_row(df: pd.DataFrame, *needles: str) -> Optional[int]:
    """Return first row whose joined string contains ALL needles (lowercase substring match)."""
    needles_lc = [n.lower() for n in needles]
    for i in range(min(50, len(df))):
        row_str = " | ".join(str(v).lower() for v in df.iloc[i].tolist())
        if all(n in row_str for n in needles_lc):
            return i
    return None


# ---------- ICICI ----------

class ICICIProfile(Profile):
    name = "icici"

    @classmethod
    def detect(cls, path, raw_text, raw_df):
        text_up = raw_text.upper()
        if "ICICI" in text_up and "TRANSACTION REMARKS" in text_up:
            return True
        if raw_df is None:
            return False
        hdr = _find_header_row(raw_df, "transaction remarks")
        return hdr is not None

    @classmethod
    def parse(cls, path, header_row=None):
        _, raw = _load_raw(path)
        if raw is None:
            raise RuntimeError("ICICI: couldn't load file as a table")
        return _parse_from_df(raw, cls.extract, header_row=header_row)

    @classmethod
    def extract(cls, remarks):
        return _icici_extract(remarks)


def _icici_extract(remarks):
    r = str(remarks).strip()
    if not r:
        return None
    raw = _icici_extract_raw(r)
    if raw is None:
        return None
    if _norm(raw) in _GENERIC_NOTES:
        return None
    return _apply_alias(raw)


def _icici_extract_raw(r):
    # UPI/[<ref>/]<note>/<vpa>/<bank>/<txnid>  OR  UPI/<vpa>/<note>/<bank>/...
    if r.startswith("UPI/"):
        parts = [p.strip() for p in r.split("/")[1:]]
        vpas = [p for p in parts if "@" in p]
        for v in vpas:
            cleaned = _clean_vpa(v)
            if cleaned and cleaned not in _GATEWAY_VPA:
                if "." in cleaned:
                    head, _, tail = cleaned.partition(".")
                    if head in _GATEWAY_VPA and tail:
                        return tail
                return cleaned
        BANK_SUFFIXES = ("BANK", "BANK LTD", "BANK LIMITED", "BANK OF I", "BANK OF INDIA")
        for p in parts:
            if not p or p == "-" or p.isdigit() or "@" in p:
                continue
            up = p.upper()
            if any(s in up for s in BANK_SUFFIXES):
                continue
            if len(p) >= 10 and any(c.isdigit() for c in p) and " " not in p:
                continue
            n = _norm(p)
            if n and n not in _GENERIC_NOTES and len(n) >= 4:
                return n
        for v in vpas:
            cleaned = _clean_vpa(v)
            if cleaned:
                return cleaned
        return None

    if r.startswith("NEFT-") or r.startswith("NEFT/"):
        body = r[5:]
        parts = re.split(r"[-/]", body, maxsplit=2)
        if len(parts) >= 2:
            return _norm(parts[1]) or None
        return None

    if r.startswith("RTGS"):
        body = r[5:] if r.startswith(("RTGS-", "RTGS/")) else r[4:]
        parts = re.split(r"[-/]", body, maxsplit=2)
        if len(parts) >= 2:
            return _norm(parts[1]) or None
        return None

    if r.startswith("BIL/"):
        BIL_SKIP = {"ONL", "INFT", "CCA", "BBPS", "ECU", "BIL"}
        for p in r.split("/")[1:]:
            p2 = p.strip()
            if not p2 or p2 in BIL_SKIP or p2.isdigit():
                continue
            if re.fullmatch(r"[A-Z]{2,4}\d{4,}", p2):
                continue
            return _norm(p2) or None
        return None

    if r.startswith("CMS/"):
        parts = r.split("/")
        if len(parts) >= 3:
            return _norm(parts[2]) or None
        return None

    if r.startswith("ACH/"):
        sub = r.split("/")[1] if "/" in r else ""
        sub_parts = sub.split("-")
        if len(sub_parts) >= 2:
            return _norm(sub_parts[1]) or None
        return _norm(sub) or None

    if r.startswith("CLG/"):
        parts = r.split("/")
        if len(parts) >= 2:
            return _norm(parts[1]) or None
        return None

    if r.startswith("VPS/"):
        parts = r.split("/")
        if len(parts) >= 2:
            return _norm(parts[1]) or None
        return None

    if r.startswith("MMT/") or r.startswith("IMPS/"):
        parts = r.split("/")
        for p in parts[2:]:
            p2 = p.strip()
            if p2 and not p2.isdigit():
                return _norm(p2) or None
        return None

    if r.startswith("GIB/"):
        parts = r.split("/")
        if len(parts) >= 3:
            return _norm(parts[2]) or None
        return None

    if r.startswith("CAM/"):
        if "CASH WDL" in r.upper():
            return "CASH WDL"
        parts = r.split("/")
        if len(parts) >= 3:
            return _norm(parts[2]) or None
        return None

    if r.startswith("NFS/") or r.startswith("ATM/"):
        return "ATM CASH WDL"

    if re.match(r"^\d{6,}:Int\.Pd", r):
        return "SB INTEREST"

    return None


# ---------- Axis ----------

class AxisProfile(Profile):
    name = "axis"

    @classmethod
    def detect(cls, path, raw_text, raw_df):
        text_up = raw_text.upper()
        if "PARTICULARS" in text_up and ("UTIB" in text_up or "AXIS BANK" in text_up):
            return True
        if raw_df is None:
            return False
        hdr = _find_header_row(raw_df, "particulars")
        if hdr is None:
            return False
        head_text = " ".join(str(v) for v in raw_df.iloc[:hdr].values.flatten() if v == v)
        return "UTIB" in head_text or "AXIS" in head_text.upper()

    @classmethod
    def parse(cls, path, header_row=None):
        _, raw = _load_raw(path)
        if raw is None:
            raise RuntimeError("Axis: couldn't load file as a table")
        return _parse_from_df(raw, cls.extract, header_row=header_row)

    @classmethod
    def extract(cls, remarks):
        r = str(remarks).strip()
        if not r:
            return None

        # UPI/P2M|P2A/<ref>/<MERCHANT>/<note>/<bank>
        if r.startswith("UPI/") or r.startswith("IMPS/"):
            parts = [p.strip() for p in r.split("/")]
            # Skip prefix (UPI/IMPS), then P2M/P2A, then ref, then merchant
            if len(parts) >= 4 and parts[1] in ("P2M", "P2A"):
                merchant = parts[3]
                # Sometimes the merchant is followed by trailing dashes/spaces - clean
                n = _norm(merchant)
                if n and n not in _GENERIC_NOTES and len(n) >= 3:
                    return _apply_alias(n)
            # Fallback: any non-empty, non-numeric slot
            for p in parts[2:]:
                if not p or p.isdigit() or "@" in p or p in ("P2M", "P2A"):
                    continue
                n = _norm(p)
                if n and n not in _GENERIC_NOTES and len(n) >= 4 and "BANK" not in n:
                    return _apply_alias(n)
            return None

        # ATM-CASH-AXIS/<atmid>/<...>
        if r.startswith("ATM-CASH"):
            return "ATM CASH WDL"

        # Long format with date/ref like 'IMPERIAL HOSPIT/12042024Imperial117'
        if "/" in r:
            head = r.split("/")[0]
            n = _norm(head)
            if n and n not in _GENERIC_NOTES and len(n) >= 4:
                return _apply_alias(n)

        return _apply_alias(extract_generic(r)) if extract_generic(r) else None


# ---------- HDFC ----------

class HDFCProfile(Profile):
    name = "hdfc"

    @classmethod
    def detect(cls, path, raw_text, raw_df):
        text_up = raw_text.upper()
        # Text-based detection (works for HTML-format .xls files where raw_df is None)
        if "HDFC" in text_up and "NARRATION" in text_up:
            return True
        if raw_df is None:
            return False
        hdr = _find_header_row(raw_df, "narration")
        if hdr is None:
            return False
        head_text = " ".join(str(v) for v in raw_df.iloc[:hdr].values.flatten() if v == v)
        return "HDFC" in head_text.upper()

    @classmethod
    def parse(cls, path, header_row=None):
        _, raw = _load_raw(path)
        if raw is None:
            raise RuntimeError("HDFC: couldn't load file as a table")
        return _parse_from_df(raw, cls.extract, header_row=header_row)

    @classmethod
    def extract(cls, remarks):
        r = str(remarks).strip()
        if not r:
            return None

        # PRIN AND INT AUTO_REDEEM <acc>  -> FD auto-redemption (interest+principal)
        if r.startswith("PRIN AND INT") or "AUTO_REDEEM" in r:
            return "FD AUTO-REDEEM"

        # CBDT/BANK REFERENCE NO:...  -> Income tax (CBDT) payment
        if r.startswith("CBDT/") or "CBDT" in r[:10]:
            return "CBDT (Income Tax)"

        # MC ISSUED - <branch> - <ref> -  - **** <NAME> ****  (manager's cheque)
        m = re.search(r"\*{2,}\s*([^\*]+?)\s*\*{2,}", r)
        if m and "MC ISSUED" in r:
            return _apply_alias(_norm(m.group(1)))

        # SELF - CHQ PAID - <branch>  -> own withdrawal
        if r.startswith("SELF - CHQ PAID") or r.startswith("SELF-CHQ"):
            return "SELF CHEQUE"

        # ACH D- <merchant>-<ref>  or  ACH C-<merchant>-<ref>
        m = re.match(r"^ACH\s+[CD]-\s*([^-]+?)-\d", r)
        if m:
            return _apply_alias(_norm(m.group(1)))

        # NEFT CR-<IFSC>-<NAME>-<rest>  or  RTGS CR-...  or  NEFT DR-...
        m = re.match(r"^(?:NEFT|RTGS|IMPS)\s+(?:CR|DR)-([^-]+)-(.+?)(?:-|$)", r)
        if m:
            name = m.group(2).strip()
            n = _norm(name)
            if n:
                return _apply_alias(n)

        # FT -<NAME> DR - <acc> - <NAME>
        m = re.match(r"^FT\s*-\s*([^\-]+)(?:\s+DR|\s+CR)", r)
        if m:
            return _apply_alias(_norm(m.group(1)))

        # CHQ PAID-MICR CTS-CH-<NAME>
        m = re.match(r"^CHQ PAID-MICR CTS-CH-(.+)$", r)
        if m:
            return _apply_alias(_norm(m.group(1)))

        # IB BILLPAY DR-<merchant>-<acc>
        m = re.match(r"^IB BILLPAY (?:DR|CR)-([^-]+)-", r)
        if m:
            return _apply_alias(_norm(m.group(1)))

        # FD THROUGH NET-<ref>:<NAME>  -> bucket as 'FD'
        if r.startswith("FD THROUGH NET") or r.startswith("FD-OPENED"):
            return "FD (Fixed Deposit creation)"

        # INTEREST CREDIT <acc>
        if r.startswith("INTEREST CREDIT"):
            return "SB INTEREST"

        # POS <ref> <MERCHANT> <city>
        m = re.match(r"^POS\s+\d+\s+(.+?)(?:\s+[A-Z]{3,}\s*$)?$", r)
        if m:
            return _apply_alias(_norm(m.group(1)))

        # UPI-<merchant>-<vpa>-<...>
        if r.startswith("UPI-") or r.startswith("UPI/"):
            parts = re.split(r"[-/]", r)
            for p in parts[1:]:
                p2 = p.strip()
                if not p2 or p2.isdigit() or "@" in p2:
                    continue
                n = _norm(p2)
                if n and n not in _GENERIC_NOTES and len(n) >= 3:
                    return _apply_alias(n)

        # RFX <date><ref> <description>  (forex transactions)
        if r.startswith("RFX ") or re.match(r"^\d{6}RTT\d", r):
            return "FOREX TRANSACTION"

        # NWD-<atm>-<rest>
        if r.startswith("NWD-") or "ATM" in r[:10]:
            return "ATM CASH WDL"

        return _apply_alias(extract_generic(r)) if extract_generic(r) else None


# ---------- SBI ----------

class SBIProfile(Profile):
    name = "sbi"

    @classmethod
    def detect(cls, path, raw_text, raw_df):
        text_up = raw_text.upper()
        if "STATE BANK OF INDIA" in text_up:
            return True
        # Tab-separated text export — be lenient about the IFSC label variant
        if "ACCOUNT NAME" in text_up and ("IFS CODE" in text_up
                                          or "IFSC" in text_up
                                          or "IFS (INDIAN FINANCIAL SYSTEM)" in text_up
                                          or "SBIN0" in text_up):
            return True
        return False

    @classmethod
    def parse(cls, path, header_row=None):
        # If the file loads as a table (HTML-wrapped .xls or real .xls/.xlsx),
        # use the shared header-driven parser.
        _, raw_df = _load_raw(path)
        if raw_df is not None and (
            header_row is not None or _find_header_row_any(raw_df) is not None
        ):
            return _parse_from_df(raw_df, cls.extract, header_row=header_row)

        # Otherwise fall back to the tab-separated text export.
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            text = f.read()
        lines = text.split("\n")
        hdr_idx = None
        for i, ln in enumerate(lines):
            low = ln.lower()
            if "txn date" in low and "description" in low and "debit" in low:
                hdr_idx = i
                break
        if hdr_idx is None:
            raise RuntimeError("SBI: couldn't locate transaction header line")

        header_fields = [h.strip().lower() for h in lines[hdr_idx].split("\t")]

        def find_col(*needles):
            for n in needles:
                for idx, h in enumerate(header_fields):
                    if _header_match(n, h):
                        return idx
            return None

        col_date = find_col(*_DATE_TOKENS)
        col_rem = find_col(*_REMARKS_TOKENS)
        col_wd = find_col(*_DEBIT_TOKENS)
        avoid = {col_wd} if col_wd is not None else set()
        col_dep = None
        for n in _CREDIT_TOKENS:
            for idx, h in enumerate(header_fields):
                if idx in avoid:
                    continue
                if _header_match(n, h):
                    col_dep = idx
                    break
            if col_dep is not None:
                break
        col_bal = find_col(*_BALANCE_TOKENS)

        def _to_num(s):
            s = str(s).strip().strip('"').replace(",", "").strip()
            if not s or s.lower() == "nan":
                return 0.0
            try:
                return float(s)
            except ValueError:
                return 0.0

        rows = []
        for ln in lines[hdr_idx + 1:]:
            if not ln.strip():
                continue
            fields = ln.split("\t")
            if len(fields) <= max(c for c in (col_date, col_rem, col_wd, col_dep) if c is not None):
                continue
            rows.append(fields)

        def get(fields, i):
            return fields[i] if (i is not None and i < len(fields)) else ""

        df = pd.DataFrame({
            "txn_date": [get(r, col_date) for r in rows],
            "remarks": [get(r, col_rem) for r in rows],
            "withdrawal": [_to_num(get(r, col_wd)) for r in rows],
            "deposit": [_to_num(get(r, col_dep)) for r in rows],
            "balance": [_to_num(get(r, col_bal)) for r in rows],
        })
        df = df[(df["withdrawal"] > 0) | (df["deposit"] > 0)].reset_index(drop=True)
        df["counterparty"] = df["remarks"].map(cls.extract)
        return df[["txn_date", "remarks", "withdrawal", "deposit", "balance", "counterparty"]]

    @classmethod
    def extract(cls, remarks):
        r = str(remarks).strip().lstrip().rstrip("-").strip()
        if not r:
            return None

        # TO TRANSFER-INB <merchant>  or  TO TRANSFER-UPI/<...>
        m = re.match(r"^TO TRANSFER-INB\s+(.+?)(?:\s*--)?$", r)
        if m:
            cleaned = re.sub(r"\s+Payments?\s*$", "", m.group(1), flags=re.I).strip()
            return _apply_alias(_norm(cleaned))

        # BY TRANSFER-NEFT*<ifsc>*<ref>*<NAME>  -> NAME slot
        m = re.match(r"^BY TRANSFER-NEFT\*[^*]+\*[^*]+\*(.+?)$", r)
        if m:
            return _apply_alias(_norm(m.group(1)))

        # BY TRANSFER-INB or TO TRANSFER-NEFT (without stars)
        m = re.match(r"^(?:TO|BY) TRANSFER-(?:NEFT|RTGS|IMPS)[*\-](.+?)$", r)
        if m:
            tail = m.group(1)
            # If contains '*', take last *-segment
            if "*" in tail:
                segs = [s for s in tail.split("*") if s.strip()]
                if segs:
                    return _apply_alias(_norm(segs[-1]))
            return _apply_alias(_norm(tail))

        # ATM WDL-ATM CASH <atmid> <loc> -> bucket
        if r.startswith("ATM WDL"):
            return "ATM CASH WDL"

        # TO CLEARING-Chq No. <num> <branch> <NAME>--<num>
        m = re.match(r"^TO CLEARING-Chq No\.\s+\d+\s+\w+\s+(.+?)(?:--\d+)?$", r)
        if m:
            name = re.sub(r"\s*--\s*\d+\s*$", "", m.group(1))
            return _apply_alias(_norm(name))

        # by debit card-OTHPOS<ref> <MERCHANT> <city>
        m = re.match(r"^by debit card-(?:OTHPOS|POS)\d*\s+(.+?)\s+[A-Z]+\s*$", r, re.I)
        if m:
            return _apply_alias(_norm(m.group(1)))

        # TO TRANSFER-UPI/DR/<ref>/<merchant>/<...> or similar
        m = re.match(r"^TO TRANSFER-UPI/[A-Z]+/\d+/([^/]+)/", r)
        if m:
            return _apply_alias(_norm(m.group(1)))

        # BY TRANSFER-Other Bank ... <NAME>
        m = re.match(r"^BY TRANSFER-(.+?)$", r)
        if m:
            tail = m.group(1)
            for sep in ["*", "/"]:
                if sep in tail:
                    segs = [s for s in tail.split(sep) if s.strip() and not s.strip().isdigit()]
                    if segs:
                        return _apply_alias(_norm(segs[-1]))
            return _apply_alias(_norm(tail))

        return _apply_alias(extract_generic(r)) if extract_generic(r) else None


# ---------- Generic fallback ----------

class GenericProfile(Profile):
    name = "generic"

    @classmethod
    def detect(cls, path, raw_text, raw_df):
        return True  # always last in registry

    @classmethod
    def parse(cls, path, header_row=None):
        _, raw = _load_raw(path)
        if raw is not None:
            return _parse_from_df(raw, cls.extract, header_row=header_row)
        # File didn't load as Excel/HTML — try tab/comma-separated text.
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            text = f.read()
        sep = "\t" if text.count("\t") > text.count(",") else ","
        rows = list(csv.reader(io.StringIO(text), delimiter=sep))
        df = pd.DataFrame(rows)
        return _parse_from_df(df, cls.extract, header_row=header_row)

    @classmethod
    def extract(cls, remarks):
        n = extract_generic(remarks)
        return _apply_alias(n) if n else None


# Order matters: more-specific profiles first; Generic always last.
PROFILES = [HDFCProfile, AxisProfile, SBIProfile, ICICIProfile, GenericProfile]


def detect_profile(path: Path) -> type[Profile]:
    raw_text, raw_df = _load_raw(path)
    for prof in PROFILES:
        try:
            if prof.detect(path, raw_text, raw_df):
                return prof
        except Exception:
            continue
    return GenericProfile


# ============================================================
# Writer (unchanged from v1)
# ============================================================

THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
HDR_FILL = PatternFill("solid", start_color="305496")
GROUP_FILL = PatternFill("solid", start_color="D9E1F2")
SUBTOTAL_FILL = PatternFill("solid", start_color="FCE4D6")
GRAND_FILL = PatternFill("solid", start_color="FFE699")

HDR_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
GROUP_FONT = Font(name="Calibri", size=11, bold=True, color="1F3864")
BOLD = Font(name="Calibri", size=11, bold=True)
REG = Font(name="Calibri", size=11)


def _write_sheet(ws, title, df, amount_col, threshold=2):
    raw_keys = [k for k in df["counterparty"] if k]
    canon_map = consolidate(raw_keys)
    df = df.copy()
    df["bucket"] = df["counterparty"].map(lambda k: canon_map.get(k) if k else None)
    df["bucket"] = df["bucket"].fillna("UNSORTED")

    counts = df["bucket"].value_counts()
    named = [b for b in counts.index if b != "UNSORTED" and counts[b] >= threshold]
    totals = df.groupby("bucket")[amount_col].sum()
    named.sort(key=lambda b: -totals.get(b, 0))

    other_df = df[~df["bucket"].isin(named)].copy()
    other_df["bucket"] = "OTHER"

    ws["A1"] = title
    ws["A1"].font = Font(name="Calibri", size=14, bold=True, color="1F3864")
    ws.merge_cells("A1:E1")

    for i, h in enumerate(["Date", "Counterparty (raw)", "Remarks", "Amount", "Balance"], start=1):
        c = ws.cell(row=3, column=i, value=h)
        c.font = HDR_FONT
        c.fill = HDR_FILL
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = BORDER

    row = 4

    def write_group(label, sub):
        nonlocal row
        gc = ws.cell(row=row, column=1, value=label)
        gc.font = GROUP_FONT
        gc.fill = GROUP_FILL
        gc.border = BORDER
        for col in range(2, 6):
            ws.cell(row=row, column=col).fill = GROUP_FILL
            ws.cell(row=row, column=col).border = BORDER
        row += 1
        first = row
        for _, r in sub.iterrows():
            ws.cell(row=row, column=1, value=str(r["txn_date"])).font = REG
            ws.cell(row=row, column=2, value=r["counterparty"] or "").font = REG
            ws.cell(row=row, column=3, value=str(r["remarks"])).font = REG
            ws.cell(row=row, column=4, value=float(r[amount_col])).font = REG
            ws.cell(row=row, column=4).number_format = "#,##0.00"
            bal = r["balance"]
            if pd.notna(bal):
                ws.cell(row=row, column=5, value=float(bal)).font = REG
                ws.cell(row=row, column=5).number_format = "#,##0.00"
            for col in range(1, 6):
                ws.cell(row=row, column=col).border = BORDER
            row += 1
        last = row - 1
        st = ws.cell(row=row, column=3, value="Subtotal")
        st.font = BOLD
        st.alignment = Alignment(horizontal="right")
        st.fill = SUBTOTAL_FILL
        amt = ws.cell(row=row, column=4, value=f"=SUM(D{first}:D{last})")
        amt.font = BOLD
        amt.number_format = "#,##0.00"
        amt.fill = SUBTOTAL_FILL
        for col in range(1, 6):
            ws.cell(row=row, column=col).fill = SUBTOTAL_FILL
            ws.cell(row=row, column=col).border = BORDER
        row += 1
        row += 1  # spacer

    def _date_sort(sub_df):
        return sub_df.sort_values(
            "txn_date",
            key=lambda s: pd.to_datetime(s, errors="coerce", dayfirst=True),
            kind="stable",
            na_position="last",
        )

    for b in named:
        write_group(b, _date_sort(df[df["bucket"] == b]))
    if not other_df.empty:
        write_group("OTHER (one-offs)", _date_sort(other_df))

    gt_label = ws.cell(row=row, column=3, value="GRAND TOTAL")
    gt_label.font = BOLD
    gt_label.alignment = Alignment(horizontal="right")
    gt_label.fill = GRAND_FILL
    gt_amt = ws.cell(row=row, column=4, value=f"=SUM(D4:D{row-1})/2")
    gt_amt.font = BOLD
    gt_amt.number_format = "#,##0.00"
    gt_amt.fill = GRAND_FILL
    for col in range(1, 6):
        ws.cell(row=row, column=col).fill = GRAND_FILL
        ws.cell(row=row, column=col).border = BORDER

    for i, w in enumerate([12, 38, 60, 14, 14], start=1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"


def write_workbook(df, out_path, bank_name=""):
    wb = Workbook()
    wb.remove(wb.active)
    title_suffix = f" — {bank_name.upper()}" if bank_name else ""

    dep_df = df[df["deposit"] > 0].copy()
    wd_df = df[df["withdrawal"] > 0].copy()

    ws_dep = wb.create_sheet("Deposits")
    _write_sheet(ws_dep, f"Deposits — grouped by counterparty{title_suffix}", dep_df, "deposit")

    ws_wd = wb.create_sheet("Withdrawals")
    _write_sheet(ws_wd, f"Withdrawals — grouped by counterparty{title_suffix}", wd_df, "withdrawal")

    wb.save(out_path)


# ============================================================
# CLI
# ============================================================

def main():
    if len(sys.argv) < 2:
        print("Usage: python sort_statement.py <statement_file> [output.xlsx]")
        sys.exit(1)
    inp = Path(sys.argv[1]).expanduser()
    out = Path(sys.argv[2]) if len(sys.argv) > 2 else inp.with_name(inp.stem + "_sorted.xlsx")

    prof = detect_profile(inp)
    print(f"Detected bank profile: {prof.name}")

    df = prof.parse(inp)
    print(f"Parsed {len(df)} rows.")
    extracted = df["counterparty"].notna().sum()
    print(f"Counterparty extracted: {extracted}/{len(df)} ({100*extracted/max(len(df),1):.0f}%)")

    write_workbook(df, out, bank_name=prof.name)
    print(f"Wrote: {out}")


if __name__ == "__main__":
    main()
