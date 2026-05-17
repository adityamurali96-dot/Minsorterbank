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
    def parse(cls, path: Path) -> pd.DataFrame:
        raise NotImplementedError

    @classmethod
    def extract(cls, remarks: str) -> Optional[str]:
        return extract_generic(remarks)


def _load_raw(path: Path):
    """Return (raw_text_first_8k, raw_df_or_None). Used for detection.

    Reads the file as text and (if BIFF) as a dataframe. Doesn't fail on either side."""
    raw_text = ""
    try:
        with open(path, "rb") as f:
            raw_text = f.read(8192).decode("utf-8", errors="replace")
    except Exception:
        pass
    raw_df = None
    try:
        raw_df = pd.read_excel(path, sheet_name=0, engine="xlrd", header=None)
    except Exception:
        try:
            raw_df = pd.read_excel(path, sheet_name=0, engine="openpyxl", header=None)
        except Exception:
            pass
    return raw_text, raw_df


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
        if raw_df is None:
            return False
        # ICICI .xls has 'Transaction Remarks' header AND 'ICICI' somewhere up top
        hdr = _find_header_row(raw_df, "transaction remarks")
        if hdr is None:
            return False
        head_text = " ".join(str(v).lower() for v in raw_df.iloc[:hdr].values.flatten() if v == v)
        return "icici" in head_text or "icicibank" in head_text or hdr is not None  # ICICI is the default 'transaction remarks' format

    @classmethod
    def parse(cls, path):
        raw = pd.read_excel(path, sheet_name=0, engine="xlrd", header=None)
        hdr = _find_header_row(raw, "transaction remarks")
        df = raw.iloc[hdr + 1:].copy()
        df.columns = ["_pad", "sno", "value_date", "txn_date", "cheque",
                      "remarks", "withdrawal", "deposit", "balance"]
        df = df.dropna(subset=["remarks"]).reset_index(drop=True)
        df["withdrawal"] = pd.to_numeric(df["withdrawal"], errors="coerce").fillna(0)
        df["deposit"] = pd.to_numeric(df["deposit"], errors="coerce").fillna(0)
        df["balance"] = pd.to_numeric(df["balance"], errors="coerce")
        df = df[(df["withdrawal"] > 0) | (df["deposit"] > 0)].reset_index(drop=True)
        df["counterparty"] = df["remarks"].map(cls.extract)
        return df[["txn_date", "remarks", "withdrawal", "deposit", "balance", "counterparty"]]

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
        if raw_df is None:
            return False
        hdr = _find_header_row(raw_df, "particulars")
        if hdr is None:
            return False
        # Axis distinctive: 'PARTICULARS' header AND 'SOL' column AND 'IFSC' in metadata mentions UTIB
        head_text = " ".join(str(v) for v in raw_df.iloc[:hdr].values.flatten() if v == v)
        return "UTIB" in head_text or "AXIS" in head_text.upper()

    @classmethod
    def parse(cls, path):
        raw = pd.read_excel(path, sheet_name=0, engine="xlrd", header=None)
        hdr = _find_header_row(raw, "particulars")
        # Axis cols (from sample): SRL NO, Tran Date, CHQNO, PARTICULARS, DR, CR, BAL, SOL
        df = raw.iloc[hdr + 1:].copy()
        df.columns = ["sno", "txn_date", "cheque", "remarks", "withdrawal", "deposit", "balance", "sol"]
        df = df.dropna(subset=["remarks"]).reset_index(drop=True)
        df["withdrawal"] = pd.to_numeric(df["withdrawal"], errors="coerce").fillna(0)
        df["deposit"] = pd.to_numeric(df["deposit"], errors="coerce").fillna(0)
        df["balance"] = pd.to_numeric(df["balance"], errors="coerce")
        df = df[(df["withdrawal"] > 0) | (df["deposit"] > 0)].reset_index(drop=True)
        df["counterparty"] = df["remarks"].map(cls.extract)
        return df[["txn_date", "remarks", "withdrawal", "deposit", "balance", "counterparty"]]

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
        if raw_df is None:
            return False
        # HDFC: 'Narration' header AND 'HDFC' in metadata
        hdr = _find_header_row(raw_df, "narration")
        if hdr is None:
            return False
        head_text = " ".join(str(v) for v in raw_df.iloc[:hdr].values.flatten() if v == v)
        return "HDFC" in head_text.upper()

    @classmethod
    def parse(cls, path):
        raw = pd.read_excel(path, sheet_name=0, engine="xlrd", header=None)
        hdr = _find_header_row(raw, "narration")
        # HDFC cols: Date, Narration, Chq./Ref.No., Value Dt, Withdrawal Amt., Deposit Amt., Closing Balance, Remarks
        df = raw.iloc[hdr + 2:].copy()  # +2 because HDFC has a separator row of asterisks
        df.columns = ["txn_date", "remarks", "chq_ref", "value_dt",
                      "withdrawal", "deposit", "balance", "extra_remarks"]
        df = df.dropna(subset=["remarks"]).reset_index(drop=True)
        # Drop the row of asterisks and any header-style rows
        df = df[~df["remarks"].astype(str).str.startswith("*")].reset_index(drop=True)
        df["withdrawal"] = pd.to_numeric(df["withdrawal"], errors="coerce").fillna(0)
        df["deposit"] = pd.to_numeric(df["deposit"], errors="coerce").fillna(0)
        df["balance"] = pd.to_numeric(df["balance"], errors="coerce")
        df = df[(df["withdrawal"] > 0) | (df["deposit"] > 0)].reset_index(drop=True)
        df["counterparty"] = df["remarks"].map(cls.extract)
        return df[["txn_date", "remarks", "withdrawal", "deposit", "balance", "counterparty"]]

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
        # SBI: tab-separated text file with 'Account Name' in first line
        return "Account Name" in raw_text and "IFS (Indian Financial System)" in raw_text

    @classmethod
    def parse(cls, path):
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            text = f.read()
        # Find the header line containing 'Txn Date' and 'Description'
        lines = text.split("\n")
        hdr_idx = None
        for i, ln in enumerate(lines):
            low = ln.lower()
            if "txn date" in low and "description" in low and "debit" in low:
                hdr_idx = i
                break
        if hdr_idx is None:
            raise RuntimeError("SBI: couldn't locate transaction header line")

        # SBI cols: Txn Date, Value Date, Description, Ref No./Cheque No., Debit, Credit, Balance
        rows = []
        for ln in lines[hdr_idx + 1:]:
            if not ln.strip():
                continue
            fields = ln.split("\t")
            if len(fields) < 7:
                continue
            rows.append(fields[:7])

        df = pd.DataFrame(rows, columns=["txn_date", "value_date", "remarks",
                                          "chq_ref", "withdrawal", "deposit", "balance"])

        def _to_num(s):
            s = str(s).strip().strip('"').replace(",", "").strip()
            if not s or s == "nan":
                return 0.0
            try:
                return float(s)
            except ValueError:
                return 0.0

        df["withdrawal"] = df["withdrawal"].map(_to_num)
        df["deposit"] = df["deposit"].map(_to_num)
        df["balance"] = df["balance"].map(_to_num)
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
    def parse(cls, path):
        # Best-effort column detection for unknown formats.
        raw = None
        for engine in ("xlrd", "openpyxl"):
            try:
                raw = pd.read_excel(path, sheet_name=0, engine=engine, header=None)
                break
            except Exception:
                continue
        if raw is None:
            # Try as tab-separated text
            with open(path, "r", encoding="utf-8", errors="replace") as f:
                text = f.read()
            return _parse_generic_text(text)

        return _parse_generic_df(raw)

    @classmethod
    def extract(cls, remarks):
        n = extract_generic(remarks)
        return _apply_alias(n) if n else None


def _parse_generic_df(raw: pd.DataFrame) -> pd.DataFrame:
    # Find header row by looking for date- and amount-like words.
    DATE_TOKENS = ("date", "txn date", "tran date", "transaction date")
    AMT_DEBIT = ("withdrawal", "debit", "dr", "withdrawal amt")
    AMT_CREDIT = ("deposit", "credit", "cr", "deposit amt")
    REMARKS = ("remarks", "narration", "particulars", "description", "transaction remarks")

    hdr = None
    for i in range(min(50, len(raw))):
        row_lc = [str(v).lower().strip() for v in raw.iloc[i].tolist()]
        has_date = any(any(t in v for t in DATE_TOKENS) for v in row_lc)
        has_remarks = any(any(t in v for t in REMARKS) for v in row_lc)
        has_amt = any(any(t in v for t in AMT_DEBIT + AMT_CREDIT) for v in row_lc)
        if has_date and has_remarks and has_amt:
            hdr = i
            break
    if hdr is None:
        raise RuntimeError("Generic parser: couldn't find header row")

    header = [str(v).lower().strip() for v in raw.iloc[hdr].tolist()]

    def find_col(*needles):
        for idx, h in enumerate(header):
            for n in needles:
                if n in h:
                    return idx
        return None

    col_date = find_col(*DATE_TOKENS)
    col_rem = find_col(*REMARKS)
    col_wd = find_col(*AMT_DEBIT)
    col_dep = find_col(*AMT_CREDIT)
    col_bal = find_col("balance", "bal")

    df = raw.iloc[hdr + 1:].copy().reset_index(drop=True)

    def col(i):
        return df.iloc[:, i] if i is not None else pd.Series([None] * len(df))

    out = pd.DataFrame({
        "txn_date": col(col_date),
        "remarks": col(col_rem),
        "withdrawal": pd.to_numeric(col(col_wd), errors="coerce").fillna(0),
        "deposit": pd.to_numeric(col(col_dep), errors="coerce").fillna(0),
        "balance": pd.to_numeric(col(col_bal), errors="coerce"),
    })
    out = out.dropna(subset=["remarks"])
    out = out[(out["withdrawal"] > 0) | (out["deposit"] > 0)].reset_index(drop=True)
    out["counterparty"] = out["remarks"].map(lambda r: _apply_alias(extract_generic(r)) if extract_generic(r) else None)
    return out


def _parse_generic_text(text: str) -> pd.DataFrame:
    # Tab- or comma-separated; sniff which.
    sep = "\t" if text.count("\t") > text.count(",") else ","
    rows = list(csv.reader(io.StringIO(text), delimiter=sep))
    # Find header row
    hdr = None
    for i, row in enumerate(rows):
        joined = " | ".join(c.lower() for c in row)
        if any(d in joined for d in ("date",)) and any(r in joined for r in ("remarks", "narration", "particulars", "description")) and any(a in joined for a in ("debit", "credit", "withdrawal", "deposit")):
            hdr = i
            break
    if hdr is None:
        raise RuntimeError("Generic text parser: couldn't find header row")
    df = pd.DataFrame(rows[hdr + 1:], columns=[c.strip().lower() for c in rows[hdr]])

    def pick(*needles):
        for c in df.columns:
            for n in needles:
                if n in c:
                    return c
        return None

    c_date = pick("txn date", "tran date", "date")
    c_rem = pick("description", "narration", "particulars", "remarks")
    c_wd = pick("debit", "withdrawal", "dr")
    c_dep = pick("credit", "deposit", "cr")
    c_bal = pick("balance")

    def numify(s):
        s = str(s).strip().strip('"').replace(",", "")
        try:
            return float(s)
        except (ValueError, TypeError):
            return 0.0

    out = pd.DataFrame({
        "txn_date": df[c_date] if c_date else "",
        "remarks": df[c_rem] if c_rem else "",
        "withdrawal": df[c_wd].map(numify) if c_wd else 0,
        "deposit": df[c_dep].map(numify) if c_dep else 0,
        "balance": df[c_bal].map(numify) if c_bal else 0,
    })
    out = out.dropna(subset=["remarks"])
    out = out[(out["withdrawal"] > 0) | (out["deposit"] > 0)].reset_index(drop=True)
    out["counterparty"] = out["remarks"].map(lambda r: _apply_alias(extract_generic(r)) if extract_generic(r) else None)
    return out


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

    for b in named:
        sub = df[df["bucket"] == b].sort_values("txn_date")
        write_group(b, sub)
    if not other_df.empty:
        write_group("OTHER (one-offs)", other_df.sort_values("txn_date"))

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
