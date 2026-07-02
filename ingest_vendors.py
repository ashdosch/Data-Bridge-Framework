#!/usr/bin/env python3
"""
Vendor Excel ingestion -> canonical rows -> SQL Server upserts (project tables only).
 - Can be extended to include other formats for ingestion -

Windows-safe behaviors:
- Reads sheets via a single Excel handle: with pd.ExcelFile(...) as xl + xl.parse(...)
- Copy-then-delete for archiving/error moves, with retries/backoff
- Row-level error logging to CSV (does not kill the run)
- Conversion errors are coerced to NULL and logged (does not kill the run)
- File-level fatal errors move the file to --error-dir (copy-then-delete)

Schema: etl (Dynamic depending on below changes)

Deps:
  pip install pandas openpyxl pyodbc pyyaml
"""

import argparse
import csv
import gc
import json
import math
import re
import shutil
import time
from datetime import datetime
from decimal import Decimal, InvalidOperation
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
import pyodbc
import yaml

SCHEMA = "etl"


# ============================================================
# Conversions (NaN/blank safe) + conversion application
# ============================================================

_NULL_STRINGS = {"", "nan", "none", "null", "n/a", "na", "-", "—"}


def _is_blank(v: Any) -> bool:
    if v is None:
        return True
    try:
        if isinstance(v, float) and math.isnan(v):
            return True
    except Exception:
        pass
    s = str(v).strip()
    return s == "" or s.lower() in _NULL_STRINGS


def to_decimal(v: Any) -> Optional[float]:
    """
    Best-effort decimal:
      - blanks/NaN -> None
      - "$1,234.50" -> 1234.5
      - "120 - 277" or "2700-3000-..." -> first number
      - invalid -> None
    """
    if _is_blank(v):
        return None
    s = str(v).strip()
    s = s.replace("$", "").replace(",", "")

    # choose first number in ranges like "120 - 277" or "2700-3000-..."
    if re.search(r"\d+\s*-\s*\d+", s):
        s = re.split(r"\s*-\s*", s)[0]

    try:
        return float(Decimal(str(s).strip()))
    except (InvalidOperation, ValueError):
        return None


def to_money(v: Any) -> Optional[float]:
    return to_decimal(v)


def to_int(v: Any) -> Optional[int]:
    """
    Best-effort integer:
      - blanks/NaN -> None
      - "22W" -> 22
      - "2700-3000-..." -> 2700
      - invalid -> None
    """
    if _is_blank(v):
        return None
    s = str(v).strip()
    m = re.search(r"-?\d+", s)
    if not m:
        return None
    try:
        return int(m.group(0))
    except ValueError:
        return None


def to_bool(v: Any) -> Optional[bool]:
    """
    Best-effort boolean:
      - blanks/NaN -> None
      - Yes/No, Y/N, True/False, 1/0, Included/Not Included
      - unknown -> None
    """
    if _is_blank(v):
        return None
    s = str(v).strip().lower()
    if s in {"y", "yes", "true", "1", "included", "include", "t"}:
        return True
    if s in {"n", "no", "false", "0", "not included", "exclude", "f"}:
        return False
    return None


def apply_conversions_inplace(
    rec: Dict[str, Any],
    cfg: Dict[str, Any],
    errors: Optional[List[Dict[str, Any]]] = None,
    context: Optional[Dict[str, Any]] = None,
) -> None:
    """
    Applies conversions declared in cfg['conversions'] to the record dict in-place.
    - If conversion fails: sets field to None (keeps pipeline moving)
    - Optionally appends a structured error into `errors`
    """
    conv = cfg.get("conversions") or {}
    context = context or {}

    def _safe(field: str, fn):
        if field not in rec:
            return
        try:
            rec[field] = fn(rec[field])
        except Exception as e:
            bad_val = rec.get(field)
            rec[field] = None
            if errors is not None:
                errors.append(
                    {
                        "type": "conversion",
                        "field": field,
                        "value": str(bad_val),
                        "error": f"{type(e).__name__}: {e}",
                        **context,
                    }
                )

    for f in (conv.get("money") or []):
        _safe(f, to_money)
    for f in (conv.get("decimals") or []):
        _safe(f, to_decimal)
    for f in (conv.get("int") or []):
        _safe(f, to_int)
    for f in (conv.get("bool") or []):
        _safe(f, to_bool)


# ============================================================
# Normalization helpers
# ============================================================

def norm_col(c: Any) -> str:
    if c is None or (isinstance(c, float) and pd.isna(c)):
        return ""
    s = str(c).strip().lower()
    s = re.sub(r"\s+", " ", s)
    s = re.sub(r"[^\w\s]", "", s)  # drop punctuation
    s = s.replace(" ", "_")
    s = re.sub(r"_+", "_", s).strip("_")
    return s


def first_non_null(row: pd.Series, candidates) -> Any:
    """
    candidates can be: list[str] | str | None
    """
    if candidates is None:
        return None
    if isinstance(candidates, str):
        candidates = [candidates]
    if not isinstance(candidates, (list, tuple)):
        return None

    for c in candidates:
        if not c:
            continue
        if c in row.index:
            v = row[c]
            if v is None:
                continue
            if isinstance(v, float) and pd.isna(v):
                continue
            if isinstance(v, str) and v.strip() == "":
                continue
            return v
    return None


# ============================================================
# FS helpers (copy-then-delete)
# ============================================================

def ensure_dir(p: Path):
    p.mkdir(parents=True, exist_ok=True)


def safe_copy_delete(src: Path, dest_dir: Path, attempts: int = 12, sleep_s: float = 0.5) -> Path:
    """
    Windows-safe archive: copy to destination then delete source.
    Retries on PermissionError (AV/indexer lock) with backoff.
    """
    ensure_dir(dest_dir)
    dest = dest_dir / src.name
    if dest.exists():
        dest = dest_dir / f"{src.stem}_{datetime.now().strftime('%Y%m%d_%H%M%S')}{src.suffix}"

    last_err: Optional[Exception] = None
    for _ in range(attempts):
        try:
            shutil.copy2(src, dest)
            src.unlink()
            return dest
        except PermissionError as e:
            last_err = e
            time.sleep(sleep_s)
        except OSError as e:
            last_err = e
            time.sleep(sleep_s)

    raise last_err if last_err else RuntimeError("Failed to copy/delete file for unknown reasons.")


# ============================================================
# Mapping loading + classification
# ============================================================

def load_vendor_mappings(mappings_path: Path) -> List[Dict[str, Any]]:
    mappings: List[Dict[str, Any]] = []
    if mappings_path.is_dir():
        for p in sorted(mappings_path.glob("*.yaml")):
            with open(p, "r", encoding="utf-8") as f:
                obj = yaml.safe_load(f)
            if isinstance(obj, dict) and "vendors" in obj:
                mappings.extend(obj["vendors"])
            elif isinstance(obj, dict) and "vendor_code" in obj:
                mappings.append(obj)
    else:
        with open(mappings_path, "r", encoding="utf-8") as f:
            obj = yaml.safe_load(f)
        if isinstance(obj, dict) and "vendors" in obj:
            mappings.extend(obj["vendors"])
        elif isinstance(obj, dict) and "vendor_code" in obj:
            mappings.append(obj)
    return mappings


def pick_vendor_mapping(file_name: str, mappings: List[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
    lower = file_name.lower()
    for m in mappings:
        for pat in (m.get("file_match") or []):
            if str(pat).lower() in lower:
                return m
    return None


def classify_category(row: pd.Series, cfg: Dict[str, Any]) -> str:
    """
    Layered category classification.
    """
    import re as _re

    def _as_text(v: Any) -> str:
        if v is None:
            return ""
        try:
            if pd.isna(v):
                return ""
        except Exception:
            pass
        return str(v).strip()

    def _norm_text(v: Any) -> str:
        return _as_text(v).strip().lower()

    def _any_field_present(tokens: List[str]) -> bool:
        tok_l = [t.strip().lower() for t in (tokens or []) if str(t).strip() != ""]
        if not tok_l:
            return False
        for col in row.index:
            col_l = str(col).lower()
            val = _as_text(row.get(col))
            if val == "":
                continue
            for t in tok_l:
                if col_l == t or t in col_l:
                    return True
        return False

    def _column_value_match(columns: List[str], mappings: Dict[str, List[str]]) -> str:
        if not columns or not mappings:
            return ""
        parts = []
        for c in columns:
            if c in row.index:
                tv = _norm_text(row.get(c))
                if tv:
                    parts.append(tv)
        if not parts:
            return ""
        joined = " | ".join(parts)

        for cat_type, vals in mappings.items():
            for v in vals or []:
                vv = str(v).strip().lower()
                if vv and joined == vv:
                    return cat_type
        for cat_type, vals in mappings.items():
            for v in vals or []:
                vv = str(v).strip().lower()
                if vv and vv in joined:
                    return cat_type
        return ""

    def _regex_scan(columns: List[str], rules: List[Dict[str, Any]]) -> str:
        if not columns or not rules:
            return ""
        parts = []
        for c in columns:
            if c in row.index:
                tv = _as_text(row.get(c))
                if tv:
                    parts.append(tv)
        if not parts:
            return ""
        hay = "\n".join(parts)
        for rule in rules:
            cat_type = rule.get("category_type")
            for pat in (rule.get("any_regex") or []):
                try:
                    if _re.search(pat, hay):
                        return cat_type
                except _re.error:
                    continue
        return ""

    cat = cfg.get("category") or {}
    strat = (cat.get("strategy") or "default").strip().lower()
    default = cat.get("default", "lighting")

    if strat == "default":
        return default

    if strat == "rules":
        for rule in (cat.get("rules") or []):
            if "default" in rule:
                return rule["default"]
            if _any_field_present(rule.get("if_any_field_present") or []):
                return rule.get("category_type", default)
        return default

    if strat == "layered":
        cv = cat.get("column_value") or {}
        hit = _column_value_match(cv.get("columns") or [], cv.get("mappings") or {})
        if hit:
            return hit

        rs = cat.get("regex_scan") or {}
        hit = _regex_scan(rs.get("columns") or [], rs.get("rules") or [])
        if hit:
            return hit

        fp = cat.get("feature_presence") or {}
        for rule in (fp.get("rules") or []):
            if _any_field_present(rule.get("if_any_field_present") or []):
                return rule.get("category_type", default)

        return default

    return default


# ============================================================
# Record building
# ============================================================

def build_records(row: pd.Series, cfg: Dict[str, Any]) -> Tuple[Dict[str, Any], Dict[str, Any], str]:
    vendor_code = cfg["vendor_code"]
    cm = cfg.get("column_map") or {}

    core_map = cm.get("core") or {}
    lighting_map = cm.get("lighting") or {}
    fans_map = cm.get("fans") or {}
    bulbs_map = cm.get("light_bulbs") or {}

    category_type = classify_category(row, cfg)

    core: Dict[str, Any] = {"vendor_code": vendor_code, "category_type": category_type}
    for canon, candidates in core_map.items():
        core[canon] = first_non_null(row, candidates)

    core.setdefault("currency_code", (cfg.get("defaults") or {}).get("currency_code", "USD"))

    cat: Dict[str, Any] = {"vendor_code": vendor_code}
    if category_type == "lighting":
        for canon, candidates in lighting_map.items():
            cat[canon] = first_non_null(row, candidates)
    elif category_type == "fans":
        for canon, candidates in fans_map.items():
            cat[canon] = first_non_null(row, candidates)
    elif category_type == "light_bulbs":
        for canon, candidates in bulbs_map.items():
            cat[canon] = first_non_null(row, candidates)

    return core, cat, category_type


# ============================================================
# SQL upsert (schema etl)
# ============================================================

CORE_MERGE_SQL = f"""
MERGE {SCHEMA}.product_core AS tgt
USING (SELECT ? AS vendor_code, ? AS inventory_partno) AS src
  ON tgt.vendor_code = src.vendor_code AND tgt.inventory_partno = src.inventory_partno
WHEN MATCHED THEN UPDATE SET
  category_type = ?,
  upc = ?,
  wholesale_price = ?,
  imap_price = ?,
  is_imap = ?,
  width_in = ?,
  height_in = ?,
  length_in = ?,
  weight_lb = ?,
  shipping_weight_lb = ?,
  carton_height_in = ?,
  carton_length_in = ?,
  carton_width_in = ?,
  bulb_quantity = ?,
  marketing_description = ?,
  finish = ?,
  subcategory = ?,
  shipping_type = ?,
  collection = ?,
  location = ?,
  currency_code = COALESCE(?, currency_code),
  source_file_name = ?,
  source_sheet_name = ?,
  source_row_number = ?,
  run_id = ?
WHEN NOT MATCHED THEN
  INSERT (vendor_code, inventory_partno, category_type, upc, wholesale_price, imap_price, is_imap,
          width_in, height_in, length_in, weight_lb,
          shipping_weight_lb, carton_height_in, carton_length_in, carton_width_in,
          bulb_quantity, marketing_description, finish, subcategory, shipping_type, collection, location,
          currency_code, source_file_name, source_sheet_name, source_row_number, run_id)
  VALUES (?, ?, ?, ?, ?, ?, ?,
          ?, ?, ?, ?,
          ?, ?, ?, ?,
          ?, ?, ?, ?, ?, ?, ?,
          COALESCE(?, 'USD'), ?, ?, ?, ?);
"""

SELECT_PRODUCT_ID_SQL = f"""
SELECT product_id
FROM {SCHEMA}.product_core
WHERE vendor_code = ? AND inventory_partno = ?;
"""

LIGHTING_MERGE_SQL = f"""
MERGE {SCHEMA}.product_lighting AS tgt
USING (SELECT ? AS product_id) AS src
  ON tgt.product_id = src.product_id
WHEN MATCHED THEN UPDATE SET
  type1=?, type2=?, subtype1=?, room_used=?,
  certification_type=?, ada=?, energy_star=?,
  material=?, style=?, shade=?,
  voltage=?, bulb_socket=?, bulb_included=?, bulb_watts=?, bulb_shape=?,
  kelvin=?, cri=?, bulb_dimmable=?, lumens=?,
  canopy_height_in=?, canopy_length_in=?, canopy_width_in=?, canopy_depth_in=?,
  minimum_height_in=?, maximum_height_in=?
WHEN NOT MATCHED THEN
  INSERT (product_id, type1, type2, subtype1, room_used,
          certification_type, ada, energy_star,
          material, style, shade,
          voltage, bulb_socket, bulb_included, bulb_watts, bulb_shape,
          kelvin, cri, bulb_dimmable, lumens,
          canopy_height_in, canopy_length_in, canopy_width_in, canopy_depth_in,
          minimum_height_in, maximum_height_in)
  VALUES (?, ?, ?, ?, ?,
          ?, ?, ?,
          ?, ?, ?,
          ?, ?, ?, ?, ?,
          ?, ?, ?, ?,
          ?, ?, ?, ?,
          ?, ?);
"""

FANS_MERGE_SQL = f"""
MERGE {SCHEMA}.product_fans AS tgt
USING (SELECT ? AS product_id) AS src
  ON tgt.product_id = src.product_id
WHEN MATCHED THEN UPDATE SET
  type1=?, type2=?, subtype1=?, room_used=?, certification_type=?,
  blade_span_in=?, blade_count=?, blade_color=?, blade_pitch=?,
  cfm_average=?, cfm_high=?, cfm_average_watts=?, cfm_high_watts=?,
  max_angle=?, low_rpm=?, high_rpm=?, motor_watts=?, motor_type=?,
  is_flush_mount=?, warranty=?, downrod=?, control_type=?,
  light_kit_included=?, light_kit_watts=?, light_kit_kelvin=?, light_kit_lumens=?,
  diffuser_type=?, style=?
WHEN NOT MATCHED THEN
  INSERT (product_id, type1, type2, subtype1, room_used, certification_type,
          blade_span_in, blade_count, blade_color, blade_pitch,
          cfm_average, cfm_high, cfm_average_watts, cfm_high_watts,
          max_angle, low_rpm, high_rpm, motor_watts, motor_type,
          is_flush_mount, warranty, downrod, control_type,
          light_kit_included, light_kit_watts, light_kit_kelvin, light_kit_lumens,
          diffuser_type, style)
  VALUES (?, ?, ?, ?, ?, ?,
          ?, ?, ?, ?,
          ?, ?, ?, ?,
          ?, ?, ?, ?, ?,
          ?, ?, ?, ?,
          ?, ?, ?, ?,
          ?, ?);
"""

BULBS_MERGE_SQL = f"""
MERGE {SCHEMA}.product_light_bulbs AS tgt
USING (SELECT ? AS product_id) AS src
  ON tgt.product_id = src.product_id
WHEN MATCHED THEN UPDATE SET diameter_in=?
WHEN NOT MATCHED THEN INSERT (product_id, diameter_in) VALUES (?, ?);
"""


def upsert_core(cur, core: Dict[str, Any], lineage: Dict[str, Any]) -> int:
    vendor_code = core.get("vendor_code")
    partno = core.get("inventory_partno")
    if not vendor_code or not partno:
        raise ValueError("Missing vendor_code or inventory_partno (required).")

    params = (
        # match
        vendor_code,
        partno,
        # update
        core.get("category_type"),
        core.get("upc"),
        core.get("wholesale_price"),
        core.get("imap_price"),
        core.get("is_imap"),
        core.get("width_in"),
        core.get("height_in"),
        core.get("length_in"),
        core.get("weight_lb"),
        core.get("shipping_weight_lb"),
        core.get("carton_height_in"),
        core.get("carton_length_in"),
        core.get("carton_width_in"),
        core.get("bulb_quantity"),
        core.get("marketing_description"),
        core.get("finish"),
        core.get("subcategory"),
        core.get("shipping_type"),
        core.get("collection"),
        core.get("location"),
        core.get("currency_code"),
        lineage.get("source_file_name"),
        lineage.get("source_sheet_name"),
        lineage.get("source_row_number"),
        lineage.get("run_id"),
        # insert
        vendor_code,
        partno,
        core.get("category_type"),
        core.get("upc"),
        core.get("wholesale_price"),
        core.get("imap_price"),
        core.get("is_imap"),
        core.get("width_in"),
        core.get("height_in"),
        core.get("length_in"),
        core.get("weight_lb"),
        core.get("shipping_weight_lb"),
        core.get("carton_height_in"),
        core.get("carton_length_in"),
        core.get("carton_width_in"),
        core.get("bulb_quantity"),
        core.get("marketing_description"),
        core.get("finish"),
        core.get("subcategory"),
        core.get("shipping_type"),
        core.get("collection"),
        core.get("location"),
        core.get("currency_code"),
        lineage.get("source_file_name"),
        lineage.get("source_sheet_name"),
        lineage.get("source_row_number"),
        lineage.get("run_id"),
    )

    cur.execute(CORE_MERGE_SQL, params)
    cur.execute(SELECT_PRODUCT_ID_SQL, (vendor_code, partno))
    r = cur.fetchone()
    if not r:
        raise RuntimeError("product_id lookup failed after upsert.")
    return int(r[0])


def upsert_lighting(cur, product_id: int, cat: Dict[str, Any]):
    p = (
        # match
        product_id,
        # update
        cat.get("type1"),
        cat.get("type2"),
        cat.get("subtype1"),
        cat.get("room_used"),
        cat.get("certification_type"),
        cat.get("ada"),
        cat.get("energy_star"),
        cat.get("material"),
        cat.get("style"),
        cat.get("shade"),
        cat.get("voltage"),
        cat.get("bulb_socket"),
        cat.get("bulb_included"),
        cat.get("bulb_watts"),
        cat.get("bulb_shape"),
        cat.get("kelvin"),
        cat.get("cri"),
        cat.get("bulb_dimmable"),
        cat.get("lumens"),
        cat.get("canopy_height_in"),
        cat.get("canopy_length_in"),
        cat.get("canopy_width_in"),
        cat.get("canopy_depth_in"),
        cat.get("minimum_height_in"),
        cat.get("maximum_height_in"),
        # insert dup
        product_id,
        cat.get("type1"),
        cat.get("type2"),
        cat.get("subtype1"),
        cat.get("room_used"),
        cat.get("certification_type"),
        cat.get("ada"),
        cat.get("energy_star"),
        cat.get("material"),
        cat.get("style"),
        cat.get("shade"),
        cat.get("voltage"),
        cat.get("bulb_socket"),
        cat.get("bulb_included"),
        cat.get("bulb_watts"),
        cat.get("bulb_shape"),
        cat.get("kelvin"),
        cat.get("cri"),
        cat.get("bulb_dimmable"),
        cat.get("lumens"),
        cat.get("canopy_height_in"),
        cat.get("canopy_length_in"),
        cat.get("canopy_width_in"),
        cat.get("canopy_depth_in"),
        cat.get("minimum_height_in"),
        cat.get("maximum_height_in"),
    )
    cur.execute(LIGHTING_MERGE_SQL, p)


def upsert_fans(cur, product_id: int, cat: Dict[str, Any]):
    p = (
        product_id,
        cat.get("type1"),
        cat.get("type2"),
        cat.get("subtype1"),
        cat.get("room_used"),
        cat.get("certification_type"),
        cat.get("blade_span_in"),
        cat.get("blade_count"),
        cat.get("blade_color"),
        cat.get("blade_pitch"),
        cat.get("cfm_average"),
        cat.get("cfm_high"),
        cat.get("cfm_average_watts"),
        cat.get("cfm_high_watts"),
        cat.get("max_angle"),
        cat.get("low_rpm"),
        cat.get("high_rpm"),
        cat.get("motor_watts"),
        cat.get("motor_type"),
        cat.get("is_flush_mount"),
        cat.get("warranty"),
        cat.get("downrod"),
        cat.get("control_type"),
        cat.get("light_kit_included"),
        cat.get("light_kit_watts"),
        cat.get("light_kit_kelvin"),
        cat.get("light_kit_lumens"),
        cat.get("diffuser_type"),
        cat.get("style"),
        # insert dup
        product_id,
        cat.get("type1"),
        cat.get("type2"),
        cat.get("subtype1"),
        cat.get("room_used"),
        cat.get("certification_type"),
        cat.get("blade_span_in"),
        cat.get("blade_count"),
        cat.get("blade_color"),
        cat.get("blade_pitch"),
        cat.get("cfm_average"),
        cat.get("cfm_high"),
        cat.get("cfm_average_watts"),
        cat.get("cfm_high_watts"),
        cat.get("max_angle"),
        cat.get("low_rpm"),
        cat.get("high_rpm"),
        cat.get("motor_watts"),
        cat.get("motor_type"),
        cat.get("is_flush_mount"),
        cat.get("warranty"),
        cat.get("downrod"),
        cat.get("control_type"),
        cat.get("light_kit_included"),
        cat.get("light_kit_watts"),
        cat.get("light_kit_kelvin"),
        cat.get("light_kit_lumens"),
        cat.get("diffuser_type"),
        cat.get("style"),
    )
    cur.execute(FANS_MERGE_SQL, p)


def upsert_bulbs(cur, product_id: int, cat: Dict[str, Any]):
    cur.execute(BULBS_MERGE_SQL, (product_id, cat.get("diameter_in"), product_id, cat.get("diameter_in")))


def upsert_presence(cur, table_name: str, product_id: int):
    cur.execute(
        f"IF NOT EXISTS (SELECT 1 FROM {table_name} WHERE product_id = ?) "
        f"INSERT INTO {table_name}(product_id) VALUES (?);",
        (product_id, product_id),
    )


# ============================================================
# Ingest file
# ============================================================

def ingest_file(
    conn_str: str,
    excel_path: Path,
    cfg: Dict[str, Any],
    run_id: Optional[int],
    error_rows: List[Dict[str, Any]],
    conversion_errors: List[Dict[str, Any]],
) -> Dict[str, Any]:
    res = {
        "file": excel_path.name,
        "vendor_code": cfg["vendor_code"],
        "rows_total": 0,
        "rows_upserted": 0,
        "rows_skipped_missing_key": 0,
        "by_category": {},
    }

    with pd.ExcelFile(excel_path) as xl:
        with pyodbc.connect(conn_str) as conn:
            conn.autocommit = False
            cur = conn.cursor()

            for sheet_cfg in (cfg.get("sheets") or []):
                sheet_name = sheet_cfg["sheet_name"]
                if sheet_name not in xl.sheet_names:
                    continue

                header_row = int(sheet_cfg.get("header_row", 0))
                df = xl.parse(sheet_name=sheet_name, header=header_row, dtype=object)
                if df is None or df.empty:
                    continue

                df.columns = [norm_col(c) for c in df.columns]

                # optional quick debug - comment out once stable
                # print("HEADER ROW:", header_row)
                # print("COLUMNS:", df.columns.tolist()[:40])
                # print("FIRST DATA ROW:", df.iloc[0].to_dict())

                for idx, row in df.iterrows():
                    res["rows_total"] += 1
                    core: Optional[Dict[str, Any]] = None
                    try:
                        core, cat, category_type = build_records(row, cfg)

                        # conversions happen HERE (correct scope)
                        context = {
                            "vendor": cfg.get("vendor_code"),
                            "file": excel_path.name,
                            "sheet": sheet_name,
                            "row_index": int(idx),
                        }
                        apply_conversions_inplace(core, cfg, errors=conversion_errors, context=context)
                        apply_conversions_inplace(cat, cfg, errors=conversion_errors, context=context)

                        partno = core.get("inventory_partno")
                        if partno is None or str(partno).strip() == "":
                            res["rows_skipped_missing_key"] += 1

                            # Show first few examples so we can fix YAML quickly
                            if res["rows_skipped_missing_key"] <= 10:
                                available_cols = list(row.index)[:40]
                                print(f"[MISSING KEY] file={excel_path.name} sheet={sheet_name} row={int(idx)}")
                                print("  cols(first40)=", available_cols)
                                # print a few likely columns if present
                                for probe in ["product_number", "config_sku", "base_model", "sku", "item", "model", "part_number", "catalog_no"]:
                                    if probe in row.index:
                                        print(f"  {probe}={row.get(probe)}")
                            continue

                        lineage = {
                            "source_file_name": excel_path.name,
                            "source_sheet_name": sheet_name,
                            "source_row_number": int(idx),
                            "run_id": run_id,
                        }

                        product_id = upsert_core(cur, core, lineage)

                        if category_type == "lighting":
                            upsert_lighting(cur, product_id, cat)
                        elif category_type == "fans":
                            upsert_fans(cur, product_id, cat)
                        elif category_type == "light_bulbs":
                            upsert_bulbs(cur, product_id, cat)
                        elif category_type == "accessories":
                            upsert_presence(cur, f"{SCHEMA}.product_accessories", product_id)
                        elif category_type == "home_decor":
                            upsert_presence(cur, f"{SCHEMA}.product_home_decor", product_id)

                        res["rows_upserted"] += 1
                        res["by_category"][category_type] = res["by_category"].get(category_type, 0) + 1

                    except Exception as e:
                        error_rows.append(
                            {
                                "timestamp": datetime.now().isoformat(),
                                "file": excel_path.name,
                                "sheet": sheet_name,
                                "row_index": int(idx),
                                "vendor_code": cfg.get("vendor_code"),
                                "inventory_partno": str(core.get("inventory_partno", "")) if isinstance(core, dict) else "",
                                "category_type": classify_category(row, cfg),
                                "error_type": type(e).__name__,
                                "error_message": str(e)[:3500],
                            }
                        )

                del df
                gc.collect()

            conn.commit()

    gc.collect()
    return res


# ============================================================
# Main
# ============================================================

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--input", required=True, help="Directory containing vendor Excel files")
    ap.add_argument("--mappings", required=True, help="Directory of YAML mappings or a combined YAML file")
    ap.add_argument("--conn", required=True, help="ODBC connection string for SQL Server")
    ap.add_argument("--run-id", type=int, default=None, help="Optional run_id to stamp into product_core")
    ap.add_argument("--report-out", default="ingest_report.json", help="JSON report filename (under --log-dir)")
    ap.add_argument("--log-dir", required=True, help="Directory for logs (JSON report + error CSV)")
    ap.add_argument("--archive-dir", required=True, help="Directory to copy successful files into")
    ap.add_argument("--error-dir", required=True, help="Directory to copy failed files into")
    args = ap.parse_args()

    in_dir = Path(args.input)
    mappings = load_vendor_mappings(Path(args.mappings))

    log_dir = Path(args.log_dir)
    archive_dir = Path(args.archive_dir)
    error_dir = Path(args.error_dir)
    ensure_dir(log_dir)
    ensure_dir(archive_dir)
    ensure_dir(error_dir)

    error_rows: List[Dict[str, Any]] = []
    conversion_errors: List[Dict[str, Any]] = []

    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    error_csv_path = log_dir / f"ingest_errors_{ts}.csv"
    conv_csv_path = log_dir / f"conversion_errors_{ts}.csv"
    report_json_path = log_dir / args.report_out

    results: Dict[str, Any] = {"input_dir": str(in_dir.resolve()), "runs": []}

    def _is_under(child: Path, parent: Path) -> bool:
        try:
            child.resolve().relative_to(parent.resolve())
            return True
        except Exception:
            return False

    for p in sorted(in_dir.rglob("*")):
        if not p.is_file():
            continue
        if p.name.startswith("~$"):
            continue
        if p.suffix.lower() not in {".xlsx", ".xlsm", ".xls"}:
            continue
        if _is_under(p, archive_dir) or _is_under(p, error_dir) or _is_under(p, log_dir):
            continue

        cfg = pick_vendor_mapping(p.name, mappings)
        if not cfg:
            results["runs"].append({"file": p.name, "skipped": True, "reason": "No vendor mapping matched file_name"})
            continue

        try:
            run_res = ingest_file(args.conn, p, cfg, args.run_id, error_rows, conversion_errors)
            results["runs"].append(run_res)
            print(f"{p.name}: upserted {run_res['rows_upserted']} / {run_res['rows_total']} (vendor={run_res['vendor_code']})")

            gc.collect()
            dest = safe_copy_delete(p, archive_dir)
            run_res["archived_to"] = str(dest)

        except Exception as e:
            results["runs"].append(
                {
                    "file": p.name,
                    "vendor_code": cfg.get("vendor_code"),
                    "skipped": True,
                    "reason": f"Fatal error: {type(e).__name__}: {e}",
                }
            )
            print(f"{p.name}: FAILED ({type(e).__name__}: {e})")

            try:
                gc.collect()
                safe_copy_delete(p, error_dir)
            except Exception as move_err:
                results["runs"][-1]["move_error"] = f"{type(move_err).__name__}: {move_err}"

    with open(report_json_path, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2)

    if error_rows:
        with open(error_csv_path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=list(error_rows[0].keys()))
            w.writeheader()
            w.writerows(error_rows)

    if conversion_errors:
        keys = sorted({k for e in conversion_errors for k in e.keys()})
        with open(conv_csv_path, "w", newline="", encoding="utf-8") as f:
            w = csv.DictWriter(f, fieldnames=keys)
            w.writeheader()
            w.writerows(conversion_errors)

    print(f"Wrote JSON report to: {report_json_path.resolve()}")
    if error_rows:
        print(f"Wrote error CSV to: {error_csv_path.resolve()}")
    if conversion_errors:
        print(f"Wrote conversion CSV to: {conv_csv_path.resolve()}")


if __name__ == "__main__":
    main()
