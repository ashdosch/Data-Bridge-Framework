import pandas as pd
import pyodbc
import re
import hashlib
from pathlib import Path
from datetime import datetime

# ================= CONFIG =================
INPUT_DIR = r"REDACTED"
CONN_STR = r"DRIVER={ODBC Driver 18 for SQL Server};SERVER=INSERTSERVERHERE;UID=USER;PWD=PASSWORD;DATABASE=pim;TrustServerCertificate=Yes"
SCHEMA = "SCHEMA"
TABLE_PREFIX = "REQUIREDPREFIX"
MAX_HEADER_SCAN_ROWS = 30

# SQL Server identifier max is 128. 120 leaves room for suffixes in the nomenclature.
MAX_COL_LEN = 120
# =========================================


def ensure_schema(cur):
    cur.execute(f"""
                
    IF NOT EXISTS (SELECT 1 FROM sys.schemas WHERE name = '{SCHEMA}')
    BEGIN
        EXEC('CREATE SCHEMA {SCHEMA}')
    END
    """)


def ensure_colmap_table(cur):
    """
    Optional helper table so you can trace original headers to stored columns.
    If you choose to use this, the below statement will create the table if not already created in the database.
    """
    cur.execute(f"""
    IF OBJECT_ID('{SCHEMA}.raw_excel_column_map', 'U') IS NULL
    BEGIN
        CREATE TABLE {SCHEMA}.raw_excel_column_map (
            map_id BIGINT IDENTITY(1,1) PRIMARY KEY,
            loaded_at DATETIME2 NOT NULL,
            source_file NVARCHAR(260) NOT NULL,
            source_sheet NVARCHAR(200) NOT NULL,
            original_header NVARCHAR(MAX) NULL,
            normalized_header NVARCHAR(512) NULL,
            stored_column_name NVARCHAR(128) NOT NULL
        );
    END
    """)


def normalize_header(s: str) -> str:
    s = str(s).strip().lower()
    s = re.sub(r"\s+", "_", s)
    s = re.sub(r"[^\w]", "", s)
    s = re.sub(r"_+", "_", s).strip("_")
    return s or "col"


def make_safe_identifier(norm: str, max_len: int = MAX_COL_LEN) -> str:
    """
    Truncate + hash to avoid >128 and avoid collisions.
    """
    norm = norm[:max_len]
    return norm


def shorten_and_dedupe_headers(original_headers):
    """
    Produce SQL-safe column names <= MAX_COL_LEN and unique.
    This also helps with consistency if you need to perform DML or DDL updates on the server itself.

    Returns: (safe_cols, mapping_rows)
      safe_cols: list[str] final unique col names
      mapping_rows: list[dict] with original/normalized/stored
    """
    safe_cols = []
    seen = {}

    mapping = []
    for h in original_headers:
        orig = "" if h is None else str(h)
        norm = normalize_header(orig)

        base = make_safe_identifier(norm)

        # If too long, append hash
        # (Also helps disambiguate if multiple long headers truncate similarly)
        hsh = hashlib.sha1(norm.encode("utf-8")).hexdigest()[:8]
        if len(norm) > MAX_COL_LEN:
            base = (base[: MAX_COL_LEN - 9] + "_" + hsh)

        # Ensure uniqueness with counters
        final = base
        if final in seen:
            seen[final] += 1
            suffix = f"_{seen[final]}"
            # ensure suffix fits
            final = final[: MAX_COL_LEN - len(suffix)] + suffix
        else:
            seen[final] = 1

        safe_cols.append(final)
        mapping.append({
            "original_header": orig,
            "normalized_header": norm,
            "stored_column_name": final
        })

    return safe_cols, mapping


def detect_header_row(raw_df: pd.DataFrame, max_scan: int = MAX_HEADER_SCAN_ROWS) -> int:
    best_row, best_score = 0, float("-inf")
    scan_n = min(max_scan, len(raw_df))

    for i in range(scan_n):
        row = raw_df.iloc[i]
        non_null = row.notna().sum()
        str_cells = sum(isinstance(x, str) and x.strip() != "" for x in row)
        num_cells = sum(isinstance(x, (int, float)) for x in row)

        score = (str_cells * 3) + non_null - (num_cells * 0.5)
        if score > best_score:
            best_score = score
            best_row = i

    return best_row


def safe_table_name(file_stem: str, sheet_name: str) -> str:
    t = f"{TABLE_PREFIX}_{file_stem}_{sheet_name}"
    t = re.sub(r"[^\w]", "_", t)
    t = re.sub(r"_+", "_", t).strip("_")
    return t[:120]


def to_str_or_none(x):
    try:
        if pd.isna(x):
            return None
    except Exception:
        pass

    if x is None:
        return None

    s = str(x).strip()
    if s == "":
        return None
    if s.lower() in {"nan", "none", "n/a", "na", "null"}:
        return None
    return s


def drop_and_create_table(cur, table_name: str, df_cols):
    col_defs = ",\n        ".join([f"[{c}] NVARCHAR(MAX) NULL" for c in df_cols])

    sql = f"""
    IF OBJECT_ID('{SCHEMA}.{table_name}', 'U') IS NOT NULL
        DROP TABLE {SCHEMA}.{table_name};

    CREATE TABLE {SCHEMA}.{table_name} (
        raw_id BIGINT IDENTITY(1,1) PRIMARY KEY,
        source_file NVARCHAR(260) NOT NULL,
        source_sheet NVARCHAR(200) NOT NULL,
        header_row BIGINT NOT NULL,
        source_row_number BIGINT NOT NULL,
        loaded_at DATETIME2 NOT NULL,
        {col_defs}
    );
    """
    cur.execute(sql)


def write_colmap(cur, file_name: str, sheet_name: str, mapping_rows, loaded_at: datetime):
    ensure_colmap_table(cur)
    rows = [
        (loaded_at, file_name, sheet_name, m["original_header"], m["normalized_header"], m["stored_column_name"])
        for m in mapping_rows
    ]
    cur.fast_executemany = True
    cur.executemany(
        f"""
        INSERT INTO {SCHEMA}.raw_excel_column_map
        (loaded_at, source_file, source_sheet, original_header, normalized_header, stored_column_name)
        VALUES (?, ?, ?, ?, ?, ?)
        """,
        rows
    )


def bulk_insert_rows(cur, table_name: str, df: pd.DataFrame, file_name: str, sheet_name: str, header_row: int):
    loaded_at = datetime.now()

    df_cols = list(df.columns)
    cols_sql = ", ".join(f"[{c}]" for c in df_cols)

    placeholders = ", ".join(["?"] * (len(df_cols) + 5))
    insert_sql = f"""
    INSERT INTO {SCHEMA}.{table_name}
    (source_file, source_sheet, header_row, source_row_number, loaded_at, {cols_sql})
    VALUES ({placeholders})
    """

    rows = []
    for i, row in enumerate(df.itertuples(index=False, name=None)):
        safe_cells = tuple(to_str_or_none(v) for v in row)
        rows.append((file_name, sheet_name, int(header_row), int(i), loaded_at, *safe_cells))

    qm = insert_sql.count("?")
    if rows and qm != len(rows[0]):
        raise RuntimeError(f"Placeholder mismatch: SQL has {qm} ?, row has {len(rows[0])}")

    try:
        cur.fast_executemany = True
        cur.executemany(insert_sql, rows)
        return len(rows), loaded_at
    except pyodbc.DataError:
        # fallback to locate offending row
        print("    Bulk insert DataError. Trying row-by-row to locate offending record...")
        cur.fast_executemany = False
        for idx, r in enumerate(rows):
            try:
                cur.execute(insert_sql, r)
            except Exception as inner:
                print(f"    FAILED at row #{idx} (source_row_number={r[3]}): {inner}")
                print("    First 20 params:", r[:20])
                raise
        raise


def process_file(conn, file_path: Path):
    print(f"\nProcessing: {file_path.name}")

    with pd.ExcelFile(file_path) as xl:
        sheet_names = xl.sheet_names

    for sheet_name in sheet_names:
        print(f"  Sheet: {sheet_name}")

        with pd.ExcelFile(file_path) as xl:
            raw = xl.parse(sheet_name=sheet_name, header=None, dtype=object)
            if raw is None or raw.empty:
                print("    Skipped (empty raw)")
                continue

            header_row = detect_header_row(raw)

            df = xl.parse(sheet_name=sheet_name, header=header_row, dtype=object)
            if df is None or df.empty:
                print("    Skipped (empty after header parse)")
                continue

        # drop empty rows/cols
        df = df.dropna(axis=0, how="all").dropna(axis=1, how="all")
        if df.empty:
            print("    Skipped (empty after cleanup)")
            continue

        # normalize, shorten, dedupe
        safe_cols, mapping_rows = shorten_and_dedupe_headers(df.columns)
        df.columns = safe_cols

        table_name = safe_table_name(file_path.stem, sheet_name)

        cur = conn.cursor()
        ensure_schema(cur)

        # create table + write mapping + insert rows
        drop_and_create_table(cur, table_name, df.columns)

        inserted, loaded_at = bulk_insert_rows(cur, table_name, df, file_path.name, sheet_name, header_row)

        # store column map (optional but recommended)
        write_colmap(cur, file_path.name, sheet_name, mapping_rows, loaded_at)

        conn.commit()
        print(f"    Imported {inserted:,} rows -> {SCHEMA}.{table_name} (header_row={header_row})")


def main():
    input_dir = Path(INPUT_DIR)
    if not input_dir.exists():
        raise FileNotFoundError(f"INPUT_DIR does not exist: {INPUT_DIR}")

    conn = pyodbc.connect(CONN_STR)

    files = sorted([p for p in input_dir.rglob("*") if p.is_file() and p.suffix.lower() in {".xlsx", ".xlsm", ".xlsb", ".xls"}])
    if not files:
        print("No Excel files found.")
        return

    for fp in files:
        if fp.name.startswith("~$"):
            continue
        process_file(conn, fp)

    conn.close()
    print("\n✅ Done.")


if __name__ == "__main__":
    main()
