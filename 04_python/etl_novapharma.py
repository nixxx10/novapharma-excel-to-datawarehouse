"""
NovaPharma Iberia - monthly sales ETL
=====================================
Loads one ERP sales extract (CSV) into the MySQL warehouse.

    python etl_novapharma.py ERP_Sales_202609.csv

Same rules as the prototype notebook (01_pipeline_prototype.ipynb), without the exploration:
extract -> transform (8 cleaning rules) -> load (idempotent) -> log.
"""
import sys, os, shutil, configparser
from pathlib import Path
from getpass import getpass

import pandas as pd
from sqlalchemy import create_engine, text

HERE = Path(__file__).resolve().parent

# ----------------------------------------------------------------------------- 0. config
def load_config():
    cfg = configparser.ConfigParser()
    cfg.read(HERE / "config.ini")
    pwd = os.environ.get("NOVAPHARMA_DB_PWD") or getpass("MySQL password: ")
    m = cfg["mysql"]
    url = f"mysql+pymysql://{m['user']}:{pwd}@{m['host']}:{m['port']}/{m['database']}"
    return cfg, create_engine(url)


# ----------------------------------------------------------------------------- 1. extract
def extract_sales(path: Path) -> pd.DataFrame:
    """Read the ERP extract (CSV) exactly as it arrives."""
    df = pd.read_csv(path, parse_dates=["Invoice_Date"])
    expected = ["Transaction_ID", "Invoice_Date", "SKU_Code", "Customer_Code",
                "Units", "Gross_Sales_EUR", "Discount_EUR", "Rebate_EUR"]
    missing = [c for c in expected if c not in df.columns]
    if missing:
        raise ValueError(f"Header mismatch, missing columns: {missing}")
    return df


def fetch_dimensions(engine):
    """The masters the FK rules are checked against."""
    with engine.connect() as con:
        skus  = pd.read_sql("SELECT sku_code FROM dim_product", con)["sku_code"]
        custs = pd.read_sql("SELECT customer_code FROM dim_customer", con)["customer_code"]
        closed = con.execute(text("SELECT MAX(month_key) FROM dim_calendar WHERE is_closed='Y'")).scalar()
    return set(skus), set(custs), closed


# ----------------------------------------------------------------------------- 2. transform
RENAME = {"Transaction_ID": "transaction_id", "Invoice_Date": "invoice_date",
          "SKU_Code": "sku_code", "Customer_Code": "customer_code", "Units": "units",
          "Gross_Sales_EUR": "gross_sales_eur", "Discount_EUR": "discount_eur", "Rebate_EUR": "rebate_eur"}
COLS = ["transaction_id", "invoice_date", "month_key", "sku_code", "customer_code",
        "units", "gross_sales_eur", "discount_eur", "rebate_eur"]


def transform_sales(df: pd.DataFrame, skus: set, custs: set, source_file: str):
    """Apply the 8 cleaning rules. Returns (clean, quarantine)."""
    s = df.copy()
    for c in ["SKU_Code", "Customer_Code"]:                       # rule 1: normalise codes
        s[c] = s[c].astype(str).str.strip().str.upper()
    s = s.rename(columns=RENAME)                                   # rule 8: snake_case
    s["month_key"] = s["invoice_date"].dt.year * 100 + s["invoice_date"].dt.month   # rule 7

    rejected = []
    def reject(mask, reason):
        nonlocal s
        rejected.append(s[mask].assign(reason=reason))
        s = s[~mask]

    reject(s.duplicated(subset="transaction_id", keep="first"), "duplicate transaction_id")   # rule 2
    reject(~s["sku_code"].isin(skus),        "sku not in dim_product")                       # rule 3
    reject(~s["customer_code"].isin(custs),  "customer not in dim_customer")                 # rule 4
    reject(s["invoice_date"].isna(),         "null invoice_date")                            # rule 6
    # rule 5 (negative units = returns) -> kept on purpose

    clean = s[COLS]
    quarantine = pd.concat(rejected, ignore_index=True)[COLS + ["reason"]] if rejected else pd.DataFrame(columns=COLS + ["reason"])
    quarantine["source_file"] = source_file
    return clean, quarantine


# ----------------------------------------------------------------------------- 3. load
def load_sales(engine, clean: pd.DataFrame, quarantine: pd.DataFrame) -> int:
    """Idempotent load: rows whose transaction_id already exists are skipped, not duplicated."""
    with engine.connect() as con:
        existing = pd.read_sql(
            text("SELECT transaction_id FROM fact_sales WHERE month_key BETWEEN :a AND :b"),
            con, params={"a": int(clean["month_key"].min()), "b": int(clean["month_key"].max())})
    new = clean[~clean["transaction_id"].isin(existing["transaction_id"])]

    new.to_sql("fact_sales", engine, if_exists="append", index=False, chunksize=5000)
    if len(quarantine):
        quarantine.to_sql("quarantine_sales", engine, if_exists="append", index=False)

    # close the months that now have actuals
    with engine.begin() as con:
        con.execute(text("UPDATE dim_calendar SET is_closed='Y' WHERE month_key <= :m"),
                    {"m": int(clean["month_key"].max())})
    return len(new)


# ----------------------------------------------------------------------------- 4. log
def write_log(engine, source_file, rows_read, rows_loaded, rows_quar, status, message=""):
    with engine.begin() as con:
        con.execute(text("""INSERT INTO etl_log (source_file, rows_read, rows_loaded, rows_quarantined, status, message)
                            VALUES (:f, :r, :l, :q, :s, :m)"""),
                    {"f": source_file, "r": rows_read, "l": rows_loaded, "q": rows_quar, "s": status, "m": message[:255]})


# ----------------------------------------------------------------------------- main
def main(file_arg: str):
    cfg, engine = load_config()
    path = Path(file_arg)
    if not path.is_absolute():
        path = Path(cfg["paths"]["inbox"]) / path
    src = path.name
    print(f"[ETL] {src}")

    try:
        raw = extract_sales(path)
        skus, custs, closed = fetch_dimensions(engine)
        clean, quar = transform_sales(raw, skus, custs, src)
        loaded = load_sales(engine, clean, quar)
        skipped = len(clean) - loaded
        msg = f"read {len(raw)} | clean {len(clean)} | loaded {loaded} | already present {skipped} | quarantined {len(quar)}"
        write_log(engine, src, len(raw), loaded, len(quar), "OK", msg)
        print("[ETL] OK  ", msg)

        archive = Path(cfg["paths"]["archive"]); archive.mkdir(parents=True, exist_ok=True)
        shutil.copy2(path, archive / src)          # keep a copy of what was loaded
    except Exception as e:
        write_log(engine, src, None, None, None, "FAIL", str(e))
        print("[ETL] FAIL", e)
        raise


if __name__ == "__main__":
    if len(sys.argv) != 2:
        sys.exit("usage: python etl_novapharma.py <ERP_Sales_YYYYMM.csv>")
    main(sys.argv[1])
