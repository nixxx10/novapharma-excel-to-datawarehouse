# NovaPharma Iberia — from an Excel model to a data warehouse

> **Status: in progress** (Steps 0–4 done, Step 5 closure next). Synthetic data, fictional company.

## The problem

A pharma commercial team in Iberia ran sales, budget, forecast and market share out of one Excel file.

- 17 sheets, 626k formulas, 6 VBA modules. Only one person understood the file.
- Lookups failed silently: 70 rows with lowercase SKU codes matched nothing and nobody noticed.
- Five layers mixed in one workbook: raw data, masters, calculations, reporting, parameters.
- Monthly refresh was manual: paste the ERP extract, run macros, check by eye.

## The proposed solution

```
Excel (as-is)  ──►  Python ETL  ──►  MySQL warehouse  ──►  Power BI
                    (clean, load,     (constellation:       (report,
                     log, schedule)    3 facts, 6 dims)      3 pages)
```

Each tool does one job and hands over to the next:

- **Python** (pandas + SQLAlchemy) reads the ERP extract, applies 8 cleaning rules and writes to MySQL. Bad rows go to a quarantine table, never deleted. Every run is logged and can be repeated without duplicating data. Windows Task Scheduler runs it monthly.
- **MySQL** stores the data once, in a constellation schema (3 fact tables, 6 dimensions). Keys and constraints refuse what Excel silently accepted. It runs locally because no cloud platform (Snowflake, BigQuery, Azure) was available; the design would move there unchanged.
- **Power BI** connects to two SQL views in the warehouse and only shows: 9 DAX measures, no calculations hidden in cells. The report is saved as a Power BI project (PBIP), so every page and visual is a readable text file in this repo.

What changes for the team: one source of truth instead of one file per person, refresh in seconds instead of a manual morning, errors visible instead of silent, and anyone can read the model, not only its author.

![ER diagram](02_data_model/er_diagram_novapharma_dw.png)

![Power BI report: Overview](05_powerbi/screenshots/01_overview.png)

## Steps

| Step | What | Where |
|---|---|---|
| 0 | Excel as-is: sheet dictionary, five layers found inside one file | `01_excel_as_is/`, `docs/Step0_*.pdf` |
| 1 | Reverse engineering the data model (type, grain, keys, measures) | `docs/Step1_*.pdf` |
| 2 | MySQL schema: DDL + ER diagram | `03_sql/01_schema.sql`, `02_data_model/` |
| 3.A | Python pipeline prototype: profile → rules → transform → load → reconcile | `04_python/01_pipeline_prototype.ipynb`, `docs/Step3A_*.pdf` |
| 3.B | Production script, idempotent monthly load, Task Scheduler | `04_python/etl_novapharma.py`, `docs/Step3B_*.pdf` |
| 4.A | Power BI model: SQL views, import, relationships | `03_sql/02_views.sql`, `docs/Step4A_*.pdf` |
| 4.B | DAX measures: why measures and not columns | `05_powerbi/`, `docs/Step4B_*.pdf` |
| 4.C | Visuals: Overview, Budget vs Actual, About | `05_powerbi/`, `docs/Step4C_*.pdf` |

## How to reproduce

1. MySQL 8: run `03_sql/01_schema.sql` (creates `novapharma_dw`, 11 tables).
2. `pip install -r requirements.txt`
3. Historical load: open `04_python/01_pipeline_prototype.ipynb`, set the Excel path, run all.
4. Monthly load: set the env var `NOVAPHARMA_DB_PWD`, adjust `04_python/config.ini`, then
   `python 04_python/etl_novapharma.py ERP_Sales_202609.csv`
5. Views: run `03_sql/02_views.sql`.
6. Report: open `05_powerbi/NovaPharma Iberia.pbip` in Power BI Desktop and refresh (MySQL on `localhost:3307`).

## LinkedIn series

| Post | Steps | Link |
|---|---|---|
| 1/5 | Step 0 — the strategy | [linkedin.com](https://www.linkedin.com/feed/update/urn:li:activity:7508117104609390592/) |
| 2/5 | Steps 1 & 2 — 3 lessons from reverse-engineering the model | [linkedin.com](https://www.linkedin.com/feed/update/urn:li:activity:7508830718685483009/) |

## What I learned

- Excel hides errors; a database refuses them. Quarantine bad rows, never delete them.
- One fact table per grain. Budget and forecast share a grain, so they share a table.
- Anything qualitative is a dimension; anything quantitative is a fact.
- A load must be re-runnable without duplicating data (idempotent).
- Reconcile to the cent: Excel total minus quarantine equals the SQL total, or the load is wrong.


---
Author: Nizar El Ouarma · MIT License
