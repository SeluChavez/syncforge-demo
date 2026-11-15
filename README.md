# Supplier feed cleanup (proof of concept)

Turns a messy supplier CSV into safe Matrixify import files for Shopify, plus an audit report.
Offline only: it reads two CSV files and never connects to Shopify. Python standard library only.

## Run

```
python feedclean.py --supplier sample_data/supplier_feed.csv --store sample_data/shopify_export.csv --out output
python -m unittest -v
```

### SyncForge demo UI (optional)

A local Streamlit page on top of `feedclean.py`: upload a supplier feed and a product export
(or toggle the sample data), choose currency and location, click **Analyze Feed**, download the files.

```
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\streamlit.exe run app.py --server.address localhost
```

It never contacts Shopify. Uploads are limited to 2 MB and 5,000 rows per file, are kept in memory,
and are written only to a temporary folder that is deleted after each analysis.
`.venv\Scripts\python.exe -m unittest -v` also runs the UI smoke test (it is skipped without Streamlit).

Options:
- `--max-price-change 20`: percent; larger changes are held for review.
- `--location "Shop location"`: for stores with several locations. Stock is read from and written to
  `Inventory Available: <location>` only. Without it, `Variant Inventory Qty` is used, which Matrixify
  defines as the total across all locations.
- `--store` and `--supplier` accept CSV or XLSX (e.g. a Matrixify export). XLSX files are only read.
- `--currency EUR`: the store currency. Prices marked with another currency (`$18`, `12 USD`, `£5`)
  are held for review and never converted. Prices without a symbol are assumed to be in the store
  currency, and the report warns when the feed mixes currencies.

## Output

- `Products_update_price_and_stock.csv`, `Products_update_price_only.csv`, `Products_update_stock_only.csv`:
  one file per set of changed fields. Matrixify leaves absent columns untouched but clears fields
  whose cell is empty, so a field that must not change is absent rather than blank.
  "Products" in the file name lets Matrixify detect the data type of a CSV.
- `audit_report.md`: what will change, what was held and why.

## Safety rules

- Missing or unreadable stock (`""`, `N/A`, `-3`, `5.5`) is never written as 0; the stock is left untouched.
- Price changes above the threshold are held; the stock update for that SKU still goes through.
- Conflicting duplicate SKUs in the feed, and SKUs shared by several store variants, are held entirely.
- Store products missing from the feed are reported, not zeroed. Feed products missing from the store are reported, not created.
- Only actual changes are exported.

## Known limitations

- Matrixify column names (`Command`, `Variant Command`, `Variant Price`, `Variant Inventory Qty`) follow
  Matrixify's documented format but have not been tested against a real store. Do a test import on a
  development store first.
- Without `--location`, `Variant Inventory Qty` assumes a single inventory location.
- The supplier price is used as the selling price as-is (no markup rules).
- Ambiguous price formats: a single separator followed by exactly three digits (`1.249`) is read as a
  thousands separator. A wrong guess like this would show up as a held price jump.
- Column detection uses a fixed list of header names (see `SKU_ALIASES` etc. in `feedclean.py`).
