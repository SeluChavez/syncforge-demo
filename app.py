"""SyncForge: demo UI for the supplier feed cleanup tool.

Run locally:  .venv\\Scripts\\streamlit run app.py --server.address localhost

All transformation logic lives in feedclean.py (unchanged). This file only stages files in a
temporary folder (deleted right after each analysis), calls feedclean.run() and shows the results.
Nothing is sent to Shopify and nothing is stored permanently.
"""
import io
import tempfile
import zipfile
from contextlib import contextmanager
from decimal import Decimal
from pathlib import Path

import streamlit as st

import feedclean

ROOT = Path(__file__).parent
SAMPLE_SUPPLIER = ROOT / "sample_data" / "supplier_feed.csv"
SAMPLE_STORE = ROOT / "sample_data" / "shopify_export.csv"
STORE_TOTAL = "Store total (Variant Inventory Qty)"

# Limits for a public demo. The upload size is also enforced by server.maxUploadSize.
MAX_UPLOAD_MB = 2
MAX_ROWS = 5000
MAX_XLSX_UNPACKED_MB = 20
READ_ERRORS = (ValueError, KeyError, StopIteration, zipfile.BadZipFile, UnicodeDecodeError)

st.set_page_config(page_title="SyncForge · Supplier Feed Cleanup", page_icon="🔄", layout="wide")


def check_size(name, data):
    if len(data) > MAX_UPLOAD_MB * 1024 * 1024:
        raise ValueError(f"{name} is larger than {MAX_UPLOAD_MB} MB (demo limit).")
    if name.lower().endswith(".xlsx"):
        with zipfile.ZipFile(io.BytesIO(data)) as z:
            if sum(i.file_size for i in z.infolist()) > MAX_XLSX_UNPACKED_MB * 1024 * 1024:
                raise ValueError(f"{name} unpacks to more than {MAX_XLSX_UNPACKED_MB} MB (demo limit).")


@contextmanager
def staged(files):
    """Write {key: (name, bytes)} to a temporary folder that is deleted on exit."""
    for name, data in files.values():
        check_size(name, data)
    with tempfile.TemporaryDirectory(prefix="syncforge_") as tmp:
        paths = {}
        for key, (name, data) in files.items():
            paths[key] = Path(tmp) / f"{key}{Path(name).suffix.lower()}"
            paths[key].write_bytes(data)
        for key, path in paths.items():
            if len(feedclean.read_table(path)[1]) > MAX_ROWS:
                raise ValueError(f"{files[key][0]} has more than {MAX_ROWS} rows (demo limit).")
        yield Path(tmp), paths


def locations_in(name, data):
    """Locations offered by the export's 'Inventory Available: <location>' columns."""
    try:
        with staged({"store": (name, data)}) as (_, paths):
            headers, _ = feedclean.read_table(paths["store"])
    except READ_ERRORS:
        return []
    prefix = "Inventory Available: "
    return [h[len(prefix):] for h in headers if h.startswith(prefix)]


def analyze(files, currency, location, max_change):
    with staged(files) as (tmp, paths):
        plan, written, report = feedclean.run(
            paths["supplier"], paths["store"], tmp / "out",
            max_price_change=Decimal(max_change) / 100, store_currency=currency, location=location)
        text = report.read_text(encoding="utf-8")
        for key, path in paths.items():  # show the uploaded file names, not temporary paths
            text = text.replace(str(path), files[key][0])
        return {"plan": plan, "files": {p.name: p.read_bytes() for p in written}, "report": text,
                "stock_column": feedclean.stock_column_for(location)}


def table(title, rows, empty="None."):
    st.markdown(f"**{title}**")
    if rows:
        st.dataframe([{k: "" if v is None else str(v) for k, v in row.items()} for row in rows],
                     hide_index=True, width="stretch")
    else:
        st.caption(empty)


# ---------------------------------------------------------------- header

st.title("SyncForge")
st.markdown("##### Supplier feed cleanup for Shopify")
st.caption(f"Demo. Files are processed in memory and deleted after each analysis; nothing is sent to "
           f"Shopify. Limits: {MAX_UPLOAD_MB} MB and {MAX_ROWS} rows per file. "
           "Please don't upload confidential data.")

# ---------------------------------------------------------------- inputs

left, right = st.columns(2)
supplier_upload = left.file_uploader("1. Supplier feed", type=["csv", "xlsx"])
store_upload = right.file_uploader("2. Shopify / Matrixify product export", type=["csv", "xlsx"])

c1, c2, c3 = st.columns(3)
currencies = list(feedclean.CURRENCY_CODES)
currency = c1.selectbox("Store currency", currencies, index=currencies.index(feedclean.DEFAULT_CURRENCY))
store_locations = locations_in(store_upload.name, store_upload.getvalue()) if store_upload else []
location = c2.selectbox("Inventory location", [STORE_TOTAL] + store_locations,
                        help="Stores with several locations: pick the location to update. "
                             "Locations are read from the export's 'Inventory Available' columns.")
location = None if location == STORE_TOTAL else location
max_change = c3.number_input("Hold price changes above (%)", min_value=1, max_value=100, value=20)

b1, b2 = st.columns([1, 3])
run_uploads = b1.button("Analyze Feed", type="primary", key="analyze",
                        disabled=not (supplier_upload and store_upload))
run_sample = b2.button("Try the demo with sample data", key="demo",
                       help="A messy synthetic supplier feed and a small coffee-equipment store (EUR).")

files = None
if run_uploads:
    files = {"supplier": (supplier_upload.name, supplier_upload.getvalue()),
             "store": (store_upload.name, store_upload.getvalue())}
elif run_sample:
    files = {"supplier": (SAMPLE_SUPPLIER.name, SAMPLE_SUPPLIER.read_bytes()),
             "store": (SAMPLE_STORE.name, SAMPLE_STORE.read_bytes())}
    location = None  # the sample export has a single stock column

if files:
    try:
        st.session_state.result = analyze(files, currency, location, max_change)
        st.session_state.result["source"] = "sample data" if run_sample else "your files"
    except READ_ERRORS as exc:
        st.session_state.pop("result", None)
        st.error(f"Could not process the files: {exc}")

# ---------------------------------------------------------------- results

result = st.session_state.get("result")
if result:
    plan = result["plan"]
    warnings = (len(plan.unusable_stock) + len(plan.unusable_price) + len(plan.unmatched)
                + len(plan.missing_from_feed) + len(plan.skipped_rows))

    st.divider()
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Safe updates", len(plan.updates))
    m2.metric("Blocked for review", plan.held_count)
    m3.metric("Warnings", warnings)
    m4.metric("Unchanged", len(plan.unchanged))
    st.caption(f"Results for {result['source']} · Store currency: {plan.store_currency} · "
               f"Stock column: `{result['stock_column']}`")

    foreign = sorted(plan.feed_currencies - {plan.store_currency})
    if foreign:
        st.warning(f"The feed contains prices in {', '.join(foreign)}. Prices without a currency symbol "
                   f"were assumed to be {plan.store_currency}. Confirm this with the supplier.")

    t_safe, t_blocked, t_warn, t_report = st.tabs(["Safe updates", "Blocked", "Warnings", "Audit report"])

    with t_safe:
        table("Changes in the import files", [
            {"SKU": u.variant.raw_sku, "Product": u.variant.title,
             "Price": f"{u.variant.price} → {u.new_price}" if u.new_price is not None else "unchanged",
             "Stock": f"{u.variant.stock} → {u.new_stock}" if u.new_stock is not None else "unchanged"}
            for u in plan.updates], empty="No safe updates.")

    with t_blocked:
        table(f"Price changes above {max_change}% (not imported)", [
            {"SKU": v.raw_sku, "Product": v.title, "Current": v.price, "Supplier": i.price,
             "Change": feedclean._pct(c)} for v, i, c in plan.price_review])
        table(f"Price in a currency other than {plan.store_currency} (not converted)", [
            {"SKU": v.raw_sku, "Product": v.title, "Current": v.price, "Supplier price": i.raw_price,
             "Detected": ", ".join(sorted(c))} for v, i, c in plan.currency_mismatch])
        table("Conflicting duplicate SKUs in the supplier feed", [
            {"SKU": g[0].raw_sku.strip(), "Feed lines": ", ".join(str(i.line) for i in g),
             "Prices": ", ".join(i.raw_price for i in g), "Stock": ", ".join(i.raw_stock for i in g)}
            for g in plan.feed_conflicts.values()])
        table("SKU shared by several store variants", [
            {"SKU": g[0].raw_sku, "Variants": ", ".join(f"{v.handle} ({v.variant_id})" for v in g)}
            for g in plan.store_duplicates.values()])

    with t_warn:
        table("Stock missing or unreadable (left untouched, never set to 0)", [
            {"SKU": v.raw_sku, "Product": v.title, "Supplier value": i.raw_stock or "(empty)",
             "Current stock": v.stock} for v, i in plan.unusable_stock])
        table("Price missing or unreadable (left untouched)", [
            {"SKU": v.raw_sku, "Product": v.title, "Supplier value": i.raw_price or "(empty)",
             "Current price": v.price} for v, i in plan.unusable_price])
        table("In the supplier feed but not in the store (not created)", [
            {"Feed line": i.line, "SKU": i.raw_sku, "Price": i.raw_price, "Stock": i.raw_stock}
            for i in plan.unmatched])
        table("In the store but not in the supplier feed (left untouched)", [
            {"SKU": v.raw_sku, "Product": v.title, "Current stock": v.stock} for v in plan.missing_from_feed])
        table("Skipped feed rows", [{"Feed line": line, "Reason": reason} for line, reason in plan.skipped_rows])

    with t_report:
        st.markdown(result["report"])

    st.divider()
    st.markdown("**Downloads**")
    st.caption("Import the files one at a time in Matrixify. A field absent from a file is not touched. "
               "Matrixify's free plan imports up to 10 items per file.")
    cols = st.columns(len(result["files"]) + 1)
    for col, (name, data) in zip(cols, result["files"].items()):
        col.download_button(name, data, file_name=name, mime="text/csv", key=f"dl_{name}")
    cols[-1].download_button("audit_report.md", result["report"].encode("utf-8"),
                             file_name="audit_report.md", mime="text/markdown", key="dl_report")
    if not result["files"]:
        st.caption("No import files: there is nothing safe to update.")
