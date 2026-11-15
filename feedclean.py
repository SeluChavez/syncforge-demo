"""Supplier feed cleanup: turn a messy supplier CSV into safe Matrixify import files.

Offline only. Reads a supplier CSV and a Shopify product export, writes import
CSVs plus a Markdown audit report. It never connects to Shopify.

Safety rules:
  * Missing or unreadable stock is never turned into 0: the stock is left untouched.
  * Price changes above the threshold (default 20%) are held for human review.
  * Prices marked with a currency other than the store currency are held for review.
    There is no currency conversion. Unmarked prices are assumed to be in the store currency.
  * Ambiguous SKUs (conflicting duplicates in the feed, or the same SKU on several
    store variants) are held entirely.
  * A field that should not change is left out of the import file rather than
    sent as a blank cell, so we never depend on how blank cells are interpreted.
"""

import argparse
import csv
import io
import re
import sys
import xml.etree.ElementTree as ET
import zipfile
from dataclasses import dataclass, field
from decimal import Decimal, InvalidOperation
from pathlib import Path

SKU_ALIASES = ("sku", "variant sku", "art.nr", "art. nr", "article", "article no",
               "item code", "product code", "ref", "reference")
PRICE_ALIASES = ("price", "unit price", "variant price", "retail price", "rrp")
STOCK_ALIASES = ("stock", "qty", "quantity", "inventory", "available", "on hand",
                 "variant inventory qty")

MISSING_TOKENS = {"", "-", "--", "n/a", "na", "null", "none", "?", "tbc", "unknown"}
DEFAULT_MAX_PRICE_CHANGE = Decimal("0.20")
DEFAULT_CURRENCY = "EUR"

# Longest symbols first, so "CA$" is not read as "$". A bare "$" is treated as USD;
# if the store uses another dollar currency, that price is held, which is the safe side.
CURRENCY_SYMBOLS = {"US$": "USD", "CA$": "CAD", "C$": "CAD", "AU$": "AUD", "A$": "AUD",
                    "NZ$": "NZD", "$": "USD", "€": "EUR", "£": "GBP", "¥": "JPY"}
CURRENCY_CODES = ("EUR", "USD", "GBP", "CHF", "CAD", "AUD", "NZD", "SEK", "NOK", "DKK",
                  "PLN", "CZK", "JPY")
_CURRENCY_CODE_RE = re.compile(r"(?<![A-Z])(" + "|".join(CURRENCY_CODES) + r")(?![A-Z])")


# ---------------------------------------------------------------- parsing

_SKU_SEPARATORS = re.compile(r"[\s\-_./]+")


def normalize_sku(raw):
    """Matching key for a SKU: uppercase, no separators, no leading zeros on numeric SKUs.

    'ab-1001 ', 'AB_1001' and 'AB 1001' all become 'AB1001'; '000456' and '456'
    (Excel drops leading zeros) both become '456'. Collisions this might cause are
    caught later by duplicate detection.
    """
    sku = _SKU_SEPARATORS.sub("", (raw or "").upper())
    if sku.isdigit():
        sku = sku.lstrip("0") or "0"
    return sku


def parse_price(raw):
    """Parse '€24,90', '19.99 EUR', '1.249,00', '1,234.50', '$18' into Decimal.

    Returns None for anything missing, ambiguous, zero or negative.
    """
    text = (raw or "").strip()
    if text.lower() in MISSING_TOKENS:
        return None
    text = re.sub(r"[^\d.,\-]", "", text)  # drop currency symbols, codes, spaces
    if not text or "-" in text:
        return None
    if "," in text and "." in text:
        decimal_sep = "," if text.rfind(",") > text.rfind(".") else "."
        thousands_sep = "." if decimal_sep == "," else ","
        text = text.replace(thousands_sep, "").replace(decimal_sep, ".")
    elif "," in text or "." in text:
        sep = "," if "," in text else "."
        parts = text.split(sep)
        if len(parts) == 2 and 1 <= len(parts[1]) <= 2:
            text = parts[0] + "." + parts[1]          # decimal separator
        elif parts[0] and len(parts[0]) <= 3 and all(len(p) == 3 for p in parts[1:]):
            text = "".join(parts)                     # thousands grouping, e.g. 1.249
        else:
            return None
    try:
        value = Decimal(text)
    except InvalidOperation:
        return None
    if value <= 0:
        return None
    return value.quantize(Decimal("0.01"))


def detect_currencies(raw):
    """Currencies explicitly written in a price: '€24,90' -> {'EUR'}, '45,00' -> set()."""
    text = (raw or "").upper()
    found = set()
    for symbol, code in CURRENCY_SYMBOLS.items():
        if symbol in text:
            found.add(code)
            text = text.replace(symbol, " ")
    found.update(_CURRENCY_CODE_RE.findall(text))
    return found


def parse_stock(raw):
    """Parse a stock quantity. Returns None (= unknown, do not touch) unless clearly valid.

    '15' -> 15, '0' -> 0, '12.0' -> 12, '1,200' -> 1200.
    '', 'N/A', '-3', '5.5', '>50' -> None. Unknown stock is never treated as 0.
    """
    text = (raw or "").strip().replace(" ", "")
    if text.lower() in MISSING_TOKENS:
        return None
    if re.fullmatch(r"\d{1,3}(,\d{3})+", text):
        text = text.replace(",", "")
    if re.fullmatch(r"\d+(\.0+)?", text):
        return int(text.split(".")[0])
    return None


# ---------------------------------------------------------------- reading CSVs

def read_csv(path):
    """Read a CSV with unknown encoding (UTF-8, UTF-8 BOM, cp1252) and delimiter."""
    data = Path(path).read_bytes()
    try:
        text = data.decode("utf-8-sig")
    except UnicodeDecodeError:
        text = data.decode("cp1252")
    header_line = text.splitlines()[0] if text else ""
    delimiter = max(",;\t|", key=header_line.count)
    reader = csv.DictReader(io.StringIO(text), delimiter=delimiter)
    headers = [(h or "").strip() for h in (reader.fieldnames or [])]
    reader.fieldnames = headers
    rows = []
    for row in reader:
        rows.append({k: (v or "").strip() for k, v in row.items() if isinstance(k, str)})
    return headers, rows


_XLSX_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
_XLSX_REL_ID = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id"


def _xlsx_column(ref):
    n = 0
    for ch in re.match(r"[A-Z]*", ref).group():
        n = n * 26 + ord(ch) - 64
    return n - 1


def _xlsx_number(text):
    """Excel stores numbers as floats: '699.95000000000005' -> '699.95', '6.8E+13' -> digits."""
    try:
        value = Decimal(text)
    except InvalidOperation:
        return text
    if value == value.to_integral_value():
        return str(int(value))
    return repr(float(text))


def read_xlsx(path):
    """Read the 'Products' sheet (or the first sheet) of an XLSX file. Read-only, standard library."""
    with zipfile.ZipFile(path) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            for si in ET.fromstring(z.read("xl/sharedStrings.xml")).iter(_XLSX_MAIN + "si"):
                shared.append("".join(t.text or "" for t in si.iter(_XLSX_MAIN + "t")))
        sheets = list(ET.fromstring(z.read("xl/workbook.xml")).iter(_XLSX_MAIN + "sheet"))
        sheet = next((s for s in sheets if s.get("name", "").lower() == "products"), sheets[0])
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
        target = next(r.get("Target") for r in rels if r.get("Id") == sheet.get(_XLSX_REL_ID)).lstrip("/")
        target = target if target.startswith("xl/") else "xl/" + target
        grid = []
        for row in ET.fromstring(z.read(target)).iter(_XLSX_MAIN + "row"):
            cells = {}
            for c in row.iter(_XLSX_MAIN + "c"):
                kind, v = c.get("t"), c.find(_XLSX_MAIN + "v")
                if kind == "inlineStr":
                    value = "".join(t.text or "" for t in c.iter(_XLSX_MAIN + "t"))
                elif v is None:
                    value = ""
                elif kind == "s":
                    value = shared[int(v.text)]
                elif kind in ("str", "b", "e"):
                    value = v.text or ""
                else:
                    value = _xlsx_number(v.text or "")
                cells[_xlsx_column(c.get("r", "")) if c.get("r") else len(cells)] = value.strip()
            grid.append(cells)
    if not grid or not grid[0]:
        return [], []
    headers = [grid[0].get(i, "") for i in range(max(grid[0]) + 1)]
    rows = [{h: cells.get(i, "") for i, h in enumerate(headers) if h} for cells in grid[1:]]
    return headers, rows


def read_table(path):
    return read_xlsx(path) if Path(path).suffix.lower() == ".xlsx" else read_csv(path)


def find_column(headers, aliases, required=True):
    lookup = {h.lower(): h for h in headers}
    for alias in aliases:
        if alias in lookup:
            return lookup[alias]
    if required:
        raise ValueError(f"None of the columns {headers} matches any of {list(aliases)}")
    return None


@dataclass
class FeedItem:
    line: int
    raw_sku: str
    sku: str
    raw_price: str
    raw_stock: str
    price: Decimal | None
    stock: int | None


@dataclass
class StoreVariant:
    handle: str
    title: str
    variant_id: str
    raw_sku: str
    sku: str
    price: Decimal | None
    stock: int | None


def load_feed(path):
    headers, rows = read_table(path)
    sku_col = find_column(headers, SKU_ALIASES)
    price_col = find_column(headers, PRICE_ALIASES)
    stock_col = find_column(headers, STOCK_ALIASES)
    items = []
    for line, row in enumerate(rows, start=2):  # line 1 is the header
        raw_sku, raw_price, raw_stock = row.get(sku_col, ""), row.get(price_col, ""), row.get(stock_col, "")
        items.append(FeedItem(line, raw_sku, normalize_sku(raw_sku), raw_price, raw_stock,
                              parse_price(raw_price), parse_stock(raw_stock)))
    return items


def stock_column_for(location):
    """Matrixify column for stock: one location's column, or the store-wide total."""
    return f"Inventory Available: {location}" if location else "Variant Inventory Qty"


def load_store(path, location=None):
    headers, rows = read_table(path)
    handle_col = find_column(headers, ("handle",))
    sku_col = find_column(headers, ("variant sku", "sku"))
    price_col = find_column(headers, ("variant price", "price"))
    if location:
        stock_col = find_column(headers, (stock_column_for(location).lower(),))
    else:
        stock_col = find_column(headers, ("variant inventory qty", "inventory", "stock"), required=False)
    title_col = find_column(headers, ("title",), required=False)
    id_col = find_column(headers, ("variant id",), required=False)
    variants = []
    for row in rows:
        raw_sku = row.get(sku_col, "")
        variants.append(StoreVariant(
            handle=row.get(handle_col, ""),
            title=row.get(title_col, "") if title_col else "",
            variant_id=row.get(id_col, "") if id_col else "",
            raw_sku=raw_sku,
            sku=normalize_sku(raw_sku),
            price=parse_price(row.get(price_col, "")),
            stock=parse_stock(row.get(stock_col, "")) if stock_col else None,
        ))
    return variants


# ---------------------------------------------------------------- planning

@dataclass
class Update:
    variant: StoreVariant
    item: FeedItem
    new_price: Decimal | None  # None = leave the price untouched
    new_stock: int | None      # None = leave the stock untouched

    @property
    def fields(self):
        return tuple(name for name, value in (("price", self.new_price), ("stock", self.new_stock))
                     if value is not None)


@dataclass
class Plan:
    updates: list = field(default_factory=list)
    unchanged: list = field(default_factory=list)          # (variant, item)
    price_review: list = field(default_factory=list)       # (variant, item, change or None)
    currency_mismatch: list = field(default_factory=list)  # (variant, item, currencies), held
    unusable_price: list = field(default_factory=list)     # (variant, item)
    unusable_stock: list = field(default_factory=list)     # (variant, item)
    feed_conflicts: dict = field(default_factory=dict)     # sku -> [FeedItem], held
    feed_merged: dict = field(default_factory=dict)        # sku -> [FeedItem], identical rows merged
    store_duplicates: dict = field(default_factory=dict)   # sku -> [StoreVariant], held
    unmatched: list = field(default_factory=list)          # FeedItem not in the store
    missing_from_feed: list = field(default_factory=list)  # StoreVariant not in the feed
    skipped_rows: list = field(default_factory=list)       # (line, reason)
    store_currency: str = DEFAULT_CURRENCY
    feed_currencies: set = field(default_factory=set)      # all currencies written in the feed

    @property
    def held_count(self):
        return (len(self.price_review) + len(self.currency_mismatch)
                + len(self.feed_conflicts) + len(self.store_duplicates))


def _group(objects):
    groups = {}
    for obj in objects:
        groups.setdefault(obj.sku, []).append(obj)
    return groups


def build_plan(items, variants, max_price_change=DEFAULT_MAX_PRICE_CHANGE,
               store_currency=DEFAULT_CURRENCY):
    store_currency = store_currency.upper()
    plan = Plan(store_currency=store_currency)
    for item in items:
        plan.feed_currencies |= detect_currencies(item.raw_price)

    usable_items = []
    for item in items:
        if item.sku:
            usable_items.append(item)
        else:
            plan.skipped_rows.append((item.line, "empty SKU"))

    feed = {}
    for sku, group in _group(usable_items).items():
        if len(group) > 1:
            if len({(i.price, i.stock, frozenset(detect_currencies(i.raw_price))) for i in group}) > 1:
                plan.feed_conflicts[sku] = group
                continue
            plan.feed_merged[sku] = group
        feed[sku] = group[0]

    store = {}
    for sku, group in _group(v for v in variants if v.sku).items():
        if len(group) > 1:
            plan.store_duplicates[sku] = group
        else:
            store[sku] = group[0]

    for sku, item in feed.items():
        if sku in plan.store_duplicates:
            continue  # ambiguous: would update several variants at once
        variant = store.get(sku)
        if variant is None:
            plan.unmatched.append(item)
            continue

        new_price = None
        currencies = detect_currencies(item.raw_price)
        if item.price is None:
            plan.unusable_price.append((variant, item))
        elif currencies - {store_currency}:
            plan.currency_mismatch.append((variant, item, currencies))
        elif item.price != variant.price:
            if variant.price is None:
                plan.price_review.append((variant, item, None))
            else:
                change = (item.price - variant.price) / variant.price
                if abs(change) > max_price_change:
                    plan.price_review.append((variant, item, change))
                else:
                    new_price = item.price

        new_stock = None
        if item.stock is None:
            plan.unusable_stock.append((variant, item))
        elif item.stock != variant.stock:
            new_stock = item.stock

        if new_price is not None or new_stock is not None:
            plan.updates.append(Update(variant, item, new_price, new_stock))
        elif item.price == variant.price and item.stock == variant.stock and not currencies - {store_currency}:
            plan.unchanged.append((variant, item))

    feed_skus = {i.sku for i in usable_items}
    plan.missing_from_feed = [v for v in store.values() if v.sku not in feed_skus]
    return plan


# ---------------------------------------------------------------- output

# Matrixify detects the data type of a CSV from its file name, so it must contain "Products".
FILE_NAMES = {
    ("price", "stock"): "Products_update_price_and_stock.csv",
    ("price",): "Products_update_price_only.csv",
    ("stock",): "Products_update_stock_only.csv",
}


def write_matrixify(plan, out_dir, stock_column="Variant Inventory Qty"):
    """Write one Matrixify file per set of changed fields. Returns the written paths."""
    out_dir = Path(out_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    written = []
    for fields, name in FILE_NAMES.items():
        updates = [u for u in plan.updates if u.fields == fields]
        if not updates:
            continue
        columns = ["Handle", "Command", "Variant ID", "Variant SKU", "Variant Command"]
        if "price" in fields:
            columns.append("Variant Price")
        if "stock" in fields:
            columns.append(stock_column)
        path = out_dir / name
        with path.open("w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerow(columns)
            for u in updates:
                row = [u.variant.handle, "UPDATE", u.variant.variant_id, u.variant.raw_sku, "UPDATE"]
                if "price" in fields:
                    row.append(str(u.new_price))
                if "stock" in fields:
                    row.append(str(u.new_stock))
                writer.writerow(row)
        written.append(path)
    return written


def _cell(value):
    return "" if value is None else str(value).replace("|", "\\|")


def _table(headers, rows):
    if not rows:
        return "_None._\n"
    lines = ["| " + " | ".join(headers) + " |", "|" + "---|" * len(headers)]
    lines += ["| " + " | ".join(_cell(c) for c in row) + " |" for row in rows]
    return "\n".join(lines) + "\n"


def _pct(change):
    return "n/a" if change is None else f"{change * 100:+.1f}%"


def _raw(text):
    return f"`{text}`" if text else "_(empty)_"


def render_report(plan, supplier_path, store_path, written, max_price_change,
                  stock_column="Variant Inventory Qty"):
    foreign = sorted(plan.feed_currencies - {plan.store_currency})
    out = [
        "# Supplier feed audit report\n",
        f"- Supplier feed: `{supplier_path}`",
        f"- Store export: `{store_path}`",
        f"- Store currency: {plan.store_currency}",
        f"- Stock column: `{stock_column}`",
        f"- Currencies written in the feed: {', '.join(sorted(plan.feed_currencies)) or 'none'}",
        f"- Price change threshold: {max_price_change * 100:.0f}%\n",
    ]
    if foreign:
        out.append(f"> **Warning:** the feed contains prices in {', '.join(foreign)}. Prices without a "
                   f"currency symbol were assumed to be {plan.store_currency}. Confirm this with the supplier.\n")
    out += [
        "## Summary\n",
        _table(["Result", "Count"], [
            ["Variants with safe updates", len(plan.updates)],
            ["Unchanged", len(plan.unchanged)],
            ["Held for review (price jump, currency mismatch or ambiguous SKU)", plan.held_count],
            ["Stock missing/unreadable (left untouched)", len(plan.unusable_stock)],
            ["Price missing/unreadable (left untouched)", len(plan.unusable_price)],
            ["In feed but not in store", len(plan.unmatched)],
            ["In store but not in feed (not zeroed)", len(plan.missing_from_feed)],
            ["Feed rows skipped", len(plan.skipped_rows)],
        ]),
        "Generated import files: " + (", ".join(f"`{p.name}`" for p in written) or "_none_") + "\n",
        "## Ready to import\n",
        _table(["SKU", "Product", "Price", "Stock"], [
            [u.variant.raw_sku, u.variant.title,
             f"{u.variant.price} → {u.new_price}" if u.new_price is not None else "unchanged",
             f"{u.variant.stock} → {u.new_stock}" if u.new_stock is not None else "unchanged"]
            for u in plan.updates]),
        f"## Price changes above {max_price_change * 100:.0f}% (held for review, not imported)\n",
        _table(["SKU", "Product", "Current", "Supplier", "Change"], [
            [v.raw_sku, v.title, v.price, i.price, _pct(c)] for v, i, c in plan.price_review]),
        f"## Price in a currency other than {plan.store_currency} (held for review, not converted)\n",
        _table(["SKU", "Product", "Current price", "Raw supplier price", "Detected"], [
            [v.raw_sku, v.title, v.price, _raw(i.raw_price), ", ".join(sorted(c))]
            for v, i, c in plan.currency_mismatch]),
        "## Conflicting duplicate SKUs in the supplier feed (held)\n",
        _table(["SKU", "Feed lines", "Raw prices", "Raw stock"], [
            [g[0].raw_sku.strip(), ", ".join(str(i.line) for i in g),
             ", ".join(_raw(i.raw_price) for i in g), ", ".join(_raw(i.raw_stock) for i in g)]
            for g in plan.feed_conflicts.values()]),
        "## Identical duplicate rows in the supplier feed (merged)\n",
        _table(["SKU", "Feed lines"], [
            [g[0].raw_sku.strip(), ", ".join(str(i.line) for i in g)] for g in plan.feed_merged.values()]),
        "## SKU used by several store variants (held: an import would update all of them)\n",
        _table(["SKU", "Variants"], [
            [g[0].raw_sku, ", ".join(f"{v.handle} ({v.variant_id})" for v in g)]
            for g in plan.store_duplicates.values()]),
        "## Stock missing or unreadable (stock left untouched, never set to 0)\n",
        _table(["SKU", "Product", "Raw stock value", "Current stock"], [
            [v.raw_sku, v.title, _raw(i.raw_stock), v.stock] for v, i in plan.unusable_stock]),
        "## Price missing or unreadable (price left untouched)\n",
        _table(["SKU", "Product", "Raw price value", "Current price"], [
            [v.raw_sku, v.title, _raw(i.raw_price), v.price] for v, i in plan.unusable_price]),
        "## In the supplier feed but not in the store\n",
        _table(["Feed line", "Raw SKU", "Price", "Stock"], [
            [i.line, _raw(i.raw_sku), _raw(i.raw_price), _raw(i.raw_stock)] for i in plan.unmatched]),
        "## In the store but not in the supplier feed (left untouched)\n",
        _table(["SKU", "Product", "Current stock"], [
            [v.raw_sku, v.title, v.stock] for v in plan.missing_from_feed]),
        "## Skipped feed rows\n",
        _table(["Feed line", "Reason"], plan.skipped_rows),
    ]
    return "\n".join(out)


def run(supplier_path, store_path, out_dir, max_price_change=DEFAULT_MAX_PRICE_CHANGE,
        store_currency=DEFAULT_CURRENCY, location=None):
    stock_column = stock_column_for(location)
    plan = build_plan(load_feed(supplier_path), load_store(store_path, location),
                      max_price_change, store_currency)
    written = write_matrixify(plan, out_dir, stock_column)
    report = Path(out_dir) / "audit_report.md"
    report.write_text(render_report(plan, supplier_path, store_path, written, max_price_change,
                                    stock_column), encoding="utf-8")
    return plan, written, report


def main(argv=None):
    parser = argparse.ArgumentParser(description="Clean a supplier CSV into safe Matrixify import files.")
    parser.add_argument("--supplier", required=True, help="supplier feed (CSV or XLSX)")
    parser.add_argument("--store", required=True, help="Shopify/Matrixify product export (CSV or XLSX)")
    parser.add_argument("--out", default="output", help="output folder (default: output)")
    parser.add_argument("--max-price-change", type=Decimal, default=Decimal("20"),
                        help="hold price changes above this percent (default: 20)")
    parser.add_argument("--currency", default=DEFAULT_CURRENCY,
                        help="store currency code; prices marked with another currency are held (default: EUR)")
    parser.add_argument("--location",
                        help="update stock at this location only, via 'Inventory Available: <location>' "
                             "(use for stores with several locations)")
    args = parser.parse_args(argv)

    plan, written, report = run(args.supplier, args.store, args.out, args.max_price_change / 100,
                                args.currency, args.location)
    print(f"Safe updates:        {len(plan.updates)}")
    print(f"Held for review:     {plan.held_count}")
    print(f"Stock left untouched (missing/unreadable): {len(plan.unusable_stock)}")
    print(f"Not in store: {len(plan.unmatched)}   Not in feed: {len(plan.missing_from_feed)}")
    for path in written:
        print(f"Wrote {path}")
    print(f"Wrote {report}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
