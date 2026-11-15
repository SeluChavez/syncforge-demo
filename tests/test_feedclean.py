import csv
import tempfile
import unittest
import zipfile
from decimal import Decimal
from pathlib import Path

from feedclean import (FeedItem, StoreVariant, _xlsx_number, build_plan, detect_currencies,
                       normalize_sku, parse_price, parse_stock, read_csv, read_xlsx, run)

ROOT = Path(__file__).resolve().parent.parent
SUPPLIER = ROOT / "sample_data" / "supplier_feed.csv"
STORE = ROOT / "sample_data" / "shopify_export.csv"


def item(sku, price="10,00", stock="5", line=2):
    return FeedItem(line, sku, normalize_sku(sku), price, stock, parse_price(price), parse_stock(stock))


def variant(sku, price="10.00", stock="5", handle="h", variant_id="1"):
    return StoreVariant(handle, "Title", variant_id, sku, normalize_sku(sku),
                        parse_price(price), parse_stock(stock))


def read_rows(path):
    with open(path, newline="", encoding="utf-8") as f:
        return list(csv.DictReader(f))


class NormalizeSkuTest(unittest.TestCase):
    def test_case_whitespace_and_separators(self):
        for raw in ("  ab-1001 ", "AB_1001", "ab 1001", "AB.1001", "AB1001"):
            self.assertEqual(normalize_sku(raw), "AB1001", raw)

    def test_leading_zeros_only_stripped_from_numeric_skus(self):
        self.assertEqual(normalize_sku("000456"), "456")
        self.assertEqual(normalize_sku("456"), "456")
        self.assertEqual(normalize_sku("0AB1"), "0AB1")
        self.assertEqual(normalize_sku("000"), "0")

    def test_empty(self):
        self.assertEqual(normalize_sku(""), "")
        self.assertEqual(normalize_sku(None), "")


class ParsePriceTest(unittest.TestCase):
    def test_formats(self):
        cases = {
            "€24,90": "24.90", "19.99 EUR": "19.99", "1.249,00": "1249.00",
            "1,234.50": "1234.50", "$18": "18.00", "€ 9,50": "9.50", "15,5": "15.50",
            "1.249": "1249.00", "12.50": "12.50", "1199.00": "1199.00",
        }
        for raw, expected in cases.items():
            self.assertEqual(parse_price(raw), Decimal(expected), raw)

    def test_invalid_prices_return_none(self):
        for raw in ("", "N/A", "price on request", "0,00", "-5", "12.5.0", "1,2345"):
            self.assertIsNone(parse_price(raw), raw)


class ParseStockTest(unittest.TestCase):
    def test_valid(self):
        for raw, expected in {"15": 15, "0": 0, "12.0": 12, "1,200": 1200, " 7 ": 7}.items():
            self.assertEqual(parse_stock(raw), expected, raw)

    def test_missing_or_invalid_is_none_never_zero(self):
        for raw in ("", "N/A", "n/a", "-", "-3", "5.5", ">50", "unknown", None):
            self.assertIsNone(parse_stock(raw), raw)


class DetectCurrenciesTest(unittest.TestCase):
    def test_symbols_and_codes(self):
        cases = {
            "€24,90": {"EUR"}, "19.99 EUR": {"EUR"}, "eur 5": {"EUR"}, "$18": {"USD"},
            "US$18": {"USD"}, "18 USD": {"USD"}, "CA$18": {"CAD"}, "£5,00": {"GBP"},
            "€10 / $12": {"EUR", "USD"},
        }
        for raw, expected in cases.items():
            self.assertEqual(detect_currencies(raw), expected, raw)

    def test_no_currency(self):
        for raw in ("45,00", "1.249,00", "", None, "price on request"):
            self.assertEqual(detect_currencies(raw), set(), raw)


class CurrencyPlanTest(unittest.TestCase):
    def test_eur_price_accepted_in_eur_store(self):
        plan = build_plan([item("A-1", price="€11,00")], [variant("A-1")], store_currency="EUR")
        self.assertEqual(plan.updates[0].new_price, Decimal("11.00"))
        self.assertEqual(plan.currency_mismatch, [])

    def test_usd_price_held_in_eur_store_without_conversion(self):
        plan = build_plan([item("A-1", price="$11", stock="9")], [variant("A-1", stock="5")],
                          store_currency="EUR")
        self.assertEqual(len(plan.currency_mismatch), 1)
        self.assertEqual(plan.currency_mismatch[0][2], {"USD"})
        self.assertIsNone(plan.updates[0].new_price)       # price untouched, not converted
        self.assertEqual(plan.updates[0].new_stock, 9)     # stock handled as usual

    def test_usd_price_accepted_in_usd_store_and_eur_held(self):
        plan = build_plan([item("A-1", price="$11"), item("B-1", price="11 EUR")],
                          [variant("A-1"), variant("B-1")], store_currency="usd")
        self.assertEqual([u.variant.sku for u in plan.updates], ["A1"])
        self.assertEqual([v.sku for v, _, _ in plan.currency_mismatch], ["B1"])

    def test_missing_currency_symbol_assumes_store_currency(self):
        plan = build_plan([item("A-1", price="11,00")], [variant("A-1")], store_currency="EUR")
        self.assertEqual(plan.updates[0].new_price, Decimal("11.00"))
        self.assertEqual(plan.currency_mismatch, [])

    def test_mixed_currency_feed_holds_only_foreign_prices(self):
        items = [item("A-1", price="€11,00"), item("B-1", price="$11"), item("C-1", price="11,00"),
                 item("D-1", price="£11")]
        variants = [variant(s) for s in ("A-1", "B-1", "C-1", "D-1")]
        plan = build_plan(items, variants, store_currency="EUR")
        self.assertEqual(sorted(u.variant.sku for u in plan.updates), ["A1", "C1"])
        self.assertEqual(sorted(v.sku for v, _, _ in plan.currency_mismatch), ["B1", "D1"])
        self.assertEqual(plan.feed_currencies, {"EUR", "USD", "GBP"})

    def test_mixed_currency_in_one_value_is_held(self):
        plan = build_plan([item("A-1", price="€10 / $12")], [variant("A-1")], store_currency="EUR")
        self.assertEqual(len(plan.currency_mismatch), 1)

    def test_held_price_with_same_stock_changes_nothing(self):
        plan = build_plan([item("A-1", price="$10", stock="5")], [variant("A-1", price="10.00", stock="5")],
                          store_currency="EUR")
        self.assertEqual(plan.updates, [])
        self.assertEqual(plan.unchanged, [])
        self.assertEqual(len(plan.currency_mismatch), 1)

    def test_duplicates_differing_only_by_currency_conflict(self):
        plan = build_plan([item("A-1", price="€18", line=2), item("A-1", price="$18", line=3)],
                          [variant("A-1")], store_currency="EUR")
        self.assertIn("A1", plan.feed_conflicts)
        self.assertEqual(plan.updates, [])


class ReadCsvTest(unittest.TestCase):
    def test_semicolon_and_utf8_bom(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "feed.csv"
            path.write_bytes("﻿SKU;Price;Stock\nA-1;€1,50;3\n".encode("utf-8"))
            headers, rows = read_csv(path)
        self.assertEqual(headers, ["SKU", "Price", "Stock"])
        self.assertEqual(rows, [{"SKU": "A-1", "Price": "€1,50", "Stock": "3"}])

    def test_cp1252_fallback(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "feed.csv"
            path.write_bytes("SKU,Price\nA-1,€2\n".encode("cp1252"))
            _, rows = read_csv(path)
        self.assertEqual(rows[0]["Price"], "€2")


def write_xlsx(path, header, rows):
    """Minimal XLSX: header as shared strings, values as numbers or inline strings."""
    strings = list(header)

    def cell(ref, value):
        if isinstance(value, (int, float)) or (isinstance(value, str) and value.replace(".", "").isdigit()):
            return f'<c r="{ref}"><v>{value}</v></c>'
        return f'<c r="{ref}" t="inlineStr"><is><t>{value}</t></is></c>'

    letters = "ABCDEFGHIJ"
    sheet_rows = ['<row r="1">' + "".join(f'<c r="{letters[i]}1" t="s"><v>{i}</v></c>'
                                          for i in range(len(header))) + "</row>"]
    for n, row in enumerate(rows, start=2):
        sheet_rows.append(f'<row r="{n}">' + "".join(cell(f"{letters[i]}{n}", v)
                                                     for i, v in enumerate(row) if v != "") + "</row>")
    main = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    with zipfile.ZipFile(path, "w") as z:
        z.writestr("xl/workbook.xml", f'<workbook xmlns="{main}" xmlns:r="{rel}"><sheets>'
                   '<sheet name="Products" sheetId="1" r:id="rId1"/></sheets></workbook>')
        z.writestr("xl/_rels/workbook.xml.rels",
                   '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                   '<Relationship Id="rId1" Target="worksheets/sheet1.xml" '
                   f'Type="{rel}/worksheet"/></Relationships>')
        z.writestr("xl/sharedStrings.xml", f'<sst xmlns="{main}">' +
                   "".join(f"<si><t>{s}</t></si>" for s in strings) + "</sst>")
        z.writestr("xl/worksheets/sheet1.xml", f'<worksheet xmlns="{main}"><sheetData>' +
                   "".join(sheet_rows) + "</sheetData></worksheet>")


class XlsxTest(unittest.TestCase):
    def test_excel_number_cleanup(self):
        self.assertEqual(_xlsx_number("699.95000000000005"), "699.95")
        self.assertEqual(_xlsx_number("68237159268644"), "68237159268644")
        self.assertEqual(_xlsx_number("6.8237159268644E13"), "68237159268644")
        self.assertEqual(_xlsx_number("600"), "600")

    def test_reads_matrixify_style_export_with_location_column(self):
        header = ["ID", "Handle", "Variant ID", "Variant SKU", "Variant Price",
                  "Variant Inventory Qty", "Inventory Available: Shop location"]
        rows = [["15450996441380", "minimal", "68237159366948", "SNOW-001", "885.95000000000005", "50", "50"],
                ["15450996572452", "out-of-stock", "68237159891236", "SNOW-003", "885.95", "0", "0"],
                ["###### YOUR PLAN ALLOWS FILE SIZE TILL HERE ######", "", "", "", "", "", ""]]
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "Export.xlsx"
            write_xlsx(path, header, rows)
            headers, parsed = read_xlsx(path)
            feed = Path(tmp) / "feed.csv"
            feed.write_text("SKU;Price;Stock\nsnow_001;$899.95;45\nSNOW-003;885.95;12\n", encoding="utf-8")
            plan, written, _ = run(feed, path, Path(tmp) / "out", store_currency="USD",
                                   location="Shop location")
            files = {p.name: read_rows(p) for p in written}
        self.assertEqual(headers, header)
        self.assertEqual(parsed[0]["Variant Price"], "885.95")
        self.assertEqual(parsed[0]["Variant ID"], "68237159366948")
        self.assertEqual(files["Products_update_price_and_stock.csv"][0],
                         {"Handle": "minimal", "Command": "UPDATE", "Variant ID": "68237159366948",
                          "Variant SKU": "SNOW-001", "Variant Command": "UPDATE",
                          "Variant Price": "899.95", "Inventory Available: Shop location": "45"})
        self.assertNotIn("Variant Inventory Qty", files["Products_update_stock_only.csv"][0])
        self.assertEqual(files["Products_update_stock_only.csv"][0]["Inventory Available: Shop location"], "12")

    def test_unknown_location_is_an_error(self):
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "Export.xlsx"
            write_xlsx(path, ["Handle", "Variant SKU", "Variant Price"], [["h", "A-1", "1.00"]])
            feed = Path(tmp) / "feed.csv"
            feed.write_text("SKU;Price;Stock\nA-1;1,00;3\n", encoding="utf-8")
            with self.assertRaises(ValueError):
                run(feed, path, Path(tmp) / "out", location="Nowhere")


class BuildPlanTest(unittest.TestCase):
    def test_missing_stock_leaves_stock_untouched(self):
        plan = build_plan([item("A-1", price="11,00", stock="")], [variant("A-1", stock="5")])
        self.assertEqual(len(plan.updates), 1)
        self.assertIsNone(plan.updates[0].new_stock)
        self.assertEqual(plan.updates[0].fields, ("price",))
        self.assertEqual(len(plan.unusable_stock), 1)

    def test_real_zero_stock_is_applied(self):
        plan = build_plan([item("A-1", stock="0")], [variant("A-1", stock="5")])
        self.assertEqual(plan.updates[0].new_stock, 0)

    def test_price_change_above_threshold_is_held_but_stock_still_updates(self):
        plan = build_plan([item("A-1", price="13,00", stock="9")], [variant("A-1", price="10.00")])
        self.assertEqual(len(plan.price_review), 1)
        self.assertEqual(plan.price_review[0][2], Decimal("0.3"))
        self.assertIsNone(plan.updates[0].new_price)
        self.assertEqual(plan.updates[0].new_stock, 9)

    def test_price_change_at_threshold_is_allowed(self):
        plan = build_plan([item("A-1", price="12,00")], [variant("A-1", price="10.00")])
        self.assertEqual(plan.updates[0].new_price, Decimal("12.00"))
        self.assertEqual(plan.price_review, [])

    def test_price_drop_above_threshold_is_held(self):
        plan = build_plan([item("A-1", price="7,00")], [variant("A-1", price="10.00")])
        self.assertEqual(len(plan.price_review), 1)
        self.assertEqual(plan.updates, [])

    def test_conflicting_feed_duplicates_are_held(self):
        plan = build_plan([item("A-1", stock="10", line=2), item("a_1", stock="8", line=3)],
                          [variant("A-1")])
        self.assertIn("A1", plan.feed_conflicts)
        self.assertEqual(plan.updates, [])
        self.assertEqual(plan.missing_from_feed, [])

    def test_identical_feed_duplicates_are_merged(self):
        plan = build_plan([item("A-1", stock="9", line=2), item("A-1", stock="9", line=3)],
                          [variant("A-1")])
        self.assertIn("A1", plan.feed_merged)
        self.assertEqual(len(plan.updates), 1)

    def test_sku_on_several_store_variants_is_held(self):
        plan = build_plan([item("A-1", stock="9")],
                          [variant("A-1", variant_id="1"), variant("A-1", variant_id="2")])
        self.assertIn("A1", plan.store_duplicates)
        self.assertEqual(plan.updates, [])

    def test_unmatched_missing_and_skipped(self):
        plan = build_plan([item("NEW-1"), item("", line=7)], [variant("OLD-1")])
        self.assertEqual([i.sku for i in plan.unmatched], ["NEW1"])
        self.assertEqual([v.sku for v in plan.missing_from_feed], ["OLD1"])
        self.assertEqual(plan.skipped_rows, [(7, "empty SKU")])
        self.assertEqual(plan.updates, [])

    def test_unchanged(self):
        plan = build_plan([item("A-1", price="10,00", stock="5")], [variant("A-1")])
        self.assertEqual(len(plan.unchanged), 1)
        self.assertEqual(plan.updates, [])


class EndToEndTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.out = Path(self.tmp.name)
        self.plan, self.written, self.report = run(SUPPLIER, STORE, self.out)
        self.rows = {p.name: read_rows(p) for p in self.written}

    def tearDown(self):
        self.tmp.cleanup()

    def skus_in(self, name):
        return sorted(r["Variant SKU"] for r in self.rows.get(name, []))

    def test_files_split_by_changed_fields(self):
        self.assertEqual(self.skus_in("Products_update_price_and_stock.csv"),
                         ["000456", "AB-1001", "AB-1003", "AB-1008"])
        self.assertEqual(self.skus_in("Products_update_price_only.csv"), ["AB-1004", "AB-1009", "AB-1011"])
        self.assertEqual(self.skus_in("Products_update_stock_only.csv"),
                         ["AB-1006", "AB-1010", "AB-1012", "AB-1014", "AB-1016"])

    def test_usd_price_in_sample_feed_is_held_not_converted(self):
        self.assertEqual([v.raw_sku for v, _, _ in self.plan.currency_mismatch], ["AB-1012"])
        row = next(r for r in self.rows["Products_update_stock_only.csv"] if r["Variant SKU"] == "AB-1012")
        self.assertEqual(row["Variant Inventory Qty"], "1200")
        self.assertNotIn("Variant Price", row)

    def test_file_names_contain_products_for_matrixify_detection(self):
        self.assertTrue(all("Products" in p.name for p in self.written))

    def test_price_only_file_has_no_inventory_column(self):
        header = next(iter(self.rows["Products_update_price_only.csv"])).keys()
        self.assertNotIn("Variant Inventory Qty", header)

    def test_only_genuine_zero_stock_is_written_as_zero(self):
        zeros = [r["Variant SKU"] for rows in self.rows.values() for r in rows
                 if r.get("Variant Inventory Qty") == "0"]
        self.assertEqual(zeros, ["AB-1006"])

    def test_held_and_ambiguous_skus_are_never_imported(self):
        imported = {r["Variant SKU"] for rows in self.rows.values() for r in rows}
        for sku in ("AB-1005", "AB-1018", "AB-1007", "AB-1017", "AB-1015", "AB-1020", "AB-1002"):
            self.assertNotIn(sku, imported)

    def test_uses_store_sku_and_ids_with_update_commands(self):
        row = next(r for r in self.rows["Products_update_price_and_stock.csv"] if r["Variant SKU"] == "AB-1001")
        self.assertEqual(row, {"Handle": "ceramic-mug-350ml", "Command": "UPDATE",
                               "Variant ID": "44120000000001", "Variant SKU": "AB-1001",
                               "Variant Command": "UPDATE", "Variant Price": "24.90",
                               "Variant Inventory Qty": "15"})

    def test_report_mentions_every_problem(self):
        text = self.report.read_text(encoding="utf-8")
        for expected in ("AB-1005", "+40.6%", "AB-1018", "AB-1007", "AB-1017",
                         "XY-2001", "AB-1020", "`N/A`", "`price on request`", "empty SKU",
                         "`$18`", "Warning:** the feed contains prices in USD"):
            self.assertIn(expected, text)


if __name__ == "__main__":
    unittest.main()
