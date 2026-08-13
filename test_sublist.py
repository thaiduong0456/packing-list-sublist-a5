from io import BytesIO
from decimal import Decimal

import pymupdf
from openpyxl import Workbook

from sublist import Carton, Item, build_sublist_pdf, decimal_text, excel_text, paginate_cartons, parse_packing_list


def _sample_file():
    wb = Workbook()
    ws = wb.active
    ws.title = "Packing List"
    ws.append(["OR No.", "Ref No.", "SKU#", "BarCode/UPC", "Quantity", "Carton#", "Packaging code", "Weight (KG)"])
    ws.append(["OR1196", "ia32674", "SKU-01", "4894961083099", 1, "1/2", "PKG0001", 0.45])
    ws.append(["OR1196", "ia32674", "SKU-02", "4894961081064", 1, "2/2", "PKG0002", 0.6])
    ws.append(["OR1196", "ia32674", "SKU-03", "4894961081248", 1, None, None, None])
    ws.append(["OR1196", "ia32674", "SKU-04", "4894961081903", 1, None, None, None])
    output = BytesIO()
    wb.save(output)
    output.seek(0)
    return output


def test_sample_packing_list_creates_two_a5_pages():
    cartons = parse_packing_list(_sample_file())
    assert [carton.carton for carton in cartons] == ["1/2", "2/2"]
    assert [len(carton.items) for carton in cartons] == [1, 3]
    assert cartons[1].items[-1].ean == "4894961081903"
    assert sum(int(item.qty) for item in cartons[1].items) == 3

    doc = pymupdf.open(stream=build_sublist_pdf(cartons), filetype="pdf")
    assert doc.page_count == 2
    assert round(doc[0].rect.width, 1) == 419.5
    assert round(doc[0].rect.height, 1) == 595.3
    assert "PKG0002" in doc[1].get_text()


def test_integer_zeroes_are_never_removed():
    assert excel_text(10.0) == "10"
    assert excel_text(20.0) == "20"
    assert decimal_text(Decimal("100")) == "100"
    assert decimal_text(Decimal("250")) == "250"
    assert decimal_text(Decimal("2.500")) == "2.5"


def test_long_carton_is_complete_and_total_is_250():
    quantities = [12, 12, 15, 15, 15, 8, 15, 8, 15, 10, 6, 6, 10, 12, 5, 10, 10, 8, 12, 12, 12, 12, 10]
    assert sum(quantities) == 250
    items = [Item(f"SKU-{n:02d}", f"4890000000{n:03d}", str(qty)) for n, qty in enumerate(quantities, 1)]
    carton = Carton("15/48", "OR1173", "po38535", "24.95", "PGKEC7HHGPOL5740002", items)
    doc = pymupdf.open(stream=build_sublist_pdf([carton]), filetype="pdf")
    text = doc[0].get_text()
    assert doc.page_count == 1
    assert "SKU-01" in text and "SKU-23" in text
    assert text.strip().endswith("250")


def test_long_reference_wraps_without_losing_content():
    long_ref = "po38534 / po38535 / po38536 / po38537 / po38538"
    carton = Carton("12/48", "OR1174", long_ref, "24.65", "PGKEC53HJEHR8470003", [
        Item("TP-WST-R06-BKR-00", "4895227934032", "10")
    ])
    doc = pymupdf.open(stream=build_sublist_pdf([carton]), filetype="pdf")
    text = doc[0].get_text()
    for ref in ("po38534", "po38535", "po38536", "po38537", "po38538"):
        assert ref in text


def _carton_with_rows(count: int) -> Carton:
    return Carton("15/48", "OR1173", "po38535", "24.95", "PKG001", [
        Item(f"SKU-{number:02d}", f"4890000000{number:03d}", "1")
        for number in range(1, count + 1)
    ])


def test_pagination_boundaries_and_carton_labels():
    expected = {25: 1, 26: 2, 50: 2, 51: 3}
    for row_count, page_count in expected.items():
        pages = paginate_cartons([_carton_with_rows(row_count)])
        assert len(pages) == page_count
        assert all(len(page.items) <= 25 for page in pages)
        doc = pymupdf.open(stream=build_sublist_pdf([_carton_with_rows(row_count)]), filetype="pdf")
        assert doc.page_count == page_count

    pages = paginate_cartons([_carton_with_rows(51)])
    assert [page.carton for page in pages] == ["15-1/48", "15-2/48", "15-3/48"]
    assert [len(page.items) for page in pages] == [25, 25, 1]
    assert [sum(int(item.qty) for item in page.items) for page in pages] == [25, 25, 1]
    assert [page.show_total for page in pages] == [False, False, True]
    assert [page.total_qty for page in pages] == ["51", "51", "51"]


def test_split_carton_shows_full_total_on_last_page_only():
    carton = _carton_with_rows(26)
    doc = pymupdf.open(stream=build_sublist_pdf([carton]), filetype="pdf")
    assert doc.page_count == 2
    first_text = doc[0].get_text().strip().splitlines()
    last_text = doc[1].get_text().strip().splitlines()
    assert first_text[-1] == "1"  # Last item's QTY, no carton total below the table.
    assert last_text[-1] == "26"  # Full-carton total, not the last-page subtotal of 1.
