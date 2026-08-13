from io import BytesIO
from decimal import Decimal

import pymupdf
from openpyxl import Workbook

from sublist import Carton, Item, build_sublist_pdf, decimal_text, excel_text, parse_packing_list


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
