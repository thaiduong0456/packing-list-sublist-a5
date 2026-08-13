from io import BytesIO

import pymupdf
from openpyxl import Workbook

from sublist import build_sublist_pdf, parse_packing_list


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
