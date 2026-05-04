from flask import Flask, request, send_file
import io
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.utils import get_column_letter

app = Flask(__name__)

DARK_BLUE  = "1A3C5E"
MID_BLUE   = "2E6DA4"
LIGHT_BLUE = "D6E4F0"
YELLOW     = "FFF2CC"
WHITE      = "FFFFFF"
GRAY       = "ECECEC"
GREEN      = "E2EFDA"

def tb(color="BBBBBB"):
    s = Side(style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def sc(ws, cell, value=None, bold=False, bg=None, fg="000000",
       sz=9, ha="left", va="center", wrap=False, locked=True,
       italic=False, fmt=None):
    c = ws[cell] if isinstance(cell, str) else cell
    if value is not None:
        c.value = value
    c.font = Font(name="Arial", bold=bold, size=sz, color=fg, italic=italic)
    if bg:
        c.fill = PatternFill("solid", start_color=bg)
    c.alignment = Alignment(horizontal=ha, vertical=va, wrap_text=wrap)
    c.border = tb()
    c.protection = Protection(locked=locked)
    if fmt:
        c.number_format = fmt
    return c

def build_iungo_xlsx(order):
    wb = Workbook()
    ws = wb.active
    ws.title = "ORDINE"

    # Column widths
    for i, w in enumerate([6,18,18,32,6,20,16,10,14,10,16,4], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Row heights
    ws.row_dimensions[1].height = 32
    ws.row_dimensions[2].height = 16
    ws.row_dimensions[3].height = 8
    for r in range(4, 14):
        ws.row_dimensions[r].height = 22
    ws.row_dimensions[13].height = 8
    ws.row_dimensions[14].height = 40

    doc_number    = order.get("docNumber", "ORD-AUTO")
    issue_date    = order.get("issueDate", "")
    delivery_date = order.get("deliveryDate", "")
    destination   = order.get("destination", "")
    customer_name = order.get("customerName", "")
    customer_vat  = order.get("customerVat", "")
    supplier_name = order.get("supplierName", "REGALIDEA S.R.L.")
    supplier_vat  = order.get("supplierVat", "IT00000000000")
    notes         = order.get("notes", "")
    order_number  = order.get("orderNumber", "")
    items         = order.get("items", [])

    # ── ROW 1: Title ──────────────────────────────────────────────────
    ws.merge_cells("A1:L1")
    sc(ws, "A1", "REGALIDEA S.R.L. — ORDINE FORNITORE",
       bold=True, sz=14, bg=DARK_BLUE, fg=WHITE, ha="center")

    # ── ROW 2: Subtitle ───────────────────────────────────────────────
    ws.merge_cells("A2:L2")
    sc(ws, "A2",
       "I campi GIALLI sono da compilare dal fornitore. Non modificare la struttura.",
       italic=True, sz=8, bg=MID_BLUE, fg=WHITE, ha="center")

    # ── ROWS 4-7: Header block ────────────────────────────────────────
    header_left = [
        (4, "N. DOCUMENTO:",   doc_number,    GRAY,   True),
        (5, "DATA EMISSIONE:", issue_date,    GRAY,   True),
        (6, "DATA CONSEGNA:",  delivery_date, YELLOW, False),
        (7, "DESTINAZIONE:",   destination,   YELLOW, False),
    ]
    header_right = [
        (4, "NOME CLIENTE:",    customer_name, YELLOW, False),
        (5, "P.IVA CLIENTE:",   customer_vat,  YELLOW, False),
        (6, "NOME FORNITORE:",  supplier_name, GRAY,   True),
        (7, "P.IVA FORNITORE:", supplier_vat,  GRAY,   True),
    ]

    for row, label, value, bg, locked in header_left:
        ws.merge_cells(f"A{row}:B{row}")
        sc(ws, f"A{row}", label, bold=True, fg=DARK_BLUE, bg=GRAY, ha="right")
        ws.merge_cells(f"C{row}:F{row}")
        sc(ws, f"C{row}", value, bg=bg, locked=locked, bold=not locked)

    for row, label, value, bg, locked in header_right:
        ws.merge_cells(f"G{row}:H{row}")
        sc(ws, f"G{row}", label, bold=True, fg=DARK_BLUE, bg=GRAY, ha="right")
        ws.merge_cells(f"I{row}:L{row}")
        sc(ws, f"I{row}", value, bg=bg, locked=locked, bold=not locked)

    # ── ROW 9: Totals ─────────────────────────────────────────────────
    last_item_row = 15 + len(items) - 1
    ws.row_dimensions[9].height = 22
    ws.merge_cells("A9:B9")
    sc(ws, "A9", "TOTALE QUANTITÀ:", bold=True, fg=DARK_BLUE, bg=GRAY, ha="right")
    ws.merge_cells("C9:F9")
    sc(ws, "C9", f"=SUM(H15:H{last_item_row})",
       bold=True, fg=DARK_BLUE, bg=LIGHT_BLUE, ha="center", fmt="#,##0")
    ws.merge_cells("G9:H9")
    sc(ws, "G9", "TOTALE IMPORTO:", bold=True, fg=DARK_BLUE, bg=GRAY, ha="right")
    ws.merge_cells("I9:L9")
    sc(ws, "I9", f"=SUM(K15:K{last_item_row})",
       bold=True, fg=DARK_BLUE, bg=LIGHT_BLUE, ha="center", fmt="€#,##0.00")

    # ── ROW 10: Notes ─────────────────────────────────────────────────
    ws.row_dimensions[10].height = 22
    ws.merge_cells("A10:B10")
    sc(ws, "A10", "NOTE:", bold=True, fg=DARK_BLUE, bg=GRAY, ha="right")
    ws.merge_cells("C10:L10")
    sc(ws, "C10", notes, bg=YELLOW, locked=False)

    # ── ROW 11: Order number ──────────────────────────────────────────
    ws.row_dimensions[11].height = 22
    ws.merge_cells("A11:B11")
    sc(ws, "A11", "N. ORDINE CLIENTE:", bold=True, fg=DARK_BLUE, bg=GRAY, ha="right")
    ws.merge_cells("C11:F11")
    sc(ws, "C11", order_number, bg=YELLOW, locked=False)

    # ── ROW 14: Column headers ────────────────────────────────────────
    headers = [
        "#", "CODICE\nARTICOLO\n(Loro)", "CODICE\nFORNITORE\n(Loro)",
        "DESCRIZIONE", "U.M.", "NOME CLIENTE\n(Riga)", "P.IVA CLIENTE\n(Riga)",
        "QTÀ", "PREZZO\nLORDO €", "SCONTO\n%", "IMPORTO\nNETTO €", ""
    ]
    for col, h in enumerate(headers, 1):
        c = ws.cell(row=14, column=col, value=h)
        c.font = Font(name="Arial", bold=True, size=8, color=WHITE)
        c.fill = PatternFill("solid", start_color=DARK_BLUE)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = tb(WHITE)
        c.protection = Protection(locked=True)

    # ── ROWS 15+: Line items ──────────────────────────────────────────
    for idx, item in enumerate(items):
        row = 15 + idx
        ws.row_dimensions[row].height = 18
        is_even = idx % 2 == 0
        amt_bg = LIGHT_BLUE if is_even else GREEN

        cols = [
            (1,  idx+1,                      GRAY,   "center", True,  False, "#,##0"),
            (2,  item.get("itemCode",""),     GRAY,   "center", True,  False, None),
            (3,  item.get("supplierCode",""), GRAY,   "center", True,  False, None),
            (4,  item.get("description",""),  GRAY,   "left",   True,  False, None),
            (5,  item.get("um","PZ"),         GRAY,   "center", True,  False, None),
            (6,  item.get("custName",""),     YELLOW, "left",   False, False, None),
            (7,  item.get("custVat",""),      YELLOW, "left",   False, False, None),
            (8,  item.get("qty",0),           YELLOW, "center", False, True,  "#,##0"),
            (9,  item.get("grossPrice",0),    GRAY,   "right",  True,  False, "€#,##0.00"),
            (10, item.get("discount",0),      YELLOW, "center", False, False, "0.00%"),
            (11, f"=IF(H{row}=0,\"\",IF(J{row}=0,H{row}*I{row},H{row}*I{row}*(1-J{row})))",
                 amt_bg, "right", True, False, "€#,##0.00"),
        ]
        for col, val, bg, ha, locked, bold, fmt in cols:
            c = ws.cell(row=row, column=col, value=val)
            c.font = Font(name="Arial", size=9, bold=bold,
                          color=DARK_BLUE if col == 11 else "000000")
            c.fill = PatternFill("solid", start_color=bg)
            c.alignment = Alignment(horizontal=ha, vertical="center")
            c.border = tb()
            c.protection = Protection(locked=locked)
            if fmt:
                c.number_format = fmt

    # Empty filler rows
    for i in range(len(items), 30):
        row = 15 + i
        ws.row_dimensions[row].height = 18
        for col in range(1, 12):
            is_even = i % 2 == 0
            amt_bg = LIGHT_BLUE if is_even else GREEN
            bg = YELLOW if col in [6,7,8,10] else (amt_bg if col == 11 else GRAY)
            c = ws.cell(row=row, column=col)
            c.fill = PatternFill("solid", start_color=bg)
            c.border = tb()
            c.protection = Protection(locked=col not in [6,7,8,10])

    # ── Instructions sheet ────────────────────────────────────────────
    ws2 = wb.create_sheet("ISTRUZIONI")
    ws2.column_dimensions["A"].width = 80
    for row, text, bold, sz, bg, fg in [
        (1, "ISTRUZIONI PER LA COMPILAZIONE", True, 13, DARK_BLUE, WHITE),
        (3, "LEGENDA COLORI", True, 10, MID_BLUE, WHITE),
        (4, "🟡 GIALLO = da compilare", False, 9, None, "000000"),
        (5, "⬜ GRIGIO = pre-compilato (NON modificare)", False, 9, None, "555555"),
        (6, "🔵 BLU/VERDE = calcolato automaticamente", False, 9, None, "1A3C5E"),
        (8, "INVIO ORDINE", True, 10, MID_BLUE, WHITE),
        (9, "• Inviare a: ordini@regalidea.it", False, 9, None, "000000"),
        (10, "• NON modificare la struttura del file", False, 9, None, "CC0000"),
    ]:
        ws2.row_dimensions[row].height = 22
        c = ws2.cell(row=row, column=1, value=text)
        c.font = Font(name="Arial", bold=bold, size=sz, color=fg)
        if bg:
            c.fill = PatternFill("solid", start_color=bg)
        c.alignment = Alignment(horizontal="left", vertical="center")
    ws2.protection.sheet = True
    ws2.protection.password = "regalidea2025"

    # ── Protect main sheet ────────────────────────────────────────────
    ws.protection.sheet = True
    ws.protection.password = "regalidea2025"
    ws.protection.selectLockedCells = False
    ws.protection.selectUnlockedCells = False
    ws.print_area = "A1:L50"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1

    # Save to buffer
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf


@app.route("/generate", methods=["POST"])
def generate():
    data = request.get_json()
    if not data:
        return {"error": "No JSON body"}, 400

    order = data.get("order", data)

    # Map from n8n order format to our format
    isPAC       = "PAC 2000" in (order.get("customer") or "")
    isGabrielli = "GABRIELLI" in (order.get("customer") or "")
    isDrink     = "DRINK" in (order.get("customer") or "")

    delivery_date = order.get("deliveryWindow") or order.get("deliveryDate") or ""
    destination   = order.get("deliveryLocation") or order.get("luogoConsegna") or order.get("pointOfSale") or ""
    customer_name = order.get("pointOfSale") or order.get("customer") or ""
    from datetime import datetime
    issue_date = datetime.now().strftime("%d/%m/%Y")
    import time
    doc_number = f"ORD-{datetime.now().year}-{str(int(time.time()))[-6:]}"

    items = []
    for i in order.get("items", []):
        if isPAC:
            items.append({
                "itemCode":     i.get("codice", ""),
                "supplierCode": i.get("codice", ""),
                "description":  i.get("descrizione", ""),
                "um":           "PZ",
                "custName":     customer_name,
                "custVat":      "",
                "qty":          i.get("quantita", 0),
                "grossPrice":   i.get("fascia", 0),
                "discount":     0,
            })
        elif isGabrielli:
            items.append({
                "itemCode":     "",
                "supplierCode": "",
                "description":  i.get("descrizione", ""),
                "um":           "PZ",
                "custName":     "GABRIELLI",
                "custVat":      "",
                "qty":          i.get("quantitaGabrielli") or i.get("quantita", 0),
                "grossPrice":   0,
                "discount":     0,
            })
        elif isDrink:
            items.append({
                "itemCode":     i.get("codice", ""),
                "supplierCode": i.get("codice", ""),
                "description":  i.get("descrizione", ""),
                "um":           "PZ",
                "custName":     order.get("luogoConsegna", ""),
                "custVat":      "",
                "qty":          i.get("quantita", 0),
                "grossPrice":   0,
                "discount":     0,
            })
        else:
            items.append({
                "itemCode":     i.get("codice", ""),
                "supplierCode": i.get("codice", ""),
                "description":  i.get("descrizione") or i.get("description", ""),
                "um":           i.get("um", "PZ"),
                "custName":     customer_name,
                "custVat":      "",
                "qty":          i.get("quantita") or i.get("quantity", 0),
                "grossPrice":   i.get("prezzo") or i.get("fascia", 0),
                "discount":     i.get("sconto", 0),
            })

    order_data = {
        "docNumber":    doc_number,
        "issueDate":    issue_date,
        "deliveryDate": delivery_date,
        "destination":  destination,
        "customerName": customer_name,
        "customerVat":  order.get("customerVat", ""),
        "supplierName": "REGALIDEA S.R.L.",
        "supplierVat":  "IT00000000000",
        "notes":        order.get("notes") or order.get("pagamento", ""),
        "orderNumber":  order.get("orderNumber", ""),
        "items":        items,
    }

    buf = build_iungo_xlsx(order_data)

    safe_customer = (order.get("customer") or "ORDER").replace(" ", "_")
    safe_date = delivery_date.replace("/", "-")
    filename = f"IUNGO_{safe_customer}_{safe_date}.xlsx"

    return send_file(
        buf,
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        as_attachment=True,
        download_name=filename
    )


@app.route("/health", methods=["GET"])
def health():
    return {"status": "ok", "service": "IUNGO Excel Generator"}


if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port)
