from flask import Flask, request, send_file
import io
import datetime
import time
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, Protection
from openpyxl.utils import get_column_letter
from fpdf import FPDF # Pure Python PDF library (no Railway errors!)

app = Flask(__name__)

# ==========================================
# EXCEL STYLING & HELPERS (Your exact code)
# ==========================================
DARK_BLUE  = "203764"   
MID_BLUE   = "305496"   
LIGHT_BLUE = "D9E1F2"   
YELLOW     = "FFF2CC"   
WHITE      = "FFFFFF"
GRAY       = "F2F2F2"   
BORDER_COL = "B7B7B7"

def get_border():
    s = Side(style="thin", color=BORDER_COL)
    return Border(left=s, right=s, top=s, bottom=s)

def style_cell(ws, cell, value=None, bold=False, bg=None, fg="000000",
               sz=9, ha="left", va="center", wrap=False, locked=True,
               italic=False, fmt=None):
    c = ws[cell] if isinstance(cell, str) else cell
    if value is not None:
        c.value = value
    c.font = Font(name="Arial", bold=bold, size=sz, color=fg, italic=italic)
    if bg:
        c.fill = PatternFill("solid", start_color=bg)
    c.alignment = Alignment(horizontal=ha, vertical=va, wrap_text=wrap)
    c.border = get_border()
    c.protection = Protection(locked=locked)
    if fmt:
        c.number_format = fmt
    return c

def build_iungo_xlsx(order):
    wb = Workbook()
    ws = wb.active
    ws.title = "ORDINE"

    # Column Widths
    widths = [6, 20, 20, 40, 8, 12, 15, 12, 18]
    for i, w in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    # Headers
    ws.merge_cells("A1:I1")
    style_cell(ws, "A1", "REGALIDEA S.R.L. — ORDINE FORNITORE",
               bold=True, sz=14, bg=DARK_BLUE, fg=WHITE, ha="center")

    ws.merge_cells("A2:I2")
    style_cell(ws, "A2", "I campi GIALLI sono da compilare dal fornitore. Non modificare la struttura.",
               italic=True, sz=9, bg=MID_BLUE, fg=WHITE, ha="center")

    # Info Block
    info_left = [
        (4, "N. DOCUMENTO:",    order.get("docNumber", ""), GRAY,   True),
        (5, "DATA EMISSIONE:",  order.get("issueDate", ""), GRAY,   True),
        (6, "DATA CONSEGNA:",   order.get("deliveryDate", ""), YELLOW, False),
        (7, "DESTINAZIONE:",    order.get("destination", ""), YELLOW, False),
    ]
    info_right = [
        (4, "NOME CLIENTE:",    order.get("customerName", ""), YELLOW, False),
        (5, "P.IVA CLIENTE:",   order.get("customerVat", ""),  YELLOW, False),
        (6, "NOME FORNITORE:",  "REGALIDEA S.R.L.", GRAY, True),
        (7, "P.IVA FORNITORE:", "IT00000000000", GRAY, True),
    ]

    for r, label, val, bg, lck in info_left:
        ws.merge_cells(f"A{r}:B{r}")
        style_cell(ws, f"A{r}", label, bold=True, fg=DARK_BLUE, ha="right")
        ws.merge_cells(f"C{r}:E{r}")
        style_cell(ws, f"C{r}", val, bg=bg, locked=lck)

    for r, label, val, bg, lck in info_right:
        ws.cell(row=r, column=7, value=label).font = Font(bold=True, color=DARK_BLUE)
        ws.cell(row=r, column=7).alignment = Alignment(horizontal="right")
        ws.merge_cells(start_row=r, start_column=8, end_row=r, end_column=9)
        style_cell(ws, ws.cell(row=r, column=8), val, bg=bg, locked=lck)

    # Totals
    style_cell(ws, "B9", "TOTALE QUANTITÀ:", bold=True, fg=DARK_BLUE, ha="right")
    ws.merge_cells("C9:E9")
    style_cell(ws, "C9", "=SUM(F15:F500)", bg=LIGHT_BLUE, bold=True, ha="center", fmt="#,##0")

    style_cell(ws, "G9", "TOTALE IMPORTO:", bold=True, fg=DARK_BLUE, ha="right")
    ws.merge_cells("H9:I9")
    style_cell(ws, "H9", "=SUM(I15:I500)", bg=LIGHT_BLUE, bold=True, ha="center", fmt='€ #,##0.00')

    # Table Headers
    headers = ["#", "CODICE ARTICOLO", "CODICE FORNITORE", "DESCRIZIONE", "U.M.", "QTÀ", "PREZZO LORDO €", "SCONTO %", "IMPORTO NETTO €"]
    for col, text in enumerate(headers, 1):
        c = ws.cell(row=14, column=col, value=text)
        style_cell(ws, c, bold=True, bg=DARK_BLUE, fg=WHITE, ha="center", sz=8)

    # Line Items
    items = order.get("items", [])
    for idx, item in enumerate(items):
        row = 15 + idx
        netto_formula = f"=F{row}*G{row}*(1-H{row}/100)"
        
        row_data = [
            (1, idx+1, GRAY, True),                        
            (2, item.get("itemCode",""), YELLOW, False),   
            (3, item.get("supplierCode",""), YELLOW, False),
            (4, item.get("description",""), YELLOW, False),
            (5, item.get("um","PZ"), YELLOW, False),       
            (6, item.get("qty",0), YELLOW, False),         
            (7, item.get("grossPrice",0), YELLOW, False),  
            (8, item.get("discount",0), YELLOW, False),    
            (9, netto_formula, LIGHT_BLUE, True)           
        ]

        for col, val, bg, lck in row_data:
            fmt = '€ #,##0.00' if col in [7, 9] else ("#,##0" if col == 6 else None)
            style_cell(ws, ws.cell(row=row, column=col), val, bg=bg, locked=lck, fmt=fmt, ha="center" if col != 4 else "left")

    # Empty filler rows
    for r in range(15 + len(items), 101):
        for c in range(1, 10):
            bg = YELLOW if 2 <= c <= 8 else (GRAY if c == 1 else LIGHT_BLUE)
            style_cell(ws, ws.cell(row=r, column=c), bg=bg, locked=not (2 <= c <= 8))
            if c == 9: 
                ws.cell(row=r, column=9).value = f"=F{r}*G{r}*(1-H{r}/100)"
                ws.cell(row=r, column=9).number_format = '€ #,##0.00'

    # Protection
    ws.protection.sheet = True
    ws.protection.password = "regalidea2026"
    
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

# ==========================================
# ROUTE 1: THE EXCEL GENERATOR (Your exact code)
# ==========================================
@app.route("/generate", methods=["POST"])
def generate():
    data = request.get_json()
    if not data: return {"error": "No JSON body"}, 400
    
    order = data.get("order", data)
    now = datetime.datetime.now()
    doc_num = f"ORD-{now.year}-{str(int(time.time()))[-6:]}"
    issue_date = now.strftime("%d/%m/%Y")

    items = []
    for i in order.get("items", []):
        item_code = i.get("itemCode") or i.get("codice") or ""
        supplier_code = i.get("supplierCode", "") 
        
        items.append({
            "itemCode": item_code,
            "supplierCode": supplier_code,
            "description": i.get("descrizione") or i.get("description", ""),
            "um": i.get("um", "PZ"),
            "qty": i.get("quantitaGabrielli") or i.get("quantita") or i.get("quantity", 0),
            "grossPrice": i.get("prezzo") or i.get("fascia", 0),
            "discount": i.get("sconto", 0),
        })

    order_data = {
        "docNumber": doc_num,
        "issueDate": issue_date,
        "deliveryDate": order.get("deliveryDate") or order.get("deliveryWindow") or "",
        "destination": order.get("luogoConsegna") or order.get("destination") or "",
        "customerName": order.get("customer") or order.get("pointOfSale") or "",
        "customerVat": order.get("customerVat", ""),
        "items": items
    }

    buf = build_iungo_xlsx(order_data)
    filename = f"IUNGO_ORDER_{order_data['customerName'].replace(' ', '_')}.xlsx"

    return send_file(buf, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                     as_attachment=True, download_name=filename)


# ==========================================
# PDF STYLING & CLASS
# ==========================================
class RegalideaPDF(FPDF):
    def header(self):
        self.set_fill_color(32, 55, 100) # #203764
        self.set_text_color(255, 255, 255)
        self.set_font("helvetica", "B", 14)
        self.cell(0, 10, "REGALIDEA S.R.L. - ORDINE FORNITORE", align="C", fill=True, ln=True)
        self.set_font("helvetica", "I", 10)
        self.cell(0, 6, "I campi GIALLI sono da compilare dal fornitore.", align="C", fill=True, ln=True)
        self.ln(5)

# ==========================================
# ROUTE 2: THE PDF GENERATOR
# ==========================================
@app.route("/generate-pdf", methods=["POST"])
def generate_pdf():
    data = request.get_json()
    if not data: return {"error": "No JSON body"}, 400
    
    order = data.get("order", data)

    pdf = RegalideaPDF(orientation="L", unit="mm", format="A4")
    pdf.add_page()

    pdf.set_font("helvetica", "B", 9)
    blue_text = (32, 55, 100)
    yellow_fill = (255, 242, 204)
    
    now = datetime.datetime.now()
    doc_num = f"ORD-{now.year}-{str(int(time.time()))[-6:]}"
    customer_name = str(order.get("customer") or order.get("pointOfSale") or "")

    # ROW 1
    pdf.set_text_color(*blue_text)
    pdf.cell(35, 6, "N. DOCUMENTO:", border=1, align="R")
    pdf.set_text_color(0, 0, 0); pdf.set_font("helvetica", "", 9)
    pdf.cell(80, 6, doc_num, border=1)

    pdf.set_text_color(*blue_text); pdf.set_font("helvetica", "B", 9)
    pdf.cell(35, 6, "NOME CLIENTE:", border=1, align="R")
    pdf.set_fill_color(*yellow_fill)
    pdf.set_text_color(0, 0, 0); pdf.set_font("helvetica", "", 9)
    pdf.cell(127, 6, customer_name, border=1, fill=True, ln=True)

    # ROW 2
    del_date = str(order.get("deliveryDate") or order.get("deliveryWindow") or "")[:10]
    dest = str(order.get("luogoConsegna") or order.get("destination") or "")[:65]

    pdf.set_text_color(*blue_text); pdf.set_font("helvetica", "B", 9)
    pdf.cell(35, 6, "DATA CONSEGNA:", border=1, align="R")
    pdf.set_text_color(0, 0, 0); pdf.set_font("helvetica", "", 9)
    pdf.cell(80, 6, del_date, border=1)

    pdf.set_text_color(*blue_text); pdf.set_font("helvetica", "B", 9)
    pdf.cell(35, 6, "DESTINAZIONE:", border=1, align="R")
    pdf.set_fill_color(*yellow_fill)
    pdf.set_text_color(0, 0, 0); pdf.set_font("helvetica", "", 9)
    pdf.cell(127, 6, dest, border=1, fill=True, ln=True)
    pdf.ln(5)

    # TABLE HEADERS
    pdf.set_fill_color(32, 55, 100)
    pdf.set_text_color(255, 255, 255)
    pdf.set_font("helvetica", "B", 8)
    cols = [("CODICE", 25), ("DESCRIZIONE", 130), ("U.M.", 10), ("QTÀ", 15), 
            ("PREZZO €", 25), ("SCONTO %", 22), ("NETTO €", 50)]

    for col_name, width in cols:
        pdf.cell(width, 7, col_name, border=1, align="C", fill=True)
    pdf.ln()

    # TABLE ROWS
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("helvetica", "", 8)

    items = order.get("items", [])
    total_qty = 0
    total_amount = 0

    for i in items:
        # Extract fields using your exact logic
        code = str(i.get("itemCode") or i.get("codice") or "")
        desc = str(i.get("descrizione") or i.get("description", ""))[:75]
        qty = float(i.get("quantitaGabrielli") or i.get("quantita") or i.get("quantity") or 0)
        price = float(i.get("prezzo") or i.get("fascia") or i.get("grossPrice") or 0)
        net = qty * price

        total_qty += qty
        total_amount += net

        pdf.set_fill_color(*yellow_fill)
        pdf.cell(25, 6, code, border=1, align="C", fill=True)
        pdf.cell(130, 6, desc, border=1, fill=True)
        pdf.cell(10, 6, "PZ", border=1, align="C")
        pdf.cell(15, 6, str(int(qty)), border=1, align="C")
        pdf.cell(25, 6, f"€ {price:.2f}", border=1, align="R", fill=True)
        pdf.cell(22, 6, "0", border=1, align="C", fill=True)
        pdf.cell(50, 6, f"€ {net:.2f}", border=1, align="R", fill=True)
        pdf.ln()

    # FOOTER TOTALS
    pdf.ln(2)
    pdf.set_font("helvetica", "B", 10)
    pdf.cell(180, 8, f"TOTALE QUANTITÀ:  {int(total_qty)}", border=1, align="R")
    pdf.set_fill_color(*yellow_fill)
    pdf.cell(97, 8, f"TOTALE IMPORTO:  € {total_amount:.2f}", border=1, align="R", fill=True)

    pdf_buffer = io.BytesIO()
    pdf.output(pdf_buffer)
    pdf_buffer.seek(0)

    safe_name = customer_name.replace(' ', '_').replace('/', '')
    return send_file(
        pdf_buffer,
        mimetype="application/pdf",
        as_attachment=True,
        download_name=f"REGALIDEA_PDF_{safe_name}.pdf"
    )

if __name__ == "__main__":
    import os
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)))
