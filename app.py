from flask import Flask, render_template, request, send_file
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Image,
    Table, TableStyle, HRFlowable, KeepTogether)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch, mm
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from datetime import datetime
import os, io, random, string, struct, zlib

app = Flask(__name__)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
W, H = A4  # 595.27 x 841.89 pts

DARK  = colors.HexColor('#0d2b5e')
CYAN  = colors.HexColor('#00aec7')
BG    = colors.HexColor('#fdf6e3')   # warm cream like OMVSAB
WHITE = colors.white
GRAY  = colors.HexColor('#555555')
LGRAY = colors.HexColor('#dddddd')

def gp(name): return os.path.join(BASE_DIR, "static", name)

def ref_no():
    return "APC/HRD/" + datetime.now().strftime('%Y') + "/OFF-" + \
           ''.join(random.choices(string.digits, k=3))

# ── Numbered section paragraph ──────────────────────────────────────────────
def sec(num, title, text, body_st):
    bold = ParagraphStyle('bold_sec', fontSize=10, fontName='Helvetica-Bold',
                          textColor=DARK, spaceAfter=2)
    items = [Paragraph(f"{num}. {title}", bold),
             Paragraph(text, body_st)]
    return KeepTogether(items)

# ── Canvas callback: background + header + footer every page ────────────────
class LetterCanvas(canvas.Canvas):
    def __init__(self, buf, data):
        super().__init__(buf, pagesize=A4)
        self.data = data
        self._saved_page_states = []

    def showPage(self):
        self._saved_page_states.append(dict(self.__dict__))
        self._startPage()

    def save(self):
        for state in self._saved_page_states:
            self.__dict__.update(state)
            self._draw_page()
            canvas.Canvas.showPage(self)
        canvas.Canvas.save(self)

    def _draw_page(self):
        self.saveState()
        # Cream background
        self.setFillColor(BG)
        self.rect(0, 0, W, H, fill=1, stroke=0)

        # ── HEADER ──
        logo = gp("logo.png")
        if os.path.exists(logo):
            self.drawImage(logo, 36, H-90, width=90, height=50,
                           preserveAspectRatio=True, mask='auto')

        self.setFillColor(DARK)
        self.setFont("Helvetica-Bold", 22)
        self.drawRightString(W-36, H-58, "APARAITECH SOFTWARE COMPANY")
        self.setFont("Helvetica", 10)
        self.setFillColor(CYAN)
        self.drawRightString(W-36, H-72, "We Build Your Vision")

        # Top divider
        self.setStrokeColor(LGRAY)
        self.setLineWidth(0.8)
        self.line(36, H-98, W-36, H-98)

        # ── FOOTER ──
        self.setStrokeColor(LGRAY)
        self.line(36, 52, W-36, 52)
        self.setFont("Helvetica", 7.5)
        self.setFillColor(GRAY)
        footer = ("Registered Office: Baramati, Pune – 412306, Maharashtra, India   |   "
                  "Email: hr@aparaitech.com   |   Web: www.aparaitech.com")
        self.drawCentredString(W/2, 38, footer)

        self.restoreState()

def build_pdf(data):
    buf = io.BytesIO()

    body = ParagraphStyle('body', fontSize=10, fontName='Helvetica',
        leading=16, textColor=colors.HexColor('#222222'),
        alignment=TA_JUSTIFY, spaceAfter=6)
    bold = ParagraphStyle('bold', fontSize=10, fontName='Helvetica-Bold',
        leading=16, textColor=DARK, spaceAfter=4)
    center_title = ParagraphStyle('ct', fontSize=14, fontName='Helvetica-Bold',
        textColor=DARK, alignment=TA_CENTER, spaceAfter=2, spaceBefore=4)
    center_sub = ParagraphStyle('cs', fontSize=10, fontName='Helvetica',
        textColor=GRAY, alignment=TA_CENTER, spaceAfter=8)
    right_st = ParagraphStyle('rt', fontSize=10, fontName='Helvetica',
        textColor=colors.HexColor('#222222'), alignment=TA_RIGHT, spaceAfter=4)
    small = ParagraphStyle('sm', fontSize=8.5, fontName='Helvetica',
        textColor=GRAY, spaceAfter=2)

    E = []
    SP = lambda n=6: Spacer(1, n)

    ref = data.get('ref', ref_no())
    date_str = datetime.now().strftime('%d %B %Y')
    joining = data['joining_date']
    try:
        joining = datetime.strptime(joining, '%Y-%m-%d').strftime('%d %B %Y')
    except: pass

    # Ref + Date
    E.append(SP(4))
    ref_tbl = Table([
        [Paragraph(f"<b>Ref:</b> {ref}", body),
         Paragraph(f"<b>Date:</b> {date_str}", right_st)]
    ], colWidths=[240, 240])
    ref_tbl.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'TOP')]))
    E.append(ref_tbl)
    E.append(SP(10))

    # Title
    E.append(Paragraph("-", center_sub))
    E.append(Paragraph("OFFER &amp; APPOINTMENT LETTER", center_title))
    E.append(Paragraph("-", center_sub))
    E.append(SP(8))

    # To block
    E.append(Paragraph("To,", body))
    E.append(Paragraph(f"<b>Mr./Ms. {data['employee_name']}</b>", bold))
    E.append(Paragraph(f"{data['college']}, {data['department']}.", body))
    E.append(SP(6))
    E.append(Paragraph(f"<b>Subject:</b> Appointment for the post of {data['position']}", body))
    E.append(SP(6))
    E.append(Paragraph(f"Dear {data['employee_name']},", body))
    E.append(SP(4))
    E.append(Paragraph(
        f"Congratulations! With reference to your interview, we are pleased to offer you the position of "
        f"<b>{data['position']}</b> at <b>APARAITECH SOFTWARE COMPANY</b>, subject to the terms and "
        f"conditions mentioned below.", body))
    E.append(SP(10))

    # Sections
    E.append(sec("1","Appointment Details",
        f"You are appointed as <b>{data['position']}</b> at our Baramati Office. Your services may be "
        f"transferred to any department or location of the company, as and when required.<br/>"
        f"Your joining date shall be <b>{joining}</b>.<br/>"
        f"You will be on internship for <b>4 Months</b>, after which you will be confirmed subject to "
        f"satisfactory performance.", body))
    E.append(SP(6))

    E.append(sec("2","Monthly Emoluments",
        f"You will be paid a monthly stipend/salary of <b>\u20b9{data['stipend']}</b> during the internship period.",
        body))
    E.append(SP(6))

    E.append(sec("3","Notice Period",
        "During the probation period, either party may terminate employment by giving <b>15 days'</b> written "
        "notice. If you resign within 1 year from confirmation, the notice period shall be 15 days. "
        "After 1 year, the notice period shall be <b>30 days</b>.", body))
    E.append(SP(6))

    E.append(sec("4","Working Hours",
        "The company follows a <b>6-day work week (9 hours/day)</b>, Monday to Saturday, 10:00 AM – 7:30 PM. "
        "The company reserves the right to modify policies as required.", body))
    E.append(SP(6))

    E.append(sec("5","Leave Entitlement",
        "<b>7 days</b> Casual Leave and <b>15 days</b> Paid Leave per calendar year.", body))
    E.append(SP(6))

    E.append(sec("6","Confidentiality &amp; Conduct",
        "You shall not disclose any confidential information, source code, or client data. All work developed "
        "during your tenure remains the property of the company.", body))
    E.append(SP(6))

    E.append(sec("7","Termination Clause",
        "If any provided information is found to be incorrect, the company reserves the right to terminate "
        "services without notice.", body))
    E.append(SP(14))

    E.append(Paragraph(
        "Please sign and return the duplicate copy of this letter as a token of acceptance. "
        "We look forward to a long and mutually rewarding association with you.", body))
    E.append(SP(16))

    # Signature block
    E.append(Paragraph("<b>For APARAITECH SOFTWARE COMPANY</b>", bold))
    E.append(SP(6))

    sign_p, stamp_p = gp("signature.png"), gp("stamp.png")
    se, te = os.path.exists(sign_p), os.path.exists(stamp_p)

    sig_cell, stm_cell = [], []
    if se: sig_cell = [Image(sign_p, width=1.6*inch, height=0.65*inch)]
    if te: stm_cell = [Image(stamp_p, width=1.1*inch, height=1.1*inch)]

    if se or te:
        row = Table([[sig_cell or [""], stm_cell or [""]]],
                    colWidths=[3*inch, 3*inch])
        row.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
        E.append(row)
        E.append(SP(4))

    digi = ParagraphStyle('digi', fontSize=8, fontName='Courier',
        textColor=GRAY, alignment=TA_RIGHT)
    E.append(Paragraph(
        f"Digitally Signed by<br/>Authorized Signatory<br/>"
        f"Date: {datetime.now().strftime('%Y.%m.%d %H:%M:%S')} IST", digi))
    E.append(SP(4))
    E.append(Paragraph("<b>Managing Director</b>", bold))

    # Build
    doc = SimpleDocTemplate(buf, pagesize=A4,
        rightMargin=36, leftMargin=36,
        topMargin=110, bottomMargin=66,
        encrypt=None)

    lc = LetterCanvas(buf, data)
    doc.build(E, canvasmaker=lambda *a, **kw: LetterCanvas(buf, data))

    # Encrypt to make non-editable (no password to open, no editing)
    raw = buf.getvalue()
    encrypted = encrypt_pdf(raw)
    out = io.BytesIO(encrypted)
    out.seek(0)
    return out

# ── Minimal PDF encryption (RC4 40-bit, restrict editing) ───────────────────
def encrypt_pdf(pdf_bytes):
    """Add owner password encryption to restrict editing/copying."""
    import hashlib, re
    # We'll use pikepdf if available, else return as-is
    try:
        import pikepdf
        src = io.BytesIO(pdf_bytes)
        dst = io.BytesIO()
        with pikepdf.open(src) as pdf:
            pdf.save(dst, encryption=pikepdf.Encryption(
                owner="aparaitech_owner_2026",
                user="",           # no password to open
                allow=pikepdf.Permissions(
                    print_highres=True,
                    print_lowres=True,
                    extract=False,
                    modify_annotation=False,
                    modify_assembly=False,
                    modify_form=False,
                    modify_other=False,
                    accessibility=True,
                )
            ))
        dst.seek(0)
        return dst.read()
    except ImportError:
        return pdf_bytes   # fallback: return unencrypted


@app.route("/")
def home(): return render_template("index.html")

@app.route("/generate", methods=["POST"])
def generate():
    data = {k: request.form[k] for k in
            ["employee_name","college","department","position","joining_date","stipend"]}
    buf = build_pdf(data)
    fname = f"{data['employee_name'].replace(' ','_')}_Aparaitech_Offer_Letter.pdf"
    return send_file(buf, as_attachment=True, download_name=fname, mimetype="application/pdf")

if __name__ == "__main__":
    app.run(debug=True)
