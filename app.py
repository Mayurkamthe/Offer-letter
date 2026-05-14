from flask import Flask, render_template, request, send_file
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Image, Table,
    TableStyle, HRFlowable, BaseDocTemplate, Frame, PageTemplate, KeepTogether)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_RIGHT
from datetime import datetime
import os, io, random, string

app = Flask(__name__)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
W, H = A4
DARK = colors.HexColor('#0d2b5e')
CYAN = colors.HexColor('#00aec7')

def gp(f): return os.path.join(BASE_DIR, "static", f)

def draw_page(c, doc):
    c.saveState()
    # White bg
    c.setFillColorRGB(1, 1, 1)
    c.rect(0, 0, W, H, fill=1, stroke=0)
    # Dark blue header bar
    c.setFillColor(DARK)
    c.rect(0, H-88, W, 88, fill=1, stroke=0)
    # Cyan accent stripe below header
    c.setFillColor(CYAN)
    c.rect(0, H-92, W, 4, fill=1, stroke=0)
    # Cyan left sidebar
    c.setFillColor(CYAN)
    c.rect(0, 0, 5, H-92, fill=1, stroke=0)
    # Logo in header
    logo = gp("logo.png")
    if os.path.exists(logo):
        c.drawImage(logo, 14, H-80, width=80, height=60,
                    preserveAspectRatio=True, mask='auto')
    # Company name
    c.setFillColorRGB(1, 1, 1)
    c.setFont('Helvetica-Bold', 17)
    c.drawRightString(W-30, H-46, "APARAITECH SOFTWARE COMPANY")
    c.setFont('Helvetica', 9)
    c.setFillColor(CYAN)
    c.drawRightString(W-30, H-62, "We Build Your Vision")
    # Footer line
    c.setStrokeColor(CYAN)
    c.setLineWidth(1)
    c.line(20, 52, W-20, 52)
    c.setFont('Helvetica', 7.5)
    c.setFillColor(colors.HexColor('#555555'))
    c.drawCentredString(W/2, 38,
        "Baramati, Pune – 412306, Maharashtra  |  hr@aparaitech.com  |  www.aparaitech.com")
    c.restoreState()

def build_pdf(data):
    buf = io.BytesIO()

    body  = ParagraphStyle('body',  fontSize=10, fontName='Helvetica',     leading=16, textColor=colors.HexColor('#222222'), alignment=TA_JUSTIFY, spaceAfter=6)
    bold  = ParagraphStyle('bold',  fontSize=10, fontName='Helvetica-Bold', leading=16, textColor=DARK, spaceAfter=4)
    title = ParagraphStyle('title', fontSize=14, fontName='Helvetica-Bold', textColor=DARK, alignment=TA_CENTER, spaceAfter=4)
    rgt   = ParagraphStyle('rgt',   fontSize=10, fontName='Helvetica',      textColor=colors.HexColor('#222222'), alignment=TA_RIGHT)
    digi  = ParagraphStyle('digi',  fontSize=8,  fontName='Courier',        textColor=colors.HexColor('#555555'), alignment=TA_RIGHT)

    def sec(n, head, text):
        return KeepTogether([
            Paragraph(f"<b>{n}. {head}</b>", bold),
            Paragraph(text, body),
            Spacer(1, 4)
        ])

    SP = lambda n=6: Spacer(1, n)
    date_str = datetime.now().strftime('%d %B %Y')
    ref = "APC/HRD/" + datetime.now().strftime('%Y') + "/OFF-" + \
          ''.join(random.choices(string.digits, k=3))
    joining = data['joining_date']
    try: joining = datetime.strptime(joining, '%Y-%m-%d').strftime('%d %B %Y')
    except: pass

    E = []
    E.append(SP(4))
    E.append(Table([
        [Paragraph(f"<b>Ref:</b> {ref}", body),
         Paragraph(f"<b>Date:</b> {date_str}", rgt)]
    ], colWidths=[240, 240]))
    E.append(SP(12))
    E.append(Paragraph("OFFER &amp; APPOINTMENT LETTER", title))
    E.append(SP(12))
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
    E.append(sec("1","Appointment Details",
        f"You are appointed as <b>{data['position']}</b> at our Baramati Office. Your services may be "
        f"transferred to any department or location of the company, as and when required.<br/>"
        f"Your joining date shall be <b>{joining}</b>.<br/>"
        f"You will be on internship for <b>4 Months</b>, after which you will be confirmed subject to "
        f"satisfactory performance."))
    E.append(sec("2","Monthly Emoluments",
        f"You will be paid a monthly stipend/salary of <b>\u20b9{data['stipend']}</b> during the internship period."))
    E.append(sec("3","Notice Period",
        "During the probation period, either party may terminate employment by giving <b>15 days'</b> "
        "written notice. After 1 year, the notice period shall be <b>30 days</b>."))
    E.append(sec("4","Working Hours",
        "The company follows a <b>6-day work week (9 hours/day)</b>, Monday–Saturday, 10:00 AM – 7:30 PM. "
        "The company reserves the right to modify policies as required."))
    E.append(sec("5","Leave Entitlement",
        "<b>7 days</b> Casual Leave and <b>15 days</b> Paid Leave per calendar year."))
    E.append(sec("6","Confidentiality &amp; Conduct",
        "You shall not disclose any confidential information, source code, or client data. All work "
        "developed during your tenure remains the property of the company."))
    E.append(sec("7","Termination Clause",
        "If any provided information is found to be incorrect, the company reserves the right to "
        "terminate services without notice."))
    E.append(SP(14))
    E.append(Paragraph(
        "Please sign and return the duplicate copy of this letter as a token of acceptance. "
        "We look forward to a long and mutually rewarding association with you.", body))
    E.append(SP(16))
    E.append(Paragraph("<b>For APARAITECH SOFTWARE COMPANY</b>", bold))
    E.append(SP(6))

    sp, st = gp("signature.png"), gp("stamp.png")
    c1 = Image(sp, width=1.5*inch, height=0.6*inch) if os.path.exists(sp) else Paragraph("", body)
    c2 = Image(st, width=1.1*inch, height=1.1*inch) if os.path.exists(st) else Paragraph("", body)
    sig_tbl = Table([[c1, c2]], colWidths=[3*inch, 3*inch])
    sig_tbl.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'MIDDLE')]))
    E.append(sig_tbl)
    E.append(SP(4))
    E.append(Paragraph(
        f"Digitally Signed by<br/>Authorized Signatory<br/>"
        f"Date: {datetime.now().strftime('%Y.%m.%d %H:%M:%S')} IST", digi))
    E.append(SP(4))
    E.append(Paragraph("<b>Managing Director</b>", bold))

    frame = Frame(25, 60, W-50, H-160, id='main')
    pt = PageTemplate(id='Letter', frames=[frame], onPage=draw_page)
    doc = BaseDocTemplate(buf, pagesize=A4, pageTemplates=[pt])
    doc.build(E)
    buf.seek(0)

    # Encrypt: non-editable, no open password
    try:
        import pikepdf
        src = io.BytesIO(buf.read())
        dst = io.BytesIO()
        with pikepdf.open(src) as pdf:
            pdf.save(dst, encryption=pikepdf.Encryption(
                owner="aparaitech2026", user="",
                allow=pikepdf.Permissions(
                    print_highres=True, print_lowres=True,
                    extract=False, modify_annotation=False,
                    modify_assembly=False, modify_form=False,
                    modify_other=False, accessibility=True)))
        dst.seek(0)
        return dst
    except Exception:
        buf.seek(0)
        return buf

@app.route("/")
def home(): return render_template("index.html")

@app.route("/generate", methods=["POST"])
def generate():
    data = {k: request.form[k] for k in
            ["employee_name","college","department","position","joining_date","stipend"]}
    buf = build_pdf(data)
    fname = f"{data['employee_name'].replace(' ','_')}_Aparaitech_Offer.pdf"
    return send_file(buf, as_attachment=True, download_name=fname, mimetype="application/pdf")

if __name__ == "__main__":
    app.run(debug=True)
