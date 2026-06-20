from flask import Flask, render_template, request, send_file, jsonify, session, redirect, url_for, flash
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Image, Table,
    TableStyle, BaseDocTemplate, Frame, PageTemplate, KeepTogether, PageBreak)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_RIGHT, TA_LEFT
from datetime import datetime
from functools import wraps
from dotenv import load_dotenv
import os, io, random, string
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

load_dotenv()

app = Flask(__name__)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
W, H = A4

# ── Auth Config ───────────────────────────────────────────────
app.secret_key  = os.environ.get("SECRET_KEY", "change-me-in-env")
LOGIN_USER      = os.environ.get("LOGIN_UID", "admin")
LOGIN_PASS      = os.environ.get("LOGIN_PASS", "aparaitech@123")
# ── SMTP Config ──────────────────────────────────────────────
SMTP_HOST     = os.environ.get("SMTP_HOST", "smtp.gmail.com")
SMTP_PORT     = int(os.environ.get("SMTP_PORT", 587))
SMTP_USER     = os.environ.get("SMTP_USER", "your@gmail.com")
SMTP_PASS     = os.environ.get("SMTP_PASS", "your_app_password")
SMTP_FROM     = os.environ.get("SMTP_FROM", SMTP_USER)
# ─────────────────────────────────────────────────────────────

def login_required(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('logged_in'):
            return redirect(url_for('login'))
        return f(*args, **kwargs)
    return decorated

# Aparaitech Brand Colors
DARK = colors.HexColor('#0d2b5e')
CYAN = colors.HexColor('#00aec7')
GREY = colors.HexColor('#555555')

def gp(f): return os.path.join(BASE_DIR, "static", f)

def draw_page(c, doc):
    c.saveState()
    # White background for the whole page
    c.setFillColorRGB(1, 1, 1)
    c.rect(0, 0, W, H, fill=1, stroke=0)
    
    # Cyan accent stripe below header
    c.setFillColor(CYAN)
    c.rect(0, H-92, W, 4, fill=1, stroke=0)
    
    # Logo in header
    logo = gp("logo.png")
    if os.path.exists(logo):
        c.drawImage(logo, 14, H-80, width=80, height=60, preserveAspectRatio=True, mask='auto')
        
    # Company name and Tagline in header
    c.setFillColor(DARK)
    c.setFont('Helvetica-Bold', 17)
    c.drawRightString(W-30, H-46, "APARAITECH SOFTWARE COMPANY")
    c.setFont('Helvetica', 9)
    c.setFillColor(CYAN)
    c.drawRightString(W-30, H-62, "We Build Your Vision")
    
    # Footer line
    c.setStrokeColor(CYAN)
    c.setLineWidth(1)
    c.line(40, 52, W-40, 52)
    
    # Footer Text
    c.setFont('Helvetica', 7.5)
    c.setFillColor(GREY)
    c.drawCentredString(W/2, 38, "Baramati, Pune – 412306, Maharashtra  |  info@aparaitechsoftware.org  |  www.aparaitech.org")
        
    # Page Number
    c.setFont('Helvetica', 8)
    c.drawRightString(W - 40, 20, f"Page {doc.page}")
    c.restoreState()

def build_pdf(data):
    buf = io.BytesIO()

    body  = ParagraphStyle('body', fontSize=9.5, fontName='Helvetica', leading=13, textColor=colors.black, alignment=TA_JUSTIFY, spaceAfter=4)
    bold  = ParagraphStyle('bold', fontSize=9.5, fontName='Helvetica-Bold', leading=13, textColor=DARK, spaceAfter=2)
    title = ParagraphStyle('title', fontSize=13, fontName='Helvetica-Bold', textColor=DARK, alignment=TA_CENTER, spaceAfter=8)
    rgt   = ParagraphStyle('rgt', fontSize=9.5, fontName='Helvetica', textColor=colors.black, alignment=TA_RIGHT)
    digi  = ParagraphStyle('digi', fontSize=8, fontName='Courier', textColor=GREY, alignment=TA_LEFT)

    def sec(n, head, text):
        return [
            Paragraph(f"<b>{n}. {head}</b>", bold),
            Paragraph(text, body),
            Spacer(1, 4)
        ]

    SP = lambda n=4: Spacer(1, n)
    date_str = datetime.now().strftime('%d %B %Y')
    ref_year = datetime.now().strftime('%Y')
    ref = f"APC/HRD/{ref_year}/OFF-" + ''.join(random.choices(string.digits, k=3))
    
    # Format joining / training dates
    joining = data.get('joining_date', '')
    joining_raw = joining
    try: joining = datetime.strptime(joining, '%Y-%m-%d').strftime('%d %B %Y')
    except: pass

    end_raw = data.get('training_end_date', '')
    end_date = end_raw
    try: end_date = datetime.strptime(end_raw, '%Y-%m-%d').strftime('%d %B %Y')
    except: pass

    # Training duration in months
    try:
        d1 = datetime.strptime(joining_raw, '%Y-%m-%d')
        d2 = datetime.strptime(end_raw, '%Y-%m-%d')
        months = (d2.year - d1.year) * 12 + (d2.month - d1.month)
        duration_str = f"{months:02d} Months (Fixed Term)"
    except:
        duration_str = "04 Months (Fixed Term)" 

    E = []
    
    # Top Reference & Date
    top_tbl = Table([
        [Paragraph(f"<b>Ref:</b> {ref}", body),
         Paragraph(f"<b>Date:</b> {date_str}", rgt)]
    ], colWidths=[255, 255], hAlign='LEFT')
    top_tbl.setStyle(TableStyle([
        ('LEFTPADDING', (0,0), (0,0), 0),
        ('RIGHTPADDING', (-1,-1), (-1,-1), 0),
    ]))
    E.append(top_tbl)
    E.append(SP(8))
    
    # Title
    E.append(Paragraph("OFFER OF EMPLOYMENT &amp; APPOINTMENT LETTER", title))
    E.append(SP(4))
    
    # Addressee Data
    emp_name = data.get('employee_name', 'Employee')
    E.append(Paragraph("To,", body))
    E.append(Paragraph(f"<b>Mr./Ms. {emp_name}</b>", bold))
    if data.get('college') and data.get('department'):
        E.append(Paragraph(f"{data.get('college')}, {data.get('department')}", body))
    E.append(Paragraph("Maharashtra, India", body))
    if data.get("email"):
        E.append(Paragraph(f"<b>Email:</b> {data.get('email')}", body))
    E.append(SP(6))
    
    # Subject
    position = data.get('position', 'Developer')
    E.append(Paragraph(f"<b>Subject: Offer of Employment for the position of {position}</b>", body))
    E.append(SP(6))
    
    # Salutation & Intro
    E.append(Paragraph(f"Dear {emp_name},", body))
    E.append(Paragraph(f"We are pleased to confirm your selection for the position of <b>{position}</b> at <b>APARAITECH SOFTWARE COMPANY</b>.", body))
    E.append(Paragraph("This letter outlines the terms and conditions of your employment with us. We are confident that your skills and experience will be a valuable addition to our team.", body))
    E.append(SP(6))
    
    # Terms and Conditions Sections
    E.extend(sec("1", "Position &amp; Appointment", 
        f"You are hereby appointed as <b>{position}</b> and shall report to the designated reporting manager at our Baramati office. Your services may be transferred to any department, project, or location as per business requirements."))
    
    E.extend(sec("2", "INTERNSHIP / TRAINING PERIOD",
        f"&#x2022; &nbsp;<b>Training Duration:</b> {duration_str}<br/>"
        f"&#x2022; &nbsp;<b>Training Start Date:</b> {joining}<br/>"
        f"&#x2022; &nbsp;<b>Training End Date:</b> {end_date}<br/>"
        f"You are required to report at our Baramati office on the training start date along with all original documents for verification."))
    
    E.extend(sec("3", "Probation Period", 
        "You will be on probation/internship for a period of six (6) months from the date of joining. During this period, your performance will be evaluated, and upon successful completion, you will be confirmed as a regular employee. The company reserves the right to extend the probation period if deemed necessary."))
    
    stipend = data.get('stipend', '0')
    E.extend(sec("4", "Compensation &amp; Benefits", 
        f"Your monthly gross salary/stipend shall be <b>{stipend} (Indian Rupees)</b>. The detailed compensation structure, including all allowances and deductions, will be provided separately in the compensation annexure. Salary will be credited to your designated bank account by the last working day of each month."))
    
    E.extend(sec("5", "Pre-Placement Offer (PPO) &amp; Full-Time Employment", 
        "After successful completion of the training period, candidates may be considered for a PPO based on performance, project requirements, academic completion, and position availability. The offered package, if applicable, may range between 2.5 LPA to 4.5 LPA depending on the final evaluation. APARAITECH reserves the right to extend or decline the PPO at its sole discretion. Completion of the internship does not guarantee full-time employment."))
    
    E.extend(sec("6", "Working Hours &amp; Attendance", 
        "The company follows a 6-day work week (9 hours/day), Monday through Saturday, 10:00 AM to 7:30 PM. You may be required to work additional hours during critical project phases. Regular and punctual attendance is essential."))
    
    E.extend(sec("7", "Leave Entitlement", 
        "You shall be entitled to 15 days of Paid Leave and 7 days of Casual Leave per calendar year, in accordance with company policy. Leave availed shall be subject to prior approval from your reporting manager."))
    
    E.extend(sec("8", "Notice Period &amp; Termination", 
        "During the probation period, either party may terminate the employment by providing fifteen (15) days' written notice. Post confirmation, the notice period shall be thirty (30) days from either side. The company may terminate your services without notice in cases of misconduct, breach of trust, or violation of company policies."))
    
    E.extend(sec("9", "Confidentiality &amp; Intellectual Property", 
        "During the course of your employment and thereafter, you shall maintain strict confidentiality regarding all proprietary information, trade secrets, client data, source code, and business strategies. All work products, innovations, and intellectual property created during your employment shall remain the exclusive property of the company."))
    
    E.extend(sec("10", "Code of Conduct", 
        "You are expected to conduct yourself professionally and ethically at all times. You shall comply with all company policies, rules, and regulations as may be communicated from time to time. Any violation may result in disciplinary action."))
    
    # Section 11 - Mandatory Documents Checklist
    docs_list = [
        "<b>Signed Offer Letter:</b> 1 copy signed on all pages.",
        "<b>Academic Records:</b> SSC, HSC, and Degree/Diploma certificates (Photocopy + Original for verification).",
        "<b>Identity Proof:</b> PAN Card and Aadhaar/Voter ID/Driving Licence (Photocopy + Original).",
        "<b>Photographs:</b> 1 recent passport-size color photograph.",
        "<b>Personal Details:</b> Bank account details (AC No, IFSC) and Emergency contact info.",
        "<b>Institutional Docs:</b> Bonafide Certificate / NOC (if applicable)."
    ]
    
    bul = ParagraphStyle('bul', fontSize=9, fontName='Helvetica', leading=13, leftIndent=14, textColor=colors.black, spaceAfter=2)
    E.append(Paragraph("<b>11. MANDATORY DOCUMENTS – JOINING DAY CHECKLIST</b>", bold))
    E.append(Spacer(1, 2))
    for d in docs_list:
        E.append(Paragraph(f"&#x2022;  {d}", bul))
    E.append(Spacer(1, 4))

    # Closing
    E.append(Paragraph("We are delighted to welcome you to the APARAITECH SOFTWARE COMPANY family. Please sign and return the duplicate copy of this letter as your acceptance of the terms and conditions mentioned herein.", body))
    E.append(Paragraph("We look forward to a long and mutually rewarding association.", body))
    E.append(SP(12))

    # -------------------------------------------------------------
    # Company Signatures Section - flows naturally after content
    # -------------------------------------------------------------
    E.append(Paragraph("<b>For APARAITECH SOFTWARE COMPANY</b>", bold))
    E.append(SP(6))

    sp, st = gp("signature.png"), gp("stamp.png")

    from reportlab.platypus import Flowable

    class SignatureBlock(Flowable):
        """Draws signature + digi text + stamp — always on same page."""
        def __init__(self, sig_path, stamp_path):
            Flowable.__init__(self)
            self.sig_path   = sig_path
            self.stamp_path = stamp_path
            self.width  = 4 * inch
            self.height = 1.6 * inch

        def draw(self):
            c = self.canv
            if os.path.exists(self.sig_path):
                c.drawImage(self.sig_path,
                            0, 0.95*inch,
                            width=1.5*inch, height=0.6*inch,
                            preserveAspectRatio=True, mask='auto')

            from reportlab.lib.styles import ParagraphStyle
            from reportlab.lib.enums import TA_LEFT
            from reportlab.platypus import Paragraph
            import datetime as dt
            now = dt.datetime.now().strftime('%d-%m-%Y %H:%M')
            style = ParagraphStyle('d', fontName='Courier', fontSize=8,
                                   textColor=colors.HexColor('#555555'),
                                   leading=12, alignment=TA_LEFT)
            lines = ["Digitally Signed by", f"Date: {now}", "<b>Managing Director</b>"]
            y = 0.6 * inch
            for line in lines:
                p = Paragraph(line, style)
                pw, ph = p.wrap(2.5*inch, 20)
                p.drawOn(c, 0, y)
                y -= ph + 1

            if os.path.exists(self.stamp_path):
                c.drawImage(self.stamp_path,
                            0.8*inch, -0.05*inch,
                            width=1.0*inch, height=1.0*inch,
                            preserveAspectRatio=True, mask='auto')

    E.append(SignatureBlock(
        sig_path   = sp if os.path.exists(sp) else "",
        stamp_path = st if os.path.exists(st) else ""
    ))

    # ----- Page Break for Acceptance Page (Page 3) -----
    E.append(PageBreak())
    
    E.append(SP(20))
    E.append(Paragraph("ACCEPTANCE BY EMPLOYEE", title))
    E.append(SP(10))
    E.append(Paragraph("I have read and understood the terms and conditions of employment as stated above. I hereby accept this offer and agree to abide by the company's policies and regulations.", body))
    E.append(SP(40))
    
    sig_accept_tbl = Table([
        [Paragraph("<b>Signature of Employee</b>", body), Paragraph("<b>Date</b>", rgt)]
    ], colWidths=[255, 255], hAlign='LEFT')
    sig_accept_tbl.setStyle(TableStyle([
        ('LEFTPADDING', (0,0), (0,0), 0),
        ('RIGHTPADDING', (-1,-1), (-1,-1), 0),
    ]))
    E.append(sig_accept_tbl)

    # Document assembly 
    frame = Frame(40, 60, W - 80, H - 165, id='main')
    pt = PageTemplate(id='Letter', frames=[frame], onPage=draw_page)
    doc = BaseDocTemplate(buf, pagesize=A4, pageTemplates=[pt])
    doc.build(E)
    buf.seek(0)

    # Encrypt the PDF
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
    except Exception as e:
        buf.seek(0)
        return buf

# Flask Routes
def send_offer_email(to_email, emp_name, pdf_buf, fname):
    msg = MIMEMultipart()
    msg['From']    = SMTP_FROM
    msg['To']      = to_email
    msg['Subject'] = f"Offer Letter – Aparaitech Software Company"

    body = f"""Dear {emp_name},

Congratulations! Please find your offer letter from Aparaitech Software Company attached to this email.

Kindly sign and return a copy as your acceptance of the terms mentioned.

We look forward to welcoming you to the team!

Warm regards,
HR Department
Aparaitech Software Company
Baramati, Pune – 412306
info@aparaitechsoftware.org | www.aparaitech.org
"""
    msg.attach(MIMEText(body, 'plain'))

    part = MIMEBase('application', 'octet-stream')
    part.set_payload(pdf_buf.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', f'attachment; filename="{fname}"')
    msg.attach(part)

    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as server:
        server.ehlo()
        server.starttls()
        server.login(SMTP_USER, SMTP_PASS)
        server.sendmail(SMTP_FROM, to_email, msg.as_string())

@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    if request.method == "POST":
        uid = request.form.get("uid", "").strip()
        pwd = request.form.get("password", "").strip()
        if uid == LOGIN_USER and pwd == LOGIN_PASS:
            session['logged_in'] = True
            return redirect(url_for('home'))
        else:
            error = "Invalid username or password."
    return render_template("login.html", error=error)

@app.route("/logout")
def logout():
    session.clear()
    return redirect(url_for('login'))

@app.route("/")
@login_required
def home(): 
    return render_template("index.html")

@app.route("/generate", methods=["POST"])
@login_required
def generate():
    keys = ["employee_name", "email", "college", "department", "position", "joining_date", "training_end_date", "stipend"]
    data = {k: request.form.get(k, '') for k in keys}
    
    buf = build_pdf(data)
    fname = f"{data['employee_name'].replace(' ','_')}_Aparaitech_Offer.pdf"

    # Send email
    email_status = "sent"
    email_error  = ""
    if data.get('email'):
        try:
            buf.seek(0)
            send_offer_email(data['email'], data['employee_name'], buf, fname)
        except Exception as e:
            email_status = "failed"
            email_error  = str(e)

    # Also return PDF for download
    buf.seek(0)
    response = send_file(buf, as_attachment=True, download_name=fname, mimetype="application/pdf")
    response.headers['X-Email-Status'] = email_status
    response.headers['X-Email-Error']  = email_error
    response.headers['X-Filename']     = fname
    response.headers['Access-Control-Expose-Headers'] = 'X-Email-Status, X-Email-Error, X-Filename'
    return response


# ── Employee ID Card Generator ────────────────────────────────────────────────

def _make_qr_image(text):
    """Generate a QR code as a PIL Image."""
    import qrcode as qrcode_lib
    qr = qrcode_lib.QRCode(version=1, box_size=4, border=1,
                            error_correction=qrcode_lib.constants.ERROR_CORRECT_M)
    qr.add_data(text)
    qr.make(fit=True)
    return qr.make_image(fill_color="#0d2b5e", back_color="white").convert('RGB')


def _pil_to_reportlab(pil_img):
    """Convert a PIL image to a ReportLab-compatible BytesIO PNG."""
    buf = io.BytesIO()
    pil_img.save(buf, format='PNG')
    buf.seek(0)
    return buf


def _clip_circle(photo_bytes, size=300):
    """Crop a photo to a circle and return as PNG BytesIO (white background)."""
    from PIL import Image as PILImage, ImageDraw
    src = PILImage.open(io.BytesIO(photo_bytes)).convert('RGBA')
    src = src.resize((size, size), PILImage.LANCZOS)
    mask = PILImage.new('L', (size, size), 0)
    draw = ImageDraw.Draw(mask)
    draw.ellipse((0, 0, size, size), fill=255)
    out = PILImage.new('RGBA', (size, size), (255, 255, 255, 0))
    out.paste(src, (0, 0), mask)
    result = PILImage.new('RGBA', (size, size), (255, 255, 255, 255))
    result.alpha_composite(out)
    buf = io.BytesIO()
    result.convert('RGB').save(buf, format='PNG')
    buf.seek(0)
    return buf


def build_id_card(data, photo_bytes=None):
    """Generate a professional vertical employee ID card (front + back) matching Aparaitech design."""
    from reportlab.platypus import BaseDocTemplate, Frame, PageTemplate, Flowable, SimpleDocTemplate
    from reportlab.lib.utils import ImageReader
    import math

    buf = io.BytesIO()

    # ── Vertical CR80 card dimensions ─────────────────────────────
    CARD_W = 2.125 * inch
    CARD_H = 3.375 * inch
    RADIUS = 10  # corner radius in points

    # ── Brand colours ──────────────────────────────────────────────
    DARK    = colors.HexColor('#0d2b5e')
    MID     = colors.HexColor('#1a4a8a')
    CYAN    = colors.HexColor('#00aec7')
    WHITE   = colors.white
    LGREY   = colors.HexColor('#f5f7fb')
    GREY    = colors.HexColor('#555555')
    LWAVE   = colors.HexColor('#cce4f7')  # light wave accent

    emp_id        = data.get('emp_id', 'APC-0000')
    emp_name      = data.get('employee_name', 'Employee Name')
    position      = data.get('position', 'Software Engineer')
    department    = data.get('department', 'Engineering')
    email         = data.get('email', 'info@aparaitech.org')
    phone         = data.get('phone', '+91 63643 26542')
    blood_group   = data.get('blood_group', 'O+')
    joining_date  = data.get('joining_date', '')
    try:
        joining_display = datetime.strptime(joining_date, '%Y-%m-%d').strftime('%d %b %Y')
    except:
        joining_display = joining_date or '01 Jan 2025'

    logo_path = gp("logo.png")

    # ── Pre-generate QR code ───────────────────────────────────────
    qr_data = f"ID:{emp_id}|Name:{emp_name}|Role:{position}|Dept:{department}|DOJ:{joining_display}"
    qr_pil  = _make_qr_image(qr_data)
    qr_buf  = _pil_to_reportlab(qr_pil)

    # ── Prepare clipped photo ──────────────────────────────────────
    photo_buf = None
    if photo_bytes:
        try:
            photo_buf = _clip_circle(photo_bytes, size=300)
        except Exception:
            photo_buf = None

    # ── Helper: draw wave shapes using bezier paths ────────────────
    def draw_waves_front(c, w, h):
        """Draw decorative wave lines on the front card (light blue)."""
        c.saveState()
        c.setStrokeColor(LWAVE)
        c.setLineWidth(1.0)
        # Upper-right wave arc
        p = c.beginPath()
        p.moveTo(w * 0.5, h)
        p.curveTo(w * 0.8, h * 0.92, w * 1.1, h * 0.75, w * 0.95, h * 0.6)
        c.drawPath(p, stroke=1, fill=0)
        p2 = c.beginPath()
        p2.moveTo(w * 0.65, h)
        p2.curveTo(w * 0.95, h * 0.90, w * 1.15, h * 0.72, w, h * 0.55)
        c.drawPath(p2, stroke=1, fill=0)
        # Lower-left wave
        p3 = c.beginPath()
        p3.moveTo(0, h * 0.28)
        p3.curveTo(w * 0.15, h * 0.32, w * 0.35, h * 0.22, w * 0.5, h * 0.28)
        c.drawPath(p3, stroke=1, fill=0)
        c.restoreState()

    def draw_dark_wave_bottom(c, w, h):
        """Draw the large dark blue wave at the bottom of the front card."""
        c.saveState()
        # Dark wave fill – bottom ~35% of card
        wave_h = h * 0.36
        p = c.beginPath()
        p.moveTo(0, 0)
        p.lineTo(w, 0)
        p.lineTo(w, wave_h * 0.55)
        p.curveTo(w * 0.75, wave_h * 0.45, w * 0.55, wave_h * 0.75, w * 0.30, wave_h * 0.65)
        p.curveTo(w * 0.10, wave_h * 0.55, 0, wave_h * 0.75, 0, wave_h * 0.85)
        p.lineTo(0, 0)
        p.close()
        c.setFillColor(DARK)
        c.drawPath(p, stroke=0, fill=1)
        # Second lighter layer on top
        p2 = c.beginPath()
        p2.moveTo(0, wave_h * 0.30)
        p2.curveTo(w * 0.25, wave_h * 0.20, w * 0.55, wave_h * 0.52, w * 0.78, wave_h * 0.38)
        p2.curveTo(w * 0.92, wave_h * 0.28, w, wave_h * 0.42, w, wave_h * 0.55)
        p2.lineTo(w, 0)
        p2.lineTo(0, 0)
        p2.close()
        c.setFillColor(MID)
        c.drawPath(p2, stroke=0, fill=1)
        c.restoreState()

    # ══════════════════════════════════════════════════════
    # FRONT CARD
    # ══════════════════════════════════════════════════════
    class FrontCard(Flowable):
        def __init__(self):
            Flowable.__init__(self)
            self.width  = CARD_W
            self.height = CARD_H

        def draw(self):
            c  = self.canv
            w, h = CARD_W, CARD_H

            # ── White card base ──────────────────────────
            c.saveState()
            c.setFillColor(WHITE)
            c.roundRect(0, 0, w, h, RADIUS, fill=1, stroke=0)
            c.restoreState()

            # ── Clip everything to the card shape ────────
            p = c.beginPath()
            p.roundRect(0, 0, w, h, RADIUS)
            c.clipPath(p, stroke=0)

            # ── Wave decorations (light, background) ─────
            draw_waves_front(c, w, h)

            # ── Top header: logo + company name ──────────
            header_h = 0.90 * inch
            # Logo
            if os.path.exists(logo_path):
                logo_size = 0.48 * inch
                c.drawImage(logo_path,
                            w/2 - logo_size/2, h - header_h + 0.28*inch,
                            width=logo_size, height=logo_size,
                            preserveAspectRatio=True, mask='auto')

            c.setFillColor(DARK)
            c.setFont('Helvetica-Bold', 9)
            c.drawCentredString(w/2, h - 0.46*inch, 'APARAITECH')
            c.setFont('Helvetica', 6)
            pass  # letter spacing not supported
            c.drawCentredString(w/2, h - 0.60*inch, 'SOFTWARE COMPANY')
            pass  # letter spacing not supported

            # ── Divider line with dot ─────────────────────
            line_y = h - 0.70*inch
            c.setStrokeColor(DARK)
            c.setLineWidth(0.6)
            c.line(0.18*inch, line_y, w - 0.18*inch, line_y)
            c.setFillColor(CYAN)
            c.circle(w/2, line_y, 2.5, fill=1, stroke=0)

            # ── Photo circle ──────────────────────────────
            photo_cx = w / 2
            photo_cy = h - 1.42 * inch
            photo_r  = 0.52 * inch

            # Blue ring
            c.setStrokeColor(DARK)
            c.setLineWidth(3)
            c.circle(photo_cx, photo_cy, photo_r + 3, fill=0, stroke=1)

            if photo_buf:
                photo_buf.seek(0)
                img_reader = ImageReader(photo_buf)
                # Clip to circle then draw
                p2 = c.beginPath()
                p2.circle(photo_cx, photo_cy, photo_r)
                c.saveState()
                c.clipPath(p2, stroke=0)
                c.drawImage(img_reader,
                            photo_cx - photo_r, photo_cy - photo_r,
                            width=photo_r*2, height=photo_r*2,
                            preserveAspectRatio=False, mask='auto')
                c.restoreState()
            else:
                # Placeholder circle
                c.setFillColor(LGREY)
                c.circle(photo_cx, photo_cy, photo_r, fill=1, stroke=0)
                c.setFillColor(colors.HexColor('#aabbcc'))
                c.setFont('Helvetica', 7)
                c.drawCentredString(photo_cx, photo_cy - 3, 'PHOTO')

            # ── Name & designation ────────────────────────
            name_y = h - 2.20 * inch
            c.setFillColor(DARK)
            c.setFont('Helvetica-Bold', 11)
            name_upper = emp_name.upper()
            # Scale font if name is long
            if c.stringWidth(name_upper, 'Helvetica-Bold', 11) > w - 0.30*inch:
                c.setFont('Helvetica-Bold', 8.5)
            c.drawCentredString(w/2, name_y, name_upper)

            c.setFillColor(CYAN)
            c.setFont('Helvetica-Bold', 7.5)
            c.drawCentredString(w/2, name_y - 0.18*inch, position.upper())

            # Short divider with dot
            div2_y = name_y - 0.30*inch
            c.setStrokeColor(DARK)
            c.setLineWidth(0.5)
            c.line(w/2 - 0.45*inch, div2_y, w/2 + 0.45*inch, div2_y)
            c.setFillColor(CYAN)
            c.circle(w/2, div2_y, 2, fill=1, stroke=0)

            # Employee ID label
            c.setFillColor(GREY)
            c.setFont('Helvetica', 7)
            c.drawCentredString(w/2, div2_y - 0.18*inch, 'Employee ID')
            c.setFillColor(CYAN)
            c.setFont('Helvetica-Bold', 9)
            c.drawCentredString(w/2, div2_y - 0.32*inch, emp_id)

            # ── Dark wave bottom section ──────────────────
            draw_dark_wave_bottom(c, w, h)

            # ── Bottom info box (white rounded rect) ──────
            box_x = 0.13*inch
            box_y = 0.10*inch
            box_w = w - 0.26*inch
            box_h = 0.78*inch
            c.setFillColor(WHITE)
            c.setStrokeColor(colors.HexColor('#ddeeff'))
            c.setLineWidth(0.5)
            c.roundRect(box_x, box_y, box_w, box_h, 6, fill=1, stroke=1)

            # QR code
            qr_size = 0.58*inch
            qr_buf.seek(0)
            c.drawImage(ImageReader(qr_buf),
                        box_x + 0.06*inch, box_y + 0.10*inch,
                        width=qr_size, height=qr_size, mask='auto')

            # Date of Joining & Blood Group
            info_x = box_x + qr_size + 0.14*inch
            # Calendar icon sim
            c.setFillColor(GREY)
            c.setFont('Helvetica', 6.5)
            c.drawString(info_x, box_y + 0.56*inch, '📅  Date of Joining')
            c.setFillColor(DARK)
            c.setFont('Helvetica-Bold', 8)
            c.drawString(info_x, box_y + 0.40*inch, joining_display)

            c.setFillColor(GREY)
            c.setFont('Helvetica', 6.5)
            c.drawString(info_x, box_y + 0.26*inch, '👤  Blood Group')
            c.setFillColor(DARK)
            c.setFont('Helvetica-Bold', 8)
            c.drawString(info_x, box_y + 0.11*inch, blood_group)

    # ══════════════════════════════════════════════════════
    # BACK CARD
    # ══════════════════════════════════════════════════════
    class BackCard(Flowable):
        def __init__(self):
            Flowable.__init__(self)
            self.width  = CARD_W
            self.height = CARD_H

        def draw(self):
            c  = self.canv
            w, h = CARD_W, CARD_H

            # ── White card base ──────────────────────────
            c.setFillColor(WHITE)
            c.roundRect(0, 0, w, h, RADIUS, fill=1, stroke=0)

            # Clip
            p = c.beginPath()
            p.roundRect(0, 0, w, h, RADIUS)
            c.clipPath(p, stroke=0)

            # ── Top header section ────────────────────────
            # Logo
            if os.path.exists(logo_path):
                logo_size = 0.36*inch
                c.drawImage(logo_path,
                            0.14*inch, h - 0.58*inch,
                            width=logo_size, height=logo_size,
                            preserveAspectRatio=True, mask='auto')

            c.setFillColor(DARK)
            c.setFont('Helvetica-Bold', 10)
            c.drawString(0.58*inch, h - 0.34*inch, 'APARAITECH')
            c.setFont('Helvetica', 5.5)
            pass  # letter spacing not supported
            c.drawString(0.58*inch, h - 0.50*inch, 'SOFTWARE COMPANY')
            pass  # letter spacing not supported

            # ── Divider with dot ─────────────────────────
            div_y = h - 0.66*inch
            c.setStrokeColor(DARK)
            c.setLineWidth(0.7)
            c.line(0.14*inch, div_y, w - 0.14*inch, div_y)
            c.setFillColor(CYAN)
            c.circle(w/2, div_y, 2.5, fill=1, stroke=0)

            # ── Info rows ────────────────────────────────
            def info_row(label, value, y):
                c.setFillColor(GREY)
                c.setFont('Helvetica', 7.5)
                c.drawString(0.16*inch, y, label)
                c.setFillColor(GREY)
                c.drawString(0.16*inch + 0.80*inch, y, ':')
                c.setFillColor(DARK)
                c.setFont('Helvetica-Bold', 7.5)
                c.drawString(0.16*inch + 0.92*inch, y, value)

            row_top = h - 0.86*inch
            gap     = 0.175*inch

            info_row('Employee ID',   emp_id,          row_top)
            info_row('Name',          emp_name,         row_top - gap)
            info_row('Designation',   position,         row_top - 2*gap)
            info_row('Department',    department,       row_top - 3*gap)
            info_row('Date of Joining', joining_display, row_top - 4*gap)
            info_row('Blood Group',   blood_group,      row_top - 5*gap)

            # ── Terms & Conditions ───────────────────────
            terms_y = row_top - 5*gap - 0.22*inch
            c.setFillColor(CYAN)
            c.setFont('Helvetica-Bold', 7.5)
            c.drawString(0.16*inch, terms_y, 'Terms & Conditions')

            terms = [
                'This ID card is the property of Aparaitech Software Company.',
                'It must be worn during office hours.',
                'This card is non-transferable.',
                'If found, please return to the HR Department.',
            ]
            ty = terms_y - 0.14*inch
            c.setFillColor(GREY)
            c.setFont('Helvetica', 6.2)
            for t in terms:
                c.drawString(0.20*inch, ty, f'\u2022  {t}')
                ty -= 0.12*inch

            # ── Signature area ────────────────────────────
            sig_y = ty - 0.08*inch
            # Signature placeholder line
            c.setStrokeColor(DARK)
            c.setLineWidth(0.5)
            c.line(0.16*inch, sig_y, 1.20*inch, sig_y)
            c.setFillColor(DARK)
            c.setFont('Helvetica-Bold', 7)
            c.drawString(0.16*inch, sig_y - 0.14*inch, 'Pratik Pawar')
            c.setFillColor(GREY)
            c.setFont('Helvetica', 6.5)
            c.drawString(0.16*inch, sig_y - 0.26*inch, 'Founder & CEO')

            # ── Dark bottom wave ─────────────────────────
            wave_h = 0.50 * inch
            p2 = c.beginPath()
            p2.moveTo(0, 0)
            p2.lineTo(w, 0)
            p2.lineTo(w, wave_h * 0.65)
            p2.curveTo(w * 0.70, wave_h * 0.55, w * 0.40, wave_h * 0.90, 0, wave_h * 0.70)
            p2.close()
            c.setFillColor(DARK)
            c.drawPath(p2, stroke=0, fill=1)

            p3 = c.beginPath()
            p3.moveTo(0, 0)
            p3.lineTo(w, 0)
            p3.lineTo(w, wave_h * 0.40)
            p3.curveTo(w * 0.60, wave_h * 0.30, w * 0.30, wave_h * 0.60, 0, wave_h * 0.45)
            p3.close()
            c.setFillColor(MID)
            c.drawPath(p3, stroke=0, fill=1)

            # Footer text
            c.setFillColor(WHITE)
            c.setFont('Helvetica', 5.0)
            footer = f'\u260e  {phone}   \u2709  {email}   \u2316  www.aparaitech.org'
            c.drawCentredString(w/2, 0.10*inch, footer)

    # ── Build the two-page PDF ─────────────────────────────────────
    from reportlab.platypus import SimpleDocTemplate, Spacer

    MARGIN   = 0.45 * inch
    PAGE_W   = CARD_W + 2 * MARGIN
    PAGE_H   = CARD_H * 2 + 3 * MARGIN   # front + gap + back + margins

    lbl  = ParagraphStyle('lbl', fontSize=7.5, fontName='Helvetica-Bold',
                           textColor=colors.HexColor('#0d2b5e'), alignment=TA_LEFT,
                           spaceAfter=5)

    doc = SimpleDocTemplate(buf, pagesize=(PAGE_W, PAGE_H),
                            leftMargin=MARGIN, rightMargin=MARGIN,
                            topMargin=MARGIN * 0.7, bottomMargin=MARGIN * 0.7)
    story = [
        Paragraph('▶ FRONT', lbl),
        FrontCard(),
        Spacer(1, 0.22*inch),
        Paragraph('▶ BACK', lbl),
        BackCard(),
    ]
    doc.build(story)
    buf.seek(0)
    return buf


@app.route("/generate-id-card", methods=["POST"])
@login_required
def generate_id_card():
    keys = ["employee_name", "emp_id", "position", "department",
            "email", "phone", "blood_group", "joining_date"]
    data = {k: request.form.get(k, '') for k in keys}

    # Handle optional photo upload
    photo_bytes = None
    photo_file = request.files.get('photo')
    if photo_file and photo_file.filename:
        photo_bytes = photo_file.read()

    buf  = build_id_card(data, photo_bytes=photo_bytes)
    safe  = data['employee_name'].replace(' ', '_')
    fname = f"{safe}_{data['emp_id']}_IDCard.pdf"
    return send_file(buf, as_attachment=True, download_name=fname, mimetype="application/pdf")


if __name__ == "__main__":
    app.run(debug=True)

