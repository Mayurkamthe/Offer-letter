from flask import Flask, render_template, request, send_file
from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Image, Table,
    TableStyle, BaseDocTemplate, Frame, PageTemplate, KeepTogether, PageBreak)
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_RIGHT, TA_LEFT
from datetime import datetime
import os, io, random, string

app = Flask(__name__)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
W, H = A4

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
    c.drawCentredString(W/2, 38, "Baramati, Pune – 412306, Maharashtra  |  hr@aparaitech.com  |  www.aparaitech.com")
        
    # Page Number
    c.setFont('Helvetica', 8)
    c.drawRightString(W - 40, 20, f"Page {doc.page}")
    c.restoreState()

def build_pdf(data):
    buf = io.BytesIO()

    body  = ParagraphStyle('body', fontSize=10, fontName='Helvetica', leading=15, textColor=colors.black, alignment=TA_JUSTIFY, spaceAfter=6)
    bold  = ParagraphStyle('bold', fontSize=10, fontName='Helvetica-Bold', leading=15, textColor=DARK, spaceAfter=4)
    title = ParagraphStyle('title', fontSize=13, fontName='Helvetica-Bold', textColor=DARK, alignment=TA_CENTER, spaceAfter=12)
    rgt   = ParagraphStyle('rgt', fontSize=10, fontName='Helvetica', textColor=colors.black, alignment=TA_RIGHT)
    digi  = ParagraphStyle('digi', fontSize=8, fontName='Courier', textColor=GREY, alignment=TA_LEFT)

    def sec(n, head, text):
        return KeepTogether([
            Paragraph(f"<b>{n}. {head}</b>", bold),
            Paragraph(text, body),
            Spacer(1, 6)
        ])

    SP = lambda n=6: Spacer(1, n)
    date_str = datetime.now().strftime('%d %B %Y')
    ref_year = datetime.now().strftime('%Y')
    ref = f"APC/HRD/{ref_year}/OFF-" + ''.join(random.choices(string.digits, k=3))
    
    # Format joining date
    joining = data.get('joining_date', '')
    try: joining = datetime.strptime(joining, '%Y-%m-%d').strftime('%d %B %Y')
    except: pass

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
    E.append(SP(12))
    
    # Title
    E.append(Paragraph("OFFER OF EMPLOYMENT &amp; APPOINTMENT LETTER", title))
    E.append(SP(6))
    
    # Addressee Data
    emp_name = data.get('employee_name', 'Employee')
    E.append(Paragraph("To,", body))
    E.append(Paragraph(f"<b>Mr./Ms. {emp_name}</b>", bold))
    if data.get('college') and data.get('department'):
        E.append(Paragraph(f"{data.get('college')}, {data.get('department')}", body))
    E.append(Paragraph("Maharashtra, India", body))
    if data.get("email"):
        E.append(Paragraph(f"<b>Email:</b> {data.get('email')}", body))
    E.append(SP(8))
    
    # Subject
    position = data.get('position', 'Developer')
    E.append(Paragraph(f"<b>Subject: Offer of Employment for the position of {position}</b>", body))
    E.append(SP(8))
    
    # Salutation & Intro
    E.append(Paragraph(f"Dear {emp_name},", body))
    E.append(Paragraph(f"We are pleased to confirm your selection for the position of <b>{position}</b> at <b>APARAITECH SOFTWARE COMPANY</b>.", body))
    E.append(Paragraph("This letter outlines the terms and conditions of your employment with us. We are confident that your skills and experience will be a valuable addition to our team.", body))
    E.append(SP(8))
    
    # Terms and Conditions Sections
    E.append(sec("1", "Position &amp; Appointment", 
        f"You are hereby appointed as <b>{position}</b> and shall report to the designated reporting manager at our Baramati office. Your services may be transferred to any department, project, or location as per business requirements."))
    
    E.append(sec("2", "Date of Commencement", 
        f"Your employment shall commence from <b>{joining}</b>. You are required to report at our office on the joining date along with all original documents for verification."))
    
    E.append(sec("3", "Probation Period", 
        "You will be on probation/internship for a period of six (6) months from the date of joining. During this period, your performance will be evaluated, and upon successful completion, you will be confirmed as a regular employee. The company reserves the right to extend the probation period if deemed necessary."))
    
    stipend = data.get('stipend', '0')
    E.append(sec("4", "Compensation &amp; Benefits", 
        f"Your monthly gross salary/stipend shall be <b>{stipend} (Indian Rupees)</b>. The detailed compensation structure, including all allowances and deductions, will be provided separately in the compensation annexure. Salary will be credited to your designated bank account by the last working day of each month."))
    
    E.append(sec("5", "Working Hours &amp; Attendance", 
        "The company follows a 6-day work week (9 hours/day), Monday through Saturday, 10:00 AM to 7:30 PM. You may be required to work additional hours during critical project phases. Regular and punctual attendance is essential."))
    
    E.append(sec("6", "Leave Entitlement", 
        "You shall be entitled to 15 days of Paid Leave and 7 days of Casual Leave per calendar year, in accordance with company policy. Leave availed shall be subject to prior approval from your reporting manager."))
    
    E.append(sec("7", "Notice Period &amp; Termination", 
        "During the probation period, either party may terminate the employment by providing fifteen (15) days' written notice. Post confirmation, the notice period shall be thirty (30) days from either side. The company may terminate your services without notice in cases of misconduct, breach of trust, or violation of company policies."))
    
    E.append(sec("8", "Confidentiality &amp; Intellectual Property", 
        "During the course of your employment and thereafter, you shall maintain strict confidentiality regarding all proprietary information, trade secrets, client data, source code, and business strategies. All work products, innovations, and intellectual property created during your employment shall remain the exclusive property of the company."))
    
    E.append(sec("9", "Code of Conduct", 
        "You are expected to conduct yourself professionally and ethically at all times. You shall comply with all company policies, rules, and regulations as may be communicated from time to time. Any violation may result in disciplinary action."))
    
    # Section 10 - Mandatory Documents Checklist
    docs_list = [
        "1 signed copy of this Offer Letter (duly signed on all pages)",
        "SSC (10th Std) Marksheet &amp; Certificate – 1 photocopy + Original for verification",
        "HSC (12th Std) Marksheet &amp; Certificate – 1 photocopy + Original for verification",
        "Degree / Diploma / Highest Qualification Certificate + All Semester Marksheets – 1 photocopy + Original for verification",
        "1 recent passport-size colour photograph (white background)",
        "PAN Card – Scanned copy + Original for verification",
        "Aadhaar Card / Voter ID / Driving Licence – Scanned copy + Original for verification",
        "Bank Account Details: Bank Name, Account Holder Name, Account Number, IFSC Code, Branch Address",
        "Emergency Contact Details: Name, Relationship, Mobile Number, Address",
        "Academic Institution Bonafide Certificate / NOC (if applicable)",
    ]
    bul = ParagraphStyle('bul', fontSize=9.5, fontName='Helvetica', leading=14, leftIndent=14, textColor=colors.black, spaceAfter=4)
    doc_items = [Paragraph(f"&#x2022;  {d}", bul) for d in docs_list]
    E.append(KeepTogether([
        Paragraph("<b>10. MANDATORY DOCUMENTS – JOINING DAY CHECKLIST</b>", bold),
        Spacer(1, 4),
        *doc_items,
        Spacer(1, 6)
    ]))

    # Closing
    E.append(Paragraph("We are delighted to welcome you to the APARAITECH SOFTWARE COMPANY family. Please sign and return the duplicate copy of this letter as your acceptance of the terms and conditions mentioned herein.", body))
    E.append(Paragraph("We look forward to a long and mutually rewarding association.", body))
    E.append(SP(16))
    
    # -------------------------------------------------------------
    # Company Signatures Section - FIXED for Stamp/Digital Sign
    # -------------------------------------------------------------
    E.append(Paragraph("<b>For APARAITECH SOFTWARE COMPANY</b>", bold))
    E.append(SP(6))
    
    sp, st = gp("signature.png"), gp("stamp.png")
    sig_img  = Image(sp, width=1.5*inch, height=0.6*inch) if os.path.exists(sp) else Paragraph("<i>(Signature)</i>", digi)
    digi_text = Paragraph(
        f"Digitally Signed by<br/>"
        f"Date: {datetime.now().strftime('%d-%m-%Y %H:%M')}<br/>"
        f"<b>Managing Director</b>", digi)

    # Draw signature + digi text block normally
    E.append(sig_img)
    E.append(Spacer(1, 4))
    E.append(digi_text)

    # Stamp overlaid on top using absolute canvas drawing via a custom flowable
    if os.path.exists(st):
        from reportlab.platypus import Flowable
        stamp_path = st
        class StampOverlay(Flowable):
            def __init__(self, path, w=1.0*inch, h=1.0*inch, x_off=10, y_off=-10):
                Flowable.__init__(self)
                self.path = path
                self.w = w; self.h = h
                self.x_off = x_off; self.y_off = y_off
                self.width = 0; self.height = 0  # takes no vertical space
            def draw(self):
                self.canv.drawImage(
                    self.path,
                    self.x_off, self.y_off,
                    width=self.w, height=self.h,
                    preserveAspectRatio=True, mask='auto'
                )
        # Negative height offsets stamp to sit ON top of the digi text above
        E.append(StampOverlay(stamp_path, w=1.05*inch, h=1.05*inch, x_off=5, y_off=2))
    
    # ----- Page Break for Acceptance Page -----
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
@app.route("/")
def home(): 
    return render_template("index.html")

@app.route("/generate", methods=["POST"])
def generate():
    keys = ["employee_name", "email", "college", "department", "position", "joining_date", "stipend"]
    data = {k: request.form.get(k, '') for k in keys}
    
    buf = build_pdf(data)
    fname = f"{data['employee_name'].replace(' ','_')}_Aparaitech_Offer.pdf"
    
    return send_file(buf, as_attachment=True, download_name=fname, mimetype="application/pdf")

if __name__ == "__main__":
    app.run(debug=True)
    
