from flask import Flask, render_template, request, send_file
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Image, PageBreak, Table, TableStyle, HRFlowable
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import inch
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY
from datetime import datetime
import os, io

app = Flask(__name__)
BASE_DIR = os.path.dirname(os.path.abspath(__file__))

def get_path(f): return os.path.join(BASE_DIR, "static", f)

def build_pdf(data):
    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=0.75*inch, leftMargin=0.75*inch, topMargin=0.5*inch, bottomMargin=0.5*inch)
    styles = getSampleStyleSheet()

    h1 = ParagraphStyle('H1', fontSize=18, fontName='Helvetica-Bold', textColor=colors.HexColor('#0d2b5e'), alignment=TA_CENTER, spaceAfter=4)
    h2 = ParagraphStyle('H2', fontSize=13, fontName='Helvetica-Bold', textColor=colors.HexColor('#00aec7'), alignment=TA_CENTER, spaceAfter=6)
    body = ParagraphStyle('Body', fontSize=10.5, fontName='Helvetica', leading=18, textColor=colors.HexColor('#222222'), alignment=TA_JUSTIFY, spaceAfter=8)
    sec = ParagraphStyle('Sec', fontSize=14, fontName='Helvetica-Bold', textColor=colors.HexColor('#0d2b5e'), spaceBefore=10, spaceAfter=8)
    bul = ParagraphStyle('Bul', fontSize=10.5, fontName='Helvetica', leading=18, leftIndent=20, textColor=colors.HexColor('#222222'), spaceAfter=6)

    E = []
    HR = lambda: HRFlowable(width="100%", thickness=1.5, color=colors.HexColor('#00aec7'), spaceAfter=10)

    # Page 1 - Offer Letter
    logo = get_path("logo.png")
    if os.path.exists(logo):
        img = Image(logo, width=1.8*inch, height=0.9*inch); img.hAlign='LEFT'; E.append(img)
    E.append(HRFlowable(width="100%", thickness=2, color=colors.HexColor('#00aec7'), spaceAfter=8))
    E.append(Paragraph("APARAITECH SOFTWARE COMPANY", h1))
    E.append(Paragraph("OFFER LETTER", h2))
    E.append(HRFlowable(width="100%", thickness=1, color=colors.HexColor('#dddddd'), spaceAfter=12))

    E.append(Paragraph(f"<b>Date:</b> {datetime.now().strftime('%d %B %Y')}", body))
    E.append(Spacer(1,6))
    E.append(Paragraph(f"<b>To,</b><br/>{data['employee_name']}<br/>{data['department']}<br/>{data['college']}", body))
    E.append(Spacer(1,10))
    E.append(Paragraph(f"Dear <b>{data['employee_name']}</b>,", body))
    E.append(Paragraph(f"We are pleased to offer you the position of <b>{data['position']}</b> at <b>Aparaitech Software Company</b>, Baramati. After reviewing your profile and interview performance, we are confident that you will be a valuable addition to our team.", body))

    rows = [["Designation", data['position']], ["Date of Joining", data['joining_date']], ["Training Duration","4 Months"], ["Location","Baramati, Maharashtra"], ["Stipend",f"\u20b9{data['stipend']} per month"], ["Employment Type","Internship / Training"]]
    tbl = Table(rows, colWidths=[2.2*inch, 4.5*inch])
    tbl.setStyle(TableStyle([('BACKGROUND',(0,0),(0,-1),colors.HexColor('#e8f6fb')),('TEXTCOLOR',(0,0),(0,-1),colors.HexColor('#0d2b5e')),('FONTNAME',(0,0),(0,-1),'Helvetica-Bold'),('FONTNAME',(1,0),(1,-1),'Helvetica'),('FONTSIZE',(0,0),(-1,-1),10.5),('GRID',(0,0),(-1,-1),0.5,colors.HexColor('#cccccc')),('TOPPADDING',(0,0),(-1,-1),7),('BOTTOMPADDING',(0,0),(-1,-1),7),('LEFTPADDING',(0,0),(-1,-1),10)]))
    E.append(Spacer(1,8)); E.append(tbl); E.append(Spacer(1,12))
    E.append(Paragraph("This offer is subject to successful completion of your training period and adherence to company policies. We look forward to your contribution and growth with us.", body))
    E.append(Spacer(1,16))
    E.append(Paragraph("Yours sincerely,", body))

    sp, st = get_path("signature.png"), get_path("stamp.png")
    if os.path.exists(sp) and os.path.exists(st):
        sig = Image(sp, width=1.8*inch, height=0.7*inch)
        stm = Image(st, width=1.3*inch, height=1.3*inch)
        row = Table([[sig, stm]], colWidths=[3*inch, 3.7*inch])
        row.setStyle(TableStyle([('VALIGN',(0,0),(-1,-1),'MIDDLE')])); E.append(row)
    elif os.path.exists(sp):
        E.append(Image(sp, width=1.8*inch, height=0.7*inch))
    E.append(Paragraph("<b>HR Department</b><br/>Aparaitech Software Company", body))

    # Page 2 - Training Policy
    E.append(PageBreak()); E.append(Paragraph("Training Policy", sec)); E.append(HR())
    for p in ["Candidate must follow all company rules and regulations strictly.","Working hours: <b>10:00 AM to 7:30 PM</b> (Monday to Saturday).","Notice period is <b>1 Month</b>.","Confidentiality of company data and client information must be maintained at all times.","Pre-Placement Offer (PPO) is performance-based and not guaranteed.","Candidates are expected to maintain professional conduct throughout the training.","Any violation of policies may result in immediate termination of training."]:
        E.append(Paragraph(f"&#x2022;  {p}", bul))

    # Page 3 - Rules
    E.append(PageBreak()); E.append(Paragraph("Rules & Regulations", sec)); E.append(HR())
    for r in ["Office discipline is mandatory at all times.","Attendance and punctuality must be strictly maintained.","<b>English communication</b> is required in the workplace.","Leave must be informed and approved in advance.","Misconduct or insubordination may lead to immediate termination.","Mobile phone usage during working hours should be limited to official purposes.","Candidates must maintain a clean and organized work environment."]:
        E.append(Paragraph(f"&#x2022;  {r}", bul))

    # Page 4 - Annexure
    E.append(PageBreak()); E.append(Paragraph("Annexure \u2013 Documents Required", sec)); E.append(HR())
    E.append(Paragraph("Please bring the following original documents on the date of joining:", body))
    for d in ["SSC & HSC Marksheets (Originals + Photocopies)","Latest College Marksheet","2 Passport Size Photographs","PAN Card / Aadhaar Card","Bank Account Details (Passbook or Cancelled Cheque)","Original Documents for verification"]:
        E.append(Paragraph(f"&#x2022;  {d}", bul))

    # Page 5 - Salary
    E.append(PageBreak()); E.append(Paragraph("Performance Based Salary Structure", sec)); E.append(HR())
    srows = [["Parameter","Details"],["Revenue Target","\u20b91,00,000 per month"],["Training Period","4 Months"],["Stipend",f"\u20b9{data['stipend']} per month"],["Stipend Eligibility","On achieving 100% monthly target"],["Salary Type","Performance-based"],["PPO","Based on overall performance"]]
    st2 = Table(srows, colWidths=[3*inch, 3.7*inch])
    st2.setStyle(TableStyle([('BACKGROUND',(0,0),(-1,0),colors.HexColor('#0d2b5e')),('TEXTCOLOR',(0,0),(-1,0),colors.white),('FONTNAME',(0,0),(-1,0),'Helvetica-Bold'),('FONTSIZE',(0,0),(-1,-1),10.5),('BACKGROUND',(0,1),(0,-1),colors.HexColor('#e8f6fb')),('FONTNAME',(0,1),(0,-1),'Helvetica-Bold'),('TEXTCOLOR',(0,1),(0,-1),colors.HexColor('#0d2b5e')),('GRID',(0,0),(-1,-1),0.5,colors.HexColor('#aaaaaa')),('TOPPADDING',(0,0),(-1,-1),8),('BOTTOMPADDING',(0,0),(-1,-1),8),('LEFTPADDING',(0,0),(-1,-1),10)]))
    E.append(st2); E.append(Spacer(1,20))
    E.append(Paragraph("Note: Stipend will be credited only upon achieving the defined monthly revenue target. Aparaitech reserves the right to modify the salary structure based on business requirements.", body))

    doc.build(E)
    buffer.seek(0)
    return buffer

@app.route("/")
def home(): return render_template("index.html")

@app.route("/generate", methods=["POST"])
def generate():
    data = {k: request.form[k] for k in ["employee_name","college","department","position","joining_date","stipend"]}
    buf = build_pdf(data)
    fname = f"{data['employee_name'].replace(' ','_')}_offer_letter.pdf"
    return send_file(buf, as_attachment=True, download_name=fname, mimetype="application/pdf")

if __name__ == "__main__":
    app.run(debug=True)
