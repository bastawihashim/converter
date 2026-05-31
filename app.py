import os
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import pdfplumber
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import io

app = Flask(__name__)
CORS(app)

# مائیکروسافٹ ورڈ کو مجبور کرنا کہ وہ ٹیکسٹ کو الٹا نہ کرے اور اُردو/عربی رسم الخط میں جوڑے
def apply_perfect_rtl(paragraph, font_name, font_size_pt):
    # پیراگراف کو دائیں سے بائیں (RTL) سیٹ کرنا
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pPr = paragraph._p.get_or_add_pPr()
    
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)
    
    # مائیکروسافٹ ورڈ کے پیچیدہ رسم الخط (Complex Script) کو ایکٹو کرنا
    # یہی وہ جادوئی سیٹنگ ہے جو حروف کو الٹا (Mirror) ہونے سے روکتی ہے
    run = paragraph.runs[0] if paragraph.runs else paragraph.add_run()
    rPr = run._r.get_or_add_rPr()
    
    # ٹیکسٹ کو بتانا کہ یہ عربی/اُردو کی کیٹیگری میں ہے
    rtl_element = OxmlElement('w:rtl')
    rtl_element.set(qn('w:val'), '1')
    rPr.append(rtl_element)
    
    # فونٹس اور سائز کو ورڈ کے اندرونی سسٹم (CS - Complex Script) میں لاک کرنا
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:cs'), font_name)
    rFonts.set(qn('w:ascii'), font_name)
    rFonts.set(qn('w:hAnsi'), font_name)
    rPr.append(rFonts)
    
    sz = OxmlElement('w:sz')
    sz.set(qn('w:val'), str(int(font_size_pt * 2)))
    rPr.append(sz)
    
    szCs = OxmlElement('w:szCs')
    szCs.set(qn('w:val'), str(int(font_size_pt * 2)))
    rPr.append(szCs)

# یہ چیک کرنے کا فنکشن کہ لائن قرآنی آیت/عربی ہے یا عام اُردو
def check_arabic(text):
    # اگر قرآنی اعراب یا علامات موجود ہوں
    for char in text:
        if '\u064b' <= char <= '\u065f' or char in ['\u0671', '\u06d6', '\u06d7', '\u06d8', '\u06d9', '\u06da', '\u06db', '\u06dc']:
            return True
    return False

@app.route('/')
def home():
    return "Perfect Urdu & Arabic PDF to Word Converter Backend is Running!"

@app.route('/convert', methods=['POST'])
def convert_pdf_to_word():
    if 'pdfFile' not in request.files:
        return jsonify({"error": "Koi file upload nahi ki gayi"}), 400
        
    file = request.files['pdfFile']
    if file.filename == '':
        return jsonify({"error": "File ka naam khali hai"}), 400

    try:
        pdf_bytes = file.read()
        extracted_text = ""
        
        # پی ڈی ایف سے اصل حالت میں ٹیکسٹ نکالنا (بغیر کسی الٹی پلٹی کے)
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    extracted_text += text + "\n"

        if not extracted_text.strip():
            return jsonify({"error": "Is PDF me koi text nahi mila."}), 400

        doc = Document()
        
        lines = extracted_text.split('\n')
        for line in lines:
            if line.strip():
                p = doc.add_paragraph()
                run = p.add_run(line) # ٹیکسٹ کو براہِ راست اصل ترتیب میں ڈالنا
                
                # خودکار طریقے سے فونٹ اور سائز کا انتخاب
                if check_arabic(line):
                    # قرآنی آیات اور عربی کے لیے مستند فونٹ اور بڑا سائز
                    font_name = 'Traditional Arabic'
                    font_size = 16
                else:
                    # عام اُردو کے لیے نستعلیق فونٹ
                    font_name = 'Noto Nastaliq Urdu'
                    font_size = 14
                
                # ورڈ کے اندرونی سسٹم پر فکسڈ رائٹ ٹو لیفٹ لاگو کرنا
                apply_perfect_rtl(p, font_name, font_size)

        word_io = io.BytesIO()
        doc.save(word_io)
        word_io.seek(0)
        
        return send_file(
            word_io,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name="Perfect_Urdu_Arabic_Converter.docx"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
