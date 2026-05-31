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
import re

# اعراب اور قرآنی رموز کی حفاظت کے لیے لائبریریز
import arabic_reshaper
from bidi.algorithm import get_display

app = Flask(__name__)
CORS(app)

# قرآنی رموز، اعراب، اُردو اور عربی کے تمام حروف کو ایک ساتھ فعال کرنے کی سیٹنگ
configuration = {
    'delete_harakat': False,     # زیر، زبر، پیش (اعراب) کو حذف نہیں کرنا
    'support_ligatures': True,   # لاطینی اور قرآنی جوڑوں کو درست رکھنا
    'arabic': True,              # عربی ٹیکسٹ کی مکمل سپورٹ
    'farsi': True,               # اُردو/فارسی کے مخصوص حروف کی سپورٹ
}
reshaper = arabic_reshaper.ArabicReshaper(configuration=configuration)

# ورڈ ڈاکومنٹ میں رائٹ ٹو لیفٹ (RTL) سپورٹ ایکٹیو کرنے کا فنکشن
def set_paragraph_rtl(p):
    p_pr = p._p.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    p_pr.append(bidi)

# یہ معلوم کرنے کا فنکشن کہ لائن میں زیادہ عربی/قرآنی آیت ہے یا اُردو
def is_pure_arabic(text):
    arabic_char_count = len(re.findall(r'[\u064b-\u065f\u0671\u06d6-\u06dc]', text))
    if arabic_char_count > 0 or ("|" in text):
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
                # اُردو، عربی اور قرآنی اعراب کو درست شکل دینا
                reshaped_text = reshaper.reshape(line)
                bidi_text = get_display(reshaped_text)
                
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
                set_paragraph_rtl(p)
                
                run = p.add_run(bidi_text)
                
                # خودکار فونٹ سلیکشن
                if is_pure_arabic(line):
                    run.font.name = 'Traditional Arabic'
                    run.font.size = Pt(16)
                    font_tag = 'Traditional Arabic'
                else:
                    run.font.name = 'Noto Nastaliq Urdu'
                    run.font.size = Pt(14)
                    font_tag = 'Noto Nastaliq Urdu'
                
                rPr = run._r.get_or_add_rPr()
                rFonts = OxmlElement('w:rFonts')
                rFonts.set(qn('w:cs'), font_tag)
                rPr.append(rFonts)

        word_io = io.BytesIO()
        doc.save(word_io)
        word_io.seek(0)
        
        return send_file(
            word_io,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name="Perfect-Urdu-Arabic-Converted.docx"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
