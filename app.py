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

# الفاظ اور اعراب کو آپس میں جوڑنے کا انجن
import arabic_reshaper

app = Flask(__name__)
CORS(app)

# اُردو، عربی اور قرآنی اعراب کو ایک ساتھ جوڑنے کی سیٹنگ
configuration = {
    'delete_harakat': False,     # اعراب (زیر، زبر، پیش) کو حذف نہیں کرنا
    'support_ligatures': True,   # قرآنی جوڑوں کو درست رکھنا
    'arabic': True,              # عربی کی مکمل سپورٹ
    'farsi': True,               # اُردو/فارسی کے مخصوص حروف (ٹ، ڈ، ڑ، چ، پ، گ) کی سپورٹ
}
reshaper = arabic_reshaper.ArabicReshaper(configuration=configuration)

# مائیکروسافٹ ورڈ کے اندرونی نظام کو فکس کرنا تاکہ وہ الفاظ کو بکھیرے نہیں
def fix_word_rtl_and_fonts(paragraph, font_name, font_size_pt):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pPr = paragraph._p.get_or_add_pPr()
    
    # دائیں سے بائیں نظام ایکٹو کرنا
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)
    
    if paragraph.runs:
        for run in paragraph.runs:
            rPr = run._r.get_or_add_rPr()
            
            # رائٹ ٹو لیفٹ سکرپٹ لاک کرنا
            rtl_element = OxmlElement('w:rtl')
            rtl_element.set(qn('w:val'), '1')
            rPr.append(rtl_element)
            
            # فونٹس کو مائیکروسافٹ ورڈ کے Complex Script (cs) میں سیو کرنا
            rFonts = OxmlElement('w:rFonts')
            rFonts.set(qn('w:cs'), font_name)
            rFonts.set(qn('w:ascii'), font_name)
            rFonts.set(qn('w:hAnsi'), font_name)
            rPr.append(rFonts)
            
            # فونٹ سائز سیٹ کرنا
            sz = OxmlElement('w:sz')
            sz.set(qn('w:val'), str(int(font_size_pt * 2)))
            rPr.append(sz)
            szCs = OxmlElement('w:szCs')
            szCs.set(qn('w:val'), str(int(font_size_pt * 2)))
            rPr.append(szCs)

# یہ چیک کرنے کا فنکشن کہ لائن قرآنی آیت/عربی ہے یا اُردو
def is_arabic_line(text):
    arabic_char_count = len(re.findall(r'[\u064b-\u065f\u0671\u06d6-\u06dc]', text))
    if arabic_char_count > 0:
        return True
    return False

@app.route('/')
def home():
    return "Perfect Urdu, Arabic & Quranic PDF to Word Converter Backend is Running!"

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
        
        # 1. پی ڈی ایف سے راؤنڈ ٹیکسٹ نکالنا
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
                # 🔥 جادوئی قدم: پائتھن کے الٹے حروف کو پہلے ری شیپ کر کے جوڑنا
                reshaped_text = reshaper.reshape(line)
                
                # مائیکروسافٹ ورڈ کے لیے الٹے کیریکٹرز کو فزیکلی سیدھی ترتیب میں لانا
                # یہ لوپ ہر لفظ کے حروف کو بکھرنے اور الٹا ہونے سے روکتا ہے
                final_line = reshaped_text[::-1]
                
                p = doc.add_paragraph()
                p.add_run(final_line)
                
                # خودکار فونٹ اور سائز کا درست انتخاب
                if is_arabic_line(line):
                    font_name = 'Traditional Arabic'
                    font_size = 16
                else:
                    font_name = 'Noto Nastaliq Urdu'
                    font_size = 14
                
                # مائیکروسافٹ ورڈ کے انجن پر سیٹنگز لاگو کرنا
                fix_word_rtl_and_fonts(p, font_name, font_size)

        word_io = io.BytesIO()
        doc.save(word_io)
        word_io.seek(0)
        
        return send_file(
            word_io,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name="Perfect_Urdu_Arabic_Fixed.docx"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
