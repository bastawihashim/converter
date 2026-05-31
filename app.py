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

import arabic_reshaper
from bidi.algorithm import get_display

app = Flask(__name__)
CORS(app)

# اُردو، عربی اور قرآنی اعراب کی مکمل سپورٹ کی کنفیگریشن
configuration = {
    'delete_harakat': False,     # اعراب کو حذف نہیں کرنا
    'support_ligatures': True,   # قرآنی جوڑوں کو برقرار رکھنا
    'arabic': True,              # عربی کی مکمل سپورٹ
    'farsi': True,               # اُردو اور فارسی کے مخصوص حروف کی سپورٹ
}
reshaper = arabic_reshaper.ArabicReshaper(configuration=configuration)

# مائیکروسافٹ ورڈ کو مینوئل الٹنے کے بجائے آفیشل مڈل ایسٹ لینگویج ٹیگ دینا
def setup_microsoft_word_rtl(paragraph, font_name, font_size_pt):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pPr = paragraph._p.get_or_add_pPr()
    
    # دائیں سے بائیں (RTL) الائنمنٹ لاگو کرنا
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)
    
    if paragraph.runs:
        for run in paragraph.runs:
            rPr = run._r.get_or_add_rPr()
            
            # رائٹ ٹو لیفٹ سکرپٹ ایکٹو کرنا
            rtl_element = OxmlElement('w:rtl')
            rtl_element.set(qn('w:val'), '1')
            rPr.append(rtl_element)
            
            # مائیکروسافٹ ورڈ کو بتانا کہ یہ کمپلیکس اسکرپٹ (CS) ہے تاکہ حروف نہ بکھریں
            rFonts = OxmlElement('w:rFonts')
            rFonts.set(qn('w:cs'), font_name)
            rFonts.set(qn('w:ascii'), font_name)
            rFonts.set(qn('w:hAnsi'), font_name)
            rPr.append(rFonts)
            
            # فونٹ کا سائز سیٹ کرنا
            sz = OxmlElement('w:sz')
            sz.set(qn('w:val'), str(int(font_size_pt * 2)))
            rPr.append(sz)
            szCs = OxmlElement('w:szCs')
            szCs.set(qn('w:val'), str(int(font_size_pt * 2)))
            rPr.append(szCs)

# قرآنی علامات اور اعراب کو چیک کرنے کا فنکشن
def is_quranic_or_arabic(text):
    arabic_char_count = len(re.findall(r'[\u064b-\u065f\u0671\u06d6-\u06dc]', text))
    if arabic_char_count > 0:
        return True
    return False

@app.route('/')
def home():
    return "Official Urdu & Arabic PDF to Word Converter Backend is Live!"

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
        
        # پی ڈی ایف سے ٹیکسٹ نکالنا
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
                # 1. پہلے اعراب اور حروف کو معیاری خوبصورت شکل میں جوڑیں
                reshaped_text = reshaper.reshape(line)
                # 2. پھر دائیں سے بائیں کی بائی ڈائریکشنل ترتیب کو آفیشل طریقے سے ترتیب دیں (کوئی مینوئل الٹ پھیر نہیں)
                display_line = get_display(reshaped_text)
                
                p = doc.add_paragraph()
                p.add_run(display_line)
                
                # خودکار طریقے سے فونٹ کا انتخاب
                if is_quranic_or_arabic(line):
                    font_name = 'Traditional Arabic'
                    font_size = 16
                else:
                    font_name = 'Noto Nastaliq Urdu'
                    font_size = 14
                
                # مائیکروسافٹ ورڈ کی اندرونی سیٹنگز لاگو کرنا
                setup_microsoft_word_rtl(p, font_name, font_size)

        word_io = io.BytesIO()
        doc.save(word_io)
        word_io.seek(0)
        
        return send_file(
            word_io,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name="Urdu_Arabic_Converted.docx"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
