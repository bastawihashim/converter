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
import unicodedata  # یونیکوڈ کو اصلی حالت میں لانے کا آفیشل انجن
import re

app = Flask(__name__)
CORS(app)

# مائیکروسافٹ ورڈ کو آفیشل اردو/عربی اسکرپٹ (RTL) پر لاک کرنے کا فنکشن
def apply_strict_word_rtl(paragraph, font_name, font_size_pt, lang_code):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    pPr = paragraph._p.get_or_add_pPr()
    
    # 1. پیراگراف دائیں سے بائیں سیٹ کریں
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)
    
    if paragraph.runs:
        for run in paragraph.runs:
            rPr = run._r.get_or_add_rPr()
            
            # 2. رائٹ ٹو لیفٹ سکرپٹ ایکٹو کریں
            rtl_element = OxmlElement('w:rtl')
            rtl_element.set(qn('w:val'), '1')
            rPr.append(rtl_element)
            
            # 3. مائیکروسافٹ ورڈ کا ڈیفالٹ لینگویج انجن لاک کریں
            lang = OxmlElement('w:lang')
            lang.set(qn('w:bidi'), lang_code)
            rPr.append(lang)
            
            # 4. فونٹس کی ترتیب کمپلیکس اسکرپٹ (cs) میں سیو کریں
            rFonts = OxmlElement('w:rFonts')
            rFonts.set(qn('w:cs'), font_name)
            rFonts.set(qn('w:ascii'), font_name)
            rFonts.set(qn('w:hAnsi'), font_name)
            rPr.append(rFonts)
            
            # 5. فونٹ سائز لاک کریں
            sz = OxmlElement('w:sz')
            sz.set(qn('w:val'), str(int(font_size_pt * 2)))
            rPr.append(sz)
            szCs = OxmlElement('w:szCs')
            szCs.set(qn('w:val'), str(int(font_size_pt * 2)))
            rPr.append(szCs)

# قرآنی اعراب اور علامات چیک کرنے کا فنکشن
def is_arabic_line(text):
    arabic_char_count = len(re.findall(r'[\u064b-\u065f\u0671\u06d6-\u06dc]', text))
    if arabic_char_count > 0:
        return True
    return False

# 🔥 جادوئی یونیکوڈ نارملائزر: جو ہر قسم کی الٹی پلٹی شکلوں کو اصلی اردو ٹیکسٹ میں بدل دے گا
def convert_to_clean_unicode(text):
    # NFKC فارمیٹ تمام ظاہری پکسل کوڈز کو حقیقی لاجیکل یونیکوڈ میں تبدیل کر دیتا ہے
    return unicodedata.normalize('NFKC', text)

@app.route('/')
def home():
    return "Pure Logical Unicode Urdu & Arabic PDF to Word Converter Backend is Live!"

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

        # 🔥 یہاں ہو رہا ہے اصل یونیکوڈ کا جادو
        perfect_unicode_text = convert_to_clean_unicode(extracted_text)

        doc = Document()
        
        lines = perfect_unicode_text.split('\n')
        for line in lines:
            if line.strip():
                p = doc.add_paragraph()
                p.add_run(line) # اب خالص لاجیکل یونیکوڈ ٹیکسٹ ورڈ میں جائے گا
                
                # خودکار طریقے سے فونٹ اور زبان کا انتخاب
                if is_arabic_line(line):
                    font_name = 'Traditional Arabic'
                    font_size = 16
                    lang_code = 'ar-SA'
                else:
                    font_name = 'Noto Nastaliq Urdu'
                    font_size = 14
                    lang_code = 'ur-PK'
                
                # مائیکروسافٹ ورڈ کی سیٹنگز لاگو کرنا
                apply_strict_word_rtl(p, font_name, font_size, lang_code)

        word_io = io.BytesIO()
        doc.save(word_io)
        word_io.seek(0)
        
        return send_file(
            word_io,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name="Urdu_Arabic_Perfect_Unicode.docx"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
