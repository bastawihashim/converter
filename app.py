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
    
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    pPr.append(bidi)
    
    if paragraph.runs:
        for run in paragraph.runs:
            rPr = run._r.get_or_add_rPr()
            
            rtl_element = OxmlElement('w:rtl')
            rtl_element.set(qn('w:val'), '1')
            rPr.append(rtl_element)
            
            lang = OxmlElement('w:lang')
            lang.set(qn('w:bidi'), lang_code)
            rPr.append(lang)
            
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

# قرآنی اعراب اور علامات چیک کرنے کا فنکشن
def is_arabic_line(text):
    arabic_char_count = len(re.findall(r'[\u064b-\u065f\u0671\u06d6-\u06dc]', text))
    if arabic_char_count > 0:
        return True
    return False

# 🔥 ورڈ 10 کی پی ڈی ایف خامیوں کو جڑ سے ختم کرنے کا فائنل ٹوکن انجن
def fix_word10_text_structure(text):
    lines = text.split('\n')
    fixed_lines = []
    
    for line in lines:
        if not line.strip():
            fixed_lines.append("")
            continue
            
        # لائن کو انفرادی الفاظ (Tokens) میں تقسیم کرنا
        tokens = line.split()
        fixed_tokens = []
        
        for token in tokens:
            # اگر اس ٹوکن میں اردو یا عربی حروف/شکلیں موجود ہیں
            if re.search(r'[\u0600-\u06FF\uFB50-\uFDFF\uFE70-\uFEFF]', token):
                # 1. صرف اس مخصوص لفظ کے الٹے حروف کو سیدھا کرنا
                reversed_token = token[::-1]
                # 2. اس کے ظاہری پکسل کوڈز کو حقیقی لاجیکل یونیکوڈ اردو میں تبدیل کرنا
                clean_unicode = unicodedata.normalize('NFKC', reversed_token)
                fixed_tokens.append(clean_unicode)
            else:
                # انگلش الفاظ یا نمبروں (جیسے 2026) کو بالکل نہیں چھیڑنا
                fixed_tokens.append(token)
                
        # الفاظ کی اصل ترتیب کو برقرار رکھتے ہوئے انہیں سنگل اسپیس سے جوڑنا
        fixed_line = " ".join(fixed_tokens)
        fixed_lines.append(fixed_line)
        
    return "\n".join(fixed_lines)

@app.route('/')
def home():
    return "Final Core Word 10 Urdu & Arabic Logical Converter Backend is Live!"

@app.route('/convert', methods=['POST'])
def convert_pdf_to_word():
    if 'pdfFile' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
        
    file = request.files['pdfFile']
    if file.filename == '':
        return jsonify({"error": "Empty file name"}), 400

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

        # 🔥 یہاں ہو رہا ہے فائنل یونیکوڈ اور ٹوکن فکسنگ کا جادو
        perfect_logical_text = fix_word10_text_structure(extracted_text)

        doc = Document()
        
        lines = perfect_logical_text.split('\n')
        for line in lines:
            if line.strip():
                p = doc.add_paragraph()
                p.add_run(line) # اب خالص لاجیکل ٹیکسٹ ورڈ میں جائے گا
                
                if is_arabic_line(line):
                    font_name = 'Traditional Arabic'
                    font_size = 16
                    lang_code = 'ar-SA'
                else:
                    font_name = 'Noto Nastaliq Urdu'
                    font_size = 14
                    lang_code = 'ur-PK'
                
                apply_strict_word_rtl(p, font_name, font_size, lang_code)

        word_io = io.BytesIO()
        doc.save(word_io)
        word_io.seek(0)
        
        return send_file(
            word_io,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name="Urdu_Perfect_Logical_Fixed.docx"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
