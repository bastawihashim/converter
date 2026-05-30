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
# CORS کو آن کیا ہے تاکہ آپ کی ہوسٹنگر والی ویب سائٹ اس سرور سے بات کر سکے
CORS(app)

# مائیکروسافٹ ورڈ میں رائٹ ٹو لیفٹ (RTL) اردو ٹیکسٹ سیٹ کرنے کا فنکشن
def set_cell_margins_and_rtl(p):
    p_pr = p._p.get_or_add_pPr()
    bidi = OxmlElement('w:bidi')
    bidi.set(qn('w:val'), '1')
    p_pr.append(bidi)

@app.route('/')
def home():
    return "Urdu PDF to Word Converter Backend is Running Successfully!"

@app.route('/convert', methods=['POST'])
def convert_pdf_to_word():
    if 'pdfFile' not in request.files:
        return jsonify({"error": "Koi file upload nahi ki gayi"}), 400
        
    file = request.files['pdfFile']
    if file.filename == '':
        return jsonify({"error": "File ka naam khali hai"}), 400

    try:
        # 1. PDF فائل کو پڑھنا اور ٹیکسٹ نکالنا
        pdf_bytes = file.read()
        extracted_text = ""
        
        with pdfplumber.open(io.BytesIO(pdf_bytes)) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    extracted_text += text + "\n"

        if not extracted_text.strip():
            return jsonify({"error": "Is PDF me koi text nahi mila. Shayed ye scanned image hai."}), 400

        # 2. نیا مائیکروسافٹ ورڈ ڈاکومنٹ بنانا
        doc = Document()
        
        # ٹیکسٹ کو لائن بائی لائن توڑ کر ورڈ میں ایڈ کرنا
        lines = extracted_text.split('\n')
        for line in lines:
            if line.strip():
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.RIGHT # اردو کے لیے رائٹ الائنمنٹ
                set_cell_margins_and_rtl(p) # اردو فارمیٹنگ ایکٹیو کرنا
                
                run = p.add_run(line)
                run.font.name = 'Noto Nastaliq Urdu' # فونٹ کا نام
                run.font.size = Pt(14) # فونٹ سائز 14pt
                
                # ورڈ فائل کو لازمی بتانا کہ یہ اردو فونٹ ہے
                rPr = run._r.get_or_add_rPr()
                rFonts = OxmlElement('w:rFonts')
                rFonts.set(qn('w:cs'), 'Noto Nastaliq Urdu')
                rPr.append(rFonts)

        # 3. ورڈ فائل کو میموری میں ہی سیو کر کے یوزر کو بھیجنا
        word_io = io.BytesIO()
        doc.save(word_io)
        word_io.seek(0)
        
        return send_file(
            word_io,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            as_attachment=True,
            download_name="Urdu-Converted.docx"
        )

    except Exception as e:
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)