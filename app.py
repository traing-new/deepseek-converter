import os
import tempfile
import subprocess
import shutil
import re
from flask import Flask, request, send_file, render_template

app = Flask(__name__)

CN_FONTS = ["微软雅黑", "宋体", "黑体", "楷体", "仿宋", "等线"]
EN_FONTS = ["Times New Roman", "Arial", "Calibri", "Cambria", "Consolas", "Courier New"]

PANDOC_PATH = shutil.which('pandoc') or 'pandoc'

def detect_input_format(text):
    if re.match(r'^\s*<(!DOCTYPE|html|head|body|div|table|p|ul|ol|li|span|h[1-6])', text, re.IGNORECASE):
        return 'html'
    return 'markdown'

def set_docx_fonts(docx_path, body_cn, body_en, heading_cn, heading_en):
    from docx import Document
    from docx.shared import Pt, RGBColor
    from docx.oxml.ns import qn
    from docx.oxml import OxmlElement

    doc = Document(docx_path)

    def set_style_font(style, cn, en):
        rPr = style.element.get_or_add_rPr()
        for theme in rPr.findall(qn('w:themeFonts')):
            rPr.remove(theme)
        for old in rPr.findall(qn('w:rFonts')):
            rPr.remove(old)
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:eastAsia'), cn)
        rFonts.set(qn('w:ascii'), en)
        rFonts.set(qn('w:hAnsi'), en)
        rPr.insert(0, rFonts)

    normal = doc.styles['Normal']
    set_style_font(normal, body_cn, body_en)
    normal.font.size = Pt(11)
    normal.font.color.rgb = RGBColor(0,0,0)

    for i in range(1, 10):
        name = f'Heading {i}'
        if name in doc.styles:
            set_style_font(doc.styles[name], heading_cn, heading_en)
            doc.styles[name].font.color.rgb = RGBColor(0,0,0)

    for table in doc.tables:
        tbl = table._tbl
        tblPr = tbl.find(qn('w:tblPr'))
        if tblPr is None:
            tblPr = OxmlElement('w:tblPr')
            tbl.insert(0, tblPr)
        borders = OxmlElement('w:tblBorders')
        for border_name in ['top','left','bottom','right','insideH','insideV']:
            border = OxmlElement(f'w:{border_name}')
            border.set(qn('w:val'), 'single')
            border.set(qn('w:sz'), '4')
            border.set(qn('w:space'), '0')
            border.set(qn('w:color'), '000000')
            borders.append(border)
        tblPr.append(borders)

    doc.save(docx_path)

@app.route('/')
def index():
    return render_template('index.html', cn_fonts=CN_FONTS, en_fonts=EN_FONTS)

@app.route('/convert', methods=['POST'])
def convert():
    text = request.form.get('text', '').strip()
    if not text:
        return '内容为空', 400

    input_mode = request.form.get('input_mode', 'auto')
    output_fmt = request.form.get('output_fmt', 'docx')
    body_cn = request.form.get('body_cn', '宋体')
    body_en = request.form.get('body_en', 'Times New Roman')
    heading_cn = request.form.get('heading_cn', '黑体')
    heading_en = request.form.get('heading_en', 'Arial')

    if input_mode == 'auto':
        detected = detect_input_format(text)
        from_fmt = 'html' if detected == 'html' else 'markdown'
    else:
        from_fmt = input_mode

    tmp_input = None
    tmp_docx = None
    try:
        with tempfile.NamedTemporaryFile(mode='w', suffix='.tmp', delete=False, encoding='utf-8') as f:
            f.write(text)
            tmp_input = f.name

        if output_fmt in ('docx', 'pdf'):
            tmp_docx = tempfile.NamedTemporaryFile(suffix='.docx', delete=False)
            tmp_docx.close()
            subprocess.run([PANDOC_PATH, tmp_input, '-f', from_fmt, '-t', 'docx', '-o', tmp_docx.name], check=True)
            set_docx_fonts(tmp_docx.name, body_cn, body_en, heading_cn, heading_en)

            if output_fmt == 'docx':
                return send_file(tmp_docx.name, as_attachment=True, download_name='converted.docx')
            else:
                tmp_pdf = tempfile.NamedTemporaryFile(suffix='.pdf', delete=False)
                tmp_pdf.close()
                subprocess.run(['libreoffice', '--headless', '--convert-to', 'pdf', '--outdir',
                                os.path.dirname(tmp_pdf.name), tmp_docx.name], check=True)
                pdf_path = tmp_docx.name.rsplit('.',1)[0] + '.pdf'
                return send_file(pdf_path, as_attachment=True, download_name='converted.pdf')
        else:
            tmp_out = tempfile.NamedTemporaryFile(suffix=f'.{output_fmt}', delete=False)
            tmp_out.close()
            subprocess.run([PANDOC_PATH, tmp_input, '-f', from_fmt, '-t', output_fmt, '-o', tmp_out.name], check=True)
            return send_file(tmp_out.name, as_attachment=True, download_name=f'converted.{output_fmt}')
    except Exception as e:
        return str(e), 500
    finally:
        if tmp_input and os.path.exists(tmp_input):
            os.unlink(tmp_input)
        if tmp_docx and os.path.exists(tmp_docx.name):
            os.unlink(tmp_docx.name)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 10000)))