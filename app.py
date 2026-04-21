from flask import Flask, request, send_file
from flask_cors import CORS
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import io

app = Flask(__name__)
CORS(app)

def add_page_numbers(section):
    footer = section.footer
    para = footer.paragraphs[0]
    para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = para.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.text = 'PAGE'
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)

@app.route('/format', methods=['POST'])
def format_book():
    file = request.files['file']
    trim_width = float(request.form['trim_width'])
    trim_height = float(request.form['trim_height'])
    margin_top = float(request.form['margin_top'])
    margin_bottom = float(request.form['margin_bottom'])
    margin_inside = float(request.form['margin_inside'])
    margin_outside = float(request.form['margin_outside'])
    author_name = request.form['author_name']
    book_title = request.form['book_title']
    font_name = request.form.get('font_name', 'Garamond')
    font_size = float(request.form.get('font_size', 12))

    doc = Document(file)
    section = doc.sections[0]

    section.page_width = Inches(trim_width)
    section.page_height = Inches(trim_height)
    section.top_margin = Inches(margin_top)
    section.bottom_margin = Inches(margin_bottom)
    section.left_margin = Inches(margin_inside)
    section.right_margin = Inches(margin_outside)

    for para in doc.paragraphs:
        is_heading = para.style.name.startswith('Heading')
        is_title = para.style.name in ['Title', 'Subtitle']
        is_empty = len(para.text.strip()) == 0

        if is_empty:
            para.paragraph_format.space_before = Pt(0)
            para.paragraph_format.space_after = Pt(0)
            continue

        if is_heading or is_title:
            para.paragraph_format.space_before = Pt(24)
            para.paragraph_format.space_after = Pt(6)
            continue

        para.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        para.paragraph_format.space_before = Pt(0)
        para.paragraph_format.space_after = Pt(0)
        para.paragraph_format.first_line_indent = Inches(0.3)

        for run in para.runs:
            run.font.name = font_name
            run.font.size = Pt(font_size)

    header = section.header
    header_para = header.paragraphs[0] if header.paragraphs else header.add_paragraph()
    header_para.clear()
    header_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = header_para.add_run(f"{author_name}  |  {book_title}")
    run.font.name = font_name
    run.font.size = Pt(9)
    run.font.color.rgb = RGBColor(0x88, 0x88, 0x88)

    add_page_numbers(section)

    output = io.BytesIO()
    doc.save(output)
    output.seek(0)

    return send_file(
        output,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        as_attachment=True,
        download_name='formatted_manuscript.docx'
    )

if __name__ == '__main__':
    app.run(debug=True)