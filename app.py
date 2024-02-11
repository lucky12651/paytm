from flask import Flask, render_template, request, send_file
from docx import Document
from docx2pdf import convert
import random
import os

app = Flask(__name__)

def replace_tags(doc, substitutes):
    for paragraph in doc.paragraphs:
        for key, value in substitutes.items():
            tag = f'{{{{{key}}}}}'
            if tag in paragraph.text:
                paragraph.text = paragraph.text.replace(tag, value)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        date = request.form['date']
        order = ''.join(str(random.randint(0, 9)) for _ in range(11))
        ref = ''.join(str(random.randint(0, 9)) for _ in range(11))

        substitutes = {
            'date': date,
            'order': order,
            'ref': ref,
            'chars': '[ < > ? . * { } & % " \' ]',
            'blank': ''
        }

        # Opening template file
        tpl_path = 'paytm.docx'
        result_path = 'Recharge.docx'
        pdf_result_path = 'Recharge.pdf'

        # Load the document
        doc = Document(tpl_path)

        # Replace tags in the document
        replace_tags(doc, substitutes)

        # Save the modified document
        doc.save(result_path)

        # Convert the DOCX file to PDF
        convert(result_path, pdf_result_path)

        # Remove the temporary DOCX file
        os.remove(result_path)

        return send_file(pdf_result_path, as_attachment=True)

    return render_template('index.html')

if __name__ == "__main__":
    app.run(debug=True)
