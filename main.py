import os

import docx
from docxtpl import DocxTemplate
from flask import Flask, request, render_template
import sqlite3

app = Flask(__name__)

pointers = []
context = {}
filename = ''
pointer_sample = []
results = []


@app.route('/', methods=['POST', 'GET'])
def main():
    if request.method == 'GET':
        file_names = os.listdir('tpl')
        return render_template('form1.html', file_names=file_names)
    elif request.method == 'POST':
        global filename
        filename = request.form['file_name']
        db_name = request.form['db_name']
        doc = docx.Document(f'tpl/{filename}')
        all_paragraphs = doc.paragraphs
        for paragraph in all_paragraphs:
            if '{' in paragraph.text:
                a = (paragraph.text[paragraph.text.find('{') + 2: paragraph.text.find('}')])
                if a not in pointers:
                    pointers.append(a)
        con = sqlite3.connect(f"{db_name}")
        cur = con.cursor()
        result = cur.execute("""SELECT * FROM fields""").fetchall()
        for elem in result:
            f = {}
            f[elem[1]] = elem[2]
            pointer_sample.append(f)
        con.close()
        return save()


def save():
    return render_template('form2.html', pointers=pointers, pointer_sample=pointer_sample)


@app.route('/savefile', methods=['GET'])
def savefile():
    file_name = request.args.get('file_name')
    for i in pointers:
        a = request.args.get(f'{i}')
        results.append(a)
    for i in range(len(pointers)):
        context[pointers[i]] = results[i]
    print(context)
    doc = DocxTemplate(f'tpl/{filename}')
    doc.render(context)
    doc.save(f'result/{file_name}')
    return 'OK'


if __name__ == '__main__':
    app.run(port=8080, host='127.0.0.1')
