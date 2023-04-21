import os

import docx
from flask import Flask, request, render_template

app = Flask(__name__)

pointers = []
variables = []
context = {}


@app.route('/', methods=['POST', 'GET'])
def main():
    if request.method == 'GET':
        file_names = os.listdir('templates1')
        return render_template('form1.html', file_names=file_names)
    elif request.method == 'POST':
        a = request.form['file_name']
        print(a)
        doc = docx.Document(f'templates1/{a}')
        all_paragraphs = doc.paragraphs
        for paragraph in all_paragraphs:
            if '{' in paragraph.text:
                a = (paragraph.text[paragraph.text.find('{') + 2: paragraph.text.find('}')])
                if a not in pointers:
                    pointers.append(a)
        print(pointers)
        return save()


def save():
    pointer1 = ''
    if len(pointers) >= 1:
        pointer1 = pointers[0]
        visibility1 = ''
    else:
        visibility1 = 'hidden'

    pointer2 = ''
    if len(pointers) >= 2:
        pointer2 = pointers[1]
        visibility2 = ''
    else:
        visibility2 = 'hidden'

    pointer3 = ''
    if len(pointers) >= 3:
        pointer3 = pointers[2]
        visibility3 = ''
    else:
        visibility3 = 'hidden'

    pointer4 = ''
    if len(pointers) >= 4:
        pointer4 = pointers[3]
        visibility4 = ''
    else:
        visibility4 = 'hidden'
    return render_template('form2.html', pointer1=pointer1, visibility1=visibility1,
                           pointer2=pointer2, visibility2=visibility2,
                           pointer3=pointer3, visibility3=visibility3,
                           pointer4=pointer4, visibility4=visibility4)


if __name__ == '__main__':
    app.run(port=8080, host='127.0.0.1')
