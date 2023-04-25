import os

import docx
from docxtpl import DocxTemplate
from flask import Flask, request, render_template

app = Flask(__name__)

pointers = []
variables = []
context = {}
filename = ''


@app.route('/', methods=['POST', 'GET'])
def main():
    if request.method == 'GET':
        file_names = os.listdir('templates1')
        return render_template('form1.html', file_names=file_names)
    elif request.method == 'POST':
        global filename
        filename = request.form['file_name']
        doc = docx.Document(f'templates1/{filename}')
        all_paragraphs = doc.paragraphs
        for paragraph in all_paragraphs:
            if '{' in paragraph.text:
                a = (paragraph.text[paragraph.text.find('{') + 2: paragraph.text.find('}')])
                if a not in pointers:
                    pointers.append(a)
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
    pointer5 = ''
    if len(pointers) >= 5:
        pointer5 = pointers[4]
        visibility5 = ''
    else:
        visibility5 = 'hidden'
    pointer6 = ''
    if len(pointers) >= 6:
        pointer6 = pointers[5]
        visibility6 = ''
    else:
        visibility6 = 'hidden'
    pointer7 = ''
    if len(pointers) >= 7:
        pointer7 = pointers[6]
        visibility7 = ''
    else:
        visibility7 = 'hidden'
    pointer8 = ''
    if len(pointers) >= 8:
        pointer8 = pointers[7]
        visibility8 = ''
    else:
        visibility8 = 'hidden'
    pointer9 = ''
    if len(pointers) >= 9:
        pointer9 = pointers[8]
        visibility9 = ''
    else:
        visibility9 = 'hidden'
    pointer10 = ''
    if len(pointers) >= 10:
        pointer10 = pointers[9]
        visibility10 = ''
    else:
        visibility10 = 'hidden'
    pointer11 = ''
    if len(pointers) >= 11:
        pointer11 = pointers[10]
        visibility11 = ''
    else:
        visibility11 = 'hidden'
    pointer12 = ''
    if len(pointers) >= 12:
        pointer12 = pointers[11]
        visibility12 = ''
    else:
        visibility12 = 'hidden'
    pointer13 = ''
    if len(pointers) >= 13:
        pointer13 = pointers[12]
        visibility13 = ''
    else:
        visibility13 = 'hidden'
    pointer14 = ''
    if len(pointers) >= 14:
        pointer14 = pointers[13]
        visibility14 = ''
    else:
        visibility14 = 'hidden'
    pointer15 = ''
    if len(pointers) >= 15:
        pointer15 = pointers[14]
        visibility15 = ''
    else:
        visibility15 = 'hidden'
    pointer16 = ''
    if len(pointers) >= 16:
        pointer16 = pointers[15]
        visibility16 = ''
    else:
        visibility16 = 'hidden'
    pointer17 = ''
    if len(pointers) >= 17:
        pointer17 = pointers[16]
        visibility17 = ''
    else:
        visibility17 = 'hidden'
    pointer18 = ''
    if len(pointers) >= 18:
        pointer18 = pointers[17]
        visibility18 = ''
    else:
        visibility18 = 'hidden'
    pointer19 = ''
    if len(pointers) >= 19:
        pointer19 = pointers[18]
        visibility19 = ''
    else:
        visibility19 = 'hidden'
    pointer20 = ''
    if len(pointers) >= 20:
        pointer20 = pointers[19]
        visibility20 = ''
    else:
        visibility20 = 'hidden'

    return render_template('form2.html', pointer1=pointer1, visibility1=visibility1,
                           pointer2=pointer2, visibility2=visibility2,
                           pointer3=pointer3, visibility3=visibility3,
                           pointer4=pointer4, visibility4=visibility4,
                           pointer5=pointer5, visibility5=visibility5,
                           pointer6=pointer6, visibility6=visibility6,
                           pointer7=pointer7, visibility7=visibility7,
                           pointer8=pointer8, visibility8=visibility8,
                           pointer9=pointer9, visibility9=visibility9,
                           pointer10=pointer10, visibility10=visibility10,
                           pointer11=pointer11, visibility11=visibility11,
                           pointer12=pointer12, visibility12=visibility12,
                           pointer13=pointer13, visibility13=visibility13,
                           pointer14=pointer14, visibility14=visibility14,
                           pointer15=pointer15, visibility15=visibility15,
                           pointer16=pointer16, visibility16=visibility16,
                           pointer17=pointer17, visibility17=visibility17,
                           pointer18=pointer18, visibility18=visibility18,
                           pointer19=pointer19, visibility19=visibility19,
                           pointer20=pointer20, visibility20=visibility20)


@app.route('/savefile', methods=['GET'])
def savefile():
    if len(request.args.get('variable1')) > 0:
        variables.append(request.args.get('variable1'))
    if len(request.args.get('variable2')) > 0:
        variables.append(request.args.get('variable2'))
    if len(request.args.get('variable3')) > 0:
        variables.append(request.args.get('variable3'))
    if len(request.args.get('variable4')) > 0:
        variables.append(request.args.get('variable4'))
    if len(request.args.get('variable5')) > 0:
        variables.append(request.args.get('variable5'))
    if len(request.args.get('variable6')) > 0:
        variables.append(request.args.get('variable6'))
    if len(request.args.get('variable7')) > 0:
        variables.append(request.args.get('variable7'))
    if len(request.args.get('variable8')) > 0:
        variables.append(request.args.get('variable8'))
    if len(request.args.get('variable9')) > 0:
        variables.append(request.args.get('variable9'))
    if len(request.args.get('variable10')) > 0:
        variables.append(request.args.get('variable10'))
    if len(request.args.get('variable11')) > 0:
        variables.append(request.args.get('variable11'))
    if len(request.args.get('variable12')) > 0:
        variables.append(request.args.get('variable12'))
    if len(request.args.get('variable13')) > 0:
        variables.append(request.args.get('variable13'))
    if len(request.args.get('variable14')) > 0:
        variables.append(request.args.get('variable14'))
    if len(request.args.get('variable15')) > 0:
        variables.append(request.args.get('variable15'))
    if len(request.args.get('variable16')) > 0:
        variables.append(request.args.get('variable16'))
    if len(request.args.get('variable17')) > 0:
        variables.append(request.args.get('variable17'))
    if len(request.args.get('variable18')) > 0:
        variables.append(request.args.get('variable18'))
    if len(request.args.get('variable19')) > 0:
        variables.append(request.args.get('variable19'))
    if len(request.args.get('variable20')) > 0:
        variables.append(request.args.get('variable20'))
    file_name = request.args.get('file_name')

    if len(pointers) >= 1 and len(variables) >= 1:
        context[pointers[0]] = variables[0]
    if len(pointers) >= 2 and len(variables) >= 2:
        context[pointers[1]] = variables[1]
    if len(pointers) >= 3 and len(variables) >= 3:
        context[pointers[2]] = variables[2]
    if len(pointers) >= 4 and len(variables) >= 4:
        context[pointers[3]] = variables[3]
    if len(pointers) >= 5 and len(variables) >= 5:
        context[pointers[4]] = variables[4]
    if len(pointers) >= 6 and len(variables) >= 6:
        context[pointers[5]] = variables[5]
    if len(pointers) >= 7 and len(variables) >= 7:
        context[pointers[6]] = variables[6]
    if len(pointers) >= 8 and len(variables) >= 8:
        context[pointers[7]] = variables[7]
    if len(pointers) >= 9 and len(variables) >= 9:
        context[pointers[8]] = variables[8]
    if len(pointers) >= 10 and len(variables) >= 10:
        context[pointers[9]] = variables[9]
    if len(pointers) >= 11 and len(variables) >= 11:
        context[pointers[10]] = variables[10]
    if len(pointers) >= 12 and len(variables) >= 12:
        context[pointers[11]] = variables[11]
    if len(pointers) >= 13 and len(variables) >= 13:
        context[pointers[12]] = variables[12]
    if len(pointers) >= 14 and len(variables) >= 14:
        context[pointers[13]] = variables[13]
    if len(pointers) >= 15 and len(variables) >= 15:
        context[pointers[14]] = variables[14]
    if len(pointers) >= 16 and len(variables) >= 16:
        context[pointers[15]] = variables[15]
    if len(pointers) >= 17 and len(variables) >= 17:
        context[pointers[16]] = variables[16]
    if len(pointers) >= 18 and len(variables) >= 18:
        context[pointers[17]] = variables[17]
    if len(pointers) >= 19 and len(variables) >= 19:
        context[pointers[18]] = variables[18]
    if len(pointers) >= 20 and len(variables) >= 20:
        context[pointers[19]] = variables[19]

    doc = DocxTemplate(f'templates1/{filename}')
    doc.render(context)
    doc.save(f'result/{file_name}')
    return 'OK'


if __name__ == '__main__':
    app.run(port=8080, host='127.0.0.1')
