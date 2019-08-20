#!/usr/bin/python3.5
import base64
import platform
import random
import sys
import tkinter
from datetime import date, datetime, timedelta
from tkinter import messagebox, ttk
from unicodedata import normalize

from PIL import Image, ImageTk, ImageFont

from ttkcalendar import Calendar


class FalseRoutine():
    ''' Classe que instancia objetos fakes '''
    def __init__ (self):
        pass
    
    def get(self, value=None):
        return value


class Widgets():
    ''' Classe que instancia os widgets nos formulários. '''
    def __init__(self, instance, color=''):
        self.instance = instance
        self.color = color

    def button(self, field, cmd, textwidth, height, row, col, colspan=0, fontwidth=9):
        ''' Inserir button em formulários. '''
        text = tkinter.Button(self.instance, text=field, width=textwidth, command=cmd, font=(fonte, fontwidth, 'bold'))
        if colspan:
            text.grid(row=row, column=col, columnspan=colspan)
        else:
            text.grid(row=row, column=col)
        return text

    def label(self, field, textwidth, row, col, colspan=1, rowspan=1, fg='', height=0, stick=None, fontwidth=10):
        ''' Inserir texto livre em formulários. '''
        label = tkinter.Label(self.instance, text=field, width=textwidth, font=(fonte, fontwidth))
        if self.color:
            label['bg'] = self.color
        if height:
            label['height'] = height
        if stick:
            label.grid(row=row, column=col, columnspan=colspan, rowspan=rowspan, sticky=stick)
        else:
            label.grid(row=row, column=col, columnspan=colspan, rowspan=rowspan)
        return label

    def textbox(self, field, textwidth, row, col, default='', show='', cmd='', fontwidth=10):
        ''' Inserir textbox em formulários. '''
        label = tkinter.Label(self.instance, text=field, font=(fonte, fontwidth))
        text = tkinter.Entry(self.instance, width=textwidth, font=(fonte, fontwidth))
        if self.color:
            label['bg'] = self.color
        if show:
            text['show'] = show
        if cmd:
            text.bind("<FocusOut>", cmd)
            #text['validate'] = 'focusout'
            #text['validatecommand'] = cmd
        label.grid(row=row, column=col, sticky='E')
        text.grid(row=row, column=col + 1, sticky='W')
        text.insert(0, default)
        return text

    def calendar(self, row, col, default='', rowspan=1, columnspan=1, command=False, fontwidth=10):
        ''' Inserir calendário em formulários. '''
        try:
            calendar = Calendar(self.instance, locale='pt_BR.utf8')
        except:
            calendar = Calendar(self.instance, locale='ptb_bra')
        if command:
            calendar.bind("<1>", command)
        calendar.grid(row=row, column=col, rowspan=rowspan, columnspan=columnspan, stick='W')
        return calendar

    def check(self, field, textwidth, checktext, row, col, selected=False, seek='', fontwidth=10):
        ''' Inserir checkbox em formulários. '''
        label = tkinter.Label(self.instance, text=field, font=(fonte, fontwidth))
        var = tkinter.IntVar()
        checkbox = tkinter.Checkbutton(self.instance, text=checktext,
                                       width=textwidth, variable=var, font=(fonte, fontwidth))
        # var = True
        if self.color:
            label['bg'] = self.color
            checkbox['bg'] = self.color
        label.grid(row=row, column=col, stick='E')
        checkbox.grid(row=row, column=col + 1, sticky='W')
        if selected:
            checkbox.select()
        else:
            checkbox.deselect()
        if seek:
            checkbox.bind("<FocusOut>", seek)
        return var

    def combobox(self, field, textwidth, combolist, row, col, default='', cmd='', seek='', fontwidth=10):
        ''' Inserir combobox em formulários. '''
        label = tkinter.Label(self.instance, text=field, font=(fonte, fontwidth))
        text = ttk.Combobox(self.instance, values=combolist, width=textwidth, font=(fonte, fontwidth))
        if self.color:
            label['bg'] = self.color
        if cmd:
            text.bind("<<ComboboxSelected>>", cmd)
        if seek:
            text.bind("<FocusOut>", seek)
        label.grid(row=row, column=col, stick='E')
        text.grid(row=row, column=col + 1, sticky='W')
        text.insert(0, default)
        return text

    def combobox_return(self, field, lista=[]):
        if field.get():
            value = ''
            for row in lista:
                if field.get().lower() == row.lower():
                    value = row
                    break
            if not value:
                for row in lista:
                    if field.get().lower() in row.lower():
                        value = row
                        break
            field.delete(0, 'end')
            field.insert(0, value)
        
    def listbox(self, field, textwidth, height, combolist, row, col, cmd=None, fontwidth=10):
        ''' Inserir combobox em formulários. '''
        label = tkinter.Label(self.instance, text=field, font=(fonte, fontwidth))
        text = tkinter.Listbox(self.instance, height=height, width=textwidth, font=(fonte, fontwidth))
        for rowl in combolist:
            text.insert('end', rowl)
        if self.color:
            label['bg'] = self.color
        if cmd:
            text.bind("<Double-1>", cmd)
        label.grid(row=row, column=col, stick='E')
        text.grid(row=row, column=col + 1, sticky='W')
        return text

    def grid(self, columns, headers, combolist, ordlist, height, row, col, colspan=1, rowspan=1, cmd=None, order=True):
        ''' Inserir gridbox em formulários. 
        columns é uma tupla.
        headers é um dicionário de dicionário.
        combolist é um dicionario de tupla.
        ordlist é a lista ordenada de preenchimento.'''

        def treeview_sort_column(tv, col, reverse):
            l = [(tv.set(k, col), k) for k in tv.get_children('')]
            indice = columns
            if 'format' in headers[indice[col]]:
                if headers[indice[col]]['format'] == 'float':
                    l.sort(key=lambda t: float(num_usa(t[0])), reverse=reverse)
                if headers[indice[col]]['format'] == 'int':
                    l.sort(key=lambda t: int(num_usa(t[0])), reverse=reverse)
                if headers[indice[col]]['format'] == 'date/time':
                    try:
                        l.sort(key=lambda t: datetime.strptime(t[0], '%d/%m/%Y %H:%M:%S'), reverse=reverse)
                    except:
                        l.sort(key=lambda t: datetime.strptime(t[0], '%d/%m/%Y %H:%M'), reverse=reverse)
                if headers[indice[col]]['format'] == 'date':
                    l.sort(key=lambda t: datetime.strptime(t[0], '%d/%m/%Y'), reverse=reverse)
            else:
                l.sort(key=lambda t: t[0].lower(), reverse=reverse)
                # l.sort(reverse=reverse)

            # rearrange items in sorted positions
            for index, (val, k) in enumerate(l):
                tv.move(k, '', index)

            # reverse sort next time
            tv.heading(col, command=lambda: \
                    treeview_sort_column(tv, col, not reverse))

        def orderable(event):
            region = text.identify("region", event.x, event.y)
            if region == "heading":
                cabec = {'0-60': 1}
                last = 60
                ind = 0
                for row in headers:
                    cabec[str(last + 1) + '-' + str(headers[row]['width'] + last)] = ind
                    ind += 1
                    last += headers[row]['width']
                for row in cabec:
                    verif = row.split('-')
                    if event.x >= int(verif[0]) and event.x <= int(verif[1]):
                        treeview_sort_column(text, cabec[row], grid_order[0])
                if grid_order[0]:
                    grid_order[0] = False
                else:
                    grid_order[0] = True

        text = ttk.Treeview(self.instance, columns=columns, height=height)
        style = ttk.Style()
        style.configure(".", font=(fonte, 10))
        style.configure("Treeview.Heading", font=(fonte, 10, 'bold'))
        text.heading('#0', text='Id', anchor='center')
        text.column('#0', anchor='w', width=60)
        if order:
            text.bind("<Button-1>", orderable)
        if cmd:
            text.bind("<Double-1>", cmd)
        colm = 0
        for rows in columns:
            text.heading(
                rows, 
                text=headers[rows]['text'], 
                anchor='center', 
            )
            anchor = 'w'
            if 'anchor' in headers[rows]:
                anchor = headers[rows]['anchor']
            text.column(rows, anchor=anchor, width=headers[rows]['width'])
            colm += 1
        for rows in ordlist:
            text.insert('', 'end', text=rows, values=combolist[rows])
        text.grid(row=row, column=col, columnspan=colspan, rowspan=rowspan, stick='E')
        grid_order = [True]
        return text
    
    def image(self, field, textwidth, imagefile, row, col, colspam=1, rowspan=1, bg='', imagewidth=None, cmd=None):
        ''' Inserir imagem em formulários. '''
        image = Image.open(imagefile)
        if imagewidth:
            image = image.resize(imagewidth, Image.ANTIALIAS)
        logo = ImageTk.PhotoImage(image)
        text = tkinter.Label(self.instance, image=logo)
        if bg:
            text['bg'] = bg
        if self.color:
            text['bg'] = self.color
        if textwidth:
            text['width'] = textwidth
        if cmd:
            text.bind("<Button-1>", cmd)
        text.grid(row=row, column=col, columnspan=colspam, rowspan=rowspan)
        text.photo = logo
        return text

    def geometry(self, x, y):
        return str(y) + 'x' + str(x) + '+' + str(int(510 - y / 2)) + '+172'


class Application():
    def __init__(self, param={}, menu=False):    
        self.root = tkinter.Tk()
        if 'title' in param.keys():
            self.root.title(param['title'])
        if 'backcolor' in param.keys():
            self.root['bg'] = param['backcolor']
        if 'geometry' in param.keys():
            self.root.geometry(param['geometry'])
        if menu:
            pass

    def destroy(self):
        self.root.quit()

    def form(self, config={}):
        self.config = config
        self.child = tkinter.Toplevel(self.root)
        self.child.title("Formulário")
        # self.child.resizable(0, 0)
        if 'dimension' in self.config.keys():
            self.child.geometry(self.config['dimension'])
        if 'title' in self.config.keys():
            self.child.title(self.config['title'])
        if 'color' in self.config.keys():
            self.child['bg'] = self.config['color']
        return self.child

    def mainwindow(self, param={}):
        self.root.mainloop()

    def menu(self, menus=[], opcoes=[]):
        # Configura o menu da Aplicação a ser construída.
        #   menus
        # title = Nome a ser exibido no cabeçalho do menu
        #   opcoes
        # title = Nome a ser exibido no item do menu
        # submenu = True ou False
        main_menu = tkinter.Menu(self.root)
        menu = []
        for row in menus:
            menu.append(tkinter.Menu(main_menu, tearoff=0))
        for row in opcoes:
            if 'command' in row.keys():
                menu[row['root']].add_command(label=row['title'], command=row['command'], font=(fonte, 10))
            else:
                if 'root' in row.keys():
                    menu[row['root']].add_cascade(label=row['title'], menu=menu[row['menu']], font=(fonte, 10))
                else:
                    main_menu.add_cascade(label=row['title'], menu=menu[row['menu']], font=(fonte, 10))
                
        '''
        main_menu = tkinter.Menu(self.root)
        for row in menus:
            command = tkinter.Menu(main_menu, tearoff=0)
            for row2 in opcoes:
                if row2['menu'] == row['title']:
                    if 'submenu' in row2.keys():
                        subcommand = tkinter.Menu(command, tearoff=0)
                        command.add_cascade(label=row2['title'], menu=subcommand)
                        for row3 in opcoes:
                            if row3['menu'] == row2['title']:
                                subcommand.add_command(label=row3['title'], command=row3['command'])
                    else:
                        command.add_command(label=row2['title'], command=row2['command'])
            main_menu.add_cascade(label=row['title'], menu=command)
        '''
        self.root.config(menu=main_menu)


class RHEnconder():
    def __init__(self, value=None):
        self.value = value

    def encode(self, key, clear):
        enc = []
        for i in range(len(clear)):
            key_c = key[i % len(key)]
            enc_c = chr((ord(clear[i]) + ord(key_c)) % 256)
            enc.append(enc_c)
        return base64.urlsafe_b64encode("".join(enc))

    def decode(self, key, enc):
        dec = []
        enc = base64.urlsafe_b64decode(enc)
        for i in range(len(enc)):
            key_c = key[i % len(key)]
            dec_c = chr((256 + ord(enc[i]) - ord(key_c)) % 256)
            dec.append(dec_c)
        return "".join(dec)
    

def _date(datefill):
    ActualDate = str(datetime.now())[0:10]
    datesplit = datefill.split('/')
    datereturn = ''
    if len(datesplit) == 3:
        datereturn = datesplit[2] + '-' + str.zfill(datesplit[1], 2) + '-' + str.zfill(datesplit[0], 2)
    elif len(datesplit) == 2:
        datereturn = ActualDate[0:4] + '-' + str.zfill(datesplit[1], 2) + '-' + str.zfill(datesplit[0], 2)
    elif len(datesplit) == 1 and datesplit[0]:
        datereturn = ActualDate[0:7] + '-' + str.zfill(datesplit[0], 2)
    return datereturn


def data_cmd(data_return):
    if data_return.get():
        resp = _date(data_return.get())
        data_return.delete(0, 'end')
        data_return.insert(0, data_brasil(resp))


def data_brasil(dia):
    # Converte a data do banco de dados para o formato dd/mm/aaaa
    # Usado para exibição em textos e formulários
    if dia:
        data_retorno = dia.split('-')
        resp = data_retorno[2].zfill(2) + '/' + data_retorno[1].zfill(2) + '/' + data_retorno[0]
    else:
        resp = ''
    return resp


def datahora_brasil(dia):
    # Converte a data do banco de dados para o formato dd/mm/aaaa hh:mm:ss
    # Usado para exibição em textos e formulários
    if dia:
        completo = dia.split(' ')
        data_retorno = completo[0].split('-')
        resp = data_retorno[2].zfill(2) + '/' + data_retorno[1].zfill(2) + '/' + data_retorno[0] + ' ' + completo[1][0:8]
    else:
        resp = ''
    return resp


def lastdaymonth(diaini):
    months = {
        1: 31, 2: 28, 3: 31, 4: 30, 5: 31, 6:30,
        7: 31, 8: 31, 9: 30, 10: 31, 11: 30, 12: 31
    }
    mes = int(str(diaini)[5:7])
    ano = int(str(diaini)[0:4])
    bis = ano % 4
    if bis == 0:
        months[2] = 29
    dia = months[mes]
    lastday = date(ano, mes, dia)
    return lastday


def num_brasil(valor):
    numero = valor.split('.')
    ts = 0
    milhar = numero[0]
    acum = ''
    while not ts:
        num = milhar[-3:] + acum
        if len(milhar) > 3:
            acum = '.' + milhar[-3:] + acum
            milhar = milhar[:-3]
        else:
            ts = 1
    if len(numero[1]) == 1:
        numero[1] += '0'
    final = num + ',' + numero[1]
    return final


def num_usa(val):
    valorfill = val.replace('.', ';')
    valorfill = valorfill.replace(',', ';')
    valorfill = valorfill.split(';')
    if len(valorfill) == 1:
        valorfill.append('0')
    valorfill = ''.join(valorfill[:-1]) + '.' + valorfill[-1:][0]
    return valorfill

def remover_acentos(txt):
    return normalize('NFKD', txt).encode('ASCII', 'ignore').decode('ASCII')

def generator(id, data_efetiva, validade):
    #numero 48 a 57 - 10
    #letras 65 a 90 - 26
    #validade em dias * (dig[1] + 1) Máx 360 com 5 casas (serve para data e período de vigência)
    compoe = []
    dig = str(id).zfill(3)
    values = int(dig[0]) + int(dig[1]) + int(dig[2])
    values -= int(dig[2])
    passwd = [
        '4KE3',
        '4716',
        '9TR8',
        '55PU',
        '4RIS',
        '1SYP',
        '61YS',
        '9L9N',
        '13V5',
        '38IS',
        '38EA',
        '6ED0',
        '8Q0S',
        '9M66',
        '9QJC',
        '8476',
        '17MN',
        '9BFP',
        '42B2',
        '2ZBE',
        '7SWP',
        '96RS',
        '1POR',
        '21UG',
        '1EEX',
        '2BQG',
        '9OLU',
        '4AP1',
        '2A8A',
        '5XYZ',
        '9K6Y',
        '0BZ8',
        '9B9Q',
        '8Q3O',
        '3LDL',
        '9BBU',
        '7WOW',
        '9Y2Z',
        '4LQD',
        '5BL3',
        '4W2P',
        '0V36',
        '4V7N',
        '2USS',
        '8WID',
        '96YA',
    ]
    code = ''
    for row in range(3):
        alpha = random.randrange(2)
        if alpha:
            code += chr(random.randrange(26) + 65)
        else:
            code += chr(random.randrange(10) + 48)
    ld = chr(int(((int(data_efetiva[5:7]) - 1) * 30 + int(data_efetiva[-2:])) / 361 * 25) + 65)
    return passwd[values][0:2] + data_efetiva[2:4] + passwd[values][2:4] + str.zfill(str((validade * int(dig[1])) + 15 * int(dig[0])), 4) + code + ld

def validator(id, data_efetiva, codigo):
    #numero 48 a 57 - 10
    #letras 65 a 90 - 26
    #validade em dias * (dig[1] + 1) Máx 360 com 5 casas (serve para data e período de vigência)
    compoe = []
    dig = str(id).zfill(3)
    values = int(dig[0]) + int(dig[1]) + int(dig[2])
    values -= int(dig[2])
    passwd = [
        '4KE3',
        '4716',
        '9TR8',
        '55PU',
        '4RIS',
        '1SYP',
        '61YS',
        '9L9N',
        '13V5',
        '38IS',
        '38EA',
        '6ED0',
        '8Q0S',
        '9M66',
        '9QJC',
        '8476',
        '17MN',
        '9BFP',
        '42B2',
        '2ZBE',
        '7SWP',
        '96RS',
        '1POR',
        '21UG',
        '1EEX',
        '2BQG',
        '9OLU',
        '4AP1',
        '2A8A',
        '5XYZ',
        '9K6Y',
        '0BZ8',
        '9B9Q',
        '8Q3O',
        '3LDL',
        '9BBU',
        '7WOW',
        '9Y2Z',
        '4LQD',
        '5BL3',
        '4W2P',
        '0V36',
        '4V7N',
        '2USS',
        '8WID',
        '96YA',
    ]
    code = ''
    for row in range(2):
        alpha = random.randrange(2)
        if alpha:
            code += chr(random.randrange(26) + 65)
        else:
            code += chr(random.randrange(10) + 48)
    ld = chr(int(((int(data_efetiva[5:7]) - 1) * 30 + int(data_efetiva[-2:])) / 361 * 25) + 65)
    periodo = '0000'
    tempo = 0.0
    if len(codigo) == 14:
        periodo = codigo[6:10]
        try:
            tempo = (int(periodo) - 15 * int(dig[0])) / int(dig[1])
            if tempo not in [30, 60, 90, 180, 360]:
                periodo = '0000'
                tempo = 0.0
        except:
            tempo = 0.0

    valor_correto = passwd[values][0:2] + data_efetiva[2:4] + passwd[values][2:4] + periodo + codigo[10:13] + ld
    #print([codigo, valor_correto])
    #print(len(valor_correto))
    if codigo == valor_correto:
        return [True, tempo]
    else:
        return [False, 0.0]

if platform.system() == 'Windows':
    fonte = 'Tahoma'
else:
    fonte = 'Helvetica'
if len(sys.argv) > 1:
    if sys.argv[1] == '-g':
        if sys.argv[3] == 'hoje':
            day = str(datetime.now().date())
        else:
            day = sys.argv[3]
        print(generator(int(sys.argv[2]), day, int(sys.argv[4])))       
#print(generator(648, '2018-04-27', 60))
#print(validator(688, '2018-03-02', '9Q18JC0330OT1E'))
