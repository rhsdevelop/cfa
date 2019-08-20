#!/usr/bin/python3.5
import base64
import csv
import os
import random
import smtplib
import sqlite3
import tkinter as tk
from datetime import date, datetime, timedelta
from email.message import EmailMessage
from tkinter import messagebox, ttk

from base import (Application, FalseRoutine, Widgets, _date, data_brasil,
                  data_cmd, datahora_brasil, generator, lastdaymonth,
                  num_brasil, num_usa, remover_acentos, validator)
from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive


def usuarios():
    def newreg():
        edit('new')

    def editreg():
        edit('edit')

    def removereg():
        edit('remove')

    def edit(mode):
        def exit():
            form_edit.destroy()
            Edit.focus()
        
        def updatethis():
            executethis = True
            
            if mode in ['new', 'edit']:
                mens = 'Tem certeza que deseja atualizar ' + _nome.get() + '?'
                erros = []
                # Validar se tem repetido.
                c.execute('SELECT Id FROM Usuarios WHERE Nome = "' + _nome.get().upper() + '" AND Id <> ' + _id.get())
                doc = c.fetchall()
                if doc:
                    erros.append('-Já existe outro usuário com o mesmo nome.')
                # Validar não vazios.
                values = {'Nome do usuário': '', 'Senha': '', 'Senha redigitada': ''}
                values['Nome do usuário'] = _nome.get()
                values['Senha'] = _senha1.get()
                values['Senha redigitada'] = _senha2.get()
                empty = []
                for row in sorted(values.keys()):
                    if values[row] == '':
                        empty.append(row)
                if empty:
                    if len(empty) == 1:
                        msgerro = '-O campo ' + empty[0] + ' não foi preenchido.'
                        erros.append(msgerro)
                    else:
                        msgerro = '-Os campos ' + empty[0]
                        for row in empty[1:-1]:
                            msgerro = msgerro + ', ' + row
                        msgerro = msgerro + ' e ' + empty[-1:][0] + ' não foram preenchidos.'
                        erros.append(msgerro)
                # Validar se as 2 senhas digitadas conferem
                if _senha1.get() != _senha2.get():
                    erros.append('-Senhas digitadas são diferentes.')
                if erros:
                    msg = 'Não é possível atualizar esse cadastro:\n\n' + str.join('\n', erros) + '\n\nRevise os dados e tente novamente.'
                    messagebox.showerror(title='Aviso', message=msg, parent=form_edit)
                    executethis = False
                    update.focus()
            else:
                mens = 'Tem certeza que deseja remover ' + _nome.get() + '?'
            if executethis:
                if messagebox.askyesno(title="Atualização de sistema", message=mens, parent=form_edit):
                    command = []
                    if mode == 'new':
                        command.append("INSERT INTO Usuarios VALUES (")
                        command.append(_id.get() + ", ")
                        command.append("'" + _nome.get().upper() + "', ")
                        command.append("'" + _senha1.get() + "', ")
                        command.append("'" + _obs.get() + "'")
                        command.append(")")
                        c.execute(str.join('', command))
                        command = []
                        command.append("INSERT INTO Usuarios_acessos VALUES (")
                        command.append(_id.get() + ", ")
                        command.append("1, ")
                        if _id.get() == '1':
                            command.append("1, ")
                            command.append("4, ")
                            command.append("4, ")
                            command.append("4, ")
                            command.append("-1")
                        else:
                            command.append("0, ")
                            command.append("0, ")
                            command.append("0, ")
                            command.append("0, ")
                            command.append("-1")
                        command.append(")")
                    if mode == 'edit':
                        command.append("UPDATE Usuarios SET ")
                        command.append("Nome = '" + _nome.get().upper() + "', ")
                        command.append("Senha = '" + _senha1.get() + "', ")
                        command.append("Observacao = '" + _obs.get() + "' ")
                        command.append("WHERE Id = " + _id.get())
                    if mode == 'remove':
                        command.append('DELETE FROM Usuarios WHERE Id = ' + _id.get())
                        c.execute(str.join('', command))
                        command = []
                        command.append('DELETE FROM Usuarios_acessos WHERE Id = ' + _id.get())
                    c.execute(str.join('', command))
                    conn.commit()
                    nomeselect.delete(0, 'end')
                    c.execute('SELECT Nome FROM Usuarios ORDER BY Nome')
                    doc = c.fetchall()
                    lista = []
                    for row in doc:
                        lista.append(row[0])
                    for row in lista:
                        nomeselect.insert('end', row)
                    nomeselect.focus()
                    exit()
                else:
                    update.focus()

        dimension = widgets.geometry(240, 460)
        config = {'title': 'Edição de Usuários',
                  'dimension': dimension,
                  'color': 'pale goldenrod'}
        form_edit = main.form(config)
        wgedit = Widgets(form_edit, 'pale goldenrod')
        wgedit.label('', 5, 0, 0)
        wgedit.label('', 5, 1, 0, rowspan=6, height=10)
        wgedit.label('', 5, 6, 0)
        wgedit.label('', 5, 14, 0)
        if mode in ['edit', 'remove']:
            c.execute('SELECT * FROM Usuarios WHERE Nome = "' + nomeselect.get('active') + '"')
            doc = c.fetchall()
        _id = wgedit.textbox('Id ', 1, 1, 1)
        _nome = wgedit.textbox('Usuário do sistema ', 30, 2, 1)
        _senha1 = wgedit.textbox('Digite a senha ', 18, 3, 1, show='*')
        _senha2 = wgedit.textbox('Repita a senha ', 18, 4, 1, show='*')
        _obs = wgedit.textbox('Observação ', 30, 5, 1)

        if mode in ['edit', 'remove']:
            _id.insert(0, str(doc[0][0]))
            _nome.insert(0, str(doc[0][1]))
            _senha1.insert(0, str(doc[0][2]))
            _senha2.insert(0, str(doc[0][2]))
            _obs.insert(0, str(doc[0][3]))
        else:
            c.execute('SELECT Id FROM Usuarios ORDER BY Id')
            rows = c.fetchall()
            NewId = 1
            if rows:
                Ids = []
                for row in rows:
                    Ids.append(row[0])
                Ids.sort() 
                NewId = Ids[-1:][0] + 1
            _id.insert(0, str(NewId))

        _id.configure(state='disabled')
        if mode in ['edit', 'new']:
            _nome.focus()
            update = wgedit.button('Atualizar', updatethis, 20, 0, 7, 1, 2)
            cancel = wgedit.button('Cancelar', exit, 20, 0, 8, 1, 2)
        else:
            _nome.configure(state='disabled')
            _senha1.configure(state='disabled')
            _senha2.configure(state='disabled')
            _obs.configure(state='disabled')
            wgedit.label('Registro que será excluído!', 0, 7, 1, 2)
            update = wgedit.button('Confirmar exclusão', updatethis, 20, 0, 8, 1, 2)
            update.focus()
    
    dimension = widgets.geometry(300, 560)
    config = {'title': 'Selecione um Banco e a ação',
              'dimension': dimension,
              'color': 'light goldenrod'}
    form_selec = main.form(config)
    wg = Widgets(form_selec, 'light goldenrod')
    wg.label('', 5, 0, 0)
    wg.label('', 5, 2, 0)
    wg.label('', 5, 6, 0)
    wg.image('',0,'images/usuario.png', 1, 1, 1, 6)
    Create = wg.button('Incluir', newreg, 10, 0, 3, 3, 2)
    Edit = wg.button('Editar', editreg, 10, 0, 4, 3, 2)
    Remove = wg.button('Excluir', removereg, 10, 0, 5, 3, 2)
    if userauth:
        if userauth[3]:
            c.execute('SELECT Nome FROM Usuarios ORDER BY Nome')
        else:
            c.execute('SELECT Nome FROM Usuarios WHERE Id = ' + str(userauth[0]))
            Create.configure(state='disabled')
            Remove.configure(state='disabled')
    else:
        c.execute('SELECT Nome FROM Usuarios ORDER BY Nome')
    doc = c.fetchall()
    lista = []
    for row in doc:
        lista.append(row[0])
    nomeselect = wg.listbox('Nome:  ', 20, 8, lista, 1, 3)
    nomeselect.focus()

def bancos(tipo):
    def newreg():
        edit('new')

    def editreg():
        edit('edit')

    def removereg():
        edit('remove')

    def edit(mode):
        def exit():
            form_edit.destroy()
            bancoselect.focus()
        
        def updatethis():
            executethis = True
            
            if mode in ['new', 'edit']:
                mens = 'Tem certeza que deseja atualizar ' + _nomebanco.get() + '?'
                erros = []
                # Validar se tem repetido.
                c.execute('SELECT Id FROM Bancos WHERE NomeBanco = "' + _nomebanco.get() + '" AND Id <> ' + _id.get())
                doc = c.fetchall()
                if doc:
                    erros.append('-Já existe outro banco com o mesmo nome.')
                # Validar não vazios.
                values = {'Nome do banco': '', 'Tipo de movimento': '', 'Tipo de conta': ''}
                values['Nome do banco'] = _nomebanco.get()
                values['Tipo de movimento'] = _tipomov.get()
                values['Tipo de conta'] = _tipoconta.get()
                empty = []
                for row in sorted(values.keys()):
                    if values[row] == '':
                        empty.append(row)
                if empty:
                    if len(empty) == 1:
                        msgerro = '-O campo ' + empty[0] + ' não foi preenchido.'
                        erros.append(msgerro)
                    else:
                        msgerro = '-Os campos ' + empty[0]
                        for row in empty[1:-1]:
                            msgerro = msgerro + ', ' + row
                        msgerro = msgerro + ' e ' + empty[-1:][0] + ' não foram preenchidos.'
                        erros.append(msgerro)
                if erros:
                    msg = 'Não é possível atualizar esse cadastro:\n\n' + str.join('\n', erros) + '\n\nRevise os dados e tente novamente.'
                    messagebox.showerror(title='Aviso', message=msg, parent=form_edit)
                    executethis = False
                    update.focus()
            else:
                mens = 'Tem certeza que deseja remover ' + _nomebanco.get() + '?'
                c.execute('SELECT Diario.Id FROM Diario JOIN Bancos ON Bancos.Id = Banco WHERE Bancos.NomeBanco = "' + _nomebanco.get() + '"')
                doc = c.fetchall()
                if doc:
                    messagebox.showerror(title='Aviso', message='Você não pode apagar um banco com movimento registrado!', parent=form_edit)
                    executethis = False
                    update.focus()
            if executethis:
                if messagebox.askyesno(title="Atualização de sistema", message=mens, parent=form_edit):
                    command = []
                    if _diavenc.get() == '':
                        _diavenc.insert(0, '0')
                    if mode == 'new':
                        command.append("INSERT INTO Bancos VALUES (")
                        command.append(_id.get() + ", ")
                        command.append("'" + ActualDate + "', ")
                        command.append("'" + _nomebanco.get() + "', ")
                        command.append(str(lista1.index(_tipomov.get())) + ", ")
                        command.append("'" + _numero.get() + "', ")
                        command.append(_diavenc.get() + ", ")
                        command.append(str(int(_gerafatura.get())) + ", ")
                        command.append("'" + _agencia.get() + "', ")
                        command.append("'" + _conta.get() + "', ")
                        command.append(str(lista2.index(_tipoconta.get())))
                        command.append(")")
                    if mode == 'edit':
                        command.append("UPDATE Bancos SET ")
                        command.append("NomeBanco = '" + _nomebanco.get() + "', ")
                        command.append("TipoMov = " + str(lista1.index(_tipomov.get())) + ", ")
                        command.append("Numero = '" + _numero.get() + "', ")
                        command.append("DiaVenc = " + _diavenc.get() + ", ")
                        command.append("GeraFatura = " + str(int(_gerafatura.get())) + ", ")
                        command.append("Agencia = '" + _agencia.get() + "', ")
                        command.append("Conta = '" + _conta.get() + "', ")
                        command.append("TipoConta = " + str(lista2.index(_tipoconta.get())) + " ")
                        command.append("WHERE Id = " + _id.get())
                    if mode == 'remove':
                        command.append('DELETE FROM Bancos WHERE Id = ' + _id.get())
                    c.execute(str.join('', command))
                    conn.commit()
                    bancoselect.delete(0, 'end')
                    if tipo == 'mov':
                        c.execute('SELECT NomeBanco FROM Bancos WHERE TipoMov <= 1 ORDER BY NomeBanco')
                    elif tipo == 'cc':
                        c.execute('SELECT NomeBanco FROM Bancos WHERE TipoMov = 2 ORDER BY NomeBanco')
                    elif tipo == 'cpp':
                        c.execute('SELECT NomeBanco FROM Bancos WHERE TipoMov = 3 ORDER BY NomeBanco')
                    doc = c.fetchall()
                    lista = []
                    for row in doc:
                        lista.append(row[0])
                    for row in lista:
                        bancoselect.insert('end', row)
                    bancoselect.focus()
                    exit()
                else:
                    update.focus()

        def _tipomov_cmd(value=None):
            widgets.combobox_return(_tipomov, lista1)

        def _tipoconta_cmd(value=None):
            widgets.combobox_return(_tipoconta, lista2)

        dimension = widgets.geometry(308, 500)
        config = {'title': 'Edição de Banco',
                  'dimension': dimension,
                  'color': 'pale goldenrod'}
        form_edit = main.form(config)
        wgedit = Widgets(form_edit, 'pale goldenrod')
        wgedit.label('', 5, 0, 0)
        wgedit.label('', 5, 11, 0)
        wgedit.label('', 5, 14, 0)
        if mode in ['edit', 'remove']:
            c.execute('SELECT * FROM Bancos WHERE NomeBanco = "' + bancoselect.get('active') + '"')
            doc = c.fetchall()
        _id = wgedit.textbox('Id ', 2, 1, 1)
        _datacadastro = wgedit.textbox('Data do cadastro ', 12, 2, 1)
        _nomebanco = wgedit.textbox('Nome do banco ', 30, 3, 1)
        lista1 = ['Dinheiro em mãos', 'Conta bancária', 'Cartão de crédito', 'Cartão pré-pago']
        _tipomov = wgedit.combobox('Tipo de movimento ', 15, lista1, 4, 1, cmd=_tipomov_cmd, seek=_tipomov_cmd)
        _numero = wgedit.textbox('Número do cartão ', 18, 5, 1)
        _diavenc = wgedit.textbox('Vencimento da fatura ', 3, 6, 1)
        if mode in ['edit', 'remove']:
            if doc[0][6]:
                _gerafatura = wgedit.check('Gera fatura automática? ', 15, ' Gerar                        ', 7, 1, True)
            else:
                _gerafatura = wgedit.check('Gera fatura automática? ', 15, ' Gerar                        ', 7, 1, False)
        else:
            _gerafatura = wgedit.check('Gera fatura automática? ', 15, ' Gerar                        ', 7, 1, False)
        _agencia = wgedit.textbox('Código da agência ', 15, 8, 1)
        _conta = wgedit.textbox('Número da conta ', 15, 9, 1)
        lista2 = ['Conta corrente', 'Conta poupança ', 'CDB/Outras operações']
        _tipoconta = wgedit.combobox('Tipo da conta bancária ', 18, lista2, 10, 1, cmd=_tipoconta_cmd, seek=_tipoconta_cmd)

        if mode in ['edit', 'remove']:
            _id.insert(0, str(doc[0][0]))
            _datacadastro.insert(0, data_brasil(doc[0][1]))
            _nomebanco.insert(0, str(doc[0][2]))
            _tipomov.insert(0, lista1[doc[0][3]])
            _numero.insert(0, str(doc[0][4]))
            _diavenc.insert(0, str(doc[0][5]))
            _agencia.insert(0, str(doc[0][7]))
            _conta.insert(0, str(doc[0][8]))
            _tipoconta.insert(0, lista2[doc[0][9]])
        else:
            c.execute('SELECT Id FROM Bancos ORDER BY Id')
            rows = c.fetchall()
            NewId = 1
            if rows:
                Ids = []
                for row in rows:
                    Ids.append(row[0])
                Ids.sort() 
                NewId = Ids[-1:][0] + 1
            _id.insert(0, str(NewId))
            _datacadastro.insert(0, data_brasil(ActualDate))

        _id.configure(state='disabled')
        _datacadastro.configure(state='disabled')
        if mode in ['edit', 'new']:
            _nomebanco.focus()
            update = wgedit.button('Atualizar', updatethis, 20, 0, 12, 1, 2)
            cancel = wgedit.button('Cancelar', exit, 20, 0, 13, 1, 2)
        else:
            _nomebanco.configure(state='disabled')
            _tipomov.configure(state='disabled')
            _numero.configure(state='disabled')
            _diavenc.configure(state='disabled')
            _agencia.configure(state='disabled')
            _conta.configure(state='disabled')
            _tipoconta.configure(state='disabled')
            wgedit.label('Registro que será excluído!', 0, 12, 1, 2)
            update = wgedit.button('Confirmar exclusão', updatethis, 20, 0, 13, 1, 2)
    
    dimension = widgets.geometry(330, 580)
    config = {'title': 'Selecione um Banco e a ação',
              'dimension': dimension,
              'color': 'light goldenrod'}
    form_selec = main.form(config)
    wg = Widgets(form_selec, 'light goldenrod')
    wg.label('', 5, 0, 0)
    wg.label('', 5, 2, 0)
    wg.label('', 5, 6, 0)
    if tipo == 'mov':
        wg.image('',0,'images/bancos.png', 1, 1, 1, 6)
        c.execute('SELECT NomeBanco FROM Bancos WHERE TipoMov <= 1 ORDER BY NomeBanco')
    elif tipo == 'cc':
        wg.image('',0,'images/creditcard.png', 1, 1, 1, 6)
        c.execute('SELECT NomeBanco FROM Bancos WHERE TipoMov = 2 ORDER BY NomeBanco')
    elif tipo == 'cpp':
        wg.image('',0,'images/creditcard.png', 1, 1, 1, 6)
        c.execute('SELECT NomeBanco FROM Bancos WHERE TipoMov = 3 ORDER BY NomeBanco')
    doc = c.fetchall()
    lista = []
    for row in doc:
        lista.append(row[0])
    bancoselect = wg.listbox('Banco:  ', 20, 8, lista, 1, 3)
    bancoselect.focus()
    Create = wg.button('Incluir', newreg, 10, 0, 3, 3, 2)
    Edit = wg.button('Editar', editreg, 10, 0, 4, 3, 2)
    Remove = wg.button('Excluir', removereg, 10, 0, 5, 3, 2)

def bancos_mov():
    bancos('mov')

def bancos_cc():
    bancos('cc')

def bancos_cpp():
    bancos('cpp')

def categorias(in_value=False, out_value=False):
    def newreg():
        edit('new')

    def editreg():
        edit('edit')

    def removereg():
        edit('remove')

    def edit(mode):
        def exit():
            form_edit.destroy()
            itemselect.focus()
        
        def updatethis():
            executethis = True
            if mode in ['new', 'edit']:
                mens = 'Tem certeza que deseja atualizar ' + _categoria.get() + '?'
                erros = []
                # Validar se tem repetido.
                c.execute('SELECT Id FROM Categorias WHERE Categoria = "' + _categoria.get() + '" AND Id <> ' + _id.get())
                doc = c.fetchall()
                if doc:
                    erros.append('-Já existe categoria com o mesmo nome.')
                # Validar não vazios.
                values = {'Categoria': '', 'Tipo de movimento': ''}
                values['Categoria'] = _categoria.get()
                values['Tipo de movimento'] = _tipomov.get()
                empty = []
                for row in sorted(values.keys()):
                    if values[row] == '':
                        empty.append(row)
                if empty:
                    if len(empty) == 1:
                        msgerro = '-O campo ' + empty[0] + ' não foi preenchido.'
                        erros.append(msgerro)
                    else:
                        msgerro = '-Os campos ' + empty[0]
                        for row in empty[1:-1]:
                            msgerro = msgerro + ', ' + row
                        msgerro = msgerro + ' e ' + empty[-1:][0] + ' não foram preenchidos.'
                        erros.append(msgerro)
                if erros:
                    msg = 'Não é possível atualizar esse cadastro:\n\n' + str.join('\n', erros) + '\n\nRevise os dados e tente novamente.'
                    messagebox.showerror(title='Aviso', message=msg, parent=form_edit)
                    update.focus()
                    executethis = False
            else:
                mens = 'Tem certeza que deseja remover ' + _categoria.get() + '?'
                c.execute('SELECT Diario.Id FROM Diario JOIN Categorias ON Categorias.Id = CategoriaMov WHERE Categorias.Categoria = "' + _categoria.get() + '"')
                doc = c.fetchall()
                if doc:
                    messagebox.showerror(title='Aviso', message='Você não pode apagar uma categoria com movimento já registrado!', parent=form_edit)
                    update.focus()
                    executethis = False
            if executethis:
                if messagebox.askyesno(title="Atualização de sistema", message=mens, parent=form_edit):
                    command = []
                    if mode == 'new':
                        command.append("INSERT INTO Categorias VALUES (")
                        command.append(_id.get() + ", ")
                        command.append("'" + _categoria.get() + "', ")
                        command.append(str(lista1.index(_tipomov.get())) + ", ")
                        command.append(str(int(_repete.get())))
                        command.append(")")
                    if mode == 'edit':
                        command.append("UPDATE Categorias SET ")
                        command.append("Categoria = '" + _categoria.get() + "', ")
                        command.append("TipoMov = " + str(lista1.index(_tipomov.get())) + ", ")
                        command.append("Repete = " + str(int(_repete.get())) + " ")
                        command.append("WHERE Id = " + _id.get())
                    if mode == 'remove':
                        command.append('DELETE FROM Categorias WHERE Id = ' + _id.get())
                    c.execute(str.join('', command))
                    conn.commit()
                    itemselect.delete(0, 'end')
                    c.execute('SELECT Categoria FROM Categorias WHERE ' + whereval + ' ORDER BY Categoria')
                    doc = c.fetchall()
                    lista = []
                    for row in doc:
                        lista.append(row[0])
                    for row in lista:
                        itemselect.insert('end', row)
                    itemselect.focus()
                    exit()
                else:
                    update.focus()

        def _tipomov_cmd(value=None):
            widgets.combobox_return(_tipomov, lista1)

        dimension = widgets.geometry(190, 490)
        config = {'title': 'Edição de Categoria',
                  'dimension': dimension,
                  'color': 'pale goldenrod'}
        form_edit = main.form(config)
        wgedit = Widgets(form_edit, 'pale goldenrod')
        wgedit.label('', 5, 0, 0)
        wgedit.label('', 5, 10, 0)
        wgedit.label('', 5, 13, 0)
        if mode in ['edit', 'remove']:
            c.execute('SELECT * FROM Categorias WHERE Categoria = "' + itemselect.get('active') + '"')
            doc = c.fetchall()
        _id = wgedit.textbox('Id ', 2, 1, 1)
        _categoria = wgedit.textbox('Categoria ', 30, 2, 1)
        lista1 = ['Receita', 'Despesa', 'Transferência']
        _tipomov = wgedit.combobox('Tipo de movimento ', 15, lista1, 3, 1, cmd=_tipomov_cmd, seek=_tipomov_cmd)
        if mode in ['edit', 'remove']:
            if doc[0][1]:
                _repete = wgedit.check('Classificar nas despesas ', 15, ' Contemplar               ', 4, 1, True)
            else:
                _repete = wgedit.check('Classificar nas despesas ', 15, ' Contemplar               ', 4, 1, False)
        else:
            _repete = wgedit.check('Classificar nas despesas ', 15, ' Contemplar               ', 4, 1, False)

        if mode in ['edit', 'remove']:
            _id.insert(0, str(doc[0][0]))
            _categoria.insert(0, str(doc[0][1]))
            _tipomov.insert(0, lista1[doc[0][2]])
        else:
            c.execute('SELECT Id FROM Categorias ORDER BY Id')
            rows = c.fetchall()
            NewId = 1
            if rows:
                Ids = []
                for row in rows:
                    Ids.append(row[0])
                Ids.sort() 
                NewId = Ids[-1:][0] + 1
            _id.insert(0, str(NewId))

        _id.configure(state='disabled')
        if mode in ['edit', 'new']:
            _categoria.focus()
            update = wgedit.button('Atualizar', updatethis, 20, 0, 11, 1, 2)
            cancel = wgedit.button('Cancelar', exit, 20, 0, 12, 1, 2)
        else:
            _categoria.configure(state='disabled')
            _tipomov.configure(state='disabled')
            wgedit.label('Registro que será excluído!', 0, 12, 1, 2)
            update = wgedit.button('Confirmar exclusão', updatethis, 20, 0, 13, 1, 2)

    def seek():
        if itemseek.get():
            nome = "'%" + itemseek.get() + "%'"
            c.execute('SELECT Categoria FROM Categorias WHERE Categoria LIKE ' + nome + ' AND ' + whereval + ' ORDER BY Categoria')
            itemselect.delete(0, 'end')
            for row in c.fetchall():
                itemselect.insert('end', row[0])
        else:
            c.execute('SELECT Categoria FROM Categorias WHERE ' + whereval + ' ORDER BY Categoria')
            doc = c.fetchall()
            itemselect.delete(0, 'end')
            lista = []
            for row in doc:
                itemselect.insert('end', row[0])
    
    dimension = widgets.geometry(330, 580)
    config = {'title': 'Selecione uma Categoria e a ação',
              'dimension': dimension,
              'color': 'light goldenrod'}
    form_selec = main.form(config)
    if in_value:
        whereval = 'TipoMov = 0'
    elif out_value:
        whereval = 'TipoMov > 0'
    wg = Widgets(form_selec, 'light goldenrod')
    if in_value:
        wg.image('',0,'images/receitas.png', 1, 1, 1, 5)
    else:
        wg.image('',0,'images/despesas.png', 1, 1, 1, 5)
    wg.label('', 5, 0, 0)
    wg.label('', 5, 2, 0)
    wg.label('', 5, 6, 0)
    itemseek = wg.textbox('   Digite: ', 20, 1, 3)
    c.execute('SELECT Categoria FROM Categorias WHERE ' + whereval + ' ORDER BY Categoria')
    doc = c.fetchall()
    lista = []
    for row in doc:
        lista.append(row[0])
    itemselect = wg.listbox('Nome:  ', 20, 8, lista, 3, 3)
    Seek = wg.button('Procurar', seek, 10, 0, 2, 3, 2)
    itemseek.focus()
    Create = wg.button('Incluir', newreg, 10, 0, 4, 3, 2)
    Edit = wg.button('Editar', editreg, 10, 0, 5, 3, 2)
    Remove = wg.button('Excluir', removereg, 10, 0, 6, 3, 2)

def categorias_receitas():
    categorias(in_value=True)

def categorias_despesas():
    categorias(out_value=True)

def parceiros():
    def newreg():
        edit('new')

    def editreg():
        edit('edit')

    def removereg():
        edit('remove')

    def edit(mode):
        def exit():
            form_edit.destroy()
            Seek.focus()
        
        def updatethis():
            executethis = True
            if mode in ['new', 'edit']:
                mens = 'Tem certeza que deseja atualizar ' + _nome.get() + '?'
                erros = []
                # Validar se tem repetido.
                c.execute('SELECT Id FROM Parceiros WHERE Nome = "' + _nome.get() + '" AND Id <> ' + _id.get())
                doc = c.fetchall()
                if doc:
                    erros.append('-Já existe fornecedor com o mesmo nome.')
                # Validar não vazios.
                values = {'Nome': '', 'Nome Completo': '', 'Tipo': '', 'Endereco': ''}
                values['Nome'] = _nome.get()
                values['Nome Completo'] = _nomecompleto.get()
                values['Tipo'] = _tipo.get()
                values['Endereco'] = _endereco.get()
                empty = []
                for row in sorted(values.keys()):
                    if values[row] == '':
                        empty.append(row)
                if empty:
                    if len(empty) == 1:
                        msgerro = '-O campo ' + empty[0] + ' não foi preenchido.'
                        erros.append(msgerro)
                    else:
                        msgerro = '-Os campos ' + empty[0]
                        for row in empty[1:-1]:
                            msgerro = msgerro + ', ' + row
                        msgerro = msgerro + ' e ' + empty[-1:][0] + ' não foram preenchidos.'
                        erros.append(msgerro)
                if erros:
                    msg = 'Não é possível atualizar esse cadastro:\n\n' + str.join('\n', erros) + '\n\nRevise os dados e tente novamente.'
                    messagebox.showerror(title='Aviso', message=msg, parent=form_edit)
                    update.focus()
                    executethis = False
            else:
                mens = 'Tem certeza que deseja remover ' + _nome.get() + '?'
                c.execute('SELECT Diario.Id FROM Diario JOIN Parceiros ON Parceiros.Id = Parceiro WHERE Parceiros.Nome = "' + _nome.get() + '"')
                doc = c.fetchall()
                if doc:
                    messagebox.showerror(title='Aviso', message='Você não pode apagar um parceiro que possui movimento registrado!', parent=form_edit)
                    update.focus()
                    executethis = False
            if executethis:
                if messagebox.askyesno(title="Atualização de sistema", message=mens):
                    command = []
                    if mode == 'new':
                        command.append("INSERT INTO Parceiros VALUES (")
                        command.append(_id.get() + ", ")
                        command.append("'" + ActualDate + "', ")
                        command.append("'" + _nome.get() + "', ")
                        command.append("'" + _nomecompleto.get() + "', ")
                        command.append(str(lista1.index(_tipo.get())) + ", ")
                        command.append("'" + _doc.get() + "', ")
                        command.append("'" + _endereco.get() + "', ")
                        command.append("'" + _telefone.get() + "', ")
                        command.append("'" + _observacao.get() + "', ")
                        command.append(str(lista2.index(_modo.get())))
                        command.append(")")
                    if mode == 'edit':
                        command.append("UPDATE Parceiros SET ")
                        command.append("Nome = '" + _nome.get() + "', ")
                        command.append("NomeCompleto = '" + _nomecompleto.get() + "', ")
                        command.append("Tipo = " + str(lista1.index(_tipo.get())) + ", ")
                        command.append("Doc = '" + _doc.get() + "', ")
                        command.append("Endereco = '" + _endereco.get() + "', ")
                        command.append("Telefone = '" + _telefone.get() + "', ")
                        command.append("Observacao = '" + _observacao.get() + "', ")
                        command.append("Modo = " + str(lista2.index(_modo.get())) + " ")
                        command.append("WHERE Id = " + _id.get())
                    if mode == 'remove':
                        command.append('DELETE FROM Parceiros WHERE Id = ' + _id.get())
                    c.execute(str.join('', command))
                    conn.commit()
                    parceiroselect.delete(0, 'end')
                    c.execute('SELECT Nome FROM Parceiros ORDER BY Nome')
                    doc = c.fetchall()
                    lista = []
                    for row in doc:
                        lista.append(row[0])
                    for row in lista:
                        parceiroselect.insert('end', row)
                    parceiroselect.focus()
                    exit()
                else:
                    update.focus()

        def _tipo_cmd(value=None):
            widgets.combobox_return(_tipo, lista1)

        def _modo_cmd(value=None):
            widgets.combobox_return(_modo, lista2)

        dimension = widgets.geometry(388, 520)
        config = {'title': 'Edição de Parceiro',
                  'dimension': dimension,
                  'color': 'pale goldenrod'}
        form_edit = main.form(config)
        wgedit = Widgets(form_edit, 'pale goldenrod')
        wgedit.label('\n' * 15, 5, 0, 0, rowspan=10)
        wgedit.label('', 5, 11, 0)
        wgedit.label('', 5, 13, 0)
        if mode in ['edit', 'remove']:
            c.execute('SELECT * FROM Parceiros WHERE Nome = "' + parceiroselect.get('active') + '"')
            doc = c.fetchall()
        _id = wgedit.textbox('Id ', 2, 1, 1)
        _datacadastro = wgedit.textbox('Data do cadastro ', 12, 2, 1)
        _nome = wgedit.textbox('Nome ', 15, 3, 1)
        _nomecompleto = wgedit.textbox('Nome Completo ', 40, 4, 1)
        lista1 = ['Pessoa Física', 'Pessoa Jurídica']
        _tipo = wgedit.combobox('Tipo de parceiro ', 15, lista1, 5, 1, cmd=_tipo_cmd, seek=_tipo_cmd)
        _doc = wgedit.textbox('Número CPF/CNPJ ', 18, 6, 1)
        _endereco = wgedit.textbox('Endereço ', 40, 7, 1)
        _telefone = wgedit.textbox('Telefone ', 15, 8, 1)
        _observacao = wgedit.textbox('Observação ', 40, 9, 1)
        lista2 = ['Ambos', 'Cliente', 'Fornecedor']
        _modo = wgedit.combobox('Tipo de parceiro ', 15, lista2, 10, 1, cmd=_modo_cmd, seek=_modo_cmd)

        if mode in ['edit', 'remove']:
            _id.insert(0, str(doc[0][0]))
            _datacadastro.insert(0, data_brasil(doc[0][1]))
            _nome.insert(0, str(doc[0][2]))
            _nomecompleto.insert(0, str(doc[0][3]))
            _tipo.insert(0, lista1[doc[0][4]])
            _doc.insert(0, str(doc[0][5]))
            _endereco.insert(0, str(doc[0][6]))
            _telefone.insert(0, str(doc[0][7]))
            _observacao.insert(0, str(doc[0][8]))
            _modo.insert(0, lista2[doc[0][9]])
        else:
            c.execute('SELECT Id FROM Parceiros ORDER BY Id')
            rows = c.fetchall()
            NewId = 1
            if rows:
                Ids = []
                for row in rows:
                    Ids.append(row[0])
                Ids.sort() 
                NewId = Ids[-1:][0] + 1
            _id.insert(0, str(NewId))
            _datacadastro.insert(0, data_brasil(ActualDate))

        _id.configure(state='disabled')
        _datacadastro.configure(state='disabled')
        if mode in ['edit', 'new']:
            _nome.focus()
            update = wgedit.button('Atualizar', updatethis, 20, 0, 12, 1, 2)
            cancel = wgedit.button('Cancelar', exit, 20, 0, 14, 1, 2)
        else:
            _nome.configure(state='disabled')
            _nomecompleto.configure(state='disabled')
            _tipo.configure(state='disabled')
            _doc.configure(state='disabled')
            _endereco.configure(state='disabled')
            _telefone.configure(state='disabled')
            _observacao.configure(state='disabled')
            _modo.configure(state='disabled')
            wgedit.label('Registro que será excluído!', 0, 12, 1, 2)
            update = wgedit.button('Confirmar exclusão', updatethis, 20, 0, 13, 1, 2)

    def seek():
        if parceiroseek.get():
            nome = "'%" + parceiroseek.get() + "%'"
            c.execute('SELECT Nome FROM Parceiros WHERE Nome LIKE ' + nome + ' ORDER BY Nome')
            parceiroselect.delete(0, 'end')
            for row in c.fetchall():
                parceiroselect.insert('end', row[0])
        else:
            c.execute('SELECT Nome FROM Parceiros ORDER BY Nome')
            doc = c.fetchall()
            parceiroselect.delete(0, 'end')
            lista = []
            for row in doc:
                parceiroselect.insert('end', row[0])
    
    dimension = widgets.geometry(330, 580)
    config = {'title': 'Selecione um Parceiro e a ação',
              'dimension': dimension,
              'color': 'light goldenrod'}
    form_selec = main.form(config)
    wg = Widgets(form_selec, 'light goldenrod')
    wg.image('',0,'images/parceiros.png', 1, 1, 1, 5)
    wg.label('', 5, 0, 0)
    wg.label('', 5, 2, 0)
    wg.label('', 5, 6, 0)
    parceiroseek = wg.textbox('   Digite: ', 20, 1, 3)
    c.execute('SELECT Nome FROM Parceiros ORDER BY Nome')
    doc = c.fetchall()
    lista = []
    for row in doc:
        lista.append(row[0])
    Seek = wg.button('Procurar', seek, 10, 0, 2, 3, 2)
    parceiroselect = wg.listbox('Nome:  ', 20, 8, lista, 3, 3)
    Create = wg.button('Incluir', newreg, 10, 0, 4, 3, 2)
    Edit = wg.button('Editar', editreg, 10, 0, 5, 3, 2)
    Remove = wg.button('Excluir', removereg, 10, 0, 6, 3, 2)
    parceiroseek.focus()

def parceiros_listar():
    def generate_sql():
        adic = ''
        command = ('SELECT * ' +
                   'FROM Parceiros ' +
                   'ORDER BY Nome')
        c.execute(command)
        doc = c.fetchall()
        return doc

    def seek():
        doc = generate_sql()
        colsf = ['datacadastro', 'nome', 'nomecompleto', 'tipo', 'doc', 'endereco', 'telefone', 'observacao']
        headf = {
            'datacadastro': {'text': 'Data inserido', 'width': 70},
            'nome': {'text': 'Nome fantasia', 'width': 120},
            'nomecompleto': {'text': 'Nome completo', 'width': 150},
            'tipo': {'text': 'Tipo', 'width': 30},
            'doc': {'text': 'CPF/CNPJ', 'width': 137},
            'endereco': {'text': 'Endereço', 'width': 120},
            'telefone': {'text': 'Telefone', 'width': 120},
            'observacao': {'text': 'Obs', 'width': 150},
        }
        combolist = {}
        ordlista = []
        for row in doc:
            texts = []
            rowact = 0
            for rows in row[1:]:
                if rowact in [0]:
                    texts.append(data_brasil(rows))
                elif rowact in [3]:
                    if rows == 0:
                        texts.append('PF')
                    else:
                        texts.append('PJ')
                else:
                    texts.append(rows)
                rowact += 1
            ordlista.append(row[0])
            combolist[row[0]] = tuple(texts)
        itemselect = wg.grid(colsf, headf, combolist, ordlista, 23, 5, 1, 4)
        #for rows in ordlista:
        #    itemselect.insert('', 'end', text=rows, values=combolist[rows])

    dimension = widgets.geometry(487, 960)
    config = {'title': 'Relação de parceiros cadastrados',
              'dimension': dimension,
              'color': 'light goldenrod'}
    form_selec = main.form(config)
    wg = Widgets(form_selec, 'light goldenrod')
    seek()

def movimentos(in_value=False, out_value=False, crd_card=False):
    def exitsc():
        form_selec.destroy()

    def newreg():
        edit('new')

    def dupreg():
        edit('dup')

    def editreg():
        edit('edit')

    def pagreg():
        if itemselect.selection():
            for i in itemselect.selection():
                gridselected = str(itemselect.item(i, 'text'))
            mens = ('Tem certeza que deseja confirmar a baixa dos documentos selecionados?\n\n' +
                    'Os documentos serão baixados na data de hoje: ' + data_brasil(ActualDate))
            if in_value:
                title = 'Baixa de Recebimentos'
            elif out_value:
                title = 'Baixa de Pagamentos'
            else:
                if not crd_card:
                    title = 'Confirmação de transferências'
                else:
                    title = 'Confirmação de pagamentos'
            if messagebox.askyesno(title=title, message=mens, parent=form_selec):
                notrelease = []
                for i in itemselect.selection():
                    c.execute('SELECT DataPago, Diario.TipoMov, Banco, Fatura, Bancos.TipoMov, Bancos.GeraFatura FROM Diario '\
                        'JOIN Bancos ON Bancos.Id = Banco WHERE Diario.Id = ' + str(itemselect.item(i, 'text')))
                    opt = c.fetchone()
                    if not opt[0]:
                        command = []
                        command.append("UPDATE Diario SET ")
                        command.append("DataPago = '" + ActualDate + "' ")
                        command.append("WHERE Id = " + str(itemselect.item(i, 'text')))
                        c.execute(str.join('', command))
                        conn.commit()
                        if opt[1] == 3:
                            command = []
                            command.append("UPDATE Diario SET ")
                            command.append("DataPago = '" + ActualDate + "' ")
                            command.append("WHERE Id = " + str(itemselect.item(i, 'text') - 1))
                            c.execute(str.join('', command))
                            conn.commit()
                            if opt[4] == 2 and opt[5]:
                                opt_bank_v = str(opt[2])
                                c.execute('UPDATE Diario SET DataPago = "' + ActualDate + '" WHERE Fatura = "' + opt[3] + '" AND Banco = ' + opt_bank_v + '')
                                messagebox.showinfo('Aviso!', 'Todos os pagamentos referente à fatura ' + opt[3] + ' foram baixados no sistema.', parent=form_selec)
                            Sair.focus()
                        elif opt[1] == 4:
                            command = []
                            command.append("UPDATE Diario SET ")
                            command.append("DataPago = '" + ActualDate + "' ")
                            command.append("WHERE Id = " + str(itemselect.item(i, 'text') + 1))
                            c.execute(str.join('', command))
                            conn.commit()
                            if opt[3]:
                                c.execute('SELECT Banco, TipoMov, GeraFatura FROM Diario WHERE Id = ' + str(itemselect.item(i, 'text') + 1))
                                opt_bank = c.fetchone()
                                if opt_bank[1] == 2 and opt_bank[2]:
                                    c.execute('UPDATE Diario SET DataPago = "' + ActualDate + '" WHERE Fatura = "' + opt[3] + '" AND Banco = ' + str(opt_bank[0]) + '')
                                    messagebox.showinfo('Aviso!', 'Todos os pagamentos referente à fatura ' + opt[3] + ' foram baixados no sistema.', parent=form_selec)
                                Sair.focus()
                    else:
                        notrelease.append(str(itemselect.item(i, 'text')))
                if notrelease:
                    if len(notrelease) < 2:
                        message = 'O documento ' + notrelease[0] + ' já foi baixado! Revise.'
                    else:
                        message = 'Os documentos ' 
                        for i in notrelease:
                            if i not in notrelease[-2:]:
                                message = message + str(i) + ', '
                            elif i == notrelease[-2]:
                                message = message + str(i) + ' e '
                            else:
                                message = message + str(i) + ' '
                        message = message + ' já foram baixados! Revise.'
                    messagebox.showerror(title='Aviso', message=message, parent=form_selec)
                    Sair.focus()
            else:
                Sair.focus()
        else:
            messagebox.showerror(title='Aviso', message='Não foi selecionado nenhum documento para a baixa!', parent=form_selec)
            Sair.focus()
    
    def removereg():
        edit('remove')

    def edit(mode):
        def exit():
            form_edit.destroy()
            Create.focus()
        
        def updatethis():
            executethis = True
            if in_value or out_value: # atualiza recebimentos e pagamentos
                if mode in ['new', 'edit', 'dup']:
                    mens = 'Tem certeza que deseja atualizar ' + _categoria.get() + '?'
                    erros = []
                    # Validar não vazios.
                    values = {'Parceiro': '', 'Banco': '', 'Categoria': '', 'Data de lançamento': '', 'Vencimento': '', 'Valor': ''}
                    values['Parceiro'] = _parceiro.get()
                    values['Banco'] = _banco.get()
                    values['Categoria'] = _categoria.get()
                    values['Data de lançamento'] = _datadoc.get()
                    values['Vencimento'] = _datavenc.get()
                    values['Valor'] = _valor.get()
                    empty = []
                    for row in sorted(values.keys()):
                        if values[row] == '':
                            empty.append(row)
                    if empty:
                        if len(empty) == 1:
                            msgerro = '-O campo ' + empty[0] + ' não foi preenchido.'
                            erros.append(msgerro)
                        else:
                            msgerro = '-Os campos ' + empty[0]
                            for row in empty[1:-1]:
                                msgerro = msgerro + ', ' + row
                            msgerro = msgerro + ' e ' + empty[-1:][0] + ' não foram preenchidos.'
                            erros.append(msgerro)
                    valorfill = _valor.get().replace('.', '')
                    valorfill = valorfill.replace(',', '.')
                    try:
                        teste = float(valorfill)
                    except:
                        erros.append('-O valor preenchido está incorreto.')
                    if erros:
                        msg = 'Não é possível atualizar esse cadastro:\n\n' + str.join('\n', erros) + '\n\nRevise os dados e tente novamente.'
                        messagebox.showerror(title='Aviso', message=msg, parent=form_edit)
                        update.focus()
                        executethis = False
                else:
                    mens = 'Tem certeza que deseja remover ' + _categoria.get() + '?'
                if executethis:
                    if messagebox.askyesno(title="Atualização de sistema", message=mens, parent=form_edit):
                        command = []
                        c.execute('SELECT Fatura, Diario.Banco, Bancos.GeraFatura, Bancos.NomeBanco FROM Diario JOIN Bancos ON Bancos.Id = Diario.Banco WHERE Diario.Id = ' + _id.get())
                        resp_old = c.fetchone()
                        if mode in ['new', 'dup']:
                            command.append("INSERT INTO Diario VALUES (")
                            command.append(_id.get() + ", ")
                            command.append("'" + ActualDate + "', ")
                            command.append("'" + ActualDate + "', ")
                            command.append("'" + _date(_datadoc.get()) + "', ")
                            command.append("'" + _date(_datavenc.get()) + "', ")
                            command.append("'" + _date(_datapago.get()) + "', ")
                            c.execute('SELECT Id FROM Parceiros WHERE Nome = "' + _parceiro.get() + '"')
                            command.append(str(c.fetchone()[0]) + ", ")
                            c.execute('SELECT Id, TipoMov, GeraFatura, DiaVenc FROM Bancos WHERE NomeBanco = "' + _banco.get() + '"')
                            resp_banco = c.fetchone()
                            command.append(str(resp_banco[0]) + ", ")
                            command.append("'" + _fatura.get() + "', ")
                            command.append("'" + _descricao.get() + "', ")
                            if in_value:
                                command.append(valorfill + ', ')
                                command.append("0, ")
                            elif out_value:
                                if valorfill[0] == '-':
                                    command.append(valorfill[1:] + ', ')
                                else:
                                    command.append('-' + valorfill + ', ')
                                command.append("1, ")
                            else:
                                command.append("2, ")
                            c.execute('SELECT Id FROM Categorias WHERE Categoria = "' + _categoria.get() + '"')
                            command.append(str(c.fetchone()[0]))
                            command.append(")")
                        if mode == 'edit':
                            command.append("UPDATE Diario SET ")
                            command.append("DataLastUpdate = '" + ActualDate + "', ")
                            command.append("DataDoc = '" + _date(_datadoc.get()) + "', ")
                            c.execute('SELECT Id FROM Parceiros WHERE Nome = "' + _parceiro.get() + '"')
                            command.append("Parceiro = " + str(c.fetchone()[0]) + ", ")
                            c.execute('SELECT Id, TipoMov, GeraFatura, DiaVenc FROM Bancos WHERE NomeBanco = "' + _banco.get() + '"')
                            resp_banco = c.fetchone()
                            command.append("Banco = " + str(resp_banco[0]) + ", ")
                            c.execute('SELECT Id FROM Categorias WHERE Categoria = "' + _categoria.get() + '"')
                            command.append("CategoriaMov = " + str(c.fetchone()[0]) + ", ")
                            command.append("Fatura = '" + _fatura.get() + "', ")
                            command.append("Descricao = '" + _descricao.get() + "', ")
                            if in_value:
                                command.append("Valor = " + valorfill + ", ")
                            else:
                                if valorfill[0] == '-':
                                    command.append("Valor = " + valorfill[1:] + ", ")
                                else:
                                    command.append("Valor = -" + valorfill + ", ")
                            command.append("DataVenc = '" + _date(_datavenc.get()) + "', ")
                            command.append("DataPago = '" + _date(_datapago.get()) + "' ")
                            command.append("WHERE Id = " + _id.get())
                        if mode == 'remove':
                            c.execute('SELECT Id, TipoMov, GeraFatura, DiaVenc FROM Bancos WHERE NomeBanco = "' + _banco.get() + '"')
                            resp_banco = c.fetchone()
                            command.append('DELETE FROM Diario WHERE Id = ' + _id.get())
                        c.execute(str.join('', command))
                        # rotina criada em 20/04/2018 para atualizar faturas de cartão de crédito automaticamente.
                        if resp_banco[1] == 2 and resp_banco[2]:
                            c.execute('SELECT Id, Valor FROM Diario WHERE Banco = ' + str(resp_banco[0]) + ' AND Fatura = "' + _fatura.get() + '" AND TipoMov = 3')
                            resp_fat = c.fetchone()
                            if resp_fat:
                                c.execute('SELECT SUM(Valor) FROM Diario WHERE Banco = ' + str(resp_banco[0]) + ' AND Fatura = "' + _fatura.get() + '" AND TipoMov IN (0, 1)')
                                value_fat = c.fetchone()
                                if None not in value_fat:
                                    c.execute("UPDATE Diario SET "\
                                        "DataLastUpdate = '" + ActualDate + "', "\
                                        "Valor = " + str(round(-value_fat[0], 2)) + " "\
                                        "WHERE Id = " + str(resp_fat[0]) + ""\
                                    )
                                    c.execute("UPDATE Diario SET "\
                                        "DataLastUpdate = '" + ActualDate + "', "\
                                        "Valor = " + str(round(value_fat[0], 2)) + " "\
                                        "WHERE Id = " + str(resp_fat[0] - 1) + ""\
                                    )
                                    mensagem = 'A fatura do seu cartão ' + _banco.get() + ' do período ' + _fatura.get() + ' foi atualizada.'\
                                        '\nValor atual: ' + num_brasil(str(round(-value_fat[0], 2))) + '.'
                                    messagebox.showinfo('Aviso', mensagem, parent=form_edit)
                                else:
                                    c.execute("DELETE FROM Diario WHERE Id = " + str(resp_fat[0]))
                                    c.execute("DELETE FROM Diario WHERE Id = " + str(resp_fat[0] - 1))
                                    mensagem = 'A fatura do seu cartão ' + _banco.get() + ' do período ' + _fatura.get() + ' foi apagada.'\
                                        '\nNão existe mais nenhum movimento para essa fatura.'
                                    messagebox.showinfo('Aviso', mensagem, parent=form_edit)
                            else:
                                c.execute('SELECT SUM(Valor) FROM Diario WHERE Banco = ' + str(resp_banco[0]) + ' AND Fatura = "' + _fatura.get() + '" AND TipoMov IN (0, 1)')
                                value_fat = c.fetchone()
                                if value_fat:
                                    c.execute('SELECT Id FROM Diario ORDER BY Id')
                                    rows = c.fetchall()
                                    NewId = 0
                                    if rows:
                                        Ids = []
                                        for row in rows:
                                            Ids.append(row[0])
                                        Ids.sort() 
                                        NewId = Ids[-1:][0] + 1
                                    new_doc = _fatura.get() + '-01'
                                    new_venc = _fatura.get() + '-' + str(resp_banco[3]).zfill(2)
                                    for a in range(2):
                                        command = []
                                        command.append("INSERT INTO Diario VALUES (")
                                        command.append(str(NewId) + ", ")
                                        command.append("'" + ActualDate + "', ")
                                        command.append("'" + ActualDate + "', ")
                                        command.append("'" + new_doc + "', ")
                                        command.append("'" + new_venc + "', ")
                                        command.append("'', ")
                                        command.append("1, ")
                                        if a == 0:
                                            command.append("1, ")
                                        else:
                                            command.append(str(resp_banco[0]) + ", ")
                                        command.append("'" + _fatura.get() + "', ")
                                        command.append("'<CRED.CARD>', ")
                                        if a == 0:
                                            command.append(str(round(value_fat[0], 2)) + ', ')
                                            command.append("4, ")
                                        else:
                                            command.append(str(round(-value_fat[0], 2)) + ', ')
                                            command.append("3, ")
                                        c.execute('SELECT Id FROM Categorias WHERE Categoria = "TRANSFERENCIAS"')
                                        command.append(str(c.fetchone()[0]))
                                        command.append(")")
                                        c.execute(str.join('', command))
                                        NewId += 1

                                    mensagem = 'A fatura do seu cartão ' + _banco.get() + ' do período ' + _fatura.get() + ' foi atualizada.'\
                                        '\nPrimeiro evento da nova fatura! Verifique o banco usado para o pagamento.'\
                                        '\nValor atual: ' + num_brasil(str(round(-value_fat[0], 2))) + '.'
                                    messagebox.showinfo('Aviso', mensagem, parent=form_edit)
                        if resp_old:
                            if resp_old[2]:
                                if resp_old[1] == resp_banco[0] and resp_old[0] == _fatura.get():
                                    pass
                                else:
                                    c.execute('SELECT Id FROM Diario WHERE TipoMov = 3 AND Banco = ' + str(resp_old[1]) + ' AND Fatura = "' + str(resp_old[0]) + '"')
                                    resp_fat = c.fetchone()
                                    if None not in resp_fat:
                                        c.execute('SELECT SUM(Valor) FROM Diario WHERE Banco = ' + str(resp_old[1]) + ' AND Fatura = "' + str(resp_old[0]) + '" AND TipoMov IN (0, 1)')
                                        value_fat = c.fetchone()
                                        if None not in value_fat:
                                            c.execute("UPDATE Diario SET "\
                                                "DataLastUpdate = '" + ActualDate + "', "\
                                                "Valor = " + str(round(-value_fat[0], 2)) + " "\
                                                "WHERE Id = " + str(resp_fat[0]) + ""\
                                            )
                                            c.execute("UPDATE Diario SET "\
                                                "DataLastUpdate = '" + ActualDate + "', "\
                                                "Valor = " + str(round(value_fat[0], 2)) + " "\
                                                "WHERE Id = " + str(resp_fat[0] - 1) + ""\
                                            )
                                            mensagem = 'A fatura do seu cartão ' + resp_old[3] + ' do período ' + resp_old[0] + ' foi atualizada.'\
                                                '\nValor atual: ' + num_brasil(str(round(-value_fat[0], 2))) + '.'
                                            messagebox.showinfo('Aviso', mensagem, parent=form_edit)
                                        else:
                                            c.execute("DELETE FROM Diario WHERE Id = " + str(resp_fat[0]))
                                            c.execute("DELETE FROM Diario WHERE Id = " + str(resp_fat[0] - 1))
                                            mensagem = 'A fatura do seu cartão ' + resp_old[3] + ' do período ' + resp_old[0] + ' foi apagada.'\
                                                '\nNão existe mais nenhum movimento para essa fatura.'
                                            messagebox.showinfo('Aviso', mensagem, parent=form_edit)
                        # fim da rotina.
                        conn.commit()
                        seek()
                        exit()
                    else:
                        update.focus()
            else: # atualiza transferências
                if mode in ['new', 'edit', 'dup']:
                    mens = 'Tem certeza que deseja atualizar ' + _descricao.get() + '?'
                    erros = []
                    # Validar não vazios.
                    values = {'Data do movimento': '', 'Banco Origem': '', 'Banco Destino': '', 'Valor': '', 'Vencimento': ''}
                    values['Data do movimento'] = _datadoc.get()
                    values['Banco Origem'] = _bancoorigem.get()
                    values['Banco Destino'] = _bancodestino.get()
                    values['Data de lançamento'] = _datadoc.get()
                    values['Valor'] = _valor.get()
                    values['Vencimento'] = _datavenc.get()
                    empty = []
                    for row in sorted(values.keys()):
                        if values[row] == '':
                            empty.append(row)
                    if empty:
                        if len(empty) == 1:
                            msgerro = '-O campo ' + empty[0] + ' não foi preenchido.'
                            erros.append(msgerro)
                        else:
                            msgerro = '-Os campos ' + empty[0]
                            for row in empty[1:-1]:
                                msgerro = msgerro + ', ' + row
                            msgerro = msgerro + ' e ' + empty[-1:][0] + ' não foram preenchidos.'
                            erros.append(msgerro)
                    valorfill = _valor.get().replace('.', '')
                    valorfill = valorfill.replace(',', '.')
                    try:
                        teste = float(valorfill)
                    except:
                        erros.append('-O valor preenchido está incorreto.')
                    if values['Banco Origem'] == values['Banco Destino']:
                        erros.append('-Banco origem e banco destino idênticos.')
                    if erros:
                        msg = 'Não é possível atualizar esse cadastro:\n\n' + str.join('\n', erros) + '\n\nRevise os dados e tente novamente.'
                        messagebox.showerror(title='Aviso', message=msg, parent=form_edit)
                        update.focus()
                        executethis = False
                else:
                    mens = 'Tem certeza que deseja remover ' + _descricao.get() + '?'
                if executethis:
                    if messagebox.askyesno(title="Atualização de sistema", message=mens, parent=form_edit):
                        if mode in ['new', 'dup']:
                            NewId = int(_id.get())
                            for a in range(2):
                                command = []
                                command.append("INSERT INTO Diario VALUES (")
                                command.append(str(NewId) + ", ")
                                command.append("'" + ActualDate + "', ")
                                command.append("'" + ActualDate + "', ")
                                command.append("'" + _date(_datadoc.get()) + "', ")
                                command.append("'" + _date(_datavenc.get()) + "', ")
                                command.append("'" + _date(_datapago.get()) + "', ")
                                command.append("1, ")
                                if a == 0:
                                    c.execute('SELECT Id FROM Bancos WHERE NomeBanco = "' + _bancoorigem.get() + '"')
                                else:
                                    c.execute('SELECT Id FROM Bancos WHERE NomeBanco = "' + _bancodestino.get() + '"')
                                command.append(str(c.fetchone()[0]) + ", ")
                                if crd_card:
                                    command.append("'" + _fatura.get() + "', ")
                                    command.append("'<CRED.CARD>', ")
                                else:
                                    command.append("'', ")
                                    command.append("'" + _descricao.get() + "', ")
                                if a == 0:
                                    command.append('-' + valorfill + ', ')
                                    command.append("4, ")
                                else:
                                    command.append(valorfill + ', ')
                                    command.append("3, ")
                                c.execute('SELECT Id FROM Categorias WHERE Categoria = "TRANSFERENCIAS"')
                                command.append(str(c.fetchone()[0]))
                                command.append(")")
                                c.execute(str.join('', command))
                                NewId += 1
                            if crd_card:
                                if _fatura.get() and _datapago.get():
                                    c.execute('SELECT Id FROM Bancos WHERE NomeBanco = "' + _bancodestino.get() + '"')
                                    banco = c.fetchone()
                                    c.execute('UPDATE Diario SET DataPago = "' + _date(_datapago.get()) + '" WHERE Fatura = "' + _fatura.get() + '" AND Banco = ' + str(banco[0]) + '')
                                    messagebox.showinfo('Aviso!', 'Todos os pagamentos referente à fatura ' + _fatura.get() + ' foram baixados no sistema.', parent=form_edit)
                        if mode == 'edit':
                            for a in range(2):
                                command = []
                                command.append("UPDATE Diario SET ")
                                command.append("DataLastUpdate = '" + ActualDate + "', ")
                                command.append("DataDoc = '" + _date(_datadoc.get()) + "', ")
                                if a == 0:
                                    c.execute('SELECT Id FROM Bancos WHERE NomeBanco = "' + _bancoorigem.get() + '"')
                                else:
                                    c.execute('SELECT Id FROM Bancos WHERE NomeBanco = "' + _bancodestino.get() + '"')
                                command.append("Banco = " + str(c.fetchone()[0]) + ", ")
                                if crd_card:
                                    command.append("Fatura = '" + _fatura.get() + "', ")
                                    command.append("Descricao = '<CRED.CARD>', ")
                                else:
                                    command.append("Descricao = '" + _descricao.get() + "', ")
                                if a == 0:
                                    command.append("Valor = -" + valorfill + ", ")
                                else:
                                    command.append("Valor = " + valorfill + ", ")
                                command.append("DataVenc = '" + _date(_datavenc.get()) + "', ")
                                command.append("DataPago = '" + _date(_datapago.get()) + "' ")
                                command.append("WHERE Id = " +  str(idedit[a]))
                                c.execute(str.join('', command))
                            if crd_card:
                                if _fatura.get() and _datapago.get():
                                    c.execute('SELECT Id FROM Bancos WHERE NomeBanco = "' + _bancodestino.get() + '"')
                                    banco = c.fetchone()
                                    c.execute('UPDATE Diario SET DataPago = "' + _date(_datapago.get()) + '" WHERE Fatura = "' + _fatura.get() + '" AND Banco = ' + str(banco[0]) + '')
                                    messagebox.showinfo('Aviso!', 'Todos os pagamentos referente à fatura ' + _fatura.get() + ' foram baixados no sistema.', parent=form_edit)
                                if _fatura.get() and not _datapago.get():
                                    c.execute('SELECT Id FROM Bancos WHERE NomeBanco = "' + _bancodestino.get() + '"')
                                    banco = c.fetchone()
                                    c.execute('UPDATE Diario SET DataPago = "" WHERE Fatura = "' + _fatura.get() + '" AND Banco = ' + str(banco[0]) + '')
                                    messagebox.showinfo('Aviso!', 'Todos os pagamentos referente à fatura ' + _fatura.get() + ' foram estornados no sistema.', parent=form_edit)
                        if mode == 'remove':
                            c.execute('DELETE FROM Diario WHERE Id = ' + str(idedit[0]))
                            c.execute('DELETE FROM Diario WHERE Id = ' + str(idedit[1]))
                            if crd_card:
                                if _fatura.get() and _datapago.get():
                                    c.execute('SELECT Id FROM Bancos WHERE NomeBanco = "' + _bancodestino.get() + '"')
                                    banco = c.fetchone()
                                    c.execute('UPDATE Diario SET DataPago = "" WHERE Fatura = "' + _fatura.get() + '" AND Banco = ' + str(banco[0]) + '')
                                    messagebox.showinfo('Aviso!', 'Todos os pagamentos referente à fatura ' + _fatura.get() + ' foram estornados no sistema.', parent=form_edit)
                        conn.commit()
                        seek()
                        exit()
                    else:
                        update.focus()

        def busca_fatura(Value=None):
            if not crd_card:
                c.execute('SELECT GeraFatura, DiaVenc FROM Bancos WHERE NomeBanco = "' + _banco.get() + '"')
            else:
                c.execute('SELECT GeraFatura, DiaVenc FROM Bancos WHERE NomeBanco = "' + _bancodestino.get() + '"')
            resp = c.fetchone()
            try:
                if resp[0]:
                    hoje = datetime.strptime(_datadoc.get(), '%d/%m/%Y') + timedelta(days=10)
                    dia = hoje.day
                    mes = hoje.month
                    ano = hoje.year
                    if resp[1] >= dia:
                        pass
                    else:
                        mes = mes + 1
                        if mes == 13:
                            mes = 1
                            ano += 1
                    new_fat = str(ano) + '-' + str(mes).zfill(2)
                    new_venc = datetime(ano, mes, resp[1]).strftime('%d/%m/%Y')
                    if crd_card:
                        _descricao.configure(state='normal')
                        _descricao.delete(0, 'end')
                        _descricao.insert(0, crd_descricao(_bancodestino.get(), new_fat))
                        _descricao.configure(state='disabled')
                    _fatura.delete(0, 'end')
                    _fatura.insert(0, new_fat)
                    _datavenc.delete(0, 'end')
                    _datavenc.insert(0, new_venc)
            except:
                pass
                
        def busca_parceiro(Value=None):
            if _parceiro.get():
                widgets.combobox_return(_parceiro, lista1)

        def busca_banco(Value=None):
            if _banco.get():
                widgets.combobox_return(_banco, lista2)
                busca_fatura()

        def busca_bancoo(Value=None):
            if _bancoorigem.get():
                if not crd_card:
                    widgets.combobox_return(_bancoorigem, lista2)
                if crd_card:
                    widgets.combobox_return(_bancoorigem, lista_origem)

        def busca_bancod(Value=None):
            if _bancodestino.get():
                if not crd_card:
                    widgets.combobox_return(_bancodestino, lista2)
                if crd_card:
                    widgets.combobox_return(_bancodestino, lista_destino)
                    busca_fatura()

        def busca_categoria(Value=None):
            if _categoria.get():
                widgets.combobox_return(_categoria, lista3)

        def _datadoc_cmd(Value=None):
            data_cmd(_datadoc)

        def _datavenc_cmd(Value=None):
            data_cmd(_datavenc)

        def _datapago_cmd(Value=None):
            data_cmd(_datapago)

        def _fatura_cmd(Value=None):
            fatura_teste = _fatura.get().split('-')
            if len(fatura_teste) == 2:
                if not crd_card:
                    c.execute('SELECT GeraFatura, DiaVenc FROM Bancos WHERE NomeBanco = "' + _banco.get() + '"')
                else:
                    c.execute('SELECT GeraFatura, DiaVenc FROM Bancos WHERE NomeBanco = "' + _bancodestino.get() + '"')
                resp = c.fetchone()
                if resp:
                    if resp[0]:
                        if crd_card:
                            _descricao.configure(state='normal')
                            _descricao.delete(0, 'end')
                            _descricao.insert(0, crd_descricao(_bancodestino.get(), _fatura.get()))
                            _descricao.configure(state='disabled')
                        _datavenc.delete(0, 'end')
                        _datavenc.insert(0, datetime(int(fatura_teste[0]), int(fatura_teste[1]), resp[1]).strftime('%d/%m/%Y'))

        def _valor_cmd(Value=None):
            brvalue = _valor.get()
            brvalue = num_usa(brvalue)
            try:
                float(brvalue)
                _valor.delete(0, 'end')
                _valor.insert(0, num_brasil(brvalue))
            except:
                _valor.delete(0, 'end')
                _valor.insert(0, '0,00')

        dimension = widgets.geometry(350, 510)
        if in_value:
            title = 'Edição de pagamento'
        elif out_value:
            title = 'Edição de recebimento'
        else:
            if not crd_card:
                title = 'Transferência'
            else:
                title = 'Fatura de cartão de crédito'
        config = {'title': title,
                  'dimension': dimension,
                  'color': 'pale goldenrod'}
        form_edit = main.form(config)
        wgedit = Widgets(form_edit, 'pale goldenrod')
        wgedit.label('', 5, 0, 0)
        wgedit.label('', 5, 10, 0)
        wgedit.label('', 5, 13, 0)
        if in_value or out_value:
            if mode in ['edit', 'remove', 'dup']:
                if itemselect.selection():
                    for i in itemselect.selection():
                        gridselected = str(itemselect.item(i, 'text'))
                else:
                    messagebox.showerror(title='Aviso', message='Não foi selecionado nenhum documento!', parent=form_edit)
                    exit()
                command = ('SELECT Diario.Id, DataFirstUpdate, DataLastUpdate, DataDoc, Parceiros.Nome, ' +
                           'Bancos.NomeBanco, Categorias.Categoria, Fatura, Descricao, Valor, ' +
                           'DataVenc, DataPago FROM Diario ' +
                           'JOIN Parceiros ON Parceiros.Id = Parceiro ' +
                           'JOIN Bancos ON Bancos.Id = Banco ' +
                           'JOIN Categorias ON Categorias.Id = CategoriaMov ' +
                           'WHERE Diario.Id = ' + gridselected)
                c.execute(command)
                doc = c.fetchall()
            _id = wgedit.textbox('Id ', 5, 1, 1)
            _datafirstupdate = wgedit.textbox('Data inclusão ', 12, 2, 1)
            _datalastupdate = wgedit.textbox('Data última atualização ', 12, 3, 1)
            _datadoc = wgedit.textbox('Data do movimento ', 12, 4, 1, cmd=_datadoc_cmd)
            if in_value:
                c.execute('SELECT Nome FROM Parceiros WHERE Modo IN (0, 1) ORDER BY Nome')
            elif out_value:
                c.execute('SELECT Nome FROM Parceiros WHERE Modo IN (0, 2) ORDER BY Nome')
            lista1 = []
            for row in c.fetchall():
                lista1.append(row[0])
            _parceiro = wgedit.combobox(classe_parceiro[0] + ' ', 20, lista1, 5, 1, cmd=busca_parceiro, seek=busca_parceiro)
            c.execute('SELECT NomeBanco FROM Bancos ORDER BY NomeBanco')
            lista2 = []
            for row in c.fetchall():
                lista2.append(row[0])
            _banco = wgedit.combobox('Banco ', 20, lista2, 6, 1, cmd=busca_banco, seek=busca_banco)
            c.execute('SELECT Categoria FROM Categorias WHERE TipoMov = ' + cat + ' ORDER BY Categoria')
            lista3 = []
            for row in c.fetchall():
                lista3.append(row[0])
            _categoria = wgedit.combobox('Categoria ', 20, lista3, 7, 1, cmd=busca_categoria, seek=busca_categoria)
            _fatura = wgedit.textbox('Fatura ', 7, 8, 1, cmd=_fatura_cmd)
            _descricao = wgedit.textbox('Descrição ', 35, 9, 1)
            _valor = wgedit.textbox('Valor ', 8, 10, 1, cmd=_valor_cmd)
            _datavenc = wgedit.textbox('Vence em ', 12, 11, 1, cmd=_datavenc_cmd)
            _datapago = wgedit.textbox('Pago em ', 12, 12, 1, cmd=_datapago_cmd)

            if mode in ['edit', 'remove', 'dup']:
                if mode not in ['dup']:
                    _id.insert(0, str(doc[0][0]))
                    _datafirstupdate.insert(0, data_brasil(str(doc[0][1])))
                    _datalastupdate.insert(0, data_brasil(str(doc[0][2])))
                    _datadoc.insert(0, data_brasil(str(doc[0][3])))
                    _datavenc.insert(0, data_brasil(str(doc[0][10])))
                    _datapago.insert(0, data_brasil(str(doc[0][11])))
                _parceiro.insert(0, str(doc[0][4]))
                _banco.insert(0, str(doc[0][5]))
                _categoria.insert(0, str(doc[0][6]))
                _fatura.insert(0, str(doc[0][7]))
                _descricao.insert(0, str(doc[0][8]))
                _valor.insert(0, num_brasil(str(abs(doc[0][9]))))
            if mode in ['new', 'dup']:
                c.execute('SELECT Id FROM Diario ORDER BY Id')
                rows = c.fetchall()
                NewId = 1
                if rows:
                    Ids = []
                    for row in rows:
                        Ids.append(row[0])
                    Ids.sort() 
                    NewId = Ids[-1:][0] + 1
                _id.insert(0, str(NewId))
                _datafirstupdate.insert(0, data_brasil(ActualDate))
                _datalastupdate.insert(0, data_brasil(ActualDate))
                _datadoc.insert(0, data_brasil(ActualDate))
                _datavenc.insert(0, data_brasil(ActualDate))

            _id.configure(state='disabled')
            _datafirstupdate.configure(state='disabled')
            _datalastupdate.configure(state='disabled')
            if mode in ['edit', 'new', 'dup']:
                _datadoc.focus()
                update = wgedit.button('Atualizar', updatethis, 20, 0, 14, 1, 2)
                cancel = wgedit.button('Cancelar', exit, 20, 0, 15, 1, 2)
            if mode in ['remove']:
                _datadoc.configure(state='disabled')
                _parceiro.configure(state='disabled')
                _banco.configure(state='disabled')
                _categoria.configure(state='disabled')
                _fatura.configure(state='disabled')
                _descricao.configure(state='disabled')
                _valor.configure(state='disabled')
                _datavenc.configure(state='disabled')
                _datapago.configure(state='disabled')
                wgedit.label('Registro que será excluído!', 0, 14, 1, 2)
                update = wgedit.button('Confirmar exclusão', updatethis, 20, 0, 15, 1, 2)
        else: # linhas para edição de transferências.
            if mode in ['edit', 'remove', 'dup']:
                if itemselect.selection():
                    for i in itemselect.selection():
                        gridselected = str(itemselect.item(i, 'text'))
                else:
                    messagebox.showerror(title='Aviso', message='Não foi selecionado nenhum documento!')
                    exit()
                idedit = []
                c.execute('SELECT TipoMov FROM Diario WHERE Id = ' + gridselected)
                lin = c.fetchone()
                if lin[0] == 3:
                    idedit.append(int(gridselected) - 1)
                    idedit.append(int(gridselected))
                else:
                    idedit.append(int(gridselected))
                    idedit.append(int(gridselected) + 1)
                command = ('SELECT Diario.Id, DataFirstUpdate, DataLastUpdate, DataDoc, ' +
                        'Bancos.NomeBanco, Fatura, Descricao, Valor, DataVenc, DataPago FROM Diario ' +
                        'JOIN Parceiros ON Parceiros.Id = Parceiro ' +
                        'JOIN Bancos ON Bancos.Id = Banco ' +
                        'JOIN Categorias ON Categorias.Id = CategoriaMov ' +
                        'WHERE Diario.Id = ' + str(idedit[0]))
                c.execute(command)
                doc = c.fetchall()
                command = ('SELECT Bancos.NomeBanco FROM Diario ' +
                           'JOIN Bancos ON Bancos.Id = Banco WHERE Diario.Id = ' + str(idedit[1]))
                c.execute(command)
                doc2 = c.fetchall()
            _id = wgedit.textbox('Id ', 5, 1, 1)
            _datafirstupdate = wgedit.textbox('Data inclusão ', 12, 2, 1)
            _datalastupdate = wgedit.textbox('Data última atualização ', 12, 3, 1)
            _datadoc = wgedit.textbox('Data do movimento ', 12, 4, 1, cmd=_datadoc_cmd)
            if not crd_card:
                c.execute('SELECT NomeBanco FROM Bancos WHERE TipoMov != 2 ORDER BY NomeBanco')
                lista2 = []
                for row in c.fetchall():
                    lista2.append(row[0])
                _bancoorigem = wgedit.combobox('Banco Origem', 20, lista2, 5, 1, cmd=busca_bancoo, seek=busca_bancoo)
                _bancodestino = wgedit.combobox('Banco Destino', 20, lista2, 6, 1, cmd=busca_bancod, seek=busca_bancod)
            if crd_card:
                c.execute('SELECT NomeBanco FROM Bancos WHERE TipoMov != 2 ORDER BY NomeBanco')
                lista_origem = []
                for row in c.fetchall():
                    lista_origem.append(row[0])
                c.execute('SELECT NomeBanco FROM Bancos WHERE TipoMov = 2 ORDER BY NomeBanco')
                lista_destino = []
                for row in c.fetchall():
                    lista_destino.append(row[0])
                _bancoorigem = wgedit.combobox('Banco pagamento', 20, lista_origem, 5, 1, cmd=busca_bancoo, seek=busca_bancoo)
                _bancodestino = wgedit.combobox('Cartão', 20, lista_destino, 6, 1, cmd=busca_bancod, seek=busca_bancod)
                _fatura = wgedit.textbox('Fatura ', 8, 7, 1, default=datetime.now().strftime('%Y-%m'), cmd=_fatura_cmd)
            _descricao = wgedit.textbox('Descrição ', 35, 9, 1)
            _valor = wgedit.textbox('Valor ', 8, 10, 1, cmd=_valor_cmd)
            _datavenc = wgedit.textbox('Vence em ', 12, 11, 1, cmd=_datavenc_cmd)
            _datapago = wgedit.textbox('Pago em ', 12, 12, 1, cmd=_datapago_cmd)

            if mode in ['edit', 'remove', 'dup']:
                if mode not in ['dup']:
                    _id.insert(0, str(doc[0][0]))
                    _datafirstupdate.insert(0, data_brasil(str(doc[0][1])))
                    _datalastupdate.insert(0, data_brasil(str(doc[0][2])))
                    _datadoc.insert(0, data_brasil(str(doc[0][3])))
                    _datavenc.insert(0, data_brasil(str(doc[0][8])))
                    _datapago.insert(0, data_brasil(str(doc[0][9])))
                _bancoorigem.insert(0, str(doc[0][4]))
                _bancodestino.insert(0, str(doc2[0][0]))
                if crd_card:
                    _fatura.delete(0, 'end')
                    _fatura.insert(0, str(doc[0][5]))
                if not crd_card:
                    _descricao.insert(0, str(doc[0][6]))
                else:
                    _descricao.insert(0, crd_descricao(str(doc2[0][0]), str(doc[0][5])))
                _valor.insert(0, num_brasil(str(abs(doc[0][7]))))
            if mode in ['new', 'dup']:
                c.execute('SELECT Id FROM Diario ORDER BY Id')
                rows = c.fetchall()
                NewId = 1
                if rows:
                    Ids = []
                    for row in rows:
                        Ids.append(row[0])
                    Ids.sort() 
                    NewId = Ids[-1:][0] + 1
                _id.insert(0, str(NewId))
                _datafirstupdate.insert(0, data_brasil(ActualDate))
                _datalastupdate.insert(0, data_brasil(ActualDate))
                _datadoc.insert(0, data_brasil(ActualDate))
                _datavenc.insert(0, data_brasil(ActualDate))

            _id.configure(state='disabled')
            _datafirstupdate.configure(state='disabled')
            _datalastupdate.configure(state='disabled')
            if mode in ['edit', 'new', 'dup']:
                if crd_card:
                    _descricao.configure(state='disabled')
                _datadoc.focus()
                update = wgedit.button('Atualizar', updatethis, 20, 0, 14, 1, 2)
                cancel = wgedit.button('Cancelar', exit, 20, 0, 15, 1, 2)
            if mode in ['remove']:
                _datadoc.configure(state='disabled')
                _bancoorigem.configure(state='disabled')
                _bancodestino.configure(state='disabled')
                _fatura.configure(state='disabled')
                _descricao.configure(state='disabled')
                _valor.configure(state='disabled')
                _datavenc.configure(state='disabled')
                _datapago.configure(state='disabled')
                wgedit.label('Registro que será excluído!', 0, 14, 1, 2)
                update = wgedit.button('Confirmar exclusão', updatethis, 20, 0, 15, 1, 2)

    def generate_sql(value=False):
        def _fc(first):
            if first:
                resp = 'WHERE '
            else:
                resp = ' AND '
            return resp

        first = True
        command = []
        if value:
            command.append('SELECT SUM(Valor) ')
        else:
            if not in_value and not out_value:
                command.append('SELECT Diario.Id, DataDoc, Diario.TipoMov, Bancos.NomeBanco, Descricao, DataVenc, Valor, Fatura ')
            else:
                command.append('SELECT Diario.Id, DataDoc, Parceiros.Nome, Descricao, Categorias.Categoria, DataVenc, Valor ')
        command.append('FROM Diario JOIN Parceiros ON Parceiros.Id = Parceiro JOIN Categorias ON Categorias.Id = CategoriaMov ')
        command.append('JOIN Bancos ON Bancos.Id = Banco ')
        if in_value or out_value:
            if parceiroini.get():
                command.append(_fc(first) + 'Parceiros.Nome = "' + parceiroini.get() + '"')
                first = False
            if categini.get():
                command.append(_fc(first) + 'Categorias.Categoria = "' + categini.get() + '"')
                first = False
        if bancoini.get():
            command.append(_fc(first) + 'Bancos.NomeBanco = "' + bancoini.get() + '"')
            first = False
        if faturaini.get():
            command.append(_fc(first) + 'Fatura = "' + faturaini.get() + '"')
            first = False
        if datadocini.get():
            datafill = _date(datadocini.get())
            command.append(_fc(first) + 'DataDoc >= "' + datafill + '"')
            first = False
        if datadocfim.get():
            datafill = _date(datadocfim.get())
            command.append(_fc(first) + 'DataDoc <= "' + datafill + '"')
            first = False
        if datavencini.get():
            datafill = _date(datavencini.get())
            command.append(_fc(first) + 'DataVenc >= "' + datafill + '"')
            first = False
        if datavencfim.get():
            datafill = _date(datavencfim.get())
            command.append(_fc(first) + 'DataVenc <= "' + datafill + '"')
            first = False
        if whereval:
            command.append(_fc(first) + whereval)
            first = False
        command.append(' ORDER BY DataDoc')
        if value:
            c.execute(str.join('', command))
            resp = c.fetchall()[0][0]
            if not resp:
                resp = 0.0
            return resp
        else:
            return str.join('', command)

    def seek():
        itemselect.delete(*itemselect.get_children())
        soma['text'] = 'Valor Total: 0,00'
        c.execute(generate_sql())
        doc = c.fetchall()
        combolist = {}
        ordlista = []
        somalc = 0.0
        for row in doc:
            texts = []
            rowact = 0
            for rows in row[1:]:
                if rowact in [1] and not in_value and not out_value:
                    texts.append(trftype[rows])
                elif rowact in [0, 4]:
                    texts.append(data_brasil(rows))
                elif rowact in [3] and crd_card:
                    texts.append(crd_descricao(row[3], row[7]))
                elif rowact in [5]:
                    somalc += rows
                    if texts[1] == 'Saída' and not in_value and not out_value:
                        texts.append('-' + num_brasil(str(abs(rows))))
                    else:
                        texts.append(num_brasil(str(abs(rows))))
                else:
                    texts.append(rows)
                rowact += 1
            ordlista.append(row[0])
            combolist[row[0]] = tuple(texts)
        for rows in ordlista:
            itemselect.insert('', 'end', text=rows, values=combolist[rows])
        if not in_value and not out_value:
            soma['text'] = 'Valor Total: ' + num_brasil(format(abs(somalc), '.2f'))
        else:
            soma['text'] = 'Valor Total: ' + num_brasil(format(abs(generate_sql(value=True)), '.2f'))

    def filecsv():
        file_output = "csv/detalhe_movimentos.csv"
        # soma['text'] = 'Valor Total: 0,00'
        c.execute(generate_sql())
        doc = c.fetchall()
        combolist = {}
        ordlista = []
        somalc = 0.0
        for row in doc:
            texts = {}
            rowact = 0
            for rows in row[1:]:
                if rowact in [1] and not in_value and not out_value:
                    texts[colsf[rowact]] = trftype[rows]
                elif rowact in [0, 4]:
                    texts[colsf[rowact]] = data_brasil(rows)
                elif rowact in [5]:
                    somalc += rows
                    if texts[colsf[1]] == 'Saída' and not in_value and not out_value:
                        texts[colsf[rowact]] = '-' + num_brasil(str(abs(rows)))
                    else:
                        texts[colsf[rowact]] = num_brasil(str(abs(rows)))
                else:
                    texts[colsf[rowact]] = rows
                rowact += 1
            ordlista.append(row[0])
            combolist[row[0]] = texts
        with open(file_output, 'w') as csvfile:
            fieldnames = colsf
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames, delimiter=';')
            writer.writeheader()
            for row in combolist:
                #rowact = 0
                writer.writerow(combolist[row])

        '''
        for rows in ordlista:
            itemselect.insert('', 'end', text=rows, values=combolist[rows])
        if not in_value and not out_value:
            soma['text'] = 'Valor Total: ' + num_brasil(format(abs(somalc), '.2f'))
        else:
            soma['text'] = 'Valor Total: ' + num_brasil(format(abs(generate_sql(value=True)), '.2f'))
        '''

    def seekparceiro(Value=None):
        if parceiroini.get():
            widgets.combobox_return(parceiroini, lista_p)    

    def seekcategoria(Value=None):
        if categini.get():
            widgets.combobox_return(categini, lista_c)    

    def seekbanco(Value=None):
        if bancoini.get():
            widgets.combobox_return(bancoini, lista_b)

    def datadocini_cmd(Value=None):
        data_cmd(datadocini)

    def datadocfim_cmd(Value=None):
        data_cmd(datadocfim)

    def datavencini_cmd(Value=None):
        data_cmd(datavencini)

    def datavencfim_cmd(Value=None):
        data_cmd(datavencfim)

    def crd_descricao(banco='', fatura=''):
        if banco:
            c.execute('SELECT GeraFatura, DiaVenc FROM Bancos WHERE NomeBanco = "' + banco + '"')
            resp = c.fetchone()
        periodo_fatura = fatura.split('-')
        descricao_fatura = ''
        if resp:
            if resp[0] and len(periodo_fatura) == 2:
                if resp[1] > 10:
                    data_fatfim = datetime(int(periodo_fatura[0]), int(periodo_fatura[1]), resp[1] - 10)
                    if int(periodo_fatura[1]) == 1:
                        month_fat = 12
                        year_fat = int(periodo_fatura[0]) - 1
                    else:
                        month_fat = int(periodo_fatura[1]) - 1
                        year_fat = int(periodo_fatura[0])
                    data_fatini = datetime(year_fat, month_fat, resp[1] - 10) + timedelta(days=1)
                else:
                    days_month = (datetime(int(periodo_fatura[0]), int(periodo_fatura[1]), 1) - timedelta(days=1)).day
                    day_fat = resp[1] + 20
                    if day_fat > days_month:
                        day_fat = days_month
                    if int(periodo_fatura[1]) == 1:
                        month_fat = 12
                        year_fat = int(periodo_fatura[0]) - 1
                    else:
                        month_fat = int(periodo_fatura[1]) - 1
                        year_fat = int(periodo_fatura[0])
                    data_fatfim = datetime(year_fat, month_fat, day_fat)

                    days_month = (datetime(year_fat, month_fat, 1) - timedelta(days=1)).day
                    day_fat = resp[1] + 20
                    if day_fat > days_month:
                        day_fat = days_month
                    if month_fat == 1:
                        month_fat = 12
                        year_fat -= 1
                    else:
                        month_fat = month_fat - 1
                    data_fatini = datetime(year_fat, month_fat, day_fat) + timedelta(days=1)
                descricao_fatura += 'Movimentação ('
                descricao_fatura += data_fatini.strftime('%d/%m/%Y') + ' e '
                descricao_fatura += data_fatfim.strftime('%d/%m/%Y') + ')'
        return descricao_fatura

    dimension = widgets.geometry(502, 950)
    if in_value:
        title = 'CONTAS À RECEBER - Selecione um evento e a ação desejada'
        color = 'spring green'
    elif out_value:
        title = 'CONTAS À PAGAR - Selecione um evento e a ação desejada'
        color = 'light salmon'
    else:
        if not crd_card:
            title = 'TRANSFERÊNCIAS - Selecione um evento e a ação desejada'
            color = 'light goldenrod'
        else:
            title = 'FATURAS DE CARTÕES DE CRÉDITO - Selecione um evento e a ação desejada'
            color = 'light salmon'
    config = {'title': title,
              'dimension': dimension,
              'color': color}
    form_selec = main.form(config)
    if in_value:
        whereval = 'Categorias.TipoMov = 0'
    elif out_value:
        whereval = 'Categorias.TipoMov = 1'
    else:
        if not crd_card:
            whereval = 'Categorias.TipoMov = 2 AND Diario.Descricao != "<CRED.CARD>"'
        else:
            whereval = 'Categorias.TipoMov = 2 AND Diario.Descricao = "<CRED.CARD>" AND Valor > 0.0'
    wg = Widgets(form_selec, color)
    wg.label('', 3, 0, 0)
    wg.label('', 3, 3, 0)
    wg.label('', 3, 5, 0)
    wg.label('', 3, 7, 0)
    if in_value or out_value:
        if in_value:
            c.execute('SELECT Nome FROM Parceiros WHERE Modo IN (0, 1) ORDER BY Nome')
            classe_parceiro = ['Cliente']
        elif out_value:
            c.execute('SELECT Nome FROM Parceiros WHERE Modo IN (0, 2) ORDER BY Nome')
            classe_parceiro = ['Fornecedor']
        lista_p = []
        for row in c.fetchall():
            lista_p.append(row[0])
        parceiroini = wg.combobox('  ' + classe_parceiro[0] + ': ', 16, lista_p, 1, 1, cmd=seekparceiro, seek=seekparceiro)
        if in_value:
            cat = '0'
        elif out_value:
            cat = '1'
        else:
            cat = '2'
        c.execute('SELECT Categoria FROM Categorias WHERE TipoMov = ' + cat + ' ORDER BY Categoria')
        lista_c = []
        for row in c.fetchall():
            lista_c.append(row[0])
        categini = wg.combobox('Categoria: ', 18, lista_c, 2, 1, cmd=seekcategoria, seek=seekcategoria)
    if in_value or out_value:
        c.execute('SELECT NomeBanco FROM Bancos ORDER BY NomeBanco')
    else:
        if not crd_card:
            c.execute('SELECT NomeBanco FROM Bancos WHERE TipoMov != 2 ORDER BY NomeBanco')
        else:
            c.execute('SELECT NomeBanco FROM Bancos WHERE TipoMov = 2 ORDER BY NomeBanco')
    lista_b = []
    for row in c.fetchall():
        lista_b.append(row[0])
    if in_value or out_value:
        colss = [3, 5]
    else:
        colss = [1, 3]
    if not crd_card:
        bancoini = wg.combobox('  Banco: ', 16, lista_b, 1, colss[0], seek=seekbanco)
    else:
        bancoini = wg.combobox('  Cartão: ', 16, lista_b, 1, colss[0], seek=seekbanco)
    faturaini = wg.textbox('  Fatura: ', 8, 2, colss[0])
    datadocini = wg.textbox('  Data Inicial: ', 10, 1, colss[1], default=data_brasil(ActualDate[0:7] + '-01'), cmd=datadocini_cmd)
    datadocfim = wg.textbox('  Data Final: ', 10, 2, colss[1], default=data_brasil(str(lastdaymonth(ActualDate))), cmd=datadocfim_cmd)
    datavencini = wg.textbox('  Venc Inicial: ', 10, 1, 7, cmd=datavencini_cmd)
    datavencfim = wg.textbox('  Venc Final: ', 10, 2, 7, cmd=datavencfim_cmd)
    Seek = wg.button('Procurar', seek, 10, 0, 4, 1, 4)
    File = wg.button('Gerar csv', filecsv, 10, 0, 4, 3, 4)
    soma = wg.label('', 30, 4, 5, 4)
    if in_value or out_value:
        soma['text'] = 'Valor Total: ' + num_brasil(format(abs(generate_sql(value=True)), '.2f'))
    soma['fg'] = 'red'
    soma['font'] = 'Arial 12 bold'
    # grid setup ini
    if not in_value and not out_value:
        colsf = ['data', 'tipo', 'banco', 'descr', 'venc', 'valor']
        if not crd_card:
            bank_desc = 'Banco'
        else:
            bank_desc = 'Cartão'
        headf = {
            'data': {'text': 'Data', 'width': 80, 'format': 'date'},
            'tipo': {'text': 'E/S', 'width': 80},
            'banco': {'text': bank_desc, 'width': 210},
            'descr': {'text': 'Descrição', 'width': 300},
            'venc': {'text': 'Vencimento', 'width': 90, 'format': 'date'},
            'valor': {'text': 'Valor', 'width': 70, 'anchor': 'e', 'format': 'float'},
        }
    else:
        if in_value:
            _categ = 'Tipo de Receita'
        elif out_value:
            _categ = 'Tipo de Despesa'
        colsf = ['data', 'parceiro', 'descr', 'categ', 'venc', 'valor']
        headf = {
            'data': {'text': 'Data', 'width': 80, 'format': 'date'},
            'parceiro': {'text': classe_parceiro[0], 'width': 150},
            'descr': {'text': 'Descrição', 'width': 250},
            'categ': {'text': _categ, 'width': 190},
            'venc': {'text': 'Vencimento', 'width': 90, 'format': 'date'},
            'valor': {'text': 'Valor', 'width': 70, 'anchor': 'e', 'format': 'float'},
        }
    c.execute(generate_sql())
    doc = c.fetchall()
    lista = {}
    listaord = []
    trftype = {3: 'Entrada', 4: 'Saída'}
    for row in doc:
        texts = []
        rowact = 0
        for rows in row[1:]:
            if rowact in [1] and not in_value and not out_value:
                texts.append(trftype[rows])
            elif rowact in [0, 4]:
                texts.append(data_brasil(rows))
            elif rowact in [3] and crd_card:
                texts.append(crd_descricao(row[3], row[7]))
            elif rowact in [5]:
                if texts[1] == 'Saída' and not in_value and not out_value:
                    val = rows
                    texts.append('-' + num_brasil(str(abs(rows))))
                else:
                    texts.append(num_brasil(str(abs(rows))))
            else:
                texts.append(rows)
            rowact += 1
        listaord.append(row[0])
        lista[row[0]] = tuple(texts)
    itemselect = wg.grid(colsf, headf, lista, listaord, 14, 6, 1, colspan=8)
    # grid setup fim
    Seek.focus()
    Create = wg.button('Incluir', newreg, 10, 0, 8, 1, 1)
    Duplic = wg.button('Duplicar', dupreg, 10, 0, 8, 2, 1)
    Edit = wg.button('Editar', editreg, 10, 0, 8, 3, 1)
    if in_value:
        texto = 'Confirmar Recebimento'
    else:
        texto = 'Confirmar Pagamento'
    Pag = wg.button(texto, pagreg, 17, 0, 8, 4, 2)
    Remove = wg.button('Excluir', removereg, 10, 0, 8, 6, 2)
    Sair = wg.button('Sair', exitsc, 10, 0, 8, 8, 1)

def movimentos_in():
    movimentos(in_value=True)

def movimentos_out():
    movimentos(out_value=True)

def movimentos_crd():
    movimentos(crd_card=True)

def materiais_categorias():
    def newreg():
        edit('new')

    def editreg():
        edit('edit')

    def removereg():
        edit('remove')

    def edit(mode):
        def exit():
            form_edit.destroy()
            categselect.focus()
        
        def updatethis():
            executethis = True
            
            if mode in ['new', 'edit']:
                mens = 'Tem certeza que deseja atualizar ' + _categoria.get() + '?'
                erros = []
                # Validar se tem repetido.
                c.execute('SELECT Id FROM Materiais_Categorias WHERE Categoria = "' + _categoria.get() + '" AND Id <> ' + _id.get())
                doc = c.fetchall()
                if doc:
                    erros.append('-Já existe outra categoria com o mesmo nome.')
                # Validar não vazios.
                values = {'Categoria': '', 'Tipo de movimento': ''}
                values['Categoria'] = _categoria.get()
                values['Tipo de movimento'] = _tipomov.get()
                empty = []
                for row in sorted(values.keys()):
                    if values[row] == '':
                        empty.append(row)
                if empty:
                    if len(empty) == 1:
                        msgerro = '-O campo ' + empty[0] + ' não foi preenchido.'
                        erros.append(msgerro)
                    else:
                        msgerro = '-Os campos ' + empty[0]
                        for row in empty[1:-1]:
                            msgerro = msgerro + ', ' + row
                        msgerro = msgerro + ' e ' + empty[-1:][0] + ' não foram preenchidos.'
                        erros.append(msgerro)
                if erros:
                    msg = 'Não é possível atualizar esse cadastro:\n\n' + str.join('\n', erros) + '\n\nRevise os dados e tente novamente.'
                    messagebox.showerror(title='Aviso', message=msg, parent=form_edit)
                    update.focus()
                    executethis = False
            else:
                mens = 'Tem certeza que deseja remover ' + _categoria.get() + '?'
                '''
                c.execute('SELECT Diario.Id FROM Diario JOIN Bancos ON Bancos.Id = Banco WHERE Bancos.NomeBanco = "' + _nomebanco.get() + '"')
                doc = c.fetchall()
                if doc:
                    messagebox.showerror(title='Aviso', message='Você não pode apagar um banco com movimento registrado!', parent=form_edit)
                    executethis = False
                '''
            if executethis:
                if messagebox.askyesno(title="Atualização de sistema", message=mens, parent=form_edit):
                    command = []
                    if mode == 'new':
                        command.append("INSERT INTO Materiais_Categorias VALUES (")
                        command.append(_id.get() + ", ")
                        command.append("'" + str(datetime.now()) + "', ")
                        command.append("'" + str(datetime.now()) + "', ")
                        command.append("'" + _categoria.get() + "', ")
                        command.append(str(lista1.index(_tipomov.get())))
                        command.append(")")
                    if mode == 'edit':
                        command.append("UPDATE Materiais_Categorias SET ")
                        command.append("DataAtualiza = '" + str(datetime.now()) + "', ")
                        command.append("Categoria = '" + _categoria.get() + "', ")
                        command.append("Tipo = " + str(lista1.index(_tipomov.get())) + " ")
                        command.append("WHERE Id = " + _id.get())
                    if mode == 'remove':
                        command.append('DELETE FROM Materiais_Categorias WHERE Id = ' + _id.get())
                    c.execute(str.join('', command))
                    conn.commit()
                    categselect.delete(0, 'end')
                    c.execute('SELECT Categoria FROM Materiais_Categorias ORDER BY Categoria')
                    '''
                    if tipo == 'mov':
                        c.execute('SELECT NomeBanco FROM Bancos WHERE TipoMov <= 1 ORDER BY NomeBanco')
                    elif tipo == 'cc':
                        c.execute('SELECT NomeBanco FROM Bancos WHERE TipoMov = 2 ORDER BY NomeBanco')
                    elif tipo == 'cpp':
                        c.execute('SELECT NomeBanco FROM Bancos WHERE TipoMov = 3 ORDER BY NomeBanco')
                    '''
                    doc = c.fetchall()
                    lista = []
                    for row in doc:
                        lista.append(row[0])
                    for row in lista:
                        categselect.insert('end', row)
                    categselect.focus()
                    exit()
                else:
                    update.focus()

        def _tipomov_cmd(value=None):
            widgets.combobox_return(_tipomov, lista1)

        dimension = widgets.geometry(308, 500)
        config = {'title': 'Edição de Categorias de Materiais',
                  'dimension': dimension,
                  'color': 'pale goldenrod'}
        form_edit = main.form(config)
        wgedit = Widgets(form_edit, 'pale goldenrod')
        wgedit.label('', 5, 0, 0)
        wgedit.label('', 5, 11, 0)
        wgedit.label('', 5, 14, 0)
        if mode in ['edit', 'remove']:
            c.execute('SELECT * FROM Materiais_Categorias WHERE Categoria = "' + categselect.get('active') + '"')
            doc = c.fetchall()
        _id = wgedit.textbox('Id ', 2, 1, 1)
        _datacadastro = wgedit.textbox('Data da criação ', 17, 2, 1)
        _dataatualiza = wgedit.textbox('Data última edição ', 17, 3, 1)
        _categoria = wgedit.textbox('Categoria ', 30, 4, 1)
        lista1 = ['Compra', 'Venda', 'Revenda']
        _tipomov = wgedit.combobox('Tipo de movimento ', 15, lista1, 5, 1, cmd=_tipomov_cmd, seek=_tipomov_cmd)

        if mode in ['edit', 'remove']:
            _id.insert(0, str(doc[0][0]))
            _datacadastro.insert(0, datahora_brasil(doc[0][1]))
            _dataatualiza.insert(0, datahora_brasil(doc[0][2]))
            _categoria.insert(0, str(doc[0][3]))
            _tipomov.insert(0, lista1[doc[0][4]])
        else:
            c.execute('SELECT Id FROM Materiais_Categorias ORDER BY Id')
            rows = c.fetchall()
            NewId = 1
            if rows:
                Ids = []
                for row in rows:
                    Ids.append(row[0])
                Ids.sort() 
                NewId = Ids[-1:][0] + 1
            _id.insert(0, str(NewId))
            _datacadastro.insert(0, datahora_brasil(str(datetime.now())))
            _dataatualiza.insert(0, datahora_brasil(str(datetime.now())))

        _id.configure(state='disabled')
        _datacadastro.configure(state='disabled')
        _dataatualiza.configure(state='disabled')
        if mode in ['edit', 'new']:
            _categoria.focus()
            update = wgedit.button('Atualizar', updatethis, 20, 0, 12, 1, 2)
            cancel = wgedit.button('Cancelar', exit, 20, 0, 13, 1, 2)
        else:
            _categoria.configure(state='disabled')
            _tipomov.configure(state='disabled')
            wgedit.label('Registro que será excluído!', 0, 12, 1, 2)
            update = wgedit.button('Confirmar exclusão', updatethis, 20, 0, 13, 1, 2)
    
    dimension = widgets.geometry(330, 580)
    config = {'title': 'Selecione uma Categoria e a ação',
              'dimension': dimension,
              'color': 'light goldenrod'}
    form_selec = main.form(config)
    wg = Widgets(form_selec, 'light goldenrod')
    wg.label('', 5, 0, 0)
    wg.label('', 5, 2, 0)
    wg.label('', 5, 6, 0)
    wg.image('',0,'images/pag-categorias.png', 1, 1, 1, 6, imagewidth=(256, 256))
    c.execute('SELECT Categoria FROM Materiais_Categorias ORDER BY Categoria')
    '''
    if tipo == 'mov':
        wg.image('',0,'images/bancos.png', 1, 1, 1, 6)
        c.execute('SELECT NomeBanco FROM Bancos WHERE TipoMov <= 1 ORDER BY NomeBanco')
    elif tipo == 'cc':
        wg.image('',0,'images/creditcard.png', 1, 1, 1, 6)
        c.execute('SELECT NomeBanco FROM Bancos WHERE TipoMov = 2 ORDER BY NomeBanco')
    elif tipo == 'cpp':
        wg.image('',0,'images/creditcard.png', 1, 1, 1, 6)
        c.execute('SELECT NomeBanco FROM Bancos WHERE TipoMov = 3 ORDER BY NomeBanco')
    '''
    doc = c.fetchall()
    lista = []
    for row in doc:
        lista.append(row[0])
    categselect = wg.listbox('    Categoria:  ', 20, 8, lista, 1, 3)
    categselect.focus()
    Create = wg.button('Incluir', newreg, 10, 0, 3, 3, 2)
    Edit = wg.button('Editar', editreg, 10, 0, 4, 3, 2)
    Remove = wg.button('Excluir', removereg, 10, 0, 5, 3, 2)

def materiais_itens():
    def newreg():
        edit('new')

    def editreg():
        edit('edit')

    def removereg():
        edit('remove')

    def edit(mode):
        def exit():
            form_edit.destroy()
            Create.focus()
        
        def updatethis():
            executethis = True
            
            if mode in ['new', 'edit']:
                mens = 'Tem certeza que deseja atualizar ' + _produto.get() + '?'
                erros = []
                # Validar se tem repetido.
                c.execute('SELECT Id FROM Materiais_Itens WHERE Produto = "' + _produto.get() + '" AND Id <> ' + _id.get())
                doc = c.fetchall()
                if doc:
                    erros.append('-Já existe outro produto com o mesmo nome.')
                # Validar não vazios.
                values = {'Produto': '', 'Categoria': '', 'Tipo de movimento': '', 'Unidade de medida': '', 'Marca': '',
                    'Custo unitário': '', 'Margem de lucro': '', 'Preço unitário': '', 'Estoque mínimo': ''
                }
                values['Produto'] = _produto.get()
                values['Categoria'] = _categoria.get()
                values['Tipo de movimento'] = _tipo.get()
                values['Unidade de medida'] = _unidade.get()
                values['Marca'] = _marca.get()
                values['Custo unitário'] = _valorcusto.get()
                values['Margem de lucro'] = _margem.get()
                values['Preço unitário'] = _valorvenda.get()
                values['Estoque mínimo'] = _estoquemin.get()
                empty = []
                for row in sorted(values.keys()):
                    if values[row] == '':
                        empty.append(row)
                if empty:
                    if len(empty) == 1:
                        msgerro = '-O campo ' + empty[0] + ' não foi preenchido.'
                        erros.append(msgerro)
                    else:
                        msgerro = '-Os campos ' + empty[0]
                        for row in empty[1:-1]:
                            msgerro = msgerro + ', ' + row
                        msgerro = msgerro + ' e ' + empty[-1:][0] + ' não foram preenchidos.'
                        erros.append(msgerro)
                if erros:
                    msg = 'Não é possível atualizar esse cadastro:\n\n' + str.join('\n', erros) + '\n\nRevise os dados e tente novamente.'
                    messagebox.showerror(title='Aviso', message=msg, parent=form_edit)
                    update.focus()
                    executethis = False
            else:
                mens = 'Tem certeza que deseja remover ' + _categoria.get() + '?'
                c.execute('SELECT Materiais_Movimentos.Id FROM Materiais_Movimentos JOIN Materiais_Itens ON Materiais_Itens.Id = Materiais_Movimentos.Produto WHERE Materiais_Itens.Produto = "' + _produto.get() + '"')
                doc = c.fetchall()
                if doc:
                    messagebox.showerror(title='Aviso', message='Você não pode apagar um produto que possui movimentação!', parent=form_edit)
                    update.focus()
                    executethis = False
            if executethis:
                if messagebox.askyesno(title="Atualização de sistema", message=mens, parent=form_edit):
                    command = []
                    if mode == 'new':
                        command.append("INSERT INTO Materiais_Itens VALUES (")
                        command.append(_id.get() + ", ")
                        command.append("'" + str(datetime.now()) + "', ")
                        command.append("'" + str(datetime.now()) + "', ")
                        command.append("'" + _produto.get() + "', ")
                        c.execute('SELECT Id FROM Materiais_Categorias WHERE Categoria = "' + _categoria.get() + '"')
                        command.append(str(c.fetchone()[0]) + ", ")
                        command.append(str(lista2.index(_tipo.get())) + ", ")
                        command.append("'" + _unidade.get() + "', ")
                        command.append("'" + _marca.get() + "', ")
                        command.append(num_usa(_valorcusto.get()) + ", ")
                        command.append(num_usa(_margem.get()) + ", ")
                        command.append(num_usa(_valorvenda.get()) + ", ")
                        command.append(num_usa(_estoquemin.get()))
                        command.append(")")
                    if mode == 'edit':
                        command.append("UPDATE Materiais_Itens SET ")
                        command.append("DataAtualiza = '" + str(datetime.now()) + "', ")
                        command.append("Produto = '" + _produto.get() + "', ")
                        c.execute('SELECT Id FROM Materiais_Categorias WHERE Categoria = "' + _categoria.get() + '"')
                        command.append("Categoria = " + str(c.fetchone()[0]) + ", ")
                        command.append("Tipo = " + str(lista2.index(_tipo.get())) + ", ")
                        command.append("Unidade = '" + _unidade.get() + "', ")
                        command.append("Marca = '" + _marca.get() + "', ")
                        command.append("ValorCusto = " + num_usa(_valorcusto.get()) + ", ")
                        command.append("Margem = " + num_usa(_margem.get()) + ", ")
                        command.append("ValorVenda = " + num_usa(_valorvenda.get()) + ", ")
                        command.append("EstoqueMin = " + num_usa(_estoquemin.get()) + " ")
                        command.append("WHERE Id = " + _id.get())
                    if mode == 'remove':
                        command.append('DELETE FROM Materiais_Itens WHERE Id = ' + _id.get())
                    c.execute(str.join('', command))
                    conn.commit()
                    prodselect.delete(0, 'end')
                    c.execute('SELECT Produto FROM Materiais_Itens ORDER BY Produto')
                    '''
                    if tipo == 'mov':
                        c.execute('SELECT NomeBanco FROM Bancos WHERE TipoMov <= 1 ORDER BY NomeBanco')
                    elif tipo == 'cc':
                        c.execute('SELECT NomeBanco FROM Bancos WHERE TipoMov = 2 ORDER BY NomeBanco')
                    elif tipo == 'cpp':
                        c.execute('SELECT NomeBanco FROM Bancos WHERE TipoMov = 3 ORDER BY NomeBanco')
                    '''
                    doc = c.fetchall()
                    lista = []
                    for row in doc:
                        lista.append(row[0])
                    for row in lista:
                        prodselect.insert('end', row)
                    prodselect.focus()
                    exit()
                else:
                    update.focus()

        def _categoria_cmd(value=None):
            widgets.combobox_return(_categoria, lista1)
            if _categoria.get():
                c.execute('SELECT Tipo FROM Materiais_Categorias WHERE Categoria = "' + _categoria.get() + '"')
                _tipo.delete(0, 'end')
                _tipo.insert(0, lista2[c.fetchone()[0]])

        def _tipo_cmd(value=None):
            widgets.combobox_return(_tipo, lista2)

        def _valorcusto_cmd(value=None):
            sval = num_usa(_valorcusto.get())
            try:
                sval = round(float(sval), 2)
                if float(num_usa(_valorvenda.get())) < sval:
                    _valorcusto.delete(0, 'end')
                    _valorcusto.insert(0, num_brasil(str(sval)))
                    _margem.delete(0, 'end')
                    _margem.insert(0, '0,00')
                    _valorvenda.delete(0, 'end')
                    _valorvenda.insert(0, num_brasil(str(sval)))
                else:
                    _valorcusto.delete(0, 'end')
                    _valorcusto.insert(0, num_brasil(str(sval)))
                    _margem.delete(0, 'end')
                    _margem.insert(0, num_brasil(str(round(float(num_usa(_valorvenda.get())) - sval, 2))))
            except:
                messagebox.showwarning('Valor incorreto', 'Preencha um valor em formato numérico', parent=form_edit)
                _valorcusto.delete(0, 'end')
                _valorcusto.insert(0, '0,00')
                _valorvenda.delete(0, 'end')
                _valorvenda.insert(0, '0,00')
                update.focus()

        def _margem_cmd(value=None):
            sval = num_usa(_margem.get())
            try:
                sval = round(float(sval), 2)
                if sval < 0.0:
                    messagebox.showwarning('Valor incorreto', 'Você não pode preencher uma margem de lucro negativa.', parent=form_edit)
                    update.focus()
                    sval = 0.0
                _margem.delete(0, 'end')
                _margem.insert(0, num_brasil(str(sval)))
                _valorvenda.delete(0, 'end')
                _valorvenda.insert(0, num_brasil(str(round(float(num_usa(_valorcusto.get())) + sval, 2))))
            except:
                messagebox.showwarning('Valor incorreto', 'Preencha um valor em formato numérico', parent=form_edit)
                _margem.delete(0, 'end')
                _margem.insert(0, '0,00')
                _valorvenda.delete(0, 'end')
                _valorvenda.insert(0, _valorcusto.get())
                update.focus

        def _valorvenda_cmd(value=None):
            sval = num_usa(_valorvenda.get())
            try:
                sval = round(float(sval), 2)
                if sval < 0.0:
                    messagebox.showwarning('Valor incorreto', 'Você não pode preencher um preço negativa.', parent=form_edit)
                    sval = 0.0
                    update.focus()
                if sval < float(num_usa(_valorcusto.get())):
                    messagebox.showwarning('Valor incorreto', 'Você não pode preencher um preço de venda inferior ao custo do material.', parent=form_edit)
                    sval = float(num_usa(_valorcusto.get()))
                    update.focus()
                _margem.delete(0, 'end')
                _margem.insert(0, num_brasil(str(round(sval - float(num_usa(_valorcusto.get())), 2))))
                _valorvenda.delete(0, 'end')
                _valorvenda.insert(0, num_brasil(str(sval)))
            except:
                messagebox.showwarning('Valor incorreto', 'Preencha um valor em formato numérico', parent=form_selec)
                _margem.delete(0, 'end')
                _margem.insert(0, '0,00')
                _valorvenda.delete(0, 'end')
                _valorvenda.insert(0, _valorcusto.get())
                update.focus()

        dimension = widgets.geometry(368, 500)
        config = {'title': 'Edição de Itens Materiais',
                  'dimension': dimension,
                  'color': 'pale goldenrod'}
        form_edit = main.form(config)
        wgedit = Widgets(form_edit, 'pale goldenrod')
        wgedit.label('', 5, 0, 0)
        wgedit.label('', 5, 13, 0)
        wgedit.label('', 5, 16, 0)
        if mode in ['edit', 'remove']:
            if prodselect.get('active'):
                c.execute('SELECT * FROM Materiais_Itens WHERE Produto = "' + prodselect.get('active') + '"')
                doc = c.fetchall()
            else:
                messagebox.showerror(title='Aviso', message='Ainda não foi cadastrado nenhum produto!', parent=form_edit)
                update.focus()
        _id = wgedit.textbox('Id ', 2, 1, 1)
        _datacadastro = wgedit.textbox('Data da criação ', 17, 2, 1)
        _dataatualiza = wgedit.textbox('Data última edição ', 17, 3, 1)
        _produto = wgedit.textbox('Produto ', 30, 4, 1)
        c.execute('SELECT Categoria FROM Materiais_Categorias ORDER BY Categoria')
        lista1 = []
        for row in c.fetchall():
            lista1.append(row[0])
        _categoria = wgedit.combobox('Categoria ', 20, lista1, 5, 1, cmd=_categoria_cmd, seek=_categoria_cmd)
        lista2 = ['Compra', 'Venda', 'Revenda']
        _tipo = wgedit.combobox('Tipo de movimento ', 15, lista2, 6, 1, cmd=_tipo_cmd, seek=_tipo_cmd)
        _unidade = wgedit.textbox('Unidade de medida ', 15, 7, 1)
        _marca = wgedit.textbox('Marca ', 30, 8, 1)
        _valorcusto = wgedit.textbox('Custo unitário ', 12, 9, 1, cmd=_valorcusto_cmd)
        _margem = wgedit.textbox('Margem ', 12, 10, 1, cmd=_margem_cmd)
        _valorvenda = wgedit.textbox('Preço unitário ', 12, 11, 1, cmd=_valorvenda_cmd)
        _estoquemin = wgedit.textbox('Estoque mínimo ', 12, 12, 1)

        if mode in ['edit', 'remove']:
            if prodselect.get('active'):
                _id.insert(0, str(doc[0][0]))
                _datacadastro.insert(0, datahora_brasil(doc[0][1]))
                _dataatualiza.insert(0, datahora_brasil(doc[0][2]))
                _produto.insert(0, str(doc[0][3]))
                c.execute('SELECT Categoria FROM Materiais_Categorias WHERE Id = ' + str(doc[0][4]))
                _categoria.insert(0, str(c.fetchone()[0]))
                _tipo.insert(0, lista2[doc[0][5]])
                _unidade.insert(0, str(doc[0][6]))
                _marca.insert(0, str(doc[0][7]))
                _valorcusto.insert(0, num_brasil(str(abs(doc[0][8]))))
                _margem.insert(0, num_brasil(str(abs(doc[0][9]))))
                _valorvenda.insert(0, num_brasil(str(abs(doc[0][10]))))
                _estoquemin.insert(0, num_brasil(str(abs(doc[0][11]))))
        else:
            c.execute('SELECT Id FROM Materiais_Itens ORDER BY Id')
            rows = c.fetchall()
            NewId = 1
            if rows:
                Ids = []
                for row in rows:
                    Ids.append(row[0])
                Ids.sort() 
                NewId = Ids[-1:][0] + 1
            _id.insert(0, str(NewId))
            _datacadastro.insert(0, datahora_brasil(str(datetime.now())))
            _dataatualiza.insert(0, datahora_brasil(str(datetime.now())))
            _tipo.insert(0, 'Venda')
            _unidade.insert(0, 'peça')
            _marca.insert(0, 'Genérico')
            _valorcusto.insert(0, '0,00')
            _margem.insert(0, '0,00')
            _valorvenda.insert(0, '0,00')
            _estoquemin.insert(0, '0,00')

        _id.configure(state='disabled')
        _datacadastro.configure(state='disabled')
        _dataatualiza.configure(state='disabled')
        if mode in ['edit', 'new']:
            _produto.focus()
            update = wgedit.button('Atualizar', updatethis, 20, 0, 14, 1, 2)
            cancel = wgedit.button('Cancelar', exit, 20, 0, 15, 1, 2)
        else:
            _produto.configure(state='disabled')
            _categoria.configure(state='disabled')
            _tipo.configure(state='disabled')
            _unidade.configure(state='disabled')
            _marca.configure(state='disabled')
            _valorcusto.configure(state='disabled')
            _margem.configure(state='disabled')
            _valorvenda.configure(state='disabled')
            _estoquemin.configure(state='disabled')
            wgedit.label('Registro que será excluído!', 0, 14, 1, 2)
            update = wgedit.button('Confirmar exclusão', updatethis, 20, 0, 15, 1, 2)
        if mode in ['edit', 'remove']:
            if not prodselect.get('active'):
                exit()
    
    def seek():
        if materialseek.get():
            nome = "'%" + materialseek.get() + "%'"
            c.execute('SELECT Produto FROM Materiais_Itens WHERE Produto LIKE ' + nome + ' ORDER BY Produto')
            prodselect.delete(0, 'end')
            for row in c.fetchall():
                prodselect.insert('end', row[0])
        else:
            c.execute('SELECT Produto FROM Materiais_Itens ORDER BY Produto')
            doc = c.fetchall()
            prodselect.delete(0, 'end')
            lista = []
            for row in doc:
                prodselect.insert('end', row[0])

    dimension = widgets.geometry(330, 580)
    config = {'title': 'Selecione um Produto e a ação',
              'dimension': dimension,
              'color': 'light goldenrod'}
    form_selec = main.form(config)
    wg = Widgets(form_selec, 'light goldenrod')
    wg.label('', 5, 0, 0)
    wg.label('', 5, 2, 0)
    wg.label('', 5, 6, 0)
    wg.image('',0,'images/produtos.png', 1, 1, 1, 6, imagewidth=(256, 256))
    materialseek = wg.textbox('   Digite: ', 20, 1, 3)
    c.execute('SELECT Produto FROM Materiais_Itens ORDER BY Produto')
    '''
    if tipo == 'mov':
        wg.image('',0,'images/bancos.png', 1, 1, 1, 6)
        c.execute('SELECT NomeBanco FROM Bancos WHERE TipoMov <= 1 ORDER BY NomeBanco')
    elif tipo == 'cc':
        wg.image('',0,'images/creditcard.png', 1, 1, 1, 6)
        c.execute('SELECT NomeBanco FROM Bancos WHERE TipoMov = 2 ORDER BY NomeBanco')
    elif tipo == 'cpp':
        wg.image('',0,'images/creditcard.png', 1, 1, 1, 6)
        c.execute('SELECT NomeBanco FROM Bancos WHERE TipoMov = 3 ORDER BY NomeBanco')
    '''
    doc = c.fetchall()
    lista = []
    for row in doc:
        lista.append(row[0])
    Seek = wg.button('Procurar', seek, 10, 0, 2, 3, 2)
    prodselect = wg.listbox('    Produto:  ', 20, 8, lista, 3, 3)
    materialseek.focus()
    Create = wg.button('Incluir', newreg, 10, 0, 5, 3, 2)
    Edit = wg.button('Editar', editreg, 10, 0, 6, 3, 2)
    Remove = wg.button('Excluir', removereg, 10, 0, 7, 3, 2)

def materiais_movimentos(tipo_mov):
    # tipo_mov options: 0-Recebimento 1-Produção 2-Consumo 3-Venda
    def exitsc():
        form_selec.destroy()

    def newreg():
        edit('new')

    def dupreg():
        edit('dup')

    def editreg():
        edit('edit')

    def pagreg():
        if itemselect.selection():
            for i in itemselect.selection():
                gridselected = str(itemselect.item(i, 'text'))
            mens = ('Tem certeza que deseja confirmar a baixa dos documentos selecionados?\n\n' +
                    'Os documentos serão baixados na data de hoje: ' + data_brasil(ActualDate))
            if in_value:
                title = 'Baixa de Recebimentos'
            elif out_value:
                title = 'Baixa de Pagamentos'
            else:
                title = 'Confirmação de transferências'
            if messagebox.askyesno(title=title, message=mens, parent=form_selec):
                notrelease = []
                for i in itemselect.selection():
                    c.execute('SELECT DataPago, TipoMov, Banco, Fatura FROM Diario WHERE Id = ' + str(itemselect.item(i, 'text')))
                    opt = c.fetchone()
                    if not opt[0]:
                        command = []
                        command.append("UPDATE Diario SET ")
                        command.append("DataPago = '" + ActualDate + "' ")
                        command.append("WHERE Id = " + str(itemselect.item(i, 'text')))
                        c.execute(str.join('', command))
                        conn.commit()
                        if opt[1] == 3:
                            command = []
                            command.append("UPDATE Diario SET ")
                            command.append("DataPago = '" + ActualDate + "' ")
                            command.append("WHERE Id = " + str(itemselect.item(i, 'text') - 1))
                            c.execute(str.join('', command))
                            conn.commit()
                            opt_bank_v = str(opt[2])
                            c.execute('UPDATE Diario SET DataPago = "' + ActualDate + '" WHERE Fatura = "' + opt[3] + '" AND Banco = ' + opt_bank_v + '')
                            messagebox.showinfo('Aviso!', 'Todos os pagamentos referente à fatura ' + opt[3] + ' foram baixados no sistema.', parent=form_selec)
                        elif opt[1] == 4:
                            command = []
                            command.append("UPDATE Diario SET ")
                            command.append("DataPago = '" + ActualDate + "' ")
                            command.append("WHERE Id = " + str(itemselect.item(i, 'text') + 1))
                            c.execute(str.join('', command))
                            conn.commit()
                            if opt[3]:
                                c.execute('SELECT Banco FROM Diario WHERE Id = ' + str(itemselect.item(i, 'text') + 1))
                                opt_bank = c.fetchone()
                                c.execute('UPDATE Diario SET DataPago = "' + ActualDate + '" WHERE Fatura = "' + opt[3] + '" AND Banco = ' + str(opt_bank[0]) + '')
                                messagebox.showinfo('Aviso!', 'Todos os pagamentos referente à fatura ' + opt[3] + ' foram baixados no sistema.', parent=form_selec)
                    else:
                        notrelease.append(str(itemselect.item(i, 'text')))
                if notrelease:
                    if len(notrelease) < 2:
                        message = 'O documento ' + notrelease[0] + ' já foi baixado! Revise.'
                    else:
                        message = 'Os documentos ' 
                        for i in notrelease:
                            if i not in notrelease[-2:]:
                                message = message + str(i) + ', '
                            elif i == notrelease[-2]:
                                message = message + str(i) + ' e '
                            else:
                                message = message + str(i) + ' '
                        message = message + ' já foram baixados! Revise.'
                    messagebox.showerror(title='Aviso', message=message, parent=form_selec)
        else:
            messagebox.showerror(title='Aviso', message='Não foi selecionado nenhum documento para a baixa!', parent=form_selec)
    
    def removereg():
        edit('remove')

    def edit(mode):
        def edit_desconto(data):
            if tipo_mov in [0, 1, 2]:
                try:
                    valuedesc = num_usa(_vldesconto.get())
                    valueajust = float(num_usa(_totalcusto.get())) - float(valuedesc)
                    valueajust = round(valueajust, 2)
                    _vldesconto.delete(0, 'end')
                    _vldesconto.insert(0, num_brasil(valuedesc))
                    _vltotal.delete(0, 'end')
                    _vltotal.insert(0, num_brasil(str(valueajust)))
                except:
                    msg = 'Digite o número num formato válido.'
                    messagebox.showerror(title='Aviso', message=msg, parent=form_edit)
                    _vldesconto.delete(0, 'end')
                    _vldesconto.insert(0, '0,00')
                    _vltotal.delete(0, 'end')
                    _vltotal.insert(0, _totalcusto.get())
                    update.focus()
            else:
                valorcomp = _vltotal['text'].split(' ')
                try:
                    valuedesc = num_usa(_troco.get())
                    valueajust = float(valuedesc) + float(num_usa(valorcomp[2]))
                    valueajust = round(valueajust, 2)
                    if float(valuedesc) < 0.0:
                        valueajust = valueajust / 0
                    _dinheiro.delete(0, 'end')
                    _dinheiro.insert(0, num_brasil(str(valueajust)))
                    _troco.delete(0, 'end')
                    _troco.insert(0, num_brasil(str(float(valuedesc))))
                except:
                    msg = 'Confira o valor informado!'
                    messagebox.showerror(title='Aviso', message=msg, parent=form_edit)
                    _troco.delete(0, 'end')
                    _troco.insert(0, '0,00')
                    _dinheiro.delete(0, 'end')
                    _dinheiro.insert(0, valorcomp[2])
                    update.focus()
        
        def edit_total(data):
            if tipo_mov in [0, 1, 2]:
                try:
                    valuedesc = num_usa(_vltotal.get())
                    valueajust = float(num_usa(_totalcusto.get())) - float(valuedesc)
                    valueajust = round(valueajust, 2)
                    _vldesconto.delete(0, 'end')
                    _vldesconto.insert(0, num_brasil(str(valueajust)))
                    _vltotal.delete(0, 'end')
                    _vltotal.insert(0, num_brasil(valuedesc))
                except:
                    msg = 'Digite o número num formato válido.'
                    messagebox.showerror(title='Aviso', message=msg, parent=form_edit)
                    _vldesconto.delete(0, 'end')
                    _vldesconto.insert(0, '0,00')
                    _vltotal.delete(0, 'end')
                    _vltotal.insert(0, _totalcusto.get())
                    update.focus()
            else:
                valorcomp = _vltotal['text'].split(' ')
                try:
                    valuedesc = num_usa(_dinheiro.get())
                    valueajust = float(valuedesc) - float(num_usa(valorcomp[2]))
                    valueajust = round(valueajust, 2)
                    if valueajust < 0.0:
                        valueajust = valueajust / 0
                    _troco.delete(0, 'end')
                    _troco.insert(0, num_brasil(str(valueajust)))
                    _dinheiro.delete(0, 'end')
                    _dinheiro.insert(0, num_brasil(str(float(valuedesc))))
                except:
                    msg = 'Confira o valor informado!'
                    messagebox.showerror(title='Aviso', message=msg, parent=form_edit)
                    _troco.delete(0, 'end')
                    _troco.insert(0, '0,00')
                    _dinheiro.delete(0, 'end')
                    _dinheiro.insert(0, valorcomp[2])
                    update.focus()

        def edit_new():
            edit_item('new')

        def edit_exist():
            if tipo_mov != 3:
                edit_item('edit')
            else:
                datas = [_date(_dataefetiva.get()), _date(_datanova.get())]
                altera = {}
                for row in datas:
                    altera[row] = []
                    c.execute("SELECT MAX(MovimentoFin), SUM(TotalCusto), SUM(TotalVenda) FROM Materiais_Documentos WHERE Tipo = 3 AND DataEfetiva = '" + row + "'")
                    altera[row].append(c.fetchone())
                command = []
                command.append("UPDATE Materiais_Documentos SET ")
                command.append("DataAtualiza = '" + str(datetime.now()) + "', ")
                command.append("DataDoc = '" + _date(_datanova.get()) + "', ")
                command.append("DataEfetiva = '" + _date(_datanova.get()) + "', ")
                command.append("MovimentoFin = 0 ")
                command.append("WHERE Id = " + gridselected)
                c.execute(str.join('', command))
                command = []
                command.append("UPDATE Materiais_Movimentos SET ")
                command.append("DataAtualiza = '" + str(datetime.now()) + "', ")
                command.append("DataEfetiva = '" + _date(_datanova.get()) + "' ")
                command.append("WHERE Documento = " + gridselected)
                c.execute(str.join('', command))
                for row in datas:
                    c.execute("SELECT MAX(MovimentoFin), SUM(TotalCusto), SUM(TotalVenda) FROM Materiais_Documentos WHERE Tipo = 3 AND DataEfetiva = '" + row + "'")
                    altera[row].append(c.fetchone())
                if not altera[datas[0]][1][1]:
                    c.execute('DELETE FROM Diario WHERE Id = ' + str(altera[datas[0]][0][0]))
                elif altera[datas[0]][1][1]:
                    command = []
                    command.append("UPDATE Diario SET ")
                    command.append("DataLastUpdate = '" + ActualDate + "', ")
                    command.append("DataDoc = '" + datas[0] + "', ")
                    command.append("DataVenc = '" + datas[0] + "', ")
                    command.append("DataPago = '" + datas[0] + "', ")
                    command.append("Valor = " + str(altera[datas[0]][1][2]) + ", ")
                    command.append("Descricao = 'Venda de balcão - Custo: " + num_brasil(str(round(altera[datas[0]][1][1], 2))) + "' ")
                    command.append("WHERE Id = " + str(altera[datas[0]][0][0]))
                    c.execute(str.join('', command))
                    #c.execute('UPDATE Materiais_Documentos SET MovimentoFin = ' + str(altera[datas[0]][0][0]) + ' WHERE Id = ' + gridselected)
                if not altera[datas[1]][0][1]:
                    c.execute('SELECT MAX(Id) FROM Diario ORDER BY Id')
                    rows = c.fetchone()
                    NewId = 1
                    if rows[0]:
                        NewId = rows[0] + 1
                    command = []
                    command.append("INSERT INTO Diario VALUES (")
                    command.append(str(NewId) + ", ")
                    command.append("'" + ActualDate + "', ")
                    command.append("'" + ActualDate + "', ")
                    command.append("'" + datas[1] + "', ")
                    command.append("'" + datas[1] + "', ")
                    command.append("'" + datas[1] + "', ")
                    command.append("1, ") # Sempre a receita cai no primeiro PARCEIRO cadastrado.
                    command.append("1, ") # Sempre a receita cai no primeiro BANCO cadastrado.
                    command.append("'', ")
                    command.append("'Venda de balcão - Custo: " + num_brasil(str(round(altera[datas[1]][1][1], 2))) + "', ")
                    command.append(str(altera[datas[1]][1][2]) + ', ')
                    command.append("0, ")
                    command.append('21') # Sempre usa o código de receita 21 - VENDA DIRETA
                    command.append(")")
                    c.execute(str.join('', command))
                    c.execute('UPDATE Materiais_Documentos SET MovimentoFin = ' + str(NewId) + ' WHERE Id = ' + gridselected)
                elif altera[datas[1]][0][1]:
                    command = []
                    command.append("UPDATE Diario SET ")
                    command.append("DataLastUpdate = '" + ActualDate + "', ")
                    command.append("DataDoc = '" + datas[1] + "', ")
                    command.append("DataVenc = '" + datas[1] + "', ")
                    command.append("DataPago = '" + datas[1] + "', ")
                    command.append("Valor = " + str(altera[datas[1]][1][2]) + ", ")
                    command.append("Descricao = 'Venda de balcão - Custo: " + num_brasil(str(round(altera[datas[1]][1][1], 2))) + "' ")
                    command.append("WHERE Id = " + str(altera[datas[1]][1][0]))
                    c.execute(str.join('', command))
                    c.execute('UPDATE Materiais_Documentos SET MovimentoFin = ' + str(altera[datas[1]][0][0]) + ' WHERE Id = ' + gridselected)
                conn.commit()
                exit()
                seek()

        def edit_remove():
            if editemselect.selection():
                for i in editemselect.selection():
                    _actualitem = str(editemselect.item(i, 'text'))
                    for row in itensdata:
                        if row[0] == _actualitem:
                            _actual = itensdata.index(row)
                            if itensdata[_actual] == max(itensdata):
                                _it_id[0] = str(int(_it_id[0]) - 1)
                            break
                del itensdata[_actual]
                updateitens(itensdata)
            else:
                messagebox.showerror(title='Aviso', message='Nenhum item foi selecionado!', parent=form_edit)
                update.focus()

        def edit_item(imode):
            def edit_exit():
                form_edit_item.destroy()

            def _it_produto_cmd(data):
                if _it_produto.get():
                    widgets.combobox_return(_it_produto, produtos.keys())
                    if tipo_mov in [2, 3, 4]:
                        _it_vlunit.configure(state='normal')
                    _it_vlunit.delete(0, 'end')
                    if _it_produto.get():
                        _it_vlunit.insert(0, num_brasil(str(produtos[_it_produto.get()][0])))
                    if tipo_mov in [2, 3, 4]:
                        _it_vlunit.configure(state='disabled')
                    _it_unidade.configure(state='normal')
                    _it_unidade.delete(0, 'end')
                    if _it_produto.get():
                        if produtos[_it_produto.get()][1][0].lower() == 'k' and tipo_mov in [3, 4]:
                            _it_unidade.insert(0, 'gramas *')
                        else:
                            _it_unidade.insert(0, produtos[_it_produto.get()][1])
                    _it_unidade.configure(state='disabled')
                    if tipo_mov in [2, 3, 4]:
                        c.execute('SELECT SUM(Qtde) FROM Materiais_Movimentos JOIN Materiais_Itens ON Materiais_Itens.Id = Materiais_Movimentos.Produto WHERE Materiais_Itens.Produto = "' + _it_produto.get() + '"')
                        value_saldo = c.fetchone()
                        _it_saldoatual['text'] = ''
                        if _it_produto.get():
                            if value_saldo[0]:                        
                                _it_saldoatual['text'] = num_brasil(str(round(value_saldo[0], 2)))
                            else:
                                _it_saldoatual['text'] = '0,00'
                    valor_calc(None)

            def valor_calc(data):
                try:
                    brqt = str(float(num_usa(_it_qtde.get())))
                    _it_qtde.delete(0, 'end')
                    _it_qtde.insert(0, num_brasil(brqt))
                except:
                    _it_qtde.delete(0, 'end')
                    _it_qtde.insert(0, '0,00')
                if _it_qtde.get():
                    try:
                        try:
                            valueunit = str(float(num_usa(_it_vlunit.get())))
                        except:
                            valueunit = '0.0'
                        if _it_unidade.get() == 'gramas *' and tipo_mov in [3, 4]:
                            div = 1000.0
                        else:
                            div = 1.0
                        valuedit = float(valueunit) * (float(num_usa(_it_qtde.get())) / div)
                        valuedit = round(valuedit, 2)
                        if tipo_mov == 0:
                            _it_vlunit.delete(0, 'end')
                            _it_vlunit.insert(0, num_brasil(valueunit))
                            _it_vlcusto.configure(state='normal')
                            _it_vlcusto.delete(0, 'end')
                            _it_vlcusto.insert(0, num_brasil(str(valuedit)))
                            _it_vlcusto.configure(state='disabled')
                        elif tipo_mov in [1, 2, 3, 4]:
                            if tipo_mov in [2, 3, 4]:
                                _it_vlunit.configure(state='normal')
                            _it_vlunit.delete(0, 'end')
                            _it_vlunit.insert(0, num_brasil(valueunit))
                            if tipo_mov in [2, 3, 4]:
                                _it_vlunit.configure(state='disabled')
                            _it_vlvenda.configure(state='normal')
                            _it_vlvenda.delete(0, 'end')
                            _it_vlvenda.insert(0, num_brasil(str(valuedit)))
                            _it_vlvenda.configure(state='disabled')
                    except:
                        msg = 'Digite o número num formato válido.'
                        messagebox.showerror(title='Aviso', message=msg, parent=form_edit)
                        _it_vlunit.delete(0, 'end')
                        _it_vlunit.insert(0, '0,00')
                        _it_qtde.delete(0, 'end')
                        _it_qtde.insert(0, '0,00')

            def _it_qtde_cmd(data):
                try:
                    brqt = str(float(num_usa(_it_qtde.get())))
                    _it_qtde.delete(0, 'end')
                    _it_qtde.insert(0, num_brasil(brqt))
                except:
                    _it_qtde.delete(0, 'end')
                    _it_qtde.insert(0, '0,00')

            def itens_dbfake():
                executedit = True
                if _it_produto.get() == '':
                    executedit = False
                if _it_vlunit.get() == '':
                    executedit = False
                if _it_qtde.get() == '':
                    executedit = False
                if tipo_mov in [2, 3, 4] and executedit:
                    actualqt = _it_qtde.get()
                    if _it_unidade.get() == 'gramas *':
                        actualqt = num_brasil(str(round(float(num_usa(_it_qtde.get())) / 1000, 2)))
                        valid = float(num_usa(_it_saldoatual['text'])) - float(num_usa(actualqt))
                    else:
                        valid = float(num_usa(_it_saldoatual['text'])) - float(num_usa(_it_qtde.get()))
                    if valid < 0:
                        executedit = False
                if executedit:
                    if tipo_mov in [0]:
                        valor_calc(None)
                        if imode == 'edit':
                            lin = itensdata[_actual][0]
                            itensdata[_actual] = [lin, produtos[_it_produto.get()][2], _it_produto.get(), _it_vlunit.get(), _it_qtde.get(), _it_vlcusto.get()]
                        else:
                            itensdata.append([_it_id[0], produtos[_it_produto.get()][2], _it_produto.get(), _it_vlunit.get(), _it_qtde.get(), _it_vlcusto.get()])
                            _it_id[0] = str(int(_it_id[0]) + 1)
                    elif tipo_mov in [1, 2, 3, 4]:
                        valor_calc(None)
                        if imode == 'edit':
                            lin = itensdata[_actual][0]
                            itensdata[_actual] = [lin, produtos[_it_produto.get()][2], _it_produto.get(), _it_vlunit.get(), _it_qtde.get(), _it_vlvenda.get()]
                        else:
                            if tipo_mov in [1]:
                                itensdata.append([_it_id[0], produtos[_it_produto.get()][2], _it_produto.get(), _it_vlunit.get(), _it_qtde.get(), _it_vlvenda.get()])
                            else:
                                itensdata.append([_it_id[0], produtos[_it_produto.get()][2], _it_produto.get(), _it_vlunit.get(), actualqt, _it_vlvenda.get()])
                            _it_id[0] = str(int(_it_id[0]) + 1)
                    form_edit_item.destroy()
                    updateitens(itensdata)
                else:
                    msg = 'Verifique o produto e valores informados. Não é possível atualizar.'
                    messagebox.showerror(title='Aviso', message=msg, parent=form_edit)

            Cons = 0  # pra diferenciar quando é baixa que exibe o saldo atual de mercadoria
            dimension = widgets.geometry(248, 360)
            title_item = 'Editando'
            if imode == 'new':
                title_item = 'Incluindo novo item'
            if imode == 'edit':
                title_item = 'Editando valores'
            config = {'title': title_item,
                    'dimension': dimension,
                    'color': 'pale goldenrod'}
            form_edit_item = main.form(config)
            wgedit_item = Widgets(form_edit_item, 'pale goldenrod')
            wgedit_item.label('', 5, 0, 0)
            wgedit_item.label('', 5, 6, 0)
            wgedit_item.label('', 5, 7, 0)
            wgedit_item.label('', 5, 9, 0)
            if imode in ['edit', 'remove']:
                if editemselect.selection():
                    for i in editemselect.selection():
                        _actual = str(editemselect.item(i, 'text'))
                        for row in itensdata:
                            if row[0] == _actual:
                                _actual = itensdata.index(row)
                                break
                else:
                    messagebox.showerror(title='Aviso', message='Nenhum item foi selecionado!', parent=form_edit)
                    edit_exit()
            if tipo_mov == 0:
                c.execute('SELECT Produto, ValorCusto, Unidade, Id FROM Materiais_Itens WHERE Tipo IN (0, 2) ORDER BY Produto')
            elif tipo_mov == 1:
                c.execute('SELECT Produto, ValorVenda, Unidade, Id FROM Materiais_Itens WHERE Tipo = 1 ORDER BY Produto')
            elif tipo_mov == 2:
                c.execute('SELECT Produto, ValorCusto, Unidade, Id FROM Materiais_Itens ORDER BY Produto')
            elif tipo_mov in [3, 4]:
                c.execute('SELECT Produto, ValorVenda, Unidade, Id, ValorCusto FROM Materiais_Itens WHERE Tipo IN (1, 2) ORDER BY Produto')
            listaedit1 = []
            produtos = {}
            for row in c.fetchall():
                listaedit1.append(row[0])
                produtos[row[0]] = [row[1], row[2], row[3]]
            _it_produto = wgedit_item.combobox('Produto ', 20, listaedit1, 1, 1, cmd=_it_produto_cmd, seek=_it_produto_cmd)
            _it_vlunit = wgedit_item.textbox('Preço unitário ', 14, 2, 1, cmd=valor_calc)
            _it_qtde = wgedit_item.textbox('Quantidade ', 10, 3, 1, cmd=valor_calc)
            _it_unidade = wgedit_item.textbox('Unidade ', 14, 4, 1)
            if tipo_mov == 0:
                _it_vlcusto = wgedit_item.textbox('Valor ', 14, 5, 1)
                _it_vlvenda = None
            elif tipo_mov in [1, 2, 3, 4]:
                _it_vlcusto = None
                _it_vlvenda = wgedit_item.textbox('Valor ', 14, 5, 1)
            if tipo_mov in [2, 3, 4]:
                rot = wgedit_item.label('Qtde estoque ', 14, 6, 1, stick='E')
                rot['anchor'] = 'e'
                _it_saldoatual = wgedit_item.label('0,00 ', 14, 6, 2, stick='W')
                _it_saldoatual['bg'] = 'yellow'
                _it_saldoatual['fg'] = 'blue'
                _it_saldoatual['relief'] = 'groove'
                Cons = 1
            _it_documento = None

            if imode in ['edit', 'remove']:
                _it_produto.insert(0, itensdata[_actual][2])
                _it_vlunit.insert(0, itensdata[_actual][3])
                _it_qtde.insert(0, itensdata[_actual][4])
                _it_unidade.insert(0, produtos[str(itensdata[_actual][2])][1])
                if tipo_mov == 0:
                    _it_vlcusto.insert(0, itensdata[_actual][5])
                elif tipo_mov in [1, 2, 3, 4]:
                    _it_vlvenda.insert(0, itensdata[_actual][5])
                _it_produto_cmd(None)
            else:
                pass
                # _it_datacadastro.insert(0, data_brasil(ActualDate))
            _it_unidade.configure(state='disabled')
            if tipo_mov == 0:
                _it_vlcusto.configure(state='disabled')
            elif tipo_mov in [1, 2, 3, 4]:
                if tipo_mov in [2, 3, 4]:
                    _it_vlunit.configure(state='disabled')
                _it_vlvenda.configure(state='disabled')
            _it_produto.focus()
            updateitem = wgedit_item.button('Atualizar', itens_dbfake, 20, 0, 7 + Cons, 1, 2)
            cancelitem = wgedit_item.button('Cancelar', edit_exit, 20, 0, 8 + Cons, 1, 2)

        def edit_cancel():
            mens = 'Tem certeza que deseja sair? A edição realizada será perdida.'
            if messagebox.askyesno(title="Atualização de sistema", message=mens, parent=form_edit):
                exit()

        def edit_cp(Venda=False):
            def cp_update():
                cpvalid = True
                if not _cpbanco.get() or not _cpcategoria.get() or not _cpdatavenc.get():
                    msg = 'Os dados estão incompletos.'
                    messagebox.showerror(title='Aviso', message=msg)
                    cpvalid = False
                if cpvalid:
                    updatethis_efetivate()
                    c.execute('SELECT MovimentoFin FROM Materiais_Documentos WHERE Id = ' + _id.get())
                    resp = c.fetchone()[0]
                    command = []
                    if not resp:
                        c.execute('SELECT Id FROM Diario ORDER BY Id')
                        rows = c.fetchall()
                        NewId = 1
                        if rows:
                            Ids = []
                            for row in rows:
                                Ids.append(row[0])
                            Ids.sort() 
                            NewId = Ids[-1:][0] + 1
                        command.append("INSERT INTO Diario VALUES (")
                        command.append(str(NewId) + ", ")
                        command.append("'" + ActualDate + "', ")
                        command.append("'" + ActualDate + "', ")
                        command.append("'" + _date(_dataefetiva.get()) + "', ")
                        command.append("'" + _date(_cpdatavenc.get()) + "', ")
                        command.append("'" + _date(_cpdatapago.get()) + "', ")
                        c.execute('SELECT Id FROM Parceiros WHERE Nome = "' + _parceiro.get() + '"')
                        command.append(str(c.fetchone()[0]) + ", ")
                        c.execute('SELECT Id, TipoMov FROM Bancos WHERE NomeBanco = "' + _cpbanco.get() + '"')
                        resp_bank = c.fetchone()
                        command.append(str(resp_bank[0]) + ", ")
                        new_fat = ''
                        if resp_bank[1] == 2:
                            new_fat = _cpdatavenc.get()[0:7]
                        command.append("'" + new_fat + "', ")
                        command.append("'Compra de materiais', ")
                        valorfill1 = _vltotal.get().replace('.', '')
                        valorfill1 = valorfill1.replace(',', '.')
                        command.append('-' + valorfill1 + ', ')
                        command.append("1, ")
                        c.execute('SELECT Id FROM Categorias WHERE Categoria = "' + _cpcategoria.get() + '"')
                        command.append(str(c.fetchone()[0]))
                        command.append(")")
                        c.execute(str.join('', command))
                        c.execute('UPDATE Materiais_Documentos SET MovimentoFin = ' + str(NewId) + ' WHERE Id = ' + _id.get())
                    else:
                        command.append("UPDATE Diario SET ")
                        command.append("DataLastUpdate = '" + ActualDate + "', ")
                        command.append("DataDoc = '" + _date(_dataefetiva.get()) + "', ")
                        c.execute('SELECT Id FROM Parceiros WHERE Nome = "' + _parceiro.get() + '"')
                        command.append("Parceiro = " + str(c.fetchone()[0]) + ", ")
                        c.execute('SELECT Id, TipoMov FROM Bancos WHERE NomeBanco = "' + _cpbanco.get() + '"')
                        resp_bank = c.fetchone()
                        if resp_bank[1] == 2:
                            new_fat = _date(_cpdatavenc.get())[0:7]
                            command.append("Fatura = '" + new_fat + "', ")
                        else:
                            command.append("Fatura = '', ")
                        command.append("Banco = " + str(resp_bank[0]) + ", ")
                        c.execute('SELECT Id FROM Categorias WHERE Categoria = "' + _cpcategoria.get() + '"')
                        command.append("CategoriaMov = " + str(c.fetchone()[0]) + ", ")
                        valorfill1 = _vltotal.get().replace('.', '')
                        valorfill1 = valorfill1.replace(',', '.')
                        command.append("Valor = -" + valorfill1 + ", ")
                        command.append("DataVenc = '" + _date(_cpdatavenc.get()) + "', ")
                        command.append("DataPago = '" + _date(_cpdatapago.get()) + "' ")
                        command.append("WHERE Id = " + str(resp))
                        c.execute(str.join('', command))
                        #if mode == 'remove':
                        #    command.append('DELETE FROM Diario WHERE Id = ' + _id.get())
                    conn.commit()
                    form_edit_pagar.destroy()
                    exit()

            def _cpbanco_cmd(value=None):
                widgets.combobox_return(_cpbanco, lista1)
                if _cpbanco.get() != '':
                    c.execute('SELECT GeraFatura, DiaVenc FROM Bancos WHERE NomeBanco = "' + _cpbanco.get() + '"')
                    resp = c.fetchone()
                    try:
                        if resp[0]:
                            hoje = datetime.strptime(_datadoc.get(), '%d/%m/%Y') + timedelta(days=10)
                            dia = hoje.day
                            mes = hoje.month
                            ano = hoje.year
                            if resp[1] >= dia:
                                pass
                            else:
                                mes = mes + 1
                                if mes == 13:
                                    mes = 1
                                    ano += 1
                            new_fat = str(ano) + '-' + str(mes).zfill(2)
                            new_venc = datetime(ano, mes, resp[1]).strftime('%d/%m/%Y')
                            _cpdatavenc.delete(0, 'end')
                            _cpdatavenc.insert(0, new_venc)
                            _cpdatapago.delete(0, 'end')
                            _cpdatapago.configure(state='disabled')
                        else:
                            _cpdatapago.configure(state='normal')
                    except:
                        _cpdatapago.configure(state='normal')

            def _cpcategoria_cmd(value=None):
                widgets.combobox_return(_cpcategoria, lista2)

            def _cpdatavenc_cmd(Value=None):
                data_cmd(_cpdatavenc)

            def _cpdatapago_cmd(Value=None):
                data_cmd(_cpdatapago)

            if Venda:
                updatethis_efetivate()
                if tipo_mov == 3:
                    c.execute("SELECT MAX(MovimentoFin), SUM(TotalCusto), SUM(TotalVenda) FROM Materiais_Documentos WHERE Tipo = 3 AND DataEfetiva = '" + ActualDate + "'")
                    resp = c.fetchone()
                    command = []
                    if not resp[0]:
                        if mode != 'remove':
                            c.execute('SELECT MAX(Id) FROM Diario ORDER BY Id')
                            rows = c.fetchone()
                            NewId = 1
                            if rows[0]:
                                NewId = rows[0] + 1
                            command.append("INSERT INTO Diario VALUES (")
                            command.append(str(NewId) + ", ")
                            command.append("'" + ActualDate + "', ")
                            command.append("'" + ActualDate + "', ")
                            command.append("'" + ActualDate + "', ")
                            command.append("'" + ActualDate + "', ")
                            command.append("'" + ActualDate + "', ")
                            command.append("1, ") # Sempre a receita cai no primeiro PARCEIRO cadastrado.
                            command.append("1, ") # Sempre a receita cai no primeiro BANCO cadastrado.
                            command.append("'', ")
                            command.append("'Venda de balcão - Custo: " + num_brasil(str(round(resp[1], 2))) + "', ")
                            command.append(str(resp[2]) + ', ')
                            command.append("0, ")
                            command.append('21') # Sempre usa o código de receita 21 - VENDA DIRETA
                            command.append(")")
                            c.execute(str.join('', command))
                            c.execute('UPDATE Materiais_Documentos SET MovimentoFin = ' + str(NewId) + ' WHERE Id = ' + _id)
                        else:
                            command.append('DELETE FROM Diario WHERE Id = ' + str(resp[0]))
                    else:
                        command.append("UPDATE Diario SET ")
                        command.append("Valor = " + str(resp[2]) + ", ")
                        command.append("Descricao = 'Venda de balcão - Custo: " + num_brasil(str(round(resp[1], 2))) + "' ")
                        command.append("WHERE Id = " + str(resp[0]))
                        c.execute(str.join('', command))
                    conn.commit()
                    seek()
                    exit()
                if tipo_mov == 4:
                    c.execute('SELECT MovimentoFin, TotalCusto, TotalVenda FROM Materiais_Documentos WHERE Id = ' + _id)
                    resp = c.fetchone()
                    if not resp:
                        resp = [0]
                    if not resp[0]:
                        if mode != 'remove':
                            c.execute('SELECT MAX(Id) FROM Diario ORDER BY Id')
                            rows = c.fetchone()
                            NewId = 1
                            if rows[0]:
                                NewId = rows[0] + 1
                            custo = resp[1] / len(_vencimentos)
                            a = 0
                            parcel = False
                            if len(_vencimentos) > 1:
                                parcel = True
                            for row in _vencimentos:
                                command = []
                                command.append("INSERT INTO Diario VALUES (")
                                command.append(str(NewId) + ", ")
                                command.append("'" + ActualDate + "', ")
                                command.append("'" + ActualDate + "', ")
                                command.append("'" + ActualDate + "', ")
                                command.append("'" + str(row['Vencimento'].date()) + "', ")
                                if row['Vencimento'].date() == datetime.now().date():
                                    command.append("'" + str(row['Vencimento'].date()) + "', ")
                                else:
                                    command.append("'', ")
                                c.execute('SELECT Id FROM Parceiros WHERE Nome = "' + _parceiro.get() + '"')
                                command.append(str(c.fetchone()[0]) + ", ")
                                command.append("1, ") # Sempre a receita cai no primeiro BANCO cadastrado.
                                command.append("'', ")
                                acres = ''
                                if parcel:
                                    acres = ' (' + str(a + 1) + '/' + str(len(_vencimentos)) + ')'
                                command.append("'Venda cliente - Custo: " + num_brasil(str(round(custo, 2))) + acres + "', ")
                                command.append(str(round(row['Valor'], 2)) + ', ')
                                command.append("0, ")
                                command.append('21') # Sempre usa o código de receita 21 - VENDA DIRETA
                                command.append(")")
                                c.execute(str.join('', command))
                                if a == 0:
                                    c.execute('UPDATE Materiais_Documentos SET MovimentoFin = ' + str(NewId) + ' WHERE Id = ' + _id)
                                a += 1
                                NewId += 1
                    else:
                        command.append("UPDATE Diario SET ")
                        command.append("Valor = " + str(resp[2]) + ", ")
                        command.append("Descricao = 'Venda de balcão - Custo: " + num_brasil(str(round(resp[1], 2))) + "' ")
                        command.append("WHERE Id = " + str(resp[0]))
                        c.execute(str.join('', command))
                    conn.commit()
                    seek()
                    exit()
            else:
                dimension = widgets.geometry(170, 350)
                title = 'Edição de dados no Contas à Pagar'
                config = {'title': title,
                        'dimension': dimension,
                        'color': 'pale goldenrod'}
                form_edit_pagar = main.form(config)
                cpedit = Widgets(form_edit_pagar, 'pale goldenrod')
                cpedit.label('', 5, 0, 0)
                cpedit.label('', 5, 5, 0)
                c.execute('SELECT NomeBanco FROM Bancos ORDER BY NomeBanco')
                lista1 = []
                for row in c.fetchall():
                    lista1.append(row[0])
                _cpbanco = cpedit.combobox('Banco ', 20, lista1, 1, 1, cmd=_cpbanco_cmd, seek=_cpbanco_cmd)
                c.execute('SELECT Categoria FROM Categorias WHERE TipoMov = 1 ORDER BY Categoria')
                lista2 = []
                for row in c.fetchall():
                    lista2.append(row[0])
                _cpcategoria = cpedit.combobox('Categoria ', 20, lista2, 2, 1, cmd=_cpcategoria_cmd, seek=_cpcategoria_cmd)
                _cpdatavenc = cpedit.textbox('Vencimento ', 12, 3, 1, cmd=_cpdatavenc_cmd)
                _cpdatapago = cpedit.textbox('Pagamento ', 12, 4, 1, cmd=_cpdatapago_cmd)
                _cpdatavenc.insert(0, data_brasil(ActualDate))
                _cpatualiza = cpedit.button('Atualizar', cp_update, 15, 0, 6, 1, 2)
                if tipo_mov in [0, 1, 2]:
                    c.execute('SELECT MovimentoFin FROM Materiais_Documentos WHERE Id = ' + _id.get())
                elif tipo_mov in [3]:
                    c.execute('SELECT MovimentoFin FROM Materiais_Documentos WHERE Id = ' + _id)
                resp = c.fetchone()
                if resp:
                    c.execute('SELECT NomeBanco, Categorias.Categoria FROM Diario JOIN Bancos ON Bancos.Id = Banco JOIN Categorias ON Categorias.Id = CategoriaMov WHERE Diario.Id = ' + str(resp[0]))
                    resp2 = c.fetchone()
                    if resp2:
                        _cpbanco.insert(0, resp2[0])
                        _cpcategoria.insert(0, resp2[1])
                _cpbanco.focus()

        def exit():
            form_edit.destroy()
            Create.focus()
        
        def updateitens(doc, firstuse=False):
            edlista.clear()
            edlistaord.clear()
            editemselect.delete(*editemselect.get_children())
            # soma['text'] = 'Valor Total: 0,00'
            if firstuse:
                if mode == 'dup':
                    idref = str(dup_id)
                else:
                    if tipo_mov in [0, 1, 2]:
                        idref = _id.get()
                    elif tipo_mov in [3, 4]:
                        idref = _id
                if tipo_mov == 0:
                    c.execute('SELECT Materiais_Movimentos.Id, Materiais_Movimentos.Produto, Materiais_Itens.Produto, VlUnit, Qtde, VlCusto '\
                            'FROM Materiais_Movimentos JOIN Materiais_Itens ON Materiais_Itens.Id = Materiais_Movimentos.Produto '\
                            'WHERE Materiais_Movimentos.Documento = ' + idref)
                elif tipo_mov in [1, 2, 3, 4]:
                    c.execute('SELECT Materiais_Movimentos.Id, Materiais_Movimentos.Produto, Materiais_Itens.Produto, VlUnit, Qtde, VlVenda '\
                            'FROM Materiais_Movimentos JOIN Materiais_Itens ON Materiais_Itens.Id = Materiais_Movimentos.Produto '\
                            'WHERE Materiais_Movimentos.Documento = ' + idref)
                values = c.fetchall()
                doc = []
                for row in values:
                    newitem = []
                    index = 0
                    for rowitem in row:
                        if index in [0]:
                            if mode == 'dup':
                                newitem.append(_it_id[0])
                                _it_id[0] = str(int(_it_id[0]) + 1)
                            else:
                                newitem.append(str(rowitem))
                        elif index in [3, 4, 5]:
                            if index == 4 and tipo_mov in [2, 3, 4]:
                                newitem.append(num_brasil(str(-rowitem)))
                            else:
                                newitem.append(num_brasil(str(rowitem)))
                        else:
                            newitem.append(rowitem)
                        index += 1
                    doc.append(newitem)
                    itensdata.append(newitem)
            somacusto = 0.0
            somalc = 0.0
            for row in doc:
                texts = []
                rowact = 0
                c.execute('SELECT ValorCusto FROM Materiais_Itens WHERE Id = ' + str(row[1]))
                valueresp = c.fetchone()
                somacusto += round(float(valueresp[0]) * float(num_usa(row[4])), 2)
                for rows in row[1:]:
                    if rowact == 4:
                        somalc += float(num_usa(rows))
                    texts.append(rows)
                    rowact += 1
                edlistaord.append(row[0])
                edlista[row[0]] = tuple(texts)
            for rows in edlistaord:
                editemselect.insert('', 'end', text=rows, values=edlista[rows])
            if tipo_mov not in [3, 4]:
                _totalcusto.configure(state='normal')
                _totalcusto.delete(0, 'end')
                if tipo_mov == 0:
                    _totalcusto.insert(0, num_brasil(str(somalc)))
                elif tipo_mov in [1, 2]:
                    _totalcusto.insert(0, num_brasil(str(somacusto)))
                    _vlmargem.configure(state='normal')
                    _vlmargem.delete(0, 'end')
                    _vlmargem.insert(0, num_brasil(str(somalc - somacusto)))
                    _vlmargem.configure(state='disabled')
                _totalcusto.configure(state='disabled')
            else:
                _totalcustos[0] = str(somacusto)
                _vltotal['text'] = 'Valor mercadorias: ' + num_brasil(str(round(somalc, 2)))
                _dinheiro.delete(0, 'end')
                _dinheiro.insert(0, num_brasil(str(round(somalc, 2))))
                _troco.delete(0, 'end')
                _troco.insert(0, '0,00')
                if tipo_mov == 4:
                    if _vencimentos:
                        _prazo['text'] = ''
                        _valor_parc = somalc / len(_vencimentos)
                        separa = 6
                        if len(_vencimentos) == 4:
                            separa = 2
                        elif len(_vencimentos) in [5, 6]:
                            separa = 3
                        a = 0
                        for row in _vencimentos:
                            _vencimentos[a]['Valor'] = _valor_parc
                            a += 1
                            if a >= 2:
                                if a != separa + 1:
                                    _prazo['text'] += ' - '
                            _prazo['text'] += row['Vencimento'].strftime('%d/%m/%y') + ' R$' + '{:03.2f}'.format(round(row['Valor'], 2)).replace('.', ',')
                            if separa == a:
                                _prazo['text'] += '\n'
            if tipo_mov == 0:
                try:
                    somalc = somalc - float(num_usa(_vldesconto.get()))
                except:
                    _vldesconto.delete(0, 'end')
                    _vldesconto.insert(0, '0,0')
                _vltotal.delete(0, 'end')
                _vltotal.insert(0, num_brasil(str(somalc)))
            if tipo_mov in [1, 2]:
                _totalvenda.configure(state='normal')
                _totalvenda.delete(0, 'end')
                _totalvenda.insert(0, num_brasil(str(somalc)))
                _totalvenda.configure(state='disabled')

        def updatethis():
            def valid():
                c.execute('SELECT Senha, Usuarios.Id, Nome, Ativo, Adm, Financ, Estoque, Vendas FROM Usuarios JOIN Usuarios_Acessos ON Usuarios_Acessos.Id = Usuarios.Id WHERE Nome = "' + _usuario.get().upper() + '"')
                #c.execute('SELECT Senha, Id FROM Usuarios WHERE Nome = "' + _usuario.get().upper() + '"')
                resp = c.fetchone()
                if resp:
                    if resp[0] == _senha.get() and resp[4]:
                        executar()
                    elif resp[0] == _senha.get() and resp[7] == 4:
                        executar()
                    elif resp[0] == _senha.get() and resp[7] < 4:
                        messagebox.showerror(title='Aviso', message='Operação não autorizada. Contate o administrador!')
                        update.focus()
                    else:
                        messagebox.showerror(title='Aviso', message='Senha digitada não confere!')
                        update.focus()
                    form_access.destroy()
                else:
                    if _usuario.get() == 'admin' and _senha.get() == 'root':
                        executar()
                    form_access.destroy()

            def executar():
                executethis = True
                if mode in ['new', 'edit', 'dup']:
                    mens = 'Tem certeza que deseja atualizar esse documento?'
                    erros = []
                    # Validar não vazios.
                    if tipo_mov == 0:
                        values = {'Nota Fiscal/Fatura': '', 'Emissão': '', 'Recebimento': '', 'Valor merc R$': '', 'Parceiro': '', 'Desconto': '', 'Valor cobrado': ''}
                        values['Nota Fiscal/Fatura'] = _doc.get()
                        values['Emissão'] = _datadoc.get()
                        values['Recebimento'] = _dataefetiva.get()
                        values['Valor merc R$'] = _totalcusto.get()
                        values['Parceiro'] = _parceiro.get()
                        values['Desconto'] = _vldesconto.get()
                        values['Valor cobrado'] = _vltotal.get()
                    if tipo_mov in [1, 2]:
                        values = {'Ordem Produção': '', 'Emissão': '', 'Data Entrega': '', 'Valor produzido R$': '', 'Margem': ''}
                        values['Ordem Produção'] = _doc.get()
                        values['Emissão'] = _datadoc.get()
                        values['Data Entrega'] = _dataefetiva.get()
                        values['Valor produzido R$'] = _totalcusto.get()
                        values['Margem'] = _vlmargem.get()
                    if tipo_mov in [3, 4]:
                        values = {'Cliente': '', 'Dinheiro': '', 'Troco': ''}
                        values['Cliente'] = _parceiro.get()
                        values['Dinheiro'] = _dinheiro.get()
                        values['Troco'] = _troco.get()
                    empty = []
                    for row in sorted(values.keys()):
                        if values[row] == '':
                            empty.append(row)
                    if empty:
                        if len(empty) == 1:
                            msgerro = '-O campo ' + empty[0] + ' não foi preenchido.'
                            erros.append(msgerro)
                        else:
                            msgerro = '-Os campos ' + empty[0]
                            for row in empty[1:-1]:
                                msgerro = msgerro + ', ' + row
                            msgerro = msgerro + ' e ' + empty[-1:][0] + ' não foram preenchidos.'
                            erros.append(msgerro)
                    # validar números errados
                    if tipo_mov in [0, 1, 2]:
                        valorfill1 = num_usa(_totalcusto.get())
                        if tipo_mov == 0:
                            valorfill2 = num_usa(_vldesconto.get())
                            valorfill3 = num_usa(_vltotal.get())
                        elif tipo_mov in [1, 2]:
                            valorfill2 = num_usa(_totalvenda.get())
                    elif tipo_mov in [3, 4]:
                        valorfill1 = num_usa(_dinheiro.get())
                        valorfill2 = num_usa(_troco.get())
                    try:
                        teste = float(valorfill1)
                        teste = float(valorfill2)
                        if tipo_mov == 0:
                            teste = float(valorfill3)
                        if float(valorfill1) <= 0.0:
                            if tipo_mov in [0, 1, 2]:
                                erros.append('-Não é possível cadastrar um documento sem mercadorias.')
                            elif tipo_mov in [3, 4]:
                                erros.append('-O cliente precisa efetuar o pagamento.')
                        if tipo_mov in [3, 4]:
                            if float(valorfill2) - float(valorfill1) == 0.0:
                                erros.append('-Não é possível cadastrar um documento sem mercadorias.')
                    except:
                        erros.append('-Um dos campos contendo valores ou quantidades foi preenchido incorretamente.')
                    # impedir cadastro de documentos sem produtos
                    if erros:
                        msg = 'Não é possível atualizar esse cadastro:\n\n' + str.join('\n', erros) + '\n\nRevise os dados e tente novamente.'
                        messagebox.showerror(title='Aviso', message=msg)
                        executethis = False
                else:
                    mens = 'O movimento de estoque dos produtos movimentados será estornado. Tem certeza que deseja remover esse documento?'
                if executethis:
                    if messagebox.askyesno(title="Atualização de sistema", message=mens):
                        if tipo_mov == 0 and mode != 'remove':
                            if _financeiro.get():
                                edit_cp()
                            else:
                                updatethis_efetivate()
                                exit()
                        elif tipo_mov in [3, 4]:
                            edit_cp(True)
                        else:
                            updatethis_efetivate()
                            exit()
                    else:
                        update.focus()
            permission = True
            if userauth:
                if tipo_mov in [0, 1, 2] and mode == 'remove' and userauth[5] < 4:
                    permission = False
                if tipo_mov in [3, 4] and mode == 'remove' and userauth[6] < 4:
                    permission = False
                if not permission:
                    config = {'title': 'AUTENTICAÇÃO',
                            'dimension': '315x215+400+200',
                            'color': 'orange'}
                    form_access = main.form(config)
                    uwg = Widgets(form_access, 'orange')
                    uwg.label('', 5, 0, 0)
                    uwg.label('', 5, 1, 0, rowspan=4, height=10)
                    uwg.label('', 5, 3, 0)
                    uwg.label('', 5, 6, 0)
                    _usuario = uwg.textbox('Usuário: ', 20, 1, 1)
                    _senha = uwg.textbox('Senha: ', 20, 2, 1, show='*')
                    uwg.button('Autenticar', valid, 20, 1, 4, 1, colspan=2)
                    #uwg.button('Sair', valid, 20, 1, 5, 1, colspan=2)
                    _usuario.focus()
                else:
                    executar()
            else:
                executar()

        def updatethis_efetivate():
            command = []
            if tipo_mov in [0]:
                if not _financeiro.get(): # Eliminar movimento no Contas à Pagar.
                    c.execute('SELECT MovimentoFin FROM Materiais_Documentos WHERE Id = ' + _id.get())
                    valorkill = c.fetchone()
                    if valorkill:
                        c.execute('UPDATE Materiais_Documentos SET MovimentoFin = 0 WHERE Id = ' + _id.get())
                        c.execute('DELETE FROM Diario WHERE Id = ' + str(valorkill[0]))
            if mode in ['new', 'dup']:
                command.append("INSERT INTO Materiais_Documentos VALUES (")
                if tipo_mov in [0, 1, 2]:
                    command.append(_id.get() + ", ")
                    command.append("'" + str(datetime.now()) + "', ")
                elif tipo_mov in [3, 4]:
                    command.append(_id + ", ")
                    command.append("'" + _datacadastro + "', ")
                command.append("'" + str(datetime.now()) + "', ")
                if tipo_mov != 4:
                    command.append(str(tipo_mov) + ", ") # Tipo 0 = Recebimento, 1 = Produção
                else:
                    command.append("3, ") # Tipo 0 = Recebimento, 1 = Produção
                command.append("0, ") # Doc Financeiro 0 - Construir formulário
                if tipo_mov in [0, 1, 2]:
                    command.append("'" + _doc.get() + "', ")
                    command.append("'" + _date(_datadoc.get()) + "', ")
                    command.append("'" + _date(_dataefetiva.get()) + "', ")
                    valorfill1 = _totalcusto.get().replace('.', '')
                    valorfill1 = valorfill1.replace(',', '.')
                    command.append(valorfill1 + ', ')
                elif tipo_mov in [3, 4]:
                    command.append("'" + _doc + "', ") # criar robot criação auto
                    command.append("'" + ActualDate + "', ")
                    command.append("'" + ActualDate + "', ")
                    command.append(_totalcustos[0] + ", ")
                if tipo_mov == 0:
                    command.append(valorfill1 + ', ') # Em Recebimentos, o valor venda é o mesmo valor custo.
                    c.execute('SELECT Id FROM Parceiros WHERE Nome = "' + _parceiro.get() + '"')
                    command.append(str(c.fetchone()[0]) + ", ")
                    command.append(str(int(_financeiro.get())) + ", ")
                    valorfill2 = _vldesconto.get().replace('.', '')
                    valorfill2 = valorfill2.replace(',', '.')
                    valorfill3 = _vltotal.get().replace('.', '')
                    valorfill3 = valorfill3.replace(',', '.')
                    command.append(valorfill2 + ', ')
                    command.append(valorfill3 + ', ')
                if tipo_mov in [1, 2]:
                    valorfill2 = _totalvenda.get().replace('.', '')
                    valorfill2 = valorfill2.replace(',', '.')
                    command.append(valorfill2 + ', ') # Em Recebimentos, o total venda é o mesmo que vl total.
                    command.append("0, ")
                    command.append("0, ")
                    command.append('0.0, ')
                    command.append(valorfill1 + ', ')
                if tipo_mov in [3, 4]:
                    valorfill1 = _vltotal['text'].split(' ')
                    valorfill2 = valorfill1[2].replace('.', '')
                    valorfill2 = valorfill2.replace(',', '.')
                    command.append(valorfill2 + ', ') # Em Faturamento, o total venda é o mesmo que vl total.
                    c.execute('SELECT Id FROM Parceiros WHERE Nome = "' + _parceiro.get() + '"')
                    command.append(str(c.fetchone()[0]) + ", ")
                    command.append("1, ")
                    command.append('0.0, ')
                    command.append(valorfill2 + ', ')
                    if tipo_mov == 3:
                        command.append('"Venda de balcão", ')
                    elif tipo_mov == 4:
                        command.append("'" + _obs.get() + "', ")
                    else:
                        command.append("'" + _descricao.get() + "', ")
                if tipo_mov in [0, 1, 2]:
                    command.append("'" + _descricao.get() + "', ")
                _user = '0'
                if userauth:
                    _user = str(userauth[0])
                command.append(_user) # Em Recebimentos e Produção sempre 0. Não há entrada como perda. ALTERAR 2 incluindo Perda
                command.append(")")
                c.execute(str.join('', command))
            if mode == 'edit':
                command.append("UPDATE Materiais_Documentos SET ")
                command.append("DataAtualiza = '" + str(datetime.now()) + "', ")
                command.append("Doc = '" + _doc.get() + "', ")
                command.append("DataDoc = '" + _date(_datadoc.get()) + "', ")
                command.append("DataEfetiva = '" + _date(_dataefetiva.get()) + "', ")
                valorfill1 = _totalcusto.get().replace('.', '')
                valorfill1 = valorfill1.replace(',', '.')
                command.append("TotalCusto = " + valorfill1 + ", ")
                if tipo_mov == 0:
                    command.append("TotalVenda = " + valorfill1 + ", ")
                    c.execute('SELECT Id FROM Parceiros WHERE Nome = "' + _parceiro.get() + '"')
                    command.append("Parceiro = " + str(c.fetchone()[0]) + ", ")
                    command.append("Financeiro = " + str(int(_financeiro.get())) + ", ")
                    valorfill2 = _vldesconto.get().replace('.', '')
                    valorfill2 = valorfill2.replace(',', '.')
                    valorfill3 = _vltotal.get().replace('.', '')
                    valorfill3 = valorfill3.replace(',', '.')
                    command.append("VlDesconto = " + valorfill2 + ", ")
                    command.append("VlTotal = " + valorfill3 + ", ")
                if tipo_mov in [1, 2]:
                    valorfill2 = _totalvenda.get().replace('.', '')
                    valorfill2 = valorfill2.replace(',', '.')
                    command.append("TotalVenda = " + valorfill2 + ", ")
                    command.append("Parceiro = 0, ")
                    command.append("Financeiro = 0, ")
                    command.append("VlDesconto = 0.0, ")
                    command.append("VlTotal = " + valorfill1 + ", ")
                command.append("Descricao = '" + _descricao.get() + "' ")
                command.append("WHERE Id = " + _id.get())
                c.execute(str.join('', command))
            if mode == 'remove':
                if tipo_mov in [0, 1, 2]:
                    c.execute('SELECT MovimentoFin FROM Materiais_Documentos WHERE Id = ' + _id.get())
                elif tipo_mov in [3, 4]:
                    c.execute('SELECT MovimentoFin FROM Materiais_Documentos WHERE Id = ' + _id)
                valorkill = c.fetchone()[0]
                if valorkill:
                    if tipo_mov != 4:
                        c.execute('DELETE FROM Diario WHERE Id = ' + str(valorkill))
                    elif tipo_mov == 4:
                        for row in range(len(_vencimentos)):
                            c.execute('DELETE FROM Diario WHERE Id = ' + str(valorkill + row))
                if tipo_mov in [0, 1, 2]:
                    c.execute('DELETE FROM Materiais_Documentos WHERE Id = ' + _id.get())
                elif tipo_mov in [3, 4]:
                    c.execute('DELETE FROM Materiais_Documentos WHERE Id = ' + _id)
            command = []
            if tipo_mov in [0, 1, 2]:
                command.append('DELETE FROM Materiais_Movimentos WHERE Documento = ' + _id.get())
            elif tipo_mov in [3, 4]:
                command.append('DELETE FROM Materiais_Movimentos WHERE Documento = ' + _id)
            c.execute(str.join('', command))
            if mode != 'remove':
                for row in itensdata:
                    command = []
                    command.append("INSERT INTO Materiais_Movimentos VALUES (")
                    command.append(str(row[0]) + ", ")
                    command.append("'" + str(datetime.now())[0:19] + "', ")
                    command.append("'" + str(datetime.now())[0:19] + "', ")
                    if tipo_mov in [0, 1, 2]:
                        command.append("'" + _date(_dataefetiva.get()) + "', ")
                    elif tipo_mov in [3, 4]:
                        command.append("'" + ActualDate + "', ")
                    command.append(str(row[1]) + ", ")
                    command.append(num_usa(str(row[3])) + ", ")
                    if tipo_mov in [2, 3, 4]:
                        command.append(num_usa(str('-' + row[4])) + ", ")
                    else:
                        command.append(num_usa(str(row[4])) + ", ")
                    command.append(num_usa(str(row[5])) + ", ")
                    command.append(num_usa(str(row[5])) + ", ")
                    if tipo_mov in [0, 1, 2]:
                        command.append(_id.get())
                    elif tipo_mov in [3, 4]:
                        command.append(_id)
                    command.append(")")
                    c.execute(str.join('', command))
                    if tipo_mov in [0, 1]:
                        c.execute('SELECT Margem, ValorVenda, ValorCusto FROM Materiais_Itens WHERE Id = ' + str(row[1]))
                        value = c.fetchone()
                        if tipo_mov == 0:
                            if value[1] <= float(num_usa(str(row[3]))):
                                newmarg = 0.0
                                newvend = round(float(num_usa(str(row[3]))), 2)
                            else:
                                newmarg = round(value[1] - round(float(num_usa(str(row[3]))), 2), 2)
                                newvend = value[1]
                        elif tipo_mov == 1:
                            if value[2] >= float(num_usa(str(row[3]))):
                                newmarg = 0.0
                                newvend = round(float(num_usa(str(row[3]))), 2)
                            else:
                                newmarg = round(round(float(num_usa(str(row[3]))), 2) - value[2], 2)
                                newvend = value[2]
                        command = []
                        command.append("UPDATE Materiais_Itens SET ")
                        if tipo_mov == 0:
                            command.append("ValorCusto = " + num_usa(str(row[3])) + ", ")
                            command.append("Margem = " + str(newmarg) + ", ")
                            command.append("ValorVenda = " + str(newvend) + " ")
                        elif tipo_mov == 1:
                            command.append("ValorCusto = " + str(newvend) + ", ")
                            command.append("Margem = " + str(newmarg) + ", ")
                            command.append("ValorVenda = " + num_usa(str(row[3])) + " ")
                        command.append("WHERE Id = " + str(row[1]))
                        c.execute(str.join('', command))
            if tipo_mov == 0: # rotina pra integrar Contas à Pagar
                # Atualizar: banco, data_venc, data_pag, tipo_mov
                pass
            conn.commit()
            seek()

        def _parceiro_cmd(resp=None):
            if _parceiro.get():
                if classe_parceiro[0] == 'Cliente':
                    c.execute('SELECT Nome FROM Parceiros WHERE Modo IN (0, 1) ORDER BY Nome')
                    rows = c.fetchall()
                    for row in rows:
                        if not row[0] in lista1:
                            lista1.append(row[0])
                    _parceiro['values'] = lista1
                elif classe_parceiro[0] == 'Fornecedor':
                    c.execute('SELECT Nome FROM Parceiros WHERE Modo IN (0, 2) ORDER BY Nome')
                    rows = c.fetchall()
                    for row in rows:
                        if not row[0] in lista1:
                            lista1.append(row[0])
                    _parceiro['values'] = lista1
                widgets.combobox_return(_parceiro, lista1)
                value = ['']
                dadosparc = 'Dados do cliente\n\n\n'
                command = ('SELECT Nome, NomeCompleto, Tipo, Doc, Endereco, Telefone FROM Parceiros WHERE Nome = "' + _parceiro.get() + '"')
                c.execute(command)
                try:
                    value = c.fetchall()[0]
                    td = ['CPF', 'CNPJ']
                    dadosparc = 'Cliente: ' + value[1] + '\n' + td[value[2]] + ': ' + value[3] + '\nEndereço: ' + value[4] + '\nTelefone: ' + value[5]
                except:
                    if userauth:
                        if userauth[6] < 4:
                            messagebox.showinfo('Informação', 'Não foi encontrado parceiro nessa pesquisa. Caso precise incluir, procure o administrador do sistema.')
                            update.focus()
                        else:
                            if messagebox.askyesno('Informação', 'Não foi encontrado parceiro nessa pesquisa. Deseja incluir agora?'):
                                update.focus()
                                parceiros()
                    else:
                        if messagebox.askyesno('Informação', 'Não foi encontrado parceiro nessa pesquisa. Deseja incluir agora?'):
                            update.focus()
                            parceiros()
                if tipo_mov in [3, 4]:
                    _parceirodata['text'] = dadosparc

        def _datadoc_cmd(Value=None):
            data_cmd(_datadoc)

        def _dataefetiva_cmd(Value=None):
            data_cmd(_dataefetiva)

        def _totalcusto_cmd(Value=None):
            value = num_usa(_totalcusto.get())
            try:
                float(value)
            except:
                value = '0.00'
            _totalcusto.delete(0, 'end')
            _totalcusto.insert(0, num_brasil(value))

        def _obs_cmd(Value=None):
            _prazo['text'] = ''
            var1 = _obs.get()
            ini = 0
            fim = 0
            for row in range(len(var1)):
                if var1[0 + row:6 + row].lower() == 'prazo(':
                    ini = 6 + row
                elif var1[row] == ')':
                    fim = row
            if ini and fim:
                a = _vltotal['text'].split(' ')
                a = float(num_usa(a[-1:][0]))
                var2 = var1[ini:fim].split('/')
                a = a / len(var2)
                _vencimentos.clear()
                for row in var2:
                    _vencimentos.append(
                        {'Vencimento': datetime.now() + timedelta(days=int(row)), 'Valor': a}
                    )
                separa = 6
                if len(_vencimentos) == 4:
                    separa = 2
                elif len(_vencimentos) in [5, 6]:
                    separa = 3
                a = 0
                for row in _vencimentos:
                    a += 1
                    if a >= 2:
                        if a != separa + 1:
                            _prazo['text'] += ' - '
                    _prazo['text'] += row['Vencimento'].strftime('%d/%m/%y') + ' R$' + '{:03.2f}'.format(round(row['Valor'], 2)).replace('.', ',')
                    if separa == a:
                        _prazo['text'] += '\n'

        if tipo_mov in [0, 1, 2]:
            dimension = widgets.geometry(500, 750)
            if tipo_mov == 0:
                title = 'Edição de faturas e/ou Notas Fiscais de Entrada'
            elif tipo_mov == 1:
                title = 'Ordens de Produção'
            elif tipo_mov == 2:
                title = 'Consumo de materiais'
            else:
                title = 'Transferência'
            config = {'title': title,
                    'dimension': dimension,
                    'color': 'pale goldenrod'}
            form_edit = main.form(config)
            wgedit = Widgets(form_edit, 'pale goldenrod')
            wgedit.label('', 5, 0, 0)
            wgedit.label('', 5, 7, 0)
            wgedit.label('', 5, 9, 0)
            wgedit.label('', 5, 13, 0)
            if mode in ['edit', 'remove', 'dup']:
                if itemselect.selection():
                    for i in itemselect.selection():
                        gridselected = str(itemselect.item(i, 'text'))
                else:
                    messagebox.showerror(title='Aviso', message='Não foi selecionado nenhum documento!')
                    exit()
                if tipo_mov == 0:
                    command = ('SELECT Materiais_Documentos.Id, Materiais_Documentos.DataCadastro, Materiais_Documentos.DataAtualiza, ' +
                            'Materiais_Documentos.Tipo, MovimentoFin, Materiais_Documentos.Doc, DataDoc, DataEfetiva, TotalCusto, TotalVenda, Parceiros.Nome, ' +
                            'Financeiro, VlDesconto, VlTotal, Descricao, Usuario FROM Materiais_Documentos ' +
                            'JOIN Parceiros ON Parceiros.Id = Parceiro ' +
                            'WHERE Materiais_Documentos.Id = ' + gridselected)
                if tipo_mov in [1, 2]:
                    command = ('SELECT Materiais_Documentos.Id, Materiais_Documentos.DataCadastro, Materiais_Documentos.DataAtualiza, ' +
                            'Materiais_Documentos.Tipo, MovimentoFin, Materiais_Documentos.Doc, DataDoc, DataEfetiva, TotalCusto, TotalVenda, Parceiro, ' +
                            'Financeiro, VlDesconto, VlTotal, Descricao, Usuario FROM Materiais_Documentos ' +
                            'WHERE Materiais_Documentos.Id = ' + gridselected)
                c.execute(command)
                doc = c.fetchall()
            if tipo_mov == 0:
                _id = wgedit.textbox('Id ', 5, 1, 1)
                _datacadastro = wgedit.textbox('Data da criação ', 17, 2, 1)
                _dataatualiza = wgedit.textbox('Data última edição ', 17, 3, 1)
                _tipo = None
                _movimentofin = None
                _doc = wgedit.textbox('Nota Fiscal/Fatura ', 12, 4, 1)
                _datadoc = wgedit.textbox('Emissão ', 12, 5, 1, cmd=_datadoc_cmd)
                _dataefetiva = wgedit.textbox('Recebimento ', 12, 6, 1, cmd=_dataefetiva_cmd)
                _totalcusto = wgedit.textbox('Valor merc R$ ', 8, 1, 3)
                _totalvenda = None
                c.execute('SELECT Nome FROM Parceiros WHERE Modo IN (0, 2) ORDER BY Nome')
                lista1 = []
                for row in c.fetchall():
                    lista1.append(row[0])
                _parceiro = wgedit.combobox('Fornecedor ', 20, lista1, 2, 3, cmd=_parceiro_cmd, seek=_parceiro_cmd)
                if mode in ['edit', 'remove', 'dup']:
                    if doc[0][11]:
                        _financeiro = wgedit.check('Gera financeiro? ', 15, ' Gerar                        ', 3, 3, True)
                    else:
                        _financeiro = wgedit.check('Gera financeiro? ', 15, ' Gerar                        ', 3, 3, False)
                else:
                    _financeiro = wgedit.check('Gera financeiro? ', 15, ' Gerar                        ', 3, 3, True)
                _vldesconto = wgedit.textbox('Desconto R$ ', 8, 4, 3, cmd=edit_desconto)
                _vltotal = wgedit.textbox('Valor cobrado R$ ', 8, 5, 3, cmd=edit_total)
                _descricao = wgedit.textbox('Descrição ', 30, 6, 3)
                _perda = None
            elif tipo_mov == 1:
                _id = wgedit.textbox('Id ', 5, 1, 1)
                _datacadastro = wgedit.textbox('Data da criação ', 17, 2, 1)
                _dataatualiza = wgedit.textbox('Data última edição ', 17, 3, 1)
                _tipo = None
                _movimentofin = None
                _datadoc = wgedit.textbox('Emissão ', 12, 4, 1, cmd=_datadoc_cmd)
                _dataefetiva = wgedit.textbox('Data de Entrega ', 12, 5, 1, cmd=_dataefetiva_cmd)
                _doc = wgedit.textbox('Ordem Produção ', 12, 1, 3)
                _totalcusto = wgedit.textbox('Valor Custo R$ ', 8, 2, 3)
                _totalvenda = wgedit.textbox('Valor Venda R$ ', 8, 3, 3)
                _financeiro = None
                _vlmargem = wgedit.textbox('Margem Lucro R$ ', 8, 4, 3) # campo calculado
                _vldesconto = None
                _vltotal = None  # = Total Venda
                _descricao = wgedit.textbox('Descrição ', 30, 5, 3)
                _perda = None
            elif tipo_mov == 2:
                _id = wgedit.textbox('Id ', 5, 1, 1)
                _datacadastro = wgedit.textbox('Data da criação ', 17, 2, 1)
                _dataatualiza = wgedit.textbox('Data última edição ', 17, 3, 1)
                _tipo = None
                _movimentofin = None
                _datadoc = wgedit.textbox('Emissão ', 12, 4, 1, cmd=_datadoc_cmd)
                _dataefetiva = wgedit.textbox('Data de Entrega ', 12, 5, 1, cmd=_dataefetiva_cmd)
                _doc = wgedit.textbox('Ordem Produção ', 12, 1, 3)
                _totalcusto = wgedit.textbox('Valor Custo R$ ', 8, 2, 3)
                _totalvenda = wgedit.textbox('Valor Venda R$ ', 8, 3, 3)
                _financeiro = None
                _vlmargem = wgedit.textbox('Margem Lucro R$ ', 8, 4, 3) # campo calculado
                _vldesconto = None
                _vltotal = None  # = Total Venda
                _descricao = wgedit.textbox('Descrição ', 30, 5, 3)
                _perda = None
            incluir = wgedit.button('Incluir produto', edit_new, 15, 0, 8, 1, 2)
            editar = wgedit.button('Editar produto', edit_exist, 15, 0, 8, 3, 1)
            excluir = wgedit.button('Excluir produto', edit_remove, 15, 0, 8, 4, 1)

            # grid setup ini
            if tipo_mov in [0, 1, 2]:
                edcolsf = ['Produto', 'Descricao', 'VlUnit', 'Qtde', 'VlCusto']
                edheadf = {
                    'Produto': {'text': 'Código', 'width': 70},
                    'Descricao': {'text': 'Produto', 'width': 240},
                    'VlUnit': {'text': 'Valor unitário', 'width': 100, 'anchor': 'e', 'format': 'float'},
                    'Qtde': {'text': 'Quantidade', 'width': 90, 'anchor': 'e', 'format': 'float'},
                    'VlCusto': {'text': 'Valor total', 'width': 100, 'anchor': 'e', 'format': 'float'},
                }
            edlista = {}
            edlistaord = []
            itensdata = []
            editemselect = wgedit.grid(edcolsf, edheadf, edlista, edlistaord, 10, 10, 1, colspan=6)
            _it_id = [None]
            c.execute('SELECT Id FROM Materiais_Movimentos ORDER BY Id')
            rows = c.fetchall()
            NewId = 1
            if rows:
                Ids = []
                for row in rows:
                    Ids.append(row[0])
                Ids.sort() 
                NewId = Ids[-1:][0] + 1
            _it_id[0] = str(NewId)
            # grid setup fim (Executar a primeira edição foi movida para o final do bloco)

            if mode in ['edit', 'remove', 'dup']:
                dup_id = str(doc[0][0]) # para passbuscar os itens do documento duplicado
                if mode not in ['dup']:
                    _id.insert(0, str(doc[0][0]))
                    _datacadastro.insert(0, datahora_brasil(str(doc[0][1])))
                    _dataatualiza.insert(0, datahora_brasil(str(doc[0][2])))
                    _datadoc.insert(0, data_brasil(str(doc[0][6])))
                    _dataefetiva.insert(0, data_brasil(str(doc[0][7])))
                if tipo_mov == 0:
                    _parceiro.insert(0, str(doc[0][10]))
                _doc.insert(0, str(doc[0][5]))
                _totalcusto.insert(0, num_brasil(str(abs(doc[0][8]))))
                if tipo_mov in [1, 2]:
                    _totalvenda.insert(0, num_brasil(str(abs(doc[0][9]))))
                if tipo_mov == 0:
                    _vldesconto.delete(0, 'end')
                    _vldesconto.insert(0, num_brasil(str(abs(doc[0][12]))))
                    _vltotal.delete(0, 'end')
                    _vltotal.insert(0, num_brasil(str(abs(doc[0][13]))))
                _descricao.insert(0, str(doc[0][14]))
            if mode in ['new', 'dup']:
                c.execute('SELECT Id FROM Materiais_Documentos ORDER BY Id')
                rows = c.fetchall()
                NewId = 1
                if rows:
                    Ids = []
                    for row in rows:
                        Ids.append(row[0])
                    Ids.sort() 
                    NewId = Ids[-1:][0] + 1
                _id.insert(0, str(NewId))
                _datacadastro.insert(0, datahora_brasil(str(datetime.now())))
                _dataatualiza.insert(0, datahora_brasil(str(datetime.now())))
                _datadoc.insert(0, data_brasil(ActualDate))
                _dataefetiva.insert(0, data_brasil(ActualDate))
                if tipo_mov == 0:
                    if mode == 'new':
                        _vldesconto.insert(0, '0,00')
                    _vltotal.insert(0, '0,00')                

            _id.configure(state='disabled')
            _datacadastro.configure(state='disabled')
            _dataatualiza.configure(state='disabled')
            _totalcusto.configure(state='disabled')
            if tipo_mov in [1, 2]:
                _totalvenda.configure(state='disabled')
                _vlmargem.configure(state='disabled')
            if mode in ['edit', 'new', 'dup']:
                if tipo_mov == 0:
                    _doc.focus()
                elif tipo_mov in [1, 2]:
                    _datadoc.focus()
                update = wgedit.button('Atualizar', updatethis, 20, 0, 15, 1, 2)
                cancel = wgedit.button('Cancelar', edit_cancel, 20, 0, 15, 4, 1)
            if mode in ['remove']:
                incluir.configure(state='disabled')
                editar.configure(state='disabled')
                excluir.configure(state='disabled')
                _datadoc.configure(state='disabled')
                if tipo_mov == 0:
                    _parceiro.configure(state='disabled')
                _doc.configure(state='disabled')
                _descricao.configure(state='disabled')
                if tipo_mov == 0:
                    _vldesconto.configure(state='disabled')
                    _vltotal.configure(state='disabled')
                _dataefetiva.configure(state='disabled')
                wgedit.label('Registro que será excluído!', 0, 15, 1, 4)
                update = wgedit.button('Confirmar exclusão', updatethis, 20, 0, 16, 1, 4)
            updateitens(itensdata, True)
        elif tipo_mov in [3, 4]:
            _vencimentos = []
            if mode not in ['edit']:
                dimension = widgets.geometry(500, 750)
                title = 'Venda de mercadorias e serviços'
                config = {'title': title,
                        'dimension': dimension,
                        'color': 'PaleGreen'}
                form_edit = main.form(config)
                wgedit = Widgets(form_edit, 'PaleGreen')
                wgedit.label('', 5, 0, 0)
                wgedit.label('', 5, 7, 0)
                wgedit.label('', 5, 9, 0)
                wgedit.label('', 5, 11, 0)
                wgedit.label('', 5, 15, 0)
                if mode in ['edit', 'remove', 'dup']:
                    if itemselect.selection():
                        for i in itemselect.selection():
                            gridselected = str(itemselect.item(i, 'text'))
                    else:
                        messagebox.showerror(title='Aviso', message='Não foi selecionado nenhum documento!')
                        exit()
                    command = ('SELECT Materiais_Documentos.Id, Materiais_Documentos.DataCadastro, Materiais_Documentos.DataAtualiza, ' +
                            'Materiais_Documentos.Tipo, MovimentoFin, Materiais_Documentos.Doc, DataDoc, DataEfetiva, TotalCusto, TotalVenda, Parceiros.Nome, ' +
                            'Financeiro, VlDesconto, VlTotal, Descricao, Usuario FROM Materiais_Documentos ' +
                            'JOIN Parceiros ON Parceiros.Id = Parceiro ' +
                            'WHERE Materiais_Documentos.Id = ' + gridselected)
                    c.execute(command)
                    doc = c.fetchall()
                _id = None
                _totalcustos = ['0.0']
                c.execute('SELECT Nome FROM Parceiros WHERE Modo IN (0, 1) ORDER BY Nome')
                lista1 = []
                for row in c.fetchall():
                    lista1.append(row[0])
                if userauth:
                    c.execute('SELECT Nome FROM Usuarios WHERE Id = ' + str(userauth[0]))
                    row = c.fetchone()[0]
                    wgedit.label('Vendedor: ' + row.upper(), 25, 3, 1, 2)
                wgedit.label('Cliente', 20, 1, 1, 2)
                wgedit.label('', 1, 3, 1)
                _parceiro = wgedit.combobox('', 20, lista1, 2, 1, cmd=_parceiro_cmd, seek=_parceiro_cmd)
                _parceirodata = wgedit.label('Dados do cliente\n\n\n', 50, 1, 3, 3, 3)
                _parceirodata['bg'] = 'white'
                _parceirodata['relief'] = 'sunken'
                if tipo_mov == 4:
                    _obs = wgedit.textbox('', 20, 4, 1, cmd=_obs_cmd)
                    _obs.insert(0, 'Prazo(0)')
                    _prazo = wgedit.label('', 50, 4, 3, 3)
                incluir = wgedit.button('Adicionar item', edit_new, 15, 0, 8, 1, 2)
                excluir = wgedit.button('Remover da lista', edit_remove, 15, 0, 8, 5, 1)
                edcolsf = ['Produto', 'Descricao', 'VlUnit', 'Qtde', 'VlCusto']
                edheadf = {
                    'Produto': {'text': 'Código', 'width': 70},
                    'Descricao': {'text': 'Produto', 'width': 240},
                    'VlUnit': {'text': 'Valor unitário', 'width': 100, 'anchor': 'e', 'format': 'float'},
                    'Qtde': {'text': 'Quantidade', 'width': 90, 'anchor': 'e', 'format': 'float'},
                    'VlCusto': {'text': 'Valor total', 'width': 100, 'anchor': 'e', 'format': 'float'},
                }
                edlista = {}
                edlistaord = []
                itensdata = []
                editemselect = wgedit.grid(edcolsf, edheadf, edlista, edlistaord, 8, 10, 1, colspan=6)
                _it_id = [None]
                c.execute('SELECT Id FROM Materiais_Movimentos ORDER BY Id')
                rows = c.fetchall()
                NewId = 1
                if rows:
                    Ids = []
                    for row in rows:
                        Ids.append(row[0])
                    Ids.sort() 
                    NewId = Ids[-1:][0] + 1
                _it_id[0] = str(NewId)

                _vltotal = wgedit.label('Valor mercadorias: 0,00', 30, 12, 2, 2, 2)
                _vltotal['bg'] = 'yellow2'
                _vltotal['fg'] = 'red'
                _vltotal['font'] = 'Arial 14 bold'
                _vltotal['relief'] = 'sunken'
                #_vltotal = wgedit.textbox('Valor mercadorias ', 8, 12, 2, '0,00')
                _dinheiro = wgedit.textbox('Dinheiro ', 8, 12, 4, '0,00', cmd=edit_total)
                _troco = wgedit.textbox('Troco ', 8, 13, 4, '0,00', cmd=edit_desconto)
                _parceiro.focus()
                if mode in ['edit', 'remove', 'dup']:
                    dup_id = str(doc[0][0]) # para buscar os itens do documento duplicado
                    if mode not in ['dup']:
                        _id = str(doc[0][0])
                    _parceiro.insert(0, str(doc[0][10]))
                    _parceiro_cmd()
                    if tipo_mov == 4:
                        _obs.delete(0, 'end')
                        _obs.insert(0, doc[0][14])
                        _obs_cmd()
                    _vltotal['text'] = 'Valor mercadorias: ' + num_brasil(str(abs(doc[0][9])))
                    _dinheiro.delete(0, 'end')
                    _dinheiro.insert(0, num_brasil(str(abs(doc[0][9]))))
                    _troco.delete(0, 'end')
                    _troco.insert(0, '0,00')
                if mode in ['edit', 'new', 'dup']:
                    update = wgedit.button('Concluir compra', updatethis, 20, 0, 16, 1, 2)
                    cancel = wgedit.button('Cancelar', edit_cancel, 20, 0, 16, 5, 1)
                if mode in ['new', 'dup']:
                    c.execute('SELECT MAX(Id) FROM Materiais_Documentos ORDER BY Id')
                    rows = c.fetchone()
                    NewId = 1
                    if rows[0]:
                        NewId = rows[0] + 1
                    _id = str(NewId)
                    c.execute('SELECT MAX(Doc) FROM Materiais_Documentos WHERE Tipo = 3 ORDER BY Doc')
                    rows = c.fetchone()
                    NewId = 'CP000001'
                    if rows[0]:
                        LastId = int(rows[0][-6:]) + 1
                        NewId = 'CP' + str(LastId).zfill(6)
                    _doc = str(NewId)
                    _datacadastro = str(datetime.now())
                if mode in ['remove']:
                    incluir.configure(state='disabled')
                    excluir.configure(state='disabled')
                    _parceiro.configure(state='disabled')
                    if tipo_mov == 4:
                        _obs.configure(state='disabled')
                    _dinheiro.configure(state='disabled')
                    _troco.configure(state='disabled')
                    wgedit.label('Dados do Cupom de compra.', 0, 15, 1, 4)
                    update = wgedit.button('Cancelar compra', updatethis, 20, 0, 16, 1, 4)
                updateitens(itensdata, True)
            else:
                if itemselect.selection():
                    for i in itemselect.selection():
                        gridselected = str(itemselect.item(i, 'text'))
                    dimension = widgets.geometry(185, 380)
                    title = 'Venda em balcão - Alterar data da compra'
                    config = {'title': title,
                            'dimension': dimension,
                            'color': 'PaleGreen'}
                    form_edit = main.form(config)
                    wgedit = Widgets(form_edit, 'PaleGreen')
                    wgedit.label('\n\n\n\n\n\n', 3, 0, 0, rowspan=5)
                    wgedit.label('', 3, 5, 0)
                    wgedit.label('É permitido ajustar a data de compra do cupom.', 39, 1, 1, 2)
                    wgedit.label('Cuidado! Mudanças afetam o fluxo de caixa.', 39, 2, 1, 2)
                    command = ('SELECT DataDoc, DataEfetiva FROM Materiais_Documentos ' +
                            'WHERE Materiais_Documentos.Id = ' + gridselected)
                    c.execute(command)
                    doc = c.fetchone()
                    _dataefetiva = wgedit.textbox('Data atual: ', 12, 3, 1, data_brasil(doc[0]))
                    _datanova = wgedit.textbox('Nova data: ', 12, 4, 1, data_brasil(doc[0]))
                    _confirmar = wgedit.button('Ok', edit_exist, 12, 1, 6, 1, 2)
                    _dataefetiva.configure(state='disabled')
                    _datanova.focus()
                else:
                    messagebox.showerror(title='Aviso', message='Não foi selecionado nenhum documento!')
                    exit()

    def generate_sql(value=False):
        def _fc(first):
            if first:
                resp = 'WHERE '
            else:
                resp = ' AND '
            return resp

        first = True
        command = []
        if value:
            command.append('SELECT SUM(VlTotal) ')
        else:
            if tipo_mov in [0, 3, 4]:
                command.append('SELECT Materiais_Documentos.Id, Materiais_Documentos.Doc, Parceiros.Nome, DataDoc, DataEfetiva, VlTotal, Descricao ')
                command.append('FROM Materiais_Documentos JOIN Parceiros ON Parceiros.Id = Parceiro ')
            elif tipo_mov in [1, 2]:
                command.append('SELECT Materiais_Documentos.Id, Materiais_Documentos.Doc, DataDoc, DataEfetiva, TotalCusto, TotalVenda, Descricao ')
                command.append('FROM Materiais_Documentos ')
        if parceiro_sel.get():
            command.append(_fc(first) + 'Parceiros.Nome = "' + parceiro_sel.get() + '"')
            first = False
        if documento_in.get():
            command.append(_fc(first) + 'Materiais_Documentos.Doc = "' + documento_in.get() + '"')
            first = False
        if datadocini.get():
            datafill = _date(datadocini.get())
            command.append(_fc(first) + 'DataDoc >= "' + datafill + '"')
            first = False
        if datadocfim.get():
            datafill = _date(datadocfim.get())
            command.append(_fc(first) + 'DataDoc <= "' + datafill + '"')
            first = False
        if dataefetivaini.get():
            datafill = _date(dataefetivaini.get())
            command.append(_fc(first) + 'DataEfetiva >= "' + datafill + '"')
            first = False
        if dataefetivafim.get():
            datafill = _date(dataefetivafim.get())
            command.append(_fc(first) + 'DataEfetiva <= "' + datafill + '"')
            first = False
        if tipo_mov != 4:
            command.append(_fc(first) + 'Materiais_Documentos.Tipo = ' + str(tipo_mov))
        else:
            command.append(_fc(first) + 'Materiais_Documentos.Tipo = 3')
        first = False
        command.append(' ORDER BY DataEfetiva')
        if value:
            c.execute(str.join('', command))
            resp = c.fetchall()[0][0]
            if not resp:
                resp = 0.0
            return resp
        else:
            return str.join('', command)

    def seek():
        itemselect.delete(*itemselect.get_children())
        soma['text'] = 'Valor Total: 0,00'
        c.execute(generate_sql())
        doc = c.fetchall()
        combolist = {}
        ordlista = []
        somalc = 0.0
        if tipo_mov in [0, 3, 4]:
            for row in doc:
                texts = []
                rowact = 0
                for rows in row[1:]:
                    if rowact in [2, 3]:
                        texts.append(data_brasil(rows))
                    elif rowact in [4]:
                        somalc += rows
                        texts.append(num_brasil(str(abs(rows))))
                    else:
                        texts.append(rows)
                    rowact += 1
                ordlista.append(row[0])
                combolist[row[0]] = tuple(texts)
        elif tipo_mov in [1, 2]:
            for row in doc:
                texts = []
                rowact = 0
                for rows in row[1:]:
                    if rowact in [1, 2]:
                        texts.append(data_brasil(rows))
                    elif rowact in [3, 4]:
                        if rowact == 4:
                            somalc += rows
                        texts.append(num_brasil(str(abs(rows))))
                    else:
                        texts.append(rows)
                    rowact += 1
                ordlista.append(row[0])
                combolist[row[0]] = tuple(texts)
        for rows in ordlista:
            itemselect.insert('', 'end', text=rows, values=combolist[rows])
        soma['text'] = 'Valor Total: ' + num_brasil(format(abs(somalc), '.2f'))

    def filecsv():
        file_output = "csv/detalhe_movimentos.csv"
        # soma['text'] = 'Valor Total: 0,00'
        c.execute(generate_sql())
        doc = c.fetchall()
        combolist = {}
        ordlista = []
        somalc = 0.0
        for row in doc:
            texts = {}
            rowact = 0
            for rows in row[1:]:
                if rowact in [1] and not in_value and not out_value:
                    texts[colsf[rowact]] = trftype[rows]
                elif rowact in [0, 4]:
                    texts[colsf[rowact]] = data_brasil(rows)
                elif rowact in [5]:
                    somalc += rows
                    if texts[colsf[1]] == 'Saída' and not in_value and not out_value:
                        texts[colsf[rowact]] = '-' + num_brasil(str(abs(rows)))
                    else:
                        texts[colsf[rowact]] = num_brasil(str(abs(rows)))
                else:
                    texts[colsf[rowact]] = rows
                rowact += 1
            ordlista.append(row[0])
            combolist[row[0]] = texts
        with open(file_output, 'w') as csvfile:
            fieldnames = colsf
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames, delimiter=';')
            writer.writeheader()
            for row in combolist:
                #rowact = 0
                writer.writerow(combolist[row])

        '''
        for rows in ordlista:
            itemselect.insert('', 'end', text=rows, values=combolist[rows])
        if not in_value and not out_value:
            soma['text'] = 'Valor Total: ' + num_brasil(format(abs(somalc), '.2f'))
        else:
            soma['text'] = 'Valor Total: ' + num_brasil(format(abs(generate_sql(value=True)), '.2f'))
        '''
    
    def parceiro_cmd(value=None):
        widgets.combobox_return(parceiro_sel, lista1)

    def datadocini_cmd(Value=None):
        data_cmd(datadocini)

    def datadocfim_cmd(Value=None):
        data_cmd(datadocfim)

    def dataefetivaini_cmd(Value=None):
        data_cmd(dataefetivaini)

    def dataefetivafim_cmd(Value=None):
        data_cmd(dataefetivafim)

    dimension = widgets.geometry(502, 950)
    if tipo_mov == 0:
        title = 'RECEBIMENTO DE MERCADORIAS'
        color = 'light goldenrod'
    elif tipo_mov == 1:
        title = 'PRODUÇÃO'
        color = 'powder blue'
    elif tipo_mov == 2:
        title = 'CONSUMO DE MATERIAIS'
        color = 'light salmon'
    elif tipo_mov == 3:
        title = 'VENDA DE BALCÃO - À VISTA'
        color = 'spring green'
    elif tipo_mov == 4:
        title = 'VENDA PERSONALIZADA'
        color = 'OliveDrab2'
    config = {'title': title,
              'dimension': dimension,
              'color': color}
    form_selec = main.form(config)
    wg = Widgets(form_selec, color)
    wg.label('', 3, 0, 0)
    wg.label('', 3, 3, 0)
    wg.label('', 3, 5, 0)
    wg.label('', 3, 7, 0)
    c.execute('SELECT Nome FROM Parceiros ORDER BY Nome')
    if tipo_mov in [3, 4]:
        c.execute('SELECT Nome FROM Parceiros WHERE Modo IN (0, 1) ORDER BY Nome')
        classe_parceiro = ['Cliente']
    elif tipo_mov in [0]:
        c.execute('SELECT Nome FROM Parceiros WHERE Modo IN (0, 2) ORDER BY Nome')
        classe_parceiro = ['Fornecedor']
    else:
        c.execute('SELECT Nome FROM Parceiros ORDER BY Nome')
        classe_parceiro = ['Parceiro']
    lista1 = []
    for row in c.fetchall():
        lista1.append(row[0])
    #Formulário base:
    if tipo_mov == 0:
        parceiro_sel = wg.combobox('  ' + classe_parceiro[0] + ': ', 16, lista1, 1, 1, cmd=parceiro_cmd, seek=parceiro_cmd)
        documento_in = wg.textbox('  Fatura: ', 8, 2, 1)
        datadocini = wg.textbox('  Emissão Inicial: ', 10, 1, 3, cmd=datadocini_cmd)
        datadocfim = wg.textbox('  Emissão Final: ', 10, 2, 3, cmd=datadocfim_cmd)
        dataefetivaini = wg.textbox('  Data Inicial: ', 10, 1, 5, default=data_brasil(ActualDate[0:7] + '-01'), cmd=dataefetivaini_cmd)
        dataefetivafim = wg.textbox('  Data Final: ', 10, 2, 5, default=data_brasil(str(lastdaymonth(ActualDate))), cmd=dataefetivafim_cmd)
    elif tipo_mov == 1:
        parceiro_sel = FalseRoutine()
        documento_in = wg.textbox('  Ordem Produção: ', 8, 1, 1)
        datadocini = wg.textbox('  Emissão Inicial: ', 10, 1, 3, cmd=datadocini_cmd)
        datadocfim = wg.textbox('  Emissão Final: ', 10, 2, 3, cmd=datadocfim_cmd)
        dataefetivaini = wg.textbox('  Data Inicial: ', 10, 1, 5, default=data_brasil(ActualDate[0:7] + '-01'), cmd=dataefetivaini_cmd)
        dataefetivafim = wg.textbox('  Data Final: ', 10, 2, 5, default=data_brasil(str(lastdaymonth(ActualDate))), cmd=dataefetivafim_cmd)
    elif tipo_mov == 2:
        parceiro_sel = FalseRoutine()
        documento_in = wg.textbox('  Ordem Produção: ', 8, 1, 1)
        datadocini = wg.textbox('  Emissão Inicial: ', 10, 1, 3, cmd=datadocini_cmd)
        datadocfim = wg.textbox('  Emissão Final: ', 10, 2, 3, cmd=datadocfim_cmd)
        dataefetivaini = wg.textbox('  Data Inicial: ', 10, 1, 5, default=data_brasil(ActualDate[0:7] + '-01'), cmd=dataefetivaini_cmd)
        dataefetivafim = wg.textbox('  Data Final: ', 10, 2, 5, default=data_brasil(str(lastdaymonth(ActualDate))), cmd=dataefetivafim_cmd)
    elif tipo_mov in [3, 4]:
        parceiro_sel = wg.combobox('  ' + classe_parceiro[0] + ': ', 16, lista1, 1, 1, cmd=parceiro_cmd, seek=parceiro_cmd)
        documento_in = wg.textbox('  Cupom: ', 8, 2, 1)
        datadocini = wg.textbox('  Emissão Inicial: ', 10, 1, 3, cmd=datadocini_cmd)
        datadocfim = wg.textbox('  Emissão Final: ', 10, 2, 3, cmd=datadocfim_cmd)
        dataefetivaini = wg.textbox('  Data Inicial: ', 10, 1, 5, default=data_brasil(ActualDate[0:7] + '-01'), cmd=dataefetivaini_cmd)
        dataefetivafim = wg.textbox('  Data Final: ', 10, 2, 5, default=data_brasil(str(lastdaymonth(ActualDate))), cmd=dataefetivafim_cmd)
    Seek = wg.button('Procurar', seek, 10, 0, 4, 1, 4)
    File = wg.button('Gerar csv', filecsv, 10, 0, 4, 3, 4)
    soma = wg.label('', 30, 4, 5, 4)
    '''
    if in_value or out_value:
        soma['text'] = 'Valor Total: ' + num_brasil(format(abs(generate_sql(value=True)), '.2f'))
    '''
    soma['fg'] = 'red'
    soma['font'] = 'Arial 12 bold'
    # grid setup ini
    if tipo_mov in [0, 3, 4]:
        colsf = ['Doc', 'Parceiro', 'DataDoc', 'DataEfetiva', 'VlTotal', 'Descricao']
        headf = {
            'Doc': {'text': 'NF/Fatura', 'width': 100},
            'Parceiro': {'text': classe_parceiro[0], 'width': 210},
            'DataDoc': {'text': 'Emissão', 'width': 80, 'format': 'date'},
            'DataEfetiva': {'text': 'Efetivação', 'width': 90, 'format': 'date'},
            'VlTotal': {'text': 'Valor Total', 'width': 80, 'anchor': 'e', 'format': 'float'},
            'Descricao': {'text': 'Descrição', 'width': 270},
        }
        if tipo_mov in [3, 4]:
            headf['Doc']['text'] = 'Cupom'
    elif tipo_mov in [1, 2]:
        colsf = ['Doc', 'DataDoc', 'DataEfetiva', 'TotalCusto', 'TotalVenda', 'Descricao']
        headf = {
            'Doc': {'text': 'Ordem Produção', 'width': 130},
            'DataDoc': {'text': 'Emissão', 'width': 80, 'format': 'date'},
            'DataEfetiva': {'text': 'Efetivação', 'width': 80, 'format': 'date'},
            'TotalCusto': {'text': 'Custo da Produção', 'width': 140, 'anchor': 'e', 'format': 'float'},
            'TotalVenda': {'text': 'Valor de Venda', 'width': 130, 'anchor': 'e', 'format': 'float'},
            'Descricao': {'text': 'Descrição', 'width': 270},
        }
    c.execute(generate_sql())
    doc = c.fetchall()
    lista = {}
    listaord = []
    if tipo_mov in [0, 3, 4]:
        for row in doc:
            texts = []
            rowact = 0
            for rows in row[1:]:
                if rowact in [2, 3]:
                    texts.append(data_brasil(rows))
                elif rowact in [4]:
                    texts.append(num_brasil(str(abs(rows))))
                else:
                    texts.append(rows)
                rowact += 1
            listaord.append(row[0])
            lista[row[0]] = tuple(texts)
    if tipo_mov in [1, 2]:
        for row in doc:
            texts = []
            rowact = 0
            for rows in row[1:]:
                if rowact in [1, 2]:
                    texts.append(data_brasil(rows))
                elif rowact in [3, 4]:
                    texts.append(num_brasil(str(abs(rows))))
                else:
                    texts.append(rows)
                rowact += 1
            listaord.append(row[0])
            lista[row[0]] = tuple(texts)
    itemselect = wg.grid(colsf, headf, lista, listaord, 14, 6, 1, colspan=8)
    # grid setup fim
    Seek.focus()
    Create = wg.button('Incluir', newreg, 10, 0, 8, 1, 1)
    Duplic = wg.button('Duplicar', dupreg, 10, 0, 8, 2, 1)
    if tipo_mov in [0, 1, 2]:
        texto = 'Editar'
    if tipo_mov in [3, 4]:
        texto = 'Ajuste data'
    Edit = wg.button(texto, editreg, 10, 0, 8, 3, 1)
    if tipo_mov in [0, 1, 2]:
        texto = 'Excluir'
    if tipo_mov in [3, 4]:
        texto = 'Detalhar'
    Remove = wg.button(texto, removereg, 10, 0, 8, 5, 2)
    Sair = wg.button('Sair', exitsc, 10, 0, 8, 8, 1)
    if userauth:
        if tipo_mov in [0, 1, 2] and userauth[5] < 3:
            Edit.configure(state='disabled')
        if tipo_mov  in [0, 1, 2] and userauth[5] < 2:
            Create.configure(state='disabled')
            Duplic.configure(state='disabled')
        if tipo_mov in [3, 4] and userauth[6] < 4:
            Edit.configure(state='disabled')
        if tipo_mov in [3, 4] and userauth[6] < 2:
            Create.configure(state='disabled')
            Duplic.configure(state='disabled')
    seek()

def materiais_movimentos_rec():
    materiais_movimentos(0)

def materiais_movimentos_prod():
    materiais_movimentos(1)

def materiais_movimentos_cons():
    materiais_movimentos(2)

def materiais_movimentos_vend():
    materiais_movimentos(3)

def materiais_movimentos_vend2():
    materiais_movimentos(4)

def materiais_consultas():
    def generate_sql(dtseek=''):
        adic = ''
        cabec = 'WHERE '
        # Filtrando DATA DO SALDO
        if Date.selection:
            dateseek = str(Date.selection.date())
            adic += cabec + 'Materiais_Movimentos.DataEfetiva <= "' + dateseek + '"'
            cabec = ' AND '
        # Filtrando TIPO DE MATERIAL
        if Compra.get() and Venda.get() and Revenda.get():
            pass
        else:
            itens = []
            if Compra.get():
                itens.append('0')
            if Venda.get():
                itens.append('1')
            if Revenda.get():
                itens.append('2')
            if len(itens) == 1:
                adic += cabec + 'Materiais_Itens.Tipo = ' + itens[0]
                cabec = ' AND '
            elif len(itens) == 2:
                adic += cabec + 'Materiais_Itens.Tipo IN (' + ', '.join(itens) + ')'
                cabec = ' AND '
            else:
                adic += cabec + 'Materiais_Itens.Tipo NOT IN (0, 1, 2)'
                cabec = ' AND '
        # Filtrando CATEGORIA
        if Categoria.get():
            c.execute('SELECT Id FROM Materiais_Categorias WHERE Categoria = "' + Categoria.get() + '"')
            adic += cabec + 'Materiais_Itens.Categoria = ' + str(c.fetchone()[0])
            cabec = ' AND '
        # Filtrando PRODUTO (contém caractere)
        if Produto.get():
            adic += cabec + 'Materiais_Itens.Produto LIKE "%' + Produto.get() + '%"'
            cabec = ' AND '
        command = ('SELECT Materiais_Itens.Id, Materiais_Itens.Produto, SUM(Qtde), ValorVenda, Unidade ' +
                   'FROM Materiais_Movimentos ' +
                   'JOIN Materiais_Itens ON Materiais_Itens.Id = Materiais_Movimentos.Produto ' + adic + ' '
                   'GROUP BY Materiais_Itens.Produto ORDER BY Materiais_Itens.Produto')
        c.execute(command)
        doc = c.fetchall()
        saldoatual = []
        soma = 0.0
        for row in doc:
            valorprodatual = round(row[2] * row[3], 2)
            soma += valorprodatual
            saldoatual.append((row[0], row[1] + ' (' + row[4] + ')', row[3], row[2], num_brasil(str(valorprodatual))))
        ValorEstoque['text'] = 'Valor R$: ' + num_brasil(format(abs(soma), '.2f'))
        return saldoatual

    def seek():
        #dtseek = ''
        #if Date.selection:
        #    dtseek = str(Date.selection.date())
        itemselect.delete(*itemselect.get_children())
        doc = generate_sql()
        combolist = {}
        ordlista = []
        somalc = [0.0, 0.0]
        for row in doc:
            texts = []
            rowact = 0
            for rows in row[1:]:
                if rowact in [1, 2]:
                    if rows < 0.0:
                        rows = -rows
                        texts.append('-' + num_brasil(format(rows, '.2f')))
                    else:
                        texts.append(num_brasil(format(rows, '.2f')))
                else:
                    texts.append(rows)
                rowact += 1
            ordlista.append(row[0])
            combolist[row[0]] = tuple(texts)
        for rows in ordlista:
            itemselect.insert('', 'end', text=rows, values=combolist[rows])
   
    def selected_item(arg=None):
        if itemselect.selection():
            for i in itemselect.selection():
                gridselected = str(itemselect.item(i, 'text'))
            materiais_extrato(produto=gridselected)

    def Categoria_cmd(value=None):
        widgets.combobox_return(Categoria, lista1)

    dimension = widgets.geometry(470, 920)
    config = {'title': 'Consulta saldo em estoque de mercadorias',
              'dimension': dimension,
              'color': 'light goldenrod'}
    form_selec = main.form(config)
    wg = Widgets(form_selec, 'light goldenrod')
    wg.label('', 5, 0, 0)
    wg.label('', 5, 2, 3)
    wg.label('', 5, 10, 0)
    Date = wg.calendar(1, 1, columnspan=2, command=seek)
    Compra = wg.check('Itens de compra: ', 15, ' Gerar                   ', 3, 1, True)
    Venda = wg.check('Itens de venda: ', 15, ' Gerar                   ', 4, 1, True)
    Revenda = wg.check('Itens de revenda: ', 15, ' Gerar                   ', 5, 1, True)
    c.execute('SELECT Categoria FROM Materiais_Categorias ORDER BY Categoria')
    lista1 = []
    for row in c.fetchall():
        lista1.append(row[0])
    Categoria =  wg.combobox('Categoria: ', 15, lista1, 6, 1, seek = Categoria_cmd)
    Produto = wg.textbox('Produto: ', 15, 7, 1)
    Seek = wg.button('Saldo na data', seek, 20, 0, 9, 1, 2)
    ValorEstoque = wg.label('Valor R$:', 30, 11, 1, colspan=2)
    ValorEstoque['fg'] = 'red'
    ValorEstoque['font'] = 'Arial 12 bold'
    doc = generate_sql()
    colsf = ['produto', 'unit', 'saldo', 'valor']
    headf = {
        'produto': {'text': 'Produto', 'width': 255},
        'unit': {'text': 'Preço R$', 'width': 65, 'anchor': 'e', 'format': 'float'},
        'saldo': {'text': 'Qtde (un)', 'width': 70, 'anchor': 'e', 'format': 'float'},
        'valor': {'text': 'Valor R$', 'width': 65, 'anchor': 'e', 'format': 'float'},
    }
    doc = generate_sql()
    lista = {}
    listaord = []
    itemselect = wg.grid(colsf, headf, lista, listaord, 20, 1, 4, rowspan=20, cmd=selected_item)

    seek()
    Categoria.focus()

def materiais_extrato(produto=None):
    def generate_sql(dtseek=''):
        adic = 'WHERE Materiais_Movimentos.Produto = ' + produto
        cabec = ' AND '
        # Filtrando DATA DO SALDO
        if dataini.get():
            datafill = _date(dataini.get())
            adic += cabec + 'Materiais_Movimentos.DataEfetiva >= "' + datafill + '"'
        if datafim.get():
            datafill = _date(datafim.get())
            adic += cabec + 'Materiais_Movimentos.DataEfetiva <= "' + datafill + '"'
        # Filtrando TIPO DE MOVIMENTO
        if Recebimento.get() and Producao.get() and Consumo.get() and Venda.get():
            pass
        else:
            itens = []
            if Recebimento.get():
                itens.append('0')
            if Producao.get():
                itens.append('1')
            if Consumo.get():
                itens.append('2')
            if Venda.get():
                itens.append('3')
            if len(itens) == 1:
                adic += cabec + 'Materiais_Documentos.Tipo = ' + itens[0]
            elif len(itens) in [2, 3]:
                adic += cabec + 'Materiais_Documentos.Tipo IN (' + ', '.join(itens) + ')'
            else:
                adic += cabec + 'Materiais_Documentos.Tipo NOT IN (0, 1, 2, 3)'
        # Filtrando DATA DO SALDO
        if Op.get():
            adic += cabec + 'Materiais_Documentos.Doc LIKE "%' + Op.get() + '%"'

        command = ('SELECT Materiais_Movimentos.Id, Materiais_Movimentos.DataAtualiza, Materiais_Movimentos.DataEfetiva, ' +
                    'Materiais_Documentos.Tipo, Materiais_Documentos.Parceiro, Materiais_Documentos.Doc, Qtde, Materiais_Documentos.Descricao ' +
                    'FROM Materiais_Movimentos ' +
                    'JOIN Materiais_Itens ON Materiais_Itens.Id = Materiais_Movimentos.Produto ' +
                    'JOIN Materiais_Documentos ON Materiais_Documentos.Id = Materiais_Movimentos.Documento ' + adic + ' '
                    'ORDER BY Materiais_Movimentos.DataEfetiva')
        c.execute(command)
        doc = c.fetchall()
        saldoatual = []
        c.execute('SELECT Id, Nome FROM Parceiros')
        parceiros_nomes = {}
        parceiros_nomes[0] = '-Mov interno-'
        for row in c.fetchall():
            parceiros_nomes[row[0]]= row[1]
        for row in doc:
            saldoatual.append((row[0], row[1], row[2], row[3], parceiros_nomes[row[4]], row[5], row[6]))
        #ValorEstoque['text'] = 'Valor R$: ' + num_brasil(format(abs(soma), '.2f'))
        return saldoatual
    
    def seek():
        itemselect.delete(*itemselect.get_children())
        doc = generate_sql()
        combolist = {}
        ordlista = []
        somalc = [0.0, 0.0]

        tipos = ['Recebimento', 'Produção', 'Consumo', 'Venda']
        saldo_final = 0.0
        if IncSaldo.get():
            c.execute('SELECT SUM(Qtde) FROM Materiais_Movimentos WHERE Produto = ' + produto + ' AND DataEfetiva < "' + _date(dataini.get()) + '"')
            resp = c.fetchone()[0]
            if resp:
                combolist['0'] = '', data_brasil(_date(dataini.get())), '   Saldo Inicial', '', '', num_brasil(str(resp)), ''
                saldo_final = resp
            else:
                combolist['0'] = '', data_brasil(_date(dataini.get())), '   Saldo Inicial', '', '', '0,00', ''
            ordlista.append('0')
        for row in doc:
            texts = []
            rowact = 0
            for rows in row[1:]:
                if rowact in [0]:
                    texts.append(datahora_brasil(rows))
                elif rowact in [1]:
                    texts.append(data_brasil(rows))
                elif rowact in [2]:
                    texts.append(tipos[int(rows)])
                elif rowact in [5]:
                    saldo_final += rows
                    if rows < 0.0:
                        rows = -rows
                        texts.append('-' + num_brasil(format(rows, '.2f')))
                    else:
                        texts.append(num_brasil(format(rows, '.2f')))
                else:
                    texts.append(rows)
                rowact += 1
            ordlista.append(row[0])
            combolist[row[0]] = tuple(texts)
        if IncSaldo.get():
            combolist['SF'] = '', data_brasil(_date(datafim.get())), '   Saldo Final', '', '', num_brasil(str(round(saldo_final, 2))), ''
            ordlista.append('SF')
        for rows in ordlista:
            itemselect.insert('', 'end', text=rows, values=combolist[rows])
        c.execute('SELECT Produto, Unidade FROM Materiais_Itens WHERE Id = ' + produto)
        resp = c.fetchone()
        Produto.configure(state='normal')
        Produto.delete(0, 'end')
        Produto.insert(0, resp[0] + ' (' + resp[1] + ')')
        Produto.configure(state='disabled')

    def dataini_cmd(Value=None):
        data_cmd(dataini)

    def datafim_cmd(Value=None):
        data_cmd(datafim)

    dimension = widgets.geometry(470, 920)
    config = {'title': 'Consulta movimentação de produto',
              'dimension': dimension,
              'color': 'light goldenrod'}
    form_selec = main.form(config)
    wg = Widgets(form_selec, 'light goldenrod')
    wg.label('', 3, 0, 0)
    wg.label('', 3, 5, 3)
    wg.label('', 3, 10, 0)
    Recebimento = wg.check('Recebimento: ', 10, ' Exibir             ', 1, 1, True)
    Producao = wg.check('Produção: ', 10, ' Exibir             ', 2, 1, True)
    Consumo = wg.check('Consumo: ', 10, ' Exibir             ', 3, 1, True)
    Venda = wg.check('Venda: ', 10, ' Exibir             ', 4, 1, True)
    Produto = wg.textbox('Produto: ', 35, 1, 3)
    Op = wg.textbox('Doc/Ordem Prod: ', 15, 2, 3)
    dataini = wg.textbox('Data Inicial: ', 10, 3, 3, default=data_brasil(ActualDate[0:7] + '-01'), cmd=dataini_cmd)
    datafim = wg.textbox('Data Final: ', 10, 4, 3, default=data_brasil(str(lastdaymonth(ActualDate))), cmd=datafim_cmd)
    Seek = wg.button('Buscar', seek, 15, 0, 5, 2, 1)
    IncSaldo = wg.check('Incluir saldos iniciais e finais: ', 10, ' Sim             ', 5, 3, False)
    colsf = ['data_atualiza', 'data_efetiva', 'tipo', 'parceiro', 'doc', 'qtde', 'descricao']
    headf = {
        'data_atualiza': {'text': 'Última Edição', 'width': 145, 'format': 'date/time'},
        'data_efetiva': {'text': 'Data Mov', 'width': 90, 'format': 'date/time'},
        'tipo': {'text': 'Tipo', 'width': 95},
        'parceiro': {'text': 'Parceiro', 'width': 140,},
        'doc': {'text': 'Doc', 'width': 80,},
        'qtde': {'text': 'Volume', 'width': 80, 'anchor': 'e', 'format': 'float'},
        'descricao': {'text': 'Descrição', 'width': 170,},
    }
    #doc = generate_sql()
    lista = {}
    listaord = []
    itemselect = wg.grid(colsf, headf, lista, listaord, 15, 6, 1, colspan=4, rowspan=20)
    Produto.configure(state='disabled')
    seek()
    Op.focus()

def materiais_listacompras():
    def generate_sql(tipo=[], categoria='', produto=''):
        adic = 'WHERE '
        texto = ''
        if len(tipo) < 3 and tipo:
            texto += adic + 'Materiais_Itens.Tipo IN (' + ', '.join(tipo) + ') '
            adic = 'AND '
        elif not tipo:
            texto += adic + 'Materiais_Itens.Tipo IN (4) '
            adic = 'AND '
        if categoria:
            texto += adic + 'Materiais_Categorias.Categoria = "' + categoria + '" '
            adic = 'AND '
        if produto:
            texto += adic + 'Materiais_Itens.Produto LIKE "%' + produto + '%" '
        cmd = 'SELECT Materiais_Itens.Id, Materiais_Itens.Produto, Materiais_Categorias.Categoria, '\
              'EstoqueMin, SUM(Qtde), ValorCusto, Unidade '\
              'FROM Materiais_Itens '\
              'JOIN Materiais_Categorias ON Materiais_Categorias.Id = Materiais_Itens.Categoria '\
              'JOIN Materiais_Movimentos ON Materiais_Movimentos.Produto = Materiais_Itens.Id ' + texto + ''\
              'GROUP BY Materiais_Itens.Id'
        c.execute(cmd)
        doc1 = c.fetchall()
        csv_dicts = []
        new_dict = {}
        ord_dict = []
        linha = 1
        try:
            if doc1[0][0]:
                for row in doc1:
                    if float(row[3]) > float(row[4]):
                        # nova query para tela
                        new_list = []
                        new_list.append(row[1] + ' (' + row[6] + ')')
                        new_list.append(row[2])
                        new_list.append(num_brasil(str(round(row[3], 2))))
                        new_list.append(num_brasil(str(round(row[4], 2))))
                        cmd = 'SELECT MAX(Materiais_Movimentos.DataEfetiva), Qtde, Parceiros.Nome '\
                            'FROM Materiais_Movimentos '\
                            'JOIN Materiais_Documentos ON Materiais_Documentos.Id = Materiais_Movimentos.Documento '\
                            'JOIN Parceiros ON Parceiros.Id = Materiais_Documentos.Parceiro '\
                            'WHERE Produto = ' + str(row[0]) + ' AND Materiais_Documentos.Tipo = 0'
                        c.execute(cmd)
                        doc2 = c.fetchone()
                        new_list.append(data_brasil(doc2[0]))
                        new_list.append(num_brasil(str(round(row[5], 2))))
                        new_list.append(num_brasil(str(round(doc2[1], 2))))
                        new_dict[linha] = tuple(new_list)
                        ord_dict.append(linha)
                        # nova query para arquivo csv
                        csv_dict = {}
                        csv_dict['Produto'] = row[1] + ' (' + row[6] + ')'
                        csv_dict['Categoria'] = row[2]
                        csv_dict['Estoque mínimo'] = num_brasil(str(row[3]))
                        csv_dict['Estoque atual'] = num_brasil(str(round(row[4], 2)))
                        csv_dict['Data da última compra'] = data_brasil(doc2[0])
                        csv_dict['Valor da última compra'] = num_brasil(str(row[5]))
                        csv_dict['Qtde comprada'] = num_brasil(str(doc2[1]))
                        csv_dict['Fornecedor'] = doc2[2]
                        csv_dicts.append(csv_dict)
                        linha += 1
        except:
            pass
        return new_dict, ord_dict, csv_dicts

    def seek():
        tipo = []
        if Compra.get():
            tipo.append('0')
        if Venda.get():
            tipo.append('1')
        if Revenda.get():
            tipo.append('2')
        doc = generate_sql(tipo, Categoria.get(), Material.get())
        headf = {
            'produto': {'text': 'Produto', 'width': 205},
            'categoria': {'text': 'Categoria', 'width': 165},
            'estoque_minimo': {'text': 'Qtde Mín', 'width': 85, 'anchor': 'e', 'format': 'float'},
            'estoque_atual': {'text': 'Qtde Atual', 'width': 90, 'anchor': 'e', 'format': 'float'},
            'data_ult': {'text': 'Última Compra', 'width': 110, 'format': 'date'},
            'valor_ult': {'text': 'Preço', 'width': 80, 'anchor': 'e', 'format': 'float'},
            'qtde_ult': {'text': 'Volume', 'width': 80, 'anchor': 'e', 'format': 'float'},
        }
        #doc = generate_sql()
        combolist = doc[0]
        ordlista = doc[1]
        itemselect = wg.grid(colsf, headf, combolist, ordlista, 13, 5, 1, 4)

    def generate_csv():
        tipo = []
        if Compra.get():
            tipo.append('0')
        if Venda.get():
            tipo.append('1')
        if Revenda.get():
            tipo.append('2')
        doc = generate_sql(tipo, Categoria.get(), Material.get())
        post_file = datetime.now().strftime('%d_%m_%Y')
        file_output = "csv/compras_" + post_file + ".csv"
        with open(file_output, 'w') as csvfile:
            fieldnames = list(doc[2][0].keys())
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames, delimiter=';')
            writer.writeheader()
            writer.writerows(doc[2])
        messagebox.showinfo(title="Atividade concluída!",
                            message="Arquivo 'compras_" + post_file + ".csv' gerado na pasta csv")
        Seek.focus()

    def generate_email():
        tipo = []
        if Compra.get():
            tipo.append('0')
        if Venda.get():
            tipo.append('1')
        if Revenda.get():
            tipo.append('2')
        doc = generate_sql(tipo, Categoria.get(), Material.get())
        try:
            c.execute('SELECT Nome, Email FROM Auth')
            resp = c.fetchone()
            if resp[1]:
                if userauth:
                    user = userauth[1]
                else:
                    user = 'ADMIN'
                msg = EmailMessage()
                msg['Subject'] = 'Lista de compras emitida em ' + datetime.now().strftime('%d/%m/%y às %H:%M')
                msg['From'] = 'rhsprogramas@gmail.com'
                msg['To'] = resp[1]
                information = ['NECESSIDADE DE COMPRAS - SOFTWARE: CFA\n\n']
                if resp[0]:
                    information.append(resp[0] + ',\n\n')
                else:
                    information.append('Prezado usuário,\n\n')
                information.append('Foi gerada no programa CFA a relação de materiais que possuem volume em estoque '\
                    'inferior ao saldo mínimo. Seguem as informações abaixo para proceder com a cotação.\n\n'
                )
                for row in doc[2]:
                    lista_field = list(row.keys())
                    for row_f in lista_field:
                        information.append(row_f + ': ' + row[row_f] + '\n')
                    information.append('\n')
                information.append('\nInformação enviada por requisição do usuário ' + user + '.\n\n')
                information.append('Enviado por:\n\nDisparo automático CFA\nRHS Programas e Desenvolvimentos')
                information = ''.join(information)
                msg.set_content(information)
        
                # Send the message via our own SMTP server.
                s = smtplib.SMTP('smtp.gmail.com:587')
                s.starttls()
                s.login('rhsprogramas@gmail.com', 'Asdf1234#')
                s.send_message(msg)
                s.quit()
                messagebox.showinfo('Aviso', 'O email com a lista de compras foi enviado com sucesso.')
                Seek.focus()
            else:
                messagebox.showwarning('Atenção', 'Não há email cadastrado para enviar a lista de compras.\nAcesse o menu Cadastros/Sistema/Licença e informe cadastre seus dados.')
                Seek.focus()
        except:
            messagebox.showwarning('Atenção', 'Não foi possível enviar o email.\nVerifique se você está conectado à internet. Em caso de dúvidas, procure o administrador do sistema.')
            Seek.focus()

    dimension = widgets.geometry(500, 920)
    config = {'title': 'Geração de lista de compras',
              'dimension': dimension,
              'color': 'light goldenrod'}
    config['title'] = 'Lista de compra de materiais com volume em estoque abaixo do estoque mínimo.'
    colsf = ['produto', 'categoria', 'estoque_minimo', 'estoque_atual', 'data_ult', 'valor_ult', 'qtde_ult']
    form_selec = main.form(config)
    wg = Widgets(form_selec, 'light goldenrod')
    wg.label('', 2, 0, 0)
    wg.label('', 2, 2, 2)
    wg.label('', 2, 4, 0)
    wg.image('',0,'images/compras.png', 1, 1, 1, 4, imagewidth=(128, 128))
    Year = ActualDate[0:4]
    #lista1 = ['Compra', 'Venda', 'Revenda']
    Tipo = None
    c.execute('SELECT Categoria FROM Materiais_Categorias')
    lista1 = []
    for row in c.fetchall():
        lista1.append(row[0])
    Compra = wg.check('Itens de compra: ', 15, ' Gerar                   ', 1, 2, True)
    Venda = wg.check('Itens de venda: ', 15, ' Gerar                   ', 2, 2, True)
    Revenda = wg.check('Itens de revenda: ', 15, ' Gerar                   ', 3, 2, True)
    Categoria = wg.combobox('Categorias: ', 20, lista1, 1, 3)
    Material = wg.textbox('Material: ', 20, 2, 3)
    Seek = wg.button('Exibir', seek, 20, 0, 4, 2, 1)
    Seek = wg.button('Gerar CSV', generate_csv, 20, 0, 4, 3, 1)
    Seek = wg.button('Enviar email', generate_email, 20, 0, 4, 4, 1)
    #doc = generate_sql()
    Seek.focus()
    seek()

def fluxocaixa():
    def generate_sql(value=False):
        def _fc(first):
            if first:
                resp = 'WHERE '
            else:
                resp = ' AND '
            return resp

        SaldoInicial = 0.0
        NaoRealizado = 0.0
        JaRealizado = 0.0
        dados = []
        for i in range(6):
            first = True
            command = []
            if i in [0, 1, 2, 3, 5]: # 0 Saldo Inicial   1 Não Efetivado   2 e 3 Já efetivado
                command.append('SELECT SUM(Valor) ')
            else:
                command.append('SELECT Diario.Id, Bancos.NomeBanco, Parceiros.Nome, Categorias.Categoria, Valor, DataVenc, DataPago, Diario.TipoMov ')
            command.append('FROM Diario JOIN Parceiros ON Parceiros.Id = Parceiro JOIN Categorias ON Categorias.Id = CategoriaMov ')
            command.append('JOIN Bancos ON Bancos.Id = Banco ')
            if bancoini.get():
                if i == 5:
                    if bancoini.get() == 'Todos os bancos':
                        command.append(_fc(first) + 'Bancos.TipoMov = 2')
                else:
                    if bancoini.get() == 'Todos os bancos':
                        command.append(_fc(first) + 'Bancos.TipoMov < 2')
                    else:
                        command.append(_fc(first) + 'Bancos.NomeBanco = "' + bancoini.get() + '"')
                first = False
            if datavencini.get():
                datafill = _date(datavencini.get())
                if i in [0]:
                    command.append(_fc(first) + 'DataPago < "' + datafill + '" AND DataPago <> ""')
                elif i in [1]:
                    command.append(_fc(first) + 'DataVenc < "' + datafill + '" AND DataPago = ""')
                elif i in [2]:
                    command.append(_fc(first) + 'DataVenc >= "' + datafill + '" AND DataPago < "' + datafill + '" AND DataPago <> ""')
                elif i in [3]:
                    command.append(_fc(first) + 'DataVenc < "' + datafill + '" AND DataPago >= "' + datafill + '" AND DataPago <> ""')
                else:
                    command.append(_fc(first) + 'DataVenc >= "' + datafill + '"')
                first = False
            else:
                if i in [0, 1, 2, 3]:
                    continue
            if datavencfim.get():
                if i in [0, 1, 2, 3]:
                    pass
                else:
                    datafill = _date(datavencfim.get())
                    command.append(_fc(first) + 'DataVenc <= "' + datafill + '"')
                if i in [7]:
                    command.append(_fc(first) + 'DataPago = ""')
                first = False
            #command.append(' ORDER BY DataVenc')
            command.append(' ORDER BY CASE WHEN DataPago = "" THEN "ZZZZ" END, DataPago ASC, DataVenc ASC')
            c.execute(str.join('', command))
            if i in [0]:
                var = c.fetchone()[0]
                if var:
                    SaldoInicial += var
                dados.append(('A', '', '', 'Saldo Inicial', SaldoInicial, 0.0, 0.0, '', ''))
            elif i in [1]:
                var = c.fetchone()[0]
                if var:
                    NaoRealizado += var
                if NaoRealizado:
                    dados.append(('B', '', '', 'Não efetivado', NaoRealizado, 0.0, SaldoInicial + NaoRealizado, '', ''))
            elif i in [2]:
                var = c.fetchone()[0]
                if var:
                    JaRealizado -= var
            elif i in [3]:
                var = c.fetchone()[0]
                if var:
                    JaRealizado += var
                if JaRealizado:
                    dados.append(('C', '', '', 'Já efetivado', JaRealizado, 0.0, SaldoInicial + NaoRealizado + JaRealizado, '', ''))
            elif i in [4]:
                SaldoAtual = SaldoInicial + NaoRealizado + JaRealizado
                var = c.fetchall()
                for rows in var:
                    linha = []
                    for rows2 in rows[:-1]:
                        linha.append(rows2)
                        if len(linha) == 4 and rows[-1:][0] in [1, 4]:
                            linha.append(0.0)
                            SaldoAtual += rows[4]
                        if len(linha) == 5 and rows[-1:][0] in [0, 3]:
                            linha.append(0.0)
                            SaldoAtual += rows[4]
                        if len(linha) == 6:
                            linha.append(SaldoAtual)
                    dados.append(tuple(linha))
            elif i in [5] and bancoini.get() == 'Todos os bancos':
                var = c.fetchone()[0]
                if var:
                    if abs(var) >= 0.01:
                        Faturas = var
                        dados.append(('F', '', '', 'Faturas cartões crédito', 0.0, Faturas, SaldoAtual + Faturas, '', ''))
        return dados

    def pagreg():
        if itemselect.selection():
            for i in itemselect.selection():
                gridselected = str(itemselect.item(i, 'text'))
            mens = ('Tem certeza que deseja confirmar a baixa dos documentos selecionados?\n\n' +
                    'Os documentos serão baixados na data de hoje: ' + data_brasil(ActualDate))
            title = 'Baixa de documentos pelo Fluxo de Caixa'
            if messagebox.askyesno(title=title, message=mens, parent=form_selec):
                notrelease = []
                for i in itemselect.selection():
                    c.execute('SELECT DataPago, TipoMov FROM Diario WHERE Id = ' + str(itemselect.item(i, 'text')))
                    opt = c.fetchone()
                    if not opt[0] and opt[1] < 2:
                        command = []
                        command.append("UPDATE Diario SET ")
                        command.append("DataPago = '" + ActualDate + "' ")
                        command.append("WHERE Id = " + str(itemselect.item(i, 'text')))
                        c.execute(str.join('', command))
                        conn.commit()
                    else:
                        notrelease.append(str(itemselect.item(i, 'text')))
                if notrelease:
                    if len(notrelease) < 2:
                        message = 'O documento ' + notrelease[0] + ' já foi baixado ou é uma transferência! Revise.'
                    else:
                        message = 'Os documentos ' 
                        for i in notrelease:
                            if i not in notrelease[-2:]:
                                message = message + str(i) + ', '
                            elif i == notrelease[-2]:
                                message = message + str(i) + ' e '
                            else:
                                message = message + str(i) + ' '
                        message = message + ' já foram baixados ou são transferências! Revise.'
                    messagebox.showerror(title='Aviso', message=message)
                seek()
            bancoini.focus()
        else:
            messagebox.showerror(title='Aviso', message='Não foi selecionado nenhum documento para a baixa!')
            bancoini.focus()

    def seek():
        itemselect.delete(*itemselect.get_children())
        Entradas['text'] = 'Entradas no período: 0,00'
        Saidas['text'] = 'Saídas no período: 0,00'
        Saldo['text'] = 'Balanço líquido: 0,00'
        doc = generate_sql()
        combolist = {}
        ordlista = []
        somalc = [0.0, 0.0]
        for row in doc:
            texts = []
            rowact = 0
            for rows in row[1:]:
                if rowact in [6, 7]:
                    texts.append(data_brasil(rows))
                elif rowact in [3, 4, 5]:
                    if rows < 0.0:
                        rows = -rows
                        texts.append('-' + num_brasil(format(rows, '.2f')))
                    else:
                        texts.append(num_brasil(format(rows, '.2f')))
                    if rowact == 3 and isinstance(row[0], int):
                        somalc[0] += rows
                    if rowact == 4:
                        somalc[1] += rows
                else:
                    texts.append(rows)
                rowact += 1
            ordlista.append(row[0])
            combolist[row[0]] = tuple(texts)
        for rows in ordlista:
            itemselect.insert('', 'end', text=rows, values=combolist[rows])
        Entradas['text'] = 'Entradas no período: ' + num_brasil(format(somalc[0], '.2f'))
        Saidas['text'] = 'Saídas no período: ' + num_brasil(format(somalc[1], '.2f'))
        somas = somalc[0] - somalc[1]
        if somas < 0.0:
            somas = -somas
            Saldo['text'] = 'Balanço líquido: -' + num_brasil(format(somas, '.2f'))
        else:
            Saldo['text'] = 'Balanço líquido: ' + num_brasil(format(somas, '.2f'))

    def bancoini_cmd(value=None):
        widgets.combobox_return(bancoini, lista1)

    def data_cmd(value=None):
        if datavencini.get():
            resp = _date(datavencini.get())
            datavencini.delete(0, 'end')
            datavencini.insert(0, data_brasil(resp))

        if datavencfim.get():
            resp = _date(datavencfim.get())
            datavencfim.delete(0, 'end')
            datavencfim.insert(0, data_brasil(resp))

    dimension = widgets.geometry(502, 950)
    title = 'FLUXO DE CAIXA - Visualização por eventos'
    color = 'light goldenrod'
    config = {'title': title,
              'dimension': dimension,
              'color': color}
    form_selec = main.form(config)
    wg = Widgets(form_selec, color)
    wg.label('', 3, 0, 0)
    wg.label('', 3, 3, 0)
    wg.label('', 3, 5, 0)
    wg.label('', 3, 7, 0)
    wg.image('',0,'images/fluxo-de-caixa.png', 1, 1, colspam=1, rowspan=4)
    c.execute('SELECT NomeBanco FROM Bancos ORDER BY NomeBanco')
    lista1 = []
    lista1.append('Todos os bancos')
    for row in c.fetchall():
        lista1.append(row[0])
    bancoini = wg.combobox('  Banco: ', 16, lista1, 1, 3, seek=bancoini_cmd)
    datavencini = wg.textbox('  Venc Inicial: ', 10, 2, 3, cmd=data_cmd)
    datavencfim = wg.textbox('  Venc Final: ', 10, 3, 3, cmd=data_cmd)
    Seek = wg.button('Procurar', seek, 10, 0, 4, 3, 2)
    Pag = wg.button('Pagar compromissos selecionados', pagreg, 40, 0, 4, 5, 4)
    Entradas = wg.label('', 30, 1, 5, 4)
    #Entradas['fg'] = 'blue'
    Entradas['font'] = 'Arial 12 bold'
    Saidas = wg.label('', 30, 2, 5, 4)
    #Saidas['fg'] = 'blue'
    Saidas['font'] = 'Arial 12 bold'
    Saldo = wg.label('', 30, 3, 5, 4)
    #Saldo['fg'] = 'blue'
    Saldo['font'] = 'Arial 12 bold'
    # grid setup ini
    colsf = ['banco', 'parceiro', 'categoria', 'entrada', 'saida', 'saldo', 'vencimento', 'pagamento']
    headf = {
        'banco': {'text': 'Banco', 'width': 140},
        'parceiro': {'text': 'Parceiro', 'width': 140},
        'categoria': {'text': 'Categoria', 'width': 160},
        'entrada': {'text': 'Vl Ent', 'width': 70, 'anchor': 'e'},
        'saida': {'text': 'Vl Sai', 'width': 70, 'anchor': 'e'},
        'saldo': {'text': 'Saldo', 'width': 70, 'anchor': 'e'},
        'vencimento': {'text': 'Vencimento', 'width': 90},
        'pagamento': {'text': 'Pagamento', 'width': 90},
    }
    doc = generate_sql()
    lista = {}
    listaord = []
    itemselect = wg.grid(colsf, headf, lista, listaord, 14, 6, 1, colspan=8, order=False)
    # grid setup fim
    bancoini.focus()

def saldobancario():
    def generate_sql(dtseek=''):
        adic = ''
        if dtseek:
            adic = 'AND DataPago <= "' + dtseek + '" '
        command = ('SELECT Id, NomeBanco ' +
                   'FROM Bancos ' +
                   'WHERE TipoMov <= 1 ' +
                   'ORDER BY NomeBanco')
        c.execute(command)
        doc = c.fetchall()
        command = ('SELECT Bancos.Id, NomeBanco, SUM(Valor) ' +
                   'FROM Diario JOIN Bancos ON Bancos.Id = Banco ' +
                   'WHERE DataPago <> "" ' + adic +
                   'AND Bancos.TipoMov <= 1 ' +
                   'GROUP BY NomeBanco ORDER BY NomeBanco')
        c.execute(command)
        doc1 = c.fetchall()
        saldoatual = []
        for row in doc:
            idatual = row[0]
            bancoatual = row[1]
            saldobancoatual = 0.0
            for row1 in doc1:
                if idatual == row1[0]:
                    saldobancoatual += row1[2]
            saldoatual.append((idatual, bancoatual, saldobancoatual))
        return saldoatual

    def seek():
        dtseek = ''
        if Date.selection:
            dtseek = str(Date.selection.date())
        itemselect.delete(*itemselect.get_children())
        doc = generate_sql(dtseek)
        combolist = {}
        ordlista = []
        somalc = [0.0, 0.0]
        for row in doc:
            texts = []
            rowact = 0
            for rows in row[1:]:
                if rowact in [1]:
                    if rows < 0.0:
                        rows = -rows
                        texts.append('-' + num_brasil(format(rows, '.2f')))
                    else:
                        texts.append(num_brasil(format(rows, '.2f')))
                else:
                    texts.append(rows)
                rowact += 1
            ordlista.append(row[0])
            combolist[row[0]] = tuple(texts)
        for rows in ordlista:
            itemselect.insert('', 'end', text=rows, values=combolist[rows])
   
    dimension = widgets.geometry(370, 640)
    config = {'title': 'Consulta a saldo de contas de movimentação',
              'dimension': dimension,
              'color': 'light goldenrod'}
    form_selec = main.form(config)
    wg = Widgets(form_selec, 'light goldenrod')
    wg.label('', 5, 0, 0)
    wg.label('', 5, 2, 2)
    wg.label('', 5, 4, 0)
    wg.image('',0,'images/bancos.png', 1, 1, 1, 6)
    Date = wg.calendar(1, 3, command=seek)
    Seek = wg.button('Saldo na data', seek, 20, 0, 3, 3)
    doc = generate_sql()
    colsf = ['banco', 'saldo']
    headf = {
        'banco': {'text': 'Banco', 'width': 140},
        'saldo': {'text': 'Saldo', 'width': 70, 'anchor': 'e', 'format': 'float'},
    }
    doc = generate_sql()
    lista = {}
    listaord = []
    itemselect = wg.grid(colsf, headf, lista, listaord, 4, 5, 3)

    seek()
    Date.focus()

def sintetico(value):
    def generate_sql(dtseek=''):
        adic = ''
        if value == 0:
            command = ('SELECT Id, Categoria ' +
                    'FROM Categorias ' +
                    'WHERE TipoMov = 1 AND Repete = 1 ' +
                    'ORDER BY Categoria')
        elif value == 1:
            command = ('SELECT Id, Nome ' +
                    'FROM Parceiros ' +
                    'ORDER BY Nome')
        c.execute(command)
        doc = c.fetchall()
        if dtseek:
            periodos = dtseek.split(';')
            if periodos[0]:
                adic = 'AND DataDoc >= "' + periodos[0] + '-01" '
            if periodos[1]:
                adic += 'AND DataDoc <= "' + periodos[1] + '-31" '
        if value == 0:
            command = ('SELECT Categorias.Id, Categoria, Valor, DataDoc ' +
                    'FROM Diario JOIN Categorias ON Categorias.Id = CategoriaMov ' +
                    'WHERE DataDoc <> "" ' + adic +
                    'AND Diario.TipoMov IN (1, 4) ')
        elif value == 1:
            command = ('SELECT Parceiros.Id, Nome, Valor, DataDoc ' +
                    'FROM Diario JOIN Parceiros ON Parceiros.Id = Parceiro ' +
                    'WHERE DataDoc <> "" ' + adic +
                    'AND Diario.TipoMov IN (1) ')
        c.execute(command)
        doc1 = c.fetchall()
        Meses = []
        for row in doc1:
            if row[3][0:7] not in Meses:
                Meses.append(row[3][0:7])
        Meses.sort()
        saldoatual = []
        somameses = {}
        for row in doc:
            idatual = row[0]
            categoria = row[1]
            valmeses = {}
            for row1 in Meses:
                if row1 not in somameses.keys():
                    somameses[row1] = 0.0
                valmeses[row1] = 0.0
            for row1 in doc1:
                if idatual == row1[0]:
                    valmeses[row1[3][0:7]] += row1[2]
                    somameses[row1[3][0:7]] += row1[2]
            retornar = []
            retornar.append(idatual)
            retornar.append(categoria)
            for row1 in Meses:
                retornar.append(valmeses[row1])
            saldoatual.append(tuple(retornar))
        retornar = []
        retornar.append('T')
        retornar.append('Total')
        for row1 in Meses:
            retornar.append(somameses[row1])
        saldoatual.append(tuple(retornar))
        return Meses, saldoatual

    def seek():
        dtseek = ''
        if Mesini.get() or Mesfim.get():
            dtseek = Mesini.get() + ';' + Mesfim.get()
        doc = generate_sql(dtseek)
        colsf = ['categoria']
        headf = {
            'categoria': {'text': 'Categoria', 'width': 140},
        }
        if value == 1:
            headf['categoria']['text'] = 'Parceiro'
        width = int(675 / len(doc[0]))
        for row in doc[0]:
            colsf.append(row)
            headf[row] = {'text': row, 'width': width, 'anchor': 'e', 'format': 'float'}
        combolist = {}
        ordlista = []
        somalc = [0.0, 0.0]
        for row in doc[1]:
            texts = []
            rowact = 0
            for rows in row[1:]:
                if rowact not in [0]:
                    if rows < 0.0:
                        rows = abs(rows)
                        texts.append(num_brasil(format(rows, '.2f')))
                    else:
                        texts.append(num_brasil(format(rows, '.2f')))
                else:
                    texts.append(rows)
                rowact += 1
            ordlista.append(row[0])
            combolist[row[0]] = tuple(texts)
        itemselect = wg.grid(colsf, headf, combolist, ordlista, 13, 5, 1, 4)

    def generate_csv():
        file_output = "csv/resumo_despesas.csv"
        dtseek = ''
        if Mesini.get() or Mesfim.get():
            dtseek = Mesini.get() + ';' + Mesfim.get()
        doc = generate_sql(dtseek)
        if value == 0:
            colsf = ['Categoria']
        elif value == 1:
            colsf = ['Parceiro']
        headf = {
            'categoria': {'text': 'Categoria', 'width': 140},
        }
        if value == 1:
            headf['categoria']['text'] = 'Parceiro'
        width = int(675 / len(doc[0]))
        for row in doc[0]:
            colsf.append(row)
            headf[row] = {'text': row, 'width': width, 'anchor': 'e', 'format': 'float'}
        somalc = [0.0, 0.0]
        with open(file_output, 'w') as csvfile:
            fieldnames = colsf
            writer = csv.DictWriter(csvfile, fieldnames=fieldnames, delimiter=';')
            writer.writeheader()
            for row in doc[1]:
                combolist = {}
                rowact = 0
                for rows in row[1:]:
                    if rowact != 0:
                        if -rows < 0.1:
                            combolist[fieldnames[rowact]] = num_brasil(format(rows, '.2f'))
                        else:
                            combolist[fieldnames[rowact]] = '-' + num_brasil(format(-rows, '.2f'))
                    else:
                        combolist[fieldnames[rowact]] = rows
                    rowact += 1
                writer.writerow(combolist)
        messagebox.showinfo(title="Atividade concluída!",
                            message="Arquivo 'resumo_despesas.csv' gerado na pasta csv")
        Mesini.focus()

    dimension = widgets.geometry(500, 920)
    config = {'title': 'Consulta sintética de despesas por categoria',
              'dimension': dimension,
              'color': 'light goldenrod'}
    if value == 1:
        config['title'] = 'Consulta sintética de despesas por parceiro'
    form_selec = main.form(config)
    wg = Widgets(form_selec, 'light goldenrod')
    wg.label('', 2, 0, 0)
    wg.label('', 2, 2, 2)
    wg.label('', 2, 4, 0)
    if value == 0:
        wg.image('',0,'images/categoria.png', 1, 1, 1, 3)
    if value == 1:
        wg.image('',0,'images/parceiros.png', 1, 1, 1, 3, imagewidth=(128, 128))
    Year = ActualDate[0:4]
    Mesini = wg.textbox('Mês inicial:', 8, 1, 2, Year + '-01')
    Mesfim = wg.textbox('Mês final:', 8, 2, 2, Year + '-12')
    Seek = wg.button('Exibir', seek, 20, 0, 3, 2, 1)
    Seek = wg.button('Gerar CSV', generate_csv, 20, 0, 3, 3, 1)
    doc = generate_sql()
    Mesini.focus()
    seek()

def sinteticocategoria():
    sintetico(0)

def sinteticoempresa():
    sintetico(1)

def import_data():
    if messagebox.askyesno(parent=main.root,
                           title="Importar registros do servidor",
                           message='Tem certeza que deseja importar os dados da última atualização do servidor? \
                                    Não será possível desfazer a operação!'):
        try:
            if os.path.exists('settings/1.dat'):
                c.execute('SELECT * FROM Drive WHERE Arquivo = "finance.db"')
                fields = c.fetchall()
                if fields:
                    gauth = GoogleAuth()
                    gauth.LocalWebserverAuth()
                    drive = GoogleDrive(gauth)
                    namefile = fields[0][1]
                    idfile = fields[0][2]
                    file1 = drive.CreateFile({'id': idfile})  # Initialize GoogleDriveFile instance with file id
                    file1.GetContentFile(namefile)            # Download file as 'catlove.png'
                    lastupdate = file1['modifiedDate'][0:10]
                    lastupdate = lastupdate.split('-')
                    lastupdate = [int(row) for row in lastupdate]
                    lasthour = file1['modifiedDate'][11:16]
                    lasthour = lasthour.split(':')
                    lasthour = [int(row) for row in lasthour]
                    international = datetime(lastupdate[0], lastupdate[1], lastupdate[2], lasthour[0], lasthour[1])
                    brasildate = international - timedelta(hours=3)
                    a = datetime(1,2,3)
                    lastupdate = str(brasildate.day).zfill(2) + '/' + str(brasildate.month).zfill(2) + '/' + str(brasildate.year)
                    lasthour = str(brasildate.hour) + ':' + str(brasildate.minute).zfill(2)
                    message = 'Banco de dados sincronizado com atualização realizada em ' + lastupdate + ' às ' + lasthour + '. Reabra o aplicativo.'
                    messagebox.showinfo(title="Importar registros do servidor",
                                        message=message)
                    main.root.quit()
            else:
                messagebox.showerror(title="Importar registros do servidor",
                                    message="Ainda não foi exportado nenhum arquivo!")
        except:
            messagebox.showerror(title="Importar registros do servidor",
                                message="O processo não foi realizado com sucesso!")
    else:
        pass

def export_data():
    if messagebox.askyesno(parent=main.root,
                            title="Exportar registros para o servidor",
                            message='Tem certeza que deseja exportar os dados em seu aplicativo para o servidor? \
                                    Não será possível desfazer a operação!'):
        try:
            gauth = GoogleAuth()
            gauth.LocalWebserverAuth()
            drive = GoogleDrive(gauth)
            c.execute('SELECT * FROM Drive WHERE Arquivo = "finance.db"')
            fields = c.fetchall()
            if fields:
                namefile = fields[0][1]
                idfile = fields[0][2]
                # print(namefile, idfile)
                file1 = drive.CreateFile({'id': idfile})  # Initialize GoogleDriveFile instance with file id
                file1.SetContentFile(namefile)            # Download file as 'catlove.png'
                file1.Upload()                            # Upload it
            else:
                # primeiro uso
                # verificar a existência do arquivo settings.yaml na raiz. Criar pasta Settings, para salvar os arquivos (1.dat, 2.dat, etc.)
                file1 = drive.CreateFile()  # Initialize GoogleDriveFile instance with file id
                file1.SetContentFile('finance.db') # Read file and set it as a content of this instance.
                file1.Upload() # Upload it
                commands = 'INSERT OR REPLACE INTO Drive (Id, Arquivo, Id_Google) VALUES ('
                commands += '1, "finance.db", "' + file1['id'] + '")'
                c.execute(commands)
                conn.commit()
                file1 = drive.CreateFile({'id': file1['id']})  # Initialize GoogleDriveFile instance with file id
                file1.SetContentFile('finance.db')            # Download file as 'catlove.png'
                file1.Upload()                            # Upload it
        except:
            messagebox.showerror(title="Exportar registros para o servidor",
                                message="O processo não foi realizado com sucesso!")
        messagebox.showinfo(title="Exportar registros para o servidor",
                            message="Atividade concluída!")
    else:
        pass

def dashboard():
    def parceiros_call(value):
        parceiros()

    def produtos_call(value):
        materiais_itens()

    def movin_call(value):
        movimentos_in()

    def movout_call(value):
        movimentos_out()

    def transf_call(value):
        movimentos()

    def fc_call(value):
        fluxocaixa()

    def sb_call(value):
        saldobancario()

    def sc_call(value):
        sinteticocategoria()
    
    def gp_call(value):
        materiais_consultas()

    def compras_call(value):
        materiais_movimentos_rec()

    def producao_call(value):
        materiais_movimentos_prod()

    def consumo_call(value):
        materiais_movimentos_cons()

    def vendas_call(value):
        materiais_movimentos_vend()

    def vendas2_call(value):
        materiais_movimentos_vend2()

    def lista_call(value):
        materiais_listacompras()

    def sobre(value):
        dimension = widgets.geometry(190, 440)
        config = {'title': 'Sobre ...',
                'dimension': dimension,
                'color': 'light goldenrod'}
        form_selec = main.form(config)
        wg = Widgets(form_selec, 'light goldenrod')
        wg.label('', 3, 0, 0)
        wg.label('', 3, 2, 2)
        wg.label('', 3, 6, 0)
        wg.image('', 0,'images/finance.png', 1, 1, rowspan=3, imagewidth=(64, 64))
        wg.label('CFA - Consultor Financeiro e Administrativo', 40, 1, 2)
        wg.label('Cuidando bem de seus recursos!', 40, 2, 2)
        wg.label('Versão 1.1', 40, 3, 2)
        wg.label('(c) 2018 RHS Desenvolvimento', 40, 4, 2)
        wg.label('Responsável: Renan Hernandes de Souza', 40, 5, 2)
        Seek = wg.button('Ok', form_selec.destroy, 20, 0, 7, 1, colspan=2)

    def quit(value):
        appquit()

    if userauth:
        opcs_dic = {
            '0': ['Parceiros', 'images/users.png', parceiros_call],
            '1': ['Produtos', 'images/produto.png', produtos_call],
            '2': ['Gestão Produtos', 'images/gestaoprod.png', gp_call],
            '3': ['Recebimento', 'images/recebimento.gif', compras_call],
            '4': ['Produção', 'images/producao.png', producao_call],
            '5': ['Consumo', 'images/consumo.png', consumo_call],
            '6': ['Lista de compras', 'images/compras.png', lista_call],
            '7': ['Venda de balcão', 'images/vendas.png', vendas_call],
            '8': ['Venda personalizada', 'images/money.png', vendas2_call],
            '9': ['Contas à Receber', 'images/salary.png', movin_call],
            '10': ['Contas à Pagar', 'images/pay.png', movout_call],
            '11': ['Transferências', 'images/transference.png', transf_call],
            '12': ['Fluxo de Caixa', 'images/calendar.png', fc_call],
            '13': ['Saldo em Contas', 'images/cashmachine.png', sb_call],
            '14': ['Despesas Categoria', 'images/pag-categorias.png', sc_call],
            '15': ['Despesas Parceiro', 'images/pag-categorias.png', sc_call],
        }
        border = tk.Canvas(main.root, width=960, height=130, bg = 'white')
        border.grid(row=0, column=0, rowspan=2, columnspan=7)
        border.create_rectangle(1, 1, 959, 129)
        widgets.image('', 400, 'images/finance.png', 2, 0, colspam=7, bg='orange', cmd=sobre)
        c.execute('SELECT Dash FROM Usuarios_acessos WHERE Id = ' + str(userauth[0]))
        itensview = c.fetchone()
        if itensview[0]:
            if itensview[0] != '-1':
                itensview = itensview[0].split(',')
                col = 0
                for row in itensview:
                    widgets.label(opcs_dic[row][0], 17, 1, col, fg='white')
                    widgets.image('', 70, opcs_dic[row][1], 0, col, bg='white', imagewidth=(64, 64), cmd=opcs_dic[row][2])
                    col += 1
        widgets.image('', 70, 'images/exit.png', 2, 6, bg='orange', imagewidth=(64, 64), cmd=quit)
    else:
        border = tk.Canvas(main.root, width=960, height=130, bg = 'white')
        border.grid(row=0, column=0, rowspan=2, columnspan=7)
        border.create_rectangle(1, 1, 959, 129)
        widgets.image('', 400, 'images/finance.png', 2, 0, colspam=7, bg='orange', cmd=sobre)
        widgets.label('Parceiros', 14, 1, 0, fg='white')
        widgets.label('Entradas', 14, 1, 1, fg='white')
        widgets.label('Saídas', 14, 1, 2, fg='white')
        widgets.label('Transferências', 14, 1, 3, fg='white')
        widgets.label('Fluxo de Caixa', 14, 1, 4, fg='white')
        widgets.label('Saldo em bancos', 14, 1, 5, fg='white')
        widgets.label('Despesas', 14, 1, 6, fg='white')
        resp = widgets.image('', 70, 'images/users.png', 0, 0, bg='white', imagewidth=(64, 64), cmd=parceiros_call)
        widgets.image('', 70, 'images/salary.png', 0, 1, bg='white', imagewidth=(64, 64), cmd=movin_call)
        widgets.image('', 70, 'images/buy.png', 0, 2, bg='white', imagewidth=(64, 64), cmd=movout_call)
        widgets.image('', 70, 'images/transference.png', 0, 3, bg='white', imagewidth=(64, 64), cmd=transf_call)
        widgets.image('', 70, 'images/calendar.png', 0, 4, bg='white', imagewidth=(64, 64), cmd=fc_call)
        widgets.image('', 70, 'images/cashmachine.png', 0, 5, bg='white', imagewidth=(64, 64), cmd=sb_call)
        widgets.image('', 70, 'images/pag-categorias.png', 0, 6, bg='white', imagewidth=(64, 64), cmd=sc_call)
        widgets.image('', 70, 'images/exit.png', 2, 6, bg='orange', imagewidth=(64, 64), cmd=quit)

def setup_dash():
    def execute():
        _id = str(userauth[0])
        itensview = []
        for row in labels:
            if row.get():
                itensview.append(str(opcs_dic[row.get()]))
        if not itensview:
            itensview.append('-1')
        command = []
        command.append("UPDATE Usuarios_Acessos SET ")
        command.append("Dash = '" + str.join(',', itensview) + "' ")
        command.append("WHERE Id = " + _id)
        c.execute(str.join('', command))
        conn.commit()
        messagebox.showwarning('Aviso', 'Reinicie o programa para que as ações sejam aplicadas.')
        form_selec.destroy()

    if userauth:
        dimension = widgets.geometry(290, 340)
        config = {'title': 'Configurações globais do sistema',
                'dimension': dimension,
                'color': 'light goldenrod'}
        form_selec = main.form(config)
        wg = Widgets(form_selec, 'light goldenrod')
        wg.label('\n' * 13, 5, 0, 0, rowspan=9)
        wg.label('', 5, 0, 1)
        wg.label('', 5, 2, 0)
        wg.label('', 5, 10, 0)
        wg.label('Defina seus ícones no Dashboard:', 30, 1, 1, 2)
        labels = []
        opcs_dic = {
            'Parceiros': 0,
            'Produtos': 1,
            'Gestão de Produtos': 2,
            'Recebimento': 3,
            'Produção': 4,
            'Consumo': 5,
            'Lista de compras': 6,
            'Venda de balcão': 7,
            'Venda personalizada': 8,
            'Contas à Receber': 9,
            'Contas à Pagar': 10,
            'Transferências': 11,
            'Fluxo de Caixa': 12,
            'Saldo em Contas': 13,
            'Despesas por Categoria': 14,
            'Despesas por Parceiro': 15,
        }
        cad = [
            'Parceiros'
            'Produtos',
        ]
        fin = [
            'Contas à Receber',
            'Contas à Pagar',
            'Transferências',
            'Fluxo de Caixa',
            'Saldo em Contas',
            'Despesas por Categoria',
            'Despesas por Parceiro',
        ]        
        est1 = [
            'Gestão de Produtos',
            'Recebimento',
            'Produção',
            'Consumo',
            'Lista de compras',
        ]
        est2 = [
            'Gestão de Produtos',
            'Lista de compras',
        ]
        est3 = [
            'Gestão de Produtos',
        ]
        ven = [
            'Venda de balcão',
            'Venda personalizada',
        ]
        opcs2 = [
            'Parceiros',
            'Produtos',
            'Gestão de Produtos',
            'Recebimento',
            'Produção',
            'Consumo',
            'Lista de compras',
            'Venda de balcão',
            'Venda personalizada',
            'Contas à Receber',
            'Contas à Pagar',
            'Transferências',
            'Fluxo de Caixa',
            'Saldo em Contas',
            'Despesas por Categoria',
            'Despesas por Parceiro',
        ]
        if not userauth[3]:
            opcs = []
            if userauth[4]:
                opcs += fin
            if userauth[5] == 1:
                opcs += est3
            elif userauth[5] == 2:
                opcs += est2
            elif userauth[5] > 2:
                opcs += est1
            if userauth[6]:
                opcs += ven
        else:
            opcs = opcs2
        itensview = []
        num = 0
        c.execute('SELECT Dash FROM Usuarios_acessos WHERE Id = ' + str(userauth[0]))
        resp = c.fetchone()
        if resp:
            itensview = resp[0].split(',')
            if itensview[0] != '-1':
                num = len(itensview)
        for row in range(7):
            if row < num and itensview[0] != '-1':
                texto = opcs2[int(itensview[row])]
            else:
                texto = ''
            labels.append(wg.combobox('Atalho ' + str(row + 1) + ': ', 20, opcs, row + 3, 1, texto))
        wg.button('Confirmar', execute, 15, 2, 11, 1, 2)
    else:
        messagebox.showerror(title='Aviso', message='O sistema não possui usuários. Contate o administrador!')

def setup_cfa():
    def fillusuario(valuereturn):
        if _usuario.get():
            _financ.configure(state='normal')
            _estoque.configure(state='normal')
            _vendas.configure(state='normal')
            value = ''
            c.execute('SELECT Nome, Ativo, Adm, Financ, Estoque, Vendas FROM Usuarios JOIN Usuarios_Acessos ON Usuarios_Acessos.Id = Usuarios.Id WHERE Nome LIKE "%' + _usuario.get() + '%"')
            resp = c.fetchone()
            _usuario.delete(0, 'end')
            _financ.delete(0, 'end')
            _estoque.delete(0, 'end')
            _vendas.delete(0, 'end')
            if resp:
                value = resp[0]
                _usuario.insert(0, value)
                if resp[1]:
                    _ativo.set(1)
                else:
                    _ativo.set(0)
                if resp[2]:
                    _adm.set(1)
                else:
                    _adm.set(0)
                _financ.insert(0, lista1[resp[3]])
                _estoque.insert(0, lista1[resp[4]])
                _vendas.insert(0, lista1[resp[5]])
                if resp[2]:
                    _financ.configure(state='disabled')
                    _estoque.configure(state='disabled')
                    _vendas.configure(state='disabled')
                else:
                    _financ.configure(state='normal')
                    _estoque.configure(state='normal')
                    _vendas.configure(state='normal')

    def parceiro_cmd(value=None):
        widgets.combobox_return(parceiro_sel, lista1)

    def fillacesso(valuereturn):
        if _financ.get() or _estoque.get() or _vendas.get():
            rows = [0, 0, 0]
            widgets.combobox_return(_financ, lista1)
            widgets.combobox_return(_estoque, lista1)
            widgets.combobox_return(_vendas, lista1)

    def habilita(valuereturn):
        if _adm.get():
            _financ.configure(state='normal')
            _estoque.configure(state='normal')
            _vendas.configure(state='normal')
            _financ.delete(0, 'end')
            _estoque.delete(0, 'end')
            _vendas.delete(0, 'end')
            _financ.insert(0, lista1[4])
            _estoque.insert(0, lista1[4])
            _vendas.insert(0, lista1[4])
            _financ.configure(state='disabled')
            _estoque.configure(state='disabled')
            _vendas.configure(state='disabled')
        else:
            _financ.configure(state='normal')
            _estoque.configure(state='normal')
            _vendas.configure(state='normal')
            _financ.delete(0, 'end')
            _estoque.delete(0, 'end')
            _vendas.delete(0, 'end')
            _financ.insert(0, lista1[0])
            _estoque.insert(0, lista1[0])
            _vendas.insert(0, lista1[0])
                
    def execute():
        if _usuario.get():
            c.execute('SELECT Id FROM Usuarios WHERE Nome = "' + _usuario.get() + '"')
            _id = str(c.fetchone()[0])
            command = []
            command.append("UPDATE Usuarios_Acessos SET ")
            command.append("Ativo = " + str(_ativo.get()) + ", ")
            command.append("Adm = " + str(_adm.get()) + ", ")
            command.append("Financ = " + str(lista1.index(_financ.get())) + ", ")
            command.append("Estoque = " + str(lista1.index(_estoque.get())) + ", ")
            command.append("Vendas = " + str(lista1.index(_vendas.get())) + " ")
            command.append("WHERE Id = " + _id)
            c.execute(str.join('', command))
            conn.commit()
            form_selec.destroy()

    dimension = widgets.geometry(280, 380)
    config = {'title': 'Configurações globais do sistema',
              'dimension': dimension,
              'color': 'light goldenrod'}
    form_selec = main.form(config)
    wg = Widgets(form_selec, 'light goldenrod')
    wg.label('\n' * 11, 5, 0, 0, rowspan=8)
    wg.label('', 5, 0, 1)
    wg.label('', 5, 2, 0)
    wg.label('', 5, 8, 0)
    c.execute('SELECT Nome FROM Usuarios ORDER BY Nome')
    lista = []
    for row in c.fetchall():
        lista.append(row[0])
    _usuario = wg.combobox('Usuário: ', 20, lista, 1, 1, cmd=fillusuario, seek=fillusuario)
    _ativo = wg.check('Ativo: ', 15, 'Sim', 3, 1)
    _adm = wg.check('Administrador: ', 15, 'Sim', 4, 1, seek=habilita)
    lista1 = ['Sem acesso', 'Somente Consulta', 'Incluir', 'Edição básica', 'Edição completa']
    _financ = wg.combobox('Finanças: ', 15, lista1, 5, 1, cmd=fillacesso, seek=fillacesso)
    _estoque = wg.combobox('Compras e Estoque: ', 15, lista1, 6, 1, cmd=fillacesso, seek=fillacesso)
    _vendas = wg.combobox('Vendas: ', 15, lista1, 7, 1, cmd=fillacesso, seek=fillacesso)
    Create = wg.button('Atualizar', execute, 10, 0, 10, 1, 2)
    _usuario.focus()

def notimplemented():
    messagebox.showinfo(title='Aviso', message='Módulo não implementado!')

def example():
    def print_tela():
        print(nome.get())
        print(endereco.get())
        print(telefone.get())
        print(cep.get())

    config = {'title': 'Edição de publicadores',
              'columns': 9,
              #'connections': [conn, c],
              'dimension': '770x262+168+202',
              'focus': 0,
              #'fields_nb': nb,
              'file': 'Publicadores',
              #'fields_list': pb_col,
              'post_signal': ''}
    teste = main.form(config)
    wg = Widgets(teste)
    dimdim = wg.image('Imagem', 0, 'images/image.png', 0, 0, 1, 6)
    intro = wg.label('Digite os dados abaixo:', 20, 0, 1, 2)
    nome = wg.textbox('Nome', 20, 1, 1, 'Renan')
    endereco = wg.textbox('Endereço', 30, 2, 1)
    telefone = wg.textbox('Telefone', 15, 3, 1)
    cep = wg.textbox('CEP', 7, 4, 1)
    estado = wg.combobox('Estado', 5, ['MG', 'SP'], 5, 1, 'SP')
    lista = ['Sim', 'Não']
    status = wg.listbox('Condição', 20, 4, lista, 6, 1)
    valor = wg.textbox('Valor', 10, 7, 1)
    nome.focus()
    resp = tk.Button(teste, width=20, text="Imprimir", command=print_tela)
    resp.grid(row=8, column=1, columnspan=1)

def appquit():
    c.execute('SELECT Id FROM Drive')
    if c.fetchone():
        export_data()
    main.destroy()

def auth():
    def valid():
        c.execute('SELECT Senha, Usuarios.Id, Nome, Ativo, Adm, Financ, Estoque, Vendas FROM Usuarios JOIN Usuarios_Acessos ON Usuarios_Acessos.Id = Usuarios.Id WHERE Nome = "' + _usuario.get().upper() + '"')
        #c.execute('SELECT Senha, Id FROM Usuarios WHERE Nome = "' + _usuario.get().upper() + '"')
        resp = c.fetchone()
        if resp:
            if resp[0] == _senha.get() and resp[3]:
                AUTHENTICATE[0] = True
                main.destroy()
                userauth.append(resp[1])
                userauth.append(resp[2])
                userauth.append(resp[3])
                userauth.append(resp[4])
                userauth.append(resp[5])
                userauth.append(resp[6])
                userauth.append(resp[7])
        else:
            if _usuario.get() == 'admin' and _senha.get() == 'root':
                AUTHENTICATE[0] = True
                main.destroy()
        if not AUTHENTICATE[0]:
            messagebox.showerror(title='Aviso', message='Usuário desconhecido ou senha incorreta')

    main = tk.Tk()
    main.title('Login')
    main.geometry('295x215+400+200')
    main['bg'] = 'orange'
    uwg = Widgets(main, 'orange')
    uwg.label('', 5, 0, 0)
    uwg.label('', 5, 1, 0, rowspan=5, height=10)
    uwg.label('', 5, 3, 0)
    uwg.label('', 5, 6, 0)
    _usuario = uwg.textbox('Usuário: ', 20, 1, 1)
    _senha = uwg.textbox('Senha: ', 20, 2, 1, show='*')
    uwg.button('Autenticar', valid, 20, 1, 4, 1, colspan=2)
    uwg.button('Sair', main.destroy, 20, 1, 5, 1, colspan=2)
    _usuario.focus()
    main.mainloop()

def recharge(rid=0, antecipado=False):
    def valid():
        if tentativa[0] >= 3:
            if not antecipado:
                validade[0] = False
            main.destroy()
        else:
            new_key = int(_usuario.get())
            check = validator(new_key, str(_datetime[0]), _senha.get())
            if check[0]:
                new_value = str(_datetime[0]) + ',' + str(int(check[1])) + ',' + _senha.get()
                new_value = new_value.encode('utf-8')
                new_value = str(base64.b64encode(new_value))[2:-1]
                c.execute('UPDATE Auth SET Id = ' + str(new_key) + ', Data = "' + new_value + '"')
                conn.commit()
                validade[0] = True
                smain.destroy()
                export_data()
            else:
                messagebox.showerror(title='Aviso', message='Código digitado não habilita o uso do aplicativo')
                tentativa[0] += 1
                if not antecipado:
                    validade[0] = False

    smain = tk.Tk()
    smain.geometry('565x285+400+200')
    smain.title('Renovação de licença')
    smain['bg'] = 'orange'
    uwg = Widgets(smain, 'orange')
    uwg.label('', 5, 0, 0)
    uwg.label('', 5, 1, 0, rowspan=6, height=12)
    uwg.label('', 5, 3, 0)
    uwg.label('', 5, 5, 0)
    if not antecipado:
        mensagem = 'Atenção! Seu acesso ao CFA expirou. Contate nosso suporte técnico:'\
        '\nEmail: renanhs.80@gmail.com\nOu fale com o representante de sua região.'
    else:
        mensagem = 'Insira o código recebido para proceder com a renovação.'
    uwg.label(mensagem, 60, 1, 1, colspan=3)
    _usuario = uwg.textbox('Id: ', 20, 3, 1, default=str(rid))
    _senha = uwg.textbox('Chave: ', 20, 4, 1)
    uwg.button('Autenticar', valid, 20, 1, 6, 1, colspan=2)
    uwg.button('Sair', smain.destroy, 20, 1, 7, 1, colspan=2)
    _usuario.configure(state='disabled')
    _senha.focus()
    tentativa = [0]
    if not antecipado:
        _datetime = [datetime.now().date()]
    else:
        _datetime = [antecipado]
    validade = [False]
    smain.mainloop()
    if validade[0]:
        return True
    if not validade[0]:
        return False

system_init = datetime.now()
param = {
    'title': 'Consultor Financeiro e Administrativo',
    'backcolor': 'orange',
    'geometry': '960x600+30+30',
}

menus = [
    # menus
    {'title': 'Cadastros'},
    {'title': 'Finanças'},
    {'title': 'Estoque'},
    {'title': 'Vendas'},
    {'title': 'Consultas'},
    {'title': 'Backup'},
    # submenus
    {'title': 'Financeiro'},
    {'title': 'Bancos'},
    {'title': 'Categorias'},
    {'title': 'Parceiros'},
    {'title': 'Materiais'},
    {'title': 'Backup'},
    {'title': 'Sistema'},
]

opc_cad1 = [
    {'title': 'Cadastros', 'menu': 0},
    {'root': 0, 'title': 'Sistema', 'menu': 12},
    {'root': 12, 'title': 'Usuários', 'command': usuarios},
    {'root': 12, 'title': 'Nivel de acesso', 'command': setup_cfa},
    {'root': 12, 'title': 'Dashboard', 'command': setup_dash},
    {'root': 0, 'title': 'Financeiro', 'menu': 6},
    {'root': 6, 'title': 'Bancos', 'menu': 7},
    {'root': 7, 'title': 'Editar contas de movimentação', 'command': bancos_mov},
    {'root': 7, 'title': 'Editar cartões de crédito', 'command': bancos_cc},
    {'root': 7, 'title': 'Editar cartões pré pago', 'command': bancos_cpp},
    {'root': 6, 'title': 'Categorias', 'menu': 8},
    {'root': 8, 'title': 'Editar receitas', 'command': categorias_receitas},
    {'root': 8, 'title': 'Editar despesas', 'command': categorias_despesas},
    {'root': 6, 'title': 'Parceiros', 'menu': 9},
    {'root': 9, 'title': 'Editar', 'command': parceiros},
    {'root': 9, 'title': 'Importar por csv', 'command': notimplemented},
    {'root': 9, 'title': 'Listar', 'command': parceiros_listar},
    {'root': 0, 'title': 'Materiais', 'menu': 10},
    {'root': 10, 'title': 'Categorias', 'command': materiais_categorias},
    {'root': 10, 'title': 'Produtos', 'command': materiais_itens},
    {'root': 0, 'title': 'Sair', 'command': appquit}
]
opc_cad2 = [
    {'title': 'Cadastros', 'menu': 0},
    {'root': 0, 'title': 'Sistema', 'menu': 12},
    {'root': 12, 'title': 'Usuários', 'command': usuarios},
    {'root': 12, 'title': 'Dashboard', 'command': setup_dash},
    {'root': 0, 'title': 'Sair', 'command': appquit}
]
opc_fin1 = [
    {'title': 'Finanças', 'menu': 1},
    {'root': 1, 'title': 'Contas à receber', 'command': movimentos_in},
    {'root': 1, 'title': 'Contas à pagar', 'command': movimentos_out},
    {'root': 1, 'title': 'Cartões de crédito', 'command': movimentos_crd},
    {'root': 1, 'title': 'Transferências internas', 'command': movimentos},
    {'root': 1, 'title': 'Fluxo de caixa', 'command': fluxocaixa},
]
opc_est1 = [
    {'title': 'Estoque', 'menu': 2},
    {'root': 2, 'title': 'Gestão de Produtos', 'command': materiais_consultas},
    {'root': 2, 'title': 'Recebimento de mercadorias', 'command': materiais_movimentos_rec},
    {'root': 2, 'title': 'Produção', 'command': materiais_movimentos_prod},
    {'root': 2, 'title': 'Consumo de materiais', 'command': materiais_movimentos_cons},
    {'root': 2, 'title': 'Lista de compras', 'command': materiais_listacompras},
]
opc_est2 = [
    {'title': 'Estoque', 'menu': 2},
    {'root': 2, 'title': 'Gestão de Produtos', 'command': materiais_consultas},
    {'root': 2, 'title': 'Lista de compras', 'command': materiais_listacompras},
]
opc_est3 = [
    {'title': 'Estoque', 'menu': 2},
    {'root': 2, 'title': 'Gestão de Produtos', 'command': materiais_consultas},
]
opc_ven1 = [
    {'title': 'Vendas', 'menu': 3},
    {'root': 3, 'title': 'Venda ao público - À vista', 'command': materiais_movimentos_vend},
    {'root': 3, 'title': 'Venda personalizada', 'command': materiais_movimentos_vend2},
]
opc_cad3 = [
    {'title': 'Consultas', 'menu': 4},
    {'title': 'Backup', 'menu': 5},
    {'root': 4, 'title': 'Saldo em contas', 'command': saldobancario},
    {'root': 4, 'title': 'Despesas por categoria', 'command': sinteticocategoria},
    {'root': 4, 'title': 'Despesas por parceiro', 'command': sinteticoempresa},
    {'root': 5, 'title': 'Exportar Google Drive', 'command': export_data},
    {'root': 5, 'title': 'Importar Google Drive', 'command': import_data},
]
opcoes = [
    {'title': 'Cadastros', 'menu': 0},
    {'title': 'Finanças', 'menu': 1},
    {'title': 'Estoque', 'menu': 2},
    {'title': 'Vendas', 'menu': 3},
    {'title': 'Consultas', 'menu': 4},
    {'title': 'Backup', 'menu': 5},
    {'root': 0, 'title': 'Sistema', 'menu': 12},
    {'root': 12, 'title': 'Usuários', 'command': usuarios},
    {'root': 12, 'title': 'Nivel de acesso', 'command': setup_cfa},
    {'root': 12, 'title': 'Dashboard', 'command': setup_dash},
    {'root': 0, 'title': 'Financeiro', 'menu': 6},
    {'root': 6, 'title': 'Bancos', 'menu': 7},
    {'root': 7, 'title': 'Editar contas de movimentação', 'command': bancos_mov},
    {'root': 7, 'title': 'Editar cartões de crédito', 'command': bancos_cc},
    {'root': 7, 'title': 'Editar cartões pré pago', 'command': bancos_cpp},
    {'root': 6, 'title': 'Categorias', 'menu': 8},
    {'root': 8, 'title': 'Editar receitas', 'command': categorias_receitas},
    {'root': 8, 'title': 'Editar despesas', 'command': categorias_despesas},
    {'root': 6, 'title': 'Parceiros', 'menu': 9},
    {'root': 9, 'title': 'Editar', 'command': parceiros},
    {'root': 9, 'title': 'Importar por csv', 'command': notimplemented},
    {'root': 9, 'title': 'Listar', 'command': parceiros_listar},
    {'root': 0, 'title': 'Materiais', 'menu': 10},
    {'root': 10, 'title': 'Categorias', 'command': materiais_categorias},
    {'root': 10, 'title': 'Produtos', 'command': materiais_itens},
    {'root': 1, 'title': 'Contas à receber', 'command': movimentos_in},
    {'root': 1, 'title': 'Contas à pagar', 'command': movimentos_out},
    {'root': 1, 'title': 'Cartões de crédito', 'command': movimentos_crd},
    {'root': 1, 'title': 'Transferências internas', 'command': movimentos},
    {'root': 1, 'title': 'Fluxo de caixa', 'command': fluxocaixa},
    {'root': 2, 'title': 'Gestão de Produtos', 'command': materiais_consultas},
    {'root': 2, 'title': 'Recebimento de mercadorias', 'command': materiais_movimentos_rec},
    {'root': 2, 'title': 'Produção', 'command': materiais_movimentos_prod},
    {'root': 2, 'title': 'Consumo de materiais', 'command': materiais_movimentos_cons},
    {'root': 2, 'title': 'Lista de compras', 'command': materiais_listacompras},
    {'root': 3, 'title': 'Venda ao público - À vista', 'command': materiais_movimentos_vend},
    {'root': 3, 'title': 'Venda personalizada', 'command': materiais_movimentos_vend2},
    {'root': 4, 'title': 'Saldo em contas', 'command': saldobancario},
    {'root': 4, 'title': 'Despesas por categoria', 'command': sinteticocategoria},
    {'root': 4, 'title': 'Despesas por parceiro', 'command': sinteticoempresa},
    {'root': 5, 'title': 'Exportar Google Drive', 'command': export_data},
    {'root': 5, 'title': 'Importar Google Drive', 'command': import_data},
    {'root': 0, 'title': 'Sair', 'command': appquit}
]
conn = sqlite3.connect('finance.db')
c = conn.cursor()
userauth = []
c.execute('SELECT Id FROM Usuarios')
if c.fetchall():
    AUTHENTICATE = [False]
    paramaut = {
        'title': 'Autenticação - Controle Financeiro Pessoal',
        'backcolor': 'orange',
        'geometry': '400x300+60+60',
    }
    auth()
else:
    AUTHENTICATE = [True]

if AUTHENTICATE[0]:
    ActualDate = str(datetime.now())[0:10]
    main = Application(param)
    widgets = Widgets(main.root)
    dashboard()
    c.execute('SELECT Id FROM Drive')
    if c.fetchone():
        import_data()
    if userauth:
        if userauth[3]:
            opcoes = opc_cad1 + opc_fin1 + opc_est1 + opc_ven1 + opc_cad3
        else:
            opcoes = opc_cad2
            if userauth[4]:
                opcoes += opc_fin1
            if userauth[5]:
                if userauth[5] == 1:
                    opcoes += opc_est2
                else:                        
                    opcoes += opc_est1
            if userauth[6]:
                opcoes += opc_ven1
    main.menu(menus=menus, opcoes=opcoes)
    main.mainwindow()
    if int((datetime.now() - system_init).total_seconds()) < 60:
        print(str(int((datetime.now() - system_init).total_seconds())) + ' segundos de uso.')
    else:
        print(str(int((datetime.now() - system_init).total_seconds() / 60)) + ' minutos de uso.')
c.close()
conn.close()
