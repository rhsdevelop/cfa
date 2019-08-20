# -*- coding: utf-8 -*-
'''
Programa teste para implantação de banco de dados para o aplicativo do salão.
'''

import datetime
import sqlite3

def create_tables():
    '''
    Rotina de criação de tabelas.
    Cria novas tabelas, caso não exista.
    '''
    # Autenticação de software
    command = []
    command.append('CREATE TABLE IF NOT EXISTS Auth')
    command.append('(')
    command.append('Id INTEGER, ') # 0 a 999
    command.append('Data BLOB NOT NULL,') # data_efetiva,minutos_saldo,senha
    command.append('Nome VARCHAR(100), ')
    command.append('Empresa VARCHAR(100), ')
    command.append('Email VARCHAR(100), ')
    command.append('Telefone VARCHAR(50)')
    command.append(')')
    c.execute(str.join('', command))

    # Usuarios
    command = []
    command.append('CREATE TABLE IF NOT EXISTS Usuarios')
    command.append('(')
    command.append('Id INTEGER PRIMARY KEY AUTOINCREMENT, ')
    command.append('Nome VARCHAR(50) NOT NULL, ')
    command.append('Senha VARCHAR(50) NOT NULL, ')
    command.append('Observacao VARCHAR(200) NOT NULL')
    command.append(')')
    c.execute(str.join('', command))

    # Nível de acesso
    command = []
    command.append('CREATE TABLE IF NOT EXISTS Usuarios_acessos')
    command.append('(')
    command.append('Id INTEGER PRIMARY KEY AUTOINCREMENT, ') # Mesmo id do usuário
    command.append('Ativo INTEGER NOT NULL, ') # Booleano
    command.append('Adm INTEGER NOT NULL, ') # Booleano
    command.append('Financ INTEGER NOT NULL, ') # lista1 = ['Sem acesso', 'Somente Consulta', 'Incluir', 'Edição básica', 'Edição completa']
    command.append('Estoque INTEGER NOT NULL, ')
    command.append('Vendas INTEGER NOT NULL, ')
    command.append('Dash VARCHAR(18)') # Campos separados por vírgula
    command.append(')')
    c.execute(str.join('', command))

    # Bancos
    command = []
    command.append('CREATE TABLE IF NOT EXISTS Bancos')
    command.append('(')
    command.append('Id INTEGER PRIMARY KEY AUTOINCREMENT, ')
    command.append('DataCadastro VARCHAR(10) NOT NULL, ')
    command.append('NomeBanco VARCHAR(50) NOT NULL, ')
    command.append('TipoMov INTEGER NOT NULL, ') # 0 Dinheiro em mãos  1 Conta bancária  2 Cartão de Crédito  3 Pré pago
    command.append('Numero VARCHAR(19), ') # numero do cartão de crédito
    command.append('DiaVenc INTEGER, ') # pra cartão de crédito (inteiro de 1 a 30)
    command.append('GeraFatura INTEGER, ') # 0 Não e 1 Sim 
    command.append('Agencia VARCHAR(10), ') # pra conta bancária
    command.append('Conta VARCHAR(12), ') # pra conta bancária
    command.append('TipoConta INTEGER') # pra conta bancária: 0 Conta Corrente  1 Poupança  2 CDB/Outras
    command.append(')')
    c.execute(str.join('', command))

    # Categorias
    command = []
    command.append('CREATE TABLE IF NOT EXISTS Categorias')
    command.append('(')
    command.append('Id INTEGER PRIMARY KEY AUTOINCREMENT, ')
    command.append('Categoria VARCHAR(50) NOT NULL, ')
    command.append('TipoMov INTEGER NOT NULL, ') # 0 para Receita e 1 para Despesa
    command.append('Repete INTEGER NOT NULL') # 0 para False e 1 para True
    command.append(')')
    c.execute(str.join('', command))

    # Parceiros
    command = []
    command.append('CREATE TABLE IF NOT EXISTS Parceiros')
    command.append('(')
    command.append('Id INTEGER PRIMARY KEY AUTOINCREMENT, ')
    command.append('DataCadastro VARCHAR(10) NOT NULL, ')
    command.append('Nome VARCHAR(30) NOT NULL, ')
    command.append('NomeCompleto VARCHAR(100) NOT NULL, ')
    command.append('Tipo INTEGER NOT NULL, ') # 0 para Pessoa Física e 1 para Pessoa Jurídica
    command.append('Doc VARCHAR(30), ')
    command.append('Endereco VARCHAR(100), ')
    command.append('Telefone VARCHAR(20), ')
    command.append('Observacao VARCHAR(200), ')
    command.append('Modo INTEGER') # 0 Ambos     1 Cliente      2 Fornecedor
    command.append(')')
    c.execute(str.join('', command))

    # Movimento bancário
    command = []
    command.append('CREATE TABLE IF NOT EXISTS Diario')
    command.append('(')
    command.append('Id INTEGER PRIMARY KEY AUTOINCREMENT, ')
    command.append('DataFirstUpdate VARCHAR(10), ')
    command.append('DataLastUpdate VARCHAR(10), ')
    command.append('DataDoc VARCHAR(10), ')
    command.append('DataVenc VARCHAR(10), ')
    command.append('DataPago VARCHAR(10), ')
    command.append('Parceiro INTEGER NOT NULL, ')
    command.append('Banco INTEGER NOT NULL, ')
    command.append('Fatura VARCHAR(20), ') # Se o pagamento for cartão de crédito, ano/mes. Sugere e usuario pode mudar
    command.append('Descricao VARCHAR(100) NOT NULL, ')
    command.append('Valor REAL NOT NULL, ')
    command.append('TipoMov INTEGER NOT NULL, ') # 0 para Receita e 1 para Despesa 3 Transf Ent 4 Transf Sai
    command.append('CategoriaMov VARCHAR(10) NOT NULL') # Categoria de Receita ou Despesa 
    command.append(')')
    c.execute(str.join('', command))

    # Arquivos exportados para o Google Drive
    command = []
    command.append('CREATE TABLE IF NOT EXISTS Drive')
    command.append('(')
    command.append('Id INTEGER PRIMARY KEY AUTOINCREMENT, ')
    command.append('Arquivo VARCHAR(30) NOT NULL, ')
    command.append('Id_Google VARCHAR(100) NOT NULL')
    command.append(')')
    c.execute(str.join('', command))

    # Materiais - Cadastro de produtos
    command = []
    command.append('CREATE TABLE IF NOT EXISTS Materiais_Itens')
    command.append('(')
    command.append('Id INTEGER PRIMARY KEY AUTOINCREMENT, ')
    command.append('DataCadastro VARCHAR(19) NOT NULL, ')
    command.append('DataAtualiza VARCHAR(19) NOT NULL, ')
    command.append('Produto VARCHAR(60) NOT NULL, ')
    command.append('Categoria INTEGER NOT NULL, ')
    command.append('Tipo INTEGER NOT NULL, ')  # [0] = 'Compra' [1] = 'Venda' [2] = 'Revenda'
    command.append('Unidade VARCHAR(30) NOT NULL, ')
    command.append('Marca VARCHAR(60), ')
    command.append('ValorCusto REAL NOT NULL, ')
    command.append('Margem REAL NOT NULL, ')
    command.append('ValorVenda REAL NOT NULL, ')
    command.append('EstoqueMin REAL NOT NULL')
    command.append(')')
    c.execute(str.join('', command))

    # Materiais - Cadastro de categorias
    command = []
    command.append('CREATE TABLE IF NOT EXISTS Materiais_Categorias')
    command.append('(')
    command.append('Id INTEGER PRIMARY KEY AUTOINCREMENT, ')
    command.append('DataCadastro VARCHAR(19) NOT NULL, ')
    command.append('DataAtualiza VARCHAR(19) NOT NULL, ')
    command.append('Categoria VARCHAR(60) NOT NULL, ')
    command.append('Tipo INTEGER NOT NULL')  # [0] = 'Compra' [1] = 'Venda' [2] = 'Revenda'
    command.append(')')
    c.execute(str.join('', command))

    # Materiais - Movimentação
    command = []
    command.append('CREATE TABLE IF NOT EXISTS Materiais_Movimentos')
    command.append('(')
    command.append('Id INTEGER PRIMARY KEY AUTOINCREMENT, ')
    command.append('DataCadastro VARCHAR(19) NOT NULL, ')
    command.append('DataAtualiza VARCHAR(19) NOT NULL, ')
    command.append('DataEfetiva VARCHAR(10) NOT NULL, ')
    command.append('Produto INTEGER NOT NULL, ')
    command.append('VlUnit REAL NOT NULL, ')
    command.append('Qtde REAL NOT NULL, ')
    command.append('VlCusto REAL NOT NULL, ')
    command.append('VlVenda REAL NOT NULL, ')
    command.append('Documento INTEGER NOT NULL')
    command.append(')')
    c.execute(str.join('', command))

    # Materiais - Documentos
    command = []
    command.append('CREATE TABLE IF NOT EXISTS Materiais_Documentos')
    command.append('(')
    command.append('Id INTEGER PRIMARY KEY AUTOINCREMENT, ')
    command.append('DataCadastro VARCHAR(19) NOT NULL, ')
    command.append('DataAtualiza VARCHAR(19) NOT NULL, ')
    command.append('Tipo INTEGER NOT NULL, ')  # [0] = 'Recebimento' [1] = 'Produção' [2] = 'Consumo' [3] = 'Venda'
    command.append('MovimentoFin INTEGER, ') # Chave estrangeira
    command.append('Doc VARCHAR(20), ')
    command.append('DataDoc VARCHAR(10), ')
    command.append('DataEfetiva VARCHAR(10), ')
    command.append('TotalCusto REAL NOT NULL, ')
    command.append('TotalVenda REAL NOT NULL, ')
    command.append('Parceiro INTEGER NOT NULL, ') # Chave estrangeira
    command.append('Financeiro INTEGER NOT NULL, ') # Booleano
    command.append('VlDesconto REAL NOT NULL, ')
    command.append('VlTotal REAL NOT NULL, ')
    command.append('Descricao VARCHAR(200), ')
    command.append('Usuario INTEGER NOT NULL')
    command.append(')')
    c.execute(str.join('', command))


def inserir_dados():
    '''
    Rotina de inserção de registros.
    Incluir novos registros.
    '''
    # Inserir categorias padrão
    insert_data = [
        '(1, "ADMINISTRATIVO", 1, 1)',
        '(2, "AGUA E LUZ", 1, 1)',
        '(3, "CARRO COMBUSTIVEL", 1, 1)',
        '(4, "CARRO MANUTENCAO", 1, 1)',
        '(5, "CLUBE E RECREAÇÃO", 1, 1)',
        '(6, "COMPRAS MP", 1, 1)',
        '(7, "COMPRAS REVENDA", 1, 1)',
        '(8, "CONSORCIO", 1, 1)',
        '(9, "DONATIVOS", 1, 1)',
        '(10, "EMPRESTIMOS", 1, 0)',
        '(11, "EQUIPAMENTOS", 1, 1)',
        '(12, "HABITAÇÃO ALUGUEL", 1, 1)',
        '(13, "HABITAÇÃO PRESTAÇÃO", 1, 1)',
        '(14, "IMPOSTO IPTU", 1, 1)',
        '(15, "IMPOSTO IPVA", 1, 1)',
        '(16, "IMPOSTOS", 1, 1)',
        '(17, "INVESTIMENTOS", 1, 1)',
        '(18, "RESTAURANTE", 1, 1)',
        '(19, "SAUDE", 1, 1)',
        '(20, "SEGUROS", 1, 1)',
        '(21, "VENDA DIRETA", 0, 0)',
        '(22, "TAXAS E JUROS", 1, 1)',
        '(23, "TELEFONE E INTERNET", 1, 1)',
        '(24, "VESTUARIO", 1, 1)',
        '(25, "VIAGENS", 1, 1)',
        '(26, "TRANSFERENCIAS", 2, 0)',
        '(27, "SUPERMERCADO", 1, 1)',
        '(28, "SERVIÇOS", 0, 0)',
        '(29, "RENDIMENTOS", 0, 0)',
        '(30, "OUTRAS ENTRADAS", 0, 0)',
        '(31, "SALÁRIO-PRO LABORE", 0, 0)',
    ]
    for row in insert_data:
        c.execute('INSERT INTO Categorias VALUES ' + row)

    # Inserir parceiro genérico
    row = ('(1, "' + str(datetime.datetime.now().date()) + '", "MOVIMENTOS GERAIS", "MOVIMENTOS SEM PARCEIRO CADASTRADO", ' +
           '0, "99.999.999/9999-99", "Sem localização", "", "Sem controle", 0)')
    c.execute('INSERT INTO Parceiros VALUES ' + row)

    # Inserir conta de movimentação Dinheiro em mãos
    row = ('(1, "' + str(datetime.datetime.now().date()) + '", "DINHEIRO EM MÃOS", 0, "", 0, 0, "", "", 0)')
    c.execute('INSERT INTO Bancos VALUES ' + row)
    
    conn.commit()

def atualiza_2208():
    command = ('SELECT Id, Valor FROM Diario WHERE TipoMov IN (1, 4)')
    # command = ('SELECT Id, Valor FROM Diario')
    c.execute(command)
    resp = c.fetchall()
    dados = {}
    for row in resp:
        dados[row[0]] = row[1]
        command = ('UPDATE Diario SET Valor = -' + str(row[1]) + ' WHERE Id = ' + str(row[0]))
        c.execute(command)
    conn.commit()
    print(dados)


conn = sqlite3.connect('finance.db')
c = conn.cursor()
create_tables()
inserir_dados()
# atualiza_2208()
c.close()
conn.close()