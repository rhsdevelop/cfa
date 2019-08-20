# CFA - CONSULTOR FINANCEIRO E ADMINISTRATIVO
ERP completo para empresas de pequeno porte ou profissionais autônomos.

RHS Desenvolvimentos
Programador: Renan Hernandes de Souza

O CFA contém funções de um ERP. Um sistema completo para controle financeiro, administração de materiais e faturamento.
O CFA foi construído em Python 3 usando a interface gráfica Tkinter. Tem seu código fonte aberto e está disponível para estudantes e/ou profissionais que possuem conhecimentos de programação e desejam ter controle total sobre sua gestão administrativa. Usa nativamente o banco de dados SQLite.


BENEFÍCIOS
- Baixo custo de implantação.
- Múltiplos usuários com acesso personalizável pelo administrador.
- Possibilidade de criar múltiplos caixas, incluindo gestão sobre Cartões de Crédito.
- Classificação de despesas por mês de compra, independente do regime de caixa.
- Integração com o Google Drive para realização de Backup ou para uso em múltiplas máquinas.


INSTALAÇÃO

Sistema operacional:
Windows, Mac e Linux

Passos:
- clonar o repositório do CFA. Comando: git clone https://github.com/rhsdevelop/cfa.git
- instalar Python 3
- instalar pip, se não estiver presente na distribuição.

Opcional:
- instalar virtualenv compatível com Python 3. Comando: virtualenv -p python3 env
- instanciar a pasta do environment criado:
Linux/Mac: source ./finance/bin/activate
Windows: ./finance/Scripts/activate.bat

Primeiro uso:
- instalar os módulos Python necessários para o funcionamento. Comando: pip install -r requirements.txt
- inicializar o banco de dados do primeiro projeto. Comando: python database.py
- executar o CFA. Comando: python finance.py


MANUTENÇÃO DO BANCO DE DADOS

Arquivo de dados
- finance.db (SQLite3)

Uso em múltiplas máquinas
Se o programa estiver integrado ao Google Drive, copiar a pasta ./settings e o arquivo finance.db para o destino na máquina que recebe a réplica do banco.
Ao exportar os dados para a nuvem, todos os computadores utilizados estarão sincronizados na mesma conta Google, se ao iniciarem for aceita a integração automática com os dados no Google Drive.