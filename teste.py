from designInicial import Ui_MainWindow
import sys
from PyQt5.QtWidgets import QMainWindow, QApplication, QMessageBox
from PyQt5 import QtWidgets

import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

import pymysql
import pymysql.cursors

from openpyxl import load_workbook
from openpyxl import Workbook

from datetime import datetime

from os import listdir, remove

import sys
import traceback

# \/VARIAVEIS ABAIXO CRIADA PARA SEREM GLOBAIS USADO NA FUNÇÃO
##############################################
# atualizar_tabelinha_ultimos_itens_criados
mainListaTabelinha = []

# on_selectionChanged
linha_selecionada = []
coluna_selecionada = []
itemSelecionadoNome = ''

contadorItensAdicionadosParaSair = 0

# atualizar_tabelinha_cardapio
mainTabelinhaCardapio = []
##############################################

class MyMainWindow(QMainWindow, Ui_MainWindow):
    def __init__(self, parent=None):
        super(MyMainWindow, self).__init__(parent)
        self.setupUi(self)

        # RETORNO DE ERRO
        ##############################################
        sys._excepthook = sys.excepthook

        def exception_hook(exctype, value, traceback):
            print(exctype, value, traceback)
            sys._excepthook(exctype, value, traceback)

        sys.excepthook = exception_hook

        # PÁGINA INICIAL
        ##############################################
        self.mainStackedWidget.setCurrentWidget(self.mainPage)

        # PÁGINAS

        ##############################################
        # PÁGINA MERCADORIAS
        ##############################################

        # IR PARA A PÁGINA DE MERCADORIAS / CONTROLE DO BANCO DE DADOS
        self.mercadorias.clicked.connect(lambda: self.mainStackedWidget.setCurrentWidget(self.mercadoriasPage))

        # VOLTAR PARA DA PÁGINA MERCADORIAS PARA A INICIAL
        self.voltarDeMercadorias.clicked.connect(lambda: self.mainStackedWidget.setCurrentWidget(self.mainPage))

        ## PÁGINA ADICIONAR ITENS AO ESTOQUE E CARDAPIO
        # FRAME ADCIONAR ITENS AO ESTOQUE E CARDAPIO
        self.btn1adicionarItem.clicked.connect(
            lambda: self.mercadoriasStackedWidget.setCurrentWidget(self.adicionarItem))

        # FINALIZAR CRIAÇÃO DE ITEM NO ESTOQUE
        self.unidadeComboBox.setChecked(True)
        self.btnFinalizarCriarItemEstoque.clicked.connect(self.finalizar_de_criar_item_estoque)

        # ATUALIZAR TABELINHA DE ULTIMOS ITENS CRIADOS NO ESTOQUE
        self.btnFinalizarCriarItemEstoque.clicked.connect(self.atualizar_tabelinha_ultimos_itens_criados)

        # ZERAR CAIXAS DE TEXTO DEPOIS DE CRIAR ITEM NO ESTOQUE
        self.btnFinalizarCriarItemEstoque.clicked.connect(self.zerar_caixas_de_texto_criar_item_estoque)

        # ADICIONAR ITENS AO CARDAPIO (DESABILITANDO E ABILITANDO CAMPOS)
        self.comboBoxMedida1.currentIndexChanged.connect(self.habilitar_ingrediente_2)
        self.comboBoxMedida2.currentIndexChanged.connect(self.habilitar_ingrediente_3)
        self.comboBoxMedida3.currentIndexChanged.connect(self.habilitar_ingrediente_4)
        self.comboBoxMedida4.currentIndexChanged.connect(self.habilitar_ingrediente_5)

        self.nomeItemGasto2.setDisabled(True)
        self.quantidadeItemCriado2.setDisabled(True)
        self.comboBoxMedida2.setDisabled(True)

        self.nomeItemGasto3.setDisabled(True)
        self.quantidadeItemCriado3.setDisabled(True)
        self.comboBoxMedida3.setDisabled(True)

        self.nomeItemGasto4.setDisabled(True)
        self.quantidadeItemCriado4.setDisabled(True)
        self.comboBoxMedida4.setDisabled(True)

        self.nomeItemGasto5.setDisabled(True)
        self.quantidadeItemCriado5.setDisabled(True)
        self.comboBoxMedida5.setDisabled(True)

        # FINALIZAR DE CRIAR ITENS NO CARDAPIO
        self.btnFinalizarCriarItemCardapio.clicked.connect(self.finalizar_criar_item_cardapio)

        # ATUZALIZAR TABELINHA DO CARDAPIO
        self.btnFinalizarCriarItemCardapio.clicked.connect(self.atualizar_tabelinha_cardapio)

        # LIMPAR CAIXAS DE TEXTO DO CRIAR ITEM CARDAPIO
        self.btnFinalizarCriarItemCardapio.clicked.connect(self.zerar_caixas_de_texto_criar_item_cardapio)

        # BOTÕES PARA VER TABELA ESTOQUE(MERCADORIAS) E CARDÁPIO
        self.btnVerEstoque.clicked.connect(lambda: self.mercadoriasStackedWidget.setCurrentWidget(self.verTabelaMercadorias))
        self.btnVerCardapio.clicked.connect(lambda: self.mercadoriasStackedWidget.setCurrentWidget(self.verTabelaCardapio))

        ## TABELA ESTOQUE
        # FRAME ACESSAR TEBELA MERCADORIAS
        self.btn2acessarTabelaMercadorias.clicked.connect(
            lambda: self.mercadoriasStackedWidget.setCurrentWidget(self.verTabelaMercadorias))
        self.btnVoltarSecreto.clicked.connect(
            lambda: self.mercadoriasStackedWidget.setCurrentWidget(self.verTabelaMercadorias))

        # ATUALIZAR TABELA / CARREGAR TODOS OS DADOS
        self.btnAtualizarTabelaMercadorias.clicked.connect(self.listar_dados)

        # PESQUISAR ITEM NA TABELA
        self.btnPesquisarTabelaMercadorias.clicked.connect(self.pesquisar_item_tabela_mercadoria)

        # FECHAR E MINIMIZAR A TELA
        self.fechar.clicked.connect(self.close)
        self.minimizar.clicked.connect(self.showMinimized)

        # ADICIONAR OU ALTERAR NA TABELA ESTOQUE
        self.btnAdicionarQuantidadeItemEstoque.clicked.connect(self.adicionar_quantia_ao_estoque)
        self.btnAlterarQuantidadeItemEstoque.clicked.connect(self.alterar_quantia_do_estoque)

        ## TABELA CARDAPIO
        # ACESSAR TABELA CARDAPIO
        self.btn3acessarCardapio.clicked.connect(lambda: self.mercadoriasStackedWidget.setCurrentWidget(self.verTabelaCardapio))

        # LISTAR TEBAL CARDAPRIO
        self.btnAtualizarTabelaCardapio.clicked.connect(self.listar_cardapio)

        ##############################################
        # PÁGINA RESGISTRAR SAIDA
        ##############################################
        # ACESSARREGISTRAR SAIDA
        self.registrarSaidaBtn.clicked.connect(lambda: self.mainStackedWidget.setCurrentWidget(self.registrarSaida))

        # ATUALIZAR TABELA NA ABA DE REGISTRAR SAIDA
        self.btnAtualizarTabelaCardapioParaSair.clicked.connect(self.listar_itens_cardapio_para_sair)

        # VOLTAR PARA A PAGINA INICIAL
        self.btnVoltarResgistrarSaida.clicked.connect(lambda: self.mainStackedWidget.setCurrentWidget(self.mainPage))

        # PESQUISAR ITEM NO CARDAPIO PARA SAIDA
        self.btnPesquisarTabelaCardapioParaSair.clicked.connect(self.pesquisar_item_cardapio_para_saida)

        # SLECIONAR ITENS NA TABELEA CARDAPIO PARA SAIR
        self.tabelaCardapioParaSaida.selectionModel().selectionChanged.connect(self.on_selectionChanged)
        self.btnSelecionarItemTabelaCardapioParaSair.clicked.connect(self.slecionar_item_para_sair)

        # CONFIRMAR SAIDA DE ITENS
        self.btnFinalizarSaidaDeItens.clicked.connect(self.registrar_saida)
        self.btnFinalizarSaidaDeItens.clicked.connect(self.zerar_tabela_itens_selecionados_para_sair)

        # TABELA ITENS SELECIONADOS
        self.btnZerarTabelaItensParaSair.clicked.connect(self.zerar_tabela_itens_selecionados_para_sair)

        # REDIMENSIONANDO TABELA
        self.tabelaCardapioParaSaida.setColumnWidth(0, 200)
        self.tabelaCardapioParaSaida.setColumnWidth(1, 200)
        self.tabelaItensSelecionadosParaSair.setColumnWidth(0, 200)
        self.tabelaItensSelecionadosParaSair.setRowHeight(0, 50)
        self.tabelaItensSelecionadosParaSair.setRowHeight(1, 50)
        self.tabelaItensSelecionadosParaSair.setRowHeight(2, 50)
        self.tabelaItensSelecionadosParaSair.setRowHeight(3, 50)
        self.tabelaItensSelecionadosParaSair.setRowHeight(4, 50)

        # CORRIGIR ITEM SELECIONADO PARA SAIR
        self.btnCorrigirItemSelecionado.clicked.connect(self.corrigir_item_selecionado_para_sair)

        ##############################################
        # PÁGINA OPÇÕES
        ##############################################
        # ACESSAR OPÇÕES
        self.opcoesBtn.clicked.connect(lambda: self.mainStackedWidget.setCurrentWidget(self.opcoesPage))

        # VOLTAR DE OPÇÕES
        self.btnVoltarDeOpcoes.clicked.connect(lambda: self.mainStackedWidget.setCurrentWidget(self.mainPage))

        # ENVIAR EMAI/ RELATÓRIO - ENCERRAR O DIA
        self.btnFinalizarDia.clicked.connect(self.enviar_email)

        ## AJUDA
        # ACESSAR AJUDA
        self.btnAjuda.clicked.connect(lambda: self.mainStackedWidget.setCurrentWidget(self.ajudaPage))

        ## CONTATO
        # ACESSAR CONTATO
        self.btnContato.clicked.connect(lambda: self.mainStackedWidget.setCurrentWidget(self.contatoPage))

        ## VOLTAR DA PÁGINA CONTATO E AJUDA
        self.btnVoltarContato.clicked.connect(lambda: self.mainStackedWidget.setCurrentWidget(self.opcoesPage))
        self.btnVoltarAjuda.clicked.connect(lambda: self.mainStackedWidget.setCurrentWidget(self.opcoesPage))

        ## TESTE ##

        self.btnAdicionarQuantidadeItemEstoque.clicked.connect(self.adicionar_quantia_ao_estoque)

        self.btnAlterarQuantidadeItemEstoque.clicked.connect(self.alterar_quantia_do_estoque)

    def on_selectionChanged(self, selected):
        global linha_selecionada
        global coluna_selecionada
        linha = ''
        coluna = ''

        for ix in selected.indexes():
            print('Selected Cell Location Row: {0}, Column: {1}'.format(ix.row(), ix.column()))
            linha = ix.row()
            coluna = 0

        linha_selecionada.append(linha)
        coluna_selecionada.append(coluna)

    def enviar_email(self):
        try:
            # PEGANDO INFORMAÇÕES DO ESTOQUE
            # database connection
            connection = pymysql.connect(host="sql886.main-hosting.eu",
                                         port=3306,
                                         user="u147843696_Teste",
                                         passwd="Teste12345",
                                         database="u147843696_Teste")
            cursor = connection.cursor()

            ## ESCREVER COMANDOS AQUI

            comando = f'SELECT * FROM estoque'
            cursor.execute(comando)

            dados_lidos = cursor.fetchall()  # ler banco de dados
            print(dados_lidos)

            cursor.close()
            connection.close()

            # DATA ATUAL
            data = datetime.today().strftime('%d-%m-%Y')
            data_texto = str(data)

            # CRIANDO ARQUIVO EXCEL PARA ENVIAR
            path = "planilha/"
            dir = listdir(path)
            for file in dir:
                print(file)
                if f'{data_texto}' not in file:
                    remove(f'planilha/{file}')

            lista = dados_lidos

            wb = Workbook()

            sh = wb.active

            sh['A1'] = 'ID'

            sh['B1'] = 'Nome'

            sh['C1'] = 'Quantidade'

            sh['D1'] = 'Medida'

            for indice, item in enumerate(lista):
                linha = indice + 2
                sh[f'A{linha}'] = lista[indice][0]

                sh[f'B{linha}'] = lista[indice][1]

                sh[f'C{linha}'] = lista[indice][2]

                sh[f'D{linha}'] = lista[indice][3]

            wb.save(filename=f'planilha/relatório {data_texto}.xlsx')

            # 1- STARTAR O SERVIDOR
            ##############################################
            host = "smtp.gmail.com"
            port = 587
            login = "midrajsistema@gmail.com"
            senha = "cjjmlrqeawrrlfcj"

            server = smtplib.SMTP(host, port)

            server.ehlo()
            server.starttls()
            server.login(login, senha)

            # 2- CONSTRUIR O EMAIL TIPO MIME
            ##############################################
            corpo = "Mensagem Teste"

            email_msg = MIMEMultipart()
            email_msg['From'] = login
            email_msg['To'] = "arthur.1055@hotmail.com"
            email_msg['Subject'] = "Email enviado por Arthur"
            email_msg.attach(MIMEText(corpo, 'plain'))

            # anexo
            caminhoArquivo = f'planilha/relatório {data_texto}.xlsx'
            attchment = open(caminhoArquivo, 'rb')

            att = MIMEBase('application', 'octet-stream')
            att.set_payload(attchment.read())
            encoders.encode_base64(att)

            att.add_header('Content-Disposition', f'attachment; filename=relatorio {data_texto}.xlsx')
            attchment.close()

            email_msg.attach(att)

            # 3- EVIAR O EMAIL TIPO MIME NO SERVIDOR SMTP
            ##############################################
            server.sendmail(email_msg['From'], email_msg['To'], email_msg.as_string())

            server.quit()

            print('oi')
            self.email_enviado()

        except:
            print('erro')

    def listar_dados(self):

        # database connection
        connection = pymysql.connect(host="sql886.main-hosting.eu",
                                     port=3306,
                                     user="u147843696_Teste",
                                     passwd="Teste12345",
                                     database="u147843696_Teste")
        cursor = connection.cursor()

        ## ESCREVER COMANDOS AQUI

        comando = f'SELECT * FROM estoque'
        cursor.execute(comando)

        dados_lidos = cursor.fetchall()  # ler banco de dados
        print(dados_lidos)

        # DEFININDO O TAMANHO DA TABELA
        self.tabelaMercadorias.setRowCount(len(dados_lidos))
        self.tabelaMercadorias.setColumnCount(4)

        # COLOCANDO OS ITENS NA TABELA
        for i in range(0, len(dados_lidos)):
            for j in range(0, 4):
                self.tabelaMercadorias.setItem(i, j, QtWidgets.QTableWidgetItem(str(dados_lidos[i][j])))

        cursor.close()
        connection.close()

    def finalizar_de_criar_item_estoque(self):
        nome = self.nomeItemCriado.text().lower()
        quantidade_str = self.quantidadeItemCriado.text()
        quantidade = quantidade_str.replace(',', '.')

        print(f"{nome}, {quantidade}")

        # database connection
        connection = pymysql.connect(host="sql886.main-hosting.eu",
                                     port=3306,
                                     user="u147843696_Teste",
                                     passwd="Teste12345",
                                     database="u147843696_Teste")
        cursor = connection.cursor()

        ## ESCREVER COMANDOS AQUI
        nome_item = nome
        valor = float(quantidade)
        comando = ''
        if self.unidadeComboBox.isChecked():
            comando = f'INSERT INTO u147843696_Teste.estoque (nome, quantidade, medida) VALUES ("{nome_item}", {valor}, "UNI")'
        elif self.kgComboBox.isChecked():
            valor = valor * 1000
            comando = f'INSERT INTO u147843696_Teste.estoque (nome, quantidade, medida) VALUES ("{nome_item}", {valor}, "gramas")'
        elif self.gramasComboBox.isChecked():
            comando = f'INSERT INTO u147843696_Teste.estoque (nome, quantidade, medida) VALUES ("{nome_item}", {valor}, "gramas")'

        cursor.execute(comando)
        connection.commit()  # edita banco de dados

        ## NÃO ESCREVER DEPOIS DAQUI

        cursor.close()
        connection.close()

    def atualizar_tabelinha_ultimos_itens_criados(self):
        nome = self.nomeItemCriado.text()
        quantidade_str = self.quantidadeItemCriado.text()
        quantidade_arrumada = quantidade_str.replace(',', '.')

        # QUANTIDADE EM FLOAT
        quantidade = float(quantidade_arrumada)

        #  CONDIÇÕES PARA CADA RADIOBUTTON CHECK DIFERENTE
        medida = ''
        if self.unidadeComboBox.isChecked():
            medida = "UNI"
        elif self.kgComboBox.isChecked():
            quantidade = quantidade * 1000
            medida = 'gramas'
        elif self.gramasComboBox.isChecked():
            medida = 'gramas'

        listaTabelinha = [nome, quantidade, medida]

        # ATUALIZANDO UMA LISTA AVULSA PARA SIMULAR OS UTILMOS ITENS ADCIONADOS E DEFININDO O TAMANHO MAXIMO DELA
        global mainListaTabelinha
        mainListaTabelinha.insert(0, listaTabelinha)
        print(mainListaTabelinha)
        if len(mainListaTabelinha) > 5:
            mainListaTabelinha.pop()

        for i in range(0, len(mainListaTabelinha)):
            for j in range(0, 3):
                self.tabelaUltimosItensAdicionados.setItem(i, j,
                                                           QtWidgets.QTableWidgetItem(str(mainListaTabelinha[i][j])))

    def atualizar_tabelinha_cardapio(self):
        nome = self.nomeItemCriadoCardapio.text()
        preco_str = self.precoItemCriadoCardapio.text()
        preco_arrumado = preco_str.replace(',', '.')

        # PREÇO EM FLOAT
        preco = float(preco_arrumado)

        print(preco)

        listaDaTabelinha = [nome, preco]

        # ATUALIZANDO UMA LISTA AVULSA PARA SIMULAR OS UTILMOS ITENS ADCIONADOS E DEFININDO O TAMANHO MAXIMO DELA
        global mainTabelinhaCardapio
        mainTabelinhaCardapio.insert(0, listaDaTabelinha)
        print(mainTabelinhaCardapio)

        if len(mainTabelinhaCardapio) > 5:
            mainTabelinhaCardapio.pop()

        for i in range(0, len(mainTabelinhaCardapio)):
            for j in range(0, 2):
                self.tabelaUltimosItensAdicionadosCardapio.setItem(i, j,
                                                           QtWidgets.QTableWidgetItem(str(mainTabelinhaCardapio[i][j])))

    def pesquisar_item_tabela_mercadoria(self):
        nome = self.nomeItemPesquisarTabelaMercadorias.text().lower()

        # database connection
        connection = pymysql.connect(host="sql886.main-hosting.eu",
                                     port=3306,
                                     user="u147843696_Teste",
                                     passwd="Teste12345",
                                     database="u147843696_Teste")
        cursor = connection.cursor()
        ##

        comando = f'SELECT * FROM estoque WHERE nome like "%{nome}%"'
        cursor.execute(comando)

        dados_lidos = cursor.fetchall()  # ler banco de dados
        print(dados_lidos)

        # DEFININDO O TAMANHO DA TABELA
        self.tabelaMercadorias.setRowCount(len(dados_lidos))
        self.tabelaMercadorias.setColumnCount(4)

        # COLOCANDO OS ITENS NA TABELA
        for i in range(0, len(dados_lidos)):
            for j in range(0, 4):
                self.tabelaMercadorias.setItem(i, j, QtWidgets.QTableWidgetItem(str(dados_lidos[i][j])))
        print(nome)

    def finalizar_criar_item_cardapio(self):
        try:
            nome = self.nomeItemCriadoCardapio.text().lower()
            preco = self.precoItemCriadoCardapio.text()
            preco_float = preco.replace(',', '.')

            # database connection
            connection = pymysql.connect(host="sql886.main-hosting.eu",
                                         port=3306,
                                         user="u147843696_Teste",
                                         passwd="Teste12345",
                                         database="u147843696_Teste")
            cursor = connection.cursor()

            # CREATE
            nome_item = nome
            preco_do_item = float(preco_float)
            comando = f'INSERT INTO cardapio (Nome, Preco) VALUES ("{nome_item}", {preco_do_item})'
            cursor.execute(comando)

            connection.commit()  # edita banco de dados

            cursor.close()
            connection.close()

            contador_de_ingredientes = 0

            ingredientes = []

            ingrediente1 = ''
            ingrediente1 = self.nomeItemGasto1.text().lower()
            ingrediente1_quantidade_str = ''
            ingrediente1_quantidade_str = self.quantidadeItemCriado1.text().lower()
            ingrediente1_medida = ''
            ingrediente1_medida = self.comboBoxMedida1.currentText().lower()

            ingrediente2 = ''
            ingrediente2 = self.nomeItemGasto2.text().lower()
            ingrediente2_quantidade_str = ''
            ingrediente2_quantidade_str = self.quantidadeItemCriado2.text().lower()
            ingrediente2_medida = ''
            ingrediente2_medida = self.comboBoxMedida2.currentText().lower()

            ingrediente3 = ''
            ingrediente3 = self.nomeItemGasto3.text().lower()
            ingrediente3_quantidade_str = ''
            ingrediente3_quantidade_str = self.quantidadeItemCriado3.text().lower()
            ingrediente3_medida = ''
            ingrediente3_medida = self.comboBoxMedida3.currentText().lower()

            ingrediente4 = ''
            ingrediente4 = self.nomeItemGasto4.text().lower()
            ingrediente4_quantidade_str = ''
            ingrediente4_quantidade_str = self.quantidadeItemCriado4.text().lower()
            ingrediente4_medida = ''
            ingrediente4_medida = self.comboBoxMedida4.currentText().lower()

            ingrediente5 = ''
            ingrediente5 = self.nomeItemGasto5.text().lower()
            ingrediente5_quantidade_str = ''
            ingrediente5_quantidade_str = self.quantidadeItemCriado5.text().lower()
            ingrediente5_medida = ''
            ingrediente5_medida = self.comboBoxMedida5.currentText().lower()

            ingrediente1_quantidade = ''
            ingrediente2_quantidade = ''
            ingrediente3_quantidade = ''
            ingrediente4_quantidade = ''
            ingrediente5_quantidade = ''

            if ingrediente1_quantidade_str != '':
                ingrediente1_quantidade = float(ingrediente1_quantidade_str)
                contador_de_ingredientes += 1
            if ingrediente2_quantidade_str != '':
                ingrediente2_quantidade = float(ingrediente2_quantidade_str)
                contador_de_ingredientes += 1
            if ingrediente3_quantidade_str != '':
                ingrediente3_quantidade = float(ingrediente3_quantidade_str)
                contador_de_ingredientes += 1
            if ingrediente4_quantidade_str != '':
                ingrediente4_quantidade = float(ingrediente4_quantidade_str)
                contador_de_ingredientes += 1
            if ingrediente5_quantidade_str != '':
                ingrediente5_quantidade = float(ingrediente5_quantidade_str)
                contador_de_ingredientes += 1
            print(contador_de_ingredientes)

            '''
            ingredientes = [ingrediente1, ingrediente1_quantidade, ingrediente1_medida], \
                           [ingrediente2, ingrediente2_quantidade, ingrediente2_medida], \
                           [ingrediente3, ingrediente3_quantidade, ingrediente3_medida], \
                           [ingrediente4, ingrediente4_quantidade, ingrediente4_medida], \
                           [ingrediente5, ingrediente5_quantidade, ingrediente5_medida]
    
            print(ingredientes)
            '''

            # database connection
            connection = pymysql.connect(host="sql886.main-hosting.eu",
                                         port=3306,
                                         user="u147843696_Teste",
                                         passwd="Teste12345",
                                         database="u147843696_Teste")
            cursor = connection.cursor()

            ## ESCREVER COMANDOS AQUI
            if contador_de_ingredientes == 5:
                comando = f'INSERT INTO cardapio_ingredientes (nome, ingrediente1Nome, ingrediente1Quantia, ingrediente1Medida, ingrediente2Nome, ingrediente2Quantia, ingrediente2Medida, ingrediente3Nome, ingrediente3Quantia, ingrediente3Medida, ingrediente4Nome, ingrediente4Quantia, ingrediente4Medida, ingrediente5Nome, ingrediente5Quantia, ingrediente5Medida) VALUES ("{nome}", "{ingrediente1}", {ingrediente1_quantidade}, "{ingrediente1_medida}", "{ingrediente2}", {ingrediente2_quantidade}, "{ingrediente2_medida}", "{ingrediente3}", {ingrediente3_quantidade}, "{ingrediente3_medida}", "{ingrediente4}", {ingrediente4_quantidade}, "{ingrediente4_medida}", "{ingrediente5}", {ingrediente5_quantidade}, "{ingrediente5_medida}")'

            if contador_de_ingredientes == 4:
                comando = f'INSERT INTO cardapio_ingredientes (nome, ingrediente1Nome, ingrediente1Quantia, ingrediente1Medida, ingrediente2Nome, ingrediente2Quantia, ingrediente2Medida, ingrediente3Nome, ingrediente3Quantia, ingrediente3Medida, ingrediente4Nome, ingrediente4Quantia, ingrediente4Medida) VALUES ("{nome}", "{ingrediente1}", {ingrediente1_quantidade}, "{ingrediente1_medida}", "{ingrediente2}", {ingrediente2_quantidade}, "{ingrediente2_medida}", "{ingrediente3}", {ingrediente3_quantidade}, "{ingrediente3_medida}", "{ingrediente4}", {ingrediente4_quantidade}, "{ingrediente4_medida}")'

            if contador_de_ingredientes == 3:
                comando = f'INSERT INTO cardapio_ingredientes (nome, ingrediente1Nome, ingrediente1Quantia, ingrediente1Medida, ingrediente2Nome, ingrediente2Quantia, ingrediente2Medida, ingrediente3Nome, ingrediente3Quantia, ingrediente3Medida) VALUES ("{nome}", "{ingrediente1}", {ingrediente1_quantidade}, "{ingrediente1_medida}", "{ingrediente2}", {ingrediente2_quantidade}, "{ingrediente2_medida}", "{ingrediente3}", {ingrediente3_quantidade}, "{ingrediente3_medida}")'

            if contador_de_ingredientes == 2:
                comando = f'INSERT INTO cardapio_ingredientes (nome, ingrediente1Nome, ingrediente1Quantia, ingrediente1Medida, ingrediente2Nome, ingrediente2Quantia, ingrediente2Medida) VALUES ("{nome}", "{ingrediente1}", {ingrediente1_quantidade}, "{ingrediente1_medida}", "{ingrediente2}", {ingrediente2_quantidade}, "{ingrediente2_medida}")'

            if contador_de_ingredientes == 1:
                comando = f'INSERT INTO cardapio_ingredientes (nome, ingrediente1Nome, ingrediente1Quantia, ingrediente1Medida) VALUES ("{nome}", "{ingrediente1}", {ingrediente1_quantidade}, "{ingrediente1_medida}")'

            cursor.execute(comando)
            connection.commit()

            ## NÃO ESCREVER DEPOIS DAQUI

            cursor.close()
            connection.close()

        except:
            self.mostrar_menssagem_erro()

    def habilitar_ingrediente_2(self):
        self.nomeItemGasto2.setDisabled(False)
        self.quantidadeItemCriado2.setDisabled(False)
        self.comboBoxMedida2.setDisabled(False)

    def habilitar_ingrediente_3(self):
        self.nomeItemGasto3.setDisabled(False)
        self.quantidadeItemCriado3.setDisabled(False)
        self.comboBoxMedida3.setDisabled(False)

    def habilitar_ingrediente_4(self):
        self.nomeItemGasto4.setDisabled(False)
        self.quantidadeItemCriado4.setDisabled(False)
        self.comboBoxMedida4.setDisabled(False)

    def habilitar_ingrediente_5(self):
        self.nomeItemGasto5.setDisabled(False)
        self.quantidadeItemCriado5.setDisabled(False)
        self.comboBoxMedida5.setDisabled(False)

    def registrar_saida(self):
        global contadorItensAdicionadosParaSair
        lista_selecionados = []

        for i in range(0, contadorItensAdicionadosParaSair):
            selecionados = self.tabelaItensSelecionadosParaSair.item(i, 0).text()
            lista_selecionados.append(selecionados)
            print(i)

        print(lista_selecionados)

        for index, item in enumerate(lista_selecionados):
            nome = lista_selecionados[index]

            quantidadeSaindo = self.tabelaItensSelecionadosParaSair.item(index, 1).text()
            quantidadeSaindo_int = int(quantidadeSaindo)
            print(f'{quantidadeSaindo_int} quantidade saindo')

            # database connection
            connection = pymysql.connect(host="sql886.main-hosting.eu",
                                         port=3306,
                                         user="u147843696_Teste",
                                         passwd="Teste12345",
                                         database="u147843696_Teste")
            cursor = connection.cursor()

            ## ESCREVER COMANDOS AQUI

            comando = f'SELECT * FROM cardapio_ingredientes WHERE nome like "{nome}"'
            cursor.execute(comando)

            resultado = cursor.fetchall()  # ler banco de dados

            ## NÃO ESCREVER DEPOIS DAQUI

            cursor.close()
            connection.close()

            lista_ingredientes = []
            indice = 0
            while indice < 15:
                ingrediente = resultado[0][indice+1]
                if indice == 0:
                    nome_do_prato = resultado[0][indice]
                indice += 1

                if ingrediente != None:
                    lista_ingredientes.append(ingrediente)

            print(len(lista_ingredientes))

            if len(lista_ingredientes) == 3:
                nome_ingrediente = lista_ingredientes[0]
                ingrediente_quantidade = lista_ingredientes[1]
                medida_ingrediente = lista_ingredientes[2]

                # database connection
                connection = pymysql.connect(host="sql886.main-hosting.eu",
                                             port=3306,
                                             user="u147843696_Teste",
                                             passwd="Teste12345",
                                             database="u147843696_Teste")
                cursor = connection.cursor()

                ## ESCREVER COMANDOS AQUI

                comando = f'UPDATE estoque SET quantidade = quantidade - {ingrediente_quantidade} * {quantidadeSaindo_int} WHERE nome = "{nome_ingrediente}" AND medida = "{medida_ingrediente}"'
                cursor.execute(comando)
                connection.commit()

                ## NÃO ESCREVER DEPOIS DAQUI

                cursor.close()
                connection.close()

            if len(lista_ingredientes) == 6:
                nome_ingrediente = lista_ingredientes[0]
                ingrediente_quantidade = lista_ingredientes[1]
                medida_ingrediente = lista_ingredientes[2]

                # database connection
                connection = pymysql.connect(host="sql886.main-hosting.eu",
                                             port=3306,
                                             user="u147843696_Teste",
                                             passwd="Teste12345",
                                             database="u147843696_Teste")
                cursor = connection.cursor()

                ## ESCREVER COMANDOS AQUI

                comando = f'UPDATE estoque SET quantidade = quantidade - {ingrediente_quantidade} * {quantidadeSaindo_int} WHERE nome = "{nome_ingrediente}" AND medida = "{medida_ingrediente}"'
                cursor.execute(comando)
                connection.commit()

                ###
                nome_ingrediente = lista_ingredientes[3]
                ingrediente_quantidade = lista_ingredientes[4]
                medida_ingrediente = lista_ingredientes[5]

                ## ESCREVER COMANDOS AQUI

                comando = f'UPDATE estoque SET quantidade = quantidade - {ingrediente_quantidade} * {quantidadeSaindo_int} WHERE nome = "{nome_ingrediente}" AND medida = "{medida_ingrediente}"'
                cursor.execute(comando)
                connection.commit()

                ## NÃO ESCREVER DEPOIS DAQUI

                cursor.close()
                connection.close()

            if len(lista_ingredientes) == 9:
                nome_ingrediente = lista_ingredientes[0]
                ingrediente_quantidade = lista_ingredientes[1]
                medida_ingrediente = lista_ingredientes[2]

                # database connection
                connection = pymysql.connect(host="sql886.main-hosting.eu",
                                             port=3306,
                                             user="u147843696_Teste",
                                             passwd="Teste12345",
                                             database="u147843696_Teste")
                cursor = connection.cursor()

                ## ESCREVER COMANDOS AQUI

                comando = f'UPDATE estoque SET quantidade = quantidade - {ingrediente_quantidade} * {quantidadeSaindo_int} WHERE nome = "{nome_ingrediente}" AND medida = "{medida_ingrediente}"'
                cursor.execute(comando)
                connection.commit()

                ## NÃO ESCREVER DEPOIS DAQUI

                nome_ingrediente = lista_ingredientes[3]
                ingrediente_quantidade = lista_ingredientes[4]
                medida_ingrediente = lista_ingredientes[5]

                ## ESCREVER COMANDOS AQUI

                comando = f'UPDATE estoque SET quantidade = quantidade - {ingrediente_quantidade} * {quantidadeSaindo_int} WHERE nome = "{nome_ingrediente}" AND medida = "{medida_ingrediente}"'
                cursor.execute(comando)
                connection.commit()

                ## NÃO ESCREVER DEPOIS DAQUI

                nome_ingrediente = lista_ingredientes[6]
                ingrediente_quantidade = lista_ingredientes[7]
                medida_ingrediente = lista_ingredientes[8]

                ## ESCREVER COMANDOS AQUI

                comando = f'UPDATE estoque SET quantidade = quantidade - {ingrediente_quantidade} * {quantidadeSaindo_int} WHERE nome = "{nome_ingrediente}" AND medida = "{medida_ingrediente}"'
                cursor.execute(comando)
                connection.commit()

                ## NÃO ESCREVER DEPOIS DAQUI

                cursor.close()
                connection.close()

            if len(lista_ingredientes) == 12:
                nome_ingrediente = lista_ingredientes[0]
                ingrediente_quantidade = lista_ingredientes[1]
                medida_ingrediente = lista_ingredientes[2]

                # database connection
                connection = pymysql.connect(host="sql886.main-hosting.eu",
                                             port=3306,
                                             user="u147843696_Teste",
                                             passwd="Teste12345",
                                             database="u147843696_Teste")
                cursor = connection.cursor()

                ## ESCREVER COMANDOS AQUI

                comando = f'UPDATE estoque SET quantidade = quantidade - {ingrediente_quantidade} * {quantidadeSaindo_int} WHERE nome = "{nome_ingrediente}" AND medida = "{medida_ingrediente}"'
                cursor.execute(comando)
                connection.commit()

                ## NÃO ESCREVER DEPOIS DAQUI

                nome_ingrediente = lista_ingredientes[3]
                ingrediente_quantidade = lista_ingredientes[4]
                medida_ingrediente = lista_ingredientes[5]

                ## ESCREVER COMANDOS AQUI

                comando = f'UPDATE estoque SET quantidade = quantidade - {ingrediente_quantidade} * {quantidadeSaindo_int} WHERE nome = "{nome_ingrediente}" AND medida = "{medida_ingrediente}"'
                cursor.execute(comando)
                connection.commit()

                ## NÃO ESCREVER DEPOIS DAQUI

                nome_ingrediente = lista_ingredientes[6]
                ingrediente_quantidade = lista_ingredientes[7]
                medida_ingrediente = lista_ingredientes[8]

                ## ESCREVER COMANDOS AQUI

                comando = f'UPDATE estoque SET quantidade = quantidade - {ingrediente_quantidade} * {quantidadeSaindo_int} WHERE nome = "{nome_ingrediente}" AND medida = "{medida_ingrediente}"'
                cursor.execute(comando)
                connection.commit()

                ## NÃO ESCREVER DEPOIS DAQUI

                nome_ingrediente = lista_ingredientes[9]
                ingrediente_quantidade = lista_ingredientes[10]
                medida_ingrediente = lista_ingredientes[11]

                ## ESCREVER COMANDOS AQUI

                comando = f'UPDATE estoque SET quantidade = quantidade - {ingrediente_quantidade} * {quantidadeSaindo_int} WHERE nome = "{nome_ingrediente}" AND medida = "{medida_ingrediente}"'
                cursor.execute(comando)
                connection.commit()

                ## NÃO ESCREVER DEPOIS DAQUI

                cursor.close()
                connection.close()

            if len(lista_ingredientes) == 15:
                nome_ingrediente = lista_ingredientes[0]
                ingrediente_quantidade = lista_ingredientes[1]
                medida_ingrediente = lista_ingredientes[2]

                # database connection
                connection = pymysql.connect(host="sql886.main-hosting.eu",
                                             port=3306,
                                             user="u147843696_Teste",
                                             passwd="Teste12345",
                                             database="u147843696_Teste")
                cursor = connection.cursor()

                ## ESCREVER COMANDOS AQUI

                comando = f'UPDATE estoque SET quantidade = quantidade - {ingrediente_quantidade} * {quantidadeSaindo_int} WHERE nome = "{nome_ingrediente}" AND medida = "{medida_ingrediente}"'
                cursor.execute(comando)
                connection.commit()

                ## NÃO ESCREVER DEPOIS DAQUI

                nome_ingrediente = lista_ingredientes[3]
                ingrediente_quantidade = lista_ingredientes[4]
                medida_ingrediente = lista_ingredientes[5]

                ## ESCREVER COMANDOS AQUI

                comando = f'UPDATE estoque SET quantidade = quantidade - {ingrediente_quantidade} * {quantidadeSaindo_int} WHERE nome = "{nome_ingrediente}" AND medida = "{medida_ingrediente}"'
                cursor.execute(comando)
                connection.commit()

                ## NÃO ESCREVER DEPOIS DAQUI

                nome_ingrediente = lista_ingredientes[6]
                ingrediente_quantidade = lista_ingredientes[7]
                medida_ingrediente = lista_ingredientes[8]

                ## ESCREVER COMANDOS AQUI

                comando = f'UPDATE estoque SET quantidade = quantidade - {ingrediente_quantidade} * {quantidadeSaindo_int} WHERE nome = "{nome_ingrediente}" AND medida = "{medida_ingrediente}"'
                cursor.execute(comando)
                connection.commit()

                ## NÃO ESCREVER DEPOIS DAQUI

                nome_ingrediente = lista_ingredientes[9]
                ingrediente_quantidade = lista_ingredientes[10]
                medida_ingrediente = lista_ingredientes[11]

                ## ESCREVER COMANDOS AQUI

                comando = f'UPDATE estoque SET quantidade = quantidade - {ingrediente_quantidade} * {quantidadeSaindo_int} WHERE nome = "{nome_ingrediente}" AND medida = "{medida_ingrediente}"'
                cursor.execute(comando)
                connection.commit()

                ## NÃO ESCREVER DEPOIS DAQUI

                nome_ingrediente = lista_ingredientes[12]
                ingrediente_quantidade = lista_ingredientes[13]
                medida_ingrediente = lista_ingredientes[14]

                ## ESCREVER COMANDOS AQUI

                comando = f'UPDATE estoque SET quantidade = quantidade - {ingrediente_quantidade} * {quantidadeSaindo_int} WHERE nome = "{nome_ingrediente}" AND medida = "{medida_ingrediente}"'
                cursor.execute(comando)
                connection.commit()

                ## NÃO ESCREVER DEPOIS DAQUI

                cursor.close()
                connection.close()

                print('oi')

        print('Feito')

    def zerar_caixas_de_texto_criar_item_cardapio(self):
        self.nomeItemCriadoCardapio.setText('')
        self.precoItemCriadoCardapio.setText('')

        self.nomeItemGasto1.setText('')
        self.quantidadeItemCriado1.setText('')
        self.nomeItemGasto2.setText('')
        self.quantidadeItemCriado2.setText('')
        self.nomeItemGasto3.setText('')
        self.quantidadeItemCriado3.setText('')
        self.nomeItemGasto4.setText('')
        self.quantidadeItemCriado4.setText('')
        self.nomeItemGasto4.setText('')
        self.quantidadeItemCriado4.setText('')
        self.nomeItemGasto5.setText('')
        self.quantidadeItemCriado5.setText('')

        self.comboBoxMedida1.setCurrentIndex(0)
        self.comboBoxMedida2.setCurrentIndex(0)
        self.comboBoxMedida3.setCurrentIndex(0)
        self.comboBoxMedida4.setCurrentIndex(0)
        self.comboBoxMedida5.setCurrentIndex(0)

        self.nomeItemGasto2.setDisabled(True)
        self.quantidadeItemCriado2.setDisabled(True)
        self.comboBoxMedida2.setDisabled(True)

        self.nomeItemGasto3.setDisabled(True)
        self.quantidadeItemCriado3.setDisabled(True)
        self.comboBoxMedida3.setDisabled(True)

        self.nomeItemGasto4.setDisabled(True)
        self.quantidadeItemCriado4.setDisabled(True)
        self.comboBoxMedida4.setDisabled(True)

        self.nomeItemGasto5.setDisabled(True)
        self.quantidadeItemCriado5.setDisabled(True)
        self.comboBoxMedida5.setDisabled(True)

    def listar_itens_cardapio_para_sair(self):
        # database connection
        connection = pymysql.connect(host="sql886.main-hosting.eu",
                                     port=3306,
                                     user="u147843696_Teste",
                                     passwd="Teste12345",
                                     database="u147843696_Teste")
        cursor = connection.cursor()

        ## ESCREVER COMANDOS AQUI

        comando = f'SELECT * FROM cardapio'
        cursor.execute(comando)

        dados_lidos = cursor.fetchall()  # ler banco de dados
        print(dados_lidos)

        # DEFININDO O TAMANHO DA TABELA
        self.tabelaCardapioParaSaida.setRowCount(len(dados_lidos))
        self.tabelaCardapioParaSaida.setColumnCount(2)

        # COLOCANDO OS ITENS NA TABELA
        for i in range(0, len(dados_lidos)):
            for j in range(0, 2):
                self.tabelaCardapioParaSaida.setItem(i, j, QtWidgets.QTableWidgetItem(str(dados_lidos[i][j])))

        cursor.close()
        connection.close()

    def slecionar_item_para_sair(self):
        global linha_selecionada
        global coluna_selecionada
        global itemSelecionadoNome
        global contadorItensAdicionadosParaSair

        if len(linha_selecionada) > 1:
            del linha_selecionada[:-1]
        if len(coluna_selecionada) > 1:
            del coluna_selecionada[:-1]

        row = linha_selecionada[0]
        col = coluna_selecionada[0]
        itemSelecionadoNome = self.tabelaCardapioParaSaida.item(row, col).text()
        print(itemSelecionadoNome)

        if contadorItensAdicionadosParaSair > 4:
            contadorItensAdicionadosParaSair = 0
            self.zerar_tabela_itens_selecionados_para_sair()

        if contadorItensAdicionadosParaSair == 0:
            self.tabelaItensSelecionadosParaSair.setItem(0, 0, QtWidgets.QTableWidgetItem(itemSelecionadoNome))

        if contadorItensAdicionadosParaSair == 1:
            self.tabelaItensSelecionadosParaSair.setItem(1, 0, QtWidgets.QTableWidgetItem(itemSelecionadoNome))

        if contadorItensAdicionadosParaSair == 2:
            self.tabelaItensSelecionadosParaSair.setItem(2, 0, QtWidgets.QTableWidgetItem(itemSelecionadoNome))

        if contadorItensAdicionadosParaSair == 3:
            self.tabelaItensSelecionadosParaSair.setItem(3, 0, QtWidgets.QTableWidgetItem(itemSelecionadoNome))

        if contadorItensAdicionadosParaSair == 4:
            self.tabelaItensSelecionadosParaSair.setItem(4, 0, QtWidgets.QTableWidgetItem(itemSelecionadoNome))

        contadorItensAdicionadosParaSair += 1
        print(contadorItensAdicionadosParaSair)

    def pesquisar_item_cardapio_para_saida(self):
        nome = self.nomeItemPesquisarParaSaida.text().lower()

        # database connection
        connection = pymysql.connect(host="sql886.main-hosting.eu",
                                     port=3306,
                                     user="u147843696_Teste",
                                     passwd="Teste12345",
                                     database="u147843696_Teste")
        cursor = connection.cursor()
        ##

        comando = f'SELECT * FROM cardapio WHERE nome like "%{nome}%"'
        cursor.execute(comando)

        dados_lidos = cursor.fetchall()  # ler banco de dados

        # DEFININDO O TAMANHO DA TABELA
        self.tabelaCardapioParaSaida.setRowCount(len(dados_lidos))
        self.tabelaCardapioParaSaida.setColumnCount(2)

        # COLOCANDO OS ITENS NA TABELA
        for i in range(0, len(dados_lidos)):
            for j in range(0, 1):
                self.tabelaCardapioParaSaida.setItem(i, j, QtWidgets.QTableWidgetItem(str(dados_lidos[i][j])))

    def finalizar_saida_de_itens(self):
        global contadorItensAdicionadosParaSair
        lista_selecionados = []

        for i in range(0, contadorItensAdicionadosParaSair):
            selecionados = self.tabelaItensSelecionadosParaSair.item(i, 0).text()
            lista_selecionados.append(selecionados)

        print(lista_selecionados)

    def zerar_tabela_itens_selecionados_para_sair(self):
        global contadorItensAdicionadosParaSair

        contadorItensAdicionadosParaSair = 0
        for i in range(0, 5):
            self.tabelaItensSelecionadosParaSair.setItem(i, 0, QtWidgets.QTableWidgetItem(''))

        for i in range(0, 5):
            self.tabelaItensSelecionadosParaSair.setItem(i, 1, QtWidgets.QTableWidgetItem('1'))

    def zerar_caixas_de_texto_criar_item_estoque(self):
        self.nomeItemCriado.setText('')
        self.quantidadeItemCriado.setText('')

    def corrigir_item_selecionado_para_sair(self):
        global contadorItensAdicionadosParaSair

        for i in range(1, 6):
            if contadorItensAdicionadosParaSair == i:
                self.tabelaItensSelecionadosParaSair.setItem(i-1, 0, QtWidgets.QTableWidgetItem(''))
                contadorItensAdicionadosParaSair -= 1

    def mostrar_menssagem_erro(self):
        msg = QMessageBox()
        msg.setWindowTitle("Erro")
        msg.setText("Algo deu errado")
        msg.setIcon(QMessageBox.Warning)
        msg.setInformativeText("Por Favor, checar as entradas de texto, se falta algo etc. \n\nSe o erro persistir procure por:\nOpções - Ajuda ou Contato")
        x = msg.exec_()

    def email_enviado(self):
        msg = QMessageBox()
        msg.setWindowTitle("Concluido")
        msg.setText("\nRelatório enviado!")
        msg.setIcon(QMessageBox.Information)
        x = msg.exec_()

    def adicionar_quantia_ao_estoque(self):
        ID = self.idItemParaAdicionar.text()
        adicionar_quantidade = self.quantidadeParaAdicionar.text()
        adicionar_quantidade_float = adicionar_quantidade.replace(',', '.')
        quantidade = float(adicionar_quantidade_float)

        #  CONDIÇÕES PARA CADA RADIOBUTTON CHECK DIFERENTE
        medida = ''
        if self.radioBtnUniAdicionarEstoque.isChecked():
            medida = "UNI"
        elif self.radioBtnKgAdicionarEstoque.isChecked():
            quantidade = quantidade * 1000
            medida = 'gramas'
        elif self.radioBtnGramasAdicionarEstoque.isChecked():
            medida = 'gramas'

        # database connection
        connection = pymysql.connect(host="sql886.main-hosting.eu",
                                     port=3306,
                                     user="u147843696_Teste",
                                     passwd="Teste12345",
                                     database="u147843696_Teste")
        cursor = connection.cursor()

        ## ESCREVER COMANDOS AQUI

        comando = f'UPDATE estoque SET quantidade = quantidade + {quantidade}  WHERE ID = "{ID}" AND medida = "{medida}"'

        cursor.execute(comando)
        connection.commit()

        ## NÃO ESCREVER DEPOIS DAQUI

        cursor.close()
        connection.close()

        print('oi')

        self.idItemParaAdicionar.setText('')
        self.quantidadeParaAdicionar.setText('')

    def alterar_quantia_do_estoque(self):
        ID = self.idItemParaAlterar.text()
        adicionar_quantidade = self.quantidadeParaAlterar.text()
        adicionar_quantidade_float = adicionar_quantidade.replace(',', '.')
        quantidade = float(adicionar_quantidade_float)

        #  CONDIÇÕES PARA CADA RADIOBUTTON CHECK DIFERENTE
        medida = ''
        if self.radioBtnUniAlterarEstoque.isChecked():
            medida = "UNI"
        elif self.radioBtnKgAlterarEstoque.isChecked():
            quantidade = quantidade * 1000
            medida = 'gramas'
        elif self.radioBtnGramasAlterarEstoque.isChecked():
            medida = 'gramas'

        # database connection
        connection = pymysql.connect(host="sql886.main-hosting.eu",
                                     port=3306,
                                     user="u147843696_Teste",
                                     passwd="Teste12345",
                                     database="u147843696_Teste")
        cursor = connection.cursor()

        ## ESCREVER COMANDOS AQUI

        comando = f'UPDATE estoque SET quantidade = {quantidade}  WHERE ID = "{ID}" AND medida = "{medida}"'

        cursor.execute(comando)
        connection.commit()

        ## NÃO ESCREVER DEPOIS DAQUI

        cursor.close()
        connection.close()

        print('oi')

        self.idItemParaAlterar.setText('')
        self.quantidadeParaAlterar.setText('')

    def listar_cardapio(self):
        # database connection
        connection = pymysql.connect(host="sql886.main-hosting.eu",
                                     port=3306,
                                     user="u147843696_Teste",
                                     passwd="Teste12345",
                                     database="u147843696_Teste")
        cursor = connection.cursor()

        ## ESCREVER COMANDOS AQUI

        comando = f'SELECT * FROM cardapio'
        cursor.execute(comando)

        dados_lidos = cursor.fetchall()  # ler banco de dados
        print(dados_lidos)

        # DEFININDO O TAMANHO DA TABELA
        self.tabelaCardapio.setRowCount(len(dados_lidos))
        self.tabelaCardapio.setColumnCount(2)

        # COLOCANDO OS ITENS NA TABELA
        for i in range(0, len(dados_lidos)):
            for j in range(0, 2):
                self.tabelaCardapio.setItem(i, j, QtWidgets.QTableWidgetItem(str(dados_lidos[i][j])))

        cursor.close()
        connection.close()

    def relatorio(self):
        pass


if __name__ == '__main__':
    app = QApplication(sys.argv)
    myWin = MyMainWindow()
    myWin.showFullScreen()
    sys.exit(app.exec_())
