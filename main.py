import xmltodict
print(xmltodict)
import os
import sys
import requests
from PIL import Image
from PySide6 import QtWidgets, QtGui, QtCore
from PySide6.QtWidgets import QApplication, QFileDialog, QMessageBox, QWidget
from PySide6.QtGui import QShortcut, QKeySequence
import pandas as pd


def pegar_infos(nome_arquivo, valores):
    print("PEGOU AS INFORMAÇÕES DO arquivo", nome_arquivo)
    with open(nome_arquivo, "rb") as arquivo_xml:
        dic_arquivo = xmltodict.parse(arquivo_xml)

        if 'nfeProc' in dic_arquivo:
            infos_nf = dic_arquivo['nfeProc']['NFe']['infNFe']
            chave_nf = infos_nf['@Id']
            fornecedor = infos_nf['emit'].get('xNome', 'Não informado')
            cnpj_fornecedor = infos_nf['emit'].get('CNPJ', 'Não informado')
            inscricao_estadual = infos_nf['emit'].get('IE', 'Não informado')
            numero_nota = infos_nf['ide'].get('nNF', 'Não informado')
            serie = infos_nf['ide'].get('serie', 'Não informado')
            data_emissao = infos_nf['ide'].get('dhEmi', 'Não informado')
            data_emissao = data_emissao.split('T')[0]
            data_emissao = data_emissao.replace('-', '/')
            data_sai_ent = infos_nf['ide'].get('dhSaiEnt', 'Não informado')
            data_sai_ent = data_sai_ent.split('T')[0]
            data_sai_ent = data_sai_ent.replace('-', '/')
            

            nome_cliente = infos_nf['dest'].get('xNome', 'Não informado')
            cnpj_cliente = infos_nf['dest'].get('CNPJ', 'Não informado')
            nome_do_arquivo = nome_arquivo.split('/')[-1]
            nome_do_arquivo = nome_do_arquivo.split("\\")[-1]

            produtos = infos_nf['det']

            if not isinstance(produtos, list):
                produtos = [produtos]

        try:

            lista_insercao = []
            for produto in produtos:
                numero_item = produto['@nItem']
                cod_produto = produto['prod'].get('cProd', 'Não informado')
                ean = produto['prod'].get('cEAN', 'Não informado')
                nome_produto = produto['prod'].get('xProd', 'Não informado')
                ncm = produto['prod'].get('NCM', 'Não informado')
                cfop = produto['prod'].get('CFOP', 'Não informado')
                quantidade = int(float(produto['prod'].get('qCom', 'Não informado')))
                valor_unitario = produto['prod'].get('vUnCom', 'Não informado')
                valor_frete = produto['prod'].get('vFrete', 'Não informado')
                valor_seguro = produto['prod'].get('vSeg', 'Não informado')
                valor_desconto = produto['prod'].get('vDesc', 'Não informado')
                valor_outros = produto['prod'].get('vOutro', 'Não informado')
                valor_produto = produto['prod'].get('vProd', 'Não informado')
                ean_tributado = produto['prod'].get('cEANTrib', 'Não informado')
                unidade_tributada = produto['prod'].get('uTrib', 'Não informado')
                quantidade_tributada = int(float(produto['prod'].get('qTrib', 'Não informado')))
                valor_unitario_tributado = produto['prod'].get('vUnTrib', 'Não informado')
                ind_tot = produto['prod'].get('indTot', 'Não informado')
                if valor_unitario != 'Não informado':
                    valor_unitario = f"{float(valor_unitario):.2f}"
                if valor_unitario_tributado != 'Não informado':
                    valor_unitario_tributado = f"{float(valor_unitario_tributado):.2f}"
                

                icms = produto['imposto'].get('ICMS', {})
                icms_original = 'Não informado'
                icms_cst = 'Não informado'
                icms_mod_bc = 'Não informado'
                icms_vbc = 'Não informado'
                icms_picms = 'Não informado'
                icms_valor = 'Não informado'

                for chave, valor in icms.items():
                    if isinstance(valor, dict):
                        icms_original = valor.get('orig', icms_original)
                        icms_cst = valor.get('CST', icms_cst)
                        icms_mod_bc = valor.get('modBC', icms_mod_bc)
                        icms_vbc = valor.get('vBC', icms_vbc)
                        icms_picms = valor.get('pICMS', icms_picms)
                        icms_valor = valor.get('vICMS', icms_valor)
                
                pis = produto['imposto'].get('PIS', {})
                pis_cst = 'Não informado'
                pis_vbc = 'Não informado'
                pis_ppis = 'Não informado'
                pis_vpis = 'Não informado'

                for chave, valor in pis.items():
                    if isinstance(valor, dict):
                        pis_cst = valor.get('CST', pis_cst)
                        pis_vbc = valor.get('vBC', pis_vbc)
                        pis_ppis = valor.get('pPIS', pis_ppis)
                        pis_vpis = valor.get('vPIS', pis_vpis)

                cofins = produto['imposto'].get('COFINS', {})
                cofins_cst = 'Não informado'
                cofins_vbc = 'Não informado'
                cofins_pcofins = 'Não informado'
                cofins_vcofins = 'Não informado'

                for chave, valor in cofins.items():
                    if isinstance(valor, dict):
                        cofins_cst = valor.get('CST', cofins_cst)
                        cofins_vbc = valor.get('vBC', cofins_vbc)
                        cofins_pcofins = valor.get('pCOFINS', cofins_pcofins)
                        cofins_vcofins = valor.get('vCOFINS', cofins_vcofins)

                lista_insercao.append((nome_cliente, cnpj_cliente, fornecedor, cnpj_fornecedor, inscricao_estadual, numero_nota, serie, data_emissao, data_sai_ent, chave_nf, numero_item, cod_produto, ean, nome_produto, ncm, cfop,  quantidade,  valor_unitario,  valor_frete, valor_seguro, valor_desconto, valor_outros, valor_produto, ean_tributado, unidade_tributada, quantidade_tributada, valor_unitario_tributado, ind_tot, icms_original, icms_cst, icms_mod_bc, icms_vbc, icms_picms, icms_valor, pis_cst, pis_vbc, pis_ppis, pis_vpis, cofins_cst, cofins_vbc, cofins_pcofins, cofins_vcofins,  nome_do_arquivo))

                print(f"empresa: {nome_cliente}, cnpj_cliente: {cnpj_cliente}, fornecedor: {fornecedor}, cnpj_fornecedor:{cnpj_fornecedor}, ie:{inscricao_estadual}, numero_nota {numero_nota} serie {serie} data_emissao {data_emissao} data_sai_ent {data_sai_ent}  chave: {chave_nf}, , numero_item: {numero_item} cod_produto: {cod_produto}, ean:{ean} nome_produto: {nome_produto}, ncm: {ncm}, cfop:{cfop}  quantidade: {quantidade}, valor_unitario: {valor_unitario}, valor_frete: {valor_frete}, valor_seguro: {valor_seguro}, valor_desconto: {valor_desconto}, valor_outros: {valor_outros}, valor_produto: {valor_produto}, ean_tributado: {ean_tributado}, unidade_tributada:{unidade_tributada}  quantidade_tributada: {quantidade_tributada}, valor_unitario_tributado: {valor_unitario_tributado}, ind_tot:{ind_tot}, icms_original:{icms_original}, icms_cst:{icms_cst}, icms_mod_bc: {icms_mod_bc}, icms_vbc:{icms_vbc}, icms_picms:{icms_picms} , icms_valor{icms_valor}, nome_do_arquivo: {nome_do_arquivo}")


                valores.append([nome_cliente, cnpj_cliente, fornecedor, cnpj_fornecedor, inscricao_estadual, numero_nota, serie, data_emissao, data_sai_ent, chave_nf, numero_item, cod_produto, ean, nome_produto, ncm, cfop,  quantidade,  valor_unitario,  valor_frete, valor_seguro, valor_desconto, valor_outros, valor_produto, ean_tributado ,unidade_tributada ,quantidade_tributada, valor_unitario_tributado, ind_tot, icms_original, icms_cst, icms_mod_bc, icms_vbc, icms_picms, icms_valor, pis_cst, pis_vbc, pis_ppis, pis_vpis, cofins_cst, cofins_vbc, cofins_pcofins, cofins_vcofins,nome_do_arquivo])

            
    

            print("Transação concluída e dados inseridos com sucesso!")

        except Exception as e:
            # Rollback e log do erro
            # conexao.rollback()
            mensagem_aviso("Erro", f"Erro ao inserir dados: {e}")
            print(f"Erro ao inserir dados: {e}")

        finally:
            # Fechamento do cursor
            # cursor.close()
            print("Conexão encerrada.") # Fechar apenas o cursor

def selecionar_pasta():
    folder_path = QFileDialog.getExistingDirectory(None, "Selecione a pasta com os arquivos XML")
    if not folder_path:
        mensagem_aviso("Aviso", "Nenhuma pasta foi selecionada.")
        return

    lista_arquivos = [os.path.join(folder_path, arquivo) for arquivo in os.listdir(folder_path) if arquivo.endswith('.xml')]

    if not lista_arquivos:
        QMessageBox.warning(None, "Aviso", "A pasta selecionada não contém arquivos XML.")
        return

    colunas = ["nome_cliente", "cnpj_cliente", "fornecedor", "cnpj_fornecedor", "inscricao_estadual", "numero_nota", "serie", "data_emissao", "data_sai_ent", "chave_nf", "numero_item", "cod_produto", "ean", "nome_produto", "ncm", "cfop", "quantidade", "valor_unitario",  "valor_frete", "valor_seguro", "valor_desconto", "valor_outros", "valor_produto", "ean_tributado", "unidade_tributada", "quantidade_tributada", "valor_unitario_tributado", "ind_tot", "icms_original", "icms_cst", "icms_mod_bc", "icms_vbc", "icms_picms", "icms_valor", "pis_cst", "pis_vbc", "pis_ppis", "pis_vpis", "cofins_cst", "cofins_vbc", "cofins_pcofins", "cofins_vcofins", "nome_do_arquivo"]
    valores = []

    try:
        for arquivo in lista_arquivos:
            pegar_infos(arquivo, valores)
            
    finally:
        
        print("Conexão encerrada")

    tabela = pd.DataFrame(columns=colunas, data=valores)

    # Abrir diálogo para salvar o arquivo
    options = QFileDialog.Options()
    options |= QFileDialog.DontUseNativeDialog
    save_path, _ = QFileDialog.getSaveFileName(None, "Salvar arquivo Excel", "", "Arquivos Excel (*.xlsx);;Todos os Arquivos (*)", options=options)

    if not save_path:
        mensagem_aviso("Aviso", "Nenhum caminho de destino foi selecionado.")
        return

    # Garantir que o arquivo tenha a extensão correta
    if not save_path.endswith('.xlsx'):
        save_path += '.xlsx'

    tabela.to_excel(save_path, index=False)
    mensagem_aviso("Sucesso", f"Arquivo Excel gerado com sucesso em: {save_path}")
    print(f"Arquivo Excel gerado com sucesso em: {save_path}")
    return

def baixar_icone(url, caminho):
    dir_name = os.path.dirname(caminho)
    if not os.path.exists(dir_name):
        os.makedirs(dir_name)
    
    resposta = requests.get(url)
    with open(caminho, 'wb') as arquivo:
        arquivo.write(resposta.content)

def usar_icone(janela):
        url_icone = "https://assertivuscontabil.com.br/wp-content/uploads/2023/11/76.png"  
        caminho_icone = "images/icone.png" 
        
        baixar_icone(url_icone, caminho_icone)
        icon = Image.open(caminho_icone)
        icon.save(os.path.splitext(caminho_icone)[0] + '.ico', format="ICO")
        
        if os.path.exists(caminho_icone):
            if sys.platform == "win32":
                janela.setWindowIcon(QtGui.QIcon(os.path.splitext(caminho_icone)[0] + '.ico'))
            else:
                janela.setWindowIcon(QtGui.QIcon(caminho_icone))
        else:
            print(f"Erro: Arquivo de ícone não encontrado em {caminho_icone}")

def mensagem_aviso(titulo, texto):
    msg = QMessageBox()
    msg.setWindowTitle(titulo)
    msg.setText(texto)
    usar_icone(msg)
    msg.exec()

def main():
    app = QtWidgets.QApplication(sys.argv)

    janela = QtWidgets.QMainWindow()

    usar_icone(janela)
    
    janela.setWindowTitle("XML")
    janela.setGeometry(100, 100, 800, 600)

    widget_central = QtWidgets.QWidget()
    janela.setCentralWidget(widget_central)

    layout = QtWidgets.QVBoxLayout(widget_central)

    botao_frame = QtWidgets.QHBoxLayout()
    layout.addLayout(botao_frame)
    layout.addStretch()

    botoes = [
        ("Inserir nfes", lambda: selecionar_pasta()),
    ]

    for texto, funcao in botoes:
        botao = QtWidgets.QPushButton(texto)
        botao.clicked.connect(funcao)
        botao.setFont(QtGui.QFont("Arial", 14))
        botao.setStyleSheet("""
            QPushButton {
                background-color: #001F3F;
                color: white;
                border: none;
                padding: 5px 10px;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #2E236C;
            }
        """)
        botao.setCursor(QtGui.QCursor(QtCore.Qt.PointingHandCursor))
        botao_frame.addWidget(botao)

    


    janela.showMaximized()
    app.exec()

if __name__ == "__main__":
    main()
    