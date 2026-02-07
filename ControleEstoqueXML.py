import os
import shutil
import xml.etree.ElementTree as ET
import pandas as pd
from datetime import datetime
import tkinter as tk
from tkinter import messagebox

NS = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

BASE_DIR = os.getcwd()

PASTA_ENTRADA = os.path.join(BASE_DIR, 'entrada')
PASTA_SAIDA = os.path.join(BASE_DIR, 'saida')

PASTA_PROCESSADOS = os.path.join(BASE_DIR, 'processados')
PASTA_PROC_ENTRADA = os.path.join(PASTA_PROCESSADOS, 'entrada')
PASTA_PROC_SAIDA = os.path.join(PASTA_PROCESSADOS, 'saida')

PASTA_RESULTADO = os.path.join(BASE_DIR, 'resultado')
PASTA_ERROS = os.path.join(BASE_DIR, 'erros')

ARQ_ESTOQUE_GERAL = os.path.join(PASTA_RESULTADO, 'estoque_geral.xlsx')
ARQ_PROCESSADAS = os.path.join(PASTA_RESULTADO, 'processadas.xlsx')

for pasta in [
    PASTA_ENTRADA, PASTA_SAIDA,
    PASTA_PROC_ENTRADA, PASTA_PROC_SAIDA,
    PASTA_RESULTADO, PASTA_ERROS
]:
    os.makedirs(pasta, exist_ok=True)

def registrar_erro(msg):
    data = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
    with open(os.path.join(PASTA_ERROS, f'erro_{data}.txt'), 'w', encoding='utf-8') as f:
        f.write(str(msg))

def carregar_processadas():
    if os.path.exists(ARQ_PROCESSADAS):
        return set(pd.read_excel(ARQ_PROCESSADAS)['chNFe'])
    return set()

def salvar_processadas(chaves):
    pd.DataFrame({'chNFe': list(chaves)}).to_excel(ARQ_PROCESSADAS, index=False)

def carregar_estoque():
    if os.path.exists(ARQ_ESTOQUE_GERAL):
        return pd.read_excel(ARQ_ESTOQUE_GERAL)
    return pd.DataFrame(columns=['Codigo', 'Produto', 'Estoque_KG'])

def processar_pasta(pasta_origem, pasta_destino, tipo, fator, chaves_processadas):
    movimentos = []

    for arquivo in os.listdir(pasta_origem):
        if not arquivo.lower().endswith('.xml'):
            continue

        origem = os.path.join(pasta_origem, arquivo)

        try:
            tree = ET.parse(origem)
            root = tree.getroot()

            chNFe = root.find('.//nfe:chNFe', NS).text
            if chNFe in chaves_processadas:
                shutil.move(origem, os.path.join(pasta_destino, arquivo))
                continue

            for det in root.findall('.//nfe:det', NS):
                prod = det.find('nfe:prod', NS)

                movimentos.append({
                    'Data': datetime.now().strftime('%Y-%m-%d'),
                    'Tipo': tipo,
                    'Codigo': prod.find('nfe:cProd', NS).text,
                    'Produto': prod.find('nfe:xProd', NS).text,
                    'Quantidade_KG': float(prod.find('nfe:qCom', NS).text) * fator,
                    'chNFe': chNFe,
                    'Arquivo_XML': arquivo
                })

            chaves_processadas.add(chNFe)
            shutil.move(origem, os.path.join(pasta_destino, arquivo))

        except Exception as e:
            registrar_erro(f'Arquivo: {arquivo}\nErro: {e}')

    return movimentos

def main():
    try:
        chaves_processadas = carregar_processadas()
        estoque = carregar_estoque()

        movimentos = []
        movimentos += processar_pasta(PASTA_ENTRADA, PASTA_PROC_ENTRADA, 'Entrada', 1, chaves_processadas)
        movimentos += processar_pasta(PASTA_SAIDA, PASTA_PROC_SAIDA, 'Saída', -1, chaves_processadas)

        if movimentos:
            mov_df = pd.DataFrame(movimentos)

            # Excel do movimento do dia
            nome_mov = f"movimento_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
            mov_df.to_excel(os.path.join(PASTA_RESULTADO, nome_mov), index=False)

            # Atualiza estoque geral
            saldo = mov_df.groupby(['Codigo', 'Produto'])['Quantidade_KG'].sum().reset_index()
            estoque = pd.concat([estoque, saldo], ignore_index=True)
            estoque = estoque.groupby(['Codigo', 'Produto'])['Estoque_KG'].sum().reset_index()

            estoque.to_excel(ARQ_ESTOQUE_GERAL, index=False)
            salvar_processadas(chaves_processadas)

        messagebox.showinfo(
            'Concluído',
            'Processamento finalizado com sucesso!\n\n'
            'Foram gerados:\n'
            '- Movimento do dia\n'
            '- Estoque geral atualizado'
        )

    except Exception as e:
        registrar_erro(e)
        messagebox.showerror(
            'Erro',
            'Ocorreu um erro.\nVerifique a pasta "erros".'
        )

if __name__ == '__main__':
    root = tk.Tk()
    root.withdraw()
    main()
