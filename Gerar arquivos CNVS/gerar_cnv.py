from docxtpl import DocxTemplate
from datetime import datetime
import pandas as pd
import os
import time


# =========================
# CONFIGURAÇÕES
# =========================


ARQUIVO_EXCEL = "vigilantes.xlsx"
MODELO_WORD = "modelo.docx"
PASTA_SAIDA = "documentos_gerados"
CIDADE = "Salvador"


# IP DA IMPRESSORA
IP_IMPRESSORA = "192.168.15.21"


# =========================
# CRIAR PASTA DE SAÍDA
# =========================


os.makedirs(PASTA_SAIDA, exist_ok=True)


# =========================
# DATA ATUAL
# =========================


meses = {
    1: "janeiro",
    2: "fevereiro",
    3: "março",
    4: "abril",
    5: "maio",
    6: "junho",
    7: "julho",
    8: "agosto",
    9: "setembro",
    10: "outubro",
    11: "novembro",
    12: "dezembro"
}


agora = datetime.now()


data_hoje = f"{agora.day} de {meses[agora.month]} de {agora.year}"


# =========================
# LISTA DE ARQUIVOS
# =========================


arquivos_gerados = []


# =========================
# LEITURA DA PLANILHA
# =========================


try:


    dados = pd.read_excel(ARQUIVO_EXCEL)


    dados.columns = dados.columns.str.strip().str.lower()


    dados = dados.apply(
        lambda x: x.str.strip() if x.dtype == "object" else x
    )


except Exception as erro:


    print(f"Erro ao ler planilha: {erro}")
    exit()


# =========================
# GERAÇÃO DOS DOCUMENTOS
# =========================


for index, linha in dados.iterrows():


    try:


        doc = DocxTemplate(MODELO_WORD)


        contexto = {
            "nome": linha["nome"],
            "turma": linha["turma"],
            "posto": linha["posto"],
            "data": data_hoje,
            "cidade": CIDADE
        }


        doc.render(contexto)


        nome_limpo = (
            str(linha["nome"])
            .replace("/", "-")
            .replace("\\", "-")
            .replace(":", "-")
            .replace("*", "")
            .replace("?", "")
            .replace('"', "")
            .replace("<", "")
            .replace(">", "")
            .replace("|", "")
        )


        nome_arquivo = os.path.abspath(
    os.path.join(PASTA_SAIDA, f"CNV_{nome_limpo}.docx")
)


        doc.save(nome_arquivo)


        arquivos_gerados.append(nome_arquivo)


        print(f"Documento gerado: {nome_arquivo}")


    except Exception as erro:


        print(f"Erro ao gerar documento: {erro}")


import socket


def verificar_impressora(ip, porta=9100):


    try:
        socket.create_connection((ip, porta), timeout=5)
        return True


    except:
        return False




# =========================
# IMPRESSÃO
# =========================


resposta = input("\nDeseja imprimir os documentos? (s/n): ")


if resposta.lower() == "s":


    print("\nVerificando impressora...\n")


    if verificar_impressora(IP_IMPRESSORA):


        print("Impressora conectada!")


        for arquivo in arquivos_gerados:


            try:


                os.startfile(arquivo, "print")


                print(f"Imprimindo: {arquivo}")


                time.sleep(2)


            except Exception as erro:


                print(f"Erro ao imprimir: {erro}")


    else:


        print("\nImpressora offline.")

