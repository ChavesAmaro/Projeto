import pickle
from docx import Document
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd
from datetime import datetime
import requests
#import creds
from dotenv import load_dotenv
import os
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle

df = pd.DataFrame()

data_inicio = []
data_final = []
data_registo = []
autor = []
titulo = []
ano_publicacao = []
edicao = []
paginas = []
preco = []
tema_principal = []
avaliacao = []
link_google = []

def configure():
    load_dotenv()

def obter_link_google_books(titulo, autor):
    query = f"{titulo} {autor}"
    #url = f"https://www.googleapis.com/books/v1/volumes?q={query}&key={creds.api_key}"
    url = f"https://www.googleapis.com/books/v1/volumes?q={query}&key={os.getenv('api_key')}"
    response = requests.get(url)
    data = response.json()
    
    if "items" in data:
        livro = data["items"][0]["volumeInfo"]
        return livro.get("infoLink", "Link não disponível")
    else:
        return "Livro não encontrado"

def adicionar_lista():
    global df
    loop_registo = True
    while loop_registo == True:
        try:
            inicio_d = input("\nIndique quando começou a ler o livro (formato DD/MM/YYYY) ou deixe em branco caso ainda não tenha começado: ").strip()
            if inicio_d:
                data_i = datetime.strptime(inicio_d, "%d/%m/%Y")
            else:
                data_i = "Por ler"
            fim_d = input("Indique quando terminou de ler o livro (formato DD/MM/YYYY) ou deixe em branco caso ainda não tenha terminado: ").strip()
            if fim_d:
                data_f = datetime.strptime(fim_d, "%d/%m/%Y")
            elif data_i == "Por ler":
                data_f = "Por ler"
            else:
                data_f = "Em Leitura"
            data_r = datetime.now().strftime("%d-%m-%Y %H:%M:%S")
            autor_r = input("Indique o autor do livro a registar: ")
            titulo_r = input("Indique o título do livro: ")
            ano_p = input("Indique o ano de publicação do livro: ")
            ano_p_r = datetime.strptime(ano_p, "%Y")
            preco_r = float(input("Indique o preço do livro: "))
            edicao_r = int(input("Indique a edição do livro: "))
            paginas_r = int(input("Indique o número de páginas do livro: "))
            tema_r = input("Indique o tema principal do livro: ")
            
            loop_avaliacao = True
            while loop_avaliacao == True:
                try:
                    avaliacao_r = float(input("Indique a sua avaliação do livro (0-5): "))
                    
                    if 0 <= avaliacao_r <= 5:
                        loop_avaliacao = False
                    else:
                        print("Erro! A avaliação deve estar entre 0 e 5.")
                except ValueError:
                    print("Erro! Introduza um número real na avaliação.")
            
            data_inicio.append(data_i)
            data_final.append(data_f)
            data_registo.append(data_r)
            autor.append(autor_r)
            titulo.append(titulo_r)
            ano_publicacao.append(ano_p_r)
            preco.append(preco_r)
            edicao.append(edicao_r)
            paginas.append(paginas_r)
            tema_principal.append(tema_r)
            avaliacao.append(avaliacao_r)
            
            
            print("Livro registado com sucesso!")
        except ValueError:
            print("Erro! Formato dos valores errado!")
        
        link_google_books = obter_link_google_books(titulo_r, autor_r)
        link_google.append(link_google_books)

        loop_verificacao = True
        while loop_verificacao == True:
            continuar_registar = input("\nDeseja registar mais um livro? (s/n): ").strip().lower()
            
            if continuar_registar in ['s', 'n']:
                loop_verificacao = False
            else:
                print("Erro! Resposta inválida introduza 's' para sim ou 'n' para não.")
        
        if continuar_registar == 'n':
            loop_registo = False
    
    df_novo = pd.DataFrame({
        "Data do Início da Leitura": data_inicio,
        "Data do Fim da Leitura": data_final,
        "Data de Registo": data_registo,
        "Autor": autor,
        "Título": titulo,
        "Ano de Publicação": ano_publicacao,
        "Preço": preco,
        "Edicao": edicao,
        "Páginas": paginas,
        "Tema Principal": tema_principal,
        "Avaliação Pessoal": avaliacao,
        "Link Google Books": link_google
    })
    df = pd.concat([df, df_novo], ignore_index=True)

    data_inicio.clear()
    data_final.clear()
    data_registo.clear()
    autor.clear()
    titulo.clear()
    ano_publicacao.clear()
    edicao.clear()
    paginas.clear()
    preco.clear()
    tema_principal.clear()
    avaliacao.clear()
    link_google.clear()
      
def consultar_lista():
    global df
    if df.empty:
        print("\nA lista está vazia!")
        return
    
    df_formatado = df.copy()
    
    df_formatado["Data do Início da Leitura"] = df_formatado["Data do Início da Leitura"].apply(
        lambda x: x.strftime("%d/%m/%Y") if isinstance(x, datetime) else x)
    df_formatado["Data do Fim da Leitura"] = df_formatado["Data do Fim da Leitura"].apply(
        lambda x: x.strftime("%d/%m/%Y") if isinstance(x, datetime) else x)
    df_formatado["Ano de Publicação"] = df_formatado["Ano de Publicação"].apply(
        lambda x: x.strftime("%Y") if isinstance(x, datetime) else x)
    
    print("Lista de livros")
    print(df_formatado)
    print()
    
def exportar_excel():
    global df
    if df.empty:
        print("Erro! Lista vazia! ")
        return
    
    nome_ficheiro = input("Introduza o nome do ficheiro a guardar: ")
    if not nome_ficheiro.endswith('.xlsx'):
        nome_ficheiro += '.xlsx'
    
    try:
        df_formatado = df.copy()
        
        df_formatado["Data do Início da Leitura"] = df_formatado["Data do Início da Leitura"].apply(
        lambda x: x.strftime("%d/%m/%Y") if isinstance(x, datetime) else x)
        df_formatado["Data do Fim da Leitura"] = df_formatado["Data do Fim da Leitura"].apply(
        lambda x: x.strftime("%d/%m/%Y") if isinstance(x, datetime) else x)
        df_formatado["Ano de Publicação"] = df_formatado["Ano de Publicação"].apply(
        lambda x: x.strftime("%Y") if isinstance(x, datetime) else x)
        
        df_formatado.to_excel(nome_ficheiro, index=False)
        
        print(f"Dados exportados com sucesso para: {nome_ficheiro}")
    except Exception as erroexportexcel:
        print(f"Erro ao gerar ficheiro: {erroexportexcel}")
    print()
    
def importar_excel():
    global df
    
    nome_ficheiro = input("Introduza o nome do ficheiro a importar: ")
    
    if not nome_ficheiro.endswith('.xlsx'):
        nome_ficheiro += '.xlsx'
    
    try:
        df = pd.read_excel(nome_ficheiro)
        print(f"{nome_ficheiro} importado com sucesso!")
    except FileNotFoundError:
        print(f"{nome_ficheiro} não encontrado!")
    except Exception as erroimportexcel:
        print(f"Erro a importar o ficheiro: {erroimportexcel}")
    print()
    
def export_pdf():
    global df
    if df.empty:
        print("Erro! Lista vazia! ")
        return
    
    nome_arquivo = input("Introduza o nome do ficheiro PDF a guardar: ")
    if not nome_arquivo.endswith('.pdf'):
        nome_arquivo += '.pdf'
    
    
    doc = SimpleDocTemplate(nome_arquivo, pagesize=A4)
    elementos = []
    df_formatado = df.copy()
        
    df_formatado["Data do Início da Leitura"] = df_formatado["Data do Início da Leitura"].apply(
    lambda x: x.strftime("%d/%m/%Y") if isinstance(x, datetime) else x)
    df_formatado["Data do Fim da Leitura"] = df_formatado["Data do Fim da Leitura"].apply(
    lambda x: x.strftime("%d/%m/%Y") if isinstance(x, datetime) else x)
    df_formatado["Ano de Publicação"] = df_formatado["Ano de Publicação"].apply(
    lambda x: x.strftime("%Y") if isinstance(x, datetime) else x)
    
    dados = [df_formatado.columns.tolist()] + df.values.tolist()
    tabela = Table(dados)
    
    
    estilo = TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.lightblue),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.black),
        ('ALIGN', (0, 0), (-1, -1), 'LEFT'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica'),
        ('FONTSIZE', (0, 0), (-1, -1), 4),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 6),
        ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ])
    tabela.setStyle(estilo)

    elementos.append(tabela)
    doc.build(elementos)
    print(f"PDF criado com sucesso com o nome: {nome_arquivo}")

def main():
    loop_menu = True
    while loop_menu == True:
        print("\nMenu Principal")
        print("1. Adicionar Livro à Lista")
        print("2. Consultar Lista de Livros")
        print("3. Exportar a Lista Para Ficheiro .xlsx")
        print("4. Importar a Lista de Ficheiro .xlsx")
        print("5. Exportar para .pdf")
        print("6. Sair")

        escolha_menu_principal = input("\nEscolha uma opção de 1 a 6: ").strip() #mudar no final para o nº de opções
        
        if escolha_menu_principal == '1':
            adicionar_lista()
        elif escolha_menu_principal == '2':
            consultar_lista()
        elif escolha_menu_principal == '3':
            exportar_excel()
        elif escolha_menu_principal == '4':
            importar_excel()
        elif escolha_menu_principal == '5':
            export_pdf()
        elif escolha_menu_principal == '6':
            print("\nA encerrar o programa...")
            loop_menu = False
        else:
            print("Erro! Escolha um número de 1 a 6.") #mudar no final para o nº de opções

if __name__ == "__main__":
    print("\n\n")
    main()
    print("\n\n")
