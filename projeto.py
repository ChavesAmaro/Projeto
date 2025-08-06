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


def formatar_tempo(minutos):
    dias = minutos // 1440
    resto = minutos % 1440
    horas = resto // 60
    mins = resto % 60
    return f"{dias} dia{'s' if dias != 1 else ''} {horas} hora{'s' if horas != 1 else ''} {mins} minuto{'s' if mins != 1 else ''}"

def tempo_estimado_leitura_formatado():
    global df
    if df.empty:
        print("\nA lista está vazia!")
        return

    
    print("\nEscolha um livro para ver o tempo estimado de leitura:\n")
    for i, row in df.iterrows():
        print(f"{i + 1}. {row['Título']} ({row['Autor']})")

    try:
        escolha = int(input("\nDigite o número do livro: ")) - 1
        if 0 <= escolha < len(df):
            paginas = df.loc[escolha, "Páginas"]
            titulo = df.loc[escolha, "Título"]
            autor = df.loc[escolha, "Autor"]
            tempo_min = formatar_tempo(paginas * 1)
            tempo_medio = formatar_tempo(paginas * 2)
            tempo_max = formatar_tempo(paginas * 3)

            print(f"\nTempo estimado de leitura para '{titulo}' ({autor}):")
            print(f"  → Mínimo: {tempo_min}")
            print(f"  → Médio: {tempo_medio}")
            print(f"  → Máximo: {tempo_max}\n")
        else:
            print("Número inválido.")
    except ValueError:
        print("Entrada inválida. Por favor, insira um número.")


    
def remover_livro():
    global df
    if df.empty:
        print("Lista vazio! ")
        return
    consultar_lista()
    
    try:
        linha = int(input("Indique o livro a remover: "))
        
        if linha not in df.index:
            print(f"Erro! O índice '{linha}' é inválido!")
            return None
        
        df.drop(index = linha, inplace=True)
        print(f"Linha {linha} removida com sucesso! ")
    except ValueError:
        print("Erro! O índice da linha deve ser um número inteiro! ")
    print()

def editar_valor():
    global df
    if df.empty:
        print("Lista vazia! ")
        return
    consultar_lista()
    try:
        linha = int(input("Indique o id da linha a alterar: "))
        nome_coluna = input("Indique em qual coluna pretende alterar o dado: ")

        if linha not in df.index:
            print(f"Erro! O índice '{linha}' é inválido!")
            return None
        
        if nome_coluna not in df.columns:
            print(f"Erro! A coluna '{nome_coluna}' não está registada!")
            return None
        
        tipo = df[nome_coluna].dtype
        
        while True:
            novo_dado = input(f"Indique o novo dado para o índice {linha}, na coluna {nome_coluna} ({tipo}): ").strip()
            try:
                if tipo == 'float64':
                    dado_convertido = float(novo_dado)
                elif tipo == 'int64':
                    dado_convertido = int(novo_dado)
                elif tipo == 'datetime64[ns]':
                    dado_convertido = pd.to_datetime(novo_dado, dayfirst=True)
                else:
                    dado_convertido = novo_dado
                break
            except ValueError:
                print(f"Erro! O dado inserido não corresponde ao tipo esperado ({tipo})")

        df.at[linha, nome_coluna] = dado_convertido
        print("Dados atualizados com sucesso! ")
    except ValueError:
        print("Erro! O id da linha deve ser um número inteiro! ")
    print()

def filtro_lista():
    global df
    if df.empty:
        print("Lista vazia! ")
        return
    
    print("\nColunas da Lista: "," | ".join(df.columns))
    
    nome_coluna = input("\nIntroduza o nome da coluna onde pretende filtrar dados: ")
    
    if nome_coluna not in df.columns:
        print(f"Erro! A coluna {nome_coluna} não existe na lista!")
        return None
    
    if nome_coluna in ["Data do Início da Leitura", "Data do Fim da Leitura"]:
        try:
            df[nome_coluna] = df[nome_coluna].astype(str)
            df[nome_coluna] = pd.to_datetime(df[nome_coluna], format="%d/%m/%Y", errors="coerce")
            intervalo = input(f"Introduza duas datas (ex: 01/01/2023 31/12/2023): ")
            data1, data2 = map(lambda x: pd.to_datetime(x, dayfirst=True), intervalo.replace(",", " ").split())
            data_min, data_max = min(data1, data2), max(data1, data2)
            filtro = df[(df[nome_coluna] >= data_min) & (df[nome_coluna] <= data_max)]
        except:
            print("Erro! Introduza duas datas válidas no formato DD/MM/YYYY.")
            return

    elif nome_coluna in "Data de Registo":
        try:
            df[nome_coluna] = df[nome_coluna].astype(str)
            df[nome_coluna] = pd.to_datetime(df[nome_coluna], format="%d-%m-%Y %H:%M:%S", errors="coerce")
            intervalo = input(f"Introduza duas datas (ex: 01/01/2023 00:00:00,30/01/2023 00:00:00): ")
            data1, data2 = map(lambda x: pd.to_datetime(x, dayfirst=True), intervalo.replace(",", " ").split())
            data_min, data_max = min(data1, data2), max(data1, data2)
            filtro = df[(df[nome_coluna] >= data_min) & (df[nome_coluna] <= data_max)]
        except:
            print("Erro! Introduza duas datas válidas no formato DD/MM/YYYY HH:MM:SS")
            return
    
    elif nome_coluna == "Ano de Publicação":
        try:
            df[nome_coluna] = df[nome_coluna].astype(str)
            df[nome_coluna] = pd.to_datetime(df[nome_coluna], format="%Y", errors="coerce")
            intervalo = input(f"Introduza dois anos (ex: 2010 2020): ")
            ano1, ano2 = map(lambda x: pd.to_datetime(x, format="%Y"), intervalo.replace(",", " ").split())
            ano_min, ano_max = min(ano1, ano2), max(ano1, ano2)
            filtro = df[(df[nome_coluna] >= ano_min) & (df[nome_coluna] <= ano_max)]
        except:
            print("Erro! Introduza dois anos válidos no formato YYYY.")
            return
    
    elif pd.api.types.is_numeric_dtype(df[nome_coluna]):
        try:
            intervalo = input(f"Introduza dois valores (ex: 10 50) para filtrar '{nome_coluna}': ")
            val1, val2 = map(float, intervalo.replace(",", " ").split())
            minimo, maximo = min(val1, val2), max(val1, val2)
            filtro = df[(df[nome_coluna] >= minimo) & (df[nome_coluna] <= maximo)]
        except:
            print("Erro! Introduza dois valores numéricos válidos.")
            return

    else:
        critério = input(f"Indique o critério a filtrar na coluna '{nome_coluna}': ")       
        filtro = df[df[nome_coluna] == critério]
    
    filtro_formatado = filtro.copy()
    
    filtro_formatado["Data do Início da Leitura"] = filtro_formatado["Data do Início da Leitura"].apply(
        lambda x: x.strftime("%d/%m/%Y") if isinstance(x, datetime) else x)
    filtro_formatado["Data do Fim da Leitura"] = filtro_formatado["Data do Fim da Leitura"].apply(
        lambda x: x.strftime("%d/%m/%Y") if isinstance(x, datetime) else x)
    filtro_formatado["Ano de Publicação"] = filtro_formatado["Ano de Publicação"].apply(
        lambda x: x.strftime("%Y") if isinstance(x, datetime) else x)
    
    print(f"Dados na coluna: {nome_coluna} ")
    if filtro.empty:
        print("Nenhum resultado encontrado!")
    else:    
        print(filtro_formatado)   
    print()      
    

def main():
    loop_menu = True
    while loop_menu == True:
        print("\nMenu Principal")
        print("1. Adicionar Livro à Lista")
        print("2. Consultar Lista de Livros")
        print("3. Exportar a Lista Para Ficheiro .xlsx")
        print("4. Importar a Lista de Ficheiro .xlsx")
        print("5. Exportar Para .pdf")
        print("6. Calcular Tempo Estimado de Leitura")
        print("7. Remover Livro")
        print("8. Editar Valor")
        print("9. Filtrar Lista")
        print("10. Sair")

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
            tempo_estimado_leitura_formatado()
        elif escolha_menu_principal == '7':
            remover_livro()
        elif escolha_menu_principal == '8':
            editar_valor()
        elif escolha_menu_principal == '9':
            filtro_lista()
        elif escolha_menu_principal == '10':
            print("\nA encerrar o programa...")
            loop_menu = False
        else:
            print("Erro! Escolha um número de 1 a 6.") #mudar no final para o nº de opções

if __name__ == "__main__":
    print("\n\n")
    main()
    print("\n\n")
