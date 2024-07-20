import tkinter as tk
from datetime import datetime
from tkinter import ttk
from tkinter.filedialog import askopenfilename

import numpy as np
import pandas as pd
import requests
from tkcalendar import DateEntry

janela = tk.Tk()
janela.title("Ferramenta de cotação de moedas.")


requisicao = requests.get("https://economia.awesomeapi.com.br/json/all")
dicionario_moedas = requisicao.json()

lista_moedas = list(dicionario_moedas.keys())


def pegar_cotacao():
    moeda = combobox_selecionarmoedas.get()
    data_cotacao = calendario_moeda.get()
    dia = data_cotacao[:2]
    mes = data_cotacao[3:5]
    ano = data_cotacao[6:]
    link = f"https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}"
    requisicao_moeda = requests.get(link)
    cotacao = requisicao_moeda.json()
    valor_moeda = cotacao[0]["bid"]
    label_textocotacao["text"] = (
        f"A cotação da {moeda} no dia {data_cotacao} foi de: R${valor_moeda}"
    )


def selecionar_arquivo():
    caminho_arquivo = askopenfilename(title="Selecione o Arquivo de Moeda")
    var_caminhoarquivo.set(caminho_arquivo)
    if caminho_arquivo:
        label_arquivoselecionado["text"] = f"Arquivo Selecionado: {caminho_arquivo}"
    else:
        label_arquivoselecionado["text"] = "Nenhum arquivo selecionado"


def atualizar_cotacoes():
    try:
        df = pd.read_excel(var_caminhoarquivo.get())
        moedas = df.iloc[:, 0]
        data_inicial = calendario_datainicial.get()
        dia_inicial = data_inicial[:2]
        mes_inicial = data_inicial[3:5]
        ano_inicial = data_inicial[6:]

        data_final = calendario_datafinal.get()
        dia_final = data_final[:2]
        mes_final = data_final[3:5]
        ano_final = data_final[6:]

        for moeda in moedas:
            link = f"https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?start_date={ano_inicial}{mes_inicial}{dia_inicial}&end_date={ano_final}{mes_final}{dia_final}"
            requisicao_moeda = requests.get(link)
            cotacoes = requisicao_moeda.json()
            for cotacao in cotacoes:
                timestamp = int(cotacao["timestamp"])
                bid = float(cotacao["bid"])
                data = datetime.fromtimestamp(timestamp)
                data = data.strftime("%d/%m/%Y")
                if data not in df:
                    df[data] = np.nan

                df.loc[df.iloc[:, 0] == moeda, data] = bid
        df.to_excel("teste.xlsx")
        label_atualizacotacoes["text"] = "Arquivo atualizado com sucesso!"
    except:
        label_atualizacotacoes["text"] = "Houve um erro, arquivo no formato errado!"


label_cotacaomoeda = tk.Label(
    text="Cotação de 1 moeda especifica",
    borderwidth=2,
    relief="solid",
    padx=10,
    pady=10,
    fg="white",
    bg="black",
)
label_cotacaomoeda.grid(row=0, column=0, padx=10, pady=10, sticky="nswe", columnspan=3)


# label do combobox
label_selecionarmoeda = tk.Label(text="Selecionar Moeda", anchor="e")
label_selecionarmoeda.grid(
    row=1, column=0, padx=10, pady=10, sticky="nswe", columnspan=2
)

# combobox, cascata de opção
combobox_selecionarmoedas = ttk.Combobox(values=lista_moedas)
combobox_selecionarmoedas.grid(row=1, column=2, padx=10, pady=10, sticky="nswe")

# label para o calendario
label_selecionardia = tk.Label(
    text="Selecione o dia que deseja pegar a cotação", anchor="e"
)
label_selecionardia.grid(row=2, column=0, padx=10, pady=10, sticky="nswe", columnspan=2)


# calendario
calendario_moeda = DateEntry(year=2024, locale="pt_br")
calendario_moeda.grid(row=2, column=2, padx=10, pady=10, sticky="nswe")

# label texto cotacao
label_textocotacao = tk.Label(text="")
label_textocotacao.grid(row=3, column=0, columnspan=2, padx=10, pady=10, sticky="nswe")

botao_pegarcotacao = tk.Button(text="Pegar cotação", command=pegar_cotacao)
botao_pegarcotacao.grid(row=3, column=2, padx=10, pady=10, sticky="nswe")

# cotação várias moedas
label_cotacaomoeda = tk.Label(
    text="Cotação de multiplas moedas",
    borderwidth=2,
    relief="solid",
    padx=10,
    pady=10,
    bg="black",
    fg="white",
)
label_cotacaomoeda.grid(row=4, column=0, padx=10, pady=10, sticky="nswe", columnspan=3)

# label selecionar arquivo
label_selecionararquivo = tk.Label(
    text="Selecione um arquivo em Excel com as Moedas na Coluna A", anchor="e"
)
label_selecionararquivo.grid(
    row=5, column=0, columnspan=2, padx=10, pady=10, sticky="nswe"
)

var_caminhoarquivo = tk.StringVar()

botao_selecionararquivo = tk.Button(
    text="Clique aqui para selecionar", command=selecionar_arquivo
)
botao_selecionararquivo.grid(row=5, column=2, padx=10, pady=10, sticky="nswe")


label_arquivoselecionado = tk.Label(text="Nenhum arquivo selecionado", anchor="e")
label_arquivoselecionado.grid(
    row=6, column=0, columnspan=3, padx=10, pady=10, sticky="nswe"
)
label_datainicial = tk.Label(text="Data inicial", anchor="e")
label_datafinal = tk.Label(text="Data final", anchor="e")
label_datainicial.grid(row=7, column=0, padx=10, pady=10, sticky="nswe")
label_datafinal.grid(row=8, column=0, padx=10, pady=10, sticky="nswe")


calendario_datainicial = DateEntry(
    year=2024,
    locale="pt_br",
)
calendario_datafinal = DateEntry(year=2024, locale="pt_br")
calendario_datainicial.grid(row=7, column=1, padx=10, pady=10, sticky="nswe")
calendario_datafinal.grid(row=8, column=1, padx=10, pady=10, sticky="nswe")


botao_atualizarcotacoes = tk.Button(
    text="Atualizar Cotações", command=atualizar_cotacoes
)

botao_atualizarcotacoes.grid(row=9, column=0, padx=10, pady=10, sticky="nsew")


label_atualizacotacoes = tk.Label(text="")
label_atualizacotacoes.grid(
    row=9, column=1, columnspan=2, padx=10, pady=10, sticky="nsew"
)

botao_fechar = tk.Button(text="Fechar", command=janela.quit)
botao_fechar.grid(row=10, column=2, padx=10, pady=10, columnspan=3, sticky="nsew")

janela.mainloop()
