import tkinter as tk                           # importando a biblioteca inteira do tkinter com nome de tk
import pandas as pd
import requests
import pandas
from tkinter import ttk                        # importando o ttk para pegar o combobox
from tkinter.filedialog import askopenfilename # importando biblioteca de abrir arquivo no tkinter
from tkcalendar import DateEntry               # importando calendario
from datetime import datetime

# pegando API para atualizar as moedas em tempo real
requisicao = requests.get('https://economia.awesomeapi.com.br/json/all') # link do API das moedas
dicionario_moedas = requisicao.json() # pegando o dicionario com todas as moedas
lista_moedas = list(dicionario_moedas.keys())  # criando a lista de moedas com base nas keys() do dicionário
def valor_moeda():
    try:
        moedas = moeda.get()  # pegando nome da moeda
        data_cotacao = calendario_moeda.get()  # pegando a data da moeda no calendario
        ano = data_cotacao[-4:]  # pegando apenas o ano no formato xx/xx/xxxx
        mes = data_cotacao[3:5]  # pegando apenas o mês no formato xx/xx/xxxx
        dia = data_cotacao[:2]  # pegando apenas o dia no formato xx/xx/xxxx
        link = f'https://economia.awesomeapi.com.br/json/daily/{moedas}--BRL/' \
               f'?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}'
        requisicao_moeda = requests.get(link)  # vendo as informações da moeda em um dicionário no link
        cotacao = requisicao_moeda.json()  # pegando as informações da moeda do link
        valor_moeda = cotacao[0]['bid']  # pegando o valor da moeda
        valor_moeda = float(valor_moeda)  # transformando o valor em float para poder formatar no print
        nome_moeda = cotacao[0]['name']  # pegando nome da moeda
        nome_moeda = nome_moeda.split('/')  # transformando em um dicionário de nomes
        nome_moedaentrada = nome_moeda[0]  # pegando o nome do dicionário da moeda escolhida
        valor_entry = valor_converter.get()
        valor_entry = float(valor_entry)
        resultado = valor_entry * valor_moeda
        mensagem_converter['text'] = f' {valor_entry:.2f} {nome_moedaentrada} para Real Brasileiro: R${resultado:,.2f}'
        mensagem_dia['text']= f'{dia}/{mes}/{ano}'
    except:
        mensagem_converter['text'] = 'Selecione uma moeda e um valor correto (00.00)'

# criando a função do botão para ele funcionar
def buscar_cotacao():
    try:
         moedas = moeda.get()  # pegando nome da moeda
         data_cotacao = calendario_moeda.get() # pegando a data da moeda no calendario
         ano = data_cotacao[-4:]  # pegando apenas o ano no formato xx/xx/xxxx
         mes = data_cotacao[3:5]  # pegando apenas o mês no formato xx/xx/xxxx
         dia = data_cotacao[:2]   # pegando apenas o dia no formato xx/xx/xxxx
         link = f'https://economia.awesomeapi.com.br/json/daily/{moedas}--BRL/' \
                f'?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}'
         requisicao_moeda = requests.get(link) # vendo as informações da moeda em um dicionário no link
         cotacao = requisicao_moeda.json() # pegando as informações da moeda do link
         valor_moeda = cotacao[0]['bid'] # pegando o valor da moeda
         valor_moeda = float(valor_moeda) # transformando o valor em float para poder formatar no print
         nome_moeda = cotacao[0]['name'] # pegando nome da moeda
         nome_moeda = nome_moeda.split('/')  # transformando em um dicionário de nomes
         nome_moedaentrada = nome_moeda[0] # pegando o nome do dicionário da moeda escolhida
         mensagem_cotacao['text'] = f'A cotação da {nome_moedaentrada} no dia {data_cotacao} foi de: R${valor_moeda:,.2f}'
    except:
         mensagem_cotacao['text'] = 'Selecione uma moeda'

def selecionar_arquivo():
    caminho_arquivo = askopenfilename(title='Selecione o Arquivo de Moeda')
    var_caminhoarquivo.set(caminho_arquivo)
    if calendario_moeda:
        label_arquivoselecionado['text'] = f'Arquivo Selecionado: {caminho_arquivo}'

def atualizar_cotacoes():
#try:
    df = pd.read_excel(var_caminhoarquivo.get())# ler o dataframe de moedas
    moedas = df.iloc[:, 0]
    data_inicial = calendario_datainicial.get()# pegar a data de inicio das cotações
    data_final = calendario_datafinal.get() # pegar a data final das cotações
    ano_inicial = data_inicial[-4:]  # pegando apenas o ano no formato xx/xx/xxxx
    mes_inicial = data_inicial[3:5]  # pegando apenas o mês no formato xx/xx/xxxx
    dia_inicial = data_inicial[:2]  # pegando apenas o dia no formato xx/xx/xxxx

    ano_final = data_final[-4:]  # pegando apenas o ano no formato xx/xx/xxxx
    mes_final = data_final[3:5]  # pegando apenas o mês no formato xx/xx/xxxx
    dia_final = data_final[:2]  # pegando apenas o dia no formato xx/xx/xxxx

    for moedas in moeda:    # para cada moeda
        link = f'https://economia.awesomeapi.com.br/json/daily/{moedas}--BRL/' \
               f'?start_date={ano_inicial}{mes_inicial}{dia_inicial}&end_date={ano_final}{mes_final}{dia_final}'# vendo as informações da moeda em um dicionário no link
        requisicao_moeda = requests.get(link)
        cotacoes = requisicao_moeda.json()  # pegando as informações da moeda do link
        for cotacao in cotacoes:
            timestamp = int(cotacao['timestamp'])
            bid = float(cotacao['bid'])
            data = datetime.fromtimestamp(timestamp)
            data = data.strftime('%d/%m/%Y')
            if data not in df:
                df[data] = np.nan
            df.loc[df.iloc[:, 0] == moedas, data] = bid

    df.to_excel('CotaçãoMoedas.xlsx')
    label_atualizarcotacoes['text'] = 'Arquivo Atualizado com Sucesso'
#except:
    #label_atualizarcotacoes['text'] = 'Selecione um arquivo Excel no Formato Correto'

janela = tk.Tk()     # criando janela tkinter com o Tk()
janela.title('Cotação de Moedas') # título da janela


                                        # COTAÇÃO DE 1 MOEDA

# criando mensagens com .Label(text='')
mensagem = tk.Label(text='COTAÇÃO DE 1 MOEDA ESPECÍFICA',borderwidth=2, relief='solid',font='Arial 10',background='#dbdbdb' )
mensagem.grid(row=0, column=0, padx=10, pady=10, columnspan=3, sticky='NSEW')  # exibindo e organizando com .grid()

label_selecionarmoeda = tk.Label(text='Selecione a moeda', anchor='e')
label_selecionarmoeda.grid(row=1, column=0, pady=10, padx=10, columnspan=2, sticky='NSEW')

# criando um combobox, onde seleciona o que quer ao invés de escrever o nome
moedas = list(lista_moedas)  # transformando o dicionário em uma lista
moeda = ttk.Combobox(janela, values=moedas) # criando o combox
moeda.grid(row=1, column=2, padx=10, pady=10, sticky='nsew')

mensagem_data = tk.Label(text='Selecione o dia que deseja pegar a cotação (2016+)', anchor='e')
mensagem_data.grid(row=2, column=0, pady=10, padx=10, columnspan=2, sticky='NSEW')  # exibindo e organizando com .grid()

calendario_moeda = DateEntry(year=2023, locale='pt_br')
calendario_moeda.grid(row=2, column=2 ,pady=10, padx=10, sticky='NSEW')

botao_buscarcotacao = tk.Button(text='Pesquisar cotação', command=buscar_cotacao)
botao_buscarcotacao.grid(row=3, column=2, padx=10, pady=10, sticky='nsew')

mensagem_cotacao = tk.Label(text='', borderwidth=1, relief='groove')
mensagem_cotacao.grid(row=3, column=0, columnspan=2, pady=10, sticky='nsew')



                                        # VALOR DA MOEDA EM REAIS

mensagem_valor = tk.Label(text='VALOR DA MOEDA EM REAIS',font='Arial 10', borderwidth=1 , relief='solid')
mensagem_valor.grid(row=4, column=0, padx=10, pady=10, columnspan=3, sticky='NSEW')  # exibindo e organizando com .grid()

botao_buscarcotacao = tk.Button(text='Converter', command=valor_moeda)
botao_buscarcotacao.grid(row=5, column=2, padx=10, pady=10, sticky='nsew')

valor_converter = tk.Entry(text='Insira um valor', borderwidth=1, relief='groove')
valor_converter.grid(row=5, column=1, padx=10, pady=10, sticky='nsew')

mensagem_entry = tk.Label(text='Insira um valor a ser convertido:', anchor='e')
mensagem_entry.grid(row=5, column=0, padx=10, pady=10, sticky='nsew')

mensagem_converter = tk.Label(text='', borderwidth=1, relief='groove')
mensagem_converter.grid(row=6, column=1,columnspan=2, padx=10, pady=10, sticky='nsew')

mensagem_dia = tk.Label(text='00/00/0000', borderwidth=1, relief='groove')
mensagem_dia.grid(row=6, column=0, padx=10, pady=10, sticky='nsew')

                                        # COTAÇÃO VARIAS MOEDAS

mensagem_variascotacoes = tk.Label(text='COTAÇÃO DE MULTIPLAS MOEDAS',borderwidth=2, relief='solid', font='Arial 10',background='#dbdbdb')
mensagem_variascotacoes.grid(row=7, column=0,padx=10, pady=10, columnspan=3, sticky='NSEW')  # exibindo e organizando com .grid()

mensagem_arquivo = tk.Label(text='Selecione um arquivo em Excel com as Moedas na COLUNA A')
mensagem_arquivo.grid(row=8, column=0, sticky='NSEW',columnspan=2, pady=10, padx=10)

var_caminhoarquivo = tk.StringVar()

botao_arquivo = tk.Button(text='Selecionar', command=selecionar_arquivo)
botao_arquivo.grid(row=8, column=2, pady=10, padx=10, sticky='NSEW')

label_arquivoselecionado = tk.Label(text='Nenhum arquivo selecionado', anchor='e', borderwidth=1, relief='groove')
label_arquivoselecionado.grid(row=9,column=0, columnspan=3, pady=10, padx=10, sticky='NSEW')

label_datainicial = tk.Label(text='Data Inicial', anchor='e')
label_datainicial.grid(row=10, column=0, padx=10, pady=10, sticky='NSEW')

label_datafinal = tk.Label(text='Data Final', anchor='e')
label_datafinal.grid(row=11, column=0, padx=10, pady=10, sticky='NSEW')

calendario_datainicial = DateEntry(year=2023, locale='pt_br')
calendario_datainicial.grid(row=10, column=1, pady=10, padx=10, sticky='NSEW')

calendario_datafinal = DateEntry(year=2023, locale='pt_br')
calendario_datafinal.grid(row=11, column=1, pady=10, padx=10, sticky='NSEW')

botao_atualizarcotacoes = tk.Button(text='Atualizar Cotações', command=atualizar_cotacoes)
botao_atualizarcotacoes.grid(row=12, column=0, padx=10, pady=10, sticky='NSEW')

label_atualizarcotacoes = tk.Label(text='')
label_atualizarcotacoes.grid(row=12, column=1, columnspan=2, padx=10, pady=10, sticky='NSEW')

botao_fechar = tk.Button(text='Fechar', command=janela.quit)  # botao de fechar aplicação
botao_fechar.grid(row=13,column=2, padx=10, pady=10, sticky='NSEW')

janela.mainloop()    # deixar a janela aberta, por loop com .mainloop()

