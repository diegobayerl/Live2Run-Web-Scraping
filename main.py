import requests
from requests.adapters import HTTPAdapter
from requests.packages.urllib3.util.retry import Retry
from bs4 import BeautifulSoup
import pandas as pd
from urllib.parse import urlparse, parse_qs
import tkinter as tk
from tkinter import ttk
from ttkthemes import ThemedTk
import time

# Definindo as variáveis
base_url = "https://site.com.br"
page_part = "calendario.asp"
headers = []

# Função para fazer a requisição com retry
def get_with_retry(url, retries=3):
    session = requests.Session()
    retry = Retry(total=retries, backoff_factor=1, status_forcelist=[500, 502, 503, 504])
    adapter = HTTPAdapter(max_retries=retry)
    session.mount('http://', adapter)
    session.mount('https://', adapter)
    
    try:
        response = session.get(url)
        response.raise_for_status()
        return response
    except requests.RequestException as e:
        print(f"Erro ao acessar a URL: {e}")
        return None

# Função para extrair o número do link
def extract_number_from_link(link):
    parsed_url = urlparse(link)
    query_params = parse_qs(parsed_url.query)
    return query_params.get('escolha', [None])[0]

def gerar(go_part):
    rows = []
    page_number = 0

    while True:
        progresso['value'] += 10
        janela.update_idletasks()
        time.sleep(0.1)
        
        page_part = f"calendario{1 * page_number}.asp" if page_number > 0 else "calendario.asp"
        url = f"{base_url}/{go_part}/{page_part}"
        print(f"Processando página: {url}")
        
        data = get_with_retry(url)
        if data:
            soup = BeautifulSoup(data.text, 'html.parser')
            tables = [table for table in soup.find_all('table') if table.get('height') != '40']
            
            if len(tables) >= 8:
                table_8 = tables[7]
                headers = [header.text.strip() for header in table_8.find_all('th')]
                
                for row in table_8.find_all('tr'):
                    if (not row.find('td', attrs={'colspan': True}) and 
                        not any(cell.text.strip() == "Data" for cell in row.find_all('td')) and
                        not any(cell.text.strip() == "Próximas Corridas:" for cell in row.find_all('td'))):
                        
                        data = [cell.text.strip() for cell in row.find_all('td')]
                        original_link = row.find_all('td')[2].find('a')['href'] if len(row.find_all('td')) >= 3 and row.find_all('td')[2].find('a') else None
                        number = f"https://site.com.br/siteevento.asp?c={extract_number_from_link(original_link)}" if original_link else None
                        if data:
                            data.append(number)
                            data.append(go_part.upper())  # Adiciona a sigla do estado na coluna UF
                            rows.append(data)
                
                if headers:
                    headers.extend(["Link", "UF"])  # Adiciona "Link" e "UF" aos cabeçalhos
                else:
                    headers = [f"Column {i+1}" for i in range(len(rows[0]))]
            else:
                print("Menos de oito tabelas foram encontradas na página.")
            
            page_number += 1
        else:
            print(f"Falha ao acessar a página {url}. Parando a execução.")
            break

    return rows, headers

# Função para gerar e salvar os dados de todos os estados
def gerar_todos():
    all_data = []
    for estado in estados[1:]:  # Ignorando o primeiro item "Todos"
        rows, headers = gerar(estado)
        all_data.extend(rows)
    
    # Salvando todos os estados no mesmo arquivo Excel
    df = pd.DataFrame(all_data, columns=headers)
    df.to_excel('Tabela-Todos-Estados.xlsx', index=False)
    print("Tabela completa foi salva com sucesso em Tabela-Todos-Estados.xlsx")

# Função para ser executada quando o botão for clicado
def imprimir_texto():
    estado_selecionado = lista_estados.get()
    if estado_selecionado == "Todos":
        gerar_todos()
        progresso['value']=100
    else:
        go_part = estado_selecionado
        rows, headers = gerar(go_part)
        df = pd.DataFrame(rows, columns=headers)
        df.to_excel(f'Tabela-{go_part.upper()}.xlsx', index=False)
        print(f"Tabela foi salva com sucesso em Tabela-{go_part.upper()}.xlsx")
        progresso['value']=100

# Cria a janela principal com o tema aplicado
janela = ThemedTk(theme="arc")
janela.title("Live4Run")
janela.geometry("420x150")
janela.configure(bg="white")

# Estilos personalizados
estilo = ttk.Style()
estilo.configure("TLabel", font=("Helvetica", 12), background="white")
estilo.configure("TButton", font=("Helvetica", 12, "bold"), foreground="#333333", borderwidth=2, relief="solid")
estilo.map("TButton", foreground=[("active", "#333333")], background=[("active", "#cccccc")], bordercolor=[("active", "#333333")])
estilo.configure("TCombobox", font=("Helvetica", 12), background="#ffffff", padding=10)
estilo.configure("TProgressbar", thickness=15)

# Lista de estados para a lista suspensa
estados = ["Todos", "ac", "al", "am", "ap", "ba", "ce", "df", "es", "go", "ma", "mg", "ms", "mt", "pa", 
           "pb", "pe", "pi", "pr", "rj", "rn", "ro", "rr", "rs", "sc", "se", "sp", "to"]

# Cria uma lista suspensa (Combobox) com os estados
lista_estados = ttk.Combobox(janela, values=estados, width=27, style="TCombobox")
lista_estados.grid(row=0, column=0, padx=10, pady=20, sticky='w')
lista_estados.current(0)

# Cria um botão que chama a função imprimir_texto quando clicado
botao = ttk.Button(janela, text="OK", command=imprimir_texto, style="TButton")
botao.grid(row=0, column=1, padx=10, pady=20, sticky='w')

# Cria uma barra de progresso (Progressbar)
progresso = ttk.Progressbar(janela, orient='horizontal', length=350, mode='determinate', style="TProgressbar")
progresso.grid(row=1, column=0, columnspan=3, pady=20)

# Inicia o loop principal da interface
janela.mainloop()
