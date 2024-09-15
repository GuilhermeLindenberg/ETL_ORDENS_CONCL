import pandas as pd
import time
from tqdm import tqdm
import warnings
import os
# Suprimir avisos de UserWarning
warnings.simplefilter("ignore", UserWarning)

# Função para exibir um cronômetro simples
def cronometro():
    start_time = time.time()
    while True:
        elapsed_time = time.time() - start_time
        mins, secs = divmod(elapsed_time, 60)
        timer = '{:02}:{:02}'.format(int(mins), int(secs))
        print(f'\rTempo decorrido: {timer}', end='', flush=True)
        time.sleep(1)

# Carregar a planilha Excel
file_path = r'C:\Users\guilherme.ferreira\Desktop\IRR\todas_ordens_de_serviço_concluídas..xlsx'
print("Carregando planilha...")
import threading
cronometro_thread = threading.Thread(target=cronometro, daemon=True)
cronometro_thread.start()
df = pd.read_excel(file_path)

# Iniciar o cronômetro em uma thread separada


# Filtrar as linhas da coluna 39 (considerando índice zero, é a coluna de índice 38)
siglas_desejadas = [
    'ALR', 'AFH', 'AHE', 'APER', 'CIM', 'CMBC', 'CPS', 'CTG', 'CCA', 'CGS',
    'CDI', 'DUS', 'GRI', 'GUU', 'ITIP', 'ILA', 'IOC', 'IEM', 'ITA', 'JOO',
    'LJM', 'MACC', 'MRZS', 'MCM', 'MRE', 'NTE', 'PMA', 'PCA', 'SDP', 'SEA',
    'SFD', 'SJUB', 'TOS', 'VIA', 'VVA', 'VTA'
]
print("Filtrando as linhas...")
df_filtered = df[df.iloc[:, 38].isin(siglas_desejadas)]

# Excluir as colunas indesejadas (considerando que o índice começa em 0)
colunas_a_excluir = [
    0, 1, 2, 3, 4, 8, 9, 10, 12, 19, 20, 24, 25, 28, 29, 33, 35, 36, 37
]
print("Excluindo as colunas...")
df_final = df_filtered.drop(df.columns[colunas_a_excluir], axis=1)

# Salvar o resultado em um novo arquivo Excel
output_path = 'todas_ordens_de_serviço_concluídas..xlsx'
print("Salvando o arquivo final...")
df_final.to_excel(output_path, index=False)

# Finalizar o cronômetro
print("\nProcesso finalizado!")
