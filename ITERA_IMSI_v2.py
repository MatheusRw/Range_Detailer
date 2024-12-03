import pandas as pd
import csv

# Nome do arquivo Excel de entrada
input_file = "C:\\Users\\matheus.weinert\\Downloads\\BASE_IMSI_VIRTUEYES_V2.xlsx"  # Substitua pelo caminho do seu arquivo

# Nome do arquivo de saída
output_file = "imsi_iccid_ranges_virtueyes2.csv"

# Ler o arquivo Excel
df = pd.read_excel(input_file)

# Garantir que o arquivo contém as colunas corretas
# As colunas devem ser: qtd, IMSI START, IMSI END, ICCID START, ICCID END
if not {"qtd", "IMSI START", "IMSI END", "ICCID START", "ICCID END"}.issubset(df.columns):
    raise ValueError("O arquivo de entrada deve conter as colunas: 'qtd', 'IMSI START', 'IMSI END', 'ICCID START', 'ICCID END'")

# Converter os dados do DataFrame em uma lista de tuplas
data = df[["qtd", "IMSI START", "IMSI END", "ICCID START", "ICCID END"]].values.tolist()

# Função para gerar o intervalo de IMSIs e ICCIDs
def generate_ranges(data):
    for qtd, imsi_start, imsi_end, iccid_start, iccid_end in data:
        imsies = list(range(imsi_start, imsi_end + 1))
        iccids = list(range(iccid_start, iccid_end + 1))
        yield from zip(imsies, iccids)

# Gerar o arquivo CSV
with open(output_file, mode='w', newline='') as file:
    writer = csv.writer(file)
    writer.writerow(["IMSI", "ICCID"])  # Cabeçalho do arquivo CSV
    for imsi, iccid in generate_ranges(data):
        writer.writerow([imsi, iccid])

print(f"Arquivo '{output_file}' gerado com sucesso!")
