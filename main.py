import jpype
import jpype.imports
from tabula.io import read_pdf
import pandas as pd

# Inicializar o JVM manualmente
jpype.startJVM("C:/Program Files/Java/jdk-21/bin/server/jvm.dll")

# Função para identificar se a linha é uma seção
def is_section(row_text):
    return isinstance(row_text, str) and row_text.startswith(("1.", "2.", "3.", "4.")) and "RPM" in row_text

# Função principal
def process_pdf_to_rpm_tables(pdf_path, output_excel_path, output_csv_path):
    # Extrai todas as tabelas do PDF
    tables = read_pdf(pdf_path, pages="all", multiple_tables=True, pandas_options={"header": None})

    # Inicializa DataFrame final
    columns = ['ORD', 'INSCRIÇÃO', 'NOME', 'TOTAL', 'RPM']
    final_df = pd.DataFrame(columns=columns)

    current_rpm = None  # Para armazenar a RPM atual

    # Processa as tabelas
    for table in tables:
        # Verifica se há uma linha indicando a seção/RPM
        if isinstance(table.iloc[0, 0], str) and is_section(table.iloc[0, 0]):
            current_rpm = table.iloc[0, 0]  # Define a RPM atual
            table = table.iloc[1:]  # Remove o título da tabela

        # Processa apenas tabelas válidas
        if current_rpm:
            # Renomeia colunas e organiza
            try:
                table.columns = ['ORD', 'INSCRIÇÃO', 'NOME', 'TOTAL']
                table = table[['ORD', 'INSCRIÇÃO', 'NOME', 'TOTAL']]  # Seleciona apenas colunas relevantes
                table['RPM'] = current_rpm  # Adiciona a coluna RPM
                final_df = pd.concat([final_df, table], ignore_index=True)
            except ValueError:
                print("Erro ao processar tabela. Verifique o formato das tabelas no PDF.")

    # Salva os resultados em Excel e CSV
    final_df.to_excel(output_excel_path, index=False)
    final_df.to_csv(output_csv_path, index=False)
    print(f"Arquivo processado salvo como:\nExcel: {output_excel_path}\nCSV: {output_csv_path}")

# Caminhos de entrada e saída
pdf_path = r"C:\Users\valfr\Downloads\221120241528123230.pdf"
output_excel_path = r"C:\Users\valfr\Downloads\processed_rpm_tables.xlsx"
output_csv_path = r"C:\Users\valfr\Downloads\processed_rpm_tables.csv"

# Processa o PDF
process_pdf_to_rpm_tables(pdf_path, output_excel_path, output_csv_path)
