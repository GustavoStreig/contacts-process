import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import re
import os

# Funções de seleção de arquivos e diretórios
def select_excel_file():
    file_path = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xls;*.xlsx')])
    excel_file_var.set(file_path)

def select_csv_file():
    file_path = filedialog.askopenfilename(filetypes=[('CSV Files', '*.csv')])
    csv_file_var.set(file_path)

def select_output_directory():
    directory = filedialog.askdirectory()
    output_dir_var.set(directory)

# Função para processar os dados do Excel e opcionalmente do CSV
def process_file():
    excel_file_path = excel_file_var.get()
    csv_file_path = csv_file_var.get()
    output_folder = output_dir_var.get()
    
    if not excel_file_path or not output_folder:
        messagebox.showwarning("Warning", "Please select both the Excel file and the output folder.")
        return
    
    # Carregar o arquivo Excel
    df_excel = pd.read_excel(excel_file_path)
    
    # Verificar se as colunas de nome existem no DataFrame, caso contrário, cria colunas vazias
    name_columns = ['Nome', 'Segundo nome', 'Sobrenome']
    for col in name_columns:
        if col not in df_excel.columns:
            df_excel[col] = ''
    
    # Substituir valores nulos por strings vazias e concatenar as colunas de nome para formar 'Nome Completo'
    df_excel[name_columns] = df_excel[name_columns].fillna('')
    df_excel[name_columns] = df_excel[name_columns].astype(str)  # Garantir que todos sejam strings
    df_excel['Nome Completo'] = df_excel[name_columns].agg(' '.join, axis=1)
    
    # Remover caracteres que não sejam letras ou números do campo 'Nome Completo'
    df_excel['Nome Completo'] = df_excel['Nome Completo'].apply(lambda x: re.sub(r'[^a-zA-Z0-9 ]', '', x))
    
    # Função para reposicionar números no 'Nome Completo'
    def reposicionar_numero(nome):
        match = re.search(r'\d+', nome)
        if match:
            numero = match.group(0)
            letras = re.sub(r'\d+', '', nome).strip()
            return f"{numero} - {letras}"
        return nome
    
    # Aplicar a função de reposicionar números no 'Nome Completo'
    df_excel['Nome Completo'] = df_excel['Nome Completo'].apply(reposicionar_numero)
    
    # Lista de colunas de telefone
    phone_columns = [
        'Celular', 'Celular2', 'Celular3', 'Celular4',
        'Telefone trabalhar', 'Telefone trabalhar2', 'Telefone trabalhar3',
        'Telefone residencial', 'Outro telefone', 'CelularTelefone',
        'CelularTelefone2', 'AntigoTelefone', 'AntigoTelefone2',
        'CasaTelefone', 'CasaTelefone2', 'ResidencialTelefone',
        'ResidencialTelefone2', 'ComercialTelefone', 'ComercialTelefone2',
        'OutrosTelefone', 'OutrosTelefone2', 'Cel.Telefone', 'Cel.Telefone2',
        'WhatsAppTelefone', 'OutroTelefone'
    ]
    
    # Filtrar apenas as colunas de telefone que existem no DataFrame
    existing_phone_columns = [col for col in phone_columns if col in df_excel.columns]
    
    # Transformar todos os campos de telefone em apenas um e duplicar linhas conforme necessário
    rows = []
    for index, row in df_excel.iterrows():
        nome_completo = row['Nome Completo']
        for col in existing_phone_columns:
            telefone = row[col]
            if pd.notna(telefone):
                telefone_limpo = re.sub(r'\D', '', str(telefone))
                rows.append({'Nome Completo': nome_completo, 'Telefone': telefone_limpo})
    
    # Criar um novo DataFrame com as linhas duplicadas
    df_final = pd.DataFrame(rows)
    
    # Remover duplicatas com base no número de telefone
    df_final = df_final.drop_duplicates(subset=['Telefone'])
    
    # Remover o primeiro dígito do número se ele for igual a zero
    df_final['Telefone'] = df_final['Telefone'].apply(lambda x: x[1:] if x.startswith('0') else x)
    
    # Remover os dois primeiros dígitos do número se ele começar com 41
    df_final['Telefone'] = df_final['Telefone'].apply(lambda x: x[2:] if x.startswith('41') else x)
    
    # Adicionar 55 na frente de todos os números que não começam com 55
    df_final['Telefone'] = df_final['Telefone'].apply(lambda x: f"55{x}" if not x.startswith('55') else x)
    
    # Filtrar os nomes que não possuem código de cliente no 'Nome Completo'
    df_sem_codigo = df_final[~df_final['Nome Completo'].str.contains(r'^\d+ - ')]
    
    # Remover esses nomes do DataFrame original
    df_final_com_codigo = df_final[df_final['Nome Completo'].str.contains(r'^\d+ - ')]
    
    # Filtrar os números de telefone que estão fora do padrão 55DDXXXXXXXXX ou 55DDXXXXXXXX
    padrao_telefone = r'^55\d{10,11}$'
    df_telefone_valido = df_final_com_codigo[df_final_com_codigo['Telefone'].str.match(padrao_telefone)]
    df_telefone_invalido = df_final_com_codigo[~df_final_com_codigo['Telefone'].str.match(padrao_telefone)]
    
    # Remover duplicatas com base no número de telefone
    df_telefone_valido = df_telefone_valido.drop_duplicates(subset=['Telefone'])
    
    # Adicionar uma coluna de comprimento do número de telefone
    df_telefone_valido['Num_Length'] = df_telefone_valido['Telefone'].apply(len)
    
    # Função para remover os primeiros 4 dígitos
    def remove_first_4_digits(num):
        return num[4:]
    
    # Função para remover os primeiros 5 dígitos
    def remove_first_5_digits(num):
        return num[5:]
    
    # Agrupar por nome
    grouped = df_telefone_valido.groupby('Nome Completo')
    
    # Lista para armazenar índices a serem removidos
    indices_to_remove = []
    
    # Iterar sobre cada grupo
    for name, group in grouped:
        # Obter números com menor e maior comprimento
        for i, row in group.iterrows():
            for j, compare_row in group.iterrows():
                if row['Num_Length'] < compare_row['Num_Length']:
                    short_num = remove_first_4_digits(row['Telefone'])
                    long_num = remove_first_5_digits(compare_row['Telefone'])
                    if short_num == long_num:
                        indices_to_remove.append(i)
    
    # Remover os índices identificados
    df_telefone_valido = df_telefone_valido.drop(indices_to_remove)
    
    # Remover a coluna auxiliar Num_Length
    df_telefone_valido = df_telefone_valido.drop(columns=['Num_Length'])
    
    # Carregar o arquivo CSV apenas se existir
    if csv_file_path:
        try:
            df_csv = pd.read_csv(csv_file_path)
        except FileNotFoundError:
            df_csv = pd.DataFrame(columns=['number'])  # DataFrame vazio se o arquivo CSV não for encontrado
        
        # Verificar se a coluna 'number' existe no DataFrame CSV
        if 'number' not in df_csv.columns:
            df_csv['number'] = ''
        
        # Converter números de telefone para string em ambos os DataFrames
        df_telefone_valido['Telefone'] = df_telefone_valido['Telefone'].astype(str)
        df_csv['number'] = df_csv['number'].astype(str)
        
        # Encontrar números duplicados
        duplicated_numbers = df_telefone_valido[df_telefone_valido['Telefone'].isin(df_csv['number'])]
        
        # Criar o caminho para o relatório dos números duplicados
        duplicated_report_path = os.path.join(output_folder, 'Duplicated_Numbers_Report.xlsx')
        duplicated_numbers.to_excel(duplicated_report_path, index=False)
    
    # Criar nomes de arquivos para os resultados
    base_name = os.path.splitext(os.path.basename(excel_file_path))[0]
    output_path_final = os.path.join(output_folder, f'{base_name}_Contatos_Processados_Validos_Atualizado.xlsx')
    output_path_invalid = os.path.join(output_folder, f'{base_name}_Contatos_Telefones_Invalidos.xlsx')
    output_path_sem_codigo = os.path.join(output_folder, f'{base_name}_Contatos_Sem_Codigo.xlsx')
    
    # Salvar os DataFrames em arquivos Excel
    df_telefone_valido.to_excel(output_path_final, index=False)
    df_telefone_invalido.to_excel(output_path_invalid, index=False)
    df_sem_codigo.to_excel(output_path_sem_codigo, index=False)
    
    messagebox.showinfo("Info", f"Files processed and saved to:\n{output_path_final}\n{output_path_invalid}\n{output_path_sem_codigo}\nDuplicated numbers report saved to:\n{duplicated_report_path if csv_file_path else 'N/A'}")

# Criação da interface gráfica
root = tk.Tk()
root.title("Contacts Processor")

frame = tk.Frame(root)
frame.pack(padx=10, pady=10)

# Variáveis de texto para os campos
excel_file_var = tk.StringVar()
csv_file_var = tk.StringVar()
output_dir_var = tk.StringVar()

# Título centralizado
tk.Label(frame, text="Contact Process Kentro", font=("Helvetica", 10)).grid(row=0, column=0, columnspan=3, pady=10)

# Campos e botões para selecionar o arquivo Excel
tk.Label(frame, text="Select Excel File:").grid(row=1, column=0, sticky="e")
tk.Entry(frame, textvariable=excel_file_var, width=50).grid(row=1, column=1, padx=5, pady=5)
tk.Button(frame, text="Browse", command=select_excel_file).grid(row=1, column=2, padx=5, pady=5)

# Campos e botões para selecionar o arquivo CSV
tk.Label(frame, text="Select CSV File (optional):").grid(row=2, column=0, sticky="e")
tk.Entry(frame, textvariable=csv_file_var, width=50).grid(row=2, column=1, padx=5, pady=5)
tk.Button(frame, text="Browse", command=select_csv_file).grid(row=2, column=2, padx=5, pady=5)

# Campos e botões para selecionar o diretório de saída
tk.Label(frame, text="Select Output Directory:").grid(row=3, column=0, sticky="e")
tk.Entry(frame, textvariable=output_dir_var, width=50).grid(row=3, column=1, padx=5, pady=5)
tk.Button(frame, text="Browse", command=select_output_directory).grid(row=3, column=2, padx=5, pady=5)

# Botão para processar o arquivo
process_button = tk.Button(frame, text="Process", command=process_file)
process_button.grid(row=4, columnspan=3, pady=10)


root.mainloop()
