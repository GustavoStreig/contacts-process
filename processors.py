import pandas as pd
import re
import os
from tkinter import messagebox

class DataProcessor:
    def __init__(self, excel_file_path, csv_file_path, output_folder):
        self.excel_file_path = excel_file_path
        self.csv_file_path = csv_file_path
        self.output_folder = output_folder

    def process_file(self):
        if not self.excel_file_path or not self.output_folder:
            messagebox.showwarning("Warning", "Please select both the Excel file and the output folder.")
            return

        try:
            # Carregar o arquivo Excel
            df_excel = pd.read_excel(self.excel_file_path)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to read Excel file: {e}")
            return

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

        # Remover os primeiros 4 dígitos se coincidirem
        df_telefone_valido['Num_Length'] = df_telefone_valido['Telefone'].apply(len)

        def remove_first_4_digits(num):
            return num[4:]

        def remove_first_5_digits(num):
            return num[5:]

        grouped = df_telefone_valido.groupby('Nome Completo')
        indices_to_remove = []
        for name, group in grouped:
            for i, row in group.iterrows():
                for j, compare_row in group.iterrows():
                    if row['Num_Length'] < compare_row['Num_Length']:
                        short_num = remove_first_4_digits(row['Telefone'])
                        long_num = remove_first_5_digits(compare_row['Telefone'])
                        if short_num == long_num:
                            indices_to_remove.append(i)

        df_telefone_valido = df_telefone_valido.drop(indices_to_remove)
        df_telefone_valido = df_telefone_valido.drop(columns=['Num_Length'])

        duplicated_report_path = ''
        if self.csv_file_path:
            try:
                df_csv = pd.read_csv(self.csv_file_path, delimiter=';')
            except FileNotFoundError:
                df_csv = pd.DataFrame(columns=['number'])

            if 'number' not in df_csv.columns:
                df_csv['number'] = ''

            def normalize_phone_number(phone):
                if pd.isna(phone):
                    return ""
                phone = ''.join(filter(str.isdigit, str(phone)))
                return phone

            df_telefone_valido['Telefone'] = df_telefone_valido['Telefone'].apply(normalize_phone_number)
            df_csv['number'] = df_csv['number'].apply(normalize_phone_number)
            df_telefone_valido['Telefone'] = df_telefone_valido['Telefone'].astype(str)
            df_csv['number'] = df_csv['number'].astype(str)
            duplicated_numbers = df_telefone_valido[df_telefone_valido['Telefone'].isin(df_csv['number'])]
            duplicated_report_path = os.path.join(self.output_folder, 'Duplicated_Numbers_Report.xlsx')
            duplicated_numbers.to_excel(duplicated_report_path, index=False)
            df_telefone_valido = df_telefone_valido[~df_telefone_valido['Telefone'].isin(df_csv['number'])]

        base_name = os.path.splitext(os.path.basename(self.excel_file_path))[0]
        output_path_final = os.path.join(self.output_folder, f'{base_name}_Contatos_Processados_Validos_Atualizado.xlsx')
        output_path_invalid = os.path.join(self.output_folder, f'{base_name}_Contatos_Telefones_Invalidos.xlsx')
        output_path_nocodes = os.path.join(self.output_folder, f'{base_name}_Nomes_Sem_Codigo_Cliente.xlsx')

        df_telefone_valido.to_excel(output_path_final, index=False)
        df_telefone_invalido.to_excel(output_path_invalid, index=False)
        df_sem_codigo.to_excel(output_path_nocodes, index=False)

        messagebox.showinfo("Success", f"Processing completed.\n\nValid Contacts: {output_path_final}\nInvalid Contacts: {output_path_invalid}\nNo Client Code: {output_path_nocodes}\nDuplicated Numbers Report: {duplicated_report_path}")
