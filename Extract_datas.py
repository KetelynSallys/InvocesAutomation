import PyPDF2
import pandas as pd
import os
import re


def extract_text_from_pdf(pdf_path):
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        text = ''
        for page in reader.pages:
            text += page.extract_text()
    return text

def find_group(pattern, lines, grupo=1):
    for line in lines:
        match = pattern.search(line)
        if match:
            return match.group(grupo)
    return 'Not found' 


def extract_data(text):
    data = {}
    lines = text.split('\n')

    source_pattern = re.compile(r'Source(\d+)')
    issue_data_pattern = re.compile(r'(\d{2}/\d{2}/\d{4})')
    value_pattern = re.compile(r'(\d{1,3}(?:\.\d{3})*,\d{2})')
    service_code_pattern = re.compile(r'(02\d{3})')
    other_corporate_cpf_cnpj_pattern = re.compile(r'(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})')
    cpf_cnpj_pattern = re.compile(r'(11\.111\.111/\d{4}-\d{2})|(11\.222\.222/\d{4}-\d{2})')
    corporate_reason_pattern = re.compile(r"(corporate reason 1)|(corporate reason 2)")
    
    data['Title 1'] = next((other_corporate_cpf_cnpj_pattern.search(line).group(1) for line in lines if other_corporate_cpf_cnpj_pattern.search(line)), 'Not found')
    data['Title 2'] = next((corporate_reason_pattern.search(line).group(0) for line in lines if corporate_reason_pattern.search(line)), 'Not found')
    data['Title 3'] = next((cpf_cnpj_pattern.search(line).group(1) for line in lines if cpf_cnpj_pattern.search(line)), 'Not found')
    data['Title 4'] = next((source_pattern.search(line).group(1) for line in lines if source_pattern.search(line)), 'Não encontrado')
    data['Title 5'] = next((issue_data_pattern .search(line).group(1) for line in lines if issue_data_pattern .search(line)), 'Not found')
    data['Title 6'] = next((service_code_pattern.search(line).group(1) for line in lines if service_code_pattern.search(line)), 'Not found')
    data['Title 7'] = next(( value_pattern.search(line).group(1) for line in lines if  value_pattern.search(line)), 'Not found')
    
    return data


folder_path = r'C:\Users\insert\your\path\for\pdf'


all_data = []

for filename in os.listdir(folder_path):
    if filename.endswith('.pdf'):
        pdf_path = os.path.join(folder_path, filename)
        text = extract_text_from_pdf(pdf_path)
        #print(f"extracted text of {filename}:\n{text}\n")
        data = extract_data(text)
        all_data.append(data)


df = pd.DataFrame(all_data)
df.to_excel('name_of_your_spreadsheet.xlsx', index=False)


print("Extracted and saved data")

