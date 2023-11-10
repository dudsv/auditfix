import os
import pandas as pd

# Função para unir listas e remover vírgulas extras
def unir_e_formatar_lista(lista):
    return ', '.join(filter(None, map(str, lista)))

# Função para varrer as abas e extrair URLs, títulos e descrições sugeridos
def extrair_urls_titulos_descricoes(arquivo_excel):
    urls_info = {}  # Dicionário para armazenar informações agrupadas por Page URL
    
    # Carregando o arquivo Excel
    xls = pd.ExcelFile(arquivo_excel)
    
    # Varrendo todas as abas no arquivo
    for sheet_name in xls.sheet_names:
        # Lendo os dados da aba atual
        df = pd.read_excel(arquivo_excel, sheet_name=sheet_name)
        
        # Verificando se as colunas necessárias estão presentes na aba atual
        if 'Page URL' in df.columns:
            for index, row in df.iterrows():
                page_url = row['Page URL']
                
                # Verificando se 'Page Title Suggested' está presente, caso contrário, deixando vazio
                page_title = row.get('Page Title Suggested', '')  # Convertendo para string
                
                # Verificando se 'Meta Description Suggested' está presente, caso contrário, deixando vazio
                meta_description = row.get('Meta Description Suggested', '')  # Convertendo para string
                
                # Verificando se 'URL sugerida' está presente, caso contrário, deixando vazio
                url_suggested = row.get('URL sugerida', '')  # Convertendo para string
                
                # Verificando se 'Suggested' está presente na aba "Missing or empty H1 tags"
                if 'Missing or empty H1 tags' in sheet_name:
                    h1_suggested = row.get('Suggested', '')  # Convertendo para string
                else:
                    h1_suggested = ''  # Deixando vazio se não estiver na aba correta
                
                # Agrupando informações por Page URL
                if page_url in urls_info:
                    urls_info[page_url]['Page Title Suggested'].append(page_title)
                    urls_info[page_url]['Meta Description Suggested'].append(meta_description)
                    urls_info[page_url]['URL Suggested'].append(url_suggested)
                    urls_info[page_url]['H1 Suggested'].append(h1_suggested)
                else:
                    urls_info[page_url] = {
                        'Page Title Suggested': [page_title],
                        'Meta Description Suggested': [meta_description],
                        'URL Suggested': [url_suggested],
                        'H1 Suggested': [h1_suggested]
                    }
        
        # Adicionando lógica específica para a aba "Meta títulos (>60) y (50<)"
        elif sheet_name == 'Meta títulos (>60) y (50<)':
            for index, row in df.iterrows():
                page_url = row.get('Dirección', '')  # Convertendo para string
                page_title = row.get('Título Propuesto', '')  # Convertendo para string
                
                # Agrupando informações por Page URL
                if page_url in urls_info:
                    urls_info[page_url]['Page Title Suggested'].append(page_title)
                else:
                    urls_info[page_url] = {
                        'Page Title Suggested': [page_title],
                        'Meta Description Suggested': [''],  # Adicionando valores vazios para manter consistência
                        'URL Suggested': [''],
                        'H1 Suggested': ['']
                    }
        
        # Adicionando lógica específica para a aba "H1 (>70) y (20<)"
        elif sheet_name == 'H1 (>70) y (20<)':
            for index, row in df.iterrows():
                page_url = row.get('Dirección', '')  # Convertendo para string
                h1_suggested = row.get('H1 Propuesto', '')  # Convertendo para string
                
                # Agrupando informações por Page URL
                if page_url in urls_info:
                    urls_info[page_url]['H1 Suggested'].append(h1_suggested)
                else:
                    urls_info[page_url] = {
                        'Page Title Suggested': [''],  # Adicionando valores vazios para manter consistência
                        'Meta Description Suggested': [''],
                        'URL Suggested': [''],
                        'H1 Suggested': [h1_suggested]
                    }
        
        else:
            print(f"A coluna 'Page URL' não foi encontrada na aba '{sheet_name}'.")
    
    return urls_info

# Solicitando o nome do arquivo Excel de origem ao usuário
arquivo_excel_origem = input('Digite o caminho do arquivo Excel de origem: ')

# Chamando a função para extrair URLs, títulos e descrições sugeridos
urls_info = extrair_urls_titulos_descricoes(arquivo_excel_origem)

# Se o dicionário de informações estiver vazio, informe ao usuário
if not urls_info:
    print("Nenhum URL com título, descrição ou H1 sugeridos foi encontrado no arquivo fornecido.")
else:
    # Criando uma lista de dicionários para o DataFrame final
    urls_titulos_descricoes = []
    for page_url, info in urls_info.items():
        page_title_combined = unir_e_formatar_lista(info['Page Title Suggested'])
        meta_description_combined = unir_e_formatar_lista(info['Meta Description Suggested'])
        url_suggested_combined = unir_e_formatar_lista(info['URL Suggested'])
        h1_suggested_combined = unir_e_formatar_lista(info['H1 Suggested'])
        urls_titulos_descricoes.append({
            'Page URL': page_url,
            'Page Title Suggested': page_title_combined,
            'Meta Description Suggested': meta_description_combined,
            'URL Suggested': url_suggested_combined,
            'H1 Suggested': h1_suggested_combined
        })

    # Solicitando o nome do arquivo Excel de destino ao usuário
    arquivo_excel_resultado = input('Digite o nome do arquivo Excel de destino para os URLs com títulos, descrições e H1 sugeridos (inclua a extensão .xlsx): ')

    # Adicionando a extensão .xlsx se não estiver presente
    if not arquivo_excel_resultado.endswith('.xlsx'):
        arquivo_excel_resultado += '.xlsx'

    # Criando um novo DataFrame com URLs, títulos, descrições, URLs sugeridos e H1 sugeridos
    urls_df = pd.DataFrame(urls_titulos_descricoes, columns=['Page URL', 'Page Title Suggested', 'Meta Description Suggested', 'URL Suggested', 'H1 Suggested'])

    # Criando um novo arquivo Excel com os URLs, títulos, descrições, URLs sugeridos e H1 sugeridos usando o mecanismo 'openpyxl'
    urls_df.to_excel(arquivo_excel_resultado, index=False, engine='openpyxl')

    print(f'URLs com títulos, descrições, URLs sugeridos e H1 sugeridos foram salvos no arquivo: {arquivo_excel_resultado}')
