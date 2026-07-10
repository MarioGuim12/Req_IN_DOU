import requests
from bs4 import BeautifulSoup
import json
import re
import pandas as pd
import time
from datetime import datetime
hoje = datetime.now()
data_formatada = hoje.strftime("%d-%m-%Y")
print(f"Executando script para a data de hoje: {data_formatada}\n")
url_principal = f"https://www.in.gov.br/leiturajornal?data={data_formatada}&secao=do3"
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36'
}
print("1. Buscando a lista de publicações no DOU...")
response = requests.get(url_principal, headers=headers)
col_titulos = []
col_datas = []
col_processos = []
col_especies = []
col_objetos = []
col_contratados = []
col_links = []
if response.status_code == 200:
    soup = BeautifulSoup(response.text, 'html.parser')
    script_tag = soup.find('script', id='params') or soup.find('script', type='application/json')
    links_artigos = []
    if script_tag:
        try:
            dados_json = json.loads(script_tag.string)
            materias = dados_json.get('jsonArray', []) or dados_json.get('artigos', [])
            for m in materias:
                hierarchy = m.get('hierarchyStr', '')
                if "Advocacia-Geral da União" in hierarchy:
                    url_artigo = m.get('urlTitle', '')
                    if url_artigo:
                        links_artigos.append(f"https://www.in.gov.br/web/dou/-/{url_artigo}")
        except Exception as e:
            print(f"Erro ao ler JSON principal: {e}")
    if not links_artigos:
        for script in soup.find_all('script'):
            if script.string and 'hierarchyStr' in script.string:
                urls_encontradas = re.findall(r'"urlTitle":"([^"]+)"', script.string)
                hierarquias = re.findall(r'"hierarchyStr":"([^"]+)"', script.string)
                for h, u in zip(hierarquias, urls_encontradas):
                    if "Advocacia-Geral da União" in h:
                        links_artigos.append(f"https://www.in.gov.br/web/dou/-/{u}")
    print(f"-> Encontradas {len(links_artigos)} publicações da AGU para o dia {data_formatada}. Iniciando extração interna...\n")
    for i, link in enumerate(links_artigos, 1):
        print(f"[{i}/{len(links_artigos)}] Acessando: {link}")
        time.sleep(1) 
        res_artigo = requests.get(link, headers=headers)
        if res_artigo.status_code == 200:
            soup_artigo = BeautifulSoup(res_artigo.text, 'html.parser')
            script_materia = soup_artigo.find('script', id='params'            
            texto_completo = ""
            titulo_materia = "Sem Título"
            data_publicacao = "Sem Data"
            if script_materia:
                try:
                    json_materia = json.loads(script_materia.string)
                    materia_obj = json_materia.get('materia', json_materia)
                    titulo_materia = materia_obj.get('title', materia_obj.get('name', titulo_materia))
                    data_publicacao = materia_obj.get('pubDate', data_publicacao)
                    texto_completo = materia_obj.get('content', '')
                except:
                    pass
            if not texto_completo:
                titulo_tag = soup_artigo.find('h1', class_='texto-artigo') or soup_artigo.find('p', class_='identifica')
                if titulo_tag: titulo_materia = titulo_tag.get_text(strip=True)
                data_tag = soup_artigo.find('span', class_='publicado-dou-data')
                if data_tag: data_publicacao = data_tag.get_text(strip=True).replace('Publicado em:', '')
                corpo_tag = soup_artigo.find('div', class_='texto-conteudo') or soup_artigo.find('body')
                if corpo_tag: texto_completo = corpo_tag.get_text(separator=" ", strip=True)
            proc_match = re.search(r'(?:Processo|Proc\.?|Nº?\s*Processo)[\sºn°:]*([\d\.\-\/]+)', texto_completo, re.IGNORECASE)
            processo = proc_match.group(1).strip() if proc_match else "Não identificado"
            
            esp_match = re.search(r'(?:Espécie|Especie)[:\s-]([^.]+)', texto_completo, re.IGNORECASE)
            especie = esp_match.group(1).strip() if esp_match else "Não identificado"
            
            obj_match = re.search(r'(?:Objeto)[:\s-]([^.]+)', texto_completo, re.IGNORECASE)
            objeto = obj_match.group(1).strip() if obj_match else "Não identificado"
            
            cont_match = re.search(r'(?:Contratado|Contratada|Beneficiário|Favorecido)[:\s-]([^.]+)', texto_completo, re.IGNORECASE)
            contratado = cont_match.group(1).strip() if cont_match else "Não identificado"
            
            col_titulos.append(titulo_materia)
            col_datas.append(data_publicacao)
            col_processos.append(processo)
            col_especies.append(especie)
            col_objetos.append(objeto)
            col_contratados.append(contratado)
            col_links.append(link)
        else:
            print(f"Erro ao acessar o artigo {link}: Status {res_artigo.status_code}")
    df_detalhado = pd.DataFrame({
        'Data Publicação': col_datas,
        'Título': col_titulos,
        'Nº Processo': col_processos,
        'Espécie': col_especies,
        'Objeto': col_objetos,
        'Contratado': col_contratados,
        'Link': col_links
    })
    df_detalhado.drop_duplicates(inplace=True)
    print("\n--- Processamento Concluído ---")
    if not df_detalhado.empty:
        pd.set_option('display.max_columns', None)
        pd.set_option('display.width', 1200)
        print(df_detalhado.to_string(index=False))
        nome_arquivo = f"publicacoes_agu_{data_formatada}.xlsx"
        df_detalhado.to_excel(nome_arquivo, index=False)
        print(f"\n[Sucesso] Dados do dia salvos em '{nome_arquivo}'!")
    else:
        print(f"Nenhuma publicação da AGU encontrada para o dia {data_formatada}.")
else:
    print(f"Erro na requisição principal: {response.status_code}")
