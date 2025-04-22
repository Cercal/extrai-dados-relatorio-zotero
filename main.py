import csv
from bs4 import BeautifulSoup
import re
from datetime import datetime
from collections import defaultdict

INPUT_HTML = ''
OUTPUT_CSV = ''


def processar_relatorio(html_path):
    with open(html_path, 'r', encoding='utf-8') as f:
        soup = BeautifulSoup(f, 'html.parser')

    dados = {
        'tipos_item': defaultdict(int),
        'autores': defaultdict(int),
        'palavras_resumo': defaultdict(int),
        'datas': defaultdict(int),
        'idiomas': defaultdict(int),
        'periodicos': defaultdict(int),
        'tags': defaultdict(int)
    }

    for item in soup.find_all('li', class_='item'):
        try:
            processar_item(item, dados)
        except Exception as e:
            print(f"Erro ao processar item: {e}")
            continue

    return dados


def processar_item(item, dados):
    # 1. Tipo de item (case-insensitive)
    tipo_th = item.find('th', string=lambda text: text and text.lower().strip() == 'tipo do item')
    if tipo_th:
        tipo = tipo_th.find_next_sibling('td').text.strip().lower()
        dados['tipos_item'][tipo.capitalize()] += 1

    # 2. Autores (normalizado para lowercase)
    for autor_td in item.select('th.author + td'):
        autor = autor_td.text.strip().lower()
        if autor:
            dados['autores'][autor.title()] += 1

    # 5. Palavras em resumos (case-insensitive)
    resumo_th = item.find('th', string='Resumo')
    if resumo_th:
        resumo_td = resumo_th.find_next_sibling('td')
        if resumo_td:
            resumo = resumo_td.text.lower()
            palavras_chave = [
                'patrimônio ambiental', 'environmental heritage', 'património ambiental',
                'sustentabilidade', 'sustainability', 'sostenibilidad',
                'cidade sustentável', 'cidades sustentáveis', 'sustainable city',
                'sustainable cities', 'ciudade sostenible', 'ciudades sostenibles',
                'lixo eletrônico', 'resíduo eletrônico', 'resíduos eletrônicos',
                'e-waste', 'electronic waste', 'basura electrónica',
                'residuo electrónico', 'residuos electrónicos'
            ]

            for palavra in palavras_chave:
                palavra_lower = palavra.lower()
                contagem = len(re.findall(rf'\b{re.escape(palavra_lower)}\b', resumo))
                dados['palavras_resumo'][palavra_lower] += contagem

    # 6. Extração robusta de ano
    data_th = item.find('th', string='Data')
    if data_th:
        data_td = data_th.find_next_sibling('td')
        if data_td:
            data_raw = data_td.text.strip()
            ano = 'Desconhecido'

            # Primeiro busca por qualquer sequência de 4 dígitos
            match = re.search(r'\b\d{4}\b', data_raw)
            if match:
                ano = int(match.group())
            else:
                # Tenta formatos de data conhecidos
                formatos = [
                    '%Y-%m-%d', '%d-%m-%Y', '%Y', '%m/%Y', '%b %Y', '%B %Y'
                ]
                for fmt in formatos:
                    try:
                        dt = datetime.strptime(data_raw, fmt)
                        ano = dt.year
                        break
                    except:
                        continue

                # Último recurso: verifica se é um número válido
                if ano == 'Desconhecido' and data_raw.isdigit():
                    try:
                        ano = int(data_raw)
                    except:
                        pass

            dados['datas'][ano] += 1

    # 7. Idioma (normalizado)
    idioma_th = item.find('th', string='Idioma')
    if idioma_th:
        idioma_td = idioma_th.find_next_sibling('td')
        if idioma_td:
            idioma = idioma_td.text.strip().lower()
            mapeamento = {
                'por': 'Português',
                'en': 'Inglês',
                'eng': 'Inglês',
                'esp': 'Espanhol',
                'spa': 'Espanhol'
            }
            dados['idiomas'][mapeamento.get(idioma, 'Não Identificado')] += 1

    # 8. Periódicos (mantém case original)
    if tipo_th and 'periódico' in tipo.lower():
        periodico_th = item.find('th', string='Título da publicação')
        if periodico_th:
            periodico_td = periodico_th.find_next_sibling('td')
            if periodico_td:
                periodico = periodico_td.text.strip()
                if periodico:
                    dados['periodicos'][periodico] += 1

    # 10. Tags (case-insensitive)
    for tag in item.select('ul.tags li'):
        tag_text = tag.text.strip().lower()
        if tag_text:
            dados['tags'][tag_text] += 1


def gerar_csv(dados, output_path):
    with open(output_path, 'w', newline='', encoding='utf-8') as f:
        writer = csv.writer(f)

        # 1. Tipos de item
        writer.writerow(['Por tipo de item'])
        writer.writerow(['Tipo', 'Quantidade'])
        for tipo, qtd in sorted(dados['tipos_item'].items(), key=lambda x: x[0].lower()):
            writer.writerow([tipo, qtd])
        writer.writerow([])

        # 2. Autores
        writer.writerow(['Por autor'])
        writer.writerow(['Autor', 'Publicações'])
        for autor, qtd in sorted(dados['autores'].items(), key=lambda x: x[0].lower()):
            writer.writerow([autor, qtd])
        writer.writerow([])

        # 5. Palavras em resumos
        writer.writerow(['Incidência de palavras em resumos'])
        writer.writerow(['Termo', 'Ocorrências'])
        ordem_palavras = [
            'patrimônio ambiental', 'environmental heritage', 'património ambiental',
            'sustentabilidade', 'sustainability', 'sostenibilidad',
            'cidade sustentável', 'cidades sustentáveis', 'sustainable city',
            'sustainable cities', 'ciudade sostenible', 'ciudades sostenibles',
            'lixo eletrônico', 'resíduo eletrônico', 'resíduos eletrônicos',
            'e-waste', 'electronic waste', 'basura electrónica',
            'residuo electrónico', 'residuos electrónicos'
        ]
        for palavra in ordem_palavras:
            palavra_lower = palavra.lower()
            writer.writerow([palavra, dados['palavras_resumo'].get(palavra_lower, 0)])
        writer.writerow([])

        # 6. Datas
        writer.writerow(['Por data'])
        writer.writerow(['Ano', 'Publicações'])
        anos_validos = [d for d in dados['datas'].items() if isinstance(d[0], int)]
        for ano, qtd in sorted(anos_validos, reverse=True):
            writer.writerow([ano, qtd])
        if 'Desconhecido' in dados['datas']:
            writer.writerow(['Desconhecido', dados['datas']['Desconhecido']])
        writer.writerow([])

        # 7. Idiomas
        writer.writerow(['Por idioma'])
        writer.writerow(['Idioma', 'Publicações'])
        for idioma in ['Português', 'Inglês', 'Espanhol', 'Não Identificado']:
            writer.writerow([idioma, dados['idiomas'].get(idioma, 0)])
        writer.writerow([])

        # 8. Periódicos
        writer.writerow(['Por periódico'])
        writer.writerow(['Periódico', 'Publicações'])
        for periodo, qtd in sorted(dados['periodicos'].items(), key=lambda x: x[0].lower()):
            writer.writerow([periodo, qtd])
        writer.writerow([])

        # 10. Palavras-chave
        writer.writerow(['Palavras-chave'])
        writer.writerow(['Tag', 'Frequência'])
        tags_ordenadas = sorted(dados['tags'].items(), key=lambda x: (-x[1], x[0]))
        for tag, qtd in tags_ordenadas:
            writer.writerow([tag.title(), qtd])


try:
    dados = processar_relatorio(INPUT_HTML)
    gerar_csv(dados, OUTPUT_CSV)
    print("Processamento concluído com sucesso!")
except Exception as e:
    print(f"Erro fatal: {str(e)}")