from docling.document_converter import DocumentConverter
import pandas as pd
import locale
import re

# Configurar locale para português brasileiro
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')


def formatar_numero(texto):
    if not texto or texto == 'None':
        return texto

    # Tenta converter o texto para número
    try:
        # Remove caracteres não numéricos, exceto ponto, hífen e vírgula
        numero_texto = re.sub(r'[^\d.-]', '', texto.replace(',', '.'))
        numero = float(numero_texto)

        # Se for porcentagem
        if '%' in texto or abs(numero) <= 1:
            return f"{numero:.2%}".replace('.', ',')

        # Se for valor monetário (maior que 100)
        if abs(numero) >= 100:
            return f"{numero:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')

        # Se for valor decimal
        if isinstance(numero, float):
            return f"{numero:.2f}".replace('.', ',')

        # Se for número inteiro
        if numero.is_integer():
            return f"{int(numero):,}".replace(',', '.')

        return texto
    except:
        return texto


source = input("Digite o nome do arquivo XLSX: ")
if not source.endswith('.xlsx'):
    raise ValueError("O arquivo deve ser um arquivo XLSX.")
converter = DocumentConverter()
result = converter.convert(source)

# Obtém o conteúdo markdown atual
markdown_content = result.document.export_to_markdown()

# Processa linha por linha
linhas_processadas = []
primeira_linha_tabela = True

for linha in markdown_content.split('\n'):
    # Se for uma linha da tabela (contém |)
    if '|' in linha:
        if primeira_linha_tabela:
            primeira_linha_tabela = False
            linhas_processadas.append(linha)
            continue

        # Divide a linha em colunas
        colunas = [col.strip() for col in linha.split('|')]
        # Formata cada coluna
        colunas_formatadas = [formatar_numero(col) for col in colunas]
        # Reconstrói a linha
        linha = '|'.join(colunas_formatadas)
    linhas_processadas.append(linha)

# Reconstrói o conteúdo markdown
novo_markdown = '\n'.join(linhas_processadas)

# nomeia o arquivo de saída
arquivo_saida = source.replace('.xlsx', '.md')

# Salva o arquivo formatado
with open(arquivo_saida, "w", encoding='utf-8') as f:
    f.write(novo_markdown)
