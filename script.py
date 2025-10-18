from docling.document_converter import DocumentConverter
import pandas as pd
import locale
import re
import tempfile
import os
import glob

# Configurar locale para português brasileiro
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')


def formatar_numero(texto):
    if not texto or texto == 'None':
        return texto

    # Tenta converter o texto para número
    try:
        # Remove espaços em branco extras
        texto = str(texto).strip()
        
        # Detecta e converte notação científica (ex: 2.76508e+06)
        if 'e+' in texto.lower() or 'e-' in texto.lower():
            numero = float(texto)
        else:
            # Remove caracteres não numéricos, exceto ponto, hífen e vírgula
            numero_texto = re.sub(r'[^\d.-]', '', texto.replace(',', '.'))
            if not numero_texto or numero_texto == '-':
                return texto
            numero = float(numero_texto)

        # Se for porcentagem (valores muito pequenos ou texto contém %)
        if '%' in texto and abs(numero) <= 1:
            return f"{numero:.2%}".replace('.', ',')

        # Se for valor monetário (maior que 100)
        if abs(numero) >= 100:
            return f"{numero:,.2f}".replace(',', '_').replace('.', ',').replace('_', '.')

        # Se for valor decimal pequeno
        if isinstance(numero, float):
            return f"{numero:.2f}".replace('.', ',')

        # Se for número inteiro
        if numero.is_integer():
            return f"{int(numero):,}".replace(',', '.')

        return texto
    except:
        return texto


# Cria as pastas se não existirem
os.makedirs('xlsx', exist_ok=True)
os.makedirs('md', exist_ok=True)

# Lista todos os arquivos XLSX na pasta xlsx
arquivos_xlsx = glob.glob('xlsx/*.xlsx')

if not arquivos_xlsx:
    print("Nenhum arquivo XLSX encontrado na pasta 'xlsx'.")
    print("Por favor, coloque seus arquivos XLSX na pasta 'xlsx' e execute o script novamente.")
    exit(1)
else:
    print("Arquivos XLSX disponíveis na pasta 'xlsx':")
    for idx, arquivo in enumerate(arquivos_xlsx, 1):
        # Mostra apenas o nome do arquivo sem o caminho
        nome_arquivo = os.path.basename(arquivo)
        print(f"{idx}. {nome_arquivo}")
    
    escolha_arquivo = input("\nDigite o número do arquivo que deseja processar: ").strip()
    try:
        idx_arquivo = int(escolha_arquivo) - 1
        if 0 <= idx_arquivo < len(arquivos_xlsx):
            source = arquivos_xlsx[idx_arquivo]
        else:
            raise ValueError("Número inválido.")
    except ValueError:
        raise ValueError("Entrada inválida. Digite um número válido.")

# Lista todas as abas disponíveis
excel_file = pd.ExcelFile(source)
abas = excel_file.sheet_names

print(f"\nAbas disponíveis no arquivo:")
for idx, aba in enumerate(abas, 1):
    print(f"{idx}. {aba}")

# Pergunta qual aba o usuário quer processar
escolha = input("\nDigite o número da aba que deseja processar (ou 'todas' para processar todas): ").strip()

if escolha.lower() == 'todas':
    abas_processar = abas
else:
    try:
        idx_escolhido = int(escolha) - 1
        if 0 <= idx_escolhido < len(abas):
            abas_processar = [abas[idx_escolhido]]
        else:
            raise ValueError("Número inválido.")
    except ValueError:
        raise ValueError("Entrada inválida. Digite um número válido ou 'todas'.")

# Processa cada aba selecionada
for aba_nome in abas_processar:
    print(f"\nProcessando aba: {aba_nome}")
    
    # Lê o DataFrame original para ter os valores reais
    df_original = pd.read_excel(source, sheet_name=aba_nome)
    
    # Cria um arquivo temporário com apenas a aba selecionada
    with tempfile.NamedTemporaryFile(suffix='.xlsx', delete=False) as temp_file:
        temp_path = temp_file.name
        df_original.to_excel(temp_path, index=False)
    
    # Converte o arquivo temporário
    converter = DocumentConverter()
    result = converter.convert(temp_path)
    
    # Remove o arquivo temporário
    os.unlink(temp_path)
    
    # Obtém o conteúdo markdown atual
    markdown_content = result.document.export_to_markdown()

    # Processa linha por linha
    linhas_processadas = []
    linha_dados_idx = 0  # Índice para as linhas de dados do DataFrame
    primeira_linha_tabela = True

    for linha in markdown_content.split('\n'):
        # Se for uma linha da tabela (contém |)
        if '|' in linha:
            if primeira_linha_tabela:
                primeira_linha_tabela = False
                linhas_processadas.append(linha)
                continue
            
            # Se for a linha separadora (só contém - e |)
            if re.match(r'^[\s\-|]+$', linha):
                linhas_processadas.append(linha)
                continue

            # Divide a linha em colunas
            colunas = [col.strip() for col in linha.split('|')]
            
            # Se temos dados do DataFrame, usa os valores reais
            if linha_dados_idx < len(df_original):
                colunas_formatadas = []
                for idx, col in enumerate(colunas):
                    # Índice da coluna no DataFrame (ajusta por causa do | inicial)
                    df_col_idx = idx - 1
                    
                    # Se existe essa coluna no DataFrame, usa o valor real
                    if 0 <= df_col_idx < len(df_original.columns):
                        valor_original = df_original.iloc[linha_dados_idx, df_col_idx]
                        # Se o valor é numérico, formata usando o valor real
                        if pd.notna(valor_original) and isinstance(valor_original, (int, float)):
                            colunas_formatadas.append(formatar_numero(str(valor_original)))
                        else:
                            # Para valores não numéricos, usa o valor original do DataFrame
                            colunas_formatadas.append(str(valor_original) if pd.notna(valor_original) else '')
                    else:
                        colunas_formatadas.append(col)
                
                linha_dados_idx += 1
            else:
                # Se não há mais dados no DataFrame, formata normalmente
                colunas_formatadas = [formatar_numero(col) for col in colunas]
            
            # Reconstrói a linha
            linha = '|'.join(colunas_formatadas)
        linhas_processadas.append(linha)

    # Reconstrói o conteúdo markdown
    novo_markdown = '\n'.join(linhas_processadas)

    # nomeia o arquivo de saída
    nome_base = os.path.basename(source).replace('.xlsx', '')
    if len(abas_processar) == 1:
        arquivo_saida = os.path.join('md', f'{nome_base}.md')
    else:
        # Sanitiza o nome da aba para usar como nome de arquivo
        nome_aba_limpo = re.sub(r'[<>:"/\\|?*]', '_', aba_nome)
        arquivo_saida = os.path.join('md', f'{nome_base}_{nome_aba_limpo}.md')

    # Salva o arquivo formatado
    with open(arquivo_saida, "w", encoding='utf-8') as f:
        f.write(novo_markdown)
    
    print(f"Arquivo salvo: {arquivo_saida}")
