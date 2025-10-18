# Financial Data Processor

Processador de dados financeiros que converte arquivos Excel (XLSX) em Markdown formatado, com suporte para múltiplas abas e formatação automática de números brasileiros.

## 📋 Funcionalidades

-   ✅ Listagem automática de arquivos XLSX disponíveis
-   ✅ Seleção interativa de arquivos e abas
-   ✅ Processamento de uma aba específica ou todas as abas
-   ✅ Formatação automática de números no padrão brasileiro
-   ✅ Conversão de valores monetários
-   ✅ Formatação de porcentagens
-   ✅ Organização automática em pastas (xlsx/md)

## 🚀 Como Usar

### 1. Requisitos

Certifique-se de ter Python instalado e as dependências necessárias:

```bash
pip install docling pandas openpyxl
```

### 2. Estrutura de Pastas

O script organiza automaticamente os arquivos em pastas:

```
financial-data-processor/
├── xlsx/          # Coloque seus arquivos Excel aqui
├── md/            # Arquivos Markdown gerados aqui
├── script.py      # Script principal
└── README.md      # Este arquivo
```

### 3. Executar o Script

```bash
python script.py
```

### 4. Processo Interativo

O script guiará você através de um processo interativo:

**Passo 1:** Escolha o arquivo XLSX

```
Arquivos XLSX disponíveis na pasta 'xlsx':
1. DRE CONDO.xlsx
2. Relatorio_2024.xlsx

Digite o número do arquivo que deseja processar: 1
```

**Passo 2:** Escolha a aba para processar

```
Abas disponíveis no arquivo:
1. Janeiro
2. Fevereiro
3. Março

Digite o número da aba que deseja processar (ou 'todas' para processar todas): 2
```

**Passo 3:** Aguarde o processamento

```
Processando aba: Fevereiro
Arquivo salvo: md\DRE CONDO.md
```

## 📊 Formatação de Números

O script formata automaticamente os números seguindo o padrão brasileiro:

| Tipo            | Entrada         | Saída      |
| --------------- | --------------- | ---------- |
| Porcentagem     | `0.15` ou `15%` | `15,00%`   |
| Valor Monetário | `1234.56`       | `1.234,56` |
| Decimal         | `3.14`          | `3,14`     |
| Inteiro         | `1000`          | `1.000`    |

## 📁 Arquivos de Saída

-   **Uma aba processada:** `md/NomeDoArquivo.md`
-   **Múltiplas abas processadas:** `md/NomeDoArquivo_NomeDaAba.md`

## 🔧 Configurações

O script usa a localização brasileira (pt_BR.UTF-8) para formatação de números. Se necessário, ajuste a linha:

```python
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')
```

## 💡 Exemplos de Uso

### Processar uma aba específica

```bash
$ python script.py
Arquivos XLSX disponíveis na pasta 'xlsx':
1. DRE CONDO.xlsx

Digite o número do arquivo que deseja processar: 1

Abas disponíveis no arquivo:
1. Resumo
2. Detalhado

Digite o número da aba que deseja processar (ou 'todas' para processar todas): 1

Processando aba: Resumo
Arquivo salvo: md\DRE CONDO.md
```

### Processar todas as abas

```bash
$ python script.py
Arquivos XLSX disponíveis na pasta 'xlsx':
1. DRE CONDO.xlsx

Digite o número do arquivo que deseja processar: 1

Abas disponíveis no arquivo:
1. Janeiro
2. Fevereiro
3. Março

Digite o número da aba que deseja processar (ou 'todas' para processar todas): todas

Processando aba: Janeiro
Arquivo salvo: md\DRE CONDO_Janeiro.md

Processando aba: Fevereiro
Arquivo salvo: md\DRE CONDO_Fevereiro.md

Processando aba: Março
Arquivo salvo: md\DRE CONDO_Março.md
```

## ⚠️ Observações

-   Os arquivos Excel devem estar na pasta `xlsx/`
-   Os arquivos Markdown serão salvos na pasta `md/`
-   As pastas são criadas automaticamente se não existirem
-   O script preserva a formatação de tabelas do Excel
-   Nomes de abas com caracteres especiais são sanitizados

## 🛠️ Tecnologias

-   **Python 3.x**
-   **docling** - Conversão de documentos
-   **pandas** - Manipulação de dados
-   **openpyxl** - Leitura de arquivos Excel

## 📝 Licença

Este projeto é de código aberto e está disponível para uso livre.

## 🤝 Contribuições

Contribuições são bem-vindas! Sinta-se à vontade para abrir issues ou pull requests.
