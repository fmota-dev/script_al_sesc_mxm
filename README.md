# Processador de Arquivos Financeiros (ME, OD, RF)

Este script em Python automatiza o processamento de arquivos Excel contendo informações de mensalidades nas áreas de **Esporte**, **Odontologia** e **Refeições/Lanches**, gerando planilhas formatadas para importação contábil e logs de processamento.

## ⚙️ Funcionalidades

- Lê arquivos `ME.xlsx`, `OD.xlsx` e `RF.xlsx` com dados de CPF e valores.
- Aplica regras de negócio específicas para cada tipo de arquivo:
  - Arredondamentos e distribuições proporcionais.
  - Substituição de CPFs por titulares quando aplicável.
  - Geração de lançamentos contábeis (débito/crédito).
  - Adição de linhas de ajuste de 50% e 100% conforme necessário.
- Gera arquivos Excel prontos para importação e logs detalhados.
- Solicita o código AL do usuário, padronizando o valor.
- Gera nomes de documentos com base no mês anterior.
- Cria automaticamente as pastas `arquivos_importacao` e `logs`.

## 📂 Estrutura de Pastas Esperada

```plaintext
/<pasta_do_script>
├── ME.xlsx
├── OD.xlsx
├── RF.xlsx
├── arquivos_importacao/
└── logs/
```

## 📌 Como Executar

1. Coloque os arquivos `ME.xlsx`, `OD.xlsx`, `RF.xlsx` na mesma pasta que o script.
2. Execute o script com Python:

```bash
python main.py
```

3. Informe o código AL solicitado para cada arquivo.
4. Os arquivos processados serão salvos em `arquivos_importacao` e os logs em `logs`.

## 📥 Saída

- Arquivos `.xlsx` gerados com nome no formato `ME202403.xlsx`, `OD202403.xlsx`, etc.
- Logs nomeados por tipo: `log_processamento_me202403.txt`, etc.
- Log de erros centralizado em `logs/erros_processamento.txt`.

## ✅ Dependências

- Python 3.7+
- Pandas
- openpyxl (para salvar arquivos .xlsx)

Instale as dependências com:

```bash
pip install pandas openpyxl
```

## 📎 Observações

- O script sempre considera o **mês anterior** ao da execução.
- A última linha dos arquivos de entrada é automaticamente removida (normalmente somatórios).
- Os arquivos devem conter colunas como `CPF`, `CPF_TITULAR` e `VALOR` (ou equivalentes).
- Os dados são agrupados por CPF e tratados conforme regras específicas de cada área (ME, OD, RF).

---

🚀 Feito por @fmota.dev
