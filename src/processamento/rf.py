import pandas as pd

from config import TEMPLATE_IMPORTACAO_BASE
from utils.helpers import (
    arredondar,
    formatar_historico,
    nome_documento,
    ultimo_dia_mes_anterior,
)
from utils.log import escrever_no_log


def salvar_excel_formatado(df: pd.DataFrame, caminho_saida: str, caminho_log: str):
    try:
        df.to_excel(caminho_saida, index=False)
    except Exception as e:
        escrever_no_log(f"⚠️ Erro ao salvar o arquivo {caminho_saida}: {e}", caminho_log)


def processar_rf(caminho, caminho_saida, codigo_al, caminho_log):
    escrever_no_log("🔹 Iniciando processamento do arquivo RF...", caminho_log)
    df = pd.read_excel(caminho, dtype=str)

    cpf_colunas = [col for col in df.columns if col.lower() == "cpf"]
    if not cpf_colunas:
        escrever_no_log("⚠️ Coluna 'CPF' não encontrada.", caminho_log)
        return
    cpf_coluna = cpf_colunas[0]

    if "VALOR" in df.columns:
        valor_col = "VALOR"
    elif "VALOR_TOTAL" in df.columns:
        valor_col = "VALOR_TOTAL"
    elif "ValorTotalProduto" in df.columns:
        valor_col = "ValorTotalProduto"
    else:
        escrever_no_log(
            "⚠️ Nenhuma das colunas de valor esperadas foi encontrada.", caminho_log
        )
        return

    df = df.iloc[:-1]
    escrever_no_log("📝 Última linha removida do DataFrame", caminho_log)

    for i, row in df.iterrows():
        cpf_original = row[cpf_coluna]
        cpf_titular = row.get("CPF_TITULAR")
        if pd.notna(cpf_titular):
            df.at[i, cpf_coluna] = str(cpf_titular)
            escrever_no_log(
                f"🔄 Alterando CPF: {cpf_original} → {cpf_titular}", caminho_log
            )

    df = df.dropna(subset=[cpf_coluna, valor_col])
    escrever_no_log("🧹 Removendo linhas com CPF ou VALOR nulos", caminho_log)

    df[valor_col] = pd.to_numeric(df[valor_col], errors="coerce").fillna(0)

    df_agrupado = df.groupby(cpf_coluna)[valor_col].sum().reset_index()
    df_agrupado[valor_col] = df_agrupado[valor_col].apply(arredondar)

    planilha_final = []
    sequencia = 1
    data_lancamento = ultimo_dia_mes_anterior()
    documento = nome_documento("RF")
    area = "FORNECIMENTO DE REFEIÇÕES E LANCHES"

    for cpf in df_agrupado[cpf_coluna]:
        valores_cpf = df_agrupado[df_agrupado[cpf_coluna] == cpf]
        if not valores_cpf.empty:
            template = TEMPLATE_IMPORTACAO_BASE.copy()
            template.update(
                {
                    "DATA DO LANCAMENTO": data_lancamento,
                    "DOCUMENTO": documento,
                    "CONTA CONTABIL": "11381010101001",
                    "INDICADOR DE CONTA": "D",
                    "VALOR": valores_cpf[valor_col].values[0],
                    "HISTORICO": formatar_historico(codigo_al, area),
                    "SEQUENCIA": sequencia,
                    "CLIENTE": cpf,
                }
            )
            planilha_final.append(template)
            sequencia += 1

    valor_total = df_agrupado[valor_col].sum()

    linha_total = TEMPLATE_IMPORTACAO_BASE.copy()
    linha_total.update(
        {
            "DATA DO LANCAMENTO": data_lancamento,
            "DOCUMENTO": documento,
            "CONTA CONTABIL": "21881010101001",
            "INDICADOR DE CONTA": "C",
            "VALOR": arredondar(valor_total),
            "HISTORICO": formatar_historico(codigo_al, area),
            "SEQUENCIA": sequencia,
            "CLIENTE": "",
        }
    )

    planilha_final.append(linha_total)
    escrever_no_log("📊 Adicionando linha de total", caminho_log)

    df_final = pd.DataFrame(planilha_final)
    salvar_excel_formatado(df_final, caminho_saida, caminho_log)
    escrever_no_log(f"✅ Arquivo RF gerado com sucesso: {caminho_saida}", caminho_log)
