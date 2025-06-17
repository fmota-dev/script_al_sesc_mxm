import math

import pandas as pd

from config import TEMPLATE_IMPORTACAO_BASE
from utils.helpers import (
    arredondar,
    formatar_historico,
    nome_documento,
    truncar_se_mais_de_duas_casas,
    ultimo_dia_mes_anterior,
)
from utils.log import escrever_no_log


def salvar_excel_formatado(df: pd.DataFrame, caminho_saida: str, caminho_log: str):
    try:
        df.to_excel(caminho_saida, index=False)
    except Exception as e:
        escrever_no_log(f"⚠️ Erro ao salvar o arquivo {caminho_saida}: {e}", caminho_log)


def processar_od(caminho, caminho_saida, codigo_al, caminho_log):
    escrever_no_log("🔹 Iniciando processamento do arquivo OD...", caminho_log)
    acumulador = {"truncado": 0.0}
    df = pd.read_excel(caminho, dtype={"CPF": str, "CPF_TITULAR": str})

    valor_col = None
    if "VALOR" in df.columns:
        valor_col = "VALOR"
    elif "VALOR_TOTAL" in df.columns:
        valor_col = "VALOR_TOTAL"
    else:
        escrever_no_log(
            "⚠️ Nenhuma das colunas 'VALOR' ou 'VALOR_TOTAL' encontrada.", caminho_log
        )
        return

    df = df.dropna(subset=["CPF", valor_col])
    escrever_no_log(
        "🧹 Removendo linhas com CPF ou VALOR/VALOR_TOTAL nulos", caminho_log
    )

    df["CPF"] = df.apply(
        lambda row: row["CPF_TITULAR"]
        if pd.notnull(row["CPF_TITULAR"]) and row["CPF_TITULAR"] != row["CPF"]
        else row["CPF"],
        axis=1,
    )

    df["index_original"] = df.index

    df_agrupado = df.groupby("CPF")[valor_col].sum().reset_index()
    df_agrupado[valor_col] = df_agrupado[valor_col].apply(arredondar)

    df_agrupado = pd.merge(
        df_agrupado, df[["CPF", "index_original"]], on="CPF", how="left"
    )
    df_agrupado = (
        df_agrupado.drop_duplicates(subset="CPF")
        .sort_values("index_original")
        .drop("index_original", axis=1)
    )

    planilha_final = []
    sequencia = 1
    data_lancamento = ultimo_dia_mes_anterior()
    documento = nome_documento("OD")
    area = "CONSULTAS ODONTOLÓGICAS"

    for cpf in df_agrupado["CPF"]:
        valores_cpf = df_agrupado[df_agrupado["CPF"] == cpf]
        if not valores_cpf.empty:
            template = TEMPLATE_IMPORTACAO_BASE.copy()
            template.update(
                {
                    "DATA DO LANCAMENTO": data_lancamento,
                    "DOCUMENTO": documento,
                    "CONTA CONTABIL": "11381010101001",
                    "INDICADOR DE CONTA": "D",
                    "VALOR": truncar_se_mais_de_duas_casas(
                        valores_cpf[valor_col].values[0] * 0.5, acumulador
                    ),
                    "HISTORICO": formatar_historico(codigo_al, area),
                    "SEQUENCIA": sequencia,
                    "CLIENTE": cpf,
                }
            )
            planilha_final.append(template)
            sequencia += 1

    valor_total = df_agrupado[valor_col].sum()
    valor_desconto_50 = valor_total * 0.5

    linha_50 = TEMPLATE_IMPORTACAO_BASE.copy()
    linha_50.update(
        {
            "DATA DO LANCAMENTO": data_lancamento,
            "DOCUMENTO": documento,
            "CONTA CONTABIL": "31321019901001",
            "INDICADOR DE CONTA": "D",
            "VALOR": arredondar(valor_desconto_50),
            "HISTORICO": formatar_historico(codigo_al, area),
            "CENTRO DE CUSTO": "02050",
            "SEQUENCIA": sequencia,
            "PROJETO": "20001",
        }
    )
    planilha_final.append(linha_50)
    sequencia += 1

    linha_total = TEMPLATE_IMPORTACAO_BASE.copy()
    linha_total.update(
        {
            "DATA DO LANCAMENTO": data_lancamento,
            "DOCUMENTO": documento,
            "CONTA CONTABIL": "21881010101001",
            "INDICADOR DE CONTA": "C",
            "VALOR": arredondar(valor_desconto_50 * 2),
            "HISTORICO": formatar_historico(codigo_al, area),
            "SEQUENCIA": sequencia,
        }
    )
    planilha_final.append(linha_total)

    df_final = pd.DataFrame(planilha_final)
    salvar_excel_formatado(df_final, caminho_saida, caminho_log)
    escrever_no_log(f"✅ Arquivo OD gerado com sucesso: {caminho_saida}", caminho_log)
