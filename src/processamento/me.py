import pandas as pd

from config import TEMPLATE_IMPORTACAO_BASE
from utils.helpers import (
    arredondar,
    formatar_historico,
    nome_documento,
    ultimo_dia_mes_anterior,
)
from utils.log import escrever_no_log


def verificar_valores(valor_total_final, valor_50_final, planilha_final, caminho_log):
    if arredondar(valor_total_final) != arredondar(valor_50_final * 2):
        nova_metade = valor_total_final / 2
        escrever_no_log(
            "⚠️  O valor total {valor_total_final} é diferente do valor de 50% {valor_50_final * 2}",
            caminho_log,
        )
        escrever_no_log("⚠️  Ajustando linha de 50% para: {nova_metade}", caminho_log)
        planilha_final[-2]["VALOR"] = nova_metade
    else:
        escrever_no_log(
            "✅ O valor total {valor_total_final} é igual ao valor de 50% {valor_50_final * 2}",
            caminho_log,
        )


def salvar_excel_formatado(df: pd.DataFrame, caminho_saida: str, caminho_log: str):
    try:
        df.to_excel(caminho_saida, index=False)
    except Exception as e:
        escrever_no_log(f"⚠️ Erro ao salvar o arquivo {caminho_saida}: {e}", caminho_log)


def processar_me(caminho, caminho_saida, codigo_al, caminho_log):
    escrever_no_log("🔹 Iniciando processamento do arquivo ME...", caminho_log)
    try:
        df = pd.read_excel(
            caminho, dtype={"CPF": str, "CPF_TITULAR": str, "VALOR": float}
        )
    except Exception as e:
        escrever_no_log(f"⚠️ Erro ao ler o arquivo {caminho}: {e}", caminho_log)
        return

    df = df.iloc[:-1]
    escrever_no_log("📝 Última linha removida do DataFrame", caminho_log)

    df["CPF"] = df.apply(
        lambda row: row["CPF_TITULAR"] if pd.notna(row["CPF_TITULAR"]) else row["CPF"],
        axis=1,
    )
    escrever_no_log("🔄 CPFs substituídos por titulares quando aplicável", caminho_log)

    df = df.dropna(subset=["CPF", "VALOR"])
    escrever_no_log("🧹 Linhas com CPF ou VALOR nulos removidas", caminho_log)

    df_agrupado = df.groupby("CPF", as_index=False)["VALOR"].sum()
    data_lancamento = ultimo_dia_mes_anterior()
    documento = nome_documento("ME")
    area = "ESPORTE"
    planilha_final = []
    sequencia = 1

    for cpf in df_agrupado["CPF"]:
        valores_cpf = df_agrupado[df_agrupado["CPF"] == cpf]
        template = TEMPLATE_IMPORTACAO_BASE.copy()
        template.update(
            {
                "DATA DO LANCAMENTO": data_lancamento,
                "DOCUMENTO": documento,
                "VALOR": valores_cpf["VALOR"].values[0] * 0.5,
                "HISTORICO": formatar_historico(codigo_al, area),
                "SEQUENCIA": sequencia,
                "CLIENTE": cpf,
            }
        )
        planilha_final.append(template)
        sequencia += 1

    valor_total = df_agrupado["VALOR"].sum()
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
            "CENTRO DE CUSTO": "02053",
            "SEQUENCIA": len(planilha_final) + 1,
            "PROJETO": "20001",
        }
    )
    planilha_final.append(linha_50)

    linha_total = TEMPLATE_IMPORTACAO_BASE.copy()
    linha_total.update(
        {
            "DATA DO LANCAMENTO": data_lancamento,
            "DOCUMENTO": documento,
            "CONTA CONTABIL": "21881010101001",
            "INDICADOR DE CONTA": "C",
            "VALOR": arredondar(valor_desconto_50 * 2),
            "HISTORICO": formatar_historico(codigo_al, area),
            "SEQUENCIA": len(planilha_final) + 1,
        }
    )
    planilha_final.append(linha_total)

    verificar_valores(linha_total["VALOR"], linha_50["VALOR"], planilha_final, caminho_log)

    df_final = pd.DataFrame(planilha_final)
    salvar_excel_formatado(df_final, caminho_saida, caminho_log)
    escrever_no_log(f"✅ Arquivo ME gerado com sucesso: {caminho_saida}", caminho_log)
