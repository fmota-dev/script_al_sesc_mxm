import pandas as pd

from config import TEMPLATE_IMPORTACAO_BASE
from utils.helpers import (
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


def processar_me(caminho, caminho_saida, codigo_al, caminho_log):
    escrever_no_log("🔹 Iniciando processamento do arquivo ME...", caminho_log)
    acumulador = {"truncado": 0.0}
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
        valor_com_desconto = valores_cpf["VALOR"].values[0] * 0.5
        valor_truncado = truncar_se_mais_de_duas_casas(valor_com_desconto, acumulador)

        template = TEMPLATE_IMPORTACAO_BASE.copy()
        template.update(
            {
                "DATA DO LANCAMENTO": data_lancamento,
                "DOCUMENTO": documento,
                "VALOR": valor_truncado,
                "HISTORICO": f"MENSALIDADE {formatar_historico(codigo_al, area)}",
                "SEQUENCIA": sequencia,
                "CLIENTE": cpf,
            }
        )
        planilha_final.append(template)
        sequencia += 1

    # Se o truncamento total foi de R$0.01, adiciona a algum CPF que tenha 1 ou 0 casas decimais
    if round(acumulador["truncado"], 6) == 0.01:
        aplicado = False
        for item in planilha_final:
            if "CLIENTE" in item:  # Garante que é uma linha de CPF e não de totalização
                casas_decimais = str(item["VALOR"]).split(".")
                if len(casas_decimais) == 2 and len(casas_decimais[1]) < 2:
                    item["VALOR"] += 0.01
                    aplicado = True
                    escrever_no_log(
                        f"🩹 Ajuste aplicado: R$0.01 adicionado ao CPF {item['CLIENTE']} para compensar truncamento.",
                        caminho_log,
                    )
                    break

        if not aplicado:
            escrever_no_log(
                "⚠️ Nenhum CPF com 1 ou 0 casas decimais encontrado para aplicar o ajuste de R$0.01.",
                caminho_log,
            )

    valor_total = df_agrupado["VALOR"].sum()
    valor_desconto_50 = valor_total * 0.5

    linha_50 = TEMPLATE_IMPORTACAO_BASE.copy()
    linha_50.update(
        {
            "DATA DO LANCAMENTO": data_lancamento,
            "DOCUMENTO": documento,
            "CONTA CONTABIL": "31321019901001",
            "INDICADOR DE CONTA": "D",
            "VALOR": valor_desconto_50,
            "HISTORICO": f"MENSALIDADE {formatar_historico(codigo_al, area)}",
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
            "VALOR": valor_desconto_50 * 2,
            "HISTORICO": f"MENSALIDADE {formatar_historico(codigo_al, area)}",
            "SEQUENCIA": len(planilha_final) + 1,
        }
    )
    planilha_final.append(linha_total)

    df_final = pd.DataFrame(planilha_final)
    salvar_excel_formatado(df_final, caminho_saida, caminho_log)
    escrever_no_log(f"✅ Arquivo ME gerado com sucesso: {caminho_saida}", caminho_log)
