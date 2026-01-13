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

    # Calcular valores para as linhas contábeis
    # Valor original (antes do truncamento)
    valor_total_original = df_agrupado["VALOR"].sum()

    # Linha 50% (débito) - sempre será calculado para garantir que débitos = crédito
    # Soma dos CPFs truncados (antes de qualquer ajuste)
    valor_cpfs_truncado = sum(item["VALOR"] for item in planilha_final)

    # A linha 50% é ajustada para que: CPFs + Linha50% = Valor_Original
    valor_linha_50 = valor_total_original - valor_cpfs_truncado

    # Calcular diferença para distribuir entre os CPFs (em centavos)
    # Queremos que soma_cpfs fique o mais próximo possível de 50% do original
    valor_50_ideal = round(valor_total_original * 0.5, 2)
    diferenca_total = valor_50_ideal - valor_cpfs_truncado
    centavos_para_ajustar = round(diferenca_total * 100)

    # Distribuir centavos entre os CPFs (bidirecional: adicionar ou subtrair)
    if centavos_para_ajustar != 0:
        cpfs_ajustados = 0
        ajuste_unitario = 0.01 if centavos_para_ajustar > 0 else -0.01
        centavos_restantes = abs(centavos_para_ajustar)

        for item in planilha_final:
            if centavos_restantes <= 0:
                break
            if "CLIENTE" in item:  # Garante que é uma linha de CPF
                item["VALOR"] = round(item["VALOR"] + ajuste_unitario, 2)
                centavos_restantes -= 1
                cpfs_ajustados += 1

        # Recalcular linha 50% após ajustes nos CPFs
        valor_cpfs_ajustado = sum(item["VALOR"] for item in planilha_final)
        valor_linha_50 = valor_total_original - valor_cpfs_ajustado

        acao = "adicionado" if centavos_para_ajustar > 0 else "subtraído"
        escrever_no_log(
            f"🩹 Ajuste de truncamento: R$0.01 {acao} de {cpfs_ajustados} CPF(s) para compensar diferença.",
            caminho_log,
        )

    linha_50 = TEMPLATE_IMPORTACAO_BASE.copy()
    linha_50.update(
        {
            "DATA DO LANCAMENTO": data_lancamento,
            "DOCUMENTO": documento,
            "CONTA CONTABIL": "31321019901001",
            "INDICADOR DE CONTA": "D",
            "VALOR": valor_linha_50,
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
            "VALOR": valor_total_original,
            "HISTORICO": f"MENSALIDADE {formatar_historico(codigo_al, area)}",
            "SEQUENCIA": len(planilha_final) + 1,
        }
    )
    planilha_final.append(linha_total)

    df_final = pd.DataFrame(planilha_final)
    salvar_excel_formatado(df_final, caminho_saida, caminho_log)
    escrever_no_log(f"✅ Arquivo ME gerado com sucesso: {caminho_saida}", caminho_log)
