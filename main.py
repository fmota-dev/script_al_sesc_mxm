import datetime
import os
import sys
from decimal import ROUND_HALF_UP, Decimal
from pathlib import Path

import pandas as pd

# Definindo o template de importação

TEMPLATE_IMPORTACAO_BASE = {
    # O que estiver como None será preenchido depois
    "CODIGO DA EMPRESA": "07  ",
    "LOTE": "CTB",
    "DATA DO LANCAMENTO": None,
    "DOCUMENTO": None,
    "CONTA CONTABIL": "11381010101001",
    "INDICADOR DE CONTA": "D",
    "VALOR": None,
    "HISTORICO": None,
    "BRANCO": "",
    "VALOR SEGUNDA MOEDA": "",
    "BRANCO2": "",
    "CENTRO DE CUSTO": "",
    "SEQUENCIA": None,
    "PROJETO": "",
    "FORNECEDOR": "",
    "CLIENTE": None,
    "VALOR SEGUNDA MOEDA2": "",
    "HIST PADRAO": "",
    "HIST. PADRAO - COMPLEMENTO 1": "",
    "HIST. PADRAO - COMPLEMENTO 2": "",
    "HIST. PADRAO - COMPLEMENTO 3": "",
    "NUMERO DO TITULO": "",
    "CONVERTER MOEDA": "N",
    "EXCLUIR LANÇAMENTOS": "N",
}

# Função para pegar o último dia do mês anterior


def ultimo_dia_mes_anterior():
    hoje = datetime.datetime.now()
    primeiro_dia_mes = datetime.datetime(hoje.year, hoje.month, 1)
    ultimo_dia = primeiro_dia_mes - datetime.timedelta(days=1)
    return ultimo_dia.strftime("%Y%m%d")


# Função para pegar o ano e mês anterior no formato YYYYMM


def ano_mes_anterior():
    hoje = datetime.datetime.now()
    mes = hoje.month - 1 or 12
    ano = hoje.year if hoje.month > 1 else hoje.year - 1
    return f"{ano}{mes:02d}"


# Função para formatar histórico


def formatar_historico(codigo_al, area):
    ano_mes = ano_mes_anterior()
    return f"MENSALIDADE {area} ({codigo_al} SESC) {ano_mes[4:]}/{ano_mes[:4]}"


# Função para arredondar valores


def arredondar(valor):
    return float(Decimal(valor).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))


# Função para escrever no ‘log’


def escrever_no_log(mensagem, caminho_log):
    from datetime import datetime

    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    mensagem_com_timestamp = f"[{timestamp}] {mensagem}"
    try:
        with open(caminho_log, "a", encoding="utf-8") as log_file:
            log_file.write(mensagem_com_timestamp + "\n")
        print(mensagem_com_timestamp)
    except Exception as e:
        print(f"Erro ao escrever no log: {e}")


# Função para gerar o nome do documento


def nome_documento(tipo: str) -> str:
    hoje = datetime.datetime.now()
    mes = hoje.month - 1 or 12
    ano = hoje.year if hoje.month > 1 else hoje.year - 1
    return f"{tipo.upper()}{mes:02d}{ano % 100:02d}"


# Função para salvar o novo DataFrame em Excel


def salvar_excel_formatado(
    df: pd.DataFrame,
    caminho_saida: str,
    caminho_log: str,
):
    try:
        # Salva o DataFrame diretamente sem aplicar formatação
        df.to_excel(caminho_saida, index=False)

    except Exception as e:
        escrever_no_log(f"⚠️ Erro ao salvar o arquivo {caminho_saida}: {e}", caminho_log)


# Função para verificar se o valor total é igual ao valor de 50% * 2


def verificar_valores(valor_total_final, valor_50_final, resultado):
    if arredondar(valor_total_final) != arredondar(valor_50_final * 2):
        nova_metade = valor_total_final / 2
        print(
            f"⚠️  O valor total {valor_total_final} é diferente do valor de 50% {valor_50_final * 2}"
        )
        print(f"⚠️  Ajustando linha de 50% para: {nova_metade}")
        resultado[-2]["VALOR"] = nova_metade
    else:
        print(
            f"✅ O valor total {valor_total_final} é igual ao valor de 50% {valor_50_final * 2}"
        )


# Função para processar o arquivo ME


def processar_me(caminho, caminho_saida, codigo_al, caminho_log):
    escrever_no_log("🔹 Iniciando processamento do arquivo ME...", caminho_log)
    try:
        df = pd.read_excel(
            caminho, dtype={"CPF": str, "CPF_TITULAR": str, "VALOR": float}
        )
    except Exception as e:
        escrever_no_log(f"⚠️ Erro ao ler o arquivo {caminho}: {e}", caminho_log)
        return

    df = df.iloc[:-1]  # Remove última linha
    escrever_no_log("📝 Última linha removida do DataFrame", caminho_log)

    # Substitui CPF pelo CPF_TITULAR, se houver
    df["CPF"] = df.apply(
        lambda row: row["CPF_TITULAR"] if pd.notna(row["CPF_TITULAR"]) else row["CPF"],
        axis=1,
    )
    escrever_no_log("🔄 CPFs substituídos por titulares quando aplicável", caminho_log)

    # Remove linhas inválidas
    df = df.dropna(subset=["CPF", "VALOR"])
    escrever_no_log("🧹 Linhas com CPF ou VALOR nulos removidas", caminho_log)

    # Agrupa os valores por CPF
    df_agrupado = df.groupby("CPF", as_index=False)["VALOR"].sum()

    # Dados fixos
    data_lancamento = ultimo_dia_mes_anterior()
    documento = nome_documento("ME")
    area = "ESPORTE"
    planilha_final = []
    sequencia = 1

    # Gera os lançamentos para cada CPF
    for cpf in df_agrupado["CPF"]:
        valores_cpf = df_agrupado[df_agrupado["CPF"] == cpf]
        template_importacao = TEMPLATE_IMPORTACAO_BASE.copy()
        template_importacao.update(
            {
                "DATA DO LANCAMENTO": data_lancamento,
                "DOCUMENTO": documento,
                "VALOR": valores_cpf["VALOR"].values[0] * 0.5,
                "HISTORICO": formatar_historico(codigo_al, area),
                "SEQUENCIA": sequencia,
                "CLIENTE": cpf,
            }
        )
        planilha_final.append(template_importacao)
        # escrever_no_log(
        #     f"📌 Adicionando linha para CPF {cpf} com valor {template_importacao['VALOR']}",
        #     caminho_log,
        # )
        sequencia += 1

    valor_total = df_agrupado["VALOR"].sum()
    valor_desconto_50 = valor_total * 0.5

    # Linha de 50%
    linha_50_porcento = TEMPLATE_IMPORTACAO_BASE.copy()
    linha_50_porcento.update(
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
    planilha_final.append(linha_50_porcento)

    # Linha de 100%
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

    valor_total_final = linha_total["VALOR"]
    valor_50_final = linha_50_porcento["VALOR"]

    print(f"Valor total: {valor_total_final}")
    print(f"Valor 50%: {valor_50_final}")

    # Verifica se o valor total é igual ao valor de 50% * 2

    verificar_valores(valor_total_final, valor_50_final, planilha_final)

    # Salvar o resultado
    df_final = pd.DataFrame(planilha_final)
    salvar_excel_formatado(df_final, caminho_saida, caminho_log)
    escrever_no_log(f"✅ Arquivo ME gerado com sucesso: {caminho_saida}", caminho_log)


# Função para processar o arquivo OD


def processar_od(caminho, caminho_saida, codigo_al, caminho_log):
    escrever_no_log("🔹 Iniciando processamento do arquivo OD...", caminho_log)
    df = pd.read_excel(caminho, dtype={"CPF": str, "CPF_TITULAR": str})

    # Verificar se as colunas 'VALOR' ou 'VALOR_TOTAL' estão presentes
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

    # Substituindo CPF pelo CPF_TITULAR se houver valor na coluna CPF_TITULAR e for diferente
    df["CPF"] = df.apply(
        lambda row: (
            row["CPF_TITULAR"]
            if pd.notnull(row["CPF_TITULAR"]) and row["CPF_TITULAR"] != row["CPF"]
            else row["CPF"]
        ),
        axis=1,
    )

    # Manter a ordem original dos CPFs
    # Adiciona uma coluna de índice original para manter a ordem
    df["index_original"] = df.index

    # Agrupar os valores por CPF (somando os valores)
    df_agrupado = df.groupby("CPF")[valor_col].sum().reset_index()
    df_agrupado[valor_col] = df_agrupado[valor_col].apply(arredondar)

    # Utilizar merge para manter a ordem original com base no CPF
    df_agrupado = pd.merge(
        df_agrupado, df[["CPF", "index_original"]], on="CPF", how="left"
    )
    # Remover duplicatas, mantendo uma linha por CPF
    df_agrupado = df_agrupado.drop_duplicates(subset="CPF")
    df_agrupado = df_agrupado.sort_values("index_original").drop(
        "index_original", axis=1
    )

    # Processamento final
    planilha_final = []
    sequencia = 1
    data_lancamento = ultimo_dia_mes_anterior()
    documento = nome_documento("OD")
    area = "CONSULTAS ODONTÓLOGICAS"

    # Iterando sobre os CPFs únicos
    for cpf in df_agrupado["CPF"]:
        valores_cpf = df_agrupado[df_agrupado["CPF"] == cpf]
        if not valores_cpf.empty:
            template_importacao = TEMPLATE_IMPORTACAO_BASE.copy()
            template_importacao.update(
                {
                    "DATA DO LANCAMENTO": data_lancamento,
                    "DOCUMENTO": documento,
                    "CONTA CONTABIL": "11381010101001",
                    "INDICADOR DE CONTA": "D",
                    "VALOR": valores_cpf[valor_col].values[0] * 0.5,
                    "HISTORICO": formatar_historico(codigo_al, area),
                    "SEQUENCIA": sequencia,
                    "CLIENTE": cpf,
                }
            )

            planilha_final.append(template_importacao)
            # escrever_no_log(
            #     f"📌 Adicionando linha para CPF {cpf} com valor {template_importacao['VALOR']}",
            #     caminho_log,
            # )
            sequencia += 1

    # Adicionando as duas linhas ao final (50% e 100%)
    valor_total = df_agrupado[valor_col].sum()
    valor_desconto_50 = valor_total * 0.5

    # Linha de 50% do valor total
    linha_50_porcento = TEMPLATE_IMPORTACAO_BASE.copy()
    linha_50_porcento.update(
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
    planilha_final.append(linha_50_porcento)
    sequencia += 1

    # Linha de 100% do valor total
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

    valor_total_final = linha_total["VALOR"]
    valor_50_final = linha_50_porcento["VALOR"]

    print(f"Valor total: {valor_total_final}")
    print(f"Valor 50%: {valor_50_final}")

    verificar_valores(valor_total_final, valor_50_final, planilha_final)

    # Salvar em Excel
    df_final = pd.DataFrame(planilha_final)
    salvar_excel_formatado(df_final, caminho_saida, caminho_log)
    escrever_no_log(f"✅ Arquivo OD gerado com sucesso: {caminho_saida}", caminho_log)


# Função para processar o arquivo RF


def processar_rf(caminho, caminho_saida, codigo_al, caminho_log):
    escrever_no_log("🔹 Iniciando processamento do arquivo RF...", caminho_log)
    df = pd.read_excel(caminho, dtype=str)

    # Verificar a coluna 'CPF' independentemente de maiúsculas ou minúsculas
    cpf_colunas = [col for col in df.columns if col.lower() == "cpf"]
    if cpf_colunas:
        cpf_coluna = cpf_colunas[0]
    else:
        escrever_no_log("⚠️ Coluna 'CPF' não encontrada.", caminho_log)
        return

    # Verificar se as colunas 'VALOR', 'VALOR_TOTAL' ou 'ValorTotalProduto' estão presentes
    if "VALOR" in df.columns:
        valor_col = "VALOR"
    elif "VALOR_TOTAL" in df.columns:
        valor_col = "VALOR_TOTAL"
    elif "ValorTotalProduto" in df.columns:
        valor_col = "ValorTotalProduto"
    else:
        escrever_no_log(
            "⚠️ Nenhuma das colunas 'VALOR', 'VALOR_TOTAL' ou 'ValorTotalProduto' encontrada.",
            caminho_log,
        )
        return

    df = df.iloc[:-1]  # Remover a última linha
    escrever_no_log("📝 Última linha removida do DataFrame", caminho_log)

    # Alteração de CPF, garantindo que seja string e levando em conta a coluna correta
    for i, cpf in df.iterrows():
        cpf_original = cpf[cpf_coluna]
        # Verificar se existe a coluna 'CPF_TITULAR'
        cpf_titular = cpf.get("CPF_TITULAR", None)
        if pd.notna(cpf_titular):
            df.at[i, cpf_coluna] = str(cpf_titular)  # Garantir que seja string
            escrever_no_log(
                f"🔄 Alterando CPF: {cpf_original} → {cpf_titular}", caminho_log
            )

    # Remover linhas sem CPF ou VALOR
    df = df.dropna(subset=[cpf_coluna, valor_col])
    escrever_no_log("🧹 Removendo linhas com CPF ou VALOR nulos", caminho_log)

    # Garantir que a coluna 'VALOR' seja do tipo float antes de aplicar o arredondamento
    # Converte para float, erros são convertidos para NaN
    df[valor_col] = pd.to_numeric(df[valor_col], errors="coerce")
    df[valor_col] = df[valor_col].fillna(0)  # Substitui NaN por 0, se houver

    # Agrupar os valores por CPF
    df_agrupado = df.groupby(cpf_coluna)[valor_col].sum().reset_index()
    df_agrupado[valor_col] = df_agrupado[valor_col].apply(arredondar)

    # Processamento final
    planilha_final = []
    sequencia = 1
    data_lancamento = ultimo_dia_mes_anterior()
    documento = nome_documento("RF")
    area = "FORNECIMENTO DE REFEIÇÕES E LANCHES"

    # Iterar sobre os CPFs únicos e adicionar ao resultado
    for cpf in df[cpf_coluna].unique():
        valores_cpf = df_agrupado[df_agrupado[cpf_coluna] == cpf]
        if not valores_cpf.empty:
            template_importacao = TEMPLATE_IMPORTACAO_BASE.copy()
            template_importacao.update(
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

            planilha_final.append(template_importacao)
            # escrever_no_log(
            #     f"📌 Adicionando linha para CPF {cpf} com valor {template_importacao['VALOR']}",
            #     caminho_log,
            # )
            sequencia += 1

    # Adicionar linha de total
    valor_total = df_agrupado[valor_col].sum()

    linha_total = TEMPLATE_IMPORTACAO_BASE.copy()
    linha_total.update(
        {
            "DATA DO LANCAMENTO": data_lancamento,
            "DOCUMENTO": documento,
            "CONTA CONTABIL": "21881010101001",  # Conta específica para o total
            "INDICADOR DE CONTA": "C",  # Indicador de conta para o total
            "VALOR": arredondar(valor_total),
            "HISTORICO": formatar_historico(codigo_al, area),
            "SEQUENCIA": sequencia,  # Número sequencial para a linha de total
            "CLIENTE": "",  # Coluna CLIENTE em branco na linha de total
        }
    )

    planilha_final.append(linha_total)
    escrever_no_log("📊 Adicionando linha de total", caminho_log)

    # Salvar em Excel
    df_final = pd.DataFrame(planilha_final)
    salvar_excel_formatado(df_final, caminho_saida, caminho_log)
    escrever_no_log(f"✅ Arquivo RF gerado com sucesso: {caminho_saida}", caminho_log)


# Função para pedir o código AL de forma genérica para qualquer arquivo


def pedir_codigo_al(arquivo):
    """Função para pedir o código AL de forma genérica para qualquer arquivo."""
    while True:
        codigo_al = input(f"Digite o código AL para o arquivo {arquivo}: ").strip()

        if not codigo_al:
            print("⚠️ Código AL inválido ou não informado. Tente novamente.")
            continue  # Volta ao início do loop para pedir novamente

        # Verifica se o código já começa com "AL " (com espaço)
        if codigo_al.upper().startswith("AL "):
            break
        # Se começa com "AL" mas sem o espaço, adiciona o espaço
        elif codigo_al.upper().startswith("AL"):
            codigo_al = "AL " + codigo_al[2:].strip()
            break
        else:
            # Se não começa com AL, adiciona do zero
            codigo_al = "AL " + codigo_al
            break
    return codigo_al


# Função para iterar por arquivos e processar conforme necessário


def processar_arquivos(pasta_entrada, pasta_saida, pasta_logs):
    arquivos = ["ME.xlsx", "OD.xlsx", "RF.xlsx"]
    arquivos_gerados = []  # Lista para armazenar os arquivos que foram gerados com sucesso

    for arquivo in arquivos:
        caminho_entrada = os.path.join(pasta_entrada, arquivo)

        if not os.path.exists(caminho_entrada):
            mensagem_erro = (
                f"Erro: Arquivo {arquivo} não encontrado em {pasta_entrada}."
            )
            caminho_log_erro = os.path.join(pasta_logs, "erros_processamento.txt")
            continue  # Pula para o próximo arquivo

        try:
            # Definir o caminho de saída e o caminho de log antes das condições
            caminho_saida = None
            caminho_log = None

            if arquivo == "ME.xlsx":
                caminho_saida = os.path.join(
                    pasta_saida, f"ME{ano_mes_anterior()}.xlsx"
                )
                caminho_log = os.path.join(
                    pasta_logs, f"log_processamento_me{ano_mes_anterior()}.txt"
                )
                codigo_al_me = pedir_codigo_al(arquivo)
                processar_me(caminho_entrada, caminho_saida, codigo_al_me, caminho_log)

            elif arquivo == "OD.xlsx":
                caminho_saida = os.path.join(
                    pasta_saida, f"OD{ano_mes_anterior()}.xlsx"
                )
                caminho_log = os.path.join(
                    pasta_logs, f"log_processamento_od{ano_mes_anterior()}.txt"
                )
                codigo_al_od = pedir_codigo_al(arquivo)
                processar_od(caminho_entrada, caminho_saida, codigo_al_od, caminho_log)

            elif arquivo == "RF.xlsx":
                caminho_saida = os.path.join(
                    pasta_saida, f"RF{ano_mes_anterior()}.xlsx"
                )
                caminho_log = os.path.join(
                    pasta_logs, f"log_processamento_rf{ano_mes_anterior()}.txt"
                )
                codigo_al_rf = pedir_codigo_al(arquivo)
                processar_rf(caminho_entrada, caminho_saida, codigo_al_rf, caminho_log)

            # Verifica se o arquivo foi gerado e adiciona à lista
            if caminho_saida and os.path.exists(caminho_saida):
                arquivos_gerados.append(caminho_saida)

            escrever_no_log(f"🆗 Processamento concluído para {arquivo}.", caminho_log)

        except Exception as e:
            mensagem_erro = (
                f"⚠️ Erro ao processar o arquivo {arquivo}: {type(e).__name__} - {e}"
            )
            caminho_log_erro = os.path.join(pasta_logs, "erros_processamento.txt")
            escrever_no_log(mensagem_erro, caminho_log_erro)
            input("Pressione enter para sair...")

    # Mensagem de finalização se pelo menos um arquivo foi gerado
    if arquivos_gerados:
        print(
            "Processamento finalizado. Verifique os arquivos gerados na pasta 'arquivos_importacao'."
        )
        input("Pressione enter para sair...")
    else:
        print("Nenhum arquivo foi gerado. Verifique os logs para mais detalhes.")
        input("Pressione enter para sair...")


if __name__ == "__main__":
    if getattr(sys, "frozen", False):  # Verifica se está rodando como executável
        pasta_entrada = Path(sys.executable).parent  # Diretório do executável
    else:
        pasta_entrada = Path(__file__).parent  # Diretório do script

    pasta_saida = pasta_entrada / "arquivos_importacao"
    pasta_logs = pasta_entrada / "logs"

    # Chama a função para processar os arquivos
    processar_arquivos(
        pasta_entrada,
        pasta_saida,
        pasta_logs,
    )
