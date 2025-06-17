import os
import sys
from pathlib import Path

from processamento.me import processar_me
from processamento.od import processar_od
from processamento.rf import processar_rf
from utils.helpers import ano_mes_anterior, pedir_codigo_al
from utils.log import escrever_no_log


def processar_arquivos(pasta_entrada, pasta_saida, pasta_logs):
    arquivos = ["ME.xlsx", "OD.xlsx", "RF.xlsx"]
    arquivos_gerados = []

    for arquivo in arquivos:
        caminho_entrada = os.path.join(pasta_entrada, arquivo)
        if not os.path.exists(caminho_entrada):
            continue

        caminho_saida = os.path.join(
            pasta_saida, f"{arquivo[:2]}{ano_mes_anterior()}.xlsx"
        )
        caminho_log = os.path.join(
            pasta_logs,
            f"log_processamento_{arquivo[:2].lower()}{ano_mes_anterior()}.txt",
        )
        codigo_al = pedir_codigo_al(arquivo)

        try:
            if arquivo == "ME.xlsx":
                processar_me(caminho_entrada, caminho_saida, codigo_al, caminho_log)
            elif arquivo == "OD.xlsx":
                processar_od(caminho_entrada, caminho_saida, codigo_al, caminho_log)
            elif arquivo == "RF.xlsx":
                processar_rf(caminho_entrada, caminho_saida, codigo_al, caminho_log)

            if os.path.exists(caminho_saida):
                arquivos_gerados.append(caminho_saida)

            escrever_no_log(f"🆗 Processamento concluído para {arquivo}.", caminho_log)

        except Exception as e:
            escrever_no_log(
                f"⚠️ Erro ao processar {arquivo}: {e}",
                os.path.join(pasta_logs, "erros_processamento.txt"),
            )

    if arquivos_gerados:
        print("✅ Processamento finalizado com sucesso.")
    else:
        print("⚠️ Nenhum arquivo foi gerado.")


if __name__ == "__main__":
    pasta_entrada = (
        Path(sys.executable).parent
        if getattr(sys, "frozen", False)
        else Path(__file__).parent.parent
    )
    pasta_saida = pasta_entrada / "arquivos_importacao"
    pasta_logs = pasta_entrada / "logs"

    try:
        processar_arquivos(pasta_entrada, pasta_saida, pasta_logs)
    except Exception as e:
        escrever_no_log(
            f"⚠️ Erro inesperado: {e}",
            os.path.join(pasta_logs, "erros_processamento.txt"),
        )
        print(f"⚠️ Erro inesperado: {e}")
    except KeyboardInterrupt:
        print(f"\n❌ Processamento interrompido pelo usuário.")
